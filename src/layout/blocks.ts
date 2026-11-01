/**
 * Blocks → boxes. A box is an unbreakable vertical slice (a line, some spacing, a table row)
 * with keep rules; the paginator stacks boxes into columns and pages.
 */
import { blocksToText, type Block, type DocParagraph, type DocTable, type NoteRefInline } from '../model/content.js';
import type { NumberingCounter } from '../model/numbering.js';
import { isVisibleBorder, mergeRunProps, shadingColor, type Border, type ParagraphBorders, type TableBorders } from '../model/props.js';
import type { Theme } from '../model/theme.js';
import { breakLines, buildAtoms, styleFont, type Line, type RangeTracker } from './lines.js';
import { addItem, borderLine, paintLine, type Painter } from './paint.js';
import { resolveRunStyle } from './style.js';
import type { FontProvider } from './types.js';

export interface NoteRef {
  type: 'footnote' | 'endnote';
  id: string;
}

export interface Box {
  height: number;
  /** Do not break between this box and the next. */
  keepWithNext: boolean;
  breakBefore?: 'page' | 'column';
  breakAfter?: 'page' | 'column';
  /** Spacing that disappears at the top of a page reached by overflow. */
  discardAtTop?: boolean;
  /** May be cut off at the bottom of a page instead of moving (space after). */
  canTruncate?: boolean;
  notes?: NoteRef[];
  /** Split into a part that fits `avail` and the rest. */
  split?(avail: number): [Box, Box] | undefined;
  /** Table header rows to repeat when this box starts a new page. */
  repeatHeader?: Box[];
  paint(p: Painter, x: number, y: number): void;
}

export interface BlockContext {
  fonts: FontProvider;
  theme: Theme;
  tracker: RangeTracker;
  counter: NumberingCounter;
  defaultTab: number;
  noteLabel(inline: NoteRefInline): string;
}

/** Private-use symbol-font bullets mapped to Unicode equivalents. */
const SYMBOL_BULLETS: Record<number, string> = {
  0xf0b7: '•',
  0xf0a7: '▪',
  0xf0a8: '◆',
  0xf0d8: '➢',
  0xf076: '❖',
  0xf0fc: '✓',
  0xf06e: '■',
  0xf071: '❑',
  0xf0e0: '➢',
  0xf06f: '□',
};

function normaliseBullet(text: string): string {
  let out = '';
  for (const ch of text) {
    const cp = ch.codePointAt(0)!;
    out += SYMBOL_BULLETS[cp] ?? (cp >= 0xf000 && cp <= 0xf0ff ? '•' : ch);
  }
  return out;
}

function sameBorders(a: ParagraphBorders | undefined, b: ParagraphBorders | undefined): boolean {
  return JSON.stringify(a ?? {}) === JSON.stringify(b ?? {});
}

interface ParagraphState {
  styleId?: string;
  contextual?: boolean;
  afterBox?: Box;
  borders?: ParagraphBorders;
  bottomBorderBox?: Box;
}

/** Lay out a sequence of blocks at a given width. */
export function blockBoxes(blocks: Block[], width: number, ctx: BlockContext): Box[] {
  const boxes: Box[] = [];
  const state: ParagraphState = {};
  for (const block of blocks) {
    if (block.kind === 'paragraph') boxes.push(...paragraphBoxes(block, width, ctx, state));
    else {
      boxes.push(...tableBoxes(block, width, ctx));
      state.styleId = undefined;
      state.afterBox = undefined;
      state.borders = undefined;
    }
  }
  return boxes;
}

export function paragraphBoxes(para: DocParagraph, width: number, ctx: BlockContext, state: ParagraphState = {}): Box[] {
  const props = para.props;
  const ind = props.indent ?? {};
  const left = ind.left ?? 0;
  const right = ind.right ?? 0;
  const firstLine = ind.firstLine ?? (ind.hanging !== undefined ? -ind.hanging : 0);
  const markStyle = resolveRunStyle(para.markProps, ctx.theme);
  const mark = styleFont(ctx.fonts, markStyle, 0x20);

  let label: { text: string; props: typeof para.markProps; suffix: 'tab' | 'space' | 'nothing' } | undefined;
  if (para.numbering) {
    const next = ctx.counter.next(para.numbering.numId, para.numbering.ilvl);
    if (next && next.label) {
      const labelProps = mergeRunProps(para.markProps, next.level.rPr);
      label = { text: normaliseBullet(next.label), props: labelProps, suffix: next.level.suffix };
    }
  }

  const atoms = buildAtoms(para, { fonts: ctx.fonts, theme: ctx.theme, tracker: ctx.tracker, noteLabel: ctx.noteLabel }, label);
  const lines = breakLines(atoms, {
    firstLeft: left + firstLine,
    left,
    right: width - right,
    indentLeft: left,
    tabs: props.tabs ?? [],
    defaultTab: ctx.defaultTab,
    jc: props.jc ?? 'left',
    markAscent: mark.font.ascent * mark.size,
    markDescent: mark.font.descent * mark.size,
  });

  const spacing = props.spacing ?? {};
  let before = spacing.beforeAutospacing ? 14 : (spacing.before ?? 0);
  const after = spacing.afterAutospacing ? 14 : (spacing.after ?? 0);
  if (props.contextualSpacing && state.styleId === para.styleId) before = 0;
  if (state.contextual && state.styleId === para.styleId && state.afterBox) state.afterBox.height = 0;

  const borders = props.borders;
  const hasBorders = !!borders && (isVisibleBorder(borders.top) || isVisibleBorder(borders.bottom) || isVisibleBorder(borders.left) || isVisibleBorder(borders.right));
  const grouped = hasBorders && sameBorders(state.borders, borders);
  const shading = shadingColor(props.shading);
  const boxes: Box[] = [];

  // Decoration bounds (borders sit outside the text indents).
  const decoLeft = left - (isVisibleBorder(borders?.left) ? borders!.left!.space + borders!.left!.width : 0);
  const decoRight = width - right + (isVisibleBorder(borders?.right) ? borders!.right!.space + borders!.right!.width : 0);

  const decorate = (p: Painter, x: number, y: number, h: number) => {
    if (shading) addItem(p, { type: 'rect', x: x + decoLeft, y, w: decoRight - decoLeft, h, fill: shading });
    if (hasBorders) {
      const lb = borders!.left;
      const rb = borders!.right;
      if (isVisibleBorder(lb)) borderLine(p, lb, x + decoLeft + lb.width / 2, y, x + decoLeft + lb.width / 2, y + h);
      if (isVisibleBorder(rb)) borderLine(p, rb, x + decoRight - rb.width / 2, y, x + decoRight - rb.width / 2, y + h);
    }
  };

  if (before > 0) {
    boxes.push({ height: before, keepWithNext: true, discardAtTop: true, paint: () => {} });
  }

  if (hasBorders) {
    if (grouped && state.bottomBorderBox) {
      // Paragraphs with identical borders form one box: drop the previous bottom border, draw "between".
      state.bottomBorderBox.height = 0;
      state.bottomBorderBox.paint = () => {};
      const between = borders!.between;
      if (isVisibleBorder(between)) {
        const h = between.space + between.width;
        boxes.push({
          height: h,
          keepWithNext: true,
          paint: (p, x, y) => {
            decorate(p, x, y, h);
            borderLine(p, between, x + decoLeft, y + between.width / 2, x + decoRight, y + between.width / 2);
          },
        });
      }
    } else if (isVisibleBorder(borders!.top)) {
      const top = borders!.top;
      const h = top.space + top.width;
      boxes.push({
        height: h,
        keepWithNext: true,
        paint: (p, x, y) => {
          decorate(p, x, y, h);
          borderLine(p, top, x + decoLeft, y + top.width / 2, x + decoRight, y + top.width / 2);
        },
      });
    }
  }

  const n = lines.length;
  const widow = props.widowControl ?? true;
  const outline = props.outlineLevel !== undefined && props.outlineLevel < 9 ? { level: props.outlineLevel, title: blocksToText([para]).replace(/\s+/g, ' ').trim() } : undefined;
  lines.forEach((line, i) => {
    const { height, baseline } = lineHeight(line, spacing.line, spacing.lineRule);
    const notes: NoteRef[] = [];
    for (const f of line.fragments) if (f.kind === 'mark' && (f.mark === 'footnote' || f.mark === 'endnote')) notes.push({ type: f.mark, id: f.name });
    let keep = false;
    if (props.keepLines && i < n - 1) keep = true;
    if (widow && n >= 2 && (i === 0 || i === n - 2)) keep = true;
    if (i === n - 1 && props.keepNext) keep = true;
    boxes.push({
      height,
      keepWithNext: keep,
      breakBefore: i === 0 && props.pageBreakBefore ? 'page' : undefined,
      breakAfter: line.breakAfter,
      notes: notes.length ? notes : undefined,
      paint: (p, x, y) => {
        decorate(p, x, y, height);
        if (i === 0 && outline?.title) p.page.headings.push({ ...outline, x: x + line.left, y });
        paintLine(p, line, x, { top: y, height, baseline: y + baseline });
      },
    });
  });
  if (props.pageBreakBefore && boxes[0]) boxes[0].breakBefore = 'page';

  let bottomBorderBox: Box | undefined;
  if (hasBorders && isVisibleBorder(borders!.bottom)) {
    const bottom = borders!.bottom;
    const h = bottom.space + bottom.width;
    bottomBorderBox = {
      height: h,
      keepWithNext: !!props.keepNext,
      paint: (p, x, y) => {
        decorate(p, x, y, h);
        borderLine(p, bottom, x + decoLeft, y + h - bottom.width / 2, x + decoRight, y + h - bottom.width / 2);
      },
    };
    boxes.push(bottomBorderBox);
  }

  const afterBox: Box = { height: after, keepWithNext: !!props.keepNext, canTruncate: true, paint: () => {} };
  boxes.push(afterBox);

  state.styleId = para.styleId;
  state.contextual = props.contextualSpacing;
  state.afterBox = afterBox;
  state.borders = hasBorders ? borders : undefined;
  state.bottomBorderBox = bottomBorderBox;
  return boxes;
}

export function lineHeight(line: Line, value: number | undefined, rule: 'auto' | 'exact' | 'atLeast' | undefined): { height: number; baseline: number } {
  const natural = line.ascent + line.descent;
  let height = natural;
  if (rule === 'exact' && value !== undefined) height = value;
  else if (rule === 'atLeast' && value !== undefined) height = Math.max(natural, value);
  else if (value !== undefined) height = natural * value;
  return { height, baseline: height - line.descent };
}

// ---------------------------------------------------------------------------------------------
// Tables

interface CellLayout {
  x: number;
  width: number;
  col: number;
  span: number;
  boxes: Box[];
  contentHeight: number;
  margins: { top: number; right: number; bottom: number; left: number };
  shading?: string;
  borders: { top?: Border; right?: Border; bottom?: Border; left?: Border };
  vMerge?: 'restart' | 'continue';
  vAlign: 'top' | 'center' | 'bottom';
  /** For a vMerge restart cell: number of rows it spans. */
  rowSpan: number;
}

interface RowLayout {
  height: number;
  cells: CellLayout[];
  isHeader: boolean;
  cantSplit: boolean;
  exact: boolean;
}

function heavier(a: Border | undefined, b: Border | undefined): Border | undefined {
  const av = isVisibleBorder(a) ? a : undefined;
  const bv = isVisibleBorder(b) ? b : undefined;
  if (!av) return bv ?? a;
  if (!bv) return av;
  return bv.width > av.width ? bv : av;
}

export function tableBoxes(table: DocTable, available: number, ctx: BlockContext): Box[] {
  const props = table.props;
  let grid = table.grid.length ? [...table.grid] : [];
  const maxCols = Math.max(1, ...table.rows.map((r) => r.cells.reduce((n, c) => n + c.gridSpan, 0) + (r.props.gridBefore ?? 0)));
  if (grid.length < maxCols) {
    const known = grid.reduce((a, b) => a + b, 0);
    const missing = maxCols - grid.length;
    const each = Math.max(18, (available - known) / missing);
    while (grid.length < maxCols) grid.push(each);
  }
  let total = grid.reduce((a, b) => a + b, 0);
  if (props.width?.type === 'pct' && props.width.value > 0) {
    const target = available * props.width.value;
    grid = grid.map((g) => (g * target) / total);
    total = target;
  } else if (total > available + 1 && props.layout !== 'fixed') {
    grid = grid.map((g) => (g * available) / total);
    total = available;
  }
  const colX: number[] = [0];
  for (const g of grid) colX.push(colX[colX.length - 1]! + g);
  let offset = props.indent ?? 0;
  if (props.jc === 'center') offset = (available - total) / 2;
  else if (props.jc === 'right') offset = available - total;

  const tb: TableBorders = props.borders ?? {};
  const defaultMargins = { top: props.cellMargins?.top ?? 0, right: props.cellMargins?.right ?? 5.4, bottom: props.cellMargins?.bottom ?? 0, left: props.cellMargins?.left ?? 5.4 };
  const nRows = table.rows.length;

  // First pass: geometry, borders and content of every cell.
  const rows: RowLayout[] = table.rows.map((row, r) => {
    let col = row.props.gridBefore ?? 0;
    const cells: CellLayout[] = row.cells.map((cell, ci) => {
      const span = Math.min(cell.gridSpan, Math.max(1, grid.length - col));
      const x = colX[col] ?? 0;
      const width = (colX[col + span] ?? total) - x;
      const cm = cell.props.margins ?? {};
      const margins = { top: cm.top ?? defaultMargins.top, right: cm.right ?? defaultMargins.right, bottom: cm.bottom ?? defaultMargins.bottom, left: cm.left ?? defaultMargins.left };
      const isFirstCol = col === 0;
      const isLastCol = ci === row.cells.length - 1;
      const cb = cell.props.borders ?? {};
      const layout: CellLayout = {
        x,
        width,
        col,
        span,
        boxes: [],
        contentHeight: 0,
        margins,
        shading: shadingColor(cell.props.shading),
        borders: {
          top: cb.top ?? (r === 0 ? tb.top : tb.insideH),
          bottom: cb.bottom ?? (r === nRows - 1 ? tb.bottom : tb.insideH),
          left: cb.left ?? (isFirstCol ? tb.left : tb.insideV),
          right: cb.right ?? (isLastCol ? tb.right : tb.insideV),
        },
        vMerge: cell.props.vMerge,
        vAlign: cell.props.vAlign ?? 'top',
        rowSpan: 1,
      };
      if (cell.props.vMerge !== 'continue') {
        layout.boxes = blockBoxes(cell.blocks, Math.max(1, width - margins.left - margins.right), ctx);
        layout.contentHeight = layout.boxes.reduce((h, b) => h + b.height, 0);
      }
      col += span;
      return layout;
    });
    const contentMax = Math.max(0, ...cells.filter((c) => c.vMerge !== 'continue' && c.vMerge !== 'restart').map((c) => c.contentHeight + c.margins.top + c.margins.bottom));
    const restartMin = Math.max(0, ...cells.filter((c) => c.vMerge === 'restart').map((c) => c.margins.top + c.margins.bottom));
    let height = Math.max(contentMax, restartMin, 1);
    const rule = row.props.heightRule ?? 'atLeast';
    if (row.props.height !== undefined) {
      if (rule === 'exact') height = row.props.height;
      else if (rule === 'atLeast') height = Math.max(height, row.props.height);
    }
    return { height, cells, isHeader: !!row.props.isHeader, cantSplit: !!row.props.cantSplit, exact: rule === 'exact' && row.props.height !== undefined };
  });

  // Vertical merges: make the spanned rows tall enough for the merged content.
  rows.forEach((row, r) => {
    for (const cell of row.cells) {
      if (cell.vMerge !== 'restart') continue;
      let span = 1;
      while (r + span < rows.length && rows[r + span]!.cells.some((c) => c.col === cell.col && c.vMerge === 'continue')) span++;
      cell.rowSpan = span;
      const need = cell.contentHeight + cell.margins.top + cell.margins.bottom;
      const have = rows.slice(r, r + span).reduce((h, rr) => h + rr.height, 0);
      if (need > have) rows[r + span - 1]!.height += need - have;
    }
  });

  // Conflict resolution with neighbours: the heavier border wins on a shared edge.
  rows.forEach((row, r) => {
    for (const cell of row.cells) {
      const above = rows[r - 1]?.cells.find((c) => c.col <= cell.col && cell.col < c.col + c.span);
      if (above) cell.borders.top = heavier(cell.borders.top, above.borders.bottom);
      const leftCell = row.cells.find((c) => c.col + c.span === cell.col);
      if (leftCell) cell.borders.left = heavier(cell.borders.left, leftCell.borders.right);
      if (cell.vMerge === 'continue') cell.borders.top = undefined;
    }
  });
  rows.forEach((row, r) => {
    for (const cell of row.cells) {
      const below = rows[r + 1]?.cells.find((c) => c.col <= cell.col && cell.col < c.col + c.span);
      if (below?.vMerge === 'continue' && (cell.vMerge === 'restart' || cell.vMerge === 'continue')) cell.borders.bottom = undefined;
    }
  });

  const rowBox = (row: RowLayout, r: number, cellBoxes?: Map<CellLayout, Box[]>, heightOverride?: number, continuation = false): Box => {
    const height = heightOverride ?? row.height;
    const box: Box = {
      height,
      keepWithNext: false,
      paint: (p, x, y) => paintRow(p, row, r, x + offset, y, height, rows, cellBoxes, continuation),
    };
    if (!row.cantSplit && !row.exact) {
      box.split = (avail) => splitRow(row, r, avail, cellBoxes);
    }
    return box;
  };

  const splitRow = (row: RowLayout, r: number, avail: number, cellBoxes?: Map<CellLayout, Box[]>): [Box, Box] | undefined => {
    const firstParts = new Map<CellLayout, Box[]>();
    const rest = new Map<CellLayout, Box[]>();
    let firstHeight = 0;
    let restHeight = 0;
    let placedAny = false;
    for (const cell of row.cells) {
      const boxes = cellBoxes?.get(cell) ?? cell.boxes;
      const limit = avail - cell.margins.top - cell.margins.bottom;
      let used = 0;
      let i = 0;
      while (i < boxes.length && used + boxes[i]!.height <= limit) used += boxes[i++]!.height;
      if (i > 0) placedAny = true;
      firstParts.set(cell, boxes.slice(0, i));
      rest.set(cell, boxes.slice(i));
      firstHeight = Math.max(firstHeight, used + cell.margins.top + cell.margins.bottom);
      const remaining = boxes.slice(i).reduce((h, b) => h + b.height, 0);
      restHeight = Math.max(restHeight, remaining + cell.margins.top + cell.margins.bottom);
    }
    if (!placedAny || restHeight <= 0) return undefined;
    return [rowBox(row, r, firstParts, Math.min(avail, Math.max(firstHeight, 1))), rowBox(row, r, rest, Math.max(restHeight, 1), true)];
  };

  const boxes = rows.map((row, r) => rowBox(row, r));
  const headerCount = rows.findIndex((row) => !row.isHeader);
  const headers = headerCount > 0 ? boxes.slice(0, headerCount) : [];
  for (let r = headerCount > 0 ? headerCount : 1; r < boxes.length; r++) if (headers.length) boxes[r]!.repeatHeader = headers;
  // Keep header rows with the first body row.
  for (let r = 0; r < headerCount && r < boxes.length - 1; r++) boxes[r]!.keepWithNext = true;
  return boxes;
}

function paintRow(p: Painter, row: RowLayout, r: number, x0: number, y: number, height: number, rows: RowLayout[], cellBoxes: Map<CellLayout, Box[]> | undefined, continuation: boolean): void {
  for (const cell of row.cells) {
    const x = x0 + cell.x;
    if (cell.shading) addItem(p, { type: 'rect', x, y, w: cell.width, h: height, fill: cell.shading });
  }
  for (const cell of row.cells) {
    if (cell.vMerge === 'continue') continue;
    const x = x0 + cell.x;
    const boxes = cellBoxes?.get(cell) ?? cell.boxes;
    const contentHeight = boxes.reduce((h, b) => h + b.height, 0);
    let spanHeight = height;
    if (cell.vMerge === 'restart' && !cellBoxes) {
      for (let k = 1; k < cell.rowSpan; k++) spanHeight += rows[r + k]?.height ?? 0;
    }
    let cy = y + cell.margins.top;
    const free = spanHeight - cell.margins.top - cell.margins.bottom - contentHeight;
    if (free > 0 && !continuation) {
      if (cell.vAlign === 'center') cy += free / 2;
      else if (cell.vAlign === 'bottom') cy += free;
    }
    const savedColumn = p.column;
    p.column = { x: x + cell.margins.left, width: cell.width - cell.margins.left - cell.margins.right };
    for (const b of boxes) {
      b.paint(p, x + cell.margins.left, cy);
      cy += b.height;
    }
    p.column = savedColumn;
  }
  for (const cell of row.cells) {
    const x = x0 + cell.x;
    const { top, bottom, left, right } = cell.borders;
    borderLine(p, top, x, y, x + cell.width, y);
    borderLine(p, bottom, x, y + height, x + cell.width, y + height);
    borderLine(p, left, x, y, x, y + height);
    borderLine(p, right, x + cell.width, y, x + cell.width, y + height);
  }
}
