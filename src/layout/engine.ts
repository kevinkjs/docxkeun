/**
 * Pagination: stacks boxes into columns and pages per section, places headers, footers and
 * footnotes, resolves page-number fields, and collects anchors for links and comments.
 */
import { FontRegistry } from '../fonts/registry.js';
import type { Block, FloatInfo, NoteRefInline } from '../model/content.js';
import type { DocxDocument } from '../model/document.js';
import { formatNumber, NumberingCounter } from '../model/numbering.js';
import { parseSectionProps, type SectionProps } from '../model/props.js';
import { blockBoxes, type BlockContext, type Box, type NoteRef, type ParagraphMeta } from './blocks.js';
import type { LayoutCache } from './cache.js';
import { createDrawingContext } from './shapes.js';
import type { DrawingContext } from './drawing.js';
import { RegistryFontProvider } from './fonts.js';
import { RangeTracker } from './lines.js';
import { floatPosition, measureChars, pageNumberLabel, type Painter } from './paint.js';
import type { LineBounds } from './lines.js';
import type { DisplayItem, FontProvider, LayoutResult, Page, TextItem } from './types.js';

export interface LayoutOptions {
  /** Registered fonts (or a custom provider, e.g. canvas metrics in a browser). */
  fonts: FontRegistry | FontProvider;
  /** Re-use the layout of unchanged paragraphs from earlier runs (keep one per document). */
  cache?: LayoutCache;
}

const EPS = 0.01;

const providers = new WeakMap<FontRegistry, RegistryFontProvider>();
/** One provider per registry, so cached paragraph layout (which holds its fonts) stays valid. */
function providerFor(registry: FontRegistry): RegistryFontProvider {
  let p = providers.get(registry);
  if (!p) providers.set(registry, (p = new RegistryFontProvider(registry)));
  return p;
}

interface SectionGroup {
  props: SectionProps;
  blocks: Block[];
}

function splitSections(blocks: Block[], finalSection: SectionProps | undefined): SectionGroup[] {
  const groups: SectionGroup[] = [];
  let current: Block[] = [];
  for (const b of blocks) {
    current.push(b);
    if (b.kind === 'paragraph' && b.sectionEnd) {
      groups.push({ props: b.sectionEnd, blocks: current });
      current = [];
    }
  }
  groups.push({ props: finalSection ?? groups[groups.length - 1]?.props ?? parseSectionProps(undefined), blocks: current });
  // Header/footer references are inherited from the previous section when missing.
  for (let i = 1; i < groups.length; i++) {
    const prev = groups[i - 1]!.props;
    const cur = groups[i]!.props;
    for (const kind of ['headers', 'footers'] as const) {
      for (const ref of prev[kind]) if (!cur[kind].some((r) => r.type === ref.type)) cur[kind] = [...cur[kind], ref];
    }
  }
  return groups;
}

interface Placement {
  box: Box;
  index: number;
  x: number;
  y: number;
  column: number;
  atTop: boolean;
}

class Paginator {
  readonly pages: Page[] = [];
  page!: Page;
  painter!: Painter;
  section!: SectionProps;
  sectionIndex = -1;
  columns: Array<{ x: number; width: number }> = [];
  col = 0;
  regionTop = 0;
  /** First placement of the current column region (placements restart on every page). */
  regionStart = 0;
  bodyTop = 0;
  bottom = 0;
  y = 0;
  atColumnTop = true;
  startedByOverflow = false;
  pageNumber = 0;
  placements: Placement[] = [];
  /** Areas text must flow around on this page (floating objects with wrapping, plus their distance from text). */
  exclusions: Array<{ left: number; top: number; right: number; bottom: number; wrap: FloatInfo['wrap']; side: FloatInfo['wrapText'] }> = [];
  /** Paragraphs whose floats are already placed on this page. */
  private floatsPlaced = new Set<ParagraphMeta>();
  notes: Array<{ ref: NoteRef; boxes: Box[]; height: number }> = [];
  notesHeight = 0;
  readonly endnotes: NoteRef[] = [];
  readonly sectionPageCounts: number[] = [];

  constructor(private readonly engine: LayoutEngine) {}

  get contentWidth(): number {
    const m = this.section.margin;
    return this.section.pageWidth - m.left - m.right - m.gutter;
  }

  private computeColumns(section: SectionProps): Array<{ x: number; width: number }> {
    const m = section.margin;
    const left = m.left + m.gutter;
    const width = section.pageWidth - m.left - m.right - m.gutter;
    const c = section.columns;
    if (c.count <= 1) return [{ x: left, width }];
    if (!c.equalWidth && c.widths.length >= c.count) {
      let x = left;
      return c.widths.map((w) => {
        const col = { x, width: w.width };
        x += w.width + w.space;
        return col;
      });
    }
    const colWidth = (width - c.space * (c.count - 1)) / c.count;
    return Array.from({ length: c.count }, (_, i) => ({ x: left + i * (colWidth + c.space), width: colWidth }));
  }

  startSection(section: SectionProps, index: number): void {
    const first = this.sectionIndex < 0;
    const prev = this.section;
    this.section = section;
    this.sectionIndex = index;
    this.sectionPageCounts[index] = 0;
    const samePage = prev && prev.pageWidth === section.pageWidth && prev.pageHeight === section.pageHeight;
    if (!first && section.type === 'continuous' && samePage) {
      // Word balances the columns of a section that ends with a continuous break.
      if (this.columns.length > 1) this.balanceColumns();
      // New column layout on the same page, below the content so far.
      this.flushColumnsTo();
      this.columns = this.computeColumns(section);
      this.col = 0;
      this.regionTop = this.y;
      this.regionStart = this.placements.length;
      this.atColumnTop = true;
      this.sectionPageCounts[index] = 1;
      return;
    }
    if (!first && section.type === 'nextColumn' && samePage) {
      this.nextColumn(false);
      return;
    }
    if (section.pageNumberStart !== undefined) this.pageNumber = section.pageNumberStart - 1;
    if (!first && (section.type === 'evenPage' || section.type === 'oddPage')) {
      const nextNumber = this.pageNumber + 1;
      const wantEven = section.type === 'evenPage';
      if ((nextNumber % 2 === 0) !== wantEven) this.newPage(false);
    }
    this.newPage(false);
  }

  /**
   * Redistribute the current page's part of a multi-column section so its columns end at about the
   * same height: find the smallest column height that holds everything (splitting paragraphs and
   * tables between lines and rows as usual), then place the boxes again with that height.
   */
  private balanceColumns(): void {
    const n = this.columns.length;
    const region = this.placements.slice(this.regionStart);
    const boxes = region.filter((p) => p.index >= 0).map((p) => p.box);
    if (boxes.length === 0) return;
    // Explicit column or page breaks decide the columns themselves.
    if (boxes.some((b) => b.breakBefore || b.breakAfter)) return;
    const total = boxes.reduce((sum, b) => sum + b.height, 0);
    const fits = (h: number): boolean => {
      const queue = [...boxes];
      let col = 0;
      let used = 0;
      for (let i = 0; i < queue.length; i++) {
        const b = queue[i]!;
        const height = used === 0 && b.discardAtTop ? 0 : b.height;
        if (used + height <= h + EPS) {
          used += height;
          continue;
        }
        const parts = b.split && h - used > 12 ? b.split(h - used) : undefined;
        if (parts) {
          queue.splice(i, 1, parts[0], parts[1]);
        } else {
          if (used === 0) return false;
          i--;
        }
        col++;
        used = 0;
        if (col >= n) return false;
      }
      return true;
    };
    let hi = this.y - this.regionTop;
    for (const p of region) hi = Math.max(hi, p.y + p.box.height - this.regionTop);
    if (!fits(hi)) return;
    let lo = total / n;
    for (let k = 0; k < 24 && hi - lo > 0.25; k++) {
      const mid = (lo + hi) / 2;
      if (fits(mid)) hi = mid;
      else lo = mid;
    }
    const page = this.page;
    const savedBottom = this.bottom;
    this.placements.length = this.regionStart;
    this.col = 0;
    this.y = this.regionTop;
    this.atColumnTop = true;
    this.bottom = this.regionTop + hi + EPS + this.notesHeight;
    this.run(boxes);
    if (this.page === page) this.bottom = savedBottom;
  }

  /** All columns of the current region are done: continue below the tallest. */
  private flushColumnsTo(): void {
    let maxY = this.y;
    for (const p of this.placements) maxY = Math.max(maxY, p.y + p.box.height);
    this.y = maxY;
  }

  newPage(overflow: boolean): void {
    if (this.page) this.finishPage();
    this.pageNumber++;
    const s = this.section;
    const index = this.pages.length;
    const page: Page = {
      index,
      number: this.pageNumber,
      label: pageNumberLabel(this.pageNumber, s.pageNumberFormat),
      width: s.pageWidth,
      height: s.pageHeight,
      section: s,
      sectionIndex: this.sectionIndex,
      items: [],
      links: [],
      comments: [],
      perms: [],
      locks: [],
      anchors: new Map(),
      lines: [],
      storyLines: [],
      headings: [],
      content: { x: 0, y: 0, w: 0, h: 0 },
    };
    this.pages.push(page);
    this.sectionPageCounts[this.sectionIndex] = (this.sectionPageCounts[this.sectionIndex] ?? 0) + 1;
    this.page = page;
    this.columns = this.computeColumns(s);
    this.painter = this.engine.painterFor(page, this.columns[0]!);
    const firstOfSection = this.sectionPageCounts[this.sectionIndex] === 1;
    const { headerBottom, footerTop } = this.engine.paintHeaderFooter(this.painter, s, page, firstOfSection);
    const m = s.margin;
    this.bodyTop = Math.max(m.top, headerBottom);
    this.bottom = Math.min(s.pageHeight - m.bottom, footerTop);
    page.content = { x: m.left + m.gutter, y: this.bodyTop, w: this.contentWidth, h: this.bottom - this.bodyTop };
    this.col = 0;
    this.regionTop = this.bodyTop;
    this.regionStart = 0;
    this.y = this.bodyTop;
    this.atColumnTop = true;
    this.startedByOverflow = overflow;
    this.placements = [];
    this.exclusions = [];
    this.floatsPlaced = new Set();
    this.notes = [];
    this.notesHeight = 0;
  }

  /** Place a paragraph's floating objects (anchored at `paraTop`) and record the areas text wraps around. */
  private placeFloats(meta: ParagraphMeta, paraTop: number): void {
    if (this.floatsPlaced.has(meta)) return;
    this.floatsPlaced.add(meta);
    const column = this.columns[this.col]!;
    for (const image of meta.floats) {
      const f = image.float!;
      const at = floatPosition(this.page, column, f, image, column.x, paraTop);
      this.painter.floatRects?.set(image, at);
      if (f.wrap === 'none') continue;
      this.exclusions.push({
        left: at.x - f.dist.left,
        top: at.y - f.dist.top,
        right: at.x + image.width + f.dist.right,
        bottom: at.y + image.height + f.dist.bottom,
        wrap: f.wrap,
        side: f.wrapText,
      });
    }
  }

  /** Room for a line at an absolute height in the current column, around this page's exclusions. */
  private boundsAt(column: { x: number; width: number }, top: number, height: number): LineBounds | undefined {
    let left = 0;
    let right = column.width;
    let next: number | undefined;
    let touched = false;
    for (const e of this.exclusions) {
      if (e.bottom <= top + EPS || e.top >= top + height - EPS) continue;
      const exLeft = e.left - column.x;
      const exRight = e.right - column.x;
      if (exRight <= left || exLeft >= right) continue;
      touched = true;
      if (e.wrap === 'topAndBottom') {
        right = left;
      } else {
        const roomLeft = exLeft - left;
        const roomRight = right - exRight;
        const textLeft = e.side === 'left' || ((e.side === 'bothSides' || e.side === 'largest') && roomLeft >= roomRight);
        if (textLeft) right = Math.min(right, exLeft);
        else left = Math.max(left, exRight);
      }
      next = Math.max(next ?? 0, e.bottom);
    }
    if (!touched) return undefined;
    return { left, right, ...(next !== undefined ? { next } : {}) };
  }

  /** Lay a paragraph out around this page's floats when it overlaps any; returns the boxes to use. */
  private wrapParagraph(queue: Box[], i: number): boolean {
    const meta = queue[i]!.paragraph;
    if (!meta || meta.wrapped) return false;
    const paraTop = this.y + meta.textOffset;
    this.placeFloats(meta, paraTop);
    if (!this.exclusions.length) return false;
    let height = 0;
    for (let k = i; k < i + meta.count && k < queue.length; k++) height += queue[k]!.height;
    const column = this.columns[this.col]!;
    const overlaps = this.exclusions.some((e) => e.bottom > this.y && e.top < this.y + height && e.right > column.x && e.left < column.x + column.width);
    if (!overlaps) return false;
    const boxes = meta.reflow((top, h) => {
      const b = this.boundsAt(column, paraTop + top, h);
      return b && { left: b.left, right: b.right, ...(b.next !== undefined ? { next: b.next - paraTop } : {}) };
    });
    queue.splice(i, meta.count, ...boxes);
    return true;
  }

  nextColumn(overflow: boolean): void {
    if (this.col < this.columns.length - 1) {
      this.col++;
      this.y = this.regionTop;
      this.atColumnTop = true;
      this.startedByOverflow = overflow;
      return;
    }
    this.newPage(overflow);
  }

  breakTo(kind: 'page' | 'column'): void {
    if (kind === 'column') this.nextColumn(false);
    else this.newPage(false);
  }

  private notesFor(box: Box): { refs: Array<{ ref: NoteRef; boxes: Box[]; height: number }>; height: number } {
    const refs: Array<{ ref: NoteRef; boxes: Box[]; height: number }> = [];
    let height = 0;
    for (const ref of box.notes ?? []) {
      if (ref.type === 'endnote') continue;
      if (this.notes.some((n) => n.ref.id === ref.id)) continue;
      const laid = this.engine.noteBoxes(ref, this.columns[this.col]!.width);
      refs.push({ ref, ...laid });
      height += laid.height;
    }
    if (refs.length && this.notes.length === 0) height += NOTE_SEPARATOR_SPACE;
    return { refs, height };
  }

  place(box: Box, index: number, height = box.height): void {
    const column = this.columns[this.col]!;
    this.placements.push({ box, index, x: column.x, y: this.y, column: this.col, atTop: this.atColumnTop });
    if (height !== box.height) {
      // Truncated or discarded spacing: paint with the reduced height.
      const original = box;
      box = { ...original, height };
      this.placements[this.placements.length - 1]!.box = box;
    }
    for (const ref of box.notes ?? []) if (ref.type === 'endnote' && !this.endnotes.some((e) => e.id === ref.id)) this.endnotes.push(ref);
    this.y += height;
    if (height > 0 || box.height > 0) this.atColumnTop = false;
  }

  run(boxes: Box[]): void {
    const queue = [...boxes];
    const backtracked = new Set<number>();
    let i = 0;
    let guard = 0;
    while (i < queue.length) {
      if (++guard > queue.length * 20 + 10000) throw new Error('Layout did not converge');
      if (queue[i]!.breakBefore && !this.atColumnTop) this.breakTo(queue[i]!.breakBefore!);
      // Text wraps around floating objects: re-break the paragraph where it meets them.
      this.wrapParagraph(queue, i);
      const box = queue[i]!;

      // Continuing a table on a new page or column: repeat its header rows.
      if (box.repeatHeader && this.atColumnTop) {
        const prev = queue[i - 1];
        if (prev && (prev.repeatHeader === box.repeatHeader || box.repeatHeader.includes(prev))) {
          for (const h of box.repeatHeader) this.place(h, -1);
        }
      }

      let height = box.height;
      if (box.discardAtTop && this.atColumnTop && this.startedByOverflow) height = 0;
      const notes = this.notesFor(box);
      const avail = this.bottom - this.notesHeight - this.y;

      if (height + notes.height <= avail + EPS) {
        this.place(box, i, height);
        this.addNotes(notes);
      } else if (box.canTruncate) {
        this.place(box, i, Math.max(0, Math.min(height, avail)));
      } else {
        // Move a keep-together chain to the next page if it does not start at the top.
        let chain = i;
        while (chain > 0 && queue[chain - 1]!.keepWithNext) chain--;
        const start = this.placements.find((p) => p.index === chain);
        if (chain < i && start && start.column === this.col && !start.atTop && !backtracked.has(chain)) {
          backtracked.add(chain);
          this.placements = this.placements.filter((p) => p.index < chain || (p.index === -1 && p.y < start.y));
          this.y = start.y;
          this.nextColumn(true);
          i = chain;
          continue;
        }
        const splittable = box.split && avail - notes.height > 12;
        const parts = splittable ? box.split!(avail - notes.height) : undefined;
        if (parts) {
          parts[0].repeatHeader = box.repeatHeader;
          parts[1].repeatHeader = box.repeatHeader;
          queue.splice(i, 1, parts[0], parts[1]);
          this.place(parts[0], i);
          this.addNotes(notes);
          this.nextColumn(true);
          i++;
          continue;
        }
        if (this.atColumnTop) {
          // Taller than a whole page: place it anyway (it will be clipped).
          this.place(box, i, height);
          this.addNotes(notes);
        } else {
          this.nextColumn(true);
          continue;
        }
      }
      if (box.breakAfter) this.breakTo(box.breakAfter);
      i++;
    }
  }

  private addNotes(notes: { refs: Array<{ ref: NoteRef; boxes: Box[]; height: number }>; height: number }): void {
    if (!notes.refs.length) return;
    this.notes.push(...notes.refs);
    this.notesHeight += notes.height;
  }

  /** Paint everything placed on the current page. */
  finishPage(): void {
    const p = this.painter;
    for (const pl of this.placements) {
      p.column = this.columns[pl.column] ?? this.columns[0]!;
      pl.box.paint(p, pl.x, pl.y);
    }
    if (this.notes.length) {
      p.story = 'footnote';
      const column = this.columns[0]!;
      let y = this.bottom - this.notesHeight;
      p.page.items.push({ type: 'line', x1: column.x, y1: y + NOTE_SEPARATOR_SPACE / 2, x2: column.x + Math.min(144, column.width / 3), y2: y + NOTE_SEPARATOR_SPACE / 2, width: 0.5, color: '000000', style: 'solid' });
      y += NOTE_SEPARATOR_SPACE;
      for (const n of this.notes) {
        p.story = n.ref.type;
        p.storyRef = { part: this.engine.notePart(n.ref.type) ?? '', noteId: n.ref.id };
        for (const b of n.boxes) {
          b.paint(p, column.x, y);
          y += b.height;
        }
      }
    }
    p.story = 'body';
    p.storyRef = undefined;
    if (p.behind.length) p.page.items.unshift(...p.behind);
    this.placements = [];
  }
}

const NOTE_SEPARATOR_SPACE = 9;

class LayoutEngine {
  readonly fonts: FontProvider;
  private readonly tracker = new RangeTracker();
  private readonly counter: NumberingCounter;
  private readonly noteNumbers = { footnote: new Map<string, number>(), endnote: new Map<string, number>() };
  private currentNote: { type: 'footnote' | 'endnote'; id: string } | undefined;
  private readonly noteCache = new Map<string, { boxes: Box[]; height: number }>();
  private numPagesPlaceholder = '0';
  private readonly cache: LayoutCache | undefined;
  private readonly faceCount: number;

  constructor(private readonly doc: DocxDocument, options: LayoutOptions) {
    this.fonts = options.fonts instanceof FontRegistry ? providerFor(options.fonts) : options.fonts;
    this.cache = options.cache;
    this.faceCount = options.fonts instanceof FontRegistry ? Array.from(options.fonts.faces()).length : 0;
    this.counter = new NumberingCounter(doc.numbering, (s) => doc.styles.get(s)?.pPr.numId);
  }

  private noteLabel = (inline: NoteRefInline): string => {
    if (!inline.id) {
      const cur = this.currentNote;
      if (!cur) return '';
      return this.formatNote(cur.type, this.noteNumbers[cur.type].get(cur.id) ?? 0);
    }
    const map = this.noteNumbers[inline.noteType];
    let n = map.get(inline.id);
    if (n === undefined) {
      n = map.size + 1;
      map.set(inline.id, n);
    }
    return this.formatNote(inline.noteType, n);
  };

  private drawingCtx: DrawingContext | undefined;

  /** Document and fonts for painting drawings (shapes, charts, metafiles). */
  drawingContext(): DrawingContext {
    return (this.drawingCtx ??= createDrawingContext(this.doc, this.fonts, (blocks, width) => blockBoxes(blocks, width, this.context(new RangeTracker(), new NumberingCounter(this.doc.numbering)))));
  }

  /** Part holding footnotes or endnotes. */
  notePart(type: 'footnote' | 'endnote'): string | undefined {
    return type === 'footnote' ? this.doc.parts.footnotes : this.doc.parts.endnotes;
  }

  private formatNote(type: 'footnote' | 'endnote', n: number): string {
    return formatNumber(n, type === 'footnote' ? 'decimal' : 'lowerRoman');
  }

  private context(tracker = this.tracker, counter = this.counter): BlockContext {
    return { fonts: this.fonts, theme: this.doc.theme, tracker, counter, defaultTab: this.doc.defaultTabStop, noteLabel: this.noteLabel, cache: this.cache };
  }

  painterFor(page: Page, column: { x: number; width: number }): Painter {
    return {
      page,
      behind: [],
      column,
      story: 'body',
      floatRects: new Map(),
      drawing: this.drawingContext(),
      fieldValue: (field) => {
        if (field === 'PAGE') return page.label;
        return this.numPagesPlaceholder;
      },
    };
  }

  noteBoxes(ref: NoteRef, width: number): { boxes: Box[]; height: number } {
    const key = `${ref.type}:${ref.id}`;
    let cached = this.noteCache.get(key);
    if (!cached) {
      this.currentNote = ref;
      const ctx = this.context(new RangeTracker(), new NumberingCounter(this.doc.numbering));
      const boxes = blockBoxes(this.doc.note(ref.type, ref.id), width, ctx);
      this.currentNote = undefined;
      cached = { boxes, height: boxes.reduce((h, b) => h + b.height, 0) };
      this.noteCache.set(key, cached);
    }
    return cached;
  }

  paintHeaderFooter(p: Painter, s: SectionProps, page: Page, firstOfSection: boolean): { headerBottom: number; footerTop: number } {
    const m = s.margin;
    const width = s.pageWidth - m.left - m.right - m.gutter;
    const x = m.left + m.gutter;
    const pick = (refs: SectionProps['headers']) => {
      let type: 'default' | 'first' | 'even' = 'default';
      if (s.titlePage && firstOfSection) type = 'first';
      else if (this.doc.evenAndOddHeaders && page.number % 2 === 0) type = 'even';
      const ref = refs.find((r) => r.type === type);
      if (!ref && type !== 'default') return undefined; // a missing first/even header is blank
      return ref ?? refs.find((r) => r.type === 'default');
    };
    const story = (ref: { rId: string } | undefined): Box[] => {
      if (!ref) return [];
      const part = this.doc.headerFooterPart(ref.rId);
      if (!part) return [];
      const ctx = this.context(new RangeTracker(), new NumberingCounter(this.doc.numbering));
      return blockBoxes(this.doc.headerFooter(part).blocks, width, ctx);
    };
    const saved = p.column;
    p.column = { x, width };
    const headerRef = pick(s.headers);
    p.story = 'header';
    p.storyRef = headerRef ? { part: this.doc.headerFooterPart(headerRef.rId) ?? '' } : undefined;
    let y = m.header;
    for (const b of story(headerRef)) {
      b.paint(p, x, y);
      y += b.height;
    }
    const headerBottom = y;
    const footerRef = pick(s.footers);
    p.story = 'footer';
    p.storyRef = footerRef ? { part: this.doc.headerFooterPart(footerRef.rId) ?? '' } : undefined;
    const footer = story(footerRef);
    const footerHeight = footer.reduce((h, b) => h + b.height, 0);
    let fy = s.pageHeight - m.footer - footerHeight;
    const footerTop = fy;
    for (const b of footer) {
      b.paint(p, x, fy);
      fy += b.height;
    }
    p.column = saved;
    p.story = 'body';
    p.storyRef = undefined;
    return { headerBottom, footerTop: Math.min(footerTop, s.pageHeight - m.footer) };
  }

  run(): LayoutResult {
    this.cache?.begin([this.fonts, this.faceCount, this.doc.theme, this.doc.styles, this.doc.numbering]);
    const story = this.doc.body();
    const groups = splitSections(story.blocks, story.finalSection);
    const paginator = new Paginator(this);
    groups.forEach((group, index) => {
      paginator.startSection(group.props, index);
      const width = Math.min(...paginator.columns.map((c) => c.width));
      paginator.run(blockBoxes(group.blocks, width, this.context()));
    });
    if (paginator.endnotes.length) {
      const width = paginator.columns[0]!.width;
      const boxes: Box[] = [{ height: NOTE_SEPARATOR_SPACE * 2, keepWithNext: true, paint: (p, x, y) => p.page.items.push({ type: 'line', x1: x, y1: y + NOTE_SEPARATOR_SPACE, x2: x + Math.min(144, width / 3), y2: y + NOTE_SEPARATOR_SPACE, width: 0.5, color: '000000', style: 'solid' }) }];
      for (const ref of paginator.endnotes) {
        const storyRef = { part: this.notePart('endnote') ?? '', noteId: ref.id };
        // Endnotes flow like body text but are their own story for editing.
        for (const b of this.noteBoxes(ref, width).boxes) {
          boxes.push({
            ...b,
            paint: (p, x, y) => {
              p.story = 'endnote';
              p.storyRef = storyRef;
              b.paint(p, x, y);
              p.story = 'body';
              p.storyRef = undefined;
            },
          });
        }
      }
      paginator.run(boxes);
    }
    paginator.finishPage();
    const pages = paginator.pages;
    this.resolveCountFields(pages, paginator.sectionPageCounts);

    const result: LayoutResult = { pages, anchors: new Map(), commentAnchors: new Map() };
    for (const page of pages) {
      for (const [name, pos] of page.anchors) if (!result.anchors.has(name)) result.anchors.set(name, { page: page.index, ...pos });
      for (const c of page.comments) if (!result.commentAnchors.has(c.id)) result.commentAnchors.set(c.id, { page: page.index, rect: { x: c.x, y: c.y, w: c.w, h: c.h } });
    }
    return result;
  }

  /** NUMPAGES / SECTIONPAGES are only known after pagination. */
  private resolveCountFields(pages: Page[], sectionCounts: number[]): void {
    for (const page of pages) {
      for (let i = 0; i < page.items.length; i++) {
        const item = page.items[i] as DisplayItem;
        if (item.type !== 'text' || (item.field !== 'NUMPAGES' && item.field !== 'SECTIONPAGES')) continue;
        const value = String(item.field === 'NUMPAGES' ? pages.length : (sectionCounts[page.sectionIndex] ?? 0));
        const m = measureChars(value, item.font, item.size, item.x);
        const updated: TextItem = { ...item, text: value, chars: m.chars, xs: m.xs, widths: m.widths, glyphs: m.chars.map((c) => item.font.glyphId(c.codePointAt(0)!)) };
        page.items[i] = updated;
      }
    }
  }
}

/** Lay out a document into pages. */
export function layoutDocument(doc: DocxDocument, options: LayoutOptions): LayoutResult {
  return new LayoutEngine(doc, options).run();
}
