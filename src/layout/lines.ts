/**
 * Paragraph text → lines: turns inlines into measured atoms, breaks them into lines (greedy,
 * like Word), resolves tab stops, and applies alignment and justification.
 */
import type { DocParagraph, ImageInline, Inline, NoteRefInline, TextInline } from '../model/content.js';
import type { LinkTarget } from '../model/fields.js';
import type { RunProps, TabStop } from '../model/props.js';
import type { Theme } from '../model/theme.js';
import { fontNameForSlot, isEastAsian, resolveRunStyle, scriptSlot, type RunStyle } from './style.js';
import { layoutMath } from '../math/index.js';
import { mirrorChar } from '../text/bidi.js';
import { shape } from '../fonts/shaping/index.js';
import { analyzeBidi, bidiOf, visualAtoms } from './bidi.js';
import { attr, firstChild, type XmlElement } from '../xml/index.js';
import type { DisplayItem, FontProvider, LayoutFont } from './types.js';

// ---------------------------------------------------------------------------------------------
// Range state (links, comments, permission ranges) that spans inlines and paragraphs

export interface PermInfo {
  id: string;
  editor?: string;
  group?: string;
}

export interface LockInfo {
  id: string;
  title?: string;
}

export interface RangeCtx {
  readonly link?: LinkTarget;
  readonly comments: readonly string[];
  readonly perms: readonly PermInfo[];
  readonly locks: readonly LockInfo[];
}

export const EMPTY_CTX: RangeCtx = Object.freeze({ comments: [], perms: [], locks: [] });

export class RangeTracker {
  ctx: RangeCtx = EMPTY_CTX;

  /** Apply a marker inline; returns true if it was a marker. */
  apply(inline: Inline): boolean {
    const c = this.ctx;
    switch (inline.kind) {
      case 'linkStart':
        this.ctx = { ...c, link: inline.target };
        return true;
      case 'linkEnd':
        this.ctx = { ...c, link: undefined };
        return true;
      case 'commentStart':
        if (!c.comments.includes(inline.id)) this.ctx = { ...c, comments: [...c.comments, inline.id] };
        return true;
      case 'commentEnd':
        this.ctx = { ...c, comments: c.comments.filter((id) => id !== inline.id) };
        return true;
      case 'permStart':
        if (!c.perms.some((p) => p.id === inline.id)) this.ctx = { ...c, perms: [...c.perms, { id: inline.id, editor: inline.editor, group: inline.group }] };
        return true;
      case 'permEnd':
        this.ctx = { ...c, perms: c.perms.filter((p) => p.id !== inline.id) };
        return true;
      case 'lockStart':
        this.ctx = { ...c, locks: [...c.locks, { id: inline.id, title: inline.title }] };
        return true;
      case 'lockEnd': {
        // Remove one level (nested controls may share an id when ids are missing).
        const i = c.locks.map((l) => l.id).lastIndexOf(inline.id);
        if (i >= 0) this.ctx = { ...c, locks: c.locks.filter((_, j) => j !== i) };
        return true;
      }
      case 'bookmarkStart':
      case 'bookmarkEnd':
      case 'commentRef':
        return true;
      default:
        return false;
    }
  }
}

// ---------------------------------------------------------------------------------------------
// Atoms

interface AtomBase {
  ctx: RangeCtx;
}
/** One glyph from the shaping engine, in points: position relative to the pen, and its advance. */
export interface PlacedGlyph {
  id: number;
  dx: number;
  dy: number;
  advance: number;
  /** Text the glyph stands for (a ligature covers several characters), for copy and search in PDFs. */
  text: string;
}

export interface CharAtom extends AtomBase {
  kind: 'char';
  ch: string;
  /** Glyphs from the shaping engine (complex scripts, kerning, ligatures); absent: one glyph per code point. */
  glyphs?: PlacedGlyph[];
  width: number;
  font: LayoutFont;
  size: number;
  shift: number;
  style: RunStyle;
  inline?: TextInline;
  offset: number;
  space: boolean;
  softHyphen: boolean;
  breakAfter: boolean;
}
export interface TabAtom extends AtomBase {
  kind: 'tab';
  font: LayoutFont;
  size: number;
  style: RunStyle;
  /** Tab that belongs to a list label suffix. */
  labelTab?: boolean;
}
export interface BreakAtom extends AtomBase {
  kind: 'break';
  type: 'line' | 'page' | 'column';
}
export interface ImageAtom extends AtomBase {
  kind: 'image';
  inline: ImageInline;
  width: number;
  height: number;
}
export interface FloatAtom extends AtomBase {
  kind: 'float';
  inline: ImageInline;
}
export interface FieldAtom extends AtomBase {
  kind: 'field';
  field: 'PAGE' | 'NUMPAGES' | 'SECTIONPAGES';
  format?: string;
  width: number;
  font: LayoutFont;
  size: number;
  shift: number;
  style: RunStyle;
}
export interface MarkAtom extends AtomBase {
  kind: 'mark';
  mark: 'bookmark' | 'comment' | 'footnote' | 'endnote';
  name: string;
}

/** A typeset object placed like a character: an equation. Its items are relative to its baseline origin. */
export interface ObjectAtom extends AtomBase {
  kind: 'object';
  width: number;
  ascent: number;
  descent: number;
  items: DisplayItem[];
  /** Display equations sit alone on their line, aligned on their own. */
  align?: 'left' | 'center' | 'right';
}

export type Atom = CharAtom | TabAtom | BreakAtom | ImageAtom | FloatAtom | FieldAtom | MarkAtom | ObjectAtom;

// ---------------------------------------------------------------------------------------------
// Fragments (positioned pieces of a line)

export interface TextFragment {
  kind: 'text';
  x: number;
  width: number;
  chars: string[];
  /** x of each char relative to the line origin (column left). */
  xs: number[];
  widths: number[];
  font: LayoutFont;
  size: number;
  shift: number;
  style: RunStyle;
  ctx: RangeCtx;
  inline?: TextInline;
  offsets: number[];
  /** Width of trailing spaces that hang past the line end (not underlined). */
  trailing: number;
  /**
   * Right-to-left or complex-script text (Arabic, Hebrew, Indic…): its characters in logical
   * order, for renderers that shape text themselves. `chars` and `xs` are in visual order.
   */
  complex?: { text: string; rtl: boolean };
  /** Shaped glyphs per cluster (same order as `chars`). */
  glyphs?: PlacedGlyph[][];
}
export interface TabFragment {
  kind: 'tab';
  x: number;
  width: number;
  leader: TabStop['leader'];
  font: LayoutFont;
  size: number;
  style: RunStyle;
  ctx: RangeCtx;
}
export interface ImageFragment {
  kind: 'image';
  x: number;
  width: number;
  height: number;
  inline: ImageInline;
  ctx: RangeCtx;
}
export interface FloatFragment {
  kind: 'float';
  x: number;
  inline: ImageInline;
}
export interface FieldFragment {
  kind: 'field';
  x: number;
  width: number;
  field: FieldAtom['field'];
  format?: string;
  font: LayoutFont;
  size: number;
  shift: number;
  style: RunStyle;
  ctx: RangeCtx;
}
export interface MarkFragment {
  kind: 'mark';
  x: number;
  mark: MarkAtom['mark'];
  name: string;
}

export interface ObjectFragment {
  kind: 'object';
  x: number;
  width: number;
  items: DisplayItem[];
  ctx: RangeCtx;
}
export type Fragment = TextFragment | TabFragment | ImageFragment | FloatFragment | FieldFragment | MarkFragment | ObjectFragment;

export interface Line {
  fragments: Fragment[];
  /** Natural ascent/descent of the tallest content (pt). */
  ascent: number;
  descent: number;
  /** Where the line's content box starts and ends (relative to the column's left edge). */
  left: number;
  right: number;
  breakAfter?: 'page' | 'column';
  last: boolean;
  /** Range state at the end of the line (for zero-length ranges at line ends). */
  ctx: RangeCtx;
  /** Empty space above the line: it was moved below a floating object that left no room. */
  gapBefore?: number;
}

/** Horizontal room for a line, relative to the column's left edge (text wrapping around floats). */
export interface LineBounds {
  left: number;
  right: number;
  /** When there is too little room here: the y (relative to the paragraph top) where room starts. */
  next?: number;
}

// ---------------------------------------------------------------------------------------------
// Building atoms

export interface AtomContext {
  fonts: FontProvider;
  theme: Theme;
  tracker: RangeTracker;
  /** Label for a footnote/endnote reference. */
  noteLabel(inline: NoteRefInline): string;
}

const NO_LINE_START = new Set([...'、。，．：；？！）」』】〕〉》ー々〻ぁぃぅぇぉっゃゅょゎァィゥェォッャュョヮヵヶ・‐–—…‥〜!),.:;?]}'].map((c) => c.codePointAt(0)!));
const NO_LINE_END = new Set([...'（「『【〔〈《([{'].map((c) => c.codePointAt(0)!));

function clusters(text: string): string[] {
  const out: string[] = [];
  for (const ch of text) {
    const cp = ch.codePointAt(0)!;
    const combining = (cp >= 0x300 && cp <= 0x36f) || (cp >= 0xfe00 && cp <= 0xfe0f) || cp === 0x200d || (cp >= 0x1f3fb && cp <= 0x1f3ff);
    if (combining && out.length) out[out.length - 1] += ch;
    else out.push(ch);
  }
  return out;
}

export function styleFont(fonts: FontProvider, style: RunStyle, cp: number): { font: LayoutFont; size: number } {
  const slot = scriptSlot(cp, style);
  const cs = slot === 'cs';
  const font = fonts.get(fontNameForSlot(style, slot), cs ? style.boldCs : style.bold, cs ? style.italicCs : style.italic);
  return { font: fonts.forChar(font, cp), size: cs ? style.sizeCs : style.size };
}

/** Effective size and baseline shift for super/subscript and raised text. */
function sizeAndShift(style: RunStyle, size: number): { size: number; shift: number } {
  let shift = style.position;
  if (style.vertAlign === 'superscript') {
    shift += size * 0.33;
    size *= 0.65;
  } else if (style.vertAlign === 'subscript') {
    shift -= size * 0.08;
    size *= 0.65;
  }
  return { size, shift };
}

export function pushText(ctx: AtomContext, atoms: Atom[], text: string, props: RunProps, inline?: TextInline): void {
  const style = resolveRunStyle(props, ctx.theme);
  if (style.hidden) return;
  const first = atoms.length;
  pushClusters(ctx, atoms, text, style, inline);
  shapeAtoms(atoms, first);
}

function pushClusters(ctx: AtomContext, atoms: Atom[], text: string, style: RunStyle, inline: TextInline | undefined): void {
  let offset = 0;
  for (const cluster of clusters(text)) {
    const cp = cluster.codePointAt(0)!;
    let display = cluster;
    let smallCap = false;
    if (style.caps || style.smallCaps) {
      const upper = cluster.toUpperCase();
      if (upper !== cluster) {
        display = upper;
        smallCap = style.smallCaps && !style.caps;
      }
    }
    const { font, size: baseSize } = styleFont(ctx.fonts, style, display.codePointAt(0)!);
    const { size: shiftedSize, shift } = sizeAndShift(style, baseSize);
    const size = smallCap ? shiftedSize * 0.8 : shiftedSize;
    const softHyphen = cp === 0xad;
    let width = 0;
    if (!softHyphen) {
      for (const ch of display) width += font.advance(ch.codePointAt(0)!) * size;
      width = width * style.scale + style.spacing;
    }
    atoms.push({
      kind: 'char',
      ch: display,
      width,
      font,
      size,
      shift,
      style,
      inline,
      offset,
      space: cp === 0x20 || cp === 0x3000,
      softHyphen,
      breakAfter: false,
      ctx: ctx.tracker.ctx,
    });
    offset += cluster.length;
  }
}

/** Build the atom list of a paragraph, starting with its list label if any. */
export function buildAtoms(para: DocParagraph, ctx: AtomContext, label?: { text: string; props: RunProps; suffix: 'tab' | 'space' | 'nothing' }): Atom[] {
  const atoms: Atom[] = [];
  if (label && label.text) {
    pushText(ctx, atoms, label.text, label.props);
    if (label.suffix === 'tab') {
      const style = resolveRunStyle(label.props, ctx.theme);
      const { font, size } = styleFont(ctx.fonts, style, 0x20);
      atoms.push({ kind: 'tab', font, size, style, ctx: ctx.tracker.ctx, labelTab: true });
    } else if (label.suffix === 'space') {
      pushText(ctx, atoms, ' ', label.props);
    }
  }
  for (const inline of para.inlines) {
    if (ctx.tracker.apply(inline)) {
      if (inline.kind === 'bookmarkStart') atoms.push({ kind: 'mark', mark: 'bookmark', name: inline.name, ctx: ctx.tracker.ctx });
      else if (inline.kind === 'commentStart') atoms.push({ kind: 'mark', mark: 'comment', name: inline.id, ctx: ctx.tracker.ctx });
      continue;
    }
    switch (inline.kind) {
      case 'text':
        pushText(ctx, atoms, inline.text, inline.props, inline);
        break;
      case 'tab': {
        const style = resolveRunStyle(inline.props, ctx.theme);
        if (style.hidden) break;
        const { font, size } = styleFont(ctx.fonts, style, 0x20);
        atoms.push({ kind: 'tab', font, size, style, ctx: ctx.tracker.ctx });
        break;
      }
      case 'break':
        if (!resolveRunStyle(inline.props, ctx.theme).hidden) atoms.push({ kind: 'break', type: inline.type, ctx: ctx.tracker.ctx });
        break;
      case 'image':
        if (inline.float) atoms.push({ kind: 'float', inline, ctx: ctx.tracker.ctx });
        else atoms.push({ kind: 'image', inline, width: inline.width, height: inline.height, ctx: ctx.tracker.ctx });
        break;
      case 'math': {
        const style = resolveRunStyle(inline.props, ctx.theme);
        // An equation alone in its paragraph is a display equation, as in Word.
        const display = inline.display || para.inlines.every((i) => i.kind === 'math' || (i.kind !== 'text' && i.kind !== 'tab' && i.kind !== 'image' && i.kind !== 'field'));
        const m = layoutMath(inline.node, { fonts: ctx.fonts, size: style.size, color: style.color === 'auto' ? '000000' : style.color, display });
        const jc = mathJc(inline.node);
        atoms.push({ kind: 'object', width: m.width, ascent: m.ascent, descent: m.descent, items: m.items, align: display ? jc : undefined, ctx: ctx.tracker.ctx });
        break;
      }
      case 'field': {
        const style = resolveRunStyle(inline.props, ctx.theme);
        if (style.hidden) break;
        const placeholder = inline.cached.trim() || '0';
        const { font, size: base } = styleFont(ctx.fonts, style, 0x30);
        const { size, shift } = sizeAndShift(style, base);
        let width = 0;
        for (const ch of placeholder) width += font.advance(ch.codePointAt(0)!) * size;
        atoms.push({ kind: 'field', field: inline.field, format: inline.format, width, font, size, shift, style, ctx: ctx.tracker.ctx });
        break;
      }
      case 'noteRef': {
        if (inline.id) atoms.push({ kind: 'mark', mark: inline.noteType, name: inline.id, ctx: ctx.tracker.ctx });
        if (inline.customMark) break;
        pushText(ctx, atoms, ctx.noteLabel(inline), inline.props);
        break;
      }
      default:
        break;
    }
  }
  markBreakOpportunities(atoms);
  analyzeBidi(atoms, !!para.props.bidi);
  return atoms;
}

function markBreakOpportunities(atoms: Atom[]): void {
  for (let i = 0; i < atoms.length; i++) {
    const a = atoms[i]!;
    if (a.kind !== 'char') {
      // Images and fields can break on both sides.
      if (a.kind === 'image' || a.kind === 'field' || a.kind === 'object') {
        const prev = atoms[i - 1];
        if (prev && prev.kind === 'char') prev.breakAfter = true;
        if (a.kind === 'field') continue;
      }
      continue;
    }
    const cp = a.ch.codePointAt(0)!;
    const next = atoms[i + 1];
    const nextCp = next && next.kind === 'char' ? next.ch.codePointAt(0)! : undefined;
    if (a.space || a.softHyphen || cp === 0x200b || cp === 0x2014 || cp === 0x2013 || cp === 0x2010) {
      a.breakAfter = true;
      continue;
    }
    if (cp === 0x2d && nextCp !== undefined && /[\p{L}\p{N}]/u.test(String.fromCodePoint(nextCp))) {
      const prev = atoms[i - 1];
      if (prev && prev.kind === 'char' && /[\p{L}\p{N}]/u.test(prev.ch)) a.breakAfter = true;
      continue;
    }
    if (nextCp !== undefined && (isEastAsian(cp) || isEastAsian(nextCp)) && !NO_LINE_START.has(nextCp) && !NO_LINE_END.has(cp)) {
      a.breakAfter = true;
    }
    if (next && (next.kind === 'image' || next.kind === 'object')) a.breakAfter = true;
  }
}

// ---------------------------------------------------------------------------------------------
// Line breaking

export interface LineGeometry {
  /** Left edge of the first line / other lines, relative to the column's left edge. */
  firstLeft: number;
  left: number;
  /** Right edge, relative to the column's left edge. */
  right: number;
  /** Paragraph indent (left), used for the implicit hanging-indent tab stop. */
  indentLeft: number;
  tabs: TabStop[];
  defaultTab: number;
  jc: 'left' | 'center' | 'right' | 'both' | 'distribute';
  /** Metrics used for empty lines and the paragraph mark. */
  markAscent: number;
  markDescent: number;
  /** Room at a given height (relative to the paragraph's first line), for wrapping around floats. */
  bounds?: (top: number, height: number) => LineBounds | undefined;
  /** The height a finished line will take (to know where the next line starts). */
  lineHeight?: (line: Line) => number;
  /**
   * A right-to-left paragraph: the geometry above is measured from the right edge (start),
   * and lines are mirrored onto the page within this width.
   */
  mirror?: number;
}

/** Narrowest line worth placing beside a floating object; anything narrower moves below it. */
const MIN_WRAP_WIDTH = 24;

const EPS = 0.01;

function atomWidth(a: Atom): number {
  switch (a.kind) {
    case 'char':
      return a.width;
    case 'image':
    case 'object':
      return a.width;
    case 'field':
      return a.width;
    default:
      return 0;
  }
}

function findTabStop(g: LineGeometry, x: number, first: boolean, labelTab: boolean | undefined): TabStop {
  // Hanging indents create an implicit stop at the left indent for the first line.
  const implicit = first && g.indentLeft > x + EPS ? g.indentLeft : undefined;
  for (const t of g.tabs) {
    if (t.align === 'bar' || t.align === 'clear') continue;
    if (t.pos > x + EPS) {
      if (implicit !== undefined && implicit < t.pos && labelTab) return { pos: implicit, align: 'left', leader: 'none' };
      return t;
    }
  }
  if (implicit !== undefined) return { pos: implicit, align: 'left', leader: 'none' };
  const step = g.defaultTab > 0 ? g.defaultTab : 36;
  return { pos: (Math.floor(x / step + EPS) + 1) * step, align: 'left', leader: 'none' };
}

/** Width of the segment after a tab, up to the next tab or break (decimal: up to the separator). */
function segmentWidth(atoms: Atom[], from: number, decimal: boolean): number {
  let w = 0;
  for (let i = from; i < atoms.length; i++) {
    const a = atoms[i]!;
    if (a.kind === 'tab' || a.kind === 'break') break;
    if (decimal && a.kind === 'char' && (a.ch === '.' || a.ch === ',')) break;
    w += atomWidth(a);
  }
  return w;
}

/** Geometry for the line starting at `top`: the paragraph's own, narrowed around floats. */
function lineGeometry(g: LineGeometry, first: boolean, top: number): { g: LineGeometry; left: number; gap: number } {
  const own = first ? g.firstLeft : g.left;
  if (!g.bounds) return { g, left: own, gap: 0 };
  const estimate = g.markAscent + g.markDescent;
  let gap = 0;
  const W = g.mirror;
  for (let tries = 0; tries < 20; tries++) {
    const raw = g.bounds(top + gap, estimate);
    // Room beside floats is on the page; a right-to-left paragraph measures from the right.
    const b = raw && W !== undefined ? { ...raw, left: W - raw.right, right: W - raw.left } : raw;
    if (!b) return { g, left: own, gap };
    const left = Math.max(own, b.left);
    const right = Math.min(g.right, b.right);
    if (right - left >= MIN_WRAP_WIDTH || b.next === undefined || b.next <= top + gap) {
      return { g: { ...g, right, firstLeft: first ? left : g.firstLeft, left: first ? g.left : left }, left, gap };
    }
    gap = b.next - top;
  }
  return { g, left: own, gap };
}

export function breakLines(atoms: Atom[], geometry: LineGeometry): Line[] {
  const lines: Line[] = [];
  let start = 0;
  let first = true;
  let top = 0;
  if (atoms.length === 0) {
    const { g, left, gap } = lineGeometry(geometry, true, 0);
    return [{ fragments: [], ascent: g.markAscent, descent: g.markDescent, left, right: g.right, last: true, ctx: EMPTY_CTX, ...(gap ? { gapBefore: gap } : {}) }];
  }
  while (start < atoms.length) {
    const lg = lineGeometry(geometry, first, top);
    const g = lg.g;
    const left = lg.left;
    const pos = new Map<number, number>();
    const tabWidths = new Map<number, number>();
    const tabStops = new Map<number, TabStop>();
    let x = left;
    let end = atoms.length;
    let forced: BreakAtom['type'] | undefined;
    let lastBreak = -1;
    let visible = false;
    for (let i = start; i < atoms.length; i++) {
      const a = atoms[i]!;
      if (a.kind === 'break') {
        pos.set(i, x);
        end = i + 1;
        forced = a.type;
        break;
      }
      if (a.kind === 'tab') {
        const stop = findTabStop(g, x, first, a.labelTab);
        let w = stop.pos - x;
        if (stop.align === 'right') w -= segmentWidth(atoms, i + 1, false);
        else if (stop.align === 'center') w -= segmentWidth(atoms, i + 1, false) / 2;
        else if (stop.align === 'decimal') w -= segmentWidth(atoms, i + 1, true);
        w = Math.max(0, w);
        if (x + w > g.right + EPS && visible && stop.align === 'left' && stop.pos > g.right) {
          // A tab past the right edge wraps the following content.
          end = i;
          break;
        }
        pos.set(i, x);
        tabWidths.set(i, w);
        tabStops.set(i, stop);
        x += w;
        lastBreak = i - 1 >= start ? i - 1 : lastBreak;
        visible = true;
        continue;
      }
      const w = atomWidth(a);
      if (a.kind === 'char' && a.space) {
        pos.set(i, x);
        x += w;
        lastBreak = i;
        continue;
      }
      if (w > 0 && x + w > g.right + EPS && visible) {
        end = lastBreak >= start ? lastBreak + 1 : i;
        break;
      }
      pos.set(i, x);
      x += w;
      if (w > 0 || a.kind === 'image' || a.kind === 'object') visible = true;
      if (a.kind === 'char' && a.breakAfter) lastBreak = i;
    }
    if (end <= start) end = start + 1; // always make progress
    const last = end >= atoms.length;
    const line = finishLine(atoms, start, end, pos, tabWidths, tabStops, g, left, forced, last);
    if (lg.gap) line.gapBefore = lg.gap;
    lines.push(line);
    top += lg.gap + (geometry.lineHeight ? geometry.lineHeight(line) : line.ascent + line.descent);
    start = end;
    first = false;
  }
  return lines;
}

function finishLine(
  atoms: Atom[],
  start: number,
  end: number,
  pos: Map<number, number>,
  tabWidths: Map<number, number>,
  tabStops: Map<number, TabStop>,
  g: LineGeometry,
  left: number,
  forced: BreakAtom['type'] | undefined,
  last: boolean,
): Line {
  // A soft hyphen at the end of a line becomes visible.
  const lastAtom = atoms[end - 1];
  if (lastAtom && lastAtom.kind === 'char' && lastAtom.softHyphen && !last) {
    lastAtom.ch = '-';
    lastAtom.width = lastAtom.font.advance(0x2d) * lastAtom.size;
  }

  // Content end ignoring trailing spaces and breaks.
  let contentEnd = left;
  let lastTab = -1;
  for (let i = start; i < end; i++) {
    const a = atoms[i]!;
    if (a.kind === 'tab') lastTab = i;
    if (a.kind === 'break' || (a.kind === 'char' && a.space)) continue;
    const p = pos.get(i);
    if (p === undefined) continue;
    const w = a.kind === 'tab' ? tabWidths.get(i)! : atomWidth(a);
    contentEnd = Math.max(contentEnd, p + w);
  }

  // Alignment and justification shift positions. A display equation alone on its line uses its own alignment.
  const extra = g.right - contentEnd;
  const shift = new Map<number, number>();
  let jc = g.jc;
  const objects = [];
  for (let i = start; i < end; i++) {
    const a = atoms[i]!;
    if (a.kind === 'object' && a.align) objects.push(a);
    else if (a.kind !== 'mark' && a.kind !== 'break' && a.kind !== 'float' && !(a.kind === 'char' && a.space)) objects.push(undefined);
  }
  if (objects.length === 1 && objects[0]) jc = objects[0].align!;
  const justify = (jc === 'both' && !last && forced !== 'page' && forced !== 'column') || jc === 'distribute';
  if (extra > EPS && (jc === 'center' || jc === 'right')) {
    const offset = jc === 'center' ? extra / 2 : extra;
    for (let i = start; i < end; i++) shift.set(i, offset);
  } else if (extra > EPS && justify) {
    let lastContent = start - 1;
    for (let i = start; i < end; i++) {
      const a = atoms[i]!;
      if (!(a.kind === 'break' || (a.kind === 'char' && a.space))) lastContent = i;
    }
    const spaces: number[] = [];
    for (let i = Math.max(start, lastTab + 1); i < lastContent; i++) {
      const a = atoms[i]!;
      if (a.kind === 'char' && a.space) spaces.push(i);
    }
    if (spaces.length > 0) {
      const each = extra / spaces.length;
      let acc = 0;
      let si = 0;
      for (let i = start; i < end; i++) {
        shift.set(i, acc);
        if (si < spaces.length && spaces[si] === i) {
          acc += each;
          si++;
        }
      }
      for (const i of spaces) {
        const a = atoms[i] as CharAtom;
        a.width += each;
      }
    } else if (jc === 'distribute') {
      const chars: number[] = [];
      for (let i = Math.max(start, lastTab + 1); i < lastContent; i++) if (atoms[i]!.kind === 'char') chars.push(i);
      if (chars.length > 0) {
        const each = extra / chars.length;
        let acc = 0;
        for (let i = start; i < end; i++) {
          shift.set(i, acc);
          if (chars.includes(i)) {
            (atoms[i] as CharAtom).width += each;
            acc += each;
          }
        }
      }
    }
  }

  // Visual order: right-to-left runs are reversed (Unicode bidi), and a right-to-left
  // paragraph's line is mirrored onto the page.
  let order: number[] = [];
  for (let i = start; i < end; i++) order.push(i);
  let xOf = (i: number) => (pos.get(i) ?? left) + (shift.get(i) ?? 0);
  const levelOf = new Map<number, number>();
  const info = bidiOf(atoms);
  if (info || g.mirror !== undefined) {
    const widthOf = (i: number) => {
      const a = atoms[i]!;
      if (a.kind === 'tab') return tabWidths.get(i) ?? 0;
      if (a.kind === 'char' && a.softHyphen && a.ch !== '-') return 0;
      return atomWidth(a);
    };
    let lo = Infinity;
    let hi = -Infinity;
    for (const i of order) {
      if (!pos.has(i)) continue;
      lo = Math.min(lo, xOf(i));
      hi = Math.max(hi, xOf(i) + widthOf(i));
    }
    if (lo === Infinity) lo = hi = left;
    if (info) {
      const visual = visualAtoms(info, start, end);
      order = visual.map((v) => v.atom);
      for (const v of visual) levelOf.set(v.atom, v.level);
    }
    const placed = new Map<number, number>();
    let x = g.mirror !== undefined ? g.mirror - hi : lo;
    for (const i of order) {
      placed.set(i, x);
      x += widthOf(i);
    }
    xOf = (i) => placed.get(i) ?? left;
  }

  // Group atoms into fragments.
  const fragments: Fragment[] = [];
  let ascent = 0;
  let descent = 0;
  let current: TextFragment | undefined;
  let lineCtx = atoms[start]?.ctx ?? EMPTY_CTX;
  for (const i of order) {
    const a = atoms[i]!;
    const x = xOf(i);
    lineCtx = a.ctx;
    if (a.kind === 'char') {
      if (a.softHyphen && a.ch !== '-') continue;
      ascent = Math.max(ascent, a.font.ascent * a.size + a.shift);
      descent = Math.max(descent, a.font.descent * a.size - a.shift);
      const level = levelOf.get(i) ?? 0;
      const rtl = level % 2 === 1;
      const complex = rtl || needsShaping(a.ch);
      if (!current || current.font !== a.font || current.size !== a.size || current.shift !== a.shift || current.style !== a.style || current.ctx !== a.ctx || current.inline !== a.inline || !!current.complex !== complex || (current.complex && current.complex.rtl !== rtl)) {
        current = { kind: 'text', x, width: 0, chars: [], xs: [], widths: [], font: a.font, size: a.size, shift: a.shift, style: a.style, ctx: a.ctx, inline: a.inline, offsets: [], trailing: 0, ...(complex ? { complex: { text: '', rtl } } : {}) };
        fragments.push(current);
      }
      // Rule L4: brackets and similar characters are mirrored in right-to-left text.
      const shown = rtl && a.ch.length <= 2 ? String.fromCodePoint(mirrorChar(a.ch.codePointAt(0)!) ?? a.ch.codePointAt(0)!) : a.ch;
      if (a.glyphs || current.glyphs) {
        // Shaped glyphs per cluster; unshaped clusters in the same fragment get their plain glyphs.
        const plain = (atom: CharAtom, ch: string): PlacedGlyph[] =>
          [...ch].map((c, k) => ({ id: atom.font.glyphId(c.codePointAt(0)!), dx: 0, dy: 0, advance: atom.font.advance(c.codePointAt(0)!) * atom.size, text: k === 0 ? ch : '' }));
        current.glyphs ??= current.chars.map((c) => plain(a, c));
        current.glyphs.push(a.glyphs ?? plain(a, shown));
      }
      current.chars.push(shown);
      current.xs.push(x);
      current.widths.push(a.width);
      current.offsets.push(a.offset);
      current.width = Math.max(current.width, x + a.width - current.x);
      if (current.complex) current.complex.text = rtl ? a.ch + current.complex.text : current.complex.text + a.ch;
      continue;
    }
    current = undefined;
    switch (a.kind) {
      case 'tab': {
        const w = tabWidths.get(i) ?? 0;
        ascent = Math.max(ascent, a.font.ascent * a.size);
        descent = Math.max(descent, a.font.descent * a.size);
        fragments.push({ kind: 'tab', x, width: w, leader: tabStops.get(i)?.leader ?? 'none', font: a.font, size: a.size, style: a.style, ctx: a.ctx });
        break;
      }
      case 'image':
        ascent = Math.max(ascent, a.height);
        fragments.push({ kind: 'image', x, width: a.width, height: a.height, inline: a.inline, ctx: a.ctx });
        break;
      case 'float':
        fragments.push({ kind: 'float', x, inline: a.inline });
        break;
      case 'object':
        ascent = Math.max(ascent, a.ascent);
        descent = Math.max(descent, a.descent);
        fragments.push({ kind: 'object', x, width: a.width, items: a.items, ctx: a.ctx });
        break;
      case 'field':
        ascent = Math.max(ascent, a.font.ascent * a.size + a.shift);
        descent = Math.max(descent, a.font.descent * a.size - a.shift);
        fragments.push({ kind: 'field', x, width: a.width, field: a.field, format: a.format, font: a.font, size: a.size, shift: a.shift, style: a.style, ctx: a.ctx });
        break;
      case 'mark':
        fragments.push({ kind: 'mark', x, mark: a.mark, name: a.name });
        break;
      default:
        break;
    }
  }

  // Trailing spaces hang: record them so underlines stop at the last visible character.
  for (let f = fragments.length - 1; f >= 0; f--) {
    const frag = fragments[f]!;
    if (frag.kind !== 'text') {
      if (frag.kind === 'mark' || frag.kind === 'float') continue;
      break;
    }
    let trailing = 0;
    for (let c = frag.chars.length - 1; c >= 0 && (frag.chars[c] === ' ' || frag.chars[c] === '　'); c--) trailing += frag.widths[c]!;
    frag.trailing = trailing;
    if (trailing < frag.width) break;
  }

  const hasHeight = ascent > 0 || descent > 0;
  if (!hasHeight || last) {
    ascent = Math.max(ascent, g.markAscent);
    descent = Math.max(descent, g.markDescent);
  }
  return {
    fragments,
    ascent,
    descent,
    // A right-to-left line's content box, mirrored like its content.
    left: g.mirror !== undefined ? g.mirror - g.right : left,
    right: g.mirror !== undefined ? g.mirror - left : g.right,
    breakAfter: forced === 'page' || forced === 'column' ? forced : undefined,
    last,
    ctx: lineCtx,
  };
}

/** Alignment of a display equation (m:oMathParaPr/m:jc; Word's default centres it). */
function mathJc(node: XmlElement): 'left' | 'center' | 'right' {
  const jc = attr(firstChild(firstChild(node, 'm:oMathParaPr'), 'm:jc'), 'm:val');
  return jc === 'left' ? 'left' : jc === 'right' ? 'right' : 'center';
}

/** Scripts whose letters join or reorder (Arabic, Syriac, Thaana, N'Ko, Indic, Thai, Lao, Tibetan, Myanmar, Khmer, Mongolian). */
const SHAPED = /[\u0600-\u08ff\u0900-\u0dff\u0e00-\u0eff\u0f00-\u0fff\u1000-\u109f\u1780-\u18af\ua800-\uabff\ufb50-\ufdff\ufe70-\ufeff]/;

/** Whether a character needs a shaping engine to look right (its renderer should shape the whole run). */
export function needsShaping(ch: string): boolean {
  return SHAPED.test(ch);
}

/** Scripts written right to left (Hebrew, Arabic, Syriac, Thaana, N'Ko, Samaritan, Mandaic). */
const RTL_SCRIPT = /[\u0590-\u08ff\ufb1d-\ufdff\ufe70-\ufeff]/;

/**
 * Run the OpenType shaping engine over runs of atoms that need it: complex scripts (Arabic
 * joining, Indic reordering, mark placement), and Latin text where Word turns on kerning
 * (w:kern, from that size up) or ligatures (w14:ligatures). Each atom keeps its own cluster;
 * a ligature's glyph goes to the first atom it covers and the others get no glyphs.
 */
function shapeAtoms(atoms: Atom[], from: number): void {
  let i = from;
  while (i < atoms.length) {
    const a = atoms[i]!;
    if (a.kind !== 'char' || !wantsShaping(a)) {
      i++;
      continue;
    }
    // A run with one font, size and style (spaces included, so kerning and joining see whole words).
    let j = i + 1;
    while (j < atoms.length) {
      const b = atoms[j]!;
      if (b.kind !== 'char' || b.font !== a.font || b.size !== a.size || b.style !== a.style || b.softHyphen) break;
      j++;
    }
    shapeRun(atoms.slice(i, j) as CharAtom[]);
    i = j;
  }
}

function wantsShaping(a: CharAtom): boolean {
  if (!a.font.face || a.softHyphen) return false;
  if (needsShaping(a.ch)) return true;
  if (a.style.kern !== undefined && a.size >= a.style.kern) return true;
  return !!a.style.ligatures;
}

function shapeRun(run: CharAtom[]): void {
  const first = run[0]!;
  const font = first.font.face!.font;
  let text = '';
  const starts: number[] = [];
  for (const a of run) {
    starts.push(text.length);
    text += a.ch;
  }
  const rtl = RTL_SCRIPT.test(text);
  const features: Record<string, boolean> = {};
  const complex = run.some((a) => needsShaping(a.ch));
  if (!complex) {
    // Latin: only what Word asks for (it applies no ligatures by default).
    const lig = first.style.ligatures;
    features.liga = !!lig && lig !== 'none';
    features.clig = !!lig && /all|contextual/i.test(lig);
    features.dlig = !!lig && /all|discretional/i.test(lig);
    features.kern = first.style.kern !== undefined && first.size >= first.style.kern;
  }
  let glyphs;
  try {
    glyphs = shape(font, text, { direction: rtl ? 'rtl' : 'ltr', features });
  } catch {
    return;
  }
  const scale = first.size / (font.unitsPerEm || 1000);
  const owner = (cluster: number) => {
    let k = 0;
    while (k + 1 < starts.length && starts[k + 1]! <= cluster) k++;
    return k;
  };
  const per: PlacedGlyph[][] = run.map(() => []);
  // The text each glyph stands for: from its cluster to the next cluster's start.
  const clusterStarts = [...new Set(glyphs.map((g) => g.cluster))].sort((x, y) => x - y);
  const seen = new Set<number>();
  for (const g of glyphs) {
    const k = owner(g.cluster);
    const next = clusterStarts.find((c) => c > g.cluster) ?? text.length;
    const glyphText = seen.has(g.cluster) ? '' : text.slice(g.cluster, next);
    seen.add(g.cluster);
    per[k]!.push({ id: g.glyph, dx: g.offsetX * scale, dy: g.offsetY * scale, advance: g.advance * scale, text: glyphText });
  }
  run.forEach((a, k) => {
    const gl = per[k]!;
    a.glyphs = gl;
    const width = gl.reduce((s, g) => s + g.advance, 0);
    a.width = width * a.style.scale + (gl.length ? a.style.spacing : 0);
  });
}
