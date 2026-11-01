/** Turns positioned lines into display items and collects link / comment / permission areas. */
import type { Border } from '../model/props.js';
import type { FloatInfo, ImageInline } from '../model/content.js';
import { formatNumber } from '../model/numbering.js';
import type { Fragment, Line, RangeCtx, TextFragment } from './lines.js';
import type { DisplayItem, LayoutFont, Page, Rect, TextItem } from './types.js';

export interface Painter {
  page: Page;
  /** Items drawn behind the text (floating images with behindDoc). */
  behind: DisplayItem[];
  /** Region used for column/margin-relative float positioning. */
  column: { x: number; width: number };
  /** Text of a page-number field at this position. */
  fieldValue(field: string, format?: string): string;
  /** Which story is being painted; only body lines are recorded for editing. */
  story: 'body' | 'other';
}

export function colorOf(c: string | undefined, fallback = '000000'): string {
  if (!c || c === 'auto') return fallback;
  return c;
}

export function addItem(p: Painter, item: DisplayItem): void {
  p.page.items.push(item);
}

function addArea(list: Rect[] & Array<Rect & Record<string, unknown>>, rect: Rect & Record<string, unknown>, sameKey: (a: Record<string, unknown>) => boolean): void {
  const prev = list[list.length - 1];
  // Merge with the previous rect when it continues it on the same line.
  if (prev && sameKey(prev) && Math.abs(prev.y - rect.y) < 0.5 && Math.abs(prev.h - rect.h) < 0.5 && rect.x - (prev.x + prev.w) < 1 && rect.x >= prev.x) {
    prev.w = rect.x + rect.w - prev.x;
    return;
  }
  list.push(rect);
}

function recordRanges(p: Painter, ctx: RangeCtx, rect: Rect): void {
  if (rect.w <= 0) return;
  if (ctx.link) {
    const target = ctx.link;
    addArea(p.page.links as never, { ...rect, target }, (a) => a.target === target);
  }
  for (const id of ctx.comments) addArea(p.page.comments as never, { ...rect, id }, (a) => a.id === id);
  for (const perm of ctx.perms) addArea(p.page.perms as never, { ...rect, id: perm.id, editor: perm.editor, group: perm.group }, (a) => a.id === perm.id);
  for (const lock of ctx.locks) addArea(p.page.locks as never, { ...rect, id: lock.id, title: lock.title }, (a) => a.id === lock.id);
}

export function textItem(x: number, y: number, chars: string[], xs: number[], widths: number[], font: LayoutFont, size: number, color: string, scale = 1): TextItem {
  return {
    type: 'text',
    x,
    y,
    text: chars.join(''),
    chars,
    xs,
    widths,
    glyphs: chars.map((c) => font.glyphId(c.codePointAt(0)!)),
    font,
    size,
    color,
    scale,
  };
}

/** Measure a plain string in a font (used for resolved fields). */
export function measureChars(text: string, font: LayoutFont, size: number, x: number): { chars: string[]; xs: number[]; widths: number[]; width: number } {
  const chars = [...text];
  const xs: number[] = [];
  const widths: number[] = [];
  let cx = x;
  for (const ch of chars) {
    const w = font.advance(ch.codePointAt(0)!) * size;
    xs.push(cx);
    widths.push(w);
    cx += w;
  }
  return { chars, xs, widths, width: cx - x };
}

function decorate(p: Painter, frag: { style: TextFragment['style']; font: LayoutFont; size: number; shift: number }, x: number, width: number, baseline: number): void {
  const { style, font, size } = frag;
  if (width <= 0) return;
  const color = colorOf(style.color);
  if (style.underline) {
    const thickness = Math.max(font.underlineThickness * size, 0.5) * (style.underline === 'thick' ? 2 : 1);
    const y = baseline - frag.shift + font.underlinePosition * size;
    const lineStyle = style.underline.startsWith('dot') ? 'dotted' : style.underline.startsWith('dash') ? 'dashed' : style.underline === 'double' ? 'double' : 'solid';
    addItem(p, { type: 'line', x1: x, y1: y, x2: x + width, y2: y, width: thickness, color: colorOf(style.underlineColor, color), style: lineStyle });
  }
  if (style.strike || style.dstrike) {
    const y = baseline - frag.shift - font.ascent * size * 0.3;
    const thickness = Math.max(size / 18, 0.5);
    addItem(p, { type: 'line', x1: x, y1: y, x2: x + width, y2: y, width: thickness, color, style: style.dstrike ? 'double' : 'solid' });
  }
}

export interface LineBand {
  top: number;
  height: number;
  baseline: number;
}

/** Paint a line whose top-left is (originX, top) with its baseline at `baseline`. */
export function paintLine(p: Painter, line: Line, originX: number, band: LineBand): void {
  const { top, height, baseline } = band;
  for (const frag of line.fragments) paintFragment(p, frag, originX, top, height, baseline);
}

function paintFragment(p: Painter, frag: Fragment, originX: number, top: number, height: number, baseline: number): void {
  switch (frag.kind) {
    case 'text': {
      const x = originX + frag.x;
      const bg = frag.style.highlight ?? frag.style.shading;
      const drawnWidth = frag.width - frag.trailing;
      if (bg) addItem(p, { type: 'rect', x, y: top, w: frag.width, h: height, fill: bg });
      const item = textItem(
        x,
        baseline - frag.shift,
        frag.chars,
        frag.xs.map((v) => originX + v),
        frag.widths,
        frag.font,
        frag.size,
        colorOf(frag.style.color),
        frag.style.scale,
      );
      if (frag.inline) item.source = { inline: frag.inline, offsets: frag.offsets };
      addItem(p, item);
      decorate(p, frag, x, drawnWidth, baseline);
      recordRanges(p, frag.ctx, { x, y: top, w: frag.width, h: height });
      break;
    }
    case 'tab': {
      const x = originX + frag.x;
      if (frag.leader !== 'none' && frag.width > 0) {
        const ch = frag.leader === 'dot' ? '.' : frag.leader === 'hyphen' ? '-' : frag.leader === 'middleDot' ? '·' : '_';
        const cw = frag.font.advance(ch.codePointAt(0)!) * frag.size;
        if (cw > 0) {
          const n = Math.floor((frag.width - cw * 0.5) / cw);
          const chars = Array<string>(Math.max(0, n)).fill(ch);
          const start = x + frag.width - n * cw;
          addItem(p, textItem(start, baseline, chars, chars.map((_, i) => start + i * cw), chars.map(() => cw), frag.font, frag.size, colorOf(frag.style.color)));
        }
      }
      if (frag.style.underline) decorate(p, { ...frag, shift: 0 }, x, frag.width, baseline);
      recordRanges(p, frag.ctx, { x, y: top, w: frag.width, h: height });
      break;
    }
    case 'image': {
      const x = originX + frag.x;
      const y = baseline - frag.height;
      addItem(p, { type: 'image', x, y, w: frag.width, h: frag.height, target: frag.inline.target, external: frag.inline.external, alt: frag.inline.alt, inline: frag.inline });
      recordRanges(p, frag.ctx, { x, y, w: frag.width, h: frag.height });
      break;
    }
    case 'float':
      paintFloat(p, frag.inline.float!, frag.inline, originX + frag.x, top);
      break;
    case 'field': {
      const x = originX + frag.x;
      const value = p.fieldValue(frag.field, frag.format);
      const m = measureChars(value, frag.font, frag.size, x);
      const item = textItem(x, baseline - frag.shift, m.chars, m.xs, m.widths, frag.font, frag.size, colorOf(frag.style.color));
      item.field = frag.field;
      addItem(p, item);
      decorate(p, frag, x, m.width, baseline);
      recordRanges(p, frag.ctx, { x, y: top, w: m.width, h: height });
      break;
    }
    case 'mark': {
      const x = originX + frag.x;
      if (frag.mark === 'bookmark') {
        if (!p.page.anchors.has(frag.name)) p.page.anchors.set(frag.name, { x, y: top });
      } else if (frag.mark === 'comment') {
        // A zero-width marker keeps empty comment ranges anchored.
        p.page.comments.push({ id: frag.name, x, y: top, w: 0, h: height });
      }
      break;
    }
  }
}

function paintFloat(p: Painter, f: FloatInfo, image: ImageInline, charX: number, lineTop: number): void {
  const page = p.page;
  const m = page.section.margin;
  const area = (rel: string, horizontal: boolean): { start: number; size: number } => {
    if (horizontal) {
      switch (rel) {
        case 'page':
          return { start: 0, size: page.width };
        case 'margin':
          return { start: m.left, size: page.width - m.left - m.right };
        case 'leftMargin':
        case 'insideMargin':
          return { start: 0, size: m.left };
        case 'rightMargin':
        case 'outsideMargin':
          return { start: page.width - m.right, size: m.right };
        case 'character':
          return { start: charX, size: 0 };
        default:
          return { start: p.column.x, size: p.column.width };
      }
    }
    switch (rel) {
      case 'page':
        return { start: 0, size: page.height };
      case 'margin':
        return { start: m.top, size: page.height - m.top - m.bottom };
      case 'topMargin':
        return { start: 0, size: m.top };
      case 'bottomMargin':
        return { start: page.height - m.bottom, size: m.bottom };
      default:
        return { start: lineTop, size: 0 };
    }
  };
  const h = area(f.hRelative, true);
  const v = area(f.vRelative, false);
  const align = (a: { start: number; size: number }, al: string | undefined, offset: number, extent: number): number => {
    if (al === 'center') return a.start + (a.size - extent) / 2;
    if (al === 'right' || al === 'bottom' || al === 'outside') return a.start + a.size - extent;
    if (al === 'left' || al === 'top' || al === 'inside') return a.start;
    return a.start + offset;
  };
  const x = align(h, f.hAlign, f.hOffset, image.width);
  const y = align(v, f.vAlign, f.vOffset, image.height);
  const item: DisplayItem = { type: 'image', x, y, w: image.width, h: image.height, target: image.target, external: image.external, alt: image.alt, inline: image };
  if (f.behindDoc) p.behind.push(item);
  else addItem(p, item);
}

/** Draw a border segment centred on the given edge. */
export function borderLine(p: Painter, b: Border | undefined, x1: number, y1: number, x2: number, y2: number): void {
  if (!b || b.style === 'nil' || b.style === 'none' || b.width <= 0) return;
  const style = b.style.startsWith('dot') ? 'dotted' : b.style.startsWith('dash') ? 'dashed' : b.style === 'double' ? 'double' : 'solid';
  addItem(p, { type: 'line', x1, y1, x2, y2, width: b.width, color: colorOf(b.color), style });
}

export function pageNumberLabel(n: number, format: string | undefined): string {
  return formatNumber(n, format ?? 'decimal');
}
