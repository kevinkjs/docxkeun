/**
 * Drawing objects into the display list: pictures (rotated and flipped), metafile pictures as
 * vectors, and — through `DrawingContext.content` — shapes, text boxes, groups, charts and SmartArt.
 */
import type { ImageInline } from '../model/content.js';
import type { DocxDocument } from '../model/document.js';
import { metafileKind, renderMetafile } from '../metafile/index.js';
import type { Painter } from './paint.js';
import type { DisplayItem, FontProvider, ImageItem, PathCommand, Rect } from './types.js';

/** What painting drawings needs beyond the page: the document and fonts. */
export interface DrawingContext {
  doc: DocxDocument;
  fonts: FontProvider;
  /** Draws an object's `content` (shapes, groups, charts, SmartArt) into `frame`; returns false when it cannot. */
  content?(p: Painter, image: ImageInline, frame: Rect): boolean;
}

/** A transform from an object's own box to the page: rotation about the centre and flips. */
export function objectTransform(frame: Rect, rotation = 0, flipH = false, flipV = false): ((x: number, y: number) => [number, number]) | undefined {
  if (!rotation && !flipH && !flipV) return undefined;
  const cx = frame.x + frame.w / 2;
  const cy = frame.y + frame.h / 2;
  const a = (rotation * Math.PI) / 180;
  const cos = Math.cos(a);
  const sin = Math.sin(a);
  return (x, y) => {
    let dx = x - cx;
    let dy = y - cy;
    if (flipH) dx = -dx;
    if (flipV) dy = -dy;
    return [cx + dx * cos - dy * sin, cy + dx * sin + dy * cos];
  };
}

/** Map path coordinates through a transform. */
export function transformCommands(commands: PathCommand[], t: (x: number, y: number) => [number, number]): PathCommand[] {
  return commands.map((c) => {
    if (c.op === 'Z') return c;
    if (c.op === 'C') {
      const [x1, y1] = t(c.x1, c.y1);
      const [x2, y2] = t(c.x2, c.y2);
      const [x, y] = t(c.x, c.y);
      return { op: 'C', x1, y1, x2, y2, x, y };
    }
    const [x, y] = t(c.x, c.y);
    return { op: c.op, x, y };
  });
}

/** Move display items by (dx, dy). */
export function translateItems(items: DisplayItem[], dx: number, dy: number): DisplayItem[] {
  return items.map((item) => {
    switch (item.type) {
      case 'text':
        return { ...item, x: item.x + dx, y: item.y + dy, xs: item.xs.map((v) => v + dx) };
      case 'path':
        return { ...item, commands: transformCommands(item.commands, (x, y) => [x + dx, y + dy]) };
      case 'line':
        return { ...item, x1: item.x1 + dx, y1: item.y1 + dy, x2: item.x2 + dx, y2: item.y2 + dy };
      default:
        return { ...item, x: item.x + dx, y: item.y + dy };
    }
  });
}

/** Rotate / flip display items drawn for an object (paths and images; text keeps its baseline). */
export function transformItems(items: DisplayItem[], t: ((x: number, y: number) => [number, number]) | undefined): DisplayItem[] {
  if (!t) return items;
  return items.map((item) => {
    switch (item.type) {
      case 'path':
        return { ...item, commands: transformCommands(item.commands, t) };
      case 'rect': {
        const c = [t(item.x, item.y), t(item.x + item.w, item.y), t(item.x + item.w, item.y + item.h), t(item.x, item.y + item.h)];
        return { type: 'path', commands: [{ op: 'M', x: c[0]![0], y: c[0]![1] }, ...c.slice(1).map(([x, y]) => ({ op: 'L' as const, x, y })), { op: 'Z' }], fill: item.fill };
      }
      case 'line': {
        const [x1, y1] = t(item.x1, item.y1);
        const [x2, y2] = t(item.x2, item.y2);
        return { ...item, x1, y1, x2, y2 };
      }
      case 'text': {
        // Text moves with its object but stays upright.
        const [x, y] = t(item.x, item.y);
        const dx = x - item.x;
        const dy = y - item.y;
        return { ...item, x, y: item.y + dy, xs: item.xs.map((v) => v + dx) };
      }
      default:
        return item;
    }
  });
}

/**
 * Draw an inline or floating object. Returns the items (the first is the object's frame, used
 * for hit-testing and the editing caret) for the caller to place in front of or behind the text.
 */
export function drawObject(p: Painter, image: ImageInline, frame: Rect): DisplayItem[] {
  const ctx = p.drawing;
  const marker: ImageItem = { type: 'image', ...frame, target: image.target, external: image.external, alt: image.alt, inline: image };
  if (image.rotation) marker.rotation = image.rotation;
  if (image.flipH) marker.flipH = true;
  if (image.flipV) marker.flipV = true;
  if (!ctx) return [marker];
  if (image.content) {
    const saved = p.page.items;
    const out: DisplayItem[] = [];
    p.page.items = out;
    let ok = false;
    try {
      ok = !!ctx.content?.(p, image, frame);
    } finally {
      p.page.items = saved;
    }
    if (ok) return [{ ...marker, frameOnly: true }, ...out];
    return [{ ...marker, frameOnly: true }, placeholder(frame)];
  }
  const bytes = ctx.doc.mediaSync(image.target);
  if (bytes && metafileKind(bytes)) {
    const items = renderMetafile(bytes, frame, ctx.fonts);
    return [{ ...marker, frameOnly: true }, ...transformItems(items, objectTransform(frame, image.rotation, image.flipH, image.flipV))];
  }
  return [marker];
}

/** A light box for objects that cannot be drawn. */
export function placeholder(frame: Rect): DisplayItem {
  return { type: 'path', commands: [{ op: 'M', x: frame.x, y: frame.y }, { op: 'L', x: frame.x + frame.w, y: frame.y }, { op: 'L', x: frame.x + frame.w, y: frame.y + frame.h }, { op: 'L', x: frame.x, y: frame.y + frame.h }, { op: 'Z' }], fill: 'F1F3F4', stroke: 'C0C4C8', strokeWidth: 0.5 };
}
