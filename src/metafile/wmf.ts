/**
 * Windows metafile player ([MS-WMF]), for the 16-bit pictures old Word documents embed
 * (Equation Editor 3.0 equations, clip art, early charts). Playback follows how applications
 * show a WMF: the viewport is the target frame and the picture's own window origin/extent map
 * onto it (the Aldus placeable header's bounding box is the window until a record sets one).
 * Viewport records and map modes are therefore not applied: they would size the picture in
 * device pixels instead of fitting it to the frame.
 */
import type { DisplayItem, FontProvider } from '../layout/types.js';
import { Canvas } from './canvas.js';
import type { ArcKind } from './canvas.js';
import { averageColor, cropRaster, decodeDib, dibIsBottomUp } from './dib.js';
import type { Matrix, Pt } from './geometry.js';
import { RecordReader } from './reader.js';
import { BLACKNESS, ETO_CLIPPED, ETO_OPAQUE, PATCOPY, SRCCOPY, WHITENESS, colorRef, makeBrush, makePen, selectInto } from './state.js';
import type { DCState, Font, GdiObject } from './state.js';
import { decodeAnsi } from './text.js';

const PLACEABLE_KEY = 0x9ac6cdd7;
const META_SETWINDOWEXT = 0x020c;

interface WmfHeader {
  placeable?: { left: number; top: number; right: number; bottom: number; inch: number };
  /** Offset of the first record. */
  records: number;
}

export function readWmfHeader(bytes: Uint8Array): WmfHeader | undefined {
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  let off = 0;
  let placeable: WmfHeader['placeable'];
  if (bytes.length >= 22 && view.getUint32(0, true) === PLACEABLE_KEY) {
    placeable = { left: view.getInt16(6, true), top: view.getInt16(8, true), right: view.getInt16(10, true), bottom: view.getInt16(12, true), inch: view.getUint16(14, true) };
    off = 22;
  }
  if (off + 18 > bytes.length) return undefined;
  const type = view.getUint16(off, true);
  const headerWords = view.getUint16(off + 2, true);
  const version = view.getUint16(off + 4, true);
  // Without the placeable key the standard header is all there is to recognise the format by.
  if (!placeable && (!(type === 1 || type === 2) || headerWords !== 9 || !(version === 0x100 || version === 0x300))) return undefined;
  return { placeable, records: off + Math.max(9, headerWords) * 2 };
}

/** Visit records (function number, parameter reader) until META_EOF or the data runs out. */
function eachRecord(bytes: Uint8Array, start: number, visit: (fn: number, r: RecordReader) => boolean | void): void {
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  for (let off = start; off + 6 <= bytes.length; ) {
    const words = view.getUint32(off, true);
    const fn = view.getUint16(off + 4, true);
    if (fn === 0 || words < 3) break;
    const end = Math.min(bytes.length, off + words * 2);
    if (visit(fn, new RecordReader(bytes, off + 6, end, view)) === false) break;
    off += words * 2;
  }
}

function firstWindowExt(bytes: Uint8Array, start: number): Pt | undefined {
  let ext: Pt | undefined;
  eachRecord(bytes, start, (fn, r) => {
    if (fn !== META_SETWINDOWEXT) return;
    ext = { x: r.i16(2), y: r.i16(0) };
    return false;
  });
  return ext;
}

/** Natural size in points: the placeable bounding box at its units per inch, else the window extent in twips. */
export function wmfSize(bytes: Uint8Array): { width: number; height: number } | undefined {
  const h = readWmfHeader(bytes);
  if (!h) return undefined;
  if (h.placeable) {
    const inch = h.placeable.inch || 1440;
    const width = (Math.abs(h.placeable.right - h.placeable.left) / inch) * 72;
    const height = (Math.abs(h.placeable.bottom - h.placeable.top) / inch) * 72;
    if (width > 0 && height > 0) return { width, height };
  }
  const ext = firstWindowExt(bytes, h.records);
  if (ext && ext.x && ext.y) return { width: (Math.abs(ext.x) / 1440) * 72, height: (Math.abs(ext.y) / 1440) * 72 };
  return undefined;
}

export function renderWmf(bytes: Uint8Array, frame: { x: number; y: number; w: number; h: number }, fonts: FontProvider): DisplayItem[] {
  const header = readWmfHeader(bytes);
  if (!header || !(frame.w > 0) || !(frame.h > 0)) return [];
  let window: { org: Pt; ext: Pt } | undefined;
  const p = header.placeable;
  if (p && p.right !== p.left && p.bottom !== p.top) window = { org: { x: p.left, y: p.top }, ext: { x: p.right - p.left, y: p.bottom - p.top } };
  if (!window && !firstWindowExt(bytes, header.records)) {
    // Nothing says where the picture is: draw it once in logical units to find out.
    const probe = play(bytes, header, { x: 0, y: 0, w: 1, h: 1 }, fonts, { org: { x: 0, y: 0 }, ext: { x: 1, y: 1 } }, 1);
    const box = itemBounds(probe);
    if (!box) return [];
    window = { org: { x: box.x, y: box.y }, ext: { x: box.w || 1, y: box.h || 1 } };
  }
  const natural = wmfSize(bytes);
  return play(bytes, header, frame, fonts, window, natural ? frame.h / natural.height : 1);
}

function play(
  bytes: Uint8Array,
  header: WmfHeader,
  frame: { x: number; y: number; w: number; h: number },
  fonts: FontProvider,
  window: { org: Pt; ext: Pt } | undefined,
  scale: number,
): DisplayItem[] {
  const pageToPoints = (s: DCState): Matrix => {
    const sx = frame.w / (s.windowExt.x || 1);
    const sy = frame.h / (s.windowExt.y || 1);
    return { a: sx, b: 0, c: 0, d: sy, e: frame.x - s.windowOrg.x * sx, f: frame.y - s.windowOrg.y * sy };
  };
  // One pixel of a 96 dpi screen and its 12pt default font, at the scale the picture is drawn.
  const canvas = new Canvas({ fonts, pageToPoints, pixel: 0.75 * scale, defaultFontSize: 12 * scale });
  if (window) {
    canvas.state.windowOrg = window.org;
    canvas.state.windowExt = window.ext;
  }
  const player = new WmfPlayer(canvas);
  eachRecord(bytes, header.records, (fn, r) => {
    try {
      player.play(fn, r);
    } catch {
      // A record we cannot make sense of is skipped; the rest of the picture still draws.
    }
  });
  return canvas.items;
}

class WmfPlayer {
  /** The WMF object table: a new object takes the lowest free slot, DeleteObject frees its slot. */
  private readonly objects: Array<GdiObject | undefined> = [];

  constructor(private readonly canvas: Canvas) {}

  private add(obj: GdiObject): void {
    let i = 0;
    while (this.objects[i] !== undefined) i++;
    this.objects[i] = obj;
  }

  play(fn: number, r: RecordReader): void {
    const c = this.canvas;
    const s = c.state;
    switch (fn) {
      case 0x0102: // META_SETBKMODE
        s.bkMode = r.u16(0);
        break;
      case 0x0106: // META_SETPOLYFILLMODE
        s.polyFillMode = r.u16(0);
        break;
      case 0x012e: // META_SETTEXTALIGN
        s.textAlign = r.u16(0);
        break;
      case 0x0108: // META_SETTEXTCHAREXTRA
        s.charExtra = r.i16(0);
        break;
      case 0x0201: // META_SETBKCOLOR
        s.bkColor = color(r, 0);
        break;
      case 0x0209: // META_SETTEXTCOLOR
        s.textColor = color(r, 0);
        break;
      case 0x020b: // META_SETWINDOWORG (y, x)
        s.windowOrg = { x: r.i16(2), y: r.i16(0) };
        break;
      case META_SETWINDOWEXT:
        s.windowExt = { x: r.i16(2), y: r.i16(0) };
        break;
      case 0x020f: // META_OFFSETWINDOWORG
        s.windowOrg = { x: s.windowOrg.x + r.i16(2), y: s.windowOrg.y + r.i16(0) };
        break;
      case 0x0410: {
        // META_SCALEWINDOWEXT: yDenom, yNum, xDenom, xNum
        const [yDen, yNum, xDen, xNum] = [r.i16(0), r.i16(2), r.i16(4), r.i16(6)];
        if (xDen && yDen) s.windowExt = { x: (s.windowExt.x * xNum) / xDen, y: (s.windowExt.y * yNum) / yDen };
        break;
      }
      case 0x001e: // META_SAVEDC
        c.save();
        break;
      case 0x0127: // META_RESTOREDC
        c.restore(r.i16(0));
        break;
      case 0x02fa: // META_CREATEPENINDIRECT: style, width (POINTS), colour
        this.add(makePen(r.u16(0), r.i16(2), color(r, 6)));
        break;
      case 0x02fc: // META_CREATEBRUSHINDIRECT: style, colour, hatch
        this.add(makeBrush(r.u16(0), color(r, 2)));
        break;
      case 0x0142: {
        // META_DIBCREATEPATTERNBRUSH: drawn as a flat fill of the pattern's average colour.
        const img = decodeDib(r.slice(4, r.size - 4), undefined, { paletteIndices: r.u16(2) === 1 });
        this.add({ kind: 'brush', style: 'pattern', color: img ? averageColor(img) : '808080' });
        break;
      }
      case 0x01f9: // META_CREATEPATTERNBRUSH (device-dependent bitmap)
        this.add(makeBrush(3, '808080'));
        break;
      case 0x02fb: // META_CREATEFONTINDIRECT
        this.add(logFont(r));
        break;
      case 0x00f7: // META_CREATEPALETTE
      case 0x06ff: // META_CREATEREGION
        this.add({ kind: 'other' });
        break;
      case 0x012d: // META_SELECTOBJECT
        selectInto(s, this.objects[r.u16(0)]);
        break;
      case 0x01f0: // META_DELETEOBJECT
        this.objects[r.u16(0)] = undefined;
        break;
      case 0x0214: // META_MOVETO (y, x)
        c.moveTo(r.i16(2), r.i16(0));
        break;
      case 0x0213: // META_LINETO (y, x)
        c.lineTo(r.i16(2), r.i16(0));
        break;
      case 0x0325: // META_POLYLINE
      case 0x0324: // META_POLYGON
        c.poly(readPoints(r, 2, r.fit(2, r.i16(0), 4)), fn === 0x0324);
        break;
      case 0x0538: {
        // META_POLYPOLYGON: count, points per polygon, then all points.
        const polys = r.fit(2, r.u16(0), 2);
        let at = 2 + polys * 2;
        const out: Pt[][] = [];
        for (let i = 0; i < polys; i++) {
          const n = r.fit(at, r.u16(2 + i * 2), 4);
          out.push(readPoints(r, at, n));
          at += n * 4;
        }
        c.polyPoly(out, true);
        break;
      }
      case 0x041b: // META_RECTANGLE (bottom, right, top, left)
        c.rectangle(r.i16(6), r.i16(4), r.i16(2), r.i16(0));
        break;
      case 0x0418: // META_ELLIPSE
        c.ellipse(r.i16(6), r.i16(4), r.i16(2), r.i16(0));
        break;
      case 0x061c: // META_ROUNDRECT (height, width, bottom, right, top, left)
        c.roundRect(r.i16(10), r.i16(8), r.i16(6), r.i16(4), r.i16(2), r.i16(0));
        break;
      case 0x0817: // META_ARC
      case 0x081a: // META_PIE
      case 0x0830: {
        // META_CHORD: yEnd, xEnd, yStart, xStart, bottom, right, top, left
        const kind: ArcKind = fn === 0x0817 ? 'arc' : fn === 0x081a ? 'pie' : 'chord';
        c.arc(kind, r.i16(14), r.i16(12), r.i16(10), r.i16(8), r.i16(6), r.i16(4), r.i16(2), r.i16(0));
        break;
      }
      case 0x0521: {
        // META_TEXTOUT: length, string (padded to a word), y, x
        const len = Math.max(0, r.i16(0));
        const padded = len + (len & 1);
        this.text(r.i16(4 + padded), r.i16(2 + padded), r.slice(2, len), undefined);
        break;
      }
      case 0x0a32:
        this.extTextOut(r);
        break;
      case 0x0f43: {
        // META_STRETCHDIB: rop, usage, srcH, srcW, ySrc, xSrc, destH, destW, yDst, xDst, DIB
        const info = r.slice(22, r.size - 22);
        this.dib(info, r.u32(0), r.u16(4) === 1, [r.i16(12), r.i16(10), r.i16(8), r.i16(6)], [r.i16(20), r.i16(18), r.i16(16), r.i16(14)], true);
        break;
      }
      case 0x0b41: {
        // META_DIBSTRETCHBLT; the bitmap-less form is a pattern blit with a reserved word.
        if (r.size <= 22) {
          this.patBlt(r.u32(0), r.i16(20), r.i16(18), r.i16(16), r.i16(14));
          break;
        }
        const info = r.slice(20, r.size - 20);
        this.dib(info, r.u32(0), false, [r.i16(10), r.i16(8), r.i16(6), r.i16(4)], [r.i16(18), r.i16(16), r.i16(14), r.i16(12)], false);
        break;
      }
      case 0x0940: {
        // META_DIBBITBLT: rop, ySrc, xSrc, height, width, yDest, xDest, DIB
        if (r.size <= 18) {
          this.patBlt(r.u32(0), r.i16(16), r.i16(14), r.i16(12), r.i16(10));
          break;
        }
        const [w, h] = [r.i16(10), r.i16(8)];
        this.dib(r.slice(16, r.size - 16), r.u32(0), false, [r.i16(6), r.i16(4), w, h], [r.i16(14), r.i16(12), w, h], false);
        break;
      }
      case 0x061d: // META_PATBLT: rop, height, width, y, x
        this.patBlt(r.u32(0), r.i16(10), r.i16(8), r.i16(6), r.i16(4));
        break;
      // Viewport and map mode records (see the file comment), clipping, palettes, regions,
      // escapes and ROP2 do not change what we draw.
    }
  }

  private text(x: number, y: number, bytes: Uint8Array, dx: number[] | undefined, opaque?: [number, number, number, number]): void {
    const font = this.canvas.state.font;
    const decoded = decodeAnsi(bytes, font.charset, font.faceName, dx);
    this.canvas.text(x, y, decoded.chars, decoded.dx, opaque);
  }

  private extTextOut(r: RecordReader): void {
    // META_EXTTEXTOUT: y, x, length, options, [rectangle], string (padded), [dx per byte]
    const y = r.i16(0);
    const x = r.i16(2);
    const len = Math.max(0, r.i16(4));
    const options = r.u16(6);
    let at = 8;
    let opaque: [number, number, number, number] | undefined;
    if (options & (ETO_OPAQUE | ETO_CLIPPED)) {
      if (options & ETO_OPAQUE) opaque = [r.i16(8), r.i16(10), r.i16(12), r.i16(14)];
      at = 16;
    }
    const str = r.slice(at, len);
    at += len + (len & 1);
    let dx: number[] | undefined;
    if (len && r.fit(at, len, 2) === len) {
      dx = [];
      for (let i = 0; i < len; i++) dx.push(r.i16(at + i * 2));
    }
    this.text(x, y, str, dx, opaque);
  }

  /** Draw a packed DIB: src and dest are [x, y, w, h]; `fromBottom` for StretchDIB's bottom-up source origin. */
  private dib(info: Uint8Array, rop: number, paletteIndices: boolean, src: number[], dest: number[], fromBottom: boolean): void {
    if (rop !== SRCCOPY) return;
    const img = decodeDib(info, undefined, { paletteIndices });
    if (!img) return;
    const [sx, sy, sw, sh] = src as [number, number, number, number];
    const top = fromBottom && dibIsBottomUp(info) ? img.height - sy - sh : sy;
    const [dx, dy, dw, dh] = dest as [number, number, number, number];
    this.canvas.image(cropRaster(img, sx, top, sw, sh), dx, dy, dw, dh);
  }

  private patBlt(op: number, x: number, y: number, w: number, h: number): void {
    const fill = op === PATCOPY ? undefined : op === BLACKNESS ? '000000' : op === WHITENESS ? 'FFFFFF' : null;
    if (fill !== null) this.canvas.fillRect(x, y, x + w, y + h, fill);
  }
}

function color(r: RecordReader, at: number): string {
  return colorRef(r.u8(at), r.u8(at + 1), r.u8(at + 2));
}

function readPoints(r: RecordReader, at: number, n: number): Pt[] {
  const out: Pt[] = [];
  for (let i = 0; i < n; i++) out.push({ x: r.i16(at + i * 4), y: r.i16(at + i * 4 + 2) });
  return out;
}

/** The 16-bit LOGFONT of META_CREATEFONTINDIRECT (face name in ANSI). */
function logFont(r: RecordReader): Font {
  let faceName = '';
  for (let i = 0; i < 32; i++) {
    const b = r.u8(18 + i);
    if (!b) break;
    faceName += String.fromCharCode(b);
  }
  return {
    kind: 'font',
    height: r.i16(0),
    weight: r.i16(8),
    italic: r.u8(10) !== 0,
    underline: r.u8(11) !== 0,
    strikeOut: r.u8(12) !== 0,
    charset: r.u8(13),
    faceName,
  };
}

/** Bounding box of display items in page points (for pictures that do not say how big they are). */
function itemBounds(items: DisplayItem[]): { x: number; y: number; w: number; h: number } | undefined {
  let x0 = Infinity;
  let y0 = Infinity;
  let x1 = -Infinity;
  let y1 = -Infinity;
  const add = (x: number, y: number): void => {
    if (!Number.isFinite(x) || !Number.isFinite(y)) return;
    x0 = Math.min(x0, x);
    y0 = Math.min(y0, y);
    x1 = Math.max(x1, x);
    y1 = Math.max(y1, y);
  };
  for (const it of items) {
    switch (it.type) {
      case 'path':
        for (const cmd of it.commands) {
          if (cmd.op === 'Z') continue;
          if (cmd.op === 'C') {
            add(cmd.x1, cmd.y1);
            add(cmd.x2, cmd.y2);
          }
          add(cmd.x, cmd.y);
        }
        break;
      case 'text': {
        const width = it.widths.reduce((a, b) => a + b, 0);
        add(it.x, it.y - it.font.ascent * it.size);
        add(it.x + width, it.y + it.font.descent * it.size);
        break;
      }
      case 'line':
        add(it.x1, it.y1);
        add(it.x2, it.y2);
        break;
      default:
        add(it.x, it.y);
        add(it.x + it.w, it.y + it.h);
    }
  }
  return x0 <= x1 && y0 <= y1 ? { x: x0, y: y0, w: x1 - x0, h: y1 - y0 } : undefined;
}
