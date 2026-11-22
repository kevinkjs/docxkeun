/**
 * Enhanced metafile player ([MS-EMF]). Reads the header to find the picture frame on the
 * reference device, then plays the records through a Canvas: logical units go through the world
 * transform, the map mode's window/viewport mapping to reference-device pixels, and finally the
 * scaling that puts the header's frame rectangle onto the requested frame. EMF+ records (inside
 * EMR_COMMENT) are skipped: a dual EMF+/EMF file still draws from its EMF records, an EMF+-only
 * file draws nothing.
 */
import type { DisplayItem, FontProvider, RasterImage } from '../layout/types.js';
import { Canvas } from './canvas.js';
import type { ArcKind } from './canvas.js';
import { averageColor, cropRaster, decodeDib, dibIsBottomUp, keyOut } from './dib.js';
import { IDENTITY, compose } from './geometry.js';
import type { Matrix, Pt } from './geometry.js';
import { RecordReader } from './reader.js';
import {
  BLACKNESS,
  BS_DIBPATTERNPT,
  BS_NULL,
  BS_PATTERN,
  ETO_GLYPH_INDEX,
  ETO_NO_RECT,
  ETO_OPAQUE,
  ETO_PDY,
  ETO_SMALL_CHARS,
  MM_ANISOTROPIC,
  MM_HIENGLISH,
  MM_HIMETRIC,
  MM_ISOTROPIC,
  MM_LOENGLISH,
  MM_LOMETRIC,
  MM_TWIPS,
  PATCOPY,
  PS_GEOMETRIC,
  PS_NULL,
  SRCCOPY,
  WHITENESS,
  colorRef,
  makeBrush,
  makePen,
  selectInto,
  stockObject,
} from './state.js';
import type { DCState, Font, GdiObject } from './state.js';
import { decodeAnsi, decodeWide } from './text.js';

const EMR_HEADER = 1;
const EMR_EOF = 14;
const EMF_SIGNATURE = 0x464d4520;

export interface EmfHeader {
  /** Bounds of the drawing in reference-device pixels (inclusive). */
  bounds: [number, number, number, number];
  /** The picture frame in 0.01 mm (inclusive). */
  frame: [number, number, number, number];
  /** Reference-device pixels per millimetre. */
  pxPerMm: Pt;
}

export function isEmf(bytes: Uint8Array): boolean {
  if (bytes.length < 88) return false;
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  return view.getUint32(0, true) === EMR_HEADER && view.getUint32(40, true) === EMF_SIGNATURE;
}

export function readEmfHeader(bytes: Uint8Array): EmfHeader | undefined {
  if (!isEmf(bytes)) return undefined;
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  const r = new RecordReader(bytes, 0, Math.min(bytes.length, view.getUint32(4, true)), view);
  const rect = (o: number): [number, number, number, number] => [r.i32(o), r.i32(o + 4), r.i32(o + 8), r.i32(o + 12)];
  let mmX = r.i32(80);
  let mmY = r.i32(84);
  // Header extension 2 gives the device size in micrometres, more precise than whole millimetres.
  // It is present when the record is long enough and the description/pixel format do not start inside it.
  const offDescription = r.u32(64);
  const offPixelFormat = r.u32(92);
  if (r.size >= 108 && (offDescription === 0 || offDescription >= 108) && (offPixelFormat === 0 || offPixelFormat >= 108)) {
    const umX = r.i32(100);
    const umY = r.i32(104);
    if (umX > 0 && umY > 0) {
      mmX = umX / 1000;
      mmY = umY / 1000;
    }
  }
  const devX = r.i32(72);
  const devY = r.i32(76);
  // A header without device metrics: assume a 96 dpi reference device.
  const pxPerMm = devX > 0 && devY > 0 && mmX > 0 && mmY > 0 ? { x: devX / mmX, y: devY / mmY } : { x: 96 / 25.4, y: 96 / 25.4 };
  return { bounds: rect(8), frame: rect(24), pxPerMm };
}

/** The picture frame in reference-device pixels: the header frame, or the drawing bounds when the frame is empty. */
function deviceFrame(h: EmfHeader): [number, number, number, number] | undefined {
  const [fl, ft, fr, fb] = h.frame;
  const f: [number, number, number, number] = [(fl / 100) * h.pxPerMm.x, (ft / 100) * h.pxPerMm.y, (fr / 100) * h.pxPerMm.x, (fb / 100) * h.pxPerMm.y];
  if (f[2] > f[0] && f[3] > f[1]) return f;
  const [bl, bt, br, bb] = h.bounds;
  if (br >= bl && bb >= bt) return [bl, bt, br + 1, bb + 1];
  return undefined;
}

/** Natural size in points, from the frame (0.01 mm), else from the bounds. */
export function emfSize(bytes: Uint8Array): { width: number; height: number } | undefined {
  const h = readEmfHeader(bytes);
  if (!h) return undefined;
  const [fl, ft, fr, fb] = h.frame;
  const toPt = 72 / 25.4;
  if (fr > fl && fb > ft) return { width: ((fr - fl) / 100) * toPt, height: ((fb - ft) / 100) * toPt };
  const [bl, bt, br, bb] = h.bounds;
  if (br >= bl && bb >= bt) return { width: ((br - bl + 1) / h.pxPerMm.x) * toPt, height: ((bb - bt + 1) / h.pxPerMm.y) * toPt };
  return undefined;
}

/** Page space → reference-device pixels for the state's map mode, window and viewport. */
function pageToDevice(s: DCState, pxPerMm: Pt): Matrix {
  let sx = 1;
  let sy = 1;
  const metric = (mm: number): void => {
    // The fixed map modes have y pointing up.
    sx = mm * pxPerMm.x;
    sy = -mm * pxPerMm.y;
  };
  switch (s.mapMode) {
    case MM_LOMETRIC:
      metric(0.1);
      break;
    case MM_HIMETRIC:
      metric(0.01);
      break;
    case MM_LOENGLISH:
      metric(0.254);
      break;
    case MM_HIENGLISH:
      metric(0.0254);
      break;
    case MM_TWIPS:
      metric(25.4 / 1440);
      break;
    case MM_ISOTROPIC:
    case MM_ANISOTROPIC: {
      sx = s.viewportExt.x / (s.windowExt.x || 1);
      sy = s.viewportExt.y / (s.windowExt.y || 1);
      if (s.mapMode === MM_ISOTROPIC) {
        // Isotropic keeps units square by shrinking the larger scale, keeping each axis' direction.
        const k = Math.min(Math.abs(sx), Math.abs(sy));
        sx = Math.sign(sx || 1) * k;
        sy = Math.sign(sy || 1) * k;
      }
      break;
    }
  }
  return { a: sx, b: 0, c: 0, d: sy, e: s.viewportOrg.x - s.windowOrg.x * sx, f: s.viewportOrg.y - s.windowOrg.y * sy };
}

/** Whether the picture was recorded in GM_ADVANCED (world transforms or advanced-mode text), which changes how arcs turn. */
function usesAdvancedMode(bytes: Uint8Array, view: DataView): boolean {
  for (let off = 0; off + 8 <= bytes.length; ) {
    const type = view.getUint32(off, true);
    const size = view.getUint32(off + 4, true);
    if (size < 8 || off + size > bytes.length || type === EMR_EOF) break;
    if (type === 35 || (type === 36 && size >= 36 && view.getUint32(off + 32, true) !== 1)) return true;
    if ((type === 83 || type === 84) && size >= 28 && view.getUint32(off + 24, true) === 2) return true;
    off += size;
  }
  return false;
}

export function renderEmf(bytes: Uint8Array, frame: { x: number; y: number; w: number; h: number }, fonts: FontProvider): DisplayItem[] {
  const header = readEmfHeader(bytes);
  if (!header) return [];
  const dev = deviceFrame(header);
  if (!dev || !(frame.w > 0) || !(frame.h > 0)) return [];
  const sx = frame.w / (dev[2] - dev[0]);
  const sy = frame.h / (dev[3] - dev[1]);
  const deviceToPoints: Matrix = { a: sx, b: 0, c: 0, d: sy, e: frame.x - dev[0] * sx, f: frame.y - dev[1] * sy };
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  const canvas = new Canvas({
    fonts,
    pageToPoints: (s) => compose(pageToDevice(s, header.pxPerMm), deviceToPoints),
    pixel: Math.sqrt(sx * sy),
    // The reference device's default font is about 12pt tall.
    defaultFontSize: (12 / 72) * 25.4 * header.pxPerMm.y * sy,
  });
  canvas.state.advanced = usesAdvancedMode(bytes, view);
  const player = new EmfPlayer(canvas, header);
  for (let off = 0; off + 8 <= bytes.length; ) {
    const type = view.getUint32(off, true);
    const size = view.getUint32(off + 4, true);
    if (size < 8 || off + size > bytes.length || type === EMR_EOF) break;
    try {
      player.play(type, new RecordReader(bytes, off, off + size, view));
    } catch {
      // A record we cannot make sense of is skipped; the rest of the picture still draws.
    }
    off += size;
  }
  return canvas.items;
}

class EmfPlayer {
  private readonly objects = new Map<number, GdiObject>();

  constructor(
    private readonly canvas: Canvas,
    private readonly header: EmfHeader,
  ) {}

  private get state(): DCState {
    return this.canvas.state;
  }

  play(type: number, r: RecordReader): void {
    const c = this.canvas;
    const s = this.state;
    switch (type) {
      case 2: // EMR_POLYBEZIER
      case 85:
        c.polyBezier(points(r, 24, type === 85));
        break;
      case 3: // EMR_POLYGON
      case 86:
        c.poly(points(r, 24, type === 86), true);
        break;
      case 4: // EMR_POLYLINE
      case 87:
        c.poly(points(r, 24, type === 87), false);
        break;
      case 5: // EMR_POLYBEZIERTO
      case 88:
        c.polyBezierTo(points(r, 24, type === 88));
        break;
      case 6: // EMR_POLYLINETO
      case 89:
        c.polylineTo(points(r, 24, type === 89));
        break;
      case 7: // EMR_POLYPOLYLINE
      case 90:
        c.polyPoly(polyPoints(r, type === 90), false);
        break;
      case 8: // EMR_POLYPOLYGON
      case 91:
        c.polyPoly(polyPoints(r, type === 91), true);
        break;
      case 56: // EMR_POLYDRAW
      case 92: {
        const pts = points(r, 24, type === 92);
        const at = 28 + pts.length * (type === 92 ? 4 : 8);
        c.polyDraw(pts, Array.from(r.slice(at, pts.length)));
        break;
      }
      case 9: // EMR_SETWINDOWEXTEX
        s.windowExt = { x: r.i32(8), y: r.i32(12) };
        break;
      case 10: // EMR_SETWINDOWORGEX
        s.windowOrg = { x: r.i32(8), y: r.i32(12) };
        break;
      case 11: // EMR_SETVIEWPORTEXTEX
        s.viewportExt = { x: r.i32(8), y: r.i32(12) };
        break;
      case 12: // EMR_SETVIEWPORTORGEX
        s.viewportOrg = { x: r.i32(8), y: r.i32(12) };
        break;
      case 31: // EMR_SCALEVIEWPORTEXTEX
      case 32: {
        const [xNum, xDen, yNum, yDen] = [r.i32(8), r.i32(12), r.i32(16), r.i32(20)];
        if (!xDen || !yDen) break;
        const ext = type === 31 ? s.viewportExt : s.windowExt;
        const scaled = { x: (ext.x * xNum) / xDen, y: (ext.y * yNum) / yDen };
        if (type === 31) s.viewportExt = scaled;
        else s.windowExt = scaled;
        break;
      }
      case 17: // EMR_SETMAPMODE
        s.mapMode = r.u32(8);
        break;
      case 18: // EMR_SETBKMODE
        s.bkMode = r.u32(8);
        break;
      case 19: // EMR_SETPOLYFILLMODE
        s.polyFillMode = r.u32(8);
        break;
      case 22: // EMR_SETTEXTALIGN
        s.textAlign = r.u32(8);
        break;
      case 24: // EMR_SETTEXTCOLOR
        s.textColor = color(r, 8);
        break;
      case 25: // EMR_SETBKCOLOR
        s.bkColor = color(r, 8);
        break;
      case 57: // EMR_SETARCDIRECTION
        s.arcDirection = r.u32(8);
        break;
      case 27: // EMR_MOVETOEX
        c.moveTo(r.i32(8), r.i32(12));
        break;
      case 54: // EMR_LINETO
        c.lineTo(r.i32(8), r.i32(12));
        break;
      case 33: // EMR_SAVEDC
        c.save();
        break;
      case 34: // EMR_RESTOREDC
        c.restore(r.i32(8));
        break;
      case 35: // EMR_SETWORLDTRANSFORM
        s.world = xform(r, 8);
        break;
      case 36: {
        // EMR_MODIFYWORLDTRANSFORM
        const x = xform(r, 8);
        switch (r.u32(32)) {
          case 1:
            s.world = IDENTITY;
            break;
          case 2:
            s.world = compose(x, s.world);
            break;
          case 3:
            s.world = compose(s.world, x);
            break;
          case 4:
            s.world = x;
            break;
        }
        break;
      }
      case 37: // EMR_SELECTOBJECT
        selectInto(s, this.object(r.u32(8)));
        break;
      case 40: // EMR_DELETEOBJECT
        this.objects.delete(r.u32(8));
        break;
      case 38: // EMR_CREATEPEN (LOGPEN: style, width as POINTL, colour)
        this.objects.set(r.u32(8), makePen(r.u32(12), r.i32(16), color(r, 24)));
        break;
      case 95:
        this.extCreatePen(r);
        break;
      case 39: // EMR_CREATEBRUSHINDIRECT
        this.objects.set(r.u32(8), makeBrush(r.u32(12), color(r, 16)));
        break;
      case 93: // EMR_CREATEMONOBRUSH
      case 94: {
        // EMR_CREATEDIBPATTERNBRUSHPT: drawn as a flat fill of the pattern's average colour.
        const img = decodeDib(r.slice(r.u32(16), r.u32(20)), r.slice(r.u32(24), r.u32(28)), { paletteIndices: r.u32(12) === 1 });
        this.objects.set(r.u32(8), { kind: 'brush', style: 'pattern', color: img ? averageColor(img) : '808080' });
        break;
      }
      case 82: // EMR_EXTCREATEFONTINDIRECTW
        this.objects.set(r.u32(8), logFont(r, 12));
        break;
      case 49: // EMR_CREATEPALETTE
      case 99: // EMR_CREATECOLORSPACE
      case 122: // EMR_CREATECOLORSPACEW
        this.objects.set(r.u32(8), { kind: 'other' });
        break;
      case 41: // EMR_ANGLEARC
        c.angleArc(r.i32(8), r.i32(12), r.u32(16), r.f32(20), r.f32(24));
        break;
      case 42: // EMR_ELLIPSE
        c.ellipse(r.i32(8), r.i32(12), r.i32(16), r.i32(20));
        break;
      case 43: // EMR_RECTANGLE
        c.rectangle(r.i32(8), r.i32(12), r.i32(16), r.i32(20));
        break;
      case 44: // EMR_ROUNDRECT
        c.roundRect(r.i32(8), r.i32(12), r.i32(16), r.i32(20), r.i32(24), r.i32(28));
        break;
      case 45: // EMR_ARC
      case 46: // EMR_CHORD
      case 47: // EMR_PIE
      case 55: {
        // EMR_ARCTO
        const kind: ArcKind = type === 45 ? 'arc' : type === 46 ? 'chord' : type === 47 ? 'pie' : 'arcto';
        c.arc(kind, r.i32(8), r.i32(12), r.i32(16), r.i32(20), r.i32(24), r.i32(28), r.i32(32), r.i32(36));
        break;
      }
      case 59: // EMR_BEGINPATH
        c.beginPath();
        break;
      case 60: // EMR_ENDPATH
        c.endPath();
        break;
      case 61: // EMR_CLOSEFIGURE
        c.closeFigure();
        break;
      case 62: // EMR_FILLPATH
        c.drawPath(true, false);
        break;
      case 63: // EMR_STROKEANDFILLPATH
        c.drawPath(true, true);
        break;
      case 64: // EMR_STROKEPATH
        c.drawPath(false, true);
        break;
      case 67: // EMR_SELECTCLIPPATH: clipping is not applied, but the path is used up.
      case 68: // EMR_ABORTPATH
        c.abortPath();
        break;
      case 83: // EMR_EXTTEXTOUTA
      case 84:
        this.extTextOut(r, type === 84);
        break;
      case 108:
        this.smallTextOut(r);
        break;
      case 76: // EMR_BITBLT
      case 77: // EMR_STRETCHBLT
      case 114: // EMR_ALPHABLEND
      case 116: // EMR_TRANSPARENTBLT
        this.blt(type, r);
        break;
      case 80:
        this.setDIBitsToDevice(r);
        break;
      case 81:
        this.stretchDIBits(r);
        break;
      // Everything else (clipping, palettes, colour management, comments with EMF+ data,
      // stretch mode, mapper flags, ROP2, miter limit, …) does not change what we draw.
    }
  }

  private object(index: number): GdiObject | undefined {
    if (index & 0x80000000) return stockObject(index & 0x7fffffff);
    return this.objects.get(index);
  }

  private extCreatePen(r: RecordReader): void {
    // EMR_EXTCREATEPEN: ihPen, DIB offsets, then LOGPENEX at 28.
    const style = r.u32(28);
    const width = r.u32(32);
    const brushStyle = r.u32(36);
    let penColor = color(r, 40);
    const count = r.fit(52, r.u32(48), 4);
    const dashes: number[] = [];
    for (let i = 0; i < count; i++) dashes.push(r.u32(52 + i * 4));
    if (brushStyle === BS_PATTERN || brushStyle === BS_DIBPATTERNPT) {
      const img = decodeDib(r.slice(r.u32(12), r.u32(16)), r.slice(r.u32(20), r.u32(24)));
      if (img) penColor = averageColor(img);
    }
    const pen = makePen(style, width, penColor, dashes);
    pen.cosmetic = (style & PS_GEOMETRIC) === 0;
    // A pen painted with a hollow brush draws nothing.
    if (brushStyle === BS_NULL) pen.style = PS_NULL;
    this.objects.set(r.u32(8), pen);
  }

  private extTextOut(r: RecordReader, wide: boolean): void {
    // EMR_EXTTEXTOUTA/W: Bounds, iGraphicsMode, exScale, eyScale, then EMRTEXT at 36.
    const x = r.i32(36);
    const y = r.i32(40);
    const count = r.u32(44);
    const offString = r.u32(48);
    const options = r.u32(52);
    let opaque: [number, number, number, number] | undefined;
    let offDx: number;
    if (options & ETO_NO_RECT) {
      offDx = r.u32(56);
    } else {
      if (options & ETO_OPAQUE) opaque = [r.i32(56), r.i32(60), r.i32(64), r.i32(68)];
      offDx = r.u32(72);
    }
    const n = r.fit(offString, count, wide ? 2 : 1);
    // With ETO_PDY each character has an x and a y advance; only x is used.
    const step = options & ETO_PDY ? 2 : 1;
    let dx: number[] | undefined;
    if (offDx > 0 && r.fit(offDx, n * step, 4) === n * step) {
      dx = [];
      for (let i = 0; i < n; i++) dx.push(r.i32(offDx + i * 4 * step));
    }
    // Glyph indices refer to the original font's glyph order, which we cannot map back to text.
    if (options & ETO_GLYPH_INDEX) {
      this.canvas.text(x, y, [], undefined, opaque);
      return;
    }
    const font = this.state.font;
    const decoded = wide ? decodeWide(units16(r, offString, n), font.faceName, dx) : decodeAnsi(r.slice(offString, n), font.charset, font.faceName, dx);
    this.canvas.text(x, y, decoded.chars, decoded.dx, opaque);
  }

  private smallTextOut(r: RecordReader): void {
    // EMR_SMALLTEXTOUT: x, y, cChars, fuOptions, iGraphicsMode, exScale, eyScale, [Bounds], text.
    const options = r.u32(20);
    const count = r.u32(16);
    const at = options & ETO_NO_RECT ? 36 : 52;
    if (options & ETO_GLYPH_INDEX) return;
    const font = this.state.font;
    const decoded =
      options & ETO_SMALL_CHARS
        ? decodeWide(Array.from(r.slice(at, count)), font.faceName, undefined)
        : decodeWide(units16(r, at, r.fit(at, count, 2)), font.faceName, undefined);
    this.canvas.text(r.i32(8), r.i32(12), decoded.chars, undefined);
  }

  private stretchDIBits(r: RecordReader): void {
    const [xDest, yDest, xSrc, ySrc, cxSrc, cySrc] = [r.i32(24), r.i32(28), r.i32(32), r.i32(36), r.i32(40), r.i32(44)];
    const [offBmi, cbBmi, offBits, cbBits, usage, rop] = [r.u32(48), r.u32(52), r.u32(56), r.u32(60), r.u32(64), r.u32(68)];
    const [cxDest, cyDest] = [r.i32(72), r.i32(76)];
    if (!cbBmi) {
      this.patBlt(rop, xDest, yDest, cxDest, cyDest);
      return;
    }
    if (rop !== SRCCOPY) return;
    const info = r.slice(offBmi, cbBmi);
    const img = decodeDib(info, r.slice(offBits, cbBits), { paletteIndices: usage === 1 });
    if (!img) return;
    // StretchDIBits measures the source rectangle from the bottom of a bottom-up DIB.
    const top = dibIsBottomUp(info) ? img.height - ySrc - cySrc : ySrc;
    this.canvas.image(cropRaster(img, xSrc, top, cxSrc, cySrc), xDest, yDest, cxDest, cyDest);
  }

  private setDIBitsToDevice(r: RecordReader): void {
    const [xDest, yDest, xSrc, ySrc, cxSrc, cySrc] = [r.i32(24), r.i32(28), r.i32(32), r.i32(36), r.i32(40), r.i32(44)];
    const info = r.slice(r.u32(48), r.u32(52));
    const img = decodeDib(info, r.slice(r.u32(56), r.u32(60)), { paletteIndices: r.u32(64) === 1 });
    if (!img) return;
    // The destination size is in device pixels: turn it back into logical units.
    const pd = pageToDevice(this.state, this.header.pxPerMm);
    const top = dibIsBottomUp(info) ? img.height - ySrc - cySrc : ySrc;
    this.canvas.image(cropRaster(img, xSrc, top, cxSrc, cySrc), xDest, yDest, cxSrc / (pd.a || 1), cySrc / (pd.d || 1));
  }

  private blt(type: number, r: RecordReader): void {
    const [xDest, yDest, cxDest, cyDest, op] = [r.i32(24), r.i32(28), r.i32(32), r.i32(36), r.u32(40)];
    const [xSrc, ySrc] = [r.i32(44), r.i32(48)];
    const src = xform(r, 52);
    const [usage, offBmi, cbBmi, offBits, cbBits] = [r.u32(80), r.u32(84), r.u32(88), r.u32(92), r.u32(96)];
    if (!cbBmi) {
      if (type === 76 || type === 77) this.patBlt(op, xDest, yDest, cxDest, cyDest);
      return;
    }
    // For BitBlt/StretchBlt the dword at 40 is the raster operation; AlphaBlend and TransparentBlt put their blend/key there.
    if ((type === 76 || type === 77) && op !== SRCCOPY) return;
    const alphaFormat = type === 114 ? (op >>> 24) & 0xff : 0;
    let img: RasterImage | undefined = decodeDib(r.slice(offBmi, cbBmi), r.slice(offBits, cbBits), { paletteIndices: usage === 1, alpha: alphaFormat === 1 });
    if (!img) return;
    // The source rectangle is in the source DC's logical units, mapped to bitmap pixels by XformSrc (top-down).
    const cxSrc = type === 76 ? cxDest : r.i32(100);
    const cySrc = type === 76 ? cyDest : r.i32(104);
    const sx = src.a || 1;
    const sy = src.d || 1;
    img = cropRaster(img, xSrc * sx + src.e, ySrc * sy + src.f, cxSrc * sx, cySrc * sy);
    if (type === 114) img = withConstantAlpha(img, (op >>> 16) & 0xff);
    if (type === 116) img = keyOut(img, colorRef(op & 0xff, (op >>> 8) & 0xff, (op >>> 16) & 0xff));
    this.canvas.image(img, xDest, yDest, cxDest, cyDest);
  }

  /** Raster operations without a source bitmap: brush, black or white fills. */
  private patBlt(op: number, x: number, y: number, w: number, h: number): void {
    const fill = op === PATCOPY ? undefined : op === BLACKNESS ? '000000' : op === WHITENESS ? 'FFFFFF' : null;
    if (fill === null) return;
    this.canvas.fillRect(x, y, x + w, y + h, fill);
  }
}

function withConstantAlpha(img: RasterImage, alpha: number): RasterImage {
  if (alpha >= 255) return img;
  const rgba = img.rgba.slice();
  for (let i = 3; i < rgba.length; i += 4) rgba[i] = Math.round((rgba[i]! * alpha) / 255);
  return { width: img.width, height: img.height, rgba };
}

function color(r: RecordReader, at: number): string {
  return colorRef(r.u8(at), r.u8(at + 1), r.u8(at + 2));
}

function xform(r: RecordReader, at: number): Matrix {
  return { a: r.f32(at), b: r.f32(at + 4), c: r.f32(at + 8), d: r.f32(at + 12), e: r.f32(at + 16), f: r.f32(at + 20) };
}

function units16(r: RecordReader, at: number, n: number): number[] {
  const out: number[] = [];
  for (let i = 0; i < n; i++) out.push(r.u16(at + i * 2));
  return out;
}

/** Count at `at`, then that many POINTL (or POINTS when `short`). */
function points(r: RecordReader, at: number, short: boolean): Pt[] {
  const size = short ? 4 : 8;
  const n = r.fit(at + 4, r.u32(at), size);
  return readPoints(r, at + 4, n, short);
}

function readPoints(r: RecordReader, at: number, n: number, short: boolean): Pt[] {
  const out: Pt[] = [];
  for (let i = 0; i < n; i++) {
    out.push(short ? { x: r.i16(at + i * 4), y: r.i16(at + i * 4 + 2) } : { x: r.i32(at + i * 8), y: r.i32(at + i * 8 + 4) });
  }
  return out;
}

/** EMR_POLYPOLY*: Bounds, number of polys, total points, a count per poly, then the points. */
function polyPoints(r: RecordReader, short: boolean): Pt[][] {
  const polys = r.fit(32, r.u32(24), 4);
  const total = r.u32(28);
  const at = 32 + polys * 4;
  const all = readPoints(r, at, r.fit(at, total, short ? 4 : 8), short);
  const out: Pt[][] = [];
  let k = 0;
  for (let i = 0; i < polys; i++) {
    const n = r.u32(32 + i * 4);
    out.push(all.slice(k, k + n));
    k += n;
  }
  return out;
}

/** LOGFONTW at `at`. */
function logFont(r: RecordReader, at: number): Font {
  let faceName = '';
  for (let i = 0; i < 32; i++) {
    const u = r.u16(at + 28 + i * 2);
    if (!u) break;
    faceName += String.fromCharCode(u);
  }
  return {
    kind: 'font',
    height: r.i32(at),
    weight: r.i32(at + 16),
    italic: r.u8(at + 20) !== 0,
    underline: r.u8(at + 21) !== 0,
    strikeOut: r.u8(at + 22) !== 0,
    charset: r.u8(at + 23),
    faceName,
  };
}
