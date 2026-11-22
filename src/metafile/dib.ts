/**
 * Device-independent bitmaps (BITMAPINFO header, colour table, pixels) as EMF and WMF records
 * embed them, decoded to RGBA. Covers 1/4/8/16/24/32 bits per pixel with BI_RGB, BI_BITFIELDS
 * and the RLE4/RLE8 run-length forms old WMFs use. JPEG/PNG-compressed DIBs are not decoded.
 */
import type { RasterImage } from '../layout/types.js';
import { colorRef } from './state.js';

const BI_RGB = 0;
const BI_RLE8 = 1;
const BI_RLE4 = 2;
const BI_BITFIELDS = 3;
const BI_ALPHABITFIELDS = 6;

/** Refuse pictures whose header would make us allocate absurd amounts of memory. */
const MAX_PIXELS = 1 << 26;

export interface DibOptions {
  /** DIB_PAL_COLORS: the colour table holds palette indices, not colours. */
  paletteIndices?: boolean;
  /** Take 32-bpp alpha as premultiplied (AlphaBlend with AC_SRC_ALPHA); otherwise that byte is ignored. */
  alpha?: boolean;
}

interface Header {
  width: number;
  height: number;
  topDown: boolean;
  bpp: number;
  compression: number;
  masks: [number, number, number, number] | undefined;
  palette: Uint8Array[];
  /** Offset of the pixels when they directly follow the header (packed DIB). */
  bitsOffset: number;
}

function readHeader(info: Uint8Array, options: DibOptions): Header | undefined {
  if (info.length < 12) return undefined;
  const view = new DataView(info.buffer, info.byteOffset, info.byteLength);
  const size = view.getUint32(0, true);
  let width: number;
  let rawHeight: number;
  let bpp: number;
  let compression = BI_RGB;
  let used = 0;
  let entrySize = 4;
  if (size === 12) {
    // BITMAPCOREHEADER (OS/2 1.x): 16-bit sizes and RGBTRIPLE colours.
    width = view.getUint16(4, true);
    rawHeight = view.getUint16(6, true);
    bpp = view.getUint16(10, true);
    entrySize = 3;
  } else {
    if (size < 40 || info.length < 40) return undefined;
    width = view.getInt32(4, true);
    rawHeight = view.getInt32(8, true);
    bpp = view.getUint16(14, true);
    compression = view.getUint32(16, true);
    used = view.getUint32(32, true);
  }
  const height = Math.abs(rawHeight);
  if (width <= 0 || height <= 0 || width * height > MAX_PIXELS) return undefined;
  let tableOffset = size;
  let masks: Header['masks'];
  if (compression === BI_BITFIELDS || compression === BI_ALPHABITFIELDS) {
    const count = compression === BI_ALPHABITFIELDS ? 4 : 3;
    // BITMAPV2+ headers carry the masks inside; a plain BITMAPINFOHEADER is followed by them.
    const at = size >= 52 ? 40 : size;
    const mask = (i: number): number => (at + i * 4 + 4 <= info.length ? view.getUint32(at + i * 4, true) : 0);
    masks = [mask(0), mask(1), mask(2), count === 4 || size >= 56 ? mask(3) : 0];
    if (size < 52) tableOffset += count * 4;
  } else if (bpp === 16) {
    masks = [0x7c00, 0x03e0, 0x001f, 0];
  } else if (bpp === 32) {
    masks = [0xff0000, 0x00ff00, 0x0000ff, 0];
  }
  // Above 8 bpp a colour table is only an optimisation hint, but it still sits before packed pixels.
  const colors = Math.min(used > 0 ? used : bpp <= 8 ? 1 << bpp : 0, 1 << 16);
  const palette: Uint8Array[] = [];
  if (options.paletteIndices && bpp <= 8) {
    // Indices into a logical palette we do not track: a grey ramp keeps the image legible.
    for (let i = 0; i < Math.min(colors, 256); i++) {
      const v = colors > 1 ? Math.round((i * 255) / (Math.min(colors, 256) - 1)) : 0;
      palette.push(Uint8Array.of(v, v, v));
    }
    tableOffset += colors * 2;
  } else {
    for (let i = 0; i < Math.min(colors, 256); i++) {
      const o = tableOffset + i * entrySize;
      palette.push(o + 3 <= info.length ? Uint8Array.of(info[o + 2]!, info[o + 1]!, info[o]!) : Uint8Array.of(0, 0, 0));
    }
    tableOffset += colors * entrySize;
  }
  return { width, height, topDown: rawHeight < 0, bpp, compression, masks, palette, bitsOffset: tableOffset };
}

function maskShift(mask: number): { shift: number; max: number } {
  if (!mask) return { shift: 0, max: 0 };
  let shift = 0;
  while (((mask >>> shift) & 1) === 0) shift++;
  return { shift, max: mask >>> shift };
}

/** Run-length data to one palette index per pixel, rows bottom-up as stored. */
function decodeRle(data: Uint8Array, width: number, height: number, four: boolean): Uint8Array {
  const out = new Uint8Array(width * height);
  let x = 0;
  let y = 0;
  let i = 0;
  const put = (v: number): void => {
    if (x < width && y < height) out[y * width + x] = v;
    x++;
  };
  while (i + 1 < data.length && y < height) {
    const a = data[i++]!;
    const b = data[i++]!;
    if (a > 0) {
      for (let k = 0; k < a; k++) put(four ? (k & 1 ? b & 0x0f : b >> 4) : b);
    } else if (b === 0) {
      x = 0;
      y++;
    } else if (b === 1) {
      break;
    } else if (b === 2) {
      x += data[i++] ?? 0;
      y += data[i++] ?? 0;
    } else {
      // Absolute run of b pixels, padded to a 16-bit boundary.
      const bytes = four ? (b + 1) >> 1 : b;
      for (let k = 0; k < b; k++) {
        const v = data[i + (four ? k >> 1 : k)] ?? 0;
        put(four ? (k & 1 ? v & 0x0f : v >> 4) : v);
      }
      i += bytes + (bytes & 1);
    }
  }
  return out;
}

/**
 * Decode a DIB. `info` holds the BITMAPINFO; `bits` the pixels, or undefined when they follow the
 * header and colour table in `info` (a packed DIB, as WMF records store it).
 */
export function decodeDib(info: Uint8Array, bits: Uint8Array | undefined, options: DibOptions = {}): RasterImage | undefined {
  const h = readHeader(info, options);
  if (!h) return undefined;
  const data = bits ?? info.subarray(Math.min(info.length, h.bitsOffset));
  const { width, height, bpp, palette } = h;
  const rgba = new Uint8Array(width * height * 4);
  const color = (index: number, o: number): void => {
    const c = palette[index];
    if (c) {
      rgba[o] = c[0]!;
      rgba[o + 1] = c[1]!;
      rgba[o + 2] = c[2]!;
    }
    rgba[o + 3] = 255;
  };

  if (h.compression === BI_RLE8 || h.compression === BI_RLE4) {
    const indices = decodeRle(data, width, height, h.compression === BI_RLE4);
    for (let y = 0; y < height; y++) {
      const src = (height - 1 - y) * width;
      for (let x = 0; x < width; x++) color(indices[src + x]!, (y * width + x) * 4);
    }
    return { width, height, rgba };
  }
  if (h.compression !== BI_RGB && h.compression !== BI_BITFIELDS && h.compression !== BI_ALPHABITFIELDS) return undefined;
  if (![1, 4, 8, 16, 24, 32].includes(bpp)) return undefined;

  const stride = (((width * bpp + 31) >> 5) << 2) >>> 0;
  const channels = h.masks?.map(maskShift);
  let anyAlpha = false;
  for (let y = 0; y < height; y++) {
    const row = (h.topDown ? y : height - 1 - y) * stride;
    if (row >= data.length) break;
    for (let x = 0; x < width; x++) {
      const o = (y * width + x) * 4;
      switch (bpp) {
        case 1:
          color(((data[row + (x >> 3)] ?? 0) >> (7 - (x & 7))) & 1, o);
          break;
        case 4: {
          const v = data[row + (x >> 1)] ?? 0;
          color(x & 1 ? v & 0x0f : v >> 4, o);
          break;
        }
        case 8:
          color(data[row + x] ?? 0, o);
          break;
        case 24: {
          const p = row + x * 3;
          rgba[o] = data[p + 2] ?? 0;
          rgba[o + 1] = data[p + 1] ?? 0;
          rgba[o + 2] = data[p] ?? 0;
          rgba[o + 3] = 255;
          break;
        }
        default: {
          // 16 and 32 bits: channels by mask.
          const p = row + (x * bpp) / 8;
          const v = bpp === 16 ? (data[p] ?? 0) | ((data[p + 1] ?? 0) << 8) : ((data[p] ?? 0) | ((data[p + 1] ?? 0) << 8) | ((data[p + 2] ?? 0) << 16) | ((data[p + 3] ?? 0) << 24)) >>> 0;
          const ch = (i: number): number => {
            const c = channels![i]!;
            return c.max ? Math.round((((v >>> c.shift) & c.max) * 255) / c.max) : 0;
          };
          rgba[o] = ch(0);
          rgba[o + 1] = ch(1);
          rgba[o + 2] = ch(2);
          if (options.alpha && bpp === 32) {
            const a = channels![3]!.max ? ch(3) : (v >>> 24) & 0xff;
            rgba[o + 3] = a;
            if (a) anyAlpha = true;
          } else {
            rgba[o + 3] = 255;
          }
        }
      }
    }
  }
  if (options.alpha && bpp === 32) {
    if (!anyAlpha) {
      // An all-zero alpha channel means the bitmap never had one: draw it opaque.
      for (let i = 3; i < rgba.length; i += 4) rgba[i] = 255;
    } else {
      for (let i = 0; i < rgba.length; i += 4) {
        const a = rgba[i + 3]!;
        if (a && a < 255) for (let k = 0; k < 3; k++) rgba[i + k] = Math.min(255, Math.round((rgba[i + k]! * 255) / a));
      }
    }
  }
  return { width, height, rgba };
}

/** The mean colour of a bitmap (pattern brushes are drawn as a flat fill of it). */
export function averageColor(img: RasterImage): string {
  let r = 0;
  let g = 0;
  let b = 0;
  const n = img.width * img.height;
  for (let i = 0; i < n * 4; i += 4) {
    r += img.rgba[i]!;
    g += img.rgba[i + 1]!;
    b += img.rgba[i + 2]!;
  }
  return n ? colorRef(r / n, g / n, b / n) : '808080';
}

/** The part of a bitmap inside the rectangle (top-down pixel coordinates), clipped to the bitmap. */
export function cropRaster(img: RasterImage, x: number, y: number, w: number, h: number): RasterImage {
  const x0 = Math.max(0, Math.min(img.width, Math.round(x)));
  const y0 = Math.max(0, Math.min(img.height, Math.round(y)));
  const x1 = Math.max(x0, Math.min(img.width, Math.round(x + w)));
  const y1 = Math.max(y0, Math.min(img.height, Math.round(y + h)));
  if (x0 === 0 && y0 === 0 && x1 === img.width && y1 === img.height) return img;
  if (x1 <= x0 || y1 <= y0) return img;
  const width = x1 - x0;
  const height = y1 - y0;
  const rgba = new Uint8Array(width * height * 4);
  for (let row = 0; row < height; row++) {
    const from = ((y0 + row) * img.width + x0) * 4;
    rgba.set(img.rgba.subarray(from, from + width * 4), row * width * 4);
  }
  return { width, height, rgba };
}

/** Mirror a bitmap, for destinations whose width or height runs backwards. */
export function flipRaster(img: RasterImage, horizontal: boolean, vertical: boolean): RasterImage {
  if (!horizontal && !vertical) return img;
  const { width, height } = img;
  const rgba = new Uint8Array(img.rgba.length);
  for (let y = 0; y < height; y++) {
    const sy = vertical ? height - 1 - y : y;
    for (let x = 0; x < width; x++) {
      const sx = horizontal ? width - 1 - x : x;
      const s = (sy * width + sx) * 4;
      const d = (y * width + x) * 4;
      rgba[d] = img.rgba[s]!;
      rgba[d + 1] = img.rgba[s + 1]!;
      rgba[d + 2] = img.rgba[s + 2]!;
      rgba[d + 3] = img.rgba[s + 3]!;
    }
  }
  return { width, height, rgba };
}

/** Make pixels of one colour transparent (TransparentBlt). */
export function keyOut(img: RasterImage, color: string): RasterImage {
  const r = Number.parseInt(color.slice(0, 2), 16);
  const g = Number.parseInt(color.slice(2, 4), 16);
  const b = Number.parseInt(color.slice(4, 6), 16);
  const rgba = img.rgba.slice();
  for (let i = 0; i < rgba.length; i += 4) if (rgba[i] === r && rgba[i + 1] === g && rgba[i + 2] === b) rgba[i + 3] = 0;
  return { width: img.width, height: img.height, rgba };
}

/** Whether a DIB stores its rows bottom-up (positive height), which puts StretchDIBits' source origin at the bottom. */
export function dibIsBottomUp(info: Uint8Array): boolean {
  if (info.length < 12) return true;
  const view = new DataView(info.buffer, info.byteOffset, info.byteLength);
  const size = view.getUint32(0, true);
  if (size === 12) return true;
  return info.length < 12 || view.getInt32(8, true) > 0;
}
