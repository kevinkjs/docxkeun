/**
 * GIF, BMP and TIFF decoders, checked pixel by pixel against files built byte by byte here (with
 * small encoders, including LZW for GIF and TIFF), plus format sniffing and PDF embedding.
 */
import { deflateSync } from 'node:zlib';
import { describe, expect, it } from 'vitest';
import { browserShows, decodeBmp, decodeGif, decodeRaster, decodeTiff, sniffImage } from '../src/images/index.js';
import { addImage, ImageFormatError } from '../src/pdf/images.js';
import { PdfWriter } from '../src/pdf/writer.js';

type Rgb = [number, number, number];

/** Deterministic pseudo-random bytes, so failures reproduce. */
function random(seed: number): () => number {
  let s = seed >>> 0;
  return () => {
    s = (Math.imul(s, 1103515245) + 12345) >>> 0;
    return s >>> 24;
  };
}

/** Expected RGBA for palette indices; `transparent` (or -1) gives a fully transparent pixel. */
function paletteRgba(palette: readonly Rgb[], indices: readonly number[], transparent = -1): Uint8Array {
  const out = new Uint8Array(indices.length * 4);
  indices.forEach((idx, i) => {
    if (idx === transparent) return;
    const c = palette[idx]!;
    out.set([c[0], c[1], c[2], 255], i * 4);
  });
  return out;
}

const ascii = (s: string) => Array.from(s, (c) => c.charCodeAt(0));

// ---------------------------------------------------------------------------------------------
// GIF

/** GIF LZW: LSB-first codes, width grows once the next code no longer fits. */
function gifLzw(indices: readonly number[], minCodeSize: number, resetWhenFull = true): number[] {
  const clear = 1 << minCodeSize;
  let codeSize = minCodeSize + 1;
  let next = clear + 2;
  let dict = new Map<number, number>();
  const out: number[] = [];
  let acc = 0;
  let bits = 0;
  const emit = (code: number) => {
    acc |= code << bits;
    bits += codeSize;
    while (bits >= 8) {
      out.push(acc & 255);
      acc >>>= 8;
      bits -= 8;
    }
  };
  emit(clear);
  let w = indices[0]!;
  for (let i = 1; i < indices.length; i++) {
    const k = indices[i]!;
    const key = w * 256 + k;
    const hit = dict.get(key);
    if (hit !== undefined) {
      w = hit;
      continue;
    }
    emit(w);
    if (next === 4096) {
      if (resetWhenFull) {
        emit(clear);
        dict = new Map();
        codeSize = minCodeSize + 1;
        next = clear + 2;
      }
    } else {
      dict.set(key, next++);
      if (next > 1 << codeSize && codeSize < 12) codeSize++;
    }
    w = k;
  }
  emit(w);
  // The decoder adds an entry when it reads the last code, which can widen the end code.
  if (next === 1 << codeSize && codeSize < 12) codeSize++;
  emit(clear + 1);
  if (bits > 0) out.push(acc & 255);
  return out;
}

interface GifOptions {
  width: number;
  height: number;
  indices: number[];
  frame?: { left: number; top: number; width: number; height: number };
  globalPalette?: Rgb[];
  localPalette?: Rgb[];
  transparent?: number;
  interlace?: boolean;
  minCodeSize?: number;
  resetWhenFull?: boolean;
  version?: '87a' | '89a';
  secondFrame?: boolean;
}

function gif(o: GifOptions): Uint8Array {
  const b: number[] = [];
  const u16 = (v: number) => b.push(v & 255, v >> 8);
  const tableBits = (pal: Rgb[]) => {
    let n = 0;
    while (2 << n < pal.length) n++;
    return n;
  };
  const table = (pal: Rgb[]) => {
    for (let i = 0; i < 2 << tableBits(pal); i++) b.push(...(pal[i] ?? [0, 0, 0]));
  };
  b.push(...ascii(`GIF${o.version ?? '89a'}`));
  u16(o.width);
  u16(o.height);
  if (o.globalPalette) {
    b.push(0x80 | 0x70 | tableBits(o.globalPalette), 0, 0);
    table(o.globalPalette);
  } else b.push(0, 0, 0);
  if (o.version !== '87a') {
    // An application extension the decoder must skip, then the graphic control extension.
    b.push(0x21, 0xff, 11, ...ascii('NETSCAPE2.0'), 3, 1, 0, 0, 0);
    if (o.transparent !== undefined) b.push(0x21, 0xf9, 4, 0x01, 0, 0, o.transparent, 0);
  }
  const frame = o.frame ?? { left: 0, top: 0, width: o.width, height: o.height };
  const image = (indices: number[]) => {
    b.push(0x2c);
    u16(frame.left);
    u16(frame.top);
    u16(frame.width);
    u16(frame.height);
    b.push((o.localPalette ? 0x80 | tableBits(o.localPalette) : 0) | (o.interlace ? 0x40 : 0));
    if (o.localPalette) table(o.localPalette);
    let ordered = indices;
    if (o.interlace) {
      ordered = [];
      for (const [start, step] of [
        [0, 8],
        [4, 8],
        [2, 4],
        [1, 2],
      ] as const) {
        for (let y = start; y < frame.height; y += step) ordered.push(...indices.slice(y * frame.width, (y + 1) * frame.width));
      }
    }
    const max = Math.max(...indices);
    const min = o.minCodeSize ?? Math.max(2, Math.ceil(Math.log2(max + 1)));
    const data = gifLzw(ordered, min, o.resetWhenFull);
    b.push(min);
    for (let i = 0; i < data.length; i += 255) {
      const chunk = data.slice(i, i + 255);
      b.push(chunk.length, ...chunk);
    }
    b.push(0);
  };
  image(o.indices);
  if (o.secondFrame) image(o.indices.map(() => 0));
  b.push(0x3b);
  return Uint8Array.from(b);
}

const RGBW: Rgb[] = [
  [255, 0, 0],
  [0, 255, 0],
  [0, 0, 255],
  [255, 255, 255],
];

describe('GIF decoder', () => {
  it('decodes a global-palette image', () => {
    const indices = [0, 1, 2, 3, 2, 1];
    const img = decodeGif(gif({ width: 3, height: 2, indices, globalPalette: RGBW }));
    expect([img.width, img.height]).toEqual([3, 2]);
    expect(img.rgba).toEqual(paletteRgba(RGBW, indices));
  });

  it('decodes GIF87a and ignores frames after the first', () => {
    const indices = [3, 2, 1, 0];
    const img = decodeGif(gif({ width: 2, height: 2, indices, globalPalette: RGBW, version: '87a', secondFrame: true }));
    expect(img.rgba).toEqual(paletteRgba(RGBW, indices));
  });

  it('prefers the local colour table and honours the transparent index', () => {
    const local: Rgb[] = [
      [10, 20, 30],
      [40, 50, 60],
      [70, 80, 90],
      [100, 110, 120],
    ];
    const indices = [0, 1, 2, 3, 1, 1];
    const img = decodeGif(gif({ width: 3, height: 2, indices, globalPalette: RGBW, localPalette: local, transparent: 1 }));
    expect(img.rgba).toEqual(paletteRgba(local, indices, 1));
  });

  it('de-interlaces rows', () => {
    const width = 3;
    const height = 11;
    const indices = Array.from({ length: width * height }, (_, i) => Math.floor(i / width) % 4);
    const img = decodeGif(gif({ width, height, indices, globalPalette: RGBW, interlace: true }));
    expect(img.rgba).toEqual(paletteRgba(RGBW, indices));
  });

  it('places a smaller frame on a transparent canvas of the screen size', () => {
    const img = decodeGif(gif({ width: 4, height: 3, indices: [0, 1, 2, 3], frame: { left: 1, top: 1, width: 2, height: 2 }, globalPalette: RGBW }));
    const T = -1;
    const canvas = [T, T, T, T, T, 0, 1, T, T, 2, 3, T];
    expect(img.rgba).toEqual(paletteRgba(RGBW, canvas, T));
  });

  it('handles code-size growth up to 12 bits and clear codes on a full table', () => {
    const palette: Rgb[] = Array.from({ length: 256 }, (_, i) => [i, 255 - i, (i * 7) & 255]);
    const next = random(1);
    const width = 120;
    const height = 90;
    const indices = Array.from({ length: width * height }, () => next());
    const bytes = gif({ width, height, indices, globalPalette: palette });
    expect(decodeGif(bytes).rgba).toEqual(paletteRgba(palette, indices));
    // Some encoders never clear a full table and keep emitting 12-bit codes.
    const deferred = gif({ width, height, indices, globalPalette: palette, resetWhenFull: false });
    expect(deferred).not.toEqual(bytes);
    expect(decodeGif(deferred).rgba).toEqual(paletteRgba(palette, indices));
  });

  it('decodes long runs (codes that refer to the entry being defined)', () => {
    const indices = new Array<number>(64 * 64).fill(2);
    indices[100] = 1;
    const img = decodeGif(gif({ width: 64, height: 64, indices, globalPalette: RGBW }));
    expect(img.rgba).toEqual(paletteRgba(RGBW, indices));
  });

  it('decodes a 2-colour image with the minimum code size', () => {
    const pal: Rgb[] = [
      [0, 0, 0],
      [255, 255, 255],
    ];
    const indices = [0, 1, 1, 0, 1, 0, 0, 1, 1];
    const img = decodeGif(gif({ width: 3, height: 3, indices, globalPalette: pal }));
    expect(img.rgba).toEqual(paletteRgba(pal, indices));
  });

  it('rejects data that is not a GIF', () => {
    expect(() => decodeGif(Uint8Array.from(ascii('GIF90a1234567890')))).toThrow(/GIF/);
  });
});

// ---------------------------------------------------------------------------------------------
// BMP

interface BmpOptions {
  width: number;
  /** Negative for top-down rows. */
  height: number;
  bpp: number;
  compression?: number;
  headerSize?: number;
  palette?: Rgb[];
  masks?: number[];
  colorsUsed?: number;
  /** Pixel data exactly as stored. */
  data: Uint8Array;
}

function bmp(o: BmpOptions): Uint8Array {
  const hs = o.headerSize ?? 40;
  const maskBytes = hs === 40 && o.masks ? o.masks.length * 4 : 0;
  const entry = hs === 12 ? 3 : 4;
  const pal = o.palette ?? [];
  const offset = 14 + hs + maskBytes + pal.length * entry;
  const out = new Uint8Array(offset + o.data.length);
  const v = new DataView(out.buffer);
  out[0] = 0x42;
  out[1] = 0x4d;
  v.setUint32(2, out.length, true);
  v.setUint32(10, offset, true);
  v.setUint32(14, hs, true);
  if (hs === 12) {
    v.setUint16(18, o.width, true);
    v.setUint16(20, o.height, true);
    v.setUint16(22, 1, true);
    v.setUint16(24, o.bpp, true);
  } else {
    v.setInt32(18, o.width, true);
    v.setInt32(22, o.height, true);
    v.setUint16(26, 1, true);
    v.setUint16(28, o.bpp, true);
    v.setUint32(30, o.compression ?? 0, true);
    v.setUint32(34, o.data.length, true);
    v.setUint32(46, o.colorsUsed ?? 0, true);
    // Inside a V2+ header and right after a 40-byte one, the masks start at the same offset.
    o.masks?.forEach((m, i) => v.setUint32(54 + i * 4, m, true));
  }
  pal.forEach(([r, g, b], i) => out.set([b, g, r], 14 + hs + maskBytes + i * entry));
  out.set(o.data, offset);
  return out;
}

/** Rows of bytes, each padded to a 4-byte boundary. */
function padRows(rows: number[][]): Uint8Array {
  const out: number[] = [];
  for (const row of rows) {
    out.push(...row);
    for (let i = row.length; i % 4; i++) out.push(0);
  }
  return Uint8Array.from(out);
}

/** Pack palette indices MSB-first into bytes. */
function packBitsRow(indices: number[], bpp: number): number[] {
  const out = new Array<number>(Math.ceil((indices.length * bpp) / 8)).fill(0);
  indices.forEach((v, x) => {
    const bit = x * bpp;
    out[bit >> 3]! |= v << (8 - bpp - (bit & 7));
  });
  return out;
}

const le16 = (v: number) => [v & 255, v >> 8];
const le32 = (v: number) => [v & 255, (v >> 8) & 255, (v >> 16) & 255, v >>> 24];
const rgba = (list: number[][]) => Uint8Array.from(list.flat());

describe('BMP decoder', () => {
  it('decodes bottom-up 24-bit rows with padding', () => {
    // Stored bottom row first; 3 pixels = 9 bytes, padded to 12.
    const data = padRows([
      [1, 2, 3, 4, 5, 6, 7, 8, 9],
      [10, 20, 30, 40, 50, 60, 70, 80, 90],
    ]);
    const img = decodeBmp(bmp({ width: 3, height: 2, bpp: 24, data }));
    expect(img.rgba).toEqual(
      rgba([
        [30, 20, 10, 255],
        [60, 50, 40, 255],
        [90, 80, 70, 255],
        [3, 2, 1, 255],
        [6, 5, 4, 255],
        [9, 8, 7, 255],
      ]),
    );
  });

  it('decodes a top-down 8-bit palette image with a short palette', () => {
    const palette: Rgb[] = [
      [255, 0, 0],
      [0, 128, 0],
      [0, 0, 64],
    ];
    const rows = [
      [0, 1, 2],
      [2, 2, 0],
    ];
    const img = decodeBmp(bmp({ width: 3, height: -2, bpp: 8, palette, colorsUsed: 3, data: padRows(rows) }));
    expect(img.rgba).toEqual(paletteRgba(palette, rows.flat()));
  });

  it('decodes 1-, 2- and 4-bit palette images', () => {
    const palette: Rgb[] = Array.from({ length: 16 }, (_, i) => [i * 16, 255 - i * 16, i]);
    for (const bpp of [1, 2, 4]) {
      const width = 11;
      const top = Array.from({ length: width }, (_, i) => (i * 3) % (1 << bpp));
      const bottom = Array.from({ length: width }, (_, i) => (i * 5 + 1) % (1 << bpp));
      const data = padRows([packBitsRow(bottom, bpp), packBitsRow(top, bpp)]);
      const img = decodeBmp(bmp({ width, height: 2, bpp, palette: palette.slice(0, 1 << bpp), data }));
      expect(img.rgba).toEqual(paletteRgba(palette, [...top, ...bottom]));
    }
  });

  it('decodes 16-bit 5-5-5 (BI_RGB) and 5-6-5 (BI_BITFIELDS) pixels', () => {
    const rgb555 = (r: number, g: number, b: number) => (r << 10) | (g << 5) | b;
    const img555 = decodeBmp(bmp({ width: 2, height: -1, bpp: 16, data: padRows([[...le16(rgb555(31, 0, 16)), ...le16(rgb555(1, 31, 0))]]) }));
    expect(img555.rgba).toEqual(
      rgba([
        [255, 0, 132, 255],
        [8, 255, 0, 255],
      ]),
    );

    const rgb565 = (r: number, g: number, b: number) => (r << 11) | (g << 5) | b;
    const data = padRows([[...le16(rgb565(31, 63, 0)), ...le16(rgb565(0, 32, 31)), ...le16(rgb565(10, 1, 5))]]);
    const img565 = decodeBmp(bmp({ width: 3, height: 1, bpp: 16, compression: 3, masks: [0xf800, 0x07e0, 0x001f], data }));
    expect(img565.rgba).toEqual(
      rgba([
        [255, 255, 0, 255],
        [0, 130, 255, 255],
        [82, 4, 41, 255],
      ]),
    );
  });

  it('decodes 16-bit ALPHABITFIELDS (4-4-4-4) with alpha', () => {
    const px = (a: number, r: number, g: number, b: number) => (a << 12) | (r << 8) | (g << 4) | b;
    const data = padRows([[...le16(px(15, 15, 0, 0)), ...le16(px(8, 0, 15, 3))]]);
    const img = decodeBmp(bmp({ width: 2, height: 1, bpp: 16, compression: 6, masks: [0x0f00, 0x00f0, 0x000f, 0xf000], data }));
    expect(img.rgba).toEqual(
      rgba([
        [255, 0, 0, 255],
        [0, 255, 51, 136],
      ]),
    );
  });

  it('uses the fourth byte of 32-bit BI_RGB as alpha unless it is zero everywhere', () => {
    const withAlpha = decodeBmp(bmp({ width: 2, height: 1, bpp: 32, data: Uint8Array.from([1, 2, 3, 128, 4, 5, 6, 0]) }));
    expect(withAlpha.rgba).toEqual(
      rgba([
        [3, 2, 1, 128],
        [6, 5, 4, 0],
      ]),
    );
    const noAlpha = decodeBmp(bmp({ width: 2, height: 1, bpp: 32, data: Uint8Array.from([1, 2, 3, 0, 4, 5, 6, 0]) }));
    expect(noAlpha.rgba).toEqual(
      rgba([
        [3, 2, 1, 255],
        [6, 5, 4, 255],
      ]),
    );
  });

  it('decodes a V5 header with BI_BITFIELDS and an alpha mask', () => {
    const argb = (a: number, r: number, g: number, b: number) => le32(((a << 24) | (r << 16) | (g << 8) | b) >>> 0);
    const data = Uint8Array.from([...argb(255, 10, 20, 30), ...argb(0, 1, 2, 3), ...argb(77, 200, 100, 50), ...argb(128, 0, 0, 0)]);
    const img = decodeBmp(bmp({ width: 2, height: 2, bpp: 32, compression: 3, headerSize: 124, masks: [0x00ff0000, 0x0000ff00, 0x000000ff, 0xff000000], data }));
    expect(img.rgba).toEqual(
      rgba([
        [200, 100, 50, 77],
        [0, 0, 0, 128],
        [10, 20, 30, 255],
        [1, 2, 3, 0],
      ]),
    );
  });

  it('decodes RLE8 with runs, absolute runs, deltas and an early end of bitmap', () => {
    const palette: Rgb[] = [
      [0, 0, 0],
      [255, 0, 0],
      [0, 255, 0],
      [0, 0, 255],
    ];
    const data = Uint8Array.from([
      // Row 0 (bottom): run of 3 x index 1, absolute run [2, 3, 1] (odd, so padded), end of line.
      3, 1, 0, 3, 2, 3, 1, 0, 0, 0,
      // Row 1: run of 2 x index 2, end of line (the rest is skipped).
      2, 2, 0, 0,
      // Row 2: delta right 3 and up 1, landing in row 3 at x = 3; run of 2 x index 3; end of bitmap.
      0, 2, 3, 1, 2, 3, 0, 1,
    ]);
    const img = decodeBmp(bmp({ width: 6, height: 4, bpp: 8, compression: 1, palette, data }));
    const T = -1;
    // Top-down: row 3, row 2, row 1, row 0.
    const expected = [T, T, T, 3, 3, T, T, T, T, T, T, T, 2, 2, T, T, T, T, 1, 1, 1, 2, 3, 1];
    expect(img.rgba).toEqual(paletteRgba(palette, expected, T));
  });

  it('decodes RLE4 with alternating nibbles and padded absolute runs', () => {
    const palette: Rgb[] = [
      [0, 0, 0],
      [255, 0, 0],
      [0, 255, 0],
      [0, 0, 255],
    ];
    const data = Uint8Array.from([
      // Row 0: 5 pixels from nibbles 1, 2 alternating; end of line.
      5, 0x12, 0, 0,
      // Row 1: absolute run of 5 nibbles [3, 1, 2, 3, 1] in 3 bytes plus 1 byte of padding; end of bitmap.
      0, 5, 0x31, 0x23, 0x10, 0, 0, 1,
    ]);
    const img = decodeBmp(bmp({ width: 5, height: 2, bpp: 4, compression: 2, palette, data }));
    expect(img.rgba).toEqual(paletteRgba(palette, [3, 1, 2, 3, 1, 1, 2, 1, 2, 1]));
  });

  it('decodes a BITMAPCOREHEADER file with a 3-byte palette', () => {
    const palette: Rgb[] = [
      [9, 8, 7],
      [1, 2, 3],
    ];
    const data = padRows([
      [1, 0],
      [0, 1],
    ]);
    const img = decodeBmp(bmp({ width: 2, height: 2, bpp: 8, headerSize: 12, palette: [...palette, ...new Array<Rgb>(254).fill([0, 0, 0])], data }));
    expect(img.rgba).toEqual(paletteRgba(palette, [0, 1, 1, 0]));
  });

  it('refuses embedded JPEG', () => {
    expect(() => decodeBmp(bmp({ width: 1, height: 1, bpp: 24, compression: 4, data: new Uint8Array(4) }))).toThrow(/JPEG/);
  });
});

// ---------------------------------------------------------------------------------------------
// TIFF

type Entry = [tag: number, type: number, values: number[]];
const SHORT = 3;
const LONG = 4;

/** A TIFF with the given tags and strips (or tiles); offsets and byte counts are filled in. */
function tiff(le: boolean, entries: Entry[], chunks: Uint8Array[], tiled = false): Uint8Array {
  let size = 8;
  const offsets: number[] = [];
  for (const c of chunks) {
    offsets.push(size);
    size += c.length;
  }
  if (size & 1) size++;
  const located: Entry[] = [
    [tiled ? 324 : 273, LONG, offsets],
    [tiled ? 325 : 279, LONG, chunks.map((c) => c.length)],
  ];
  const all = [...entries, ...located].sort((a, b) => a[0] - b[0]);
  const ifd = size;
  let extra = ifd + 2 + all.length * 12 + 4;
  const width = (type: number) => (type === SHORT ? 2 : type === LONG ? 4 : 1);
  let total = extra;
  for (const [, type, values] of all) {
    const n = width(type) * values.length;
    if (n > 4) total += n + (n & 1);
  }
  const out = new Uint8Array(total);
  const v = new DataView(out.buffer);
  out.set(le ? [0x49, 0x49] : [0x4d, 0x4d]);
  v.setUint16(2, 42, le);
  v.setUint32(4, ifd, le);
  chunks.forEach((c, i) => out.set(c, offsets[i]!));
  v.setUint16(ifd, all.length, le);
  all.forEach(([tag, type, values], i) => {
    const e = ifd + 2 + i * 12;
    v.setUint16(e, tag, le);
    v.setUint16(e + 2, type, le);
    v.setUint32(e + 4, values.length, le);
    const n = width(type) * values.length;
    let at = e + 8;
    if (n > 4) {
      v.setUint32(e + 8, extra, le);
      at = extra;
      extra += n + (n & 1);
    }
    values.forEach((x, k) => {
      if (type === SHORT) v.setUint16(at + k * 2, x, le);
      else if (type === LONG) v.setUint32(at + k * 4, x, le);
      else v.setUint8(at + k, x);
    });
  });
  return out;
}

/** Pack sample values row by row (each row padded to a byte), optionally with horizontal differencing. */
function packSamples(values: readonly number[], width: number, rows: number, n: number, bps: number, le: boolean, predictor = false): Uint8Array {
  const max = bps === 16 ? 0xffff : (1 << bps) - 1;
  const rowBytes = Math.ceil((width * n * bps) / 8);
  const out = new Uint8Array(rowBytes * rows);
  for (let r = 0; r < rows; r++) {
    for (let k = 0; k < width * n; k++) {
      let v = values[r * width * n + k]!;
      if (predictor && k >= n) v = (v - values[r * width * n + k - n]!) & max;
      if (bps === 16) {
        const q = r * rowBytes + k * 2;
        out[q] = le ? v & 255 : v >> 8;
        out[q + 1] = le ? v >> 8 : v & 255;
      } else if (bps === 8) out[r * rowBytes + k] = v;
      else {
        const bit = k * bps;
        out[r * rowBytes + (bit >> 3)]! |= v << (8 - bps - (bit & 7));
      }
    }
  }
  return out;
}

/** PackBits: repeats of 2+ bytes as runs, everything else as literals. */
function packBits(data: Uint8Array): Uint8Array {
  const out: number[] = [];
  let i = 0;
  while (i < data.length) {
    let run = 1;
    while (i + run < data.length && run < 128 && data[i + run] === data[i]) run++;
    if (run >= 2) {
      out.push((1 - run) & 255, data[i]!);
      i += run;
      continue;
    }
    const start = i;
    while (i < data.length && i - start < 128 && !(i + 1 < data.length && data[i + 1] === data[i])) i++;
    out.push(i - start - 1, ...data.subarray(start, i));
  }
  return Uint8Array.from(out);
}

/** TIFF LZW: MSB-first codes with early change, clearing before the table overflows (as libtiff does). */
function tiffLzw(data: Uint8Array): Uint8Array {
  const out: number[] = [];
  let acc = 0;
  let bits = 0;
  let codeSize = 9;
  let next = 258;
  let dict = new Map<number, number>();
  const emit = (code: number) => {
    acc = (acc << codeSize) | code;
    bits += codeSize;
    while (bits >= 8) {
      bits -= 8;
      out.push((acc >>> bits) & 255);
    }
    acc &= (1 << bits) - 1;
  };
  emit(256);
  let w = data[0]!;
  for (let i = 1; i < data.length; i++) {
    const k = data[i]!;
    const key = w * 256 + k;
    const hit = dict.get(key);
    if (hit !== undefined) {
      w = hit;
      continue;
    }
    emit(w);
    dict.set(key, next++);
    if (next === 4094) {
      emit(256);
      dict = new Map();
      codeSize = 9;
      next = 258;
    } else if (next === 1 << codeSize) codeSize++;
    w = k;
  }
  emit(w);
  if (next + 1 === 1 << codeSize) codeSize++;
  emit(257);
  if (bits > 0) out.push((acc << (8 - bits)) & 255);
  return Uint8Array.from(out);
}

const size = (w: number, h: number): Entry[] => [
  [256, LONG, [w]],
  [257, LONG, [h]],
];

/** Expected RGBA from per-pixel RGB(A) values in 8 bits. */
function rgbRgba(values: readonly number[], n: number): Uint8Array {
  const out = new Uint8Array((values.length / n) * 4);
  for (let i = 0; i < values.length / n; i++) {
    out[i * 4] = values[i * n]!;
    out[i * 4 + 1] = values[i * n + (n >= 3 ? 1 : 0)]!;
    out[i * 4 + 2] = values[i * n + (n >= 3 ? 2 : 0)]!;
    out[i * 4 + 3] = n === 4 || n === 2 ? values[i * n + n - 1]! : 255;
  }
  return out;
}

describe('TIFF decoder', () => {
  it('decodes little-endian uncompressed RGB in several strips', async () => {
    const values = Array.from({ length: 3 * 3 * 3 }, (_, i) => (i * 37) & 255);
    const strips = [packSamples(values.slice(0, 18), 3, 2, 3, 8, true), packSamples(values.slice(18), 3, 1, 3, 8, true)];
    const bytes = tiff(true, [...size(3, 3), [258, SHORT, [8, 8, 8]], [259, SHORT, [1]], [262, SHORT, [2]], [277, SHORT, [3]], [278, LONG, [2]]], strips);
    const img = await decodeTiff(bytes);
    expect([img.width, img.height]).toEqual([3, 3]);
    expect(img.rgba).toEqual(rgbRgba(values, 3));
  });

  it('decodes big-endian PackBits greyscale', async () => {
    const values = [0, 0, 0, 0, 9, 1, 2, 3, 3, 3, 200, 200, 200, 200, 200, 7, 8, 9, 10, 11];
    const bytes = tiff(false, [...size(10, 2), [258, SHORT, [8]], [259, SHORT, [32773]], [262, SHORT, [1]]], [packBits(Uint8Array.from(values))]);
    expect((await decodeTiff(bytes)).rgba).toEqual(rgbRgba(values, 1));
  });

  it('decodes LZW with the horizontal predictor and unassociated alpha', async () => {
    const width = 40;
    const height = 30;
    const next = random(7);
    const values: number[] = [];
    for (let y = 0; y < height; y++) for (let x = 0; x < width; x++) values.push(x * 6, y * 8, (x + y) & 255 ? next() : 0, x < 20 ? 255 : x * 4);
    const strip = tiffLzw(packSamples(values, width, height, 4, 8, true, true));
    const entries: Entry[] = [...size(width, height), [258, SHORT, [8, 8, 8, 8]], [259, SHORT, [5]], [262, SHORT, [2]], [277, SHORT, [4]], [317, SHORT, [2]], [338, SHORT, [2]]];
    expect((await decodeTiff(tiff(true, entries, [strip]))).rgba).toEqual(rgbRgba(values, 4));
  });

  it('decodes LZW streams long enough to clear the table', async () => {
    const width = 100;
    const height = 100;
    const next = random(3);
    const values = Array.from({ length: width * height * 3 }, () => next());
    const strip = tiffLzw(Uint8Array.from(values));
    const bytes = tiff(false, [...size(width, height), [258, SHORT, [8, 8, 8]], [259, SHORT, [5]], [262, SHORT, [2]], [277, SHORT, [3]]], [strip]);
    expect((await decodeTiff(bytes)).rgba).toEqual(rgbRgba(values, 3));
  });

  it('decodes 16-bit RGB with LZW and the predictor, scaling to 8 bits', async () => {
    const values = [0, 65535, 0x1234, 0x8000, 0x7fff, 0xff00, 257 * 10, 257 * 20, 257 * 30, 65535, 0, 1];
    const strip = tiffLzw(packSamples(values, 2, 2, 3, 16, true, true));
    const bytes = tiff(true, [...size(2, 2), [258, SHORT, [16, 16, 16]], [259, SHORT, [5]], [262, SHORT, [2]], [277, SHORT, [3]], [317, SHORT, [2]]], [strip]);
    expect((await decodeTiff(bytes)).rgba).toEqual(rgbRgba(values.map((v) => Math.round((v * 255) / 65535)), 3));
  });

  it('decodes 16-bit big-endian greyscale', async () => {
    const values = [0, 0x0100, 0x7f7f, 0xffff];
    const bytes = tiff(false, [...size(2, 2), [258, SHORT, [16]], [262, SHORT, [1]]], [packSamples(values, 2, 2, 1, 16, false)]);
    expect((await decodeTiff(bytes)).rgba).toEqual(rgbRgba(values.map((v) => Math.round((v * 255) / 65535)), 1));
  });

  it('decodes a 4-bit palette image with a 16-bit colour map', async () => {
    const palette: Rgb[] = Array.from({ length: 16 }, (_, i) => [i * 17, 255 - i * 17, (i * 40) & 255]);
    const map = [0, 1, 2].flatMap((c) => palette.map((p) => p[c]! * 257));
    const indices = [0, 15, 3, 7, 9, 1, 2];
    const bytes = tiff(true, [...size(7, 1), [258, SHORT, [4]], [262, SHORT, [3]], [320, SHORT, map]], [packSamples(indices, 7, 1, 1, 4, true)]);
    expect((await decodeTiff(bytes)).rgba).toEqual(paletteRgba(palette, indices));
  });

  it('decodes 1-bit WhiteIsZero and 2-bit BlackIsZero images', async () => {
    const bits = [1, 0, 0, 1, 1, 1, 0, 0, 0, 1, 0, 1, 1, 0, 1, 0, 0, 0, 1, 1];
    const white = await decodeTiff(tiff(true, [...size(10, 2), [258, SHORT, [1]], [262, SHORT, [0]]], [packSamples(bits, 10, 2, 1, 1, true)]));
    expect(white.rgba).toEqual(rgbRgba(bits.map((b) => (b ? 0 : 255)), 1));

    const levels = [0, 1, 2, 3, 3, 2, 1];
    const grey = await decodeTiff(tiff(false, [...size(7, 1), [258, SHORT, [2]], [262, SHORT, [1]]], [packSamples(levels, 7, 1, 1, 2, false)]));
    expect(grey.rgba).toEqual(rgbRgba(levels.map((v) => v * 85), 1));
  });

  it('reverses bits for FillOrder 2', async () => {
    const bits = [1, 1, 0, 0, 0, 0, 0, 1, 0, 1];
    const packed = packSamples(bits, 10, 1, 1, 1, true).map((b) => Number.parseInt(b.toString(2).padStart(8, '0').split('').reverse().join(''), 2));
    const img = await decodeTiff(tiff(true, [...size(10, 1), [258, SHORT, [1]], [262, SHORT, [1]], [266, SHORT, [2]]], [packed]));
    expect(img.rgba).toEqual(rgbRgba(bits.map((b) => b * 255), 1));
  });

  it('decodes Deflate-compressed tiles, cropping the partial ones', async () => {
    const width = 20;
    const height = 18;
    const values: number[] = [];
    for (let y = 0; y < height; y++) for (let x = 0; x < width; x++) values.push(x * 12, y * 14, (x * y) & 255);
    const tiles: Uint8Array[] = [];
    for (let ty = 0; ty < 2; ty++) {
      for (let tx = 0; tx < 2; tx++) {
        const tile: number[] = [];
        for (let y = 0; y < 16; y++) {
          for (let x = 0; x < 16; x++) {
            const gx = tx * 16 + x;
            const gy = ty * 16 + y;
            if (gx < width && gy < height) tile.push(...values.slice((gy * width + gx) * 3, (gy * width + gx) * 3 + 3));
            else tile.push(99, 99, 99);
          }
        }
        tiles.push(new Uint8Array(deflateSync(packSamples(tile, 16, 16, 3, 8, true, true))));
      }
    }
    const entries: Entry[] = [...size(width, height), [258, SHORT, [8, 8, 8]], [259, SHORT, [8]], [262, SHORT, [2]], [277, SHORT, [3]], [317, SHORT, [2]], [322, SHORT, [16]], [323, SHORT, [16]]];
    expect((await decodeTiff(tiff(true, entries, tiles, true))).rgba).toEqual(rgbRgba(values, 3));
  });

  it('un-premultiplies associated alpha', async () => {
    // Grey + associated alpha: 64 under 128 is 50% grey at half opacity.
    const values = [64, 128, 255, 255, 0, 0, 10, 20];
    const bytes = tiff(true, [...size(4, 1), [258, SHORT, [8, 8]], [262, SHORT, [1]], [277, SHORT, [2]], [338, SHORT, [1]]], [packSamples(values, 4, 1, 2, 8, true)]);
    expect((await decodeTiff(bytes)).rgba).toEqual(
      rgba([
        [128, 128, 128, 128],
        [255, 255, 255, 255],
        [0, 0, 0, 0],
        [128, 128, 128, 20],
      ]),
    );
  });

  it('decodes planar (separate) sample planes', async () => {
    const width = 3;
    const height = 2;
    const values = Array.from({ length: width * height * 3 }, (_, i) => (i * 29) & 255);
    const plane = (c: number) => Uint8Array.from(values.filter((_, i) => i % 3 === c));
    const bytes = tiff(false, [...size(width, height), [258, SHORT, [8, 8, 8]], [262, SHORT, [2]], [277, SHORT, [3]], [284, SHORT, [2]]], [plane(0), plane(1), plane(2)]);
    expect((await decodeTiff(bytes)).rgba).toEqual(rgbRgba(values, 3));
  });

  it('refuses JPEG, CCITT and CMYK images with a clear error', async () => {
    const base = (extra: Entry[]) => tiff(true, [...size(1, 1), [258, SHORT, [8]], ...extra], [new Uint8Array(4)]);
    await expect(decodeTiff(base([[259, SHORT, [7]], [262, SHORT, [2]]]))).rejects.toThrow(/JPEG/);
    await expect(decodeTiff(base([[259, SHORT, [4]], [262, SHORT, [0]]]))).rejects.toThrow(/CCITT/);
    await expect(decodeTiff(tiff(true, [...size(1, 1), [258, SHORT, [8, 8, 8, 8]], [262, SHORT, [5]], [277, SHORT, [4]]], [new Uint8Array(4)]))).rejects.toThrow(/CMYK/);
  });
});

// ---------------------------------------------------------------------------------------------
// Sniffing, dispatch and PDF embedding

const tinyGif = () =>
  gif({
    width: 2,
    height: 2,
    indices: [0, 1, 2, 3],
    globalPalette: RGBW,
    transparent: 3,
  });
const tinyBmp = () => bmp({ width: 1, height: 1, bpp: 24, data: padRows([[1, 2, 3]]) });
const tinyTiff = () => tiff(false, [...size(1, 1), [258, SHORT, [8]], [262, SHORT, [1]]], [Uint8Array.from([77])]);

describe('sniffImage, browserShows and decodeRaster', () => {
  it('identifies every format by its bytes', () => {
    const emf = new Uint8Array(88);
    emf.set([1, 0, 0, 0]);
    emf.set(ascii(' EMF'), 40);
    const enc = (s: string) => new TextEncoder().encode(s);
    expect(sniffImage(Uint8Array.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]))).toBe('png');
    expect(sniffImage(Uint8Array.from([0xff, 0xd8, 0xff, 0xe0]))).toBe('jpeg');
    expect(sniffImage(tinyGif())).toBe('gif');
    expect(sniffImage(tinyBmp())).toBe('bmp');
    expect(sniffImage(tinyTiff())).toBe('tiff');
    expect(sniffImage(tiff(true, size(1, 1), [new Uint8Array(1)]))).toBe('tiff');
    expect(sniffImage(emf)).toBe('emf');
    expect(sniffImage(Uint8Array.from([0xd7, 0xcd, 0xc6, 0x9a, 0, 0]))).toBe('wmf');
    expect(sniffImage(Uint8Array.from([1, 0, 9, 0, 0, 3, 0, 0]))).toBe('wmf');
    expect(sniffImage(enc('﻿<?xml version="1.0"?>\n<!-- x -->\n<svg xmlns="http://www.w3.org/2000/svg"/>'))).toBe('svg');
    expect(sniffImage(enc('BM is not a bitmap header'))).toBeUndefined();
    expect(sniffImage(enc('<html><body/></html>'))).toBeUndefined();
    expect(sniffImage(new Uint8Array())).toBeUndefined();
  });

  it('knows which formats browsers show as they are', () => {
    for (const kind of ['png', 'jpeg', 'gif', 'bmp', 'svg'] as const) expect(browserShows(kind)).toBe(true);
    for (const kind of ['tiff', 'emf', 'wmf', undefined] as const) expect(browserShows(kind)).toBe(false);
  });

  it('dispatches to the right decoder', async () => {
    expect((await decodeRaster(tinyGif())).rgba).toEqual(paletteRgba(RGBW, [0, 1, 2, 3], 3));
    expect((await decodeRaster(tinyBmp())).rgba).toEqual(rgba([[3, 2, 1, 255]]));
    expect((await decodeRaster(tinyTiff())).rgba).toEqual(rgba([[77, 77, 77, 255]]));
    await expect(decodeRaster(Uint8Array.from([0x89, 0x50, 0x4e, 0x47]))).rejects.toThrow(/PNG/);
  });
});

describe('addImage with decoded formats', () => {
  const pdfText = (writer: PdfWriter, ref: Parameters<PdfWriter['build']>[0]) => new TextDecoder('latin1').decode(writer.build(ref, undefined, new Uint8Array(16)));

  it('embeds a transparent GIF with a soft mask', async () => {
    const writer = new PdfWriter();
    const { ref, width, height } = await addImage(writer, tinyGif());
    expect([width, height]).toEqual([2, 2]);
    const text = pdfText(writer, ref);
    expect(text).toContain('/SMask');
    expect(text).toContain('/DeviceRGB');
  });

  it('embeds BMP and TIFF images', async () => {
    const writer = new PdfWriter();
    expect(await addImage(writer, tinyBmp())).toMatchObject({ width: 1, height: 1 });
    expect(await addImage(writer, tinyTiff())).toMatchObject({ width: 1, height: 1 });
  });

  it('still rejects formats it cannot embed, and damaged files, with ImageFormatError', async () => {
    const emf = new Uint8Array(88);
    emf.set([1, 0, 0, 0]);
    emf.set(ascii(' EMF'), 40);
    await expect(addImage(new PdfWriter(), emf)).rejects.toThrow(ImageFormatError);
    const broken = tiff(true, [...size(1, 1), [258, SHORT, [8]], [259, SHORT, [7]]], [new Uint8Array(4)]);
    await expect(addImage(new PdfWriter(), broken)).rejects.toThrow(ImageFormatError);
    await expect(addImage(new PdfWriter(), broken)).rejects.toThrow(/JPEG/);
  });
});
