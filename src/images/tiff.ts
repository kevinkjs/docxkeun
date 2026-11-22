/**
 * Baseline TIFF decoder producing RGBA pixels from the first image (IFD) of a file. Browsers other
 * than Safari cannot show TIFF and PDF has no TIFF filter, so the viewer and the PDF exporter both
 * need the pixels.
 *
 * Supported: either byte order; strips or tiles; no compression, PackBits, LZW and Deflate, with
 * horizontal-differencing prediction; WhiteIsZero, BlackIsZero, RGB and Palette images at 1, 2, 4,
 * 8 or 16 bits per sample; associated or unassociated alpha; chunky or planar sample layout.
 * JPEG, CCITT fax, CMYK, YCbCr and Lab images are refused with an error that says so.
 */
import type { RasterImage } from '../layout/types.js';
import { inflate } from '../zip/deflate.js';

/** Refuse headers that would need gigabytes of RGBA before a single pixel is read. */
const MAX_PIXELS = 1 << 27;

const TAG = {
  width: 256,
  height: 257,
  bitsPerSample: 258,
  compression: 259,
  photometric: 262,
  fillOrder: 266,
  stripOffsets: 273,
  samplesPerPixel: 277,
  rowsPerStrip: 278,
  stripByteCounts: 279,
  planarConfiguration: 284,
  predictor: 317,
  colorMap: 320,
  tileWidth: 322,
  tileLength: 323,
  tileOffsets: 324,
  tileByteCounts: 325,
  extraSamples: 338,
  sampleFormat: 339,
} as const;

/** Bytes per value for each TIFF field type. */
const TYPE_SIZE: Record<number, number> = { 1: 1, 2: 1, 3: 2, 4: 4, 5: 8, 6: 1, 7: 1, 8: 2, 9: 4, 10: 8, 11: 4, 12: 8 };

export function isTiff(b: Uint8Array): boolean {
  // 42 is classic TIFF, 43 BigTIFF (recognised so that it gets a clear error rather than none).
  return b.length >= 4 && ((b[0] === 0x49 && b[1] === 0x49 && (b[2] === 42 || b[2] === 43) && b[3] === 0) || (b[0] === 0x4d && b[1] === 0x4d && b[2] === 0 && (b[3] === 42 || b[3] === 43)));
}

function readIfd(view: DataView, offset: number, le: boolean): Map<number, number[]> {
  const tags = new Map<number, number[]>();
  if (offset + 2 > view.byteLength) throw new Error('TIFF: first IFD lies outside the file');
  const count = view.getUint16(offset, le);
  for (let i = 0; i < count; i++) {
    const e = offset + 2 + i * 12;
    if (e + 12 > view.byteLength) break;
    const tag = view.getUint16(e, le);
    const type = view.getUint16(e + 2, le);
    const n = view.getUint32(e + 4, le);
    const size = TYPE_SIZE[type];
    if (!size) continue;
    const total = size * n;
    const at = total <= 4 ? e + 8 : view.getUint32(e + 8, le);
    if (at + total > view.byteLength) continue;
    const values: number[] = new Array<number>(n);
    for (let k = 0; k < n; k++) {
      const q = at + k * size;
      switch (type) {
        case 3:
          values[k] = view.getUint16(q, le);
          break;
        case 4:
          values[k] = view.getUint32(q, le);
          break;
        case 5:
          values[k] = view.getUint32(q, le) / view.getUint32(q + 4, le);
          break;
        case 6:
          values[k] = view.getInt8(q);
          break;
        case 8:
          values[k] = view.getInt16(q, le);
          break;
        case 9:
          values[k] = view.getInt32(q, le);
          break;
        case 10:
          values[k] = view.getInt32(q, le) / view.getInt32(q + 4, le);
          break;
        case 11:
          values[k] = view.getFloat32(q, le);
          break;
        case 12:
          values[k] = view.getFloat64(q, le);
          break;
        default:
          values[k] = view.getUint8(q);
      }
    }
    tags.set(tag, values);
  }
  return tags;
}

/** PackBits: a signed count byte, then n+1 literal bytes or one byte repeated 1-n times. */
function unpackBits(src: Uint8Array, size: number): Uint8Array {
  const out = new Uint8Array(size);
  let p = 0;
  let o = 0;
  while (p < src.length && o < size) {
    const n = (src[p++]! << 24) >> 24;
    if (n >= 0) {
      const len = n + 1;
      out.set(src.subarray(p, Math.min(p + len, src.length, p + size - o)), o);
      p += len;
      o += len;
    } else if (n !== -128) {
      if (p >= src.length) break;
      const v = src[p++]!;
      out.fill(v, o, Math.min(o + 1 - n, size));
      o += 1 - n;
    }
  }
  return out;
}

/**
 * TIFF's LZW differs from GIF's: codes are packed MSB-first, the literal alphabet is always 8 bits,
 * and the code width grows one code early ("early change"), as libtiff and every writer do.
 */
function unpackLzw(src: Uint8Array, size: number): Uint8Array {
  // Pre-6.0 LZW was LSB-first; its stream starts with the clear code written that way.
  if (src.length >= 2 && src[0] === 0 && (src[1]! & 1) !== 0) throw new Error('TIFF: old-style (pre-6.0) LZW is not supported');
  const out = new Uint8Array(size);
  const prefix = new Uint16Array(4096);
  const suffix = new Uint8Array(4096);
  const first = new Uint8Array(4096);
  const length = new Uint16Array(4096);
  for (let i = 0; i < 256; i++) {
    suffix[i] = i;
    first[i] = i;
    length[i] = 1;
  }
  let codeSize = 9;
  let next = 258;
  let prev = -1;
  let acc = 0;
  let bits = 0;
  let pos = 0;
  let n = 0;
  while (n < size) {
    while (bits < codeSize) {
      if (pos >= src.length) return out;
      acc = (acc << 8) | src[pos++]!;
      bits += 8;
    }
    bits -= codeSize;
    const code = (acc >>> bits) & ((1 << codeSize) - 1);
    acc &= (1 << bits) - 1;
    if (code === 256) {
      codeSize = 9;
      next = 258;
      prev = -1;
      continue;
    }
    if (code === 257) break;
    if (prev < 0) {
      if (code > 255) break;
      out[n++] = code;
      prev = code;
      continue;
    }
    if (code > next) break;
    if (next < 4096) {
      prefix[next] = prev;
      first[next] = first[prev]!;
      suffix[next] = code === next ? first[prev]! : first[code]!;
      length[next] = length[prev]! + 1;
      next++;
      if (next >= (1 << codeSize) - 1 && codeSize < 12) codeSize++;
    }
    const len = length[code]!;
    let c = code;
    for (let i = n + len - 1; i >= n; i--) {
      if (i < size) out[i] = suffix[c]!;
      c = prefix[c]!;
    }
    n += len;
    prev = code;
  }
  return out;
}

/** FillOrder 2 stores the first pixel in the least significant bit. */
function reverseBits(src: Uint8Array): Uint8Array {
  const out = new Uint8Array(src.length);
  for (let i = 0; i < src.length; i++) {
    let v = src[i]!;
    v = ((v & 0xf0) >> 4) | ((v & 0x0f) << 4);
    v = ((v & 0xcc) >> 2) | ((v & 0x33) << 2);
    out[i] = ((v & 0xaa) >> 1) | ((v & 0x55) << 1);
  }
  return out;
}

/** Decompress one strip or tile; the result is at least `size` bytes so unpacking needs no bounds checks. */
async function decompress(raw: Uint8Array, compression: number, size: number): Promise<Uint8Array> {
  if (compression === 5) return unpackLzw(raw, size);
  if (compression === 32773) return unpackBits(raw, size);
  const data = compression === 8 || compression === 32946 ? await inflate(raw, 'deflate') : raw;
  if (data.length >= size) return data;
  const padded = new Uint8Array(size);
  padded.set(data);
  return padded;
}

function unsupportedCompression(compression: number): string | undefined {
  if (compression === 1 || compression === 5 || compression === 8 || compression === 32773 || compression === 32946) return undefined;
  if (compression === 2 || compression === 3 || compression === 4 || compression === 32771) return 'CCITT fax-compressed TIFF is not supported';
  if (compression === 6 || compression === 7) return 'JPEG-compressed TIFF is not supported';
  return `TIFF compression ${compression} is not supported`;
}

function unsupportedPhotometric(photometric: number): string | undefined {
  if (photometric <= 3) return undefined;
  const what: Record<number, string> = { 4: 'transparency-mask', 5: 'CMYK (separated)', 6: 'YCbCr', 8: 'CIE Lab', 9: 'ICC Lab', 10: 'ITU Lab', 32844: 'LogL', 32845: 'LogLuv' };
  return `${what[photometric] ?? `photometric ${photometric}`} TIFF is not supported`;
}

interface Layout {
  width: number;
  height: number;
  spp: number;
  bps: number;
  le: boolean;
  predictor: number;
}

/**
 * Spread one strip or tile into the whole-image sample array. `n` is the number of samples per
 * pixel stored in the chunk (all of them when chunky, one when planar) and `plane` the first one.
 */
function unpack(data: Uint8Array, samples: Uint8Array | Uint16Array, l: Layout, x0: number, y0: number, chunkWidth: number, rows: number, n: number, plane: number): void {
  const { width, height, spp, bps, le, predictor } = l;
  const max = bps === 16 ? 0xffff : (1 << bps) - 1;
  const rowBytes = Math.ceil((chunkWidth * n * bps) / 8);
  // Prediction runs left to right, so the columns past the image edge never matter.
  const visible = Math.min(chunkWidth, width - x0);
  const prev = new Uint32Array(n);
  for (let r = 0; r < rows; r++) {
    const y = y0 + r;
    if (y >= height) break;
    const start = r * rowBytes;
    let out = (y * width + x0) * spp + plane;
    for (let x = 0; x < visible; x++, out += spp) {
      for (let s = 0; s < n; s++) {
        const k = x * n + s;
        let v: number;
        if (bps === 8) v = data[start + k]!;
        else if (bps === 16) {
          const q = start + k * 2;
          v = le ? data[q]! | (data[q + 1]! << 8) : (data[q]! << 8) | data[q + 1]!;
        } else {
          const bit = k * bps;
          v = (data[start + (bit >> 3)]! >> (8 - bps - (bit & 7))) & max;
        }
        if (predictor === 2 && x > 0) v = (v + prev[s]!) & max;
        prev[s] = v;
        samples[out + s] = v;
      }
    }
  }
}

/** Decode the first image of a TIFF file. */
export async function decodeTiff(bytes: Uint8Array): Promise<RasterImage> {
  if (!isTiff(bytes) || bytes.length < 8) throw new Error('TIFF: missing II*/MM* header');
  const le = bytes[0] === 0x49;
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  if (view.getUint16(2, le) === 43) throw new Error('TIFF: BigTIFF files are not supported');
  const tags = readIfd(view, view.getUint32(4, le), le);
  const one = (tag: number, fallback: number) => tags.get(tag)?.[0] ?? fallback;

  const width = one(TAG.width, 0);
  const height = one(TAG.height, 0);
  if (!width || !height) throw new Error('TIFF: missing image width or length');
  if (width * height > MAX_PIXELS) throw new Error(`TIFF: ${width}x${height} is too large to decode`);

  const compression = one(TAG.compression, 1);
  const badCompression = unsupportedCompression(compression);
  if (badCompression) throw new Error(`TIFF: ${badCompression}`);

  const spp = one(TAG.samplesPerPixel, 1);
  if (spp < 1 || spp > 16) throw new Error(`TIFF: ${spp} samples per pixel is not supported`);
  const bitsList = tags.get(TAG.bitsPerSample) ?? [1];
  const bps = bitsList[0] ?? 1;
  if (bitsList.slice(0, spp).some((b) => b !== bps)) throw new Error('TIFF: samples with different bit depths are not supported');
  if (bps !== 1 && bps !== 2 && bps !== 4 && bps !== 8 && bps !== 16) throw new Error(`TIFF: ${bps} bits per sample is not supported`);
  if ((tags.get(TAG.sampleFormat) ?? [1]).some((f) => f !== 1)) throw new Error('TIFF: only unsigned integer samples are supported');

  const photometric = one(TAG.photometric, spp >= 3 ? 2 : 1);
  const badPhotometric = unsupportedPhotometric(photometric);
  if (badPhotometric) throw new Error(`TIFF: ${badPhotometric}`);
  const colorSamples = photometric === 2 ? 3 : 1;
  if (spp < colorSamples) throw new Error(`TIFF: ${spp} samples per pixel is too few for the photometric interpretation`);
  // Extra samples marked 0 ("unspecified") are not alpha and are ignored.
  const extra = spp > colorSamples ? (tags.get(TAG.extraSamples)?.[0] ?? 0) : 0;
  const alpha = extra === 1 || extra === 2 ? colorSamples : -1;
  const associated = extra === 1;

  const planar = one(TAG.planarConfiguration, 1);
  if (planar !== 1 && planar !== 2) throw new Error(`TIFF: planar configuration ${planar} is not supported`);
  const predictor = one(TAG.predictor, 1);
  if (predictor === 3) throw new Error('TIFF: the floating-point predictor is not supported');
  if (predictor !== 1 && predictor !== 2) throw new Error(`TIFF: predictor ${predictor} is not supported`);
  const fillOrder = one(TAG.fillOrder, 1);

  const tiled = tags.has(TAG.tileWidth);
  const chunkWidth = tiled ? one(TAG.tileWidth, 0) : width;
  const chunkHeight = tiled ? one(TAG.tileLength, 0) : Math.min(one(TAG.rowsPerStrip, height), height) || height;
  if (!chunkWidth || !chunkHeight || chunkWidth * chunkHeight > MAX_PIXELS) throw new Error(`TIFF: invalid tile size ${chunkWidth}x${chunkHeight}`);
  const offsets = tags.get(tiled ? TAG.tileOffsets : TAG.stripOffsets) ?? [];
  const counts = tags.get(tiled ? TAG.tileByteCounts : TAG.stripByteCounts) ?? [];
  const across = Math.ceil(width / chunkWidth);
  const perPlane = across * Math.ceil(height / chunkHeight);
  const planes = planar === 2 ? spp : 1;
  const n = planar === 2 ? 1 : spp;
  if (offsets.length < perPlane * planes) throw new Error(`TIFF: expected ${perPlane * planes} ${tiled ? 'tile' : 'strip'} offsets, found ${offsets.length}`);
  const rowBytes = Math.ceil((chunkWidth * n * bps) / 8);

  const layout: Layout = { width, height, spp, bps, le, predictor };
  const samples = bps > 8 ? new Uint16Array(width * height * spp) : new Uint8Array(width * height * spp);
  for (let i = 0; i < perPlane * planes; i++) {
    const start = offsets[i]!;
    if (start >= bytes.length) continue;
    const t = i % perPlane;
    const x0 = (t % across) * chunkWidth;
    const y0 = Math.floor(t / across) * chunkHeight;
    // The last strip may be short; tiles are always stored whole.
    const rows = tiled ? chunkHeight : Math.min(chunkHeight, height - y0);
    const size = rows * rowBytes;
    // A missing or zero byte count happens in the wild; uncompressed data has a known size.
    const count = counts[i] || (compression === 1 ? size : bytes.length - start);
    let raw = bytes.subarray(start, Math.min(bytes.length, start + count));
    if (fillOrder === 2) raw = reverseBits(raw);
    const data = await decompress(raw, compression, size);
    unpack(data, samples, layout, x0, y0, chunkWidth, rows, n, planar === 2 ? Math.floor(i / perPlane) : 0);
  }
  return toRgba(samples, layout, photometric, alpha, associated, tags.get(TAG.colorMap));
}

function toRgba(samples: Uint8Array | Uint16Array, l: Layout, photometric: number, alpha: number, associated: boolean, colorMap: number[] | undefined): RasterImage {
  const { width, height, spp, bps } = l;
  const max = bps === 16 ? 0xffff : (1 << bps) - 1;
  const scale = new Uint8Array(max + 1);
  for (let v = 0; v <= max; v++) scale[v] = Math.round((v * 255) / max);

  let palette: Uint8Array | undefined;
  if (photometric === 3) {
    const entries = 1 << bps;
    if (!colorMap || colorMap.length < entries * 3) throw new Error('TIFF: palette image without a colour map');
    // The map should hold 16-bit values, but some writers store 8-bit ones; libtiff guesses the same way.
    const wide = colorMap.some((v) => v > 255);
    palette = new Uint8Array(entries * 3);
    for (let i = 0; i < entries; i++) {
      for (let c = 0; c < 3; c++) {
        const v = colorMap[c * entries + i]!;
        palette[i * 3 + c] = wide ? v >> 8 : v;
      }
    }
  }

  const pixels = width * height;
  const rgba = new Uint8Array(pixels * 4);
  // Premultiplied colour c under alpha a becomes c / a; both are in sample units, so the result is 8-bit directly.
  const unpremultiply = (c: number, a: number) => (a === 0 ? 0 : Math.min(255, Math.round((c * 255) / a)));
  for (let i = 0, s = 0, o = 0; i < pixels; i++, s += spp, o += 4) {
    const a = alpha >= 0 ? samples[s + alpha]! : max;
    if (palette) {
      const idx = samples[s]!;
      rgba[o] = palette[idx * 3]!;
      rgba[o + 1] = palette[idx * 3 + 1]!;
      rgba[o + 2] = palette[idx * 3 + 2]!;
    } else if (photometric === 2) {
      const r = samples[s]!;
      const g = samples[s + 1]!;
      const b = samples[s + 2]!;
      if (associated) {
        rgba[o] = unpremultiply(r, a);
        rgba[o + 1] = unpremultiply(g, a);
        rgba[o + 2] = unpremultiply(b, a);
      } else {
        rgba[o] = scale[r]!;
        rgba[o + 1] = scale[g]!;
        rgba[o + 2] = scale[b]!;
      }
    } else {
      let g = samples[s]!;
      if (photometric === 0) g = max - g;
      rgba[o] = rgba[o + 1] = rgba[o + 2] = associated ? unpremultiply(g, a) : scale[g]!;
    }
    rgba[o + 3] = scale[a]!;
  }
  return { width, height, rgba };
}
