/**
 * GIF decoder (GIF87a and GIF89a) producing RGBA pixels. PDF pages are static, so only the first
 * frame is decoded. It is composited onto a transparent canvas of the logical screen size; the
 * background colour index is ignored, as browsers ignore it.
 */
import type { RasterImage } from '../layout/types.js';
import { concatBytes } from '../util/bytes.js';

/** Refuse headers that would need gigabytes of RGBA before a single pixel is read. */
const MAX_PIXELS = 1 << 27;

export function isGif(b: Uint8Array): boolean {
  return b.length >= 6 && b[0] === 0x47 && b[1] === 0x49 && b[2] === 0x46 && b[3] === 0x38 && (b[4] === 0x37 || b[4] === 0x39) && b[5] === 0x61;
}

/**
 * GIF's variable-width LZW: codes are packed LSB-first, the width grows as soon as the next free
 * code no longer fits, and stays at 12 bits once the table is full until a clear code arrives
 * (encoders may defer the clear). Returns how many indices were written, so a truncated or damaged
 * stream still yields the pixels before the damage.
 */
function lzwDecode(data: Uint8Array, minCodeSize: number, out: Uint8Array): number {
  const clear = 1 << minCodeSize;
  const end = clear + 1;
  const prefix = new Uint16Array(4096);
  const suffix = new Uint8Array(4096);
  const first = new Uint8Array(4096);
  const length = new Uint16Array(4096);
  for (let i = 0; i < clear; i++) {
    suffix[i] = i;
    first[i] = i;
    length[i] = 1;
  }
  let codeSize = minCodeSize + 1;
  let next = clear + 2;
  let prev = -1;
  let acc = 0;
  let bits = 0;
  let pos = 0;
  let n = 0;
  while (n < out.length) {
    while (bits < codeSize) {
      if (pos >= data.length) return n;
      acc |= data[pos++]! << bits;
      bits += 8;
    }
    const code = acc & ((1 << codeSize) - 1);
    acc >>>= codeSize;
    bits -= codeSize;
    if (code === clear) {
      codeSize = minCodeSize + 1;
      next = clear + 2;
      prev = -1;
      continue;
    }
    if (code === end) break;
    if (prev < 0) {
      // The first code after a clear must be a literal; anything else means corrupt data.
      if (code > clear) break;
      out[n++] = code;
      prev = code;
      continue;
    }
    if (code > next) break;
    if (next < 4096) {
      // code === next is the "KwKwK" case: the string is prev's plus prev's first byte.
      prefix[next] = prev;
      first[next] = first[prev]!;
      suffix[next] = code === next ? first[prev]! : first[code]!;
      length[next] = length[prev]! + 1;
      next++;
      if (next === 1 << codeSize && codeSize < 12) codeSize++;
    }
    const len = length[code]!;
    let c = code;
    for (let i = n + len - 1; i >= n; i--) {
      if (i < out.length) out[i] = suffix[c]!;
      c = prefix[c]!;
    }
    n += len;
    prev = code;
  }
  return Math.min(n, out.length);
}

/** Row order of an interlaced image: every 8th row from 0, every 8th from 4, every 4th from 2, every 2nd from 1. */
function interlacedRows(height: number): Uint32Array {
  const rows = new Uint32Array(height);
  let i = 0;
  for (const [start, step] of [
    [0, 8],
    [4, 8],
    [2, 4],
    [1, 2],
  ] as const) {
    for (let y = start; y < height; y += step) rows[i++] = y;
  }
  return rows;
}

function skipSubBlocks(b: Uint8Array, p: number): number {
  while (p < b.length) {
    const n = b[p++]!;
    if (n === 0) break;
    p += n;
  }
  return p;
}

/** Decode the first frame of a GIF. */
export function decodeGif(bytes: Uint8Array): RasterImage {
  if (!isGif(bytes)) throw new Error('GIF: missing GIF87a/GIF89a signature');
  if (bytes.length < 13) throw new Error('GIF: truncated header');
  const u16 = (p: number) => bytes[p]! | (bytes[p + 1]! << 8);
  const screenWidth = u16(6);
  const screenHeight = u16(8);
  const flags = bytes[10]!;
  let p = 13;
  let globalTable: Uint8Array | undefined;
  if (flags & 0x80) {
    const size = 3 << ((flags & 7) + 1);
    globalTable = bytes.subarray(p, p + size);
    p += size;
  }

  let transparent = -1;
  while (p < bytes.length) {
    const block = bytes[p++]!;
    if (block === 0x2c) return decodeFrame(bytes, p, screenWidth, screenHeight, globalTable, transparent);
    if (block === 0x21) {
      const label = bytes[p++];
      if (label === 0xf9 && (bytes[p] ?? 0) >= 4) transparent = (bytes[p + 1] ?? 0) & 1 ? (bytes[p + 4] ?? -1) : -1;
      // A Graphic Control Extension applies to the next graphic block only, which may be plain text.
      else if (label === 0x01) transparent = -1;
      p = skipSubBlocks(bytes, p);
      continue;
    }
    if (block === 0x3b) break;
    throw new Error(`GIF: unexpected block 0x${block.toString(16)} at offset ${p - 1}`);
  }
  throw new Error('GIF: no image in file');
}

function decodeFrame(bytes: Uint8Array, p: number, screenWidth: number, screenHeight: number, globalTable: Uint8Array | undefined, transparent: number): RasterImage {
  if (p + 9 > bytes.length) throw new Error('GIF: truncated image descriptor');
  const u16 = (q: number) => bytes[q]! | (bytes[q + 1]! << 8);
  let left = u16(p);
  let top = u16(p + 2);
  const w = u16(p + 4);
  const h = u16(p + 6);
  const flags = bytes[p + 8]!;
  p += 9;
  let table = globalTable;
  if (flags & 0x80) {
    const size = 3 << ((flags & 7) + 1);
    table = bytes.subarray(p, p + size);
    p += size;
  }
  const interlaced = (flags & 0x40) !== 0;
  const minCodeSize = bytes[p++];
  if (minCodeSize === undefined || minCodeSize < 2 || minCodeSize > 8) throw new Error(`GIF: invalid LZW code size ${minCodeSize}`);
  const chunks: Uint8Array[] = [];
  while (p < bytes.length) {
    const n = bytes[p++]!;
    if (n === 0) break;
    chunks.push(bytes.subarray(p, p + n));
    p += n;
  }

  // Browsers grow a logical screen that is too small for the first frame (or zero) to fit it,
  // and place the frame at the origin; files written that way are common enough to follow suit.
  let width = screenWidth;
  let height = screenHeight;
  if (w > width || h > height) {
    width = Math.max(width, w);
    height = Math.max(height, h);
    left = 0;
    top = 0;
  }
  if (!width || !height) throw new Error('GIF: image has no size');
  if (width * height > MAX_PIXELS) throw new Error(`GIF: ${width}x${height} is too large to decode`);

  const rgba = new Uint8Array(width * height * 4);
  if (!w || !h) return { width, height, rgba };
  const indices = new Uint8Array(w * h);
  const count = lzwDecode(concatBytes(chunks), minCodeSize, indices);
  const rows = interlaced ? interlacedRows(h) : undefined;
  // Without any colour table the indices are shown as grey levels.
  const colors = table ? Math.floor(table.length / 3) : 256;
  for (let i = 0; i < count; i++) {
    const idx = indices[i]!;
    if (idx === transparent || idx >= colors) continue;
    const r = (i / w) | 0;
    const x = left + i - r * w;
    const y = top + (rows ? rows[r]! : r);
    if (x >= width || y >= height) continue;
    const o = (y * width + x) * 4;
    if (table) {
      rgba[o] = table[idx * 3]!;
      rgba[o + 1] = table[idx * 3 + 1]!;
      rgba[o + 2] = table[idx * 3 + 2]!;
    } else {
      rgba[o] = rgba[o + 1] = rgba[o + 2] = idx;
    }
    rgba[o + 3] = 255;
  }
  return { width, height, rgba };
}
