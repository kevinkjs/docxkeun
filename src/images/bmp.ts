/**
 * BMP / DIB decoder producing RGBA pixels. Reads BITMAPCOREHEADER (OS/2 1.x), BITMAPINFOHEADER
 * and its V2–V5 extensions, and the shorter OS/2 2.x headers; 1, 2, 4, 8, 16, 24 and 32 bits per
 * pixel; BI_RGB, BI_BITFIELDS, BI_ALPHABITFIELDS, BI_RLE8 and BI_RLE4.
 */
import type { RasterImage } from '../layout/types.js';

/** Refuse headers that would need gigabytes of RGBA before a single pixel is read. */
const MAX_PIXELS = 1 << 27;

const BI_RGB = 0;
const BI_RLE8 = 1;
const BI_RLE4 = 2;
const BI_BITFIELDS = 3;
const BI_ALPHABITFIELDS = 6;

export function isBmp(b: Uint8Array): boolean {
  if (b.length < 18 || b[0] !== 0x42 || b[1] !== 0x4d) return false;
  // "BM" alone is too weak a signature (text can start with it), so check the DIB header size too.
  const size = (b[14]! | (b[15]! << 8) | (b[16]! << 16) | (b[17]! << 24)) >>> 0;
  return size === 12 || (size >= 16 && size <= 64) || size === 108 || size === 124;
}

interface Channel {
  mask: number;
  shift: number;
  max: number;
}

function channel(mask: number): Channel {
  mask >>>= 0;
  if (!mask) return { mask: 0, shift: 0, max: 0 };
  let shift = 0;
  while (!((mask >>> shift) & 1)) shift++;
  return { mask, shift, max: mask >>> shift };
}

/** Scale a bitfield to 8 bits; a 5-bit 31 must become 255, not 248. */
function extract(px: number, c: Channel): number {
  if (!c.max) return 0;
  const v = (px & c.mask) >>> c.shift;
  return c.max === 255 ? v : Math.round((v * 255) / c.max);
}

/** Decode a .bmp file (with its 14-byte BITMAPFILEHEADER). */
export function decodeBmp(bytes: Uint8Array): RasterImage {
  if (bytes.length < 26 || bytes[0] !== 0x42 || bytes[1] !== 0x4d) throw new Error('BMP: missing "BM" signature');
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  const dataOffset = view.getUint32(10, true);
  const headerSize = view.getUint32(14, true);
  if (headerSize < 12 || 14 + headerSize > bytes.length) throw new Error('BMP: truncated header');

  let width: number;
  let height: number;
  let bpp: number;
  let compression = BI_RGB;
  let colorsUsed = 0;
  let topDown = false;
  let entrySize = 4;
  if (headerSize === 12) {
    width = view.getUint16(18, true);
    height = view.getUint16(20, true);
    bpp = view.getUint16(24, true);
    entrySize = 3;
  } else {
    width = view.getInt32(18, true);
    const h = view.getInt32(22, true);
    bpp = view.getUint16(28, true);
    // OS/2 2.x headers may stop early; missing fields are zero.
    if (headerSize >= 20) compression = view.getUint32(30, true);
    if (headerSize >= 36) colorsUsed = view.getUint32(46, true);
    topDown = h < 0;
    height = Math.abs(h);
  }
  // OS/2 2.x reuses compression 3 and 4 for Huffman 1D and RLE24, which Windows never wrote.
  const os2 = headerSize > 12 && headerSize <= 64 && headerSize !== 40 && headerSize !== 52 && headerSize !== 56;
  if (width <= 0 || height <= 0) throw new Error(`BMP: invalid size ${width}x${height}`);
  if (width * height > MAX_PIXELS) throw new Error(`BMP: ${width}x${height} is too large to decode`);
  if (![1, 2, 4, 8, 16, 24, 32].includes(bpp)) throw new Error(`BMP: ${bpp} bits per pixel is not supported`);

  const bitfields = !os2 && (compression === BI_BITFIELDS || compression === BI_ALPHABITFIELDS);
  if (compression === BI_RLE8 && bpp !== 8) throw new Error('BMP: RLE8 needs 8 bits per pixel');
  if (compression === BI_RLE4 && bpp !== 4) throw new Error('BMP: RLE4 needs 4 bits per pixel');
  if (bitfields && bpp !== 16 && bpp !== 32) throw new Error(`BMP: bitfields need 16 or 32 bits per pixel, not ${bpp}`);
  if (compression > BI_RLE4 && !bitfields) {
    const what = os2 && compression === 3 ? 'OS/2 Huffman' : os2 && compression === 4 ? 'OS/2 RLE24' : compression === 4 ? 'embedded JPEG' : compression === 5 ? 'embedded PNG' : `compression ${compression}`;
    throw new Error(`BMP: ${what} is not supported`);
  }

  let red = channel(0x7c00);
  let green = channel(0x03e0);
  let blue = channel(0x001f);
  let alphaChannel: Channel | undefined;
  if (bitfields) {
    // The masks sit right after the 40-byte core, inside a V2+ header or following a 40-byte one.
    const alphaInHeader = headerSize >= 56 || compression === BI_ALPHABITFIELDS;
    if (54 + (alphaInHeader ? 16 : 12) > bytes.length) throw new Error('BMP: truncated bitfield masks');
    red = channel(view.getUint32(54, true));
    green = channel(view.getUint32(58, true));
    blue = channel(view.getUint32(62, true));
    const alphaMask = alphaInHeader ? view.getUint32(66, true) : 0;
    if (alphaMask) alphaChannel = channel(alphaMask);
  }

  const paletteOffset = 14 + headerSize + (headerSize === 40 && compression === BI_BITFIELDS ? 12 : headerSize === 40 && compression === BI_ALPHABITFIELDS ? 16 : 0);
  let paletteCount = 0;
  if (bpp <= 8) {
    paletteCount = colorsUsed > 0 && colorsUsed < 1 << bpp ? colorsUsed : 1 << bpp;
    // Some writers store fewer entries than they declare; the pixel data offset tells.
    if (dataOffset > paletteOffset) paletteCount = Math.min(paletteCount, Math.floor((dataOffset - paletteOffset) / entrySize));
    paletteCount = Math.max(0, Math.min(paletteCount, Math.floor((bytes.length - paletteOffset) / entrySize)));
  }
  const palette = new Uint8Array(256 * 3);
  for (let i = 0; i < paletteCount; i++) {
    const q = paletteOffset + i * entrySize;
    palette[i * 3] = bytes[q + 2]!;
    palette[i * 3 + 1] = bytes[q + 1]!;
    palette[i * 3 + 2] = bytes[q]!;
  }
  const pixelOffset = dataOffset > 0 && dataOffset < bytes.length ? dataOffset : paletteOffset + paletteCount * entrySize;

  const rgba = new Uint8Array(width * height * 4);
  if (compression === BI_RLE8 || compression === BI_RLE4) {
    decodeRle(bytes, pixelOffset, width, height, topDown, palette, compression === BI_RLE4, rgba);
    return { width, height, rgba };
  }

  const stride = Math.floor((bpp * width + 31) / 32) * 4;
  const rowBytes = Math.ceil((bpp * width) / 8);
  // 32-bit BI_RGB has an unused byte that many writers fill with alpha and many leave at zero.
  const alphaByte = bpp === 32 && !bitfields;
  const hasAlpha = alphaByte || alphaChannel !== undefined;
  let anyAlpha = false;
  let decodedRows = 0;
  const indexMask = (1 << bpp) - 1;
  for (let row = 0; row < height; row++) {
    const p = pixelOffset + row * stride;
    if (p + rowBytes > bytes.length) break;
    const y = topDown ? row : height - 1 - row;
    let o = y * width * 4;
    for (let x = 0; x < width; x++, o += 4) {
      let a = 255;
      if (bpp <= 8) {
        const bit = x * bpp;
        const idx = (bytes[p + (bit >> 3)]! >> (8 - bpp - (bit & 7))) & indexMask;
        rgba[o] = palette[idx * 3]!;
        rgba[o + 1] = palette[idx * 3 + 1]!;
        rgba[o + 2] = palette[idx * 3 + 2]!;
      } else if (bpp === 24) {
        const q = p + x * 3;
        rgba[o] = bytes[q + 2]!;
        rgba[o + 1] = bytes[q + 1]!;
        rgba[o + 2] = bytes[q]!;
      } else if (alphaByte) {
        const q = p + x * 4;
        rgba[o] = bytes[q + 2]!;
        rgba[o + 1] = bytes[q + 1]!;
        rgba[o + 2] = bytes[q]!;
        a = bytes[q + 3]!;
      } else {
        const px = bpp === 16 ? view.getUint16(p + x * 2, true) : view.getUint32(p + x * 4, true);
        rgba[o] = extract(px, red);
        rgba[o + 1] = extract(px, green);
        rgba[o + 2] = extract(px, blue);
        if (alphaChannel) a = extract(px, alphaChannel);
      }
      if (a) anyAlpha = true;
      rgba[o + 3] = a;
    }
    decodedRows++;
  }
  // An alpha channel that is zero everywhere was never meant as alpha: show the image opaque.
  if (hasAlpha && !anyAlpha) {
    for (let row = 0; row < decodedRows; row++) {
      const y = topDown ? row : height - 1 - row;
      for (let x = 0; x < width; x++) rgba[(y * width + x) * 4 + 3] = 255;
    }
  }
  return { width, height, rgba };
}

/**
 * BI_RLE8 / BI_RLE4. Pixels the stream skips (with a delta, an early end of line or end of
 * bitmap) stay transparent, as browsers show them.
 */
function decodeRle(bytes: Uint8Array, start: number, width: number, height: number, topDown: boolean, palette: Uint8Array, four: boolean, rgba: Uint8Array): void {
  let x = 0;
  let row = 0;
  let p = start;
  const put = (idx: number) => {
    if (x < width && row < height) {
      const y = topDown ? row : height - 1 - row;
      const o = (y * width + x) * 4;
      rgba[o] = palette[idx * 3]!;
      rgba[o + 1] = palette[idx * 3 + 1]!;
      rgba[o + 2] = palette[idx * 3 + 2]!;
      rgba[o + 3] = 255;
    }
    x++;
  };
  while (p + 1 < bytes.length && row < height) {
    const n = bytes[p++]!;
    const c = bytes[p++]!;
    if (n > 0) {
      // Encoded run: in RLE4 the two nibbles of c alternate.
      for (let i = 0; i < n; i++) put(four ? (i & 1 ? c & 15 : c >> 4) : c);
    } else if (c === 0) {
      x = 0;
      row++;
    } else if (c === 1) {
      break;
    } else if (c === 2) {
      if (p + 1 >= bytes.length) break;
      x += bytes[p++]!;
      row += bytes[p++]!;
    } else {
      // Absolute run of c pixels, padded to a 16-bit boundary.
      const count = four ? (c + 1) >> 1 : c;
      if (p + count > bytes.length) break;
      for (let i = 0; i < c; i++) {
        const v = bytes[p + (four ? i >> 1 : i)]!;
        put(four ? (i & 1 ? v & 15 : v >> 4) : v);
      }
      p += count + (count & 1);
    }
  }
}
