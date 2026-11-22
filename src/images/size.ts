/** Pixel size (and resolution, when the file says) of PNG, JPEG, GIF, BMP and TIFF images, from their headers. */
import { sniffImage, type ImageKind } from './index.js';
import { metafileSize } from '../metafile/index.js';

export interface ImageInfo {
  kind: ImageKind;
  /** Natural size in points (pixels at the image's DPI, 96 when unknown). */
  width: number;
  height: number;
}

const u16be = (b: Uint8Array, i: number) => (b[i]! << 8) | b[i + 1]!;
const u32be = (b: Uint8Array, i: number) => ((b[i]! << 24) >>> 0) + (b[i + 1]! << 16) + (b[i + 2]! << 8) + b[i + 3]!;
const u16le = (b: Uint8Array, i: number) => b[i]! | (b[i + 1]! << 8);
const i32le = (b: Uint8Array, i: number) => b[i]! | (b[i + 1]! << 8) | (b[i + 2]! << 16) | (b[i + 3]! << 24);

/** Format and natural size of an image, or undefined when it is not one docxkeun can place. */
export function imageInfo(bytes: Uint8Array): ImageInfo | undefined {
  const kind = sniffImage(bytes);
  const px = (w: number, h: number, dpiX = 96, dpiY = dpiX): ImageInfo | undefined =>
    w > 0 && h > 0 && kind ? { kind, width: (w * 72) / (dpiX || 96), height: (h * 72) / (dpiY || 96) } : undefined;
  switch (kind) {
    case 'png': {
      if (bytes.length < 24) return undefined;
      let dpiX = 96;
      let dpiY = 96;
      // pHYs (pixels per metre), before IDAT.
      for (let i = 8; i + 8 <= bytes.length; ) {
        const len = u32be(bytes, i);
        const type = String.fromCharCode(bytes[i + 4]!, bytes[i + 5]!, bytes[i + 6]!, bytes[i + 7]!);
        if (type === 'pHYs' && bytes[i + 16] === 1) {
          dpiX = Math.round(u32be(bytes, i + 8) * 0.0254) || 96;
          dpiY = Math.round(u32be(bytes, i + 12) * 0.0254) || 96;
        }
        if (type === 'IDAT' || type === 'IEND') break;
        i += 12 + len;
      }
      return px(u32be(bytes, 16), u32be(bytes, 20), dpiX, dpiY);
    }
    case 'jpeg': {
      let dpi = 96;
      for (let i = 2; i + 9 < bytes.length; ) {
        if (bytes[i] !== 0xff) return undefined;
        const marker = bytes[i + 1]!;
        const len = u16be(bytes, i + 2);
        if (marker === 0xe0 && bytes[i + 4] === 0x4a && bytes[i + 11] === 1) dpi = u16be(bytes, i + 12) || 96; // JFIF, dots per inch
        if (marker >= 0xc0 && marker <= 0xcf && marker !== 0xc4 && marker !== 0xc8 && marker !== 0xcc) return px(u16be(bytes, i + 7), u16be(bytes, i + 5), dpi);
        i += 2 + len;
      }
      return undefined;
    }
    case 'gif':
      return px(u16le(bytes, 6), u16le(bytes, 8));
    case 'bmp': {
      const header = bytes[14]! | (bytes[15]! << 8);
      if (header === 12) return px(u16le(bytes, 18), u16le(bytes, 20));
      const ppm = i32le(bytes, 38);
      return px(Math.abs(i32le(bytes, 18)), Math.abs(i32le(bytes, 22)), ppm > 0 ? Math.round(ppm * 0.0254) : 96);
    }
    case 'tiff': {
      const le = bytes[0] === 0x49;
      const r16 = (i: number) => (le ? u16le(bytes, i) : u16be(bytes, i));
      const r32 = (i: number) => (le ? i32le(bytes, i) >>> 0 : u32be(bytes, i));
      const ifd = r32(4);
      let w = 0;
      let h = 0;
      for (let k = 0, n = r16(ifd); k < n; k++) {
        const e = ifd + 2 + k * 12;
        const tag = r16(e);
        const type = r16(e + 2);
        const v = type === 3 ? r16(e + 8) : r32(e + 8);
        if (tag === 256) w = v;
        if (tag === 257) h = v;
      }
      return px(w, h);
    }
    case 'emf':
    case 'wmf': {
      const s = metafileSize(bytes);
      return s ? { kind, width: s.width, height: s.height } : undefined;
    }
    default:
      return undefined;
  }
}
