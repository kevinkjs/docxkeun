/**
 * Image format detection and decoding for the bitmaps a .docx may embed. Word keeps the original
 * bytes of an inserted picture, and the part's file extension or content type is not reliable, so
 * the format is sniffed from the bytes. GIF, BMP and TIFF are decoded to RGBA here because PDF
 * cannot embed them as they are (and most browsers cannot show TIFF).
 */
import type { RasterImage } from '../layout/types.js';
import { decodeBmp, isBmp } from './bmp.js';
import { decodeGif, isGif } from './gif.js';
import { decodeTiff, isTiff } from './tiff.js';

export { decodeBmp, decodeGif, decodeTiff };

export type ImageKind = 'png' | 'jpeg' | 'gif' | 'bmp' | 'tiff' | 'emf' | 'wmf' | 'svg';

function isEmf(b: Uint8Array): boolean {
  // EMR_HEADER (record type 1) first, with the " EMF" signature in its dSignature field.
  return b.length >= 44 && b[0] === 1 && b[1] === 0 && b[2] === 0 && b[3] === 0 && b[40] === 0x20 && b[41] === 0x45 && b[42] === 0x4d && b[43] === 0x46;
}

function isWmf(b: Uint8Array): boolean {
  if (b.length >= 4 && b[0] === 0xd7 && b[1] === 0xcd && b[2] === 0xc6 && b[3] === 0x9a) return true; // placeable (Aldus) header
  if (b.length < 6) return false;
  const type = b[0]! | (b[1]! << 8);
  const headerWords = b[2]! | (b[3]! << 8);
  const version = b[4]! | (b[5]! << 8);
  return (type === 1 || type === 2) && headerWords === 9 && (version === 0x0100 || version === 0x0300);
}

function isSvg(b: Uint8Array): boolean {
  let text = '';
  const n = Math.min(b.length, 4096);
  for (let i = 0; i < n; i++) text += String.fromCharCode(b[i]!);
  // Allow a UTF-8 BOM, an XML declaration, comments and a doctype before the root element.
  text = text.replace(/^\xef\xbb\xbf/, '').trimStart();
  return text.startsWith('<') && /<svg[\s/>]/i.test(text);
}

/** Identify an image from its leading bytes. */
export function sniffImage(bytes: Uint8Array): ImageKind | undefined {
  const b = bytes;
  if (b.length >= 4 && b[0] === 0x89 && b[1] === 0x50 && b[2] === 0x4e && b[3] === 0x47) return 'png';
  if (b.length >= 3 && b[0] === 0xff && b[1] === 0xd8 && b[2] === 0xff) return 'jpeg';
  if (isGif(b)) return 'gif';
  if (isBmp(b)) return 'bmp';
  if (isTiff(b)) return 'tiff';
  if (isEmf(b)) return 'emf';
  if (isWmf(b)) return 'wmf';
  if (isSvg(b)) return 'svg';
  return undefined;
}

/** Whether every current browser can show this kind of image in an <img> as it is. */
export function browserShows(kind: ImageKind | undefined): boolean {
  return kind === 'png' || kind === 'jpeg' || kind === 'gif' || kind === 'bmp' || kind === 'svg';
}

/** Decode a GIF (first frame), BMP or TIFF (first page) to RGBA; throws for anything else. */
export async function decodeRaster(bytes: Uint8Array): Promise<RasterImage> {
  const kind = sniffImage(bytes);
  if (kind === 'gif') return decodeGif(bytes);
  if (kind === 'bmp') return decodeBmp(bytes);
  if (kind === 'tiff') return decodeTiff(bytes);
  throw new Error(`Cannot decode ${kind ? kind.toUpperCase() : 'unrecognised'} images to pixels`);
}
