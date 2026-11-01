/** Image XObjects: JPEG passes through; PNG is re-packed (alpha goes to a soft mask). */
import { ByteReader, concatBytes } from '../util/bytes.js';
import { inflate } from '../zip/deflate.js';
import { name, PdfStream, str, type PdfDict, type PdfRef } from './objects.js';
import type { PdfWriter } from './writer.js';

export class ImageFormatError extends Error {
  override name = 'ImageFormatError';
}

export function isJpeg(b: Uint8Array): boolean {
  return b[0] === 0xff && b[1] === 0xd8;
}

export function isPng(b: Uint8Array): boolean {
  return b[0] === 0x89 && b[1] === 0x50 && b[2] === 0x4e && b[3] === 0x47;
}

function jpegInfo(b: Uint8Array): { width: number; height: number; components: number; adobe: boolean } {
  const r = new ByteReader(b, 2);
  let adobe = false;
  while (r.pos + 4 <= b.length) {
    if (r.u8() !== 0xff) continue;
    let marker = r.u8();
    while (marker === 0xff) marker = r.u8();
    if (marker === 0xd9 || marker === 0xda) break;
    if (marker >= 0xd0 && marker <= 0xd7) continue;
    const length = r.u16be();
    const start = r.pos;
    if (marker === 0xee && b[start] === 0x41 && b[start + 1] === 0x64) adobe = true; // "Adobe" APP14
    if ((marker >= 0xc0 && marker <= 0xc3) || (marker >= 0xc5 && marker <= 0xc7) || (marker >= 0xc9 && marker <= 0xcb) || (marker >= 0xcd && marker <= 0xcf)) {
      r.skip(1);
      const height = r.u16be();
      const width = r.u16be();
      const components = r.u8();
      return { width, height, components, adobe };
    }
    r.seek(start + length - 2);
  }
  throw new ImageFormatError('JPEG without a frame header');
}

interface Png {
  width: number;
  height: number;
  bitDepth: number;
  colorType: number;
  interlace: number;
  palette?: Uint8Array;
  trns?: Uint8Array;
  idat: Uint8Array;
}

function parsePng(b: Uint8Array): Png {
  const r = new ByteReader(b, 8);
  const png: Partial<Png> = {};
  const idat: Uint8Array[] = [];
  while (r.pos + 8 <= b.length) {
    const length = r.u32be();
    const type = r.ascii(4);
    const data = r.bytesAt(length);
    r.skip(4); // crc
    if (type === 'IHDR') {
      const h = new ByteReader(data);
      png.width = h.u32be();
      png.height = h.u32be();
      png.bitDepth = h.u8();
      png.colorType = h.u8();
      h.skip(2);
      png.interlace = h.u8();
    } else if (type === 'PLTE') png.palette = data;
    else if (type === 'tRNS') png.trns = data;
    else if (type === 'IDAT') idat.push(data);
    else if (type === 'IEND') break;
  }
  if (!png.width || !png.height) throw new ImageFormatError('PNG without IHDR');
  return { ...(png as Png), idat: concatBytes(idat) };
}

const CHANNELS: Record<number, number> = { 0: 1, 2: 3, 3: 1, 4: 2, 6: 4 };

function paeth(a: number, b: number, c: number): number {
  const p = a + b - c;
  const pa = Math.abs(p - a);
  const pb = Math.abs(p - b);
  const pc = Math.abs(p - c);
  return pa <= pb && pa <= pc ? a : pb <= pc ? b : c;
}

/** Undo PNG scanline filters for one (sub)image. */
function unfilter(data: Uint8Array, offset: number, height: number, bpp: number, rowBytes: number, out: Uint8Array): number {
  let prev: Uint8Array = new Uint8Array(rowBytes);
  let p = offset;
  for (let y = 0; y < height; y++) {
    const filter = data[p++]!;
    const row = out.subarray(y * rowBytes, (y + 1) * rowBytes);
    for (let x = 0; x < rowBytes; x++) {
      const raw = data[p + x]!;
      const a = x >= bpp ? row[x - bpp]! : 0;
      const b = prev[x]!;
      const c = x >= bpp ? prev[x - bpp]! : 0;
      let v: number;
      switch (filter) {
        case 0:
          v = raw;
          break;
        case 1:
          v = raw + a;
          break;
        case 2:
          v = raw + b;
          break;
        case 3:
          v = raw + ((a + b) >> 1);
          break;
        case 4:
          v = raw + paeth(a, b, c);
          break;
        default:
          throw new ImageFormatError(`Bad PNG filter ${filter}`);
      }
      row[x] = v & 0xff;
    }
    p += rowBytes;
    prev = row;
  }
  return p;
}

/** Decode a PNG into 8-bit samples: colour channels plus an optional alpha plane. */
async function decodePng(png: Png): Promise<{ color: Uint8Array; alpha?: Uint8Array; channels: number }> {
  const raw = await inflate(png.idat, 'deflate');
  const ch = CHANNELS[png.colorType];
  if (!ch) throw new ImageFormatError(`Unsupported PNG colour type ${png.colorType}`);
  const bitsPerPixel = ch * png.bitDepth;
  const bpp = Math.max(1, bitsPerPixel >> 3);
  const { width, height } = png;
  const samples = new Uint8Array(width * height * ch);

  const place = (rows: Uint8Array, w: number, h: number, rowBytes: number, xs: (x: number) => number, ys: (y: number) => number) => {
    for (let y = 0; y < h; y++) {
      const row = rows.subarray(y * rowBytes, (y + 1) * rowBytes);
      for (let x = 0; x < w; x++) {
        for (let c = 0; c < ch; c++) {
          let v: number;
          if (png.bitDepth === 8) v = row[x * ch + c]!;
          else if (png.bitDepth === 16) v = row[(x * ch + c) * 2]!;
          else {
            const bit = (x * ch + c) * png.bitDepth;
            const byte = row[bit >> 3]!;
            const shift = 8 - png.bitDepth - (bit & 7);
            v = (byte >> shift) & ((1 << png.bitDepth) - 1);
            if (png.colorType === 0) v = Math.round((v * 255) / ((1 << png.bitDepth) - 1));
          }
          samples[(ys(y) * width + xs(x)) * ch + c] = v;
        }
      }
    }
  };

  if (png.interlace) {
    // Adam7 passes: [xStart, yStart, xStep, yStep]
    const passes = [
      [0, 0, 8, 8],
      [4, 0, 8, 8],
      [0, 4, 4, 8],
      [2, 0, 4, 4],
      [0, 2, 2, 4],
      [1, 0, 2, 2],
      [0, 1, 1, 2],
    ];
    let offset = 0;
    for (const [x0, y0, dx, dy] of passes) {
      const w = Math.ceil((width - x0!) / dx!);
      const h = Math.ceil((height - y0!) / dy!);
      if (w <= 0 || h <= 0) continue;
      const rowBytes = Math.ceil((w * bitsPerPixel) / 8);
      const rows = new Uint8Array(rowBytes * h);
      offset = unfilter(raw, offset, h, bpp, rowBytes, rows);
      place(rows, w, h, rowBytes, (x) => x0! + x * dx!, (y) => y0! + y * dy!);
    }
  } else {
    const rowBytes = Math.ceil((width * bitsPerPixel) / 8);
    const rows = new Uint8Array(rowBytes * height);
    unfilter(raw, 0, height, bpp, rowBytes, rows);
    place(rows, width, height, rowBytes, (x) => x, (y) => y);
  }

  const pixels = width * height;
  if (png.colorType === 3) {
    const pal = png.palette ?? new Uint8Array(768);
    const color = new Uint8Array(pixels * 3);
    const alpha = png.trns ? new Uint8Array(pixels) : undefined;
    for (let i = 0; i < pixels; i++) {
      const idx = samples[i]!;
      color.set(pal.subarray(idx * 3, idx * 3 + 3), i * 3);
      if (alpha) alpha[i] = idx < png.trns!.length ? png.trns![idx]! : 255;
    }
    return { color, alpha, channels: 3 };
  }
  if (png.colorType === 4 || png.colorType === 6) {
    const cc = ch - 1;
    const color = new Uint8Array(pixels * cc);
    const alpha = new Uint8Array(pixels);
    for (let i = 0; i < pixels; i++) {
      for (let c = 0; c < cc; c++) color[i * cc + c] = samples[i * ch + c]!;
      alpha[i] = samples[i * ch + cc]!;
    }
    return { color, alpha, channels: cc };
  }
  return { color: samples, channels: ch };
}

/** Add an image XObject; returns its reference and pixel size. */
export async function addImage(writer: PdfWriter, bytes: Uint8Array): Promise<{ ref: PdfRef; width: number; height: number }> {
  if (isJpeg(bytes)) {
    const info = jpegInfo(bytes);
    const space = info.components === 1 ? 'DeviceGray' : info.components === 4 ? 'DeviceCMYK' : 'DeviceRGB';
    const dict: PdfDict = {
      Type: name('XObject'),
      Subtype: name('Image'),
      Width: info.width,
      Height: info.height,
      ColorSpace: name(space),
      BitsPerComponent: 8,
      Filter: name('DCTDecode'),
    };
    // Adobe CMYK JPEGs store inverted values.
    if (info.components === 4 && info.adobe) dict.Decode = [1, 0, 1, 0, 1, 0, 1, 0];
    return { ref: writer.add(new PdfStream(dict, bytes)), width: info.width, height: info.height };
  }
  if (isPng(bytes)) {
    const png = parsePng(bytes);
    const simple = !png.interlace && (png.colorType === 0 || png.colorType === 2 || png.colorType === 3) && !png.trns && png.bitDepth <= 8;
    if (simple) {
      // The PNG's own zlib data is valid FlateDecode with the PNG predictor.
      const colors = png.colorType === 2 ? 3 : 1;
      const dict: PdfDict = {
        Type: name('XObject'),
        Subtype: name('Image'),
        Width: png.width,
        Height: png.height,
        BitsPerComponent: png.bitDepth,
        Filter: name('FlateDecode'),
        DecodeParms: { Predictor: 15, Colors: colors, BitsPerComponent: png.bitDepth, Columns: png.width },
        ColorSpace:
          png.colorType === 3
            ? [name('Indexed'), name('DeviceRGB'), (png.palette?.length ?? 3) / 3 - 1, str(png.palette ?? new Uint8Array(3), true)]
            : name(png.colorType === 2 ? 'DeviceRGB' : 'DeviceGray'),
      };
      return { ref: writer.add(new PdfStream(dict, png.idat)), width: png.width, height: png.height };
    }
    const decoded = await decodePng(png);
    let smask: PdfRef | undefined;
    if (decoded.alpha && decoded.alpha.some((a) => a !== 255)) {
      smask = await writer.stream({ Type: name('XObject'), Subtype: name('Image'), Width: png.width, Height: png.height, ColorSpace: name('DeviceGray'), BitsPerComponent: 8 }, decoded.alpha);
    }
    const ref = await writer.stream(
      {
        Type: name('XObject'),
        Subtype: name('Image'),
        Width: png.width,
        Height: png.height,
        ColorSpace: name(decoded.channels === 1 ? 'DeviceGray' : 'DeviceRGB'),
        BitsPerComponent: 8,
        SMask: smask,
      },
      decoded.color,
    );
    return { ref, width: png.width, height: png.height };
  }
  throw new ImageFormatError('Only JPEG and PNG images can be embedded in PDF');
}

