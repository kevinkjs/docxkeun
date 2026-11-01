/** PDF object model and serialisation (ISO 32000). */
import { ByteWriter } from '../util/bytes.js';

export class PdfName {
  constructor(readonly name: string) {}
}

export class PdfRef {
  constructor(readonly num: number, readonly gen = 0) {}
}

/** A string object. Text is written as PDFDocEncoding when possible, else UTF-16BE with a BOM. */
export class PdfString {
  constructor(readonly value: string | Uint8Array, readonly hex = false) {}

  bytes(): Uint8Array {
    if (this.value instanceof Uint8Array) return this.value;
    const s = this.value;
    if (/^[\x20-\x7e\t\r\n]*$/.test(s)) {
      const out = new Uint8Array(s.length);
      for (let i = 0; i < s.length; i++) out[i] = s.charCodeAt(i);
      return out;
    }
    const out = new Uint8Array(2 + s.length * 2);
    out[0] = 0xfe;
    out[1] = 0xff;
    for (let i = 0; i < s.length; i++) {
      out[2 + i * 2] = s.charCodeAt(i) >> 8;
      out[3 + i * 2] = s.charCodeAt(i) & 0xff;
    }
    return out;
  }
}

export class PdfStream {
  constructor(readonly dict: PdfDict, public data: Uint8Array) {}
}

export type PdfValue = number | boolean | null | PdfName | PdfRef | PdfString | PdfValue[] | PdfDict | PdfStream;
export interface PdfDict {
  [key: string]: PdfValue | undefined;
}

export const name = (n: string) => new PdfName(n);
export const str = (s: string | Uint8Array, hex = false) => new PdfString(s, hex);

export function num(v: number): string {
  if (!Number.isFinite(v)) return '0';
  if (Number.isInteger(v)) return String(v);
  const s = v.toFixed(4).replace(/0+$/, '').replace(/\.$/, '');
  return s === '-0' ? '0' : s;
}

function escapeName(n: string): string {
  let out = '';
  for (const ch of new TextEncoder().encode(n)) {
    if (ch < 0x21 || ch > 0x7e || '#()<>[]{}/%'.includes(String.fromCharCode(ch))) out += '#' + ch.toString(16).padStart(2, '0');
    else out += String.fromCharCode(ch);
  }
  return out;
}

function literalString(bytes: Uint8Array): Uint8Array {
  const w = new ByteWriter(bytes.length + 16);
  w.u8(0x28);
  for (const b of bytes) {
    if (b === 0x28 || b === 0x29 || b === 0x5c) w.u8(0x5c).u8(b);
    else if (b === 0x0d) w.ascii('\\r');
    else if (b === 0x0a) w.ascii('\\n');
    else w.u8(b);
  }
  w.u8(0x29);
  return w.toBytes();
}

function hexString(bytes: Uint8Array): string {
  let s = '<';
  for (const b of bytes) s += b.toString(16).padStart(2, '0');
  return s + '>';
}

/** Encrypts strings and streams of one object (set by the security handler). */
export interface ObjectCrypt {
  string(bytes: Uint8Array): Uint8Array;
  stream(bytes: Uint8Array): Uint8Array;
}

export function writeValue(w: ByteWriter, v: PdfValue | undefined, crypt?: ObjectCrypt): void {
  if (v === undefined || v === null) {
    w.ascii('null');
  } else if (typeof v === 'number') {
    w.ascii(num(v));
  } else if (typeof v === 'boolean') {
    w.ascii(v ? 'true' : 'false');
  } else if (v instanceof PdfName) {
    w.ascii('/' + escapeName(v.name));
  } else if (v instanceof PdfRef) {
    w.ascii(`${v.num} ${v.gen} R`);
  } else if (v instanceof PdfString) {
    const bytes = crypt ? crypt.string(v.bytes()) : v.bytes();
    if (v.hex || crypt) w.ascii(hexString(bytes));
    else w.bytes(literalString(bytes));
  } else if (Array.isArray(v)) {
    w.ascii('[');
    v.forEach((item, i) => {
      if (i) w.ascii(' ');
      writeValue(w, item, crypt);
    });
    w.ascii(']');
  } else if (v instanceof PdfStream) {
    const data = crypt ? crypt.stream(v.data) : v.data;
    writeValue(w, { ...v.dict, Length: data.length }, crypt);
    w.ascii('\nstream\n').bytes(data).ascii('\nendstream');
  } else {
    w.ascii('<<');
    for (const [k, val] of Object.entries(v)) {
      if (val === undefined) continue;
      w.ascii('/' + escapeName(k) + ' ');
      writeValue(w, val, crypt);
      w.ascii(' ');
    }
    w.ascii('>>');
  }
}

/** PDF date string: D:YYYYMMDDHHmmSSZ */
export function pdfDate(d: Date): string {
  const p = (n: number) => String(n).padStart(2, '0');
  return `D:${d.getUTCFullYear()}${p(d.getUTCMonth() + 1)}${p(d.getUTCDate())}${p(d.getUTCHours())}${p(d.getUTCMinutes())}${p(d.getUTCSeconds())}Z`;
}
