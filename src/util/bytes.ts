/** Binary helpers shared by the zip, cfb, font, crypto and pdf modules. */

export function concatBytes(chunks: readonly Uint8Array[]): Uint8Array {
  let length = 0;
  for (const c of chunks) length += c.length;
  const out = new Uint8Array(length);
  let offset = 0;
  for (const c of chunks) {
    out.set(c, offset);
    offset += c.length;
  }
  return out;
}

export function bytesEqual(a: Uint8Array, b: Uint8Array): boolean {
  if (a.length !== b.length) return false;
  for (let i = 0; i < a.length; i++) if (a[i] !== b[i]) return false;
  return true;
}

/** Constant-time comparison for hashes and verifiers. */
export function timingSafeEqual(a: Uint8Array, b: Uint8Array): boolean {
  if (a.length !== b.length) return false;
  let diff = 0;
  for (let i = 0; i < a.length; i++) diff |= a[i]! ^ b[i]!;
  return diff === 0;
}

export function toHex(bytes: Uint8Array): string {
  let s = '';
  for (const b of bytes) s += b.toString(16).padStart(2, '0');
  return s;
}

export function fromHex(hex: string): Uint8Array {
  const clean = hex.replace(/\s+/g, '');
  if (clean.length % 2 !== 0) throw new Error('Odd-length hex string');
  const out = new Uint8Array(clean.length / 2);
  for (let i = 0; i < out.length; i++) {
    const byte = Number.parseInt(clean.substr(i * 2, 2), 16);
    if (Number.isNaN(byte)) throw new Error(`Invalid hex at ${i * 2}`);
    out[i] = byte;
  }
  return out;
}

/** Little-endian / big-endian reader over a byte array with a moving cursor. */
export class ByteReader {
  readonly view: DataView;
  pos: number;

  constructor(readonly bytes: Uint8Array, offset = 0) {
    this.view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
    this.pos = offset;
  }

  get length(): number {
    return this.bytes.length;
  }

  seek(pos: number): this {
    if (pos < 0 || pos > this.bytes.length) throw new RangeError(`Seek out of range: ${pos}`);
    this.pos = pos;
    return this;
  }

  skip(n: number): this {
    return this.seek(this.pos + n);
  }

  u8(): number {
    return this.view.getUint8(this.pos++);
  }
  i8(): number {
    return this.view.getInt8(this.pos++);
  }
  u16le(): number {
    const v = this.view.getUint16(this.pos, true);
    this.pos += 2;
    return v;
  }
  u32le(): number {
    const v = this.view.getUint32(this.pos, true);
    this.pos += 4;
    return v;
  }
  i32le(): number {
    const v = this.view.getInt32(this.pos, true);
    this.pos += 4;
    return v;
  }
  /** 64-bit little-endian as a JS number; throws above 2^53. */
  u64le(): number {
    const lo = this.u32le();
    const hi = this.u32le();
    const v = hi * 0x1_0000_0000 + lo;
    if (!Number.isSafeInteger(v)) throw new RangeError('64-bit value too large');
    return v;
  }
  u16be(): number {
    const v = this.view.getUint16(this.pos, false);
    this.pos += 2;
    return v;
  }
  i16be(): number {
    const v = this.view.getInt16(this.pos, false);
    this.pos += 2;
    return v;
  }
  u32be(): number {
    const v = this.view.getUint32(this.pos, false);
    this.pos += 4;
    return v;
  }
  i32be(): number {
    const v = this.view.getInt32(this.pos, false);
    this.pos += 4;
    return v;
  }
  bytesAt(n: number): Uint8Array {
    if (this.pos + n > this.bytes.length) throw new RangeError('Read past end of data');
    const out = this.bytes.subarray(this.pos, this.pos + n);
    this.pos += n;
    return out;
  }
  ascii(n: number): string {
    let s = '';
    for (const b of this.bytesAt(n)) s += String.fromCharCode(b);
    return s;
  }
}

/** Growable byte buffer with little-endian / big-endian writers. */
export class ByteWriter {
  private buf: Uint8Array;
  private view: DataView;
  length = 0;

  constructor(initialSize = 1024) {
    this.buf = new Uint8Array(initialSize);
    this.view = new DataView(this.buf.buffer);
  }

  private ensure(extra: number): void {
    const need = this.length + extra;
    if (need <= this.buf.length) return;
    let size = this.buf.length * 2;
    while (size < need) size *= 2;
    const next = new Uint8Array(size);
    next.set(this.buf.subarray(0, this.length));
    this.buf = next;
    this.view = new DataView(next.buffer);
  }

  u8(v: number): this {
    this.ensure(1);
    this.buf[this.length++] = v & 0xff;
    return this;
  }
  u16le(v: number): this {
    this.ensure(2);
    this.view.setUint16(this.length, v, true);
    this.length += 2;
    return this;
  }
  u32le(v: number): this {
    this.ensure(4);
    this.view.setUint32(this.length, v >>> 0, true);
    this.length += 4;
    return this;
  }
  i32le(v: number): this {
    this.ensure(4);
    this.view.setInt32(this.length, v, true);
    this.length += 4;
    return this;
  }
  u64le(v: number): this {
    this.u32le(v % 0x1_0000_0000);
    return this.u32le(Math.floor(v / 0x1_0000_0000));
  }
  u16be(v: number): this {
    this.ensure(2);
    this.view.setUint16(this.length, v, false);
    this.length += 2;
    return this;
  }
  i16be(v: number): this {
    this.ensure(2);
    this.view.setInt16(this.length, v, false);
    this.length += 2;
    return this;
  }
  u32be(v: number): this {
    this.ensure(4);
    this.view.setUint32(this.length, v >>> 0, false);
    this.length += 4;
    return this;
  }
  bytes(data: Uint8Array): this {
    this.ensure(data.length);
    this.buf.set(data, this.length);
    this.length += data.length;
    return this;
  }
  ascii(s: string): this {
    this.ensure(s.length);
    for (let i = 0; i < s.length; i++) this.buf[this.length++] = s.charCodeAt(i) & 0xff;
    return this;
  }
  zeros(n: number): this {
    this.ensure(n);
    this.buf.fill(0, this.length, this.length + n);
    this.length += n;
    return this;
  }
  /** Overwrite a 32-bit LE value at an earlier position (for back-patching sizes). */
  patchU32le(pos: number, v: number): this {
    this.view.setUint32(pos, v >>> 0, true);
    return this;
  }
  patchU32be(pos: number, v: number): this {
    this.view.setUint32(pos, v >>> 0, false);
    return this;
  }
  toBytes(): Uint8Array {
    return this.buf.slice(0, this.length);
  }
}
