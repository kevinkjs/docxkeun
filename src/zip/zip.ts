/**
 * ZIP reader and writer (APPNOTE 6.3.x subset used by OOXML packages).
 * Reading supports stored/deflated entries, data descriptors and ZIP64 sizes.
 * Unchanged entries can be written back from their original compressed bytes.
 */
import { ByteReader, ByteWriter } from '../util/bytes.js';
import { crc32 } from '../util/crc32.js';
import { utf8Decode, utf8Encode } from '../util/text.js';
import { deflate, inflate } from './deflate.js';

const SIG_LOCAL = 0x04034b50;
const SIG_CENTRAL = 0x02014b50;
const SIG_EOCD = 0x06054b50;
const SIG_EOCD64 = 0x06064b50;
const SIG_EOCD64_LOCATOR = 0x07064b50;

export const METHOD_STORED = 0;
export const METHOD_DEFLATE = 8;

/** Compressed form of an entry as found in an archive, reusable when writing. */
export interface ZipRawData {
  method: number;
  crc32: number;
  size: number;
  compressed: Uint8Array;
}

export interface ZipFileEntry {
  readonly name: string;
  readonly raw: ZipRawData;
  readonly modified: Date;
  /** Decompress and return the entry's bytes. */
  read(): Promise<Uint8Array>;
}

export class ZipError extends Error {
  override name = 'ZipError';
}

function findEocd(r: ByteReader): number {
  const min = Math.max(0, r.length - 22 - 0xffff);
  for (let i = r.length - 22; i >= min; i--) {
    if (r.view.getUint32(i, true) === SIG_EOCD) return i;
  }
  throw new ZipError('Not a zip archive (end of central directory not found)');
}

function dosToDate(time: number, date: number): Date {
  const year = ((date >> 9) & 0x7f) + 1980;
  const month = ((date >> 5) & 0x0f) - 1;
  const day = date & 0x1f;
  return new Date(year, month, day || 1, (time >> 11) & 0x1f, (time >> 5) & 0x3f, (time & 0x1f) * 2);
}

function dateToDos(d: Date): { time: number; date: number } {
  const year = Math.min(Math.max(d.getFullYear(), 1980), 2107);
  return {
    time: (d.getHours() << 11) | (d.getMinutes() << 5) | Math.floor(d.getSeconds() / 2),
    date: ((year - 1980) << 9) | ((d.getMonth() + 1) << 5) | d.getDate(),
  };
}

export function readZip(bytes: Uint8Array): ZipFileEntry[] {
  const r = new ByteReader(bytes);
  const eocd = findEocd(r);
  r.seek(eocd + 10);
  let count = r.u16le();
  let cdSize = r.u32le();
  let cdOffset = r.u32le();

  if (count === 0xffff || cdSize === 0xffffffff || cdOffset === 0xffffffff) {
    const locator = eocd - 20;
    if (locator < 0 || r.view.getUint32(locator, true) !== SIG_EOCD64_LOCATOR) {
      throw new ZipError('ZIP64 archive without a locator');
    }
    r.seek(locator + 8);
    const eocd64 = r.u64le();
    r.seek(eocd64);
    if (r.u32le() !== SIG_EOCD64) throw new ZipError('Bad ZIP64 end of central directory');
    r.skip(8 + 2 + 2 + 4 + 4 + 8);
    count = r.u64le();
    cdSize = r.u64le();
    cdOffset = r.u64le();
  }

  const entries: ZipFileEntry[] = [];
  r.seek(cdOffset);
  for (let i = 0; i < count; i++) {
    if (r.u32le() !== SIG_CENTRAL) throw new ZipError(`Bad central directory entry ${i}`);
    r.skip(4); // version made by, version needed
    const flags = r.u16le();
    const method = r.u16le();
    const time = r.u16le();
    const date = r.u16le();
    const crc = r.u32le();
    let compressedSize = r.u32le();
    let size = r.u32le();
    const nameLength = r.u16le();
    const extraLength = r.u16le();
    const commentLength = r.u16le();
    r.skip(8); // disk start, internal attrs, external attrs
    let localOffset = r.u32le();
    const nameBytes = r.bytesAt(nameLength);
    const extraStart = r.pos;
    // ZIP64 extended information (0x0001) holds only the fields that overflowed, in this order.
    while (r.pos + 4 <= extraStart + extraLength) {
      const id = r.u16le();
      const len = r.u16le();
      const end = r.pos + len;
      if (id === 0x0001) {
        if (size === 0xffffffff) size = r.u64le();
        if (compressedSize === 0xffffffff) compressedSize = r.u64le();
        if (localOffset === 0xffffffff) localOffset = r.u64le();
      }
      r.seek(end);
    }
    r.seek(extraStart + extraLength + commentLength);

    if (flags & 0x1) throw new ZipError('Encrypted zip entries are not supported');
    // Bit 11 means UTF-8; legacy names are CP437, and ASCII is the same in both.
    const name = utf8Decode(nameBytes);

    const local = new ByteReader(bytes, localOffset);
    if (local.u32le() !== SIG_LOCAL) throw new ZipError(`Bad local header for ${name}`);
    local.skip(22);
    const localNameLength = local.u16le();
    const localExtraLength = local.u16le();
    const dataStart = localOffset + 30 + localNameLength + localExtraLength;
    if (dataStart + compressedSize > bytes.length) throw new ZipError(`Truncated entry ${name}`);
    const compressed = bytes.subarray(dataStart, dataStart + compressedSize);
    const raw: ZipRawData = { method, crc32: crc, size, compressed };

    entries.push({
      name,
      raw,
      modified: dosToDate(time, date),
      async read() {
        let data: Uint8Array;
        if (method === METHOD_STORED) data = compressed;
        else if (method === METHOD_DEFLATE) data = await inflate(compressed, 'deflate-raw');
        else throw new ZipError(`Unsupported compression method ${method} for ${name}`);
        if (data.length !== size) throw new ZipError(`Size mismatch for ${name}`);
        if (crc32(data) !== crc) throw new ZipError(`CRC mismatch for ${name}`);
        return data;
      },
    });
  }
  return entries;
}

export type ZipWriteEntry =
  | { name: string; data: Uint8Array; compress?: boolean }
  | { name: string; raw: ZipRawData };

export interface WriteZipOptions {
  /** Timestamp stored on every entry. Defaults to now; fix it for reproducible output. */
  modified?: Date;
}

export async function writeZip(entries: readonly ZipWriteEntry[], options: WriteZipOptions = {}): Promise<Uint8Array> {
  const { time, date } = dateToDos(options.modified ?? new Date());
  const prepared: Array<{ name: Uint8Array; raw: ZipRawData; offset: number }> = [];
  const seen = new Set<string>();

  for (const entry of entries) {
    if (seen.has(entry.name)) throw new ZipError(`Duplicate zip entry ${entry.name}`);
    seen.add(entry.name);
    let raw: ZipRawData;
    if ('raw' in entry) {
      raw = entry.raw;
    } else {
      const crc = crc32(entry.data);
      const shouldCompress = entry.compress ?? entry.data.length > 64;
      if (shouldCompress) {
        const compressed = await deflate(entry.data, 'deflate-raw');
        raw = compressed.length < entry.data.length
          ? { method: METHOD_DEFLATE, crc32: crc, size: entry.data.length, compressed }
          : { method: METHOD_STORED, crc32: crc, size: entry.data.length, compressed: entry.data };
      } else {
        raw = { method: METHOD_STORED, crc32: crc, size: entry.data.length, compressed: entry.data };
      }
    }
    if (raw.size > 0xfffffffe || raw.compressed.length > 0xfffffffe) {
      throw new ZipError(`Entry ${entry.name} needs ZIP64, which the writer does not support`);
    }
    prepared.push({ name: utf8Encode(entry.name), raw, offset: 0 });
  }

  const w = new ByteWriter(1 << 16);
  for (const p of prepared) {
    p.offset = w.length;
    w.u32le(SIG_LOCAL)
      .u16le(20) // version needed: 2.0
      .u16le(0x0800) // UTF-8 names
      .u16le(p.raw.method)
      .u16le(time)
      .u16le(date)
      .u32le(p.raw.crc32)
      .u32le(p.raw.compressed.length)
      .u32le(p.raw.size)
      .u16le(p.name.length)
      .u16le(0)
      .bytes(p.name)
      .bytes(p.raw.compressed);
  }
  const cdStart = w.length;
  for (const p of prepared) {
    w.u32le(SIG_CENTRAL)
      .u16le(20)
      .u16le(20)
      .u16le(0x0800)
      .u16le(p.raw.method)
      .u16le(time)
      .u16le(date)
      .u32le(p.raw.crc32)
      .u32le(p.raw.compressed.length)
      .u32le(p.raw.size)
      .u16le(p.name.length)
      .u16le(0) // extra
      .u16le(0) // comment
      .u16le(0) // disk
      .u16le(0) // internal attrs
      .u32le(0) // external attrs
      .u32le(p.offset)
      .bytes(p.name);
  }
  const cdSize = w.length - cdStart;
  if (prepared.length > 0xfffe || cdStart > 0xfffffffe) throw new ZipError('Archive needs ZIP64');
  w.u32le(SIG_EOCD)
    .u16le(0)
    .u16le(0)
    .u16le(prepared.length)
    .u16le(prepared.length)
    .u32le(cdSize)
    .u32le(cdStart)
    .u16le(0);
  return w.toBytes();
}
