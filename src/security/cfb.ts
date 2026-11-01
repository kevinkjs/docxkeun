/**
 * Compound File Binary format (MS-CFB): the OLE container that encrypted Office files use.
 * Reader supports v3/v4, DIFAT and mini streams; the writer produces v3 (512-byte sectors).
 */
import { ByteReader, ByteWriter } from '../util/bytes.js';
import { utf16leDecode, utf16leEncode } from '../util/text.js';

const SIGNATURE = [0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1];
const FREESECT = 0xffffffff;
const ENDOFCHAIN = 0xfffffffe;
const FATSECT = 0xfffffffd;
const DIFSECT = 0xfffffffc;
const NOSTREAM = 0xffffffff;
const MINI_CUTOFF = 4096;
const MINI_SECTOR = 64;

export class CfbError extends Error {
  override name = 'CfbError';
}

/** A storage (folder) in the compound file: child storages and streams by name. */
export interface CfbStorage {
  storages: Map<string, CfbStorage>;
  streams: Map<string, Uint8Array>;
}

export function emptyStorage(): CfbStorage {
  return { storages: new Map(), streams: new Map() };
}

interface DirEntry {
  name: string;
  type: number;
  left: number;
  right: number;
  child: number;
  start: number;
  size: number;
}

export function readCfb(bytes: Uint8Array): CfbStorage {
  for (let i = 0; i < 8; i++) if (bytes[i] !== SIGNATURE[i]) throw new CfbError('Not a compound file');
  const r = new ByteReader(bytes);
  r.seek(26);
  const major = r.u16le();
  r.seek(30);
  const sectorSize = 1 << r.u16le();
  const miniSectorSize = 1 << r.u16le();
  if ((major === 3 && sectorSize !== 512) || (major === 4 && sectorSize !== 4096)) throw new CfbError('Unexpected sector size');
  r.seek(44);
  const fatCount = r.u32le();
  const firstDir = r.u32le();
  r.skip(4);
  const cutoff = r.u32le();
  const firstMiniFat = r.u32le();
  const miniFatCount = r.u32le();
  let difatSector = r.u32le();
  const difatCount = r.u32le();
  const fatSectors: number[] = [];
  for (let i = 0; i < 109; i++) {
    const s = r.u32le();
    if (s !== FREESECT) fatSectors.push(s);
  }
  const sectorOffset = (s: number) => (s + 1) * sectorSize;
  for (let i = 0; i < difatCount && difatSector !== ENDOFCHAIN && difatSector !== FREESECT; i++) {
    const d = new ByteReader(bytes, sectorOffset(difatSector));
    for (let j = 0; j < sectorSize / 4 - 1; j++) {
      const s = d.u32le();
      if (s !== FREESECT) fatSectors.push(s);
    }
    difatSector = d.u32le();
  }
  if (fatSectors.length < fatCount) throw new CfbError('Truncated FAT');
  const per = sectorSize / 4;
  const fat = new Uint32Array(fatSectors.length * per);
  fatSectors.forEach((s, i) => {
    const off = sectorOffset(s);
    if (off + sectorSize > bytes.length) throw new CfbError('FAT sector out of range');
    const view = new DataView(bytes.buffer, bytes.byteOffset + off, sectorSize);
    for (let j = 0; j < per; j++) fat[i * per + j] = view.getUint32(j * 4, true);
  });

  const chain = (start: number, table: Uint32Array): number[] => {
    const out: number[] = [];
    const seen = new Set<number>();
    for (let s = start; s !== ENDOFCHAIN && s !== FREESECT; s = table[s] ?? ENDOFCHAIN) {
      if (seen.has(s) || s >= table.length) throw new CfbError('Corrupt sector chain');
      seen.add(s);
      out.push(s);
    }
    return out;
  };
  const readChain = (start: number, size?: number): Uint8Array => {
    const sectors = chain(start, fat);
    const out = new Uint8Array(sectors.length * sectorSize);
    sectors.forEach((s, i) => out.set(bytes.subarray(sectorOffset(s), sectorOffset(s) + sectorSize), i * sectorSize));
    return size === undefined ? out : out.subarray(0, size);
  };

  const dirBytes = readChain(firstDir);
  const entries: DirEntry[] = [];
  for (let off = 0; off + 128 <= dirBytes.length; off += 128) {
    const e = new ByteReader(dirBytes, off);
    const nameBytes = e.bytesAt(64);
    const nameLength = e.u16le();
    const type = e.u8();
    e.skip(1);
    const left = e.u32le();
    const right = e.u32le();
    const child = e.u32le();
    e.skip(36);
    const start = e.u32le();
    const sizeLo = e.u32le();
    const sizeHi = e.u32le();
    const size = major === 3 ? sizeLo : sizeHi * 0x1_0000_0000 + sizeLo;
    entries.push({ name: utf16leDecode(nameBytes.subarray(0, Math.max(0, nameLength - 2))), type, left, right, child, start, size });
  }
  const root = entries[0];
  if (!root || root.type !== 5) throw new CfbError('Missing root entry');

  const miniStream = root.size > 0 ? readChain(root.start, root.size) : new Uint8Array();
  const miniFat = new Uint32Array(miniFatCount * per);
  if (miniFatCount > 0) {
    const raw = readChain(firstMiniFat);
    const view = new DataView(raw.buffer, raw.byteOffset, raw.byteLength);
    for (let j = 0; j < miniFat.length; j++) miniFat[j] = view.getUint32(j * 4, true);
  }
  const readStream = (e: DirEntry): Uint8Array => {
    if (e.size === 0) return new Uint8Array();
    if (e.size < cutoff) {
      const sectors = chain(e.start, miniFat);
      const out = new Uint8Array(sectors.length * miniSectorSize);
      sectors.forEach((s, i) => out.set(miniStream.subarray(s * miniSectorSize, (s + 1) * miniSectorSize), i * miniSectorSize));
      return out.subarray(0, e.size);
    }
    return readChain(e.start, e.size);
  };

  const build = (entry: DirEntry, depth: number): CfbStorage => {
    if (depth > 64) throw new CfbError('Storage tree too deep');
    const storage = emptyStorage();
    const visit = (id: number, guard: Set<number>) => {
      if (id === NOSTREAM || guard.has(id)) return;
      guard.add(id);
      const e = entries[id];
      if (!e) return;
      visit(e.left, guard);
      if (e.type === 1) storage.storages.set(e.name, build(e, depth + 1));
      else if (e.type === 2) storage.streams.set(e.name, readStream(e));
      visit(e.right, guard);
    };
    visit(entry.child, new Set());
    return storage;
  };
  return build(root, 0);
}

/** Directory ordering: shorter names first, then case-insensitive code-unit order. */
function compareNames(a: string, b: string): number {
  if (a.length !== b.length) return a.length - b.length;
  const ua = a.toUpperCase();
  const ub = b.toUpperCase();
  return ua < ub ? -1 : ua > ub ? 1 : 0;
}

interface WriteEntry {
  name: string;
  type: number;
  data?: Uint8Array;
  children: WriteEntry[];
  id: number;
  left: number;
  right: number;
  child: number;
  start: number;
  size: number;
}

export function writeCfb(root: CfbStorage): Uint8Array {
  const sectorSize = 512;
  const per = sectorSize / 4;
  const all: WriteEntry[] = [];
  const make = (name: string, type: number, data?: Uint8Array): WriteEntry => {
    const e: WriteEntry = { name, type, data, children: [], id: all.length, left: NOSTREAM, right: NOSTREAM, child: NOSTREAM, start: ENDOFCHAIN, size: data?.length ?? 0 };
    all.push(e);
    return e;
  };
  const rootEntry = make('Root Entry', 5);
  const addStorage = (entry: WriteEntry, storage: CfbStorage) => {
    for (const [name, child] of storage.storages) {
      if (name.length > 31) throw new CfbError(`Name too long: ${name}`);
      const e = make(name, 1);
      entry.children.push(e);
      addStorage(e, child);
    }
    for (const [name, data] of storage.streams) {
      if (name.length > 31) throw new CfbError(`Name too long: ${name}`);
      entry.children.push(make(name, 2, data));
    }
  };
  addStorage(rootEntry, root);
  // Siblings form a balanced binary search tree (all nodes black).
  const link = (entry: WriteEntry) => {
    const sorted = [...entry.children].sort((a, b) => compareNames(a.name, b.name));
    const balance = (lo: number, hi: number): number => {
      if (lo > hi) return NOSTREAM;
      const mid = (lo + hi) >> 1;
      const node = sorted[mid]!;
      node.left = balance(lo, mid - 1);
      node.right = balance(mid + 1, hi);
      return node.id;
    };
    entry.child = balance(0, sorted.length - 1);
    for (const c of entry.children) if (c.type === 1) link(c);
  };
  link(rootEntry);

  // Mini stream: small streams packed in 64-byte mini sectors.
  const small = all.filter((e) => e.type === 2 && e.size > 0 && e.size < MINI_CUTOFF);
  const big = all.filter((e) => e.type === 2 && e.size >= MINI_CUTOFF);
  const miniFat: number[] = [];
  const miniData = new ByteWriter(4096);
  for (const e of small) {
    const n = Math.ceil(e.size / MINI_SECTOR);
    e.start = miniFat.length;
    for (let i = 0; i < n; i++) miniFat.push(i === n - 1 ? ENDOFCHAIN : miniFat.length + 1);
    miniData.bytes(e.data!).zeros(n * MINI_SECTOR - e.size);
  }
  const miniStream = miniData.toBytes();
  rootEntry.size = miniStream.length;

  const sectorsFor = (n: number) => Math.ceil(n / sectorSize);
  const dirSectors = sectorsFor(all.length * 128);
  const miniFatSectors = sectorsFor(miniFat.length * 4);
  const miniStreamSectors = sectorsFor(miniStream.length);
  const bigSectors = big.reduce((n, e) => n + sectorsFor(e.size), 0);
  const content = dirSectors + miniFatSectors + miniStreamSectors + bigSectors;
  // FAT and DIFAT sectors also need FAT entries: iterate to a fixed point.
  let fatSectors = 1;
  let difatSectors = 0;
  for (;;) {
    const total = content + fatSectors + difatSectors;
    const needFat = Math.ceil(total / per);
    const needDifat = needFat > 109 ? Math.ceil((needFat - 109) / (per - 1)) : 0;
    if (needFat === fatSectors && needDifat === difatSectors) break;
    fatSectors = needFat;
    difatSectors = needDifat;
  }
  const totalSectors = content + fatSectors + difatSectors;
  const fat = new Uint32Array(fatSectors * per).fill(FREESECT);
  let next = 0;
  const allocate = (count: number, mark?: number): number => {
    const start = next;
    for (let i = 0; i < count; i++) fat[next + i] = mark ?? (i === count - 1 ? ENDOFCHAIN : next + i + 1);
    next += count;
    return count ? start : ENDOFCHAIN;
  };
  const fatStart = allocate(fatSectors, FATSECT);
  const difatStart = allocate(difatSectors, DIFSECT);
  const dirStart = allocate(dirSectors);
  const miniFatStart = allocate(miniFatSectors);
  rootEntry.start = miniStreamSectors ? allocate(miniStreamSectors) : ENDOFCHAIN;
  for (const e of big) e.start = allocate(sectorsFor(e.size));
  if (next !== totalSectors) throw new CfbError('Sector accounting error');

  const out = new ByteWriter((totalSectors + 1) * sectorSize);
  // Header
  out.bytes(Uint8Array.from(SIGNATURE)).zeros(16).u16le(0x3e).u16le(3).u16le(0xfffe).u16le(9).u16le(6).zeros(6);
  out.u32le(0).u32le(fatSectors).u32le(dirStart).u32le(0).u32le(MINI_CUTOFF);
  out.u32le(miniFatSectors ? miniFatStart : ENDOFCHAIN).u32le(miniFatSectors);
  out.u32le(difatSectors ? difatStart : ENDOFCHAIN).u32le(difatSectors);
  const fatList = Array.from({ length: fatSectors }, (_, i) => fatStart + i);
  for (let i = 0; i < 109; i++) out.u32le(fatList[i] ?? FREESECT);
  // FAT sectors
  for (const v of fat) out.u32le(v);
  // DIFAT sectors: 127 FAT locations each, then the next DIFAT sector.
  for (let d = 0; d < difatSectors; d++) {
    for (let j = 0; j < per - 1; j++) out.u32le(fatList[109 + d * (per - 1) + j] ?? FREESECT);
    out.u32le(d === difatSectors - 1 ? ENDOFCHAIN : difatStart + d + 1);
  }
  // Directory
  for (const e of all) {
    const name = utf16leEncode(e.name);
    out.bytes(name).zeros(64 - name.length).u16le(e.name.length ? name.length + 2 : 0);
    out.u8(e.type).u8(1).u32le(e.left).u32le(e.right).u32le(e.child).zeros(16).u32le(0).zeros(16);
    out.u32le(e.type === 1 ? 0 : e.size === 0 && e.type === 2 ? ENDOFCHAIN : e.start).u32le(e.size).u32le(0);
  }
  // Unused directory slots are empty entries with no siblings or children.
  for (let i = all.length; i < dirSectors * 4; i++) out.zeros(66).u8(0).u8(0).u32le(NOSTREAM).u32le(NOSTREAM).u32le(NOSTREAM).zeros(48);
  for (let i = 0; i < miniFatSectors * per; i++) out.u32le(miniFat[i] ?? FREESECT);
  out.bytes(miniStream).zeros(miniStreamSectors * sectorSize - miniStream.length);
  for (const e of big) out.bytes(e.data!).zeros(sectorsFor(e.size) * sectorSize - e.size);
  return out.toBytes();
}
