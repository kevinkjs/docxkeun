/**
 * TrueType subsetting for PDF embedding. Glyph ids stay the same (unused glyphs become empty),
 * so content streams can use the original glyph ids with Identity CID mapping.
 */
import { ByteWriter } from '../util/bytes.js';
import type { Font } from './sfnt.js';

const KEEP_TABLES = ['head', 'hhea', 'maxp', 'hmtx', 'cmap', 'cvt ', 'fpgm', 'prep', 'OS/2', 'name', 'post'];

/** Glyph ids used by composite glyphs, recursively. */
export function closeOverComposites(font: Font, glyphs: Set<number>): Set<number> {
  const out = new Set(glyphs);
  out.add(0);
  const queue = [...out];
  const view = new DataView(font.data.buffer, font.data.byteOffset, font.data.byteLength);
  while (queue.length) {
    const gid = queue.pop()!;
    const range = font.glyphRange(gid);
    if (!range || range.length < 10) continue;
    const numberOfContours = view.getInt16(range.offset);
    if (numberOfContours >= 0) continue;
    let p = range.offset + 10;
    const end = range.offset + range.length;
    for (;;) {
      if (p + 4 > end) break;
      const flags = view.getUint16(p);
      const component = view.getUint16(p + 2);
      if (!out.has(component)) {
        out.add(component);
        queue.push(component);
      }
      p += 4;
      p += flags & 0x0001 ? 4 : 2; // ARG_1_AND_2_ARE_WORDS
      if (flags & 0x0008) p += 2; // WE_HAVE_A_SCALE
      else if (flags & 0x0040) p += 4; // X_AND_Y_SCALE
      else if (flags & 0x0080) p += 8; // TWO_BY_TWO
      if (!(flags & 0x0020)) break; // MORE_COMPONENTS
    }
  }
  return out;
}

function checksum(bytes: Uint8Array): number {
  let sum = 0;
  const padded = bytes.length % 4 === 0 ? bytes : (() => {
    const p = new Uint8Array(bytes.length + (4 - (bytes.length % 4)));
    p.set(bytes);
    return p;
  })();
  const view = new DataView(padded.buffer, padded.byteOffset, padded.byteLength);
  for (let i = 0; i < padded.length; i += 4) sum = (sum + view.getUint32(i)) >>> 0;
  return sum;
}

/** Assemble an sfnt file from tables, recomputing checksums and head.checkSumAdjustment. */
export function buildSfnt(tables: Map<string, Uint8Array>, version = 0x00010000): Uint8Array {
  const tags = [...tables.keys()].sort();
  const numTables = tags.length;
  let searchRange = 1;
  let entrySelector = 0;
  while (searchRange * 2 <= numTables) {
    searchRange *= 2;
    entrySelector++;
  }
  searchRange *= 16;
  const out = new ByteWriter(16 * 1024);
  out.u32be(version).u16be(numTables).u16be(searchRange).u16be(entrySelector).u16be(numTables * 16 - searchRange);
  let offset = 12 + numTables * 16;
  const placements: Array<{ tag: string; offset: number; bytes: Uint8Array }> = [];
  for (const tag of tags) {
    const bytes = tables.get(tag)!;
    out.ascii(tag).u32be(checksum(bytes)).u32be(offset).u32be(bytes.length);
    placements.push({ tag, offset, bytes });
    offset += bytes.length + ((4 - (bytes.length % 4)) % 4);
  }
  for (const p of placements) {
    out.bytes(p.bytes);
    while (out.length % 4) out.u8(0);
  }
  const result = out.toBytes();
  const head = placements.find((p) => p.tag === 'head');
  if (head) {
    const view = new DataView(result.buffer);
    view.setUint32(head.offset + 8, 0);
    view.setUint32(head.offset + 8, (0xb1b0afba - checksum(result)) >>> 0);
  }
  return result;
}

/** A face as a standalone font file (extracts one face from a .ttc collection). */
export function standaloneFont(font: Font): Uint8Array {
  if (font.offset === 0) return font.data;
  const tables = new Map<string, Uint8Array>();
  for (const tag of font.tables.keys()) tables.set(tag, font.tableBytes(tag)!.slice());
  return buildSfnt(tables, font.isCFF ? 0x4f54544f : 0x00010000);
}

/**
 * Build a TrueType font containing only the given glyphs. CFF fonts are returned whole
 * (they are embedded as FontFile3/OpenType).
 */
export function subsetTrueType(font: Font, glyphs: Set<number>): Uint8Array {
  if (font.isCFF || !font.tables.has('glyf') || !font.tables.has('loca')) return standaloneFont(font);
  const keep = closeOverComposites(font, glyphs);

  // New glyf and loca (long format).
  const glyf = new ByteWriter(64 * 1024);
  const loca = new ByteWriter((font.numGlyphs + 1) * 4);
  for (let gid = 0; gid < font.numGlyphs; gid++) {
    loca.u32be(glyf.length);
    if (!keep.has(gid)) continue;
    const range = font.glyphRange(gid);
    if (!range || range.length === 0) continue;
    glyf.bytes(font.data.subarray(range.offset, range.offset + range.length));
    while (glyf.length % 4) glyf.u8(0);
  }
  loca.u32be(glyf.length);

  const tables = new Map<string, Uint8Array>();
  for (const tag of KEEP_TABLES) {
    const bytes = font.tableBytes(tag);
    if (bytes) tables.set(tag, bytes.slice());
  }
  tables.set('glyf', glyf.toBytes());
  tables.set('loca', loca.toBytes());
  const head = tables.get('head')!;
  new DataView(head.buffer).setInt16(50, 1); // indexToLocFormat = long
  const post = tables.get('post');
  if (post && post.length >= 32) {
    // Version 3: no glyph names (they are not needed and can be large).
    const short = post.slice(0, 32);
    new DataView(short.buffer).setUint32(0, 0x00030000);
    tables.set('post', short);
  }
  return buildSfnt(tables);
}
