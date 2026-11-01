/**
 * TrueType / OpenType (sfnt) parser: the metrics layout needs and the data PDF embedding needs.
 * Supports TrueType outlines (glyf) and CFF-flavoured OpenType, plus .ttc collections.
 */
import { ByteReader } from '../util/bytes.js';

export interface TableRecord {
  offset: number;
  length: number;
}

export class FontParseError extends Error {
  override name = 'FontParseError';
}

export class Font {
  readonly tables = new Map<string, TableRecord>();
  unitsPerEm = 1000;
  numGlyphs = 0;
  bbox: [number, number, number, number] = [0, 0, 0, 0];
  indexToLocFormat = 0;
  ascender = 800;
  descender = -200;
  lineGap = 0;
  winAscent = 0;
  winDescent = 0;
  typoAscender = 0;
  typoDescender = 0;
  typoLineGap = 0;
  useTypoMetrics = false;
  capHeight = 700;
  xHeight = 500;
  weightClass = 400;
  italicAngle = 0;
  underlinePosition = -100;
  underlineThickness = 50;
  strikeoutPosition = 250;
  strikeoutSize = 50;
  isFixedPitch = false;
  /** OS/2 fsType: 0x0002 means restricted-license embedding (must not embed). */
  fsType = 0;
  familyName = '';
  subfamilyName = '';
  fullName = '';
  postscriptName = '';
  macStyle = 0;
  isSymbolic = false;
  private numberOfHMetrics = 0;
  private advances!: Uint16Array;
  private cmap = new Map<number, number>();
  private kernPairs: Map<number, number> | undefined;

  /**
   * @param data the font file (or a whole .ttc collection)
   * @param offset start of the table directory (for collections)
   */
  constructor(readonly data: Uint8Array, readonly offset = 0) {
    const r = new ByteReader(data, offset);
    const version = r.u32be();
    if (version === 0x74746366 /* ttcf */) throw new FontParseError('Font collection: use Font.fromCollection');
    if (version !== 0x00010000 && version !== 0x4f54544f /* OTTO */ && version !== 0x74727565 /* true */) {
      throw new FontParseError('Not a TrueType/OpenType font');
    }
    const numTables = r.u16be();
    r.skip(6);
    for (let i = 0; i < numTables; i++) {
      const tag = r.ascii(4);
      r.skip(4); // checksum
      const tableOffset = r.u32be();
      const length = r.u32be();
      if (tableOffset + length > data.length) throw new FontParseError(`Table ${tag} out of bounds`);
      this.tables.set(tag, { offset: tableOffset, length });
    }
    for (const required of ['head', 'hhea', 'hmtx', 'maxp', 'cmap']) {
      if (!this.tables.has(required)) throw new FontParseError(`Missing ${required} table`);
    }
    this.parseHead();
    this.parseHhea();
    this.parseMaxp();
    this.parseHmtx();
    this.parseOs2();
    this.parsePost();
    this.parseName();
    this.parseCmap();
  }

  /** Fonts inside a .ttc collection. */
  static fromCollection(data: Uint8Array): Font[] {
    const r = new ByteReader(data);
    if (r.u32be() !== 0x74746366) return [new Font(data)];
    r.skip(4);
    const count = r.u32be();
    const fonts: Font[] = [];
    for (let i = 0; i < count; i++) fonts.push(new Font(data, r.u32be()));
    return fonts;
  }

  /** Parse any supported font file; for collections the first face is returned. */
  static parse(data: Uint8Array): Font {
    return Font.fromCollection(data)[0]!;
  }

  get isCFF(): boolean {
    return this.tables.has('CFF ') || this.tables.has('CFF2');
  }

  get isBold(): boolean {
    return this.weightClass >= 600 || (this.macStyle & 1) !== 0;
  }

  get isItalic(): boolean {
    return (this.macStyle & 2) !== 0 || this.italicAngle !== 0;
  }

  get embeddingAllowed(): boolean {
    return (this.fsType & 0x000f) !== 0x0002;
  }

  table(tag: string): ByteReader | undefined {
    const t = this.tables.get(tag);
    return t ? new ByteReader(this.data.subarray(0, t.offset + t.length), t.offset) : undefined;
  }

  tableBytes(tag: string): Uint8Array | undefined {
    const t = this.tables.get(tag);
    return t ? this.data.subarray(t.offset, t.offset + t.length) : undefined;
  }

  private parseHead(): void {
    const r = this.table('head')!;
    r.skip(18);
    this.unitsPerEm = r.u16be() || 1000;
    r.skip(16);
    this.bbox = [r.i16be(), r.i16be(), r.i16be(), r.i16be()];
    this.macStyle = r.u16be();
    r.skip(4);
    this.indexToLocFormat = r.i16be();
  }

  private parseHhea(): void {
    const r = this.table('hhea')!;
    r.skip(4);
    this.ascender = r.i16be();
    this.descender = r.i16be();
    this.lineGap = r.i16be();
    r.skip(24);
    this.numberOfHMetrics = r.u16be();
  }

  private parseMaxp(): void {
    const r = this.table('maxp')!;
    r.skip(4);
    this.numGlyphs = r.u16be();
  }

  private parseHmtx(): void {
    const r = this.table('hmtx')!;
    const n = Math.max(1, Math.min(this.numberOfHMetrics, this.numGlyphs || this.numberOfHMetrics));
    this.advances = new Uint16Array(Math.max(this.numGlyphs, n));
    let last = 0;
    for (let i = 0; i < n; i++) {
      last = r.u16be();
      this.advances[i] = last;
      r.skip(2);
    }
    for (let i = n; i < this.advances.length; i++) this.advances[i] = last;
  }

  private parseOs2(): void {
    const r = this.table('OS/2');
    this.winAscent = this.ascender;
    this.winDescent = -this.descender;
    if (!r) return;
    const version = r.u16be();
    r.skip(2);
    this.weightClass = r.u16be();
    r.skip(2);
    this.fsType = r.u16be();
    r.skip(16); // subscript / superscript metrics
    this.strikeoutSize = r.i16be();
    this.strikeoutPosition = r.i16be();
    r.skip(2 + 10 + 16 + 4);
    const fsSelection = r.u16be();
    this.useTypoMetrics = (fsSelection & 0x80) !== 0;
    r.skip(4);
    this.typoAscender = r.i16be();
    this.typoDescender = r.i16be();
    this.typoLineGap = r.i16be();
    this.winAscent = r.u16be();
    this.winDescent = r.u16be();
    if (version >= 2 && this.tables.get('OS/2')!.length >= 96) {
      r.skip(8);
      this.xHeight = r.i16be();
      this.capHeight = r.i16be();
    } else {
      this.capHeight = Math.round(this.ascender * 0.7);
      this.xHeight = Math.round(this.ascender * 0.5);
    }
  }

  private parsePost(): void {
    const r = this.table('post');
    if (!r) return;
    r.skip(4);
    this.italicAngle = r.i32be() / 65536;
    this.underlinePosition = r.i16be();
    this.underlineThickness = r.i16be();
    this.isFixedPitch = r.u32be() !== 0;
  }

  private parseName(): void {
    const r = this.table('name');
    if (!r) return;
    const base = r.pos;
    r.skip(2);
    const count = r.u16be();
    const stringOffset = r.u16be();
    const found: Record<number, { text: string; score: number }> = {};
    for (let i = 0; i < count; i++) {
      const platform = r.u16be();
      const encoding = r.u16be();
      const language = r.u16be();
      const nameId = r.u16be();
      const length = r.u16be();
      const off = r.u16be();
      if (nameId > 6) continue;
      const start = base + stringOffset + off;
      if (start + length > this.data.length) continue;
      const bytes = this.data.subarray(start, start + length);
      let text: string;
      let score: number;
      if (platform === 3 || platform === 0) {
        text = '';
        for (let j = 0; j + 1 < bytes.length; j += 2) text += String.fromCharCode((bytes[j]! << 8) | bytes[j + 1]!);
        score = platform === 3 && (language === 0x409 || encoding === 0) ? 3 : 2;
      } else if (platform === 1 && encoding === 0) {
        text = String.fromCharCode(...bytes);
        score = 1;
      } else continue;
      if (!found[nameId] || found[nameId]!.score < score) found[nameId] = { text, score };
    }
    this.familyName = found[1]?.text ?? '';
    this.subfamilyName = found[2]?.text ?? '';
    this.fullName = found[4]?.text ?? this.familyName;
    this.postscriptName = (found[6]?.text ?? this.fullName).replace(/[^\x21-\x7e]|[[\](){}<>/%]/g, '');
  }

  private parseCmap(): void {
    const r = this.table('cmap')!;
    const base = r.pos;
    r.skip(2);
    const count = r.u16be();
    let best: { offset: number; score: number } | undefined;
    for (let i = 0; i < count; i++) {
      const platform = r.u16be();
      const encoding = r.u16be();
      const offset = r.u32be();
      let score = 0;
      if (platform === 3 && encoding === 10) score = 5;
      else if (platform === 0 && (encoding === 4 || encoding === 6)) score = 4;
      else if (platform === 3 && encoding === 1) score = 3;
      else if (platform === 0) score = 2;
      else if (platform === 3 && encoding === 0) {
        score = 1;
        this.isSymbolic = true;
      }
      if (score && (!best || score > best.score)) best = { offset: base + offset, score };
    }
    if (best && best.score > 1) this.isSymbolic = false;
    if (!best) return;
    const t = new ByteReader(this.data, best.offset);
    const format = t.u16be();
    if (format === 4) {
      t.skip(4);
      const segX2 = t.u16be();
      const segs = segX2 / 2;
      t.skip(6);
      const endsAt = t.pos;
      const startsAt = endsAt + segX2 + 2;
      const deltasAt = startsAt + segX2;
      const rangesAt = deltasAt + segX2;
      const view = new DataView(this.data.buffer, this.data.byteOffset, this.data.byteLength);
      for (let s = 0; s < segs; s++) {
        const end = view.getUint16(endsAt + s * 2);
        const start = view.getUint16(startsAt + s * 2);
        const delta = view.getInt16(deltasAt + s * 2);
        const rangeOffset = view.getUint16(rangesAt + s * 2);
        if (start === 0xffff) break;
        for (let c = start; c <= end; c++) {
          let gid: number;
          if (rangeOffset === 0) gid = (c + delta) & 0xffff;
          else {
            const addr = rangesAt + s * 2 + rangeOffset + (c - start) * 2;
            if (addr + 2 > this.data.length) continue;
            gid = view.getUint16(addr);
            if (gid !== 0) gid = (gid + delta) & 0xffff;
          }
          if (gid !== 0) this.cmap.set(c, gid);
        }
      }
    } else if (format === 12) {
      t.skip(10);
      const groups = t.u32be();
      for (let g = 0; g < groups; g++) {
        const start = t.u32be();
        const end = t.u32be();
        const gid = t.u32be();
        for (let c = start; c <= end && c - start < 0x10000; c++) this.cmap.set(c, gid + (c - start));
      }
    } else if (format === 0) {
      t.skip(4);
      for (let c = 0; c < 256; c++) {
        const gid = t.u8();
        if (gid) this.cmap.set(c, gid);
      }
    } else if (format === 6) {
      t.skip(4);
      const first = t.u16be();
      const n = t.u16be();
      for (let i = 0; i < n; i++) {
        const gid = t.u16be();
        if (gid) this.cmap.set(first + i, gid);
      }
    }
  }

  /** Glyph id for a code point, or 0 (.notdef) when the font lacks it. */
  glyphId(codePoint: number): number {
    const gid = this.cmap.get(codePoint);
    if (gid !== undefined) return gid;
    if (this.isSymbolic && codePoint < 0x100) return this.cmap.get(codePoint | 0xf000) ?? 0;
    return 0;
  }

  hasGlyph(codePoint: number): boolean {
    return this.glyphId(codePoint) !== 0;
  }

  /** Advance width in font units. */
  advance(gid: number): number {
    return this.advances[gid] ?? this.advances[this.advances.length - 1] ?? 0;
  }

  /** Every mapped code point → glyph id (for ToUnicode maps and tests). */
  codePoints(): ReadonlyMap<number, number> {
    return this.cmap;
  }

  /** Kerning from the legacy 'kern' table (format 0), in font units. */
  kerning(left: number, right: number): number {
    if (!this.kernPairs) {
      this.kernPairs = new Map();
      const r = this.table('kern');
      if (r) {
        const version = r.u16be();
        if (version === 0) {
          const nTables = r.u16be();
          for (let t = 0; t < nTables; t++) {
            const start = r.pos;
            r.skip(2);
            const length = r.u16be();
            const coverage = r.u16be();
            if (coverage >> 8 === 0 && (coverage & 1) === 1) {
              const nPairs = r.u16be();
              r.skip(6);
              for (let i = 0; i < nPairs; i++) this.kernPairs.set((r.u16be() << 16) | r.u16be(), r.i16be());
            }
            r.seek(start + length);
          }
        }
      }
    }
    return this.kernPairs.get((left << 16) | right) ?? 0;
  }

  /** Ascent / descent (positive, font units) that Word uses for line height. */
  get lineAscent(): number {
    return this.winAscent || this.ascender;
  }
  get lineDescent(): number {
    return this.winDescent || -this.descender;
  }

  /** Byte range of a TrueType glyph in the glyf table. */
  glyphRange(gid: number): { offset: number; length: number } | undefined {
    const loca = this.tables.get('loca');
    const glyf = this.tables.get('glyf');
    if (!loca || !glyf || gid >= this.numGlyphs) return undefined;
    const view = new DataView(this.data.buffer, this.data.byteOffset, this.data.byteLength);
    let start: number;
    let end: number;
    if (this.indexToLocFormat === 0) {
      start = view.getUint16(loca.offset + gid * 2) * 2;
      end = view.getUint16(loca.offset + gid * 2 + 2) * 2;
    } else {
      start = view.getUint32(loca.offset + gid * 4);
      end = view.getUint32(loca.offset + gid * 4 + 4);
    }
    if (end < start) return undefined;
    return { offset: glyf.offset + start, length: end - start };
  }
}
