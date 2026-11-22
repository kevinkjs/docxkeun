/**
 * OpenType Layout common table formats (shared by GSUB and GPOS): ScriptList / LangSys,
 * FeatureList, LookupList, Coverage, ClassDef and the contextual (5/6 in GSUB, 7/8 in GPOS)
 * subtables. Everything is parsed lazily and cached per table, because a shaping call only
 * touches the few lookups its features select, and fonts like Noto Nastaliq carry hundreds.
 *
 * Reads are bounds-checked and return 0 past the end: a malformed offset then reads as an
 * empty structure instead of throwing in the middle of layout.
 */

/** Big-endian reader over one table's bytes, addressed by absolute offset into the table. */
export class BinaryData {
  readonly view: DataView;
  readonly length: number;
  /** Coverage and ClassDef tables are often shared between subtables: parse each once. */
  readonly coverageCache = new Map<number, Coverage>();
  readonly classDefCache = new Map<number, ClassDef>();

  constructor(readonly bytes: Uint8Array) {
    this.view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
    this.length = bytes.byteLength;
  }

  u16(off: number): number {
    return off >= 0 && off + 2 <= this.length ? this.view.getUint16(off) : 0;
  }

  i16(off: number): number {
    return off >= 0 && off + 2 <= this.length ? this.view.getInt16(off) : 0;
  }

  u32(off: number): number {
    return off >= 0 && off + 4 <= this.length ? this.view.getUint32(off) : 0;
  }

  tag(off: number): string {
    let s = '';
    for (let i = 0; i < 4; i++) s += String.fromCharCode(off + i < this.length ? this.bytes[off + i]! : 0x20);
    return s;
  }

  /** u16 array of `count` items at `off` (clamped to the table). */
  u16s(off: number, count: number): number[] {
    const n = Math.max(0, Math.min(count, Math.floor((this.length - off) / 2)));
    const out = new Array<number>(n);
    for (let i = 0; i < n; i++) out[i] = this.view.getUint16(off + i * 2);
    return out;
  }
}

/** Glyph id → coverage index. */
export type Coverage = Map<number, number>;
/** Glyph id → class (absent means class 0). */
export type ClassDef = Map<number, number>;

const EMPTY_COVERAGE: Coverage = new Map();

export function parseCoverage(d: BinaryData, off: number): Coverage {
  if (!off) return EMPTY_COVERAGE;
  const cached = d.coverageCache.get(off);
  if (cached) return cached;
  const cov: Coverage = new Map();
  const format = d.u16(off);
  if (format === 1) {
    const count = d.u16(off + 2);
    const glyphs = d.u16s(off + 4, count);
    for (let i = 0; i < glyphs.length; i++) if (!cov.has(glyphs[i]!)) cov.set(glyphs[i]!, i);
  } else if (format === 2) {
    const count = d.u16(off + 2);
    for (let i = 0; i < count; i++) {
      const r = off + 4 + i * 6;
      if (r + 6 > d.length) break;
      const start = d.u16(r);
      const end = d.u16(r + 2);
      const startIndex = d.u16(r + 4);
      for (let g = start; g <= end; g++) if (!cov.has(g)) cov.set(g, startIndex + g - start);
    }
  }
  d.coverageCache.set(off, cov);
  return cov;
}

export function parseClassDef(d: BinaryData, off: number): ClassDef {
  if (!off) return new Map();
  const cached = d.classDefCache.get(off);
  if (cached) return cached;
  const cd: ClassDef = new Map();
  const format = d.u16(off);
  if (format === 1) {
    const start = d.u16(off + 2);
    const values = d.u16s(off + 6, d.u16(off + 4));
    for (let i = 0; i < values.length; i++) if (values[i]) cd.set(start + i, values[i]!);
  } else if (format === 2) {
    const count = d.u16(off + 2);
    for (let i = 0; i < count; i++) {
      const r = off + 4 + i * 6;
      if (r + 6 > d.length) break;
      const start = d.u16(r);
      const end = d.u16(r + 2);
      const cls = d.u16(r + 4);
      if (cls) for (let g = start; g <= end; g++) cd.set(g, cls);
    }
  }
  d.classDefCache.set(off, cd);
  return cd;
}

export function classOf(cd: ClassDef, glyph: number): number {
  return cd.get(glyph) ?? 0;
}

/** Lookup flags (OpenType LookupFlag bit enumeration). */
export const LookupFlag = {
  RightToLeft: 0x0001,
  IgnoreBaseGlyphs: 0x0002,
  IgnoreLigatures: 0x0004,
  IgnoreMarks: 0x0008,
  IgnoreFlags: 0x000e,
  UseMarkFilteringSet: 0x0010,
  MarkAttachmentType: 0xff00,
} as const;

export interface LangSys {
  /** Feature index of the required feature, or -1. */
  required: number;
  features: number[];
}

interface ScriptRecord {
  defaultLangSys: LangSys | undefined;
  langSys: Map<string, LangSys>;
}

export interface LookupHeader {
  type: number;
  flag: number;
  /** Mark filtering set index (only meaningful with UseMarkFilteringSet). */
  markFilteringSet: number;
  /** Subtables with extension lookups already resolved to their real type and offset. */
  subtables: Array<{ type: number; offset: number }>;
}

/** Selected script and language system of a GSUB/GPOS table. */
export interface ScriptSelection {
  /** The script tag actually found in the font (may be 'DFLT' / 'latn' fallbacks). */
  scriptTag: string;
  langSys: LangSys;
}

/** ScriptList / FeatureList / LookupList of a GSUB or GPOS table. */
export class LayoutTable {
  readonly data: BinaryData;
  private readonly scripts = new Map<string, ScriptRecord>();
  readonly featureTags: string[] = [];
  private readonly featureOffsets: number[] = [];
  private readonly featureLookupCache = new Map<number, number[]>();
  private readonly lookupOffsets: number[];
  private readonly lookupHeaders = new Map<number, LookupHeader>();

  /** @param extensionType the Extension lookup type of this table (7 for GSUB, 9 for GPOS) */
  constructor(bytes: Uint8Array, private readonly extensionType: number) {
    const d = (this.data = new BinaryData(bytes));
    const scriptList = d.u16(4);
    const featureList = d.u16(6);
    const lookupList = d.u16(8);
    // Version 1.1 adds FeatureVariations (variable fonts); the default features are what we use.

    if (scriptList) {
      const count = d.u16(scriptList);
      for (let i = 0; i < count; i++) {
        const rec = scriptList + 2 + i * 6;
        const tag = d.tag(rec);
        const off = scriptList + d.u16(rec + 4);
        const defOff = d.u16(off);
        const record: ScriptRecord = { defaultLangSys: defOff ? this.parseLangSys(off + defOff) : undefined, langSys: new Map() };
        const langCount = d.u16(off + 2);
        for (let j = 0; j < langCount; j++) {
          const lrec = off + 4 + j * 6;
          record.langSys.set(d.tag(lrec), this.parseLangSys(off + d.u16(lrec + 4)));
        }
        if (!this.scripts.has(tag)) this.scripts.set(tag, record);
      }
    }
    if (featureList) {
      const count = d.u16(featureList);
      for (let i = 0; i < count; i++) {
        const rec = featureList + 2 + i * 6;
        this.featureTags.push(d.tag(rec));
        this.featureOffsets.push(featureList + d.u16(rec + 4));
      }
    }
    this.lookupOffsets = lookupList ? d.u16s(lookupList + 2, d.u16(lookupList)).map((o) => lookupList + o) : [];
  }

  private parseLangSys(off: number): LangSys {
    const d = this.data;
    const req = d.u16(off + 2);
    return { required: req === 0xffff ? -1 : req, features: d.u16s(off + 6, d.u16(off + 4)) };
  }

  get lookupCount(): number {
    return this.lookupOffsets.length;
  }

  hasScript(tag: string): boolean {
    return this.scripts.has(tag);
  }

  /**
   * Picks the first candidate script the font has, then falls back like HarfBuzz to 'DFLT',
   * 'dflt' and finally 'latn' (old fonts park their features there even for other scripts).
   * Within the script, the requested language system wins, else the default one.
   */
  select(scriptCandidates: readonly string[], language?: string): ScriptSelection | undefined {
    let tag: string | undefined;
    for (const t of [...scriptCandidates, 'DFLT', 'dflt', 'latn']) {
      if (this.scripts.has(t)) {
        tag = t;
        break;
      }
    }
    if (!tag) return undefined;
    const script = this.scripts.get(tag)!;
    const langSys = (language && script.langSys.get(language)) || script.defaultLangSys || { required: -1, features: [] };
    return { scriptTag: tag, langSys };
  }

  /** Index of the feature with `tag` in the language system, or -1. */
  findFeature(langSys: LangSys, tag: string): number {
    for (const fi of langSys.features) if (this.featureTags[fi] === tag) return fi;
    return -1;
  }

  featureLookups(featureIndex: number): number[] {
    const cached = this.featureLookupCache.get(featureIndex);
    if (cached) return cached;
    const off = this.featureOffsets[featureIndex];
    const list = off === undefined ? [] : this.data.u16s(off + 4, this.data.u16(off + 2)).filter((l) => l < this.lookupCount);
    this.featureLookupCache.set(featureIndex, list);
    return list;
  }

  lookupHeader(index: number): LookupHeader | undefined {
    const cached = this.lookupHeaders.get(index);
    if (cached) return cached;
    const off = this.lookupOffsets[index];
    if (off === undefined) return undefined;
    const d = this.data;
    const type = d.u16(off);
    const flag = d.u16(off + 2);
    const count = d.u16(off + 4);
    const subtables: LookupHeader['subtables'] = [];
    for (const rel of d.u16s(off + 6, count)) {
      const st = off + rel;
      if (type === this.extensionType) {
        // ExtensionSubstFormat1 / ExtensionPosFormat1: format, real type, 32-bit offset.
        if (d.u16(st) !== 1) continue;
        const realType = d.u16(st + 2);
        if (realType === this.extensionType) continue; // an extension may not point at another one
        subtables.push({ type: realType, offset: st + d.u32(st + 4) });
      } else subtables.push({ type, offset: st });
    }
    const header: LookupHeader = {
      type: subtables[0]?.type ?? type,
      flag,
      markFilteringSet: flag & LookupFlag.UseMarkFilteringSet ? d.u16(off + 6 + count * 2) : 0,
      subtables,
    };
    this.lookupHeaders.set(index, header);
    return header;
  }
}

// ---------------------------------------------------------------------------------------
// Contextual subtables (GSUB 5/6, GPOS 7/8): the three formats of both, unified as
// "chain rules" whose backtrack/lookahead are empty for the non-chaining variants.

export interface LookupRecord {
  sequenceIndex: number;
  lookupIndex: number;
}

export interface ChainRule {
  /** Backtrack values, nearest glyph first (as stored in the font). */
  backtrack: number[];
  /** Input values after the first glyph (the first is matched by coverage / class). */
  input: number[];
  lookahead: number[];
  lookups: LookupRecord[];
}

export type ContextualSubtable =
  | { kind: 'ctx-glyphs'; coverage: Coverage; ruleSets: Array<ChainRule[] | undefined> }
  | {
      kind: 'ctx-classes';
      coverage: Coverage;
      backtrackClasses: ClassDef;
      inputClasses: ClassDef;
      lookaheadClasses: ClassDef;
      ruleSets: Array<ChainRule[] | undefined>;
    }
  | { kind: 'ctx-coverage'; backtrack: Coverage[]; input: Coverage[]; lookahead: Coverage[]; lookups: LookupRecord[] };

function lookupRecords(d: BinaryData, off: number, count: number): LookupRecord[] {
  const out: LookupRecord[] = [];
  for (let i = 0; i < count; i++) out.push({ sequenceIndex: d.u16(off + i * 4), lookupIndex: d.u16(off + i * 4 + 2) });
  return out;
}

function parseRuleSets(d: BinaryData, off: number, countAt: number, chained: boolean): Array<ChainRule[] | undefined> {
  const sets: Array<ChainRule[] | undefined> = [];
  const count = d.u16(countAt);
  for (let i = 0; i < count; i++) {
    const rel = d.u16(countAt + 2 + i * 2);
    if (!rel) {
      sets.push(undefined);
      continue;
    }
    const setOff = off + rel;
    const rules: ChainRule[] = [];
    const ruleCount = d.u16(setOff);
    for (let j = 0; j < ruleCount; j++) {
      const r = setOff + d.u16(setOff + 2 + j * 2);
      if (chained) {
        const bc = d.u16(r);
        const backtrack = d.u16s(r + 2, bc);
        let p = r + 2 + bc * 2;
        const ic = d.u16(p);
        const input = d.u16s(p + 2, Math.max(0, ic - 1));
        p += 2 + Math.max(0, ic - 1) * 2;
        const lc = d.u16(p);
        const lookahead = d.u16s(p + 2, lc);
        p += 2 + lc * 2;
        if (ic === 0) continue;
        rules.push({ backtrack, input, lookahead, lookups: lookupRecords(d, p + 2, d.u16(p)) });
      } else {
        const gc = d.u16(r);
        const lc = d.u16(r + 2);
        if (gc === 0) continue;
        const input = d.u16s(r + 4, gc - 1);
        rules.push({ backtrack: [], input, lookahead: [], lookups: lookupRecords(d, r + 4 + (gc - 1) * 2, lc) });
      }
    }
    sets.push(rules);
  }
  return sets;
}

/** Parses a Context (chained = false) or ChainContext subtable of any format. */
export function parseContextual(d: BinaryData, off: number, chained: boolean): ContextualSubtable | undefined {
  const format = d.u16(off);
  if (format === 1) {
    return { kind: 'ctx-glyphs', coverage: parseCoverage(d, off + d.u16(off + 2)), ruleSets: parseRuleSets(d, off, off + 4, chained) };
  }
  if (format === 2) {
    const coverage = parseCoverage(d, off + d.u16(off + 2));
    if (chained) {
      const cd = (at: number) => parseClassDef(d, d.u16(at) ? off + d.u16(at) : 0);
      return {
        kind: 'ctx-classes',
        coverage,
        backtrackClasses: cd(off + 4),
        inputClasses: cd(off + 6),
        lookaheadClasses: cd(off + 8),
        ruleSets: parseRuleSets(d, off, off + 10, true),
      };
    }
    const classes = parseClassDef(d, off + d.u16(off + 4));
    return { kind: 'ctx-classes', coverage, backtrackClasses: classes, inputClasses: classes, lookaheadClasses: classes, ruleSets: parseRuleSets(d, off, off + 6, false) };
  }
  if (format === 3) {
    const covs = (at: number, count: number) => d.u16s(at, count).map((o) => parseCoverage(d, off + o));
    if (chained) {
      let p = off + 2;
      const bc = d.u16(p);
      const backtrack = covs(p + 2, bc);
      p += 2 + bc * 2;
      const ic = d.u16(p);
      const input = covs(p + 2, ic);
      p += 2 + ic * 2;
      const lc = d.u16(p);
      const lookahead = covs(p + 2, lc);
      p += 2 + lc * 2;
      if (!ic) return undefined;
      return { kind: 'ctx-coverage', backtrack, input, lookahead, lookups: lookupRecords(d, p + 2, d.u16(p)) };
    }
    const gc = d.u16(off + 2);
    const lc = d.u16(off + 4);
    if (!gc) return undefined;
    return { kind: 'ctx-coverage', backtrack: [], input: covs(off + 6, gc), lookahead: [], lookups: lookupRecords(d, off + 6 + gc * 2, lc) };
  }
  return undefined;
}

/** Normalises a user-supplied OpenType tag ('nko' → 'nko ', 'URD' → 'URD '). */
export function normalizeTag(tag: string): string {
  return tag.length >= 4 ? tag.slice(0, 4) : tag.padEnd(4, ' ');
}
