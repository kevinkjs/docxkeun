/**
 * Indic shaper (Devanagari, Bengali, Gurmukhi, Gujarati, Oriya, Tamil, Telugu, Kannada,
 * Malayalam), following the OpenType Indic shaping spec as implemented by HarfBuzz /
 * Uniscribe:
 *
 * 1. Characters get a syllabic category (consonant, vowel, matra, nukta, halant, ...) and a
 *    position class; the text is split into syllables with the Indic syllable grammar.
 * 2. Initial reordering per syllable: find the base consonant, detect a reph (Ra+Halant at
 *    the start when the font's 'rphf' forms one), move pre-base matras before the base, and
 *    set per-glyph masks so rphf/half/blwf/pstf/pref apply only where the spec says.
 * 3. The basic features run one stage at a time: nukt akhn rphf rkrf pref blwf abvf half
 *    pstf vatu cjct.
 * 4. Final reordering: the pre-base matra moves after any half forms, the reph moves to its
 *    script-specific position, a pre-base-reordering consonant (Malayalam Ra) moves before
 *    the base.
 * 5. The presentation features init pres abvs blws psts haln (+ calt...), then GPOS.
 *
 * Every glyph of a syllable ends up in one cluster (as in Uniscribe), which keeps the
 * reordering invisible to cluster-based layout.
 */
import { isLigated, isMultiplied, isSubstituted, ligatedAndDidntMultiply, mergeClusters, type GlyphInfo } from './buffer.js';
import { GlyphProps } from './gdef.js';
import type { PlanBuilder, ShapeContext } from './plan.js';
import type { Shaper } from './shaper.js';

// Syllabic categories (values match HarfBuzz's for easy cross-checking).
const X = 0;
const C = 1;
const V = 2;
const N = 3;
const H = 4;
const ZWNJ = 5;
const ZWJ = 6;
const M = 7;
const SM = 8;
const A = 9;
const PLACEHOLDER = 10;
const DOTTEDCIRCLE = 11;
const RS = 12;
const Repha = 14;
const Ra = 15;
const CM = 16;
const Symbol = 17;
const CS = 18;

// Positions, in final visual order within a syllable.
const POS_START = 0;
const POS_RA_TO_BECOME_REPH = 1;
const POS_PRE_M = 2;
const POS_PRE_C = 3;
const POS_BASE_C = 4;
const POS_AFTER_MAIN = 5;
const POS_ABOVE_C = 6;
const POS_BEFORE_SUB = 7;
const POS_BELOW_C = 8;
const POS_AFTER_SUB = 9;
const POS_BEFORE_POST = 10;
const POS_POST_C = 11;
const POS_AFTER_POST = 12;
const POS_FINAL_C = 13;
const POS_SMVD = 14;
const POS_END = 15;

const bit = (n: number): number => 1 << n;
const CONSONANT_FLAGS = bit(C) | bit(CS) | bit(Ra) | bit(CM) | bit(V) | bit(PLACEHOLDER) | bit(DOTTEDCIRCLE);
const JOINER_FLAGS = bit(ZWJ) | bit(ZWNJ);

/**
 * IndicSyllabicCategory / IndicPositionalCategory for the nine scripts, as
 * "first[-last] category [side]". Sides: L left, R right, T top, B bottom and their
 * combinations (two-part matras); only matras use them.
 */
const INDIC_DATA = `
00A0 PH;200C ZWNJ;200D ZWJ;2010-2014 PH;25CC DC
0900-0902 SM T;0903 SM R;0904-0914 V;0915-0939 C;093A M T;093B M R;093C N B;093D Sym;093E M R;093F M L;0940 M R
0941-0944 M B;0945-0948 M T;0949-094C M R;094D H B;094E M L;094F M R;0951 A T;0952 A B;0953-0954 SM T;0955 M T
0956-0957 M B;0958-095F C;0960-0961 V;0962-0963 M B;0966-096F PH;0972-0977 V;0978-097F C
0980 PH;0981 SM T;0982-0983 SM R;0985-098C V;098F-0990 V;0993-0994 V;0995-09A8 C;09AA-09B0 C;09B2 C;09B6-09B9 C
09BC N B;09BD Sym;09BE M R;09BF M L;09C0 M R;09C1-09C4 M B;09C7-09C8 M L;09CB-09CC M LR;09CD H B;09CE C;09D7 M R
09DC-09DD C;09DF C;09E0-09E1 V;09E2-09E3 M B;09E6-09EF PH;09F0-09F1 C;09FC PH;09FE SM T
0A01-0A02 SM T;0A03 SM R;0A05-0A0A V;0A0F-0A10 V;0A13-0A14 V;0A15-0A28 C;0A2A-0A30 C;0A32-0A33 C;0A35-0A36 C
0A38-0A39 C;0A3C N B;0A3E M R;0A3F M L;0A40 M R;0A41-0A42 M B;0A47-0A48 M T;0A4B-0A4C M T;0A4D H B;0A51 M B
0A59-0A5C C;0A5E C;0A66-0A6F PH;0A70-0A71 SM T;0A72-0A73 C;0A75 CM B
0A81-0A82 SM T;0A83 SM R;0A85-0A8D V;0A8F-0A91 V;0A93-0A94 V;0A95-0AA8 C;0AAA-0AB0 C;0AB2-0AB3 C;0AB5-0AB9 C
0ABC N B;0ABD Sym;0ABE M R;0ABF M L;0AC0 M R;0AC1-0AC4 M B;0AC5 M T;0AC7-0AC8 M T;0AC9 M TR;0ACB-0ACC M R;0ACD H B
0AE0-0AE1 V;0AE2-0AE3 M B;0AE6-0AEF PH;0AF9 C;0AFA A T;0AFB N T;0AFC A T;0AFD-0AFF N T
0B01 SM T;0B02-0B03 SM R;0B05-0B0C V;0B0F-0B10 V;0B13-0B14 V;0B15-0B28 C;0B2A-0B30 C;0B32-0B33 C;0B35-0B39 C
0B3C N B;0B3D Sym;0B3E M R;0B3F M T;0B40 M R;0B41-0B44 M B;0B47 M L;0B48 M TL;0B4B M LR;0B4C M TLR;0B4D H B
0B55 N T;0B56 M T;0B57 M TR;0B5C-0B5D C;0B5F C;0B60-0B61 V;0B62-0B63 M B;0B66-0B6F PH;0B71 C
0B82 SM T;0B85-0B8A V;0B8E-0B90 V;0B92-0B94 V;0B95 C;0B99-0B9A C;0B9C C;0B9E-0B9F C;0BA3-0BA4 C;0BA8-0BAA C
0BAE-0BB9 C;0BBE-0BBF M R;0BC0 M T;0BC1-0BC2 M R;0BC6-0BC8 M L;0BCA-0BCC M LR;0BCD H T;0BD7 M R;0BE6-0BEF PH
0C00 SM T;0C01-0C03 SM R;0C04 SM T;0C05-0C0C V;0C0E-0C10 V;0C12-0C14 V;0C15-0C28 C;0C2A-0C39 C;0C3C N B;0C3D Sym
0C3E-0C40 M T;0C41-0C44 M R;0C46-0C47 M T;0C48 M TB;0C4A-0C4C M T;0C4D H T;0C55 M T;0C56 M B;0C58-0C5A C;0C5D C
0C60-0C61 V;0C62-0C63 M B;0C66-0C6F PH
0C80 PH;0C81 SM T;0C82-0C83 SM R;0C85-0C8C V;0C8E-0C90 V;0C92-0C94 V;0C95-0CA8 C;0CAA-0CB3 C;0CB5-0CB9 C;0CBC N B
0CBD Sym;0CBE M R;0CBF M T;0CC0 M TR;0CC1-0CC4 M R;0CC6 M T;0CC7-0CC8 M TR;0CCA-0CCB M TR;0CCC M T;0CCD H T
0CD5-0CD6 M R;0CDD-0CDE C;0CE0-0CE1 V;0CE2-0CE3 M B;0CE6-0CEF PH;0CF1-0CF2 CS;0CF3 SM R
0D00-0D01 SM T;0D02-0D03 SM R;0D05-0D0C V;0D0E-0D10 V;0D12-0D14 V;0D15-0D3A C;0D3B-0D3C M T;0D3D Sym;0D3E-0D42 M R
0D43-0D44 M B;0D46-0D48 M L;0D4A-0D4C M LR;0D4D H T;0D4E Repha;0D54-0D56 C;0D57 M R;0D5F-0D61 V;0D62-0D63 M B
0D66-0D6F PH;0D7A-0D7F C
1CD0-1CD2 A T;1CD4-1CE8 A;1CE9-1CEC Sym;1CED A B;1CEE-1CF1 Sym;1CF2-1CF3 SM R;1CF4 A T;1CF5-1CF6 C;1CF8-1CF9 A T
A8E0-A8F1 A T;A8F2-A8F7 Sym;A8FF M T
`;

const CATEGORY_CODES: Record<string, number> = {
  X, C, V, N, H, M, SM, A, PH: PLACEHOLDER, DC: DOTTEDCIRCLE, Sym: Symbol, Repha, CM, CS, ZWJ, ZWNJ,
};
const SIDE_CODES: Record<string, number> = {
  L: POS_PRE_C, R: POS_POST_C, T: POS_ABOVE_C, B: POS_BELOW_C,
  // Two- and three-part matras resolve to the position of their last part.
  LR: POS_POST_C, TR: POS_POST_C, TL: POS_ABOVE_C, TLR: POS_POST_C, TB: POS_BELOW_C, BR: POS_POST_C,
};
const RA = new Set([0x0930, 0x09b0, 0x09f0, 0x0a30, 0x0ab0, 0x0b30, 0x0bb0, 0x0c30, 0x0cb0, 0x0d30]);

const TABLE = new Map<number, number>();
{
  for (const entry of INDIC_DATA.split(/[;\n]/)) {
    const m = /^\s*([0-9A-F]{4})(?:-([0-9A-F]{4}))?\s+(\w+)(?:\s+(\w+))?\s*$/.exec(entry);
    if (!m) continue;
    const first = parseInt(m[1]!, 16);
    const last = m[2] ? parseInt(m[2], 16) : first;
    const cat = CATEGORY_CODES[m[3]!] ?? X;
    const side = m[4] ? (SIDE_CODES[m[4]] ?? POS_END) : POS_END;
    for (let cp = first; cp <= last; cp++) TABLE.set(cp, cat | (side << 8));
  }
}

function blockScript(cp: number): string {
  if ((cp >= 0x0900 && cp <= 0x097f) || (cp >= 0x1cd0 && cp <= 0x1cff) || (cp >= 0xa8e0 && cp <= 0xa8ff)) return 'deva';
  if (cp >= 0x0980 && cp <= 0x09ff) return 'beng';
  if (cp >= 0x0a00 && cp <= 0x0a7f) return 'guru';
  if (cp >= 0x0a80 && cp <= 0x0aff) return 'gujr';
  if (cp >= 0x0b00 && cp <= 0x0b7f) return 'orya';
  if (cp >= 0x0b80 && cp <= 0x0bff) return 'taml';
  if (cp >= 0x0c00 && cp <= 0x0c7f) return 'telu';
  if (cp >= 0x0c80 && cp <= 0x0cff) return 'knda';
  if (cp >= 0x0d00 && cp <= 0x0d7f) return 'mlym';
  return '';
}

/** Where a matra on a given side sits relative to the consonants, per script. */
function matraPosition(cp: number, side: number): number {
  const s = blockScript(cp);
  switch (side) {
    case POS_PRE_C:
      return POS_PRE_M;
    case POS_POST_C:
      if (s === 'deva') return POS_AFTER_SUB;
      if (s === 'telu') return cp <= 0x0c42 ? POS_BEFORE_SUB : POS_AFTER_SUB;
      if (s === 'knda') return cp < 0x0cc3 || cp > 0x0cd6 ? POS_BEFORE_SUB : POS_AFTER_SUB;
      if (s === 'beng' || s === 'guru' || s === 'gujr' || s === 'orya' || s === 'taml' || s === 'mlym') return POS_AFTER_POST;
      return POS_AFTER_SUB;
    case POS_ABOVE_C:
      if (s === 'guru') return POS_AFTER_POST; // deviates from the spec, like Uniscribe
      if (s === 'orya') return POS_AFTER_MAIN;
      if (s === 'telu' || s === 'knda') return POS_BEFORE_SUB;
      return POS_AFTER_SUB;
    case POS_BELOW_C:
      if (s === 'guru' || s === 'gujr' || s === 'taml' || s === 'mlym') return POS_AFTER_POST;
      if (s === 'telu' || s === 'knda') return POS_BEFORE_SUB;
      return POS_AFTER_SUB;
    default:
      return side;
  }
}

/** [category, position] of a code point. */
export function indicProperties(cp: number): [number, number] {
  const entry = TABLE.get(cp);
  let cat = entry === undefined ? X : entry & 0xff;
  const side = entry === undefined ? POS_END : entry >> 8;
  let pos = side;
  if (bit(cat) & CONSONANT_FLAGS) {
    pos = POS_BASE_C;
    if (RA.has(cp)) cat = Ra;
  } else if (cat === M) pos = matraPosition(cp, side);
  else if (cat === SM || cat === A || cat === Symbol) pos = POS_SMVD;
  if (cp === 0x0b01) pos = POS_BEFORE_SUB; // Oriya candrabindu
  return [cat, pos];
}

/**
 * Characters HarfBuzz always decomposes for Indic: two-part matras (so each part can be
 * reordered) and nukta consonants excluded from composition (the font's 'nukt' forms them).
 */
const ALWAYS_DECOMPOSE = new Set([
  0x09cb, 0x09cc, 0x0b48, 0x0b4b, 0x0b4c, 0x0bca, 0x0bcb, 0x0bcc, 0x0c48, 0x0cc0, 0x0cc7, 0x0cc8, 0x0cca, 0x0ccb, 0x0d4a,
  0x0d4b, 0x0d4c, 0x0958, 0x0959, 0x095a, 0x095b, 0x095c, 0x095d, 0x095e, 0x095f, 0x09df, 0x0a33, 0x0a36, 0x0a59, 0x0a5a,
  0x0a5b, 0x0a5e, 0x0b5c, 0x0b5d,
]);

// --- Syllables --------------------------------------------------------------------------

const SYL_CONSONANT = 0;
const SYL_VOWEL = 1;
const SYL_STANDALONE = 2;
const SYL_SYMBOL = 3;
const SYL_BROKEN = 4;
const SYL_OTHER = 5;

/** A pattern maps a start index to every index where a match can end (a tiny NFA). */
type Pattern = (cats: readonly number[], i: number) => number[];

const cat =
  (...cs: number[]): Pattern =>
  (s, i) =>
    i < s.length && cs.includes(s[i]!) ? [i + 1] : [];
const seq =
  (...ps: Pattern[]): Pattern =>
  (s, i) => {
    let cur = [i];
    for (const p of ps) {
      const next = new Set<number>();
      for (const c of cur) for (const e of p(s, c)) next.add(e);
      cur = [...next];
      if (!cur.length) break;
    }
    return cur;
  };
const alt =
  (...ps: Pattern[]): Pattern =>
  (s, i) => [...new Set(ps.flatMap((p) => p(s, i)))];
const opt =
  (p: Pattern): Pattern =>
  (s, i) => [...new Set([i, ...p(s, i)])];
const star =
  (p: Pattern): Pattern =>
  (s, i) => {
    const seen = new Set([i]);
    let frontier = [i];
    while (frontier.length) {
      const next: number[] = [];
      for (const f of frontier) for (const e of p(s, f)) if (!seen.has(e)) (seen.add(e), next.push(e));
      frontier = next;
    }
    return [...seen];
  };

// The Indic syllable grammar (HarfBuzz's hb-ot-shaper-indic-machine.rl).
const c = cat(C, Ra);
const n = seq(opt(seq(opt(cat(ZWNJ)), cat(RS))), opt(seq(cat(N), opt(cat(N)))));
const z = cat(ZWJ, ZWNJ);
const reph = alt(seq(cat(Ra), cat(H)), cat(Repha));
const cn = seq(c, opt(cat(ZWJ)), n);
const symbol = seq(cat(Symbol), opt(cat(N)));
const matraGroup = seq(star(z), cat(M), opt(cat(N)), opt(cat(H)));
const syllableTail = seq(opt(seq(opt(z), cat(SM), opt(cat(SM)), opt(cat(ZWNJ)))), star(cat(A)));
const halantGroup = seq(opt(z), cat(H), opt(seq(cat(ZWJ), opt(cat(N)))));
const finalHalantGroup = alt(halantGroup, seq(cat(H), cat(ZWNJ)));
const medialGroup = opt(cat(CM));
const halantOrMatraGroup = alt(finalHalantGroup, star(matraGroup));
const complexSyllableTail = seq(star(seq(halantGroup, cn)), medialGroup, halantOrMatraGroup, syllableTail);
const SYLLABLES: Array<[number, Pattern]> = [
  [SYL_CONSONANT, seq(opt(cat(Repha, CS)), cn, complexSyllableTail)],
  [SYL_VOWEL, seq(opt(reph), cat(V), n, alt(cat(ZWJ), complexSyllableTail))],
  [SYL_STANDALONE, seq(alt(seq(opt(cat(Repha, CS)), cat(PLACEHOLDER)), seq(opt(reph), cat(DOTTEDCIRCLE))), n, complexSyllableTail)],
  [SYL_SYMBOL, seq(symbol, syllableTail)],
  [SYL_BROKEN, seq(opt(reph), n, complexSyllableTail)],
];

/** Splits categories into syllables: [start, end, type][], longest match first. */
export function findSyllables(cats: readonly number[]): Array<[number, number, number]> {
  const out: Array<[number, number, number]> = [];
  let i = 0;
  while (i < cats.length) {
    let bestEnd = i + 1;
    let bestType = SYL_OTHER;
    let bestLen = 0;
    for (const [type, p] of SYLLABLES) {
      const end = Math.max(...p(cats, i));
      if (end - i > bestLen) {
        bestLen = end - i;
        bestEnd = end;
        bestType = type;
      }
    }
    out.push([i, bestEnd, bestType]);
    i = bestEnd;
  }
  return out;
}

// --- Per-script configuration -----------------------------------------------------------

interface IndicConfig {
  hasOldSpec: boolean;
  virama: number;
  rephPos: number;
  rephMode: 'implicit' | 'explicit' | 'log-repha';
  blwfMode: 'pre-and-post' | 'post-only';
}

const CONFIGS: Record<string, IndicConfig> = {
  deva: { hasOldSpec: true, virama: 0x094d, rephPos: POS_BEFORE_POST, rephMode: 'implicit', blwfMode: 'pre-and-post' },
  beng: { hasOldSpec: true, virama: 0x09cd, rephPos: POS_AFTER_SUB, rephMode: 'implicit', blwfMode: 'pre-and-post' },
  guru: { hasOldSpec: true, virama: 0x0a4d, rephPos: POS_BEFORE_SUB, rephMode: 'implicit', blwfMode: 'pre-and-post' },
  gujr: { hasOldSpec: true, virama: 0x0acd, rephPos: POS_BEFORE_POST, rephMode: 'implicit', blwfMode: 'pre-and-post' },
  orya: { hasOldSpec: true, virama: 0x0b4d, rephPos: POS_AFTER_MAIN, rephMode: 'implicit', blwfMode: 'pre-and-post' },
  taml: { hasOldSpec: true, virama: 0x0bcd, rephPos: POS_AFTER_POST, rephMode: 'implicit', blwfMode: 'pre-and-post' },
  telu: { hasOldSpec: true, virama: 0x0c4d, rephPos: POS_AFTER_POST, rephMode: 'explicit', blwfMode: 'post-only' },
  knda: { hasOldSpec: true, virama: 0x0ccd, rephPos: POS_AFTER_POST, rephMode: 'implicit', blwfMode: 'post-only' },
  mlym: { hasOldSpec: true, virama: 0x0d4d, rephPos: POS_AFTER_MAIN, rephMode: 'log-repha', blwfMode: 'pre-and-post' },
};
const DEFAULT_CONFIG: IndicConfig = { hasOldSpec: false, virama: 0, rephPos: POS_BEFORE_POST, rephMode: 'implicit', blwfMode: 'pre-and-post' };

interface IndicPlanData {
  script: string;
  config: IndicConfig;
  /** The font only has the old 'deva'-style tables (different halant/base rules). */
  isOldSpec: boolean;
  zeroContext: boolean;
  viramaGlyph: number;
  consonantPositions: Map<number, number>;
}

function planData(ctx: ShapeContext): IndicPlanData {
  let d = ctx.plan.shaperData as IndicPlanData | undefined;
  if (!d) {
    const script = ctx.plan.script;
    const config = CONFIGS[script] ?? DEFAULT_CONFIG;
    const isOldSpec = config.hasOldSpec && !(ctx.plan.gsubScript ?? '').endsWith('2');
    d = {
      script,
      config,
      isOldSpec,
      zeroContext: !isOldSpec && script !== 'mlym',
      viramaGlyph: config.virama ? ctx.font.glyphId(config.virama) : 0,
      consonantPositions: new Map(),
    };
    ctx.plan.shaperData = d;
  }
  return d;
}

function wouldSubstitute(ctx: ShapeContext, d: IndicPlanData, tag: string, glyphs: number[]): boolean {
  const gsub = ctx.plan.gsub;
  return !!gsub && gsub.wouldSubstitute(ctx.plan.gsubLookups(tag), glyphs, d.zeroContext);
}

const isOneOf = (info: GlyphInfo, flags: number): boolean => !isLigated(info) && (bit(info.cat) & flags) !== 0; // a ligature has no category any more
const isConsonant = (info: GlyphInfo): boolean => isOneOf(info, CONSONANT_FLAGS);
const isJoiner = (info: GlyphInfo): boolean => isOneOf(info, JOINER_FLAGS);
const isHalant = (info: GlyphInfo): boolean => isOneOf(info, bit(H));

function reverseRange(buf: GlyphInfo[], start: number, end: number): void {
  for (let i = start, j = end - 1; i < j; i++, j--) [buf[i], buf[j]] = [buf[j]!, buf[i]!];
}

/** Moves the glyph at `from` to index `to` (as it reads after removal), shifting the rest. */
function moveGlyph(buf: GlyphInfo[], from: number, to: number): void {
  const [g] = buf.splice(from, 1);
  buf.splice(to, 0, g!);
}

function forEachSyllable(buf: GlyphInfo[], fn: (start: number, end: number) => void): void {
  let start = 0;
  while (start < buf.length) {
    let end = start + 1;
    while (end < buf.length && buf[end]!.syllable === buf[start]!.syllable) end++;
    fn(start, end);
    start = end;
  }
}

// --- Shaping stages ---------------------------------------------------------------------

function setupMasks(ctx: ShapeContext): void {
  for (const info of ctx.buffer) [info.cat, info.pos] = indicProperties(info.codepoint);
}

function setupSyllables(ctx: ShapeContext): void {
  const buf = ctx.buffer;
  const syllables = findSyllables(buf.map((g) => g.cat));
  let serial = 1;
  for (const [start, end, type] of syllables) {
    for (let i = start; i < end; i++) buf[i]!.syllable = serial * 16 + type;
    mergeClusters(buf, start, end);
    serial++;
  }
}

/** Below-base / post-base consonants are those the font forms with a halant (blwf, pstf...). */
function consonantPosition(ctx: ShapeContext, d: IndicPlanData, consonant: number): number {
  const cached = d.consonantPositions.get(consonant);
  if (cached !== undefined) return cached;
  // New-spec tables match Halant,Consonant; old-spec (and fonts copying them) Consonant,Halant.
  const g1 = [d.viramaGlyph, consonant];
  const g2 = [consonant, d.viramaGlyph];
  const ws = (tag: string) => wouldSubstitute(ctx, d, tag, g1) || wouldSubstitute(ctx, d, tag, g2);
  let pos = POS_BASE_C;
  if (ws('blwf') || ws('vatu')) pos = POS_BELOW_C;
  else if (ws('pstf') || ws('pref')) pos = POS_POST_C;
  d.consonantPositions.set(consonant, pos);
  return pos;
}

function initialReordering(ctx: ShapeContext): void {
  const d = planData(ctx);
  const buf = ctx.buffer;
  if (d.viramaGlyph) {
    for (const info of buf) if (info.pos === POS_BASE_C) info.pos = consonantPosition(ctx, d, info.glyph);
  }
  forEachSyllable(buf, (start, end) => {
    const type = buf[start]!.syllable & 15;
    if (type === SYL_CONSONANT || type === SYL_VOWEL || type === SYL_STANDALONE || type === SYL_BROKEN) initialReorderSyllable(ctx, d, start, end);
  });
}

function initialReorderSyllable(ctx: ShapeContext, d: IndicPlanData, start: number, end: number): void {
  const buf = ctx.buffer;
  const plan = ctx.plan;
  const g = (i: number) => buf[i]!;

  // Kannada: Ra,Halant,ZWJ behaves like Ra,ZWJ,Halant (legacy encoding of the eyelash form).
  if (d.script === 'knda' && start + 3 <= end && isOneOf(g(start), bit(Ra)) && isOneOf(g(start + 1), bit(H)) && isOneOf(g(start + 2), bit(ZWJ))) {
    [buf[start + 1], buf[start + 2]] = [g(start + 2), g(start + 1)];
  }

  // 1. Find the base consonant, excluding an initial Ra+Halant that forms a reph.
  let base = end;
  let hasReph = false;
  let limit = start;
  const rphf = plan.maskOf('rphf');
  const mode = d.config.rephMode;
  if (rphf && start + 3 <= end && ((mode === 'implicit' && !isJoiner(g(start + 2))) || (mode === 'explicit' && g(start + 2).cat === ZWJ))) {
    const glyphs = [g(start).glyph, g(start + 1).glyph];
    if (wouldSubstitute(ctx, d, 'rphf', glyphs) || (mode === 'explicit' && wouldSubstitute(ctx, d, 'rphf', [...glyphs, g(start + 2).glyph]))) {
      limit += 2;
      while (limit < end && isJoiner(g(limit))) limit++;
      base = start;
      hasReph = true;
    }
  } else if (mode === 'log-repha' && g(start).cat === Repha) {
    limit += 1;
    while (limit < end && isJoiner(g(limit))) limit++;
    base = start;
    hasReph = true;
  }

  // Walk back from the end to the first consonant without a below-/post-base form.
  {
    let i = end;
    let seenBelow = false;
    do {
      i--;
      const info = g(i);
      if (isConsonant(info)) {
        if (info.pos !== POS_BELOW_C && (info.pos !== POS_POST_C || seenBelow)) {
          base = i;
          break;
        }
        if (info.pos === POS_BELOW_C) seenBelow = true;
        base = i;
      } else if (start < i && info.cat === ZWJ && g(i - 1).cat === H) {
        // Halant,ZWJ asks for an explicit half form: the base is after it.
        break;
      }
    } while (i > limit);
  }
  // A lone Ra+Halant (no other consonant) doesn't form a reph; Ra is the base.
  if (hasReph && base === start && limit - base <= 2) hasReph = false;

  // 2. Assign final positions.
  for (let i = start; i < base; i++) g(i).pos = Math.min(POS_PRE_C, g(i).pos);
  if (base < end) g(base).pos = POS_BASE_C;
  // A consonant after a matra is a final consonant.
  for (let i = base + 1; i < end; i++) {
    if (g(i).cat === M) {
      for (let j = i + 1; j < end; j++) {
        if (isConsonant(g(j))) {
          g(j).pos = POS_FINAL_C;
          break;
        }
      }
      break;
    }
  }
  if (hasReph) g(start).pos = POS_RA_TO_BECOME_REPH;

  if (d.isOldSpec) {
    // Old spec: the first post-base halant moves after the last consonant.
    const noDoubleHalant = d.script === 'knda';
    for (let i = base + 1; i < end; i++) {
      if (g(i).cat === H) {
        let j = end - 1;
        for (; j > i; j--) if (isConsonant(g(j)) || (noDoubleHalant && g(j).cat === H)) break;
        if (g(j).cat !== H && j > i) moveGlyph(buf, i, j);
        break;
      }
    }
  }

  // Joiners, nuktas, halants and medials travel with the character before them.
  {
    let lastPos = POS_START;
    for (let i = start; i < end; i++) {
      const info = g(i);
      if (bit(info.cat) & (JOINER_FLAGS | bit(N) | bit(RS) | bit(CM) | bit(H))) {
        info.pos = lastPos;
        if (info.cat === H && info.pos === POS_PRE_M) {
          // Uniscribe doesn't move a halant along with a left matra.
          for (let j = i; j > start; j--) {
            if (g(j - 1).pos !== POS_PRE_M) {
              info.pos = g(j - 1).pos;
              break;
            }
          }
        }
      } else if (info.pos !== POS_SMVD) {
        lastPos = info.pos;
      }
    }
  }
  // Post-base consonants own the marks since the previous consonant or matra.
  {
    let last = base;
    for (let i = base + 1; i < end; i++) {
      if (isConsonant(g(i))) {
        for (let j = last + 1; j < i; j++) if (g(j).pos < POS_SMVD) g(j).pos = g(i).pos;
        last = i;
      } else if (g(i).cat === M) last = i;
    }
  }

  // Stable sort by position.
  const sorted = buf
    .slice(start, end)
    .map((info, k) => ({ info, k }))
    .sort((a, b) => a.info.pos - b.info.pos || a.k - b.k);
  for (let k = 0; k < sorted.length; k++) buf[start + k] = sorted[k]!.info;

  // Find the base again; several left matras keep their logical order.
  let firstLeft = end;
  let lastLeft = end;
  base = end;
  for (let i = start; i < end; i++) {
    if (g(i).pos === POS_BASE_C) {
      base = i;
      break;
    }
    if (g(i).pos === POS_PRE_M) {
      if (firstLeft === end) firstLeft = i;
      lastLeft = i;
    }
  }
  if (firstLeft < lastLeft) {
    reverseRange(buf, firstLeft, lastLeft + 1);
    let i = firstLeft;
    for (let j = i; j <= lastLeft; j++) {
      if (g(j).cat === M) {
        reverseRange(buf, i, j + 1);
        i = j + 1;
      }
    }
  }

  // 3. Masks: reph, pre-base (half, and below forms in new-spec), post-base forms.
  for (let i = start; i < end && g(i).pos === POS_RA_TO_BECOME_REPH; i++) g(i).mask |= rphf;
  let mask = plan.maskOf('half');
  if (!d.isOldSpec && d.config.blwfMode === 'pre-and-post') mask |= plan.maskOf('blwf');
  for (let i = start; i < base; i++) g(i).mask |= mask;
  mask = plan.maskOf('blwf') | plan.maskOf('abvf') | plan.maskOf('pstf');
  for (let i = base + 1; i < end; i++) g(i).mask |= mask;

  if (d.isOldSpec && d.script === 'deva') {
    // Old-spec eyelash Ra: Ra,Halant before the base takes 'blwf' unless followed by ZWJ.
    const blwf = plan.maskOf('blwf');
    for (let i = start; i + 1 < base; i++) {
      if (g(i).cat === Ra && g(i + 1).cat === H && (i + 2 === base || g(i + 2).cat !== ZWJ)) {
        g(i).mask |= blwf;
        g(i + 1).mask |= blwf;
      }
    }
  }

  const pref = plan.maskOf('pref');
  if (pref && base + 2 < end) {
    // A Halant,Ra pair the font forms with 'pref' is a pre-base-reordering consonant.
    for (let i = base + 1; i + 1 < end; i++) {
      if (wouldSubstitute(ctx, d, 'pref', [g(i).glyph, g(i + 1).glyph])) {
        g(i).mask |= pref;
        g(i + 1).mask |= pref;
        break;
      }
    }
  }

  // ZWNJ blocks half forms of the consonant before it.
  const half = plan.maskOf('half');
  for (let i = start + 1; i < end; i++) {
    if (isJoiner(g(i))) {
      const nonJoiner = g(i).cat === ZWNJ;
      let j = i;
      do {
        j--;
        if (nonJoiner) g(j).mask &= ~half;
      } while (j > start && !isConsonant(g(j)));
    }
  }
}

function finalReordering(ctx: ShapeContext): void {
  const d = planData(ctx);
  forEachSyllable(ctx.buffer, (start, end) => finalReorderSyllable(ctx, d, start, end));
}

function finalReorderSyllable(ctx: ShapeContext, d: IndicPlanData, start: number, end: number): void {
  const buf = ctx.buffer;
  const plan = ctx.plan;
  const g = (i: number) => buf[i]!;
  const script = d.script;

  // A halant that went through a ligate-then-decompose round trip is still a halant.
  if (d.viramaGlyph) {
    for (let i = start; i < end; i++) {
      const info = g(i);
      if (info.glyph === d.viramaGlyph && isLigated(info) && isMultiplied(info)) {
        info.cat = H;
        info.props &= ~(GlyphProps.Ligated | GlyphProps.Multiplied);
      }
    }
  }

  const prefMask = plan.maskOf('pref');
  let tryPref = !!prefMask;

  // Find the base again (basic features may have merged glyphs).
  let base = start;
  for (; base < end; base++) {
    if (g(base).pos >= POS_BASE_C) {
      if (tryPref && base + 1 < end) {
        for (let i = base + 1; i < end; i++) {
          if (g(i).mask & prefMask) {
            if (!(isSubstituted(g(i)) && ligatedAndDidntMultiply(g(i)))) {
              // A 'pref' candidate that didn't form: the base is around here.
              base = i;
              while (base < end && isHalant(g(base))) base++;
              if (base < end) g(base).pos = POS_BASE_C;
              tryPref = false;
            }
            break;
          }
        }
        if (base === end) break;
      }
      if (script === 'mlym') {
        // Malayalam: skip over unformed below (but not post) forms.
        for (let i = base + 1; i < end; i++) {
          while (i < end && isJoiner(g(i))) i++;
          if (i === end || !isHalant(g(i))) break;
          i++;
          while (i < end && isJoiner(g(i))) i++;
          if (i < end && isConsonant(g(i)) && g(i).pos === POS_BELOW_C) {
            base = i;
            g(base).pos = POS_BASE_C;
          }
        }
      }
      if (start < base && g(base).pos > POS_BASE_C) base--;
      break;
    }
  }
  if (base === end && start < base && isOneOf(g(base - 1), bit(ZWJ))) base--;
  if (base < end) while (start < base && isOneOf(g(base), bit(N) | bit(H))) base--;

  // Pre-base matra: move it after the last standalone halant before the base (i.e. after
  // any half forms that didn't ligate), unless that halant is followed by ZWJ.
  if (start + 1 < end && start < base) {
    let newPos = base === end ? base - 2 : base - 1;
    if (script !== 'mlym' && script !== 'taml') {
      for (;;) {
        while (newPos > start && !isOneOf(g(newPos), bit(M) | bit(H))) newPos--;
        if (isHalant(g(newPos)) && g(newPos).pos !== POS_PRE_M) {
          if (newPos + 1 < end && g(newPos + 1).cat === ZWJ && newPos > start) {
            newPos--;
            continue;
          }
        } else newPos = start;
        break;
      }
    }
    if (start < newPos && g(newPos).pos !== POS_PRE_M) {
      for (let i = newPos; i > start; i--) {
        if (g(i - 1).pos === POS_PRE_M) {
          const oldPos = i - 1;
          if (oldPos < base && base <= newPos) base--;
          moveGlyph(buf, oldPos, newPos);
          newPos--;
        }
      }
    }
  }

  // Reph: only move it when Ra+Halant actually ligated into a reph glyph (or, for an
  // atomically encoded repha, when it did NOT ligate).
  if (start + 1 < end && g(start).pos === POS_RA_TO_BECOME_REPH && (g(start).cat === Repha) !== ligatedAndDidntMultiply(g(start))) {
    const rephPos = d.config.rephPos;
    const afterFirstHalant = (): number => {
      let p = start + 1;
      while (p < base && !isHalant(g(p))) p++;
      if (p < base && isHalant(g(p))) {
        if (p + 1 < base && isJoiner(g(p + 1))) p++;
        return p;
      }
      return -1;
    };
    const target = ((): number => {
      if (rephPos !== POS_AFTER_POST) {
        // 2. After the first explicit halant between the reph and the main consonant.
        const p = afterFirstHalant();
        if (p >= 0) return p;
        // 3. After the main consonant (and anything ligated / placed right after it).
        if (rephPos === POS_AFTER_MAIN) {
          let q = base;
          while (q + 1 < end && g(q + 1).pos <= POS_AFTER_MAIN) q++;
          if (q < end) return q;
        }
        // 4. Before the first post-base consonant / matra / modifier.
        if (rephPos === POS_AFTER_SUB) {
          let q = base;
          while (q + 1 < end && !(bit(g(q + 1).pos) & (bit(POS_POST_C) | bit(POS_AFTER_POST) | bit(POS_SMVD)))) q++;
          if (q < end) return q;
        }
      }
      // 5. Same as 2 (for after-post scripts).
      const p = afterFirstHalant();
      if (p >= 0) return p;
      // 6. The end of the syllable, before syllable modifiers / vedic signs.
      let q = end - 1;
      while (q > start && g(q).pos === POS_SMVD) q--;
      if (isHalant(g(q))) {
        // After a Matra,Halant sequence, stay before the halant so reph can meet the matra.
        for (let i = base + 1; i < q; i++) if (g(i).cat === M) q--;
      }
      return q;
    })();
    moveGlyph(buf, start, target);
    if (start < base && base <= target) base--;
  }

  // Pre-base-reordering consonant (e.g. Malayalam Ra): move it before the base, after
  // any half forms, when 'pref' actually formed it.
  if (tryPref && base + 1 < end) {
    for (let i = base + 1; i < end; i++) {
      if (g(i).mask & prefMask) {
        if (ligatedAndDidntMultiply(g(i))) {
          let newPos = base;
          if (script !== 'mlym' && script !== 'taml') {
            while (newPos > start && !isOneOf(g(newPos - 1), bit(M) | bit(H))) newPos--;
          }
          if (newPos > start && isHalant(g(newPos - 1)) && newPos < end && isJoiner(g(newPos))) newPos++;
          moveGlyph(buf, i, newPos);
          if (newPos <= base && base < i) base++;
        }
        break;
      }
    }
  }

  // 'init' applies to a left matra at the start of a word.
  if (g(start).pos === POS_PRE_M && (start === 0 || !isWordCharacter(g(start - 1).codepoint))) g(start).mask |= plan.maskOf('init');
}

const RE_WORD = /[\p{L}\p{M}\p{Cf}\p{Co}\p{Cn}]/u;
function isWordCharacter(cp: number): boolean {
  return RE_WORD.test(String.fromCodePoint(cp));
}

// --- Features ---------------------------------------------------------------------------

function collectFeatures(b: PlanBuilder): void {
  const manual = { manualZwj: true, manualZwnj: true, perSyllable: true };
  b.pauseGsub(setupSyllables);
  b.enable('locl', { perSyllable: true });
  b.enable('ccmp', { perSyllable: true });
  b.pauseGsub(initialReordering);
  for (const tag of ['nukt', 'akhn', 'rphf', 'rkrf', 'pref', 'blwf', 'abvf', 'half', 'pstf', 'vatu', 'cjct']) {
    const global = tag === 'nukt' || tag === 'akhn' || tag === 'rkrf' || tag === 'vatu' || tag === 'cjct';
    b.add(tag, { ...manual, global });
    b.pauseGsub();
  }
  b.pauseGsub(finalReordering);
  b.add('init', manual);
  for (const tag of ['pres', 'abvs', 'blws', 'psts', 'haln']) b.enable(tag, manual);
  for (const tag of ['dist', 'abvm', 'blwm']) b.enable(tag);
}

export const indicShaper: Shaper = {
  name: 'indic',
  collectFeatures,
  // Uniscribe doesn't apply 'liga' to Indic scripts; fonts put Latin-only ligatures there.
  overrideFeatures: (b) => b.remove('liga'),
  setupMasks,
  zeroMarks: false,
  compose: false,
  alwaysDecompose: (cp) => ALWAYS_DECOMPOSE.has(cp),
};
