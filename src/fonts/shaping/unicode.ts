/**
 * The few Unicode properties shaping needs beyond scripts: mark general categories (glyph
 * classes for fonts without GDEF, Arabic transparency, cluster formation), default-ignorable
 * code points (ZWJ, ZWNJ, bidi controls: hidden after shaping), and bidi mirroring for
 * right-to-left runs. General categories come from the JS regex engine's Unicode tables
 * (\p{...}), so no data has to be bundled for them.
 */

export const UFlag = {
  /** General category Mn (non-spacing mark). */
  NonSpacingMark: 0x01,
  /** Mc (spacing mark). */
  SpacingMark: 0x02,
  /** Me (enclosing mark). */
  EnclosingMark: 0x04,
  /** Cf (format control). */
  Format: 0x08,
  DefaultIgnorable: 0x10,
  Zwnj: 0x20,
  Zwj: 0x40,
  /** Joins the previous character's cluster (marks, variation selectors, emoji modifiers). */
  Continuation: 0x80,
} as const;

const RE_MN = /\p{Mn}/u;
const RE_MC = /\p{Mc}/u;
const RE_ME = /\p{Me}/u;
const RE_CF = /\p{Cf}/u;
const RE_PICTO = /\p{Extended_Pictographic}/u;

const flagCache = new Map<number, number>();

/** Default_Ignorable_Code_Point (DerivedCoreProperties), the ranges that occur in text. */
export function isDefaultIgnorable(cp: number): boolean {
  if (cp < 0xad) return false;
  return (
    cp === 0xad ||
    cp === 0x34f ||
    cp === 0x61c ||
    (cp >= 0x115f && cp <= 0x1160) ||
    (cp >= 0x17b4 && cp <= 0x17b5) ||
    (cp >= 0x180b && cp <= 0x180f) ||
    (cp >= 0x200b && cp <= 0x200f) ||
    (cp >= 0x202a && cp <= 0x202e) ||
    (cp >= 0x2060 && cp <= 0x206f) ||
    cp === 0x3164 ||
    (cp >= 0xfe00 && cp <= 0xfe0f) ||
    cp === 0xfeff ||
    cp === 0xffa0 ||
    (cp >= 0xfff0 && cp <= 0xfff8) ||
    (cp >= 0x1bca0 && cp <= 0x1bca3) ||
    (cp >= 0x1d173 && cp <= 0x1d17a) ||
    (cp >= 0xe0000 && cp <= 0xe0fff)
  );
}

/** UFlag bits for a code point (cached: text reuses few distinct characters). */
export function unicodeFlags(cp: number): number {
  if (cp < 0x300) return cp === 0xad ? UFlag.Format | UFlag.DefaultIgnorable : 0;
  const cached = flagCache.get(cp);
  if (cached !== undefined) return cached;
  const s = String.fromCodePoint(cp);
  let f = 0;
  if (RE_MN.test(s)) f |= UFlag.NonSpacingMark;
  else if (RE_MC.test(s)) f |= UFlag.SpacingMark;
  else if (RE_ME.test(s)) f |= UFlag.EnclosingMark;
  else if (RE_CF.test(s)) f |= UFlag.Format;
  if (isDefaultIgnorable(cp)) f |= UFlag.DefaultIgnorable;
  if (cp === 0x200c) f |= UFlag.Zwnj;
  if (cp === 0x200d) f |= UFlag.Zwj;
  if (
    f & (UFlag.NonSpacingMark | UFlag.SpacingMark | UFlag.EnclosingMark) ||
    (cp >= 0xfe00 && cp <= 0xfe0f) ||
    (cp >= 0xe0100 && cp <= 0xe01ef) ||
    (cp >= 0x1f3fb && cp <= 0x1f3ff) ||
    (cp >= 0xe0020 && cp <= 0xe007f)
  ) {
    f |= UFlag.Continuation;
  }
  if (flagCache.size < 4096) flagCache.set(cp, f);
  return f;
}

export function isExtendedPictographic(cp: number): boolean {
  return cp >= 0xa9 && RE_PICTO.test(String.fromCodePoint(cp));
}

/**
 * HarfBuzz's "modified combining classes" for Hebrew points (SBL order: shin/sin dot and
 * dagesh before vowels, so fonts can compose shin+dot) and Arabic harakat (shadda first, so
 * shadda+vowel ligatures and mark-to-mark stacking work as typed).
 */
const MODIFIED_CCC = new Map<number, number>([
  [0x05b0, 22], [0x05b1, 15], [0x05b2, 16], [0x05b3, 17], [0x05b4, 23], [0x05b5, 18], [0x05b6, 19], [0x05b7, 20],
  [0x05b8, 21], [0x05b9, 14], [0x05ba, 14], [0x05bb, 24], [0x05bc, 12], [0x05bd, 25], [0x05bf, 13], [0x05c1, 10],
  [0x05c2, 11], [0x05c7, 21],
  [0x064b, 28], [0x064c, 29], [0x064d, 30], [0x064e, 31], [0x064f, 32], [0x0650, 33], [0x0651, 27], [0x0652, 34],
  [0x0670, 35],
]);

/**
 * Whether mark `b` must come before mark `a` (canonical reordering). The JS normalizer knows
 * every combining class: NFD swaps a pair exactly when ccc(a) > ccc(b) > 0.
 */
export function markGoesBefore(b: number, a: number): boolean {
  if (a === b) return false; // NFD of a doubled mark reads the same "swapped"
  const ma = MODIFIED_CCC.get(a);
  const mb = MODIFIED_CCC.get(b);
  if (ma !== undefined && mb !== undefined) return mb < ma;
  const sa = String.fromCodePoint(a);
  const sb = String.fromCodePoint(b);
  if (sa.normalize('NFD') !== sa || sb.normalize('NFD') !== sb) return false;
  return (sa + sb).normalize('NFD') === sb + sa;
}

/** Bidi_Mirroring_Glyph pairs for the brackets and math signs that occur in documents. */
const MIRROR_PAIRS: Array<[number, number]> = [
  [0x28, 0x29], [0x3c, 0x3e], [0x5b, 0x5d], [0x7b, 0x7d], [0xab, 0xbb], [0x0f3a, 0x0f3b], [0x0f3c, 0x0f3d],
  [0x169b, 0x169c], [0x2039, 0x203a], [0x2045, 0x2046], [0x207d, 0x207e], [0x208d, 0x208e], [0x2208, 0x220b],
  [0x2209, 0x220c], [0x220a, 0x220d], [0x2264, 0x2265], [0x2266, 0x2267], [0x226a, 0x226b], [0x2282, 0x2283],
  [0x2286, 0x2287], [0x2308, 0x2309], [0x230a, 0x230b], [0x2329, 0x232a], [0x2768, 0x2769], [0x276a, 0x276b],
  [0x276c, 0x276d], [0x276e, 0x276f], [0x2770, 0x2771], [0x2772, 0x2773], [0x2774, 0x2775], [0x27e6, 0x27e7],
  [0x27e8, 0x27e9], [0x27ea, 0x27eb], [0x2983, 0x2984], [0x2985, 0x2986], [0x2987, 0x2988], [0x2989, 0x298a],
  [0x3008, 0x3009], [0x300a, 0x300b], [0x300c, 0x300d], [0x300e, 0x300f], [0x3010, 0x3011], [0x3014, 0x3015],
  [0x3016, 0x3017], [0x3018, 0x3019], [0x301a, 0x301b], [0xfe59, 0xfe5a], [0xfe5b, 0xfe5c], [0xfe5d, 0xfe5e],
  [0xfe64, 0xfe65], [0xff08, 0xff09], [0xff1c, 0xff1e], [0xff3b, 0xff3d], [0xff5b, 0xff5d], [0xff5f, 0xff60],
  [0xff62, 0xff63],
];
const MIRRORS = new Map<number, number>();
for (const [a, b] of MIRROR_PAIRS) {
  MIRRORS.set(a, b);
  MIRRORS.set(b, a);
}

/** The mirrored counterpart of a code point in right-to-left text, or undefined. */
export function mirrorOf(cp: number): number | undefined {
  return MIRRORS.get(cp);
}
