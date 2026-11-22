/**
 * Arabic-style shaper (Arabic, Syriac, N'Ko, Mongolian): joining analysis picks each letter's
 * contextual form (isol / fina / medi / init, plus Syriac fin2 / fin3 / med2 after Alaph),
 * and those forms are applied as per-glyph features one stage at a time before rlig (where
 * fonts form lam-alef), calt, liga and mset. Positioning then runs curs (cursive joining,
 * e.g. Nastaliq), kern, mark and mkmk.
 *
 * Joining types are from Unicode's ArabicShaping.txt for the Arabic, Arabic Supplement,
 * Arabic Extended-A/B, Syriac, N'Ko and Mongolian blocks. Characters not listed are
 * non-joining, except marks and format controls, which are transparent.
 *
 * For fonts with no Arabic GSUB forms at all, a fallback maps letters to their Unicode
 * presentation forms (U+FB50..U+FEFC) when the font has those glyphs.
 */
import { deleteGlyph, isLigated, mergeClusters, type GlyphInfo } from './buffer.js';
import { GlyphProps } from './gdef.js';
import type { PlanBuilder, ShapeContext } from './plan.js';
import type { Shaper } from './shaper.js';
import { UFlag } from './unicode.js';

const JT_U = 0;
const JT_L = 1;
const JT_R = 2;
const JT_D = 3;
const JG_ALAPH = 4;
const JG_DALATH_RISH = 5;
const JT_T = 7;

/** "range type" entries; C (join-causing) behaves as D. */
const JOINING_DATA = `
0600-0605 U;0620 D;0621 U;0622-0625 R;0626 D;0627 R;0628 D;0629 R;062A-062E D;062F-0632 R;0633-063F D;0640 D;
0641-0647 D;0648 R;0649-064A D;066E-066F D;0671-0673 R;0674 U;0675-0677 R;0678-0687 D;0688-0699 R;069A-06BF D;
06C0 R;06C1-06C2 D;06C3-06CB R;06CC D;06CD R;06CE D;06CF R;06D0-06D1 D;06D2-06D3 R;06D5 R;06DD U;06EE-06EF R;
06FA-06FC D;06FF D;
0710 A;0712-0714 D;0715-0716 S;0717-0719 R;071A-071D D;071E R;071F-0727 D;0728 R;0729 D;072A S;072B D;072C R;
072D-072E D;072F S;074D R;074E-074F D;
0750-0758 D;0759-075B R;075C-076A D;076B-076C R;076D-0770 D;0771 R;0772 D;0773-0774 R;0775-0777 D;0778-077B R;
077C-077F D;
07CA-07EA D;07FA D;
0860 D;0862-0865 D;0867 R;0868 D;0869-086A R;
0870-0882 R;0883-0885 D;0886 D;0889-088D D;088E R;0890-0891 U;
08A0-08A9 D;08AA-08AC R;08AD U;08AE R;08AF-08B0 D;08B1-08B2 R;08B3-08B8 D;08B9 R;08BA-08C8 D;08E2 U;
1807 D;180A D;1820-1878 D;1887-18A8 D;18AA D;
200C U;200D D;202F U
`;

const JOINING = new Map<number, number>();
{
  const codes: Record<string, number> = { U: JT_U, L: JT_L, R: JT_R, D: JT_D, A: JG_ALAPH, S: JG_DALATH_RISH };
  for (const entry of JOINING_DATA.split(/[;\n]/)) {
    const m = /^\s*([0-9A-F]{4,5})(?:-([0-9A-F]{4,5}))?\s+([ULRDAS])\s*$/.exec(entry);
    if (!m) continue;
    const first = parseInt(m[1]!, 16);
    const last = m[2] ? parseInt(m[2], 16) : first;
    for (let cp = first; cp <= last; cp++) JOINING.set(cp, codes[m[3]!]!);
  }
}

function joiningType(info: GlyphInfo): number {
  const jt = JOINING.get(info.codepoint);
  if (jt !== undefined) return jt;
  return info.uflags & (UFlag.NonSpacingMark | UFlag.EnclosingMark | UFlag.Format) ? JT_T : JT_U;
}

export const ISOL = 0;
export const FINA = 1;
export const FIN2 = 2;
export const FIN3 = 3;
export const MEDI = 4;
export const MED2 = 5;
export const INIT = 6;
export const NONE = 7;

export const ARABIC_FEATURES = ['isol', 'fina', 'fin2', 'fin3', 'medi', 'med2', 'init'] as const;

/**
 * Joining state machine (HarfBuzz's, after the Unicode joining rules): for each state and
 * joining type, the action for the previous letter, the action for this one, the next state.
 * Columns: U, L, R, D, Alaph, Dalath/Rish.
 */
const STATE_TABLE: ReadonlyArray<ReadonlyArray<readonly [number, number, number]>> = [
  // 0: prev was U, not willing to join.
  [[NONE, NONE, 0], [NONE, ISOL, 2], [NONE, ISOL, 1], [NONE, ISOL, 2], [NONE, ISOL, 1], [NONE, ISOL, 6]],
  // 1: prev was R or ISOL/ALAPH, not willing to join.
  [[NONE, NONE, 0], [NONE, ISOL, 2], [NONE, ISOL, 1], [NONE, ISOL, 2], [NONE, FIN2, 5], [NONE, ISOL, 6]],
  // 2: prev was D/L in ISOL form, willing to join.
  [[NONE, NONE, 0], [NONE, ISOL, 2], [INIT, FINA, 1], [INIT, FINA, 3], [INIT, FINA, 4], [INIT, FINA, 6]],
  // 3: prev was D in FINA form, willing to join.
  [[NONE, NONE, 0], [NONE, ISOL, 2], [MEDI, FINA, 1], [MEDI, FINA, 3], [MEDI, FINA, 4], [MEDI, FINA, 6]],
  // 4: prev was FINA ALAPH, not willing to join.
  [[NONE, NONE, 0], [NONE, ISOL, 2], [MED2, ISOL, 1], [MED2, ISOL, 2], [MED2, FIN2, 5], [MED2, ISOL, 6]],
  // 5: prev was FIN2/FIN3 ALAPH, not willing to join.
  [[NONE, NONE, 0], [NONE, ISOL, 2], [ISOL, ISOL, 1], [ISOL, ISOL, 2], [ISOL, FIN2, 5], [ISOL, ISOL, 6]],
  // 6: prev was DALATH/RISH, not willing to join.
  [[NONE, NONE, 0], [NONE, ISOL, 2], [NONE, ISOL, 1], [NONE, ISOL, 2], [NONE, FIN3, 5], [NONE, ISOL, 6]],
];

/** Stores each glyph's joining action (ISOL..INIT or NONE) in `info.cat`. */
export function arabicJoining(buf: GlyphInfo[]): void {
  let prev = -1;
  let state = 0;
  for (let i = 0; i < buf.length; i++) {
    const info = buf[i]!;
    const type = joiningType(info);
    if (type === JT_T) {
      info.cat = NONE;
      continue;
    }
    const [prevAction, curAction, next] = STATE_TABLE[state]![type]!;
    if (prevAction !== NONE && prev >= 0) buf[prev]!.cat = prevAction;
    info.cat = curAction;
    prev = i;
    state = next;
  }
}

function setupMasks(ctx: ShapeContext): void {
  const buf = ctx.buffer;
  arabicJoining(buf);
  if (ctx.plan.script === 'mong') {
    // Mongolian free variation selectors take the form of the letter they follow.
    for (let i = 1; i < buf.length; i++) {
      const cp = buf[i]!.codepoint;
      if ((cp >= 0x180b && cp <= 0x180d) || cp === 0x180f) buf[i]!.cat = buf[i - 1]!.cat;
    }
  }
  const masks = ARABIC_FEATURES.map((tag) => ctx.plan.maskOf(tag));
  for (const info of buf) if (info.cat !== NONE) info.mask |= masks[info.cat] ?? 0;
}

// --- Presentation-form fallback --------------------------------------------------------

/** Code point → [isol, fina, init, medi] presentation forms (0 where the form doesn't exist). */
const PRESENTATION = new Map<number, [number, number, number, number]>();
{
  // Presentation Forms-B lists U+0621..U+064A in order, with 1, 2 or 4 forms per letter.
  let form = 0xfe80;
  const letters: number[] = [];
  for (let cp = 0x0621; cp <= 0x063a; cp++) letters.push(cp);
  for (let cp = 0x0641; cp <= 0x064a; cp++) letters.push(cp);
  for (const cp of letters) {
    const jt = cp === 0x0649 ? JT_R : (JOINING.get(cp) ?? JT_U); // alef maksura has only isol/fina here
    if (jt === JT_U) PRESENTATION.set(cp, [form++, 0, 0, 0]);
    else if (jt === JT_R) {
      PRESENTATION.set(cp, [form, form + 1, 0, 0]);
      form += 2;
    } else {
      PRESENTATION.set(cp, [form, form + 1, form + 2, form + 3]);
      form += 4;
    }
  }
  // Common Persian / Urdu letters from Presentation Forms-A.
  const extra: Array<[number, number, number]> = [
    [0x0679, 0xfb66, 4], [0x067e, 0xfb56, 4], [0x0686, 0xfb7a, 4], [0x0688, 0xfb88, 2], [0x0691, 0xfb8c, 2],
    [0x0698, 0xfb8a, 2], [0x06a9, 0xfb8e, 4], [0x06af, 0xfb92, 4], [0x06ba, 0xfb9e, 2], [0x06be, 0xfbaa, 4],
    [0x06c1, 0xfba6, 4], [0x06cc, 0xfbfc, 4], [0x06d2, 0xfbae, 2], [0x06d3, 0xfbb0, 2],
  ];
  for (const [cp, base, n] of extra) PRESENTATION.set(cp, n === 4 ? [base, base + 1, base + 2, base + 3] : [base, base + 1, 0, 0]);
}

/** Lam + alef variant → [isolated, final] lam-alef ligature. */
const LAM_ALEF = new Map<number, [number, number]>([
  [0x0622, [0xfef5, 0xfef6]],
  [0x0623, [0xfef7, 0xfef8]],
  [0x0625, [0xfef9, 0xfefa]],
  [0x0627, [0xfefb, 0xfefc]],
]);

/**
 * Shapes Arabic with the font's presentation-form glyphs when its GSUB has none of the
 * joining features (older fonts, or AAT-only fonts whose cmap still covers U+FExx).
 */
function fallbackShape(ctx: ShapeContext): void {
  const plan = ctx.plan;
  if (['isol', 'fina', 'medi', 'init'].some((tag) => plan.hasFeature(0, tag))) return;
  const buf = ctx.buffer;
  const font = ctx.font;
  const rlig = plan.maskOf('rlig');

  for (let i = 0; i < buf.length; i++) {
    const lam = buf[i]!;
    if (lam.codepoint !== 0x0644 || (lam.cat !== INIT && lam.cat !== MEDI) || !(lam.mask & rlig) || isLigated(lam)) continue;
    let j = i + 1;
    while (j < buf.length && buf[j]!.props & GlyphProps.Mark) j++;
    const alef = buf[j];
    const forms = alef && alef.cat === FINA ? LAM_ALEF.get(alef.codepoint) : undefined;
    if (!forms) continue;
    const glyph = font.glyphId(lam.cat === INIT ? forms[0] : forms[1]);
    if (!glyph) continue;
    mergeClusters(buf, i, j + 1);
    lam.glyph = glyph;
    lam.props = (lam.props & ~GlyphProps.Mark) | GlyphProps.Substituted | GlyphProps.Ligated | GlyphProps.Ligature;
    lam.cat = lam.cat === INIT ? ISOL : FINA;
    deleteGlyph(buf, j);
  }

  const masks = ARABIC_FEATURES.map((tag) => plan.maskOf(tag));
  const slot = [0, 1, -1, -1, 3, -1, 2]; // action → index into [isol, fina, init, medi]
  for (const info of buf) {
    if (info.cat === NONE || isLigated(info) || !(info.mask & (masks[info.cat] ?? 0))) continue;
    const forms = PRESENTATION.get(info.codepoint);
    const s = slot[info.cat] ?? -1;
    const cp = forms && s >= 0 ? forms[s] : 0;
    const glyph = cp ? font.glyphId(cp) : 0;
    if (glyph) {
      info.glyph = glyph;
      info.props |= GlyphProps.Substituted;
    }
  }
}

function collectFeatures(b: PlanBuilder, script: string): void {
  b.enable('ccmp', { manualZwj: true });
  b.enable('locl', { manualZwj: true });
  b.pauseGsub();
  for (const tag of ARABIC_FEATURES) {
    // The Syriac-only forms have no presentation-form fallback.
    b.add(tag, { fallback: script === 'arab' && !tag.endsWith('2') && !tag.endsWith('3') });
    b.pauseGsub();
  }
  b.enable('rlig', { manualZwj: true, fallback: true });
  if (script === 'arab') b.pauseGsub(fallbackShape);
  b.enable('rclt', { manualZwj: true });
  b.enable('calt', { manualZwj: true });
  b.pauseGsub();
  b.enable('mset');
}

export const arabicShaper: Shaper = {
  name: 'arabic',
  collectFeatures,
  setupMasks,
  zeroMarks: true,
  compose: true,
};
