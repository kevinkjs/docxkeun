/**
 * Character knowledge for equations: the atom class of each operator (which decides the TeX-style
 * spacing Word applies), the function names Word sets upright, the Unicode mathematical
 * alphanumerics, and the accents, delimiters and big operators that layout draws itself.
 */
import type { AtomClass } from './boxes.js';

const REL = new Set([
  ...'=<>≠≤≥≦≧≨≩≪≫≮≯≰≱≈≉≊≋≡≢≣≅≆≇≃≄≂∼∽≍≐≑≒≓≔≕≜≝≞≟∝∈∉∊∋∌∍⊂⊃⊄⊅⊆⊇⊈⊉⊊⊋⊏⊐⊑⊒≺≻≼≽⪯⪰∣∤∥∦⊢⊣⊨⊩⊪⊫⊲⊳⊴⊵',
  ...'→←↔↑↓↕⇒⇐⇔⇑⇓⇕↦↤⟶⟵⟷⟹⟸⟺⟼↪↩⇀⇁↼↽⇌⇋⇄⇆↗↘↙↖:∶∷',
]);
const BIN = new Set([...'+-−±∓×÷·⋅∗∘∙∪∩∧∨⊕⊖⊗⊘⊙⊚⊛⊎⊓⊔⋄⋆⋉⋊⋋⋌⋎⋏⋒⋓△▽◁▷∖≀⊞⊟⊠⊡*']);
const OPEN = new Set([...'([{⟨⌈⌊⟦⟪⦃〈❲⁅']);
const CLOSE = new Set([...')]}⟩⌉⌋⟧⟫⦄〉❳⁆!']);
const PUNCT = new Set([...',;']);
const LARGE = new Set([...'∑∏∐∫∬∭∮∯∰∱∲∳⋀⋁⋂⋃⨀⨁⨂⨃⨄⨅⨆⨉⨌']);

export function charClass(ch: string): AtomClass {
  if (REL.has(ch)) return 'rel';
  if (BIN.has(ch)) return 'bin';
  if (OPEN.has(ch)) return 'open';
  if (CLOSE.has(ch)) return 'close';
  if (PUNCT.has(ch)) return 'punct';
  if (LARGE.has(ch)) return 'op';
  return 'ord';
}

/** Word shows these typed characters as their mathematical forms. */
const SHOWN: Record<string, string> = { '-': '−', '*': '∗', "'": '′' };

export function displayChar(ch: string): string {
  return SHOWN[ch] ?? ch;
}

/** Look-alikes for symbols no registered font has (the bundled fonts lack many math symbols). */
export const GLYPH_FALLBACK: Record<string, string> = {
  '⋅': '·',
  '∙': '·',
  '∗': '*',
  '∼': '~',
  '∣': '|',
  '∥': '‖',
  '−': '-',
  '′': "'",
  '″': '"',
  '∶': ':',
  '⋯': '…',
  '⟶': '→',
  '⟵': '←',
  '⟷': '↔',
  '⟹': '⇒',
  '〈': '⟨',
  '〉': '⟩',
};

/** Function names Word sets upright when they appear as a whole word. */
export const FUNCTION_NAMES = new Set([
  'sin', 'cos', 'tan', 'cot', 'sec', 'csc', 'sinh', 'cosh', 'tanh', 'coth', 'sech', 'csch',
  'arcsin', 'arccos', 'arctan', 'arccot', 'arcsec', 'arccsc', 'arsinh', 'arcosh', 'artanh',
  'log', 'ln', 'lg', 'exp', 'lim', 'liminf', 'limsup', 'max', 'min', 'sup', 'inf', 'det', 'dim',
  'ker', 'deg', 'gcd', 'lcm', 'arg', 'hom', 'Pr', 'mod', 'sgn', 'tr', 'rank', 'Re', 'Im', 'erf',
]);

export function isDigit(ch: string): boolean {
  return ch >= '0' && ch <= '9';
}

/** Latin and Greek letters: the characters Word's math italic applies to. */
export function isLetter(cp: number): boolean {
  return (
    (cp >= 0x41 && cp <= 0x5a) ||
    (cp >= 0x61 && cp <= 0x7a) ||
    (cp >= 0xc0 && cp <= 0x24f && cp !== 0xd7 && cp !== 0xf7) ||
    (cp >= 0x391 && cp <= 0x3a9) ||
    (cp >= 0x3b1 && cp <= 0x3c9) ||
    (cp >= 0x3d0 && cp <= 0x3d6) ||
    (cp >= 0x3f0 && cp <= 0x3f5)
  );
}

export function isSpace(ch: string): boolean {
  return ch === ' ' || ch === ' ' || ch === ' ' || ch === ' ' || ch === ' ' || ch === ' ' || ch === ' ' || ch === ' ';
}

// ---------------------------------------------------------------------------------------------
// Mathematical alphanumeric symbols (U+1D400…): documents may contain them directly (𝑥), and
// m:scr asks for them. Fonts without them get the base letter in an approximating style.

export type Script = 'roman' | 'script' | 'fraktur' | 'double-struck' | 'sans-serif' | 'monospace';

/** Latin block order: [script, bold, italic] for each 52-letter run from U+1D400. */
const LATIN_STYLES: Array<[Script, boolean, boolean]> = [
  ['roman', true, false],
  ['roman', false, true],
  ['roman', true, true],
  ['script', false, false],
  ['script', true, false],
  ['fraktur', false, false],
  ['double-struck', false, false],
  ['fraktur', true, false],
  ['sans-serif', false, false],
  ['sans-serif', true, false],
  ['sans-serif', false, true],
  ['sans-serif', true, true],
  ['monospace', false, false],
];
const GREEK_STYLES: Array<[Script, boolean, boolean]> = [
  ['roman', true, false],
  ['roman', false, true],
  ['roman', true, true],
  ['sans-serif', true, false],
  ['sans-serif', true, true],
];
const GREEK_ORDER = [...'ΑΒΓΔΕΖΗΘϴΙΚΛΜΝΞΟΠΡΣΤΥΦΧΨΩ∇αβγδεζηθικλμνξοπρςστυφχψω∂ϵϑϰϕϱϖ'];
const DIGIT_STYLES: Array<[Script, boolean]> = [
  ['roman', true],
  ['double-struck', false],
  ['sans-serif', false],
  ['sans-serif', true],
  ['monospace', false],
];
/** Letters the block leaves out because they were encoded earlier (ℎ, ℬ, ℂ…). */
const HOLES: Record<string, number> = {
  '1:h': 0x210e,
  '3:B': 0x212c, '3:E': 0x2130, '3:F': 0x2131, '3:H': 0x210b, '3:I': 0x2110, '3:L': 0x2112, '3:M': 0x2133, '3:R': 0x211b,
  '3:e': 0x212f, '3:g': 0x210a, '3:o': 0x2134,
  '5:C': 0x212d, '5:H': 0x210c, '5:I': 0x2111, '5:R': 0x211c, '5:Z': 0x2128,
  '6:C': 0x2102, '6:H': 0x210d, '6:N': 0x2115, '6:P': 0x2119, '6:Q': 0x211a, '6:R': 0x211d, '6:Z': 0x2124,
};
const HOLE_DECODE = new Map<number, string>(Object.entries(HOLES).map(([k, cp]) => [cp, k]));

export interface MathAlpha {
  base: string;
  script: Script;
  bold: boolean;
  italic: boolean;
}

export function decodeMathAlpha(cp: number): MathAlpha | undefined {
  const hole = HOLE_DECODE.get(cp);
  if (hole) {
    const [index, base] = hole.split(':') as [string, string];
    const [script, bold, italic] = LATIN_STYLES[Number(index)]!;
    return { base, script, bold, italic };
  }
  if (cp >= 0x1d400 && cp < 0x1d400 + 13 * 52) {
    const i = cp - 0x1d400;
    const [script, bold, italic] = LATIN_STYLES[Math.floor(i / 52)]!;
    const k = i % 52;
    return { base: String.fromCharCode(k < 26 ? 65 + k : 97 + k - 26), script, bold, italic };
  }
  if (cp === 0x1d6a4 || cp === 0x1d6a5) return { base: cp === 0x1d6a4 ? 'ı' : 'ȷ', script: 'roman', bold: false, italic: true };
  if (cp >= 0x1d6a8 && cp < 0x1d6a8 + 5 * 58) {
    const i = cp - 0x1d6a8;
    const [script, bold, italic] = GREEK_STYLES[Math.floor(i / 58)]!;
    return { base: GREEK_ORDER[i % 58]!, script, bold, italic };
  }
  if (cp >= 0x1d7ce && cp <= 0x1d7ff) {
    const i = cp - 0x1d7ce;
    const [script, bold] = DIGIT_STYLES[Math.floor(i / 10)]!;
    return { base: String(i % 10), script, bold, italic: false };
  }
  return undefined;
}

/** The mathematical alphanumeric for a letter or digit in a style, when Unicode has one. */
export function encodeMathAlpha(base: string, script: Script, bold: boolean, italic: boolean): number | undefined {
  const cp = base.codePointAt(0)!;
  if (isDigitCp(cp)) {
    const i = DIGIT_STYLES.findIndex(([s, b]) => s === script && b === bold);
    return i < 0 || (script === 'roman' && !bold) ? undefined : 0x1d7ce + i * 10 + (cp - 48);
  }
  const latin = (cp >= 65 && cp <= 90) || (cp >= 97 && cp <= 122);
  if (latin) {
    // Styles without a bold or italic variant fall back to the plain one.
    let i = LATIN_STYLES.findIndex(([s, b, it]) => s === script && b === bold && it === italic);
    if (i < 0) i = LATIN_STYLES.findIndex(([s, b]) => s === script && b === bold);
    if (i < 0) i = LATIN_STYLES.findIndex(([s]) => s === script);
    if (i < 0) return undefined;
    const hole = HOLES[`${i}:${base}`];
    if (hole) return hole;
    return 0x1d400 + i * 52 + (cp <= 90 ? cp - 65 : 26 + cp - 97);
  }
  const g = GREEK_ORDER.indexOf(base);
  if (g >= 0) {
    const i = GREEK_STYLES.findIndex(([s, b, it]) => s === script && b === bold && it === italic);
    return i < 0 ? undefined : 0x1d6a8 + i * 58 + g;
  }
  return undefined;
}

function isDigitCp(cp: number): boolean {
  return cp >= 48 && cp <= 57;
}

// ---------------------------------------------------------------------------------------------
// Shapes layout draws as paths: they must stretch, or the bundled fonts lack them.

export type DelimiterKind = 'paren' | 'bracket' | 'ceil' | 'floor' | 'brace' | 'angle' | 'bar' | 'dbar';

/** Stretchy delimiters. `side` is the side the character opens from (its outer edge). */
const DELIMITERS: Record<string, { kind: DelimiterKind; side: 'left' | 'right' | 'both' }> = {
  '(': { kind: 'paren', side: 'left' },
  ')': { kind: 'paren', side: 'right' },
  '[': { kind: 'bracket', side: 'left' },
  ']': { kind: 'bracket', side: 'right' },
  '{': { kind: 'brace', side: 'left' },
  '}': { kind: 'brace', side: 'right' },
  '⌈': { kind: 'ceil', side: 'left' },
  '⌉': { kind: 'ceil', side: 'right' },
  '⌊': { kind: 'floor', side: 'left' },
  '⌋': { kind: 'floor', side: 'right' },
  '⟨': { kind: 'angle', side: 'left' },
  '⟩': { kind: 'angle', side: 'right' },
  '〈': { kind: 'angle', side: 'left' },
  '〉': { kind: 'angle', side: 'right' },
  '|': { kind: 'bar', side: 'both' },
  '∣': { kind: 'bar', side: 'both' },
  '│': { kind: 'bar', side: 'both' },
  '‖': { kind: 'dbar', side: 'both' },
  '∥': { kind: 'dbar', side: 'both' },
};

export function delimiterOf(ch: string): { kind: DelimiterKind; side: 'left' | 'right' | 'both' } | undefined {
  return DELIMITERS[ch];
}

/** Horizontal stretchy characters of m:groupChr: the shape and the side its outer edge faces. */
const GROUP: Record<string, { kind: DelimiterKind | 'arrow'; outer: 'top' | 'bottom' }> = {
  '⏞': { kind: 'brace', outer: 'top' },
  '⏟': { kind: 'brace', outer: 'bottom' },
  '⏜': { kind: 'paren', outer: 'top' },
  '⏝': { kind: 'paren', outer: 'bottom' },
  '⎴': { kind: 'bracket', outer: 'top' },
  '⎵': { kind: 'bracket', outer: 'bottom' },
  '︷': { kind: 'brace', outer: 'top' },
  '︸': { kind: 'brace', outer: 'bottom' },
};

export function groupShapeOf(ch: string): { kind: DelimiterKind | 'arrow'; outer: 'top' | 'bottom' } | undefined {
  return GROUP[ch];
}

/** Arrows that stretch (group characters, accents): which ends get heads, and whether doubled. */
const ARROWS: Record<string, { left: boolean; right: boolean; double: boolean; harpoon: boolean }> = {
  '→': { left: false, right: true, double: false, harpoon: false },
  '←': { left: true, right: false, double: false, harpoon: false },
  '↔': { left: true, right: true, double: false, harpoon: false },
  '⟶': { left: false, right: true, double: false, harpoon: false },
  '⟵': { left: true, right: false, double: false, harpoon: false },
  '⟷': { left: true, right: true, double: false, harpoon: false },
  '⇒': { left: false, right: true, double: true, harpoon: false },
  '⇐': { left: true, right: false, double: true, harpoon: false },
  '⇔': { left: true, right: true, double: true, harpoon: false },
  '⟹': { left: false, right: true, double: true, harpoon: false },
  '⟸': { left: true, right: false, double: true, harpoon: false },
  '⟺': { left: true, right: true, double: true, harpoon: false },
  '⃗': { left: false, right: true, double: false, harpoon: false },
  '⃖': { left: true, right: false, double: false, harpoon: false },
  '⃡': { left: true, right: true, double: false, harpoon: false },
  '⃑': { left: false, right: true, double: false, harpoon: true },
  '⃐': { left: true, right: false, double: false, harpoon: true },
  '⇀': { left: false, right: true, double: false, harpoon: true },
  '↼': { left: true, right: false, double: false, harpoon: true },
};

export function arrowOf(ch: string): { left: boolean; right: boolean; double: boolean; harpoon: boolean } | undefined {
  return ARROWS[ch];
}

export type AccentKind = 'hat' | 'check' | 'tilde' | 'bar' | 'dot' | 'ddot' | 'dddot' | 'acute' | 'grave' | 'breve' | 'ring' | 'arrow';

const ACCENTS: Record<string, AccentKind> = {
  '̂': 'hat', '^': 'hat', 'ˆ': 'hat',
  '̌': 'check', 'ˇ': 'check',
  '̃': 'tilde', '~': 'tilde', '˜': 'tilde',
  '̄': 'bar', '̅': 'bar', '¯': 'bar', '‾': 'bar',
  '̇': 'dot', '˙': 'dot',
  '̈': 'ddot', '¨': 'ddot',
  '⃛': 'dddot',
  '́': 'acute', '´': 'acute',
  '̀': 'grave', '`': 'grave',
  '̆': 'breve', '˘': 'breve',
  '̊': 'ring', '˚': 'ring',
};

export function accentOf(ch: string): AccentKind | undefined {
  return ACCENTS[ch] ?? (ARROWS[ch] ? 'arrow' : undefined);
}

/** Integral signs: how many strokes, and whether a contour circle crosses them. */
const INTEGRALS: Record<string, { count: number; contour: boolean }> = {
  '∫': { count: 1, contour: false },
  '∬': { count: 2, contour: false },
  '∭': { count: 3, contour: false },
  '⨌': { count: 4, contour: false },
  '∮': { count: 1, contour: true },
  '∯': { count: 2, contour: true },
  '∰': { count: 3, contour: true },
  '∱': { count: 1, contour: true },
  '∲': { count: 1, contour: true },
  '∳': { count: 1, contour: true },
};

export function integralOf(ch: string): { count: number; contour: boolean } | undefined {
  return INTEGRALS[ch];
}

/** Big operators layout always draws (the bundled fonts lack them, or their glyphs would not scale well). */
export const DRAWN_OPERATORS = new Set([...'∐⋃⋂⋁⋀⨀⨁⨂⨄⨆⨅']);
