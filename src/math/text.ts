/**
 * Text in equations: picking fonts and code points for math runs (Word's math italic, m:sty,
 * m:scr, m:nor and run formatting), splitting runs into atoms for operator spacing, and reading
 * glyph ink bounds, which place scripts, limits and fraction parts far more tightly than the
 * font's line ascent and descent would.
 */
import type { FontProvider, LayoutFont } from '../layout/types.js';
import { measureChars, textItem } from '../layout/paint.js';
import { attr, firstChild } from '../xml/dom.js';
import type { XmlElement } from '../xml/dom.js';
import { SCALE } from './boxes.js';
import type { Atom, AtomClass, Box, Ctx } from './boxes.js';
import { onOff, runText, wOnOff } from './omml.js';
import {
  FUNCTION_NAMES,
  GLYPH_FALLBACK,
  charClass,
  decodeMathAlpha,
  displayChar,
  encodeMathAlpha,
  isDigit,
  isLetter,
  isSpace,
} from './symbols.js';
import type { Script } from './symbols.js';

/**
 * Word typesets equations in Cambria Math. Registries without it map it (and Cambria) to the
 * metric-compatible Caladea; a serif with Greek comes next, then whatever font has the character.
 */
const MATH_FAMILIES = ['Cambria Math', 'Cambria', 'Times New Roman'];

export class MathFonts {
  private readonly cache = new Map<string, LayoutFont>();

  constructor(readonly provider: FontProvider) {}

  family(name: string, bold: boolean, italic: boolean): LayoutFont {
    const key = `${name}|${bold}|${italic}`;
    let font = this.cache.get(key);
    if (!font) {
      font = this.provider.get(name, bold, italic);
      this.cache.set(key, font);
    }
    return font;
  }

  /** The first font of the math chain (after `family`, when given) that has the character. */
  pick(cp: number, bold: boolean, italic: boolean, family?: string): LayoutFont {
    const chain = family ? [family, ...MATH_FAMILIES] : MATH_FAMILIES;
    for (const name of chain) {
      const font = this.family(name, bold, italic);
      if (font.has(cp)) return font;
    }
    return this.provider.forChar(this.family(chain[0]!, bold, italic), cp);
  }

  /** Whether some registered font can show the character. */
  canShow(cp: number): boolean {
    return this.pick(cp, false, false).has(cp);
  }
}

// ---------------------------------------------------------------------------------------------
// Glyph ink.

/** Ink bounds of a glyph in em, y up. */
export interface Ink {
  xMin: number;
  yMin: number;
  xMax: number;
  yMax: number;
}

const inkCache = new WeakMap<LayoutFont, Map<number, Ink | null>>();

/**
 * Exact bounds from the font's glyf table (null for a glyph without ink, undefined when the
 * font file is not available, as with canvas-measured fonts, or has CFF outlines).
 */
export function fontInk(font: LayoutFont, cp: number): Ink | null | undefined {
  const sfnt = font.face?.font;
  if (!sfnt) return undefined;
  let cache = inkCache.get(font);
  if (!cache) {
    cache = new Map();
    inkCache.set(font, cache);
  }
  if (cache.has(cp)) return cache.get(cp);
  const range = sfnt.glyphRange(font.glyphId(cp));
  if (!range) return undefined;
  let ink: Ink | null = null;
  if (range.length >= 10) {
    const view = new DataView(sfnt.data.buffer, sfnt.data.byteOffset + range.offset, 10);
    const upem = sfnt.unitsPerEm;
    ink = { xMin: view.getInt16(2) / upem, yMin: view.getInt16(4) / upem, xMax: view.getInt16(6) / upem, yMax: view.getInt16(8) / upem };
    // Renderers slant a synthetic italic by about 12°.
    if (font.syntheticItalic) ink = { ...ink, xMin: ink.xMin + Math.min(0, ink.yMin) * 0.21, xMax: ink.xMax + Math.max(0, ink.yMax) * 0.21 };
  }
  cache.set(cp, ink);
  return ink;
}

const ASCENDERS = new Set([...'bdfhklt']);
const DESCENDERS = new Set([...'gjpqyQ']);

/** Typical bounds when the font file is not available. */
function approximateInk(font: LayoutFont, cp: number): Ink | null {
  const ch = String.fromCodePoint(cp);
  if (isSpace(ch)) return null;
  const xMax = font.advance(cp);
  let yMax = 0.7;
  let yMin = 0;
  if (ch >= 'a' && ch <= 'z' && !ASCENDERS.has(ch)) yMax = 0.47;
  if (DESCENDERS.has(ch) || '()[]{}|'.includes(ch)) yMin = -0.2;
  if ('()[]{}|'.includes(ch)) yMax = 0.75;
  if ('+−-=±×÷<>≤≥≠≈·∙'.includes(ch)) {
    yMax = 0.5;
    yMin = 0.05;
  }
  return { xMin: 0, yMin, xMax, yMax };
}

export function inkOf(font: LayoutFont, cp: number): Ink | null {
  const exact = fontInk(font, cp);
  return exact === undefined ? approximateInk(font, cp) : exact;
}

/** Height of the math axis (em): the centre of the minus sign, which fraction bars line up with. */
export function axisHeight(fonts: MathFonts): number {
  const minus = 0x2212;
  const font = fonts.pick(minus, false, false);
  const ink = font.has(minus) ? fontInk(font, minus) : undefined;
  if (ink && ink.yMax > ink.yMin && ink.yMax < 0.6 && ink.yMin > 0.1) return (ink.yMin + ink.yMax) / 2;
  return 0.25;
}

// ---------------------------------------------------------------------------------------------
// Boxes of glyphs.

export interface Glyph {
  /** The character as drawn (one code point). */
  ch: string;
  font: LayoutFont;
}

/** A box of glyphs on the baseline; neighbouring glyphs of one font share a text item. */
export function glyphsBox(glyphs: Glyph[], size: number, color: string, simple = true): Box {
  const items: Box['items'] = [];
  let x = 0;
  let ascent = 0;
  let descent = 0;
  let italic = 0;
  let i = 0;
  while (i < glyphs.length) {
    const font = glyphs[i]!.font;
    let j = i + 1;
    while (j < glyphs.length && glyphs[j]!.font === font) j++;
    const text = glyphs
      .slice(i, j)
      .map((g) => g.ch)
      .join('');
    const m = measureChars(text, font, size, x);
    items.push(textItem(x, 0, m.chars, m.xs, m.widths, font, size, color));
    m.chars.forEach((ch, k) => {
      const cp = ch.codePointAt(0)!;
      const ink = inkOf(font, cp);
      if (!ink) {
        italic = 0;
        return;
      }
      ascent = Math.max(ascent, ink.yMax * size);
      descent = Math.max(descent, -ink.yMin * size);
      italic = font.italic ? Math.max(0, ink.xMax * size - m.widths[k]!) : 0;
    });
    x += m.width;
    i = j;
  }
  return { width: x, ascent, descent, items, italic, simple };
}

/** A symbol in the math fonts, with a look-alike when no font has it. */
export function symbolGlyph(ch: string, fonts: MathFonts, bold: boolean): Glyph {
  const cp = ch.codePointAt(0)!;
  const font = fonts.pick(cp, bold, false);
  if (font.has(cp)) return { ch, font };
  const alt = GLYPH_FALLBACK[ch];
  if (alt) {
    const altFont = fonts.pick(alt.codePointAt(0)!, bold, false);
    if (altFont.has(alt.codePointAt(0)!)) return { ch: alt, font: altFont };
  }
  return { ch, font };
}

/** A string of symbols in the math fonts (delimiter glyphs, the lin fraction slash, fallbacks). */
export function symbolBox(text: string, size: number, color: string, fonts: MathFonts, bold = false): Box {
  return glyphsBox(
    [...text].map((ch) => symbolGlyph(ch, fonts, bold)),
    size,
    color,
  );
}

// ---------------------------------------------------------------------------------------------
// Runs.

interface RunFormat {
  sty: 'p' | 'b' | 'i' | 'bi' | undefined;
  scr: Script | undefined;
  nor: boolean;
  color: string | undefined;
  /** Level-0 size from w:sz (pt). */
  size: number | undefined;
  family: string | undefined;
  bold: boolean | undefined;
  italic: boolean | undefined;
}

const STY = new Set(['p', 'b', 'i', 'bi']);
const SCRIPTS = new Set<string>(['roman', 'script', 'fraktur', 'double-struck', 'sans-serif', 'monospace']);

function readFormat(run: XmlElement): RunFormat {
  const m = firstChild(run, 'm:rPr');
  const w = firstChild(run, 'w:rPr');
  const sty = attr(firstChild(m, 'm:sty'), 'm:val');
  const scr = attr(firstChild(m, 'm:scr'), 'm:val');
  const color = attr(firstChild(w, 'w:color'), 'w:val');
  const sz = Number(attr(firstChild(w, 'w:sz'), 'w:val'));
  const fonts = firstChild(w, 'w:rFonts');
  return {
    sty: sty && STY.has(sty) ? (sty as RunFormat['sty']) : undefined,
    scr: scr && SCRIPTS.has(scr) ? (scr as Script) : undefined,
    nor: onOff(m, 'm:nor', false),
    color: color && /^[0-9a-fA-F]{6}$/.test(color) ? color.toUpperCase() : undefined,
    size: Number.isFinite(sz) && sz > 0 ? sz / 2 : undefined,
    family: attr(fonts, 'w:ascii') ?? attr(fonts, 'w:hAnsi'),
    bold: wOnOff(w, 'w:b'),
    italic: wOnOff(w, 'w:i'),
  };
}

/** Size of a run at the current level (w:sz replaces the base size). */
function runSize(f: RunFormat, ctx: Ctx): number {
  return (f.size ?? ctx.base) * SCALE[ctx.level];
}

/** A letter in a math style: real math alphanumerics when a font has them, else an approximation. */
function letterGlyph(base: string, script: Script, bold: boolean, italic: boolean, fonts: MathFonts): Glyph {
  const cp = base.codePointAt(0)!;
  if (script !== 'roman') {
    const mapped = encodeMathAlpha(base, script, bold, italic);
    if (mapped !== undefined) {
      const font = fonts.pick(mapped, false, false);
      if (font.has(mapped)) return { ch: String.fromCodePoint(mapped), font };
    }
    switch (script) {
      case 'sans-serif':
        return { ch: base, font: fonts.pick(cp, bold, italic, 'Calibri') };
      case 'monospace':
        return { ch: base, font: fonts.pick(cp, bold, italic, 'Courier New') };
      case 'double-struck':
        return { ch: base, font: fonts.pick(cp, true, false) };
      case 'script':
        return { ch: base, font: fonts.pick(cp, bold, true) };
      default:
        return { ch: base, font: fonts.pick(cp, bold, false) };
    }
  }
  if (italic) {
    // An upright math font with the math italic letters (the real Cambria Math has no italic face).
    const mapped = encodeMathAlpha(base, 'roman', bold, true);
    const math = fonts.family('Cambria Math', false, false);
    if (mapped !== undefined && math.face && math.has(mapped)) return { ch: String.fromCodePoint(mapped), font: math };
  }
  return { ch: base, font: fonts.pick(cp, bold, italic) };
}

function atomOf(box: Box, cls: AtomClass): Atom {
  return { box, left: cls, right: cls };
}

/** Normal text (m:nor, or a w:r inside an equation): upright, its own font, no operator spacing. */
function normalTextAtoms(text: string, f: RunFormat, ctx: Ctx): Atom[] {
  if (!text) return [];
  const bold = f.bold ?? false;
  const italic = f.italic ?? false;
  const glyphs = [...text].map((ch) => ({ ch, font: ctx.fonts.pick(ch.codePointAt(0)!, bold, italic, f.family) }));
  return [atomOf(glyphsBox(glyphs, runSize(f, ctx), f.color ?? ctx.color, false), 'ord')];
}

/** A w:r inside an equation is plain text with its run formatting. */
export function textRunAtoms(run: XmlElement, ctx: Ctx): Atom[] {
  return normalTextAtoms(runText(run), readFormat(run), ctx);
}

/**
 * The atoms of a math run. Letters are italic unless styled or inside a function name, digits
 * and operators upright; numbers and known function names stay whole; every operator is its own
 * atom so the row can space it by class.
 */
export function runAtoms(run: XmlElement, ctx: Ctx): Atom[] {
  const f = readFormat(run);
  const text = runText(run);
  if (!text) return [];
  if (f.nor) return normalTextAtoms(text, f, ctx);
  const { fonts } = ctx;
  const size = runSize(f, ctx);
  const color = f.color ?? ctx.color;
  const bold = f.sty === 'b' || f.sty === 'bi';
  const styItalic = f.sty === undefined ? undefined : f.sty === 'i' || f.sty === 'bi';
  const script = f.scr ?? 'roman';
  const chars = [...text];
  const atoms: Atom[] = [];
  const push = (glyphs: Glyph[], cls: AtomClass) => atoms.push(atomOf(glyphsBox(glyphs, size, color), cls));
  let i = 0;
  while (i < chars.length) {
    const ch = chars[i]!;
    const cp = ch.codePointAt(0)!;
    const alpha = decodeMathAlpha(cp);
    if (alpha) {
      // A mathematical alphanumeric typed directly: keep it when a real font has it.
      const font = fonts.pick(cp, false, false);
      const glyph = font.has(cp) ? { ch, font } : letterGlyph(alpha.base, alpha.script, alpha.bold || bold, alpha.italic, fonts);
      push([glyph], 'ord');
      i++;
      continue;
    }
    if (isDigit(ch)) {
      let j = i + 1;
      while (j < chars.length && (isDigit(chars[j]!) || (chars[j] === '.' && j + 1 < chars.length && isDigit(chars[j + 1]!)))) j++;
      const glyphs = chars.slice(i, j).map((c) => letterGlyph(c, script, bold, styItalic === true, fonts));
      push(glyphs, 'ord');
      i = j;
      continue;
    }
    if (isLetter(cp)) {
      let j = i + 1;
      while (j < chars.length && isLetter(chars[j]!.codePointAt(0)!)) j++;
      const word = chars.slice(i, j).join('');
      if (styItalic !== true && j - i > 1 && FUNCTION_NAMES.has(word)) {
        push(
          chars.slice(i, j).map((c) => letterGlyph(c, script, bold, false, fonts)),
          'op',
        );
      } else {
        const italic = styItalic ?? !ctx.fn;
        for (let k = i; k < j; k++) push([letterGlyph(chars[k]!, script, bold, italic, fonts)], 'ord');
      }
      i = j;
      continue;
    }
    i++;
    if (isSpace(ch)) {
      const font = fonts.pick(32, bold, false);
      atoms.push({ box: { width: font.advance(32) * size, ascent: 0, descent: 0, items: [], italic: 0 }, left: 'ord', right: 'ord', kind: 'glue' });
      continue;
    }
    if (ch === '&' && ctx.align) {
      atoms.push({ box: { width: 0, ascent: 0, descent: 0, items: [], italic: 0 }, left: 'ord', right: 'ord', kind: 'mark' });
      continue;
    }
    // A run of primes is one atom so it reads ′′ rather than spaced symbols.
    if (ch === "'" || ch === '′') {
      let primes = '′';
      while (i < chars.length && (chars[i] === "'" || chars[i] === '′')) {
        primes += '′';
        i++;
      }
      push([...primes].map((p) => symbolGlyph(p, fonts, bold)), 'ord');
      continue;
    }
    const cls = charClass(ch);
    push([symbolGlyph(displayChar(ch), fonts, bold)], cls);
  }
  return atoms;
}
