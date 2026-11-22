/**
 * Equation layout (src/math): OMML built in the test, laid out with the bundled fonts, checked
 * by geometry (where scripts, bars, limits and delimiters land) and by linear-format text.
 */
import { describe, expect, it } from 'vitest';
import { RegistryFontProvider } from '../src/layout/fonts.js';
import type { FontProvider, LayoutFont, PathItem, RectItem, TextItem } from '../src/layout/types.js';
import { layoutMath, mathToText } from '../src/math/index.js';
import type { MathLayout } from '../src/math/index.js';
import { elements } from '../src/xml/dom.js';
import type { XmlElement } from '../src/xml/dom.js';
import { parseXml } from '../src/xml/parser.js';
import { testRegistry } from './helpers/fonts.js';

const fonts = new RegistryFontProvider(testRegistry());
const NS = 'xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
const SIZE = 12;
const EM = SIZE;

// --- OMML builders ----------------------------------------------------------------------------

const esc = (s: string) => s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
const r = (text: string, rPr = '') => `<m:r>${rPr}<m:t xml:space="preserve">${esc(text)}</m:t></m:r>`;
const sty = (v: string) => `<m:rPr><m:sty m:val="${v}"/></m:rPr>`;
const sSup = (base: string, sup: string) => `<m:sSup><m:e>${base}</m:e><m:sup>${sup}</m:sup></m:sSup>`;
const sSub = (base: string, sub: string) => `<m:sSub><m:e>${base}</m:e><m:sub>${sub}</m:sub></m:sSub>`;
const sSubSup = (base: string, sub: string, sup: string) => `<m:sSubSup><m:e>${base}</m:e><m:sub>${sub}</m:sub><m:sup>${sup}</m:sup></m:sSubSup>`;
const frac = (num: string, den: string, type?: string) => `<m:f>${type ? `<m:fPr><m:type m:val="${type}"/></m:fPr>` : ''}<m:num>${num}</m:num><m:den>${den}</m:den></m:f>`;
const rad = (body: string, deg?: string) => `<m:rad><m:radPr>${deg ? '' : '<m:degHide m:val="1"/>'}</m:radPr><m:deg>${deg ?? ''}</m:deg><m:e>${body}</m:e></m:rad>`;
const nary = (chr: string, sub: string, sup: string, body: string, props = '') =>
  `<m:nary><m:naryPr><m:chr m:val="${chr}"/>${props}</m:naryPr><m:sub>${sub}</m:sub><m:sup>${sup}</m:sup><m:e>${body}</m:e></m:nary>`;
const delim = (inner: string[], beg?: string, end?: string) =>
  `<m:d><m:dPr>${beg !== undefined ? `<m:begChr m:val="${beg}"/>` : ''}${end !== undefined ? `<m:endChr m:val="${end}"/>` : ''}</m:dPr>${inner.map((x) => `<m:e>${x}</m:e>`).join('')}</m:d>`;
const matrix = (rows: string[][]) => `<m:m>${rows.map((row) => `<m:mr>${row.map((c) => `<m:e>${c}</m:e>`).join('')}</m:mr>`).join('')}</m:m>`;
const func = (name: string, arg: string) => `<m:func><m:fName>${name}</m:fName><m:e>${arg}</m:e></m:func>`;

function parse(xml: string): XmlElement {
  const root = parseXml(`<root ${NS}>${xml}</root>`).root;
  return elements(root)[0]!;
}

const omath = (inner: string) => parse(`<m:oMath>${inner}</m:oMath>`);

function lay(inner: string, display = false, size = SIZE): MathLayout {
  const layout = layoutMath(omath(inner), { fonts, size, color: '000000', display });
  expectFinite(layout);
  return layout;
}

const text = (inner: string) => mathToText(omath(inner));

// --- Inspecting the display list --------------------------------------------------------------

interface Found {
  item: TextItem;
  x: number;
  w: number;
  y: number;
  size: number;
}

/** The nth occurrence of a character among the text items. */
function glyph(layout: MathLayout, ch: string, nth = 0): Found {
  let seen = 0;
  for (const item of layout.items) {
    if (item.type !== 'text') continue;
    for (let i = 0; i < item.chars.length; i++) {
      if (item.chars[i] !== ch) continue;
      if (seen++ === nth) return { item, x: item.xs[i]!, w: item.widths[i]!, y: item.y, size: item.size };
    }
  }
  throw new Error(`no ${ch} (#${nth}) in ${layout.items.map((i) => (i.type === 'text' ? i.text : i.type)).join(' | ')}`);
}

function allText(layout: MathLayout): string {
  return layout.items
    .filter((i): i is TextItem => i.type === 'text')
    .map((i) => i.text)
    .join('');
}

const paths = (layout: MathLayout) => layout.items.filter((i): i is PathItem => i.type === 'path');
const rects = (layout: MathLayout) => layout.items.filter((i): i is RectItem => i.type === 'rect');

function bounds(p: PathItem): { minX: number; maxX: number; minY: number; maxY: number } {
  const xs: number[] = [];
  const ys: number[] = [];
  for (const c of p.commands) {
    if (c.op === 'Z') continue;
    xs.push(c.x);
    ys.push(c.y);
  }
  return { minX: Math.min(...xs), maxX: Math.max(...xs), minY: Math.min(...ys), maxY: Math.max(...ys) };
}

function expectFinite(layout: MathLayout): void {
  const nums: number[] = [layout.width, layout.ascent, layout.descent];
  for (const it of layout.items) {
    if (it.type === 'text') nums.push(it.x, it.y, it.size, ...it.xs, ...it.widths);
    else if (it.type === 'rect') nums.push(it.x, it.y, it.w, it.h);
    else if (it.type === 'line') nums.push(it.x1, it.y1, it.x2, it.y2);
    else if (it.type === 'path') for (const c of it.commands) if (c.op !== 'Z') nums.push(c.x, c.y);
  }
  for (const n of nums) expect(Number.isFinite(n)).toBe(true);
}

const center = (g: Found) => g.x + g.w / 2;

// ----------------------------------------------------------------------------------------------

describe('math runs', () => {
  it('sets single-letter variables italic and digits and operators upright', () => {
    const l = lay(r('2x+α=y'));
    expect(glyph(l, 'x').item.font.italic).toBe(true);
    expect(glyph(l, 'y').item.font.italic).toBe(true);
    expect(glyph(l, 'α').item.font.italic).toBe(true);
    expect(glyph(l, '2').item.font.italic).toBe(false);
    expect(glyph(l, '+').item.font.italic).toBe(false);
    expect(glyph(l, '=').item.font.italic).toBe(false);
    // Word's math font is Cambria Math; the bundled stand-in is Caladea.
    expect(glyph(l, 'x').item.font.family).toBe('Caladea');
  });

  it('keeps known function names upright and honours m:sty', () => {
    expect(glyph(lay(r('sin')), 's').item.font.italic).toBe(false);
    expect(glyph(lay(r('ab')), 'a').item.font.italic).toBe(true);
    expect(glyph(lay(r('ab', sty('p'))), 'a').item.font.italic).toBe(false);
    const bold = glyph(lay(r('v', sty('b'))), 'v').item.font;
    expect(bold.bold).toBe(true);
    expect(bold.italic).toBe(false);
    const boldItalic = glyph(lay(r('v', sty('bi'))), 'v').item.font;
    expect(boldItalic.bold && boldItalic.italic).toBe(true);
  });

  it('sets normal text (m:nor) upright without operator spacing', () => {
    const l = lay(r('if x=1', '<m:rPr><m:nor/></m:rPr>'));
    expect(glyph(l, 'i').item.font.italic).toBe(false);
    expect(glyph(l, 'x').item.font.italic).toBe(false);
    const x = glyph(l, 'x');
    expect(glyph(l, '=').x).toBeCloseTo(x.x + x.w, 3);
  });

  it('decodes mathematical alphanumerics the fonts lack into styled letters', () => {
    const l = lay(r('\u{1D465}'));
    const x = glyph(l, 'x');
    expect(x.item.font.italic).toBe(true);
  });

  it('applies run colour and size from w:rPr', () => {
    const l = lay(r('x', '<w:rPr><w:color w:val="C00000"/><w:sz w:val="32"/></w:rPr>'));
    const x = glyph(l, 'x');
    expect(x.item.color).toBe('C00000');
    expect(x.size).toBe(16);
  });

  it('spaces relations thickly, binary operators medium, and unary signs not at all', () => {
    const rel = lay(r('x=1'));
    const x = glyph(rel, 'x');
    const eq = glyph(rel, '=');
    expect(eq.x - (x.x + x.w)).toBeCloseTo((5 / 18) * EM, 2);

    const bin = lay(r('a+b'));
    const a = glyph(bin, 'a');
    const plus = glyph(bin, '+');
    expect(plus.x - (a.x + a.w)).toBeCloseTo((4 / 18) * EM, 2);

    const unary = lay(r('-b'));
    const minus = glyph(unary, '−');
    expect(glyph(unary, 'b').x).toBeCloseTo(minus.x + minus.w, 3);
    expect(allText(unary)).not.toContain('-');

    const paren = lay(r('(-1)'));
    const open = glyph(paren, '(');
    expect(glyph(paren, '−').x).toBeCloseTo(open.x + open.w, 3);

    // After a relation a minus is unary too.
    const neg = lay(r('x=-1'));
    const m = glyph(neg, '−');
    expect(glyph(neg, '1').x).toBeCloseTo(m.x + m.w, 3);
  });

  it('drops operator spacing in scripts', () => {
    const l = lay(sSub(r('x'), r('i=1')));
    const i = glyph(l, 'i');
    expect(glyph(l, '=').x).toBeCloseTo(i.x + i.w, 3);
  });

  it('lays out an empty equation as an empty box', () => {
    const l = lay('');
    expect(l.width).toBe(0);
    expect(l.items).toHaveLength(0);
  });

  it('ignores bookmarks and deleted text, and keeps inserted text', () => {
    const l = lay(`<w:bookmarkStart w:id="1" w:name="eq"/>${r('a')}<w:del>${r('b')}</w:del><w:ins>${r('c')}</w:ins><m:ctrlPr/>`);
    expect(allText(l)).toBe('ac');
    expect(text(`${r('a')}<w:ins>${r('c')}</w:ins><w:del>${r('b')}</w:del>`)).toBe('ac');
  });

  it('works with fonts that have no font file (canvas measurement)', () => {
    const font: LayoutFont = {
      key: 'c0',
      family: 'serif',
      ascent: 0.9,
      descent: 0.22,
      underlinePosition: 0.1,
      underlineThickness: 0.05,
      advance: () => 0.5,
      glyphId: (cp) => cp,
      has: () => true,
      syntheticBold: false,
      syntheticItalic: false,
      bold: false,
      italic: false,
    };
    const canvas: FontProvider = { get: () => font, forChar: (f) => f };
    const xml = r('x=') + frac(r('-b±') + rad(sSup(r('b'), r('2')) + r('-4ac')), r('2a')) + nary('∑', r('i'), r('n'), r('a')) + delim([frac(r('1'), r('2'))]);
    const l = layoutMath(omath(xml), { fonts: canvas, size: SIZE, color: '000000', display: true });
    expectFinite(l);
    expect(l.width).toBeGreaterThan(0);
    expect(l.ascent).toBeGreaterThan(EM);
    expect(rects(l).length).toBeGreaterThan(0);
  });
});

describe('scripts', () => {
  it('x²: a smaller superscript above the baseline', () => {
    const l = lay(sSup(r('x'), r('2')));
    const x = glyph(l, 'x');
    const two = glyph(l, '2');
    expect(two.size).toBeCloseTo(SIZE * 0.73, 5);
    expect(two.y).toBeLessThan(0);
    expect(-two.y).toBeGreaterThan(0.3 * EM);
    expect(-two.y).toBeLessThan(0.45 * EM);
    expect(two.x).toBeGreaterThanOrEqual(x.x + x.w - 1e-6);
    expect(two.item.font.italic).toBe(false);
    expect(l.ascent).toBeGreaterThan(lay(r('x')).ascent);
    expect(text(sSup(r('x'), r('2')))).toBe('x^2');
  });

  it('x_i: a smaller subscript below the baseline', () => {
    const l = lay(sSub(r('x'), r('i')));
    const i = glyph(l, 'i');
    expect(i.size).toBeCloseTo(SIZE * 0.73, 5);
    expect(i.y).toBeGreaterThan(0);
    expect(i.item.font.italic).toBe(true);
    expect(l.descent).toBeGreaterThan(0);
    expect(text(sSub(r('x'), r('i')))).toBe('x_i');
  });

  it('x_i^2: both scripts, stacked with a gap', () => {
    const l = lay(sSubSup(r('x'), r('i'), r('2')));
    const x = glyph(l, 'x');
    const i = glyph(l, 'i');
    const two = glyph(l, '2');
    expect(two.y).toBeLessThan(0);
    expect(i.y).toBeGreaterThan(0);
    expect(i.x).toBeCloseTo(x.x + x.w, 3);
    // The superscript may move right by the italic correction, never left.
    expect(two.x).toBeGreaterThanOrEqual(i.x - 1e-6);
    // Bottom of the superscript stays clear of the top of the subscript.
    const subTop = i.y - 0.68 * i.size;
    expect(subTop - two.y).toBeGreaterThan(0.1 * EM);
    expect(text(sSubSup(r('x'), r('i'), r('2')))).toBe('x_i^2');
  });

  it('raises superscripts on tall bases further', () => {
    const plain = glyph(lay(sSup(r('x'), r('2'))), '2');
    const tall = glyph(lay(sSup(delim([frac(r('a'), r('b'))]), r('2')), true), '2');
    expect(tall.y).toBeLessThan(plain.y);
  });

  it('places prescripts left of the base', () => {
    const xml = `<m:sPre><m:sub>${r('1')}</m:sub><m:sup>${r('2')}</m:sup><m:e>${r('X')}</m:e></m:sPre>`;
    const l = lay(xml);
    const X = glyph(l, 'X');
    expect(glyph(l, '1').x).toBeLessThan(X.x);
    expect(glyph(l, '2').x).toBeLessThan(X.x);
    expect(glyph(l, '2').y).toBeLessThan(0);
    expect(glyph(l, '1').y).toBeGreaterThan(0);
  });
});

describe('fractions', () => {
  it('stacks numerator over denominator around a bar on the math axis', () => {
    const l = lay(frac(r('a'), r('b')), true);
    const plain = lay(r('a'), true);
    expect(l.ascent + l.descent).toBeGreaterThan(plain.ascent + plain.descent);
    const [bar] = rects(l);
    expect(bar).toBeDefined();
    const axis = -(bar!.y + bar!.h / 2);
    expect(axis).toBeGreaterThan(0.2 * EM);
    expect(axis).toBeLessThan(0.32 * EM);
    expect(bar!.h).toBeCloseTo(0.06 * EM, 5);
    const a = glyph(l, 'a');
    const b = glyph(l, 'b');
    expect(a.y).toBeLessThan(bar!.y);
    expect(b.y - 0.46 * EM).toBeGreaterThan(bar!.y + bar!.h);
    // Centred on the bar, and full size in display style.
    expect(center(a)).toBeCloseTo(bar!.x + bar!.w / 2, 5);
    expect(center(b)).toBeCloseTo(bar!.x + bar!.w / 2, 5);
    expect(bar!.w).toBeGreaterThan(Math.max(a.w, b.w));
    expect(a.size).toBe(SIZE);
    expect(text(frac(r('a'), r('b')))).toBe('a/b');
  });

  it('scales numerator and denominator down inline', () => {
    const l = lay(frac(r('1'), r('2')));
    expect(glyph(l, '1').size).toBeCloseTo(SIZE * 0.73, 5);
    expect(glyph(l, '2').size).toBeCloseTo(SIZE * 0.73, 5);
  });

  it('supports linear, skewed and bar-less fractions', () => {
    const lin = lay(frac(r('a'), r('b'), 'lin'));
    expect(allText(lin)).toBe('a/b');
    expect(glyph(lin, 'b').y).toBe(0);
    const skw = lay(frac(r('a'), r('b'), 'skw'));
    expect(glyph(skw, 'a').y).toBeLessThan(glyph(skw, 'b').y);
    expect(glyph(skw, 'a').x).toBeLessThan(glyph(skw, 'b').x);
    expect(paths(skw)).toHaveLength(1);
    const noBar = lay(frac(r('n'), r('k'), 'noBar'));
    expect(rects(noBar)).toHaveLength(0);
    expect(glyph(noBar, 'n').y).toBeLessThan(glyph(noBar, 'k').y);
    expect(text(frac(r('n'), r('k'), 'noBar'))).toBe('n¦k');
  });

  it('typesets the quadratic formula', () => {
    const xml = r('x=') + frac(r('-b±') + rad(sSup(r('b'), r('2')) + r('-4ac')), r('2a'));
    expect(text(xml)).toBe('x=(-b±√(b^2-4ac))/(2a)');
    const l = lay(xml, true);
    for (const ch of ['x', 'b', 'a', 'c']) expect(glyph(l, ch).item.font.italic).toBe(true);
    for (const ch of ['2', '4', '±', '=', '−']) expect(glyph(l, ch).item.font.italic).toBe(false);
    expect(allText(l)).not.toContain('-');
    // Radical sign drawn as a path, fraction bar as a rule.
    expect(paths(l)).toHaveLength(1);
    expect(rects(l)).toHaveLength(1);
    // "−b" is unary (tight); "±" is binary (medium space on both sides).
    const minus = glyph(l, '−', 0);
    const b = glyph(l, 'b', 0);
    expect(b.x).toBeCloseTo(minus.x + minus.w, 3);
    const pm = glyph(l, '±');
    expect(pm.x - (b.x + b.w)).toBeCloseTo((4 / 18) * EM, 2);
    // The numerator sits above the denominator.
    expect(glyph(l, 'b', 0).y).toBeLessThan(glyph(l, 'a', 1).y);
    expect(l.ascent).toBeGreaterThan(EM);
  });
});

describe('radicals', () => {
  it('√x: a sign drawn over and left of the radicand', () => {
    const l = lay(rad(r('x')));
    const x = glyph(l, 'x');
    const [sign] = paths(l);
    expect(sign).toBeDefined();
    const box = bounds(sign!);
    expect(box.minX).toBeLessThan(x.x);
    expect(box.maxX).toBeGreaterThanOrEqual(x.x + x.w);
    expect(box.minY).toBeLessThan(-0.47 * EM);
    expect(box.maxY).toBeGreaterThan(0);
    expect(l.width).toBeGreaterThan(x.w);
    expect(text(rad(r('x')))).toBe('√x');
  });

  it('∛x: a small degree tucked over the hook', () => {
    const l = lay(rad(r('x'), r('3')));
    const x = glyph(l, 'x');
    const three = glyph(l, '3');
    expect(three.size).toBeCloseTo(SIZE * 0.6, 5);
    expect(three.x + three.w).toBeLessThan(x.x);
    expect(three.y).toBeLessThan(0);
    expect(text(rad(r('x'), r('3')))).toBe('∛x');
    expect(text(rad(r('x+1'), r('n')))).toBe('√(n&x+1)');
  });

  it('grows with its content', () => {
    const small = bounds(paths(lay(rad(r('x')), true))[0]!);
    const big = bounds(paths(lay(rad(frac(r('a'), r('b'))), true))[0]!);
    expect(big.maxY - big.minY).toBeGreaterThan(small.maxY - small.minY);
  });
});

describe('n-ary operators', () => {
  const sum = nary('∑', r('i=1'), r('n'), r('a'));

  it('∑ in display style: a big operator with limits under and over', () => {
    const l = lay(sum, true);
    const sigma = glyph(l, '∑');
    expect(sigma.size).toBeGreaterThan(SIZE * 1.5);
    const n = glyph(l, 'n');
    const one = glyph(l, '1');
    // Caladea's ∑ ink spans 0..0.667 em from its baseline.
    const opTop = sigma.y - 0.667 * sigma.size;
    const opBottom = sigma.y;
    expect(n.y).toBeLessThan(opTop);
    expect(one.y - 0.67 * one.size).toBeGreaterThan(opBottom);
    const opCentre = sigma.x + ((0.135 + 0.578) / 2) * sigma.size;
    expect(Math.abs(center(n) - opCentre)).toBeLessThan(0.5);
    expect(n.size).toBeCloseTo(SIZE * 0.73, 5);
    // The summand follows the operator with a thin space.
    expect(glyph(l, 'a').x).toBeGreaterThan(sigma.x + 0.578 * sigma.size);
    expect(text(sum)).toBe('∑_(i=1)^n▒a');
  });

  it('∑ inline: a smaller operator with limits at its side', () => {
    const inline = lay(sum, false);
    const display = lay(sum, true);
    const sigma = glyph(inline, '∑');
    expect(sigma.size).toBeLessThan(glyph(display, '∑').size);
    const n = glyph(inline, 'n');
    const one = glyph(inline, '1');
    expect(n.x).toBeGreaterThan(sigma.x + 0.5 * sigma.size);
    expect(n.y).toBeLessThan(0);
    expect(one.y).toBeGreaterThan(0);
    expect(inline.ascent + inline.descent).toBeLessThan(display.ascent + display.descent);
  });

  it('∫ with limits at the side, the lower one tucked under the slant', () => {
    const xml = nary('∫', r('0'), r('1'), r('x'));
    const l = lay(xml, true);
    const [integral] = paths(l);
    expect(integral).toBeDefined();
    const box = bounds(integral!);
    expect(box.maxY - box.minY).toBeGreaterThan(2 * EM);
    const lower = glyph(l, '0');
    const upper = glyph(l, '1');
    expect(upper.y).toBeLessThan(0);
    expect(lower.y).toBeGreaterThan(0);
    expect(upper.x).toBeGreaterThan(lower.x);
    expect(lower.x).toBeGreaterThan(box.minX);
    expect(glyph(l, 'x').x).toBeGreaterThan(upper.x);
    expect(text(xml)).toBe('∫_0^1▒x');
    // Inline the integral is shorter.
    const inline = bounds(paths(lay(xml))[0]!);
    expect(inline.maxY - inline.minY).toBeLessThan(box.maxY - box.minY);
  });

  it('honours limLoc, hidden limits and drawn operators', () => {
    const side = lay(nary('∑', r('k'), r('n'), r('a'), '<m:limLoc m:val="subSup"/>'), true);
    expect(glyph(side, 'n').x).toBeGreaterThan(glyph(side, '∑').x + 0.5 * glyph(side, '∑').size);
    const hidden = lay(nary('∑', '', '', r('a'), '<m:subHide m:val="1"/><m:supHide m:val="1"/>'), true);
    expect(allText(hidden)).toBe('∑a');
    // ⋃ is not in the bundled fonts: drawn.
    const union = lay(nary('⋃', r('i'), '', r('A')), true);
    expect(paths(union).length).toBeGreaterThan(0);
    expect(allText(union)).not.toContain('⋃');
  });
});

describe('delimiters', () => {
  it('uses the font glyphs around short content', () => {
    const l = lay(delim([r('x+1')]));
    expect(glyph(l, '(').item.font.italic).toBe(false);
    expect(glyph(l, ')')).toBeDefined();
    expect(paths(l)).toHaveLength(0);
    expect(text(delim([r('x+1')]))).toBe('(x+1)');
  });

  it('stretches parentheses around a fraction', () => {
    const l = lay(delim([frac(r('1'), r('2'))]), true);
    const ps = paths(l);
    expect(ps).toHaveLength(2);
    const one = glyph(l, '1');
    const two = glyph(l, '2');
    const contentTop = one.y - 0.67 * one.size;
    const contentBottom = two.y;
    for (const p of ps) {
      const box = bounds(p);
      expect(box.minY).toBeLessThanOrEqual(contentTop);
      expect(box.maxY).toBeGreaterThanOrEqual(contentBottom);
    }
    const [open, close] = ps.map(bounds);
    expect(open!.maxX).toBeLessThanOrEqual(one.x);
    expect(close!.minX).toBeGreaterThanOrEqual(one.x + one.w);
    expect(text(delim([frac(r('1'), r('2'))]))).toBe('(1/2)');
  });

  it('draws delimiters the fonts lack, and separators between arguments', () => {
    const angle = lay(delim([r('a'), r('b')], '⟨', '⟩'));
    expect(paths(angle).length).toBeGreaterThanOrEqual(2);
    expect(glyph(angle, 'a').x).toBeLessThan(glyph(angle, 'b').x);
    expect(text(delim([r('a'), r('b')], '⟨', '⟩'))).toBe('⟨a|b⟩');
    const abs = lay(delim([r('x')], '|', '|'));
    expect(allText(abs)).toBe('|x|');
    const open = lay(delim([r('x')], '{', ''));
    expect(allText(open)).toBe('{x');
    expect(text(delim([r('x')], '{', ''))).toBe('{x┤');
  });

  it('sets a 2×2 matrix in stretched brackets', () => {
    const xml = delim([matrix([[r('a'), r('b')], [r('c'), r('d')]])], '[', ']');
    const l = lay(xml, true);
    const [a, b, c, d] = ['a', 'b', 'c', 'd'].map((ch) => glyph(l, ch));
    expect(a!.y).toBeCloseTo(b!.y, 5);
    expect(c!.y).toBeCloseTo(d!.y, 5);
    expect(a!.y).toBeLessThan(c!.y);
    expect(a!.x).toBeLessThan(b!.x);
    expect(c!.x).toBeLessThan(d!.x);
    expect(center(a!)).toBeCloseTo(center(c!), 5);
    // Centred on the math axis: one row above, one below.
    expect(a!.y).toBeLessThan(0);
    expect(c!.y).toBeGreaterThan(0);
    const brackets = paths(l).map(bounds);
    expect(brackets).toHaveLength(2);
    for (const box of brackets) {
      expect(box.minY).toBeLessThan(a!.y - 0.46 * EM);
      expect(box.maxY).toBeGreaterThan(c!.y);
    }
    expect(text(xml)).toBe('[■(a&b@c&d)]');
  });

  it('aligns matrix columns as m:mcJc says', () => {
    const xml = `<m:m><m:mPr><m:mcs><m:mc><m:mcPr><m:count m:val="1"/><m:mcJc m:val="left"/></m:mcPr></m:mc></m:mcs></m:mPr><m:mr><m:e>${r('1')}</m:e></m:mr><m:mr><m:e>${r('100')}</m:e></m:mr></m:m>`;
    const l = lay(xml);
    expect(glyph(l, '1', 0).x).toBeCloseTo(glyph(l, '1', 1).x, 5);
  });
});

describe('accents, bars, functions and limits', () => {
  it('draws a vector arrow above its base', () => {
    const xml = `<m:acc><m:accPr><m:chr m:val="⃗"/></m:accPr><m:e>${r('v')}</m:e></m:acc>`;
    const l = lay(xml);
    const v = glyph(l, 'v');
    const [arrow] = paths(l);
    expect(arrow).toBeDefined();
    const box = bounds(arrow!);
    expect(box.maxY).toBeLessThan(v.y - 0.44 * EM);
    expect(box.maxX - box.minX).toBeGreaterThanOrEqual(v.w * 0.9);
    expect(l.ascent).toBeGreaterThan(lay(r('v')).ascent);
    expect(text(xml)).toBe('v⃗');
  });

  it('draws a hat by default, stretching it over wide bases', () => {
    const narrow = bounds(paths(lay(`<m:acc><m:e>${r('x')}</m:e></m:acc>`))[0]!);
    const wide = bounds(paths(lay(`<m:acc><m:e>${r('xyz')}</m:e></m:acc>`))[0]!);
    expect(wide.maxX - wide.minX).toBeGreaterThan(2 * (narrow.maxX - narrow.minX));
  });

  it('draws over- and underbars (m:bar defaults to below)', () => {
    const over = lay(`<m:bar><m:barPr><m:pos m:val="top"/></m:barPr><m:e>${r('z')}</m:e></m:bar>`);
    expect(rects(over)[0]!.y).toBeLessThan(-0.46 * EM);
    const under = lay(`<m:bar><m:e>${r('z')}</m:e></m:bar>`);
    expect(rects(under)[0]!.y).toBeGreaterThan(0);
  });

  it('sin x: an upright name, a thin space, an italic argument', () => {
    const xml = func(r('sin', sty('p')), r('x'));
    const l = lay(xml);
    const n = glyph(l, 'n');
    const x = glyph(l, 'x');
    expect(glyph(l, 's').item.font.italic).toBe(false);
    expect(n.item.font.italic).toBe(false);
    expect(x.item.font.italic).toBe(true);
    expect(x.x - (n.x + n.w)).toBeCloseTo((3 / 18) * EM, 2);
    expect(text(xml)).toBe('sin⁡x');
    // No space before a parenthesised argument.
    const paren = lay(func(r('f'), delim([r('x')])));
    const f = glyph(paren, 'f');
    expect(f.item.font.italic).toBe(false);
    expect(glyph(paren, '(').x).toBeCloseTo(f.x + f.w, 3);
  });

  it('lim under: a small limit centred below an upright "lim"', () => {
    const xml = func(`<m:limLow><m:e>${r('lim', sty('p'))}</m:e><m:lim>${r('n→∞')}</m:lim></m:limLow>`, sSub(r('a'), r('n')));
    const l = lay(xml, true);
    const lim = glyph(l, 'l');
    const inf = glyph(l, '∞');
    const arrow = glyph(l, '→');
    expect(lim.item.font.italic).toBe(false);
    // Letters in the limit are variables again, not part of the function name.
    expect(glyph(l, 'n', 0).item.font.italic).toBe(true);
    expect(inf.size).toBeCloseTo(SIZE * 0.73, 5);
    expect(inf.y).toBeGreaterThan(0.35 * EM);
    const limCentre = (lim.x + glyph(l, 'm').x + glyph(l, 'm').w) / 2;
    expect(Math.abs(center(arrow) - limCentre)).toBeLessThan(0.15 * EM);
    expect(glyph(l, 'a').x).toBeGreaterThan(glyph(l, 'm').x);
    expect(text(xml)).toBe('lim┬(n→∞)⁡〖a_n〗');
  });

  it('draws an underbrace, and an upper limit above a base', () => {
    const brace = lay(`<m:groupChr><m:groupChrPr><m:chr m:val="⏟"/><m:pos m:val="bot"/><m:vertJc m:val="top"/></m:groupChrPr><m:e>${r('a+b+c')}</m:e></m:groupChr>`);
    const [p] = paths(brace);
    expect(p).toBeDefined();
    expect(bounds(p!).minY).toBeGreaterThan(0);
    expect(glyph(brace, 'a').y).toBe(0);
    const upp = lay(`<m:limUpp><m:e>${r('x')}</m:e><m:lim>${r('k')}</m:lim></m:limUpp>`);
    expect(glyph(upp, 'k').y).toBeLessThan(-0.46 * EM);
  });

  it('boxes, border boxes and phantoms', () => {
    const border = lay(`<m:borderBox><m:borderBoxPr><m:hideTop m:val="1"/></m:borderBoxPr><m:e>${r('x')}</m:e></m:borderBox>`);
    expect(border.items.filter((i) => i.type === 'line')).toHaveLength(3);
    const phantom = lay(`<m:phant><m:phantPr><m:show m:val="0"/></m:phantPr><m:e>${r('xx')}</m:e></m:phant>${r('y')}`);
    expect(allText(phantom)).toBe('y');
    expect(glyph(phantom, 'y').x).toBeGreaterThan(0);
    const smash = lay(`<m:phant><m:phantPr><m:zeroWid m:val="1"/></m:phantPr><m:e>${r('x')}</m:e></m:phant>${r('y')}`);
    expect(glyph(smash, 'y').x).toBe(0);
    expect(allText(lay(`<m:box><m:e>${r('dx')}</m:e></m:box>`))).toBe('dx');
  });
});

describe('equation arrays and math paragraphs', () => {
  it('aligns rows of an equation array at &', () => {
    const xml = `<m:eqArr><m:e>${r('x&=1')}</m:e><m:e>${r('2y&=3+z')}</m:e></m:eqArr>`;
    const l = lay(xml, true);
    const eq1 = glyph(l, '=', 0);
    const eq2 = glyph(l, '=', 1);
    expect(eq1.x).toBeCloseTo(eq2.x, 5);
    expect(eq2.y).toBeGreaterThan(eq1.y);
    const x = glyph(l, 'x');
    const y = glyph(l, 'y');
    expect(x.x + x.w).toBeCloseTo(y.x + y.w, 5);
    // The thick space before "=" survives the split at &.
    expect(eq1.x - (x.x + x.w)).toBeCloseTo((5 / 18) * EM, 2);
    expect(allText(l)).not.toContain('&');
    expect(text(xml)).toBe('█(x&=1@2y&=3+z)');
  });

  it('stacks the equations of an m:oMathPara, aligned by m:jc', () => {
    const para = (jc: string) =>
      parse(`<m:oMathPara><m:oMathParaPr><m:jc m:val="${jc}"/></m:oMathParaPr><m:oMath>${r('a=b')}</m:oMath><m:oMath>${r('c=d+e+f')}</m:oMath></m:oMathPara>`);
    const centred = layoutMath(para('center'), { fonts, size: SIZE, color: '000000', display: true });
    expectFinite(centred);
    const a = glyph(centred, 'a');
    const c = glyph(centred, 'c');
    expect(c.y).toBeGreaterThan(a.y + EM);
    expect(a.y).toBe(0);
    expect(c.x).toBeCloseTo(0, 5);
    const lineOne = lay(r('a=b'), true).width;
    expect(centred.width).toBeCloseTo(lay(r('c=d+e+f'), true).width, 5);
    expect(a.x).toBeCloseTo((centred.width - lineOne) / 2, 5);
    expect(centred.descent).toBeGreaterThan(EM);

    const left = layoutMath(para('left'), { fonts, size: SIZE, color: '000000', display: true });
    expect(glyph(left, 'a').x).toBeCloseTo(0, 5);
    const right = layoutMath(para('right'), { fonts, size: SIZE, color: '000000', display: true });
    expect(glyph(right, 'a').x).toBeCloseTo(right.width - lineOne, 5);

    expect(mathToText(para('center'))).toBe('a=b\nc=d+e+f');
  });
});
