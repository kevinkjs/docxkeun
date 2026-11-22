/**
 * OpenType shaping engine: lookup machinery on synthetic GSUB/GPOS/GDEF tables (built here
 * byte by byte and fed through `shape` with a fake font), script itemization, and real fonts:
 * the bundled Carlito / Caladea / Liberation fonts, plus Arabic and Devanagari fonts of this
 * Mac (skipped when missing).
 */
import { existsSync, readFileSync } from 'node:fs';
import { describe, expect, it } from 'vitest';
import { Font } from '../src/fonts/sfnt.js';
import { arabicJoining, FINA, INIT, ISOL, MEDI, NONE } from '../src/fonts/shaping/arabic.js';
import { newGlyphInfo } from '../src/fonts/shaping/buffer.js';
import { Gsub } from '../src/fonts/shaping/gsub.js';
import { findSyllables, indicProperties } from '../src/fonts/shaping/indic.js';
import { hasOpenTypeLayout, itemizeScripts, scriptOf, shape, type ShapedGlyph, type ShapeOptions, type ShapingFont } from '../src/fonts/shaping/index.js';
import { BinaryData, LayoutTable, parseClassDef, parseCoverage } from '../src/fonts/shaping/layout-common.js';
import { unicodeFlags } from '../src/fonts/shaping/unicode.js';
import { FONTS_DIR, readFont } from './helpers/fonts.js';

// ---------------------------------------------------------------------------------------
// A tiny OpenType table builder: numbers are 16-bit fields (negative ones int16), a nested
// Table is an Offset16 to it (laid out after its parent), null is a null offset.

type Item = number | Table | null | { tag: string } | { off32: Table } | { u32: number };

class Table {
  constructor(readonly items: Item[]) {}
}
const T = (...items: Item[]) => new Table(items);

function itemSize(item: Item): number {
  if (item !== null && typeof item === 'object' && !(item instanceof Table)) return 'tag' in item || 'off32' in item || 'u32' in item ? 4 : 2;
  return 2;
}

function serialize(table: Table): Uint8Array {
  const head = table.items.reduce<number>((n, item) => n + itemSize(item), 0);
  const children: Table[] = [];
  for (const item of table.items) {
    const child = item instanceof Table ? item : item && typeof item === 'object' && 'off32' in item ? item.off32 : undefined;
    if (child && !children.includes(child)) children.push(child);
  }
  const childBytes = children.map(serialize);
  const childAt = new Map<Table, number>();
  let size = head;
  children.forEach((c, i) => {
    childAt.set(c, size);
    size += childBytes[i]!.length;
  });
  const out = new Uint8Array(size);
  const view = new DataView(out.buffer);
  let p = 0;
  for (const item of table.items) {
    if (item === null) view.setUint16(p, 0);
    else if (typeof item === 'number') {
      if (item < 0) view.setInt16(p, item);
      else view.setUint16(p, item);
    } else if (item instanceof Table) view.setUint16(p, childAt.get(item)!);
    else if ('tag' in item) for (let k = 0; k < 4; k++) out[p + k] = item.tag.charCodeAt(k);
    else if ('off32' in item) view.setUint32(p, childAt.get(item.off32)!);
    else view.setUint32(p, item.u32);
    p += itemSize(item);
  }
  childBytes.forEach((b, i) => out.set(b, childAt.get(children[i]!)!));
  return out;
}

const cov1 = (...glyphs: number[]) => T(1, glyphs.length, ...glyphs);
const cov2 = (...ranges: Array<[number, number, number]>) => T(2, ranges.length, ...ranges.flat());
const anchor = (x: number, y: number) => T(1, x, y);
const lookup = (type: number, flag: number, ...subtables: Table[]) => T(type, flag, subtables.length, ...subtables);

/** A GSUB/GPOS table with one script (default LangSys enabling every feature). */
function layoutTable(features: Array<[string, number[]]>, lookups: Table[], scripts: string[] = ['DFLT']): Uint8Array {
  const langSys = T(null, 0xffff, features.length, ...features.map((_, i) => i));
  const script = T(langSys, 0);
  const scriptList = T(scripts.length, ...scripts.flatMap((tag) => [{ tag }, script]));
  const featureList = T(features.length, ...features.flatMap(([tag, lks]): Item[] => [{ tag }, T(null, lks.length, ...lks)]));
  return serialize(T(1, 0, scriptList, featureList, T(lookups.length, ...lookups)));
}

// ---------------------------------------------------------------------------------------
// A fake font: a..z → glyphs 1..26, space 27, U+0301 acute 30, U+0323 dot below 31.

const A = (ch: string) => ch.charCodeAt(0) - 96;
const SPACE = 27;
const ACUTE = 30;
const DOTBELOW = 31;

interface FakeFontOptions {
  cmap?: Map<number, number>;
  tables?: Record<string, Uint8Array>;
  advance?: (glyph: number) => number;
  kerning?: (left: number, right: number) => number;
}

function latinCmap(): Map<number, number> {
  const cmap = new Map<number, number>();
  for (let c = 0x61; c <= 0x7a; c++) cmap.set(c, c - 0x60);
  cmap.set(0x20, SPACE);
  cmap.set(0x301, ACUTE);
  cmap.set(0x323, DOTBELOW);
  return cmap;
}

function fakeFont(opts: FakeFontOptions = {}): ShapingFont {
  const cmap = opts.cmap ?? latinCmap();
  return {
    glyphId: (cp) => cmap.get(cp) ?? 0,
    advance: opts.advance ?? ((g) => (g === ACUTE || g === DOTBELOW ? 200 : g >= 100 ? 700 : 500)),
    tableBytes: (tag) => opts.tables?.[tag],
    kerning: opts.kerning ?? (() => 0),
    unitsPerEm: 1000,
  };
}

/** GDEF: letters are bases, 30/31 marks (attachment classes 1 and 2), 100+ ligatures; mark set 0 = {acute}. */
function latinGdef(): Uint8Array {
  const classes = T(2, 3, 1, 27, 1, 30, 31, 3, 100, 120, 2);
  const attach = T(2, 2, 30, 30, 1, 31, 31, 2);
  const markSets = T(1, 1, { off32: cov1(ACUTE) });
  return serialize(T(1, 2, classes, null, null, attach, markSets));
}

const ids = (glyphs: ShapedGlyph[]) => glyphs.map((g) => g.glyph);
const ltr: ShapeOptions = { direction: 'ltr', script: 'latn' };

// ---------------------------------------------------------------------------------------

describe('common tables', () => {
  it('parses Coverage formats 1 and 2', () => {
    // Wrapped in a parent so the table sits at offset 2 (offset 0 means "no table").
    const d1 = new BinaryData(serialize(T(cov1(5, 9, 12))));
    expect([...parseCoverage(d1, 2)]).toEqual([[5, 0], [9, 1], [12, 2]]);
    const d2 = new BinaryData(serialize(T(cov2([10, 12, 0], [20, 21, 3]))));
    expect([...parseCoverage(d2, 2)]).toEqual([[10, 0], [11, 1], [12, 2], [20, 3], [21, 4]]);
  });

  it('parses ClassDef formats 1 and 2', () => {
    const c1 = parseClassDef(new BinaryData(serialize(T(T(1, 10, 3, 1, 0, 2)))), 2);
    expect([c1.get(10), c1.get(11), c1.get(12), c1.get(13)]).toEqual([1, undefined, 2, undefined]);
    const c2 = parseClassDef(new BinaryData(serialize(T(T(2, 2, 5, 6, 4, 9, 9, 1)))), 2);
    expect([c2.get(5), c2.get(6), c2.get(7), c2.get(9)]).toEqual([4, 4, undefined, 1]);
  });

  it('selects scripts with DFLT fallback and prefers new Indic tags', () => {
    const table = new LayoutTable(layoutTable([['liga', []]], [], ['DFLT', 'dev2', 'deva', 'latn']), 7);
    expect(table.select(['dev2', 'deva'])!.scriptTag).toBe('dev2');
    expect(table.select(['arab'])!.scriptTag).toBe('DFLT');
    const latnOnly = new LayoutTable(layoutTable([['liga', []]], [], ['latn']), 7);
    expect(latnOnly.select(['thai'])!.scriptTag).toBe('latn');
    expect(latnOnly.select(['latn'], 'TRK ')!.langSys.features).toEqual([0]); // unknown language → default LangSys
  });
});

describe('GSUB', () => {
  // Lookup 0: 'f','f','i' → 101 and 'f','i' → 100 (IgnoreMarks); lookup 1: 'z' → first alternate;
  // lookup 2: 'q' → 'q','u'; lookup 3: 'k' → 'k' + 50 (delta, single format 1).
  const ligatures = lookup(4, 0x0008, T(1, cov1(A('f')), 1, T(2, T(101, 3, A('f'), A('i')), T(100, 2, A('i')))));
  const alternates = lookup(3, 0, T(1, cov1(A('z')), 1, T(2, 60, 61)));
  const multiple = lookup(2, 0, T(1, cov1(A('q')), 1, T(2, A('q'), A('u'))));
  const delta = lookup(1, 0, T(1, cov1(A('k')), 50));
  const font = fakeFont({
    tables: {
      GSUB: layoutTable([['liga', [0]], ['salt', [1]], ['ccmp', [2]], ['smcp', [3]]], [ligatures, alternates, multiple, delta]),
      GDEF: latinGdef(),
    },
  });

  it('forms ligatures (longest first) and lets features turn them off', () => {
    expect(ids(shape(font, 'fi ffi', ltr))).toEqual([100, SPACE, 101]);
    expect(ids(shape(font, 'fi', { ...ltr, features: { liga: false } }))).toEqual([A('f'), A('i')]);
    const lig = shape(font, 'xffiy', ltr);
    expect(lig.map((g) => g.cluster)).toEqual([0, 1, 4]);
    expect(lig[1]!.advance).toBe(700);
  });

  it('skips marks inside a ligature (IgnoreMarks) and keeps them after it', () => {
    const out = shape(font, 'f́i', ltr);
    expect(ids(out)).toEqual([100, ACUTE]);
    expect(out.map((g) => g.cluster)).toEqual([0, 0]);
  });

  it('applies multiple substitution with one cluster, alternates only when enabled', () => {
    const out = shape(font, 'aq', ltr);
    expect(ids(out)).toEqual([A('a'), A('q'), A('u')]);
    expect(out.map((g) => g.cluster)).toEqual([0, 1, 1]);
    expect(ids(shape(font, 'z', ltr))).toEqual([A('z')]);
    expect(ids(shape(font, 'z', { ...ltr, features: { salt: true } }))).toEqual([60]);
    expect(ids(shape(font, 'k', { ...ltr, features: { smcp: true } }))).toEqual([A('k') + 50]);
  });

  it('applies chaining context format 3 with backtrack, lookahead and a nested lookup', () => {
    // calt: x [a] b → a.alt (70), marks ignored while matching.
    const chain = lookup(6, 0x0008, T(3, 1, cov1(A('x')), 1, cov1(A('a')), 1, cov1(A('b')), 1, 0, 1));
    const single = lookup(1, 0, T(2, cov1(A('a')), 1, 70));
    const f = fakeFont({ tables: { GSUB: layoutTable([['calt', [0]]], [chain, single]), GDEF: latinGdef() } });
    expect(ids(shape(f, 'xab', ltr))).toEqual([A('x'), 70, A('b')]);
    expect(ids(shape(f, 'x́ab', ltr))).toEqual([A('x'), ACUTE, 70, A('b')]);
    expect(ids(shape(f, 'yab', ltr))).toEqual([A('y'), A('a'), A('b')]);
    expect(ids(shape(f, 'xac', ltr))).toEqual([A('x'), A('a'), A('c')]);
  });

  it('applies chaining context formats 1 and 2 and context formats 1 and 3', () => {
    // Lookup 0 (chain format 1): x a [b] → b becomes 71 (nested lookup 4 at sequence index 1).
    const chain1 = lookup(6, 0, T(1, cov1(A('a')), 1, T(1, T(1, A('x'), 2, A('b'), 0, 1, 1, 4))));
    // Lookup 1 (chain format 2): a vowel followed by 't' → vowel + 40 (nested lookup 5).
    const vowels = T(2, 5, A('a'), A('a'), 1, A('e'), A('e'), 1, A('i'), A('i'), 1, A('o'), A('o'), 1, A('u'), A('u'), 1);
    const ahead = T(2, 1, A('t'), A('t'), 1);
    const chain2 = lookup(6, 0, T(2, cov1(A('a'), A('e'), A('i'), A('o'), A('u')), T(1, 0, 0), vowels, ahead, 2, null, T(1, T(0, 1, 1, 1, 1, 0, 5))));
    // Lookup 2 (context format 1): 'g' 'h' → g becomes 72; lookup 3 (context format 3): 'm' 'n' → n becomes 73.
    const ctx1 = lookup(5, 0, T(1, cov1(A('g')), 1, T(1, T(2, 1, A('h'), 0, 6))));
    const ctx3 = lookup(5, 0, T(3, 2, 1, cov1(A('m')), cov1(A('n')), 1, 7));
    const toB = lookup(1, 0, T(2, cov1(A('b')), 1, 71));
    const plus40 = lookup(1, 0, T(1, cov2([1, 26, 0]), 40));
    const toG = lookup(1, 0, T(2, cov1(A('g')), 1, 72));
    const toN = lookup(1, 0, T(2, cov1(A('n')), 1, 73));
    const f = fakeFont({ tables: { GSUB: layoutTable([['calt', [0, 1, 2, 3]]], [chain1, chain2, ctx1, ctx3, toB, plus40, toG, toN]) } });
    expect(ids(shape(f, 'xab', ltr))).toEqual([A('x'), A('a'), 71]);
    expect(ids(shape(f, 'yab', ltr))).toEqual([A('y'), A('a'), A('b')]);
    expect(ids(shape(f, 'ate', ltr))).toEqual([A('a') + 40, A('t'), A('e')]);
    expect(ids(shape(f, 'gh', ltr))).toEqual([72, A('h')]);
    expect(ids(shape(f, 'mn nm', ltr))).toEqual([A('m'), 73, SPACE, A('n'), A('m')]);
  });

  it('applies reverse chaining single substitution from the end', () => {
    // 'a' followed by 'b' or by an already substituted 'a' (80) becomes 80.
    const reverse = lookup(8, 0, T(1, cov1(A('a')), 0, 1, cov1(A('b'), 80), 1, 80));
    const f = fakeFont({ tables: { GSUB: layoutTable([['calt', [0]]], [reverse]) } });
    expect(ids(shape(f, 'aaab', ltr))).toEqual([80, 80, 80, A('b')]);
    expect(ids(shape(f, 'aaac', ltr))).toEqual([A('a'), A('a'), A('a'), A('c')]);
  });

  it('honours mark filtering sets and mark attachment types', () => {
    // Ligature 'a','b' → 102 with UseMarkFilteringSet (set 0 = acute only): other marks are skipped.
    const filtered = lookup(4, 0x0010, T(1, cov1(A('a')), 1, T(1, T(102, 2, A('b')))));
    const withSet = T(...filtered.items, 0); // markFilteringSet index after the subtable offsets
    // Ligature 'c','d' → 103 with MarkAttachmentType 2: marks of other classes are skipped.
    const typed = lookup(4, 0x0200, T(1, cov1(A('c')), 1, T(1, T(103, 2, A('d')))));
    const f = fakeFont({ tables: { GSUB: layoutTable([['liga', [0, 1]]], [withSet, typed]), GDEF: latinGdef() } });
    expect(ids(shape(f, 'ạb', ltr))[0]).toBe(102);
    expect(ids(shape(f, 'áb', ltr))).toEqual([A('a'), ACUTE, A('b')]);
    expect(ids(shape(f, 'ćd', ltr))[0]).toBe(103);
    expect(ids(shape(f, 'c̣d', ltr))).toEqual([A('c'), DOTBELOW, A('d')]);
  });

  it('answers would-substitute queries for ligatures and context rules', () => {
    const gsub = new Gsub(layoutTable([['liga', [0]]], [ligatures]));
    expect(gsub.wouldSubstitute([0], [A('f'), A('i')], true)).toBe(true);
    expect(gsub.wouldSubstitute([0], [A('f'), A('f')], true)).toBe(false);
  });
});

describe('GPOS', () => {
  // Lookup 0: pair format 1 a+v -80 (IgnoreMarks); lookup 1: pair format 2 {t}×{o,e} -50 with
  // +10 x placement on the second glyph; lookup 2: single k (+10, +20, +30).
  const pair1 = lookup(2, 0x0008, T(1, cov1(A('a')), 4, 0, 1, T(1, A('v'), -80)));
  const classes1 = T(2, 1, A('t'), A('t'), 1);
  const classes2 = T(2, 2, A('e'), A('e'), 1, A('o'), A('o'), 1);
  const pair2 = lookup(2, 0, T(2, cov1(A('t')), 4, 1, classes1, classes2, 2, 2, 0, 0, 0, 0, 0, 0, -50, 10));
  const single = lookup(1, 0, T(1, cov1(A('k')), 7, 10, 20, 30));
  // Lookup 3: mark-to-base: acute (anchor 100,0) / dot below (anchor 100,0, class 1) on letters.
  const markArray = T(2, 0, anchor(100, 0), 1, anchor(100, 0));
  const letters = cov2([1, 26, 0]);
  const baseArray = T(26, ...Array.from({ length: 26 }, () => [anchor(250, 600), anchor(250, -50)]).flat());
  const markBase = lookup(4, 0, T(1, cov1(ACUTE, DOTBELOW), letters, 2, markArray, baseArray));
  // Lookup 4: mark-to-mark: acute on acute (mark2 anchor 100,300).
  const markMark = lookup(6, 0, T(1, cov1(ACUTE), cov1(ACUTE), 1, T(1, 0, anchor(100, 0)), T(1, anchor(100, 300))));
  // Lookup 5: cursive on 'c': entry (50,0), exit (450,100).
  const cursive = lookup(3, 0, T(1, cov1(A('c')), 1, anchor(50, 0), anchor(450, 100)));
  const gpos = layoutTable(
    [['kern', [0, 1, 2]], ['mark', [3]], ['mkmk', [4]], ['curs', [5]]],
    [pair1, pair2, single, markBase, markMark, cursive],
  );
  const font = fakeFont({ tables: { GPOS: gpos, GDEF: latinGdef() } });

  it('kerns with pair adjustment formats 1 and 2', () => {
    const av = shape(font, 'av', ltr);
    expect(av.map((g) => g.advance)).toEqual([420, 500]);
    expect(shape(font, 'av', { ...ltr, features: { kern: false } }).map((g) => g.advance)).toEqual([500, 500]);
    const to = shape(font, 'tote', ltr);
    expect(to.map((g) => g.advance)).toEqual([450, 500, 450, 500]);
    expect(to.map((g) => g.offsetX)).toEqual([0, 10, 0, 10]);
    expect(shape(font, 'ta', ltr).map((g) => g.advance)).toEqual([500, 500]); // class 0 pair: zero
  });

  it('skips ignored marks between a kerning pair', () => {
    expect(shape(font, 'áv', ltr)[0]!.advance).toBe(420);
  });

  it('applies single adjustment value records', () => {
    expect(shape(font, 'k', ltr)[0]).toMatchObject({ advance: 530, offsetX: 10, offsetY: 20 });
  });

  it('attaches marks to bases and to marks, with zero advance', () => {
    const out = shape(font, 'á́', ltr);
    expect(out[1]).toMatchObject({ glyph: ACUTE, advance: 0, offsetX: 250 - 100 - 500, offsetY: 600 });
    expect(out[2]).toMatchObject({ glyph: ACUTE, advance: 0, offsetX: 250 - 100 - 500, offsetY: 900 });
    const below = shape(font, 'ḅ', ltr);
    expect(below[1]).toMatchObject({ advance: 0, offsetX: -350, offsetY: -50 });
  });

  it('offsets marks in right-to-left runs for reversed drawing', () => {
    const out = shape(font, 'á', { direction: 'rtl', script: 'latn' });
    // Drawn reversed, the mark comes first; the base starts at the same pen position.
    expect(out[1]).toMatchObject({ advance: 0, offsetX: 150, offsetY: 600 });
  });

  it('connects cursive attachments', () => {
    const out = shape(font, 'cc', ltr);
    expect(out[0]!.advance).toBe(450);
    expect(out[1]).toMatchObject({ advance: 450, offsetX: -50, offsetY: 100 });
  });
});

describe('mark-to-ligature', () => {
  it('attaches each mark to the ligature component it followed', () => {
    const liga = lookup(4, 0x0008, T(1, cov1(A('f')), 1, T(1, T(100, 2, A('i')))));
    const markLig = lookup(5, 0, T(1, cov1(ACUTE), cov1(100), 1, T(1, 0, anchor(50, 0)), T(1, T(2, anchor(100, 500), anchor(400, 500)))));
    const font = fakeFont({
      tables: { GSUB: layoutTable([['liga', [0]]], [liga]), GPOS: layoutTable([['mark', [0]]], [markLig]), GDEF: latinGdef() },
    });
    const first = shape(font, 'f́i', ltr);
    expect(ids(first)).toEqual([100, ACUTE]);
    expect(first[1]!.offsetX).toBe(100 - 50 - 700);
    const second = shape(font, 'fí', ltr);
    expect(second[1]!.offsetX).toBe(400 - 50 - 700);
  });
});

describe('fallback without layout tables', () => {
  it('uses cmap, hmtx and legacy kerning, zeroing mark advances over their base', () => {
    const font = fakeFont({ kerning: (l, r) => (l === A('a') && r === A('v') ? -70 : 0) });
    expect(hasOpenTypeLayout(font)).toBe(false);
    const out = shape(font, 'av á', ltr);
    expect(ids(out)).toEqual([A('a'), A('v'), SPACE, A('a'), ACUTE]);
    expect(out.map((g) => g.advance)).toEqual([430, 500, 500, 500, 0]);
    expect(out[4]!.offsetX).toBe(-200);
    expect(shape(font, 'av', { ...ltr, features: { kern: false } })[0]!.advance).toBe(500);
  });

  it('hides default ignorables as zero-width spaces and mirrors brackets in RTL', () => {
    const cmap = latinCmap();
    cmap.set(0x28, 40);
    cmap.set(0x29, 41);
    const font = fakeFont({ cmap });
    const out = shape(font, 'a\u200db', ltr);
    expect(out[1]).toMatchObject({ glyph: SPACE, advance: 0, cluster: 0 });
    expect(ids(shape(font, '(a)', { direction: 'rtl', script: 'latn' }))).toEqual([41, A('a'), 40]);
  });

  it('composes base + mark when the font has the precomposed character', () => {
    const cmap = latinCmap();
    cmap.set(0xe9, 90);
    const out = shape(fakeFont({ cmap }), 'é', ltr);
    expect(out).toEqual([{ glyph: 90, cluster: 0, advance: 500, offsetX: 0, offsetY: 0 }]);
    // And decomposes a precomposed character the font lacks.
    expect(ids(shape(fakeFont(), 'á', ltr))).toEqual([A('a'), ACUTE]);
  });
});

// ---------------------------------------------------------------------------------------
// Arabic and Indic shapers on synthetic fonts (no system fonts needed).

describe('Arabic shaper', () => {
  it('analyses joining with transparent marks and joiners', () => {
    const actions = (text: string) => {
      const buf = [...text].map((c, i) => newGlyphInfo(c.codePointAt(0)!, i, unicodeFlags(c.codePointAt(0)!)));
      arabicJoining(buf);
      return buf.map((b) => b.cat);
    };
    expect(actions('سلام')).toEqual([INIT, MEDI, FINA, ISOL]);
    expect(actions('بِسْمِ')).toEqual([INIT, NONE, MEDI, NONE, FINA, NONE]);
    expect(actions('ب\u200cب')).toEqual([ISOL, NONE, ISOL]);
    expect(actions('ب\u200d')).toEqual([INIT, FINA]); // ZWJ is join-causing, so it takes a form too
    expect(actions('ܐܒ')).toEqual([ISOL, ISOL]); // Syriac Alaph (right-joining) then beth
  });

  // seen 10, lam 11, alef 12, meem 13, beh 14, yeh 15, teh 16, fatha 30; forms: fina +100,
  // medi +200, init +300; rlig: lam.medi/lam.init + alef.fina → 400/401.
  const letters: Array<[number, number]> = [[0x633, 10], [0x644, 11], [0x627, 12], [0x645, 13], [0x628, 14], [0x64a, 15], [0x62a, 16]];
  const cmap = new Map<number, number>([...letters, [0x64e, 30], [0x20, 27]]);
  const joining = cov2([10, 16, 0]);
  const form = (delta: number) => lookup(1, 0, T(1, joining, delta));
  const lamAlef = lookup(4, 0x0008, T(1, cov1(211, 311), 2, T(1, T(400, 2, 112)), T(1, T(401, 2, 112))));
  const gsub = layoutTable([['fina', [0]], ['medi', [1]], ['init', [2]], ['rlig', [3]]], [form(100), form(200), form(300), lamAlef], ['arab']);
  const markBase = lookup(4, 0, T(1, cov1(30), cov1(314), 1, T(1, 0, anchor(0, 0)), T(1, anchor(200, 700))));
  const gpos = layoutTable([['mark', [0]]], [markBase], ['arab']);
  const gdef = serialize(T(1, 0, T(2, 4, 10, 29, 1, 30, 30, 3, 31, 399, 1, 400, 401, 2), null, null, null));
  const font = fakeFont({ cmap, tables: { GSUB: gsub, GPOS: gpos, GDEF: gdef }, advance: (g) => (g === 30 ? 0 : 600) });
  const rtl: ShapeOptions = { direction: 'rtl' };

  it('applies joining forms and the lam-alef ligature', () => {
    const salam = shape(font, 'سلام', rtl);
    expect(ids(salam)).toEqual([310, 400, 13]);
    expect(salam.map((g) => g.cluster)).toEqual([0, 1, 3]);
    expect(ids(shape(font, 'لا', rtl))).toEqual([401]);
    expect(ids(shape(font, 'بيت', rtl))).toEqual([314, 215, 116]);
    expect(ids(shape(font, 'بيت', { ...rtl, features: { init: false } }))).toEqual([14, 215, 116]);
  });

  it('positions harakat on the joined letter', () => {
    const out = shape(font, 'بَب', rtl);
    expect(ids(out)).toEqual([314, 30, 114]);
    expect(out[1]).toMatchObject({ advance: 0, offsetX: 200, offsetY: 700, cluster: 0 });
  });
});

describe('Indic shaper', () => {
  it('classifies characters and finds syllables', () => {
    expect(indicProperties(0x0915)).toEqual([1, 4]); // consonant, base
    expect(indicProperties(0x093f)[1]).toBe(2); // i-matra: pre-base
    const cats = [...'नमस्ते'].map((c) => indicProperties(c.codePointAt(0)!)[0]);
    expect(findSyllables(cats).map(([s, e]) => [s, e])).toEqual([[0, 1], [1, 2], [2, 6]]);
  });

  // ka 1, nga 2, ta 3, ra 4, virama 5, i-matra 6, ssa 8; reph 200, ka.half 201, ksha 202.
  const cmap = new Map<number, number>([[0x915, 1], [0x919, 2], [0x924, 3], [0x930, 4], [0x94d, 5], [0x93f, 6], [0x937, 8], [0x20, 27]]);
  const akhn = lookup(4, 0, T(1, cov1(1), 1, T(1, T(202, 3, 5, 8))));
  const rphf = lookup(4, 0, T(1, cov1(4), 1, T(1, T(200, 2, 5))));
  const half = lookup(4, 0, T(1, cov1(1), 1, T(1, T(201, 2, 5))));
  const gsub = layoutTable([['akhn', [0]], ['rphf', [1]], ['half', [2]]], [akhn, rphf, half], ['dev2']);
  const font = fakeFont({ cmap, tables: { GSUB: gsub } });
  const deva: ShapeOptions = { direction: 'ltr' };

  it('forms a reph and moves it after the base, with the i-matra before it', () => {
    expect(ids(shape(font, 'र्क', deva))).toEqual([1, 200]);
    const out = shape(font, 'र्कि', deva);
    expect(ids(out)).toEqual([6, 1, 200]);
    expect(out.every((g) => g.cluster === 0)).toBe(true);
  });

  it('places the pre-base matra before half forms, or after a standalone halant', () => {
    expect(ids(shape(font, 'क्ति', deva))).toEqual([6, 201, 3]);
    expect(ids(shape(font, 'ङ्ति', deva))).toEqual([2, 5, 6, 3]);
    expect(ids(shape(font, 'क्षि', deva))).toEqual([6, 202]);
  });
});

describe('scripts', () => {
  it('maps code points to OpenType script tags', () => {
    const tags = ['a', 'ب', 'א', 'क', 'ক', 'ਕ', 'ક', 'କ', 'க', 'క', 'ಕ', 'ക', 'ก', 'Ж', 'λ', '中', 'あ', 'カ', '한', '1', ' ', '́'];
    expect(tags.map((c) => scriptOf(c.codePointAt(0)!))).toEqual([
      'latn', 'arab', 'hebr', 'deva', 'beng', 'guru', 'gujr', 'orya', 'taml', 'telu', 'knda', 'mlym', 'thai', 'cyrl', 'grek', 'hani', 'kana', 'kana', 'hang', 'DFLT', 'DFLT', 'DFLT',
    ]);
  });

  it('itemizes text with neutral characters joining the surrounding run', () => {
    expect(itemizeScripts('Hello שלום world')).toEqual([
      { start: 0, end: 6, script: 'latn' },
      { start: 6, end: 11, script: 'hebr' },
      { start: 11, end: 16, script: 'latn' },
    ]);
    expect(itemizeScripts('123 abc')).toEqual([{ start: 0, end: 7, script: 'latn' }]);
    expect(itemizeScripts('\u00e9, नमस्ते!')).toEqual([
      { start: 0, end: 3, script: 'latn' },
      { start: 3, end: 10, script: 'deva' },
    ]);
    // A closing bracket follows its opening partner's script, not the run it ends up in.
    expect(itemizeScripts('שלום (abc) x')).toEqual([
      { start: 0, end: 6, script: 'hebr' },
      { start: 6, end: 9, script: 'latn' },
      { start: 9, end: 11, script: 'hebr' },
      { start: 11, end: 12, script: 'latn' },
    ]);
    expect(itemizeScripts('')).toEqual([]);
  });
});

// ---------------------------------------------------------------------------------------
// Real fonts.

const bundled = (file: string) => new Font(readFont(new URL(file, FONTS_DIR)));

/** First face of a font file (reads the .ttc header itself). */
function loadFace(path: string): Font {
  const data = new Uint8Array(readFileSync(path));
  const view = new DataView(data.buffer, data.byteOffset, data.byteLength);
  const offset = view.getUint32(0) === 0x74746366 ? view.getUint32(12) : 0;
  return new Font(data, offset);
}

/** The glyph a font's single-substitution feature maps `glyph` to (or the glyph itself). */
function featureForm(font: Font, script: string, feature: string, glyph: number): number {
  const gsub = new Gsub(font.tableBytes('GSUB')!);
  const sel = gsub.layout.select([script])!;
  const fi = gsub.layout.findFeature(sel.langSys, feature);
  for (const li of fi >= 0 ? gsub.layout.featureLookups(fi) : []) {
    for (const st of gsub.lookup(li)?.subtables ?? []) {
      if (st.kind !== 'single') continue;
      const ci = st.coverage.get(glyph);
      if (ci !== undefined) return st.delta !== undefined ? (glyph + st.delta) & 0xffff : st.substitutes[ci]!;
    }
  }
  return glyph;
}

describe('bundled fonts', () => {
  const carlito = bundled('Carlito-Regular.ttf');
  const liberation = bundled('LiberationSans-Regular.ttf');
  const caladea = bundled('Caladea-Regular.ttf');
  const gid = (font: Font, ch: string) => font.glyphId(ch.codePointAt(0)!);

  it('has OpenType layout', () => {
    expect(hasOpenTypeLayout(carlito)).toBe(true);
    expect(hasOpenTypeLayout(liberation)).toBe(true);
  });

  it('Carlito forms fi / ffi ligatures (liga) unless turned off', () => {
    const fi = shape(carlito, 'fi', ltr);
    expect(fi).toHaveLength(1);
    expect(fi[0]!.glyph).not.toBe(gid(carlito, 'f'));
    const ffi = shape(carlito, 'office', ltr);
    expect(ffi).toHaveLength(4);
    expect(ffi.map((g) => g.cluster)).toEqual([0, 1, 4, 5]);
    expect(ids(shape(carlito, 'office', { ...ltr, features: { liga: false } }))).toEqual([...'office'].map((c) => gid(carlito, c)));
  });

  it('kerns AV and To through GPOS (Carlito, Caladea, Liberation Sans)', () => {
    for (const font of [carlito, caladea, liberation]) {
      for (const pair of ['AV', 'To']) {
        const kerned = shape(font, pair, ltr);
        const plain = shape(font, pair, { ...ltr, features: { kern: false } });
        expect(plain[0]!.advance).toBe(font.advance(gid(font, pair[0]!)));
        expect(kerned[0]!.advance).toBeLessThan(plain[0]!.advance);
      }
    }
  });

  it('Liberation Sans positions Hebrew points as marks', () => {
    const out = shape(liberation, 'שָׁלוֹם', { direction: 'rtl' });
    const qamats = out.find((g) => g.glyph === gid(liberation, 'ָ'))!;
    expect(qamats).toMatchObject({ advance: 0, cluster: 0 });
    expect(qamats.offsetX).not.toBe(0);
    // Shin + shin dot (reordered before the vowel) compose through the font's ccmp.
    expect(out[0]!.glyph).not.toBe(gid(liberation, 'ש'));
    expect(out.filter((g) => g.advance > 0)).toHaveLength(4);
  });

  it('positions a Latin combining mark with mark-to-base', () => {
    const out = shape(liberation, 'q́', ltr);
    expect(out[1]).toMatchObject({ glyph: gid(liberation, '́'), advance: 0 });
    expect(out[1]!.offsetY !== 0 || out[1]!.offsetX !== 0).toBe(true);
  });
});

const SUPPLEMENTAL = '/System/Library/Fonts/Supplemental';
const arabicFontPath = [`${SUPPLEMENTAL}/Times New Roman.ttf`, `${SUPPLEMENTAL}/Arial.ttf`, `${SUPPLEMENTAL}/Tahoma.ttf`].find(existsSync);
const devanagariFontPath = ['/System/Library/Fonts/Kohinoor.ttc', `${SUPPLEMENTAL}/ITFDevanagari.ttc`, `${SUPPLEMENTAL}/Devanagari Sangam MN.ttc`].find(existsSync);
const nastaliqPath = '/System/Library/Fonts/NotoNastaliq.ttc';

// A skipped describe still runs its body to collect the tests: load a system font only where it exists.
describe.skipIf(!arabicFontPath)('system Arabic font', () => {
  const font = arabicFontPath ? loadFace(arabicFontPath) : (undefined as unknown as Font);
  const rtl: ShapeOptions = { direction: 'rtl' };
  const gid = (ch: string) => font.glyphId(ch.codePointAt(0)!);

  it('shapes سلام: initial seen, lam-alef ligature, isolated meem', () => {
    const out = shape(font, 'سلام', rtl);
    expect(out).toHaveLength(3);
    expect(out[0]!.glyph).toBe(featureForm(font, 'arab', 'init', gid('س')));
    expect(out[0]!.glyph).not.toBe(gid('س'));
    expect(out[1]!.glyph).not.toBe(gid('ل'));
    expect(out[1]!.glyph).not.toBe(gid('ا'));
    expect(out.map((g) => g.cluster)).toEqual([0, 1, 3]);
    expect(out[2]!.glyph).toBe(gid('م'));
    const lamAlef = shape(font, 'لا', rtl);
    expect(lamAlef).toHaveLength(1);
    expect(lamAlef[0]!.glyph).not.toBe(gid('ل'));
  });

  it('gives beh its isolated, initial, medial and final forms', () => {
    const out = shape(font, 'ب ببب', rtl);
    const beh = gid('ب');
    expect(ids(out)).toEqual([
      beh,
      gid(' '),
      featureForm(font, 'arab', 'init', beh),
      featureForm(font, 'arab', 'medi', beh),
      featureForm(font, 'arab', 'fina', beh),
    ]);
    expect(out[2]!.glyph).not.toBe(beh);
    expect(out[3]!.glyph).not.toBe(beh);
    if (arabicFontPath!.includes('Times')) expect(new Set(ids(out)).size).toBe(5); // all four forms distinct
  });

  it('attaches harakat as zero-width marks', () => {
    const out = shape(font, 'بِسْمِ', rtl);
    const marks = out.filter((g) => [gid('ِ'), gid('ْ')].includes(g.glyph));
    expect(marks).toHaveLength(3);
    for (const m of marks) expect(m.advance).toBe(0);
    if (!arabicFontPath!.includes('Tahoma')) expect(marks.some((m) => m.offsetX !== 0 || m.offsetY !== 0)).toBe(true);
    expect(out[2]!.glyph).toBe(featureForm(font, 'arab', 'medi', gid('س')));
  });
});

describe.skipIf(!existsSync(nastaliqPath))('system Nastaliq font', () => {
  it('raises glyphs through cursive attachment', () => {
    const out = shape(loadFace(nastaliqPath), 'سلام', { direction: 'rtl' });
    expect(out.some((g) => g.offsetY !== 0)).toBe(true);
    expect(out.every((g) => Number.isFinite(g.advance))).toBe(true);
  });
});

describe.skipIf(!devanagariFontPath)('system Devanagari font', () => {
  const font = devanagariFontPath ? loadFace(devanagariFontPath) : (undefined as unknown as Font);
  const deva: ShapeOptions = { direction: 'ltr' };
  const gid = (ch: string) => font.glyphId(ch.codePointAt(0)!);

  it('moves the i-matra of क्षि before the conjunct', () => {
    const out = shape(font, 'क्षि', deva);
    const ksha = shape(font, 'क्ष', deva);
    expect(ksha).toHaveLength(1);
    expect(out).toHaveLength(2);
    expect(out[1]!.glyph).toBe(ksha[0]!.glyph);
    expect(out[0]!.glyph).not.toBe(ksha[0]!.glyph);
    expect(out.every((g) => g.cluster === 0)).toBe(true);
    // Same for a plain consonant: the matra glyph comes first.
    expect(ids(shape(font, 'कि', deva))[1]).toBe(gid('क'));
  });

  it('forms a reph from र्क after the base consonant', () => {
    const out = shape(font, 'र्क', deva);
    expect(out).toHaveLength(2);
    expect(out[0]!.glyph).toBe(gid('क'));
    expect([gid('र'), gid('्')]).not.toContain(out[1]!.glyph);
    expect(out[1]!.advance).toBe(0);
  });

  it('forms a half form before ZWJ', () => {
    const out = shape(font, 'क्\u200dत', deva);
    const gsub = new Gsub(font.tableBytes('GSUB')!);
    const sel = gsub.layout.select(['dev2', 'deva'])!;
    const half = gsub.layout.featureLookups(gsub.layout.findFeature(sel.langSys, 'half'));
    expect(gsub.wouldSubstitute(half, [gid('क'), gid('्')], false)).toBe(true);
    expect(out[0]!.glyph).not.toBe(gid('क'));
    expect(out[1]).toMatchObject({ advance: 0 }); // the hidden joiner
    expect(out[out.length - 1]!.glyph).toBe(gid('त'));
  });
});
