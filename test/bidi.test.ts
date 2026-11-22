/**
 * Tests for the Unicode Bidirectional Algorithm (src/text/bidi.ts), in the style of
 * BidiCharacterTest.txt: input text, paragraph direction, then the expected resolved paragraph
 * level, per-code-unit levels and visual order. Every expectation was worked out by hand from the
 * UAX #9 rules; the comments name the rules that decide each case.
 *
 * Notation: uppercase ASCII letters stand for Hebrew letters (strong R) via he(), or Arabic letters
 * (strong AL) via ar(), as in the UAX #9 examples. `display` is the line as drawn left to right,
 * mapped back to that notation, with explicit controls dropped and mirroring (L4) applied.
 */
import { describe, expect, it } from 'vitest';
import { bidiClass, isRtlChar, lineLevels, mirrorChar, resolveLevels, visualOrder } from '../src/text/bidi.js';

const he = (s: string) => s.replace(/[A-Z]/g, (c) => String.fromCharCode(0x05d0 + c.charCodeAt(0) - 65));
const ar = (s: string) => s.replace(/[A-Z]/g, (c) => String.fromCharCode(0x0627 + c.charCodeAt(0) - 65));

const LRE = '‪';
const RLE = '‫';
const PDF = '‬';
const LRO = '‭';
const RLO = '‮';
const LRI = '⁦';
const RLI = '⁧';
const FSI = '⁨';
const PDI = '⁩';

function display(text: string, levels: ArrayLike<number>, order: readonly number[]): string {
  let out = '';
  for (const i of order) {
    const unit = text.charCodeAt(i);
    if ((unit >= 0x202a && unit <= 0x202e) || (unit >= 0x2066 && unit <= 0x2069)) continue;
    out += String.fromCharCode((levels[i]! & 1 ? mirrorChar(unit) : undefined) ?? unit);
  }
  return out
    .replace(/[א-ש]/g, (c) => String.fromCharCode(c.charCodeAt(0) - 0x05d0 + 65))
    .replace(/[ا-ـ]/g, (c) => String.fromCharCode(c.charCodeAt(0) - 0x0627 + 65));
}

function bidi(text: string, paragraph: 0 | 1 | 'auto') {
  const { levels, paragraphLevel } = resolveLevels(text, paragraph);
  const order = visualOrder(levels, 0, text.length);
  return { paragraphLevel, levels: [...levels], order, display: display(text, levels, order) };
}

interface Case {
  name: string;
  text: string;
  paragraph: 0 | 1 | 'auto';
  /** Resolved paragraph level. */
  level: 0 | 1;
  /** One digit per UTF-16 code unit. */
  levels: string;
  order: number[];
  display: string;
}

function check(cases: Case[]) {
  for (const c of cases) {
    it(c.name, () => {
      const result = bidi(c.text, c.paragraph);
      expect(result.paragraphLevel).toBe(c.level);
      expect(result.levels).toEqual([...c.levels].map(Number));
      expect(result.order).toEqual(c.order);
      expect(result.display).toBe(c.display);
    });
  }
}

describe('bidi: strong text', () => {
  check([
    { name: 'pure Hebrew finds RTL by itself (P2)', text: he('ABC'), paragraph: 'auto', level: 1, levels: '111', order: [2, 1, 0], display: 'CBA' },
    { name: 'pure Hebrew in an LTR paragraph is raised to level 1 (I1)', text: he('ABC'), paragraph: 0, level: 0, levels: '111', order: [2, 1, 0], display: 'CBA' },
    { name: 'spaces between Hebrew words stay in the run (N1)', text: he('ABC DEF'), paragraph: 'auto', level: 1, levels: '1111111', order: [6, 5, 4, 3, 2, 1, 0], display: 'FED CBA' },
    {
      // Spaces between R and L take the embedding direction (N2); L at an odd level is 2 (I2).
      name: 'Hebrew sentence with an English word',
      text: he('ABC def GHI'),
      paragraph: 'auto',
      level: 1,
      levels: '11112221111',
      order: [10, 9, 8, 7, 4, 5, 6, 3, 2, 1, 0],
      display: 'IHG def CBA',
    },
    {
      name: 'English sentence with a Hebrew word (UAX #9 "car means CAR.")',
      text: he('car means CAR.'),
      paragraph: 'auto',
      level: 0,
      levels: '00000000001110',
      order: [0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 12, 11, 10, 13],
      display: 'car means RAC.',
    },
    {
      name: 'Hebrew sentence ending in an English word, final full stop goes left',
      text: he('CAR MEANS car.'),
      paragraph: 'auto',
      level: 1,
      levels: '11111111112221',
      order: [13, 10, 11, 12, 9, 8, 7, 6, 5, 4, 3, 2, 1, 0],
      display: '.car SNAEM RAC',
    },
    {
      name: 'UAX #9 "car MEANS CAR." in an RTL paragraph',
      text: he('car MEANS CAR.'),
      paragraph: 1,
      level: 1,
      levels: '22211111111111',
      order: [13, 12, 11, 10, 9, 8, 7, 6, 5, 4, 3, 0, 1, 2],
      display: '.RAC SNAEM car',
    },
    {
      // W2/W7 keep the digits EN after R; comma+space between numbers resolve to R (N1).
      name: 'UAX #9 "he said THE VALUES ARE 123, 456, 789, OK"',
      text: he('he said "THE VALUES ARE 123, 456, 789, OK".'),
      paragraph: 'auto',
      level: 0,
      levels: '0000000001111111111111112221122211222111100',
      order: [
        0, 1, 2, 3, 4, 5, 6, 7, 8, 40, 39, 38, 37, 34, 35, 36, 33, 32, 29, 30, 31, 28, 27, 24, 25, 26, 23, 22,
        21, 20, 19, 18, 17, 16, 15, 14, 13, 12, 11, 10, 9, 41, 42,
      ],
      display: 'he said "KO ,789 ,456 ,123 ERA SEULAV EHT".',
    },
    {
      name: 'neutrals between RTL text in an LTR paragraph join it (N1)',
      text: he('ABC, DEF'),
      paragraph: 0,
      level: 0,
      levels: '11111111',
      order: [7, 6, 5, 4, 3, 2, 1, 0],
      display: 'FED ,CBA',
    },
    {
      name: 'neutrals between LTR words in an RTL paragraph join them (N1)',
      text: he('A b, c D'),
      paragraph: 'auto',
      level: 1,
      levels: '11222211',
      order: [7, 6, 2, 3, 4, 5, 1, 0],
      display: 'D b, c A',
    },
    { name: 'an explicit paragraph level wins over the text', text: 'abc', paragraph: 1, level: 1, levels: '222', order: [0, 1, 2], display: 'abc' },
  ]);
});

describe('bidi: numbers', () => {
  check([
    {
      // W2 turns EN after AL into AN; AN at an odd level is 2.
      name: 'Arabic with European digits',
      text: ar('A 123 B'),
      paragraph: 'auto',
      level: 1,
      levels: '1122211',
      order: [6, 5, 2, 3, 4, 1, 0],
      display: 'B 123 A',
    },
    {
      name: 'Arabic with Arabic-Indic digits (AN)',
      text: ar('A ') + '١٢٣' + ar(' B'),
      paragraph: 'auto',
      level: 1,
      levels: '1122211',
      order: [6, 5, 2, 3, 4, 1, 0],
      display: 'B ١٢٣ A',
    },
    {
      // Extended Arabic-Indic digits are EN, so with no strong text W7 makes them L.
      name: 'Extended Arabic-Indic digits alone are EN (level 0)',
      text: '۱۲۳',
      paragraph: 'auto',
      level: 0,
      levels: '000',
      order: [0, 1, 2],
      display: '۱۲۳',
    },
    {
      name: 'Arabic-Indic digits alone are AN (level 2) and not strong (P2)',
      text: '١٢٣',
      paragraph: 'auto',
      level: 0,
      levels: '222',
      order: [0, 1, 2],
      display: '١٢٣',
    },
    {
      // After AL both numbers are AN; W4 joins only CS between AN, so the ES stays neutral and
      // resolves to R: the two numbers swap places.
      name: '"123-456" inside Arabic',
      text: ar('A 123-456 B'),
      paragraph: 'auto',
      level: 1,
      levels: '11222122211',
      order: [10, 9, 6, 7, 8, 5, 2, 3, 4, 1, 0],
      display: 'B 456-123 A',
    },
    {
      name: '"123-456" inside Hebrew stays one number (W4)',
      text: he('A 123-456 B'),
      paragraph: 'auto',
      level: 1,
      levels: '11222222211',
      order: [10, 9, 2, 3, 4, 5, 6, 7, 8, 1, 0],
      display: 'B 123-456 A',
    },
    {
      name: '"$100.00" in Hebrew: ET and CS join the number (W4, W5)',
      text: he('A $100.00'),
      paragraph: 'auto',
      level: 1,
      levels: '112222222',
      order: [2, 3, 4, 5, 6, 7, 8, 1, 0],
      display: '$100.00 A',
    },
    {
      name: '"$100.00" in Arabic: ET next to AN stays neutral',
      text: ar('A $100.00'),
      paragraph: 'auto',
      level: 1,
      levels: '111222222',
      order: [3, 4, 5, 6, 7, 8, 2, 1, 0],
      display: '100.00$ A',
    },
    { name: '"1,234" in Hebrew (W4, CS between EN)', text: he('A 1,234'), paragraph: 'auto', level: 1, levels: '1122222', order: [2, 3, 4, 5, 6, 1, 0], display: '1,234 A' },
    { name: '"1,234" in Arabic (W4, CS between AN)', text: ar('A 1,234'), paragraph: 'auto', level: 1, levels: '1122222', order: [2, 3, 4, 5, 6, 1, 0], display: '1,234 A' },
    { name: '"50%" in Hebrew (W5)', text: he('A 50%'), paragraph: 'auto', level: 1, levels: '11222', order: [2, 3, 4, 1, 0], display: '50% A' },
    { name: '"50%" in Arabic: the percent sign goes left', text: ar('A 50%'), paragraph: 'auto', level: 1, levels: '11221', order: [4, 2, 3, 1, 0], display: '%50 A' },
    { name: '"1+2" in Hebrew (W4, ES between EN)', text: he('A 1+2'), paragraph: 'auto', level: 1, levels: '11222', order: [2, 3, 4, 1, 0], display: '1+2 A' },
    { name: 'two separators do not join numbers (W4 needs a single one)', text: he('A 1++2'), paragraph: 'auto', level: 1, levels: '112112', order: [5, 4, 3, 2, 1, 0], display: '2++1 A' },
    { name: 'numbers separated by spaces in RTL keep RTL order', text: he('A 1 2 B'), paragraph: 'auto', level: 1, levels: '1121211', order: [6, 5, 4, 3, 2, 1, 0], display: 'B 2 1 A' },
    { name: 'a number between Hebrew words in an LTR paragraph', text: he('A 1 B'), paragraph: 0, level: 0, levels: '11211', order: [4, 3, 2, 1, 0], display: 'B 1 A' },
    { name: 'European digits after Arabic in an LTR paragraph become AN', text: 'a ' + ar('B') + ' 12', paragraph: 'auto', level: 0, levels: '001122', order: [0, 1, 4, 5, 3, 2], display: 'a 12 B' },
    { name: 'digits without strong text keep an auto paragraph LTR', text: '123', paragraph: 'auto', level: 0, levels: '000', order: [0, 1, 2], display: '123' },
  ]);
});

describe('bidi: paired brackets (N0)', () => {
  check([
    {
      name: '"a (b C) d": content matches the LTR embedding (N0 b)',
      text: he('a (b C) d'),
      paragraph: 'auto',
      level: 0,
      levels: '000001000',
      order: [0, 1, 2, 3, 4, 5, 6, 7, 8],
      display: 'a (b C) d',
    },
    {
      name: '"A (B c) D": content matches the RTL embedding, brackets mirror (N0 b, L4)',
      text: he('A (B c) D'),
      paragraph: 'auto',
      level: 1,
      levels: '111112111',
      order: [8, 7, 6, 5, 4, 3, 2, 1, 0],
      display: 'D (c B) A',
    },
    {
      name: 'RTL content preceded by RTL context takes R (N0 c.1)',
      text: he('A (B) c'),
      paragraph: 0,
      level: 0,
      levels: '1111100',
      order: [4, 3, 2, 1, 0, 5, 6],
      display: '(B) A c',
    },
    {
      name: 'RTL content preceded by LTR context takes the embedding direction (N0 c.2)',
      text: he('a (B) c'),
      paragraph: 0,
      level: 0,
      levels: '0001000',
      order: [0, 1, 2, 3, 4, 5, 6],
      display: 'a (B) c',
    },
    {
      name: 'a European number inside brackets counts as R (N0)',
      text: he('A (1) b'),
      paragraph: 0,
      level: 0,
      levels: '1112100',
      order: [4, 3, 2, 1, 0, 5, 6],
      display: '(1) A b',
    },
    {
      name: 'brackets with only neutrals inside are left to N1/N2 (N0 d)',
      text: he('A (!) b'),
      paragraph: 0,
      level: 0,
      levels: '1000000',
      order: [0, 1, 2, 3, 4, 5, 6],
      display: 'A (!) b',
    },
    {
      name: 'UAX #9 N0 example "AB(CD[&ef]!)gh" in an RTL paragraph',
      text: he('AB(CD[&ef]!)gh'),
      paragraph: 1,
      level: 1,
      levels: '11111112211122',
      order: [12, 13, 11, 10, 9, 7, 8, 6, 5, 4, 3, 2, 1, 0],
      display: 'gh(![ef&]DC)BA',
    },
    {
      name: 'UAX #9 N0 example "smith (fabrikam ARABIC) HEBREW" in an RTL paragraph',
      text: he('smith (fabrikam ARABIC) HEBREW'),
      paragraph: 1,
      level: 1,
      levels: '222221122222222111111111111111',
      order: [29, 28, 27, 26, 25, 24, 23, 22, 21, 20, 19, 18, 17, 16, 15, 7, 8, 9, 10, 11, 12, 13, 14, 6, 5, 0, 1, 2, 3, 4],
      display: 'WERBEH (CIBARA fabrikam) smith',
    },
    {
      // The first pair resolves to L (its content has L); that L, not the C inside it, is the
      // context the second pair sees, so N0 c.2 makes the second pair L too.
      name: 'a pair resolved earlier is context for the next pair (N0)',
      text: he('A(bC)(D)'),
      paragraph: 0,
      level: 0,
      levels: '10010010',
      order: [0, 1, 2, 3, 4, 5, 6, 7],
      display: 'A(bC)(D)',
    },
    {
      name: 'U+2329 pairs with its canonical equivalent U+3009 (BD16)',
      text: he('A ') + '〈' + he('B') + '〉 c',
      paragraph: 0,
      level: 0,
      levels: '1111100',
      order: [4, 3, 2, 1, 0, 5, 6],
      display: '〈B〉 A c',
    },
    {
      name: 'a mark after a resolved bracket follows the bracket (N0)',
      text: he('A (B)') + '́ c',
      paragraph: 0,
      level: 0,
      levels: '11111100',
      order: [5, 4, 3, 2, 1, 0, 6, 7],
      display: '́(B) A c',
    },
  ]);

  it('gives up pairing after 63 open brackets (BD16)', () => {
    // A, n openers, B, ")", " c": the last opener pairs with ")" (N0 c.1 -> R) only while the
    // stack has room for it.
    const withOpeners = (n: number) => resolveLevels(he('A') + '('.repeat(n) + he('B') + ') c', 0).levels;
    expect(withOpeners(63)[65]).toBe(1);
    expect(withOpeners(64)[66]).toBe(0);
  });

  it('leaves unmatched brackets to the neutral rules', () => {
    // "A (b": the "(" sits between R and L, so N2 gives it the embedding direction.
    expect(bidi(he('A (b'), 0).levels).toEqual([1, 0, 0, 0]);
    expect(bidi(he('A (b'), 1).levels).toEqual([1, 1, 1, 2]);
  });
});

describe('bidi: nonspacing marks (W1)', () => {
  check([
    { name: 'a Hebrew point takes R from its letter', text: he('A') + 'ָ b', paragraph: 0, level: 0, levels: '1100', order: [1, 0, 2, 3], display: 'ָA b' },
    { name: 'a mark after a digit is EN, so a following % joins (W1, W5)', text: he('A 1') + '̀%', paragraph: 'auto', level: 1, levels: '11222', order: [2, 3, 4, 1, 0], display: '1̀% A' },
    { name: 'a mark after AL passes AL on to W2', text: ar('A') + 'َ1%', paragraph: 'auto', level: 1, levels: '1121', order: [3, 2, 1, 0], display: '%1َA' },
    { name: 'a mark at the paragraph start takes sos', text: '̀a', paragraph: 1, level: 1, levels: '12', order: [1, 0], display: 'à' },
    { name: 'a mark starting an isolate takes its sos, not the outer text', text: `a${RLI}̀${PDI}b`, paragraph: 'auto', level: 0, levels: '00100', order: [0, 1, 2, 3, 4], display: 'àb' },
  ]);
});

describe('bidi: explicit embeddings and overrides', () => {
  check([
    { name: 'RLE around LTR text changes nothing visible', text: `a${RLE}b c${PDF}d`, paragraph: 'auto', level: 0, levels: '0022220', order: [0, 1, 2, 3, 4, 5, 6], display: 'ab cd' },
    { name: 'RLE around Hebrew', text: `a${RLE}${he('B C')}${PDF}d`, paragraph: 'auto', level: 0, levels: '0011110', order: [0, 1, 5, 4, 3, 2, 6], display: 'aC Bd' },
    { name: 'LRE inside an RTL paragraph', text: `${he('A')}${LRE}b c${PDF}${he('D')}`, paragraph: 'auto', level: 1, levels: '1122221', order: [6, 2, 3, 4, 5, 1, 0], display: 'Db cA' },
    { name: 'RLO forces Latin letters right to left', text: `a${RLO}bc${PDF}d`, paragraph: 'auto', level: 0, levels: '001110', order: [0, 1, 4, 3, 2, 5], display: 'acbd' },
    { name: 'LRO forces Hebrew letters left to right', text: `${he('A')}${LRO}${he('BC')}${PDF}${he('D')}`, paragraph: 'auto', level: 1, levels: '112221', order: [5, 2, 3, 4, 1, 0], display: 'DBCA' },
    {
      name: 'nested RLE and LRE',
      text: `a${RLE}${he('B')}${LRE}c${PDF}${he('D')}${PDF}e`,
      paragraph: 'auto',
      level: 0,
      levels: '001122110',
      order: [0, 1, 7, 6, 4, 5, 3, 2, 8],
      display: 'aDcBe',
    },
    {
      // The run after the PDF starts where the embedding (level 1) ends, so its sos is R and W7
      // leaves the digit EN.
      name: 'a number after an RTL embedding sees sos R (X10)',
      text: `${RLE}${he('A')}${PDF}1`,
      paragraph: 0,
      level: 0,
      levels: '0112',
      order: [0, 3, 2, 1],
      display: '1A',
    },
    { name: 'P2 does not skip embeddings', text: `${RLE}a${PDF}${he('B')}`, paragraph: 'auto', level: 0, levels: '0221', order: [0, 3, 1, 2], display: 'Ba' },
  ]);

  it('stops at depth 125 and counts overflowed embeddings (X5, X7)', () => {
    const text = RLE.repeat(70) + 'b' + PDF.repeat(7) + 'c' + PDF + 'd';
    const { levels } = resolveLevels(text, 0);
    // 63 RLEs reach level 125; the other 7 overflow, and the first 7 PDFs only cancel those.
    expect(levels[70]).toBe(126); // b: L at 125
    expect(levels[78]).toBe(126); // c: still at 125
    expect(levels[80]).toBe(124); // d: back at 123
    expect(resolveLevels(LRE.repeat(70) + 'b', 0).levels[70]).toBe(124);
  });

  it('an overflowed isolate swallows PDFs until its PDI (X5a, X6a, X7)', () => {
    // 62 LREs reach 124; the LRI would need 126 and overflows, so the RLE inside it is ignored
    // without counting, the PDF inside does nothing, and the PDF after the PDI pops one LRE.
    const text = LRE.repeat(62) + LRI + RLE + 'x' + PDF + PDI + 'y' + PDF + 'z';
    const { levels } = resolveLevels(text, 0);
    expect(levels[64]).toBe(124); // x
    expect(levels[67]).toBe(124); // y
    expect(levels[69]).toBe(122); // z
  });
});

describe('bidi: isolates', () => {
  check([
    { name: 'LRI inside an RTL paragraph', text: he('A ') + `${LRI}b c${PDI}` + he(' D'), paragraph: 'auto', level: 1, levels: '111222111', order: [8, 7, 6, 3, 4, 5, 2, 1, 0], display: 'D b c A' },
    {
      // The isolated "b" is invisible to W2/W7 outside, so "1" still follows R and stays EN.
      name: 'isolated content does not affect the number after it',
      text: he('A') + `${LRI}b${PDI}1`,
      paragraph: 0,
      level: 0,
      levels: '11212',
      order: [4, 3, 2, 1, 0],
      display: '1bA',
    },
    {
      // With an embedding instead, "1" starts a new run whose sos is L, so W7 makes it L.
      name: 'the same with LRE/PDF lets the embedding decide the number',
      text: he('A') + `${LRE}b${PDF}1`,
      paragraph: 0,
      level: 0,
      levels: '11220',
      order: [2, 3, 1, 0, 4],
      display: 'bA1',
    },
    { name: 'FSI with RTL content acts as RLI (X5c)', text: `a ${FSI}${he('B c')}${PDI} d`, paragraph: 'auto', level: 0, levels: '000112000', order: [0, 1, 2, 5, 4, 3, 6, 7, 8], display: 'a c B d' },
    { name: 'FSI with LTR content acts as LRI (X5c)', text: he('A ') + `${FSI}b ${he('C')}${PDI}` + he(' D'), paragraph: 'auto', level: 1, levels: '111223111', order: [8, 7, 6, 3, 4, 5, 2, 1, 0], display: 'D b C A' },
    { name: 'auto direction skips a matched isolate (P2)', text: `${RLI}${he('A')}${PDI}b`, paragraph: 'auto', level: 0, levels: '0100', order: [0, 1, 2, 3], display: 'Ab' },
    { name: 'auto direction skips an LTR isolate to find Hebrew (P2)', text: `${LRI}a${PDI}${he('B')}`, paragraph: 'auto', level: 1, levels: '1211', order: [3, 2, 1, 0], display: 'Ba' },
    { name: 'an unmatched isolate runs to the paragraph end (BD9, P2)', text: `${RLI}${he('A')}b`, paragraph: 'auto', level: 0, levels: '012', order: [0, 2, 1], display: 'bA' },
  ]);
});

describe('bidi: L1 and paragraphs', () => {
  it('resets trailing whitespace left inside an embedding', () => {
    expect([...resolveLevels(`${RLE}${he('A ')}`, 0).levels]).toEqual([0, 1, 0]);
  });

  it('resets trailing whitespace and isolate controls, even inside the isolate', () => {
    expect([...resolveLevels(`${RLI}${he('A ')}${PDI} `, 0).levels]).toEqual([0, 1, 0, 0, 0]);
  });

  it('resets trailing whitespace inside an LTR embedding of an RTL paragraph', () => {
    expect([...resolveLevels(he('A ') + `${LRE}b `, 'auto').levels]).toEqual([1, 1, 1, 2, 1]);
  });

  it('resets a tab and the spaces before it (segment separator)', () => {
    expect(bidi('a\tb', 1)).toMatchObject({ levels: [2, 1, 2], order: [2, 1, 0] });
    expect(bidi('a \tb', 1)).toMatchObject({ levels: [2, 1, 1, 2], order: [3, 2, 1, 0] });
  });

  it('resets trailing format characters (BN)', () => {
    expect([...resolveLevels(he('A') + '‍', 0).levels]).toEqual([1, 0]);
    expect([...resolveLevels(he('A') + '‍' + he('B'), 'auto').levels]).toEqual([1, 1, 1]);
  });

  it('ends a paragraph at a paragraph separator', () => {
    expect([...resolveLevels('a b\nc', 1).levels]).toEqual([2, 2, 2, 1, 2]);
    // With auto each paragraph finds its own direction; the result reports the first one.
    expect(resolveLevels(he('A') + '\nb', 'auto')).toEqual({ levels: new Uint8Array([1, 1, 0]), paragraphLevel: 1 });
    // CR LF is one separator, so the LF stays with the RTL paragraph.
    expect([...resolveLevels(he('A') + '\r\nb', 'auto').levels]).toEqual([1, 1, 1, 0]);
  });

  it('applies L1 at a line break with lineLevels', () => {
    const text = he('A b c');
    const { levels } = resolveLevels(text, 'auto');
    expect([...levels]).toEqual([1, 1, 2, 2, 2]);
    const line = lineLevels(text, levels, 1, 0, 4);
    expect([...line]).toEqual([1, 1, 2, 1, 2]);
    expect([...levels]).toEqual([1, 1, 2, 2, 2]); // the input is not modified
    // The trailing space of the first line moves to its visual end (the left, in RTL).
    expect(visualOrder(line, 0, 4)).toEqual([3, 2, 1, 0]);
    expect(visualOrder(levels, 0, 4)).toEqual([2, 3, 1, 0]);
  });

  it('handles empty text', () => {
    expect(resolveLevels('', 'auto')).toEqual({ levels: new Uint8Array(0), paragraphLevel: 0 });
    expect(resolveLevels('', 1).paragraphLevel).toBe(1);
  });
});

describe('bidi: code units', () => {
  it('gives both halves of a surrogate pair the same level', () => {
    // U+10900 and U+10901 are Phoenician letters (R).
    const { levels } = resolveLevels('a \u{10900}\u{10901}', 'auto');
    expect([...levels]).toEqual([0, 0, 1, 1, 1, 1]);
  });

  it('tolerates lone surrogates', () => {
    expect([...resolveLevels('\ud800a', 'auto').levels]).toEqual([0, 0]);
  });
});

describe('visualOrder (L2)', () => {
  it('reverses runs from the highest level down to the lowest odd level', () => {
    expect(visualOrder([0, 0, 1, 1, 0], 0, 5)).toEqual([0, 1, 3, 2, 4]);
    expect(visualOrder([1, 1, 2, 2, 1], 0, 5)).toEqual([4, 2, 3, 1, 0]);
    expect(visualOrder([0, 1, 2, 3, 0], 0, 5)).toEqual([0, 2, 3, 1, 4]);
  });

  it('orders one line of a longer paragraph', () => {
    expect(visualOrder([0, 1, 1, 1, 0], 1, 4)).toEqual([3, 2, 1]);
    expect(visualOrder([], 0, 0)).toEqual([]);
  });

  it('leaves even-only levels in logical order', () => {
    expect(visualOrder([0, 2, 2, 0], 0, 4)).toEqual([0, 1, 2, 3]);
  });
});

describe('mirrorChar (L4)', () => {
  it('maps mirrored pairs both ways', () => {
    const pairs: [string, string][] = [['(', ')'], ['[', ']'], ['{', '}'], ['<', '>'], ['«', '»'], ['≤', '≥'], ['‹', '›'], ['〈', '〉'], ['∈', '∋']];
    for (const [a, b] of pairs) {
      expect(mirrorChar(a.codePointAt(0)!)).toBe(b.codePointAt(0));
      expect(mirrorChar(b.codePointAt(0)!)).toBe(a.codePointAt(0));
    }
  });

  it('has no mirror for unmirrored characters', () => {
    expect(mirrorChar(0x61)).toBeUndefined();
    expect(mirrorChar(0x2f)).toBeUndefined();
    expect(mirrorChar(0x2211)).toBeUndefined(); // ∑ is Bidi_Mirrored but has no mirror glyph
  });
});

describe('bidiClass', () => {
  it('classifies representative characters', () => {
    const expected: [number, string][] = [
      [0x41, 'L'], [0x3b1, 'L'], [0x436, 'L'], [0x4e00, 'L'], [0xac00, 'L'], [0x915, 'L'], [0x93f, 'L'],
      [0x200e, 'L'], [0x5d0, 'R'], [0x5be, 'R'], [0x200f, 'R'], [0x7c0, 'R'], [0x800, 'R'], [0x840, 'R'],
      [0xfb1d, 'R'], [0x5ff, 'R'], [0x10900, 'R'], [0x1e900, 'R'], [0x627, 'AL'], [0x61c, 'AL'], [0x710, 'AL'],
      [0x780, 'AL'], [0x750, 'AL'], [0x8a0, 'AL'], [0xfb50, 'AL'], [0xfe70, 'AL'], [0x1ee00, 'AL'],
      [0x30, 'EN'], [0x39, 'EN'], [0xb2, 'EN'], [0x6f0, 'EN'], [0xff10, 'EN'], [0x1d7ce, 'EN'], [0x2b, 'ES'],
      [0x2d, 'ES'], [0x2212, 'ES'], [0x23, 'ET'], [0x24, 'ET'], [0x25, 'ET'], [0xa3, 'ET'], [0x20ac, 'ET'],
      [0x66a, 'ET'], [0xe3f, 'ET'], [0x660, 'AN'], [0x669, 'AN'], [0x600, 'AN'], [0x66b, 'AN'], [0x10e60, 'AN'],
      [0x2c, 'CS'], [0x2e, 'CS'], [0x2f, 'CS'], [0x3a, 'CS'], [0xa0, 'CS'], [0x60c, 'CS'], [0x300, 'NSM'],
      [0x36f, 'NSM'], [0x5b0, 'NSM'], [0x64b, 'NSM'], [0x670, 'NSM'], [0x941, 'NSM'], [0xe31, 'NSM'],
      [0x20d0, 'NSM'], [0xfe00, 'NSM'], [0xfb1e, 'NSM'], [0xe0100, 'NSM'], [0x200b, 'BN'], [0x200c, 'BN'],
      [0x200d, 'BN'], [0xad, 'BN'], [0xfeff, 'BN'], [0x2060, 'BN'], [0x0, 'BN'], [0xe0001, 'BN'], [0x10ffff, 'BN'],
      [0xa, 'B'], [0xd, 'B'], [0x85, 'B'], [0x2029, 'B'], [0x9, 'S'], [0xb, 'S'], [0x1f, 'S'], [0x20, 'WS'],
      [0xc, 'WS'], [0x2028, 'WS'], [0x3000, 'WS'], [0x2003, 'WS'], [0x21, 'ON'], [0x28, 'ON'], [0x2018, 'ON'],
      [0x2190, 'ON'], [0x1f600, 'ON'], [0xfdfd, 'ON'], [0x202a, 'LRE'], [0x202b, 'RLE'], [0x202c, 'PDF'],
      [0x202d, 'LRO'], [0x202e, 'RLO'], [0x2066, 'LRI'], [0x2067, 'RLI'], [0x2068, 'FSI'], [0x2069, 'PDI'],
    ];
    for (const [cp, cls] of expected) expect([cp.toString(16), bidiClass(cp)]).toEqual([cp.toString(16), cls]);
  });

  it('returns a valid class for every code point', () => {
    const classes = new Set<string>();
    for (let cp = 0; cp <= 0x10ffff; cp++) classes.add(bidiClass(cp));
    expect(classes.size).toBe(23);
    expect(bidiClass(-1)).toBe('L');
    expect(bidiClass(0x110000)).toBe('L');
  });

  it('isRtlChar is true for R and AL only', () => {
    expect([0x5d0, 0x627, 0x10900, 0x1e900].map(isRtlChar)).toEqual([true, true, true, true]);
    expect([0x41, 0x661, 0x30, 0x20].map(isRtlChar)).toEqual([false, false, false, false]);
  });
});
