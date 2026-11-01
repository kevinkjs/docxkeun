/** numbering.xml: list definitions and the counters that produce list labels during layout. */
import { attr, elements, firstChild, type XmlDocument, type XmlElement } from '../xml/index.js';
import { parseParagraphProps, parseRunProps, type ParagraphProps, type RunProps } from './props.js';
import { parseIntSafe, parseOnOff } from './units.js';

export interface LevelDef {
  ilvl: number;
  start: number;
  format: string;
  text: string;
  jc: 'left' | 'center' | 'right';
  suffix: 'tab' | 'space' | 'nothing';
  isLegal: boolean;
  restart?: number;
  pStyle?: string;
  pPr: ParagraphProps;
  rPr: RunProps;
}

interface AbstractNum {
  id: string;
  levels: Map<number, LevelDef>;
  numStyleLink?: string;
  styleLink?: string;
}

interface NumInstance {
  id: string;
  abstractId: string;
  overrides: Map<number, { start?: number; level?: LevelDef }>;
}

function parseLevel(node: XmlElement): LevelDef {
  const v = (name: string) => attr(firstChild(node, name), 'w:val');
  const jc = v('w:lvlJc');
  const suffix = v('w:suff');
  const restart = parseIntSafe(v('w:lvlRestart'));
  return {
    ilvl: parseIntSafe(attr(node, 'w:ilvl')) ?? 0,
    start: parseIntSafe(v('w:start')) ?? 1,
    format: v('w:numFmt') ?? 'decimal',
    text: v('w:lvlText') ?? '',
    jc: jc === 'center' ? 'center' : jc === 'right' || jc === 'end' ? 'right' : 'left',
    suffix: suffix === 'space' ? 'space' : suffix === 'nothing' ? 'nothing' : 'tab',
    isLegal: firstChild(node, 'w:isLgl') !== undefined && parseOnOff(v('w:isLgl')),
    restart,
    pStyle: v('w:pStyle'),
    pPr: parseParagraphProps(firstChild(node, 'w:pPr')),
    rPr: parseRunProps(firstChild(node, 'w:rPr')),
  };
}

export class NumberingDefs {
  private readonly abstracts = new Map<string, AbstractNum>();
  private readonly nums = new Map<string, NumInstance>();

  constructor(readonly doc?: XmlDocument) {
    if (!doc) return;
    for (const a of elements(doc.root, 'w:abstractNum')) {
      const id = attr(a, 'w:abstractNumId');
      if (id === undefined) continue;
      const levels = new Map<number, LevelDef>();
      for (const l of elements(a, 'w:lvl')) {
        const level = parseLevel(l);
        levels.set(level.ilvl, level);
      }
      this.abstracts.set(id, {
        id,
        levels,
        numStyleLink: attr(firstChild(a, 'w:numStyleLink'), 'w:val'),
        styleLink: attr(firstChild(a, 'w:styleLink'), 'w:val'),
      });
    }
    for (const n of elements(doc.root, 'w:num')) {
      const id = attr(n, 'w:numId');
      const abstractId = attr(firstChild(n, 'w:abstractNumId'), 'w:val');
      if (id === undefined || abstractId === undefined) continue;
      const overrides = new Map<number, { start?: number; level?: LevelDef }>();
      for (const o of elements(n, 'w:lvlOverride')) {
        const ilvl = parseIntSafe(attr(o, 'w:ilvl')) ?? 0;
        const lvl = firstChild(o, 'w:lvl');
        overrides.set(ilvl, {
          start: parseIntSafe(attr(firstChild(o, 'w:startOverride'), 'w:val')),
          level: lvl ? { ...parseLevel(lvl), ilvl } : undefined,
        });
      }
      this.nums.set(id, { id, abstractId, overrides });
    }
  }

  /** Abstract definition of a num, following numStyleLink → styleLink indirection. */
  abstractFor(numId: string, styleNumId?: (styleId: string) => string | undefined): AbstractNum | undefined {
    const num = this.nums.get(numId);
    let abs = num ? this.abstracts.get(num.abstractId) : undefined;
    if (abs?.numStyleLink && styleNumId) {
      const linkedNum = styleNumId(abs.numStyleLink);
      const linked = linkedNum ? this.nums.get(linkedNum) : undefined;
      const target = linked ? this.abstracts.get(linked.abstractId) : undefined;
      if (target) abs = target;
    }
    return abs;
  }

  level(numId: string | undefined, ilvl: number, styleNumId?: (styleId: string) => string | undefined): LevelDef | undefined {
    if (!numId || numId === '0') return undefined; // numId 0 removes numbering
    const num = this.nums.get(numId);
    if (!num) return undefined;
    const override = num.overrides.get(ilvl)?.level;
    if (override) return override;
    return this.abstractFor(numId, styleNumId)?.levels.get(ilvl);
  }

  levels(numId: string, styleNumId?: (styleId: string) => string | undefined): LevelDef[] {
    const out: LevelDef[] = [];
    for (let i = 0; i < 9; i++) {
      const l = this.level(numId, i, styleNumId);
      if (l) out[i] = l;
    }
    return out;
  }

  startOverride(numId: string, ilvl: number): number | undefined {
    return this.nums.get(numId)?.overrides.get(ilvl)?.start;
  }

  abstractIdOf(numId: string): string | undefined {
    return this.nums.get(numId)?.abstractId;
  }

  has(numId: string | undefined): boolean {
    return !!numId && this.nums.has(numId);
  }
}

/**
 * Stateful counters, advanced in document order. Lists that share an abstract definition share
 * counters (Word behaviour); a num with a startOverride restarts them the first time it is used.
 */
export class NumberingCounter {
  private readonly counters = new Map<string, Array<number | undefined>>();
  private readonly startedNums = new Set<string>();

  constructor(private readonly defs: NumberingDefs, private readonly styleNumId?: (styleId: string) => string | undefined) {}

  /** Advance the counter for a paragraph and return its label, or undefined if not numbered. */
  next(numId: string | undefined, ilvl: number): { label: string; level: LevelDef } | undefined {
    if (!numId) return undefined;
    const level = this.defs.level(numId, ilvl, this.styleNumId);
    if (!level) return undefined;
    const key = this.defs.abstractIdOf(numId) ?? numId;
    let counts = this.counters.get(key);
    if (!counts) {
      counts = [];
      this.counters.set(key, counts);
    }
    if (!this.startedNums.has(numId)) {
      this.startedNums.add(numId);
      for (let i = 0; i < 9; i++) {
        const start = this.defs.startOverride(numId, i);
        if (start !== undefined) counts[i] = start - 1;
      }
    }
    const levels = this.defs.levels(numId, this.styleNumId);
    counts[ilvl] = (counts[ilvl] ?? (level.start - 1)) + 1;
    // Reset deeper levels (respecting w:lvlRestart, where 0 means never restart).
    for (let i = ilvl + 1; i < 9; i++) {
      const deeper = levels[i];
      const restartAt = deeper?.restart;
      if (restartAt === 0) continue;
      // w:lvlRestart is 1-based: restart when a level at or above it advances.
      if (restartAt === undefined || ilvl + 1 <= restartAt) counts[i] = undefined;
    }
    const label = level.text.replace(/%([1-9])/g, (_, d: string) => {
      const i = Number(d) - 1;
      const l = levels[i];
      const value = counts[i] ?? l?.start ?? 1;
      const format = level.isLegal && i < ilvl ? 'decimal' : (l?.format ?? 'decimal');
      return formatNumber(value, format);
    });
    return { label: level.format === 'none' ? '' : label, level };
  }
}

const ROMAN: Array<[number, string]> = [
  [1000, 'm'],
  [900, 'cm'],
  [500, 'd'],
  [400, 'cd'],
  [100, 'c'],
  [90, 'xc'],
  [50, 'l'],
  [40, 'xl'],
  [10, 'x'],
  [9, 'ix'],
  [5, 'v'],
  [4, 'iv'],
  [1, 'i'],
];

function roman(n: number): string {
  let s = '';
  for (const [v, r] of ROMAN) {
    while (n >= v) {
      s += r;
      n -= v;
    }
  }
  return s;
}

/** Word's letter numbering repeats the letter: a..z, aa..zz, aaa.. */
function letters(n: number): string {
  if (n <= 0) return '';
  const letter = String.fromCharCode(97 + ((n - 1) % 26));
  return letter.repeat(Math.floor((n - 1) / 26) + 1);
}

const ONES = ['zero', 'one', 'two', 'three', 'four', 'five', 'six', 'seven', 'eight', 'nine', 'ten', 'eleven', 'twelve', 'thirteen', 'fourteen', 'fifteen', 'sixteen', 'seventeen', 'eighteen', 'nineteen'];
const TENS = ['', '', 'twenty', 'thirty', 'forty', 'fifty', 'sixty', 'seventy', 'eighty', 'ninety'];

function cardinal(n: number): string {
  if (n < 20) return ONES[n]!;
  if (n < 100) return TENS[Math.floor(n / 10)]! + (n % 10 ? '-' + ONES[n % 10] : '');
  if (n < 1000) return ONES[Math.floor(n / 100)] + ' hundred' + (n % 100 ? ' ' + cardinal(n % 100) : '');
  return String(n);
}

function ordinalSuffix(n: number): string {
  const mod100 = n % 100;
  if (mod100 >= 11 && mod100 <= 13) return 'th';
  return ['th', 'st', 'nd', 'rd'][n % 10] ?? 'th';
}

export function formatNumber(n: number, format: string): string {
  switch (format) {
    case 'decimal':
      return String(n);
    case 'decimalZero':
      return n < 10 ? `0${n}` : String(n);
    case 'upperRoman':
      return roman(n).toUpperCase();
    case 'lowerRoman':
      return roman(n);
    case 'upperLetter':
      return letters(n).toUpperCase();
    case 'lowerLetter':
      return letters(n);
    case 'ordinal':
      return `${n}${ordinalSuffix(n)}`;
    case 'cardinalText': {
      const s = cardinal(n);
      return s.charAt(0).toUpperCase() + s.slice(1);
    }
    case 'ordinalText':
      return cardinal(n) + ordinalSuffix(n);
    case 'decimalEnclosedParen':
      return `(${n})`;
    case 'decimalEnclosedFullstop':
      return `${n}.`;
    case 'bullet':
    case 'none':
      return '';
    default:
      return String(n);
  }
}
