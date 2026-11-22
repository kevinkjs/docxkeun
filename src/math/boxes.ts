/**
 * Boxes, atoms and the style context of equation layout. A box is a group of display items
 * measured around its own baseline origin, so structures can stack and shift boxes without
 * knowing what is inside. Operator spacing follows the TeXbook's rules (appendix G), which
 * Word's math layout matches closely.
 */
import type { DisplayItem } from '../layout/types.js';
import type { MathFonts } from './text.js';

export type AtomClass = 'ord' | 'op' | 'bin' | 'rel' | 'open' | 'close' | 'punct' | 'inner';

export interface Box {
  width: number;
  /** Ink height above the baseline and depth below it (both ≥ 0). */
  ascent: number;
  descent: number;
  /** Items relative to the box origin: left end of the baseline, y down. */
  items: DisplayItem[];
  /** How far the last glyph's ink leans past its advance (italic letters); superscripts move right by it. */
  italic: number;
  /** A plain character or number, not a structure: scripts attach with TeX's character rules. */
  simple?: boolean;
}

export interface Atom {
  box: Box;
  /** Class seen by the atom to the left and to the right (structures such as m:d differ). */
  left: AtomClass;
  right: AtomClass;
  /** Typed spaces and m:eqArr alignment marks take no part in operator spacing. */
  kind?: 'glue' | 'mark';
}

/** Style of the current position: TeX's display / text / script / scriptscript, cramped or not. */
export interface Ctx {
  fonts: MathFonts;
  /** Font size of level-0 text (pt). */
  base: number;
  level: 0 | 1 | 2;
  /** Display style (level 0 of a display equation): big operators, full-size fractions. */
  display: boolean;
  /** Cramped (under bars, in denominators and subscripts): superscripts sit lower. */
  cramped: boolean;
  color: string;
  /** Height of the math axis in em (the centre of the minus sign). */
  axis: number;
  /** Inside m:fName: letters default to upright. */
  fn: boolean;
  /** Inside m:eqArr: `&` marks an alignment point. */
  align: boolean;
}

/** Script sizes relative to the base: Cambria Math's ScriptPercentScaleDown (73) and ScriptScript (60). */
export const SCALE = [1, 0.73, 0.6] as const;
/** Default rule thickness (fraction bars, overbars, radicals), in em. */
export const RULE = 0.06;
export const THIN = 3 / 18;
export const MEDIUM = 4 / 18;
export const THICK = 5 / 18;
export const X_HEIGHT = 0.46;

export function em(ctx: Ctx): number {
  return ctx.base * SCALE[ctx.level];
}

export function scriptStyle(ctx: Ctx, cramped = ctx.cramped): Ctx {
  return { ...ctx, level: Math.min(2, ctx.level + 1) as 0 | 1 | 2, display: false, cramped, align: false, fn: false };
}

/** Numerators and denominators: display → text size, text → script, and so on. */
export function fractionStyle(ctx: Ctx, cramped: boolean): Ctx {
  if (ctx.display && ctx.level === 0) return { ...ctx, display: false, cramped, align: false, fn: false };
  return scriptStyle(ctx, cramped);
}

export function emptyBox(width = 0): Box {
  return { width, ascent: 0, descent: 0, items: [], italic: 0 };
}

/** Move items in place (boxes are owned by exactly one parent, so nothing is copied). */
export function shiftItems(items: DisplayItem[], dx: number, dy: number): void {
  if (dx === 0 && dy === 0) return;
  for (const it of items) {
    switch (it.type) {
      case 'text':
        it.x += dx;
        it.y += dy;
        for (let i = 0; i < it.xs.length; i++) it.xs[i]! += dx;
        break;
      case 'rect':
      case 'image':
        it.x += dx;
        it.y += dy;
        break;
      case 'line':
        it.x1 += dx;
        it.x2 += dx;
        it.y1 += dy;
        it.y2 += dy;
        break;
      case 'path':
        for (const c of it.commands) {
          if (c.op === 'Z') continue;
          c.x += dx;
          c.y += dy;
          if (c.op === 'C') {
            c.x1 += dx;
            c.y1 += dy;
            c.x2 += dx;
            c.y2 += dy;
          }
        }
        break;
    }
  }
}

/** Collects boxes and drawn items into a new box, tracking its extent. */
export class Builder {
  readonly items: DisplayItem[] = [];
  ascent = 0;
  descent = 0;
  right = 0;

  /** Place a box with its origin at (x, y); y is down, so a raised box has negative y. */
  put(box: Box, x: number, y: number): void {
    shiftItems(box.items, x, y);
    for (const it of box.items) this.items.push(it);
    this.ascent = Math.max(this.ascent, box.ascent - y);
    this.descent = Math.max(this.descent, box.descent + y);
    this.right = Math.max(this.right, x + box.width);
  }

  /** Add an item the caller drew, covering `top`..`bottom` (y down). */
  draw(item: DisplayItem, top: number, bottom: number): void {
    this.items.push(item);
    this.ascent = Math.max(this.ascent, -top);
    this.descent = Math.max(this.descent, bottom);
  }

  done(width = this.right, italic = 0): Box {
    return { width, ascent: this.ascent, descent: this.descent, items: this.items, italic };
  }
}

// ---------------------------------------------------------------------------------------------
// Operator spacing.

const CLASSES: AtomClass[] = ['ord', 'op', 'bin', 'rel', 'open', 'close', 'punct', 'inner'];
/**
 * TeX's inter-atom spacing table (rows: left atom, columns: right atom). 1 thin, 2 medium,
 * 3 thick; negative entries apply only in display and text style, not in scripts.
 */
const TABLE: number[][] = [
  /* ord   */ [0, 1, -2, -3, 0, 0, 0, -1],
  /* op    */ [1, 1, 0, -3, 0, 0, 0, -1],
  /* bin   */ [-2, -2, 0, 0, -2, 0, 0, -2],
  /* rel   */ [-3, -3, 0, 0, -3, 0, 0, -3],
  /* open  */ [0, 0, 0, 0, 0, 0, 0, 0],
  /* close */ [0, 1, -2, -3, 0, 0, 0, -1],
  /* punct */ [-1, -1, 0, -1, -1, -1, -1, -1],
  /* inner */ [-1, 1, -2, -3, -1, 0, -1, -1],
];
const AMOUNT = [0, THIN, MEDIUM, THICK];

export function spaceBetween(left: AtomClass, right: AtomClass, ctx: Ctx): number {
  const code = TABLE[CLASSES.indexOf(left)]![CLASSES.indexOf(right)]!;
  if (code < 0 && ctx.level > 0) return 0;
  return AMOUNT[Math.abs(code)]! * em(ctx);
}

const UNARY_AFTER = new Set<AtomClass>(['bin', 'op', 'rel', 'open', 'punct']);
const UNARY_BEFORE = new Set<AtomClass>(['rel', 'close', 'punct']);

/**
 * A binary operator with no left operand (start of a row, after an opening delimiter or another
 * operator) or no right operand is unary: it becomes ordinary and gets no space ("−b", "(−1)").
 */
export function resolveBinaries(atoms: Atom[]): void {
  const real = atoms.filter((a) => !a.kind);
  for (let i = 0; i < real.length; i++) {
    const a = real[i]!;
    if (a.left !== 'bin' || a.right !== 'bin') continue;
    const prev = real[i - 1];
    if (!prev || UNARY_AFTER.has(prev.right)) a.left = a.right = 'ord';
  }
  for (let i = 0; i < real.length; i++) {
    const a = real[i]!;
    if (a.left !== 'bin' || a.right !== 'bin') continue;
    const next = real[i + 1];
    if (!next || UNARY_BEFORE.has(next.left)) a.left = a.right = 'ord';
  }
}

export interface Row {
  box: Box;
  /** Classes of the first and last atom (undefined for an empty row). */
  first: AtomClass | undefined;
  last: AtomClass | undefined;
}

/** Lay atoms side by side with operator spacing. `before` is the class of the atom to the left. */
export function hpack(atoms: Atom[], ctx: Ctx, before?: AtomClass): Row {
  const b = new Builder();
  let x = 0;
  let prev = before;
  let first: AtomClass | undefined;
  let last: AtomClass | undefined;
  let real = 0;
  let lastBox: Box | undefined;
  for (const a of atoms) {
    if (a.kind === 'mark') continue;
    if (!a.kind) {
      if (prev) x += spaceBetween(prev, a.left, ctx);
      prev = a.right;
      first ??= a.left;
      last = a.right;
      real++;
    }
    b.put(a.box, x, 0);
    x += a.box.width;
    lastBox = a.box;
  }
  const box = b.done(x, lastBox?.italic ?? 0);
  if (real === 1 && atoms.length === 1 && lastBox?.simple) box.simple = true;
  return { box, first, last };
}
