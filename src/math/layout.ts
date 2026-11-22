/**
 * Equation layout: each OMML structure (fraction, scripts, radical, n-ary operator, delimiters,
 * matrix, accent, …) is built into a box from the boxes of its arguments. Sizes and shifts follow
 * TeX's rules and the Office math constants, scaled by the current style's em.
 */
import type { DisplayItem, LineItem, RectItem } from '../layout/types.js';
import { measureChars, textItem } from '../layout/paint.js';
import { elements, firstChild } from '../xml/dom.js';
import type { XmlElement } from '../xml/dom.js';
import { Builder, RULE, X_HEIGHT, em, emptyBox, fractionStyle, hpack, resolveBinaries, scriptStyle, shiftItems, spaceBetween } from './boxes.js';
import type { Atom, AtomClass, Box, Ctx, Row } from './boxes.js';
import { charProp, ctrlColor, isPropertyElement, numProp, onOff, val } from './omml.js';
import {
  accentMetrics,
  accentPaths,
  arrowPaths,
  bigOperatorPaths,
  delimiterShape,
  delimiterWidth,
  ellipse,
  filled,
  horizontalFrame,
  integralShape,
  mapShape,
  polygons,
  radicalShape,
  strokeOutline,
  stroked,
  verticalFrame,
} from './paths.js';
import { DRAWN_OPERATORS, accentOf, arrowOf, delimiterOf, groupShapeOf, integralOf } from './symbols.js';
import { inkOf, runAtoms, symbolBox, textRunAtoms } from './text.js';

/** Wrappers whose content belongs to the equation (tracked insertions, smart tags…). */
const WRAPPERS = new Set(['w:ins', 'w:moveTo', 'w:smartTag', 'w:customXml', 'w:hyperlink', 'w:fldSimple']);

function ord(box: Box): Atom {
  return { box, left: 'ord', right: 'ord' };
}

function rect(x: number, y: number, w: number, h: number, color: string): RectItem {
  return { type: 'rect', x, y, w, h, fill: color };
}

/** Whether an argument element holds anything besides properties. */
function hasContent(el: XmlElement | undefined): el is XmlElement {
  return !!el && elements(el).some((c) => !isPropertyElement(c.name));
}

export function rowAtoms(container: XmlElement | undefined, ctx: Ctx): Atom[] {
  if (!container) return [];
  const out: Atom[] = [];
  for (const child of elements(container)) for (const a of elementAtoms(child, ctx)) out.push(a);
  return out;
}

/** Lay out an argument (m:e, m:num, m:sub, m:oMath…) as a spaced row. */
export function row(container: XmlElement | undefined, ctx: Ctx, before?: AtomClass): Row {
  const atoms = rowAtoms(container, ctx);
  resolveBinaries(atoms);
  return hpack(atoms, ctx, before);
}

function elementAtoms(el: XmlElement, ctx: Ctx): Atom[] {
  switch (el.name) {
    case 'm:r':
      return runAtoms(el, ctx);
    case 'w:r':
      return textRunAtoms(el, ctx);
    case 'm:f':
      return [fraction(el, ctx)];
    case 'm:sSup':
    case 'm:sSub':
    case 'm:sSubSup':
      return [scripted(el, ctx)];
    case 'm:sPre':
      return [prescripted(el, ctx)];
    case 'm:rad':
      return [radical(el, ctx)];
    case 'm:nary':
      return [nary(el, ctx)];
    case 'm:d':
      return [delimited(el, ctx)];
    case 'm:m':
      return [matrix(el, ctx)];
    case 'm:acc':
      return [accent(el, ctx)];
    case 'm:bar':
      return [overUnderBar(el, ctx)];
    case 'm:func':
      return [func(el, ctx)];
    case 'm:limLow':
    case 'm:limUpp':
      return [limit(el, ctx)];
    case 'm:groupChr':
      return [groupChr(el, ctx)];
    case 'm:box':
      return [boxed(el, ctx)];
    case 'm:borderBox':
      return [borderBox(el, ctx)];
    case 'm:eqArr':
      return [eqArray(el, ctx)];
    case 'm:phant':
      return [phantom(el, ctx)];
    case 'w:sdt':
      return rowAtoms(firstChild(el, 'w:sdtContent'), ctx);
    default:
      if (isPropertyElement(el.name)) return [];
      // Unknown math elements still show their content; other markup (bookmarks, deletions) does not.
      if (el.name.startsWith('m:') || WRAPPERS.has(el.name)) return rowAtoms(el, ctx);
      return [];
  }
}

// ---------------------------------------------------------------------------------------------
// Fractions.

function fraction(el: XmlElement, ctx: Ctx): Atom {
  const pr = firstChild(el, 'm:fPr');
  const type = val(pr, 'm:type') ?? 'bar';
  const numEl = firstChild(el, 'm:num');
  const denEl = firstChild(el, 'm:den');
  const s = em(ctx);
  const color = ctrlColor(pr) ?? ctx.color;
  const axis = ctx.axis * s;
  if (type === 'lin') {
    // Linear: num/den on the baseline at full size.
    const num = row(numEl, ctx).box;
    const den = row(denEl, ctx).box;
    const slash = symbolBox('/', s, color, ctx.fonts);
    const b = new Builder();
    b.put(num, 0, 0);
    b.put(slash, num.width, 0);
    b.put(den, num.width + slash.width, 0);
    return ord(b.done());
  }
  if (type === 'skw') {
    // Skewed: small numerator up left, denominator down right, a slash between them.
    const num = row(numEl, scriptStyle(ctx)).box;
    const den = row(denEl, scriptStyle(ctx, true)).box;
    const up = axis + num.descent + 0.04 * s;
    const down = den.ascent - axis + 0.04 * s;
    const top = up + num.ascent;
    const bottom = -(down + den.descent);
    const slant = 0.32 * (top - bottom);
    const gap = 0.04 * s;
    const mid = num.width + gap + slant / 2;
    const b = new Builder();
    b.put(num, 0, -up);
    const t = RULE * s * 0.9;
    const outline = strokeOutline([[mid - slant / 2, bottom], [mid + slant / 2, top]], [t]);
    b.draw(filled(polygons([outline]), color), -top, -bottom);
    b.put(den, mid + slant / 2 + gap, down);
    return ord(b.done());
  }
  const display = ctx.display && ctx.level === 0;
  const num = row(numEl, fractionStyle(ctx, ctx.cramped)).box;
  const den = row(denEl, fractionStyle(ctx, true)).box;
  const th = RULE * s;
  const gap = (display ? 0.12 : 0.07) * s;
  // Numerator and denominator clear the bar by `gap`, with minimum shifts so that fractions
  // side by side line up (TeX's num/denom parameters, a little tighter, as in Word).
  const u = Math.max(axis + th / 2 + gap + num.descent, (display ? 0.5 : 0.36) * s);
  const v = Math.max(den.ascent + gap + th / 2 - axis, (display ? 0.45 : 0.3) * s);
  const over = 0.07 * s;
  const margin = 0.03 * s;
  const inner = Math.max(num.width, den.width) + 2 * over;
  const b = new Builder();
  b.put(num, margin + (inner - num.width) / 2, -u);
  b.put(den, margin + (inner - den.width) / 2, v);
  if (type !== 'noBar') b.draw(rect(margin, -(axis + th / 2), inner, th, color), -(axis + th / 2), -(axis - th / 2));
  return ord(b.done(inner + 2 * margin));
}

// ---------------------------------------------------------------------------------------------
// Scripts.

interface ScriptOptions {
  pre?: boolean;
  /** Superscript and subscript start at the same x (m:alnScr), ignoring the italic correction. */
  aligned?: boolean;
  /** Move the subscript left under a slanted operator (integrals). */
  subKern?: number;
}

/**
 * Attach scripts to a base (TeX rules 18a–f): a character base uses fixed shifts, a structure
 * shifts scripts from its own top and bottom; superscript and subscript keep a minimum gap.
 */
function attachScripts(base: Box, sub: Box | undefined, sup: Box | undefined, ctx: Ctx, o: ScriptOptions = {}): Box {
  const s = em(ctx);
  const simple = base.simple === true;
  const u0 = simple ? 0 : base.ascent - 0.28 * s;
  const v0 = simple ? 0 : base.descent + 0.05 * s;
  let u = 0;
  let v = 0;
  if (sup) u = Math.max(u0, (ctx.cramped ? 0.29 : 0.35) * s, sup.descent + 0.12 * s);
  if (sub && !sup) v = Math.max(v0, 0.2 * s, sub.ascent - 0.8 * X_HEIGHT * s);
  if (sub && sup) {
    v = Math.max(v0, 0.25 * s);
    const minGap = 0.16 * s;
    const gap = u - sup.descent - (sub.ascent - v);
    if (gap < minGap) {
      v += minGap - gap;
      const psi = 0.8 * X_HEIGHT * s - (u - sup.descent);
      if (psi > 0) {
        u += psi;
        v -= psi;
      }
    }
  }
  const space = 0.04 * s;
  const b = new Builder();
  if (o.pre) {
    const w = Math.max(sup?.width ?? 0, sub?.width ?? 0);
    if (sup) b.put(sup, w - sup.width, -u);
    if (sub) b.put(sub, w - sub.width, v);
    b.put(base, w + space, 0);
    return b.done(w + space + base.width, base.italic);
  }
  const baseWidth = base.width;
  const italic = base.italic;
  b.put(base, 0, 0);
  const supX = baseWidth + (o.aligned ? 0 : italic);
  const subX = baseWidth - (o.subKern ?? 0);
  if (sup) b.put(sup, supX, -u);
  if (sub) b.put(sub, subX, v);
  const right = Math.max(baseWidth, sup ? supX + sup.width : 0, sub ? subX + sub.width : 0);
  return b.done(right + (sub || sup ? space : 0));
}

function scripted(el: XmlElement, ctx: Ctx): Atom {
  const base = row(firstChild(el, 'm:e'), ctx);
  const supEl = el.name === 'm:sSub' ? undefined : firstChild(el, 'm:sup');
  const subEl = el.name === 'm:sSup' ? undefined : firstChild(el, 'm:sub');
  const sup = supEl ? row(supEl, scriptStyle(ctx)).box : undefined;
  const sub = subEl ? row(subEl, scriptStyle(ctx, true)).box : undefined;
  const aligned = onOff(firstChild(el, 'm:sSubSupPr'), 'm:alnScr', false);
  const box = attachScripts(base.box, sub, sup, ctx, { aligned });
  // A script keeps its base's class: sin² is still an operator, (x)² still closes.
  return { box, left: base.first ?? 'ord', right: base.last ?? 'ord' };
}

function prescripted(el: XmlElement, ctx: Ctx): Atom {
  const base = row(firstChild(el, 'm:e'), ctx);
  const subEl = firstChild(el, 'm:sub');
  const supEl = firstChild(el, 'm:sup');
  const sub = hasContent(subEl) ? row(subEl, scriptStyle(ctx, true)).box : undefined;
  const sup = hasContent(supEl) ? row(supEl, scriptStyle(ctx)).box : undefined;
  return { box: attachScripts(base.box, sub, sup, ctx, { pre: true }), left: 'ord', right: base.last ?? 'ord' };
}

/** Limits centred over and under a base (n-ary limits in display, m:limLow, m:limUpp). */
function stackLimits(base: Box, under: Box | undefined, over: Box | undefined, overGap: number, underGap: number): Box {
  const w = Math.max(base.width, under?.width ?? 0, over?.width ?? 0);
  const b = new Builder();
  const baseAscent = base.ascent;
  const baseDescent = base.descent;
  b.put(base, (w - base.width) / 2, 0);
  if (over) b.put(over, (w - over.width) / 2, -(baseAscent + overGap + over.descent));
  if (under) b.put(under, (w - under.width) / 2, baseDescent + underGap + under.ascent);
  return b.done(w);
}

// ---------------------------------------------------------------------------------------------
// Radicals.

function radical(el: XmlElement, ctx: Ctx): Atom {
  const pr = firstChild(el, 'm:radPr');
  const s = em(ctx);
  const color = ctrlColor(pr) ?? ctx.color;
  const body = row(firstChild(el, 'm:e'), { ...ctx, cramped: true }).box;
  const degEl = firstChild(el, 'm:deg');
  const deg = !onOff(pr, 'm:degHide', false) && hasContent(degEl) ? row(degEl, { ...ctx, level: 2, display: false, cramped: true, align: false }).box : undefined;
  const display = ctx.display && ctx.level === 0;
  const rule = RULE * s;
  const gap = (display ? 0.12 : 0.08) * s;
  const top = Math.max(body.ascent, X_HEIGHT * s) + gap + rule;
  // The sign reaches a little below the content and is never shorter than a text radical.
  const depth = Math.max(body.descent + 0.06 * s, 0.9 * s - top);
  const height = top + depth;
  const signWidth = Math.min(0.56 + 0.1 * Math.max(0, height / s - 1.2), 0.95) * s;
  const b = new Builder();
  let x0 = 0;
  if (deg) {
    // The degree sits over the hook, its right edge above the downstroke, raised to 60% of the
    // sign; a wide degree pushes the sign right.
    const anchor = 0.3 * s;
    x0 = Math.max(0, deg.width - anchor + 0.03 * s);
    const hook = Math.min(0.48 * s, height * 0.5);
    const baseline = Math.max(-depth + 0.6 * height, -depth + hook + 0.08 * s) + deg.descent;
    b.put(deg, x0 + anchor - deg.width, -baseline);
  }
  const bodyX = x0 + signWidth + 0.04 * s;
  const barEnd = bodyX + body.width + 0.06 * s;
  b.put(body, bodyX, 0);
  b.draw(filled(polygons([radicalShape(x0, -depth, signWidth, top, barEnd, rule, s)]), color), -top, depth + 0.05 * s);
  return ord(b.done(barEnd + 0.03 * s));
}

// ---------------------------------------------------------------------------------------------
// N-ary operators.

interface BigOp {
  box: Box;
  /** How far the lower limit tucks under a slanted integral. */
  subKern: number;
}

function bigOperator(chr: string, ctx: Ctx, display: boolean, growTo: Box | undefined, color: string): BigOp {
  const s = em(ctx);
  const axis = ctx.axis * s;
  // m:grow: the operator grows to cover its argument, centred on the axis.
  const grown = (h: number) => (growTo ? Math.max(h, 2 * Math.max(growTo.ascent - axis, growTo.descent + axis) + 0.1 * s) : h);
  const box = (width: number, height: number, items: DisplayItem[]): Box => ({ width, ascent: axis + height / 2, descent: Math.max(0, height / 2 - axis), items, italic: 0 });
  const integral = integralOf(chr);
  if (integral) {
    const natural = (display ? 2.2 : 1.25) * s;
    const H = grown(natural);
    const W = (display ? 0.8 : 0.52) * s * Math.pow(H / natural, 0.3);
    const step = 0.62 * W;
    const bearing = 0.04 * s;
    const bottom = axis - H / 2;
    const commands = [];
    for (let k = 0; k < integral.count; k++) {
      const shape = integralShape(bearing + k * step, bottom, W, H, s);
      commands.push(...polygons([shape.outline]));
      for (const [cx, cy, r] of shape.balls) commands.push(...ellipse(cx, cy, r, r));
    }
    const items: DisplayItem[] = [filled(commands, color)];
    if (integral.contour) {
      const span = (integral.count - 1) * step;
      const cx = bearing + (span + W) / 2;
      items.push(stroked(ellipse(cx, axis, span / 2 + (display ? 0.2 : 0.14) * s, (display ? 0.2 : 0.14) * s), color, 0.05 * s));
    }
    return { box: box(2 * bearing + W + (integral.count - 1) * step, H, items), subKern: W * (display ? 0.35 : 0.3) };
  }
  const target = grown((display ? 1.4 : 1.0) * s);
  const cp = chr.codePointAt(0)!;
  const bearing = 0.06 * s;
  if (DRAWN_OPERATORS.has(chr) || !ctx.fonts.canShow(cp)) {
    const ratio = '⋃⋂'.includes(chr) ? 0.72 : '⋁⋀'.includes(chr) ? 0.8 : '⨀⨁⨂'.includes(chr) ? 1 : 0.85;
    const W = target * ratio;
    const drawn = bigOperatorPaths(chr, bearing, axis - target / 2, W, target, s, color);
    if (drawn) return { box: box(W + 2 * bearing, target, drawn), subKern: 0 };
  }
  // A glyph from the math fonts, scaled so its ink has the target height, centred on the axis.
  const font = ctx.fonts.pick(cp, false, false);
  const ink = inkOf(font, cp) ?? { xMin: 0, yMin: 0, xMax: font.advance(cp), yMax: 0.7 };
  const inkHeight = Math.max(ink.yMax - ink.yMin, 0.1);
  const size = target / inkHeight;
  const x = bearing - ink.xMin * size;
  const baseline = axis - ((ink.yMin + ink.yMax) / 2) * size;
  const m = measureChars(chr, font, size, x);
  const item = textItem(x, -baseline, m.chars, m.xs, m.widths, font, size, color);
  return { box: box((ink.xMax - ink.xMin) * size + 2 * bearing, target, [item]), subKern: 0 };
}

function nary(el: XmlElement, ctx: Ctx): Atom {
  const pr = firstChild(el, 'm:naryPr');
  const chr = charProp(pr, 'm:chr', '∫') || '∫';
  const integral = integralOf(chr);
  const limLoc = val(pr, 'm:limLoc') ?? (integral ? 'subSup' : 'undOvr');
  const display = ctx.display && ctx.level === 0;
  const color = ctrlColor(pr) ?? ctx.color;
  const body = row(firstChild(el, 'm:e'), ctx);
  const subEl = firstChild(el, 'm:sub');
  const supEl = firstChild(el, 'm:sup');
  const sub = !onOff(pr, 'm:subHide', false) && hasContent(subEl) ? row(subEl, scriptStyle(ctx, true)).box : undefined;
  const sup = !onOff(pr, 'm:supHide', false) && hasContent(supEl) ? row(supEl, scriptStyle(ctx)).box : undefined;
  const op = bigOperator(chr, ctx, display, onOff(pr, 'm:grow', false) ? body.box : undefined, color);
  const s = em(ctx);
  // Limits go under and over in display style; inline equations put them at the side, as Word does.
  const opBox = limLoc === 'undOvr' && display ? stackLimits(op.box, sub, sup, 0.12 * s, 0.1 * s) : attachScripts(op.box, sub, sup, ctx, { subKern: op.subKern });
  const space = body.first ? spaceBetween('op', body.first, ctx) : 0;
  const b = new Builder();
  const opWidth = opBox.width;
  b.put(opBox, 0, 0);
  b.put(body.box, opWidth + space, 0);
  return { box: b.done(), left: 'op', right: body.last ?? 'op' };
}

// ---------------------------------------------------------------------------------------------
// Delimiters.

/**
 * A delimiter covering content of the given ascent and descent: the font's glyph when it is
 * tall enough, otherwise a shape drawn at the needed height (centred on the axis unless `match`).
 */
function delimiterBox(ch: string, ascent: number, descent: number, ctx: Ctx, grow: boolean, match: boolean, color: string): Box {
  if (!ch) return emptyBox();
  const s = em(ctx);
  const axis = ctx.axis * s;
  const cp = ch.codePointAt(0)!;
  const shape = delimiterOf(ch);
  const font = ctx.fonts.pick(cp, false, false);
  const half = Math.max(ascent - axis, descent + axis);
  let top = match ? ascent : axis + half;
  let bottom = match ? descent : half - axis;
  if (font.has(cp)) {
    const ink = inkOf(font, cp);
    const tolerance = 0.15 * s;
    const fits = !ink || (top <= ink.yMax * s + tolerance && bottom <= -ink.yMin * s + tolerance);
    if (!shape || !grow || fits) return symbolBox(ch, s, color, ctx.fonts);
  } else if (!shape) return symbolBox(ch, s, color, ctx.fonts);
  if (!shape) return emptyBox();
  const pad = 0.06 * s;
  top = Math.max(top + pad, axis + 0.5 * s);
  bottom = Math.max(bottom + pad, 0.5 * s - axis);
  if (!grow) {
    top = axis + 0.5 * s;
    bottom = 0.5 * s - axis;
  }
  const L = top + bottom;
  const W = delimiterWidth(shape.kind, L, s);
  const bearing = (shape.kind === 'bar' || shape.kind === 'dbar' ? 0.08 : 0.05) * s;
  const frame = verticalFrame(bearing, W, -bottom, shape.side === 'right' ? 'right' : 'left');
  const item = filled(polygons(mapShape(delimiterShape(shape.kind, L, W, s), frame)), color);
  return { width: W + 2 * bearing, ascent: top, descent: Math.max(0, bottom), items: [item], italic: 0 };
}

function delimited(el: XmlElement, ctx: Ctx): Atom {
  const pr = firstChild(el, 'm:dPr');
  const beg = charProp(pr, 'm:begChr', '(');
  const end = charProp(pr, 'm:endChr', ')');
  const sep = charProp(pr, 'm:sepChr', '|');
  const grow = onOff(pr, 'm:grow', true);
  const match = val(pr, 'm:shp') === 'match';
  const color = ctrlColor(pr) ?? ctx.color;
  const s = em(ctx);
  const parts = elements(el, 'm:e').map((e) => row(e, ctx));
  const ascent = Math.max(0, ...parts.map((p) => p.box.ascent));
  const descent = Math.max(0, ...parts.map((p) => p.box.descent));
  const b = new Builder();
  let x = 0;
  const open = delimiterBox(beg, ascent, descent, ctx, grow, match, color);
  b.put(open, 0, 0);
  x += open.width;
  parts.forEach((p, i) => {
    if (i > 0 && sep) {
      const bar = delimiterBox(sep, ascent, descent, ctx, grow, match, color);
      x += 0.05 * s;
      b.put(bar, x, 0);
      x += bar.width + 0.05 * s;
    }
    const w = p.box.width;
    b.put(p.box, x, 0);
    x += w;
  });
  const close = delimiterBox(end, ascent, descent, ctx, grow, match, color);
  b.put(close, x, 0);
  x += close.width;
  return { box: b.done(x), left: beg ? 'open' : (parts[0]?.first ?? 'ord'), right: end ? 'close' : (parts[parts.length - 1]?.last ?? 'ord') };
}

// ---------------------------------------------------------------------------------------------
// Matrices and equation arrays.

/** Gap between rows (pt) from m:rSpRule / m:rSp: single, 1.5, double, exactly (twips), multiple (half lines). */
function rowGap(rule: number, value: number, s: number): number {
  switch (rule) {
    case 1:
      return 0.2 * s + 0.6 * s;
    case 2:
      return 0.2 * s + 1.2 * s;
    case 3:
      return value / 20;
    case 4:
      return Math.max(0, 0.2 * s + (value / 2 - 1) * 1.2 * s);
    default:
      return 0.2 * s;
  }
}

/** Gap between matrix columns (pt) from m:cGapRule / m:cGap. */
function columnGap(rule: number, value: number, s: number): number {
  switch (rule) {
    case 1:
      return 1.5 * s;
    case 2:
      return 2 * s;
    case 3:
      return value / 20;
    case 4:
      return (value / 2) * s;
    default:
      return s;
  }
}

/**
 * Stack rows of cells. Rows get at least a strut's height so their spacing is even; the block is
 * centred on the axis (baseJc center), or its top or bottom row sits on the baseline.
 */
function stackRows(rows: Array<Array<{ box: Box; x: number }>>, ctx: Ctx, gap: number, baseJc: string, width: number): Box {
  const s = em(ctx);
  const metrics = rows.map((cells) => ({
    ascent: Math.max(0.7 * s, ...cells.map((c) => c.box.ascent)),
    descent: Math.max(0.25 * s, ...cells.map((c) => c.box.descent)),
  }));
  const baselines: number[] = [];
  let y = 0;
  metrics.forEach((m, i) => {
    if (i > 0) y += gap;
    y += m.ascent;
    baselines.push(y);
    y += m.descent;
  });
  const height = y;
  const shift = baseJc === 'top' ? -(baselines[0] ?? 0) : baseJc === 'bot' ? -(baselines[baselines.length - 1] ?? 0) : -(ctx.axis * s + height / 2);
  const b = new Builder();
  rows.forEach((cells, i) => {
    for (const c of cells) b.put(c.box, c.x, baselines[i]! + shift);
  });
  if (rows.length) {
    b.ascent = Math.max(b.ascent, -shift);
    b.descent = Math.max(b.descent, height + shift);
  }
  return b.done(width);
}

function matrix(el: XmlElement, ctx: Ctx): Atom {
  const pr = firstChild(el, 'm:mPr');
  const s = em(ctx);
  const cellCtx: Ctx = { ...ctx, display: false, align: false };
  const cells = elements(el, 'm:mr').map((mr) => elements(mr, 'm:e').map((e) => row(e, cellCtx).box));
  const jcs: string[] = [];
  const mcs = firstChild(pr, 'm:mcs');
  for (const mc of mcs ? elements(mcs, 'm:mc') : []) {
    const mcPr = firstChild(mc, 'm:mcPr');
    const count = Math.max(1, Math.min(64, numProp(mcPr, 'm:count') ?? 1));
    const jc = val(mcPr, 'm:mcJc') ?? 'center';
    for (let k = 0; k < count; k++) jcs.push(jc);
  }
  const columns = Math.max(0, ...cells.map((r) => r.length));
  const minWidth = (numProp(pr, 'm:cSp') ?? 0) / 20;
  const widths: number[] = [];
  for (let j = 0; j < columns; j++) widths.push(Math.max(minWidth, ...cells.map((r) => r[j]?.width ?? 0)));
  const gap = columnGap(numProp(pr, 'm:cGapRule') ?? 0, numProp(pr, 'm:cGap') ?? 0, s);
  const side = 0.08 * s;
  const xs: number[] = [];
  let x = side;
  for (let j = 0; j < columns; j++) {
    xs.push(x);
    x += widths[j]! + (j < columns - 1 ? gap : 0);
  }
  const placed = cells.map((r) =>
    r.map((box, j) => {
      const jc = jcs[j] ?? 'center';
      const free = widths[j]! - box.width;
      return { box, x: xs[j]! + (jc === 'left' ? 0 : jc === 'right' ? free : free / 2) };
    }),
  );
  const box = stackRows(placed, ctx, rowGap(numProp(pr, 'm:rSpRule') ?? 0, numProp(pr, 'm:rSp') ?? 0, s), val(pr, 'm:baseJc') ?? 'center', x + side);
  return ord(box);
}

/**
 * An equation array: rows aligned at `&` marks. Segments alternate right- and left-aligned
 * columns (like LaTeX's align), so "x&=1" lines up the equals signs.
 */
function eqArray(el: XmlElement, ctx: Ctx): Atom {
  const pr = firstChild(el, 'm:eqArrPr');
  const s = em(ctx);
  const rows = elements(el, 'm:e').map((e) => {
    const atoms = rowAtoms(e, { ...ctx, align: true });
    resolveBinaries(atoms);
    const segments: Atom[][] = [[]];
    for (const a of atoms) {
      if (a.kind === 'mark') segments.push([]);
      else segments[segments.length - 1]!.push(a);
    }
    // Each segment is packed knowing what precedes it, so the space before "=" survives the split.
    let before: AtomClass | undefined;
    return segments.map((seg) => {
      const r = hpack(seg, ctx, before);
      before = r.last ?? before;
      return r.box;
    });
  });
  const columns = Math.max(1, ...rows.map((r) => r.length));
  const aligned = columns > 1;
  const widths: number[] = [];
  for (let j = 0; j < columns; j++) widths.push(Math.max(0, ...rows.map((r) => r[j]?.width ?? 0)));
  const xs: number[] = [];
  let x = 0;
  for (const w of widths) {
    xs.push(x);
    x += w;
  }
  const placed = rows.map((r) =>
    r.map((box, j) => {
      const free = widths[j]! - box.width;
      return { box, x: xs[j]! + (aligned ? (j % 2 === 0 ? free : 0) : free / 2) };
    }),
  );
  const box = stackRows(placed, ctx, rowGap(numProp(pr, 'm:rSpRule') ?? 0, numProp(pr, 'm:rSp') ?? 0, s), val(pr, 'm:baseJc') ?? 'center', x);
  return ord(box);
}

// ---------------------------------------------------------------------------------------------
// Accents, bars, group characters, limits.

function isItalicGlyph(box: Box): boolean {
  const first = box.items[0];
  return box.simple === true && first?.type === 'text' && first.font.italic;
}

function accent(el: XmlElement, ctx: Ctx): Atom {
  const pr = firstChild(el, 'm:accPr');
  const chr = charProp(pr, 'm:chr', '̂') || '̂';
  const s = em(ctx);
  const color = ctrlColor(pr) ?? ctx.color;
  const base = row(firstChild(el, 'm:e'), { ...ctx, cramped: true });
  const bb = base.box;
  const baseWidth = bb.width;
  // Over an italic letter the accent follows the slant to the letter's top.
  const skew = isItalicGlyph(bb) ? 0.21 * bb.ascent * 0.5 : 0;
  const lift = Math.max(bb.ascent, X_HEIGHT * s) + 0.05 * s;
  const kind = accentOf(chr);
  let items: DisplayItem[];
  let x0: number;
  let w: number;
  let height: number;
  if (kind) {
    const metrics = accentMetrics(kind);
    const natural = metrics.width * s;
    w = natural;
    if (kind === 'bar' || kind === 'arrow') w = Math.max(natural, baseWidth);
    else if (metrics.stretch && baseWidth > 1.3 * natural) w = baseWidth * 0.92;
    x0 = baseWidth / 2 + skew - w / 2;
    const drawn = accentPaths(kind, x0, w, lift, s, color, arrowOf(chr));
    items = drawn.items;
    height = drawn.height;
  } else {
    // An accent we do not draw: its glyph, centred above the base.
    const glyph = symbolBox(chr, s, color, ctx.fonts);
    w = glyph.width;
    x0 = baseWidth / 2 + skew - w / 2;
    height = glyph.ascent + glyph.descent;
    shiftItems(glyph.items, x0, -(lift + glyph.descent));
    items = glyph.items;
  }
  const dx = Math.max(0, -x0);
  const b = new Builder();
  b.put(bb, dx, 0);
  shiftItems(items, dx, 0);
  for (const it of items) b.draw(it, -(lift + height), -lift);
  return { box: b.done(Math.max(dx + baseWidth, dx + x0 + w)), left: base.first ?? 'ord', right: base.last ?? 'ord' };
}

function overUnderBar(el: XmlElement, ctx: Ctx): Atom {
  const pr = firstChild(el, 'm:barPr');
  const top = (val(pr, 'm:pos') ?? 'bot') === 'top';
  const s = em(ctx);
  const color = ctrlColor(pr) ?? ctx.color;
  const base = row(firstChild(el, 'm:e'), top ? { ...ctx, cramped: true } : ctx);
  const bb = base.box;
  const th = RULE * s;
  const gap = 0.1 * s;
  const b = new Builder();
  const width = bb.width;
  if (top) {
    const y = -(Math.max(bb.ascent, X_HEIGHT * s) + gap + th);
    b.draw(rect(0, y, width, th, color), y, y + th);
  } else {
    const y = bb.descent + gap;
    b.draw(rect(0, y, width, th, color), y, y + th);
  }
  b.put(bb, 0, 0);
  return { box: b.done(width), left: base.first ?? 'ord', right: base.last ?? 'ord' };
}

function func(el: XmlElement, ctx: Ctx): Atom {
  const name = row(firstChild(el, 'm:fName'), { ...ctx, fn: true });
  const arg = row(firstChild(el, 'm:e'), ctx);
  // Function application: a thin space before an ordinary argument, none before a parenthesis.
  const space = arg.first ? spaceBetween('op', arg.first, ctx) : 0;
  const b = new Builder();
  const nameWidth = name.box.width;
  b.put(name.box, 0, 0);
  b.put(arg.box, nameWidth + space, 0);
  return { box: b.done(), left: 'op', right: arg.last ?? 'op' };
}

function limit(el: XmlElement, ctx: Ctx): Atom {
  const upper = el.name === 'm:limUpp';
  const s = em(ctx);
  const base = row(firstChild(el, 'm:e'), ctx);
  const lim = row(firstChild(el, 'm:lim'), scriptStyle(ctx, !upper)).box;
  const box = upper ? stackLimits(base.box, undefined, lim, 0.1 * s, 0) : stackLimits(base.box, lim, undefined, 0, 0.1 * s);
  return { box, left: base.first ?? 'ord', right: base.last ?? 'ord' };
}

function groupChr(el: XmlElement, ctx: Ctx): Atom {
  const pr = firstChild(el, 'm:groupChrPr');
  const chr = charProp(pr, 'm:chr', '⏟') || '⏟';
  const pos = val(pr, 'm:pos') ?? 'bot';
  // vertJc names the part aligned with the text: by default the base stays on the baseline.
  const vertJc = val(pr, 'm:vertJc') ?? (pos === 'bot' ? 'top' : 'bot');
  const s = em(ctx);
  const color = ctrlColor(pr) ?? ctx.color;
  const base = row(firstChild(el, 'm:e'), ctx);
  const bb = base.box;
  const L = Math.max(bb.width, 0.5 * s);
  const shape = groupShapeOf(chr);
  const arrow = arrowOf(chr);
  let sign: Box;
  if (shape && shape.kind !== 'arrow') {
    const depth = Math.min(0.2 * s + 0.03 * L, 0.42 * s);
    const frame = horizontalFrame(0, depth, 0, shape.outer);
    sign = { width: L, ascent: depth, descent: 0, items: [filled(polygons(mapShape(delimiterShape(shape.kind, L, depth, s), frame)), color)], italic: 0 };
  } else if (arrow) {
    const h = 0.3 * s;
    sign = { width: L, ascent: h, descent: 0, items: arrowPaths(0, L, h / 2, s, arrow, color, false), italic: 0 };
  } else if (accentOf(chr) === 'bar' || chr === '_') {
    const th = RULE * s;
    sign = { width: L, ascent: th, descent: 0, items: [rect(0, -th, L, th, color)], italic: 0 };
  } else sign = symbolBox(chr, s, color, ctx.fonts);
  const gap = 0.06 * s;
  const w = Math.max(bb.width, sign.width);
  const signHeight = sign.ascent + sign.descent;
  const signBaseline = pos === 'bot' ? bb.descent + gap + sign.ascent : -(bb.ascent + gap + sign.descent);
  const b = new Builder();
  b.put(bb, (w - bb.width) / 2, 0);
  b.put(sign, (w - sign.width) / 2, signBaseline);
  const box = b.done(w);
  const signOnBaseline = (pos === 'bot' && vertJc === 'bot') || (pos === 'top' && vertJc === 'top');
  if (signOnBaseline) {
    // The character sits on the text line (centred on the axis) and the base moves above or below it.
    const signCentre = signBaseline - sign.ascent + signHeight / 2;
    const dy = -ctx.axis * s - signCentre;
    shiftItems(box.items, 0, dy);
    const ascent = box.ascent - dy;
    const descent = box.descent + dy;
    box.ascent = Math.max(0, ascent);
    box.descent = Math.max(0, descent);
    const cls: AtomClass = arrow ? 'rel' : 'ord';
    return { box, left: cls, right: cls };
  }
  return ord(box);
}

// ---------------------------------------------------------------------------------------------
// Boxes, borders, phantoms.

function boxed(el: XmlElement, ctx: Ctx): Atom {
  const r = row(firstChild(el, 'm:e'), ctx);
  // An operator emulator spaces like the operator it holds ("==" as one relation).
  if (onOff(firstChild(el, 'm:boxPr'), 'm:opEmu', false) && r.first) return { box: r.box, left: r.first, right: r.last ?? r.first };
  return ord(r.box);
}

function borderBox(el: XmlElement, ctx: Ctx): Atom {
  const pr = firstChild(el, 'm:borderBoxPr');
  const s = em(ctx);
  const color = ctrlColor(pr) ?? ctx.color;
  const inner = row(firstChild(el, 'm:e'), ctx).box;
  const pad = 0.1 * s;
  const t = 0.05 * s;
  const w = inner.width + 2 * pad;
  const top = -(Math.max(inner.ascent, X_HEIGHT * s) + pad);
  const bottom = inner.descent + pad;
  const b = new Builder();
  b.put(inner, pad, 0);
  const line = (x1: number, y1: number, x2: number, y2: number): LineItem => ({ type: 'line', x1, y1, x2, y2, width: t, color, style: 'solid' });
  const add = (on: boolean, l: LineItem) => {
    if (on) b.draw(l, Math.min(l.y1, l.y2) - t / 2, Math.max(l.y1, l.y2) + t / 2);
  };
  add(!onOff(pr, 'm:hideTop', false), line(0, top, w, top));
  add(!onOff(pr, 'm:hideBot', false), line(0, bottom, w, bottom));
  add(!onOff(pr, 'm:hideLeft', false), line(0, top, 0, bottom));
  add(!onOff(pr, 'm:hideRight', false), line(w, top, w, bottom));
  add(onOff(pr, 'm:strikeH', false), line(0, -ctx.axis * s, w, -ctx.axis * s));
  add(onOff(pr, 'm:strikeV', false), line(w / 2, top, w / 2, bottom));
  add(onOff(pr, 'm:strikeBLTR', false), line(0, bottom, w, top));
  add(onOff(pr, 'm:strikeTLBR', false), line(0, top, w, bottom));
  return ord(b.done(w));
}

function phantom(el: XmlElement, ctx: Ctx): Atom {
  const pr = firstChild(el, 'm:phantPr');
  const r = row(firstChild(el, 'm:e'), ctx);
  const box: Box = {
    width: onOff(pr, 'm:zeroWid', false) ? 0 : r.box.width,
    ascent: onOff(pr, 'm:zeroAsc', false) ? 0 : r.box.ascent,
    descent: onOff(pr, 'm:zeroDesc', false) ? 0 : r.box.descent,
    items: onOff(pr, 'm:show', true) ? r.box.items : [],
    italic: 0,
  };
  return { box, left: r.first ?? 'ord', right: r.last ?? 'ord' };
}

// ---------------------------------------------------------------------------------------------

/** Join neighbouring text items of one font, size and colour on one baseline (fewer items to draw). */
export function mergeText(items: DisplayItem[]): DisplayItem[] {
  const out: DisplayItem[] = [];
  for (const it of items) {
    const prev = out[out.length - 1];
    if (
      prev?.type === 'text' &&
      it.type === 'text' &&
      prev.font === it.font &&
      prev.size === it.size &&
      prev.color === it.color &&
      prev.scale === it.scale &&
      Math.abs(prev.y - it.y) < 1e-6 &&
      it.xs[0] !== undefined &&
      it.xs[0] >= (prev.xs[prev.xs.length - 1] ?? -Infinity)
    ) {
      prev.text += it.text;
      prev.chars.push(...it.chars);
      prev.xs.push(...it.xs);
      prev.widths.push(...it.widths);
      prev.glyphs.push(...it.glyphs);
      continue;
    }
    out.push(it);
  }
  return out;
}
