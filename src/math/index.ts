/**
 * Office Math (OMML) typesetting: lays out an m:oMath or m:oMathPara into a box of display items,
 * and renders it as linear-format text. Line layout places the box like an inline object.
 */
import type { DisplayItem, FontProvider } from '../layout/types.js';
import { elements, firstChild } from '../xml/dom.js';
import type { XmlElement } from '../xml/dom.js';
import { Builder } from './boxes.js';
import type { Ctx } from './boxes.js';
import { mergeText, row } from './layout.js';
import { val } from './omml.js';
import { MathFonts, axisHeight } from './text.js';

export { mathToText } from './linear.js';

export interface MathLayoutOptions {
  fonts: FontProvider;
  /** Font size of the equation's text (pt); scripts and limits are scaled from it. */
  size: number;
  /** Default colour (RRGGBB); run colours (w:color) override it. */
  color: string;
  /** Display style (an equation on its own line): big operators, limits above and below, full-size fractions. */
  display: boolean;
}

export interface MathLayout {
  width: number;
  /**
   * Ink extent above and below the baseline. Line layout should take the larger of these and
   * the paragraph font's own ascent and descent, as it would for any inline object.
   */
  ascent: number;
  descent: number;
  /** Items relative to the left end of the (first) baseline: x right, y down. */
  items: DisplayItem[];
}

function rootContext(options: MathLayoutOptions): Ctx {
  const fonts = new MathFonts(options.fonts);
  return {
    fonts,
    base: options.size,
    level: 0,
    display: options.display,
    cramped: false,
    color: options.color,
    axis: axisHeight(fonts),
    fn: false,
    align: false,
  };
}

/** Lay out an equation. An m:oMathPara stacks its equations, aligned by m:oMathParaPr/m:jc. */
export function layoutMath(node: XmlElement, options: MathLayoutOptions): MathLayout {
  const ctx = rootContext(options);
  const lines = node.name === 'm:oMathPara' ? elements(node, 'm:oMath') : [node];
  const boxes = lines.map((line) => row(line, ctx).box);
  if (boxes.length === 1) {
    const box = boxes[0]!;
    return { width: box.width, ascent: box.ascent, descent: box.descent, items: mergeText(box.items) };
  }
  const jc = val(firstChild(node, 'm:oMathParaPr'), 'm:jc') ?? 'centerGroup';
  const width = Math.max(0, ...boxes.map((b) => b.width));
  const size = options.size;
  const b = new Builder();
  let y = 0;
  boxes.forEach((box, i) => {
    const prev = boxes[i - 1];
    // Baselines at least a single line apart (Cambria's line height), further when the equations are tall.
    if (prev) y += Math.max(prev.descent + 0.3 * size + box.ascent, 1.3 * size);
    const x = jc === 'left' ? 0 : jc === 'right' ? width - box.width : (width - box.width) / 2;
    b.put(box, x, y);
  });
  return { width, ascent: b.ascent, descent: b.descent, items: mergeText(b.items) };
}
