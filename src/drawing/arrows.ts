/**
 * Line end decorations (`a:headEnd` / `a:tailEnd`, ECMA-376 §20.1.2.2.15 and §20.1.2.2.37): the
 * arrowheads, diamonds and dots drawn where a line or open path stops.
 *
 * The spec only names the kinds (triangle, stealth, diamond, oval, arrow) and three relative sizes;
 * the proportions here follow what Office draws: each size is a multiple of the line width
 * (sm 2×, med 3×, lg 5×), measured against a floor so that heads on hairlines stay visible.
 */
import type { PathCommand } from '../layout/types.js';
import type { XmlElement } from '../xml/dom.js';
import { attr } from '../xml/dom.js';

export type LineEndType = 'none' | 'triangle' | 'stealth' | 'diamond' | 'oval' | 'arrow';
export type LineEndSize = 'sm' | 'med' | 'lg';

export interface LineEndSpec {
  type: LineEndType;
  w: LineEndSize;
  len: LineEndSize;
}

export interface LineEndShape {
  /** The decoration in the same coordinates as `tip`; empty for `none`. */
  commands: PathCommand[];
  /** Filled with the line colour (true) or stroked with the line's pen (the open `arrow`). */
  filled: boolean;
  /** How far to pull the line's end back from `tip` so its butt end hides inside the decoration. */
  inset: number;
}

interface Point {
  x: number;
  y: number;
}

const TYPES = new Set<string>(['none', 'triangle', 'stealth', 'diamond', 'oval', 'arrow']);
const FACTOR: Record<LineEndSize, number> = { sm: 2, med: 3, lg: 5 };
/**
 * Below about 2pt Office stops shrinking heads with the line; LibreOffice uses the same floor
 * (0.7 mm) when it imports these files.
 */
const MIN_BASE_WIDTH = 2;
/** Where the stealth head's back notch sits, as a fraction of its length from the tip. */
const STEALTH_NOTCH = 2 / 3;
/** Cubic Bézier circle constant: control distance for a quarter arc of a unit circle. */
const KAPPA = 0.5522847498307936;

function size(v: string | undefined): LineEndSize {
  return v === 'sm' || v === 'lg' ? v : 'med';
}

/** Reads an `a:headEnd` or `a:tailEnd`; undefined when absent or `type="none"`. */
export function readLineEnd(el: XmlElement | undefined): LineEndSpec | undefined {
  if (!el) return undefined;
  const type = attr(el, 'type') ?? 'none';
  if (!TYPES.has(type) || type === 'none') return undefined;
  return { type: type as LineEndType, w: size(attr(el, 'w')), len: size(attr(el, 'len')) };
}

/**
 * The decoration for a line ending at `tip` and arriving from `from` (the previous point on the
 * line, which gives the direction). `w` and `len` are the spec's sm/med/lg sizes.
 */
export function lineEnd(type: string, w: string | undefined, len: string | undefined, tip: Point, from: Point, lineWidth: number): LineEndShape {
  const none: LineEndShape = { commands: [], filled: false, inset: 0 };
  if (!TYPES.has(type) || type === 'none') return none;
  const lw = Math.max(0, lineWidth);
  const base = Math.max(lw, MIN_BASE_WIDTH);
  const W = FACTOR[size(w)] * base;
  const L = FACTOR[size(len)] * base;

  // Unit vector along the line toward the tip, and its normal; a degenerate line points right.
  let dx = tip.x - from.x;
  let dy = tip.y - from.y;
  const d = Math.hypot(dx, dy);
  if (d > 0) {
    dx /= d;
    dy /= d;
  } else {
    dx = 1;
    dy = 0;
  }
  const nx = -dy;
  const ny = dx;
  /** A point `along` the line from the tip (negative = back toward `from`) and `across` it. */
  const at = (along: number, across: number): Point => ({ x: tip.x + dx * along + nx * across, y: tip.y + dy * along + ny * across });
  const poly = (pts: Point[], close: boolean): PathCommand[] => {
    const out: PathCommand[] = pts.map((p, i) => ({ op: i === 0 ? 'M' : 'L', x: p.x, y: p.y }));
    if (close) out.push({ op: 'Z' });
    return out;
  };

  switch (type) {
    case 'triangle':
      // Stop the line half a pen width inside the base so no seam shows; the base is at least
      // twice the pen width, so the line's corners stay inside the head.
      return { commands: poly([at(0, 0), at(-L, W / 2), at(-L, -W / 2)], true), filled: true, inset: Math.max(0, L - lw / 2) };
    case 'stealth':
      return {
        commands: poly([at(0, 0), at(-L, W / 2), at(-L * STEALTH_NOTCH, 0), at(-L, -W / 2)], true),
        filled: true,
        inset: L * STEALTH_NOTCH,
      };
    case 'diamond':
      // Diamonds and ovals are centred on the end point and cover it, so the line keeps its length.
      return { commands: poly([at(L / 2, 0), at(0, W / 2), at(-L / 2, 0), at(0, -W / 2)], true), filled: true, inset: 0 };
    case 'oval': {
      const a = L / 2;
      const b = W / 2;
      const k = KAPPA;
      const p = (along: number, across: number) => at(along, across);
      const c = (c1: Point, c2: Point, e: Point): PathCommand => ({ op: 'C', x1: c1.x, y1: c1.y, x2: c2.x, y2: c2.y, x: e.x, y: e.y });
      const start = p(a, 0);
      return {
        commands: [
          { op: 'M', x: start.x, y: start.y },
          c(p(a, k * b), p(k * a, b), p(0, b)),
          c(p(-k * a, b), p(-a, k * b), p(-a, 0)),
          c(p(-a, -k * b), p(-k * a, -b), p(0, -b)),
          c(p(k * a, -b), p(a, -k * b), p(a, 0)),
          { op: 'Z' },
        ],
        filled: true,
        inset: 0,
      };
    }
    case 'arrow':
      // An open chevron stroked with the line's own pen; the line runs into its joint.
      return { commands: poly([at(-L, W / 2), at(0, 0), at(-L, -W / 2)], false), filled: false, inset: Math.min(lw / 2, L) };
    default:
      return none;
  }
}
