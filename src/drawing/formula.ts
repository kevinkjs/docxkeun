/**
 * DrawingML shape guide formulas (ECMA-376 §20.1.9.11).
 *
 * Preset and custom shapes describe their geometry as an ordered list of named guides, each one
 * operator applied to numeric literals or to earlier guides. This module evaluates them. Angles are
 * in 60000ths of a degree, as in the files, and lengths are in whatever unit the caller chose for
 * `w` and `h` (the formulas are unit-agnostic as long as every length shares that unit).
 */

/** Guide values by name, in evaluation order. */
export type Guides = Map<string, number>;

/** A named formula: an `a:gd` element. */
export interface GuideDef {
  name: string;
  fmla: string;
}

const RAD_PER_UNIT = Math.PI / (180 * 60000);

/** 60000ths of a degree → radians. */
export function angleToRadians(a: number): number {
  return a * RAD_PER_UNIT;
}

/** Radians → 60000ths of a degree. */
export function radiansToAngle(r: number): number {
  return r / RAD_PER_UNIT;
}

const NUMBER = /^[-+]?(\d+\.?\d*|\.\d+)(e[-+]?\d+)?$/i;

/**
 * The value of one formula argument: a guide name or a numeric literal. Guides are looked up first
 * because some preset guide names start with a digit (`2a1`, `3cd4`) and would otherwise half-parse
 * as numbers. Unknown names count as 0 rather than throwing, so one bad guide in a document cannot
 * stop the rest of the page from drawing.
 */
export function resolveValue(token: string, guides: Guides): number {
  const v = guides.get(token);
  if (v !== undefined) return v;
  if (NUMBER.test(token)) return Number(token);
  return 0;
}

/**
 * The guides every shape gets for free: the shape box, its fractions, and common angles. The spec
 * lists a fixed set (wd2 … wd32 and so on); a few extra divisors are defined too since they cost
 * nothing and some producers use them.
 */
export function builtinGuides(w: number, h: number): Guides {
  const ss = Math.min(w, h);
  const ls = Math.max(w, h);
  const g: Guides = new Map<string, number>([
    ['l', 0],
    ['t', 0],
    ['r', w],
    ['b', h],
    ['w', w],
    ['h', h],
    ['hc', w / 2],
    ['vc', h / 2],
    ['ss', ss],
    ['ls', ls],
    ['cd2', 10800000],
    ['cd4', 5400000],
    ['cd8', 2700000],
    ['3cd4', 16200000],
    ['3cd8', 8100000],
    ['5cd8', 13500000],
    ['7cd8', 18900000],
  ]);
  for (const d of [2, 3, 4, 5, 6, 8, 10, 12, 16, 32]) {
    g.set(`wd${d}`, w / d);
    g.set(`hd${d}`, h / d);
    g.set(`ssd${d}`, ss / d);
  }
  return g;
}

/** x·y/z and (x+y)/z with a zero divisor give 0: a collapsed (zero-size) shape must still evaluate. */
function div(n: number, d: number): number {
  return d === 0 ? 0 : n / d;
}

/** Evaluates one formula such as `+- r 0 x1` against the guides defined so far. */
export function evaluateFormula(fmla: string, guides: Guides): number {
  const parts = fmla.trim().split(/\s+/);
  const op = parts[0] ?? '';
  const arg = (i: number): number => {
    const t = parts[i + 1];
    return t === undefined ? 0 : resolveValue(t, guides);
  };
  const x = arg(0);
  let v: number;
  switch (op) {
    case '*/':
      v = div(x * arg(1), arg(2));
      break;
    case '+-':
      v = x + arg(1) - arg(2);
      break;
    case '+/':
      v = div(x + arg(1), arg(2));
      break;
    case '?:':
      v = x > 0 ? arg(1) : arg(2);
      break;
    case 'abs':
      v = Math.abs(x);
      break;
    case 'at2':
      // Arguments are (x, y); the result is the angle of that vector.
      v = radiansToAngle(Math.atan2(arg(1), x));
      break;
    case 'cat2':
      v = x * Math.cos(Math.atan2(arg(2), arg(1)));
      break;
    case 'sat2':
      v = x * Math.sin(Math.atan2(arg(2), arg(1)));
      break;
    case 'cos':
      v = x * Math.cos(angleToRadians(arg(1)));
      break;
    case 'sin':
      v = x * Math.sin(angleToRadians(arg(1)));
      break;
    case 'tan':
      v = x * Math.tan(angleToRadians(arg(1)));
      break;
    case 'max':
      v = Math.max(x, arg(1));
      break;
    case 'min':
      v = Math.min(x, arg(1));
      break;
    case 'mod': {
      const y = arg(1);
      const z = arg(2);
      v = Math.sqrt(x * x + y * y + z * z);
      break;
    }
    case 'pin': {
      const y = arg(1);
      const z = arg(2);
      v = y < x ? x : y > z ? z : y;
      break;
    }
    case 'sqrt':
      v = Math.sqrt(Math.max(0, x));
      break;
    case 'val':
      v = x;
      break;
    default:
      v = 0;
  }
  return Number.isFinite(v) ? v : 0;
}

/**
 * Evaluates a shape's guides: the built-ins for a `w`×`h` box, then the adjust values (`avLst`,
 * where a value from `adjust` replaces the default formula), then `gdLst` in document order so
 * each guide can use the ones before it.
 */
export function evaluateGuides(w: number, h: number, avLst: readonly GuideDef[], gdLst: readonly GuideDef[], adjust: Readonly<Record<string, number>> = {}): Guides {
  const g = builtinGuides(w, h);
  for (const gd of avLst) {
    const override = adjust[gd.name];
    g.set(gd.name, override !== undefined && Number.isFinite(override) ? override : evaluateFormula(gd.fmla, g));
  }
  for (const gd of gdLst) g.set(gd.name, evaluateFormula(gd.fmla, g));
  return g;
}
