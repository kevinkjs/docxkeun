/**
 * Geometry for the metafile players: affine matrices laid out like a GDI XFORM, and the
 * conversion of GDI ellipses and arcs into cubic Bézier segments (the display list has no arcs).
 */

export interface Pt {
  x: number;
  y: number;
}

/** x' = a·x + c·y + e, y' = b·x + d·y + f — the same layout as XFORM (eM11, eM12, eM21, eM22, eDx, eDy). */
export interface Matrix {
  a: number;
  b: number;
  c: number;
  d: number;
  e: number;
  f: number;
}

export const IDENTITY: Matrix = { a: 1, b: 0, c: 0, d: 1, e: 0, f: 0 };

/** The transform that applies `first`, then `second` (GDI's row-vector product first × second). */
export function compose(first: Matrix, second: Matrix): Matrix {
  return {
    a: first.a * second.a + first.b * second.c,
    b: first.a * second.b + first.b * second.d,
    c: first.c * second.a + first.d * second.c,
    d: first.c * second.b + first.d * second.d,
    e: first.e * second.a + first.f * second.c + second.e,
    f: first.e * second.b + first.f * second.d + second.f,
  };
}

export function apply(m: Matrix, x: number, y: number): Pt {
  return { x: m.a * x + m.c * y + m.e, y: m.b * x + m.d * y + m.f };
}

export function determinant(m: Matrix): number {
  return m.a * m.d - m.b * m.c;
}

export function isIdentity(m: Matrix): boolean {
  return m.a === 1 && m.b === 0 && m.c === 0 && m.d === 1 && m.e === 0 && m.f === 0;
}

/** One cubic segment: two control points and the end point (the start is the previous end). */
export type Cubic = [Pt, Pt, Pt];

/**
 * Cubic segments for the elliptical arc centred on (cx, cy) with radii rx, ry, starting at
 * parameter t0 and sweeping by `sweep` radians. The point at t is (cx + rx·cos t, cy + ry·sin t),
 * so with y pointing down a positive sweep runs clockwise on screen. Segments span at most 90°,
 * where the standard 4/3·tan(θ/4) control distance stays visually exact.
 */
export function arcCubics(cx: number, cy: number, rx: number, ry: number, t0: number, sweep: number): Cubic[] {
  const n = Math.max(1, Math.ceil(Math.abs(sweep) / (Math.PI / 2) - 1e-9));
  const step = sweep / n;
  const k = (4 / 3) * Math.tan(step / 4);
  const out: Cubic[] = [];
  for (let i = 0; i < n; i++) {
    const ta = t0 + i * step;
    const tb = ta + step;
    const ca = Math.cos(ta);
    const sa = Math.sin(ta);
    const cb = Math.cos(tb);
    const sb = Math.sin(tb);
    const p0 = { x: cx + rx * ca, y: cy + ry * sa };
    const p3 = { x: cx + rx * cb, y: cy + ry * sb };
    out.push([
      { x: p0.x - k * rx * sa, y: p0.y + k * ry * ca },
      { x: p3.x + k * rx * sb, y: p3.y - k * ry * cb },
      p3,
    ]);
  }
  return out;
}

export interface Arc {
  start: Pt;
  end: Pt;
  cubics: Cubic[];
  centre: Pt;
}

/**
 * The arc GDI's Arc/Chord/Pie/ArcTo draw: on the ellipse inscribed in the box, from where the ray
 * from the centre through (xs, ys) meets it to where the ray through (xe, ye) does. `clockwise`
 * is in the y-down sense of the coordinates given. Equal rays mean a full ellipse, as in GDI.
 */
export function gdiArc(left: number, top: number, right: number, bottom: number, xs: number, ys: number, xe: number, ye: number, clockwise: boolean): Arc {
  const cx = (left + right) / 2;
  const cy = (top + bottom) / 2;
  const rx = Math.abs(right - left) / 2;
  const ry = Math.abs(bottom - top) / 2;
  // Scaling the ray direction onto the unit circle gives the ellipse parameter where the ray meets it.
  const ts = Math.atan2((ys - cy) / (ry || 1), (xs - cx) / (rx || 1));
  const te = Math.atan2((ye - cy) / (ry || 1), (xe - cx) / (rx || 1));
  let sweep = clockwise ? te - ts : ts - te;
  while (sweep <= 1e-9) sweep += 2 * Math.PI;
  while (sweep > 2 * Math.PI + 1e-9) sweep -= 2 * Math.PI;
  const signed = clockwise ? sweep : -sweep;
  const cubics = arcCubics(cx, cy, rx, ry, ts, signed);
  const start = { x: cx + rx * Math.cos(ts), y: cy + ry * Math.sin(ts) };
  const end = cubics.length ? cubics[cubics.length - 1]![2] : start;
  return { start, end, cubics, centre: { x: cx, y: cy } };
}

/** Cubic segments of a full ellipse in a box, starting at its rightmost point. */
export function ellipseCubics(left: number, top: number, right: number, bottom: number): { start: Pt; cubics: Cubic[] } {
  const cx = (left + right) / 2;
  const cy = (top + bottom) / 2;
  const rx = Math.abs(right - left) / 2;
  const ry = Math.abs(bottom - top) / 2;
  return { start: { x: cx + rx, y: cy }, cubics: arcCubics(cx, cy, rx, ry, 0, 2 * Math.PI) };
}
