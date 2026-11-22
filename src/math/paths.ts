/**
 * Vector shapes for the parts of an equation that must stretch or that the bundled fonts lack:
 * radicals, tall delimiters, horizontal braces, integrals, big operators, accents and arrows.
 * Shapes are built with y pointing up (the way glyphs are designed) and flipped to the display
 * list's y-down coordinates when turned into path commands.
 */
import type { PathCommand, PathItem } from '../layout/types.js';
import type { AccentKind, DelimiterKind } from './symbols.js';

export type Pt = [number, number];

/** Closed polygons (y up) as path commands (y down). */
export function polygons(shapes: Pt[][]): PathCommand[] {
  const out: PathCommand[] = [];
  for (const pts of shapes) {
    if (pts.length < 2) continue;
    pts.forEach(([x, y], i) => out.push({ op: i === 0 ? 'M' : 'L', x, y: -y }));
    out.push({ op: 'Z' });
  }
  return out;
}

/** An ellipse (y up) as four cubic arcs (y down). */
export function ellipse(cx: number, cy: number, rx: number, ry: number): PathCommand[] {
  const k = 0.5523;
  const y = -cy;
  return [
    { op: 'M', x: cx + rx, y },
    { op: 'C', x1: cx + rx, y1: y - ry * k, x2: cx + rx * k, y2: y - ry, x: cx, y: y - ry },
    { op: 'C', x1: cx - rx * k, y1: y - ry, x2: cx - rx, y2: y - ry * k, x: cx - rx, y },
    { op: 'C', x1: cx - rx, y1: y + ry * k, x2: cx - rx * k, y2: y + ry, x: cx, y: y + ry },
    { op: 'C', x1: cx + rx * k, y1: y + ry, x2: cx + rx, y2: y + ry * k, x: cx + rx, y },
    { op: 'Z' },
  ];
}

export function filled(commands: PathCommand[], color: string): PathItem {
  return { type: 'path', commands, fill: color };
}

export function stroked(commands: PathCommand[], color: string, width: number): PathItem {
  return { type: 'path', commands, stroke: color, strokeWidth: width, lineCap: 'round', lineJoin: 'round' };
}

function intersect(p: Pt, d: Pt, q: Pt, e: Pt): Pt | undefined {
  const den = d[0] * e[1] - d[1] * e[0];
  if (Math.abs(den) < 1e-9) return undefined;
  const t = ((q[0] - p[0]) * e[1] - (q[1] - p[1]) * e[0]) / den;
  return [p[0] + d[0] * t, p[1] + d[1] * t];
}

/**
 * Outline of a polyline stroked with a width per segment. Joins are mitred, falling back to a
 * bevel when the mitre would reach further than `limit` half-widths (very sharp corners).
 */
export function strokeOutline(pts: Pt[], widths: number[], limit = 3): Pt[] {
  const n = pts.length;
  const dirs: Pt[] = [];
  for (let i = 0; i < n - 1; i++) {
    const [x0, y0] = pts[i]!;
    const [x1, y1] = pts[i + 1]!;
    const len = Math.hypot(x1 - x0, y1 - y0) || 1;
    dirs.push([(x1 - x0) / len, (y1 - y0) / len]);
  }
  const side = (sign: number): Pt[] => {
    const out: Pt[] = [];
    const offset = (p: Pt, seg: number): Pt => {
      const d = dirs[seg]!;
      const h = (widths[seg]! / 2) * sign;
      return [p[0] - d[1] * h, p[1] + d[0] * h];
    };
    for (let i = 0; i < n; i++) {
      const p = pts[i]!;
      if (i === 0) out.push(offset(p, 0));
      else if (i === n - 1) out.push(offset(p, n - 2));
      else {
        const a = offset(p, i - 1);
        const b = offset(p, i);
        const hit = intersect(a, dirs[i - 1]!, b, dirs[i]!);
        const reach = limit * Math.max(widths[i - 1]!, widths[i]!) * 0.5;
        if (hit && Math.hypot(hit[0] - p[0], hit[1] - p[1]) <= reach) out.push(hit);
        else out.push(a, b);
      }
    }
    return out;
  };
  return [...side(1), ...side(-1).reverse()];
}

/** Outline of a sampled curve stroked with a width that varies along it (s from 0 to 1). */
export function taperOutline(pts: Pt[], width: (s: number) => number): Pt[] {
  const n = pts.length;
  const lengths = [0];
  for (let i = 1; i < n; i++) lengths.push(lengths[i - 1]! + Math.hypot(pts[i]![0] - pts[i - 1]![0], pts[i]![1] - pts[i - 1]![1]));
  const total = lengths[n - 1] || 1;
  const left: Pt[] = [];
  const right: Pt[] = [];
  for (let i = 0; i < n; i++) {
    const a = pts[Math.max(0, i - 1)]!;
    const b = pts[Math.min(n - 1, i + 1)]!;
    const len = Math.hypot(b[0] - a[0], b[1] - a[1]) || 1;
    const nx = -(b[1] - a[1]) / len;
    const ny = (b[0] - a[0]) / len;
    const h = width(lengths[i]! / total) / 2;
    const p = pts[i]!;
    left.push([p[0] + nx * h, p[1] + ny * h]);
    right.push([p[0] - nx * h, p[1] - ny * h]);
  }
  return [...left, ...right.reverse()];
}

/** Points along a cubic Bézier, without its first point (so pieces concatenate). */
function cubic(p0: Pt, p1: Pt, p2: Pt, p3: Pt, steps: number): Pt[] {
  const out: Pt[] = [];
  for (let i = 1; i <= steps; i++) {
    const t = i / steps;
    const u = 1 - t;
    const a = u * u * u;
    const b = 3 * u * u * t;
    const c = 3 * u * t * t;
    const d = t * t * t;
    out.push([a * p0[0] + b * p1[0] + c * p2[0] + d * p3[0], a * p0[1] + b * p1[1] + c * p2[1] + d * p3[1]]);
  }
  return out;
}

function line(p0: Pt, p1: Pt, steps: number): Pt[] {
  const out: Pt[] = [];
  for (let i = 1; i <= steps; i++) out.push([p0[0] + ((p1[0] - p0[0]) * i) / steps, p0[1] + ((p1[1] - p0[1]) * i) / steps]);
  return out;
}

/** Thickness that rises from `thin` at both ends to `thick`, reaching it a fraction `ramp` in. */
function profile(thin: number, thick: number, ramp: number): (s: number) => number {
  return (s) => {
    const t = Math.min(1, Math.min(s, 1 - s) / ramp);
    return thin + (thick - thin) * Math.sin((t * Math.PI) / 2);
  };
}

// ---------------------------------------------------------------------------------------------
// Stretchy delimiters. Shapes are built in (along, across) coordinates: `along` runs the length
// of the delimiter (0..L), `across` its thickness (0 = the outer side, W = the side facing the
// content). A frame maps them onto the page for a left/right or top/bottom delimiter.

export type Frame = (along: number, across: number) => Pt;

/** A vertical delimiter occupying x0..x0+W and y bottom..bottom+L (y up). */
export function verticalFrame(x0: number, width: number, bottom: number, side: 'left' | 'right'): Frame {
  return side === 'left' ? (a, c) => [x0 + c, bottom + a] : (a, c) => [x0 + width - c, bottom + a];
}

/** A horizontal delimiter occupying x0..x0+L and y bottom..bottom+W (y up). */
export function horizontalFrame(x0: number, depth: number, bottom: number, outer: 'top' | 'bottom'): Frame {
  return outer === 'top' ? (a, c) => [x0 + a, bottom + depth - c] : (a, c) => [x0 + a, bottom + c];
}

/** Natural thickness (across) of a stretchy delimiter of length L, before side bearings. */
export function delimiterWidth(kind: DelimiterKind, L: number, em: number): number {
  const grow = Math.max(0, L / em - 1);
  switch (kind) {
    case 'paren':
      return Math.min(0.3 + 0.05 * grow, 0.55) * em;
    case 'bracket':
    case 'ceil':
    case 'floor':
      return Math.min(0.26 + 0.02 * grow, 0.36) * em;
    case 'brace':
      return Math.min(0.36 + 0.04 * grow, 0.55) * em;
    case 'angle':
      return Math.min(0.3 + 0.05 * grow, 0.5) * em;
    case 'bar':
      return 0.07 * em;
    case 'dbar':
      return 0.2 * em;
  }
}

/** Polygons of a stretchy delimiter in (along, across) coordinates. */
export function delimiterShape(kind: DelimiterKind, L: number, W: number, em: number): Pt[][] {
  const grow = Math.max(0, L / em - 1);
  switch (kind) {
    case 'paren': {
      const thick = Math.min(0.075 + 0.012 * grow, 0.11) * em;
      const thin = 0.022 * em;
      const mid = thick / 2;
      const r = Math.min(L / 2, 0.9 * em);
      let pts: Pt[];
      if (L / 2 <= r + 1e-9) {
        // One symmetric arc; the control points put its middle at `mid`.
        const c = (8 * mid - 2 * W) / 6;
        pts = [[L, W], ...cubic([L, W], [L * 0.7, c], [L * 0.3, c], [0, W], 32)];
      } else {
        // Tall: curved ends joined by a straight stretch, as in big parenthesis glyphs.
        const top = cubic([L, W], [L - r * 0.25, W - (W - mid) * 0.5], [L - r * 0.45, mid], [L - r, mid], 16);
        const straight = line([L - r, mid], [r, mid], 8);
        const bottom = cubic([r, mid], [r * 0.45, mid], [r * 0.25, W - (W - mid) * 0.5], [0, W], 16);
        pts = [[L, W], ...top, ...straight, ...bottom];
      }
      return [taperOutline(pts, profile(thin, thick, Math.min(0.5, r / L)))];
    }
    case 'bracket':
    case 'ceil':
    case 'floor': {
      const stem = Math.min(0.06 + 0.006 * grow, 0.085) * em;
      const arm = 0.045 * em;
      const top = kind !== 'floor';
      const bottom = kind !== 'ceil';
      const pts: Pt[] = [[L, 0]];
      if (top) pts.push([L, W], [L - arm, W], [L - arm, stem]);
      else pts.push([L, stem]);
      if (bottom) pts.push([arm, stem], [arm, W], [0, W]);
      else pts.push([0, stem]);
      pts.push([0, 0]);
      return [pts];
    }
    case 'brace': {
      const half = L / 2;
      const stemAt = W / 2;
      const r = Math.min(0.2 * em, half / 2);
      const thick = Math.min(0.075 + 0.008 * grow, 0.1) * em;
      const thin = 0.02 * em;
      // Upper half from the tip to the cusp; the lower half mirrors it.
      const tip: Pt = [L, W];
      const pts: Pt[] = [
        tip,
        ...cubic(tip, [L, W - (W - stemAt) * 0.55], [L - r * 0.45, stemAt], [L - r, stemAt], 10),
        ...line([L - r, stemAt], [half + r, stemAt], 6),
        ...cubic([half + r, stemAt], [half + r * 0.45, stemAt], [half, stemAt * 0.55], [half, 0], 10),
      ];
      const upper = taperOutline(pts, profile(thin, thick, 0.35));
      const lower = upper.map(([a, c]): Pt => [L - a, c]);
      return [upper, lower];
    }
    case 'angle': {
      const t = 0.05 * em;
      return [strokeOutline([[L, W - t / 2], [L / 2, t / 2], [0, W - t / 2]], [t, t], 6)];
    }
    case 'bar': {
      const t = 0.055 * em;
      const c = W / 2;
      return [[[0, c - t / 2], [L, c - t / 2], [L, c + t / 2], [0, c + t / 2]]];
    }
    case 'dbar': {
      const t = 0.05 * em;
      const gap = 0.09 * em;
      const c = W / 2;
      return [c - gap / 2 - t / 2, c + gap / 2 + t / 2].map((m): Pt[] => [[0, m - t / 2], [L, m - t / 2], [L, m + t / 2], [0, m + t / 2]]);
    }
  }
}

export function mapShape(shape: Pt[][], frame: Frame): Pt[][] {
  return shape.map((pts) => pts.map(([a, c]) => frame(a, c)));
}

// ---------------------------------------------------------------------------------------------
// Radical sign: a short rising hook, a thick downstroke, a thin upstroke and the overbar, all one
// polygon so the joins are clean. Coordinates y up; the sign starts at x0, its vertex at `bottom`.

export function radicalShape(x0: number, bottom: number, signWidth: number, top: number, barEnd: number, rule: number, em: number): Pt[] {
  const height = top - bottom;
  const hook = Math.min(0.48 * em, height * 0.5);
  const thick = 0.085 * em;
  const p0: Pt = [x0 + 0.02 * em, bottom + hook * 0.86];
  const p1: Pt = [x0 + 0.12 * em, bottom + hook];
  const p2: Pt = [x0 + Math.min(0.29 * em, signWidth * 0.5), bottom + 0.05 * em];
  const p3: Pt = [x0 + signWidth, top - rule / 2];
  const p4: Pt = [barEnd, top - rule / 2];
  return strokeOutline([p0, p1, p2, p3, p4], [rule * 0.9, thick, rule, rule], 4);
}

// ---------------------------------------------------------------------------------------------
// Integral signs: a slanted stroke with curled ends and ball terminals, scaled to any height.

/** Polygons and terminal circles of an integral in the box x0..x0+W, bottom..bottom+H (y up). */
export function integralShape(x0: number, bottom: number, W: number, H: number, em: number): { outline: Pt[]; balls: Array<[number, number, number]> } {
  const P = (x: number, y: number): Pt => [x0 + x * W, bottom + y * H];
  const start = P(0.98, 0.9);
  const pts: Pt[] = [
    start,
    ...cubic(start, P(0.95, 1.0), P(0.78, 1.02), P(0.68, 0.88), 12),
    ...line(P(0.68, 0.88), P(0.32, 0.12), 12),
    ...cubic(P(0.32, 0.12), P(0.22, -0.02), P(0.05, 0.0), P(0.02, 0.1), 12),
  ];
  const big = H > 1.6 * em;
  const thick = (big ? 0.085 : 0.07) * em;
  const outline = taperOutline(pts, profile(0.02 * em, thick, 0.3));
  const r = (big ? 0.05 : 0.04) * em;
  const end = pts[pts.length - 1]!;
  return { outline, balls: [[start[0] - r * 0.4, start[1] - r * 0.2, r], [end[0] + r * 0.4, end[1] + r * 0.2, r]] };
}

/** Big operators drawn as shapes (y up) in x0..x0+W, bottom..bottom+H. Returns undefined for unknown ones. */
export function bigOperatorPaths(ch: string, x0: number, bottom: number, W: number, H: number, em: number, color: string): PathItem[] | undefined {
  const t = Math.max(0.065, Math.min(0.09, H / em / 18)) * em;
  const top = bottom + H;
  const rect = (xa: number, ya: number, xb: number, yb: number): Pt[] => [[xa, ya], [xb, ya], [xb, yb], [xa, yb]];
  switch (ch) {
    case '∏':
    case '∐': {
      const bar = t * 0.8;
      const stem = t * 1.15;
      const serif = 0.08 * em;
      const inset = 0.1 * W;
      const shapes: Pt[][] = [
        rect(x0, top - bar, x0 + W, top),
        rect(x0 + inset, bottom, x0 + inset + stem, top),
        rect(x0 + W - inset - stem, bottom, x0 + W - inset, top),
        rect(x0 + inset - serif, bottom, x0 + inset + stem + serif, bottom + bar * 0.7),
        rect(x0 + W - inset - stem - serif, bottom, x0 + W - inset + serif, bottom + bar * 0.7),
      ];
      // ∐ is ∏ upside down.
      const flipped = ch === '∐' ? shapes.map((pts) => pts.map(([x, y]): Pt => [x, bottom + top - y])) : shapes;
      return [filled(polygons(flipped), color)];
    }
    case '∑': {
      const pts = strokeOutline(
        [[x0 + W, top - t * 0.4], [x0 + t * 0.4, top - t * 0.4], [x0 + W * 0.55, bottom + H / 2], [x0 + t * 0.4, bottom + t * 0.4], [x0 + W, bottom + t * 0.4]],
        [t * 0.8, t * 1.2, t * 0.7, t * 0.8],
        8,
      );
      return [filled(polygons([pts]), color)];
    }
    case '⋃':
    case '⋂': {
      const r = W / 2 - t / 2;
      const cx = x0 + W / 2;
      const up = ch === '⋃';
      const yEnd = up ? top : bottom;
      const yTurn = up ? bottom + t / 2 + r : top - t / 2 - r;
      const yTip = up ? bottom + t / 2 : top - t / 2;
      const k = 0.5523 * r;
      const d: PathCommand[] = [
        { op: 'M', x: cx - r, y: -yEnd },
        { op: 'L', x: cx - r, y: -yTurn },
        { op: 'C', x1: cx - r, y1: -(yTurn + (up ? -k : k)), x2: cx - k, y2: -yTip, x: cx, y: -yTip },
        { op: 'C', x1: cx + k, y1: -yTip, x2: cx + r, y2: -(yTurn + (up ? -k : k)), x: cx + r, y: -yTurn },
        { op: 'L', x: cx + r, y: -yEnd },
      ];
      return [{ ...stroked(d, color, t), lineCap: 'butt' }];
    }
    case '⋁':
    case '⋀': {
      const down = ch === '⋁';
      const pts: Pt[] = down
        ? [[x0 + t / 2, top], [x0 + W / 2, bottom + t / 2], [x0 + W - t / 2, top]]
        : [[x0 + t / 2, bottom], [x0 + W / 2, top - t / 2], [x0 + W - t / 2, bottom]];
      return [filled(polygons([strokeOutline(pts, [t * 0.6, t * 1.1], 8)]), color)];
    }
    case '⨀':
    case '⨁':
    case '⨂': {
      const r = Math.min(W, H) / 2 - t / 2;
      const cx = x0 + W / 2;
      const cy = bottom + H / 2;
      const items: PathItem[] = [stroked(ellipse(cx, cy, r, r), color, t * 0.8)];
      if (ch === '⨀') items.push(filled(ellipse(cx, cy, t * 1.2, t * 1.2), color));
      else {
        const d = ch === '⨁' ? r : r * Math.SQRT1_2;
        const segs: Array<[Pt, Pt]> = ch === '⨁' ? [[[cx - d, cy], [cx + d, cy]], [[cx, cy - d], [cx, cy + d]]] : [[[cx - d, cy - d], [cx + d, cy + d]], [[cx - d, cy + d], [cx + d, cy - d]]];
        for (const [a, b] of segs) items.push({ ...stroked([{ op: 'M', x: a[0], y: -a[1] }, { op: 'L', x: b[0], y: -b[1] }], color, t * 0.8), lineCap: 'butt' });
      }
      return items;
    }
    default:
      return undefined;
  }
}

// ---------------------------------------------------------------------------------------------
// Arrows (stretchy group characters and vector accents), y up, from x0 to x1 centred on y.

export function arrowPaths(x0: number, x1: number, y: number, em: number, spec: { left: boolean; right: boolean; double: boolean; harpoon: boolean }, color: string, small: boolean): PathItem[] {
  const t = (small ? 0.04 : 0.05) * em;
  const headLength = (small ? 0.16 : 0.24) * em;
  const headHalf = (small ? 0.075 : 0.12) * em * (spec.double ? 1.5 : 1);
  const shapes: Pt[][] = [];
  const notch = headLength * (spec.double ? 0.45 : 0.6);
  const sx0 = spec.left ? x0 + notch : x0;
  const sx1 = spec.right ? x1 - notch : x1;
  const shafts = spec.double ? [y + 0.08 * em, y - 0.08 * em] : [y];
  for (const sy of shafts) shapes.push([[sx0, sy - t / 2], [sx1, sy - t / 2], [sx1, sy + t / 2], [sx0, sy + t / 2]]);
  const head = (tip: number, dir: 1 | -1): Pt[] => {
    const back = tip - dir * headLength;
    const inner = tip - dir * notch;
    if (spec.harpoon) return [[tip, y - t / 2], [back, y + headHalf], [inner, y + t / 2], [back + dir * 0.0, y - t / 2]];
    return [[tip, y], [back, y + headHalf], [inner, y], [back, y - headHalf]];
  };
  if (spec.right) shapes.push(head(x1, 1));
  if (spec.left) shapes.push(head(x0, -1));
  return [filled(polygons(shapes), color)];
}

// ---------------------------------------------------------------------------------------------
// Accents, drawn above a base from x0 over width w with their lowest ink at y (y up).

/** Natural width of an accent (em), and whether it stretches over wide bases. */
export function accentMetrics(kind: AccentKind): { width: number; stretch: boolean } {
  switch (kind) {
    case 'hat':
    case 'check':
      return { width: 0.32, stretch: true };
    case 'tilde':
      return { width: 0.34, stretch: true };
    case 'bar':
    case 'arrow':
      return { width: 0.36, stretch: true };
    case 'ddot':
      return { width: 0.26, stretch: false };
    case 'dddot':
      return { width: 0.42, stretch: false };
    default:
      return { width: 0.2, stretch: false };
  }
}

/** Draw an accent; returns the items and the accent's ink height. */
export function accentPaths(kind: AccentKind, x0: number, w: number, y: number, em: number, color: string, arrow?: { left: boolean; right: boolean; double: boolean; harpoon: boolean }): { items: PathItem[]; height: number } {
  const t = 0.045 * em;
  const cx = x0 + w / 2;
  switch (kind) {
    case 'hat':
    case 'check': {
      const h = Math.min(0.13 * em + 0.04 * w, 0.3 * em);
      const pts: Pt[] = kind === 'hat' ? [[x0, y + t / 2], [cx, y + h - t / 2], [x0 + w, y + t / 2]] : [[x0, y + h - t / 2], [cx, y + t / 2], [x0 + w, y + h - t / 2]];
      return { items: [filled(polygons([strokeOutline(pts, [t * 0.8, t * 1.2], 8)]), color)], height: h };
    }
    case 'tilde': {
      const h = Math.min(0.09 * em + 0.025 * w, 0.2 * em);
      const pts: Pt[] = [];
      for (let i = 0; i <= 24; i++) {
        const s = i / 24;
        pts.push([x0 + s * w, y + t / 2 + ((h - t) / 2) * (1 + Math.sin(2 * Math.PI * s - Math.PI * 0.15))]);
      }
      return { items: [filled(polygons([taperOutline(pts, profile(t * 0.5, t * 1.2, 0.3))]), color)], height: h };
    }
    case 'bar':
      return { items: [filled(polygons([[[x0, y], [x0 + w, y], [x0 + w, y + t], [x0, y + t]]]), color)], height: t };
    case 'dot':
    case 'ddot':
    case 'dddot': {
      const r = 0.05 * em;
      const n = kind === 'dot' ? 1 : kind === 'ddot' ? 2 : 3;
      const gap = 0.16 * em;
      const commands: PathCommand[] = [];
      for (let i = 0; i < n; i++) commands.push(...ellipse(cx + (i - (n - 1) / 2) * gap, y + r, r, r));
      return { items: [filled(commands, color)], height: 2 * r };
    }
    case 'acute':
    case 'grave': {
      const h = 0.16 * em;
      const pts: Pt[] = kind === 'acute' ? [[cx - 0.05 * em, y], [cx + 0.07 * em, y + h]] : [[cx + 0.05 * em, y], [cx - 0.07 * em, y + h]];
      return { items: [filled(polygons([strokeOutline(pts, [t * 1.4])]), color)], height: h };
    }
    case 'breve': {
      const h = 0.12 * em;
      const r = 0.12 * em;
      const pts: Pt[] = [];
      for (let i = 0; i <= 16; i++) {
        const a = Math.PI + (Math.PI * i) / 16;
        pts.push([cx + r * Math.cos(a), y + h + h * Math.sin(a) * 0.9]);
      }
      return { items: [filled(polygons([taperOutline(pts, profile(t * 0.5, t * 1.2, 0.4))]), color)], height: h };
    }
    case 'ring': {
      const r = 0.075 * em;
      return { items: [stroked(ellipse(cx, y + r, r - t / 3, r - t / 3), color, t * 0.7)], height: 2 * r };
    }
    case 'arrow': {
      const spec = arrow ?? { left: false, right: true, double: false, harpoon: false };
      const h = 0.15 * em;
      return { items: arrowPaths(x0, x0 + w, y + h / 2, em, spec, color, true), height: h };
    }
  }
}
