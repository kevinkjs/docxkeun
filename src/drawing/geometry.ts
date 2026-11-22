/**
 * Turns DrawingML shape geometry (`a:prstGeom` and `a:custGeom`) into vector paths in shape-local
 * points: origin at the shape's top-left, x right, y down, spanning 0..w and 0..h.
 *
 * Guides are evaluated in EMU (12700 per point) because a custom geometry's literal coordinates are
 * EMU when its path has no `w`/`h` of its own; presets are unit-free so the choice does not affect
 * them. Arcs become cubic Béziers so every renderer (SVG, PDF) only needs M/L/C/Z.
 */
import type { PathCommand } from '../layout/types.js';
import type { XmlElement } from '../xml/dom.js';
import { attr, elements, firstChild } from '../xml/dom.js';
import { angleToRadians, evaluateFormula, evaluateGuides, resolveValue } from './formula.js';
import type { GuideDef, Guides } from './formula.js';
import { presetDefinition } from './presets.js';
import type { PathCommandDef, PathDef, PathFill, PathPoint, ShapeDefinition } from './presets.js';

export type { PathFill } from './presets.js';

/** One drawn path: how to fill it (relative to the shape fill) and whether to stroke it. */
export interface ShapePath {
  commands: PathCommand[];
  fill: PathFill;
  stroke: boolean;
}

export interface ShapeGeometry {
  paths: ShapePath[];
  /** Where the shape's text goes, in shape-local points. */
  textRect: { x: number; y: number; w: number; h: number };
}

const EMU_PER_PT = 12700;

/**
 * Evaluates a parsed definition for a `w`×`h` box given in geometry units, of which there are
 * `unitsPerPoint` to an output point.
 */
export function evaluateDefinition(def: ShapeDefinition, w: number, h: number, adjust: Readonly<Record<string, number>> = {}, unitsPerPoint = 1): ShapeGeometry {
  const g = evaluateGuides(w, h, def.avLst, def.gdLst, adjust);
  const paths = def.pathLst.map((p) => evaluatePath(p, g, w, h, unitsPerPoint));
  const tr = def.textRect;
  let textRect = { x: 0, y: 0, w: w / unitsPerPoint, h: h / unitsPerPoint };
  if (tr) {
    const l = resolveValue(tr.l, g) / unitsPerPoint;
    const t = resolveValue(tr.t, g) / unitsPerPoint;
    const r = resolveValue(tr.r, g) / unitsPerPoint;
    const b = resolveValue(tr.b, g) / unitsPerPoint;
    textRect = { x: l, y: t, w: r - l, h: b - t };
  }
  return { paths, textRect };
}

/**
 * The geometry of a preset shape (`a:prstGeom prst="…"`) sized `w`×`h` points, or undefined when
 * the preset is not known. `adjust` holds the shape's `a:avLst` overrides (see readAdjustments).
 */
export function presetGeometry(name: string, w: number, h: number, adjust?: Record<string, number>): ShapeGeometry | undefined {
  const def = presetDefinition(name);
  if (!def) return undefined;
  return evaluateDefinition(def, w * EMU_PER_PT, h * EMU_PER_PT, normaliseAdjust(def, adjust), EMU_PER_PT);
}

/**
 * Some producers write a single handle as `adj` for shapes whose handle is called `adj1` (and the
 * other way round); match it to the name the shape actually has rather than dropping it.
 */
function normaliseAdjust(def: ShapeDefinition, adjust: Record<string, number> | undefined): Record<string, number> {
  if (!adjust) return {};
  const names = new Set(def.avLst.map((g) => g.name));
  const out: Record<string, number> = {};
  for (const [k, v] of Object.entries(adjust)) {
    if (names.has(k)) out[k] = v;
    else if (k === 'adj' && names.has('adj1') && adjust['adj1'] === undefined) out['adj1'] = v;
    else if (k === 'adj1' && names.has('adj') && adjust['adj'] === undefined) out['adj'] = v;
  }
  return out;
}

/** The adjust values of an `a:prstGeom` (its `a:avLst/a:gd`, normally `fmla="val N"`). */
export function readAdjustments(prstGeom: XmlElement): Record<string, number> {
  const out: Record<string, number> = {};
  const av = firstChild(prstGeom, 'a:avLst');
  if (!av) return out;
  for (const gd of elements(av, 'a:gd')) {
    const name = attr(gd, 'name');
    const fmla = attr(gd, 'fmla');
    if (!name || !fmla) continue;
    // Adjust values are formulas too; anything but `val` is rare, and with no guides in scope the
    // evaluator still handles constant forms.
    out[name] = evaluateFormula(fmla, new Map());
  }
  return out;
}

/** Parses an `a:custGeom` element into a definition (guides, text rectangle and paths). */
export function parseCustomGeometry(custGeom: XmlElement): ShapeDefinition {
  const guides = (list: XmlElement | undefined): GuideDef[] =>
    list ? elements(list, 'a:gd').map((gd) => ({ name: attr(gd, 'name') ?? '', fmla: attr(gd, 'fmla') ?? 'val 0' })) : [];
  const def: ShapeDefinition = {
    avLst: guides(firstChild(custGeom, 'a:avLst')),
    gdLst: guides(firstChild(custGeom, 'a:gdLst')),
    pathLst: [],
  };
  const rect = firstChild(custGeom, 'a:rect');
  if (rect) def.textRect = { l: attr(rect, 'l') ?? 'l', t: attr(rect, 't') ?? 't', r: attr(rect, 'r') ?? 'r', b: attr(rect, 'b') ?? 'b' };
  const pathLst = firstChild(custGeom, 'a:pathLst');
  for (const p of pathLst ? elements(pathLst, 'a:path') : []) def.pathLst.push(parsePathElement(p));
  return def;
}

const FILL_MODES = new Set<string>(['norm', 'none', 'lighten', 'lightenLess', 'darken', 'darkenLess']);

function xsdBool(v: string | undefined, dflt: boolean): boolean {
  if (v === undefined) return dflt;
  return v === 'true' || v === '1';
}

function parsePathElement(p: XmlElement): PathDef {
  const fill = attr(p, 'fill') ?? 'norm';
  const def: PathDef = {
    w: Number(attr(p, 'w') ?? 0) || 0,
    h: Number(attr(p, 'h') ?? 0) || 0,
    fill: FILL_MODES.has(fill) ? (fill as PathFill) : 'norm',
    stroke: xsdBool(attr(p, 'stroke'), true),
    extrusionOk: xsdBool(attr(p, 'extrusionOk'), true),
    commands: [],
  };
  const pts = (el: XmlElement): PathPoint[] => elements(el, 'a:pt').map((pt) => ({ x: attr(pt, 'x') ?? '0', y: attr(pt, 'y') ?? '0' }));
  for (const c of elements(p)) {
    switch (c.name) {
      case 'a:moveTo':
      case 'a:lnTo': {
        const pt = pts(c)[0];
        if (pt) def.commands.push({ op: c.name === 'a:moveTo' ? 'moveTo' : 'lnTo', pt });
        break;
      }
      case 'a:arcTo':
        def.commands.push({ op: 'arcTo', wR: attr(c, 'wR') ?? '0', hR: attr(c, 'hR') ?? '0', stAng: attr(c, 'stAng') ?? '0', swAng: attr(c, 'swAng') ?? '0' });
        break;
      case 'a:quadBezTo': {
        const [a, b] = pts(c);
        if (a && b) def.commands.push({ op: 'quadBezTo', pts: [a, b] });
        break;
      }
      case 'a:cubicBezTo': {
        const [a, b, d] = pts(c);
        if (a && b && d) def.commands.push({ op: 'cubicBezTo', pts: [a, b, d] });
        break;
      }
      case 'a:close':
        def.commands.push({ op: 'close' });
        break;
    }
  }
  return def;
}

/** The geometry of an `a:custGeom` for a shape sized `w`×`h` points. */
export function customGeometry(custGeom: XmlElement, w: number, h: number): ShapeGeometry {
  return evaluateDefinition(parseCustomGeometry(custGeom), w * EMU_PER_PT, h * EMU_PER_PT, {}, EMU_PER_PT);
}

/**
 * Evaluates one path. Coordinates are worked out in the path's own space (its `w`/`h`, or the
 * shape's when those are 0) and scaled into the shape at output, so an arc's radii and angles are
 * interpreted where the spec defines them.
 */
function evaluatePath(def: PathDef, g: Guides, w: number, h: number, unitsPerPoint: number): ShapePath {
  const sx = (def.w > 0 ? w / def.w : 1) / unitsPerPoint;
  const sy = (def.h > 0 ? h / def.h : 1) / unitsPerPoint;
  const v = (s: string) => resolveValue(s, g);
  const out: PathCommand[] = [];
  let cx = 0;
  let cy = 0;
  let startX = 0;
  let startY = 0;
  let open = false;
  const ensureOpen = () => {
    // PDF needs a current point before a segment; the spec starts at the subpath start after close.
    if (open) return;
    out.push({ op: 'M', x: cx * sx, y: cy * sy });
    startX = cx;
    startY = cy;
    open = true;
  };
  for (const c of def.commands as readonly PathCommandDef[]) {
    switch (c.op) {
      case 'moveTo':
        cx = v(c.pt.x);
        cy = v(c.pt.y);
        out.push({ op: 'M', x: cx * sx, y: cy * sy });
        startX = cx;
        startY = cy;
        open = true;
        break;
      case 'lnTo':
        ensureOpen();
        cx = v(c.pt.x);
        cy = v(c.pt.y);
        out.push({ op: 'L', x: cx * sx, y: cy * sy });
        break;
      case 'arcTo': {
        ensureOpen();
        const end = arcToBeziers(cx, cy, v(c.wR), v(c.hR), v(c.stAng), v(c.swAng), (x1, y1, x2, y2, x, y) => {
          out.push({ op: 'C', x1: x1 * sx, y1: y1 * sy, x2: x2 * sx, y2: y2 * sy, x: x * sx, y: y * sy });
        });
        cx = end.x;
        cy = end.y;
        break;
      }
      case 'quadBezTo': {
        ensureOpen();
        const [q, p] = c.pts;
        const qx = v(q.x);
        const qy = v(q.y);
        const px = v(p.x);
        const py = v(p.y);
        // A quadratic is the cubic whose controls sit 2/3 of the way to the quadratic control.
        out.push({
          op: 'C',
          x1: (cx + (2 / 3) * (qx - cx)) * sx,
          y1: (cy + (2 / 3) * (qy - cy)) * sy,
          x2: (px + (2 / 3) * (qx - px)) * sx,
          y2: (py + (2 / 3) * (qy - py)) * sy,
          x: px * sx,
          y: py * sy,
        });
        cx = px;
        cy = py;
        break;
      }
      case 'cubicBezTo': {
        ensureOpen();
        const [a, b, p] = c.pts;
        cx = v(p.x);
        cy = v(p.y);
        out.push({ op: 'C', x1: v(a.x) * sx, y1: v(a.y) * sy, x2: v(b.x) * sx, y2: v(b.y) * sy, x: cx * sx, y: cy * sy });
        break;
      }
      case 'close':
        if (open) {
          out.push({ op: 'Z' });
          open = false;
          cx = startX;
          cy = startY;
        }
        break;
    }
  }
  return { commands: out, fill: def.fill, stroke: def.stroke };
}

/**
 * `a:arcTo` from the current point (x0, y0): an arc of the ellipse with radii wR, hR, starting at
 * angle stAng and sweeping swAng (60000ths of a degree, clockwise with y down). The angles are
 * directions from the ellipse centre (the preset guides compute start points that way, e.g. `arc`
 * and `moon`), so they are converted to the ellipse's parametric angles before drawing. The arc is
 * emitted as cubic Béziers of at most 90° each, through `emit`; returns the end point.
 */
export function arcToBeziers(
  x0: number,
  y0: number,
  wR: number,
  hR: number,
  stAng: number,
  swAng: number,
  emit: (x1: number, y1: number, x2: number, y2: number, x: number, y: number) => void,
): { x: number; y: number } {
  if ((wR === 0 && hR === 0) || swAng === 0) return { x: x0, y: y0 };
  const st = angleToRadians(stAng);
  const sw = angleToRadians(swAng);
  const param = (a: number) => Math.atan2(wR * Math.sin(a), hR * Math.cos(a));
  const t0 = param(st);
  let dt = param(st + sw) - t0;
  // The visual→parametric map keeps whole turns, so pick the parametric sweep nearest the visual one.
  dt += 2 * Math.PI * Math.round((sw - dt) / (2 * Math.PI));
  const cxe = x0 - wR * Math.cos(t0);
  const cye = y0 - hR * Math.sin(t0);
  const n = Math.max(1, Math.ceil(Math.abs(dt) / (Math.PI / 2) - 1e-9));
  const step = dt / n;
  const k = (4 / 3) * Math.tan(step / 4);
  let px = x0;
  let py = y0;
  for (let i = 0; i < n; i++) {
    const a1 = t0 + i * step;
    const a2 = a1 + step;
    const c1 = Math.cos(a1);
    const s1 = Math.sin(a1);
    const c2 = Math.cos(a2);
    const s2 = Math.sin(a2);
    const ex = cxe + wR * c2;
    const ey = cye + hR * s2;
    emit(px - k * wR * s1, py + k * hR * c1, ex + k * wR * s2, ey - k * hR * c2, ex, ey);
    px = ex;
    py = ey;
  }
  return { x: px, y: py };
}

/**
 * The colour a path with a `lighten`/`darken` fill mode is painted, given the shape's fill colour
 * as RRGGBB: darken scales toward black (60%, darkenLess 80%), lighten mixes toward white (40%,
 * lightenLess 20%), which is how PowerPoint shades the faces of 3-D-looking presets like `cube`.
 */
export function shadeFill(color: string, mode: PathFill): string {
  const hex = color.replace(/^#/, '');
  if (!/^[0-9a-fA-F]{6}$/.test(hex)) return color;
  if (mode === 'norm' || mode === 'none') return hex.toUpperCase();
  // Darkening scales each channel; lightening moves it that fraction of the way to 255.
  const darken = mode === 'darken' ? 0.6 : mode === 'darkenLess' ? 0.8 : 1;
  const lighten = mode === 'lighten' ? 0.4 : mode === 'lightenLess' ? 0.2 : 0;
  return [0, 2, 4]
    .map((i) => Number.parseInt(hex.slice(i, i + 2), 16))
    .map((c) => c * darken + (255 - c) * lighten)
    .map((c) => Math.round(Math.min(255, Math.max(0, c))).toString(16).padStart(2, '0'))
    .join('')
    .toUpperCase();
}
