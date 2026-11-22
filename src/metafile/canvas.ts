/**
 * The drawing back end both metafile players share. It keeps the device context (with the
 * SaveDC stack and the path bracket), maps logical coordinates through the whole transform chain
 * (world → page → window/viewport → frame) into page points, and turns GDI drawing calls into
 * display items. Path points are transformed when they are added, as GDI does, so a transform
 * changed halfway through a path does not move what was already recorded.
 */
import { textItem } from '../layout/paint.js';
import type { DisplayItem, FontProvider, LayoutFont, PathCommand, PathItem, RasterImage } from '../layout/types.js';
import { flipRaster } from './dib.js';
import { apply, arcCubics, compose, determinant, ellipseCubics, gdiArc } from './geometry.js';
import type { Cubic, Matrix, Pt } from './geometry.js';
import {
  AD_CLOCKWISE,
  ALTERNATE,
  OPAQUE,
  PS_ALTERNATE,
  PS_DASH,
  PS_DASHDOT,
  PS_DASHDOTDOT,
  PS_DOT,
  PS_NULL,
  PS_USERSTYLE,
  TA_BASELINE,
  TA_BOTTOM,
  TA_CENTER,
  TA_RIGHT,
  TA_UPDATECP,
  defaultState,
  mixColors,
} from './state.js';
import type { DCState } from './state.js';

/** Pen width 0 and cosmetic pens are one device pixel wide: drawn as a hairline this wide (pt). */
export const HAIRLINE = 0.25;

/** How much of the area a GDI hatch covers, for drawing hatched brushes as a flat tint. */
const HATCH_COVERAGE = 0.35;

export interface CanvasOptions {
  fonts: FontProvider;
  /** Page space (after the world transform) → page points, for the current window/viewport state. */
  pageToPoints(state: DCState): Matrix;
  /** Page points per device pixel: the unit of cosmetic pen dash patterns. */
  pixel: number;
  /** Em size in points for fonts whose height is 0 (the device's default font). */
  defaultFontSize: number;
}

type StrokeStyle = Pick<PathItem, 'stroke' | 'strokeWidth' | 'dash' | 'lineCap' | 'lineJoin'>;
type FillStyle = Pick<PathItem, 'fill' | 'fillOpacity' | 'evenOdd'>;

export type ArcKind = 'arc' | 'arcto' | 'chord' | 'pie';

const M = (p: Pt): PathCommand => ({ op: 'M', x: p.x, y: p.y });
const L = (p: Pt): PathCommand => ({ op: 'L', x: p.x, y: p.y });

export class Canvas {
  readonly items: DisplayItem[] = [];
  state: DCState = defaultState();
  private readonly saved: DCState[] = [];
  /** Commands of the path bracket, in page points; undefined when there is no path. */
  private path: PathCommand[] | undefined;
  /** Between BeginPath and EndPath drawing calls add to the path instead of drawing. */
  private recording = false;
  /** Whether the path's last figure is open and ends at the current position (so "…To" calls extend it). */
  private figureOpen = false;
  /** The last stroked open path, which a following LineTo from its end point extends instead of starting a new item. */
  private chain: { item: PathItem; end: Pt } | undefined;

  constructor(private readonly options: CanvasOptions) {}

  /** Logical → page points for the current state. */
  matrix(): Matrix {
    return compose(this.state.world, this.options.pageToPoints(this.state));
  }

  // ---- device context ----

  save(): void {
    this.saved.push({ ...this.state });
  }

  /** RestoreDC: negative counts back from the latest SaveDC, positive names a SaveDC instance (1-based). */
  restore(n: number): void {
    const index = n < 0 ? this.saved.length + n : n - 1;
    const state = this.saved[index];
    if (!state || n === 0) return;
    this.state = state;
    this.saved.length = index;
  }

  // ---- paths ----

  beginPath(): void {
    this.path = [];
    this.recording = true;
    this.figureOpen = false;
  }

  endPath(): void {
    this.recording = false;
  }

  abortPath(): void {
    this.path = undefined;
    this.recording = false;
    this.figureOpen = false;
  }

  closeFigure(): void {
    if (this.recording && this.path && this.figureOpen) {
      this.path.push({ op: 'Z' });
      this.figureOpen = false;
    }
  }

  /** FillPath / StrokePath / StrokeAndFillPath: draw the recorded path and discard it. */
  drawPath(fill: boolean, stroke: boolean): void {
    const path = this.path;
    this.abortPath();
    if (!path || !path.length) return;
    // Filling closes every open figure, and StrokeAndFillPath strokes the closed outline.
    this.output(fill ? closeFigures(path) : path, fill, stroke);
  }

  // ---- lines and curves ----

  moveTo(x: number, y: number): void {
    this.state.position = { x, y };
    if (this.recording) this.figureOpen = false;
  }

  lineTo(x: number, y: number): void {
    const m = this.matrix();
    this.openFrom(m, [L(apply(m, x, y))]);
    this.state.position = { x, y };
  }

  polylineTo(points: Pt[]): void {
    if (!points.length) return;
    const m = this.matrix();
    this.openFrom(m, points.map((p) => L(apply(m, p.x, p.y))));
    this.state.position = points[points.length - 1]!;
  }

  polyBezierTo(points: Pt[]): void {
    const n = points.length - (points.length % 3);
    if (!n) return;
    const m = this.matrix();
    this.openFrom(m, bezierCommands(m, points, 0, n));
    this.state.position = points[n - 1]!;
  }

  /** Polyline (open, stroked) or Polygon (closed, filled and stroked). */
  poly(points: Pt[], closed: boolean): void {
    if (points.length < 2) return;
    const m = this.matrix();
    const cmds = points.map((p, i) => (i ? L(apply(m, p.x, p.y)) : M(apply(m, p.x, p.y))));
    if (closed) cmds.push({ op: 'Z' });
    this.shape(cmds, closed, true);
  }

  /** PolyPolygon draws one filled shape (holes follow the fill mode); PolyPolyline separate strokes. */
  polyPoly(polys: Pt[][], closed: boolean): void {
    const m = this.matrix();
    const cmds: PathCommand[] = [];
    for (const points of polys) {
      if (points.length < 2) continue;
      points.forEach((p, i) => cmds.push(i ? L(apply(m, p.x, p.y)) : M(apply(m, p.x, p.y))));
      if (closed) cmds.push({ op: 'Z' });
    }
    if (cmds.length) this.shape(cmds, closed, true);
  }

  polyBezier(points: Pt[]): void {
    if (points.length < 4) return;
    const m = this.matrix();
    const n = 1 + Math.floor((points.length - 1) / 3) * 3;
    const cmds = [M(apply(m, points[0]!.x, points[0]!.y)), ...bezierCommands(m, points, 1, n)];
    this.shape(cmds, false, true);
  }

  /** PolyDraw: PT_MOVETO 6, PT_LINETO 2, PT_BEZIERTO 4, each optionally | PT_CLOSEFIGURE 1. */
  polyDraw(points: Pt[], types: number[]): void {
    const m = this.matrix();
    const cmds: PathCommand[] = [];
    let last: Pt | undefined;
    for (let i = 0; i < points.length; i++) {
      const p = points[i]!;
      const type = (types[i] ?? 2) & ~1;
      if (type === 4 && i + 2 < points.length) {
        cmds.push(...bezierCommands(m, points, i, i + 3));
        i += 2;
      } else {
        cmds.push(type === 6 ? M(apply(m, p.x, p.y)) : L(apply(m, p.x, p.y)));
      }
      last = points[i]!;
      if ((types[i] ?? 0) & 1) cmds.push({ op: 'Z' });
    }
    if (!last) return;
    if (cmds[0]?.op === 'M') {
      this.shape(cmds, false, true);
    } else {
      this.openFrom(m, cmds);
    }
    this.state.position = last;
  }

  // ---- shapes ----

  rectangle(l: number, t: number, r: number, b: number): void {
    const m = this.matrix();
    this.shape([M(apply(m, l, t)), L(apply(m, r, t)), L(apply(m, r, b)), L(apply(m, l, b)), { op: 'Z' }], true, true);
  }

  roundRect(l: number, t: number, r: number, b: number, w: number, h: number): void {
    const left = Math.min(l, r);
    const right = Math.max(l, r);
    const top = Math.min(t, b);
    const bottom = Math.max(t, b);
    const rx = Math.min(Math.abs(w) / 2, (right - left) / 2);
    const ry = Math.min(Math.abs(h) / 2, (bottom - top) / 2);
    if (rx <= 0 || ry <= 0) {
      this.rectangle(left, top, right, bottom);
      return;
    }
    const m = this.matrix();
    const cmds: PathCommand[] = [M(apply(m, left + rx, top)), L(apply(m, right - rx, top))];
    const corner = (cx: number, cy: number, t0: number): void => {
      cmds.push(...cubicCommands(m, arcCubics(cx, cy, rx, ry, t0, Math.PI / 2)));
    };
    corner(right - rx, top + ry, -Math.PI / 2);
    cmds.push(L(apply(m, right, bottom - ry)));
    corner(right - rx, bottom - ry, 0);
    cmds.push(L(apply(m, left + rx, bottom)));
    corner(left + rx, bottom - ry, Math.PI / 2);
    cmds.push(L(apply(m, left, top + ry)));
    corner(left + rx, top + ry, Math.PI);
    cmds.push({ op: 'Z' });
    this.shape(cmds, true, true);
  }

  ellipse(l: number, t: number, r: number, b: number): void {
    const m = this.matrix();
    const e = ellipseCubics(l, t, r, b);
    this.shape([M(apply(m, e.start.x, e.start.y)), ...cubicCommands(m, e.cubics), { op: 'Z' }], true, true);
  }

  arc(kind: ArcKind, l: number, t: number, r: number, b: number, xs: number, ys: number, xe: number, ye: number): void {
    const m = this.matrix();
    const arc = gdiArc(l, t, r, b, xs, ys, xe, ye, this.clockwise(m));
    const start = apply(m, arc.start.x, arc.start.y);
    const curve = cubicCommands(m, arc.cubics);
    switch (kind) {
      case 'arc':
        this.shape([M(start), ...curve], false, true);
        break;
      case 'arcto':
        this.openFrom(m, [L(start), ...curve]);
        this.state.position = arc.end;
        break;
      case 'chord':
        this.shape([M(start), ...curve, { op: 'Z' }], true, true);
        break;
      case 'pie':
        this.shape([M(apply(m, arc.centre.x, arc.centre.y)), L(start), ...curve, { op: 'Z' }], true, true);
        break;
    }
  }

  /** AngleArc: a line from the current position to the arc's start, then the arc (angles in degrees, counterclockwise). */
  angleArc(cx: number, cy: number, radius: number, startDeg: number, sweepDeg: number): void {
    const m = this.matrix();
    // The angles turn counterclockwise on a y-down device; a mirrored compatible-mode mapping reverses that in logical space.
    const sign = this.mirrored(m) ? 1 : -1;
    const sweep = Math.max(-360, Math.min(360, sweepDeg));
    const t0 = (sign * startDeg * Math.PI) / 180;
    const cubics = arcCubics(cx, cy, radius, radius, t0, (sign * sweep * Math.PI) / 180);
    const start = { x: cx + radius * Math.cos(t0), y: cy + radius * Math.sin(t0) };
    this.openFrom(m, [L(apply(m, start.x, start.y)), ...cubicCommands(m, cubics)]);
    const end = cubics.length ? cubics[cubics.length - 1]![2] : start;
    this.state.position = end;
  }

  /** Fill a rectangle with a colour, or the current brush (PatBlt, bitmap-less BitBlt). */
  fillRect(l: number, t: number, r: number, b: number, color: string | undefined): void {
    const m = this.matrix();
    const commands: PathCommand[] = [M(apply(m, l, t)), L(apply(m, r, t)), L(apply(m, r, b)), L(apply(m, l, b)), { op: 'Z' }];
    if (color) {
      this.push({ type: 'path', commands, fill: color });
      return;
    }
    const fill = this.fillStyle();
    if (fill) this.push({ type: 'path', commands, ...fill });
  }

  // ---- bitmaps ----

  /** Draw a bitmap into a destination rectangle given in logical units. */
  image(raster: RasterImage, x: number, y: number, w: number, h: number): void {
    const m = this.matrix();
    const p0 = apply(m, x, y);
    const p1 = apply(m, x + w, y + h);
    const width = Math.abs(p1.x - p0.x);
    const height = Math.abs(p1.y - p0.y);
    if (!(width > 0 && height > 0) || !Number.isFinite(width + height)) return;
    // A destination that runs backwards in page space mirrors the bitmap, as StretchBlt does.
    this.push({
      type: 'image',
      x: Math.min(p0.x, p1.x),
      y: Math.min(p0.y, p1.y),
      w: width,
      h: height,
      target: undefined,
      external: false,
      raster: flipRaster(raster, p1.x < p0.x, p1.y < p0.y),
    });
  }

  // ---- text ----

  /**
   * ExtTextOut / TextOut. (x, y) is the reference point (ignored under TA_UPDATECP); `dx` the
   * advance after each character in logical units; `opaque` a rectangle filled with the
   * background colour first (ETO_OPAQUE), which pictures also use to paint plain rectangles.
   */
  text(x: number, y: number, chars: string[], dx: number[] | undefined, opaque?: [number, number, number, number]): void {
    const st = this.state;
    const m = this.matrix();
    if (opaque) {
      const [l, t, r, b] = opaque;
      if (r !== l && b !== t) this.fillRect(l, t, r, b, st.bkColor);
    }
    if (!chars.length) return;
    const { fonts } = this.options;
    const f = st.font;
    const base = fonts.get(f.faceName || undefined, f.weight >= 600, f.italic);
    const xscale = Math.hypot(m.a, m.b);
    const yscale = Math.hypot(m.c, m.d);
    const cell = base.ascent + base.descent || 1;
    // Negative heights give the em size; positive ones the cell height (ascent + descent).
    const size = f.height === 0 ? this.options.defaultFontSize : f.height < 0 ? -f.height * yscale : (f.height * yscale) / cell;
    if (!(size > 0) || !Number.isFinite(size)) return;

    const cps = chars.map((c) => c.codePointAt(0) ?? 0x20);
    const charFonts = cps.map((cp) => fonts.forChar(base, cp));
    const advances = cps.map((cp, i) => (dx ? (dx[i] ?? 0) * xscale : charFonts[i]!.advance(cp) * size + st.charExtra * xscale));
    const total = advances.reduce((a, b) => a + b, 0);

    const updateCP = (st.textAlign & TA_UPDATECP) !== 0;
    const ref = updateCP ? st.position : { x, y };
    const origin = apply(m, ref.x, ref.y);
    const horizontal = st.textAlign & TA_CENTER;
    const vertical = st.textAlign & TA_BASELINE;
    // GDI draws text upright on the device whatever the mapping, so alignment works in page points.
    const left = origin.x - (horizontal === TA_CENTER ? total / 2 : horizontal === TA_RIGHT ? total : 0);
    const baseline = origin.y + (vertical === TA_BASELINE ? 0 : vertical === TA_BOTTOM ? -base.descent * size : base.ascent * size);

    const xs: number[] = [];
    let cx = left;
    for (const a of advances) {
      xs.push(cx);
      cx += a;
    }
    let runStart = -1;
    let runFont: LayoutFont | undefined;
    const flush = (end: number): void => {
      if (runStart < 0 || !runFont) return;
      const slice = (arr: readonly number[]): number[] => arr.slice(runStart, end);
      this.push(textItem(xs[runStart]!, baseline, chars.slice(runStart, end), slice(xs), slice(advances), runFont, size, st.textColor));
      runStart = -1;
      runFont = undefined;
    };
    for (let i = 0; i < chars.length; i++) {
      // Control characters advance but draw nothing (fonts would show a missing-glyph box).
      if (cps[i]! < 0x20) {
        flush(i);
        continue;
      }
      if (runFont !== charFonts[i]) {
        flush(i);
        runStart = i;
        runFont = charFonts[i];
      }
    }
    flush(chars.length);

    if (f.underline || f.strikeOut) {
      const thickness = Math.max(base.underlineThickness * size, HAIRLINE);
      const line = (yy: number): void => {
        this.push({ type: 'line', x1: left, y1: yy, x2: left + total, y2: yy, width: thickness, color: st.textColor, style: 'solid' });
      };
      if (f.underline) line(baseline + base.underlinePosition * size);
      if (f.strikeOut) line(baseline - base.ascent * size * 0.3);
    }
    if (updateCP && xscale > 0) {
      const dir = horizontal === TA_CENTER ? 0 : horizontal === TA_RIGHT ? -1 : 1;
      st.position = { x: ref.x + (dir * total * Math.sign(m.a || 1)) / xscale, y: ref.y };
    }
  }

  // ---- internals ----

  private push(item: DisplayItem): void {
    this.items.push(item);
    this.chain = undefined;
  }

  /** A mirrored mapping in compatible mode: GDI then keeps arc directions in device space. */
  private mirrored(m: Matrix): boolean {
    return !this.state.advanced && determinant(m) < 0;
  }

  /** Arc direction in logical space, as "clockwise when y points down". */
  private clockwise(m: Matrix): boolean {
    const cw = this.state.arcDirection === AD_CLOCKWISE;
    return this.mirrored(m) ? !cw : cw;
  }

  /** A closed or self-contained shape: recorded into the path bracket, or drawn. */
  private shape(cmds: PathCommand[], fill: boolean, stroke: boolean): void {
    if (this.recording && this.path) {
      this.path.push(...cmds);
      // These calls neither use nor move the current position, so a later LineTo starts afresh.
      this.figureOpen = false;
      return;
    }
    this.output(cmds, fill, stroke);
  }

  /** Commands that continue from the current position (LineTo, PolylineTo, ArcTo, …): stroked only. */
  private openFrom(m: Matrix, cmds: PathCommand[]): void {
    const from = apply(m, this.state.position.x, this.state.position.y);
    if (this.recording && this.path) {
      if (!this.figureOpen) this.path.push(M(from));
      this.path.push(...cmds);
      this.figureOpen = cmds[cmds.length - 1]?.op !== 'Z';
      return;
    }
    const stroke = this.strokeStyle(m);
    if (!stroke) return;
    const end = lastPoint(cmds) ?? from;
    // Runs of LineTo calls (common in charts) become one path rather than thousands of items.
    const chain = this.chain;
    if (chain && this.items[this.items.length - 1] === chain.item && near(chain.end, from) && sameStroke(chain.item, stroke)) {
      chain.item.commands.push(...cmds);
      chain.end = end;
      return;
    }
    const item: PathItem = { type: 'path', commands: [M(from), ...cmds], ...stroke };
    this.push(item);
    this.chain = { item, end };
  }

  private output(commands: PathCommand[], fill: boolean, stroke: boolean): void {
    const m = this.matrix();
    const f = fill ? this.fillStyle() : undefined;
    const s = stroke ? this.strokeStyle(m) : undefined;
    if (!f && !s) return;
    this.push({ type: 'path', commands, ...f, ...s });
  }

  private fillStyle(): FillStyle | undefined {
    const st = this.state;
    const brush = st.brush;
    let style: FillStyle;
    switch (brush.style) {
      case 'null':
        return undefined;
      case 'hatched':
        // A hatch drawn as a flat tint of roughly its coverage; opaque mode fills the gaps with the background.
        style = st.bkMode === OPAQUE ? { fill: mixColors(brush.color, st.bkColor, HATCH_COVERAGE) } : { fill: brush.color, fillOpacity: HATCH_COVERAGE };
        break;
      default:
        style = { fill: brush.color };
    }
    if (st.polyFillMode === ALTERNATE) style.evenOdd = true;
    return style;
  }

  private strokeStyle(m: Matrix): StrokeStyle | undefined {
    const pen = this.state.pen;
    if (pen.style === PS_NULL) return undefined;
    const hair = pen.cosmetic || pen.width === 0;
    const scale = Math.sqrt(Math.abs(determinant(m)));
    const width = hair ? HAIRLINE : Math.max(HAIRLINE, pen.width * scale);
    const style: StrokeStyle = { stroke: pen.color, strokeWidth: width, lineCap: pen.cap, lineJoin: pen.join };
    // Wine's patterns: cosmetic dashes in device pixels, geometric ones in multiples of the width.
    const unit = hair ? Math.max(this.options.pixel, HAIRLINE) : width;
    let pattern: number[] | undefined;
    switch (pen.style) {
      case PS_DASH:
        pattern = hair ? [18, 6] : [3, 1];
        break;
      case PS_DOT:
        pattern = hair ? [3, 3] : [1, 1];
        break;
      case PS_DASHDOT:
        pattern = hair ? [9, 6, 3, 6] : [3, 1, 1, 1];
        break;
      case PS_DASHDOTDOT:
        pattern = hair ? [9, 3, 3, 3, 3, 3] : [3, 1, 1, 1, 1, 1];
        break;
      case PS_ALTERNATE:
        pattern = [1, 1];
        break;
      case PS_USERSTYLE:
        if (pen.dashes?.length) style.dash = pen.dashes.map((d) => Math.abs(d) * (hair ? unit : scale));
        break;
    }
    if (pattern) style.dash = pattern.map((d) => d * unit);
    if (style.dash && !style.dash.some((d) => d > 0)) delete style.dash;
    return style;
  }
}

function near(a: Pt, b: Pt): boolean {
  return Math.abs(a.x - b.x) < 1e-6 && Math.abs(a.y - b.y) < 1e-6;
}

function sameStroke(item: PathItem, s: StrokeStyle): boolean {
  return (
    !item.fill &&
    item.stroke === s.stroke &&
    item.strokeWidth === s.strokeWidth &&
    item.lineCap === s.lineCap &&
    item.lineJoin === s.lineJoin &&
    (item.dash ?? []).join() === (s.dash ?? []).join()
  );
}

function lastPoint(cmds: PathCommand[]): Pt | undefined {
  for (let i = cmds.length - 1; i >= 0; i--) {
    const c = cmds[i]!;
    if (c.op !== 'Z') return { x: c.x, y: c.y };
  }
  return undefined;
}

function cubicCommands(m: Matrix, cubics: Cubic[]): PathCommand[] {
  return cubics.map(([c1, c2, p]) => {
    const a = apply(m, c1.x, c1.y);
    const b = apply(m, c2.x, c2.y);
    const e = apply(m, p.x, p.y);
    return { op: 'C', x1: a.x, y1: a.y, x2: b.x, y2: b.y, x: e.x, y: e.y };
  });
}

/** Bézier triples (control, control, end) of points[from, to). */
function bezierCommands(m: Matrix, points: Pt[], from: number, to: number): PathCommand[] {
  const cubics: Cubic[] = [];
  for (let i = from; i + 2 < to && i + 2 < points.length; i += 3) cubics.push([points[i]!, points[i + 1]!, points[i + 2]!]);
  return cubicCommands(m, cubics);
}

/** Close every figure that is not closed yet (FillPath and StrokeAndFillPath do this). */
function closeFigures(path: PathCommand[]): PathCommand[] {
  const out: PathCommand[] = [];
  let open = false;
  for (const c of path) {
    if (c.op === 'M' && open) out.push({ op: 'Z' });
    out.push(c);
    // A figure needs a segment to be worth closing; a bare MoveTo stays as it is.
    open = c.op === 'L' || c.op === 'C';
  }
  if (open) out.push({ op: 'Z' });
  return out;
}
