/**
 * Draws a chart model into display items inside a frame (points, y down): chart and plot area,
 * title, legend, axes with automatic "nice" scales, gridlines, bars, lines, areas, scatter and bubble
 * points, pies and doughnuts, radar, and data labels. Layout follows Office's order: the title and
 * legend take their space first, axis labels next, and the plot gets what is left (or the manual
 * layout written in the file). Fonts keep their point sizes whatever the frame size, as in Office.
 */
import { measureChars, textItem } from '../layout/paint.js';
import type { DisplayItem, FontProvider, LayoutFont, PathCommand, PathItem } from '../layout/types.js';
import type { XmlElement } from '../xml/dom.js';
import { dashPattern, type Paint } from './color.js';
import { formatNumber, isDateCode } from './format.js';
import { readChart, type ChartAxis, type ChartModel, type ChartPlot, type ChartSeries, type DataLabelFlags, type LineStyle, type ManualLayout, type MarkerStyle, type TextStyle } from './model.js';
import { hasOutlines, textOutline } from './outline.js';

export interface ChartFrame {
  x: number;
  y: number;
  w: number;
  h: number;
}

export interface RenderChartOptions {
  fonts: FontProvider;
  /** Theme colours by scheme name (accent1, tx1, bg1, …) → RRGGBB. */
  schemeColor(name: string): string | undefined;
  /** Font for chart text that names none (default Calibri, the Office theme's body font). */
  defaultFont?: string;
}

/** Read a c:chartSpace and draw it into `frame`. */
export function renderChart(chartSpace: XmlElement, frame: ChartFrame, options: RenderChartOptions): DisplayItem[] {
  const model = readChart(chartSpace, { schemeColor: (name) => options.schemeColor(name) });
  return drawChart(model, frame, options);
}

type Box = ChartFrame;

const TICK = 3.5;
const LABEL_GAP = 3;

// ---------------------------------------------------------------------------------------------
// Drawing primitives

class Draw {
  readonly items: DisplayItem[] = [];

  constructor(
    private readonly fonts: FontProvider,
    private readonly defaultFont: string,
  ) {}

  font(st: TextStyle): LayoutFont {
    return this.fonts.get(st.font ?? this.defaultFont, st.bold, st.italic);
  }

  lineHeight(st: TextStyle): number {
    const f = this.font(st);
    return (f.ascent + f.descent) * st.size;
  }

  ascent(st: TextStyle): number {
    return this.font(st).ascent * st.size;
  }

  /** Split text into runs that one font can draw (fallback fonts for missing glyphs). */
  private runs(text: string, st: TextStyle): Array<{ font: LayoutFont; chars: string[] }> {
    const base = this.font(st);
    const out: Array<{ font: LayoutFont; chars: string[] }> = [];
    for (const ch of text) {
      const f = this.fonts.forChar(base, ch.codePointAt(0)!);
      const last = out[out.length - 1];
      if (last && last.font === f) last.chars.push(ch);
      else out.push({ font: f, chars: [ch] });
    }
    return out;
  }

  width(text: string, st: TextStyle): number {
    let w = 0;
    for (const run of this.runs(text, st)) for (const ch of run.chars) w += run.font.advance(ch.codePointAt(0)!) * st.size;
    return w;
  }

  /** One line of text with its left end at x and its baseline at y. */
  textLine(text: string, x: number, baseline: number, st: TextStyle): void {
    let cx = x;
    for (const run of this.runs(text, st)) {
      const m = measureChars(run.chars.join(''), run.font, st.size, cx);
      this.items.push(textItem(cx, baseline, m.chars, m.xs, m.widths, run.font, st.size, st.color));
      cx += m.width;
    }
  }

  /**
   * Lines of text around an anchor: `h` says which part of the block sits at x (left end, centre,
   * right end), `v` which part sits at y.
   */
  textBlock(lines: string[], x: number, y: number, st: TextStyle, h: 'left' | 'center' | 'right', v: 'top' | 'middle' | 'bottom'): void {
    const lh = this.lineHeight(st);
    const total = lh * lines.length;
    let top = v === 'top' ? y : v === 'middle' ? y - total / 2 : y - total;
    const asc = this.ascent(st);
    for (const line of lines) {
      const w = this.width(line, st);
      const left = h === 'left' ? x : h === 'center' ? x - w / 2 : x - w;
      if (line) this.textLine(line, left, top + asc, st);
      top += lh;
    }
  }

  /** Can text in this style be drawn rotated (as glyph outlines)? */
  canRotate(st: TextStyle): boolean {
    return hasOutlines(this.font(st));
  }

  /**
   * Rotated text as filled outlines. The point (x, y) is where the text's anchor sits: `along` 0, 0.5
   * or 1 picks the start, middle or end of the line, vertically the line's middle. Degrees clockwise.
   */
  rotatedText(text: string, x: number, y: number, st: TextStyle, degrees: number, along: number): boolean {
    const font = this.font(st);
    if (!hasOutlines(font)) return false;
    const w = this.width(text, st);
    const asc = font.ascent * st.size;
    const desc = font.descent * st.size;
    const rad = (degrees * Math.PI) / 180;
    const cos = Math.cos(rad);
    const sin = Math.sin(rad);
    let pen = 0;
    const commands: PathCommand[] = [];
    for (const run of this.runs(text, st)) {
      const start = pen;
      const cmds = textOutline(run.chars, run.font, st.size, (u, v) => {
        const lu = start + u - along * w;
        const lv = v + (asc - desc) / 2;
        return { x: x + lu * cos - lv * sin, y: y + lu * sin + lv * cos };
      });
      if (!cmds) return false;
      commands.push(...cmds);
      for (const ch of run.chars) pen += run.font.advance(ch.codePointAt(0)!) * st.size;
    }
    const item: PathItem = { type: 'path', commands, fill: st.color };
    if (font.syntheticBold) {
      item.stroke = st.color;
      item.strokeWidth = st.size * 0.03;
    }
    this.items.push(item);
    return true;
  }

  /** Greedy word wrap to `maxWidth`; words longer than the width stay whole. Honours "\n". */
  wrap(text: string, maxWidth: number, st: TextStyle): string[] {
    const out: string[] = [];
    for (const para of text.split('\n')) {
      const words = para.split(/ +/);
      let line = '';
      for (const word of words) {
        const candidate = line ? `${line} ${word}` : word;
        if (line && this.width(candidate, st) > maxWidth) {
          out.push(line);
          line = word;
        } else line = candidate;
      }
      out.push(line);
    }
    return out;
  }

  path(commands: PathCommand[], fill: Paint | null | undefined, line: LineStyle | null | undefined, extra: Partial<PathItem> = {}): void {
    if (commands.length === 0) return;
    const item: PathItem = { type: 'path', commands, ...extra };
    if (fill) {
      item.fill = fill.color;
      if (fill.opacity < 1) item.fillOpacity = fill.opacity;
    }
    if (line && line.paint && line.width > 0) {
      item.stroke = line.paint.color;
      if (line.paint.opacity < 1) item.strokeOpacity = line.paint.opacity;
      item.strokeWidth = line.width;
      const dash = dashPattern(line.dash, line.width);
      if (dash) item.dash = dash;
      if (line.cap) item.lineCap = line.cap;
      if (line.join) item.lineJoin = line.join;
    }
    if (item.fill === undefined && item.stroke === undefined) return;
    this.items.push(item);
  }

  segment(x1: number, y1: number, x2: number, y2: number, line: LineStyle | undefined): void {
    if (!line?.paint) return;
    this.path(
      [
        { op: 'M', x: x1, y: y1 },
        { op: 'L', x: x2, y: y2 },
      ],
      null,
      line,
    );
  }
}

function rectPath(b: Box, radius = 0): PathCommand[] {
  if (radius <= 0) {
    return [{ op: 'M', x: b.x, y: b.y }, { op: 'L', x: b.x + b.w, y: b.y }, { op: 'L', x: b.x + b.w, y: b.y + b.h }, { op: 'L', x: b.x, y: b.y + b.h }, { op: 'Z' }];
  }
  const r = Math.min(radius, b.w / 2, b.h / 2);
  const k = r * 0.5523;
  const { x, y, w, h } = b;
  return [
    { op: 'M', x: x + r, y },
    { op: 'L', x: x + w - r, y },
    { op: 'C', x1: x + w - r + k, y1: y, x2: x + w, y2: y + r - k, x: x + w, y: y + r },
    { op: 'L', x: x + w, y: y + h - r },
    { op: 'C', x1: x + w, y1: y + h - r + k, x2: x + w - r + k, y2: y + h, x: x + w - r, y: y + h },
    { op: 'L', x: x + r, y: y + h },
    { op: 'C', x1: x + r - k, y1: y + h, x2: x, y2: y + h - r + k, x, y: y + h - r },
    { op: 'L', x, y: y + r },
    { op: 'C', x1: x, y1: y + r - k, x2: x + r - k, y2: y, x: x + r, y },
    { op: 'Z' },
  ];
}

/** Point on a circle at `deg` clockwise from 12 o'clock. */
function polar(cx: number, cy: number, r: number, deg: number): { x: number; y: number } {
  const a = (deg * Math.PI) / 180;
  return { x: cx + r * Math.sin(a), y: cy - r * Math.cos(a) };
}

/** Arc from a1 to a2 (degrees clockwise from 12 o'clock; a2 < a1 runs counter-clockwise) as Béziers. */
function arcTo(cmds: PathCommand[], cx: number, cy: number, r: number, a1: number, a2: number): void {
  const sweep = a2 - a1;
  const n = Math.max(1, Math.ceil(Math.abs(sweep) / 90 - 1e-9));
  const step = sweep / n;
  for (let i = 0; i < n; i++) {
    const s = a1 + i * step;
    const e = s + step;
    const k = (4 / 3) * Math.tan(((step * Math.PI) / 180) / 4) * r;
    const sr = (s * Math.PI) / 180;
    const er = (e * Math.PI) / 180;
    const p1 = polar(cx, cy, r, s);
    const p2 = polar(cx, cy, r, e);
    // Tangent of (sin a, -cos a) is (cos a, sin a).
    cmds.push({ op: 'C', x1: p1.x + k * Math.cos(sr), y1: p1.y + k * Math.sin(sr), x2: p2.x - k * Math.cos(er), y2: p2.y - k * Math.sin(er), x: p2.x, y: p2.y });
  }
}

function circlePath(cx: number, cy: number, r: number): PathCommand[] {
  const p = polar(cx, cy, r, 0);
  const cmds: PathCommand[] = [{ op: 'M', x: p.x, y: p.y }];
  arcTo(cmds, cx, cy, r, 0, 360);
  cmds.push({ op: 'Z' });
  return cmds;
}

/** Smooth curve through points (Catmull-Rom as cubic Béziers), as Office's "smoothed line". */
function smoothPath(pts: Array<{ x: number; y: number }>): PathCommand[] {
  const cmds: PathCommand[] = [];
  if (pts.length === 0) return cmds;
  cmds.push({ op: 'M', x: pts[0]!.x, y: pts[0]!.y });
  for (let i = 0; i < pts.length - 1; i++) {
    const p0 = pts[i - 1] ?? pts[i]!;
    const p1 = pts[i]!;
    const p2 = pts[i + 1]!;
    const p3 = pts[i + 2] ?? p2;
    cmds.push({ op: 'C', x1: p1.x + (p2.x - p0.x) / 6, y1: p1.y + (p2.y - p0.y) / 6, x2: p2.x - (p3.x - p1.x) / 6, y2: p2.y - (p3.y - p1.y) / 6, x: p2.x, y: p2.y });
  }
  return cmds;
}

function polyline(pts: Array<{ x: number; y: number }>): PathCommand[] {
  return pts.map((p, i) => ({ op: i === 0 ? 'M' : 'L', x: p.x, y: p.y }) as PathCommand);
}

function drawMarker(d: Draw, m: MarkerStyle, x: number, y: number): void {
  if (m.symbol === 'none') return;
  const s = m.size;
  const h = s / 2;
  const line: LineStyle | null = m.line ? { paint: m.line, width: m.lineWidth, join: 'miter' } : null;
  let cmds: PathCommand[];
  let strokeOnly = false;
  switch (m.symbol) {
    case 'circle':
      cmds = circlePath(x, y, h);
      break;
    case 'diamond':
      cmds = [{ op: 'M', x, y: y - h }, { op: 'L', x: x + h, y }, { op: 'L', x, y: y + h }, { op: 'L', x: x - h, y }, { op: 'Z' }];
      break;
    case 'triangle':
      cmds = [{ op: 'M', x, y: y - h }, { op: 'L', x: x + h, y: y + h }, { op: 'L', x: x - h, y: y + h }, { op: 'Z' }];
      break;
    case 'x':
      cmds = [{ op: 'M', x: x - h, y: y - h }, { op: 'L', x: x + h, y: y + h }, { op: 'M', x: x + h, y: y - h }, { op: 'L', x: x - h, y: y + h }];
      strokeOnly = true;
      break;
    case 'plus':
      cmds = [{ op: 'M', x, y: y - h }, { op: 'L', x, y: y + h }, { op: 'M', x: x - h, y }, { op: 'L', x: x + h, y }];
      strokeOnly = true;
      break;
    case 'star':
      cmds = [
        { op: 'M', x: x - h, y: y - h },
        { op: 'L', x: x + h, y: y + h },
        { op: 'M', x: x + h, y: y - h },
        { op: 'L', x: x - h, y: y + h },
        { op: 'M', x, y: y - h },
        { op: 'L', x, y: y + h },
      ];
      strokeOnly = true;
      break;
    case 'dash':
      cmds = rectPath({ x: x - h, y: y - s / 8, w: s, h: s / 4 });
      break;
    case 'dot':
      cmds = rectPath({ x: x - s / 4, y: y - s / 8, w: s / 2, h: s / 4 });
      break;
    default:
      cmds = rectPath({ x: x - h, y: y - h, w: s, h: s });
  }
  if (strokeOnly) d.path(cmds, null, { paint: m.line ?? m.fill, width: Math.max(m.lineWidth, 0.75) });
  else d.path(cmds, m.fill, line);
}

// ---------------------------------------------------------------------------------------------
// Scales

interface Scale {
  min: number;
  max: number;
  major: number;
  minor: number;
  log?: number;
}

function roundNice(v: number): number {
  return Number(v.toPrecision(12));
}

/** Tick values of a scale. */
function ticks(s: Scale): number[] {
  const out: number[] = [];
  if (s.log) {
    for (let v = s.min, i = 0; v <= s.max * (1 + 1e-9) && i < 100; v *= s.log, i++) out.push(roundNice(v));
    return out;
  }
  const n = Math.round((s.max - s.min) / s.major);
  for (let i = 0; i <= Math.min(n, 1000); i++) out.push(roundNice(s.min + i * s.major));
  return out;
}

/**
 * Office-like automatic axis scale: zero-based unless the data sits in a narrow band far from zero,
 * about 5% headroom beyond the data, and the smallest 1-2-5 major unit whose ticks are not crowded
 * (`minSpacing` says how far apart ticks must be, in points, for a candidate).
 */
function niceScale(lo: number, hi: number, axis: ChartAxis, length: number, percent: boolean, minSpacing: (s: Scale) => number): Scale {
  if (axis.logBase) {
    const b = axis.logBase;
    const pos = (v: number): number => (v > 0 ? v : 1);
    const min = axis.min && axis.min > 0 ? axis.min : b ** Math.floor(Math.log(pos(Number.isFinite(lo) ? lo : 1)) / Math.log(b) + 1e-9);
    let max = axis.max && axis.max > min ? axis.max : b ** Math.ceil(Math.log(pos(Number.isFinite(hi) ? hi : b)) / Math.log(b) - 1e-9);
    if (max <= min) max = min * b;
    return { min, max, major: b, minor: b, log: b };
  }
  if (!Number.isFinite(lo) || !Number.isFinite(hi)) {
    lo = 0;
    hi = percent ? 1 : 1;
  }
  if (percent) {
    hi = hi > 0 ? 1 : 0;
    lo = lo < 0 ? -1 : 0;
  }
  if (lo === hi) {
    if (lo === 0) hi = 1;
    else if (lo > 0) lo = 0;
    else hi = 0;
  }
  const fixedMin = axis.min;
  const fixedMax = axis.max !== undefined && (fixedMin === undefined || axis.max > fixedMin) ? axis.max : undefined;
  const headroom = percent ? 0 : 0.05;

  const build = (unit: number): Scale => {
    let min: number;
    if (fixedMin !== undefined) min = fixedMin;
    else if (lo >= 0) {
      // A narrow band of positive values gets a non-zero minimum, as Office does.
      min = hi > 0 && (hi - lo) / hi < 1 / 6 ? Math.max(0, Math.floor((lo - (hi - lo) / 2) / unit) * unit) : 0;
    } else min = Math.floor((lo - headroom * (Math.max(hi, 0) - lo)) / unit + 1e-9) * unit;
    let max: number;
    if (fixedMax !== undefined) max = fixedMax;
    else if (hi <= 0) max = lo < 0 && (hi - lo) / -lo < 1 / 6 ? Math.min(0, Math.ceil((hi + (hi - lo) / 2) / unit) * unit) : 0;
    else max = Math.ceil((hi + headroom * (hi - min)) / unit - 1e-9) * unit;
    if (percent) {
      if (fixedMax === undefined) max = hi > 0 ? 1 : 0;
      if (fixedMin === undefined) min = lo < 0 ? -1 : 0;
    }
    if (max <= min) max = min + unit;
    return { min: roundNice(min), max: roundNice(max), major: unit, minor: unit / 5 };
  };

  if (axis.majorUnit) {
    const s = build(axis.majorUnit);
    s.minor = axis.minorUnit ?? axis.majorUnit / 5;
    return s;
  }
  const span = Math.max((fixedMax ?? hi) - (fixedMin ?? Math.min(lo, 0)), Math.abs(hi) * 1e-9, 1e-12);
  let exp = Math.floor(Math.log10(span / 50));
  for (let guard = 0; guard < 40; guard++, exp++) {
    for (const m of [1, 2, 5]) {
      const unit = roundNice(m * 10 ** exp);
      const s = build(unit);
      const intervals = (s.max - s.min) / unit;
      if (intervals > 200) continue;
      if (intervals < 1 || length / intervals >= minSpacing(s)) {
        if (axis.minorUnit) s.minor = axis.minorUnit;
        return s;
      }
    }
  }
  return build(span);
}

// ---------------------------------------------------------------------------------------------
// Legend

interface LegendEntry {
  label: string;
  kind: 'fill' | 'line';
  fill: Paint | null;
  line?: LineStyle;
  marker?: MarkerStyle;
}

function pointFill(series: ChartSeries, i: number): Paint | null {
  const pt = series.points.find((p) => p.index === i);
  if (pt?.fill !== undefined) return pt.fill;
  const auto = series.pointColors?.[i];
  if (auto) return { color: auto, opacity: 1 };
  return series.fill;
}

function pointLine(series: ChartSeries, i: number): LineStyle {
  return series.points.find((p) => p.index === i)?.line ?? series.line;
}

function isVaried(plot: ChartPlot): boolean {
  return plot.varyColors && (plot.type === 'pie' || plot.type === 'doughnut' || (plot.series.length === 1 && (plot.type === 'bar' || plot.type === 'bubble')));
}

function legendEntries(model: ChartModel): LegendEntry[] {
  const entries: LegendEntry[] = [];
  const first = model.plots[0];
  if (model.plots.length === 1 && first && isVaried(first) && first.series[0]) {
    const s = first.series[0];
    const n = Math.max(s.categories.length, s.values.length);
    for (let i = 0; i < n; i++) entries.push({ label: s.categories[i] ?? String(i + 1), kind: 'fill', fill: pointFill(s, i), line: pointLine(s, i) });
  } else {
    for (const plot of model.plots) {
      const plotEntries: LegendEntry[] = [];
      for (const s of plot.series) {
        const lineKind = plot.type === 'line' || plot.type === 'scatter' || (plot.type === 'radar' && plot.radarStyle !== 'filled');
        if (lineKind) {
          const e: LegendEntry = { label: s.name, kind: 'line', fill: null, line: s.line };
          if (s.marker.symbol !== 'none') e.marker = s.marker;
          plotEntries.push(e);
        } else plotEntries.push({ label: s.name, kind: 'fill', fill: s.fill, line: s.line });
      }
      // Stacked columns and areas list the top series first in a vertical legend, matching the stack.
      const vertical = model.legend && ['r', 'l', 'tr'].includes(model.legend.position);
      if (vertical && (plot.grouping === 'stacked' || plot.grouping === 'percentStacked') && ((plot.type === 'bar' && plot.barDir === 'col') || plot.type === 'area')) plotEntries.reverse();
      entries.push(...plotEntries);
    }
  }
  const deleted = new Set(model.legend?.deletedEntries ?? []);
  return entries.filter((_, i) => !deleted.has(i));
}

interface LegendPlan {
  box: Box;
  rows: Array<Array<{ entry: LegendEntry; x: number; y: number; lines: string[]; keyW: number }>>;
  remaining: Box;
}

function planLegend(d: Draw, model: ChartModel, frame: Box, area: Box): LegendPlan | undefined {
  const legend = model.legend!;
  const entries = legendEntries(model);
  if (entries.length === 0) return undefined;
  const st = legend.style;
  const lh = d.lineHeight(st);
  const pad = st.size * 0.35;
  const keyGap = st.size * 0.4;
  const keyW = (e: LegendEntry): number => (e.kind === 'line' ? st.size * 2.2 : st.size * 0.75);
  const vertical = legend.position === 'r' || legend.position === 'l' || legend.position === 'tr';
  const manual = legend.layout;
  const maxW = manual?.w !== undefined && manual.wMode !== 'edge' ? manual.w * frame.w : vertical ? area.w * 0.35 : area.w;

  const rows: LegendPlan['rows'] = [];
  let boxW: number;
  let boxH: number;
  if (vertical) {
    let y = pad;
    let widest = 0;
    const col: LegendPlan['rows'][number] = [];
    for (const e of entries) {
      const kw = keyW(e);
      const lines = d.wrap(e.label, Math.max(maxW - kw - keyGap - 2 * pad, st.size * 3), st);
      const h = lines.length * lh;
      if (y + h > area.h && col.length > 0) break; // entries that do not fit are left out
      col.push({ entry: e, x: pad, y, lines, keyW: kw });
      widest = Math.max(widest, kw + keyGap + Math.max(...lines.map((l) => d.width(l, st))));
      y += h;
    }
    rows.push(col);
    boxW = widest + 2 * pad;
    boxH = y + pad;
  } else {
    const spacing = st.size * 1.1;
    let x = pad;
    let y = pad;
    let row: LegendPlan['rows'][number] = [];
    let widest = 0;
    for (const e of entries) {
      const kw = keyW(e);
      const w = kw + keyGap + d.width(e.label, st);
      if (row.length > 0 && x + w + pad > maxW) {
        rows.push(row);
        widest = Math.max(widest, x - spacing);
        row = [];
        x = pad;
        y += lh;
      }
      row.push({ entry: e, x, y, lines: [e.label], keyW: kw });
      x += w + spacing;
    }
    rows.push(row);
    widest = Math.max(widest, x - spacing);
    boxW = Math.min(widest + pad, area.w);
    boxH = y + lh + pad;
    // Centre each row within the legend box.
    for (const r of rows) {
      const last = r[r.length - 1];
      if (!last) continue;
      const rowW = last.x + last.keyW + keyGap + d.width(last.lines[0] ?? '', st) + pad;
      const shift = (boxW - rowW) / 2;
      for (const item of r) item.x += Math.max(0, shift);
    }
  }

  let box: Box;
  const remaining = { ...area };
  const gap = Math.max(4, st.size * 0.5);
  if (manual && (manual.x !== undefined || manual.y !== undefined)) {
    const x = manual.x !== undefined ? (manual.xMode === 'edge' ? frame.x + manual.x * frame.w : area.x + manual.x * frame.w) : area.x + area.w - boxW;
    const y = manual.y !== undefined ? (manual.yMode === 'edge' ? frame.y + manual.y * frame.h : area.y + manual.y * frame.h) : area.y;
    const w = manual.w !== undefined ? (manual.wMode === 'edge' ? frame.x + manual.w * frame.w - x : manual.w * frame.w) : boxW;
    const h = manual.h !== undefined ? (manual.hMode === 'edge' ? frame.y + manual.h * frame.h - y : manual.h * frame.h) : boxH;
    box = { x, y, w, h };
    return { box, rows, remaining };
  }
  switch (legend.position) {
    case 'l':
      box = { x: area.x, y: area.y + (area.h - boxH) / 2, w: boxW, h: boxH };
      if (!legend.overlay) {
        remaining.x += boxW + gap;
        remaining.w -= boxW + gap;
      }
      break;
    case 't':
      box = { x: area.x + (area.w - boxW) / 2, y: area.y, w: boxW, h: boxH };
      if (!legend.overlay) {
        remaining.y += boxH + gap;
        remaining.h -= boxH + gap;
      }
      break;
    case 'b':
      box = { x: area.x + (area.w - boxW) / 2, y: area.y + area.h - boxH, w: boxW, h: boxH };
      if (!legend.overlay) remaining.h -= boxH + gap;
      break;
    case 'tr':
      box = { x: area.x + area.w - boxW, y: area.y, w: boxW, h: boxH };
      if (!legend.overlay) remaining.w -= boxW + gap;
      break;
    default:
      box = { x: area.x + area.w - boxW, y: area.y + (area.h - boxH) / 2, w: boxW, h: boxH };
      if (!legend.overlay) remaining.w -= boxW + gap;
  }
  return { box, rows, remaining };
}

function drawLegend(d: Draw, model: ChartModel, plan: LegendPlan): void {
  const legend = model.legend!;
  const st = legend.style;
  const lh = d.lineHeight(st);
  const keyGap = st.size * 0.4;
  d.path(rectPath(plan.box), legend.fill ?? null, legend.line ?? null);
  for (const row of plan.rows) {
    for (const item of row) {
      const x = plan.box.x + item.x;
      const top = plan.box.y + item.y;
      const cy = top + lh / 2;
      const e = item.entry;
      if (e.kind === 'fill') {
        const s = st.size * 0.75;
        // Swatches keep a visible edge only when the series has an outline.
        d.path(rectPath({ x: x + (item.keyW - s) / 2, y: cy - s / 2, w: s, h: s }), e.fill, e.line?.paint ? { ...e.line, width: Math.min(e.line.width, 1), dash: undefined } : null);
      } else {
        if (e.line?.paint) d.segment(x, cy, x + item.keyW, cy, { ...e.line, width: Math.min(e.line.width, 3) });
        if (e.marker) drawMarker(d, e.marker, x + item.keyW / 2, cy);
      }
      d.textBlock(item.lines, x + item.keyW + keyGap, top, st, 'left', 'top');
    }
  }
}

// ---------------------------------------------------------------------------------------------
// Titles

interface TitlePlan {
  draw(): void;
  remaining: Box;
}

function planTitle(d: Draw, model: ChartModel, frame: Box, area: Box): TitlePlan {
  const title = model.title!;
  const st = title.style;
  const lines = title.lines.flatMap((l) => d.wrap(l, area.w * 0.8, st));
  const lh = d.lineHeight(st);
  const h = lines.length * lh;
  const w = Math.max(...lines.map((l) => d.width(l, st)), 0);
  const remaining = { ...area };
  let cx = area.x + area.w / 2;
  let top = area.y;
  const layout = title.layout;
  if (layout && (layout.x !== undefined || layout.y !== undefined)) {
    // Manual position: x/y are the top-left corner of the title box.
    if (layout.x !== undefined) cx = (layout.xMode === 'edge' ? frame.x + layout.x * frame.w : cx - w / 2 + layout.x * frame.w) + w / 2;
    if (layout.y !== undefined) top = layout.yMode === 'edge' ? frame.y + layout.y * frame.h : top + layout.y * frame.h;
    if (!title.overlay && top < area.y + h) {
      const used = Math.max(0, top + h - area.y) + 4;
      remaining.y += used;
      remaining.h -= used;
    }
  } else if (!title.overlay) {
    const used = h + Math.max(4, st.size * 0.4);
    remaining.y += used;
    remaining.h -= used;
  }
  return {
    remaining,
    draw: () => {
      if (title.fill || title.line) d.path(rectPath({ x: cx - w / 2 - 3, y: top - 1, w: w + 6, h: h + 2 }), title.fill ?? null, title.line ?? null);
      d.textBlock(lines, cx, top, st, 'center', 'top');
    },
  };
}

// ---------------------------------------------------------------------------------------------
// Data labels

interface LabelJob {
  text: string;
  x: number;
  y: number;
  h: 'left' | 'center' | 'right';
  v: 'top' | 'middle' | 'bottom';
  style: TextStyle;
  fill?: Paint | null;
}

function labelFlags(series: ChartSeries, i: number): DataLabelFlags | undefined {
  const dl = series.dataLabels;
  if (!dl) return undefined;
  const flags = dl.points.get(i) ?? dl;
  if (flags.deleted) return undefined;
  if (!flags.showVal && !flags.showPercent && !flags.showCatName && !flags.showSerName && !flags.showBubbleSize) return undefined;
  return flags;
}

function labelText(flags: DataLabelFlags, series: ChartSeries, i: number, value: number, percent: number | undefined, xValue?: number): string {
  const parts: string[] = [];
  const code = flags.numFmt && !flags.numFmt.sourceLinked ? flags.numFmt.code : series.formatCode;
  if (flags.showSerName) parts.push(series.name);
  if (flags.showCatName) parts.push(xValue !== undefined ? formatNumber(xValue, series.categoryFormat) : (series.categories[i] ?? String(i + 1)));
  if (flags.showVal) parts.push(formatNumber(value, code));
  if (flags.showPercent && percent !== undefined) parts.push(formatNumber(percent, flags.numFmt && !flags.numFmt.sourceLinked && flags.numFmt.code.includes('%') ? flags.numFmt.code : '0%'));
  if (flags.showBubbleSize) {
    const b = series.bubbleSizes?.[i];
    if (b !== null && b !== undefined) parts.push(formatNumber(b, 'General'));
  }
  return parts.join(flags.separator ?? ', ');
}

function drawLabels(d: Draw, jobs: LabelJob[]): void {
  for (const job of jobs) {
    const lines = job.text.split('\n');
    if (job.fill) {
      const lh = d.lineHeight(job.style);
      const w = Math.max(...lines.map((l) => d.width(l, job.style)));
      const h = lh * lines.length;
      const left = job.h === 'left' ? job.x : job.h === 'center' ? job.x - w / 2 : job.x - w;
      const top = job.v === 'top' ? job.y : job.v === 'middle' ? job.y - h / 2 : job.y - h;
      d.path(rectPath({ x: left - 2, y: top, w: w + 4, h }), job.fill, null);
    }
    d.textBlock(lines, job.x, job.y, job.style, job.h, job.v);
  }
}

// ---------------------------------------------------------------------------------------------
// Axis charts (bar, line, area, scatter, bubble)

interface AxisInfo {
  axis: ChartAxis;
  /** 'cat': indexes 0..count-1; 'val': a numeric scale. */
  role: 'cat' | 'val';
  horizontal: boolean;
  partner?: AxisInfo;
  count: number;
  between: boolean;
  labels: string[];
  groups?: Array<Array<string | undefined>>;
  lo: number;
  hi: number;
  percent: boolean;
  format: string;
  scale?: Scale;
  reverse: boolean;
  /** Side of the plot the labels and title go on. */
  side: 'b' | 't' | 'l' | 'r';
  visible: boolean;
  /** Computed label layout. */
  catLayout?: CatLabelLayout;
}

interface CatLabelLayout {
  lines: string[][];
  skip: number;
  rotation: number;
  thickness: number;
}

const AXIS_TYPES = new Set(['bar', 'line', 'area', 'scatter', 'bubble']);

function syntheticAxis(id: string, kind: 'cat' | 'val', position: ChartAxis['position'], text: TextStyle): ChartAxis {
  return {
    id,
    kind,
    position,
    deleted: true,
    orientation: 'minMax',
    majorTickMark: 'none',
    minorTickMark: 'none',
    tickLblPos: 'none',
    crosses: 'autoZero',
    line: { paint: null, width: 0.75 },
    textStyle: text,
  };
}

/** Values of a series as plotted (blanks per dispBlanksAs). */
function plotted(v: number | null | undefined, blanks: ChartModel['dispBlanksAs']): number | null {
  if (v === null || v === undefined) return blanks === 'zero' ? 0 : null;
  return v;
}

function stackRange(plot: ChartPlot, model: ChartModel): { lo: number; hi: number } {
  let lo = Number.POSITIVE_INFINITY;
  let hi = Number.NEGATIVE_INFINITY;
  const n = Math.max(0, ...plot.series.map((s) => s.values.length));
  const stacked = plot.grouping === 'stacked' || plot.grouping === 'percentStacked';
  const splitSigns = plot.type === 'bar';
  for (let i = 0; i < n; i++) {
    if (!stacked) {
      for (const s of plot.series) {
        const v = plotted(s.values[i], model.dispBlanksAs);
        if (v === null) continue;
        lo = Math.min(lo, v);
        hi = Math.max(hi, v);
      }
      continue;
    }
    const total = plot.series.reduce((a, s) => a + Math.abs(s.values[i] ?? 0), 0);
    let pos = 0;
    let neg = 0;
    let cum = 0;
    for (const s of plot.series) {
      let v = s.values[i] ?? 0;
      if (plot.grouping === 'percentStacked') v = total ? v / total : 0;
      if (splitSigns) {
        if (v >= 0) pos += v;
        else neg += v;
        lo = Math.min(lo, neg);
        hi = Math.max(hi, pos);
      } else {
        cum += v;
        lo = Math.min(lo, cum);
        hi = Math.max(hi, cum);
      }
    }
  }
  return { lo, hi };
}

function catFrac(info: AxisInfo, i: number): number {
  if (info.between) return (i + 0.5) / Math.max(info.count, 1);
  return info.count > 1 ? i / (info.count - 1) : 0.5;
}

function valFrac(info: AxisInfo, v: number): number {
  const s = info.scale!;
  if (s.log) {
    const lv = Math.log(Math.max(v, s.min * 1e-6));
    return (lv - Math.log(s.min)) / (Math.log(s.max) - Math.log(s.min));
  }
  return (v - s.min) / (s.max - s.min || 1);
}

class AxisLayout {
  inner: Box = { x: 0, y: 0, w: 0, h: 0 };

  coord(info: AxisInfo, f: number): number {
    const g = info.reverse ? 1 - f : f;
    return info.horizontal ? this.inner.x + g * this.inner.w : this.inner.y + this.inner.h - g * this.inner.h;
  }

  /** Coordinate of a value on an axis, clamped to the plot. */
  at(info: AxisInfo, v: number): number {
    const f = info.role === 'cat' ? catFrac(info, v) : valFrac(info, v);
    return this.coord(info, Math.max(0, Math.min(1, f)));
  }

  /**
   * Where `info`'s axis line sits across the plot, as a fraction of its partner's length from the
   * bottom/left edge (reversal applied): 0 and 1 are the plot edges.
   */
  crossFrac(info: AxisInfo): number {
    const other = info.partner;
    if (!other) return 0;
    const a = info.axis;
    let f: number;
    if (other.role === 'val') {
      const s = other.scale!;
      const v = a.crossesAt ?? (a.crosses === 'max' ? s.max : a.crosses === 'min' ? s.min : s.log ? s.min : 0);
      f = valFrac(other, Math.max(s.min, Math.min(s.max, v)));
    } else if (a.crossesAt !== undefined) {
      const k = a.crossesAt - 1;
      f = other.between ? k / Math.max(other.count, 1) : other.count > 1 ? k / (other.count - 1) : 0;
    } else f = a.crosses === 'max' ? 1 : 0;
    f = Math.max(0, Math.min(1, f));
    return other.reverse ? 1 - f : f;
  }

  /** Page coordinate of `info`'s axis line (y for a horizontal axis, x for a vertical one). */
  cross(info: AxisInfo): number {
    const g = this.crossFrac(info);
    return info.horizontal ? this.inner.y + this.inner.h - g * this.inner.h : this.inner.x + g * this.inner.w;
  }

  /** Value (on the partner value axis) where `info` crosses: the base of bars and areas. */
  crossValue(info: AxisInfo): number {
    const other = info.partner;
    if (!other || other.role !== 'val') return 0;
    const s = other.scale!;
    const a = info.axis;
    const v = a.crossesAt ?? (a.crosses === 'max' ? s.max : a.crosses === 'min' ? s.min : s.log ? s.min : 0);
    return Math.max(s.min, Math.min(s.max, v));
  }
}

function axisFormat(info: AxisInfo, firstSeries: ChartSeries | undefined, xRole: boolean): string {
  const nf = info.axis.numFmt;
  let code: string;
  if (nf && !nf.sourceLinked) code = nf.code;
  else code = (xRole ? firstSeries?.categoryFormat : firstSeries?.formatCode) ?? 'General';
  if (info.percent && (code === 'General' || (nf?.sourceLinked ?? true))) code = '0%';
  return code;
}

function valueLabels(info: AxisInfo): string[] {
  const div = info.axis.dispUnits ?? 1;
  return ticks(info.scale!).map((v) => formatNumber(v / div, info.format));
}

function layoutCategoryLabels(d: Draw, info: AxisInfo, length: number, maxOther: number): CatLabelLayout {
  const st = info.axis.textStyle;
  const lh = d.lineHeight(st);
  const n = Math.max(info.count, 1);
  const slots = info.between ? n : Math.max(n - 1, 1);
  const band = length / slots;
  let skip = info.axis.tickLblSkip ?? 1;
  const explicitRot = st.rotation !== undefined && Math.abs(st.rotation) <= 90 ? st.rotation : undefined;
  if (info.horizontal) {
    if (explicitRot && explicitRot !== 0 && d.canRotate(st)) {
      const rad = (Math.abs(explicitRot) * Math.PI) / 180;
      const widths = info.labels.map((l) => d.width(l, st));
      const thick = Math.max(0, ...widths.map((w) => w * Math.sin(rad) + lh * Math.cos(rad)));
      if (info.axis.tickLblSkip === undefined && Math.abs(Math.sin(rad)) > 0.1) skip = Math.max(1, Math.ceil((lh / Math.sin(rad)) / band));
      return { lines: info.labels.map((l) => [l]), skip, rotation: explicitRot, thickness: thick };
    }
    const maxW = Math.max(band * skip - st.size * 0.4, st.size * 1.5);
    const lines = info.labels.map((l) => d.wrap(l, maxW, st).slice(0, 3));
    const fits = lines.every((ls) => ls.every((l) => d.width(l, st) <= band * skip));
    if (fits) return { lines, skip, rotation: 0, thickness: Math.max(1, ...lines.map((l) => l.length)) * lh };
    if (info.axis.tickLblSkip === undefined && d.canRotate(st) && explicitRot === undefined) {
      // Office turns crowded category labels 45° before it starts skipping them.
      const widths = info.labels.map((l) => Math.min(d.width(l, st), maxOther));
      const rotSkip = Math.max(1, Math.ceil((lh * Math.SQRT2) / band));
      return { lines: info.labels.map((l) => [l]), skip: rotSkip, rotation: -45, thickness: Math.max(0, ...widths.map((w) => (w + lh) * Math.SQRT1_2)) };
    }
    const widest = Math.max(0, ...info.labels.map((l) => d.width(l, st)));
    if (info.axis.tickLblSkip === undefined) skip = Math.max(1, Math.ceil((widest + st.size * 0.5) / band));
    return { lines: info.labels.map((l) => [l]), skip, rotation: 0, thickness: lh };
  }
  // Vertical category axis (bar charts): labels wrap to a share of the width, then skip if crowded.
  const maxW = Math.max(maxOther * 0.4, st.size * 3);
  const lines = info.labels.map((l) => d.wrap(l, maxW, st).slice(0, 3));
  const tallest = Math.max(1, ...lines.map((l) => l.length)) * lh;
  if (info.axis.tickLblSkip === undefined) skip = Math.max(1, Math.ceil(tallest / band - 1e-6));
  const thickness = Math.max(0, ...lines.map((ls) => Math.max(0, ...ls.map((l) => d.width(l, st)))));
  return { lines, skip, rotation: 0, thickness };
}

function tickOut(axis: ChartAxis): number {
  return axis.majorTickMark === 'out' || axis.majorTickMark === 'cross' ? TICK : 0;
}

/** Space an axis needs beside the plot: its labels (with ticks) and its title, in points. */
function axisThickness(d: Draw, info: AxisInfo, inner: Box, outer: Box): { labels: number; title: number } {
  if (!info.visible) return { labels: 0, title: 0 };
  const a = info.axis;
  const st = a.textStyle;
  let t = tickOut(a);
  if (a.tickLblPos !== 'none') {
    t += LABEL_GAP;
    const length = info.horizontal ? inner.w : inner.h;
    if (info.role === 'cat') {
      info.catLayout = layoutCategoryLabels(d, info, length, info.horizontal ? outer.h * 0.4 : outer.w);
      t += info.catLayout.thickness;
      if (info.horizontal && info.groups) t += info.groups.length * (d.lineHeight(st) + LABEL_GAP);
    } else if (info.horizontal) t += d.lineHeight(st);
    else t += Math.max(0, ...valueLabels(info).map((l) => d.width(l, st)));
  }
  const title = a.title ? a.title.lines.length * d.lineHeight(a.title.style) + LABEL_GAP * 2 : 0;
  return { labels: t, title };
}

/**
 * Which side of its line an axis' labels go: "low"/"high" follow the partner's minimum/maximum end;
 * "nextTo" labels go outside the plot when the line sits on an edge (a reversed partner moves it to
 * the far edge), and on the axPos side when the line crosses inside the plot.
 */
function labelSide(info: AxisInfo, layout: AxisLayout): AxisInfo['side'] {
  const far: AxisInfo['side'] = info.horizontal ? 't' : 'r';
  const near: AxisInfo['side'] = info.horizontal ? 'b' : 'l';
  const minAtFar = info.partner?.reverse ?? false;
  if (info.axis.tickLblPos === 'low') return minAtFar ? far : near;
  if (info.axis.tickLblPos === 'high') return minAtFar ? near : far;
  const g = layout.crossFrac(info);
  if (g >= 1 - 1e-9) return far;
  if (g <= 1e-9) return near;
  return info.horizontal ? (info.axis.position === 't' ? 't' : 'b') : info.axis.position === 'r' ? 'r' : 'l';
}

/** The plot edge on an axis' label side. */
function edgeOf(inner: Box, side: AxisInfo['side']): number {
  return side === 'b' ? inner.y + inner.h : side === 't' ? inner.y : side === 'l' ? inner.x : inner.x + inner.w;
}

function drawAxisChart(d: Draw, model: ChartModel, outer: Box, manualInner: Box | undefined, labels: LabelJob[]): void {
  const plots = model.plots.filter((p) => AXIS_TYPES.has(p.type));
  const axesById = new Map(model.axes.map((a) => [a.id, a]));
  const infos = new Map<string, AxisInfo>();
  const pairs: Array<{ plot: ChartPlot; x: AxisInfo; y: AxisInfo }> = [];

  plots.forEach((plot, pi) => {
    const xy = plot.type === 'scatter' || plot.type === 'bubble';
    const horizontalCats = !(plot.type === 'bar' && plot.barDir === 'bar');
    const ids = [plot.axisIds[0] ?? `#x${pi}`, plot.axisIds[1] ?? `#y${pi}`] as const;
    const make = (id: string, isX: boolean): AxisInfo => {
      let info = infos.get(id);
      if (info) return info;
      const horizontal = isX ? horizontalCats : !horizontalCats;
      const axis = axesById.get(id) ?? syntheticAxis(id, isX && !xy ? 'cat' : 'val', horizontal ? 'b' : 'l', model.textStyle);
      const role = isX && !xy ? 'cat' : 'val';
      const side: AxisInfo['side'] = horizontal ? (axis.position === 't' ? 't' : 'b') : axis.position === 'r' ? 'r' : 'l';
      info = {
        axis,
        role,
        horizontal,
        count: 0,
        between: true,
        labels: [],
        lo: Number.POSITIVE_INFINITY,
        hi: Number.NEGATIVE_INFINITY,
        percent: false,
        format: 'General',
        reverse: axis.orientation === 'maxMin',
        side,
        visible: !axis.deleted,
      };
      infos.set(id, info);
      return info;
    };
    const x = make(ids[0], true);
    const y = make(ids[1], false);
    x.partner ??= y;
    y.partner ??= x;
    pairs.push({ plot, x, y });
  });

  // Data extents and category counts.
  for (const { plot, x, y } of pairs) {
    const first = plot.series[0];
    if (x.role === 'cat') {
      const n = Math.max(0, ...plot.series.map((s) => Math.max(s.values.length, s.categories.length)));
      if (n > x.count) {
        x.count = n;
        const src = plot.series.find((s) => s.categories.length > 0) ?? first;
        const nf = x.axis.numFmt;
        x.labels = Array.from({ length: n }, (_, i) => {
          const num = src?.categoryNumbers?.[i];
          if (num !== null && num !== undefined && nf && !nf.sourceLinked) return formatNumber(num, nf.code);
          if (num !== null && num !== undefined && src?.categoryFormat && isDateCode(src.categoryFormat)) return formatNumber(num, src.categoryFormat);
          return src?.categories[i] ?? String(i + 1);
        });
        if (src?.categoryGroups) x.groups = src.categoryGroups;
      }
      // Bars sit between tick marks; lines too unless the value axis says midCat (areas default to it).
      const cb = y.axis.crossBetween ?? (plot.type === 'area' ? 'midCat' : 'between');
      if (plot.type === 'bar') x.between = true;
      else if (!pairs.some((p) => p.x === x && p.plot.type === 'bar')) x.between = cb === 'between';
    } else {
      for (const s of plot.series) {
        s.values.forEach((_, i) => {
          const xv = s.categoryNumbers?.[i] ?? (s.categoryNumbers ? null : i + 1);
          if (xv === null) return;
          x.lo = Math.min(x.lo, xv);
          x.hi = Math.max(x.hi, xv);
        });
      }
      x.format = axisFormat(x, first, true);
    }
    const r = stackRange(plot, model);
    y.lo = Math.min(y.lo, r.lo);
    y.hi = Math.max(y.hi, r.hi);
    if (plot.grouping === 'percentStacked') y.percent = true;
    y.format = axisFormat(y, first, false);
  }

  const all = [...infos.values()];
  const layout = new AxisLayout();
  const computeScales = (inner: Box): void => {
    for (const info of all) {
      if (info.role !== 'val') continue;
      const st = info.axis.textStyle;
      const lh = d.lineHeight(st);
      const length = info.horizontal ? inner.w : inner.h;
      info.scale = niceScale(info.lo, info.hi, info.axis, length, info.percent, (s) => {
        if (!info.horizontal) return lh * 1.4;
        const div = info.axis.dispUnits ?? 1;
        const widest = Math.max(...ticks(s).map((v) => d.width(formatNumber(v / div, info.format), st)));
        return widest * 1.5 + st.size * 1.5;
      });
    }
  };

  let inner = manualInner ?? { ...outer };
  for (let pass = 0; pass < 3; pass++) {
    computeScales(inner);
    if (manualInner) {
      layout.inner = inner;
      for (const info of all) info.side = labelSide(info, layout);
      for (const info of all) axisThickness(d, info, inner, outer);
      break;
    }
    const t = { b: 0, t: 0, l: 0, r: 0 };
    layout.inner = inner;
    for (const info of all) info.side = labelSide(info, layout);
    for (const info of all) {
      const th = axisThickness(d, info, inner, outer);
      // Labels "next to" an axis line inside the plot (negative values) hang inside it, as in Office.
      const inside = info.axis.tickLblPos === 'nextTo' && info.partner ? Math.abs(edgeOf(inner, info.side) - layout.cross(info)) : 0;
      t[info.side] += Math.max(0, th.labels - inside) + th.title;
    }
    // Value labels centred on the end ticks stick out past the plot edges.
    for (const info of all) {
      if (!info.visible || info.axis.tickLblPos === 'none') continue;
      const st = info.axis.textStyle;
      if (info.role === 'val' && !info.horizontal) {
        const half = d.lineHeight(st) / 2;
        t.t = Math.max(t.t, half);
        t.b = Math.max(t.b, half);
      } else if (info.role === 'val') {
        const labs = valueLabels(info);
        t.l = Math.max(t.l, d.width(labs[0] ?? '', st) / 2);
        t.r = Math.max(t.r, d.width(labs[labs.length - 1] ?? '', st) / 2);
      }
    }
    inner = { x: outer.x + t.l, y: outer.y + t.t, w: Math.max(outer.w - t.l - t.r, 10), h: Math.max(outer.h - t.t - t.b, 10) };
  }
  layout.inner = inner;

  // Plot area, then gridlines.
  d.path(rectPath(inner), model.plotArea.fill, null);
  for (const info of all) drawGridlines(d, layout, info, 'minor');
  for (const info of all) drawGridlines(d, layout, info, 'major');

  // Plots: areas at the back, then bars, then lines and points.
  const rank: Record<string, number> = { area: 0, bar: 1, line: 2, scatter: 3, bubble: 3 };
  const ordered = [...pairs].sort((a, b) => (rank[a.plot.type] ?? 9) - (rank[b.plot.type] ?? 9));
  for (const pair of ordered) {
    switch (pair.plot.type) {
      case 'bar':
        drawBars(d, model, layout, pair.plot, pair.x, pair.y, labels);
        break;
      case 'area':
        drawAreas(d, model, layout, pair.plot, pair.x, pair.y, labels);
        break;
      case 'line':
        drawLines(d, model, layout, pair.plot, pair.x, pair.y, labels);
        break;
      default:
        drawScatter(d, model, layout, pair.plot, pair.x, pair.y, labels);
    }
  }

  d.path(rectPath(inner), null, model.plotArea.line);
  for (const info of all) drawAxis(d, layout, info, outer);
}

function drawGridlines(d: Draw, layout: AxisLayout, info: AxisInfo, which: 'major' | 'minor'): void {
  const style = which === 'major' ? info.axis.majorGridlines : info.axis.minorGridlines;
  if (!style?.paint) return;
  const { inner } = layout;
  const positions: number[] = [];
  if (info.role === 'val') {
    const s = info.scale!;
    if (which === 'major') for (const v of ticks(s)) positions.push(layout.coord(info, valFrac(info, v)));
    else if (!s.log) {
      const n = Math.round((s.max - s.min) / s.minor);
      for (let i = 0; i <= Math.min(n, 500); i++) positions.push(layout.coord(info, valFrac(info, s.min + i * s.minor)));
    }
  } else {
    const n = info.count;
    if (info.between) for (let i = 0; i <= n; i++) positions.push(layout.coord(info, i / Math.max(n, 1)));
    else for (let i = 0; i < n; i++) positions.push(layout.coord(info, catFrac(info, i)));
  }
  for (const p of positions) {
    if (info.horizontal) d.segment(p, inner.y, p, inner.y + inner.h, style);
    else d.segment(inner.x, p, inner.x + inner.w, p, style);
  }
}

function drawAxis(d: Draw, layout: AxisLayout, info: AxisInfo, outer: Box): void {
  if (!info.visible) return;
  const a = info.axis;
  const { inner } = layout;
  const st = a.textStyle;
  const lh = d.lineHeight(st);
  const line = layout.cross(info);
  const outward = info.side === 'b' || info.side === 'r' ? 1 : -1;

  // Axis line.
  if (info.horizontal) d.segment(inner.x, line, inner.x + inner.w, line, a.line);
  else d.segment(line, inner.y, line, inner.y + inner.h, a.line);

  // Tick positions along the axis.
  const tickPos: number[] = [];
  if (info.role === 'val') for (const v of ticks(info.scale!)) tickPos.push(layout.coord(info, valFrac(info, v)));
  else if (info.between) for (let i = 0; i <= info.count; i++) tickPos.push(layout.coord(info, i / Math.max(info.count, 1)));
  else for (let i = 0; i < info.count; i++) tickPos.push(layout.coord(info, catFrac(info, i)));
  const mark = a.majorTickMark;
  if (mark !== 'none' && a.line.paint) {
    const from = mark === 'in' ? -TICK : mark === 'cross' ? -TICK : 0;
    const to = mark === 'in' ? 0 : TICK;
    for (const p of tickPos) {
      if (info.horizontal) d.segment(p, line + outward * from, p, line + outward * to, a.line);
      else d.segment(line + outward * from, p, line + outward * to, p, a.line);
    }
  }

  // Labels sit next to the axis line, or at the plot edge for low/high.
  let base = line;
  if (a.tickLblPos === 'low' || a.tickLblPos === 'high') base = info.horizontal ? (info.side === 'b' ? inner.y + inner.h : inner.y) : info.side === 'l' ? inner.x : inner.x + inner.w;
  const labelStart = base + outward * (tickOut(a) + LABEL_GAP);
  let extent = 0;
  if (a.tickLblPos !== 'none') {
    if (info.role === 'val') {
      const values = ticks(info.scale!);
      const labs = valueLabels(info);
      values.forEach((v, k) => {
        const p = layout.coord(info, valFrac(info, v));
        const text = labs[k] ?? '';
        if (info.horizontal) d.textBlock([text], p, labelStart, st, 'center', outward > 0 ? 'top' : 'bottom');
        else d.textBlock([text], labelStart, p, st, outward > 0 ? 'left' : 'right', 'middle');
      });
      extent = info.horizontal ? lh : Math.max(0, ...labs.map((l) => d.width(l, st)));
    } else {
      const cl = info.catLayout ?? layoutCategoryLabels(d, info, info.horizontal ? inner.w : inner.h, outer.w);
      for (let i = 0; i < info.count; i += cl.skip) {
        const p = layout.coord(info, catFrac(info, i));
        const lines = cl.lines[i] ?? [];
        if (cl.rotation !== 0 && info.horizontal) {
          // Rotated labels end at the tick (Office's -45° labels read up towards it).
          const text = lines.join(' ');
          const rad = (cl.rotation * Math.PI) / 180;
          const along = cl.rotation < 0 ? 1 : 0;
          const cy = labelStart + outward * (lh / 2) * Math.abs(Math.cos(rad));
          if (!d.rotatedText(text, p, cy, st, cl.rotation, along)) d.textBlock([text], p, labelStart, st, 'center', outward > 0 ? 'top' : 'bottom');
        } else if (info.horizontal) d.textBlock(lines, p, labelStart, st, 'center', outward > 0 ? 'top' : 'bottom');
        else d.textBlock(lines, labelStart, p, st, outward > 0 ? 'left' : 'right', 'middle');
      }
      extent = cl.thickness;
      // Outer levels of multi-level categories: one row per level, labels centred over their groups.
      if (info.horizontal && info.groups) {
        let rowStart = labelStart + outward * (extent + LABEL_GAP);
        for (const level of info.groups) {
          const starts: number[] = [];
          level.forEach((v, i) => {
            if (v !== undefined) starts.push(i);
          });
          starts.forEach((s, k) => {
            const e = (starts[k + 1] ?? info.count) - 1;
            const p1 = layout.coord(info, s / Math.max(info.count, 1));
            const p2 = layout.coord(info, (e + 1) / Math.max(info.count, 1));
            d.textBlock([level[s] ?? ''], (p1 + p2) / 2, rowStart, st, 'center', outward > 0 ? 'top' : 'bottom');
            // Dividers between groups, and closing the last one.
            d.segment(p1, rowStart - outward * LABEL_GAP, p1, rowStart + outward * lh, a.line);
            if (k === starts.length - 1) d.segment(p2, rowStart - outward * LABEL_GAP, p2, rowStart + outward * lh, a.line);
          });
          rowStart += outward * (lh + LABEL_GAP);
          extent += lh + LABEL_GAP;
        }
      }
    }
  }

  // Axis title beyond the labels; vertical axes turn it to read bottom to top.
  if (a.title) {
    const t = a.title;
    const tlh = d.lineHeight(t.style);
    const labelsEnd = labelStart + outward * extent;
    const edge = edgeOf(inner, info.side);
    const pos = (outward > 0 ? Math.max(labelsEnd, edge) : Math.min(labelsEnd, edge)) + outward * LABEL_GAP;
    if (info.horizontal) d.textBlock(t.lines, inner.x + inner.w / 2, pos, t.style, 'center', outward > 0 ? 'top' : 'bottom');
    else {
      const rot = t.style.rotation !== undefined && Math.abs(t.style.rotation) <= 90 ? t.style.rotation : -90;
      const cx = pos + (outward * (t.lines.length * tlh)) / 2;
      const cy = inner.y + inner.h / 2;
      const text = t.lines.join(' ');
      if (rot === 0 || !d.rotatedText(text, cx, cy, t.style, rot, 0.5)) {
        // Without outlines the title stays horizontal, above the plot on the axis' side.
        if (outward > 0) d.textBlock(t.lines, outer.x + outer.w, inner.y - LABEL_GAP, t.style, 'right', 'bottom');
        else d.textBlock(t.lines, outer.x, inner.y - LABEL_GAP, t.style, 'left', 'bottom');
      }
    }
  }
}

function barRect(layout: AxisLayout, x: AxisInfo, y: AxisInfo, f0: number, f1: number, from: number, to: number): Box {
  const a = layout.coord(x, f0);
  const b = layout.coord(x, f1);
  const c = layout.coord(y, Math.max(0, Math.min(1, valFrac(y, from))));
  const e = layout.coord(y, Math.max(0, Math.min(1, valFrac(y, to))));
  if (x.horizontal) return { x: Math.min(a, b), y: Math.min(c, e), w: Math.abs(b - a), h: Math.abs(e - c) };
  return { x: Math.min(c, e), y: Math.min(a, b), w: Math.abs(e - c), h: Math.abs(b - a) };
}

function drawBars(d: Draw, model: ChartModel, layout: AxisLayout, plot: ChartPlot, x: AxisInfo, y: AxisInfo, labels: LabelJob[]): void {
  const n = Math.max(x.count, 1);
  const stacked = plot.grouping === 'stacked' || plot.grouping === 'percentStacked';
  const slots = stacked ? 1 : Math.max(plot.series.length, 1);
  const ov = Math.max(-100, Math.min(100, plot.overlap)) / 100;
  const gap = Math.max(0, Math.min(500, plot.gapWidth)) / 100;
  const band = 1 / n;
  const barW = band / (slots - (slots - 1) * ov + gap);
  const groupW = barW * (slots - (slots - 1) * ov);
  const base = layout.crossValue(x);
  const pos = new Array<number>(n).fill(0);
  const neg = new Array<number>(n).fill(0);
  const totals = Array.from({ length: n }, (_, i) => plot.series.reduce((a, s) => a + Math.abs(s.values[i] ?? 0), 0));
  const varied = isVaried(plot);
  plot.series.forEach((s, si) => {
    for (let i = 0; i < n; i++) {
      const raw = plotted(s.values[i], model.dispBlanksAs);
      if (raw === null) continue;
      const v = plot.grouping === 'percentStacked' ? (totals[i] ? raw / totals[i]! : 0) : raw;
      let from: number;
      let to: number;
      if (stacked) {
        if (v >= 0) {
          from = pos[i]!;
          to = from + v;
          pos[i] = to;
        } else {
          from = neg[i]!;
          to = from + v;
          neg[i] = to;
        }
      } else {
        from = y.scale!.log ? y.scale!.min : base;
        to = v;
      }
      const f0 = i * band + (band - groupW) / 2 + (stacked ? 0 : si) * barW * (1 - ov);
      const box = barRect(layout, x, y, f0, f0 + barW, from, to);
      const pt = s.points.find((p) => p.index === i);
      let fill = pt?.fill !== undefined ? pt.fill : varied && s.pointColors?.[i] ? { color: s.pointColors[i]!, opacity: 1 } : s.fill;
      const line = pt?.line ?? s.line;
      if ((pt?.invertIfNegative ?? s.invertIfNegative) && v < 0) fill = { color: 'FFFFFF', opacity: 1 };
      d.path(rectPath(box), fill, line, { lineJoin: 'miter' });

      const flags = labelFlags(s, i);
      if (flags) {
        const text = labelText(flags, s, i, raw, undefined);
        const position = flags.position ?? (stacked ? 'ctr' : 'outEnd');
        const up = to >= from !== y.reverse;
        labels.push(barLabel(text, box, x.horizontal, up, position, flags));
      }
    }
  });
}

function barLabel(text: string, box: Box, vertical: boolean, positive: boolean, position: string, flags: DataLabelFlags): LabelJob {
  const job = (xx: number, yy: number, h: LabelJob['h'], v: LabelJob['v']): LabelJob => {
    const j: LabelJob = { text, x: xx, y: yy, h, v, style: flags.style };
    if (flags.fill) j.fill = flags.fill;
    return j;
  };
  const g = 2;
  if (vertical) {
    const cx = box.x + box.w / 2;
    const endY = positive ? box.y : box.y + box.h;
    const baseY = positive ? box.y + box.h : box.y;
    switch (position) {
      case 'ctr':
        return job(cx, box.y + box.h / 2, 'center', 'middle');
      case 'inEnd':
        return job(cx, endY + (positive ? g : -g), 'center', positive ? 'top' : 'bottom');
      case 'inBase':
        return job(cx, baseY + (positive ? -g : g), 'center', positive ? 'bottom' : 'top');
      default:
        return job(cx, endY + (positive ? -g : g), 'center', positive ? 'bottom' : 'top');
    }
  }
  const cy = box.y + box.h / 2;
  const endX = positive ? box.x + box.w : box.x;
  const baseX = positive ? box.x : box.x + box.w;
  switch (position) {
    case 'ctr':
      return job(box.x + box.w / 2, cy, 'center', 'middle');
    case 'inEnd':
      return job(endX + (positive ? -g : g), cy, positive ? 'right' : 'left', 'middle');
    case 'inBase':
      return job(baseX + (positive ? g : -g), cy, positive ? 'left' : 'right', 'middle');
    default:
      return job(endX + (positive ? g : -g), cy, positive ? 'left' : 'right', 'middle');
  }
}

/** Plotted points of each series of a line or area plot (stacked and percent-stacked applied). */
function seriesPoints(model: ChartModel, layout: AxisLayout, plot: ChartPlot, x: AxisInfo, y: AxisInfo, area: boolean): Array<Array<{ x: number; y: number; v: number; raw: number; i: number; lower: number } | null>> {
  const n = x.count;
  const cum = new Array<number>(n).fill(0);
  const stacked = plot.grouping === 'stacked' || plot.grouping === 'percentStacked';
  const totals = Array.from({ length: n }, (_, i) => plot.series.reduce((a, s) => a + Math.abs(s.values[i] ?? 0), 0));
  return plot.series.map((s) => {
    const pts: Array<{ x: number; y: number; v: number; raw: number; i: number; lower: number } | null> = [];
    for (let i = 0; i < n; i++) {
      // Stacked and area series treat blanks as zero so the stack stays closed.
      const raw = stacked || area ? (s.values[i] ?? 0) : plotted(s.values[i], model.dispBlanksAs);
      if (raw === null || i >= s.values.length) {
        pts.push(null);
        continue;
      }
      let v = plot.grouping === 'percentStacked' ? (totals[i] ? raw / totals[i]! : 0) : raw;
      const lower = stacked ? cum[i]! : 0;
      if (stacked) {
        v += cum[i]!;
        cum[i] = v;
      }
      pts.push({ x: 0, y: 0, v, raw, i, lower });
    }
    for (const p of pts) {
      if (!p) continue;
      const a = layout.coord(x, catFrac(x, p.i));
      const b = layout.at(y, p.v);
      if (x.horizontal) {
        p.x = a;
        p.y = b;
      } else {
        p.x = b;
        p.y = a;
      }
    }
    return pts;
  });
}

/** Split a point list at gaps (or join across them for "span"). */
function runsOf<T>(pts: Array<T | null>, span: boolean): T[][] {
  const runs: T[][] = [];
  let cur: T[] = [];
  for (const p of pts) {
    if (p === null) {
      if (!span && cur.length) {
        runs.push(cur);
        cur = [];
      }
      continue;
    }
    cur.push(p);
  }
  if (cur.length) runs.push(cur);
  return runs;
}

function pointLabel(text: string, px: number, py: number, position: string, flags: DataLabelFlags): LabelJob {
  const g = 3;
  const job: LabelJob = { text, x: px + g, y: py, h: 'left', v: 'middle', style: flags.style };
  switch (position) {
    case 't':
      Object.assign(job, { x: px, y: py - g, h: 'center', v: 'bottom' });
      break;
    case 'b':
      Object.assign(job, { x: px, y: py + g, h: 'center', v: 'top' });
      break;
    case 'l':
      Object.assign(job, { x: px - g, h: 'right' });
      break;
    case 'ctr':
      Object.assign(job, { x: px, h: 'center' });
      break;
    default:
      break;
  }
  if (flags.fill) job.fill = flags.fill;
  return job;
}

function drawLines(d: Draw, model: ChartModel, layout: AxisLayout, plot: ChartPlot, x: AxisInfo, y: AxisInfo, labels: LabelJob[]): void {
  const all = seriesPoints(model, layout, plot, x, y, false);
  plot.series.forEach((s, si) => {
    const pts = all[si]!;
    if (s.line.paint) {
      for (const run of runsOf(pts, model.dispBlanksAs === 'span')) {
        if (run.length < 2) continue;
        d.path(s.smooth ? smoothPath(run) : polyline(run), null, s.line);
      }
    }
    for (const p of pts) {
      if (!p) continue;
      const pt = s.points.find((q) => q.index === p.i);
      drawMarker(d, pt?.marker ?? s.marker, p.x, p.y);
      const flags = labelFlags(s, p.i);
      if (flags) labels.push(pointLabel(labelText(flags, s, p.i, p.raw, undefined), p.x, p.y, flags.position ?? 'r', flags));
    }
  });
}

function drawAreas(d: Draw, model: ChartModel, layout: AxisLayout, plot: ChartPlot, x: AxisInfo, y: AxisInfo, labels: LabelJob[]): void {
  const all = seriesPoints(model, layout, plot, x, y, true);
  const stacked = plot.grouping === 'stacked' || plot.grouping === 'percentStacked';
  const baseValue = layout.crossValue(x);
  plot.series.forEach((s, si) => {
    const pts = all[si]!.filter((p): p is NonNullable<typeof p> => p !== null);
    if (pts.length === 0) return;
    const cmds: PathCommand[] = polyline(pts);
    // Down the lower edge: the previous stack level, or the axis crossing.
    for (let k = pts.length - 1; k >= 0; k--) {
      const p = pts[k]!;
      const lowerV = stacked ? p.lower : baseValue;
      const c = layout.at(y, lowerV);
      cmds.push(x.horizontal ? { op: 'L', x: p.x, y: c } : { op: 'L', x: c, y: p.y });
    }
    cmds.push({ op: 'Z' });
    d.path(cmds, s.fill, s.line.paint ? s.line : null, { lineJoin: 'round' });
    for (const p of pts) {
      const flags = labelFlags(s, p.i);
      if (!flags) continue;
      const lowerC = layout.at(y, stacked ? p.lower : baseValue);
      const mx = x.horizontal ? p.x : (p.x + lowerC) / 2;
      const my = x.horizontal ? (p.y + lowerC) / 2 : p.y;
      labels.push(pointLabel(labelText(flags, s, p.i, p.raw, undefined), mx, my, 'ctr', flags));
    }
  });
}

function drawScatter(d: Draw, model: ChartModel, layout: AxisLayout, plot: ChartPlot, x: AxisInfo, y: AxisInfo, labels: LabelJob[]): void {
  const bubble = plot.type === 'bubble';
  const maxSize = bubble ? Math.max(0, ...plot.series.flatMap((s) => (s.bubbleSizes ?? []).map((b) => Math.abs(b ?? 0)))) : 0;
  // The largest bubble's diameter is a quarter of the plot's smaller side at 100% scale.
  const maxR = (Math.min(layout.inner.w, layout.inner.h) / 8) * (plot.bubbleScale / 100);
  for (const s of plot.series) {
    const pts: Array<{ x: number; y: number; i: number; xv: number; v: number } | null> = s.values.map((v, i) => {
      const yv = plotted(v, model.dispBlanksAs);
      const xv = s.categoryNumbers ? s.categoryNumbers[i] : i + 1;
      if (yv === null || xv === null || xv === undefined) return null;
      const px = layout.at(x, xv);
      const py = layout.at(y, yv);
      return x.horizontal ? { x: px, y: py, i, xv, v: yv } : { x: py, y: px, i, xv, v: yv };
    });
    if (bubble) {
      for (const p of pts) {
        if (!p) continue;
        const size = Math.abs(s.bubbleSizes?.[p.i] ?? 0);
        const r = maxSize > 0 ? Math.sqrt(size / maxSize) * maxR : 0;
        if (r > 0) d.path(circlePath(p.x, p.y, r), pointFill(s, p.i), pointLine(s, p.i));
      }
    } else {
      if (s.line.paint) {
        for (const run of runsOf(pts, model.dispBlanksAs === 'span')) {
          if (run.length < 2) continue;
          d.path(s.smooth ? smoothPath(run) : polyline(run), null, s.line);
        }
      }
      for (const p of pts) if (p) drawMarker(d, s.points.find((q) => q.index === p.i)?.marker ?? s.marker, p.x, p.y);
    }
    for (const p of pts) {
      if (!p) continue;
      const flags = labelFlags(s, p.i);
      if (flags) labels.push(pointLabel(labelText(flags, s, p.i, p.v, undefined, p.xv), p.x, p.y, flags.position ?? (bubble ? 'ctr' : 'r'), flags));
    }
  }
}

// ---------------------------------------------------------------------------------------------
// Pie and doughnut

function drawPie(d: Draw, model: ChartModel, outer: Box, labels: LabelJob[]): void {
  const plot = model.plots.find((p) => p.type === 'pie' || p.type === 'doughnut')!;
  const doughnut = plot.type === 'doughnut';
  const series = doughnut ? plot.series : plot.series.slice(0, 1);
  if (series.length === 0) return;
  const cx = outer.x + outer.w / 2;
  const cy = outer.y + outer.h / 2;

  // Slices of each ring: angles in degrees clockwise from 12 o'clock.
  const rings = series.map((s) => {
    const total = s.values.reduce<number>((a, v) => a + Math.abs(v ?? 0), 0);
    let angle = plot.firstSliceAng;
    return s.values.map((v, i) => {
      const value = Math.abs(v ?? 0);
      const sweep = total > 0 ? (value / total) * 360 : 0;
      const slice = { i, start: angle, sweep, value: v ?? 0, percent: total > 0 ? value / total : 0 };
      angle += sweep;
      return slice;
    });
  });
  const maxExplosion = Math.max(0, ...series.flatMap((s) => [s.explosion, ...s.points.map((p) => p.explosion ?? 0)]));

  // Labels decide how much room the pie leaves around itself.
  const first = series[series.length - 1]!;
  const ringSlices = rings[series.length - 1]!;
  const texts = ringSlices.map((sl) => {
    const flags = labelFlags(first, sl.i);
    return flags ? { flags, text: labelText(flags, first, sl.i, sl.value, sl.percent) } : undefined;
  });
  const lhOf = (st: TextStyle): number => d.lineHeight(st);
  let radius = Math.max(4, Math.min(outer.w, outer.h) / 2) / (1 + maxExplosion / 100);
  const outside = new Set<number>();
  if (!doughnut) {
    const r0 = radius * 0.95;
    texts.forEach((t, k) => {
      if (!t) return;
      const pos = t.flags.position ?? 'bestFit';
      if (pos === 'outEnd') outside.add(k);
      else if (pos === 'bestFit') {
        const sl = ringSlices[k]!;
        const lines = t.text.split('\n');
        const w = Math.max(...lines.map((l) => d.width(l, t.flags.style)));
        const h = lines.length * lhOf(t.flags.style);
        const rho = r0 * 0.62;
        const arc = (Math.min(sl.sweep, 180) * Math.PI) / 180 * rho;
        if (arc < w * 0.9 || r0 - rho < h * 0.6) outside.add(k);
      }
    });
    if (outside.size) {
      let mx = 0;
      let my = 0;
      for (const k of outside) {
        const t = texts[k]!;
        const lines = t.text.split('\n');
        mx = Math.max(mx, ...lines.map((l) => d.width(l, t.flags.style)));
        my = Math.max(my, lines.length * lhOf(t.flags.style));
      }
      radius = Math.max(4, Math.min((outer.w - 2 * (mx + 6)) / 2, (outer.h - 2 * (my + 6)) / 2)) / (1 + maxExplosion / 100);
    } else radius *= 0.95;
  } else radius *= 0.95;

  const hole = doughnut ? (Math.max(10, Math.min(90, plot.holeSize)) / 100) * radius : 0;
  const ringW = (radius - hole) / series.length;
  const varied = plot.varyColors;
  series.forEach((s, si) => {
    const r1 = doughnut ? hole + si * ringW : 0;
    const r2 = doughnut ? hole + (si + 1) * ringW : radius;
    for (const sl of rings[si]!) {
      if (sl.sweep <= 0) continue;
      const pt = s.points.find((p) => p.index === sl.i);
      const ex = ((pt?.explosion ?? s.explosion) / 100) * radius;
      const mid = sl.start + sl.sweep / 2;
      const off = ex > 0 ? polar(0, 0, ex, mid) : { x: 0, y: 0 };
      const ox = cx + off.x;
      const oy = cy + off.y;
      const cmds: PathCommand[] = [];
      if (sl.sweep >= 359.999) {
        cmds.push(...circlePath(ox, oy, r2));
        if (r1 > 0) cmds.push(...circlePath(ox, oy, r1));
      } else if (r1 > 0) {
        const a = polar(ox, oy, r2, sl.start);
        cmds.push({ op: 'M', x: a.x, y: a.y });
        arcTo(cmds, ox, oy, r2, sl.start, sl.start + sl.sweep);
        const b = polar(ox, oy, r1, sl.start + sl.sweep);
        cmds.push({ op: 'L', x: b.x, y: b.y });
        arcTo(cmds, ox, oy, r1, sl.start + sl.sweep, sl.start);
        cmds.push({ op: 'Z' });
      } else {
        const a = polar(ox, oy, r2, sl.start);
        cmds.push({ op: 'M', x: ox, y: oy }, { op: 'L', x: a.x, y: a.y });
        arcTo(cmds, ox, oy, r2, sl.start, sl.start + sl.sweep);
        cmds.push({ op: 'Z' });
      }
      const fill = pt?.fill !== undefined ? pt.fill : varied && s.pointColors?.[sl.i] ? { color: s.pointColors[sl.i]!, opacity: 1 } : s.fill;
      d.path(cmds, fill, pt?.line ?? s.line, { evenOdd: sl.sweep >= 359.999 && r1 > 0, lineJoin: 'round' });

      const flags = labelFlags(s, sl.i);
      if (!flags) continue;
      const text = labelText(flags, s, sl.i, sl.value, sl.percent);
      if (!text) continue;
      const pos = flags.position ?? 'bestFit';
      const job: LabelJob = { text, x: 0, y: 0, h: 'center', v: 'middle', style: flags.style };
      if (flags.fill) job.fill = flags.fill;
      if (doughnut || !outside.has(sl.i) || pos === 'ctr' || pos === 'inEnd' || pos === 'inBase') {
        const rho = doughnut ? (r1 + r2) / 2 : pos === 'inEnd' ? r2 * 0.78 : pos === 'ctr' ? r2 * 0.5 : r2 * 0.62;
        const p = polar(ox, oy, rho, mid);
        job.x = p.x;
        job.y = p.y;
      } else {
        const p = polar(ox, oy, r2 + 5, mid);
        job.x = p.x;
        job.y = p.y;
        const sin = Math.sin((mid * Math.PI) / 180);
        job.h = sin > 0.15 ? 'left' : sin < -0.15 ? 'right' : 'center';
        const cos = Math.cos((mid * Math.PI) / 180);
        job.v = job.h === 'center' ? (cos > 0 ? 'bottom' : 'top') : 'middle';
      }
      labels.push(job);
    }
  });
}

// ---------------------------------------------------------------------------------------------
// Radar

function drawRadar(d: Draw, model: ChartModel, outer: Box, labels: LabelJob[]): void {
  const plot = model.plots.find((p) => p.type === 'radar')!;
  const axes = new Map(model.axes.map((a) => [a.id, a]));
  const catAxis = axes.get(plot.axisIds[0] ?? '') ?? syntheticAxis('#rc', 'cat', 'b', model.textStyle);
  const valAxis = axes.get(plot.axisIds[1] ?? '') ?? syntheticAxis('#rv', 'val', 'l', model.textStyle);
  const first = plot.series[0];
  const n = Math.max(0, ...plot.series.map((s) => Math.max(s.values.length, s.categories.length)));
  if (n === 0) return;
  const cats = Array.from({ length: n }, (_, i) => first?.categories[i] ?? String(i + 1));
  const st = catAxis.textStyle;
  const lh = d.lineHeight(st);
  const showCats = !catAxis.deleted && catAxis.tickLblPos !== 'none';
  const labelW = showCats ? Math.max(0, ...cats.map((c) => d.width(c, st))) : 0;
  const radius = Math.max(4, Math.min(outer.w / 2 - labelW - 6, outer.h / 2 - (showCats ? lh * 1.5 : 0) - 6));
  const cx = outer.x + outer.w / 2;
  const cy = outer.y + outer.h / 2;

  let lo = Number.POSITIVE_INFINITY;
  let hi = Number.NEGATIVE_INFINITY;
  for (const s of plot.series) for (const v of s.values) if (v !== null && v !== undefined) {
    lo = Math.min(lo, v);
    hi = Math.max(hi, v);
  }
  const vst = valAxis.textStyle;
  const scale = niceScale(lo, hi, valAxis, radius, false, () => d.lineHeight(vst) * 1.6);
  const rOf = (v: number): number => (Math.max(scale.min, Math.min(scale.max, v)) - scale.min) / (scale.max - scale.min || 1) * radius;
  const angle = (i: number): number => (i * 360) / n;

  // Web: rings at each major tick, spokes per category.
  const ring = valAxis.majorGridlines ?? { paint: { color: 'D9D9D9', opacity: 1 }, width: 0.75 };
  for (const v of ticks(scale)) {
    const r = rOf(v);
    if (r <= 0) continue;
    const pts = Array.from({ length: n }, (_, i) => polar(cx, cy, r, angle(i)));
    d.path([...polyline(pts), { op: 'Z' }], null, ring);
  }
  const spoke = catAxis.majorGridlines ?? ring;
  for (let i = 0; i < n; i++) {
    const p = polar(cx, cy, radius, angle(i));
    d.segment(cx, cy, p.x, p.y, spoke);
  }
  if (!valAxis.deleted) {
    d.segment(cx, cy, cx, cy - radius, valAxis.line);
    if (valAxis.tickLblPos !== 'none') {
      const fmt = valAxis.numFmt && !valAxis.numFmt.sourceLinked ? valAxis.numFmt.code : (first?.formatCode ?? 'General');
      for (const v of ticks(scale)) d.textBlock([formatNumber(v, fmt)], cx - LABEL_GAP, cy - rOf(v), vst, 'right', 'middle');
    }
  }
  if (showCats) {
    cats.forEach((c, i) => {
      const a = angle(i);
      const sin = Math.sin((a * Math.PI) / 180);
      const cos = Math.cos((a * Math.PI) / 180);
      const p = polar(cx, cy, radius + 4 + (Math.abs(cos) * d.lineHeight(vst)) / 2, a);
      const h = sin > 0.1 ? 'left' : sin < -0.1 ? 'right' : 'center';
      const v = h !== 'center' ? 'middle' : cos > 0 ? 'bottom' : 'top';
      d.textBlock([c], p.x, p.y, st, h, v);
    });
  }

  const filled = plot.radarStyle === 'filled';
  for (const s of plot.series) {
    const pts = Array.from({ length: n }, (_, i) => polar(cx, cy, rOf(s.values[i] ?? 0), angle(i)));
    d.path([...polyline(pts), { op: 'Z' }], filled ? s.fill : null, filled ? (s.line.paint ? s.line : null) : s.line, { lineJoin: 'round' });
    if (!filled) pts.forEach((p) => drawMarker(d, s.marker, p.x, p.y));
    pts.forEach((p, i) => {
      const flags = labelFlags(s, i);
      const v = s.values[i];
      if (flags && v !== null && v !== undefined) labels.push(pointLabel(labelText(flags, s, i, v, undefined), p.x, p.y, flags.position ?? 't', flags));
    });
  }
}

// ---------------------------------------------------------------------------------------------
// The chart

function manualBox(layout: ManualLayout, frame: Box, dflt: Box): Box {
  const x = layout.x === undefined ? dflt.x : layout.xMode === 'edge' ? frame.x + layout.x * frame.w : dflt.x + layout.x * frame.w;
  const y = layout.y === undefined ? dflt.y : layout.yMode === 'edge' ? frame.y + layout.y * frame.h : dflt.y + layout.y * frame.h;
  const w = layout.w === undefined ? dflt.x + dflt.w - x : layout.wMode === 'edge' ? frame.x + layout.w * frame.w - x : layout.w * frame.w;
  const h = layout.h === undefined ? dflt.y + dflt.h - y : layout.hMode === 'edge' ? frame.y + layout.h * frame.h - y : layout.h * frame.h;
  return { x, y, w: Math.max(w, 1), h: Math.max(h, 1) };
}

/** Draw an already-read chart model into `frame`. */
export function drawChart(model: ChartModel, frame: ChartFrame, options: RenderChartOptions): DisplayItem[] {
  const d = new Draw(options.fonts, options.defaultFont ?? 'Calibri');
  if (!(frame.w > 0 && frame.h > 0)) return d.items;
  const ca = model.chartArea;
  d.path(rectPath(frame, ca.roundedCorners ? Math.min(frame.w, frame.h) * 0.06 : 0), ca.fill, ca.line);

  const pad = Math.max(2, Math.min(7, Math.min(frame.w, frame.h) * 0.03));
  let area: Box = { x: frame.x + pad, y: frame.y + pad, w: frame.w - 2 * pad, h: frame.h - 2 * pad };
  let title: TitlePlan | undefined;
  if (model.title && model.title.text.trim()) {
    title = planTitle(d, model, frame, area);
    area = title.remaining;
  }
  const legend = model.legend ? planLegend(d, model, frame, area) : undefined;
  if (legend) area = legend.remaining;

  const layout = model.plotArea.layout;
  const placed = layout ? manualBox(layout, frame, area) : area;
  const outer = layout && layout.target === 'outer' ? placed : area;
  const manualInner = layout && layout.target === 'inner' ? placed : undefined;

  const labels: LabelJob[] = [];
  if (model.plots.some((p) => AXIS_TYPES.has(p.type))) drawAxisChart(d, model, outer, manualInner, labels);
  else {
    const box = manualInner ?? outer;
    d.path(rectPath(box), model.plotArea.fill, model.plotArea.line);
    if (model.plots.some((p) => p.type === 'pie' || p.type === 'doughnut')) drawPie(d, model, box, labels);
    else if (model.plots.some((p) => p.type === 'radar')) drawRadar(d, model, box, labels);
  }
  drawLabels(d, labels);
  title?.draw();
  if (legend) drawLegend(d, model, legend);
  return d.items;
}
