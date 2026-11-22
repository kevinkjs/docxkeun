/**
 * Reads a chart part (c:chartSpace) into a plain model: plots and their series with the cached
 * values (there is no spreadsheet engine), axes, title, legend, data labels and resolved colours.
 * Automatic formatting (series palette, default line widths) is resolved here so the renderer and
 * any other consumer see the colours Office would show.
 */
import { attr, elements, firstChild, textContent, type XmlElement } from '../xml/dom.js';
import { adjustLightness, fillOf, lineOf, lookupScheme, schemeVariant, type LineProps, type Paint, type SchemeColorFn } from './color.js';
import { formatNumber } from './format.js';

export type PlotType = 'bar' | 'line' | 'area' | 'pie' | 'doughnut' | 'scatter' | 'radar' | 'bubble' | 'surface';
export type Grouping = 'clustered' | 'stacked' | 'percentStacked' | 'standard';
export type MarkerSymbol = 'circle' | 'square' | 'diamond' | 'triangle' | 'x' | 'star' | 'plus' | 'dash' | 'dot' | 'picture' | 'none';
export type AxisPos = 'b' | 't' | 'l' | 'r';
export type TickMark = 'cross' | 'in' | 'out' | 'none';

export interface TextStyle {
  /** Points. */
  size: number;
  bold: boolean;
  italic: boolean;
  color: string;
  /** Typeface; undefined = the chart's default font. */
  font?: string;
  /** Rotation in degrees (a:bodyPr rot), clockwise. */
  rotation?: number;
}

/** A resolved line: `paint: null` means no line. Width in points. */
export interface LineStyle {
  paint: Paint | null;
  width: number;
  dash?: string;
  cap?: 'butt' | 'round' | 'square';
  join?: 'miter' | 'round' | 'bevel';
}

/** c:layout/c:manualLayout: fractions of the chart area. */
export interface ManualLayout {
  target: 'inner' | 'outer';
  xMode: 'edge' | 'factor';
  yMode: 'edge' | 'factor';
  wMode: 'edge' | 'factor';
  hMode: 'edge' | 'factor';
  x?: number;
  y?: number;
  w?: number;
  h?: number;
}

export interface NumFmt {
  code: string;
  sourceLinked: boolean;
}

/** A title (chart or axis). */
export interface ChartText {
  text: string;
  /** Paragraphs (lines) of the text. */
  lines: string[];
  style: TextStyle;
  overlay: boolean;
  /** True when the text was not written but derived (the single series' name). */
  auto: boolean;
  layout?: ManualLayout;
  fill?: Paint | null;
  line?: LineStyle;
}

export interface DataLabelFlags {
  showVal: boolean;
  showPercent: boolean;
  showCatName: boolean;
  showSerName: boolean;
  showLegendKey: boolean;
  showBubbleSize: boolean;
  separator?: string;
  /** c:dLblPos: outEnd, inEnd, ctr, inBase, bestFit, t, b, l, r. */
  position?: string;
  numFmt?: NumFmt;
  style: TextStyle;
  fill?: Paint | null;
  /** c:delete: the labels are removed. */
  deleted: boolean;
}

export interface DataLabels extends DataLabelFlags {
  /** Per-point overrides (c:dLbl) by point index. */
  points: Map<number, DataLabelFlags>;
}

export interface MarkerStyle {
  symbol: MarkerSymbol;
  /** Points (c:size is 2..72, in points). */
  size: number;
  fill: Paint | null;
  line: Paint | null;
  lineWidth: number;
}

/** Explicit formatting of one data point (c:dPt). */
export interface ChartPoint {
  index: number;
  fill?: Paint | null;
  line?: LineStyle;
  explosion?: number;
  marker?: MarkerStyle;
  invertIfNegative?: boolean;
}

export interface ChartSeries {
  /** c:idx: drives the automatic colour. */
  index: number;
  /** c:order: drawing and legend order. */
  order: number;
  name: string;
  /** Category labels by point index (formatted when the categories are numbers or dates). */
  categories: string[];
  /** Outer levels of multi-level categories, outermost first; each sparse by point index. */
  categoryGroups?: Array<Array<string | undefined>>;
  /** Numeric categories (scatter x values, date axes), when the categories are numbers. */
  categoryNumbers?: Array<number | null>;
  categoryFormat?: string;
  values: Array<number | null>;
  /** Format code of the cached values (General when none). */
  formatCode: string;
  bubbleSizes?: Array<number | null>;
  /** Main colour: the fill for bars/areas, the line for lines (explicit or automatic). */
  color: string;
  /** Effective fill (bars, areas, slices, bubbles). */
  fill: Paint | null;
  /** Effective line (the series line, or the outline of bars and slices). */
  line: LineStyle;
  marker: MarkerStyle;
  smooth: boolean;
  /** Pie explosion as a percentage of the radius. */
  explosion: number;
  invertIfNegative: boolean;
  points: ChartPoint[];
  /** Automatic colour per point, when the plot varies colours by point (pies). */
  pointColors?: string[];
  dataLabels?: DataLabels;
}

export interface ChartPlot {
  type: PlotType;
  /** The element it came from, e.g. c:bar3DChart. */
  element: string;
  is3D: boolean;
  barDir: 'col' | 'bar';
  grouping: Grouping;
  varyColors: boolean;
  /** Bar charts: gap between clusters and overlap of bars, in percent of a bar width. */
  gapWidth: number;
  overlap: number;
  /** Pie: angle of the first slice in degrees clockwise from 12 o'clock. */
  firstSliceAng: number;
  /** Doughnut hole as a percentage of the diameter. */
  holeSize: number;
  scatterStyle?: string;
  radarStyle?: string;
  /** c:lineChart/c:marker: series without an explicit marker get automatic markers. */
  markers: boolean;
  /** Bubble charts: c:bubbleScale percent. */
  bubbleScale: number;
  axisIds: string[];
  series: ChartSeries[];
  dataLabels?: DataLabels;
}

export interface ChartAxis {
  id: string;
  kind: 'cat' | 'val' | 'date' | 'ser';
  position: AxisPos;
  deleted: boolean;
  orientation: 'minMax' | 'maxMin';
  min?: number;
  max?: number;
  majorUnit?: number;
  minorUnit?: number;
  logBase?: number;
  numFmt?: NumFmt;
  title?: ChartText;
  majorGridlines?: LineStyle;
  minorGridlines?: LineStyle;
  majorTickMark: TickMark;
  minorTickMark: TickMark;
  tickLblPos: 'nextTo' | 'low' | 'high' | 'none';
  crossAx?: string;
  crosses: 'autoZero' | 'min' | 'max';
  crossesAt?: number;
  crossBetween?: 'between' | 'midCat';
  line: LineStyle;
  textStyle: TextStyle;
  tickLblSkip?: number;
  /** Display units divisor (c:dispUnits: thousands → 1000). */
  dispUnits?: number;
}

export interface ChartLegend {
  position: 'r' | 'l' | 't' | 'b' | 'tr';
  overlay: boolean;
  layout?: ManualLayout;
  /** Legend entries removed with c:legendEntry/c:delete. */
  deletedEntries: number[];
  style: TextStyle;
  fill?: Paint | null;
  line?: LineStyle;
}

export interface ChartModel {
  title?: ChartText;
  plots: ChartPlot[];
  axes: ChartAxis[];
  legend?: ChartLegend;
  dispBlanksAs: 'gap' | 'zero' | 'span';
  /** c:style (1–48): picks the automatic palette (2 = the accent colours). */
  style: number;
  /** Chart-wide default text (c:chartSpace/c:txPr). */
  textStyle: TextStyle;
  chartArea: { fill: Paint | null; line: LineStyle; roundedCorners: boolean };
  plotArea: { fill: Paint | null; line: LineStyle; layout?: ManualLayout };
}

export interface ReadChartOptions {
  /** Theme colours; Office's default theme is used for names it does not know. */
  schemeColor?: SchemeColorFn;
}

// ---------------------------------------------------------------------------------------------
// Small readers

/** CT_Boolean: an element without `val` means true. */
function bool(parent: XmlElement | undefined, name: string, dflt: boolean): boolean {
  const node = firstChild(parent, name);
  if (!node) return dflt;
  const v = attr(node, 'val');
  return v === undefined || v === '1' || v === 'true';
}

function num(parent: XmlElement | undefined, name: string): number | undefined {
  const v = attr(firstChild(parent, name), 'val');
  if (v === undefined) return undefined;
  const n = Number.parseFloat(v);
  return Number.isFinite(n) ? n : undefined;
}

function str(parent: XmlElement | undefined, name: string): string | undefined {
  return attr(firstChild(parent, name), 'val');
}

/** Elements may sit inside mc:AlternateContent; take the fallback (it uses the c: namespace). */
function childOrFallback(parent: XmlElement | undefined, name: string): XmlElement | undefined {
  const direct = firstChild(parent, name);
  if (direct || !parent) return direct;
  for (const alt of elements(parent, 'mc:AlternateContent')) {
    for (const branch of [...elements(alt, 'mc:Choice'), ...elements(alt, 'mc:Fallback')]) {
      const found = firstChild(branch, name);
      if (found) return found;
    }
  }
  return undefined;
}

function readLayout(parent: XmlElement | undefined): ManualLayout | undefined {
  const m = firstChild(firstChild(parent, 'c:layout'), 'c:manualLayout');
  if (!m) return undefined;
  const mode = (name: string): 'edge' | 'factor' => (str(m, name) === 'edge' ? 'edge' : 'factor');
  const layout: ManualLayout = {
    target: str(m, 'c:layoutTarget') === 'inner' ? 'inner' : 'outer',
    xMode: mode('c:xMode'),
    yMode: mode('c:yMode'),
    wMode: mode('c:wMode'),
    hMode: mode('c:hMode'),
  };
  const x = num(m, 'c:x');
  const y = num(m, 'c:y');
  const w = num(m, 'c:w');
  const h = num(m, 'c:h');
  if (x !== undefined) layout.x = x;
  if (y !== undefined) layout.y = y;
  if (w !== undefined) layout.w = w;
  if (h !== undefined) layout.h = h;
  return layout.x === undefined && layout.y === undefined && layout.w === undefined && layout.h === undefined ? undefined : layout;
}

function resolveLine(props: LineProps, dflt: LineStyle): LineStyle {
  const out: LineStyle = { paint: props.paint === undefined ? dflt.paint : props.paint, width: props.width ?? dflt.width };
  const dash = props.dash ?? dflt.dash;
  if (dash) out.dash = dash;
  const cap = props.cap ?? dflt.cap;
  if (cap) out.cap = cap;
  const join = props.join ?? dflt.join;
  if (join) out.join = join;
  return out;
}

/** Character properties (a:defRPr / a:rPr) over a base style. */
function applyRunProps(base: TextStyle, rPr: XmlElement | undefined, scheme: SchemeColorFn | undefined): TextStyle {
  if (!rPr) return base;
  const out: TextStyle = { ...base };
  const sz = attr(rPr, 'sz');
  if (sz !== undefined && Number.isFinite(Number(sz))) out.size = Number(sz) / 100;
  const b = attr(rPr, 'b');
  if (b !== undefined) out.bold = b === '1' || b === 'true';
  const i = attr(rPr, 'i');
  if (i !== undefined) out.italic = i === '1' || i === 'true';
  const fill = fillOf(rPr, scheme);
  if (fill) out.color = fill.color;
  const typeface = attr(firstChild(rPr, 'a:latin'), 'typeface');
  // "+mn-lt"/"+mj-lt" are theme font references; the chart's default font stands in for them.
  if (typeface && !typeface.startsWith('+')) out.font = typeface;
  return out;
}

/** Text properties of a c:txPr or c:rich: body rotation plus the first paragraph's defaults. */
function readTextProps(txPr: XmlElement | undefined, base: TextStyle, scheme: SchemeColorFn | undefined): TextStyle {
  if (!txPr) return base;
  let style = { ...base };
  const bodyPr = firstChild(txPr, 'a:bodyPr');
  const rot = attr(bodyPr, 'rot');
  const vert = attr(bodyPr, 'vert');
  if (rot !== undefined && Number.isFinite(Number(rot))) style.rotation = Number(rot) / 60000;
  else if (vert === 'vert270') style.rotation = -90;
  else if (vert === 'vert' || vert === 'eaVert') style.rotation = 90;
  const p = firstChild(txPr, 'a:p');
  style = applyRunProps(style, firstChild(firstChild(p, 'a:pPr'), 'a:defRPr'), scheme);
  return style;
}

function richText(rich: XmlElement): string[] {
  const lines: string[] = [];
  for (const p of elements(rich, 'a:p')) {
    let line = '';
    for (const c of elements(p)) {
      if (c.name === 'a:r' || c.name === 'a:fld') line += textContent(firstChild(c, 'a:t') ?? c);
      else if (c.name === 'a:br') {
        lines.push(line);
        line = '';
      }
    }
    lines.push(line);
  }
  return lines;
}

// ---------------------------------------------------------------------------------------------
// Cached data

interface PointList<T> {
  count: number;
  values: T[];
  formatCode?: string;
}

function readPoints(cache: XmlElement | undefined): PointList<string | undefined> & { formats: Map<number, string> } {
  const out = { count: 0, values: [] as Array<string | undefined>, formats: new Map<number, string>(), formatCode: undefined as string | undefined };
  if (!cache) return out;
  const fc = firstChild(cache, 'c:formatCode');
  if (fc) out.formatCode = textContent(fc);
  const count = num(cache, 'c:ptCount');
  let max = -1;
  for (const pt of elements(cache, 'c:pt')) {
    const idx = Number.parseInt(attr(pt, 'idx') ?? '-1', 10);
    if (!(idx >= 0)) continue;
    out.values[idx] = textContent(firstChild(pt, 'c:v') ?? pt);
    const f = attr(pt, 'formatCode');
    if (f) out.formats.set(idx, f);
    max = Math.max(max, idx);
  }
  out.count = Math.max(count ?? 0, max + 1);
  return out;
}

interface DataRef {
  kind: 'num' | 'str' | 'multi' | 'none';
  count: number;
  numbers: Array<number | null>;
  strings: Array<string | undefined>;
  formatCode?: string;
  /** Multi-level categories: levels from the innermost (leaf) outwards. */
  levels?: Array<Array<string | undefined>>;
}

/** Read c:cat / c:val / c:xVal / c:yVal / c:tx / c:bubbleSize from whichever cache or literal it carries. */
function readData(node: XmlElement | undefined): DataRef {
  const empty: DataRef = { kind: 'none', count: 0, numbers: [], strings: [] };
  if (!node) return empty;
  const numCache = firstChild(firstChild(node, 'c:numRef'), 'c:numCache') ?? firstChild(node, 'c:numLit');
  if (numCache) {
    const p = readPoints(numCache);
    const numbers: Array<number | null> = [];
    for (let i = 0; i < p.count; i++) {
      const v = p.values[i];
      const n = v === undefined || v.trim() === '' ? Number.NaN : Number(v);
      numbers.push(Number.isFinite(n) ? n : null);
    }
    const ref: DataRef = { kind: 'num', count: p.count, numbers, strings: p.values };
    if (p.formatCode !== undefined) ref.formatCode = p.formatCode;
    return ref;
  }
  const strCache = firstChild(firstChild(node, 'c:strRef'), 'c:strCache') ?? firstChild(node, 'c:strLit');
  if (strCache) {
    const p = readPoints(strCache);
    return { kind: 'str', count: p.count, numbers: p.values.map((v) => (v !== undefined && v.trim() !== '' && Number.isFinite(Number(v)) ? Number(v) : null)), strings: p.values };
  }
  const multi = firstChild(firstChild(node, 'c:multiLvlStrRef'), 'c:multiLvlStrCache');
  if (multi) {
    const count = num(multi, 'c:ptCount') ?? 0;
    const levels = elements(multi, 'c:lvl').map((lvl) => readPoints(lvl).values);
    const leaf = levels[0] ?? [];
    return { kind: 'multi', count: Math.max(count, leaf.length), numbers: [], strings: leaf, levels };
  }
  // c:tx can hold the text directly.
  const v = firstChild(node, 'c:v');
  if (v) return { kind: 'str', count: 1, numbers: [], strings: [textContent(v)] };
  return empty;
}

// ---------------------------------------------------------------------------------------------
// Automatic colours

/** The chart colour style's variations for series 7+ (Office's colors1.xml, cycling the accents). */
const VARIATIONS: Array<[number, number]> = [
  [1, 0],
  [0.6, 0],
  [0.8, 0.2],
  [0.8, 0],
  [0.6, 0.4],
  [0.5, 0],
  [0.7, 0.3],
  [0.7, 0],
  [0.5, 0.5],
];

/** Automatic colour of series (or point) `i` of `count`, for chart style `style`. */
export function autoColor(i: number, count: number, style: number, scheme: SchemeColorFn | undefined): string {
  const column = style >= 1 && style <= 48 ? (style - 1) % 8 : 1;
  if (column === 1) {
    const [lumMod, lumOff] = VARIATIONS[Math.floor(i / 6) % VARIATIONS.length]!;
    return schemeVariant(`accent${(i % 6) + 1}`, scheme, lumMod, lumOff);
  }
  // Grayscale (column 0) and single-accent styles spread shades from dark to light.
  const base = column === 0 ? '808080' : lookupScheme(`accent${column - 1}`, scheme);
  if (count <= 1) return base;
  const t = i / (count - 1) - 0.5; // -0.5 … 0.5
  return t < 0 ? adjustLightness(base, 1 + t * 0.9) : adjustLightness(base, 1 - t * 0.9, t * 0.9);
}

// ---------------------------------------------------------------------------------------------
// Data labels, markers, series

function readLabelFlags(node: XmlElement, base: TextStyle, scheme: SchemeColorFn | undefined, inherit?: DataLabelFlags): DataLabelFlags {
  const b = (name: string, key: keyof DataLabelFlags): boolean => bool(node, name, inherit ? Boolean(inherit[key]) : false);
  const flags: DataLabelFlags = {
    showVal: b('c:showVal', 'showVal'),
    showPercent: b('c:showPercent', 'showPercent'),
    showCatName: b('c:showCatName', 'showCatName'),
    showSerName: b('c:showSerName', 'showSerName'),
    showLegendKey: b('c:showLegendKey', 'showLegendKey'),
    showBubbleSize: b('c:showBubbleSize', 'showBubbleSize'),
    style: readTextProps(firstChild(node, 'c:txPr'), inherit?.style ?? base, scheme),
    deleted: bool(node, 'c:delete', false),
  };
  const sep = firstChild(node, 'c:separator');
  if (sep) flags.separator = textContent(sep);
  else if (inherit?.separator !== undefined) flags.separator = inherit.separator;
  const pos = str(node, 'c:dLblPos') ?? inherit?.position;
  if (pos) flags.position = pos;
  const nf = firstChild(node, 'c:numFmt');
  if (nf) flags.numFmt = { code: attr(nf, 'formatCode') ?? 'General', sourceLinked: attr(nf, 'sourceLinked') === '1' };
  else if (inherit?.numFmt) flags.numFmt = inherit.numFmt;
  const fill = fillOf(firstChild(node, 'c:spPr'), scheme);
  if (fill !== undefined) flags.fill = fill;
  else if (inherit?.fill !== undefined) flags.fill = inherit.fill;
  return flags;
}

function readDataLabels(node: XmlElement | undefined, base: TextStyle, scheme: SchemeColorFn | undefined): DataLabels | undefined {
  if (!node) return undefined;
  const flags = readLabelFlags(node, base, scheme);
  const points = new Map<number, DataLabelFlags>();
  for (const d of elements(node, 'c:dLbl')) {
    const idx = num(d, 'c:idx');
    if (idx !== undefined) points.set(idx, readLabelFlags(d, base, scheme, flags));
  }
  return { ...flags, points };
}

const SYMBOLS = new Set<MarkerSymbol>(['circle', 'square', 'diamond', 'triangle', 'x', 'star', 'plus', 'dash', 'dot', 'picture', 'none']);
/** Excel's automatic marker sequence. */
const AUTO_SYMBOLS: MarkerSymbol[] = ['diamond', 'square', 'triangle', 'x', 'star', 'circle', 'plus'];

function readMarker(node: XmlElement | undefined, auto: MarkerStyle, scheme: SchemeColorFn | undefined): MarkerStyle {
  if (!node) return auto;
  const sym = str(node, 'c:symbol') as MarkerSymbol | undefined;
  const spPr = firstChild(node, 'c:spPr');
  const fill = fillOf(spPr, scheme);
  const ln = lineOf(spPr, scheme);
  return {
    symbol: sym && SYMBOLS.has(sym) ? sym : auto.symbol,
    size: num(node, 'c:size') ?? auto.size,
    fill: fill === undefined ? auto.fill : fill,
    line: ln.paint === undefined ? auto.line : ln.paint,
    lineWidth: ln.width ?? auto.lineWidth,
  };
}

interface PlotContext {
  type: PlotType;
  varyColors: boolean;
  /** Series and point counts, to spread monochrome palettes. */
  seriesCount: number;
  markers: boolean;
  scatterStyle?: string;
  radarStyle?: string;
  style: number;
  scheme: SchemeColorFn | undefined;
  text: TextStyle;
}

function readSeries(ser: XmlElement, ctx: PlotContext, position: number): ChartSeries {
  const { scheme, type } = ctx;
  const index = num(ser, 'c:idx') ?? position;
  const order = num(ser, 'c:order') ?? index;

  const tx = readData(firstChild(ser, 'c:tx'));
  const name = tx.strings[0] ?? tx.numbers[0]?.toString() ?? `Series ${index + 1}`;

  const xy = type === 'scatter' || type === 'bubble';
  const catRef = readData(firstChild(ser, xy ? 'c:xVal' : 'c:cat'));
  const valRef = readData(firstChild(ser, xy ? 'c:yVal' : 'c:val'));
  const values = valRef.numbers.slice();
  const count = Math.max(valRef.count, catRef.count);
  while (values.length < valRef.count) values.push(null);

  const categories: string[] = [];
  let categoryNumbers: Array<number | null> | undefined;
  let categoryFormat: string | undefined;
  if (catRef.kind === 'num') {
    categoryNumbers = catRef.numbers;
    categoryFormat = catRef.formatCode;
    for (let i = 0; i < catRef.count; i++) {
      const n = catRef.numbers[i];
      categories.push(n === null || n === undefined ? '' : formatNumber(n, catRef.formatCode));
    }
  } else if (catRef.kind !== 'none') {
    for (let i = 0; i < catRef.count; i++) categories.push(catRef.strings[i] ?? '');
  } else {
    // No categories: Excel numbers the points from 1.
    for (let i = 0; i < count; i++) categories.push(String(i + 1));
  }

  // Automatic colour: per series, or per point when the plot varies colours.
  const color = autoColor(index, ctx.seriesCount, ctx.style, scheme);
  const spPr = firstChild(ser, 'c:spPr');
  const explicitFill = fillOf(spPr, scheme);
  const lineProps = lineOf(spPr, scheme);
  const lineLike = type === 'line' || type === 'scatter' || (type === 'radar' && ctx.radarStyle !== 'filled');

  let main = color;
  let fill: Paint | null;
  let line: LineStyle;
  if (lineLike) {
    line = resolveLine(lineProps, { paint: { color, opacity: 1 }, width: 2.25, cap: 'round', join: 'round' });
    if (line.paint) main = line.paint.color;
    fill = explicitFill === undefined ? { color: main, opacity: 1 } : explicitFill;
  } else {
    fill = explicitFill === undefined ? { color, opacity: 1 } : explicitFill;
    if (fill) main = fill.color;
    // Slices get a thin white edge by default, as Office draws them; bars and areas have none.
    const dfltPaint = type === 'pie' || type === 'doughnut' ? { color: lookupScheme('lt1', scheme), opacity: 1 } : null;
    line = resolveLine(lineProps, { paint: dfltPaint, width: 0.75 });
  }

  // Markers: line/scatter/radar series; automatic when the plot asks for markers and the series is silent.
  const markerNode = firstChild(ser, 'c:marker');
  let autoSymbol: MarkerSymbol = 'none';
  if (type === 'line' && ctx.markers) autoSymbol = AUTO_SYMBOLS[index % AUTO_SYMBOLS.length]!;
  if (type === 'scatter' && !['line', 'smooth'].includes(ctx.scatterStyle ?? '')) autoSymbol = AUTO_SYMBOLS[index % AUTO_SYMBOLS.length]!;
  if (type === 'radar' && ctx.radarStyle === 'marker') autoSymbol = AUTO_SYMBOLS[index % AUTO_SYMBOLS.length]!;
  const marker = readMarker(markerNode, { symbol: autoSymbol, size: 5, fill: { color: main, opacity: 1 }, line: { color: main, opacity: 1 }, lineWidth: 0.75 }, scheme);

  const series: ChartSeries = {
    index,
    order,
    name,
    categories,
    values,
    formatCode: valRef.formatCode ?? 'General',
    color: main,
    fill,
    line,
    marker,
    smooth: bool(ser, 'c:smooth', type === 'scatter' && (ctx.scatterStyle === 'smooth' || ctx.scatterStyle === 'smoothMarker')),
    explosion: num(ser, 'c:explosion') ?? 0,
    invertIfNegative: bool(ser, 'c:invertIfNegative', false),
    points: [],
  };
  // A scatter series with style "marker" has no line unless it sets one.
  if (type === 'scatter' && ctx.scatterStyle === 'marker' && lineProps.paint === undefined) series.line = { ...series.line, paint: null };
  if (categoryNumbers) series.categoryNumbers = categoryNumbers;
  if (categoryFormat !== undefined) series.categoryFormat = categoryFormat;
  if (catRef.levels && catRef.levels.length > 1) series.categoryGroups = catRef.levels.slice(1).reverse();

  const bubble = readData(firstChild(ser, 'c:bubbleSize'));
  if (type === 'bubble') series.bubbleSizes = bubble.kind === 'none' ? values.map(() => 1) : bubble.numbers;

  // Explicit series formatting wins over per-point automatic colours.
  if (ctx.varyColors && explicitFill === undefined) {
    const n = Math.max(values.length, categories.length);
    series.pointColors = Array.from({ length: n }, (_, i) => autoColor(i, n, ctx.style, scheme));
  }

  for (const dPt of elements(ser, 'c:dPt')) {
    const idx = num(dPt, 'c:idx');
    if (idx === undefined) continue;
    const pSpPr = firstChild(dPt, 'c:spPr');
    const point: ChartPoint = { index: idx };
    const pf = fillOf(pSpPr, scheme);
    if (pf !== undefined) point.fill = pf;
    const pl = lineOf(pSpPr, scheme);
    if (pl.paint !== undefined || pl.width !== undefined) point.line = resolveLine(pl, series.line);
    const ex = num(dPt, 'c:explosion');
    if (ex !== undefined) point.explosion = ex;
    const mk = firstChild(dPt, 'c:marker');
    if (mk) point.marker = readMarker(mk, series.marker, scheme);
    if (firstChild(dPt, 'c:invertIfNegative')) point.invertIfNegative = bool(dPt, 'c:invertIfNegative', false);
    series.points.push(point);
  }

  const dl = readDataLabels(firstChild(ser, 'c:dLbls'), ctx.text, scheme);
  if (dl) series.dataLabels = dl;
  return series;
}

// ---------------------------------------------------------------------------------------------
// Plots, axes, titles, legend

const PLOT_TYPES: Record<string, PlotType> = {
  'c:barChart': 'bar',
  'c:bar3DChart': 'bar',
  'c:lineChart': 'line',
  'c:line3DChart': 'line',
  'c:stockChart': 'line',
  'c:areaChart': 'area',
  'c:area3DChart': 'area',
  'c:pieChart': 'pie',
  'c:pie3DChart': 'pie',
  'c:ofPieChart': 'pie',
  'c:doughnutChart': 'doughnut',
  'c:scatterChart': 'scatter',
  'c:radarChart': 'radar',
  'c:bubbleChart': 'bubble',
  'c:surfaceChart': 'surface',
  'c:surface3DChart': 'surface',
};

function readTitle(node: XmlElement | undefined, base: TextStyle, scheme: SchemeColorFn | undefined, fallbackText: string | undefined): ChartText | undefined {
  if (!node) return undefined;
  const tx = firstChild(node, 'c:tx');
  const rich = firstChild(tx, 'c:rich');
  let lines: string[];
  let auto = false;
  let style = readTextProps(firstChild(node, 'c:txPr'), base, scheme);
  if (rich) {
    lines = richText(rich);
    style = readTextProps(rich, style, scheme);
    // The first run's own properties win over the paragraph defaults.
    const firstRun = elements(rich, 'a:p').flatMap((p) => elements(p, 'a:r'))[0];
    style = applyRunProps(style, firstChild(firstRun, 'a:rPr'), scheme);
  } else if (tx) {
    const data = readData(tx);
    lines = [data.strings[0] ?? ''];
  } else {
    if (fallbackText === undefined) return undefined;
    lines = [fallbackText];
    auto = true;
  }
  const text = lines.join('\n');
  const title: ChartText = { text, lines, style, overlay: bool(node, 'c:overlay', false), auto };
  const layout = readLayout(node);
  if (layout) title.layout = layout;
  const spPr = firstChild(node, 'c:spPr');
  const fill = fillOf(spPr, scheme);
  if (fill !== undefined) title.fill = fill;
  const ln = lineOf(spPr, scheme);
  if (ln.paint) title.line = resolveLine(ln, { paint: null, width: 0.75 });
  return title;
}

function readAxis(node: XmlElement, base: TextStyle, scheme: SchemeColorFn | undefined): ChartAxis {
  const kind = node.name === 'c:valAx' ? 'val' : node.name === 'c:dateAx' ? 'date' : node.name === 'c:serAx' ? 'ser' : 'cat';
  const scaling = firstChild(node, 'c:scaling');
  const pos = str(node, 'c:axPos');
  const textStyle = readTextProps(firstChild(node, 'c:txPr'), base, scheme);
  const spPr = firstChild(node, 'c:spPr');
  const axis: ChartAxis = {
    id: str(node, 'c:axId') ?? '',
    kind,
    position: pos === 't' || pos === 'l' || pos === 'r' ? pos : 'b',
    deleted: bool(node, 'c:delete', false),
    orientation: str(scaling, 'c:orientation') === 'maxMin' ? 'maxMin' : 'minMax',
    majorTickMark: (str(node, 'c:majorTickMark') as TickMark | undefined) ?? 'out',
    minorTickMark: (str(node, 'c:minorTickMark') as TickMark | undefined) ?? 'none',
    tickLblPos: (str(node, 'c:tickLblPos') as ChartAxis['tickLblPos'] | undefined) ?? 'nextTo',
    crosses: (str(node, 'c:crosses') as ChartAxis['crosses'] | undefined) ?? 'autoZero',
    line: resolveLine(lineOf(spPr, scheme), { paint: { color: '868686', opacity: 1 }, width: 0.75 }),
    textStyle,
  };
  const min = num(scaling, 'c:min');
  const max = num(scaling, 'c:max');
  const logBase = num(scaling, 'c:logBase');
  if (min !== undefined) axis.min = min;
  if (max !== undefined) axis.max = max;
  if (logBase !== undefined && logBase >= 2) axis.logBase = logBase;
  const major = num(node, 'c:majorUnit');
  const minor = num(node, 'c:minorUnit');
  if (major !== undefined && major > 0) axis.majorUnit = major;
  if (minor !== undefined && minor > 0) axis.minorUnit = minor;
  const nf = firstChild(node, 'c:numFmt');
  if (nf) axis.numFmt = { code: attr(nf, 'formatCode') ?? 'General', sourceLinked: attr(nf, 'sourceLinked') === '1' };
  const crossAx = str(node, 'c:crossAx');
  if (crossAx) axis.crossAx = crossAx;
  const crossesAt = num(node, 'c:crossesAt');
  if (crossesAt !== undefined) axis.crossesAt = crossesAt;
  const cb = str(node, 'c:crossBetween');
  if (cb === 'between' || cb === 'midCat') axis.crossBetween = cb;
  const skip = num(node, 'c:tickLblSkip');
  if (skip !== undefined && skip >= 1) axis.tickLblSkip = Math.floor(skip);
  const gridDefault: LineStyle = { paint: { color: 'D9D9D9', opacity: 1 }, width: 0.75 };
  const major$ = firstChild(node, 'c:majorGridlines');
  if (major$) axis.majorGridlines = resolveLine(lineOf(firstChild(major$, 'c:spPr'), scheme), gridDefault);
  const minor$ = firstChild(node, 'c:minorGridlines');
  if (minor$) axis.minorGridlines = resolveLine(lineOf(firstChild(minor$, 'c:spPr'), scheme), { ...gridDefault, paint: { color: 'F2F2F2', opacity: 1 } });
  const titleBase: TextStyle = { ...textStyle, rotation: undefined, bold: true };
  // An axis title element without text shows Office's placeholder "Axis Title".
  const title = readTitle(firstChild(node, 'c:title'), titleBase, scheme, 'Axis Title');
  if (title) axis.title = title;
  const du = firstChild(node, 'c:dispUnits');
  if (du) {
    const custom = num(du, 'c:custUnit');
    const builtIn = str(du, 'c:builtInUnit');
    const units: Record<string, number> = { hundreds: 1e2, thousands: 1e3, tenThousands: 1e4, hundredThousands: 1e5, millions: 1e6, tenMillions: 1e7, hundredMillions: 1e8, billions: 1e9, trillions: 1e12 };
    const divisor = custom ?? (builtIn ? units[builtIn] : undefined);
    if (divisor && divisor > 0) axis.dispUnits = divisor;
  }
  return axis;
}

/** Read a c:chartSpace into a chart model. */
export function readChart(chartSpace: XmlElement, options: ReadChartOptions = {}): ChartModel {
  const scheme = options.schemeColor;
  const chart = firstChild(chartSpace, 'c:chart');
  const plotArea = firstChild(chart, 'c:plotArea');
  const styleValue = Number(attr(childOrFallback(chartSpace, 'c:style'), 'val') ?? '2');
  const style = Number.isFinite(styleValue) && styleValue >= 1 ? styleValue : 2;

  const baseText = readTextProps(firstChild(chartSpace, 'c:txPr'), { size: 10, bold: false, italic: false, color: '000000' }, scheme);

  // Plots in document order, each with its series sorted by c:order.
  const plots: ChartPlot[] = [];
  for (const node of plotArea ? elements(plotArea) : []) {
    const type = PLOT_TYPES[node.name];
    if (!type) continue;
    const barDir = str(node, 'c:barDir') === 'bar' ? 'bar' : 'col';
    const groupingRaw = str(node, 'c:grouping');
    let grouping: Grouping;
    if (type === 'bar') grouping = groupingRaw === 'stacked' || groupingRaw === 'percentStacked' || groupingRaw === 'standard' ? groupingRaw : 'clustered';
    else grouping = groupingRaw === 'stacked' || groupingRaw === 'percentStacked' ? groupingRaw : 'standard';
    const sers = elements(node, 'c:ser');
    // Office writes varyColors explicitly; for pies the visible default is one colour per slice.
    const varyDefault = type === 'pie' || type === 'doughnut';
    const varyColors = bool(node, 'c:varyColors', varyDefault) && (type === 'pie' || type === 'doughnut' || sers.length === 1);
    const scatterStyle = str(node, 'c:scatterStyle');
    const radarStyle = str(node, 'c:radarStyle');
    const plotLabels = readDataLabels(firstChild(node, 'c:dLbls'), baseText, scheme);
    const ctx: PlotContext = {
      type,
      varyColors,
      seriesCount: sers.length,
      markers: bool(node, 'c:marker', true),
      style,
      scheme,
      text: plotLabels?.style ?? baseText,
    };
    if (scatterStyle) ctx.scatterStyle = scatterStyle;
    if (radarStyle) ctx.radarStyle = radarStyle;
    const series = sers.map((s, i) => readSeries(s, ctx, i)).sort((a, b) => a.order - b.order);
    // Series without their own c:dLbls follow the plot's.
    if (plotLabels) for (const s of series) s.dataLabels ??= plotLabels;
    const plot: ChartPlot = {
      type,
      element: node.name,
      is3D: node.name.includes('3D'),
      barDir,
      grouping,
      varyColors,
      gapWidth: num(node, 'c:gapWidth') ?? 150,
      overlap: num(node, 'c:overlap') ?? (grouping === 'stacked' || grouping === 'percentStacked' ? 100 : 0),
      firstSliceAng: num(node, 'c:firstSliceAng') ?? 0,
      holeSize: num(node, 'c:holeSize') ?? (type === 'doughnut' ? 50 : 0),
      markers: ctx.markers,
      bubbleScale: num(node, 'c:bubbleScale') ?? 100,
      axisIds: elements(node, 'c:axId').map((a) => attr(a, 'val') ?? ''),
      series,
    };
    if (scatterStyle) plot.scatterStyle = scatterStyle;
    if (radarStyle) plot.radarStyle = radarStyle;
    if (plotLabels) plot.dataLabels = plotLabels;
    plots.push(plot);
  }

  const axes = plotArea ? elements(plotArea).filter((e) => ['c:catAx', 'c:valAx', 'c:dateAx', 'c:serAx'].includes(e.name)).map((a) => readAxis(a, baseText, scheme)) : [];

  // Title: explicit, or the single series' name when the chart shows exactly one series (Office's auto title).
  const allSeries = plots.flatMap((p) => p.series);
  const autoTitleDeleted = bool(chart, 'c:autoTitleDeleted', false);
  const titleNode = firstChild(chart, 'c:title');
  const titleBase: TextStyle = { ...baseText, size: Math.round(baseText.size * 1.4 * 2) / 2 };
  let title: ChartText | undefined;
  const single = allSeries.length === 1 ? allSeries[0]!.name : undefined;
  if (titleNode) title = readTitle(titleNode, titleBase, scheme, single ?? 'Chart Title');
  else if (!autoTitleDeleted && single !== undefined) title = { text: single, lines: [single], style: titleBase, overlay: false, auto: true };

  let legend: ChartLegend | undefined;
  const legendNode = firstChild(chart, 'c:legend');
  if (legendNode) {
    const pos = str(legendNode, 'c:legendPos');
    legend = {
      position: pos === 'l' || pos === 't' || pos === 'b' || pos === 'tr' ? pos : 'r',
      overlay: bool(legendNode, 'c:overlay', false),
      deletedEntries: elements(legendNode, 'c:legendEntry')
        .filter((e) => bool(e, 'c:delete', false))
        .map((e) => num(e, 'c:idx') ?? -1),
      style: readTextProps(firstChild(legendNode, 'c:txPr'), baseText, scheme),
    };
    const layout = readLayout(legendNode);
    if (layout) legend.layout = layout;
    const spPr = firstChild(legendNode, 'c:spPr');
    const fill = fillOf(spPr, scheme);
    if (fill !== undefined) legend.fill = fill;
    const ln = lineOf(spPr, scheme);
    if (ln.paint) legend.line = resolveLine(ln, { paint: null, width: 0.75 });
  }

  const blanks = str(chart, 'c:dispBlanksAs');
  const csSpPr = firstChild(chartSpace, 'c:spPr');
  const csFill = fillOf(csSpPr, scheme);
  const paSpPr = firstChild(plotArea, 'c:spPr');
  const paFill = fillOf(paSpPr, scheme);
  const model: ChartModel = {
    plots,
    axes,
    dispBlanksAs: blanks === 'zero' || blanks === 'span' ? blanks : 'gap',
    style,
    textStyle: baseText,
    chartArea: {
      fill: csFill === undefined ? { color: 'FFFFFF', opacity: 1 } : csFill,
      line: resolveLine(lineOf(csSpPr, scheme), { paint: csSpPr ? null : { color: '868686', opacity: 1 }, width: 0.75 }),
      roundedCorners: bool(chartSpace, 'c:roundedCorners', false),
    },
    plotArea: {
      fill: paFill === undefined ? null : paFill,
      line: resolveLine(lineOf(paSpPr, scheme), { paint: null, width: 0.75 }),
    },
  };
  const paLayout = readLayout(plotArea);
  if (paLayout) model.plotArea.layout = paLayout;
  if (title) model.title = title;
  if (legend) model.legend = legend;
  return model;
}

// ---------------------------------------------------------------------------------------------
// Description

function plotName(plot: ChartPlot): string {
  const dir = plot.barDir === 'bar' ? 'bar' : 'column';
  const g = plot.grouping;
  const prefix = g === 'stacked' ? 'stacked ' : g === 'percentStacked' ? '100% stacked ' : '';
  let name: string;
  switch (plot.type) {
    case 'bar':
      name = `${g === 'clustered' || g === 'standard' ? 'clustered ' : prefix}${dir}`;
      break;
    case 'line':
      name = plot.element === 'c:stockChart' ? 'stock' : `${prefix}line`;
      break;
    case 'area':
      name = `${prefix}area`;
      break;
    case 'scatter':
      name = 'scatter';
      break;
    default:
      name = plot.type;
  }
  return (plot.is3D ? '3-D ' : '') + name;
}

function capitalize(s: string): string {
  return s.charAt(0).toUpperCase() + s.slice(1);
}

const MAX_POINTS = 12;

/** One short plain-text description of the chart, for screen readers and AI tools. */
export function describeChart(model: ChartModel): string {
  const plots = model.plots;
  let kind: string;
  if (plots.length === 0) kind = 'Chart';
  else if (plots.length === 1) kind = `${capitalize(plotName(plots[0]!))} chart`;
  else kind = `Combination chart (${[...new Set(plots.map(plotName))].join(', ')})`;
  let out = kind;
  if (model.title && model.title.text.trim()) out += ` "${model.title.text.replace(/\s*\n\s*/g, ' ').trim()}"`;

  const seriesParts: string[] = [];
  for (const plot of plots) {
    for (const s of plot.series) {
      const fmt = (v: number): string => formatNumber(v, s.formatCode);
      const total = s.values.reduce<number>((a, v) => a + Math.abs(v ?? 0), 0);
      const pts: string[] = [];
      const n = s.values.length;
      for (let i = 0; i < Math.min(n, MAX_POINTS); i++) {
        const v = s.values[i];
        const shown = v === null || v === undefined ? 'blank' : fmt(v);
        if (plot.type === 'scatter' || plot.type === 'bubble') {
          const x = s.categoryNumbers?.[i];
          const xs = x === null || x === undefined ? (s.categories[i] ?? String(i + 1)) : formatNumber(x, s.categoryFormat);
          pts.push(`(${xs}, ${shown})`);
        } else {
          const cat = s.categories[i] ?? String(i + 1);
          const pct = (plot.type === 'pie' || plot.type === 'doughnut') && total > 0 && v !== null && v !== undefined ? ` (${Math.round((Math.abs(v) / total) * 100)}%)` : '';
          pts.push(`${cat} ${shown}${pct}`);
        }
      }
      if (n > MAX_POINTS) pts.push(`… ${n - MAX_POINTS} more`);
      seriesParts.push(`${s.name} (${pts.join(', ')})`);
    }
  }
  if (seriesParts.length) out += `: series ${seriesParts.join(', ')}`;

  const used = new Set(plots.flatMap((p) => p.axisIds));
  const axisTitles = model.axes.filter((a) => used.has(a.id) && !a.deleted && a.title && !a.title.auto).map((a) => `${a.position === 'b' || a.position === 't' ? 'horizontal' : 'vertical'} axis "${a.title!.text.replace(/\s*\n\s*/g, ' ')}"`);
  if (axisTitles.length) out += `; ${axisTitles.join(', ')}`;
  return out;
}
