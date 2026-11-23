/**
 * Writing chart parts: a new c:chartSpace built from a table of data (as Word writes a chart it has
 * just inserted), any chart's data read back as that table, and an existing chart patched with new
 * data while its formatting stays. Element order follows the DrawingML chart schema throughout.
 */
import {
  attr,
  childPath,
  cloneElement,
  el,
  elements,
  ensureChild,
  firstChild,
  insertOrdered,
  removeChildren,
  text,
  textContent,
  type XmlElement,
  type XmlNode,
} from '../xml/dom.js';
import { NS } from '../xml/ns.js';
import { plotName, readChart, VARIATIONS } from './model.js';
import { chartRanges, type ChartSeriesData } from './workbook.js';

export type { ChartSeriesData } from './workbook.js';

export type NewChartType = 'column' | 'bar' | 'line' | 'area' | 'pie' | 'doughnut' | 'scatter';
export type LegendPosition = 'bottom' | 'top' | 'left' | 'right' | 'none';

export interface NewChart {
  type: NewChartType;
  /** Column, bar and area charts only. */
  stacked?: 'stacked' | 'percent';
  title?: string;
  /** Category labels; for scatter charts the x values (numbers). */
  categories: Array<string | number>;
  series: ChartSeriesData[];
  /** Default bottom. */
  legend?: LegendPosition;
  axisTitles?: { category?: string; value?: string };
}

/** What to change in an existing chart; an empty string removes a title. */
export interface ChartChanges {
  title?: string;
  axisTitles?: { category?: string; value?: string };
  legend?: LegendPosition;
  categories?: Array<string | number>;
  series?: ChartSeriesData[];
}

/** A chart's data as one table: categories down the side, one column per series. */
export interface ChartTable {
  /** "clustered column", "stacked bar", "line", "pie", …; several plots joined with " + ". */
  kind: string;
  title?: string;
  axisTitles: { category?: string; value?: string };
  legend: LegendPosition;
  categories: Array<string | number>;
  series: ChartSeriesData[];
  /** Number format of the first series' values. */
  formatCode: string;
}

export class ChartDataError extends Error {
  override name = 'ChartDataError';
}

const STACKABLE = new Set<NewChartType>(['column', 'bar', 'area']);

/**
 * Check a table before it is written into a chart: every series has one value per category, all
 * numbers are finite, and the shape suits the chart type (when one is given).
 */
export function validateChartData(type: NewChartType | undefined, stacked: NewChart['stacked'], categories: Array<string | number>, series: ChartSeriesData[]): void {
  if (series.length === 0) throw new ChartDataError('a chart needs at least one series');
  if (categories.length === 0) throw new ChartDataError('a chart needs at least one category');
  categories.forEach((c, i) => {
    if (typeof c === 'number' ? !Number.isFinite(c) : typeof c !== 'string') throw new ChartDataError(`category ${i + 1} is not text or a finite number`);
  });
  for (const s of series) {
    if (s.values.length !== categories.length) throw new ChartDataError(`series "${s.name}" has ${s.values.length} values for ${categories.length} categories`);
    s.values.forEach((v, i) => {
      if (v !== null && !Number.isFinite(v)) throw new ChartDataError(`series "${s.name}": value ${i + 1} is not a finite number`);
    });
  }
  if (type === undefined) return;
  if (stacked && !STACKABLE.has(type)) throw new ChartDataError(`a ${type} chart cannot be stacked; only column, bar and area charts can`);
  if (type === 'pie' && series.length > 1) throw new ChartDataError(`a pie chart has one series, not ${series.length}`);
  if (type === 'scatter' && categories.some((c) => typeof c !== 'number')) throw new ChartDataError('a scatter chart needs numbers as its categories (the x values)');
}

// ---------------------------------------------------------------------------------------------
// Schema sequences (for inserting into existing elements)

const CHART_ORDER = ['c:title', 'c:autoTitleDeleted', 'c:pivotFmts', 'c:view3D', 'c:floor', 'c:sideWall', 'c:backWall', 'c:plotArea', 'c:legend', 'c:plotVisOnly', 'c:dispBlanksAs', 'c:showDLblsOverMax', 'c:extLst'];
const TITLE_ORDER = ['c:tx', 'c:layout', 'c:overlay', 'c:spPr', 'c:txPr', 'c:extLst'];
const LEGEND_ORDER = ['c:legendPos', 'c:legendEntry', 'c:layout', 'c:overlay', 'c:spPr', 'c:txPr', 'c:extLst'];
/** CT_CatAx, CT_ValAx, CT_DateAx and CT_SerAx merged: their orders agree. */
const AXIS_ORDER = [
  'c:axId', 'c:scaling', 'c:delete', 'c:axPos', 'c:majorGridlines', 'c:minorGridlines', 'c:title', 'c:numFmt', 'c:majorTickMark', 'c:minorTickMark', 'c:tickLblPos', 'c:spPr', 'c:txPr',
  'c:crossAx', 'c:crosses', 'c:crossesAt', 'c:crossBetween', 'c:auto', 'c:lblAlgn', 'c:lblOffset', 'c:baseTimeUnit', 'c:majorUnit', 'c:majorTimeUnit', 'c:minorUnit', 'c:minorTimeUnit',
  'c:tickLblSkip', 'c:tickMarkSkip', 'c:dispUnits', 'c:noMultiLvlLbl', 'c:extLst',
];
/** Every series type's sequence (bar, line, area, pie, scatter, bubble, radar, surface) merged: their orders agree. */
const SER_ORDER = [
  'c:idx', 'c:order', 'c:tx', 'c:spPr', 'c:invertIfNegative', 'c:pictureOptions', 'c:marker', 'c:explosion', 'c:dPt', 'c:dLbls', 'c:trendline', 'c:errBars',
  'c:cat', 'c:xVal', 'c:val', 'c:yVal', 'c:bubbleSize', 'c:bubble3D', 'c:shape', 'c:smooth', 'c:extLst',
];
/** Plot children that come before the series. */
const BEFORE_SER = new Set(['c:barDir', 'c:grouping', 'c:varyColors', 'c:scatterStyle', 'c:radarStyle', 'c:ofPieType', 'c:wireframe']);
const AXES = new Set(['c:catAx', 'c:valAx', 'c:dateAx', 'c:serAx']);

// ---------------------------------------------------------------------------------------------
// Small builders and readers

const v = (name: string, value: string | number): XmlElement => el(name, { val: value });
const valOf = (parent: XmlElement | undefined, name: string): string | undefined => attr(firstChild(parent, name), 'val');
/** Every plot element (c:barChart, c:pie3DChart, …) ends in "Chart". */
const isPlot = (e: XmlElement): boolean => e.name.endsWith('Chart');

function numOf(parent: XmlElement | undefined, name: string): number | undefined {
  const s = valOf(parent, name);
  if (s === undefined) return undefined;
  const n = Number.parseFloat(s);
  return Number.isFinite(n) ? n : undefined;
}

/** Series colour `i` as Word picks it: the six accents, then rounds of darker and lighter variations. */
function accentColor(i: number): XmlElement {
  const [lumMod, lumOff] = VARIATIONS[Math.floor(i / 6) % VARIATIONS.length]!;
  const mods: XmlElement[] = [];
  if (lumMod !== 1) mods.push(v('a:lumMod', Math.round(lumMod * 100000)));
  if (lumOff !== 0) mods.push(v('a:lumOff', Math.round(lumOff * 100000)));
  return el('a:schemeClr', { val: `accent${(i % 6) + 1}` }, mods);
}

/** Word's light grey (text colour at 15%) for gridlines and the category axis line. */
function greyLine(): XmlElement {
  const grey = el('a:schemeClr', { val: 'tx1' }, [v('a:lumMod', 15000), v('a:lumOff', 85000)]);
  return el('a:ln', { w: 9525, cap: 'flat', cmpd: 'sng', algn: 'ctr' }, [el('a:solidFill', {}, [grey]), el('a:round')]);
}

// Data: with a formula, a reference (c:strRef, c:numRef) and its cache; without one, a literal
// (c:strLit, c:numLit, or c:v for a series name) for data that has no cells.

function strData(formula: string | undefined, values: string[]): XmlElement {
  const points = [v('c:ptCount', values.length), ...values.map((s, i) => el('c:pt', { idx: i }, [el('c:v', {}, [text(s)])]))];
  return formula === undefined ? el('c:strLit', {}, points) : el('c:strRef', {}, [el('c:f', {}, [text(formula)]), el('c:strCache', {}, points)]);
}

/** A missing value has no c:pt. */
function numData(formula: string | undefined, values: Array<number | null>, formatCode: string): XmlElement {
  const points = [el('c:formatCode', {}, [text(formatCode)]), v('c:ptCount', values.length)];
  values.forEach((n, i) => {
    if (n !== null) points.push(el('c:pt', { idx: i }, [el('c:v', {}, [text(String(n))])]));
  });
  return formula === undefined ? el('c:numLit', {}, points) : el('c:numRef', {}, [el('c:f', {}, [text(formula)]), el('c:numCache', {}, points)]);
}

function seriesName(formula: string | undefined, name: string): XmlElement {
  return el('c:tx', {}, [formula === undefined ? el('c:v', {}, [text(name)]) : strData(formula, [name])]);
}

/** c:cat or c:xVal: numbers when every category is a number, text otherwise. */
function categoryData(name: string, formula: string | undefined, categories: Array<string | number>, formatCode = 'General'): XmlElement {
  const numeric = categories.every((c): c is number => typeof c === 'number');
  return el(name, {}, [numeric ? numData(formula, categories as number[], formatCode) : strData(formula, categories.map(String))]);
}

function formatCodeOf(data: XmlElement | undefined): string | undefined {
  const fc = firstChild(childPath(data, 'c:numRef', 'c:numCache') ?? firstChild(data, 'c:numLit'), 'c:formatCode');
  return fc ? textContent(fc) : undefined;
}

/** The cell reference of a series name, categories or values; undefined for literal data. */
function formulaOf(data: XmlElement | undefined): string | undefined {
  const ref = firstChild(data, 'c:strRef') ?? firstChild(data, 'c:numRef') ?? firstChild(data, 'c:multiLvlStrRef');
  const f = firstChild(ref, 'c:f');
  return f ? textContent(f) : undefined;
}

/** Replace the text of a c:rich (or c:txPr) body, keeping its first paragraph and run formatting. */
function setRichText(rich: XmlElement, value: string): void {
  const paragraphs = elements(rich, 'a:p');
  const pPr = firstChild(paragraphs[0], 'a:pPr');
  const rPr = firstChild(paragraphs.flatMap((p) => elements(p, 'a:r'))[0], 'a:rPr');
  const lines = value.split(/\r?\n/).map((line) =>
    el('a:p', {}, [...(pPr ? [cloneElement(pPr)] : []), el('a:r', {}, [rPr ? cloneElement(rPr) : el('a:rPr', { lang: 'en-US' }), el('a:t', {}, [text(line)])])]),
  );
  rich.children = [...rich.children.filter((c) => !(c.type === 'element' && c.name === 'a:p')), ...lines];
}

function richText(rich: XmlElement): string {
  return elements(rich, 'a:p')
    .map((p) =>
      elements(p)
        .filter((r) => r.name === 'a:r' || r.name === 'a:fld')
        .map((r) => textContent(firstChild(r, 'a:t') ?? r))
        .join(''),
    )
    .join('\n');
}

/** An empty title body: `size` in hundredths of a point, not bold; vertical titles read bottom to top. */
function newRich(size: number, vertical: boolean): XmlElement {
  return el('c:rich', {}, [
    el('a:bodyPr', vertical ? { rot: -5400000, vert: 'horz' } : {}),
    el('a:lstStyle'),
    el('a:p', {}, [el('a:pPr', {}, [el('a:defRPr', { sz: size, b: 0 })]), el('a:r', {}, [el('a:rPr', { lang: 'en-US' })])]),
  ]);
}

function newTitle(value: string, size: number, vertical: boolean): XmlElement {
  const rich = newRich(size, vertical);
  setRichText(rich, value);
  return el('c:title', {}, [el('c:tx', {}, [rich]), v('c:overlay', 0)]);
}

const LEGEND_POS: Record<Exclude<LegendPosition, 'none'>, string> = { bottom: 'b', top: 't', left: 'l', right: 'r' };
const LEGEND_NAMES: Record<string, LegendPosition> = { b: 'bottom', t: 'top', l: 'left', r: 'right', tr: 'right' };

function newLegend(position: Exclude<LegendPosition, 'none'>): XmlElement {
  return el('c:legend', {}, [v('c:legendPos', LEGEND_POS[position]), v('c:overlay', 0)]);
}

// ---------------------------------------------------------------------------------------------
// New charts

type SeriesStyle = 'bar' | 'line' | 'area' | 'pie' | 'scatter';
type Ranges = ReturnType<typeof chartRanges>;

const CAT_AX = '100000001';
const VAL_AX = '100000002';

function newSeries(style: SeriesStyle, i: number, data: ChartSeriesData, categories: Array<string | number>, ranges: Ranges): XmlElement {
  const color = accentColor(i);
  const children: XmlElement[] = [v('c:idx', i), v('c:order', i), seriesName(ranges.name(i), data.name)];
  switch (style) {
    case 'bar':
    case 'area':
      children.push(el('c:spPr', {}, [el('a:solidFill', {}, [color]), el('a:ln', {}, [el('a:noFill')]), el('a:effectLst')]));
      if (style === 'bar') children.push(v('c:invertIfNegative', 0));
      break;
    case 'line':
      children.push(
        el('c:spPr', {}, [el('a:ln', { w: 28575, cap: 'rnd' }, [el('a:solidFill', {}, [color]), el('a:round')]), el('a:effectLst')]),
        el('c:marker', {}, [v('c:symbol', 'none')]),
      );
      break;
    case 'scatter':
      // Markers only, as Word's default scatter chart.
      children.push(
        el('c:spPr', {}, [el('a:ln', { w: 25400, cap: 'rnd' }, [el('a:noFill'), el('a:round')]), el('a:effectLst')]),
        el('c:marker', {}, [
          v('c:symbol', 'circle'),
          v('c:size', 5),
          el('c:spPr', {}, [el('a:solidFill', {}, [color]), el('a:ln', { w: 9525 }, [el('a:solidFill', {}, [cloneElement(color)])]), el('a:effectLst')]),
        ]),
      );
      break;
    case 'pie':
      // No series fill: the plot varies colours by point.
      break;
  }
  const xy = style === 'scatter';
  children.push(categoryData(xy ? 'c:xVal' : 'c:cat', ranges.categories, categories));
  children.push(el(xy ? 'c:yVal' : 'c:val', {}, [numData(ranges.values(i), data.values, 'General')]));
  if (style === 'line' || style === 'scatter') children.push(v('c:smooth', 0));
  return el('c:ser', {}, children);
}

function newAxis(
  kind: 'c:catAx' | 'c:valAx',
  id: string,
  crossId: string,
  pos: 'b' | 'l',
  opts: { title?: string | undefined; gridlines?: boolean; axisLine?: boolean; crossBetween?: 'between' | 'midCat' },
): XmlElement {
  const children: XmlElement[] = [v('c:axId', id), el('c:scaling', {}, [v('c:orientation', 'minMax')]), v('c:delete', 0), v('c:axPos', pos)];
  if (opts.gridlines) children.push(el('c:majorGridlines', {}, [el('c:spPr', {}, [greyLine(), el('a:effectLst')])]));
  if (opts.title) children.push(newTitle(opts.title, 1000, pos === 'l'));
  children.push(
    el('c:numFmt', { formatCode: 'General', sourceLinked: 1 }),
    v('c:majorTickMark', 'none'),
    v('c:minorTickMark', 'none'),
    v('c:tickLblPos', 'nextTo'),
    el('c:spPr', {}, [el('a:noFill'), opts.axisLine ? greyLine() : el('a:ln', {}, [el('a:noFill')]), el('a:effectLst')]),
    v('c:crossAx', crossId),
    v('c:crosses', 'autoZero'),
  );
  if (kind === 'c:catAx') children.push(v('c:auto', 1), v('c:lblAlgn', 'ctr'), v('c:lblOffset', 100), v('c:noMultiLvlLbl', 0));
  else children.push(v('c:crossBetween', opts.crossBetween ?? 'between'));
  return el(kind, {}, children);
}

function newPlotArea(chart: NewChart): XmlElement {
  const ranges = chartRanges(chart.series.length, chart.categories.length);
  const series = (style: SeriesStyle): XmlElement[] => chart.series.map((s, i) => newSeries(style, i, s, chart.categories, ranges));
  const axIds = [v('c:axId', CAT_AX), v('c:axId', VAL_AX)];
  const titles = chart.axisTitles ?? {};
  const grouping = chart.stacked === 'stacked' ? 'stacked' : chart.stacked === 'percent' ? 'percentStacked' : undefined;
  const layout = el('c:layout');

  let plot: XmlElement;
  switch (chart.type) {
    case 'column':
    case 'bar':
      plot = el('c:barChart', {}, [
        v('c:barDir', chart.type === 'bar' ? 'bar' : 'col'),
        v('c:grouping', grouping ?? 'clustered'),
        v('c:varyColors', 0),
        ...series('bar'),
        v('c:gapWidth', grouping ? 150 : 219),
        v('c:overlap', grouping ? 100 : -27),
        ...axIds,
      ]);
      break;
    case 'line':
      plot = el('c:lineChart', {}, [v('c:grouping', 'standard'), v('c:varyColors', 0), ...series('line'), v('c:marker', 1), ...axIds]);
      break;
    case 'area':
      plot = el('c:areaChart', {}, [v('c:grouping', grouping ?? 'standard'), v('c:varyColors', 0), ...series('area'), ...axIds]);
      break;
    case 'pie':
    case 'doughnut': {
      const shared = [v('c:varyColors', 1), ...series('pie'), v('c:firstSliceAng', 0)];
      return el('c:plotArea', {}, [layout, chart.type === 'pie' ? el('c:pieChart', {}, shared) : el('c:doughnutChart', {}, [...shared, v('c:holeSize', 75)])]);
    }
    case 'scatter':
      return el('c:plotArea', {}, [
        layout,
        el('c:scatterChart', {}, [v('c:scatterStyle', 'lineMarker'), v('c:varyColors', 0), ...series('scatter'), ...axIds]),
        newAxis('c:valAx', CAT_AX, VAL_AX, 'b', { title: titles.category, axisLine: true, crossBetween: 'midCat' }),
        newAxis('c:valAx', VAL_AX, CAT_AX, 'l', { title: titles.value, gridlines: true, crossBetween: 'midCat' }),
      ]);
  }
  // Columns stand on a horizontal category axis; bars lie along a vertical one.
  const catPos = chart.type === 'bar' ? 'l' : 'b';
  return el('c:plotArea', {}, [
    layout,
    plot,
    newAxis('c:catAx', CAT_AX, VAL_AX, catPos, { title: titles.category, axisLine: true }),
    // Areas run from edge to edge of the plot; columns, bars and lines sit between the tick marks.
    newAxis('c:valAx', VAL_AX, CAT_AX, catPos === 'b' ? 'l' : 'b', { title: titles.value, gridlines: true, crossBetween: chart.type === 'area' ? 'midCat' : 'between' }),
  ]);
}

/** A new chart part's root, looking like a chart Word has just inserted; its formulas point into `Sheet1` of the workbook `workbookRelId`. */
export function newChartSpace(chart: NewChart, workbookRelId: string): XmlElement {
  validateChartData(chart.type, chart.stacked, chart.categories, chart.series);
  const children: XmlElement[] = [];
  if (chart.title) children.push(newTitle(chart.title, 1400, false));
  children.push(v('c:autoTitleDeleted', chart.title ? 0 : 1), newPlotArea(chart));
  const legend = chart.legend ?? 'bottom';
  if (legend !== 'none') children.push(newLegend(legend));
  children.push(v('c:plotVisOnly', 1), v('c:dispBlanksAs', 'gap'));
  return el('c:chartSpace', { 'xmlns:c': NS.c, 'xmlns:a': NS.a, 'xmlns:r': NS.r }, [
    v('c:date1904', 0),
    v('c:lang', 'en-US'),
    v('c:roundedCorners', 0),
    el('c:chart', {}, children),
    el('c:externalData', { 'r:id': workbookRelId }, [v('c:autoUpdate', 0)]),
  ]);
}

// ---------------------------------------------------------------------------------------------
// Reading

/** The category (x) and value (y) axes of the first plot that has two. */
function chartAxes(chartSpace: XmlElement): { category?: XmlElement | undefined; value?: XmlElement | undefined } {
  const plotArea = childPath(chartSpace, 'c:chart', 'c:plotArea');
  if (!plotArea) return {};
  const axes = elements(plotArea).filter((e) => AXES.has(e.name));
  for (const plot of elements(plotArea).filter(isPlot)) {
    const own = elements(plot, 'c:axId')
      .map((id) => axes.find((a) => valOf(a, 'c:axId') === attr(id, 'val')))
      .filter((a): a is XmlElement => a !== undefined);
    if (own.length < 2) continue;
    // Scatter and bubble charts have two value axes: x first.
    if (own.every((a) => a.name === 'c:valAx')) return { category: own[0], value: own[1] };
    return { category: own.find((a) => a.name !== 'c:valAx'), value: own.find((a) => a.name === 'c:valAx') };
  }
  return {};
}

/** A chart's data as one table (series in the reader's order: by plot, then c:order). */
export function readChartTable(chartSpace: XmlElement): ChartTable {
  const model = readChart(chartSpace);
  const all = model.plots.flatMap((p) => p.series);
  const first = all[0];
  const categories: Array<string | number> = first ? (first.categoryNumbers ? first.categoryNumbers.map((n) => n ?? '') : first.categories.slice()) : [];
  const count = Math.max(categories.length, ...all.map((s) => s.values.length));
  while (categories.length < count) categories.push('');
  const table: ChartTable = {
    kind: model.plots.map(plotName).join(' + ') || 'chart',
    axisTitles: {},
    legend: model.legend ? (LEGEND_NAMES[model.legend.position] ?? 'right') : 'none',
    categories,
    series: all.map((s) => ({ name: s.name, values: Array.from({ length: count }, (_, i) => s.values[i] ?? null) })),
    formatCode: first?.formatCode ?? 'General',
  };
  if (model.title && !model.title.auto) table.title = model.title.text;
  const axes = chartAxes(chartSpace);
  for (const key of ['category', 'value'] as const) {
    const id = valOf(axes[key], 'c:axId');
    const title = id === undefined ? undefined : model.axes.find((a) => a.id === id)?.title;
    if (title && !title.auto) table.axisTitles[key] = title.text;
  }
  return table;
}

// ---------------------------------------------------------------------------------------------
// Patching

interface SeriesSlot {
  plot: XmlElement;
  ser: XmlElement;
  /** Added by this patch. */
  added?: boolean;
}

/** Series in the order the reader lists them: plots in document order, each sorted by c:order. */
function orderedSeries(plotArea: XmlElement | undefined): SeriesSlot[] {
  const out: SeriesSlot[] = [];
  for (const plot of plotArea ? elements(plotArea).filter(isPlot) : []) {
    const sers = elements(plot, 'c:ser').map((ser, i) => {
      const idx = numOf(ser, 'c:idx') ?? i;
      return { plot, ser, order: numOf(ser, 'c:order') ?? idx };
    });
    out.push(...sers.sort((a, b) => a.order - b.order));
  }
  return out;
}

function replaceChild(parent: XmlElement, node: XmlElement, order: readonly string[]): void {
  const i = parent.children.findIndex((c) => c.type === 'element' && c.name === node.name);
  if (i >= 0) parent.children[i] = node;
  else insertOrdered(parent, node, order);
}

function recolor(spPr: XmlElement | undefined, color: XmlElement): void {
  for (const fill of [firstChild(spPr, 'a:solidFill'), childPath(spPr, 'a:ln', 'a:solidFill')]) {
    if (fill) fill.children = [cloneElement(color)];
  }
}

/** A new series after the last one, styled like it in the next accent colour. */
function addSeries(plotArea: XmlElement | undefined, slots: SeriesSlot[]): SeriesSlot {
  const all = plotArea ? elements(plotArea).filter(isPlot).flatMap((p) => elements(p, 'c:ser')) : [];
  const idx = Math.max(-1, ...all.map((s, i) => numOf(s, 'c:idx') ?? i)) + 1;
  const order = Math.max(-1, ...all.map((s, i) => numOf(s, 'c:order') ?? numOf(s, 'c:idx') ?? i)) + 1;
  const last = slots[slots.length - 1];
  let slot: SeriesSlot;
  if (last) {
    const ser = cloneElement(last.ser);
    // Extensions hold the original's unique id and cell-linked extras.
    removeChildren(ser, 'c:extLst');
    const color = accentColor(idx);
    recolor(firstChild(ser, 'c:spPr'), color);
    recolor(childPath(ser, 'c:marker', 'c:spPr'), color);
    last.plot.children.splice(last.plot.children.indexOf(last.ser) + 1, 0, ser);
    slot = { plot: last.plot, ser };
  } else {
    const plot = plotArea ? elements(plotArea).find(isPlot) : undefined;
    if (!plot) throw new ChartDataError('the chart has no plot to add a series to');
    const ser = el('c:ser');
    let at = 0;
    plot.children.forEach((c, i) => {
      if (c.type === 'element' && BEFORE_SER.has(c.name)) at = i + 1;
    });
    plot.children.splice(at, 0, ser);
    slot = { plot, ser };
  }
  ensureChild(slot.ser, 'c:idx', SER_ORDER).attrs.val = String(idx);
  ensureChild(slot.ser, 'c:order', SER_ORDER).attrs.val = String(order);
  slot.added = true;
  return slot;
}

function fillSeries(slot: SeriesSlot, i: number, data: ChartSeriesData, categories: Array<string | number>, ranges: Ranges, keepFormulas: boolean): void {
  const { plot, ser } = slot;
  const xy = plot.name === 'c:scatterChart' || plot.name === 'c:bubbleChart';
  const tx = firstChild(ser, 'c:tx');
  const cat = firstChild(ser, xy ? 'c:xVal' : 'c:cat');
  const val = firstChild(ser, xy ? 'c:yVal' : 'c:val');
  // With keepFormulas, existing references stay; added series have no cells, so they hold literals.
  const formula = (old: XmlElement | undefined, fresh: string): string | undefined => (!keepFormulas ? fresh : slot.added ? undefined : formulaOf(old));
  replaceChild(ser, seriesName(formula(tx, ranges.name(i)), data.name), SER_ORDER);
  replaceChild(ser, categoryData(xy ? 'c:xVal' : 'c:cat', formula(cat, ranges.categories), categories, formatCodeOf(cat)), SER_ORDER);
  replaceChild(ser, el(xy ? 'c:yVal' : 'c:val', {}, [numData(formula(val, ranges.values(i)), data.values, formatCodeOf(val) ?? 'General')]), SER_ORDER);
  // Formatting of points past the new last one goes.
  const n = categories.length;
  const pastEnd = (c: XmlNode, name: string): boolean => c.type === 'element' && c.name === name && (numOf(c, 'c:idx') ?? -1) >= n;
  ser.children = ser.children.filter((c) => !pastEnd(c, 'c:dPt'));
  const labels = firstChild(ser, 'c:dLbls');
  if (labels) labels.children = labels.children.filter((c) => !pastEnd(c, 'c:dLbl'));
}

function patchData(chart: XmlElement, categories: Array<string | number>, series: ChartSeriesData[], keepFormulas: boolean): void {
  const plotArea = firstChild(chart, 'c:plotArea');
  const slots = orderedSeries(plotArea);
  const removed = slots.splice(series.length);
  for (const { plot, ser } of removed) plot.children = plot.children.filter((c) => c !== ser);
  // A plot left without series goes, and so do the axes only it used (a combination chart's secondary axes).
  const emptied = new Set(removed.map((s) => s.plot).filter((p) => elements(p, 'c:ser').length === 0));
  if (plotArea && emptied.size > 0) {
    plotArea.children = plotArea.children.filter((c) => !(c.type === 'element' && emptied.has(c)));
    const used = new Set(elements(plotArea).filter(isPlot).flatMap((p) => elements(p, 'c:axId').map((a) => attr(a, 'val'))));
    plotArea.children = plotArea.children.filter((c) => !(c.type === 'element' && AXES.has(c.name) && !used.has(valOf(c, 'c:axId'))));
  }
  while (slots.length < series.length) slots.push(addSeries(plotArea, slots));
  const ranges = chartRanges(series.length, categories.length);
  slots.forEach((slot, i) => fillSeries(slot, i, series[i]!, categories, ranges, keepFormulas));
}

/** Set, replace or (with '') remove the c:title of a chart or axis, keeping the formatting it had. */
function setTitle(parent: XmlElement, order: readonly string[], value: string, size: number, vertical: boolean): void {
  if (value === '') {
    removeChildren(parent, 'c:title');
    return;
  }
  const title = firstChild(parent, 'c:title');
  if (!title) {
    insertOrdered(parent, newTitle(value, size, vertical), order);
    return;
  }
  let rich = childPath(title, 'c:tx', 'c:rich');
  if (rich && richText(rich) === value) return;
  if (!rich) {
    // An automatic title or one linked to a cell: write the text in the formatting it was shown with.
    const txPr = firstChild(title, 'c:txPr');
    rich = txPr ? { ...cloneElement(txPr), name: 'c:rich' } : newRich(size, vertical);
    removeChildren(title, 'c:tx');
    insertOrdered(title, el('c:tx', {}, [rich]), TITLE_ORDER);
  }
  setRichText(rich, value);
}

function setLegend(chart: XmlElement, position: LegendPosition): void {
  if (position === 'none') {
    removeChildren(chart, 'c:legend');
    return;
  }
  const legend = firstChild(chart, 'c:legend');
  if (!legend) {
    insertOrdered(chart, newLegend(position), CHART_ORDER);
    return;
  }
  const pos = ensureChild(legend, 'c:legendPos', LEGEND_ORDER);
  if (pos.attrs.val === LEGEND_POS[position]) return;
  pos.attrs.val = LEGEND_POS[position];
  // A manual layout would keep the legend where it was.
  removeChildren(legend, 'c:layout');
}

/**
 * Change an existing chart in place. New data rewrites every series' name, categories and values;
 * series keep their formatting by position, added series copy the last one's in the next accent
 * colour, and extra series are removed from the end, along with any plot (and the axes only it
 * used) left without series. Formulas point into Sheet1 of the chart's own workbook (series i in
 * column i + 1); with `keepFormulas`, for a chart linked to an outside workbook, existing references
 * stay and only their caches change, while added series hold literal data. Categories or series
 * alone are combined with the chart's current data. Throws ChartDataError (before changing
 * anything) when the resulting table is invalid.
 */
export function patchChartSpace(chartSpace: XmlElement, changes: ChartChanges, options: { keepFormulas?: boolean } = {}): void {
  const chart = firstChild(chartSpace, 'c:chart');
  if (!chart) throw new ChartDataError('the chart part has no c:chart element');
  if (changes.categories !== undefined || changes.series !== undefined) {
    const current = changes.categories !== undefined && changes.series !== undefined ? undefined : readChartTable(chartSpace);
    const categories = changes.categories ?? current!.categories;
    const series = changes.series ?? current!.series;
    validateChartData(undefined, undefined, categories, series);
    patchData(chart, categories, series, options.keepFormulas === true);
  }
  if (changes.title !== undefined) {
    setTitle(chart, CHART_ORDER, changes.title, 1400, false);
    ensureChild(chart, 'c:autoTitleDeleted', CHART_ORDER).attrs.val = changes.title === '' ? '1' : '0';
  }
  if (changes.legend !== undefined) setLegend(chart, changes.legend);
  if (changes.axisTitles) {
    const axes = chartAxes(chartSpace);
    for (const key of ['category', 'value'] as const) {
      const value = changes.axisTitles[key];
      const axis = axes[key];
      if (value === undefined || !axis) continue;
      const pos = valOf(axis, 'c:axPos');
      setTitle(axis, AXIS_ORDER, value, 1000, pos === 'l' || pos === 'r');
    }
  }
}
