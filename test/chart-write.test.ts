/** Writing charts: the embedded workbook, and building chart parts. */
import { describe, expect, it } from 'vitest';
import { chartRanges, columnName, isWorkbook, workbookFor, workbookHasMore } from '../src/charts/workbook.js';
import { parseXml } from '../src/xml/parser.js';
import { attr, childPath, cloneElement, el, elements, findAll, firstChild, serializeNode, text, textContent, walk, type XmlElement } from '../src/xml/index.js';
import { readChart, type Grouping, type PlotType } from '../src/charts/index.js';
import {
  ChartDataError,
  newChartSpace,
  patchChartSpace,
  readChartTable,
  validateChartData,
  type ChartSeriesData,
  type NewChart,
  type NewChartType,
} from '../src/charts/write.js';
import { readZip, writeZip } from '../src/zip/zip.js';

describe('chart workbook', () => {
  it('names columns and chart ranges', () => {
    expect([0, 25, 26, 701, 702].map(columnName)).toEqual(['A', 'Z', 'AA', 'ZZ', 'AAA']);
    const r = chartRanges(2, 4);
    expect([r.name(0), r.categories, r.values(1)]).toEqual(['Sheet1!$B$1', 'Sheet1!$A$2:$A$5', 'Sheet1!$C$2:$C$5']);
  });

  it('workbookFor writes one sheet with the table, escaped, and leaves missing values empty', async () => {
    const bytes = await workbookFor({ categories: ['Q1', 'R&D <2026>'], series: [{ name: 'Sales "EU"', values: [1.5, null] }] });
    const entries = readZip(bytes);
    expect(entries.map((e) => e.name).sort()).toEqual([
      '[Content_Types].xml',
      '_rels/.rels',
      'xl/_rels/workbook.xml.rels',
      'xl/styles.xml',
      'xl/workbook.xml',
      'xl/worksheets/sheet1.xml',
    ]);
    const sheet = entries.find((e) => e.name === 'xl/worksheets/sheet1.xml')!;
    const root = parseXml(new TextDecoder().decode(await sheet.read())).root;
    const cells = Object.fromEntries(findAll(root, 'c').map((c) => [attr(c, 'r'), textContent(c)]));
    expect(cells).toEqual({ B1: 'Sales "EU"', A2: 'Q1', B2: '1.5', A3: 'R&D <2026>' });
  });

  it('workbookHasMore spots data beyond the chart table', async () => {
    const bytes = await workbookFor({ categories: ['Q1', 'R&D <2026>'], series: [{ name: 'Sales "EU"', values: [1.5, null] }] });
    expect(await workbookHasMore(bytes, 2, 1)).toBe(false);
    expect(await workbookHasMore(await workbookFor({ categories: ['a', 'b', 'c'], series: [{ name: 's', values: [1, 2, 3] }] }), 2, 1)).toBe(true);
    expect(await workbookHasMore(new Uint8Array([1, 2, 3]), 2, 1)).toBe(false);
  });

  it('isWorkbook tells a workbook package from bytes that are not one', async () => {
    expect(isWorkbook(await workbookFor({ categories: ['Q1'], series: [{ name: 's', values: [1] }] }))).toBe(true);
    expect(isWorkbook(new Uint8Array([1, 2, 3]))).toBe(false);
    expect(isWorkbook(await writeZip([{ name: 'word/document.xml', data: new TextEncoder().encode('<w:document/>') }]))).toBe(false);
  });
});

// ---------------------------------------------------------------------------------------------
// Chart XML

/** DrawingML chart schema sequences (ECMA-376 dml-chart.xsd), to check element order. */
const SEQUENCES: Record<string, string[]> = {
  'c:chartSpace': ['c:date1904', 'c:lang', 'c:roundedCorners', 'mc:AlternateContent', 'c:style', 'c:clrMapOvr', 'c:pivotSource', 'c:protection', 'c:chart', 'c:spPr', 'c:txPr', 'c:externalData', 'c:printSettings', 'c:userShapes', 'c:extLst'],
  'c:chart': ['c:title', 'c:autoTitleDeleted', 'c:pivotFmts', 'c:view3D', 'c:floor', 'c:sideWall', 'c:backWall', 'c:plotArea', 'c:legend', 'c:plotVisOnly', 'c:dispBlanksAs', 'c:showDLblsOverMax', 'c:extLst'],
  'c:plotArea': ['c:layout', 'CHART', 'AXIS', 'c:dTable', 'c:spPr', 'c:extLst'],
  'c:barChart': ['c:barDir', 'c:grouping', 'c:varyColors', 'c:ser', 'c:dLbls', 'c:gapWidth', 'c:overlap', 'c:serLines', 'c:axId', 'c:extLst'],
  'c:lineChart': ['c:grouping', 'c:varyColors', 'c:ser', 'c:dLbls', 'c:dropLines', 'c:hiLowLines', 'c:upDownBars', 'c:marker', 'c:smooth', 'c:axId', 'c:extLst'],
  'c:areaChart': ['c:grouping', 'c:varyColors', 'c:ser', 'c:dLbls', 'c:dropLines', 'c:axId', 'c:extLst'],
  'c:pieChart': ['c:varyColors', 'c:ser', 'c:dLbls', 'c:firstSliceAng', 'c:extLst'],
  'c:doughnutChart': ['c:varyColors', 'c:ser', 'c:dLbls', 'c:firstSliceAng', 'c:holeSize', 'c:extLst'],
  'c:scatterChart': ['c:scatterStyle', 'c:varyColors', 'c:ser', 'c:dLbls', 'c:axId', 'c:extLst'],
  'c:barChart/c:ser': ['c:idx', 'c:order', 'c:tx', 'c:spPr', 'c:invertIfNegative', 'c:pictureOptions', 'c:dPt', 'c:dLbls', 'c:trendline', 'c:errBars', 'c:cat', 'c:val', 'c:shape', 'c:extLst'],
  'c:lineChart/c:ser': ['c:idx', 'c:order', 'c:tx', 'c:spPr', 'c:marker', 'c:dPt', 'c:dLbls', 'c:trendline', 'c:errBars', 'c:cat', 'c:val', 'c:smooth', 'c:extLst'],
  'c:areaChart/c:ser': ['c:idx', 'c:order', 'c:tx', 'c:spPr', 'c:pictureOptions', 'c:dPt', 'c:dLbls', 'c:trendline', 'c:errBars', 'c:cat', 'c:val', 'c:extLst'],
  'c:pieChart/c:ser': ['c:idx', 'c:order', 'c:tx', 'c:spPr', 'c:explosion', 'c:dPt', 'c:dLbls', 'c:cat', 'c:val', 'c:extLst'],
  'c:doughnutChart/c:ser': ['c:idx', 'c:order', 'c:tx', 'c:spPr', 'c:explosion', 'c:dPt', 'c:dLbls', 'c:cat', 'c:val', 'c:extLst'],
  'c:scatterChart/c:ser': ['c:idx', 'c:order', 'c:tx', 'c:spPr', 'c:marker', 'c:dPt', 'c:dLbls', 'c:trendline', 'c:errBars', 'c:xVal', 'c:yVal', 'c:smooth', 'c:extLst'],
  'c:catAx': ['c:axId', 'c:scaling', 'c:delete', 'c:axPos', 'c:majorGridlines', 'c:minorGridlines', 'c:title', 'c:numFmt', 'c:majorTickMark', 'c:minorTickMark', 'c:tickLblPos', 'c:spPr', 'c:txPr', 'c:crossAx', 'c:crosses', 'c:crossesAt', 'c:auto', 'c:lblAlgn', 'c:lblOffset', 'c:tickLblSkip', 'c:tickMarkSkip', 'c:noMultiLvlLbl', 'c:extLst'],
  'c:valAx': ['c:axId', 'c:scaling', 'c:delete', 'c:axPos', 'c:majorGridlines', 'c:minorGridlines', 'c:title', 'c:numFmt', 'c:majorTickMark', 'c:minorTickMark', 'c:tickLblPos', 'c:spPr', 'c:txPr', 'c:crossAx', 'c:crosses', 'c:crossesAt', 'c:crossBetween', 'c:majorUnit', 'c:minorUnit', 'c:dispUnits', 'c:extLst'],
  'c:title': ['c:tx', 'c:layout', 'c:overlay', 'c:spPr', 'c:txPr', 'c:extLst'],
  'c:legend': ['c:legendPos', 'c:legendEntry', 'c:layout', 'c:overlay', 'c:spPr', 'c:txPr', 'c:extLst'],
  'c:marker': ['c:symbol', 'c:size', 'c:spPr', 'c:extLst'],
  'c:spPr': ['a:xfrm', 'a:custGeom', 'a:prstGeom', 'a:noFill', 'a:solidFill', 'a:gradFill', 'a:blipFill', 'a:pattFill', 'a:grpFill', 'a:ln', 'a:effectLst', 'a:effectDag', 'a:scene3d', 'a:sp3d', 'a:extLst'],
  'a:ln': ['a:noFill', 'a:solidFill', 'a:gradFill', 'a:pattFill', 'a:prstDash', 'a:custDash', 'a:round', 'a:bevel', 'a:miter', 'a:headEnd', 'a:tailEnd', 'a:extLst'],
  'c:rich': ['a:bodyPr', 'a:lstStyle', 'a:p'],
  'c:txPr': ['a:bodyPr', 'a:lstStyle', 'a:p'],
  'a:p': ['a:pPr', 'a:r', 'a:br', 'a:fld', 'a:endParaRPr'],
  'a:defRPr': ['a:ln', 'a:noFill', 'a:solidFill', 'a:gradFill', 'a:blipFill', 'a:pattFill', 'a:grpFill', 'a:effectLst', 'a:effectDag', 'a:highlight', 'a:uLnTx', 'a:uLn', 'a:uFillTx', 'a:uFill', 'a:latin', 'a:ea', 'a:cs', 'a:sym', 'a:hlinkClick', 'a:hlinkMouseOver', 'a:rtl', 'a:extLst'],
  'c:strCache': ['c:ptCount', 'c:pt', 'c:extLst'],
  'c:numCache': ['c:formatCode', 'c:ptCount', 'c:pt', 'c:extLst'],
  'c:strLit': ['c:ptCount', 'c:pt', 'c:extLst'],
  'c:numLit': ['c:formatCode', 'c:ptCount', 'c:pt', 'c:extLst'],
};

const AXES = new Set(['c:catAx', 'c:valAx', 'c:dateAx', 'c:serAx']);

/** Children out of schema order (or unknown to the sequence), anywhere under `root`. */
function orderProblems(root: XmlElement): string[] {
  const problems: string[] = [];
  walk(root, (e, parent) => {
    const seq = SEQUENCES[e.name === 'c:ser' && parent ? `${parent.name}/c:ser` : e.name];
    if (!seq) return;
    let last = -1;
    for (const c of elements(e)) {
      const key = e.name === 'c:plotArea' && c.name.endsWith('Chart') ? 'CHART' : e.name === 'c:plotArea' && AXES.has(c.name) ? 'AXIS' : c.name;
      const rank = seq.indexOf(key);
      if (rank < 0) problems.push(`${e.name}: unexpected ${c.name}`);
      else if (rank < last) problems.push(`${e.name}: ${c.name} out of order`);
      else last = rank;
    }
  });
  return problems;
}

const names = (e: XmlElement | undefined): string[] => (e ? elements(e).map((c) => c.name) : []);
const val = (e: XmlElement | undefined, name: string): string | undefined => attr(firstChild(e, name), 'val');
const formula = (e: XmlElement | undefined): string | undefined => {
  const f = e ? findAll(e, 'c:f')[0] : undefined;
  return f ? textContent(f) : undefined;
};
/** Serialise and parse again, as a saved file would be read. */
const reparse = (e: XmlElement): XmlElement => parseXml(serializeNode(e)).root;

const QUARTERS = ['Q1', 'Q2', 'Q3', 'Q4'];
const TWO: ChartSeriesData[] = [
  { name: 'North', values: [1, 2, 3, 4] },
  { name: 'South', values: [5, 6, 7, 8] },
];

describe('chart writer', () => {
  it('writes every new-chart type so the reader reads it back', () => {
    const cases: Array<[NewChartType, NewChart['stacked'], PlotType, 'col' | 'bar', Grouping, string]> = [
      ['column', undefined, 'bar', 'col', 'clustered', 'clustered column'],
      ['column', 'stacked', 'bar', 'col', 'stacked', 'stacked column'],
      ['column', 'percent', 'bar', 'col', 'percentStacked', '100% stacked column'],
      ['bar', undefined, 'bar', 'bar', 'clustered', 'clustered bar'],
      ['bar', 'stacked', 'bar', 'bar', 'stacked', 'stacked bar'],
      ['bar', 'percent', 'bar', 'bar', 'percentStacked', '100% stacked bar'],
      ['line', undefined, 'line', 'col', 'standard', 'line'],
      ['area', undefined, 'area', 'col', 'standard', 'area'],
      ['area', 'stacked', 'area', 'col', 'stacked', 'stacked area'],
      ['area', 'percent', 'area', 'col', 'percentStacked', '100% stacked area'],
      ['pie', undefined, 'pie', 'col', 'standard', 'pie'],
      ['doughnut', undefined, 'doughnut', 'col', 'standard', 'doughnut'],
      ['scatter', undefined, 'scatter', 'col', 'standard', 'scatter'],
    ];
    for (const [type, stacked, plotType, barDir, grouping, kind] of cases) {
      const categories = type === 'scatter' ? [1, 2, 3] : ['Q1', 'Q2', 'Q3'];
      const series = type === 'pie' ? [{ name: 'Share', values: [10, 20, 30] }] : [{ name: 'North', values: [1, 2, 3] }, { name: 'South', values: [4, null, 6] }];
      const chart: NewChart = { type, title: 'Sales', categories, series };
      if (stacked) chart.stacked = stacked;
      const cs = reparse(newChartSpace(chart, 'rId1'));
      const label = `${type} ${stacked ?? ''}`;
      expect(orderProblems(cs), label).toEqual([]);

      const model = readChart(cs);
      expect(model.plots, label).toHaveLength(1);
      const plot = model.plots[0]!;
      expect([plot.type, plot.barDir, plot.grouping], label).toEqual([plotType, barDir, grouping]);
      expect(plot.series.map((s) => s.name), label).toEqual(series.map((s) => s.name));
      expect(plot.series.map((s) => s.values), label).toEqual(series.map((s) => s.values));
      if (type === 'scatter') expect(plot.series[0]!.categoryNumbers, label).toEqual([1, 2, 3]);
      else expect(plot.series[0]!.categories, label).toEqual(categories);
      expect(model.title?.text, label).toBe('Sales');
      expect(model.legend?.position, label).toBe('b');

      expect(readChartTable(cs), label).toEqual({ kind, title: 'Sales', axisTitles: {}, legend: 'bottom', categories, series, formatCode: 'General' });
    }
  });

  it('escapes text and keeps missing values', () => {
    const chart: NewChart = {
      type: 'line',
      title: 'R&D <2026> "Q1"',
      categories: ['A & B', '<C>'],
      series: [{ name: 'S "1" & <2>', values: [1.5, null] }],
      legend: 'right',
      axisTitles: { category: 'Quarter <q>', value: 'USD & "EUR"' },
    };
    const xml = serializeNode(newChartSpace(chart, 'rId1'));
    expect(xml).toContain('R&amp;D &lt;2026&gt; "Q1"');
    const cs = parseXml(xml).root;
    expect(readChartTable(cs)).toEqual({
      kind: 'line',
      title: 'R&D <2026> "Q1"',
      axisTitles: { category: 'Quarter <q>', value: 'USD & "EUR"' },
      legend: 'right',
      categories: ['A & B', '<C>'],
      series: [{ name: 'S "1" & <2>', values: [1.5, null] }],
      formatCode: 'General',
    });
    const cache = findAll(firstChild(findAll(cs, 'c:ser')[0]!, 'c:val')!, 'c:numCache')[0]!;
    expect(val(cache, 'c:ptCount')).toBe('2');
    expect(findAll(cache, 'c:pt').map((p) => attr(p, 'idx'))).toEqual(['0']);
  });

  it('orders the elements as the schema requires', () => {
    const cs = newChartSpace({ type: 'column', categories: QUARTERS, series: TWO, axisTitles: { category: 'Quarter', value: 'USD' } }, 'rId1');
    expect(orderProblems(cs)).toEqual([]);
    expect(names(cs)).toEqual(['c:date1904', 'c:lang', 'c:roundedCorners', 'c:chart', 'c:spPr', 'c:txPr', 'c:externalData']);
    // As Word inserts a chart: a white chart area without a border, and 9 pt text in the text colour at 65%.
    const area = firstChild(cs, 'c:spPr')!;
    expect(names(area)).toEqual(['a:solidFill', 'a:ln', 'a:effectLst']);
    expect(val(firstChild(area, 'a:solidFill'), 'a:schemeClr')).toBe('bg1');
    expect(names(firstChild(area, 'a:ln'))[0]).toBe('a:noFill');
    expect(names(firstChild(cs, 'c:txPr'))).toEqual(['a:bodyPr', 'a:lstStyle', 'a:p']);
    const defRPr = childPath(cs, 'c:txPr', 'a:p', 'a:pPr', 'a:defRPr')!;
    expect(attr(defRPr, 'sz')).toBe('900');
    const textColour = childPath(defRPr, 'a:solidFill', 'a:schemeClr')!;
    expect([attr(textColour, 'val'), val(textColour, 'a:lumMod'), val(textColour, 'a:lumOff')]).toEqual(['tx1', '65000', '35000']);
    const model = readChart(cs);
    expect(model.chartArea).toMatchObject({ fill: { color: 'FFFFFF' }, line: { paint: null } });
    expect(model.textStyle).toMatchObject({ size: 9, color: '595959' });
    const ext = firstChild(cs, 'c:externalData');
    expect(attr(ext, 'r:id')).toBe('rId1');
    expect(val(ext, 'c:autoUpdate')).toBe('0');

    const plotArea = childPath(cs, 'c:chart', 'c:plotArea');
    expect(names(plotArea)).toEqual(['c:layout', 'c:barChart', 'c:catAx', 'c:valAx']);
    const bar = firstChild(plotArea, 'c:barChart')!;
    expect(names(bar).slice(0, 4)).toEqual(['c:barDir', 'c:grouping', 'c:varyColors', 'c:ser']);
    expect(names(bar).slice(-4)).toEqual(['c:gapWidth', 'c:overlap', 'c:axId', 'c:axId']);
    expect([val(bar, 'c:gapWidth'), val(bar, 'c:overlap')]).toEqual(['219', '-27']);
    const sers = elements(bar, 'c:ser');
    for (const s of sers) expect(names(s)).toEqual(['c:idx', 'c:order', 'c:tx', 'c:spPr', 'c:invertIfNegative', 'c:cat', 'c:val']);
    expect(sers.map((s) => val(childPath(s, 'c:spPr', 'a:solidFill'), 'a:schemeClr'))).toEqual(['accent1', 'accent2']);

    // Value axis: gridlines, then the title, then the number format.
    const valAx = firstChild(plotArea, 'c:valAx')!;
    const order = names(valAx);
    expect(order.indexOf('c:majorGridlines')).toBeLessThan(order.indexOf('c:title'));
    expect(order.indexOf('c:title')).toBeLessThan(order.indexOf('c:numFmt'));
    expect(attr(firstChild(valAx, 'c:numFmt'), 'formatCode')).toBe('General');
    expect(attr(firstChild(valAx, 'c:numFmt'), 'sourceLinked')).toBe('1');
    expect(val(valAx, 'c:crossBetween')).toBe('between');
    expect(firstChild(firstChild(plotArea, 'c:catAx'), 'c:majorGridlines')).toBeUndefined();

    expect(names(firstChild(cs, 'c:chart'))).toEqual(['c:autoTitleDeleted', 'c:plotArea', 'c:legend', 'c:plotVisOnly', 'c:dispBlanksAs']);
    expect(val(childPath(cs, 'c:chart', 'c:legend'), 'c:legendPos')).toBe('b');
    expect(val(childPath(cs, 'c:chart', 'c:legend'), 'c:overlay')).toBe('0');
    expect(val(firstChild(cs, 'c:chart'), 'c:plotVisOnly')).toBe('1');
    expect(val(firstChild(cs, 'c:chart'), 'c:dispBlanksAs')).toBe('gap');

    const stacked = childPath(newChartSpace({ type: 'bar', stacked: 'stacked', categories: QUARTERS, series: TWO }, 'rId1'), 'c:chart', 'c:plotArea', 'c:barChart')!;
    expect(names(stacked).slice(-4)).toEqual(['c:gapWidth', 'c:overlap', 'c:axId', 'c:axId']);
    expect([val(stacked, 'c:gapWidth'), val(stacked, 'c:overlap')]).toEqual(['150', '100']);

    // Lines: a 2.25 pt round-capped line in the accent colour and no markers.
    const line = findAll(newChartSpace({ type: 'line', categories: QUARTERS, series: TWO }, 'rId1'), 'c:ser')[1]!;
    const ln = childPath(line, 'c:spPr', 'a:ln')!;
    expect([attr(ln, 'w'), attr(ln, 'cap'), val(firstChild(ln, 'a:solidFill'), 'a:schemeClr')]).toEqual(['28575', 'rnd', 'accent2']);
    expect(val(firstChild(line, 'c:marker'), 'c:symbol')).toBe('none');

    // The seventh series starts the second round of accents, darker.
    const many = Array.from({ length: 7 }, (_, i) => ({ name: `S${i}`, values: [i, i, i, i] }));
    const seventh = findAll(newChartSpace({ type: 'column', categories: QUARTERS, series: many }, 'rId1'), 'c:ser')[6]!;
    const clr = childPath(seventh, 'c:spPr', 'a:solidFill', 'a:schemeClr')!;
    expect([attr(clr, 'val'), val(clr, 'a:lumMod')]).toEqual(['accent1', '60000']);

    // Pies vary colours by point; doughnuts have a 75% hole; scatter has two value axes.
    const pie = childPath(newChartSpace({ type: 'pie', categories: QUARTERS, series: [TWO[0]!] }, 'rId1'), 'c:chart', 'c:plotArea')!;
    expect(names(pie)).toEqual(['c:layout', 'c:pieChart']);
    expect([val(firstChild(pie, 'c:pieChart'), 'c:varyColors'), val(firstChild(pie, 'c:pieChart'), 'c:firstSliceAng')]).toEqual(['1', '0']);
    const doughnut = childPath(newChartSpace({ type: 'doughnut', categories: QUARTERS, series: TWO }, 'rId1'), 'c:chart', 'c:plotArea', 'c:doughnutChart');
    expect([val(doughnut, 'c:varyColors'), val(doughnut, 'c:firstSliceAng'), val(doughnut, 'c:holeSize')]).toEqual(['1', '0', '75']);
    const scatter = childPath(newChartSpace({ type: 'scatter', categories: [1, 2, 3, 4], series: TWO }, 'rId1'), 'c:chart', 'c:plotArea')!;
    expect(names(scatter)).toEqual(['c:layout', 'c:scatterChart', 'c:valAx', 'c:valAx']);
    expect(elements(scatter, 'c:valAx').map((a) => val(a, 'c:crossBetween'))).toEqual(['midCat', 'midCat']);
    // Areas run from edge to edge; lines sit between the tick marks.
    const crossBetween = (type: NewChartType): string | undefined => val(childPath(newChartSpace({ type, categories: QUARTERS, series: TWO }, 'rId1'), 'c:chart', 'c:plotArea', 'c:valAx'), 'c:crossBetween');
    expect([crossBetween('area'), crossBetween('line'), crossBetween('column')]).toEqual(['midCat', 'between', 'between']);
    expect(names(findAll(scatter, 'c:ser')[0])).toContain('c:xVal');
  });

  it('refers to the workbook cells', () => {
    const cs = newChartSpace({ type: 'column', categories: QUARTERS, series: TWO }, 'rId1');
    const [first, second] = findAll(cs, 'c:ser');
    expect(formula(firstChild(first, 'c:tx'))).toBe('Sheet1!$B$1');
    expect(formula(firstChild(first, 'c:cat'))).toBe('Sheet1!$A$2:$A$5');
    expect(formula(firstChild(first, 'c:val'))).toBe('Sheet1!$B$2:$B$5');
    expect(formula(firstChild(second, 'c:tx'))).toBe('Sheet1!$C$1');
    expect(formula(firstChild(second, 'c:val'))).toBe('Sheet1!$C$2:$C$5');
    // Number categories are numbers in the cache.
    const nums = newChartSpace({ type: 'line', categories: [2024, 2025], series: [{ name: 'A', values: [1, 2] }] }, 'rId1');
    expect(names(firstChild(findAll(nums, 'c:ser')[0], 'c:cat'))).toEqual(['c:numRef']);
  });

  it('names no title and deletes the auto title when there is none', () => {
    const cs = newChartSpace({ type: 'column', categories: QUARTERS, series: [TWO[0]!] }, 'rId1');
    expect(findAll(cs, 'c:title')).toEqual([]);
    expect(val(firstChild(cs, 'c:chart'), 'c:autoTitleDeleted')).toBe('1');
    expect(readChart(cs).title).toBeUndefined();
    const titled = newChartSpace({ type: 'column', title: 'Sales', categories: QUARTERS, series: TWO }, 'rId1');
    expect(val(firstChild(titled, 'c:chart'), 'c:autoTitleDeleted')).toBe('0');
    expect(val(childPath(titled, 'c:chart', 'c:title'), 'c:overlay')).toBe('0');
    expect(childPath(titled, 'c:chart', 'c:title', 'c:tx', 'c:rich')).toBeDefined();
    // No legend when asked for none.
    expect(childPath(newChartSpace({ type: 'column', categories: QUARTERS, series: TWO, legend: 'none' }, 'rId1'), 'c:chart', 'c:legend')).toBeUndefined();
  });

  it('validates the data', () => {
    const cats = ['A', 'B'];
    const one = [{ name: 'S', values: [1, 2] }];
    const fails = (fn: () => void): ChartDataError => {
      try {
        fn();
      } catch (e) {
        expect(e).toBeInstanceOf(ChartDataError);
        return e as ChartDataError;
      }
      throw new Error('expected a ChartDataError');
    };
    expect(fails(() => validateChartData('column', undefined, cats, [{ name: 'Sales', values: [1] }])).message).toBe('series "Sales" has 1 values for 2 categories');
    fails(() => validateChartData('line', 'stacked', cats, one));
    fails(() => validateChartData('pie', undefined, cats, [...one, { name: 'T', values: [3, 4] }]));
    fails(() => validateChartData('scatter', undefined, [1, 'B'], one));
    fails(() => validateChartData('column', undefined, cats, []));
    fails(() => validateChartData('column', undefined, [], [{ name: 'S', values: [] }]));
    fails(() => validateChartData('column', undefined, cats, [{ name: 'S', values: [1, Number.NaN] }]));
    fails(() => validateChartData(undefined, undefined, cats, [{ name: 'S', values: [Number.POSITIVE_INFINITY, 1] }]));
    fails(() => validateChartData('scatter', undefined, [1, Number.NEGATIVE_INFINITY], one));
    fails(() => newChartSpace({ type: 'pie', categories: cats, series: [...one, ...one] }, 'rId1'));

    expect(() => validateChartData('column', 'percent', cats, [...one, { name: 'T', values: [null, 4] }])).not.toThrow();
    expect(() => validateChartData('doughnut', undefined, cats, [...one, ...one])).not.toThrow();
    expect(() => validateChartData('scatter', undefined, [1, 2.5], one)).not.toThrow();
    expect(() => validateChartData(undefined, undefined, cats, one)).not.toThrow();
  });

  it('patches a chart and keeps its formatting', () => {
    const cs = newChartSpace({ type: 'column', title: 'Old', categories: QUARTERS, series: TWO }, 'rId1');
    const s0 = findAll(cs, 'c:ser')[0]!;
    firstChild(s0, 'c:spPr')!.children = [el('a:solidFill', {}, [el('a:srgbClr', { val: '123456' })])];
    const dPt = el('c:dPt', {}, [el('c:idx', { val: 3 }), el('c:invertIfNegative', { val: 0 }), el('c:bubble3D', { val: 0 })]);
    const dLbls = el('c:dLbls', {}, [el('c:dLbl', {}, [el('c:idx', { val: 3 }), el('c:showVal', { val: 1 })]), el('c:showVal', { val: 1 })]);
    s0.children.splice(s0.children.indexOf(firstChild(s0, 'c:cat')!), 0, dPt, dLbls);
    // Bold title text keeps its run formatting.
    const titleRun = findAll(childPath(cs, 'c:chart', 'c:title')!, 'a:rPr')[0]!;
    titleRun.attrs.b = '1';

    patchChartSpace(cs, {
      title: 'New',
      categories: ['A', 'B'],
      series: [
        { name: 'One', values: [1, 2] },
        { name: 'Two', values: [3, 4] },
        { name: 'Three', values: [5, 6] },
      ],
    });
    expect(orderProblems(cs)).toEqual([]);
    const sers = findAll(cs, 'c:ser');
    expect(sers).toHaveLength(3);
    expect(attr(childPath(sers[0], 'c:spPr', 'a:solidFill', 'a:srgbClr'), 'val')).toBe('123456');
    expect(firstChild(sers[0], 'c:dPt')).toBeUndefined();
    expect(findAll(firstChild(sers[0], 'c:dLbls')!, 'c:dLbl')).toEqual([]);
    expect(val(firstChild(sers[0], 'c:dLbls'), 'c:showVal')).toBe('1');
    expect([val(sers[2], 'c:idx'), val(sers[2], 'c:order')]).toEqual(['2', '2']);
    expect(val(childPath(sers[2], 'c:spPr', 'a:solidFill'), 'a:schemeClr')).toBe('accent3');
    for (const s of sers) {
      expect(val(childPath(s, 'c:cat', 'c:strRef', 'c:strCache'), 'c:ptCount')).toBe('2');
      expect(val(childPath(s, 'c:val', 'c:numRef', 'c:numCache'), 'c:ptCount')).toBe('2');
    }
    expect(formula(firstChild(sers[2], 'c:tx'))).toBe('Sheet1!$D$1');
    expect(formula(firstChild(sers[2], 'c:val'))).toBe('Sheet1!$D$2:$D$3');
    expect(formula(firstChild(sers[0], 'c:cat'))).toBe('Sheet1!$A$2:$A$3');
    expect(attr(findAll(childPath(cs, 'c:chart', 'c:title')!, 'a:rPr')[0], 'b')).toBe('1');
    expect(readChartTable(reparse(cs))).toMatchObject({
      kind: 'clustered column',
      title: 'New',
      categories: ['A', 'B'],
      series: [
        { name: 'One', values: [1, 2] },
        { name: 'Two', values: [3, 4] },
        { name: 'Three', values: [5, 6] },
      ],
    });

    // Fewer series: the extra ones go; categories alone keep the series.
    patchChartSpace(cs, { series: [{ name: 'Only', values: [7, null] }] });
    expect(findAll(cs, 'c:ser')).toHaveLength(1);
    expect(readChartTable(cs)).toMatchObject({ categories: ['A', 'B'], series: [{ name: 'Only', values: [7, null] }] });
    expect(() => patchChartSpace(cs, { categories: ['A', 'B', 'C'] })).toThrow(ChartDataError);

    patchChartSpace(cs, { title: '' });
    expect(childPath(cs, 'c:chart', 'c:title')).toBeUndefined();
    expect(val(firstChild(cs, 'c:chart'), 'c:autoTitleDeleted')).toBe('1');
    patchChartSpace(cs, { title: 'Back' });
    expect(readChartTable(cs).title).toBe('Back');
    expect(val(firstChild(cs, 'c:chart'), 'c:autoTitleDeleted')).toBe('0');

    patchChartSpace(cs, { legend: 'none' });
    expect(childPath(cs, 'c:chart', 'c:legend')).toBeUndefined();
    patchChartSpace(cs, { legend: 'top' });
    expect(val(childPath(cs, 'c:chart', 'c:legend'), 'c:legendPos')).toBe('t');

    patchChartSpace(cs, { axisTitles: { value: 'USD' } });
    const valAx = childPath(cs, 'c:chart', 'c:plotArea', 'c:valAx')!;
    expect(textContent(firstChild(valAx, 'c:title')!)).toBe('USD');
    expect(readChartTable(cs).axisTitles).toEqual({ value: 'USD' });
    patchChartSpace(cs, { axisTitles: { value: '', category: 'When' } });
    expect(readChartTable(cs).axisTitles).toEqual({ category: 'When' });
    expect(orderProblems(reparse(cs))).toEqual([]);
  });

  it('styles an added series like the last one, without its per-point formatting, in the next colour', () => {
    const cs = newChartSpace({ type: 'column', categories: QUARTERS, series: TWO }, 'rId1');
    const last = findAll(cs, 'c:ser')[1]!;
    // A half-transparent fill with a matching outline, a highlighted point, a labelled point, a trendline and error bars.
    firstChild(last, 'c:spPr')!.children = parseXml(
      '<c:spPr xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">' +
        '<a:solidFill><a:schemeClr val="accent2"><a:alpha val="50000"/></a:schemeClr></a:solidFill><a:ln w="12700"><a:solidFill><a:schemeClr val="accent2"><a:shade val="50000"/></a:schemeClr></a:solidFill></a:ln></c:spPr>',
    ).root.children;
    const extras = parseXml(
      '<x xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:dPt><c:idx val="1"/><c:invertIfNegative val="0"/><c:bubble3D val="0"/></c:dPt>' +
        '<c:dLbls><c:dLbl><c:idx val="2"/><c:showVal val="1"/></c:dLbl><c:showVal val="1"/></c:dLbls><c:trendline><c:trendlineType val="linear"/></c:trendline>' +
        '<c:errBars><c:errDir val="y"/><c:errBarType val="both"/><c:errValType val="stdErr"/></c:errBars></x>',
    ).root.children;
    last.children.splice(last.children.indexOf(firstChild(last, 'c:cat')!), 0, ...extras);

    patchChartSpace(cs, { series: [...TWO, { name: 'Added', values: [1, 1, 1, 1] }] });
    expect(orderProblems(cs)).toEqual([]);
    const added = findAll(cs, 'c:ser')[2]!;
    expect(names(added).filter((n) => ['c:dPt', 'c:dLbls', 'c:trendline', 'c:errBars'].includes(n))).toEqual(['c:dLbls']);
    // The series' own labels stay; the labels of single points go.
    expect(names(firstChild(added, 'c:dLbls'))).toEqual(['c:showVal']);
    // The next accent, still half transparent; the outline, a darker shade of the fill's colour, follows with its shade.
    const fill = childPath(added, 'c:spPr', 'a:solidFill', 'a:schemeClr')!;
    expect([attr(fill, 'val'), names(fill), val(fill, 'a:alpha')]).toEqual(['accent3', ['a:alpha'], '50000']);
    const outline = childPath(added, 'c:spPr', 'a:ln', 'a:solidFill', 'a:schemeClr')!;
    expect([attr(outline, 'val'), val(outline, 'a:shade')]).toEqual(['accent3', '50000']);
    // The original keeps everything.
    expect(names(last).filter((n) => ['c:dPt', 'c:dLbls', 'c:trendline', 'c:errBars'].includes(n))).toEqual(['c:dPt', 'c:dLbls', 'c:trendline', 'c:errBars']);

    // A filled series whose outline has a colour of its own keeps that outline.
    const own = newChartSpace({ type: 'column', categories: QUARTERS, series: [TWO[0]!] }, 'rId1');
    childPath(findAll(own, 'c:ser')[0], 'c:spPr', 'a:ln')!.children = [el('a:solidFill', {}, [el('a:srgbClr', { val: '000000' })])];
    patchChartSpace(own, { series: [TWO[0]!, TWO[1]!] });
    const second = findAll(own, 'c:ser')[1]!;
    expect(val(childPath(second, 'c:spPr', 'a:solidFill'), 'a:schemeClr')).toBe('accent2');
    expect(val(childPath(second, 'c:spPr', 'a:ln', 'a:solidFill'), 'a:srgbClr')).toBe('000000');
  });

  it('patches multi-level categories, scatter values and line colours', () => {
    const line = newChartSpace({ type: 'line', categories: QUARTERS, series: TWO }, 'rId1');
    const cat = firstChild(findAll(line, 'c:ser')[0]!, 'c:cat')!;
    cat.children = parseXml(
      '<c:cat xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:multiLvlStrRef><c:f>Sheet1!$A$2:$B$5</c:f><c:multiLvlStrCache><c:ptCount val="1"/><c:lvl><c:pt idx="0"><c:v>Jan</c:v></c:pt></c:lvl></c:multiLvlStrCache></c:multiLvlStrRef></c:cat>',
    ).root.children;
    patchChartSpace(line, { categories: ['X', 'Y'], series: [...TWO.map((s) => ({ ...s, values: [1, 2] })), { name: 'West', values: [3, 4] }] });
    expect(orderProblems(line)).toEqual([]);
    expect(names(firstChild(findAll(line, 'c:ser')[0], 'c:cat'))).toEqual(['c:strRef']);
    const added = findAll(line, 'c:ser')[2]!;
    expect(val(childPath(added, 'c:spPr', 'a:ln', 'a:solidFill'), 'a:schemeClr')).toBe('accent3');

    const scatter = newChartSpace({ type: 'scatter', categories: [1, 2, 3], series: TWO.map((s) => ({ ...s, values: [1, 2, 3] })) }, 'rId1');
    patchChartSpace(scatter, { categories: [10, 20], series: [{ name: 'Only', values: [5, 6] }] });
    expect(orderProblems(scatter)).toEqual([]);
    const ser = findAll(scatter, 'c:ser')[0]!;
    expect(names(firstChild(ser, 'c:xVal'))).toEqual(['c:numRef']);
    expect(formula(firstChild(ser, 'c:yVal'))).toBe('Sheet1!$B$2:$B$3');
    expect(readChart(scatter).plots[0]!.series[0]).toMatchObject({ categoryNumbers: [10, 20], values: [5, 6] });
  });
  it('keeps the formulas of a linked chart and gives added series literal data', () => {
    const linked = newChartSpace({ type: 'column', categories: QUARTERS, series: TWO }, 'rId1');
    for (const f of findAll(findAll(linked, 'c:ser')[0]!, 'c:f')) f.children = [text(textContent(f).replace('Sheet1', '[1]Data'))];
    const plain = cloneElement(linked);
    const three: ChartSeriesData[] = [
      { name: 'One', values: [1, 2] },
      { name: 'Two', values: [3, 4] },
      { name: 'Three', values: [5, null] },
    ];

    patchChartSpace(linked, { categories: ['A', 'B'], series: three }, { keepFormulas: true });
    expect(orderProblems(linked)).toEqual([]);
    const sers = findAll(linked, 'c:ser');
    expect(findAll(sers[0]!, 'c:f').map(textContent)).toEqual(['[1]Data!$B$1', '[1]Data!$A$2:$A$5', '[1]Data!$B$2:$B$5']);
    expect(findAll(sers[1]!, 'c:f').map(textContent)).toEqual(['Sheet1!$C$1', 'Sheet1!$A$2:$A$5', 'Sheet1!$C$2:$C$5']);
    expect(val(childPath(sers[0], 'c:val', 'c:numRef', 'c:numCache'), 'c:ptCount')).toBe('2');
    const added = sers[2]!;
    expect(findAll(added, 'c:f')).toEqual([]);
    expect(textContent(childPath(added, 'c:tx', 'c:v')!)).toBe('Three');
    expect(names(firstChild(added, 'c:cat'))).toEqual(['c:strLit']);
    expect(names(firstChild(added, 'c:val'))).toEqual(['c:numLit']);
    expect(findAll(firstChild(added, 'c:val')!, 'c:pt').map((p) => attr(p, 'idx'))).toEqual(['0']);
    expect(readChartTable(reparse(linked))).toMatchObject({ categories: ['A', 'B'], series: three });

    // Literal data stays literal; references keep their formulas when the categories become numbers.
    patchChartSpace(linked, { categories: [2024, 2025] }, { keepFormulas: true });
    expect(orderProblems(linked)).toEqual([]);
    expect(names(firstChild(findAll(linked, 'c:ser')[2], 'c:cat'))).toEqual(['c:numLit']);
    expect(findAll(findAll(linked, 'c:ser')[2]!, 'c:f')).toEqual([]);
    expect(formula(firstChild(findAll(linked, 'c:ser')[0], 'c:cat'))).toBe('[1]Data!$A$2:$A$5');
    expect(readChartTable(linked).categories).toEqual([2024, 2025]);

    // Without the option every formula points into Sheet1 of the chart's own workbook.
    patchChartSpace(plain, { categories: ['A', 'B'], series: three });
    expect(findAll(findAll(plain, 'c:ser')[0]!, 'c:f').map(textContent)).toEqual(['Sheet1!$B$1', 'Sheet1!$A$2:$A$3', 'Sheet1!$B$2:$B$3']);
    expect(findAll(findAll(plain, 'c:ser')[2]!, 'c:f').map(textContent)).toEqual(['Sheet1!$D$1', 'Sheet1!$A$2:$A$3', 'Sheet1!$D$2:$D$3']);

    // Scatter: an added series gets literal x and y values.
    const scatter = newChartSpace({ type: 'scatter', categories: [1, 2], series: [{ name: 'A', values: [1, 2] }] }, 'rId1');
    patchChartSpace(scatter, { series: [{ name: 'A', values: [3, 4] }, { name: 'B', values: [5, 6] }] }, { keepFormulas: true });
    expect(orderProblems(scatter)).toEqual([]);
    const [a, b] = findAll(scatter, 'c:ser');
    expect(formula(firstChild(a, 'c:yVal'))).toBe('Sheet1!$B$2:$B$3');
    expect([names(firstChild(b, 'c:xVal')), names(firstChild(b, 'c:yVal'))]).toEqual([['c:numLit'], ['c:numLit']]);
    expect(findAll(b!, 'c:f')).toEqual([]);
    expect(readChart(scatter).plots[0]!.series[1]).toMatchObject({ name: 'B', categoryNumbers: [1, 2], values: [5, 6] });
  });

  it('removes plots and axes left without series', () => {
    /** Columns North and South, plus a line plot with Trend on the same axes or its own secondary ones. */
    const combo = (secondary: boolean): XmlElement => {
      const cs = newChartSpace({ type: 'column', categories: QUARTERS, series: TWO }, 'rId1');
      const lineArea = childPath(newChartSpace({ type: 'line', categories: QUARTERS, series: [...TWO, { name: 'Trend', values: [2, 3, 4, 5] }] }, 'rId1'), 'c:chart', 'c:plotArea')!;
      const line = firstChild(lineArea, 'c:lineChart')!;
      line.children = line.children.filter((c) => !(c.type === 'element' && c.name === 'c:ser' && val(c, 'c:idx') !== '2'));
      const plotArea = childPath(cs, 'c:chart', 'c:plotArea')!;
      plotArea.children.splice(2, 0, line);
      if (secondary) {
        const [catId, valId] = elements(line, 'c:axId');
        catId!.attrs.val = '3';
        valId!.attrs.val = '4';
        const cat = firstChild(lineArea, 'c:catAx')!;
        const value = firstChild(lineArea, 'c:valAx')!;
        firstChild(cat, 'c:axId')!.attrs.val = '3';
        firstChild(cat, 'c:crossAx')!.attrs.val = '4';
        firstChild(cat, 'c:delete')!.attrs.val = '1';
        firstChild(value, 'c:axId')!.attrs.val = '4';
        firstChild(value, 'c:crossAx')!.attrs.val = '3';
        firstChild(value, 'c:axPos')!.attrs.val = 'r';
        plotArea.children.push(cat, value);
      }
      return cs;
    };
    const plotNames = (cs: XmlElement): string[] => names(childPath(cs, 'c:chart', 'c:plotArea'));

    for (const secondary of [false, true]) {
      const cs = combo(secondary);
      const label = secondary ? 'secondary axes' : 'shared axes';
      expect(orderProblems(cs), label).toEqual([]);
      expect(readChartTable(cs), label).toMatchObject({ kind: 'clustered column + line', series: [{ name: 'North' }, { name: 'South' }, { name: 'Trend' }] });
      const before = plotNames(cs);
      expect(before, label).toEqual(['c:layout', 'c:barChart', 'c:lineChart', 'c:catAx', 'c:valAx', ...(secondary ? ['c:catAx', 'c:valAx'] : [])]);

      // Every plot keeps a series: nothing goes.
      patchChartSpace(cs, { series: [...TWO, { name: 'Trend', values: [1, 1, 1, 1] }] });
      expect(plotNames(cs), label).toEqual(before);

      patchChartSpace(cs, { series: TWO });
      expect(plotNames(cs), label).toEqual(['c:layout', 'c:barChart', 'c:catAx', 'c:valAx']);
      const plotArea = childPath(cs, 'c:chart', 'c:plotArea')!;
      expect([...elements(plotArea, 'c:catAx'), ...elements(plotArea, 'c:valAx')].map((a) => val(a, 'c:axId')), label).toEqual(['100000001', '100000002']);
      expect(readChartTable(cs), label).toMatchObject({ kind: 'clustered column', series: TWO });
      expect(orderProblems(reparse(cs)), label).toEqual([]);
    }
  });
});
