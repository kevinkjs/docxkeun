/** Charts: reading c:chartSpace parts into a model, describing them, and drawing them. */
import { describe, expect, it } from 'vitest';
import { autoColor, describeChart, formatNumber, readChart, renderChart, resolveColorElement } from '../src/charts/index.js';
import { RegistryFontProvider } from '../src/layout/fonts.js';
import type { DisplayItem, PathCommand, PathItem, TextItem } from '../src/layout/types.js';
import { parseXml } from '../src/xml/parser.js';
import { testRegistry } from './helpers/fonts.js';

const fonts = new RegistryFontProvider(testRegistry());
const THEME: Record<string, string> = {
  dk1: '000000',
  lt1: 'FFFFFF',
  dk2: '44546A',
  lt2: 'E7E6E6',
  accent1: '4472C4',
  accent2: 'ED7D31',
  accent3: 'A5A5A5',
  accent4: 'FFC000',
  accent5: '5B9BD5',
  accent6: '70AD47',
};
const schemeColor = (name: string): string | undefined => THEME[name];
const FRAME = { x: 50, y: 100, w: 432, h: 252 };

// ---------------------------------------------------------------------------------------------
// XML builders

const NS = 'xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';

function strRef(values: string[]): string {
  return `<c:strRef><c:f>Sheet1!$A$2</c:f><c:strCache><c:ptCount val="${values.length}"/>${values.map((v, i) => `<c:pt idx="${i}"><c:v>${v}</c:v></c:pt>`).join('')}</c:strCache></c:strRef>`;
}

function numRef(values: Array<number | null>, format = 'General'): string {
  const pts = values.map((v, i) => (v === null ? '' : `<c:pt idx="${i}"><c:v>${v}</c:v></c:pt>`)).join('');
  return `<c:numRef><c:f>Sheet1!$B$2</c:f><c:numCache><c:formatCode>${format}</c:formatCode><c:ptCount val="${values.length}"/>${pts}</c:numCache></c:numRef>`;
}

function solid(color: string): string {
  return `<a:solidFill><a:srgbClr val="${color}"/></a:solidFill>`;
}

function ser(idx: number, name: string, cats: string[], values: Array<number | null>, extra = '', format = 'General'): string {
  return `<c:ser><c:idx val="${idx}"/><c:order val="${idx}"/><c:tx>${strRef([name])}</c:tx>${extra}<c:cat>${strRef(cats)}</c:cat><c:val>${numRef(values, format)}</c:val></c:ser>`;
}

function catAx(id: number, cross: number, pos = 'b', extra = ''): string {
  return `<c:catAx><c:axId val="${id}"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="${pos}"/><c:numFmt formatCode="General" sourceLinked="1"/><c:majorTickMark val="out"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/><c:crossAx val="${cross}"/><c:crosses val="autoZero"/>${extra}</c:catAx>`;
}

function valAx(id: number, cross: number, pos = 'l', extra = '', fmt = '<c:numFmt formatCode="General" sourceLinked="1"/>'): string {
  return `<c:valAx><c:axId val="${id}"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="${pos}"/><c:majorGridlines/>${fmt}<c:majorTickMark val="out"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/><c:crossAx val="${cross}"/><c:crosses val="autoZero"/><c:crossBetween val="between"/>${extra}</c:valAx>`;
}

function title(text: string): string {
  return `<c:title><c:tx><c:rich><a:bodyPr/><a:p><a:pPr><a:defRPr sz="1400" b="0"/></a:pPr><a:r><a:rPr lang="en-US"/><a:t>${text}</a:t></a:r></a:p></c:rich></c:tx><c:overlay val="0"/></c:title>`;
}

function chartSpace(plotArea: string, opts: { title?: string; legend?: string; autoTitleDeleted?: boolean } = {}) {
  const xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<c:chartSpace ${NS}><c:roundedCorners val="0"/><c:chart>${opts.title ?? ''}<c:autoTitleDeleted val="${opts.autoTitleDeleted === false ? 0 : 1}"/><c:plotArea><c:layout/>${plotArea}</c:plotArea>${opts.legend !== undefined ? `<c:legend><c:legendPos val="${opts.legend}"/><c:overlay val="0"/></c:legend>` : ''}<c:plotVisOnly val="1"/><c:dispBlanksAs val="gap"/></c:chart></c:chartSpace>`;
  return parseXml(xml).root;
}

// ---------------------------------------------------------------------------------------------
// Display-list helpers

const paths = (items: DisplayItem[]): PathItem[] => items.filter((i): i is PathItem => i.type === 'path');
const texts = (items: DisplayItem[]): TextItem[] => items.filter((i): i is TextItem => i.type === 'text');
const textOf = (items: DisplayItem[]): string[] => texts(items).map((t) => t.text);

/** Axis-aligned rectangle paths (M + 3 L + Z) with this fill. */
function rects(items: DisplayItem[], fill: string): Array<{ x: number; y: number; w: number; h: number }> {
  return paths(items)
    .filter((p) => p.fill === fill && p.commands.length === 5 && p.commands[4]!.op === 'Z' && p.commands.slice(1, 4).every((c) => c.op === 'L'))
    .map((p) => {
      const pts = p.commands.filter((c): c is Extract<PathCommand, { op: 'M' | 'L' }> => c.op === 'M' || c.op === 'L');
      const xs = pts.map((c) => c.x);
      const ys = pts.map((c) => c.y);
      return { x: Math.min(...xs), y: Math.min(...ys), w: Math.max(...xs) - Math.min(...xs), h: Math.max(...ys) - Math.min(...ys) };
    });
}

function bbox(p: PathItem): { x1: number; y1: number; x2: number; y2: number } {
  const xs: number[] = [];
  const ys: number[] = [];
  for (const c of p.commands) {
    if (c.op === 'Z') continue;
    xs.push(c.x);
    ys.push(c.y);
  }
  return { x1: Math.min(...xs), y1: Math.min(...ys), x2: Math.max(...xs), y2: Math.max(...ys) };
}

// ---------------------------------------------------------------------------------------------

const QUARTERS = ['Q1', 'Q2', 'Q3', 'Q4'];

function clusteredColumn(extraValAx = '', plotLayout = '') {
  return chartSpace(
    `${plotLayout}<c:barChart><c:barDir val="col"/><c:grouping val="clustered"/><c:varyColors val="0"/>` +
      ser(0, 'Revenue', QUARTERS, [120, 150, 90, 180], `<c:spPr>${solid('4472C4')}</c:spPr>`) +
      ser(1, 'Costs', QUARTERS, [80, 95, 70, 110]) +
      `<c:gapWidth val="150"/><c:axId val="1"/><c:axId val="2"/></c:barChart>${catAx(1, 2)}${valAx(2, 1, 'l', '', extraValAx || undefined)}`,
    { title: title('Revenue by quarter'), legend: 'r' },
  );
}

describe('chart model', () => {
  it('reads a clustered column chart', () => {
    const model = readChart(clusteredColumn());
    expect(model.title?.text).toBe('Revenue by quarter');
    expect(model.plots).toHaveLength(1);
    const plot = model.plots[0]!;
    expect(plot).toMatchObject({ type: 'bar', barDir: 'col', grouping: 'clustered', gapWidth: 150, overlap: 0 });
    expect(plot.series.map((s) => s.name)).toEqual(['Revenue', 'Costs']);
    expect(plot.series[0]!.categories).toEqual(QUARTERS);
    expect(plot.series[1]!.values).toEqual([80, 95, 70, 110]);
    // Explicit spPr colour, then the accent palette by series index.
    expect(plot.series[0]!.color).toBe('4472C4');
    expect(plot.series[1]!.color).toBe('ED7D31');
    expect(model.legend?.position).toBe('r');
    expect(model.axes.map((a) => [a.kind, a.position])).toEqual([
      ['cat', 'b'],
      ['val', 'l'],
    ]);
    expect(model.axes[1]!.majorGridlines).toBeDefined();
  });

  it('describes a chart in one line', () => {
    expect(describeChart(readChart(clusteredColumn()))).toBe('Clustered column chart "Revenue by quarter": series Revenue (Q1 120, Q2 150, Q3 90, Q4 180), Costs (Q1 80, Q2 95, Q3 70, Q4 110)');
  });

  it('reads literals, multi-level categories, blanks and the automatic title', () => {
    const xml = chartSpace(
      `<c:lineChart><c:grouping val="standard"/><c:varyColors val="0"/><c:ser><c:idx val="0"/><c:order val="0"/><c:tx><c:v>Visitors</c:v></c:tx>` +
        `<c:cat><c:multiLvlStrRef><c:f>x</c:f><c:multiLvlStrCache><c:ptCount val="4"/><c:lvl><c:pt idx="0"><c:v>Jan</c:v></c:pt><c:pt idx="1"><c:v>Feb</c:v></c:pt><c:pt idx="2"><c:v>Jan</c:v></c:pt><c:pt idx="3"><c:v>Feb</c:v></c:pt></c:lvl>` +
        `<c:lvl><c:pt idx="0"><c:v>2023</c:v></c:pt><c:pt idx="2"><c:v>2024</c:v></c:pt></c:lvl></c:multiLvlStrCache></c:multiLvlStrRef></c:cat>` +
        `<c:val><c:numLit><c:formatCode>#,##0</c:formatCode><c:ptCount val="4"/><c:pt idx="0"><c:v>1200</c:v></c:pt><c:pt idx="1"><c:v>1500</c:v></c:pt><c:pt idx="3"><c:v>2100</c:v></c:pt></c:numLit></c:val></c:ser>` +
        `<c:marker val="1"/><c:axId val="1"/><c:axId val="2"/></c:lineChart>${catAx(1, 2)}${valAx(2, 1)}`,
      { autoTitleDeleted: false },
    );
    const model = readChart(xml);
    const s = model.plots[0]!.series[0]!;
    expect(s.name).toBe('Visitors');
    expect(s.categories).toEqual(['Jan', 'Feb', 'Jan', 'Feb']);
    expect(s.categoryGroups).toEqual([['2023', undefined, '2024']]);
    expect(s.values).toEqual([1200, 1500, null, 2100]);
    expect(s.formatCode).toBe('#,##0');
    // A single series without a written title shows its name as the title.
    expect(model.title).toMatchObject({ text: 'Visitors', auto: true });
    expect(describeChart(model)).toBe('Line chart "Visitors": series Visitors (Jan 1,200, Feb 1,500, Jan blank, Feb 2,100)');
  });
});

describe('chart colours', () => {
  const clr = (xml: string): string | undefined => resolveColorElement(parseXml(`<a:x xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">${xml}</a:x>`).root.children.find((c) => c.type === 'element') as never, schemeColor)?.color;

  it('applies lumMod/lumOff in HSL as Office does', () => {
    expect(clr('<a:schemeClr val="accent1"><a:lumMod val="75000"/></a:schemeClr>')).toBe('2F5597');
    expect(clr('<a:schemeClr val="accent1"><a:lumMod val="60000"/><a:lumOff val="40000"/></a:schemeClr>')).toBe('8FAADC');
    expect(clr('<a:schemeClr val="tx1"><a:lumMod val="65000"/><a:lumOff val="35000"/></a:schemeClr>')).toBe('595959');
    expect(clr('<a:schemeClr val="bg1"><a:lumMod val="85000"/></a:schemeClr>')).toBe('D9D9D9');
  });

  it('applies tint and shade in linear light, and alpha', () => {
    expect(clr('<a:srgbClr val="FF0000"><a:shade val="50000"/></a:srgbClr>')).toBe('BC0000');
    expect(clr('<a:srgbClr val="000000"><a:tint val="50000"/></a:srgbClr>')).toBe('BCBCBC');
    const paint = resolveColorElement(parseXml(`<a:srgbClr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" val="00FF00"><a:alpha val="40000"/></a:srgbClr>`).root, schemeColor);
    expect(paint).toEqual({ color: '00FF00', opacity: 0.4 });
  });

  it('cycles the accents, then darker and lighter variants for series 7+', () => {
    expect([0, 1, 2, 3, 4, 5].map((i) => autoColor(i, 12, 2, schemeColor))).toEqual(['4472C4', 'ED7D31', 'A5A5A5', 'FFC000', '5B9BD5', '70AD47']);
    expect(autoColor(6, 12, 2, schemeColor)).toBe('264478');
    expect(autoColor(12, 18, 2, schemeColor)).toBe(clr('<a:schemeClr val="accent1"><a:lumMod val="80000"/><a:lumOff val="20000"/></a:schemeClr>'));
  });
});

describe('number formats', () => {
  it('formats the common codes', () => {
    expect(formatNumber(0.1 + 0.2, 'General')).toBe('0.3');
    expect(formatNumber(1234567, undefined)).toBe('1234567');
    expect(formatNumber(1.5e12, 'General')).toBe('1.5E+12');
    expect(formatNumber(3.456, '0')).toBe('3');
    expect(formatNumber(3.45, '0.0')).toBe('3.5');
    expect(formatNumber(1.005, '0.00')).toBe('1.01');
    expect(formatNumber(1234.5, '#,##0')).toBe('1,235');
    expect(formatNumber(1234.5, '#,##0.00')).toBe('1,234.50');
    expect(formatNumber(-1234.5, '#,##0.00')).toBe('-1,234.50');
    expect(formatNumber(0.256, '0%')).toBe('26%');
    expect(formatNumber(0.256, '0.0%')).toBe('25.6%');
    expect(formatNumber(1234, '$#,##0')).toBe('$1,234');
    expect(formatNumber(-1234, '$#,##0;($#,##0)')).toBe('($1,234)');
    expect(formatNumber(0, '#,##0;-#,##0;"zero"')).toBe('zero');
    expect(formatNumber(1500, '[$€-407] #,##0')).toBe('€ 1,500');
    expect(formatNumber(12, '0 "pcs"')).toBe('12 pcs');
    expect(formatNumber(12345678, '#,##0,,"M"')).toBe('12M');
    expect(formatNumber(12345, '#,##0.0,"K"')).toBe('12.3K');
    expect(formatNumber(-5, '[Red]0.0')).toBe('-5.0');
    expect(formatNumber(0.5, '#.00')).toBe('.50');
    expect(formatNumber(12345, '0.00E+00')).toBe('1.23E+04');
    // "_(" pads like a character, "* " (fill) has no width outside a cell.
    expect(formatNumber(5, '_($* #,##0_)')).toBe(' $5 ');
  });

  it('formats date serials', () => {
    expect(formatNumber(44927, 'mmm-yy')).toBe('Jan-23');
    expect(formatNumber(44927, 'm/d/yyyy')).toBe('1/1/2023');
    expect(formatNumber(44927.75, 'h:mm AM/PM')).toBe('6:00 PM');
    expect(formatNumber(45000, 'dd mmmm yyyy')).toBe('15 March 2023');
  });
});

describe('chart drawing', () => {
  it('draws a clustered column chart with a title and a right legend', () => {
    const items = renderChart(clusteredColumn(), FRAME, { fonts, schemeColor });
    const revenue = rects(items, '4472C4');
    const costs = rects(items, 'ED7D31');
    // Four bars per series plus a legend swatch each.
    const bars = (r: typeof revenue) => r.filter((b) => b.w > 12);
    expect(bars(revenue)).toHaveLength(4);
    expect(bars(costs)).toHaveLength(4);
    expect(revenue.filter((b) => b.w <= 12)).toHaveLength(1);

    // Heights are proportional to the values and share a baseline.
    const all = [...bars(revenue), ...bars(costs)];
    const values = [120, 150, 90, 180, 80, 95, 70, 110];
    const ratio = all[0]!.h / values[0]!;
    all.forEach((b, i) => expect(b.h / values[i]!).toBeCloseTo(ratio, 6));
    const baseline = all[0]!.y + all[0]!.h;
    for (const b of all) expect(b.y + b.h).toBeCloseTo(baseline, 6);
    // Series sit side by side within each category, left to right.
    for (let i = 0; i < 4; i++) expect(bars(costs)[i]!.x).toBeCloseTo(bars(revenue)[i]!.x + bars(revenue)[i]!.w, 6);

    const t = texts(items);
    const strings = textOf(items);
    expect(strings).toContain('Revenue by quarter');
    for (const q of QUARTERS) expect(strings).toContain(q);
    // Legend to the right of every bar.
    const right = Math.max(...all.map((b) => b.x + b.w));
    const legendRevenue = t.find((x) => x.text === 'Revenue')!;
    const legendCosts = t.find((x) => x.text === 'Costs')!;
    expect(legendRevenue.x).toBeGreaterThan(right);
    expect(legendCosts.y).toBeGreaterThan(legendRevenue.y);
    // Title centred at the top.
    const titleItem = t.find((x) => x.text === 'Revenue by quarter')!;
    const titleW = titleItem.widths.reduce((a, b) => a + b, 0);
    expect(titleItem.x + titleW / 2).toBeCloseTo(FRAME.x + FRAME.w / 2, 0);
    expect(titleItem.y).toBeLessThan(Math.min(...all.map((b) => b.y)));

    // Automatic value axis: 0 to a round maximum above 180, even steps.
    const axisLabels = t.filter((x) => /^\d+$/.test(x.text) && x.x < Math.min(...all.map((b) => b.x))).map((x) => Number(x.text));
    expect(axisLabels[0]).toBe(0);
    expect(Math.max(...axisLabels)).toBeGreaterThanOrEqual(180);
    const step = axisLabels[1]! - axisLabels[0]!;
    expect([10, 20, 25, 50]).toContain(step);
    axisLabels.forEach((v, i) => expect(v).toBe(i * step));
    // The bar for 180 reaches the 180 level of the axis scale.
    const zeroLabel = t.find((x) => x.text === '0' && x.x < right)!;
    const topLabel = t.find((x) => x.text === String(Math.max(...axisLabels)))!;
    const pxPerUnit = (zeroLabel.y - topLabel.y) / Math.max(...axisLabels);
    expect(all[3]!.h).toBeCloseTo(180 * pxPerUnit, 4);

    // Gridlines (default colour) span the plot.
    expect(paths(items).filter((p) => p.stroke === 'D9D9D9').length).toBe(axisLabels.length);
  });

  it('formats value axis labels with the axis number format', () => {
    const items = renderChart(clusteredColumn('<c:numFmt formatCode="$#,##0" sourceLinked="0"/>'), FRAME, { fonts, schemeColor });
    const labels = textOf(items).filter((s) => s.startsWith('$'));
    expect(labels[0]).toBe('$0');
    expect(labels.length).toBeGreaterThan(3);
    expect(labels.every((l) => /^\$\d{1,3}(,\d{3})*$/.test(l))).toBe(true);
  });

  it('honours a manual inner plot layout', () => {
    const layout = '<c:layout><c:manualLayout><c:layoutTarget val="inner"/><c:xMode val="edge"/><c:yMode val="edge"/><c:x val="0.1"/><c:y val="0.2"/><c:w val="0.6"/><c:h val="0.6"/></c:manualLayout></c:layout>';
    const cs = clusteredColumn('', '');
    // Put the manual layout on the plot area itself.
    const xml = parseXml(
      `<c:chartSpace ${NS}><c:chart><c:autoTitleDeleted val="1"/><c:plotArea>${layout}<c:barChart><c:barDir val="col"/><c:grouping val="clustered"/>${ser(0, 'A', QUARTERS, [1, 2, 3, 4])}<c:axId val="1"/><c:axId val="2"/></c:barChart>${catAx(1, 2)}${valAx(2, 1)}</c:plotArea></c:chart></c:chartSpace>`,
    ).root;
    void cs;
    const items = renderChart(xml, FRAME, { fonts, schemeColor });
    const grid = paths(items).filter((p) => p.stroke === 'D9D9D9').map(bbox);
    for (const g of grid) {
      expect(g.x1).toBeCloseTo(FRAME.x + 0.1 * FRAME.w, 6);
      expect(g.x2).toBeCloseTo(FRAME.x + 0.7 * FRAME.w, 6);
    }
    const ys = grid.map((g) => g.y1);
    expect(Math.min(...ys)).toBeCloseTo(FRAME.y + 0.2 * FRAME.h, 6);
    expect(Math.max(...ys)).toBeCloseTo(FRAME.y + 0.8 * FRAME.h, 6);
  });

  it('draws a stacked bar chart with bars end to end', () => {
    const xml = chartSpace(
      `<c:barChart><c:barDir val="bar"/><c:grouping val="stacked"/><c:varyColors val="0"/>` +
        ser(0, 'North', ['Apples', 'Pears', 'Plums'], [30, 20, 10]) +
        ser(1, 'South', ['Apples', 'Pears', 'Plums'], [10, 25, 40]) +
        `<c:gapWidth val="100"/><c:overlap val="100"/><c:axId val="1"/><c:axId val="2"/></c:barChart>${catAx(1, 2, 'l')}${valAx(2, 1, 'b')}`,
      { legend: 'b' },
    );
    const model = readChart(xml);
    expect(model.plots[0]).toMatchObject({ barDir: 'bar', grouping: 'stacked', overlap: 100 });
    expect(describeChart(model)).toMatch(/^Stacked bar chart: series North \(Apples 30, Pears 20, Plums 10\)/);

    const items = renderChart(xml, FRAME, { fonts, schemeColor });
    const north = rects(items, '4472C4').filter((b) => b.h > 10);
    const south = rects(items, 'ED7D31').filter((b) => b.h > 10);
    expect(north).toHaveLength(3);
    expect(south).toHaveLength(3);
    // First category at the bottom; each South bar starts where the North bar ends.
    expect(north[0]!.y).toBeGreaterThan(north[2]!.y);
    const per = north[0]!.w / 30;
    for (let i = 0; i < 3; i++) {
      expect(south[i]!.x).toBeCloseTo(north[i]!.x + north[i]!.w, 6);
      expect(south[i]!.y).toBeCloseTo(north[i]!.y, 6);
      expect(north[i]!.w).toBeCloseTo([30, 20, 10][i]! * per, 6);
      expect(south[i]!.w).toBeCloseTo([10, 25, 40][i]! * per, 6);
    }
    // Category labels to the left of the bars, legend below them.
    const t = texts(items);
    for (const c of ['Apples', 'Pears', 'Plums']) expect(t.find((x) => x.text === c)!.x).toBeLessThan(north[0]!.x);
    expect(t.find((x) => x.text === 'North')!.y).toBeGreaterThan(Math.max(...north.map((b) => b.y + b.h)));
  });

  it('draws a line chart with markers and a smoothed series', () => {
    const marker = (sym: string, size: number) => `<c:marker><c:symbol val="${sym}"/><c:size val="${size}"/></c:marker>`;
    const xml = chartSpace(
      `<c:lineChart><c:grouping val="standard"/><c:varyColors val="0"/>` +
        ser(0, 'Plan', QUARTERS, [10, 30, 20, 40], `<c:spPr><a:ln w="28575" cap="rnd">${solid('70AD47')}</a:ln></c:spPr>${marker('circle', 7)}`) +
        ser(1, 'Actual', QUARTERS, [15, 25, null, 35], marker('square', 5)).replace('</c:val>', '</c:val><c:smooth val="1"/>') +
        `<c:marker val="1"/><c:axId val="1"/><c:axId val="2"/></c:lineChart>${catAx(1, 2)}${valAx(2, 1)}`,
      { legend: 't' },
    );
    const model = readChart(xml);
    const [plan, actual] = model.plots[0]!.series;
    expect(plan!.marker).toMatchObject({ symbol: 'circle', size: 7 });
    expect(plan!.line).toMatchObject({ width: 2.25, paint: { color: '70AD47' } });
    expect(actual!.smooth).toBe(true);
    expect(actual!.color).toBe('ED7D31');

    const items = renderChart(xml, FRAME, { fonts, schemeColor });
    const p = paths(items);
    const planLine = p.filter((x) => x.stroke === '70AD47' && !x.fill && x.commands.length === 4);
    expect(planLine).toHaveLength(1);
    expect(planLine[0]!.strokeWidth).toBe(2.25);
    expect(planLine[0]!.commands.map((c) => c.op)).toEqual(['M', 'L', 'L', 'L']);
    // Circle markers: 4 points on the line plus one in the legend key.
    const circles = p.filter((x) => x.fill === '70AD47' && x.commands.filter((c) => c.op === 'C').length === 4);
    expect(circles).toHaveLength(5);
    // The markers sit on the line's vertices.
    const vertices = planLine[0]!.commands as Array<{ x: number; y: number }>;
    for (const v of vertices) expect(circles.some((c) => Math.abs((bbox(c).x1 + bbox(c).x2) / 2 - v.x) < 1e-6 && Math.abs((bbox(c).y1 + bbox(c).y2) / 2 - v.y) < 1e-6)).toBe(true);
    // The smoothed series breaks at the blank: one Bézier run of two points; three square markers (+1 legend).
    const actualLines = p.filter((x) => x.stroke === 'ED7D31' && !x.fill && x.commands.length > 1 && x.strokeWidth === 2.25);
    expect(actualLines.some((x) => x.commands.some((c) => c.op === 'C'))).toBe(true);
    expect(rects(items, 'ED7D31').filter((r) => Math.abs(r.w - 5) < 1e-6)).toHaveLength(4);
    // Higher values are drawn higher up.
    expect(vertices[3]!.y).toBeLessThan(vertices[0]!.y);
    expect(textOf(items)).toEqual(expect.arrayContaining(['Plan', 'Actual']));
  });

  it('draws a pie chart with percentage labels', () => {
    const values = [10, 20, 30, 40];
    const xml = chartSpace(
      `<c:pieChart><c:varyColors val="1"/>` +
        ser(0, 'Share', ['North', 'South', 'East', 'West'], values, '', 'General').replace('<c:cat>', '<c:dLbls><c:showLegendKey val="0"/><c:showVal val="0"/><c:showCatName val="0"/><c:showSerName val="0"/><c:showPercent val="1"/><c:showBubbleSize val="0"/></c:dLbls><c:cat>') +
        `<c:firstSliceAng val="0"/></c:pieChart>`,
      { legend: 'r', autoTitleDeleted: false },
    );
    const model = readChart(xml);
    const s = model.plots[0]!.series[0]!;
    expect(s.pointColors?.slice(0, 4)).toEqual(['4472C4', 'ED7D31', 'A5A5A5', 'FFC000']);
    expect(model.title?.text).toBe('Share');
    expect(describeChart(model)).toBe('Pie chart "Share": series Share (North 10 (10%), South 20 (20%), East 30 (30%), West 40 (40%))');

    const items = renderChart(xml, FRAME, { fonts, schemeColor });
    const slices = paths(items).filter((p) => p.commands[0]?.op === 'M' && p.commands[1]?.op === 'L' && p.commands.some((c) => c.op === 'C'));
    expect(slices).toHaveLength(4);
    expect(slices.map((p) => p.fill)).toEqual(['4472C4', 'ED7D31', 'A5A5A5', 'FFC000']);
    // Each slice: centre, edge start, arc to edge end. Sweeps sum to 360° and follow the values.
    const centre = slices[0]!.commands[0] as { x: number; y: number };
    const angleOf = (pt: { x: number; y: number }) => (Math.atan2(pt.x - centre.x, -(pt.y - centre.y)) * 180) / Math.PI;
    let total = 0;
    slices.forEach((p, i) => {
      const start = p.commands[1] as { x: number; y: number };
      const end = p.commands[p.commands.length - 2] as { x: number; y: number };
      const sweep = (((angleOf(end) - angleOf(start)) % 360) + 360) % 360;
      expect(sweep).toBeCloseTo((values[i]! / 100) * 360, 6);
      total += sweep;
    });
    expect(total).toBeCloseTo(360, 6);
    // First slice starts at 12 o'clock.
    expect(angleOf(slices[0]!.commands[1] as { x: number; y: number })).toBeCloseTo(0, 6);
    const strings = textOf(items);
    expect(strings).toEqual(expect.arrayContaining(['10%', '20%', '30%', '40%', 'North', 'South', 'East', 'West', 'Share']));
  });

  it('draws a doughnut with one ring per series', () => {
    const xml = chartSpace(
      `<c:doughnutChart><c:varyColors val="1"/>` +
        ser(0, '2023', ['A', 'B', 'C'], [1, 1, 2]) +
        ser(1, '2024', ['A', 'B', 'C'], [2, 1, 1]) +
        `<c:firstSliceAng val="90"/><c:holeSize val="50"/></c:doughnutChart>`,
      { legend: 'b' },
    );
    const model = readChart(xml);
    expect(model.plots[0]).toMatchObject({ type: 'doughnut', holeSize: 50, firstSliceAng: 90 });
    const items = renderChart(xml, FRAME, { fonts, schemeColor });
    const rings = paths(items).filter((p) => p.commands.filter((c) => c.op === 'L').length === 1 && p.commands.some((c) => c.op === 'C') && p.commands[p.commands.length - 1]!.op === 'Z');
    expect(rings).toHaveLength(6);
    const first = rings[0]!.commands[0] as { x: number; y: number };
    // Ring 1 (first series) is inside ring 2; the hole is half the outer diameter.
    const radii = rings.map((p) => {
      const pts = p.commands.filter((c): c is Extract<PathCommand, { op: 'L' | 'M' }> => c.op === 'M' || c.op === 'L');
      return pts;
    });
    expect(radii).toHaveLength(6);
    const bb = paths(items).filter((p) => rings.includes(p)).map(bbox);
    const outer = Math.max(...bb.map((b) => b.x2 - b.x1)) / 2;
    const cx = (bb[0]!.x1 + bb[0]!.x2) / 2;
    // The first slice of the inner ring starts at 3 o'clock (firstSliceAng 90) on the hole's edge.
    expect(first.y).toBeCloseTo(bb.reduce((a, b) => Math.max(a, b.y2), 0) - outer, 0);
    expect(first.x - cx).toBeGreaterThan(0);
    expect(textOf(items)).toEqual(expect.arrayContaining(['A', 'B', 'C']));
  });

  it('draws a scatter chart on two value axes', () => {
    const xml = chartSpace(
      `<c:scatterChart><c:scatterStyle val="lineMarker"/><c:varyColors val="0"/>` +
        `<c:ser><c:idx val="0"/><c:order val="0"/><c:tx>${strRef(['Height'])}</c:tx><c:spPr><a:ln w="19050"><a:noFill/></a:ln></c:spPr><c:marker><c:symbol val="diamond"/><c:size val="6"/></c:marker>` +
        `<c:xVal>${numRef([1, 2, 4, 8])}</c:xVal><c:yVal>${numRef([3, 5, 4, 9], '0.0')}</c:yVal><c:smooth val="0"/></c:ser>` +
        `<c:axId val="1"/><c:axId val="2"/></c:scatterChart>${valAx(1, 2, 'b')}${valAx(2, 1, 'l')}`,
    );
    const model = readChart(xml);
    const s = model.plots[0]!.series[0]!;
    expect(s.categoryNumbers).toEqual([1, 2, 4, 8]);
    expect(s.values).toEqual([3, 5, 4, 9]);
    expect(s.line.paint).toBeNull();
    expect(describeChart(model)).toBe('Scatter chart: series Height ((1, 3.0), (2, 5.0), (4, 4.0), (8, 9.0))');

    const items = renderChart(xml, FRAME, { fonts, schemeColor });
    const p = paths(items);
    // No connecting line, four diamonds.
    expect(p.filter((x) => x.stroke === '4472C4' && !x.fill && x.commands.length === 4)).toHaveLength(0);
    const diamonds = p.filter((x) => x.fill === '4472C4' && x.commands.length === 5);
    expect(diamonds).toHaveLength(4);
    const cxs = diamonds.map((x) => (x.commands[0] as { x: number }).x);
    // x positions are proportional to the x values (axis starts at 0).
    const unit = (cxs[1]! - cxs[0]!) / 1;
    expect(cxs[2]! - cxs[0]!).toBeCloseTo(3 * unit, 6);
    expect(cxs[3]! - cxs[0]!).toBeCloseTo(7 * unit, 6);
    // Y labels use the source format (0.0), X labels are General.
    const strings = textOf(items);
    expect(strings).toContain('0.0');
    expect(strings).toContain('8');
  });

  it('draws a bar + line combination on primary and secondary axes', () => {
    const xml = chartSpace(
      `<c:barChart><c:barDir val="col"/><c:grouping val="clustered"/><c:varyColors val="0"/>` +
        ser(0, 'Units', QUARTERS, [100, 200, 150, 250]) +
        `<c:gapWidth val="100"/><c:axId val="1"/><c:axId val="2"/></c:barChart>` +
        `<c:lineChart><c:grouping val="standard"/><c:varyColors val="0"/>` +
        ser(1, 'Margin', QUARTERS, [0.1, 0.25, 0.2, 0.3], '<c:marker><c:symbol val="none"/></c:marker>', '0%') +
        `<c:marker val="1"/><c:axId val="3"/><c:axId val="4"/></c:lineChart>` +
        `${catAx(1, 2)}${valAx(2, 1)}` +
        `<c:catAx><c:axId val="3"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="1"/><c:axPos val="b"/><c:majorTickMark val="out"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/><c:crossAx val="4"/><c:crosses val="autoZero"/></c:catAx>` +
        valAx(4, 3, 'r', '', '<c:numFmt formatCode="0%" sourceLinked="1"/>').replace('<c:crosses val="autoZero"/>', '<c:crosses val="max"/>').replace('<c:majorGridlines/>', ''),
      { legend: 'b', title: title('Units and margin') },
    );
    const model = readChart(xml);
    expect(model.plots.map((p) => p.type)).toEqual(['bar', 'line']);
    expect(describeChart(model)).toMatch(/^Combination chart \(clustered column, line\) "Units and margin": series Units \(Q1 100, .*\), Margin \(Q1 10%, Q2 25%, Q3 20%, Q4 30%\)$/);

    const items = renderChart(xml, FRAME, { fonts, schemeColor });
    const bars = rects(items, '4472C4').filter((b) => b.w > 12);
    expect(bars).toHaveLength(4);
    const line = paths(items).find((x) => x.stroke === 'ED7D31' && !x.fill && x.commands.length === 4)!;
    expect(line).toBeDefined();
    // Line points sit over the bar centres.
    line.commands.forEach((c, i) => expect((c as { x: number }).x).toBeCloseTo(bars[i]!.x + bars[i]!.w / 2, 6));
    // Secondary axis labels (percent) on the right of the plot, primary on the left.
    const t = texts(items);
    const right = Math.max(...bars.map((b) => b.x + b.w));
    const pct = t.filter((x) => /^\d+%$/.test(x.text));
    expect(pct.length).toBeGreaterThan(2);
    for (const l of pct) expect(l.x).toBeGreaterThan(right);
    const left = Math.min(...bars.map((b) => b.x));
    expect(t.find((x) => x.text === '0' && x.x < left)).toBeDefined();
    // The 30% point is the highest point of the line; the 10% one the lowest.
    const ys = line.commands.map((c) => (c as { y: number }).y);
    expect(Math.min(...ys)).toBe(ys[3]);
    expect(Math.max(...ys)).toBe(ys[0]);
  });

  it('draws rotated axis titles as outlines and keeps chart-area fills', () => {
    const axisTitle = (text: string) => `<c:title><c:tx><c:rich><a:bodyPr rot="-5400000" vert="horz"/><a:p><a:r><a:t>${text}</a:t></a:r></a:p></c:rich></c:tx><c:overlay val="0"/></c:title>`;
    const xml = parseXml(
      `<c:chartSpace ${NS}><c:chart><c:autoTitleDeleted val="1"/><c:plotArea><c:barChart><c:barDir val="col"/><c:grouping val="clustered"/>${ser(0, 'A', QUARTERS, [1, 2, 3, 4])}<c:axId val="1"/><c:axId val="2"/></c:barChart>` +
        catAx(1, 2, 'b', '').replace('<c:numFmt', `${axisTitle('Quarter')}<c:numFmt`) +
        valAx(2, 1, 'l').replace('<c:numFmt', `${axisTitle('Units')}<c:numFmt`) +
        `<c:spPr>${solid('F2F2F2')}</c:spPr></c:plotArea></c:chart><c:spPr>${solid('FFFFEE')}<a:ln><a:noFill/></a:ln></c:spPr></c:chartSpace>`,
    ).root;
    const items = renderChart(xml, FRAME, { fonts, schemeColor });
    const area = paths(items)[0]!;
    expect(area.fill).toBe('FFFFEE');
    expect(area.stroke).toBeUndefined();
    expect(bbox(area)).toEqual({ x1: FRAME.x, y1: FRAME.y, x2: FRAME.x + FRAME.w, y2: FRAME.y + FRAME.h });
    expect(paths(items).some((p) => p.fill === 'F2F2F2')).toBe(true);
    // The horizontal axis title is text; the vertical one is a tall, narrow outline left of the plot.
    expect(textOf(items)).toContain('Quarter');
    expect(textOf(items)).not.toContain('Units');
    const outline = paths(items).find((p) => p.fill === '000000' && p.commands.length > 20)!;
    const b = bbox(outline);
    expect(b.y2 - b.y1).toBeGreaterThan(b.x2 - b.x1);
    expect(b.x1).toBeLessThan(Math.min(...rects(items, '4472C4').map((r) => r.x)));
  });

  it('draws 3-D variants in 2-D and percent-stacked columns on a 0–100% axis', () => {
    const xml = chartSpace(
      `<c:bar3DChart><c:barDir val="col"/><c:grouping val="percentStacked"/><c:varyColors val="0"/>` +
        ser(0, 'A', ['X', 'Y'], [1, 3]) +
        ser(1, 'B', ['X', 'Y'], [3, 1]) +
        `<c:gapWidth val="150"/><c:overlap val="100"/><c:shape val="box"/><c:axId val="1"/><c:axId val="2"/><c:axId val="3"/></c:bar3DChart>${catAx(1, 2)}${valAx(2, 1, 'l', '', '<c:numFmt formatCode="0%" sourceLinked="1"/>')}`,
    );
    const model = readChart(xml);
    expect(model.plots[0]).toMatchObject({ type: 'bar', is3D: true, grouping: 'percentStacked' });
    expect(describeChart(model)).toMatch(/^3-D 100% stacked column chart: /);
    const items = renderChart(xml, FRAME, { fonts, schemeColor });
    const a = rects(items, '4472C4');
    const b = rects(items, 'ED7D31');
    expect(a).toHaveLength(2);
    expect(b).toHaveLength(2);
    // Every stack fills the whole 0–100% range: 25% + 75%, 75% + 25%.
    expect(a[0]!.h + b[0]!.h).toBeCloseTo(a[1]!.h + b[1]!.h, 6);
    expect(a[0]!.h / (a[0]!.h + b[0]!.h)).toBeCloseTo(0.25, 6);
    const labels = textOf(items).filter((t) => t.endsWith('%'));
    expect(labels[0]).toBe('0%');
    expect(labels[labels.length - 1]).toBe('100%');
  });

  it('places data labels with their number format', () => {
    const labels = '<c:dLbls><c:numFmt formatCode="#,##0.0" sourceLinked="0"/><c:showLegendKey val="0"/><c:showVal val="1"/><c:showCatName val="0"/><c:showSerName val="0"/><c:showPercent val="0"/><c:showBubbleSize val="0"/></c:dLbls>';
    const xml = chartSpace(`<c:barChart><c:barDir val="col"/><c:grouping val="clustered"/>${ser(0, 'A', ['X', 'Y'], [1200, 3400], labels)}<c:axId val="1"/><c:axId val="2"/></c:barChart>${catAx(1, 2)}${valAx(2, 1)}`);
    const items = renderChart(xml, FRAME, { fonts, schemeColor });
    const bars = rects(items, '4472C4');
    const t = texts(items);
    const l1 = t.find((x) => x.text === '1,200.0')!;
    const l2 = t.find((x) => x.text === '3,400.0')!;
    // Outside end: centred over the bar, just above its top.
    for (const [label, bar] of [
      [l1, bars[0]!],
      [l2, bars[1]!],
    ] as const) {
      const w = label.widths.reduce((s, v) => s + v, 0);
      expect(label.x + w / 2).toBeCloseTo(bar.x + bar.w / 2, 6);
      expect(label.y).toBeLessThan(bar.y);
      expect(label.y).toBeGreaterThan(bar.y - 15);
    }
  });

  it('draws bubbles sized by area', () => {
    const xml = chartSpace(
      `<c:bubbleChart><c:varyColors val="0"/><c:ser><c:idx val="0"/><c:order val="0"/><c:tx>${strRef(['B'])}</c:tx><c:xVal>${numRef([1, 2])}</c:xVal><c:yVal>${numRef([1, 2])}</c:yVal><c:bubbleSize>${numRef([1, 4])}</c:bubbleSize></c:ser><c:bubbleScale val="100"/><c:axId val="1"/><c:axId val="2"/></c:bubbleChart>${valAx(1, 2, 'b')}${valAx(2, 1, 'l')}`,
    );
    const items = renderChart(xml, FRAME, { fonts, schemeColor });
    const circles = paths(items).filter((p) => p.fill === '4472C4' && p.commands.filter((c) => c.op === 'C').length === 4).map(bbox);
    expect(circles).toHaveLength(2);
    // Four times the size is twice the diameter.
    expect((circles[1]!.x2 - circles[1]!.x1) / (circles[0]!.x2 - circles[0]!.x1)).toBeCloseTo(2, 6);
  });

  it('copes with empty and degenerate charts', () => {
    const empty = parseXml(`<c:chartSpace ${NS}><c:chart><c:plotArea><c:layout/></c:plotArea></c:chart></c:chartSpace>`).root;
    expect(readChart(empty).plots).toEqual([]);
    expect(describeChart(readChart(empty))).toBe('Chart');
    expect(paths(renderChart(empty, FRAME, { fonts, schemeColor }))).toHaveLength(1);
    expect(renderChart(empty, { x: 0, y: 0, w: 0, h: 0 }, { fonts, schemeColor })).toEqual([]);

    const cases = [
      `<c:barChart><c:barDir val="col"/><c:grouping val="clustered"/>${ser(0, 'A', ['X'], [null])}<c:axId val="1"/><c:axId val="2"/></c:barChart>${catAx(1, 2)}${valAx(2, 1)}`,
      `<c:barChart><c:barDir val="col"/><c:grouping val="stacked"/>${ser(0, 'A', ['X', 'Y'], [-5, -10])}<c:axId val="1"/><c:axId val="2"/></c:barChart>${catAx(1, 2)}${valAx(2, 1)}`,
      `<c:lineChart><c:grouping val="standard"/>${ser(0, 'A', ['X', 'Y'], [5, 5])}<c:axId val="9"/><c:axId val="8"/></c:lineChart>`,
      `<c:pieChart><c:varyColors val="1"/>${ser(0, 'A', ['X', 'Y'], [0, 0])}</c:pieChart>`,
      `<c:pieChart><c:varyColors val="1"/>${ser(0, 'A', ['X'], [7])}</c:pieChart>`,
      `<c:lineChart><c:grouping val="standard"/>${ser(0, 'A', ['X', 'Y', 'Z'], [1, 100, 10000])}<c:axId val="1"/><c:axId val="2"/></c:lineChart>${catAx(1, 2)}${valAx(2, 1).replace('<c:orientation val="minMax"/>', '<c:logBase val="10"/><c:orientation val="minMax"/>')}`,
      `<c:radarChart><c:radarStyle val="filled"/>${ser(0, 'A', ['X', 'Y', 'Z'], [1, 2, 3])}<c:axId val="1"/><c:axId val="2"/></c:radarChart>${catAx(1, 2)}${valAx(2, 1)}`,
      `<c:surfaceChart>${ser(0, 'A', ['X'], [1])}<c:axId val="1"/><c:axId val="2"/></c:surfaceChart>`,
    ];
    for (const plot of cases) {
      for (const legend of ['l', 'tr', 't']) {
        const items = renderChart(chartSpace(plot, { legend, autoTitleDeleted: false }), { x: 0, y: 0, w: 200, h: 120 }, { fonts, schemeColor });
        for (const item of items) {
          if (item.type === 'path') for (const c of item.commands) if (c.op !== 'Z') expect(Number.isFinite(c.x) && Number.isFinite(c.y)).toBe(true);
          if (item.type === 'text') expect(Number.isFinite(item.x) && Number.isFinite(item.y)).toBe(true);
        }
      }
    }
    // The log axis labels are powers of ten.
    const log = renderChart(chartSpace(cases[5]!), FRAME, { fonts, schemeColor });
    expect(textOf(log)).toEqual(expect.arrayContaining(['1', '10', '100', '1000', '10000']));
  });

  it('moves the value axis with reversed categories', () => {
    const bar = (crosses: string) =>
      chartSpace(
        `<c:barChart><c:barDir val="bar"/><c:grouping val="clustered"/>${ser(0, 'A', ['First', 'Second', 'Third'], [10, 20, 30])}<c:axId val="1"/><c:axId val="2"/></c:barChart>` +
          catAx(1, 2, 'l').replace('<c:orientation val="minMax"/>', '<c:orientation val="maxMin"/>') +
          valAx(2, 1, 'b').replace('<c:crosses val="autoZero"/>', `<c:crosses val="${crosses}"/>`),
      );
    for (const [crosses, above] of [
      ['autoZero', true],
      ['max', false],
    ] as const) {
      const items = renderChart(bar(crosses), FRAME, { fonts, schemeColor });
      const bars = rects(items, '4472C4');
      expect(bars).toHaveLength(3);
      // First category on top.
      expect(bars[0]!.y).toBeLessThan(bars[2]!.y);
      const top = Math.min(...bars.map((b) => b.y));
      const bottom = Math.max(...bars.map((b) => b.y + b.h));
      const valueLabels = texts(items).filter((t) => /^\d+$/.test(t.text));
      expect(valueLabels.length).toBeGreaterThan(2);
      for (const l of valueLabels) {
        if (above) expect(l.y).toBeLessThan(top);
        else expect(l.y).toBeGreaterThan(bottom);
      }
    }
  });
});

