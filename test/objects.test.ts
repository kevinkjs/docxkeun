/** Objects in the editor: listing them; inserting, reading, updating and copying charts; inserting shapes and text boxes. */
import { describe, expect, it } from 'vitest';
import { DocxEditor, EditError, LockedRegionError, OBJECT_CHAR, paragraphText, type DocObject, type DocPosition, type ShapeKind } from '../src/api/index.js';
import { ChartDataError, type NewChart } from '../src/charts/index.js';
import { workbookFor } from '../src/charts/workbook.js';
import { createDocx } from '../src/index.js';
import { layoutDocument, type PathItem, type TextItem } from '../src/layout/index.js';
import { DocxDocument, type ImageInline } from '../src/model/index.js';
import { CONTENT_TYPE, REL } from '../src/opc/index.js';
import { AccessPolicy } from '../src/security/index.js';
import { attr, childPath, elements, findAll, firstChild, NS, parseXml, serializeXml, textContent, type XmlElement } from '../src/xml/index.js';
import { readZip } from '../src/zip/zip.js';
import { buildDocx, p } from './helpers/docx.js';
import { inline, textBoxRun, textBoxRunHolding, textBoxShape, vmlTextBoxRun } from './helpers/drawings.js';
import { testRegistry } from './helpers/fonts.js';

const QUARTERS = ['Q1', 'Q2', 'Q3', 'Q4'];
const TWO = [
  { name: 'Revenue', values: [120, 150, 90, 180] },
  { name: 'Costs', values: [80, 95, 70, 110] },
];
const COLUMNS: NewChart = { type: 'column', title: 'Sales', categories: QUARTERS, series: TWO };
const CHART1 = 'word/charts/chart1.xml';
const WORKBOOK1 = 'word/embeddings/Microsoft_Excel_Worksheet1.xlsx';
const OLE_OBJECT = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/oleObject';
/** A 2 × 1 PNG at 96 dpi. */
const PNG = Uint8Array.from(atob('iVBORw0KGgoAAAANSUhEUgAAAAIAAAABCAYAAAD0In+KAAAAEUlEQVR4nGP4z8DwnwGIGQAd8AP9Fr8f4QAAAABJRU5ErkJggg=='), (c) => c.charCodeAt(0));

async function open(body = p('Sales by quarter'), extra: Partial<Parameters<typeof buildDocx>[0]> = {}) {
  const doc = await DocxDocument.open(await buildDocx({ body, ...extra }));
  return { doc, editor: new DocxEditor(doc) };
}

const reopen = async (doc: DocxDocument): Promise<DocxDocument> => DocxDocument.open(await doc.save());

/** The end of paragraph `i`'s last text. */
function endOf(editor: DocxEditor, i = 0): DocPosition {
  const t = findAll(editor.paragraphs()[i]!.el, 'w:t').at(-1)!;
  return { node: t, offset: textContent(t).length };
}

/** Cell references and values of the first worksheet of a workbook part. */
async function sheetCells(doc: DocxDocument, part: string): Promise<Record<string, string>> {
  const entry = readZip(await doc.pkg.getBytes(part)).find((e) => e.name === 'xl/worksheets/sheet1.xml')!;
  const root = parseXml(new TextDecoder().decode(await entry.read())).root;
  return Object.fromEntries(findAll(root, 'c').map((c) => [attr(c, 'r')!, textContent(c)]));
}

/**
 * The number format of each styled cell of a workbook part's first sheet: a built-in format's id
 * ("#17") or a custom format's code.
 */
async function cellFormats(doc: DocxDocument, part: string): Promise<Record<string, string>> {
  const entries = readZip(await doc.pkg.getBytes(part));
  const xml = async (name: string) => parseXml(new TextDecoder().decode(await entries.find((e) => e.name === name)!.read())).root;
  const styles = await xml('xl/styles.xml');
  const custom = new Map(findAll(styles, 'numFmt').map((f) => [attr(f, 'numFmtId')!, attr(f, 'formatCode')!]));
  const xfs = elements(findAll(styles, 'cellXfs')[0]!, 'xf');
  expect(attr(findAll(styles, 'cellXfs')[0], 'count')).toBe(String(xfs.length));
  const out: Record<string, string> = {};
  for (const c of findAll(await xml('xl/worksheets/sheet1.xml'), 'c')) {
    const xf = xfs[Number(attr(c, 's') ?? 0)]!;
    const id = attr(xf, 'numFmtId')!;
    if (id !== '0') out[attr(c, 'r')!] = custom.get(id) ?? `#${id}`;
  }
  return out;
}

/** The part a chart's c:externalData points at, and that relationship. */
async function workbookOf(doc: DocxDocument, chartPart: string) {
  const cs = doc.xmlSync(chartPart)!.root;
  const id = attr(firstChild(cs, 'c:externalData'), 'r:id');
  const rel = id === undefined ? undefined : await doc.pkg.relationshipById(chartPart, id);
  return { rel, part: rel && !rel.external ? doc.pkg.resolve(chartPart, rel) : undefined };
}

describe('objects in the editor', () => {
  it('lists every kind of object with its run, drawing, size and placement', async () => {
    const EMU = 12700;
    const ns = 'xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"';
    const picture =
      `<w:r><w:drawing><wp:inline><wp:extent cx="${72 * EMU}" cy="${36 * EMU}"/><wp:docPr id="1" name="Picture 1" descr="Logo"/>` +
      '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:blipFill><a:blip r:embed="rIdNone"/></pic:blipFill></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>';
    const textBox =
      '<w:r><mc:AlternateContent><mc:Choice Requires="wps"><w:drawing><wp:anchor behindDoc="0" relativeHeight="1"><wp:simplePos x="0" y="0"/>' +
      '<wp:positionH relativeFrom="column"><wp:posOffset>0</wp:posOffset></wp:positionH><wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV>' +
      `<wp:extent cx="${144 * EMU}" cy="${72 * EMU}"/><wp:wrapSquare wrapText="bothSides"/><wp:docPr id="2" name="Text Box 2"/>` +
      '<a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><wps:wsp><wps:cNvSpPr txBox="1"/><wps:spPr/>' +
      '<wps:txbx><w:txbxContent><w:p><w:r><w:t>Inside</w:t></w:r></w:p></w:txbxContent></wps:txbx><wps:bodyPr/></wps:wsp></a:graphicData></a:graphic></wp:anchor></w:drawing></mc:Choice>' +
      '<mc:Fallback><w:pict><v:rect style="width:144pt;height:72pt"><v:textbox><w:txbxContent><w:p><w:r><w:t>Old</w:t></w:r></w:p></w:txbxContent></v:textbox></v:rect></w:pict></mc:Fallback></mc:AlternateContent></w:r>';
    const shape =
      `<w:r><w:drawing><wp:inline><wp:extent cx="${50 * EMU}" cy="${50 * EMU}"/><wp:docPr id="3" name="Oval 3"/>` +
      '<a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><wps:wsp><wps:cNvSpPr/><wps:spPr><a:prstGeom prst="ellipse"/></wps:spPr><wps:bodyPr/></wps:wsp></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>';
    const group =
      `<w:r><w:drawing><wp:inline><wp:extent cx="${100 * EMU}" cy="${100 * EMU}"/><wp:docPr id="4" name="Group 4"/>` +
      '<a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup"><wpg:wgp><wpg:grpSpPr/>' +
      '<wps:wsp><wps:cNvSpPr/><wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="10" cy="10"/></a:xfrm></wps:spPr><wps:bodyPr/></wps:wsp>' +
      '<wps:wsp><wps:cNvSpPr/><wps:spPr><a:xfrm><a:off x="10" y="0"/><a:ext cx="10" cy="10"/></a:xfrm></wps:spPr><wps:bodyPr/></wps:wsp></wpg:wgp></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>';
    const chart =
      `<w:r><w:drawing><wp:inline><wp:extent cx="${200 * EMU}" cy="${100 * EMU}"/><wp:docPr id="5" name="Chart 5"/>` +
      `<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart ${ns} r:id="rIdChart"/></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>`;
    const vml = '<w:r><w:pict><v:rect style="width:150pt;height:60pt"><v:textbox><w:txbxContent><w:p><w:r><w:t>VML</w:t></w:r></w:p></w:txbxContent></v:textbox></v:rect></w:pict></w:r>';
    const body =
      `<w:p>${picture}<w:r><w:t>text</w:t></w:r>${textBox}</w:p>` +
      p('No objects') +
      `<w:tbl><w:tr><w:tc><w:p>${shape}${group}</w:p></w:tc></w:tr></w:tbl>` +
      `<w:p>${chart}${vml}</w:p>`;
    const chartXml = `<?xml version="1.0" encoding="UTF-8"?><c:chartSpace ${ns}><c:chart><c:plotArea/></c:chart></c:chartSpace>`;
    const { editor } = await open(body, {
      rels: [['rIdChart', 'chart', 'charts/chart1.xml']],
      parts: { [CHART1]: { data: chartXml, contentType: CONTENT_TYPE.chart } },
    });

    const objects = editor.objects();
    expect(objects.map((o) => [o.n, o.kind, o.vml, o.floating, o.width, o.height, o.drawing.name])).toEqual([
      [1, 'picture', false, false, 72, 36, 'w:drawing'],
      [2, 'textBox', false, true, 144, 72, 'w:drawing'],
      [3, 'shape', false, false, 50, 50, 'w:drawing'],
      [4, 'group', false, false, 100, 100, 'w:drawing'],
      [5, 'chart', false, false, 200, 100, 'w:drawing'],
      [6, 'textBox', true, false, 150, 60, 'w:pict'],
    ]);
    const paras = editor.paragraphs();
    // The text box's paragraph comes right after the first paragraph.
    expect(textContent(paras[1]!.el)).toBe('Inside');
    expect(objects.map((o) => paras.indexOf(o.paragraph))).toEqual([0, 0, 3, 3, 4, 4]);
    expect(objects[0]!.alt).toBe('Logo');
    expect(objects[4]!.chartPart).toBe(CHART1);
    // The real run, not a stand-in, and the w:drawing inside mc:Choice.
    for (const o of objects) {
      expect(o.run.name).toBe('w:r');
      expect(o.paragraph.el.children).toContain(o.run);
      expect(findAll(o.run, o.drawing.name)).toContain(o.drawing);
    }
    expect(objects[1]!.drawing).toBe(findAll(objects[1]!.run, 'mc:Choice').flatMap((c) => findAll(c, 'w:drawing'))[0]);
  });
});

describe('charts in the editor', () => {
  it('inserts a chart with its own part and workbook', async () => {
    const { doc, editor } = await open();
    const drawing = await editor.insertChart(endOf(editor), COLUMNS);
    expect(drawing.name).toBe('w:drawing');

    const again = await reopen(doc);
    const ed = new DocxEditor(again);
    expect(ed.objects()).toMatchObject([{ kind: 'chart', n: 1, chartPart: CHART1 }]);
    const image = ed.paragraphs()[0]!.inlines.find((i): i is ImageInline => i.kind === 'image');
    expect(image?.content?.type).toBe('chart');
    expect(again.pkg.has(CHART1)).toBe(true);
    expect(again.pkg.contentType(CHART1)).toBe(CONTENT_TYPE.chart);
    const packages = (await again.pkg.relationships(CHART1)).filter((r) => r.type === REL.package);
    expect(packages).toEqual([{ id: expect.any(String), type: REL.package, target: '../embeddings/Microsoft_Excel_Worksheet1.xlsx', external: false }]);
    expect(again.pkg.contentType(WORKBOOK1)).toBe(CONTENT_TYPE.xlsx);
    expect(readZip(await again.pkg.getBytes(WORKBOOK1)).map((e) => e.name)).toContain('xl/worksheets/sheet1.xml');
    expect(await sheetCells(again, WORKBOOK1)).toMatchObject({ B1: 'Revenue', C1: 'Costs', A2: 'Q1', B5: '180' });
    const ids = findAll(again.documentXml.root, 'wp:docPr').map((d) => attr(d, 'id'));
    expect(new Set(ids).size).toBe(ids.length);
    const extent = findAll(again.documentXml.root, 'wp:extent')[0];
    expect([attr(extent, 'cx'), attr(extent, 'cy')]).toEqual(['5486400', '3200400']);
    const graphicData = findAll(again.documentXml.root, 'a:graphicData')[0]!;
    expect(attr(graphicData, 'uri')).toBe('http://schemas.openxmlformats.org/drawingml/2006/chart');
  });

  it('shrinks a chart to the text width', async () => {
    // Section 1 (paragraph 1) has 3 in of text width; section 2 the default 6.5 in.
    const narrow = '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="3960" w:bottom="1440" w:left="3960" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr>';
    const { doc, editor } = await open(p('Narrow', narrow) + p('Wide'));
    await editor.insertChart(endOf(editor, 0), COLUMNS);
    await editor.insertChart(endOf(editor, 1), COLUMNS);
    const extents = findAll(doc.documentXml.root, 'wp:extent').map((e) => [attr(e, 'cx'), attr(e, 'cy')]);
    expect(extents).toEqual([
      ['2743200', '1600200'],
      ['5486400', '3200400'],
    ]);
    const ids = findAll(doc.documentXml.root, 'wp:docPr').map((d) => attr(d, 'id'));
    expect(new Set(ids).size).toBe(2);
    expect(doc.pkg.has('word/charts/chart2.xml') && doc.pkg.has('word/embeddings/Microsoft_Excel_Worksheet2.xlsx')).toBe(true);
  });

  it('reads a chart back', async () => {
    const { doc, editor } = await open();
    await editor.insertChart(endOf(editor), { ...COLUMNS, axisTitles: { value: 'USD' } });
    const ed = new DocxEditor(await reopen(doc));
    expect(ed.chartInfo(ed.objects()[0]!)).toEqual({
      kind: 'clustered column',
      title: 'Sales',
      axisTitles: { value: 'USD' },
      legend: 'bottom',
      categories: QUARTERS,
      series: TWO,
      formatCode: 'General',
      source: 'embedded',
    });
  });

  it('updates the data and rewrites the workbook', async () => {
    const { doc, editor } = await open();
    await editor.insertChart(endOf(editor), COLUMNS);
    const three = [...TWO, { name: 'Profit', values: [40, 55, 20, 70] }];
    const result = await editor.updateChart(editor.objects()[0]!, { series: three, title: 'Sales and profit' });
    expect(result).toEqual({ source: 'embedded', workbookRewritten: true, workbookHadMore: false });
    expect(editor.chartInfo(editor.objects()[0]!)).toMatchObject({ title: 'Sales and profit', series: three });

    const again = await reopen(doc);
    const ed = new DocxEditor(again);
    expect(ed.chartInfo(ed.objects()[0]!).series).toEqual(three);
    expect((await workbookOf(again, CHART1)).part).toBe(WORKBOOK1);
    expect(await sheetCells(again, WORKBOOK1)).toMatchObject({ D1: 'Profit', D2: '40', D5: '70' });
  });

  it('keeps number formats through a read, an update and the workbook', async () => {
    // A chart as Word saves one of monthly shares: dates (Excel serial numbers shown "mmm-yy") and percentages.
    const JAN = 45292;
    const months = [JAN, JAN + 31, JAN + 60];
    const { doc, editor } = await open();
    await editor.insertChart(endOf(editor), {
      type: 'line',
      categories: months,
      series: [
        { name: 'Share', values: [0.1, 0.25, 0.3] },
        { name: 'Growth', values: [0.012, 0.034, 0.051] },
      ],
    });
    const cs = doc.xmlSync(CHART1)!.root;
    const formatCodes = (path: string) => findAll(cs, 'c:ser').map((ser) => firstChild(childPath(ser, path, 'c:numRef', 'c:numCache'), 'c:formatCode')!);
    for (const fc of formatCodes('c:cat')) fc.children = [{ type: 'text', value: 'mmm-yy' }];
    const [share, growth] = formatCodes('c:val');
    share!.children = [{ type: 'text', value: '0%' }];
    growth!.children = [{ type: 'text', value: '0.0%' }];
    doc.touch(CHART1);
    const obj = editor.objects()[0]!;

    const info = editor.chartInfo(obj);
    expect(info).toMatchObject({ categories: months, categoryFormat: 'mmm-yy', formatCode: '0%' });
    expect(info.series).toEqual([
      { name: 'Share', values: [0.1, 0.25, 0.3], formatCode: '0%' },
      { name: 'Growth', values: [0.012, 0.034, 0.051], formatCode: '0.0%' },
    ]);

    // An update by position keeps each series' format and the categories' (still numbers, not text).
    await editor.updateChart(obj, { series: [{ name: 'Share', values: [0.2, 0.3, 0.4] }, { name: 'Growth', values: [0.02, 0.03, 0.04] }, { name: 'Added', values: [0.5, 0.6, 0.7] }] });
    const after = editor.chartInfo(obj);
    expect(after).toMatchObject({ categories: months, categoryFormat: 'mmm-yy' });
    expect(after.series.map((x) => x.formatCode)).toEqual(['0%', '0.0%', '0.0%']);
    for (const ser of findAll(cs, 'c:ser')) expect(childPath(ser, 'c:cat', 'c:numRef', 'c:numCache')).toBeDefined();
    expect(await sheetCells(doc, WORKBOOK1)).toMatchObject({ A2: String(JAN), B2: '0.2', C2: '0.02', D1: 'Added' });
    // In the workbook: "mmm-yy" and "0%" are Excel's built-in formats 17 and 9; "0.0%" is a custom one.
    expect(await cellFormats(doc, WORKBOOK1)).toEqual({
      A2: '#17', A3: '#17', A4: '#17',
      B2: '#9', B3: '#9', B4: '#9',
      C2: '0.0%', C3: '0.0%', C4: '0.0%',
      D2: '0.0%', D3: '0.0%', D4: '0.0%',
    });

    // A missing date stays a gap in numeric categories, also when given back as an empty string.
    const catCaches = findAll(cs, 'c:ser').map((ser) => childPath(ser, 'c:cat', 'c:numRef', 'c:numCache')!);
    for (const cache of catCaches) cache.children = cache.children.filter((c) => !(c.type === 'element' && c.name === 'c:pt' && attr(c, 'idx') === '1'));
    doc.touch(CHART1);
    expect(editor.chartInfo(obj).categories).toEqual([JAN, null, JAN + 60]);
    await editor.updateChart(obj, { categories: [JAN, '', JAN + 60] });
    expect(editor.chartInfo(obj)).toMatchObject({ categories: [JAN, null, JAN + 60], categoryFormat: 'mmm-yy' });
    for (const ser of findAll(cs, 'c:ser')) expect(childPath(ser, 'c:cat', 'c:strRef')).toBeUndefined();
    const cells = await sheetCells(doc, WORKBOOK1);
    expect([cells.A2, cells.A3, cells.A4]).toEqual([String(JAN), undefined, String(JAN + 60)]);

    // New charts take formats too.
    await editor.insertChart(endOf(editor), { type: 'column', categories: months, categoryFormat: 'mmm-yy', series: [{ name: 'Share', values: [0.1, 0.2, 0.3], formatCode: '0%' }] });
    const made = editor.objects()[1]!;
    expect(editor.chartInfo(made)).toMatchObject({ categoryFormat: 'mmm-yy', formatCode: '0%' });
    expect(await cellFormats(doc, 'word/embeddings/Microsoft_Excel_Worksheet2.xlsx')).toEqual({ A2: '#17', A3: '#17', A4: '#17', B2: '#9', B3: '#9', B4: '#9' });
  });

  it('says when the old workbook held more than the chart', async () => {
    const { editor } = await open();
    await editor.insertChart(endOf(editor), COLUMNS);
    const obj = editor.objects()[0]!;
    // Fewer categories: the old workbook has rows the chart no longer uses, which is fine…
    expect(await editor.updateChart(obj, { categories: ['H1', 'H2'], series: [{ name: 'Revenue', values: [270, 270] }] })).toMatchObject({ workbookHadMore: false });
    // …but data beyond the table the chart used is reported.
    const { doc } = editor;
    doc.pkg.setBytes(WORKBOOK1, await workbookFor({ categories: ['a', 'b', 'c'], series: [{ name: 'x', values: [1, 2, 3] }, { name: 'y', values: [1, 2, 3] }] }));
    expect(await editor.updateChart(obj, { title: 'Halves' })).toEqual({ source: 'embedded', workbookRewritten: false, workbookHadMore: false });
    expect(await editor.updateChart(obj, { series: [{ name: 'Revenue', values: [1, 2] }] })).toEqual({ source: 'embedded', workbookRewritten: true, workbookHadMore: true });
  });

  it('says a workbook it cannot read may have held more', async () => {
    const { doc, editor } = await open();
    await editor.insertChart(endOf(editor), COLUMNS);
    const obj = editor.objects()[0]!;
    doc.pkg.setBytes(WORKBOOK1, new Uint8Array([0x50, 0x4b, 1, 2, 3]));
    // Not overwritten yet: nothing to report.
    expect(await editor.updateChart(obj, { title: 'Sales' })).toEqual({ source: 'embedded', workbookRewritten: false, workbookHadMore: false });
    expect(await editor.updateChart(obj, { series: [{ name: 'Revenue', values: [1, 2, 3, 4] }] })).toEqual({ source: 'embedded', workbookRewritten: true, workbookHadMore: true });
    expect(await sheetCells(doc, WORKBOOK1)).toMatchObject({ B1: 'Revenue', B2: '1' });
  });

  it('checks new data against the chart type', async () => {
    const { editor } = await open(p('Pie: ') + p('Scatter: '));
    await editor.insertChart(endOf(editor, 0), { type: 'pie', categories: QUARTERS, series: [TWO[0]!] });
    await editor.insertChart(endOf(editor, 1), { type: 'scatter', categories: [1, 2, 3, 4], series: [TWO[0]!] });
    const [pie, scatter] = editor.objects();
    await expect(editor.updateChart(pie!, { series: TWO })).rejects.toThrow(/a pie chart has one series/);
    await expect(editor.updateChart(scatter!, { categories: QUARTERS })).rejects.toThrow(ChartDataError);
    await expect(editor.updateChart(pie!, { series: [{ name: 'Short', values: [1] }] })).rejects.toThrow(/has 1 values for 4 categories/);
    expect(editor.chartInfo(pie!).series).toEqual([TWO[0]]);
    await expect(editor.insertChart(endOf(editor), { type: 'line', stacked: 'stacked', categories: QUARTERS, series: TWO })).rejects.toThrow(ChartDataError);
    expect(editor.objects()).toHaveLength(2);
  });

  it('updates a linked chart in the chart only', async () => {
    const { doc, editor } = await open();
    await editor.insertChart(endOf(editor), COLUMNS);
    // Point the chart at an outside workbook, as a linked Word chart does.
    const link = await doc.pkg.addRelationship(CHART1, OLE_OBJECT, 'file:///C:/data.xlsx', true);
    firstChild(doc.xmlSync(CHART1)!.root, 'c:externalData')!.attrs['r:id'] = link;
    doc.touch(CHART1);
    const again = await reopen(doc);
    const ed = new DocxEditor(again);
    const obj = ed.objects()[0]!;
    expect(ed.chartInfo(obj).source).toBe('linked');
    const workbook = await again.pkg.getBytes(WORKBOOK1);

    const series = [{ name: 'Revenue', values: [1, 2, 3, 4] }, TWO[1]!];
    expect(await ed.updateChart(obj, { series })).toEqual({ source: 'linked', workbookRewritten: false, workbookHadMore: false });
    expect(ed.chartInfo(obj).series).toEqual(series);
    // The formulas still point at the outside workbook's cells; the embedded copy is untouched.
    const ser = findAll(again.xmlSync(CHART1)!.root, 'c:ser')[0]!;
    expect(textContent(childPath(ser, 'c:val', 'c:numRef', 'c:f')!)).toBe('Sheet1!$B$2:$B$5');
    expect(await again.pkg.getBytes(WORKBOOK1)).toBe(workbook);
    expect((await workbookOf(again, CHART1)).rel).toMatchObject({ external: true, target: 'file:///C:/data.xlsx' });
  });

  it('gives a chart without a workbook, or with an old embedded object, a new workbook', async () => {
    const { doc, editor } = await open(p('One: ') + p('Two: '));
    await editor.insertChart(endOf(editor, 0), COLUMNS);
    await editor.insertChart(endOf(editor, 1), COLUMNS);
    const [none, ole] = editor.objects();
    // Chart 1: no c:externalData at all.
    const cs1 = doc.xmlSync(CHART1)!.root;
    cs1.children = cs1.children.filter((c) => !(c.type === 'element' && c.name === 'c:externalData'));
    // Chart 2: an old .xls OLE object instead of an .xlsx package.
    const CHART2 = 'word/charts/chart2.xml';
    const OLD = 'word/embeddings/Microsoft_Excel_97-2003_Worksheet1.xls';
    doc.pkg.setBytes(OLD, new Uint8Array([0xd0, 0xcf, 0x11, 0xe0]), 'application/vnd.ms-excel');
    const oldId = await doc.pkg.addRelationship(CHART2, OLE_OBJECT, OLD);
    firstChild(doc.xmlSync(CHART2)!.root, 'c:externalData')!.attrs['r:id'] = oldId;
    await doc.refreshRels(CHART2);
    expect(editor.chartInfo(none!).source).toBe('none');
    expect(editor.chartInfo(ole!).source).toBe('embedded');

    const series = [{ name: 'Revenue', values: [5, 6, 7, 8] }];
    expect(await editor.updateChart(none!, { series })).toEqual({ source: 'embedded', workbookRewritten: true, workbookHadMore: false });
    // What the old object held can't be read, so the caller is told it may have held more.
    expect(await editor.updateChart(ole!, { series })).toEqual({ source: 'embedded', workbookRewritten: true, workbookHadMore: true });

    const again = await reopen(doc);
    const ed = new DocxEditor(again);
    for (const [obj, part] of ed.objects().map((o) => [o, o.chartPart!] as const)) {
      expect(ed.chartInfo(obj)).toMatchObject({ source: 'embedded', series });
      const { rel, part: workbook } = await workbookOf(again, part);
      expect(rel?.type).toBe(REL.package);
      expect(again.pkg.contentType(workbook!)).toBe(CONTENT_TYPE.xlsx);
      expect(await sheetCells(again, workbook!)).toMatchObject({ B1: 'Revenue', B2: '5' });
      // The externalData element sits where the schema wants it.
      const names = (again.xmlSync(part)!.root.children.filter((c) => c.type === 'element') as XmlElement[]).map((c) => c.name);
      expect(names.slice(-2)).toEqual(['c:chart', 'c:externalData']);
    }
    // The old object and its relationship are gone.
    expect(again.pkg.has(OLD)).toBe(false);
    expect((await again.pkg.relationships(CHART2)).some((r) => r.type === OLE_OBJECT)).toBe(false);
  });

  it('records a new chart as an insertion with Track Changes on', async () => {
    const doc = await DocxDocument.open(await buildDocx({ body: p('Sales by quarter') }));
    const editor = new DocxEditor(doc, undefined, { trackChanges: true, author: 'AI' });
    const drawing = await editor.insertChart(endOf(editor), COLUMNS);
    const ins = findAll(doc.documentXml.root, 'w:ins').find((i) => findAll(i, 'w:drawing').includes(drawing));
    expect(ins).toBeDefined();
    expect(attr(ins, 'w:author')).toBe('AI');
    expect(editor.revisions()).toMatchObject([{ kind: 'insert', author: 'AI' }]);
  });

  it('refuses to insert in a locked region', async () => {
    const { doc, editor } = await open(p('The rate is 3.75 percent per year'));
    editor.lockRange(editor.find('3.75 percent')[0]!, { title: 'Agreed rate' });
    const inside = editor.find('.75')[0]!.start;
    await expect(editor.insertChart(inside, COLUMNS)).rejects.toThrow(LockedRegionError);
    expect(doc.pkg.has(CHART1)).toBe(false);
    expect(editor.objects()).toEqual([]);
  });

  it('refuses to update a chart in a locked region', async () => {
    const { editor } = await open(p('Chart: ') + p('After'));
    await editor.insertChart(endOf(editor), COLUMNS);
    editor.lockParagraphs(editor.paragraphs()[0]!.el, undefined, { title: 'Figures' });
    await expect(editor.updateChart(editor.objects()[0]!, { title: 'Changed' })).rejects.toThrow(LockedRegionError);
    expect(editor.chartInfo(editor.objects()[0]!).title).toBe('Sales');
  });

  it('gives copied charts their own chart part', async () => {
    const { doc, editor } = await open(p('Chart: ') + p('After'));
    await editor.insertChart(endOf(editor), COLUMNS);
    const para = editor.paragraphs()[0]!.el;
    await editor.copyBlocks(para, para, { after: para });

    const charts = findAll(doc.documentXml.root, 'c:chart');
    expect(charts).toHaveLength(2);
    const ids = charts.map((c) => attr(c, 'r:id')!);
    expect(ids[0]).not.toBe(ids[1]);
    const parts = ids.map((id) => doc.rels(doc.mainPart).get(id)!.target);
    expect(parts[0]).toBe(CHART1);
    expect(parts[1]).not.toBe(CHART1);
    expect(doc.pkg.contentType(parts[1]!)).toBe(CONTENT_TYPE.chart);
    const workbooks = await Promise.all(parts.map(async (part) => (await workbookOf(doc, part)).part));
    expect(workbooks[0]).toBe(WORKBOOK1);
    expect(workbooks[1]).not.toBe(WORKBOOK1);
    expect(doc.pkg.contentType(workbooks[1]!)).toBe(CONTENT_TYPE.xlsx);

    const [original, copy] = editor.objects();
    expect(copy!.chartPart).toBe(parts[1]);
    await editor.updateChart(copy!, { series: [{ name: 'Only', values: [1, 2, 3, 4] }] });
    expect(editor.chartInfo(original!).series).toEqual(TWO);
    expect(editor.chartInfo(copy!).series).toEqual([{ name: 'Only', values: [1, 2, 3, 4] }]);

    const again = await reopen(doc);
    const ed = new DocxEditor(again);
    expect(ed.objects().map((o) => ed.chartInfo(o).series.map((s) => s.name))).toEqual([['Revenue', 'Costs'], ['Only']]);
    expect(await sheetCells(again, workbooks[0]!)).toMatchObject({ B1: 'Revenue', C1: 'Costs' });
    expect(await sheetCells(again, workbooks[1]!)).toMatchObject({ B1: 'Only' });
  });
});

describe('shapes and text boxes', () => {
  const EMU = 12700;
  /** The paragraphs' text, as the editor lists them (text box paragraphs right after their holder). */
  const texts = (editor: DocxEditor) => editor.paragraphs().map((para) => paragraphText(para).text);
  const shapeOf = (drawing: XmlElement) => childPath(elements(drawing)[0], 'a:graphic', 'a:graphicData', 'wps:wsp')!;
  const position = (anchor: XmlElement, axis: 'H' | 'V') => {
    const pos = firstChild(anchor, `wp:position${axis}`);
    return [attr(pos, 'relativeFrom'), textContent(firstChild(pos, 'wp:posOffset')!)];
  };

  it('adds a floating text box with text, white with a black outline', async () => {
    const { doc, editor } = await open(p('Before') + p('After'));
    const drawing = await editor.insertShape(endOf(editor), { shape: 'textBox', text: 'Line one\nLine two' });
    expect(drawing.name).toBe('w:drawing');

    const again = await reopen(doc);
    const run = findAll(again.documentXml.root, 'w:r').find((r) => firstChild(r, 'mc:AlternateContent'))!;
    const alternate = firstChild(run, 'mc:AlternateContent')!;
    expect(elements(alternate).map((e) => [e.name, attr(e, 'Requires')])).toEqual([['mc:Choice', 'wps']]);
    const anchor = childPath(alternate, 'mc:Choice', 'w:drawing', 'wp:anchor')!;
    expect(elements(anchor).map((e) => e.name)).toEqual([
      'wp:simplePos', 'wp:positionH', 'wp:positionV', 'wp:extent', 'wp:effectExtent', 'wp:wrapSquare', 'wp:docPr', 'wp:cNvGraphicFramePr', 'a:graphic',
    ]);
    expect(anchor.attrs).toMatchObject({ behindDoc: '0', distL: '114300', distR: '114300', simplePos: '0', locked: '0', layoutInCell: '1', allowOverlap: '1', relativeHeight: '251659264' });
    expect(attr(firstChild(anchor, 'wp:wrapSquare'), 'wrapText')).toBe('bothSides');
    expect(position(anchor, 'H')).toEqual(['column', '0']);
    expect(position(anchor, 'V')).toEqual(['paragraph', '0']);
    expect(firstChild(anchor, 'wp:extent')!.attrs).toEqual({ cx: String(180 * EMU), cy: String(72 * EMU) });
    expect(attr(firstChild(anchor, 'wp:docPr'), 'name')).toMatch(/^Text Box \d+$/);

    const wsp = childPath(anchor, 'a:graphic', 'a:graphicData', 'wps:wsp')!;
    expect(elements(wsp).map((e) => e.name)).toEqual(['wps:cNvSpPr', 'wps:spPr', 'wps:txbx', 'wps:bodyPr']);
    expect(attr(firstChild(wsp, 'wps:cNvSpPr'), 'txBox')).toBe('1');
    const spPr = firstChild(wsp, 'wps:spPr')!;
    expect(elements(spPr).map((e) => e.name)).toEqual(['a:xfrm', 'a:prstGeom', 'a:solidFill', 'a:ln']);
    expect(attr(childPath(spPr, 'a:prstGeom'), 'prst')).toBe('rect');
    expect(attr(childPath(spPr, 'a:solidFill', 'a:srgbClr'), 'val')).toBe('FFFFFF');
    const ln = firstChild(spPr, 'a:ln')!;
    expect(attr(ln, 'w')).toBe('9525');
    expect(attr(childPath(ln, 'a:solidFill', 'a:srgbClr'), 'val')).toBe('000000');
    expect(firstChild(wsp, 'wps:bodyPr')!.attrs).toMatchObject({ lIns: '91440', tIns: '45720', rIns: '91440', bIns: '45720', anchor: 't' });
    const paras = elements(childPath(wsp, 'wps:txbx', 'w:txbxContent')!, 'w:p');
    expect(paras.map((x) => textContent(x))).toEqual(['Line one', 'Line two']);

    const ed = new DocxEditor(again);
    expect(texts(ed)).toEqual(['Before', 'Line one', 'Line two', 'After']);
    expect(ed.objects()).toMatchObject([{ n: 1, kind: 'textBox', floating: true, width: 180, height: 72 }]);
    expect(ed.objects()[0]!.paragraph.el).toBe(ed.paragraphs()[0]!.el);
  });

  it('adds a shape in the theme accent with white centred text', async () => {
    const { editor } = await open();
    const drawing = await editor.insertShape(endOf(editor), { shape: 'wedgeRectCallout', text: 'Note' });
    const wsp = shapeOf(drawing);
    expect(elements(wsp).map((e) => e.name)).toEqual(['wps:cNvSpPr', 'wps:spPr', 'wps:style', 'wps:txbx', 'wps:bodyPr']);
    expect(attr(firstChild(wsp, 'wps:cNvSpPr'), 'txBox')).toBeUndefined();
    const spPr = firstChild(wsp, 'wps:spPr')!;
    expect(attr(firstChild(spPr, 'a:prstGeom'), 'prst')).toBe('wedgeRectCallout');
    // The style gives the fill and outline.
    expect(elements(spPr).map((e) => e.name)).toEqual(['a:xfrm', 'a:prstGeom']);
    const style = firstChild(wsp, 'wps:style')!;
    expect(elements(style).map((e) => e.name)).toEqual(['a:lnRef', 'a:fillRef', 'a:effectRef', 'a:fontRef']);
    const fillRef = firstChild(style, 'a:fillRef')!;
    expect(attr(fillRef, 'idx')).toBe('1');
    expect(attr(firstChild(fillRef, 'a:schemeClr'), 'val')).toBe('accent1');
    const lnRef = firstChild(style, 'a:lnRef')!;
    expect(attr(lnRef, 'idx')).toBe('2');
    expect(attr(firstChild(lnRef, 'a:schemeClr'), 'val')).toBe('accent1');
    expect(attr(childPath(lnRef, 'a:schemeClr', 'a:shade'), 'val')).toBe('50000');
    expect(attr(firstChild(wsp, 'wps:bodyPr'), 'anchor')).toBe('ctr');
    const para = findAll(wsp, 'w:p')[0]!;
    expect(attr(childPath(para, 'w:pPr', 'w:jc'), 'w:val')).toBe('center');
    expect(attr(childPath(firstChild(para, 'w:r'), 'w:rPr', 'w:color'), 'w:val')).toBe('FFFFFF');
    expect(textContent(para)).toBe('Note');
    expect(firstChild(elements(drawing)[0], 'wp:extent')!.attrs).toEqual({ cx: String(108 * EMU), cy: String(72 * EMU) });
    expect(editor.objects()).toMatchObject([{ kind: 'shape', floating: true }]);
    expect(texts(editor)).toEqual(['Sales by quarter', 'Note']);
  });

  it('places a shape inline or in front of or behind text', async () => {
    const { doc, editor } = await open(p('One') + p('Two') + p('Three') + p('Four') + p('Five'));
    const inline = await editor.insertShape(endOf(editor, 0), { shape: 'rect', place: { inline: true } });
    const front = await editor.insertShape(endOf(editor, 1), { shape: 'ellipse', place: { x: 0, y: 0, wrap: 'inFront' } });
    const behind = await editor.insertShape(endOf(editor, 2), { shape: 'star5', place: { x: 0, y: 0, wrap: 'behind' } });
    const page = await editor.insertShape(endOf(editor, 3), { shape: 'rightArrow', place: { x: 36, y: 18, from: 'page', wrap: 'topAndBottom' } });
    const margin = await editor.insertShape(endOf(editor, 4), { shape: 'line', place: { x: -9, y: 0, from: 'margin' } });

    expect(elements(inline).map((e) => e.name)).toEqual(['wp:inline']);
    expect(findAll(doc.documentXml.root, 'mc:Choice').map((c) => firstChild(c, 'w:drawing'))).toEqual([inline, front, behind, page, margin]);

    const anchors = [front, behind, page, margin].map((d) => firstChild(d, 'wp:anchor')!);
    const wraps = anchors.map((a) => elements(a).find((e) => e.name.startsWith('wp:wrap'))!.name);
    expect(wraps).toEqual(['wp:wrapNone', 'wp:wrapNone', 'wp:wrapTopAndBottom', 'wp:wrapSquare']);
    expect(anchors.map((a) => attr(a, 'behindDoc'))).toEqual(['0', '1', '0', '0']);
    expect(position(anchors[2]!, 'H')).toEqual(['page', '457200']);
    expect(position(anchors[2]!, 'V')).toEqual(['page', '228600']);
    expect(position(anchors[3]!, 'H')).toEqual(['margin', String(-9 * EMU)]);
    expect(position(anchors[3]!, 'V')).toEqual(['margin', '0']);
    // Each new floating object goes on top of the ones before it.
    expect(anchors.map((a) => Number(attr(a, 'relativeHeight')))).toEqual([251659264, 251659265, 251659266, 251659267]);

    // A line: a connector, 144 × 0 pt, with no text.
    const line = shapeOf(margin);
    expect(elements(line).map((e) => e.name)).toEqual(['wps:cNvCnPr', 'wps:spPr', 'wps:style', 'wps:bodyPr']);
    expect(attr(childPath(line, 'wps:spPr', 'a:prstGeom'), 'prst')).toBe('line');
    expect(firstChild(anchors[3]!, 'wp:extent')!.attrs).toEqual({ cx: String(144 * EMU), cy: '0' });

    const ids = findAll(doc.documentXml.root, 'wp:docPr').map((d) => attr(d, 'id'));
    expect(new Set(ids).size).toBe(5);
    expect(editor.objects().map((o) => [o.kind, o.floating])).toEqual([
      ['shape', false],
      ['shape', true],
      ['shape', true],
      ['shape', true],
      ['shape', true],
    ]);
  });

  it('takes colours or none', async () => {
    const { editor } = await open(p('One') + p('Two') + p('Three'));
    const plain = await editor.insertShape(endOf(editor, 0), { shape: 'ellipse', fill: 'none', outline: 'C00000' });
    const box = await editor.insertShape(endOf(editor, 1), { shape: 'textBox', fill: '#ffeeaa', outline: 'none' });
    const spPr = (d: XmlElement) => firstChild(shapeOf(d), 'wps:spPr')!;
    expect(elements(spPr(plain)).map((e) => e.name)).toEqual(['a:xfrm', 'a:prstGeom', 'a:noFill', 'a:ln']);
    expect(attr(childPath(spPr(plain), 'a:ln', 'a:solidFill', 'a:srgbClr'), 'val')).toBe('C00000');
    expect(attr(childPath(spPr(box), 'a:solidFill', 'a:srgbClr'), 'val')).toBe('FFEEAA');
    expect(elements(firstChild(spPr(box), 'a:ln')!).map((e) => e.name)).toEqual(['a:noFill']);
    // An empty text box still has a paragraph to type in.
    expect(elements(childPath(shapeOf(box), 'wps:txbx', 'w:txbxContent')!).map((e) => e.name)).toEqual(['w:p']);

    const before = editor.objects().length;
    const at = endOf(editor, 3);
    await expect(editor.insertShape(at, { shape: 'rect', fill: 'red' })).rejects.toThrow(EditError);
    await expect(editor.insertShape(at, { shape: 'rect', outline: '12345' })).rejects.toThrow(EditError);
    await expect(editor.insertShape(at, { shape: 'rect', width: 0 })).rejects.toThrow(EditError);
    await expect(editor.insertShape(at, { shape: 'rect', height: Number.NaN })).rejects.toThrow(EditError);
    await expect(editor.insertShape(at, { shape: 'line', text: 'No' })).rejects.toThrow(EditError);
    await expect(editor.insertShape(at, { shape: 'hexagon' as ShapeKind })).rejects.toThrow(EditError);
    await expect(editor.insertShape(at, { shape: 'rect', place: { x: Number.POSITIVE_INFINITY, y: 0 } })).rejects.toThrow(EditError);
    await expect(editor.insertShape(at, { shape: 'rect', place: { x: 0, y: 0, wrap: 'tight' as 'square' } })).rejects.toThrow(EditError);
    expect(editor.objects()).toHaveLength(before);
  });

  it('works in a document made from scratch', async () => {
    const doc = await createDocx();
    const editor = new DocxEditor(doc);
    editor.insertParagraph('end', 'Arrow here');
    editor.insertParagraph('end', 'Box here');
    await editor.insertChart({ node: editor.paragraphs()[0]!.el, offset: 0 }, COLUMNS);
    await editor.insertShape(editor.find('Arrow here')[0]!.end, { shape: 'rightArrow', text: 'Go', place: { inline: true } });
    await editor.insertShape(editor.find('Box here')[0]!.end, { shape: 'textBox', text: 'Boxed', place: { x: 0, y: 24 } });

    const again = await reopen(doc);
    const root = again.documentXml.root;
    for (const prefix of ['wp', 'a', 'wps', 'mc', 'r']) expect(attr(root, `xmlns:${prefix}`)).toBe(NS[prefix as keyof typeof NS]);
    const ed = new DocxEditor(again);
    expect(ed.objects().map((o) => o.kind)).toEqual(['chart', 'shape', 'textBox']);
    expect(texts(ed)).toEqual(['', 'Arrow here', 'Go', 'Box here', 'Boxed']);

    const items = layoutDocument(again, { fonts: testRegistry() }).pages[0]!.items;
    const words = items.filter((i): i is TextItem => i.type === 'text').map((t) => t.text);
    expect(words).toEqual(expect.arrayContaining(['Sales', 'Q1', 'Go', 'Boxed']));
    const shapes = items.filter((i): i is PathItem => i.type === 'path');
    // The arrow in the theme's accent 1, the text box white with a black outline.
    expect(shapes.some((s) => s.fill === '4472C4' && s.stroke !== undefined)).toBe(true);
    expect(shapes.some((s) => s.fill === 'FFFFFF' && s.stroke === '000000')).toBe(true);
    expect(items.filter((i): i is TextItem => i.type === 'text').find((t) => t.text === 'Go')!.color).toBe('FFFFFF');
  });

  it('records a new shape as an insertion with Track Changes on', async () => {
    const doc = await DocxDocument.open(await buildDocx({ body: p('Notes') }));
    const editor = new DocxEditor(doc, undefined, { trackChanges: true, author: 'AI' });
    const drawing = await editor.insertShape(endOf(editor), { shape: 'textBox', text: 'Tracked' });
    const ins = findAll(doc.documentXml.root, 'w:ins').find((i) => findAll(i, 'w:drawing').includes(drawing));
    expect(ins).toBeDefined();
    expect(attr(ins, 'w:author')).toBe('AI');
    expect(editor.revisions()).toMatchObject([{ kind: 'insert', author: 'AI' }]);
    // Rejecting it takes the whole box away.
    editor.rejectRevisions();
    expect(editor.objects()).toEqual([]);
    expect(texts(editor)).toEqual(['Notes']);
  });

  it('refuses to insert in a locked region', async () => {
    const { editor } = await open(p('The rate is 3.75 percent per year'));
    editor.lockRange(editor.find('3.75 percent')[0]!, { title: 'Agreed rate' });
    await expect(editor.insertShape(editor.find('.75')[0]!.start, { shape: 'textBox', text: 'No' })).rejects.toThrow(LockedRegionError);
    expect(editor.objects()).toEqual([]);
  });

  it('refuses a shape, text box or chart inside a text box, but takes a picture there', async () => {
    const { doc, editor } = await open(`<w:p>${textBoxRun('Inside the box')}</w:p>`);
    const inBox = () => editor.find('Inside the box')[0]!.end;
    for (const shape of ['textBox', 'rect'] as const) {
      await expect(editor.insertShape(inBox(), { shape, place: { inline: true } })).rejects.toThrow(new EditError("Shapes and text boxes can't go inside a text box"));
    }
    // Word puts only pictures in a text box.
    await expect(editor.insertChart(inBox(), COLUMNS)).rejects.toThrow(new EditError('Only pictures can go inside a text box'));
    expect(doc.pkg.has(CHART1)).toBe(false);
    // (The fixture's box is a shape holding text: its cNvSpPr has no txBox flag.)
    expect(editor.objects().map((o) => o.kind)).toEqual(['shape']);
    // Nothing changed, so the box keeps its old copy.
    expect(findAll(doc.documentXml.root, 'mc:Fallback')).toHaveLength(1);

    await editor.insertImage(inBox(), PNG);
    expect(editor.objects().map((o) => o.kind)).toEqual(['shape', 'picture']);
  });
});

describe('changing objects', () => {
  const EMU = 12700;
  const PICTURE = 'http://schemas.openxmlformats.org/drawingml/2006/picture';
  /** A w × h pt inline picture's w:drawing, with drawing id `id`. */
  const pictureDrawing = (id = 1, w = 200, h = 100) =>
    `<w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="${w * EMU}" cy="${h * EMU}"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:docPr id="${id}" name="Picture ${id}"/><wp:cNvGraphicFramePr/>` +
    `<a:graphic><a:graphicData uri="${PICTURE}"><pic:pic><pic:nvPicPr><pic:cNvPr id="0" name="Picture ${id}"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="rIdNone"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>` +
    `<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${w * EMU}" cy="${h * EMU}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing>`;
  const picture = (id = 1, w = 200, h = 100) => `<w:r>${pictureDrawing(id, w, h)}</w:r>`;
  /** A 50 × 50 pt inline oval (wps:wsp). */
  const oval = () => inline(`<wps:wsp><wps:cNvSpPr/><wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${50 * EMU}" cy="${50 * EMU}"/></a:xfrm><a:prstGeom prst="ellipse"><a:avLst/></a:prstGeom></wps:spPr><wps:bodyPr/></wps:wsp>`, 50, 50);
  const frameOf = (o: DocObject) => elements(o.drawing)[0]!;
  const extentOf = (o: DocObject) => {
    const extent = firstChild(frameOf(o), 'wp:extent')!;
    return [attr(extent, 'cx'), attr(extent, 'cy')];
  };
  const position = (anchor: XmlElement, axis: 'H' | 'V') => {
    const pos = firstChild(anchor, `wp:position${axis}`);
    return [attr(pos, 'relativeFrom'), textContent(firstChild(pos, 'wp:posOffset')!)];
  };
  /** Which paragraphs (by index in `paragraphs()`) hold the objects. */
  const where = (editor: DocxEditor) => {
    const paras = editor.paragraphs().map((x) => x.el);
    return editor.objects().map((o) => paras.indexOf(o.paragraph.el));
  };
  const texts = (editor: DocxEditor) => editor.paragraphs().map((para) => paragraphText(para).text);
  /** The w:ins or w:del elements holding a node named `name`. */
  const marked = (root: XmlElement, kind: 'w:ins' | 'w:del', name: string) => findAll(root, kind).filter((m) => findAll(m, name).length > 0);
  const drawingIds = (root: XmlElement) => findAll(root, 'wp:docPr').map((d) => attr(d, 'id'));
  const ANCHOR_CHILDREN = ['wp:simplePos', 'wp:positionH', 'wp:positionV', 'wp:extent', 'wp:effectExtent', 'wp:wrapSquare', 'wp:docPr', 'wp:cNvGraphicFramePr', 'a:graphic'];

  it('resizes keeping proportions', async () => {
    const group =
      `<w:r><w:drawing><wp:inline><wp:extent cx="${100 * EMU}" cy="${100 * EMU}"/><wp:docPr id="7" name="Group 7"/><a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup"><wpg:wgp><wpg:cNvGrpSpPr/>` +
      `<wpg:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${100 * EMU}" cy="${100 * EMU}"/><a:chOff x="0" y="0"/><a:chExt cx="${100 * EMU}" cy="${100 * EMU}"/></a:xfrm></wpg:grpSpPr>` +
      `<wps:wsp><wps:cNvSpPr/><wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${50 * EMU}" cy="${100 * EMU}"/></a:xfrm></wps:spPr><wps:bodyPr/></wps:wsp></wpg:wgp></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>`;
    const { doc, editor } = await open(`<w:p>${picture()}</w:p>` + p('Chart: ') + `<w:p>${group}</w:p>`);
    await editor.insertChart(endOf(editor, 1), COLUMNS);
    const [pic, chart, grp] = editor.objects();

    editor.resizeObject(pic!, { width: 100 });
    expect(extentOf(pic!)).toEqual(['1270000', '635000']);
    const ext = childPath(frameOf(pic!), 'a:graphic', 'a:graphicData', 'pic:pic', 'pic:spPr', 'a:xfrm', 'a:ext')!;
    expect([attr(ext, 'cx'), attr(ext, 'cy')]).toEqual(['1270000', '635000']);
    // The proportions are the object's current ones; both sides are taken as given.
    editor.resizeObject(pic!, { height: 25 });
    expect(extentOf(pic!)).toEqual([String(50 * EMU), String(25 * EMU)]);
    editor.resizeObject(pic!, { width: 30, height: 40 });
    expect(extentOf(pic!)).toEqual([String(30 * EMU), String(40 * EMU)]);
    expect([attr(ext, 'cx'), attr(ext, 'cy')]).toEqual([String(30 * EMU), String(40 * EMU)]);

    // A chart: its frame changes, its part does not.
    const chartXml = serializeXml(doc.xmlSync(CHART1)!);
    const relId = attr(findAll(chart!.drawing, 'c:chart')[0], 'r:id');
    editor.resizeObject(chart!, { width: 216 });
    expect(extentOf(chart!)).toEqual([String(216 * EMU), String(126 * EMU)]);
    expect(attr(findAll(chart!.drawing, 'c:chart')[0], 'r:id')).toBe(relId);
    expect(serializeXml(doc.xmlSync(CHART1)!)).toBe(chartXml);

    // A group: its own size changes and its child coordinates stay, so the shapes in it scale.
    editor.resizeObject(grp!, { width: 50 });
    const xfrm = childPath(frameOf(grp!), 'a:graphic', 'a:graphicData', 'wpg:wgp', 'wpg:grpSpPr', 'a:xfrm')!;
    expect(firstChild(xfrm, 'a:ext')!.attrs).toEqual({ cx: String(50 * EMU), cy: String(50 * EMU) });
    expect(firstChild(xfrm, 'a:chOff')!.attrs).toEqual({ x: '0', y: '0' });
    expect(firstChild(xfrm, 'a:chExt')!.attrs).toEqual({ cx: String(100 * EMU), cy: String(100 * EMU) });
    expect(childPath(findAll(grp!.drawing, 'wps:wsp')[0], 'wps:spPr', 'a:xfrm', 'a:ext')!.attrs).toEqual({ cx: String(50 * EMU), cy: String(100 * EMU) });

    const ed = new DocxEditor(await reopen(doc));
    expect(ed.objects().map((o) => [o.kind, o.width, o.height])).toEqual([
      ['picture', 30, 40],
      ['chart', 216, 126],
      ['group', 50, 50],
    ]);
    expect(ed.objects()[1]!.chartPart).toBe(CHART1);

    // A size is a positive number of points.
    for (const size of [{}, { width: 0 }, { height: -5 }, { width: Number.NaN }, { width: 10, height: Number.POSITIVE_INFINITY }]) {
      expect(() => editor.resizeObject(pic!, size)).toThrow(EditError);
    }
    expect(extentOf(pic!)).toEqual([String(30 * EMU), String(40 * EMU)]);
  });

  it('moves an inline object to another paragraph', async () => {
    // The picture shares a run with text: only the picture moves.
    const { doc, editor } = await open(`<w:p><w:r><w:t xml:space="preserve">Logo: </w:t>${pictureDrawing(4)}<w:t xml:space="preserve"> (2026)</w:t></w:r></w:p>` + p('Middle') + p('Target'));
    const [pic] = editor.objects();
    await editor.moveObject(pic!, endOf(editor, 2));
    const paras = editor.paragraphs();
    expect(editor.objects()[0]!.paragraph.el).toBe(paras[2]!.el);
    expect(findAll(paras[0]!.el, 'w:drawing')).toEqual([]);
    expect(texts(editor)).toEqual(['Logo:  (2026)', 'Middle', 'Target']);
    // The same object: no copy, no tracked change.
    expect(editor.objects()[0]!.drawing).toBe(pic!.drawing);
    expect(drawingIds(doc.documentXml.root)).toEqual(['4']);
    expect(findAll(doc.documentXml.root, 'w:ins')).toEqual([]);
    // It sits after the paragraph's text.
    expect(elements(paras[2]!.el).at(-1)).toBe(editor.objects()[0]!.run);

    const ed = new DocxEditor(await reopen(doc));
    expect(where(ed)).toEqual([2]);
    // Moving it next to itself changes nothing.
    await ed.moveObject(ed.objects()[0]!, { node: ed.objects()[0]!.drawing, offset: 1 });
    expect(where(ed)).toEqual([2]);
  });

  it('tracks a move as a deletion and an insertion', async () => {
    const doc = await DocxDocument.open(await buildDocx({ body: p('Chart: ') + p('Middle') + p('Target') }));
    const setup = new DocxEditor(doc);
    await setup.insertChart(endOf(setup, 0), COLUMNS);
    const editor = new DocxEditor(doc, undefined, { trackChanges: true, author: 'AI' });
    const [chart] = editor.objects();
    await editor.moveObject(chart!, endOf(editor, 2));

    const root = doc.documentXml.root;
    const [del] = marked(root, 'w:del', 'c:chart');
    const [ins] = marked(root, 'w:ins', 'c:chart');
    const paras = editor.paragraphs();
    expect(findAll(paras[0]!.el, 'w:del')).toEqual([del]);
    expect(elements(del!)[0]).toBe(chart!.run);
    expect(findAll(paras[2]!.el, 'w:ins')).toEqual([ins]);
    expect([attr(del, 'w:author'), attr(ins, 'w:author')]).toEqual(['AI', 'AI']);
    expect(editor.revisions().map((r) => r.kind)).toEqual(['delete', 'insert']);
    // The copy has its own drawing id, chart part and workbook.
    const ids = drawingIds(root);
    expect(ids).toHaveLength(2);
    expect(new Set(ids).size).toBe(2);
    const copyChart = findAll(ins!, 'c:chart')[0]!;
    const copyPart = doc.rels(doc.mainPart).get(attr(copyChart, 'r:id')!)!.target;
    expect(copyPart).not.toBe(CHART1);
    expect(doc.pkg.contentType(copyPart)).toBe(CONTENT_TYPE.chart);
    expect((await workbookOf(doc, copyPart)).part).not.toBe(WORKBOOK1);
    expect(editor.chartInfo(editor.objects().find((o) => o.chartPart === copyPart)!).series).toEqual(TWO);
    // Only the copy is listed now; the deleted original can't be changed.
    expect(editor.objects().map((o) => o.chartPart)).toEqual([copyPart]);
    expect(() => editor.resizeObject(chart!, { width: 100 })).toThrow(/is deleted/);
    // Moving the author's own insertion again takes it out of its old place outright.
    const again = new DocxEditor(await reopen(doc), undefined, { trackChanges: true, author: 'AI' });
    await again.moveObject(again.objects()[0]!, endOf(again, 1));
    const againRoot = again.doc.documentXml.root;
    expect(findAll(againRoot, 'w:ins').map((i) => [where(again)[0], findAll(i, 'c:chart').length])).toEqual([[1, 1]]);
    expect(findAll(againRoot, 'w:ins').every((i) => elements(i).length > 0)).toBe(true);
    expect(again.revisions().map((r) => r.kind)).toEqual(['delete', 'insert']);

    // Rejecting puts the chart back where it was; accepting leaves it at the target.
    const rejected = new DocxEditor(await reopen(doc));
    rejected.rejectRevisions();
    expect(where(rejected)).toEqual([0]);
    expect(rejected.objects()[0]!.chartPart).toBe(CHART1);
    const accepted = new DocxEditor(await reopen(doc));
    accepted.acceptRevisions();
    expect(where(accepted)).toEqual([2]);
    expect(accepted.objects()[0]!.chartPart).toBe(copyPart);
  });

  it('repositions and re-wraps a floating object, untracked', async () => {
    const { doc, editor } = await open(p('Box: ') + p('After'));
    await editor.insertShape(endOf(editor, 0), { shape: 'textBox', text: 'Boxed' });
    editor.trackChanges = true;
    const [box] = editor.objects();
    const anchor = frameOf(box!);
    await editor.placeObject(box!, { x: 72, y: 36, from: 'margin', wrap: 'topAndBottom' });
    expect(frameOf(box!)).toBe(anchor);
    expect(position(anchor, 'H')).toEqual(['margin', '914400']);
    expect(position(anchor, 'V')).toEqual(['margin', '457200']);
    expect(elements(anchor).map((e) => e.name)).toEqual(ANCHOR_CHILDREN.map((n) => (n === 'wp:wrapSquare' ? 'wp:wrapTopAndBottom' : n)));
    const root = doc.documentXml.root;
    expect([...findAll(root, 'w:ins'), ...findAll(root, 'w:del')]).toEqual([]);
    expect(editor.revisions()).toEqual([]);

    // Leaving out the reference or the wrap keeps the object's own.
    await editor.placeObject(box!, { x: 0, y: 12 });
    expect(position(anchor, 'H')).toEqual(['margin', '0']);
    expect(position(anchor, 'V')).toEqual(['margin', String(12 * EMU)]);
    expect(firstChild(anchor, 'wp:wrapTopAndBottom')).toBeDefined();
    await editor.placeObject(box!, { x: 0, y: 0, wrap: 'behind' });
    expect([elements(anchor).filter((e) => e.name.startsWith('wp:wrap')).map((e) => e.name), attr(anchor, 'behindDoc')]).toEqual([['wp:wrapNone'], '1']);
    await editor.placeObject(box!, { x: 0, y: 0, wrap: 'inFront', from: 'page' });
    expect([elements(anchor).filter((e) => e.name.startsWith('wp:wrap')).map((e) => e.name), attr(anchor, 'behindDoc')]).toEqual([['wp:wrapNone'], '0']);
    expect(position(anchor, 'H')).toEqual(['page', '0']);
    expect(editor.revisions()).toEqual([]);
    await expect(editor.placeObject(box!, { x: Number.NaN, y: 0 })).rejects.toThrow(EditError);

    // A box Word placed and sized relative to the page (percentages, with a VML copy).
    const wp14 = 'xmlns:wp14="http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing"';
    const wordBox =
      `<w:r><mc:AlternateContent><mc:Choice Requires="wps"><w:drawing><wp:anchor ${wp14} distT="0" distB="0" distL="114300" distR="114300" simplePos="1" relativeHeight="251660000" behindDoc="1" locked="0" layoutInCell="1" allowOverlap="1">` +
      '<wp:simplePos x="5" y="5"/><mc:AlternateContent><mc:Choice Requires="wp14"><wp:positionH relativeFrom="page"><wp14:pctPosHOffset>10000</wp14:pctPosHOffset></wp:positionH></mc:Choice>' +
      '<mc:Fallback><wp:positionH relativeFrom="page"><wp:posOffset>777240</wp:posOffset></wp:positionH></mc:Fallback></mc:AlternateContent><wp:positionV relativeFrom="paragraph"><wp:align>top</wp:align></wp:positionV>' +
      `<wp:extent cx="${200 * EMU}" cy="${80 * EMU}"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:wrapNone/><wp:docPr id="9" name="Text Box 9"/><wp:cNvGraphicFramePr/>` +
      `<a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">${textBoxShape('Word box')}</a:graphicData></a:graphic>` +
      '<wp14:sizeRelH relativeFrom="margin"><wp14:pctWidth>40000</wp14:pctWidth></wp14:sizeRelH><wp14:sizeRelV relativeFrom="margin"><wp14:pctHeight>0</wp14:pctHeight></wp14:sizeRelV></wp:anchor></w:drawing></mc:Choice>' +
      '<mc:Fallback><w:pict><v:rect style="position:absolute;width:200pt;height:80pt"><v:textbox><w:txbxContent><w:p><w:r><w:t>Old copy</w:t></w:r></w:p></w:txbxContent></v:textbox></v:rect></w:pict></mc:Fallback></mc:AlternateContent></w:r>';
    const word = await open(`<w:p>${wordBox}</w:p>`);
    const [placed] = word.editor.objects();
    const wordAnchor = frameOf(placed!);
    await word.editor.placeObject(placed!, { x: 36, y: 6 });
    expect(elements(wordAnchor).map((e) => e.name)).toEqual([...ANCHOR_CHILDREN.map((n) => (n === 'wp:wrapSquare' ? 'wp:wrapNone' : n)), 'wp14:sizeRelH', 'wp14:sizeRelV']);
    expect(position(wordAnchor, 'H')).toEqual(['page', String(36 * EMU)]);
    expect(position(wordAnchor, 'V')).toEqual(['paragraph', String(6 * EMU)]);
    expect([attr(wordAnchor, 'simplePos'), attr(wordAnchor, 'behindDoc')]).toEqual(['0', '1']);
    // Word's old copy of the box is stale now.
    expect(findAll(word.doc.documentXml.root, 'mc:Fallback')).toEqual([]);
    // A new size replaces the relative one.
    word.editor.resizeObject(placed!, { width: 100 });
    expect(elements(wordAnchor).map((e) => e.name).filter((n) => n.startsWith('wp14:'))).toEqual([]);
    const again = new DocxEditor(await reopen(word.doc));
    expect(again.objects()).toMatchObject([{ kind: 'shape', floating: true, width: 100, height: 40 }]);
  });

  it('re-wraps or re-references a floating object without a new position', async () => {
    const { doc, editor } = await open(p('Box: ') + p('After') + `<w:p>${picture(3)}</w:p>`);
    await editor.insertShape(endOf(editor, 0), { shape: 'textBox', text: 'Boxed', place: { x: 36, y: 12, from: 'margin' } });
    const [box, pic] = editor.objects();
    const anchor = frameOf(box!);
    const wraps = () => [elements(anchor).filter((e) => e.name.startsWith('wp:wrap')).map((e) => e.name), attr(anchor, 'behindDoc')];
    await editor.placeObject(box!, { wrap: 'behind' });
    expect(position(anchor, 'H')).toEqual(['margin', String(36 * EMU)]);
    expect(position(anchor, 'V')).toEqual(['margin', String(12 * EMU)]);
    expect(wraps()).toEqual([['wp:wrapNone'], '1']);
    expect(elements(anchor).map((e) => e.name)).toEqual(ANCHOR_CHILDREN.map((n) => (n === 'wp:wrapSquare' ? 'wp:wrapNone' : n)));
    // One side alone moves that side only; a reference alone keeps both offsets.
    await editor.placeObject(box!, { y: 24 });
    expect([position(anchor, 'H'), position(anchor, 'V')]).toEqual([
      ['margin', String(36 * EMU)],
      ['margin', String(24 * EMU)],
    ]);
    await editor.placeObject(box!, { from: 'page' });
    expect([position(anchor, 'H'), position(anchor, 'V')]).toEqual([
      ['page', String(36 * EMU)],
      ['page', String(24 * EMU)],
    ]);
    expect(wraps()).toEqual([['wp:wrapNone'], '1']);
    expect(editor.revisions()).toEqual([]);
    expect(new DocxEditor(await reopen(doc)).objects()[0]).toMatchObject({ kind: 'textBox', floating: true });

    // An inline object, or a new shape, still needs both x and y to float.
    const needsXY = new EditError('A floating object needs x and y in points');
    await expect(editor.placeObject(pic!, { wrap: 'square' })).rejects.toThrow(needsXY);
    await expect(editor.placeObject(pic!, { x: 5 })).rejects.toThrow(needsXY);
    await expect(editor.insertShape(endOf(editor, 2), { shape: 'rect', place: { wrap: 'behind' } })).rejects.toThrow(needsXY);
    expect(editor.objects().map((o) => o.floating)).toEqual([true, false]);

    // A position aligned rather than offset keeps its alignment under a new reference.
    const alignedPicture = pictureDrawing(4).replace(
      /<wp:inline[^>]*>/,
      '<wp:anchor distT="0" distB="0" distL="114300" distR="114300" simplePos="0" relativeHeight="251659264" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1"><wp:simplePos x="0" y="0"/>' +
        `<wp:positionH relativeFrom="page"><wp:align>center</wp:align></wp:positionH><wp:positionV relativeFrom="paragraph"><wp:posOffset>${6 * EMU}</wp:posOffset></wp:positionV>`,
    ).replace('<wp:effectExtent l="0" t="0" r="0" b="0"/>', '<wp:effectExtent l="0" t="0" r="0" b="0"/><wp:wrapSquare wrapText="bothSides"/>').replace('</wp:inline>', '</wp:anchor>');
    const aligned = await open(`<w:p><w:r>${alignedPicture}</w:r></w:p>`);
    const [floating] = aligned.editor.objects();
    await aligned.editor.placeObject(floating!, { from: 'margin', wrap: 'topAndBottom' });
    const frame = frameOf(floating!);
    expect(childPath(frame, 'wp:positionH', 'wp:align')).toBeDefined();
    expect(attr(firstChild(frame, 'wp:positionH'), 'relativeFrom')).toBe('margin');
    expect(position(frame, 'V')).toEqual(['margin', String(6 * EMU)]);
    expect(firstChild(frame, 'wp:wrapTopAndBottom')).toBeDefined();
  });

  it('switches between inline and floating', async () => {
    const { doc, editor } = await open(`<w:p><w:r><w:t xml:space="preserve">Logo: </w:t></w:r>${picture(3)}</w:p>`);
    const [pic] = editor.objects();
    const before = frameOf(pic!);
    const [docPr, extent, graphic] = ['wp:docPr', 'wp:extent', 'a:graphic'].map((n) => firstChild(before, n));
    await editor.placeObject(pic!, { x: 0, y: 0 });
    const anchor = frameOf(pic!);
    expect(elements(pic!.drawing).map((e) => e.name)).toEqual(['wp:anchor']);
    expect(elements(anchor).map((e) => e.name)).toEqual(ANCHOR_CHILDREN);
    expect(firstChild(anchor, 'wp:docPr')).toBe(docPr);
    expect(firstChild(anchor, 'wp:extent')).toBe(extent);
    expect(firstChild(anchor, 'a:graphic')).toBe(graphic);
    expect(position(anchor, 'H')).toEqual(['column', '0']);
    expect(position(anchor, 'V')).toEqual(['paragraph', '0']);
    expect(Number(attr(anchor, 'relativeHeight'))).toBe(251659264);
    expect(editor.objects()).toMatchObject([{ floating: true, width: 200, height: 100 }]);

    await editor.placeObject(editor.objects()[0]!, { inline: true });
    const back = frameOf(pic!);
    expect(back.name).toBe('wp:inline');
    expect(elements(back).map((e) => e.name)).toEqual(['wp:extent', 'wp:effectExtent', 'wp:docPr', 'wp:cNvGraphicFramePr', 'a:graphic']);
    expect(firstChild(back, 'wp:docPr')).toBe(docPr);
    expect(editor.objects()).toMatchObject([{ floating: false, width: 200, height: 100 }]);
    // Already inline: nothing to do.
    await editor.placeObject(editor.objects()[0]!, { inline: true });
    expect(frameOf(pic!)).toBe(back);

    // With Track Changes on, the inline original is deleted and a floating copy inserted right after it.
    editor.trackChanges = true;
    await editor.placeObject(editor.objects()[0]!, { x: 36, y: 0, wrap: 'inFront' });
    const root = doc.documentXml.root;
    const [del] = marked(root, 'w:del', 'wp:inline');
    const [ins] = marked(root, 'w:ins', 'wp:anchor');
    expect(del).toBeDefined();
    expect(ins).toBeDefined();
    const kids = elements(editor.paragraphs()[0]!.el);
    expect(kids.at(-2)).toBe(del);
    expect(kids.at(-1)).toBe(ins);
    expect(new Set(drawingIds(root)).size).toBe(2);
    expect(texts(editor)).toEqual(['Logo: ']);
    const accepted = new DocxEditor(await reopen(doc));
    accepted.acceptRevisions();
    expect(accepted.objects()).toMatchObject([{ kind: 'picture', floating: true, width: 200, height: 100 }]);
    expect(firstChild(frameOf(accepted.objects()[0]!), 'wp:wrapNone')).toBeDefined();
  });

  it('deletes an object, tracked with Track Changes on', async () => {
    const { doc, editor } = await open(`<w:p><w:r><w:t xml:space="preserve">Logo: </w:t>${pictureDrawing(1)}<w:t xml:space="preserve"> (2026)</w:t></w:r></w:p>` + `<w:p>${textBoxRun('Boxed')}</w:p>`);
    editor.deleteObject(editor.objects()[0]!);
    expect(editor.objects().map((o) => o.kind)).toEqual(['shape']);
    expect(texts(editor)).toEqual(['Logo:  (2026)', '', 'Boxed']);
    const root = doc.documentXml.root;
    expect(findAll(root, 'w:del')).toEqual([]);

    editor.trackChanges = true;
    const [box] = editor.objects();
    editor.deleteObject(box!);
    const [del] = marked(root, 'w:del', 'w:drawing');
    expect(elements(del!)[0]).toBe(box!.run);
    // Word's old copy of the box goes; the box's own text is marked deleted with it.
    expect(findAll(root, 'mc:Fallback')).toEqual([]);
    expect(findAll(box!.run, 'w:t')).toEqual([]);
    expect(editor.revisions().every((r) => r.kind === 'delete')).toBe(true);
    editor.rejectRevisions();
    expect(editor.objects().map((o) => o.kind)).toEqual(['shape']);
    expect(texts(editor)).toEqual(['Logo:  (2026)', '', 'Boxed']);
  });

  it('sets alt text', async () => {
    const { doc, editor } = await open(`<w:p>${picture()}</w:p>` + `<w:p>${textBoxRun('Boxed')}</w:p>`);
    const [pic, box] = editor.objects();
    editor.setObjectAlt(pic!, 'Company logo');
    expect(attr(firstChild(frameOf(pic!), 'wp:docPr'), 'descr')).toBe('Company logo');
    editor.trackChanges = true;
    editor.setObjectAlt(box!, 'Contact details');
    expect(attr(firstChild(frameOf(box!), 'wp:docPr'), 'descr')).toBe('Contact details');
    // Not tracked (Word doesn't), and the box's old copy is stale.
    expect(editor.revisions()).toEqual([]);
    expect(findAll(doc.documentXml.root, 'mc:Fallback')).toEqual([]);
    expect(new DocxEditor(await reopen(doc)).objects().map((o) => o.alt)).toEqual(['Company logo', 'Contact details']);
    editor.setObjectAlt(pic!, '');
    expect(attr(firstChild(frameOf(pic!), 'wp:docPr'), 'descr')).toBeUndefined();
  });

  it('refuses untrackable changes when the document only allows tracked changes', async () => {
    const doc = await DocxDocument.open(await buildDocx({ body: p('Chart: ') + p('Box: ') + p('Target'), settings: '<w:documentProtection w:edit="trackedChanges" w:enforcement="1"/>' }));
    const setup = new DocxEditor(doc, undefined, { trackChanges: false });
    await setup.insertChart(endOf(setup, 0), COLUMNS);
    await setup.insertShape(endOf(setup, 1), { shape: 'textBox', text: 'Boxed' });
    const editor = new DocxEditor(doc, new AccessPolicy(doc, { user: { id: 'u1', name: 'Sari' }, role: 'editor' }));
    expect(editor.trackChanges).toBe(true);
    const [chart, box] = editor.objects();
    const onlyTracked = /only allows changes that are tracked/;
    expect(() => editor.resizeObject(chart!, { width: 100 })).toThrow(onlyTracked);
    expect(() => editor.setObjectAlt(chart!, 'Sales chart')).toThrow(onlyTracked);
    await expect(editor.updateChart(chart!, { title: 'Changed' })).rejects.toThrow(onlyTracked);
    await expect(editor.placeObject(box!, { x: 10, y: 10 })).rejects.toThrow(onlyTracked);
    expect(extentOf(chart!)).toEqual([String(432 * EMU), String(252 * EMU)]);
    expect(editor.chartInfo(chart!).title).toBe('Sales');
    expect(position(frameOf(box!), 'H')).toEqual(['column', '0']);

    // Deleting, moving, switching and inserting are tracked, so they are allowed.
    const root = doc.documentXml.root;
    editor.deleteObject(box!);
    expect(marked(root, 'w:del', 'wps:wsp')).toHaveLength(1);
    const target = () => editor.find('Target')[0]!.end;
    await editor.moveObject(chart!, target());
    expect(marked(root, 'w:ins', 'c:chart')).toHaveLength(1);
    await editor.insertShape(target(), { shape: 'rect' });
    expect(marked(root, 'w:ins', 'wps:wsp')).toHaveLength(1);
    const moved = editor.objects().find((o) => marked(root, 'w:ins', 'c:chart')[0]!.children.includes(o.run))!;
    await editor.placeObject(moved, { x: 0, y: 0 });
    expect(marked(root, 'w:ins', 'wp:anchor').filter((i) => findAll(i, 'c:chart').length)).toHaveLength(1);
    expect(new Set(editor.revisions().map((r) => r.author))).toEqual(new Set(['Sari']));
  });

  it('refuses changes to an object in a locked region', async () => {
    const { editor } = await open(p('Chart: ') + p('After'));
    await editor.insertChart(endOf(editor), COLUMNS);
    editor.lockParagraphs(editor.paragraphs()[0]!.el, undefined, { title: 'Figures' });
    const [chart] = editor.objects();
    await expect(editor.updateChart(chart!, { title: 'Changed' })).rejects.toThrow(LockedRegionError);
    expect(() => editor.resizeObject(chart!, { width: 100 })).toThrow(LockedRegionError);
    expect(() => editor.deleteObject(chart!)).toThrow(LockedRegionError);
    expect(() => editor.setObjectAlt(chart!, 'Sales')).toThrow(LockedRegionError);
    await expect(editor.moveObject(chart!, endOf(editor, 1))).rejects.toThrow(LockedRegionError);
    await expect(editor.placeObject(chart!, { x: 0, y: 0 })).rejects.toThrow(LockedRegionError);
    expect(editor.objects()).toMatchObject([{ kind: 'chart', floating: false, width: 432, height: 252 }]);
    expect(editor.objects()[0]!.alt).toBeUndefined();
    expect(editor.chartInfo(chart!).title).toBe('Sales');

    // Nor can an object move into one.
    const other = await open(`<w:p>${picture()}</w:p>` + p('Locked text'));
    other.editor.lockParagraphs(other.editor.paragraphs()[1]!.el, undefined, { title: 'Terms' });
    await expect(other.editor.moveObject(other.editor.objects()[0]!, other.editor.find('Locked')[0]!.end)).rejects.toThrow(LockedRegionError);
    expect(where(other.editor)).toEqual([0]);
  });

  it('handles objects in table cells and text boxes', async () => {
    const cell = `<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid><w:tr><w:tc><w:p>${picture()}</w:p></w:tc></w:tr></w:tbl>`;
    const box = `<w:p>${textBoxRunHolding(`<w:p><w:r>${oval()}</w:r></w:p>`)}</w:p>`;
    const { doc, editor } = await open(p('Intro') + cell + box + p('Outro'));
    const objects = editor.objects();
    expect(objects.map((o) => [o.n, o.kind])).toEqual([
      [1, 'picture'],
      [2, 'shape'],
      [3, 'shape'],
    ]);
    // In paragraph order: the cell, the paragraph holding the box, the box's own paragraph.
    expect(where(editor)).toEqual([1, 2, 3]);
    const [pic, , inner] = objects;
    editor.resizeObject(pic!, { width: 100 });
    editor.resizeObject(inner!, { width: 25 });
    expect(extentOf(pic!)).toEqual([String(100 * EMU), String(50 * EMU)]);
    expect(extentOf(inner!)).toEqual([String(25 * EMU), String(25 * EMU)]);
    const ext = childPath(frameOf(inner!), 'a:graphic', 'a:graphicData', 'wps:wsp', 'wps:spPr', 'a:xfrm', 'a:ext')!;
    expect(ext.attrs).toEqual({ cx: String(25 * EMU), cy: String(25 * EMU) });
    // A change inside the box makes Word's old copy of it stale.
    expect(findAll(doc.documentXml.root, 'mc:Fallback')).toEqual([]);
    const ed = new DocxEditor(await reopen(doc));
    expect(ed.objects().map((o) => [o.width, o.height])).toEqual([
      [100, 50],
      [200, 80],
      [25, 25],
    ]);

    // A picture may move into the box; a shape may not (Word doesn't nest them).
    const inBox = () => ({ node: ed.paragraphs()[3]!.el, offset: 0 });
    const boxed = ed.objects()[0]!;
    await ed.moveObject(boxed, inBox());
    expect(where(ed)).toEqual([2, 3, 3]);
    await ed.insertShape(ed.find('Outro')[0]!.end, { shape: 'rect' });
    const rect = ed.objects().at(-1)!;
    await expect(ed.moveObject(rect, inBox())).rejects.toThrow(new EditError("Shapes and text boxes can't go inside a text box"));
    expect(where(ed)).toEqual([2, 3, 3, 4]);
    // In a box an object can only be inline: it can't float there, nor can a floating one move in.
    const inlineOnly = new EditError('An object inside a text box can only be inline');
    await expect(ed.placeObject(boxed, { x: 0, y: 0 })).rejects.toThrow(inlineOnly);
    await ed.moveObject(boxed, ed.find('Outro')[0]!.end);
    await ed.placeObject(boxed, { x: 0, y: 0 });
    await expect(ed.moveObject(boxed, inBox())).rejects.toThrow(inlineOnly);
    expect(ed.objects().map((o) => [o.kind, o.floating])).toEqual([
      ['shape', false],
      ['shape', false],
      ['picture', true],
      ['shape', true],
    ]);
  });

  it('moves only pictures into a text box', async () => {
    const DIAGRAM = 'http://schemas.openxmlformats.org/drawingml/2006/diagram';
    const smartArt = `<w:r>${inline('<dgm:relIds r:dm="rIdNone" r:lo="rIdNone" r:qs="rIdNone" r:cs="rIdNone"/>', 100, 50, DIAGRAM)}</w:r>`;
    // A graphic docxkeun doesn't know (it lists it as a picture) and one that only wraps a picture.
    const unknown = `<w:r>${inline('<x:thing xmlns:x="urn:example:graphic"/>', 30, 30, 'urn:example:graphic')}</w:r>`;
    const wrapped = `<w:r>${inline('<x:frame xmlns:x="urn:example:graphic"><a:blip r:embed="rIdNone"/></x:frame>', 30, 30, 'urn:example:graphic')}</w:r>`;
    const { editor } = await open(p('Chart: ') + `<w:p>${smartArt}${unknown}${wrapped}${picture(2)}</w:p>` + `<w:p>${textBoxRun('Inside the box')}</w:p>`);
    await editor.insertChart(endOf(editor, 0), COLUMNS);
    const inBox = () => editor.find('Inside the box')[0]!.end;
    const [chart, diagram, other, frame, pic] = editor.objects();
    expect(editor.objects().map((o) => o.kind)).toEqual(['chart', 'diagram', 'picture', 'picture', 'picture', 'shape']);
    for (const o of [chart, diagram, other, frame]) await expect(editor.moveObject(o!, inBox())).rejects.toThrow(new EditError('Only pictures can go inside a text box'));
    expect(where(editor)).toEqual([0, 1, 1, 1, 1, 2]);
    await editor.moveObject(pic!, inBox());
    expect(where(editor)).toEqual([0, 1, 1, 1, 2, 3]);
  });

  it('refuses to move or resize an old VML object but deletes it', async () => {
    const { editor } = await open(p('Before') + `<w:p>${vmlTextBoxRun('Old box')}</w:p>`);
    const [old] = editor.objects();
    expect(old!.vml).toBe(true);
    const refusal = new EditError('This is an old-style (VML) object: it can only be deleted');
    expect(() => editor.resizeObject(old!, { width: 10 })).toThrow(refusal);
    expect(() => editor.setObjectAlt(old!, 'Old')).toThrow(refusal);
    await expect(editor.moveObject(old!, endOf(editor, 0))).rejects.toThrow(refusal);
    await expect(editor.placeObject(old!, { x: 0, y: 0 })).rejects.toThrow(refusal);
    await expect(editor.placeObject(old!, { inline: true })).rejects.toThrow(refusal);
    expect(editor.objects()).toHaveLength(1);
    editor.deleteObject(old!);
    expect(editor.objects()).toEqual([]);
    expect(texts(editor)).toEqual(['Before', '']);
    // A stale object is refused rather than changed.
    expect(() => editor.deleteObject(old!)).toThrow(EditError);
  });
});

describe('rewriting text around objects', () => {
  const PICTURE = 'http://schemas.openxmlformats.org/drawingml/2006/picture';
  const pictureRun = `<w:r>${inline('<pic:pic><pic:blipFill><a:blip r:embed="rIdNone"/></pic:blipFill></pic:pic>', 20, 20, PICTURE)}</w:r>`;
  const t = (s: string) => `<w:r><w:t xml:space="preserve">${s}</w:t></w:r>`;
  /** "Head ‹picture›office ‹text box›here": a picture and a text box (holding "Inside") in the middle of the text. */
  const BODY = `<w:p>${t('Head ')}${pictureRun}${t('office ')}${textBoxRun('Inside')}${t('here')}</w:p>` + p('After');
  const O = OBJECT_CHAR;
  const texts = (editor: DocxEditor) => editor.paragraphs().map((para) => paragraphText(para).text);
  const withObjects = (editor: DocxEditor) => paragraphText(editor.paragraphs()[0]!, { objects: true }).text;
  const wholeFirst = (editor: DocxEditor) => editor.paragraphRange(editor.paragraphs()[0]!)!;

  it('keeps the objects of a rewritten range after the new text, in order', async () => {
    const { doc, editor } = await open(BODY);
    expect(withObjects(editor)).toBe(`Head ${O}office ${O}here`);
    expect(editor.replaceRange(wholeFirst(editor), 'New words', undefined, { keepObjects: true })).toBe(2);
    expect(withObjects(editor)).toBe(`New words${O}${O}`);
    expect(editor.objects().map((o) => o.kind)).toEqual(['picture', 'shape']);
    // The text box keeps its paragraph (and, unchanged, Word's old copy of it).
    expect(texts(editor)).toEqual(['New words', 'Inside', 'After']);
    expect(findAll(doc.documentXml.root, 'mc:Fallback')).toHaveLength(1);
    const again = new DocxEditor(await reopen(doc));
    expect(texts(again)).toEqual(['New words', 'Inside', 'After']);
    expect(again.objects().map((o) => o.kind)).toEqual(['picture', 'shape']);

    // Without the option the objects go with the text, as when a user selects and types over them.
    const plain = await open(BODY);
    expect(plain.editor.replaceRange(wholeFirst(plain.editor), 'New words')).toBe(0);
    expect(plain.editor.objects()).toEqual([]);
    expect(texts(plain.editor)).toEqual(['New words', 'After']);

    // An object someone inserted with Track Changes on stays their insertion.
    const theirs = await open(`<w:p>${t('Head ')}<w:ins w:id="7" w:author="Budi" w:date="2026-01-01T00:00:00Z">${pictureRun}</w:ins>${t('here')}</w:p>`);
    theirs.editor.replaceRange(wholeFirst(theirs.editor), 'New', undefined, { keepObjects: true });
    expect(withObjects(theirs.editor)).toBe(`New${O}`);
    const [ins] = findAll(theirs.doc.documentXml.root, 'w:ins');
    expect([attr(ins, 'w:author'), findAll(ins!, 'w:drawing').length]).toEqual(['Budi', 1]);
    expect(theirs.editor.revisions().map((r) => [r.kind, r.author])).toEqual([['insert', 'Budi']]);
  });

  it('keeps the objects unmarked with Track Changes on', async () => {
    const { doc, editor } = await open(BODY);
    editor.trackChanges = true;
    expect(editor.replaceRange(wholeFirst(editor), 'New words', undefined, { keepObjects: true })).toBe(2);
    const root = doc.documentXml.root;
    // The old text is marked deleted and the new text inserted; the objects are neither, nor is the box's text.
    expect(findAll(root, 'w:del').flatMap((d) => [...findAll(d, 'w:drawing'), ...findAll(d, 'mc:AlternateContent')])).toEqual([]);
    expect(findAll(root, 'w:ins').flatMap((d) => [...findAll(d, 'w:drawing'), ...findAll(d, 'mc:AlternateContent')])).toEqual([]);
    expect(editor.revisions().map((r) => [r.kind, r.text])).toEqual([
      ['delete', 'Head '],
      ['delete', 'office '],
      ['delete', 'here'],
      ['insert', 'New words'],
    ]);
    // (The text as read leaves out what is marked deleted.)
    expect(withObjects(editor)).toBe(`New words${O}${O}`);
    expect(editor.objects().map((o) => o.kind)).toEqual(['picture', 'shape']);

    const accepted = new DocxEditor(await reopen(doc));
    accepted.acceptRevisions();
    expect(withObjects(accepted)).toBe(`New words${O}${O}`);
    expect(texts(accepted)).toEqual(['New words', 'Inside', 'After']);
    const rejected = new DocxEditor(await reopen(doc));
    rejected.rejectRevisions();
    expect(withObjects(rejected)).toBe(`Head office here${O}${O}`);
    expect(texts(rejected)).toEqual(['Head office here', 'Inside', 'After']);
  });
});
