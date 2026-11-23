/** Objects in the editor: listing them; inserting, reading, updating and copying charts; inserting shapes and text boxes. */
import { describe, expect, it } from 'vitest';
import { DocxEditor, EditError, LockedRegionError, paragraphText, type DocPosition, type ShapeKind } from '../src/api/index.js';
import { ChartDataError, type NewChart } from '../src/charts/index.js';
import { workbookFor } from '../src/charts/workbook.js';
import { createDocx } from '../src/index.js';
import { layoutDocument, type PathItem, type TextItem } from '../src/layout/index.js';
import { DocxDocument, type ImageInline } from '../src/model/index.js';
import { CONTENT_TYPE, REL } from '../src/opc/index.js';
import { attr, childPath, elements, findAll, firstChild, NS, parseXml, textContent, type XmlElement } from '../src/xml/index.js';
import { readZip } from '../src/zip/zip.js';
import { buildDocx, p } from './helpers/docx.js';
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
});
