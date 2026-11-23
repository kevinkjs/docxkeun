/**
 * docxkeun-mcp with text boxes and objects: a text box's paragraphs are numbered right after the
 * paragraph holding it, every text tool works on them, objects get numbers, and live views get the
 * viewer's paragraph numbers.
 */
import { mkdtemp, readFile, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { Client } from '@modelcontextprotocol/client';
import { InMemoryTransport } from '@modelcontextprotocol/server';
import { afterEach, beforeEach, describe, expect, it } from 'vitest';
import { DocxDocument, type DocParagraph } from 'docxkeun';
import { DocxEditor, paragraphsOf, paragraphText, type DocPosition } from 'docxkeun/editor';
import { buildDocx, p } from '../../../test/helpers/docx.js';
import { EMU, inline, textBoxRun, textBoxRunHolding, textBoxShape, vmlTextBoxRun } from '../../../test/helpers/drawings.js';
import { serializeXml } from '../../../src/xml/index.js';
import { createServer } from '../src/index.js';
import { contextFor, numberedObjects, OPERATIONS, targetParagraphs } from '../src/operations.js';
import { listParagraphs, viewerParagraphs } from '../src/outline.js';

let dir: string;
const clients: Client[] = [];

beforeEach(async () => {
  dir = await mkdtemp(join(tmpdir(), 'docxkeun-mcp-objects-'));
});

afterEach(async () => {
  for (const c of clients.splice(0)) await c.close();
  await rm(dir, { recursive: true, force: true });
});

async function connect(): Promise<Client> {
  const server = await createServer({ folders: [dir] });
  const [clientSide, serverSide] = InMemoryTransport.createLinkedPair();
  await server.connect(serverSide);
  const client = new Client({ name: 'test', version: '1.0.0' });
  await client.connect(clientSide);
  clients.push(client);
  return client;
}

async function call(client: Client, name: string, args: Record<string, unknown>): Promise<{ text: string; isError: boolean }> {
  const r = await client.callTool({ name, arguments: args });
  const text = (r.content as Array<{ type: string; text?: string }>).map((c) => c.text ?? '').join('\n');
  return { text, isError: r.isError === true };
}

async function ok(client: Client, name: string, args: Record<string, unknown>): Promise<string> {
  const r = await call(client, name, args);
  if (r.isError) throw new Error(`${name} failed: ${r.text}`);
  return r.text;
}

const BOX = p('Before') + `<w:p>${textBoxRun('Inside the box')}</w:p>` + p('After');

async function write(body: string, name = 'boxes.docx'): Promise<string> {
  const file = join(dir, name);
  await writeFile(file, await buildDocx({ body }));
  return file;
}

async function reopen(file: string): Promise<DocxDocument> {
  return DocxDocument.open(new Uint8Array(await readFile(file)));
}

const texts = (doc: DocxDocument) => new DocxEditor(doc).paragraphs().map((x) => paragraphText(x).text);
/** The lines of read_document that show paragraphs. */
const paragraphLines = (read: string) => read.split('\n').filter((l) => l.startsWith('['));
const documentXml = async (file: string) => serializeXml((await reopen(file)).documentXml);
/** Whether a paragraph is inside a text box. */
const inBox = (para: DocParagraph, doc: DocxDocument) => !paragraphsOf(doc.body().blocks).some((x) => x.el === para.el);

describe('text box paragraphs', () => {
  it("numbers a text box's paragraphs right after the paragraph holding it", async () => {
    await write(BOX);
    const read = await ok(await connect(), 'read_document', { path: 'boxes.docx' });
    expect(read).toContain('boxes.docx: 4 paragraphs');
    expect(paragraphLines(read)).toEqual(['[1] Before', '[2] {object 1: text box "Inside the box" (200×80pt, inline)}', '[3] (text box 1) Inside the box', '[4] After']);
  });

  it('numbers paragraphs exactly as DocxEditor.paragraphs() lists them', async () => {
    const cell = (inner: string) => `<w:tc><w:tcPr><w:tcW w:w="4000" w:type="dxa"/></w:tcPr>${inner}</w:tc>`;
    const table = (inner: string) => `<w:tbl><w:tblPr><w:tblW w:w="4000" w:type="dxa"/></w:tblPr><w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid><w:tr>${cell(inner)}</w:tr></w:tbl>`;
    const group =
      `<wpg:wgp><wpg:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${600 * EMU}" cy="${80 * EMU}"/><a:chOff x="0" y="0"/><a:chExt cx="${600 * EMU}" cy="${80 * EMU}"/></a:xfrm></wpg:grpSpPr>` +
      `${textBoxShape('One')}<wpg:grpSp><wpg:grpSpPr/>${textBoxShape('Two')}</wpg:grpSp>${textBoxShape('Three')}</wpg:wgp>`;
    const body =
      p('Before') +
      `<w:p>${textBoxRun('Inside the box')}</w:p>` +
      table(`<w:p>${textBoxRun('In a cell')}</w:p>`) +
      `<w:p><w:r>${inline(group, 600, 80, 'http://schemas.microsoft.com/office/word/2010/wordprocessingGroup')}</w:r></w:p>` +
      `<w:p>${textBoxRunHolding(p('Above a table') + table(p('Boxed cell')))}</w:p>` +
      `<w:p>${vmlTextBoxRun('Old box')}</w:p>` +
      p('After');
    const doc = await reopen(await write(body));
    const entries = listParagraphs(doc);
    const paras = new DocxEditor(doc).paragraphs();
    expect(entries.map((e) => e.para.el)).toHaveLength(paras.length);
    entries.forEach((e, i) => expect(e.para.el).toBe(paras[i]!.el));
    expect(entries.map((e) => e.text)).toEqual(['Before', '', 'Inside the box', '', 'In a cell', '', 'One', 'Two', 'Three', '', 'Above a table', 'Boxed cell', '', 'Old box', 'After']);
    // Boxes are numbered in the order their paragraphs come; `host` is the paragraph holding the box.
    expect(entries.map((e) => (e.box ? `${e.box.index}@${e.box.host}` : ''))).toEqual(['', '', '1@2', '', '2@4', '', '3@6', '4@6', '5@6', '', '6@10', '6@10', '', '7@13', '']);
    // Table numbers stay those of the body's tables; a table inside a box is part of the box.
    expect(entries.map((e) => (e.cell ? `${e.cell.table}` : ''))).toEqual(['', '', '', '1', '', '', '', '', '', '', '', '', '', '', '']);
    // `view` is the viewer's numbering (paragraphsOf: body and table cells); box paragraphs take their host's.
    const viewer = paragraphsOf(doc.body().blocks);
    expect(entries.filter((e) => !e.box).map((e) => e.para.el)).toEqual(viewer.map((x) => x.el));
    expect(entries.map((e) => e.view)).toEqual([1, 2, 2, 3, 3, 4, 4, 4, 4, 5, 5, 5, 6, 6, 7]);
  });

  it('edits a text box paragraph with formatting tags, and drops the old copy of the box', async () => {
    const file = await write(BOX);
    const client = await connect();
    expect(await ok(client, 'edit_paragraph', { path: 'boxes.docx', paragraph: 3, text: '<b>New</b> text' })).toBe('Paragraph 3: text replaced.\nSaved boxes.docx.');
    const doc = await reopen(file);
    expect(texts(doc)).toEqual(['Before', '', 'New text', 'After']);
    const box = new DocxEditor(doc).paragraphs()[2]!;
    expect(inBox(box, doc)).toBe(true);
    expect(box.inlines.flatMap((i) => (i.kind === 'text' ? [[i.text, !!i.props.bold]] : []))).toEqual([
      ['New', true],
      [' text', false],
    ]);
    const xml = await documentXml(file);
    expect(xml).not.toContain('mc:Fallback');
    expect(xml).not.toContain('Old copy');
  });

  it('inserts paragraphs into a text box, and keeps one paragraph in it when deleting', async () => {
    const file = await write(BOX);
    const client = await connect();
    expect(await ok(client, 'insert_paragraphs', { path: 'boxes.docx', after: 3, paragraphs: [{ text: 'Second <i>line</i>' }] })).toBe('Inserted 1 paragraph (now paragraph 4).\nSaved boxes.docx.');
    let doc = await reopen(file);
    expect(texts(doc)).toEqual(['Before', '', 'Inside the box', 'Second line', 'After']);
    expect(inBox(new DocxEditor(doc).paragraphs()[3]!, doc)).toBe(true);
    expect(paragraphLines(await ok(client, 'read_document', { path: 'boxes.docx' }))).toContain('[4] (text box 1) Second <i>line</i>');

    await ok(client, 'delete_paragraphs', { path: 'boxes.docx', from: 3, to: 4 });
    doc = await reopen(file);
    // The box keeps one (empty) paragraph.
    expect(texts(doc)).toEqual(['Before', '', '', 'After']);
    expect(paragraphLines(await ok(client, 'read_document', { path: 'boxes.docx' }))).toContain('[3] (text box 1)');
  });

  it('still needs a body paragraph when only text box paragraphs would be left', async () => {
    await write(p('Before') + `<w:p>${textBoxRunHolding(p('One') + p('Two'))}</w:p>`);
    const r = await call(await connect(), 'delete_paragraphs', { path: 'boxes.docx', from: 1, to: 2 });
    expect(r.isError).toBe(true);
    expect(r.text).toContain('A document needs at least one paragraph outside tables');
  });

  it('finds and replaces text in text boxes, and still lists them as stories', async () => {
    const file = await write(BOX);
    const client = await connect();
    expect(await ok(client, 'find_text', { path: 'boxes.docx', query: 'Inside' })).toBe('1 match for "Inside":\n[3] «Inside» the box');
    expect(await ok(client, 'replace_text', { path: 'boxes.docx', find: 'Inside', replace: 'Within' })).toContain('Replaced 1 of 1 match');
    expect(texts(await reopen(file))).toEqual(['Before', '', 'Within the box', 'After']);
    expect(await ok(client, 'read_headers_footers', { path: 'boxes.docx' })).toBe('textbox1: Text box 1 (in paragraph 2)\n  Within the box');
  });

  it('edits a text box through replace_story_text, dropping its old copy', async () => {
    const file = await write(BOX);
    const client = await connect();
    expect(await ok(client, 'replace_story_text', { path: 'boxes.docx', story: 'textbox1', find: 'Inside', replace: 'Within' })).toContain('in Text box 1 (in paragraph 2)');
    expect(texts(await reopen(file))).toEqual(['Before', '', 'Within the box', 'After']);
    expect(await documentXml(file)).not.toContain('mc:Fallback');
  });

  it('formats, comments on and lists text box paragraphs by number', async () => {
    const file = await write(BOX);
    const client = await connect();
    await ok(client, 'format_text', { path: 'boxes.docx', paragraph: 3, text: 'box', format: { italic: true } });
    await ok(client, 'format_paragraph', { path: 'boxes.docx', from: 3, align: 'center' });
    const added = await ok(client, 'add_comment', { path: 'boxes.docx', paragraph: 3, text: 'Inside', comment: 'Check' });
    expect(added).toContain('on "Inside" in paragraph 3.');
    const id = /comment (\S+) on/.exec(added)![1]!;
    const read = await ok(client, 'read_document', { path: 'boxes.docx' });
    expect(paragraphLines(read)).toContain(`[3] (text box 1; centered; comment ${id}) Inside the <i>box</i>`);
    expect(await ok(client, 'list_comments', { path: 'boxes.docx' })).toContain('On paragraph 3: "Inside"');
    expect(texts(await reopen(file))).toEqual(['Before', '', 'Inside the box', 'After']);
  });

  it('refuses to copy paragraphs from or into a text box', async () => {
    await write(BOX);
    const client = await connect();
    const message = 'copy_paragraphs copies body paragraphs and tables; to put text in a text box use insert_paragraphs or edit_paragraph';
    for (const args of [{ from: 3 }, { from: 1, to: 3 }, { from: 1, after: 3 }, { from: 4, before: 3 }]) {
      const r = await call(client, 'copy_paragraphs', { path: 'boxes.docx', ...args });
      expect(r.isError).toBe(true);
      expect(r.text).toContain(message);
    }
    // The holding paragraph is a body paragraph: copying it copies the box too.
    expect(await ok(client, 'copy_paragraphs', { path: 'boxes.docx', from: 2 })).toContain('the copy is paragraph 5');
    expect(texts(await reopen(join(dir, 'boxes.docx')))).toEqual(['Before', '', 'Inside the box', 'After', '', 'Inside the box']);
  });
});

describe('object numbers', () => {
  async function objectsDocument(): Promise<string> {
    const doc = await DocxDocument.open(await buildDocx({ body: BOX + p('Chart here') }));
    const editor = new DocxEditor(doc);
    const end = (i: number): DocPosition => editor.paragraphRange(editor.paragraphs()[i]!)!.end;
    await editor.insertChart(end(4), { type: 'column', title: 'Revenue', categories: ['Q1', 'Q2'], series: [{ name: 'Sales', values: [4.3, 2.5] }] });
    // At the end of the text, so before the chart.
    await editor.insertShape(end(4), { shape: 'rect', width: 100, height: 50 });
    const file = join(dir, 'objects.docx');
    await writeFile(file, await doc.save());
    return file;
  }

  it('numbers every object in read_document and list_objects, saying inline or floating', async () => {
    await objectsDocument();
    const client = await connect();
    const listed = await ok(client, 'list_objects', { path: 'objects.docx' });
    expect(listed.split('\n')).toEqual([
      '[2] object 1: text box "Inside the box" (200×80pt, inline)',
      '[5] object 2: shape (rect) (100×50pt, floating)',
      '[5] object 3: clustered column chart "Revenue": series Sales (Q1 4.3, Q2 2.5) (432×252pt, inline)',
    ]);
    const read = await ok(client, 'read_document', { path: 'objects.docx' });
    expect(paragraphLines(read)[4]).toBe('[5] Chart here {object 2: shape (rect) (100×50pt, floating)} {object 3: clustered column chart "Revenue": series Sales (Q1 4.3, Q2 2.5) (432×252pt, inline)}');
  });

  it('matches the numbers DocxEditor.objects() gives', async () => {
    const doc = await reopen(await objectsDocument());
    const editor = new DocxEditor(doc);
    const entries = listParagraphs(doc);
    const numbered = numberedObjects(editor, entries);
    expect(numbered.map((o) => [o.n, o.paragraph])).toEqual([
      [1, 2],
      [2, 5],
      [3, 5],
    ]);
    editor.objects().forEach((o, i) => {
      expect(numbered[i]!.object.drawing).toBe(o.drawing);
      expect(entries[numbered[i]!.paragraph - 1]!.para.el).toBe(o.paragraph.el);
    });
  });
});

describe('live views', () => {
  it("translates read_document's numbers to the viewer's", async () => {
    const doc = await reopen(await write(BOX));
    const entries = listParagraphs(doc);
    expect(viewerParagraphs(entries, [3, 4])).toEqual([2, 3]);
    expect(viewerParagraphs(entries, [1, 2, 3])).toEqual([1, 2]);
    expect(viewerParagraphs(entries, [])).toEqual([]);
    // Numbers past the end are left out.
    expect(viewerParagraphs(entries, [4, 9])).toEqual([3]);
  });

  it('works on the paragraph holding an object', async () => {
    const doc = await reopen(await write(BOX));
    const entries = listParagraphs(doc);
    const objects = numberedObjects(new DocxEditor(doc), entries);
    expect(targetParagraphs(entries, 'delete_object', { object: 1 }, objects)).toEqual([2]);
    expect(targetParagraphs(entries, 'delete_object', { object: 7 }, objects)).toEqual([]);
    expect(targetParagraphs(entries, 'delete_object', { object: 1 })).toEqual([]);
  });

  it('works on the paragraph edit_paragraph rewrites, not where its new text is', async () => {
    const entries = listParagraphs(await reopen(await write(BOX)));
    expect(targetParagraphs(entries, 'edit_paragraph', { paragraph: 3, text: 'New text' })).toEqual([3]);
    expect(viewerParagraphs(entries, targetParagraphs(entries, 'edit_paragraph', { paragraph: 3, text: 'New text' }))).toEqual([2]);
  });
});

describe('chart tools', () => {
  const revenue = {
    path: 'charts.docx',
    paragraph: 2,
    type: 'column',
    title: 'Revenue',
    categories: ['Q1', 'Q2', 'Q3', 'Q4'],
    series: [
      { name: 'Revenue', values: [4.3, 2.5, 3.5, 4.5] },
      { name: 'Costs', values: [2.4, 4.4, 1.8, 2.8] },
    ],
    categoryTitle: 'Quarter',
    valueTitle: 'USD m',
  };
  const chartInfo = async (file: string, n = 1) => {
    const editor = new DocxEditor(await reopen(file));
    return editor.chartInfo(editor.objects().find((o) => o.n === n)!);
  };

  it('inserts a chart, lists it and reads it back as a table', async () => {
    const file = await write(p('Intro') + p('Chart goes here'), 'charts.docx');
    const client = await connect();
    expect(await ok(client, 'insert_chart', revenue)).toBe('Inserted a clustered column chart (object 1) in paragraph 2: 2 series × 4 categories.\nSaved charts.docx.');
    expect((await ok(client, 'list_objects', { path: 'charts.docx' })).split('\n')[0]).toMatch(/^\[2\] object 1: clustered column chart "Revenue"/);
    expect(await ok(client, 'read_chart', { path: 'charts.docx', object: 1 })).toBe(
      [
        'Chart (object 1, paragraph 2): clustered column, title "Revenue", category axis "Quarter", value axis "USD m", legend bottom. Data in the embedded workbook.',
        '| Q1 | Q2 | Q3 | Q4 |',
        'Revenue | 4.3 | 2.5 | 3.5 | 4.5',
        'Costs | 2.4 | 4.4 | 1.8 | 2.8',
      ].join('\n'),
    );
    expect((await chartInfo(file)).series).toHaveLength(2);
  });

  it('shows a missing value as an empty cell', async () => {
    await write(p('Intro') + p('Here'), 'charts.docx');
    const client = await connect();
    await ok(client, 'insert_chart', { ...revenue, type: 'line', series: [{ name: 'Costs', values: [1, null, 3, 4] }] });
    expect((await ok(client, 'read_chart', { path: 'charts.docx', object: 1 })).split('\n').at(-1)).toBe('Costs | 1 |  | 3 | 4');
  });

  it('updates the data and titles, keeping the chart', async () => {
    const file = await write(p('Intro') + p('Chart goes here'), 'charts.docx');
    const client = await connect();
    await ok(client, 'insert_chart', revenue);
    const text = await ok(client, 'update_chart', {
      path: 'charts.docx',
      object: 1,
      title: 'Sales',
      series: [...revenue.series, { name: 'Profit', values: [1.9, -1.9, 1.7, 1.7] }],
    });
    expect(text).toContain('Changed the title, data of the chart (object 1, paragraph 2).');
    const info = await chartInfo(file);
    expect(info.title).toBe('Sales');
    expect(info.series.map((s) => s.name)).toEqual(['Revenue', 'Costs', 'Profit']);
    expect(info.axisTitles).toEqual({ category: 'Quarter', value: 'USD m' });
    expect(info.source).toBe('embedded');
  });

  it('refuses data that does not fit, with the library message', async () => {
    await write(p('Intro') + p('Chart goes here'), 'charts.docx');
    const client = await connect();
    const bad = await call(client, 'insert_chart', { ...revenue, series: [{ name: 'Costs', values: [1, 2, 3] }] });
    expect(bad.isError).toBe(true);
    expect(bad.text).toContain('series "Costs" has 3 values for 4 categories');
    await ok(client, 'insert_chart', revenue);
    const bad2 = await call(client, 'update_chart', { path: 'charts.docx', object: 1, series: [{ name: 'Costs', values: [1, 2, 3] }] });
    expect(bad2.isError).toBe(true);
    expect(bad2.text).toContain('series "Costs" has 3 values for 4 categories');
  });

  it('says when an object is not a chart or does not exist, or nothing is asked', async () => {
    await write(BOX, 'charts.docx');
    const client = await connect();
    expect((await call(client, 'read_chart', { path: 'charts.docx', object: 1 })).text).toContain('Object 1 is a text box, not a chart.');
    expect((await call(client, 'update_chart', { path: 'charts.docx', object: 1, title: 'X' })).text).toContain('Object 1 is a text box, not a chart.');
    expect((await call(client, 'read_chart', { path: 'charts.docx', object: 5 })).text).toContain('There is no object 5');
    await ok(client, 'insert_chart', { ...revenue, paragraph: 1 });
    expect((await call(client, 'update_chart', { path: 'charts.docx', object: 1 })).text).toContain('Nothing to change');
  });

  it('rejects wrong types and sizes from JSON callers', async () => {
    await write(p('Intro') + p('Here'), 'charts.docx');
    const client = await connect();
    expect((await call(client, 'insert_chart', { ...revenue, width: '300' })).isError).toBe(true);
    expect((await call(client, 'insert_chart', { ...revenue, height: 5000 })).isError).toBe(true);
    expect((await call(client, 'insert_chart', { ...revenue, categories: [] })).isError).toBe(true);
    expect((await call(client, 'insert_chart', { ...revenue, series: [{ name: 'A', values: ['1', '2', '3', '4'] }] })).isError).toBe(true);
  });

  it('inserts as a tracked change, but chart data changes are never tracked', async () => {
    const file = await write(p('Intro') + p('Chart goes here'), 'charts.docx');
    const client = await connect();
    await ok(client, 'set_track_changes', { path: 'charts.docx', on: true });
    await ok(client, 'insert_chart', revenue);
    expect(await ok(client, 'list_revisions', { path: 'charts.docx' })).toMatch(/insert/i);
    const text = await ok(client, 'update_chart', { path: 'charts.docx', object: 1, title: 'Sales' });
    expect(text).toContain("Track Changes can't record chart data changes, so this change is not tracked.");
    expect((await chartInfo(file)).title).toBe('Sales');
  });
});

describe('shape and object tools', () => {
  const NAME = 'shapes.docx';
  const open = async () => new DocxEditor(await reopen(join(dir, NAME)));

  it('inserts a floating text box with tagged text as paragraphs, and says where they are', async () => {
    await write(p('Intro') + p('Here'), NAME);
    const client = await connect();
    const text = await ok(client, 'insert_shape', { path: NAME, paragraph: 2, shape: 'text box', text: '<b>Hello</b> there\nSecond line' });
    expect(text).toBe(
      `Added a text box (object 1) in paragraph 2, floating 0 pt right and 0 pt down from the paragraph, with 2 paragraphs (paragraphs 3\u20134).\nSaved ${NAME}.`,
    );
    const read = paragraphLines(await ok(client, 'read_document', { path: NAME }));
    expect(read[2]).toBe('[3] (text box 1) <b>Hello</b> there');
    expect(read[3]).toBe('[4] (text box 1) Second line');
    const editor = await open();
    expect(editor.objects()[0]).toMatchObject({ floating: true });
    expect(editor.find('Hello')).toHaveLength(1);
  });

  it('inserts an inline shape, or a floating one at a position with a wrap style and alt text', async () => {
    await write(p('Intro') + p('Here'), NAME);
    const client = await connect();
    expect(await ok(client, 'insert_shape', { path: NAME, paragraph: 2, shape: 'oval', placement: 'inline', width: 50, height: 40, fill: '00AA00', alt: 'A dot' })).toContain(
      'Added an oval (object 1) in paragraph 2, inline.',
    );
    const floating = await ok(client, 'insert_shape', { path: NAME, paragraph: 1, shape: 'rounded callout', text: 'Look', x: 36, y: -10, from: 'page', wrap: 'in front of text' });
    expect(floating).toContain('Added a rounded callout (object 1) in paragraph 1, floating 36 pt right and -10 pt down from the page');
    expect(floating).toContain('in front of text');
    const objects = (await open()).objects();
    expect(objects.map((o) => [o.floating, o.alt])).toEqual([
      [true, undefined],
      [false, 'A dot'],
    ]);
    expect(await ok(client, 'list_objects', { path: NAME })).toContain('inline');
  });

  it('refuses a position for an inline shape, and a shape inside a text box, with plain messages', async () => {
    await write(BOX, NAME);
    const client = await connect();
    expect((await call(client, 'insert_shape', { path: NAME, paragraph: 1, shape: 'rectangle', placement: 'inline', x: 5, y: 5 })).text).toContain('x, y, from and wrap only apply to floating shapes');
    expect((await call(client, 'insert_shape', { path: NAME, paragraph: 3, shape: 'rectangle' })).text).toContain("Shapes and text boxes can't go inside a text box");
  });

  it('rejects wrong types and sizes from JSON callers', async () => {
    await write(p('Intro') + p('Here'), NAME);
    const client = await connect();
    const base = { path: NAME, paragraph: 2, shape: 'rectangle' };
    for (const bad of [{ width: '300' }, { width: 0 }, { height: 2001 }, { x: 2001 }, { y: -2001 }, { x: '5' }, { shape: 'hexagon' }, { wrap: 'around' }, { from: 'cell' }, { fill: 5 }, { text: 5 }]) {
      expect((await call(client, 'insert_shape', { ...base, ...bad })).isError, JSON.stringify(bad)).toBe(true);
    }
  });

  it('highlights the target paragraph, not a search for the shape text', () => {
    const entries = [
      { n: 1, text: 'Hello world' },
      { n: 2, text: 'Another line' },
    ];
    expect(targetParagraphs(entries, 'insert_shape', { paragraph: 2, text: 'Hello world' })).toEqual([2]);
  });

  it('resizes an object, sets its alt text and places it, saying what changed', async () => {
    await write(p('Intro') + p('Here'), NAME);
    const client = await connect();
    await ok(client, 'insert_shape', { path: NAME, paragraph: 2, shape: 'rectangle', width: 100, height: 50 });
    expect(await call(client, 'edit_object', { path: NAME, object: 1 })).toMatchObject({ isError: true, text: expect.stringContaining('Say what to change: size, place or alt text.') });
    const text = await ok(client, 'edit_object', { path: NAME, object: 1, width: 200, height: 80, alt: 'A box', placement: 'floating', x: 10, y: 20, from: 'margin', wrap: 'behind text' });
    expect(text).toContain('Changed the shape (object 1, paragraph 2): size 200 × 80 pt, alt text, position.');
    const o = (await open()).objects()[0]!;
    expect([o.width, o.height, o.alt]).toEqual([200, 80, 'A box']);
    expect(await ok(client, 'edit_object', { path: NAME, object: 1, placement: 'inline' })).toContain('now inline');
    expect((await open()).objects()[0]!.floating).toBe(false);
    expect((await call(client, 'edit_object', { path: NAME, object: 1, placement: 'floating' })).text).toContain('needs x and y');
    expect((await call(client, 'edit_object', { path: NAME, object: 9, width: 5 })).text).toContain('There is no object 9');
    expect((await call(client, 'edit_object', { path: NAME, object: 1, width: 5000 })).isError).toBe(true);
  });

  it('moves an object to another paragraph, after some text', async () => {
    await write(p('Intro') + p('Here, now') + p('There'), NAME);
    const client = await connect();
    await ok(client, 'insert_shape', { path: NAME, paragraph: 1, shape: 'diamond' });
    const text = await ok(client, 'edit_object', { path: NAME, object: 1, paragraph: 3 });
    expect(text).toContain('moved to paragraph 3');
    const after = await open();
    expect(after.paragraphs().indexOf(after.objects()[0]!.paragraph)).toBe(2);
    expect(await ok(client, 'edit_object', { path: NAME, object: 1, paragraph: 2, after: 'Here' })).toContain('moved to paragraph 2');
  });

  it('deletes an object, and says when that is tracked', async () => {
    await write(p('Intro') + p('Here'), NAME);
    const client = await connect();
    await ok(client, 'insert_chart', { path: NAME, paragraph: 1, type: 'column', categories: ['A'], series: [{ name: 'S', values: [1] }] });
    await ok(client, 'insert_shape', { path: NAME, paragraph: 2, shape: 'star' });
    expect(await ok(client, 'delete_object', { path: NAME, object: 2 })).toContain('Deleted the shape (object 2) from paragraph 2.');
    expect((await open()).objects()).toHaveLength(1);
    await ok(client, 'set_track_changes', { path: NAME, on: true });
    const tracked = await ok(client, 'delete_object', { path: NAME, object: 1 });
    expect(tracked).toContain('Deleted the clustered column chart (object 1) from paragraph 1 (tracked).');
    expect(await ok(client, 'list_revisions', { path: NAME })).toMatch(/delet/i);
    expect((await call(client, 'delete_object', { path: NAME, object: 7 })).text).toContain('There is no object 7');
  });

  it('says which parts of an edit are not tracked when Track Changes is on', async () => {
    await write(p('Intro') + p('Here'), NAME);
    const client = await connect();
    await ok(client, 'insert_shape', { path: NAME, paragraph: 1, shape: 'rectangle' });
    await ok(client, 'set_track_changes', { path: NAME, on: true });
    const text = await ok(client, 'edit_object', { path: NAME, object: 1, width: 90, alt: 'Box', x: 5, y: 5 });
    expect(text).toContain("Track Changes can't record size, alt text or position, so those changes are not tracked.");
    const moved = await ok(client, 'edit_object', { path: NAME, object: 1, paragraph: 2 });
    expect(moved).toContain('moved to paragraph 2');
    expect(moved).not.toContain("can't record");
    expect((await call(client, 'edit_object', { path: NAME, object: 1, paragraph: 1, x: 1, y: 1 })).text).toContain('separate calls');
  });

  it('refuses an untracked change under tracked-changes-only protection, but deletes', async () => {
    const file = join(dir, NAME);
    const setup = await DocxDocument.open(await buildDocx({ body: p('Intro') + p('Here'), settings: '<w:documentProtection w:edit="trackedChanges" w:enforcement="1"/>' }));
    const editor = new DocxEditor(setup, undefined, { trackChanges: false });
    await editor.insertShape(editor.paragraphRange(editor.paragraphs()[1]!)!.end, { shape: 'rect' });
    await writeFile(file, await setup.save());
    const client = await connect();
    const refused = await call(client, 'edit_object', { path: NAME, object: 1, width: 50 });
    expect(refused.isError).toBe(true);
    expect(refused.text).toContain('only allows changes that are tracked');
    expect(await ok(client, 'delete_object', { path: NAME, object: 1 })).toContain('(tracked)');
  });

  it('shows an image the object numbering could not pair, without a number', async () => {
    await write(p('Intro') + p('Here'), NAME);
    const client = await connect();
    await ok(client, 'insert_shape', { path: NAME, paragraph: 2, shape: 'rectangle' });
    const doc = await reopen(join(dir, NAME));
    const editor = new DocxEditor(doc);
    const blind = new Proxy(editor, { get: (t, k) => (k === 'objects' ? () => [] : Reflect.get(t, k, t)) });
    for (const name of ['list_objects', 'read_document']) {
      const out = await OPERATIONS.find((op) => op.name === name)!.run(contextFor(doc, blind as DocxEditor, NAME, 'AI'), {});
      expect(out, name).toContain('shape (rect)');
      expect(out, name).not.toContain('object 1');
    }
  });

  it('describes text boxes where it describes table cells', async () => {
    const client = await connect();
    expect(client.getInstructions()).toContain('text box');
    const tools = (await client.listTools()).tools;
    const desc = (n: string) => tools.find((t) => t.name === n)!.description!;
    expect(desc('insert_paragraphs')).toContain('text box');
    expect(desc('delete_paragraphs')).toContain('text box');
    expect(desc('copy_paragraphs')).toContain('text box');
  });

  it('refuses a footnote in a text box', async () => {
    await write(BOX, NAME);
    const client = await connect();
    expect((await call(client, 'add_footnote', { path: NAME, paragraph: 3, text: 'Note' })).text).toContain("Footnotes can't go in a text box");
  });

  it('tells the AI about charts, shapes and objects in its instructions', async () => {
    const instructions = (await connect()).getInstructions()!;
    for (const word of ['insert_chart', 'read_chart', 'update_chart', 'insert_shape', 'edit_object', 'delete_object', 'list_objects']) expect(instructions).toContain(word);
  });
});
