/** Copying blocks (with tables and lists), building tables, editing their rows, and set_list's friendlier rules. */
import { describe, expect, it } from 'vitest';
import { DocxEditor, paragraphsOf, paragraphText } from '../src/api/index.js';
import { DocxDocument } from '../src/model/index.js';
import { elements, findAll, firstChild, type XmlElement } from '../src/xml/index.js';
import { buildDocx, p } from './helpers/docx.js';

const cell = (t: string) => `<w:tc><w:tcPr><w:tcW w:w="2000" w:type="dxa"/><w:shd w:val="clear" w:fill="DDEEFF"/></w:tcPr>${p(t)}</w:tc>`;
const table = (rows: string[][]) => `<w:tbl><w:tblPr><w:tblW w:w="4000" w:type="dxa"/></w:tblPr><w:tblGrid><w:gridCol w:w="2000"/><w:gridCol w:w="2000"/></w:tblGrid>${rows.map((r) => `<w:tr>${r.map(cell).join('')}</w:tr>`).join('')}</w:tbl>`;
const TABLE_GRID = '<w:style w:type="table" w:styleId="TableGrid"><w:name w:val="Table Grid"/><w:tblPr><w:tblBorders><w:top w:val="single" w:sz="4"/></w:tblBorders></w:tblPr></w:style>';

async function open(body: string, styles = '') {
  const doc = await DocxDocument.open(await buildDocx({ body, styles }));
  return { doc, editor: new DocxEditor(doc), paras: () => paragraphsOf(doc.body().blocks), texts: () => paragraphsOf(doc.body().blocks).map((x) => paragraphText(x).text) };
}
const reopen = async (doc: DocxDocument) => DocxDocument.open(await doc.save());
const bodyBlocks = (doc: DocxDocument) => elements(firstChild(doc.documentXml.root, 'w:body')!).filter((c) => c.name !== 'w:sectPr');

describe('copying blocks', () => {
  it('copies paragraphs and a table whole, formatting included, onto a new page at the end', async () => {
    const t = await open(p('Title', '<w:jc w:val="center"/>', '<w:b/>') + p('Intro') + table([['No', 'Item'], ['1', 'Land']]) + p('After the table'));
    const [first, , , , , , last] = t.paras();
    const { blocks, dropped } = await t.editor.copyBlocks(first!.el, last!.el, 'end', { pageBreak: true });
    expect(dropped).toEqual({ comments: 0, notes: 0, bookmarks: 0 });
    expect(blocks.map((b) => b.name)).toEqual(['w:p', 'w:p', 'w:tbl', 'w:p']);
    expect(t.texts()).toEqual(['Title', 'Intro', 'No', 'Item', '1', 'Land', 'After the table', 'Title', 'Intro', 'No', 'Item', '1', 'Land', 'After the table']);
    const copyTitle = t.paras()[7]!;
    expect(copyTitle.props.pageBreakBefore).toBe(true);
    expect(copyTitle.props.jc).toBe('center');
    expect(findAll(copyTitle.el, 'w:b')).toHaveLength(1);
    // The copied table keeps its cell shading.
    expect(findAll(blocks[2]!, 'w:shd')).toHaveLength(4);
    const again = await reopen(t.doc);
    expect(paragraphsOf(again.body().blocks)).toHaveLength(14);
  });

  it('copies the whole table when a range starts or ends inside it', async () => {
    const t = await open(p('Before') + table([['a', 'b'], ['c', 'd']]) + p('After'));
    // From the 3rd cell paragraph ("c") to "After": the table is copied whole.
    const ps = t.paras();
    const { blocks } = await t.editor.copyBlocks(ps[3]!.el, ps[5]!.el, { after: ps[5]!.el });
    expect(blocks.map((b) => b.name)).toEqual(['w:tbl', 'w:p']);
    expect(t.texts().slice(6)).toEqual(['a', 'b', 'c', 'd', 'After']);
  });

  it('starts the copy of a list again at 1, or continues it when asked', async () => {
    const t = await open(p('One') + p('Two') + p('Three'));
    await t.editor.setList(t.paras().map((x) => x.el), { format: 'decimal' });
    const [a, , c] = t.paras();
    await t.editor.copyBlocks(a!.el, c!.el, 'end');
    const labels = () => {
      const items = t.editor.listItems();
      return t.paras().map((x) => items.get(x.el)?.label);
    };
    expect(labels()).toEqual(['1.', '2.', '3.', '1.', '2.', '3.']);
    const u = await open(p('One') + p('Two'));
    await u.editor.setList(u.paras().map((x) => x.el), { format: 'decimal' });
    const [x, y] = u.paras();
    await u.editor.copyBlocks(x!.el, y!.el, 'end', { restartLists: false });
    const items = u.editor.listItems();
    expect(u.paras().map((q) => items.get(q.el)?.label)).toEqual(['1.', '2.', '3.', '4.']);
  });

  it('leaves out what may exist only once, and gives drawings and Word ids fresh values', async () => {
    const t = await open(
      '<w:p w14:paraId="11111111" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"><w:bookmarkStart w:id="0" w:name="terms"/><w:r><w:t>Terms</w:t></w:r><w:bookmarkEnd w:id="0"/></w:p>' + p('Body'),
    );
    const [terms, body] = t.paras();
    await t.editor.addComment({ start: { node: findAll(body!.el, 'w:t')[0]!, offset: 0 }, end: { node: findAll(body!.el, 'w:t')[0]!, offset: 4 } }, { author: 'Rina', text: 'Check' });
    const { blocks, dropped } = await t.editor.copyBlocks(terms!.el, t.paras()[1]!.el, 'end');
    expect(dropped).toEqual({ comments: 1, notes: 0, bookmarks: 1 });
    for (const b of blocks) {
      expect(findAll(b, 'w:bookmarkStart')).toHaveLength(0);
      expect(findAll(b, 'w:commentReference')).toHaveLength(0);
      expect(findAll(b, 'w:commentRangeStart')).toHaveLength(0);
      expect(b.attrs['w14:paraId']).toBeUndefined();
    }
    // The original keeps its bookmark and comment, and the copied text is there.
    expect(t.editor.bookmarks().map((x) => x.name)).toEqual(['terms']);
    expect(findAll(t.doc.documentXml.root, 'w:commentReference')).toHaveLength(1);
    expect(t.texts()).toEqual(['Terms', 'Body', 'Terms', 'Body']);
    await reopen(t.doc);
  });

  it('puts a page-break paragraph first when the copy starts with a table', async () => {
    const t = await open(table([['a', 'b']]) + p('End'));
    const ps = t.paras();
    const { blocks } = await t.editor.copyBlocks(ps[0]!.el, ps[0]!.el, 'end', { pageBreak: true });
    expect(blocks.map((b) => b.name)).toEqual(['w:p', 'w:tbl']);
    expect(firstChild(firstChild(blocks[0]!, 'w:pPr'), 'w:pageBreakBefore')).toBeDefined();
    // A table ending the document gets a paragraph after it.
    expect(bodyBlocks(t.doc).map((b) => b.name)).toEqual(['w:tbl', 'w:p', 'w:p', 'w:tbl', 'w:p']);
  });

  it('records a copy as a tracked insertion that rejecting takes away again', async () => {
    const t = await open(p('Keep') + table([['x', 'y']]));
    const editor = new DocxEditor(t.doc, undefined, { trackChanges: true, author: 'Claude' });
    const ps = t.paras();
    const { blocks } = await editor.copyBlocks(ps[0]!.el, ps[2]!.el, 'end');
    expect(findAll(blocks[0]!, 'w:ins').length).toBeGreaterThan(0);
    expect(findAll(blocks[1]!, 'w:trPr').every((r) => !!firstChild(r, 'w:ins'))).toBe(true);
    editor.rejectRevisions();
    expect(t.texts()).toEqual(['Keep', 'x', 'y', '']);
  });
});

describe('making tables', () => {
  it('builds a table across the text width, with a bold header row that repeats', async () => {
    const t = await open(p('Intro') + p('End'));
    const tbl = t.editor.insertTable({ after: t.paras()[0]!.el }, [['No', 'Description', 'Total'], ['1', 'Land\nand trees', 'Rp. 201.801.600']], { header: true, widths: [1, 4, 2] });
    expect(t.texts()).toEqual(['Intro', 'No', 'Description', 'Total', '1', 'Land', 'and trees', 'Rp. 201.801.600', 'End']);
    const rows = elements(tbl, 'w:tr');
    expect(firstChild(firstChild(rows[0]!, 'w:trPr'), 'w:tblHeader')).toBeDefined();
    expect(findAll(rows[0]!, 'w:b')).toHaveLength(3);
    expect(findAll(rows[1]!, 'w:b')).toHaveLength(0);
    // Default page: 12240 twips wide with 1440 margins = 9360; widths in proportion 1:4:2.
    const cols = elements(firstChild(tbl, 'w:tblGrid')!, 'w:gridCol').map((g) => Number(g.attrs['w:w']));
    expect(cols.reduce((a, b) => a + b, 0)).toBe(9360);
    expect(cols[1]! / cols[0]!).toBeCloseTo(4, 1);
    // No Table Grid style here: the table draws its own grid.
    expect(findAll(tbl, 'w:insideV')).toHaveLength(1);
    const again = await reopen(t.doc);
    expect(again.body().blocks.map((b) => b.kind)).toEqual(['paragraph', 'table', 'paragraph']);
  });

  it('uses Table Grid when the document has it, pads short rows, and checks widths', async () => {
    const t = await open(p('Only'), TABLE_GRID);
    const tbl = t.editor.insertTable('end', [['a', 'b', 'c'], ['d']]);
    expect(firstChild(firstChild(tbl, 'w:tblPr'), 'w:tblStyle')?.attrs['w:val']).toBe('TableGrid');
    expect(findAll(tbl, 'w:tblBorders')).toHaveLength(0);
    expect(elements(elements(tbl, 'w:tr')[1]!, 'w:tc')).toHaveLength(3);
    // At the end of the document a paragraph follows the table.
    expect(bodyBlocks(t.doc).map((b) => b.name)).toEqual(['w:p', 'w:tbl', 'w:p']);
    expect(() => t.editor.insertTable('end', [['a', 'b']], { widths: [1] })).toThrow(/widths needs 2 positive numbers/);
    expect(() => t.editor.insertTable('end', [])).toThrow(/at least one row/);
    expect(() => t.editor.insertTable('end', [['a']], { style: 'Fancy' })).toThrow(/no table style "Fancy"/);
  });

  it('records a new table as a tracked insertion', async () => {
    const t = await open(p('Only'));
    const editor = new DocxEditor(t.doc, undefined, { trackChanges: true, author: 'Claude' });
    editor.insertTable('end', [['a', 'b']]);
    expect(editor.revisions().some((r) => r.kind === 'insert')).toBe(true);
    editor.rejectRevisions();
    expect(t.doc.body().blocks.filter((b) => b.kind === 'table')).toHaveLength(0);
  });
});

describe('editing table rows', () => {
  async function withTable() {
    const t = await open(p('Before') + table([['No', 'Item'], ['1', 'Land'], ['2', 'Trees']]) + p('After'));
    const tbl = t.doc.body().blocks.find((b) => b.kind === 'table')!.el;
    return { ...t, tbl };
  }
  const rowTexts = (tbl: XmlElement) => elements(tbl, 'w:tr').map((r) => elements(r, 'w:tc').map((c) => findAll(c, 'w:t').map((x) => x.children.map((n) => ('value' in n ? n.value : '')).join('')).join('|')));

  it('adds rows that copy the formatting of the row next to them', async () => {
    const t = await withTable();
    t.editor.insertTableRows(t.tbl, 3, [['3', 'Buildings'], ['4']]);
    expect(rowTexts(t.tbl)).toEqual([['No', 'Item'], ['1', 'Land'], ['2', 'Trees'], ['3', 'Buildings'], ['4', '']]);
    // The shading of the cells came along.
    expect(findAll(elements(t.tbl, 'w:tr')[3]!, 'w:shd')).toHaveLength(2);
    t.editor.insertTableRows(t.tbl, 1, [['0', 'First']]);
    expect(rowTexts(t.tbl)[1]).toEqual(['0', 'First']);
    expect(() => t.editor.insertTableRows(t.tbl, 1, [['a', 'b', 'c']])).toThrow(/at most 2 texts/);
    expect(() => t.editor.insertTableRows(t.tbl, 99, [['a']])).toThrow(/Rows go from 0 to 6/);
    await reopen(t.doc);
  });

  it('deletes rows, and the whole table when every row goes', async () => {
    const t = await withTable();
    t.editor.deleteTableRows(t.tbl, 1, 1);
    expect(rowTexts(t.tbl)).toEqual([['No', 'Item'], ['2', 'Trees']]);
    t.editor.deleteTableRows(t.tbl, 0, 1);
    expect(t.doc.body().blocks.map((b) => b.kind)).toEqual(['paragraph', 'paragraph']);
  });

  it('deletes a table, or marks its rows deleted when tracking', async () => {
    const t = await withTable();
    const tracked = new DocxEditor(t.doc, undefined, { trackChanges: true, author: 'Claude' });
    tracked.deleteTable(t.tbl);
    expect(elements(t.tbl, 'w:tr').every((r) => !!firstChild(firstChild(r, 'w:trPr'), 'w:del'))).toBe(true);
    tracked.acceptRevisions();
    expect(t.doc.body().blocks.map((b) => b.kind)).toEqual(['paragraph', 'paragraph']);
    const u = await withTable();
    u.editor.deleteTable(u.tbl);
    expect(u.texts()).toEqual(['Before', 'After']);
  });
});

describe('set_list rules for joining a list', () => {
  async function list() {
    const t = await open(p('Item') + p('Gap') + p('Next') + p('Plain'));
    const ps = t.paras();
    await t.editor.setList([ps[0]!.el], { format: 'decimal' });
    return { ...t, ps };
  }

  it("accepts sameListAs with the list's own format", async () => {
    const t = await list();
    await t.editor.setList([t.ps[2]!.el], { sameListAs: t.ps[0]!.el, format: 'decimal' });
    const items = t.editor.listItems();
    expect(items.get(t.ps[2]!.el)?.label).toBe('2.');
  });

  it('says what to do instead when the request cannot work', async () => {
    const t = await list();
    await expect(t.editor.setList([t.ps[2]!.el], { sameListAs: t.ps[0]!.el, format: 'bullet' })).rejects.toThrow(/That list is numbered 1\.; to continue it leave out format/);
    await expect(t.editor.setList([t.ps[2]!.el], { sameListAs: t.ps[0]!.el, start: 3 })).rejects.toThrow(/use start without sameListAs/);
    await expect(t.editor.setList([t.ps[2]!.el], { sameListAs: t.ps[3]!.el })).rejects.toThrow(/not in a list yet: make it one first/);
  });
});
