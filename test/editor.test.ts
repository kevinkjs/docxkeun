import { describe, expect, it } from 'vitest';
import { DocxEditor } from '../src/api/index.js';
import { readComments } from '../src/comments/read.js';
import { DocxDocument, blocksToText, type DocParagraph } from '../src/model/index.js';
import { attr, findAll, parseXml, serializeXml } from '../src/xml/index.js';
import { readZip } from '../src/zip/zip.js';
import { buildDocx, p } from './helpers/docx.js';

async function open(body: string, extra: Parameters<typeof buildDocx>[0] extends infer T ? Partial<T> : never = {}) {
  const doc = await DocxDocument.open(await buildDocx({ body, ...extra }));
  return { doc, editor: new DocxEditor(doc) };
}

const text = (doc: DocxDocument) => blocksToText(doc.body().blocks);

/** Save, re-open with a fresh parser, and return the document. */
async function roundTrip(doc: DocxDocument): Promise<DocxDocument> {
  return DocxDocument.open(await doc.save());
}

describe('text editing', () => {
  it('finds text across runs and replaces it keeping the first run formatting', async () => {
    const { doc, editor } = await open('<w:p><w:r><w:rPr><w:b/></w:rPr><w:t>Hel</w:t></w:r><w:r><w:t>lo wor</w:t></w:r><w:r><w:rPr><w:i/></w:rPr><w:t>ld!</w:t></w:r></w:p>');
    const [m] = editor.find('lo world');
    expect(m!.text).toBe('lo world');
    expect(editor.textOf(m!)).toBe('lo world');
    editor.replaceRange(m!, 'p, Earth');
    expect(text(doc)).toBe('Help, Earth!');
    const reopened = await roundTrip(doc);
    expect(text(reopened)).toBe('Help, Earth!');
    const para = reopened.body().blocks[0] as DocParagraph;
    const bold = para.inlines.filter((i) => i.kind === 'text' && i.props.bold).map((i) => (i.kind === 'text' ? i.text : ''));
    expect(bold.join('')).toBe('Hel');
  });

  it('fills {{placeholders}} split across runs and content controls', async () => {
    const body =
      '<w:p><w:r><w:t>Dear {{</w:t></w:r><w:r><w:rPr><w:b/></w:rPr><w:t>name</w:t></w:r><w:r><w:t>}}, total {{ amount }} due {{unknown}}.</w:t></w:r></w:p>' +
      '<w:sdt><w:sdtPr><w:tag w:val="company"/><w:showingPlcHdr/></w:sdtPr><w:sdtContent><w:p><w:r><w:rPr><w:rStyle w:val="PlaceholderText"/></w:rPr><w:t>Click here</w:t></w:r></w:p></w:sdtContent></w:sdt>';
    const { doc, editor } = await open(body);
    const used = editor.fillTemplate({ name: 'Rina', amount: 'USD 1,200', company: 'PT Contoh' });
    expect(used.sort()).toEqual(['amount', 'company', 'name']);
    expect(text(doc)).toBe('Dear Rina, total USD 1,200 due {{unknown}}.\nPT Contoh');
    const xml = serializeXml(doc.documentXml);
    expect(xml).not.toContain('showingPlcHdr');
    expect(xml).not.toContain('PlaceholderText');
  });

  it('inserts, formats and deletes across paragraphs', async () => {
    const { doc, editor } = await open(p('First paragraph') + p('Second paragraph') + p('Third paragraph'));
    const [a] = editor.find('paragraph');
    editor.insertText(a!.start, 'big ', { bold: true });
    expect(text(doc).split('\n')[0]).toBe('First big paragraph');
    const [, second] = editor.find('Second');
    editor.formatRange(second ? second : editor.find('Second')[0]!, { italic: true, color: 'ff0000' });
    const range = { start: editor.find('big')[0]!.start, end: editor.find('Third')[0]!.end };
    editor.deleteRange(range);
    expect(text(doc)).toBe('First  paragraph');
    const reopened = await roundTrip(doc);
    expect(text(reopened)).toBe('First  paragraph');
  });

  it('inserts and styles paragraphs', async () => {
    const { doc, editor } = await open(p('Body'));
    const first = editor.paragraphs()[0]!.el;
    const heading = editor.insertParagraph({ before: first }, 'Heading', { style: 'Heading1' });
    editor.insertParagraph('end', 'Closing\tline\nnext');
    editor.setParagraphStyle(first, 'Quote');
    const reopened = await roundTrip(doc);
    const paras = reopened.body().blocks as DocParagraph[];
    const styleOf = (x: DocParagraph) => attr(findAll(x.el, 'w:pStyle')[0], 'w:val');
    expect(paras.map(styleOf)).toEqual(['Heading1', 'Quote', undefined]);
    expect(text(reopened)).toBe('Heading\nBody\nClosing\tline\nnext');
    expect(heading.name).toBe('w:p');
  });

  it('preserves unknown markup when editing', async () => {
    const body = '<w:p><w:r><w:t>Keep </w:t></w:r><mc:AlternateContent><mc:Choice Requires="w14"><w:r><w:t>new</w:t></w:r></mc:Choice><mc:Fallback><w:r><w:t>old</w:t></w:r></mc:Fallback></mc:AlternateContent><w:customXml w:element="x"><w:r><w:t> me</w:t></w:r></w:customXml></w:p>';
    const { doc, editor } = await open(body);
    editor.replaceRange(editor.find('Keep')[0]!, 'Kept');
    const xml = serializeXml(doc.documentXml);
    expect(xml).toContain('<mc:Fallback><w:r><w:t>old</w:t></w:r></mc:Fallback>');
    expect(xml).toContain('<w:customXml w:element="x">');
  });
});

describe('links and bookmarks', () => {
  it('adds external and internal hyperlinks and bookmarks', async () => {
    const { doc, editor } = await open(p('See the website and the appendix.') + p('Appendix'));
    await editor.addHyperlink(editor.find('the website')[0]!, { url: 'https://example.com/?a=1&b=2' }, 'Open site');
    editor.addBookmark(editor.find('Appendix')[0]!, 'appendix');
    await editor.addHyperlink(editor.find('appendix')[0]!, { bookmark: 'appendix' });
    await expect(editor.addHyperlink(editor.find('See')[0]!, { url: 'javascript:alert(1)' })).rejects.toThrow(/scheme/);
    expect(() => editor.addBookmark(editor.find('See')[0]!, '1bad name')).toThrow(/Bookmark names/);

    const reopened = await roundTrip(doc);
    const e2 = new DocxEditor(reopened);
    expect(e2.hyperlinks().map(({ url, anchor, tooltip, text }) => ({ url, anchor, tooltip, text }))).toEqual([
      { url: 'https://example.com/?a=1&b=2', anchor: undefined, tooltip: 'Open site', text: 'the website' },
      { url: undefined, anchor: 'appendix', tooltip: undefined, text: 'appendix' },
    ]);
    expect(e2.bookmarks()).toEqual([{ id: '0', name: 'appendix' }]);
    expect(reopened.styles.get('Hyperlink')).toBeDefined();
    const para = reopened.body().blocks[0] as DocParagraph;
    expect(para.inlines.map((i) => i.kind)).toContain('linkStart');

    e2.removeHyperlink(e2.hyperlinks()[0]!.el);
    e2.removeBookmark('appendix');
    expect(e2.hyperlinks()).toHaveLength(1);
    expect(e2.bookmarks()).toEqual([]);
  });
});

describe('comments', () => {
  it('adds, replies, resolves, edits and deletes comments across all comment parts', async () => {
    const { doc, editor } = await open(p('The royalty rate is 3.75 percent.'));
    const id = await editor.addComment(editor.find('3.75 percent')[0]!, { author: 'Rina Pratama', text: 'Confirm with finance', date: new Date('2026-10-28T09:00:00Z') });
    const reply = await editor.replyToComment(id, { author: 'Budi Santoso', text: 'Confirmed.' });
    await editor.resolveComment(id);

    const saved = await doc.save();
    const names = readZip(saved).map((e) => e.name);
    for (const part of ['word/comments.xml', 'word/commentsExtended.xml', 'word/commentsIds.xml', 'word/commentsExtensible.xml', 'word/people.xml']) expect(names).toContain(part);

    const reopened = await DocxDocument.open(saved);
    const comments = readComments(reopened);
    expect(comments).toHaveLength(2);
    expect(comments[0]).toMatchObject({ id, author: 'Rina Pratama', initials: 'RP', text: 'Confirm with finance', resolved: true });
    expect(comments[0]!.durableId).toMatch(/^[0-9A-F]{8}$/);
    expect(comments[0]!.dateUtc).toBe('2026-10-28T09:00:00Z');
    expect(comments[1]).toMatchObject({ id: reply, parentId: id, text: 'Confirmed.' });
    const para = reopened.body().blocks[0] as DocParagraph;
    expect(para.inlines.filter((i) => i.kind === 'commentStart').map((i) => (i.kind === 'commentStart' ? i.id : ''))).toEqual([id, reply]);
    const people = parseXml(new TextDecoder().decode(await reopened.pkg.getBytes('word/people.xml')));
    expect(findAll(people.root, 'w15:person').map((x) => attr(x, 'w15:author'))).toEqual(['Rina Pratama', 'Budi Santoso']);

    const e2 = new DocxEditor(reopened);
    await e2.editComment(id, 'Confirm with finance by Friday');
    expect(readComments(reopened)[0]!.text).toBe('Confirm with finance by Friday');
    await e2.deleteComment(id);
    expect(readComments(reopened)).toEqual([]);
    const xml = serializeXml(reopened.documentXml);
    expect(xml).not.toMatch(/commentRange|commentReference/);
    expect(text(reopened)).toBe('The royalty rate is 3.75 percent.');
  });
});
