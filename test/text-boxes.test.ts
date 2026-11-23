/**
 * Text inside text boxes through the editor: a box's paragraphs come right after the paragraph
 * holding it, so find, replace, template filling and deletion reach them, and Word's stale VML
 * copy of a box goes when its content changes.
 */
import { describe, expect, it } from 'vitest';
import { DocxEditor, paragraphsOf, paragraphText } from '../src/api/index.js';
import { DocxDocument, type DocParagraph } from '../src/model/index.js';
import { elements, findAll, serializeXml } from '../src/xml/index.js';
import { buildDocx, p } from './helpers/docx.js';
import { EMU, inline, textBoxRun, textBoxShape, vmlTextBoxRun } from './helpers/drawings.js';

async function open(body: string) {
  const doc = await DocxDocument.open(await buildDocx({ body }));
  return { doc, editor: new DocxEditor(doc) };
}

const withBox = (text = 'Inside the box') => p('Before') + `<w:p>${textBoxRun(text)}</w:p>` + p('After');
const text = (para: DocParagraph) => paragraphText(para).text;
const reopen = async (doc: DocxDocument) => DocxDocument.open(await doc.save());
const texts = (doc: DocxDocument) => new DocxEditor(doc).paragraphs().map(text);

describe('text boxes in the editor', () => {
  it("lists a text box's paragraphs right after the paragraph holding it", async () => {
    const { doc, editor } = await open(withBox());
    // The holding paragraph has no text of its own here.
    expect(editor.paragraphs().map(text)).toEqual(['Before', '', 'Inside the box', 'After']);
    expect(paragraphsOf(doc.body().blocks)).toHaveLength(3);
  });

  it('lists the text boxes of a group depth first', async () => {
    const group =
      `<wpg:wgp><wpg:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${600 * EMU}" cy="${80 * EMU}"/><a:chOff x="0" y="0"/><a:chExt cx="${600 * EMU}" cy="${80 * EMU}"/></a:xfrm></wpg:grpSpPr>` +
      `${textBoxShape('One')}<wpg:grpSp><wpg:grpSpPr/>${textBoxShape('Two')}</wpg:grpSp>${textBoxShape('Three')}</wpg:wgp>`;
    const { editor } = await open(`<w:p><w:r>${inline(group, 600, 80, 'http://schemas.microsoft.com/office/word/2010/wordprocessingGroup')}</w:r></w:p>` + p('After'));
    expect(editor.paragraphs().map(text)).toEqual(['', 'One', 'Two', 'Three', 'After']);
  });

  it('finds and replaces inside text boxes', async () => {
    const { doc, editor } = await open(withBox());
    expect(editor.replaceAll('Inside', 'Within')).toBe(1);
    expect(texts(await reopen(doc))).toEqual(['Before', '', 'Within the box', 'After']);
  });

  it('drops the stale VML copy after an edit inside the box', async () => {
    const { doc, editor } = await open(withBox());
    editor.replaceAll('Inside', 'Within');
    const xml = serializeXml((await reopen(doc)).documentXml);
    expect(xml).toContain('mc:Choice');
    expect(xml).not.toContain('mc:Fallback');
    expect(xml).not.toContain('Old copy');

    // Any change inside the box, not only replacing text.
    const formatted = await open(withBox());
    formatted.editor.formatRange(formatted.editor.find('box')[0]!, { bold: true });
    expect(serializeXml(formatted.doc.documentXml)).not.toContain('mc:Fallback');

    const elsewhere = await open(withBox());
    elsewhere.editor.replaceAll('Before', 'First');
    const kept = serializeXml((await reopen(elsewhere.doc)).documentXml);
    expect(kept).toContain('mc:Fallback');
    expect(kept).toContain('Old copy');
  });

  it('fills template fields in text boxes', async () => {
    const { doc, editor } = await open(withBox('Dear {{name}}'));
    expect(editor.fillTemplate({ name: 'Ada' })).toEqual(['name']);
    expect(texts(await reopen(doc))).toEqual(['Before', '', 'Dear Ada', 'After']);
  });

  it('keeps a paragraph in an emptied text box', async () => {
    const { doc, editor } = await open(withBox());
    const box = findAll(doc.documentXml.root, 'w:txbxContent')[0]!;
    editor.deleteParagraph(elements(box, 'w:p')[0]!);
    const reopened = await reopen(doc);
    const [content] = findAll(reopened.documentXml.root, 'w:txbxContent');
    expect(elements(content!, 'w:p')).toHaveLength(1);
    expect(texts(reopened)).toEqual(['Before', '', '', 'After']);
  });

  it('covers old VML-only text boxes', async () => {
    const { doc, editor } = await open(p('Before') + `<w:p>${vmlTextBoxRun('VML text')}</w:p>`);
    expect(editor.paragraphs().map(text)).toEqual(['Before', '', 'VML text']);
    expect(editor.replaceAll('VML', 'Old-style')).toBe(1);
    expect(texts(await reopen(doc))).toEqual(['Before', '', 'Old-style text']);
  });
});
