/**
 * Text inside text boxes through the editor: a box's paragraphs come right after the paragraph
 * holding it, so find, replace, template filling and deletion reach them, and Word's stale VML
 * copy of a box goes when its content changes.
 */
import { describe, expect, it } from 'vitest';
import { DocxEditor, paragraphsOf, paragraphText } from '../src/api/index.js';
import { DocxDocument, type DocParagraph } from '../src/model/index.js';
import { attr, elements, findAll, serializeXml, type XmlElement } from '../src/xml/index.js';
import { buildDocx, p } from './helpers/docx.js';
import { EMU, inline, textBoxRun, textBoxRunHolding, textBoxShape, vmlTextBoxRun } from './helpers/drawings.js';

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

  it('drops the stale VML copy when typing or pressing Enter at a locked edge inside the box', async () => {
    const lockedText = '<w:p><w:r><w:t xml:space="preserve">Name: </w:t></w:r><w:sdt><w:sdtPr><w:lock w:val="contentLocked"/></w:sdtPr><w:sdtContent><w:r><w:t>Ada</w:t></w:r></w:sdtContent></w:sdt></w:p>';
    const typed = await open(`<w:p>${textBoxRunHolding(lockedText)}</w:p>`);
    // Typing at the end of locked text goes just outside it.
    typed.editor.insertText({ node: findAll(typed.doc.documentXml.root, 'w:sdt').flatMap((s) => findAll(s, 'w:t'))[0]!, offset: 3 }, '!');
    const afterTyping = await reopen(typed.doc);
    expect(serializeXml(afterTyping.documentXml)).not.toContain('mc:Fallback');
    expect(texts(afterTyping)).toEqual(['', 'Name: Ada!']);

    const lockedBlock = '<w:sdt><w:sdtPr><w:lock w:val="sdtContentLocked"/></w:sdtPr><w:sdtContent><w:p><w:r><w:t>Fixed</w:t></w:r></w:p></w:sdtContent></w:sdt>';
    const split = await open(`<w:p>${textBoxRunHolding(lockedBlock)}</w:p>`);
    // Enter at the end of a locked block adds a paragraph after it.
    split.editor.splitParagraph({ node: findAll(split.doc.documentXml.root, 'w:sdt').flatMap((s) => findAll(s, 'w:t'))[0]!, offset: 5 });
    const afterSplit = await reopen(split.doc);
    expect(serializeXml(afterSplit.documentXml)).not.toContain('mc:Fallback');
    expect(texts(afterSplit)).toEqual(['', 'Fixed', '']);
  });

  it('drops the stale VML copy when a link, picture, bookmark or lock inside the box changes', async () => {
    const picture =
      `<w:r><w:drawing><wp:inline><wp:extent cx="${20 * EMU}" cy="${20 * EMU}"/><wp:docPr id="9" name="Picture 9"/>` +
      '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:blipFill><a:blip r:embed="rIdNone"/></pic:blipFill></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>';
    const content =
      '<w:p><w:hyperlink w:anchor="top"><w:r><w:t>Link</w:t></w:r></w:hyperlink><w:bookmarkStart w:id="7" w:name="mark"/><w:r><w:t>Text</w:t></w:r><w:bookmarkEnd w:id="7"/>' +
      `<w:sdt><w:sdtPr><w:lock w:val="contentLocked"/></w:sdtPr><w:sdtContent><w:r><w:t>Locked</w:t></w:r></w:sdtContent></w:sdt>${picture}</w:p>`;
    const link = (root: XmlElement) => findAll(root, 'w:hyperlink')[0]!;
    const changes: Record<string, (editor: DocxEditor, root: XmlElement) => unknown> = {
      'link target': (editor, root) => editor.setHyperlinkTarget(link(root), { bookmark: 'mark' }),
      'link removed': (editor, root) => editor.removeHyperlink(link(root)),
      'alt text': (editor, root) => editor.setImageAlt(findAll(root, 'wp:docPr').find((d) => attr(d, 'name') === 'Picture 9')!, 'Logo'),
      'bookmark removed': (editor) => editor.removeBookmark('mark'),
      unlocked: (editor) => editor.unlock(editor.lockedRegions()[0]!.id),
    };
    const stale: string[] = [];
    for (const [name, change] of Object.entries(changes)) {
      const { doc, editor } = await open(`<w:p>${textBoxRunHolding(content)}</w:p>`);
      await change(editor, doc.documentXml.root);
      if (serializeXml(doc.documentXml).includes('mc:Fallback')) stale.push(name);
    }
    expect(stale).toEqual([]);
  });

  it('keeps a paragraph where a deleted table was all a text box or cell held', async () => {
    const table = '<w:tbl><w:tr><w:tc><w:p><w:r><w:t>Cell</w:t></w:r></w:p></w:tc></w:tr></w:tbl>';
    const boxed = await open(`<w:p>${textBoxRunHolding(table)}</w:p>`);
    boxed.editor.deleteTable(findAll(boxed.doc.documentXml.root, 'w:tbl')[0]!);
    const [content] = findAll((await reopen(boxed.doc)).documentXml.root, 'w:txbxContent');
    expect(elements(content!).map((c) => c.name)).toEqual(['w:p']);

    // A table cell, as before.
    const nested = await open(`<w:tbl><w:tr><w:tc>${table}</w:tc></w:tr></w:tbl>` + p('After'));
    const [outer, inner] = findAll(nested.doc.documentXml.root, 'w:tbl');
    nested.editor.deleteTable(inner!);
    expect(elements(findAll(outer!, 'w:tc')[0]!).map((c) => c.name)).toEqual(['w:p']);
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
