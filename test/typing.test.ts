import { describe, expect, it } from 'vitest';
import { DocxEditor, paragraphText } from '../src/api/editor.js';
import { blocksToText, DocxDocument, type DocParagraph } from '../src/model/index.js';
import { attr, findAll, firstChild, serializeXml } from '../src/xml/index.js';
import { buildDocx, p } from './helpers/docx.js';

const STYLES = `<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>
<w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:next w:val="Normal"/><w:rPr><w:b/></w:rPr></w:style>`;

async function open(body: string) {
  const doc = await DocxDocument.open(await buildDocx({ body, styles: STYLES }));
  return { doc, editor: new DocxEditor(doc) };
}

const text = (doc: DocxDocument) => blocksToText(doc.body().blocks);

describe('editing primitives for typing', () => {
  it('types into existing text without creating runs', async () => {
    const { doc, editor } = await open('<w:p><w:r><w:t>Helo</w:t></w:r></w:p>');
    const t = findAll(doc.documentXml.root, 'w:t')[0]!;
    const pos = editor.insertText({ node: t, offset: 3 }, 'l');
    expect(pos.node).toBe(t);
    expect(pos.offset).toBe(4);
    const end = editor.insertText({ node: t, offset: 5 }, ' world');
    expect(end.offset).toBe(11);
    expect(text(doc)).toBe('Hello world');
    expect(findAll(doc.documentXml.root, 'w:r')).toHaveLength(1);
    expect(attr(t, 'xml:space')).toBeUndefined();
    editor.insertText({ node: t, offset: 11 }, ' ');
    expect(attr(t, 'xml:space')).toBe('preserve');
  });

  it('types into an empty paragraph using its paragraph-mark formatting', async () => {
    const { doc, editor } = await open('<w:p><w:pPr><w:rPr><w:b/></w:rPr></w:pPr></w:p>');
    const para = findAll(doc.documentXml.root, 'w:p')[0]!;
    const pos = editor.insertText({ node: para, offset: 0 }, 'Bold start');
    expect(pos.node.name).toBe('w:t');
    expect(text(doc)).toBe('Bold start');
    const inline = (doc.body().blocks[0] as DocParagraph).inlines[0]!;
    expect(inline.kind === 'text' && inline.props.bold).toBe(true);
  });

  it('deletes inside a text node in place', async () => {
    const { doc, editor } = await open(p('abcdef'));
    const t = findAll(doc.documentXml.root, 'w:t')[0]!;
    editor.deleteRange({ start: { node: t, offset: 2 }, end: { node: t, offset: 4 } });
    expect(text(doc)).toBe('abef');
  });

  it('splits paragraphs, moving the section break and following the heading "next" style', async () => {
    const body = p('Title', '<w:pStyle w:val="Heading1"/>') + p('First second', '<w:sectPr><w:type w:val="continuous"/></w:sectPr>');
    const { doc, editor } = await open(body);
    const [title, para] = findAll(doc.documentXml.root, 'w:t');
    const afterTitle = editor.splitParagraph({ node: title!, offset: 5 });
    expect(attr(firstChild(firstChild(afterTitle, 'w:pPr'), 'w:pStyle'), 'w:val')).toBe('Normal');
    const second = editor.splitParagraph({ node: para!, offset: 6 });
    expect(text(doc)).toBe('Title\n\nFirst \nsecond');
    const paras = doc.body().blocks as DocParagraph[];
    expect(paras[2]!.sectionEnd).toBeUndefined();
    expect(paras[3]!.sectionEnd?.type).toBe('continuous');
    expect(second.name).toBe('w:p');
    // Splitting in the middle of a heading keeps the heading style on both halves.
    const { doc: d2, editor: e2 } = await open(p('Long title', '<w:pStyle w:val="Heading1"/>'));
    const half = e2.splitParagraph({ node: findAll(d2.documentXml.root, 'w:t')[0]!, offset: 4 });
    expect(attr(firstChild(firstChild(half, 'w:pPr'), 'w:pStyle'), 'w:val')).toBe('Heading1');
  });

  it('merges paragraphs (backspace at the start of a paragraph)', async () => {
    const { doc, editor } = await open(p('One', '<w:jc w:val="center"/>') + p('Two', '<w:sectPr/>'));
    const [a, b] = findAll(doc.documentXml.root, 'w:p');
    editor.mergeParagraphs(a!, b!);
    expect(text(doc)).toBe('OneTwo');
    const pPr = firstChild(findAll(doc.documentXml.root, 'w:p')[0], 'w:pPr')!;
    expect(firstChild(pPr, 'w:jc')).toBeDefined();
    expect(firstChild(pPr, 'w:sectPr')).toBeDefined();
    const reopened = await DocxDocument.open(await doc.save());
    expect(text(reopened)).toBe('OneTwo');
  });

  it('refuses to merge paragraphs in different table cells', async () => {
    const { doc, editor } = await open(`<w:tbl><w:tr><w:tc>${p('a')}</w:tc><w:tc>${p('b')}</w:tc></w:tr></w:tbl><w:p/>`);
    const [a, b] = findAll(doc.documentXml.root, 'w:p');
    expect(() => editor.mergeParagraphs(a!, b!)).toThrow(/table cells/);
    expect(serializeXml(doc.documentXml)).toContain('<w:t xml:space="preserve">a</w:t>');
  });

  it('maps tabs and line breaks to positions', async () => {
    const { doc } = await open('<w:p><w:r><w:t>a</w:t><w:tab/><w:t>b</w:t><w:br/><w:t>c</w:t></w:r></w:p>');
    const pt = paragraphText(doc.body().blocks[0] as DocParagraph);
    expect(pt.text).toBe('a\tb\nc');
    expect(pt.segments.map((s) => s.node.name)).toEqual(['w:t', 'w:tab', 'w:t', 'w:br', 'w:t']);
  });
});
