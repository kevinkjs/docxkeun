import { describe, expect, it } from 'vitest';
import { DocxEditor } from '../src/api/index.js';
import { layoutDocument, type TextItem } from '../src/layout/index.js';
import { DocxDocument } from '../src/model/index.js';
import { serializeNode } from '../src/xml/index.js';
import { buildDocx, p } from './helpers/docx.js';
import { testRegistry } from './helpers/fonts.js';

const PAGE = '<w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>';
const fonts = testRegistry();

async function open(body = p('Body text.'), settings = '') {
  const doc = await DocxDocument.open(await buildDocx({ body, settings }));
  return { doc, editor: new DocxEditor(doc) };
}

const pageTexts = (doc: DocxDocument, page = 0) =>
  layoutDocument(doc, { fonts })
    .pages[page]!.items.filter((i): i is TextItem => i.type === 'text')
    .map((t) => ({ text: t.text, x: t.x, y: t.y }));

const many = (n: number) => Array.from({ length: n }, (_, i) => p(`Paragraph ${i + 1} with enough words to take some room on the page.`)).join('');

describe('adding headers and footers', () => {
  it('adds a header with left, centre and right parts and a page-number footer', async () => {
    const { doc, editor } = await open(many(60));
    await editor.setHeader('ACME Ltd.\tService Agreement\tConfidential');
    await editor.setFooter('Page {PAGE} of {NUMPAGES}', { align: 'center' });
    const again = await DocxDocument.open(await doc.save());
    const info = new DocxEditor(again).headersAndFooters();
    expect(info.map((h) => [h.kind, h.type, h.sections, h.text])).toEqual([
      ['header', 'default', [1], 'ACME Ltd.\tService Agreement\tConfidential'],
      ['footer', 'default', [1], 'Page {PAGE} of {NUMPAGES}'],
    ]);
    const pages = layoutDocument(again, { fonts }).pages;
    expect(pages.length).toBeGreaterThan(1);
    const second = pages[1]!.items.filter((i): i is TextItem => i.type === 'text');
    // Page numbers are filled in per page; the header's parts sit left, centre and right.
    expect(second.map((t) => t.text).join('')).toContain(`Page 2 of ${pages.length}`);
    const header = second.filter((t) => t.y < 72);
    const left = header.find((t) => t.text.startsWith('ACME'))!;
    const right = header.find((t) => t.text.includes('Confidential'))!;
    expect(left.x).toBeCloseTo(72, 0);
    expect(right.xs[right.xs.length - 1]! + right.widths[right.widths.length - 1]!).toBeCloseTo(612 - 72, 0);
  });

  it('replaces a header in place, and gives a first page its own footer', async () => {
    const { doc, editor } = await open(many(60));
    const part = await editor.setHeader('Draft');
    expect(await editor.setHeader('Final')).toBe(part);
    expect(editor.headersAndFooters()).toHaveLength(1);
    await editor.setFooter('Cover page', { type: 'first' });
    await editor.setFooter('Page {PAGE}');
    const body = serializeNode(doc.documentXml.root);
    expect(body).toContain('<w:titlePg/>');
    // References come first in the section, headers before footers.
    expect(body).toMatch(/<w:sectPr><w:headerReference w:type="default" r:id="[^"]+"\/><w:footerReference w:type="first" r:id="[^"]+"\/><w:footerReference w:type="default" r:id="[^"]+"\/><w:pgSz/);
    expect(pageTexts(doc, 0).some((t) => t.text === 'Cover page')).toBe(true);
    expect(pageTexts(doc, 1).some((t) => t.text.includes('Page 2') || t.text === '2')).toBe(true);
  });

  it('sets even-page headers (and Word\'s even/odd setting) and removes them again', async () => {
    const { doc, editor } = await open(many(60), '<w:defaultTabStop w:val="720"/>');
    await editor.setHeader('Odd pages');
    await editor.setHeader('Even pages', { type: 'even' });
    expect(serializeNode(doc.xmlSync(doc.parts.settings!)!.root)).toContain('<w:evenAndOddHeaders/>');
    expect(pageTexts(doc, 1).some((t) => t.text === 'Even pages')).toBe(true);
    expect(await editor.removeHeader({ type: 'even' })).toBe(1);
    expect(serializeNode(doc.xmlSync(doc.parts.settings!)!.root)).not.toContain('evenAndOddHeaders');
    expect(pageTexts(doc, 1).some((t) => t.text === 'Odd pages')).toBe(true);
  });

  it('works per section, and removing deletes parts nothing uses', async () => {
    const sect = `<w:sectPr>${PAGE}</w:sectPr>`;
    const body = p('Section one') + `<w:p><w:pPr>${sect}</w:pPr><w:r><w:t>End of one</w:t></w:r></w:p>` + p('Section two');
    const { doc, editor } = await open(body);
    await editor.setHeader('Everywhere');
    await editor.setHeader('Only section two', { section: 2 });
    const info = editor.headersAndFooters();
    expect(info.map((h) => [h.text, h.sections])).toEqual([
      ['Everywhere', [1]],
      ['Only section two', [2]],
    ]);
    await expect(editor.setHeader('x', { section: 3 })).rejects.toThrow('There is no section 3');
    expect(await editor.removeHeader()).toBe(2);
    expect(editor.headersAndFooters()).toEqual([]);
    expect(doc.pkg.partNames().filter((n) => /header/.test(n))).toEqual([]);
    const saved = await DocxDocument.open(await doc.save());
    expect(new DocxEditor(saved).headersAndFooters()).toEqual([]);
  });

  it('records changed header text as tracked changes', async () => {
    const { doc, editor } = await open();
    await editor.setHeader('Draft');
    const tracked = new DocxEditor(doc, undefined, { trackChanges: true, author: 'Rina' });
    await tracked.setHeader('Final');
    const xml = serializeNode(doc.xmlSync(editor.headersAndFooters()[0]!.part)!.root);
    expect(xml).toMatch(/<w:del w:id="\d+" w:author="Rina"[^>]*><w:r><w:delText>Draft<\/w:delText><\/w:r><\/w:del>/);
    expect(xml).toMatch(/<w:ins w:id="\d+" w:author="Rina"[^>]*><w:r><w:t>Final<\/w:t><\/w:r><\/w:ins>/);
    expect(new DocxEditor(doc, undefined, { part: editor.headersAndFooters()[0]!.part }).revisions().map((r) => r.kind)).toContain('insert');
  });
});
