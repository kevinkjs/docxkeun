import { describe, expect, it } from 'vitest';
import { layoutDocument } from '../src/layout/index.js';
import { DocxDocument } from '../src/model/index.js';
import { collectNavTargets, nextTarget } from '../src/render/navigation.js';
import { buildDocx, p } from './helpers/docx.js';
import { testRegistry } from './helpers/fonts.js';

const STYLES = `<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="22"/></w:rPr></w:rPrDefault></w:docDefaults>
<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>
<w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:pPr><w:outlineLvl w:val="0"/></w:pPr><w:rPr><w:sz w:val="32"/></w:rPr></w:style>`;

const link = (text: string, attrs: string) => `<w:hyperlink ${attrs}><w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:hyperlink>`;
const filler = (n: number) => Array.from({ length: n }, (_, i) => p(`Filler line ${i + 1}`)).join('');

async function targets() {
  const header = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:p><w:hyperlink r:id="rIdH"><w:r><w:t>Header link</w:t></w:r></w:hyperlink></w:p></w:hdr>`;
  const body =
    '<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:bookmarkStart w:id="0" w:name="top"/><w:r><w:t>Introduction</w:t></w:r><w:bookmarkEnd w:id="0"/></w:p>' +
    `<w:p>${link('Go to terms', 'w:anchor="terms"')}<w:r><w:t xml:space="preserve"> or read </w:t></w:r>${link('the website', 'r:id="rIdW"')}</w:p>` +
    `<w:p>${link('A long link whose text runs on and on, and then on some more, until it certainly wraps onto the next line of the page', 'r:id="rIdW"')}</w:p>` +
    '<w:p><w:bookmarkStart w:id="1" w:name="_Toc12345"/><w:r><w:t>hidden bookmark here</w:t></w:r><w:bookmarkEnd w:id="1"/></w:p>' +
    filler(60) +
    '<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:bookmarkStart w:id="2" w:name="terms"/><w:r><w:t>Terms</w:t></w:r><w:bookmarkEnd w:id="2"/></w:p>' +
    `<w:p>${link('Back to top', 'w:anchor="top"')}</w:p>`;
  const sect =
    '<w:sectPr><w:headerReference w:type="default" r:id="rIdHdr"/><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr>';
  const doc = await DocxDocument.open(
    await buildDocx({
      styles: STYLES,
      body,
      sectPr: sect,
      rels: [
        ['rIdW', 'hyperlink', 'https://example.com/', true],
        ['rIdH', 'hyperlink', 'https://example.com/header', true],
        ['rIdHdr', 'header', 'header1.xml'],
      ],
      parts: { 'word/header1.xml': { data: header, contentType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml' } },
    }),
  );
  return collectNavTargets(layoutDocument(doc, { fonts: testRegistry() }));
}

describe('navigation targets', () => {
  it('lists headings, bookmarks and body links in document order', async () => {
    const t = await targets();
    expect(t.headings.map((h) => [h.label, h.page, h.level])).toEqual([
      ['Introduction', 0, 0],
      ['Terms', 1, 0],
    ]);
    expect(t.bookmarks.map((b) => [b.name, b.page, b.hidden])).toEqual([
      ['top', 0, false],
      ['_Toc12345', 0, true],
      ['terms', 1, false],
    ]);
    expect(t.links.map((l) => l.label)).toEqual([
      'Go to terms',
      'the website',
      'A long link whose text runs on and on, and then on some more, until it certainly wraps onto the next line of the page',
      'Back to top',
    ]);
    // The header's link is left out even though it appears on every page.
    expect(t.links.some((l) => l.link?.url === 'https://example.com/header')).toBe(false);
  });

  it('keeps a wrapped link as one target with a rectangle per line', async () => {
    const t = await targets();
    const wrapped = t.links[2]!;
    expect(wrapped.rects.length).toBe(2);
    expect(wrapped.rects[1]!.y).toBeGreaterThan(wrapped.rects[0]!.y);
  });

  it('resolves internal links to their bookmark', async () => {
    const t = await targets();
    const toTerms = t.links[0]!;
    expect(toTerms.link).toMatchObject({ anchor: 'terms' });
    expect(toTerms.destination).toMatchObject({ page: 1 });
    expect(toTerms.destination!.y).toBeCloseTo(t.bookmarks.find((b) => b.name === 'terms')!.y, 5);
    expect(t.links[3]!.destination).toMatchObject({ page: 0, y: 72 });
    expect(t.links[1]!.destination).toBeUndefined();
  });

  it('finds the next and previous target from a location', async () => {
    const t = await targets();
    const top = { page: 0, x: 0, y: 0 };
    expect(nextTarget(t.links, top, 1)?.label).toBe('Go to terms');
    expect(nextTarget(t.links, { page: 0, x: 0, y: 700 }, 1)?.label).toBe('Back to top');
    expect(nextTarget(t.links, { page: 1, x: 0, y: 700 }, 1)).toBeUndefined();
    expect(nextTarget(t.links, { page: 1, x: 0, y: 0 }, -1)?.label).toMatch(/^A long link/);
    expect(nextTarget(t.headings, { page: 1, x: 0, y: 0 }, 1)?.label).toBe('Terms');
  });
});
