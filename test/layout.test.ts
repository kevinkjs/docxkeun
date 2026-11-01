import { describe, expect, it } from 'vitest';
import { layoutDocument, type Page, type TextItem } from '../src/layout/index.js';
import { DocxDocument } from '../src/model/index.js';
import { buildDocx, p } from './helpers/docx.js';
import { haveSystemFonts, testRegistry } from './helpers/fonts.js';

const STYLES = `
<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="22"/></w:rPr></w:rPrDefault>
<w:pPrDefault><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults>
<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>
<w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:basedOn w:val="Normal"/><w:pPr><w:keepNext/><w:spacing w:before="240"/></w:pPr><w:rPr><w:b/><w:sz w:val="32"/></w:rPr></w:style>
<w:style w:type="character" w:styleId="Hyperlink"><w:name w:val="Hyperlink"/><w:rPr><w:color w:val="0563C1"/><w:u w:val="single"/></w:rPr></w:style>
<w:style w:type="table" w:styleId="TableGrid"><w:name w:val="Table Grid"/><w:tblPr><w:tblBorders><w:top w:val="single" w:sz="4"/><w:left w:val="single" w:sz="4"/><w:bottom w:val="single" w:sz="4"/><w:right w:val="single" w:sz="4"/><w:insideH w:val="single" w:sz="4"/><w:insideV w:val="single" w:sz="4"/></w:tblBorders></w:tblPr></w:style>`;

const LOREM = 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut aliquip ex ea commodo consequat.';

async function layout(body: string, extra: Parameters<typeof buildDocx>[0] extends infer T ? Partial<T> : never = {}) {
  const doc = await DocxDocument.open(await buildDocx({ styles: STYLES, body, ...extra }));
  return { doc, result: layoutDocument(doc, { fonts: testRegistry() }) };
}

function texts(page: Page): TextItem[] {
  return page.items.filter((i): i is TextItem => i.type === 'text');
}

function lineYs(page: Page): number[] {
  return [...new Set(texts(page).map((t) => Math.round(t.y * 100) / 100))].sort((a, b) => a - b);
}

describe.skipIf(!haveSystemFonts)('layout', () => {
  it('places the first baseline at the top margin plus the ascent', async () => {
    const { result } = await layout(p('Hello world'));
    expect(result.pages).toHaveLength(1);
    const [t] = texts(result.pages[0]!);
    expect(t!.text).toBe('Hello world');
    expect(t!.x).toBeCloseTo(72, 5);
    // Arial winAscent 1854/2048 at 11pt.
    expect(t!.y).toBeCloseTo(72 + (1854 / 2048) * 11, 3);
    expect(t!.glyphs.every((g) => g > 0)).toBe(true);
  });

  it('wraps long paragraphs within the margins and justifies', async () => {
    const { result } = await layout(p(LOREM.repeat(3), '<w:jc w:val="both"/>'));
    const page = result.pages[0]!;
    const ys = lineYs(page);
    expect(ys.length).toBeGreaterThan(5);
    // Line pitch = (winAscent + winDescent) / upem * 11pt.
    expect(ys[1]! - ys[0]!).toBeCloseTo(((1854 + 434) / 2048) * 11, 1);
    const right = 612 - 72;
    for (const y of ys.slice(0, -1)) {
      const items = texts(page).filter((t) => Math.abs(t.y - y) < 0.01);
      let end = 0;
      for (const t of items) {
        const visible = t.chars.map((c, i) => (c === ' ' ? 0 : t.xs[i]! + t.widths[i]!));
        end = Math.max(end, ...visible);
      }
      expect(end).toBeCloseTo(right, 1);
    }
  });

  it('paginates long documents and resolves page-number fields in footers', async () => {
    const footer = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:ftr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:p><w:r><w:t xml:space="preserve">Page </w:t></w:r><w:fldSimple w:instr="PAGE"><w:r><w:t>1</w:t></w:r></w:fldSimple><w:r><w:t xml:space="preserve"> of </w:t></w:r><w:fldSimple w:instr="NUMPAGES"><w:r><w:t>1</w:t></w:r></w:fldSimple></w:p></w:ftr>`;
    const body = Array.from({ length: 60 }, (_, i) => p(`${i + 1}. ${LOREM}`)).join('');
    const sect = '<w:sectPr><w:footerReference w:type="default" r:id="rIdF"/><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr>';
    const { result } = await layout(body, {
      sectPr: sect,
      rels: [['rIdF', 'footer', 'footer1.xml']],
      parts: { 'word/footer1.xml': { data: footer, contentType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml' } },
    });
    const n = result.pages.length;
    expect(n).toBeGreaterThan(3);
    result.pages.forEach((page, i) => {
      const footerText = texts(page).filter((t) => t.y > 720).map((t) => t.text).join('');
      expect(footerText).toBe(`Page ${i + 1} of ${n}`);
      // Body text stays above the bottom margin.
      for (const t of texts(page).filter((t) => t.y < 700)) expect(t.y).toBeLessThanOrEqual(792 - 72);
    });
  });

  it('keeps a heading with the next paragraph', async () => {
    // Fill the page so the heading lands at the very bottom.
    const filler = Array.from({ length: 52 }, () => p('filler')).join('');
    const { result } = await layout(filler + p('Heading', '<w:pStyle w:val="Heading1"/>') + p('Body after heading'));
    expect(result.pages.length).toBe(2);
    const second = texts(result.pages[1]!).map((t) => t.text);
    expect(second).toEqual(['Heading', 'Body after heading']);
  });

  it('honours page breaks and section page sizes', async () => {
    const body = p('One') + '<w:p><w:r><w:br w:type="page"/></w:r></w:p>' + p('Two', '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>') + p('Landscape');
    const sect = '<w:sectPr><w:pgSz w:w="15840" w:h="12240" w:orient="landscape"/><w:pgMar w:top="720" w:right="720" w:bottom="720" w:left="720" w:header="360" w:footer="360" w:gutter="0"/></w:sectPr>';
    const { result } = await layout(body, { sectPr: sect });
    expect(result.pages.map((pg) => [pg.width, pg.height])).toEqual([
      [612, 792],
      [612, 792],
      [792, 612],
    ]);
    expect(texts(result.pages[2]!)[0]!.x).toBeCloseTo(36, 3);
  });

  it('records link areas, bookmark anchors, comment areas and permission ranges', async () => {
    const body =
      '<w:p><w:bookmarkStart w:id="0" w:name="target"/><w:r><w:t>Anchor</w:t></w:r><w:bookmarkEnd w:id="0"/></w:p>' +
      '<w:p><w:hyperlink r:id="rIdL"><w:r><w:rPr><w:rStyle w:val="Hyperlink"/></w:rPr><w:t>External</w:t></w:r></w:hyperlink><w:r><w:t xml:space="preserve"> and </w:t></w:r><w:hyperlink w:anchor="target"><w:r><w:t>internal</w:t></w:r></w:hyperlink></w:p>' +
      '<w:p><w:commentRangeStart w:id="1"/><w:r><w:t>commented</w:t></w:r><w:commentRangeEnd w:id="1"/><w:permStart w:id="5" w:edGrp="everyone"/><w:r><w:t xml:space="preserve"> editable</w:t></w:r><w:permEnd w:id="5"/></w:p>';
    const { result } = await layout(body, { rels: [['rIdL', 'hyperlink', 'https://example.com', true]] });
    const page = result.pages[0]!;
    expect(page.links.map((l) => l.target)).toEqual([{ url: 'https://example.com', anchor: undefined, tooltip: undefined }, { url: undefined, anchor: 'target', tooltip: undefined }]);
    expect(page.links[0]!.w).toBeGreaterThan(30);
    expect(result.anchors.get('target')).toMatchObject({ page: 0, y: 72 });
    const comment = page.comments.filter((c) => c.w > 0);
    expect(comment).toHaveLength(1);
    expect(result.commentAnchors.get('1')?.page).toBe(0);
    expect(page.perms).toHaveLength(1);
    expect(page.perms[0]).toMatchObject({ id: '5', group: 'everyone' });
    // The underline of the hyperlink is drawn.
    expect(page.items.some((i) => i.type === 'line' && i.color === '0563C1')).toBe(true);
  });

  it('lays out tables with borders, spans and repeated header rows', async () => {
    const row = (a: string, b: string, header = false) => `<w:tr>${header ? '<w:trPr><w:tblHeader/></w:trPr>' : ''}<w:tc>${p(a)}</w:tc><w:tc>${p(b)}</w:tc></w:tr>`;
    const rows = Array.from({ length: 70 }, (_, i) => row(`r${i}`, `value ${i}`)).join('');
    const tbl = `<w:tbl><w:tblPr><w:tblStyle w:val="TableGrid"/></w:tblPr><w:tblGrid><w:gridCol w:w="4680"/><w:gridCol w:w="4680"/></w:tblGrid>${row('Name', 'Value', true)}${rows}</w:tbl><w:p/>`;
    const { result } = await layout(tbl);
    expect(result.pages.length).toBe(2);
    const first = texts(result.pages[0]!).map((t) => t.text);
    const second = texts(result.pages[1]!).map((t) => t.text);
    expect(first.slice(0, 2)).toEqual(['Name', 'Value']);
    expect(second.slice(0, 2)).toEqual(['Name', 'Value']);
    const cellText = texts(result.pages[0]!).find((t) => t.text === 'r0')!;
    expect(cellText.x).toBeCloseTo(72 + 5.4, 3);
    const lines = result.pages[0]!.items.filter((i) => i.type === 'line');
    expect(lines.length).toBeGreaterThan(100);
  });

  it('numbers list items and aligns text at the hanging indent', async () => {
    const numbering = `<w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/><w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr></w:lvl></w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>`;
    const item = (t: string) => p(t, '<w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr>');
    const { result } = await layout(item('First') + item('Second'), { numbering });
    const t = texts(result.pages[0]!);
    expect(t.map((x) => x.text)).toEqual(['1.', 'First', '2.', 'Second']);
    expect(t[0]!.x).toBeCloseTo(72 + 18, 3);
    expect(t[1]!.x).toBeCloseTo(72 + 36, 3);
  });

  it('applies tab stops with leaders', async () => {
    const body = p('Chapter\tText', '<w:tabs><w:tab w:val="right" w:leader="dot" w:pos="9360"/></w:tabs>').replace('Chapter\tText', 'Chapter</w:t></w:r><w:r><w:tab/></w:r><w:r><w:t>12');
    const { result } = await layout(body);
    const t = texts(result.pages[0]!);
    const num = t.find((x) => x.text === '12')!;
    expect(num.xs[1]! + num.widths[1]!).toBeCloseTo(72 + 468, 3);
    expect(t.some((x) => /^\.+$/.test(x.text))).toBe(true);
  });
});
