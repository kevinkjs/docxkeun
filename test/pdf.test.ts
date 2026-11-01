import { deflateSync } from 'node:zlib';
import { describe, expect, it } from 'vitest';
import { DocxEditor } from '../src/api/index.js';
import { DocxDocument } from '../src/model/index.js';
import { exportPdf } from '../src/pdf/index.js';
import { crc32 } from '../src/util/crc32.js';
import { buildDocx, p } from './helpers/docx.js';
import { testRegistry } from './helpers/fonts.js';
import { dictOf, pages, pageText, parsePdf, resolve, streamData, type PdfDictObj, type PdfNameObj, type PdfObj, type PdfStr } from './helpers/pdf.js';

const STYLES = `<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="22"/></w:rPr></w:rPrDefault></w:docDefaults>
<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>
<w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:pPr><w:outlineLvl w:val="0"/></w:pPr><w:rPr><w:b/><w:sz w:val="32"/></w:rPr></w:style>
<w:style w:type="paragraph" w:styleId="Heading2"><w:name w:val="heading 2"/><w:pPr><w:outlineLvl w:val="1"/></w:pPr><w:rPr><w:b/><w:sz w:val="26"/></w:rPr></w:style>`;

/** A tiny RGBA PNG (2x2, one transparent pixel). */
function rgbaPng(): Uint8Array {
  const chunk = (type: string, data: Uint8Array) => {
    const out = new Uint8Array(12 + data.length);
    const v = new DataView(out.buffer);
    v.setUint32(0, data.length);
    for (let i = 0; i < 4; i++) out[4 + i] = type.charCodeAt(i);
    out.set(data, 8);
    v.setUint32(8 + data.length, crc32(out.subarray(4, 8 + data.length)));
    return out;
  };
  const ihdr = new Uint8Array(13);
  new DataView(ihdr.buffer).setUint32(0, 2);
  new DataView(ihdr.buffer).setUint32(4, 2);
  ihdr.set([8, 6, 0, 0, 0], 8);
  const raw = Uint8Array.from([0, 255, 0, 0, 255, 0, 255, 0, 255, 0, 0, 0, 255, 255, 255, 255, 255, 0]);
  const parts = [Uint8Array.from([137, 80, 78, 71, 13, 10, 26, 10]), chunk('IHDR', ihdr), chunk('IDAT', new Uint8Array(deflateSync(raw))), chunk('IEND', new Uint8Array())];
  const out = new Uint8Array(parts.reduce((n, x) => n + x.length, 0));
  let o = 0;
  for (const x of parts) {
    out.set(x, o);
    o += x.length;
  }
  return out;
}

describe('PDF export', () => {
  it('writes a structurally valid PDF with searchable text, links, outline and comments', async () => {
    const image = `<w:p><w:r><w:drawing><wp:inline><wp:extent cx="914400" cy="914400"/><wp:docPr id="1" name="dot"/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:blipFill><a:blip r:embed="rIdImg"/></pic:blipFill></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>`;
    const body =
      p('Introduction', '<w:pStyle w:val="Heading1"/>') +
      '<w:p><w:r><w:t xml:space="preserve">Read the </w:t></w:r><w:hyperlink r:id="rIdL"><w:r><w:t>policy</w:t></w:r></w:hyperlink><w:r><w:t xml:space="preserve"> or jump to </w:t></w:r><w:hyperlink w:anchor="details"><w:r><w:t>details</w:t></w:r></w:hyperlink><w:r><w:t>. Ünïcödé ✓</w:t></w:r></w:p>' +
      image +
      '<w:p><w:pPr><w:pStyle w:val="Heading2"/></w:pPr><w:bookmarkStart w:id="0" w:name="details"/><w:r><w:t>Details</w:t></w:r><w:bookmarkEnd w:id="0"/></w:p>' +
      p('The fee is due monthly.');
    const doc = await DocxDocument.open(
      await buildDocx({ styles: STYLES, body, rels: [['rIdL', 'hyperlink', 'https://example.com/policy', true], ['rIdImg', 'image', 'media/dot.png']], parts: { 'word/media/dot.png': { data: rgbaPng() } } }),
    );
    const editor = new DocxEditor(doc);
    const id = await editor.addComment(editor.find('fee is due')[0]!, { author: 'Reviewer', text: 'Monthly or quarterly?' });
    await editor.replyToComment(id, { author: 'Owner', text: 'Monthly.' });

    const bytes = await exportPdf(doc, { fonts: testRegistry(), metadata: { title: 'Test document' } });
    const pdf = parsePdf(bytes);
    const [page] = pages(pdf);
    const text = pageText(pdf, page!);
    expect(text).toContain('Introduction');
    expect(text).toContain('Ünïcödé ✓');
    expect(text.replace(/\n/g, '')).toContain('Read the policy or jump to details');

    // Fonts are embedded subsets.
    const fonts = dictOf(pdf, dictOf(pdf, page!.get('Resources')).get('Font'));
    for (const ref of fonts.values()) {
      const font = dictOf(pdf, ref);
      expect((font.get('Subtype') as PdfNameObj).name).toBe('Type0');
      const cid = dictOf(pdf, (font.get('DescendantFonts') as PdfObj[])[0]);
      const descriptor = dictOf(pdf, cid.get('FontDescriptor'));
      expect((descriptor.get('FontName') as PdfNameObj).name).toMatch(/^[A-Z]{6}\+/);
      expect(streamData(pdf, descriptor.get('FontFile2')).length).toBeGreaterThan(1000);
    }

    // Links: one URI action, one internal destination.
    const annots = (page!.get('Annots') as PdfObj[]).map((a) => dictOf(pdf, a));
    const links = annots.filter((a) => (a.get('Subtype') as PdfNameObj).name === 'Link');
    expect(links).toHaveLength(2);
    const uri = dictOf(pdf, links[0]!.get('A')).get('URI') as PdfStr;
    expect(new TextDecoder().decode(uri.bytes)).toBe('https://example.com/policy');
    expect(Array.isArray(links[1]!.get('Dest'))).toBe(true);

    // Comments: a highlight with a reply in reply to it.
    const highlight = annots.find((a) => (a.get('Subtype') as PdfNameObj).name === 'Highlight')!;
    expect(new TextDecoder().decode((highlight.get('Contents') as PdfStr).bytes)).toBe('Monthly or quarterly?');
    expect((highlight.get('QuadPoints') as number[]).length).toBe(8);
    const reply = annots.find((a) => a.has('IRT'))!;
    expect(new TextDecoder().decode((reply.get('Contents') as PdfStr).bytes)).toBe('Monthly.');

    // Outline from headings, named destination for the bookmark, metadata.
    const root = dictOf(pdf, pdf.trailer.get('Root'));
    const outlines = dictOf(pdf, root.get('Outlines'));
    expect(outlines.get('Count')).toBe(2);
    const first = dictOf(pdf, outlines.get('First'));
    expect(new TextDecoder().decode((first.get('Title') as PdfStr).bytes)).toBe('Introduction');
    expect(dictOf(pdf, first.get('First')).has('Dest')).toBe(true);
    const names = dictOf(pdf, dictOf(pdf, root.get('Names')).get('Dests')).get('Names') as PdfObj[];
    expect(new TextDecoder().decode((names[0] as PdfStr).bytes)).toBe('details');
    const info = dictOf(pdf, pdf.trailer.get('Info'));
    expect(new TextDecoder().decode((info.get('Title') as PdfStr).bytes)).toBe('Test document');

    // The RGBA image has a soft mask.
    const xobjects = dictOf(pdf, dictOf(pdf, page!.get('Resources')).get('XObject'));
    const img = resolve(pdf, [...xobjects.values()][0]) as PdfDictObj;
    expect(img.dict.has('SMask')).toBe(true);
    expect(streamData(pdf, img).length).toBe(2 * 2 * 3);
  });

  it('positions glyphs exactly as layout did (justified text)', async () => {
    const lorem = 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. '.repeat(3);
    const doc = await DocxDocument.open(await buildDocx({ styles: STYLES, body: p(lorem, '<w:jc w:val="both"/>') }));
    const pdf = parsePdf(await exportPdf(doc, { fonts: testRegistry() }));
    const content = new TextDecoder('latin1').decode(streamData(pdf, pages(pdf)[0]!.get('Contents')));
    // Justification shows up as TJ adjustments between glyph runs.
    expect(content).toMatch(/> -\d+(\.\d+)? </);
    expect(pageText(pdf, pages(pdf)[0]!).replace(/\n/g, '')).toBe(lorem);
  });
});
