import { describe, expect, it } from 'vitest';
import { DocxEditor, OBJECT_CHAR, paragraphText } from '../src/api/index.js';
import { layoutDocument } from '../src/layout/index.js';
import { blocksToText, DocxDocument } from '../src/model/index.js';
import { htmlToParagraphs } from '../src/render/clipboard.js';
import { serializeNode as serializeXml } from '../src/xml/index.js';
import { buildDocx, p, W_NS } from './helpers/docx.js';
import { testRegistry } from './helpers/fonts.js';

const HDR = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:hdr ${W_NS}><w:p><w:r><w:t>Company header</w:t></w:r></w:p></w:hdr>`;
const FTR = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:ftr ${W_NS}><w:p><w:r><w:t>Footer text</w:t></w:r></w:p></w:ftr>`;
const NOTES = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:footnotes ${W_NS}><w:footnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:footnote><w:footnote w:id="1"><w:p><w:r><w:t>The note.</w:t></w:r></w:p></w:footnote></w:footnotes>`;
const SECT = '<w:sectPr><w:headerReference w:type="default" r:id="rIdH"/><w:footerReference w:type="default" r:id="rIdF"/><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr>';

async function withStories() {
  const body = p('Body text.') + '<w:p><w:r><w:t>See note</w:t></w:r><w:r><w:footnoteReference w:id="1"/></w:r></w:p>';
  return DocxDocument.open(
    await buildDocx({
      body,
      sectPr: SECT,
      rels: [
        ['rIdH', 'header', 'header1.xml'],
        ['rIdF', 'footer', 'footer1.xml'],
        ['rIdN', 'footnotes', 'footnotes.xml'],
      ],
      parts: {
        'word/header1.xml': { data: HDR, contentType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml' },
        'word/footer1.xml': { data: FTR, contentType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml' },
        'word/footnotes.xml': { data: NOTES, contentType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml' },
      },
    }),
  );
}

describe('editing headers, footers and notes', () => {
  it('lists the story parts and edits a header', async () => {
    const doc = await withStories();
    const parts = doc.storyParts();
    expect(parts.map((s) => `${s.kind}:${s.part}`)).toEqual(['header:word/header1.xml', 'footer:word/footer1.xml', 'footnotes:word/footnotes.xml']);
    const header = new DocxEditor(doc, undefined, { part: 'word/header1.xml' });
    expect(header.paragraphs().map((x) => paragraphText(x).text)).toEqual(['Company header']);
    header.replaceRange(header.find('Company')[0]!, 'Acme');
    const again = await DocxDocument.open(await doc.save());
    expect(blocksToText(again.headerFooter('word/header1.xml').blocks)).toBe('Acme header');
    // The body is untouched.
    expect(blocksToText(again.body().blocks)).toContain('Body text.');
  });

  it('edits footnote text, skipping separator notes', async () => {
    const doc = await withStories();
    const notes = new DocxEditor(doc, undefined, { part: 'word/footnotes.xml' });
    expect(notes.paragraphs().map((x) => paragraphText(x).text)).toEqual(['The note.']);
    notes.insertText(notes.find('note')[0]!.start, 'first ');
    expect(blocksToText(doc.note('footnote', '1'))).toBe('The first note.');
  });

  it('records header, footer and note lines for in-page editing', async () => {
    const doc = await withStories();
    const page = layoutDocument(doc, { fonts: testRegistry() }).pages[0]!;
    const kinds = page.storyLines.map((l) => `${l.story}:${l.part}${l.noteId ? `#${l.noteId}` : ''}`);
    expect(kinds).toEqual(['header:word/header1.xml', 'footer:word/footer1.xml', 'footnote:word/footnotes.xml#1']);
    expect(page.lines.every((l) => !('story' in l))).toBe(true);
    // Header lines sit in the top margin, the footer's near the bottom.
    expect(page.storyLines[0]!.y).toBeLessThan(72);
    expect(page.storyLines[1]!.y).toBeGreaterThan(720);
  });
});

describe('paragraph formatting', () => {
  it('sets alignment, indents, spacing and direction in schema order', async () => {
    const doc = await DocxDocument.open(await buildDocx({ body: p('One') }));
    const editor = new DocxEditor(doc);
    const para = editor.paragraphs()[0]!.el;
    editor.formatParagraph(para, { align: 'center', indentLeft: 36, indentFirstLine: -18, spaceAfter: 12, lineSpacing: 1.5, keepWithNext: true });
    expect(serializeXml(para)).toContain('<w:pPr><w:keepNext/><w:spacing w:after="240" w:line="360" w:lineRule="auto"/><w:ind w:left="720" w:hanging="360"/><w:jc w:val="center"/></w:pPr>');
    const props = editor.paragraphs()[0]!.props;
    expect(props.jc).toBe('center');
    expect(props.indent).toMatchObject({ left: 36, hanging: 18 });
    // Right-to-left: visual right alignment is stored as Word's logical "left".
    editor.formatParagraph(para, { direction: 'rtl' });
    editor.formatParagraph(para, { align: 'right' });
    expect(serializeXml(para)).toContain('<w:bidi/>');
    expect(serializeXml(para)).toContain('<w:jc w:val="left"/>');
  });
});

describe('inline pictures as caret positions', () => {
  it('counts an inline picture as one character that can be deleted', async () => {
    const pic = '<w:r><w:drawing><wp:inline><wp:extent cx="254000" cy="254000"/><wp:docPr id="1" name="P"/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:blipFill><a:blip r:embed="rIdX"/></pic:blipFill></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>';
    const doc = await DocxDocument.open(await buildDocx({ body: `<w:p><w:r><w:t>ab</w:t></w:r>${pic}<w:r><w:t>cd</w:t></w:r></w:p>` }));
    const editor = new DocxEditor(doc);
    const para = editor.paragraphs()[0]!;
    expect(paragraphText(para).text).toBe('abcd');
    const pt = paragraphText(para, { objects: true });
    expect(pt.text).toBe(`ab${OBJECT_CHAR}cd`);
    const seg = pt.segments[1]!;
    expect(seg.node.name).toBe('w:drawing');
    editor.deleteRange({ start: { node: seg.node, offset: 0 }, end: { node: seg.node, offset: 1 } });
    expect(serializeXml(doc.documentXml.root)).not.toContain('w:drawing');
    expect(paragraphText(editor.paragraphs()[0]!).text).toBe('abcd');
  });
});

describe('pasting HTML', () => {
  it('keeps bold, italic, links, headings and lists', () => {
    const html = '<html><body><!--StartFragment--><h2>Title here</h2><p>Some <b>bold</b> and <i>italic</i> and <a href="https://example.com">a link</a>.</p><ul><li>One</li><li>Two</li></ul><ol><li>First</li></ol><!--EndFragment--></body></html>';
    const out = htmlToParagraphs(html);
    expect(out.map((x) => [x.style, x.list, x.runs.map((r) => r.text).join('')])).toEqual([
      ['Heading2', undefined, 'Title here'],
      [undefined, undefined, 'Some bold and italic and a link.'],
      [undefined, 'bullet', 'One'],
      [undefined, 'bullet', 'Two'],
      [undefined, 'decimal', 'First'],
    ]);
    const runs = out[1]!.runs;
    expect(runs.find((r) => r.text === 'bold')!.format).toEqual({ bold: true });
    expect(runs.find((r) => r.text === 'italic')!.format).toEqual({ italic: true });
    expect(runs.find((r) => r.text === 'a link')!.link).toBe('https://example.com');
  });

  it("reads Google Docs' CSS and ignores scripts, styles and unsafe links", () => {
    const html =
      '<meta charset="utf-8"><style>p{color:red}</style><b style="font-weight:normal;" id="docs-internal-guid-1"><p dir="ltr"><span style="font-size:14pt;font-family:Arial;color:#c00000;font-weight:700;">Red bold</span><span style="font-style:italic;text-decoration:underline;"> under</span></p></b>' +
      '<script>alert(1)</script><p><a href="javascript:alert(1)">bad</a>&nbsp;&amp;&#8364;</p>';
    const out = htmlToParagraphs(html);
    expect(out[0]!.runs).toEqual([
      { text: 'Red bold', format: { bold: true, size: 14, font: 'Arial', color: 'C00000' }, link: undefined },
      { text: ' under', format: { italic: true, underline: true }, link: undefined },
    ]);
    expect(out[1]!.runs.map((r) => [r.text, r.link])).toEqual([['bad &€', undefined]]);
    expect(JSON.stringify(out)).not.toContain('alert');
  });

  it('turns line breaks and table cells into breaks and tabs', () => {
    const out = htmlToParagraphs('<table><tr><td>a</td><td>b</td></tr></table><p>line<br>next</p>');
    expect(out.map((x) => x.runs.map((r) => r.text).join(''))).toEqual(['a\tb', 'line\nnext']);
  });
});
