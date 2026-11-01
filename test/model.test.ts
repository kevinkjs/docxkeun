import { describe, expect, it } from 'vitest';
import { DocxDocument, NumberingCounter, blocksToText, formatNumber, parseFieldInstruction, type DocParagraph, type DocTable, type Inline } from '../src/model/index.js';
import { buildDocx, p } from './helpers/docx.js';

const STYLES = `
<w:docDefaults>
  <w:rPrDefault><w:rPr><w:rFonts w:asciiTheme="minorHAnsi" w:hAnsiTheme="minorHAnsi"/><w:sz w:val="22"/></w:rPr></w:rPrDefault>
  <w:pPrDefault><w:pPr><w:spacing w:after="160" w:line="259" w:lineRule="auto"/></w:pPr></w:pPrDefault>
</w:docDefaults>
<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>
<w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:basedOn w:val="Normal"/>
  <w:pPr><w:keepNext/><w:spacing w:before="240" w:after="0"/><w:outlineLvl w:val="0"/></w:pPr>
  <w:rPr><w:b/><w:sz w:val="32"/><w:color w:val="2F5496"/></w:rPr></w:style>
<w:style w:type="character" w:styleId="Strong"><w:name w:val="Strong"/><w:rPr><w:b/></w:rPr></w:style>
<w:style w:type="table" w:styleId="Grid"><w:name w:val="Grid"/>
  <w:tblPr><w:tblBorders><w:top w:val="single" w:sz="8"/><w:insideH w:val="single" w:sz="4"/></w:tblBorders></w:tblPr>
  <w:tblStylePr w:type="firstRow"><w:rPr><w:b/></w:rPr><w:tcPr><w:shd w:val="clear" w:fill="DDDDDD"/></w:tcPr></w:tblStylePr>
</w:style>`;

const NUMBERING = `
<w:abstractNum w:abstractNumId="0">
  <w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/><w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr></w:lvl>
  <w:lvl w:ilvl="1"><w:start w:val="1"/><w:numFmt w:val="lowerLetter"/><w:lvlText w:val="%1.%2)"/><w:pPr><w:ind w:left="1440" w:hanging="360"/></w:pPr></w:lvl>
</w:abstractNum>
<w:abstractNum w:abstractNumId="1"><w:lvl w:ilvl="0"><w:numFmt w:val="bullet"/><w:lvlText w:val="•"/></w:lvl></w:abstractNum>
<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>
<w:num w:numId="2"><w:abstractNumId w:val="0"/><w:lvlOverride w:ilvl="0"><w:startOverride w:val="5"/></w:lvlOverride></w:num>
<w:num w:numId="3"><w:abstractNumId w:val="1"/></w:num>`;

function paragraphs(doc: DocxDocument): DocParagraph[] {
  return doc.body().blocks.filter((b): b is DocParagraph => b.kind === 'paragraph');
}

function kinds(inlines: Inline[]): string[] {
  return inlines.map((i) => (i.kind === 'text' ? `text:${i.text}` : i.kind));
}

describe('styles and properties', () => {
  it('resolves paragraph and run properties through the style chain', async () => {
    const doc = await DocxDocument.open(
      await buildDocx({
        styles: STYLES,
        body: p('Title', '<w:pStyle w:val="Heading1"/>') + p('Body text') + '<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:rPr><w:rStyle w:val="Strong"/></w:rPr><w:t>x</w:t></w:r><w:r><w:rPr><w:b w:val="0"/><w:i/></w:rPr><w:t>y</w:t></w:r></w:p>',
      }),
    );
    const [h, body, mixed] = paragraphs(doc);
    expect(h!.styleId).toBe('Heading1');
    expect(h!.props.spacing).toEqual({ before: 12, after: 0, line: 259 / 240, lineRule: 'auto' });
    expect(h!.props.keepNext).toBe(true);
    const hRun = h!.inlines[0]!;
    expect(hRun.kind === 'text' && hRun.props).toMatchObject({ bold: true, size: 16, color: '2F5496', fonts: { asciiTheme: 'minorHAnsi' } });
    expect(body!.styleId).toBe('Normal');
    const bodyRun = body!.inlines[0]!;
    expect(bodyRun.kind === 'text' && bodyRun.props).toMatchObject({ bold: false, size: 11 });
    // Heading (bold) + Strong (bold) toggle each other off; direct formatting wins absolutely.
    const [x, y] = mixed!.inlines;
    expect(x!.kind === 'text' && x!.props.bold).toBe(false);
    expect(y!.kind === 'text' && y!.props).toMatchObject({ bold: false, italic: true });
  });

  it('applies table style conditional formatting', async () => {
    const tbl = `<w:tbl><w:tblPr><w:tblStyle w:val="Grid"/><w:tblLook w:firstRow="1" w:lastRow="0" w:firstColumn="0" w:lastColumn="0" w:noHBand="1" w:noVBand="1"/></w:tblPr>
      <w:tblGrid><w:gridCol w:w="2000"/><w:gridCol w:w="3000"/></w:tblGrid>
      <w:tr><w:trPr><w:tblHeader/></w:trPr><w:tc>${p('H1')}</w:tc><w:tc>${p('H2')}</w:tc></w:tr>
      <w:tr><w:tc><w:tcPr><w:gridSpan w:val="2"/></w:tcPr>${p('Body')}</w:tc></w:tr></w:tbl>`;
    const doc = await DocxDocument.open(await buildDocx({ styles: STYLES, body: tbl + '<w:p/>' }));
    const table = doc.body().blocks[0] as DocTable;
    expect(table.kind).toBe('table');
    expect(table.grid).toEqual([100, 150]);
    expect(table.props.borders?.top?.width).toBe(1);
    expect(table.rows[0]!.props.isHeader).toBe(true);
    expect(table.rows[0]!.cells[0]!.props.shading?.fill).toBe('DDDDDD');
    expect(table.rows[1]!.cells[0]!.props.shading).toBeUndefined();
    expect(table.rows[1]!.cells[0]!.gridSpan).toBe(2);
    const headerRun = (table.rows[0]!.cells[0]!.blocks[0] as DocParagraph).inlines[0]!;
    expect(headerRun.kind === 'text' && headerRun.props.bold).toBe(true);
    const bodyRun = (table.rows[1]!.cells[0]!.blocks[0] as DocParagraph).inlines[0]!;
    expect(bodyRun.kind === 'text' && bodyRun.props.bold).toBe(false);
  });

  it('reads sections', async () => {
    const body = p('a', '<w:sectPr><w:pgSz w:w="16838" w:h="11906" w:orient="landscape"/><w:type w:val="continuous"/></w:sectPr>') + p('b');
    const doc = await DocxDocument.open(await buildDocx({ body }));
    const [first] = paragraphs(doc);
    expect(first!.sectionEnd).toMatchObject({ pageWidth: 16838 / 20, orientation: 'landscape', type: 'continuous' });
    expect(doc.body().finalSection).toMatchObject({ pageWidth: 612, pageHeight: 792 });
  });
});

describe('numbering', () => {
  it('produces labels with nesting, shared counters and start overrides', async () => {
    const num = (id: string, lvl: number, text: string) => p(text, `<w:numPr><w:ilvl w:val="${lvl}"/><w:numId w:val="${id}"/></w:numPr>`);
    const doc = await DocxDocument.open(
      await buildDocx({
        numbering: NUMBERING,
        body: num('1', 0, 'one') + num('1', 1, 'one-a') + num('1', 1, 'one-b') + num('1', 0, 'two') + num('1', 1, 'two-a') + num('2', 0, 'restart') + num('3', 0, 'bullet'),
      }),
    );
    const counter = new NumberingCounter(doc.numbering);
    const labels = paragraphs(doc).map((para) => counter.next(para.numbering?.numId, para.numbering?.ilvl ?? 0)?.label);
    expect(labels).toEqual(['1.', '1.a)', '1.b)', '2.', '2.a)', '5.', '•']);
    expect(paragraphs(doc)[1]!.props.indent).toEqual({ left: 72, hanging: 18 });
  });

  it('formats numbers', () => {
    expect(formatNumber(14, 'upperRoman')).toBe('XIV');
    expect(formatNumber(28, 'lowerLetter')).toBe('bb');
    expect(formatNumber(3, 'ordinal')).toBe('3rd');
    expect(formatNumber(12, 'ordinal')).toBe('12th');
    expect(formatNumber(21, 'cardinalText')).toBe('Twenty-one');
  });
});

describe('inline content', () => {
  it('flattens hyperlinks, bookmarks, comments and permission ranges', async () => {
    const body =
      '<w:p><w:bookmarkStart w:id="0" w:name="intro"/><w:r><w:t>Hi </w:t></w:r><w:bookmarkEnd w:id="0"/>' +
      '<w:hyperlink r:id="rIdLink" w:tooltip="Go"><w:r><w:t>site</w:t></w:r></w:hyperlink>' +
      '<w:hyperlink w:anchor="intro"><w:r><w:t>top</w:t></w:r></w:hyperlink>' +
      '<w:commentRangeStart w:id="3"/><w:r><w:t>noted</w:t></w:r><w:commentRangeEnd w:id="3"/><w:r><w:commentReference w:id="3"/></w:r>' +
      '<w:permStart w:id="9" w:edGrp="everyone"/><w:r><w:t>editable</w:t></w:r><w:permEnd w:id="9"/>' +
      '<w:del w:id="1"><w:r><w:delText>gone</w:delText></w:r></w:del><w:ins w:id="2"><w:r><w:t>added</w:t></w:r></w:ins></w:p>';
    const doc = await DocxDocument.open(await buildDocx({ body, rels: [['rIdLink', 'hyperlink', 'https://example.com/', true]] }));
    const [para] = paragraphs(doc);
    expect(kinds(para!.inlines)).toEqual([
      'bookmarkStart',
      'text:Hi ',
      'bookmarkEnd',
      'linkStart',
      'text:site',
      'linkEnd',
      'linkStart',
      'text:top',
      'linkEnd',
      'commentStart',
      'text:noted',
      'commentEnd',
      'commentRef',
      'permStart',
      'text:editable',
      'permEnd',
      'text:added',
    ]);
    const links = para!.inlines.filter((i) => i.kind === 'linkStart');
    expect(links.map((l) => l.kind === 'linkStart' && l.target)).toEqual([
      { url: 'https://example.com/', anchor: undefined, tooltip: 'Go' },
      { url: undefined, anchor: 'intro', tooltip: undefined },
    ]);
    const perm = para!.inlines.find((i) => i.kind === 'permStart');
    expect(perm).toMatchObject({ id: '9', group: 'everyone' });
  });

  it('turns complex fields into links and dynamic page numbers', async () => {
    const fld = (instr: string, result: string) =>
      `<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve">${instr}</w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>${result}</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>`;
    const body =
      `<w:p>${fld(' HYPERLINK "https://x.test/a" \\o "tip" ', 'link text')}<w:r><w:t xml:space="preserve"> page </w:t></w:r>${fld(' PAGE \\* MERGEFORMAT ', '7')}<w:r><w:t xml:space="preserve"> of </w:t></w:r><w:fldSimple w:instr=" NUMPAGES "><w:r><w:t>9</w:t></w:r></w:fldSimple></w:p>` +
      `<w:p>${fld(' PAGEREF _Toc1 \\h ', '3')}</w:p>`;
    const doc = await DocxDocument.open(await buildDocx({ body }));
    const [a, b] = paragraphs(doc);
    expect(kinds(a!.inlines)).toEqual(['linkStart', 'text:link text', 'linkEnd', 'text: page ', 'field', 'text: of ', 'field']);
    const [page, numpages] = a!.inlines.filter((i) => i.kind === 'field');
    expect(page).toMatchObject({ field: 'PAGE', cached: '7', format: 'MERGEFORMAT' });
    expect(numpages).toMatchObject({ field: 'NUMPAGES', cached: '9' });
    expect(a!.inlines[0]).toMatchObject({ target: { url: 'https://x.test/a', tooltip: 'tip' } });
    expect(kinds(b!.inlines)).toEqual(['linkStart', 'text:3', 'linkEnd']);
    expect(b!.inlines[0]).toMatchObject({ target: { anchor: '_Toc1' } });
  });

  it('reads images', async () => {
    const drawing = `<w:p><w:r><w:drawing><wp:inline><wp:extent cx="914400" cy="457200"/><wp:docPr id="1" name="Pic" descr="A chart"/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:blipFill><a:blip r:embed="rIdImg"/></pic:blipFill></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>`;
    const doc = await DocxDocument.open(
      await buildDocx({ body: drawing, rels: [['rIdImg', 'image', 'media/image1.png']], parts: { 'word/media/image1.png': { data: new Uint8Array([137, 80, 78, 71]) } } }),
    );
    const img = paragraphs(doc)[0]!.inlines[0]!;
    expect(img).toMatchObject({ kind: 'image', target: 'word/media/image1.png', width: 72, height: 36, alt: 'A chart' });
    expect([...(await doc.media('word/media/image1.png'))]).toEqual([137, 80, 78, 71]);
  });

  it('extracts plain text', async () => {
    const doc = await DocxDocument.open(await buildDocx({ body: p('Line one') + '<w:p><w:r><w:t>a</w:t><w:tab/><w:t>b</w:t></w:r></w:p>' }));
    expect(blocksToText(doc.body().blocks)).toBe('Line one\na\tb');
  });

  it('parses field instructions', () => {
    const f = parseFieldInstruction(' HYPERLINK "http://a.b/c d" \\l "sec 2" \\o "Tip \\"quoted\\"" \\n ');
    expect(f.type).toBe('HYPERLINK');
    expect(f.args).toEqual(['http://a.b/c d']);
    expect(f.switches.get('\\l')).toBe('sec 2');
    expect(f.switches.get('\\o')).toBe('Tip "quoted"');
    expect(f.switches.get('\\n')).toBe(true);
  });
});

describe('document', () => {
  it('creates a blank document that re-opens', async () => {
    const doc = await DocxDocument.create();
    const reopened = await DocxDocument.open(await doc.save());
    expect(reopened.styles.get('Heading1')?.name).toBe('heading 1');
    expect(reopened.body().finalSection?.pageWidth).toBeCloseTo(595.3, 1);
  });
});
