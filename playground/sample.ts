/** A sample document exercising headings, lists, tables, links, bookmarks, comments and ranges. */
import { buildDocx, p } from '../test/helpers/docx.js';

const STYLES = `
<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/><w:sz w:val="22"/></w:rPr></w:rPrDefault>
<w:pPrDefault><w:pPr><w:spacing w:after="160" w:line="259" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults>
<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>
<w:style w:type="paragraph" w:styleId="Title"><w:name w:val="Title"/><w:basedOn w:val="Normal"/><w:pPr><w:spacing w:after="80"/></w:pPr><w:rPr><w:rFonts w:ascii="Calibri Light" w:hAnsi="Calibri Light"/><w:sz w:val="56"/><w:color w:val="1F3864"/></w:rPr></w:style>
<w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:basedOn w:val="Normal"/><w:pPr><w:keepNext/><w:spacing w:before="240" w:after="80"/><w:outlineLvl w:val="0"/></w:pPr><w:rPr><w:rFonts w:ascii="Calibri Light" w:hAnsi="Calibri Light"/><w:color w:val="2F5496"/><w:sz w:val="32"/></w:rPr></w:style>
<w:style w:type="character" w:styleId="Hyperlink"><w:name w:val="Hyperlink"/><w:rPr><w:color w:val="0563C1"/><w:u w:val="single"/></w:rPr></w:style>
<w:style w:type="table" w:styleId="GridTable"><w:name w:val="Grid Table"/><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr>
  <w:tblPr><w:tblBorders><w:top w:val="single" w:sz="4" w:color="8EAADB"/><w:left w:val="single" w:sz="4" w:color="8EAADB"/><w:bottom w:val="single" w:sz="4" w:color="8EAADB"/><w:right w:val="single" w:sz="4" w:color="8EAADB"/><w:insideH w:val="single" w:sz="4" w:color="8EAADB"/><w:insideV w:val="single" w:sz="4" w:color="8EAADB"/></w:tblBorders><w:tblCellMar><w:top w:w="40" w:type="dxa"/><w:left w:w="108" w:type="dxa"/><w:bottom w:w="40" w:type="dxa"/><w:right w:w="108" w:type="dxa"/></w:tblCellMar></w:tblPr>
  <w:tblStylePr w:type="firstRow"><w:rPr><w:b/><w:color w:val="FFFFFF"/></w:rPr><w:tcPr><w:shd w:val="clear" w:color="auto" w:fill="4472C4"/></w:tcPr></w:tblStylePr>
  <w:tblStylePr w:type="band1Horz"><w:tcPr><w:shd w:val="clear" w:color="auto" w:fill="D9E2F3"/></w:tcPr></w:tblStylePr>
</w:style>`;

const NUMBERING = `
<w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/><w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr></w:lvl>
<w:lvl w:ilvl="1"><w:start w:val="1"/><w:numFmt w:val="lowerLetter"/><w:lvlText w:val="(%2)"/><w:pPr><w:ind w:left="1440" w:hanging="360"/></w:pPr></w:lvl></w:abstractNum>
<w:abstractNum w:abstractNumId="1"><w:lvl w:ilvl="0"><w:numFmt w:val="bullet"/><w:lvlText w:val=""/><w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr><w:rPr><w:rFonts w:ascii="Symbol" w:hAnsi="Symbol"/></w:rPr></w:lvl></w:abstractNum>
<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num><w:num w:numId="2"><w:abstractNumId w:val="1"/></w:num>`;

const COMMENTS = `
<w:comment w:id="0" w:author="Rina Pratama" w:date="2026-10-28T09:12:00Z" w:initials="RP"><w:p w14:paraId="1A2B3C01"><w:r><w:t>Please confirm the royalty rate with the finance team.</w:t></w:r></w:p></w:comment>
<w:comment w:id="1" w:author="Budi Santoso" w:date="2026-10-28T10:40:00Z" w:initials="BS"><w:p w14:paraId="1A2B3C02"><w:r><w:t>Confirmed: 3.75% as agreed in the term sheet.</w:t></w:r></w:p></w:comment>`;

const LOREM =
  'The Contractor shall carry out the mining operations in accordance with good mining practice, applicable laws and the approved work plan. All activities shall be conducted in a manner that protects the environment, the health and safety of workers, and the interests of local communities.';

const num = (text: string, lvl = 0, id = 1) => p(text, `<w:numPr><w:ilvl w:val="${lvl}"/><w:numId w:val="${id}"/></w:numPr>`);
const cell = (text: string, w = 3120) => `<w:tc><w:tcPr><w:tcW w:w="${w}" w:type="dxa"/></w:tcPr>${p(text)}</w:tc>`;

export async function sampleDocx(): Promise<Uint8Array> {
  const body = [
    p('Mining Services Agreement', '<w:pStyle w:val="Title"/>'),
    '<w:p><w:r><w:t xml:space="preserve">Draft for review · prepared with </w:t></w:r><w:hyperlink r:id="rIdSite"><w:r><w:rPr><w:rStyle w:val="Hyperlink"/></w:rPr><w:t>docxkit</w:t></w:r></w:hyperlink><w:r><w:t xml:space="preserve">. Jump to </w:t></w:r><w:hyperlink w:anchor="payment"><w:r><w:rPr><w:rStyle w:val="Hyperlink"/></w:rPr><w:t>Payment terms</w:t></w:r></w:hyperlink><w:r><w:t>.</w:t></w:r></w:p>',
    p('1. Scope of work', '<w:pStyle w:val="Heading1"/>'),
    p(LOREM, '<w:jc w:val="both"/>'),
    num('Exploration drilling and sampling'),
    num('Overburden removal', 0),
    num('Topsoil stockpiling', 1),
    num('Progressive rehabilitation', 1),
    num('Coal hauling to the port'),
    '<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:bookmarkStart w:id="10" w:name="payment"/><w:r><w:t>2. Payment terms</w:t></w:r><w:bookmarkEnd w:id="10"/></w:p>',
    '<w:p><w:pPr><w:jc w:val="both"/></w:pPr><w:r><w:t xml:space="preserve">The Company shall pay a royalty of </w:t></w:r><w:commentRangeStart w:id="0"/><w:r><w:rPr><w:b/></w:rPr><w:t>3.75% of the gross revenue</w:t></w:r><w:commentRangeEnd w:id="0"/><w:r><w:commentReference w:id="0"/></w:r><w:r><w:t xml:space="preserve"> within thirty (30) days after the end of each quarter. Late payments bear interest at the rate set out below.</w:t></w:r></w:p>',
    `<w:tbl><w:tblPr><w:tblStyle w:val="GridTable"/><w:tblW w:w="0" w:type="auto"/><w:tblLook w:firstRow="1" w:lastRow="0" w:firstColumn="0" w:lastColumn="0" w:noHBand="0" w:noVBand="1"/></w:tblPr><w:tblGrid><w:gridCol w:w="3120"/><w:gridCol w:w="3120"/><w:gridCol w:w="3120"/></w:tblGrid>
      <w:tr><w:trPr><w:tblHeader/></w:trPr>${cell('Milestone')}${cell('Due')}${cell('Amount (USD)')}</w:tr>
      <w:tr>${cell('Mobilisation')}${cell('Signing + 30 days')}${cell('250,000')}</w:tr>
      <w:tr>${cell('First coal')}${cell('Q2 2027')}${cell('1,200,000')}</w:tr>
      <w:tr>${cell('Steady state')}${cell('Quarterly')}${cell('Per royalty schedule')}</w:tr></w:tbl>`,
    p(''),
    p('3. Editable section', '<w:pStyle w:val="Heading1"/>'),
    '<w:p><w:r><w:t xml:space="preserve">Counterparty notes: </w:t></w:r><w:permStart w:id="20" w:edGrp="everyone"/><w:r><w:t>[to be completed by the counterparty]</w:t></w:r><w:permEnd w:id="20"/></w:p>',
    num('Bulleted point one', 0, 2),
    num('Bulleted point two', 0, 2),
    p('4. General', '<w:pStyle w:val="Heading1"/>'),
    ...Array.from({ length: 14 }, (_, i) => p(`${i + 1}. ${LOREM}`, '<w:jc w:val="both"/>')),
  ].join('');
  const footer = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:ftr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:rPr><w:color w:val="7F7F7F"/><w:sz w:val="18"/></w:rPr><w:t xml:space="preserve">Page </w:t></w:r><w:fldSimple w:instr="PAGE"><w:r><w:rPr><w:color w:val="7F7F7F"/><w:sz w:val="18"/></w:rPr><w:t>1</w:t></w:r></w:fldSimple><w:r><w:rPr><w:color w:val="7F7F7F"/><w:sz w:val="18"/></w:rPr><w:t xml:space="preserve"> of </w:t></w:r><w:fldSimple w:instr="NUMPAGES"><w:r><w:rPr><w:color w:val="7F7F7F"/><w:sz w:val="18"/></w:rPr><w:t>1</w:t></w:r></w:fldSimple></w:p></w:ftr>`;
  const header = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:p><w:pPr><w:pBdr><w:bottom w:val="single" w:sz="4" w:space="4" w:color="BFBFBF"/></w:pBdr><w:tabs><w:tab w:val="right" w:pos="9360"/></w:tabs></w:pPr><w:r><w:rPr><w:sz w:val="18"/><w:color w:val="7F7F7F"/></w:rPr><w:t>PT Contoh Tambang</w:t></w:r><w:r><w:rPr><w:sz w:val="18"/><w:color w:val="7F7F7F"/></w:rPr><w:tab/><w:t>CONFIDENTIAL</w:t></w:r></w:p></w:hdr>`;
  const sect =
    '<w:sectPr><w:headerReference w:type="default" r:id="rIdH"/><w:footerReference w:type="default" r:id="rIdF"/><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr>';
  const commentsReply = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w15:commentsEx xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml"><w15:commentEx w15:paraId="1A2B3C01" w15:done="0"/><w15:commentEx w15:paraId="1A2B3C02" w15:paraIdParent="1A2B3C01" w15:done="0"/></w15:commentsEx>`;
  return buildDocx({
    styles: STYLES,
    numbering: NUMBERING,
    comments: COMMENTS,
    body,
    sectPr: sect,
    rels: [
      ['rIdSite', 'hyperlink', 'https://example.com/docxkit', true],
      ['rIdH', 'header', 'header1.xml'],
      ['rIdF', 'footer', 'footer1.xml'],
      ['rIdCx', 'http://schemas.microsoft.com/office/2011/relationships/commentsExtended', 'commentsExtended.xml'],
    ],
    parts: {
      'word/header1.xml': { data: header, contentType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml' },
      'word/footer1.xml': { data: footer, contentType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml' },
      'word/commentsExtended.xml': { data: commentsReply, contentType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.commentsExtended+xml' },
    },
  });
}
