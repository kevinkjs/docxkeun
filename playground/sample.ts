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
<w:abstractNum w:abstractNumId="1"><w:lvl w:ilvl="0"><w:numFmt w:val="bullet"/><w:lvlText w:val="\uF0B7"/><w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr><w:rPr><w:rFonts w:ascii="Symbol" w:hAnsi="Symbol"/></w:rPr></w:lvl></w:abstractNum>
<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num><w:num w:numId="2"><w:abstractNumId w:val="1"/></w:num>`;

const COMMENTS = `
<w:comment w:id="0" w:author="Rina Pratama" w:date="2026-10-28T09:12:00Z" w:initials="RP"><w:p w14:paraId="1A2B3C01"><w:r><w:t>Please confirm the royalty rate with the finance team.</w:t></w:r></w:p></w:comment>
<w:comment w:id="1" w:author="Budi Santoso" w:date="2026-10-28T10:40:00Z" w:initials="BS"><w:p w14:paraId="1A2B3C02"><w:r><w:t>Confirmed: 3.75% as agreed in the term sheet.</w:t></w:r></w:p></w:comment>`;

const LOREM =
  'The Contractor shall carry out the mining operations in accordance with good mining practice, applicable laws and the approved work plan. All activities shall be conducted in a manner that protects the environment, the health and safety of workers, and the interests of local communities.';

const num = (text: string, lvl = 0, id = 1) => p(text, `<w:numPr><w:ilvl w:val="${lvl}"/><w:numId w:val="${id}"/></w:numPr>`);
const cell = (text: string, w = 3120) => `<w:tc><w:tcPr><w:tcW w:w="${w}" w:type="dxa"/></w:tcPr>${p(text)}</w:tc>`;

const series = (idx: number, name: string, color: string, values: number[]) =>
  `<c:ser><c:idx val="${idx}"/><c:order val="${idx}"/><c:tx><c:v>${name}</c:v></c:tx><c:spPr><a:solidFill><a:srgbClr val="${color}"/></a:solidFill></c:spPr>` +
  `<c:cat><c:strLit><c:ptCount val="4"/>${['Q1', 'Q2', 'Q3', 'Q4'].map((q, i) => `<c:pt idx="${i}"><c:v>${q}</c:v></c:pt>`).join('')}</c:strLit></c:cat>` +
  `<c:val><c:numLit><c:formatCode>#,##0</c:formatCode><c:ptCount val="4"/>${values.map((v, i) => `<c:pt idx="${i}"><c:v>${v}</c:v></c:pt>`).join('')}</c:numLit></c:val></c:ser>`;
const CHART =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><c:chart>' +
  '<c:title><c:tx><c:rich><a:bodyPr/><a:p><a:r><a:t>Coal hauled, 2027</a:t></a:r></a:p></c:rich></c:tx><c:overlay val="0"/></c:title><c:autoTitleDeleted val="0"/><c:plotArea><c:barChart><c:barDir val="col"/><c:grouping val="clustered"/><c:varyColors val="0"/>' +
  series(0, 'Plan', '9DC3E6', [420, 610, 680, 700]) +
  series(1, 'Actual', '2E74B5', [380, 640, 655, 720]) +
  '<c:gapWidth val="80"/><c:axId val="1"/><c:axId val="2"/></c:barChart><c:catAx><c:axId val="1"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="b"/><c:crossAx val="2"/></c:catAx><c:valAx><c:axId val="2"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="l"/><c:majorGridlines/><c:numFmt formatCode="#,##0" sourceLinked="0"/><c:crossAx val="1"/></c:valAx></c:plotArea><c:legend><c:legendPos val="b"/><c:overlay val="0"/></c:legend><c:plotVisOnly val="1"/></c:chart></c:chartSpace>';

export async function sampleDocx(): Promise<Uint8Array> {
  const link = (text: string, anchor: string) => `<w:hyperlink w:anchor="${anchor}"><w:r><w:rPr><w:rStyle w:val="Hyperlink"/></w:rPr><w:t xml:space="preserve">${text}</w:t></w:r></w:hyperlink>`;
  const heading = (text: string, bookmark: string, id: number) =>
    `<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:bookmarkStart w:id="${id}" w:name="${bookmark}"/><w:r><w:t>${text}</w:t></w:r><w:bookmarkEnd w:id="${id}"/></w:p>`;
  const sep = '<w:r><w:t xml:space="preserve"> · </w:t></w:r>';
  const body = [
    '<w:p><w:pPr><w:pStyle w:val="Title"/></w:pPr><w:bookmarkStart w:id="9" w:name="top"/><w:r><w:t>Mining Services Agreement</w:t></w:r><w:bookmarkEnd w:id="9"/></w:p>',
    '<w:p><w:r><w:t xml:space="preserve">Draft for review · prepared with </w:t></w:r><w:hyperlink r:id="rIdSite"><w:r><w:rPr><w:rStyle w:val="Hyperlink"/></w:rPr><w:t>docxkeun</w:t></w:r></w:hyperlink><w:r><w:t>.</w:t></w:r></w:p>',
    `<w:p><w:r><w:rPr><w:b/></w:rPr><w:t xml:space="preserve">Contents: </w:t></w:r>${link('Scope', 'scope')}${sep}${link('Payment terms', 'payment')}${sep}${link('Editable section', 'editable')}${sep}${link('General', 'general')}${sep}${link('Figures', 'figures')}${sep}${link('Signatures', 'signatures')}</w:p>`,
    heading('1. Scope of work', 'scope', 11),
    p(LOREM, '<w:jc w:val="both"/>'),
    num('Exploration drilling and sampling'),
    num('Overburden removal', 0),
    num('Topsoil stockpiling', 1),
    num('Progressive rehabilitation', 1),
    num('Coal hauling to the port'),
    heading('2. Payment terms', 'payment', 10),
    '<w:p><w:pPr><w:jc w:val="both"/></w:pPr><w:r><w:t xml:space="preserve">The Company shall pay a royalty of </w:t></w:r><w:commentRangeStart w:id="0"/><w:sdt><w:sdtPr><w:alias w:val="Agreed royalty rate"/><w:tag w:val="docxkeun-lock"/><w:id w:val="501"/><w:lock w:val="contentLocked"/></w:sdtPr><w:sdtContent><w:r><w:rPr><w:b/></w:rPr><w:t>3.75% of the gross revenue</w:t></w:r></w:sdtContent></w:sdt><w:commentRangeEnd w:id="0"/><w:r><w:commentReference w:id="0"/></w:r><w:r><w:t xml:space="preserve"> within thirty (30) days after the end of each quarter. Late payments bear interest at the rate set out below. See also </w:t></w:r>' + link('the general terms', 'general') + '<w:r><w:t>.</w:t></w:r></w:p>',
    `<w:tbl><w:tblPr><w:tblStyle w:val="GridTable"/><w:tblW w:w="0" w:type="auto"/><w:tblLook w:firstRow="1" w:lastRow="0" w:firstColumn="0" w:lastColumn="0" w:noHBand="0" w:noVBand="1"/></w:tblPr><w:tblGrid><w:gridCol w:w="3120"/><w:gridCol w:w="3120"/><w:gridCol w:w="3120"/></w:tblGrid>
      <w:tr><w:trPr><w:tblHeader/></w:trPr>${cell('Milestone')}${cell('Due')}${cell('Amount (USD)')}</w:tr>
      <w:tr>${cell('Mobilisation')}${cell('Signing + 30 days')}${cell('250,000')}</w:tr>
      <w:tr>${cell('First coal')}${cell('Q2 2027')}${cell('1,200,000')}</w:tr>
      <w:tr>${cell('Steady state')}${cell('Quarterly')}${cell('Per royalty schedule')}</w:tr></w:tbl>`,
    p(''),
    heading('3. Editable section', 'editable', 12),
    '<w:p><w:r><w:t xml:space="preserve">Counterparty notes: </w:t></w:r><w:permStart w:id="20" w:edGrp="everyone"/><w:r><w:t>[to be completed by the counterparty]</w:t></w:r><w:permEnd w:id="20"/></w:p>',
    num('Bulleted point one', 0, 2),
    num('Bulleted point two', 0, 2),
    heading('4. General', 'general', 13),
    ...Array.from({ length: 14 }, (_, i) => p(`${i + 1}. ${LOREM}`, '<w:jc w:val="both"/>')),
    `<w:p>${link('Back to the payment terms', 'payment')}${sep}${link('Back to top', 'top')}</w:p>`,
    heading('5. Site plan, figures and translations', 'figures', 15),
    // A floating text box (DrawingML, with Word's VML copy) that the next paragraphs wrap around.
    `<w:p><w:r><mc:AlternateContent><mc:Choice Requires="wps"><w:drawing><wp:anchor distT="0" distB="0" distL="114300" distR="114300" simplePos="0" relativeHeight="2" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1"><wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="column"><wp:align>right</wp:align></wp:positionH><wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV><wp:extent cx="2286000" cy="1028700"/><wp:wrapSquare wrapText="bothSides"/><wp:docPr id="20" name="Note box"/><a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><wps:wsp><wps:cNvSpPr txBox="1"/><wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="2286000" cy="1028700"/></a:xfrm><a:prstGeom prst="roundRect"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val="EAF1FB"/></a:solidFill><a:ln w="12700"><a:solidFill><a:srgbClr val="2E74B5"/></a:solidFill></a:ln></wps:spPr><wps:txbx><w:txbxContent><w:p><w:r><w:rPr><w:b/><w:color w:val="1F4E79"/></w:rPr><w:t>Note</w:t></w:r></w:p><w:p><w:r><w:rPr><w:sz w:val="18"/></w:rPr><w:t>Rehabilitation starts within six months of the end of mining in each pit.</w:t></w:r></w:p></w:txbxContent></wps:txbx><wps:bodyPr lIns="91440" tIns="45720" rIns="91440" bIns="45720" anchor="ctr"/></wps:wsp></a:graphicData></a:graphic></wp:anchor></w:drawing></mc:Choice><mc:Fallback><w:pict><v:rect style="width:180pt;height:81pt"><v:textbox><w:txbxContent><w:p><w:r><w:t>Note</w:t></w:r></w:p></w:txbxContent></v:textbox></v:rect></w:pict></mc:Fallback></mc:AlternateContent></w:r><w:r><w:t xml:space="preserve">The haul road runs from the pit to the crusher, then along the river to the port stockpile. </w:t></w:r><w:del w:id="31" w:author="Rina" w:date="2026-11-20T09:00:00Z"><w:r><w:delText xml:space="preserve">Trucks run day and night. </w:delText></w:r></w:del><w:ins w:id="32" w:author="Rina" w:date="2026-11-20T09:00:00Z"><w:r><w:t xml:space="preserve">Trucks run from 06:00 to 22:00 to limit noise near the village. </w:t></w:r></w:ins><w:r><w:t>${LOREM}</w:t></w:r></w:p>`,
    // Shapes: an arrow from the pit to the port.
    `<w:p><w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="1143000" cy="457200"/><wp:docPr id="21" name="Pit"/><a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><wps:wsp><wps:cNvSpPr/><wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1143000" cy="457200"/></a:xfrm><a:prstGeom prst="flowChartProcess"><a:avLst/></a:prstGeom></wps:spPr><wps:style><a:lnRef idx="2"><a:schemeClr val="accent1"><a:shade val="50000"/></a:schemeClr></a:lnRef><a:fillRef idx="1"><a:schemeClr val="accent1"/></a:fillRef><a:effectRef idx="0"><a:schemeClr val="accent1"/></a:effectRef><a:fontRef idx="minor"><a:schemeClr val="lt1"/></a:fontRef></wps:style><wps:txbx><w:txbxContent><w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:rPr><w:color w:val="FFFFFF"/></w:rPr><w:t>Pit</w:t></w:r></w:p></w:txbxContent></wps:txbx><wps:bodyPr anchor="ctr"/></wps:wsp></a:graphicData></a:graphic></wp:inline></w:drawing></w:r><w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="1371600" cy="457200"/><wp:docPr id="22" name="Haul"/><a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><wps:wsp><wps:cNvSpPr/><wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1371600" cy="457200"/></a:xfrm><a:prstGeom prst="rightArrow"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val="FFC000"/></a:solidFill></wps:spPr><wps:bodyPr/></wps:wsp></a:graphicData></a:graphic></wp:inline></w:drawing></w:r><w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="1143000" cy="457200"/><wp:docPr id="23" name="Port"/><a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><wps:wsp><wps:cNvSpPr/><wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1143000" cy="457200"/></a:xfrm><a:prstGeom prst="can"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val="70AD47"/></a:solidFill></wps:spPr><wps:txbx><w:txbxContent><w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:rPr><w:color w:val="FFFFFF"/></w:rPr><w:t>Port</w:t></w:r></w:p></w:txbxContent></wps:txbx><wps:bodyPr anchor="b"/></wps:wsp></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>`,
    p('Quarterly tonnage (thousand tonnes):'),
    `<w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="4572000" cy="2286000"/><wp:docPr id="24" name="Tonnage chart"/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" r:id="rIdChart"/></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>`,
    p('The royalty for each quarter follows the formula below, where r is the rate and P the price per tonne:'),
    `<w:p><m:oMathPara><m:oMath><m:r><m:t>R</m:t></m:r><m:r><m:t>=</m:t></m:r><m:nary><m:naryPr><m:chr m:val="∑"/></m:naryPr><m:sub><m:r><m:t>i=1</m:t></m:r></m:sub><m:sup><m:r><m:t>n</m:t></m:r></m:sup><m:e><m:sSub><m:e><m:r><m:t>q</m:t></m:r></m:e><m:sub><m:r><m:t>i</m:t></m:r></m:sub></m:sSub><m:r><m:t>r</m:t></m:r><m:sSub><m:e><m:r><m:t>P</m:t></m:r></m:e><m:sub><m:r><m:t>i</m:t></m:r></m:sub></m:sSub></m:e></m:nary><m:r><m:t>,  </m:t></m:r><m:r><m:t>x=</m:t></m:r><m:f><m:num><m:r><m:t>-b±</m:t></m:r><m:rad><m:radPr><m:degHide m:val="1"/></m:radPr><m:deg/><m:e><m:sSup><m:e><m:r><m:t>b</m:t></m:r></m:e><m:sup><m:r><m:t>2</m:t></m:r></m:sup></m:sSup><m:r><m:t>-4ac</m:t></m:r></m:e></m:rad></m:num><m:den><m:r><m:t>2a</m:t></m:r></m:den></m:f></m:oMath></m:oMathPara></w:p>`,
    p('Notice to residents, in Arabic and Hebrew:'),
    p('يبدأ النقل في الساعة السادسة صباحاً (06:00) وينتهي في العاشرة مساءً.', '<w:bidi/>'),
    p('ההובלה מתחילה בשעה 06:00 ומסתיימת בשעה 22:00.', '<w:bidi/>'),
    heading('6. Signatures', 'signatures', 14),
    '<w:sdt><w:sdtPr><w:alias w:val="Signature clause"/><w:tag w:val="docxkeun-lock"/><w:id w:val="502"/><w:lock w:val="sdtContentLocked"/></w:sdtPr><w:sdtContent>' +
      p('Signed for and on behalf of PT Contoh Tambang and the Contractor.') +
      p('This clause is locked: it cannot be edited or deleted.') +
      '</w:sdtContent></w:sdt>',
    `<w:p>${link('Back to top', 'top')}</w:p>`,
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
      ['rIdSite', 'hyperlink', 'https://example.com/docxkeun', true],
      ['rIdH', 'header', 'header1.xml'],
      ['rIdF', 'footer', 'footer1.xml'],
      ['rIdCx', 'http://schemas.microsoft.com/office/2011/relationships/commentsExtended', 'commentsExtended.xml'],
      ['rIdChart', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart', 'charts/chart1.xml'],
    ],
    parts: {
      'word/header1.xml': { data: header, contentType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml' },
      'word/footer1.xml': { data: footer, contentType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml' },
      'word/commentsExtended.xml': { data: commentsReply, contentType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.commentsExtended+xml' },
      'word/charts/chart1.xml': { data: CHART, contentType: 'application/vnd.openxmlformats-officedocument.drawingml.chart+xml' },
    },
  });
}
