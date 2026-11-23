/** Drawing markup for tests: inline frames, text box shapes and the two ways Word stores text boxes. */

export const EMU = 12700;

const NS = 'xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram"';

/** A w:drawing with an inline frame of `w` × `h` points around `graphic` (the a:graphicData content). */
export function inline(graphic: string, w: number, h: number, uri = 'http://schemas.microsoft.com/office/word/2010/wordprocessingShape'): string {
  return `<w:drawing><wp:inline ${NS}><wp:extent cx="${w * EMU}" cy="${h * EMU}"/><wp:docPr id="5" name="Shape"/><a:graphic><a:graphicData uri="${uri}">${graphic}</a:graphicData></a:graphic></wp:inline></w:drawing>`;
}

/** A 200 × 80 pt filled shape (wps:wsp) holding one paragraph of text. */
export const textBoxShape = (text: string, prst = 'roundRect') =>
  `<wps:wsp><wps:cNvSpPr/><wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${200 * EMU}" cy="${80 * EMU}"/></a:xfrm><a:prstGeom prst="${prst}"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val="DDEEFF"/></a:solidFill><a:ln w="${2 * EMU}"><a:solidFill><a:srgbClr val="2E74B5"/></a:solidFill><a:tailEnd type="triangle"/></a:ln></wps:spPr>` +
  `<wps:txbx><w:txbxContent><w:p><w:r><w:t>${text}</w:t></w:r></w:p></w:txbxContent></wps:txbx><wps:bodyPr anchor="ctr"/></wps:wsp>`;

/** A run holding a text box as Word saves it: DrawingML in mc:Choice, and an old VML copy saying "Old copy" in mc:Fallback. */
export const textBoxRun = (text: string) =>
  `<w:r><mc:AlternateContent><mc:Choice Requires="wps">${inline(textBoxShape(text), 200, 80)}</mc:Choice><mc:Fallback><w:pict><v:rect style="width:200pt;height:80pt"><v:textbox><w:txbxContent><w:p><w:r><w:t>Old copy</w:t></w:r></w:p></w:txbxContent></v:textbox></v:rect></w:pict></mc:Fallback></mc:AlternateContent></w:r>`;

/** A run holding an old VML-only text box (no DrawingML version). */
export const vmlTextBoxRun = (text: string) =>
  `<w:r><w:pict><v:rect style="width:150pt;height:60pt" fillcolor="#FFF2CC" strokecolor="#BF9000"><v:textbox><w:txbxContent><w:p><w:r><w:t>${text}</w:t></w:r></w:p></w:txbxContent></v:textbox></v:rect></w:pict></w:r>`;
