/** Builds small .docx packages from hand-written XML for tests. */
import { writeZip } from '../../src/zip/zip.js';
import { utf8Encode } from '../../src/util/text.js';

export const W_NS =
  'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ' +
  'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ' +
  'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" ' +
  'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ' +
  'xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture" ' +
  'xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" ' +
  'xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml" ' +
  'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="w14 w15"';

const DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n';

export interface FixtureOptions {
  body: string;
  styles?: string;
  numbering?: string;
  settings?: string;
  comments?: string;
  /** Extra document relationships: [id, typeSuffix or full type, target, external?]. */
  rels?: Array<[string, string, string, boolean?]>;
  /** Extra parts: name → XML string or bytes, with optional content type override. */
  parts?: Record<string, { data: string | Uint8Array; contentType?: string }>;
  sectPr?: string;
}

const RT = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const CT = 'application/vnd.openxmlformats-officedocument.wordprocessingml';

export function documentXml(body: string, sectPr?: string): string {
  const sect =
    sectPr ??
    '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr>';
  return `${DECL}<w:document ${W_NS}><w:body>${body}${sect}</w:body></w:document>`;
}

export async function buildDocx(options: FixtureOptions): Promise<Uint8Array> {
  const overrides: string[] = [`<Override PartName="/word/document.xml" ContentType="${CT}.document.main+xml"/>`];
  const rels: string[] = [];
  const files: Array<{ name: string; data: Uint8Array }> = [];
  const add = (name: string, data: string | Uint8Array) => files.push({ name, data: typeof data === 'string' ? utf8Encode(data) : data });

  add('word/document.xml', documentXml(options.body, options.sectPr));
  let n = 1;
  const addPart = (name: string, xml: string, kind: string) => {
    add(`word/${name}`, xml);
    overrides.push(`<Override PartName="/word/${name}" ContentType="${CT}.${kind}+xml"/>`);
    rels.push(`<Relationship Id="rIdFx${n++}" Type="${RT}/${kind}" Target="${name}"/>`);
  };
  if (options.styles) addPart('styles.xml', `${DECL}<w:styles ${W_NS}>${options.styles}</w:styles>`, 'styles');
  if (options.numbering) addPart('numbering.xml', `${DECL}<w:numbering ${W_NS}>${options.numbering}</w:numbering>`, 'numbering');
  if (options.settings !== undefined) addPart('settings.xml', `${DECL}<w:settings ${W_NS}>${options.settings}</w:settings>`, 'settings');
  if (options.comments) addPart('comments.xml', `${DECL}<w:comments ${W_NS}>${options.comments}</w:comments>`, 'comments');
  for (const [id, type, target, external] of options.rels ?? []) {
    const fullType = type.includes('://') ? type : `${RT}/${type}`;
    rels.push(`<Relationship Id="${id}" Type="${fullType}" Target="${target}"${external ? ' TargetMode="External"' : ''}/>`);
  }
  for (const [name, part] of Object.entries(options.parts ?? {})) {
    add(name, part.data);
    if (part.contentType) overrides.push(`<Override PartName="/${name}" ContentType="${part.contentType}"/>`);
  }

  const contentTypes =
    `${DECL}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
    '<Default Extension="xml" ContentType="application/xml"/>' +
    '<Default Extension="png" ContentType="image/png"/>' +
    `${overrides.join('')}</Types>`;
  const pkgRels =
    `${DECL}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
    `<Relationship Id="rId1" Type="${RT}/officeDocument" Target="word/document.xml"/></Relationships>`;
  const docRels = `${DECL}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${rels.join('')}</Relationships>`;

  return writeZip(
    [
      { name: '[Content_Types].xml', data: utf8Encode(contentTypes) },
      { name: '_rels/.rels', data: utf8Encode(pkgRels) },
      { name: 'word/_rels/document.xml.rels', data: utf8Encode(docRels) },
      ...files,
    ],
    { modified: new Date(2026, 0, 1) },
  );
}

/** Shorthand for a paragraph of plain runs. */
export function p(text: string, pPr = '', rPr = ''): string {
  return `<w:p>${pPr ? `<w:pPr>${pPr}</w:pPr>` : ''}<w:r>${rPr ? `<w:rPr>${rPr}</w:rPr>` : ''}<w:t xml:space="preserve">${text}</w:t></w:r></w:p>`;
}
