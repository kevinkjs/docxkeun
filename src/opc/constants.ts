/** Relationship types (matched by their last path segment, so Strict and Transitional both work). */
export const REL = {
  officeDocument: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument',
  coreProperties: 'http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties',
  extendedProperties: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties',
  styles: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles',
  numbering: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering',
  settings: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings',
  webSettings: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/webSettings',
  fontTable: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/fontTable',
  theme: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme',
  footnotes: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes',
  endnotes: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes',
  header: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/header',
  footer: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer',
  comments: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments',
  commentsExtended: 'http://schemas.microsoft.com/office/2011/relationships/commentsExtended',
  commentsIds: 'http://schemas.microsoft.com/office/2016/09/relationships/commentsIds',
  commentsExtensible: 'http://schemas.microsoft.com/office/2018/08/relationships/commentsExtensible',
  people: 'http://schemas.microsoft.com/office/2011/relationships/people',
  hyperlink: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink',
  image: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image',
  font: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/font',
} as const;

export type RelKind = keyof typeof REL;

const WML = 'application/vnd.openxmlformats-officedocument.wordprocessingml';

export const CONTENT_TYPE = {
  relationships: 'application/vnd.openxmlformats-package.relationships+xml',
  xml: 'application/xml',
  document: `${WML}.document.main+xml`,
  template: `${WML}.template.main+xml`,
  documentMacro: 'application/vnd.ms-word.document.macroEnabled.main+xml',
  styles: `${WML}.styles+xml`,
  numbering: `${WML}.numbering+xml`,
  settings: `${WML}.settings+xml`,
  webSettings: `${WML}.webSettings+xml`,
  fontTable: `${WML}.fontTable+xml`,
  footnotes: `${WML}.footnotes+xml`,
  endnotes: `${WML}.endnotes+xml`,
  header: `${WML}.header+xml`,
  footer: `${WML}.footer+xml`,
  comments: `${WML}.comments+xml`,
  commentsExtended: `${WML}.commentsExtended+xml`,
  commentsIds: `${WML}.commentsIds+xml`,
  commentsExtensible: `${WML}.commentsExtensible+xml`,
  people: `${WML}.people+xml`,
  theme: 'application/vnd.openxmlformats-officedocument.theme+xml',
  coreProperties: 'application/vnd.openxmlformats-package.core-properties+xml',
  extendedProperties: 'application/vnd.openxmlformats-officedocument.extended-properties+xml',
  obfuscatedFont: 'application/vnd.openxmlformats-officedocument.obfuscatedFont',
} as const;

export const IMAGE_CONTENT_TYPES: Record<string, string> = {
  png: 'image/png',
  jpg: 'image/jpeg',
  jpeg: 'image/jpeg',
  gif: 'image/gif',
  bmp: 'image/bmp',
  tif: 'image/tiff',
  tiff: 'image/tiff',
  svg: 'image/svg+xml',
  emf: 'image/x-emf',
  wmf: 'image/x-wmf',
  webp: 'image/webp',
};

export function relTypeName(type: string): string {
  const i = type.lastIndexOf('/');
  return i < 0 ? type : type.slice(i + 1);
}

export function isRelType(type: string, kind: RelKind): boolean {
  return relTypeName(type) === relTypeName(REL[kind]);
}
