/** Creating optional Word parts (comments, people, numbering...) the first time they are needed. */
import { randomBytes } from '../crypto/index.js';
import type { DocxDocument, WordParts } from '../model/document.js';
import { CONTENT_TYPE, REL } from '../opc/index.js';
import { el, NS, type XmlDocument } from '../xml/index.js';

interface PartSpec {
  name: string;
  rel: string;
  contentType: string;
  root: string;
}

export const PART_SPECS: Partial<Record<keyof WordParts, PartSpec>> = {
  comments: { name: 'word/comments.xml', rel: REL.comments, contentType: CONTENT_TYPE.comments, root: 'w:comments' },
  commentsExtended: { name: 'word/commentsExtended.xml', rel: REL.commentsExtended, contentType: CONTENT_TYPE.commentsExtended, root: 'w15:commentsEx' },
  commentsIds: { name: 'word/commentsIds.xml', rel: REL.commentsIds, contentType: CONTENT_TYPE.commentsIds, root: 'w16cid:commentsIds' },
  commentsExtensible: { name: 'word/commentsExtensible.xml', rel: REL.commentsExtensible, contentType: CONTENT_TYPE.commentsExtensible, root: 'w16cex:commentsExtensible' },
  people: { name: 'word/people.xml', rel: REL.people, contentType: CONTENT_TYPE.people, root: 'w15:people' },
  settings: { name: 'word/settings.xml', rel: REL.settings, contentType: CONTENT_TYPE.settings, root: 'w:settings' },
  styles: { name: 'word/styles.xml', rel: REL.styles, contentType: CONTENT_TYPE.styles, root: 'w:styles' },
};

/** Namespace declarations put on the roots of parts docxkeun creates. */
export const ROOT_NAMESPACES: Record<string, string> = {
  'xmlns:mc': NS.mc,
  'xmlns:r': NS.r,
  'xmlns:w': NS.w,
  'xmlns:w14': NS.w14,
  'xmlns:w15': NS.w15,
  'xmlns:w16cid': NS.w16cid,
  'xmlns:w16cex': NS.w16cex,
  'mc:Ignorable': 'w14 w15 w16cid w16cex',
};

/** Make sure a root element declares the namespaces docxkeun writes into it. */
export function ensureNamespaces(doc: XmlDocument): void {
  const attrs = doc.root.attrs;
  for (const [k, v] of Object.entries(ROOT_NAMESPACES)) {
    if (k === 'mc:Ignorable') continue;
    if (attrs[k] === undefined) attrs[k] = v;
  }
  const ignorable = new Set((attrs['mc:Ignorable'] ?? '').split(/\s+/).filter(Boolean));
  for (const p of ['w14', 'w15', 'w16cid', 'w16cex']) ignorable.add(p);
  attrs['mc:Ignorable'] = [...ignorable].join(' ');
}

/** Existing part XML, or a new empty part wired into the package. */
export async function ensurePart(doc: DocxDocument, kind: keyof WordParts): Promise<{ name: string; xml: XmlDocument }> {
  const existing = doc.parts[kind];
  if (existing) {
    const xml = doc.xmlSync(existing) ?? (await doc.pkg.getXml(existing));
    return { name: existing, xml };
  }
  const spec = PART_SPECS[kind];
  if (!spec) throw new Error(`docxkeun cannot create a ${kind} part`);
  let name = spec.name;
  for (let n = 2; doc.pkg.has(name); n++) name = spec.name.replace(/\.xml$/, `${n}.xml`);
  const xml: XmlDocument = { declaration: null, prolog: [], root: el(spec.root, { ...ROOT_NAMESPACES }), epilog: [] };
  doc.pkg.setXml(name, xml, spec.contentType);
  await doc.pkg.addRelationship(doc.mainPart, spec.rel, name);
  await doc.refreshRels(doc.mainPart);
  doc.parts[kind] = name;
  return { name, xml };
}

/** 8 hex digits below 0x80000000, as Word requires for w14:paraId and durable ids. */
export function randomHexId(): string {
  const b = randomBytes(4);
  b[0]! &= 0x7f;
  if (b[0] === 0 && b[1] === 0 && b[2] === 0 && b[3] === 0) b[3] = 1;
  return [...b].map((x) => x.toString(16).padStart(2, '0')).join('').toUpperCase();
}
