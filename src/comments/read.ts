/**
 * Comments: w:comment entries in comments.xml plus Word's modern-comment parts
 * (commentsExtended: threads and resolved state, commentsIds: durable ids,
 * commentsExtensible: UTC dates, people: author identities).
 */
import type { Block } from '../model/content.js';
import { blocksToText } from '../model/content.js';
import type { DocxDocument } from '../model/document.js';
import { attr, elements, findAll, type XmlElement } from '../xml/index.js';

export interface CommentInfo {
  id: string;
  author: string;
  initials?: string;
  date?: string;
  /** UTC date from commentsExtensible, when present. */
  dateUtc?: string;
  text: string;
  blocks: Block[];
  /** w14:paraId of the comment's last paragraph (the key used by the modern-comment parts). */
  paraId?: string;
  durableId?: string;
  /** Id of the comment this one replies to. */
  parentId?: string;
  resolved: boolean;
  el: XmlElement;
}

export function readComments(doc: DocxDocument): CommentInfo[] {
  const part = doc.parts.comments;
  const xml = doc.xmlSync(part);
  if (!part || !xml) return [];
  const ext = doc.xmlSync(doc.parts.commentsExtended);
  const ids = doc.xmlSync(doc.parts.commentsIds);
  const extensible = doc.xmlSync(doc.parts.commentsExtensible);

  const exByPara = new Map<string, { parent?: string; done: boolean }>();
  if (ext) {
    for (const e of findAll(ext.root, 'w15:commentEx')) {
      const paraId = attr(e, 'w15:paraId');
      if (paraId) exByPara.set(paraId.toUpperCase(), { parent: attr(e, 'w15:paraIdParent')?.toUpperCase(), done: attr(e, 'w15:done') === '1' });
    }
  }
  const durableByPara = new Map<string, string>();
  if (ids) for (const e of findAll(ids.root, 'w16cid:commentId')) {
    const paraId = attr(e, 'w16cid:paraId');
    const durable = attr(e, 'w16cid:durableId');
    if (paraId && durable) durableByPara.set(paraId.toUpperCase(), durable);
  }
  const utcByDurable = new Map<string, string>();
  if (extensible) for (const e of findAll(extensible.root, 'w16cex:commentExtensible')) {
    const durable = attr(e, 'w16cex:durableId');
    const utc = attr(e, 'w16cex:dateUtc');
    if (durable && utc) utcByDurable.set(durable, utc);
  }

  const comments: CommentInfo[] = [];
  for (const c of elements(xml.root, 'w:comment')) {
    const id = attr(c, 'w:id');
    if (id === undefined) continue;
    const paragraphs = elements(c, 'w:p');
    const paraId = attr(paragraphs[paragraphs.length - 1], 'w14:paraId')?.toUpperCase();
    const blocks = doc.readBlocks(part, c);
    const durableId = paraId ? durableByPara.get(paraId) : undefined;
    comments.push({
      id,
      author: attr(c, 'w:author') ?? '',
      initials: attr(c, 'w:initials'),
      date: attr(c, 'w:date'),
      dateUtc: durableId ? utcByDurable.get(durableId) : undefined,
      text: blocksToText(blocks),
      blocks,
      paraId,
      durableId,
      resolved: paraId ? (exByPara.get(paraId)?.done ?? false) : false,
      el: c,
    });
  }
  // Resolve parent paraIds to comment ids.
  const byPara = new Map(comments.filter((c) => c.paraId).map((c) => [c.paraId!, c.id]));
  for (const c of comments) {
    const parentPara = c.paraId ? exByPara.get(c.paraId)?.parent : undefined;
    if (parentPara) c.parentId = byPara.get(parentPara);
  }
  return comments;
}
