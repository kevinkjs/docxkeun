/**
 * Tracked changes (revisions), the way Word stores them:
 * - inserted runs inside `w:ins`, deleted runs inside `w:del` (their `w:t` become `w:delText`);
 * - a paragraph mark inserted or deleted as `w:pPr/w:rPr/w:ins|w:del`;
 * - formatting changes as `w:rPrChange` / `w:pPrChange` holding the old properties;
 * - moves as `w:moveFrom` / `w:moveTo` (treated as a deletion and an insertion).
 * This module records changes for the editor, lists them, and accepts or rejects them.
 */
import { attr, cloneElement, el, elements, findAll, firstChild, insertOrdered, removeChildren, textContent, type XmlElement, type XmlNode } from '../xml/index.js';

export interface RevisionAuthor {
  author: string;
  /** ISO date without milliseconds, as Word writes it. */
  date: string;
}

export type RevisionKind = 'insert' | 'delete' | 'move-from' | 'move-to' | 'format' | 'paragraph-format' | 'paragraph-insert' | 'paragraph-delete' | 'row-insert' | 'row-delete' | 'property';

export interface Revision {
  id: string;
  kind: RevisionKind;
  author: string;
  date?: string;
  /** Text the revision covers (inserted or deleted text; empty for formatting). */
  text: string;
  /** The paragraph it is in (or the first one it touches). */
  paragraph?: XmlElement;
  /** The element that carries the revision (w:ins, w:del, w:rPrChange, …). */
  el: XmlElement;
}

export interface RevisionFilter {
  ids?: string[];
  author?: string;
}

const RUN_REVISIONS = new Set(['w:ins', 'w:del', 'w:moveFrom', 'w:moveTo']);
const PROPERTY_CHANGES = new Set(['w:rPrChange', 'w:pPrChange', 'w:sectPrChange', 'w:tblPrChange', 'w:tblGridChange', 'w:trPrChange', 'w:tcPrChange', 'w:tblPrExChange', 'w:numberingChange']);
const RANGE_MARKERS = new Set(['w:moveFromRangeStart', 'w:moveFromRangeEnd', 'w:moveToRangeStart', 'w:moveToRangeEnd']);

export function revisionStamp(author: string, now = new Date()): RevisionAuthor {
  return { author, date: now.toISOString().replace(/\.\d{3}Z$/, 'Z') };
}

/** The next free revision id (ids are shared by all revision elements in a part). */
export function nextRevisionId(root: XmlElement): () => string {
  let max = 0;
  for (const name of [...RUN_REVISIONS, ...PROPERTY_CHANGES, 'w:cellIns', 'w:cellDel', 'w:cellMerge']) {
    for (const node of findAll(root, name)) max = Math.max(max, Number.parseInt(attr(node, 'w:id') ?? '', 10) || 0);
  }
  return () => String(++max);
}

function stampAttrs(id: string, who: RevisionAuthor): Record<string, string> {
  return { 'w:id': id, 'w:author': who.author, 'w:date': who.date };
}

/** Run properties of a new run must not carry another run's revision record. */
export function cleanRunProps(run: XmlElement): void {
  const rPr = firstChild(run, 'w:rPr');
  if (rPr) removeChildren(rPr, 'w:rPrChange');
}

/** Wrap runs (already in the tree, adjacent siblings) in one w:ins. */
export function wrapInserted(parent: XmlElement, runs: XmlElement[], id: string, who: RevisionAuthor): XmlElement {
  const first = parent.children.indexOf(runs[0]!);
  const ins = el('w:ins', stampAttrs(id, who), runs);
  parent.children.splice(first, runs.length, ins);
  return ins;
}

/** A new w:ins element holding runs (not yet in the tree). */
export function insertedRuns(runs: XmlElement[], id: string, who: RevisionAuthor): XmlElement {
  for (const r of runs) cleanRunProps(r);
  return el('w:ins', stampAttrs(id, who), runs);
}

function toDeletedText(run: XmlElement): void {
  const walk = (node: XmlElement): void => {
    for (const c of node.children) {
      if (c.type !== 'element') continue;
      if (c.name === 'w:t') c.name = 'w:delText';
      else if (c.name === 'w:instrText') c.name = 'w:delInstrText';
      walk(c);
    }
  };
  walk(run);
}

function toLiveText(node: XmlElement): void {
  for (const c of node.children) {
    if (c.type !== 'element') continue;
    if (c.name === 'w:delText') c.name = 'w:t';
    else if (c.name === 'w:delInstrText') c.name = 'w:instrText';
    toLiveText(c);
  }
}

/** The w:ins / w:del / w:moveFrom / w:moveTo a run sits directly in, if any. */
export function runRevision(run: XmlElement, parentOf: (n: XmlElement) => XmlElement | undefined): XmlElement | undefined {
  const parent = parentOf(run);
  return parent && RUN_REVISIONS.has(parent.name) ? parent : undefined;
}

/**
 * Mark a run as deleted by `who`. A run the same author inserted is removed outright (as Word does
 * when you delete your own insertion); a run that is already deleted is left alone.
 */
export function deleteRun(run: XmlElement, parentOf: (n: XmlElement) => XmlElement | undefined, nextId: () => string, who: RevisionAuthor): void {
  const parent = parentOf(run);
  if (!parent) return;
  if (parent.name === 'w:del' || parent.name === 'w:moveFrom') return;
  if ((parent.name === 'w:ins' || parent.name === 'w:moveTo') && attr(parent, 'w:author') === who.author) {
    parent.children.splice(parent.children.indexOf(run), 1);
    if (!elements(parent).length) {
      const grand = parentOf(parent);
      if (grand) grand.children.splice(grand.children.indexOf(parent), 1);
    }
    return;
  }
  toDeletedText(run);
  const del = el('w:del', stampAttrs(nextId(), who), [run]);
  parent.children.splice(parent.children.indexOf(run), 1, del);
}

function markRPr(p: XmlElement): XmlElement {
  let pPr = firstChild(p, 'w:pPr');
  if (!pPr) {
    pPr = el('w:pPr');
    p.children.unshift(pPr);
  }
  let rPr = firstChild(pPr, 'w:rPr');
  if (!rPr) {
    rPr = el('w:rPr');
    // In w:pPr, rPr comes after the paragraph properties and before sectPr / pPrChange.
    const after = elements(pPr).find((c) => c.name === 'w:sectPr' || c.name === 'w:pPrChange');
    if (after) pPr.children.splice(pPr.children.indexOf(after), 0, rPr);
    else pPr.children.push(rPr);
  }
  return rPr;
}

/** How a paragraph's mark (its end) is tracked: inserted, deleted, or not at all. */
export function paragraphMarkRevision(p: XmlElement): { kind: 'ins' | 'del'; el: XmlElement } | undefined {
  const rPr = firstChild(firstChild(p, 'w:pPr'), 'w:rPr');
  const ins = firstChild(rPr, 'w:ins') ?? firstChild(rPr, 'w:moveTo');
  if (ins) return { kind: 'ins', el: ins };
  const del = firstChild(rPr, 'w:del') ?? firstChild(rPr, 'w:moveFrom');
  if (del) return { kind: 'del', el: del };
  return undefined;
}

export function markParagraph(p: XmlElement, kind: 'ins' | 'del', id: string, who: RevisionAuthor): void {
  const rPr = markRPr(p);
  removeChildren(rPr, 'w:ins');
  removeChildren(rPr, 'w:del');
  // w:ins / w:del come first in a paragraph mark's rPr.
  rPr.children.unshift(el(kind === 'ins' ? 'w:ins' : 'w:del', stampAttrs(id, who)));
}

/** Record the run's current formatting before it changes (once: the oldest state is what counts). */
export function recordRunFormat(run: XmlElement, id: string, who: RevisionAuthor): void {
  let rPr = firstChild(run, 'w:rPr');
  if (!rPr) {
    rPr = el('w:rPr');
    run.children.unshift(rPr);
  }
  if (firstChild(rPr, 'w:rPrChange')) return;
  const old = el('w:rPr', {}, elements(rPr).map(cloneElement));
  rPr.children.push(el('w:rPrChange', stampAttrs(id, who), [old]));
}

/** Record a paragraph's current properties before they change. */
export function recordParagraphFormat(p: XmlElement, id: string, who: RevisionAuthor, order: readonly string[]): void {
  let pPr = firstChild(p, 'w:pPr');
  if (!pPr) {
    pPr = el('w:pPr');
    p.children.unshift(pPr);
  }
  if (firstChild(pPr, 'w:pPrChange')) return;
  const old = el(
    'w:pPr',
    {},
    elements(pPr)
      .filter((c) => c.name !== 'w:rPr' && c.name !== 'w:sectPr' && c.name !== 'w:pPrChange')
      .map(cloneElement),
  );
  insertOrdered(pPr, el('w:pPrChange', stampAttrs(id, who), [old]), order);
}

// ---- listing ---------------------------------------------------------------------------------------

function textOfRuns(node: XmlElement): string {
  let s = '';
  const walk = (n: XmlElement): void => {
    for (const c of n.children) {
      if (c.type !== 'element') continue;
      if (c.name === 'w:t' || c.name === 'w:delText') s += textContent(c);
      else if (c.name === 'w:tab') s += '\t';
      else if (c.name === 'w:br' || c.name === 'w:cr') s += '\n';
      else walk(c);
    }
  };
  walk(node);
  return s;
}

interface Parents {
  parent: Map<XmlElement, XmlElement>;
}

function parentsOf(root: XmlElement): Parents {
  const parent = new Map<XmlElement, XmlElement>();
  const walk = (n: XmlElement): void => {
    for (const c of n.children) {
      if (c.type !== 'element') continue;
      parent.set(c, n);
      walk(c);
    }
  };
  walk(root);
  return { parent };
}

function ancestor(p: Parents, node: XmlElement, name: string): XmlElement | undefined {
  for (let cur = p.parent.get(node); cur; cur = p.parent.get(cur)) if (cur.name === name) return cur;
  return undefined;
}

/** Every revision in a part, in document order. */
export function listRevisions(root: XmlElement): Revision[] {
  const p = parentsOf(root);
  const out: Revision[] = [];
  const walk = (node: XmlElement): void => {
    for (const c of node.children) {
      if (c.type !== 'element') continue;
      const id = attr(c, 'w:id') ?? '';
      const author = attr(c, 'w:author') ?? '';
      const date = attr(c, 'w:date');
      const base = { id, author, ...(date ? { date } : {}), el: c };
      const parentName = p.parent.get(c)?.name;
      if (RUN_REVISIONS.has(c.name)) {
        if (parentName === 'w:rPr') {
          // A paragraph mark (w:pPr/w:rPr) or a table row (w:trPr) revision.
          const para = ancestor(p, c, 'w:p');
          const kind = c.name === 'w:ins' || c.name === 'w:moveTo' ? 'paragraph-insert' : 'paragraph-delete';
          if (para && p.parent.get(p.parent.get(c)!)?.name === 'w:pPr') out.push({ ...base, kind, text: '', paragraph: para });
        } else if (parentName === 'w:trPr') {
          const row = ancestor(p, c, 'w:tr')!;
          out.push({ ...base, kind: c.name === 'w:ins' ? 'row-insert' : 'row-delete', text: textOfRuns(row), paragraph: findAll(row, 'w:p')[0] });
        } else {
          const kind = c.name === 'w:ins' ? 'insert' : c.name === 'w:del' ? 'delete' : c.name === 'w:moveFrom' ? 'move-from' : 'move-to';
          out.push({ ...base, kind, text: textOfRuns(c), ...(ancestor(p, c, 'w:p') ? { paragraph: ancestor(p, c, 'w:p')! } : {}) });
          walk(c);
        }
        continue;
      }
      if (c.name === 'w:p') {
        // The paragraph mark is the end of the paragraph: list its revision after the text's.
        const start = out.length;
        walk(c);
        const marks = out.slice(start).filter((r) => (r.kind === 'paragraph-insert' || r.kind === 'paragraph-delete') && r.paragraph === c);
        if (marks.length) {
          const rest = out.slice(start).filter((r) => !marks.includes(r));
          out.splice(start, out.length - start, ...rest, ...marks);
        }
        continue;
      }
      if (PROPERTY_CHANGES.has(c.name)) {
        const kind = c.name === 'w:rPrChange' ? 'format' : c.name === 'w:pPrChange' ? 'paragraph-format' : 'property';
        const run = ancestor(p, c, 'w:r');
        out.push({ ...base, kind, text: run ? textOfRuns(run) : '', ...(ancestor(p, c, 'w:p') ? { paragraph: ancestor(p, c, 'w:p')! } : {}) });
        continue;
      }
      walk(c);
    }
  };
  walk(root);
  return out;
}

// ---- accepting and rejecting -----------------------------------------------------------------------

function matches(r: Revision, filter: RevisionFilter | undefined): boolean {
  if (!filter) return true;
  if (filter.ids && !filter.ids.includes(r.id)) return false;
  if (filter.author !== undefined && r.author !== filter.author) return false;
  return true;
}

function unwrap(p: Parents, node: XmlElement): void {
  const parent = p.parent.get(node);
  if (!parent) return;
  parent.children.splice(parent.children.indexOf(node), 1, ...node.children);
  for (const c of node.children) if (c.type === 'element') p.parent.set(c, parent);
}

function remove(p: Parents, node: XmlElement): void {
  const parent = p.parent.get(node);
  if (parent) parent.children.splice(parent.children.indexOf(node), 1);
}

/** Join a paragraph into the one after it (its mark goes away); the next paragraph keeps its properties. */
function joinWithNext(p: Parents, para: XmlElement): void {
  const parent = p.parent.get(para);
  if (!parent) return;
  const i = parent.children.indexOf(para);
  let next: XmlElement | undefined;
  for (let j = i + 1; j < parent.children.length; j++) {
    const c = parent.children[j]!;
    if (c.type === 'element' && c.name === 'w:p') {
      next = c;
      break;
    }
    if (c.type === 'element' && (c.name === 'w:tbl' || c.name === 'w:sdt')) break;
  }
  const content = para.children.filter((c): c is XmlNode => !(c.type === 'element' && c.name === 'w:pPr'));
  if (!next) {
    // Nothing to join with (last paragraph of its container): just drop the mark's revision.
    const rPr = firstChild(firstChild(para, 'w:pPr'), 'w:rPr');
    if (rPr) for (const n of ['w:ins', 'w:del', 'w:moveTo', 'w:moveFrom']) removeChildren(rPr, n);
    return;
  }
  const nextPPr = firstChild(next, 'w:pPr');
  const at = nextPPr ? next.children.indexOf(nextPPr) + 1 : 0;
  next.children.splice(at, 0, ...content);
  for (const c of content) if (c.type === 'element') p.parent.set(c, next);
  parent.children.splice(i, 1);
}

function restoreProperties(p: Parents, change: XmlElement): void {
  const owner = p.parent.get(change);
  if (!owner) return;
  const old = elements(change)[0];
  const keep = owner.name === 'w:pPr' ? elements(owner).filter((c) => c.name === 'w:rPr' || c.name === 'w:sectPr') : [];
  owner.children = [...(old ? old.children : []), ...keep];
}

function apply(root: XmlElement, filter: RevisionFilter | undefined, accept: boolean): number {
  const p = parentsOf(root);
  const revisions = listRevisions(root).filter((r) => matches(r, filter));
  // Innermost and latest first, so nested revisions are resolved before their containers.
  for (const r of [...revisions].reverse()) {
    const node = r.el;
    switch (r.kind) {
      case 'insert':
      case 'move-to':
        if (accept) unwrap(p, node);
        else remove(p, node);
        break;
      case 'delete':
      case 'move-from':
        if (accept) remove(p, node);
        else {
          toLiveText(node);
          unwrap(p, node);
        }
        break;
      case 'format':
      case 'paragraph-format':
      case 'property':
        if (!accept) restoreProperties(p, node);
        remove(p, node);
        break;
      case 'paragraph-insert':
        remove(p, node);
        if (!accept && r.paragraph) joinWithNext(p, r.paragraph);
        break;
      case 'paragraph-delete':
        remove(p, node);
        if (accept && r.paragraph) joinWithNext(p, r.paragraph);
        break;
      case 'row-insert': {
        const row = p.parent.get(p.parent.get(node)!);
        remove(p, node);
        if (!accept && row) remove(p, row);
        break;
      }
      case 'row-delete': {
        const row = p.parent.get(p.parent.get(node)!);
        remove(p, node);
        if (accept && row) remove(p, row);
        break;
      }
    }
  }
  // Move range markers mean nothing once their moves are resolved.
  if (revisions.some((r) => r.kind === 'move-from' || r.kind === 'move-to')) {
    for (const name of RANGE_MARKERS) for (const m of findAll(root, name)) remove(parentsOf(root), m);
  }
  return revisions.length;
}

export function acceptRevisions(root: XmlElement, filter?: RevisionFilter): number {
  return apply(root, filter, true);
}

export function rejectRevisions(root: XmlElement, filter?: RevisionFilter): number {
  return apply(root, filter, false);
}
