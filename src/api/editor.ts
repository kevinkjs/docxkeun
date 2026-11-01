/**
 * DocxEditor: the code API for changing a document. Every change edits the XML directly (so
 * everything docxkit does not understand is preserved), passes through the access guard, and
 * notifies listeners such as the viewer.
 */
import { REL } from '../opc/index.js';
import type { DocParagraph, Block, Inline } from '../model/content.js';
import type { DocxDocument } from '../model/document.js';
import { attr, el, elements, ensureChild, findAll, firstChild, insertOrdered, removeChildren, textContent, type XmlElement } from '../xml/index.js';
import { ensureNamespaces, ensurePart, randomHexId } from './parts.js';
import {
  between,
  EditError,
  indexOfChild,
  isolate,
  PPR_ORDER,
  replaceMark,
  RPR_ORDER,
  textNode,
  TreeIndex,
  type DocPosition,
  type DocRange,
} from './positions.js';

export type EditAction = 'edit' | 'format' | 'comment' | 'reply' | 'resolveComment' | 'deleteComment' | 'link' | 'bookmark' | 'protect';

/** Checks every edit (see docxkit/security's AccessPolicy). Throw to refuse. */
export interface EditGuard {
  assert(action: EditAction, range?: DocRange, detail?: { commentAuthor?: string }): void;
}

export interface TextFormat {
  bold?: boolean;
  italic?: boolean;
  underline?: boolean | 'single' | 'double' | 'dotted' | 'dash' | 'wave' | 'thick';
  strike?: boolean;
  color?: string;
  highlight?: string;
  /** Points. */
  size?: number;
  font?: string;
  style?: string;
  vertAlign?: 'superscript' | 'subscript' | 'baseline';
  caps?: boolean;
  smallCaps?: boolean;
  hidden?: boolean;
}

export interface CommentInput {
  author: string;
  text: string;
  initials?: string;
  date?: Date;
}

export interface Match extends DocRange {
  text: string;
  /** Capture groups for regular-expression searches. */
  groups: string[];
  paragraph: XmlElement;
}

interface TextSegment {
  start: number;
  end: number;
  node: XmlElement;
}

/** Concatenated text of a paragraph and where each character lives in the XML. */
interface ParagraphText {
  paragraph: XmlElement;
  text: string;
  segments: TextSegment[];
}

function paragraphsOf(blocks: Block[], out: DocParagraph[] = []): DocParagraph[] {
  for (const b of blocks) {
    if (b.kind === 'paragraph') out.push(b);
    else for (const row of b.rows) for (const cell of row.cells) paragraphsOf(cell.blocks, out);
  }
  return out;
}

function paragraphText(para: DocParagraph): ParagraphText {
  let text = '';
  const segments: TextSegment[] = [];
  for (const inline of para.inlines as Inline[]) {
    if (inline.kind === 'text') {
      segments.push({ start: text.length, end: text.length + inline.text.length, node: inline.node });
      text += inline.text;
    } else if (inline.kind === 'tab') text += '\t';
    else if (inline.kind === 'break' && inline.type === 'line') text += '\n';
  }
  return { paragraph: para.el, text, segments };
}

function positionAt(pt: ParagraphText, k: number, side: 'start' | 'end'): DocPosition | undefined {
  for (const s of pt.segments) {
    const inside = side === 'start' ? k >= s.start && k < s.end : k > s.start && k <= s.end;
    if (inside) {
      if (s.node.name === 'w:t') return { node: s.node, offset: k - s.start };
      return { node: s.node, offset: side === 'start' ? 0 : 1 };
    }
  }
  return undefined;
}

/** Run children for a string: text, tabs and line breaks. */
export function runContent(text: string): XmlElement[] {
  const out: XmlElement[] = [];
  for (const part of text.split(/(\t|\r?\n)/)) {
    if (part === '') continue;
    if (part === '\t') out.push(el('w:tab'));
    else if (part === '\n' || part === '\r\n') out.push(el('w:br'));
    else out.push(textNode(part));
  }
  return out;
}

/** Apply TextFormat to a w:rPr element (created when missing). */
export function applyFormat(run: XmlElement, format: TextFormat): void {
  const rPr = firstChild(run, 'w:rPr') ?? (() => {
    const created = el('w:rPr');
    run.children.unshift(created);
    return created;
  })();
  const toggle = (name: string, value: boolean | undefined) => {
    if (value === undefined) return;
    removeChildren(rPr, name);
    insertOrdered(rPr, el(name, value ? {} : { 'w:val': '0' }), RPR_ORDER);
  };
  const set = (name: string, attrs: Record<string, string> | undefined) => {
    removeChildren(rPr, name);
    if (attrs) insertOrdered(rPr, el(name, attrs), RPR_ORDER);
  };
  toggle('w:b', format.bold);
  toggle('w:bCs', format.bold);
  toggle('w:i', format.italic);
  toggle('w:iCs', format.italic);
  toggle('w:strike', format.strike);
  toggle('w:caps', format.caps);
  toggle('w:smallCaps', format.smallCaps);
  toggle('w:vanish', format.hidden);
  if (format.underline !== undefined) set('w:u', { 'w:val': format.underline === true ? 'single' : format.underline === false ? 'none' : format.underline });
  if (format.color !== undefined) set('w:color', { 'w:val': format.color.replace(/^#/, '').toUpperCase() });
  if (format.highlight !== undefined) set('w:highlight', format.highlight === 'none' ? undefined : { 'w:val': format.highlight });
  if (format.size !== undefined) {
    set('w:sz', { 'w:val': String(Math.round(format.size * 2)) });
    set('w:szCs', { 'w:val': String(Math.round(format.size * 2)) });
  }
  if (format.font !== undefined) set('w:rFonts', { 'w:ascii': format.font, 'w:hAnsi': format.font, 'w:cs': format.font, 'w:eastAsia': format.font });
  if (format.style !== undefined) set('w:rStyle', { 'w:val': format.style });
  if (format.vertAlign !== undefined) set('w:vertAlign', { 'w:val': format.vertAlign });
  if (rPr.children.length === 0) run.children.splice(run.children.indexOf(rPr), 1);
}

function newRun(text: string, template?: XmlElement, format?: TextFormat): XmlElement {
  const rPr = template ? firstChild(template, 'w:rPr') : undefined;
  const run = el('w:r', {}, [...(rPr ? [structuredClone(rPr)] : []), ...runContent(text)]);
  if (format) applyFormat(run, format);
  return run;
}

const BOOKMARK_NAME = /^[\p{L}_][\p{L}\p{N}_]{0,39}$/u;

export class DocxEditor {
  constructor(readonly doc: DocxDocument, readonly guard?: EditGuard) {}

  private get root(): XmlElement {
    return this.doc.documentXml.root;
  }

  private get body(): XmlElement {
    const body = firstChild(this.root, 'w:body');
    if (!body) throw new EditError('document has no body');
    return body;
  }

  private index(): TreeIndex {
    return new TreeIndex(this.root);
  }

  private changed(part = this.doc.mainPart): void {
    this.doc.touch(part);
  }

  private check(action: EditAction, range?: DocRange, detail?: { commentAuthor?: string }): void {
    this.guard?.assert(action, range, detail);
  }

  // ---- reading ------------------------------------------------------------------------------

  /** Convert a viewer selection (inline positions) into an editor range. */
  rangeFromSelection(sel: { start: { inline: { node: XmlElement }; offset: number }; end: { inline: { node: XmlElement }; offset: number } }): DocRange {
    return { start: { node: sel.start.inline.node, offset: sel.start.offset }, end: { node: sel.end.inline.node, offset: sel.end.offset } };
  }

  /** Every paragraph of the body, including those in tables, in document order. */
  paragraphs(): DocParagraph[] {
    return paragraphsOf(this.doc.body().blocks);
  }

  /** Range covering a paragraph's text (undefined for paragraphs without text). */
  paragraphRange(paragraph: XmlElement | DocParagraph): DocRange | undefined {
    const para = 'kind' in paragraph ? paragraph : this.paragraphs().find((p) => p.el === paragraph);
    if (!para) return undefined;
    const pt = paragraphText(para);
    const start = positionAt(pt, 0, 'start');
    const end = positionAt(pt, pt.text.length, 'end');
    return start && end ? { start, end } : undefined;
  }

  /** Find text (or a regular expression) within paragraphs, across run boundaries. */
  find(query: string | RegExp): Match[] {
    const out: Match[] = [];
    const regex = typeof query === 'string' ? new RegExp(query.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'), 'g') : new RegExp(query.source, query.flags.includes('g') ? query.flags : query.flags + 'g');
    for (const para of this.paragraphs()) {
      const pt = paragraphText(para);
      if (!pt.text) continue;
      regex.lastIndex = 0;
      for (let m = regex.exec(pt.text); m; m = regex.exec(pt.text)) {
        if (m[0].length === 0) {
          regex.lastIndex++;
          continue;
        }
        const start = positionAt(pt, m.index, 'start');
        const end = positionAt(pt, m.index + m[0].length, 'end');
        if (start && end) out.push({ start, end, text: m[0], groups: m.slice(1).map((g) => g ?? ''), paragraph: para.el });
      }
    }
    return out;
  }

  /** Text covered by a range. */
  textOf(range: DocRange): string {
    for (const para of this.paragraphs()) {
      const pt = paragraphText(para);
      const s = pt.segments.find((seg) => seg.node === range.start.node);
      const e = pt.segments.find((seg) => seg.node === range.end.node);
      if (s && e) {
        const a = s.start + (s.node.name === 'w:t' ? range.start.offset : range.start.offset ? s.end - s.start : 0);
        const b = e.start + (e.node.name === 'w:t' ? range.end.offset : range.end.offset ? e.end - e.start : 0);
        return pt.text.slice(a, b);
      }
    }
    return '';
  }

  // ---- text edits ---------------------------------------------------------------------------

  /** Insert text at a position, formatted like the run it lands in (plus optional format). */
  insertText(at: DocPosition, text: string, format?: TextFormat): void {
    this.check('edit', { start: at, end: at });
    const index = this.index();
    const template = index.parentOf(at.node);
    const { startMark, endMark } = isolate(index, { start: at, end: at });
    const idx = this.index();
    replaceMark(idx, startMark, [newRun(text, template, format)]);
    replaceMark(idx, endMark, []);
    this.changed();
  }

  /** Append text at the end of a paragraph. */
  appendText(paragraph: XmlElement, text: string, format?: TextFormat): XmlElement {
    const range = this.paragraphRange(paragraph);
    this.check('edit', range);
    const runs = elements(paragraph, 'w:r');
    const run = newRun(text, runs[runs.length - 1], format);
    paragraph.children.push(run);
    this.changed();
    return run;
  }

  /** Delete a range (it may span paragraphs, but not table cells). */
  deleteRange(range: DocRange): void {
    this.check('edit', range);
    this.removeRange(range);
    this.changed();
  }

  /** Replace a range with text formatted like the range's first run. */
  replaceRange(range: DocRange, text: string, format?: TextFormat): void {
    this.check('edit', range);
    const index = this.index();
    const template = index.parentOf(range.start.node);
    const { startMark, endMark, idx } = this.cut(range);
    replaceMark(idx, startMark, text || format ? [newRun(text, template, format)] : []);
    replaceMark(idx, endMark, []);
    this.changed();
  }

  /** Replace every match; returns how many were replaced. */
  replaceAll(query: string | RegExp, replacement: string | ((m: Match) => string)): number {
    const matches = this.find(query);
    // Back to front, so earlier positions stay valid.
    for (const m of matches.reverse()) {
      const value = typeof replacement === 'string' ? replacement.replace(/\$(\d)/g, (_, d: string) => m.groups[Number(d) - 1] ?? '') : replacement(m);
      this.replaceRange(m, value);
    }
    return matches.length;
  }

  /**
   * Fill `{{name}}` placeholders (and content controls whose tag or title is a key).
   * Returns the keys that were found.
   */
  fillTemplate(data: Record<string, string | number | boolean | null | undefined>, pattern = /\{\{\s*([\w.-]+)\s*\}\}/g): string[] {
    const used = new Set<string>();
    this.replaceAll(pattern, (m) => {
      const key = m.groups[0] ?? '';
      if (!(key in data)) return m.text;
      used.add(key);
      return String(data[key] ?? '');
    });
    for (const key of this.fillContentControls(data)) used.add(key);
    return [...used];
  }

  /** Set the text of content controls (w:sdt) by tag or title. */
  fillContentControls(data: Record<string, string | number | boolean | null | undefined>): string[] {
    const used: string[] = [];
    for (const sdt of findAll(this.body, 'w:sdt')) {
      const sdtPr = firstChild(sdt, 'w:sdtPr');
      const key = attr(firstChild(sdtPr, 'w:tag'), 'w:val') ?? attr(firstChild(sdtPr, 'w:alias'), 'w:val');
      if (key === undefined || !(key in data)) continue;
      const content = firstChild(sdt, 'w:sdtContent');
      if (!content) continue;
      this.check('edit');
      const value = String(data[key] ?? '');
      const firstRun = findAll(content, 'w:r')[0];
      if (sdtPr) removeChildren(sdtPr, 'w:showingPlcHdr');
      const firstPara = firstChild(content, 'w:p');
      if (firstPara) {
        const pPr = firstChild(firstPara, 'w:pPr');
        content.children = [el('w:p', {}, [...(pPr ? [pPr] : []), newRun(value, firstRun)])];
      } else {
        content.children = [newRun(value, firstRun)];
      }
      // Placeholder text usually uses the "PlaceholderText" style; drop it from the new run.
      for (const run of findAll(content, 'w:r')) {
        const rStyle = firstChild(firstChild(run, 'w:rPr'), 'w:rStyle');
        if (attr(rStyle, 'w:val') === 'PlaceholderText') removeChildren(firstChild(run, 'w:rPr')!, 'w:rStyle');
      }
      used.push(key);
    }
    if (used.length) this.changed();
    return used;
  }

  /** Apply character formatting to a range. */
  formatRange(range: DocRange, format: TextFormat): void {
    this.check('format', range);
    const index = this.index();
    const { startMark, endMark } = isolate(index, range);
    const idx = this.index();
    for (const run of between(idx, startMark, endMark, 'w:r')) applyFormat(run, format);
    replaceMark(idx, startMark, []);
    replaceMark(idx, endMark, []);
    this.changed();
  }

  /** Remove the content of a range, returning the sentinels at its (now adjacent) ends. */
  private cut(range: DocRange): { startMark: XmlElement; endMark: XmlElement; idx: TreeIndex } {
    const index = this.index();
    const { startMark, endMark } = isolate(index, range);
    let idx = this.index();
    const startCell = idx.ancestor(startMark, 'w:tc');
    const endCell = idx.ancestor(endMark, 'w:tc');
    if (startCell !== endCell) {
      replaceMark(idx, startMark, []);
      replaceMark(idx, endMark, []);
      throw new EditError('The range crosses table cells');
    }
    const startPara = idx.ancestor(startMark, 'w:p')!;
    const endPara = idx.ancestor(endMark, 'w:p')!;
    // Remove runs between the marks, then whole paragraphs and tables in between.
    for (const run of between(idx, startMark, endMark, 'w:r')) {
      const parent = idx.parent.get(run);
      if (parent) parent.children.splice(parent.children.indexOf(run), 1);
    }
    if (startPara !== endPara) {
      idx = this.index();
      for (const node of between(idx, startMark, endMark)) {
        if ((node.name === 'w:p' || node.name === 'w:tbl') && node !== endPara && !idx.ancestor(node, 'w:p')) {
          const parent = idx.parent.get(node);
          if (parent && parent.children.includes(node) && !isAncestor(idx, node, endMark)) parent.children.splice(parent.children.indexOf(node), 1);
        }
      }
      // Merge the end paragraph's remaining content into the start paragraph.
      idx = this.index();
      const endContainer = idx.parentOf(endMark);
      const tail = endContainer.children.splice(indexOfChild(endContainer, endMark));
      startPara.children.push(...tail.filter((n) => !(n.type === 'element' && n.name === 'w:pPr')));
      const endParent = idx.parent.get(endPara);
      if (endParent) endParent.children.splice(endParent.children.indexOf(endPara), 1);
    }
    return { startMark, endMark, idx: this.index() };
  }

  private removeRange(range: DocRange): void {
    const { startMark, endMark, idx } = this.cut(range);
    replaceMark(idx, startMark, []);
    replaceMark(idx, endMark, []);
  }

  // ---- paragraphs ---------------------------------------------------------------------------

  /** Insert a paragraph before or after a block (or at the start/end of the body). */
  insertParagraph(where: { before?: XmlElement; after?: XmlElement } | 'start' | 'end', text = '', options: { style?: string; format?: TextFormat } = {}): XmlElement {
    this.check('edit');
    const pPr = options.style ? [el('w:pPr', {}, [el('w:pStyle', { 'w:val': options.style })])] : [];
    const para = el('w:p', {}, [...pPr, ...(text ? [newRun(text, undefined, options.format)] : [])]);
    const body = this.body;
    if (where === 'start') body.children.unshift(para);
    else if (where === 'end') {
      const sect = firstChild(body, 'w:sectPr');
      body.children.splice(sect ? body.children.indexOf(sect) : body.children.length, 0, para);
    } else {
      const ref = where.before ?? where.after;
      if (!ref) throw new EditError('insertParagraph needs a reference block');
      const idx = this.index();
      const parent = idx.parentOf(ref);
      parent.children.splice(indexOfChild(parent, ref) + (where.after ? 1 : 0), 0, para);
    }
    this.changed();
    return para;
  }

  deleteParagraph(paragraph: XmlElement): void {
    this.check('edit', this.paragraphRange(paragraph));
    const idx = this.index();
    const parent = idx.parentOf(paragraph);
    // A table cell must keep at least one paragraph.
    if (parent.name === 'w:tc' && elements(parent, 'w:p').length === 1) {
      paragraph.children = paragraph.children.filter((c) => c.type === 'element' && c.name === 'w:pPr');
    } else parent.children.splice(indexOfChild(parent, paragraph), 1);
    this.changed();
  }

  setParagraphStyle(paragraph: XmlElement, styleId: string): void {
    this.check('format', this.paragraphRange(paragraph));
    const pPr = ensureChild(paragraph, 'w:pPr');
    if (paragraph.children[0] !== pPr) {
      paragraph.children.splice(paragraph.children.indexOf(pPr), 1);
      paragraph.children.unshift(pPr);
    }
    removeChildren(pPr, 'w:pStyle');
    insertOrdered(pPr, el('w:pStyle', { 'w:val': styleId }), PPR_ORDER);
    this.changed();
  }

  // ---- hyperlinks and bookmarks -------------------------------------------------------------

  /** Turn a range (within one paragraph) into a hyperlink to a URL or a bookmark. */
  async addHyperlink(range: DocRange, target: { url: string } | { bookmark: string }, tooltip?: string): Promise<XmlElement> {
    this.check('link', range);
    if ('url' in target && !/^(https?|mailto|tel|ftp):/i.test(target.url)) throw new EditError(`Refusing link to "${target.url}": unsupported scheme`);
    const attrs: Record<string, string> = {};
    if ('url' in target) {
      attrs['r:id'] = await this.doc.pkg.addRelationship(this.doc.mainPart, REL.hyperlink, target.url, true);
      await this.doc.refreshRels(this.doc.mainPart);
    } else attrs['w:anchor'] = target.bookmark;
    if (tooltip) attrs['w:tooltip'] = tooltip;
    attrs['w:history'] = '1';
    await this.ensureCharacterStyle('Hyperlink', 'Hyperlink', '<w:color w:val="0563C1" w:themeColor="hyperlink"/><w:u w:val="single"/>');

    const index = this.index();
    const { startMark, endMark } = isolate(index, range);
    const idx = this.index();
    const container = idx.parentOf(startMark);
    if (idx.parentOf(endMark) !== container) {
      replaceMark(idx, startMark, []);
      replaceMark(idx, endMark, []);
      throw new EditError('A hyperlink must start and end in the same paragraph (and not cut across another link or field)');
    }
    const a = indexOfChild(container, startMark);
    const b = indexOfChild(container, endMark);
    const inner = container.children.slice(a + 1, b);
    const link = el('w:hyperlink', attrs, inner);
    for (const run of findAll(link, 'w:r')) {
      const rPr = firstChild(run, 'w:rPr');
      if (!firstChild(rPr, 'w:rStyle')) applyFormat(run, { style: 'Hyperlink' });
    }
    container.children.splice(a, b - a + 1, link);
    this.changed();
    return link;
  }

  /** Unwrap a w:hyperlink, keeping its text. */
  removeHyperlink(link: XmlElement): void {
    this.check('link');
    const idx = this.index();
    const parent = idx.parentOf(link);
    const runs = link.children;
    for (const run of findAll(link, 'w:r')) {
      const rPr = firstChild(run, 'w:rPr');
      const rStyle = firstChild(rPr, 'w:rStyle');
      if (attr(rStyle, 'w:val') === 'Hyperlink') removeChildren(rPr!, 'w:rStyle');
    }
    parent.children.splice(indexOfChild(parent, link), 1, ...runs);
    this.changed();
  }

  /** Every w:hyperlink in the body with its target. */
  hyperlinks(): Array<{ el: XmlElement; url?: string; anchor?: string; tooltip?: string; text: string }> {
    const rels = this.doc.rels(this.doc.mainPart);
    return findAll(this.body, 'w:hyperlink').map((h) => {
      const rId = attr(h, 'r:id');
      return { el: h, url: rId ? rels.get(rId)?.target : undefined, anchor: attr(h, 'w:anchor'), tooltip: attr(h, 'w:tooltip'), text: textContent(h) };
    });
  }

  addBookmark(range: DocRange, name: string): void {
    this.check('bookmark', range);
    if (!BOOKMARK_NAME.test(name)) throw new EditError('Bookmark names start with a letter and use up to 40 letters, digits or underscores');
    if (this.bookmarks().some((b) => b.name === name)) throw new EditError(`A bookmark named "${name}" already exists`);
    const id = String(1 + Math.max(-1, ...findAll(this.root, 'w:bookmarkStart').map((b) => Number(attr(b, 'w:id') ?? -1))));
    const index = this.index();
    const { startMark, endMark } = isolate(index, range);
    const idx = this.index();
    replaceMark(idx, startMark, [el('w:bookmarkStart', { 'w:id': id, 'w:name': name })]);
    replaceMark(idx, endMark, [el('w:bookmarkEnd', { 'w:id': id })]);
    this.changed();
  }

  bookmarks(): Array<{ id: string; name: string }> {
    return findAll(this.root, 'w:bookmarkStart').map((b) => ({ id: attr(b, 'w:id') ?? '', name: attr(b, 'w:name') ?? '' }));
  }

  removeBookmark(name: string): void {
    this.check('bookmark');
    const start = findAll(this.root, 'w:bookmarkStart').find((b) => attr(b, 'w:name') === name);
    if (!start) return;
    const id = attr(start, 'w:id');
    const idx = this.index();
    for (const node of [start, ...findAll(this.root, 'w:bookmarkEnd').filter((b) => attr(b, 'w:id') === id)]) {
      const parent = idx.parentOf(node);
      parent.children.splice(indexOfChild(parent, node), 1);
    }
    this.changed();
  }

  // ---- comments -----------------------------------------------------------------------------

  /** Add a comment anchored to a range. Returns the new comment id. */
  async addComment(range: DocRange, input: CommentInput): Promise<string> {
    this.check('comment', range);
    const id = await this.writeComment(input);
    const index = this.index();
    const { startMark, endMark } = isolate(index, range);
    const idx = this.index();
    replaceMark(idx, startMark, [el('w:commentRangeStart', { 'w:id': id })]);
    replaceMark(idx, endMark, [el('w:commentRangeEnd', { 'w:id': id }), referenceRun(id)]);
    this.changed();
    return id;
  }

  /** Reply to a comment (Word threads replies on the same anchor). */
  async replyToComment(parentId: string, input: CommentInput): Promise<string> {
    this.check('reply');
    const parent = await this.commentElement(parentId);
    if (!parent) throw new EditError(`No comment with id ${parentId}`);
    const id = await this.writeComment(input, parentId);
    const idx = this.index();
    const start = findAll(this.root, 'w:commentRangeStart').find((n) => attr(n, 'w:id') === parentId);
    const end = findAll(this.root, 'w:commentRangeEnd').find((n) => attr(n, 'w:id') === parentId);
    if (start) {
      const p = idx.parentOf(start);
      p.children.splice(indexOfChild(p, start) + 1, 0, el('w:commentRangeStart', { 'w:id': id }));
    }
    const refRun = findAll(this.root, 'w:commentReference').find((n) => attr(n, 'w:id') === parentId);
    const anchorNode = refRun ? idx.parentOf(refRun) : end;
    if (anchorNode) {
      const p = idx.parentOf(anchorNode);
      p.children.splice(indexOfChild(p, anchorNode) + 1, 0, el('w:commentRangeEnd', { 'w:id': id }), referenceRun(id));
    }
    this.changed();
    return id;
  }

  /** Mark a comment thread resolved (done) or open again. */
  async resolveComment(id: string, resolved = true): Promise<void> {
    this.check('resolveComment');
    const comment = await this.commentElement(id);
    if (!comment) throw new EditError(`No comment with id ${id}`);
    const paraId = this.ensureParaId(comment);
    const { name, xml } = await ensurePart(this.doc, 'commentsExtended');
    let entry = findAll(xml.root, 'w15:commentEx').find((e) => attr(e, 'w15:paraId')?.toUpperCase() === paraId);
    if (!entry) {
      entry = el('w15:commentEx', { 'w15:paraId': paraId, 'w15:done': '0' });
      xml.root.children.push(entry);
    }
    entry.attrs['w15:done'] = resolved ? '1' : '0';
    this.doc.touch(name);
    this.doc.touch(this.doc.parts.comments!);
  }

  /** Replace a comment's text. */
  async editComment(id: string, text: string): Promise<void> {
    const comment = await this.commentElement(id);
    if (!comment) throw new EditError(`No comment with id ${id}`);
    this.check('comment', undefined, { commentAuthor: attr(comment, 'w:author') });
    const paraId = this.ensureParaId(comment);
    comment.children = [commentParagraph(text, paraId)];
    this.doc.touch(this.doc.parts.comments!);
  }

  /** Delete a comment, its replies and its anchors. */
  async deleteComment(id: string): Promise<void> {
    const comment = await this.commentElement(id);
    if (!comment) return;
    this.check('deleteComment', undefined, { commentAuthor: attr(comment, 'w:author') });
    const { readComments } = await import('../comments/read.js');
    const all = readComments(this.doc);
    const doomed = new Set<string>([id]);
    for (let grew = true; grew; ) {
      grew = false;
      for (const c of all) if (c.parentId && doomed.has(c.parentId) && !doomed.has(c.id)) {
        doomed.add(c.id);
        grew = true;
      }
    }
    const paraIds = new Set(all.filter((c) => doomed.has(c.id) && c.paraId).map((c) => c.paraId!));
    const durableIds = new Set(all.filter((c) => doomed.has(c.id) && c.durableId).map((c) => c.durableId!));

    const commentsXml = this.doc.xmlSync(this.doc.parts.comments)!;
    commentsXml.root.children = commentsXml.root.children.filter((n) => !(n.type === 'element' && n.name === 'w:comment' && doomed.has(attr(n, 'w:id') ?? '')));
    this.doc.touch(this.doc.parts.comments!);

    const prune = async (kind: 'commentsExtended' | 'commentsIds' | 'commentsExtensible', test: (n: XmlElement) => boolean) => {
      const part = this.doc.parts[kind];
      const xml = this.doc.xmlSync(part);
      if (!part || !xml) return;
      xml.root.children = xml.root.children.filter((n) => !(n.type === 'element' && test(n)));
      this.doc.touch(part);
    };
    await prune('commentsExtended', (n) => paraIds.has(attr(n, 'w15:paraId')?.toUpperCase() ?? ''));
    await prune('commentsIds', (n) => paraIds.has(attr(n, 'w16cid:paraId')?.toUpperCase() ?? ''));
    await prune('commentsExtensible', (n) => durableIds.has(attr(n, 'w16cex:durableId') ?? ''));

    // Anchors in the body: range markers and the runs holding references.
    const idx = this.index();
    for (const name of ['w:commentRangeStart', 'w:commentRangeEnd']) {
      for (const n of findAll(this.root, name)) {
        if (!doomed.has(attr(n, 'w:id') ?? '')) continue;
        const parent = idx.parentOf(n);
        parent.children.splice(indexOfChild(parent, n), 1);
      }
    }
    for (const ref of findAll(this.root, 'w:commentReference')) {
      if (!doomed.has(attr(ref, 'w:id') ?? '')) continue;
      const run = idx.parentOf(ref);
      const parent = idx.parentOf(run);
      run.children.splice(indexOfChild(run, ref), 1);
      if (!run.children.some((c) => c.type === 'element' && c.name !== 'w:rPr')) parent.children.splice(indexOfChild(parent, run), 1);
    }
    this.changed();
  }

  private async commentElement(id: string): Promise<XmlElement | undefined> {
    const xml = this.doc.xmlSync(this.doc.parts.comments);
    return xml ? elements(xml.root, 'w:comment').find((c) => attr(c, 'w:id') === id) : undefined;
  }

  private ensureParaId(comment: XmlElement): string {
    const paras = elements(comment, 'w:p');
    let last = paras[paras.length - 1];
    if (!last) {
      last = el('w:p');
      comment.children.push(last);
    }
    let paraId = attr(last, 'w14:paraId')?.toUpperCase();
    if (!paraId) {
      paraId = randomHexId();
      last.attrs['w14:paraId'] = paraId;
      last.attrs['w14:textId'] = '77777777';
    }
    return paraId;
  }

  private async writeComment(input: CommentInput, parentId?: string): Promise<string> {
    if (!input.author.trim()) throw new EditError('A comment needs an author');
    const comments = await ensurePart(this.doc, 'comments');
    ensureNamespaces(comments.xml);
    const ids = elements(comments.xml.root, 'w:comment').map((c) => Number(attr(c, 'w:id') ?? -1));
    // Comment ids share a space with other annotation ids in Word; stay above all of them.
    const bodyIds = findAll(this.root, 'w:commentRangeStart').map((c) => Number(attr(c, 'w:id') ?? -1));
    const id = String(1 + Math.max(-1, ...ids, ...bodyIds));
    const date = input.date ?? new Date();
    const paraId = randomHexId();
    const initials = input.initials ?? input.author.split(/\s+/).map((w) => w[0] ?? '').join('').slice(0, 3).toUpperCase();
    comments.xml.root.children.push(
      el('w:comment', { 'w:id': id, 'w:author': input.author, 'w:date': localDate(date), 'w:initials': initials }, [commentParagraph(input.text, paraId)]),
    );
    this.doc.touch(comments.name);

    let parentParaId: string | undefined;
    if (parentId) {
      const parent = elements(comments.xml.root, 'w:comment').find((c) => attr(c, 'w:id') === parentId);
      if (parent) parentParaId = this.ensureParaId(parent);
    }
    const ext = await ensurePart(this.doc, 'commentsExtended');
    ext.xml.root.children.push(el('w15:commentEx', { 'w15:paraId': paraId, 'w15:paraIdParent': parentParaId, 'w15:done': '0' }));
    this.doc.touch(ext.name);

    const durableId = randomHexId();
    const idsPart = await ensurePart(this.doc, 'commentsIds');
    idsPart.xml.root.children.push(el('w16cid:commentId', { 'w16cid:paraId': paraId, 'w16cid:durableId': durableId }));
    this.doc.touch(idsPart.name);

    const extensible = await ensurePart(this.doc, 'commentsExtensible');
    extensible.xml.root.children.push(el('w16cex:commentExtensible', { 'w16cex:durableId': durableId, 'w16cex:dateUtc': date.toISOString().replace(/\.\d{3}Z$/, 'Z') }));
    this.doc.touch(extensible.name);

    const people = await ensurePart(this.doc, 'people');
    if (!elements(people.xml.root, 'w15:person').some((p) => attr(p, 'w15:author') === input.author)) {
      people.xml.root.children.push(el('w15:person', { 'w15:author': input.author }, [el('w15:presenceInfo', { 'w15:providerId': 'None', 'w15:userId': input.author })]));
      this.doc.touch(people.name);
    }
    return id;
  }

  // ---- styles -------------------------------------------------------------------------------

  private async ensureCharacterStyle(id: string, name: string, rPrXml: string): Promise<void> {
    if (this.doc.styles.get(id)) return;
    const styles = await ensurePart(this.doc, 'styles');
    const { parseXml } = await import('../xml/index.js');
    const fragment = parseXml(`<w:style xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:type="character" w:styleId="${id}"><w:name w:val="${name}"/><w:uiPriority w:val="99"/><w:unhideWhenUsed/><w:rPr>${rPrXml}</w:rPr></w:style>`).root;
    delete fragment.attrs['xmlns:w'];
    styles.xml.root.children.push(fragment);
    this.doc.touch(styles.name);
  }
}

function isAncestor(idx: TreeIndex, node: XmlElement, descendant: XmlElement): boolean {
  let cur = idx.parent.get(descendant);
  while (cur) {
    if (cur === node) return true;
    cur = idx.parent.get(cur);
  }
  return false;
}

function referenceRun(id: string): XmlElement {
  return el('w:r', {}, [el('w:rPr', {}, [el('w:rStyle', { 'w:val': 'CommentReference' })]), el('w:commentReference', { 'w:id': id })]);
}

function commentParagraph(text: string, paraId: string): XmlElement {
  return el('w:p', { 'w14:paraId': paraId, 'w14:textId': '77777777' }, [
    el('w:pPr', {}, [el('w:pStyle', { 'w:val': 'CommentText' })]),
    el('w:r', {}, [el('w:rPr', {}, [el('w:rStyle', { 'w:val': 'CommentReference' })]), el('w:annotationRef')]),
    el('w:r', {}, runContent(text)),
  ]);
}

/** Word writes w:date as local time without a zone designator. */
function localDate(d: Date): string {
  const pad = (n: number) => String(n).padStart(2, '0');
  return `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}T${pad(d.getHours())}:${pad(d.getMinutes())}:${pad(d.getSeconds())}Z`;
}

export type { DocPosition, DocRange };
export { EditError };
