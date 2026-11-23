/**
 * DocxEditor: the code API for changing a document. Every change edits the XML directly (so
 * everything docxkeun does not understand is preserved), passes through the access guard, and
 * notifies listeners such as the viewer.
 */
import { CONTENT_TYPE, IMAGE_CONTENT_TYPES, REL } from '../opc/index.js';
import { newChartSpace, patchChartSpace, readChartTable, validateChartData, type ChartChanges, type NewChart } from '../charts/write.js';
import { isWorkbook, workbookFor, workbookHasMore } from '../charts/workbook.js';
import { imageInfo } from '../images/size.js';
import { hasEvenPages, hasFirstPage, headerParagraphs, headerText, referencesOf, sectionProperties, SECTPR_ORDER, textWidth, type HeaderFooterInfo, type HeaderFooterKind, type HeaderFooterOptions, type HeaderFooterType } from './headers.js';
import type { DocParagraph, Block, DrawingContent, Inline } from '../model/content.js';
import type { DocxDocument } from '../model/document.js';
import { attr, el, elements, ensureChild, findAll, firstChild, insertOrdered, NS, removeChildren, textContent, type XmlElement, type XmlNode } from '../xml/index.js';
import { createList, listItems, type ListFormat, type ListItem } from './lists.js';
import { buildTable, cleanCopy, columnWidths, markInserted, markRow, type CopyDropped, type CopyIds, type TableOptions } from './tables.js';
import { cloneElement } from '../xml/index.js';
import {
  acceptRevisions,
  cleanRunProps,
  deleteRun,
  insertedRuns,
  listRevisions,
  markParagraph,
  nextRevisionId,
  paragraphMarkRevision,
  recordParagraphFormat,
  recordRunFormat,
  rejectRevisions,
  revisionStamp,
  runRevision,
  type Revision,
  type RevisionAuthor,
  type RevisionFilter,
} from './revisions.js';
import { SETTINGS_ORDER } from '../security/protection.js';
import { ensureNamespaces, ensurePart, randomHexId, ROOT_NAMESPACES } from './parts.js';
import {
  addWorkbook,
  anchorFrame,
  chartData,
  chartGraphic,
  chartKind,
  chartOf,
  chartSize,
  checkPlacement,
  checkShape,
  cloneObject,
  copyPart,
  dropFallback,
  frameOf,
  freePartName,
  inlineFrame,
  inTextBox,
  listObjects,
  nextRelativeHeight,
  placeFrame,
  replaceWorkbook,
  resized,
  resizeFrame,
  sectionAt,
  SHAPE_NAMES,
  shapeGraphic,
  takeOut,
  type ChartInfo,
  type ChartUpdate,
  type DocObject,
  type NewShape,
  type ObjectPlacement,
} from './objects.js';
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

export type EditAction = 'edit' | 'format' | 'comment' | 'reply' | 'resolveComment' | 'editComment' | 'deleteComment' | 'link' | 'bookmark' | 'fillForm' | 'protect' | 'review';

/**
 * How a locked region is locked (Word's content-control locking, w:sdtPr/w:lock):
 * 'content' = the text cannot be edited; 'contentAndControl' = and the region cannot be deleted;
 * 'control' = the region cannot be deleted, but its text can be edited.
 */
export type LockMode = 'content' | 'contentAndControl' | 'control';

const LOCK_VALUES: Record<LockMode, string> = { content: 'contentLocked', contentAndControl: 'sdtContentLocked', control: 'sdtLocked' };
const LOCK_MODES: Record<string, LockMode> = { contentLocked: 'content', sdtContentLocked: 'contentAndControl', sdtLocked: 'control' };

/** Tag docxkeun puts on the content controls it creates to lock text. */
export const LOCK_TAG = 'docxkeun-lock';

export interface LockedRegion {
  /** The content control's w:id. */
  id: string;
  mode: LockMode;
  title?: string;
  tag?: string;
  text: string;
  /** Whether it wraps whole paragraphs (block) or text inside a paragraph (inline). */
  block: boolean;
  el: XmlElement;
}

/** An edit touched text inside a locked region. */
export class LockedRegionError extends EditError {
  override name = 'LockedRegionError';
  constructor(readonly region: XmlElement, title: string | undefined) {
    super(title ? `"${title}" is locked against editing` : 'This text is locked against editing');
  }
}

/** The lock mode of a content control, if it is locked. */
export function lockModeOf(sdt: XmlElement): LockMode | undefined {
  const v = attr(firstChild(firstChild(sdt, 'w:sdtPr'), 'w:lock'), 'w:val');
  return v ? LOCK_MODES[v] : undefined;
}

function sdtTitle(sdt: XmlElement): string | undefined {
  return attr(firstChild(firstChild(sdt, 'w:sdtPr'), 'w:alias'), 'w:val');
}

function locksContent(mode: LockMode | undefined): boolean {
  return mode === 'content' || mode === 'contentAndControl';
}

/** Checks every edit (see docxkeun/security's AccessPolicy). Throw to refuse. */
export interface EditGuard {
  /** `detail.tracked` is set when the edit is recorded as a tracked change. */
  assert(action: EditAction, range?: DocRange, detail?: { commentAuthor?: string; tracked?: boolean }): void;
}

export interface EditorOptions {
  /** Record edits as tracked changes. Default: the document's own Track Changes setting. */
  trackChanges?: boolean;
  /** Author of tracked changes (default: the guard's user name, else "docxkeun"). */
  author?: string;
  /**
   * Edit another story: a header or footer part, or the footnotes or endnotes part (see
   * `DocxDocument.storyParts()`). Default: the main document (body and text boxes).
   */
  part?: string;
}

/** Whether Word would record changes in this document: Track Changes on, or protected for tracked changes. */
export function documentTracksChanges(doc: DocxDocument): boolean {
  const settings = doc.parts.settings ? doc.xmlSync(doc.parts.settings)?.root : undefined;
  if (!settings) return false;
  const track = firstChild(settings, 'w:trackRevisions');
  if (track && attr(track, 'w:val') !== 'false' && attr(track, 'w:val') !== '0') return true;
  const protection = firstChild(settings, 'w:documentProtection');
  const enforced = attr(protection, 'w:enforcement');
  return attr(protection, 'w:edit') === 'trackedChanges' && (enforced === '1' || enforced === 'true' || enforced === 'on');
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

/** A piece of new text for `replaceRange`: formatted like the run `like` (default: the range's first run), then `format`. */
export interface TextPiece {
  text: string;
  /** A w:r element whose formatting the piece copies. */
  like?: XmlElement;
  format?: TextFormat;
}

/** Paragraph formatting for `formatParagraph` (lengths in points; undefined leaves a property alone). */
export interface ParagraphFormat {
  align?: 'left' | 'center' | 'right' | 'justify';
  /** Right-to-left paragraph (Arabic, Hebrew): w:bidi. */
  direction?: 'ltr' | 'rtl';
  indentLeft?: number;
  indentRight?: number;
  /** Positive: first-line indent; negative: hanging indent. */
  indentFirstLine?: number;
  spaceBefore?: number;
  spaceAfter?: number;
  /** Line spacing as a multiple of single spacing (1, 1.15, 1.5, 2). */
  lineSpacing?: number;
  keepWithNext?: boolean;
  keepLinesTogether?: boolean;
  pageBreakBefore?: boolean;
}

export interface SetListOptions {
  /** A list format, or 'none' to take the paragraphs out of their list. */
  format?: ListFormat | 'none';
  /** List level, 0–8. */
  level?: number;
  /** First number of the list. */
  start?: number;
  /** Join this paragraph's list instead of starting a new one. */
  sameListAs?: XmlElement;
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

export interface TextSegment {
  start: number;
  end: number;
  /** w:t, or a one-character item such as w:tab, w:br or w:sym. */
  node: XmlElement;
}

/** Concatenated text of a paragraph and where each character lives in the XML. */
export interface ParagraphText {
  paragraph: XmlElement;
  text: string;
  segments: TextSegment[];
}

/**
 * Every paragraph of a block list (including table cells), in document order. With `textBoxes`,
 * the paragraphs of the text boxes in a paragraph follow it (a group's boxes depth first).
 */
export function paragraphsOf(blocks: Block[], out: DocParagraph[] = [], options: { textBoxes?: boolean } = {}): DocParagraph[] {
  for (const b of blocks) {
    if (b.kind === 'paragraph') {
      out.push(b);
      if (options.textBoxes) for (const inline of b.inlines) if (inline.kind === 'image' && inline.content) boxParagraphs(inline.content, out, options);
    } else for (const row of b.rows) for (const cell of row.cells) paragraphsOf(cell.blocks, out, options);
  }
  return out;
}

function boxParagraphs(content: DrawingContent, out: DocParagraph[], options: { textBoxes?: boolean }): void {
  if (content.type === 'shape' && content.shape.text) paragraphsOf(content.shape.text.blocks, out, options);
  else if (content.type === 'group') for (const child of content.children) boxParagraphs(child.content, out, options);
}

export function paragraphText(para: DocParagraph, options: { objects?: boolean } = {}): ParagraphText {
  let text = '';
  const segments: TextSegment[] = [];
  const add = (value: string, node: XmlElement) => {
    segments.push({ start: text.length, end: text.length + value.length, node });
    text += value;
  };
  for (const inline of para.inlines as Inline[]) {
    if (inline.kind === 'text') add(inline.text, inline.node);
    else if (inline.kind === 'tab') add('\t', inline.node);
    else if (inline.kind === 'break' && inline.type === 'line') add('\n', inline.node);
    else if (options.objects && inline.kind === 'image' && !inline.float) {
      // An inline picture or shape is one character (U+FFFC) the caret can step over and delete.
      const node = objectNode(inline);
      if (node) add(OBJECT_CHAR, node);
    }
  }
  return { paragraph: para.el, text, segments };
}

/** Stands for an inline picture, shape or chart in `paragraphText(p, { objects: true })`. */
export const OBJECT_CHAR = '\ufffc';

/** The run child (w:drawing, w:pict, mc:AlternateContent) holding an inline object. */
export function objectNode(inline: { run: XmlElement; drawing: XmlElement }): XmlElement | undefined {
  return elements(inline.run).find((c) => c === inline.drawing || findAll(c, inline.drawing.name).includes(inline.drawing));
}

/**
 * XML position of character offset `k` in a paragraph. With side 'start' the position is at the
 * start of the segment holding character k; with 'end' it is at the end of the segment holding
 * character k-1 (so insertions there inherit the formatting of the text before).
 */
export function positionAt(pt: ParagraphText, k: number, side: 'start' | 'end'): DocPosition | undefined {
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

/** Position just after a run's content (its last text node, or after its last item). */
function endOfRun(run: XmlElement, fallback: XmlElement): DocPosition {
  const items = run.children.filter((c): c is XmlElement => c.type === 'element' && c.name !== 'w:rPr');
  const last = items[items.length - 1];
  if (!last) return { node: fallback, offset: 0 };
  return last.name === 'w:t' ? { node: last, offset: textContent(last).length } : { node: last, offset: 1 };
}

/** Whether a list level is in a set_list format (decimal "1." and decimalParen "1)" differ by their text). */
function levelMatches(level: { format: string; text: string }, format: ListFormat): boolean {
  if (format === 'decimal') return level.format === 'decimal' && !/\)\s*$/.test(level.text);
  if (format === 'decimalParen') return level.format === 'decimal' && /\)\s*$/.test(level.text);
  return level.format === format;
}

function describeLevel(level: { format: string; text: string }): string {
  if (level.format === 'bullet') return 'bulleted';
  if (level.format === 'decimal') return /\)\s*$/.test(level.text) ? 'numbered 1)' : 'numbered 1.';
  return `numbered with ${level.format}`;
}

function newRun(text: string, template?: XmlElement, format?: TextFormat): XmlElement {
  const rPr = template ? firstChild(template, 'w:rPr') : undefined;
  const run = el('w:r', {}, [...(rPr ? [structuredClone(rPr)] : []), ...runContent(text)]);
  // A new run starts with no revision history of its own.
  cleanRunProps(run);
  if (format) applyFormat(run, format);
  return run;
}

const BOOKMARK_NAME = /^[\p{L}_][\p{L}\p{N}_]{0,39}$/u;

/** Run children that are objects: pictures, charts, shapes and text boxes (DrawingML, VML or with Word's old copy), and embedded OLE objects. */
const OBJECT_NODES = new Set(['w:drawing', 'w:pict', 'w:object', 'mc:AlternateContent']);

/** Objects `replaceRange` keeps: the runs holding them (see `takeObjects`) and how many there are. */
interface KeptObjects {
  runs: XmlNode[];
  count: number;
  nextId: () => string;
}

const OLD_VML = 'This is an old-style (VML) object: it can only be deleted';
// Word doesn't put shapes or text boxes in a text box, nor floating objects.
const SHAPE_IN_BOX = "Shapes and text boxes can't go inside a text box";
const INLINE_IN_BOX = 'An object inside a text box can only be inline';

export class DocxEditor {
  private trackOverride: boolean | undefined;
  private readonly authorName: string | undefined;
  /** The part being edited. */
  readonly part: string;

  constructor(
    readonly doc: DocxDocument,
    readonly guard?: EditGuard,
    options: EditorOptions = {},
  ) {
    this.trackOverride = options.trackChanges;
    this.authorName = options.author;
    this.part = options.part ?? doc.mainPart;
    if (this.part !== doc.mainPart && !doc.xmlSync(this.part)) throw new EditError(`No part "${this.part}" to edit`);
  }

  /** Whether this editor works on a header, footer or notes part rather than the main document. */
  private get isStory(): boolean {
    return this.part !== this.doc.mainPart;
  }

  // ---- tracked changes ----------------------------------------------------------------------

  /** Whether edits are recorded as tracked changes (default: the document's Track Changes setting). */
  get trackChanges(): boolean {
    return this.trackOverride ?? documentTracksChanges(this.doc);
  }

  set trackChanges(on: boolean) {
    this.trackOverride = on;
  }

  /** The author tracked changes are recorded under. */
  get author(): string {
    const user = (this.guard as { user?: { name?: string; id?: string } } | undefined)?.user;
    return this.authorName ?? user?.name ?? user?.id ?? 'docxkeun';
  }

  /** Revision stamp and id source for one edit, or undefined when not tracking. */
  private tracking(): { who: RevisionAuthor; nextId: () => string } | undefined {
    if (!this.trackChanges) return undefined;
    return { who: revisionStamp(this.author), nextId: nextRevisionId(this.root) };
  }

  /** Turn Word's own Track Changes setting (settings.xml w:trackRevisions) on or off. */
  async setTrackChangesSetting(on: boolean): Promise<void> {
    this.guard?.assert('edit');
    const { name, xml } = await ensurePart(this.doc, 'settings');
    removeChildren(xml.root, 'w:trackRevisions');
    if (on) insertOrdered(xml.root, el('w:trackRevisions'), SETTINGS_ORDER);
    this.trackOverride = undefined;
    this.doc.touch(name);
  }

  /** Tracked changes in the main document, in document order. */
  revisions(): Revision[] {
    return listRevisions(this.root);
  }

  /** Accept tracked changes (all, or those matching ids / author). Returns how many. */
  acceptRevisions(filter?: RevisionFilter): number {
    this.guard?.assert('review');
    const n = acceptRevisions(this.root, filter);
    if (n) this.changed();
    return n;
  }

  /** Reject tracked changes (all, or those matching ids / author). Returns how many. */
  rejectRevisions(filter?: RevisionFilter): number {
    this.guard?.assert('review');
    const n = rejectRevisions(this.root, filter);
    if (n) this.changed();
    return n;
  }

  /** The new run as it goes into the document: wrapped in a w:ins when tracking. */
  private placed(run: XmlElement, track: ReturnType<DocxEditor['tracking']>): XmlElement {
    return track ? insertedRuns([run], track.nextId(), track.who) : run;
  }

  /** Whether a run sits in an insertion by the current author (edits to it need no new revision). */
  private inOwnInsertion(run: XmlElement, idx: TreeIndex): boolean {
    const rev = runRevision(run, (n) => idx.parent.get(n));
    return !!rev && (rev.name === 'w:ins' || rev.name === 'w:moveTo') && attr(rev, 'w:author') === this.author;
  }

  /**
   * A sentinel inside someone else's w:ins cannot take a nested w:ins: split that insertion at the
   * sentinel so the new one goes between its two halves.
   */
  private liftOutOfInsertion(mark: XmlElement, idx: TreeIndex): void {
    const container = idx.parent.get(mark);
    if (!container || (container.name !== 'w:ins' && container.name !== 'w:moveTo')) return;
    const outer = idx.parentOf(container);
    const at = container.children.indexOf(mark);
    const tail = container.children.splice(at);
    tail.shift();
    const pos = outer.children.indexOf(container);
    const pieces: XmlNode[] = [mark];
    if (tail.some((n) => n.type === 'element')) pieces.push(el(container.name, { ...container.attrs }, tail));
    outer.children.splice(pos + 1, 0, ...pieces);
    if (!container.children.some((n) => n.type === 'element')) outer.children.splice(outer.children.indexOf(container), 1);
  }

  /** Mark the runs of a range as deleted (and the paragraph marks inside it), keeping the text. */
  private trackedCut(
    range: DocRange,
    track: NonNullable<ReturnType<DocxEditor['tracking']>>,
    keep?: KeptObjects,
  ): { startMark: XmlElement; endMark: XmlElement; idx: TreeIndex } {
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
    if (keep) {
      this.takeObjects(idx, startMark, endMark, keep);
      idx = this.index();
    }
    const startPara = idx.ancestor(startMark, 'w:p')!;
    const endPara = idx.ancestor(endMark, 'w:p')!;
    const paragraphs = startPara === endPara ? [] : [startPara, ...between(idx, startMark, endMark, 'w:p').filter((p) => p !== endPara)];
    for (const run of between(idx, startMark, endMark, 'w:r')) deleteRun(run, (n) => idx.parent.get(n), track.nextId, track.who);
    for (const p of paragraphs) markParagraph(p, 'del', track.nextId(), track.who);
    idx = this.index();
    return { startMark, endMark, idx };
  }

  private get root(): XmlElement {
    return this.isStory ? this.doc.xmlSync(this.part)!.root : this.doc.documentXml.root;
  }

  private get body(): XmlElement {
    if (this.isStory) return this.root;
    const body = firstChild(this.root, 'w:body');
    if (!body) throw new EditError('document has no body');
    return body;
  }

  private index(): TreeIndex {
    return new TreeIndex(this.root);
  }

  private changed(part = this.part): void {
    this.doc.touch(part);
  }

  private check(action: EditAction, range?: DocRange, detail?: { commentAuthor?: string }): void {
    const tracked = this.trackChanges && (action === 'edit' || action === 'format' || action === 'fillForm');
    this.guard?.assert(action, range, tracked ? { ...detail, tracked: true } : detail);
    if (!range) return;
    const idx = this.index();
    if (action === 'edit' || action === 'format' || action === 'link') this.assertUnlocked(range, idx);
    this.dropStaleFallbacks([range.start.node, range.end.node], idx);
  }

  /**
   * Permission for a change Word can't track (an object's size, a floating object's position or
   * wrap, alt text, chart data): it applies directly even with Track Changes on, so it is refused
   * where only tracked changes are allowed. Locks hold as for any edit.
   */
  private checkUntracked(range: DocRange): void {
    this.guard?.assert('format', range, { tracked: false });
    const idx = this.index();
    this.assertUnlocked(range, idx);
    this.dropStaleFallbacks([range.start.node, range.end.node], idx);
  }

  /**
   * Word stores a text box twice: DrawingML in mc:Choice and an old VML copy in mc:Fallback. A
   * change inside the box (at any of `nodes`) leaves the copy stale, so it goes.
   */
  private dropStaleFallbacks(nodes: XmlElement[], idx = this.index()): void {
    for (const node of nodes) {
      let inBox = false;
      for (let cur: XmlElement | undefined = node; cur; cur = idx.parent.get(cur)) {
        if (cur.name === 'w:txbxContent') inBox = true;
        else if (cur.name === 'mc:Choice' && inBox) {
          const alternate = idx.parent.get(cur);
          if (alternate) alternate.children = alternate.children.filter((c) => !(c.type === 'element' && c.name === 'mc:Fallback'));
        }
      }
    }
  }

  /** The locked content control around a node, if any. */
  private lockAround(idx: TreeIndex, node: XmlElement): XmlElement | undefined {
    for (let cur = idx.parent.get(node); cur; cur = idx.parent.get(cur)) {
      if (cur.name === 'w:sdt' && locksContent(lockModeOf(cur))) return cur;
    }
    return undefined;
  }

  /** Throw LockedRegionError when a range touches a locked region (or would delete one). */
  private assertUnlocked(range: DocRange, idx = this.index()): void {
    if (this.bypassLocks) return;
    for (const pos of [range.start, range.end]) {
      if (!idx.contains(pos.node)) continue;
      const lock = this.lockAround(idx, pos.node);
      if (lock) throw new LockedRegionError(lock, sdtTitle(lock));
    }
    const a = idx.order.get(range.start.node);
    const b = idx.order.get(range.end.node);
    if (a === undefined || b === undefined || (a === b && range.start.offset === range.end.offset)) return;
    const [lo, hi] = a <= b ? [a, b] : [b, a];
    for (const sdt of findAll(this.root, 'w:sdt')) {
      const o = idx.order.get(sdt)!;
      if (o > lo && o < hi && lockModeOf(sdt)) throw new LockedRegionError(sdt, sdtTitle(sdt));
    }
  }

  /** Whether a range touches a locked region. */
  isLocked(range: DocRange): boolean {
    try {
      this.assertUnlocked(range);
      return false;
    } catch (e) {
      if (e instanceof LockedRegionError) return true;
      throw e;
    }
  }

  /** Set while filling locked content on purpose (fillTemplate with includeLocked). */
  private bypassLocks = false;

  // ---- reading ------------------------------------------------------------------------------

  /** Convert a viewer selection (inline positions) into an editor range. */
  rangeFromSelection(sel: { start: { inline: { node: XmlElement }; offset: number }; end: { inline: { node: XmlElement }; offset: number } }): DocRange {
    return { start: { node: sel.start.inline.node, offset: sel.start.offset }, end: { node: sel.end.inline.node, offset: sel.end.offset } };
  }

  /**
   * Every paragraph of the story (the body by default), including those in tables, in document
   * order. In the body, a text box's paragraphs come right after the paragraph holding the box.
   */
  paragraphs(): DocParagraph[] {
    if (!this.isStory) return paragraphsOf(this.doc.body().blocks, [], { textBoxes: true });
    const root = this.root;
    if (root.name === 'w:footnotes' || root.name === 'w:endnotes') {
      // Separator notes (w:type) are not content.
      return elements(root)
        .filter((n) => !attr(n, 'w:type') || attr(n, 'w:type') === 'normal')
        .flatMap((n) => paragraphsOf(this.doc.readBlocks(this.part, n)));
    }
    return paragraphsOf(this.doc.headerFooter(this.part).blocks);
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

  /**
   * Insert text at a position, formatted like the run it lands in (plus optional format).
   * Returns the position right after the inserted text. A position whose node is a w:p means
   * "inside this (empty) paragraph".
   */
  insertText(at: DocPosition, text: string, format?: TextFormat): DocPosition {
    const edge = this.bypassLocks ? undefined : this.lockEdge(at);
    const track = this.tracking();
    if (edge) {
      // Typing at the very edge of a locked region goes just outside it, as in Word.
      this.guard?.assert('edit', { start: at, end: at }, track ? { tracked: true } : undefined);
      const idx = this.index();
      this.dropStaleFallbacks([at.node], idx);
      const run = newRun(text, idx.parentOf(at.node), format);
      const parent = idx.parentOf(edge.sdt);
      parent.children.splice(indexOfChild(parent, edge.sdt) + (edge.after ? 1 : 0), 0, this.placed(run, track));
      this.changed();
      return endOfRun(run, run);
    }
    this.check('edit', { start: at, end: at });
    const plain = !format && !/[\t\r\n]/.test(text);
    // Typing into existing text: change the text node itself instead of adding a run per keystroke
    // (when tracking, only inside the author's own insertion).
    if (plain && at.node.name === 'w:t' && (!track || this.inOwnInsertion(this.index().parentOf(at.node), this.index()))) {
      const current = textContent(at.node);
      const offset = Math.max(0, Math.min(at.offset, current.length));
      const value = current.slice(0, offset) + text + current.slice(offset);
      at.node.children = [{ type: 'text', value }];
      if (/^\s|\s$/.test(value)) at.node.attrs['xml:space'] = 'preserve';
      this.changed();
      return { node: at.node, offset: offset + text.length };
    }
    if (at.node.name === 'w:p') {
      // An empty paragraph: the new run takes the paragraph mark's formatting.
      const markRPr = firstChild(firstChild(at.node, 'w:pPr'), 'w:rPr');
      const run = el('w:r', {}, [...(markRPr ? [structuredClone(markRPr)] : []), ...runContent(text)]);
      // The paragraph mark's own revision markers do not belong on a text run.
      const runPr = firstChild(run, 'w:rPr');
      if (runPr) for (const n of ['w:ins', 'w:del', 'w:moveFrom', 'w:moveTo', 'w:rPrChange']) removeChildren(runPr, n);
      if (format) applyFormat(run, format);
      const pPr = firstChild(at.node, 'w:pPr');
      at.node.children.splice(pPr ? at.node.children.indexOf(pPr) + 1 : 0, 0, this.placed(run, track));
      this.changed();
      return endOfRun(run, at.node);
    }
    const index = this.index();
    const template = index.parentOf(at.node);
    const { startMark, endMark } = isolate(index, { start: at, end: at });
    let idx = this.index();
    const run = newRun(text, template, format);
    replaceMark(idx, endMark, []);
    idx = this.index();
    if (track && !this.inOwnInsertion(template, index)) this.liftOutOfInsertion(startMark, idx);
    idx = this.index();
    replaceMark(idx, startMark, [track && !this.inOwnInsertion(template, index) ? this.placed(run, track) : run]);
    this.changed();
    return endOfRun(run, run);
  }

  /** Insert a run holding `content` (a drawing, a note reference) at a position; returns the run. */
  private insertRunContent(at: DocPosition, content: XmlElement[], rPr?: XmlElement): XmlElement {
    // Insert a placeholder character in a run of its own (a format forces a new run), then swap in the content.
    const end = this.insertText(at, '\u2063', {});
    const t = end.node;
    const run = this.index().parentOf(t);
    run.children.splice(run.children.indexOf(t), 1, ...content);
    if (rPr) {
      removeChildren(run, 'w:rPr');
      run.children.unshift(rPr);
    }
    this.changed();
    return run;
  }

  /**
   * Insert a picture (PNG, JPEG, GIF, BMP, TIFF, EMF or WMF) inline at a position. Without a size it
   * gets its natural size, shrunk to `maxWidth` (default 468pt, a Letter page's text width).
   */
  async insertImage(at: DocPosition, bytes: Uint8Array, options: { width?: number; height?: number; alt?: string; name?: string; maxWidth?: number } = {}): Promise<XmlElement> {
    this.check('edit', { start: at, end: at });
    const info = imageInfo(bytes);
    if (!info) throw new EditError('Unsupported image: use PNG, JPEG, GIF, BMP, TIFF, EMF or WMF');
    let { width, height } = options;
    if (width === undefined && height === undefined) {
      width = info.width;
      height = info.height;
      const max = options.maxWidth ?? 468;
      if (width > max) {
        height = (height * max) / width;
        width = max;
      }
    } else if (width === undefined) width = (info.width * height!) / info.height;
    else if (height === undefined) height = (info.height * width) / info.width;
    const ext = info.kind === 'jpeg' ? 'jpeg' : info.kind;
    let n = 1;
    while (this.doc.pkg.has(`word/media/image${n}.${ext}`)) n++;
    const part = `word/media/image${n}.${ext}`;
    this.doc.pkg.setBytes(part, bytes, IMAGE_CONTENT_TYPES[ext]);
    const rId = await this.doc.pkg.addRelationship(this.part, REL.image, part);
    await this.doc.refreshRels(this.part);
    const id = this.newDrawingId();
    const name = options.name ?? `Picture ${id}`;
    const cx = String(Math.round(width! * 12700));
    const cy = String(Math.round(height! * 12700));
    const docPr: Record<string, string> = { id: String(id), name };
    if (options.alt) docPr.descr = options.alt;
    const drawing = el('w:drawing', {}, [
      el('wp:inline', { distT: '0', distB: '0', distL: '0', distR: '0' }, [
        el('wp:extent', { cx, cy }),
        el('wp:effectExtent', { l: '0', t: '0', r: '0', b: '0' }),
        el('wp:docPr', docPr),
        el('wp:cNvGraphicFramePr', {}, [el('a:graphicFrameLocks', { 'xmlns:a': NS.a, noChangeAspect: '1' })]),
        el('a:graphic', { 'xmlns:a': NS.a }, [
          el('a:graphicData', { uri: 'http://schemas.openxmlformats.org/drawingml/2006/picture' }, [
            el('pic:pic', { 'xmlns:pic': NS.pic }, [
              el('pic:nvPicPr', {}, [el('pic:cNvPr', { id: '0', name }), el('pic:cNvPicPr')]),
              el('pic:blipFill', {}, [el('a:blip', { 'r:embed': rId }), el('a:stretch', {}, [el('a:fillRect')])]),
              el('pic:spPr', {}, [el('a:xfrm', {}, [el('a:off', { x: '0', y: '0' }), el('a:ext', { cx, cy })]), el('a:prstGeom', { prst: 'rect' }, [el('a:avLst')])]),
            ]),
          ]),
        ]),
      ]),
    ]);
    this.insertRunContent(at, [drawing]);
    return drawing;
  }

  /** Declare namespace prefixes on the story's root element where they are missing. */
  private declare(...prefixes: Array<keyof typeof NS>): void {
    for (const prefix of prefixes) if (this.root.attrs[`xmlns:${prefix}`] === undefined) this.root.attrs[`xmlns:${prefix}`] = NS[prefix];
  }

  /** An id for a new drawing (wp:docPr ids must be unique in the document); the root gets the prefixes drawings use. */
  private newDrawingId(): number {
    let id = 0;
    for (const root of [this.doc.documentXml.root, this.root]) for (const d of findAll(root, 'wp:docPr')) id = Math.max(id, Number(attr(d, 'id')) || 0);
    this.declare('wp', 'r');
    return id + 1;
  }

  /** Fresh ids for copied content (see `cleanCopy`): tracked changes, drawings and content controls stay unique. */
  private copyIds(): CopyIds {
    let drawing = this.newDrawingId() - 1;
    const sdtIds = new Set(findAll(this.root, 'w:sdtPr').map((s) => attr(firstChild(s, 'w:id'), 'w:val')));
    const sdt = () => {
      let id: string;
      do id = String(1 + Math.floor(Math.random() * 2147483646));
      while (sdtIds.has(id));
      sdtIds.add(id);
      return id;
    };
    return { revision: nextRevisionId(this.root), drawing: () => String(++drawing), sdt };
  }

  // ---- objects and charts ---------------------------------------------------------------------

  /** Pictures, charts, SmartArt, shapes, text boxes and groups, numbered from 1 in document order. */
  objects(): DocObject[] {
    return listObjects(this.paragraphs());
  }

  /**
   * The run's child holding an object (its w:drawing or w:pict, or the mc:AlternateContent around
   * it), found in the tree as it is now; EditError when the object is no longer in the document or
   * is marked deleted (as the original of a tracked move is).
   */
  private objectNodeOf(obj: DocObject, idx = this.index()): XmlElement {
    if (!idx.contains(obj.drawing)) throw new EditError(`Object ${obj.n} is no longer in the document (list the objects again)`);
    if (idx.ancestor(obj.drawing, 'w:del') || idx.ancestor(obj.drawing, 'w:moveFrom')) throw new EditError(`Object ${obj.n} is deleted (a tracked change; list the objects again)`);
    return this.runChildHolding(obj.drawing, idx);
  }

  /** The child of the nearest run that holds `node` (or is it). */
  private runChildHolding(node: XmlElement, idx: TreeIndex): XmlElement {
    for (let parent = idx.parentOf(node); parent.name !== 'w:r'; parent = idx.parentOf(parent)) node = parent;
    return node;
  }

  /** The range of an object, for access and lock checks. */
  private objectRange(obj: DocObject, idx = this.index()): DocRange {
    const node = this.objectNodeOf(obj, idx);
    return { start: { node, offset: 0 }, end: { node, offset: 1 } };
  }

  /**
   * Insert a chart inline at a position, with its data in an embedded workbook (so Word's "Edit
   * Data" shows it). Without a size it gets Word's 432 × 252 pt, shrunk to the section's text width;
   * with one side, the other keeps those proportions. Returns the w:drawing.
   */
  async insertChart(at: DocPosition, chart: NewChart & { width?: number; height?: number; alt?: string }): Promise<XmlElement> {
    this.check('edit', { start: at, end: at });
    validateChartData(chart.type, chart.stacked, chart.categories, chart.series);
    const idx = this.index();
    if (!idx.contains(at.node)) throw new EditError('That position is not in the document (stale position?)');
    const size = chartSize(chart, textWidth(sectionAt(this.body, idx, at.node)) / 20);
    const { pkg } = this.doc;
    const part = freePartName(pkg, 'word/charts/chart', '.xml');
    const workbook = await addWorkbook(pkg, part, await workbookFor(chart));
    pkg.setXml(part, { declaration: null, prolog: [], root: newChartSpace(chart, workbook), epilog: [] }, CONTENT_TYPE.chart);
    await this.doc.refreshRels(part);
    const rId = await pkg.addRelationship(this.part, REL.chart, part);
    await this.doc.refreshRels(this.part);
    const id = this.newDrawingId();
    const drawing = inlineFrame(id, `Chart ${id}`, size.width, size.height, chartGraphic(rId), chart.alt);
    this.insertRunContent(at, [drawing]);
    return drawing;
  }

  /**
   * Insert a text box or a preset shape at a position, with optional text (each line a paragraph,
   * listed by `paragraphs()` right after the paragraph holding it). By default it floats at the top
   * left of the paragraph with square wrap; see `NewShape` for sizes and colours. Not inside a
   * text box (Word doesn't nest them). Returns the w:drawing.
   */
  async insertShape(at: DocPosition, shape: NewShape): Promise<XmlElement> {
    if (inTextBox(this.index(), at.node)) throw new EditError(SHAPE_IN_BOX);
    this.check('edit', { start: at, end: at });
    const { width, height, place } = checkShape(shape);
    const graphic = shapeGraphic(shape, width, height, (text, format) => newRun(text, undefined, format));
    const id = this.newDrawingId();
    this.declare('a', 'wps', 'mc');
    const name = `${SHAPE_NAMES[shape.shape]} ${id}`;
    const drawing =
      'inline' in place
        ? inlineFrame(id, name, width, height, graphic, shape.alt)
        : anchorFrame(id, name, width, height, place, graphic, shape.alt, nextRelativeHeight([this.doc.documentXml.root, this.root]));
    // Word 2010+ shapes go in mc:Choice; no old VML copy (mc:Fallback) is written.
    this.insertRunContent(at, [el('mc:AlternateContent', {}, [el('mc:Choice', { Requires: 'wps' }, [drawing])])]);
    return drawing;
  }

  /** A chart's data (type, titles, legend, categories, series) and where it comes from. */
  chartInfo(obj: DocObject): ChartInfo {
    const { part, chartSpace } = chartOf(this.doc, obj);
    return { ...readChartTable(chartSpace), source: chartData(this.doc, part, chartSpace).source };
  }

  /**
   * Change a chart's titles, legend or data, keeping its formatting (see `patchChartSpace`). New
   * data must suit the chart's type, and rewrites the embedded workbook; a chart with an old embedded
   * object or no workbook gets a new one. A chart linked to an outside file only has its cached data
   * changed: the file is not, and Word's refresh brings its numbers back. Word does not track chart
   * changes, so neither does this.
   */
  async updateChart(obj: DocObject, changes: ChartChanges): Promise<ChartUpdate> {
    const { part, chartSpace } = chartOf(this.doc, obj);
    this.checkUntracked(this.objectRange(obj));
    const { pkg } = this.doc;
    const data = chartData(this.doc, part, chartSpace);
    const result: ChartUpdate = { source: data.source, workbookRewritten: false, workbookHadMore: false };
    let workbook: Uint8Array | undefined;
    if (changes.categories !== undefined || changes.series !== undefined) {
      const current = readChartTable(chartSpace);
      const categories = changes.categories ?? current.categories;
      const series = changes.series ?? current.series;
      const { type, stacked } = chartKind(chartSpace);
      validateChartData(type, stacked, categories, series);
      if (data.source !== 'linked') {
        // What a workbook that can't be read (an old .xls object, broken bytes) held is unknown: it may have held more.
        if (data.workbook) {
          const old = await pkg.getBytes(data.workbook);
          result.workbookHadMore = !isWorkbook(old) || (await workbookHasMore(old, current.categories.length, current.series.length));
        } else if (data.source === 'embedded') result.workbookHadMore = true;
        workbook = await workbookFor({ categories, series });
      }
    }
    patchChartSpace(chartSpace, changes, { keepFormulas: data.source === 'linked' });
    if (workbook) {
      if (data.workbook) pkg.setBytes(data.workbook, workbook);
      else await replaceWorkbook(this.doc, part, chartSpace, data.relId, workbook);
      result.source = 'embedded';
      result.workbookRewritten = true;
    }
    this.changed(part);
    return result;
  }

  /** Give every chart under `nodes` its own copy of its chart part (and workbook), so copies change independently. */
  private async copyChartParts(nodes: XmlElement[]): Promise<void> {
    const { pkg } = this.doc;
    let copied = false;
    for (const chart of nodes.flatMap((n) => findAll(n, 'c:chart'))) {
      const rel = await pkg.relationshipById(this.part, attr(chart, 'r:id') ?? '');
      if (!rel || rel.external || !pkg.has(pkg.resolve(this.part, rel))) continue;
      const copy = await copyPart(pkg, pkg.resolve(this.part, rel));
      await this.doc.refreshRels(copy);
      chart.attrs['r:id'] = await pkg.addRelationship(this.part, rel.type, copy);
      copied = true;
    }
    if (copied) await this.doc.refreshRels(this.part);
  }

  /**
   * Resize an object (points). With one side given, the other keeps the object's proportions. A
   * group's shapes scale with it; a chart redraws to fit. Word doesn't track size, so this applies
   * directly, and is refused where only tracked changes are allowed.
   */
  resizeObject(obj: DocObject, size: { width?: number; height?: number }): void {
    if (obj.vml) throw new EditError(OLD_VML);
    const range = this.objectRange(obj);
    const frame = frameOf(obj.drawing);
    const { width, height } = resized(frame, size);
    this.checkUntracked(range);
    resizeFrame(frame, width, height);
    dropFallback(range.start.node);
    this.changed();
  }

  /** Set an object's alt text (empty removes it). Not tracked, as in Word (see `resizeObject`). */
  setObjectAlt(obj: DocObject, alt: string): void {
    if (obj.vml) throw new EditError(OLD_VML);
    const range = this.objectRange(obj);
    const docPr = firstChild(frameOf(obj.drawing), 'wp:docPr');
    if (!docPr) throw new EditError('This object has no description to edit');
    this.checkUntracked(range);
    if (alt) docPr.attrs['descr'] = alt;
    else delete docPr.attrs['descr'];
    dropFallback(range.start.node);
    this.changed();
  }

  /** Delete an object. With Track Changes on the deletion is tracked: its run (and any text in it) is marked deleted. */
  deleteObject(obj: DocObject): void {
    const range = this.objectRange(obj);
    this.deleteRange(range);
    // What a tracked deletion keeps is the object, not Word's old copy of it.
    dropFallback(range.start.node);
  }

  /**
   * Move an object to `at`, keeping its size and placement (a floating object is anchored there
   * instead). With Track Changes on this is tracked as Word tracks a cut and paste: the original is
   * marked deleted and a copy inserted (a chart's copy gets its own chart part and workbook). Shapes,
   * text boxes and groups can't go inside a text box, and an object there can only be inline.
   */
  async moveObject(obj: DocObject, at: DocPosition): Promise<void> {
    if (obj.vml) throw new EditError(OLD_VML);
    const idx = this.index();
    const range = this.objectRange(obj, idx);
    const node = range.start.node;
    if (!idx.contains(at.node)) throw new EditError('That position is not in the document (stale position?)');
    // Right before or after itself: already there.
    if (at.node === node) return;
    for (let cur = idx.parent.get(at.node); cur; cur = idx.parent.get(cur)) if (cur === node) throw new EditError("An object can't move inside itself");
    if (inTextBox(idx, at.node)) {
      if (obj.kind === 'shape' || obj.kind === 'textBox' || obj.kind === 'group') throw new EditError(SHAPE_IN_BOX);
      if (frameOf(obj.drawing).name === 'wp:anchor') throw new EditError(INLINE_IN_BOX);
    }
    this.check('edit', range);
    if (this.trackChanges) {
      await this.replaceTracked(obj, node, at);
      return;
    }
    const run = this.objectRun(at, idx.parentOf(node));
    takeOut(this.index(), node);
    run.children.push(node);
    dropFallback(node);
    this.changed();
  }

  /**
   * Place an object inline, or floating at a position with a wrap style. Moving or re-wrapping a
   * floating object (what `place` leaves out of x, y, `from` and `wrap`, it keeps) isn't tracked by Word, so it applies
   * directly and is refused where only tracked changes are allowed. Switching between inline and
   * floating is tracked like a move when Track Changes is on: the original is marked deleted and a
   * copy inserted right after it.
   */
  async placeObject(obj: DocObject, place: ObjectPlacement): Promise<void> {
    if (obj.vml) throw new EditError(OLD_VML);
    const floating = frameOf(obj.drawing).name === 'wp:anchor';
    checkPlacement(place, floating);
    const idx = this.index();
    const range = this.objectRange(obj, idx);
    const node = range.start.node;
    const relativeHeight = () => nextRelativeHeight([this.doc.documentXml.root, this.root]);
    if ('inline' in place && !floating) return;
    if (!('inline' in place) && floating) {
      // Moving or re-wrapping a floating object: Word doesn't track it.
      this.checkUntracked(range);
    } else {
      // Switching between inline and floating.
      if (!floating && inTextBox(idx, node)) throw new EditError(INLINE_IN_BOX);
      this.check('edit', range);
      if (this.trackChanges) {
        await this.replaceTracked(obj, node, { node, offset: 1 }, (drawing) => placeFrame(drawing, place, relativeHeight));
        return;
      }
    }
    placeFrame(obj.drawing, place, relativeHeight);
    dropFallback(node);
    this.changed();
  }

  /** A new, empty run at `at` (a tracked insertion when tracking) formatted like `like`, for an object. */
  private objectRun(at: DocPosition, like: XmlElement): XmlElement {
    const run = this.insertRunContent(at, []);
    removeChildren(run, 'w:rPr');
    const rPr = firstChild(like, 'w:rPr');
    if (rPr) {
      run.children.unshift(cloneElement(rPr));
      cleanRunProps(run);
    }
    return run;
  }

  /**
   * A tracked move: a copy of the object (`node` is the run's child holding it), changed by
   * `change`, goes in at `at` as an insertion, with its own drawing id and chart part; the original
   * is marked deleted.
   */
  private async replaceTracked(obj: DocObject, node: XmlElement, at: DocPosition, change?: (drawing: XmlElement) => void): Promise<void> {
    const idx = this.index();
    const original = idx.parentOf(node);
    const { copy, drawing } = cloneObject(idx, node, obj.drawing);
    change?.(drawing);
    const run = this.objectRun(at, original);
    await this.copyChartParts([copy]);
    // Ids from the tree as it is now, with the new insertion in it.
    cleanCopy(copy, this.copyIds(), { comments: 0, notes: 0, bookmarks: 0 });
    run.children.push(copy);
    this.deleteRange({ start: { node, offset: 0 }, end: { node, offset: 1 } });
    dropFallback(node);
  }

  /** Add a footnote whose reference mark goes at a position. Returns the note's id. */
  async addFootnote(at: DocPosition, text: string): Promise<string> {
    this.check('edit', { start: at, end: at });
    if (this.isStory) throw new EditError('Footnotes can only be added to the main document');
    let part = this.doc.parts.footnotes;
    if (!part || !this.doc.xmlSync(part)) {
      part = 'word/footnotes.xml';
      for (let n = 2; this.doc.pkg.has(part); n++) part = `word/footnotes${n}.xml`;
      const separator = (type: string, id: string, mark: string) =>
        el('w:footnote', { 'w:type': type, 'w:id': id }, [el('w:p', {}, [el('w:pPr', {}, [el('w:spacing', { 'w:after': '0', 'w:line': '240', 'w:lineRule': 'auto' })]), el('w:r', {}, [el(mark)])])]);
      const xml = { declaration: null, prolog: [], epilog: [], root: el('w:footnotes', { ...ROOT_NAMESPACES }, [separator('separator', '-1', 'w:separator'), separator('continuationSeparator', '0', 'w:continuationSeparator')]) };
      this.doc.pkg.setXml(part, xml, CONTENT_TYPE.footnotes);
      await this.doc.pkg.addRelationship(this.doc.mainPart, REL.footnotes, part);
      await this.doc.refreshRels(this.doc.mainPart);
      this.doc.parts.footnotes = part;
    }
    const notes = this.doc.xmlSync(part)!.root;
    const id = String(Math.max(0, ...elements(notes, 'w:footnote').map((n) => Number(attr(n, 'w:id') ?? 0))) + 1);
    const styled = (styleId: string, fallback: XmlElement[]) => (this.doc.styles.get(styleId) ? [el('w:rStyle', { 'w:val': styleId })] : fallback);
    const refPr = () => el('w:rPr', {}, styled('FootnoteReference', [el('w:vertAlign', { 'w:val': 'superscript' })]));
    const pPr = this.doc.styles.get('FootnoteText') ? [el('w:pPr', {}, [el('w:pStyle', { 'w:val': 'FootnoteText' })])] : [el('w:pPr', {}, [el('w:spacing', { 'w:after': '0' })])];
    const textPr = this.doc.styles.get('FootnoteText') ? [] : [el('w:rPr', {}, [el('w:sz', { 'w:val': '20' })])];
    notes.children.push(
      el('w:footnote', { 'w:id': id }, [
        el('w:p', {}, [...pPr, el('w:r', {}, [refPr(), el('w:footnoteRef')]), el('w:r', {}, [...textPr, ...runContent(` ${text}`)])]),
      ]),
    );
    this.doc.touch(part);
    this.insertRunContent(at, [el('w:footnoteReference', { 'w:id': id })], refPr());
    return id;
  }

  // ---- headers and footers -----------------------------------------------------------------

  /** Headers and footers by section and page type, with their text ({PAGE} for page fields). */
  headersAndFooters(): HeaderFooterInfo[] {
    const sections = sectionProperties(this.mainBody());
    const out: HeaderFooterInfo[] = [];
    sections.forEach((sp, i) => {
      for (const kind of ['header', 'footer'] as const) {
        for (const ref of referencesOf(sp, kind)) {
          const part = this.doc.headerFooterPart(attr(ref, 'r:id') ?? '');
          if (!part) continue;
          const type = (attr(ref, 'w:type') ?? 'default') as HeaderFooterType;
          const known = out.find((h) => h.part === part && h.type === type && h.kind === kind);
          if (known) known.sections.push(i + 1);
          else out.push({ kind, type, sections: [i + 1], part, text: headerText(this.doc.xmlSync(part)?.root ?? el('w:hdr')) });
        }
      }
    });
    return out;
  }

  /** Give the document's sections a header (see `setHeaderFooter`). Returns the part name. */
  setHeader(text: string, options: HeaderFooterOptions = {}): Promise<string> {
    return this.setHeaderFooter('header', text, options);
  }

  /** Give the document's sections a footer (see `setHeaderFooter`). Returns the part name. */
  setFooter(text: string, options: HeaderFooterOptions = {}): Promise<string> {
    return this.setHeaderFooter('footer', text, options);
  }

  /**
   * Create or replace a header or footer. `text`: lines become paragraphs, tabs split it into
   * left / centre / right parts, and {PAGE}, {NUMPAGES}, {SECTIONPAGES} become page number fields
   * ("Page {PAGE} of {NUMPAGES}"). Every section by default, or one (`section`, from 1); all pages,
   * the first page or even pages (`type`). A part used only by these sections is rewritten in place,
   * otherwise a new part is made, so other sections keep theirs. With Track Changes on, the text
   * changes are recorded.
   */
  async setHeaderFooter(kind: HeaderFooterKind, text: string, options: HeaderFooterOptions = {}): Promise<string> {
    if (this.isStory) throw new EditError('Headers and footers are set from an editor of the main document');
    this.check('edit');
    const type = options.type ?? 'default';
    const body = this.mainBody();
    let sections = sectionProperties(body);
    if (!sections.length) {
      body.children.push(el('w:sectPr', {}, [el('w:pgSz', { 'w:w': '12240', 'w:h': '15840' }), el('w:pgMar', { 'w:top': '1440', 'w:right': '1440', 'w:bottom': '1440', 'w:left': '1440', 'w:header': '720', 'w:footer': '720', 'w:gutter': '0' })]));
      sections = sectionProperties(body);
    }
    const targets = this.sectionTargets(sections, options.section);
    const refOf = (sp: XmlElement) => referencesOf(sp, kind).find((r) => (attr(r, 'w:type') ?? 'default') === type);
    const partOf = (r: XmlElement | undefined) => (r ? this.doc.headerFooterPart(attr(r, 'r:id') ?? '') : undefined);
    // Rewrite in place only when the targets share one part that nothing else uses.
    const current = targets.map((sp) => partOf(refOf(sp)));
    const shared = current[0] && current.every((p) => p === current[0]) ? current[0] : undefined;
    const usedElsewhere =
      shared &&
      sections.some((sp) => {
        const refs = [...referencesOf(sp, 'header'), ...referencesOf(sp, 'footer')];
        return refs.some((r) => partOf(r) === shared && (!targets.includes(sp) || r.name !== (kind === 'header' ? 'w:headerReference' : 'w:footerReference') || (attr(r, 'w:type') ?? 'default') !== type));
      });
    const paragraphs = headerParagraphs(this.doc, kind, text, options.align, textWidth(targets[0]));
    const track = this.tracking();
    if (track) {
      for (const p of paragraphs) {
        const pPr = firstChild(p, 'w:pPr');
        const content = p.children.filter((c): c is XmlElement => c.type === 'element' && c !== pPr);
        p.children = [...(pPr ? [pPr] : []), ...(content.length ? [el('w:ins', insertedRuns([], track.nextId(), track.who).attrs, content)] : [])];
        markParagraph(p, 'ins', track.nextId(), track.who);
      }
    }
    let part: string;
    if (shared && !usedElsewhere && this.doc.xmlSync(shared)) {
      part = shared;
      const root = this.doc.xmlSync(part)!.root;
      if (track) {
        // The old text stays, marked deleted.
        const idx = new TreeIndex(root);
        for (const run of findAll(root, 'w:r')) deleteRun(run, (n) => idx.parent.get(n), track.nextId, track.who);
        for (const p of findAll(root, 'w:p')) markParagraph(p, 'del', track.nextId(), track.who);
        root.children.push(...paragraphs);
      } else root.children = paragraphs;
    } else {
      let n = 1;
      while (this.doc.pkg.has(`word/${kind}${n}.xml`)) n++;
      part = `word/${kind}${n}.xml`;
      const root = el(kind === 'header' ? 'w:hdr' : 'w:ftr', { ...ROOT_NAMESPACES }, paragraphs);
      this.doc.pkg.setXml(part, { declaration: null, prolog: [], root, epilog: [] }, CONTENT_TYPE[kind]);
      const rId = await this.doc.pkg.addRelationship(this.doc.mainPart, REL[kind], part);
      await this.doc.refreshRels(this.doc.mainPart);
      await this.doc.refreshRels(part);
      for (const sp of targets) {
        const ref = refOf(sp);
        if (ref) ref.attrs['r:id'] = rId;
        else insertOrdered(sp, el(kind === 'header' ? 'w:headerReference' : 'w:footerReference', { 'w:type': type, 'r:id': rId }), SECTPR_ORDER);
      }
      if (!this.root.attrs['xmlns:r']) this.root.attrs['xmlns:r'] = NS.r;
      await this.dropUnusedParts(current.filter((p): p is string => !!p));
    }
    if (type === 'first') for (const sp of targets) if (!firstChild(sp, 'w:titlePg')) insertOrdered(sp, el('w:titlePg'), SECTPR_ORDER);
    if (type === 'even') await this.setEvenAndOdd(true);
    this.doc.touch(part);
    this.changed();
    return part;
  }

  /** Remove headers (see `removeHeaderFooter`). Returns how many references were removed. */
  removeHeader(options: Omit<HeaderFooterOptions, 'type' | 'align'> & { type?: HeaderFooterType | 'all' } = {}): Promise<number> {
    return this.removeHeaderFooter('header', options);
  }

  /** Remove footers (see `removeHeaderFooter`). Returns how many references were removed. */
  removeFooter(options: Omit<HeaderFooterOptions, 'type' | 'align'> & { type?: HeaderFooterType | 'all' } = {}): Promise<number> {
    return this.removeHeaderFooter('footer', options);
  }

  /**
   * Remove headers or footers: of every type unless `type` says which, from every section unless
   * `section` says which (a later section without its own then shows the previous section's). Parts
   * nothing uses any more are deleted. Not recorded as a tracked change.
   */
  async removeHeaderFooter(kind: HeaderFooterKind, options: Omit<HeaderFooterOptions, 'type' | 'align'> & { type?: HeaderFooterType | 'all' } = {}): Promise<number> {
    if (this.isStory) throw new EditError('Headers and footers are removed from an editor of the main document');
    this.guard?.assert('edit');
    const type = options.type ?? 'all';
    const sections = sectionProperties(this.mainBody());
    const targets = this.sectionTargets(sections, options.section);
    const removed: string[] = [];
    let count = 0;
    let firstGone = false;
    let evenGone = false;
    for (const sp of targets) {
      for (const ref of referencesOf(sp, kind)) {
        const t = attr(ref, 'w:type') ?? 'default';
        if (type !== 'all' && t !== type) continue;
        sp.children.splice(sp.children.indexOf(ref), 1);
        const part = this.doc.headerFooterPart(attr(ref, 'r:id') ?? '');
        if (part) removed.push(part);
        if (t === 'first') firstGone = true;
        if (t === 'even') evenGone = true;
        count++;
      }
      if (firstGone && !hasFirstPage(sp)) removeChildren(sp, 'w:titlePg');
    }
    if (!count) return 0;
    if (evenGone && !hasEvenPages(sections)) await this.setEvenAndOdd(false);
    await this.dropUnusedParts(removed);
    this.changed();
    return count;
  }

  private mainBody(): XmlElement {
    const body = firstChild(this.doc.documentXml.root, 'w:body');
    if (!body) throw new EditError('document has no body');
    return body;
  }

  private sectionTargets(sections: XmlElement[], section: number | undefined): XmlElement[] {
    if (section === undefined) return sections;
    const sp = sections[section - 1];
    if (!sp) throw new EditError(`There is no section ${section}; the document has ${sections.length}`);
    return [sp];
  }

  /** Delete header/footer parts (and their relationships) that no section refers to any more. */
  private async dropUnusedParts(parts: string[]): Promise<void> {
    const sections = sectionProperties(this.mainBody());
    const inUse = new Set<string>();
    for (const sp of sections) for (const r of [...referencesOf(sp, 'header'), ...referencesOf(sp, 'footer')]) inUse.add(attr(r, 'r:id') ?? '');
    for (const part of new Set(parts)) {
      const rels = [...this.doc.rels(this.doc.mainPart)].filter(([, rel]) => rel.target === part);
      if (rels.some(([id]) => inUse.has(id))) continue;
      for (const [id] of rels) await this.doc.pkg.removeRelationship(this.doc.mainPart, id);
      if (this.doc.pkg.has(part)) this.doc.pkg.removePart(part);
    }
    await this.doc.refreshRels(this.doc.mainPart);
  }

  private async setEvenAndOdd(on: boolean): Promise<void> {
    const { name, xml } = await ensurePart(this.doc, 'settings');
    const has = !!firstChild(xml.root, 'w:evenAndOddHeaders');
    if (on === has) return;
    if (on) insertOrdered(xml.root, el('w:evenAndOddHeaders'), SETTINGS_ORDER);
    else removeChildren(xml.root, 'w:evenAndOddHeaders');
    this.doc.touch(name);
  }

  /** Join a paragraph with the one right after it (Backspace at the start of the second). */
  mergeParagraphs(first: XmlElement, second: XmlElement): void {
    const firstRange = this.paragraphRange(first);
    const secondRange = this.paragraphRange(second);
    this.check('edit', { start: firstRange?.end ?? { node: first, offset: 0 }, end: secondRange?.start ?? { node: second, offset: 0 } });
    const track = this.tracking();
    const mark = paragraphMarkRevision(first);
    if (track && !(mark?.kind === 'ins' && attr(mark.el, 'w:author') === this.author)) {
      // Deleting a paragraph mark is recorded; the paragraphs join when the change is accepted.
      markParagraph(first, 'del', track.nextId(), track.who);
      this.changed();
      return;
    }
    const idx = this.index();
    const parent = idx.parentOf(first);
    if (idx.parentOf(second) !== parent) throw new EditError('Cannot join paragraphs in different table cells');
    const a = indexOfChild(parent, first);
    const b = indexOfChild(parent, second);
    if (parent.children.slice(a + 1, b).some((n) => n.type === 'element' && (n.name === 'w:p' || n.name === 'w:tbl' || n.name === 'w:sdt'))) {
      throw new EditError('The paragraphs are not next to each other');
    }
    const secondPPr = firstChild(second, 'w:pPr');
    const sectPr = firstChild(secondPPr, 'w:sectPr');
    if (sectPr) {
      // Keep the section break that ended at the second paragraph.
      const pPr = ensureChild(first, 'w:pPr');
      if (first.children[0] !== pPr) {
        first.children.splice(first.children.indexOf(pPr), 1);
        first.children.unshift(pPr);
      }
      removeChildren(pPr, 'w:sectPr');
      insertOrdered(pPr, sectPr, PPR_ORDER);
    }
    first.children.push(...second.children.filter((n) => !(n.type === 'element' && n.name === 'w:pPr')));
    parent.children.splice(indexOfChild(parent, second), 1);
    this.changed();
  }

  /** Split a paragraph at a position (Enter). Returns the new, second paragraph. */
  splitParagraph(at: DocPosition): XmlElement {
    const blockEdge = this.bypassLocks ? undefined : this.blockLockEdge(at);
    if (blockEdge) {
      // Enter at the start or end of a locked block adds a paragraph outside it.
      this.guard?.assert('edit', { start: at, end: at });
      const idx = this.index();
      this.dropStaleFallbacks([at.node], idx);
      const pPr = firstChild(blockEdge.paragraph, 'w:pPr');
      const newPPr = pPr ? structuredClone(pPr) : undefined;
      if (newPPr) removeChildren(newPPr, 'w:sectPr');
      const para = el('w:p', {}, newPPr ? [newPPr] : []);
      const parent = idx.parentOf(blockEdge.sdt);
      parent.children.splice(indexOfChild(parent, blockEdge.sdt) + (blockEdge.after ? 1 : 0), 0, para);
      this.changed();
      return para;
    }
    this.check('edit', { start: at, end: at });
    let para: XmlElement;
    let mark: XmlElement;
    let idx: TreeIndex;
    if (at.node.name === 'w:p') {
      para = at.node;
      mark = el('docxkeun:split');
      const pPr = firstChild(para, 'w:pPr');
      para.children.splice(pPr ? para.children.indexOf(pPr) + 1 : 0, 0, mark);
      idx = this.index();
    } else {
      const index = this.index();
      const marks = isolate(index, { start: at, end: at });
      idx = this.index();
      replaceMark(idx, marks.endMark, []);
      mark = marks.startMark;
      idx = this.index();
      para = idx.ancestor(mark, 'w:p')!;
      // A split inside a hyperlink or content control happens after it.
      for (let container = idx.parentOf(mark); container !== para; container = idx.parentOf(mark)) {
        container.children.splice(indexOfChild(container, mark), 1);
        const outer = idx.parentOf(container);
        outer.children.splice(indexOfChild(outer, container) + 1, 0, mark);
        idx = this.index();
      }
    }
    const tail = para.children.splice(indexOfChild(para, mark));
    tail.shift();
    const pPr = firstChild(para, 'w:pPr');
    const newPPr = pPr ? structuredClone(pPr) : undefined;
    // The section break stays with the paragraph that now ends the section.
    if (pPr && firstChild(pPr, 'w:sectPr')) removeChildren(pPr, 'w:sectPr');
    const tailHasContent = tail.some((n) => n.type === 'element' && (findAll(n, 'w:t').some((t) => textContent(t) !== '') || findAll(n, 'w:tab').length > 0 || findAll(n, 'w:drawing').length > 0));
    if (newPPr && !tailHasContent) {
      // Enter at the end of a heading continues with the style's "next" style.
      const styleId = attr(firstChild(newPPr, 'w:pStyle'), 'w:val');
      const next = this.doc.styles.get(styleId)?.next;
      if (styleId && next && next !== styleId) {
        removeChildren(newPPr, 'w:pStyle');
        if (this.doc.styles.get(next)) insertOrdered(newPPr, el('w:pStyle', { 'w:val': next }), PPR_ORDER);
        removeChildren(newPPr, 'w:numPr');
      }
    }
    const newPara = el('w:p', {}, [...(newPPr ? [newPPr] : []), ...tail]);
    const parent = idx.parentOf(para);
    parent.children.splice(indexOfChild(parent, para) + 1, 0, newPara);
    const track = this.tracking();
    // Enter inserts a paragraph mark: the first half's mark is the new one.
    if (track) {
      const newPr = firstChild(newPara, 'w:pPr');
      const newMark = firstChild(newPr, 'w:rPr');
      if (newMark) for (const n of ['w:ins', 'w:del']) removeChildren(newMark, n);
      markParagraph(para, 'ins', track.nextId(), track.who);
    }
    this.changed();
    return newPara;
  }

  /** Append text at the end of a paragraph. */
  appendText(paragraph: XmlElement, text: string, format?: TextFormat): XmlElement {
    const range = this.paragraphRange(paragraph);
    this.check('edit', range);
    const runs = elements(paragraph, 'w:r');
    const run = newRun(text, runs[runs.length - 1], format);
    paragraph.children.push(this.placed(run, this.tracking()));
    this.changed();
    return run;
  }

  /** Delete a range (it may span paragraphs, but not table cells). */
  deleteRange(range: DocRange): void {
    this.check('edit', range);
    const track = this.tracking();
    if (track) {
      const { startMark, endMark, idx } = this.trackedCut(range, track);
      const holders = new Set([idx.parentOf(startMark), idx.parentOf(endMark)]);
      replaceMark(idx, startMark, []);
      replaceMark(this.index(), endMark, []);
      // Deleting all of the author's own insertion takes it out, w:ins and all.
      const after = this.index();
      for (const holder of holders) {
        const parent = after.parent.get(holder);
        if (parent && (holder.name === 'w:ins' || holder.name === 'w:moveTo') && !elements(holder).length) parent.children.splice(parent.children.indexOf(holder), 1);
      }
      this.changed();
      return;
    }
    // Deleting inside one text node (backspace): edit it in place.
    if (range.start.node === range.end.node && range.start.node.name === 'w:t') {
      const current = textContent(range.start.node);
      const a = Math.min(range.start.offset, range.end.offset);
      const b = Math.max(range.start.offset, range.end.offset);
      range.start.node.children = [{ type: 'text', value: current.slice(0, a) + current.slice(b) }];
      this.changed();
      return;
    }
    this.removeRange(range);
    this.changed();
  }

  /**
   * Replace a range with text formatted like the range's first run, or with pieces of text that are
   * each formatted like their own `like` run (for example to rewrite a sentence and keep which words
   * were bold). `format` is applied on top. Objects in the range (pictures, charts, shapes, text
   * boxes) go with the old text, as when a user selects and types over them; with `keepObjects` they
   * stay, right after the new text in their original order (not marked deleted when tracking).
   * Returns how many objects were kept.
   */
  replaceRange(range: DocRange, text: string | TextPiece[], format?: TextFormat, options: { keepObjects?: boolean } = {}): number {
    this.check('edit', range);
    const index = this.index();
    const template = index.parentOf(range.start.node);
    // Built before the cut, from the runs as they are now.
    const runs =
      typeof text === 'string'
        ? text || format
          ? [newRun(text, template, format)]
          : []
        : text.filter((p) => p.text).map((p) => newRun(p.text, p.like ?? template, format || p.format ? { ...format, ...p.format } : undefined));
    const track = this.tracking();
    const keep: KeptObjects | undefined = options.keepObjects ? { runs: [], count: 0, nextId: track?.nextId ?? nextRevisionId(this.root) } : undefined;
    if (track) {
      // The old text stays as a deletion; the new text follows it as an insertion.
      const { startMark, endMark } = this.trackedCut(range, track, keep);
      replaceMark(this.index(), startMark, []);
      this.liftOutOfInsertion(endMark, this.index());
      replaceMark(this.index(), endMark, [...(runs.length ? [insertedRuns(runs, track.nextId(), track.who)] : []), ...(keep?.runs ?? [])]);
      this.changed();
      return keep?.count ?? 0;
    }
    const { startMark, endMark, idx } = this.cut(range, keep);
    replaceMark(idx, startMark, [...runs, ...(keep?.runs ?? [])]);
    replaceMark(idx, endMark, []);
    this.changed();
    return keep?.count ?? 0;
  }

  /**
   * Take the objects (see OBJECT_NODES) between two marks out of their runs, for `replaceRange`
   * with keepObjects: they come back in `keep.runs`, in a run like the one they were in, and still in
   * a tracked insertion when they were in one. Objects already marked deleted, and anything inside an
   * object (a text box's own runs), are left alone.
   */
  private takeObjects(idx: TreeIndex, startMark: XmlElement, endMark: XmlElement, keep: KeptObjects): void {
    for (const run of between(idx, startMark, endMark, 'w:r')) {
      if (idx.ancestor(run, 'w:r') || idx.ancestor(run, 'w:del') || idx.ancestor(run, 'w:moveFrom')) continue;
      const objects = elements(run).filter((c) => OBJECT_NODES.has(c.name));
      if (!objects.length) continue;
      const rPr = firstChild(run, 'w:rPr');
      const copy = el('w:r', {}, [...(rPr ? [cloneElement(rPr)] : []), ...objects]);
      const revision = runRevision(run, (n) => idx.parent.get(n));
      for (const node of objects) takeOut(idx, node);
      keep.count += objects.length;
      keep.runs.push(
        revision && (revision.name === 'w:ins' || revision.name === 'w:moveTo')
          ? el('w:ins', { 'w:id': keep.nextId(), 'w:author': attr(revision, 'w:author') ?? '', ...(attr(revision, 'w:date') ? { 'w:date': attr(revision, 'w:date')! } : {}) }, [copy])
          : copy,
      );
    }
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
  fillTemplate(
    data: Record<string, string | number | boolean | null | undefined>,
    pattern = /\{\{\s*([\w.-]+)\s*\}\}/g,
    options: { includeLocked?: boolean } = {},
  ): string[] {
    const used = new Set<string>();
    // Back to front, so earlier positions stay valid.
    for (const m of this.find(pattern).reverse()) {
      const key = m.groups[0] ?? '';
      if (!(key in data)) continue;
      if (!options.includeLocked && this.isLocked(m)) continue;
      this.withLocks(!!options.includeLocked, () => this.replaceRange(m, String(data[key] ?? '')));
      used.add(key);
    }
    for (const key of this.fillContentControls(data, options)) used.add(key);
    return [...used];
  }

  private withLocks<T>(bypass: boolean, fn: () => T): T {
    const saved = this.bypassLocks;
    this.bypassLocks = saved || bypass;
    try {
      return fn();
    } finally {
      this.bypassLocks = saved;
    }
  }

  /** Set the text of content controls (w:sdt) by tag or title. Locked ones are skipped unless `includeLocked`. */
  fillContentControls(data: Record<string, string | number | boolean | null | undefined>, options: { includeLocked?: boolean } = {}): string[] {
    const used: string[] = [];
    const idx = this.index();
    for (const sdt of findAll(this.body, 'w:sdt')) {
      const sdtPr = firstChild(sdt, 'w:sdtPr');
      const key = attr(firstChild(sdtPr, 'w:tag'), 'w:val') ?? attr(firstChild(sdtPr, 'w:alias'), 'w:val');
      if (key === undefined || !(key in data)) continue;
      const content = firstChild(sdt, 'w:sdtContent');
      if (!content) continue;
      if (!options.includeLocked && !this.bypassLocks && (locksContent(lockModeOf(sdt)) || this.lockAround(idx, sdt))) continue;
      this.check('fillForm');
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
    const track = this.tracking();
    for (const run of between(idx, startMark, endMark, 'w:r')) {
      if (track && !this.inOwnInsertion(run, idx)) recordRunFormat(run, track.nextId(), track.who);
      applyFormat(run, format);
    }
    replaceMark(idx, startMark, []);
    replaceMark(idx, endMark, []);
    this.changed();
  }

  /** Remove direct character formatting (and character styles, except Hyperlink) from a range. */
  clearFormatting(range: DocRange): void {
    this.check('format', range);
    const index = this.index();
    const { startMark, endMark } = isolate(index, range);
    const idx = this.index();
    const track = this.tracking();
    for (const run of between(idx, startMark, endMark, 'w:r')) {
      if (track && !this.inOwnInsertion(run, idx)) recordRunFormat(run, track.nextId(), track.who);
      const rPr = firstChild(run, 'w:rPr');
      if (!rPr) continue;
      const keep = rPr.children.filter((c) => c.type === 'element' && ((c.name === 'w:rStyle' && attr(c, 'w:val') === 'Hyperlink') || c.name === 'w:rPrChange'));
      rPr.children = keep;
      if (!keep.length) run.children.splice(run.children.indexOf(rPr), 1);
    }
    replaceMark(idx, startMark, []);
    replaceMark(idx, endMark, []);
    this.changed();
  }

  /** Remove the content of a range, returning the sentinels at its (now adjacent) ends. */
  private cut(range: DocRange, keep?: KeptObjects): { startMark: XmlElement; endMark: XmlElement; idx: TreeIndex } {
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
    if (keep) {
      this.takeObjects(idx, startMark, endMark, keep);
      idx = this.index();
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
    const ref = typeof where === 'object' ? (where.before ?? where.after) : undefined;
    this.check('edit', ref ? { start: { node: ref, offset: 0 }, end: { node: ref, offset: 0 } } : undefined);
    const pPr = options.style ? [el('w:pPr', {}, [el('w:pStyle', { 'w:val': options.style })])] : [];
    const track = this.tracking();
    const para = el('w:p', {}, [...pPr, ...(text ? [this.placed(newRun(text, undefined, options.format), track)] : [])]);
    if (track) markParagraph(para, 'ins', track.nextId(), track.who);
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
    this.check('edit', this.paragraphRange(paragraph) ?? { start: { node: paragraph, offset: 0 }, end: { node: paragraph, offset: 0 } });
    const track = this.tracking();
    const mark = paragraphMarkRevision(paragraph);
    const ownNew = mark?.kind === 'ins' && attr(mark.el, 'w:author') === this.author;
    if (track && !ownNew) {
      // The text is marked deleted and so is the paragraph mark; Word joins them away on accept.
      const tidx = this.index();
      for (const run of findAll(paragraph, 'w:r')) deleteRun(run, (n) => tidx.parent.get(n), track.nextId, track.who);
      markParagraph(paragraph, 'del', track.nextId(), track.who);
      this.changed();
      return;
    }
    const idx = this.index();
    const parent = idx.parentOf(paragraph);
    // A table cell, and a text box, must keep at least one paragraph.
    if ((parent.name === 'w:tc' || parent.name === 'w:txbxContent') && elements(parent, 'w:p').length === 1) {
      paragraph.children = paragraph.children.filter((c) => c.type === 'element' && c.name === 'w:pPr');
    } else parent.children.splice(indexOfChild(parent, paragraph), 1);
    this.changed();
  }

  // ---- copies and tables -----------------------------------------------------------------------

  /** The block (paragraph, table or content control) of the body that holds `node`. */
  private topBlock(node: XmlElement, idx = this.index()): XmlElement {
    let n = node;
    for (let p = idx.parent.get(n); p && p !== this.body; p = idx.parent.get(p)) n = p;
    if (idx.parent.get(n) !== this.body) throw new EditError('That paragraph is not in the body of the document');
    return n;
  }

  /** Put new blocks into the body before/after the block holding `ref`, or at the start or end. */
  private placeBlocks(blocks: XmlElement[], where: { before?: XmlElement; after?: XmlElement } | 'start' | 'end'): void {
    const body = this.body;
    const sect = firstChild(body, 'w:sectPr');
    let at: number;
    if (where === 'start') at = 0;
    else if (where === 'end') at = sect ? body.children.indexOf(sect) : body.children.length;
    else {
      const ref = where.before ?? where.after;
      if (!ref) throw new EditError('Pass the paragraph to place it before or after');
      const top = this.topBlock(ref);
      at = body.children.indexOf(top) + (where.after ? 1 : 0);
    }
    body.children.splice(at, 0, ...blocks);
    // Word wants a paragraph after a table that ends the document.
    const last = blocks[blocks.length - 1];
    const next = body.children[at + blocks.length];
    if (last?.name === 'w:tbl' && (!next || (next.type === 'element' && next.name === 'w:sectPr'))) body.children.splice(at + blocks.length, 0, el('w:p'));
  }

  /** Permission to add content next to `where`. */
  private checkPlace(where: { before?: XmlElement; after?: XmlElement } | 'start' | 'end'): void {
    const ref = typeof where === 'object' ? (where.before ?? where.after) : undefined;
    this.check('edit', ref ? { start: { node: ref, offset: 0 }, end: { node: ref, offset: 0 } } : undefined);
  }

  /**
   * Copy the body's blocks from the one holding `from` to the one holding `to` (a table is copied
   * whole) and put the copies at `where`, with their formatting, lists, tables and pictures. Lists
   * in the copy number from the start again unless `restartLists` is false, and `pageBreak` starts
   * the copy on a new page. What may exist only once (bookmarks, comment anchors, note references)
   * is left out of the copy and counted in `dropped`.
   */
  async copyBlocks(
    from: XmlElement,
    to: XmlElement,
    where: { before?: XmlElement; after?: XmlElement } | 'start' | 'end',
    options: { pageBreak?: boolean; restartLists?: boolean } = {},
  ): Promise<{ blocks: XmlElement[]; dropped: CopyDropped }> {
    this.checkPlace(where);
    const idx = this.index();
    const kids = this.body.children;
    const first = kids.indexOf(this.topBlock(from, idx));
    const last = kids.indexOf(this.topBlock(to, idx));
    if (last < first) throw new EditError('The last paragraph to copy comes before the first');
    const originals = kids.slice(first, last + 1).filter((c): c is XmlElement => c.type === 'element' && c.name !== 'w:sectPr');
    const copies = originals.map(cloneElement);

    const ids = this.copyIds();
    const dropped: CopyDropped = { comments: 0, notes: 0, bookmarks: 0 };
    for (const c of copies) cleanCopy(c, ids, dropped);

    // Lists in the copy: a new list of the same look for each list copied, so numbering restarts.
    if (options.restartLists ?? true) {
      const styleNumId = (s: string) => this.doc.styles.get(s)?.pPr.numId;
      const renumbered = new Map<string, string>();
      for (const p of copies.flatMap((c) => (c.name === 'w:p' ? [c] : findAll(c, 'w:p')))) {
        const pPr = firstChild(p, 'w:pPr');
        const numPr = firstChild(pPr, 'w:numPr');
        const styleId = attr(firstChild(pPr, 'w:pStyle'), 'w:val');
        const numId = numPr ? attr(firstChild(numPr, 'w:numId'), 'w:val') : styleId ? this.doc.styles.get(styleId)?.pPr.numId : undefined;
        if (!numId || numId === '0') continue;
        let fresh = renumbered.get(numId);
        if (!fresh) {
          const abstractId = this.doc.numbering.abstractFor(numId, styleNumId)?.id;
          fresh = await createList(this.doc, { format: 'keep', levels: new Set(), base: { numId, abstractId } });
          renumbered.set(numId, fresh);
        }
        const ilvl = numPr ? attr(firstChild(numPr, 'w:ilvl'), 'w:val') ?? '0' : String(styleId ? (this.doc.styles.get(styleId)?.pPr.ilvl ?? 0) : 0);
        const target = ensureChild(p, 'w:pPr');
        if (p.children[0] !== target) {
          p.children.splice(p.children.indexOf(target), 1);
          p.children.unshift(target);
        }
        removeChildren(target, 'w:numPr');
        insertOrdered(target, el('w:numPr', {}, [el('w:ilvl', { 'w:val': ilvl }), el('w:numId', { 'w:val': fresh })]), PPR_ORDER);
      }
    }

    const blocks = [...copies];
    if (options.pageBreak) {
      const head = blocks[0];
      if (head?.name === 'w:p') {
        const pPr = ensureChild(head, 'w:pPr');
        if (head.children[0] !== pPr) {
          head.children.splice(head.children.indexOf(pPr), 1);
          head.children.unshift(pPr);
        }
        removeChildren(pPr, 'w:pageBreakBefore');
        insertOrdered(pPr, el('w:pageBreakBefore'), PPR_ORDER);
      } else blocks.unshift(el('w:p', {}, [el('w:pPr', {}, [el('w:pageBreakBefore')])]));
    }
    await this.copyChartParts(copies);
    const track = this.tracking();
    if (track) for (const b of blocks) markInserted(b, track.nextId, track.who, (p) => markParagraph(p, 'ins', track.nextId(), track.who));
    this.placeBlocks(blocks, where);
    this.changed();
    return { blocks, dropped };
  }

  /**
   * Insert a table of text (`rows` of cell texts; "\n" in a cell starts a new paragraph) before or
   * after the block holding a paragraph, or at the start or end. It spans the text width.
   */
  insertTable(where: { before?: XmlElement; after?: XmlElement } | 'start' | 'end', rows: string[][], options: TableOptions = {}): XmlElement {
    this.checkPlace(where);
    const columns = Math.max(0, ...rows.map((r) => r.length));
    if (!rows.length || !columns) throw new EditError('A table needs at least one row and one column');
    if (rows.length > 500 || columns > 63) throw new EditError('A table can have at most 500 rows and 63 columns');
    if (options.widths && (options.widths.length !== columns || options.widths.some((w) => !(w > 0)))) {
      throw new EditError(`widths needs ${columns} positive numbers, one per column`);
    }
    const sections = sectionProperties(this.body);
    const total = textWidth(sections[sections.length - 1]);
    const grid = this.doc.styles.get('TableGrid') ? 'TableGrid' : undefined;
    const style = options.style ?? (options.borders === false ? undefined : grid);
    if (options.style && !this.doc.styles.get(options.style)) throw new EditError(`There is no table style "${options.style}"`);
    const track = this.tracking();
    const table = buildTable(rows, columnWidths(columns, total, options.widths), {
      header: !!options.header,
      style,
      borders: options.borders !== false,
      align: options.align,
      paragraph: (text, header) =>
        el('w:p', {}, [
          // Without a table style, cells get Table Grid's tight spacing.
          ...(style ? [] : [el('w:pPr', {}, [el('w:spacing', { 'w:before': '0', 'w:after': '0' })])]),
          ...(text ? [newRun(text, undefined, header ? { bold: true } : undefined)] : []),
        ]),
    });
    if (track) markInserted(table, track.nextId, track.who, (p) => markParagraph(p, 'ins', track.nextId(), track.who));
    this.placeBlocks([table], where);
    this.changed();
    return table;
  }

  /** The table holding a paragraph (the innermost one). */
  tableOf(paragraph: XmlElement): XmlElement {
    const idx = this.index();
    const table = idx.ancestor(paragraph, 'w:tbl');
    if (!table) throw new EditError('That paragraph is not in a table');
    return table;
  }

  /** Permission to change a table: its whole text. */
  private checkTable(table: XmlElement): void {
    const paras = findAll(table, 'w:p');
    const a = this.paragraphRange(paras[0]!);
    const b = this.paragraphRange(paras[paras.length - 1]!);
    this.check('edit', a && b ? { start: a.start, end: b.end } : { start: { node: table, offset: 0 }, end: { node: table, offset: 0 } });
  }

  /**
   * Insert rows into a table before row `at` (0-based; the row count appends). Each new row copies
   * the formatting of the row next to it and gets the texts of one entry of `rows`.
   */
  insertTableRows(table: XmlElement, at: number, rows: string[][]): XmlElement[] {
    this.checkTable(table);
    const existing = elements(table, 'w:tr');
    if (!Number.isInteger(at) || at < 0 || at > existing.length) throw new EditError(`Rows go from 0 to ${existing.length}`);
    if (!rows.length) throw new EditError('Pass at least one row');
    const template = existing[at - 1] ?? existing[at];
    if (!template) throw new EditError('The table has no rows to copy the formatting of');
    const cells = elements(template, 'w:tc');
    const widest = Math.max(...rows.map((r) => r.length));
    if (widest > cells.length) throw new EditError(`That row has ${cells.length} cells, so a new row can have at most ${cells.length} texts`);
    const track = this.tracking();
    const made = rows.map((texts) => {
      const row = cloneElement(template);
      delete row.attrs['w14:paraId'];
      delete row.attrs['w14:textId'];
      // Only the first rows of a table repeat as a header.
      const trPr = firstChild(row, 'w:trPr');
      if (trPr && at > 0) {
        removeChildren(trPr, 'w:tblHeader');
        removeChildren(trPr, 'w:ins');
        removeChildren(trPr, 'w:del');
      }
      elements(row, 'w:tc').forEach((tc, i) => {
        const tcPr = firstChild(tc, 'w:tcPr');
        if (tcPr) removeChildren(tcPr, 'w:vMerge');
        const p = firstChild(tc, 'w:p') ?? el('w:p');
        const pPr = firstChild(p, 'w:pPr');
        const runTemplate = findAll(p, 'w:r')[0];
        const text = texts[i] ?? '';
        const lines = text.split('\n');
        const paragraphs = lines.map((line) => el('w:p', {}, [...(pPr ? [cloneElement(pPr)] : []), ...(line ? [newRun(line, runTemplate)] : [])]));
        tc.children = [...(tcPr ? [tcPr] : []), ...paragraphs];
      });
      if (track) markInserted(row, track.nextId, track.who, (p) => markParagraph(p, 'ins', track.nextId(), track.who));
      return row;
    });
    const anchor = existing[at];
    table.children.splice(anchor ? table.children.indexOf(anchor) : table.children.length, 0, ...made);
    this.changed();
    return made;
  }

  /** Delete rows `from` to `to` (0-based, inclusive) of a table; deleting them all deletes the table. */
  deleteTableRows(table: XmlElement, from: number, to = from): void {
    this.checkTable(table);
    const rows = elements(table, 'w:tr');
    if (!Number.isInteger(from) || !Number.isInteger(to) || from < 0 || to >= rows.length || to < from) throw new EditError(`Rows go from 0 to ${rows.length - 1}`);
    if (from === 0 && to === rows.length - 1) {
      this.deleteTable(table);
      return;
    }
    const doomed = rows.slice(from, to + 1);
    const track = this.tracking();
    if (track) for (const row of doomed) this.markRowDeleted(row, track);
    else table.children = table.children.filter((c) => !(c.type === 'element' && doomed.includes(c)));
    this.changed();
  }

  /** Delete a whole table. */
  deleteTable(table: XmlElement): void {
    this.checkTable(table);
    const track = this.tracking();
    if (track) for (const row of elements(table, 'w:tr')) this.markRowDeleted(row, track);
    else {
      const idx = this.index();
      const parent = idx.parentOf(table);
      parent.children.splice(indexOfChild(parent, table), 1);
      // A table cell, a text box, and the body, must keep a paragraph.
      if ((parent.name === 'w:tc' || parent.name === 'w:txbxContent') && !elements(parent, 'w:p').length) parent.children.push(el('w:p'));
      if (parent === this.body && !elements(parent).some((c) => c.name === 'w:p' || c.name === 'w:tbl' || c.name === 'w:sdt')) {
        const sect = firstChild(parent, 'w:sectPr');
        parent.children.splice(sect ? parent.children.indexOf(sect) : parent.children.length, 0, el('w:p'));
      }
    }
    this.changed();
  }

  private markRowDeleted(row: XmlElement, track: { who: RevisionAuthor; nextId: () => string }): void {
    const idx = this.index();
    for (const run of findAll(row, 'w:r')) deleteRun(run, (n) => idx.parent.get(n), track.nextId, track.who);
    for (const p of findAll(row, 'w:p')) markParagraph(p, 'del', track.nextId(), track.who);
    markRow(row, 'del', track.nextId(), track.who);
  }

  setParagraphStyle(paragraph: XmlElement, styleId: string): void {
    this.check('format', this.paragraphRange(paragraph));
    const track = this.tracking();
    if (track) recordParagraphFormat(paragraph, track.nextId(), track.who, PPR_ORDER);
    const pPr = ensureChild(paragraph, 'w:pPr');
    if (paragraph.children[0] !== pPr) {
      paragraph.children.splice(paragraph.children.indexOf(pPr), 1);
      paragraph.children.unshift(pPr);
    }
    removeChildren(pPr, 'w:pStyle');
    insertOrdered(pPr, el('w:pStyle', { 'w:val': styleId }), PPR_ORDER);
    this.changed();
  }

  /** Change paragraph formatting: alignment, direction, indents, spacing and keep rules. */
  formatParagraph(paragraph: XmlElement, format: ParagraphFormat): void {
    this.check('format', this.paragraphRange(paragraph));
    const track = this.tracking();
    if (track) recordParagraphFormat(paragraph, track.nextId(), track.who, PPR_ORDER);
    const pPr = ensureChild(paragraph, 'w:pPr');
    if (paragraph.children[0] !== pPr) {
      paragraph.children.splice(paragraph.children.indexOf(pPr), 1);
      paragraph.children.unshift(pPr);
    }
    const twips = (pt: number) => String(Math.round(pt * 20));
    const onOff = (name: string, on: boolean | undefined) => {
      if (on === undefined) return;
      removeChildren(pPr, name);
      insertOrdered(pPr, on ? el(name) : el(name, { 'w:val': '0' }), PPR_ORDER);
    };
    onOff('w:keepNext', format.keepWithNext);
    onOff('w:keepLines', format.keepLinesTogether);
    onOff('w:pageBreakBefore', format.pageBreakBefore);
    if (format.direction) onOff('w:bidi', format.direction === 'rtl');
    if (format.align) {
      removeChildren(pPr, 'w:jc');
      // In a right-to-left paragraph Word reads left and right as start and end, so a visually
      // right-aligned RTL paragraph is stored as "left".
      const bidi = firstChild(pPr, 'w:bidi');
      const rtl = !!bidi && attr(bidi, 'w:val') !== '0' && attr(bidi, 'w:val') !== 'false';
      const visual = format.align === 'justify' ? 'both' : format.align;
      const jc = rtl && visual === 'left' ? 'right' : rtl && visual === 'right' ? 'left' : visual;
      insertOrdered(pPr, el('w:jc', { 'w:val': jc }), PPR_ORDER);
    }
    if (format.indentLeft !== undefined || format.indentRight !== undefined || format.indentFirstLine !== undefined) {
      const ind = ensureOrdered(pPr, 'w:ind', PPR_ORDER);
      if (format.indentLeft !== undefined) {
        delete ind.attrs['w:start'];
        ind.attrs['w:left'] = twips(format.indentLeft);
      }
      if (format.indentRight !== undefined) {
        delete ind.attrs['w:end'];
        ind.attrs['w:right'] = twips(format.indentRight);
      }
      if (format.indentFirstLine !== undefined) {
        delete ind.attrs['w:firstLine'];
        delete ind.attrs['w:hanging'];
        if (format.indentFirstLine >= 0) ind.attrs['w:firstLine'] = twips(format.indentFirstLine);
        else ind.attrs['w:hanging'] = twips(-format.indentFirstLine);
      }
    }
    if (format.spaceBefore !== undefined || format.spaceAfter !== undefined || format.lineSpacing !== undefined) {
      const sp = ensureOrdered(pPr, 'w:spacing', PPR_ORDER);
      if (format.spaceBefore !== undefined) {
        delete sp.attrs['w:beforeAutospacing'];
        sp.attrs['w:before'] = twips(format.spaceBefore);
      }
      if (format.spaceAfter !== undefined) {
        delete sp.attrs['w:afterAutospacing'];
        sp.attrs['w:after'] = twips(format.spaceAfter);
      }
      if (format.lineSpacing !== undefined) {
        sp.attrs['w:line'] = String(Math.round(format.lineSpacing * 240));
        sp.attrs['w:lineRule'] = 'auto';
      }
    }
    this.changed();
  }

  // ---- lists --------------------------------------------------------------------------------

  /** The list paragraphs of the body with their labels ("1.", "a)", "•"), keyed by w:p. */
  listItems(): Map<XmlElement, ListItem> {
    return listItems(this.doc, this.paragraphs());
  }

  /**
   * Change list formatting of paragraphs:
   * - `format`: make them one list in that format ('bullet', 'decimal' 1., 'decimalParen' 1),
   *   'lowerLetter' a., 'upperLetter' A., 'lowerRoman' i., 'upperRoman' I.), numbered from `start`
   *   (default 1). Paragraphs already in a list keep its indentation and their level. 'none' takes
   *   them out of their list.
   * - `start` alone: restart the numbering, keeping the format.
   * - `level` (0–8): the list level (indent); alone, only the level changes.
   * - `sameListAs`: join the list of that paragraph, continuing its numbering.
   * Returns the numId the paragraphs now use (undefined after 'none').
   */
  async setList(paragraphs: XmlElement[], options: SetListOptions): Promise<string | undefined> {
    const { format, level, start, sameListAs } = options;
    if (paragraphs.length === 0) throw new EditError('setList needs at least one paragraph');
    if (level !== undefined && (!Number.isInteger(level) || level < 0 || level > 8)) throw new EditError('List levels go from 0 to 8');
    if (start !== undefined && (!Number.isInteger(start) || start < 0 || start > 32767)) throw new EditError('start must be a whole number from 0 to 32767');
    const all = new Map(this.paragraphs().map((p) => [p.el, p]));
    const paras = paragraphs.map((el) => {
      const p = all.get(el);
      if (!p) throw new EditError('setList only works on paragraphs of the document body');
      return p;
    });
    for (const p of paras) this.check('format', this.paragraphRange(p) ?? { start: { node: p.el, offset: 0 }, end: { node: p.el, offset: 0 } });

    if (format === 'none') {
      if (sameListAs || start !== undefined || level !== undefined) throw new EditError("format 'none' cannot be combined with other list options");
      for (const p of paras) this.writeNumPr(p, undefined);
      this.changed();
      return undefined;
    }
    if (sameListAs) {
      const target = all.get(sameListAs)?.numbering;
      if (!target) throw new EditError('The paragraph to join is not in a list yet: make it one first (with a format), then join the others to it');
      if (start !== undefined) throw new EditError("sameListAs continues that list's numbering; to number from a given value, use start without sameListAs");
      if (format !== undefined) {
        // Naming the list's own format is harmless; another format would start a different list.
        const styleNumId = (s: string) => this.doc.styles.get(s)?.pPr.numId;
        const lvl = this.doc.numbering.level(target.numId, level ?? target.ilvl, styleNumId);
        if (!lvl || !levelMatches(lvl, format)) {
          throw new EditError(`That list is ${lvl ? describeLevel(lvl) : 'in another format'}; to continue it leave out format, or to start a new list use format without sameListAs`);
        }
      }
      for (const p of paras) this.writeNumPr(p, { numId: target.numId, ilvl: level ?? (p.numbering ? p.numbering.ilvl : target.ilvl) });
      this.changed();
      return target.numId;
    }
    const first = paras.find((p) => p.numbering)?.numbering;
    if (format === undefined) {
      if (!first) throw new EditError('These paragraphs are not in a list; pass a format to make them one');
      if (start === undefined) {
        if (level === undefined) throw new EditError('Nothing to change: pass format, level, start or sameListAs');
        for (const p of paras) {
          if (!p.numbering) throw new EditError('Some of these paragraphs are not in a list; pass a format to make them one');
          this.writeNumPr(p, { numId: p.numbering.numId, ilvl: level });
        }
        this.changed();
        return first.numId;
      }
    }
    const ilvlOf = (p: DocParagraph) => level ?? p.numbering?.ilvl ?? 0;
    const styleNumId = (s: string) => this.doc.styles.get(s)?.pPr.numId;
    const base = first ? { numId: first.numId, abstractId: this.doc.numbering.abstractFor(first.numId, styleNumId)?.id } : undefined;
    const numId = await createList(this.doc, {
      format: format ?? 'keep',
      levels: new Set(paras.map(ilvlOf)),
      ...(start !== undefined ? { start } : {}),
      ...(base ? { base } : {}),
    });
    for (const p of paras) this.writeNumPr(p, { numId, ilvl: ilvlOf(p) });
    this.changed();
    return numId;
  }

  /** Set or clear a paragraph's w:numPr (clearing style-given numbering with numId 0). */
  private writeNumPr(para: DocParagraph, num: { numId: string; ilvl: number } | undefined): void {
    const p = para.el;
    const track = this.tracking();
    if (track) recordParagraphFormat(p, track.nextId(), track.who, PPR_ORDER);
    const pPr = ensureChild(p, 'w:pPr');
    if (p.children[0] !== pPr) {
      p.children.splice(p.children.indexOf(pPr), 1);
      p.children.unshift(pPr);
    }
    removeChildren(pPr, 'w:numPr');
    if (num) {
      insertOrdered(pPr, el('w:numPr', {}, [el('w:ilvl', { 'w:val': String(num.ilvl) }), el('w:numId', { 'w:val': num.numId })]), PPR_ORDER);
      return;
    }
    const styleNum = para.styleId ? this.doc.styles.get(para.styleId)?.pPr.numId : undefined;
    if (styleNum && styleNum !== '0') insertOrdered(pPr, el('w:numPr', {}, [el('w:numId', { 'w:val': '0' })]), PPR_ORDER);
  }

  // ---- hyperlinks and bookmarks -------------------------------------------------------------

  /** Turn a range (within one paragraph) into a hyperlink to a URL or a bookmark. */
  async addHyperlink(range: DocRange, target: { url: string } | { bookmark: string }, tooltip?: string): Promise<XmlElement> {
    this.check('link', range);
    if ('url' in target && !/^(https?|mailto|tel|ftp):/i.test(target.url)) throw new EditError(`Refusing link to "${target.url}": unsupported scheme`);
    const attrs: Record<string, string> = {};
    if ('url' in target) {
      attrs['r:id'] = await this.doc.pkg.addRelationship(this.part, REL.hyperlink, target.url, true);
      await this.doc.refreshRels(this.part);
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

  /** Point an existing w:hyperlink at another URL or a bookmark (same scheme rules as `addHyperlink`). */
  async setHyperlinkTarget(hyperlink: XmlElement, target: { url: string } | { bookmark: string }): Promise<void> {
    this.check('link');
    if ('url' in target && !/^(https?|mailto|tel|ftp):/i.test(target.url)) throw new EditError(`Refusing link to "${target.url}": unsupported scheme`);
    this.dropStaleFallbacks([hyperlink]);
    const old = hyperlink.attrs['r:id'];
    if ('url' in target) {
      hyperlink.attrs['r:id'] = await this.doc.pkg.addRelationship(this.part, REL.hyperlink, target.url, true);
      delete hyperlink.attrs['w:anchor'];
      if (!this.root.attrs['xmlns:r']) this.root.attrs['xmlns:r'] = NS.r;
    } else {
      hyperlink.attrs['w:anchor'] = target.bookmark;
      delete hyperlink.attrs['r:id'];
    }
    // The old tooltip described the old target.
    delete hyperlink.attrs['w:tooltip'];
    delete hyperlink.attrs['w:docLocation'];
    await this.dropHyperlinkRel(old);
    this.changed();
  }

  /** Remove a hyperlink relationship that nothing in this part refers to any more. */
  private async dropHyperlinkRel(id: string | undefined): Promise<void> {
    if (id) {
      const used = findAll(this.root, 'w:hyperlink').some((h) => h.attrs['r:id'] === id);
      const rel = this.doc.rels(this.part).get(id);
      if (!used && rel && rel.external) await this.doc.pkg.removeRelationship(this.part, id);
    }
    await this.doc.refreshRels(this.part);
  }

  /** Set the alt text (wp:docPr/@descr) of a picture or drawing. Not tracked, as in Word (see `setObjectAlt`). */
  setImageAlt(drawing: XmlElement, text: string): void {
    const props = drawing.name === 'wp:docPr' ? drawing : findAll(drawing, 'wp:docPr')[0];
    if (!props) throw new EditError('This object has no description to edit');
    const node = this.runChildHolding(props, this.index());
    this.checkUntracked({ start: { node, offset: 0 }, end: { node, offset: 1 } });
    props.attrs['descr'] = text;
    dropFallback(node);
    this.changed();
  }

  /** Unwrap a w:hyperlink, keeping its text. */
  async removeHyperlink(link: XmlElement): Promise<void> {
    this.check('link');
    const idx = this.index();
    this.dropStaleFallbacks([link], idx);
    const parent = idx.parentOf(link);
    const runs = link.children;
    for (const run of findAll(link, 'w:r')) {
      const rPr = firstChild(run, 'w:rPr');
      const rStyle = firstChild(rPr, 'w:rStyle');
      if (attr(rStyle, 'w:val') === 'Hyperlink') removeChildren(rPr!, 'w:rStyle');
    }
    parent.children.splice(indexOfChild(parent, link), 1, ...runs);
    this.changed();
    await this.dropHyperlinkRel(attr(link, 'r:id'));
  }

  /** Every w:hyperlink in the body with its target. */
  hyperlinks(): Array<{ el: XmlElement; url?: string; anchor?: string; tooltip?: string; text: string }> {
    const rels = this.doc.rels(this.part);
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
    const nodes = [start, ...findAll(this.root, 'w:bookmarkEnd').filter((b) => attr(b, 'w:id') === id)];
    this.dropStaleFallbacks(nodes, idx);
    for (const node of nodes) {
      const parent = idx.parentOf(node);
      parent.children.splice(indexOfChild(parent, node), 1);
    }
    this.changed();
  }

  // ---- editable ranges (exceptions to document protection) ---------------------------------

  /** Let a user (w:ed, e.g. an e-mail or DOMAIN\\user) or a group (w:edGrp, e.g. "everyone") edit a range. */
  addEditableRange(range: DocRange, who: { editor?: string; group?: 'everyone' | 'editors' | 'owners' | 'administrators' | 'contributors' | 'current' | 'none' | string }): string {
    this.check('protect', range);
    if (!who.editor && !who.group) throw new EditError('An editable range needs an editor or a group');
    const id = String(1 + Math.max(0, ...findAll(this.root, 'w:permStart').map((n) => Number(attr(n, 'w:id') ?? 0))));
    const index = this.index();
    const { startMark, endMark } = isolate(index, range);
    const idx = this.index();
    replaceMark(idx, startMark, [el('w:permStart', { 'w:id': id, 'w:ed': who.editor, 'w:edGrp': who.editor ? undefined : who.group })]);
    replaceMark(idx, endMark, [el('w:permEnd', { 'w:id': id })]);
    this.changed();
    return id;
  }

  editableRanges(): Array<{ id: string; editor?: string; group?: string }> {
    return findAll(this.root, 'w:permStart').map((n) => ({ id: attr(n, 'w:id') ?? '', editor: attr(n, 'w:ed'), group: attr(n, 'w:edGrp') }));
  }

  removeEditableRange(id: string): void {
    this.check('protect');
    const idx = this.index();
    for (const node of [...findAll(this.root, 'w:permStart'), ...findAll(this.root, 'w:permEnd')]) {
      if (attr(node, 'w:id') !== id) continue;
      const parent = idx.parentOf(node);
      parent.children.splice(indexOfChild(parent, node), 1);
    }
    this.changed();
  }

  // ---- locked regions (content controls that cannot be edited) ------------------------------

  /** The locked inline region whose first or last character is at `pos`. */
  private lockEdge(pos: DocPosition): { sdt: XmlElement; after: boolean } | undefined {
    const idx = this.index();
    if (!idx.contains(pos.node) || pos.node.name === 'w:p') return undefined;
    const sdt = this.lockAround(idx, pos.node);
    const content = firstChild(sdt, 'w:sdtContent');
    if (!sdt || !content || firstChild(content, 'w:p')) return undefined;
    const items = findAll(content, 'w:r').flatMap((r) => elements(r).filter((c) => c.name !== 'w:rPr'));
    const first = items[0];
    const last = items[items.length - 1];
    if (pos.node === first && pos.offset === 0) return { sdt, after: false };
    const end = last?.name === 'w:t' ? textContent(last).length : 1;
    if (pos.node === last && pos.offset >= end) return { sdt, after: true };
    return undefined;
  }

  /** A locked block of paragraphs that `pos` sits at the very start or end of. */
  private blockLockEdge(pos: DocPosition): { sdt: XmlElement; paragraph: XmlElement; after: boolean } | undefined {
    const idx = this.index();
    if (!idx.contains(pos.node)) return undefined;
    const paragraph = pos.node.name === 'w:p' ? pos.node : idx.ancestor(pos.node, 'w:p');
    if (!paragraph) return undefined;
    const content = idx.parent.get(paragraph);
    const sdt = content?.name === 'w:sdtContent' ? idx.parent.get(content) : undefined;
    if (!sdt || !locksContent(lockModeOf(sdt))) return undefined;
    const paras = elements(content!, 'w:p');
    const range = this.paragraphRange(paragraph);
    const atStart = !range || (pos.node === range.start.node && pos.offset === range.start.offset);
    const atEnd = !range || (pos.node === range.end.node && pos.offset === range.end.offset);
    if (atEnd && paragraph === paras[paras.length - 1]) return { sdt, paragraph, after: true };
    if (atStart && paragraph === paras[0]) return { sdt, paragraph, after: false };
    return undefined;
  }

  /**
   * Lock a range against editing. Inside one paragraph the exact text is locked; a range that
   * spans paragraphs locks those whole paragraphs. Stored as a Word content control, so Word
   * enforces the lock too. Returns the region's id.
   */
  lockRange(range: DocRange, options: { title?: string; tag?: string; mode?: LockMode } = {}): string {
    this.check('protect', range);
    const mode = options.mode ?? 'content';
    const used = new Set(findAll(this.root, 'w:id').map((n) => attr(n, 'w:val')));
    let id: string;
    do id = String(1 + Math.floor(Math.random() * 0x7ffffffe));
    while (used.has(id));
    const sdtPr = el('w:sdtPr', {}, [
      ...(options.title ? [el('w:alias', { 'w:val': options.title })] : []),
      el('w:tag', { 'w:val': options.tag ?? LOCK_TAG }),
      el('w:id', { 'w:val': id }),
      el('w:lock', { 'w:val': LOCK_VALUES[mode] }),
    ]);
    let idx = this.index();
    const paraOf = (node: XmlElement) => (node.name === 'w:p' ? node : idx.ancestor(node, 'w:p'));
    const startPara = paraOf(range.start.node);
    const endPara = paraOf(range.end.node);
    if (!startPara || !endPara) throw new EditError('The range is not inside a paragraph');
    if (startPara === endPara && range.start.node.name !== 'w:p') {
      const { startMark, endMark } = isolate(idx, range);
      idx = this.index();
      const container = idx.parentOf(startMark);
      if (idx.parentOf(endMark) !== container) {
        replaceMark(idx, startMark, []);
        replaceMark(idx, endMark, []);
        throw new EditError('A locked region inside a paragraph must not cut across a link or field');
      }
      const a = indexOfChild(container, startMark);
      const b = indexOfChild(container, endMark);
      const inner = container.children.slice(a + 1, b);
      if (!inner.length) {
        container.children.splice(a, b - a + 1);
        throw new EditError('Select some text to lock');
      }
      container.children.splice(a, b - a + 1, el('w:sdt', {}, [sdtPr, el('w:sdtContent', {}, inner)]));
    } else {
      const parent = idx.parentOf(startPara);
      if (idx.parentOf(endPara) !== parent) throw new EditError('A locked region must not cross table cells');
      const a = indexOfChild(parent, startPara);
      const b = indexOfChild(parent, endPara);
      const [lo, hi] = a <= b ? [a, b] : [b, a];
      const blocks = parent.children.slice(lo, hi + 1);
      parent.children.splice(lo, hi - lo + 1, el('w:sdt', {}, [sdtPr, el('w:sdtContent', {}, blocks)]));
    }
    this.changed();
    return id;
  }

  /** Lock whole paragraphs (from `first` to `last`). */
  lockParagraphs(first: XmlElement, last: XmlElement = first, options: { title?: string; tag?: string; mode?: LockMode } = {}): string {
    return this.lockRange({ start: { node: first, offset: 0 }, end: { node: last, offset: 0 } }, options);
  }

  /** Every locked content control in the body. */
  lockedRegions(): LockedRegion[] {
    return findAll(this.body, 'w:sdt')
      .filter((sdt) => lockModeOf(sdt))
      .map((sdt) => {
        const sdtPr = firstChild(sdt, 'w:sdtPr');
        const content = firstChild(sdt, 'w:sdtContent');
        return {
          id: attr(firstChild(sdtPr, 'w:id'), 'w:val') ?? '',
          mode: lockModeOf(sdt)!,
          title: sdtTitle(sdt),
          tag: attr(firstChild(sdtPr, 'w:tag'), 'w:val'),
          text: content ? findAll(content, 'w:p').length ? findAll(content, 'w:p').map((p) => textContent(p)).join('\n') : textContent(content) : '',
          block: !!firstChild(content, 'w:p'),
          el: sdt,
        };
      });
  }

  /**
   * Remove a lock. Regions docxkeun created are unwrapped completely; other content controls
   * keep their control and only lose the lock.
   */
  unlock(id: string): void {
    this.check('protect');
    const region = this.lockedRegions().find((r) => r.id === id);
    if (!region) throw new EditError(`No locked region with id ${id}`);
    this.dropStaleFallbacks([region.el]);
    if (region.tag === LOCK_TAG) {
      const idx = this.index();
      const parent = idx.parentOf(region.el);
      parent.children.splice(indexOfChild(parent, region.el), 1, ...(firstChild(region.el, 'w:sdtContent')?.children ?? []));
    } else {
      removeChildren(firstChild(region.el, 'w:sdtPr')!, 'w:lock');
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
    this.check('editComment', undefined, { commentAuthor: attr(comment, 'w:author') });
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

/** The child `name` of `parent`, created in schema order when missing. */
function ensureOrdered(parent: XmlElement, name: string, order: readonly string[]): XmlElement {
  const existing = firstChild(parent, name);
  if (existing) return existing;
  const node = el(name);
  insertOrdered(parent, node, order);
  return node;
}
