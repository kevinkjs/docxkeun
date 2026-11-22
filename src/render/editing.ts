/**
 * Typing directly on the rendered pages: a caret and selection drawn over the SVG, keyboard and
 * IME input through a hidden textarea, and every change made through DocxEditor (so access
 * guards apply and unknown markup survives). The page is laid out again after each change.
 *
 * Positions are kept as (paragraph index, character offset) so they survive re-layout and
 * the run splitting that edits cause.
 */
import { DocxEditor, OBJECT_CHAR, paragraphsOf, paragraphText as plainText, positionAt, type EditGuard, type ParagraphFormat, type TextFormat } from '../api/editor.js';
import type { ListFormat } from '../api/lists.js';
import { htmlToParagraphs, paragraphsToHtml, type PastedParagraph } from './clipboard.js';
import { selectionInfoFor, type SelectionInfo } from '../api/selection.js';
import type { DocPosition } from '../api/positions.js';
import type { BodyLine, ImageItem, Page, StoryLine, TextItem } from '../layout/types.js';
import type { DocParagraph, TextInline } from '../model/content.js';
import { cloneElement, elements, findAll, type XmlElement } from '../xml/index.js';
import { UndoHistory, type UndoEntry } from './undo.js';
import type { DocxViewer } from './viewer.js';

export interface EditingOptions {
  /** Checks every change (e.g. an AccessPolicy from docxkeun/security). */
  guard?: EditGuard;
  /** Record typing as tracked changes (default: the document's Track Changes setting). */
  trackChanges?: boolean;
  /** Author of tracked changes. */
  author?: string;
}

/** A place in the text: paragraph index (document order, tables included) and UTF-16 offset. */
export interface Caret {
  para: number;
  char: number;
}

/** Which text the caret is in: the body, a header or footer, a note, or a text box. */
export type EditStory =
  | { kind: 'body' }
  | { kind: 'header' | 'footer'; part: string }
  | { kind: 'footnote' | 'endnote'; part: string; noteId: string }
  | { kind: 'textbox'; container: XmlElement };

type StoryEntry = Extract<UndoEntry, { kind: 'story' }>;

function sameStory(a: EditStory, b: EditStory): boolean {
  if (a.kind !== b.kind) return false;
  if (a.kind === 'textbox') return a.container === (b as typeof a).container;
  if (a.kind === 'footnote' || a.kind === 'endnote') return a.part === (b as typeof a).part && a.noteId === (b as typeof a).noteId;
  if (a.kind === 'header' || a.kind === 'footer') return a.part === (b as typeof a).part;
  return true;
}

/** Which w:txbxContent of `root` a text-box story is (undefined for other stories or if it is gone). */
function textboxIndex(root: XmlElement, story: EditStory): number | undefined {
  if (story.kind !== 'textbox') return undefined;
  const i = findAll(root, 'w:txbxContent').indexOf(story.container);
  return i < 0 ? undefined : i;
}

/** The story again after the XML trees were replaced: a text box by its index, else the body. */
function resolveStory(root: XmlElement, story: EditStory, index: number | undefined): EditStory {
  if (story.kind !== 'textbox') return story;
  const container = index === undefined ? undefined : findAll(root, 'w:txbxContent')[index];
  return container ? { kind: 'textbox', container } : { kind: 'body' };
}

function storyOfLine(line: StoryLine): EditStory | undefined {
  switch (line.story) {
    case 'header':
    case 'footer':
      return line.part ? { kind: line.story, part: line.part } : undefined;
    case 'footnote':
    case 'endnote':
      return line.part && line.noteId ? { kind: line.story, part: line.part, noteId: line.noteId } : undefined;
    case 'textbox':
      return line.container ? { kind: 'textbox', container: line.container } : undefined;
  }
}

const within = (l: BodyLine, x: number, y: number, slack = 2) => x >= l.x - slack && x <= l.x + l.w + slack && y >= l.y - slack && y <= l.y + l.h + slack;

interface ScreenPos {
  page: number;
  x: number;
  y: number;
  h: number;
}

/** Inline pictures and shapes count as one character (U+FFFC) for the caret. */
const paragraphText = (p: DocParagraph) => plainText(p, { objects: true });

/** Formatting at the caret, for toolbars (see the viewer's `caretchange` event). */
export interface CaretFormat {
  bold: boolean;
  italic: boolean;
  underline: boolean;
  strike: boolean;
  /** Points. */
  size?: number;
  font?: string;
  color?: string;
  highlight?: string;
  align: 'left' | 'center' | 'right' | 'justify';
  direction: 'ltr' | 'rtl';
  styleId?: string;
  list?: 'bullet' | 'number';
  story: EditStory['kind'];
  hasSelection: boolean;
  trackChanges: boolean;
  canUndo: boolean;
  canRedo: boolean;
}

type LineItem = TextItem | ImageItem;

const isMac = typeof navigator !== 'undefined' && /Mac|iPhone|iPad/.test(navigator.platform);

function compare(a: Caret, b: Caret): number {
  return a.para - b.para || a.char - b.char;
}

export class InlineEditing {
  private enabled = true;
  private storyValue: EditStory = { kind: 'body' };
  private anchor: Caret = { para: 0, char: 0 };
  private focusPos: Caret = { para: 0, char: 0 };
  private hasCaret = false;
  private readonly input: HTMLTextAreaElement;
  private composing = false;
  private dragging = false;
  private goalX: number | undefined;
  private readonly history = new UndoHistory();
  private historyDoc: unknown;
  private lastEdit = { kind: '', at: 0 };
  /** Formatting chosen with no selection: applied to the next text typed. */
  private pending: TextFormat | undefined;
  private lastFormat = '';
  private paras: DocParagraph[] = [];
  private readonly nodeIndex = new Map<XmlElement, { para: number; start: number }>();
  private readonly paraIndex = new Map<XmlElement, number>();
  private readonly layers = new Map<number, HTMLDivElement>();
  private readonly cleanup: Array<() => void> = [];

  private readonly options: EditingOptions;

  constructor(private readonly viewer: DocxViewer, options: EditingOptions = {}) {
    this.options = { ...options };
    const root = viewer.root;
    this.input = document.createElement('textarea');
    this.input.className = 'docxkeun-input';
    this.input.setAttribute('aria-label', 'Document text');
    this.input.setAttribute('autocapitalize', 'off');
    this.input.setAttribute('autocomplete', 'off');
    this.input.spellcheck = false;
    root.appendChild(this.input);
    root.classList.add('docxkeun-editing');

    const listen = <K extends keyof HTMLElementEventMap>(target: HTMLElement | Document, type: K, fn: (e: HTMLElementEventMap[K]) => void) => {
      target.addEventListener(type, fn as EventListener);
      this.cleanup.push(() => target.removeEventListener(type, fn as EventListener));
    };
    listen(root, 'mousedown', (e) => this.onMouseDown(e));
    listen(document, 'mousemove', (e) => this.onMouseMove(e));
    listen(document, 'mouseup', () => (this.dragging = false));
    listen(root, 'dblclick', (e) => this.onDoubleClick(e));
    listen(this.input, 'keydown', (e) => this.onKeyDown(e));
    listen(this.input, 'input', () => this.onInput());
    listen(this.input, 'compositionstart', () => (this.composing = true));
    listen(this.input, 'compositionend', (e) => {
      this.composing = false;
      this.input.value = '';
      if (e.data) this.typeText(e.data);
    });
    listen(this.input, 'copy', (e) => this.onCopy(e, false));
    listen(this.input, 'cut', (e) => this.onCopy(e, true));
    listen(this.input, 'paste', (e) => this.onPaste(e));
    listen(this.input, 'focus', () => root.classList.add('docxkeun-focused'));
    listen(this.input, 'blur', () => root.classList.remove('docxkeun-focused'));
    this.cleanup.push(viewer.on('layout', () => this.onLayout()));
    this.cleanup.push(viewer.on('pagemount', ({ index, element }) => this.onPageMount(index, element)));
    if (viewer.layout) this.onLayout();
  }

  setEnabled(on: boolean): void {
    this.enabled = on;
    this.viewer.root.classList.toggle('docxkeun-editing', on);
    if (!on) {
      this.input.blur();
      this.hasCaret = false;
    }
    this.redraw();
  }

  get isEnabled(): boolean {
    return this.enabled;
  }

  /** The current selection (anchor and focus; equal when it is just a caret). */
  get selection(): { anchor: Caret; focus: Caret } | undefined {
    return this.hasCaret ? { anchor: { ...this.anchor }, focus: { ...this.focusPos } } : undefined;
  }

  /** Place the caret (and optionally extend a selection) programmatically. */
  select(anchor: Caret, focus: Caret = anchor): void {
    this.anchor = this.clamp(anchor);
    this.focusPos = this.clamp(focus);
    this.hasCaret = true;
    this.input.focus({ preventScroll: true });
    this.redraw(true);
  }

  focus(): void {
    this.input.focus({ preventScroll: true });
  }

  /** The selection (or caret) with paragraph numbers, text and pages, for AI assistants and toolbars. */
  selectionInfo(): SelectionInfo | undefined {
    const doc = this.doc;
    if (!doc || !this.hasCaret || !this.enabled) return undefined;
    const s = this.storyValue;
    const part = s.kind === 'header' || s.kind === 'footer' || s.kind === 'footnote' || s.kind === 'endnote' ? s.part : undefined;
    return selectionInfoFor(doc, this.paras, this.anchor, this.focusPos, { story: s.kind, ...(part ? { part } : {}), ...(this.viewer.layout ? { layout: this.viewer.layout } : {}) });
  }

  /** Whether typing is recorded as tracked changes. */
  get trackChanges(): boolean {
    return this.doc ? new DocxEditor(this.doc, this.options.guard, { trackChanges: this.options.trackChanges }).trackChanges : false;
  }

  set trackChanges(on: boolean) {
    this.options.trackChanges = on;
    this.lastFormat = '';
    this.reportFormat();
  }

  /** The story the caret is in. */
  get story(): EditStory {
    return this.storyValue;
  }

  /** Move the caret into a header, footer, note or text box (or back to the body). */
  enterStory(story: EditStory, caret: Caret = { para: 0, char: 0 }): void {
    this.storyValue = story;
    this.refreshModel();
    this.viewer.root.dataset.editStory = story.kind;
    this.goalX = undefined;
    this.select(caret);
    this.viewer.emit('storychange', { story });
  }

  /** The guard checking this layer's edits, if any. */
  get guard(): EditGuard | undefined {
    return this.options.guard;
  }

  /** Who new comments and replies are written by: the `author` option, else the guard's user. */
  get author(): string {
    const user = (this.options.guard as { user?: { name?: string; id?: string } } | undefined)?.user;
    return this.options.author ?? user?.name ?? user?.id ?? 'docxkeun';
  }

  /** Put the caret at a point on a page (points), as a click would. No-op while editing is off. */
  placeCaretAt(page: number, x: number, y: number): void {
    if (!this.enabled) return;
    const caret = this.hitTest(page, x, y);
    if (caret) this.select(caret);
  }

  /** Whether a point on a page (points) lies inside the current, non-empty selection. */
  selectionContains(page: number, x: number, y: number): boolean {
    if (!this.enabled || !this.hasCaret || compare(this.anchor, this.focusPos) === 0) return false;
    const caret = this.hitTest(page, x, y);
    if (!caret) return false;
    const [start, end] = this.ordered();
    return compare(caret, start) >= 0 && compare(caret, end) <= 0;
  }

  /** Where the start of the selection is: its page and the line's top (points on the page). */
  selectionSpot(): { page: number; x: number; y: number; h: number } | undefined {
    if (!this.hasCaret) return undefined;
    return this.screen(this.ordered()[0]);
  }

  /** An editor range (for DocxEditor calls such as addComment) between two carets. */
  rangeOf(a: Caret, b: Caret): { start: DocPosition; end: DocPosition } {
    const [start, end] = compare(a, b) <= 0 ? [a, b] : [b, a];
    return { start: this.toDoc(start, 'start'), end: this.toDoc(end, 'end') };
  }

  destroy(): void {
    for (const fn of this.cleanup) fn();
    this.input.remove();
    for (const layer of this.layers.values()) layer.remove();
    this.viewer.root.classList.remove('docxkeun-editing', 'docxkeun-focused');
  }

  // ---- model ---------------------------------------------------------------------------------

  private get doc() {
    return this.viewer.document;
  }

  /** Part holding the current story. */
  private get storyPart(): string {
    const s = this.storyValue;
    return s.kind === 'body' || s.kind === 'textbox' ? this.doc!.mainPart : s.part;
  }

  private editor(): DocxEditor {
    return new DocxEditor(this.doc!, this.options.guard, { part: this.storyPart, trackChanges: this.options.trackChanges, author: this.options.author });
  }

  private storyParagraphs(): DocParagraph[] {
    const doc = this.doc;
    if (!doc) return [];
    const s = this.storyValue;
    switch (s.kind) {
      case 'body':
        return paragraphsOf(doc.body().blocks);
      case 'header':
      case 'footer':
        return doc.xmlSync(s.part) ? paragraphsOf(doc.headerFooter(s.part).blocks) : [];
      case 'footnote':
      case 'endnote':
        return paragraphsOf(doc.note(s.kind, s.noteId));
      case 'textbox':
        return paragraphsOf(doc.readBlocks(doc.mainPart, s.container));
    }
  }

  private refreshModel(): void {
    const doc = this.doc;
    this.nodeIndex.clear();
    this.paraIndex.clear();
    this.paras = this.storyParagraphs();
    // A story that disappeared (its text box was deleted): back to the body.
    if (!this.paras.length && this.storyValue.kind !== 'body') {
      this.storyValue = { kind: 'body' };
      this.paras = doc ? paragraphsOf(doc.body().blocks) : [];
    }
    this.paras.forEach((p, i) => {
      this.paraIndex.set(p.el, i);
      for (const seg of paragraphText(p).segments) this.nodeIndex.set(seg.node, { para: i, start: seg.start });
    });
  }

  private textOf(para: number): string {
    const p = this.paras[para];
    return p ? paragraphText(p).text : '';
  }

  private clamp(c: Caret): Caret {
    if (!this.paras.length) return { para: 0, char: 0 };
    const para = Math.max(0, Math.min(c.para, this.paras.length - 1));
    return { para, char: Math.max(0, Math.min(c.char, this.textOf(para).length)) };
  }

  /** XML position for a caret. 'end' attaches to the text before (it inherits that formatting). */
  private toDoc(c: Caret, side: 'start' | 'end'): DocPosition {
    const para = this.paras[c.para]!;
    const pt = paragraphText(para);
    if (!pt.segments.length) return { node: para.el, offset: 0 };
    const k = Math.max(0, Math.min(c.char, pt.text.length));
    const pos = positionAt(pt, k, side) ?? positionAt(pt, k, side === 'start' ? 'end' : 'start');
    if (pos) return pos;
    const last = pt.segments[pt.segments.length - 1]!;
    return { node: last.node, offset: last.node.name === 'w:t' ? last.end - last.start : 1 };
  }

  private fromDoc(pos: DocPosition): Caret | undefined {
    if (pos.node.name === 'w:p') {
      const i = this.paraIndex.get(pos.node);
      return i === undefined ? undefined : { para: i, char: 0 };
    }
    const e = this.nodeIndex.get(pos.node);
    if (!e) return undefined;
    return { para: e.para, char: e.start + (pos.node.name === 'w:t' ? pos.offset : pos.offset > 0 ? 1 : 0) };
  }

  // ---- geometry ------------------------------------------------------------------------------

  private pages(): Page[] {
    return this.viewer.layout?.pages ?? [];
  }

  /** Lines of the current story on a page. */
  private bodyLines(page: Page): BodyLine[] {
    return (this.storyValue.kind === 'body' ? page.lines : page.storyLines).filter((l) => this.paraIndex.has(l.paragraph));
  }

  private linesOf(page: Page): BodyLine[] {
    return this.storyValue.kind === 'body' ? page.lines : page.storyLines;
  }

  /** The run child an inline image stands for, when it is in the current story. */
  private imageNode(item: ImageItem): XmlElement | undefined {
    const inline = item.inline;
    if (!inline || inline.float) return undefined;
    for (const c of inline.run.children) if (c.type === 'element' && this.nodeIndex.has(c) && (c === inline.drawing || findAll(c, inline.drawing.name).includes(inline.drawing))) return c;
    return undefined;
  }

  private lineItems(page: Page, line: BodyLine): LineItem[] {
    const out: Array<{ item: LineItem; x: number }> = [];
    for (const i of page.items) {
      if (i.type === 'text') {
        if (!i.source || !this.nodeIndex.has(i.source.inline.node)) continue;
        if (i.y > line.y - 0.5 && i.y < line.y + line.h + 0.5 && i.xs[0]! >= line.x - 1 && i.xs[0]! <= line.x + line.w + 1) out.push({ item: i, x: i.xs[0]! });
      } else if (i.type === 'image' && this.imageNode(i)) {
        const bottom = i.y + i.h;
        if (bottom > line.y - 0.5 && bottom < line.y + line.h + 0.5 && i.x >= line.x - 1 && i.x <= line.x + line.w + 1) out.push({ item: i, x: i.x });
      }
    }
    return out.sort((a, b) => a.x - b.x).map((e) => e.item);
  }

  private clusterLength(inline: TextInline, offset: number): number {
    const cp = inline.text.codePointAt(offset) ?? 0;
    return cp > 0xffff ? 2 : 1;
  }

  /** Caret for a point on a page (in points). */
  private hitTest(pageIndex: number, x: number, y: number): Caret | undefined {
    const page = this.pages()[pageIndex];
    if (!page) return undefined;
    const lines = this.bodyLines(page);
    if (!lines.length) return undefined;
    const within = lines.filter((l) => x >= l.x - 6 && x <= l.x + l.w + 6);
    const pool = within.length ? within : lines;
    const dist = (l: BodyLine) => (y < l.y ? l.y - y : y > l.y + l.h ? y - l.y - l.h : 0);
    const line = pool.reduce((best, l) => (dist(l) < dist(best) ? l : best));
    const items = this.lineItems(page, line);
    const para = this.paraIndex.get(line.paragraph)!;
    if (!items.length) {
      // No text on this line: an empty paragraph, or a line of tabs or images.
      const sameParaLater = lines.some((l) => l.paragraph === line.paragraph && l.y > line.y);
      return { para, char: sameParaLater ? 0 : this.textOf(para).length };
    }
    for (const item of items) {
      if (item.type === 'image') {
        if (x < item.x + item.w / 2) return this.fromDoc({ node: this.imageNode(item)!, offset: 0 });
        continue;
      }
      for (let i = 0; i < item.chars.length; i++) {
        if (x < item.xs[i]! + item.widths[i]! / 2) return this.itemCaret(item, i, true);
      }
    }
    const last = items[items.length - 1]!;
    if (last.type === 'image') return this.fromDoc({ node: this.imageNode(last)!, offset: 1 });
    return this.itemCaret(last, last.chars.length - 1, false);
  }

  private itemCaret(item: TextItem, i: number, before: boolean): Caret | undefined {
    const { inline, offsets } = item.source!;
    const node = inline.node;
    let offset: number;
    if (node.name !== 'w:t') offset = before ? 0 : 1;
    else offset = before ? offsets[i]! : offsets[i]! + this.clusterLength(inline, offsets[i]!);
    return this.fromDoc({ node, offset });
  }

  /** Where a caret is drawn. */
  private screen(c: Caret): ScreenPos | undefined {
    const len = this.textOf(c.para).length;
    const tries: DocPosition[] = [];
    if (c.char < len) tries.push(this.toDoc(c, 'start'));
    tries.push(this.toDoc(c, 'end'));
    if (c.char < len) tries.push(this.toDoc({ para: c.para, char: c.char + 1 }, 'start'));
    const pages = this.pages();
    for (const pos of tries) {
      for (const page of pages) {
        if (pos.node.name === 'w:p') {
          const line = this.linesOf(page).find((l) => l.paragraph === pos.node);
          if (line) return { page: page.index, x: line.x, y: line.y, h: line.h };
          continue;
        }
        for (const item of page.items) {
          if (item.type === 'image' && pos.node.name !== 'w:t' && item.inline && !item.inline.float && this.imageNode(item) === pos.node) {
            return { page: page.index, x: pos.offset ? item.x + item.w : item.x, y: item.y, h: item.h };
          }
          if (item.type !== 'text' || item.source?.inline.node !== pos.node) continue;
          const { offsets, inline } = item.source;
          const top = item.y - item.font.ascent * item.size;
          const h = (item.font.ascent + item.font.descent) * item.size;
          if (pos.node.name !== 'w:t') {
            return { page: page.index, x: pos.offset ? item.xs[0]! + item.widths[0]! : item.xs[0]!, y: top, h };
          }
          const i = offsets.indexOf(pos.offset);
          if (i >= 0) return { page: page.index, x: item.xs[i]!, y: top, h };
          const lastI = offsets.length - 1;
          const end = offsets[lastI]! + this.clusterLength(inline, offsets[lastI]!);
          if (pos.offset === end) return { page: page.index, x: item.xs[lastI]! + item.widths[lastI]!, y: top, h };
        }
      }
    }
    // An empty paragraph or one with only tabs: fall back to its first line.
    const el = this.paras[c.para]?.el;
    for (const page of pages) {
      const line = this.linesOf(page).find((l) => l.paragraph === el);
      if (line) return { page: page.index, x: line.x, y: line.y, h: line.h };
    }
    return undefined;
  }

  // ---- drawing -------------------------------------------------------------------------------

  private onLayout(): void {
    if (this.viewer.document !== this.historyDoc) {
      this.historyDoc = this.viewer.document;
      this.history.clear();
      this.lastEdit = { kind: '', at: 0 };
    }
    this.refreshModel();
    if (this.hasCaret) {
      this.anchor = this.clamp(this.anchor);
      this.focusPos = this.clamp(this.focusPos);
    }
    this.redraw();
  }

  private onPageMount(index: number, element: HTMLDivElement): void {
    const layer = document.createElement('div');
    layer.className = 'docxkeun-edit-layer';
    element.appendChild(layer);
    this.layers.set(index, layer);
    this.drawPage(index);
  }

  private redraw(scrollIntoView = false): void {
    for (const index of this.layers.keys()) this.drawPage(index);
    this.reportFormat();
    if (!this.hasCaret || !this.enabled) return;
    const pos = this.screen(this.focusPos);
    if (!pos) return;
    const pageEl = this.viewer.pageElement(pos.page);
    if (!pageEl) return;
    const s = this.viewer.scale;
    // Keep the hidden input at the caret so IME candidate windows appear in the right place.
    this.input.style.left = `${pageEl.offsetLeft + pos.x * s}px`;
    this.input.style.top = `${pageEl.offsetTop + pos.y * s}px`;
    if (scrollIntoView) {
      const root = this.viewer.root;
      const top = pageEl.offsetTop + pos.y * s;
      const bottom = top + pos.h * s;
      if (top < root.scrollTop + 8) root.scrollTop = top - 40;
      else if (bottom > root.scrollTop + root.clientHeight - 8) root.scrollTop = bottom - root.clientHeight + 40;
    }
  }

  private drawPage(index: number): void {
    const layer = this.layers.get(index);
    if (!layer || !layer.isConnected) {
      this.layers.delete(index);
      return;
    }
    layer.replaceChildren();
    if (!this.enabled || !this.hasCaret) return;
    const page = this.pages()[index];
    if (!page) return;
    const s = this.viewer.scale;
    const box = (cls: string, x: number, y: number, w: number | undefined, h: number) => {
      const el = document.createElement('div');
      el.className = cls;
      el.style.left = `${x * s}px`;
      el.style.top = `${y * s}px`;
      // The caret's width comes from the stylesheet (it must not scale to zero).
      if (w !== undefined) el.style.width = `${w * s}px`;
      el.style.height = `${h * s}px`;
      layer.appendChild(el);
    };
    const [start, end] = compare(this.anchor, this.focusPos) <= 0 ? [this.anchor, this.focusPos] : [this.focusPos, this.anchor];
    if (compare(start, end) !== 0) {
      for (const item of page.items) {
        if (item.type !== 'text' || !item.source) continue;
        const e = this.nodeIndex.get(item.source.inline.node);
        if (!e) continue;
        const isText = item.source.inline.node.name === 'w:t';
        let runStart = -1;
        const flush = (i: number) => {
          if (runStart < 0) return;
          const x = item.xs[runStart]!;
          box('docxkeun-selection', x, item.y - item.font.ascent * item.size, item.xs[i - 1]! + item.widths[i - 1]! - x, (item.font.ascent + item.font.descent) * item.size);
          runStart = -1;
        };
        for (let i = 0; i < item.chars.length; i++) {
          const c = { para: e.para, char: e.start + (isText ? item.source.offsets[i]! : 0) };
          const selected = compare(c, start) >= 0 && compare(c, end) < 0;
          if (selected && runStart < 0) runStart = i;
          if (!selected) flush(i);
        }
        flush(item.chars.length);
      }
      // Inline pictures inside the selection.
      for (const item of page.items) {
        if (item.type !== 'image') continue;
        const node = this.imageNode(item);
        const e = node && this.nodeIndex.get(node);
        if (!e) continue;
        const c = { para: e.para, char: e.start };
        if (compare(c, start) >= 0 && compare(c, end) < 0) box('docxkeun-selection', item.x, item.y, item.w, item.h);
      }
      // Empty paragraphs inside the selection.
      for (const line of this.bodyLines(page)) {
        const para = this.paraIndex.get(line.paragraph)!;
        if (para >= start.para && para < end.para && this.textOf(para) === '') box('docxkeun-selection', line.x, line.y, 4, line.h);
      }
    }
    const caret = this.screen(this.focusPos);
    if (caret && caret.page === index) box('docxkeun-caret', caret.x, caret.y, undefined, caret.h);
  }

  // ---- mouse ---------------------------------------------------------------------------------

  private pagePoint(e: MouseEvent): { page: number; x: number; y: number } | undefined {
    const target = document.elementFromPoint(e.clientX, e.clientY);
    const pageEl = (target?.closest('[data-page]') ?? (e.target as Element | null)?.closest?.('[data-page]')) as HTMLElement | null;
    if (!pageEl || !this.viewer.root.contains(pageEl)) return undefined;
    const rect = pageEl.getBoundingClientRect();
    const s = this.viewer.scale;
    return { page: Number(pageEl.dataset.page), x: (e.clientX - rect.left) / s, y: (e.clientY - rect.top) / s };
  }

  private pointToCaret(e: MouseEvent): { caret: Caret; page: number; x: number; y: number } | undefined {
    const at = this.pagePoint(e);
    if (!at) return undefined;
    const caret = this.hitTest(at.page, at.x, at.y);
    return caret ? { caret, ...at } : undefined;
  }

  /**
   * Switch stories for a click, as Word does: a click in a note or text box enters it, a
   * double-click enters a header or footer, and a click in the body leaves them.
   */
  private storyForClick(at: { page: number; x: number; y: number }, double: boolean): EditStory | 'ignore' | undefined {
    const page = this.pages()[at.page];
    if (!page) return undefined;
    const line = page.storyLines.find((l) => within(l, at.x, at.y)) ?? page.storyLines.find((l) => within(l, at.x, at.y, 8));
    const target = line && storyOfLine(line);
    const current = this.storyValue;
    if (target) {
      if (sameStory(target, current)) return undefined;
      // A single click on a header or footer waits for the second click (as in Word).
      if (target.kind === 'header' || target.kind === 'footer') return double ? target : 'ignore';
      return target;
    }
    if (current.kind === 'body') return undefined;
    // Leaving a header or footer takes a click in the body area; any click outside a note or text box leaves it.
    const c = page.content;
    const inBody = at.y >= c.y && at.y <= c.y + c.h;
    if (current.kind === 'header' || current.kind === 'footer') return inBody ? { kind: 'body' } : 'ignore';
    return { kind: 'body' };
  }

  private onMouseDown(e: MouseEvent): void {
    if (!this.enabled || e.button !== 0) return;
    if (!(e.target as Element).closest?.('.docxkeun-page')) return;
    const hit = this.pointToCaret(e);
    if (!hit) return;
    // Cmd/Ctrl-click follows links, as in Word.
    if (e.metaKey || e.ctrlKey) {
      const link = this.pages()[hit.page]?.links.find((l) => hit.x >= l.x && hit.x <= l.x + l.w && hit.y >= l.y && hit.y <= l.y + l.h);
      if (link) {
        e.preventDefault();
        this.viewer.followLink(link.target);
        return;
      }
    }
    e.preventDefault();
    const at = { page: hit.page, x: hit.x, y: hit.y };
    const found = this.storyForClick(at, e.detail >= 2);
    if (found === 'ignore') return;
    const next = found;
    if (next) {
      this.storyValue = next;
      this.refreshModel();
      this.viewer.root.dataset.editStory = next.kind;
      this.viewer.emit('storychange', { story: next });
      const caret = this.hitTest(at.page, at.x, at.y);
      if (!caret) return;
      hit.caret = caret;
    } else if (this.storyValue.kind !== 'body' && e.detail >= 2) {
      // Fall through to word selection in the story.
    }
    if (e.shiftKey && this.hasCaret && !next) this.focusPos = hit.caret;
    else this.anchor = this.focusPos = hit.caret;
    this.hasCaret = true;
    this.dragging = true;
    this.goalX = undefined;
    this.input.focus({ preventScroll: true });
    this.redraw();
  }

  private onMouseMove(e: MouseEvent): void {
    if (!this.dragging) return;
    const hit = this.pointToCaret(e);
    if (!hit) return;
    this.focusPos = hit.caret;
    this.redraw();
  }

  private onDoubleClick(e: MouseEvent): void {
    if (!this.enabled) return;
    const hit = this.pointToCaret(e);
    if (!hit) return;
    const text = this.textOf(hit.caret.para);
    const word = /[\p{L}\p{N}_]/u;
    let a = hit.caret.char;
    let b = hit.caret.char;
    while (a > 0 && word.test(text[a - 1]!)) a--;
    while (b < text.length && word.test(text[b]!)) b++;
    this.select({ para: hit.caret.para, char: a }, { para: hit.caret.para, char: b });
  }

  // ---- keyboard ------------------------------------------------------------------------------

  private onKeyDown(e: KeyboardEvent): void {
    if (!this.enabled || this.composing || e.isComposing) return;
    const mod = isMac ? e.metaKey : e.ctrlKey;
    const wordMod = isMac ? e.altKey : e.ctrlKey;
    let handled = true;
    switch (e.key) {
      case 'ArrowLeft':
      case 'ArrowRight': {
        const dir = e.key === 'ArrowLeft' ? -1 : 1;
        if (isMac && e.metaKey) this.moveToLineEdge(dir, e.shiftKey);
        else this.moveHorizontal(dir, e.shiftKey, wordMod);
        break;
      }
      case 'ArrowUp':
      case 'ArrowDown':
        this.moveVertical(e.key === 'ArrowUp' ? -1 : 1, e.shiftKey);
        break;
      case 'Home':
      case 'End':
        this.moveToLineEdge(e.key === 'Home' ? -1 : 1, e.shiftKey);
        break;
      case 'Backspace':
        this.deleteBackward(wordMod);
        break;
      case 'Delete':
        this.deleteForward();
        break;
      case 'Enter':
        if (e.shiftKey) this.lineBreak();
        else this.newParagraph();
        break;
      case 'Tab':
        this.typeText('\t');
        break;
      case 'Escape':
        this.anchor = this.focusPos;
        this.redraw();
        break;
      default:
        handled = false;
    }
    if (!handled && mod) {
      handled = true;
      switch (e.key.toLowerCase()) {
        case 'z':
          if (e.shiftKey) this.redo();
          else this.undo();
          break;
        case 'y':
          this.redo();
          break;
        case 'a':
          this.select({ para: 0, char: 0 }, { para: this.paras.length - 1, char: this.textOf(this.paras.length - 1).length });
          break;
        case 'b':
          this.toggle('bold');
          break;
        case 'e':
          this.formatParagraphs({ align: 'center' });
          break;
        case 'l':
          this.formatParagraphs({ align: 'left' });
          break;
        case 'r':
          this.formatParagraphs({ align: 'right' });
          break;
        case 'j':
          this.formatParagraphs({ align: 'justify' });
          break;
        case 'i':
          this.toggle('italic');
          break;
        case 'u':
          this.toggle('underline');
          break;
        default:
          handled = false; // copy / cut / paste go through their own events
      }
    }
    if (handled) e.preventDefault();
  }

  private onInput(): void {
    if (this.composing) return;
    const text = this.input.value;
    this.input.value = '';
    if (text) this.typeText(text);
  }

  private selectedText(): string {
    const [start, end] = this.ordered();
    let out = '';
    for (let p = start.para; p <= end.para; p++) {
      const text = this.textOf(p);
      out += text.slice(p === start.para ? start.char : 0, p === end.para ? end.char : text.length);
      if (p < end.para) out += '\n';
    }
    return out;
  }

  private onCopy(e: ClipboardEvent, cut: boolean): void {
    if (!this.hasCaret || compare(this.anchor, this.focusPos) === 0) return;
    e.preventDefault();
    e.clipboardData?.setData('text/plain', this.selectedText().replaceAll(OBJECT_CHAR, ''));
    e.clipboardData?.setData('text/html', this.selectedHtml());
    if (cut) this.apply('cut', () => this.deleteSelectionRaw());
  }

  /**
   * Copy (or cut) the selection with the async clipboard: HTML and plain text, or plain text only
   * where rich items are refused. For menus; the keyboard shortcuts use the native events above.
   */
  async copySelection(cut = false): Promise<void> {
    if (!this.enabled || !this.hasCaret || compare(this.anchor, this.focusPos) === 0) return;
    const text = this.selectedText().replaceAll(OBJECT_CHAR, '');
    const html = this.selectedHtml();
    const clipboard = typeof navigator === 'undefined' ? undefined : navigator.clipboard;
    if (!clipboard) return;
    let done = false;
    if (clipboard.write && typeof ClipboardItem !== 'undefined') {
      try {
        await clipboard.write([new ClipboardItem({ 'text/html': new Blob([html], { type: 'text/html' }), 'text/plain': new Blob([text], { type: 'text/plain' }) })]);
        done = true;
      } catch {
        // Fall back to plain text below.
      }
    }
    if (!done) {
      try {
        await clipboard.writeText(text);
      } catch {
        return;
      }
    }
    if (cut) this.apply('cut', () => this.deleteSelectionRaw());
  }

  private onPaste(e: ClipboardEvent): void {
    e.preventDefault();
    const html = e.clipboardData?.getData('text/html');
    if (html && this.pasteHtml(html)) return;
    const text = e.clipboardData?.getData('text/plain');
    if (text) this.pasteText(text);
  }

  /** Paste HTML (rich paste: formatting, headings, lists and links). False when it holds no paragraphs. */
  pasteHtml(html: string): boolean {
    const pasted = htmlToParagraphs(html);
    if (!pasted?.length) return false;
    this.pasteParagraphs(pasted);
    return true;
  }

  /** Paste plain text at the caret (replacing the selection). */
  pasteText(text: string): void {
    if (text) this.typeText(text.replace(/\r\n?/g, '\n'), 'paste');
  }

  /** The selection as HTML (formatting, headings and lists), for pasting into other programs. */
  private selectedHtml(): string {
    const [start, end] = this.ordered();
    const out: Array<{ para: DocParagraph; from: number; to: number }> = [];
    for (let p = start.para; p <= end.para; p++) {
      const para = this.paras[p]!;
      out.push({ para, from: p === start.para ? start.char : 0, to: p === end.para ? end.char : paragraphText(para).text.length });
    }
    return paragraphsToHtml(out, (para) => this.listKind(para));
  }

  private listKind(para: DocParagraph): 'bullet' | 'number' | undefined {
    if (!para.numbering) return undefined;
    return para.numbering.level.format === 'bullet' ? 'bullet' : 'number';
  }

  /** Insert formatted paragraphs (from pasted HTML) at the caret. */
  pasteParagraphs(paragraphs: PastedParagraph[]): void {
    const links: Array<{ para: number; from: number; to: number; url: string }> = [];
    const lists: Array<{ para: number; format: ListFormat }> = [];
    this.apply('paste', () => {
      let caret = this.deleteSelectionRaw();
      const editor = this.editor();
      paragraphs.forEach((para, i) => {
        if (i > 0) {
          const el = editor.splitParagraph(this.toDoc(caret, caret.char === 0 ? 'start' : 'end'));
          this.refreshModel();
          caret = { para: this.paraIndex.get(el) ?? caret.para + 1, char: 0 };
          if (para.style) editor.setParagraphStyle(el, para.style);
        } else if (para.style && caret.char === 0 && this.textOf(caret.para) === '') editor.setParagraphStyle(this.paras[caret.para]!.el, para.style);
        if (para.list) lists.push({ para: caret.para, format: para.list });
        for (const run of para.runs) {
          if (!run.text) continue;
          const from = caret.char;
          editor.insertText(this.toDoc(caret, caret.char === 0 ? 'start' : 'end'), run.text);
          this.refreshModel();
          caret = { para: caret.para, char: caret.char + run.text.length };
          // Pasted text keeps its own formatting: the on/off properties are set either way, since
          // text typed after a bold word would otherwise come out bold too.
          const format: TextFormat = { bold: false, italic: false, underline: false, strike: false, ...run.format };
          editor.formatRange(this.rangeOf({ para: caret.para, char: from }, caret), format);
          this.refreshModel();
          if (run.link) links.push({ para: caret.para, from, to: caret.char, url: run.link });
        }
      });
      return caret;
    });
    if (!links.length && !lists.length) return;
    // Links and lists need new relationships and numbering definitions (asynchronous).
    void (async () => {
      const editor = this.editor();
      try {
        for (const l of links.reverse()) {
          this.refreshModel();
          await editor.addHyperlink(this.rangeOf({ para: l.para, char: l.from }, { para: l.para, char: l.to }), { url: l.url });
        }
        for (const format of ['bullet', 'decimal'] as const) {
          const els = lists.filter((l) => l.format === format).map((l) => this.paras[l.para]?.el).filter((e): e is XmlElement => !!e);
          if (els.length) await editor.setList(els, { format });
        }
      } catch (error) {
        this.viewer.emit('editerror', { error: error instanceof Error ? error : new Error(String(error)) });
      }
    })();
  }

  // ---- movement ------------------------------------------------------------------------------

  private ordered(): [Caret, Caret] {
    return compare(this.anchor, this.focusPos) <= 0 ? [this.anchor, this.focusPos] : [this.focusPos, this.anchor];
  }

  private setCaret(c: Caret, extend: boolean): void {
    this.pending = undefined;
    this.focusPos = this.clamp(c);
    if (!extend) this.anchor = this.focusPos;
    this.hasCaret = true;
    this.redraw(true);
  }

  private moveHorizontal(dir: -1 | 1, extend: boolean, byWord: boolean): void {
    this.goalX = undefined;
    const [start, end] = this.ordered();
    if (!extend && compare(start, end) !== 0) return this.setCaret(dir < 0 ? start : end, false);
    let { para, char } = this.focusPos;
    const text = this.textOf(para);
    if (dir < 0) {
      if (char === 0) {
        if (para === 0) return;
        para--;
        char = this.textOf(para).length;
      } else if (byWord) {
        while (char > 0 && /\s/.test(text[char - 1]!)) char--;
        while (char > 0 && !/\s/.test(text[char - 1]!)) char--;
      } else char -= /[\udc00-\udfff]/.test(text[char - 1]!) && char > 1 ? 2 : 1;
    } else {
      if (char >= text.length) {
        if (para >= this.paras.length - 1) return;
        para++;
        char = 0;
      } else if (byWord) {
        while (char < text.length && !/\s/.test(text[char]!)) char++;
        while (char < text.length && /\s/.test(text[char]!)) char++;
      } else char += /[\ud800-\udbff]/.test(text[char]!) ? 2 : 1;
    }
    this.setCaret({ para, char }, extend);
  }

  private allLines(): Array<{ page: number; line: BodyLine }> {
    // Visual order: page, then top to bottom, then left to right (table cells, columns).
    return this.pages().flatMap((page) =>
      this.bodyLines(page)
        .sort((a, b) => a.y - b.y || a.x - b.x)
        .map((line) => ({ page: page.index, line })),
    );
  }

  private currentLine(): { index: number; lines: Array<{ page: number; line: BodyLine }>; pos: ScreenPos } | undefined {
    const pos = this.screen(this.focusPos);
    if (!pos) return undefined;
    const lines = this.allLines();
    const mid = pos.y + pos.h / 2;
    let index = lines.findIndex(({ page, line }) => page === pos.page && mid >= line.y && mid <= line.y + line.h && pos.x >= line.x - 6 && pos.x <= line.x + line.w + 6);
    if (index < 0) index = lines.findIndex(({ page, line }) => page === pos.page && mid >= line.y && mid <= line.y + line.h);
    return index < 0 ? undefined : { index, lines, pos };
  }

  private moveVertical(dir: -1 | 1, extend: boolean): void {
    const cur = this.currentLine();
    if (!cur) return;
    this.goalX ??= cur.pos.x;
    const target = cur.lines[cur.index + dir];
    if (!target) return this.setCaret(dir < 0 ? { para: 0, char: 0 } : { para: this.paras.length - 1, char: this.textOf(this.paras.length - 1).length }, extend);
    const caret = this.hitTest(target.page, this.goalX, target.line.y + target.line.h / 2);
    if (!caret) return;
    const goal = this.goalX;
    this.setCaret(caret, extend);
    this.goalX = goal;
  }

  private moveToLineEdge(dir: -1 | 1, extend: boolean): void {
    this.goalX = undefined;
    const cur = this.currentLine();
    if (!cur) return;
    const { page, line } = cur.lines[cur.index]!;
    const caret = this.hitTest(page, dir < 0 ? line.x - 1 : line.x + line.w + 1000, line.y + line.h / 2);
    if (caret) this.setCaret(caret, extend);
  }

  // ---- edits ---------------------------------------------------------------------------------

  /** Run an edit with undo support; report refused edits instead of throwing. */
  private apply(kind: string, fn: () => Caret | undefined): void {
    const doc = this.doc;
    if (!doc || !this.enabled || !this.hasCaret || this.history.busy) return;
    const now = Date.now();
    // Consecutive typing within a second is one undo step.
    const coalesce = kind === 'type' && this.lastEdit.kind === 'type' && now - this.lastEdit.at < 1000;
    const snapshot = this.snapshot();
    try {
      const next = fn();
      if (this.storyValue.kind === 'textbox') dropStaleFallback(doc.documentXml.root, this.storyValue.container);
      if (!coalesce) this.history.push(snapshot);
      this.lastEdit = { kind, at: now };
      this.refreshModel();
      if (next) {
        this.anchor = this.focusPos = next;
        this.goalX = undefined;
      }
      this.redraw(true);
    } catch (error) {
      // Edits that failed half-way are rolled back.
      this.putRoot(snapshot);
      this.refreshModel();
      this.viewer.emit('editerror', { error: error instanceof Error ? error : new Error(String(error)) });
    }
  }

  /** Delete the selection (no undo bookkeeping); returns the caret where it was. */
  private deleteSelectionRaw(): Caret {
    const [start, end] = this.ordered();
    if (compare(start, end) === 0) return start;
    const editor = this.editor();
    if (start.para === end.para) {
      editor.deleteRange({ start: this.toDoc(start, 'start'), end: this.toDoc(end, 'end') });
      this.refreshModel();
      return start;
    }
    // Trim the end paragraph's head, drop whole paragraphs between, trim the start's tail, join.
    if (end.char > 0) {
      editor.deleteRange({ start: this.toDoc({ para: end.para, char: 0 }, 'start'), end: this.toDoc(end, 'end') });
      this.refreshModel();
    }
    for (let p = end.para - 1; p > start.para; p--) {
      editor.deleteParagraph(this.paras[p]!.el);
      this.refreshModel();
    }
    const startLen = this.textOf(start.para).length;
    if (start.char < startLen) {
      editor.deleteRange({ start: this.toDoc(start, 'start'), end: this.toDoc({ para: start.para, char: startLen }, 'end') });
      this.refreshModel();
    }
    editor.mergeParagraphs(this.paras[start.para]!.el, this.paras[start.para + 1]!.el);
    this.refreshModel();
    return start;
  }

  /** Type text at the caret, replacing the selection. Newlines start new paragraphs. */
  typeText(text: string, kind = 'type'): void {
    this.apply(kind, () => {
      let caret = this.deleteSelectionRaw();
      const editor = this.editor();
      const lines = text.split('\n');
      lines.forEach((line, i) => {
        if (i > 0) {
          const para = editor.splitParagraph(this.toDoc(caret, caret.char === 0 ? 'start' : 'end'));
          this.refreshModel();
          caret = { para: this.paraIndex.get(para) ?? caret.para + 1, char: 0 };
        }
        if (!line) return;
        const from = caret;
        editor.insertText(this.toDoc(caret, caret.char === 0 ? 'start' : 'end'), line);
        this.refreshModel();
        caret = { para: caret.para, char: caret.char + line.length };
        if (this.pending && kind === 'type') {
          editor.formatRange(this.rangeOf(from, caret), this.pending);
          this.refreshModel();
        }
      });
      return caret;
    });
  }

  /** Shift+Enter inserts a line break inside the paragraph. */
  private lineBreak(): void {
    this.apply('type', () => {
      const caret = this.deleteSelectionRaw();
      this.editor().insertText(this.toDoc(caret, caret.char === 0 ? 'start' : 'end'), '\n');
      return { para: caret.para, char: caret.char + 1 };
    });
  }

  private newParagraph(): void {
    this.apply('enter', () => {
      const caret = this.deleteSelectionRaw();
      const len = this.textOf(caret.para).length;
      const para = this.editor().splitParagraph(this.toDoc(caret, caret.char === 0 ? 'start' : caret.char >= len ? 'end' : 'start'));
      this.refreshModel();
      // Usually the paragraph after the caret; before it when Enter was pressed in front of a locked block.
      return { para: this.paraIndex.get(para) ?? caret.para + 1, char: 0 };
    });
  }

  private deleteBackward(byWord: boolean): void {
    this.apply('delete', () => {
      const [start, end] = this.ordered();
      if (compare(start, end) !== 0) return this.deleteSelectionRaw();
      const { para, char } = this.focusPos;
      if (char === 0) {
        if (para === 0) return undefined;
        const prevLen = this.textOf(para - 1).length;
        this.editor().mergeParagraphs(this.paras[para - 1]!.el, this.paras[para]!.el);
        return { para: para - 1, char: prevLen };
      }
      const text = this.textOf(para);
      let from = char - (/[\udc00-\udfff]/.test(text[char - 1]!) && char > 1 ? 2 : 1);
      if (byWord) {
        from = char;
        while (from > 0 && /\s/.test(text[from - 1]!)) from--;
        while (from > 0 && !/\s/.test(text[from - 1]!)) from--;
      }
      this.anchor = { para, char: from };
      return this.deleteSelectionRaw();
    });
  }

  private deleteForward(): void {
    this.apply('delete', () => {
      const [start, end] = this.ordered();
      if (compare(start, end) !== 0) return this.deleteSelectionRaw();
      const { para, char } = this.focusPos;
      const text = this.textOf(para);
      if (char >= text.length) {
        if (para >= this.paras.length - 1) return undefined;
        this.editor().mergeParagraphs(this.paras[para]!.el, this.paras[para + 1]!.el);
        return { para, char };
      }
      this.anchor = { para, char };
      this.focusPos = { para, char: char + (/[\ud800-\udbff]/.test(text[char]!) ? 2 : 1) };
      return this.deleteSelectionRaw();
    });
  }

  /** Turn bold, italic, underline or strikethrough on or off for the selection (or the next text typed). */
  toggle(prop: 'bold' | 'italic' | 'underline' | 'strike'): void {
    const now = this.caretFormat();
    if (!now) return;
    this.format({ [prop]: !now[prop] });
  }

  /** Format the selected text; with only a caret, the format applies to the next text typed. */
  format(format: TextFormat): void {
    if (!this.hasCaret) return;
    const [start, end] = this.ordered();
    if (compare(start, end) === 0) {
      this.pending = { ...this.pending, ...format };
      this.reportFormat();
      return;
    }
    this.apply('format', () => {
      this.editor().formatRange(this.rangeOf(start, end), format);
      return undefined;
    });
    // Formatting keeps the selection.
    this.anchor = start;
    this.focusPos = end;
    this.redraw();
  }

  /** Paragraphs the selection touches. */
  private selectedParagraphs(): XmlElement[] {
    const [start, end] = this.ordered();
    const out: XmlElement[] = [];
    for (let p = start.para; p <= end.para; p++) if (this.paras[p]) out.push(this.paras[p]!.el);
    return out;
  }

  /** Alignment, direction, indents or spacing of the selected paragraphs. */
  formatParagraphs(format: ParagraphFormat): Promise<void> {
    return this.act('format', (editor) => {
      for (const p of this.selectedParagraphs()) editor.formatParagraph(p, format);
    });
  }

  /** Apply a paragraph style (e.g. "Heading1") to the selected paragraphs. */
  setStyle(styleId: string): Promise<void> {
    return this.act('style', (editor) => {
      for (const p of this.selectedParagraphs()) editor.setParagraphStyle(p, styleId);
    });
  }

  /** Make the selected paragraphs a bulleted or numbered list, or take them out of one. */
  setList(format: ListFormat | 'none'): Promise<void> {
    return this.act('list', (editor) => editor.setList(this.selectedParagraphs(), { format }));
  }

  /** Increase (1) or decrease (-1) the indent: the list level in a list, else the left indent by half an inch. */
  async indent(dir: 1 | -1): Promise<void> {
    if (!this.hasCaret) return;
    const first = this.paras[this.ordered()[0].para];
    if (first?.numbering) {
      const level = Math.max(0, Math.min(8, first.numbering.ilvl + dir));
      return this.act('indent', (editor) => editor.setList(this.selectedParagraphs(), { level }));
    }
    const left = first?.props.indent?.left ?? 0;
    return this.formatParagraphs({ indentLeft: Math.max(0, left + dir * 36) });
  }

  /** Whether there is something to undo or redo (for toolbars). */
  get canUndo(): boolean {
    return this.history.canUndo;
  }

  get canRedo(): boolean {
    return this.history.canRedo;
  }

  /**
   * Run an edit that may touch any part of the document as one undo step. A failure puts the
   * document back, reports `editerror` and rethrows. Inside `fn`, change the document through the
   * `editor` it receives. A nested `transaction()`, `setList()` and the like joins this step only
   * before `fn`'s first `await`; after it, it waits for this step and gives up after
   * `UndoHistory.queueTimeout` (10 s by default): a nested `transaction()` rejects with an
   * `EditError`, while `setList()` and the like report it through `editerror`. Edits requested
   * elsewhere meanwhile wait their turn, act on the selection they were requested with, and are
   * undo steps of their own.
   */
  async transaction<T>(label: string, fn: (editor: DocxEditor) => T | Promise<T>): Promise<T> {
    const doc = this.doc;
    if (!doc) throw new Error('No document is open');
    if (this.history.joinable) {
      // Nested: the outer transaction checkpoints, records, restores and redraws.
      try {
        return await fn(this.editor());
      } catch (error) {
        this.reportError(error);
        throw error;
      }
    }
    const story = this.storyValue;
    const where = { story, anchor: { ...this.anchor }, focus: { ...this.focusPos } };
    const container = textboxIndex(doc.documentXml.root, story);
    let started = false;
    try {
      const result = await this.history.transaction(
        doc,
        label,
        () => {
          started = true;
          // A queued edit acts on the story and selection it was requested with, not on the ones
          // the transaction ahead of it put back (act() and the menus read them as they run).
          this.putSelection(doc.documentXml.root, story, container, where);
          return fn(this.editor());
        },
        { ...where, ...(container !== undefined ? { container } : {}) },
      );
      this.lastEdit = { kind: '', at: 0 };
      this.refreshModel();
      this.anchor = this.clamp(where.anchor);
      this.focusPos = this.clamp(where.focus);
      // The selection is put back unchanged: no scrolling to it (a comment card's Resolve on
      // page 8 must not jump to a caret left on page 1).
      this.redraw();
      return result;
    } catch (error) {
      // Never started (it gave up waiting for another transaction, or no checkpoint could be
      // taken): the document is untouched, and the one still running redraws when it ends.
      if (started) {
        // The failure restore replaced the XML trees: find the text box again.
        this.putSelection(doc.documentXml.root, story, container, where);
        this.redraw();
      }
      this.reportError(error);
      throw error;
    }
  }

  /** Make `story` (a text box found again by its index) and the selection current, with a fresh model. */
  private putSelection(root: XmlElement, story: EditStory, container: number | undefined, at: { anchor: Caret; focus: Caret }): void {
    this.storyValue = resolveStory(root, story, container);
    this.viewer.root.dataset.editStory = this.storyValue.kind;
    this.refreshModel();
    this.anchor = this.clamp(at.anchor);
    this.focusPos = this.clamp(at.focus);
  }

  /** Errors already reported through `editerror` (a nested transaction's failure reaching the outer one). */
  private readonly reported = new WeakSet<object>();

  private reportError(error: unknown): void {
    if (typeof error === 'object' && error !== null) {
      if (this.reported.has(error)) return;
      this.reported.add(error);
    }
    this.viewer.emit('editerror', { error: error instanceof Error ? error : new Error(String(error)) });
  }

  /** A transaction on the selection whose failure has already been reported through `editerror`. */
  private async act(label: string, fn: (editor: DocxEditor) => void | Promise<unknown>): Promise<void> {
    if (!this.doc || !this.hasCaret) return;
    try {
      await this.transaction(label, async (editor) => {
        await fn(editor);
        if (this.storyValue.kind === 'textbox') dropStaleFallback(this.doc!.documentXml.root, this.storyValue.container);
      });
    } catch {
      // editerror was emitted by transaction().
    }
  }

  /** Formatting at the caret (start of the selection), with any format waiting for the next text typed. */
  caretFormat(): CaretFormat | undefined {
    if (!this.hasCaret || !this.paras.length) return undefined;
    const [start, end] = this.ordered();
    const para = this.paras[start.para]!;
    // The text before the caret decides, as in Word (the first character of a selection).
    const k = compare(start, end) === 0 ? Math.max(0, start.char - 1) : start.char;
    const pt = paragraphText(para);
    const seg = pt.segments.find((sg) => k >= sg.start && k < sg.end) ?? pt.segments[pt.segments.length - 1];
    const inline = seg && para.inlines.find((i) => (i.kind === 'text' || i.kind === 'tab') && i.node === seg.node);
    const props = inline && 'props' in inline ? inline.props : para.markProps;
    const pending = this.pending ?? {};
    const underline = pending.underline !== undefined ? !!pending.underline : !!props.underline && props.underline !== 'none';
    const jc = para.props.jc ?? 'left';
    const rtl = !!para.props.bidi;
    const align = jc === 'both' || jc === 'distribute' ? 'justify' : jc === 'center' ? 'center' : (jc === 'right') !== rtl ? 'right' : 'left';
    return {
      bold: pending.bold ?? !!props.bold,
      italic: pending.italic ?? !!props.italic,
      underline,
      strike: pending.strike ?? !!props.strike,
      size: pending.size ?? props.size,
      font: pending.font ?? props.fonts?.ascii,
      color: pending.color ?? (props.color && props.color !== 'auto' ? props.color : undefined),
      highlight: pending.highlight ?? props.highlight,
      align,
      direction: rtl ? 'rtl' : 'ltr',
      styleId: para.styleId,
      list: this.listKind(para),
      story: this.storyValue.kind,
      hasSelection: compare(start, end) !== 0,
      trackChanges: this.trackChanges,
      canUndo: this.history.canUndo,
      canRedo: this.history.canRedo,
    };
  }

  private reportFormat(): void {
    this.viewer.selectionMoved();
    const format = this.enabled ? this.caretFormat() : undefined;
    const key = JSON.stringify(format ?? null);
    if (key === this.lastFormat) return;
    this.lastFormat = key;
    this.viewer.emit('caretchange', { format });
  }

  /** The current story's part as it is now, with the selection. */
  private snapshot(story: EditStory = this.storyValue): StoryEntry {
    const doc = this.doc!;
    const part = story.kind === 'body' || story.kind === 'textbox' ? doc.mainPart : story.part;
    const root = part === doc.mainPart ? doc.documentXml.root : doc.xmlSync(part)!.root;
    const container = story.kind === 'textbox' ? findAll(root, 'w:txbxContent').indexOf(story.container) : undefined;
    return { kind: 'story', part, story, container, root: cloneElement(root), anchor: { ...this.anchor }, focus: { ...this.focusPos } };
  }

  /** Put a snapshot's XML back into its part. */
  private putRoot(snap: StoryEntry): void {
    const doc = this.doc!;
    if (snap.part === doc.mainPart) doc.documentXml.root = snap.root;
    else doc.xmlSync(snap.part)!.root = snap.root;
    doc.touch(snap.part);
  }

  /** The entry for the state being left, of the same kind as the one being restored. */
  private currentFor(restoring: UndoEntry): UndoEntry {
    const liveRoot = this.doc!.documentXml.root;
    if (restoring.kind === 'story') return this.snapshot(resolveStory(liveRoot, restoring.story, restoring.container));
    const container = textboxIndex(liveRoot, this.storyValue);
    return {
      kind: 'package',
      checkpoint: this.doc!.checkpoint(),
      story: this.storyValue,
      ...(container !== undefined ? { container } : {}),
      anchor: { ...this.anchor },
      focus: { ...this.focusPos },
      label: restoring.label,
    };
  }

  undo(): void {
    if (!this.doc || this.history.busy) return;
    const entry = this.history.undo((e) => this.currentFor(e));
    if (entry) this.restore(entry);
  }

  redo(): void {
    if (!this.doc || this.history.busy) return;
    const entry = this.history.redo((e) => this.currentFor(e));
    if (entry) this.restore(entry);
  }

  private restore(entry: UndoEntry): void {
    const doc = this.doc!;
    let root: XmlElement;
    if (entry.kind === 'story') {
      this.putRoot(entry);
      root = entry.part === doc.mainPart ? doc.documentXml.root : doc.xmlSync(entry.part)!.root;
    } else {
      doc.restore(entry.checkpoint);
      root = doc.documentXml.root;
    }
    // A text box container is replaced along with the document's root: find its copy.
    const story = resolveStory(root, entry.story, entry.container);
    this.storyValue = story;
    this.viewer.root.dataset.editStory = story.kind;
    this.refreshModel();
    this.lastEdit = { kind: '', at: 0 };
    this.anchor = this.clamp(entry.anchor);
    this.focusPos = this.clamp(entry.focus);
    this.hasCaret = true;
    this.redraw(true);
  }
}

/**
 * Word stores a text box twice: the DrawingML version in mc:Choice and a VML copy in
 * mc:Fallback. After editing the first, the copy is out of date, so it is removed (Word and
 * other current readers use the Choice; the Fallback is only for very old ones).
 */
function dropStaleFallback(root: XmlElement, container: XmlElement): void {
  for (const ac of findAll(root, 'mc:AlternateContent')) {
    const choice = elements(ac).find((n) => n.name === 'mc:Choice');
    if (!choice || !findAll(choice, 'w:txbxContent').includes(container)) continue;
    ac.children = ac.children.filter((n) => !(n.type === 'element' && n.name === 'mc:Fallback'));
  }
}
