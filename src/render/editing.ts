/**
 * Typing directly on the rendered pages: a caret and selection drawn over the SVG, keyboard and
 * IME input through a hidden textarea, and every change made through DocxEditor (so access
 * guards apply and unknown markup survives). The page is laid out again after each change.
 *
 * Positions are kept as (paragraph index, character offset) so they survive re-layout and
 * the run splitting that edits cause.
 */
import { DocxEditor, paragraphsOf, paragraphText, positionAt, type EditGuard, type TextFormat } from '../api/editor.js';
import type { DocPosition } from '../api/positions.js';
import type { BodyLine, Page, TextItem } from '../layout/types.js';
import type { DocParagraph, TextInline } from '../model/content.js';
import { cloneElement, type XmlElement } from '../xml/index.js';
import type { DocxViewer } from './viewer.js';

export interface EditingOptions {
  /** Checks every change (e.g. an AccessPolicy from docxkit/security). */
  guard?: EditGuard;
}

/** A place in the text: paragraph index (document order, tables included) and UTF-16 offset. */
export interface Caret {
  para: number;
  char: number;
}

interface Snapshot {
  root: XmlElement;
  anchor: Caret;
  focus: Caret;
}

interface ScreenPos {
  page: number;
  x: number;
  y: number;
  h: number;
}

const isMac = typeof navigator !== 'undefined' && /Mac|iPhone|iPad/.test(navigator.platform);

function compare(a: Caret, b: Caret): number {
  return a.para - b.para || a.char - b.char;
}

export class InlineEditing {
  private enabled = true;
  private anchor: Caret = { para: 0, char: 0 };
  private focusPos: Caret = { para: 0, char: 0 };
  private hasCaret = false;
  private readonly input: HTMLTextAreaElement;
  private composing = false;
  private dragging = false;
  private goalX: number | undefined;
  private readonly undoStack: Snapshot[] = [];
  private readonly redoStack: Snapshot[] = [];
  private lastEdit = { kind: '', at: 0 };
  private paras: DocParagraph[] = [];
  private readonly nodeIndex = new Map<XmlElement, { para: number; start: number }>();
  private readonly paraIndex = new Map<XmlElement, number>();
  private readonly layers = new Map<number, HTMLDivElement>();
  private readonly cleanup: Array<() => void> = [];

  constructor(private readonly viewer: DocxViewer, private readonly options: EditingOptions = {}) {
    const root = viewer.root;
    this.input = document.createElement('textarea');
    this.input.className = 'docxkit-input';
    this.input.setAttribute('aria-label', 'Document text');
    this.input.setAttribute('autocapitalize', 'off');
    this.input.setAttribute('autocomplete', 'off');
    this.input.spellcheck = false;
    root.appendChild(this.input);
    root.classList.add('docxkit-editing');

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
    listen(this.input, 'focus', () => root.classList.add('docxkit-focused'));
    listen(this.input, 'blur', () => root.classList.remove('docxkit-focused'));
    this.cleanup.push(viewer.on('layout', () => this.onLayout()));
    this.cleanup.push(viewer.on('pagemount', ({ index, element }) => this.onPageMount(index, element)));
    if (viewer.layout) this.onLayout();
  }

  setEnabled(on: boolean): void {
    this.enabled = on;
    this.viewer.root.classList.toggle('docxkit-editing', on);
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

  /** An editor range (for DocxEditor calls such as addComment) between two carets. */
  rangeOf(a: Caret, b: Caret): { start: DocPosition; end: DocPosition } {
    const [start, end] = compare(a, b) <= 0 ? [a, b] : [b, a];
    return { start: this.toDoc(start, 'start'), end: this.toDoc(end, 'end') };
  }

  destroy(): void {
    for (const fn of this.cleanup) fn();
    this.input.remove();
    for (const layer of this.layers.values()) layer.remove();
    this.viewer.root.classList.remove('docxkit-editing', 'docxkit-focused');
  }

  // ---- model ---------------------------------------------------------------------------------

  private get doc() {
    return this.viewer.document;
  }

  private editor(): DocxEditor {
    return new DocxEditor(this.doc!, this.options.guard);
  }

  private refreshModel(): void {
    const doc = this.doc;
    this.nodeIndex.clear();
    this.paraIndex.clear();
    this.paras = doc ? paragraphsOf(doc.body().blocks) : [];
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

  private bodyLines(page: Page): BodyLine[] {
    return page.lines.filter((l) => this.paraIndex.has(l.paragraph));
  }

  private lineItems(page: Page, line: BodyLine): TextItem[] {
    return page.items
      .filter((i): i is TextItem => i.type === 'text' && !!i.source && this.nodeIndex.has(i.source.inline.node))
      .filter((i) => i.y > line.y - 0.5 && i.y < line.y + line.h + 0.5 && i.xs[0]! >= line.x - 1 && i.xs[0]! <= line.x + line.w + 1)
      .sort((a, b) => a.xs[0]! - b.xs[0]!);
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
      for (let i = 0; i < item.chars.length; i++) {
        if (x < item.xs[i]! + item.widths[i]! / 2) return this.itemCaret(item, i, true);
      }
    }
    const last = items[items.length - 1]!;
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
          const line = page.lines.find((l) => l.paragraph === pos.node);
          if (line) return { page: page.index, x: line.x, y: line.y, h: line.h };
          continue;
        }
        for (const item of page.items) {
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
      const line = page.lines.find((l) => l.paragraph === el);
      if (line) return { page: page.index, x: line.x, y: line.y, h: line.h };
    }
    return undefined;
  }

  // ---- drawing -------------------------------------------------------------------------------

  private onLayout(): void {
    this.refreshModel();
    if (this.hasCaret) {
      this.anchor = this.clamp(this.anchor);
      this.focusPos = this.clamp(this.focusPos);
    }
    this.redraw();
  }

  private onPageMount(index: number, element: HTMLDivElement): void {
    const layer = document.createElement('div');
    layer.className = 'docxkit-edit-layer';
    element.appendChild(layer);
    this.layers.set(index, layer);
    this.drawPage(index);
  }

  private redraw(scrollIntoView = false): void {
    for (const index of this.layers.keys()) this.drawPage(index);
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
    const box = (cls: string, x: number, y: number, w: number, h: number) => {
      const el = document.createElement('div');
      el.className = cls;
      el.style.left = `${x * s}px`;
      el.style.top = `${y * s}px`;
      el.style.width = `${w * s}px`;
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
          box('docxkit-selection', x, item.y - item.font.ascent * item.size, item.xs[i - 1]! + item.widths[i - 1]! - x, (item.font.ascent + item.font.descent) * item.size);
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
      // Empty paragraphs inside the selection.
      for (const line of this.bodyLines(page)) {
        const para = this.paraIndex.get(line.paragraph)!;
        if (para >= start.para && para < end.para && this.textOf(para) === '') box('docxkit-selection', line.x, line.y, 4, line.h);
      }
    }
    const caret = this.screen(this.focusPos);
    if (caret && caret.page === index) box('docxkit-caret', caret.x, caret.y, 0, caret.h);
  }

  // ---- mouse ---------------------------------------------------------------------------------

  private pointToCaret(e: MouseEvent): { caret: Caret; page: number; x: number; y: number } | undefined {
    const target = document.elementFromPoint(e.clientX, e.clientY);
    const pageEl = (target?.closest('[data-page]') ?? (e.target as Element | null)?.closest?.('[data-page]')) as HTMLElement | null;
    if (!pageEl || !this.viewer.root.contains(pageEl)) return undefined;
    const page = Number(pageEl.dataset.page);
    const rect = pageEl.getBoundingClientRect();
    const s = this.viewer.scale;
    const x = (e.clientX - rect.left) / s;
    const y = (e.clientY - rect.top) / s;
    const caret = this.hitTest(page, x, y);
    return caret ? { caret, page, x, y } : undefined;
  }

  private onMouseDown(e: MouseEvent): void {
    if (!this.enabled || e.button !== 0) return;
    if (!(e.target as Element).closest?.('.docxkit-page')) return;
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
    if (e.shiftKey && this.hasCaret) this.focusPos = hit.caret;
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
    e.clipboardData?.setData('text/plain', this.selectedText());
    if (cut) this.apply('cut', () => this.deleteSelectionRaw());
  }

  private onPaste(e: ClipboardEvent): void {
    e.preventDefault();
    const text = e.clipboardData?.getData('text/plain');
    if (text) this.typeText(text.replace(/\r\n?/g, '\n'), 'paste');
  }

  // ---- movement ------------------------------------------------------------------------------

  private ordered(): [Caret, Caret] {
    return compare(this.anchor, this.focusPos) <= 0 ? [this.anchor, this.focusPos] : [this.focusPos, this.anchor];
  }

  private setCaret(c: Caret, extend: boolean): void {
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
    if (!doc || !this.enabled || !this.hasCaret) return;
    const now = Date.now();
    // Consecutive typing within a second is one undo step.
    const coalesce = kind === 'type' && this.lastEdit.kind === 'type' && now - this.lastEdit.at < 1000;
    const snapshot: Snapshot = { root: cloneElement(doc.documentXml.root), anchor: { ...this.anchor }, focus: { ...this.focusPos } };
    try {
      const next = fn();
      if (!coalesce) this.undoStack.push(snapshot);
      if (this.undoStack.length > 200) this.undoStack.shift();
      this.redoStack.length = 0;
      this.lastEdit = { kind, at: now };
      this.refreshModel();
      if (next) {
        this.anchor = this.focusPos = next;
        this.goalX = undefined;
      }
      this.redraw(true);
    } catch (error) {
      // Edits that failed half-way are rolled back.
      doc.documentXml.root = snapshot.root;
      doc.touch(doc.mainPart);
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
          editor.splitParagraph(this.toDoc(caret, caret.char === 0 ? 'start' : 'end'));
          this.refreshModel();
          caret = { para: caret.para + 1, char: 0 };
        }
        if (!line) return;
        editor.insertText(this.toDoc(caret, caret.char === 0 ? 'start' : 'end'), line);
        this.refreshModel();
        caret = { para: caret.para, char: caret.char + line.length };
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
      this.editor().splitParagraph(this.toDoc(caret, caret.char === 0 ? 'start' : caret.char >= len ? 'end' : 'start'));
      return { para: caret.para + 1, char: 0 };
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

  private toggle(prop: 'bold' | 'italic' | 'underline'): void {
    const [start, end] = this.ordered();
    if (compare(start, end) === 0) return;
    const para = this.paras[start.para];
    const at = this.toDoc(start, 'start');
    const inline = para?.inlines.find((i) => i.kind === 'text' && i.node === at.node);
    const current = inline && inline.kind === 'text' ? (prop === 'underline' ? !!inline.props.underline && inline.props.underline !== 'none' : !!inline.props[prop]) : false;
    const format: TextFormat = { [prop]: !current };
    this.apply('format', () => {
      this.editor().formatRange({ start: at, end: this.toDoc(end, 'end') }, format);
      return undefined;
    });
    // Formatting keeps the selection.
    this.anchor = start;
    this.focusPos = end;
    this.redraw();
  }

  undo(): void {
    const doc = this.doc;
    const snap = this.undoStack.pop();
    if (!doc || !snap) return;
    this.redoStack.push({ root: cloneElement(doc.documentXml.root), anchor: { ...this.anchor }, focus: { ...this.focusPos } });
    this.restore(snap);
  }

  redo(): void {
    const doc = this.doc;
    const snap = this.redoStack.pop();
    if (!doc || !snap) return;
    this.undoStack.push({ root: cloneElement(doc.documentXml.root), anchor: { ...this.anchor }, focus: { ...this.focusPos } });
    this.restore(snap);
  }

  private restore(snap: Snapshot): void {
    const doc = this.doc!;
    doc.documentXml.root = snap.root;
    doc.touch(doc.mainPart);
    this.refreshModel();
    this.lastEdit = { kind: '', at: 0 };
    this.anchor = this.clamp(snap.anchor);
    this.focusPos = this.clamp(snap.focus);
    this.hasCaret = true;
    this.redraw(true);
  }
}
