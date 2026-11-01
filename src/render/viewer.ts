/**
 * DocxViewer: renders a document into a scrollable element, one SVG per page, mounting only the
 * pages near the viewport. Handles links, bookmarks, comment cards and selection mapping.
 */
import { readComments, type CommentInfo } from '../comments/read.js';
import { FontRegistry } from '../fonts/registry.js';
import { standaloneFont } from '../fonts/subset.js';
import { layoutDocument } from '../layout/engine.js';
import type { FontProvider, LayoutFont, LayoutResult, LinkArea, Page, PermArea } from '../layout/types.js';
import { DocxDocument, type OpenOptions } from '../model/document.js';
import type { LinkTarget } from '../model/fields.js';
import { extension } from '../opc/paths.js';
import { CanvasFontProvider } from './canvas-fonts.js';
import { InlineEditing, type EditingOptions } from './editing.js';
import { selectionPositions, type InlinePosition } from './selection.js';
import { pageToSvg } from './svg.js';

export interface ViewerOptions {
  /** Font files for exact layout. Without them, system fonts are measured with the canvas API. */
  fonts?: FontRegistry;
  zoom?: number;
  /** Show the comment sidebar (default true). */
  showComments?: boolean;
  /** URL schemes a link may open (default http, https, mailto, tel). */
  allowedProtocols?: string[];
  /** Load images referenced by URL (r:link). Off by default: they can track readers. */
  allowExternalImages?: boolean;
  /** Which permission ranges to shade (e.g. the ones the current user may edit). */
  highlightRange?(area: PermArea): boolean;
  /**
   * Let users type directly on the pages. Pass `{ guard }` (e.g. an AccessPolicy) to restrict
   * what they may change; refused edits are reported through the `editerror` event.
   */
  editable?: boolean | EditingOptions;
}

export interface ViewerEvents {
  /** A link was clicked. Call `preventDefault()` to stop the viewer from following it. */
  linkclick: { target: LinkTarget; url?: string; preventDefault(): void };
  commentclick: { comment: CommentInfo };
  selectionchange: { start: InlinePosition; end: InlinePosition } | undefined;
  layout: { pages: Page[] };
  /** A page's SVG was (re)drawn into its element. */
  pagemount: { index: number; element: HTMLDivElement; page: Page };
  /** An edit typed on the page was refused (e.g. a PermissionError) or failed. */
  editerror: { error: Error };
}

type Listener<K extends keyof ViewerEvents> = (event: ViewerEvents[K]) => void;

const STYLE_ID = 'docxkit-viewer-style';
const CSS = `
.docxkit-viewer{position:relative;overflow:auto;background:#e8eaed;height:100%;box-sizing:border-box;padding:16px 0}
.docxkit-row{display:flex;justify-content:center;align-items:flex-start;gap:12px;margin:0 auto 16px}
.docxkit-page{position:relative;background:#fff;box-shadow:0 1px 3px rgba(60,64,67,.3),0 1px 4px rgba(60,64,67,.15);flex:none}
.docxkit-page svg{display:block;user-select:text}
.docxkit-overlay{position:absolute;inset:0;pointer-events:none}
.docxkit-link{position:absolute;pointer-events:auto;cursor:pointer;display:block}
.docxkit-link:focus-visible{outline:2px solid #1a73e8;outline-offset:1px}
.docxkit-comment-hl{position:absolute;background:rgba(251,188,4,.28);border-bottom:2px solid rgba(251,188,4,.9);mix-blend-mode:multiply}
.docxkit-comment-hl.active{background:rgba(251,188,4,.5)}
.docxkit-perm{position:absolute;background:rgba(26,115,232,.12);border-bottom:1px dashed rgba(26,115,232,.7)}
.docxkit-gutter{position:relative;width:260px;flex:none}
.docxkit-card{position:absolute;left:0;right:0;background:#fff;border-radius:8px;padding:8px 10px;font:13px/1.4 system-ui,sans-serif;color:#202124;box-shadow:0 1px 2px rgba(60,64,67,.3);cursor:pointer;box-sizing:border-box}
.docxkit-card.resolved{opacity:.6}
.docxkit-card .who{font-weight:600}.docxkit-card .when{color:#5f6368;font-size:12px;margin-left:6px}
.docxkit-card .body{white-space:pre-wrap;margin-top:2px;overflow-wrap:anywhere}
.docxkit-card .reply{border-top:1px solid #e8eaed;margin-top:6px;padding-top:6px}
.docxkit-editing .docxkit-page svg{user-select:none;cursor:text}
.docxkit-editing .docxkit-link{pointer-events:none}
.docxkit-edit-layer{position:absolute;inset:0;pointer-events:none}
.docxkit-selection{position:absolute;background:rgba(26,115,232,.28)}
.docxkit-caret{position:absolute;width:1.5px;background:#202124;animation:docxkit-blink 1.06s steps(1) infinite}
.docxkit-viewer:not(.docxkit-focused) .docxkit-caret{display:none}
@keyframes docxkit-blink{50%{opacity:0}}
.docxkit-input{position:absolute;opacity:0;width:1px;height:1em;padding:0;border:0;margin:0;resize:none;overflow:hidden;white-space:pre;pointer-events:none;font-size:16px}
`;

function ensureStyle(): void {
  if (typeof document === 'undefined' || document.getElementById(STYLE_ID)) return;
  const style = document.createElement('style');
  style.id = STYLE_ID;
  style.textContent = CSS;
  document.head.appendChild(style);
}

const IMAGE_MIME: Record<string, string> = {
  png: 'image/png',
  jpg: 'image/jpeg',
  jpeg: 'image/jpeg',
  gif: 'image/gif',
  bmp: 'image/bmp',
  svg: 'image/svg+xml',
  webp: 'image/webp',
};

/** Allow only links whose scheme is on the allow-list. */
export function safeUrl(url: string | undefined, allowed: string[]): string | undefined {
  if (!url) return undefined;
  try {
    const parsed = new URL(url);
    return allowed.includes(parsed.protocol) ? parsed.href : undefined;
  } catch {
    return undefined;
  }
}

export class DocxViewer {
  readonly root: HTMLDivElement;
  private doc: DocxDocument | undefined;
  private layoutResult: LayoutResult | undefined;
  private fonts: FontProvider | undefined;
  private readonly registeredFaces = new Set<string>();
  private readonly imageUrls = new Map<string, string | null>();
  private readonly pagesEl: HTMLDivElement;
  private readonly rows: HTMLDivElement[] = [];
  private readonly mounted = new Set<number>();
  private observer: IntersectionObserver | undefined;
  private comments = new Map<string, CommentInfo>();
  private readonly listeners = new Map<keyof ViewerEvents, Set<Listener<never>>>();
  private unsubscribe: (() => void) | undefined;
  private relayoutScheduled = false;
  private zoomLevel: number;
  private activeComment: string | undefined;
  private editingLayer: InlineEditing | undefined;
  private renderedShape = '';

  constructor(container: HTMLElement, private readonly options: ViewerOptions = {}) {
    ensureStyle();
    this.zoomLevel = options.zoom ?? 1;
    this.root = document.createElement('div');
    this.root.className = 'docxkit-viewer';
    this.pagesEl = document.createElement('div');
    this.root.appendChild(this.pagesEl);
    container.appendChild(this.root);
    document.addEventListener('selectionchange', this.onSelectionChange);
    if (options.editable) this.editingLayer = new InlineEditing(this, options.editable === true ? {} : options.editable);
  }

  /** In-page editing (present when the viewer was created with `editable`). */
  get editing(): InlineEditing | undefined {
    return this.editingLayer;
  }

  /** Turn typing on the pages on or off. */
  setEditable(on: boolean, options: EditingOptions = {}): void {
    if (on && !this.editingLayer) this.editingLayer = new InlineEditing(this, options);
    this.editingLayer?.setEnabled(on);
  }

  get document(): DocxDocument | undefined {
    return this.doc;
  }

  get layout(): LayoutResult | undefined {
    return this.layoutResult;
  }

  on<K extends keyof ViewerEvents>(event: K, listener: Listener<K>): () => void {
    let set = this.listeners.get(event);
    if (!set) this.listeners.set(event, (set = new Set()));
    set.add(listener as Listener<never>);
    return () => set!.delete(listener as Listener<never>);
  }

  /** @internal Used by the editing layer. */
  emit<K extends keyof ViewerEvents>(event: K, payload: ViewerEvents[K]): void {
    for (const l of this.listeners.get(event) ?? []) (l as Listener<K>)(payload);
  }

  async load(input: DocxDocument | Uint8Array | ArrayBuffer | Blob, openOptions?: OpenOptions): Promise<DocxDocument> {
    const doc = input instanceof DocxDocument ? input : await DocxDocument.open(input, openOptions);
    this.unsubscribe?.();
    this.doc = doc;
    this.unsubscribe = doc.onChange(() => this.scheduleRelayout());
    for (const url of this.imageUrls.values()) if (url) URL.revokeObjectURL(url);
    this.imageUrls.clear();
    await this.prepareFonts();
    this.relayout();
    return doc;
  }

  private async prepareFonts(): Promise<void> {
    const registry = this.options.fonts;
    if (!registry || registry.isEmpty) {
      this.fonts = new CanvasFontProvider();
      return;
    }
    const { RegistryFontProvider } = await import('../layout/fonts.js');
    this.fonts = new RegistryFontProvider(registry);
    const loads: Promise<unknown>[] = [];
    for (const face of registry.faces()) {
      if (this.registeredFaces.has(face.id)) continue;
      this.registeredFaces.add(face.id);
      const data = standaloneFont(face.font);
      const ff = new FontFace(`docxkit-${face.id}`, data.buffer.slice(data.byteOffset, data.byteOffset + data.byteLength) as ArrayBuffer);
      document.fonts.add(ff);
      loads.push(ff.load().catch(() => undefined));
    }
    await Promise.all(loads);
  }

  private fontFamily = (font: LayoutFont): string => (font.face ? `docxkit-${font.face.id}` : font.family);

  /** Re-run layout (after edits, zoom or option changes) and redraw. */
  relayout(): void {
    if (!this.doc || !this.fonts) return;
    this.layoutResult = layoutDocument(this.doc, { fonts: this.fonts });
    this.comments = new Map(readComments(this.doc).map((c) => [c.id, c]));
    // Same pages and sizes (the usual case while typing): redraw the mounted pages in place.
    const shape = this.shapeKey();
    if (shape === this.renderedShape && this.rows.length) {
      for (const i of [...this.mounted]) this.mountPage(i);
    } else this.buildRows();
    this.emit('layout', { pages: this.layoutResult.pages });
  }

  private shapeKey(): string {
    const showComments = (this.options.showComments ?? true) && this.comments.size > 0;
    return `${this.scale}|${showComments}|` + (this.layoutResult?.pages.map((p) => `${p.width}x${p.height}`).join(',') ?? '');
  }

  private scheduleRelayout(): void {
    if (this.relayoutScheduled) return;
    this.relayoutScheduled = true;
    queueMicrotask(() => {
      this.relayoutScheduled = false;
      this.relayout();
    });
  }

  get zoom(): number {
    return this.zoomLevel;
  }

  set zoom(value: number) {
    this.zoomLevel = Math.min(5, Math.max(0.25, value));
    this.buildRows();
  }

  /** CSS pixels per point at the current zoom. */
  get scale(): number {
    return (this.zoomLevel * 96) / 72;
  }

  /** The element of a page (undefined before the first layout). */
  pageElement(index: number): HTMLDivElement | undefined {
    return this.rows[index]?.firstElementChild as HTMLDivElement | undefined;
  }

  /** Indexes of the pages currently drawn. */
  get mountedPages(): number[] {
    return [...this.mounted];
  }

  private buildRows(): void {
    const result = this.layoutResult;
    if (!result) return;
    const scrollRatio = this.root.scrollHeight ? this.root.scrollTop / this.root.scrollHeight : 0;
    this.observer?.disconnect();
    this.pagesEl.replaceChildren();
    this.rows.length = 0;
    this.renderedShape = this.shapeKey();
    this.mounted.clear();
    const showComments = (this.options.showComments ?? true) && this.comments.size > 0;
    this.observer = new IntersectionObserver((entries) => this.onIntersect(entries), { root: this.root, rootMargin: '600px 0px' });
    result.pages.forEach((page, i) => {
      const row = document.createElement('div');
      row.className = 'docxkit-row';
      const pageEl = document.createElement('div');
      pageEl.className = 'docxkit-page';
      pageEl.dataset.page = String(i);
      pageEl.style.width = `${page.width * this.scale}px`;
      pageEl.style.height = `${page.height * this.scale}px`;
      pageEl.setAttribute('aria-label', `Page ${i + 1}`);
      row.appendChild(pageEl);
      if (showComments) {
        const gutter = document.createElement('div');
        gutter.className = 'docxkit-gutter';
        gutter.style.height = pageEl.style.height;
        row.appendChild(gutter);
      }
      this.pagesEl.appendChild(row);
      this.rows.push(row);
      this.observer!.observe(row);
    });
    this.root.scrollTop = scrollRatio * this.root.scrollHeight;
  }

  private onIntersect(entries: IntersectionObserverEntry[]): void {
    for (const entry of entries) {
      const i = this.rows.indexOf(entry.target as HTMLDivElement);
      if (i < 0) continue;
      if (entry.isIntersecting && !this.mounted.has(i)) this.mountPage(i);
    }
  }

  private mountPage(i: number): void {
    const page = this.layoutResult?.pages[i];
    const row = this.rows[i];
    if (!page || !row) return;
    this.mounted.add(i);
    const pageEl = row.firstElementChild as HTMLDivElement;
    const missing = this.missingImages(page);
    pageEl.innerHTML = pageToSvg(page, { fontFamily: this.fontFamily, imageUrl: (t, ext) => this.imageUrl(t, ext) });
    pageEl.appendChild(this.overlay(page));
    const gutter = row.children[1] as HTMLDivElement | undefined;
    if (gutter) this.renderComments(page, gutter);
    this.emit('pagemount', { index: i, element: pageEl, page });
    if (missing.length) {
      void this.loadImages(missing).then(() => {
        if (this.layoutResult?.pages[i] === page) {
          this.mounted.delete(i);
          this.mountPage(i);
        }
      });
    }
  }

  private missingImages(page: Page): string[] {
    const out: string[] = [];
    for (const item of page.items) if (item.type === 'image' && item.target && !item.external && !this.imageUrls.has(item.target)) out.push(item.target);
    return out;
  }

  private imageUrl(target: string | undefined, external: boolean): string | undefined {
    if (!target) return undefined;
    if (external) return this.options.allowExternalImages ? safeUrl(target, ['https:', 'http:']) : undefined;
    return this.imageUrls.get(target) ?? undefined;
  }

  private async loadImages(targets: string[]): Promise<void> {
    const doc = this.doc;
    if (!doc) return;
    await Promise.all(
      targets.map(async (target) => {
        if (this.imageUrls.has(target)) return;
        const mime = IMAGE_MIME[extension(target)];
        if (!mime || !doc.pkg.has(target)) {
          this.imageUrls.set(target, null);
          return;
        }
        const bytes = await doc.media(target);
        this.imageUrls.set(target, URL.createObjectURL(new Blob([bytes as BlobPart], { type: mime })));
      }),
    );
  }

  private overlay(page: Page): HTMLDivElement {
    const s = this.scale;
    const layer = document.createElement('div');
    layer.className = 'docxkit-overlay';
    const place = (el: HTMLElement, r: { x: number; y: number; w: number; h: number }) => {
      el.style.left = `${r.x * s}px`;
      el.style.top = `${r.y * s}px`;
      el.style.width = `${Math.max(r.w, 1) * s}px`;
      el.style.height = `${r.h * s}px`;
    };
    for (const perm of page.perms) {
      if (!this.options.highlightRange?.(perm)) continue;
      const el = document.createElement('div');
      el.className = 'docxkit-perm';
      place(el, perm);
      layer.appendChild(el);
    }
    for (const c of page.comments) {
      if (c.w <= 0 || !this.comments.has(c.id)) continue;
      const el = document.createElement('div');
      el.className = 'docxkit-comment-hl' + (this.activeComment === c.id ? ' active' : '');
      el.dataset.comment = c.id;
      place(el, c);
      layer.appendChild(el);
    }
    for (const link of page.links) layer.appendChild(this.linkElement(link, place));
    return layer;
  }

  private linkElement(link: LinkArea, place: (el: HTMLElement, r: LinkArea) => void): HTMLAnchorElement {
    const a = document.createElement('a');
    a.className = 'docxkit-link';
    place(a, link);
    const allowed = this.options.allowedProtocols ?? ['http:', 'https:', 'mailto:', 'tel:'];
    const url = safeUrl(link.target.url, allowed);
    if (url) {
      a.href = url;
      a.target = '_blank';
      a.rel = 'noopener noreferrer nofollow';
    } else if (link.target.anchor) {
      a.href = `#${encodeURIComponent(link.target.anchor)}`;
    }
    const title = link.target.tooltip ?? url ?? (link.target.anchor ? `Go to ${link.target.anchor}` : undefined);
    if (title) a.title = title;
    a.setAttribute('aria-label', title ?? 'Link');
    a.addEventListener('click', (e) => {
      e.preventDefault();
      this.followLink(link.target);
    });
    return a;
  }

  /** Follow a link the way a click does: emit `linkclick`, then open it or jump to the bookmark. */
  followLink(target: LinkTarget): void {
    const url = safeUrl(target.url, this.options.allowedProtocols ?? ['http:', 'https:', 'mailto:', 'tel:']);
    let prevented = false;
    this.emit('linkclick', { target, url, preventDefault: () => (prevented = true) });
    if (prevented) return;
    if (url) window.open(url, '_blank', 'noopener,noreferrer');
    else if (target.anchor) this.scrollToAnchor(target.anchor);
  }

  private renderComments(page: Page, gutter: HTMLDivElement): void {
    gutter.replaceChildren();
    const s = this.scale;
    const roots: Array<{ comment: CommentInfo; y: number }> = [];
    const seen = new Set<string>();
    for (const area of page.comments) {
      const comment = this.comments.get(area.id);
      if (!comment || comment.parentId || seen.has(area.id)) continue;
      const first = this.layoutResult?.commentAnchors.get(area.id);
      if (first && first.page !== page.index) continue;
      seen.add(area.id);
      roots.push({ comment, y: area.y });
    }
    roots.sort((a, b) => a.y - b.y);
    let nextFree = 0;
    for (const { comment, y } of roots) {
      const card = document.createElement('div');
      card.className = 'docxkit-card' + (comment.resolved ? ' resolved' : '');
      card.dataset.comment = comment.id;
      card.appendChild(this.commentBody(comment));
      for (const reply of [...this.comments.values()].filter((c) => c.parentId === comment.id)) {
        const r = this.commentBody(reply);
        r.className = 'reply';
        card.appendChild(r);
      }
      const top = Math.max(y * s, nextFree);
      card.style.top = `${top}px`;
      card.addEventListener('click', () => {
        this.activeComment = comment.id;
        this.emit('commentclick', { comment });
        for (const el of this.root.querySelectorAll('.docxkit-comment-hl')) el.classList.toggle('active', (el as HTMLElement).dataset.comment === comment.id);
      });
      gutter.appendChild(card);
      nextFree = top + card.offsetHeight + 8;
    }
  }

  private commentBody(c: CommentInfo): HTMLDivElement {
    const wrap = document.createElement('div');
    const head = document.createElement('div');
    const who = document.createElement('span');
    who.className = 'who';
    who.textContent = c.author || 'Unknown';
    head.appendChild(who);
    const date = c.dateUtc ?? c.date;
    if (date) {
      const when = document.createElement('span');
      when.className = 'when';
      const d = new Date(date);
      when.textContent = Number.isNaN(d.getTime()) ? date : d.toLocaleString();
      head.appendChild(when);
    }
    const body = document.createElement('div');
    body.className = 'body';
    body.textContent = c.text;
    wrap.append(head, body);
    return wrap;
  }

  scrollToPage(index: number, y = 0): void {
    const row = this.rows[index];
    if (!row) return;
    this.root.scrollTop = row.offsetTop - 16 + y * this.scale;
  }

  /** Scroll to a bookmark. Returns false when the bookmark does not exist. */
  scrollToAnchor(name: string): boolean {
    const anchor = this.layoutResult?.anchors.get(name);
    if (!anchor) return false;
    this.scrollToPage(anchor.page, anchor.y);
    return true;
  }

  scrollToComment(id: string): boolean {
    const anchor = this.layoutResult?.commentAnchors.get(id);
    if (!anchor) return false;
    this.scrollToPage(anchor.page, anchor.rect.y);
    return true;
  }

  /** The current text selection as document positions. */
  getSelection(): { start: InlinePosition; end: InlinePosition } | undefined {
    const pages = this.layoutResult?.pages;
    if (!pages) return undefined;
    const sel = document.getSelection();
    if (!sel || !sel.anchorNode || !this.root.contains(sel.anchorNode)) return undefined;
    return selectionPositions(sel, pages);
  }

  private onSelectionChange = (): void => {
    if (!this.listeners.get('selectionchange')?.size) return;
    this.emit('selectionchange', this.getSelection());
  };

  destroy(): void {
    this.editingLayer?.destroy();
    this.observer?.disconnect();
    this.unsubscribe?.();
    document.removeEventListener('selectionchange', this.onSelectionChange);
    for (const url of this.imageUrls.values()) if (url) URL.revokeObjectURL(url);
    this.root.remove();
  }
}
