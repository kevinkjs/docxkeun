/**
 * DocxViewer: renders a document into a scrollable element, one SVG per page, mounting only the
 * pages near the viewport. Handles links, bookmarks, comment cards and selection mapping.
 */
import { readComments, type CommentInfo } from '../comments/read.js';
import { FontRegistry } from '../fonts/registry.js';
import { standaloneFont } from '../fonts/subset.js';
import { LayoutCache } from '../layout/cache.js';
import { layoutDocument } from '../layout/engine.js';
import type { FontProvider, LayoutFont, LayoutResult, LinkArea, Page, PermArea, Rect } from '../layout/types.js';
import { DocxDocument, type OpenOptions } from '../model/document.js';
import type { LinkTarget } from '../model/fields.js';
import { browserShows, decodeRaster, sniffImage } from '../images/index.js';
import { encodePng } from '../images/png-encode.js';
import { extension } from '../opc/paths.js';
import { CanvasFontProvider } from './canvas-fonts.js';
import { InlineEditing, type CaretFormat, type EditingOptions, type EditStory } from './editing.js';
import { selectionInfoFor, type SelectionInfo } from '../api/selection.js';
import { paragraphsOf, paragraphText, type EditAction } from '../api/editor.js';
import type { DocParagraph } from '../model/content.js';
import { ToolbarController, type ToolbarOptions } from './toolbar.js';
import { NavigationController, type NavigationOptions } from './navigation-panel.js';
import { ContextMenuController, type ContextMenuOptions } from './context-menu.js';
import { CommentsController, type CommentThread, type CommentsOptions } from './comments.js';
import { AiActivity, type ActivityOptions } from './activity.js';
import { collectNavTargets, nextTarget, type DocLocation, type NavKind, type NavTarget, type NavTargets } from './navigation.js';
import { selectionPositions, type InlinePosition } from './selection.js';
import { pageToSvg } from './svg.js';
import { ensureUiStyle } from './ui/theme.js';

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
  /** Shade text in locked regions (default true). */
  showLockedRegions?: boolean;
  /**
   * Let users type directly on the pages. Pass `{ guard }` (e.g. an AccessPolicy) to restrict
   * what they may change; refused edits are reported through the `editerror` event.
   */
  editable?: boolean | EditingOptions;
  /**
   * How an AI assistant's work shows (`viewer.activity`): `false` turns it off, or pass options
   * (`animate: false` for highlights without motion, `marks: false` for no change bars).
   */
  aiActivity?: false | ActivityOptions;
  /** Show a formatting toolbar above the pages while editing (or put it in this element). */
  toolbar?: boolean | HTMLElement | ToolbarOptions;
  /** Show the navigation panel (outline, bookmarks, links) beside the pages. */
  navigation?: boolean | NavigationOptions;
  /** The comment column: `mode` (`auto` switches between cards and bubbles by width) and `showResolved`. `showComments: false` hides it. */
  comments?: CommentsOptions;
  /** Right-click menu (off by default). Shift + right-click still shows the browser's own menu. */
  contextMenu?: boolean | ContextMenuOptions;
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
  /** The viewer jumped somewhere (a link, a bookmark, next/previous, back/forward). */
  navigate: { location: DocLocation; target?: NavTarget };
  /** What is selected (or where the caret is) changed: paragraph numbers, text and pages. */
  selectioninfo: { selection: SelectionInfo | undefined };
  /** The caret moved into a header, footer, note or text box, or back to the body. */
  storychange: { story: EditStory };
  /** The selection or caret moved, or the formatting at it changed (for toolbars). */
  caretchange: { format: CaretFormat | undefined };
  /** A chrome part (toolbar, navigation panel, comment column, context menu) was shown, hidden or folded. */
  chromechange: { part: 'toolbar' | 'navigation' | 'comments' | 'contextMenu'; visible: boolean; collapsed?: boolean };
}

type Listener<K extends keyof ViewerEvents> = (event: ViewerEvents[K]) => void;

const STYLE_ID = 'docxkeun-viewer-style';
const CSS = `
.docxkeun-viewer{position:relative;overflow:auto;background:#e8eaed;height:100%;box-sizing:border-box;padding:16px 0}
.docxkeun-row{display:flex;justify-content:center;align-items:flex-start;gap:12px;margin:0 auto 16px;padding:0 16px;min-width:max-content;box-sizing:border-box}
.docxkeun-page{position:relative;background:#fff;box-shadow:0 1px 3px rgba(60,64,67,.3),0 1px 4px rgba(60,64,67,.15);flex:none}
.docxkeun-page svg{display:block;user-select:text}
.docxkeun-overlay{position:absolute;inset:0;pointer-events:none}
.docxkeun-link{position:absolute;pointer-events:auto;cursor:pointer;display:block}
.docxkeun-link:focus-visible{outline:2px solid #1a73e8;outline-offset:1px}
.docxkeun-comment-hl{position:absolute;background:rgba(251,188,4,.28);border-bottom:2px solid rgba(251,188,4,.9);mix-blend-mode:multiply}
.docxkeun-comment-hl.active{background:rgba(251,188,4,.5)}
.docxkeun-perm{position:absolute;background:rgba(26,115,232,.12);border-bottom:1px dashed rgba(26,115,232,.7)}
.docxkeun-locked{position:absolute;background:repeating-linear-gradient(135deg,rgba(95,99,104,.10) 0 3px,rgba(95,99,104,.03) 3px 7px);box-shadow:inset 0 -1px 0 rgba(95,99,104,.55)}
.docxkeun-gutter{flex:none;height:1px}
.docxkeun-editing .docxkeun-page svg{user-select:none;cursor:text}
.docxkeun-editing .docxkeun-link{pointer-events:none}
.docxkeun-edit-layer{position:absolute;inset:0;pointer-events:none}
.docxkeun-selection{position:absolute;background:rgba(26,115,232,.28)}
.docxkeun-caret{position:absolute;width:2px;margin-left:-1px;background:#202124;animation:docxkeun-blink 1.06s steps(1) infinite}
.docxkeun-viewer:not(.docxkeun-focused) .docxkeun-caret{display:none}
@keyframes docxkeun-blink{50%{opacity:0}}
.docxkeun-flash{position:absolute;pointer-events:none;border-radius:2px;background:rgba(251,188,4,.45);box-shadow:0 0 0 2px rgba(251,188,4,.8);animation:docxkeun-flash 1.6s ease-out forwards}
@keyframes docxkeun-flash{0%,35%{opacity:1}100%{opacity:0}}
.docxkeun-input{position:absolute;opacity:0;width:1px;height:1em;padding:0;border:0;margin:0;resize:none;overflow:hidden;white-space:pre;pointer-events:none;font-size:16px}
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
  jpeg: 'image/jpeg',
  gif: 'image/gif',
  bmp: 'image/bmp',
  svg: 'image/svg+xml',
  webp: 'image/webp',
};

/**
 * Media bytes as something an <img> shows: the format is sniffed (extensions lie); TIFF is
 * decoded to PNG, and EMF/WMF are drawn into SVG. Undefined for images that cannot be shown.
 */
async function browserImage(bytes: Uint8Array, ext: string, _fonts: FontProvider | undefined): Promise<{ bytes: Uint8Array; mime: string } | undefined> {
  const kind = sniffImage(bytes);
  if (kind && browserShows(kind)) return { bytes, mime: IMAGE_MIME[kind]! };
  try {
    if (kind === 'tiff') return { bytes: encodePng(await decodeRaster(bytes)), mime: 'image/png' };
  } catch {
    return undefined;
  }
  // Formats we cannot sniff (WebP): trust the extension.
  const mime = IMAGE_MIME[ext === 'jpg' ? 'jpeg' : ext];
  return mime && !kind ? { bytes, mime } : undefined;
}

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
  private commentMap = new Map<string, CommentInfo>();
  private threadCache: CommentThread[] | undefined;
  private readonly listeners = new Map<keyof ViewerEvents, Set<Listener<never>>>();
  private unsubscribe: (() => void) | undefined;
  private relayoutScheduled = false;
  private layoutCache = new LayoutCache();
  private zoomLevel: number;
  private activeComments = new Set<string>();
  private editingLayer: InlineEditing | undefined;
  private toolbarCtl: ToolbarController;
  private navigationCtl: NavigationController;
  private commentsCtl: CommentsController;
  private contextMenuCtl: ContextMenuController;
  private readonly containerEl: HTMLElement;
  private renderedShape = '';
  private navTargets: NavTargets | undefined;
  private navCursor: { kind: NavKind; index: number; scrollTop: number } | undefined;
  private readonly past: DocLocation[] = [];
  private readonly future: DocLocation[] = [];

  constructor(container: HTMLElement, private readonly options: ViewerOptions = {}) {
    ensureStyle();
    this.containerEl = container;
    this.zoomLevel = options.zoom ?? 1;
    this.root = document.createElement('div');
    this.root.className = 'docxkeun-viewer';
    this.pagesEl = document.createElement('div');
    this.root.appendChild(this.pagesEl);
    container.appendChild(this.root);
    document.addEventListener('selectionchange', this.onSelectionChange);
    if (options.editable) this.editingLayer = new InlineEditing(this, options.editable === true ? {} : options.editable);
    const tb = options.toolbar;
    const tbOptions: ToolbarOptions = !tb || tb === true ? {} : 'nodeType' in tb ? { container: tb } : tb;
    this.toolbarCtl = new ToolbarController(this, tbOptions);
    if (tb) {
      if (options.editable) this.toolbarCtl.show();
      else this.toolbarCtl.deferShow();
    }
    const nav = options.navigation;
    this.navigationCtl = new NavigationController(this, !nav || nav === true ? {} : nav);
    if (nav) this.navigationCtl.show();
    this.commentsCtl = new CommentsController(this.commentsHost(), { ...options.comments, visible: options.showComments ?? true });
    const menu = options.contextMenu;
    this.contextMenuCtl = new ContextMenuController(this, !menu || menu === true ? {} : menu);
    if (menu) this.contextMenuCtl.enable();
  }

  /** The threads of the document, in order: each root comment with its replies and anchor. */
  private commentThreads(): CommentThread[] {
    if (this.threadCache) return this.threadCache;
    const layout = this.layoutResult;
    if (!layout) return [];
    const areas = new Map<string, Array<{ page: number; rect: Rect }>>();
    for (const page of layout.pages) {
      for (const a of page.comments) if (a.w > 0) (areas.get(a.id) ?? areas.set(a.id, []).get(a.id)!).push({ page: page.index, rect: a });
    }
    const replies = new Map<string, CommentInfo[]>();
    for (const c of this.commentMap.values()) if (c.parentId) (replies.get(c.parentId) ?? replies.set(c.parentId, []).get(c.parentId)!).push(c);
    const threads: CommentThread[] = [];
    for (const comment of this.commentMap.values()) {
      const anchor = layout.commentAnchors.get(comment.id);
      if (comment.parentId || !anchor) continue;
      const others = (areas.get(comment.id) ?? []).filter((a) => a.page === anchor.page && a.rect !== anchor.rect).map((a) => a.rect);
      threads.push({ comment, replies: replies.get(comment.id) ?? [], anchor: { page: anchor.page, y: anchor.rect.y, rects: [anchor.rect, ...others] } });
    }
    threads.sort((a, b) => a.anchor.page - b.anchor.page || a.anchor.y - b.anchor.y);
    return (this.threadCache = threads);
  }

  /** What the comment column needs from this viewer. */
  private commentsHost(): ConstructorParameters<typeof CommentsController>[0] {
    const guardAllows = (action: EditAction, detail?: { commentAuthor?: string }): boolean => {
      const guard = this.editingLayer?.guard;
      if (!guard) return true;
      try {
        guard.assert(action, undefined, detail);
        return true;
      } catch {
        return false;
      }
    };
    return {
      root: this.root,
      on: (event, listener) => this.on(event, listener),
      emit: (event, payload) => this.emit(event, payload),
      threads: () => this.commentThreads(),
      anchorOf: (thread) => {
        const page = this.pageElement(thread.anchor.page);
        const rect = thread.anchor.rects[0];
        if (!page || !rect) return undefined;
        // The last line of the comment's text on that page (the rects are its highlighted areas there).
        const last = thread.anchor.rects.reduce((m, r) => Math.max(m, r.y + r.h), rect.y + rect.h);
        return { top: page.offsetTop + rect.y * this.scale, pageRight: page.offsetLeft + page.offsetWidth, height: rect.h * this.scale, bottom: page.offsetTop + last * this.scale };
      },
      columnX: () => {
        let right = 0;
        for (const row of this.rows) {
          const page = row.firstElementChild as HTMLElement | null;
          if (page) right = Math.max(right, page.offsetLeft + page.offsetWidth);
        }
        return right + 12;
      },
      widestPage: () => ((this.layoutResult?.pages.reduce((m, p) => Math.max(m, p.width), 0) ?? 0) * 96) / 72,
      editable: () => !!this.editingLayer?.isEnabled,
      can: guardAllows,
      author: () => this.editingLayer?.author ?? 'docxkeun',
      transaction: (label, fn) => {
        if (!this.editingLayer) return Promise.reject(new Error('Comments can only be changed while the viewer is editable'));
        return this.editingLayer.transaction(label, fn);
      },
      selection: () => {
        const editing = this.editingLayer;
        const sel = editing?.selection;
        const spot = editing?.selectionSpot();
        const page = spot ? this.pageElement(spot.page) : undefined;
        if (!editing || !sel || !spot || !page || (sel.anchor.para === sel.focus.para && sel.anchor.char === sel.focus.char)) return undefined;
        return { range: editing.rangeOf(sel.anchor, sel.focus), top: page.offsetTop + spot.y * this.scale, pageRight: page.offsetLeft + page.offsetWidth, height: spot.h * this.scale };
      },
      setActive: (ids) => {
        this.activeComments = new Set(ids);
        for (const hl of this.root.querySelectorAll<HTMLElement>('.docxkeun-comment-hl')) hl.classList.toggle('active', this.activeComments.has(hl.dataset.comment ?? ''));
      },
      rebuild: () => this.buildRows(),
      currentDocument: () => this.doc,
    };
  }

  private frameParts: { frame: HTMLDivElement; toolbarSlot: HTMLDivElement; body: HTMLDivElement } | undefined;

  /**
   * The frame around the viewer, created on first call: `container > .docxkeun-frame >
   * [.docxkeun-toolbar-slot, .docxkeun-body > root]`. Only chrome parts call it; without them
   * the root stays a direct child of the container.
   * @internal
   */
  frame(): { frame: HTMLDivElement; toolbarSlot: HTMLDivElement; body: HTMLDivElement } {
    if (this.frameParts) return this.frameParts;
    ensureUiStyle();
    const frame = document.createElement('div');
    frame.className = 'docxkeun-frame';
    const toolbarSlot = document.createElement('div');
    toolbarSlot.className = 'docxkeun-toolbar-slot';
    const body = document.createElement('div');
    body.className = 'docxkeun-body';
    frame.append(toolbarSlot, body);
    this.root.replaceWith(frame);
    body.appendChild(this.root);
    return (this.frameParts = { frame, toolbarSlot, body });
  }

  private activityLayer: AiActivity | undefined;

  /**
   * An AI assistant's presence on the pages: `work([4, 5])` shimmers the paragraphs it is
   * changing (with its caret), `reveal(before)` animates what changed. Created on first use.
   */
  get activity(): AiActivity {
    const o = this.options.aiActivity;
    return (this.activityLayer ??= new AiActivity(this, o === false ? { enabled: false } : (o ?? {})));
  }

  /** The formatting toolbar: `show()`, `hide()`, `collapsed`. Shown when the `toolbar` option is set. */
  get toolbar(): ToolbarController {
    return this.toolbarCtl;
  }

  /** The comment column: `show()`, `hide()`, `mode`, `open(id)`, `compose()`, `width`. */
  get comments(): CommentsController {
    return this.commentsCtl;
  }

  /** The right-click menu: `enable()`, `disable()`, `open(x, y)`, `close()`. Enabled by the `contextMenu` option. */
  get contextMenu(): ContextMenuController {
    return this.contextMenuCtl;
  }

  /** The element the viewer was created in (where chrome such as the menu is placed). @internal */
  get container(): HTMLElement {
    return this.containerEl;
  }

  /** A picture as a blob a browser shows (for copying it). Undefined for external or unreadable images. @internal */
  async imageBlob(image: { target: string | undefined; external: boolean }): Promise<Blob | undefined> {
    const doc = this.doc;
    if (!doc || !image.target || image.external) return undefined;
    const shown = await browserImage(await doc.media(image.target), extension(image.target), this.fonts).catch(() => undefined);
    return shown ? new Blob([shown.bytes as BlobPart], { type: shown.mime }) : undefined;
  }

  /** In-page editing (present when the viewer was created with `editable`). */
  get editing(): InlineEditing | undefined {
    return this.editingLayer;
  }

  /** Turn typing on the pages on or off. */
  setEditable(on: boolean, options: EditingOptions = {}): void {
    if (on && !this.editingLayer) this.editingLayer = new InlineEditing(this, options);
    this.editingLayer?.setEnabled(on);
    this.toolbarCtl.setEditable(on);
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
    this.layoutCache = new LayoutCache();
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
      const ff = new FontFace(`docxkeun-${face.id}`, data.buffer.slice(data.byteOffset, data.byteOffset + data.byteLength) as ArrayBuffer);
      document.fonts.add(ff);
      loads.push(ff.load().catch(() => undefined));
    }
    await Promise.all(loads);
  }

  private fontFamily = (font: LayoutFont): string => (font.face ? `docxkeun-${font.face.id}` : font.family);

  /** Re-run layout (after edits, zoom or option changes) and redraw. */
  relayout(): void {
    if (!this.doc || !this.fonts) return;
    this.layoutResult = layoutDocument(this.doc, { fonts: this.fonts, cache: this.layoutCache });
    this.navTargets = undefined;
    this.commentMap = new Map(readComments(this.doc).map((c) => [c.id, c]));
    this.threadCache = undefined;
    this.commentsCtl.sync();
    // Same pages and sizes (the usual case while typing): redraw the mounted pages in place.
    const shape = this.shapeKey();
    if (shape === this.renderedShape && this.rows.length) {
      for (const i of [...this.mounted]) this.mountPage(i);
    } else this.buildRows();
    this.emit('layout', { pages: this.layoutResult.pages });
  }

  private shapeKey(): string {
    return `${this.scale}|${this.commentsCtl.width}|` + (this.layoutResult?.pages.map((p) => `${p.width}x${p.height}`).join(',') ?? '');
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
    this.commentsCtl.refresh();
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
    const gutterWidth = Math.max(0, this.commentsCtl.width - 12);
    this.observer = new IntersectionObserver((entries) => this.onIntersect(entries), { root: this.root, rootMargin: '600px 0px' });
    result.pages.forEach((page, i) => {
      const row = document.createElement('div');
      row.className = 'docxkeun-row';
      const pageEl = document.createElement('div');
      pageEl.className = 'docxkeun-page';
      pageEl.dataset.page = String(i);
      pageEl.style.width = `${page.width * this.scale}px`;
      pageEl.style.height = `${page.height * this.scale}px`;
      pageEl.setAttribute('aria-label', `Page ${i + 1}`);
      row.appendChild(pageEl);
      if (gutterWidth) {
        const gutter = document.createElement('div');
        gutter.className = 'docxkeun-gutter';
        gutter.style.width = `${gutterWidth}px`;
        row.appendChild(gutter);
      }
      this.pagesEl.appendChild(row);
      this.rows.push(row);
      this.observer!.observe(row);
    });
    this.root.scrollTop = scrollRatio * this.root.scrollHeight;
    this.commentsCtl.rowsBuilt();
    // The AI's marks and effects sit at page positions: follow the new rows (zoom, column width).
    this.activityLayer?.rowsBuilt();
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
    for (const item of page.items) if (item.type === 'image' && !item.frameOnly && item.target && !item.external && !this.imageUrls.has(item.target)) out.push(item.target);
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
        if (!doc.pkg.has(target)) {
          this.imageUrls.set(target, null);
          return;
        }
        const shown = await browserImage(await doc.media(target), extension(target), this.fonts);
        this.imageUrls.set(target, shown ? URL.createObjectURL(new Blob([shown.bytes as BlobPart], { type: shown.mime })) : null);
      }),
    );
  }

  private overlay(page: Page): HTMLDivElement {
    const s = this.scale;
    const layer = document.createElement('div');
    layer.className = 'docxkeun-overlay';
    const place = (el: HTMLElement, r: { x: number; y: number; w: number; h: number }) => {
      el.style.left = `${r.x * s}px`;
      el.style.top = `${r.y * s}px`;
      el.style.width = `${Math.max(r.w, 1) * s}px`;
      el.style.height = `${r.h * s}px`;
    };
    if (this.options.showLockedRegions ?? true) {
      for (const lock of page.locks) {
        const el = document.createElement('div');
        el.className = 'docxkeun-locked';
        el.dataset.lock = lock.id;
        place(el, lock);
        layer.appendChild(el);
      }
    }
    for (const perm of page.perms) {
      if (!this.options.highlightRange?.(perm)) continue;
      const el = document.createElement('div');
      el.className = 'docxkeun-perm';
      place(el, perm);
      layer.appendChild(el);
    }
    for (const c of page.comments) {
      if (c.w <= 0 || !this.commentMap.has(c.id)) continue;
      const el = document.createElement('div');
      el.className = 'docxkeun-comment-hl' + (this.activeComments.has(c.id) ? ' active' : '');
      el.dataset.comment = c.id;
      place(el, c);
      layer.appendChild(el);
    }
    for (const link of page.links) layer.appendChild(this.linkElement(link, place));
    return layer;
  }

  private linkElement(link: LinkArea, place: (el: HTMLElement, r: LinkArea) => void): HTMLAnchorElement {
    const a = document.createElement('a');
    a.className = 'docxkeun-link';
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
    else if (target.anchor) {
      const bookmark = this.collectedTargets().bookmarks.find((b) => b.name === target.anchor);
      // Remember where the link was, so back() returns to it.
      if (bookmark) this.goTo(bookmark, { history: true });
    }
  }

  scrollToPage(index: number, y = 0): void {
    const row = this.rows[index];
    if (!row) return;
    this.root.scrollTop = row.offsetTop - 16 + y * this.scale;
  }

  /** Scroll to a bookmark. Returns false when the bookmark does not exist. */
  scrollToAnchor(name: string): boolean {
    const bookmark = this.collectedTargets().bookmarks.find((b) => b.name === name);
    if (!bookmark) return false;
    this.goTo(bookmark, { flash: false });
    return true;
  }

  // ---- navigation -----------------------------------------------------------------------------

  /**
   * The navigation panel: `show()`, `hide()`, `collapsed`, `tab`. It also lists the headings,
   * bookmarks and hyperlinks of the current layout (`headings`, `bookmarks`, `links`), each in
   * document order.
   */
  get navigation(): NavigationController {
    return this.navigationCtl;
  }

  /** The navigation targets of the current layout, collected once per layout. @internal */
  collectedTargets(): NavTargets {
    if (!this.navTargets) this.navTargets = this.layoutResult ? collectNavTargets(this.layoutResult) : { headings: [], bookmarks: [], links: [] };
    return this.navTargets;
  }

  /** The location at the top of the viewport. */
  get location(): DocLocation {
    const top = this.root.scrollTop + 16;
    let page = 0;
    for (let i = 0; i < this.rows.length; i++) {
      if (this.rows[i]!.offsetTop <= top + 0.5) page = i;
      else break;
    }
    const row = this.rows[page];
    return { page, x: 0, y: row ? Math.max(0, (top - row.offsetTop) / this.scale) : 0 };
  }

  /**
   * Jump to a target (or any location). The target briefly lights up. With `history`, back()
   * returns to where the viewer was before the jump.
   */
  goTo(target: NavTarget | DocLocation, options: { history?: boolean; flash?: boolean } = {}): void {
    if (options.history) {
      this.past.push(this.location);
      this.future.length = 0;
    }
    // Leave some room above the target so it is not glued to the edge.
    this.scrollToPage(target.page, Math.max(0, target.y - 24));
    const nav = 'kind' in target ? target : undefined;
    if (nav) {
      const list = this.collectedTargets()[`${nav.kind}s` as keyof NavTargets];
      this.navCursor = { kind: nav.kind, index: list.indexOf(nav), scrollTop: this.root.scrollTop };
      if (options.flash ?? true) this.flash(nav);
    }
    this.emit('navigate', { location: { page: target.page, x: target.x, y: target.y }, target: nav });
  }

  /** Jump to the next heading, bookmark or link after the current one (or the viewport top). */
  next(kind: NavKind, options: { includeHidden?: boolean } = {}): NavTarget | undefined {
    return this.step(kind, 1, options.includeHidden ?? false);
  }

  /** Jump to the previous heading, bookmark or link. */
  previous(kind: NavKind, options: { includeHidden?: boolean } = {}): NavTarget | undefined {
    return this.step(kind, -1, options.includeHidden ?? false);
  }

  private step(kind: NavKind, dir: 1 | -1, includeHidden: boolean): NavTarget | undefined {
    const all = this.collectedTargets()[`${kind}s` as keyof NavTargets];
    const list = includeHidden ? all : all.filter((t) => !t.hidden);
    let target: NavTarget | undefined;
    const cursor = this.navCursor;
    if (cursor && cursor.kind === kind && Math.abs(this.root.scrollTop - cursor.scrollTop) < 2 && all[cursor.index]) {
      // Not scrolled since the last jump: continue from that target, by index so that two
      // links on the same line are both visited.
      const from = all[cursor.index]!;
      const i = list.indexOf(from);
      target = i >= 0 ? list[i + dir] : nextTarget(list, from, dir, 0);
    } else {
      target = nextTarget(list, this.location, dir);
    }
    if (target) this.goTo(target);
    return target;
  }

  get canGoBack(): boolean {
    return this.past.length > 0;
  }

  get canGoForward(): boolean {
    return this.future.length > 0;
  }

  /** Return to where the viewer was before the last jump (like a browser's Back). */
  back(): boolean {
    const to = this.past.pop();
    if (!to) return false;
    this.future.push(this.location);
    this.scrollToPage(to.page, to.y);
    this.navCursor = undefined;
    this.emit('navigate', { location: to });
    return true;
  }

  forward(): boolean {
    const to = this.future.pop();
    if (!to) return false;
    this.past.push(this.location);
    this.scrollToPage(to.page, to.y);
    this.navCursor = undefined;
    this.emit('navigate', { location: to });
    return true;
  }

  /** Briefly highlight a target's areas. */
  flash(target: Pick<NavTarget, 'page' | 'rects'>): void {
    const pageEl = this.pageElement(target.page);
    if (!pageEl) return;
    const s = this.scale;
    for (const r of target.rects) {
      const el = document.createElement('div');
      el.className = 'docxkeun-flash';
      // Positioned in the scroll container so page re-renders do not remove it.
      el.style.left = `${pageEl.offsetLeft + r.x * s - 2}px`;
      el.style.top = `${pageEl.offsetTop + r.y * s - 1}px`;
      el.style.width = `${Math.max(r.w, 2) * s + 4}px`;
      el.style.height = `${r.h * s + 2}px`;
      el.addEventListener('animationend', () => el.remove());
      this.root.appendChild(el);
    }
  }

  scrollToComment(id: string): boolean {
    const anchor = this.layoutResult?.commentAnchors.get(id);
    if (!anchor) return false;
    this.scrollToPage(anchor.page, anchor.rect.y);
    this.commentsCtl.open(id);
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
    this.selectionMoved();
    if (!this.listeners.get('selectionchange')?.size) return;
    this.emit('selectionchange', this.getSelection());
  };

  private selectionTimer: ReturnType<typeof setTimeout> | undefined;
  private lastSelectionKey = '';

  /** @internal The caret or the selection moved: tell `selectioninfo` listeners (batched, when it changed). */
  selectionMoved(): void {
    if (!this.listeners.get('selectioninfo')?.size || this.selectionTimer) return;
    // A timer rather than an animation frame: pages in background tabs still report selections.
    this.selectionTimer = setTimeout(() => {
      this.selectionTimer = undefined;
      const selection = this.selectionInfo();
      const key = selection ? `${selection.story}|${selection.part ?? ''}|${selection.start.paragraph}:${selection.start.offset}|${selection.end.paragraph}:${selection.end.offset}|${selection.text.length}` : '';
      if (key === this.lastSelectionKey) return;
      this.lastSelectionKey = key;
      this.emit('selectioninfo', { selection });
    }, 16);
  }

  /**
   * What the user has selected, or where the caret is: paragraph numbers (as docxkeun-mcp's
   * read_document numbers them), offsets, pages and text. The editing caret when typing is on,
   * otherwise the browser's text selection on the pages. `describeSelection` turns it into a prompt.
   */
  selectionInfo(): SelectionInfo | undefined {
    const doc = this.doc;
    if (!doc) return undefined;
    const typed = this.editingLayer?.isEnabled ? this.editingLayer.selectionInfo() : undefined;
    if (typed) return typed;
    const sel = this.getSelection();
    if (!sel) return undefined;
    const layout = this.layoutResult;
    // The paragraphs of each story, looked up by the text nodes the selection points at.
    const stories: Array<{ story: SelectionInfo['story']; part?: string; paras: DocParagraph[] }> = [{ story: 'body', paras: paragraphsOf(doc.body().blocks) }];
    for (const sp of doc.storyParts()) {
      if (sp.kind === 'header' || sp.kind === 'footer') stories.push({ story: sp.kind, part: sp.part, paras: paragraphsOf(doc.headerFooter(sp.part).blocks) });
    }
    const locate = (pos: InlinePosition) => {
      for (const s of stories) {
        for (let i = 0; i < s.paras.length; i++) {
          for (const seg of paragraphText(s.paras[i]!, { objects: true }).segments) {
            if (seg.node !== pos.inline.node) continue;
            return { s, caret: { para: i, char: seg.start + (seg.node.name === 'w:t' ? pos.offset : pos.offset > 0 ? 1 : 0) } };
          }
        }
      }
      return undefined;
    };
    const a = locate(sel.start);
    const b = locate(sel.end);
    if (!a || !b || a.s !== b.s) return undefined;
    return selectionInfoFor(doc, a.s.paras, a.caret, b.caret, { story: a.s.story, ...(a.s.part ? { part: a.s.part } : {}), ...(layout ? { layout } : {}) });
  }

  destroy(): void {
    this.activityLayer?.destroy();
    this.contextMenuCtl.destroy();
    this.toolbarCtl.destroy();
    this.navigationCtl.destroy();
    this.commentsCtl.destroy();
    this.editingLayer?.destroy();
    this.observer?.disconnect();
    this.unsubscribe?.();
    document.removeEventListener('selectionchange', this.onSelectionChange);
    for (const url of this.imageUrls.values()) if (url) URL.revokeObjectURL(url);
    this.frameParts ? this.frameParts.frame.remove() : this.root.remove();
  }
}
