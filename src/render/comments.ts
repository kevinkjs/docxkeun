/**
 * The comment column beside the pages: one layer for the whole document, with a card per thread
 * placed at its anchor (cards mode) or a strip of bubbles (bubbles mode, when the viewer is too
 * narrow). `viewer.comments` is its controller. The controller takes a small host (the viewer, or a
 * stub in tests) that supplies the threads, where they sit and how edits are made.
 */
import type { CommentInfo } from '../comments/read.js';
import type { DocRange } from '../api/editor.js';
import type { DocxEditor } from '../api/editor.js';
import type { Rect } from '../layout/types.js';
import { REVISION_COLORS } from '../model/content.js';
import type { DocxViewer } from './viewer.js';
import { icon } from './ui/icons.js';
import { BUBBLE_SIZE, bubbleGroups, commentMode, layoutCards, type CardBox } from './ui/comment-layout.js';
import { ensureUiStyle } from './ui/theme.js';

export interface CommentThread {
  /** The first comment of the thread. */
  comment: CommentInfo;
  replies: CommentInfo[];
  /** Where the thread is anchored: the page, the top of its first line (points) and its highlighted areas. */
  anchor: { page: number; y: number; rects: Rect[] };
}

export type CommentsMode = 'auto' | 'cards' | 'bubbles';

export interface CommentsOptions {
  /** `auto` picks cards or bubbles from the viewer's width. Default: auto. */
  mode?: CommentsMode;
  /** Show resolved threads (faded). Default: false. */
  showResolved?: boolean;
}

export interface CommentsControllerOptions extends CommentsOptions {
  /** Start shown (default true). */
  visible?: boolean;
}

type GuardedAction = 'comment' | 'reply' | 'resolveComment' | 'deleteComment';

/** What the column needs from the viewer. */
export interface CommentsHost extends Pick<DocxViewer, 'root' | 'on' | 'emit'> {
  /** Every thread (resolved ones too), in document order. */
  threads(): CommentThread[];
  /**
   * Where a thread's highlight is, in the scroll container's pixels: the top and height of its
   * first line, `bottom` of its last line on that page (default `top + height`), and `pageRight`,
   * the right edge of its page (where the connector starts).
   */
  anchorOf(thread: CommentThread): CommentSpot | undefined;
  /** Left edge of the column in the scroll container's pixels. */
  columnX(): number;
  /** The widest page at 100% zoom, in CSS px. */
  widestPage(): number;
  editable(): boolean;
  /** Whether the guard allows an action (true without a guard). */
  can(action: GuardedAction, detail?: { commentAuthor?: string }): boolean;
  /** Who new comments and replies are written by. */
  author(): string;
  /** One undo step; a failure restores the document and rethrows. */
  transaction<T>(label: string, fn: (editor: DocxEditor) => T | Promise<T>): Promise<T>;
  /** The selection to comment on, with where it is. */
  selection(): ({ range: DocRange } & CommentSpot) | undefined;
  /** Mark the highlights of these comments as active. */
  setActive(ids: string[]): void;
  /** The column's width changed: rebuild the rows (which read `width`). */
  rebuild(): void;
  /** The open document (any value that changes when another one is loaded). */
  currentDocument?(): unknown;
}

/** Where a thread's text (or the selection) is, in the scroll container's pixels. */
export interface CommentSpot {
  top: number;
  pageRight: number;
  height: number;
  bottom?: number;
}

const STYLE_ID = 'docxkeun-comments-style';
const COLUMN = 260;
const STRIP = 36;
const GAP = 12;
const CARD_GAP = 8;
const SWITCH_FLOOR = 32;
/** Space between a bubble popover and the text, and from the edges of the view. */
const POP_GAP = 6;
const EDGE = 4;
const DRAFT = '__draft';

const CSS = `
.docxkeun-comments{position:absolute;top:0;z-index:3;pointer-events:none;font:var(--dk-ui-font);color:var(--dk-ui-fg)}
.docxkeun-comments button{font:inherit;color:inherit}
.docxkeun-comments .dk-card{position:absolute;left:0;right:0;box-sizing:border-box;pointer-events:auto;background:var(--dk-ui-bg);border:1px solid var(--dk-ui-line);border-radius:var(--dk-ui-radius);padding:8px 10px;cursor:pointer;box-shadow:0 1px 2px rgba(60,64,67,.25)}
@media (prefers-reduced-motion:no-preference){.docxkeun-comments .dk-card{transition:top .18s ease}}
.docxkeun-comments .dk-card:hover{border-color:var(--dk-ui-muted)}
.docxkeun-comments .dk-card:focus-visible,.docxkeun-comments button:focus-visible,.docxkeun-comments textarea:focus-visible{outline:2px solid var(--dk-ui-accent);outline-offset:1px}
.docxkeun-comments .dk-card.open,.docxkeun-comments .dk-card.dk-draft{cursor:default;border-color:var(--dk-ui-accent);box-shadow:var(--dk-ui-shadow)}
.docxkeun-comments .dk-card.resolved{opacity:.6}
.docxkeun-comments .dk-head{display:flex;align-items:center;gap:6px;min-width:0}
.docxkeun-comments .dk-avatar{flex:none;width:22px;height:22px;border-radius:50%;color:#fff;font-size:10px;font-weight:600;display:inline-flex;align-items:center;justify-content:center}
.docxkeun-comments .dk-who{font-weight:600;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
.docxkeun-comments .dk-when{flex:none;margin-left:auto;font-size:12px;color:var(--dk-ui-muted)}
.docxkeun-comments .dk-text{margin-top:4px;white-space:pre-wrap;overflow-wrap:anywhere}
.docxkeun-comments .dk-clamp{display:-webkit-box;-webkit-box-orient:vertical;-webkit-line-clamp:2;line-clamp:2;overflow:hidden}
.docxkeun-comments .dk-replies{margin-top:4px;font-size:12px;color:var(--dk-ui-accent)}
.docxkeun-comments .dk-reply{border-top:1px solid var(--dk-ui-line);margin-top:8px;padding-top:8px}
.docxkeun-comments .dk-connector{position:absolute;right:100%;height:1px;background:var(--dk-ui-accent);pointer-events:none}
.docxkeun-comments .dk-actions,.docxkeun-comments .dk-composer-buttons{display:flex;gap:4px;margin-top:8px}
.docxkeun-comments .dk-actions button,.docxkeun-comments .dk-composer-buttons button,.docxkeun-comments .dk-resolved-switch,.docxkeun-comments .dk-pop-nav button{height:26px;min-width:26px;padding:0 8px;border:1px solid var(--dk-ui-line);border-radius:4px;background:transparent;cursor:pointer;display:inline-flex;align-items:center;justify-content:center}
.docxkeun-comments button:hover:not(:disabled){background:var(--dk-ui-hover)}
.docxkeun-comments button:disabled{opacity:.4;cursor:default}
.docxkeun-comments .dk-composer-buttons .dk-post{background:var(--dk-ui-accent);border-color:var(--dk-ui-accent);color:#fff}
.docxkeun-comments .dk-composer-buttons .dk-post:hover:not(:disabled){background:var(--dk-ui-accent)}
.docxkeun-comments textarea{display:block;width:100%;box-sizing:border-box;min-height:56px;margin-top:8px;padding:6px 8px;border:1px solid var(--dk-ui-line);border-radius:4px;background:var(--dk-ui-bg);color:inherit;font:inherit;resize:vertical}
.docxkeun-comments .dk-resolved-switch{position:absolute;top:4px;right:0;pointer-events:auto;font-size:12px;color:var(--dk-ui-muted);background:var(--dk-ui-bg)}
.docxkeun-comments .dk-resolved-switch[aria-checked=true]{color:var(--dk-ui-fg);background:var(--dk-ui-accent-soft)}
.docxkeun-comments .dk-bubble{position:absolute;left:4px;width:${BUBBLE_SIZE}px;height:${BUBBLE_SIZE}px;padding:0;border-radius:50%;border:1px solid var(--dk-ui-line);background:var(--dk-ui-bg);cursor:pointer;pointer-events:auto;box-shadow:0 1px 2px rgba(60,64,67,.25);font-size:12px;font-weight:600;display:inline-flex;align-items:center;justify-content:center;color:var(--dk-ui-accent)}
.docxkeun-comments .dk-bubble.open{background:var(--dk-ui-accent);border-color:var(--dk-ui-accent);color:#fff}
.docxkeun-comments .dk-pop{position:absolute;width:${COLUMN}px;pointer-events:auto}
.docxkeun-comments .dk-pop .dk-card{position:relative;left:auto;right:auto;transition:none}
.docxkeun-comments .dk-pop-nav{display:flex;align-items:center;justify-content:flex-end;gap:6px;margin-bottom:4px;font-size:12px;color:var(--dk-ui-muted)}
`;

function ensureStyle(): void {
  ensureUiStyle();
  if (typeof document === 'undefined' || document.getElementById(STYLE_ID)) return;
  const style = document.createElement('style');
  style.id = STYLE_ID;
  style.textContent = CSS;
  document.head.appendChild(style);
}

/** "just now", "5m ago", "3h ago", "2d ago", else a short date. Empty without a date. */
export function relativeTime(date: string | undefined, now = Date.now()): string {
  if (!date) return '';
  const t = Date.parse(date);
  if (Number.isNaN(t)) return date;
  const s = Math.max(0, (now - t) / 1000);
  if (s < 45) return 'just now';
  if (s < 3600) return `${Math.max(1, Math.round(s / 60))}m ago`;
  if (s < 86400) return `${Math.round(s / 3600)}h ago`;
  if (s < 7 * 86400) return `${Math.round(s / 86400)}d ago`;
  const d = new Date(t);
  return d.toLocaleDateString(undefined, { month: 'short', day: 'numeric', ...(d.getFullYear() === new Date(now).getFullYear() ? {} : { year: 'numeric' }) });
}

function initialsOf(c: CommentInfo): string {
  const given = c.initials?.trim();
  if (given) return given.slice(0, 3).toUpperCase();
  const words = c.author.trim().split(/\s+/).filter(Boolean);
  return (words.length > 1 ? words[0]![0]! + words[words.length - 1]![0]! : (words[0]?.slice(0, 2) ?? '?')).toUpperCase();
}

function el<K extends keyof HTMLElementTagNameMap>(tag: K, className?: string, text?: string): HTMLElementTagNameMap[K] {
  const e = document.createElement(tag);
  if (className) e.className = className;
  if (text !== undefined) e.textContent = text;
  return e;
}

function signature(t: CommentThread, open: boolean, replying: boolean, actions: object): string {
  const part = (c: CommentInfo) => [c.id, c.author, c.initials, c.text, c.date, c.dateUtc, c.resolved];
  return JSON.stringify([part(t.comment), t.replies.map(part), open, replying, actions]);
}

function textButton(label: string, run: () => void, className?: string): HTMLButtonElement {
  const b = el('button', className, label);
  b.type = 'button';
  b.addEventListener('click', (e) => {
    e.stopPropagation();
    run();
  });
  return b;
}

/** Controls the comment column of a viewer (`viewer.comments`). */
export class CommentsController {
  private shown: boolean;
  private pref: CommentsMode;
  private auto: 'cards' | 'bubbles' | undefined;
  private eff: 'cards' | 'bubbles' = 'cards';
  private announced: 'cards' | 'bubbles' | undefined;
  private showRes: boolean;
  private threadList: CommentThread[] = [];
  private openId: string | undefined;
  private pendingOpen: string | undefined;
  private draft: ({ range: DocRange } & CommentSpot) | undefined;
  private docSeen: unknown;
  private readonly replying = new Set<string>();
  private colWidth = 0;
  private builtWidth = 0;
  private layer: HTMLDivElement | undefined;
  private readonly cards = new Map<string, { el: HTMLElement; sig: string }>();
  private readonly colors = new Map<string, string>();
  private readonly cleanup: Array<() => void> = [];

  constructor(private readonly host: CommentsHost, options: CommentsControllerOptions = {}) {
    this.shown = options.visible ?? true;
    this.pref = options.mode ?? 'auto';
    this.showRes = options.showResolved ?? false;
    this.cleanup.push(
      host.on('layout', () => {
        this.refresh();
        // A new comment that still has no place after a layout is not opened later.
        this.pendingOpen = undefined;
      }),
    );
    if (typeof document !== 'undefined') {
      document.addEventListener('click', this.onDocClick, true);
      document.addEventListener('keydown', this.onKey, true);
      this.cleanup.push(() => {
        document.removeEventListener('click', this.onDocClick, true);
        document.removeEventListener('keydown', this.onKey, true);
      });
    }
    if (typeof ResizeObserver !== 'undefined') {
      let last = host.root.clientWidth;
      const ro = new ResizeObserver(() => {
        if (host.root.clientWidth === last) return;
        last = host.root.clientWidth;
        this.refresh();
      });
      ro.observe(host.root);
      this.cleanup.push(() => ro.disconnect());
    }
  }

  get visible(): boolean {
    return this.shown;
  }

  show(): void {
    this.setVisible(true);
  }

  hide(): void {
    this.setVisible(false);
  }

  private setVisible(on: boolean): void {
    if (on === this.shown) return;
    this.shown = on;
    this.refresh();
    this.host.emit('chromechange', { part: 'comments', visible: on, collapsed: this.eff === 'bubbles' });
  }

  get mode(): CommentsMode {
    return this.pref;
  }

  set mode(value: CommentsMode) {
    if (value === this.pref) return;
    this.pref = value;
    this.refresh();
  }

  get effectiveMode(): 'cards' | 'bubbles' {
    return this.eff;
  }

  get showResolved(): boolean {
    return this.showRes;
  }

  set showResolved(value: boolean) {
    if (value === this.showRes) return;
    this.showRes = value;
    this.refresh();
  }

  /** The space the column takes beside the pages now, in CSS px (0 when hidden or without comments). */
  get width(): number {
    return this.colWidth;
  }

  /** The column's element, when there is one. */
  get element(): HTMLElement | undefined {
    return this.layer;
  }

  /** The id of the thread that is open. */
  get openThread(): string | undefined {
    return this.openId;
  }

  /** @internal The viewer rebuilt its rows with the current width. */
  rowsBuilt(): void {
    this.builtWidth = this.colWidth;
  }

  /** Open a thread (a reply's id opens its thread): line it up with its anchor and draw the connector. */
  open(id: string): void {
    if (!this.shown) return;
    this.threadList = this.host.threads();
    if (!this.openThreadById(id)) return;
    this.refresh();
  }

  /** Open without redrawing. False when there is no such thread or it is already open. */
  private openThreadById(id: string): boolean {
    const t = this.find(id);
    if (!t || this.openId === t.comment.id) return false;
    if (this.openId) this.replying.delete(this.openId);
    this.openId = t.comment.id;
    this.host.setActive([t.comment.id, ...t.replies.map((r) => r.id)]);
    this.host.emit('commentclick', { comment: t.comment });
    return true;
  }

  /** Open a thread and start its reply box (no-op where the viewer cannot reply). */
  reply(id: string): void {
    if (!this.host.editable() || !this.host.can('reply')) return;
    if (!this.shown) this.show();
    this.threadList = this.host.threads();
    const t = this.find(id);
    if (!t) return;
    this.openThreadById(id);
    this.replying.add(t.comment.id);
    this.refresh();
  }

  close(): void {
    if (!this.openId) return;
    this.replying.delete(this.openId);
    this.openId = undefined;
    this.host.setActive([]);
    this.refresh();
  }

  /** Start a comment on the current selection: a draft card with a text box beside it. */
  compose(): void {
    if (!this.host.editable() || !this.host.can('comment')) return;
    const selection = this.host.selection();
    if (!selection) return;
    if (!this.shown) this.show();
    this.draft = selection;
    this.refresh();
  }

  /** Recompute what is shown from the host and redraw. Called on layout, resize and state changes. */
  refresh(): void {
    this.sync();
    if (this.pendingOpen && this.openThreadById(this.pendingOpen)) this.pendingOpen = undefined;
    if (this.colWidth !== this.builtWidth) {
      this.host.rebuild();
      this.builtWidth = this.colWidth;
    }
    this.place();
  }

  /** Update the threads, the mode and the width, without touching the DOM. */
  sync(): void {
    const doc = this.host.currentDocument?.();
    if (doc !== this.docSeen) {
      // The draft's range, a reply being typed and a comment waiting to open belong to the old document.
      this.docSeen = doc;
      this.draft = undefined;
      this.replying.clear();
      this.pendingOpen = undefined;
    }
    this.threadList = this.host.threads();
    if (this.openId && !this.find(this.openId)) {
      this.openId = undefined;
      this.host.setActive([]);
    }
    this.eff = this.computeMode();
    const modeChanged = this.announced !== undefined && this.announced !== this.eff;
    this.announced = this.eff;
    this.colWidth = this.shown && this.has() ? (this.eff === 'cards' ? COLUMN : STRIP) + GAP : 0;
    if (modeChanged) this.host.emit('chromechange', { part: 'comments', visible: this.shown, collapsed: this.eff === 'bubbles' });
  }

  destroy(): void {
    for (const fn of this.cleanup) fn();
    this.layer?.remove();
    this.layer = undefined;
  }

  // ---- state ---------------------------------------------------------------------------------

  private has(): boolean {
    return this.threadList.length > 0 || !!this.draft;
  }

  private computeMode(): 'cards' | 'bubbles' {
    if (this.pref !== 'auto') return this.pref;
    const width = this.host.root.clientWidth;
    // Without a measured width (not laid out yet) keep what was chosen.
    if (!width) return this.auto ?? 'cards';
    return (this.auto = commentMode(width, this.host.widestPage(), this.auto));
  }

  private find(id: string): CommentThread | undefined {
    return this.threadList.find((t) => t.comment.id === id || t.replies.some((r) => r.id === id));
  }

  private resolvedCount(): number {
    return this.threadList.filter((t) => t.comment.resolved).length;
  }

  /** Threads shown now: resolved ones only when switched on (or open), and only those with a place. */
  private visibleThreads(): Array<{ thread: CommentThread; at: CommentSpot }> {
    const out: Array<{ thread: CommentThread; at: CommentSpot }> = [];
    for (const thread of this.threadList) {
      if (thread.comment.resolved && !this.showRes && this.openId !== thread.comment.id) continue;
      const at = this.host.anchorOf(thread);
      if (at) out.push({ thread, at });
    }
    return out;
  }

  private actionsFor(t: CommentThread): { reply: boolean; resolve: boolean; del: boolean } {
    if (!this.host.editable()) return { reply: false, resolve: false, del: false };
    return {
      reply: this.host.can('reply'),
      resolve: this.host.can('resolveComment'),
      del: this.host.can('deleteComment', { commentAuthor: t.comment.author }),
    };
  }

  // ---- events --------------------------------------------------------------------------------

  private onDocClick = (e: MouseEvent): void => {
    const target = e.target as Node | null;
    if (!target || this.layer?.contains(target)) return;
    if (this.host.root.contains(target)) {
      const hit = this.highlightAt(e.clientX, e.clientY);
      const thread = hit ? this.find(hit) : undefined;
      if (thread) {
        this.open(thread.comment.id);
        return;
      }
    }
    if (this.openId) this.close();
  };

  /** Esc closes the open thread from the page, the editing input or the column (capture: the editor handles Esc itself). */
  private onKey = (e: KeyboardEvent): void => {
    if (e.key !== 'Escape' || !this.openId) return;
    const target = e.target as Node | null;
    if (!target || !this.host.root.contains(target)) return;
    // A text box in the column cancels itself.
    if ((target as HTMLElement).tagName === 'TEXTAREA' && this.layer?.contains(target)) return;
    this.close();
  };

  private highlightAt(x: number, y: number): string | undefined {
    for (const hl of this.host.root.querySelectorAll<HTMLElement>('.docxkeun-comment-hl')) {
      const r = hl.getBoundingClientRect();
      if (r.width > 0 && x >= r.left && x <= r.right && y >= r.top && y <= r.bottom) return hl.dataset.comment;
    }
    return undefined;
  }

  // ---- drawing -------------------------------------------------------------------------------

  private place(): void {
    if (!this.shown || !this.has()) {
      this.layer?.remove();
      this.layer = undefined;
      this.cards.clear();
      return;
    }
    ensureStyle();
    const mode = this.eff;
    if (!this.layer || this.layer.dataset.mode !== mode) {
      this.layer?.remove();
      this.cards.clear();
      const layer = el('div', 'docxkeun-comments docxkeun-ui-root');
      layer.dataset.mode = mode;
      layer.setAttribute('role', 'complementary');
      layer.setAttribute('aria-label', 'Comments');
      this.host.root.appendChild(layer);
      this.layer = layer;
    }
    this.layer.style.left = `${this.host.columnX()}px`;
    this.layer.style.width = `${mode === 'cards' ? COLUMN : STRIP}px`;
    if (mode === 'cards') this.placeCards();
    else this.placeBubbles();
  }

  private placeCards(): void {
    const layer = this.layer!;
    for (const stale of layer.querySelectorAll('.dk-bubble,.dk-pop')) stale.remove();
    const visible = this.visibleThreads();
    const wanted = new Set<string>();
    const boxes: CardBox[] = [];
    const floor = this.resolvedCount() > 0 ? SWITCH_FLOOR : 0;
    const order: HTMLElement[] = [];
    const fresh = new Set<HTMLElement>();
    const connectors = new Map<string, CommentSpot>();

    const cardFor = (id: string, create: (card: HTMLElement) => void) => {
      wanted.add(id);
      let rec = this.cards.get(id);
      if (!rec) {
        const card = el('div', 'dk-card');
        card.dataset.comment = id;
        create(card);
        rec = { el: card, sig: '' };
        this.cards.set(id, rec);
        layer.appendChild(card);
        fresh.add(card);
      }
      order.push(rec.el);
      return rec;
    };

    for (const { thread, at } of visible) {
      const id = thread.comment.id;
      const open = this.openId === id;
      const actions = this.actionsFor(thread);
      const sig = signature(thread, open, this.replying.has(id), actions);
      const rec = cardFor(id, (c) => this.wireCard(c, id));
      if (rec.sig !== sig) {
        this.fillCard(rec.el, thread, open, actions);
        rec.sig = sig;
      }
      if (open) connectors.set(id, at);
      boxes.push({ id, anchorY: Math.max(at.top, floor), height: 0 });
    }
    if (this.draft) {
      const rec = cardFor(DRAFT, (c) => c.classList.add('dk-draft', 'open'));
      if (!rec.sig) {
        this.fillDraft(rec.el);
        rec.sig = 'draft';
      }
      connectors.set(DRAFT, this.draft);
      boxes.push({ id: DRAFT, anchorY: Math.max(this.draft.top, floor), height: 0 });
    }

    for (const [id, rec] of this.cards) {
      if (wanted.has(id)) continue;
      rec.el.remove();
      this.cards.delete(id);
    }
    this.fillSwitch();
    // Tab order follows the cards' order on the page; moving nodes only when it differs keeps focus.
    const present = [...layer.querySelectorAll<HTMLElement>('.dk-card')];
    if (present.some((c, i) => c !== order[i])) for (const c of order) layer.appendChild(c);

    for (const box of boxes) box.height = this.cards.get(box.id)!.el.offsetHeight;
    const tops = layoutCards(boxes, this.draft ? DRAFT : this.openId, CARD_GAP);
    for (const box of boxes) {
      const rec = this.cards.get(box.id)!;
      const top = tops.get(box.id) ?? box.anchorY;
      // A card that was just created appears in place instead of sliding in from the top.
      if (fresh.has(rec.el)) rec.el.style.transition = 'none';
      rec.el.style.top = `${top}px`;
      if (fresh.has(rec.el)) {
        void rec.el.offsetHeight;
        rec.el.style.transition = '';
      }
    }
    for (const [id, at] of connectors) this.fillConnector(this.cards.get(id)!.el, at);
  }

  private fillSwitch(): void {
    const layer = this.layer!;
    let sw = layer.querySelector<HTMLButtonElement>('.dk-resolved-switch');
    const n = this.resolvedCount();
    if (!n) {
      sw?.remove();
      return;
    }
    if (!sw) {
      sw = textButton('', () => (this.showResolved = !this.showRes), 'dk-resolved-switch');
      sw.setAttribute('role', 'switch');
      layer.appendChild(sw);
    }
    sw.textContent = `Show resolved (${n})`;
    sw.setAttribute('aria-checked', String(this.showRes));
  }

  private placeBubbles(): void {
    const layer = this.layer!;
    const visible = this.visibleThreads();
    const groups = bubbleGroups(
      visible.map(({ thread, at }) => ({ id: thread.comment.id, y: at.top })),
      BUBBLE_SIZE + 2,
    );
    const mark = (ids: string[]) => ids.includes(this.openId ?? '\0');
    // Bubbles and the popover stay across refreshes, so a reply being typed keeps its text and focus.
    const existing = new Map([...layer.querySelectorAll<HTMLButtonElement>('.dk-bubble')].map((b) => [b.dataset.key!, b]));
    for (const group of groups) {
      const key = group.ids.join(',');
      let b = existing.get(key);
      if (b) existing.delete(key);
      else {
        b = textButton('', () => this.open(group.ids[0]!), 'dk-bubble');
        b.dataset.key = key;
        if (group.ids.length > 1) b.textContent = String(group.ids.length);
        else b.innerHTML = icon('comment');
        b.setAttribute('aria-label', group.ids.length > 1 ? `${group.ids.length} comments` : 'Comment');
        layer.appendChild(b);
      }
      b.style.top = `${group.y}px`;
      b.classList.toggle('open', mark(group.ids));
    }
    for (const stale of existing.values()) stale.remove();

    const pop = layer.querySelector<HTMLElement>('.dk-pop');
    const drop = () => pop?.remove();
    const group = groups.find((g) => mark(g.ids));
    const thread = this.openId ? this.find(this.openId) : undefined;
    if (this.draft) {
      if (pop?.dataset.sig === 'draft') {
        this.placePopover(pop, this.draft);
        return;
      }
      drop();
      const next = el('div', 'dk-pop docxkeun-ui-root');
      next.dataset.sig = 'draft';
      const card = el('div', 'dk-card dk-draft open');
      card.dataset.comment = DRAFT;
      this.fillDraft(card);
      next.appendChild(card);
      layer.appendChild(next);
      this.placePopover(next, this.draft);
      return;
    }
    if (!group || !thread) return drop();
    const at = visible.find((v) => v.thread.comment.id === thread.comment.id)?.at ?? { top: group.y, pageRight: 0, height: 0 };
    const actions = this.actionsFor(thread);
    const sig = signature(thread, true, this.replying.has(thread.comment.id), actions) + group.ids.join(',');
    if (pop?.dataset.sig === sig) {
      this.placePopover(pop, at);
      return;
    }
    drop();
    const next = el('div', 'dk-pop docxkeun-ui-root');
    next.dataset.sig = sig;
    if (group.ids.length > 1) {
      const at = group.ids.indexOf(thread.comment.id);
      const nav = el('div', 'dk-pop-nav');
      const prev = textButton('', () => this.step(group.ids, -1), 'dk-prev');
      prev.innerHTML = icon('back');
      prev.setAttribute('aria-label', 'Previous comment');
      prev.disabled = at <= 0;
      const forward = textButton('', () => this.step(group.ids, 1), 'dk-next');
      forward.innerHTML = icon('forward');
      forward.setAttribute('aria-label', 'Next comment');
      forward.disabled = at >= group.ids.length - 1;
      nav.append(el('span', 'dk-count', `${at + 1} of ${group.ids.length}`), prev, forward);
      next.appendChild(nav);
    }
    const card = el('div', 'dk-card');
    card.dataset.comment = thread.comment.id;
    this.fillCard(card, thread, true, actions);
    card.querySelector('.dk-connector')?.remove();
    next.appendChild(card);
    layer.appendChild(next);
    this.placePopover(next, at);
  }

  /**
   * Put a bubble popover below the last line of its text (so it never covers it), its right edge
   * on the bubble strip's, inside the visible part of the viewer; above the text when there is
   * no room below.
   */
  private placePopover(pop: HTMLElement, at: CommentSpot): void {
    const root = this.host.root;
    const x = parseFloat(this.layer!.style.left) || 0;
    let left = x + STRIP - COLUMN;
    if (root.clientWidth > 0) left = Math.max(root.scrollLeft + EDGE, Math.min(left, root.scrollLeft + root.clientWidth - COLUMN - EDGE));
    pop.style.left = `${left - x}px`;
    const below = (at.bottom ?? at.top + at.height) + POP_GAP;
    let top = below;
    const h = pop.offsetHeight;
    const view = root.clientHeight;
    if (view > 0 && h > 0) {
      const viewTop = root.scrollTop + EDGE;
      const viewBottom = root.scrollTop + view - EDGE;
      const above = at.top - POP_GAP - h;
      if (below + h > viewBottom) top = above >= viewTop ? above : Math.max(viewTop, viewBottom - h);
    }
    pop.style.top = `${top}px`;
  }

  private step(ids: string[], by: number): void {
    const at = ids.indexOf(this.openId ?? '');
    const id = ids[at + by];
    if (id) this.open(id);
  }

  private wireCard(card: HTMLElement, id: string): void {
    card.tabIndex = 0;
    card.setAttribute('role', 'group');
    card.addEventListener('click', () => {
      if (this.openId !== id) this.open(id);
    });
    card.addEventListener('keydown', (e) => {
      if (e.target !== card || (e.key !== 'Enter' && e.key !== ' ')) return;
      e.preventDefault();
      if (this.openId !== id) this.open(id);
    });
  }

  private fillConnector(card: HTMLElement, at: CommentSpot): void {
    const line = card.querySelector<HTMLElement>('.dk-connector');
    if (!line) return;
    const cardTop = parseFloat(card.style.top) || at.top;
    const x = parseFloat(this.layer!.style.left) || 0;
    line.style.width = `${Math.max(0, x - at.pageRight)}px`;
    line.style.top = `${Math.max(8, Math.min(Math.max(8, card.offsetHeight - 8), at.top + at.height / 2 - cardTop))}px`;
  }

  private colorOf(author: string): string {
    let color = this.colors.get(author);
    if (!color) {
      color = `#${REVISION_COLORS[this.colors.size % REVISION_COLORS.length]!}`;
      this.colors.set(author, color);
    }
    return color;
  }

  private head(c: CommentInfo): HTMLElement {
    const head = el('div', 'dk-head');
    const avatar = el('span', 'dk-avatar', initialsOf(c));
    avatar.style.background = this.colorOf(c.author);
    avatar.setAttribute('aria-hidden', 'true');
    const date = c.dateUtc ?? c.date;
    const when = el('span', 'dk-when', relativeTime(date));
    if (date && !Number.isNaN(Date.parse(date))) when.title = new Date(date).toLocaleString();
    head.append(avatar, el('span', 'dk-who', c.author || 'Unknown'), when);
    return head;
  }

  private fillCard(card: HTMLElement, t: CommentThread, open: boolean, actions: { reply: boolean; resolve: boolean; del: boolean }): void {
    const c = t.comment;
    card.replaceChildren();
    card.classList.toggle('open', open);
    card.classList.toggle('resolved', c.resolved);
    card.setAttribute('aria-expanded', String(open));
    card.setAttribute('aria-label', `Comment by ${c.author || 'Unknown'}`);
    if (!open) {
      card.append(this.head(c), el('div', 'dk-text dk-clamp', c.text));
      if (t.replies.length) card.appendChild(el('div', 'dk-replies', `${t.replies.length} ${t.replies.length === 1 ? 'reply' : 'replies'}`));
      return;
    }
    card.append(el('div', 'dk-connector'), this.head(c), el('div', 'dk-text', c.text));
    for (const r of t.replies) {
      const item = el('div', 'dk-reply');
      item.append(this.head(r), el('div', 'dk-text', r.text));
      card.appendChild(item);
    }
    if (this.replying.has(c.id)) {
      card.appendChild(this.composer('Reply', 'Reply…', (text) => this.postReply(c.id, text), () => {
        this.replying.delete(c.id);
        this.refresh();
      }));
    } else if (actions.reply || actions.resolve || actions.del) {
      const row = el('div', 'dk-actions');
      if (actions.reply) {
        row.appendChild(textButton('Reply', () => {
          this.replying.add(c.id);
          this.refresh();
        }));
      }
      if (actions.resolve) row.appendChild(textButton(c.resolved ? 'Reopen' : 'Resolve', () => void this.resolve(t)));
      if (actions.del) row.appendChild(textButton('Delete', () => void this.remove(t)));
      card.appendChild(row);
    }
  }

  private fillDraft(card: HTMLElement): void {
    const me: CommentInfo = { id: DRAFT, author: this.host.author(), text: '', blocks: [], resolved: false, el: {} as never };
    card.replaceChildren(el('div', 'dk-connector'), this.head(me), this.composer('Comment', 'Add a comment…', (text) => this.postDraft(text), () => {
      this.draft = undefined;
      this.refresh();
    }));
    card.querySelector('.dk-when')?.remove();
    card.setAttribute('aria-label', 'New comment');
  }

  /** A text box with Post and Cancel. `post` rejects to keep the box (the viewer reports the error). */
  private composer(label: string, placeholder: string, post: (text: string) => Promise<void>, cancel: () => void): HTMLElement {
    const wrap = el('div', 'dk-composer');
    const box = el('textarea');
    box.placeholder = placeholder;
    box.setAttribute('aria-label', label);
    queueMicrotask(() => box.isConnected && box.focus());
    const send = textButton('Post', () => {
      const text = box.value.trim();
      if (!text) return;
      send.disabled = true;
      box.disabled = true;
      post(text).catch(() => {
        send.disabled = !box.value.trim();
        box.disabled = false;
        box.focus();
      });
    }, 'dk-post');
    send.disabled = true;
    box.addEventListener('input', () => (send.disabled = !box.value.trim()));
    box.addEventListener('keydown', (e) => {
      if (e.key === 'Escape') {
        e.stopPropagation();
        cancel();
      } else if (e.key === 'Enter' && (e.metaKey || e.ctrlKey)) {
        e.preventDefault();
        if (!send.disabled) send.click();
      }
    });
    const buttons = el('div', 'dk-composer-buttons');
    buttons.append(send, textButton('Cancel', cancel));
    wrap.append(box, buttons);
    return wrap;
  }

  // ---- edits ---------------------------------------------------------------------------------

  private async postReply(id: string, text: string): Promise<void> {
    const author = this.host.author();
    await this.host.transaction('Reply', (editor) => editor.replyToComment(id, { author, text }));
    this.replying.delete(id);
    this.refresh();
  }

  private async postDraft(text: string): Promise<void> {
    const draft = this.draft;
    if (!draft) return;
    const author = this.host.author();
    const id = await this.host.transaction('Add comment', (editor) => editor.addComment(draft.range, { author, text }));
    this.draft = undefined;
    this.pendingOpen = id;
    this.refresh();
  }

  private async resolve(t: CommentThread): Promise<void> {
    const resolved = !t.comment.resolved;
    try {
      await this.host.transaction(resolved ? 'Resolve comment' : 'Reopen comment', (editor) => editor.resolveComment(t.comment.id, resolved));
    } catch {
      return;
    }
    if (resolved) this.close();
  }

  private async remove(t: CommentThread): Promise<void> {
    try {
      await this.host.transaction('Delete comment', (editor) => editor.deleteComment(t.comment.id));
    } catch {
      return;
    }
    this.close();
  }
}
