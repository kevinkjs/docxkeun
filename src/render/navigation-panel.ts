/**
 * The navigation panel beside the pages: an outline of the headings, the bookmarks and the links,
 * with search, previous/next, back/forward, and a fold into a thin rail. `viewer.navigation` is its
 * controller; the DOM is built on the first `show()`. The controller also keeps the lists
 * (`headings`, `bookmarks`, `links`) that `viewer.navigation` has always offered.
 */
import type { DocxViewer } from './viewer.js';
import type { DocLocation, NavKind, NavTarget, NavTargets } from './navigation.js';
import { icon, type IconName } from './ui/icons.js';
import { currentHeading, filterTargets } from './ui/nav-model.js';
import { ensureUiStyle } from './ui/theme.js';

export type NavTab = 'outline' | 'bookmarks' | 'links';

export interface NavigationOptions {
  /** The tab shown first. Default: outline. */
  tab?: NavTab;
  /** Start folded into the rail. Default: folded when the frame is narrower than 720px. */
  collapsed?: boolean;
}

/** What the panel needs from the viewer. */
export type NavigationHost = Pick<DocxViewer, 'root' | 'on' | 'emit' | 'frame' | 'goTo' | 'back' | 'forward' | 'canGoBack' | 'canGoForward' | 'location' | 'next' | 'previous' | 'layout' | 'scrollToComment' | 'collectedTargets'>;

const STYLE_ID = 'docxkeun-navigation-style';
const NARROW = 720;
const CSS = `
.docxkeun-nav{flex:none;box-sizing:border-box;width:240px;overflow:hidden;display:flex;flex-direction:column;background:var(--dk-ui-bg);color:var(--dk-ui-fg);border-right:1px solid var(--dk-ui-line);font:var(--dk-ui-font)}
.docxkeun-nav[hidden]{display:none}
.docxkeun-nav.dk-rail{width:36px;align-items:center;padding:4px 0;gap:2px}
@media (prefers-reduced-motion:no-preference){.docxkeun-nav{transition:width .15s ease}}
.docxkeun-nav button{height:28px;min-width:28px;border:1px solid transparent;border-radius:4px;background:transparent;color:inherit;font:inherit;padding:0 6px;cursor:pointer;display:inline-flex;align-items:center;justify-content:center}
.docxkeun-nav button:hover:not(:disabled){background:var(--dk-ui-hover)}
.docxkeun-nav button:focus-visible,.docxkeun-nav input:focus-visible{outline:2px solid var(--dk-ui-accent);outline-offset:1px}
.docxkeun-nav button:disabled{opacity:.4;cursor:default}
.docxkeun-nav .dk-nav-bar{display:flex;align-items:center;gap:2px;padding:4px 6px}
.docxkeun-nav .dk-nav-bar .dk-nav-spacer{flex:1 1 auto}
.docxkeun-nav .dk-nav-tabs{display:flex;gap:2px;padding:0 6px;border-bottom:1px solid var(--dk-ui-line)}
.docxkeun-nav .dk-nav-tabs button{border-radius:4px 4px 0 0;height:30px;flex:1 1 0;min-width:0;padding:0 4px;font-size:12px;white-space:nowrap;color:var(--dk-ui-muted);border-bottom:2px solid transparent}
.docxkeun-nav .dk-nav-tabs button[aria-selected=true]{color:var(--dk-ui-fg);border-bottom-color:var(--dk-ui-accent)}
.docxkeun-nav input[type=search]{flex:1 1 auto;min-width:0;height:28px;box-sizing:border-box;border:1px solid var(--dk-ui-line);border-radius:4px;background:var(--dk-ui-bg);color:inherit;font:inherit;padding:0 8px}
.docxkeun-nav .dk-nav-hidden[hidden],.docxkeun-nav .dk-nav-empty[hidden]{display:none}
.docxkeun-nav .dk-nav-hidden{display:flex;align-items:center;gap:6px;padding:2px 10px;font-size:12px;color:var(--dk-ui-muted)}
.docxkeun-nav .dk-nav-count{font-size:12px;color:var(--dk-ui-muted);padding:0 4px;white-space:nowrap}
.docxkeun-nav .dk-nav-list{list-style:none;margin:0;padding:2px 4px;overflow:auto;flex:1 1 auto;min-height:0}
.docxkeun-nav .dk-nav-list button{width:100%;height:auto;min-height:26px;justify-content:flex-start;text-align:left;gap:8px;padding-top:3px;padding-bottom:3px}
.docxkeun-nav .dk-nav-list [aria-current=location]{background:var(--dk-ui-accent-soft)}
.docxkeun-nav .dk-nav-label{flex:1 1 auto;min-width:0;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
.docxkeun-nav .dk-nav-meta{flex:none;font-size:11px;color:var(--dk-ui-muted);max-width:90px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
.docxkeun-nav .dk-nav-empty{padding:16px 12px;color:var(--dk-ui-muted)}
.docxkeun-nav .dk-nav-foot{display:flex;padding:4px 6px;border-top:1px solid var(--dk-ui-line)}
`;

function ensureStyle(): void {
  ensureUiStyle();
  if (typeof document === 'undefined' || document.getElementById(STYLE_ID)) return;
  const style = document.createElement('style');
  style.id = STYLE_ID;
  style.textContent = CSS;
  document.head.appendChild(style);
}

const TABS: Array<{ tab: NavTab; label: string; icon: IconName; kind: NavKind; key: keyof NavTargets; empty: string }> = [
  { tab: 'outline', label: 'Outline', icon: 'outline', kind: 'heading', key: 'headings', empty: 'No headings' },
  { tab: 'bookmarks', label: 'Bookmarks', icon: 'bookmark', kind: 'bookmark', key: 'bookmarks', empty: 'No bookmarks' },
  { tab: 'links', label: 'Links', icon: 'link', kind: 'link', key: 'links', empty: 'No links' },
];

function button(label: string, html: string, run: () => void, extra = ''): HTMLButtonElement {
  const b = document.createElement('button');
  b.type = 'button';
  b.innerHTML = html;
  b.title = label;
  b.setAttribute('aria-label', label);
  if (extra) b.className = extra;
  b.addEventListener('click', run);
  return b;
}

/** Controls the navigation panel of a viewer (`viewer.navigation`); the DOM is built on the first `show()`. */
export class NavigationController {
  private el: HTMLElement | undefined;
  private shown = false;
  private fold: boolean | undefined;
  private active: NavTab;
  private query = '';
  private showHidden = false;
  private current: NavTarget | undefined;
  private readonly items = new Map<NavTarget, HTMLElement>();
  private readonly cleanup: Array<() => void> = [];
  private scrollTimer: ReturnType<typeof setTimeout> | undefined;
  private list: HTMLUListElement | undefined;
  /** The row last marked current (revealed once, not on every scroll). */
  private marked: HTMLElement | undefined;
  private count: HTMLElement | undefined;
  private back: HTMLButtonElement | undefined;
  private forward: HTMLButtonElement | undefined;
  private tabButtons: HTMLButtonElement[] = [];
  private hiddenLabel: HTMLElement | undefined;
  private emptyEl: HTMLElement | undefined;

  constructor(private readonly viewer: NavigationHost, options: NavigationOptions = {}) {
    this.fold = options.collapsed;
    this.active = options.tab ?? 'outline';
  }

  /** The headings of the document, in order. */
  get headings(): NavTarget[] {
    return this.viewer.collectedTargets().headings;
  }

  /** The bookmarks (including hidden `_Toc…` ones), in order. */
  get bookmarks(): NavTarget[] {
    return this.viewer.collectedTargets().bookmarks;
  }

  /** The hyperlinks in the body, in order. */
  get links(): NavTarget[] {
    return this.viewer.collectedTargets().links;
  }

  get element(): HTMLElement | undefined {
    return this.el;
  }

  get visible(): boolean {
    return this.shown;
  }

  get collapsed(): boolean {
    return !!this.fold;
  }

  set collapsed(value: boolean) {
    if (value === !!this.fold && this.fold !== undefined) return;
    this.fold = value;
    if (!this.el) return;
    this.build();
    this.announce();
  }

  get tab(): NavTab {
    return this.active;
  }

  set tab(value: NavTab) {
    if (value === this.active) return;
    this.active = value;
    this.query = '';
    if (!this.el) return;
    if (this.fold) return;
    this.build();
  }

  /** Show the panel. With a tab, unfold onto that tab; without one, keep the fold as it is. */
  show(tab?: NavTab): void {
    let changed = !this.shown;
    if (tab) {
      this.active = tab;
      this.query = '';
      if (this.fold) changed = true;
      this.fold = false;
    }
    if (!this.el) this.mount();
    else if (tab) this.build();
    this.shown = true;
    this.el!.hidden = false;
    if (changed) this.announce();
  }

  hide(): void {
    if (!this.shown) return;
    this.shown = false;
    if (this.el) this.el.hidden = true;
    this.announce();
  }

  /** Rebuild the lists from the current layout (done automatically after each layout). */
  refresh(): void {
    if (!this.el || this.fold) {
      if (this.el) this.build();
      return;
    }
    this.renderList();
  }

  destroy(): void {
    clearTimeout(this.scrollTimer);
    for (const fn of this.cleanup.splice(0)) fn();
    this.el?.remove();
    this.el = undefined;
  }

  private announce(): void {
    this.viewer.emit('chromechange', { part: 'navigation', visible: this.shown, collapsed: !!this.fold });
  }

  private mount(): void {
    ensureStyle();
    const { frame, body } = this.viewer.frame();
    if (this.fold === undefined) {
      const width = frame.getBoundingClientRect().width || frame.clientWidth;
      // A width of 0 means the frame is not laid out yet; leave the panel open then.
      this.fold = width > 0 && width < NARROW;
    }
    const el = (this.el = document.createElement('nav'));
    el.className = 'docxkeun-nav docxkeun-ui-root';
    el.setAttribute('aria-label', 'Navigation');
    el.hidden = true;
    body.insertBefore(el, this.viewer.root);
    this.build();

    this.cleanup.push(this.viewer.on('layout', () => this.refresh()));
    this.cleanup.push(
      this.viewer.on('navigate', ({ target }) => {
        this.current = target?.kind === 'heading' ? target : undefined;
        this.syncHistory();
        this.markCurrent();
      }),
    );
    const onScroll = () => {
      if (this.scrollTimer) return;
      this.scrollTimer = setTimeout(() => {
        this.scrollTimer = undefined;
        this.current = undefined;
        this.markCurrent();
      }, 100);
    };
    this.viewer.root.addEventListener('scroll', onScroll, { passive: true });
    this.cleanup.push(() => this.viewer.root.removeEventListener('scroll', onScroll));
  }

  private build(): void {
    const el = this.el!;
    // Rebuilding drops the focused control; when focus was inside the panel, put it on the counterpart.
    const hadFocus = el.contains(document.activeElement);
    this.items.clear();
    this.list = this.count = this.back = this.forward = undefined;
    this.tabButtons = [];
    el.classList.toggle('dk-rail', !!this.fold);
    if (this.fold) {
      el.replaceChildren(...this.rail());
      if (hadFocus) el.querySelector<HTMLButtonElement>('[data-unfold]')?.focus();
      return;
    }
    el.replaceChildren(...this.panel());
    this.renderList();
    if (hadFocus) this.tabButtons.find((b) => b.dataset.tab === this.active)?.focus();
  }

  private rail(): HTMLElement[] {
    const out: HTMLElement[] = [];
    for (const t of TABS) out.push(button(t.label, icon(t.icon), () => this.unfoldTo(t.tab)));
    if ((this.viewer.layout?.commentAnchors.size ?? 0) > 0) {
      out.push(button('Comments', icon('comment'), () => this.goToNextComment()));
    }
    const spacer = document.createElement('span');
    spacer.style.flex = '1 1 auto';
    const unfold = button('Show navigation', icon('forward'), () => (this.collapsed = false));
    unfold.dataset.unfold = '';
    out.push(spacer, unfold);
    return out;
  }

  private panel(): HTMLElement[] {
    const bar = document.createElement('div');
    bar.className = 'dk-nav-bar';
    this.back = button('Back', icon('back'), () => this.viewer.back());
    this.forward = button('Forward', icon('forward'), () => this.viewer.forward());
    bar.append(this.back, this.forward);
    const spacer = document.createElement('span');
    spacer.className = 'dk-nav-spacer';
    bar.appendChild(spacer);

    const tabs = document.createElement('div');
    tabs.className = 'dk-nav-tabs';
    tabs.setAttribute('role', 'tablist');
    tabs.setAttribute('aria-label', 'Navigation views');
    for (const t of TABS) {
      const b = button(t.label, t.label, () => (this.tab = t.tab));
      b.setAttribute('role', 'tab');
      b.dataset.tab = t.tab;
      tabs.appendChild(b);
      this.tabButtons.push(b);
    }
    tabs.addEventListener('keydown', (e) => {
      const i = TABS.findIndex((t) => t.tab === this.active);
      const next = e.key === 'ArrowRight' ? i + 1 : e.key === 'ArrowLeft' ? i - 1 : undefined;
      if (next === undefined) return;
      e.preventDefault();
      this.tab = TABS[(next + TABS.length) % TABS.length]!.tab;
      this.tabButtons.find((b) => b.getAttribute('aria-selected') === 'true')?.focus();
    });

    const search = document.createElement('div');
    search.className = 'dk-nav-bar';
    const input = document.createElement('input');
    input.type = 'search';
    input.placeholder = 'Search';
    input.setAttribute('aria-label', 'Search the list');
    input.value = this.query;
    input.addEventListener('input', () => {
      this.query = input.value;
      this.renderList();
    });
    const prev = button('Previous', icon('fold'), () => this.step(-1));
    const next = button('Next', icon('unfold'), () => this.step(1));
    prev.dataset.step = 'prev';
    next.dataset.step = 'next';
    this.count = document.createElement('span');
    this.count.className = 'dk-nav-count';
    search.append(input, this.count, prev, next);

    const list = (this.list = document.createElement('ul'));
    list.className = 'dk-nav-list';

    const foot = document.createElement('div');
    foot.className = 'dk-nav-foot';
    foot.appendChild(button('Fold navigation', icon('back'), () => (this.collapsed = true)));

    const label = (this.hiddenLabel = document.createElement('label'));
    label.className = 'dk-nav-hidden';
    const sw = document.createElement('input');
    sw.type = 'checkbox';
    sw.setAttribute('role', 'switch');
    sw.checked = this.showHidden;
    sw.addEventListener('change', () => {
      this.showHidden = sw.checked;
      this.renderList();
    });
    label.append(sw, 'Show hidden bookmarks');
    const empty = (this.emptyEl = document.createElement('div'));
    empty.className = 'dk-nav-empty';
    empty.hidden = true;

    return [bar, tabs, search, label, empty, list, foot];
  }

  private get spec() {
    return TABS.find((t) => t.tab === this.active)!;
  }

  private targets(): NavTarget[] {
    const spec = this.spec;
    let list = this.viewer.collectedTargets()[spec.key];
    if (spec.kind === 'bookmark' && !this.showHidden) list = list.filter((b) => !b.hidden);
    return filterTargets(list, this.query);
  }

  private renderList(): void {
    const list = this.list;
    if (!list) return;
    const spec = this.spec;
    this.items.clear();
    for (const b of this.tabButtons) {
      const on = b.dataset.tab === this.active;
      b.setAttribute('aria-selected', String(on));
      b.tabIndex = on ? 0 : -1;
    }
    const prev = this.el!.querySelector<HTMLButtonElement>('[data-step=prev]');
    const next = this.el!.querySelector<HTMLButtonElement>('[data-step=next]');
    prev?.setAttribute('aria-label', `Previous ${spec.kind}`);
    next?.setAttribute('aria-label', `Next ${spec.kind}`);
    if (prev) prev.title = `Previous ${spec.kind}`;
    if (next) next.title = `Next ${spec.kind}`;
    this.syncHistory();

    // The "show hidden" switch belongs to the Bookmarks tab only.
    this.hiddenLabel!.hidden = spec.kind !== 'bookmark';

    const all = this.viewer.collectedTargets()[spec.key].length;
    const targets = this.targets();
    this.count!.textContent = String(targets.length);
    const empty = this.emptyEl!;
    empty.hidden = targets.length > 0;
    if (!targets.length) {
      empty.textContent = all && this.query.trim() ? 'No matches' : spec.empty;
      list.replaceChildren();
      return;
    }
    const rows = targets.map((t) => {
      const li = document.createElement('li');
      const b = document.createElement('button');
      b.type = 'button';
      b.title = t.link?.url ?? t.label;
      if (t.level) b.style.paddingLeft = `${6 + t.level * 14}px`;
      const text = document.createElement('span');
      text.className = 'dk-nav-label';
      text.textContent = t.label;
      const meta = document.createElement('span');
      meta.className = 'dk-nav-meta';
      meta.textContent = t.kind === 'link' ? (t.link?.anchor ? `#${t.link.anchor}` : (t.link?.url ?? '')) : `p. ${t.page + 1}`;
      b.append(text, meta);
      b.addEventListener('click', () => this.viewer.goTo(t, { history: true }));
      li.appendChild(b);
      this.items.set(t, b);
      return li;
    });
    list.replaceChildren(...rows);
    this.markCurrent();
  }

  /** Highlight the heading the viewport is at (outline tab only). */
  private markCurrent(): void {
    if (this.active !== 'outline' || !this.items.size) return;
    const at: DocLocation = this.viewer.location;
    // A jump leaves 24pt above its target, so a heading that close below the top is the current one.
    const now = this.current && this.items.has(this.current) ? this.current : currentHeading([...this.items.keys()], at, 30);
    for (const [t, b] of this.items) {
      if (t === now) {
        b.setAttribute('aria-current', 'location');
        if (b !== this.marked) this.reveal(b);
        this.marked = b;
      } else b.removeAttribute('aria-current');
    }
  }

  /** Scroll the list (never the page around it) so a row is in view. */
  private reveal(row: HTMLElement): void {
    const list = this.list;
    if (!list) return;
    const box = list.getBoundingClientRect();
    if (!box.height) return;
    const r = row.getBoundingClientRect();
    if (r.top < box.top) list.scrollTop += r.top - box.top;
    else if (r.bottom > box.bottom) list.scrollTop += Math.min(r.bottom - box.bottom, r.top - box.top);
  }

  private syncHistory(): void {
    if (this.back) this.back.disabled = !this.viewer.canGoBack;
    if (this.forward) this.forward.disabled = !this.viewer.canGoForward;
  }

  private step(dir: 1 | -1): void {
    const opts = { includeHidden: this.showHidden };
    if (dir > 0) this.viewer.next(this.spec.kind, opts);
    else this.viewer.previous(this.spec.kind, opts);
  }

  private unfoldTo(tab: NavTab): void {
    this.active = tab;
    this.query = '';
    this.collapsed = false;
  }

  /** The first comment below the top of the viewport, wrapping to the first one in the document. */
  private goToNextComment(): void {
    const anchors = [...(this.viewer.layout?.commentAnchors ?? [])].map(([id, a]) => ({ id, page: a.page, y: a.rect.y }));
    if (!anchors.length) return;
    anchors.sort((a, b) => a.page - b.page || a.y - b.y);
    const at = this.viewer.location;
    const next = anchors.find((a) => a.page > at.page || (a.page === at.page && a.y > at.y)) ?? anchors[0]!;
    this.viewer.scrollToComment(next.id);
  }
}
