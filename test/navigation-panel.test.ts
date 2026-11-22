// @vitest-environment happy-dom
import { afterEach, describe, expect, it, vi } from 'vitest';
import { NavigationController } from '../src/render/navigation-panel.js';
import type { NavTarget } from '../src/render/navigation.js';

const target = (kind: NavTarget['kind'], label: string, page: number, y: number, extra: Partial<NavTarget> = {}): NavTarget => ({ kind, label, page, x: 0, y, rects: [], ...extra });

function stub(opts: { width?: number; comments?: boolean; empty?: boolean } = {}) {
  const listeners = new Map<string, Array<(p: any) => void>>();
  const emitted: Array<{ event: string; payload: any }> = [];
  const root = document.createElement('div');
  const frame = document.createElement('div');
  const body = document.createElement('div');
  body.appendChild(root);
  frame.appendChild(body);
  document.body.appendChild(frame);
  if (opts.width !== undefined) Object.defineProperty(frame, 'clientWidth', { value: opts.width });
  const targets = opts.empty
    ? { headings: [], bookmarks: [], links: [] }
    : {
        headings: [target('heading', '1. Introduction', 0, 50, { level: 0 }), target('heading', '2. Payment terms', 0, 400, { level: 0 }), target('heading', 'Net 30', 1, 40, { level: 1 })],
        bookmarks: [target('bookmark', 'top', 0, 50, { name: 'top' }), target('bookmark', '_Toc1', 0, 60, { name: '_Toc1', hidden: true })],
        links: [target('link', 'the website', 0, 90, { link: { url: 'https://example.com/pay' } as never })],
      };
  const viewer: any = {
    root,
    layout: { commentAnchors: new Map(opts.comments ? [['c1', { page: 0, rect: { y: 10 } }], ['c2', { page: 1, rect: { y: 20 } }]] : []) },
    location: { page: 0, x: 0, y: 100 },
    canGoBack: false,
    canGoForward: false,
    goTo: vi.fn(),
    back: vi.fn(),
    forward: vi.fn(),
    next: vi.fn(),
    previous: vi.fn(),
    scrollToComment: vi.fn(() => true),
    collectedTargets: () => targets,
    frame: () => ({ frame, toolbarSlot: document.createElement('div'), body }),
    on: (event: string, fn: (p: any) => void) => {
      (listeners.get(event) ?? listeners.set(event, []).get(event)!).push(fn);
      return () => {};
    },
    emit: (event: string, payload: any) => {
      emitted.push({ event, payload });
      for (const l of listeners.get(event) ?? []) l(payload);
    },
  };
  return { viewer, body, root, frame, targets, emitted };
}

const visibleEmpty = (nav: NavigationController) => [...nav.element!.querySelectorAll('.dk-nav-empty')].find((e) => !(e as HTMLElement).hidden)?.textContent;
const labels = (nav: NavigationController) => [...nav.element!.querySelectorAll('.dk-nav-list button')].map((b) => b.querySelector('.dk-nav-label')!.textContent);

afterEach(() => {
  document.body.replaceChildren();
  vi.useRealTimers();
});

describe('NavigationController', () => {
  it('builds nothing until shown, then puts the panel before the viewer root', () => {
    const s = stub();
    const nav = new NavigationController(s.viewer);
    expect(nav.element).toBeUndefined();
    expect(nav.visible).toBe(false);
    nav.show();
    expect(nav.visible).toBe(true);
    expect(nav.element!.parentElement).toBe(s.body);
    expect(nav.element!.nextElementSibling).toBe(s.root);
    expect(s.emitted.at(-1)).toEqual({ event: 'chromechange', payload: { part: 'navigation', visible: true, collapsed: false } });
  });

  it('exposes the targets through the compat getters', () => {
    const s = stub();
    const nav = new NavigationController(s.viewer);
    expect(nav.headings).toBe(s.targets.headings);
    expect(nav.bookmarks).toBe(s.targets.bookmarks);
    expect(nav.links).toBe(s.targets.links);
  });

  it('renders a tablist with three tabs and the outline first', () => {
    const s = stub();
    const nav = new NavigationController(s.viewer);
    nav.show();
    const list = nav.element!.querySelector('[role=tablist]')!;
    const tabs = [...list.querySelectorAll('[role=tab]')];
    expect(tabs.map((t) => t.getAttribute('aria-label'))).toEqual(['Outline', 'Bookmarks', 'Links']);
    expect(tabs.map((t) => t.getAttribute('aria-selected'))).toEqual(['true', 'false', 'false']);
    expect(labels(nav)).toEqual(['1. Introduction', '2. Payment terms', 'Net 30']);
  });

  it('switches tabs and hides hidden bookmarks until the switch is on', () => {
    const s = stub();
    const nav = new NavigationController(s.viewer, { tab: 'bookmarks' });
    nav.show();
    expect(nav.tab).toBe('bookmarks');
    expect(labels(nav)).toEqual(['top']);
    const sw = nav.element!.querySelector<HTMLInputElement>('[role=switch]')!;
    expect(sw.closest('label')!.textContent).toContain('Show hidden bookmarks');
    sw.checked = true;
    sw.dispatchEvent(new Event('change'));
    expect(labels(nav)).toEqual(['top', '_Toc1']);
    nav.tab = 'links';
    expect(nav.element!.querySelector<HTMLElement>('.dk-nav-hidden')!.hidden).toBe(true);
    expect(labels(nav)).toEqual(['the website']);
  });

  it('filters the list while typing in the search box', () => {
    const s = stub();
    const nav = new NavigationController(s.viewer);
    nav.show();
    const input = nav.element!.querySelector<HTMLInputElement>('input[type=search]')!;
    input.value = 'pay';
    input.dispatchEvent(new Event('input'));
    expect(labels(nav)).toEqual(['2. Payment terms']);
    input.value = 'zzz';
    input.dispatchEvent(new Event('input'));
    expect(visibleEmpty(nav)).toBe('No matches');
  });

  it('jumps with history when an item is clicked, and indents by level', () => {
    const s = stub();
    const nav = new NavigationController(s.viewer);
    nav.show();
    const buttons = nav.element!.querySelectorAll<HTMLButtonElement>('.dk-nav-list button');
    buttons[1]!.click();
    expect(s.viewer.goTo).toHaveBeenCalledWith(s.targets.headings[1], { history: true });
    expect(buttons[2]!.style.paddingLeft).not.toBe(buttons[0]!.style.paddingLeft);
  });

  it('steps with the per-tab previous and next buttons and the history buttons', () => {
    const s = stub();
    const nav = new NavigationController(s.viewer);
    nav.show();
    nav.element!.querySelector<HTMLButtonElement>('[aria-label="Next heading"]')!.click();
    expect(s.viewer.next).toHaveBeenCalledWith('heading', { includeHidden: false });
    nav.element!.querySelector<HTMLButtonElement>('[aria-label="Previous heading"]')!.click();
    expect(s.viewer.previous).toHaveBeenCalledWith('heading', { includeHidden: false });
    const back = nav.element!.querySelector<HTMLButtonElement>('[aria-label="Back"]')!;
    expect(back.disabled).toBe(true);
    s.viewer.canGoBack = true;
    s.viewer.emit('navigate', { location: { page: 0, x: 0, y: 0 } });
    expect(back.disabled).toBe(false);
    back.click();
    expect(s.viewer.back).toHaveBeenCalled();
  });

  it('renders empty states', () => {
    const s = stub({ empty: true });
    const nav = new NavigationController(s.viewer);
    nav.show();
    const empty = () => visibleEmpty(nav);
    expect(empty()).toBe('No headings');
    nav.tab = 'bookmarks';
    expect(empty()).toBe('No bookmarks');
    nav.tab = 'links';
    expect(empty()).toBe('No links');
  });

  it('folds into a 36px rail with four buttons and unfolds on the clicked tab', () => {
    const s = stub({ comments: true });
    const nav = new NavigationController(s.viewer);
    nav.show();
    nav.collapsed = true;
    expect(nav.collapsed).toBe(true);
    expect(nav.element!.classList.contains('dk-rail')).toBe(true);
    const buttons = [...nav.element!.querySelectorAll('button')];
    expect(buttons.map((b) => b.getAttribute('aria-label'))).toEqual(['Outline', 'Bookmarks', 'Links', 'Comments', 'Show navigation']);
    expect(buttons.slice(0, 4).every((b) => !!b.title)).toBe(true);
    expect(s.emitted.at(-1)!.payload).toEqual({ part: 'navigation', visible: true, collapsed: true });
    buttons[1]!.click();
    expect(nav.collapsed).toBe(false);
    expect(nav.tab).toBe('bookmarks');
    expect(nav.element!.classList.contains('dk-rail')).toBe(false);
    expect(s.emitted.at(-1)!.payload).toEqual({ part: 'navigation', visible: true, collapsed: false });
  });

  it('shows the rail Comments button only when comments exist, and goes to the next one below the viewport top', () => {
    const none = stub();
    const a = new NavigationController(none.viewer, { collapsed: true });
    a.show();
    expect(a.element!.querySelector('[aria-label="Comments"]')).toBeNull();

    const s = stub({ comments: true });
    const nav = new NavigationController(s.viewer, { collapsed: true });
    nav.show();
    const btn = nav.element!.querySelector<HTMLButtonElement>('[aria-label="Comments"]')!;
    btn.click();
    expect(s.viewer.scrollToComment).toHaveBeenLastCalledWith('c2');
    s.viewer.location = { page: 1, x: 0, y: 500 };
    btn.click();
    expect(s.viewer.scrollToComment).toHaveBeenLastCalledWith('c1');
  });

  it('starts collapsed when the frame is narrower than 720px and no collapsed option was given', () => {
    const narrow = new NavigationController(stub({ width: 600 }).viewer);
    narrow.show();
    expect(narrow.collapsed).toBe(true);
    const wide = new NavigationController(stub({ width: 900 }).viewer);
    wide.show();
    expect(wide.collapsed).toBe(false);
  });

  it('keeps an explicit collapsed: false open even in a narrow frame', () => {
    const nav = new NavigationController(stub({ width: 600 }).viewer, { collapsed: false });
    nav.show();
    expect(nav.collapsed).toBe(false);
  });

  it('hides and shows again, announcing each change', () => {
    const s = stub();
    const nav = new NavigationController(s.viewer);
    nav.show();
    nav.hide();
    expect(nav.visible).toBe(false);
    expect(nav.element!.hidden).toBe(true);
    expect(s.emitted.at(-1)!.payload).toMatchObject({ part: 'navigation', visible: false });
    nav.show();
    expect(nav.element!.hidden).toBe(false);
  });

  it('refreshes the list after each layout', () => {
    const s = stub();
    const nav = new NavigationController(s.viewer);
    nav.show();
    s.targets.headings.push(target('heading', 'Appendix', 2, 10, { level: 0 }));
    s.viewer.emit('layout', { pages: [] });
    expect(labels(nav)).toContain('Appendix');
  });

  it('marks the current heading as the viewer scrolls (throttled)', () => {
    vi.useFakeTimers();
    const s = stub();
    const nav = new NavigationController(s.viewer);
    nav.show();
    const current = () => nav.element!.querySelector('[aria-current=location] .dk-nav-label')?.textContent;
    expect(current()).toBe('1. Introduction');
    s.viewer.location = { page: 0, x: 0, y: 500 };
    s.root.dispatchEvent(new Event('scroll'));
    s.viewer.location = { page: 1, x: 0, y: 100 };
    s.root.dispatchEvent(new Event('scroll'));
    expect(current()).toBe('1. Introduction');
    vi.advanceTimersByTime(200);
    expect(current()).toBe('Net 30');
  });

  it('keeps the current row in view by scrolling the list, not the page', () => {
    vi.useFakeTimers();
    const s = stub();
    const nav = new NavigationController(s.viewer);
    const scrollIntoView = vi.fn();
    const proto = HTMLElement.prototype as { scrollIntoView?: unknown };
    const had = Object.getOwnPropertyDescriptor(proto, 'scrollIntoView');
    proto.scrollIntoView = scrollIntoView;
    try {
      nav.show();
      const list = nav.element!.querySelector<HTMLElement>('.dk-nav-list')!;
      const row = [...list.querySelectorAll('button')].find((b) => b.textContent?.includes('Net 30'))!;
      const rect = (top: number, height: number) => ({ top, bottom: top + height, height, left: 0, right: 200, width: 200, x: 0, y: top, toJSON() {} }) as DOMRect;
      // The list shows 100px; the "Net 30" row sits below that.
      list.getBoundingClientRect = () => rect(0, 100);
      row.getBoundingClientRect = () => rect(150 - list.scrollTop, 26);
      s.viewer.location = { page: 1, x: 0, y: 100 };
      s.root.dispatchEvent(new Event('scroll'));
      vi.advanceTimersByTime(200);
      expect(row.getAttribute('aria-current')).toBe('location');
      expect(list.scrollTop).toBe(76);
      expect(scrollIntoView).not.toHaveBeenCalled();
    } finally {
      if (had) Object.defineProperty(proto, 'scrollIntoView', had);
      else delete proto.scrollIntoView;
    }
  });

  it('show(tab) unfolds onto that tab; show() keeps the fold', () => {
    const s = stub();
    const nav = new NavigationController(s.viewer, { collapsed: true });
    nav.show();
    expect(nav.collapsed).toBe(true);
    nav.hide();
    nav.show();
    expect(nav.collapsed).toBe(true);
    nav.show('links');
    expect(nav.collapsed).toBe(false);
    expect(nav.tab).toBe('links');
    expect(nav.element!.classList.contains('dk-rail')).toBe(false);
    expect(nav.element!.querySelector('[role=tab][aria-selected=true]')!.textContent).toBe('Links');
    expect(s.emitted.at(-1)).toEqual({ event: 'chromechange', payload: { part: 'navigation', visible: true, collapsed: false } });
    // From hidden and folded, in one step.
    const t = stub();
    const other = new NavigationController(t.viewer, { collapsed: true });
    other.show('bookmarks');
    expect(other.collapsed).toBe(false);
    expect(other.visible).toBe(true);
    expect(other.tab).toBe('bookmarks');
    expect(t.emitted.filter((e) => e.event === 'chromechange')).toEqual([{ event: 'chromechange', payload: { part: 'navigation', visible: true, collapsed: false } }]);
  });

  it('shows text labels on the tabs', () => {
    const nav = new NavigationController(stub().viewer);
    nav.show();
    expect([...nav.element!.querySelectorAll('[role=tab]')].map((t) => t.textContent)).toEqual(['Outline', 'Bookmarks', 'Links']);
  });

  it('keeps the same hidden-bookmarks switch, with focus, when toggled', () => {
    const nav = new NavigationController(stub().viewer, { tab: 'bookmarks' });
    nav.show();
    const sw = nav.element!.querySelector<HTMLInputElement>('[role=switch]')!;
    sw.focus();
    sw.checked = true;
    sw.dispatchEvent(new Event('change'));
    expect(nav.element!.querySelector('[role=switch]')).toBe(sw);
    expect(document.activeElement).toBe(sw);
  });

  it('moves focus to the counterpart control when folding and unfolding from inside the panel', () => {
    const nav = new NavigationController(stub({ width: 900 }).viewer);
    nav.show();
    const fold = nav.element!.querySelector<HTMLButtonElement>('[aria-label="Fold navigation"]')!;
    fold.focus();
    fold.click();
    expect(nav.element!.contains(document.activeElement)).toBe(true);
    expect(document.activeElement!.getAttribute('aria-label')).toBe('Show navigation');
    const rail = nav.element!.querySelector<HTMLButtonElement>('[aria-label="Bookmarks"]')!;
    rail.focus();
    rail.click();
    expect(document.activeElement!.getAttribute('role')).toBe('tab');
    expect(document.activeElement!.getAttribute('aria-selected')).toBe('true');
    expect(document.activeElement!.textContent).toBe('Bookmarks');
  });

  it('does not steal focus when collapsed is set from code', () => {
    const nav = new NavigationController(stub().viewer);
    nav.show();
    const other = document.createElement('button');
    document.body.appendChild(other);
    other.focus();
    nav.collapsed = true;
    expect(document.activeElement).toBe(other);
    nav.collapsed = false;
    expect(document.activeElement).toBe(other);
  });
});
