/**
 * A keyboard-accessible popup menu (the context menu is built on it): grouped rows with icons and
 * right-aligned shortcuts, submenus, disabled items with a reason. `position: fixed`, so `at` is
 * a viewport point such as a contextmenu event's clientX/clientY.
 */
import { icon, isIconName } from './icons.js';
import { ensureUiStyle } from './theme.js';

export type MenuItem =
  | '-'
  | {
      id?: string;
      label: string;
      icon?: import('./icons.js').IconName | string;
      shortcut?: string;
      /** `true` disables the item; a string also shows as the tooltip (the reason). */
      disabled?: boolean | string;
      run?: () => void | Promise<void>;
      items?: MenuItem[];
    };

export interface MenuHandle {
  close(): void;
  readonly element: HTMLElement;
}

export interface MenuOptions {
  label?: string;
  onClose?: () => void;
  /** Where to append the menu (default `document.body`). */
  container?: HTMLElement;
}

/** Place a `size` box at `at`, flipping left/up when it would overflow, then clamping to the margin. */
export function positionMenu(
  at: { x: number; y: number },
  size: { w: number; h: number },
  viewport: { w: number; h: number },
  margin = 4,
): { x: number; y: number } {
  let x = at.x;
  let y = at.y;
  if (x + size.w + margin > viewport.w) x = at.x - size.w;
  if (y + size.h + margin > viewport.h) y = at.y - size.h;
  x = Math.max(margin, Math.min(x, viewport.w - size.w - margin));
  y = Math.max(margin, Math.min(y, viewport.h - size.h - margin));
  return { x, y };
}

const STYLE_ID = 'docxkeun-menu-style';
const CSS = `
.dk-menu{position:fixed;z-index:2147483000;min-width:220px;max-width:min(360px,calc(100vw - 8px));max-height:calc(100vh - 8px);overflow-y:auto;box-sizing:border-box;padding:4px;background:var(--dk-ui-bg);color:var(--dk-ui-fg);border:1px solid var(--dk-ui-line);border-radius:var(--dk-ui-radius);box-shadow:var(--dk-ui-shadow);font:var(--dk-ui-font);outline:none}
.dk-menu-item{display:flex;align-items:center;gap:8px;min-height:28px;padding:3px 8px;border-radius:4px;cursor:pointer;white-space:nowrap;outline:none}
.dk-menu-item:hover,.dk-menu-item:focus{background:var(--dk-ui-hover)}
.dk-menu-item:focus-visible{outline:2px solid var(--dk-ui-accent);outline-offset:-2px}
.dk-menu-item[aria-disabled=true]{opacity:.45;cursor:default}
.dk-menu-item[aria-disabled=true]:hover{background:transparent}
.dk-menu-icon{flex:none;width:16px;height:16px;display:inline-flex;align-items:center;justify-content:center;color:var(--dk-ui-muted);font-size:13px;line-height:1}
.dk-menu-label{flex:1 1 auto;overflow:hidden;text-overflow:ellipsis}
.dk-menu-shortcut{flex:none;margin-left:24px;color:var(--dk-ui-muted);font-size:12px}
.dk-menu-arrow{flex:none;margin-left:16px;color:var(--dk-ui-muted)}
.dk-menu-sep{height:1px;margin:4px 6px;background:var(--dk-ui-line)}
`;

function ensureMenuStyle(): void {
  ensureUiStyle();
  if (typeof document === 'undefined' || document.getElementById(STYLE_ID)) return;
  const style = document.createElement('style');
  style.id = STYLE_ID;
  style.textContent = CSS;
  document.head.appendChild(style);
}

type Item = Exclude<MenuItem, '-'>;
const isDisabled = (item: Item) => !!item.disabled;

let current: MenuHandle | undefined;

interface Level {
  el: HTMLElement;
  parentRow?: HTMLElement;
}

export function openMenu(items: MenuItem[], at: { x: number; y: number }, options: MenuOptions = {}): MenuHandle {
  current?.close();
  ensureMenuStyle();
  const container = options.container ?? document.body;
  const previous = document.activeElement instanceof HTMLElement ? document.activeElement : null;
  const levels: Level[] = [];
  let closed = false;

  const close = (): void => {
    if (closed) return;
    closed = true;
    if (current === handle) current = undefined;
    document.removeEventListener('mousedown', onOutside, true);
    for (const l of levels) l.el.remove();
    levels.length = 0;
    if (previous?.isConnected) previous.focus();
    options.onClose?.();
  };

  const onOutside = (e: Event): void => {
    const t = e.target as Node;
    if (!levels.some((l) => l.el.contains(t))) close();
  };

  const closeFrom = (depth: number): void => {
    while (levels.length > depth) {
      const l = levels.pop()!;
      l.parentRow?.setAttribute('aria-expanded', 'false');
      l.el.remove();
    }
  };

  const enabledRows = (el: HTMLElement) =>
    [...el.querySelectorAll<HTMLElement>('[role=menuitem]')].filter((r) => r.getAttribute('aria-disabled') !== 'true');

  const activate = (item: Item, row: HTMLElement, depth: number, viaKeyboard: boolean): void => {
    if (isDisabled(item)) return;
    if (item.items?.length) {
      openSub(item.items, row, depth, viaKeyboard);
      return;
    }
    if (!item.run) return;
    const run = item.run;
    close();
    try {
      const result = run();
      if (result && typeof (result as Promise<void>).catch === 'function') {
        (result as Promise<void>).catch((err) => console.error(err));
      }
    } catch (err) {
      console.error(err);
    }
  };

  const build = (list: MenuItem[], depth: number, label?: string): HTMLElement => {
    const el = document.createElement('div');
    el.className = 'dk-menu docxkeun-ui-root';
    el.setAttribute('role', 'menu');
    el.tabIndex = -1;
    if (label) el.setAttribute('aria-label', label);
    for (const item of list) {
      if (item === '-') {
        const sep = document.createElement('div');
        sep.className = 'dk-menu-sep';
        sep.setAttribute('role', 'separator');
        el.appendChild(sep);
        continue;
      }
      const row = document.createElement('div');
      row.className = 'dk-menu-item';
      row.setAttribute('role', 'menuitem');
      row.tabIndex = -1;
      if (item.id) row.dataset.id = item.id;
      if (isDisabled(item)) {
        row.setAttribute('aria-disabled', 'true');
        if (typeof item.disabled === 'string') row.title = item.disabled;
      }
      const ic = document.createElement('span');
      ic.className = 'dk-menu-icon';
      if (item.icon) {
        if (isIconName(item.icon)) ic.innerHTML = icon(item.icon);
        else ic.textContent = item.icon;
      }
      ic.setAttribute('aria-hidden', 'true');
      const text = document.createElement('span');
      text.className = 'dk-menu-label';
      text.textContent = item.label;
      row.append(ic, text);
      if (item.shortcut) {
        const sc = document.createElement('span');
        sc.className = 'dk-menu-shortcut';
        sc.textContent = item.shortcut;
        row.appendChild(sc);
      }
      if (item.items?.length) {
        row.setAttribute('aria-haspopup', 'menu');
        row.setAttribute('aria-expanded', 'false');
        const arrow = document.createElement('span');
        arrow.className = 'dk-menu-arrow';
        arrow.setAttribute('aria-hidden', 'true');
        arrow.innerHTML = icon('forward');
        row.appendChild(arrow);
      }
      row.addEventListener('mouseenter', () => {
        if (isDisabled(item)) return;
        row.focus();
        if (item.items?.length) openSub(item.items, row, depth, false);
        else closeFrom(depth + 1);
      });
      row.addEventListener('click', (e) => {
        e.stopPropagation();
        activate(item, row, depth, false);
      });
      (row as any)._dkItem = item;
      el.appendChild(row);
    }
    el.addEventListener('keydown', (e) => onKey(e, el, depth));
    return el;
  };

  const place = (el: HTMLElement, pos: { x: number; y: number }, parentRow?: HTMLElement): void => {
    const rect = el.getBoundingClientRect();
    const size = { w: rect.width, h: rect.height };
    const vp = { w: window.innerWidth, h: window.innerHeight };
    let p: { x: number; y: number };
    if (parentRow) {
      const pr = parentRow.getBoundingClientRect();
      p = positionMenu({ x: pr.right, y: pr.top }, size, vp);
      if (p.x < pr.right) p = { x: Math.max(4, pr.left - size.w), y: p.y };
    } else p = positionMenu(pos, size, vp);
    el.style.left = `${p.x}px`;
    el.style.top = `${p.y}px`;
  };

  const show = (list: MenuItem[], pos: { x: number; y: number }, depth: number, parentRow?: HTMLElement, label?: string): HTMLElement => {
    const el = build(list, depth, label);
    el.style.left = '0px';
    el.style.top = '0px';
    container.appendChild(el);
    place(el, pos, parentRow);
    levels.push({ el, parentRow });
    return el;
  };

  function openSub(list: MenuItem[], row: HTMLElement, depth: number, focusFirst: boolean): void {
    if (levels[depth + 1]?.parentRow === row) {
      if (focusFirst) enabledRows(levels[depth + 1]!.el)[0]?.focus();
      return;
    }
    closeFrom(depth + 1);
    row.setAttribute('aria-expanded', 'true');
    const el = show(list, { x: 0, y: 0 }, depth + 1, row);
    if (focusFirst) enabledRows(el)[0]?.focus();
  }

  function onKey(e: KeyboardEvent, el: HTMLElement, depth: number): void {
    const rows = enabledRows(el);
    const active = rows.indexOf(document.activeElement as HTMLElement);
    const focusAt = (i: number) => rows[(i + rows.length) % rows.length]?.focus();
    const stop = () => {
      e.preventDefault();
      e.stopPropagation();
    };
    switch (e.key) {
      case 'ArrowDown':
        stop();
        focusAt(active < 0 ? 0 : active + 1);
        break;
      case 'ArrowUp':
        stop();
        focusAt(active < 0 ? rows.length - 1 : active - 1);
        break;
      case 'Home':
        stop();
        focusAt(0);
        break;
      case 'End':
        stop();
        focusAt(rows.length - 1);
        break;
      case 'ArrowRight': {
        const row = rows[active];
        const item = row && ((row as any)._dkItem as Item);
        if (row && item?.items?.length) {
          stop();
          openSub(item.items, row, depth, true);
        }
        break;
      }
      case 'ArrowLeft':
        if (depth > 0) {
          stop();
          const parentRow = levels[depth]!.parentRow;
          closeFrom(depth);
          parentRow?.focus();
        }
        break;
      case 'Enter':
      case ' ': {
        const row = rows[active];
        if (row) {
          stop();
          activate((row as any)._dkItem as Item, row, depth, true);
        }
        break;
      }
      case 'Escape':
      case 'Tab':
        stop();
        close();
        break;
    }
  }

  const root = show(items, at, 0, undefined, options.label);
  const handle: MenuHandle = { close, element: root };
  current = handle;
  document.addEventListener('mousedown', onOutside, true);
  (enabledRows(root)[0] ?? root).focus();
  return handle;
}
