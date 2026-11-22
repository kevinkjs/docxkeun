// @vitest-environment happy-dom
import { afterEach, describe, expect, it, vi } from 'vitest';
import { openMenu, type MenuHandle } from '../src/render/ui/menu.js';
import { promptPopover } from '../src/render/ui/popover.js';

const key = (el: Element, k: string) => el.dispatchEvent(new KeyboardEvent('keydown', { key: k, bubbles: true, cancelable: true }));
let open: MenuHandle | undefined;
afterEach(() => {
  open?.close();
  open = undefined;
  document.body.innerHTML = '';
});
const rows = (el: HTMLElement) => [...el.querySelectorAll<HTMLElement>('[role=menuitem]')];

describe('openMenu', () => {
  it('renders menu, items and separators with roles', () => {
    open = openMenu([{ label: 'A' }, '-', { label: 'B', shortcut: 'Ctrl+B' }], { x: 10, y: 10 });
    expect(open.element.getAttribute('role')).toBe('menu');
    expect(open.element.classList.contains('docxkeun-ui-root')).toBe(true);
    expect(rows(open.element)).toHaveLength(2);
    expect(open.element.querySelectorAll('[role=separator]')).toHaveLength(1);
    expect(open.element.textContent).toContain('Ctrl+B');
  });

  it('focuses the first enabled item and arrows skip disabled items and separators', () => {
    open = openMenu([{ label: 'A', disabled: true }, { label: 'B' }, '-', { label: 'C', disabled: 'no' }, { label: 'D' }], { x: 0, y: 0 });
    const r = rows(open.element);
    expect(document.activeElement).toBe(r[1]);
    key(open.element, 'ArrowDown');
    expect(document.activeElement).toBe(r[3]);
    key(open.element, 'ArrowDown');
    expect(document.activeElement).toBe(r[1]);
    key(open.element, 'ArrowUp');
    expect(document.activeElement).toBe(r[3]);
  });

  it('Enter runs the item and closes the menu first', () => {
    let openAtRun: boolean | undefined;
    const run = vi.fn(() => {
      openAtRun = open!.element.isConnected;
    });
    open = openMenu([{ label: 'A', run }], { x: 0, y: 0 });
    key(open.element, 'Enter');
    expect(run).toHaveBeenCalledOnce();
    expect(openAtRun).toBe(false);
    expect(open.element.isConnected).toBe(false);
  });

  it('does not run disabled items and sets aria-disabled and title', () => {
    const run = vi.fn();
    open = openMenu([{ label: 'A', disabled: 'Because', run }, { label: 'B' }], { x: 0, y: 0 });
    const a = rows(open.element)[0]!;
    expect(a.getAttribute('aria-disabled')).toBe('true');
    expect(a.title).toBe('Because');
    a.click();
    expect(run).not.toHaveBeenCalled();
    expect(open.element.isConnected).toBe(true);
  });

  it('closes before run and reports a rejected run promise via console.error', async () => {
    const err = vi.spyOn(console, 'error').mockImplementation(() => {});
    try {
      const boom = new Error('x');
      let openAtRun: boolean | undefined;
      open = openMenu([{ label: 'A', run: () => { openAtRun = open!.element.isConnected; return Promise.reject(boom); } }], { x: 0, y: 0 });
      rows(open.element)[0]!.click();
      await new Promise((r) => setTimeout(r, 0));
      expect(openAtRun).toBe(false);
      expect(err).toHaveBeenCalledWith(boom);
    } finally {
      err.mockRestore();
    }
  });

  it('opens a submenu with ArrowRight and closes it with ArrowLeft', () => {
    const sub = vi.fn();
    open = openMenu([{ label: 'Format', items: [{ label: 'Bold', run: sub }] }], { x: 0, y: 0 });
    const parent = rows(open.element)[0]!;
    expect(parent.getAttribute('aria-haspopup')).toBe('menu');
    expect(parent.getAttribute('aria-expanded')).toBe('false');
    key(open.element, 'ArrowRight');
    expect(parent.getAttribute('aria-expanded')).toBe('true');
    const menus = document.querySelectorAll('[role=menu]');
    expect(menus).toHaveLength(2);
    const submenu = menus[1] as HTMLElement;
    expect(document.activeElement?.textContent).toContain('Bold');
    key(submenu, 'ArrowLeft');
    expect(document.querySelectorAll('[role=menu]')).toHaveLength(1);
    expect(parent.getAttribute('aria-expanded')).toBe('false');
    expect(document.activeElement).toBe(parent);
  });

  it('clicking an item with items opens the submenu', () => {
    open = openMenu([{ label: 'Format', items: [{ label: 'Bold' }] }], { x: 0, y: 0 });
    rows(open.element)[0]!.click();
    expect(document.querySelectorAll('[role=menu]')).toHaveLength(2);
  });

  it('Esc closes everything and calls onClose; focus returns', () => {
    const btn = document.createElement('button');
    document.body.appendChild(btn);
    btn.focus();
    const onClose = vi.fn();
    open = openMenu([{ label: 'A' }], { x: 0, y: 0 }, { onClose });
    key(open.element, 'Escape');
    expect(onClose).toHaveBeenCalledOnce();
    expect(document.querySelectorAll('[role=menu]')).toHaveLength(0);
    expect(document.activeElement).toBe(btn);
  });

  it('closes on an outside mousedown but not an inside one', () => {
    const onClose = vi.fn();
    open = openMenu([{ label: 'A' }], { x: 0, y: 0 }, { onClose });
    rows(open.element)[0]!.dispatchEvent(new MouseEvent('mousedown', { bubbles: true }));
    expect(onClose).not.toHaveBeenCalled();
    document.body.dispatchEvent(new MouseEvent('mousedown', { bubbles: true }));
    expect(onClose).toHaveBeenCalledOnce();
  });

  it('appends into options.container', () => {
    const host = document.createElement('div');
    document.body.appendChild(host);
    open = openMenu([{ label: 'A' }], { x: 0, y: 0 }, { container: host });
    expect(open.element.parentElement).toBe(host);
  });

  it('only one menu is open at a time', () => {
    const closed = vi.fn();
    const first = openMenu([{ label: 'A' }], { x: 0, y: 0 }, { onClose: closed });
    open = openMenu([{ label: 'B' }], { x: 0, y: 0 });
    expect(closed).toHaveBeenCalledOnce();
    expect(first.element.isConnected).toBe(false);
    expect(document.querySelectorAll('[role=menu]')).toHaveLength(1);
  });

  it('renders a non-icon string icon as text, never HTML', () => {
    open = openMenu([{ label: 'A', icon: '✦' }, { label: 'B', icon: '<b>x</b>' }, { label: 'C', icon: 'copy' }], { x: 0, y: 0 });
    const r = rows(open.element);
    expect(r[0]!.querySelector('.dk-menu-icon')!.textContent).toBe('✦');
    expect(r[1]!.querySelector('b')).toBeNull();
    expect(r[1]!.querySelector('.dk-menu-icon')!.textContent).toBe('<b>x</b>');
    expect(r[2]!.querySelector('svg')).not.toBeNull();
  });
});

describe('promptPopover', () => {
  const input = () => document.querySelector('input') as HTMLInputElement;
  it('resolves the typed value on Enter', async () => {
    const p = promptPopover({ title: 'Link', value: 'a', at: { x: 0, y: 0 } });
    input().value = 'https://x';
    key(input(), 'Enter');
    expect(await p).toBe('https://x');
    expect(document.querySelector('.dk-popover')).toBeNull();
  });
  it('resolves undefined on Esc, Cancel and outside click', async () => {
    let p = promptPopover({ title: 'T', at: { x: 0, y: 0 } });
    key(input(), 'Escape');
    expect(await p).toBeUndefined();
    p = promptPopover({ title: 'T', at: { x: 0, y: 0 } });
    (document.querySelector('[data-role=cancel]') as HTMLElement).click();
    expect(await p).toBeUndefined();
    p = promptPopover({ title: 'T', at: { x: 0, y: 0 } });
    document.body.dispatchEvent(new MouseEvent('mousedown', { bubbles: true }));
    expect(await p).toBeUndefined();
  });
  it('shows validate message and stays open', async () => {
    const p = promptPopover({ title: 'T', at: { x: 0, y: 0 }, validate: (v) => (v ? undefined : 'Required') });
    key(input(), 'Enter');
    expect(document.querySelector('[role=alert]')!.textContent).toBe('Required');
    expect(document.querySelector('.dk-popover')).not.toBeNull();
    input().value = 'ok';
    key(input(), 'Enter');
    expect(await p).toBe('ok');
  });
  it('honours the container option and carries the theme root class', () => {
    const host = document.createElement('div');
    document.body.appendChild(host);
    const p = promptPopover({ title: 'T', at: { x: 0, y: 0 }, container: host });
    const pop = host.querySelector('.dk-popover')!;
    expect(pop.classList.contains('docxkeun-ui-root')).toBe(true);
    key(input(), 'Escape');
    return p;
  });
});
