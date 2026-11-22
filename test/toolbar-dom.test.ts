// @vitest-environment happy-dom
import { describe, expect, it } from 'vitest';
import { ToolbarController } from '../src/render/toolbar.js';

function stub(withCaret = true) {
  const listeners = new Map<string, Array<(p: any) => void>>();
  const emitted: Array<{ event: string; payload: any }> = [];
  const root = document.createElement('div');
  const slot = document.createElement('div');
  document.body.append(slot, root);
  const format = { bold: true, italic: false, underline: false, strike: false, align: 'left', direction: 'ltr', story: 'body', hasSelection: false, trackChanges: false, canUndo: false, canRedo: false };
  const editing = { caretFormat: () => (withCaret ? format : undefined), canUndo: true, canRedo: false, focus() {}, undo() {}, redo() {} };
  const viewer: any = {
    root,
    editing,
    document: undefined,
    on: (event: string, fn: (p: any) => void) => {
      (listeners.get(event) ?? listeners.set(event, []).get(event)!).push(fn);
      return () => {};
    },
    emit: (event: string, payload: any) => {
      emitted.push({ event, payload });
      for (const l of listeners.get(event) ?? []) l(payload);
    },
    frame: () => ({ toolbarSlot: slot }),
  };
  return { viewer, slot, emitted, format, fire: (e: string, p: any) => viewer.emit(e, p) };
}

describe('ToolbarController', () => {
  it('builds a role=toolbar into the frame slot on first show', () => {
    const s = stub();
    const tb = new ToolbarController(s.viewer);
    expect(tb.element).toBeUndefined();
    tb.show();
    expect(tb.visible).toBe(true);
    expect(tb.element!.getAttribute('role')).toBe('toolbar');
    expect(tb.element!.parentElement).toBe(s.slot);
  });

  it('orders the groups with separators', () => {
    const tb = new ToolbarController(stub().viewer);
    tb.show();
    const groups = [...tb.element!.querySelectorAll('.dk-tb-group')].map((g) => g.getAttribute('data-group'));
    expect(groups).toEqual(['history', 'text', 'font', 'align', 'list', 'direction', 'track']);
    expect(tb.element!.querySelectorAll('.dk-tb-sep').length).toBe(6);
    const first = tb.element!.querySelector('[data-group=history]')!;
    expect([...first.querySelectorAll('button')].map((b) => b.getAttribute('aria-label'))).toEqual(['Undo', 'Redo']);
  });

  it('uses a given container', () => {
    const c = document.createElement('div');
    const tb = new ToolbarController(stub().viewer, { container: c });
    tb.show();
    expect(tb.element!.parentElement).toBe(c);
  });

  it('folds into a strip and unfolds', () => {
    const s = stub();
    const tb = new ToolbarController(s.viewer);
    tb.show();
    (tb.element!.querySelector('[aria-label="Hide toolbar"]') as HTMLElement).click();
    expect(tb.collapsed).toBe(true);
    const strip = tb.element!.querySelector('.dk-tb-strip') as HTMLButtonElement;
    expect(strip.getAttribute('aria-label')).toBe('Show toolbar');
    expect(strip.hidden).toBe(false);
    expect(s.emitted.at(-1)).toEqual({ event: 'chromechange', payload: { part: 'toolbar', visible: true, collapsed: true } });
    strip.click();
    expect(tb.collapsed).toBe(false);
    expect(strip.hidden).toBe(true);
  });

  it('hide emits and sets hidden', () => {
    const s = stub();
    const tb = new ToolbarController(s.viewer);
    tb.show();
    tb.hide();
    expect(tb.element!.hidden).toBe(true);
    expect(tb.visible).toBe(false);
    expect(s.emitted.at(-1)).toEqual({ event: 'chromechange', payload: { part: 'toolbar', visible: false, collapsed: false } });
  });

  it('setEditable restores the previous visibility', () => {
    const tb = new ToolbarController(stub().viewer);
    tb.show();
    tb.setEditable(false);
    expect(tb.visible).toBe(false);
    tb.setEditable(true);
    expect(tb.visible).toBe(true);
    tb.hide();
    tb.setEditable(false);
    tb.setEditable(true);
    expect(tb.visible).toBe(false);
  });

  it('mirrors the caret and disables without one', () => {
    const s = stub();
    const tb = new ToolbarController(s.viewer);
    tb.show();
    const bold = tb.element!.querySelector('[aria-label="Bold"]') as HTMLButtonElement;
    expect(bold.getAttribute('aria-pressed')).toBe('true');
    expect(bold.disabled).toBe(false);
    s.fire('caretchange', { format: undefined });
    expect(bold.disabled).toBe(true);
  });

  it('roves focus with arrow keys', () => {
    const tb = new ToolbarController(stub().viewer);
    tb.show();
    const all = [...tb.element!.querySelectorAll<HTMLElement>('[data-nav]')];
    expect(all.filter((n) => n.tabIndex === 0).length).toBe(1);
    const nav = all.filter((n) => !(n as HTMLButtonElement).disabled);
    const stop = all.find((n) => n.tabIndex === 0) as HTMLButtonElement;
    expect(stop.disabled).toBe(false);
    nav[0]!.focus();
    nav[0]!.dispatchEvent(new KeyboardEvent('keydown', { key: 'ArrowRight', bubbles: true }));
    expect(document.activeElement).toBe(nav[1]);
    expect(nav[1]!.tabIndex).toBe(0);
    expect(nav[0]!.tabIndex).toBe(-1);
  });

  it('moves trailing groups into the more menu when narrow', () => {
    const tb = new ToolbarController(stub().viewer);
    tb.show();
    tb.relayout([40, 60, 80, 50, 50, 50, 50], 200);
    const more = tb.element!.querySelector('.dk-tb-more') as HTMLElement;
    expect(more.hidden).toBe(false);
    const pop = tb.element!.querySelector('.dk-tb-pop') as HTMLElement;
    expect([...pop.querySelectorAll('.dk-tb-group')].length).toBe(5);
    tb.relayout([40, 60, 80, 50, 50, 50, 50], 1000);
    expect(more.hidden).toBe(true);
    expect(pop.querySelectorAll('.dk-tb-group').length).toBe(0);
  });

  describe('overflow popover placement', () => {
    function open() {
      const tb = new ToolbarController(stub().viewer);
      tb.show();
      tb.relayout([40, 60, 80, 50, 50, 50, 50], 200);
      const el = tb.element!;
      const more = el.querySelector('.dk-tb-more button') as HTMLElement;
      const pop = el.querySelector('.dk-tb-pop') as HTMLElement;
      let btn = { left: 300, right: 330, top: 10, bottom: 40, width: 30, height: 30, x: 300, y: 10 };
      more.getBoundingClientRect = () => btn as DOMRect;
      pop.getBoundingClientRect = () => ({ left: 0, top: 0, right: 100, bottom: 80, width: 100, height: 80, x: 0, y: 0 }) as DOMRect;
      more.click();
      return { tb, more, pop, move: (b: Partial<typeof btn>) => (btn = { ...btn, ...b }) };
    }

    it('re-places on window resize and on relayout while open', () => {
      const { tb, pop, move } = open();
      expect(pop.hidden).toBe(false);
      expect(pop.style.left).toBe('230px');
      expect(pop.style.top).toBe('40px');
      move({ right: 430, bottom: 60 });
      window.dispatchEvent(new Event('resize'));
      expect(pop.style.left).toBe('330px');
      expect(pop.style.top).toBe('60px');
      move({ right: 530, bottom: 70 });
      tb.relayout([40, 60, 80, 50, 50, 50, 50], 200);
      expect(pop.hidden).toBe(false);
      expect(pop.style.left).toBe('430px');
      expect(pop.style.top).toBe('70px');
    });

    it('closes on page scroll but not on a scroll inside the popover, and detaches listeners after close', () => {
      const { pop, more } = open();
      pop.dispatchEvent(new Event('scroll'));
      expect(pop.hidden).toBe(false);
      document.body.dispatchEvent(new Event('scroll'));
      expect(pop.hidden).toBe(true);
      expect(more.getAttribute('aria-expanded')).toBe('false');
      expect(() => {
        document.body.dispatchEvent(new Event('scroll'));
        window.dispatchEvent(new Event('resize'));
      }).not.toThrow();
      expect(pop.style.left).toBe('230px');
    });

    it('Esc closes the popover and returns focus to the more button', () => {
      const { pop, more } = open();
      more.focus();
      more.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', bubbles: true }));
      expect(pop.hidden).toBe(true);
      expect(document.activeElement).toBe(more);
    });
  });

  it('keeps the tab stop on an enabled control without a caret', () => {
    const s = stub(false);
    const tb = new ToolbarController(s.viewer);
    tb.show();
    const all = [...tb.element!.querySelectorAll<HTMLElement>('[data-nav]')];
    const stops = all.filter((n) => n.tabIndex === 0);
    expect(stops.length).toBe(1);
    expect((stops[0] as HTMLButtonElement).disabled).toBe(false);
    expect(stops[0]!.getAttribute('aria-label')).toBe('Hide toolbar');
    s.fire('caretchange', { format: s.format });
    const after = all.filter((n) => n.tabIndex === 0);
    expect(after.length).toBe(1);
    expect((after[0] as HTMLButtonElement).disabled).toBe(false);
  });

  it('marks a custom-container toolbar as a ui root', () => {
    const c = document.createElement('div');
    const tb = new ToolbarController(stub().viewer, { container: c });
    tb.show();
    expect(tb.element!.classList.contains('docxkeun-ui-root')).toBe(true);
  });
});
