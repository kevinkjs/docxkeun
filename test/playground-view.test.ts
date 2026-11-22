import { describe, expect, it } from 'vitest';
import { FOLD_KEY, fitZoom, readFold, writeFold } from '../playground/view.js';

function memory(initial?: string) {
  const data = new Map<string, string>(initial === undefined ? [] : [[FOLD_KEY, initial]]);
  return { data, store: () => ({ getItem: (k: string) => data.get(k) ?? null, setItem: (k: string, v: string) => void data.set(k, v) }) };
}

describe('playground fold memory', () => {
  it('keeps the toolbar fold but never a hidden toolbar (editing off hides it)', () => {
    const m = memory();
    // setEditable(false) hides the folded toolbar and announces it.
    writeFold(m.store, 'toolbar', false, true);
    expect(readFold(m.store)).toEqual({ toolbar: { collapsed: true } });
    expect(JSON.parse(m.data.get(FOLD_KEY)!).toolbar).not.toHaveProperty('visible');
  });

  it('ignores a hidden toolbar stored by an older playground', () => {
    const m = memory(JSON.stringify({ toolbar: { visible: false, collapsed: false }, navigation: { visible: false, collapsed: true } }));
    expect(readFold(m.store)).toEqual({ toolbar: { collapsed: false }, navigation: { visible: false, collapsed: true } });
  });

  it('keeps both states for the navigation panel, next to the toolbar', () => {
    const m = memory();
    writeFold(m.store, 'toolbar', true, true);
    writeFold(m.store, 'navigation', false, false);
    expect(readFold(m.store)).toEqual({ toolbar: { collapsed: true }, navigation: { visible: false, collapsed: false } });
  });

  it('survives blocked or broken storage', () => {
    const blocked = () => {
      throw new Error('SecurityError');
    };
    expect(readFold(blocked)).toEqual({});
    expect(() => writeFold(blocked, 'toolbar', true, true)).not.toThrow();
    expect(readFold(memory('{not json').store)).toEqual({});
  });
});

describe('playground fit to width', () => {
  const letter = [612]; // 816 CSS px

  it('fits the page and the comment column, never above 100%', () => {
    expect(fitZoom(2000, letter, 272)).toBe(1);
    // ~1280px window with the 240px navigation panel: the root is ~1040px and cards take 272.
    const z = fitZoom(1040, letter, 272);
    expect(z).toBeCloseTo((1040 - 32 - 272 - 16) / 816, 5);
    expect(816 * z + 32 + 272 + 16).toBeLessThanOrEqual(1040);
  });

  it('never goes below 50%, and uses the widest page', () => {
    expect(fitZoom(300, letter, 272)).toBe(0.5);
    expect(fitZoom(1040, [612, 792], 0)).toBeCloseTo((1040 - 48) / 1056, 5);
  });

  it('leaves 100% without pages or a measured width', () => {
    expect(fitZoom(1040, [], 0)).toBe(1);
    expect(fitZoom(0, letter, 0)).toBe(1);
  });
});
