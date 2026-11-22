import { describe, expect, it } from 'vitest';
import { currentHeading, filterTargets } from '../src/render/ui/nav-model.js';
import type { NavTarget } from '../src/render/navigation.js';

const t = (label: string, page = 0, y = 0, extra: Partial<NavTarget> = {}): NavTarget => ({ kind: 'heading', label, page, x: 0, y, rects: [], ...extra });

describe('filterTargets', () => {
  const list = [t('1. Introduction'), t('2. Payment terms'), t('Café menu')];
  it('matches case-insensitively on the label', () => {
    expect(filterTargets(list, 'pay').map((x) => x.label)).toEqual(['2. Payment terms']);
  });
  it('ignores accents in both directions', () => {
    expect(filterTargets(list, 'cafe')).toHaveLength(1);
    expect(filterTargets([t('Cafe')], 'é')).toHaveLength(1);
  });
  it('returns everything for an empty query', () => {
    expect(filterTargets(list, '  ')).toBe(list);
  });
  it('matches links on URL and anchor', () => {
    const links = [t('click here', 0, 0, { kind: 'link', link: { url: 'https://example.com/Docs' } as never }), t('see terms', 0, 5, { kind: 'link', link: { anchor: 'terms' } as never })];
    expect(filterTargets(links, 'example.com').map((x) => x.label)).toEqual(['click here']);
    expect(filterTargets(links, 'TERMS').map((x) => x.label)).toEqual(['see terms']);
  });
});

describe('currentHeading', () => {
  const heads = [t('A', 0, 100), t('B', 0, 400), t('C', 2, 50)];
  it('is undefined before the first heading', () => {
    expect(currentHeading(heads, { page: 0, x: 0, y: 50 })).toBeUndefined();
    expect(currentHeading([], { page: 0, x: 0, y: 50 })).toBeUndefined();
  });
  it('picks the last heading at or above the location', () => {
    expect(currentHeading(heads, { page: 0, x: 0, y: 100 })?.label).toBe('A');
    expect(currentHeading(heads, { page: 0, x: 0, y: 399 })?.label).toBe('A');
    expect(currentHeading(heads, { page: 1, x: 0, y: 10 })?.label).toBe('B');
  });
  it('works on a later page', () => {
    expect(currentHeading(heads, { page: 2, x: 0, y: 60 })?.label).toBe('C');
  });
  it('counts a heading just below the location within the tolerance', () => {
    expect(currentHeading(heads, { page: 0, x: 0, y: 80 }, 24)?.label).toBe('A');
  });
});
