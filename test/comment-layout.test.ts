import { describe, expect, it } from 'vitest';
import { BUBBLE_SIZE, bubbleGroups, commentMode, layoutCards } from '../src/render/ui/comment-layout.js';

describe('layoutCards', () => {
  it('stacks resting cards in anchor order without overlap', () => {
    const tops = layoutCards(
      [
        { id: 'a', anchorY: 50, height: 80 },
        { id: 'b', anchorY: 60, height: 60 },
        { id: 'c', anchorY: 200, height: 40 },
      ],
      undefined,
      8,
    );
    expect(tops.get('a')).toBe(50);
    expect(tops.get('b')).toBe(138);
    expect(tops.get('c')).toBe(206);
  });

  it('keeps the open card at its anchor when there is room', () => {
    const tops = layoutCards(
      [
        { id: 'a', anchorY: 100, height: 40 },
        { id: 'b', anchorY: 200, height: 50 },
        { id: 'c', anchorY: 300, height: 40 },
      ],
      'b',
      8,
    );
    expect(tops.get('a')).toBe(100);
    expect(tops.get('b')).toBe(200);
    expect(tops.get('c')).toBe(300);
  });

  it('pushes close neighbours away from the open card', () => {
    const tops = layoutCards(
      [
        { id: 'a', anchorY: 100, height: 40 },
        { id: 'b', anchorY: 110, height: 50 },
        { id: 'c', anchorY: 120, height: 40 },
      ],
      'b',
      8,
    );
    expect(tops.get('a')).toBe(62);
    expect(tops.get('b')).toBe(110);
    expect(tops.get('c')).toBe(168);
  });

  it('shifts the open card and cards above down when there is no room above', () => {
    const tops = layoutCards(
      [
        { id: 'a', anchorY: 50, height: 80 },
        { id: 'b', anchorY: 60, height: 60 },
        { id: 'c', anchorY: 200, height: 40 },
      ],
      'b',
      8,
    );
    expect(tops.get('a')).toBe(0);
    expect(tops.get('b')).toBe(88);
    expect(tops.get('c')).toBe(200);
  });

  it('opens the 15th of 30 cards stacked at one anchor without overlap', () => {
    const cards = Array.from({ length: 30 }, (_, i) => ({ id: `c${i}`, anchorY: 300, height: 40 }));
    const tops = layoutCards(cards, 'c14', 8);
    expect(tops.get('c14')).toBe(672);
    for (const card of cards) expect(tops.get(card.id)!).toBeGreaterThanOrEqual(0);
    for (let i = 0; i + 1 < cards.length; i++) {
      const current = cards[i]!;
      expect(tops.get(cards[i + 1]!.id)!).toBeGreaterThanOrEqual(tops.get(current.id)! + current.height + 8);
    }
  });
});

describe('commentMode', () => {
  it('uses the plain threshold with no previous mode', () => {
    // page 612 -> threshold 0.75 * 612 + 300 = 759
    expect(commentMode(760, 612, undefined)).toBe('cards');
    expect(commentMode(758, 612, undefined)).toBe('bubbles');
  });

  it('applies hysteresis around the threshold', () => {
    expect(commentMode(780, 612, 'bubbles')).toBe('bubbles');
    expect(commentMode(800, 612, 'bubbles')).toBe('cards');
    expect(commentMode(730, 612, 'cards')).toBe('cards');
    expect(commentMode(718, 612, 'cards')).toBe('bubbles');
  });
});

describe('bubbleGroups', () => {
  it('keeps neighbouring 28px bubbles from overlapping by default (at least 30 apart)', () => {
    const groups = bubbleGroups([
      { id: 'a', y: 0 },
      { id: 'b', y: 25 },
      { id: 'c', y: 30 },
      { id: 'd', y: 59 },
      { id: 'e', y: 60 },
    ]);
    expect(groups).toEqual([
      { y: 0, ids: ['a', 'b'] },
      { y: 30, ids: ['c', 'd'] },
      { y: 60, ids: ['e'] },
    ]);
    for (let i = 1; i < groups.length; i++) expect(groups[i]!.y - groups[i - 1]!.y).toBeGreaterThanOrEqual(BUBBLE_SIZE + 2);
  });

  it('compares each anchor to the group y, not the previous anchor', () => {
    expect(
      bubbleGroups([
        { id: 'a', y: 0 },
        { id: 'b', y: 15 },
        { id: 'c', y: 30 },
        { id: 'd', y: 45 },
      ]),
    ).toEqual([
      { y: 0, ids: ['a', 'b'] },
      { y: 30, ids: ['c', 'd'] },
    ]);
  });

  it('processes unsorted input in y order', () => {
    expect(
      bubbleGroups([
        { id: 'd', y: 45 },
        { id: 'b', y: 15 },
        { id: 'a', y: 0 },
        { id: 'c', y: 30 },
      ]),
    ).toEqual([
      { y: 0, ids: ['a', 'b'] },
      { y: 30, ids: ['c', 'd'] },
    ]);
  });

  it('merges anchors closer than the minimum gap', () => {
    expect(
      bubbleGroups(
        [
          { id: 'a', y: 10 },
          { id: 'b', y: 20 },
          { id: 'c', y: 50 },
        ],
        22,
      ),
    ).toEqual([
      { y: 10, ids: ['a', 'b'] },
      { y: 50, ids: ['c'] },
    ]);
  });
});
