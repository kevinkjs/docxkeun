/**
 * Pure layout for the viewer's comment column. Nothing here touches the DOM; all values are in
 * the same unit as the anchors (Task 7 passes CSS px).
 */

export interface CardBox {
  id: string;
  anchorY: number;
  height: number;
}

/**
 * Tops for comment cards, keyed by card id. Cards are ordered by `anchorY` (stable for ties).
 *
 * Without `openId`, each card sits at its anchor unless that would overlap the card before it.
 * With `openId`, the open card stays at its anchor; cards above are stacked upward from it and
 * cards below are stacked downward from it. If the topmost card would start above 0, the open
 * card and every card above it move down by that overflow; cards below follow the open card.
 */
export function layoutCards(cards: CardBox[], openId?: string, gap = 8): Map<string, number> {
  const sorted = [...cards].sort((a, b) => a.anchorY - b.anchorY);
  const tops = new Map<string, number>();
  const openIndex = openId === undefined ? -1 : sorted.findIndex((c) => c.id === openId);

  if (openIndex === -1) {
    let previousBottom = -Infinity;
    for (const card of sorted) {
      const top = Math.max(card.anchorY, previousBottom + gap);
      tops.set(card.id, top);
      previousBottom = top + card.height;
    }
    return tops;
  }

  const open = sorted[openIndex] as CardBox;
  // Tops for cards 0..openIndex before any shift; the open card stays at its anchor.
  const upTops: number[] = [open.anchorY];
  for (let i = openIndex - 1; i >= 0; i--) {
    const card = sorted[i] as CardBox;
    const nextTop = upTops[0] as number;
    upTops.unshift(Math.min(card.anchorY, nextTop - card.height - gap));
  }
  const overflow = Math.max(0, -(upTops[0] as number));
  upTops.forEach((top, i) => tops.set((sorted[i] as CardBox).id, top + overflow));

  let previousBottom = (open.anchorY + overflow) + open.height;
  for (let i = openIndex + 1; i < sorted.length; i++) {
    const card = sorted[i] as CardBox;
    const top = Math.max(card.anchorY, previousBottom + gap);
    tops.set(card.id, top);
    previousBottom = top + card.height;
  }
  return tops;
}

/**
 * Whether comments show as cards or as bubbles. Cards need the viewer to be at least
 * `0.75 × widestPageAt100 + 300` wide. The `previous` mode adds a 40-unit band so the mode does
 * not flicker when the width hovers near the threshold.
 */
export function commentMode(
  viewerWidth: number,
  widestPageAt100: number,
  previous: 'cards' | 'bubbles' | undefined,
): 'cards' | 'bubbles' {
  const threshold = 0.75 * widestPageAt100 + 300;
  if (previous === 'bubbles') return viewerWidth >= threshold + 40 ? 'cards' : 'bubbles';
  if (previous === 'cards') return viewerWidth < threshold - 40 ? 'bubbles' : 'cards';
  return viewerWidth >= threshold ? 'cards' : 'bubbles';
}

/** Height (and width) of a comment bubble, in CSS px. */
export const BUBBLE_SIZE = 28;

/**
 * Groups anchors into bubbles. Anchors are sorted by `y` (stable for ties). An anchor joins the
 * current group only when it is closer than `minGap` to the group's `y`; otherwise it starts a new
 * group. A group's `y` is its first anchor's `y`, which is where the bubble is drawn. The default
 * gap (a bubble's height plus 2) keeps neighbouring bubbles from overlapping.
 */
export function bubbleGroups(
  anchors: Array<{ id: string; y: number }>,
  minGap = BUBBLE_SIZE + 2,
): Array<{ y: number; ids: string[] }> {
  const sorted = [...anchors].sort((a, b) => a.y - b.y);
  const groups: Array<{ y: number; ids: string[] }> = [];
  for (const anchor of sorted) {
    const current = groups[groups.length - 1];
    if (current && anchor.y - current.y < minGap) {
      current.ids.push(anchor.id);
    } else {
      groups.push({ y: anchor.y, ids: [anchor.id] });
    }
  }
  return groups;
}
