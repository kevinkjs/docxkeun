/** Pure helpers behind the navigation panel: searching the lists and finding the current heading. */
import { compareLocations, type DocLocation, type NavTarget } from '../navigation.js';

/** Lower-case and strip accents, so "é" matches "e" and "Pay" matches "payment". */
export function fold(text: string): string {
  return text.normalize('NFD').replace(/\p{M}+/gu, '').toLowerCase();
}

/** Targets whose label (or, for links, URL or anchor) contains the query, ignoring case and accents. */
export function filterTargets(list: NavTarget[], query: string): NavTarget[] {
  const q = fold(query.trim());
  if (!q) return list;
  return list.filter((t) => {
    if (fold(t.label).includes(q)) return true;
    const url = t.link?.url;
    const anchor = t.link?.anchor;
    return (!!url && fold(url).includes(q)) || (!!anchor && fold(anchor).includes(q));
  });
}

/**
 * The last heading at or above a location (`tolerance` points below it still count, so a heading
 * that a jump left just under the viewport top is the current one). `headings` is in document order.
 */
export function currentHeading(headings: NavTarget[], at: DocLocation, tolerance = 0): NavTarget | undefined {
  const limit = { page: at.page, x: at.x, y: at.y + tolerance };
  let found: NavTarget | undefined;
  for (const h of headings) {
    if (compareLocations(h, limit) <= 0) found = h;
    else break;
  }
  return found;
}
