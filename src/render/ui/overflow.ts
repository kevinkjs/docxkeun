/**
 * How many leading groups stay in a toolbar row of `available` pixels; the rest go to the "more"
 * menu. When everything fits, no "more" button is needed and `moreButtonWidth` is not reserved.
 */
export function fitGroups(groupWidths: number[], available: number, moreButtonWidth: number): number {
  const total = groupWidths.reduce((a, b) => a + b, 0);
  if (total <= available) return groupWidths.length;
  let used = moreButtonWidth;
  let count = 0;
  for (const w of groupWidths) {
    if (used + w > available) break;
    used += w;
    count++;
  }
  return count;
}
