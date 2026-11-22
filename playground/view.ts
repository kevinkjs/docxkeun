/** What the playground remembers between visits, and how it fits the page to the viewer. */

/**
 * Fold states remembered across visits. The toolbar keeps only its fold: it is hidden while
 * editing is off, and a remembered "hidden" would keep it away after editing is turned back on.
 * The navigation panel keeps both (the header button shows and hides it).
 */
export interface Fold {
  toolbar?: { collapsed: boolean };
  navigation?: { visible: boolean; collapsed: boolean };
}

export const FOLD_KEY = 'docxkeun-playground-chrome';

/** Storage can be missing or blocked (reading `localStorage` can throw), so it is reached through a function. */
export type GetStore = () => Pick<Storage, 'getItem' | 'setItem'>;

export function readFold(store: GetStore): Fold {
  try {
    const raw = JSON.parse(store().getItem(FOLD_KEY) ?? '{}') as { toolbar?: { collapsed?: unknown }; navigation?: { visible?: unknown; collapsed?: unknown } };
    const out: Fold = {};
    if (raw.toolbar) out.toolbar = { collapsed: raw.toolbar.collapsed === true };
    if (raw.navigation) out.navigation = { visible: raw.navigation.visible !== false, collapsed: raw.navigation.collapsed === true };
    return out;
  } catch {
    return {};
  }
}

export function writeFold(store: GetStore, part: 'toolbar' | 'navigation', visible: boolean, collapsed: boolean): void {
  try {
    const next = readFold(store);
    if (part === 'toolbar') next.toolbar = { collapsed };
    else next.navigation = { visible, collapsed };
    store().setItem(FOLD_KEY, JSON.stringify(next));
  } catch {
    // Storage is blocked or full: the folds are simply not remembered.
  }
}

/**
 * The zoom that fits the widest page and the comment column (`commentsWidth`, as the viewer sized
 * it) into a viewer `viewerWidth` CSS px wide: never above 100%, never below 50%.
 */
export function fitZoom(viewerWidth: number, pageWidthsPt: number[], commentsWidth: number): number {
  if (!pageWidthsPt.length || viewerWidth <= 0) return 1;
  const pageWidth = (Math.max(...pageWidthsPt) * 96) / 72;
  // Row padding (32) and a scrollbar (16).
  const room = viewerWidth - 32 - commentsWidth - 16;
  return Math.max(0.5, Math.min(1, room / pageWidth));
}
