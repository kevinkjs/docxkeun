/** Inline 16x16 stroke icons shared by the toolbar, navigation panel and context menu. */
export type IconName =
  | 'undo' | 'redo' | 'alignLeft' | 'alignCenter' | 'alignRight' | 'justify' | 'bullets' | 'numbers'
  | 'outdent' | 'indent' | 'ltr' | 'rtl' | 'track' | 'fold' | 'unfold' | 'outline' | 'bookmark' | 'link'
  | 'comment' | 'search' | 'back' | 'forward' | 'cut' | 'copy' | 'paste' | 'more';

const PATHS: Record<IconName, string> = {
  undo: '<path d="M5 4L2 7l3 3"/><path d="M2 7h8a4 4 0 010 8H7"/>',
  redo: '<path d="M11 4l3 3-3 3"/><path d="M14 7H6a4 4 0 000 8h3"/>',
  alignLeft: '<path d="M2 3h12M2 6.5h8M2 10h12M2 13.5h8"/>',
  alignCenter: '<path d="M2 3h12M4 6.5h8M2 10h12M4 13.5h8"/>',
  alignRight: '<path d="M2 3h12M6 6.5h8M2 10h12M6 13.5h8"/>',
  justify: '<path d="M2 3h12M2 6.5h12M2 10h12M2 13.5h12"/>',
  bullets: '<circle cx="3" cy="4" r=".8" fill="currentColor"/><circle cx="3" cy="8" r=".8" fill="currentColor"/><circle cx="3" cy="12" r=".8" fill="currentColor"/><path d="M6 4h8M6 8h8M6 12h8"/>',
  numbers: '<path d="M6 4h8M6 8h8M6 12h8"/><text x="1" y="5.5" font-size="5" fill="currentColor" stroke="none">1</text><text x="1" y="9.5" font-size="5" fill="currentColor" stroke="none">2</text><text x="1" y="13.5" font-size="5" fill="currentColor" stroke="none">3</text>',
  outdent: '<path d="M7 4h7M7 8h7M7 12h7M4 6L2 8l2 2"/>',
  indent: '<path d="M7 4h7M7 8h7M7 12h7M2 6l2 2-2 2"/>',
  rtl: '<path d="M6 2v9M9 2v9M6 2h5M5 2a2.5 2.5 0 000 5h1"/><path d="M14 14H3M5 12l-2 2 2 2"/>',
  ltr: '<path d="M6 2v9M9 2v9M6 2h5M5 2a2.5 2.5 0 000 5h1"/><path d="M2 14h11M11 12l2 2-2 2"/>',
  track: '<path d="M3 13l1-3 7-7 2 2-7 7z"/><path d="M2 15h12" stroke-dasharray="2 1.5"/>',
  fold: '<path d="M4 10l4-4 4 4"/>',
  unfold: '<path d="M4 6l4 4 4-4"/>',
  outline: '<path d="M2 3h12M5 6.5h9M5 10h9M8 13.5h6"/>',
  bookmark: '<path d="M4 2h8v12l-4-3-4 3z"/>',
  link: '<path d="M7 9a3 3 0 004 0l2.5-2.5a3 3 0 00-4-4L8.5 3.5"/><path d="M9 7a3 3 0 00-4 0L2.5 9.5a3 3 0 004 4l1-1"/>',
  comment: '<path d="M2 3h12v8H7l-3 3v-3H2z"/>',
  search: '<circle cx="7" cy="7" r="4.5"/><path d="M10.5 10.5L14 14"/>',
  back: '<path d="M10 3L5 8l5 5"/>',
  forward: '<path d="M6 3l5 5-5 5"/>',
  cut: '<circle cx="4.5" cy="12" r="2"/><circle cx="11.5" cy="12" r="2"/><path d="M5.5 10.3L11 2M10.5 10.3L5 2"/>',
  copy: '<rect x="5.5" y="5.5" width="8" height="8" rx="1"/><path d="M10.5 5.5v-3h-8v8h3"/>',
  paste: '<rect x="3" y="3" width="10" height="11" rx="1"/><path d="M6 3V2h4v1M6 7h4M6 10h4"/>',
  more: '<circle cx="3.5" cy="8" r=".9" fill="currentColor"/><circle cx="8" cy="8" r=".9" fill="currentColor"/><circle cx="12.5" cy="8" r=".9" fill="currentColor"/>',
};

const escapeAttr = (text: string) => text.replace(/&/g, '&amp;').replace(/"/g, '&quot;').replace(/</g, '&lt;');

/** Inline SVG markup for an icon. With a `label` it is announced; without one it is decorative. */
export function icon(name: IconName, label?: string): string {
  const a11y = label ? `role="img" aria-label="${escapeAttr(label)}"` : 'aria-hidden="true"';
  return `<svg width="16" height="16" viewBox="0 0 16 16" ${a11y} fill="none" stroke="currentColor" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round">${PATHS[name]}</svg>`;
}

/** True when `name` is one of the built-in icons (own keys only). */
export function isIconName(name: string): name is IconName {
  return Object.prototype.hasOwnProperty.call(PATHS, name);
}
