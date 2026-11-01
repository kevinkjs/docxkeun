/** Map DOM selections inside rendered pages back to positions in the document model. */
import type { Page, TextItem } from '../layout/types.js';
import type { TextInline } from '../model/content.js';

export interface InlinePosition {
  inline: TextInline;
  /** UTF-16 offset in `inline.text`. */
  offset: number;
}

function textElement(node: Node): Element | null {
  let cur: Node | null = node;
  while (cur && !(cur instanceof Element && cur.tagName.toLowerCase() === 'text' && cur.hasAttribute('data-i'))) cur = cur.parentNode;
  return cur as Element | null;
}

function pageIndex(el: Element): number | undefined {
  const pageEl = el.closest('[data-page]');
  const v = pageEl?.getAttribute('data-page');
  return v === null || v === undefined ? undefined : Number(v);
}

/** Position for a DOM (node, offset) pair, e.g. a Selection's anchor or focus. */
export function positionFromDom(node: Node, offset: number, pages: Page[]): InlinePosition | undefined {
  const el = textElement(node);
  if (!el) return undefined;
  const pi = pageIndex(el);
  const item = pi === undefined ? undefined : (pages[pi]?.items[Number(el.getAttribute('data-i'))] as TextItem | undefined);
  if (!item || item.type !== 'text' || !item.source) return undefined;
  let cluster: number;
  const fixed = el.getAttribute('data-c');
  if (fixed !== null) {
    cluster = Number(fixed) + (offset > 0 ? 1 : 0);
  } else {
    cluster = 0;
    let units = 0;
    while (cluster < item.chars.length && units + item.chars[cluster]!.length <= offset) units += item.chars[cluster++]!.length;
  }
  const { inline, offsets } = item.source;
  if (cluster < offsets.length) return { inline, offset: offsets[cluster]! };
  const lastOffset = offsets[offsets.length - 1] ?? 0;
  // Past the last character: the end of its cluster in the source text.
  const lastChar = [...inline.text.slice(lastOffset)][0] ?? '';
  return { inline, offset: lastOffset + lastChar.length };
}

/** Start and end positions of the current selection, in document order. */
export function selectionPositions(selection: Selection | null, pages: Page[]): { start: InlinePosition; end: InlinePosition } | undefined {
  if (!selection || selection.rangeCount === 0 || selection.isCollapsed) return undefined;
  const range = selection.getRangeAt(0);
  const start = positionFromDom(range.startContainer, range.startOffset, pages);
  const end = positionFromDom(range.endContainer, range.endOffset, pages);
  if (!start || !end) return undefined;
  return { start, end };
}
