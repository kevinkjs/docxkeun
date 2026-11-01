/**
 * Navigation targets: headings, bookmarks (anchors) and hyperlinks, located on the laid-out
 * pages, so a viewer can list them and jump or step between them.
 */
import type { BodyLine, LayoutResult, LinkArea, Page, Rect } from '../layout/types.js';
import type { LinkTarget } from '../model/fields.js';

export type NavKind = 'heading' | 'bookmark' | 'link';

/** A place in the laid-out document (points from the page's top-left corner). */
export interface DocLocation {
  page: number;
  x: number;
  y: number;
}

export interface NavTarget extends DocLocation {
  kind: NavKind;
  /** Heading text, bookmark name, or the link's text. */
  label: string;
  /** Areas on `page` to highlight when jumping here. */
  rects: Rect[];
  /** Headings: outline level (0 = Heading 1). */
  level?: number;
  /** Bookmarks: the bookmark name. */
  name?: string;
  /** Bookmarks whose name starts with "_" (_Toc, _Hlk, _GoBack); Word hides these. */
  hidden?: boolean;
  /** Links: where the link points. */
  link?: LinkTarget;
  /** Internal links: the location of the bookmark they point to. */
  destination?: DocLocation;
}

export interface NavTargets {
  headings: NavTarget[];
  bookmarks: NavTarget[];
  links: NavTarget[];
}

export function compareLocations(a: DocLocation, b: DocLocation): number {
  return a.page - b.page || a.y - b.y || a.x - b.x;
}

function lineAt(page: Page, x: number, y: number): BodyLine | undefined {
  return page.lines.find((l) => y >= l.y - 0.5 && y < l.y + l.h && x >= l.x - 1 && x <= l.x + l.w + 1) ?? page.lines.find((l) => y >= l.y - 0.5 && y < l.y + l.h);
}

/** Visible text inside a rectangle of a page. */
export function textIn(page: Page, rect: Rect): string {
  let out = '';
  for (const item of page.items) {
    if (item.type !== 'text' || item.y < rect.y || item.y > rect.y + rect.h + 0.5) continue;
    item.chars.forEach((ch, i) => {
      const mid = item.xs[i]! + item.widths[i]! / 2;
      if (mid >= rect.x && mid <= rect.x + rect.w) out += ch;
    });
  }
  return out;
}

function isBody(page: Page, r: Rect): boolean {
  const mid = r.y + r.h / 2;
  return mid >= page.content.y && mid <= page.content.y + page.content.h;
}

export function collectNavTargets(layout: LayoutResult): NavTargets {
  const headings: NavTarget[] = [];
  const bookmarks: NavTarget[] = [];
  const links: NavTarget[] = [];

  for (const page of layout.pages) {
    for (const h of page.headings) {
      const line = lineAt(page, h.x, h.y + 0.5);
      headings.push({
        kind: 'heading',
        label: h.title,
        level: h.level,
        page: page.index,
        x: h.x,
        y: h.y,
        rects: [{ x: line?.x ?? h.x, y: h.y, w: line?.w ?? 200, h: line?.h ?? 14 }],
      });
    }

    // Link areas are recorded per line; consecutive areas of the same link are one target.
    let current: { target: LinkTarget; rects: LinkArea[] } | undefined;
    const flush = () => {
      if (!current) return;
      const [first] = current.rects;
      const label = current.rects.map((r) => textIn(page, r)).join(' ').replace(/\s+/g, ' ').trim();
      const anchor = current.target.anchor ? layout.anchors.get(current.target.anchor) : undefined;
      links.push({
        kind: 'link',
        label: label || current.target.url || (current.target.anchor ? `#${current.target.anchor}` : 'Link'),
        page: page.index,
        x: first!.x,
        y: first!.y,
        rects: current.rects.map(({ x, y, w, h }) => ({ x, y, w, h })),
        link: current.target,
        destination: anchor ? { page: anchor.page, x: anchor.x, y: anchor.y } : undefined,
      });
      current = undefined;
    };
    for (const area of page.links) {
      // Links in headers and footers repeat on every page; leave them out.
      if (!isBody(page, area)) continue;
      if (current && current.target === area.target) current.rects.push(area);
      else {
        flush();
        current = { target: area.target, rects: [area] };
      }
    }
    flush();
  }

  for (const [name, a] of layout.anchors) {
    if (!name) continue;
    const page = layout.pages[a.page];
    const line = page ? lineAt(page, a.x, a.y + 0.5) : undefined;
    bookmarks.push({
      kind: 'bookmark',
      label: name,
      name,
      hidden: name.startsWith('_'),
      page: a.page,
      x: a.x,
      y: a.y,
      rects: [{ x: a.x, y: line?.y ?? a.y, w: Math.max(12, line ? line.x + line.w - a.x : 120), h: line?.h ?? 14 }],
    });
  }

  for (const list of [headings, bookmarks, links]) list.sort(compareLocations);
  return { headings, bookmarks, links };
}

/** The first target after (dir 1) or the last target before (dir -1) a location. */
export function nextTarget(list: NavTarget[], from: DocLocation, dir: 1 | -1, tolerance = 1): NavTarget | undefined {
  if (dir > 0) return list.find((t) => t.page > from.page || (t.page === from.page && t.y > from.y + tolerance));
  for (let i = list.length - 1; i >= 0; i--) {
    const t = list[i]!;
    if (t.page < from.page || (t.page === from.page && t.y < from.y - tolerance)) return t;
  }
  return undefined;
}
