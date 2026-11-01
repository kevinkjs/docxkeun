/**
 * Display list → SVG markup. Text stays real text (selectable, searchable, accessible), with an
 * explicit x for every character so the browser draws exactly what layout computed.
 * All document-derived strings are escaped: documents are untrusted input.
 */
import type { DisplayItem, LayoutFont, Page, TextItem } from '../layout/types.js';

export interface SvgOptions {
  /** CSS font-family for a layout font (registered FontFace name, or a system family list). */
  fontFamily(font: LayoutFont): string;
  /** URL for an image part, or undefined when it is not loaded / not displayable. */
  imageUrl(target: string | undefined, external: boolean): string | undefined;
}

export function esc(s: string): string {
  return s.replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' })[c]!);
}

function n(v: number): string {
  return Number.isInteger(v) ? String(v) : v.toFixed(2).replace(/\.?0+$/, '');
}

function color(c: string): string {
  return /^[0-9a-f]{6}$/i.test(c) ? `#${c}` : '#000';
}

function isSimpleText(t: TextItem): boolean {
  return t.chars.every((c) => c.length === 1);
}

function textMarkup(t: TextItem, index: number, o: SvgOptions): string {
  const family = esc(o.fontFamily(t.font));
  const weight = t.font.bold && !t.font.face ? ' font-weight="bold"' : '';
  const style = t.font.italic && !t.font.face ? ' font-style="italic"' : '';
  const stroke = t.font.syntheticBold ? ` stroke="${color(t.color)}" stroke-width="${n(t.size * 0.03)}"` : '';
  const skew = t.font.syntheticItalic ? ` transform="translate(${n(t.x)} ${n(t.y)}) skewX(-12) translate(${n(-t.x)} ${n(-t.y)})"` : '';
  // xml:space="preserve" keeps leading/trailing spaces, so the per-character x list lines up.
  const common = `xml:space="preserve" font-family="${family}" font-size="${n(t.size)}" fill="${color(t.color)}"${weight}${style}${stroke}${skew}`;
  if (isSimpleText(t)) {
    return `<text data-i="${index}" x="${t.xs.map(n).join(' ')}" y="${n(t.y)}" ${common}>${esc(t.text)}</text>`;
  }
  // Clusters longer than one UTF-16 unit cannot use a per-character x list: one element each.
  return t.chars.map((c, ci) => `<text data-i="${index}" data-c="${ci}" x="${n(t.xs[ci]!)}" y="${n(t.y)}" ${common}>${esc(c)}</text>`).join('');
}

function lineMarkup(item: Extract<DisplayItem, { type: 'line' }>): string {
  const dash = item.style === 'dashed' ? ` stroke-dasharray="${n(item.width * 3)} ${n(item.width * 2)}"` : item.style === 'dotted' ? ` stroke-dasharray="${n(item.width)} ${n(item.width)}"` : '';
  const base = (x1: number, y1: number, x2: number, y2: number, w: number) =>
    `<line x1="${n(x1)}" y1="${n(y1)}" x2="${n(x2)}" y2="${n(y2)}" stroke="${color(item.color)}" stroke-width="${n(w)}"${dash}/>`;
  if (item.style !== 'double') return base(item.x1, item.y1, item.x2, item.y2, item.width);
  const horizontal = Math.abs(item.y2 - item.y1) < Math.abs(item.x2 - item.x1);
  const w = Math.max(item.width / 3, 0.25);
  const d = item.width / 2 - w / 2 + w;
  return horizontal
    ? base(item.x1, item.y1 - d / 2, item.x2, item.y2 - d / 2, w) + base(item.x1, item.y1 + d / 2, item.x2, item.y2 + d / 2, w)
    : base(item.x1 - d / 2, item.y1, item.x2 - d / 2, item.y2, w) + base(item.x1 + d / 2, item.y1, item.x2 + d / 2, item.y2, w);
}

export function pageToSvg(page: Page, o: SvgOptions): string {
  const parts: string[] = [];
  parts.push(`<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 ${n(page.width)} ${n(page.height)}" width="100%" height="100%" style="white-space:pre" role="document">`);
  parts.push(`<rect x="0" y="0" width="${n(page.width)}" height="${n(page.height)}" fill="#fff"/>`);
  page.items.forEach((item, i) => {
    switch (item.type) {
      case 'text':
        parts.push(textMarkup(item, i, o));
        break;
      case 'rect':
        parts.push(`<rect x="${n(item.x)}" y="${n(item.y)}" width="${n(item.w)}" height="${n(item.h)}" fill="${color(item.fill)}"/>`);
        break;
      case 'line':
        parts.push(lineMarkup(item));
        break;
      case 'image': {
        const url = o.imageUrl(item.target, item.external);
        if (url) {
          parts.push(`<image href="${esc(url)}" x="${n(item.x)}" y="${n(item.y)}" width="${n(item.w)}" height="${n(item.h)}" preserveAspectRatio="none"${item.alt ? ` aria-label="${esc(item.alt)}"` : ''}/>`);
        } else {
          parts.push(`<rect x="${n(item.x)}" y="${n(item.y)}" width="${n(item.w)}" height="${n(item.h)}" fill="#f1f3f4" stroke="#c0c4c8" stroke-width="0.5"/>`);
        }
        break;
      }
    }
  });
  parts.push('</svg>');
  return parts.join('');
}
