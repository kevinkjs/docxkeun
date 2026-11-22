/**
 * Display list → SVG markup. Text stays real text (selectable, searchable, accessible), with an
 * explicit x for every character so the browser draws exactly what layout computed.
 * All document-derived strings are escaped: documents are untrusted input.
 */
import { pngDataUrl } from '../images/png-encode.js';
import type { DisplayItem, LayoutFont, Page, PathItem, RasterImage, TextItem } from '../layout/types.js';

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
  if (t.complex && t.chars.length) {
    // Right-to-left and complex scripts: the browser shapes and orders the run; textLength keeps our width.
    const left = t.xs[0]!;
    const right = t.xs[t.xs.length - 1]! + t.widths[t.widths.length - 1]!;
    const dir = t.complex.rtl ? ` direction="rtl" unicode-bidi="embed"` : '';
    return `<text data-i="${index}" data-complex="1" x="${n(t.complex.rtl ? right : left)}" y="${n(t.y)}" textLength="${n(right - left)}" lengthAdjust="spacingAndGlyphs"${dir} ${common}>${esc(t.complex.text)}</text>`;
  }
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

export function pathData(commands: PathItem['commands']): string {
  return commands
    .map((c) => (c.op === 'Z' ? 'Z' : c.op === 'C' ? `C${n(c.x1)} ${n(c.y1)} ${n(c.x2)} ${n(c.y2)} ${n(c.x)} ${n(c.y)}` : `${c.op}${n(c.x)} ${n(c.y)}`))
    .join('');
}

function pathMarkup(item: PathItem): string {
  const attrs = [`d="${pathData(item.commands)}"`];
  attrs.push(item.fill ? `fill="${color(item.fill)}"` : 'fill="none"');
  if (item.fill && item.fillOpacity !== undefined && item.fillOpacity < 1) attrs.push(`fill-opacity="${n(item.fillOpacity)}"`);
  if (item.evenOdd) attrs.push('fill-rule="evenodd"');
  if (item.stroke) {
    attrs.push(`stroke="${color(item.stroke)}" stroke-width="${n(item.strokeWidth ?? 0.75)}"`);
    if (item.strokeOpacity !== undefined && item.strokeOpacity < 1) attrs.push(`stroke-opacity="${n(item.strokeOpacity)}"`);
    if (item.dash?.length) attrs.push(`stroke-dasharray="${item.dash.map(n).join(' ')}"`);
    if (item.lineCap) attrs.push(`stroke-linecap="${item.lineCap}"`);
    if (item.lineJoin) attrs.push(`stroke-linejoin="${item.lineJoin}"`);
  }
  return `<path ${attrs.join(' ')}/>`;
}

/** Decoded pixels become a PNG data URL, cached per image. */
const rasterUrls = new WeakMap<RasterImage, string>();
function rasterUrl(image: RasterImage): string {
  let url = rasterUrls.get(image);
  if (!url) {
    url = pngDataUrl(image);
    rasterUrls.set(image, url);
  }
  return url;
}

/** SVG markup for a list of display items (without the <svg> wrapper). */
export function itemsToSvg(items: DisplayItem[], o: SvgOptions, indexBase = 0): string {
  const parts: string[] = [];
  items.forEach((item, i) => {
    switch (item.type) {
      case 'text':
        parts.push(textMarkup(item, indexBase + i, o));
        break;
      case 'rect':
        parts.push(`<rect x="${n(item.x)}" y="${n(item.y)}" width="${n(item.w)}" height="${n(item.h)}" fill="${color(item.fill)}"/>`);
        break;
      case 'line':
        parts.push(lineMarkup(item));
        break;
      case 'path':
        parts.push(pathMarkup(item));
        break;
      case 'image': {
        if (item.frameOnly) break;
        const url = item.raster ? rasterUrl(item.raster) : o.imageUrl(item.target, item.external);
        if (url) {
          const cx = item.x + item.w / 2;
          const cy = item.y + item.h / 2;
          const turns = [item.rotation ? `rotate(${n(item.rotation)} ${n(cx)} ${n(cy)})` : '', item.flipH || item.flipV ? `translate(${n(cx)} ${n(cy)}) scale(${item.flipH ? -1 : 1} ${item.flipV ? -1 : 1}) translate(${n(-cx)} ${n(-cy)})` : ''].filter(Boolean).join(' ');
          parts.push(`<image href="${esc(url)}" x="${n(item.x)}" y="${n(item.y)}" width="${n(item.w)}" height="${n(item.h)}" preserveAspectRatio="none"${turns ? ` transform="${turns}"` : ''}${item.alt ? ` aria-label="${esc(item.alt)}"` : ''}/>`);
        } else {
          parts.push(`<rect x="${n(item.x)}" y="${n(item.y)}" width="${n(item.w)}" height="${n(item.h)}" fill="#f1f3f4" stroke="#c0c4c8" stroke-width="0.5"/>`);
        }
        break;
      }
    }
  });
  return parts.join('');
}

export function pageToSvg(page: Page, o: SvgOptions): string {
  const parts: string[] = [];
  parts.push(`<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 ${n(page.width)} ${n(page.height)}" width="100%" height="100%" style="white-space:pre" role="document">`);
  parts.push(`<rect x="0" y="0" width="${n(page.width)}" height="${n(page.height)}" fill="#fff"/>`);
  parts.push(itemsToSvg(page.items, o));
  parts.push('</svg>');
  return parts.join('');
}
