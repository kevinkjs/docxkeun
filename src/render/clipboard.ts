/**
 * Rich copy and paste: clipboard HTML (from Word, Google Docs, browsers) to paragraphs of
 * formatted runs, and a selection back to HTML. A small tolerant tokenizer is used instead of
 * the DOM so it also runs in workers and tests; scripts, styles and unknown markup are ignored.
 */
import type { TextFormat } from '../api/editor.js';
import type { ListFormat } from '../api/lists.js';
import type { DocParagraph } from '../model/content.js';

export interface PastedRun {
  text: string;
  format: TextFormat;
  link?: string;
}

export interface PastedParagraph {
  /** Paragraph style id, e.g. "Heading1". */
  style?: string;
  list?: ListFormat;
  runs: PastedRun[];
}

const NAMED: Record<string, string> = {
  amp: '&', lt: '<', gt: '>', quot: '"', apos: "'", nbsp: ' ', mdash: '—', ndash: '–', hellip: '…',
  lsquo: '‘', rsquo: '’', ldquo: '“', rdquo: '”', bull: '•', copy: '©', reg: '®', trade: '™',
  euro: '€', deg: '°', times: '×', middot: '·', laquo: '«', raquo: '»', shy: '­',
};

function decodeEntities(s: string): string {
  return s.replace(/&(#x[0-9a-f]+|#\d+|[a-z]+);?/gi, (m, e: string) => {
    if (e[0] === '#') {
      const cp = e[1] === 'x' || e[1] === 'X' ? parseInt(e.slice(2), 16) : parseInt(e.slice(1), 10);
      return Number.isFinite(cp) && cp > 0 && cp <= 0x10ffff ? String.fromCodePoint(cp) : m;
    }
    return NAMED[e.toLowerCase()] ?? m;
  });
}

type Token = { type: 'text'; text: string } | { type: 'open'; name: string; attrs: Record<string, string>; selfClosing: boolean } | { type: 'close'; name: string };

function tokenize(html: string): Token[] {
  const out: Token[] = [];
  let i = 0;
  while (i < html.length) {
    const lt = html.indexOf('<', i);
    if (lt < 0) {
      out.push({ type: 'text', text: html.slice(i) });
      break;
    }
    if (lt > i) out.push({ type: 'text', text: html.slice(i, lt) });
    if (html.startsWith('<!--', lt)) {
      const end = html.indexOf('-->', lt + 4);
      i = end < 0 ? html.length : end + 3;
      continue;
    }
    if (html[lt + 1] === '!' || html[lt + 1] === '?') {
      const end = html.indexOf('>', lt);
      i = end < 0 ? html.length : end + 1;
      continue;
    }
    const m = /^<\/?\s*([a-zA-Z][\w:.-]*)/.exec(html.slice(lt, lt + 64));
    if (!m) {
      out.push({ type: 'text', text: '<' });
      i = lt + 1;
      continue;
    }
    // Find the tag end, skipping quoted attribute values.
    let j = lt + m[0].length;
    let quote = '';
    for (; j < html.length; j++) {
      const c = html[j]!;
      if (quote) {
        if (c === quote) quote = '';
      } else if (c === '"' || c === "'") quote = c;
      else if (c === '>') break;
    }
    const body = html.slice(lt + m[0].length, j);
    const name = m[1]!.toLowerCase();
    if (html[lt + 1] === '/') out.push({ type: 'close', name });
    else {
      const attrs: Record<string, string> = {};
      for (const a of body.matchAll(/([^\s=/>]+)(?:\s*=\s*("([^"]*)"|'([^']*)'|([^\s>]+)))?/g)) {
        attrs[a[1]!.toLowerCase()] = decodeEntities(a[3] ?? a[4] ?? a[5] ?? '');
      }
      out.push({ type: 'open', name, attrs, selfClosing: /\/\s*$/.test(body) });
    }
    i = j + 1;
    // Raw text elements: skip their content.
    if (html[lt + 1] !== '/' && (name === 'script' || name === 'style' || name === 'title' || name === 'xml')) {
      const close = html.toLowerCase().indexOf(`</${name}`, i);
      i = close < 0 ? html.length : close;
    }
  }
  return out;
}

function parseStyle(style: string | undefined): Record<string, string> {
  const out: Record<string, string> = {};
  for (const decl of (style ?? '').split(';')) {
    const k = decl.indexOf(':');
    if (k > 0) out[decl.slice(0, k).trim().toLowerCase()] = decl.slice(k + 1).trim();
  }
  return out;
}

/** CSS colour → RRGGBB (hex, rgb(), a few names); undefined for black/auto/unknown. */
export function cssColor(value: string | undefined): string | undefined {
  if (!value) return undefined;
  const v = value.trim().toLowerCase().replace(/\s*!important$/, '');
  let hex: string | undefined;
  const h = /^#([0-9a-f]{3}|[0-9a-f]{6})$/.exec(v);
  if (h) hex = h[1]!.length === 3 ? [...h[1]!].map((c) => c + c).join('') : h[1]!;
  const rgb = /^rgba?\(\s*(\d+)[\s,]+(\d+)[\s,]+(\d+)(?:[\s,/]+([\d.]+%?))?\s*\)$/.exec(v);
  if (rgb) {
    if (rgb[4] !== undefined && parseFloat(rgb[4]) === 0) return undefined;
    hex = [rgb[1], rgb[2], rgb[3]].map((n) => Math.min(255, Number(n)).toString(16).padStart(2, '0')).join('');
  }
  const names: Record<string, string> = { red: 'ff0000', blue: '0000ff', green: '008000', yellow: 'ffff00', orange: 'ffa500', purple: '800080', gray: '808080', grey: '808080', white: 'ffffff', black: '000000' };
  hex ??= names[v];
  return hex ? hex.toUpperCase() : undefined;
}

/** Word's highlight colours, nearest by RGB distance (only for fairly saturated backgrounds). */
function highlightFor(hex: string | undefined): string | undefined {
  if (!hex || hex === 'FFFFFF') return undefined;
  const table: Record<string, string> = { yellow: 'FFFF00', green: '00FF00', cyan: '00FFFF', magenta: 'FF00FF', blue: '0000FF', red: 'FF0000', lightGray: 'C0C0C0', darkGray: '808080' };
  const rgb = (x: string) => [0, 2, 4].map((i) => parseInt(x.slice(i, i + 2), 16));
  const [r, g, b] = rgb(hex);
  let best: string | undefined;
  let bestD = Infinity;
  for (const [name, value] of Object.entries(table)) {
    const [r2, g2, b2] = rgb(value);
    const d = (r! - r2!) ** 2 + (g! - g2!) ** 2 + (b! - b2!) ** 2;
    if (d < bestD) [best, bestD] = [name, d];
  }
  return bestD < 120 ** 2 ? best : undefined;
}

function cssSize(value: string | undefined): number | undefined {
  const m = /^([\d.]+)\s*(pt|px)?$/.exec((value ?? '').trim().toLowerCase());
  if (!m) return undefined;
  const n = parseFloat(m[1]!) * (m[2] === 'px' ? 0.75 : 1);
  return n >= 1 && n <= 400 ? Math.round(n * 2) / 2 : undefined;
}

const BLOCKS = new Set(['p', 'div', 'h1', 'h2', 'h3', 'h4', 'h5', 'h6', 'li', 'blockquote', 'pre', 'tr', 'table', 'ul', 'ol', 'section', 'article', 'header', 'footer', 'dd', 'dt', 'figcaption', 'hr']);
const VOID = new Set(['br', 'img', 'hr', 'meta', 'link', 'input', 'col', 'wbr', 'area', 'base', 'source']);

interface Frame {
  name: string;
  format: TextFormat;
  link?: string;
  pre: boolean;
}

/** Paragraphs of formatted runs from clipboard HTML. */
export function htmlToParagraphs(html: string): PastedParagraph[] {
  // Only the copied fragment when the clipboard marks it.
  const start = html.indexOf('<!--StartFragment-->');
  const end = html.indexOf('<!--EndFragment-->');
  if (start >= 0 && end > start) html = html.slice(start + 20, end);
  const paragraphs: PastedParagraph[] = [];
  let current: PastedParagraph | undefined;
  const lists: ListFormat[] = [];
  const stack: Frame[] = [{ name: '#root', format: {}, pre: false }];
  const top = () => stack[stack.length - 1]!;
  let cellCount = 0;

  const startParagraph = (p: Partial<PastedParagraph> = {}) => {
    finishParagraph();
    current = { runs: [], ...p };
  };
  const finishParagraph = () => {
    if (!current) return;
    // Trim the paragraph's outer spaces (HTML whitespace is not content).
    const runs = current.runs;
    while (runs.length && !runs[0]!.text.replace(/^[ \t]+/, '')) runs.shift();
    if (runs[0]) runs[0].text = runs[0].text.replace(/^[ ]+/, '');
    while (runs.length && !runs[runs.length - 1]!.text.replace(/[ \t]+$/, '')) runs.pop();
    if (runs.length) runs[runs.length - 1]!.text = runs[runs.length - 1]!.text.replace(/[ ]+$/, '');
    if (runs.length || current.list || current.style) paragraphs.push(current);
    current = undefined;
  };
  const addText = (raw: string, literal = false) => {
    const f = top();
    let text = literal ? raw : decodeEntities(raw);
    if (!f.pre && !literal) {
      text = text.replace(/[\s ]+/g, (ws) => (ws.includes(' ') ? ws.replace(/[^ ]/g, '').replace(/ /g, ' ') || ' ' : ' '));
      if (!current && !text.trim()) return;
    }
    if (!current) current = { runs: [] };
    const last = current.runs[current.runs.length - 1];
    // Collapse spaces across tags.
    if (!f.pre && !literal && last && /[ \n\t]$/.test(last.text)) text = text.replace(/^ /, '');
    if (!text) return;
    if (last && last.link === f.link && JSON.stringify(last.format) === JSON.stringify(f.format)) last.text += text;
    else current.runs.push({ text, format: { ...f.format }, link: f.link });
  };

  for (const t of tokenize(html)) {
    if (t.type === 'text') {
      addText(t.text);
      continue;
    }
    if (t.type === 'close') {
      const i = stack.map((f) => f.name).lastIndexOf(t.name);
      if (i > 0) stack.length = i;
      if (t.name === 'ul' || t.name === 'ol') lists.pop();
      if (BLOCKS.has(t.name)) finishParagraph();
      continue;
    }
    const name = t.name;
    const attrs = t.attrs;
    const css = parseStyle(attrs.style);
    if (css.display === 'none' || css['mso-hide'] === 'all') {
      if (!VOID.has(name) && !t.selfClosing) stack.push({ ...top(), name, format: { ...top().format, hidden: true } });
      continue;
    }
    if (name === 'br') {
      const last = current?.runs[current.runs.length - 1];
      if (last) last.text = last.text.replace(/ $/, '');
      addText('\n', true);
      continue;
    }
    if (name === 'ul' || name === 'ol') {
      finishParagraph();
      lists.push(name === 'ul' ? 'bullet' : 'decimal');
    }
    const heading = /^h([1-6])$/.exec(name);
    if (heading) startParagraph({ style: `Heading${heading[1]}` });
    else if (name === 'li') startParagraph({ list: lists[lists.length - 1] ?? 'bullet' });
    else if (name === 'tr') {
      startParagraph();
      cellCount = 0;
    } else if (name === 'td' || name === 'th') {
      if (cellCount++ > 0) addText('\t', true);
    } else if (BLOCKS.has(name)) startParagraph();
    // Word marks its own list paragraphs with a class and an mso-list style instead of <li>.
    if (name === 'p' && (/MsoListParagraph/i.test(attrs.class ?? '') || css['mso-list']) && current) current.list = 'bullet';

    if (VOID.has(name) || t.selfClosing) continue;
    const parent = top();
    const format: TextFormat = { ...parent.format };
    let link = parent.link;
    switch (name) {
      case 'b':
      case 'strong':
        format.bold = true;
        break;
      case 'i':
      case 'em':
      case 'cite':
        format.italic = true;
        break;
      case 'u':
      case 'ins':
        format.underline = true;
        break;
      case 's':
      case 'strike':
      case 'del':
        format.strike = true;
        break;
      case 'sup':
        format.vertAlign = 'superscript';
        break;
      case 'sub':
        format.vertAlign = 'subscript';
        break;
      case 'code':
      case 'kbd':
      case 'tt':
        format.font = 'Courier New';
        break;
      case 'a':
        if (attrs.href && /^(https?|mailto|tel):/i.test(attrs.href)) link = attrs.href;
        break;
      case 'font':
        if (attrs.color) format.color = cssColor(attrs.color);
        if (attrs.face) format.font = attrs.face.split(',')[0]!.replace(/['"]/g, '').trim();
        break;
    }
    // Google Docs wraps everything in <b style="font-weight:normal">: CSS decides.
    const weight = css['font-weight'];
    if (weight) format.bold = weight === 'bold' || weight === 'bolder' || Number(weight) >= 600;
    if (css['font-style']) format.italic = css['font-style'] === 'italic' || css['font-style'] === 'oblique';
    const deco = css['text-decoration'] ?? css['text-decoration-line'];
    if (deco) {
      if (deco.includes('underline')) format.underline = true;
      if (deco.includes('line-through')) format.strike = true;
      if (deco === 'none' && name !== 'a') format.underline = false;
    }
    const color = cssColor(css.color);
    if (color && color !== '000000') format.color = color;
    const highlight = highlightFor(cssColor(css['background-color'] ?? css.background));
    if (highlight) format.highlight = highlight;
    const size = cssSize(css['font-size']);
    if (size) format.size = size;
    const family = css['font-family']?.split(',')[0]?.replace(/['"]/g, '').trim();
    if (family && !/^(inherit|initial|sans-serif|serif|monospace|system-ui|-apple-system)$/i.test(family)) format.font = family;
    if (css['vertical-align'] === 'super') format.vertAlign = 'superscript';
    if (css['vertical-align'] === 'sub') format.vertAlign = 'subscript';
    stack.push({ name, format, link, pre: parent.pre || name === 'pre' || /pre/.test(css['white-space'] ?? '') });
  }
  finishParagraph();
  // Drop hidden runs and formatting that only repeats the default.
  for (const p of paragraphs) {
    p.runs = p.runs.filter((r) => !r.format.hidden);
    for (const r of p.runs) for (const k of Object.keys(r.format) as Array<keyof TextFormat>) if (r.format[k] === false || r.format[k] === undefined) delete r.format[k];
  }
  return paragraphs.filter((p) => p.runs.length || p.list || p.style);
}

const escapeHtml = (s: string) => s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');

/** HTML for parts of paragraphs (`from`/`to` are offsets in their text with objects as U+FFFC). */
export function paragraphsToHtml(parts: Array<{ para: DocParagraph; from: number; to: number }>, listKind: (p: DocParagraph) => 'bullet' | 'number' | undefined): string {
  let html = '';
  let openList: 'ul' | 'ol' | undefined;
  for (const { para, from, to } of parts) {
    let inner = '';
    let offset = 0;
    for (const inline of para.inlines) {
      let text: string | undefined;
      if (inline.kind === 'text') text = inline.text;
      else if (inline.kind === 'tab') text = '\t';
      else if (inline.kind === 'break' && inline.type === 'line') text = '\n';
      else if (inline.kind === 'image' && !inline.float) text = '￼';
      if (text === undefined) continue;
      const a = Math.max(from, offset);
      const b = Math.min(to, offset + text.length);
      offset += text.length;
      if (a >= b) continue;
      const piece = text.slice(a - (offset - text.length), b - (offset - text.length));
      if (piece === '￼') continue;
      let h = escapeHtml(piece).replace(/\n/g, '<br>');
      if (inline.kind === 'text') {
        const pr = inline.props;
        const css: string[] = [];
        if (pr.color && pr.color !== 'auto') css.push(`color:#${pr.color}`);
        if (pr.size) css.push(`font-size:${pr.size}pt`);
        if (pr.fonts?.ascii) css.push(`font-family:'${pr.fonts.ascii.replace(/'/g, '')}'`);
        if (css.length) h = `<span style="${css.join(';')}">${h}</span>`;
        if (pr.bold) h = `<b>${h}</b>`;
        if (pr.italic) h = `<i>${h}</i>`;
        if (pr.underline && pr.underline !== 'none') h = `<u>${h}</u>`;
        if (pr.strike || pr.dstrike) h = `<s>${h}</s>`;
        if (pr.vertAlign === 'superscript') h = `<sup>${h}</sup>`;
        if (pr.vertAlign === 'subscript') h = `<sub>${h}</sub>`;
      }
      inner += h;
    }
    const kind = listKind(para);
    const tag = kind ? 'ul' : undefined;
    const want = kind === 'bullet' ? 'ul' : kind === 'number' ? 'ol' : undefined;
    if (openList && openList !== want) {
      html += `</${openList}>`;
      openList = undefined;
    }
    if (want && !openList) {
      html += `<${want}>`;
      openList = want;
    }
    const heading = /^Heading([1-6])$/.exec(para.styleId ?? '');
    const block = tag ? 'li' : heading ? `h${heading[1]}` : 'p';
    const dir = para.props.bidi ? ' dir="rtl"' : '';
    html += `<${block}${dir}>${inner || '<br>'}</${block}>`;
  }
  if (openList) html += `</${openList}>`;
  return html;
}
