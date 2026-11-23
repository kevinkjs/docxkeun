/**
 * Showing an AI assistant at work on the pages:
 * - a shimmer over the paragraphs it is working on, and its own named caret gliding there;
 * - a scan line while it reads;
 * - after a change, new text "inked in" left to right with a spark at the pen tip and a fading
 *   glow, removed paragraphs floating away struck through, and red ticks where words were cut;
 * - change bars with a ✦ in the margin (hover: what the text was) and a change map on the
 *   scroll bar's edge, until the next change or `clearMarks()`.
 * Everything respects prefers-reduced-motion (static highlights instead of motion).
 */
import { paragraphsOf, paragraphText } from '../api/editor.js';
import type { Page, Rect, TextItem } from '../layout/types.js';
import type { DocParagraph } from '../model/content.js';
import type { XmlElement } from '../xml/index.js';
import type { DocxViewer } from './viewer.js';

// ---- diff ------------------------------------------------------------------------------------------

export interface ParagraphChange {
  /** Index (0-based) of the paragraph in the new text. */
  para: number;
  /** A new paragraph, or one whose text changed. */
  kind: 'inserted' | 'changed';
  /** Ranges of new text in it, [from, to) (the whole paragraph for an inserted one). */
  ranges: Array<[number, number]>;
  /** Offsets where text was removed without new text in its place. */
  cuts: number[];
  /** The paragraph's text before (changed paragraphs). */
  before?: string;
  /** Index (0-based) of the paragraph in the old text (changed paragraphs). */
  from?: number;
  /** Only the formatting changed (bold, style, alignment, list). */
  formatting?: boolean;
}

export interface ParagraphDiff {
  changes: ParagraphChange[];
  /** Removed paragraphs: their text, and the index of the new paragraph they were before (or the count, at the end). */
  deleted: Array<{ at: number; text: string }>;
  /** Paragraphs whose text did not change: [old index, new index]. */
  same: Array<[number, number]>;
}

/** A document's paragraphs at one moment: text, and formatting (to notice bold, styles, alignment…). */
export interface ParagraphSnapshot {
  texts: string[];
  formats: string[];
}

/** Longest common subsequence pairs of two sequences (indices), with a size limit (then: none). */
function lcs<T>(a: T[], b: T[], eq: (x: T, y: T) => boolean, limit = 360_000): Array<[number, number]> {
  const n = a.length;
  const m = b.length;
  if (!n || !m || n * m > limit) return [];
  const w = m + 1;
  const t = new Uint32Array((n + 1) * w);
  for (let i = n - 1; i >= 0; i--) for (let j = m - 1; j >= 0; j--) t[i * w + j] = eq(a[i]!, b[j]!) ? t[(i + 1) * w + j + 1]! + 1 : Math.max(t[(i + 1) * w + j]!, t[i * w + j + 1]!);
  const out: Array<[number, number]> = [];
  for (let i = 0, j = 0; i < n && j < m; ) {
    if (eq(a[i]!, b[j]!)) {
      out.push([i, j]);
      i++;
      j++;
    } else if (t[(i + 1) * w + j]! >= t[i * w + j + 1]!) i++;
    else j++;
  }
  return out;
}

const tokens = (s: string) => s.match(/\s+|[\p{L}\p{N}_]+|[^\s\p{L}\p{N}_]/gu) ?? [];

/** What changed inside one paragraph: ranges of new text and points where text was cut. */
export function diffText(before: string, after: string): { ranges: Array<[number, number]>; cuts: number[] } {
  if (before === after) return { ranges: [], cuts: [] };
  const a = tokens(before);
  const b = tokens(after);
  // Common prefix and suffix first, then a token LCS on the middle.
  let p = 0;
  while (p < a.length && p < b.length && a[p] === b[p]) p++;
  let s = 0;
  while (s < a.length - p && s < b.length - p && a[a.length - 1 - s] === b[b.length - 1 - s]) s++;
  const am = a.slice(p, a.length - s);
  const bm = b.slice(p, b.length - s);
  const pairs = lcs(am, bm, (x, y) => x === y);
  const keptB = new Set(pairs.map(([, j]) => j + p));
  const keptA = new Set(pairs.map(([i]) => i + p));
  const offsets: number[] = [];
  let o = 0;
  for (const t of b) {
    offsets.push(o);
    o += t.length;
  }
  offsets.push(o);
  const ranges: Array<[number, number]> = [];
  for (let j = p; j < b.length - s; j++) {
    if (keptB.has(j) || /^\s+$/.test(b[j]!)) continue;
    const from = offsets[j]!;
    const to = offsets[j + 1]!;
    const last = ranges[ranges.length - 1];
    // Join ranges separated only by spaces.
    if (last && after.slice(last[1], from).trim() === '') last[1] = to;
    else ranges.push([from, to]);
  }
  // Removed tokens with nothing new next to them: a cut at the following kept token.
  const cuts: number[] = [];
  let bi = p;
  for (let i = p; i < a.length - s; i++) {
    if (keptA.has(i)) {
      const pair = pairs.find(([x]) => x + p === i)!;
      bi = pair[1] + p + 1;
      continue;
    }
    if (/^\s+$/.test(a[i]!)) continue;
    const at = offsets[Math.min(bi, b.length)]!;
    if (!ranges.some(([f, t]) => at >= f && at <= t) && !cuts.includes(at)) cuts.push(at);
  }
  return { ranges, cuts };
}

function similarity(a: string, b: string): number {
  const wa = new Set(a.toLowerCase().match(/[\p{L}\p{N}]+/gu) ?? []);
  const wb = new Set(b.toLowerCase().match(/[\p{L}\p{N}]+/gu) ?? []);
  if (!wa.size && !wb.size) return 1;
  let common = 0;
  for (const w of wa) if (wb.has(w)) common++;
  return common / Math.max(wa.size, wb.size);
}

/** Paragraph-level changes between two versions of a document's paragraph texts. */
export function diffParagraphs(before: string[], after: string[]): ParagraphDiff {
  const changes: ParagraphChange[] = [];
  const deleted: ParagraphDiff['deleted'] = [];
  const same: ParagraphDiff['same'] = [];
  let p = 0;
  while (p < before.length && p < after.length && before[p] === after[p]) p++;
  let s = 0;
  while (s < before.length - p && s < after.length - p && before[before.length - 1 - s] === after[after.length - 1 - s]) s++;
  for (let i = 0; i < p; i++) same.push([i, i]);
  for (let k = 1; k <= s; k++) same.push([before.length - k, after.length - k]);
  const a = before.slice(p, before.length - s);
  const b = after.slice(p, after.length - s);
  const anchors = [...lcs(a, b, (x, y) => x === y && x.trim() !== ''), [a.length, b.length] as [number, number]];
  let ai = 0;
  let bi = 0;
  for (const [ae, be] of anchors) {
    // A gap between unchanged paragraphs: pair similar ones as changes, the rest are inserts and deletes.
    const oldGap = a.slice(ai, ae);
    const newGap = b.slice(bi, be);
    let k = 0;
    for (let j = 0; j < newGap.length; j++) {
      const para = p + bi + j;
      const text = newGap[j]!;
      // Find the next old paragraph that is similar enough to count as this one changed.
      let match = -1;
      for (let i = k; i < oldGap.length && match < 0; i++) if (similarity(oldGap[i]!, text) >= 0.34) match = i;
      if (match < 0 && oldGap.length === newGap.length && k < oldGap.length) match = k;
      if (match >= 0) {
        for (let i = k; i < match; i++) deleted.push({ at: para, text: oldGap[i]! });
        const d = diffText(oldGap[match]!, text);
        changes.push({ para, kind: 'changed', ranges: d.ranges, cuts: d.cuts, before: oldGap[match]!, from: p + ai + match });
        k = match + 1;
      } else {
        changes.push({ para, kind: 'inserted', ranges: text ? [[0, text.length]] : [], cuts: [] });
      }
    }
    for (let i = k; i < oldGap.length; i++) deleted.push({ at: p + be, text: oldGap[i]! });
    if (ae < a.length) same.push([p + ae, p + be]);
    ai = ae + 1;
    bi = be + 1;
  }
  same.sort((x, y) => x[1] - y[1]);
  return { changes, deleted, same };
}

/** A paragraph's formatting as a string: style, alignment, list, and each run's properties. */
function formatSignature(p: DocParagraph): string {
  const runs = p.inlines.map((i) => (i.kind === 'text' ? JSON.stringify(i.props) : i.kind));
  return `${p.styleId ?? ''}|${p.props.jc ?? ''}|${p.props.bidi ? 'rtl' : ''}|${p.numbering ? `${p.numbering.numId}.${p.numbering.ilvl}` : ''}|${runs.join(';')}`;
}

// ---- drawing ---------------------------------------------------------------------------------------

const STYLE_ID = 'docxkeun-ai-style';

/** Highlights without motion: for prefers-reduced-motion, and for `animate: false` (scoped by `prefix`). */
function still(prefix: string): string {
  return [
    `${prefix}.docxkeun-ai-shimmer{animation:none;background-image:none;background-color:color-mix(in srgb,var(--dk-ai-1) 14%,transparent)}`,
    `${prefix}.docxkeun-ai-scan,${prefix}.docxkeun-ai-ink,${prefix}.docxkeun-ai-pen{display:none}`,
    `${prefix}.docxkeun-ai-glow{animation:docxkeun-ai-fade 3s ease-out both}`,
    `${prefix}.docxkeun-ai-caret{transition:none;animation:none}`,
    `${prefix}.docxkeun-ai-ghost,${prefix}.docxkeun-ai-cut{animation:docxkeun-ai-fade 2.4s both}`,
    `${prefix}.docxkeun-ai-bar,${prefix}.docxkeun-ai-spark{animation:none}`,
    // The other working styles, held still.
    `${prefix}.docxkeun-ai-in,${prefix}.docxkeun-ai-flash{animation:none}`,
    `${prefix}.docxkeun-ai-underline{animation:none;background-image:linear-gradient(90deg,var(--dk-ai-1),var(--dk-ai-2));background-size:100% 100%}`,
    `${prefix}.docxkeun-ai-edge{animation:none;box-shadow:0 0 0 1.5px color-mix(in srgb,var(--dk-ai-1) 70%,transparent),0 0 10px color-mix(in srgb,var(--dk-ai-1) 30%,transparent)}`,
    `${prefix}.docxkeun-ai-spot,${prefix}.docxkeun-ai-pulse,${prefix}.docxkeun-ai-intro{animation:none}`,
    `${prefix}.docxkeun-ai-sparkle{animation:none;opacity:.75}`,
  ].join('\n');
}
const CSS = `
.docxkeun-viewer{--dk-ai-1:var(--docxkeun-ai-color-1,#7c5cff);--dk-ai-2:var(--docxkeun-ai-color-2,var(--docxkeun-ai-color-1,#d97757));--dk-ai-text:var(--docxkeun-ai-text,#fff);--docxkeun-ai-page:#fff}
.docxkeun-ai-layer{position:absolute;left:0;top:0;width:0;height:0;pointer-events:none}
.docxkeun-ai-layer>*{position:absolute;pointer-events:none}
.docxkeun-ai-paint{position:absolute;left:0;top:0;width:0;height:0;overflow:hidden;pointer-events:none}
.docxkeun-ai-in{animation:docxkeun-ai-stop-in .9s ease-out var(--t,0ms) both}
@keyframes docxkeun-ai-stop-in{from{stop-color:var(--from)}}
.docxkeun-ai-flash{animation:docxkeun-ai-stop-flash 3.2s ease-in-out var(--t,0ms) both}
@keyframes docxkeun-ai-stop-flash{from,to{stop-color:var(--from)}20%,65%{stop-color:var(--to)}}
.docxkeun-ai-underline{border-radius:1px;background-color:color-mix(in srgb,var(--dk-ai-1) 18%,transparent);background-image:linear-gradient(90deg,transparent,var(--dk-ai-1),var(--dk-ai-2),transparent);background-size:35% 100%;background-repeat:no-repeat;animation:docxkeun-ai-wave 1.4s ease-in-out infinite}
@keyframes docxkeun-ai-wave{from{background-position:-40% 0}to{background-position:140% 0}}
.docxkeun-ai-edge{border-radius:6px;animation:docxkeun-ai-aura 2.2s ease-in-out infinite}
@keyframes docxkeun-ai-aura{0%,100%{box-shadow:0 0 0 1.5px color-mix(in srgb,var(--dk-ai-1) 70%,transparent),0 0 12px color-mix(in srgb,var(--dk-ai-1) 35%,transparent)}50%{box-shadow:0 0 0 1.5px color-mix(in srgb,var(--dk-ai-2) 70%,transparent),0 0 16px color-mix(in srgb,var(--dk-ai-2) 40%,transparent)}}
.docxkeun-ai-spot{border-radius:4px;box-shadow:0 0 0 4000px rgba(255,255,255,.55);animation:docxkeun-ai-in .4s ease-out both}
.docxkeun-ai-pulse{width:3px;border-radius:2px;background:linear-gradient(var(--dk-ai-1),var(--dk-ai-2));animation:docxkeun-ai-pulse 1.2s ease-in-out infinite}
@keyframes docxkeun-ai-pulse{0%,100%{opacity:.35;transform:scaleY(.85)}50%{opacity:1;transform:none}}
.docxkeun-ai-sparkle{width:12px;height:12px;margin:-6px 0 0 -6px;font:11px/12px system-ui,sans-serif;text-align:center;animation:docxkeun-ai-float 2.4s ease-in-out var(--t) infinite both}
@keyframes docxkeun-ai-float{0%{opacity:0;transform:translateY(4px) scale(.4) rotate(-40deg)}30%{opacity:1;transform:translateY(-2px) scale(1)}to{opacity:0;transform:translateY(-14px) scale(.6) rotate(40deg)}}
.docxkeun-ai-mark-tint{border-radius:3px;background:linear-gradient(90deg,color-mix(in srgb,var(--dk-ai-1) 14%,transparent),color-mix(in srgb,var(--dk-ai-2) 14%,transparent));mix-blend-mode:multiply}
.docxkeun-ai-mark-underline{border-radius:1px;background:linear-gradient(90deg,var(--dk-ai-1),var(--dk-ai-2))}
.docxkeun-ai-intro{animation:docxkeun-ai-in .9s ease-out var(--t,0ms) both}
.docxkeun-ai-mark-underline.docxkeun-ai-intro{transform-origin:left;animation-name:docxkeun-ai-draw}
@keyframes docxkeun-ai-in{from{opacity:0}}
@keyframes docxkeun-ai-draw{from{transform:scaleX(0)}}
.docxkeun-ai-shimmer{border-radius:3px;background-color:color-mix(in srgb,var(--dk-ai-1) 7%,transparent);background-image:linear-gradient(100deg,color-mix(in srgb,var(--dk-ai-1) 0%,transparent) 25%,color-mix(in srgb,var(--dk-ai-1) 30%,transparent) 42%,color-mix(in srgb,var(--dk-ai-2) 30%,transparent) 58%,color-mix(in srgb,var(--dk-ai-2) 0%,transparent) 75%);background-size:300% 100%;mix-blend-mode:multiply;animation:docxkeun-ai-shimmer 1.6s ease-in-out infinite}
@keyframes docxkeun-ai-shimmer{from{background-position:100% 0}to{background-position:0 0}}
.docxkeun-ai-scan{height:46px;background:linear-gradient(180deg,color-mix(in srgb,var(--dk-ai-1) 0%,transparent),color-mix(in srgb,var(--dk-ai-1) 16%,transparent) 60%,color-mix(in srgb,var(--dk-ai-2) 45%,transparent) 96%,color-mix(in srgb,var(--dk-ai-2) 0%,transparent));mix-blend-mode:multiply;animation:docxkeun-ai-scan 1.3s cubic-bezier(.4,0,.2,1) forwards}
@keyframes docxkeun-ai-scan{from{transform:translateY(0);opacity:0}12%{opacity:1}85%{opacity:1}to{transform:translateY(var(--travel));opacity:0}}
.docxkeun-ai-caret{z-index:2;width:2px;border-radius:1px;background:linear-gradient(var(--dk-ai-1),var(--dk-ai-2));transition:left .45s cubic-bezier(.3,.7,.2,1),top .45s cubic-bezier(.3,.7,.2,1),height .2s,opacity .4s;box-shadow:0 0 6px color-mix(in srgb,var(--dk-ai-1) 60%,transparent)}
.docxkeun-ai-caret.working{animation:docxkeun-ai-blink 1s steps(2,jump-none) infinite}
@keyframes docxkeun-ai-blink{50%{opacity:.35}}
.docxkeun-ai-caret span{position:absolute;left:0;bottom:100%;margin-bottom:2px;white-space:nowrap;font:600 11px/1 system-ui,sans-serif;color:var(--dk-ai-text);padding:3px 6px;border-radius:4px 4px 4px 0;background:linear-gradient(90deg,var(--dk-ai-1),var(--dk-ai-2));box-shadow:0 1px 3px rgba(0,0,0,.2)}
.docxkeun-ai-ink{background:var(--docxkeun-ai-page);animation:docxkeun-ai-wipe var(--d) cubic-bezier(.45,.05,.55,.95) var(--t) both}
@keyframes docxkeun-ai-wipe{from{clip-path:inset(0 0 0 0)}to{clip-path:inset(0 0 0 100%)}}
.docxkeun-ai-glow{border-radius:2px;background:linear-gradient(90deg,color-mix(in srgb,var(--dk-ai-1) 26%,transparent),color-mix(in srgb,var(--dk-ai-2) 26%,transparent));mix-blend-mode:multiply;animation:docxkeun-ai-glow calc(var(--d) + 2.6s) ease-out var(--t) both}
@keyframes docxkeun-ai-glow{0%{clip-path:inset(0 100% 0 0);opacity:1}25%{clip-path:inset(0 0 0 0);opacity:1}to{clip-path:inset(0 0 0 0);opacity:0}}
.docxkeun-ai-pen{width:9px;height:9px;margin:-4px 0 0 -4px;background:radial-gradient(circle,#fff 0,var(--dk-ai-2) 35%,color-mix(in srgb,var(--dk-ai-1) 0%,transparent) 70%);border-radius:50%;animation:docxkeun-ai-pen var(--d) cubic-bezier(.45,.05,.55,.95) var(--t) both}
@keyframes docxkeun-ai-pen{from{transform:translateX(0);opacity:1}95%{opacity:1}to{transform:translateX(var(--w));opacity:0}}
.docxkeun-ai-ghost{font:12px/1.4 system-ui,sans-serif;color:#c5221f;text-decoration:line-through;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;background:rgba(252,232,230,.9);border-left:2px solid #c5221f;padding:1px 6px;border-radius:2px;animation:docxkeun-ai-ghost 2.2s ease-in both}
@keyframes docxkeun-ai-ghost{0%{opacity:0;transform:translateY(4px)}15%{opacity:1;transform:none}70%{opacity:.9;transform:translateY(-2px)}to{opacity:0;transform:translateY(-14px)}}
.docxkeun-ai-cut{width:2px;background:#c5221f;border-radius:1px;animation:docxkeun-ai-cut 1.8s ease-out both}
@keyframes docxkeun-ai-cut{0%{opacity:0;transform:scaleY(.2)}15%{opacity:1;transform:none}to{opacity:0}}
.docxkeun-ai-bar{width:3px;border-radius:2px;background:linear-gradient(var(--dk-ai-1),var(--dk-ai-2));pointer-events:auto;cursor:help;animation:docxkeun-ai-bar .5s ease-out both}
@keyframes docxkeun-ai-bar{from{transform:scaleY(0)}}
.docxkeun-ai-spark{width:12px;height:12px;margin:-6px 0 0 -5px;color:var(--dk-ai-2);font:12px/12px system-ui,sans-serif;text-align:center;animation:docxkeun-ai-spark 1.4s ease-out both}
@keyframes docxkeun-ai-spark{0%{transform:scale(0) rotate(-90deg);opacity:0}40%{transform:scale(1.5) rotate(0);opacity:1}to{transform:scale(1);opacity:.9}}
.docxkeun-ai-map{position:sticky;top:0;float:right;width:8px;margin-right:2px;z-index:6;pointer-events:none}
.docxkeun-ai-map i{position:absolute;right:0;width:8px;height:4px;border-radius:2px;background:linear-gradient(90deg,var(--dk-ai-1),var(--dk-ai-2));pointer-events:auto;cursor:pointer;opacity:.85}
.docxkeun-ai-map i:hover{opacity:1;height:6px}
@keyframes docxkeun-ai-fade{0%,40%{opacity:1}to{opacity:0}}
@media (prefers-reduced-motion:reduce){${still('')}}
${still('.docxkeun-ai-still ')}`;

function ensureStyle(): void {
  if (typeof document === 'undefined' || document.getElementById(STYLE_ID)) return;
  const style = document.createElement('style');
  style.id = STYLE_ID;
  style.textContent = CSS;
  document.head.appendChild(style);
}

/** Whether the browser accepts this as a colour (no check where `CSS.supports` is missing). */
function isCssColor(value: string): boolean {
  if (!value.trim()) return false;
  // `CSS` is this module's style sheet; the browser's CSS namespace is on globalThis.
  const css = (globalThis as { CSS?: { supports?: (property: string, value: string) => boolean } }).CSS;
  return typeof css?.supports !== 'function' || css.supports('color', value);
}

/** sRGB channels (0-255) of a colour as the browser resolves it, when it can be read. */
function resolveRgb(color: string, host: HTMLElement): [number, number, number] | undefined {
  const probe = document.createElement('span');
  probe.style.color = color;
  probe.style.display = 'none';
  host.appendChild(probe);
  const resolved = getComputedStyle(probe).color || probe.style.color || color;
  probe.remove();
  const hex = /^#([0-9a-f]{3}|[0-9a-f]{6})$/i.exec(resolved.trim());
  if (hex) {
    const h = hex[1]!.length === 3 ? [...hex[1]!].map((d) => d + d).join('') : hex[1]!;
    return [0, 2, 4].map((i) => parseInt(h.slice(i, i + 2), 16)) as [number, number, number];
  }
  const rgb = /rgba?\(\s*([\d.]+)[\s,]+([\d.]+)[\s,]+([\d.]+)/i.exec(resolved);
  return rgb ? [Number(rgb[1]), Number(rgb[2]), Number(rgb[3])] : undefined;
}

function luminance([r, g, b]: [number, number, number]): number {
  const lin = (c: number) => {
    const s = c / 255;
    return s <= 0.03928 ? s / 12.92 : ((s + 0.055) / 1.055) ** 2.4;
  };
  return 0.2126 * lin(r) + 0.7152 * lin(g) + 0.0722 * lin(b);
}

const DARK_TEXT = '#1f1f1f';

/**
 * The caret label's text colour on `colors`: the default white (`undefined`) while it keeps a
 * 3:1 contrast (WCAG's bar for bold UI text) on every colour, else `DARK_TEXT` when that reads
 * better. White stays when a colour can't be read.
 */
function labelText(colors: string[], host: HTMLElement): string | undefined {
  const lums = colors.map((c) => resolveRgb(c, host)).map((rgb) => (rgb ? luminance(rgb) : undefined));
  if (lums.some((l) => l === undefined)) return undefined;
  const darkL = luminance([31, 31, 31]);
  const onWhite = Math.min(...lums.map((l) => 1.05 / (l! + 0.05)));
  const onDark = Math.min(...lums.map((l) => (l! + 0.05) / (darkL + 0.05)));
  return onWhite < 3 && onDark > onWhite ? DARK_TEXT : undefined;
}

/** The bounding box of some lines on each page they are on. */
function boxesByPage(lines: Array<{ page: Page } & Rect>): Array<{ page: Page } & Rect> {
  const byPage = new Map<Page, Array<{ page: Page } & Rect>>();
  for (const l of lines) byPage.set(l.page, [...(byPage.get(l.page) ?? []), l]);
  return [...byPage].map(([page, ls]) => {
    const x = Math.min(...ls.map((l) => l.x));
    const y = Math.min(...ls.map((l) => l.y));
    return { page, x, y, w: Math.max(...ls.map((l) => l.x + l.w)) - x, h: Math.max(...ls.map((l) => l.y + l.h)) - y };
  });
}

/** A fixed, even-looking spread in [0.1, 0.9) for the i-th item (no randomness: redraws match). */
function spread(i: number, axis: number): number {
  const golden = axis === 1 ? 0.6180339887 : 0.7548776662;
  return 0.1 + (((i + 1) * golden) % 1) * 0.8;
}

const reducedMotion = () => typeof matchMedia !== 'undefined' && matchMedia('(prefers-reduced-motion: reduce)').matches;
const nextFrame = () => new Promise<void>((ok) => requestAnimationFrame(() => requestAnimationFrame(() => ok())));

// ---- painting the letters ------------------------------------------------------------------------

const SVG_NS = 'http://www.w3.org/2000/svg';
let paintScopes = 0;
let paintLayers = 0;

/** A text item's colour as the page draws it (see svg.ts). */
const inkOf = (t: TextItem): string => (/^[0-9a-f]{6}$/i.test(t.color) ? `#${t.color}` : '#000');
const offset = (v: number): string => String(Math.round(Math.min(1, Math.max(0, v)) * 10000) / 10000);

/** A gradient stop; `var(...)` colours go in a style so the custom property resolves. */
function stop(at: number, color: string, extra: { cls?: string; vars?: Record<string, string> } = {}): string {
  const style = [...(color.startsWith('var(') ? [`stop-color:${color}`] : []), ...Object.entries(extra.vars ?? {}).map(([k, v]) => `${k}:${v}`)].join(';');
  return `<stop offset="${offset(at)}"${color.startsWith('var(') ? '' : ` stop-color="${color}"`}${extra.cls ? ` class="${extra.cls}"` : ''}${style ? ` style="${style}"` : ''}/>`;
}

/** The text items of a paragraph on a page: the SVG <text> elements to paint (by `data-i`). */
interface Glyphs {
  page: Page;
  index: number;
  item: TextItem;
  /** Spans [from, to) of character indexes to colour (all of the item when painting whole items). */
  spans: Array<[number, number]>;
}

/** One gradient and the glyphs it paints. */
interface PaintGroup {
  gradient: (id: string) => string;
  glyphs: Glyphs[];
}

/**
 * Colours chosen letters on the pages with SVG gradients: a style sheet points their <text>
 * elements at gradients defined beside it. Only the glyphs change, never the paper, shading,
 * highlights or pictures behind them, and pages drawn later (scrolling) pick the rules up too.
 */
class GlyphPaint {
  private readonly sheet: HTMLStyleElement;
  private readonly defs: SVGSVGElement;
  private readonly scope: string;
  private readonly id: string;

  constructor(root: HTMLElement, name: string) {
    this.scope = root.getAttribute('data-docxkeun-ai') ?? String(++paintScopes);
    root.setAttribute('data-docxkeun-ai', this.scope);
    this.id = `docxkeun-ai-${this.scope}-${name}-${++paintLayers}`;
    this.sheet = document.createElement('style');
    this.sheet.setAttribute('data-docxkeun-ai-paint', name);
    this.defs = document.createElementNS(SVG_NS, 'svg') as SVGSVGElement;
    this.defs.setAttribute('class', 'docxkeun-ai-paint');
    this.defs.setAttribute('aria-hidden', 'true');
    this.defs.setAttribute('data-paint', name);
    root.append(this.sheet, this.defs);
  }

  paint(groups: PaintGroup[]): void {
    const rules: string[] = [];
    const defs: string[] = [];
    groups.forEach((g, k) => {
      if (!g.glyphs.length) return;
      const id = `${this.id}-${k}`;
      defs.push(g.gradient(id));
      const sel = (x: Glyphs) => `[data-docxkeun-ai="${this.scope}"] .docxkeun-page[data-page="${x.page.index}"] svg>text[data-i="${x.index}"]`;
      rules.push(`${g.glyphs.map(sel).join(',')}{fill:url(#${id})}`);
      // Synthetic bold strokes the outline in the text colour: it takes the gradient too.
      const stroked = g.glyphs.filter((x) => x.item.font.syntheticBold);
      if (stroked.length) rules.push(`${stroked.map(sel).join(',')}{stroke:url(#${id})}`);
    });
    this.defs.innerHTML = defs.join('');
    this.sheet.textContent = rules.join('\n');
  }

  clear(): void {
    if (this.sheet.textContent) this.sheet.textContent = '';
    if (this.defs.firstChild) this.defs.replaceChildren();
  }

  remove(): void {
    this.sheet.remove();
    this.defs.remove();
  }
}

/**
 * The working shimmer: the letters keep their own colour (`ink`) and a band of the AI's colours
 * sweeps across the page, slightly slanted so lower lines follow on. `width`: the widest page.
 */
function sweepGradient(id: string, ink: string, width: number): string {
  return (
    `<linearGradient id="${id}" class="docxkeun-ai-words" gradientUnits="userSpaceOnUse" x1="-400" y1="-60" x2="400" y2="60">` +
    stop(0, ink) +
    stop(0.3, ink) +
    stop(0.44, 'var(--dk-ai-1)') +
    stop(0.56, 'var(--dk-ai-2)') +
    stop(0.7, ink) +
    stop(1, ink) +
    `<animateTransform attributeName="gradientTransform" type="translate" values="-220 0;${Math.round(width + 220)} 0" dur="2s" repeatCount="indefinite"/>` +
    `</linearGradient>`
  );
}

/** The shimmer held still (no motion): the letters take the AI's colours across the page. */
function stillGradient(id: string, width: number): string {
  return `<linearGradient id="${id}" class="docxkeun-ai-words" gradientUnits="userSpaceOnUse" x1="0" y1="0" x2="${Math.round(width)}" y2="0">${stop(0, 'var(--dk-ai-1)')}${stop(1, 'var(--dk-ai-2)')}</linearGradient>`;
}

/**
 * Changed words in the AI's colours and the rest of the item in its own: hard stops at the
 * edges of each span. `cls` animates the coloured stops from the ink (`docxkeun-ai-in` fades
 * them in, `docxkeun-ai-flash` in and out again), starting after `delay`.
 */
function spanGradient(id: string, kind: string, g: Glyphs, cls?: string, delay = '0ms'): string {
  const t = g.item;
  const ink = inkOf(t);
  const lefts = t.xs;
  const rights = t.xs.map((x, i) => x + (t.widths[i] ?? 0));
  const left = Math.min(...lefts);
  const right = Math.max(...rights, left + 0.01);
  const at = (x: number) => (x - left) / (right - left);
  const colour = (c: string) => (cls ? { cls, vars: { '--from': ink, '--to': c, '--t': delay } } : {});
  const stops: string[] = [];
  for (const [from, to] of g.spans) {
    const a = at(lefts[from]!);
    const b = at(rights[to - 1]!);
    // A span at the item's edge runs on past it (pad), so overhanging glyph ink is coloured too.
    if (from > 0) stops.push(stop(a, ink));
    stops.push(stop(a, 'var(--dk-ai-1)', colour('var(--dk-ai-1)')), stop(b, 'var(--dk-ai-2)', colour('var(--dk-ai-2)')));
    if (to < t.chars.length) stops.push(stop(b, ink));
  }
  return `<linearGradient id="${id}" class="${kind}" gradientUnits="userSpaceOnUse" x1="${left}" y1="0" x2="${right}" y2="0">${stops.join('')}</linearGradient>`;
}

/** How the AI's work shows on the paragraphs it is changing (see `ACTIVITY_STYLES`). */
export type ActivityStyle = 'highlight' | 'words' | 'underline' | 'edge' | 'spotlight' | 'quiet' | 'sparkles';

/** The styles, with a label and a line to show in a picker, in the order to show them. */
export const ACTIVITY_STYLES: ReadonlyArray<{ style: ActivityStyle; label: string; description: string }> = [
  { style: 'highlight', label: 'Highlight band', description: 'A tinted band sweeps over the lines; the changed paragraph keeps a soft tint.' },
  { style: 'words', label: 'Word shimmer', description: 'The letters themselves shimmer; the changed words stay coloured.' },
  { style: 'underline', label: 'Underline wave', description: 'A streak runs under each line; the changed words stay underlined.' },
  { style: 'edge', label: 'Edge glow', description: 'A soft glow breathes around the paragraph.' },
  { style: 'spotlight', label: 'Spotlight', description: 'The rest of the page dims while the paragraph is changed.' },
  { style: 'quiet', label: 'Quiet', description: 'Only the named caret and a pulsing bar in the margin.' },
  { style: 'sparkles', label: 'Sparkles', description: 'Small ✦ drift over the paragraph.' },
];

/** Styles whose own mark stays on the changed text; the others leave the ✦ margin bar. */
const OWN_MARK: ReadonlySet<ActivityStyle> = new Set(['highlight', 'words', 'underline']);

export interface ActivityOptions {
  /** The name on the caret and in the change marks' tooltips (default "Claude"). */
  name?: string;
  /**
   * How the AI's work shows (default `'highlight'`), and so which mark stays after a change:
   * highlight → a still tint on the changed paragraph, words → the changed words' letters stay in the AI's colours,
   * underline → the changed words stay underlined, the others → the ✦ margin bar.
   */
  style?: ActivityStyle;
  /**
   * Colour of the caret, shimmer, glow and change marks: one CSS colour, or two for a gradient
   * (default `['#7c5cff', '#d97757']`). The caret's label text turns dark on light colours.
   */
  color?: string | [string, string];
  /** Show anything at all (default true). Off: calls do nothing on screen; `reveal` still returns the diff. */
  enabled?: boolean;
  /** Motion: shimmer sweep, ink, ghosts, scan line, the gliding caret (default true). Off: static highlights. */
  animate?: boolean;
  /** The marks that stay after a change, and the change map on the right edge (default true). */
  marks?: boolean;
}

interface Mark {
  para: number;
  /** Character ranges of the new text (none: formatting only, or only deletions). */
  ranges: Array<[number, number]>;
  before?: string;
  formatting?: boolean;
}

const isStyle = (value: unknown): value is ActivityStyle => ACTIVITY_STYLES.some((s) => s.style === value);

/** The AI's presence on the pages: see the module comment. Get it from `viewer.activity`. */
export class AiActivity {
  private aiName: string;
  private aiColor: string | [string, string] | undefined;
  private aiStyle: ActivityStyle = 'highlight';
  private readonly layer: HTMLDivElement;
  private readonly map: HTMLDivElement;
  private readonly caret: HTMLDivElement;
  private readonly label: HTMLSpanElement;
  private working: { paras: number[]; mode: 'edit' | 'read'; label?: string } | undefined;
  private shimmers: HTMLElement[] = [];
  private marks: Mark[] = [];
  private markEls: HTMLElement[] = [];
  /** Letters coloured by the marks, the passing flash of a change, and the working shimmer (later wins). */
  private readonly markPaint: GlyphPaint;
  private readonly flashPaint: GlyphPaint;
  private readonly workPaint: GlyphPaint;
  private flashTimer: ReturnType<typeof setTimeout> | undefined;
  private hideTimer: ReturnType<typeof setTimeout> | undefined;
  private readonly cleanup: Array<() => void> = [];
  private on: boolean;
  private motion: boolean;
  private keepMarks: boolean;

  constructor(private readonly viewer: DocxViewer, options: ActivityOptions = {}) {
    ensureStyle();
    this.aiName = options.name ?? 'Claude';
    this.on = options.enabled ?? true;
    this.motion = options.animate ?? true;
    this.keepMarks = options.marks ?? true;
    viewer.root.classList.toggle('docxkeun-ai-still', !this.motion);
    const root = viewer.root;
    this.map = document.createElement('div');
    this.map.className = 'docxkeun-ai-map';
    this.map.setAttribute('aria-hidden', 'true');
    root.prepend(this.map);
    this.layer = document.createElement('div');
    this.layer.className = 'docxkeun-ai-layer';
    root.appendChild(this.layer);
    this.caret = document.createElement('div');
    this.caret.className = 'docxkeun-ai-caret';
    this.caret.style.opacity = '0';
    this.label = document.createElement('span');
    this.caret.appendChild(this.label);
    this.layer.appendChild(this.caret);
    this.markPaint = new GlyphPaint(root, 'marks');
    this.flashPaint = new GlyphPaint(root, 'flash');
    this.workPaint = new GlyphPaint(root, 'work');
    if (options.color !== undefined) this.color = options.color;
    if (options.style !== undefined) this.style = options.style;
    // Positions depend on the layout and zoom: redraw what lasts.
    this.cleanup.push(viewer.on('layout', () => requestAnimationFrame(() => this.redraw())));
    const onScroll = () => this.drawMap();
    root.addEventListener('scroll', onScroll, { passive: true });
    this.cleanup.push(() => root.removeEventListener('scroll', onScroll));
  }

  // ---- options ----

  /** The name on the caret ("Claude is editing…") and in the change marks' tooltips. */
  get name(): string {
    return this.aiName;
  }

  set name(name: string) {
    this.aiName = name;
    this.label.textContent = this.working ? this.workingText(this.working) : name;
    // The change marks' tooltips name the AI too.
    if (this.marks.length) this.drawMarks();
  }

  /**
   * Colour of the caret, shimmer, glow and change marks: one CSS colour, or two for a gradient.
   * `undefined` goes back to the defaults (or to the `--docxkeun-ai-color-1` / `-2` CSS
   * variables when the page sets them). Throws a TypeError for something that isn't a CSS colour.
   */
  get color(): string | [string, string] | undefined {
    return this.aiColor;
  }

  set color(value: string | [string, string] | undefined) {
    const style = this.viewer.root.style;
    if (value === undefined) {
      this.aiColor = undefined;
      for (const name of ['--docxkeun-ai-color-1', '--docxkeun-ai-color-2', '--docxkeun-ai-text']) style.removeProperty(name);
      return;
    }
    const [first, second] = typeof value === 'string' ? [value, value] : value;
    for (const c of [first, second]) {
      if (typeof c !== 'string' || !isCssColor(c)) throw new TypeError(`"${String(c)}" is not a CSS colour`);
    }
    this.aiColor = typeof value === 'string' ? value : [first, second];
    style.setProperty('--docxkeun-ai-color-1', first);
    style.setProperty('--docxkeun-ai-color-2', second);
    // Keep the caret's label readable: dark text on light colours.
    const text = labelText([first, second], this.viewer.root);
    if (text) style.setProperty('--docxkeun-ai-text', text);
    else style.removeProperty('--docxkeun-ai-text');
  }

  /**
   * How the AI's work shows, and so which mark stays after a change (see `ActivityOptions.style`).
   * Changing it redraws the work in progress and the marks. Throws a TypeError for an unknown style.
   */
  get style(): ActivityStyle {
    return this.aiStyle;
  }

  set style(style: ActivityStyle) {
    if (!isStyle(style)) throw new TypeError(`Unknown AI effect style "${String(style)}" (use ${ACTIVITY_STYLES.map((s) => s.style).join(', ')})`);
    this.aiStyle = style;
    if (this.working) this.drawWorking(false);
    if (this.marks.length) this.drawMarks();
  }

  /** Show the AI's activity at all. Turning it off clears everything on screen. */
  get enabled(): boolean {
    return this.on;
  }

  set enabled(on: boolean) {
    this.on = on;
    if (on) return;
    this.working = undefined;
    for (const el of [...this.layer.children]) if (el !== this.caret) el.remove();
    this.shimmers = [];
    this.markEls = [];
    this.marks = [];
    for (const paint of [this.workPaint, this.markPaint, this.flashPaint]) paint.clear();
    this.caret.style.opacity = '0';
    this.drawMap();
  }

  /** Motion (sweeps, ink, ghosts, gliding caret); off leaves static highlights. */
  get animate(): boolean {
    return this.motion;
  }

  set animate(on: boolean) {
    this.motion = on;
    this.viewer.root.classList.toggle('docxkeun-ai-still', !on);
    // The word shimmer's sweep is an SVG animation, which the still style sheet cannot stop.
    if (this.working && this.aiStyle === 'words') this.drawWorking(false);
  }

  /** Change bars and the change map after a change. */
  get marksShown(): boolean {
    return this.keepMarks;
  }

  set marksShown(on: boolean) {
    this.keepMarks = on;
    this.drawMarks();
  }

  // ---- geometry ----

  private paragraphs(): DocParagraph[] {
    const doc = this.viewer.document;
    return doc ? paragraphsOf(doc.body().blocks) : [];
  }

  /** Where a page point is in the scroll container. */
  private at(page: number, x: number, y: number): { left: number; top: number } | undefined {
    const el = this.viewer.pageElement(page);
    if (!el) return undefined;
    const s = this.viewer.scale;
    return { left: el.offsetLeft + x * s, top: el.offsetTop + y * s };
  }

  private lineRects(el: XmlElement): Array<{ page: Page } & Rect> {
    const out: Array<{ page: Page } & Rect> = [];
    for (const page of this.viewer.layout?.pages ?? []) for (const l of page.lines) if (l.paragraph === el) out.push({ page, x: l.x, y: l.y, w: l.w, h: l.h });
    return out;
  }

  /** Rectangles (one per line piece) covering characters [from, to) of a paragraph. */
  private charRects(para: DocParagraph, from: number, to: number): Array<{ page: Page } & Rect> {
    const starts = new Map<XmlElement, number>();
    for (const seg of paragraphText(para, { objects: true }).segments) starts.set(seg.node, seg.start);
    const out: Array<{ page: Page } & Rect> = [];
    for (const page of this.viewer.layout?.pages ?? []) {
      for (const item of page.items) {
        if (item.type !== 'text' || !item.source) continue;
        const start = starts.get(item.source.inline.node);
        if (start === undefined) continue;
        const t = item as TextItem;
        const isText = item.source.inline.node.name === 'w:t';
        let run: Rect | undefined;
        const flush = () => {
          if (run) out.push({ page, ...run });
          run = undefined;
        };
        for (let i = 0; i < t.chars.length; i++) {
          const off = start + (isText ? item.source.offsets[i]! : 0);
          if (off < from || off >= to) {
            flush();
            continue;
          }
          const top = t.y - t.font.ascent * t.size;
          const h = (t.font.ascent + t.font.descent) * t.size;
          const x = t.xs[i]!;
          const w = t.widths[i]!;
          if (run && Math.abs(run.x + run.w - x) < 1.5) run.w = x + w - run.x;
          else {
            flush();
            run = { x, y: top, w, h };
          }
        }
        flush();
      }
    }
    // Merge pieces on the same line (items split by formatting).
    out.sort((a, b) => a.page.index - b.page.index || a.y - b.y || a.x - b.x);
    const merged: typeof out = [];
    for (const r of out) {
      const last = merged[merged.length - 1];
      if (last && last.page === r.page && Math.abs(last.y - r.y) < 2 && r.x - (last.x + last.w) < 4) {
        last.w = Math.max(last.w, r.x + r.w - last.x);
        last.h = Math.max(last.h, r.h);
      } else merged.push({ ...r });
    }
    return merged;
  }

  /**
   * A paragraph's text items, with the characters in `ranges` (all of them, list labels too, when
   * `ranges` is omitted). Items of other paragraphs on the same lines (table cells) are left out.
   */
  private glyphs(para: DocParagraph, ranges?: Array<[number, number]>): Glyphs[] {
    const starts = new Map<XmlElement, number>();
    for (const seg of paragraphText(para, { objects: true }).segments) starts.set(seg.node, seg.start);
    const out: Glyphs[] = [];
    for (const page of this.viewer.layout?.pages ?? []) {
      const lines = page.lines.filter((l) => l.paragraph === para.el);
      if (!lines.length) continue;
      page.items.forEach((item, index) => {
        if (item.type !== 'text' || !item.chars.length) return;
        const t = item as TextItem;
        if (!t.source) {
          // A list label or a field result: no source, so take it when it sits on one of the lines.
          if (!ranges && lines.some((l) => t.y >= l.y && t.y <= l.y + l.h && t.x >= l.x - 72 && t.x < l.x + l.w)) out.push({ page, index, item: t, spans: [[0, t.chars.length]] });
          return;
        }
        const start = starts.get(t.source.inline.node);
        if (start === undefined) return;
        if (!ranges) {
          out.push({ page, index, item: t, spans: [[0, t.chars.length]] });
          return;
        }
        const isText = t.source.inline.node.name === 'w:t';
        const spans: Array<[number, number]> = [];
        for (let i = 0; i < t.chars.length; i++) {
          const off = start + (isText ? t.source.offsets[i]! : 0);
          if (!ranges.some(([from, to]) => off >= from && off < to)) continue;
          const last = spans[spans.length - 1];
          if (last && last[1] === i) last[1] = i + 1;
          else spans.push([i, i + 1]);
        }
        if (spans.length) out.push({ page, index, item: t, spans });
      });
    }
    return out;
  }

  /** The working shimmer on the paragraphs' letters: one gradient per text colour. */
  private shimmerPaint(paras: DocParagraph[]): PaintGroup[] {
    const still = !this.motion || reducedMotion();
    const width = Math.max(612, ...(this.viewer.layout?.pages ?? []).map((p) => p.width));
    const byInk = new Map<string, Glyphs[]>();
    for (const para of paras) {
      for (const g of this.glyphs(para)) {
        const key = still ? '' : inkOf(g.item);
        byInk.set(key, [...(byInk.get(key) ?? []), g]);
      }
    }
    return [...byInk].map(([ink, glyphs]) => ({ glyphs, gradient: (id: string) => (still ? stillGradient(id, width) : sweepGradient(id, ink, width)) }));
  }

  private place(el: HTMLElement, r: { page: Page } & Rect, pad = 0): boolean {
    const p = this.at(r.page.index, r.x - pad, r.y - pad);
    if (!p) return false;
    const s = this.viewer.scale;
    el.style.left = `${p.left}px`;
    el.style.top = `${p.top}px`;
    el.style.width = `${(r.w + pad * 2) * s}px`;
    el.style.height = `${(r.h + pad * 2) * s}px`;
    return true;
  }

  /**
   * Add an overlay. `life` (ms): remove it then (animation events are not dispatched reliably for
   * pages in background tabs, so a timer does it); without it, it stays until redrawn.
   */
  private add(className: string, r: { page: Page } & Rect, pad = 0, vars: Record<string, string> = {}, life?: number): HTMLDivElement | undefined {
    const el = document.createElement('div');
    el.className = className;
    if (!this.place(el, r, pad)) return undefined;
    for (const [k, v] of Object.entries(vars)) el.style.setProperty(k, v);
    if (life !== undefined) setTimeout(() => el.remove(), life + 100);
    this.layer.appendChild(el);
    return el;
  }

  // ---- working ----

  /**
   * Show the AI working on paragraphs (numbers from 1, as read_document counts them): a shimmer
   * over them and its caret at the first. `mode: 'read'` sweeps a scan line over the view instead
   * when no paragraphs are given. Lasts until `idle()` or `reveal()`.
   */
  work(paragraphs: number[], options: { mode?: 'edit' | 'read'; label?: string; scroll?: boolean } = {}): void {
    if (!this.on) return;
    const paras = [...new Set(paragraphs)].filter((n) => n >= 1).sort((a, b) => a - b);
    this.working = { paras, mode: options.mode ?? 'edit', ...(options.label ? { label: options.label } : {}) };
    clearTimeout(this.hideTimer);
    this.drawWorking(options.scroll ?? true);
    if (!paras.length && this.working.mode === 'read') this.scan();
  }

  /** The AI finished (without a change to show). */
  idle(): void {
    if (!this.on) return;
    this.working = undefined;
    this.clearWorking();
    this.caret.classList.remove('working');
    this.label.textContent = this.name;
    clearTimeout(this.hideTimer);
    this.hideTimer = setTimeout(() => (this.caret.style.opacity = '0'), 2500);
  }

  private clearWorking(): void {
    for (const el of this.shimmers) el.remove();
    this.shimmers = [];
    this.workPaint.clear();
  }

  private drawWorking(scroll: boolean): void {
    this.clearWorking();
    const w = this.working;
    if (!w) return;
    const all = this.paragraphs();
    const paras = w.paras.map((n) => all[n - 1]).filter((p): p is DocParagraph => !!p);
    // The lines of each paragraph being worked on.
    const groups = paras.map((p) => this.lineRects(p.el));
    const first = groups.flat()[0];
    this.drawStyle(groups, first, paras);
    this.label.textContent = this.workingText(w);
    if (first) {
      this.moveCaret(first.page, first.x, first.y, first.h, true);
      if (scroll) this.ensureVisible(first);
    } else if (w.mode === 'read') this.parkCaret();
  }

  /** The working effect of the current style over each paragraph's lines. */
  private drawStyle(groups: Array<Array<{ page: Page } & Rect>>, first: ({ page: Page } & Rect) | undefined, paras: DocParagraph[]): void {
    const keep = (el: HTMLElement | undefined): HTMLElement | undefined => {
      if (el) this.shimmers.push(el);
      return el;
    };
    // Ripple down the lines.
    let line = 0;
    const ripple = (el: HTMLElement | undefined) => {
      if (el) el.style.animationDelay = `${(line++ % 12) * 70}ms`;
    };
    const lines = groups.flat();
    switch (this.aiStyle) {
      case 'highlight':
        for (const r of lines) ripple(keep(this.add('docxkeun-ai-shimmer', r, 1.5)));
        break;
      case 'words':
        // The letters themselves take the colours; the paper, shading and pictures stay as they are.
        this.workPaint.paint(this.shimmerPaint(paras));
        break;
      case 'underline':
        for (const r of lines) ripple(keep(this.add('docxkeun-ai-underline', { ...r, y: r.y + r.h - 0.4, h: 1.6 })));
        break;
      case 'edge':
        for (const g of groups) for (const box of boxesByPage(g)) keep(this.add('docxkeun-ai-edge', box, 3));
        break;
      case 'spotlight': {
        // One clear window around the lines on the caret's page; the rest of the view dims.
        const box = first ? boxesByPage(lines).find((b) => b.page === first.page) : undefined;
        if (box) keep(this.add('docxkeun-ai-spot', box, 4));
        break;
      }
      case 'quiet':
        for (const g of groups) {
          for (const box of boxesByPage(g)) keep(this.add('docxkeun-ai-pulse', { page: box.page, x: Math.max(4, box.page.content.x - 12), y: box.y, w: 3, h: box.h }));
        }
        break;
      case 'sparkles': {
        let i = 0;
        for (const r of lines) {
          for (let k = 0, n = Math.min(6, Math.max(1, Math.round(r.w / 70))); k < n && i < 40; k++, i++) {
            const el = keep(this.add('docxkeun-ai-sparkle', { page: r.page, x: r.x + spread(i, 1) * r.w, y: r.y + spread(i, 2) * r.h, w: 0, h: 0 }));
            if (!el) continue;
            el.textContent = '✦';
            el.style.width = el.style.height = '12px';
            el.style.color = i % 2 ? 'var(--dk-ai-1)' : 'var(--dk-ai-2)';
            el.style.setProperty('--t', `${(i * 370) % 2400}ms`);
          }
        }
        break;
      }
    }
  }

  private workingText(w: { mode: 'edit' | 'read'; label?: string }): string {
    const verb = w.label ?? (w.mode === 'read' ? 'reading' : 'editing');
    return `${this.aiName} is ${verb}…`;
  }

  private moveCaret(page: Page, x: number, y: number, h: number, working: boolean): void {
    const p = this.at(page.index, x, y);
    if (!p) return;
    this.caret.style.left = `${p.left - 2}px`;
    this.caret.style.top = `${p.top}px`;
    this.caret.style.height = `${h * this.viewer.scale}px`;
    this.caret.style.opacity = '1';
    this.caret.classList.toggle('working', working);
  }

  /** The caret near the top of the view (reading the whole document). */
  private parkCaret(): void {
    const root = this.viewer.root;
    const page = (this.viewer.layout?.pages ?? []).find((pg) => {
      const el = this.viewer.pageElement(pg.index);
      return el && el.offsetTop + el.offsetHeight > root.scrollTop;
    });
    if (!page) return;
    const el = this.viewer.pageElement(page.index)!;
    const y = Math.max(page.content.y, (root.scrollTop - el.offsetTop) / this.viewer.scale + 24);
    this.moveCaret(page, page.content.x, y, 14, true);
  }

  private scan(): void {
    const root = this.viewer.root;
    const el = document.createElement('div');
    el.className = 'docxkeun-ai-scan';
    const first = this.viewer.pageElement(0);
    el.style.left = `${first ? first.offsetLeft : 0}px`;
    el.style.width = `${first ? first.offsetWidth : root.clientWidth}px`;
    el.style.top = `${root.scrollTop}px`;
    el.style.setProperty('--travel', `${Math.max(0, root.clientHeight - 46)}px`);
    setTimeout(() => el.remove(), 1400);
    this.layer.appendChild(el);
  }

  private ensureVisible(r: { page: Page } & Rect): void {
    const p = this.at(r.page.index, r.x, r.y);
    if (!p) return;
    const root = this.viewer.root;
    const h = r.h * this.viewer.scale;
    if (p.top < root.scrollTop + 40 || p.top + h > root.scrollTop + root.clientHeight - 40) {
      root.scrollTo({ top: Math.max(0, p.top - root.clientHeight / 3), behavior: reducedMotion() ? 'auto' : 'smooth' });
    }
  }

  // ---- changes ----

  /** The body's paragraphs now (text and formatting): pass to `reveal()` after the document changed. */
  snapshot(): ParagraphSnapshot {
    const paras = this.paragraphs();
    return { texts: paras.map((p) => paragraphText(p, { objects: true }).text), formats: paras.map(formatSignature) };
  }

  /**
   * Show what changed since `before` (from `snapshot()`), once the viewer has laid out the new
   * version: ink, glow, ghosts and cuts now, change bars and the change map until the next change.
   */
  async reveal(before: ParagraphSnapshot | string[], options: { scroll?: boolean } = {}): Promise<ParagraphDiff> {
    await nextFrame();
    const paras = this.paragraphs();
    const after = paras.map((p) => paragraphText(p, { objects: true }).text);
    const old = Array.isArray(before) ? { texts: before, formats: [] } : before;
    const diff = diffParagraphs(old.texts, after);
    // Same text, other formatting (bold, a style, alignment, a list): the whole paragraph glows.
    if (old.formats.length) {
      for (const [o, n] of diff.same) {
        if (old.formats[o] !== undefined && old.formats[o] !== formatSignature(paras[n]!)) diff.changes.push({ para: n, kind: 'changed', ranges: [], cuts: [], formatting: true, before: old.texts[o]!, from: o });
      }
      diff.changes.sort((x, y) => x.para - y.para);
    }
    if (!this.on) return diff;
    this.working = undefined;
    this.clearWorking();
    // The marks of earlier changes stay (following their paragraphs); this change's join them.
    this.marks = this.mergeMarks(diff, after);
    this.drawMarks(new Set(diff.changes.map((c) => c.para)));
    // A style with its own mark lets the mark fade in; the others get the passing glow. The word
    // shimmer never tints the paper: without marks, the changed letters light up for a moment.
    const glow = !(OWN_MARK.has(this.aiStyle) && this.keepMarks) && this.aiStyle !== 'words';
    if (this.aiStyle === 'words' && !this.keepMarks) this.flash(diff, paras);

    // Ink in the new text, line after line, as if written (about 2.5 s at most).
    const pieces: Array<{ r: { page: Page } & Rect; para: number }> = [];
    for (const c of diff.changes) {
      const para = paras[c.para];
      if (!para) continue;
      for (const [from, to] of c.ranges) for (const r of this.charRects(para, from, to)) pieces.push({ r, para: c.para });
      if (!c.ranges.length && !c.cuts.length) for (const r of this.lineRects(para.el)) pieces.push({ r, para: c.para });
    }
    const total = pieces.reduce((s, p) => s + p.r.w, 0);
    const speed = Math.max(500, total / 2.4); // points per second
    let t = 0;
    for (const { r } of pieces) {
      const d = Math.min(900, Math.max(140, (r.w / speed) * 1000));
      const vars = { '--d': `${d}ms`, '--t': `${t}ms`, '--w': `${r.w * this.viewer.scale}px` };
      if (glow) this.add('docxkeun-ai-glow', r, 1, vars, t + d + 2600);
      this.add('docxkeun-ai-ink', r, 0.5, vars, t + d);
      const pen = this.add('docxkeun-ai-pen', { ...r, w: 0, h: 0, y: r.y + r.h / 2 }, 0, vars, t + d);
      if (pen) pen.style.width = pen.style.height = '9px';
      t += d * 0.85;
    }
    // Cuts: a red tick where words were removed.
    for (const c of diff.changes) {
      const para = paras[c.para];
      if (!para) continue;
      for (const at of c.cuts) {
        const r = this.charRects(para, Math.max(0, at - 1), at)[0] ?? this.charRects(para, at, at + 1)[0];
        if (r) this.add('docxkeun-ai-cut', { ...r, x: at > 0 ? r.x + r.w : r.x, w: 0 }, 1, {}, 1800);
      }
    }
    // Removed paragraphs float away where they were.
    for (const d of diff.deleted) {
      const next = paras[d.at];
      const line = next ? this.lineRects(next.el)[0] : undefined;
      const prev = !line && d.at > 0 ? this.lineRects(paras[d.at - 1]!.el).pop() : undefined;
      const r = line ?? (prev ? { ...prev, y: prev.y + prev.h } : undefined);
      if (!r) continue;
      const el = this.add('docxkeun-ai-ghost', { ...r, y: r.y - 9, h: 16, w: Math.min(r.page.content.w, 360) }, 0, {}, 2200);
      if (el) {
        el.textContent = d.text.replace(/￼/g, '').slice(0, 160) || '(empty paragraph)';
        el.style.height = 'auto';
      }
    }
    // The caret ends after the last change, then fades.
    const lastChange = diff.changes[diff.changes.length - 1];
    const end = lastChange ? (lastChange.ranges.length ? this.charRects(paras[lastChange.para]!, lastChange.ranges[lastChange.ranges.length - 1]![0], lastChange.ranges[lastChange.ranges.length - 1]![1]).pop() : this.lineRects(paras[lastChange.para]!.el).pop()) : undefined;
    if (end) {
      setTimeout(() => this.moveCaret(end.page, end.x + end.w, end.y, end.h, false), Math.min(t, 2400));
      if (options.scroll ?? true) {
        const firstChange = diff.changes[0];
        const first = firstChange ? this.lineRects(paras[firstChange.para]!.el)[0] : undefined;
        if (first) this.ensureVisible(first);
      }
    }
    this.caret.classList.remove('working');
    this.label.textContent = this.name;
    clearTimeout(this.hideTimer);
    this.hideTimer = setTimeout(() => (this.caret.style.opacity = '0'), Math.min(t, 2400) + 2500);
    return diff;
  }

  /**
   * Remove every mark and the change map. Marks build up over a run of changes (each follows its
   * paragraph as others are added or removed) until this is called, e.g. when a new request starts.
   */
  clearMarks(): void {
    this.marks = [];
    this.drawMarks();
  }

  /**
   * The marks after a change: earlier marks follow their paragraphs (and go with deleted ones), a
   * paragraph changed again is marked against its text before the first change, and new ones join.
   */
  private mergeMarks(diff: ParagraphDiff, after: string[]): Mark[] {
    const earlier = new Map(this.marks.map((m) => [m.para, m]));
    const out = new Map<number, Mark>();
    for (const [o, n] of diff.same) {
      const m = earlier.get(o);
      if (m) out.set(n, { ...m, para: n });
    }
    for (const c of diff.changes) {
      const prev = c.from !== undefined ? earlier.get(c.from) : undefined;
      if (!prev) {
        out.set(c.para, { para: c.para, ranges: c.ranges, ...(c.before !== undefined ? { before: c.before } : {}), ...(c.formatting ? { formatting: true } : {}) });
        continue;
      }
      const text = after[c.para] ?? '';
      const formatting = !!(prev.formatting || c.formatting);
      // A paragraph the AI added stays added; otherwise compare with the text before the first change.
      const d = prev.before === undefined ? { ranges: text ? ([[0, text.length]] as Array<[number, number]>) : [], cuts: [] } : diffText(prev.before, text);
      if (!d.ranges.length && !d.cuts.length && !formatting) {
        // Changed back to how it was: no mark.
        out.delete(c.para);
        continue;
      }
      out.set(c.para, { para: c.para, ranges: d.ranges, ...(prev.before !== undefined ? { before: prev.before } : {}), ...(formatting ? { formatting: true } : {}) });
    }
    return [...out.values()].sort((a, b) => a.para - b.para);
  }

  /** What a mark's tooltip says. */
  private markTitle(m: Mark): string {
    if (m.formatting && !m.ranges.length) return `Formatting changed by ${this.name}`;
    if (m.before !== undefined) return `Changed by ${this.name}. Before: ${m.before.replace(/￼/g, '').slice(0, 300) || '(empty)'}`;
    return `Added by ${this.name}`;
  }

  /** Draw the marks; those of the paragraphs in `fresh` (just changed) appear with a short animation. */
  private drawMarks(fresh: ReadonlySet<number> = new Set()): void {
    for (const el of this.markEls) el.remove();
    this.markEls = [];
    const paras = this.paragraphs();
    const painted: PaintGroup[] = [];
    let n = 0;
    let animate = false;
    // Marks of the style's own kind fade (or draw) in one after another.
    const intro = (el: HTMLElement | undefined) => {
      if (!el) return;
      if (animate) {
        el.classList.add('docxkeun-ai-intro');
        el.style.setProperty('--t', `${Math.min(n++ * 90, 1500)}ms`);
      }
      this.markEls.push(el);
    };
    for (const m of this.keepMarks ? this.marks : []) {
      animate = fresh.has(m.para);
      const para = paras[m.para];
      if (!para) continue;
      const lines = this.lineRects(para.el);
      const style = this.aiStyle;
      if (style === 'highlight') {
        // The changed paragraph keeps a still tint.
        for (const r of lines) intro(this.add('docxkeun-ai-mark-tint', r, 1.5));
        continue;
      }
      if (style === 'words') {
        // The changed words' letters keep the colours (the whole paragraph when only its formatting changed).
        const glyphs = m.ranges.length ? this.glyphs(para, m.ranges) : m.formatting ? this.glyphs(para) : [];
        // Decide now: the gradients are written after the loop, when `animate` holds the last mark's value.
        const cls = animate ? 'docxkeun-ai-in' : undefined;
        for (const g of glyphs) {
          const delay = `${Math.min(n++ * 90, 1500)}ms`;
          painted.push({ glyphs: [g], gradient: (id) => spanGradient(id, 'docxkeun-ai-mark-words', g, cls, delay) });
        }
        if (glyphs.length) continue;
        // Only deletions in this paragraph: nothing to colour, so the margin bar marks it.
      }
      if (style === 'underline') {
        // The changed words; the whole paragraph when only its formatting changed.
        const rects = m.ranges.length ? m.ranges.flatMap(([from, to]) => this.charRects(para, from, to)) : m.formatting ? lines : [];
        for (const r of rects) intro(this.add('docxkeun-ai-mark-underline', { ...r, y: r.y + r.h - 0.2, h: 1.4 }));
        if (rects.length) continue;
        // Only deletions in this paragraph: nothing to underline, so the margin bar marks it.
      }
      // One bar per page the paragraph is on, in the left margin.
      const byPage = new Map<Page, Array<{ page: Page } & Rect>>();
      for (const l of lines) byPage.set(l.page, [...(byPage.get(l.page) ?? []), l]);
      let firstBar = true;
      for (const [page, ls] of byPage) {
        const top = Math.min(...ls.map((l) => l.y));
        const bottom = Math.max(...ls.map((l) => l.y + l.h));
        const x = Math.max(4, page.content.x - 12);
        const bar = this.add('docxkeun-ai-bar', { page, x, y: top, w: 3, h: bottom - top });
        if (!bar) continue;
        if (!animate) bar.style.animation = 'none';
        bar.title = this.markTitle(m);
        this.markEls.push(bar);
        if (firstBar && animate) {
          const spark = this.add('docxkeun-ai-spark', { page, x: x + 1.5, y: top, w: 0, h: 0 });
          if (spark) {
            spark.textContent = '✦';
            spark.style.width = spark.style.height = '12px';
            this.markEls.push(spark);
          }
        }
        firstBar = false;
      }
    }
    this.markPaint.paint(painted);
    this.drawMap();
  }

  /** The changed letters light up in the AI's colours for a moment (the word shimmer without marks). */
  private flash(diff: ParagraphDiff, paras: DocParagraph[]): void {
    const groups: PaintGroup[] = [];
    let n = 0;
    for (const c of diff.changes) {
      const para = paras[c.para];
      if (!para) continue;
      const glyphs = c.ranges.length ? this.glyphs(para, c.ranges) : c.cuts.length ? [] : this.glyphs(para);
      for (const g of glyphs) {
        const delay = `${Math.min(n++ * 90, 1500)}ms`;
        groups.push({ glyphs: [g], gradient: (id) => spanGradient(id, 'docxkeun-ai-flash-words', g, 'docxkeun-ai-flash', delay) });
      }
    }
    clearTimeout(this.flashTimer);
    this.flashPaint.paint(groups);
    this.flashTimer = setTimeout(() => this.flashPaint.clear(), 3200 + Math.min(n * 90, 1500) + 100);
  }

  /** Ticks on the right edge for every change, scaled to the scroll height; click to jump. */
  private drawMap(): void {
    const root = this.viewer.root;
    this.map.replaceChildren();
    this.map.style.height = `${root.clientHeight}px`;
    this.map.style.marginBottom = `-${root.clientHeight}px`;
    if (!this.marks.length || !this.keepMarks) return;
    const paras = this.paragraphs();
    const height = Math.max(root.scrollHeight, 1);
    for (const m of this.marks) {
      const para = paras[m.para];
      const r = para ? this.lineRects(para.el)[0] : undefined;
      const p = r && this.at(r.page.index, r.x, r.y);
      if (!r || !p) continue;
      const tick = document.createElement('i');
      tick.style.top = `${(p.top / height) * root.clientHeight}px`;
      tick.title = this.markTitle(m);
      tick.addEventListener('click', () => root.scrollTo({ top: Math.max(0, p.top - root.clientHeight / 3), behavior: reducedMotion() ? 'auto' : 'smooth' }));
      this.map.appendChild(tick);
    }
  }

  /**
   * The viewer rebuilt its page rows (zoom, the comment column's width): redraw at the new positions.
   * @internal
   */
  rowsBuilt(): void {
    requestAnimationFrame(() => this.redraw());
  }

  private redraw(): void {
    // Transient animations are tied to the old positions: drop them; redraw what lasts.
    for (const el of [...this.layer.children]) {
      if (el !== this.caret && !this.markEls.includes(el as HTMLElement) && !this.shimmers.includes(el as HTMLElement)) el.remove();
    }
    this.flashPaint.clear();
    this.drawMarks();
    if (this.working) this.drawWorking(false);
  }

  destroy(): void {
    for (const fn of this.cleanup) fn();
    clearTimeout(this.hideTimer);
    clearTimeout(this.flashTimer);
    for (const paint of [this.workPaint, this.markPaint, this.flashPaint]) paint.remove();
    this.layer.remove();
    this.map.remove();
  }
}
