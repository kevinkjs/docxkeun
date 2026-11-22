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
        changes.push({ para, kind: 'changed', ranges: d.ranges, cuts: d.cuts, before: oldGap[match]! });
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
  ].join('\n');
}
const CSS = `
.docxkeun-viewer{--dk-ai-1:var(--docxkeun-ai-color-1,#7c5cff);--dk-ai-2:var(--docxkeun-ai-color-2,var(--docxkeun-ai-color-1,#d97757));--dk-ai-text:var(--docxkeun-ai-text,#fff);--docxkeun-ai-page:#fff}
.docxkeun-ai-layer{position:absolute;left:0;top:0;width:0;height:0;pointer-events:none;z-index:5}
.docxkeun-ai-layer>*{position:absolute;pointer-events:none}
.docxkeun-ai-shimmer{border-radius:3px;background-color:color-mix(in srgb,var(--dk-ai-1) 7%,transparent);background-image:linear-gradient(100deg,color-mix(in srgb,var(--dk-ai-1) 0%,transparent) 25%,color-mix(in srgb,var(--dk-ai-1) 30%,transparent) 42%,color-mix(in srgb,var(--dk-ai-2) 30%,transparent) 58%,color-mix(in srgb,var(--dk-ai-2) 0%,transparent) 75%);background-size:300% 100%;mix-blend-mode:multiply;animation:docxkeun-ai-shimmer 1.6s ease-in-out infinite}
@keyframes docxkeun-ai-shimmer{from{background-position:100% 0}to{background-position:0 0}}
.docxkeun-ai-scan{height:46px;background:linear-gradient(180deg,color-mix(in srgb,var(--dk-ai-1) 0%,transparent),color-mix(in srgb,var(--dk-ai-1) 16%,transparent) 60%,color-mix(in srgb,var(--dk-ai-2) 45%,transparent) 96%,color-mix(in srgb,var(--dk-ai-2) 0%,transparent));mix-blend-mode:multiply;animation:docxkeun-ai-scan 1.3s cubic-bezier(.4,0,.2,1) forwards}
@keyframes docxkeun-ai-scan{from{transform:translateY(0);opacity:0}12%{opacity:1}85%{opacity:1}to{transform:translateY(var(--travel));opacity:0}}
.docxkeun-ai-caret{width:2px;border-radius:1px;background:linear-gradient(var(--dk-ai-1),var(--dk-ai-2));transition:left .45s cubic-bezier(.3,.7,.2,1),top .45s cubic-bezier(.3,.7,.2,1),height .2s,opacity .4s;box-shadow:0 0 6px color-mix(in srgb,var(--dk-ai-1) 60%,transparent)}
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

const reducedMotion = () => typeof matchMedia !== 'undefined' && matchMedia('(prefers-reduced-motion: reduce)').matches;
const nextFrame = () => new Promise<void>((ok) => requestAnimationFrame(() => requestAnimationFrame(() => ok())));

export interface ActivityOptions {
  /** The name on the caret and in the change marks' tooltips (default "Claude"). */
  name?: string;
  /**
   * Colour of the caret, shimmer, glow and change marks: one CSS colour, or two for a gradient
   * (default `['#7c5cff', '#d97757']`). The caret's label text turns dark on light colours.
   */
  color?: string | [string, string];
  /** Show anything at all (default true). Off: calls do nothing on screen; `reveal` still returns the diff. */
  enabled?: boolean;
  /** Motion: shimmer sweep, ink, ghosts, scan line, the gliding caret (default true). Off: static highlights. */
  animate?: boolean;
  /** Change bars and the change map that stay after a change (default true). */
  marks?: boolean;
}

interface Mark {
  para: number;
  before?: string;
  formatting?: boolean;
}

/** The AI's presence on the pages: see the module comment. Get it from `viewer.activity`. */
export class AiActivity {
  private aiName: string;
  private aiColor: string | [string, string] | undefined;
  private readonly layer: HTMLDivElement;
  private readonly map: HTMLDivElement;
  private readonly caret: HTMLDivElement;
  private readonly label: HTMLSpanElement;
  private working: { paras: number[]; mode: 'edit' | 'read'; label?: string } | undefined;
  private shimmers: HTMLElement[] = [];
  private marks: Mark[] = [];
  private markEls: HTMLElement[] = [];
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
    if (options.color !== undefined) this.color = options.color;
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
    if (this.marks.length) this.drawMarks(false);
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
  }

  /** Change bars and the change map after a change. */
  get marksShown(): boolean {
    return this.keepMarks;
  }

  set marksShown(on: boolean) {
    this.keepMarks = on;
    this.drawMarks(false);
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
    for (const el of this.shimmers) el.remove();
    this.shimmers = [];
    this.caret.classList.remove('working');
    this.label.textContent = this.name;
    clearTimeout(this.hideTimer);
    this.hideTimer = setTimeout(() => (this.caret.style.opacity = '0'), 2500);
  }

  private drawWorking(scroll: boolean): void {
    for (const el of this.shimmers) el.remove();
    this.shimmers = [];
    const w = this.working;
    if (!w) return;
    const all = this.paragraphs();
    let first: ({ page: Page } & Rect) | undefined;
    let line = 0;
    for (const n of w.paras) {
      const para = all[n - 1];
      if (!para) continue;
      for (const r of this.lineRects(para.el)) {
        first ??= r;
        const el = this.add('docxkeun-ai-shimmer', r, 1.5);
        if (el) {
          // Ripple down the paragraph.
          el.style.animationDelay = `${(line++ % 12) * 70}ms`;
          this.shimmers.push(el);
        }
      }
    }
    this.label.textContent = this.workingText(w);
    if (first) {
      this.moveCaret(first.page, first.x, first.y, first.h, true);
      if (scroll) this.ensureVisible(first);
    } else if (w.mode === 'read') this.parkCaret();
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
        if (old.formats[o] !== undefined && old.formats[o] !== formatSignature(paras[n]!)) diff.changes.push({ para: n, kind: 'changed', ranges: [], cuts: [], formatting: true });
      }
      diff.changes.sort((x, y) => x.para - y.para);
    }
    if (!this.on) return diff;
    this.working = undefined;
    for (const el of this.shimmers) el.remove();
    this.shimmers = [];
    this.marks = diff.changes.map((c) => ({ para: c.para, ...(c.before !== undefined ? { before: c.before } : {}), ...(c.formatting ? { formatting: true } : {}) }));
    this.drawMarks(true);

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
      this.add('docxkeun-ai-glow', r, 1, vars, t + d + 2600);
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

  /** Remove the change bars and the change map. */
  clearMarks(): void {
    this.marks = [];
    this.drawMarks(false);
  }

  private drawMarks(animate: boolean): void {
    for (const el of this.markEls) el.remove();
    this.markEls = [];
    const paras = this.paragraphs();
    for (const m of this.keepMarks ? this.marks : []) {
      const para = paras[m.para];
      if (!para) continue;
      const lines = this.lineRects(para.el);
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
        bar.title = m.formatting
          ? `Formatting changed by ${this.name}`
          : m.before !== undefined
            ? `Changed by ${this.name}. Before: ${m.before.replace(/￼/g, '').slice(0, 300) || '(empty)'}`
            : `Added by ${this.name}`;
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
    this.drawMap();
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
      tick.title = `Changed by ${this.name}`;
      tick.addEventListener('click', () => root.scrollTo({ top: Math.max(0, p.top - root.clientHeight / 3), behavior: reducedMotion() ? 'auto' : 'smooth' }));
      this.map.appendChild(tick);
    }
  }

  private redraw(): void {
    // Transient animations are tied to the old positions: drop them; redraw what lasts.
    for (const el of [...this.layer.children]) {
      if (el !== this.caret && !this.markEls.includes(el as HTMLElement) && !this.shimmers.includes(el as HTMLElement)) el.remove();
    }
    this.drawMarks(false);
    if (this.working) this.drawWorking(false);
  }

  destroy(): void {
    for (const fn of this.cleanup) fn();
    clearTimeout(this.hideTimer);
    this.layer.remove();
    this.map.remove();
  }
}
