/**
 * Character formatting as light tags, so an AI can see it and keep it: read_document shows <b>, <i>,
 * <u>, <s>, <sup> and <sub> around text formatted that way, and the writing tools take the same tags.
 * Formatting the whole paragraph has (a bold title) is named once in the paragraph's description
 * instead of tagging all its text.
 */
import type { DocParagraph, TextInline } from 'docxkeun';
import type { TextFormat, TextPiece } from 'docxkeun/editor';

type XmlElement = DocParagraph['el'];

export const TAGS = ['b', 'i', 'u', 's', 'sup', 'sub'] as const;
export type Tag = (typeof TAGS)[number];
type Tags = ReadonlySet<Tag>;

const WORDS: Record<Tag, string> = { b: 'bold', i: 'italic', u: 'underlined', s: 'struck through', sup: 'superscript', sub: 'subscript' };
/** Formatting that shows on spaces and tabs too (a signature line is often underlined blanks). */
const SHOWS_ON_SPACE: ReadonlySet<Tag> = new Set(['u', 's']);
const TAG_OR_ENTITY = /<(\/?)(sup|sub|b|i|u|s)>|&(lt|gt|amp);/g;

export function tagsOf(props: TextInline['props']): Set<Tag> {
  const t = new Set<Tag>();
  if (props.bold) t.add('b');
  if (props.italic) t.add('i');
  if (props.underline && props.underline !== 'none') t.add('u');
  if (props.strike || props.dstrike) t.add('s');
  if (props.vertAlign === 'superscript') t.add('sup');
  else if (props.vertAlign === 'subscript') t.add('sub');
  return t;
}

const same = (a: Tags, b: Tags): boolean => a.size === b.size && [...a].every((t) => b.has(t));
const without = (a: Tags, b: Tags): Set<Tag> => new Set([...a].filter((t) => !b.has(t)));
const visible = (text: string): boolean => /\S/.test(text);

interface Piece {
  text: string;
  tags: Tags;
  run?: XmlElement;
  inLink: boolean;
}

/** The paragraph's text (text runs, tabs and line breaks, as read_document shows it) in formatted pieces. */
function piecesOf(para: DocParagraph): Piece[] {
  const out: Piece[] = [];
  let links = 0;
  for (const inline of para.inlines) {
    if (inline.kind === 'linkStart') links++;
    else if (inline.kind === 'linkEnd') links = Math.max(0, links - 1);
    else if (inline.kind === 'text' || inline.kind === 'tab' || (inline.kind === 'break' && inline.type === 'line')) {
      const text = inline.kind === 'text' ? inline.text : inline.kind === 'tab' ? '\t' : '\n';
      out.push({ text, tags: tagsOf(inline.props), run: inline.run, inLink: links > 0 });
    }
  }
  return out;
}

/** Formatting every visible character of the paragraph has. */
function wholeOf(pieces: Piece[]): Set<Tag> {
  const shown = pieces.filter((p) => visible(p.text));
  if (!shown.length) return new Set();
  return new Set(TAGS.filter((t) => shown.every((p) => p.tags.has(t))));
}

/** What formatting the paragraph's text has as a whole, in words ("bold, underlined"). */
export function wholeFormatting(para: DocParagraph): string[] {
  return [...wholeOf(piecesOf(para))].map((t) => WORDS[t]);
}

function escape(text: string): string {
  return text.replace(/&(?=(?:lt|gt|amp);)/g, '&amp;').replace(/<(?=\/?(?:sup|sub|b|i|u|s)>)/g, '&lt;');
}

/** The paragraph's text with tags around what only part of it is formatted with. */
export function markupOf(para: DocParagraph): string {
  const pieces = piecesOf(para);
  const whole = wholeOf(pieces);
  const shown = pieces.map((p) => ({ text: p.text, tags: without(p.tags, whole) }));
  // Bold or italic spaces look like any other spaces: they follow the text around them, so a bold
  // phrase reads <b>PT Swadaya Usaha Persada</b> however Word split it into runs.
  shown.forEach((p, i) => {
    if (visible(p.text)) return;
    const near = [shown[i - 1], shown[i + 1]].filter((n): n is (typeof shown)[number] => !!n && visible(n.text));
    const around = near.length === 2 && !same(near[0]!.tags, near[1]!.tags) ? new Set([...near[0]!.tags].filter((t) => near[1]!.tags.has(t))) : (near[0]?.tags ?? new Set<Tag>());
    p.tags = new Set([...[...p.tags].filter((t) => SHOWS_ON_SPACE.has(t)), ...[...around].filter((t) => !SHOWS_ON_SPACE.has(t))]);
  });
  let out = '';
  const merged: Array<{ text: string; tags: Tags }> = [];
  for (const p of shown) {
    const last = merged[merged.length - 1];
    if (last && same(last.tags, p.tags)) last.text += p.text;
    else merged.push({ ...p });
  }
  // Tags stay open across pieces that keep them: <b>6.337 M<sup>2</sup> (enam ribu…)</b>.
  const open: Tag[] = [];
  for (const p of merged) {
    let keep = 0;
    while (keep < open.length && p.tags.has(open[keep]!)) keep++;
    while (open.length > keep) out += `</${open.pop()}>`;
    for (const t of TAGS) {
      if (p.tags.has(t) && !open.includes(t)) {
        open.push(t);
        out += `<${t}>`;
      }
    }
    out += escape(p.text);
  }
  while (open.length) out += `</${open.pop()}>`;
  return out;
}

export interface Segment {
  text: string;
  tags: Set<Tag>;
}

/** Whether text has any formatting tags (or the escapes &lt; &gt; &amp;). */
export function hasMarkup(text: string): boolean {
  return new RegExp(TAG_OR_ENTITY.source).test(text);
}

/**
 * Split tagged text into pieces with their formatting. Lenient: a closing tag that was never opened
 * is ignored, a tag left open runs to the end, and any other "<" or "&" is ordinary text.
 */
export function parseMarkup(text: string): Segment[] {
  const out: Segment[] = [];
  const depth = new Map<Tag, number>();
  let last = 0;
  const push = (s: string): void => {
    if (!s) return;
    const tags = new Set(TAGS.filter((t) => (depth.get(t) ?? 0) > 0));
    const prev = out[out.length - 1];
    if (prev && same(prev.tags, tags)) prev.text += s;
    else out.push({ text: s, tags });
  };
  for (const m of text.matchAll(TAG_OR_ENTITY)) {
    push(text.slice(last, m.index));
    last = m.index + m[0].length;
    if (m[3]) {
      push(m[3] === 'lt' ? '<' : m[3] === 'gt' ? '>' : '&');
      continue;
    }
    const tag = m[2] as Tag;
    const n = depth.get(tag) ?? 0;
    depth.set(tag, m[1] ? Math.max(0, n - 1) : n + 1);
  }
  push(text.slice(last));
  return out;
}

/** The text without its tags. */
export function plainText(text: string): string {
  return parseMarkup(text)
    .map((s) => s.text)
    .join('');
}

/** Direct formatting that turns text formatted `have` into `want`. */
export function formatFor(want: Tags, have: Tags): TextFormat | undefined {
  const f: TextFormat = {};
  if (want.has('b') !== have.has('b')) f.bold = want.has('b');
  if (want.has('i') !== have.has('i')) f.italic = want.has('i');
  if (want.has('u') !== have.has('u')) f.underline = want.has('u');
  if (want.has('s') !== have.has('s')) f.strike = want.has('s');
  const va = (t: Tags) => (t.has('sup') ? 'superscript' : t.has('sub') ? 'subscript' : 'baseline');
  if (va(want) !== va(have)) f.vertAlign = va(want);
  return Object.keys(f).length ? f : undefined;
}

/**
 * New text for a paragraph as pieces for DocxEditor.replaceRange. The paragraph's text is read as
 * stretches of one formatting each; the new text's parts copy them in order, so the second bold part
 * of a translation copies the second bold phrase of the original (with its own font, size and colour).
 * A part whose formatting the paragraph never had copies its ordinary text and sets the difference.
 * Formatting the whole paragraph had needs no tags.
 */
export function piecesFor(para: DocParagraph, segments: Segment[]): TextPiece[] {
  const pieces = piecesOf(para);
  const whole = wholeOf(pieces);
  // Within a stretch, copy an ordinary run rather than a hyperlink's (its colour comes from the link), then the longest.
  const better = (a: Piece, b: Piece): boolean => (a.inLink !== b.inLink ? !a.inLink : a.text.length > b.text.length);
  const spans: Array<{ tags: Tags; model: Piece }> = [];
  for (const p of pieces) {
    if (!visible(p.text) || !p.run) continue;
    const last = spans[spans.length - 1];
    if (last && same(last.tags, p.tags)) {
      if (better(p, last.model)) last.model = p;
    } else spans.push({ tags: p.tags, model: p });
  }
  const plain = (spans.find((s) => same(s.tags, whole)) ?? spans[0])?.model;
  const used = new Map<string, number>();
  let previous: string | undefined;
  return segments.map((seg) => {
    const want = new Set([...seg.tags, ...whole]);
    const key = TAGS.filter((t) => want.has(t)).join(',');
    const matching = spans.filter((s) => same(s.tags, want));
    if (matching.length) {
      // Spaces between two parts of one phrase ("<b>A</b> <b>B</b>") do not start the next one.
      const advance = visible(seg.text) && key !== previous;
      if (advance) used.set(key, (used.get(key) ?? 0) + 1);
      if (visible(seg.text)) previous = key;
      const n = Math.max(1, used.get(key) ?? 0);
      return { text: seg.text, like: matching[Math.min(n, matching.length) - 1]!.model.run! };
    }
    if (visible(seg.text)) previous = key;
    const format = formatFor(want, plain?.tags ?? new Set());
    return { text: seg.text, ...(plain?.run ? { like: plain.run } : {}), ...(format ? { format } : {}) };
  });
}
