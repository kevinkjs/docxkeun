/**
 * What the user has selected, in terms an AI assistant can act on: paragraph numbers (the same
 * numbering as docxkeun-mcp's read_document), character offsets, pages, and the text with a
 * little context, like the "selected lines" an editor integration shares with a coding agent.
 */
import { paragraphText as plainText } from './editor.js';

/** Paragraph text with inline pictures as U+FFFC, the offsets the viewer's caret uses. */
const paragraphText = (p: DocParagraph) => plainText(p, { objects: true });
import type { LayoutResult, StoryKind } from '../layout/types.js';
import type { DocParagraph } from '../model/content.js';
import type { DocxDocument } from '../model/document.js';

/** A place in the text: a paragraph number (from 1) and a character offset in it. */
export interface TextSpot {
  paragraph: number;
  offset: number;
}

export interface SelectedParagraph {
  paragraph: number;
  /** The paragraph's whole text. */
  text: string;
  /** The selected part of `text`: [from, to). Equal for a caret. */
  from: number;
  to: number;
  /** Paragraph style name, when not the normal style. */
  style?: string;
}

export interface SelectionInfo {
  /** Which text the selection is in (paragraph numbers count within that story). */
  story: 'body' | StoryKind;
  /** The header, footer or notes part, for selections outside the body. */
  part?: string;
  /** Nothing selected, just a caret (or a click) at `start`. */
  collapsed: boolean;
  /** The selected text, paragraphs separated by newlines ('' for a caret). */
  text: string;
  start: TextSpot;
  end: TextSpot;
  paragraphs: SelectedParagraph[];
  /** Pages (from 1) the selection appears on. */
  pages: number[];
}

/**
 * A selection from paragraph positions (0-based indices into `paragraphs`, which are the story's
 * paragraphs in document order, table cells included).
 */
export function selectionInfoFor(
  doc: DocxDocument,
  paragraphs: DocParagraph[],
  anchor: { para: number; char: number },
  focus: { para: number; char: number },
  options: { story?: SelectionInfo['story']; part?: string; layout?: LayoutResult } = {},
): SelectionInfo | undefined {
  if (!paragraphs.length) return undefined;
  const clamp = (c: { para: number; char: number }) => {
    const para = Math.max(0, Math.min(c.para, paragraphs.length - 1));
    return { para, char: Math.max(0, Math.min(c.char, paragraphText(paragraphs[para]!).text.length)) };
  };
  let a = clamp(anchor);
  let b = clamp(focus);
  if (b.para < a.para || (b.para === a.para && b.char < a.char)) [a, b] = [b, a];
  const out: SelectedParagraph[] = [];
  const texts: string[] = [];
  for (let p = a.para; p <= b.para; p++) {
    const para = paragraphs[p]!;
    const text = paragraphText(para).text;
    const from = p === a.para ? a.char : 0;
    const to = p === b.para ? b.char : text.length;
    const style = para.styleId && para.styleId !== doc.styles.defaults.paragraph ? (doc.styles.get(para.styleId)?.name ?? para.styleId) : undefined;
    out.push({ paragraph: p + 1, text, from, to, ...(style ? { style } : {}) });
    texts.push(text.slice(from, to));
  }
  const story = options.story ?? 'body';
  return {
    story,
    ...(options.part ? { part: options.part } : {}),
    collapsed: a.para === b.para && a.char === b.char,
    text: texts.join('\n').replace(/￼/g, ''),
    start: { paragraph: a.para + 1, offset: a.char },
    end: { paragraph: b.para + 1, offset: b.char },
    paragraphs: out,
    pages: options.layout ? pagesOf(options.layout, paragraphs.slice(a.para, b.para + 1), story) : [],
  };
}

function pagesOf(layout: LayoutResult, paras: DocParagraph[], story: SelectionInfo['story']): number[] {
  const els = new Set(paras.map((p) => p.el));
  const pages: number[] = [];
  for (const page of layout.pages) {
    const lines = story === 'body' ? page.lines : page.storyLines;
    if (lines.some((l) => els.has(l.paragraph))) pages.push(page.index + 1);
  }
  return pages;
}

const range = (ns: number[]) => (ns.length === 0 ? '' : ns.length === 1 || ns[0] === ns[ns.length - 1] ? String(ns[0]) : `${ns[0]}–${ns[ns.length - 1]}`);

/**
 * The selection as a few lines of text for an AI prompt: where it is, and each paragraph with the
 * selected part in «guillemets» (or ‸ for a caret), shortened around the selection.
 */
export function describeSelection(info: SelectionInfo, options: { document?: string; maxChars?: number; context?: number } = {}): string {
  const max = options.maxChars ?? 2000;
  const around = options.context ?? 60;
  const where: string[] = [];
  const n = info.paragraphs.length;
  where.push(n > 1 ? `paragraphs ${info.start.paragraph}–${info.end.paragraph}` : `paragraph ${info.start.paragraph}`);
  if (info.story !== 'body') where.unshift(info.part ? `${info.story} (${info.part})` : info.story);
  if (info.pages.length) where.push(`page${info.pages.length > 1 ? 's' : ''} ${range(info.pages)}`);
  const doc = options.document ? ` in ${options.document}` : '';
  const head = info.collapsed
    ? `Cursor${doc}: ${where.join(', ')}, nothing selected.`
    : `Selected${doc}: ${where.join(', ')}, ${info.text.length} character${info.text.length === 1 ? '' : 's'}.`;
  let budget = max;
  const lines = info.paragraphs.map((p) => {
    const before = p.text.slice(Math.max(0, p.from - around), p.from);
    const after = p.text.slice(p.to, p.to + around);
    let chosen = p.text.slice(p.from, p.to);
    if (chosen.length > budget) chosen = `${chosen.slice(0, Math.max(0, budget))}… (${chosen.length - budget} more characters)`;
    budget = Math.max(0, budget - chosen.length);
    const lead = p.from - around > 0 ? '…' : '';
    const tail = p.to + around < p.text.length ? '…' : '';
    const mark = info.collapsed ? '‸' : `«${chosen}»`;
    const style = p.style ? ` (${p.style})` : '';
    return `[${p.paragraph}]${style} ${lead}${before}${mark}${after}${tail}`.replace(/￼/g, '').replace(/\s*\n\s*/g, ' ');
  });
  return [head, ...lines].join('\n');
}
