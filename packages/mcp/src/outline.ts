/**
 * A numbered, plain-text view of a document for AI clients: paragraphs (body, table cells and text
 * boxes) numbered from 1, with their style, table or text box, locks and comments.
 */
import { readComments, type Block, type CommentInfo, type DocParagraph, type DocxDocument, type DrawingContent } from 'docxkeun';
import type { ListItem, LockedRegion, SelectionInfo } from 'docxkeun/editor';
import { markupOf, wholeFormatting } from './markup.js';

type XmlElement = DocParagraph['el'];

export interface ParagraphEntry {
  /** 1-based paragraph number. */
  n: number;
  para: DocParagraph;
  /** The text editor.find() searches: text runs, tabs and line breaks. */
  text: string;
  cell?: { table: number; row: number; cell: number };
  /** In a text box: the box's number (from 1, in the order their paragraphs come) and the number of the paragraph holding it. */
  box?: { index: number; host: number };
  /**
   * The paragraph's number in a live viewer, whose activity effects count body and table-cell
   * paragraphs only (paragraphsOf); a text box's paragraphs have the number of the paragraph holding the box.
   */
  view: number;
}

export function paragraphTextOf(para: DocParagraph): string {
  let text = '';
  for (const inline of para.inlines) {
    if (inline.kind === 'text') text += inline.text;
    else if (inline.kind === 'tab') text += '\t';
    else if (inline.kind === 'break' && inline.type === 'line') text += '\n';
  }
  return text;
}

/**
 * Every paragraph in the body, in the order DocxEditor.paragraphs() lists them: table cells in
 * place, and a text box's paragraphs right after the paragraph holding the box.
 */
export function listParagraphs(doc: DocxDocument): ParagraphEntry[] {
  const out: ParagraphEntry[] = [];
  let tables = 0;
  let boxes = 0;
  let views = 0;
  type Where = { cell?: ParagraphEntry['cell']; box?: ParagraphEntry['box']; view?: number };
  const walk = (blocks: Block[], where: Where): void => {
    for (const b of blocks) {
      if (b.kind === 'paragraph') {
        const e: ParagraphEntry = { n: out.length + 1, para: b, text: paragraphTextOf(b), view: where.view ?? ++views };
        if (where.cell) e.cell = where.cell;
        if (where.box) e.box = where.box;
        out.push(e);
        for (const inline of b.inlines) if (inline.kind === 'image' && inline.content) visit(inline.content, e);
      } else if (where.box) {
        // A table in a text box is part of the box; table numbers count the body's tables.
        for (const row of b.rows) for (const c of row.cells) walk(c.blocks, where);
      } else {
        const table = ++tables;
        b.rows.forEach((row, r) => row.cells.forEach((c, i) => walk(c.blocks, { cell: { table, row: r + 1, cell: i + 1 } })));
      }
    }
  };
  // A text box (or a shape holding text), or the ones in a group, depth first.
  const visit = (content: DrawingContent, host: ParagraphEntry): void => {
    if (content.type === 'shape' && content.shape.text) walk(content.shape.text.blocks, { box: { index: ++boxes, host: host.n }, view: host.view });
    else if (content.type === 'group') for (const child of content.children) visit(child.content, host);
  };
  walk(doc.body().blocks, {});
  return out;
}

/** read_document's paragraph numbers as a live viewer numbers them (unique, sorted); numbers past the end are left out. */
export function viewerParagraphs(entries: ReadonlyArray<ParagraphEntry>, numbers: number[]): number[] {
  const out = new Set<number>();
  for (const n of numbers) {
    const e = entries[n - 1];
    if (e) out.add(e.view);
  }
  return [...out].sort((a, b) => a - b);
}

/** A selection in a live viewer's body (paragraphs numbered as the viewer numbers them), with read_document's paragraph numbers. */
export function renumberSelection(info: SelectionInfo, entries: ReadonlyArray<ParagraphEntry>): SelectionInfo {
  if (info.story !== 'body') return info;
  const byView = new Map<number, number>();
  for (const e of entries) if (!e.box) byView.set(e.view, e.n);
  const n = (view: number) => byView.get(view) ?? view;
  return {
    ...info,
    start: { ...info.start, paragraph: n(info.start.paragraph) },
    end: { ...info.end, paragraph: n(info.end.paragraph) },
    paragraphs: info.paragraphs.map((p) => ({ ...p, paragraph: n(p.paragraph) })),
  };
}

export function tableCount(doc: DocxDocument): number {
  let n = 0;
  const walk = (blocks: Block[]): void => {
    for (const b of blocks) {
      if (b.kind === 'table') {
        n++;
        for (const row of b.rows) for (const c of row.cells) walk(c.blocks);
      }
    }
  };
  walk(doc.body().blocks);
  return n;
}

// ---- styles ---------------------------------------------------------------------------------------

function displayName(name: string): string {
  return name.charAt(0).toUpperCase() + name.slice(1);
}

export function styleName(doc: DocxDocument, styleId: string | undefined): string | undefined {
  if (!styleId) return undefined;
  const def = doc.styles.styles.get(styleId);
  return def?.name ? displayName(def.name) : styleId;
}

export function paragraphStyles(doc: DocxDocument): Array<{ id: string; name: string }> {
  return [...doc.styles.styles.values()].filter((s) => s.type === 'paragraph').map((s) => ({ id: s.id, name: s.name ? displayName(s.name) : s.id }));
}

/** A paragraph style id from an id or a name ("Heading1", "Heading 1", "heading 1"). */
export function resolveStyle(doc: DocxDocument, style: string): string {
  const wanted = style.trim().toLowerCase();
  const compact = wanted.replace(/\s+/g, '');
  const styles = paragraphStyles(doc);
  const hit =
    styles.find((s) => s.id.toLowerCase() === wanted) ??
    styles.find((s) => s.name.toLowerCase() === wanted) ??
    styles.find((s) => s.id.toLowerCase() === compact || s.name.toLowerCase().replace(/\s+/g, '') === compact);
  if (!hit) throw new Error(`Unknown paragraph style "${style}". This document has: ${styles.map((s) => s.name).join(', ')}`);
  return hit.id;
}

// ---- locks and comments ---------------------------------------------------------------------------

function* descendants(el: XmlElement, name: string): Generator<XmlElement> {
  for (const child of el.children) {
    if (child.type !== 'element') continue;
    if (child.name === name) yield child;
    yield* descendants(child, name);
  }
}

/** Paragraph number → locked regions that cover it fully ('locked') or in part ('partly'). */
export function locksByParagraph(entries: ParagraphEntry[], regions: LockedRegion[]): Map<number, Array<{ region: LockedRegion; whole: boolean }>> {
  const out = new Map<number, Array<{ region: LockedRegion; whole: boolean }>>();
  if (regions.length === 0) return out;
  const byPara = new Map<XmlElement, number>(entries.map((e) => [e.para.el, e.n]));
  const sdtPara = new Map<XmlElement, number>();
  for (const e of entries) for (const sdt of descendants(e.para.el, 'w:sdt')) sdtPara.set(sdt, e.n);
  const add = (n: number, region: LockedRegion, whole: boolean): void => {
    const list = out.get(n) ?? [];
    list.push({ region, whole });
    out.set(n, list);
  };
  for (const region of regions) {
    if (region.block) {
      for (const p of descendants(region.el, 'w:p')) {
        const n = byPara.get(p);
        if (n !== undefined) add(n, region, true);
      }
    } else {
      const n = sdtPara.get(region.el);
      if (n !== undefined) add(n, region, false);
    }
  }
  return out;
}

export interface CommentAnchor {
  paragraph: number;
  text: string;
}

/** Where each comment is anchored in the body, and the text it covers. */
export function commentAnchors(entries: ParagraphEntry[]): Map<string, CommentAnchor> {
  const open = new Map<string, CommentAnchor>();
  const done = new Map<string, CommentAnchor>();
  for (const e of entries) {
    for (const inline of e.para.inlines) {
      if (inline.kind === 'commentStart') open.set(inline.id, { paragraph: e.n, text: '' });
      else if (inline.kind === 'commentEnd') {
        const a = open.get(inline.id);
        if (a) done.set(inline.id, a);
        open.delete(inline.id);
      } else if (inline.kind === 'commentRef') {
        // A comment without a range is anchored at its reference mark.
        if (!open.has(inline.id) && !done.has(inline.id)) done.set(inline.id, { paragraph: e.n, text: '' });
      } else {
        const t = inline.kind === 'text' ? inline.text : inline.kind === 'tab' ? '\t' : inline.kind === 'break' && inline.type === 'line' ? '\n' : '';
        if (t) for (const a of open.values()) a.text += t;
      }
    }
    // Ranges that run on into the next paragraph get a line break between paragraphs.
    for (const a of open.values()) if (a.text) a.text += '\n';
  }
  for (const [id, a] of open) done.set(id, a);
  return done;
}

export function commentsByParagraph(anchors: Map<string, CommentAnchor>): Map<number, string[]> {
  const out = new Map<number, string[]>();
  for (const [id, a] of anchors) {
    const list = out.get(a.paragraph) ?? [];
    list.push(id);
    out.set(a.paragraph, list);
  }
  return out;
}

export interface CommentThread {
  comment: CommentInfo;
  replies: CommentInfo[];
}

export function commentThreads(doc: DocxDocument): CommentThread[] {
  const all = readComments(doc);
  const byId = new Map(all.map((c) => [c.id, c]));
  const threads: CommentThread[] = [];
  const index = new Map<string, CommentThread>();
  for (const c of all) {
    let root = c;
    for (let guard = 0; root.parentId && byId.has(root.parentId) && guard < 100; guard++) root = byId.get(root.parentId)!;
    let thread = index.get(root.id);
    if (!thread) {
      thread = { comment: root, replies: [] };
      index.set(root.id, thread);
      threads.push(thread);
    }
    if (c !== root) thread.replies.push(c);
  }
  return threads;
}

// ---- formatting -----------------------------------------------------------------------------------

export function quote(text: string, max = 80): string {
  const flat = text.replace(/\s+/g, ' ').trim();
  return `"${flat.length > max ? flat.slice(0, max - 1) + '…' : flat}"`;
}

const ALIGNMENT: Record<string, string> = { center: 'centered', right: 'right-aligned', both: 'justified', distribute: 'justified' };

export function describeParagraph(
  doc: DocxDocument,
  e: ParagraphEntry,
  extras: { locks?: Array<{ region: LockedRegion; whole: boolean }>; comments?: string[]; list?: ListItem | undefined; objects?: string[]; revisions?: number } = {},
): string {
  const tags: string[] = [];
  const style = e.para.styleId && e.para.styleId !== doc.styles.defaults.paragraph ? styleName(doc, e.para.styleId) : undefined;
  if (style) tags.push(style);
  const list = extras.list;
  if (list) tags.push(`${list.format === 'bullet' ? 'bullet' : 'list item'}${list.label ? ' ' + list.label : ''}${list.level ? `, level ${list.level + 1}` : ''}`);
  else if (e.para.numbering) tags.push(`list level ${e.para.numbering.ilvl + 1}`);
  if (e.cell) tags.push(`table ${e.cell.table}, row ${e.cell.row}, cell ${e.cell.cell}`);
  if (e.box) tags.push(`text box ${e.box.index}`);
  const align = ALIGNMENT[e.para.props.jc ?? ''];
  if (align) tags.push(align);
  const whole = e.text.trim() ? wholeFormatting(e.para) : [];
  if (whole.length) tags.push(whole.join(', '));
  for (const { region, whole } of extras.locks ?? []) {
    const name = region.title ? ` "${region.title}"` : '';
    tags.push(`${whole ? 'locked' : 'partly locked'}${name}, lock ${region.id}`);
  }
  if (extras.comments?.length) tags.push(`comment${extras.comments.length > 1 ? 's' : ''} ${extras.comments.join(', ')}`);
  if (e.para.props.bidi) tags.push('right-to-left');
  if (extras.revisions) tags.push(`${extras.revisions} tracked change${extras.revisions > 1 ? 's' : ''}`);
  const objects = extras.objects?.length ? ` ${extras.objects.map((o) => `{${o}}`).join(' ')}` : '';
  const text = e.text ? markupOf(e.para) : '';
  return `[${e.n}]${tags.length ? ` (${tags.join('; ')})` : ''}${text ? ' ' + text : ''}${objects}`;
}

export function wordCount(entries: ParagraphEntry[]): number {
  let n = 0;
  for (const e of entries) n += e.text.split(/\s+/).filter(Boolean).length;
  return n;
}
