/**
 * The document operations behind the tools: reading, finding, editing, lists, comments, links and
 * locks, on one open document. The MCP server and the docxkeun chat web app share them, so this
 * module uses no Node APIs. Each operation has a zod schema for its input (without file arguments),
 * the minimum role that may use it, and whether it reads, adds or changes.
 */
import { z } from 'zod';
import { blocksToText, describeChart, readChart, type DocParagraph, type DocxDocument, type ImageInline } from 'docxkeun';
import { DocxEditor as Editor, paragraphText, paragraphsOf, type DocObject, type DocRange, type DocxEditor, type Match, type ParagraphFormat, type Revision, type TextFormat } from 'docxkeun/editor';
import { readProtection, type Role } from 'docxkeun/security';
import {
  commentAnchors,
  commentsByParagraph,
  commentThreads,
  describeParagraph,
  listParagraphs,
  locksByParagraph,
  quote,
  resolveStyle,
  tableCount,
  wordCount,
  type ParagraphEntry,
} from './outline.js';
import { formatFor, hasMarkup, parseMarkup, piecesFor, plainText, type Segment } from './markup.js';

/** One open document and who is acting on it. */
export interface DocContext {
  doc: DocxDocument;
  /** An editor guarded by the caller's AccessPolicy. */
  editor: DocxEditor;
  /** The document's name, for messages. */
  name: string;
  /** The name comments are signed with. */
  author: string;
  /** Paragraphs numbered as the operation starts (listParagraphs(doc)). */
  entries: ParagraphEntry[];
  /** Set by operations: where a live view should scroll afterwards. */
  focus?: { comment?: string };
  /** Reads a picture file the caller may use (the MCP server's allowed folders); absent where files cannot be read. */
  loadImage?: (path: string) => Promise<Uint8Array>;
}

export type OperationKind = 'read' | 'add' | 'change';

export interface OperationEnv {
  author: string;
  role: Role;
}

export interface Operation<S extends z.ZodObject = z.ZodObject> {
  name: string;
  title: string;
  description: string | ((env: OperationEnv) => string);
  /** The least role that may use it. */
  role: Role;
  kind: OperationKind;
  /** Input, without file arguments (path, password, output). */
  schema: S;
  /** Returns a short report of what happened; throws when nothing could be done. */
  run(ctx: DocContext, args: z.output<S>): Promise<string> | string;
}

function define<S extends z.ZodObject>(op: Operation<S>): Operation {
  return op as unknown as Operation;
}

export function describeOperation(op: Operation, env: OperationEnv): string {
  return typeof op.description === 'function' ? op.description(env) : op.description;
}

// ---- helpers ----------------------------------------------------------------------------------------

export function plural(n: number, word: string, many = word + 's'): string {
  return `${n} ${n === 1 ? word : many}`;
}

export function initialsOf(name: string): string {
  return (
    name
      .split(/\s+/)
      .filter(Boolean)
      .map((w) => w[0]!.toUpperCase())
      .join('')
      .slice(0, 3) || 'AI'
  );
}

export function escapeRegExp(s: string): string {
  return s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

export function searchPattern(query: string, options: { regex?: boolean | undefined; matchCase?: boolean | undefined }): string | RegExp {
  const matchCase = options.matchCase ?? true;
  // read_document's formatting tags are not part of the text.
  if (!options.regex && hasMarkup(query)) query = plainText(query);
  if (!options.regex) return matchCase ? query : new RegExp(escapeRegExp(query), 'gi');
  try {
    return new RegExp(query, matchCase ? 'g' : 'gi');
  } catch (e) {
    throw new Error(`Invalid regular expression: ${(e as Error).message}`);
  }
}

/** $& (the match), $1…$9 (groups) and $$ in a regex replacement. */
function expandReplacement(template: string, m: Match): string {
  return template.replace(/\$(\$|&|\d)/g, (_, k: string) => (k === '$' ? '$' : k === '&' ? m.text : (m.groups[Number(k) - 1] ?? '')));
}

function entry(ctx: DocContext, n: number): ParagraphEntry {
  const e = ctx.entries[n - 1];
  if (!e) throw new Error(`There is no paragraph ${n}; the document has ${plural(ctx.entries.length, 'paragraph')}.`);
  return e;
}

function matchesIn(ctx: DocContext, pattern: string | RegExp, paragraph?: number): Match[] {
  const all = ctx.editor.find(pattern);
  if (paragraph === undefined) return all;
  const el = entry(ctx, paragraph).para.el;
  return all.filter((m) => m.paragraph === el);
}

function paragraphOf(ctx: DocContext, m: Match): number {
  return ctx.entries.find((e) => e.para.el === m.paragraph)?.n ?? 0;
}

/** The range an operation acts on, from a paragraph number and/or a piece of text. */
function resolveTarget(ctx: DocContext, t: { paragraph?: number | undefined; text?: string | undefined; occurrence?: number | undefined }): { range: DocRange; label: string } {
  if (t.text !== undefined) {
    const matches = matchesIn(ctx, t.text, t.paragraph);
    const where = t.paragraph ? ` in paragraph ${t.paragraph}` : '';
    if (matches.length === 0) throw new Error(`Text not found${where}: ${quote(t.text)}. Text must match exactly, including case.`);
    const k = t.occurrence ?? 1;
    const m = matches[k - 1];
    if (!m) throw new Error(`Only ${plural(matches.length, 'match', 'matches')} of ${quote(t.text)}${where}.`);
    const which = matches.length > 1 ? ` (match ${k} of ${matches.length})` : '';
    return { range: m, label: `${quote(t.text)} in paragraph ${paragraphOf(ctx, m)}${which}` };
  }
  if (t.paragraph !== undefined) {
    const e = entry(ctx, t.paragraph);
    const range = ctx.editor.paragraphRange(e.para);
    if (!range) throw new Error(`Paragraph ${t.paragraph} is empty; there is no text to act on.`);
    return { range, label: `paragraph ${t.paragraph}` };
  }
  throw new Error('Say what to act on: pass `paragraph`, `text`, or both.');
}

function newNumbers(ctx: DocContext, els: Array<ParagraphEntry['para']['el']>): string {
  const entries = listParagraphs(ctx.doc);
  const ns = els.map((el) => entries.find((e) => e.para.el === el)?.n).filter((n): n is number => n !== undefined);
  if (ns.length === 0) return '';
  return ns.length === 1 ? `paragraph ${ns[0]}` : `paragraphs ${ns[0]}–${ns[ns.length - 1]}`;
}

// ---- shared input fields ----------------------------------------------------------------------------

export const paragraphNumber = z.number().int().min(1);

const target = {
  paragraph: paragraphNumber.optional().describe('Paragraph number from read_document'),
  text: z.string().min(1).optional().describe('Exact text to act on (case-sensitive). With `paragraph`, only that paragraph is searched'),
  occurrence: z.number().int().min(1).optional().describe('Which match of `text` to use, counting from 1 (default 1)'),
};

export const newParagraph = z.object({
  text: z.string().describe('Paragraph text; "\\n" makes a line break and "\\t" a tab. <b>, <i>, <u>, <s>, <sup> and <sub> tags format parts of it'),
  style: z.string().optional().describe('Paragraph style name or id, e.g. "Heading 1" (default: the normal style)'),
});

const textFormat = z
  .object({
    bold: z.boolean().optional(),
    italic: z.boolean().optional(),
    underline: z.boolean().optional(),
    strike: z.boolean().optional(),
    color: z
      .string()
      .regex(/^#?[0-9a-fA-F]{6}$/)
      .optional()
      .describe('Text colour as hex RGB, e.g. "C00000"'),
    highlight: z
      .enum(['yellow', 'green', 'cyan', 'magenta', 'blue', 'red', 'darkBlue', 'darkCyan', 'darkGreen', 'darkMagenta', 'darkRed', 'darkYellow', 'darkGray', 'lightGray', 'black', 'white', 'none'])
      .optional(),
    size: z.number().min(1).max(1638).optional().describe('Font size in points'),
    font: z.string().optional().describe('Font family name'),
  })
  .describe('Formatting to apply; leave out what should not change, false turns a property off');

// ---- reading ----------------------------------------------------------------------------------------

const readDocument = define({
  name: 'read_document',
  title: 'Read a document',
  description:
    'Read a Word document as numbered paragraphs. Each line is "[number] (style; list label; table position or text box; alignment; formatting of the whole paragraph; locks; comments) text {objects}", ' +
    'e.g. "[1] (centered; bold, underlined) AGREEMENT" or "[4] (list item a.; justified) …". ' +
    'Text formatted differently from the rest of its paragraph is tagged: <b>bold</b>, <i>italic</i>, <u>underlined</u>, <s>struck through</s>, <sup>superscript</sup>, <sub>subscript</sub>. ' +
    'The tags are not part of the text (find_text and replace_text ignore them); edit_paragraph and insert_paragraphs take them, so keep them when you rewrite or translate text. ' +
    "A text box's paragraphs come right after the paragraph holding it, tagged \"text box N\". Pictures, charts and shapes show as {object N: …}. " +
    'Long documents come in pages of paragraphs: pass `from` to continue.',
  role: 'viewer',
  kind: 'read',
  schema: z.object({
    from: paragraphNumber.optional().describe('First paragraph to show (default 1)'),
    count: z.number().int().min(1).max(2000).optional().describe('How many paragraphs to show at most (default 300)'),
  }),
  run(ctx, args) {
    const total = ctx.entries.length;
    // Replies share their thread's anchor, so only thread starts are shown on paragraphs.
    const threads = commentThreads(ctx.doc);
    const roots = new Set(threads.map((t) => t.comment.id));
    const anchors = new Map([...commentAnchors(ctx.entries)].filter(([id]) => roots.has(id)));
    const comments = commentsByParagraph(anchors);
    const locks = locksByParagraph(ctx.entries, ctx.editor.lockedRegions());
    const lists = ctx.editor.listItems();
    const numbered = new Map(numberedObjects(ctx.editor, ctx.entries).map((o) => [o.image, o]));
    const protection = readProtection(ctx.doc);
    const revisions = new Map<number, number>();
    for (const r of ctx.editor.revisions()) {
      const n = revisionParagraph(ctx, r);
      if (n) revisions.set(n, (revisions.get(n) ?? 0) + 1);
    }
    const head = [
      `${ctx.name}: ${plural(total, 'paragraph')}, ${plural(tableCount(ctx.doc), 'table')}, ${plural(wordCount(ctx.entries), 'word')}, ${plural(threads.length, 'comment thread')}.`,
    ];
    if (protection.enforced && protection.mode !== 'none') head.push(`Protection: ${protection.mode}${protection.hasPassword ? ' (with a password)' : ''}.`);
    if (ctx.editor.trackChanges) head.push('Track Changes is on: your edits are recorded as tracked changes.');
    if (revisions.size) head.push(`The text is shown with tracked changes accepted; list_revisions lists them.`);
    if (ctx.doc.storyParts().length) head.push('Headers, footers and notes: read_headers_footers.');
    const from = Math.min(args.from ?? 1, Math.max(total, 1));
    const limit = args.count ?? 300;
    const lines: string[] = [];
    let size = 0;
    let last = from - 1;
    for (const e of ctx.entries.slice(from - 1, from - 1 + limit)) {
      const objects = e.para.inlines.flatMap((i) => {
        const o = i.kind === 'image' ? numbered.get(i) : undefined;
        return o ? [objectLabel(ctx, o)] : i.kind === 'math' ? [`equation: ${i.text}`] : [];
      });
      const line = describeParagraph(ctx.doc, e, { locks: locks.get(e.n) ?? [], comments: comments.get(e.n) ?? [], list: lists.get(e.para.el), objects, revisions: revisions.get(e.n) ?? 0 });
      if (size + line.length > 60_000 && lines.length > 0) break;
      lines.push(line);
      size += line.length + 1;
      last = e.n;
    }
    const tail = last < total ? [`(Paragraphs ${from}–${last} of ${total}. Call read_document with from: ${last + 1} for more.)`] : [];
    return [...head, '', ...lines, ...(tail.length ? ['', ...tail] : [])].join('\n');
  },
});

const findText = define({
  name: 'find_text',
  title: 'Find text',
  description: 'Find text (or a regular expression) in a document and show each match with its paragraph number and surrounding text.',
  role: 'viewer',
  kind: 'read',
  schema: z.object({
    query: z.string().min(1).describe('Text to look for'),
    regex: z.boolean().optional().describe('Treat query as a JavaScript regular expression'),
    matchCase: z.boolean().optional().describe('Match upper/lower case exactly (default true)'),
  }),
  run(ctx, args) {
    const pattern = searchPattern(args.query, args);
    const re = typeof pattern === 'string' ? new RegExp(escapeRegExp(pattern), 'g') : pattern;
    const lines: string[] = [];
    let total = 0;
    for (const e of ctx.entries) {
      re.lastIndex = 0;
      for (let m = re.exec(e.text); m; m = re.exec(e.text)) {
        if (m[0].length === 0) {
          re.lastIndex++;
          continue;
        }
        total++;
        if (lines.length < 100) {
          const start = Math.max(0, m.index - 50);
          const end = Math.min(e.text.length, m.index + m[0].length + 50);
          const before = (start > 0 ? '…' : '') + e.text.slice(start, m.index);
          const after = e.text.slice(m.index + m[0].length, end) + (end < e.text.length ? '…' : '');
          lines.push(`[${e.n}] ${before}«${m[0]}»${after}`.replace(/\s+/g, ' '));
        }
      }
    }
    if (total === 0) return `No matches for ${quote(args.query)}.`;
    return [`${plural(total, 'match', 'matches')} for ${quote(args.query)}${total > lines.length ? ` (showing ${lines.length})` : ''}:`, ...lines].join('\n');
  },
});

const listComments = define({
  name: 'list_comments',
  title: 'List comments',
  description: 'List comment threads with their ids, authors, dates, the text they are attached to, and replies.',
  role: 'viewer',
  kind: 'read',
  schema: z.object({
    includeResolved: z.boolean().optional().describe('Include resolved threads (default true)'),
  }),
  run(ctx, args) {
    const anchors = commentAnchors(ctx.entries);
    const threads = commentThreads(ctx.doc).filter((t) => args.includeResolved !== false || !t.comment.resolved);
    if (threads.length === 0) return 'No comments.';
    // dateUtc (commentsExtensible) is real UTC; Word writes w:date in local time despite its "Z".
    const date = (c: { date?: string | undefined; dateUtc?: string | undefined }): string =>
      c.dateUtc ? `, ${c.dateUtc.slice(0, 16).replace('T', ' ')} UTC` : c.date ? `, ${c.date.slice(0, 16).replace('T', ' ')}` : '';
    const out: string[] = [];
    for (const { comment: c, replies } of threads) {
      const a = anchors.get(c.id);
      out.push(`Comment ${c.id} by ${c.author}${date(c)}${c.resolved ? ' (resolved)' : ''}`);
      if (a) out.push(`  On paragraph ${a.paragraph}${a.text ? `: ${quote(a.text)}` : ''}`);
      out.push(`  ${c.text.replace(/\n/g, '\n  ')}`);
      for (const r of replies) out.push(`  Reply ${r.id} by ${r.author}${date(r)}: ${r.text.replace(/\n/g, ' ')}`);
      out.push('');
    }
    return out.join('\n').trimEnd();
  },
});

const listLinks = define({
  name: 'list_links',
  title: 'List links and bookmarks',
  description: 'List the hyperlinks (to web addresses or to bookmarks) and the bookmarks in a document.',
  role: 'viewer',
  kind: 'read',
  schema: z.object({}),
  run(ctx) {
    const out: string[] = [];
    const links = ctx.editor.hyperlinks();
    out.push(links.length ? `Hyperlinks (${links.length}):` : 'No hyperlinks.');
    for (const l of links) out.push(`  ${quote(l.text, 60)} → ${l.url ?? `bookmark ${l.anchor}`}${l.tooltip ? ` (tooltip: ${l.tooltip})` : ''}`);
    const where = new Map<string, number>();
    for (const e of ctx.entries) for (const i of e.para.inlines) if (i.kind === 'bookmarkStart') where.set(i.name, e.n);
    const bookmarks = ctx.editor.bookmarks().filter((b) => !b.name.startsWith('_'));
    out.push(bookmarks.length ? `Bookmarks (${bookmarks.length}):` : 'No bookmarks.');
    for (const b of bookmarks) out.push(`  ${b.name}${where.has(b.name) ? ` (paragraph ${where.get(b.name)})` : ''}`);
    return out.join('\n');
  },
});

// ---- comments ---------------------------------------------------------------------------------------

const addComment = define({
  name: 'add_comment',
  title: 'Add a comment',
  description: (env) => `Attach a comment to text in a document (signed "${env.author}"). Target a whole paragraph, a piece of text, or text within a paragraph.`,
  role: 'commenter',
  kind: 'add',
  schema: z.object({ ...target, comment: z.string().min(1).describe('The comment') }),
  async run(ctx, args) {
    const { range, label } = resolveTarget(ctx, args);
    const id = await ctx.editor.addComment(range, { author: ctx.author, initials: initialsOf(ctx.author), text: args.comment });
    ctx.focus = { comment: id };
    return `Added comment ${id} on ${label}.`;
  },
});

const replyToComment = define({
  name: 'reply_to_comment',
  title: 'Reply to a comment',
  description: 'Reply to a comment thread (ids come from list_comments).',
  role: 'commenter',
  kind: 'add',
  schema: z.object({ commentId: z.string().min(1), reply: z.string().min(1).describe('The reply') }),
  async run(ctx, args) {
    const id = await ctx.editor.replyToComment(args.commentId, { author: ctx.author, initials: initialsOf(ctx.author), text: args.reply });
    ctx.focus = { comment: args.commentId };
    return `Added reply ${id} to comment ${args.commentId}.`;
  },
});

const resolveComment = define({
  name: 'resolve_comment',
  title: 'Resolve a comment',
  description: 'Mark a comment thread as resolved (done), or reopen it.',
  role: 'commenter',
  kind: 'change',
  schema: z.object({ commentId: z.string().min(1), resolved: z.boolean().optional().describe('false reopens the thread (default true)') }),
  async run(ctx, args) {
    await ctx.editor.resolveComment(args.commentId, args.resolved ?? true);
    return `Comment ${args.commentId} ${args.resolved === false ? 'reopened' : 'resolved'}.`;
  },
});

const deleteComment = define({
  name: 'delete_comment',
  title: 'Delete a comment',
  description: (env) => (env.role === 'commenter' ? `Delete a comment you added (signed "${env.author}").` : 'Delete a comment and its replies.'),
  role: 'commenter',
  kind: 'change',
  schema: z.object({ commentId: z.string().min(1) }),
  async run(ctx, args) {
    await ctx.editor.deleteComment(args.commentId);
    return `Deleted comment ${args.commentId}.`;
  },
});

// ---- editing ----------------------------------------------------------------------------------------

const replaceText = define({
  name: 'replace_text',
  title: 'Replace text',
  description:
    'Replace text everywhere in a document (or one occurrence), keeping the formatting of the replaced text. ' +
    'Matches inside locked regions are skipped. With regex, the replacement may use $1…$9 and $&.',
  role: 'editor',
  kind: 'change',
  schema: z.object({
    find: z.string().min(1).describe('Text to replace'),
    replace: z.string().describe('Replacement text (empty deletes the matches); <b>, <i>, <u>… tags add formatting to parts of it'),
    regex: z.boolean().optional().describe('Treat find as a JavaScript regular expression'),
    matchCase: z.boolean().optional().describe('Match upper/lower case exactly (default true)'),
    paragraph: paragraphNumber.optional().describe('Only replace in this paragraph'),
    occurrence: z.number().int().min(1).optional().describe('Only replace this match, counting from 1 (default: all)'),
  }),
  run(ctx, args) {
    const all = matchesIn(ctx, searchPattern(args.find, args), args.paragraph);
    if (all.length === 0) throw new Error(`No matches for ${quote(args.find)}${args.paragraph ? ` in paragraph ${args.paragraph}` : ''}; nothing changed.`);
    let chosen = all;
    if (args.occurrence !== undefined) {
      const m = all[args.occurrence - 1];
      if (!m) throw new Error(`Only ${plural(all.length, 'match', 'matches')} of ${quote(args.find)}; nothing changed.`);
      chosen = [m];
    }
    let skipped = 0;
    for (const m of [...chosen].reverse()) {
      if (ctx.editor.isLocked(m)) {
        skipped++;
        continue;
      }
      const replacement = args.regex ? expandReplacement(args.replace, m) : args.replace;
      // Untagged text keeps the formatting of what it replaces; tags add to it.
      ctx.editor.replaceRange(m, hasMarkup(replacement) ? parseMarkup(replacement).map((x) => ({ text: x.text, ...(x.tags.size ? { format: formatFor(x.tags, new Set())! } : {}) })) : replacement);
    }
    const done = chosen.length - skipped;
    if (done === 0) throw new Error(`${plural(skipped, 'match is', 'matches are')} inside locked regions; nothing changed.`);
    return `Replaced ${done} of ${plural(all.length, 'match', 'matches')}${skipped ? `; ${skipped} skipped (locked)` : ''}.`;
  },
});

const fillTemplate = define({
  name: 'fill_template',
  title: 'Fill a template',
  description: 'Fill {{placeholder}} fields and content controls (by tag or title) with values. Placeholders inside locked regions are left alone.',
  role: 'editor',
  kind: 'change',
  schema: z.object({
    data: z.record(z.string(), z.union([z.string(), z.number(), z.boolean(), z.null()])).describe('Placeholder name → value, e.g. {"client": "ACME"}'),
  }),
  run(ctx, args) {
    const placeholder = /\{\{\s*([\w.-]+)\s*\}\}/g;
    const used = ctx.editor.fillTemplate(args.data);
    const remaining = new Set<string>();
    for (const e of listParagraphs(ctx.doc)) for (const m of e.text.matchAll(placeholder)) remaining.add(m[1]!);
    if (used.length === 0) {
      throw new Error(
        `None of the keys matched a placeholder or content control; nothing changed.${remaining.size ? ` The document has: ${[...remaining].map((k) => `{{${k}}}`).join(', ')}` : ''}`,
      );
    }
    const unused = Object.keys(args.data).filter((k) => !used.includes(k));
    return [
      `Filled ${used.join(', ')}.`,
      ...(unused.length ? [`Not found in the document: ${unused.join(', ')}.`] : []),
      ...(remaining.size ? [`Still unfilled: ${[...remaining].map((k) => `{{${k}}}`).join(', ')}.`] : []),
    ].join('\n');
  },
});

const editParagraph = define({
  name: 'edit_paragraph',
  title: 'Edit a paragraph',
  description:
    "Replace a paragraph's text and/or change its style. The paragraph keeps its own formatting (alignment, indents, list, whole-paragraph bold and so on). " +
    'Use the tags read_document shows (<b>, <i>, <u>, <s>, <sup>, <sub>) for text formatted differently from the rest: each tagged part takes the formatting ' +
    '(font, size and colour too) of the text that had the same tags. Keep the tags when rewriting or translating, or that formatting is lost.',
  role: 'editor',
  kind: 'change',
  schema: z.object({
    paragraph: paragraphNumber.describe('Paragraph number from read_document'),
    text: z.string().optional().describe('New text for the whole paragraph, with formatting tags as read_document shows them'),
    style: z.string().optional().describe('New paragraph style name or id, e.g. "Heading 2"'),
  }),
  run(ctx, args) {
    if (args.text === undefined && args.style === undefined) throw new Error('Pass text, style, or both.');
    const e = entry(ctx, args.paragraph);
    const done: string[] = [];
    if (args.style !== undefined) {
      const id = resolveStyle(ctx.doc, args.style);
      ctx.editor.setParagraphStyle(e.para.el, id);
      done.push(`style set to ${args.style}`);
    }
    if (args.text !== undefined) {
      const segments = parseMarkup(args.text);
      const range = ctx.editor.paragraphRange(e.para);
      if (range) ctx.editor.replaceRange(range, piecesFor(e.para, segments));
      else if (args.text) writeTagged(ctx.editor, e.para.el, segments);
      done.push('text replaced');
    }
    return `Paragraph ${args.paragraph}: ${done.join(', ')}.`;
  },
});

const insertParagraphs = define({
  name: 'insert_paragraphs',
  title: 'Insert paragraphs',
  description:
    'Insert new paragraphs before or after a paragraph (default: at the end of the document). Inserting next to a paragraph in a table cell adds to that cell. ' +
    'Paragraph numbers after the insertion point shift.',
  role: 'editor',
  kind: 'add',
  schema: z.object({
    paragraphs: z.array(newParagraph).min(1).describe('The paragraphs to insert, in order'),
    after: paragraphNumber.optional().describe('Insert after this paragraph'),
    before: paragraphNumber.optional().describe('Insert before this paragraph'),
  }),
  run(ctx, args) {
    if (args.after !== undefined && args.before !== undefined) throw new Error('Pass after or before, not both.');
    const styles = args.paragraphs.map((p) => (p.style ? resolveStyle(ctx.doc, p.style) : undefined));
    const els: Array<ParagraphEntry['para']['el']> = [];
    args.paragraphs.forEach((p, i) => {
      const style = styles[i];
      const opts = style ? { style } : {};
      const where =
        i > 0
          ? { after: els[i - 1]! }
          : args.after !== undefined
            ? { after: entry(ctx, args.after).para.el }
            : args.before !== undefined
              ? { before: entry(ctx, args.before).para.el }
              : ('end' as const);
      const segments = parseMarkup(p.text);
      const el = ctx.editor.insertParagraph(where, segments.map((x) => x.text).join(''), opts);
      tagParagraph(ctx.editor, el, segments);
      els.push(el);
    });
    return `Inserted ${plural(els.length, 'paragraph')} (now ${newNumbers(ctx, els)}).`;
  },
});

const deleteParagraphs = define({
  name: 'delete_paragraphs',
  title: 'Delete paragraphs',
  description: 'Delete a paragraph or a run of paragraphs. The last paragraph of a table cell is emptied instead of removed. Paragraph numbers after them shift.',
  role: 'editor',
  kind: 'change',
  schema: z.object({
    from: paragraphNumber.describe('First paragraph to delete'),
    to: paragraphNumber.optional().describe('Last paragraph to delete (default: same as from)'),
  }),
  run(ctx, args) {
    const to = args.to ?? args.from;
    if (to < args.from) throw new Error('`to` must not be before `from`.');
    const doomed = [];
    for (let n = args.from; n <= to; n++) doomed.push(entry(ctx, n));
    const inBody = (e: ParagraphEntry) => !e.cell && !e.box;
    if (doomed.filter(inBody).length >= ctx.entries.filter(inBody).length) {
      throw new Error('A document needs at least one paragraph outside tables; clear the last one with edit_paragraph instead.');
    }
    for (const e of doomed.reverse()) ctx.editor.deleteParagraph(e.para.el);
    return `Deleted ${to === args.from ? `paragraph ${args.from}` : `paragraphs ${args.from}–${to}`}.`;
  },
});

/**
 * Where new body content goes: after or before a paragraph (its table, if it is in one), or at the
 * end. A paragraph in a text box is refused with `inBox`, since body content can't go there.
 */
function placeOf(ctx: DocContext, args: { after?: number; before?: number }, inBox: string): { where: { after?: Element; before?: Element } | 'end'; said: string } {
  if (args.after !== undefined && args.before !== undefined) throw new Error('Pass after or before, not both.');
  if (args.after !== undefined) {
    const e = bodyEntry(ctx, args.after, inBox);
    return { where: { after: e.para.el }, said: e.cell ? `after table ${e.cell.table}` : `after paragraph ${args.after}` };
  }
  if (args.before !== undefined) {
    const e = bodyEntry(ctx, args.before, inBox);
    return { where: { before: e.para.el }, said: e.cell ? `before table ${e.cell.table}` : `before paragraph ${args.before}` };
  }
  return { where: 'end', said: 'at the end' };
}

/** A paragraph that is not in a text box; `inBox` says why one is refused. */
function bodyEntry(ctx: DocContext, n: number, inBox: string): ParagraphEntry {
  const e = entry(ctx, n);
  if (e.box) throw new Error(`Paragraph ${n} is in a text box: ${inBox}.`);
  return e;
}
type Element = ParagraphEntry['para']['el'];

/** The descendants of an element with a name (document order). */
function findAll(node: Element, name: string): Element[] {
  const out: Element[] = [];
  for (const c of node.children) {
    if (c.type !== 'element') continue;
    if (c.name === name) out.push(c);
    out.push(...findAll(c, name));
  }
  return out;
}

/** Table `n` (1-based), counted as read_document counts them. */
function tableAt(ctx: DocContext, n: number): Element {
  const tables: Element[] = [];
  const walk = (blocks: ReturnType<DocxDocument['body']>['blocks']): void => {
    for (const b of blocks) {
      if (b.kind !== 'table') continue;
      tables.push(b.el);
      for (const row of b.rows) for (const c of row.cells) walk(c.blocks);
    }
  };
  walk(ctx.doc.body().blocks);
  const t = tables[n - 1];
  if (!t) throw new Error(`There is no table ${n}; the document has ${plural(tables.length, 'table')}.`);
  return t;
}

const copyParagraphs = define({
  name: 'copy_paragraphs',
  title: 'Copy paragraphs',
  description:
    'Copy paragraphs, with their formatting, lists, tables and pictures, to another place (default: the end of the document). ' +
    'A table the range starts or ends in is copied whole. Use it to duplicate content, for example to translate a document on a new page: ' +
    'copy everything with pageBreak, then change the copies’ text with edit_paragraph or replace_text (table cells are paragraphs too). ' +
    'Lists in the copy number from 1 again unless restartLists is false. Comments, bookmarks and footnote references are not copied. Paragraph numbers after the copy shift.',
  role: 'editor',
  kind: 'add',
  schema: z.object({
    from: paragraphNumber.describe('First paragraph to copy'),
    to: paragraphNumber.optional().describe('Last paragraph to copy (default: same as from)'),
    after: paragraphNumber.optional().describe('Put the copy after this paragraph (after its table, if it is in one). Default: at the end'),
    before: paragraphNumber.optional().describe('Put the copy before this paragraph (before its table, if it is in one)'),
    pageBreak: z.boolean().optional().describe('Start the copy on a new page'),
    restartLists: z.boolean().optional().describe('Number lists in the copy from 1 again (default true); false continues the original lists'),
  }),
  async run(ctx, args) {
    const to = args.to ?? args.from;
    if (to < args.from) throw new Error('`to` must not be before `from`.');
    const inBox = 'copy_paragraphs copies body paragraphs and tables; to put text in a text box use insert_paragraphs or edit_paragraph';
    const first = bodyEntry(ctx, args.from, inBox);
    const last = bodyEntry(ctx, to, inBox);
    const { where, said } = placeOf(ctx, args, inBox);
    const { blocks, dropped } = await ctx.editor.copyBlocks(first.para.el, last.para.el, where, {
      ...(args.pageBreak !== undefined ? { pageBreak: args.pageBreak } : {}),
      ...(args.restartLists !== undefined ? { restartLists: args.restartLists } : {}),
    });
    const paragraphs = blocks.flatMap((b) => (b.name === 'w:p' ? [b] : findAll(b, 'w:p')));
    const tables = blocks.filter((b) => b.name === 'w:tbl').length;
    const left = [
      dropped.comments ? plural(dropped.comments, 'comment anchor') : '',
      dropped.bookmarks ? plural(dropped.bookmarks, 'bookmark') : '',
      dropped.notes ? plural(dropped.notes, 'footnote reference') : '',
    ].filter(Boolean);
    return (
      `Copied ${to === args.from ? `paragraph ${args.from}` : `paragraphs ${args.from}–${to}`}${tables ? ` (with ${plural(tables, 'table')})` : ''} ${said}` +
      `${args.pageBreak ? ', starting on a new page' : ''}: the copy is ${newNumbers(ctx, paragraphs)}.` +
      (args.restartLists === false ? '' : ' Lists in the copy start again at 1.') +
      (left.length ? ` Not copied: ${left.join(', ')}.` : '')
    );
  },
});

const insertTableOp = define({
  name: 'insert_table',
  title: 'Insert a table',
  description:
    'Insert a table of text before or after a paragraph (default: at the end of the document). It spans the text width, with grid lines unless borders is false. ' +
    'Its cells become paragraphs in read_document (“table N, row R, cell C”), so their text can be changed and formatted like any paragraph. Paragraph numbers after it shift.',
  role: 'editor',
  kind: 'add',
  schema: z.object({
    rows: z.array(z.array(z.string()).min(1)).min(1).max(500).describe('The rows, each a list of cell texts; "\\n" in a cell starts a new line. Short rows get empty cells'),
    header: z.boolean().optional().describe('The first row is a header: bold, and repeated at the top of each page the table runs onto'),
    widths: z.array(z.number().positive()).optional().describe('Relative column widths, one per column (e.g. [1, 4, 2]); default: equal'),
    borders: z.boolean().optional().describe('Grid lines (default true)'),
    align: z.enum(['left', 'center', 'right']).optional().describe('Position between the margins (default left)'),
    after: paragraphNumber.optional().describe('Insert after this paragraph (after its table, if it is in one). Default: at the end'),
    before: paragraphNumber.optional().describe('Insert before this paragraph (before its table, if it is in one)'),
  }),
  run(ctx, args) {
    const { where, said } = placeOf(ctx, args, 'insert_table puts tables in the body, not in text boxes; pass a paragraph outside the box');
    const table = ctx.editor.insertTable(where, args.rows, {
      ...(args.header !== undefined ? { header: args.header } : {}),
      ...(args.widths ? { widths: args.widths } : {}),
      ...(args.borders !== undefined ? { borders: args.borders } : {}),
      ...(args.align ? { align: args.align } : {}),
    });
    const columns = Math.max(...args.rows.map((r) => r.length));
    const first = listParagraphs(ctx.doc).find((e) => e.para.el === findAll(table, 'w:p')[0]);
    return `Inserted a ${args.rows.length}×${columns} table ${said}${first?.cell ? ` (table ${first.cell.table})` : ''}: its cells are ${newNumbers(ctx, findAll(table, 'w:p'))}.`;
  },
});

const editTable = define({
  name: 'edit_table',
  title: 'Add or delete table rows',
  description:
    'Change a table’s rows: add rows (they copy the formatting of the row next to them), delete rows, or delete the whole table. ' +
    'Tables are numbered as read_document shows them (“table 2, row 1, cell 1”); to change a cell’s text, edit its paragraph. Paragraph numbers after the table shift.',
  role: 'editor',
  kind: 'change',
  schema: z.object({
    table: z.number().int().min(1).describe('Table number from read_document'),
    action: z.enum(['add_rows', 'delete_rows', 'delete_table']),
    rows: z.array(z.array(z.string())).min(1).optional().describe('add_rows: the new rows, each a list of cell texts'),
    beforeRow: z.number().int().min(1).optional().describe('add_rows: put the new rows before this row (default: after the last row)'),
    fromRow: z.number().int().min(1).optional().describe('delete_rows: first row to delete'),
    toRow: z.number().int().min(1).optional().describe('delete_rows: last row to delete (default: same as fromRow)'),
  }),
  run(ctx, args) {
    const table = tableAt(ctx, args.table);
    const rowsOf = () => table.children.filter((c) => c.type === 'element' && c.name === 'w:tr').length;
    if (args.action === 'delete_table') {
      ctx.editor.deleteTable(table);
      return `Deleted table ${args.table}.`;
    }
    if (args.action === 'add_rows') {
      if (!args.rows) throw new Error('Pass rows: the new rows’ cell texts.');
      const at = args.beforeRow === undefined ? rowsOf() : args.beforeRow - 1;
      const made = ctx.editor.insertTableRows(table, at, args.rows);
      return `Added ${plural(made.length, 'row')} to table ${args.table} (now ${plural(rowsOf(), 'row')}): the new cells are ${newNumbers(ctx, made.flatMap((r) => findAll(r, 'w:p')))}.`;
    }
    if (args.fromRow === undefined) throw new Error('Pass fromRow: the first row to delete.');
    const toRow = args.toRow ?? args.fromRow;
    if (toRow < args.fromRow) throw new Error('`toRow` must not be before `fromRow`.');
    ctx.editor.deleteTableRows(table, args.fromRow - 1, toRow - 1);
    const which = toRow === args.fromRow ? `row ${args.fromRow}` : `rows ${args.fromRow}–${toRow}`;
    const left = table.children.some((c) => c.type === 'element' && c.name === 'w:tr') && listParagraphs(ctx.doc).some((e) => findAll(table, 'w:p').includes(e.para.el));
    return left ? `Deleted ${which} of table ${args.table} (${plural(rowsOf(), 'row')} left).` : `Deleted ${which}, the whole table ${args.table}.`;
  },
});

const setList = define({
  name: 'set_list',
  title: 'Change list formatting',
  description:
    'Make paragraphs a bulleted or numbered list, or switch a list between bullets, numbers (1. or 1)), letters (a. A.) and roman numerals (i. I.). ' +
    'Also restarts numbering, changes the list level (indent), adds paragraphs to an existing list, or takes them out of a list. ' +
    'Every paragraph from `from` to `to` is affected. A changed list keeps its indentation. read_document shows the labels. ' +
    'To make one list of paragraphs that are not next to each other (other paragraphs between them), give the first a format, then add each of the others with sameListAs pointing at the first. ' +
    '(To copy a list as it is, copy_paragraphs keeps lists, formatting and tables.)',
  role: 'editor',
  kind: 'change',
  schema: z.object({
    from: paragraphNumber.describe('First paragraph'),
    to: paragraphNumber.optional().describe('Last paragraph (default: same as from)'),
    format: z
      .enum(['bullet', 'decimal', 'decimalParen', 'lowerLetter', 'upperLetter', 'lowerRoman', 'upperRoman', 'none'])
      .optional()
      .describe('bullet •, decimal 1., decimalParen 1), lowerLetter a., upperLetter A., lowerRoman i., upperRoman I., or none to remove the list. The paragraphs become one list in this format'),
    start: z.number().int().min(0).max(32767).optional().describe('First number. Without format, restarts the numbering in the current format'),
    level: z.number().int().min(1).max(9).optional().describe('List level: 1 is the outermost, 2 is indented once, and so on'),
    sameListAs: paragraphNumber.optional().describe('Add the paragraphs to the list this paragraph is in, continuing its numbering'),
  }),
  async run(ctx, args) {
    const to = args.to ?? args.from;
    if (to < args.from) throw new Error('`to` must not be before `from`.');
    const targets: Array<ParagraphEntry['para']['el']> = [];
    for (let n = args.from; n <= to; n++) targets.push(entry(ctx, n).para.el);
    if (args.sameListAs !== undefined && !entry(ctx, args.sameListAs).para.numbering) {
      throw new Error(
        `Paragraph ${args.sameListAs} is not in a list yet. Make it one first (set_list from ${args.sameListAs} with a format), then add the others with sameListAs: ${args.sameListAs}.`,
      );
    }
    await ctx.editor.setList(targets, {
      format: args.format,
      start: args.start,
      level: args.level === undefined ? undefined : args.level - 1,
      sameListAs: args.sameListAs === undefined ? undefined : entry(ctx, args.sameListAs).para.el,
    });
    const which = to === args.from ? `Paragraph ${args.from} is` : `Paragraphs ${args.from}–${to} are`;
    if (args.format === 'none') return `${which} no longer in a list.`;
    const items = ctx.editor.listItems();
    const labels = targets.map((el) => `"${items.get(el)?.label ?? ''}"`);
    return `${which} now labelled ${labels.join(', ')}.`;
  },
});

const formatText = define({
  name: 'format_text',
  title: 'Format text',
  description: 'Change character formatting (bold, italic, underline, colour, highlight, size, font) of a paragraph or of matching text.',
  role: 'editor',
  kind: 'change',
  schema: z.object({
    ...target,
    all: z.boolean().optional().describe('Format every match of `text` instead of one (matches in locked regions are skipped)'),
    format: textFormat,
  }),
  run(ctx, args) {
    const f = args.format;
    const format: TextFormat = {
      ...(f.bold !== undefined ? { bold: f.bold } : {}),
      ...(f.italic !== undefined ? { italic: f.italic } : {}),
      ...(f.underline !== undefined ? { underline: f.underline } : {}),
      ...(f.strike !== undefined ? { strike: f.strike } : {}),
      ...(f.color !== undefined ? { color: f.color.replace('#', '').toUpperCase() } : {}),
      ...(f.highlight !== undefined ? { highlight: f.highlight } : {}),
      ...(f.size !== undefined ? { size: f.size } : {}),
      ...(f.font !== undefined ? { font: f.font } : {}),
    };
    if (Object.keys(format).length === 0) throw new Error('`format` is empty; nothing to change.');
    if (args.all) {
      if (!args.text) throw new Error('`all` needs `text`.');
      const matches = matchesIn(ctx, args.text, args.paragraph);
      if (matches.length === 0) throw new Error(`Text not found: ${quote(args.text)}.`);
      let skipped = 0;
      for (const m of matches) {
        if (ctx.editor.isLocked(m)) skipped++;
        else ctx.editor.formatRange(m, format);
      }
      if (skipped === matches.length) throw new Error('Every match is inside a locked region; nothing changed.');
      return `Formatted ${plural(matches.length - skipped, 'match', 'matches')} of ${quote(args.text)}${skipped ? `; ${skipped} skipped (locked)` : ''}.`;
    }
    const { range, label } = resolveTarget(ctx, args);
    ctx.editor.formatRange(range, format);
    return `Formatted ${label}.`;
  },
});

const addHyperlink = define({
  name: 'add_hyperlink',
  title: 'Add a hyperlink',
  description: 'Turn text into a hyperlink to a web address (http, https, mailto, tel) or to a bookmark in the document. The text must be within one paragraph.',
  role: 'editor',
  kind: 'add',
  schema: z.object({
    ...target,
    url: z.string().optional().describe('Web address to link to'),
    bookmark: z.string().optional().describe('Bookmark name to link to (see list_links)'),
    tooltip: z.string().optional(),
  }),
  async run(ctx, args) {
    if ((args.url === undefined) === (args.bookmark === undefined)) throw new Error('Pass either url or bookmark.');
    if (args.url !== undefined) {
      let url: URL;
      try {
        url = new URL(args.url);
      } catch {
        throw new Error(`Not a valid URL: ${args.url}`);
      }
      if (!['http:', 'https:', 'mailto:', 'tel:'].includes(url.protocol)) throw new Error(`Links to ${url.protocol} addresses are not allowed (use http, https, mailto or tel).`);
    } else if (!ctx.editor.bookmarks().some((b) => b.name === args.bookmark)) {
      throw new Error(`No bookmark named "${args.bookmark}". Add it with add_bookmark first.`);
    }
    const { range, label } = resolveTarget(ctx, args);
    await ctx.editor.addHyperlink(range, args.url !== undefined ? { url: args.url } : { bookmark: args.bookmark! }, args.tooltip);
    return `Linked ${label} to ${args.url ?? `bookmark ${args.bookmark}`}.`;
  },
});

const addBookmark = define({
  name: 'add_bookmark',
  title: 'Add a bookmark',
  description: 'Add a bookmark (a named anchor that hyperlinks can point to) around a paragraph or a piece of text.',
  role: 'editor',
  kind: 'add',
  schema: z.object({
    ...target,
    name: z
      .string()
      .regex(/^[A-Za-z][A-Za-z0-9_]{0,39}$/)
      .describe('Bookmark name: starts with a letter; letters, digits and _ only; at most 40 characters'),
  }),
  run(ctx, args) {
    if (ctx.editor.bookmarks().some((b) => b.name.toLowerCase() === args.name.toLowerCase())) throw new Error(`There is already a bookmark named "${args.name}".`);
    const { range, label } = resolveTarget(ctx, args);
    ctx.editor.addBookmark(range, args.name);
    return `Added bookmark ${args.name} on ${label}.`;
  },
});

// ---- locks (owner) ----------------------------------------------------------------------------------

const lockRegion = define({
  name: 'lock_region',
  title: 'Lock a region',
  description:
    'Lock paragraphs or a piece of text against editing (a Word content control with "contents cannot be edited"). Word, docxkeun and these tools all respect it.',
  role: 'owner',
  kind: 'change',
  schema: z.object({
    ...target,
    toParagraph: paragraphNumber.optional().describe('With paragraph (and no text): lock paragraphs paragraph…toParagraph'),
    title: z.string().optional().describe('A name for the region, shown in Word'),
  }),
  run(ctx, args) {
    const title = args.title ? { title: args.title } : {};
    if (args.text === undefined && args.paragraph !== undefined) {
      const last = args.toParagraph ?? args.paragraph;
      if (last < args.paragraph) throw new Error('`toParagraph` must not be before `paragraph`.');
      const id = ctx.editor.lockParagraphs(entry(ctx, args.paragraph).para.el, entry(ctx, last).para.el, title);
      return `Locked ${last === args.paragraph ? `paragraph ${last}` : `paragraphs ${args.paragraph}–${last}`} as lock ${id}.`;
    }
    const { range, label } = resolveTarget(ctx, args);
    const id = ctx.editor.lockRange(range, title);
    return `Locked ${label} as lock ${id}.`;
  },
});

const unlockRegion = define({
  name: 'unlock_region',
  title: 'Unlock a region',
  description: 'Remove a lock (ids come from read_document or get_document_info).',
  role: 'owner',
  kind: 'change',
  schema: z.object({ lockId: z.string().min(1) }),
  run(ctx, args) {
    ctx.editor.unlock(args.lockId);
    return `Unlocked lock ${args.lockId}.`;
  },
});


// ---- tracked changes --------------------------------------------------------------------------------

const REVISION_WORDS: Record<Revision['kind'], string> = {
  insert: 'inserted',
  delete: 'deleted',
  'move-from': 'moved away',
  'move-to': 'moved here',
  format: 'formatting changed',
  'paragraph-format': 'paragraph formatting changed',
  'paragraph-insert': 'paragraph break inserted',
  'paragraph-delete': 'paragraph break deleted',
  'row-insert': 'table row inserted',
  'row-delete': 'table row deleted',
  property: 'properties changed',
};

function revisionParagraph(ctx: DocContext, r: Revision): number | undefined {
  return r.paragraph ? ctx.entries.find((e) => e.para.el === r.paragraph)?.n : undefined;
}

const listRevisionsOp = define({
  name: 'list_revisions',
  title: 'List tracked changes',
  description: 'List the tracked changes (insertions, deletions, moves, formatting changes) with their ids, authors, dates and paragraph numbers. read_document shows the text with every change accepted.',
  role: 'viewer',
  kind: 'read',
  schema: z.object({ author: z.string().optional().describe('Only changes by this author') }),
  run(ctx, args) {
    const revisions = ctx.editor.revisions().filter((r) => !args.author || r.author === args.author);
    const tracking = ctx.editor.trackChanges;
    const head = `Track Changes is ${tracking ? 'on: edits are recorded as tracked changes' : 'off'}.`;
    if (!revisions.length) return `${head}\nNo tracked changes${args.author ? ` by ${args.author}` : ''}.`;
    const lines = revisions.map((r) => {
      const n = revisionParagraph(ctx, r);
      const date = r.date ? `, ${r.date.slice(0, 16).replace('T', ' ')}` : '';
      return `Change ${r.id} (${REVISION_WORDS[r.kind]}) by ${r.author}${date}${n ? `, paragraph ${n}` : ''}${r.text ? `: ${quote(r.text, 120)}` : ''}`;
    });
    return [head, `${plural(revisions.length, 'tracked change')}:`, ...lines].join('\n');
  },
});

const revisionFilter = {
  ids: z.array(z.string()).optional().describe('Change ids from list_revisions'),
  author: z.string().optional().describe('All changes by this author'),
  all: z.boolean().optional().describe('Every tracked change in the document'),
};

function reviewFilter(args: { ids?: string[] | undefined; author?: string | undefined; all?: boolean | undefined }): { ids?: string[]; author?: string } | undefined {
  if (!args.ids?.length && !args.author && !args.all) throw new Error('Say which changes: pass `ids`, `author`, or `all: true`.');
  return args.all && !args.ids?.length && !args.author ? undefined : { ...(args.ids?.length ? { ids: args.ids } : {}), ...(args.author ? { author: args.author } : {}) };
}

const acceptRevisionsOp = define({
  name: 'accept_revisions',
  title: 'Accept tracked changes',
  description: 'Accept tracked changes: insertions stay, deletions go, formatting changes are kept. Pick them by id, by author, or all.',
  role: 'editor',
  kind: 'change',
  schema: z.object(revisionFilter),
  run(ctx, args) {
    const n = ctx.editor.acceptRevisions(reviewFilter(args));
    if (!n) throw new Error('No tracked changes matched.');
    return `Accepted ${plural(n, 'tracked change')}.`;
  },
});

const rejectRevisionsOp = define({
  name: 'reject_revisions',
  title: 'Reject tracked changes',
  description: 'Reject tracked changes: insertions are removed, deleted text comes back, formatting changes are undone. Pick them by id, by author, or all.',
  role: 'editor',
  kind: 'change',
  schema: z.object(revisionFilter),
  run(ctx, args) {
    const n = ctx.editor.rejectRevisions(reviewFilter(args));
    if (!n) throw new Error('No tracked changes matched.');
    return `Rejected ${plural(n, 'tracked change')}.`;
  },
});

const setTrackChanges = define({
  name: 'set_track_changes',
  title: 'Turn Track Changes on or off',
  description: ({ author }) =>
    `Turn Word's Track Changes setting on or off for a document. While it is on, every edit (yours too) is recorded as a tracked change signed "${author}" that people can accept or reject.`,
  role: 'editor',
  kind: 'change',
  schema: z.object({ on: z.boolean() }),
  async run(ctx, args) {
    await ctx.editor.setTrackChangesSetting(args.on);
    return `Track Changes is now ${args.on ? 'on' : 'off'}.`;
  },
});

// ---- paragraph formatting ---------------------------------------------------------------------------

const formatParagraph = define({
  name: 'format_paragraph',
  title: 'Format paragraphs',
  description: 'Change paragraph alignment, text direction (right-to-left for Arabic or Hebrew), indents, spacing, line spacing and keep/page-break rules.',
  role: 'editor',
  kind: 'change',
  schema: z.object({
    from: paragraphNumber.describe('First paragraph'),
    to: paragraphNumber.optional().describe('Last paragraph (default: same as from)'),
    align: z.enum(['left', 'center', 'right', 'justify']).optional(),
    direction: z.enum(['ltr', 'rtl']).optional().describe('Text direction; rtl for Arabic, Hebrew, Persian, Urdu'),
    indentLeft: z.number().min(0).max(1000).optional().describe('Left indent in points (72 = 1 inch)'),
    indentRight: z.number().min(0).max(1000).optional().describe('Right indent in points'),
    indentFirstLine: z.number().min(-1000).max(1000).optional().describe('First-line indent in points; negative for a hanging indent'),
    spaceBefore: z.number().min(0).max(1000).optional().describe('Space before in points'),
    spaceAfter: z.number().min(0).max(1000).optional().describe('Space after in points'),
    lineSpacing: z.number().min(0.5).max(5).optional().describe('Line spacing as a multiple: 1, 1.15, 1.5, 2'),
    keepWithNext: z.boolean().optional(),
    keepLinesTogether: z.boolean().optional(),
    pageBreakBefore: z.boolean().optional(),
  }),
  run(ctx, args) {
    const { from, to = args.from, ...format } = args;
    if (to < from) throw new Error('`to` must not be before `from`.');
    const set = Object.entries(format).filter(([, v]) => v !== undefined);
    if (!set.length) throw new Error('Say what to change (align, direction, indents, spacing…).');
    for (let n = from; n <= to; n++) ctx.editor.formatParagraph(entry(ctx, n).para.el, format as ParagraphFormat);
    return `Formatted ${from === to ? `paragraph ${from}` : `paragraphs ${from}–${to}`}: ${set.map(([k, v]) => `${k} ${v}`).join(', ')}.`;
  },
});

// ---- headers, footers, notes and text boxes ---------------------------------------------------------

interface StoryInfo {
  id: string;
  label: string;
  part: string;
  paragraphs: DocParagraph[];
  /** The text to show, when it differs from the paragraphs' (page fields as {PAGE}). */
  text?: string;
}

function stories(ctx: DocContext): StoryInfo[] {
  const out: StoryInfo[] = [];
  const counts = { header: 0, footer: 0 };
  const hfs = ctx.editor.headersAndFooters();
  const sectionCount = hfs.length ? Math.max(...hfs.flatMap((h) => h.sections)) : 0;
  for (const h of hfs) {
    const n = ++counts[h.kind];
    const which = h.type === 'first' ? 'first page' : h.type === 'even' ? 'even pages' : 'every page';
    const where = h.sections.length === 1 ? `section ${h.sections[0]}` : `sections ${h.sections.join(', ')}`;
    out.push({
      id: `${h.kind}${n}`,
      label: `${h.kind === 'header' ? 'Header' : 'Footer'} ${n} (${which}, ${sectionCount > 1 || h.sections[0] !== 1 ? where : 'the whole document'})`,
      part: h.part,
      paragraphs: new Editor(ctx.doc, undefined, { part: h.part }).paragraphs(),
      text: h.text,
    });
  }
  for (const sp of ctx.doc.storyParts()) {
    if (sp.kind === 'header' || sp.kind === 'footer') continue;
    const xml = ctx.doc.xmlSync(sp.part);
    const kind = sp.kind === 'footnotes' ? 'footnote' : 'endnote';
    for (const note of xml?.root.children ?? []) {
      if (note.type !== 'element' || note.name !== `w:${kind}`) continue;
      const type = note.attrs['w:type'];
      if (type && type !== 'normal') continue;
      const id = note.attrs['w:id'] ?? '';
      out.push({ id: `${kind}${id}`, label: `${kind === 'footnote' ? 'Footnote' : 'Endnote'} ${id}`, part: sp.part, paragraphs: paragraphsOf(ctx.doc.note(kind, id)) });
    }
  }
  // Text boxes, numbered as read_document numbers them ("text box 1").
  const boxes = new Map<number, StoryInfo>();
  for (const e of ctx.entries) {
    if (!e.box) continue;
    let box = boxes.get(e.box.index);
    if (!box) {
      box = { id: `textbox${e.box.index}`, label: `Text box ${e.box.index} (in paragraph ${e.box.host})`, part: ctx.doc.mainPart, paragraphs: [] };
      boxes.set(e.box.index, box);
      out.push(box);
    }
    box.paragraphs.push(e.para);
  }
  return out;
}

const readStories = define({
  name: 'read_headers_footers',
  title: 'Read headers, footers, notes and text boxes',
  description: 'Read the text outside the body: headers, footers, footnotes, endnotes and text boxes, each with an id (header1, footer1, footnote1, textbox1) for replace_story_text.',
  role: 'viewer',
  kind: 'read',
  schema: z.object({}),
  run(ctx) {
    const all = stories(ctx);
    if (!all.length) return 'This document has no headers, footers, footnotes, endnotes or text boxes. set_header_footer adds a header or footer.';
    const out: string[] = [];
    for (const s of all) {
      out.push(`${s.id}: ${s.label}`);
      const text = (s.text !== undefined ? s.text.split('\n') : s.paragraphs.map((p) => paragraphText(p).text)).filter((t) => t.trim());
      out.push(...(text.length ? text.map((t) => `  ${t.replace(/\t/g, ' → ')}`) : ['  (empty)']));
    }
    return out.join('\n');
  },
});

const replaceStoryText = define({
  name: 'replace_story_text',
  title: 'Edit a header, footer, note or text box',
  description: 'Replace text in a header, footer, footnote, endnote or text box (ids from read_headers_footers). Use an empty `replace` to delete the text.',
  role: 'editor',
  kind: 'change',
  schema: z.object({
    story: z.string().describe('Story id from read_headers_footers, e.g. "header1", "footnote2", "textbox1"'),
    find: z.string().min(1).describe('Exact text to replace (case-sensitive)'),
    replace: z.string().describe('New text'),
    all: z.boolean().optional().describe('Replace every match (default: only the first)'),
  }),
  run(ctx, args) {
    const story = stories(ctx).find((s) => s.id === args.story);
    if (!story) throw new Error(`No story "${args.story}"; read_headers_footers lists them.`);
    const editor = new Editor(ctx.doc, ctx.editor.guard, { part: story.part, author: ctx.editor.author });
    let count = 0;
    // Later matches first, so earlier positions stay valid.
    const matches: DocRange[] = [];
    for (const p of story.paragraphs) {
      const pt = paragraphText(p);
      for (let i = pt.text.indexOf(args.find); i >= 0; i = pt.text.indexOf(args.find, i + args.find.length)) {
        const range = rangeIn(pt, i, i + args.find.length);
        if (range) matches.push(range);
        if (!args.all) break;
      }
      if (matches.length && !args.all) break;
    }
    if (!matches.length) throw new Error(`Text not found in ${story.id}: ${quote(args.find)}.`);
    for (const range of matches.reverse()) {
      if (args.replace) editor.replaceRange(range, args.replace);
      else editor.deleteRange(range);
      count++;
    }
    return `Replaced ${plural(count, 'match', 'matches')} of ${quote(args.find)} in ${story.label}.`;
  },
});

/** Write tagged text into an empty paragraph. */
export function writeTagged(editor: DocxEditor, paragraph: DocParagraph['el'], segments: Segment[]): void {
  const text = segments.map((x) => x.text).join('');
  if (!text) return;
  editor.appendText(paragraph, text);
  tagParagraph(editor, paragraph, segments);
}

/** Format the parts of a paragraph's new text that `segments` tag (its text is their text, in order). */
export function tagParagraph(editor: DocxEditor, paragraph: DocParagraph['el'], segments: Segment[]): void {
  if (!segments.some((x) => x.tags.size)) return;
  const para = editor.paragraphs().find((p) => p.el === paragraph);
  if (!para) return;
  const pt = paragraphText(para);
  const parts: Array<{ range: DocRange; format: TextFormat }> = [];
  let at = 0;
  for (const seg of segments) {
    const format = seg.tags.size ? formatFor(seg.tags, new Set()) : undefined;
    const range = format ? rangeIn(pt, at, at + seg.text.length) : undefined;
    at += seg.text.length;
    if (range && format) parts.push({ range, format });
  }
  // Later parts first, so earlier positions stay valid as runs are split.
  for (const { range, format } of parts.reverse()) editor.formatRange(range, format);
}

/** A range for characters [a, b) of a paragraph's text. */
function rangeIn(pt: ReturnType<typeof paragraphText>, a: number, b: number): DocRange | undefined {
  const at = (k: number, side: 'start' | 'end') => {
    for (const s of pt.segments) {
      const inside = side === 'start' ? k >= s.start && k < s.end : k > s.start && k <= s.end;
      if (inside) return s.node.name === 'w:t' ? { node: s.node, offset: k - s.start } : { node: s.node, offset: side === 'start' ? 0 : 1 };
    }
    return undefined;
  };
  const start = at(a, 'start');
  const end = at(b, 'end');
  return start && end ? { start, end } : undefined;
}

const hfKind = z.enum(['header', 'footer']);

const setHeaderFooter = define({
  name: 'set_header_footer',
  title: 'Add or change a header or footer',
  description:
    'Add a header or footer, or replace one. Lines become paragraphs; a tab "\\t" splits a line into left, centre and right parts; ' +
    '{PAGE}, {NUMPAGES} and {SECTIONPAGES} become page numbers ("Page {PAGE} of {NUMPAGES}"). ' +
    'By default it is for every page of every section; `type` "first" is for the first page only (e.g. a cover), "even" for even pages.',
  role: 'editor',
  kind: 'change',
  schema: z.object({
    kind: hfKind,
    text: z.string().describe('The text; "\\n" between paragraphs, "\\t" between left / centre / right parts, {PAGE} {NUMPAGES} {SECTIONPAGES} for page numbers'),
    type: z.enum(['default', 'first', 'even']).optional().describe('Which pages: every page (default), the first page, or even pages'),
    section: z.number().int().min(1).optional().describe('Only this section (from 1); default every section'),
    align: z.enum(['left', 'center', 'right']).optional(),
  }),
  async run(ctx, args) {
    await ctx.editor.setHeaderFooter(args.kind, args.text, { ...(args.type ? { type: args.type } : {}), ...(args.section ? { section: args.section } : {}), ...(args.align ? { align: args.align } : {}) });
    const pages = args.type === 'first' ? 'first-page ' : args.type === 'even' ? 'even-page ' : '';
    const where = args.section ? `section ${args.section}` : 'the document';
    return `Set the ${pages}${args.kind} of ${where} to ${quote(args.text.replace(/\t/g, ' → ').replace(/\n/g, ' / '), 120)}.`;
  },
});

const removeHeaderFooter = define({
  name: 'remove_header_footer',
  title: 'Remove a header or footer',
  description: 'Remove headers or footers: all of them, or only the first-page / even-page / default one, in every section or one section.',
  role: 'editor',
  kind: 'change',
  schema: z.object({
    kind: hfKind,
    type: z.enum(['default', 'first', 'even', 'all']).optional().describe('Which ones (default all)'),
    section: z.number().int().min(1).optional().describe('Only this section (from 1); default every section'),
  }),
  async run(ctx, args) {
    const n = await ctx.editor.removeHeaderFooter(args.kind, { ...(args.type ? { type: args.type } : {}), ...(args.section ? { section: args.section } : {}) });
    if (!n) throw new Error(`There is no ${args.type && args.type !== 'all' ? `${args.type} ` : ''}${args.kind} to remove${args.section ? ` in section ${args.section}` : ''}.`);
    return `Removed ${plural(n, args.kind)}.`;
  },
});

// ---- pictures, charts, equations --------------------------------------------------------------------

function describeObject(ctx: DocContext, image: ImageInline): string {
  const c = image.content;
  const alt = image.alt ? ` "${image.alt}"` : '';
  const size = `${Math.round(image.width)}×${Math.round(image.height)}pt, ${image.float ? 'floating' : 'inline'}`;
  if (!c || c.type === 'picture') return `picture${alt} (${size})`;
  if (c.type === 'chart') {
    const xml = ctx.doc.xmlSync(c.part)?.root;
    let text = 'chart';
    if (xml) {
      try {
        text = describeChart(readChart(xml));
      } catch {
        // keep the generic name
      }
    }
    return `${text.charAt(0).toLowerCase()}${text.slice(1)} (${size})`;
  }
  if (c.type === 'diagram') {
    const data = ctx.doc.xmlSync(c.dataPart)?.root;
    const texts: string[] = [];
    const walk = (n: { type: string; name?: string; children?: unknown[]; value?: string }): void => {
      if (n.type !== 'element') return;
      if (n.name === 'a:t') texts.push((n.children as Array<{ value?: string }> | undefined)?.map((c) => c.value ?? '').join('') ?? '');
      for (const c of n.children ?? []) walk(c as typeof n);
    };
    if (data) walk(data as never);
    return `SmartArt diagram${texts.length ? `: ${texts.join(' / ')}` : ''} (${size})`;
  }
  if (c.type === 'shape') {
    const prst = c.shape.geometry?.attrs['prst'];
    const text = c.shape.text ? blocksToText(c.shape.text.blocks).replace(/\s+/g, ' ').trim() : '';
    return `${c.shape.text ? 'text box' : `shape${prst ? ` (${prst})` : ''}`}${text ? ` ${quote(text, 80)}` : ''} (${size})`;
  }
  return `group of ${c.children.length} shapes (${size})`;
}

/** An object as the tools show it: its number, the paragraph holding it, and the object itself. */
export interface NumberedObject {
  /** Object number, as DocxEditor.objects() numbers it. */
  n: number;
  /** read_document's number of the paragraph holding it. */
  paragraph: number;
  object: DocObject;
  image: ImageInline;
}

/** The objects in the body, numbered as DocxEditor.objects() numbers them, with the paragraphs holding them. */
export function numberedObjects(editor: DocxEditor, entries: ReadonlyArray<ParagraphEntry>): NumberedObject[] {
  const byParagraph = new Map<Element, DocObject[]>();
  for (const o of editor.objects()) {
    const list = byParagraph.get(o.paragraph.el);
    if (list) list.push(o);
    else byParagraph.set(o.paragraph.el, [o]);
  }
  const out: NumberedObject[] = [];
  for (const e of entries) {
    const objects = byParagraph.get(e.para.el);
    if (!objects) continue;
    // DocxEditor.objects() goes along each paragraph's pictures and drawings in order too.
    const images = e.para.inlines.filter((i): i is ImageInline => i.kind === 'image');
    images.forEach((image, k) => {
      const object = objects[k];
      if (object) out.push({ n: object.n, paragraph: e.n, object, image });
    });
  }
  return out;
}

/** "object 2: clustered column chart "Revenue": … (432×252pt, inline)" */
function objectLabel(ctx: DocContext, o: NumberedObject): string {
  return `object ${o.n}: ${describeObject(ctx, o.image)}`;
}

const listObjects = define({
  name: 'list_objects',
  title: 'List pictures, charts, shapes and equations',
  description:
    'List the pictures, charts (with their data), SmartArt, shapes, text boxes and equations in a document. Each line starts with the number of the paragraph holding it; ' +
    'pictures, charts and shapes have an object number ("object 2") and say whether they sit in the text (inline) or float.',
  role: 'viewer',
  kind: 'read',
  schema: z.object({}),
  run(ctx) {
    const numbered = new Map(numberedObjects(ctx.editor, ctx.entries).map((o) => [o.image, o]));
    const out: string[] = [];
    for (const e of ctx.entries) {
      for (const i of e.para.inlines) {
        const o = i.kind === 'image' ? numbered.get(i) : undefined;
        if (o) out.push(`[${e.n}] ${objectLabel(ctx, o)}`);
        else if (i.kind === 'math') out.push(`[${e.n}] ${i.display ? 'display ' : ''}equation: ${i.text}`);
      }
    }
    return out.length ? out.join('\n') : 'No pictures, charts, shapes or equations.';
  },
});

const insertImageOp = define({
  name: 'insert_image',
  title: 'Insert a picture',
  description: 'Insert a picture file (PNG, JPEG, GIF, BMP, TIFF, EMF or WMF) into a paragraph, at its end or after some text. Without a size it keeps its proportions and fits the page width.',
  role: 'editor',
  kind: 'add',
  schema: z.object({
    image: z.string().min(1).describe('Path of the picture file (inside the allowed folders)'),
    paragraph: paragraphNumber.describe('Paragraph to put it in'),
    after: z.string().min(1).optional().describe('Put it right after this text in the paragraph (default: at the end)'),
    width: z.number().min(1).max(2000).optional().describe('Width in points (72 = 1 inch); height follows the proportions'),
    height: z.number().min(1).max(2000).optional().describe('Height in points'),
    alt: z.string().optional().describe('Alternative text describing the picture, for accessibility'),
  }),
  async run(ctx, args) {
    if (!ctx.loadImage) throw new Error('Pictures can only be inserted from files where this server can read them.');
    const bytes = await ctx.loadImage(args.image);
    const at = insertionPoint(ctx, args.paragraph, args.after);
    await ctx.editor.insertImage(at, bytes, { ...(args.width ? { width: args.width } : {}), ...(args.height ? { height: args.height } : {}), ...(args.alt ? { alt: args.alt } : {}) });
    return `Inserted ${quote(args.image)} in paragraph ${args.paragraph}${args.after ? ` after ${quote(args.after)}` : ''}.`;
  },
});

/** Where something goes: after `after` in the paragraph, or at its end. */
function insertionPoint(ctx: DocContext, paragraph: number, after: string | undefined) {
  if (after) return resolveTarget(ctx, { paragraph, text: after }).range.end;
  const e = entry(ctx, paragraph);
  return ctx.editor.paragraphRange(e.para)?.end ?? { node: e.para.el, offset: 0 };
}

const addFootnoteOp = define({
  name: 'add_footnote',
  title: 'Add a footnote',
  description: 'Add a footnote: its number goes in a paragraph (after some text, or at the end), its text at the bottom of the page.',
  role: 'editor',
  kind: 'add',
  schema: z.object({
    paragraph: paragraphNumber.describe('Paragraph where the footnote number goes'),
    after: z.string().min(1).optional().describe('Put the number right after this text (default: end of the paragraph)'),
    text: z.string().min(1).describe('The footnote text'),
  }),
  async run(ctx, args) {
    const id = await ctx.editor.addFootnote(insertionPoint(ctx, args.paragraph, args.after), args.text);
    return `Added footnote ${id} in paragraph ${args.paragraph}.`;
  },
});

const chartType = z.enum(['column', 'bar', 'line', 'area', 'pie', 'doughnut', 'scatter']);
const legendPosition = z.enum(['bottom', 'top', 'left', 'right', 'none']);
const chartCategories = z.array(z.union([z.string(), z.number()])).min(1);
const chartSeries = z.array(z.object({ name: z.string(), values: z.array(z.number().nullable()) })).min(1);

const OBJECT_NAMES: Record<DocObject['kind'], string> = {
  picture: 'a picture',
  chart: 'a chart',
  diagram: 'a SmartArt diagram',
  shape: 'a shape',
  textBox: 'a text box',
  group: 'a group of shapes',
};

/** The chart called "object n", with the paragraph holding it. */
function chartObject(ctx: DocContext, n: number): NumberedObject {
  const all = numberedObjects(ctx.editor, ctx.entries);
  const o = all.find((x) => x.n === n);
  if (!o) throw new Error(`There is no object ${n}; the document has ${plural(all.length, 'object')}.`);
  const c = o.image.content;
  const hasText = c?.type === 'shape' && !!c.shape.text;
  if (o.object.kind !== 'chart') throw new Error(`Object ${n} is ${hasText ? 'a text box' : OBJECT_NAMES[o.object.kind]}, not a chart.`);
  return o;
}

const insertChartOp = define({
  name: 'insert_chart',
  title: 'Insert a chart',
  description:
    'Insert a chart into a paragraph, at its end or after some text. Give the categories (the labels along the bottom, or the slices of a pie) and one series per line or set of bars, each with a value for every category (null leaves a gap). ' +
    'The data sits in an embedded workbook, so "Edit Data" in Word shows it. For a scatter chart the categories are the x values (numbers). Without a size it is 6 × 3.5 inches, shrunk to the page width.',
  role: 'editor',
  kind: 'add',
  schema: z.object({
    paragraph: paragraphNumber.describe('Paragraph to put it in'),
    after: z.string().min(1).optional().describe('Put it right after this text in the paragraph (default: at the end)'),
    type: chartType.describe('column (vertical bars), bar (horizontal), line, area, pie, doughnut or scatter'),
    stacked: z.enum(['stacked', 'percent']).optional().describe('Stack the bars or areas (column, bar and area charts only)'),
    title: z.string().optional().describe('Chart title'),
    categories: chartCategories.describe('Category labels, one per value of each series'),
    series: chartSeries.describe('The data: each series has a name and one value per category (null for none)'),
    legend: legendPosition.optional().describe('Where the legend goes (default bottom; none hides it)'),
    categoryTitle: z.string().optional().describe('Title of the category axis'),
    valueTitle: z.string().optional().describe('Title of the value axis'),
    width: z.number().min(1).max(2000).optional().describe('Width in points (72 = 1 inch); height follows the proportions'),
    height: z.number().min(1).max(2000).optional().describe('Height in points'),
    alt: z.string().optional().describe('Alternative text describing the chart, for accessibility'),
  }),
  async run(ctx, args) {
    const at = insertionPoint(ctx, args.paragraph, args.after);
    const axisTitles = { ...(args.categoryTitle ? { category: args.categoryTitle } : {}), ...(args.valueTitle ? { value: args.valueTitle } : {}) };
    const drawing = await ctx.editor.insertChart(at, {
      type: args.type,
      ...(args.stacked ? { stacked: args.stacked } : {}),
      ...(args.title ? { title: args.title } : {}),
      categories: args.categories,
      series: args.series,
      ...(args.legend ? { legend: args.legend } : {}),
      ...(Object.keys(axisTitles).length ? { axisTitles } : {}),
      ...(args.width ? { width: args.width } : {}),
      ...(args.height ? { height: args.height } : {}),
      ...(args.alt ? { alt: args.alt } : {}),
    });
    const made = ctx.editor.objects().find((o) => o.drawing === drawing);
    const info = made ? ctx.editor.chartInfo(made) : undefined;
    const kind = info?.kind ?? args.type;
    return `Inserted a ${kind} chart (object ${made?.n ?? '?'}) in paragraph ${args.paragraph}: ${plural(args.series.length, 'series', 'series')} × ${plural(args.categories.length, 'category', 'categories')}.`;
  },
});

const readChartOp = define({
  name: 'read_chart',
  title: 'Read a chart',
  description: 'Show a chart\'s type, titles, legend and data as a table (one line per series). Use the object number from read_document or list_objects.',
  role: 'viewer',
  kind: 'read',
  schema: z.object({ object: z.number().int().min(1).describe('Object number of the chart') }),
  run(ctx, args) {
    const o = chartObject(ctx, args.object);
    const info = ctx.editor.chartInfo(o.object);
    const parts = [info.kind];
    if (info.title) parts.push(`title ${quote(info.title)}`);
    if (info.axisTitles.category) parts.push(`category axis ${quote(info.axisTitles.category)}`);
    if (info.axisTitles.value) parts.push(`value axis ${quote(info.axisTitles.value)}`);
    parts.push(info.legend === 'none' ? 'no legend' : `legend ${info.legend}`);
    const source = info.source === 'embedded' ? 'Data in the embedded workbook.' : info.source === 'linked' ? 'Data linked to an outside file.' : 'No workbook behind the data.';
    const table = [`| ${info.categories.join(' | ')} |`, ...info.series.map((s) => [s.name, ...s.values.map((v) => v ?? '')].join(' | '))];
    return [`Chart (object ${o.n}, paragraph ${o.paragraph}): ${parts.join(', ')}. ${source}`, ...table].join('\n');
  },
});

const updateChartOp = define({
  name: 'update_chart',
  title: 'Change a chart',
  description:
    'Change a chart\'s title, axis titles, legend or data, keeping its look. Giving categories or series replaces them all, so give the whole table (read_chart shows the current one). An empty title removes it. ' +
    'Word does not track chart changes, so with Track Changes on this change is not tracked.',
  role: 'editor',
  kind: 'change',
  schema: z.object({
    object: z.number().int().min(1).describe('Object number of the chart'),
    title: z.string().optional().describe('New title (empty removes it)'),
    categories: chartCategories.optional().describe('All category labels'),
    series: chartSeries.optional().describe('All series, each with a name and one value per category'),
    legend: legendPosition.optional().describe('Where the legend goes (none hides it)'),
    categoryTitle: z.string().optional().describe('Title of the category axis (empty removes it)'),
    valueTitle: z.string().optional().describe('Title of the value axis (empty removes it)'),
  }),
  async run(ctx, args) {
    const o = chartObject(ctx, args.object);
    const axisTitles = { ...(args.categoryTitle !== undefined ? { category: args.categoryTitle } : {}), ...(args.valueTitle !== undefined ? { value: args.valueTitle } : {}) };
    const changed = [
      ...(args.title !== undefined ? ['title'] : []),
      ...(args.categoryTitle !== undefined || args.valueTitle !== undefined ? ['axis titles'] : []),
      ...(args.legend ? ['legend'] : []),
      ...(args.categories || args.series ? ['data'] : []),
    ];
    if (!changed.length) throw new Error('Nothing to change: give a title, axis titles, legend, categories or series.');
    const result = await ctx.editor.updateChart(o.object, {
      ...(args.title !== undefined ? { title: args.title } : {}),
      ...(Object.keys(axisTitles).length ? { axisTitles } : {}),
      ...(args.legend ? { legend: args.legend } : {}),
      ...(args.categories ? { categories: args.categories } : {}),
      ...(args.series ? { series: args.series } : {}),
    });
    const notes = [`Changed the ${changed.join(', ')} of the chart (object ${o.n}, paragraph ${o.paragraph}).`];
    if (result.source === 'linked' && (args.categories || args.series))
      notes.push("This chart's data is linked to an outside file; only the chart's own copy changed, so refreshing it in Word brings the old numbers back.");
    if (result.workbookHadMore) notes.push("The old workbook held more than the chart's data; it was replaced with the chart's data only.");
    if (ctx.editor.trackChanges) notes.push("Track Changes can't record chart data changes, so this change is not tracked.");
    return notes.join(' ');
  },
});

// ---- where an operation works (for live views) ------------------------------------------------------

/**
 * The paragraphs (read_document's numbers) a call will read or change, from its arguments:
 * `paragraph`, `from`/`to`, `after`/`before`, the paragraph holding `object` (given the document's
 * objects), or the paragraphs holding `text`, `find`, `query` or a template placeholder. Empty when
 * it is the whole document or cannot be told. Partial arguments (still being written by the model)
 * are fine. A live viewer numbers paragraphs differently: pass the result through viewerParagraphs.
 */
export function targetParagraphs(
  entries: ReadonlyArray<{ n: number; text: string }>,
  name: string,
  args: Record<string, unknown>,
  objects?: ReadonlyArray<{ n: number; paragraph: number }>,
): number[] {
  const out = new Set<number>();
  const num = (v: unknown) => (typeof v === 'number' && Number.isInteger(v) && v >= 1 ? v : undefined);
  const p = num(args.paragraph);
  const object = num(args.object);
  const holder = object === undefined ? undefined : objects?.find((o) => o.n === object)?.paragraph;
  if (holder !== undefined) out.add(holder);
  const from = num(args.from);
  const to = num(args.to);
  const after = num(args.after);
  const before = num(args.before);
  if (from !== undefined) for (let n = from; n <= Math.min(to ?? from, from + 500); n++) out.add(n);
  if (after !== undefined) out.add(after);
  if (before !== undefined) out.add(before);
  const needle = [args.text, args.find, args.query].find((v): v is string => typeof v === 'string' && v.length > 1);
  // For these, `text` is what they write, not where.
  if (p !== undefined && (!needle || name === 'edit_paragraph' || name === 'insert_image' || name === 'add_footnote')) out.add(p);
  else if (needle) {
    let re: RegExp | undefined;
    if (args.regex === true) {
      try {
        re = new RegExp(needle, args.matchCase === false ? 'i' : '');
      } catch {
        re = undefined;
      }
    }
    const lower = needle.toLowerCase();
    for (const e of entries) {
      if (p !== undefined && e.n !== p) continue;
      const hit = re ? re.test(e.text) : args.matchCase === false ? e.text.toLowerCase().includes(lower) : e.text.includes(needle);
      if (hit) out.add(e.n);
    }
  } else if (name === 'fill_template') for (const e of entries) if (/\{\{[^}]+\}\}/.test(e.text)) out.add(e.n);
  return [...out].sort((a, b) => a - b).slice(0, 500);
}

/** A verb for what an operation does, for a live view's caret ("Claude is commenting…"). */
export function activityLabel(name: string): string {
  if (name === 'find_text') return 'searching';
  if (/comment/.test(name)) return 'commenting';
  if (/^(read|list|get)_/.test(name)) return 'reading';
  if (/^format_|set_list/.test(name)) return 'formatting';
  if (/^copy_/.test(name)) return 'copying';
  if (/^delete_|reject_/.test(name)) return 'deleting';
  if (/^(insert|add|create)_/.test(name)) return 'writing';
  if (/revisions/.test(name)) return 'reviewing';
  return 'editing';
}

/** Every shared operation, reading ones first. */
export const OPERATIONS: readonly Operation[] = [
  readDocument,
  findText,
  listComments,
  listLinks,
  listRevisionsOp,
  readStories,
  listObjects,
  readChartOp,
  addComment,
  replyToComment,
  resolveComment,
  deleteComment,
  replaceText,
  fillTemplate,
  editParagraph,
  insertParagraphs,
  deleteParagraphs,
  copyParagraphs,
  insertTableOp,
  editTable,
  setList,
  formatText,
  addHyperlink,
  addBookmark,
  lockRegion,
  unlockRegion,
  formatParagraph,
  setHeaderFooter,
  removeHeaderFooter,
  replaceStoryText,
  insertImageOp,
  insertChartOp,
  updateChartOp,
  addFootnoteOp,
  setTrackChanges,
  acceptRevisionsOp,
  rejectRevisionsOp,
];

const ROLE_ORDER: Role[] = ['viewer', 'commenter', 'editor', 'owner'];

/** The operations a role may use. */
export function operationsFor(role: Role): Operation[] {
  const rank = ROLE_ORDER.indexOf(role);
  return OPERATIONS.filter((op) => ROLE_ORDER.indexOf(op.role) <= rank);
}

/** Context for one call: paragraphs are numbered afresh, since earlier calls may have changed them. */
export function contextFor(doc: DocxDocument, editor: DocxEditor, name: string, author: string): DocContext {
  return { doc, editor, name, author, entries: listParagraphs(doc) };
}
