/**
 * The document operations behind the tools: reading, finding, editing, lists, comments, links and
 * locks, on one open document. The MCP server and the docxkeun chat web app share them, so this
 * module uses no Node APIs. Each operation has a zod schema for its input (without file arguments),
 * the minimum role that may use it, and whether it reads, adds or changes.
 */
import { z } from 'zod';
import { blocksToText, describeChart, readChart, type DocParagraph, type DocxDocument, type ImageInline } from 'docxkeun';
import { DocxEditor as Editor, paragraphText, paragraphsOf, type DocRange, type DocxEditor, type Match, type ParagraphFormat, type Revision, type TextFormat } from 'docxkeun/editor';
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
  text: z.string().describe('Paragraph text; "\\n" makes a line break and "\\t" a tab'),
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
    'Read a Word document as numbered paragraphs. Each line is "[number] (style; list label; table position; locks; comments) text", ' +
    'e.g. "[4] (list item a.) …" or "[7] (bullet •, level 2) …". ' +
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
      const objects = e.para.inlines.flatMap((i) => (i.kind === 'image' ? [describeObject(ctx, i)] : i.kind === 'math' ? [`equation: ${i.text}`] : []));
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
    replace: z.string().describe('Replacement text (empty deletes the matches)'),
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
      ctx.editor.replaceRange(m, args.regex ? expandReplacement(args.replace, m) : args.replace);
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
  description: "Replace a paragraph's text (keeping the formatting of its first run) and/or change its style.",
  role: 'editor',
  kind: 'change',
  schema: z.object({
    paragraph: paragraphNumber.describe('Paragraph number from read_document'),
    text: z.string().optional().describe('New text for the whole paragraph'),
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
      const range = ctx.editor.paragraphRange(e.para);
      if (range) ctx.editor.replaceRange(range, args.text);
      else if (args.text) ctx.editor.appendText(e.para.el, args.text);
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
      els.push(ctx.editor.insertParagraph(where, p.text, opts));
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
    const bodyParagraphs = ctx.entries.filter((e) => !e.cell).length;
    if (doomed.filter((e) => !e.cell).length >= bodyParagraphs) {
      throw new Error('A document needs at least one paragraph outside tables; clear the last one with edit_paragraph instead.');
    }
    for (const e of doomed.reverse()) ctx.editor.deleteParagraph(e.para.el);
    return `Deleted ${to === args.from ? `paragraph ${args.from}` : `paragraphs ${args.from}–${to}`}.`;
  },
});

const setList = define({
  name: 'set_list',
  title: 'Change list formatting',
  description:
    'Make paragraphs a bulleted or numbered list, or switch a list between bullets, numbers (1. or 1)), letters (a. A.) and roman numerals (i. I.). ' +
    'Also restarts numbering, changes the list level (indent), adds paragraphs to an existing list, or takes them out of a list. ' +
    'Every paragraph from `from` to `to` is affected. A changed list keeps its indentation. read_document shows the labels.',
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
  /** A text box's container (edited in the main document). */
  container?: ImageInline;
}

/** Text boxes in the body, with the paragraph they are anchored in. */
function textBoxes(ctx: DocContext): Array<{ image: ImageInline; paragraph: number }> {
  const out: Array<{ image: ImageInline; paragraph: number }> = [];
  const visit = (image: ImageInline, paragraph: number) => {
    const c = image.content;
    if (c?.type === 'shape' && c.shape.text) out.push({ image, paragraph });
    if (c?.type === 'group') {
      const walk = (children: typeof c.children) => {
        for (const ch of children) {
          if (ch.content.type === 'shape' && ch.content.shape.text) out.push({ image: { ...image, content: ch.content }, paragraph });
          if (ch.content.type === 'group') walk(ch.content.children);
        }
      };
      walk(c.children);
    }
  };
  for (const e of ctx.entries) for (const i of e.para.inlines) if (i.kind === 'image') visit(i, e.n);
  return out;
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
  textBoxes(ctx).forEach(({ image, paragraph }, i) => {
    const c = image.content as { type: 'shape'; shape: { text: { blocks: Parameters<typeof paragraphsOf>[0]; container: unknown } } };
    out.push({ id: `textbox${i + 1}`, label: `Text box ${i + 1} (in paragraph ${paragraph})`, part: ctx.doc.mainPart, paragraphs: paragraphsOf(c.shape.text.blocks), container: image });
  });
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
    if (story.container) dropTextBoxFallback(ctx.doc);
    return `Replaced ${plural(count, 'match', 'matches')} of ${quote(args.find)} in ${story.label}.`;
  },
});

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

/** Text boxes are stored twice (DrawingML and an old VML copy); after an edit the copy is out of date. */
function dropTextBoxFallback(doc: DocxDocument): void {
  const walk = (node: { type: string; name?: string; children?: unknown[] }): void => {
    if (node.type !== 'element') return;
    const children = (node.children ?? []) as Array<{ type: string; name?: string; children?: unknown[] }>;
    if (node.name === 'mc:AlternateContent' && children.some((c) => c.type === 'element' && c.name === 'mc:Choice' && JSON.stringify(c).includes('"w:txbxContent"'))) {
      node.children = children.filter((c) => !(c.type === 'element' && c.name === 'mc:Fallback'));
      return;
    }
    for (const c of children) walk(c);
  };
  walk(doc.documentXml.root as never);
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
  const size = `${Math.round(image.width)}×${Math.round(image.height)}pt${image.float ? ', floating' : ''}`;
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
    return `${text} (${size})`;
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
    return `${c.shape.text ? 'text box' : `shape${prst ? ` (${prst})` : ''}`}${text ? `: ${quote(text, 80)}` : ''} (${size})`;
  }
  return `group of ${c.children.length} shapes (${size})`;
}

const listObjects = define({
  name: 'list_objects',
  title: 'List pictures, charts, shapes and equations',
  description: 'List the pictures, charts (with their data), SmartArt, shapes, text boxes and equations in a document, with paragraph numbers.',
  role: 'viewer',
  kind: 'read',
  schema: z.object({}),
  run(ctx) {
    const out: string[] = [];
    for (const e of ctx.entries) {
      for (const i of e.para.inlines) {
        if (i.kind === 'image') out.push(`[${e.n}] ${describeObject(ctx, i)}`);
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

// ---- where an operation works (for live views) ------------------------------------------------------

/**
 * The paragraphs (numbers from 1) a call will read or change, from its arguments: `paragraph`,
 * `from`/`to`, `after`/`before`, or the paragraphs holding `text`, `find`, `query` or a template
 * placeholder. Empty when it is the whole document or cannot be told. Partial arguments (still
 * being written by the model) are fine.
 */
export function targetParagraphs(entries: ReadonlyArray<{ n: number; text: string }>, name: string, args: Record<string, unknown>): number[] {
  const out = new Set<number>();
  const num = (v: unknown) => (typeof v === 'number' && Number.isInteger(v) && v >= 1 ? v : undefined);
  const p = num(args.paragraph);
  const from = num(args.from);
  const to = num(args.to);
  const after = num(args.after);
  const before = num(args.before);
  if (from !== undefined) for (let n = from; n <= Math.min(to ?? from, from + 500); n++) out.add(n);
  if (after !== undefined) out.add(after);
  if (before !== undefined) out.add(before);
  const needle = [args.text, args.find, args.query].find((v): v is string => typeof v === 'string' && v.length > 1);
  if (p !== undefined && (!needle || name === 'insert_image' || name === 'add_footnote')) out.add(p);
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
  addComment,
  replyToComment,
  resolveComment,
  deleteComment,
  replaceText,
  fillTemplate,
  editParagraph,
  insertParagraphs,
  deleteParagraphs,
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
