/**
 * The docxkeun MCP server: tools that let an AI client read, edit, comment on, protect and export
 * Word documents. Every change goes through docxkeun's AccessPolicy (the server's role) and its
 * locked regions, and only files inside the allowed folders can be read or written.
 */
import { readFileSync } from 'node:fs';
import { McpServer, type CallToolResult, type ToolAnnotations } from '@modelcontextprotocol/server';
import { z } from 'zod';
import { DocxDocument, layoutDocument } from 'docxkeun';
import { describeSelection, DocxEditor, LockedRegionError, paragraphsOf, selectionInfoFor } from 'docxkeun/editor';
import { exportPdf } from 'docxkeun/pdf';
import {
  AccessPolicy,
  PasswordRequiredError,
  PermissionError,
  WrongPasswordError,
  pdfSecurity,
  readProtection,
  removeProtection,
  setProtection,
  type Role,
} from 'docxkeun/security';
import { fontsFor } from './fonts.js';
import { activityLabel, describeOperation, newParagraph, numberedObjects, operationsFor, plural, tagParagraph, targetParagraphs, writeTagged, type DocContext, type McpRole } from './operations.js';
import { MAX_INSTRUCTIONS, instructionsSection } from './instructions.js';
import { parseMarkup } from './markup.js';
import { commentThreads, listParagraphs, locksByParagraph, paragraphStyles, quote, readNumbersByView, readNumbersWithChanges, readParagraphs, renumberSelection, resolveStyle, tableCount, viewerParagraphs, wordCount } from './outline.js';
import type { Preview } from './preview.js';
import { Queue, Workspace, WorkspaceError } from './workspace.js';

/** This package's version, read from its package.json (one level above src/ and dist/). */
export const VERSION: string = JSON.parse(readFileSync(new URL('../package.json', import.meta.url), 'utf8')).version;

/** Ranked: each role may use what the roles before it may, except that suggester sees only its own list. */
export const ROLES: McpRole[] = ['viewer', 'commenter', 'suggester', 'editor', 'owner'];

export interface DocxMcpOptions {
  /** Folders the server may read and write. */
  folders: string[];
  /** What the AI may do (default 'editor'). Document protection and locked regions apply on top. */
  role?: McpRole;
  /** Author name on comments the AI adds (default "AI assistant"). */
  author?: string;
  /** Extra font folders (.ttf/.otf) for page counts and PDF export, e.g. Microsoft Office's fonts. */
  fontFolders?: string[];
  /** A running live preview (see startPreview) to tell about every change. */
  preview?: Preview;
  /** The user's own instructions, as text (not a path), added after the built-in ones. At most MAX_INSTRUCTIONS characters once trimmed. */
  instructions?: string;
}

const ROLE_SUMMARY: Record<McpRole, string> = {
  viewer: 'read documents and export PDFs',
  commenter: 'read documents, export PDFs, and add, reply to and resolve comments',
  suggester: 'read documents, add, reply to and resolve comments, and change text, lists and tables as tracked changes that a person accepts or rejects',
  editor: 'read, edit, comment on and export documents',
  owner: 'do everything an editor can, and also lock regions, change document protection and set passwords',
};

const READ: ToolAnnotations = { readOnlyHint: true, openWorldHint: false };
const ADD: ToolAnnotations = { readOnlyHint: false, destructiveHint: false, idempotentHint: false, openWorldHint: false };
const CHANGE: ToolAnnotations = { readOnlyHint: false, destructiveHint: true, idempotentHint: false, openWorldHint: false };

// ---- file fields (the document operations' own fields come from operations.ts) ----------------------

const path = z.string().min(1).describe('The .docx file: an absolute path, or a path relative to the first allowed folder');
const password = z.string().optional().describe('Password to open the file, only for encrypted (password-protected) documents');
const output = z.string().optional().describe('Save the result to this .docx instead of changing the original');
const overwrite = z.boolean().optional().describe('Allow `output` to replace an existing file (default false)');
const saveFields = { output, overwrite };

function reply(text: string): CallToolResult {
  return { content: [{ type: 'text', text }] };
}

function kb(bytes: number): string {
  return bytes < 1024 ? `${bytes} bytes` : `${(bytes / 1024).toFixed(bytes < 10240 ? 1 : 0)} KB`;
}

interface Session extends DocContext {
  policy: AccessPolicy;
  path: string;
}

// ---- the server -------------------------------------------------------------------------------------

export async function createServer(options: DocxMcpOptions): Promise<McpServer> {
  const role = options.role ?? 'editor';
  if (!ROLES.includes(role)) throw new Error(`Unknown role "${role}" (use ${ROLES.join(', ')})`);
  const author = options.author?.trim() || 'AI assistant';
  const userInstructions = instructionsSection(options.instructions);
  const instructionsLength = options.instructions?.trim().length ?? 0;
  if (instructionsLength > MAX_INSTRUCTIONS) throw new Error(`Instructions have ${instructionsLength} characters; the limit is ${MAX_INSTRUCTIONS}.`);
  const fontFolders = options.fontFolders ?? [];
  const ws = await Workspace.create(options.folders);
  const queue = new Queue();
  const user = { id: 'docxkeun-mcp', name: author };
  const rank = ROLES.indexOf(role);
  const allows = (minimum: McpRole): boolean => rank >= ROLES.indexOf(minimum);
  // docxkeun's own policy knows four roles: a suggester edits as an editor whose every change is tracked.
  const policyRole: Role = role === 'suggester' ? 'editor' : role;
  const editorOptions = role === 'suggester' ? { author, trackChanges: true } : { author };
  const preview = options.preview;

  const server = new McpServer(
    { name: 'docxkeun', version: VERSION },
    {
      instructions: [
        `This server reads and changes Word documents (.docx) in: ${ws.roots.join(', ')}.`,
        'Start with read_document. It numbers paragraphs from 1, including paragraphs in table cells and text boxes; tools that take `paragraph` use those numbers.',
        'Numbers shift after insert_paragraphs, delete_paragraphs, copy_paragraphs, table changes, insert_shape with text, delete_object of a text box and edit_object moving a text box, so read the document again before reusing numbers from earlier.',
        ...(role === 'editor' || role === 'owner'
          ? ['Tables: insert_table makes one and edit_table adds or deletes rows; cells are paragraphs, so edit their text like any paragraph. To duplicate content with its formatting, lists and tables (for example to translate a document onto a new page), use copy_paragraphs and then change the copies\u2019 text, rather than retyping it with insert_paragraphs. For a copy or new version of a whole document in another file (for example a translation), use copy_document and then change the copy; create_document makes a blank document without the original\u2019s formatting.',
             'read_document shows text formatted differently from the rest of its paragraph with <b>, <i>, <u>, <s>, <sup> and <sub> tags, and formatting of the whole paragraph in its description. Keep those tags when you rewrite or translate a paragraph with edit_paragraph, so the bold, underlined and other text stays where it was; the paragraph\u2019s own formatting stays by itself.',
             'Charts: insert_chart adds one; read_chart shows its data as a table, and update_chart changes it (give the full new table of categories and series).',
             'Shapes and text boxes: insert_shape adds one; the text inside a box is numbered right after the paragraph holding it, and every text tool works there (edit_paragraph, replace_text, format_text, comments).',
             'Objects (pictures, charts, shapes, text boxes): list_objects numbers them (\"object 2\", also shown in read_document); edit_object resizes, moves, repositions or describes one, and delete_object removes it.']
          : []),
        ...(role === 'suggester'
          ? ['Every change you make is recorded as a tracked change signed with your name, for a person to accept or reject. You cannot accept or reject changes or switch tracking off; explain each change with a comment.']
          : []),
        'Changes are saved to the file straight away; pass `output` to write a copy instead.',
        `Your role is ${role}: you may ${ROLE_SUMMARY[role]}. Comments you add are signed "${author}".`,
        "Locked regions and the document's own protection are enforced. If an edit is refused, tell the user rather than working around it.",
        ...(preview
          ? [
              `A live preview shows the documents at ${preview.url} and redraws them after every change; give the user get_preview_link's link when they want to watch.`,
              'When the user says "this", "here" or "the selected text" without saying which, call get_selection: it tells you what they selected in the live preview.',
            ]
          : []),
        ...(userInstructions ? [userInstructions] : []),
      ].join('\n'),
    },
  );

  function explain(e: unknown): Error {
    if (e instanceof PermissionError) return new Error(`Not allowed: ${e.message} (this server runs with the ${role} role).`);
    if (e instanceof LockedRegionError) return new Error(`${e.message}. Locked regions cannot be changed; ask the document's owner.`);
    if (e instanceof PasswordRequiredError) return new Error('This document is encrypted. Ask the user for its password and pass it as `password`.');
    if (e instanceof WrongPasswordError) return new Error('Wrong password for this document.');
    if (e instanceof WorkspaceError || e instanceof Error) return e;
    return new Error(String(e));
  }

  function tool<S extends z.ZodObject>(
    name: string,
    config: { title: string; description: string; inputSchema: S; annotations: ToolAnnotations },
    handler: (args: z.output<S>) => Promise<string>,
  ): void {
    const callback = async (args: z.output<S>): Promise<CallToolResult> => {
      try {
        // Reads wait for changes that are still being saved, so they see them.
        if (config.annotations.readOnlyHint) await queue.idle();
        return reply(await handler(args));
      } catch (e) {
        throw explain(e);
      }
    };
    server.registerTool(name, config, callback as never);
  }

  async function open(file: string, pass: string | undefined): Promise<Session> {
    const real = await ws.resolve(file, 'document', true);
    const doc = await DocxDocument.open(await ws.read(real), pass ? { password: pass } : {});
    const policy = new AccessPolicy(doc, { user, role: policyRole });
    return {
      doc,
      editor: new DocxEditor(doc, policy, editorOptions),
      policy,
      path: real,
      name: ws.display(real),
      author,
      entries: listParagraphs(doc),
      loadImage: async (file: string) => ws.read(await ws.resolve(file, 'image', true)),
    };
  }

  /** Save to the original file or to `output`. `password`: undefined keeps the current one, null removes it. */
  async function save(s: Session, args: { output?: string | undefined; overwrite?: boolean | undefined }, newPassword?: string | null): Promise<{ real: string; shown: string }> {
    let dest = s.path;
    if (args.output) {
      dest = await ws.resolve(args.output, 'document', false);
      if (dest !== s.path && !args.overwrite && (await ws.exists(dest))) {
        throw new Error(`${ws.display(dest)} already exists; pass overwrite: true to replace it`);
      }
    }
    const pass = newPassword === undefined ? s.doc.sourcePassword : (newPassword ?? undefined);
    await ws.write(dest, await s.doc.save({ modified: new Date(), ...(pass ? { password: pass } : {}) }));
    return { real: dest, shown: ws.display(dest) };
  }

  /** Open, change, save (one change at a time per server). */
  function change<A extends { path: string; password?: string | undefined; output?: string | undefined; overwrite?: boolean | undefined }>(
    fn: (s: Session, args: A) => Promise<string>,
  ): (args: A) => Promise<string> {
    return (args) =>
      queue.run(async () => {
        const s = await open(args.path, args.password);
        const message = await fn(s, args);
        const saved = await save(s, args);
        preview?.notify({ path: saved.real, message, ...(s.focus?.comment ? { comment: s.focus.comment } : {}) });
        return `${message}\nSaved ${saved.shown}.`;
      });
  }

  // ==== reading =======================================================================================

  tool(
    'list_documents',
    {
      title: 'List Word documents',
      description: 'List the Word documents (.docx, .docm, .dotx, .dotm) in the folders this server may use.',
      inputSchema: z.object({
        folder: z.string().optional().describe('Only look in this folder (inside the allowed folders)'),
        query: z.string().optional().describe('Only files whose name contains this text'),
      }),
      annotations: READ,
    },
    async ({ folder, query }) => {
      const { entries, truncated } = await ws.listDocuments({ ...(folder ? { folder } : {}), ...(query ? { query } : {}) });
      const lines = [`Allowed folders: ${ws.roots.join(', ')}`];
      if (entries.length === 0) lines.push('No Word documents found.');
      for (const d of entries) lines.push(`${ws.roots.length === 1 ? d.relative : d.path}  (${kb(d.size)}, modified ${d.modified.toISOString().slice(0, 16).replace('T', ' ')})`);
      if (truncated) lines.push(`Showing the first ${entries.length}; pass folder or query to narrow it down.`);
      return lines.join('\n');
    },
  );

  tool(
    'get_document_info',
    {
      title: 'Document details',
      description:
        'Properties, page count, word count, protection, encryption, locked regions, editable ranges and the paragraph styles a document has.',
      inputSchema: z.object({ path, password }),
      annotations: READ,
    },
    async (args) => {
      const s = await open(args.path, args.password);
      const lines = [ws.display(s.path)];
      const props = await s.doc.coreProperties();
      const propLines = Object.entries({
        Title: props.title,
        Author: props.creator,
        Subject: props.subject,
        Keywords: props.keywords,
        Created: props.created,
        'Last modified': props.modified,
        'Last modified by': props.lastModifiedBy,
      }).filter(([, v]) => v);
      for (const [k, v] of propLines) lines.push(`${k}: ${v}`);

      let pages = 'unknown';
      try {
        pages = String(layoutDocument(s.doc, { fonts: await fontsFor(s.doc, fontFolders) }).pages.length);
      } catch (e) {
        pages = `unknown (${(e as Error).message})`;
      }
      const threads = commentThreads(s.doc);
      const open_ = threads.filter((t) => !t.comment.resolved).length;
      const links = s.editor.hyperlinks();
      const bookmarks = s.editor.bookmarks().filter((b) => !b.name.startsWith('_'));
      lines.push(
        `Pages: ${pages} (as laid out by docxkeun; Word may differ slightly)`,
        `Paragraphs: ${s.entries.length}, tables: ${tableCount(s.doc)}, words: ${wordCount(s.entries)}`,
        `Comment threads: ${threads.length} (${open_} open, ${threads.length - open_} resolved)`,
        `Hyperlinks: ${links.length}, bookmarks: ${bookmarks.length}`,
        `Encrypted (password to open): ${s.doc.sourcePassword ? 'yes' : 'no'}`,
      );
      const p = readProtection(s.doc);
      lines.push(
        p.enforced && p.mode !== 'none'
          ? `Protection: ${p.mode}${p.hasPassword ? ', with a password' : ', no password'}${p.formatting ? ', formatting restricted' : ''}`
          : 'Protection: none',
      );
      const ranges = s.editor.editableRanges();
      if (ranges.length) lines.push(`Editable ranges: ${ranges.map((r) => `${r.id} (${r.editor ?? r.group ?? 'everyone'})`).join(', ')}`);
      const regions = s.editor.lockedRegions();
      const locks = locksByParagraph(s.entries, regions);
      if (regions.length) {
        lines.push('Locked regions:');
        for (const r of regions) {
          const paras = [...locks].filter(([, list]) => list.some((l) => l.region === r)).map(([n]) => n);
          const where = paras.length ? ` in paragraph${paras.length > 1 ? 's' : ''} ${paras.length > 2 ? `${paras[0]}–${paras[paras.length - 1]}` : paras.join(', ')}` : '';
          lines.push(`  lock ${r.id}${r.title ? ` "${r.title}"` : ''} (${r.mode})${where}: ${quote(r.text, 60)}`);
        }
      }
      lines.push(`Paragraph styles: ${paragraphStyles(s.doc).map((st) => st.name).join(', ')}`);
      lines.push(`Your role: ${role} (you may ${ROLE_SUMMARY[role]})`);
      return lines.join('\n');
    },
  );

  if (preview) {
    tool(
      'get_preview_link',
      {
        title: 'Live preview link',
        description:
          'A link to the live preview, a web page on this computer that shows a document with its pages, comments and lists, and redraws it after every change. Give it to the user when they want to watch.',
        inputSchema: z.object({ path: path.optional().describe('The document to open (default: the most recently changed one)') }),
        annotations: READ,
      },
      async (args) => {
        if (!args.path) return `Live preview: ${preview.url}`;
        const real = await ws.resolve(args.path, 'document', true);
        return `Live preview of ${ws.display(real)}: ${preview.link(real)}`;
      },
    );

    /** The selection made in the live preview, described from the file as it is now. */
    const currentSelection = async (): Promise<string> => {
      const sel = preview.selection();
      if (!sel) return `Nothing is selected in the live preview. The user can select text (or click in a paragraph) at ${preview.url}.`;
      await queue.idle();
      const s = await open(sel.path, undefined);
      // A page showing tracked changes counts deleted text and paragraphs: read its positions in the same view.
      const doc = sel.changes ? await DocxDocument.open(await ws.read(s.path), { showDeletions: true }) : s.doc;
      // The page numbers body paragraphs as its viewer does (paragraphsOf); the AI gets read_document's numbers.
      const paras = sel.story === 'body' ? paragraphsOf(doc.body().blocks) : sel.part && doc.xmlSync(sel.part) ? new DocxEditor(doc, undefined, { part: sel.part }).paragraphs() : [];
      let layout;
      try {
        layout = layoutDocument(doc, { fonts: await fontsFor(doc, fontFolders) });
      } catch {
        layout = undefined;
      }
      const info = selectionInfoFor(doc, paras, { para: sel.start.paragraph - 1, char: sel.start.offset }, { para: sel.end.paragraph - 1, char: sel.end.offset }, {
        story: sel.story,
        ...(sel.part ? { part: sel.part } : {}),
        ...(layout ? { layout } : {}),
      });
      if (!info) return `The selection in ${s.name} is no longer there (the document changed). Ask the user to select again.`;
      const age = Math.round((Date.now() - sel.at) / 1000);
      const when = age < 5 ? 'just now' : age < 120 ? `${age} seconds ago` : `${Math.round(age / 60)} minutes ago`;
      const byView = sel.changes ? readNumbersWithChanges(s.entries, s.doc, doc) : readNumbersByView(s.entries);
      const shownChanges = sel.changes && s.editor.revisions().length > 0
        ? ' The preview shows tracked changes, so the text includes deleted words and offsets count them; read_document shows the changes accepted.'
        : '';
      return `${describeSelection(renumberSelection(info, byView), { document: s.name })}\n(Selected ${when}. Paragraph numbers are read_document's; offsets count characters in the paragraph.${shownChanges})`;
    };

    tool(
      'get_selection',
      {
        title: "The user's selection",
        description:
          'What the user has selected (or where they clicked) in the live preview: the document, paragraph numbers, page and the selected text in «» with some context. ' +
          'Call it when the user says "this", "here", "the selected text" or "this paragraph" without saying which.',
        inputSchema: z.object({}),
        annotations: READ,
      },
      currentSelection,
    );

    server.registerResource(
      'selection',
      'docxkeun://selection',
      { title: "The user's selection in the live preview", description: 'The selected text, its paragraph numbers and page, in the document shown in the live preview.', mimeType: 'text/plain' },
      async (uri) => ({ contents: [{ uri: uri.href, mimeType: 'text/plain', text: await currentSelection() }] }),
    );
  }

  tool(
    'export_pdf',
    {
      title: 'Export to PDF',
      description:
        'Convert a document to PDF with docxkeun\'s own layout engine: embedded fonts, searchable text, links, a bookmark outline from headings, and comments as PDF annotations. ' +
        'Optionally protect the PDF with a password and permissions. An existing PDF at the output path is replaced.',
      inputSchema: z.object({
        path,
        password,
        output: z.string().optional().describe('The .pdf to write (default: next to the document, same name)'),
        comments: z.boolean().optional().describe('Include comments as annotations (default true)'),
        outline: z.boolean().optional().describe('Build a bookmark outline from headings (default true)'),
        pdfPassword: z.string().optional().describe('Password needed to open the PDF'),
        permissions: z
          .object({
            print: z.boolean().optional(),
            copy: z.boolean().optional(),
            modify: z.boolean().optional(),
            annotate: z.boolean().optional(),
          })
          .optional()
          .describe('What PDF readers should allow (default: everything)'),
      }),
      annotations: { readOnlyHint: false, destructiveHint: false, idempotentHint: true, openWorldHint: false },
    },
    async (args) =>
      queue.run(async () => {
        const s = await open(args.path, args.password);
        s.policy.assert('export');
        const dest = await ws.resolve(args.output ?? s.path.replace(/\.[^./\\]+$/, '.pdf'), 'pdf', false);
        const fonts = await fontsFor(s.doc, fontFolders);
        const layout = layoutDocument(s.doc, { fonts });
        const secure = args.pdfPassword !== undefined || args.permissions !== undefined;
        const pdf = await exportPdf(s.doc, {
          fonts,
          layout,
          comments: args.comments ?? true,
          outline: args.outline ?? true,
          ...(secure ? { encryption: pdfSecurity({ ...(args.pdfPassword ? { userPassword: args.pdfPassword } : {}), ...(args.permissions ? { permissions: args.permissions } : {}) }) } : {}),
        });
        await ws.write(dest, pdf);
        return `Exported ${plural(layout.pages.length, 'page')} to ${ws.display(dest)} (${kb(pdf.length)})${args.pdfPassword ? ', password protected' : ''}.`;
      }),
  );

  // ==== the shared document operations (operations.ts), as far as the role allows ===================

  for (const op of operationsFor(role)) {
    const inputSchema = z.object({ path, password, ...op.schema.shape, ...(op.kind === 'read' ? {} : saveFields) });
    const annotations = op.kind === 'read' ? READ : op.kind === 'add' ? ADD : CHANGE;
    type Args = Record<string, unknown> & { path: string; password?: string; output?: string; overwrite?: boolean };
    const run = async (s: Session, args: Args): Promise<string> => {
      // The live preview shimmers the paragraphs the AI works on before the result comes in.
      if (preview) {
        const objects = args.object !== undefined ? numberedObjects(s.editor, s.entries) : undefined;
        const targets = targetParagraphs(s.entries, op.name, args, objects);
        preview.activity({
          path: s.path,
          paragraphs: viewerParagraphs(s.entries, targets),
          read: readParagraphs(s.entries, targets),
          mode: op.kind === 'read' ? 'read' : 'edit',
          label: activityLabel(op.name),
        });
      }
      return op.run(s, args as never);
    };
    const save_ = change<Args>(run);
    tool(op.name, { title: op.title, description: describeOperation(op, { author, role }), inputSchema, annotations }, async (input) => {
      const args = input as Args;
      return op.kind === 'read' ? run(await open(args.path, args.password), args) : save_(args);
    });
  }

  // ==== editing (editor and up) =======================================================================

  if (allows('editor')) {
    tool(
      'create_document',
      {
        title: 'Create a document',
        description:
          'Create a new, blank Word document from paragraphs of plain text. The new document has the styles Normal, Heading 1 and Heading 2 and none of any other document\'s formatting. To copy a document, use copy_document (an exact copy), never retype it here.',
        inputSchema: z.object({
          path,
          paragraphs: z.array(newParagraph).describe('The paragraphs, in order'),
          overwrite,
        }),
        annotations: ADD,
      },
      async (args) =>
        queue.run(async () => {
          const dest = await ws.resolve(args.path, 'document', false);
          if (!args.overwrite && (await ws.exists(dest))) throw new Error(`${ws.display(dest)} already exists; pass overwrite: true to replace it`);
          const doc = await DocxDocument.create();
          const editor = new DocxEditor(doc, new AccessPolicy(doc, { user, role: policyRole }));
          const styles = args.paragraphs.map((p) => (p.style ? resolveStyle(doc, p.style) : undefined));
          const first = listParagraphs(doc)[0]!.para.el;
          let prev = first;
          args.paragraphs.forEach((p, i) => {
            const style = styles[i];
            const segments = parseMarkup(p.text);
            if (i === 0) {
              if (style) editor.setParagraphStyle(first, style);
              writeTagged(editor, first, segments);
            } else {
              prev = editor.insertParagraph({ after: prev }, segments.map((x) => x.text).join(''), style ? { style } : {});
              tagParagraph(editor, prev, segments);
            }
          });
          await ws.write(dest, await doc.save({ modified: new Date() }));
          const message = `Created ${ws.display(dest)} with ${plural(Math.max(args.paragraphs.length, 1), 'paragraph')}.`;
          preview?.notify({ path: dest, message });
          return message;
        }),
    );

    tool(
      'copy_document',
      {
        title: 'Copy a document',
        description:
          'Copy a document to a new file, exactly as it is: formatting, styles, tables, lists, pictures, headers, footers, comments and page setup. Use it whenever the user wants a copy or a new version of a document (for example a translation in a new file), then change the copy with the other tools.',
        inputSchema: z.object({
          path,
          to: z.string().describe('The new file, a .docx inside the allowed folders'),
          overwrite: z.boolean().optional().describe('Allow `to` to replace an existing file (default false)'),
        }),
        annotations: ADD,
      },
      async (args) =>
        queue.run(async () => {
          const source = await ws.resolve(args.path, 'document', true);
          const dest = await ws.resolve(args.to, 'document', false);
          if (dest === source) throw new Error(`${ws.display(dest)} is the same file as the original; give the copy another name`);
          if (!args.overwrite && (await ws.exists(dest))) throw new Error(`${ws.display(dest)} already exists; pass overwrite: true to replace it`);
          // The bytes as they are: nothing is re-saved, so nothing can be lost.
          const bytes = await ws.read(source);
          await ws.write(dest, bytes);
          const encrypted = bytes[0] === 0xd0 && bytes[1] === 0xcf && bytes[2] === 0x11 && bytes[3] === 0xe0;
          const shown = ws.display(dest);
          const message =
            `Copied ${ws.display(source)} to ${shown}: an exact copy with all its formatting, tables, lists, pictures, headers, footers and comments. ` +
            `Change the copy with path "${shown}".${encrypted ? ' It opens with the same password.' : ''}`;
          preview?.notify({ path: dest, message: `Copied from ${ws.display(source)}.` });
          return message;
        }),
    );

  }

  // ==== protection (owner) ============================================================================

  if (allows('owner')) {
    tool(
      'protect_document',
      {
        title: 'Protect a document',
        description:
          "Set Word's editing restrictions: readOnly, comments (only comments allowed), forms (only form fields), trackedChanges, or none to remove protection. " +
          'This restricts editing in Word but is not encryption; use set_open_password for that.',
        inputSchema: z.object({
          path,
          password,
          mode: z.enum(['none', 'readOnly', 'comments', 'forms', 'trackedChanges']),
          protectionPassword: z.string().optional().describe('Password for turning protection off later (setting), or the current one (removing)'),
          ...saveFields,
        }),
        annotations: CHANGE,
      },
      change(async (s, args) => {
        s.policy.assert('protect');
        if (args.mode === 'none') {
          await removeProtection(s.doc, args.protectionPassword ?? '');
          return 'Protection removed.';
        }
        await setProtection(s.doc, args.mode, args.protectionPassword ? { password: args.protectionPassword } : {});
        return `Protection set to ${args.mode}${args.protectionPassword ? ' with a password' : ''}.`;
      }),
    );

    tool(
      'set_open_password',
      {
        title: 'Set or remove the open password',
        description: 'Encrypt a document so it needs a password to open (AES-256, opens in Word), or remove that password.',
        inputSchema: z.object({
          path,
          password: z.string().optional().describe('The current password, if the document is already encrypted'),
          action: z.enum(['set', 'remove']),
          newPassword: z.string().min(1).optional().describe('The new password (for set)'),
          ...saveFields,
        }),
        annotations: CHANGE,
      },
      async (args) =>
        queue.run(async () => {
          if (args.action === 'set' && !args.newPassword) throw new Error('`set` needs newPassword.');
          const s = await open(args.path, args.password);
          s.policy.assert('protect');
          if (args.action === 'remove' && !s.doc.sourcePassword) throw new Error('This document has no open password.');
          const saved = await save(s, args, args.action === 'set' ? args.newPassword! : null);
          const message = `${args.action === 'set' ? 'Encrypted' : 'Removed the password from'} ${saved.shown}.`;
          preview?.notify({ path: saved.real, message });
          return message;
        }),
    );
  }

  return server;
}
