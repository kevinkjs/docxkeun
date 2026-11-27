/**
 * Claude's tools in the chat: the document operations shared with docxkeun-mcp, run on the document
 * open in the viewer. Every input is checked against the operation's schema before it runs, since
 * streamed tool input is not validated by the API.
 */
import type { BetaTool } from '@anthropic-ai/sdk/resources/beta/messages/messages';
import { z } from 'zod';
import type { DocxDocument } from 'docxkeun';
import { DocxEditor, LockedRegionError } from 'docxkeun/editor';
import { AccessPolicy, PermissionError } from 'docxkeun/security';
import { contextFor, describeOperation, operationsFor, type Operation, type OperationKind } from '../../mcp/src/operations.js';

/** What the AI may do in the chat: everything an editor may (not locks, protection or passwords). */
export const AI_ROLE = 'editor';
export const AI_AUTHOR = 'Claude';

export interface OpenDocument {
  doc: DocxDocument;
  name: string;
}

export interface ToolOutcome {
  content: string;
  isError: boolean;
  kind: OperationKind;
  /** The document the call ran on (absent when it could not tell which). */
  document?: OpenDocument;
  /** A comment to scroll to after the change. */
  comment?: string;
}

export interface Toolbox {
  definitions: BetaTool[];
  run(name: string, input: unknown): Promise<ToolOutcome>;
  /** Whether a tool reads, adds or changes (undefined for unknown tools). */
  kind(name: string): OperationKind | undefined;
  /** The document a call (or the part of it written so far) is about; undefined while that is unclear. */
  target(input: unknown): OpenDocument | undefined;
}

/** Every tool takes the document it works on, by file name: needed when two are open. */
const DOCUMENT_FIELD = {
  type: 'string',
  description: 'The file name of the document to work on, as the user’s message lists it. Needed when two documents are open; may be left out when only one is.',
};

function inputSchema(op: Operation): BetaTool.InputSchema {
  const schema = z.toJSONSchema(op.schema) as Record<string, unknown> & { properties?: Record<string, unknown> };
  delete schema.$schema;
  schema.properties = { document: DOCUMENT_FIELD, ...schema.properties };
  return schema as BetaTool.InputSchema;
}

/** The open document a call names, or why it can't run: none open, two open and no name, or a name not open. */
function pick(open: readonly OpenDocument[], name: unknown): OpenDocument | string {
  const names = open.map((d) => d.name).join(' and ');
  if (typeof name === 'string' && name.trim()) {
    const wanted = name.trim().toLowerCase();
    const found = open.find((d) => d.name.toLowerCase() === wanted);
    if (found) return found;
    return open.length ? `No open document is called ${name.trim()}. Open: ${names}.` : 'No document is open. Ask the user to open a .docx first.';
  }
  if (open.length === 1) return open[0]!;
  if (!open.length) return 'No document is open. Ask the user to open a .docx first.';
  return `Two documents are open: ${names}. Say which one with \`document\`.`;
}

/** A call's input without the document field, which the operations themselves don't take. */
function split(input: unknown): { name: unknown; args: unknown } {
  if (!input || typeof input !== 'object' || Array.isArray(input)) return { name: undefined, args: input };
  const { document: name, ...args } = input as Record<string, unknown>;
  return { name, args };
}

function explain(e: unknown): string {
  if (e instanceof PermissionError) return `Not allowed: ${e.message}. The document's protection does not allow this change.`;
  if (e instanceof LockedRegionError) return `${e.message}. Locked regions cannot be changed here; tell the user.`;
  return e instanceof Error ? e.message : String(e);
}

/** Tools that act on the open documents (at most two), as they are when the tools run. */
export function createToolbox(documents: () => readonly OpenDocument[]): Toolbox {
  const ops = operationsFor(AI_ROLE);
  const byName = new Map(ops.map((op) => [op.name, op]));
  const editors = new WeakMap<DocxDocument, DocxEditor>();
  const editorFor = (doc: DocxDocument): DocxEditor => {
    let editor = editors.get(doc);
    if (!editor) {
      editor = new DocxEditor(doc, new AccessPolicy(doc, { user: { id: 'claude', name: AI_AUTHOR }, role: AI_ROLE }));
      editors.set(doc, editor);
    }
    return editor;
  };

  return {
    definitions: ops.map((op) => ({
      name: op.name,
      description: describeOperation(op, { author: AI_AUTHOR, role: AI_ROLE }),
      input_schema: inputSchema(op),
      // Streamed requests: let large inputs (a long paragraph) stream as they are generated.
      eager_input_streaming: true,
    })),

    kind: (name) => byName.get(name)?.kind,

    target(input) {
      const found = pick(documents(), split(input).name);
      return typeof found === 'string' ? undefined : found;
    },

    async run(name, input) {
      const op = byName.get(name);
      if (!op) return { content: `There is no tool named ${name}.`, isError: true, kind: 'read' };
      const { name: docName, args } = split(input);
      const parsed = op.schema.safeParse(args);
      if (!parsed.success) {
        return { content: JSON.stringify({ INVALID_JSON: JSON.stringify(input), problem: z.prettifyError(parsed.error) }), isError: true, kind: op.kind };
      }
      const open = pick(documents(), docName);
      if (typeof open === 'string') return { content: open, isError: true, kind: op.kind };
      const ctx = contextFor(open.doc, editorFor(open.doc), open.name, AI_AUTHOR);
      try {
        const content = await op.run(ctx, parsed.data);
        return { content, isError: false, kind: op.kind, document: open, ...(ctx.focus?.comment ? { comment: ctx.focus.comment } : {}) };
      } catch (e) {
        return { content: explain(e), isError: true, kind: op.kind, document: open };
      }
    },
  };
}
