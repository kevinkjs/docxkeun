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
  /** A comment to scroll to after the change. */
  comment?: string;
}

export interface Toolbox {
  definitions: BetaTool[];
  run(name: string, input: unknown): Promise<ToolOutcome>;
  /** Whether a tool reads, adds or changes (undefined for unknown tools). */
  kind(name: string): OperationKind | undefined;
}

function inputSchema(op: Operation): BetaTool.InputSchema {
  const schema = z.toJSONSchema(op.schema) as Record<string, unknown>;
  delete schema.$schema;
  return schema as BetaTool.InputSchema;
}

function explain(e: unknown): string {
  if (e instanceof PermissionError) return `Not allowed: ${e.message}. The document's protection does not allow this change.`;
  if (e instanceof LockedRegionError) return `${e.message}. Locked regions cannot be changed here; tell the user.`;
  return e instanceof Error ? e.message : String(e);
}

/** Tools that act on whatever document is open when they run. */
export function createToolbox(current: () => OpenDocument | undefined): Toolbox {
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

    async run(name, input) {
      const op = byName.get(name);
      if (!op) return { content: `There is no tool named ${name}.`, isError: true, kind: 'read' };
      const parsed = op.schema.safeParse(input);
      if (!parsed.success) {
        return { content: JSON.stringify({ INVALID_JSON: JSON.stringify(input), problem: z.prettifyError(parsed.error) }), isError: true, kind: op.kind };
      }
      const open = current();
      if (!open) return { content: 'No document is open. Ask the user to open a .docx first.', isError: true, kind: op.kind };
      const ctx = contextFor(open.doc, editorFor(open.doc), open.name, AI_AUTHOR);
      try {
        const content = await op.run(ctx, parsed.data);
        return { content, isError: false, kind: op.kind, ...(ctx.focus?.comment ? { comment: ctx.focus.comment } : {}) };
      } catch (e) {
        return { content: explain(e), isError: true, kind: op.kind };
      }
    },
  };
}
