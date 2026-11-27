/**
 * The conversation with Claude: a streaming tool-use loop on the Messages API (beta endpoint, for
 * server-side refusal fallbacks). One call to runTurn handles one user message: it streams Claude's
 * text, runs the tools Claude asks for, sends the results back, and repeats until Claude is done.
 */
import Anthropic from '@anthropic-ai/sdk';
import type {
  BetaContentBlock,
  BetaContentBlockParam,
  BetaMessage,
  BetaMessageParam,
  BetaRawMessageStreamEvent,
  BetaToolResultBlockParam,
  BetaToolUseBlock,
  MessageCreateParamsNonStreaming,
} from '@anthropic-ai/sdk/resources/beta/messages/messages';
import { instructionsSection } from '../../mcp/src/instructions.js';
import type { Toolbox, ToolOutcome } from './tools.js';

export interface ModelChoice {
  id: string;
  label: string;
  /** Server-side refusal fallback ("default" routing). Not available for Claude Haiku 5.5. */
  fallbacks: boolean;
}

export const MODELS: ModelChoice[] = [
  { id: 'claude-opus-5-5', label: 'Claude Opus 5.5', fallbacks: true },
  { id: 'claude-sonnet-5-5', label: 'Claude Sonnet 5.5', fallbacks: true },
  { id: 'claude-haiku-5-5', label: 'Claude Haiku 5.5', fallbacks: false },
];

export const SYSTEM_PROMPT = `You are the assistant in docxkeun chat. The user has a Word document open in a live viewer next to this chat, and every change you make with your tools appears there immediately.

How to work:
- Use read_document before you talk about the document's content or change it. Paragraph numbers shift after you insert or delete paragraphs, add a text box with text, delete a text box or move one, so read again before reusing numbers.
- Make the changes the user asks for directly; they can undo all of your changes from one reply with a single click. Ask first only when a request is unclear or could reasonably mean very different things.
- Prefer targeted edits (replace_text, edit_paragraph, format_text, set_list) over rewriting large parts of the document.
- To duplicate content with its formatting, lists and tables (for example to translate the document onto a new page), use copy_paragraphs and then change the copies' text, instead of retyping it with insert_paragraphs. Table cells are paragraphs; insert_table makes a new table and edit_table adds or deletes rows.
- read_document shows text formatting as <b>, <i>, <u>, <s>, <sup> and <sub> tags (formatting of a whole paragraph is in its description). Keep the tags when you rewrite or translate text with edit_paragraph, so bold, underlined and other formatted words stay formatted.
- Charts: insert_chart adds one; read_chart shows its data as a table and update_chart changes it (give the full new table of categories and series).
- Shapes and text boxes: insert_shape adds one. The text inside a box is numbered right after the paragraph holding it, and every text tool works there.
- Objects (pictures, charts, shapes, text boxes): list_objects numbers them ("object 2", also shown in read_document); edit_object resizes, moves, repositions or describes one, and delete_object removes it.
- When the user's message says what they selected in the viewer, "this" or "here" means that selection.
- Two documents can be open side by side; each message lists the open ones. When two are open, pass its file name as \`document\` in every tool call, and read each document before you work on it. Paragraph numbers belong to one document: never use one document's numbers in the other. To move content between them, read it from one and write it into the other (insert_paragraphs or edit_paragraph, keeping the formatting tags).
- Comments you add are signed "Claude". Use comments for suggestions the user should decide on.
- Locked regions and the document's own protection are enforced by the tools. If a change is refused, explain why in plain words and suggest an alternative, such as a comment; never try to work around it.

How to reply:
- Short, plain language, in the language the user writes in. Say what you changed, not how the tools work.
- Quote only the sentences that matter; never paste the whole document.`;

/** The part of the SDK's stream this loop needs (an interface, so tests and the demo can fake it). */
export interface MessageStream {
  on(event: 'text', listener: (delta: string) => void): unknown;
  /** Raw events with the message so far: tool inputs arrive here while they are being written. */
  on(event: 'streamEvent', listener: (event: BetaRawMessageStreamEvent, snapshot: BetaMessage) => void): unknown;
  finalMessage(): Promise<BetaMessage>;
}

export type StreamFactory = (params: MessageCreateParamsNonStreaming, options: { signal?: AbortSignal }) => MessageStream;

/** The real thing: client.beta.messages.stream. */
export function anthropicStream(client: Anthropic): StreamFactory {
  return (params, options) => client.beta.messages.stream(params, options);
}

export interface TurnEvents {
  /** A new assistant step begins (one API response). */
  step?(): void;
  text(delta: string): void;
  /** A tool call is being written: its name and the input so far (partly parsed). */
  toolDraft?(name: string, input: unknown): void;
  /** A tool call is about to run. */
  toolStart?(name: string, input: unknown): void;
  tool(name: string, input: unknown, outcome: ToolOutcome): void;
  notice(text: string, kind: 'refusal' | 'cut' | 'error'): void;
}

export interface TurnOptions {
  stream: StreamFactory;
  model: ModelChoice;
  tools: Toolbox;
  /** The conversation; runTurn appends to it. On a refusal the whole turn is taken back out. */
  messages: BetaMessageParam[];
  events: TurnEvents;
  signal?: AbortSignal;
  maxSteps?: number;
  /** The user's own instructions from the Instructions dialog; added after the built-in prompt for every step of this turn. */
  instructions?: string;
}

/** Block types that must not be echoed from before a mid-output fallback (the docs' echo rule). */
const MODEL_INTERNAL = new Set(['thinking', 'redacted_thinking', 'tool_use', 'server_tool_use']);

/**
 * The assistant content to send back next time, and the tool calls to run. After a mid-output
 * refusal fallback, thinking and tool calls from before the last `fallback` block are dropped
 * (they belonged to the model that declined); the `fallback` block itself stays where it was.
 */
export function afterFallback(content: BetaContentBlock[]): { echo: BetaContentBlockParam[]; toolUses: BetaToolUseBlock[] } {
  let boundary = -1;
  content.forEach((b, i) => {
    if (b.type === 'fallback') boundary = i;
  });
  const echo = content.filter((b, i) => i >= boundary || !MODEL_INTERNAL.has(b.type)) as BetaContentBlockParam[];
  const toolUses = content.filter((b, i): b is BetaToolUseBlock => b.type === 'tool_use' && i > boundary);
  return { echo, toolUses };
}

/** The built-in prompt, followed by the user's instructions when they wrote any. */
export function systemPrompt(instructions?: string): string {
  const section = instructionsSection(instructions);
  return section ? `${SYSTEM_PROMPT}\n\n${section}` : SYSTEM_PROMPT;
}

function requestFor(options: TurnOptions): MessageCreateParamsNonStreaming {
  return {
    model: options.model.id,
    max_tokens: 64000,
    system: systemPrompt(options.instructions),
    tools: options.tools.definitions,
    messages: options.messages,
    // Cache the system prompt, tools and conversation so far: each step of a turn re-sends them.
    cache_control: { type: 'ephemeral' },
    output_config: { effort: 'medium' },
    ...(options.model.fallbacks ? { betas: ['server-side-fallback-2026-07-01'], fallbacks: 'default' as const } : {}),
  };
}

/** Run one user turn (the user message must already be the last entry of `messages`). */
export async function runTurn(options: TurnOptions): Promise<void> {
  const { messages, events } = options;
  const start = messages.length;
  const takeBack = () => messages.splice(start - 1);
  const maxSteps = options.maxSteps ?? 30;
  let retries = 0;

  for (let step = 0; step < maxSteps; step++) {
    events.step?.();
    const stream = options.stream(requestFor(options), options.signal ? { signal: options.signal } : {});
    stream.on('text', (delta) => events.text(delta));
    if (events.toolDraft) {
      stream.on('streamEvent', (event, snapshot) => {
        if (event.type !== 'content_block_delta' || event.delta.type !== 'input_json_delta') return;
        const block = snapshot.content[event.index];
        if (block?.type === 'tool_use') events.toolDraft!(block.name, block.input);
      });
    }
    let message: BetaMessage;
    try {
      message = await stream.finalMessage();
      retries = 0;
    } catch (err) {
      // A tool input the SDK could not parse at all, or a dropped connection: re-issue the step.
      // API errors (auth, rate limits, overload) and a user abort are the caller's to handle.
      if (err instanceof Anthropic.APIError || err instanceof Anthropic.APIUserAbortError || options.signal?.aborted || retries++ >= 2) throw err;
      step--;
      continue;
    }

    if (message.stop_reason === 'refusal') {
      takeBack();
      const category = message.stop_details?.category;
      events.notice(`Claude declined this request${category ? ` (${category})` : ''}. Changes already made in this reply stay until you undo them.`, 'refusal');
      return;
    }

    const { echo, toolUses } = afterFallback(message.content);
    if (toolUses.length && (message.stop_reason === 'max_tokens' || message.stop_reason === 'model_context_window_exceeded')) {
      // A tool input cut off mid-way can still look valid: never run it.
      takeBack();
      events.notice('The reply was cut off before Claude finished. Try a smaller request.', 'cut');
      return;
    }
    messages.push({ role: 'assistant', content: echo });
    if (message.stop_reason === 'pause_turn') continue;
    if (message.stop_reason !== 'tool_use' || toolUses.length === 0) {
      if (message.stop_reason === 'max_tokens') events.notice('The reply hit the length limit.', 'cut');
      return;
    }

    const results: BetaToolResultBlockParam[] = [];
    for (const use of toolUses) {
      events.toolStart?.(use.name, use.input);
      const outcome = await options.tools.run(use.name, use.input);
      events.tool(use.name, use.input, outcome);
      results.push({ type: 'tool_result', tool_use_id: use.id, content: outcome.content, ...(outcome.isError ? { is_error: true } : {}) });
    }
    // All results of one step go back in a single user message.
    messages.push({ role: 'user', content: results });
  }
  events.notice(`Stopped after ${maxSteps} steps.`, 'cut');
}
