/**
 * Development only (open the dev server with ?mock): a scripted stand-in for Claude, so the whole UI
 * can be tried without an API key or cost. It reads the document, makes one change that fits the
 * request, and reports back, streaming its text like the real API. Never part of a build.
 */
import type { BetaMessage, BetaMessageParam, BetaRawMessageStreamEvent } from '@anthropic-ai/sdk/resources/beta/messages/messages';
import type { MessageStream, StreamFactory } from './agent.js';

type Block = { type: 'text'; text: string } | { type: 'tool_use'; id: string; name: string; input: Record<string, unknown> };

let ids = 0;

function lastUserText(messages: BetaMessageParam[]): string {
  for (let i = messages.length - 1; i >= 0; i--) {
    const m = messages[i]!;
    if (m.role !== 'user') continue;
    if (typeof m.content === 'string') return m.content;
    const text = m.content.find((b) => b.type === 'text');
    if (text && text.type === 'text') return text.text;
  }
  return '';
}

function lastToolResult(messages: BetaMessageParam[]): string | undefined {
  const m = messages[messages.length - 1];
  if (!m || m.role !== 'user' || typeof m.content === 'string') return undefined;
  const r = m.content.find((b) => b.type === 'tool_result');
  if (!r || r.type !== 'tool_result') return undefined;
  return typeof r.content === 'string' ? r.content : '';
}

/** Pick the next scripted step from the conversation so far. */
function script(messages: BetaMessageParam[]): Block[] {
  const request = lastUserText(messages).toLowerCase();
  const result = lastToolResult(messages);
  const tool = (name: string, input: Record<string, unknown>): Block => ({ type: 'tool_use', id: `toolu_mock_${++ids}`, name, input });
  if (result === undefined) return [{ type: 'text', text: 'Let me read the document first.' }, tool('read_document', {})];
  if (result.includes(': ') && /paragraphs?, \d+ tables?/.test(result)) {
    const listItems = [...result.matchAll(/^\[(\d+)\] \((?:list item|bullet)[^)]*\)/gm)].map((m) => Number(m[1]));
    if (/number|list/.test(request) && listItems.length) {
      return [{ type: 'text', text: 'I will turn the list into a numbered one.' }, tool('set_list', { from: listItems[0], to: listItems[listItems.length - 1], format: 'decimal' })];
    }
    const first = /^\[(\d+)\] [^\n(]*\S/m.exec(result.split('\n\n')[1] ?? '');
    const target = Number(first?.[1] ?? 1);
    if (/bold/.test(request)) return [{ type: 'text', text: `Making paragraph ${target} bold.` }, tool('format_text', { paragraph: target, format: { bold: true } })];
    return [{ type: 'text', text: 'Adding a comment with a suggestion.' }, tool('add_comment', { paragraph: target, comment: 'Mock suggestion: consider making this clearer.' })];
  }
  return [{ type: 'text', text: `Done. ${result}\n\nThis is the **mock** assistant (dev only); open the page without \`?mock\` to talk to Claude.` }];
}

export const mockStream: StreamFactory = (params, options): MessageStream => {
  const blocks = script(params.messages);
  const listeners: Array<(delta: string) => void> = [];
  const rawListeners: Array<(event: BetaRawMessageStreamEvent, snapshot: BetaMessage) => void> = [];
  const wait = (ms: number) => {
    if (options.signal?.aborted) throw new DOMException('Aborted', 'AbortError');
    return new Promise((r) => setTimeout(r, ms));
  };
  return {
    on(event: string, listener: (...args: never[]) => void) {
      if (event === 'text') listeners.push(listener as (delta: string) => void);
      if (event === 'streamEvent') rawListeners.push(listener as (event: BetaRawMessageStreamEvent, snapshot: BetaMessage) => void);
      return this;
    },
    async finalMessage() {
      const content: Block[] = [];
      for (const [index, block] of blocks.entries()) {
        if (block.type === 'text') {
          for (const word of block.text.split(/(?<= )/)) {
            await wait(25);
            for (const l of listeners) l(word);
          }
          content.push(block);
          continue;
        }
        // Tool input "being written", so the demo shows the viewer's working animation too.
        content.push({ ...block, input: {} });
        const json = JSON.stringify(block.input);
        for (let i = 8; i <= json.length + 8; i += 8) {
          await wait(60);
          (content[index] as Extract<Block, { type: 'tool_use' }>).input = block.input;
          const event = { type: 'content_block_delta', index, delta: { type: 'input_json_delta', partial_json: json.slice(i - 8, i) } } as unknown as BetaRawMessageStreamEvent;
          for (const l of rawListeners) l(event, { content } as unknown as BetaMessage);
        }
      }
      const hasTool = blocks.some((b) => b.type === 'tool_use');
      return {
        id: `msg_mock_${++ids}`,
        type: 'message',
        role: 'assistant',
        model: params.model,
        content: blocks,
        stop_reason: hasTool ? 'tool_use' : 'end_turn',
        stop_sequence: null,
        usage: { input_tokens: 0, output_tokens: 0 },
      } as unknown as BetaMessage;
    },
  };
};
