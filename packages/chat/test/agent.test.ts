/** The chat's tool-use loop, driven by a scripted stream instead of the API. */
import Anthropic from '@anthropic-ai/sdk';
import type { BetaMessage, BetaMessageParam, MessageCreateParamsNonStreaming } from '@anthropic-ai/sdk/resources/beta/messages/messages';
import { describe, expect, it } from 'vitest';
import { DocxDocument } from 'docxkeun';
import { buildDocx, p } from '../../../test/helpers/docx.js';
import { afterFallback, MODELS, runTurn, SYSTEM_PROMPT, systemPrompt, type MessageStream, type StreamFactory } from '../src/agent.js';
import { createToolbox, type OpenDocument } from '../src/tools.js';
import { instructionsSection } from '../../mcp/src/instructions.js';
import { listParagraphs } from '../../mcp/src/outline.js';

type Block = Record<string, unknown> & { type: string };

function message(content: Block[], stop_reason: string, extra: Record<string, unknown> = {}): BetaMessage {
  return { id: 'msg', type: 'message', role: 'assistant', model: 'claude-opus-5-5', content, stop_reason, stop_sequence: null, usage: {}, ...extra } as unknown as BetaMessage;
}

/** A stream factory that answers with the scripted messages in order, recording each request. */
function scripted(...replies: Array<BetaMessage | Error>) {
  const requests: MessageCreateParamsNonStreaming[] = [];
  const factory: StreamFactory = (params) => {
    requests.push(structuredClone(params));
    const reply = replies.shift();
    if (!reply) throw new Error('no more scripted replies');
    const listeners: Array<(d: string) => void> = [];
    const stream: MessageStream = {
      on(e: string, l: (...args: never[]) => void) {
        if (e === 'text') listeners.push(l as (d: string) => void);
        return stream;
      },
      async finalMessage() {
        if (reply instanceof Error) throw reply;
        for (const b of reply.content) if (b.type === 'text') for (const l of listeners) l(b.text);
        return reply;
      },
    };
    return stream;
  };
  return { factory, requests };
}

const tool = (id: string, name: string, input: unknown): Block => ({ type: 'tool_use', id, name, input });
const text = (t: string): Block => ({ type: 'text', text: t });

async function setup() {
  const doc = await DocxDocument.open(await buildDocx({ body: p('Hello world') + p('Second paragraph') }));
  let open: OpenDocument | undefined = { doc, name: 'letter.docx' };
  const toolbox = createToolbox(() => (open ? [open] : []));
  const messages: BetaMessageParam[] = [{ role: 'user', content: 'Make it bold' }];
  const seen: Array<{ name: string; content: string; isError: boolean }> = [];
  const notices: string[] = [];
  let streamed = '';
  const events = {
    text: (d: string) => (streamed += d),
    tool: (name: string, _input: unknown, o: { content: string; isError: boolean }) => seen.push({ name, content: o.content, isError: o.isError }),
    notice: (t: string) => notices.push(t),
  };
  return { doc, toolbox, messages, seen, notices, events, streamed: () => streamed, close: () => (open = undefined) };
}

const texts = (doc: DocxDocument) => listParagraphs(doc).map((e) => e.text);

describe('the built-in prompt', () => {
  it('tells Claude how to work with two documents open', () => {
    expect(SYSTEM_PROMPT).toContain('Two documents can be open side by side');
    expect(SYSTEM_PROMPT).toContain('pass its file name as `document` in every tool call');
    expect(SYSTEM_PROMPT).toContain('Paragraph numbers belong to one document');
  });
});

describe('runTurn', () => {
  it('reports tool calls while they are written and before they run (for the viewer animation)', async () => {
    const s = await setup();
    const { factory } = scripted(message([tool('t1', 'edit_paragraph', { paragraph: 2, text: 'New second paragraph' })], 'tool_use'), message([text('Done.')], 'end_turn'));
    // The scripted stream also plays the tool input as stream events, two halves of the JSON.
    const drafting: StreamFactory = (params, options) => {
      const inner = factory(params, options);
      const raw: Array<(e: never, snap: never) => void> = [];
      return {
        on(e: string, l: (...args: never[]) => void) {
          if (e === 'streamEvent') raw.push(l as (e: never, snap: never) => void);
          else inner.on('text', l as (d: string) => void);
          return this;
        },
        async finalMessage() {
          const m = await inner.finalMessage();
          m.content.forEach((b, index) => {
            if (b.type !== 'tool_use') return;
            for (const partial of [{ paragraph: 2 }, b.input]) {
              const snapshot = { content: m.content.map((c, i) => (i === index ? { ...c, input: partial } : c)) };
              for (const l of raw) l({ type: 'content_block_delta', index, delta: { type: 'input_json_delta', partial_json: '' } } as never, snapshot as never);
            }
          });
          return m;
        },
      };
    };
    const calls: string[] = [];
    await runTurn({
      stream: drafting,
      model: MODELS[0]!,
      tools: s.toolbox,
      messages: s.messages,
      events: { ...s.events, toolDraft: (name, input) => calls.push(`draft ${name} ${JSON.stringify(input)}`), toolStart: (name) => calls.push(`start ${name} ${texts(s.doc)[1]}`) },
    });
    expect(calls).toEqual([
      'draft edit_paragraph {"paragraph":2}',
      'draft edit_paragraph {"paragraph":2,"text":"New second paragraph"}',
      // Before it runs, the document is still unchanged.
      'start edit_paragraph Second paragraph',
    ]);
    expect(s.toolbox.kind('edit_paragraph')).toBe('change');
    expect(s.toolbox.kind('read_document')).toBe('read');
  });

  it('runs tools, sends the results back and stops when Claude is done', async () => {
    const s = await setup();
    const { factory, requests } = scripted(
      message([text('Reading.'), tool('t1', 'read_document', {})], 'tool_use'),
      message([tool('t2', 'replace_text', { find: 'Hello', replace: 'Halo' })], 'tool_use'),
      message([text('Done.')], 'end_turn'),
    );
    await runTurn({ stream: factory, model: MODELS[0]!, tools: s.toolbox, messages: s.messages, events: s.events });

    expect(texts(s.doc)).toEqual(['Halo world', 'Second paragraph']);
    expect(s.seen.map((t) => t.name)).toEqual(['read_document', 'replace_text']);
    expect(s.streamed()).toBe('Reading.Done.');
    expect(s.messages.map((m) => m.role)).toEqual(['user', 'assistant', 'user', 'assistant', 'user', 'assistant']);
    const results = s.messages[2]!.content as Array<{ type: string; tool_use_id: string; content: string }>;
    expect(results[0]).toMatchObject({ type: 'tool_result', tool_use_id: 't1' });
    expect(results[0]!.content).toContain('[1] Hello world');

    const first = requests[0]!;
    expect(first).toMatchObject({ model: 'claude-opus-5-5', max_tokens: 64000, cache_control: { type: 'ephemeral' }, output_config: { effort: 'medium' }, fallbacks: 'default' });
    expect((first as { betas?: string[] }).betas).toEqual(['server-side-fallback-2026-07-01']);
    expect(first.tools!.every((t) => 'eager_input_streaming' in t && t.eager_input_streaming === true)).toBe(true);
  });

  it('sends no fallbacks for Claude Haiku 5.5', async () => {
    const s = await setup();
    const { factory, requests } = scripted(message([text('Hi')], 'end_turn'));
    await runTurn({ stream: factory, model: MODELS.find((m) => m.id === 'claude-haiku-5-5')!, tools: s.toolbox, messages: s.messages, events: s.events });
    expect(requests[0]).not.toHaveProperty('fallbacks');
    expect(requests[0]).not.toHaveProperty('betas');
  });

  it('does not run a tool whose input fails its schema, and tells Claude why', async () => {
    const s = await setup();
    const { factory } = scripted(message([tool('t1', 'replace_text', { find: 'Hello' })], 'tool_use'), message([text('Sorry.')], 'end_turn'));
    await runTurn({ stream: factory, model: MODELS[0]!, tools: s.toolbox, messages: s.messages, events: s.events });
    expect(texts(s.doc)[0]).toBe('Hello world');
    expect(s.seen[0]).toMatchObject({ name: 'replace_text', isError: true });
    const result = (s.messages[2]!.content as Array<{ content: string; is_error?: boolean }>)[0]!;
    expect(result.is_error).toBe(true);
    expect(JSON.parse(result.content)).toMatchObject({ INVALID_JSON: '{"find":"Hello"}' });
  });

  it('takes the whole turn back on a refusal', async () => {
    const s = await setup();
    const { factory } = scripted(
      message([tool('t1', 'replace_text', { find: 'Hello', replace: 'Halo' })], 'tool_use'),
      message([], 'refusal', { stop_details: { type: 'refusal', category: 'cyber' } }),
    );
    await runTurn({ stream: factory, model: MODELS[0]!, tools: s.toolbox, messages: s.messages, events: s.events });
    expect(s.messages).toEqual([]);
    expect(s.notices[0]).toContain('declined this request (cyber)');
    // Changes made before the refusal stay until the user undoes them.
    expect(texts(s.doc)[0]).toBe('Halo world');
  });

  it('never runs a tool call cut off by max_tokens', async () => {
    const s = await setup();
    const { factory } = scripted(message([tool('t1', 'replace_text', { find: 'Hello', replace: 'Hal' })], 'max_tokens'));
    await runTurn({ stream: factory, model: MODELS[0]!, tools: s.toolbox, messages: s.messages, events: s.events });
    expect(texts(s.doc)[0]).toBe('Hello world');
    expect(s.messages).toEqual([]);
    expect(s.notices[0]).toContain('cut off');
  });

  it('re-issues a step when the stream fails outside the API, and passes API errors on', async () => {
    const s = await setup();
    const { factory, requests } = scripted(new SyntaxError('Unexpected end of JSON input'), message([text('OK')], 'end_turn'));
    await runTurn({ stream: factory, model: MODELS[0]!, tools: s.toolbox, messages: s.messages, events: s.events });
    expect(requests).toHaveLength(2);
    expect(s.streamed()).toBe('OK');

    const t = await setup();
    const auth = new Anthropic.AuthenticationError(401, { type: 'error', error: { type: 'authentication_error', message: 'invalid x-api-key' } }, 'invalid x-api-key', new Headers());
    const failing = scripted(auth);
    await expect(runTurn({ stream: failing.factory, model: MODELS[0]!, tools: t.toolbox, messages: t.messages, events: t.events })).rejects.toBe(auth);
    expect(failing.requests).toHaveLength(1);
  });

  it('reports a closed document to Claude instead of failing', async () => {
    const s = await setup();
    s.close();
    const { factory } = scripted(message([tool('t1', 'read_document', {})], 'tool_use'), message([text('Please open one.')], 'end_turn'));
    await runTurn({ stream: factory, model: MODELS[0]!, tools: s.toolbox, messages: s.messages, events: s.events });
    expect(s.seen[0]).toMatchObject({ isError: true, content: 'No document is open. Ask the user to open a .docx first.' });
  });
});

describe('afterFallback', () => {
  it('echoes everything when there was no fallback', () => {
    const content = [{ type: 'thinking', thinking: '', signature: 's' }, text('a'), tool('t1', 'read_document', {})] as never;
    const { echo, toolUses } = afterFallback(content);
    expect(echo).toEqual(content);
    expect(toolUses.map((t) => t.id)).toEqual(['t1']);
  });

  it('drops thinking and tool calls from before a mid-output fallback, keeping the boundary', () => {
    const content = [
      { type: 'thinking', thinking: '', signature: 's1' },
      text('Partial answer'),
      tool('t1', 'replace_text', {}),
      { type: 'fallback', from: { model: 'claude-opus-5-5' }, to: { model: 'claude-opus-5' } },
      { type: 'thinking', thinking: '', signature: 's2' },
      tool('t2', 'read_document', {}),
    ] as never;
    const { echo, toolUses } = afterFallback(content);
    expect((echo as Block[]).map((b) => b.type)).toEqual(['text', 'fallback', 'thinking', 'tool_use']);
    expect(toolUses.map((t) => t.id)).toEqual(['t2']);
  });
});

describe('user instructions', () => {
  const run = async (s: Awaited<ReturnType<typeof setup>>, factory: StreamFactory, instructions?: string) =>
    runTurn({ stream: factory, model: MODELS[0]!, tools: s.toolbox, messages: s.messages, events: s.events, ...(instructions === undefined ? {} : { instructions }) });

  it('adds the user instructions after the built-in prompt', async () => {
    const s = await setup();
    const { factory, requests } = scripted(message([text('Done.')], 'end_turn'));
    await run(s, factory, 'You are a legal assistant.');
    expect(requests[0]!.system).toBe(`${SYSTEM_PROMPT}\n\n${instructionsSection('You are a legal assistant.')}`);
  });

  it('sends the built-in prompt alone without instructions', async () => {
    for (const instructions of [undefined, ' \n ']) {
      const s = await setup();
      const { factory, requests } = scripted(message([text('Done.')], 'end_turn'));
      await run(s, factory, instructions);
      expect(requests[0]!.system).toBe(SYSTEM_PROMPT);
    }
    expect(systemPrompt()).toBe(SYSTEM_PROMPT);
  });

  it('keeps the same instructions for every step of a turn', async () => {
    const s = await setup();
    const { factory, requests } = scripted(message([tool('t1', 'read_document', {})], 'tool_use'), message([text('Done.')], 'end_turn'));
    await run(s, factory, 'Act as a proofreader.');
    expect(requests).toHaveLength(2);
    expect(requests[1]!.system).toBe(requests[0]!.system);
    expect(requests[0]!.system).toContain('Act as a proofreader.');
  });
});
