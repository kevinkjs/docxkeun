/**
 * Opt-in check against the real Claude API, with the exact request the chat sends (streaming, tools
 * with eager input streaming, prompt caching, effort, refusal fallbacks). Costs a few cents; runs only
 * with DOCXKEUN_LIVE=1 and an ANTHROPIC_API_KEY:
 *
 *   DOCXKEUN_LIVE=1 ANTHROPIC_API_KEY=sk-ant-... pnpm vitest run packages/chat/test/live.test.ts
 */
import Anthropic from '@anthropic-ai/sdk';
import type { BetaMessageParam } from '@anthropic-ai/sdk/resources/beta/messages/messages';
import { describe, expect, it } from 'vitest';
import { DocxDocument } from 'docxkeun';
import { buildDocx, p } from '../../../test/helpers/docx.js';
import { anthropicStream, MODELS, runTurn } from '../src/agent.js';
import { createToolbox } from '../src/tools.js';

const live = process.env.DOCXKEUN_LIVE === '1' && !!process.env.ANTHROPIC_API_KEY;

describe.skipIf(!live)('live Claude API', () => {
  it('edits a document through the tools', { timeout: 180_000 }, async () => {
    const doc = await DocxDocument.open(await buildDocx({ body: p('Payment is due within 30 days.') + p('Supplier: ACME Ltd.') }));
    const toolbox = createToolbox(() => [{ doc, name: 'live.docx' }]);
    const messages: BetaMessageParam[] = [{ role: 'user', content: '[Open document: live.docx]\n\nReplace "ACME Ltd." with "Contoso GmbH", then tell me in one sentence what you did.' }];
    const tools: string[] = [];
    let text = '';
    const notices: string[] = [];
    await runTurn({
      stream: anthropicStream(new Anthropic()),
      model: MODELS[0]!,
      tools: toolbox,
      messages,
      events: { text: (d) => (text += d), tool: (name, _i, o) => tools.push(`${name}${o.isError ? ' (error)' : ''}`), notice: (n) => notices.push(n) },
    });
    expect(notices).toEqual([]);
    expect(tools).toContain('replace_text');
    expect((await toolbox.run('find_text', { query: 'Contoso GmbH' })).content).toContain('1 match');
    expect(text.length).toBeGreaterThan(0);
  });
});
