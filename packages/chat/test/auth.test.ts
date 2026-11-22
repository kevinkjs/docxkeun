/** The page's side of the relay: what its requests carry, and the messages the key dialog shows. */
import { describe, expect, it } from 'vitest';
import { giveKey, keyHint, keyProblem, NoRelayError, relayClient, relayStatus, START_THE_SERVER } from '../src/auth.js';

const KEY = 'sk-ant-api03-TESTVALUE-not-a-real-key';
const json = (status: number, body: unknown) => new Response(JSON.stringify(body), { status, headers: { 'content-type': 'application/json' } });

function recordingFetch(answer: () => Response | Promise<Response>) {
  const seen: Array<{ url: string; method: string; headers: Headers; body: string | undefined }> = [];
  const fetch = (async (url: RequestInfo | URL, init?: RequestInit) => {
    seen.push({ url: String(url), method: init?.method ?? 'GET', headers: new Headers(init?.headers), body: typeof init?.body === 'string' ? init.body : undefined });
    return answer();
  }) as typeof globalThis.fetch;
  return { fetch, seen };
}

describe('the page talking to the relay', () => {
  it('sends Claude requests to the relay without any credential', async () => {
    const f = recordingFetch(() => json(200, { type: 'model', id: 'claude-opus-5-5', display_name: 'x', created_at: '2026-01-01T00:00:00Z' }));
    await relayClient({ fetch: f.fetch, maxRetries: 0 }).models.retrieve('claude-opus-5-5');
    const [req] = f.seen;
    expect(req!.url).toBe('http://localhost/anthropic/v1/models/claude-opus-5-5');
    expect(req!.headers.get('x-docxkeun-relay')).toBe('1');
    expect(req!.headers.get('authorization')).toBeNull();
    expect(req!.headers.get('x-api-key')).toBeNull();
  });

  it('hands the key to the relay, which answers with the API verdict', async () => {
    const ok = recordingFetch(() => json(200, { hasKey: true, fromEnv: false }));
    await giveKey(KEY, 'claude-opus-5-5', ok.fetch);
    expect(ok.seen[0]).toMatchObject({ url: 'http://localhost/relay/key', method: 'POST', body: JSON.stringify({ key: KEY, model: 'claude-opus-5-5' }) });
    expect(ok.seen[0]!.headers.get('x-docxkeun-relay')).toBe('1');

    const bad = recordingFetch(() => json(401, { type: 'error', error: { type: 'authentication_error', message: 'invalid x-api-key' } }));
    const err = await giveKey(KEY, 'claude-opus-5-5', bad.fetch).catch((e: unknown) => e);
    const message = keyProblem(err, KEY);
    expect(message).toContain('The API did not accept this key (the API said: “invalid x-api-key”).');
    expect(message).not.toContain(KEY);
  });

  it('says how to start the server when there is no relay', async () => {
    const down = recordingFetch(() => Promise.reject(new TypeError('fetch failed')));
    await expect(relayStatus(down.fetch)).rejects.toBeInstanceOf(NoRelayError);
    const staticHost = recordingFetch(() => new Response('<!doctype html>', { status: 200, headers: { 'content-type': 'text/html' } }));
    const err = await relayStatus(staticHost.fetch).catch((e: unknown) => e);
    expect(keyProblem(err, KEY)).toBe(START_THE_SERVER);
  });

  it('recognises keys of the wrong kind', () => {
    expect(keyHint('sk-ant-oat01-abc')).toContain('Claude subscription token');
    expect(keyHint('sk-ant-admin01-abc')).toContain('Admin API key');
    expect(keyHint('abc')).toContain('start with "sk-ant-"');
    expect(keyHint('sk-ant-api03-ab cd')).toContain('spaces');
    expect(keyHint(KEY)).toBeUndefined();
  });
});
