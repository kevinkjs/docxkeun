/** The local relay, as a real HTTP server in front of a fake api.anthropic.com. */
import { createServer, request, type Server } from 'node:http';
import { afterEach, describe, expect, it } from 'vitest';
import { createRelay } from '../server/relay.mjs';

const GOOD = 'sk-ant-api03-GOOD-test-value';
type Seen = { url: string; method: string; headers: Record<string, string>; body: string };

/** A stand-in for the Claude API: accepts GOOD, streams a short reply for /v1/messages. */
function fakeApi() {
  const seen: Seen[] = [];
  const fetch = (async (url: RequestInfo | URL, init?: RequestInit) => {
    const headers = Object.fromEntries(new Headers(init?.headers).entries());
    const body = init?.body instanceof Uint8Array ? new TextDecoder().decode(init.body) : String(init?.body ?? '');
    seen.push({ url: String(url), method: init?.method ?? 'GET', headers, body });
    if (headers['x-api-key'] !== GOOD) {
      return new Response(JSON.stringify({ type: 'error', error: { type: 'authentication_error', message: 'invalid x-api-key' } }), { status: 401, headers: { 'content-type': 'application/json' } });
    }
    if (String(url).includes('/v1/messages')) {
      const chunks = ['event: message_start\ndata: {}\n\n', 'event: message_stop\ndata: {}\n\n'];
      const stream = new ReadableStream({
        start(c) {
          for (const ch of chunks) c.enqueue(new TextEncoder().encode(ch));
          c.close();
        },
      });
      return new Response(stream, { status: 200, headers: { 'content-type': 'text/event-stream', 'request-id': 'req_test', 'set-cookie': 'nope=1' } });
    }
    return new Response(JSON.stringify({ type: 'model', id: 'claude-opus-5-5' }), { status: 200, headers: { 'content-type': 'application/json' } });
  }) as typeof globalThis.fetch;
  return { fetch, seen };
}

let server: Server | undefined;
afterEach(() => new Promise<void>((ok) => (server ? server.close(() => ok()) : ok())));

async function start(apiKey = '') {
  const api = fakeApi();
  const relay = createRelay({ apiKey, upstream: 'https://api.test', fetch: api.fetch });
  server = createServer((req, res) => void relay(req, res, () => res.writeHead(200).end('static page')));
  await new Promise<void>((ok) => server!.listen(0, '127.0.0.1', ok));
  const port = (server!.address() as { port: number }).port;
  const base = `http://127.0.0.1:${port}`;
  const call = (path: string, init: RequestInit = {}, own = true) =>
    fetch(base + path, { ...init, headers: { ...(own ? { 'x-docxkeun-relay': '1' } : {}), ...(init.headers as Record<string, string>) } });
  return { api, base, port, call };
}

/** A GET with a chosen Host header (fetch always sends the real one). */
function rawStatus(port: number, host: string): Promise<number> {
  return new Promise((ok, fail) => {
    const req = request({ host: '127.0.0.1', port, path: '/relay/status', headers: { host, 'x-docxkeun-relay': '1' } }, (res) => {
      res.resume();
      ok(res.statusCode ?? 0);
    });
    req.on('error', fail);
    req.end();
  });
}

describe('chat relay', () => {
  it('takes a key only after the API accepts it', async () => {
    const { call, api } = await start();
    expect(await (await call('/relay/status')).json()).toEqual({ hasKey: false, fromEnv: false });
    const noKey = await call('/anthropic/v1/models/claude-opus-5-5');
    expect(noKey.status).toBe(401);
    expect((await noKey.json()).error.type).toBe('authentication_error');

    const bad = await call('/relay/key', { method: 'POST', headers: { 'content-type': 'application/json' }, body: JSON.stringify({ key: 'sk-ant-wrong', model: 'claude-opus-5-5' }) });
    expect(bad.status).toBe(401);
    expect((await bad.json()).error.message).toBe('invalid x-api-key');
    expect(await (await call('/relay/status')).json()).toEqual({ hasKey: false, fromEnv: false });

    const good = await call('/relay/key', { method: 'POST', headers: { 'content-type': 'application/json' }, body: JSON.stringify({ key: GOOD, model: 'claude-opus-5-5' }) });
    expect(good.status).toBe(200);
    expect(api.seen.at(-1)).toMatchObject({ url: 'https://api.test/v1/models/claude-opus-5-5', method: 'GET' });
    expect(await (await call('/relay/status')).json()).toEqual({ hasKey: true, fromEnv: false });

    expect((await call('/relay/key', { method: 'DELETE' })).status).toBe(204);
    expect(await (await call('/relay/status')).json()).toEqual({ hasKey: false, fromEnv: false });
  });

  it('forwards Claude calls with the key and nothing from the browser, and streams the answer back', async () => {
    const { call, api } = await start(GOOD);
    const body = JSON.stringify({ model: 'claude-opus-5-5', max_tokens: 10, messages: [], stream: true });
    const res = await call('/anthropic/v1/messages?beta=true', {
      method: 'POST',
      body,
      headers: {
        'content-type': 'application/json',
        'anthropic-version': '2026-10-21',
        'anthropic-beta': 'server-side-fallback-2026-07-01',
        authorization: 'Bearer held-by-the-local-relay',
        'anthropic-dangerous-direct-browser-access': 'true',
        cookie: 'session=abc',
      },
    });
    expect(res.status).toBe(200);
    expect(res.headers.get('content-type')).toBe('text/event-stream');
    expect(res.headers.get('request-id')).toBe('req_test');
    expect(res.headers.get('set-cookie')).toBeNull();
    expect(await res.text()).toBe('event: message_start\ndata: {}\n\nevent: message_stop\ndata: {}\n\n');

    const sent = api.seen.at(-1)!;
    expect(sent.url).toBe('https://api.test/v1/messages?beta=true');
    expect(sent.body).toBe(body);
    expect(sent.headers).toEqual({
      accept: '*/*',
      'content-type': 'application/json',
      'anthropic-version': '2026-10-21',
      'anthropic-beta': 'server-side-fallback-2026-07-01',
      'x-api-key': GOOD,
    });
  });

  it('reports a key from the environment', async () => {
    const { call } = await start(GOOD);
    expect(await (await call('/relay/status')).json()).toEqual({ hasKey: true, fromEnv: true });
  });

  it('refuses other websites, other host names and other API calls', async () => {
    const { call, port } = await start(GOOD);
    expect((await call('/relay/status', {}, false)).status).toBe(403);
    expect((await call('/anthropic/v1/messages', { method: 'POST', body: '{}', headers: { origin: 'https://evil.example', 'content-type': 'text/plain' } })).status).toBe(403);
    expect(await rawStatus(port, 'evil.example')).toBe(403);
    expect(await rawStatus(port, `localhost:${port}`)).toBe(200);
    expect((await call('/anthropic/v1/organizations/me')).status).toBe(404);
    expect((await call('/anthropic/v1/messages', { method: 'DELETE' })).status).toBe(405);
    // Everything else is the page itself.
    expect(await (await call('/index.html', {}, false)).text()).toBe('static page');
  });
});
