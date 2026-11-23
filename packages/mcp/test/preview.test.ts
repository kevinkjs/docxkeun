/** The live preview server: pages, files, safety checks and change events. */
import { execFileSync } from 'node:child_process';
import { existsSync } from 'node:fs';
import { mkdtemp, readFile, rm, writeFile } from 'node:fs/promises';
import { request } from 'node:http';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { Client } from '@modelcontextprotocol/client';
import { InMemoryTransport } from '@modelcontextprotocol/server';
import { afterEach, beforeEach, describe, expect, it } from 'vitest';
import { encryptPackage } from 'docxkeun/security';
import { buildDocx, p } from '../../../test/helpers/docx.js';
import { textBoxRun } from '../../../test/helpers/drawings.js';
import { createServer, startPreview, type Preview } from '../src/index.js';
import { PREVIEW_SCRIPT } from '../src/preview-page.js';
import { parseColors } from '../src/preview.js';

let dir: string;
let preview: Preview | undefined;
const cleanup: Array<() => Promise<void> | void> = [];

beforeEach(async () => {
  dir = await mkdtemp(join(tmpdir(), 'docxkeun-preview-'));
  await writeFile(join(dir, 'letter.docx'), await buildDocx({ body: p('Hello preview') }));
});

afterEach(async () => {
  for (const fn of cleanup.splice(0).reverse()) await fn();
  await preview?.close();
  preview = undefined;
  await rm(dir, { recursive: true, force: true });
});

async function start(options: Parameters<typeof startPreview>[1] = { port: 0 }): Promise<Preview> {
  preview = await startPreview([dir], options);
  return preview;
}

/** A raw GET with a chosen Host header (fetch always sends the real one). */
function rawGet(url: string, host: string): Promise<number> {
  return new Promise((ok, fail) => {
    const req = request(url, { headers: { host } }, (res) => {
      res.resume();
      ok(res.statusCode ?? 0);
    });
    req.on('error', fail);
    req.end();
  });
}

/** Collect server-sent events until `until` returns true. */
async function events(url: string, trigger: () => Promise<void>, until: (events: Array<Record<string, string>>) => boolean): Promise<Array<Record<string, string>>> {
  const controller = new AbortController();
  const res = await fetch(`${url}/api/events`, { signal: controller.signal });
  const reader = res.body!.getReader();
  const seen: Array<Record<string, string>> = [];
  let buffer = '';
  await trigger();
  const deadline = Date.now() + 10_000;
  try {
    while (!until(seen)) {
      if (Date.now() > deadline) throw new Error(`timed out; saw ${JSON.stringify(seen)}`);
      const { value, done } = await reader.read();
      if (done) break;
      buffer += new TextDecoder().decode(value);
      let end: number;
      while ((end = buffer.indexOf('\n\n')) >= 0) {
        const block = buffer.slice(0, end);
        buffer = buffer.slice(end + 2);
        const data = block.split('\n').find((l) => l.startsWith('data: '));
        if (data) seen.push(JSON.parse(data.slice(6)));
      }
    }
  } finally {
    controller.abort();
  }
  return seen;
}

describe('live preview server', () => {
  it('serves the page, the document list and documents', async () => {
    const { url } = await start();
    const page = await fetch(url);
    expect(page.headers.get('content-security-policy')).toContain("script-src 'self'");
    expect(await page.text()).toContain('<script type="module" src="./preview.js">');
    expect(await (await fetch(`${url}/preview.js`)).text()).toContain("from './lib/render/index.js'");
    expect(await (await fetch(`${url}/api/documents`)).json()).toMatchObject([{ path: 'letter.docx' }]);
    const file = await fetch(`${url}/api/file?path=letter.docx`);
    expect(new Uint8Array(await file.arrayBuffer())).toEqual(new Uint8Array(await readFile(join(dir, 'letter.docx'))));
    const font = await fetch(`${url}/fonts/Carlito-Regular.ttf`);
    expect(font.headers.get('content-type')).toBe('font/ttf');
  });

  it('serves a page script that parses and refits when the viewer gets narrower or wider', async () => {
    const file = join(dir, 'preview.mjs');
    await writeFile(file, PREVIEW_SCRIPT);
    expect(() => execFileSync(process.execPath, ['--check', file], { stdio: 'pipe' })).not.toThrow();
    // The navigation panel opening or folding changes the viewer's width without a window resize.
    expect(PREVIEW_SCRIPT).toContain('new ResizeObserver(');
    expect(PREVIEW_SCRIPT).toContain('.observe(viewer.root)');
    expect(PREVIEW_SCRIPT).toMatch(/setTimeout\(applyFit, 50\)/);
  });

  it("gives the page the AI's name and colours, escaped", async () => {
    const { url } = await start({ port: 0, watch: false, ai: { name: 'Ada "the" <AI> $& co', color: ['#0b57d0', 'rgb(52, 168, 83)'] } });
    const html = await (await fetch(url)).text();
    const content = /<meta name="docxkeun-ai" content="([^"]*)">/.exec(html)?.[1];
    expect(content).toBeDefined();
    expect(content).not.toMatch(/[<>"]/);
    const decoded = content!.replace(/&#(\d+);/g, (_, n: string) => String.fromCharCode(Number(n)));
    expect(JSON.parse(decoded)).toEqual({ name: 'Ada "the" <AI> $& co', color: ['#0b57d0', 'rgb(52, 168, 83)'] });
    expect(PREVIEW_SCRIPT).toContain('viewer.activity.name = ai.name');
    expect(PREVIEW_SCRIPT).toContain('viewer.activity.color = ai.color');
  });

  it('reads --ai-color as one colour or two, keeping commas inside rgb()', () => {
    expect(parseColors('#0b57d0')).toBe('#0b57d0');
    expect(parseColors(' #0b57d0 , teal ')).toEqual(['#0b57d0', 'teal']);
    expect(parseColors('rgb(11, 87, 208),hsl(140 50% 40%)')).toEqual(['rgb(11, 87, 208)', 'hsl(140 50% 40%)']);
    for (const bad of ['', '#a,#b,#c', 'red;background:url(x)', '"red"', '#a,']) expect(() => parseColors(bad), bad).toThrow(/Not a colour/);
  });

  it.skipIf(!existsSync(new URL('../../../dist/render/index.js', import.meta.url)))('serves the docxkeun browser build', async () => {
    const { url } = await start();
    const lib = await fetch(`${url}/lib/render/index.js`);
    expect(lib.status).toBe(200);
    expect(lib.headers.get('content-type')).toContain('text/javascript');
  });

  it('refuses other hosts, other folders and other files', async () => {
    const { url } = await start();
    expect(await rawGet(url, 'evil.example')).toBe(403);
    expect(await rawGet(url, `localhost:${new URL(url).port}`)).toBe(200);
    const outside = await mkdtemp(join(tmpdir(), 'docxkeun-outside-'));
    cleanup.push(() => rm(outside, { recursive: true, force: true }));
    await writeFile(join(outside, 'secret.docx'), await buildDocx({ body: p('secret') }));
    for (const path of [join(outside, 'secret.docx'), '../secret.docx', 'notes.txt']) {
      expect((await fetch(`${url}/api/file?path=${encodeURIComponent(path)}`)).status, path).toBe(404);
    }
    for (const path of ['/lib/../package.json', '/lib/%2e%2e/%2e%2e/package.json', '/fonts/../index.js']) {
      expect((await fetch(url + path)).status, path).toBe(404);
    }
    expect((await fetch(url, { method: 'POST' })).status).toBe(405);
  });

  it('explains that encrypted documents cannot be previewed', async () => {
    await writeFile(join(dir, 'locked.docx'), encryptPackage(await buildDocx({ body: p('x') }), 'pw', { spinCount: 1000 }));
    const { url } = await start();
    const res = await fetch(`${url}/api/file?path=locked.docx`);
    expect(res.status).toBe(423);
    expect(((await res.json()) as { error: string }).error).toContain('encrypted');
  });

  it('uses the next port when the first is taken', async () => {
    const first = await start({ port: 0 });
    const port = Number(new URL(first.url).port);
    const second = await startPreview([dir], { port, watch: false });
    cleanup.push(() => second.close());
    // The next free port: usually port + 1, unless something else on the machine has that one.
    const next = Number(new URL(second.url).port);
    expect(next).toBeGreaterThan(port);
    expect(next).toBeLessThanOrEqual(port + 20);
  });

  it('announces changes made through the MCP tools, with what changed', async () => {
    const pv = await start({ port: 0, watch: false });
    const server = await createServer({ folders: [dir], preview: pv });
    const [clientSide, serverSide] = InMemoryTransport.createLinkedPair();
    await server.connect(serverSide);
    const client = new Client({ name: 'test', version: '1.0.0' });
    await client.connect(clientSide);
    cleanup.push(() => client.close());

    expect(client.getInstructions()).toContain(pv.url);
    const link = await client.callTool({ name: 'get_preview_link', arguments: { path: 'letter.docx' } });
    expect((link.content as Array<{ text: string }>)[0]!.text).toBe(`Live preview of letter.docx: ${pv.url}/?doc=letter.docx`);

    let commentId = '';
    const seen = await events(
      pv.url,
      async () => {
        await client.callTool({ name: 'replace_text', arguments: { path: 'letter.docx', find: 'Hello', replace: 'Halo' } });
        const r = await client.callTool({ name: 'add_comment', arguments: { path: 'letter.docx', text: 'Halo', comment: 'Friendly!' } });
        commentId = /comment (\S+) on/.exec((r.content as Array<{ text: string }>)[0]!.text)![1]!;
      },
      (s) => s.filter((e) => e.type === 'changed').length >= 2,
    );
    const changes = seen.filter((e) => e.type === 'changed');
    expect(changes[0]).toEqual({ type: 'changed', path: 'letter.docx', message: 'Replaced 1 of 1 match.' });
    expect(changes[1]).toEqual({ type: 'changed', path: 'letter.docx', message: `Added comment ${commentId} on "Halo" in paragraph 1.`, comment: commentId });
  });

  it('has no preview link tool without a preview', async () => {
    const server = await createServer({ folders: [dir] });
    const [clientSide, serverSide] = InMemoryTransport.createLinkedPair();
    await server.connect(serverSide);
    const client = new Client({ name: 'test', version: '1.0.0' });
    await client.connect(clientSide);
    cleanup.push(() => client.close());
    expect((await client.listTools()).tools.map((t) => t.name)).not.toContain('get_preview_link');
  });

  it('announces changes made outside the server', async () => {
    const pv = await start({ port: 0 });
    const bytes = await buildDocx({ body: p('Edited in Word') });
    const seen = await events(
      pv.url,
      async () => {
        await new Promise((r) => setTimeout(r, 100));
        await writeFile(join(dir, 'letter.docx'), bytes);
      },
      (s) => s.some((e) => e.type === 'changed'),
    );
    expect(seen.find((e) => e.type === 'changed')).toEqual({ type: 'changed', path: 'letter.docx' });
  });

  describe('selection and activity', () => {
    async function connectTo(pv: Preview) {
      const server = await createServer({ folders: [dir], preview: pv });
      const [clientSide, serverSide] = InMemoryTransport.createLinkedPair();
      await server.connect(serverSide);
      const client = new Client({ name: 'test', version: '1.0.0' });
      await client.connect(clientSide);
      cleanup.push(() => client.close());
      return client;
    }
    const tokenOf = async (url: string) => /name="docxkeun-token" content="([0-9a-f]+)"/.exec(await (await fetch(`${url}/`)).text())![1]!;
    const post = (url: string, body: unknown, headers: Record<string, string>) =>
      fetch(`${url}/api/selection`, { method: 'POST', headers: { 'content-type': 'application/json', ...headers }, body: JSON.stringify(body) });

    it('tells the AI what the user selected in the preview, read from the file', async () => {
      await writeFile(join(dir, 'letter.docx'), await buildDocx({ body: p('Dear team,') + p('Payment is due within 30 days of the invoice date.') }));
      const pv = await start({ port: 0, watch: false });
      const client = await connectTo(pv);
      expect(client.getInstructions()).toContain('get_selection');
      const text = async () => ((await client.callTool({ name: 'get_selection', arguments: {} })).content as Array<{ text: string }>)[0]!.text;
      expect(await text()).toContain('Nothing is selected');
      const token = await tokenOf(pv.url);
      const r = await post(pv.url, { path: 'letter.docx', story: 'body', start: { paragraph: 2, offset: 22 }, end: { paragraph: 2, offset: 29 } }, { 'x-docxkeun-token': token, origin: pv.url });
      expect(r.status).toBe(204);
      const said = await text();
      expect(said).toContain('Selected in letter.docx: paragraph 2, page 1, 7 characters.');
      expect(said).toContain('[2] Payment is due within «30 days» of the invoice date.');
      // Also a resource, for @-mentions.
      const res = await client.readResource({ uri: 'docxkeun://selection' });
      expect((res.contents[0] as { text: string }).text).toContain('«30 days»');
      // Clearing it.
      await post(pv.url, { path: 'letter.docx', clear: true }, { 'x-docxkeun-token': token });
      expect(await text()).toContain('Nothing is selected');
    });

    it('only takes selections from its own pages', async () => {
      const pv = await start({ port: 0, watch: false });
      const token = await tokenOf(pv.url);
      const body = { path: 'letter.docx', start: { paragraph: 1, offset: 0 }, end: { paragraph: 1, offset: 5 } };
      expect((await post(pv.url, body, {})).status).toBe(403);
      expect((await post(pv.url, body, { 'x-docxkeun-token': 'f'.repeat(32) })).status).toBe(403);
      expect((await post(pv.url, body, { 'x-docxkeun-token': token, origin: 'https://evil.example' })).status).toBe(403);
      expect((await post(pv.url, { ...body, path: '../../etc/passwd.docx' }, { 'x-docxkeun-token': token })).status).toBe(404);
      expect((await post(pv.url, { ...body, start: { paragraph: 0, offset: -1 } }, { 'x-docxkeun-token': token })).status).toBe(400);
      expect(pv.selection()).toBeUndefined();
    });

    it('tells pages which paragraphs a tool is about to work on', async () => {
      await writeFile(join(dir, 'letter.docx'), await buildDocx({ body: p('One') + p('Hello preview') + p('Three') }));
      const pv = await start({ port: 0, watch: false });
      const client = await connectTo(pv);
      const seen = await events(
        pv.url,
        async () => {
          await client.callTool({ name: 'replace_text', arguments: { path: 'letter.docx', find: 'Hello', replace: 'Halo' } });
          await client.callTool({ name: 'read_document', arguments: { path: 'letter.docx' } });
        },
        (s) => s.filter((e) => e.type === 'activity').length >= 2,
      );
      const activity = seen.filter((e) => e.type === 'activity');
      expect(activity[0]).toEqual({ type: 'activity', path: 'letter.docx', paragraphs: [2], mode: 'edit', label: 'editing' });
      expect(activity[1]).toEqual({ type: 'activity', path: 'letter.docx', paragraphs: [], mode: 'read', label: 'reading' });
    });

    // read_document numbers a text box's paragraphs after the paragraph holding it; the viewer does not.
    const WITH_BOX = p('One') + `<w:p>${textBoxRun('Boxed')}</w:p>` + p('Three');

    it("gives pages the viewer's paragraph numbers when the document has a text box", async () => {
      await writeFile(join(dir, 'letter.docx'), await buildDocx({ body: WITH_BOX }));
      const pv = await start({ port: 0, watch: false });
      const client = await connectTo(pv);
      const seen = await events(
        pv.url,
        async () => {
          await client.callTool({ name: 'edit_paragraph', arguments: { path: 'letter.docx', paragraph: 3, text: 'In the box' } });
          await client.callTool({ name: 'replace_text', arguments: { path: 'letter.docx', find: 'Three', replace: '3' } });
        },
        (s) => s.filter((e) => e.type === 'activity').length >= 2,
      );
      // Paragraph 3 is in the box the viewer's paragraph 2 holds; paragraph 4 is the viewer's 3.
      expect(seen.filter((e) => e.type === 'activity').map((e) => e.paragraphs)).toEqual([[2], [3]]);
    });

    it("gives the page read_document's number for each paragraph it counts, for its selection chip", async () => {
      await writeFile(join(dir, 'letter.docx'), await buildDocx({ body: WITH_BOX }));
      const pv = await start({ port: 0, watch: false });
      // The page's paragraphs One, the one holding the box and Three are read_document's 1, 2 and 4.
      expect(await (await fetch(`${pv.url}/api/numbering?path=letter.docx`)).json()).toEqual([1, 2, 4]);
      expect((await fetch(`${pv.url}/api/numbering?path=missing.docx`)).status).toBe(404);
      expect(PREVIEW_SCRIPT).toContain('./api/numbering');
    });

    it("tells the AI read_document's paragraph numbers for a selection after a text box", async () => {
      await writeFile(join(dir, 'letter.docx'), await buildDocx({ body: WITH_BOX }));
      const pv = await start({ port: 0, watch: false });
      const client = await connectTo(pv);
      const token = await tokenOf(pv.url);
      // The page counts paragraphs as the viewer does: "Three" is its paragraph 3.
      await post(pv.url, { path: 'letter.docx', story: 'body', start: { paragraph: 1, offset: 0 }, end: { paragraph: 3, offset: 5 } }, { 'x-docxkeun-token': token, origin: pv.url });
      const said = ((await client.callTool({ name: 'get_selection', arguments: {} })).content as Array<{ text: string }>)[0]!.text;
      expect(said).toContain('Selected in letter.docx: paragraphs 1–4');
      expect(said).toContain('[4] «Three»');
      expect(said).not.toContain('[3]');
    });
  });
});
