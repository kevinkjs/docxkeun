#!/usr/bin/env node
/**
 * Serve the built docxkeun chat (dist/) with its Claude API relay, on this computer only.
 *
 *   pnpm build:chat && pnpm start:chat          # http://127.0.0.1:5260
 *   PORT=8080 ANTHROPIC_API_KEY=sk-ant-... node server/serve.mjs
 */
import { readFile, stat } from 'node:fs/promises';
import { createServer } from 'node:http';
import { extname, join, normalize, sep } from 'node:path';
import { fileURLToPath } from 'node:url';
import { createRelay } from './relay.mjs';

const dist = fileURLToPath(new URL('../dist/', import.meta.url));
const port = Number(process.env.PORT ?? 5260);
const relay = createRelay();

const TYPES = {
  '.html': 'text/html; charset=utf-8',
  '.js': 'text/javascript; charset=utf-8',
  '.css': 'text/css; charset=utf-8',
  '.json': 'application/json',
  '.svg': 'image/svg+xml',
  '.png': 'image/png',
  '.ttf': 'font/ttf',
  '.txt': 'text/plain; charset=utf-8',
};

/** @param {import('node:http').IncomingMessage} req @param {import('node:http').ServerResponse} res */
async function serveStatic(req, res) {
  if (req.method !== 'GET' && req.method !== 'HEAD') {
    res.writeHead(405).end();
    return;
  }
  const pathname = decodeURIComponent(new URL(req.url ?? '/', 'http://local').pathname);
  const file = normalize(join(dist, pathname.endsWith('/') ? pathname + 'index.html' : pathname));
  if (!file.startsWith(dist.endsWith(sep) ? dist : dist + sep)) {
    res.writeHead(404).end();
    return;
  }
  try {
    if (!(await stat(file)).isFile()) throw new Error('not a file');
    const body = await readFile(file);
    res.writeHead(200, {
      'content-type': TYPES[/** @type {keyof typeof TYPES} */ (extname(file))] ?? 'application/octet-stream',
      'cache-control': file.endsWith('index.html') ? 'no-store' : 'max-age=3600',
      'x-content-type-options': 'nosniff',
      'referrer-policy': 'no-referrer',
    });
    res.end(req.method === 'HEAD' ? undefined : body);
  } catch {
    res.writeHead(404, { 'content-type': 'text/plain; charset=utf-8' }).end('Not found. Run `pnpm build:chat` first.');
  }
}

const server = createServer((req, res) => {
  void relay(req, res, () => void serveStatic(req, res));
});

server.listen(port, '127.0.0.1', () => {
  process.stdout.write(`docxkeun chat: http://127.0.0.1:${port}${process.env.ANTHROPIC_API_KEY ? ' (using ANTHROPIC_API_KEY)' : ''}\n`);
});
