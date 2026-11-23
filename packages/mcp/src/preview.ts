/**
 * Live preview: a small web server on 127.0.0.1 that shows the documents with docxkeun's own viewer
 * and redraws a document whenever a tool (or anything else) changes it. Pages get changes through
 * server-sent events, together with what the AI did ("Added comment 3 on …").
 */
import { randomBytes } from 'node:crypto';
import { watch, type FSWatcher } from 'node:fs';
import { readFile, stat } from 'node:fs/promises';
import { createServer, type IncomingMessage, type Server, type ServerResponse } from 'node:http';
import { createRequire } from 'node:module';
import { basename, dirname, extname, join, resolve, sep } from 'node:path';
import { DocxDocument } from 'docxkeun';
import { isEncryptedPackage } from 'docxkeun/security';
import { fontUrls } from 'docxkeun-fonts';
import { listParagraphs } from './outline.js';
import { PREVIEW_HTML, PREVIEW_SCRIPT } from './preview-page.js';
import { DOCUMENT_EXTENSIONS, Workspace } from './workspace.js';

export interface PreviewEvent {
  /** The document's real path. */
  path: string;
  /** What changed, in words (shown to the person watching). */
  message?: string;
  /** A comment to scroll to. */
  comment?: string;
}

/** The AI is about to work on a document: pages shimmer those paragraphs (or scan, for reads). */
export interface ActivityEvent {
  path: string;
  /** Paragraph numbers as the viewer counts them: body and table cells, not text boxes (viewerParagraphs); none: the whole document. */
  paragraphs: number[];
  mode: 'read' | 'edit';
  /** A verb for the caret label: "editing", "reading", "commenting"… */
  label?: string;
}

/** What the person watching has selected in the preview, as positions (the text is read from the file). */
export interface PreviewSelection {
  /** The document's real path. */
  path: string;
  story: 'body' | 'header' | 'footer' | 'footnote' | 'endnote' | 'textbox';
  part?: string;
  start: { paragraph: number; offset: number };
  end: { paragraph: number; offset: number };
  /** When it was selected (ms since the epoch). */
  at: number;
}

export interface Preview {
  /** Base URL, e.g. http://127.0.0.1:5230 */
  url: string;
  /** Link that opens one document. */
  link(realPath: string): string;
  /** Tell open pages that a document changed. */
  notify(event: PreviewEvent): void;
  /** Tell open pages what the AI is about to work on. */
  activity(event: ActivityEvent): void;
  /** The latest selection made in a preview page, if any. */
  selection(): PreviewSelection | undefined;
  close(): Promise<void>;
}

export interface PreviewOptions {
  /** First port to try (default 5230); the next free one is used when it is taken. 0 picks any free port. */
  port?: number;
  /** Watch the folders for changes made outside the server (default true). */
  watch?: boolean;
  /** How the AI shows on the pages: the name on its caret and its colour (one CSS colour, or two for a gradient). */
  ai?: { name?: string; color?: string | [string, string] };
}

/**
 * `--ai-color`: one CSS colour, or two separated by a comma (commas inside `rgb(…)` don't count).
 * The browser checks the colour itself; this only refuses what could never be one.
 */
export function parseColors(value: string): string | [string, string] {
  const parts: string[] = [];
  let depth = 0;
  let current = '';
  for (const c of value) {
    if (c === '(') depth++;
    else if (c === ')') depth = Math.max(0, depth - 1);
    if (c === ',' && depth === 0) {
      parts.push(current.trim());
      current = '';
    } else current += c;
  }
  parts.push(current.trim());
  if (parts.length > 2 || parts.some((p) => !p || !/^[#a-z0-9().,%\s/+-]+$/i.test(p))) {
    throw new Error(`Not a colour: "${value}" (use a CSS colour, or two separated by a comma)`);
  }
  return parts.length === 1 ? parts[0]! : [parts[0]!, parts[1]!];
}

/** Text safe inside a double-quoted HTML attribute. */
function escapeAttribute(text: string): string {
  return text.replace(/[&<>"']/g, (c) => `&#${c.charCodeAt(0)};`);
}

const DEFAULT_PORT = 5230;
const PORT_ATTEMPTS = 20;

/** docxkeun's browser build: the dist folder next to its package.json. */
function docxkeunDist(): string {
  const pkg = createRequire(import.meta.url).resolve('docxkeun/package.json');
  return join(dirname(pkg), 'dist');
}

function listen(server: Server, port: number): Promise<number> {
  return new Promise((ok, fail) => {
    const onError = (e: Error) => {
      server.off('listening', onListening);
      fail(e);
    };
    const onListening = () => {
      server.off('error', onError);
      const address = server.address();
      ok(typeof address === 'object' && address ? address.port : port);
    };
    server.once('error', onError);
    server.once('listening', onListening);
    server.listen(port, '127.0.0.1');
  });
}

export async function startPreview(folders: string[], options: PreviewOptions = {}): Promise<Preview> {
  const ws = await Workspace.create(folders);
  const dist = docxkeunDist();
  const clients = new Set<ServerResponse>();
  // Pages prove they were served by us (not another website posting to 127.0.0.1) with this token.
  const token = randomBytes(16).toString('hex');
  let selection: PreviewSelection | undefined;
  /** Paths the server just wrote itself, so the file watcher does not announce them twice. */
  const recent = new Map<string, number>();
  let port = 0;

  const send = (data: unknown): void => {
    const chunk = `data: ${JSON.stringify(data)}\n\n`;
    for (const res of clients) res.write(chunk);
  };

  const json = (res: ServerResponse, status: number, body: unknown): void => {
    res.writeHead(status, { 'content-type': 'application/json; charset=utf-8', 'cache-control': 'no-store' });
    res.end(JSON.stringify(body));
  };

  const handle = async (req: IncomingMessage, res: ServerResponse): Promise<void> => {
    // Only this machine's own pages: a different Host header means DNS rebinding or a proxy.
    const host = req.headers.host ?? '';
    if (![`127.0.0.1:${port}`, `localhost:${port}`].includes(host)) {
      res.writeHead(403).end('Forbidden');
      return;
    }
    const url = new URL(req.url ?? '/', `http://${host}`);
    const path = url.pathname;
    if (path === '/api/selection' && req.method === 'POST') {
      await receiveSelection(req, res);
      return;
    }
    if (req.method !== 'GET' && req.method !== 'HEAD') {
      res.writeHead(405, { allow: 'GET, HEAD' }).end();
      return;
    }

    if (path === '/') {
      res.writeHead(200, { 'content-type': 'text/html; charset=utf-8', 'cache-control': 'no-store', 'content-security-policy': "default-src 'self'; img-src 'self' blob: data:; style-src 'self' 'unsafe-inline'; font-src 'self' blob: data:; connect-src 'self'; script-src 'self'" });
      // Function replacements: `$&` and the like in a name stay literal.
      res.end(PREVIEW_HTML.replace('%TOKEN%', () => token).replace('%AI%', () => escapeAttribute(JSON.stringify(options.ai ?? {}))));
      return;
    }
    if (path === '/preview.js') {
      res.writeHead(200, { 'content-type': 'text/javascript; charset=utf-8', 'cache-control': 'no-store' });
      res.end(PREVIEW_SCRIPT);
      return;
    }
    if (path === '/api/documents') {
      const { entries } = await ws.listDocuments({ limit: 500 });
      json(res, 200, entries.map((e) => ({ path: ws.display(e.path), modified: e.modified.toISOString() })));
      return;
    }
    if (path === '/api/file') {
      let real: string;
      try {
        real = await ws.resolve(url.searchParams.get('path') ?? '', 'document', true);
      } catch (e) {
        json(res, 404, { error: (e as Error).message });
        return;
      }
      const bytes = await ws.read(real);
      if (isEncryptedPackage(bytes)) {
        json(res, 423, { error: 'This document is encrypted (it needs a password to open), so it cannot be previewed.' });
        return;
      }
      res.writeHead(200, { 'content-type': 'application/octet-stream', 'cache-control': 'no-store', 'content-length': bytes.length });
      res.end(bytes);
      return;
    }
    if (path === '/api/numbering') {
      // For each paragraph the page counts (body and table cells), read_document's number: the page's selection chip says what the AI sees.
      let real: string;
      try {
        real = await ws.resolve(url.searchParams.get('path') ?? '', 'document', true);
      } catch (e) {
        json(res, 404, { error: (e as Error).message });
        return;
      }
      const bytes = await ws.read(real);
      if (isEncryptedPackage(bytes)) {
        json(res, 200, []);
        return;
      }
      const byView: number[] = [];
      for (const e of listParagraphs(await DocxDocument.open(bytes))) if (!e.box) byView[e.view - 1] = e.n;
      json(res, 200, byView);
      return;
    }
    if (path === '/api/events') {
      res.writeHead(200, { 'content-type': 'text/event-stream; charset=utf-8', 'cache-control': 'no-store', connection: 'keep-alive' });
      res.write('retry: 2000\n\n');
      clients.add(res);
      req.on('close', () => clients.delete(res));
      return;
    }
    if (path.startsWith('/lib/')) {
      const file = resolve(dist, '.' + path.slice(4));
      if (!file.startsWith(dist + sep) || extname(file) !== '.js') {
        res.writeHead(404).end();
        return;
      }
      try {
        const body = await readFile(file);
        res.writeHead(200, { 'content-type': 'text/javascript; charset=utf-8', 'cache-control': 'no-cache' });
        res.end(body);
      } catch {
        res.writeHead(404).end();
      }
      return;
    }
    if (path.startsWith('/fonts/')) {
      const fontUrl = fontUrls[basename(path)];
      if (!fontUrl) {
        res.writeHead(404).end();
        return;
      }
      res.writeHead(200, { 'content-type': 'font/ttf', 'cache-control': 'max-age=86400' });
      res.end(await readFile(fontUrl));
      return;
    }
    res.writeHead(404).end('Not found');
  };

  /**
   * A page reports what is selected. Only positions are accepted; get_selection reads the text from
   * the file itself, so a page cannot put words in the AI's mouth.
   */
  const receiveSelection = async (req: IncomingMessage, res: ServerResponse): Promise<void> => {
    const origin = req.headers.origin;
    const own = [`http://127.0.0.1:${port}`, `http://localhost:${port}`];
    if ((origin && !own.includes(origin)) || req.headers['x-docxkeun-token'] !== token || !String(req.headers['content-type'] ?? '').startsWith('application/json')) {
      res.writeHead(403).end('Forbidden');
      return;
    }
    let body = '';
    for await (const chunk of req) {
      body += chunk;
      if (body.length > 4096) {
        res.writeHead(413).end();
        return;
      }
    }
    let input: Record<string, unknown>;
    try {
      input = JSON.parse(body) as Record<string, unknown>;
    } catch {
      res.writeHead(400).end();
      return;
    }
    const spot = (v: unknown) => {
      const o = v as { paragraph?: unknown; offset?: unknown } | undefined;
      const ok = (n: unknown) => typeof n === 'number' && Number.isInteger(n) && n >= 0 && n < 1e7;
      return o && ok(o.paragraph) && ok(o.offset) && (o.paragraph as number) >= 1 ? { paragraph: o.paragraph as number, offset: o.offset as number } : undefined;
    };
    let real: string;
    try {
      real = await ws.resolve(String(input.path ?? ''), 'document', true);
    } catch {
      res.writeHead(404).end();
      return;
    }
    if (input.clear === true) {
      if (selection?.path === real) selection = undefined;
      res.writeHead(204).end();
      return;
    }
    const stories = ['body', 'header', 'footer', 'footnote', 'endnote', 'textbox'] as const;
    const story = stories.find((st) => st === input.story) ?? 'body';
    const start = spot(input.start);
    const end = spot(input.end);
    if (!start || !end) {
      res.writeHead(400).end();
      return;
    }
    const part = typeof input.part === 'string' && /^[\w/.-]{1,200}$/.test(input.part) ? input.part : undefined;
    selection = { path: real, story, ...(part ? { part } : {}), start, end, at: Date.now() };
    res.writeHead(204).end();
  };

  const server = createServer((req, res) => {
    handle(req, res).catch((e: unknown) => {
      if (!res.headersSent) json(res, 500, { error: e instanceof Error ? e.message : String(e) });
      else res.end();
    });
  });

  const first = options.port ?? DEFAULT_PORT;
  for (let attempt = 0; ; attempt++) {
    try {
      port = await listen(server, first === 0 ? 0 : first + attempt);
      break;
    } catch (e) {
      if ((e as NodeJS.ErrnoException).code !== 'EADDRINUSE' || first === 0 || attempt + 1 >= PORT_ATTEMPTS) throw e;
    }
  }
  const base = `http://127.0.0.1:${port}`;

  // Keep event streams open through proxies and idle timeouts.
  const heartbeat = setInterval(() => {
    for (const res of clients) res.write(': ping\n\n');
  }, 25_000);
  heartbeat.unref();

  // Changes made outside the server (Word, Finder, another tool).
  const watchers: FSWatcher[] = [];
  const pending = new Map<string, NodeJS.Timeout>();
  // macOS replays recent file events when a watcher starts: only a newer modification time counts.
  const watchStart = Date.now();
  const seenMtime = new Map<string, number>();
  const modifiedSinceLastSeen = async (file: string): Promise<boolean> => {
    let mtime: number;
    try {
      mtime = (await stat(file)).mtimeMs;
    } catch {
      return false; // deleted or renamed away: the document list refresh covers it
    }
    const known = seenMtime.get(file);
    seenMtime.set(file, mtime);
    return known === undefined ? mtime >= watchStart - 1000 : mtime !== known;
  };
  if (options.watch !== false) {
    for (const root of ws.roots) {
      try {
        const watcher = watch(root, { recursive: true }, (_event, name) => {
          if (!name) return;
          const file = join(root, name.toString());
          const base = basename(file);
          if (!DOCUMENT_EXTENSIONS.includes(extname(file).toLowerCase()) || base.startsWith('~$') || base.startsWith('.')) return;
          clearTimeout(pending.get(file));
          pending.set(
            file,
            setTimeout(() => {
              pending.delete(file);
              void modifiedSinceLastSeen(file).then((changed) => {
                send({ type: 'documents' });
                const wrote = recent.get(file);
                if (!changed || (wrote && Date.now() - wrote < 1500)) return;
                send({ type: 'changed', path: ws.display(file) });
              });
            }, 250),
          );
        });
        watcher.on('error', () => {});
        watchers.push(watcher);
      } catch {
        // Recursive watching is not available everywhere; tool changes are still announced.
      }
    }
  }

  return {
    url: base,
    link: (real) => `${base}/?doc=${encodeURIComponent(ws.display(real))}`,
    activity(event) {
      send({ type: 'activity', path: ws.display(event.path), paragraphs: event.paragraphs.slice(0, 500), mode: event.mode, ...(event.label ? { label: event.label } : {}) });
    },
    selection: () => selection,
    notify(event) {
      recent.set(event.path, Date.now());
      send({ type: 'changed', path: ws.display(event.path), ...(event.message ? { message: event.message } : {}), ...(event.comment ? { comment: event.comment } : {}) });
      send({ type: 'documents' });
    },
    async close() {
      clearInterval(heartbeat);
      for (const w of watchers) w.close();
      for (const t of pending.values()) clearTimeout(t);
      for (const res of clients) res.end();
      clients.clear();
      await new Promise<void>((ok) => server.close(() => ok()));
    },
  };
}
