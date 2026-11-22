/**
 * The docxkeun chat relay: makes the Claude API calls for the page, from this computer instead of
 * from the browser. Organizations whose settings block browser ("CORS") calls allow these, and the
 * key never has to be stored in the browser: the page hands it over once and the relay keeps it in
 * memory (or takes ANTHROPIC_API_KEY from the environment).
 *
 * Routes (all for pages served from this same local server only):
 *   GET    /relay/status        → { hasKey, fromEnv }
 *   POST   /relay/key           { key, model } → checks the key with the API, then keeps it
 *   DELETE /relay/key           forget the key
 *   *      /anthropic/v1/...    forwarded to https://api.anthropic.com/v1/... with the key added
 *
 * Only the calls the chat makes are forwarded (Messages and Models), only these request headers go
 * upstream, and a request from another website, another host name, or without the page's own
 * X-Docxkeun-Relay header is refused, so other sites cannot spend the key through the relay.
 */

const UPSTREAM = 'https://api.anthropic.com';
const FORWARD_REQUEST_HEADERS = ['content-type', 'accept', 'anthropic-version', 'anthropic-beta'];
const FORWARD_RESPONSE_HEADERS = ['content-type', 'request-id', 'retry-after', 'x-should-retry'];
const ALLOWED_PATHS = [/^\/v1\/messages$/, /^\/v1\/messages\/count_tokens$/, /^\/v1\/models$/, /^\/v1\/models\/[\w.-]+$/];
const LOCAL_HOSTS = ['localhost', '127.0.0.1', '[::1]'];
const MAX_BODY = 32 * 1024 * 1024;

/**
 * @typedef {import('node:http').IncomingMessage} Req
 * @typedef {import('node:http').ServerResponse} Res
 * @typedef {{ apiKey?: string, upstream?: string, fetch?: typeof fetch }} RelayOptions
 */

/** @param {Res} res @param {number} status @param {unknown} body */
function json(res, status, body) {
  res.statusCode = status;
  res.setHeader('content-type', 'application/json; charset=utf-8');
  res.setHeader('cache-control', 'no-store');
  res.end(JSON.stringify(body));
}

/** An error in the API's own format, so the SDK on the page raises the matching error class. */
/** @param {Res} res @param {number} status @param {string} type @param {string} message */
function apiError(res, status, type, message) {
  json(res, status, { type: 'error', error: { type, message } });
}

/** @param {Req} req @param {number} limit @returns {Promise<Buffer>} */
function readBody(req, limit) {
  return new Promise((resolve, reject) => {
    /** @type {Buffer[]} */
    const chunks = [];
    let size = 0;
    req.on('data', (chunk) => {
      size += chunk.length;
      if (size > limit) {
        reject(new Error('Request body too large'));
        req.destroy();
        return;
      }
      chunks.push(chunk);
    });
    req.on('end', () => resolve(Buffer.concat(chunks)));
    req.on('error', reject);
  });
}

/** Whether a request comes from a page served by this same local server. @param {Req} req */
export function fromOwnPage(req) {
  const host = req.headers.host ?? '';
  const hostname = host.replace(/:\d+$/, '');
  if (!LOCAL_HOSTS.includes(hostname)) return false;
  const origin = req.headers.origin;
  if (origin !== undefined && origin !== `http://${host}`) return false;
  // A custom header: browsers send it cross-site only after a CORS preflight, which the relay never allows.
  return req.headers['x-docxkeun-relay'] === '1';
}

/** @param {RelayOptions} [options] */
export function createRelay(options = {}) {
  const upstream = options.upstream ?? UPSTREAM;
  const fetchImpl = options.fetch ?? fetch;
  const envKey = options.apiKey ?? process.env.ANTHROPIC_API_KEY;
  /** @type {string | undefined} */
  let key = envKey || undefined;
  let fromEnv = !!key;

  /** @param {string} candidate @param {string} model */
  async function checkKey(candidate, model) {
    return fetchImpl(`${upstream}/v1/models/${encodeURIComponent(model)}`, {
      headers: { 'x-api-key': candidate, 'anthropic-version': '2023-06-01', accept: 'application/json' },
    });
  }

  /**
   * Connect-style middleware: handles /relay/* and /anthropic/*, passes everything else on.
   * @param {Req} req @param {Res} res @param {() => void} next
   */
  return async function relay(req, res, next) {
    const url = new URL(req.url ?? '/', 'http://relay.local');
    const path = url.pathname;
    if (!path.startsWith('/relay/') && !path.startsWith('/anthropic/')) return next();
    if (!fromOwnPage(req)) return apiError(res, 403, 'permission_error', 'The docxkeun chat relay only answers its own page on this computer.');

    try {
      if (path === '/relay/status' && req.method === 'GET') return json(res, 200, { hasKey: !!key, fromEnv });

      if (path === '/relay/key' && req.method === 'POST') {
        let body;
        try {
          body = JSON.parse((await readBody(req, 16 * 1024)).toString('utf8'));
        } catch {
          return apiError(res, 400, 'invalid_request_error', 'Send {"key": "...", "model": "..."}');
        }
        const candidate = typeof body?.key === 'string' ? body.key.trim() : '';
        const model = typeof body?.model === 'string' && /^[\w.-]+$/.test(body.model) ? body.model : 'claude-opus-5-5';
        if (!candidate) return apiError(res, 400, 'invalid_request_error', 'No key given.');
        let check;
        try {
          check = await checkKey(candidate, model);
        } catch (e) {
          return apiError(res, 502, 'api_error', `Could not reach the Claude API: ${e instanceof Error ? e.message : String(e)}`);
        }
        if (!check.ok) {
          // Pass the API's own answer through, so the page can say why.
          const text = await check.text();
          res.statusCode = check.status;
          res.setHeader('content-type', check.headers.get('content-type') ?? 'application/json');
          res.setHeader('cache-control', 'no-store');
          return res.end(text);
        }
        key = candidate;
        fromEnv = false;
        return json(res, 200, { hasKey: true, fromEnv });
      }

      if (path === '/relay/key' && req.method === 'DELETE') {
        key = undefined;
        fromEnv = false;
        res.statusCode = 204;
        return res.end();
      }

      if (path.startsWith('/anthropic/')) {
        const apiPath = path.slice('/anthropic'.length);
        if (!ALLOWED_PATHS.some((p) => p.test(apiPath))) return apiError(res, 404, 'not_found_error', `The relay does not forward ${apiPath}.`);
        if (req.method !== 'GET' && req.method !== 'POST') return apiError(res, 405, 'invalid_request_error', 'Method not allowed.');
        if (!key) return apiError(res, 401, 'authentication_error', 'No API key yet: enter one in the chat.');

        const headers = new Headers();
        for (const name of FORWARD_REQUEST_HEADERS) {
          const value = req.headers[name];
          if (typeof value === 'string') headers.set(name, value);
        }
        headers.set('x-api-key', key);
        const body = req.method === 'POST' ? new Uint8Array(await readBody(req, MAX_BODY)) : undefined;

        // Stop the upstream call when the page goes away (the Stop button, a reload).
        const controller = new AbortController();
        res.on('close', () => {
          if (!res.writableEnded) controller.abort();
        });
        let answer;
        try {
          answer = await fetchImpl(`${upstream}${apiPath}${url.search}`, { method: req.method, headers, body, signal: controller.signal });
        } catch (e) {
          if (controller.signal.aborted) return;
          return apiError(res, 502, 'api_error', `Could not reach the Claude API: ${e instanceof Error ? e.message : String(e)}`);
        }
        res.statusCode = answer.status;
        for (const [name, value] of answer.headers) {
          if (FORWARD_RESPONSE_HEADERS.includes(name) || name.startsWith('anthropic-ratelimit-')) res.setHeader(name, value);
        }
        res.setHeader('cache-control', 'no-store');
        if (answer.body) {
          try {
            // Streamed replies (server-sent events) are passed on chunk by chunk.
            for await (const chunk of answer.body) res.write(chunk);
          } catch {
            // The page went away or the upstream stream broke; nothing more to send.
          }
        }
        return res.end();
      }

      return apiError(res, 404, 'not_found_error', 'Not found.');
    } catch (e) {
      if (!res.headersSent) return apiError(res, 500, 'api_error', e instanceof Error ? e.message : String(e));
      res.end();
    }
  };
}
