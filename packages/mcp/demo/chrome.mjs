/** Headless Chrome over the DevTools protocol, with a temporary profile. */
import { spawn } from 'node:child_process';
import { mkdtemp, rm } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';

const CHROME = process.env.CHROME ?? '/Applications/Google Chrome.app/Contents/MacOS/Google Chrome';

export async function launchChrome() {
  const profile = await mkdtemp(join(tmpdir(), 'docxkeun-demo-chrome-'));
  const proc = spawn(CHROME, ['--headless=new', '--remote-debugging-port=0', `--user-data-dir=${profile}`, '--no-first-run', '--no-default-browser-check',
    '--hide-scrollbars', '--mute-audio', '--disable-background-timer-throttling', '--disable-renderer-backgrounding', 'about:blank'], { stdio: ['ignore', 'ignore', 'pipe'] });
  const wsUrl = await new Promise((resolve, reject) => {
    let buf = '';
    proc.stderr.on('data', (d) => {
      buf += d;
      const m = /DevTools listening on (ws:\S+)/.exec(buf);
      if (m) resolve(m[1]);
    });
    proc.once('exit', (code) => reject(new Error(`Chrome exited (${code}): ${buf}`)));
  });
  const browser = await connect(wsUrl);
  return {
    browser,
    /** A new page (tab) with its own session. */
    async newPage() {
      const { targetId } = await browser.send('Target.createTarget', { url: 'about:blank' });
      const { sessionId } = await browser.send('Target.attachToTarget', { targetId, flatten: true });
      return {
        send: (method, params = {}) => browser.send(method, params, sessionId),
        on: (event, fn) => browser.on(event, fn, sessionId),
        async evaluate(expression) {
          const r = await browser.send('Runtime.evaluate', { expression, awaitPromise: true, returnByValue: true }, sessionId);
          if (r.exceptionDetails) throw new Error(r.exceptionDetails.exception?.description ?? r.exceptionDetails.text);
          return r.result.value;
        },
      };
    },
    async close() {
      browser.close();
      proc.kill();
      await new Promise((r) => proc.once('exit', r));
      // Chrome's helper processes can still be writing to the profile for a moment.
      await rm(profile, { recursive: true, force: true, maxRetries: 10, retryDelay: 200 });
    },
  };
}

async function connect(url) {
  const ws = new WebSocket(url);
  await new Promise((resolve, reject) => { ws.onopen = resolve; ws.onerror = reject; });
  let id = 0;
  const pending = new Map();
  const listeners = [];
  ws.onmessage = (m) => {
    const msg = JSON.parse(m.data);
    if (msg.id !== undefined) {
      const p = pending.get(msg.id);
      pending.delete(msg.id);
      if (msg.error) p?.reject(new Error(`${p.method}: ${msg.error.message}`));
      else p?.resolve(msg.result);
    } else {
      for (const l of listeners) if (l.event === msg.method && l.sessionId === msg.sessionId) l.fn(msg.params);
    }
  };
  return {
    send(method, params = {}, sessionId) {
      const msgId = ++id;
      ws.send(JSON.stringify({ id: msgId, method, params, ...(sessionId ? { sessionId } : {}) }));
      return new Promise((resolve, reject) => pending.set(msgId, { resolve, reject, method }));
    },
    on(event, fn, sessionId) { listeners.push({ event, fn, sessionId }); },
    close() { ws.close(); },
  };
}
