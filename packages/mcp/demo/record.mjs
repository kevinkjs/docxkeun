/**
 * Records the docxkeun-mcp demo video: a scripted session that makes real tool calls to the built
 * server (../dist/cli.js) while its live preview redraws, filmed in headless Chrome.
 *
 *   node demo/record.mjs [output.mp4]      (after `pnpm build` and `pnpm -C packages/mcp build`)
 *
 * Needs Google Chrome (set CHROME to use another Chromium build). Chrome runs with a throwaway
 * profile; the document lives in a temporary folder, so nothing of yours is opened or changed.
 */
import { createServer } from 'node:http';
import { mkdir, mkdtemp, readFile, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { dirname, join, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import { Client } from '@modelcontextprotocol/client';
import { StdioClientTransport } from '@modelcontextprotocol/client/stdio';
import { launchChrome } from './chrome.mjs';
import { writeMp4 } from './mp4.mjs';
import { sampleDocument } from './sample.mjs';

const here = dirname(fileURLToPath(import.meta.url));
const output = resolve(process.argv[2] ?? join(here, 'out', 'docxkeun-mcp-demo.mp4'));
const cli = join(here, '..', 'dist', 'cli.js');
const DOC = 'Q3 sales update.docx';
const WIDTH = 1280;
const HEIGHT = 720;
const SCALE = 1.5; // 1920 × 1080 video
const FPS = 30;

const pause = (ms) => new Promise((r) => setTimeout(r, ms));
const J = JSON.stringify;
const textOf = (r) => r.content.map((c) => c.text ?? '').join('\n');

const work = await mkdtemp(join(tmpdir(), 'docxkeun-demo-'));
const folder = join(work, 'Documents');
const framesDir = join(work, 'frames');
await mkdir(folder);
await mkdir(framesDir);
await writeFile(join(folder, DOC), await sampleDocument());

// The MCP server, started the way Claude Desktop starts it.
const client = new Client({ name: 'docxkeun-demo', version: '1.0.0' });
await client.connect(new StdioClientTransport({ command: process.execPath, args: [cli, '--preview', '--preview-port', '0', folder], stderr: 'ignore' }));
const version = client.getServerVersion()?.version ?? '';
const link = /https?:\/\/\S+/.exec(textOf(await client.callTool({ name: 'get_preview_link', arguments: { path: DOC } })))?.[0];
if (!link) throw new Error('The server gave no preview link');

// Serves the stage and the encoder, and receives the encoded video.
let frames = [];
let clockStart = 0;
let duration = 0;
let videoDone;
const videoWritten = new Promise((r) => (videoDone = r));
const files = { '/stage.html': 'stage.html', '/encoder.html': 'encoder.html' };
const http = createServer(async (req, res) => {
  try {
    const path = new URL(req.url, 'http://x').pathname;
    if (files[path]) {
      res.setHeader('Content-Type', 'text/html; charset=utf-8');
      res.end(await readFile(join(here, files[path])));
    } else if (path === '/frames.json') {
      res.setHeader('Content-Type', 'application/json');
      res.end(J({ frames, width: WIDTH * SCALE, height: HEIGHT * SCALE, fps: FPS, duration, bitrate: 2_200_000 }));
    } else if (path.startsWith('/frames/')) {
      res.setHeader('Content-Type', 'image/jpeg');
      res.end(await readFile(join(framesDir, path.slice('/frames/'.length))));
    } else if (path === '/video' && req.method === 'POST') {
      const body = Buffer.concat(await Array.fromAsync(req));
      const headerLength = body.readUInt32LE(0);
      const { description, samples } = JSON.parse(body.subarray(4, 4 + headerLength).toString('utf8'));
      samples.forEach((s, i) => {
        if (i && s.timestamp <= samples[i - 1].timestamp) throw new Error('Samples out of order (B-frames?)');
      });
      const mp4 = writeMp4({
        width: WIDTH * SCALE, height: HEIGHT * SCALE, timescale: 90_000, description: Uint8Array.from(description),
        samples: samples.map((s) => ({ size: s.size, key: s.key, duration: 90_000 / FPS })),
        data: body.subarray(4 + headerLength),
      });
      await mkdir(dirname(output), { recursive: true });
      await writeFile(output, mp4);
      res.end('ok');
      videoDone(mp4.length);
    } else {
      res.statusCode = 404;
      res.end();
    }
  } catch (e) {
    console.error(e);
    res.statusCode = 500;
    res.end(String(e));
  }
}).listen(0, '127.0.0.1');
await new Promise((r) => http.once('listening', r));
const base = `http://127.0.0.1:${http.address().port}`;

const chrome = await launchChrome();
try {
  const page = await chrome.newPage();
  await page.send('Page.enable');
  await page.send('Emulation.setDeviceMetricsOverride', { width: WIDTH, height: HEIGHT, deviceScaleFactor: SCALE, mobile: false });
  const loaded = new Promise((r) => page.on('Page.loadEventFired', r));
  const note = `Scripted session: real tool calls to docxkeun-mcp ${version}`;
  await page.send('Page.navigate', { url: `${base}/stage.html?preview=${encodeURIComponent(link)}&note=${encodeURIComponent(note)}` });
  await loaded;
  await pause(2500); // the preview loads its fonts and draws the page behind the title card

  // Film: every frame Chrome paints, with its time. Chrome sends nothing while the page is still,
  // so the length comes from the clock: `first` maps frame times to it.
  const writes = [];
  let first;
  page.on('Page.screencastFrame', ({ data, metadata, sessionId }) => {
    void page.send('Page.screencastFrameAck', { sessionId });
    if (first === undefined) {
      first = metadata.timestamp;
      clockStart = performance.now() / 1000;
    }
    const file = `${frames.length}.jpg`;
    frames.push({ t: metadata.timestamp - first, file });
    writes.push(writeFile(join(framesDir, file), Buffer.from(data, 'base64')));
  });
  await page.send('Page.startScreencast', { format: 'jpeg', quality: 92, maxWidth: WIDTH * SCALE, maxHeight: HEIGHT * SCALE, everyNthFrame: 1 });

  const stage = (call) => page.evaluate(`stage.${call}`);
  const firstLine = (s) => s.split('\n').find((l) => l.trim()) ?? '';
  async function tool(name, args, detail, after = 2600) {
    const id = await stage(`tool(${J(name)}, ${J(detail)})`);
    await pause(650);
    const r = await client.callTool({ name, arguments: { path: DOC, ...args } });
    const out = textOf(r);
    if (r.isError) throw new Error(`${name}: ${out}`);
    await pause(200);
    await stage(`done(${id}, ${J(firstLine(out))})`);
    await pause(after);
    return out;
  }
  const paragraphOf = (listing, start) => {
    const m = listing.split('\n').map((l) => /^\[(\d+)\] (?:\([^)]*\) )?(.*)$/.exec(l)).find((x) => x && x[2].startsWith(start));
    if (!m) throw new Error(`No paragraph starting "${start}"`);
    return Number(m[1]);
  };

  await pause(2600);
  await stage('start()');
  await pause(1200);

  // 1. Rewrite a paragraph.
  await stage(`user(${J('The summary reads badly. Can you tidy it up?')})`);
  await pause(400);
  await stage('assistant()');
  const listing = await tool('read_document', {}, DOC, 900);
  const summary = paragraphOf(listing, 'This quarter');
  await tool('edit_paragraph', { paragraph: summary, text: 'Sales grew in every region this quarter, and the online store became our biggest channel for the first time since it launched.' }, `¶${summary}, the summary`, 2800);
  await stage(`reply(${J('Rewrote the summary: half the length, grammar fixed.')})`);
  await pause(1800);

  // 2. A chart from the table's numbers.
  await stage(`user(${J('Add a chart of the sales table.')})`);
  await pause(400);
  await stage('assistant()');
  const below = paragraphOf(listing, 'Next steps') - 1; // the empty paragraph after the table
  await tool('insert_chart', {
    paragraph: below, type: 'column', title: 'Cups sold by region (thousands)',
    categories: ['Q1', 'Q2', 'Q3'],
    series: [{ name: 'Java', values: [120, 135, 152] }, { name: 'Sumatra', values: [80, 86, 97] }, { name: 'Bali', values: [45, 52, 71] }],
    width: 300, alt: 'Column chart of cups sold per quarter in Java, Sumatra and Bali',
  }, `¶${below}, column chart`, 3000);
  await stage(`reply(${J('Added a column chart under the table. Word\'s Edit Data opens its numbers.')})`);
  await pause(1800);

  // 3. Tracked changes.
  await stage(`user(${J('Track your changes, and make the next steps specific.')})`);
  await pause(400);
  await stage('assistant()');
  await tool('set_track_changes', { on: true }, 'on', 700);
  const step1 = paragraphOf(listing, 'Open more stores');
  const step2 = paragraphOf(listing, 'Hire for the online store');
  await tool('edit_paragraph', { paragraph: step1, text: 'Open two stores in Bali before December, where sales grew fastest (+37%)' }, `¶${step1}`, 2400);
  await tool('edit_paragraph', { paragraph: step2, text: 'Hire a second lead for the online store by November' }, `¶${step2}`, 2600);
  await stage(`reply(${J('Both steps now say where and when. The edits are tracked, ready to accept or reject in Word.')})`);
  await pause(2000);

  // 4. A comment.
  await stage(`user(${J('Ask Rina to check the Bali numbers.')})`);
  await pause(400);
  await stage('assistant()');
  const bali = paragraphOf(listing, 'Bali');
  await tool('add_comment', { paragraph: bali, text: 'Bali', comment: 'Rina, can you confirm the Bali figures? Q3 is up 37% on Q2.' }, `¶${bali}, "Bali"`, 3000);
  await stage(`reply(${J('Left Rina a comment on the Bali row.')})`);
  await pause(1800);

  // 5. A PDF.
  await stage(`user(${J('Export it as a PDF for the meeting.')})`);
  await pause(400);
  await stage('assistant()');
  await tool('export_pdf', {}, DOC, 900);
  await stage(`reply(${J('Saved Q3 sales update.pdf, with the chart and the comment.')})`);
  await pause(2400);

  await stage('end()');
  await pause(4800);
  duration = performance.now() / 1000 - clockStart;
  await page.send('Page.stopScreencast');
  await Promise.all(writes);
  console.log(`${frames.length} frames over ${duration.toFixed(1)} s`);

  // Encode in the browser, write the MP4 here.
  const encoderLoaded = new Promise((r) => page.on('Page.loadEventFired', r));
  await page.send('Page.navigate', { url: `${base}/encoder.html` });
  await encoderLoaded;
  console.log(await page.evaluate('encode()'));
  const bytes = await videoWritten;
  console.log(`Wrote ${output} (${(bytes / 1e6).toFixed(1)} MB)`);
} finally {
  // Each step runs even when one before it fails.
  const steps = [() => chrome.close(), () => client.close(), () => http.close(), () => rm(work, { recursive: true, force: true, maxRetries: 10, retryDelay: 200 })];
  for (const step of steps) await Promise.resolve().then(step).catch((e) => console.error('Cleanup:', e.message));
}
