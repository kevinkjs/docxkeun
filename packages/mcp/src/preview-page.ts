/** The live preview page: docxkeun's viewer, a document picker, and a feed of what changed. */

export const PREVIEW_HTML: string = `<!doctype html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<meta name="docxkeun-token" content="%TOKEN%">
<meta name="docxkeun-ai" content="%AI%">
<title>docxkeun live preview</title>
<style>
  :root { --line: #dadce0; --muted: #5f6368; --text: #202124; }
  html, body { margin: 0; height: 100%; font: 14px system-ui, sans-serif; color: var(--text); background: #fff; }
  body { display: flex; flex-direction: column; }
  header { display: flex; gap: 12px; align-items: center; padding: 8px 12px; border-bottom: 1px solid var(--line); flex-wrap: wrap; }
  header h1 { font-size: 15px; margin: 0; }
  select { font: inherit; max-width: min(420px, 60vw); }
  #status { display: inline-flex; align-items: center; gap: 6px; color: var(--muted); }
  #status::before { content: ''; width: 8px; height: 8px; border-radius: 50%; background: #9aa0a6; }
  #status.live::before { background: #1e8e3e; }
  label { color: var(--muted); display: inline-flex; gap: 4px; align-items: center; }
  .zoom { display: inline-flex; gap: 4px; align-items: center; }
  .zoom button { font: inherit; min-width: 28px; }
  #zoom { min-width: 44px; text-align: center; color: var(--muted); }
  #fit[aria-pressed="true"] { background: #e8f0fe; border-color: #1a73e8; }
  #pages { color: var(--muted); margin-left: auto; }
  #sel { display: inline-flex; align-items: center; gap: 6px; max-width: min(460px, 70vw); padding: 3px 8px; border-radius: 12px; background: #f3effe; color: #4b2fb5; font-size: 13px; }
  #sel[hidden] { display: none; }
  #sel span { overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
  #sel::before { content: '✦'; color: #d97757; }
  main { flex: 1; min-height: 0; position: relative; }
  #viewer { height: 100%; }
  #message { position: absolute; inset: 0; display: flex; align-items: center; justify-content: center; background: #f8f9fa; color: var(--muted); padding: 24px; text-align: center; }
  #message[hidden] { display: none; }
  #feed { position: fixed; right: 16px; bottom: 16px; display: flex; flex-direction: column; gap: 8px; max-width: min(420px, calc(100vw - 32px)); z-index: 10; }
  .change { background: var(--text); color: #fff; padding: 10px 14px; border-radius: 8px; box-shadow: 0 2px 8px rgba(0, 0, 0, .25); white-space: pre-wrap; animation: in .2s ease-out; }
  .change small { display: block; color: #bdc1c6; margin-bottom: 2px; }
  @keyframes in { from { opacity: 0; transform: translateY(8px); } }
</style>
<script type="module" src="./preview.js"></script>
</head>
<body>
<header>
  <h1>docxkeun live preview</h1>
  <select id="doc" aria-label="Document"></select>
  <span id="status">Connecting…</span>
  <label><input type="checkbox" id="follow" checked> Follow changes</label>
  <label title="How the AI's work shows: shimmer, caret and change marks">AI effects <select id="effects"><option value="animated">Animated</option><option value="still">Still</option><option value="off">Off</option></select></label>
  <label title="How the paragraph the AI is changing looks">Style <select id="aiStyle"></select></label>
  <label title="Keep a mark on what the AI changed until its next change"><input type="checkbox" id="aiMarks" checked> Marks</label>
  <span class="zoom"><button id="zoomOut" aria-label="Zoom out">−</button><span id="zoom">100%</span><button id="zoomIn" aria-label="Zoom in">+</button><button id="fit" aria-pressed="true">Fit width</button></span>
  <span id="sel" hidden title="The AI can read this selection with get_selection"><span></span></span>
  <span id="pages"></span>
</header>
<main>
  <div id="viewer"></div>
  <div id="message" hidden></div>
</main>
<div id="feed" aria-live="polite"></div>
</body>
</html>
`;

export const PREVIEW_SCRIPT: string = `import { FontRegistry, loadBundledFonts } from './lib/index.js';
import { ACTIVITY_STYLES, DocxViewer } from './lib/render/index.js';

const $ = (id) => document.getElementById(id);
const picker = $('doc');
const status = $('status');
const message = $('message');
const feed = $('feed');
const follow = $('follow');
let current = new URLSearchParams(location.search).get('doc');
const token = document.querySelector('meta[name="docxkeun-token"]').content;
let loaded;
// read_document's number for each paragraph the viewer counts (text boxes are numbered there, not here).
let numbering = [];
let workingSince = 0;
const sleep = (ms) => new Promise((ok) => setTimeout(ok, ms));

const fonts = new FontRegistry();
try {
  await loadBundledFonts(fonts, { baseUrl: new URL('./fonts/', location.href) });
} catch (e) {
  console.warn('Bundled fonts unavailable; using system font metrics', e);
}
const viewer = new DocxViewer($('viewer'), {
  fonts,
  showComments: true,
  navigation: { collapsed: true },
  comments: { mode: 'auto' },
  // Read-only: only what copies.
  contextMenu: { items: (at, d) => d.filter((i) => i !== '-' && ['Copy', 'Copy link address', 'Open link'].includes(i.label)) },
});

// AI effects: animated, still (highlights without motion) or off; remembered in this browser.
const effects = $('effects');
function applyEffects(value) {
  viewer.activity.enabled = value !== 'off';
  viewer.activity.animate = value === 'animated';
}
try {
  const saved = localStorage.getItem('docxkeun-ai-effects');
  if (saved && ['animated', 'still', 'off'].includes(saved)) effects.value = saved;
} catch {}
applyEffects(effects.value);
effects.addEventListener('change', () => {
  applyEffects(effects.value);
  try {
    localStorage.setItem('docxkeun-ai-effects', effects.value);
  } catch {}
});

// The style of the AI's work and whether its marks stay; remembered in this browser.
const aiStyle = $('aiStyle');
const aiMarks = $('aiMarks');
for (const s of ACTIVITY_STYLES) {
  const option = new Option(s.label, s.style);
  option.title = s.description;
  aiStyle.add(option);
}
try {
  const saved = localStorage.getItem('docxkeun-ai-style');
  if (saved && ACTIVITY_STYLES.some((s) => s.style === saved)) aiStyle.value = saved;
  aiMarks.checked = localStorage.getItem('docxkeun-ai-marks') !== 'off';
} catch {}
viewer.activity.style = aiStyle.value;
viewer.activity.marksShown = aiMarks.checked;
aiStyle.addEventListener('change', () => {
  viewer.activity.style = aiStyle.value;
  try {
    localStorage.setItem('docxkeun-ai-style', aiStyle.value);
  } catch {}
});
aiMarks.addEventListener('change', () => {
  viewer.activity.marksShown = aiMarks.checked;
  try {
    localStorage.setItem('docxkeun-ai-marks', aiMarks.checked ? 'on' : 'off');
  } catch {}
});

// The AI's name and colours on the pages, from the server's --author and --ai-color.
try {
  const ai = JSON.parse(document.querySelector('meta[name="docxkeun-ai"]').content || '{}');
  if (typeof ai.name === 'string' && ai.name) viewer.activity.name = ai.name;
  if (ai.color) viewer.activity.color = ai.color;
} catch (e) {
  console.warn('The AI colour was not applied', e);
}

// Zoom: "Fit width" (the default) keeps the page and its comments inside narrow panes.
let fit = true;
function showZoom() {
  $('zoom').textContent = Math.round(viewer.zoom * 100) + '%';
  $('fit').setAttribute('aria-pressed', String(fit));
}
function applyFit() {
  const layout = viewer.layout;
  if (!fit || !layout || !layout.pages.length) return showZoom();
  const pageWidth = Math.max(...layout.pages.map((p) => p.width)) * 96 / 72;
  // Row padding (32), the comment column as the viewer sized it (0 without comments) and a scrollbar (16).
  const extra = 32 + viewer.comments.width + 16;
  const zoom = Math.max(0.25, Math.min(1, ((viewer.root.clientWidth || $('viewer').clientWidth) - extra) / pageWidth));
  if (Math.abs(viewer.zoom - zoom) > 0.01) viewer.zoom = zoom;
  showZoom();
}
function zoomBy(factor) {
  fit = false;
  viewer.zoom = Math.round(viewer.zoom * factor * 20) / 20;
  showZoom();
}
$('zoomIn').addEventListener('click', () => zoomBy(1.1));
$('zoomOut').addEventListener('click', () => zoomBy(1 / 1.1));
$('fit').addEventListener('click', () => {
  fit = true;
  applyFit();
});
viewer.on('chromechange', (e) => {
  if (e.part === 'comments') applyFit();
});
// Refit whenever the document's room changes: the window, and the navigation panel opening or
// folding beside it. Only real width changes count (a scrollbar coming and going does not).
let resizeTimer;
function refit() {
  clearTimeout(resizeTimer);
  resizeTimer = setTimeout(applyFit, 50);
}
if (typeof ResizeObserver === 'undefined') window.addEventListener('resize', refit);
else {
  let rootWidth = viewer.root.offsetWidth;
  new ResizeObserver(() => {
    if (viewer.root.offsetWidth === rootWidth) return;
    rootWidth = viewer.root.offsetWidth;
    refit();
  }).observe(viewer.root);
}

function note(path, text) {
  const item = document.createElement('div');
  item.className = 'change';
  const where = document.createElement('small');
  where.textContent = path + ' · ' + new Date().toLocaleTimeString();
  item.append(where, text);
  feed.append(item);
  while (feed.children.length > 3) feed.firstElementChild.remove();
  setTimeout(() => item.remove(), 8000);
}

function showMessage(text) {
  message.textContent = text;
  message.hidden = false;
}

async function refreshList() {
  const docs = await (await fetch('./api/documents')).json();
  if (!current && docs.length) current = [...docs].sort((a, b) => b.modified.localeCompare(a.modified))[0].path;
  picker.replaceChildren(...docs.map((d) => new Option(d.path, d.path)));
  if (current && !docs.some((d) => d.path === current)) picker.append(new Option(current + ' (missing)', current));
  picker.value = current ?? '';
  if (!docs.length) showMessage('No Word documents in this folder yet.');
}

async function load(path, focus) {
  // The AI changed the document on show: keep the old text to animate what changed.
  const before = focus && focus.message && loaded === path && viewer.document ? viewer.activity.snapshot() : undefined;
  current = path;
  picker.value = path;
  history.replaceState(null, '', '?doc=' + encodeURIComponent(path));
  const res = await fetch('./api/file?path=' + encodeURIComponent(path));
  if (!res.ok) {
    const body = await res.json().catch(() => ({ error: res.statusText }));
    showMessage(body.error);
    $('pages').textContent = '';
    return;
  }
  message.hidden = true;
  numbering = await fetch('./api/numbering?path=' + encodeURIComponent(path)).then((r) => (r.ok ? r.json() : [])).catch(() => []);
  const bytes = new Uint8Array(await res.arrayBuffer());
  // Let the "working" shimmer show for a moment before the result appears.
  if (before) await sleep(Math.max(0, workingSince + 700 - Date.now()));
  // The AI's marks build up while it works on a document; another document starts clean.
  if (loaded !== path) viewer.activity.clearMarks();
  await viewer.load(bytes);
  loaded = path;
  document.title = path + ' · docxkeun live preview';
  applyFit();
  const pages = viewer.layout ? viewer.layout.pages.length : 0;
  $('pages').textContent = pages + (pages === 1 ? ' page' : ' pages');
  if (before) viewer.activity.reveal(before, { scroll: follow.checked });
  const anchor = focus && focus.comment ? viewer.layout?.commentAnchors.get(focus.comment) : undefined;
  if (anchor) {
    viewer.scrollToComment(focus.comment);
    viewer.flash({ page: anchor.page, rects: [anchor.rect] });
  }
}

// Changes can arrive faster than a document lays out: keep only the newest request.
let next;
let busy = false;
async function show(path, focus) {
  next = { path, focus };
  if (busy) return;
  busy = true;
  try {
    while (next) {
      const job = next;
      next = undefined;
      try {
        await load(job.path, job.focus);
      } catch (e) {
        showMessage('Could not show ' + job.path + ': ' + e.message);
      }
    }
  } finally {
    busy = false;
  }
}

picker.addEventListener('change', () => show(picker.value));

let connectedBefore = false;
const events = new EventSource('./api/events');
events.onopen = () => {
  status.textContent = 'Live';
  status.classList.add('live');
  // After a reconnect, changes may have been missed.
  if (connectedBefore && current) show(current);
  connectedBefore = true;
};
events.onerror = () => {
  status.textContent = 'Reconnecting…';
  status.classList.remove('live');
};
events.onmessage = (e) => {
  const event = JSON.parse(e.data);
  if (event.type === 'documents') refreshList();
  if (event.type === 'activity') {
    if (event.path !== current) return;
    workingSince = Date.now();
    viewer.activity.work(event.paragraphs, { mode: event.mode, label: event.label, scroll: follow.checked });
    // Reads change nothing: the shimmer ends by itself.
    if (event.mode === 'read') setTimeout(() => viewer.activity.idle(), 1600);
    return;
  }
  if (event.type !== 'changed') return;
  const text = event.message ?? 'Changed on disk';
  if (event.path === current || (follow.checked && event.message)) {
    show(event.path, event);
    note(event.path, text);
  } else if (event.message) {
    note(event.path, text);
  }
};

// Share what is selected with the AI (get_selection), and show that it is shared.
let selTimer;
viewer.on('selectioninfo', ({ selection }) => {
  clearTimeout(selTimer);
  selTimer = setTimeout(() => shareSelection(selection), 200);
});
const readNumber = (n) => numbering[n - 1] ?? n;
function shareSelection(sel) {
  if (!current) return;
  const chip = $('sel');
  chip.hidden = !sel;
  if (sel) {
    const first = sel.story === 'body' ? readNumber(sel.start.paragraph) : sel.start.paragraph;
    const last = sel.story === 'body' ? readNumber(sel.end.paragraph) : sel.end.paragraph;
    const where = first === last ? '¶' + first : '¶' + first + '–' + last;
    const text = sel.text.replace(/\\s+/g, ' ').trim();
    chip.firstElementChild.textContent = sel.collapsed || !text ? 'Cursor at ' + where : where + ': “' + (text.length > 48 ? text.slice(0, 47) + '…' : text) + '”';
  }
  const body = sel ? { path: current, story: sel.story, part: sel.part, start: sel.start, end: sel.end } : { path: current, clear: true };
  fetch('./api/selection', { method: 'POST', headers: { 'content-type': 'application/json', 'x-docxkeun-token': token }, body: JSON.stringify(body) }).catch(() => {});
}

await refreshList();
if (current) await show(current);
`;
