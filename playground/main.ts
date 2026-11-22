import { DocxEditor } from '../src/api/index.js';
import { FontRegistry, loadBundledFonts } from '../src/index.js';
import { exportPdf } from '../src/pdf/index.js';
import { pdfSecurity } from '../src/security/index.js';
import { DocxViewer, promptPopover, type MenuItem } from '../src/render/index.js';
import { sampleDocx } from './sample.js';
import { fitZoom, readFold, writeFold } from './view.js';

const $ = <T extends HTMLElement>(id: string) => document.getElementById(id) as T;
const status = (text: string) => ($('status').textContent = text);

/** The bundled open fonts (Carlito, Caladea, Liberation), served next to index.html. */
async function loadFonts(): Promise<FontRegistry> {
  const registry = new FontRegistry();
  try {
    await loadBundledFonts(registry, { baseUrl: './' });
  } catch (e) {
    console.warn('Bundled fonts unavailable, using system font metrics', e);
  }
  return registry;
}

const fonts = await loadFonts();
/** Fold states remembered across visits (see view.ts: the toolbar keeps only its fold). */
const store = () => localStorage;
const saved = readFold(store);

const viewer = new DocxViewer($('viewer'), {
  fonts,
  highlightRange: () => true,
  editable: true,
  toolbar: { collapsed: saved.toolbar?.collapsed ?? false },
  navigation: { collapsed: saved.navigation?.collapsed ?? false },
  comments: { mode: 'auto' },
  contextMenu: {
    items: (_at, defaults): MenuItem[] => [
      ...defaults,
      '-',
      { label: 'Lock…', run: () => lockSelection() },
      { label: 'Unlock', run: () => unlockAtCaret() },
    ],
  },
});
if (saved.navigation?.visible === false) viewer.navigation.hide();
viewer.on('chromechange', ({ part, visible, collapsed }) => {
  if (part === 'toolbar' || part === 'navigation') writeFold(store, part, visible, !!collapsed);
  if (part === 'navigation') $('navToggle').setAttribute('aria-pressed', String(visible));
});
$('navToggle').setAttribute('aria-pressed', String(viewer.navigation.visible));
viewer.on('editerror', ({ error }) => status(`Not changed: ${error.message}`));
(window as unknown as { viewer: DocxViewer }).viewer = viewer;
$('navToggle').addEventListener('click', () => {
  if (viewer.navigation.visible) viewer.navigation.hide();
  else viewer.navigation.show();
});

// Alt+Left / Alt+Right walk the navigation history, except in text fields outside the viewer
// (the navigation search box, popovers), where they keep jumping by word.
document.addEventListener('keydown', (e) => {
  if (!e.altKey || e.metaKey || e.ctrlKey) return;
  const t = e.target instanceof HTMLElement ? e.target : undefined;
  if (t?.matches('input, textarea, select, [contenteditable=""], [contenteditable="true"]') && !viewer.root.contains(t)) return;
  if (e.key === 'ArrowLeft' && viewer.back()) e.preventDefault();
  else if (e.key === 'ArrowRight' && viewer.forward()) e.preventDefault();
});

async function show(bytes: Uint8Array | Blob, name: string) {
  const t0 = performance.now();
  try {
    await viewer.load(bytes);
    const pages = viewer.layout?.pages.length ?? 0;
    status(`${name}: ${pages} page${pages === 1 ? '' : 's'} in ${Math.round(performance.now() - t0)} ms${fonts.isEmpty ? ' (system font metrics)' : ''}`);
  } catch (e) {
    status(`Could not open ${name}: ${(e as Error).message}`);
    console.error(e);
  }
}

$<HTMLInputElement>('file').addEventListener('change', (e) => {
  const file = (e.target as HTMLInputElement).files?.[0];
  if (file) void show(file, file.name);
});
document.body.addEventListener('dragover', (e) => e.preventDefault());
document.body.addEventListener('drop', (e) => {
  e.preventDefault();
  const file = e.dataTransfer?.files[0];
  if (file) void show(file, file.name);
});
$('sample').addEventListener('click', async () => show(await sampleDocx(), 'sample.docx'));

// Zoom: the page and its comment column are fitted to the viewer's width (at most 100%, at least
// 50%) until − or + is used; then that zoom stays until the page is reloaded.
let fitWidth = true;
const showZoom = () => ($('zoom').textContent = `${Math.round(viewer.zoom * 100)}%`);
function fitToWidth(): void {
  const layout = viewer.layout;
  if (!fitWidth || !layout || layout.pages.length === 0 || !viewer.root.clientWidth) return;
  const zoom = fitZoom(viewer.root.clientWidth, layout.pages.map((p) => p.width), viewer.comments.width);
  if (Math.abs(viewer.zoom - zoom) > 0.01) viewer.zoom = zoom;
  showZoom();
}
viewer.on('layout', fitToWidth);
viewer.on('chromechange', ({ part }) => {
  if (part === 'comments') fitToWidth();
});
// The viewer's room changes with the window and with the navigation panel opening or folding.
// Only real width changes count (a scrollbar coming and going does not).
{
  let timer: ReturnType<typeof setTimeout> | undefined;
  const refit = () => {
    clearTimeout(timer);
    timer = setTimeout(fitToWidth, 50);
  };
  if (typeof ResizeObserver === 'undefined') window.addEventListener('resize', refit);
  else {
    let width = viewer.root.offsetWidth;
    new ResizeObserver(() => {
      if (viewer.root.offsetWidth === width) return;
      width = viewer.root.offsetWidth;
      refit();
    }).observe(viewer.root);
  }
}
const setZoom = (z: number) => {
  fitWidth = false;
  viewer.zoom = z;
  showZoom();
};
$('zoomIn').addEventListener('click', () => setZoom(viewer.zoom * 1.25));
$('zoomOut').addEventListener('click', () => setZoom(viewer.zoom / 1.25));
viewer.on('selectionchange', (sel) => {
  if (sel) status(`Selected "${sel.start.inline.text.slice(sel.start.offset, sel.start.inline === sel.end.inline ? sel.end.offset : undefined)}"`);
});

function button(label: string, onClick: () => unknown): void {
  const b = document.createElement('button');
  b.textContent = label;
  b.addEventListener('click', async () => {
    try {
      await onClick();
    } catch (e) {
      status((e as Error).message);
      console.error(e);
    }
  });
  $('actions').appendChild(b);
}

function download(bytes: Uint8Array, name: string, type: string): void {
  const url = URL.createObjectURL(new Blob([bytes as BlobPart], { type }));
  const a = document.createElement('a');
  a.href = url;
  a.download = name;
  a.click();
  setTimeout(() => URL.revokeObjectURL(url), 1000);
}

const editor = () => new DocxEditor(viewer.document!);

/** The current selection as an editor range, from in-page editing or the browser selection. */
function selectionRange() {
  const ed = editor();
  const sel = viewer.editing?.isEnabled ? viewer.editing.selection : undefined;
  if (sel && (sel.anchor.para !== sel.focus.para || sel.anchor.char !== sel.focus.char)) {
    return viewer.editing!.rangeOf(sel.anchor, sel.focus);
  }
  const native = viewer.getSelection();
  return native ? ed.rangeFromSelection(native) : undefined;
}

/** Show a PDF in the browser's own viewer, in an overlay. */
function showPdf(bytes: Uint8Array): void {
  document.getElementById('pdf-overlay')?.remove();
  const overlay = document.createElement('div');
  overlay.id = 'pdf-overlay';
  overlay.style.cssText = 'position:fixed;inset:48px 0 0 0;background:#525659;z-index:10;display:flex;flex-direction:column';
  const close = document.createElement('button');
  close.textContent = 'Close PDF';
  close.style.cssText = 'align-self:flex-end;margin:6px';
  close.onclick = () => overlay.remove();
  const frame = document.createElement('iframe');
  frame.style.cssText = 'flex:1;border:0;width:100%';
  frame.src = URL.createObjectURL(new Blob([bytes as BlobPart], { type: 'application/pdf' }));
  overlay.append(close, frame);
  document.body.appendChild(overlay);
}

button('Editing: on', function (this: void) {
  const on = !viewer.editing?.isEnabled;
  viewer.setEditable(on);
  const b = [...document.querySelectorAll('#actions button')].find((x) => x.textContent?.startsWith('Editing'));
  if (b) b.textContent = `Editing: ${on ? 'on' : 'off'}`;
  status(on ? 'Click in the text and type' : 'View only: select text to copy or comment');
});
button('Undo', () => viewer.editing?.undo());
button('Redo', () => viewer.editing?.redo());
/** Run an edit as one undo step (the editing layer is there once the viewer has been editable). */
function edit<T>(label: string, fn: (ed: DocxEditor) => T | Promise<T>): Promise<T> {
  return viewer.editing ? viewer.editing.transaction(label, fn) : Promise.resolve(fn(editor()));
}

/** A viewport point near the document, for popovers opened from the header buttons. */
function anchorPoint(): { x: number; y: number } {
  const r = viewer.root.getBoundingClientRect();
  return { x: r.left + r.width / 2 - 130, y: r.top + 60 };
}

async function lockSelection(at = anchorPoint()): Promise<void> {
  const range = selectionRange();
  if (!range) return void status('Select the text to lock first');
  const title = await promptPopover({ title: 'Name of the locked region (optional)', placeholder: 'Optional', at });
  if (title === undefined) return;
  const id = await edit('Lock', (ed) => ed.lockRange(range, { title: title || undefined }));
  status(`Locked region ${id}${title ? ` "${title}"` : ''}: it can no longer be edited`);
}

async function unlockAtCaret(): Promise<void> {
  const sel = viewer.editing?.selection;
  const range = selectionRange() ?? (sel ? viewer.editing!.rangeOf(sel.anchor, sel.anchor) : undefined);
  if (!range) return void status('Click inside a locked region first');
  const contains = (root: { children: unknown[] }, node: unknown): boolean =>
    root.children.some((c) => c === node || (typeof c === 'object' && c !== null && 'children' in c && contains(c as { children: unknown[] }, node)));
  const region = editor().lockedRegions().find((r) => contains(r.el, range.start.node));
  if (!region) return void status('The caret is not inside a locked region');
  await edit('Unlock', (ed) => ed.unlock(region.id));
  status(`Unlocked${region.title ? ` "${region.title}"` : ''}`);
}

button('Comment', () => {
  if (!viewer.editing?.isEnabled) return status('Turn editing on to comment');
  if (!selectionRange()) return status('Select some text first');
  viewer.comments.compose();
});
button('Link', async () => {
  const range = selectionRange();
  if (!range) return status('Select some text first');
  const url = await promptPopover({
    title: 'Link address',
    placeholder: 'https://… or #bookmark',
    at: anchorPoint(),
    validate: (v) => (v.trim() ? undefined : 'Enter a web address, or #bookmark'),
  });
  if (!url) return;
  const target = url.trim();
  await edit('Add link', (ed) => ed.addHyperlink(range, target.startsWith('#') ? { bookmark: target.slice(1) } : { url: target }));
  status('Link added');
});
button('Bold', async () => {
  const range = selectionRange();
  if (!range) return status('Select some text first');
  await edit('Bold', (ed) => ed.formatRange(range, { bold: true }));
});
button('Lock', () => lockSelection());
button('Unlock', () => unlockAtCaret());
button('Export PDF', async () => {
  if (fonts.isEmpty) return status('PDF export needs font files');
  const t0 = performance.now();
  const bytes = await exportPdf(viewer.document!, { fonts, layout: viewer.layout });
  status(`PDF: ${Math.round(bytes.length / 1024)} KB in ${Math.round(performance.now() - t0)} ms`);
  showPdf(bytes);
});
button('Protected PDF', async () => {
  if (fonts.isEmpty) return status('PDF export needs font files');
  const password = prompt('Password to open (leave empty for none; copying and editing are still restricted)') ?? '';
  const bytes = await exportPdf(viewer.document!, { fonts, layout: viewer.layout, encryption: pdfSecurity({ userPassword: password, permissions: { copy: false, modify: false } }) });
  status(`Encrypted PDF (AES-256): ${Math.round(bytes.length / 1024)} KB`);
  showPdf(bytes);
});
button('Save .docx', async () => download(await viewer.document!.save(), 'edited.docx', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'));

void show(await sampleDocx(), 'sample.docx');
