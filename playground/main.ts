import { DocxEditor } from '../src/api/index.js';
import { FontRegistry, loadBundledFonts } from '../src/index.js';
import { exportPdf } from '../src/pdf/index.js';
import { pdfSecurity } from '../src/security/index.js';
import { DocxViewer } from '../src/render/index.js';
import { sampleDocx } from './sample.js';

const $ = <T extends HTMLElement>(id: string) => document.getElementById(id) as T;
const status = (text: string) => ($('status').textContent = text);

/** The bundled open fonts (Carlito, Caladea, Liberation), served from the site root. */
async function loadFonts(): Promise<FontRegistry> {
  const registry = new FontRegistry();
  try {
    await loadBundledFonts(registry, { baseUrl: '/' });
  } catch (e) {
    console.warn('Bundled fonts unavailable, using system font metrics', e);
  }
  return registry;
}

const fonts = await loadFonts();
const viewer = new DocxViewer($('viewer'), { fonts, highlightRange: () => true });
(window as unknown as { viewer: DocxViewer }).viewer = viewer;

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
const setZoom = (z: number) => {
  viewer.zoom = z;
  $('zoom').textContent = `${Math.round(viewer.zoom * 100)}%`;
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

button('Comment', async () => {
  const sel = viewer.getSelection();
  if (!sel) return status('Select some text first');
  const text = prompt('Comment');
  if (!text) return;
  await editor().addComment(editor().rangeFromSelection(sel), { author: 'Playground user', text });
  status('Comment added');
});
button('Link', async () => {
  const sel = viewer.getSelection();
  if (!sel) return status('Select some text first');
  const url = prompt('URL (https://...) or #bookmark');
  if (!url) return;
  const ed = editor();
  await ed.addHyperlink(ed.rangeFromSelection(sel), url.startsWith('#') ? { bookmark: url.slice(1) } : { url });
  status('Link added');
});
button('Bold', () => {
  const sel = viewer.getSelection();
  if (!sel) return status('Select some text first');
  const ed = editor();
  ed.formatRange(ed.rangeFromSelection(sel), { bold: true });
});
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
