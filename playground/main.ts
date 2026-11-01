import { DocxEditor } from '../src/api/index.js';
import { FontRegistry } from '../src/index.js';
import { DocxViewer } from '../src/render/index.js';
import { sampleDocx } from './sample.js';

const $ = <T extends HTMLElement>(id: string) => document.getElementById(id) as T;
const status = (text: string) => ($('status').textContent = text);

/** Load whatever system fonts the dev server can serve; fall back to canvas metrics if none. */
async function loadFonts(): Promise<FontRegistry> {
  const registry = new FontRegistry();
  const files: Array<[string, string, boolean, boolean]> = [
    ['Arial.ttf', 'Arial', false, false],
    ['Arial Bold.ttf', 'Arial', true, false],
    ['Arial Italic.ttf', 'Arial', false, true],
    ['Arial Bold Italic.ttf', 'Arial', true, true],
    ['Times New Roman.ttf', 'Times New Roman', false, false],
    ['Times New Roman Bold.ttf', 'Times New Roman', true, false],
    ['Times New Roman Italic.ttf', 'Times New Roman', false, true],
    ['Courier New.ttf', 'Courier New', false, false],
    ['Georgia.ttf', 'Georgia', false, false],
    ['Verdana.ttf', 'Verdana', false, false],
  ];
  await Promise.all(
    files.map(async ([file, family, bold, italic]) => {
      try {
        await registry.addFromUrl(`/__fonts/${encodeURIComponent(file)}`, { family, bold, italic });
      } catch {
        // not available on this machine
      }
    }),
  );
  if (!registry.isEmpty) {
    registry.alias('Calibri', 'Arial');
    registry.alias('Calibri Light', 'Arial');
    registry.alias('Cambria', 'Times New Roman');
    registry.fallbacks = ['Arial'];
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
button('Save .docx', async () => download(await viewer.document!.save(), 'edited.docx', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'));

void show(await sampleDocx(), 'sample.docx');
