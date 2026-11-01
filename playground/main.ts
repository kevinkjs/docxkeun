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

void show(await sampleDocx(), 'sample.docx');
