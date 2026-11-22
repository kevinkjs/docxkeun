// @vitest-environment happy-dom
import { afterEach, describe, expect, it, vi } from 'vitest';
import { EditError } from '../src/api/index.js';
import { readComments } from '../src/comments/read.js';
import { UndoHistory } from '../src/render/undo.js';
import { DocxViewer } from '../src/render/viewer.js';
import { readFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { BUNDLED_FAMILIES, BUNDLED_FONTS, FontRegistry } from '../src/fonts/index.js';
import { buildDocx, p } from './helpers/docx.js';

// The bundled fonts, read by path (happy-dom's URL is not one node:fs takes).
const fonts = new FontRegistry();
const FONTS_DIR = join(dirname(fileURLToPath(import.meta.url)), '../packages/fonts');
for (const f of BUNDLED_FONTS) fonts.add(new Uint8Array(readFileSync(join(FONTS_DIR, f.file))), { family: f.family, bold: f.bold, italic: f.italic });
fonts.fallbacks = [...BUNDLED_FAMILIES];

async function editableViewer() {
  // happy-dom has no FontFace: the registry's metrics are all layout needs.
  vi.stubGlobal('FontFace', class {
    load() {
      return Promise.resolve(this);
    }
  });
  if (!(document as { fonts?: unknown }).fonts) Object.defineProperty(document, 'fonts', { value: { add() {} }, configurable: true });
  const container = document.createElement('div');
  document.body.appendChild(container);
  const viewer = new DocxViewer(container, { fonts, editable: true });
  const body = Array.from({ length: 40 }, (_, i) => p(`Paragraph number ${i + 1} with some words in it.`)).join('');
  await viewer.load(await buildDocx({ body }));
  return { viewer, editing: viewer.editing! };
}

afterEach(() => {
  document.body.replaceChildren();
  vi.unstubAllGlobals();
});

const hasList = (viewer: DocxViewer) => viewer.document!.body().blocks.some((b) => b.kind === 'paragraph' && !!b.numbering);

describe('InlineEditing.transaction', () => {
  it('redraws without scrolling to the caret (the caret did not move)', async () => {
    const { viewer, editing } = await editableViewer();
    editing.select({ para: 0, char: 0 });
    const root = viewer.root;
    // happy-dom does no layout: give the scroll container a size (rows rebuilt for the new
    // comment column keep the scroll position as a ratio of scrollHeight).
    Object.defineProperty(root, 'clientHeight', { value: 400, configurable: true });
    Object.defineProperty(root, 'scrollHeight', { value: 20000, configurable: true });
    root.scrollTop = 5000;
    const before = root.scrollTop;
    expect(before).toBe(5000);
    await editing.transaction('Add comment', (editor) => editor.addComment(editor.find('Paragraph number 30')[0]!, { author: 'A', text: 'x' }));
    expect(readComments(viewer.document!)).toHaveLength(1);
    expect(root.scrollTop).toBe(before);
    expect(editing.selection).toEqual({ anchor: { para: 0, char: 0 }, focus: { para: 0, char: 0 } });
  });

  it('runs nested calls made before the first await inside the outer one: no deadlock, one undo step that undoes it all', async () => {
    const { viewer, editing } = await editableViewer();
    editing.select({ para: 0, char: 0 }, { para: 0, char: 9 });
    const errors = vi.fn();
    viewer.on('editerror', errors);
    const done = await Promise.race([
      editing.transaction('Outer', async (editor) => {
        await Promise.all([
          editing.setList('bullet'),
          editing.transaction('Inner', (inner) => inner.addComment(inner.find('Paragraph number 3 ')[0]!, { author: 'A', text: 'inner' })),
        ]);
        // After an await, changes go through the editor the transaction gives.
        await editor.addComment(editor.find('Paragraph number 2 ')[0]!, { author: 'A', text: 'outer' });
        return 'done';
      }),
      new Promise((_, fail) => setTimeout(() => fail(new Error('deadlock')), 1000)),
    ]);
    expect(done).toBe('done');
    expect(errors).not.toHaveBeenCalled();
    expect(readComments(viewer.document!).map((c) => c.text).sort()).toEqual(['inner', 'outer']);
    expect(hasList(viewer)).toBe(true);
    expect(editing.canUndo).toBe(true);
    editing.undo();
    expect(readComments(viewer.document!)).toHaveLength(0);
    expect(hasList(viewer)).toBe(false);
    expect(editing.canUndo).toBe(false);
  });

  it('an edit requested while a transaction awaits waits for it and is its own undo step', async () => {
    const { viewer, editing } = await editableViewer();
    editing.select({ para: 0, char: 0 }, { para: 0, char: 9 });
    let started!: () => void;
    const running = new Promise<void>((r) => (started = r));
    let open!: () => void;
    const gate = new Promise<void>((r) => (open = r));
    const comment = editing.transaction('Add comment', async (editor) => {
      started();
      await gate;
      await editor.addComment(editor.find('Paragraph number 2 ')[0]!, { author: 'A', text: 'a' });
    });
    await running;
    // A toolbar click while the comment is still being added.
    const list = editing.setList('bullet');
    await new Promise((r) => setTimeout(r, 5));
    expect(hasList(viewer)).toBe(false);
    open();
    await comment;
    await list;
    expect(readComments(viewer.document!)).toHaveLength(1);
    expect(hasList(viewer)).toBe(true);
    editing.undo();
    expect(hasList(viewer)).toBe(false);
    expect(readComments(viewer.document!)).toHaveLength(1);
    editing.undo();
    expect(readComments(viewer.document!)).toHaveLength(0);
    expect(editing.canUndo).toBe(false);
  });

  it.each(['succeeds', 'fails'])('a queued edit acts on the selection made when it was requested (the one ahead %s)', async (outcome) => {
    const { viewer, editing } = await editableViewer();
    editing.select({ para: 0, char: 0 });
    let started!: () => void;
    const running = new Promise<void>((r) => (started = r));
    let open!: () => void;
    const gate = new Promise<void>((r) => (open = r));
    const first = editing.transaction('First', async () => {
      started();
      await gate;
      if (outcome === 'fails') throw new Error('refused');
    });
    const firstDone = first.then(
      () => 'ok',
      () => 'failed',
    );
    await running;
    editing.select({ para: 5, char: 0 });
    const list = editing.setList('bullet');
    open();
    expect(await firstDone).toBe(outcome === 'fails' ? 'failed' : 'ok');
    await list;
    const listed = viewer.document!.body().blocks.flatMap((b, i) => (b.kind === 'paragraph' && b.numbering ? [i] : []));
    expect(listed).toEqual([5]);
    expect(editing.selection).toEqual({ anchor: { para: 5, char: 0 }, focus: { para: 5, char: 0 } });
  });

  it('a helper called after an await inside a transaction gives up after the wait limit and reports it', async () => {
    const limit = UndoHistory.queueTimeout;
    UndoHistory.queueTimeout = 50;
    try {
      const { viewer, editing } = await editableViewer();
      editing.select({ para: 0, char: 0 }, { para: 0, char: 9 });
      const errors = vi.fn();
      viewer.on('editerror', errors);
      const done = await Promise.race([
        editing.transaction('Outer', async (editor) => {
          await editor.addComment(editor.find('Paragraph number 2 ')[0]!, { author: 'A', text: 'outer' });
          // Too late to join the outer step: waits for it, then gives up instead of hanging.
          await editing.setList('bullet');
          return 'done';
        }),
        new Promise((_, fail) => setTimeout(() => fail(new Error('hang')), 2000)),
      ]);
      expect(done).toBe('done');
      expect(errors).toHaveBeenCalledTimes(1);
      const error = errors.mock.calls[0]![0].error as Error;
      expect(error).toBeInstanceOf(EditError);
      expect(error.message).toMatch(/Inside viewer\.editing\.transaction\(\), make changes through the editor it gives you/);
      expect(hasList(viewer)).toBe(false);
      expect(readComments(viewer.document!)).toHaveLength(1);
      // Not stuck: the next edit runs, and each is its own undo step.
      await editing.setList('bullet');
      expect(hasList(viewer)).toBe(true);
      editing.undo();
      expect(hasList(viewer)).toBe(false);
      expect(readComments(viewer.document!)).toHaveLength(1);
      editing.undo();
      expect(readComments(viewer.document!)).toHaveLength(0);
    } finally {
      UndoHistory.queueTimeout = limit;
    }
  });

  it('reports a failing nested transaction once', async () => {
    const { viewer, editing } = await editableViewer();
    editing.select({ para: 0, char: 0 });
    const errors = vi.fn();
    viewer.on('editerror', errors);
    await expect(
      editing.transaction('Outer', async () => {
        await editing.transaction('Inner', () => {
          throw new Error('refused');
        });
      }),
    ).rejects.toThrow('refused');
    expect(errors).toHaveBeenCalledTimes(1);
    expect(editing.canUndo).toBe(false);
  });
});
