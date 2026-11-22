import { describe, expect, it, vi } from 'vitest';
import { DocxEditor, EditError } from '../src/api/index.js';
import { readComments } from '../src/comments/read.js';
import { DocxDocument, blocksToText } from '../src/model/index.js';
import { cloneElement } from '../src/xml/index.js';
import { UndoHistory, type UndoEntry } from '../src/render/undo.js';
import { AccessPolicy, PermissionError } from '../src/security/index.js';
import { buildDocx, p } from './helpers/docx.js';

const at = { para: 0, char: 0 };
const body = { kind: 'body' } as const;

async function open() {
  const doc = await DocxDocument.open(await buildDocx({ body: p('Hello world') + p('Second paragraph') }));
  return { doc, editor: new DocxEditor(doc) };
}

const text = (doc: DocxDocument) => blocksToText(doc.body().blocks);

function storyEntry(doc: DocxDocument): UndoEntry {
  return { kind: 'story', part: doc.mainPart, story: body, root: cloneElement(doc.documentXml.root), anchor: at, focus: at };
}
function packageEntry(doc: DocxDocument): UndoEntry {
  return { kind: 'package', checkpoint: doc.checkpoint(), story: body, anchor: at, focus: at, label: 'test' };
}
/** What the editor does with an entry it gets back from the history. */
function apply(doc: DocxDocument, e: UndoEntry): void {
  if (e.kind === 'package') doc.restore(e.checkpoint);
  else {
    doc.documentXml.root = e.root;
    doc.touch(e.part);
  }
}
const current = (doc: DocxDocument) => (next: UndoEntry) => (next.kind === 'package' ? packageEntry(doc) : storyEntry(doc));
/** A promise with its resolve and reject, to hold a transaction in the middle of its run. */
function deferred() {
  let resolve!: () => void;
  let reject!: (error: unknown) => void;
  const promise = new Promise<void>((res, rej) => {
    resolve = res;
    reject = rej;
  });
  return { promise, resolve, reject };
}
const settle = (ms = 5) => new Promise((r) => setTimeout(r, ms));
const undo = (h: UndoHistory, doc: DocxDocument) => {
  const e = h.undo(current(doc));
  if (e) apply(doc, e);
};
const redo = (h: UndoHistory, doc: DocxDocument) => {
  const e = h.redo(current(doc));
  if (e) apply(doc, e);
};

describe('undo history', () => {
  it('package entry undo and redo', async () => {
    const { doc, editor } = await open();
    const h = new UndoHistory();
    const cp = packageEntry(doc);
    await editor.addComment(editor.find('Hello')[0]!, { author: 'A', text: 'note' });
    h.push(cp);
    undo(h, doc);
    expect(readComments(doc)).toHaveLength(0);
    redo(h, doc);
    const comments = readComments(doc);
    expect(comments).toHaveLength(1);
    expect(comments[0]!.text).toBe('note');
  });

  it('failed transaction restores and records nothing', async () => {
    const { doc } = await open();
    const policy = new AccessPolicy(doc, { user: { id: 'v' }, role: 'viewer' });
    const editor = new DocxEditor(doc, policy);
    const h = new UndoHistory();
    const before = text(doc);
    await expect(
      h.transaction(doc, 'comment', () => editor.addComment(editor.find('Hello')[0]!, { author: 'A', text: 'x' }), { story: body, anchor: at, focus: at }),
    ).rejects.toBeInstanceOf(PermissionError);
    expect(text(doc)).toBe(before);
    expect(readComments(doc)).toHaveLength(0);
    expect(h.canUndo).toBe(false);
  });

  it('a successful transaction pushes one package entry', async () => {
    const { doc, editor } = await open();
    const h = new UndoHistory();
    const out = await h.transaction(doc, 'comment', async () => {
      await editor.addComment(editor.find('Hello')[0]!, { author: 'A', text: 'x' });
      return 7;
    }, { story: body, anchor: at, focus: at });
    expect(out).toBe(7);
    expect(h.canUndo).toBe(true);
    undo(h, doc);
    expect(readComments(doc)).toHaveLength(0);
    expect(h.canUndo).toBe(false);
    expect(h.canRedo).toBe(true);
  });

  it('mixed story and package undo', async () => {
    const { doc, editor } = await open();
    const h = new UndoHistory();
    const original = text(doc);
    h.push(storyEntry(doc));
    editor.insertText(editor.find('Hello')[0]!.start, 'A');
    const afterFirst = text(doc);
    h.push(packageEntry(doc));
    await editor.addComment(editor.find('Second')[0]!, { author: 'A', text: 'c' });
    h.push(storyEntry(doc));
    editor.insertText(editor.find('Second')[0]!.start, 'B');
    const final = text(doc);
    undo(h, doc);
    undo(h, doc);
    undo(h, doc);
    const reopened = await DocxDocument.open(await doc.save());
    expect(text(reopened)).toBe(original);
    expect(readComments(reopened)).toHaveLength(0);
    redo(h, doc);
    expect(text(doc)).toBe(afterFirst);
    redo(h, doc);
    expect(readComments(doc)).toHaveLength(1);
    redo(h, doc);
    expect(text(doc)).toBe(final);
    expect(readComments(doc)).toHaveLength(1);
  });

  it('cap at 200', async () => {
    const { doc } = await open();
    const h = new UndoHistory();
    for (let i = 0; i < 205; i++) h.push(storyEntry(doc));
    let n = 0;
    while (h.canUndo) {
      h.undo(current(doc));
      n++;
    }
    expect(n).toBe(200);
  });

  it('overlapping transactions run one at a time; a failing first one does not wipe the second', async () => {
    const { doc, editor } = await open();
    const h = new UndoHistory();
    const where = { story: body, anchor: at, focus: at };
    const tick = () => new Promise((r) => setTimeout(r, 5));
    const a = h.transaction(doc, 'a', async () => {
      editor.insertText(editor.find('Hello')[0]!.start, 'A');
      await tick();
      throw new Error('boom');
    }, where);
    const b = h.transaction(doc, 'b', async () => {
      await editor.addComment(editor.find('Second')[0]!, { author: 'A', text: 'b' });
    }, where);
    await expect(a).rejects.toThrow('boom');
    await b;
    expect(text(doc)).not.toContain('AHello');
    expect(readComments(doc)).toHaveLength(1);
    let n = 0;
    while (h.canUndo) {
      undo(h, doc);
      n++;
    }
    expect(n).toBe(1);
    expect(readComments(doc)).toHaveLength(0);
  });

  it('a transaction started inside another before its first await joins it: no deadlock, one undo step', async () => {
    const { doc, editor } = await open();
    const h = new UndoHistory();
    const where = { story: body, anchor: at, focus: at };
    const tick = () => new Promise((r) => setTimeout(r, 1));
    const out = await Promise.race([
      h.transaction(doc, 'outer', async () => {
        editor.insertText(editor.find('Hello')[0]!.start, 'A');
        const inner = h.transaction(doc, 'inner', async () => {
          await tick();
          await editor.addComment(editor.find('Second')[0]!, { author: 'A', text: 'nested' });
          return 'inner done';
        }, where);
        const inner2 = h.transaction(doc, 'inner 2', () => editor.insertText(editor.find('Second')[0]!.start, 'B'), where);
        await inner2;
        return await inner;
      }, where),
      new Promise((_, fail) => setTimeout(() => fail(new Error('deadlock')), 500)),
    ]);
    expect(out).toBe('inner done');
    expect(h.busy).toBe(false);
    expect(text(doc)).toContain('AHello');
    expect(text(doc)).toContain('BSecond');
    expect(readComments(doc)).toHaveLength(1);
    let n = 0;
    while (h.canUndo) {
      undo(h, doc);
      n++;
    }
    expect(n).toBe(1);
    expect(text(doc)).not.toContain('AHello');
    expect(text(doc)).not.toContain('BSecond');
    expect(readComments(doc)).toHaveLength(0);
  });

  it('a transaction requested while another awaits waits its turn; a failing first one does not wipe it', async () => {
    const { doc, editor } = await open();
    const h = new UndoHistory();
    const where = { story: body, anchor: at, focus: at };
    const started = deferred();
    const gate = deferred();
    const a = h.transaction(doc, 'a', async () => {
      editor.insertText(editor.find('Hello')[0]!.start, 'A');
      started.resolve();
      await gate.promise;
    }, where);
    await started.promise;
    let seenByB: string | undefined;
    const b = h.transaction(doc, 'b', async () => {
      seenByB = text(doc);
      await editor.addComment(editor.find('Second')[0]!, { author: 'A', text: 'b' });
    }, where);
    await settle();
    expect(seenByB).toBeUndefined();
    gate.reject(new Error('boom'));
    await expect(a).rejects.toThrow('boom');
    await b;
    // B ran on the document A's failure put back, and its change survived.
    expect(seenByB).not.toContain('AHello');
    expect(text(doc)).not.toContain('AHello');
    expect(readComments(doc)).toHaveLength(1);
    let n = 0;
    while (h.canUndo) {
      undo(h, doc);
      n++;
    }
    expect(n).toBe(1);
    expect(readComments(doc)).toHaveLength(0);
  });

  it('a transaction requested while another awaits is a separate undo step', async () => {
    const { doc, editor } = await open();
    const h = new UndoHistory();
    const where = { story: body, anchor: at, focus: at };
    const started = deferred();
    const gate = deferred();
    const a = h.transaction(doc, 'a', async () => {
      editor.insertText(editor.find('Hello')[0]!.start, 'A');
      started.resolve();
      await gate.promise;
    }, where);
    await started.promise;
    let bRan = false;
    const b = h.transaction(doc, 'b', async () => {
      bRan = true;
      await editor.addComment(editor.find('Second')[0]!, { author: 'A', text: 'b' });
    }, where);
    await settle();
    expect(bRan).toBe(false);
    gate.resolve();
    await a;
    await b;
    expect(text(doc)).toContain('AHello');
    expect(readComments(doc)).toHaveLength(1);
    undo(h, doc);
    expect(readComments(doc)).toHaveLength(0);
    expect(text(doc)).toContain('AHello');
    undo(h, doc);
    expect(text(doc)).not.toContain('AHello');
    expect(h.canUndo).toBe(false);
  });

  it('a transaction started inside another after an await gives up after the wait limit instead of hanging', async () => {
    const { doc, editor } = await open();
    const h = new UndoHistory();
    const where = { story: body, anchor: at, focus: at };
    const before = text(doc);
    const limit = UndoHistory.queueTimeout;
    UndoHistory.queueTimeout = 50;
    try {
      let innerRan = false;
      let seen: unknown;
      const outer = h.transaction(doc, 'outer', async () => {
        editor.insertText(editor.find('Hello')[0]!.start, 'A');
        await settle();
        try {
          await h.transaction(doc, 'inner', () => {
            innerRan = true;
            editor.insertText(editor.find('Second')[0]!.start, 'B');
          }, where);
        } catch (error) {
          seen = error;
          throw error;
        }
      }, where);
      await expect(outer).rejects.toThrow(/Another change is still running/);
      expect(seen).toBeInstanceOf(EditError);
      expect((seen as Error).message).toMatch(/make changes through the editor it gives you/);
      expect(innerRan).toBe(false);
      expect(text(doc)).toBe(before);
      expect(h.canUndo).toBe(false);
      expect(h.busy).toBe(false);
      await h.transaction(doc, 'next', () => editor.insertText(editor.find('Hello')[0]!.start, 'C'), where);
      expect(text(doc)).toContain('CHello');
      expect(h.canUndo).toBe(true);
    } finally {
      UndoHistory.queueTimeout = limit;
    }
  });

  it('a transaction that gave up waiting keeps its place: one queued after it still waits for the running one', async () => {
    const { doc, editor } = await open();
    const h = new UndoHistory();
    const where = { story: body, anchor: at, focus: at };
    const limit = UndoHistory.queueTimeout;
    UndoHistory.queueTimeout = 50;
    try {
      const started = deferred();
      const gate = deferred();
      const a = h.transaction(doc, 'a', async () => {
        started.resolve();
        await gate.promise;
        editor.insertText(editor.find('Hello')[0]!.start, 'A');
      }, where);
      await started.promise;
      const b = h.transaction(doc, 'b', () => 'never', where);
      const bGaveUp = expect(b).rejects.toBeInstanceOf(EditError);
      await settle(20);
      let cRan = false;
      const c = h.transaction(doc, 'c', () => {
        cRan = true;
        editor.insertText(editor.find('Second')[0]!.start, 'C');
      }, where);
      await bGaveUp;
      expect(cRan).toBe(false);
      gate.resolve();
      await a;
      await c;
      expect(text(doc)).toContain('AHello');
      expect(text(doc)).toContain('CSecond');
      undo(h, doc);
      expect(text(doc)).not.toContain('CSecond');
      expect(text(doc)).toContain('AHello');
    } finally {
      UndoHistory.queueTimeout = limit;
    }
  });

  it('a failing joined transaction rolls back the whole outer step when the outer lets it throw', async () => {
    const { doc, editor } = await open();
    const h = new UndoHistory();
    const where = { story: body, anchor: at, focus: at };
    const before = text(doc);
    await expect(
      h.transaction(doc, 'outer', async () => {
        editor.insertText(editor.find('Hello')[0]!.start, 'A');
        await h.transaction(doc, 'inner', () => {
          throw new Error('inner refused');
        }, where);
      }, where),
    ).rejects.toThrow('inner refused');
    expect(text(doc)).toBe(before);
    expect(h.canUndo).toBe(false);
    expect(h.busy).toBe(false);
  });

  it('a checkpoint that throws leaves the history usable', async () => {
    const { doc, editor } = await open();
    const h = new UndoHistory();
    const where = { story: body, anchor: at, focus: at };
    const real = doc.checkpoint.bind(doc);
    const spy = vi.spyOn(doc, 'checkpoint').mockImplementationOnce(() => {
      throw new Error('no snapshot');
    });
    await expect(h.transaction(doc, 'a', () => 1, where)).rejects.toThrow('no snapshot');
    expect(h.busy).toBe(false);
    spy.mockImplementation(real);
    await h.transaction(doc, 'b', () => editor.insertText(editor.find('Hello')[0]!.start, 'B'), where);
    expect(h.canUndo).toBe(true);
    undo(h, doc);
    expect(text(doc)).not.toContain('BHello');
  });

  it('undo and redo do nothing while a transaction is pending', async () => {
    const { doc, editor } = await open();
    const h = new UndoHistory();
    h.push(storyEntry(doc));
    let release!: () => void;
    const gate = new Promise<void>((r) => (release = r));
    const t = h.transaction(doc, 'c', async () => {
      await gate;
      await editor.addComment(editor.find('Hello')[0]!, { author: 'A', text: 'x' });
    }, { story: body, anchor: at, focus: at });
    expect(h.busy).toBe(true);
    expect(h.undo(current(doc))).toBeUndefined();
    expect(h.redo(current(doc))).toBeUndefined();
    release();
    await t;
    expect(h.busy).toBe(false);
    expect(h.canUndo).toBe(true);
    expect(h.undo(current(doc))).toBeDefined();
  });
});
