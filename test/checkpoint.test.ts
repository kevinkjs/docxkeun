import { describe, expect, it } from 'vitest';
import { DocxEditor } from '../src/api/index.js';
import { readComments } from '../src/comments/read.js';
import { DocxDocument, blocksToText } from '../src/model/index.js';
import { buildDocx, p } from './helpers/docx.js';

async function open() {
  const doc = await DocxDocument.open(await buildDocx({ body: p('Hello world') + p('Second paragraph') }));
  return { doc, editor: new DocxEditor(doc) };
}

const text = (doc: DocxDocument) => blocksToText(doc.body().blocks);

describe('document checkpoints', () => {
  it('restores comments, links and headers added after the checkpoint', async () => {
    const { doc, editor } = await open();
    const names = doc.pkg.partNames();
    const before = text(doc);
    const cp = doc.checkpoint();
    await editor.addComment(editor.find('Hello')[0]!, { author: 'A', text: 'note' });
    await editor.addHyperlink(editor.find('world')[0]!, { url: 'https://example.com' });
    await editor.setHeader('H');
    expect(readComments(doc)).toHaveLength(1);
    doc.restore(cp);
    expect(readComments(doc)).toHaveLength(0);
    expect(editor.hyperlinks()).toHaveLength(0);
    expect(doc.storyParts()).toHaveLength(0);
    expect(doc.pkg.partNames()).toEqual(names);
    expect(text(doc)).toBe(before);
  });

  it('restores a deleted part', async () => {
    const { doc, editor } = await open();
    await editor.setHeader('H');
    const cp = doc.checkpoint();
    await editor.removeHeader();
    expect(editor.headersAndFooters()).toHaveLength(0);
    doc.restore(cp);
    const [info] = editor.headersAndFooters();
    expect(info).toBeDefined();
    const headers = doc.storyParts().filter((s) => s.kind === 'header');
    expect(headers).toHaveLength(1);
    const hdr = new DocxEditor(doc, undefined, { part: headers[0]!.part });
    expect(hdr.find('H')).toHaveLength(1);
    expect(doc.pkg.has(headers[0]!.part)).toBe(true);
    expect(doc.pkg.partNames()).toContain(headers[0]!.part);
    const reopened = await DocxDocument.open(await doc.save());
    expect(reopened.storyParts().filter((s) => s.kind === 'header')).toHaveLength(1);
  });

  it('checkpoints are independent of later edits', async () => {
    const { doc, editor } = await open();
    const cp = doc.checkpoint();
    const names = doc.pkg.partNames();
    for (let i = 0; i < 2; i++) {
      await editor.addComment(editor.find('Hello')[0]!, { author: 'A', text: 'n' });
      editor.replaceRange(editor.find('Second')[0]!, 'Third');
      doc.restore(cp);
      expect(readComments(doc)).toHaveLength(0);
      expect(text(doc)).toBe('Hello world\nSecond paragraph');
      expect(doc.pkg.partNames()).toEqual(names);
    }
  });

  it('save after restore reads the restored state', async () => {
    const { doc, editor } = await open();
    const cp = doc.checkpoint();
    await editor.addComment(editor.find('Hello')[0]!, { author: 'A', text: 'n' });
    await editor.setHeader('H');
    doc.restore(cp);
    const reopened = await DocxDocument.open(await doc.save());
    expect(readComments(reopened)).toHaveLength(0);
    expect(reopened.storyParts()).toHaveLength(0);
    expect(text(reopened)).toBe(text(doc));
  });

  it('checkpoints share unchanged parts', async () => {
    const { doc, editor } = await open();
    await editor.setHeader('H');
    const a = doc.checkpoint();
    const b = doc.checkpoint();
    expect(a.pkg.parts.length).toBeGreaterThan(2);
    a.pkg.parts.forEach(([key, part], i) => {
      expect(b.pkg.parts[i]![0]).toBe(key);
      expect(b.pkg.parts[i]![1]).toBe(part);
    });
    editor.replaceRange(editor.find('Second')[0]!, 'Third');
    const c = doc.checkpoint();
    const changed = c.pkg.parts.filter(([, part], i) => part !== a.pkg.parts[i]![1]).map(([key]) => key);
    expect(changed).toEqual(['word/document.xml']);
    // Restoring must not hand a frozen copy to live state, nor let a later checkpoint reuse stale copies.
    doc.restore(a);
    editor.replaceRange(editor.find('Second')[0]!, 'Fourth');
    const d = doc.checkpoint();
    expect(text(doc)).toContain('Fourth');
    doc.restore(a);
    expect(text(doc)).toContain('Second');
    doc.restore(d);
    expect(text(doc)).toContain('Fourth');
    expect(a.pkg.parts.every(([, part]) => b.pkg.parts.some(([, q]) => q === part))).toBe(true);
  });
});
