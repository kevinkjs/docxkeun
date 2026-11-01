import { describe, expect, it } from 'vitest';
import { DocxEditor, LockedRegionError } from '../src/api/index.js';
import { layoutDocument } from '../src/layout/index.js';
import { blocksToText, DocxDocument } from '../src/model/index.js';
import { AccessPolicy, PermissionError } from '../src/security/index.js';
import { findAll, serializeXml } from '../src/xml/index.js';
import { buildDocx, p } from './helpers/docx.js';
import { testRegistry } from './helpers/fonts.js';

async function open(body: string) {
  const doc = await DocxDocument.open(await buildDocx({ body }));
  return { doc, editor: new DocxEditor(doc) };
}

const text = (doc: DocxDocument) => blocksToText(doc.body().blocks);

describe('locked regions inside a paragraph', () => {
  it('blocks edits to the locked text and allows everything around it', async () => {
    const { doc, editor } = await open(p('The royalty is 3.75 percent of revenue.'));
    const id = editor.lockRange(editor.find('3.75 percent')[0]!, { title: 'Agreed rate' });
    expect(editor.lockedRegions()).toMatchObject([{ id, mode: 'content', title: 'Agreed rate', text: '3.75 percent', block: false }]);
    expect(serializeXml(doc.documentXml)).toContain('<w:lock w:val="contentLocked"/>');

    const inside = editor.find('75')[0]!;
    expect(() => editor.insertText(inside.start, 'x')).toThrow(LockedRegionError);
    expect(() => editor.insertText(inside.start, 'x')).toThrow(/"Agreed rate" is locked/);
    expect(() => editor.replaceRange(editor.find('royalty is 3.75')[0]!, 'fee')).toThrow(LockedRegionError);
    expect(() => editor.deleteRange({ start: editor.find('is')[0]!.start, end: editor.find('of')[0]!.end })).toThrow(LockedRegionError);
    expect(() => editor.formatRange(inside, { bold: true })).toThrow(LockedRegionError);
    await expect(editor.addHyperlink(inside, { url: 'https://example.com' })).rejects.toThrow(LockedRegionError);
    expect(editor.isLocked(inside)).toBe(true);

    editor.replaceRange(editor.find('royalty')[0]!, 'fee');
    expect(text(doc)).toBe('The fee is 3.75 percent of revenue.');
    // Comments on locked text are fine: they do not change it.
    await editor.addComment(inside, { author: 'Legal', text: 'Fixed by the board' });
  });

  it('types just outside the region at its edges', async () => {
    const { doc, editor } = await open('<w:p><w:r><w:t xml:space="preserve">Rate: </w:t></w:r><w:r><w:t>3.75%</w:t></w:r></w:p>');
    editor.lockRange(editor.find('3.75%')[0]!);
    const [, locked] = findAll(doc.documentXml.root, 'w:t');
    const after = editor.insertText({ node: locked!, offset: 5 }, ' per year');
    editor.insertText(after, '!');
    editor.insertText({ node: locked!, offset: 0 }, '~');
    expect(text(doc)).toBe('Rate: ~3.75% per year!');
    expect(editor.lockedRegions()[0]!.text).toBe('3.75%');
  });

  it('survives saving and can be unlocked', async () => {
    const { doc, editor } = await open(p('Fixed clause here.'));
    const id = editor.lockRange(editor.find('Fixed clause')[0]!, { title: 'Clause 7' });
    const reopened = await DocxDocument.open(await doc.save());
    const e2 = new DocxEditor(reopened);
    expect(e2.lockedRegions().map((r) => r.id)).toEqual([id]);
    expect(() => e2.replaceRange(e2.find('Fixed')[0]!, 'Changed')).toThrow(LockedRegionError);
    e2.unlock(id);
    expect(serializeXml(reopened.documentXml)).not.toContain('w:sdt');
    e2.replaceRange(e2.find('Fixed')[0]!, 'Changed');
    expect(text(reopened)).toBe('Changed clause here.');
  });
});

describe('locked paragraphs', () => {
  it('locks whole paragraphs and adds paragraphs outside them on Enter', async () => {
    const { doc, editor } = await open(p('Intro') + p('Locked one') + p('Locked two') + p('Outro'));
    const paras = editor.paragraphs();
    const id = editor.lockParagraphs(paras[1]!.el, paras[2]!.el, { title: 'Signature block', mode: 'contentAndControl' });
    expect(editor.lockedRegions()).toMatchObject([{ id, mode: 'contentAndControl', block: true, text: 'Locked one\nLocked two' }]);

    const locked = editor.paragraphs()[1]!.el;
    expect(() => editor.deleteParagraph(locked)).toThrow(LockedRegionError);
    expect(() => editor.mergeParagraphs(editor.paragraphs()[0]!.el, locked)).toThrow(LockedRegionError);
    expect(() => editor.splitParagraph(editor.find('one')[0]!.start)).toThrow(LockedRegionError);
    expect(() => editor.insertParagraph({ after: locked }, 'nope')).toThrow(LockedRegionError);
    // Deleting a range that contains the whole region is refused too.
    expect(() => editor.deleteRange({ start: editor.find('Intro')[0]!.start, end: editor.find('Outro')[0]!.end })).toThrow(LockedRegionError);

    // Enter at the end of the block (and at its start) creates a paragraph outside it.
    const after = editor.splitParagraph(editor.find('two')[0]!.end);
    editor.insertText({ node: after, offset: 0 }, 'Added after');
    const before = editor.splitParagraph(editor.find('Locked one')[0]!.start);
    editor.insertText({ node: before, offset: 0 }, 'Added before');
    expect(text(doc)).toBe('Intro\nAdded before\nLocked one\nLocked two\nAdded after\nOutro');
    expect(editor.lockedRegions()[0]!.text).toBe('Locked one\nLocked two');
  });

  it('honours locks Word wrote, and unlocking keeps Word controls', async () => {
    const body =
      '<w:sdt><w:sdtPr><w:alias w:val="Party"/><w:tag w:val="party"/><w:id w:val="-1234"/><w:lock w:val="sdtContentLocked"/><w:text/></w:sdtPr><w:sdtContent><w:p><w:r><w:t>PT Contoh</w:t></w:r></w:p></w:sdtContent></w:sdt>' +
      p('Free text');
    const { doc, editor } = await open(body);
    expect(() => editor.replaceRange(editor.find('Contoh')[0]!, 'X')).toThrow(/"Party" is locked/);
    editor.unlock('-1234');
    const xml = serializeXml(doc.documentXml);
    expect(xml).toContain('<w:tag w:val="party"/>');
    expect(xml).not.toContain('w:lock');
    editor.replaceRange(editor.find('Contoh')[0]!, 'Baru');
    expect(text(doc)).toBe('PT Baru\nFree text');
  });
});

describe('locks with templates, permissions and layout', () => {
  it('leaves locked placeholders and controls alone unless asked', async () => {
    const body =
      p('Name: {{name}}, rate: {{rate}}') +
      '<w:p><w:sdt><w:sdtPr><w:tag w:val="company"/><w:lock w:val="contentLocked"/></w:sdtPr><w:sdtContent><w:r><w:t>Company</w:t></w:r></w:sdtContent></w:sdt></w:p><w:p/>';
    const { doc, editor } = await open(body);
    editor.lockRange(editor.find('{{rate}}')[0]!);
    expect(editor.fillTemplate({ name: 'Rina', rate: '4%', company: 'PT A' }).sort()).toEqual(['name']);
    expect(text(doc)).toBe('Name: Rina, rate: {{rate}}\nCompany\n');
    expect(editor.fillTemplate({ rate: '4%', company: 'PT A' }, undefined, { includeLocked: true }).sort()).toEqual(['company', 'rate']);
    expect(text(doc)).toBe('Name: Rina, rate: 4%\nPT A\n');
    // The lock itself is kept.
    expect(editor.lockedRegions()).toHaveLength(2);
  });

  it('only owners can lock or unlock when an access policy is used', async () => {
    const { doc } = await open(p('Clause text'));
    const asEditor = new DocxEditor(doc, new AccessPolicy(doc, { user: { id: 'e' }, role: 'editor' }));
    expect(() => asEditor.lockRange(asEditor.find('Clause')[0]!)).toThrow(PermissionError);
    const asOwner = new DocxEditor(doc, new AccessPolicy(doc, { user: { id: 'o' }, role: 'owner' }));
    const id = asOwner.lockRange(asOwner.find('Clause')[0]!);
    expect(() => asEditor.unlock(id)).toThrow(PermissionError);
    expect(() => asEditor.replaceRange(asEditor.find('Clause')[0]!, 'x')).toThrow(LockedRegionError);
  });

  it('marks locked text in the layout', async () => {
    const { doc, editor } = await open(p('Open text, locked text, open again.'));
    editor.lockRange(editor.find('locked text')[0]!, { title: 'Fixed' });
    const page = layoutDocument(doc, { fonts: testRegistry() }).pages[0]!;
    expect(page.locks).toHaveLength(1);
    expect(page.locks[0]).toMatchObject({ title: 'Fixed' });
    expect(page.locks[0]!.w).toBeGreaterThan(30);
    expect(page.items.filter((i) => i.type === 'text').map((i) => (i.type === 'text' ? i.text : '')).join('')).toBe('Open text, locked text, open again.');
  });
});
