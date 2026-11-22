import { describe, expect, it } from 'vitest';
import { DocxEditor, type DocxEditor as Editor } from '../src/api/index.js';
import { blocksToText, DocxDocument } from '../src/model/index.js';
import { AccessPolicy, PermissionError, setProtection } from '../src/security/index.js';
import { findAll, serializeXml } from '../src/xml/index.js';
import { buildDocx, p } from './helpers/docx.js';

const BODY = p('Payment is due within 30 days.') + p('Second paragraph.') + p('Third paragraph.');

async function open(options: { body?: string; settings?: string; author?: string; track?: boolean } = {}) {
  const doc = await DocxDocument.open(await buildDocx({ body: options.body ?? BODY, settings: options.settings ?? '' }));
  const editor = new DocxEditor(doc, undefined, { trackChanges: options.track ?? true, author: options.author ?? 'Rina' });
  return { doc, editor };
}

const xml = (doc: DocxDocument) => serializeXml(doc.documentXml);
const finalText = (doc: DocxDocument) => blocksToText(doc.body().blocks);

async function markupText(doc: DocxDocument): Promise<string> {
  const shown = await DocxDocument.open(await doc.save(), { showDeletions: true });
  return blocksToText(shown.body().blocks);
}

function range(editor: Editor, text: string) {
  const m = editor.find(text)[0];
  if (!m) throw new Error(`no "${text}"`);
  return m;
}

describe('recording tracked changes', () => {
  it('records a replacement as a deletion followed by an insertion', async () => {
    const { doc, editor } = await open();
    editor.replaceRange(range(editor, '30'), '45');
    expect(finalText(doc).split('\n')[0]).toBe('Payment is due within 45 days.');
    expect(await markupText(doc)).toContain('Payment is due within 3045 days.');
    const out = xml(doc);
    expect(out).toMatch(/<w:del w:id="\d+" w:author="Rina" w:date="[^"]+Z"><w:r><w:delText>30<\/w:delText><\/w:r><\/w:del><w:ins w:id="\d+" w:author="Rina" w:date="[^"]+"><w:r><w:t>45<\/w:t><\/w:r><\/w:ins>/);
    expect(editor.revisions().map((r) => [r.kind, r.author, r.text])).toEqual([
      ['delete', 'Rina', '30'],
      ['insert', 'Rina', '45'],
    ]);
  });

  it('keeps typing in the same insertion and removes text the author inserted', async () => {
    const { doc, editor } = await open();
    const after = editor.insertText(range(editor, 'Second').start, 'Very ');
    editor.insertText(after, 'very ');
    expect(findAll(doc.documentXml.root, 'w:ins')).toHaveLength(1);
    expect(finalText(doc)).toContain('Very very Second paragraph.');
    // Deleting your own insertion takes it out instead of marking it deleted.
    editor.deleteRange(range(editor, 'very '));
    expect(findAll(doc.documentXml.root, 'w:del')).toHaveLength(0);
    expect(finalText(doc)).toContain('Very Second paragraph.');
  });

  it("marks another author's insertion as deleted instead of removing it", async () => {
    const { doc, editor } = await open({ author: 'Rina' });
    editor.insertText(range(editor, 'Second').start, 'New ');
    const budi = new DocxEditor(doc, undefined, { trackChanges: true, author: 'Budi' });
    budi.deleteRange(range(budi, 'New '));
    expect(budi.revisions().map((r) => `${r.kind}:${r.author}`)).toEqual(['insert:Rina', 'delete:Budi']);
    expect(finalText(doc)).toContain('\nSecond paragraph.');
  });

  it('records formatting changes, and rejecting them restores the old formatting', async () => {
    const { doc, editor } = await open();
    editor.formatRange(range(editor, 'Payment'), { bold: true });
    expect(editor.revisions().map((r) => r.kind)).toEqual(['format']);
    expect(xml(doc)).toContain('<w:rPrChange');
    expect(editor.rejectRevisions()).toBe(1);
    expect(xml(doc)).not.toContain('w:b/>');
    expect(xml(doc)).not.toContain('w:rPrChange');
  });

  it('records paragraph deletions and joins them away on accept', async () => {
    const { doc, editor } = await open();
    editor.deleteParagraph(editor.paragraphs()[1]!.el);
    expect(finalText(doc)).toBe('Payment is due within 30 days.\nThird paragraph.');
    expect(editor.revisions().map((r) => r.kind)).toEqual(['delete', 'paragraph-delete']);
    expect(editor.acceptRevisions()).toBe(2);
    expect(findAll(doc.documentXml.root, 'w:p')).toHaveLength(2);
    expect(finalText(doc)).toBe('Payment is due within 30 days.\nThird paragraph.');
  });

  it('rejecting a deleted paragraph brings it back', async () => {
    const { doc, editor } = await open();
    editor.deleteParagraph(editor.paragraphs()[1]!.el);
    editor.rejectRevisions();
    expect(finalText(doc)).toBe('Payment is due within 30 days.\nSecond paragraph.\nThird paragraph.');
    expect(xml(doc)).not.toMatch(/w:del|w:delText/);
  });

  it('records Enter as an inserted paragraph mark, and rejecting it joins the halves', async () => {
    const { doc, editor } = await open();
    editor.splitParagraph(range(editor, 'within').start);
    expect(finalText(doc).split('\n').slice(0, 2)).toEqual(['Payment is due ', 'within 30 days.']);
    expect(editor.revisions().map((r) => r.kind)).toEqual(['paragraph-insert']);
    editor.rejectRevisions();
    expect(finalText(doc).split('\n')[0]).toBe('Payment is due within 30 days.');
  });

  it('records paragraph style changes', async () => {
    const styles = '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style><w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/></w:style>';
    const doc = await DocxDocument.open(await buildDocx({ body: BODY, styles }));
    const editor = new DocxEditor(doc, undefined, { trackChanges: true, author: 'Rina' });
    editor.setParagraphStyle(editor.paragraphs()[0]!.el, 'Heading1');
    expect(editor.revisions().map((r) => r.kind)).toEqual(['paragraph-format']);
    editor.rejectRevisions();
    expect(xml(doc)).not.toContain('Heading1');
  });

  it('accepts or rejects by author and by id', async () => {
    const { doc, editor } = await open({ author: 'Rina' });
    editor.replaceRange(range(editor, 'Second'), 'Next');
    const budi = new DocxEditor(doc, undefined, { trackChanges: true, author: 'Budi' });
    budi.replaceRange(range(budi, 'Third'), 'Last');
    expect(editor.rejectRevisions({ author: 'Budi' })).toBe(2);
    expect(finalText(doc)).toContain('Third paragraph.');
    const [del] = editor.revisions();
    expect(editor.acceptRevisions({ ids: [del!.id] })).toBe(1);
    expect(editor.revisions().map((r) => r.kind)).toEqual(['insert']);
    expect(editor.acceptRevisions()).toBe(1);
    expect(finalText(doc)).toContain('Next paragraph.');
    expect(xml(doc)).not.toMatch(/w:ins|w:del/);
  });

  it('survives saving', async () => {
    const { doc, editor } = await open();
    editor.replaceRange(range(editor, '30'), '45');
    const again = await DocxDocument.open(await doc.save());
    const reopened = new DocxEditor(again);
    expect(reopened.revisions().map((r) => r.kind)).toEqual(['delete', 'insert']);
    expect(blocksToText(again.body().blocks).split('\n')[0]).toBe('Payment is due within 45 days.');
  });
});

describe("Word's Track Changes setting", () => {
  it('tracks by default when the document has Track Changes on', async () => {
    const { doc, editor } = await open({ settings: '<w:trackRevisions/>', track: undefined as never });
    const plain = new DocxEditor(doc);
    expect(plain.trackChanges).toBe(true);
    plain.replaceRange(range(plain, '30'), '45');
    expect(plain.revisions()).toHaveLength(2);
    expect(plain.revisions()[0]!.author).toBe('docxkeun');
    await editor.setTrackChangesSetting(false);
    expect(new DocxEditor(doc).trackChanges).toBe(false);
  });

  it('can turn the setting on', async () => {
    const doc = await DocxDocument.open(await buildDocx({ body: BODY }));
    const editor = new DocxEditor(doc);
    expect(editor.trackChanges).toBe(false);
    await editor.setTrackChangesSetting(true);
    expect(editor.trackChanges).toBe(true);
    expect(serializeXml(doc.xmlSync(doc.parts.settings)!)).toContain('<w:trackRevisions/>');
  });

  it('lets editors change a document protected for tracked changes, tracked, but not review it', async () => {
    const doc = await DocxDocument.open(await buildDocx({ body: BODY, settings: '' }));
    await setProtection(doc, 'trackedChanges');
    const editor = new DocxEditor(doc, new AccessPolicy(doc, { user: { id: 'u1', name: 'Sari' }, role: 'editor' }));
    expect(editor.trackChanges).toBe(true);
    editor.replaceRange(range(editor, '30'), '45');
    expect(editor.revisions().map((r) => r.author)).toEqual(['Sari', 'Sari']);
    expect(() => editor.acceptRevisions()).toThrow(PermissionError);
    const owner = new DocxEditor(doc, new AccessPolicy(doc, { user: { id: 'o' }, role: 'owner' }));
    expect(owner.acceptRevisions()).toBe(2);
  });
});

describe('reading tracked changes made in Word', () => {
  it('resolves moves and their range markers', async () => {
    const body =
      '<w:p><w:moveFromRangeStart w:id="1" w:name="move1"/><w:moveFrom w:id="2" w:author="Ana" w:date="2026-01-01T00:00:00Z"><w:r><w:t>Moved text. </w:t></w:r></w:moveFrom><w:moveFromRangeEnd w:id="1"/><w:r><w:t>Stays.</w:t></w:r></w:p>' +
      '<w:p><w:r><w:t>Target: </w:t></w:r><w:moveToRangeStart w:id="3" w:name="move1"/><w:moveTo w:id="4" w:author="Ana" w:date="2026-01-01T00:00:00Z"><w:r><w:t>Moved text. </w:t></w:r></w:moveTo><w:moveToRangeEnd w:id="3"/></w:p>';
    const doc = await DocxDocument.open(await buildDocx({ body }));
    const editor = new DocxEditor(doc);
    expect(editor.revisions().map((r) => r.kind)).toEqual(['move-from', 'move-to']);
    expect(blocksToText(doc.body().blocks)).toBe('Stays.\nTarget: Moved text. ');
    editor.acceptRevisions();
    expect(serializeXml(doc.documentXml)).not.toMatch(/moveFrom|moveTo/);
    expect(blocksToText(doc.body().blocks)).toBe('Stays.\nTarget: Moved text. ');
  });
});
