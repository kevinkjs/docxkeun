import { describe, expect, it } from 'vitest';
import { paragraphsOf } from '../src/api/index.js';
import { layoutDocument } from '../src/layout/index.js';
import { DocxDocument } from '../src/model/index.js';
import { diffParagraphs, diffText } from '../src/render/activity.js';
import { describeSelection, selectionInfoFor } from '../src/api/selection.js';
import { buildDocx, p } from './helpers/docx.js';
import { testRegistry } from './helpers/fonts.js';

describe('what changed, for the reveal animation', () => {
  it('finds replaced words and cuts inside a paragraph', () => {
    const before = 'Payment is due within 30 days of the invoice date.';
    const after = 'Payment is due within 45 days of the invoice date.';
    const d = diffText(before, after);
    expect(d.ranges.map(([a, b]) => after.slice(a, b))).toEqual(['45']);
    const cut = diffText('The Supplier shall deliver on time and in full.', 'The Supplier shall deliver in full.');
    expect(cut.ranges).toEqual([]);
    expect(cut.cuts).toEqual(['The Supplier shall deliver '.length]);
  });

  it('finds several edits in one paragraph', () => {
    const after = 'ACME Limited pays the Client within 45 days.';
    const d = diffText('ACME Ltd. pays the Client within 30 days.', after);
    expect(d.ranges.map(([a, b]) => after.slice(a, b))).toEqual(['Limited', '45']);
  });

  it('tells inserted, changed and deleted paragraphs apart', () => {
    const before = ['Title', 'First clause.', 'Second clause.', 'Old clause to remove entirely here.', 'Last.'];
    const after = ['Title', 'First clause, amended.', 'A brand new paragraph about something else.', 'Second clause.', 'Last.'];
    const d = diffParagraphs(before, after);
    expect(d.changes.map((c) => [c.para, c.kind])).toEqual([
      [1, 'changed'],
      [2, 'inserted'],
    ]);
    expect(d.changes[0]!.before).toBe('First clause.');
    expect(d.deleted).toEqual([{ at: 4, text: 'Old clause to remove entirely here.' }]);
    // Unchanged paragraphs, old index → new index (to compare their formatting).
    expect(d.same).toEqual([
      [0, 0],
      [2, 3],
      [4, 4],
    ]);
  });

  it('reports nothing for an unchanged document', () => {
    expect(diffParagraphs(['a', 'b'], ['a', 'b'])).toEqual({ changes: [], deleted: [], same: [[0, 0], [1, 1]] });
  });
});

describe('selection info for AI assistants', () => {
  async function doc() {
    const styles = '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style><w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/></w:style>';
    const cell = (t: string) => `<w:tc>${p(t)}</w:tc>`;
    const body =
      p('Terms', '<w:pStyle w:val="Heading1"/>') +
      p('Payment is due within 30 days of the invoice date.') +
      `<w:tbl><w:tblGrid><w:gridCol w:w="3000"/><w:gridCol w:w="3000"/></w:tblGrid><w:tr>${cell('Rate')}${cell('3.75%')}</w:tr></w:tbl>` +
      p('Signed by the parties.');
    return DocxDocument.open(await buildDocx({ body, styles }));
  }

  it('numbers paragraphs as read_document does (table cells included) and finds the pages', async () => {
    const d = await doc();
    const paras = paragraphsOf(d.body().blocks);
    const layout = layoutDocument(d, { fonts: testRegistry() });
    const info = selectionInfoFor(d, paras, { para: 1, char: 22 }, { para: 1, char: 29 }, { layout })!;
    expect(info).toMatchObject({ story: 'body', collapsed: false, text: '30 days', start: { paragraph: 2, offset: 22 }, end: { paragraph: 2, offset: 29 }, pages: [1] });
    expect(describeSelection(info, { document: 'contract.docx' })).toBe(
      'Selected in contract.docx: paragraph 2, page 1, 7 characters.\n[2] Payment is due within «30 days» of the invoice date.',
    );
    // Backwards selections are put in order; the cell "3.75%" is paragraph 4.
    const across = selectionInfoFor(d, paras, { para: 4, char: 6 }, { para: 3, char: 0 })!;
    expect(across.start).toEqual({ paragraph: 4, offset: 0 });
    expect(across.text).toBe('3.75%\nSigned');
    expect(describeSelection(across)).toContain('Selected: paragraphs 4–5, 12 characters.');
  });

  it('describes a caret and shortens long context', async () => {
    const d = await doc();
    const paras = paragraphsOf(d.body().blocks);
    const caret = selectionInfoFor(d, paras, { para: 0, char: 5 }, { para: 0, char: 5 })!;
    expect(describeSelection(caret)).toBe('Cursor: paragraph 1, nothing selected.\n[1] (heading 1) Terms‸');
    const short = describeSelection(selectionInfoFor(d, paras, { para: 1, char: 30 }, { para: 1, char: 32 })!, { context: 8 });
    expect(short).toContain('[2] …30 days «of» the inv…');
  });
});
