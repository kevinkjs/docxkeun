import { describe, expect, it } from 'vitest';
import { DocxEditor, LockedRegionError } from '../src/api/index.js';
import { layoutDocument, type TextItem } from '../src/layout/index.js';
import { DocxDocument } from '../src/model/index.js';
import { AccessPolicy, PermissionError } from '../src/security/index.js';
import { attr, findAll, firstChild } from '../src/xml/index.js';
import { buildDocx, p } from './helpers/docx.js';
import { testRegistry } from './helpers/fonts.js';

/** A lettered list (a., b.) like Word makes, plus a bulleted one in a Symbol font. */
const NUMBERING =
  '<w:abstractNum w:abstractNumId="0"><w:multiLevelType w:val="hybridMultilevel"/>' +
  '<w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="lowerLetter"/><w:lvlText w:val="%1."/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="1080" w:hanging="360"/></w:pPr><w:rPr><w:u w:val="none"/></w:rPr></w:lvl>' +
  '<w:lvl w:ilvl="1"><w:start w:val="1"/><w:numFmt w:val="lowerRoman"/><w:lvlText w:val="%2)"/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="1800" w:hanging="360"/></w:pPr></w:lvl>' +
  '</w:abstractNum>' +
  '<w:abstractNum w:abstractNumId="1"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="bullet"/><w:lvlText w:val=""/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr><w:rPr><w:rFonts w:ascii="Symbol" w:hAnsi="Symbol" w:hint="default"/></w:rPr></w:lvl></w:abstractNum>' +
  '<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num><w:num w:numId="2"><w:abstractNumId w:val="1"/></w:num>';

const item = (text: string, numId: string, ilvl = 0) => p(text, `<w:numPr><w:ilvl w:val="${ilvl}"/><w:numId w:val="${numId}"/></w:numPr>`);

const BODY =
  p('Payment request') +
  item('Land release fee', '1') +
  item('Compensation', '1') +
  p('Notes') +
  item('First note', '2') +
  item('Second note', '2') +
  p('Closing');

async function open(body = BODY, numbering: string | null = NUMBERING) {
  const doc = await DocxDocument.open(await buildDocx({ body, ...(numbering ? { numbering } : {}) }));
  return { doc, editor: new DocxEditor(doc) };
}

function labels(editor: DocxEditor): string[] {
  const items = editor.listItems();
  return editor.paragraphs().map((p) => items.get(p.el)?.label ?? '-');
}

function para(editor: DocxEditor, text: string) {
  const found = editor.paragraphs().find((p) => editor.textOf(editor.paragraphRange(p)!) === text);
  if (!found) throw new Error(`no paragraph ${text}`);
  return found;
}

async function reopen(doc: DocxDocument) {
  const again = await DocxDocument.open(await doc.save());
  return { doc: again, editor: new DocxEditor(again) };
}

describe('list labels', () => {
  it('reports the label each list paragraph shows', async () => {
    const { editor } = await open();
    expect(labels(editor)).toEqual(['-', 'a.', 'b.', '-', '•', '•', '-']);
    const info = editor.listItems().get(para(editor, 'Compensation').el)!;
    expect(info).toEqual({ numId: '1', level: 0, format: 'lowerLetter', label: 'b.' });
  });
});

describe('setList', () => {
  it('turns lettered items into numbers and keeps their indentation', async () => {
    const { doc, editor } = await open();
    const items = [para(editor, 'Land release fee').el, para(editor, 'Compensation').el];
    const numId = await editor.setList(items, { format: 'decimal' });
    expect(numId).toBe('3');
    expect(labels(editor)).toEqual(['-', '1.', '2.', '-', '•', '•', '-']);

    const after = await reopen(doc);
    expect(labels(after.editor)).toEqual(['-', '1.', '2.', '-', '•', '•', '-']);
    const fee = para(after.editor, 'Land release fee');
    expect(fee.props.indent).toMatchObject({ left: 54, hanging: 18 });
    // The original definition is untouched, so other lists that use it keep their letters.
    const numbering = after.doc.xmlSync(after.doc.parts.numbering)!.root;
    const original = findAll(numbering, 'w:abstractNum').find((a) => attr(a, 'w:abstractNumId') === '0')!;
    expect(attr(firstChild(firstChild(original, 'w:lvl'), 'w:numFmt'), 'w:val')).toBe('lowerLetter');
    // The copied level keeps its label formatting.
    const copy = findAll(numbering, 'w:abstractNum').find((a) => attr(a, 'w:abstractNumId') === '2')!;
    expect(firstChild(firstChild(firstChild(copy, 'w:lvl'), 'w:rPr'), 'w:u')).toBeDefined();
  });

  it('turns bullets into numbers without the bullet symbol font', async () => {
    const { doc, editor } = await open();
    await editor.setList([para(editor, 'First note').el, para(editor, 'Second note').el], { format: 'upperRoman' });
    expect(labels(editor).slice(4, 6)).toEqual(['I.', 'II.']);
    const numbering = doc.xmlSync(doc.parts.numbering)!.root;
    const lvl = firstChild(findAll(numbering, 'w:abstractNum').at(-1)!, 'w:lvl')!;
    expect(firstChild(firstChild(lvl, 'w:rPr'), 'w:rFonts')).toBeUndefined();
  });

  it('turns numbers into bullets', async () => {
    const { editor } = await open();
    await editor.setList([para(editor, 'Land release fee').el, para(editor, 'Compensation').el], { format: 'bullet' });
    expect(labels(editor).slice(1, 3)).toEqual(['•', '•']);
  });

  it('makes plain paragraphs a list, creating numbering.xml when needed', async () => {
    const { doc, editor } = await open(p('Intro') + p('One') + p('Two') + p('Three'), null);
    expect(doc.parts.numbering).toBeUndefined();
    await editor.setList(['One', 'Two', 'Three'].map((t) => para(editor, t).el), { format: 'decimalParen', start: 4 });
    expect(labels(editor)).toEqual(['-', '4)', '5)', '6)']);
    expect(para(editor, 'One').props.indent).toMatchObject({ left: 36, hanging: 18 });

    const after = await reopen(doc);
    expect(after.doc.parts.numbering).toBe('word/numbering.xml');
    expect(labels(after.editor)).toEqual(['-', '4)', '5)', '6)']);
    const page = layoutDocument(after.doc, { fonts: testRegistry() }).pages[0]!;
    const texts = page.items.filter((i): i is TextItem => i.type === 'text').map((t) => t.text);
    expect(texts).toEqual(['Intro', '4)', 'One', '5)', 'Two', '6)', 'Three']);
  });

  it('restarts numbering and keeps the format', async () => {
    const { editor } = await open();
    await editor.setList([para(editor, 'Land release fee').el, para(editor, 'Compensation').el], { start: 3 });
    expect(labels(editor).slice(1, 3)).toEqual(['c.', 'd.']);
  });

  it('changes the level of list items', async () => {
    const { editor } = await open();
    await editor.setList([para(editor, 'Compensation').el], { level: 1 });
    expect(labels(editor).slice(1, 3)).toEqual(['a.', 'i)']);
    expect(para(editor, 'Compensation').props.indent).toMatchObject({ left: 90 });
  });

  it('adds paragraphs to an existing list', async () => {
    const { editor } = await open();
    const extra = editor.insertParagraph({ after: para(editor, 'Compensation').el }, 'Admin fee');
    await editor.setList([extra], { sameListAs: para(editor, 'Compensation').el });
    expect(labels(editor).slice(1, 4)).toEqual(['a.', 'b.', 'c.']);
  });

  it('takes paragraphs out of a list, including style numbering', async () => {
    const styles =
      '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>' +
      '<w:style w:type="paragraph" w:styleId="ListNumber"><w:name w:val="List Number"/><w:pPr><w:numPr><w:numId w:val="1"/></w:numPr></w:pPr></w:style>';
    const body = BODY + p('Styled item', '<w:pStyle w:val="ListNumber"/>');
    const doc = await DocxDocument.open(await buildDocx({ body, numbering: NUMBERING, styles }));
    const editor = new DocxEditor(doc);
    expect(labels(editor).at(-1)).toBe('c.');
    await editor.setList([para(editor, 'Compensation').el, para(editor, 'Styled item').el], { format: 'none' });
    expect(labels(editor)).toEqual(['-', 'a.', '-', '-', '•', '•', '-', '-']);
    const styled = para(editor, 'Styled item').el;
    expect(attr(firstChild(firstChild(firstChild(styled, 'w:pPr'), 'w:numPr'), 'w:numId'), 'w:val')).toBe('0');
  });

  it('respects locks and the access policy', async () => {
    const { doc, editor } = await open();
    editor.lockParagraphs(para(editor, 'Land release fee').el, para(editor, 'Compensation').el, { title: 'Amounts' });
    await expect(editor.setList([para(editor, 'Compensation').el], { format: 'decimal' })).rejects.toThrow(LockedRegionError);

    const commenter = new DocxEditor(doc, new AccessPolicy(doc, { user: { id: 'c' }, role: 'commenter' }));
    await expect(commenter.setList([para(editor, 'Closing').el], { format: 'bullet' })).rejects.toThrow(PermissionError);
  });

  it('rejects options that do not fit together', async () => {
    const { editor } = await open();
    const closing = para(editor, 'Closing').el;
    await expect(editor.setList([closing], { level: 2 })).rejects.toThrow(/not in a list/);
    await expect(editor.setList([closing], {})).rejects.toThrow(/not in a list/);
    await expect(editor.setList([closing], { format: 'none', start: 2 })).rejects.toThrow(/cannot be combined/);
    await expect(editor.setList([closing], { format: 'decimal', level: 9 })).rejects.toThrow(/0 to 8/);
  });
});
