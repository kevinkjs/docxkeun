import { describe, expect, it } from 'vitest';
import { DocxEditor } from '../src/api/index.js';
import { LayoutCache, layoutDocument, type LayoutResult } from '../src/layout/index.js';
import { DocxDocument } from '../src/model/index.js';
import { buildDocx, p } from './helpers/docx.js';
import { testRegistry } from './helpers/fonts.js';

const fonts = testRegistry();

/** Page items reduced to what a reader sees. */
function snapshot(result: LayoutResult) {
  return result.pages.map((page) => ({
    items: page.items.map((i) => (i.type === 'text' ? `${i.text}@${i.x.toFixed(2)},${i.y.toFixed(2)}` : i.type)),
    comments: page.comments.map((c) => `${c.id}@${c.x.toFixed(1)},${c.y.toFixed(1)}`),
  }));
}

const lorem = 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua.';
const numbering =
  '<w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/><w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr></w:lvl></w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>';
const item = (text: string) => p(text, '<w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr>');

describe('layout cache', () => {
  it('gives the same pages as a fresh layout after edits', async () => {
    const body =
      p(`Intro ${lorem}`) +
      '<w:p><w:commentRangeStart w:id="7"/><w:r><w:t>Commented start. </w:t></w:r></w:p>' +
      p(lorem) +
      '<w:p><w:r><w:t>Commented end.</w:t></w:r><w:commentRangeEnd w:id="7"/></w:p>' +
      Array.from({ length: 30 }, (_, i) => item(`Item ${i} ${lorem}`)).join('') +
      Array.from({ length: 30 }, (_, i) => p(`Tail ${i} ${lorem}`)).join('');
    const doc = await DocxDocument.open(await buildDocx({ body, numbering }));
    const cache = new LayoutCache();
    layoutDocument(doc, { fonts, cache });
    const editor = new DocxEditor(doc);
    // A new list item shifts every later number; a longer paragraph pushes the rest down.
    editor.splitParagraph(editor.find('Item 2 ')[0]!.end);
    editor.insertText(editor.find('Tail 3 ')[0]!.end, `${lorem} ${lorem} `);
    const cached = layoutDocument(doc, { fonts, cache });
    // Items after the split get new numbers, so only they and the two edited paragraphs are laid out again.
    expect(cache.misses).toBe(30);
    expect(cache.hits).toBe(35);
    expect(snapshot(cached)).toEqual(snapshot(layoutDocument(doc, { fonts })));
  });

  it('starts over when styles change', async () => {
    const styles = '<w:docDefaults><w:rPrDefault><w:rPr><w:sz w:val="22"/></w:rPr></w:rPrDefault></w:docDefaults>';
    const doc = await DocxDocument.open(await buildDocx({ body: p(lorem) + p(lorem), styles }));
    const cache = new LayoutCache();
    layoutDocument(doc, { fonts, cache });
    const sz = doc.xmlSync(doc.parts.styles!)!;
    const szEl = JSON.stringify(sz).includes('"w:sz"');
    expect(szEl).toBe(true);
    // Make the default size bigger.
    const walk = (n: { name?: string; attrs?: Record<string, string>; children?: unknown[] }): void => {
      if (n.name === 'w:sz') n.attrs!['w:val'] = '40';
      for (const c of n.children ?? []) walk(c as never);
    };
    walk(sz.root as never);
    doc.touch(doc.parts.styles!);
    const after = layoutDocument(doc, { fonts, cache });
    expect(cache.hits).toBe(0);
    expect(snapshot(after)).toEqual(snapshot(layoutDocument(doc, { fonts })));
  });
});
