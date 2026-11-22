/** The playground's sample document (also the docxkeun chat sample): its lists render with their labels. */
import { describe, expect, it } from 'vitest';
import { DocxEditor } from '../src/api/index.js';
import { layoutDocument, type TextItem } from '../src/layout/index.js';
import { DocxDocument } from '../src/model/index.js';
import { sampleDocx } from '../playground/sample.js';
import { testRegistry } from './helpers/fonts.js';

describe('sample document', () => {
  it('labels every list item, bullets included', async () => {
    const doc = await DocxDocument.open(await sampleDocx());
    const editor = new DocxEditor(doc);
    const items = editor.listItems();
    const labels = editor
      .paragraphs()
      .filter((p) => items.has(p.el))
      .map((p) => `${items.get(p.el)!.label} ${editor.textOf(editor.paragraphRange(p)!)}`);
    expect(labels).toContain('• Bulleted point one');
    expect(labels).toContain('• Bulleted point two');
    expect(labels.slice(0, 5)).toEqual(['1. Exploration drilling and sampling', '2. Overburden removal', '(a) Topsoil stockpiling', '(b) Progressive rehabilitation', '3. Coal hauling to the port']);
    // No list item may end up without a label.
    expect(labels.filter((l) => l.startsWith(' '))).toEqual([]);
  });

  it('draws the bullets on the page', async () => {
    const doc = await DocxDocument.open(await sampleDocx());
    const texts = layoutDocument(doc, { fonts: testRegistry() })
      .pages.flatMap((page) => page.items)
      .filter((i): i is TextItem => i.type === 'text')
      .map((t) => t.text);
    const first = texts.indexOf('Bulleted point one');
    expect(first).toBeGreaterThan(0);
    expect(texts[first - 1]).toBe('•');
    expect(texts[texts.indexOf('Bulleted point two') - 1]).toBe('•');
  });
});
