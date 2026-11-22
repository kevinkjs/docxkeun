/** The playground sample exports to a valid PDF that contains its drawings, chart and equations. */
import { expect, it } from 'vitest';
import { sampleDocx } from '../playground/sample.js';
import { layoutDocument } from '../src/layout/index.js';
import { DocxDocument } from '../src/model/index.js';
import { exportPdf } from '../src/pdf/index.js';
import { testRegistry } from './helpers/fonts.js';
import { pageText, pages, parsePdf, streamData } from './helpers/pdf.js';

it('exports the sample, drawings included, to a well-formed PDF', async () => {
  const doc = await DocxDocument.open(await sampleDocx());
  const fonts = testRegistry();
  const layout = layoutDocument(doc, { fonts });
  const bytes = await exportPdf(doc, { fonts, layout });
  const pdf = parsePdf(bytes);
  const all = pages(pdf);
  expect(all).toHaveLength(layout.pages.length);
  const figures = layout.pages.findIndex((p) => p.items.some((i) => i.type === 'text' && i.text.includes('Coal hauled')));
  expect(figures).toBeGreaterThan(0);
  const ops = Buffer.from(streamData(pdf, all[figures]!.get('Contents'))).toString('latin1');
  // Curves (rounded text box, can, chart legend markers, radical) and filled paths are drawn.
  expect((ops.match(/ c\b/g) ?? []).length).toBeGreaterThan(10);
  expect(ops).toContain('0.1804 0.4549 0.7098 RG'); // the text box outline, 2E74B5
  expect(ops).toMatch(/ f\b|B\b/);
  // Chart and equation text is real, searchable text.
  const text = pageText(pdf, all[figures]!);
  expect(text).toContain('Coal hauled, 2027');
  expect(text).toContain('Rehabilitation starts');
});
