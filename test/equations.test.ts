import { describe, expect, it } from 'vitest';
import { layoutDocument, type PathItem, type TextItem } from '../src/layout/index.js';
import { blocksToText, DocxDocument } from '../src/model/index.js';
import { buildDocx } from './helpers/docx.js';
import { testRegistry } from './helpers/fonts.js';

const M = 'xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"';
const frac = `<m:oMath><m:f><m:num><m:r><m:t>a+b</m:t></m:r></m:num><m:den><m:r><m:t>2</m:t></m:r></m:den></m:f></m:oMath>`;

async function page(body: string) {
  const doc = await DocxDocument.open(await buildDocx({ body }));
  const result = layoutDocument(doc, { fonts: testRegistry() }).pages[0]!;
  return { doc, items: result.items, texts: result.items.filter((i): i is TextItem => i.type === 'text'), paths: result.items.filter((i): i is PathItem => i.type === 'path') };
}

describe('equations', () => {
  it('typesets a display fraction, centred on its own line, with a fraction bar', async () => {
    const { doc, texts, paths, items } = await page(`<w:p ${M}><m:oMathPara>${frac}</m:oMathPara></w:p>`);
    expect(blocksToText(doc.body().blocks)).toBe('(a+b)/2');
    const num = texts.find((t) => t.text.includes('a'))!;
    const den = texts.find((t) => t.text === '2')!;
    expect(num.y).toBeLessThan(den.y);
    // The bar lies between numerator and denominator.
    const bar = paths.find((x) => x.fill && x.commands.length <= 6) ?? items.find((i) => i.type === 'rect' || i.type === 'line');
    expect(bar).toBeDefined();
    // Centred between the margins (72 … 540).
    const left = Math.min(num.x, den.x);
    const right = Math.max(num.xs[num.xs.length - 1]! + num.widths[num.widths.length - 1]!, den.x + den.widths[0]!);
    expect((left + right) / 2).toBeCloseTo(306, -1);
  });

  it('places an inline equation in running text', async () => {
    const { texts } = await page(`<w:p ${M}><w:r><w:t xml:space="preserve">Half of it is </w:t></w:r>${frac}<w:r><w:t xml:space="preserve"> here.</w:t></w:r></w:p>`);
    const before = texts.find((t) => t.text.startsWith('Half'))!;
    const after = texts.find((t) => t.text.includes('here'))!;
    const num = texts.find((t) => t.text === 'a')!;
    const den = texts.find((t) => t.text === '2')!;
    const beforeEnd = before.xs[before.xs.length - 1]! + before.widths[before.widths.length - 1]!;
    expect(num.x).toBeGreaterThan(beforeEnd - 0.5);
    expect(after.xs[0]!).toBeGreaterThan(num.x);
    // Numerator above the baseline, denominator below it.
    expect(num.y).toBeLessThan(before.y);
    expect(den.y).toBeGreaterThan(before.y);
    // Same line: the text after follows on the baseline of the text before.
    expect(after.y).toBeCloseTo(before.y, 1);
  });
});
