import { describe, expect, it } from 'vitest';
import { layoutDocument, type TextItem } from '../src/layout/index.js';
import { DocxDocument } from '../src/model/index.js';
import { buildDocx, p } from './helpers/docx.js';
import { testRegistry } from './helpers/fonts.js';

const PAGE = '<w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>';
const twoColumns = `<w:sectPr>${PAGE}<w:cols w:num="2" w:space="720"/></w:sectPr>`;
const continuousOne = `<w:sectPr><w:type w:val="continuous"/>${PAGE}<w:cols w:num="1"/></w:sectPr>`;

async function texts(body: string, sectPr = continuousOne) {
  const doc = await DocxDocument.open(await buildDocx({ body, sectPr }));
  const page = layoutDocument(doc, { fonts: testRegistry() }).pages[0]!;
  return page.items.filter((i): i is TextItem => i.type === 'text');
}

describe('column balancing', () => {
  it('balances a multi-column section that ends with a continuous break', async () => {
    const lines = Array.from({ length: 6 }, (_, i) => p(`Line ${i + 1}`));
    const body = lines.slice(0, 5).join('') + p('Line 6', twoColumns) + p('After the columns');
    const t = await texts(body);
    const at = (s: string) => t.find((x) => x.text === s)!;
    const left = 72;
    // Three lines in each column, the second column starting at the top again.
    expect(['Line 1', 'Line 2', 'Line 3'].map((s) => at(s).x)).toEqual([left, left, left]);
    expect(['Line 4', 'Line 5', 'Line 6'].every((s) => at(s).x > 300)).toBe(true);
    expect(at('Line 4').y).toBeCloseTo(at('Line 1').y, 3);
    // The next section starts below the balanced columns, not below six lines.
    expect(at('After the columns').y).toBeGreaterThan(at('Line 3').y);
    expect(at('After the columns').y - at('Line 3').y).toBeLessThan((at('Line 3').y - at('Line 1').y) / 2 + 20);
  });

  it('splits a long paragraph across the balanced columns', async () => {
    const long = p(Array.from({ length: 40 }, (_, i) => `word${i}`).join(' '), twoColumns);
    const t = await texts(long + p('After'));
    const left = t.filter((x) => x.x < 300 && x.text !== 'After');
    const right = t.filter((x) => x.x > 300);
    expect(left.length).toBeGreaterThan(0);
    expect(right.length).toBeGreaterThan(0);
    const leftLines = new Set(left.map((x) => Math.round(x.y))).size;
    const rightLines = new Set(right.map((x) => Math.round(x.y))).size;
    expect(Math.abs(leftLines - rightLines)).toBeLessThanOrEqual(1);
  });

  it('leaves columns alone at the end of the document and around explicit column breaks', async () => {
    const lines = Array.from({ length: 4 }, (_, i) => p(`Line ${i + 1}`)).join('');
    const end = await texts(lines, twoColumns);
    expect(end.every((x) => x.x < 300)).toBe(true);
    const withBreak = p('One') + '<w:p><w:r><w:br w:type="column"/></w:r></w:p>' + p('Two', twoColumns) + p('After');
    const t = await texts(withBreak);
    expect(t.find((x) => x.text === 'Two')!.x).toBeGreaterThan(300);
  });
});
