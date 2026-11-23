/** Writing charts: the embedded workbook, and building chart parts. */
import { describe, expect, it } from 'vitest';
import { chartRanges, columnName, workbookFor, workbookHasMore } from '../src/charts/workbook.js';
import { parseXml } from '../src/xml/parser.js';
import { findAll, attr, textContent } from '../src/xml/index.js';
import { readZip } from '../src/zip/zip.js';

describe('chart workbook', () => {
  it('names columns and chart ranges', () => {
    expect([0, 25, 26, 701, 702].map(columnName)).toEqual(['A', 'Z', 'AA', 'ZZ', 'AAA']);
    const r = chartRanges(2, 4);
    expect([r.name(0), r.categories, r.values(1)]).toEqual(['Sheet1!$B$1', 'Sheet1!$A$2:$A$5', 'Sheet1!$C$2:$C$5']);
  });

  it('workbookFor writes one sheet with the table, escaped, and leaves missing values empty', async () => {
    const bytes = await workbookFor({ categories: ['Q1', 'R&D <2026>'], series: [{ name: 'Sales "EU"', values: [1.5, null] }] });
    const entries = readZip(bytes);
    expect(entries.map((e) => e.name).sort()).toEqual([
      '[Content_Types].xml',
      '_rels/.rels',
      'xl/_rels/workbook.xml.rels',
      'xl/styles.xml',
      'xl/workbook.xml',
      'xl/worksheets/sheet1.xml',
    ]);
    const sheet = entries.find((e) => e.name === 'xl/worksheets/sheet1.xml')!;
    const root = parseXml(new TextDecoder().decode(await sheet.read())).root;
    const cells = Object.fromEntries(findAll(root, 'c').map((c) => [attr(c, 'r'), textContent(c)]));
    expect(cells).toEqual({ B1: 'Sales "EU"', A2: 'Q1', B2: '1.5', A3: 'R&D <2026>' });
  });

  it('workbookHasMore spots data beyond the chart table', async () => {
    const bytes = await workbookFor({ categories: ['Q1', 'R&D <2026>'], series: [{ name: 'Sales "EU"', values: [1.5, null] }] });
    expect(await workbookHasMore(bytes, 2, 1)).toBe(false);
    expect(await workbookHasMore(await workbookFor({ categories: ['a', 'b', 'c'], series: [{ name: 's', values: [1, 2, 3] }] }), 2, 1)).toBe(true);
    expect(await workbookHasMore(new Uint8Array([1, 2, 3]), 2, 1)).toBe(false);
  });
});
