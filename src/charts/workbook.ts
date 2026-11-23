/** The small Excel workbook a Word chart embeds, so "Edit Data" works. */
import { readZip, writeZip } from '../zip/zip.js';
import { DEFAULT_DECLARATION, el, localName, parseXml, serializeNode, text, type XmlElement } from '../xml/index.js';

export interface ChartSeriesData {
  name: string;
  values: Array<number | null>;
}

const SHEET = 'Sheet1';
const MAIN = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const REL_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const PKG_REL = 'http://schemas.openxmlformats.org/package/2006/relationships';
const CT = 'http://schemas.openxmlformats.org/package/2006/content-types';
const SML = 'application/vnd.openxmlformats-officedocument.spreadsheetml';

/** Zero-based column index to its letters: 0 is A, 25 is Z, 26 is AA. */
export function columnName(index: number): string {
  let n = index + 1;
  let out = '';
  while (n > 0) {
    const rem = (n - 1) % 26;
    out = String.fromCharCode(65 + rem) + out;
    n = Math.floor((n - 1) / 26);
  }
  return out;
}

/** The sheet ranges a chart's formulas point at: series i sits in column i + 1. */
export function chartRanges(
  _seriesCount: number,
  categoryCount: number,
): { categories: string; name(i: number): string; values(i: number): string } {
  const last = categoryCount + 1;
  return {
    categories: `${SHEET}!$A$2:$A$${last}`,
    name: (i) => `${SHEET}!$${columnName(i + 1)}$1`,
    values: (i) => `${SHEET}!$${columnName(i + 1)}$2:$${columnName(i + 1)}$${last}`,
  };
}

function xml(root: XmlElement): Uint8Array {
  return new TextEncoder().encode(`<?xml ${DEFAULT_DECLARATION}?>\r\n${serializeNode(root)}`);
}

function stringCell(ref: string, value: string): XmlElement {
  return el('c', { r: ref, t: 'inlineStr' }, [el('is', {}, [el('t', { 'xml:space': 'preserve' }, [text(value)])])]);
}

function numberCell(ref: string, value: number): XmlElement {
  return el('c', { r: ref }, [el('v', {}, [text(String(value))])]);
}

export async function workbookFor(data: { categories: Array<string | number>; series: ChartSeriesData[] }): Promise<Uint8Array> {
  const rows: XmlElement[] = [];
  const header: XmlElement[] = data.series.map((s, i) => stringCell(`${columnName(i + 1)}1`, s.name));
  rows.push(el('row', { r: 1 }, header));
  data.categories.forEach((category, r) => {
    const row = r + 2;
    const cells: XmlElement[] = [
      typeof category === 'number' ? numberCell(`A${row}`, category) : stringCell(`A${row}`, category),
    ];
    data.series.forEach((s, i) => {
      const v = s.values[r];
      if (v !== null && v !== undefined) cells.push(numberCell(`${columnName(i + 1)}${row}`, v));
    });
    rows.push(el('row', { r: row }, cells));
  });
  const last = `${columnName(data.series.length)}${data.categories.length + 1}`;

  const sheet = el('worksheet', { xmlns: MAIN }, [el('dimension', { ref: `A1:${last}` }), el('sheetData', {}, rows)]);
  const workbook = el('workbook', { xmlns: MAIN, 'xmlns:r': REL_NS }, [
    el('sheets', {}, [el('sheet', { name: SHEET, sheetId: 1, 'r:id': 'rId1' })]),
  ]);
  const styles = el('styleSheet', { xmlns: MAIN }, [
    el('fonts', { count: 1 }, [el('font', {}, [el('sz', { val: 11 }), el('name', { val: 'Calibri' })])]),
    el('fills', { count: 2 }, [
      el('fill', {}, [el('patternFill', { patternType: 'none' })]),
      el('fill', {}, [el('patternFill', { patternType: 'gray125' })]),
    ]),
    el('borders', { count: 1 }, [
      el('border', {}, [el('left'), el('right'), el('top'), el('bottom'), el('diagonal')]),
    ]),
    el('cellStyleXfs', { count: 1 }, [el('xf', { numFmtId: 0, fontId: 0, fillId: 0, borderId: 0 })]),
    el('cellXfs', { count: 1 }, [el('xf', { numFmtId: 0, fontId: 0, fillId: 0, borderId: 0, xfId: 0 })]),
  ]);
  const rootRels = el('Relationships', { xmlns: PKG_REL }, [
    el('Relationship', { Id: 'rId1', Type: `${REL_NS}/officeDocument`, Target: 'xl/workbook.xml' }),
  ]);
  const workbookRels = el('Relationships', { xmlns: PKG_REL }, [
    el('Relationship', { Id: 'rId1', Type: `${REL_NS}/worksheet`, Target: 'worksheets/sheet1.xml' }),
    el('Relationship', { Id: 'rId2', Type: `${REL_NS}/styles`, Target: 'styles.xml' }),
  ]);
  const types = el('Types', { xmlns: CT }, [
    el('Default', { Extension: 'rels', ContentType: 'application/vnd.openxmlformats-package.relationships+xml' }),
    el('Default', { Extension: 'xml', ContentType: 'application/xml' }),
    el('Override', { PartName: '/xl/workbook.xml', ContentType: `${SML}.sheet.main+xml` }),
    el('Override', { PartName: '/xl/worksheets/sheet1.xml', ContentType: `${SML}.worksheet+xml` }),
    el('Override', { PartName: '/xl/styles.xml', ContentType: `${SML}.styles+xml` }),
  ]);

  return writeZip(
    [
      { name: '[Content_Types].xml', data: xml(types) },
      { name: '_rels/.rels', data: xml(rootRels) },
      { name: 'xl/workbook.xml', data: xml(workbook) },
      { name: 'xl/_rels/workbook.xml.rels', data: xml(workbookRels) },
      { name: 'xl/styles.xml', data: xml(styles) },
      { name: 'xl/worksheets/sheet1.xml', data: xml(sheet) },
    ],
    { modified: new Date(Date.UTC(2026, 0, 1)) },
  );
}

function walk(node: XmlElement, visit: (e: XmlElement) => void): void {
  visit(node);
  for (const c of node.children) if (c.type === 'element') walk(c, visit);
}

function columnIndex(letters: string): number {
  let n = 0;
  for (const ch of letters) n = n * 26 + (ch.charCodeAt(0) - 64);
  return n - 1;
}

/**
 * True when the workbook holds more than the chart's own table: a second worksheet, or a
 * non-empty cell outside A1 to the last series column and category row + 1. False when unreadable.
 */
export async function workbookHasMore(bytes: Uint8Array, categoryCount: number, seriesCount: number): Promise<boolean> {
  try {
    const entries = readZip(bytes);
    const decode = async (name: string): Promise<string | undefined> => {
      const e = entries.find((x) => x.name === name);
      return e ? new TextDecoder().decode(await e.read()) : undefined;
    };
    const sheets = entries.filter((e) => /^xl\/worksheets\/[^/]+\.xml$/.test(e.name));
    if (sheets.length > 1) return true;
    const book = await decode('xl/workbook.xml');
    if (book !== undefined) {
      let count = 0;
      walk(parseXml(book).root, (e) => {
        if (localName(e.name) === 'sheet') count++;
      });
      if (count > 1) return true;
    }
    const maxCol = seriesCount;
    const maxRow = categoryCount + 1;
    for (const sheet of sheets) {
      let more = false;
      walk(parseXml(new TextDecoder().decode(await sheet.read())).root, (e) => {
        if (localName(e.name) !== 'c') return;
        const filled = e.children.some((c) => c.type === 'element' && ['v', 'is', 'f'].includes(localName(c.name)));
        if (!filled) return;
        const m = /^([A-Z]+)(\d+)$/.exec(e.attrs.r ?? '');
        if (!m || columnIndex(m[1]!) > maxCol || Number(m[2]) > maxRow) more = true;
      });
      if (more) return true;
    }
    return false;
  } catch {
    return false;
  }
}
