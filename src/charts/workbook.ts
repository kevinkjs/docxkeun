/** The small Excel workbook a Word chart embeds, so "Edit Data" works. */
import { readZip, writeZip } from '../zip/zip.js';
import { DEFAULT_DECLARATION, el, localName, parseXml, serializeNode, text, type XmlElement } from '../xml/index.js';

export interface ChartSeriesData {
  name: string;
  values: Array<number | null>;
  /** Number format of the values, as Excel writes it (e.g. "0%", "#,##0.00"); General when left out. */
  formatCode?: string;
}

/** The data a chart's workbook holds: categories down column A, a column per series, and their number formats. */
export interface WorkbookData {
  categories: Array<string | number | null>;
  series: ChartSeriesData[];
  /** Number format of numeric categories (e.g. "mmm-yy" for dates); General when left out. */
  categoryFormat?: string;
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

function numberCell(ref: string, value: number, style: number): XmlElement {
  return el('c', { r: ref, s: style || undefined }, [el('v', {}, [text(String(value))])]);
}

/** Excel's built-in number formats (ECMA-376 18.8.30) by format code; other codes get custom ids from 164. */
const BUILT_IN_FORMATS: Record<string, number> = {
  General: 0, '0': 1, '0.00': 2, '#,##0': 3, '#,##0.00': 4, '0%': 9, '0.00%': 10, '0.00E+00': 11, '# ?/?': 12, '# ??/??': 13,
  'mm-dd-yy': 14, 'd-mmm-yy': 15, 'd-mmm': 16, 'mmm-yy': 17, 'h:mm AM/PM': 18, 'h:mm:ss AM/PM': 19, 'h:mm': 20, 'h:mm:ss': 21, 'm/d/yy h:mm': 22,
  '#,##0 ;(#,##0)': 37, '#,##0 ;[Red](#,##0)': 38, '#,##0.00;(#,##0.00)': 39, '#,##0.00;[Red](#,##0.00)': 40, 'mm:ss': 45, '[h]:mm:ss': 46, 'mmss.0': 47,
  '##0.0E+0': 48, '@': 49,
};

/** The number formats of a workbook: a cell style (index into cellXfs) per format code, and the custom formats they need. */
class NumberFormats {
  /** numFmtId of each cell style after the default one (style 0, General). */
  readonly styles: number[] = [];
  readonly custom: Array<{ id: number; code: string }> = [];
  private readonly byCode = new Map<string, number>();

  /** The cell style for a format code (0 for General or none). */
  style(code: string | undefined): number {
    if (code === undefined || code === '' || code === 'General') return 0;
    let style = this.byCode.get(code);
    if (style === undefined) {
      let id = BUILT_IN_FORMATS[code];
      if (id === undefined) {
        id = 164 + this.custom.length;
        this.custom.push({ id, code });
      }
      this.styles.push(id);
      style = this.styles.length;
      this.byCode.set(code, style);
    }
    return style;
  }
}

export async function workbookFor(data: WorkbookData): Promise<Uint8Array> {
  const formats = new NumberFormats();
  const categoryStyle = formats.style(data.categoryFormat);
  const valueStyles = data.series.map((s) => formats.style(s.formatCode));
  const rows: XmlElement[] = [];
  const header: XmlElement[] = data.series.map((s, i) => stringCell(`${columnName(i + 1)}1`, s.name));
  rows.push(el('row', { r: 1 }, header));
  data.categories.forEach((category, r) => {
    const row = r + 2;
    const cells: XmlElement[] = [];
    // A missing category (a gap in numbers or dates) has no cell.
    if (typeof category === 'number') cells.push(numberCell(`A${row}`, category, categoryStyle));
    else if (category !== null) cells.push(stringCell(`A${row}`, category));
    data.series.forEach((s, i) => {
      const v = s.values[r];
      if (v !== null && v !== undefined) cells.push(numberCell(`${columnName(i + 1)}${row}`, v, valueStyles[i]!));
    });
    rows.push(el('row', { r: row }, cells));
  });
  const last = `${columnName(data.series.length)}${data.categories.length + 1}`;

  const sheet = el('worksheet', { xmlns: MAIN }, [el('dimension', { ref: `A1:${last}` }), el('sheetData', {}, rows)]);
  const workbook = el('workbook', { xmlns: MAIN, 'xmlns:r': REL_NS }, [
    el('sheets', {}, [el('sheet', { name: SHEET, sheetId: 1, 'r:id': 'rId1' })]),
  ]);
  const styles = el('styleSheet', { xmlns: MAIN }, [
    ...(formats.custom.length ? [el('numFmts', { count: formats.custom.length }, formats.custom.map((f) => el('numFmt', { numFmtId: f.id, formatCode: f.code })))] : []),
    el('fonts', { count: 1 }, [el('font', {}, [el('sz', { val: 11 }), el('name', { val: 'Calibri' })])]),
    el('fills', { count: 2 }, [
      el('fill', {}, [el('patternFill', { patternType: 'none' })]),
      el('fill', {}, [el('patternFill', { patternType: 'gray125' })]),
    ]),
    el('borders', { count: 1 }, [
      el('border', {}, [el('left'), el('right'), el('top'), el('bottom'), el('diagonal')]),
    ]),
    el('cellStyleXfs', { count: 1 }, [el('xf', { numFmtId: 0, fontId: 0, fillId: 0, borderId: 0 })]),
    el('cellXfs', { count: 1 + formats.styles.length }, [
      el('xf', { numFmtId: 0, fontId: 0, fillId: 0, borderId: 0, xfId: 0 }),
      ...formats.styles.map((id) => el('xf', { numFmtId: id, fontId: 0, fillId: 0, borderId: 0, xfId: 0, applyNumberFormat: 1 })),
    ]),
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

/** Whether bytes are a workbook package that can be read: a zip with a workbook or worksheet part. */
export function isWorkbook(bytes: Uint8Array): boolean {
  try {
    return readZip(bytes).some((e) => e.name === 'xl/workbook.xml' || /^xl\/worksheets\/[^/]+\.xml$/.test(e.name));
  } catch {
    return false;
  }
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
