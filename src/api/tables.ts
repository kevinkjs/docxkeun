/**
 * Building blocks for DocxEditor's table and copy operations: a new table's XML, cleaning a copied
 * block so it can live next to its original, and marking new content as a tracked insertion.
 */
import { el, findAll, firstChild, insertOrdered, removeChildren, type XmlElement, type XmlNode } from '../xml/index.js';
import { insertedRuns, type RevisionAuthor } from './revisions.js';

/** Options for a new table. */
export interface TableOptions {
  /** The first row is a header: bold, and repeated at the top of each page the table runs onto. */
  header?: boolean;
  /** Relative column widths (e.g. [1, 3, 1]); default: equal. The table spans the text width. */
  widths?: number[];
  /** Grid borders (default true). Ignored when `style` is given. */
  borders?: boolean;
  /** The table's position between the margins (default left). */
  align?: 'left' | 'center' | 'right';
  /** A table style id; default: Word's "Table Grid" when the document has it. */
  style?: string;
}

const BORDER_SIDES = ['w:top', 'w:left', 'w:bottom', 'w:right', 'w:insideH', 'w:insideV'];
const TRPR_ORDER = ['w:cnfStyle', 'w:divId', 'w:gridBefore', 'w:gridAfter', 'w:wBefore', 'w:wAfter', 'w:cantSplit', 'w:trHeight', 'w:tblHeader', 'w:tblCellSpacing', 'w:jc', 'w:hidden', 'w:ins', 'w:del', 'w:trPrChange'];

/**
 * A new table (w:tbl). `widths` are the columns' widths in twips; `paragraph` makes a cell's
 * paragraph for one line of text (lines come from "\n" in the cell text).
 */
export function buildTable(
  rows: string[][],
  widths: number[],
  o: { header: boolean; style: string | undefined; borders: boolean; align: TableOptions['align']; paragraph: (text: string, header: boolean) => XmlElement },
): XmlElement {
  const total = widths.reduce((a, b) => a + b, 0);
  const tblPr = el('w:tblPr', {}, [
    ...(o.style ? [el('w:tblStyle', { 'w:val': o.style })] : []),
    el('w:tblW', { 'w:w': String(total), 'w:type': 'dxa' }),
    ...(o.align && o.align !== 'left' ? [el('w:jc', { 'w:val': o.align })] : []),
    ...(o.borders && !o.style ? [el('w:tblBorders', {}, BORDER_SIDES.map((n) => el(n, { 'w:val': 'single', 'w:sz': '4', 'w:space': '0', 'w:color': 'auto' })))] : []),
    el('w:tblLayout', { 'w:type': 'fixed' }),
    el('w:tblLook', { 'w:val': '04A0', 'w:firstRow': o.header ? '1' : '0', 'w:lastRow': '0', 'w:firstColumn': '1', 'w:lastColumn': '0', 'w:noHBand': '0', 'w:noVBand': '1' }),
  ]);
  const grid = el('w:tblGrid', {}, widths.map((w) => el('w:gridCol', { 'w:w': String(w) })));
  const trs = rows.map((cells, r) => {
    const isHeader = o.header && r === 0;
    return el('w:tr', {}, [
      ...(isHeader ? [el('w:trPr', {}, [el('w:tblHeader')])] : []),
      ...widths.map((w, c) =>
        el('w:tc', {}, [el('w:tcPr', {}, [el('w:tcW', { 'w:w': String(w), 'w:type': 'dxa' })]), ...(cells[c] ?? '').split('\n').map((line) => o.paragraph(line, isHeader))]),
      ),
    ]);
  });
  return el('w:tbl', {}, [tblPr, grid, ...trs]);
}

/** Column widths in twips: `weights` (default equal) scaled to fill `total`. */
export function columnWidths(columns: number, total: number, weights?: number[]): number[] {
  const w = weights ?? Array.from({ length: columns }, () => 1);
  const sum = w.reduce((a, b) => a + b, 0);
  const out = w.map((x) => Math.max(1, Math.floor((x / sum) * total)));
  // Rounding: the last column takes what is left so the table is exactly `total` wide.
  out[out.length - 1]! += total - out.reduce((a, b) => a + b, 0);
  return out;
}

// ---- copies ---------------------------------------------------------------------------------------

/** What a copy left out because it can only exist once in a document. */
export interface CopyDropped {
  /** Comment anchors (the comments stay on the original). */
  comments: number;
  /** Footnote and endnote references. */
  notes: number;
  /** Bookmarks (their names must be unique; links keep pointing at the original). */
  bookmarks: number;
}

/** Fresh ids for the copy's elements that must be unique in the document. */
export interface CopyIds {
  revision: () => string;
  drawing: () => string;
  sdt: () => string;
}

const DROP_MARKERS = new Set(['w:bookmarkStart', 'w:bookmarkEnd', 'w:commentRangeStart', 'w:commentRangeEnd', 'w:permStart', 'w:permEnd']);
const REVISIONS_WITH_ID = new Set([
  'w:ins', 'w:del', 'w:moveFrom', 'w:moveTo', 'w:moveFromRangeStart', 'w:moveFromRangeEnd', 'w:moveToRangeStart', 'w:moveToRangeEnd',
  'w:rPrChange', 'w:pPrChange', 'w:sectPrChange', 'w:tblPrChange', 'w:tblGridChange', 'w:trPrChange', 'w:tcPrChange', 'w:tblPrExChange',
  'w:numberingChange', 'w:cellIns', 'w:cellDel', 'w:cellMerge',
]);

/**
 * Make a cloned block fit to sit next to its original: drop what may exist only once (bookmarks,
 * comment anchors, note references, editable-range markers, a paragraph's section break), give
 * drawings, content controls and tracked changes new ids, and drop Word's paragraph ids.
 */
export function cleanCopy(node: XmlElement, ids: CopyIds, dropped: CopyDropped): void {
  delete node.attrs['w14:paraId'];
  delete node.attrs['w14:textId'];
  if (node.name === 'wp:docPr') node.attrs.id = ids.drawing();
  if (REVISIONS_WITH_ID.has(node.name) && node.attrs['w:id'] !== undefined) node.attrs['w:id'] = ids.revision();
  if (node.name === 'w:sdtPr') {
    const id = firstChild(node, 'w:id');
    if (id) id.attrs['w:val'] = ids.sdt();
  }
  const kept: XmlNode[] = [];
  for (const c of node.children) {
    if (c.type === 'element') {
      if (DROP_MARKERS.has(c.name)) {
        if (c.name === 'w:bookmarkStart') dropped.bookmarks++;
        continue;
      }
      if (node.name === 'w:pPr' && c.name === 'w:sectPr') continue;
      if (c.name === 'w:r') {
        if (firstChild(c, 'w:commentReference')) {
          dropped.comments++;
          continue;
        }
        if (firstChild(c, 'w:footnoteReference') || firstChild(c, 'w:endnoteReference')) {
          dropped.notes++;
          continue;
        }
      }
      cleanCopy(c, ids, dropped);
    }
    kept.push(c);
  }
  node.children = kept;
}

// ---- tracked insertion ----------------------------------------------------------------------------

const RUN_CONTAINERS = new Set(['w:hyperlink', 'w:smartTag', 'w:fldSimple', 'w:sdt', 'w:sdtContent', 'w:customXml']);

/** Wrap the runs of a paragraph (and of its links and fields) in w:ins, a group at a time. */
function wrapRuns(container: XmlElement, nextId: () => string, who: RevisionAuthor): void {
  const out: XmlNode[] = [];
  let group: XmlElement[] = [];
  const flush = () => {
    if (group.length) out.push(insertedRuns(group, nextId(), who));
    group = [];
  };
  for (const c of container.children) {
    if (c.type === 'element' && c.name === 'w:r') {
      group.push(c);
      continue;
    }
    flush();
    if (c.type === 'element' && RUN_CONTAINERS.has(c.name)) wrapRuns(c, nextId, who);
    out.push(c);
  }
  flush();
  container.children = out;
}

/** Mark a table row inserted or deleted (w:trPr/w:ins or w:del). */
export function markRow(row: XmlElement, kind: 'ins' | 'del', id: string, who: RevisionAuthor): void {
  let trPr = firstChild(row, 'w:trPr');
  if (!trPr) {
    trPr = el('w:trPr');
    const exceptions = firstChild(row, 'w:tblPrEx');
    row.children.splice(exceptions ? row.children.indexOf(exceptions) + 1 : 0, 0, trPr);
  }
  removeChildren(trPr, 'w:ins');
  removeChildren(trPr, 'w:del');
  insertOrdered(trPr, el(kind === 'ins' ? 'w:ins' : 'w:del', { 'w:id': id, 'w:author': who.author, 'w:date': who.date }), TRPR_ORDER);
}

/**
 * Mark new content (a paragraph, a table, a row) as a tracked insertion by `who`: its runs, its
 * paragraph marks and its table rows. `markParagraph` marks one paragraph mark.
 */
export function markInserted(node: XmlElement, nextId: () => string, who: RevisionAuthor, markParagraph: (p: XmlElement) => void): void {
  for (const p of node.name === 'w:p' ? [node] : findAll(node, 'w:p')) {
    wrapRuns(p, nextId, who);
    markParagraph(p);
  }
  for (const row of node.name === 'w:tr' ? [node] : findAll(node, 'w:tr')) markRow(row, 'ins', nextId(), who);
}
