/**
 * Objects in a story (pictures, charts, SmartArt, shapes, text boxes, groups): listing them in
 * document order, the drawing XML for new ones, and the parts behind a chart (its workbook, and
 * copies of a chart part for copied content).
 */
import type { ChartTable, NewChart, NewChartType } from '../charts/write.js';
import type { DocParagraph, ImageInline } from '../model/content.js';
import type { DocxDocument } from '../model/document.js';
import { CONTENT_TYPE, REL, type OpcPackage } from '../opc/index.js';
import { attr, childPath, cloneElement, el, elements, firstChild, insertOrdered, NS, walk, type XmlElement } from '../xml/index.js';
import { sectionProperties } from './headers.js';
import { EditError, TreeIndex } from './positions.js';

export interface DocObject {
  /** Object number, from 1, in the order of `DocxEditor.paragraphs()` and then along each paragraph. */
  n: number;
  kind: 'picture' | 'chart' | 'diagram' | 'shape' | 'textBox' | 'group';
  /** The paragraph holding the object. */
  paragraph: DocParagraph;
  /** The w:r holding the object. */
  run: XmlElement;
  /** The w:drawing (inside mc:Choice when wrapped), or the w:pict / w:object of an old VML object. */
  drawing: XmlElement;
  vml: boolean;
  /** Size in points. */
  width: number;
  height: number;
  floating: boolean;
  alt?: string;
  /** The chart part, for a chart. */
  chartPart?: string;
}

/** A chart's data, and where it comes from: an embedded workbook, a linked outside file, or neither. */
export interface ChartInfo extends ChartTable {
  source: 'embedded' | 'linked' | 'none';
}

export interface ChartUpdate {
  source: 'embedded' | 'linked' | 'none';
  /** The embedded workbook was rewritten with the chart's new data. */
  workbookRewritten: boolean;
  /**
   * The old workbook held more than the chart's table (other sheets or cells), which is gone now;
   * also true when an old embedded object that can't be read (an .xls) was replaced.
   */
  workbookHadMore: boolean;
}

// ---------------------------------------------------------------------------------------------
// Listing

function kindOf(inline: ImageInline, vml: boolean): DocObject['kind'] {
  switch (inline.content?.type) {
    case 'chart':
      return 'chart';
    case 'diagram':
      return 'diagram';
    case 'group':
      return 'group';
    case 'shape': {
      // The reader only makes a VML shape of an old text box.
      if (vml) return 'textBox';
      const flag = attr(childPath(inline.drawing, 'a:graphic', 'a:graphicData', 'wps:wsp', 'wps:cNvSpPr'), 'txBox');
      return flag === '1' || flag === 'true' ? 'textBox' : 'shape';
    }
    default:
      return 'picture';
  }
}

/** The objects of some paragraphs, numbered from 1 (a group counts once, not per shape in it). */
export function listObjects(paragraphs: DocParagraph[]): DocObject[] {
  const out: DocObject[] = [];
  for (const paragraph of paragraphs) {
    let idx: TreeIndex | undefined;
    for (const inline of paragraph.inlines) {
      if (inline.kind !== 'image') continue;
      // The reader reads mc:Choice content through a stand-in run, so find the real one in the tree.
      idx ??= new TreeIndex(paragraph.el);
      const vml = inline.drawing.name === 'w:pict' || inline.drawing.name === 'w:object';
      const drawing = vml ? inline.drawing : (idx.parent.get(inline.drawing) ?? inline.drawing);
      const obj: DocObject = {
        n: out.length + 1,
        kind: kindOf(inline, vml),
        paragraph,
        run: idx.ancestor(drawing, 'w:r') ?? inline.run,
        drawing,
        vml,
        width: inline.width,
        height: inline.height,
        floating: inline.float !== undefined,
      };
      if (inline.alt) obj.alt = inline.alt;
      if (inline.content?.type === 'chart' && inline.content.part) obj.chartPart = inline.content.part;
      out.push(obj);
    }
  }
  return out;
}

// ---------------------------------------------------------------------------------------------
// Drawing XML

const emu = (pt: number): string => String(Math.round(pt * 12700));

/** A w:drawing with an inline frame of `widthPt` × `heightPt` points around `graphicData`. */
export function inlineFrame(id: number, name: string, widthPt: number, heightPt: number, graphicData: XmlElement, alt?: string): XmlElement {
  const docPr: Record<string, string> = { id: String(id), name };
  if (alt) docPr.descr = alt;
  return el('w:drawing', {}, [
    el('wp:inline', { distT: 0, distB: 0, distL: 0, distR: 0 }, [
      el('wp:extent', { cx: emu(widthPt), cy: emu(heightPt) }),
      el('wp:effectExtent', { l: 0, t: 0, r: 0, b: 0 }),
      el('wp:docPr', docPr),
      el('wp:cNvGraphicFramePr'),
      el('a:graphic', { 'xmlns:a': NS.a }, [graphicData]),
    ]),
  ]);
}

/** The graphic of a chart whose part the relationship `relId` points at. */
export function chartGraphic(relId: string): XmlElement {
  return el('a:graphicData', { uri: 'http://schemas.openxmlformats.org/drawingml/2006/chart' }, [el('c:chart', { 'xmlns:c': NS.c, 'xmlns:r': NS.r, 'r:id': relId })]);
}

/** The w:sectPr of the section holding `node`: the next paragraph section break, or the body's own. */
export function sectionAt(body: XmlElement, idx: TreeIndex, node: XmlElement): XmlElement | undefined {
  const sections = sectionProperties(body);
  const at = idx.order.get(node);
  if (at === undefined) return sections[sections.length - 1];
  for (const sectPr of sections) {
    const paragraph = idx.ancestor(sectPr, 'w:p');
    if (!paragraph) return sectPr;
    let last = paragraph;
    for (let c = elements(last).at(-1); c; c = elements(last).at(-1)) last = c;
    if (at <= idx.order.get(last)!) return sectPr;
  }
  return sections[sections.length - 1];
}

const CHART_WIDTH = 432;
const CHART_HEIGHT = 252;

/** A new chart's size in points: as given (one side keeps Word's 432 × 252 proportions), else Word's default shrunk to `maxWidth`. */
export function chartSize(size: { width?: number | undefined; height?: number | undefined }, maxWidth: number): { width: number; height: number } {
  for (const side of [size.width, size.height]) {
    if (side !== undefined && !(side > 0 && Number.isFinite(side))) throw new EditError('A chart needs a positive width and height');
  }
  const { width, height } = size;
  if (width !== undefined) return { width, height: height ?? (width * CHART_HEIGHT) / CHART_WIDTH };
  if (height !== undefined) return { width: (height * CHART_WIDTH) / CHART_HEIGHT, height };
  const scale = Math.min(1, maxWidth / CHART_WIDTH);
  return { width: CHART_WIDTH * scale, height: CHART_HEIGHT * scale };
}

// ---------------------------------------------------------------------------------------------
// Chart parts

/** `{base}{n}{ext}` with the lowest n from 1 that no part has. */
export function freePartName(pkg: OpcPackage, base: string, ext: string): string {
  let n = 1;
  while (pkg.has(`${base}${n}${ext}`)) n++;
  return `${base}${n}${ext}`;
}

/** Store a new embedded workbook and relate it to a chart part; returns the relationship id. */
export async function addWorkbook(pkg: OpcPackage, chartPart: string, bytes: Uint8Array): Promise<string> {
  const name = freePartName(pkg, 'word/embeddings/Microsoft_Excel_Worksheet', '.xlsx');
  pkg.setDefaultContentType('xlsx', CONTENT_TYPE.xlsx);
  pkg.setBytes(name, bytes);
  return pkg.addRelationship(chartPart, REL.package, name);
}

/** The chart part and its c:chartSpace behind a chart object. */
export function chartOf(doc: DocxDocument, obj: DocObject): { part: string; chartSpace: XmlElement } {
  if (obj.kind !== 'chart') throw new EditError(`Object ${obj.n} is ${obj.kind === 'textBox' ? 'a text box' : `a ${obj.kind}`}, not a chart`);
  const root = obj.chartPart ? doc.xmlSync(obj.chartPart)?.root : undefined;
  if (!obj.chartPart || root?.name !== 'c:chartSpace') throw new EditError(`The data of chart ${obj.n} is missing from the file`);
  return { part: obj.chartPart, chartSpace: root };
}

const PLOT_TYPES: Record<string, NewChartType> = { 'c:lineChart': 'line', 'c:areaChart': 'area', 'c:pieChart': 'pie', 'c:doughnutChart': 'doughnut', 'c:scatterChart': 'scatter' };

/** The new-chart type of a chart's first plot (undefined for kinds new charts can't have), and its stacking. */
export function chartKind(chartSpace: XmlElement): { type: NewChartType | undefined; stacked: NewChart['stacked'] } {
  const plotArea = childPath(chartSpace, 'c:chart', 'c:plotArea');
  const plot = plotArea ? elements(plotArea).find((e) => e.name.endsWith('Chart')) : undefined;
  if (!plot) return { type: undefined, stacked: undefined };
  const type = plot.name === 'c:barChart' ? (attr(firstChild(plot, 'c:barDir'), 'val') === 'bar' ? 'bar' : 'column') : PLOT_TYPES[plot.name];
  // Only column, bar and area charts stack as far as new data is concerned (a stacked line is still a line).
  if (type !== 'column' && type !== 'bar' && type !== 'area') return { type, stacked: undefined };
  const grouping = attr(firstChild(plot, 'c:grouping'), 'val');
  return { type, stacked: grouping === 'stacked' ? 'stacked' : grouping === 'percentStacked' ? 'percent' : undefined };
}

/**
 * Where a chart's data lives (its c:externalData relationship): an outside file ('linked'), a part
 * in the package ('embedded'; `workbook` is set when that part is an .xlsx package that can be
 * rewritten, not an old OLE object), or nowhere.
 */
export function chartData(doc: DocxDocument, chartPart: string, chartSpace: XmlElement): { source: ChartInfo['source']; relId: string | undefined; workbook: string | undefined } {
  const relId = attr(firstChild(chartSpace, 'c:externalData'), 'r:id');
  const rel = relId === undefined ? undefined : doc.rels(chartPart).get(relId);
  if (!rel) return { source: 'none', relId, workbook: undefined };
  if (rel.external) return { source: 'linked', relId, workbook: undefined };
  if (!doc.pkg.has(rel.target)) return { source: 'none', relId, workbook: undefined };
  return { source: 'embedded', relId, workbook: doc.pkg.contentType(rel.target) === CONTENT_TYPE.xlsx ? rel.target : undefined };
}

const CHARTSPACE_ORDER = [
  'c:date1904', 'c:lang', 'c:roundedCorners', 'mc:AlternateContent', 'c:style', 'c:clrMapOvr', 'c:pivotSource', 'c:protection', 'c:chart', 'c:spPr', 'c:txPr',
  'c:externalData', 'c:printSettings', 'c:userShapes', 'c:extLst',
];

/**
 * Give a chart a new embedded .xlsx workbook in place of the data source it had (an old OLE object,
 * a missing part, or none), pointing its c:externalData at it.
 */
export async function replaceWorkbook(doc: DocxDocument, chartPart: string, chartSpace: XmlElement, oldRelId: string | undefined, bytes: Uint8Array): Promise<void> {
  const { pkg } = doc;
  const relId = await addWorkbook(pkg, chartPart, bytes);
  const old = oldRelId === undefined ? undefined : await pkg.relationshipById(chartPart, oldRelId);
  if (old) {
    await pkg.removeRelationship(chartPart, old.id);
    const target = pkg.resolve(chartPart, old);
    if (!old.external && pkg.has(target)) pkg.removePart(target);
  }
  let externalData = firstChild(chartSpace, 'c:externalData');
  if (!externalData) {
    externalData = el('c:externalData', {}, [el('c:autoUpdate', { val: 0 })]);
    insertOrdered(chartSpace, externalData, CHARTSPACE_ORDER);
  }
  externalData.attrs['r:id'] = relId;
  if (chartSpace.attrs['xmlns:r'] === undefined) chartSpace.attrs['xmlns:r'] = NS.r;
  await doc.refreshRels(chartPart);
}

/**
 * Copy a part and the parts its relationships point to inside the package (a chart's workbook,
 * styles and colours), so the copy shares nothing it could change; pictures are shared. References
 * to relationships (r:id, r:embed, …) in the copied XML are renumbered to the copy's. Returns the
 * copy's name.
 */
export async function copyPart(pkg: OpcPackage, name: string, copies = new Map<string, string>()): Promise<string> {
  const done = copies.get(name);
  if (done) return done;
  // chart1.xml → chartN.xml, Microsoft_Excel_Worksheet.xlsx → Microsoft_Excel_WorksheetN.xlsx.
  const [, base = name, ext = ''] = /^(.*?)\d*(\.[^./]*)?$/.exec(name) ?? [];
  const copy = freePartName(pkg, base, ext);
  copies.set(name, copy);
  const type = pkg.contentType(name);
  const isXml = type !== undefined && /[+/]xml$/.test(type);
  let root: XmlElement | undefined;
  if (isXml) {
    const xml = await pkg.getXml(name);
    root = cloneElement(xml.root);
    pkg.setXml(copy, { declaration: xml.declaration, prolog: [...xml.prolog], root, epilog: [...xml.epilog] }, type);
  } else {
    pkg.setBytes(copy, (await pkg.getBytes(name)).slice(), type);
  }
  const ids = new Map<string, string>();
  for (const rel of [...(await pkg.relationships(name))]) {
    let target = rel.target;
    if (!rel.external) {
      target = pkg.resolve(name, rel);
      if (pkg.has(target) && !pkg.contentType(target)?.startsWith('image/')) target = await copyPart(pkg, target, copies);
    }
    ids.set(rel.id, await pkg.addRelationship(copy, rel.type, target, rel.external));
  }
  if (root && ids.size > 0) {
    walk(root, (e) => {
      for (const [key, value] of Object.entries(e.attrs)) {
        const id = key.startsWith('r:') ? ids.get(value) : undefined;
        if (id !== undefined) e.attrs[key] = id;
      }
    });
  }
  return copy;
}
