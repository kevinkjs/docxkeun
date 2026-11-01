/**
 * Typed run / paragraph / table / section properties parsed from w:rPr, w:pPr, w:tblPr, w:trPr,
 * w:tcPr and w:sectPr. All values are partial: undefined means "not specified at this level".
 */
import { attr, elements, firstChild, type XmlElement } from '../xml/index.js';
import {
  eighthPtToPt,
  halfPtToPt,
  parseIntSafe,
  parseOnOff,
  parseTableWidth,
  parseTwipsMeasure,
  type TableWidth,
} from './units.js';

// ---------------------------------------------------------------------------------------------
// Shared pieces

export interface Border {
  style: string; // single, double, dotted, dashed, thick, nil, none, ...
  width: number; // pt
  space: number; // pt
  color: string; // hex RRGGBB or 'auto'
}

export interface Shading {
  pattern: string;
  color: string;
  fill: string;
}

export function parseBorder(node: XmlElement | undefined): Border | undefined {
  if (!node) return undefined;
  const style = attr(node, 'w:val') ?? 'single';
  const color = attr(node, 'w:color');
  return {
    style,
    width: eighthPtToPt(parseIntSafe(attr(node, 'w:sz')) ?? (style === 'nil' || style === 'none' ? 0 : 4)),
    space: parseIntSafe(attr(node, 'w:space')) ?? 0,
    color: !color || color.toLowerCase() === 'auto' ? 'auto' : color,
  };
}

export function isVisibleBorder(b: Border | undefined): b is Border {
  return !!b && b.style !== 'nil' && b.style !== 'none' && b.width > 0;
}

export function parseShading(node: XmlElement | undefined): Shading | undefined {
  if (!node) return undefined;
  return {
    pattern: attr(node, 'w:val') ?? 'clear',
    color: attr(node, 'w:color') ?? 'auto',
    fill: attr(node, 'w:fill') ?? 'auto',
  };
}

/** Effective background colour of a shading, or undefined for none. */
export function shadingColor(s: Shading | undefined): string | undefined {
  if (!s) return undefined;
  if (s.pattern === 'solid' && s.color !== 'auto') return s.color;
  if (s.fill && s.fill.toLowerCase() !== 'auto') return s.fill;
  return undefined;
}

function onOff(parent: XmlElement, name: string): boolean | undefined {
  const node = firstChild(parent, name);
  return node ? parseOnOff(attr(node, 'w:val')) : undefined;
}

function val(parent: XmlElement, name: string): string | undefined {
  return attr(firstChild(parent, name), 'w:val');
}

function assignDefined<T extends object>(target: T, source: Partial<T> | undefined): T {
  if (!source) return target;
  for (const key of Object.keys(source) as Array<keyof T>) {
    const v = source[key];
    if (v !== undefined) target[key] = v as T[keyof T];
  }
  return target;
}

// ---------------------------------------------------------------------------------------------
// Run properties

export interface RunFonts {
  ascii?: string;
  hAnsi?: string;
  eastAsia?: string;
  cs?: string;
  asciiTheme?: string;
  hAnsiTheme?: string;
  eastAsiaTheme?: string;
  csTheme?: string;
  hint?: string;
}

export interface RunProps {
  styleId?: string;
  fonts?: RunFonts;
  size?: number; // pt
  sizeCs?: number;
  bold?: boolean;
  boldCs?: boolean;
  italic?: boolean;
  italicCs?: boolean;
  underline?: string; // single, double, none, ...
  underlineColor?: string;
  strike?: boolean;
  dstrike?: boolean;
  caps?: boolean;
  smallCaps?: boolean;
  vanish?: boolean;
  specVanish?: boolean;
  color?: string; // hex or 'auto'
  themeColor?: string;
  themeTint?: string;
  themeShade?: string;
  highlight?: string;
  shading?: Shading;
  vertAlign?: 'superscript' | 'subscript' | 'baseline';
  position?: number; // pt (raised/lowered)
  spacing?: number; // pt between characters
  scale?: number; // percent, w:w
  kern?: number;
  lang?: string;
  rtl?: boolean;
  border?: Border;
  emboss?: boolean;
  imprint?: boolean;
  outline?: boolean;
  shadow?: boolean;
  noProof?: boolean;
  webHidden?: boolean;
}

/** Properties toggled (XOR) across style levels, ECMA-376 §17.7.3. */
export const TOGGLE_PROPS = ['bold', 'boldCs', 'italic', 'italicCs', 'caps', 'smallCaps', 'strike', 'dstrike', 'outline', 'shadow', 'emboss', 'imprint', 'vanish'] as const;

export function parseRunProps(rPr: XmlElement | undefined): RunProps {
  const p: RunProps = {};
  if (!rPr) return p;
  p.styleId = val(rPr, 'w:rStyle');
  const f = firstChild(rPr, 'w:rFonts');
  if (f) {
    const fonts: RunFonts = {
      ascii: attr(f, 'w:ascii'),
      hAnsi: attr(f, 'w:hAnsi'),
      eastAsia: attr(f, 'w:eastAsia'),
      cs: attr(f, 'w:cs'),
      asciiTheme: attr(f, 'w:asciiTheme'),
      hAnsiTheme: attr(f, 'w:hAnsiTheme'),
      eastAsiaTheme: attr(f, 'w:eastAsiaTheme'),
      csTheme: attr(f, 'w:cstheme'),
      hint: attr(f, 'w:hint'),
    };
    p.fonts = assignDefined({}, fonts);
  }
  const sz = parseIntSafe(val(rPr, 'w:sz'));
  if (sz !== undefined) p.size = halfPtToPt(sz);
  const szCs = parseIntSafe(val(rPr, 'w:szCs'));
  if (szCs !== undefined) p.sizeCs = halfPtToPt(szCs);
  p.bold = onOff(rPr, 'w:b');
  p.boldCs = onOff(rPr, 'w:bCs');
  p.italic = onOff(rPr, 'w:i');
  p.italicCs = onOff(rPr, 'w:iCs');
  const u = firstChild(rPr, 'w:u');
  if (u) {
    p.underline = attr(u, 'w:val') ?? 'single';
    p.underlineColor = attr(u, 'w:color');
  }
  p.strike = onOff(rPr, 'w:strike');
  p.dstrike = onOff(rPr, 'w:dstrike');
  p.caps = onOff(rPr, 'w:caps');
  p.smallCaps = onOff(rPr, 'w:smallCaps');
  p.vanish = onOff(rPr, 'w:vanish');
  p.specVanish = onOff(rPr, 'w:specVanish');
  p.webHidden = onOff(rPr, 'w:webHidden');
  const color = firstChild(rPr, 'w:color');
  if (color) {
    p.color = attr(color, 'w:val');
    p.themeColor = attr(color, 'w:themeColor');
    p.themeTint = attr(color, 'w:themeTint');
    p.themeShade = attr(color, 'w:themeShade');
  }
  p.highlight = val(rPr, 'w:highlight');
  p.shading = parseShading(firstChild(rPr, 'w:shd'));
  const va = val(rPr, 'w:vertAlign');
  if (va === 'superscript' || va === 'subscript' || va === 'baseline') p.vertAlign = va;
  const pos = parseIntSafe(val(rPr, 'w:position'));
  if (pos !== undefined) p.position = halfPtToPt(pos);
  p.spacing = parseTwipsMeasure(val(rPr, 'w:spacing'));
  p.scale = parseIntSafe(val(rPr, 'w:w'));
  const kern = parseIntSafe(val(rPr, 'w:kern'));
  if (kern !== undefined) p.kern = halfPtToPt(kern);
  p.lang = val(rPr, 'w:lang');
  p.rtl = onOff(rPr, 'w:rtl');
  p.border = parseBorder(firstChild(rPr, 'w:bdr'));
  p.emboss = onOff(rPr, 'w:emboss');
  p.imprint = onOff(rPr, 'w:imprint');
  p.outline = onOff(rPr, 'w:outline');
  p.shadow = onOff(rPr, 'w:shadow');
  p.noProof = onOff(rPr, 'w:noProof');
  return assignDefined({} as RunProps, p);
}

/** Later levels override earlier ones (no toggle semantics; see StyleSheet for those). */
export function mergeRunProps(...layers: Array<RunProps | undefined>): RunProps {
  const out: RunProps = {};
  for (const layer of layers) {
    if (!layer) continue;
    const { fonts, ...rest } = layer;
    assignDefined(out, rest);
    if (fonts) out.fonts = assignDefined({ ...(out.fonts ?? {}) }, fonts);
  }
  return out;
}

// ---------------------------------------------------------------------------------------------
// Paragraph properties

export interface Spacing {
  before?: number;
  after?: number;
  /** For lineRule auto: a multiple (1 = single). Otherwise points. */
  line?: number;
  lineRule?: 'auto' | 'exact' | 'atLeast';
  beforeAutospacing?: boolean;
  afterAutospacing?: boolean;
}

export interface Indent {
  left?: number;
  right?: number;
  firstLine?: number;
  hanging?: number;
}

export interface TabStop {
  pos: number; // pt from the paragraph's left indent origin (page text area)
  align: 'left' | 'center' | 'right' | 'decimal' | 'bar' | 'clear' | 'num';
  leader: 'none' | 'dot' | 'hyphen' | 'underscore' | 'heavy' | 'middleDot';
}

export interface ParagraphBorders {
  top?: Border;
  bottom?: Border;
  left?: Border;
  right?: Border;
  between?: Border;
}

export type Justification = 'left' | 'center' | 'right' | 'both' | 'distribute';

export interface ParagraphProps {
  styleId?: string;
  keepNext?: boolean;
  keepLines?: boolean;
  pageBreakBefore?: boolean;
  widowControl?: boolean;
  numId?: string;
  ilvl?: number;
  spacing?: Spacing;
  indent?: Indent;
  jc?: Justification;
  tabs?: TabStop[];
  borders?: ParagraphBorders;
  shading?: Shading;
  contextualSpacing?: boolean;
  outlineLevel?: number;
  bidi?: boolean;
  suppressAutoHyphens?: boolean;
}

function parseJc(v: string | undefined): Justification | undefined {
  switch (v) {
    case 'left':
    case 'start':
      return 'left';
    case 'right':
    case 'end':
      return 'right';
    case 'center':
      return 'center';
    case 'both':
    case 'lowKashida':
    case 'mediumKashida':
    case 'highKashida':
    case 'thaiDistribute':
      return 'both';
    case 'distribute':
      return 'distribute';
    default:
      return undefined;
  }
}

export function parseParagraphProps(pPr: XmlElement | undefined): ParagraphProps {
  const p: ParagraphProps = {};
  if (!pPr) return p;
  p.styleId = val(pPr, 'w:pStyle');
  p.keepNext = onOff(pPr, 'w:keepNext');
  p.keepLines = onOff(pPr, 'w:keepLines');
  p.pageBreakBefore = onOff(pPr, 'w:pageBreakBefore');
  p.widowControl = onOff(pPr, 'w:widowControl');
  p.contextualSpacing = onOff(pPr, 'w:contextualSpacing');
  p.bidi = onOff(pPr, 'w:bidi');
  p.suppressAutoHyphens = onOff(pPr, 'w:suppressAutoHyphens');
  const numPr = firstChild(pPr, 'w:numPr');
  if (numPr) {
    p.numId = val(numPr, 'w:numId');
    p.ilvl = parseIntSafe(val(numPr, 'w:ilvl'));
  }
  const sp = firstChild(pPr, 'w:spacing');
  if (sp) {
    const spacing: Spacing = {
      before: parseTwipsMeasure(attr(sp, 'w:before')),
      after: parseTwipsMeasure(attr(sp, 'w:after')),
      beforeAutospacing: attr(sp, 'w:beforeAutospacing') !== undefined ? parseOnOff(attr(sp, 'w:beforeAutospacing')) : undefined,
      afterAutospacing: attr(sp, 'w:afterAutospacing') !== undefined ? parseOnOff(attr(sp, 'w:afterAutospacing')) : undefined,
    };
    const line = parseIntSafe(attr(sp, 'w:line'));
    const rule = attr(sp, 'w:lineRule');
    if (line !== undefined) {
      if (rule === 'exact' || rule === 'atLeast') {
        spacing.lineRule = rule;
        spacing.line = line / 20;
      } else {
        spacing.lineRule = 'auto';
        spacing.line = line / 240;
      }
    }
    p.spacing = assignDefined({}, spacing);
  }
  const ind = firstChild(pPr, 'w:ind');
  if (ind) {
    p.indent = assignDefined({} as Indent, {
      left: parseTwipsMeasure(attr(ind, 'w:left') ?? attr(ind, 'w:start')),
      right: parseTwipsMeasure(attr(ind, 'w:right') ?? attr(ind, 'w:end')),
      firstLine: parseTwipsMeasure(attr(ind, 'w:firstLine')),
      hanging: parseTwipsMeasure(attr(ind, 'w:hanging')),
    });
  }
  p.jc = parseJc(val(pPr, 'w:jc'));
  const tabs = firstChild(pPr, 'w:tabs');
  if (tabs) {
    p.tabs = elements(tabs, 'w:tab').map((t) => ({
      pos: parseTwipsMeasure(attr(t, 'w:pos')) ?? 0,
      align: normaliseTabAlign(attr(t, 'w:val')),
      leader: (attr(t, 'w:leader') ?? 'none') as TabStop['leader'],
    }));
  }
  const bdr = firstChild(pPr, 'w:pBdr');
  if (bdr) {
    p.borders = assignDefined({} as ParagraphBorders, {
      top: parseBorder(firstChild(bdr, 'w:top')),
      bottom: parseBorder(firstChild(bdr, 'w:bottom')),
      left: parseBorder(firstChild(bdr, 'w:left') ?? firstChild(bdr, 'w:start')),
      right: parseBorder(firstChild(bdr, 'w:right') ?? firstChild(bdr, 'w:end')),
      between: parseBorder(firstChild(bdr, 'w:between')),
    });
  }
  p.shading = parseShading(firstChild(pPr, 'w:shd'));
  p.outlineLevel = parseIntSafe(val(pPr, 'w:outlineLvl'));
  return assignDefined({} as ParagraphProps, p);
}

function normaliseTabAlign(v: string | undefined): TabStop['align'] {
  switch (v) {
    case 'start':
      return 'left';
    case 'end':
      return 'right';
    case 'center':
    case 'right':
    case 'decimal':
    case 'bar':
    case 'clear':
    case 'num':
    case 'left':
      return v;
    default:
      return 'left';
  }
}

/** Merge tab stop lists: later lists add stops and `clear` removes stops at the same position. */
export function mergeTabs(base: TabStop[] | undefined, next: TabStop[] | undefined): TabStop[] | undefined {
  if (!next) return base;
  const out = [...(base ?? [])];
  for (const t of next) {
    const i = out.findIndex((o) => Math.abs(o.pos - t.pos) < 0.05);
    if (i >= 0) out.splice(i, 1);
    if (t.align !== 'clear') out.push(t);
  }
  return out.sort((a, b) => a.pos - b.pos);
}

export function mergeParagraphProps(...layers: Array<ParagraphProps | undefined>): ParagraphProps {
  const out: ParagraphProps = {};
  for (const layer of layers) {
    if (!layer) continue;
    const { spacing, indent, tabs, borders, ...rest } = layer;
    assignDefined(out, rest);
    if (spacing) out.spacing = assignDefined({ ...(out.spacing ?? {}) }, spacing);
    if (indent) {
      const merged = { ...(out.indent ?? {}) };
      // firstLine and hanging are mutually exclusive: the later level wins.
      if (indent.firstLine !== undefined) delete merged.hanging;
      if (indent.hanging !== undefined) delete merged.firstLine;
      out.indent = assignDefined(merged, indent);
    }
    if (tabs) out.tabs = mergeTabs(out.tabs, tabs);
    if (borders) out.borders = assignDefined({ ...(out.borders ?? {}) }, borders);
  }
  return out;
}

// ---------------------------------------------------------------------------------------------
// Tables

export interface CellMargins {
  top?: number;
  left?: number;
  bottom?: number;
  right?: number;
}

export interface TableBorders {
  top?: Border;
  left?: Border;
  bottom?: Border;
  right?: Border;
  insideH?: Border;
  insideV?: Border;
}

export interface TableLook {
  firstRow: boolean;
  lastRow: boolean;
  firstColumn: boolean;
  lastColumn: boolean;
  noHBand: boolean;
  noVBand: boolean;
}

export interface TableProps {
  styleId?: string;
  width?: TableWidth;
  jc?: 'left' | 'center' | 'right';
  indent?: number;
  borders?: TableBorders;
  cellMargins?: CellMargins;
  layout?: 'fixed' | 'autofit';
  shading?: Shading;
  look?: TableLook;
  cellSpacing?: number;
  rowBandSize?: number;
  colBandSize?: number;
}

function parseCellMargins(node: XmlElement | undefined): CellMargins | undefined {
  if (!node) return undefined;
  const side = (name: string, alt?: string): number | undefined => {
    const n = firstChild(node, name) ?? (alt ? firstChild(node, alt) : undefined);
    if (!n) return undefined;
    const w = parseTableWidth(attr(n, 'w:w'), attr(n, 'w:type'));
    return w && w.type === 'dxa' ? w.value : 0;
  };
  return assignDefined({} as CellMargins, { top: side('w:top'), left: side('w:left', 'w:start'), bottom: side('w:bottom'), right: side('w:right', 'w:end') });
}

function parseTableBorders(node: XmlElement | undefined): TableBorders | undefined {
  if (!node) return undefined;
  return assignDefined({} as TableBorders, {
    top: parseBorder(firstChild(node, 'w:top')),
    left: parseBorder(firstChild(node, 'w:left') ?? firstChild(node, 'w:start')),
    bottom: parseBorder(firstChild(node, 'w:bottom')),
    right: parseBorder(firstChild(node, 'w:right') ?? firstChild(node, 'w:end')),
    insideH: parseBorder(firstChild(node, 'w:insideH')),
    insideV: parseBorder(firstChild(node, 'w:insideV')),
  });
}

function parseTableLook(node: XmlElement | undefined): TableLook | undefined {
  if (!node) return undefined;
  const hex = attr(node, 'w:val');
  if (hex && attr(node, 'w:firstRow') === undefined) {
    const v = Number.parseInt(hex, 16);
    return {
      firstRow: !!(v & 0x20),
      lastRow: !!(v & 0x40),
      firstColumn: !!(v & 0x80),
      lastColumn: !!(v & 0x100),
      noHBand: !!(v & 0x200),
      noVBand: !!(v & 0x400),
    };
  }
  const flag = (n: string, d: boolean) => (attr(node, n) === undefined ? d : parseOnOff(attr(node, n)));
  return {
    firstRow: flag('w:firstRow', true),
    lastRow: flag('w:lastRow', false),
    firstColumn: flag('w:firstColumn', true),
    lastColumn: flag('w:lastColumn', false),
    noHBand: flag('w:noHBand', false),
    noVBand: flag('w:noVBand', true),
  };
}

export function parseTableProps(tblPr: XmlElement | undefined): TableProps {
  const p: TableProps = {};
  if (!tblPr) return p;
  p.styleId = val(tblPr, 'w:tblStyle');
  const w = firstChild(tblPr, 'w:tblW');
  if (w) p.width = parseTableWidth(attr(w, 'w:w'), attr(w, 'w:type'));
  const jc = val(tblPr, 'w:jc');
  if (jc) p.jc = jc === 'center' ? 'center' : jc === 'right' || jc === 'end' ? 'right' : 'left';
  const ind = firstChild(tblPr, 'w:tblInd');
  if (ind) p.indent = parseTableWidth(attr(ind, 'w:w'), attr(ind, 'w:type'))?.value;
  p.borders = parseTableBorders(firstChild(tblPr, 'w:tblBorders'));
  p.cellMargins = parseCellMargins(firstChild(tblPr, 'w:tblCellMar'));
  const layout = attr(firstChild(tblPr, 'w:tblLayout'), 'w:type');
  if (layout) p.layout = layout === 'fixed' ? 'fixed' : 'autofit';
  p.shading = parseShading(firstChild(tblPr, 'w:shd'));
  p.look = parseTableLook(firstChild(tblPr, 'w:tblLook'));
  const spacing = firstChild(tblPr, 'w:tblCellSpacing');
  if (spacing) p.cellSpacing = parseTableWidth(attr(spacing, 'w:w'), attr(spacing, 'w:type'))?.value;
  p.rowBandSize = parseIntSafe(val(tblPr, 'w:tblStyleRowBandSize'));
  p.colBandSize = parseIntSafe(val(tblPr, 'w:tblStyleColBandSize'));
  return assignDefined({} as TableProps, p);
}

export function mergeTableProps(...layers: Array<TableProps | undefined>): TableProps {
  const out: TableProps = {};
  for (const layer of layers) {
    if (!layer) continue;
    const { borders, cellMargins, ...rest } = layer;
    assignDefined(out, rest);
    if (borders) out.borders = assignDefined({ ...(out.borders ?? {}) }, borders);
    if (cellMargins) out.cellMargins = assignDefined({ ...(out.cellMargins ?? {}) }, cellMargins);
  }
  return out;
}

export interface RowProps {
  height?: number;
  heightRule?: 'auto' | 'atLeast' | 'exact';
  isHeader?: boolean;
  cantSplit?: boolean;
  gridBefore?: number;
  gridAfter?: number;
  jc?: 'left' | 'center' | 'right';
}

export function parseRowProps(trPr: XmlElement | undefined): RowProps {
  const p: RowProps = {};
  if (!trPr) return p;
  const h = firstChild(trPr, 'w:trHeight');
  if (h) {
    p.height = parseTwipsMeasure(attr(h, 'w:val'));
    const rule = attr(h, 'w:hRule');
    p.heightRule = rule === 'exact' ? 'exact' : rule === 'auto' ? 'auto' : 'atLeast';
  }
  p.isHeader = onOff(trPr, 'w:tblHeader');
  p.cantSplit = onOff(trPr, 'w:cantSplit');
  p.gridBefore = parseIntSafe(val(trPr, 'w:gridBefore'));
  p.gridAfter = parseIntSafe(val(trPr, 'w:gridAfter'));
  return assignDefined({} as RowProps, p);
}

export interface CellProps {
  width?: TableWidth;
  gridSpan?: number;
  vMerge?: 'restart' | 'continue';
  hMerge?: 'restart' | 'continue';
  borders?: TableBorders & { tl2br?: Border; tr2bl?: Border };
  shading?: Shading;
  margins?: CellMargins;
  vAlign?: 'top' | 'center' | 'bottom';
  noWrap?: boolean;
  textDirection?: string;
}

export function parseCellProps(tcPr: XmlElement | undefined): CellProps {
  const p: CellProps = {};
  if (!tcPr) return p;
  const w = firstChild(tcPr, 'w:tcW');
  if (w) p.width = parseTableWidth(attr(w, 'w:w'), attr(w, 'w:type'));
  p.gridSpan = parseIntSafe(val(tcPr, 'w:gridSpan'));
  const vm = firstChild(tcPr, 'w:vMerge');
  if (vm) p.vMerge = attr(vm, 'w:val') === 'restart' ? 'restart' : 'continue';
  const hm = firstChild(tcPr, 'w:hMerge');
  if (hm) p.hMerge = attr(hm, 'w:val') === 'restart' ? 'restart' : 'continue';
  p.borders = parseTableBorders(firstChild(tcPr, 'w:tcBorders'));
  p.shading = parseShading(firstChild(tcPr, 'w:shd'));
  p.margins = parseCellMargins(firstChild(tcPr, 'w:tcMar'));
  const va = val(tcPr, 'w:vAlign');
  if (va) p.vAlign = va === 'center' ? 'center' : va === 'bottom' ? 'bottom' : 'top';
  p.noWrap = onOff(tcPr, 'w:noWrap');
  p.textDirection = val(tcPr, 'w:textDirection');
  return assignDefined({} as CellProps, p);
}

export function mergeCellProps(...layers: Array<CellProps | undefined>): CellProps {
  const out: CellProps = {};
  for (const layer of layers) {
    if (!layer) continue;
    const { borders, margins, ...rest } = layer;
    assignDefined(out, rest);
    if (borders) out.borders = assignDefined({ ...(out.borders ?? {}) }, borders);
    if (margins) out.margins = assignDefined({ ...(out.margins ?? {}) }, margins);
  }
  return out;
}

// ---------------------------------------------------------------------------------------------
// Sections

export interface HeaderFooterRef {
  type: 'default' | 'first' | 'even';
  rId: string;
}

export interface SectionProps {
  pageWidth: number;
  pageHeight: number;
  orientation: 'portrait' | 'landscape';
  margin: { top: number; right: number; bottom: number; left: number; header: number; footer: number; gutter: number };
  columns: { count: number; space: number; equalWidth: boolean; separator: boolean; widths: Array<{ width: number; space: number }> };
  type: 'nextPage' | 'continuous' | 'evenPage' | 'oddPage' | 'nextColumn';
  titlePage: boolean;
  headers: HeaderFooterRef[];
  footers: HeaderFooterRef[];
  pageNumberStart?: number;
  pageNumberFormat?: string;
  vAlign: 'top' | 'center' | 'both' | 'bottom';
  bidi: boolean;
}

export const DEFAULT_SECTION: SectionProps = {
  pageWidth: 612,
  pageHeight: 792,
  orientation: 'portrait',
  margin: { top: 72, right: 72, bottom: 72, left: 72, header: 36, footer: 36, gutter: 0 },
  columns: { count: 1, space: 36, equalWidth: true, separator: false, widths: [] },
  type: 'nextPage',
  titlePage: false,
  headers: [],
  footers: [],
  vAlign: 'top',
  bidi: false,
};

export function parseSectionProps(sectPr: XmlElement | undefined): SectionProps {
  const s: SectionProps = structuredClone(DEFAULT_SECTION);
  if (!sectPr) return s;
  const pgSz = firstChild(sectPr, 'w:pgSz');
  if (pgSz) {
    s.pageWidth = parseTwipsMeasure(attr(pgSz, 'w:w')) ?? s.pageWidth;
    s.pageHeight = parseTwipsMeasure(attr(pgSz, 'w:h')) ?? s.pageHeight;
    s.orientation = attr(pgSz, 'w:orient') === 'landscape' ? 'landscape' : 'portrait';
  }
  const pgMar = firstChild(sectPr, 'w:pgMar');
  if (pgMar) {
    const m = s.margin;
    m.top = parseTwipsMeasure(attr(pgMar, 'w:top')) ?? m.top;
    m.right = parseTwipsMeasure(attr(pgMar, 'w:right')) ?? m.right;
    m.bottom = parseTwipsMeasure(attr(pgMar, 'w:bottom')) ?? m.bottom;
    m.left = parseTwipsMeasure(attr(pgMar, 'w:left')) ?? m.left;
    m.header = parseTwipsMeasure(attr(pgMar, 'w:header')) ?? m.header;
    m.footer = parseTwipsMeasure(attr(pgMar, 'w:footer')) ?? m.footer;
    m.gutter = parseTwipsMeasure(attr(pgMar, 'w:gutter')) ?? m.gutter;
  }
  const cols = firstChild(sectPr, 'w:cols');
  if (cols) {
    s.columns = {
      count: Math.max(1, parseIntSafe(attr(cols, 'w:num')) ?? 1),
      space: parseTwipsMeasure(attr(cols, 'w:space')) ?? 36,
      equalWidth: attr(cols, 'w:equalWidth') === undefined ? true : parseOnOff(attr(cols, 'w:equalWidth')),
      separator: attr(cols, 'w:sep') !== undefined && parseOnOff(attr(cols, 'w:sep')),
      widths: elements(cols, 'w:col').map((c) => ({ width: parseTwipsMeasure(attr(c, 'w:w')) ?? 0, space: parseTwipsMeasure(attr(c, 'w:space')) ?? 0 })),
    };
    if (!s.columns.equalWidth && s.columns.widths.length > 0) s.columns.count = s.columns.widths.length;
  }
  const type = val(sectPr, 'w:type');
  if (type === 'continuous' || type === 'evenPage' || type === 'oddPage' || type === 'nextColumn' || type === 'nextPage') s.type = type;
  s.titlePage = onOff(sectPr, 'w:titlePg') ?? false;
  const refs = (name: string): HeaderFooterRef[] =>
    elements(sectPr, name)
      .map((r) => ({ type: (attr(r, 'w:type') ?? 'default') as HeaderFooterRef['type'], rId: attr(r, 'r:id') ?? '' }))
      .filter((r) => r.rId);
  s.headers = refs('w:headerReference');
  s.footers = refs('w:footerReference');
  const pgNum = firstChild(sectPr, 'w:pgNumType');
  if (pgNum) {
    s.pageNumberStart = parseIntSafe(attr(pgNum, 'w:start'));
    s.pageNumberFormat = attr(pgNum, 'w:fmt');
  }
  const va = val(sectPr, 'w:vAlign');
  if (va === 'center' || va === 'both' || va === 'bottom') s.vAlign = va;
  s.bidi = onOff(sectPr, 'w:bidi') ?? false;
  return s;
}
