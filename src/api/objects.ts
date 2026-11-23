/**
 * Objects in a story (pictures, charts, SmartArt, shapes, text boxes, groups): listing them in
 * document order, the drawing XML for new ones, resizing and placing existing ones, and the parts
 * behind a chart (its workbook, and copies of a chart part for copied content).
 */
import type { ChartTable, NewChart, NewChartType } from '../charts/write.js';
import type { DocParagraph, ImageInline } from '../model/content.js';
import type { DocxDocument } from '../model/document.js';
import { CONTENT_TYPE, REL, type OpcPackage } from '../opc/index.js';
import { attr, childPath, cloneElement, el, elements, findAll, firstChild, insertOrdered, NS, removeChildren, text, walk, type XmlElement } from '../xml/index.js';
import type { TextFormat } from './editor.js';
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

export type ShapeKind =
  | 'textBox'
  | 'rect'
  | 'roundRect'
  | 'ellipse'
  | 'triangle'
  | 'diamond'
  | 'rightArrow'
  | 'leftArrow'
  | 'upArrow'
  | 'downArrow'
  | 'chevron'
  | 'pentagon'
  | 'star5'
  | 'wedgeRectCallout'
  | 'wedgeRoundRectCallout'
  | 'cloudCallout'
  | 'line';

/**
 * Inline in the text, or floating at (x, y) points from the paragraph (default), margin or page,
 * with text wrapped square around it (default), above and below it, or not at all (in front of or
 * behind the text). An object that already floats keeps what is left out (its offsets, reference
 * and wrap); a new floating object, or one that was inline, needs x and y.
 */
export type ObjectPlacement = { inline: true } | { x?: number; y?: number; from?: 'margin' | 'page' | 'paragraph'; wrap?: 'square' | 'topAndBottom' | 'inFront' | 'behind' };
export type FloatingPlacement = Exclude<ObjectPlacement, { inline: true }>;

/** A new shape or text box (sizes in points; colours RRGGBB or 'none'). */
export interface NewShape {
  shape: ShapeKind;
  /** Text inside: each line becomes a paragraph. */
  text?: string;
  /** Default: text box 180 × 72, line 144 × 0, other shapes 108 × 72. */
  width?: number;
  height?: number;
  /** Default: white for a text box, the theme's accent 1 for a shape. */
  fill?: string;
  /** Default: thin black for a text box, a darker accent 1 for a shape. */
  outline?: string;
  /** Default: floating at the top left of the paragraph, square wrap. */
  place?: ObjectPlacement;
  alt?: string;
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
   * also true when a workbook that can't be read (an old .xls object, broken bytes) was replaced.
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

/** Where a floating object sits: relativeFrom of wp:positionH and wp:positionV. */
const RELATIVE_FROM: Record<'margin' | 'page' | 'paragraph', [string, string]> = { margin: ['margin', 'margin'], page: ['page', 'page'], paragraph: ['column', 'paragraph'] };
const WRAPS = ['square', 'topAndBottom', 'inFront', 'behind'];

/**
 * Check a placement's position, reference and wrap style (EditError when wrong). x and y may be left
 * out only for an object that already floats (`floating`), which keeps its own.
 */
export function checkPlacement(place: ObjectPlacement, floating = false): void {
  if ('inline' in place) return;
  for (const side of [place.x, place.y]) {
    if (side === undefined ? !floating : !Number.isFinite(side)) throw new EditError('A floating object needs x and y in points');
  }
  if (place.from !== undefined && !Object.hasOwn(RELATIVE_FROM, place.from)) throw new EditError(`Place an object from the margin, page or paragraph, not "${place.from}"`);
  if (place.wrap !== undefined && !WRAPS.includes(place.wrap)) throw new EditError(`Wrap text square, topAndBottom, inFront or behind, not "${place.wrap}"`);
}

/** The z-order for a new floating object: above every other one in `roots` (Word starts at 251659264). */
export function nextRelativeHeight(roots: XmlElement[]): number {
  let top = 0;
  for (const root of roots) for (const anchor of findAll(root, 'wp:anchor')) top = Math.max(top, Number(attr(anchor, 'relativeHeight')) || 0);
  return Math.max(top + 1, 251659264);
}

/** What an inline or floating frame is built around: its size, name and alt text, and graphic. */
interface FrameParts {
  extent: XmlElement;
  effectExtent: XmlElement | undefined;
  docPr: XmlElement;
  cNvGraphicFramePr: XmlElement | undefined;
  graphic: XmlElement;
}

/** wp:anchor's children in schema order (wp14 relative sizes last). */
const ANCHOR_ORDER = [
  'wp:simplePos', 'wp:positionH', 'wp:positionV', 'wp:extent', 'wp:effectExtent', 'wp:wrapNone', 'wp:wrapSquare', 'wp:wrapTight', 'wp:wrapThrough', 'wp:wrapTopAndBottom',
  'wp:docPr', 'wp:cNvGraphicFramePr', 'a:graphic', 'wp14:sizeRelH', 'wp14:sizeRelV',
];
const WRAP_ELEMENTS = new Set(['wp:wrapNone', 'wp:wrapSquare', 'wp:wrapTight', 'wp:wrapThrough', 'wp:wrapTopAndBottom']);

const wrapElement = (wrap: NonNullable<FloatingPlacement['wrap']>): XmlElement =>
  wrap === 'square' ? el('wp:wrapSquare', { wrapText: 'bothSides' }) : wrap === 'topAndBottom' ? el('wp:wrapTopAndBottom') : el('wp:wrapNone');

const positionElement = (axis: 'H' | 'V', relativeFrom: string, offset: number): XmlElement =>
  el(`wp:position${axis}`, { relativeFrom }, [el('wp:posOffset', {}, [text(emu(offset))])]);

const optional = (e: XmlElement | undefined): XmlElement[] => (e ? [e] : []);

/** A wp:anchor around `parts` at `place` (from the paragraph and square wrap by default), at z-order `relativeHeight`. */
function anchorOf(parts: FrameParts, place: FloatingPlacement, relativeHeight: number, attrs: Record<string, string> = {}): XmlElement {
  checkPlacement(place);
  const [h, v] = RELATIVE_FROM[place.from ?? 'paragraph'];
  const wrap = place.wrap ?? 'square';
  const frame = { ...attrs, distT: 0, distB: 0, distL: 114300, distR: 114300, simplePos: 0, relativeHeight, behindDoc: wrap === 'behind' ? 1 : 0, locked: 0, layoutInCell: 1, allowOverlap: 1 };
  return el('wp:anchor', frame, [
    el('wp:simplePos', { x: 0, y: 0 }),
    positionElement('H', h, place.x!),
    positionElement('V', v, place.y!),
    parts.extent,
    ...optional(parts.effectExtent),
    wrapElement(wrap),
    parts.docPr,
    ...optional(parts.cNvGraphicFramePr),
    parts.graphic,
  ]);
}

/**
 * A w:drawing with a floating frame (wp:anchor) of `widthPt` × `heightPt` points around
 * `graphicData`, at `place` (from the paragraph by default, square wrap by default) and at z-order
 * `relativeHeight`.
 */
export function anchorFrame(
  id: number,
  name: string,
  widthPt: number,
  heightPt: number,
  place: FloatingPlacement,
  graphicData: XmlElement,
  alt: string | undefined,
  relativeHeight: number,
): XmlElement {
  const docPr: Record<string, string> = { id: String(id), name };
  if (alt) docPr.descr = alt;
  const parts: FrameParts = {
    extent: el('wp:extent', { cx: emu(widthPt), cy: emu(heightPt) }),
    effectExtent: el('wp:effectExtent', { l: 0, t: 0, r: 0, b: 0 }),
    docPr: el('wp:docPr', docPr),
    cNvGraphicFramePr: el('wp:cNvGraphicFramePr'),
    graphic: el('a:graphic', { 'xmlns:a': NS.a }, [graphicData]),
  };
  return el('w:drawing', {}, [anchorOf(parts, place, relativeHeight)]);
}

/** The graphic of a chart whose part the relationship `relId` points at. */
export function chartGraphic(relId: string): XmlElement {
  return el('a:graphicData', { uri: 'http://schemas.openxmlformats.org/drawingml/2006/chart' }, [el('c:chart', { 'xmlns:c': NS.c, 'xmlns:r': NS.r, 'r:id': relId })]);
}

// ---------------------------------------------------------------------------------------------
// Changing objects

/** The frame (wp:inline or wp:anchor) of an object's w:drawing. */
export function frameOf(drawing: XmlElement): XmlElement {
  const frame = elements(drawing).find((e) => e.name === 'wp:inline' || e.name === 'wp:anchor');
  if (!frame) throw new EditError('This object has no frame to change');
  return frame;
}

function partsOf(frame: XmlElement): FrameParts {
  const [extent, docPr, graphic] = ['wp:extent', 'wp:docPr', 'a:graphic'].map((n) => firstChild(frame, n));
  if (!extent || !docPr || !graphic) throw new EditError('This object is missing part of its frame');
  return { extent, effectExtent: firstChild(frame, 'wp:effectExtent'), docPr, cNvGraphicFramePr: firstChild(frame, 'wp:cNvGraphicFramePr'), graphic };
}

/** The namespace declarations on an element, kept when its frame is rebuilt. */
function declarations(e: XmlElement): Record<string, string> {
  return Object.fromEntries(Object.entries(e.attrs).filter(([k]) => k === 'xmlns' || k.startsWith('xmlns:')));
}

/**
 * Word's old VML copy of an object (the mc:Fallback beside its mc:Choice) is stale once the
 * object changes, so it goes. `node` is the run's child holding the object.
 */
export function dropFallback(node: XmlElement): void {
  if (node.name === 'mc:AlternateContent') removeChildren(node, 'mc:Fallback');
}

/** An object's new size in points: as given, or with one side given, the other keeps the frame's proportions. */
export function resized(frame: XmlElement, size: { width?: number | undefined; height?: number | undefined }): { width: number; height: number } {
  const { width, height } = size;
  if (width === undefined && height === undefined) throw new EditError('Give a new width, height or both');
  for (const side of [width, height]) {
    if (side !== undefined && !(side > 0 && Number.isFinite(side))) throw new EditError('An object needs a positive width and height');
  }
  const extent = firstChild(frame, 'wp:extent');
  const cx = Number(attr(extent, 'cx')) || 0;
  const cy = Number(attr(extent, 'cy')) || 0;
  if (width !== undefined) return { width, height: height ?? (cx > 0 ? (cy * width) / cx : cy / 12700) };
  return { width: cy > 0 ? (cx * height!) / cy : cx / 12700, height: height! };
}

/** Size a frame to `widthPt` × `heightPt`, with the picture, shape or group in it (a group's children scale with it). */
export function resizeFrame(frame: XmlElement, widthPt: number, heightPt: number): void {
  const size = { cx: emu(widthPt), cy: emu(heightPt) };
  const extent = firstChild(frame, 'wp:extent');
  if (!extent) throw new EditError('This object has no size to change');
  Object.assign(extent.attrs, size);
  // A size relative to the page or margin would win over the new one.
  removeChildren(frame, 'wp14:sizeRelH');
  removeChildren(frame, 'wp14:sizeRelV');
  const graphicData = childPath(frame, 'a:graphic', 'a:graphicData');
  for (const content of graphicData ? elements(graphicData) : []) {
    const props = firstChild(content, 'pic:spPr') ?? firstChild(content, 'wps:spPr') ?? firstChild(content, 'wpg:grpSpPr');
    const ext = childPath(props, 'a:xfrm', 'a:ext');
    if (ext) Object.assign(ext.attrs, size);
  }
}

/**
 * Move a floating frame to `place`. Without `from` or `wrap` it keeps its own; a side without an
 * offset (x or y) keeps its own offset or alignment, counted from `from` when that is given. A
 * position Word stored as a percentage (wp14, inside mc:AlternateContent) is replaced when a new
 * offset is given.
 */
function moveAnchor(anchor: XmlElement, place: FloatingPlacement): void {
  const holds = (c: XmlElement, name: string) => c.name === name || (c.name === 'mc:AlternateContent' && findAll(c, name).length > 0);
  // The frame's own position elements: its child, or the ones in an mc:AlternateContent child.
  const positions = (name: string) => elements(anchor).flatMap((c) => (c.name === name ? [c] : c.name === 'mc:AlternateContent' ? findAll(c, name) : []));
  const sides = [
    ['H', place.x, 0, 'column'],
    ['V', place.y, 1, 'paragraph'],
  ] as const;
  for (const [axis, offset, i, fallback] of sides) {
    const name = `wp:position${axis}`;
    const from = place.from ? RELATIVE_FROM[place.from][i] : undefined;
    if (offset === undefined) {
      if (from) for (const position of positions(name)) position.attrs.relativeFrom = from;
      continue;
    }
    const relativeFrom = from ?? attr(positions(name)[0], 'relativeFrom') ?? fallback;
    anchor.children = anchor.children.filter((c) => !(c.type === 'element' && holds(c, name)));
    insertOrdered(anchor, positionElement(axis, relativeFrom, offset), ANCHOR_ORDER);
    // simplePos="1" would place it at wp:simplePos instead.
    anchor.attrs.simplePos = '0';
  }
  if (place.wrap) {
    anchor.children = anchor.children.filter((c) => !(c.type === 'element' && WRAP_ELEMENTS.has(c.name)));
    insertOrdered(anchor, wrapElement(place.wrap), ANCHOR_ORDER);
    anchor.attrs.behindDoc = place.wrap === 'behind' ? '1' : '0';
  }
}

/**
 * Place an object's w:drawing inline, or floating at `place`: an inline frame becomes a wp:anchor
 * (at z-order `relativeHeight()`, from the paragraph and square wrap by default) around the same
 * docPr, extent and graphic, and back; a floating one moves (see `moveAnchor`).
 */
export function placeFrame(drawing: XmlElement, place: ObjectPlacement, relativeHeight: () => number): void {
  const frame = frameOf(drawing);
  checkPlacement(place, frame.name === 'wp:anchor');
  let next: XmlElement;
  if ('inline' in place) {
    if (frame.name === 'wp:inline') return;
    const parts = partsOf(frame);
    next = el('wp:inline', { ...declarations(frame), distT: 0, distB: 0, distL: 0, distR: 0 }, [
      parts.extent,
      ...optional(parts.effectExtent),
      parts.docPr,
      ...optional(parts.cNvGraphicFramePr),
      parts.graphic,
    ]);
  } else if (frame.name === 'wp:inline') {
    next = anchorOf(partsOf(frame), place, relativeHeight(), declarations(frame));
  } else {
    moveAnchor(frame, place);
    return;
  }
  drawing.children[drawing.children.indexOf(frame)] = next;
}

/**
 * A copy of an object: of `node`, the run's child holding it (a w:drawing, or the
 * mc:AlternateContent around one), without Word's old VML copy; and the w:drawing in the copy that
 * stands where `drawing` stands in the original.
 */
export function cloneObject(idx: TreeIndex, node: XmlElement, drawing: XmlElement): { copy: XmlElement; drawing: XmlElement } {
  const path: number[] = [];
  for (let cur = drawing; cur !== node; cur = idx.parentOf(cur)) path.unshift(idx.parentOf(cur).children.indexOf(cur));
  const copy = cloneElement(node);
  let found = copy;
  for (const i of path) found = found.children[i] as XmlElement;
  dropFallback(copy);
  return { copy, drawing: found };
}

const RUN_REVISIONS = new Set(['w:ins', 'w:del', 'w:moveFrom', 'w:moveTo']);

/** Take an object's run child out of its run; a run (and a tracked change around it) left empty goes too. */
export function takeOut(idx: TreeIndex, node: XmlElement): void {
  let parent = idx.parentOf(node);
  parent.children.splice(parent.children.indexOf(node), 1);
  const empty = (e: XmlElement) => !e.children.some((c) => c.type === 'element' && c.name !== 'w:rPr');
  for (let cur = parent; (cur.name === 'w:r' || RUN_REVISIONS.has(cur.name)) && empty(cur); cur = parent) {
    parent = idx.parentOf(cur);
    parent.children.splice(parent.children.indexOf(cur), 1);
  }
}

// ---------------------------------------------------------------------------------------------
// Shapes and text boxes

/** Word's names for new shapes (the wp:docPr name, before the drawing id). */
export const SHAPE_NAMES: Record<ShapeKind, string> = {
  textBox: 'Text Box',
  rect: 'Rectangle',
  roundRect: 'Rectangle: Rounded Corners',
  ellipse: 'Oval',
  triangle: 'Isosceles Triangle',
  diamond: 'Diamond',
  rightArrow: 'Arrow: Right',
  leftArrow: 'Arrow: Left',
  upArrow: 'Arrow: Up',
  downArrow: 'Arrow: Down',
  chevron: 'Arrow: Chevron',
  pentagon: 'Regular Pentagon',
  star5: 'Star: 5 Points',
  wedgeRectCallout: 'Speech Bubble: Rectangle',
  wedgeRoundRectCallout: 'Speech Bubble: Rectangle with Corners Rounded',
  cloudCallout: 'Thought Bubble: Cloud',
  line: 'Straight Connector',
};

/** Whether a node sits inside a text box (DrawingML or VML: both keep their text in w:txbxContent). */
export function inTextBox(idx: TreeIndex, node: XmlElement): boolean {
  return idx.ancestor(node, 'w:txbxContent') !== undefined;
}

/** A colour option: RRGGBB (a leading # is fine) or 'none'. */
function checkColour(value: string | undefined, what: string): void {
  if (value !== undefined && value !== 'none' && !/^#?[0-9A-Fa-f]{6}$/.test(value)) throw new EditError(`A ${what} colour is RRGGBB (like 4472C4) or "none", not "${value}"`);
}

/**
 * Check a new shape's options (EditError when wrong) and fill in its size and placement: a text box
 * is 180 × 72 pt, a line 144 × 0 pt, other shapes 108 × 72 pt; by default it floats at the top left
 * of its paragraph.
 */
export function checkShape(shape: NewShape): { width: number; height: number; place: ObjectPlacement } {
  if (!Object.hasOwn(SHAPE_NAMES, shape.shape)) throw new EditError(`Unknown shape "${shape.shape}"`);
  const line = shape.shape === 'line';
  if (line && shape.text) throw new EditError('A line cannot hold text');
  const [w, h] = shape.shape === 'textBox' ? [180, 72] : line ? [144, 0] : [108, 72];
  const width = shape.width ?? w;
  const height = shape.height ?? h;
  // A line may be flat in one direction, not both; other shapes need both sides.
  const sized = (n: number) => Number.isFinite(n) && (line ? n >= 0 : n > 0);
  if (!sized(width) || !sized(height) || width + height === 0) throw new EditError(line ? 'A line needs a length' : 'A shape needs a positive width and height');
  checkColour(shape.fill, 'fill');
  checkColour(shape.outline, 'outline');
  const place = shape.place ?? { x: 0, y: 0 };
  checkPlacement(place);
  return { width, height, place };
}

const solid = (colour: string) => el('a:solidFill', {}, [el('a:srgbClr', { val: colour.replace(/^#/, '').toUpperCase() })]);

/**
 * The graphic of a new shape (wps:wsp) of `widthPt` × `heightPt` points, holding `shape.text` (each
 * line a paragraph, each run made by `newRun`). A text box is white with a thin black outline and
 * top-aligned text; other shapes take the theme's accent 1 with a darker outline (wps:style) and
 * white, centred text; `fill` and `outline` override either.
 */
export function shapeGraphic(shape: NewShape, widthPt: number, heightPt: number, newRun: (text: string, format?: TextFormat) => XmlElement): XmlElement {
  const box = shape.shape === 'textBox';
  const line = shape.shape === 'line';
  const spPr = el('wps:spPr', {}, [
    el('a:xfrm', {}, [el('a:off', { x: 0, y: 0 }), el('a:ext', { cx: emu(widthPt), cy: emu(heightPt) })]),
    el('a:prstGeom', { prst: box ? 'rect' : shape.shape }, [el('a:avLst')]),
  ]);
  const fill = shape.fill ?? (box ? 'FFFFFF' : undefined);
  if (fill !== undefined) spPr.children.push(fill === 'none' ? el('a:noFill') : solid(fill));
  const outline = shape.outline ?? (box ? '000000' : undefined);
  // A shape's own outline keeps the style's width.
  if (outline !== undefined) spPr.children.push(el('a:ln', { w: box ? 9525 : undefined }, [outline === 'none' ? el('a:noFill') : solid(outline)]));
  const children = [line ? el('wps:cNvCnPr') : el('wps:cNvSpPr', { txBox: box ? 1 : undefined }), spPr];
  if (!box) {
    const scheme = (...transforms: XmlElement[]) => el('a:schemeClr', { val: 'accent1' }, transforms);
    children.push(
      el('wps:style', {}, [
        el('a:lnRef', { idx: 2 }, [scheme(el('a:shade', { val: 50000 }))]),
        el('a:fillRef', { idx: 1 }, [scheme()]),
        el('a:effectRef', { idx: 0 }, [scheme()]),
        el('a:fontRef', { idx: 'minor' }, [el('a:schemeClr', { val: 'lt1' })]),
      ]),
    );
  }
  if (box || shape.text) {
    // An empty text box still has a paragraph to type in.
    const lines = (shape.text ?? '').split(/\r?\n/);
    const paragraphs = lines.map((t) =>
      box ? el('w:p', {}, t ? [newRun(t)] : []) : el('w:p', {}, [el('w:pPr', {}, [el('w:jc', { 'w:val': 'center' })]), ...(t ? [newRun(t, { color: 'FFFFFF' })] : [])]),
    );
    children.push(el('wps:txbx', {}, [el('w:txbxContent', {}, paragraphs)]));
  }
  children.push(
    el('wps:bodyPr', { rot: 0, vert: 'horz', wrap: 'square', lIns: 91440, tIns: 45720, rIns: 91440, bIns: 45720, anchor: box ? 't' : 'ctr', anchorCtr: 0 }, [el('a:noAutofit')]),
  );
  return el('a:graphicData', { uri: 'http://schemas.microsoft.com/office/word/2010/wordprocessingShape' }, [el('wps:wsp', {}, children)]);
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
