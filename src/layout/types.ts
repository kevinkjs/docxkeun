/**
 * The display list: what layout produces and what renderers (DOM/SVG, PDF) draw.
 * Coordinates are points with the origin at the page's top-left corner, y pointing down.
 */
import type { FontFace } from '../fonts/registry.js';
import type { ImageInline, TextInline } from '../model/content.js';
import type { LinkTarget } from '../model/fields.js';
import type { SectionProps } from '../model/props.js';

/** A font as layout sees it: metrics per 1pt of size, plus glyph lookup. */
export interface LayoutFont {
  /** Unique key; also the CSS family name renderers register. */
  readonly key: string;
  readonly family: string;
  /** Ascent and descent (positive) as fractions of the em. */
  readonly ascent: number;
  readonly descent: number;
  readonly underlinePosition: number;
  readonly underlineThickness: number;
  /** Advance width of a code point as a fraction of the em. */
  advance(codePoint: number): number;
  glyphId(codePoint: number): number;
  has(codePoint: number): boolean;
  /** Set when the font is a registered font file (needed for PDF embedding). */
  readonly face?: FontFace;
  readonly syntheticBold: boolean;
  readonly syntheticItalic: boolean;
  readonly bold: boolean;
  readonly italic: boolean;
}

/** Supplies fonts to layout. */
export interface FontProvider {
  /** Font for a family and style; never fails (falls back to a default). */
  get(family: string | undefined, bold: boolean, italic: boolean): LayoutFont;
  /** A font that can render the code point, preferring `font`. */
  forChar(font: LayoutFont, codePoint: number): LayoutFont;
}

export interface TextSource {
  inline: TextInline;
  /** Offset in `inline.text` of each character cluster in the item. */
  offsets: number[];
}

export interface TextItem {
  type: 'text';
  /** Baseline origin of the first character. */
  x: number;
  y: number;
  text: string;
  /** Character clusters of `text` (usually one code point each). */
  chars: string[];
  /** Absolute x of each cluster. */
  xs: number[];
  /** Advance of each cluster (pt). */
  widths: number[];
  glyphs: number[];
  font: LayoutFont;
  size: number;
  color: string;
  /** Horizontal scale (w:w), 1 = normal. */
  scale: number;
  source?: TextSource;
  /** Present when the text is a resolved field (page numbers). */
  field?: 'PAGE' | 'NUMPAGES' | 'SECTIONPAGES';
  /**
   * Right-to-left or complex-script text: the characters in logical order. `chars`, `xs` and
   * `glyphs` are in visual order; renderers that shape text (SVG) draw `text` as one run.
   */
  complex?: { text: string; rtl: boolean };
  /** Glyphs of each cluster from the shaping engine (ligatures, contextual forms), when shaped. */
  clusterGlyphs?: Array<Array<{ id: number; dx: number; dy: number; advance: number; text: string }>>;
}

export interface RectItem {
  type: 'rect';
  x: number;
  y: number;
  w: number;
  h: number;
  fill: string;
}

export interface LineItem {
  type: 'line';
  x1: number;
  y1: number;
  x2: number;
  y2: number;
  width: number;
  color: string;
  style: 'solid' | 'dashed' | 'dotted' | 'double';
}

/** Decoded pixels: for images a renderer cannot take as they are (GIF and BMP in PDF, bitmaps inside metafiles). */
export interface RasterImage {
  width: number;
  height: number;
  /** 8-bit RGBA, row by row from the top, 4 bytes per pixel (not premultiplied). */
  rgba: Uint8Array;
}

export interface ImageItem {
  type: 'image';
  x: number;
  y: number;
  w: number;
  h: number;
  /** Media part name (or external URL). */
  target: string | undefined;
  external: boolean;
  alt?: string;
  /** The image as drawn, when it does not come from a media part (or must not be read from one). */
  raster?: RasterImage;
  inline?: ImageInline;
  /** Clockwise rotation in degrees about the centre, and flips. */
  rotation?: number;
  flipH?: boolean;
  flipV?: boolean;
  /** Only marks the object's box (for hit-testing): what it shows is drawn by the items after it. */
  frameOnly?: boolean;
}

/** Path segments in points, page coordinates (y down). Arcs are converted to cubic Béziers. */
export type PathCommand =
  | { op: 'M'; x: number; y: number }
  | { op: 'L'; x: number; y: number }
  | { op: 'C'; x1: number; y1: number; x2: number; y2: number; x: number; y: number }
  | { op: 'Z' };

/** A filled and/or stroked vector path: shapes, charts, equations, metafiles. */
export interface PathItem {
  type: 'path';
  commands: PathCommand[];
  /** Fill colour as RRGGBB; no fill when absent. */
  fill?: string;
  /** 0..1, default 1. */
  fillOpacity?: number;
  /** Even-odd fill rule instead of non-zero. */
  evenOdd?: boolean;
  /** Stroke colour as RRGGBB; no stroke when absent. */
  stroke?: string;
  /** Stroke width in points (default 0.75). */
  strokeWidth?: number;
  strokeOpacity?: number;
  /** Dash pattern in points (on, off, …). */
  dash?: number[];
  lineCap?: 'butt' | 'round' | 'square';
  lineJoin?: 'miter' | 'round' | 'bevel';
}

export type DisplayItem = TextItem | RectItem | LineItem | ImageItem | PathItem;

export interface Rect {
  x: number;
  y: number;
  w: number;
  h: number;
}

export interface LinkArea extends Rect {
  target: LinkTarget;
}

export interface CommentArea extends Rect {
  id: string;
}

export interface LockArea extends Rect {
  /** Content control id of the locked region. */
  id: string;
  title?: string;
}

export interface PermArea extends Rect {
  id: string;
  editor?: string;
  group?: string;
}

export interface Page {
  index: number;
  /** Displayed page number (honours section restarts), formatted. */
  label: string;
  number: number;
  width: number;
  height: number;
  section: SectionProps;
  sectionIndex: number;
  items: DisplayItem[];
  links: LinkArea[];
  comments: CommentArea[];
  perms: PermArea[];
  /** Text in locked regions (content controls that cannot be edited). */
  locks: LockArea[];
  /** Bookmark name → position on this page. */
  anchors: Map<string, { x: number; y: number }>;
  /** Every body line (not headers, footers or notes), for caret placement and hit-testing. */
  lines: BodyLine[];
  /** Lines of headers, footers, footnotes and text boxes, for editing them in place. */
  storyLines: StoryLine[];
  /** Paragraphs with an outline level (headings), for PDF outlines and navigation panes. */
  headings: Array<{ level: number; title: string; x: number; y: number }>;
  /** Body content area (for viewers that draw margins). */
  content: Rect;
}

/** Stories other than the main body that can be edited in place. */
export type StoryKind = 'header' | 'footer' | 'footnote' | 'endnote' | 'textbox';

export interface StoryRef {
  /** Part holding the story's XML (the main document for text boxes). */
  part: string;
  /** Footnote or endnote id. */
  noteId?: string;
  /** The w:txbxContent of a text box. */
  container?: import('../xml/dom.js').XmlElement;
}

/** A line of a header, footer, note or text box. */
export interface StoryLine extends BodyLine, StoryRef {
  story: StoryKind;
}

export interface BodyLine {
  /** Left edge of the line's content box and its width (between the paragraph's indents). */
  x: number;
  w: number;
  y: number;
  h: number;
  baseline: number;
  /** The w:p this line belongs to. */
  paragraph: import('../xml/dom.js').XmlElement;
}

export interface LayoutResult {
  pages: Page[];
  /** Bookmark name → page index and position. */
  anchors: Map<string, { page: number; x: number; y: number }>;
  /** Comment id → first page it appears on and its first rect. */
  commentAnchors: Map<string, { page: number; rect: Rect }>;
}
