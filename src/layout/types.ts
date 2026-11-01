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
  inline: ImageInline;
}

export type DisplayItem = TextItem | RectItem | LineItem | ImageItem;

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
  /** Bookmark name → position on this page. */
  anchors: Map<string, { x: number; y: number }>;
  /** Every body line (not headers, footers or notes), for caret placement and hit-testing. */
  lines: BodyLine[];
  /** Paragraphs with an outline level (headings), for PDF outlines and navigation panes. */
  headings: Array<{ level: number; title: string; x: number; y: number }>;
  /** Body content area (for viewers that draw margins). */
  content: Rect;
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
