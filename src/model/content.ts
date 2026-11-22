/**
 * Reads WordprocessingML stories (body, headers, footers, notes, comments) into a flat structure
 * for layout: blocks (paragraphs, tables) whose inlines are runs plus range markers. Wrappers
 * such as hyperlinks, fields, content controls and tracked insertions are flattened into markers;
 * every inline keeps a reference to the XML it came from so edits can go back to the source.
 */
import { mathToText } from '../math/linear.js';
import { attr, elements, firstChild, localName, textContent, type XmlElement } from '../xml/index.js';
import { fieldLinkTarget, setLinkSource, isDynamicField, parseFieldInstruction, type DynamicFieldType, type LinkTarget } from './fields.js';
import type { LevelDef, NumberingDefs } from './numbering.js';
import {
  mergeCellProps,
  mergeParagraphProps,
  mergeRunProps,
  mergeTableProps,
  parseCellProps,
  parseParagraphProps,
  parseRowProps,
  parseRunProps,
  parseSectionProps,
  parseTableProps,
  type CellProps,
  type ParagraphProps,
  type RowProps,
  type RunProps,
  type SectionProps,
  type TableProps,
} from './props.js';
import type { ConditionalType, ResolvedTableStyle, StyleSheet } from './styles.js';
import { emuToPt, parseIntSafe, parseOnOff, parseTwipsMeasure } from './units.js';

// ---------------------------------------------------------------------------------------------
// Types

interface RunRef {
  props: RunProps;
  /** The w:r element (or the closest run-like element) the inline came from. */
  run: XmlElement;
}

export interface TextInline extends RunRef {
  kind: 'text';
  text: string;
  /** The w:t (or w:sym, w:noBreakHyphen...) element. */
  node: XmlElement;
}
export interface TabInline extends RunRef {
  kind: 'tab';
  /** The w:tab / w:ptab element. */
  node: XmlElement;
}
export interface BreakInline extends RunRef {
  kind: 'break';
  /** The w:br / w:cr element. */
  node: XmlElement;
  type: 'line' | 'page' | 'column';
}
export interface FloatInfo {
  behindDoc: boolean;
  wrap: 'none' | 'square' | 'tight' | 'through' | 'topAndBottom';
  wrapText: 'bothSides' | 'left' | 'right' | 'largest';
  hRelative: string;
  hAlign?: string;
  hOffset: number;
  vRelative: string;
  vAlign?: string;
  vOffset: number;
  dist: { top: number; bottom: number; left: number; right: number };
}
/** A DrawingML shape: geometry, fill and outline (kept as XML: colours resolve against the theme), text. */
export interface ShapeInfo {
  /** a:prstGeom or a:custGeom. */
  geometry: XmlElement | undefined;
  /** The shape properties (fill, outline, transform). */
  spPr: XmlElement | undefined;
  /** wps:style: theme references used when spPr leaves fill or outline out. */
  style: XmlElement | undefined;
  /** Text box content (Word paragraphs), when the shape has a text box. */
  text?: { blocks: Block[]; insets: { left: number; top: number; right: number; bottom: number }; anchor: 't' | 'ctr' | 'b'; vertical: boolean; container: XmlElement; autoFit?: boolean };
  /** DrawingML text (a:txBody) of SmartArt and canvas shapes, kept as XML. */
  txBody?: XmlElement;
  /** Fill and outline colours of an old-style VML shape (RRGGBB; absent: none). */
  vml?: { fill?: string; stroke?: string };
}

/** A placed child of a group or drawing canvas, in points relative to the group's top-left. */
export interface DrawingChild {
  x: number;
  y: number;
  w: number;
  h: number;
  rotation: number;
  flipH: boolean;
  flipV: boolean;
  content: DrawingContent;
}

/** What a drawing shows when it is not a plain picture. */
export type DrawingContent =
  | { type: 'picture'; target: string | undefined; external: boolean }
  | { type: 'shape'; shape: ShapeInfo }
  | { type: 'group'; children: DrawingChild[] }
  | { type: 'chart'; part: string | undefined }
  | { type: 'diagram'; dataPart: string | undefined };

export interface ImageInline extends RunRef {
  kind: 'image';
  /** Resolved media part name (or external URL when `external`). */
  target: string | undefined;
  external: boolean;
  width: number;
  height: number;
  alt?: string;
  float?: FloatInfo;
  drawing: XmlElement;
  /** Shapes, text boxes, groups, charts and SmartArt (absent for a plain picture). */
  content?: DrawingContent;
  /** Rotation in degrees and flips of the whole drawing. */
  rotation?: number;
  flipH?: boolean;
  flipV?: boolean;
}
export interface FieldInline extends RunRef {
  kind: 'field';
  field: DynamicFieldType;
  /** Cached result text from the file, used until layout knows the real value. */
  cached: string;
  format?: string;
}
export interface NoteRefInline extends RunRef {
  kind: 'noteRef';
  noteType: 'footnote' | 'endnote';
  id: string;
  customMark: boolean;
}
export interface CommentRefInline extends RunRef {
  kind: 'commentRef';
  id: string;
}
export interface LinkStartInline {
  kind: 'linkStart';
  target: LinkTarget;
  el: XmlElement;
}
export interface LinkEndInline {
  kind: 'linkEnd';
}
export interface BookmarkStartInline {
  kind: 'bookmarkStart';
  id: string;
  name: string;
  el: XmlElement;
}
export interface BookmarkEndInline {
  kind: 'bookmarkEnd';
  id: string;
  el: XmlElement;
}
export interface CommentStartInline {
  kind: 'commentStart';
  id: string;
  el: XmlElement;
}
export interface CommentEndInline {
  kind: 'commentEnd';
  id: string;
  el: XmlElement;
}
export interface PermStartInline {
  kind: 'permStart';
  id: string;
  editor?: string;
  group?: string;
  el: XmlElement;
}
/** Start of a locked content control (w:sdt with w:lock contentLocked / sdtContentLocked). */
export interface LockStartInline {
  kind: 'lockStart';
  id: string;
  title?: string;
  el: XmlElement;
}
export interface LockEndInline {
  kind: 'lockEnd';
  id: string;
}
export interface PermEndInline {
  kind: 'permEnd';
  id: string;
  el: XmlElement;
}

/** An equation (m:oMath, or m:oMathPara for display equations on their own line). */
export interface MathInline {
  kind: 'math';
  node: XmlElement;
  /** Display style: an m:oMathPara (big operators, limits above and below). */
  display: boolean;
  /** Linear format ("x=(-b±√(b^2-4ac))/(2a)"), for text extraction and search. */
  text: string;
  props: RunProps;
}

export type Inline =
  | MathInline
  | TextInline
  | TabInline
  | BreakInline
  | ImageInline
  | FieldInline
  | NoteRefInline
  | CommentRefInline
  | LinkStartInline
  | LinkEndInline
  | BookmarkStartInline
  | BookmarkEndInline
  | CommentStartInline
  | CommentEndInline
  | PermStartInline
  | PermEndInline
  | LockStartInline
  | LockEndInline;

export interface DocParagraph {
  kind: 'paragraph';
  el: XmlElement;
  styleId: string | undefined;
  props: ParagraphProps;
  /** Resolved properties of the paragraph mark (used for empty paragraphs and list labels). */
  markProps: RunProps;
  inlines: Inline[];
  numbering?: { numId: string; ilvl: number; level: LevelDef };
  /** Set when this paragraph's w:pPr/w:sectPr ends a section. */
  sectionEnd?: SectionProps;
  /** The paragraph mark is a tracked insertion or deletion. */
  markRevision?: 'ins' | 'del';
}

/** Word-like colours for revision authors (one per author, in order of first appearance). */
export const REVISION_COLORS = ['B5082E', '2E74B5', '548235', '7030A0', 'C55A11', '00838F', '7F6000', '595959'];

export interface DocCell {
  el: XmlElement;
  props: CellProps;
  blocks: Block[];
  gridSpan: number;
}

export interface DocRow {
  el: XmlElement;
  props: RowProps;
  cells: DocCell[];
}

export interface DocTable {
  kind: 'table';
  el: XmlElement;
  styleId: string | undefined;
  props: TableProps;
  /** Column widths from w:tblGrid, in points. */
  grid: number[];
  rows: DocRow[];
}

export type Block = DocParagraph | DocTable;

export interface Story {
  blocks: Block[];
  /** Section of the last blocks (the body's final w:sectPr); undefined for non-body stories. */
  finalSection?: SectionProps;
}

export interface ReadContext {
  styles: StyleSheet;
  numbering: NumberingDefs;
  /** Relationship id → { target, external } for the part being read. */
  rels: Map<string, { target: string; external: boolean }>;
  /** Display tracked deletions (struck through) instead of hiding them. */
  showDeletions?: boolean;
}

// ---------------------------------------------------------------------------------------------
// Reader

interface FieldState {
  phase: 'instr' | 'result';
  instr: string;
  link?: LinkTarget;
  dynamic?: DynamicFieldType;
  format?: string;
  emitted: boolean;
  props?: RunProps;
  run?: XmlElement;
  cached: string;
}

interface CellContext {
  table: ResolvedTableStyle;
  conditions: ConditionalType[];
}

/** Lock info of a content control whose content cannot be edited. */
function contentLock(sdt: XmlElement): { id: string; title?: string } | undefined {
  const sdtPr = firstChild(sdt, 'w:sdtPr');
  const v = attr(firstChild(sdtPr, 'w:lock'), 'w:val');
  if (v !== 'contentLocked' && v !== 'sdtContentLocked') return undefined;
  return { id: attr(firstChild(sdtPr, 'w:id'), 'w:val') ?? '', title: attr(firstChild(sdtPr, 'w:alias'), 'w:val') };
}

const UNDERSTOOD_REQUIRES = new Set(['w14', 'w15', 'w16se', 'w16cid', 'w16cex', 'wp14', 'w16du', 'w16sdtdh', 'wps', 'wpg', 'wpc', 'cx', 'cx1', 'cx2', 'aink', 'am3d', 'a14']);

export class StoryReader {
  private fields: FieldState[] = [];
  /** Locked block-level content controls around the paragraphs being read. */
  private blockLocks: Array<{ id: string; title?: string; el: XmlElement }> = [];
  private paragraphStyleId: string | undefined;
  private cellRunProps: RunProps | undefined;

  constructor(private readonly ctx: ReadContext) {}

  readStory(container: XmlElement): Story {
    const blocks = this.readBlocks(container);
    const sectPr = firstChild(container, 'w:sectPr');
    return { blocks, finalSection: container.name === 'w:body' ? parseSectionProps(sectPr) : undefined };
  }

  readBlocks(container: XmlElement, cell?: CellContext): Block[] {
    const out: Block[] = [];
    this.collectBlocks(container, out, cell);
    return this.ctx.showDeletions ? out : joinDeletedMarks(out);
  }

  private collectBlocks(container: XmlElement, out: Block[], cell?: CellContext): void {
    for (const node of container.children) {
      if (node.type !== 'element') continue;
      switch (node.name) {
        case 'w:p':
          out.push(this.readParagraph(node, cell));
          break;
        case 'w:tbl':
          out.push(this.readTable(node));
          break;
        case 'w:sdt': {
          const lock = contentLock(node);
          if (lock) this.blockLocks.push({ ...lock, el: node });
          this.collectBlocks(firstChild(node, 'w:sdtContent') ?? node, out, cell);
          if (lock) this.blockLocks.pop();
          break;
        }
        case 'w:customXml':
        case 'w:ins':
        case 'w:moveTo':
          this.collectBlocks(node, out, cell);
          break;
        case 'mc:AlternateContent': {
          const chosen = this.chooseAlternate(node);
          if (chosen) this.collectBlocks(chosen, out, cell);
          break;
        }
        default:
          break; // bookmarks between blocks, w:sectPr, w:del, w:altChunk... are skipped
      }
    }
  }

  readParagraph(p: XmlElement, cell?: CellContext): DocParagraph {
    const { styles, numbering } = this.ctx;
    const pPr = firstChild(p, 'w:pPr');
    const direct = parseParagraphProps(pPr);
    const styleId = styles.paragraphStyleId(direct.styleId);
    const tablePPr = cell ? conditionalParagraphProps(cell) : undefined;
    let props = styles.resolveParagraph(styleId, direct, tablePPr);

    let numInfo: DocParagraph['numbering'];
    const numId = direct.numId ?? props.numId;
    if (numId && numId !== '0') {
      const ilvl = direct.ilvl ?? props.ilvl ?? 0;
      const level = numbering.level(numId, ilvl, (s) => styles.get(s)?.pPr.numId);
      if (level) {
        numInfo = { numId, ilvl, level };
        // Numbering indents sit between the style and direct formatting.
        props = mergeParagraphProps(styles.defaultParagraphProps, tablePPr, styles.paragraphStyle(styleId).pPr, level.pPr, direct);
      }
    }

    this.paragraphStyleId = styleId;
    this.cellRunProps = cell ? conditionalRunProps(cell) : undefined;
    const markDirect = parseRunProps(firstChild(pPr, 'w:rPr'));
    const markProps = styles.resolveRun(styleId, markDirect, this.cellRunProps);

    const inlines: Inline[] = [];
    for (const lock of this.blockLocks) inlines.push({ kind: 'lockStart', id: lock.id, title: lock.title, el: lock.el });
    // A link field that started in an earlier paragraph continues here.
    for (const f of this.fields) if (f.phase === 'result' && f.link) inlines.push({ kind: 'linkStart', target: f.link, el: p });
    this.readInlines(p, inlines);
    for (const f of this.fields) if (f.phase === 'result' && f.link) inlines.push({ kind: 'linkEnd' });
    for (const lock of this.blockLocks) inlines.push({ kind: 'lockEnd', id: lock.id });

    const sectPr = firstChild(pPr, 'w:sectPr');
    const markRPr = firstChild(pPr, 'w:rPr');
    const markRevision = firstChild(markRPr, 'w:del') || firstChild(markRPr, 'w:moveFrom') ? 'del' : firstChild(markRPr, 'w:ins') || firstChild(markRPr, 'w:moveTo') ? 'ins' : undefined;
    return {
      kind: 'paragraph',
      el: p,
      styleId,
      props,
      markProps,
      inlines,
      numbering: numInfo,
      sectionEnd: sectPr ? parseSectionProps(sectPr) : undefined,
      ...(markRevision ? { markRevision } : {}),
    };
  }

  /** Display formatting for text inside a tracked change (only when deletions are shown: "all markup"). */
  private revisionProps(props: RunProps): RunProps {
    const rev = this.revisions[this.revisions.length - 1];
    if (!rev || !this.ctx.showDeletions) return props;
    let index = this.authors.indexOf(rev.author);
    if (index < 0) index = this.authors.push(rev.author) - 1;
    const color = REVISION_COLORS[index % REVISION_COLORS.length]!;
    return rev.kind === 'ins' ? { ...props, underline: 'single', underlineColor: color, color } : { ...props, strike: true, color };
  }

  private readonly revisions: Array<{ kind: 'ins' | 'del'; author: string }> = [];
  private readonly authors: string[] = [];

  private suppressed(): boolean {
    for (const f of this.fields) if (f.phase === 'instr' || f.dynamic) return true;
    return false;
  }

  private readInlines(parent: XmlElement, out: Inline[]): void {
    for (const node of parent.children) {
      if (node.type !== 'element') continue;
      switch (node.name) {
        case 'w:r':
          this.readRun(node, out);
          break;
        case 'w:hyperlink': {
          const target = this.hyperlinkTarget(node);
          out.push({ kind: 'linkStart', target, el: node });
          this.readInlines(node, out);
          out.push({ kind: 'linkEnd' });
          break;
        }
        case 'w:fldSimple': {
          const f = parseFieldInstruction(attr(node, 'w:instr') ?? '');
          const link = fieldLinkTarget(f);
          if (isDynamicField(f.type)) {
            const firstRun = firstChild(node, 'w:r');
            const props = this.runProps(firstRun);
            out.push({ kind: 'field', field: f.type, cached: textContent(node), props, run: firstRun ?? node, format: formatSwitch(f) });
          } else if (link) {
            out.push({ kind: 'linkStart', target: link, el: node });
            this.readInlines(node, out);
            out.push({ kind: 'linkEnd' });
          } else {
            this.readInlines(node, out);
          }
          break;
        }
        case 'w:ins':
        case 'w:moveTo':
          this.revisions.push({ kind: 'ins', author: attr(node, 'w:author') ?? '' });
          this.readInlines(node, out);
          this.revisions.pop();
          break;
        case 'w:smartTag':
        case 'w:customXml':
        case 'w:dir':
        case 'w:bdo':
          this.readInlines(node, out);
          break;
        case 'w:del':
        case 'w:moveFrom':
          if (this.ctx.showDeletions) {
            this.revisions.push({ kind: 'del', author: attr(node, 'w:author') ?? '' });
            this.readInlines(node, out);
            this.revisions.pop();
          }
          break;
        case 'w:sdt': {
          const lock = contentLock(node);
          if (lock) out.push({ kind: 'lockStart', ...lock, el: node });
          this.readInlines(firstChild(node, 'w:sdtContent') ?? node, out);
          if (lock) out.push({ kind: 'lockEnd', id: lock.id });
          break;
        }
        case 'mc:AlternateContent': {
          const chosen = this.chooseAlternate(node);
          if (chosen) this.readInlines(chosen, out);
          break;
        }
        case 'w:bookmarkStart': {
          const id = attr(node, 'w:id');
          if (id !== undefined) out.push({ kind: 'bookmarkStart', id, name: attr(node, 'w:name') ?? '', el: node });
          break;
        }
        case 'w:bookmarkEnd': {
          const id = attr(node, 'w:id');
          if (id !== undefined) out.push({ kind: 'bookmarkEnd', id, el: node });
          break;
        }
        case 'w:commentRangeStart': {
          const id = attr(node, 'w:id');
          if (id !== undefined) out.push({ kind: 'commentStart', id, el: node });
          break;
        }
        case 'w:commentRangeEnd': {
          const id = attr(node, 'w:id');
          if (id !== undefined) out.push({ kind: 'commentEnd', id, el: node });
          break;
        }
        case 'w:permStart': {
          const id = attr(node, 'w:id');
          if (id !== undefined) out.push({ kind: 'permStart', id, editor: attr(node, 'w:ed'), group: attr(node, 'w:edGrp'), el: node });
          break;
        }
        case 'w:permEnd': {
          const id = attr(node, 'w:id');
          if (id !== undefined) out.push({ kind: 'permEnd', id, el: node });
          break;
        }
        case 'm:oMath':
        case 'm:oMathPara':
          this.readMath(node, out);
          break;
        default:
          break;
      }
    }
  }

  private runProps(run: XmlElement | undefined): RunProps {
    const direct = parseRunProps(firstChild(run, 'w:rPr'));
    return this.ctx.styles.resolveRun(this.paragraphStyleId, direct, this.cellRunProps);
  }

  private readRun(run: XmlElement, out: Inline[]): void {
    let props: RunProps | undefined;
    const getProps = () => (props ??= this.revisionProps(this.runProps(run)));

    for (const node of run.children) {
      if (node.type !== 'element') continue;
      switch (node.name) {
        case 'w:fldChar':
          this.fieldChar(node, run, getProps, out);
          continue;
        case 'w:instrText': {
          const current = this.fields[this.fields.length - 1];
          if (current && current.phase === 'instr') current.instr += textContent(node);
          continue;
        }
        default:
          break;
      }
      if (this.suppressed()) {
        const current = this.fields[this.fields.length - 1];
        if (current?.dynamic && (node.name === 'w:t' || node.name === 'w:delText')) current.cached += textContent(node);
        continue;
      }
      switch (node.name) {
        case 'w:t':
          out.push({ kind: 'text', text: textContent(node), node, run, props: getProps() });
          break;
        case 'w:delText':
          if (this.ctx.showDeletions) out.push({ kind: 'text', text: textContent(node), node, run, props: getProps() });
          break;
        case 'w:tab':
        case 'w:ptab':
          out.push({ kind: 'tab', run, node, props: getProps() });
          break;
        case 'w:br': {
          const type = attr(node, 'w:type');
          out.push({ kind: 'break', type: type === 'page' ? 'page' : type === 'column' ? 'column' : 'line', run, node, props: getProps() });
          break;
        }
        case 'w:cr':
          out.push({ kind: 'break', type: 'line', run, node, props: getProps() });
          break;
        case 'w:noBreakHyphen':
          out.push({ kind: 'text', text: '‑', node, run, props: getProps() });
          break;
        case 'w:softHyphen':
          out.push({ kind: 'text', text: '­', node, run, props: getProps() });
          break;
        case 'w:sym': {
          const font = attr(node, 'w:font');
          let code = Number.parseInt(attr(node, 'w:char') ?? '20', 16);
          if (code >= 0xf000 && code <= 0xf0ff && font && /symbol|wingdings|webdings/i.test(font)) code -= 0xf000;
          const base = getProps();
          out.push({
            kind: 'text',
            text: String.fromCodePoint(code),
            node,
            run,
            props: font ? { ...base, fonts: { ascii: font, hAnsi: font, eastAsia: font, cs: font } } : base,
          });
          break;
        }
        case 'w:footnoteReference':
        case 'w:endnoteReference':
          out.push({
            kind: 'noteRef',
            noteType: node.name === 'w:footnoteReference' ? 'footnote' : 'endnote',
            id: attr(node, 'w:id') ?? '',
            customMark: attr(node, 'w:customMarkFollows') !== undefined && parseOnOff(attr(node, 'w:customMarkFollows')),
            run,
            props: getProps(),
          });
          break;
        case 'w:footnoteRef':
        case 'w:endnoteRef':
          // The note's own number at the start of a footnote/endnote; layout fills it in.
          out.push({ kind: 'noteRef', noteType: node.name === 'w:footnoteRef' ? 'footnote' : 'endnote', id: '', customMark: false, run, props: getProps() });
          break;
        case 'w:commentReference':
          out.push({ kind: 'commentRef', id: attr(node, 'w:id') ?? '', run, props: getProps() });
          break;
        case 'w:drawing':
          for (const d of elements(node)) {
            const img = this.readDrawing(d, run, getProps());
            if (img) out.push(img);
          }
          break;
        case 'w:pict':
        case 'w:object': {
          const img = this.readVml(node, run, getProps());
          if (img) out.push(img);
          break;
        }
        case 'mc:AlternateContent': {
          const chosen = this.chooseAlternate(node);
          if (chosen) {
            const wrapper: XmlElement = { type: 'element', name: 'w:r', attrs: run.attrs, children: [...(firstChild(run, 'w:rPr') ? [firstChild(run, 'w:rPr')!] : []), ...chosen.children] };
            this.readRun(wrapper, out);
          }
          break;
        }
        default:
          break;
      }
    }
  }

  private fieldChar(node: XmlElement, run: XmlElement, getProps: () => RunProps, out: Inline[]): void {
    const type = attr(node, 'w:fldCharType');
    if (type === 'begin') {
      this.fields.push({ phase: 'instr', instr: '', emitted: false, cached: '' });
      return;
    }
    const top = this.fields[this.fields.length - 1];
    if (!top) return;
    if (type === 'separate') {
      if (top.phase !== 'instr') return;
      this.activateField(top, run, getProps);
      top.phase = 'result';
      if (!this.outerSuppressed() && top.link) out.push({ kind: 'linkStart', target: top.link, el: run });
      return;
    }
    if (type === 'end') {
      if (top.phase === 'instr') this.activateField(top, run, getProps);
      this.fields.pop();
      if (this.suppressed()) return;
      if (top.dynamic) {
        out.push({ kind: 'field', field: top.dynamic, cached: top.cached, props: top.props ?? getProps(), run: top.run ?? run, format: top.format });
      } else if (top.link && top.phase === 'result') {
        out.push({ kind: 'linkEnd' });
      }
    }
  }

  private outerSuppressed(): boolean {
    for (let i = 0; i < this.fields.length - 1; i++) {
      const f = this.fields[i]!;
      if (f.phase === 'instr' || f.dynamic) return true;
    }
    return false;
  }

  private activateField(f: FieldState, run: XmlElement, getProps: () => RunProps): void {
    const instr = parseFieldInstruction(f.instr);
    f.link = fieldLinkTarget(instr);
    if (isDynamicField(instr.type)) {
      f.dynamic = instr.type;
      f.format = formatSwitch(instr);
      f.props = getProps();
      f.run = run;
    }
  }

  private hyperlinkTarget(node: XmlElement): LinkTarget {
    const rId = attr(node, 'r:id');
    const anchor = attr(node, 'w:anchor');
    const tooltip = attr(node, 'w:tooltip');
    let url: string | undefined;
    if (rId) {
      const rel = this.ctx.rels.get(rId);
      if (rel) url = rel.target;
    }
    const location = attr(node, 'w:docLocation');
    if (url && location) url += `#${location}`;
    return setLinkSource({ url, anchor, tooltip }, node);
  }

  private chooseAlternate(node: XmlElement): XmlElement | undefined {
    for (const choice of elements(node, 'mc:Choice')) {
      const requires = (attr(choice, 'Requires') ?? '').split(/\s+/).filter(Boolean);
      if (requires.every((r) => UNDERSTOOD_REQUIRES.has(r))) return choice;
    }
    return firstChild(node, 'mc:Fallback');
  }

  private readDrawing(d: XmlElement, run: XmlElement, props: RunProps): ImageInline | undefined {
    const isAnchor = d.name === 'wp:anchor';
    if (!isAnchor && d.name !== 'wp:inline') return undefined;
    const extent = firstChild(d, 'wp:extent');
    const width = emuToPt(Number(attr(extent, 'cx') ?? 0));
    const height = emuToPt(Number(attr(extent, 'cy') ?? 0));
    const docPr = firstChild(d, 'wp:docPr');
    const graphicData = firstChild(firstChild(d, 'a:graphic'), 'a:graphicData');
    const content = this.readGraphic(graphicData, width, height);
    const picture = content?.type === 'picture' ? content : undefined;
    const image: ImageInline = {
      kind: 'image',
      target: picture?.target,
      external: picture?.external ?? false,
      width,
      height,
      alt: attr(docPr, 'descr') || attr(docPr, 'title') || undefined,
      drawing: d,
      run,
      props,
    };
    if (content && content.type !== 'picture') image.content = content;
    const xfrm = graphicData ? findDescendant(graphicData, 'a:xfrm') : undefined;
    const transform = readTransform(xfrm);
    if (transform.rotation) image.rotation = transform.rotation;
    if (transform.flipH) image.flipH = true;
    if (transform.flipV) image.flipV = true;
    if (isAnchor) image.float = readFloat(d);
    return image;
  }

  /** The content of an a:graphicData: picture, shape, group, canvas, chart or SmartArt. */
  private readGraphic(graphicData: XmlElement | undefined, width: number, height: number): DrawingContent | undefined {
    if (!graphicData) return undefined;
    for (const c of elements(graphicData)) {
      switch (c.name) {
        case 'pic:pic':
          return this.readPicture(c);
        case 'wps:wsp':
          return { type: 'shape', shape: this.readShape(c) };
        case 'wpg:wgp':
        case 'wpc:wpc':
          return { type: 'group', children: this.readGroupChildren(c, { x: 0, y: 0, w: width, h: height }) };
        case 'c:chart': {
          const rel = this.ctx.rels.get(attr(c, 'r:id') ?? '');
          return { type: 'chart', part: rel?.target };
        }
        case 'dgm:relIds': {
          const rel = this.ctx.rels.get(attr(c, 'r:dm') ?? '');
          return { type: 'diagram', dataPart: rel?.target };
        }
        default:
          break;
      }
    }
    // Anything else with a picture in it (e.g. an unknown wrapper) still shows the picture.
    const blip = findDescendant(graphicData, 'a:blip');
    return blip ? this.readPicture(graphicData) : undefined;
  }

  private readPicture(node: XmlElement): DrawingContent {
    const blip = findDescendant(node, 'a:blip');
    const rId = attr(blip, 'r:embed') ?? attr(blip, 'r:link');
    const rel = rId ? this.ctx.rels.get(rId) : undefined;
    return { type: 'picture', target: rel?.target, external: rel?.external ?? false };
  }

  private readShape(wsp: XmlElement): ShapeInfo {
    const spPr = firstChild(wsp, 'wps:spPr') ?? firstChild(wsp, 'dsp:spPr') ?? firstChild(wsp, 'a:spPr') ?? firstChild(wsp, 'p:spPr');
    const geometry = firstChild(spPr, 'a:prstGeom') ?? firstChild(spPr, 'a:custGeom');
    const shape: ShapeInfo = { geometry, spPr, style: firstChild(wsp, 'wps:style') ?? firstChild(wsp, 'dsp:style') };
    const content = firstChild(firstChild(wsp, 'wps:txbx'), 'w:txbxContent');
    const bodyPr = firstChild(wsp, 'wps:bodyPr');
    if (content) {
      const inset = (name: string, fallback: number) => {
        const v = parseIntSafe(attr(bodyPr, name));
        return v === undefined ? fallback : emuToPt(v);
      };
      const anchor = attr(bodyPr, 'anchor');
      const vert = attr(bodyPr, 'vert');
      shape.text = {
        container: content,
        autoFit: !!firstChild(bodyPr, 'a:spAutoFit'),
        blocks: this.readBlocks(content),
        insets: { left: inset('lIns', 7.2), top: inset('tIns', 3.6), right: inset('rIns', 7.2), bottom: inset('bIns', 3.6) },
        anchor: anchor === 'ctr' ? 'ctr' : anchor === 'b' ? 'b' : 't',
        vertical: vert !== undefined && vert !== 'horz',
      };
    }
    const txBody = firstChild(wsp, 'dsp:txBody') ?? firstChild(wsp, 'wps:txBody');
    if (txBody) shape.txBody = txBody;
    return shape;
  }

  /** Children of a group or canvas, mapped from the group's child coordinates into `frame`. */
  private readGroupChildren(group: XmlElement, frame: { x: number; y: number; w: number; h: number }): DrawingChild[] {
    const grpSpPr = firstChild(group, 'wpg:grpSpPr') ?? firstChild(group, 'wpc:bg') ?? firstChild(group, 'a:grpSpPr');
    const xfrm = firstChild(grpSpPr, 'a:xfrm');
    const chOff = firstChild(xfrm, 'a:chOff');
    const chExt = firstChild(xfrm, 'a:chExt');
    const off = { x: Number(attr(chOff, 'x') ?? 0), y: Number(attr(chOff, 'y') ?? 0) };
    const ext = { cx: Number(attr(chExt, 'cx') ?? 0), cy: Number(attr(chExt, 'cy') ?? 0) };
    // Without child extents (a canvas), child coordinates are EMU from the frame's corner.
    const sx = ext.cx > 0 ? frame.w / ext.cx : 1 / 12700;
    const sy = ext.cy > 0 ? frame.h / ext.cy : 1 / 12700;
    const out: DrawingChild[] = [];
    for (const c of elements(group)) {
      let content: DrawingContent | undefined;
      let spPr: XmlElement | undefined;
      switch (c.name) {
        case 'wps:wsp':
          content = { type: 'shape', shape: this.readShape(c) };
          spPr = firstChild(c, 'wps:spPr');
          break;
        case 'pic:pic':
          content = this.readPicture(c);
          spPr = firstChild(c, 'pic:spPr');
          break;
        case 'wpg:grpSp':
        case 'wpg:wgp':
          spPr = firstChild(c, 'wpg:grpSpPr');
          break;
        case 'wpc:wpc':
          break;
        default:
          continue;
      }
      const t = readTransform(firstChild(spPr, 'a:xfrm'));
      const child = {
        x: frame.x + (t.x - off.x) * sx,
        y: frame.y + (t.y - off.y) * sy,
        w: t.cx * sx,
        h: t.cy * sy,
      };
      if (!content) content = { type: 'group', children: this.readGroupChildren(c, { x: 0, y: 0, w: child.w, h: child.h }) };
      out.push({ ...child, rotation: t.rotation, flipH: t.flipH, flipV: t.flipV, content });
    }
    return out;
  }

  private readVml(node: XmlElement, run: XmlElement, props: RunProps): ImageInline | undefined {
    const imagedata = findDescendant(node, 'v:imagedata');
    const rId = attr(imagedata, 'r:id');
    if (!rId) return this.readVmlTextBox(node, run, props);
    const shape = findDescendant(node, 'v:shape');
    const style = parseCssStyle(attr(shape, 'style') ?? '');
    const rel = this.ctx.rels.get(rId);
    return {
      kind: 'image',
      target: rel?.target,
      external: rel?.external ?? false,
      width: cssLengthToPt(style.width) ?? 72,
      height: cssLengthToPt(style.height) ?? 72,
      alt: attr(shape, 'alt'),
      drawing: node,
      run,
      props,
    };
  }

  /** An old-style (VML) text box: its Word paragraphs in a plain frame, sized from the VML style. */
  private readVmlTextBox(node: XmlElement, run: XmlElement, props: RunProps): ImageInline | undefined {
    const textbox = findDescendant(node, 'v:textbox');
    const content = firstChild(textbox, 'w:txbxContent');
    if (!content) return undefined;
    const shape = findDescendant(node, 'v:shape') ?? findDescendant(node, 'v:rect') ?? findDescendant(node, 'v:roundrect');
    const style = parseCssStyle(attr(shape, 'style') ?? '');
    const fill = attr(shape, 'fillcolor');
    const stroke = attr(shape, 'strokecolor');
    const inline: ImageInline = {
      kind: 'image',
      target: undefined,
      external: false,
      width: cssLengthToPt(style.width) ?? 144,
      height: cssLengthToPt(style.height) ?? 72,
      drawing: node,
      run,
      props,
      content: {
        type: 'shape',
        shape: {
          geometry: undefined,
          spPr: undefined,
          style: undefined,
          vml: { fill: attr(shape, 'filled') === 'f' ? undefined : (fill?.replace('#', '') ?? 'FFFFFF'), stroke: attr(shape, 'stroked') === 'f' ? undefined : (stroke?.replace('#', '') ?? '000000') },
          text: { container: content, blocks: this.readBlocks(content), insets: { left: 7.2, top: 3.6, right: 7.2, bottom: 3.6 }, anchor: 't', vertical: false },
        },
      },
    };
    if (style.position === 'absolute') {
      inline.float = {
        behindDoc: (style['z-index'] ?? '').startsWith('-'),
        wrap: 'square',
        wrapText: 'bothSides',
        hRelative: style['mso-position-horizontal-relative'] ?? 'column',
        hAlign: style['mso-position-horizontal'],
        hOffset: cssLengthToPt(style['margin-left']) ?? 0,
        vRelative: style['mso-position-vertical-relative'] ?? 'paragraph',
        vAlign: style['mso-position-vertical'],
        vOffset: cssLengthToPt(style['margin-top']) ?? 0,
        dist: { top: 0, bottom: 0, left: 9, right: 9 },
      };
    }
    return inline;
  }

  private readMath(node: XmlElement, out: Inline[]): void {
    out.push({ kind: 'math', node, display: node.name === 'm:oMathPara', text: mathToText(node), props: this.runProps(undefined) });
  }

  readTable(tbl: XmlElement): DocTable {
    const { styles } = this.ctx;
    const direct = parseTableProps(firstChild(tbl, 'w:tblPr'));
    const styleId = direct.styleId ?? styles.defaults.table;
    const tableStyle = styles.tableStyle(styleId);
    const props = mergeTableProps(tableStyle.tblPr, direct);
    const grid = elements(firstChild(tbl, 'w:tblGrid') ?? tbl, 'w:gridCol').map((g) => parseTwipsMeasure(attr(g, 'w:w')) ?? 0);

    const rowEls: XmlElement[] = [];
    const collectRows = (parent: XmlElement) => {
      for (const c of elements(parent)) {
        if (c.name === 'w:tr') rowEls.push(c);
        else if (c.name === 'w:sdt') collectRows(firstChild(c, 'w:sdtContent') ?? c);
        else if (c.name === 'w:customXml' || c.name === 'w:ins') collectRows(c);
      }
    };
    collectRows(tbl);

    const look = props.look ?? { firstRow: true, lastRow: false, firstColumn: true, lastColumn: false, noHBand: false, noVBand: true };
    const rowBand = Math.max(1, props.rowBandSize ?? 1);
    const colBand = Math.max(1, props.colBandSize ?? 1);
    const rows: DocRow[] = [];
    const savedFields = this.fields;

    rowEls.forEach((tr, rowIndex) => {
      const rowProps = { ...tableStyle.trPr, ...parseRowProps(firstChild(tr, 'w:trPr')) };
      const cellEls: XmlElement[] = [];
      const collectCells = (parent: XmlElement) => {
        for (const c of elements(parent)) {
          if (c.name === 'w:tc') cellEls.push(c);
          else if (c.name === 'w:sdt') collectCells(firstChild(c, 'w:sdtContent') ?? c);
          else if (c.name === 'w:customXml') collectCells(c);
        }
      };
      collectCells(tr);
      const isFirst = look.firstRow && rowIndex === 0;
      const isLast = look.lastRow && rowIndex === rowEls.length - 1;
      const bodyRow = rowIndex - (look.firstRow ? 1 : 0);
      const cells: DocCell[] = [];
      let colIndex = rowProps.gridBefore ?? 0;
      cellEls.forEach((tc, cellIndex) => {
        const cellDirect = parseCellProps(firstChild(tc, 'w:tcPr'));
        const span = Math.max(1, cellDirect.gridSpan ?? 1);
        const conditions: ConditionalType[] = ['wholeTable'];
        const isFirstCol = look.firstColumn && cellIndex === 0;
        const isLastCol = look.lastColumn && cellIndex === cellEls.length - 1;
        if (!look.noVBand && !isFirstCol && !isLastCol) {
          const bodyCol = colIndex - (look.firstColumn ? 1 : 0);
          conditions.push(Math.floor(bodyCol / colBand) % 2 === 0 ? 'band1Vert' : 'band2Vert');
        }
        if (!look.noHBand && !isFirst && !isLast) conditions.push(Math.floor(bodyRow / rowBand) % 2 === 0 ? 'band1Horz' : 'band2Horz');
        if (isFirstCol) conditions.push('firstCol');
        if (isLastCol) conditions.push('lastCol');
        if (isFirst) conditions.push('firstRow');
        if (isLast) conditions.push('lastRow');
        if (isFirst && isFirstCol) conditions.push('nwCell');
        if (isFirst && isLastCol) conditions.push('neCell');
        if (isLast && isFirstCol) conditions.push('swCell');
        if (isLast && isLastCol) conditions.push('seCell');
        const conditionalCell = mergeCellProps(tableStyle.tcPr, ...conditions.map((c) => tableStyle.conditional.get(c)?.tcPr));
        const cellProps = mergeCellProps(conditionalCell, cellDirect);
        // Fields cannot span cells; read each cell with a fresh field stack.
        this.fields = [];
        const blocks = this.readBlocks(tc, { table: tableStyle, conditions });
        cells.push({ el: tc, props: cellProps, blocks, gridSpan: span });
        colIndex += span;
      });
      rows.push({ el: tr, props: rowProps, cells });
    });
    this.fields = savedFields;
    return { kind: 'table', el: tbl, styleId, props, grid, rows };
  }
}

function conditionalParagraphProps(cell: CellContext): ParagraphProps {
  return mergeParagraphProps(cell.table.pPr, ...cell.conditions.map((c) => cell.table.conditional.get(c)?.pPr));
}

function conditionalRunProps(cell: CellContext): RunProps {
  return mergeRunProps(cell.table.rPr, ...cell.conditions.map((c) => cell.table.conditional.get(c)?.rPr));
}

function formatSwitch(f: ReturnType<typeof parseFieldInstruction>): string | undefined {
  const fmt = f.switches.get('\\*');
  return typeof fmt === 'string' ? fmt : undefined;
}

function readFloat(anchor: XmlElement): FloatInfo {
  const posH = firstChild(anchor, 'wp:positionH');
  const posV = firstChild(anchor, 'wp:positionV');
  const emu = (n: XmlElement | undefined, name: string) => {
    const v = parseIntSafe(attr(n, name));
    return v === undefined ? 0 : emuToPt(v);
  };
  let wrap: FloatInfo['wrap'] = 'none';
  let wrapText: FloatInfo['wrapText'] = 'bothSides';
  for (const c of elements(anchor)) {
    const name = localName(c.name);
    if (name === 'wrapSquare') wrap = 'square';
    else if (name === 'wrapTight') wrap = 'tight';
    else if (name === 'wrapThrough') wrap = 'through';
    else if (name === 'wrapTopAndBottom') wrap = 'topAndBottom';
    else if (name === 'wrapNone') wrap = 'none';
    else continue;
    const wt = attr(c, 'wrapText');
    if (wt === 'left' || wt === 'right' || wt === 'largest') wrapText = wt;
  }
  return {
    behindDoc: parseOnOff(attr(anchor, 'behindDoc') ?? '0'),
    wrap,
    wrapText,
    hRelative: attr(posH, 'relativeFrom') ?? 'column',
    hAlign: textOf(firstChild(posH, 'wp:align')),
    hOffset: emuToPt(Number(textOf(firstChild(posH, 'wp:posOffset')) ?? 0)),
    vRelative: attr(posV, 'relativeFrom') ?? 'paragraph',
    vAlign: textOf(firstChild(posV, 'wp:align')),
    vOffset: emuToPt(Number(textOf(firstChild(posV, 'wp:posOffset')) ?? 0)),
    dist: { top: emu(anchor, 'distT'), bottom: emu(anchor, 'distB'), left: emu(anchor, 'distL'), right: emu(anchor, 'distR') },
  };
}

function textOf(node: XmlElement | undefined): string | undefined {
  if (!node) return undefined;
  const t = textContent(node).trim();
  return t || undefined;
}

export function findDescendant(node: XmlElement | undefined, name: string): XmlElement | undefined {
  if (!node) return undefined;
  for (const c of node.children) {
    if (c.type !== 'element') continue;
    if (c.name === name) return c;
    const found = findDescendant(c, name);
    if (found) return found;
  }
  return undefined;
}

function findAllDescendants(node: XmlElement, name: string, out: XmlElement[] = []): XmlElement[] {
  for (const c of node.children) {
    if (c.type !== 'element') continue;
    if (c.name === name) out.push(c);
    else findAllDescendants(c, name, out);
  }
  return out;
}

function parseCssStyle(style: string): Record<string, string> {
  const out: Record<string, string> = {};
  for (const decl of style.split(';')) {
    const i = decl.indexOf(':');
    if (i > 0) out[decl.slice(0, i).trim().toLowerCase()] = decl.slice(i + 1).trim();
  }
  return out;
}

function cssLengthToPt(v: string | undefined): number | undefined {
  if (!v) return undefined;
  const m = /^(-?\d+(?:\.\d+)?)(pt|in|cm|mm|px)?$/.exec(v);
  if (!m) return undefined;
  const n = Number.parseFloat(m[1]!);
  switch (m[2]) {
    case 'in':
      return n * 72;
    case 'cm':
      return (n * 72) / 2.54;
    case 'mm':
      return (n * 72) / 25.4;
    case 'px':
      return n * 0.75;
    default:
      return n;
  }
}

/** Plain text of a block list (paragraphs separated by newlines, cells by tabs). */
export function blocksToText(blocks: Block[]): string {
  const lines: string[] = [];
  for (const b of blocks) {
    if (b.kind === 'paragraph') {
      let s = '';
      for (const i of b.inlines) {
        if (i.kind === 'text') s += i.text;
        else if (i.kind === 'tab') s += '\t';
        else if (i.kind === 'break' && i.type === 'line') s += '\n';
        else if (i.kind === 'field') s += i.cached;
        else if (i.kind === 'math') s += i.text;
      }
      lines.push(s);
    } else {
      for (const row of b.rows) lines.push(row.cells.map((c) => blocksToText(c.blocks)).join('\t'));
    }
  }
  return lines.join('\n');
}

/**
 * The final view of tracked changes: a paragraph whose mark was deleted runs on into the next
 * paragraph (which keeps its own properties), as Word shows it with markup hidden.
 */
function joinDeletedMarks(blocks: Block[]): Block[] {
  const out: Block[] = [];
  let carry: Inline[] = [];
  for (const b of blocks) {
    if (b.kind === 'paragraph' && b.markRevision === 'del') {
      carry = [...carry, ...b.inlines];
      continue;
    }
    if (carry.length && b.kind === 'paragraph') {
      b.inlines = [...carry, ...b.inlines];
      carry = [];
    } else if (carry.length) {
      // A table follows: the carried text stays as its own paragraph rather than vanishing.
      const last = blocks[blocks.indexOf(b) - 1];
      if (last && last.kind === 'paragraph') out.push({ ...last, inlines: carry, markRevision: undefined });
      carry = [];
    }
    out.push(b);
  }
  if (carry.length) {
    const last = blocks[blocks.length - 1];
    if (last && last.kind === 'paragraph') out.push({ ...last, inlines: carry, markRevision: undefined });
  }
  return out;
}

/** Position, size (EMU) and rotation (degrees) of an a:xfrm. */
function readTransform(xfrm: XmlElement | undefined): { x: number; y: number; cx: number; cy: number; rotation: number; flipH: boolean; flipV: boolean } {
  const off = firstChild(xfrm, 'a:off');
  const ext = firstChild(xfrm, 'a:ext');
  return {
    x: Number(attr(off, 'x') ?? 0),
    y: Number(attr(off, 'y') ?? 0),
    cx: Number(attr(ext, 'cx') ?? 0),
    cy: Number(attr(ext, 'cy') ?? 0),
    rotation: Number(attr(xfrm, 'rot') ?? 0) / 60000,
    flipH: parseOnOff(attr(xfrm, 'flipH') ?? '0'),
    flipV: parseOnOff(attr(xfrm, 'flipV') ?? '0'),
  };
}
