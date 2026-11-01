/**
 * Positions in the XML tree and the run-splitting primitive every edit builds on.
 * A position is a child of a run (usually a w:t) plus an offset inside it.
 */
import { cloneElement, el, firstChild, textContent, type XmlElement, type XmlNode } from '../xml/index.js';

export interface DocPosition {
  /** A run child: w:t (offset counts UTF-16 units) or another run item (offset 0 = before, >0 = after). */
  node: XmlElement;
  offset: number;
}

export interface DocRange {
  start: DocPosition;
  end: DocPosition;
}

export class EditError extends Error {
  override name = 'EditError';
}

/** Parent pointers and document order for a tree; rebuild after structural edits. */
export class TreeIndex {
  readonly parent = new Map<XmlElement, XmlElement>();
  readonly order = new Map<XmlElement, number>();

  constructor(readonly root: XmlElement) {
    let n = 0;
    const visit = (node: XmlElement) => {
      this.order.set(node, n++);
      for (const c of node.children) {
        if (c.type !== 'element') continue;
        this.parent.set(c, node);
        visit(c);
      }
    };
    visit(root);
  }

  parentOf(node: XmlElement): XmlElement {
    const p = this.parent.get(node);
    if (!p) throw new EditError(`<${node.name}> is not part of this document (stale position?)`);
    return p;
  }

  /** Closest ancestor with the given name. */
  ancestor(node: XmlElement, name: string): XmlElement | undefined {
    let cur = this.parent.get(node);
    while (cur && cur.name !== name) cur = this.parent.get(cur);
    return cur;
  }

  compare(a: DocPosition, b: DocPosition): number {
    const oa = this.order.get(a.node);
    const ob = this.order.get(b.node);
    if (oa === undefined || ob === undefined) throw new EditError('Position is not part of this document');
    return oa === ob ? a.offset - b.offset : oa - ob;
  }

  contains(node: XmlElement): boolean {
    return this.order.has(node);
  }
}

export function indexOfChild(parent: XmlElement, child: XmlNode): number {
  const i = parent.children.indexOf(child);
  if (i < 0) throw new EditError('Node is not a child of its recorded parent');
  return i;
}

/** w:t with xml:space="preserve" when needed. */
export function textNode(text: string, name = 'w:t'): XmlElement {
  const t = el(name, /^\s|\s$/.test(text) ? { 'xml:space': 'preserve' } : {}, text ? [{ type: 'text', value: text }] : []);
  return t;
}

function setText(t: XmlElement, value: string): void {
  t.children = value ? [{ type: 'text', value }] : [];
  if (/^\s|\s$/.test(value)) t.attrs['xml:space'] = 'preserve';
}

function runSkeleton(run: XmlElement): XmlElement {
  const rPr = firstChild(run, 'w:rPr');
  const copy: XmlElement = { type: 'element', name: run.name, attrs: {}, children: rPr ? [cloneElement(rPr)] : [] };
  // Keep revision ids so tracked-change attribution survives.
  for (const k of ['w:rsidR', 'w:rsidRPr', 'w:rsidDel']) if (run.attrs[k]) copy.attrs[k] = run.attrs[k]!;
  return copy;
}

export interface Boundary {
  container: XmlElement;
  /** The boundary is before `container.children[at]`. */
  at: number;
}

/**
 * Split the run that contains `pos.node` so that a run boundary falls exactly at the position,
 * and return that boundary (in the run's container: a paragraph, hyperlink, sdtContent...).
 */
export function splitAt(index: TreeIndex, pos: DocPosition): Boundary {
  const run = index.parentOf(pos.node);
  if (run.name !== 'w:r') throw new EditError(`Expected a run around <${pos.node.name}>, found <${run.name}>`);
  const container = index.parentOf(run);
  const i = indexOfChild(run, pos.node);
  let splitIndex: number;
  if (pos.node.name === 'w:t' || pos.node.name === 'w:delText' || pos.node.name === 'w:instrText') {
    const text = textContent(pos.node);
    const offset = Math.max(0, Math.min(pos.offset, text.length));
    if (offset > 0 && offset < text.length) {
      setText(pos.node, text.slice(0, offset));
      const rest = textNode(text.slice(offset), pos.node.name);
      run.children.splice(i + 1, 0, rest);
      splitIndex = i + 1;
    } else splitIndex = offset === 0 ? i : i + 1;
  } else {
    splitIndex = pos.offset > 0 ? i + 1 : i;
  }
  const runIndex = indexOfChild(container, run);
  const contentStart = firstChild(run, 'w:rPr') ? 1 : 0;
  if (splitIndex <= contentStart) return { container, at: runIndex };
  if (splitIndex >= run.children.length) return { container, at: runIndex + 1 };
  const after = runSkeleton(run);
  after.children.push(...run.children.splice(splitIndex));
  container.children.splice(runIndex + 1, 0, after);
  return { container, at: runIndex + 1 };
}

const SENTINEL = 'docxkeun:mark';

/**
 * Split runs at both ends of a range and drop sentinel elements at the two boundaries, so the
 * range covers whole runs. Call `index.rebuild` afterwards and replace or remove the sentinels.
 */
export function isolate(index: TreeIndex, range: DocRange): { startMark: XmlElement; endMark: XmlElement } {
  if (index.compare(range.start, range.end) > 0) range = { start: range.end, end: range.start };
  const endMark = el(SENTINEL);
  const startMark = el(SENTINEL);
  const end = splitAt(index, range.end);
  end.container.children.splice(end.at, 0, endMark);
  const start = splitAt(index, range.start);
  start.container.children.splice(start.at, 0, startMark);
  return { startMark, endMark };
}

export function isSentinel(node: XmlNode): boolean {
  return node.type === 'element' && node.name === SENTINEL;
}

/** Elements strictly between two marks in document order (after rebuilding the index). */
export function between(index: TreeIndex, startMark: XmlElement, endMark: XmlElement, name?: string): XmlElement[] {
  const a = index.order.get(startMark);
  const b = index.order.get(endMark);
  if (a === undefined || b === undefined) throw new EditError('Range marks are not in the document');
  const out: XmlElement[] = [];
  for (const [node, o] of index.order) {
    if (o > a && o < b && (name === undefined || node.name === name)) out.push(node);
  }
  return out.sort((x, y) => index.order.get(x)! - index.order.get(y)!);
}

/** Replace a sentinel with other nodes (or remove it). */
export function replaceMark(index: TreeIndex, mark: XmlElement, nodes: XmlNode[]): void {
  const parent = index.parentOf(mark);
  parent.children.splice(indexOfChild(parent, mark), 1, ...nodes);
}

export const RPR_ORDER = [
  'w:rStyle',
  'w:rFonts',
  'w:b',
  'w:bCs',
  'w:i',
  'w:iCs',
  'w:caps',
  'w:smallCaps',
  'w:strike',
  'w:dstrike',
  'w:outline',
  'w:shadow',
  'w:emboss',
  'w:imprint',
  'w:noProof',
  'w:snapToGrid',
  'w:vanish',
  'w:webHidden',
  'w:color',
  'w:spacing',
  'w:w',
  'w:kern',
  'w:position',
  'w:sz',
  'w:szCs',
  'w:highlight',
  'w:u',
  'w:effect',
  'w:bdr',
  'w:shd',
  'w:fitText',
  'w:vertAlign',
  'w:rtl',
  'w:cs',
  'w:em',
  'w:lang',
  'w:eastAsianLayout',
  'w:specVanish',
  'w:oMath',
];

export const PPR_ORDER = [
  'w:pStyle',
  'w:keepNext',
  'w:keepLines',
  'w:pageBreakBefore',
  'w:framePr',
  'w:widowControl',
  'w:numPr',
  'w:suppressLineNumbers',
  'w:pBdr',
  'w:shd',
  'w:tabs',
  'w:suppressAutoHyphens',
  'w:kinsoku',
  'w:wordWrap',
  'w:overflowPunct',
  'w:topLinePunct',
  'w:autoSpaceDE',
  'w:autoSpaceDN',
  'w:bidi',
  'w:adjustRightInd',
  'w:snapToGrid',
  'w:spacing',
  'w:ind',
  'w:contextualSpacing',
  'w:mirrorIndents',
  'w:suppressOverlap',
  'w:jc',
  'w:textDirection',
  'w:textAlignment',
  'w:textboxTightWrap',
  'w:outlineLvl',
  'w:divId',
  'w:cnfStyle',
  'w:rPr',
  'w:sectPr',
  'w:pPrChange',
];
