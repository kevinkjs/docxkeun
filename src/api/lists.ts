/**
 * List editing: the labels list paragraphs show ("1.", "a)", "•"), and changing a list's format,
 * level or start number. A changed list gets its own definition in numbering.xml, copied from the
 * list it was in, so its indentation, tabs and label fonts stay as they were.
 */
import type { DocParagraph } from '../model/content.js';
import type { DocxDocument } from '../model/document.js';
import { normaliseBullet, NumberingCounter } from '../model/numbering.js';
import { attr, cloneElement, el, elements, ensureChild, firstChild, insertOrdered, removeChildren, type XmlDocument, type XmlElement } from '../xml/index.js';
import { ensurePart, randomHexId } from './parts.js';

/** List formats docxkeun can apply: bullets, or numbers shown as 1. 1) a. A. i. I. */
export type ListFormat = 'bullet' | 'decimal' | 'decimalParen' | 'lowerLetter' | 'upperLetter' | 'lowerRoman' | 'upperRoman';

export const LIST_FORMATS: readonly ListFormat[] = ['bullet', 'decimal', 'decimalParen', 'lowerLetter', 'upperLetter', 'lowerRoman', 'upperRoman'];

export interface ListItem {
  numId: string;
  /** List level, 0-based (Word's w:ilvl). */
  level: number;
  /** Word's number format for that level: 'bullet', 'decimal', 'lowerLetter', ... */
  format: string;
  /** The label as Word shows it, e.g. "1.", "b)", "iv.", "•". */
  label: string;
}

/** Child order of w:lvl. */
const LVL_ORDER = ['w:start', 'w:numFmt', 'w:lvlRestart', 'w:pStyle', 'w:isLgl', 'w:suff', 'w:lvlText', 'w:lvlPicBulletId', 'w:legacy', 'w:lvlJc', 'w:pPr', 'w:rPr'];
const RPR_FONTS = ['w:rStyle', 'w:rFonts'];

/** Bullets per level (repeating): a round bullet, an open circle and a small square. */
const BULLETS: Array<{ text: string; font?: string }> = [{ text: '•' }, { text: 'o', font: 'Courier New' }, { text: '▪', font: 'Arial' }];

/** Word's default numbering for the levels a change does not pick a format for. */
const NUMBER_CYCLE: ListFormat[] = ['decimal', 'lowerLetter', 'lowerRoman'];

/** Labels of the list paragraphs among `paragraphs` (in document order, as layout numbers them). */
export function listItems(doc: DocxDocument, paragraphs: DocParagraph[]): Map<XmlElement, ListItem> {
  const counter = new NumberingCounter(doc.numbering, (s) => doc.styles.get(s)?.pPr.numId);
  const out = new Map<XmlElement, ListItem>();
  for (const p of paragraphs) {
    if (!p.numbering) continue;
    const next = counter.next(p.numbering.numId, p.numbering.ilvl);
    if (!next) continue;
    out.set(p.el, { numId: p.numbering.numId, level: p.numbering.ilvl, format: next.level.format, label: normaliseBullet(next.label) });
  }
  return out;
}

function val(node: XmlElement | undefined): string | undefined {
  return attr(node, 'w:val');
}

function setVal(lvl: XmlElement, name: string, value: string): void {
  removeChildren(lvl, name);
  insertOrdered(lvl, el(name, { 'w:val': value }), LVL_ORDER);
}

function defaultLevel(ilvl: number): XmlElement {
  const left = 720 * (ilvl + 1);
  return el('w:lvl', { 'w:ilvl': String(ilvl) }, [
    el('w:start', { 'w:val': '1' }),
    el('w:numFmt', { 'w:val': NUMBER_CYCLE[ilvl % 3]! }),
    el('w:lvlText', { 'w:val': `%${ilvl + 1}.` }),
    el('w:lvlJc', { 'w:val': ilvl % 3 === 2 ? 'right' : 'left' }),
    el('w:pPr', {}, [el('w:ind', { 'w:left': String(left), 'w:hanging': ilvl % 3 === 2 ? '180' : '360' })]),
  ]);
}

/** Give a level a new format. Numbers keep the level's own label font; bullets get theirs. */
function applyFormat(lvl: XmlElement, format: ListFormat, ilvl: number): void {
  const wasBullet = val(firstChild(lvl, 'w:numFmt')) === 'bullet';
  removeChildren(lvl, 'w:lvlPicBulletId');
  const rPr = firstChild(lvl, 'w:rPr');
  if (format === 'bullet') {
    const bullet = BULLETS[ilvl % BULLETS.length]!;
    setVal(lvl, 'w:numFmt', 'bullet');
    setVal(lvl, 'w:lvlText', bullet.text);
    const target = rPr ?? ensureChild(lvl, 'w:rPr', LVL_ORDER);
    removeChildren(target, 'w:rFonts');
    if (bullet.font) insertOrdered(target, el('w:rFonts', { 'w:ascii': bullet.font, 'w:hAnsi': bullet.font, 'w:hint': 'default' }), RPR_FONTS);
    if (val(firstChild(lvl, 'w:lvlJc')) === 'right') setVal(lvl, 'w:lvlJc', 'left');
    return;
  }
  setVal(lvl, 'w:numFmt', format === 'decimalParen' ? 'decimal' : format);
  setVal(lvl, 'w:lvlText', `%${ilvl + 1}${format === 'decimalParen' ? ')' : '.'}`);
  // A bullet's symbol font (Symbol, Wingdings, Courier New) would turn digits into symbols.
  if (wasBullet && rPr) removeChildren(rPr, 'w:rFonts');
}

function familyOf(format: string | undefined): 'bullet' | 'number' | 'none' {
  return format === 'bullet' ? 'bullet' : format === 'none' ? 'none' : 'number';
}

function maxId(nodes: XmlElement[], name: string): number {
  return nodes.reduce((m, n) => Math.max(m, Number.parseInt(attr(n, name) ?? '', 10) || 0), 0);
}

/** The w:lvl elements a num uses: its abstract definition's, with lvlOverride levels on top. */
function levelsOf(numbering: XmlElement, numId: string, abstractId: string | undefined): Map<number, XmlElement> {
  const out = new Map<number, XmlElement>();
  const abs = elements(numbering, 'w:abstractNum').find((a) => attr(a, 'w:abstractNumId') === abstractId);
  for (const l of abs ? elements(abs, 'w:lvl') : []) out.set(Number(attr(l, 'w:ilvl') ?? 0), l);
  const num = elements(numbering, 'w:num').find((n) => attr(n, 'w:numId') === numId);
  for (const o of num ? elements(num, 'w:lvlOverride') : []) {
    const lvl = firstChild(o, 'w:lvl');
    if (lvl) out.set(Number(attr(o, 'w:ilvl') ?? 0), lvl);
  }
  return out;
}

export interface NewListOptions {
  /** A format for the levels in `levels`, or 'keep' to keep the base list's formats (to restart numbering). */
  format: ListFormat | 'keep';
  /** Levels the list paragraphs use. */
  levels: Set<number>;
  /** First number for those levels. */
  start?: number;
  /** The list to copy indentation and label formatting from. */
  base?: { numId: string; abstractId: string | undefined };
}

/** Add a list definition (abstractNum + num) to numbering.xml and return its numId. */
export async function createList(doc: DocxDocument, options: NewListOptions): Promise<string> {
  const part = await ensurePart(doc, 'numbering');
  const root = (part.xml as XmlDocument).root;
  const base = options.base ? levelsOf(root, options.base.numId, options.base.abstractId) : new Map<number, XmlElement>();
  const wanted = options.format === 'keep' ? undefined : familyOf(options.format);

  const levels: XmlElement[] = [];
  for (let i = 0; i < 9; i++) {
    const from = base.get(i);
    const lvl = from ? cloneElement(from) : defaultLevel(i);
    lvl.attrs['w:ilvl'] = String(i);
    if (options.format !== 'keep' && options.levels.has(i)) applyFormat(lvl, options.format, i);
    else if (wanted && familyOf(val(firstChild(lvl, 'w:numFmt'))) !== wanted) applyFormat(lvl, wanted === 'bullet' ? 'bullet' : NUMBER_CYCLE[i % 3]!, i);
    if (options.start !== undefined && options.levels.has(i)) setVal(lvl, 'w:start', String(options.start));
    levels.push(lvl);
  }

  const abstracts = elements(root, 'w:abstractNum');
  const nums = elements(root, 'w:num');
  const abstractId = String(maxId(abstracts, 'w:abstractNumId') + 1);
  const numId = String(maxId(nums, 'w:numId') + 1);
  const abstractNum = el('w:abstractNum', { 'w:abstractNumId': abstractId }, [
    el('w:nsid', { 'w:val': randomHexId() }),
    el('w:multiLevelType', { 'w:val': 'hybridMultilevel' }),
    ...levels,
  ]);
  const num = el('w:num', { 'w:numId': numId }, [el('w:abstractNumId', { 'w:val': abstractId })]);

  // numbering.xml order: numPicBullet*, abstractNum*, num*, numIdMacAtCleanup?
  const children = root.children;
  const lastOf = (name: string) => {
    for (let i = children.length - 1; i >= 0; i--) {
      const c = children[i]!;
      if (c.type === 'element' && c.name === name) return i;
    }
    return -1;
  };
  const afterAbstracts = Math.max(lastOf('w:abstractNum'), lastOf('w:numPicBullet')) + 1;
  children.splice(afterAbstracts, 0, abstractNum);
  const lastNum = lastOf('w:num');
  const cleanup = lastOf('w:numIdMacAtCleanup');
  children.splice(lastNum >= 0 ? lastNum + 1 : cleanup >= 0 ? cleanup : children.length, 0, num);
  doc.touch(part.name);
  return numId;
}
