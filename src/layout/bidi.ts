/**
 * Right-to-left text in line layout. Paragraphs are laid out in logical order, as if left to
 * right from the paragraph's start edge (Word's w:ind and w:jc are start/end values in a
 * right-to-left paragraph); each finished line is then put in visual order with the Unicode
 * bidi algorithm, and a right-to-left paragraph's lines are mirrored onto the page.
 */
import { lineLevels, resolveLevels } from '../text/bidi.js';
import { visualOrder } from '../text/bidi.js';
import type { Atom } from './lines.js';

export interface BidiInfo {
  /** One stand-in string for the paragraph's atoms. */
  text: string;
  /** Resolved embedding levels, per UTF-16 code unit of `text`. */
  levels: Uint8Array;
  /** Where each atom starts in `text` (and the end, at index atoms.length). */
  offsets: number[];
  paragraphLevel: 0 | 1;
}

const infos = new WeakMap<Atom[], BidiInfo>();

const RTL = /[֐-ࣿיִ-﷿ﹰ-ﻼ\u{10800}-\u{10fff}\u{1e800}-\u{1efff}‏‫‮⁧]/u;

/** Text standing for an atom in the bidi algorithm. */
function standIn(a: Atom): string {
  switch (a.kind) {
    case 'char':
      return a.ch;
    case 'tab':
      return '\t';
    case 'break':
      return ' ';
    case 'image':
    case 'object':
    case 'field':
      return '￼';
    default:
      // Bookmarks, comment marks and anchors of floating objects do not take part.
      return '​';
  }
}

/** Resolve levels for a paragraph's atoms; nothing to do (undefined) for plain left-to-right text. */
export function analyzeBidi(atoms: Atom[], rtlParagraph: boolean): BidiInfo | undefined {
  let text = '';
  const offsets: number[] = [];
  let anyRtl = rtlParagraph;
  for (const a of atoms) {
    offsets.push(text.length);
    const s = standIn(a);
    if (!anyRtl && a.kind === 'char' && RTL.test(s)) anyRtl = true;
    text += s;
  }
  offsets.push(text.length);
  if (!anyRtl) return undefined;
  const { levels, paragraphLevel } = resolveLevels(text, rtlParagraph ? 1 : 0);
  const info = { text, levels, offsets, paragraphLevel };
  infos.set(atoms, info);
  return info;
}

export function bidiOf(atoms: Atom[]): BidiInfo | undefined {
  return infos.get(atoms);
}

/** Atom indices of a line in visual order (left to right on the page), with each atom's level. */
export function visualAtoms(info: BidiInfo, start: number, end: number): Array<{ atom: number; level: number }> {
  const s = info.offsets[start]!;
  const e = info.offsets[end]!;
  const levels = lineLevels(info.text, info.levels, info.paragraphLevel, s, e);
  const owner = new Int32Array(e - s);
  for (let i = start, k = s; i < end; i++) for (; k < info.offsets[i + 1]!; k++) owner[k - s] = i;
  const out: Array<{ atom: number; level: number }> = [];
  const seen = new Set<number>();
  for (const k of visualOrder(levels, s, e)) {
    const atom = owner[k - s]!;
    if (seen.has(atom)) continue;
    seen.add(atom);
    out.push({ atom, level: levels[k] ?? 0 });
  }
  return out;
}
