/**
 * GPOS: glyph positioning lookups. Types 1 single, 2 pair (formats 1 and 2: kerning),
 * 3 cursive attachment, 4 mark-to-base, 5 mark-to-ligature, 6 mark-to-mark, 7 context,
 * 8 chaining context, 9 extension. Device tables, variation deltas and anchor contour points
 * are ignored: layout works in font units, unhinted.
 *
 * Attachments are recorded as (type, relative parent index) and resolved into offsets after
 * all lookups ran, so a mark on a mark on a base picks up every adjustment of its chain.
 */
import { lookupProps, type ApplyContext, type LookupDispatcher, type Matcher } from './apply.js';
import { ATTACH_CURSIVE, ATTACH_MARK, isMark, isMultiplied, ligComp, type GlyphInfo } from './buffer.js';
import {
  BinaryData,
  classOf,
  LayoutTable,
  LookupFlag,
  parseClassDef,
  parseContextual,
  parseCoverage,
  type ClassDef,
  type ContextualSubtable,
  type Coverage,
} from './layout-common.js';

export interface Anchor {
  x: number;
  y: number;
}

interface MarkRecord {
  cls: number;
  anchor: Anchor | undefined;
}

export type GposSubtable =
  | { kind: 'single'; coverage: Coverage; valueFormat: number; valueAt: (index: number) => number }
  | { kind: 'pair1'; coverage: Coverage; vf1: number; vf2: number; pairSet: (index: number) => Map<number, number> | undefined }
  | {
      kind: 'pair2';
      coverage: Coverage;
      vf1: number;
      vf2: number;
      classDef1: ClassDef;
      classDef2: ClassDef;
      class1Count: number;
      class2Count: number;
      recordsAt: number;
    }
  | { kind: 'cursive'; coverage: Coverage; records: Array<{ entry: Anchor | undefined; exit: Anchor | undefined }> }
  | { kind: 'markBase'; markCoverage: Coverage; baseCoverage: Coverage; marks: MarkRecord[]; bases: Array<Array<Anchor | undefined>> }
  | { kind: 'markLig'; markCoverage: Coverage; ligCoverage: Coverage; marks: MarkRecord[]; ligatures: Array<Array<Array<Anchor | undefined>>> }
  | { kind: 'markMark'; markCoverage: Coverage; mark2Coverage: Coverage; marks: MarkRecord[]; mark2s: Array<Array<Anchor | undefined>> }
  | ContextualSubtable;

export interface GposLookup {
  type: number;
  props: number;
  subtables: GposSubtable[];
}

/** Number of 16-bit fields in a ValueRecord of this format. */
function valueLength(format: number): number {
  let n = 0;
  for (let f = format & 0xff; f; f >>= 1) n += f & 1;
  return n;
}

export class Gpos {
  readonly layout: LayoutTable;
  private readonly lookups = new Map<number, GposLookup | null>();
  readonly dispatch: LookupDispatcher;

  constructor(bytes: Uint8Array) {
    this.layout = new LayoutTable(bytes, 9);
    this.dispatch = (ctx, index, i) => {
      const lookup = this.lookup(index);
      if (!lookup) return -1;
      ctx.lookupProps = lookup.props;
      for (const st of lookup.subtables) {
        const next = applyPositioning(ctx, this.layout.data, st, i);
        if (next >= 0) return next;
      }
      return -1;
    };
  }

  lookup(index: number): GposLookup | undefined {
    const cached = this.lookups.get(index);
    if (cached !== undefined) return cached ?? undefined;
    const header = this.layout.lookupHeader(index);
    let lookup: GposLookup | null = null;
    if (header) {
      const subtables: GposSubtable[] = [];
      for (const st of header.subtables) {
        const parsed = parsePositioning(this.layout.data, st.type, st.offset);
        if (parsed) subtables.push(parsed);
      }
      lookup = { type: header.type, props: lookupProps(header.flag, header.markFilteringSet), subtables };
    }
    this.lookups.set(index, lookup);
    return lookup ?? undefined;
  }
}

function parseAnchor(d: BinaryData, off: number): Anchor | undefined {
  if (!off) return undefined;
  const format = d.u16(off);
  if (format < 1 || format > 3) return undefined;
  return { x: d.i16(off + 2), y: d.i16(off + 4) };
}

function parseMarkArray(d: BinaryData, off: number): MarkRecord[] {
  const count = d.u16(off);
  const out: MarkRecord[] = [];
  for (let i = 0; i < count; i++) {
    const r = off + 2 + i * 4;
    const rel = d.u16(r + 2);
    out.push({ cls: d.u16(r), anchor: rel ? parseAnchor(d, off + rel) : undefined });
  }
  return out;
}

/** BaseArray / Mark2Array / a LigatureAttach: rows of `classCount` anchor offsets. */
function parseAnchorMatrix(d: BinaryData, off: number, classCount: number): Array<Array<Anchor | undefined>> {
  const rows = d.u16(off);
  const out: Array<Array<Anchor | undefined>> = [];
  for (let r = 0; r < rows; r++) {
    const row: Array<Anchor | undefined> = [];
    for (let c = 0; c < classCount; c++) {
      const rel = d.u16(off + 2 + (r * classCount + c) * 2);
      row.push(rel ? parseAnchor(d, off + rel) : undefined);
    }
    out.push(row);
  }
  return out;
}

function parsePositioning(d: BinaryData, type: number, off: number): GposSubtable | undefined {
  const format = d.u16(off);
  const coverageAt = (at: number) => parseCoverage(d, off + d.u16(at));
  switch (type) {
    case 1: {
      const valueFormat = d.u16(off + 4);
      const len = valueLength(valueFormat) * 2;
      if (format === 1) return { kind: 'single', coverage: coverageAt(off + 2), valueFormat, valueAt: () => off + 6 };
      if (format === 2) return { kind: 'single', coverage: coverageAt(off + 2), valueFormat, valueAt: (i) => off + 8 + i * len };
      return undefined;
    }
    case 2: {
      const vf1 = d.u16(off + 4);
      const vf2 = d.u16(off + 6);
      if (format === 1) {
        const recordSize = (1 + valueLength(vf1) + valueLength(vf2)) * 2;
        const setCount = d.u16(off + 8);
        const sets = new Map<number, Map<number, number>>();
        return {
          kind: 'pair1',
          coverage: coverageAt(off + 2),
          vf1,
          vf2,
          // Pair sets are parsed on first use: big kerning tables are mostly never touched.
          pairSet: (index) => {
            if (index >= setCount) return undefined;
            let set = sets.get(index);
            if (!set) {
              set = new Map();
              const so = off + d.u16(off + 10 + index * 2);
              const n = d.u16(so);
              for (let k = 0; k < n; k++) {
                const r = so + 2 + k * recordSize;
                const second = d.u16(r);
                if (!set.has(second)) set.set(second, r + 2);
              }
              sets.set(index, set);
            }
            return set;
          },
        };
      }
      if (format === 2) {
        return {
          kind: 'pair2',
          coverage: coverageAt(off + 2),
          vf1,
          vf2,
          classDef1: parseClassDef(d, off + d.u16(off + 8)),
          classDef2: parseClassDef(d, off + d.u16(off + 10)),
          class1Count: d.u16(off + 12),
          class2Count: d.u16(off + 14),
          recordsAt: off + 16,
        };
      }
      return undefined;
    }
    case 3: {
      if (format !== 1) return undefined;
      const count = d.u16(off + 4);
      const records: Array<{ entry: Anchor | undefined; exit: Anchor | undefined }> = [];
      for (let i = 0; i < count; i++) {
        const r = off + 6 + i * 4;
        const entry = d.u16(r);
        const exit = d.u16(r + 2);
        records.push({ entry: entry ? parseAnchor(d, off + entry) : undefined, exit: exit ? parseAnchor(d, off + exit) : undefined });
      }
      return { kind: 'cursive', coverage: coverageAt(off + 2), records };
    }
    case 4:
    case 5:
    case 6: {
      if (format !== 1) return undefined;
      const classCount = d.u16(off + 6);
      const markCoverage = coverageAt(off + 2);
      const otherCoverage = coverageAt(off + 4);
      const marks = parseMarkArray(d, off + d.u16(off + 8));
      const arrayOff = off + d.u16(off + 10);
      if (type === 4) return { kind: 'markBase', markCoverage, baseCoverage: otherCoverage, marks, bases: parseAnchorMatrix(d, arrayOff, classCount) };
      if (type === 6) return { kind: 'markMark', markCoverage, mark2Coverage: otherCoverage, marks, mark2s: parseAnchorMatrix(d, arrayOff, classCount) };
      const ligCount = d.u16(arrayOff);
      const ligatures = d.u16s(arrayOff + 2, ligCount).map((rel) => parseAnchorMatrix(d, arrayOff + rel, classCount));
      return { kind: 'markLig', markCoverage, ligCoverage: otherCoverage, marks, ligatures };
    }
    case 7:
    case 8:
      return parseContextual(d, off, type === 8);
    default:
      return undefined;
  }
}

/** Adds a ValueRecord to a glyph (horizontal layout: Y advance is ignored). */
function applyValue(d: BinaryData, off: number, format: number, info: GlyphInfo): void {
  let p = off;
  if (format & 0x0001) {
    info.xOffset += d.i16(p);
    p += 2;
  }
  if (format & 0x0002) {
    info.yOffset += d.i16(p);
    p += 2;
  }
  if (format & 0x0004) info.xAdvance += d.i16(p);
}

function attachMark(ctx: ApplyContext, i: number, mark: MarkRecord | undefined, target: Anchor | undefined, j: number): number {
  if (!mark?.anchor || !target) return -1; // let a later subtable try
  const info = ctx.buffer[i]!;
  info.xOffset = Math.round(target.x - mark.anchor.x);
  info.yOffset = Math.round(target.y - mark.anchor.y);
  info.attachType = ATTACH_MARK;
  info.attachChain = j - i;
  return i + 1;
}

function applyPositioning(ctx: ApplyContext, d: BinaryData, st: GposSubtable, i: number): number {
  const buf = ctx.buffer;
  const info = buf[i]!;
  const glyph = info.glyph;
  switch (st.kind) {
    case 'single': {
      const ci = st.coverage.get(glyph);
      if (ci === undefined) return -1;
      applyValue(d, st.valueAt(ci), st.valueFormat, info);
      return i + 1;
    }
    case 'pair1':
    case 'pair2': {
      const ci = st.coverage.get(glyph);
      if (ci === undefined) return -1;
      const j = ctx.next(i, ctx.inputMatcher(i));
      if (j < 0) return -1;
      const second = buf[j]!;
      const len1 = valueLength(st.vf1) * 2;
      let rec: number | undefined;
      if (st.kind === 'pair1') {
        rec = st.pairSet(ci)?.get(second.glyph);
      } else {
        const k1 = classOf(st.classDef1, glyph);
        const k2 = classOf(st.classDef2, second.glyph);
        if (k1 < st.class1Count && k2 < st.class2Count) rec = st.recordsAt + (k1 * st.class2Count + k2) * (len1 + valueLength(st.vf2) * 2);
      }
      if (rec === undefined) return -1;
      applyValue(d, rec, st.vf1, info);
      applyValue(d, rec + len1, st.vf2, second);
      // A pair that also moved the second glyph consumes it; otherwise it may pair again.
      return valueLength(st.vf2) ? j + 1 : j;
    }
    case 'cursive':
      return applyCursive(ctx, st, i);
    case 'markBase': {
      const mi = st.markCoverage.get(glyph);
      if (mi === undefined) return -1;
      const m = baseMatcher(ctx);
      let j = i;
      for (;;) {
        j = ctx.prev(j, m);
        if (j < 0) return -1;
        // Attach to the first glyph of a multiple-substitution sequence, not the later ones.
        const b = buf[j]!;
        const p = buf[j - 1];
        if (!isMultiplied(b) || ligComp(b) === 0 || !p || isMark(p) || !isMultiplied(p) || b.ligId !== p.ligId || ligComp(b) !== ligComp(p) + 1) break;
      }
      const bi = st.baseCoverage.get(buf[j]!.glyph);
      if (bi === undefined) return -1;
      const mark = st.marks[mi];
      return attachMark(ctx, i, mark, mark && st.bases[bi]?.[mark.cls], j);
    }
    case 'markLig': {
      const mi = st.markCoverage.get(glyph);
      if (mi === undefined) return -1;
      const j = ctx.prev(i, baseMatcher(ctx));
      if (j < 0) return -1;
      const lig = buf[j]!;
      const li = st.ligCoverage.get(lig.glyph);
      const comps = li === undefined ? undefined : st.ligatures[li];
      if (!comps?.length) return -1;
      // A mark that came between the ligature's components carries its component number.
      const markComp = ligComp(info);
      const comp = lig.ligId && lig.ligId === info.ligId && markComp > 0 ? Math.min(comps.length, markComp) - 1 : comps.length - 1;
      const mark = st.marks[mi];
      return attachMark(ctx, i, mark, mark && comps[comp]?.[mark.cls], j);
    }
    case 'markMark': {
      const mi = st.markCoverage.get(glyph);
      if (mi === undefined) return -1;
      const m: Matcher = { props: ctx.lookupProps & ~LookupFlag.IgnoreFlags, ignoreZwnj: true, ignoreZwj: true, mask: ctx.lookupMask, syllable: 0 };
      const j = ctx.prev(i, m);
      if (j < 0 || !isMark(buf[j]!)) return -1;
      const prev = buf[j]!;
      const id1 = info.ligId;
      const id2 = prev.ligId;
      const comp1 = ligComp(info);
      const comp2 = ligComp(prev);
      // Both marks must sit on the same base or the same ligature component.
      const sameAnchorGroup = id1 === id2 ? id1 === 0 || comp1 === comp2 : (id1 > 0 && !comp1) || (id2 > 0 && !comp2);
      if (!sameAnchorGroup) return -1;
      const m2 = st.mark2Coverage.get(prev.glyph);
      if (m2 === undefined) return -1;
      const mark = st.marks[mi];
      return attachMark(ctx, i, mark, mark && st.mark2s[m2]?.[mark.cls], j);
    }
    default:
      return ctx.applyContextual(st, i);
  }
}

/** Mark-to-base / mark-to-ligature look back past marks regardless of the lookup's flags. */
function baseMatcher(ctx: ApplyContext): Matcher {
  return { props: LookupFlag.IgnoreMarks, ignoreZwnj: true, ignoreZwj: true, mask: ctx.lookupMask, syllable: 0 };
}

function applyCursive(ctx: ApplyContext, st: Extract<GposSubtable, { kind: 'cursive' }>, i: number): number {
  const buf = ctx.buffer;
  const ci = st.coverage.get(buf[i]!.glyph);
  const thisRec = ci === undefined ? undefined : st.records[ci];
  if (!thisRec?.entry) return -1;
  const p = ctx.prev(i, ctx.inputMatcher(i));
  if (p < 0) return -1;
  const pci = st.coverage.get(buf[p]!.glyph);
  const prevRec = pci === undefined ? undefined : st.records[pci];
  if (!prevRec?.exit) return -1;

  const exit = prevRec.exit;
  const entry = thisRec.entry;
  const a = buf[p]!;
  const b = buf[i]!;
  // Main direction: the exit point of the previous glyph meets the entry of this one.
  if (ctx.direction === 'ltr') {
    a.xAdvance = Math.round(exit.x) + a.xOffset;
    const d = Math.round(entry.x) + b.xOffset;
    b.xAdvance -= d;
    b.xOffset -= d;
  } else {
    const d = Math.round(exit.x) + a.xOffset;
    a.xAdvance -= d;
    a.xOffset -= d;
    b.xAdvance = Math.round(entry.x) + b.xOffset;
  }

  // Cross direction: the child hangs off its parent; with RightToLeft the logically earlier
  // glyph is the child (the last glyph of a word stays on the baseline).
  let child = p;
  let parent = i;
  let yOffset = Math.round(entry.y - exit.y);
  if (!(ctx.lookupProps & LookupFlag.RightToLeft)) {
    child = i;
    parent = p;
    yOffset = -yOffset;
  }
  reverseCursiveMinorOffset(buf, child, parent, 0);
  const c = buf[child]!;
  c.attachType = ATTACH_CURSIVE;
  c.attachChain = parent - child;
  c.yOffset = yOffset;
  const par = buf[parent]!;
  if (par.attachChain === -c.attachChain) {
    par.attachChain = 0;
    par.yOffset = 0;
  }
  return i + 1;
}

/** When a child gets a new parent, flip its old chain so the whole tree follows. */
function reverseCursiveMinorOffset(buf: GlyphInfo[], i: number, newParent: number, depth: number): void {
  const info = buf[i]!;
  const chain = info.attachChain;
  const type = info.attachType;
  if (!chain || !(type & ATTACH_CURSIVE) || depth > 1024) return;
  info.attachChain = 0;
  const j = i + chain;
  if (j === newParent || j < 0 || j >= buf.length) return;
  reverseCursiveMinorOffset(buf, j, newParent, depth + 1);
  const parent = buf[j]!;
  parent.yOffset = -info.yOffset;
  parent.attachChain = -chain;
  parent.attachType = type;
}

/**
 * Turns attachment chains into final offsets: a mark lands at its base's offset minus the
 * advances between them (in logical order; for RTL the run is later drawn reversed, so the
 * glyphs after the base, up to and including the mark, are the ones in between).
 */
export function propagateAttachments(buf: GlyphInfo[], direction: 'ltr' | 'rtl'): void {
  for (let i = 0; i < buf.length; i++) propagate(buf, i, direction, 0);
}

function propagate(buf: GlyphInfo[], i: number, direction: 'ltr' | 'rtl', depth: number): void {
  const info = buf[i]!;
  const chain = info.attachChain;
  if (!chain) return;
  info.attachChain = 0;
  const j = i + chain;
  if (j < 0 || j >= buf.length || depth > 1024) return;
  propagate(buf, j, direction, depth + 1);
  const parent = buf[j]!;
  if (info.attachType & ATTACH_CURSIVE) {
    info.yOffset += parent.yOffset;
    return;
  }
  info.xOffset += parent.xOffset;
  info.yOffset += parent.yOffset;
  if (direction === 'ltr') {
    for (let k = j; k < i; k++) info.xOffset -= buf[k]!.xAdvance;
  } else {
    for (let k = j + 1; k <= i; k++) info.xOffset += buf[k]!.xAdvance;
  }
}
