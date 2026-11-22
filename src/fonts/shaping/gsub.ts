/**
 * GSUB: glyph substitution lookups. Types 1 single, 2 multiple, 3 alternate (first
 * alternate), 4 ligature, 5 context, 6 chaining context, 7 extension (resolved by
 * LayoutTable) and 8 reverse chaining single. Substitutions edit the buffer in place and
 * keep the glyph-property history (substituted / ligated / multiplied) and ligature
 * component numbers that mark positioning and the Indic final reordering rely on.
 */
import { MAX_NESTING, lookupProps, type ApplyContext, type LookupDispatcher } from './apply.js';
import {
  cloneInfo,
  deleteGlyph,
  isBaseGlyph,
  isLigature,
  isMark,
  ligComp,
  ligNumComps,
  mergeClusters,
  setLigPropsForLigature,
  setLigPropsForMark,
  type GlyphInfo,
} from './buffer.js';
import { GlyphProps } from './gdef.js';
import { classOf, LayoutTable, parseContextual, parseCoverage, type ContextualSubtable, type Coverage } from './layout-common.js';
import { UFlag } from './unicode.js';

export interface Ligature {
  glyph: number;
  /** Components after the first (the first is matched by coverage). */
  components: number[];
}

export type GsubSubtable =
  | { kind: 'single'; coverage: Coverage; delta: number | undefined; substitutes: number[] }
  | { kind: 'multiple'; coverage: Coverage; sequences: number[][] }
  | { kind: 'alternate'; coverage: Coverage; sets: number[][] }
  | { kind: 'ligature'; coverage: Coverage; sets: Ligature[][] }
  | { kind: 'reverse'; coverage: Coverage; backtrack: Coverage[]; lookahead: Coverage[]; substitutes: number[] }
  | ContextualSubtable;

export interface GsubLookup {
  type: number;
  /** Packed lookup flags + mark filtering set (see lookupProps). */
  props: number;
  reverse: boolean;
  subtables: GsubSubtable[];
}

export class Gsub {
  readonly layout: LayoutTable;
  private readonly lookups = new Map<number, GsubLookup | null>();
  /** Applies one of this table's lookups at a position (also used for nested lookups). */
  readonly dispatch: LookupDispatcher;

  constructor(bytes: Uint8Array) {
    this.layout = new LayoutTable(bytes, 7);
    this.dispatch = (ctx, index, i) => {
      const lookup = this.lookup(index);
      if (!lookup) return -1;
      ctx.lookupProps = lookup.props;
      for (const st of lookup.subtables) {
        const next = applySubstitution(ctx, st, i);
        if (next >= 0) return next;
      }
      return -1;
    };
  }

  lookup(index: number): GsubLookup | undefined {
    const cached = this.lookups.get(index);
    if (cached !== undefined) return cached ?? undefined;
    const header = this.layout.lookupHeader(index);
    let lookup: GsubLookup | null = null;
    if (header) {
      const subtables: GsubSubtable[] = [];
      for (const st of header.subtables) {
        const parsed = parseSubstitution(this.layout, st.type, st.offset);
        if (parsed) subtables.push(parsed);
      }
      lookup = { type: header.type, props: lookupProps(header.flag, header.markFilteringSet), reverse: header.type === 8, subtables };
    }
    this.lookups.set(index, lookup);
    return lookup ?? undefined;
  }

  /**
   * Whether any of the lookups would substitute exactly this glyph sequence (ignoring its
   * surroundings). The Indic shaper asks this to learn, e.g., whether Ra+Halant forms a reph
   * or a consonant has a below-base form, before deciding the syllable's base.
   */
  wouldSubstitute(lookupIndices: readonly number[], glyphs: readonly number[], zeroContext: boolean): boolean {
    if (!glyphs.length) return false;
    for (const index of lookupIndices) {
      const lookup = this.lookup(index);
      if (lookup?.subtables.some((st) => wouldApply(st, glyphs, zeroContext))) return true;
    }
    return false;
  }
}

function parseSubstitution(layout: LayoutTable, type: number, off: number): GsubSubtable | undefined {
  const d = layout.data;
  const format = d.u16(off);
  const coverage = () => parseCoverage(d, off + d.u16(off + 2));
  const offsetsAt = (at: number) => d.u16s(at + 2, d.u16(at));
  const glyphArray = (at: number) => d.u16s(at + 2, d.u16(at));
  switch (type) {
    case 1:
      if (format === 1) return { kind: 'single', coverage: coverage(), delta: d.i16(off + 4), substitutes: [] };
      if (format === 2) return { kind: 'single', coverage: coverage(), delta: undefined, substitutes: glyphArray(off + 4) };
      return undefined;
    case 2:
    case 3: {
      if (format !== 1) return undefined;
      const lists = offsetsAt(off + 4).map((rel) => glyphArray(off + rel));
      return type === 2 ? { kind: 'multiple', coverage: coverage(), sequences: lists } : { kind: 'alternate', coverage: coverage(), sets: lists };
    }
    case 4: {
      if (format !== 1) return undefined;
      const sets = offsetsAt(off + 4).map((rel) => {
        const setOff = off + rel;
        return offsetsAt(setOff).map((lrel): Ligature => {
          const l = setOff + lrel;
          return { glyph: d.u16(l), components: d.u16s(l + 4, Math.max(0, d.u16(l + 2) - 1)) };
        });
      });
      return { kind: 'ligature', coverage: coverage(), sets };
    }
    case 5:
    case 6:
      return parseContextual(d, off, type === 6);
    case 8: {
      if (format !== 1) return undefined;
      let p = off + 4;
      const covs = (at: number, count: number) => d.u16s(at, count).map((o) => parseCoverage(d, off + o));
      const bc = d.u16(p);
      const backtrack = covs(p + 2, bc);
      p += 2 + bc * 2;
      const lc = d.u16(p);
      const lookahead = covs(p + 2, lc);
      p += 2 + lc * 2;
      return { kind: 'reverse', coverage: coverage(), backtrack, lookahead, substitutes: glyphArray(p) };
    }
    default:
      return undefined;
  }
}

/**
 * Sets a glyph and updates its class: from GDEF when the font has glyph classes, else from
 * the caller's guess (ligature / base), keeping the substitution history bits.
 */
export function replaceGlyph(ctx: ApplyContext, info: GlyphInfo, glyph: number, classGuess = 0, ligature = false, component = false): void {
  let props = info.props | GlyphProps.Substituted;
  if (ligature) props = (props | GlyphProps.Ligated) & ~GlyphProps.Multiplied;
  if (component) props |= GlyphProps.Multiplied;
  if (ctx.gdef.hasGlyphClasses) props = (props & GlyphProps.Preserve) | ctx.gdef.glyphProps(glyph);
  else if (classGuess) props = (props & GlyphProps.Preserve) | classGuess;
  info.props = props;
  info.glyph = glyph;
}

function applySubstitution(ctx: ApplyContext, st: GsubSubtable, i: number): number {
  const buf = ctx.buffer;
  const info = buf[i]!;
  const glyph = info.glyph;
  switch (st.kind) {
    case 'single': {
      const ci = st.coverage.get(glyph);
      if (ci === undefined) return -1;
      const out = st.delta !== undefined ? (glyph + st.delta) & 0xffff : st.substitutes[ci];
      if (out === undefined) return -1;
      replaceGlyph(ctx, info, out);
      return i + 1;
    }
    case 'alternate': {
      const ci = st.coverage.get(glyph);
      const out = ci === undefined ? undefined : st.sets[ci]?.[0];
      if (out === undefined) return -1;
      replaceGlyph(ctx, info, out);
      return i + 1;
    }
    case 'multiple': {
      const ci = st.coverage.get(glyph);
      const seq = ci === undefined ? undefined : st.sequences[ci];
      if (!seq) return -1;
      if (seq.length === 1) {
        // Not a real multiplication: keep it a plain substitution.
        replaceGlyph(ctx, info, seq[0]!);
        return i + 1;
      }
      if (seq.length === 0) {
        // Disallowed by the spec but used by fonts (and honoured by Uniscribe) to delete.
        deleteGlyph(buf, i);
        return i;
      }
      const klass = isLigature(info) ? GlyphProps.BaseGlyph : 0;
      const keepLig = info.ligId !== 0; // still attached to an earlier ligature: leave that be
      const out = seq.map((g, k) => {
        const c = cloneInfo(info);
        if (!keepLig) setLigPropsForMark(c, 0, k);
        replaceGlyph(ctx, c, g, klass, false, true);
        return c;
      });
      buf.splice(i, 1, ...out);
      return i + out.length;
    }
    case 'ligature': {
      const ci = st.coverage.get(glyph);
      const set = ci === undefined ? undefined : st.sets[ci];
      if (!set) return -1;
      for (const lig of set) {
        if (!lig.components.length) {
          replaceGlyph(ctx, info, lig.glyph);
          return i + 1;
        }
        const match = ctx.matchInput(i, lig.components.length + 1, (g, k) => g.glyph === lig.components[k - 1]);
        if (match) return ligate(ctx, match.positions, match.end, lig.glyph, match.totalComponents);
      }
      return -1;
    }
    case 'reverse': {
      // Only applied by the reverse pass itself, never as a nested lookup.
      if (ctx.nestingLeft !== MAX_NESTING) return -1;
      const ci = st.coverage.get(glyph);
      const out = ci === undefined ? undefined : st.substitutes[ci];
      if (out === undefined) return -1;
      if (
        ctx.matchBacktrack(i, st.backtrack.length, (g, k) => st.backtrack[k]?.has(g.glyph) ?? false) &&
        ctx.matchLookahead(i, st.lookahead.length, (g, k) => st.lookahead[k]?.has(g.glyph) ?? false)
      ) {
        replaceGlyph(ctx, info, out);
        return i + 1;
      }
      return -1;
    }
    default:
      return ctx.applyContextual(st, i);
  }
}

/**
 * Replaces the matched components with the ligature glyph. Marks skipped between components
 * stay (after the ligature) and are renumbered to the component they followed, so GPOS
 * mark-to-ligature can put each on the right part of the ligature.
 */
function ligate(ctx: ApplyContext, positions: number[], matchEnd: number, ligGlyph: number, totalComponents: number): number {
  const buf = ctx.buffer;
  const count = positions.length;
  const first = buf[positions[0]!]!;
  mergeClusters(buf, positions[0]!, matchEnd);

  // A base + marks ligature stays a base (later marks still attach to it); an all-marks
  // ligature keeps its ligature id so it can still sit on an earlier ligature component.
  let isBaseLigature = isBaseGlyph(first);
  let isMarkLigature = isMark(first);
  for (let k = 1; k < count; k++) {
    if (!isMark(buf[positions[k]!]!)) {
      isBaseLigature = false;
      isMarkLigature = false;
      break;
    }
  }
  const isLig = !isBaseLigature && !isMarkLigature;
  const klass = isLig ? GlyphProps.Ligature : 0;
  const ligId = isLig ? ctx.allocLigId() : 0;
  let lastLigId = first.ligId;
  let lastNumComps = ligNumComps(first);
  let compsSoFar = lastNumComps;

  if (isLig) {
    setLigPropsForLigature(first, ligId, totalComponents);
    first.uflags &= ~UFlag.NonSpacingMark;
  }
  replaceGlyph(ctx, first, ligGlyph, klass, true);

  for (let k = 1; k < count; k++) {
    if (isLig) {
      for (let j = positions[k - 1]! + 1; j < positions[k]!; j++) {
        const info = buf[j]!;
        let thisComp = ligComp(info);
        if (thisComp === 0) thisComp = lastNumComps;
        setLigPropsForMark(info, ligId, compsSoFar - lastNumComps + Math.min(thisComp, lastNumComps));
      }
    }
    const comp = buf[positions[k]!]!;
    lastLigId = comp.ligId;
    lastNumComps = ligNumComps(comp);
    compsSoFar += lastNumComps;
  }

  if (!isMarkLigature && lastLigId) {
    // Marks that followed the last component (itself a ligature) move to the new ligature.
    for (let j = positions[count - 1]! + 1; j < buf.length; j++) {
      const info = buf[j]!;
      if (info.ligId !== lastLigId) break;
      const thisComp = ligComp(info);
      if (!thisComp) break;
      setLigPropsForMark(info, ligId, compsSoFar - lastNumComps + Math.min(thisComp, lastNumComps));
    }
  }

  for (let k = count - 1; k >= 1; k--) buf.splice(positions[k]!, 1);
  return positions[count - 1]! - (count - 1) + 1;
}

function wouldApply(st: GsubSubtable, glyphs: readonly number[], zeroContext: boolean): boolean {
  const g0 = glyphs[0]!;
  const noContext = (r: { backtrack: unknown[]; lookahead: unknown[] }) => !zeroContext || (!r.backtrack.length && !r.lookahead.length);
  switch (st.kind) {
    case 'single':
    case 'multiple':
    case 'alternate':
      return glyphs.length === 1 && st.coverage.has(g0);
    case 'reverse':
      return glyphs.length === 1 && st.coverage.has(g0) && noContext(st);
    case 'ligature': {
      const ci = st.coverage.get(g0);
      const set = ci === undefined ? undefined : st.sets[ci];
      return !!set?.some((l) => l.components.length + 1 === glyphs.length && l.components.every((c, k) => c === glyphs[k + 1]));
    }
    case 'ctx-glyphs': {
      const ci = st.coverage.get(g0);
      const rules = ci === undefined ? undefined : st.ruleSets[ci];
      return !!rules?.some((r) => noContext(r) && r.input.length + 1 === glyphs.length && r.input.every((v, k) => v === glyphs[k + 1]));
    }
    case 'ctx-classes': {
      if (!st.coverage.has(g0)) return false;
      const rules = st.ruleSets[classOf(st.inputClasses, g0)];
      return !!rules?.some(
        (r) => noContext(r) && r.input.length + 1 === glyphs.length && r.input.every((v, k) => classOf(st.inputClasses, glyphs[k + 1]!) === v),
      );
    }
    case 'ctx-coverage':
      return noContext(st) && st.input.length === glyphs.length && st.input.every((c, k) => c.has(glyphs[k]!));
  }
}
