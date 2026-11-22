/**
 * Lookup application machinery shared by GSUB and GPOS: lookup-flag glyph filtering, the
 * "skipping iterator" that walks past ignored glyphs (marks for a ligature, default-ignorables,
 * glyphs of other syllables), contextual rule matching and nested-lookup application.
 *
 * The semantics follow HarfBuzz's hb_ot_apply_context_t so fonts tuned against it (and
 * Uniscribe/DirectWrite, which it mirrors) behave the same: e.g. ZWJ is transparent to
 * input matching unless a feature asks for manual joiners, ZWNJ is always visible to GSUB
 * input but ignored by GPOS.
 */
import { ligComp, ligNumComps, type GlyphInfo } from './buffer.js';
import { GlyphProps, type Gdef } from './gdef.js';
import { classOf, LookupFlag, type ContextualSubtable, type LookupRecord } from './layout-common.js';
import { UFlag } from './unicode.js';

export const MAX_NESTING = 16;
const MAX_CONTEXT = 64;

const SKIP_NO = 0;
const SKIP_YES = 1;
const SKIP_MAYBE = 2;
const MATCH_NO = 0;
const MATCH_YES = 1;
const MATCH_MAYBE = 2;

/** How the skipping iterator treats glyphs for one match. */
export interface Matcher {
  /** LookupFlag bits, with the mark filtering set index in the high 16 bits. */
  props: number;
  ignoreZwnj: boolean;
  ignoreZwj: boolean;
  /** Feature mask a matched glyph must carry (-1: any). */
  mask: number;
  /** Syllable a matched glyph must belong to (0: any). */
  syllable: number;
}

type MatchFn = (info: GlyphInfo) => boolean;

export interface InputMatch {
  positions: number[];
  end: number;
  totalComponents: number;
}

/** Applies one lookup (by index) at a buffer position; returns the next position or -1. */
export type LookupDispatcher = (ctx: ApplyContext, lookupIndex: number, i: number) => number;

/** Lookup props as HarfBuzz packs them: flags plus the mark filtering set in the high half. */
export function lookupProps(flag: number, markFilteringSet: number): number {
  return (flag & 0xffff) | (flag & LookupFlag.UseMarkFilteringSet ? (markFilteringSet & 0xffff) * 0x10000 : 0);
}

export class ApplyContext {
  lookupMask = 1;
  lookupProps = 0;
  autoZwj = true;
  autoZwnj = true;
  perSyllable = false;
  nestingLeft = MAX_NESTING;
  /** Budget of nested-lookup calls, against fonts whose contextual lookups recurse wildly. */
  opsLeft: number;
  private nextLigId = 1;

  /**
   * @param tableIndex 0 for GSUB, 1 for GPOS
   * @param dispatch applies a lookup of this table once at a position (used for nesting)
   */
  constructor(
    readonly buffer: GlyphInfo[],
    readonly gdef: Gdef,
    readonly tableIndex: 0 | 1,
    readonly direction: 'ltr' | 'rtl',
    readonly dispatch: LookupDispatcher,
  ) {
    this.opsLeft = Math.max(16384, buffer.length * 256);
  }

  allocLigId(): number {
    return this.nextLigId++;
  }

  /** Whether lookup flags let a glyph through (LookupFlag Ignore*, mark filtering). */
  checkGlyphProperty(info: GlyphInfo, props: number): boolean {
    const gp = info.props;
    if (gp & props & LookupFlag.IgnoreFlags) return false;
    if (gp & GlyphProps.Mark) {
      if (props & LookupFlag.UseMarkFilteringSet) return this.gdef.markSetCovers(props >>> 16, info.glyph);
      if (props & LookupFlag.MarkAttachmentType) return (props & LookupFlag.MarkAttachmentType) === (gp & GlyphProps.MarkAttachClass);
    }
    return true;
  }

  inputMatcher(start: number): Matcher {
    return {
      props: this.lookupProps,
      ignoreZwnj: this.tableIndex === 1,
      ignoreZwj: this.tableIndex === 1 || this.autoZwj,
      mask: this.lookupMask,
      syllable: this.tableIndex === 0 && this.perSyllable ? (this.buffer[start]?.syllable ?? 0) : 0,
    };
  }

  contextMatcher(start: number): Matcher {
    return {
      props: this.lookupProps,
      ignoreZwnj: this.tableIndex === 1 || this.autoZwnj,
      ignoreZwj: true,
      mask: -1,
      syllable: this.tableIndex === 0 && this.perSyllable ? (this.buffer[start]?.syllable ?? 0) : 0,
    };
  }

  private maySkip(info: GlyphInfo, m: Matcher): number {
    if (!this.checkGlyphProperty(info, m.props)) return SKIP_YES;
    const u = info.uflags;
    if (
      u & UFlag.DefaultIgnorable &&
      !(info.props & GlyphProps.Substituted) &&
      (m.ignoreZwnj || !(u & UFlag.Zwnj)) &&
      (m.ignoreZwj || !(u & UFlag.Zwj))
    ) {
      return SKIP_MAYBE;
    }
    return SKIP_NO;
  }

  private mayMatch(info: GlyphInfo, m: Matcher, fn: MatchFn | undefined): number {
    if (!(info.mask & m.mask) || (m.syllable && m.syllable !== info.syllable)) return MATCH_NO;
    if (fn) return fn(info) ? MATCH_YES : MATCH_NO;
    return MATCH_MAYBE;
  }

  /** Next glyph after `i` that isn't skipped and matches `fn`, or -1 at a blocking glyph. */
  next(i: number, m: Matcher, fn?: MatchFn): number {
    const buf = this.buffer;
    for (let j = i + 1; j < buf.length; j++) {
      const info = buf[j]!;
      const skip = this.maySkip(info, m);
      if (skip === SKIP_YES) continue;
      const match = this.mayMatch(info, m, fn);
      if (match === MATCH_YES || (match === MATCH_MAYBE && skip === SKIP_NO)) return j;
      if (skip === SKIP_NO) return -1;
    }
    return -1;
  }

  prev(i: number, m: Matcher, fn?: MatchFn): number {
    const buf = this.buffer;
    for (let j = i - 1; j >= 0; j--) {
      const info = buf[j]!;
      const skip = this.maySkip(info, m);
      if (skip === SKIP_YES) continue;
      const match = this.mayMatch(info, m, fn);
      if (match === MATCH_YES || (match === MATCH_MAYBE && skip === SKIP_NO)) return j;
      if (skip === SKIP_NO) return -1;
    }
    return -1;
  }

  /**
   * Matches `count` input glyphs starting at `i` (which the caller already matched as
   * component 0). Ligature-id checks keep components of different earlier ligatures (and
   * marks attached to them) from being glued together.
   */
  matchInput(i: number, count: number, fn: (info: GlyphInfo, k: number) => boolean): InputMatch | undefined {
    const buf = this.buffer;
    const m = this.inputMatcher(i);
    const first = buf[i]!;
    const firstLigId = first.ligId;
    const firstLigComp = ligComp(first);
    let total = ligNumComps(first);
    const positions = [i];
    let j = i;
    for (let k = 1; k < count; k++) {
      j = this.next(j, m, (info) => fn(info, k));
      if (j < 0) return undefined;
      const info = buf[j]!;
      const thisLigId = info.ligId;
      const thisLigComp = ligComp(info);
      if (firstLigId && firstLigComp) {
        // All components must hang off the same earlier ligature component, unless that
        // ligature itself is ignored by this lookup.
        if ((firstLigId !== thisLigId || firstLigComp !== thisLigComp) && !this.ligBaseSkippable(i, firstLigId, m)) return undefined;
      } else if (thisLigId && thisLigComp && thisLigId !== firstLigId) {
        return undefined;
      }
      total += ligNumComps(info);
      positions.push(j);
    }
    return { positions, end: j + 1, totalComponents: total };
  }

  private ligBaseSkippable(i: number, ligId: number, m: Matcher): boolean {
    const buf = this.buffer;
    for (let j = i - 1; j >= 0 && buf[j]!.ligId === ligId; j--) {
      if (ligComp(buf[j]!) === 0) return this.maySkip(buf[j]!, m) === SKIP_YES;
    }
    return false;
  }

  /** Backtrack glyphs before `i`; `fn(info, 0)` is the nearest one. */
  matchBacktrack(i: number, count: number, fn: (info: GlyphInfo, k: number) => boolean): boolean {
    const m = this.contextMatcher(i);
    let j = i;
    for (let k = 0; k < count; k++) {
      j = this.prev(j, m, (info) => fn(info, k));
      if (j < 0) return false;
    }
    return true;
  }

  /** Lookahead glyphs after `last` (the last input glyph). */
  matchLookahead(last: number, count: number, fn: (info: GlyphInfo, k: number) => boolean): boolean {
    const m = this.contextMatcher(last);
    let j = last;
    for (let k = 0; k < count; k++) {
      j = this.next(j, m, (info) => fn(info, k));
      if (j < 0) return false;
    }
    return true;
  }

  /** Applies a nested lookup at one position, with that lookup's own flags. */
  recurse(lookupIndex: number, pos: number): boolean {
    if (this.nestingLeft <= 0 || this.opsLeft-- <= 0) return false;
    const savedProps = this.lookupProps;
    this.nestingLeft--;
    const applied = this.dispatch(this, lookupIndex, pos) >= 0;
    this.nestingLeft++;
    this.lookupProps = savedProps;
    return applied;
  }

  /**
   * Runs a contextual rule's nested lookups over the matched input. Nested substitutions may
   * grow or shrink the buffer, so match positions after the edited one are shifted the way
   * HarfBuzz does it. Returns the (adjusted) end of the matched input.
   */
  applyNested(records: readonly LookupRecord[], positions: number[], matchEnd: number): number {
    const buf = this.buffer;
    let count = positions.length;
    let end = matchEnd;
    for (const rec of records) {
      const idx = rec.sequenceIndex;
      if (idx >= count) continue;
      const origLen = buf.length;
      const at = positions[idx]!;
      if (at >= origLen) continue;
      if (!this.recurse(rec.lookupIndex, at)) continue;
      let delta = buf.length - origLen;
      if (!delta) continue;
      end += delta;
      if (end < at) {
        delta += at - end;
        end = at;
      }
      let next = idx + 1;
      if (delta > 0) {
        if (delta + count > MAX_CONTEXT) break;
      } else {
        delta = Math.max(delta, next - count);
        next -= delta;
      }
      const moved = positions.slice(next, count);
      for (let k = 0; k < moved.length; k++) positions[next + delta + k] = moved[k]!;
      next += delta;
      count += delta;
      positions.length = count;
      for (let j = idx + 1; j < next; j++) positions[j] = positions[j - 1]! + 1;
      for (; next < count; next++) positions[next] = positions[next]! + delta;
    }
    return end;
  }

  /** Context / chaining-context subtable at position `i`. */
  applyContextual(st: ContextualSubtable, i: number): number {
    const glyph = this.buffer[i]!.glyph;
    switch (st.kind) {
      case 'ctx-glyphs': {
        const ci = st.coverage.get(glyph);
        const rules = ci === undefined ? undefined : st.ruleSets[ci];
        if (!rules) return -1;
        for (const r of rules) {
          const next = this.applyRule(
            i,
            r.input.length + 1,
            (info, k) => info.glyph === r.input[k - 1],
            r.backtrack.length,
            (info, k) => info.glyph === r.backtrack[k],
            r.lookahead.length,
            (info, k) => info.glyph === r.lookahead[k],
            r.lookups,
          );
          if (next >= 0) return next;
        }
        return -1;
      }
      case 'ctx-classes': {
        if (!st.coverage.has(glyph)) return -1;
        const rules = st.ruleSets[classOf(st.inputClasses, glyph)];
        if (!rules) return -1;
        for (const r of rules) {
          const next = this.applyRule(
            i,
            r.input.length + 1,
            (info, k) => classOf(st.inputClasses, info.glyph) === r.input[k - 1],
            r.backtrack.length,
            (info, k) => classOf(st.backtrackClasses, info.glyph) === r.backtrack[k],
            r.lookahead.length,
            (info, k) => classOf(st.lookaheadClasses, info.glyph) === r.lookahead[k],
            r.lookups,
          );
          if (next >= 0) return next;
        }
        return -1;
      }
      case 'ctx-coverage': {
        if (!st.input[0]?.has(glyph)) return -1;
        return this.applyRule(
          i,
          st.input.length,
          (info, k) => st.input[k]?.has(info.glyph) ?? false,
          st.backtrack.length,
          (info, k) => st.backtrack[k]?.has(info.glyph) ?? false,
          st.lookahead.length,
          (info, k) => st.lookahead[k]?.has(info.glyph) ?? false,
          st.lookups,
        );
      }
    }
  }

  private applyRule(
    i: number,
    inputCount: number,
    inputFn: (info: GlyphInfo, k: number) => boolean,
    backCount: number,
    backFn: (info: GlyphInfo, k: number) => boolean,
    aheadCount: number,
    aheadFn: (info: GlyphInfo, k: number) => boolean,
    records: readonly LookupRecord[],
  ): number {
    const match = this.matchInput(i, inputCount, inputFn);
    if (!match) return -1;
    if (!this.matchLookahead(match.end - 1, aheadCount, aheadFn)) return -1;
    if (!this.matchBacktrack(i, backCount, backFn)) return -1;
    // The end only moves back to `i` when nested lookups deleted glyphs, so this always
    // makes progress.
    return this.applyNested(records, match.positions, match.end);
  }

  /**
   * Applies a lookup over the whole buffer (or backwards for reverse-chaining lookups),
   * only at glyphs carrying `mask` and passing the lookup flags.
   */
  applyLookup(lookupIndex: number, props: number, mask: number, reverse: boolean): void {
    const buf = this.buffer;
    this.lookupMask = mask;
    this.lookupProps = props;
    if (reverse) {
      for (let i = buf.length - 1; i >= 0; i--) {
        const info = buf[i]!;
        if (info.mask & mask && this.checkGlyphProperty(info, props)) {
          this.dispatch(this, lookupIndex, i);
          this.lookupProps = props;
        }
      }
      return;
    }
    // Only successful applications can fail to advance, so only they count against the
    // budget that stops a looping or exploding font.
    const maxLength = Math.max(256, buf.length * 32);
    let budget = Math.max(1024, buf.length * 16);
    let i = 0;
    while (i < buf.length) {
      const info = buf[i]!;
      if (info.mask & mask && this.checkGlyphProperty(info, props)) {
        const next = this.dispatch(this, lookupIndex, i);
        this.lookupProps = props;
        if (next >= 0) {
          if (--budget <= 0 || buf.length > maxLength) return;
          i = next;
          continue;
        }
      }
      i++;
    }
  }
}
