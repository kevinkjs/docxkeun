/**
 * The shape plan: which GSUB/GPOS lookups run, in which order, on which glyphs. Shapers
 * request features in "stages" separated by pauses (where they reorder or analyse the
 * buffer); within a stage the lookups of all its features are merged and applied in lookup
 * list order, as fonts expect. Each feature gets a mask bit: global features share one bit,
 * per-glyph features (Arabic joining forms, Indic half/below forms...) get their own so the
 * shaper can enable them glyph by glyph.
 */
import type { GlyphInfo } from './buffer.js';
import type { Gpos } from './gpos.js';
import type { Gsub } from './gsub.js';
import type { LayoutTable, ScriptSelection } from './layout-common.js';

/** The font surface shaping needs; `Font` from sfnt.ts satisfies it. */
export interface ShapingFont {
  glyphId(codePoint: number): number;
  advance(glyph: number): number;
  tableBytes(tag: string): Uint8Array | undefined;
  kerning(left: number, right: number): number;
  unitsPerEm?: number;
}

export interface FeatureFlags {
  /** On for every glyph (shares the global mask bit). */
  global?: boolean;
  /** ZWJ is not skipped while matching (joiners steer the feature). */
  manualZwj?: boolean;
  manualZwnj?: boolean;
  /** Context matching stops at syllable boundaries (Indic). */
  perSyllable?: boolean;
  /** Keep a mask bit even when the font lacks the feature (the shaper has a fallback). */
  fallback?: boolean;
}

export interface ShapeContext {
  font: ShapingFont;
  buffer: GlyphInfo[];
  plan: ShapePlan;
  direction: 'ltr' | 'rtl';
}

export type PauseHook = (ctx: ShapeContext) => void;

export interface PlannedLookup {
  index: number;
  mask: number;
  autoZwj: boolean;
  autoZwnj: boolean;
  perSyllable: boolean;
}

export interface PlanStage {
  lookups: PlannedLookup[];
  /** Runs after this stage's lookups. */
  pause: PauseHook | undefined;
}

export const GLOBAL_MASK = 1;

interface FeatureEntry {
  tag: string;
  flags: FeatureFlags;
  stage: [number, number];
}

export class PlanBuilder {
  private readonly entries: FeatureEntry[] = [];
  private readonly stage: [number, number] = [0, 0];
  private readonly pauses: [Array<{ stage: number; hook: PauseHook | undefined }>, Array<{ stage: number; hook: PauseHook | undefined }>] = [[], []];

  /** Adds a feature to the current stage of both tables (it runs wherever the font has it). */
  add(tag: string, flags: FeatureFlags = {}): void {
    this.entries.push({ tag, flags, stage: [this.stage[0], this.stage[1]] });
  }

  enable(tag: string, flags: FeatureFlags = {}): void {
    this.add(tag, { ...flags, global: true });
  }

  has(tag: string): boolean {
    return this.entries.some((e) => e.tag === tag);
  }

  remove(tag: string): void {
    for (let i = this.entries.length - 1; i >= 0; i--) if (this.entries[i]!.tag === tag) this.entries.splice(i, 1);
  }

  pauseGsub(hook?: PauseHook): void {
    this.pauses[0].push({ stage: this.stage[0], hook });
    this.stage[0]++;
  }

  compile(gsub: Gsub | undefined, gpos: Gpos | undefined, gsubSel: ScriptSelection | undefined, gposSel: ScriptSelection | undefined): ShapePlan {
    // Merge repeated requests: the first one's joiner/syllable flags win, the last decides
    // whether it is global, and it runs in the earliest stage asked for.
    const merged = new Map<string, FeatureEntry>();
    for (const e of this.entries) {
      const prev = merged.get(e.tag);
      if (!prev) {
        merged.set(e.tag, { tag: e.tag, flags: { ...e.flags }, stage: [e.stage[0], e.stage[1]] });
        continue;
      }
      prev.flags.global = e.flags.global;
      prev.flags.fallback = prev.flags.fallback || e.flags.fallback;
      prev.stage[0] = Math.min(prev.stage[0], e.stage[0]);
      prev.stage[1] = Math.min(prev.stage[1], e.stage[1]);
    }

    const tables: [LayoutTable | undefined, LayoutTable | undefined] = [gsub?.layout, gpos?.layout];
    const sels: [ScriptSelection | undefined, ScriptSelection | undefined] = [gsubSel, gposSel];
    const plan = new ShapePlan(gsub, gpos, gsubSel?.scriptTag, gposSel?.scriptTag);
    let nextBit = 1;
    const active: Array<{ entry: FeatureEntry; mask: number; index: [number, number] }> = [];
    for (const entry of merged.values()) {
      const index: [number, number] = [-1, -1];
      for (const t of [0, 1] as const) {
        const table = tables[t];
        const sel = sels[t];
        if (table && sel) index[t] = table.findFeature(sel.langSys, entry.tag);
      }
      if (index[0] < 0 && index[1] < 0 && !entry.flags.fallback) continue;
      let mask: number;
      if (entry.flags.global) mask = GLOBAL_MASK;
      else {
        if (nextBit > 30) continue; // out of mask bits: drop the feature rather than alias another
        mask = 1 << nextBit++;
      }
      plan.masks.set(entry.tag, mask);
      plan.featureIndex[0].set(entry.tag, index[0]);
      plan.featureIndex[1].set(entry.tag, index[1]);
      active.push({ entry, mask, index });
    }

    for (const t of [0, 1] as const) {
      const table = tables[t];
      const sel = sels[t];
      const pauses = this.pauses[t];
      const stageCount = pauses.length + 1;
      const stages: PlanStage[] = [];
      for (let s = 0; s < stageCount; s++) {
        const lookups: PlannedLookup[] = [];
        if (table && sel) {
          if (s === 0 && sel.langSys.required >= 0) {
            for (const li of table.featureLookups(sel.langSys.required)) lookups.push({ index: li, mask: GLOBAL_MASK, autoZwj: true, autoZwnj: true, perSyllable: false });
          }
          for (const { entry, mask, index } of active) {
            if (entry.stage[t] !== s || index[t] < 0) continue;
            for (const li of table.featureLookups(index[t])) {
              lookups.push({
                index: li,
                mask,
                autoZwj: !entry.flags.manualZwj,
                autoZwnj: !entry.flags.manualZwnj,
                perSyllable: !!entry.flags.perSyllable,
              });
            }
          }
        }
        lookups.sort((a, b) => a.index - b.index);
        const deduped: PlannedLookup[] = [];
        for (const l of lookups) {
          const last = deduped[deduped.length - 1];
          if (last && last.index === l.index) {
            last.mask |= l.mask;
            last.autoZwj &&= l.autoZwj;
            last.autoZwnj &&= l.autoZwnj;
            last.perSyllable ||= l.perSyllable;
          } else deduped.push({ ...l });
        }
        stages.push({ lookups: deduped, pause: pauses.find((p) => p.stage === s)?.hook });
      }
      if (t === 0) plan.gsubStages = stages;
      else plan.gposStages = stages;
    }
    return plan;
  }
}

export class ShapePlan {
  gsubStages: PlanStage[] = [];
  gposStages: PlanStage[] = [];
  readonly masks = new Map<string, number>();
  readonly featureIndex: [Map<string, number>, Map<string, number>] = [new Map(), new Map()];
  /** Shaper-specific data computed once per plan. */
  shaperData: unknown;
  /** Unicode-script tag being shaped ('deva', 'arab', 'DFLT'...). */
  script = 'DFLT';

  constructor(
    readonly gsub: Gsub | undefined,
    readonly gpos: Gpos | undefined,
    /** Script tag GSUB/GPOS actually selected ('dev2', 'DFLT', ...), undefined when absent. */
    readonly gsubScript: string | undefined,
    readonly gposScript: string | undefined,
  ) {}

  /** Mask bit of a feature, 0 when the plan doesn't apply it. */
  maskOf(tag: string): number {
    return this.masks.get(tag) ?? 0;
  }

  /** Whether the font itself implements the feature (as opposed to a shaper fallback). */
  hasFeature(table: 0 | 1, tag: string): boolean {
    return (this.featureIndex[table].get(tag) ?? -1) >= 0;
  }

  /** GSUB lookups of a feature (for "would this substitute?" queries). */
  gsubLookups(tag: string): number[] {
    const fi = this.featureIndex[0].get(tag) ?? -1;
    return fi >= 0 && this.gsub ? this.gsub.layout.featureLookups(fi) : [];
  }
}
