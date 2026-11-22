/**
 * Script shapers: what features a script needs beyond the common set, in which stages, and
 * the per-script buffer work (joining analysis, syllable reordering). The default shaper
 * (Latin, Greek, Cyrillic, Hebrew, CJK, Thai and everything without a dedicated shaper)
 * adds nothing and relies on the common features.
 */
import type { PlanBuilder, ShapeContext } from './plan.js';

export interface Shaper {
  name: string;
  /** Adds the shaper's features and pauses to the plan, before the common features. */
  collectFeatures?(b: PlanBuilder, script: string): void;
  /** Adjusts features after the common ones were added (e.g. Indic turns 'liga' off). */
  overrideFeatures?(b: PlanBuilder): void;
  /** Sets per-glyph masks / categories once glyphs are mapped, before GSUB. */
  setupMasks?(ctx: ShapeContext): void;
  /** Zero the advance of GDEF mark glyphs after GPOS (marks then sit on their base). */
  zeroMarks: boolean;
  /** Recompose base + combining mark into a precomposed character the font has. */
  compose: boolean;
  /** Characters to decompose even when the font has them (Indic two-part matras). */
  alwaysDecompose?(cp: number): boolean;
}

export const defaultShaper: Shaper = {
  name: 'default',
  zeroMarks: true,
  compose: true,
};
