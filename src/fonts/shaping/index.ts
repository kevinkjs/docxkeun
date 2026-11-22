/**
 * OpenType text shaping: turns a run of text in one font, script and direction into
 * positioned glyphs using the font's GSUB/GPOS tables (ligatures, contextual forms, Arabic
 * joining, Indic reordering, kerning, mark attachment), or cmap + hmtx + legacy 'kern' for
 * fonts without OpenType layout.
 *
 * The pipeline follows HarfBuzz's hb-ot-shape: map characters to glyphs (with light
 * normalization and RTL mirroring), choose a script shaper, build a plan of feature stages,
 * run GSUB, hide default-ignorables, run GPOS, zero mark advances, resolve attachments.
 *
 * Output is in logical order. For right-to-left runs, offsets assume the caller draws the
 * glyphs reversed (visual order), advancing left to right, which is how HarfBuzz computes
 * them before its final reversal.
 */
import { arabicShaper } from './arabic.js';
import { ApplyContext } from './apply.js';
import { deleteGlyph, isMark, newGlyphInfo, type GlyphInfo } from './buffer.js';
import { Gdef, GlyphProps } from './gdef.js';
import { Gpos, propagateAttachments } from './gpos.js';
import { Gsub } from './gsub.js';
import { indicShaper } from './indic.js';
import { normalizeTag } from './layout-common.js';
import { GLOBAL_MASK, PlanBuilder, type ShapeContext, type ShapePlan, type ShapingFont } from './plan.js';
import { canonicalScript, detectScript, INDIC_NEW_TAGS, scriptTagCandidates } from './scripts.js';
import { defaultShaper, type Shaper } from './shaper.js';
import { isExtendedPictographic, markGoesBefore, mirrorOf, UFlag, unicodeFlags } from './unicode.js';

export { scriptOf, itemizeScripts, type ScriptRun } from './scripts.js';
export type { ShapingFont } from './plan.js';

export interface ShapedGlyph {
  glyph: number;
  /** UTF-16 index of the first character of the glyph's cluster. */
  cluster: number;
  /** Advance width in font units (0 for attached marks). */
  advance: number;
  /** Offset from the pen position, in font units (y up). */
  offsetX: number;
  offsetY: number;
}

export interface ShapeOptions {
  direction: 'ltr' | 'rtl';
  /** OpenType script tag ('latn', 'arab', 'deva'...); detected from the text when absent. */
  script?: string;
  /** OpenType language system tag ('URD ', 'MAR '...). */
  language?: string;
  /**
   * Turn features on or off by tag: `{ liga: false }` disables standard ligatures (and
   * kern / calt / clig likewise); `{ smcp: true, onum: true }` enables optional ones.
   */
  features?: Record<string, boolean>;
}

/** Whether the font has OpenType layout tables (otherwise shaping is cmap + kern only). */
export function hasOpenTypeLayout(font: ShapingFont): boolean {
  return !!(font.tableBytes('GSUB')?.length || font.tableBytes('GPOS')?.length);
}

interface FontLayout {
  gsub: Gsub | undefined;
  gpos: Gpos | undefined;
  gdef: Gdef;
  plans: Map<string, { plan: ShapePlan; shaper: Shaper }>;
}

const layouts = new WeakMap<object, FontLayout>();

function fontLayout(font: ShapingFont): FontLayout {
  let layout = layouts.get(font);
  if (!layout) {
    const gsub = font.tableBytes('GSUB');
    const gpos = font.tableBytes('GPOS');
    layout = {
      gsub: gsub && gsub.length >= 10 ? new Gsub(gsub) : undefined,
      gpos: gpos && gpos.length >= 10 ? new Gpos(gpos) : undefined,
      gdef: new Gdef(font.tableBytes('GDEF')),
      plans: new Map(),
    };
    layouts.set(font, layout);
  }
  return layout;
}

function chooseShaper(script: string, gsubScript: string | undefined): Shaper {
  if (INDIC_NEW_TAGS[script]) {
    // A font that only covers the script through 'DFLT'/'latn' was designed for no reordering.
    return gsubScript === 'DFLT' || gsubScript === 'latn' ? defaultShaper : indicShaper;
  }
  if (script === 'arab') return arabicShaper;
  if (script === 'syrc' || script === 'nko ' || script === 'mong') return gsubScript === 'DFLT' ? defaultShaper : arabicShaper;
  return defaultShaper;
}

function planFor(font: ShapingFont, layout: FontLayout, script: string, language: string | undefined, direction: 'ltr' | 'rtl', features: Record<string, boolean> | undefined) {
  const featureKey = features
    ? Object.entries(features)
        .map(([t, on]) => `${normalizeTag(t)}=${on ? 1 : 0}`)
        .sort()
        .join(',')
    : '';
  const key = `${script}|${language ?? ''}|${direction}|${featureKey}`;
  const cached = layout.plans.get(key);
  if (cached) return cached;

  const candidates = scriptTagCandidates(script);
  const gsubSel = layout.gsub?.layout.select(candidates, language);
  const gposSel = layout.gpos?.layout.select(candidates, language);
  const shaper = chooseShaper(script, gsubSel?.scriptTag);

  const b = new PlanBuilder();
  b.enable('rvrn');
  b.pauseGsub();
  if (direction === 'ltr') {
    b.enable('ltra');
    b.enable('ltrm');
  } else {
    b.enable('rtla');
    b.add('rtlm'); // only on characters without a mirrored counterpart in the cmap
  }
  shaper.collectFeatures?.(b, script);
  for (const tag of ['abvm', 'blwm', 'ccmp', 'locl']) b.enable(tag);
  b.enable('mark', { manualZwj: true, manualZwnj: true });
  b.enable('mkmk', { manualZwj: true, manualZwnj: true });
  b.enable('rlig');
  for (const tag of ['calt', 'clig', 'curs', 'dist']) b.enable(tag);
  b.enable('kern', { fallback: true }); // the legacy 'kern' table stands in when GPOS has none
  b.enable('liga');
  b.enable('rclt');
  shaper.overrideFeatures?.(b);
  for (const [tag, on] of Object.entries(features ?? {})) {
    const t = normalizeTag(tag);
    if (!on) b.remove(t);
    else if (!b.has(t)) b.enable(t);
  }
  const plan = b.compile(layout.gsub, layout.gpos, gsubSel, gposSel);
  plan.script = script;
  const entry = { plan, shaper };
  layout.plans.set(key, entry);
  return entry;
}

/**
 * Shapes `text` (one script, one direction, one font) into glyphs in logical order, with
 * advances and offsets in font units.
 */
export function shape(font: ShapingFont, text: string, options: ShapeOptions): ShapedGlyph[] {
  if (!text) return [];
  const direction = options.direction === 'rtl' ? 'rtl' : 'ltr';
  let script = options.script ? normalizeTag(options.script) : detectScript(text);
  script = script === 'DFLT' || script === 'dflt' || script === 'Zyyy' ? 'DFLT' : canonicalScript(script.toLowerCase());
  const language = options.language ? normalizeTag(options.language) : undefined;
  try {
    return shapeOpenType(font, text, script, language, direction, options.features);
  } catch {
    // A font whose layout tables trip the engine still gets readable text.
    return shapeSimple(font, text);
  }
}

function shapeOpenType(
  font: ShapingFont,
  text: string,
  script: string,
  language: string | undefined,
  direction: 'ltr' | 'rtl',
  features: Record<string, boolean> | undefined,
): ShapedGlyph[] {
  const layout = fontLayout(font);
  const { plan, shaper } = planFor(font, layout, script, language, direction, features);
  const buffer = buildBuffer(font, text, direction, shaper, plan);
  const ctx: ShapeContext = { font, buffer, plan, direction };
  shaper.setupMasks?.(ctx);

  // Glyph classes: from GDEF, else non-spacing marks are marks and the rest bases.
  const gdef = layout.gdef;
  for (const info of buffer) {
    if (gdef.hasGlyphClasses) info.props = gdef.glyphProps(info.glyph);
    else info.props = info.uflags & UFlag.NonSpacingMark && !(info.uflags & UFlag.DefaultIgnorable) ? GlyphProps.Mark : GlyphProps.BaseGlyph;
  }

  // GSUB, stage by stage, with the shaper's pauses in between.
  const gsub = layout.gsub;
  const gsubCtx = new ApplyContext(buffer, gdef, 0, direction, gsub ? gsub.dispatch : () => -1);
  for (const stage of plan.gsubStages) {
    if (gsub) {
      for (const l of stage.lookups) {
        const lookup = gsub.lookup(l.index);
        if (!lookup) continue;
        gsubCtx.autoZwj = l.autoZwj;
        gsubCtx.autoZwnj = l.autoZwnj;
        gsubCtx.perSyllable = l.perSyllable;
        gsubCtx.applyLookup(l.index, lookup.props, l.mask, lookup.reverse);
      }
    }
    stage.pause?.(ctx);
  }

  hideDefaultIgnorables(font, buffer);
  position(font, layout, plan, shaper, buffer, direction);

  return buffer.map((g) => ({ glyph: g.glyph, cluster: g.cluster, advance: g.xAdvance, offsetX: g.xOffset, offsetY: g.yOffset }));
}

/** Space-like characters a font may lack, drawn as its space with a standard width. */
function spaceFallbackWidth(font: ShapingFont, cp: number): number | undefined {
  const em = font.unitsPerEm ?? 1000;
  const space = font.advance(font.glyphId(0x20));
  switch (cp) {
    case 0x00a0:
      return space;
    case 0x2000:
    case 0x2002:
      return em / 2;
    case 0x2001:
    case 0x2003:
    case 0x3000:
      return em;
    case 0x2004:
      return em / 3;
    case 0x2005:
      return em / 4;
    case 0x2006:
      return em / 6;
    case 0x2007:
      return font.advance(font.glyphId(0x30));
    case 0x2008:
      return font.advance(font.glyphId(0x2e));
    case 0x2009:
      return em / 5;
    case 0x200a:
      return em / 16;
    case 0x202f:
      return space / 2;
    case 0x205f:
      return (em * 4) / 18;
    default:
      return undefined;
  }
}

/** Fixed advances for glyphs substituted by the space fallback, keyed by buffer record. */
const advanceOverrides = new WeakMap<GlyphInfo, number>();

/**
 * Characters → glyph records: RTL mirroring, canonical decomposition when the font lacks a
 * precomposed character (or the shaper always splits it), recomposition of base + mark
 * when the font has the precomposed glyph, then grapheme clusters.
 */
function buildBuffer(font: ShapingFont, text: string, direction: 'ltr' | 'rtl', shaper: Shaper, plan: ShapePlan): GlyphInfo[] {
  const buf: GlyphInfo[] = [];
  const rtlm = plan.maskOf('rtlm');
  for (let i = 0; i < text.length; ) {
    let cp = text.codePointAt(i)!;
    const cluster = i;
    i += cp > 0xffff ? 2 : 1;
    let needsRtlm = false;
    if (direction === 'rtl') {
      const mirrored = mirrorOf(cp);
      if (mirrored !== undefined) {
        if (font.glyphId(mirrored)) cp = mirrored;
        else needsRtlm = true;
      }
    }
    for (const part of decompose(font, cp, shaper)) {
      const info = newGlyphInfo(part, cluster, unicodeFlags(part));
      info.mask = GLOBAL_MASK | (needsRtlm ? rtlm : 0);
      buf.push(info);
    }
  }

  reorderMarks(buf);
  if (shaper.compose) compose(font, buf);

  // Grapheme clusters: marks, joiners and emoji continuations join the previous cluster.
  for (let i = 1; i < buf.length; i++) {
    const info = buf[i]!;
    const prev = buf[i - 1]!;
    const cp = info.codepoint;
    const isRegional = (c: number) => c >= 0x1f1e6 && c <= 0x1f1ff;
    if (
      info.uflags & (UFlag.Continuation | UFlag.Zwj) ||
      (prev.uflags & UFlag.Zwj && isExtendedPictographic(cp)) ||
      (cp >= 0xff9e && cp <= 0xff9f) ||
      (isRegional(cp) && isRegional(prev.codepoint) && (i < 2 || buf[i - 2]!.cluster !== prev.cluster))
    ) {
      info.cluster = prev.cluster;
    }
  }

  for (const info of buf) {
    let glyph = font.glyphId(info.codepoint);
    if (!glyph) {
      const width = spaceFallbackWidth(font, info.codepoint);
      if (width !== undefined && font.glyphId(0x20)) {
        glyph = font.glyphId(0x20);
        advanceOverrides.set(info, Math.round(width));
      } else if (info.codepoint === 0x2011) {
        glyph = font.glyphId(0x2010) || font.glyphId(0x2d);
      }
    }
    info.glyph = glyph;
  }
  return buf;
}

function decompose(font: ShapingFont, cp: number, shaper: Shaper): number[] {
  if (cp < 0xc0 || (cp >= 0xac00 && cp <= 0xd7a3)) return [cp];
  const always = shaper.alwaysDecompose?.(cp) ?? false;
  if (!always && font.glyphId(cp)) return [cp];
  const parts = [...String.fromCodePoint(cp).normalize('NFD')].map((ch) => ch.codePointAt(0)!);
  if (parts.length === 1 && parts[0] === cp) return [cp];
  if (!parts.every((p) => font.glyphId(p))) return [cp];
  return parts;
}

/** Canonical (combining-class) order within each run of marks, like normalization does. */
function reorderMarks(buf: GlyphInfo[]): void {
  const markFlags = UFlag.NonSpacingMark | UFlag.SpacingMark | UFlag.EnclosingMark;
  for (let i = 0; i < buf.length; ) {
    let j = i;
    while (j < buf.length && buf[j]!.uflags & markFlags) j++;
    if (j - i >= 2) {
      // Stable bubble sort: runs are a handful of marks, and only adjacent swaps are canonical.
      // At most n passes, so a comparator mixing modified and standard classes can't cycle.
      for (let swapped = true, pass = 0; swapped && pass < j - i; pass++) {
        swapped = false;
        for (let k = i; k + 1 < j; k++) {
          if (markGoesBefore(buf[k + 1]!.codepoint, buf[k]!.codepoint)) {
            [buf[k], buf[k + 1]] = [buf[k + 1]!, buf[k]!];
            swapped = true;
          }
        }
      }
    }
    i = Math.max(j, i + 1);
  }
}

function compose(font: ShapingFont, buf: GlyphInfo[]): void {
  const markFlags = UFlag.NonSpacingMark | UFlag.SpacingMark | UFlag.EnclosingMark;
  for (let i = 1; i < buf.length; i++) {
    const mark = buf[i]!;
    const base = buf[i - 1]!;
    if (!(mark.uflags & markFlags) || base.uflags & markFlags) continue;
    const composed = [...(String.fromCodePoint(base.codepoint) + String.fromCodePoint(mark.codepoint)).normalize('NFC')];
    if (composed.length !== 1) continue;
    const cp = composed[0]!.codePointAt(0)!;
    if (!font.glyphId(cp)) continue;
    base.codepoint = cp;
    base.uflags = unicodeFlags(cp);
    buf.splice(i, 1);
    i--;
  }
}

/** ZWJ, ZWNJ, bidi controls etc. become the (zero-width) space glyph after GSUB. */
function hideDefaultIgnorables(font: ShapingFont, buf: GlyphInfo[]): void {
  const space = font.glyphId(0x20);
  for (let i = 0; i < buf.length; i++) {
    const info = buf[i]!;
    if (!(info.uflags & UFlag.DefaultIgnorable) || info.props & GlyphProps.Substituted) continue;
    if (space) info.glyph = space;
    else {
      deleteGlyph(buf, i);
      i--;
    }
  }
}

function position(font: ShapingFont, layout: FontLayout, plan: ShapePlan, shaper: Shaper, buf: GlyphInfo[], direction: 'ltr' | 'rtl'): void {
  for (const info of buf) {
    info.xAdvance = advanceOverrides.get(info) ?? font.advance(info.glyph);
    info.yAdvance = 0;
    info.xOffset = 0;
    info.yOffset = 0;
    info.attachType = 0;
    info.attachChain = 0;
  }

  const gpos = layout.gpos;
  if (gpos) {
    const ctx = new ApplyContext(buf, layout.gdef, 1, direction, gpos.dispatch);
    for (const stage of plan.gposStages) {
      for (const l of stage.lookups) {
        const lookup = gpos.lookup(l.index);
        if (lookup) ctx.applyLookup(l.index, lookup.props, l.mask, false);
      }
    }
  }

  const kernMask = plan.maskOf('kern');
  if (kernMask && !plan.hasFeature(1, 'kern')) legacyKern(font, buf, kernMask);

  if (shaper.zeroMarks) {
    // Without GPOS, pull a zeroed mark back over its base (LTR) instead of after it.
    const adjust = !gpos && direction === 'ltr';
    for (const info of buf) {
      if (!isMark(info)) continue;
      if (adjust) info.xOffset -= info.xAdvance;
      info.xAdvance = 0;
    }
  }

  for (const info of buf) {
    if (info.uflags & UFlag.DefaultIgnorable && !(info.props & GlyphProps.Substituted)) {
      info.xAdvance = 0;
      info.xOffset = 0;
      info.yOffset = 0;
    }
  }

  propagateAttachments(buf, direction);
}

/** Pair kerning from the legacy 'kern' table, between consecutive non-mark glyphs. */
function legacyKern(font: ShapingFont, buf: GlyphInfo[], mask: number): void {
  for (let i = 0; i < buf.length; i++) {
    const a = buf[i]!;
    if (!(a.mask & mask) || isMark(a)) continue;
    let j = i + 1;
    while (j < buf.length && (isMark(buf[j]!) || buf[j]!.uflags & UFlag.DefaultIgnorable)) j++;
    const b = buf[j];
    if (!b || !(b.mask & mask)) continue;
    a.xAdvance += font.kerning(a.glyph, b.glyph);
  }
}

/** cmap + hmtx + 'kern', one glyph per character (the safety net). */
function shapeSimple(font: ShapingFont, text: string): ShapedGlyph[] {
  const out: ShapedGlyph[] = [];
  for (let i = 0; i < text.length; ) {
    const cp = text.codePointAt(i)!;
    const glyph = font.glyphId(cp);
    out.push({ glyph, cluster: i, advance: font.advance(glyph), offsetX: 0, offsetY: 0 });
    i += cp > 0xffff ? 2 : 1;
  }
  for (let k = 0; k + 1 < out.length; k++) out[k]!.advance += font.kerning(out[k]!.glyph, out[k + 1]!.glyph);
  return out;
}
