/**
 * The Unicode Bidirectional Algorithm (UAX #9): resolves an embedding level for every character of
 * a paragraph and the visual order of a line, so right-to-left scripts (Hebrew, Arabic, ...) and
 * mixed-direction text display in the right order.
 *
 * Implemented: P2–P3 (paragraph direction, skipping isolates), X1–X10 (explicit embeddings,
 * overrides and isolates, depth limit 125 with overflow counters, isolating run sequences), W1–W7,
 * N0 (paired brackets), N1–N2, I1–I2, L1 and L2. L3 (combining marks) and L4 (mirroring, via
 * mirrorChar) belong to the renderer.
 *
 * Characters removed by X9 (LRE, RLE, LRO, RLO, PDF and BN) stay in place: they are skipped when
 * building level runs and afterwards take the level of the preceding character, as the UAX #9
 * reference implementation does, so callers get a level for every code unit and no level break
 * appears where the controls sit.
 */
import { BC, BIDI_CLASSES, BPT_CLOSE, BPT_OPEN, bidiClassCode, mirrorGlyph, pairedBracket, pairedBracketType } from './bidi-data.js';
import type { BidiClass } from './bidi-data.js';

export type { BidiClass } from './bidi-data.js';

const { L, R, AL, EN, ES, ET, AN, CS, NSM, BN, B, S, WS, ON, LRE, LRO, RLE, RLO, PDF, LRI, RLI, FSI, PDI } = BC;

/** BD2: the deepest embedding level explicit controls may reach. */
const MAX_DEPTH = 125;
/** BD16: bracket pairing gives up beyond this many open brackets in one sequence. */
const MAX_BRACKET_STACK = 63;

export interface BidiParagraph {
  /** Resolved embedding level of every UTF-16 code unit (both halves of a surrogate pair agree). */
  levels: Uint8Array;
  /** The paragraph embedding level used (that of the first paragraph if the text has several). */
  paragraphLevel: 0 | 1;
}

export function bidiClass(cp: number): BidiClass {
  return BIDI_CLASSES[bidiClassCode(cp)]!;
}

/** True for strong right-to-left characters (classes R and AL). */
export function isRtlChar(cp: number): boolean {
  const c = bidiClassCode(cp);
  return c === R || c === AL;
}

/**
 * Bidi_Mirroring_Glyph, for rule L4: a character resolved to an odd level is drawn as its mirror,
 * so "(" in right-to-left text still opens towards its content.
 */
export function mirrorChar(cp: number): number | undefined {
  return mirrorGlyph(cp);
}

const isRemovedByX9 = (t: number) => t === BN || (t >= LRE && t <= PDF);
const isIsolateInitiator = (t: number) => t >= LRI && t <= FSI;
/** Types N1/N2 resolve; isolate controls count as neutrals there. */
const isNeutral = (t: number) => t === B || t === S || t === WS || t === ON || (t >= LRI && t <= PDI);
/** L1 resets these when they trail a line or precede a separator (X9-removed ones included). */
const isL1Whitespace = (t: number) => t === WS || t === BN || (t >= LRE && t <= PDI);

/**
 * Resolves embedding levels (rules P2–L1) for text that may hold several paragraphs: each
 * paragraph separator (class B, with CR LF counted as one) ends a paragraph, and with 'auto' each
 * paragraph finds its own direction. Line-end whitespace (the last clause of L1) is applied to the
 * end of each paragraph; use lineLevels() for the other line ends once lines are broken.
 */
export function resolveLevels(text: string, paragraphLevel: 0 | 1 | 'auto'): BidiParagraph {
  const cps: number[] = [];
  for (let i = 0; i < text.length; i++) {
    const cp = text.codePointAt(i)!;
    cps.push(cp);
    if (cp > 0xffff) i++;
  }
  const classes = new Uint8Array(cps.length);
  for (let i = 0; i < cps.length; i++) classes[i] = bidiClassCode(cps[i]!);

  const cpLevels = new Uint8Array(cps.length);
  let firstLevel: 0 | 1 | undefined;
  for (let start = 0; start < cps.length; ) {
    let end = start;
    while (end < cps.length && classes[end] !== B) end++;
    if (end < cps.length) {
      end++;
      if (cps[end - 1] === 0x0d && cps[end] === 0x0a) end++;
    }
    const paragraph = resolveParagraph(cps.slice(start, end), classes.subarray(start, end), paragraphLevel);
    cpLevels.set(paragraph.levels, start);
    firstLevel ??= paragraph.level;
    start = end;
  }

  const levels = new Uint8Array(text.length);
  for (let i = 0, unit = 0; i < cps.length; i++) {
    levels[unit++] = cpLevels[i]!;
    if (cps[i]! > 0xffff) levels[unit++] = cpLevels[i]!;
  }
  return { levels, paragraphLevel: firstLevel ?? (paragraphLevel === 'auto' ? 0 : paragraphLevel) };
}

/**
 * Rule L1 for a line break: a copy of `levels` in which the whitespace, isolate controls and
 * X9-removed characters ending the line [start, end) are reset to the paragraph level. A wrapped
 * line's trailing spaces then sit at its visual end instead of between right-to-left words.
 */
export function lineLevels(text: string, levels: Uint8Array, paragraphLevel: 0 | 1, start: number, end: number): Uint8Array {
  const out = levels.slice();
  for (let i = end; i > start; ) {
    let j = i - 1;
    const unit = text.charCodeAt(j);
    if (unit >= 0xdc00 && unit <= 0xdfff && j > start) {
      const high = text.charCodeAt(j - 1);
      if (high >= 0xd800 && high <= 0xdbff) j--;
    }
    if (!isL1Whitespace(bidiClassCode(text.codePointAt(j)!))) break;
    out.fill(paragraphLevel, j, i);
    i = j;
  }
  return out;
}

/**
 * Rule L2 for one line: the logical indices start..end-1 in visual (left-to-right) order. From the
 * highest level down to the lowest odd level, every maximal run at that level or higher is
 * reversed. Indices are whatever `levels` is indexed by; with UTF-16 levels a surrogate pair at an
 * odd level comes out low half first, so callers that place single characters should keep pairs
 * (and combining sequences, rule L3) together themselves.
 */
export function visualOrder(levels: ArrayLike<number>, start: number, end: number): number[] {
  const order: number[] = [];
  let highest = 0;
  let lowestOdd = Infinity;
  for (let i = start; i < end; i++) {
    order.push(i);
    const level = levels[i]!;
    if (level > highest) highest = level;
    if (level & 1 && level < lowestOdd) lowestOdd = level;
  }
  for (let level = highest; level >= lowestOdd; level--) {
    for (let k = 0; k < order.length; ) {
      if (levels[order[k]!]! < level) {
        k++;
        continue;
      }
      let runEnd = k;
      while (runEnd < order.length && levels[order[runEnd]!]! >= level) runEnd++;
      for (let a = k, b = runEnd - 1; a < b; a++, b--) {
        const t = order[a]!;
        order[a] = order[b]!;
        order[b] = t;
      }
      k = runEnd;
    }
  }
  return order;
}

/** P2: the first strong class (L, R or AL) in [from, to), skipping isolates; -1 if none. */
function firstStrong(types: Uint8Array, from: number, to: number, matchingPDI: Int32Array): number {
  for (let i = from; i < to; i++) {
    const t = types[i]!;
    if (t === L || t === R || t === AL) return t;
    // Jump to the matching PDI (or the paragraph end), which the loop then steps past.
    if (isIsolateInitiator(t)) i = matchingPDI[i]!;
  }
  return -1;
}

/** One paragraph (P2 to L1). `t0` holds the original classes; the last character may be a B. */
function resolveParagraph(cps: readonly number[], t0: Uint8Array, requested: 0 | 1 | 'auto'): { levels: Uint8Array; level: 0 | 1 } {
  const n = t0.length;

  // BD9: pair isolate initiators with PDIs. Unmatched initiators "match" the paragraph end (n).
  const matchingPDI = new Int32Array(n).fill(-1);
  const matchingInitiator = new Int32Array(n).fill(-1);
  const openIsolates: number[] = [];
  for (let i = 0; i < n; i++) {
    const t = t0[i]!;
    if (isIsolateInitiator(t)) openIsolates.push(i);
    else if (t === PDI && openIsolates.length > 0) {
      const initiator = openIsolates.pop()!;
      matchingPDI[initiator] = i;
      matchingInitiator[i] = initiator;
    }
  }
  for (const i of openIsolates) matchingPDI[i] = n;

  // P2, P3
  let paraLevel: 0 | 1;
  if (requested === 'auto') {
    const strong = firstStrong(t0, 0, n, matchingPDI);
    paraLevel = strong === R || strong === AL ? 1 : 0;
  } else paraLevel = requested;

  // X1–X8: explicit levels. `types` receives override results; `embedding` the explicit levels.
  const types = Uint8Array.from(t0);
  const embedding = new Uint8Array(n);
  const stackLevel = new Uint8Array(MAX_DEPTH + 2);
  const stackOverride = new Uint8Array(MAX_DEPTH + 2);
  const stackIsolate = new Uint8Array(MAX_DEPTH + 2);
  stackLevel[0] = paraLevel;
  stackOverride[0] = ON;
  let depth = 1;
  let overflowIsolates = 0;
  let overflowEmbeddings = 0;
  let validIsolates = 0;
  for (let i = 0; i < n; i++) {
    const t = t0[i]!;
    const top = depth - 1;
    if ((t >= LRE && t <= RLO) || isIsolateInitiator(t)) {
      // X2–X5c
      const isolate = isIsolateInitiator(t);
      let rtl = t === RLE || t === RLO || t === RLI;
      if (t === FSI) {
        // X5c: an FSI is an RLI or LRI depending on its content's first strong character.
        const strong = firstStrong(t0, i + 1, matchingPDI[i]!, matchingPDI);
        rtl = strong === R || strong === AL;
      }
      if (isolate) {
        // An isolate initiator belongs to the outside: it takes the current level and override.
        embedding[i] = stackLevel[top]!;
        if (stackOverride[top] !== ON) types[i] = stackOverride[top]!;
      }
      const level = rtl ? (stackLevel[top]! + 1) | 1 : (stackLevel[top]! + 2) & ~1;
      if (level <= MAX_DEPTH && overflowIsolates === 0 && overflowEmbeddings === 0) {
        if (isolate) validIsolates++;
        stackLevel[depth] = level;
        stackOverride[depth] = t === LRO ? L : t === RLO ? R : ON;
        stackIsolate[depth] = isolate ? 1 : 0;
        depth++;
        if (!isolate) embedding[i] = level;
      } else if (isolate) overflowIsolates++;
      else if (overflowIsolates === 0) overflowEmbeddings++;
    } else if (t === PDI) {
      // X6a
      if (overflowIsolates > 0) overflowIsolates--;
      else if (validIsolates > 0) {
        overflowEmbeddings = 0;
        while (!stackIsolate[depth - 1]) depth--;
        depth--;
        validIsolates--;
      }
      embedding[i] = stackLevel[depth - 1]!;
      if (stackOverride[depth - 1] !== ON) types[i] = stackOverride[depth - 1]!;
    } else if (t === PDF) {
      // X7
      if (overflowIsolates > 0) {
        // An overflowed isolate swallows PDFs until its PDI.
      } else if (overflowEmbeddings > 0) overflowEmbeddings--;
      else if (!stackIsolate[top] && depth >= 2) depth--;
      embedding[i] = stackLevel[depth - 1]!;
    } else if (t === B) {
      // X8: B only ever ends the paragraph.
      embedding[i] = paraLevel;
    } else {
      // X6 (BN gets a level too, though X9 removes it)
      embedding[i] = stackLevel[top]!;
      if (t !== BN && stackOverride[top] !== ON) types[i] = stackOverride[top]!;
    }
  }

  // X9, X10 / BD7: level runs over the characters X9 keeps.
  const runs: number[][] = [];
  const runOf = new Int32Array(n).fill(-1);
  let run: number[] | undefined;
  for (let i = 0; i < n; i++) {
    if (isRemovedByX9(t0[i]!)) continue;
    if (!run || embedding[i] !== embedding[run[0]!]) runs.push((run = []));
    run.push(i);
    runOf[i] = runs.length - 1;
  }

  // BD13: chain level runs across matched isolates into isolating run sequences. sos and eos are
  // computed before any sequence is resolved since they compare explicit levels only.
  const levels = new Uint8Array(n);
  for (const levelRun of runs) {
    const first = levelRun[0]!;
    if (t0[first] === PDI && matchingInitiator[first] !== -1) continue; // continues an earlier sequence
    const sequence: number[] = [];
    for (let current = levelRun; ; ) {
      for (const i of current) sequence.push(i);
      const last = current[current.length - 1]!;
      if (!isIsolateInitiator(t0[last]!) || matchingPDI[last]! >= n) break;
      current = runs[runOf[matchingPDI[last]!]!]!;
    }
    resolveSequence(sequence, cps, t0, types, embedding, paraLevel, levels);
  }

  // Removed characters follow the character before them, so they never split a run.
  for (let i = 0; i < n; i++) {
    if (isRemovedByX9(t0[i]!)) levels[i] = i > 0 ? levels[i - 1]! : paraLevel;
  }

  // L1: separators, the whitespace before them, and the whitespace ending the paragraph.
  for (let i = 0; i < n; i++) {
    const t = t0[i]!;
    if (t !== B && t !== S) continue;
    levels[i] = paraLevel;
    for (let j = i - 1; j >= 0 && isL1Whitespace(t0[j]!); j--) levels[j] = paraLevel;
  }
  for (let j = n - 1; j >= 0 && isL1Whitespace(t0[j]!); j--) levels[j] = paraLevel;

  return { levels, level: paraLevel };
}

/** W1–W7, N0–N2 and I1–I2 for one isolating run sequence (paragraph indices in `seq`). */
function resolveSequence(
  seq: readonly number[],
  cps: readonly number[],
  t0: Uint8Array,
  types: Uint8Array,
  embedding: Uint8Array,
  paraLevel: number,
  levels: Uint8Array,
): void {
  const n = seq.length;
  const total = t0.length;
  const first = seq[0]!;
  const last = seq[n - 1]!;
  const level = embedding[first]!;

  // X10: sos/eos from the higher of this level and the neighbouring (non-removed) character's.
  let prev = first - 1;
  while (prev >= 0 && isRemovedByX9(t0[prev]!)) prev--;
  const sos = Math.max(prev >= 0 ? embedding[prev]! : paraLevel, level) & 1 ? R : L;
  let nextLevel = paraLevel;
  if (!isIsolateInitiator(t0[last]!)) {
    let next = last + 1;
    while (next < total && isRemovedByX9(t0[next]!)) next++;
    if (next < total) nextLevel = embedding[next]!;
  }
  const eos = Math.max(nextLevel, level) & 1 ? R : L;

  const ts = new Uint8Array(n);
  for (let k = 0; k < n; k++) ts[k] = types[seq[k]!]!;

  // W1: NSM takes the type of the character before it (ON after an isolate control).
  let before: number = sos;
  for (let k = 0; k < n; k++) {
    const t = ts[k]!;
    if (t === NSM) ts[k] = before;
    else before = t >= LRI && t <= PDI ? ON : t;
  }
  // W2: EN after AL is an Arabic number.
  let strong: number = sos;
  for (let k = 0; k < n; k++) {
    const t = ts[k]!;
    if (t === L || t === R || t === AL) strong = t;
    else if (t === EN && strong === AL) ts[k] = AN;
  }
  // W3
  for (let k = 0; k < n; k++) if (ts[k] === AL) ts[k] = R;
  // W4: a single separator between two numbers of the same kind joins them.
  for (let k = 1; k < n - 1; k++) {
    const t = ts[k]!;
    if (t !== ES && t !== CS) continue;
    const a = ts[k - 1]!;
    const b = ts[k + 1]!;
    if (a === EN && b === EN) ts[k] = EN;
    else if (t === CS && a === AN && b === AN) ts[k] = AN;
  }
  // W5: terminators next to a European number become part of it.
  for (let k = 0; k < n; k++) {
    if (ts[k] !== ET) continue;
    let end = k;
    while (end < n && ts[end] === ET) end++;
    if ((k === 0 ? sos : ts[k - 1]) === EN || (end === n ? eos : ts[end]) === EN) ts.fill(EN, k, end);
    k = end - 1;
  }
  // W6
  for (let k = 0; k < n; k++) {
    const t = ts[k];
    if (t === ES || t === ET || t === CS) ts[k] = ON;
  }
  // W7: European numbers in a left-to-right context are L.
  strong = sos;
  for (let k = 0; k < n; k++) {
    const t = ts[k]!;
    if (t === L || t === R) strong = t;
    else if (t === EN && strong === L) ts[k] = L;
  }

  const direction = level & 1 ? R : L;
  resolveBrackets(seq, cps, t0, ts, direction, sos);

  // N1/N2: neutral runs between same-direction text take that direction, else the embedding's.
  for (let k = 0; k < n; k++) {
    if (!isNeutral(ts[k]!)) continue;
    let end = k;
    while (end < n && isNeutral(ts[end]!)) end++;
    let lead = k === 0 ? sos : ts[k - 1]!;
    let trail = end === n ? eos : ts[end]!;
    if (lead === EN || lead === AN) lead = R;
    if (trail === EN || trail === AN) trail = R;
    ts.fill(lead === trail ? lead : direction, k, end);
    k = end - 1;
  }

  // I1, I2
  for (let k = 0; k < n; k++) {
    const t = ts[k]!;
    let resolved = level;
    if ((level & 1) === 0) {
      if (t === R) resolved += 1;
      else if (t === AN || t === EN) resolved += 2;
    } else if (t === L || t === EN || t === AN) resolved += 1;
    levels[seq[k]!] = resolved;
  }
}

/** Canonical equivalents U+2329/U+232A ≡ U+3008/U+3009 must pair with each other (BD16). */
function canonicalBracket(cp: number): number {
  return cp === 0x2329 ? 0x3008 : cp === 0x232a ? 0x3009 : cp;
}

/** N0's view of a type: EN and AN count as R. */
function strongForN0(t: number): number {
  if (t === L) return L;
  if (t === R || t === AL || t === EN || t === AN) return R;
  return ON;
}

/** N0: paired brackets take the direction of their content, or of the context around them. */
function resolveBrackets(seq: readonly number[], cps: readonly number[], t0: Uint8Array, ts: Uint8Array, direction: number, sos: number): void {
  const n = seq.length;
  // BD16: pair brackets that are still ON (overridden ones are not brackets any more).
  const openClosers: number[] = [];
  const openPositions: number[] = [];
  const pairs: [number, number][] = [];
  for (let k = 0; k < n; k++) {
    if (ts[k] !== ON) continue;
    const cp = cps[seq[k]!]!;
    const kind = pairedBracketType(cp);
    if (kind === BPT_OPEN) {
      if (openPositions.length === MAX_BRACKET_STACK) break;
      openClosers.push(canonicalBracket(pairedBracket(cp)!));
      openPositions.push(k);
    } else if (kind === BPT_CLOSE) {
      const closer = canonicalBracket(cp);
      for (let s = openPositions.length - 1; s >= 0; s--) {
        if (openClosers[s] !== closer) continue;
        pairs.push([openPositions[s]!, k]);
        openPositions.length = s;
        openClosers.length = s;
        break;
      }
    }
  }
  pairs.sort((a, b) => a[0] - b[0]);

  // In order of the opening brackets, so pairs resolved earlier inform later ones.
  for (const [open, close] of pairs) {
    let inside: number = ON;
    for (let k = open + 1; k < close; k++) {
      const s = strongForN0(ts[k]!);
      if (s === ON) continue;
      inside = s;
      if (s === direction) break;
    }
    if (inside === ON) continue; // N0 d: no strong type inside, leave to N1/N2
    let resolved = inside; // N0 b: matches the embedding direction
    if (inside !== direction) {
      // N0 c: opposite direction inside; keep it only if the preceding context agrees.
      let context = sos;
      for (let k = open - 1; k >= 0; k--) {
        const s = strongForN0(ts[k]!);
        if (s !== ON) {
          context = s;
          break;
        }
      }
      resolved = context === inside ? inside : direction;
    }
    ts[open] = resolved;
    ts[close] = resolved;
    // Marks that originally followed a bracket (W1 gave them its ON) follow its new type too.
    for (let k = open + 1; k < n && t0[seq[k]!] === NSM; k++) ts[k] = resolved;
    for (let k = close + 1; k < n && t0[seq[k]!] === NSM; k++) ts[k] = resolved;
  }
}
