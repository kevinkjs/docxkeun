/**
 * Glyph outlines from TrueType `glyf` data, for chart text that is rotated (vertical axis titles,
 * rotated tick labels). Display-list TextItems are always horizontal, so rotated text is drawn as
 * filled paths. Fonts without glyf outlines (CFF, or fonts known only by metrics) return undefined
 * and the caller falls back to horizontal text.
 */
import type { LayoutFont, PathCommand } from '../layout/types.js';

interface Pt {
  x: number;
  y: number;
  on: boolean;
}

type Contours = Pt[][];

const cache = new WeakMap<object, Map<number, Contours | null>>();

function readGlyph(font: NonNullable<LayoutFont['face']>['font'], gid: number, depth: number): Contours | null {
  const range = font.glyphRange(gid);
  if (!range) return null;
  if (range.length === 0) return [];
  const data = font.data;
  const view = new DataView(data.buffer, data.byteOffset, data.byteLength);
  let p = range.offset;
  const numberOfContours = view.getInt16(p);
  p += 10;
  if (numberOfContours >= 0) {
    const ends: number[] = [];
    for (let i = 0; i < numberOfContours; i++, p += 2) ends.push(view.getUint16(p));
    const nPoints = numberOfContours ? ends[ends.length - 1]! + 1 : 0;
    p += 2 + view.getUint16(p); // instructions
    const flags: number[] = [];
    while (flags.length < nPoints) {
      const f = view.getUint8(p++);
      flags.push(f);
      if (f & 8) {
        const repeat = view.getUint8(p++);
        for (let r = 0; r < repeat; r++) flags.push(f);
      }
    }
    const xs: number[] = [];
    let v = 0;
    for (const f of flags) {
      if (f & 2) {
        const d = view.getUint8(p++);
        v += f & 16 ? d : -d;
      } else if (!(f & 16)) {
        v += view.getInt16(p);
        p += 2;
      }
      xs.push(v);
    }
    const ys: number[] = [];
    v = 0;
    for (const f of flags) {
      if (f & 4) {
        const d = view.getUint8(p++);
        v += f & 32 ? d : -d;
      } else if (!(f & 32)) {
        v += view.getInt16(p);
        p += 2;
      }
      ys.push(v);
    }
    const contours: Contours = [];
    let start = 0;
    for (const end of ends) {
      const c: Pt[] = [];
      for (let i = start; i <= end; i++) c.push({ x: xs[i]!, y: ys[i]!, on: (flags[i]! & 1) === 1 });
      contours.push(c);
      start = end + 1;
    }
    return contours;
  }
  // Composite glyph: components with offsets and optional scale/2x2 transforms.
  if (depth > 8) return [];
  const out: Contours = [];
  let more = true;
  while (more) {
    const flags = view.getUint16(p);
    const glyphIndex = view.getUint16(p + 2);
    p += 4;
    let dx: number;
    let dy: number;
    if (flags & 1) {
      dx = view.getInt16(p);
      dy = view.getInt16(p + 2);
      p += 4;
    } else {
      dx = view.getInt8(p);
      dy = view.getInt8(p + 1);
      p += 2;
    }
    if (!(flags & 2)) dx = dy = 0; // point matching: rare in text fonts, placed unshifted
    let a = 1;
    let b = 0;
    let c = 0;
    let d = 1;
    const f2 = (o: number): number => view.getInt16(o) / 16384;
    if (flags & 8) {
      a = d = f2(p);
      p += 2;
    } else if (flags & 0x40) {
      a = f2(p);
      d = f2(p + 2);
      p += 4;
    } else if (flags & 0x80) {
      a = f2(p);
      b = f2(p + 2);
      c = f2(p + 4);
      d = f2(p + 6);
      p += 8;
    }
    const part = readGlyph(font, glyphIndex, depth + 1) ?? [];
    for (const contour of part) out.push(contour.map((pt) => ({ x: pt.x * a + pt.y * c + dx, y: pt.x * b + pt.y * d + dy, on: pt.on })));
    more = (flags & 0x20) !== 0;
  }
  return out;
}

function glyphContours(font: LayoutFont, gid: number): Contours | null {
  const face = font.face;
  if (!face || !face.font.tables.has('glyf')) return null;
  let perFont = cache.get(face.font);
  if (!perFont) {
    perFont = new Map();
    cache.set(face.font, perFont);
  }
  let contours = perFont.get(gid);
  if (contours === undefined) {
    try {
      contours = readGlyph(face.font, gid, 0);
    } catch {
      contours = null;
    }
    perFont.set(gid, contours);
  }
  return contours;
}

/** Whether rotated text can be drawn in this font. */
export function hasOutlines(font: LayoutFont): boolean {
  return font.face !== undefined && font.face.font.tables.has('glyf');
}

/**
 * Outline of a run of text as path commands. `place(u, v)` maps text space (u along the baseline from
 * the run's start, v downward from the baseline, in points) to page space.
 */
export function textOutline(chars: string[], font: LayoutFont, size: number, place: (u: number, v: number) => { x: number; y: number }): PathCommand[] | undefined {
  const face = font.face;
  if (!face) return undefined;
  const scale = size / face.font.unitsPerEm;
  const slant = font.syntheticItalic ? 0.2 : 0;
  const cmds: PathCommand[] = [];
  let pen = 0;
  for (const ch of chars) {
    const cp = ch.codePointAt(0)!;
    const contours = glyphContours(font, font.glyphId(cp));
    if (contours === null) return undefined;
    const map = (pt: { x: number; y: number }): { x: number; y: number } => place(pen + (pt.x + pt.y * slant) * scale, -pt.y * scale);
    for (const contour of contours) {
      if (contour.length === 0) continue;
      // Start on an on-curve point (or the midpoint of two off-curve points).
      const n = contour.length;
      let startIdx = contour.findIndex((q) => q.on);
      let start: { x: number; y: number };
      if (startIdx < 0) {
        const a = contour[0]!;
        const b = contour[n - 1]!;
        start = { x: (a.x + b.x) / 2, y: (a.y + b.y) / 2 };
        startIdx = 0; // all off-curve: begin between the last and first points
      } else start = contour[startIdx]!;
      const s = map(start);
      cmds.push({ op: 'M', x: s.x, y: s.y });
      let cur = start;
      let ctrl: { x: number; y: number } | undefined;
      const allOff = !contour[startIdx]!.on;
      const total = allOff ? n : n - 1;
      const quad = (c: { x: number; y: number }, to: { x: number; y: number }): void => {
        const p0 = map(cur);
        const pc = map(c);
        const p1 = map(to);
        cmds.push({ op: 'C', x1: p0.x + (2 / 3) * (pc.x - p0.x), y1: p0.y + (2 / 3) * (pc.y - p0.y), x2: p1.x + (2 / 3) * (pc.x - p1.x), y2: p1.y + (2 / 3) * (pc.y - p1.y), x: p1.x, y: p1.y });
        cur = to;
      };
      for (let k = 1; k <= total + 1; k++) {
        const pt = k <= total ? contour[(startIdx + k + (allOff ? -1 : 0) + n) % n]! : undefined;
        const target = pt ?? { x: start.x, y: start.y, on: true };
        if (target.on) {
          if (ctrl) quad(ctrl, target);
          else {
            const t = map(target);
            cmds.push({ op: 'L', x: t.x, y: t.y });
            cur = target;
          }
          ctrl = undefined;
        } else if (ctrl) {
          const mid = { x: (ctrl.x + target.x) / 2, y: (ctrl.y + target.y) / 2 };
          quad(ctrl, mid);
          ctrl = target;
        } else ctrl = target;
      }
      cmds.push({ op: 'Z' });
    }
    pen += font.advance(cp) * size;
  }
  return cmds;
}
