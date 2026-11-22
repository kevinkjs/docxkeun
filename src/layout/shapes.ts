/**
 * Painting DrawingML content: shapes (preset and custom geometry, theme-styled fill and outline,
 * arrowheads), text boxes (Word paragraphs laid out inside the shape), groups and drawing
 * canvases, charts (their own part) and SmartArt (its pre-drawn dsp:drawing part).
 */
import { colorIn, fillOf, lineOf, dashPattern, type Paint, type SchemeColorFn } from '../charts/color.js';
import { renderChart } from '../charts/index.js';
import { customGeometry, lineEnd, presetGeometry, readAdjustments, readLineEnd, shadeFill, type ShapeGeometry } from '../drawing/index.js';
import type { Block, DrawingChild, DrawingContent, ImageInline, ShapeInfo } from '../model/content.js';
import type { DocxDocument } from '../model/document.js';
import { attr, elements, firstChild, textContent, type XmlElement } from '../xml/index.js';
import type { Box } from './blocks.js';
import { objectTransform, placeholder, transformItems, type DrawingContext } from './drawing.js';
import { measureChars, textItem, type Painter } from './paint.js';
import type { DisplayItem, FontProvider, PathCommand, Rect } from './types.js';

type Transform = ((x: number, y: number) => [number, number]) | undefined;

const compose = (inner: Transform, outer: Transform): Transform => (!inner ? outer : !outer ? inner : (x, y) => outer(...inner(x, y)));

/** Theme line widths for a:lnRef idx 1–3 (Office theme). */
const THEME_LINE_WIDTHS = [0, 0.5, 1, 1.5];

export function createDrawingContext(doc: DocxDocument, fonts: FontProvider, layoutBlocks: (blocks: Block[], width: number) => Box[]): DrawingContext {
  const scheme: SchemeColorFn = (name) => doc.theme.color(name) ?? doc.theme.colors.get(name);
  const ctx: DrawingContext = {
    doc,
    fonts,
    content(p, image, frame) {
      const content = image.content;
      if (!content) return false;
      const t = objectTransform(frame, image.rotation, image.flipH, image.flipV);
      return drawContent(p, content, frame, t);
    },
  };

  const emit = (p: Painter, items: DisplayItem[], t: Transform) => {
    for (const item of transformItems(items, t)) p.page.items.push(item);
  };

  function drawContent(p: Painter, content: DrawingContent, frame: Rect, t: Transform): boolean {
    switch (content.type) {
      case 'shape':
        drawShape(p, content.shape, frame, t);
        return true;
      case 'group':
        for (const child of content.children) drawChild(p, child, frame, t);
        return true;
      case 'picture':
        p.page.items.push({ type: 'image', ...frame, target: content.target, external: content.external });
        return true;
      case 'chart': {
        const chartSpace = doc.xmlSync(content.part)?.root;
        if (!chartSpace) return false;
        emit(p, renderChart(chartSpace, frame, { fonts, schemeColor: scheme, defaultFont: doc.theme.font('minorHAnsi') }), t);
        return true;
      }
      case 'diagram':
        return drawDiagram(p, content.dataPart, frame, t);
    }
  }

  function drawChild(p: Painter, child: DrawingChild, group: Rect, t: Transform): void {
    const frame = { x: group.x + child.x, y: group.y + child.y, w: child.w, h: child.h };
    const own = objectTransform(frame, child.rotation, child.flipH, child.flipV);
    if (child.content.type === 'picture') {
      // Pictures cannot be rotated as paths: place them by their transformed centre.
      const [cx, cy] = (compose(own, t) ?? ((x: number, y: number) => [x, y] as [number, number]))(frame.x + frame.w / 2, frame.y + frame.h / 2);
      p.page.items.push({ type: 'image', x: cx - frame.w / 2, y: cy - frame.h / 2, w: frame.w, h: frame.h, target: child.content.target, external: child.content.external, rotation: child.rotation || undefined });
      return;
    }
    drawContent(p, child.content, frame, compose(own, t));
  }

  /** Fill and outline: spPr first, then the shape style's theme references, then a VML shape's colours. */
  function paints(shape: ShapeInfo): { fill: Paint | null; line: { paint: Paint | null; width: number; dash?: number[]; cap?: 'butt' | 'round' | 'square'; join?: 'miter' | 'round' | 'bevel' } } {
    if (shape.vml) {
      return {
        fill: shape.vml.fill ? { color: shape.vml.fill.toUpperCase(), opacity: 1 } : null,
        line: { paint: shape.vml.stroke ? { color: shape.vml.stroke.toUpperCase(), opacity: 1 } : null, width: 0.75 },
      };
    }
    const style = shape.style;
    let fill = fillOf(shape.spPr, scheme);
    if (fill === undefined) {
      const ref = firstChild(style, 'a:fillRef');
      fill = ref && Number(attr(ref, 'idx') ?? 0) > 0 ? (colorIn(ref, scheme) ?? null) : null;
    }
    const ln = lineOf(shape.spPr, scheme);
    const lnRef = firstChild(style, 'a:lnRef');
    const refIdx = Number(attr(lnRef, 'idx') ?? 0);
    let paint = ln.paint;
    if (paint === undefined) paint = lnRef && refIdx > 0 ? (colorIn(lnRef, scheme) ?? null) : null;
    const width = ln.width ?? (THEME_LINE_WIDTHS[refIdx] || 0.75);
    return { fill, line: { paint, width, dash: dashPattern(ln.dash, width), cap: ln.cap, join: ln.join } };
  }

  function geometryOf(shape: ShapeInfo, w: number, h: number): ShapeGeometry {
    const g = shape.geometry;
    if (g?.name === 'a:prstGeom') {
      const out = presetGeometry(attr(g, 'prst') ?? 'rect', w, h, readAdjustments(g));
      if (out) return out;
    } else if (g?.name === 'a:custGeom') {
      try {
        return customGeometry(g, w, h);
      } catch {
        // fall back to a rectangle
      }
    }
    return { paths: [{ commands: rect(0, 0, w, h), fill: 'norm', stroke: true }], textRect: { x: 0, y: 0, w, h } };
  }

  function drawShape(p: Painter, shape: ShapeInfo, frame: Rect, t: Transform): void {
    const geometry = geometryOf(shape, frame.w, frame.h);
    const { fill, line } = paints(shape);
    const items: DisplayItem[] = [];
    const ln = firstChild(shape.spPr, 'a:ln');
    const head = readLineEnd(firstChild(ln, 'a:headEnd'));
    const tail = readLineEnd(firstChild(ln, 'a:tailEnd'));
    geometry.paths.forEach((path, i) => {
      const commands = offset(path.commands, frame.x, frame.y);
      const filled = fill && path.fill !== 'none' ? shadeFill(fill.color, path.fill) : undefined;
      const stroked = line.paint && path.stroke ? line.paint : undefined;
      if (!filled && !stroked) return;
      let drawn = commands;
      // Arrowheads on the first open path.
      const ends: DisplayItem[] = [];
      if (stroked && i === 0 && (head || tail)) {
        const pts = points(commands);
        if (pts.length >= 2) {
          if (head) {
            const e = lineEnd(head.type, head.w, head.len, pts[0]!, pts[1]!, line.width);
            ends.push(endItem(e.commands, e.filled, stroked, line.width));
            drawn = pullBack(drawn, 'start', e.inset);
          }
          if (tail) {
            const e = lineEnd(tail.type, tail.w, tail.len, pts[pts.length - 1]!, pts[pts.length - 2]!, line.width);
            ends.push(endItem(e.commands, e.filled, stroked, line.width));
            drawn = pullBack(drawn, 'end', e.inset);
          }
        }
      }
      items.push({
        type: 'path',
        commands: drawn,
        fill: filled,
        fillOpacity: filled && fill!.opacity < 1 ? fill!.opacity : undefined,
        stroke: stroked?.color,
        strokeOpacity: stroked && stroked.opacity < 1 ? stroked.opacity : undefined,
        strokeWidth: line.width,
        dash: line.dash,
        lineCap: line.cap,
        lineJoin: line.join,
      });
      items.push(...ends);
    });
    emit(p, items, t);

    const tr = geometry.textRect;
    if (shape.text) drawTextBox(p, shape.text, { x: frame.x + tr.x, y: frame.y + tr.y, w: tr.w, h: tr.h }, t);
    else if (shape.txBody) emit(p, drawTextBody(shape.txBody, { x: frame.x + tr.x, y: frame.y + tr.y, w: tr.w, h: tr.h }, shape.style), t);
  }

  function drawTextBox(p: Painter, text: NonNullable<ShapeInfo['text']>, area: Rect, t: Transform): void {
    const ins = text.insets;
    const inner = { x: area.x + ins.left, y: area.y + ins.top, w: Math.max(1, area.w - ins.left - ins.right), h: Math.max(0, area.h - ins.top - ins.bottom) };
    const boxes = layoutBlocks(text.blocks, inner.w);
    const height = boxes.reduce((s, b) => s + b.height, 0);
    let y = inner.y;
    if (text.anchor === 'ctr') y += (inner.h - height) / 2;
    else if (text.anchor === 'b') y += inner.h - height;
    // Text moves with a rotated box but stays upright; lines go to the editing layer as a text box story.
    const [ox, oy] = t ? t(inner.x + inner.w / 2, inner.y + inner.h / 2) : [inner.x + inner.w / 2, inner.y + inner.h / 2];
    const dx = ox - (inner.x + inner.w / 2);
    const dy = oy - (inner.y + inner.h / 2);
    const saved = { story: p.story, ref: p.storyRef, column: p.column };
    p.story = 'textbox';
    p.storyRef = { part: doc.mainPart, container: text.container };
    p.column = { x: inner.x + dx, width: inner.w };
    for (const b of boxes) {
      // Lines past the bottom of a fixed-size box are cut off, as in Word.
      if (!text.autoFit && y + b.height > area.y + area.h + 0.5 && y > inner.y) break;
      b.paint(p, inner.x + dx, y + dy);
      y += b.height;
    }
    p.story = saved.story;
    p.storyRef = saved.ref;
    p.column = saved.column;
  }

  /** DrawingML text (a:txBody): SmartArt and canvas shapes. Simple wrapping, per-run size, colour and weight. */
  function drawTextBody(txBody: XmlElement, area: Rect, style: XmlElement | undefined): DisplayItem[] {
    const bodyPr = firstChild(txBody, 'a:bodyPr');
    const emu = (name: string, d: number) => {
      const v = attr(bodyPr, name);
      return v === undefined ? d : Number(v) / 12700;
    };
    const inner = { x: area.x + emu('lIns', 7.2), y: area.y + emu('tIns', 3.6), w: Math.max(1, area.w - emu('lIns', 7.2) - emu('rIns', 7.2)), h: area.h - emu('tIns', 3.6) - emu('bIns', 3.6) };
    const fontRef = firstChild(style, 'a:fontRef');
    const styleColor = colorIn(fontRef, scheme)?.color;
    type Word = { text: string; width: number; font: ReturnType<FontProvider['get']>; size: number; color: string; space: boolean };
    const lines: Array<{ words: Word[]; width: number; height: number; align: string }> = [];
    for (const para of elements(txBody, 'a:p')) {
      const pPr = firstChild(para, 'a:pPr');
      const align = attr(pPr, 'algn') ?? 'ctr';
      const endSize = Number(attr(firstChild(para, 'a:endParaRPr'), 'sz') ?? 0) / 100;
      const words: Word[] = [];
      for (const r of elements(para)) {
        if (r.name !== 'a:r' && r.name !== 'a:fld') continue;
        const rPr = firstChild(r, 'a:rPr');
        const size = Number(attr(rPr, 'sz') ?? 0) / 100 || endSize || 14;
        const latin = attr(firstChild(rPr, 'a:latin'), 'typeface');
        const family = latin && !latin.startsWith('+') ? latin : (doc.theme.font(latin === '+mj-lt' ? 'majorHAnsi' : 'minorHAnsi') ?? 'Calibri');
        const font = fonts.get(family, attr(rPr, 'b') === '1', attr(rPr, 'i') === '1');
        const color = fillOf(rPr, scheme)?.color ?? styleColor ?? '000000';
        for (const piece of textContent(firstChild(r, 'a:t') ?? r).split(/(\s+)/)) {
          if (!piece) continue;
          const space = /^\s+$/.test(piece);
          words.push({ text: space ? ' ' : piece, width: measureChars(space ? ' ' : piece, font, size, 0).width, font, size, color, space });
        }
      }
      let line: Word[] = [];
      let width = 0;
      const flush = () => {
        while (line.length && line[line.length - 1]!.space) width -= line.pop()!.width;
        const height = Math.max(endSize || 14, ...line.map((w) => w.size)) * 1.2;
        lines.push({ words: line, width, height, align });
        line = [];
        width = 0;
      };
      for (const w of words) {
        if (!line.length && w.space) continue;
        if (width + w.width > inner.w && line.length) flush();
        if (!line.length && w.space) continue;
        line.push(w);
        width += w.width;
      }
      flush();
    }
    const total = lines.reduce((s, l) => s + l.height, 0);
    const anchor = attr(bodyPr, 'anchor') ?? 'ctr';
    let y = inner.y + (anchor === 'ctr' ? (inner.h - total) / 2 : anchor === 'b' ? inner.h - total : 0);
    const items: DisplayItem[] = [];
    for (const l of lines) {
      let x = inner.x + (l.align === 'ctr' ? (inner.w - l.width) / 2 : l.align === 'r' ? inner.w - l.width : 0);
      const baseline = y + l.height * 0.8;
      for (const w of l.words) {
        if (!w.space) {
          const m = measureChars(w.text, w.font, w.size, x);
          items.push(textItem(x, baseline, m.chars, m.xs, m.widths, w.font, w.size, w.color));
        }
        x += w.width;
      }
      y += l.height;
    }
    return items;
  }

  /** SmartArt: Word saves the laid-out diagram as shapes in a dsp:drawing part next to the data. */
  function drawDiagram(p: Painter, dataPart: string | undefined, frame: Rect, t: Transform): boolean {
    const data = doc.xmlSync(dataPart)?.root;
    if (!data) return false;
    let relId: string | undefined;
    for (const ext of elements(firstChild(data, 'dgm:extLst') ?? data, 'a:ext')) {
      const dme = firstChild(ext, 'dsp:dataModelExt');
      if (dme) relId = attr(dme, 'relId');
    }
    const drawingPart = relId ? doc.rels(doc.mainPart).get(relId)?.target : undefined;
    const drawing = doc.xmlSync(drawingPart)?.root;
    const tree = firstChild(drawing, 'dsp:spTree');
    if (!tree) {
      // No saved drawing (made by another program): list the diagram's text in a box.
      const texts = elements(firstChild(data, 'dgm:ptLst') ?? data, 'dgm:pt').map((pt) => textContent(firstChild(pt, 'dgm:t') ?? pt).trim()).filter(Boolean);
      emit(p, [placeholder(frame)], t);
      const fakeBody = { type: 'element', name: 'a:txBody', attrs: {}, children: texts.map((s) => ({ type: 'element', name: 'a:p', attrs: {}, children: [{ type: 'element', name: 'a:r', attrs: {}, children: [{ type: 'element', name: 'a:t', attrs: {}, children: [{ type: 'text', text: s }] }] }] })) } as unknown as XmlElement;
      emit(p, drawTextBody(fakeBody, frame, undefined), t);
      return true;
    }
    const toPt = (v: string | undefined) => Number(v ?? 0) / 12700;
    for (const sp of elements(tree, 'dsp:sp')) {
      const spPr = firstChild(sp, 'dsp:spPr');
      const xfrm = firstChild(spPr, 'a:xfrm');
      const off = firstChild(xfrm, 'a:off');
      const ext = firstChild(xfrm, 'a:ext');
      const box = { x: frame.x + toPt(attr(off, 'x')), y: frame.y + toPt(attr(off, 'y')), w: toPt(attr(ext, 'cx')), h: toPt(attr(ext, 'cy')) };
      const own = objectTransform(box, Number(attr(xfrm, 'rot') ?? 0) / 60000, attr(xfrm, 'flipH') === '1', attr(xfrm, 'flipV') === '1');
      const shape: ShapeInfo = { geometry: firstChild(spPr, 'a:prstGeom') ?? firstChild(spPr, 'a:custGeom'), spPr, style: firstChild(sp, 'dsp:style') };
      drawShape(p, shape, box, compose(own, t));
      // The text has its own rectangle (dsp:txXfrm) and is not rotated with the shape.
      const txBody = firstChild(sp, 'dsp:txBody');
      if (txBody) {
        const tx = firstChild(sp, 'dsp:txXfrm');
        const toff = firstChild(tx, 'a:off');
        const text = tx ? { x: frame.x + toPt(attr(toff, 'x')), y: frame.y + toPt(attr(toff, 'y')), w: toPt(attr(firstChild(tx, 'a:ext'), 'cx')), h: toPt(attr(firstChild(tx, 'a:ext'), 'cy')) } : box;
        emit(p, drawTextBody(txBody, text, shape.style), t);
      }
    }
    return true;
  }

  return ctx;
}

function rect(x: number, y: number, w: number, h: number): PathCommand[] {
  return [{ op: 'M', x, y }, { op: 'L', x: x + w, y }, { op: 'L', x: x + w, y: y + h }, { op: 'L', x, y: y + h }, { op: 'Z' }];
}

function offset(commands: PathCommand[], dx: number, dy: number): PathCommand[] {
  return commands.map((c) => (c.op === 'Z' ? c : c.op === 'C' ? { ...c, x1: c.x1 + dx, y1: c.y1 + dy, x2: c.x2 + dx, y2: c.y2 + dy, x: c.x + dx, y: c.y + dy } : { ...c, x: c.x + dx, y: c.y + dy }));
}

/** Points of a path in order (curve control points left out). */
function points(commands: PathCommand[]): Array<{ x: number; y: number }> {
  const out: Array<{ x: number; y: number }> = [];
  for (const c of commands) if (c.op !== 'Z') out.push({ x: c.x, y: c.y });
  return out;
}

function endItem(commands: PathCommand[], filled: boolean, paint: Paint, width: number): DisplayItem {
  return filled ? { type: 'path', commands, fill: paint.color, fillOpacity: paint.opacity < 1 ? paint.opacity : undefined } : { type: 'path', commands, stroke: paint.color, strokeWidth: width, lineJoin: 'miter' };
}

/** Shorten the first or last segment of a path by `inset` so a line's end hides inside its arrowhead. */
function pullBack(commands: PathCommand[], end: 'start' | 'end', inset: number): PathCommand[] {
  if (inset <= 0) return commands;
  const out = commands.map((c) => ({ ...c })) as PathCommand[];
  const pts = out.filter((c): c is Exclude<PathCommand, { op: 'Z' }> => c.op !== 'Z');
  if (pts.length < 2) return commands;
  const [a, b] = end === 'start' ? [pts[0]!, pts[1]!] : [pts[pts.length - 1]!, pts[pts.length - 2]!];
  const dx = b.x - a.x;
  const dy = b.y - a.y;
  const len = Math.hypot(dx, dy);
  if (len <= inset) return commands;
  a.x += (dx / len) * inset;
  a.y += (dy / len) * inset;
  return out;
}

export type { ImageInline };
