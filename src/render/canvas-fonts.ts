/**
 * Font provider for browsers without registered font files: measures with the canvas API and
 * the fonts installed on the viewer's system. Good for viewing; PDF export needs real files.
 */
import { METRIC_COMPATIBLE } from '../fonts/registry.js';
import type { FontProvider, LayoutFont } from '../layout/types.js';

const MEASURE_SIZE = 100;

function cssFamily(name: string | undefined): string {
  const families: string[] = [];
  if (name) {
    families.push(name);
    const substitute = METRIC_COMPATIBLE[name.toLowerCase()];
    if (substitute) families.push(substitute);
  }
  families.push('Arial', 'Helvetica', 'sans-serif');
  return families.map((f) => (/^[a-z-]+$/i.test(f) && f !== 'Arial' ? f : JSON.stringify(f))).join(', ');
}

class CanvasLayoutFont implements LayoutFont {
  readonly key: string;
  readonly ascent: number;
  readonly descent: number;
  readonly underlinePosition = 0.1;
  readonly underlineThickness = 0.05;
  readonly syntheticBold = false;
  readonly syntheticItalic = false;
  private readonly cache = new Map<number, number>();

  constructor(
    private readonly ctx: CanvasRenderingContext2D,
    readonly family: string,
    readonly bold: boolean,
    readonly italic: boolean,
    key: string,
  ) {
    this.key = key;
    this.setFont();
    const m = ctx.measureText('Hgjy');
    const ascent = m.fontBoundingBoxAscent ?? m.actualBoundingBoxAscent ?? MEASURE_SIZE * 0.9;
    const descent = m.fontBoundingBoxDescent ?? m.actualBoundingBoxDescent ?? MEASURE_SIZE * 0.22;
    this.ascent = ascent / MEASURE_SIZE;
    this.descent = descent / MEASURE_SIZE;
  }

  private setFont(): void {
    this.ctx.font = `${this.italic ? 'italic ' : ''}${this.bold ? 'bold ' : ''}${MEASURE_SIZE}px ${this.family}`;
  }

  advance(cp: number): number {
    let a = this.cache.get(cp);
    if (a === undefined) {
      this.setFont();
      a = this.ctx.measureText(String.fromCodePoint(cp)).width / MEASURE_SIZE;
      this.cache.set(cp, a);
    }
    return a;
  }

  glyphId(cp: number): number {
    return cp;
  }

  has(): boolean {
    return true;
  }
}

export class CanvasFontProvider implements FontProvider {
  private readonly ctx: CanvasRenderingContext2D;
  private readonly fonts = new Map<string, CanvasLayoutFont>();

  constructor() {
    const canvas = typeof OffscreenCanvas !== 'undefined' ? new OffscreenCanvas(8, 8) : document.createElement('canvas');
    const ctx = canvas.getContext('2d') as CanvasRenderingContext2D | null;
    if (!ctx) throw new Error('Canvas 2D is not available');
    this.ctx = ctx;
  }

  get(family: string | undefined, bold: boolean, italic: boolean): LayoutFont {
    const css = cssFamily(family);
    const key = `c${this.fonts.size}`;
    const id = `${css}|${bold}|${italic}`;
    let font = this.fonts.get(id);
    if (!font) {
      font = new CanvasLayoutFont(this.ctx, css, bold, italic, key);
      this.fonts.set(id, font);
    }
    return font;
  }

  forChar(font: LayoutFont): LayoutFont {
    return font;
  }
}
