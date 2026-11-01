/** Run properties turned into concrete drawing style (fonts, colours, sizes). */
import type { RunProps } from '../model/props.js';
import { shadingColor } from '../model/props.js';
import type { Theme } from '../model/theme.js';

export interface RunStyle {
  fonts: { ascii?: string; hAnsi?: string; eastAsia?: string; cs?: string };
  size: number;
  sizeCs: number;
  bold: boolean;
  italic: boolean;
  boldCs: boolean;
  italicCs: boolean;
  /** RRGGBB, or 'auto'. */
  color: string;
  underline?: string;
  underlineColor?: string;
  strike: boolean;
  dstrike: boolean;
  caps: boolean;
  smallCaps: boolean;
  hidden: boolean;
  vertAlign: 'superscript' | 'subscript' | 'baseline';
  position: number;
  spacing: number;
  scale: number;
  highlight?: string;
  shading?: string;
  rtl: boolean;
  hint?: string;
}

export const HIGHLIGHT_COLORS: Record<string, string> = {
  yellow: 'FFFF00',
  green: '00FF00',
  cyan: '00FFFF',
  magenta: 'FF00FF',
  blue: '0000FF',
  red: 'FF0000',
  darkBlue: '000080',
  darkCyan: '008080',
  darkGreen: '008000',
  darkMagenta: '800080',
  darkRed: '800000',
  darkYellow: '808000',
  darkGray: '808080',
  lightGray: 'C0C0C0',
  black: '000000',
  white: 'FFFFFF',
};

const cache = new WeakMap<RunProps, RunStyle>();

export function resolveRunStyle(props: RunProps, theme: Theme): RunStyle {
  let style = cache.get(props);
  if (style) return style;
  const f = props.fonts ?? {};
  const color = props.themeColor ? theme.color(props.themeColor, props.themeTint, props.themeShade) ?? props.color : props.color;
  style = {
    fonts: {
      ascii: theme.font(f.asciiTheme) ?? f.ascii,
      hAnsi: theme.font(f.hAnsiTheme) ?? f.hAnsi ?? theme.font(f.asciiTheme) ?? f.ascii,
      eastAsia: theme.font(f.eastAsiaTheme) ?? f.eastAsia,
      cs: theme.font(f.csTheme) ?? f.cs,
    },
    size: props.size ?? 10,
    sizeCs: props.sizeCs ?? props.size ?? 10,
    bold: props.bold ?? false,
    italic: props.italic ?? false,
    boldCs: props.boldCs ?? props.bold ?? false,
    italicCs: props.italicCs ?? props.italic ?? false,
    color: !color || color.toLowerCase() === 'auto' ? 'auto' : color.toUpperCase(),
    underline: props.underline && props.underline !== 'none' ? props.underline : undefined,
    underlineColor: props.underlineColor,
    strike: props.strike ?? false,
    dstrike: props.dstrike ?? false,
    caps: props.caps ?? false,
    smallCaps: props.smallCaps ?? false,
    hidden: (props.vanish ?? false) || (props.specVanish ?? false) || (props.webHidden ?? false),
    vertAlign: props.vertAlign ?? 'baseline',
    position: props.position ?? 0,
    spacing: props.spacing ?? 0,
    scale: (props.scale ?? 100) / 100,
    highlight: props.highlight && props.highlight !== 'none' ? HIGHLIGHT_COLORS[props.highlight] : undefined,
    shading: shadingColor(props.shading),
    rtl: props.rtl ?? false,
    hint: f.hint,
  };
  cache.set(props, style);
  return style;
}

export type ScriptSlot = 'ascii' | 'hAnsi' | 'eastAsia' | 'cs';

export function isEastAsian(cp: number): boolean {
  return (
    (cp >= 0x1100 && cp <= 0x11ff) ||
    (cp >= 0x2e80 && cp <= 0x9fff) ||
    (cp >= 0xa960 && cp <= 0xa97f) ||
    (cp >= 0xac00 && cp <= 0xd7af) ||
    (cp >= 0xf900 && cp <= 0xfaff) ||
    (cp >= 0xfe30 && cp <= 0xfe4f) ||
    (cp >= 0xff00 && cp <= 0xffef) ||
    (cp >= 0x20000 && cp <= 0x2ffff)
  );
}

export function isComplexScript(cp: number): boolean {
  return (cp >= 0x0590 && cp <= 0x08ff) || (cp >= 0x0900 && cp <= 0x0dff) || (cp >= 0x0e00 && cp <= 0x0eff) || (cp >= 0xfb1d && cp <= 0xfdff) || (cp >= 0xfe70 && cp <= 0xfeff);
}

/** Which w:rFonts slot a character uses (ECMA-376 §17.3.2.26, simplified). */
export function scriptSlot(cp: number, style: RunStyle): ScriptSlot {
  if (style.rtl || isComplexScript(cp)) return 'cs';
  if (isEastAsian(cp)) return 'eastAsia';
  if (style.hint === 'eastAsia' && cp >= 0x2000 && cp <= 0x2bff) return 'eastAsia';
  return cp < 0x80 ? 'ascii' : 'hAnsi';
}

export function fontNameForSlot(style: RunStyle, slot: ScriptSlot): string | undefined {
  const f = style.fonts;
  switch (slot) {
    case 'ascii':
      return f.ascii ?? f.hAnsi;
    case 'hAnsi':
      return f.hAnsi ?? f.ascii;
    case 'eastAsia':
      return f.eastAsia ?? f.hAnsi ?? f.ascii;
    case 'cs':
      return f.cs ?? f.hAnsi ?? f.ascii;
  }
}
