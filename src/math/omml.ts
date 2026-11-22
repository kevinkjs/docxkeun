/**
 * Reading OMML properties. Math properties sit in `m:*Pr` elements as children with an `m:val`
 * attribute; on/off properties are on when the element is present without a value.
 */
import { attr, elements, firstChild, textContent } from '../xml/dom.js';
import type { XmlElement } from '../xml/dom.js';

const OFF = new Set(['0', 'off', 'false']);

/** `m:val` of a property child, undefined when the property is absent. */
export function val(pr: XmlElement | undefined, name: string): string | undefined {
  return attr(firstChild(pr, name), 'm:val');
}

/** A character property (m:chr, m:begChr…). Present without m:val means the default. */
export function charProp(pr: XmlElement | undefined, name: string, fallback: string): string {
  const v = val(pr, name);
  return v === undefined ? fallback : v;
}

export function onOff(pr: XmlElement | undefined, name: string, fallback: boolean): boolean {
  const e = firstChild(pr, name);
  if (!e) return fallback;
  const v = attr(e, 'm:val') ?? attr(e, 'w:val');
  return v === undefined || !OFF.has(v.toLowerCase());
}

export function numProp(pr: XmlElement | undefined, name: string): number | undefined {
  const v = val(pr, name);
  if (v === undefined) return undefined;
  const n = Number(v);
  return Number.isFinite(n) ? n : undefined;
}

/** WordprocessingML on/off (w:b, w:i): `w:val` false values turn it off. */
export function wOnOff(rPr: XmlElement | undefined, name: string): boolean | undefined {
  const e = firstChild(rPr, name);
  if (!e) return undefined;
  const v = attr(e, 'w:val');
  return v === undefined || !OFF.has(v.toLowerCase());
}

/** Colour from `m:ctrlPr/w:rPr/w:color` of a structure, which Word uses for bars and delimiters. */
export function ctrlColor(pr: XmlElement | undefined): string | undefined {
  const color = attr(firstChild(firstChild(firstChild(pr, 'm:ctrlPr'), 'w:rPr'), 'w:color'), 'w:val');
  return color && /^[0-9a-fA-F]{6}$/.test(color) ? color.toUpperCase() : undefined;
}

/** Text of a run: its m:t (and stray w:t) children in order. */
export function runText(run: XmlElement): string {
  let s = '';
  for (const c of elements(run)) {
    if (c.name === 'm:t' || c.name === 'w:t') s += textContent(c);
    else if (c.name === 'w:tab') s += ' ';
  }
  return s;
}

/** Elements that only carry properties or bookkeeping, never equation content. */
export function isPropertyElement(name: string): boolean {
  return name.endsWith('Pr') || name === 'm:ctrlPr';
}
