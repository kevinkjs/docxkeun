/**
 * Headers and footers: creating, replacing and removing them, per section and per kind of page
 * (default, first, even), with page-number fields and left / centre / right parts.
 *
 * In Word a section's w:sectPr points at header and footer parts (w:headerReference /
 * w:footerReference by type); a "first" header needs w:titlePg in the section, and "even" headers
 * need w:evenAndOddHeaders in settings. A section without its own reference shows the previous
 * section's header.
 */
import type { DocxDocument } from '../model/document.js';
import { attr, el, elements, findAll, firstChild, insertOrdered, removeChildren, type XmlElement, type XmlNode } from '../xml/index.js';

export type HeaderFooterKind = 'header' | 'footer';
export type HeaderFooterType = 'default' | 'first' | 'even';

export interface HeaderFooterOptions {
  /**
   * Which pages: 'default' (every page, or the odd ones when even pages have their own), 'first'
   * (the section's first page) or 'even'. Default 'default'.
   */
  type?: HeaderFooterType;
  /** Section number from 1 (default: every section). */
  section?: number;
  /** Paragraph alignment (default: the Header/Footer style's, else left). */
  align?: 'left' | 'center' | 'right';
}

export interface HeaderFooterInfo {
  kind: HeaderFooterKind;
  type: HeaderFooterType;
  /** Sections (from 1) that use it. */
  sections: number[];
  part: string;
  /** Its text, with page fields as {PAGE}, {NUMPAGES}, {SECTIONPAGES}. */
  text: string;
}

/** Child order of w:sectPr (ECMA-376 CT_SectPr). */
export const SECTPR_ORDER = [
  'w:headerReference',
  'w:footerReference',
  'w:footnotePr',
  'w:endnotePr',
  'w:type',
  'w:pgSz',
  'w:pgMar',
  'w:paperSrc',
  'w:pgBorders',
  'w:lnNumType',
  'w:pgNumType',
  'w:cols',
  'w:formProt',
  'w:vAlign',
  'w:noEndnote',
  'w:titlePg',
  'w:textDirection',
  'w:bidi',
  'w:rtlGutter',
  'w:docGrid',
  'w:printerSettings',
  'w:sectPrChange',
];

/** The sections' w:sectPr elements in order: paragraph-level ones, then the body's last one. */
export function sectionProperties(body: XmlElement): XmlElement[] {
  const out: XmlElement[] = [];
  const walk = (node: XmlElement) => {
    for (const c of elements(node)) {
      // Not the old properties kept by a tracked change (w:sectPrChange, w:pPrChange).
      if (c.name === 'w:sectPrChange' || c.name === 'w:pPrChange') continue;
      if (c.name === 'w:sectPr') {
        if (node.name === 'w:pPr' || node === body) out.push(c);
        continue;
      }
      walk(c);
    }
  };
  walk(body);
  // The body's own sectPr is its last child; paragraph ones come in document order before it.
  const last = out.findIndex((s) => elements(body).includes(s));
  if (last >= 0 && last !== out.length - 1) out.push(...out.splice(last, 1));
  return out;
}

const REF = { header: 'w:headerReference', footer: 'w:footerReference' } as const;

export function referencesOf(sectPr: XmlElement, kind: HeaderFooterKind): XmlElement[] {
  return elements(sectPr, REF[kind]);
}

/** Text width of a section (page width minus margins), in twips. */
export function textWidth(sectPr: XmlElement | undefined): number {
  const w = Number(attr(firstChild(sectPr, 'w:pgSz'), 'w:w') ?? 12240);
  const mar = firstChild(sectPr, 'w:pgMar');
  const left = Number(attr(mar, 'w:left') ?? 1440);
  const right = Number(attr(mar, 'w:right') ?? 1440);
  const gutter = Number(attr(mar, 'w:gutter') ?? 0);
  return Math.max(1440, w - left - right - gutter);
}

const FIELDS = new Set(['PAGE', 'NUMPAGES', 'SECTIONPAGES']);

/**
 * Paragraphs for header or footer text: lines become paragraphs, a tab moves to the centre and then
 * the right edge ("Left\tCentre\tRight"), and {PAGE}, {NUMPAGES} and {SECTIONPAGES} become fields.
 */
export function headerParagraphs(doc: DocxDocument, kind: HeaderFooterKind, text: string, align: HeaderFooterOptions['align'], width: number): XmlElement[] {
  const styleId = kind === 'header' ? 'Header' : 'Footer';
  const hasStyle = !!doc.styles.get(styleId);
  return text.split(/\r?\n/).map((line) => {
    const pPr: XmlElement[] = [];
    if (hasStyle) pPr.push(el('w:pStyle', { 'w:val': styleId }));
    if (line.includes('\t') && !hasStyle) {
      pPr.push(el('w:tabs', {}, [el('w:tab', { 'w:val': 'center', 'w:pos': String(Math.round(width / 2)) }), el('w:tab', { 'w:val': 'right', 'w:pos': String(width) })]));
    }
    if (align) pPr.push(el('w:jc', { 'w:val': align }));
    const content: XmlElement[] = [];
    for (const piece of line.split(/(\t|\{(?:PAGE|NUMPAGES|SECTIONPAGES)\})/)) {
      if (!piece) continue;
      if (piece === '\t') content.push(el('w:r', {}, [el('w:tab')]));
      else if (piece.startsWith('{') && FIELDS.has(piece.slice(1, -1))) {
        content.push(el('w:fldSimple', { 'w:instr': ` ${piece.slice(1, -1)} ` }, [el('w:r', {}, [el('w:t', {}, [{ type: 'text', value: '1' }])])]));
      } else {
        const t = el('w:t', /^\s|\s$/.test(piece) ? { 'xml:space': 'preserve' } : {}, [{ type: 'text', value: piece }]);
        content.push(el('w:r', {}, [t]));
      }
    }
    return el('w:p', {}, [...(pPr.length ? [el('w:pPr', {}, pPr)] : []), ...content]);
  });
}

/** Header or footer paragraphs back to text, with page fields as {PAGE}. */
export function headerText(root: XmlElement): string {
  const lines: string[] = [];
  for (const p of findAll(root, 'w:p')) {
    let s = '';
    const walk = (n: XmlElement) => {
      for (const c of n.children as XmlNode[]) {
        if (c.type !== 'element') continue;
        if (c.name === 'w:delText' || c.name === 'w:del') continue;
        if (c.name === 'w:fldSimple') {
          const instr = (attr(c, 'w:instr') ?? '').trim().split(/\s+/)[0]?.toUpperCase() ?? '';
          s += FIELDS.has(instr) ? `{${instr}}` : '';
          continue;
        }
        if (c.name === 'w:t') s += c.children.map((t) => (t.type === 'text' ? t.value : '')).join('');
        else if (c.name === 'w:tab' && n.name === 'w:r') s += '\t';
        else walk(c);
      }
    };
    walk(p);
    lines.push(s);
  }
  return lines.join('\n');
}

/** Whether a section still has a first-page header or footer. */
export function hasFirstPage(sectPr: XmlElement): boolean {
  return [...referencesOf(sectPr, 'header'), ...referencesOf(sectPr, 'footer')].some((r) => attr(r, 'w:type') === 'first');
}

/** Whether any section has an even-page header or footer. */
export function hasEvenPages(sectPrs: XmlElement[]): boolean {
  return sectPrs.some((sp) => [...referencesOf(sp, 'header'), ...referencesOf(sp, 'footer')].some((r) => attr(r, 'w:type') === 'even'));
}
