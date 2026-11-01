/** styles.xml: document defaults, style inheritance (basedOn) and effective property resolution. */
import { attr, elements, firstChild, type XmlDocument, type XmlElement } from '../xml/index.js';
import {
  mergeCellProps,
  mergeParagraphProps,
  mergeRunProps,
  mergeTableProps,
  parseCellProps,
  parseParagraphProps,
  parseRowProps,
  parseRunProps,
  parseTableProps,
  TOGGLE_PROPS,
  type CellProps,
  type ParagraphProps,
  type RowProps,
  type RunProps,
  type TableProps,
} from './props.js';
import { parseOnOff } from './units.js';

export type StyleType = 'paragraph' | 'character' | 'table' | 'numbering';

export type ConditionalType =
  | 'wholeTable'
  | 'firstRow'
  | 'lastRow'
  | 'firstCol'
  | 'lastCol'
  | 'band1Vert'
  | 'band2Vert'
  | 'band1Horz'
  | 'band2Horz'
  | 'neCell'
  | 'nwCell'
  | 'seCell'
  | 'swCell';

export interface ConditionalFormat {
  pPr: ParagraphProps;
  rPr: RunProps;
  tblPr: TableProps;
  trPr: RowProps;
  tcPr: CellProps;
}

export interface StyleDef {
  id: string;
  type: StyleType;
  name?: string;
  basedOn?: string;
  next?: string;
  link?: string;
  isDefault: boolean;
  pPr: ParagraphProps;
  rPr: RunProps;
  tblPr: TableProps;
  trPr: RowProps;
  tcPr: CellProps;
  conditional: Map<ConditionalType, ConditionalFormat>;
  el: XmlElement;
}

export interface ResolvedTableStyle {
  tblPr: TableProps;
  trPr: RowProps;
  tcPr: CellProps;
  pPr: ParagraphProps;
  rPr: RunProps;
  conditional: Map<ConditionalType, ConditionalFormat>;
}

export class StyleSheet {
  defaultRunProps: RunProps = {};
  defaultParagraphProps: ParagraphProps = {};
  readonly styles = new Map<string, StyleDef>();
  readonly defaults: Partial<Record<StyleType, string>> = {};
  private readonly paraCache = new Map<string, { pPr: ParagraphProps; rPr: RunProps; toggles: Partial<Record<string, boolean>> }>();
  private readonly charCache = new Map<string, { rPr: RunProps; toggles: Partial<Record<string, boolean>> }>();
  private readonly tableCache = new Map<string, ResolvedTableStyle>();

  constructor(readonly doc?: XmlDocument) {
    if (!doc) return;
    const docDefaults = firstChild(doc.root, 'w:docDefaults');
    this.defaultRunProps = parseRunProps(firstChild(firstChild(docDefaults, 'w:rPrDefault'), 'w:rPr'));
    this.defaultParagraphProps = parseParagraphProps(firstChild(firstChild(docDefaults, 'w:pPrDefault'), 'w:pPr'));
    for (const s of elements(doc.root, 'w:style')) {
      const id = attr(s, 'w:styleId');
      if (!id) continue;
      const type = (attr(s, 'w:type') ?? 'paragraph') as StyleType;
      const conditional = new Map<ConditionalType, ConditionalFormat>();
      for (const c of elements(s, 'w:tblStylePr')) {
        const ctype = attr(c, 'w:type') as ConditionalType | undefined;
        if (!ctype) continue;
        conditional.set(ctype, {
          pPr: parseParagraphProps(firstChild(c, 'w:pPr')),
          rPr: parseRunProps(firstChild(c, 'w:rPr')),
          tblPr: parseTableProps(firstChild(c, 'w:tblPr')),
          trPr: parseRowProps(firstChild(c, 'w:trPr')),
          tcPr: parseCellProps(firstChild(c, 'w:tcPr')),
        });
      }
      const def: StyleDef = {
        id,
        type,
        name: attr(firstChild(s, 'w:name'), 'w:val'),
        basedOn: attr(firstChild(s, 'w:basedOn'), 'w:val'),
        next: attr(firstChild(s, 'w:next'), 'w:val'),
        link: attr(firstChild(s, 'w:link'), 'w:val'),
        isDefault: attr(s, 'w:default') !== undefined && parseOnOff(attr(s, 'w:default')),
        pPr: parseParagraphProps(firstChild(s, 'w:pPr')),
        rPr: parseRunProps(firstChild(s, 'w:rPr')),
        tblPr: parseTableProps(firstChild(s, 'w:tblPr')),
        trPr: parseRowProps(firstChild(s, 'w:trPr')),
        tcPr: parseCellProps(firstChild(s, 'w:tcPr')),
        conditional,
        el: s,
      };
      // A style's own w:pStyle / w:rStyle references are not meaningful; drop them.
      delete def.pPr.styleId;
      delete def.rPr.styleId;
      this.styles.set(id, def);
      if (def.isDefault && !this.defaults[type]) this.defaults[type] = id;
    }
  }

  get(id: string | undefined): StyleDef | undefined {
    return id ? this.styles.get(id) : undefined;
  }

  /** Find a style by its display name (e.g. "heading 1"), case-insensitively. */
  byName(name: string): StyleDef | undefined {
    const lower = name.toLowerCase();
    for (const s of this.styles.values()) if (s.name?.toLowerCase() === lower) return s;
    return undefined;
  }

  /** Root-first inheritance chain, guarded against cycles and type mismatches. */
  chain(id: string | undefined, type: StyleType): StyleDef[] {
    const out: StyleDef[] = [];
    const seen = new Set<string>();
    let cur = this.get(id);
    while (cur && !seen.has(cur.id) && cur.type === type) {
      seen.add(cur.id);
      out.unshift(cur);
      cur = this.get(cur.basedOn);
    }
    return out;
  }

  paragraphStyleId(explicit: string | undefined): string | undefined {
    if (explicit && this.styles.get(explicit)?.type === 'paragraph') return explicit;
    return this.defaults.paragraph;
  }

  /** Merged pPr / rPr of a paragraph style chain, with the chain's own toggle values. */
  paragraphStyle(id: string | undefined): { pPr: ParagraphProps; rPr: RunProps; toggles: Partial<Record<string, boolean>> } {
    const key = id ?? '';
    let cached = this.paraCache.get(key);
    if (!cached) {
      const chain = this.chain(id, 'paragraph');
      const rPr = mergeRunProps(...chain.map((s) => s.rPr));
      cached = { pPr: mergeParagraphProps(...chain.map((s) => s.pPr)), rPr, toggles: pickToggles(rPr) };
      this.paraCache.set(key, cached);
    }
    return cached;
  }

  characterStyle(id: string | undefined): { rPr: RunProps; toggles: Partial<Record<string, boolean>> } {
    const key = id ?? '';
    let cached = this.charCache.get(key);
    if (!cached) {
      const chain = this.chain(id ?? this.defaults.character, 'character');
      const rPr = mergeRunProps(...chain.map((s) => s.rPr));
      cached = { rPr, toggles: pickToggles(rPr) };
      this.charCache.set(key, cached);
    }
    return cached;
  }

  tableStyle(id: string | undefined): ResolvedTableStyle {
    const key = id ?? '';
    let cached = this.tableCache.get(key);
    if (!cached) {
      const chain = this.chain(id ?? this.defaults.table, 'table');
      const conditional = new Map<ConditionalType, ConditionalFormat>();
      for (const s of chain) {
        for (const [type, fmt] of s.conditional) {
          const prev = conditional.get(type);
          conditional.set(
            type,
            prev
              ? {
                  pPr: mergeParagraphProps(prev.pPr, fmt.pPr),
                  rPr: mergeRunProps(prev.rPr, fmt.rPr),
                  tblPr: mergeTableProps(prev.tblPr, fmt.tblPr),
                  trPr: { ...prev.trPr, ...fmt.trPr },
                  tcPr: mergeCellProps(prev.tcPr, fmt.tcPr),
                }
              : fmt,
          );
        }
      }
      cached = {
        tblPr: mergeTableProps(...chain.map((s) => s.tblPr)),
        trPr: Object.assign({}, ...chain.map((s) => s.trPr)) as RowProps,
        tcPr: mergeCellProps(...chain.map((s) => s.tcPr)),
        pPr: mergeParagraphProps(...chain.map((s) => s.pPr)),
        rPr: mergeRunProps(...chain.map((s) => s.rPr)),
        conditional,
      };
      this.tableCache.set(key, cached);
    }
    return cached;
  }

  /**
   * Effective run properties: docDefaults → table style → paragraph style → character style →
   * direct formatting. Toggle properties XOR across the style levels (ECMA-376 §17.7.3); direct
   * formatting sets them absolutely.
   */
  resolveRun(paragraphStyleId: string | undefined, direct: RunProps, tableRunProps?: RunProps): RunProps {
    const para = this.paragraphStyle(paragraphStyleId);
    const char = this.characterStyle(direct.styleId);
    const merged = mergeRunProps(this.defaultRunProps, tableRunProps, para.rPr, char.rPr, direct);
    const tableToggles = tableRunProps ? pickToggles(tableRunProps) : {};
    for (const t of TOGGLE_PROPS) {
      if (direct[t] !== undefined) continue;
      let v = this.defaultRunProps[t] ?? false;
      for (const level of [tableToggles, para.toggles, char.toggles]) if (level[t] === true) v = !v;
      merged[t] = v;
    }
    return merged;
  }

  /** Paragraph properties without numbering (numbering is merged by the content reader). */
  resolveParagraph(styleId: string | undefined, direct: ParagraphProps, tableParagraphProps?: ParagraphProps): ParagraphProps {
    return mergeParagraphProps(this.defaultParagraphProps, tableParagraphProps, this.paragraphStyle(styleId).pPr, direct);
  }
}

function pickToggles(rPr: RunProps): Partial<Record<string, boolean>> {
  const out: Partial<Record<string, boolean>> = {};
  for (const t of TOGGLE_PROPS) if (rPr[t] !== undefined) out[t] = rPr[t];
  return out;
}
