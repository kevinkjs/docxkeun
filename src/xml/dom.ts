/** A light XML tree: just enough structure to model OOXML parts and write them back faithfully. */

export interface XmlElement {
  type: 'element';
  /** Qualified name as written, e.g. `w:p` (prefixes are canonicalised on parse, see ns.ts). */
  name: string;
  /** Attributes in document order, keyed by qualified name. */
  attrs: Record<string, string>;
  children: XmlNode[];
}

export interface XmlText {
  type: 'text';
  value: string;
}

export interface XmlCData {
  type: 'cdata';
  value: string;
}

export interface XmlComment {
  type: 'comment';
  value: string;
}

export interface XmlProcessingInstruction {
  type: 'pi';
  target: string;
  value: string;
}

export type XmlNode = XmlElement | XmlText | XmlCData | XmlComment | XmlProcessingInstruction;

export interface XmlDocument {
  /** The `<?xml ...?>` declaration content, if present (without the delimiters). */
  declaration: string | null;
  /** Nodes before and after the root (comments, PIs, whitespace). */
  prolog: XmlNode[];
  root: XmlElement;
  epilog: XmlNode[];
}

export function el(name: string, attrs: Record<string, string | number | boolean | undefined | null> = {}, children: XmlNode[] = []): XmlElement {
  const clean: Record<string, string> = {};
  for (const [k, v] of Object.entries(attrs)) {
    if (v === undefined || v === null) continue;
    clean[k] = typeof v === 'string' ? v : String(v);
  }
  return { type: 'element', name, attrs: clean, children };
}

export function text(value: string): XmlText {
  return { type: 'text', value };
}

export function isElement(node: XmlNode | undefined): node is XmlElement {
  return node !== undefined && node.type === 'element';
}

/** Local part of a qualified name (`w:p` → `p`). */
export function localName(name: string): string {
  const i = name.indexOf(':');
  return i < 0 ? name : name.slice(i + 1);
}

export function prefixOf(name: string): string {
  const i = name.indexOf(':');
  return i < 0 ? '' : name.slice(0, i);
}

export function elements(parent: XmlElement, name?: string): XmlElement[] {
  const out: XmlElement[] = [];
  for (const c of parent.children) if (c.type === 'element' && (name === undefined || c.name === name)) out.push(c);
  return out;
}

export function firstChild(parent: XmlElement | undefined, name: string): XmlElement | undefined {
  if (!parent) return undefined;
  for (const c of parent.children) if (c.type === 'element' && c.name === name) return c;
  return undefined;
}

/** Follow a path of child names: `childPath(rPr, 'w:rFonts')`. */
export function childPath(parent: XmlElement | undefined, ...names: string[]): XmlElement | undefined {
  let cur = parent;
  for (const n of names) {
    cur = firstChild(cur, n);
    if (!cur) return undefined;
  }
  return cur;
}

export function attr(node: XmlElement | undefined, name: string): string | undefined {
  return node?.attrs[name];
}

/** Concatenated character data of an element and its descendants. */
export function textContent(node: XmlNode): string {
  switch (node.type) {
    case 'text':
    case 'cdata':
      return node.value;
    case 'element': {
      let s = '';
      for (const c of node.children) s += textContent(c);
      return s;
    }
    default:
      return '';
  }
}

/** Depth-first walk; return `false` from the visitor to skip an element's children. */
export function walk(node: XmlElement, visit: (el: XmlElement, parent: XmlElement | null) => void | false, parent: XmlElement | null = null): void {
  if (visit(node, parent) === false) return;
  for (const c of node.children) if (c.type === 'element') walk(c, visit, node);
}

export function findAll(root: XmlElement, name: string): XmlElement[] {
  const out: XmlElement[] = [];
  walk(root, (e) => {
    if (e.name === name) out.push(e);
  });
  return out;
}

export function cloneElement(node: XmlElement): XmlElement {
  return {
    type: 'element',
    name: node.name,
    attrs: { ...node.attrs },
    children: node.children.map((c) => (c.type === 'element' ? cloneElement(c) : { ...c })),
  };
}

/** Return the existing child with this name or append a new one (keeps schema order only if you insert carefully). */
export function ensureChild(parent: XmlElement, name: string, order?: readonly string[]): XmlElement {
  const existing = firstChild(parent, name);
  if (existing) return existing;
  const created = el(name);
  insertOrdered(parent, created, order);
  return created;
}

/**
 * Insert a child respecting a schema sequence: the child goes before the first existing sibling
 * whose name appears later in `order`. Without an order it is appended.
 */
export function insertOrdered(parent: XmlElement, child: XmlElement, order?: readonly string[]): void {
  if (!order) {
    parent.children.push(child);
    return;
  }
  const rank = order.indexOf(child.name);
  if (rank < 0) {
    parent.children.push(child);
    return;
  }
  for (let i = 0; i < parent.children.length; i++) {
    const c = parent.children[i]!;
    if (c.type !== 'element') continue;
    const r = order.indexOf(c.name);
    if (r > rank) {
      parent.children.splice(i, 0, child);
      return;
    }
  }
  parent.children.push(child);
}

export function removeChildren(parent: XmlElement, name: string): void {
  parent.children = parent.children.filter((c) => !(c.type === 'element' && c.name === name));
}
