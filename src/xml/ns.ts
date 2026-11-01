/**
 * Known OOXML namespaces and their customary prefixes. The model code matches names like `w:p`,
 * so parts written with unusual prefixes (or a default namespace) are rewritten to these.
 */
import type { XmlElement } from './dom.js';

export const NS = {
  w: 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
  r: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
  wp: 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing',
  a: 'http://schemas.openxmlformats.org/drawingml/2006/main',
  pic: 'http://schemas.openxmlformats.org/drawingml/2006/picture',
  m: 'http://schemas.openxmlformats.org/officeDocument/2006/math',
  mc: 'http://schemas.openxmlformats.org/markup-compatibility/2006',
  v: 'urn:schemas-microsoft-com:vml',
  o: 'urn:schemas-microsoft-com:office:office',
  w10: 'urn:schemas-microsoft-com:office:word',
  w14: 'http://schemas.microsoft.com/office/word/2010/wordml',
  w15: 'http://schemas.microsoft.com/office/word/2012/wordml',
  w16cid: 'http://schemas.microsoft.com/office/word/2016/wordml/cid',
  w16cex: 'http://schemas.microsoft.com/office/word/2018/wordml/cex',
  w16se: 'http://schemas.microsoft.com/office/word/2015/wordml/symex',
  wp14: 'http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing',
  wps: 'http://schemas.microsoft.com/office/word/2010/wordprocessingShape',
  wpg: 'http://schemas.microsoft.com/office/word/2010/wordprocessingGroup',
  wpc: 'http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas',
  cp: 'http://schemas.openxmlformats.org/package/2006/metadata/core-properties',
  dc: 'http://purl.org/dc/elements/1.1/',
  dcterms: 'http://purl.org/dc/terms/',
  dcmitype: 'http://purl.org/dc/dcmitype/',
  xsi: 'http://www.w3.org/2001/XMLSchema-instance',
  ep: 'http://schemas.openxmlformats.org/officeDocument/2006/extended-properties',
  vt: 'http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes',
  pkgRel: 'http://schemas.openxmlformats.org/package/2006/relationships',
  contentTypes: 'http://schemas.openxmlformats.org/package/2006/content-types',
} as const;

/** ISO/IEC 29500 Strict namespaces mapped onto the same prefixes as Transitional. */
const STRICT: Record<string, string> = {
  'http://purl.oclc.org/ooxml/wordprocessingml/main': 'w',
  'http://purl.oclc.org/ooxml/officeDocument/relationships': 'r',
  'http://purl.oclc.org/ooxml/drawingml/wordprocessingDrawing': 'wp',
  'http://purl.oclc.org/ooxml/drawingml/main': 'a',
  'http://purl.oclc.org/ooxml/drawingml/picture': 'pic',
  'http://purl.oclc.org/ooxml/officeDocument/math': 'm',
};

/** Namespaces whose elements are conventionally unprefixed (package parts); never rewritten. */
const DEFAULT_ONLY = new Set<string>([NS.pkgRel, NS.contentTypes, NS.ep]);

const CANONICAL_PREFIX: Map<string, string> = (() => {
  const m = new Map<string, string>();
  for (const [prefix, uri] of Object.entries(NS)) if (!DEFAULT_ONLY.has(uri)) m.set(uri, prefix);
  for (const [uri, prefix] of Object.entries(STRICT)) m.set(uri, prefix);
  return m;
})();

export function isStrictNamespace(uri: string): boolean {
  return uri in STRICT;
}

const PREFIX_LIST_ATTRS = ['mc:Ignorable', 'mc:ProcessContent', 'mc:MustUnderstand', 'Requires'];

/**
 * Rewrite prefixes in place. Returns true when anything changed. The fast path (all prefixes
 * already canonical, the normal case for Word output) leaves the tree untouched.
 */
export function canonicalizeNamespaces(root: XmlElement): boolean {
  let changed = false;

  const visit = (node: XmlElement, renames: Map<string, string>): void => {
    let scope = renames;
    // Find declarations on this element that need a rename.
    for (const key of Object.keys(node.attrs)) {
      let declared: string | null = null;
      if (key === 'xmlns') declared = '';
      else if (key.startsWith('xmlns:')) declared = key.slice(6);
      if (declared === null) continue;
      const uri = node.attrs[key]!;
      const canonical = CANONICAL_PREFIX.get(uri);
      if (canonical === undefined || canonical === declared) {
        if (scope.has(declared)) {
          if (scope === renames) scope = new Map(renames);
          scope.delete(declared);
        }
        continue;
      }
      if (scope === renames) scope = new Map(renames);
      scope.set(declared, canonical);
    }

    if (scope.size > 0) {
      const renamePrefix = (qname: string, isAttr: boolean): string => {
        const i = qname.indexOf(':');
        const prefix = i < 0 ? '' : qname.slice(0, i);
        if (prefix === 'xmlns' || qname === 'xmlns') return qname;
        if (i < 0 && isAttr) return qname; // unprefixed attributes have no namespace
        const target = scope.get(prefix);
        if (target === undefined) return qname;
        return `${target}:${i < 0 ? qname : qname.slice(i + 1)}`;
      };

      const newName = renamePrefix(node.name, false);
      if (newName !== node.name) {
        node.name = newName;
        changed = true;
      }

      const attrs: Record<string, string> = {};
      for (const [key, value] of Object.entries(node.attrs)) {
        let outKey = key;
        let outValue = value;
        if (key === 'xmlns' || key.startsWith('xmlns:')) {
          const declared = key === 'xmlns' ? '' : key.slice(6);
          const target = scope.get(declared);
          if (target !== undefined && CANONICAL_PREFIX.get(value) === target) {
            outKey = `xmlns:${target}`;
            changed = true;
          }
        } else {
          outKey = renamePrefix(key, true);
          if (PREFIX_LIST_ATTRS.includes(outKey)) {
            outValue = value
              .split(/\s+/)
              .map((p) => scope.get(p) ?? p)
              .join(' ');
          }
        }
        if (outKey !== key || outValue !== value) changed = true;
        attrs[outKey] = outValue;
      }
      node.attrs = attrs;
    }

    for (const c of node.children) if (c.type === 'element') visit(c, scope);
  };

  visit(root, new Map());
  return changed;
}
