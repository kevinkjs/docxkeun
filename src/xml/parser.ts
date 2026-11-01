/**
 * Non-validating XML 1.0 parser. It never loads external entities or DTDs, so it is safe on
 * untrusted input; unknown entity references are kept as literal text.
 */
import type { XmlDocument, XmlElement, XmlNode } from './dom.js';
import { canonicalizeNamespaces } from './ns.js';

export class XmlParseError extends Error {
  override name = 'XmlParseError';
  constructor(message: string, readonly position: number) {
    super(`${message} at offset ${position}`);
  }
}

const NAMED_ENTITIES: Record<string, string> = { lt: '<', gt: '>', amp: '&', quot: '"', apos: "'" };

export function decodeEntities(s: string): string {
  if (s.indexOf('&') < 0) return s;
  return s.replace(/&(#x[0-9a-fA-F]+|#[0-9]+|[A-Za-z][\w.-]*);/g, (match, body: string) => {
    if (body[0] === '#') {
      const code = body[1] === 'x' ? Number.parseInt(body.slice(2), 16) : Number.parseInt(body.slice(1), 10);
      if (!Number.isFinite(code) || code < 0 || code > 0x10ffff) return match;
      return String.fromCodePoint(code);
    }
    return NAMED_ENTITIES[body] ?? match;
  });
}

function isNameChar(c: number): boolean {
  return (
    (c >= 97 && c <= 122) || // a-z
    (c >= 65 && c <= 90) || // A-Z
    (c >= 48 && c <= 57) || // 0-9
    c === 58 || // :
    c === 95 || // _
    c === 45 || // -
    c === 46 || // .
    c >= 0x80
  );
}

function isSpace(c: number): boolean {
  return c === 32 || c === 9 || c === 10 || c === 13;
}

export interface ParseXmlOptions {
  /** Rewrite non-standard prefixes of known OOXML namespaces to their usual ones (default true). */
  canonicalize?: boolean;
}

export function parseXml(source: string, options: ParseXmlOptions = {}): XmlDocument {
  const s = source;
  const n = s.length;
  let i = 0;
  let declaration: string | null = null;
  const prolog: XmlNode[] = [];
  const epilog: XmlNode[] = [];
  let root: XmlElement | null = null;
  const stack: XmlElement[] = [];

  if (s.charCodeAt(0) === 0xfeff) i = 1;

  const push = (node: XmlNode): void => {
    const parent = stack[stack.length - 1];
    if (parent) parent.children.push(node);
    else if (root) epilog.push(node);
    else prolog.push(node);
  };

  const attach = (element: XmlElement): void => {
    if (stack.length > 0) push(element);
    else if (!root) root = element;
    else throw new XmlParseError('More than one root element', i);
  };

  const readName = (): string => {
    const start = i;
    while (i < n && isNameChar(s.charCodeAt(i))) i++;
    if (i === start) throw new XmlParseError('Expected a name', i);
    return s.slice(start, i);
  };

  const skipSpace = (): void => {
    while (i < n && isSpace(s.charCodeAt(i))) i++;
  };

  while (i < n) {
    const lt = s.indexOf('<', i);
    if (lt < 0 || lt > i) {
      const end = lt < 0 ? n : lt;
      const raw = s.slice(i, end);
      if (stack.length > 0) push({ type: 'text', value: decodeEntities(raw) });
      else if (raw.trim() !== '') throw new XmlParseError('Text outside the root element', i);
      else push({ type: 'text', value: raw });
      i = end;
      if (lt < 0) break;
    }
    // s[i] === '<'
    const next = s.charCodeAt(i + 1);
    if (next === 63 /* ? */) {
      const end = s.indexOf('?>', i + 2);
      if (end < 0) throw new XmlParseError('Unterminated processing instruction', i);
      const body = s.slice(i + 2, end);
      const sp = body.search(/\s/);
      const target = sp < 0 ? body : body.slice(0, sp);
      const value = sp < 0 ? '' : body.slice(sp + 1);
      if (target === 'xml' && !root && stack.length === 0) declaration = value.trim();
      else push({ type: 'pi', target, value });
      i = end + 2;
    } else if (next === 33 /* ! */) {
      if (s.startsWith('<!--', i)) {
        const end = s.indexOf('-->', i + 4);
        if (end < 0) throw new XmlParseError('Unterminated comment', i);
        push({ type: 'comment', value: s.slice(i + 4, end) });
        i = end + 3;
      } else if (s.startsWith('<![CDATA[', i)) {
        const end = s.indexOf(']]>', i + 9);
        if (end < 0) throw new XmlParseError('Unterminated CDATA section', i);
        push({ type: 'cdata', value: s.slice(i + 9, end) });
        i = end + 3;
      } else if (s.startsWith('<!DOCTYPE', i)) {
        // Skip the doctype, including an internal subset; entities declared there are not expanded.
        let depth = 0;
        let j = i + 9;
        for (; j < n; j++) {
          const c = s[j];
          if (c === '[') depth++;
          else if (c === ']') depth--;
          else if (c === '>' && depth <= 0) break;
        }
        i = j + 1;
      } else {
        throw new XmlParseError('Unsupported markup declaration', i);
      }
    } else if (next === 47 /* / */) {
      i += 2;
      const name = readName();
      skipSpace();
      if (s[i] !== '>') throw new XmlParseError('Expected >', i);
      i++;
      const open = stack.pop();
      if (!open || open.name !== name) {
        throw new XmlParseError(`Mismatched closing tag </${name}>${open ? ` for <${open.name}>` : ''}`, i);
      }
    } else {
      i++;
      const name = readName();
      const element: XmlElement = { type: 'element', name, attrs: {}, children: [] };
      for (;;) {
        skipSpace();
        const c = s.charCodeAt(i);
        if (c === 62 /* > */) {
          i++;
          attach(element);
          stack.push(element);
          break;
        }
        if (c === 47 /* / */) {
          if (s[i + 1] !== '>') throw new XmlParseError('Expected />', i);
          i += 2;
          attach(element);
          break;
        }
        if (i >= n) throw new XmlParseError(`Unterminated tag <${name}>`, i);
        const attrName = readName();
        skipSpace();
        if (s[i] !== '=') throw new XmlParseError(`Expected = after attribute ${attrName}`, i);
        i++;
        skipSpace();
        const quote = s[i];
        if (quote !== '"' && quote !== "'") throw new XmlParseError('Expected quoted attribute value', i);
        const end = s.indexOf(quote, i + 1);
        if (end < 0) throw new XmlParseError('Unterminated attribute value', i);
        // Attribute-value normalisation: literal whitespace characters become spaces.
        const rawValue = s.slice(i + 1, end).replace(/[\t\n\r]/g, ' ');
        element.attrs[attrName] = decodeEntities(rawValue);
        i = end + 1;
      }
    }
  }

  if (stack.length > 0) throw new XmlParseError(`Unclosed element <${stack[stack.length - 1]!.name}>`, n);
  if (!root) throw new XmlParseError('No root element', 0);
  const doc: XmlDocument = { declaration, prolog, root, epilog };
  if (options.canonicalize !== false) canonicalizeNamespaces(doc.root);
  return doc;
}
