import type { XmlDocument, XmlNode } from './dom.js';

export function escapeText(s: string): string {
  return s.replace(/[&<>]/g, (c) => (c === '&' ? '&amp;' : c === '<' ? '&lt;' : '&gt;'));
}

export function escapeAttr(s: string): string {
  return s.replace(/[&<"\t\n\r]/g, (c) => {
    switch (c) {
      case '&':
        return '&amp;';
      case '<':
        return '&lt;';
      case '"':
        return '&quot;';
      case '\t':
        return '&#9;';
      case '\n':
        return '&#10;';
      default:
        return '&#13;';
    }
  });
}

function writeNode(node: XmlNode, out: string[]): void {
  switch (node.type) {
    case 'element': {
      out.push('<', node.name);
      for (const name in node.attrs) out.push(' ', name, '="', escapeAttr(node.attrs[name]!), '"');
      if (node.children.length === 0) {
        out.push('/>');
        return;
      }
      out.push('>');
      for (const c of node.children) writeNode(c, out);
      out.push('</', node.name, '>');
      return;
    }
    case 'text':
      out.push(escapeText(node.value));
      return;
    case 'cdata':
      out.push('<![CDATA[', node.value.replace(/]]>/g, ']]]]><![CDATA[>'), ']]>');
      return;
    case 'comment':
      out.push('<!--', node.value.replace(/--/g, '- -'), '-->');
      return;
    case 'pi':
      out.push('<?', node.target, node.value ? ' ' + node.value : '', '?>');
      return;
  }
}

export function serializeNode(node: XmlNode): string {
  const out: string[] = [];
  writeNode(node, out);
  return out.join('');
}

export const DEFAULT_DECLARATION = 'version="1.0" encoding="UTF-8" standalone="yes"';

export function serializeXml(doc: XmlDocument): string {
  const out: string[] = [];
  // Word writes the declaration followed by CRLF; we always write UTF-8.
  const decl = doc.declaration ? doc.declaration.replace(/encoding=("|')[^"']*\1/, 'encoding="UTF-8"') : DEFAULT_DECLARATION;
  out.push('<?xml ', decl, '?>\r\n');
  for (const p of doc.prolog) if (p.type !== 'text') writeNode(p, out);
  writeNode(doc.root, out);
  for (const p of doc.epilog) if (p.type !== 'text') writeNode(p, out);
  return out.join('');
}
