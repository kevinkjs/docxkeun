/**
 * Equations as text in Office's linear format (UnicodeMath, as Word shows an equation switched
 * to "Linear"): x=(-b±√(b^2-4ac))/(2a). Used for search, accessibility and AI tools, so it keeps
 * Word's operators (▒ after n-ary operators, ⁡ function application, ■ matrices) where they exist.
 */
import { elements, firstChild } from '../xml/dom.js';
import type { XmlElement } from '../xml/dom.js';
import { charProp, isPropertyElement, onOff, runText, val } from './omml.js';

const WRAPPERS = new Set(['w:ins', 'w:moveTo', 'w:smartTag', 'w:customXml', 'w:hyperlink', 'w:fldSimple']);

/**
 * An operand that needs no parentheses: one character (with any combining marks), a number, or
 * something already enclosed in one matching pair of brackets.
 */
function isAtomic(s: string): boolean {
  if (/^\p{L}\p{M}*$/u.test(s) || /^\d+(\.\d+)?$/.test(s) || [...s].length === 1) return true;
  return isEnclosed(s);
}

const PAIRS: Record<string, string> = { '(': ')', '[': ']', '{': '}', '〖': '〗', '⟨': '⟩', '|': '|', '‖': '‖' };

function isEnclosed(s: string): boolean {
  const open = s[0];
  const close = open ? PAIRS[open] : undefined;
  if (!open || !close || s[s.length - 1] !== close || s.length < 2) return false;
  if (open === close) return s.indexOf(open, 1) === s.length - 1;
  let depth = 0;
  for (let i = 0; i < s.length; i++) {
    if (s[i] === open) depth++;
    else if (s[i] === close) {
      depth--;
      if (depth === 0 && i < s.length - 1) return false;
    }
  }
  return depth === 0;
}

function wrap(s: string): string {
  return isAtomic(s) ? s : `(${s})`;
}

/** Arguments of n-ary operators and functions use Word's invisible brackets 〖〗. */
function wrapArgument(s: string): string {
  return isAtomic(s) ? s : `〖${s}〗`;
}

function content(el: XmlElement | undefined): string {
  if (!el) return '';
  let s = '';
  for (const c of elements(el)) s += node(c);
  return s;
}

function arg(el: XmlElement, name: string): string {
  return content(firstChild(el, name));
}

function node(el: XmlElement): string {
  switch (el.name) {
    case 'm:r':
    case 'w:r':
      return runText(el);
    case 'm:f': {
      const type = val(firstChild(el, 'm:fPr'), 'm:type') ?? 'bar';
      const op = type === 'skw' ? '⁄' : type === 'noBar' ? '¦' : '/';
      return `${wrap(arg(el, 'm:num'))}${op}${wrap(arg(el, 'm:den'))}`;
    }
    case 'm:sSup':
      return `${wrap(arg(el, 'm:e'))}^${wrap(arg(el, 'm:sup'))}`;
    case 'm:sSub':
      return `${wrap(arg(el, 'm:e'))}_${wrap(arg(el, 'm:sub'))}`;
    case 'm:sSubSup':
      return `${wrap(arg(el, 'm:e'))}_${wrap(arg(el, 'm:sub'))}^${wrap(arg(el, 'm:sup'))}`;
    case 'm:sPre': {
      const sub = arg(el, 'm:sub');
      const sup = arg(el, 'm:sup');
      return `〖${sub ? `_${wrap(sub)}` : ''}${sup ? `^${wrap(sup)}` : ''}〗${wrap(arg(el, 'm:e'))}`;
    }
    case 'm:rad': {
      const body = arg(el, 'm:e');
      const deg = onOff(firstChild(el, 'm:radPr'), 'm:degHide', false) ? '' : arg(el, 'm:deg');
      if (!deg) return `√${wrap(body)}`;
      if (deg === '3') return `∛${wrap(body)}`;
      if (deg === '4') return `∜${wrap(body)}`;
      return `√(${deg}&${body})`;
    }
    case 'm:nary': {
      const pr = firstChild(el, 'm:naryPr');
      const chr = charProp(pr, 'm:chr', '∫') || '∫';
      const sub = onOff(pr, 'm:subHide', false) ? '' : arg(el, 'm:sub');
      const sup = onOff(pr, 'm:supHide', false) ? '' : arg(el, 'm:sup');
      return `${chr}${sub ? `_${wrap(sub)}` : ''}${sup ? `^${wrap(sup)}` : ''}▒${wrapArgument(arg(el, 'm:e'))}`;
    }
    case 'm:d': {
      const pr = firstChild(el, 'm:dPr');
      const beg = charProp(pr, 'm:begChr', '(');
      const end = charProp(pr, 'm:endChr', ')');
      const sep = charProp(pr, 'm:sepChr', '|');
      const inner = elements(el, 'm:e').map(content).join(sep);
      if (!beg && !end) return `〖${inner}〗`;
      return `${beg || '├'}${inner}${end || '┤'}`;
    }
    case 'm:m': {
      const rows = elements(el, 'm:mr').map((r) => elements(r, 'm:e').map(content).join('&'));
      return `■(${rows.join('@')})`;
    }
    case 'm:eqArr':
      return `█(${elements(el, 'm:e').map(content).join('@')})`;
    case 'm:acc': {
      const chr = charProp(firstChild(el, 'm:accPr'), 'm:chr', '̂') || '̂';
      return `${wrap(arg(el, 'm:e'))}${chr}`;
    }
    case 'm:bar': {
      const top = val(firstChild(el, 'm:barPr'), 'm:pos') === 'top';
      return top ? `${wrap(arg(el, 'm:e'))}̅` : `▁${wrap(arg(el, 'm:e'))}`;
    }
    case 'm:func':
      return `${arg(el, 'm:fName')}⁡${wrapArgument(arg(el, 'm:e'))}`;
    case 'm:limLow':
      return `${arg(el, 'm:e')}┬${wrap(arg(el, 'm:lim'))}`;
    case 'm:limUpp':
      return `${arg(el, 'm:e')}┴${wrap(arg(el, 'm:lim'))}`;
    case 'm:groupChr': {
      const chr = charProp(firstChild(el, 'm:groupChrPr'), 'm:chr', '⏟') || '⏟';
      return `${chr}${wrap(arg(el, 'm:e'))}`;
    }
    case 'm:box':
      return `□(${arg(el, 'm:e')})`;
    case 'm:borderBox':
      return `▭(${arg(el, 'm:e')})`;
    case 'm:phant':
      return `⟡(${arg(el, 'm:e')})`;
    case 'm:oMathPara':
      return elements(el, 'm:oMath').map(content).join('\n');
    case 'w:sdt':
      return content(firstChild(el, 'w:sdtContent'));
    default:
      if (isPropertyElement(el.name)) return '';
      if (el.name.startsWith('m:') || WRAPPERS.has(el.name)) return content(el);
      return '';
  }
}

/** The equation (m:oMath or m:oMathPara) in Office's linear format; equations of a paragraph on separate lines. */
export function mathToText(math: XmlElement): string {
  if (math.name === 'm:oMathPara') return node(math);
  return content(math);
}
