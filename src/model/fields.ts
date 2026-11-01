/** Field instruction parsing (`HYPERLINK "url" \l "bm"`, `PAGE`, `PAGEREF _Toc1 \h`, ...). */

export interface FieldInstruction {
  type: string;
  args: string[];
  switches: Map<string, string | true>;
  raw: string;
}

/** Switches that take an argument; everything else is a flag. */
const ARG_SWITCHES = new Set(['\\l', '\\o', '\\t', '\\*', '\\#', '\\@', '\\f', '\\s', '\\b', '\\c', '\\d', '\\p', '\\m']);

function tokenize(instr: string): string[] {
  const tokens: string[] = [];
  let i = 0;
  while (i < instr.length) {
    const c = instr[i]!;
    if (/\s/.test(c)) {
      i++;
      continue;
    }
    if (c === '"' || c === '“' || c === '”') {
      let j = i + 1;
      let s = '';
      while (j < instr.length && !['"', '“', '”'].includes(instr[j]!)) {
        if (instr[j] === '\\' && (instr[j + 1] === '"' || instr[j + 1] === '\\')) {
          s += instr[j + 1];
          j += 2;
        } else s += instr[j++];
      }
      tokens.push(s);
      i = j + 1;
      continue;
    }
    let j = i;
    while (j < instr.length && !/\s/.test(instr[j]!)) j++;
    tokens.push(instr.slice(i, j));
    i = j;
  }
  return tokens;
}

export function parseFieldInstruction(raw: string): FieldInstruction {
  const tokens = tokenize(raw.trim());
  const type = (tokens.shift() ?? '').toUpperCase();
  const args: string[] = [];
  const switches = new Map<string, string | true>();
  for (let i = 0; i < tokens.length; i++) {
    const t = tokens[i]!;
    if (t.startsWith('\\') && t.length >= 2 && !/^\\\\/.test(t)) {
      const name = t.slice(0, 2).toLowerCase();
      const inline = t.slice(2);
      if (inline) switches.set(name, inline);
      else if (ARG_SWITCHES.has(name) && i + 1 < tokens.length) switches.set(name, tokens[++i]!);
      else switches.set(name, true);
    } else {
      args.push(t);
    }
  }
  return { type, args, switches, raw };
}

export type DynamicFieldType = 'PAGE' | 'NUMPAGES' | 'SECTIONPAGES';

export function isDynamicField(type: string): type is DynamicFieldType {
  return type === 'PAGE' || type === 'NUMPAGES' || type === 'SECTIONPAGES';
}

export interface LinkTarget {
  url?: string;
  anchor?: string;
  tooltip?: string;
}

/** Link target expressed by a field, if the field is a link. */
export function fieldLinkTarget(f: FieldInstruction): LinkTarget | undefined {
  if (f.type === 'HYPERLINK') {
    const anchor = f.switches.get('\\l');
    const tooltip = f.switches.get('\\o');
    return {
      url: f.args[0] || undefined,
      anchor: typeof anchor === 'string' ? anchor : undefined,
      tooltip: typeof tooltip === 'string' ? tooltip : undefined,
    };
  }
  if ((f.type === 'PAGEREF' || f.type === 'REF' || f.type === 'NOTEREF') && f.switches.has('\\h') && f.args[0]) {
    return { anchor: f.args[0] };
  }
  return undefined;
}
