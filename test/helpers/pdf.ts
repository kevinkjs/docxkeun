/**
 * A small PDF reader for tests: parses the xref table and objects, checks structural
 * invariants, and extracts text through ToUnicode maps. Not part of the package.
 */
import { inflateSync } from 'node:zlib';

export type PdfObj = number | boolean | null | PdfNameObj | PdfRefObj | PdfStr | PdfObj[] | PdfDictObj;
export interface PdfNameObj {
  name: string;
}
export interface PdfRefObj {
  ref: number;
}
export interface PdfStr {
  bytes: Uint8Array;
}
export interface PdfDictObj {
  dict: Map<string, PdfObj>;
  stream?: Uint8Array;
}

class Lexer {
  pos: number;
  constructor(readonly s: string, pos = 0) {
    this.pos = pos;
  }
  ws(): void {
    for (;;) {
      while (this.pos < this.s.length && /[\s\0]/.test(this.s[this.pos]!)) this.pos++;
      if (this.s[this.pos] === '%') {
        while (this.pos < this.s.length && this.s[this.pos] !== '\n' && this.s[this.pos] !== '\r') this.pos++;
      } else return;
    }
  }
  value(): PdfObj {
    this.ws();
    const c = this.s[this.pos]!;
    if (c === '<' && this.s[this.pos + 1] === '<') {
      this.pos += 2;
      const dict = new Map<string, PdfObj>();
      for (;;) {
        this.ws();
        if (this.s.startsWith('>>', this.pos)) {
          this.pos += 2;
          break;
        }
        const key = this.value() as PdfNameObj;
        dict.set(key.name, this.value());
      }
      return { dict };
    }
    if (c === '[') {
      this.pos++;
      const arr: PdfObj[] = [];
      for (;;) {
        this.ws();
        if (this.s[this.pos] === ']') {
          this.pos++;
          return arr;
        }
        arr.push(this.value());
      }
    }
    if (c === '/') {
      const m = /^\/[^\s/<>[\]()%{}]*/.exec(this.s.slice(this.pos, this.pos + 200))!;
      this.pos += m[0].length;
      return { name: m[0].slice(1).replace(/#([0-9a-f]{2})/gi, (_, h: string) => String.fromCharCode(Number.parseInt(h, 16))) };
    }
    if (c === '<') {
      const end = this.s.indexOf('>', this.pos);
      const hex = this.s.slice(this.pos + 1, end).replace(/\s/g, '');
      this.pos = end + 1;
      const bytes = new Uint8Array(Math.ceil(hex.length / 2));
      for (let i = 0; i < bytes.length; i++) bytes[i] = Number.parseInt(hex.substr(i * 2, 2).padEnd(2, '0'), 16);
      return { bytes };
    }
    if (c === '(') {
      let depth = 0;
      const out: number[] = [];
      this.pos++;
      for (;;) {
        const ch = this.s[this.pos++]!;
        if (ch === '\\') {
          const e = this.s[this.pos++]!;
          const map: Record<string, number> = { n: 10, r: 13, t: 9, b: 8, f: 12 };
          if (e in map) out.push(map[e]!);
          else if (/[0-7]/.test(e)) {
            let o = e;
            while (o.length < 3 && /[0-7]/.test(this.s[this.pos]!)) o += this.s[this.pos++];
            out.push(Number.parseInt(o, 8));
          } else if (e === '\r' || e === '\n') continue;
          else out.push(e.charCodeAt(0));
        } else if (ch === '(') {
          depth++;
          out.push(40);
        } else if (ch === ')') {
          if (depth === 0) break;
          depth--;
          out.push(41);
        } else out.push(ch.charCodeAt(0));
      }
      return { bytes: Uint8Array.from(out) };
    }
    const tok = /^[^\s/<>[\]()%]+/.exec(this.s.slice(this.pos, this.pos + 40))![0];
    if (tok === 'true' || tok === 'false') {
      this.pos += tok.length;
      return tok === 'true';
    }
    if (tok === 'null') {
      this.pos += 4;
      return null;
    }
    // number, or "num gen R"
    const ref = /^(\d+)\s+(\d+)\s+R/.exec(this.s.slice(this.pos, this.pos + 40));
    if (ref) {
      this.pos += ref[0].length;
      return { ref: Number(ref[1]) };
    }
    this.pos += tok.length;
    return Number(tok);
  }
}

export interface ParsedPdf {
  objects: Map<number, PdfObj>;
  trailer: Map<string, PdfObj>;
  version: string;
  text: string;
}

export function latin1(bytes: Uint8Array): string {
  let s = '';
  for (let i = 0; i < bytes.length; i += 8192) s += String.fromCharCode(...bytes.subarray(i, i + 8192));
  return s;
}

/** Parse and validate structure: throws with a description of the first problem found. */
export function parsePdf(bytes: Uint8Array): ParsedPdf {
  const s = latin1(bytes);
  const version = /^%PDF-(\d\.\d)/.exec(s)?.[1];
  if (!version) throw new Error('Missing %PDF header');
  if (!s.trimEnd().endsWith('%%EOF')) throw new Error('Missing %%EOF');
  const startxref = Number(/startxref\s+(\d+)\s+%%EOF\s*$/.exec(s)?.[1]);
  if (!s.startsWith('xref', startxref)) throw new Error('startxref does not point at the xref table');
  const lx = new Lexer(s, startxref + 4);
  lx.ws();
  const header = /^(\d+) (\d+)/.exec(s.slice(lx.pos, lx.pos + 40))!;
  const first = Number(header[1]);
  const count = Number(header[2]);
  let p = lx.pos + header[0].length;
  while (/\s/.test(s[p]!)) p++;
  const offsets = new Map<number, number>();
  for (let i = 0; i < count; i++) {
    const entry = s.slice(p + i * 20, p + i * 20 + 20);
    if (entry.length !== 20 || !/^\d{10} \d{5} [nf](\r\n| \n| \r)$/.test(entry)) throw new Error(`Bad xref entry ${first + i}: ${JSON.stringify(entry)}`);
    if (entry[17] === 'n') offsets.set(first + i, Number(entry.slice(0, 10)));
  }
  const trailerAt = s.indexOf('trailer', p + count * 20);
  const trailer = (new Lexer(s, trailerAt + 7).value() as PdfDictObj).dict;
  if (trailer.get('Size') !== count) throw new Error('Trailer /Size does not match the xref table');

  const objects = new Map<number, PdfObj>();
  for (const [num, offset] of offsets) {
    const head = new RegExp(`^${num} 0 obj\\s`).exec(s.slice(offset, offset + 30));
    if (!head) throw new Error(`xref offset of object ${num} does not point at "${num} 0 obj"`);
    const l = new Lexer(s, offset + head[0].length);
    const value = l.value();
    l.ws();
    if (s.startsWith('stream', l.pos)) {
      let start = l.pos + 6;
      if (s[start] === '\r') start++;
      if (s[start] === '\n') start++;
      const length = (value as PdfDictObj).dict.get('Length');
      if (typeof length !== 'number') throw new Error(`Stream ${num} has no direct /Length`);
      const after = s.slice(start + length, start + length + 12);
      if (!/^\r?\n?endstream/.test(after)) throw new Error(`Stream ${num} /Length ${length} does not end at endstream`);
      (value as PdfDictObj).stream = bytes.subarray(start, start + length);
      l.pos = start + length;
    }
    if (!/^\s*(endstream\s*)?endobj/.test(s.slice(l.pos, l.pos + 30))) throw new Error(`Object ${num} is not terminated by endobj`);
    objects.set(num, value);
  }
  // Every reference must resolve.
  const visit = (v: PdfObj, from: number) => {
    if (v && typeof v === 'object') {
      if ('ref' in v) {
        if (!objects.has(v.ref)) throw new Error(`Object ${from} refers to missing object ${v.ref}`);
      } else if (Array.isArray(v)) v.forEach((x) => visit(x, from));
      else if ('dict' in v) for (const x of v.dict.values()) visit(x, from);
    }
  };
  for (const [num, v] of objects) visit(v, num);
  for (const v of trailer.values()) visit(v, 0);
  return { objects, trailer, version, text: '' };
}

export function resolve(pdf: ParsedPdf, v: PdfObj | undefined): PdfObj | undefined {
  if (v && typeof v === 'object' && 'ref' in v) return pdf.objects.get(v.ref);
  return v;
}

export function dictOf(pdf: ParsedPdf, v: PdfObj | undefined): Map<string, PdfObj> {
  const r = resolve(pdf, v);
  if (!r || typeof r !== 'object' || !('dict' in r)) throw new Error('Expected a dictionary');
  return r.dict;
}

export function streamData(pdf: ParsedPdf, v: PdfObj | undefined): Uint8Array {
  const r = resolve(pdf, v) as PdfDictObj;
  const filter = r.dict.get('Filter') as PdfNameObj | undefined;
  return filter?.name === 'FlateDecode' ? new Uint8Array(inflateSync(r.stream!)) : r.stream!;
}

export function pages(pdf: ParsedPdf): Array<Map<string, PdfObj>> {
  const root = dictOf(pdf, pdf.trailer.get('Root'));
  const tree = dictOf(pdf, root.get('Pages'));
  return (tree.get('Kids') as PdfObj[]).map((k) => dictOf(pdf, k));
}

function toUnicode(pdf: ParsedPdf, font: Map<string, PdfObj>): Map<number, string> {
  const map = new Map<number, string>();
  const cmap = latin1(streamData(pdf, font.get('ToUnicode')));
  for (const m of cmap.matchAll(/<([0-9A-Fa-f]{4})>\s*<([0-9A-Fa-f]+)>/g)) {
    let text = '';
    for (let i = 0; i < m[2]!.length; i += 4) text += String.fromCharCode(Number.parseInt(m[2]!.slice(i, i + 4), 16));
    map.set(Number.parseInt(m[1]!, 16), text);
  }
  return map;
}

/** Text of a page from its content stream, decoded through each font's ToUnicode map. */
export function pageText(pdf: ParsedPdf, page: Map<string, PdfObj>): string {
  const resources = dictOf(pdf, page.get('Resources'));
  const fontDict = dictOf(pdf, resources.get('Font'));
  const maps = new Map<string, Map<number, string>>();
  for (const [res, ref] of fontDict) maps.set(res, toUnicode(pdf, dictOf(pdf, ref)));
  const content = latin1(streamData(pdf, page.get('Contents')));
  let out = '';
  let current: Map<number, string> | undefined;
  for (const m of content.matchAll(/\/(F\d+) [\d.]+ Tf|\[([^\]]*)\] TJ|(ET)/g)) {
    if (m[1]) current = maps.get(m[1]);
    else if (m[3]) out += '\n';
    else if (m[2] && current) {
      for (const h of m[2].matchAll(/<([0-9a-f]*)>/gi)) {
        for (let i = 0; i < h[1]!.length; i += 4) out += current.get(Number.parseInt(h[1]!.slice(i, i + 4), 16)) ?? '�';
      }
    }
  }
  return out;
}
