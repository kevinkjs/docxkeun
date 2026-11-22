/**
 * Excel number format codes, enough for chart labels: General, digit placeholders (0 # ?), thousands
 * separators and scaling commas, percent, scientific notation, quoted and escaped literals, currency
 * tags ([$€-407]), colour tags (ignored), conditions, up to four sections, and common date/time codes
 * (date axes carry serial numbers). Output uses "." and "," as Excel does for the en-US locale.
 */

/** Format a number with an Excel format code ("General" when absent). */
export function formatNumber(value: number, code: string | undefined): string {
  if (!Number.isFinite(value)) return String(value);
  const source = code === undefined || code.trim() === '' ? 'General' : code;
  const sections = splitSections(source);
  const { section, signed } = pickSection(sections, value);
  if (section === undefined) return '';
  const tokens = tokenize(section);
  // Only a section that "owns" negative numbers prints the minus sign itself.
  const minus = signed && value < 0 ? '-' : '';
  const abs = Math.abs(value);
  if (tokens.some((t) => t.kind === 'general')) return minus + renderGeneral(tokens, abs);
  if (isDateFormat(tokens)) return minus + formatDate(abs, tokens);
  return renderNumber(tokens, abs, minus);
}

/** Is this format code a date/time format (so values are serial dates)? */
export function isDateCode(code: string | undefined): boolean {
  if (!code) return false;
  const first = splitSections(code)[0] ?? '';
  return isDateFormat(tokenize(first));
}

/** Excel's General format: up to 10 significant digits, scientific notation for very large or small values. */
export function formatGeneral(value: number): string {
  if (value === 0 || !Number.isFinite(value)) return Number.isFinite(value) ? '0' : String(value);
  const abs = Math.abs(value);
  if (abs >= 1e11 || abs < 1e-9) {
    const [mantissa, exponent] = value.toExponential(5).split('e') as [string, string];
    const m = mantissa.includes('.') ? mantissa.replace(/0+$/, '').replace(/\.$/, '') : mantissa;
    const e = Number(exponent);
    return `${m}E${e < 0 ? '-' : '+'}${String(Math.abs(e)).padStart(2, '0')}`;
  }
  const magnitude = Math.floor(Math.log10(abs)) + 1;
  const decimals = Math.max(0, Math.min(20, 10 - magnitude));
  let s = value.toFixed(decimals);
  if (s.includes('.')) s = s.replace(/0+$/, '').replace(/\.$/, '');
  return s === '-0' ? '0' : s;
}

function splitSections(code: string): string[] {
  const out: string[] = [];
  let cur = '';
  for (let i = 0; i < code.length; i++) {
    const ch = code[i]!;
    if (ch === '"') {
      const end = code.indexOf('"', i + 1);
      const stop = end < 0 ? code.length : end;
      cur += code.slice(i, stop + 1);
      i = stop;
    } else if (ch === '\\' || ch === '_' || ch === '*') {
      cur += code.slice(i, i + 2);
      i++;
    } else if (ch === '[') {
      const end = code.indexOf(']', i);
      const stop = end < 0 ? code.length : end;
      cur += code.slice(i, stop + 1);
      i = stop;
    } else if (ch === ';') {
      out.push(cur);
      cur = '';
    } else cur += ch;
  }
  out.push(cur);
  return out;
}

const CONDITION = /\[(<=|>=|<>|<|>|=)(-?\d+(?:\.\d+)?)\]/;

function testCondition(section: string, value: number): boolean | undefined {
  const m = CONDITION.exec(section);
  if (!m) return undefined;
  const n = Number(m[2]);
  switch (m[1]) {
    case '<':
      return value < n;
    case '<=':
      return value <= n;
    case '>':
      return value > n;
    case '>=':
      return value >= n;
    case '=':
      return value === n;
    default:
      return value !== n;
  }
}

/**
 * Choose the section for a value. `signed` means the section shows the sign itself (the single-section
 * case and conditional sections do; the negative section of a multi-section code gets the absolute value).
 */
function pickSection(sections: string[], value: number): { section: string | undefined; signed: boolean } {
  const conditional = sections.some((s) => CONDITION.test(s));
  if (conditional) {
    for (const s of sections.slice(0, 2)) if (testCondition(s, value)) return { section: s, signed: false };
    const rest = sections[2] ?? sections[sections.length - 1];
    return { section: rest, signed: true };
  }
  if (sections.length === 1) return { section: sections[0], signed: true };
  if (value > 0 || (value === 0 && sections.length === 2)) return { section: sections[0], signed: false };
  if (value < 0) return { section: sections[1], signed: false };
  return { section: sections[2], signed: false };
}

type Token =
  | { kind: 'literal'; text: string }
  | { kind: 'digit'; ch: '0' | '#' | '?' }
  | { kind: 'point' }
  | { kind: 'comma' }
  | { kind: 'percent' }
  | { kind: 'exp'; sign: '+' | '-' }
  | { kind: 'text' }
  | { kind: 'general' }
  | { kind: 'date'; code: string };

const DATE_LETTERS = /[ymdhsYMDHS]/;

function tokenize(section: string): Token[] {
  const tokens: Token[] = [];
  const lit = (text: string): void => {
    const last = tokens[tokens.length - 1];
    if (last?.kind === 'literal') last.text += text;
    else tokens.push({ kind: 'literal', text });
  };
  for (let i = 0; i < section.length; i++) {
    const ch = section[i]!;
    const rest = section.slice(i);
    if (/^general/i.test(rest)) {
      tokens.push({ kind: 'general' });
      i += 6;
    } else if (ch === '"') {
      const end = section.indexOf('"', i + 1);
      const stop = end < 0 ? section.length : end;
      lit(section.slice(i + 1, stop));
      i = stop;
    } else if (ch === '\\') {
      if (i + 1 < section.length) lit(section[i + 1]!);
      i++;
    } else if (ch === '_') {
      // Space as wide as the next character: a plain space is close enough for labels.
      lit(' ');
      i++;
    } else if (ch === '*') {
      i++; // repeat-to-fill has no meaning outside a cell
    } else if (ch === '[') {
      const end = section.indexOf(']', i);
      const stop = end < 0 ? section.length : end;
      const inner = section.slice(i + 1, stop);
      if (inner.startsWith('$')) {
        // [$€-407]: currency symbol before the locale id.
        const dash = inner.indexOf('-');
        lit(inner.slice(1, dash < 0 ? undefined : dash));
      }
      i = stop;
    } else if (ch === '0' || ch === '#' || ch === '?') {
      tokens.push({ kind: 'digit', ch });
    } else if (ch === '.') {
      tokens.push({ kind: 'point' });
    } else if (ch === ',') {
      tokens.push({ kind: 'comma' });
    } else if (ch === '%') {
      tokens.push({ kind: 'percent' });
    } else if ((ch === 'E' || ch === 'e') && (section[i + 1] === '+' || section[i + 1] === '-')) {
      tokens.push({ kind: 'exp', sign: section[i + 1] as '+' | '-' });
      i++;
    } else if (ch === '@') {
      tokens.push({ kind: 'text' });
    } else if (/^(AM\/PM|am\/pm|A\/P|a\/p)/.test(rest)) {
      const m = /^(AM\/PM|am\/pm|A\/P|a\/p)/.exec(rest)![0];
      tokens.push({ kind: 'date', code: m });
      i += m.length - 1;
    } else if (DATE_LETTERS.test(ch)) {
      let j = i;
      while (j < section.length && section[j]!.toLowerCase() === ch.toLowerCase()) j++;
      tokens.push({ kind: 'date', code: section.slice(i, j).toLowerCase() });
      i = j - 1;
    } else {
      lit(ch);
    }
  }
  return tokens;
}

function isDateFormat(tokens: Token[]): boolean {
  return tokens.some((t) => t.kind === 'date') && !tokens.some((t) => t.kind === 'digit' && t.ch !== '0');
}

function renderGeneral(tokens: Token[], value: number): string {
  let out = '';
  for (const t of tokens) {
    if (t.kind === 'general') out += formatGeneral(value);
    else if (t.kind === 'literal') out += t.text;
  }
  return out;
}

/** Round half away from zero to `decimals` places and return the digit strings (no sign). */
function splitDigits(value: number, decimals: number): { int: string; frac: string } {
  const d = Math.min(decimals, 15);
  // toPrecision(15) first so 1.005 → 100.5 (not 100.4999…) before rounding, like Excel's decimal arithmetic.
  const scaled = Number((value * 10 ** d).toPrecision(15));
  if (scaled >= 1e21) {
    const [int, frac = ''] = value.toFixed(d).split('.') as [string, string?];
    return { int, frac };
  }
  const s = String(Math.round(scaled)).padStart(d + 1, '0');
  return d === 0 ? { int: s, frac: '' } : { int: s.slice(0, -d), frac: s.slice(-d) };
}

function renderNumber(tokens: Token[], abs: number, minus: string): string {
  // The number part runs from the first to the last digit placeholder (with the point and commas in it).
  const digitIdx = tokens.map((t, i) => (t.kind === 'digit' ? i : -1)).filter((i) => i >= 0);
  if (digitIdx.length === 0) {
    // Only literals (e.g. "-" or text): Excel shows the literal text, with percent applied to nothing.
    return minus + tokens.map((t) => (t.kind === 'literal' ? t.text : t.kind === 'percent' ? '%' : '')).join('');
  }
  const first = digitIdx[0]!;
  const last = digitIdx[digitIdx.length - 1]!;
  const expAt = tokens.findIndex((t) => t.kind === 'exp');
  const mantEnd = expAt >= 0 ? expAt : last + 1;
  const pointAt = tokens.findIndex((t, i) => t.kind === 'point' && i >= first - 1 && i < mantEnd);

  let value = abs;
  const percents = tokens.filter((t) => t.kind === 'percent').length;
  value *= 100 ** percents;

  // Commas: between integer placeholders = grouping; right after the last placeholder (or before
  // the point) = ÷1000 each.
  const intEnd = pointAt >= 0 ? pointAt : mantEnd;
  let grouping = false;
  let scaleCommas = 0;
  {
    let lastIntDigit = -1;
    for (let i = first; i < intEnd; i++) if (tokens[i]!.kind === 'digit') lastIntDigit = i;
    for (let i = first; i < intEnd; i++) if (tokens[i]!.kind === 'comma' && i < lastIntDigit) grouping = true;
    if (pointAt >= 0) for (let i = pointAt - 1; i > lastIntDigit && tokens[i]!.kind === 'comma'; i--) scaleCommas++;
    for (let i = last + 1; i < tokens.length && tokens[i]!.kind === 'comma'; i++) scaleCommas++;
  }
  value /= 1000 ** scaleCommas;

  const intSlots = tokens.slice(first, intEnd).filter((t): t is Extract<Token, { kind: 'digit' }> => t.kind === 'digit');
  const fracSlots = pointAt >= 0 ? tokens.slice(pointAt + 1, mantEnd).filter((t): t is Extract<Token, { kind: 'digit' }> => t.kind === 'digit') : [];

  let exponent = 0;
  if (expAt >= 0 && value !== 0) {
    // Engineering-style formats (##0.0E+0) group the exponent by the integer placeholder count.
    const step = intSlots.length > 1 && intSlots.some((s) => s.ch === '#') ? intSlots.length : 1;
    exponent = Math.floor(Math.log10(value) / step) * step;
    value /= 10 ** exponent;
    if (Number(value.toFixed(fracSlots.length)) >= 10 ** step) {
      value /= 10 ** step;
      exponent += step;
    }
  }

  const { int: rawInt, frac: rawFrac } = splitDigits(value, fracSlots.length);
  // Integer digits: pad to the number of '0' slots; drop a lone 0 when no slot requires it.
  const required = intSlots.filter((s) => s.ch === '0').length;
  let intDigits = rawInt === '0' ? '' : rawInt;
  if (intDigits.length < required) intDigits = intDigits.padStart(required, '0');

  // Fraction digits: '0' keeps trailing zeros, '#' drops them, '?' turns them into spaces.
  const fracChars = rawFrac.split('');
  for (let i = fracSlots.length - 1; i >= 0; i--) {
    if (fracChars[i] !== '0') break;
    const slot = fracSlots[i]!.ch;
    if (slot === '0') break;
    fracChars[i] = slot === '?' ? ' ' : '';
  }
  const fracText = fracChars.join('');

  // Assign integer digits to slots right to left; the first slot takes any overflow.
  const intText: string[] = [];
  const slotCount = intSlots.length;
  const leftover = intDigits.length - slotCount;
  {
    let d = intDigits.length - 1;
    const assigned: string[] = Array(slotCount).fill('');
    for (let s = slotCount - 1; s >= 0; s--) {
      if (d >= 0) assigned[s] = intDigits[d--]!;
      else assigned[s] = intSlots[s]!.ch === '?' ? ' ' : '';
    }
    if (leftover > 0) assigned[0] = intDigits.slice(0, leftover + 1);
    intText.push(...assigned);
  }
  const groupedInt = grouping ? group(intText.join('').replace(/ /g, '')) : undefined;

  let out = '';
  let slot = 0;
  let fracSlot = 0;
  let wroteInt = false;
  for (let i = 0; i < tokens.length; i++) {
    const t = tokens[i]!;
    switch (t.kind) {
      case 'literal':
        out += t.text;
        break;
      case 'percent':
        out += '%';
        break;
      case 'digit':
        if (expAt >= 0 && i > expAt) break; // exponent digits are written with the E token
        if (pointAt < 0 || i < pointAt) {
          if (groupedInt !== undefined) {
            if (!wroteInt) out += groupedInt;
            wroteInt = true;
          } else out += intText[slot] ?? '';
          slot++;
        } else {
          out += fracText[fracSlot] ?? '';
          fracSlot++;
        }
        break;
      case 'point':
        // A format like ".00" has no integer slots; Excel still prints the integer digits.
        if (i === pointAt && slotCount === 0) out += intDigits;
        out += '.';
        break;
      case 'comma':
        // Grouping and scaling commas are not printed; a comma outside the number is a literal.
        if (i < first || (i > last && !tokens.slice(last + 1, i).every((x) => x.kind === 'comma'))) out += ',';
        break;
      case 'exp': {
        const expDigits = tokens.slice(i + 1).filter((x) => x.kind === 'digit').length;
        const e = Math.abs(exponent);
        out += `E${exponent < 0 ? '-' : t.sign === '+' ? '+' : ''}${String(e).padStart(Math.max(1, expDigits), '0')}`;
        break;
      }
      case 'text':
      case 'general':
      case 'date':
        break;
    }
  }
  // A value that rounds to zero shows no minus sign.
  const isZero = !/[1-9]/.test(rawInt + rawFrac);
  return (isZero ? '' : minus) + out;
}

function group(digits: string): string {
  let out = '';
  for (let i = 0; i < digits.length; i++) {
    if (i > 0 && (digits.length - i) % 3 === 0) out += ',';
    out += digits[i];
  }
  return out;
}

const MONTHS = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'];
const DAYS = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'];

/** Excel serial date (1900 system, with its phantom 29 Feb 1900) to a UTC Date. */
function serialToDate(serial: number): Date {
  const days = serial < 61 ? serial + 1 : serial;
  return new Date(Date.UTC(1899, 11, 30) + Math.round(days * 86400000));
}

function formatDate(serial: number, tokens: Token[]): string {
  const d = serialToDate(serial);
  const hasAmPm = tokens.some((t) => t.kind === 'date' && t.code.includes('/'));
  const parts = tokens.filter((t) => t.kind === 'date' || t.kind === 'literal' || t.kind === 'digit' || t.kind === 'point');
  let out = '';
  for (let i = 0; i < parts.length; i++) {
    const t = parts[i]!;
    if (t.kind === 'literal') {
      out += t.text;
      continue;
    }
    if (t.kind === 'point') {
      out += '.';
      continue;
    }
    if (t.kind === 'digit') {
      // Fractional seconds after "ss.": tenths, hundredths…
      out += '0';
      continue;
    }
    if (t.kind !== 'date') continue;
    const code = t.code;
    const prev = parts.slice(0, i).reverse().find((x) => x.kind === 'date');
    const next = parts.slice(i + 1).find((x) => x.kind === 'date');
    switch (code[0]) {
      case 'y':
        out += code.length <= 2 ? String(d.getUTCFullYear() % 100).padStart(2, '0') : String(d.getUTCFullYear());
        break;
      case 'm': {
        // "m" right after hours or right before seconds means minutes.
        const minutes = (prev?.kind === 'date' && prev.code[0] === 'h') || (next?.kind === 'date' && next.code[0] === 's');
        if (minutes && code.length <= 2) {
          out += code.length === 2 ? String(d.getUTCMinutes()).padStart(2, '0') : String(d.getUTCMinutes());
        } else {
          const month = d.getUTCMonth();
          const name = MONTHS[month]!;
          out += code.length === 1 ? String(month + 1) : code.length === 2 ? String(month + 1).padStart(2, '0') : code.length === 3 ? name.slice(0, 3) : code.length === 5 ? name[0] : name;
        }
        break;
      }
      case 'd': {
        const day = d.getUTCDate();
        const name = DAYS[d.getUTCDay()]!;
        out += code.length === 1 ? String(day) : code.length === 2 ? String(day).padStart(2, '0') : code.length === 3 ? name.slice(0, 3) : name;
        break;
      }
      case 'h': {
        let h = d.getUTCHours();
        if (hasAmPm) h = h % 12 === 0 ? 12 : h % 12;
        out += code.length >= 2 ? String(h).padStart(2, '0') : String(h);
        break;
      }
      case 's':
        out += code.length >= 2 ? String(d.getUTCSeconds()).padStart(2, '0') : String(d.getUTCSeconds());
        break;
      default: {
        // AM/PM and A/P markers.
        const pm = d.getUTCHours() >= 12;
        if (code.toLowerCase() === 'am/pm') out += pm ? 'PM' : 'AM';
        else out += code === 'a/p' ? (pm ? 'p' : 'a') : pm ? 'P' : 'A';
      }
    }
  }
  return out;
}
