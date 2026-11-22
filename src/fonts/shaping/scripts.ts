/**
 * Unicode scripts as OpenType script tags, and script itemization of text. The Script
 * property comes from the JS regex engine (\p{Script=...}); results are cached per code
 * point. Common and Inherited characters (spaces, digits, punctuation, combining marks) take
 * the script of the run around them, and paired brackets follow their opening partner, the
 * way ICU and HarfBuzz itemize.
 */

/** Unicode Script value → OpenType script tag. */
const SCRIPT_TAGS: ReadonlyArray<readonly [string, string]> = [
  ['Latin', 'latn'], ['Greek', 'grek'], ['Cyrillic', 'cyrl'], ['Armenian', 'armn'], ['Hebrew', 'hebr'],
  ['Arabic', 'arab'], ['Syriac', 'syrc'], ['Thaana', 'thaa'], ['Nko', 'nko '], ['Samaritan', 'samr'],
  ['Mandaic', 'mand'], ['Devanagari', 'deva'], ['Bengali', 'beng'], ['Gurmukhi', 'guru'], ['Gujarati', 'gujr'],
  ['Oriya', 'orya'], ['Tamil', 'taml'], ['Telugu', 'telu'], ['Kannada', 'knda'], ['Malayalam', 'mlym'],
  ['Sinhala', 'sinh'], ['Thai', 'thai'], ['Lao', 'lao '], ['Tibetan', 'tibt'], ['Myanmar', 'mymr'],
  ['Georgian', 'geor'], ['Hangul', 'hang'], ['Ethiopic', 'ethi'], ['Cherokee', 'cher'],
  ['Canadian_Aboriginal', 'cans'], ['Ogham', 'ogam'], ['Runic', 'runr'], ['Khmer', 'khmr'], ['Mongolian', 'mong'],
  ['Hiragana', 'kana'], ['Katakana', 'kana'], ['Bopomofo', 'bopo'], ['Han', 'hani'], ['Yi', 'yi  '],
  ['Tifinagh', 'tfng'], ['Coptic', 'copt'], ['Javanese', 'java'], ['Balinese', 'bali'], ['Sundanese', 'sund'],
  ['Tai_Tham', 'lana'], ['Tai_Le', 'tale'], ['New_Tai_Lue', 'talu'], ['Tai_Viet', 'tavt'], ['Cham', 'cham'],
  ['Lepcha', 'lepc'], ['Limbu', 'limb'], ['Buginese', 'bugi'], ['Batak', 'batk'], ['Tagalog', 'tglg'],
  ['Hanunoo', 'hano'], ['Buhid', 'buhd'], ['Tagbanwa', 'tagb'], ['Vai', 'vai '], ['Ol_Chiki', 'olck'],
  ['Meetei_Mayek', 'mtei'], ['Saurashtra', 'saur'], ['Kayah_Li', 'kali'], ['Rejang', 'rjng'], ['Adlam', 'adlm'],
  ['Hanifi_Rohingya', 'rohg'], ['Chakma', 'cakm'], ['Sharada', 'shrd'], ['Brahmi', 'brah'], ['Kaithi', 'kthi'],
  ['Glagolitic', 'glag'], ['Gothic', 'goth'], ['Braille', 'brai'], ['Bamum', 'bamu'], ['Osage', 'osge'],
  ['Lisu', 'lisu'], ['Newa', 'newa'], ['Grantha', 'gran'], ['Tirhuta', 'tirh'], ['Modi', 'modi'], ['Takri', 'takr'],
  ['Khojki', 'khoj'], ['Mahajani', 'mahj'], ['Siddham', 'sidd'], ['Ahom', 'ahom'], ['Dogra', 'dogr'],
  ['Phags_Pa', 'phag'], ['Syloti_Nagri', 'sylo'], ['Kharoshthi', 'khar'], ['Manichaean', 'mani'],
  ['Psalter_Pahlavi', 'phlp'], ['Sogdian', 'sogd'], ['Yezidi', 'yezi'], ['Old_Italic', 'ital'],
  ['Deseret', 'dsrt'], ['Shavian', 'shaw'], ['Osmanya', 'osma'], ['Cuneiform', 'xsux'],
  ['Egyptian_Hieroglyphs', 'egyp'],
];

const NEUTRAL_RES = [/\p{Script=Common}/u, /\p{Script=Inherited}/u];
let scriptRes: Array<readonly [RegExp, string]> | undefined;
const scriptCache = new Map<number, string>();

function scriptRegexes(): Array<readonly [RegExp, string]> {
  if (!scriptRes) {
    scriptRes = [];
    for (const [name, tag] of SCRIPT_TAGS) {
      try {
        scriptRes.push([new RegExp(`\\p{Script=${name}}`, 'u'), tag]);
      } catch {
        // A runtime with older Unicode data: that script just itemizes as 'DFLT'.
      }
    }
  }
  return scriptRes;
}

/** The OpenType script tag of a code point ('DFLT' for Common, Inherited and unknown). */
export function scriptOf(cp: number): string {
  if (cp < 0x80) return (cp >= 0x41 && cp <= 0x5a) || (cp >= 0x61 && cp <= 0x7a) ? 'latn' : 'DFLT';
  if (cp >= 0xc0 && cp <= 0x24f && cp !== 0xd7 && cp !== 0xf7) return 'latn';
  const cached = scriptCache.get(cp);
  if (cached) return cached;
  const s = String.fromCodePoint(cp);
  let tag = 'DFLT';
  if (!NEUTRAL_RES.some((re) => re.test(s))) {
    for (const [re, t] of scriptRegexes()) {
      if (re.test(s)) {
        tag = t;
        break;
      }
    }
  }
  if (scriptCache.size < 8192) scriptCache.set(cp, tag);
  return tag;
}

export interface ScriptRun {
  /** UTF-16 start index (inclusive). */
  start: number;
  /** UTF-16 end index (exclusive). */
  end: number;
  script: string;
}

const BRACKET_PAIRS: ReadonlyArray<readonly [number, number]> = [
  [0x28, 0x29], [0x5b, 0x5d], [0x7b, 0x7d], [0xab, 0xbb], [0x2039, 0x203a], [0x2045, 0x2046], [0x207d, 0x207e],
  [0x208d, 0x208e], [0x2329, 0x232a], [0x27e8, 0x27e9], [0x3008, 0x3009], [0x300a, 0x300b], [0x300c, 0x300d],
  [0x300e, 0x300f], [0x3010, 0x3011], [0x3014, 0x3015], [0x3016, 0x3017], [0x3018, 0x3019], [0x301a, 0x301b],
  [0xff08, 0xff09], [0xff3b, 0xff3d], [0xff5b, 0xff5d], [0xff5f, 0xff60], [0xff62, 0xff63],
];
const OPEN_BRACKETS = new Set(BRACKET_PAIRS.map(([o]) => o));
const CLOSE_TO_OPEN = new Map(BRACKET_PAIRS.map(([o, c]) => [c, o]));

/** Splits text into runs of one script each; neutral characters join the surrounding run. */
export function itemizeScripts(text: string): ScriptRun[] {
  const runs: ScriptRun[] = [];
  let runStart = 0;
  let runScript = 'DFLT';
  const open: Array<{ cp: number; script: string }> = [];
  for (let i = 0; i < text.length; ) {
    const cp = text.codePointAt(i)!;
    let script = scriptOf(cp);
    if (script === 'DFLT') {
      if (OPEN_BRACKETS.has(cp)) {
        if (open.length < 64) open.push({ cp, script: runScript });
      } else {
        const opener = CLOSE_TO_OPEN.get(cp);
        if (opener !== undefined) {
          for (let k = open.length - 1; k >= 0; k--) {
            if (open[k]!.cp === opener) {
              script = open[k]!.script;
              open.length = k;
              break;
            }
          }
        }
      }
    }
    if (script !== 'DFLT' && script !== runScript) {
      if (runScript === 'DFLT') {
        runScript = script;
        for (const o of open) if (o.script === 'DFLT') o.script = script;
      } else {
        runs.push({ start: runStart, end: i, script: runScript });
        runStart = i;
        runScript = script;
      }
    }
    i += cp > 0xffff ? 2 : 1;
  }
  if (text.length > runStart) runs.push({ start: runStart, end: text.length, script: runScript });
  return runs;
}

/** The first non-neutral script of the text, or 'DFLT'. */
export function detectScript(text: string): string {
  for (let i = 0; i < text.length; ) {
    const cp = text.codePointAt(i)!;
    const s = scriptOf(cp);
    if (s !== 'DFLT') return s;
    i += cp > 0xffff ? 2 : 1;
  }
  return 'DFLT';
}

/** Indic scripts with their new-style ("version 2") OpenType tags. */
export const INDIC_NEW_TAGS: Readonly<Record<string, string>> = {
  deva: 'dev2',
  beng: 'bng2',
  guru: 'gur2',
  gujr: 'gjr2',
  orya: 'ory2',
  taml: 'tml2',
  telu: 'tel2',
  knda: 'knd2',
  mlym: 'mlm2',
};

const INDIC_OLD_TAGS: Readonly<Record<string, string>> = Object.fromEntries(Object.entries(INDIC_NEW_TAGS).map(([o, n]) => [n, o]));

/** Maps a script tag given as 'dev2' etc. back to the Unicode-script tag ('deva'). */
export function canonicalScript(tag: string): string {
  return INDIC_OLD_TAGS[tag] ?? tag;
}

/** Tags to look up in a font for a script, preferred first ('dev2' before 'deva'). */
export function scriptTagCandidates(script: string): string[] {
  if (script === 'DFLT') return [];
  const newTag = INDIC_NEW_TAGS[script];
  return newTag ? [newTag, script] : [script];
}
