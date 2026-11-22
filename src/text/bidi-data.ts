/**
 * Unicode Character Database properties used by the bidi algorithm (UAX #9): Bidi_Class,
 * Bidi_Paired_Bracket, Bidi_Paired_Bracket_Type and Bidi_Mirroring_Glyph.
 *
 * The tables are written by hand from DerivedBidiClass.txt, BidiBrackets.txt and BidiMirroring.txt
 * (Unicode 16) as compact range lists, so the library carries no generated or downloaded data. They
 * are parsed once, on first lookup; the parser rejects overlapping or unsorted ranges, so a typo in
 * a table fails loudly in the tests instead of silently misclassifying characters.
 */

/** The 23 bidi classes, in the order of their numeric codes. */
export const BIDI_CLASSES = [
  'L', 'R', 'AL', 'EN', 'ES', 'ET', 'AN', 'CS', 'NSM', 'BN', 'B', 'S', 'WS', 'ON',
  'LRE', 'LRO', 'RLE', 'RLO', 'PDF', 'LRI', 'RLI', 'FSI', 'PDI',
] as const;

export type BidiClass = (typeof BIDI_CLASSES)[number];

/**
 * Numeric class codes (indexes into BIDI_CLASSES). The algorithm works on these small integers;
 * the explicit controls are contiguous (LRE..PDF, LRI..PDI) so range checks can test them.
 */
export const BC = {
  L: 0, R: 1, AL: 2, EN: 3, ES: 4, ET: 5, AN: 6, CS: 7, NSM: 8, BN: 9, B: 10, S: 11, WS: 12, ON: 13,
  LRE: 14, LRO: 15, RLE: 16, RLO: 17, PDF: 18, LRI: 19, RLI: 20, FSI: 21, PDI: 22,
} as const satisfies Record<BidiClass, number>;

/**
 * Block defaults for unassigned code points (the @missing lines of DerivedBidiClass.txt). They
 * matter because the explicit table below only lists characters that differ from these, so the
 * bulk of Hebrew, Arabic and the RTL SMP blocks is covered by a handful of ranges.
 */
const DEFAULT_RANGES = `
0590-05FF:R 0600-07BF:AL 07C0-085F:R 0860-08FF:AL
FB1D-FB4F:R FB50-FDCF:AL FDF0-FDFF:AL FE70-FEFF:AL
10800-10CFF:R 10D00-10D3F:AL 10D40-10EBF:R 10EC0-10EFF:AL 10F00-10F2F:R 10F30-10F6F:AL 10F70-10FFF:R
1E800-1EC6F:R 1EC70-1ECBF:AL 1ECC0-1ECFF:R 1ED00-1ED4F:AL 1ED50-1EDFF:R 1EE00-1EEFF:AL 1EF00-1EFFF:R
`;

/**
 * Explicit Bidi_Class assignments. Anything not listed takes the block default above, else L, which
 * covers Latin, Greek, Cyrillic, CJK, Indic letters and most other scripts without listing them.
 * Whole combining-mark blocks (0300, 1AB0, 1DC0, FE20, ...) are listed as NSM including their few
 * unassigned slots, since any future assignment there will be a mark too.
 */
const CLASS_RANGES = `
# Basic Latin and Latin-1
0000-0008:BN 0009:S 000A:B 000B:S 000C:WS 000D:B 000E-001B:BN 001C-001E:B 001F:S 0020:WS
0021-0022:ON 0023-0025:ET 0026-002A:ON 002B:ES 002C:CS 002D:ES 002E-002F:CS 0030-0039:EN 003A:CS
003B-0040:ON 005B-0060:ON 007B-007E:ON 007F-0084:BN 0085:B 0086-009F:BN
00A0:CS 00A1:ON 00A2-00A5:ET 00A6-00A9:ON 00AB-00AC:ON 00AD:BN 00AE-00AF:ON 00B0-00B1:ET
00B2-00B3:EN 00B4:ON 00B6-00B8:ON 00B9:EN 00BB-00BF:ON 00D7:ON 00F7:ON
# Spacing modifiers, combining diacriticals, Greek, Cyrillic, Armenian
02B9-02BA:ON 02C2-02CF:ON 02D2-02DF:ON 02E5-02ED:ON 02EF-02FF:ON 0300-036F:NSM
0374-0375:ON 037E:ON 0384-0385:ON 0387:ON 03F6:ON 0483-0489:NSM 058A:ON 058D-058E:ON 058F:ET
# Hebrew points (letters are R by block default)
0591-05BD:NSM 05BF:NSM 05C1-05C2:NSM 05C4-05C5:NSM 05C7:NSM
# Arabic (letters are AL by block default): number signs, digits, separators, harakat
0600-0605:AN 0606-0607:ON 0609-060A:ET 060C:CS 060E-060F:ON 0610-061A:NSM 064B-065F:NSM
0660-0669:AN 066A:ET 066B-066C:AN 0670:NSM 06D6-06DC:NSM 06DD:AN 06DE:ON 06DF-06E4:NSM
06E7-06E8:NSM 06E9:ON 06EA-06ED:NSM 06F0-06F9:EN
# Syriac, Thaana, N'Ko, Samaritan, Mandaic, Arabic Extended-A/B
0711:NSM 0730-074A:NSM 07A6-07B0:NSM 07EB-07F3:NSM 07F6-07F9:ON 07FD:NSM
0816-0819:NSM 081B-0823:NSM 0825-0827:NSM 0829-082D:NSM 0859-085B:NSM
0890-0891:AN 0897-089F:NSM 08CA-08E1:NSM 08E2:AN 08E3-0902:NSM
# Indic scripts: nonspacing vowel signs, viramas, nuktas; currency signs
093A:NSM 093C:NSM 0941-0948:NSM 094D:NSM 0951-0957:NSM 0962-0963:NSM
0981:NSM 09BC:NSM 09C1-09C4:NSM 09CD:NSM 09E2-09E3:NSM 09F2-09F3:ET 09FB:ET 09FE:NSM
0A01-0A02:NSM 0A3C:NSM 0A41-0A42:NSM 0A47-0A48:NSM 0A4B-0A4D:NSM 0A51:NSM 0A70-0A71:NSM 0A75:NSM
0A81-0A82:NSM 0ABC:NSM 0AC1-0AC5:NSM 0AC7-0AC8:NSM 0ACD:NSM 0AE2-0AE3:NSM 0AF1:ET 0AFA-0AFF:NSM
0B01:NSM 0B3C:NSM 0B3F:NSM 0B41-0B44:NSM 0B4D:NSM 0B55-0B56:NSM 0B62-0B63:NSM
0B82:NSM 0BC0:NSM 0BCD:NSM 0BF3-0BF8:ON 0BF9:ET 0BFA:ON
0C00:NSM 0C04:NSM 0C3C:NSM 0C3E-0C40:NSM 0C46-0C48:NSM 0C4A-0C4D:NSM 0C55-0C56:NSM 0C62-0C63:NSM 0C78-0C7E:ON
0C81:NSM 0CBC:NSM 0CCC-0CCD:NSM 0CE2-0CE3:NSM
0D00-0D01:NSM 0D3B-0D3C:NSM 0D41-0D44:NSM 0D4D:NSM 0D62-0D63:NSM
0D81:NSM 0DCA:NSM 0DD2-0DD4:NSM 0DD6:NSM
# Thai, Lao, Tibetan, Myanmar
0E31:NSM 0E34-0E3A:NSM 0E3F:ET 0E47-0E4E:NSM 0EB1:NSM 0EB4-0EBC:NSM 0EC8-0ECE:NSM
0F18-0F19:NSM 0F35:NSM 0F37:NSM 0F39:NSM 0F3A-0F3D:ON 0F71-0F7E:NSM 0F80-0F84:NSM 0F86-0F87:NSM
0F8D-0F97:NSM 0F99-0FBC:NSM 0FC6:NSM
102D-1030:NSM 1032-1037:NSM 1039-103A:NSM 103D-103E:NSM 1058-1059:NSM 105E-1060:NSM 1071-1074:NSM
1082:NSM 1085-1086:NSM 108D:NSM 109D:NSM
# Ethiopic, Canadian syllabics, Ogham, Philippine scripts, Khmer, Mongolian, Limbu
135D-135F:NSM 1390-1399:ON 1400:ON 1680:WS 169B-169C:ON
1712-1714:NSM 1732-1733:NSM 1752-1753:NSM 1772-1773:NSM
17B4-17B5:NSM 17B7-17BD:NSM 17C6:NSM 17C9-17D3:NSM 17DB:ET 17DD:NSM 17F0-17F9:ON
1800-180A:ON 180B-180D:NSM 180E:BN 180F:NSM 1885-1886:NSM 18A9:NSM
1920-1922:NSM 1927-1928:NSM 1932:NSM 1939-193B:NSM 1940:ON 1944-1945:ON 19DE-19FF:ON
# Buginese, Tai Tham, Balinese, Sundanese, Batak, Lepcha, Vedic, combining supplements
1A17-1A18:NSM 1A1B:NSM 1A56:NSM 1A58-1A5E:NSM 1A60:NSM 1A62:NSM 1A65-1A6C:NSM 1A73-1A7C:NSM
1A7F:NSM 1AB0-1AFF:NSM
1B00-1B03:NSM 1B34:NSM 1B36-1B3A:NSM 1B3C:NSM 1B42:NSM 1B6B-1B73:NSM 1B80-1B81:NSM 1BA2-1BA5:NSM
1BA8-1BA9:NSM 1BAB-1BAD:NSM 1BE6:NSM 1BE8-1BE9:NSM 1BED:NSM 1BEF-1BF1:NSM
1C2C-1C33:NSM 1C36-1C37:NSM 1CD0-1CD2:NSM 1CD4-1CE0:NSM 1CE2-1CE8:NSM 1CED:NSM 1CF4:NSM
1CF8-1CF9:NSM 1DC0-1DFF:NSM
# Greek Extended spacing accents
1FBD:ON 1FBF-1FC1:ON 1FCD-1FCF:ON 1FDD-1FDF:ON 1FED-1FEF:ON 1FFD-1FFE:ON
# General punctuation and the explicit directional controls
2000-200A:WS 200B-200D:BN 200F:R 2010-2027:ON 2028:WS 2029:B 202A:LRE 202B:RLE 202C:PDF
202D:LRO 202E:RLO 202F:CS 2030-2034:ET 2035-2043:ON 2044:CS 2045-205E:ON 205F:WS 2060-2065:BN
2066:LRI 2067:RLI 2068:FSI 2069:PDI 206A-206F:BN
# Super/subscripts, currency, combining marks for symbols, letterlike symbols, number forms
2070:EN 2074-2079:EN 207A-207B:ES 207C-207E:ON 2080-2089:EN 208A-208B:ES 208C-208E:ON
20A0-20CF:ET 20D0-20F0:NSM
2100-2101:ON 2103-2106:ON 2108-2109:ON 2114:ON 2116-2118:ON 211E-2123:ON 2125:ON 2127:ON 2129:ON
212E:ET 213A-213B:ON 2140-2144:ON 214A-214D:ON 2150-215F:ON 2189-218B:ON
# Arrows, mathematical operators, technical, enclosed alphanumerics, shapes, dingbats
2190-2211:ON 2212:ES 2213:ET 2214-2335:ON 237B-2394:ON 2396-2429:ON 2440-244A:ON 2460-2487:ON
2488-249B:EN 24EA-26AB:ON 26AD-27FF:ON 2900-2BFF:ON
# Coptic, Tifinagh, Cyrillic Extended-A, supplemental punctuation, CJK
2CE5-2CEA:ON 2CEF-2CF1:NSM 2CF9-2CFF:ON 2D7F:NSM 2DE0-2DFF:NSM 2E00-2E7F:ON 2E80-2E99:ON
2E9B-2EF3:ON 2F00-2FD5:ON 2FF0-2FFF:ON
3000:WS 3001-3004:ON 3008-3020:ON 302A-302D:NSM 3030:ON 3036-3037:ON 303D-303F:ON 3099-309A:NSM
309B-309C:ON 30A0:ON 30FB:ON 31C0-31E5:ON 31EF:ON 321D-321E:ON 3250-325F:ON 327C-327E:ON
32B1-32BF:ON 32CC-32CF:ON 3377-337A:ON 33DE-33DF:ON 33FF:ON 4DC0-4DFF:ON
# Yi, Vai, Cyrillic Extended-B, Bamum, modifier tones, and the Brahmic scripts of A800-ABFF
A490-A4C6:ON A60D-A60F:ON A66F-A672:NSM A673:ON A674-A67D:NSM A67E-A67F:ON A69E-A69F:NSM
A6F0-A6F1:NSM A700-A721:ON A788:ON
A802:NSM A806:NSM A80B:NSM A825-A826:NSM A828-A82B:ON A82C:NSM A838-A839:ET A874-A877:ON
A8C4-A8C5:NSM A8E0-A8F1:NSM A8FF:NSM A926-A92D:NSM A947-A951:NSM A980-A982:NSM A9B3:NSM
A9B6-A9B9:NSM A9BC-A9BD:NSM A9E5:NSM AA29-AA2E:NSM AA31-AA32:NSM AA35-AA36:NSM AA43:NSM AA4C:NSM
AA7C:NSM AAB0:NSM AAB2-AAB4:NSM AAB7-AAB8:NSM AABE-AABF:NSM AAC1:NSM AAEC-AAED:NSM AAF6:NSM
AB6A-AB6B:ON ABE5:NSM ABE8:NSM ABED:NSM
# Presentation forms (Hebrew R and Arabic AL by block default), variation selectors, small and
# fullwidth forms, specials
FB1E:NSM FB29:ES FD3E-FD4F:ON FDCF:ON FDD0-FDEF:BN FDFD-FDFF:ON FE00-FE0F:NSM FE10-FE19:ON
FE20-FE2F:NSM FE30-FE4F:ON FE50:CS FE51:ON FE52:CS FE54:ON FE55:CS FE56-FE5E:ON FE5F:ET
FE60-FE61:ON FE62-FE63:ES FE64-FE66:ON FE68:ON FE69-FE6A:ET FE6B:ON FEFF:BN
FF01-FF02:ON FF03-FF05:ET FF06-FF0A:ON FF0B:ES FF0C:CS FF0D:ES FF0E-FF0F:CS FF10-FF19:EN FF1A:CS
FF1B-FF20:ON FF3B-FF40:ON FF5B-FF65:ON FFE0-FFE1:ET FFE2-FFE4:ON FFE5-FFE6:ET FFE8-FFEE:ON
FFF0-FFF8:BN FFF9-FFFD:ON FFFE-FFFF:BN
# SMP: Aegean and Greek numbers, Coptic epact, Old Permic, RTL blocks' marks and numbers
10101:ON 10140-1018C:ON 10190-1019C:ON 101A0:ON 101FD:NSM 102E0:NSM 102E1-102FB:EN 10376-1037A:NSM
1091F:ON 10A01-10A03:NSM 10A05-10A06:NSM 10A0C-10A0F:NSM 10A38-10A3A:NSM 10A3F:NSM 10AE5-10AE6:NSM
10B39-10B3F:ON 10D24-10D27:NSM 10D30-10D39:AN 10D69-10D6D:NSM 10E60-10E7E:AN 10EAB-10EAC:NSM 10EFC-10EFF:NSM
10F46-10F50:NSM 10F82-10F85:NSM
# SMP Brahmic scripts
11001:NSM 11038-11046:NSM 11052-11065:ON 11070:NSM 11073-11074:NSM 1107F-11081:NSM 110B3-110B6:NSM
110B9-110BA:NSM 110C2:NSM 11100-11102:NSM 11127-1112B:NSM 1112D-11134:NSM 11173:NSM 11180-11181:NSM
111B6-111BE:NSM 111C9-111CC:NSM 111CF:NSM 1122F-11231:NSM 11234:NSM 11236-11237:NSM 1123E:NSM
11241:NSM 112DF:NSM 112E3-112EA:NSM 11300-11301:NSM 1133B-1133C:NSM 11340:NSM 11366-1136C:NSM
11370-11374:NSM 11438-1143F:NSM 11442-11444:NSM 11446:NSM 1145E:NSM 114B3-114B8:NSM 114BA:NSM
114BF-114C0:NSM 114C2-114C3:NSM 115B2-115B5:NSM 115BC-115BD:NSM 115BF-115C0:NSM 115DC-115DD:NSM
11633-1163A:NSM 1163D:NSM 1163F-11640:NSM 11660-1166C:ON 116AB:NSM 116AD:NSM 116B0-116B5:NSM
116B7:NSM 1171D-1171F:NSM 11722-11725:NSM 11727-1172B:NSM 1182F-11837:NSM 11839-1183A:NSM
1193B-1193C:NSM 1193E:NSM 11943:NSM 119D4-119D7:NSM 119DA-119DB:NSM 119E0:NSM 11A01-11A0A:NSM
11A33-11A38:NSM 11A3B-11A3E:NSM 11A47:NSM 11A51-11A56:NSM 11A59-11A5B:NSM
11A8A-11A96:NSM 11A98-11A99:NSM 11C30-11C36:NSM 11C38-11C3D:NSM 11C3F:NSM 11C92-11CA7:NSM 11CAA-11CB0:NSM
11CB2-11CB3:NSM 11CB5-11CB6:NSM 11D31-11D36:NSM 11D3A:NSM 11D3C-11D3D:NSM 11D3F-11D45:NSM 11D47:NSM
11D90-11D91:NSM 11D95:NSM 11D97:NSM 11EF3-11EF4:NSM 11F00-11F01:NSM 11F36-11F3A:NSM 11F40:NSM
11F42:NSM 11F5A:NSM 11FD5-11FDC:ON 11FDD-11FE0:ET 11FE1-11FF1:ON
# Egyptian format controls, Bassa Vah, Pahawh Hmong, Miao, ideographic symbols, shorthand
13440:NSM 13447-13455:NSM 16AF0-16AF4:NSM 16B30-16B36:NSM 16F4F:NSM 16F8F-16F92:NSM 16FE2:ON
16FE4:NSM 1BC9D-1BC9E:NSM 1BCA0-1BCA3:BN
# Znamenny, musical symbols, Tai Xuan Jing, mathematical alphanumerics, SignWriting
1CF00-1CF2D:NSM 1CF30-1CF46:NSM 1D167-1D169:NSM 1D173-1D17A:BN 1D17B-1D182:NSM 1D185-1D18B:NSM
1D1AA-1D1AD:NSM 1D1E9-1D1EA:ON 1D200-1D241:ON 1D242-1D244:NSM 1D245:ON 1D300-1D356:ON 1D6DB:ON
1D715:ON 1D74F:ON 1D789:ON 1D7C3:ON 1D7CE-1D7FF:EN 1DA00-1DA36:NSM 1DA3B-1DA6C:NSM 1DA75:NSM
1DA84:NSM 1DA9B-1DA9F:NSM 1DAA1-1DAAF:NSM
# Glagolitic supplement, Cyrillic Extended-D, Hmong, Toto, Wancho, Nag Mundari, Mende, Adlam
1E000-1E006:NSM 1E008-1E018:NSM 1E01B-1E021:NSM 1E023-1E024:NSM 1E026-1E02A:NSM 1E08F:NSM
1E130-1E136:NSM 1E2AE:NSM 1E2EC-1E2EF:NSM 1E2FF:ET 1E4EC-1E4EF:NSM 1E5EE-1E5EF:NSM 1E8D0-1E8D6:NSM 1E944-1E94A:NSM
1EEF0-1EEF1:ON
# Game pieces, enclosed alphanumerics, emoji and symbol blocks
1F000-1F0FF:ON 1F100-1F10A:EN 1F10B-1F10F:ON 1F12F:ON 1F16A-1F16F:ON 1F1AD:ON 1F260-1F265:ON
1F300-1FBEF:ON 1FBF0-1FBF9:EN
# Noncharacters, tags and variation selectors supplement
1FFFE-1FFFF:BN 2FFFE-2FFFF:BN 3FFFE-3FFFF:BN 4FFFE-4FFFF:BN 5FFFE-5FFFF:BN 6FFFE-6FFFF:BN
7FFFE-7FFFF:BN 8FFFE-8FFFF:BN 9FFFE-9FFFF:BN AFFFE-AFFFF:BN BFFFE-BFFFF:BN CFFFE-CFFFF:BN
DFFFE-DFFFF:BN E0000-E00FF:BN E0100-E01EF:NSM E01F0-E0FFF:BN EFFFE-EFFFF:BN FFFFE-FFFFF:BN
10FFFE-10FFFF:BN
`;

/** Bidi_Paired_Bracket: each opening bracket and its closing partner (BidiBrackets.txt). */
const BRACKET_PAIRS = `
0028:0029 005B:005D 007B:007D 0F3A:0F3B 0F3C:0F3D 169B:169C 2045:2046 207D:207E 208D:208E
2308:2309 230A:230B 2329:232A 2768:2769 276A:276B 276C:276D 276E:276F 2770:2771 2772:2773 2774:2775
27C5:27C6 27E6:27E7 27E8:27E9 27EA:27EB 27EC:27ED 27EE:27EF 2983:2984 2985:2986 2987:2988 2989:298A
298B:298C 298D:2990 298F:298E 2991:2992 2993:2994 2995:2996 2997:2998 29D8:29D9 29DA:29DB 29FC:29FD
2E22:2E23 2E24:2E25 2E26:2E27 2E28:2E29 2E55:2E56 2E57:2E58 2E59:2E5A 2E5B:2E5C
3008:3009 300A:300B 300C:300D 300E:300F 3010:3011 3014:3015 3016:3017 3018:3019 301A:301B
FE59:FE5A FE5B:FE5C FE5D:FE5E FF08:FF09 FF3B:FF3D FF5B:FF5D FF5F:FF60 FF62:FF63
`;

/** Bidi_Mirroring_Glyph pairs (BidiMirroring.txt); every listed pair mirrors both ways. */
const MIRROR_PAIRS = `
0028:0029 003C:003E 005B:005D 007B:007D 00AB:00BB 0F3A:0F3B 0F3C:0F3D 169B:169C 2039:203A 2045:2046
207D:207E 208D:208E 2208:220B 2209:220C 220A:220D 2215:29F5 221F:2BFE 2220:29A3 2221:299B 2222:29A0
2224:2AEE 223C:223D 2243:22CD 2245:224C 2252:2253 2254:2255 2264:2265 2266:2267 2268:2269 226A:226B
226E:226F 2270:2271 2272:2273 2274:2275 2276:2277 2278:2279 227A:227B 227C:227D 227E:227F 2280:2281
2282:2283 2284:2285 2286:2287 2288:2289 228A:228B 228F:2290 2291:2292 2298:29B8 22A2:22A3 22A6:2ADE
22A8:2AE4 22A9:2AE3 22AB:2AE5 22B0:22B1 22B2:22B3 22B4:22B5 22B6:22B7 22B8:27DC 22C9:22CA 22CB:22CC
22D0:22D1 22D6:22D7 22D8:22D9 22DA:22DB 22DC:22DD 22DE:22DF 22E0:22E1 22E2:22E3 22E4:22E5 22E6:22E7
22E8:22E9 22EA:22EB 22EC:22ED 22F0:22F1 22F2:22FA 22F3:22FB 22F4:22FC 22F6:22FD 22F7:22FE
2308:2309 230A:230B 2329:232A 2768:2769 276A:276B 276C:276D 276E:276F 2770:2771 2772:2773 2774:2775
27C3:27C4 27C5:27C6 27C8:27C9 27CB:27CD 27D5:27D6 27DD:27DE 27E2:27E3 27E4:27E5 27E6:27E7 27E8:27E9
27EA:27EB 27EC:27ED 27EE:27EF 2983:2984 2985:2986 2987:2988 2989:298A 298B:298C 298D:2990 298E:298F
2991:2992 2993:2994 2995:2996 2997:2998 29A4:29A5 29A8:29A9 29AA:29AB 29AC:29AD 29AE:29AF 29C0:29C1
29C4:29C5 29CF:29D0 29D1:29D2 29D4:29D5 29D8:29D9 29DA:29DB 29E8:29E9 29F8:29F9 29FC:29FD 2A2B:2A2C
2A2D:2A2E 2A34:2A35 2A3C:2A3D 2A64:2A65 2A79:2A7A 2A7B:2A7C 2A7D:2A7E 2A7F:2A80 2A81:2A82 2A83:2A84
2A85:2A86 2A87:2A88 2A89:2A8A 2A8B:2A8C 2A8D:2A8E 2A8F:2A90 2A91:2A92 2A93:2A94 2A95:2A96 2A97:2A98
2A99:2A9A 2A9B:2A9C 2A9D:2A9E 2A9F:2AA0 2AA1:2AA2 2AA6:2AA7 2AA8:2AA9 2AAA:2AAB 2AAC:2AAD 2AAF:2AB0
2AB1:2AB2 2AB3:2AB4 2AB5:2AB6 2AB7:2AB8 2AB9:2ABA 2ABB:2ABC 2ABD:2ABE 2ABF:2AC0 2AC1:2AC2 2AC3:2AC4
2AC5:2AC6 2AC7:2AC8 2AC9:2ACA 2ACB:2ACC 2ACD:2ACE 2ACF:2AD0 2AD1:2AD2 2AD3:2AD4 2AD5:2AD6 2AEC:2AED
2AF7:2AF8 2AF9:2AFA 2E02:2E03 2E04:2E05 2E09:2E0A 2E0C:2E0D 2E1C:2E1D 2E20:2E21 2E22:2E23 2E24:2E25
2E26:2E27 2E28:2E29 2E55:2E56 2E57:2E58 2E59:2E5A 2E5B:2E5C 3008:3009 300A:300B 300C:300D 300E:300F
3010:3011 3014:3015 3016:3017 3018:3019 301A:301B FE59:FE5A FE5B:FE5C FE5D:FE5E FE64:FE65 FF08:FF09
FF1C:FF1E FF3B:FF3D FF5B:FF5D FF5F:FF60 FF62:FF63
`;

interface RangeTable {
  starts: Int32Array;
  ends: Int32Array;
  values: Uint8Array;
}

function parseRanges(source: string): RangeTable {
  const entries = source.replace(/#.*$/gm, '').trim().split(/\s+/);
  const table: RangeTable = {
    starts: new Int32Array(entries.length),
    ends: new Int32Array(entries.length),
    values: new Uint8Array(entries.length),
  };
  let previousEnd = -1;
  entries.forEach((entry, i) => {
    const match = /^([0-9A-F]{4,6})(?:-([0-9A-F]{4,6}))?:([A-Z]+)$/.exec(entry);
    const value = match ? (BC as Record<string, number>)[match[3]!] : undefined;
    if (!match || value === undefined) throw new Error(`Malformed bidi table entry: ${entry}`);
    const start = parseInt(match[1]!, 16);
    const end = match[2] === undefined ? start : parseInt(match[2], 16);
    if (start <= previousEnd || end < start) throw new Error(`Bidi table entry out of order: ${entry}`);
    table.starts[i] = start;
    table.ends[i] = end;
    table.values[i] = value;
    previousEnd = end;
  });
  return table;
}

/** The value of the range containing cp, or -1. */
function lookup(table: RangeTable, cp: number): number {
  let lo = 0;
  let hi = table.starts.length - 1;
  while (lo <= hi) {
    const mid = (lo + hi) >> 1;
    if (cp < table.starts[mid]!) hi = mid - 1;
    else if (cp > table.ends[mid]!) lo = mid + 1;
    else return table.values[mid]!;
  }
  return -1;
}

let defaults: RangeTable | undefined;
let classes: RangeTable | undefined;
let bmp: Uint8Array | undefined;

/** Text is overwhelmingly BMP, so the BMP is flattened into a 64 KB array for O(1) lookups. */
function bmpTable(): Uint8Array {
  defaults ??= parseRanges(DEFAULT_RANGES);
  classes ??= parseRanges(CLASS_RANGES);
  const table = new Uint8Array(0x10000); // 0 = L
  for (const layer of [defaults, classes]) {
    for (let i = 0; i < layer.starts.length && layer.starts[i]! < 0x10000; i++) {
      table.fill(layer.values[i]!, layer.starts[i]!, Math.min(layer.ends[i]!, 0xffff) + 1);
    }
  }
  return (bmp = table);
}

/** Bidi_Class of a code point as a numeric code (see BC). */
export function bidiClassCode(cp: number): number {
  if (cp >= 0 && cp < 0x10000) return (bmp ?? bmpTable())[cp]!;
  if (!(cp >= 0x10000 && cp <= 0x10ffff)) return BC.L;
  if (!bmp) bmpTable();
  const explicit = lookup(classes!, cp);
  if (explicit >= 0) return explicit;
  const fallback = lookup(defaults!, cp);
  return fallback >= 0 ? fallback : BC.L;
}

function parsePairs(source: string): [number, number][] {
  return source
    .trim()
    .split(/\s+/)
    .map((entry) => {
      const [a, b] = entry.split(':');
      return [parseInt(a!, 16), parseInt(b!, 16)];
    });
}

/** Bidi_Paired_Bracket_Type values. */
export const BPT_NONE = 0;
export const BPT_OPEN = 1;
export const BPT_CLOSE = 2;

let bracketPartner: Map<number, number> | undefined;
let bracketOpeners: Set<number> | undefined;

function loadBrackets(): Map<number, number> {
  const partner = new Map<number, number>();
  const openers = new Set<number>();
  for (const [open, close] of parsePairs(BRACKET_PAIRS)) {
    partner.set(open, close);
    partner.set(close, open);
    openers.add(open);
  }
  bracketOpeners = openers;
  return (bracketPartner = partner);
}

/** Bidi_Paired_Bracket_Type: BPT_OPEN, BPT_CLOSE or BPT_NONE. */
export function pairedBracketType(cp: number): number {
  const partner = bracketPartner ?? loadBrackets();
  if (!partner.has(cp)) return BPT_NONE;
  return bracketOpeners!.has(cp) ? BPT_OPEN : BPT_CLOSE;
}

/** Bidi_Paired_Bracket: the other bracket of the pair, if cp is a paired bracket. */
export function pairedBracket(cp: number): number | undefined {
  return (bracketPartner ?? loadBrackets()).get(cp);
}

let mirrors: Map<number, number> | undefined;

/** Bidi_Mirroring_Glyph: the character that mirrors cp, if there is one. */
export function mirrorGlyph(cp: number): number | undefined {
  if (!mirrors) {
    mirrors = new Map();
    for (const [a, b] of parsePairs(MIRROR_PAIRS)) {
      mirrors.set(a, b);
      mirrors.set(b, a);
    }
  }
  return mirrors.get(cp);
}
