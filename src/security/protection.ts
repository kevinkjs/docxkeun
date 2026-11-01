/**
 * Word document protection (settings.xml w:documentProtection / w:writeProtection).
 *
 * This is an editing restriction that Word honours, not encryption: anyone can remove it from
 * the XML. Use password-to-open encryption to keep content confidential.
 *
 * Word's transitional hash: the password (max 15 chars) goes through the legacy Word 97 XOR hash,
 * whose 32-bit result is written as 8 hex digits in little-endian byte order, encoded UTF-16LE,
 * then H(salt || that) and spinCount rounds of H(hash || iterator). The ISO form
 * (w:algorithmName / w:hashValue / w:saltValue / w:spinCount) hashes the UTF-16LE password directly.
 */
import { ensurePart } from '../api/parts.js';
import { hash, hashNameFromString, randomBytes, type HashName } from '../crypto/index.js';
import type { DocxDocument } from '../model/document.js';
import { timingSafeEqual } from '../util/bytes.js';
import { base64Decode, base64Encode, utf16leEncode } from '../util/text.js';
import { attr, el, firstChild, insertOrdered, removeChildren, type XmlElement } from '../xml/index.js';

export type ProtectionMode = 'none' | 'readOnly' | 'comments' | 'trackedChanges' | 'forms';

export interface ProtectionInfo {
  mode: ProtectionMode;
  enforced: boolean;
  /** Formatting restricted to allowed styles. */
  formatting: boolean;
  hasPassword: boolean;
}

// First column of Word's XOR encryption matrix; each row continues by doubling with x^16+x^12+x^5+1.
const MATRIX_SEEDS = [0xaefc, 0x7b61, 0x4563, 0x0375, 0xd849, 0x6f45, 0xeb23, 0x47d3, 0xb861, 0x45a0, 0xaa51, 0x76b4, 0x3730, 0x3331, 0x1021];
const INITIAL_CODES = [0xe1f0, 0x1d0f, 0xcc9c, 0x84c0, 0x110c, 0x0e10, 0xf1ce, 0x313e, 0x1872, 0xe139, 0xd40f, 0x84f9, 0x280c, 0xa96a, 0x4ec3];

const MATRIX: number[][] = MATRIX_SEEDS.map((seed) => {
  const row = [seed];
  for (let j = 1; j < 7; j++) {
    let v = row[j - 1]! << 1;
    if (v & 0x10000) v = (v & 0xffff) ^ 0x1021;
    row.push(v);
  }
  return row;
});

function passwordBytes(password: string): number[] {
  return [...password.slice(0, 15)].map((ch) => {
    const c = ch.charCodeAt(0);
    return c & 0xff || (c >> 8) & 0xff;
  });
}

/** Low 16 bits of the legacy hash (also Excel's sheet protection "password" value). */
export function legacyVerifier(password: string): number {
  const bytes = passwordBytes(password);
  let v = 0;
  for (let i = bytes.length - 1; i >= 0; i--) v = (((v >> 14) & 1) | ((v << 1) & 0x7fff)) ^ bytes[i]!;
  return ((((v >> 14) & 1) | ((v << 1) & 0x7fff)) ^ bytes.length ^ 0xce4b) & 0xffff;
}

/** The legacy Word 97 32-bit password hash. */
export function legacyHash(password: string): number {
  const bytes = passwordBytes(password);
  if (bytes.length === 0) return 0;
  let key = INITIAL_CODES[bytes.length - 1]!;
  bytes.forEach((b, i) => {
    const row = MATRIX[15 - bytes.length + i]!;
    for (let bit = 0; bit < 7; bit++) if (b & (1 << bit)) key ^= row[bit]!;
  });
  return ((key << 16) | legacyVerifier(password)) >>> 0;
}

/** Legacy hash as 8 hex digits in little-endian byte order (Word's pre-processing step). */
export function legacyHashHexReversed(password: string): string {
  const h = legacyHash(password);
  return [h & 0xff, (h >>> 8) & 0xff, (h >>> 16) & 0xff, (h >>> 24) & 0xff].map((b) => b.toString(16).padStart(2, '0').toUpperCase()).join('');
}

/** H(salt || pwd), then spinCount rounds of H(hash || LE32 iterator). */
export function iteratedHash(algorithm: HashName, salt: Uint8Array, password: Uint8Array, spinCount: number): Uint8Array {
  const first = new Uint8Array(salt.length + password.length);
  first.set(salt);
  first.set(password, salt.length);
  let h = hash(algorithm, first);
  const buf = new Uint8Array(h.length + 4);
  const view = new DataView(buf.buffer);
  for (let i = 0; i < spinCount; i++) {
    buf.set(h);
    view.setUint32(h.length, i, true);
    h = hash(algorithm, buf);
  }
  return h;
}

const SID_TO_HASH: Record<string, HashName> = { '4': 'SHA-1', '12': 'SHA-256', '13': 'SHA-384', '14': 'SHA-512' };
const HASH_TO_SID: Record<HashName, string> = { 'SHA-1': '4', 'SHA-256': '12', 'SHA-384': '13', 'SHA-512': '14' };

const PASSWORD_ATTRS = [
  'w:cryptProviderType',
  'w:cryptAlgorithmClass',
  'w:cryptAlgorithmType',
  'w:cryptAlgorithmSid',
  'w:cryptSpinCount',
  'w:hash',
  'w:salt',
  'w:algorithmName',
  'w:hashValue',
  'w:saltValue',
  'w:spinCount',
  'w:cryptProvider',
  'w:algIdExt',
  'w:algIdExtSource',
  'w:cryptProviderTypeExt',
  'w:cryptProviderTypeExtSource',
];

/** Check a password against a protection element's stored hash. */
export function verifyProtectionPassword(node: XmlElement, password: string): boolean {
  const a = node.attrs;
  if (a['w:hash'] && a['w:salt']) {
    const algorithm = SID_TO_HASH[a['w:cryptAlgorithmSid'] ?? '4'];
    if (!algorithm) throw new Error(`Unsupported protection hash algorithm id ${a['w:cryptAlgorithmSid']}`);
    const spin = Number(a['w:cryptSpinCount'] ?? 0);
    const computed = iteratedHash(algorithm, base64Decode(a['w:salt']), utf16leEncode(legacyHashHexReversed(password)), spin);
    return timingSafeEqual(computed, base64Decode(a['w:hash']));
  }
  if (a['w:hashValue'] && a['w:saltValue']) {
    const algorithm = hashNameFromString(a['w:algorithmName'] ?? '');
    if (!algorithm) throw new Error(`Unsupported protection hash algorithm ${a['w:algorithmName']}`);
    const spin = Number(a['w:spinCount'] ?? 0);
    const salt = base64Decode(a['w:saltValue']);
    const expected = base64Decode(a['w:hashValue']);
    // ISO form: the password itself; some producers still apply Word's legacy step.
    const direct = iteratedHash(algorithm, salt, utf16leEncode(password), spin);
    if (timingSafeEqual(direct, expected)) return true;
    return timingSafeEqual(iteratedHash(algorithm, salt, utf16leEncode(legacyHashHexReversed(password)), spin), expected);
  }
  // No stored hash: protection without a password.
  return true;
}

/** Attributes Word writes for a password-protected element. */
export function passwordAttributes(password: string, options: { algorithm?: HashName; spinCount?: number } = {}): Record<string, string> {
  const algorithm = options.algorithm ?? 'SHA-512';
  const spinCount = options.spinCount ?? 100_000;
  const salt = randomBytes(16);
  const h = iteratedHash(algorithm, salt, utf16leEncode(legacyHashHexReversed(password)), spinCount);
  return {
    'w:cryptProviderType': 'rsaAES',
    'w:cryptAlgorithmClass': 'hash',
    'w:cryptAlgorithmType': 'typeAny',
    'w:cryptAlgorithmSid': HASH_TO_SID[algorithm],
    'w:cryptSpinCount': String(spinCount),
    'w:hash': base64Encode(h),
    'w:salt': base64Encode(salt),
  };
}

const SETTINGS_ORDER = [
  'w:writeProtection',
  'w:view',
  'w:zoom',
  'w:removePersonalInformation',
  'w:removeDateAndTime',
  'w:doNotDisplayPageBoundaries',
  'w:displayBackgroundShape',
  'w:printPostScriptOverText',
  'w:printFractionalCharacterWidth',
  'w:printFormsData',
  'w:embedTrueTypeFonts',
  'w:embedSystemFonts',
  'w:saveSubsetFonts',
  'w:saveFormsData',
  'w:mirrorMargins',
  'w:alignBordersAndEdges',
  'w:bordersDoNotSurroundHeader',
  'w:bordersDoNotSurroundFooter',
  'w:gutterAtTop',
  'w:hideSpellingErrors',
  'w:hideGrammaticalErrors',
  'w:activeWritingStyle',
  'w:proofState',
  'w:formsDesign',
  'w:attachedTemplate',
  'w:linkStyles',
  'w:stylePaneFormatFilter',
  'w:stylePaneSortMethod',
  'w:documentType',
  'w:mailMerge',
  'w:revisionView',
  'w:trackRevisions',
  'w:doNotTrackMoves',
  'w:doNotTrackFormatting',
  'w:documentProtection',
  'w:autoFormatOverride',
  'w:styleLockTheme',
  'w:styleLockQFSet',
  'w:defaultTabStop',
];

const MODE_TO_XML: Record<ProtectionMode, string> = { none: 'none', readOnly: 'readOnly', comments: 'comments', trackedChanges: 'trackedChanges', forms: 'forms' };

function settingsNode(doc: DocxDocument, name: string): XmlElement | undefined {
  return firstChild(doc.xmlSync(doc.parts.settings)?.root, name);
}

export function readProtection(doc: DocxDocument): ProtectionInfo {
  const node = settingsNode(doc, 'w:documentProtection');
  if (!node) return { mode: 'none', enforced: false, formatting: false, hasPassword: false };
  const edit = attr(node, 'w:edit') ?? 'none';
  const enforcement = attr(node, 'w:enforcement');
  // Word enforces protection when w:enforcement is missing (MS-OI29500).
  const enforced = enforcement === undefined ? true : !['0', 'false', 'off'].includes(enforcement);
  const mode = (edit === 'trackedChanges' ? 'trackedChanges' : (Object.keys(MODE_TO_XML) as ProtectionMode[]).find((m) => m === edit) ?? 'none') as ProtectionMode;
  return {
    mode,
    enforced: enforced && mode !== 'none',
    formatting: attr(node, 'w:formatting') === '1' || attr(node, 'w:formatting') === 'true',
    hasPassword: !!(node.attrs['w:hash'] || node.attrs['w:hashValue']),
  };
}

/** Turn protection on (or change it). A password is optional, as in Word. */
export async function setProtection(doc: DocxDocument, mode: ProtectionMode, options: { password?: string; formatting?: boolean; algorithm?: HashName; spinCount?: number } = {}): Promise<void> {
  const { name, xml } = await ensurePart(doc, 'settings');
  removeChildren(xml.root, 'w:documentProtection');
  if (mode !== 'none') {
    const attrs: Record<string, string> = { 'w:edit': MODE_TO_XML[mode], 'w:enforcement': '1' };
    if (options.formatting) attrs['w:formatting'] = '1';
    if (options.password) Object.assign(attrs, passwordAttributes(options.password, options));
    insertOrdered(xml.root, el('w:documentProtection', attrs), SETTINGS_ORDER);
  }
  doc.touch(name);
}

/** Stop enforcing protection. Throws when the password is wrong. */
export async function removeProtection(doc: DocxDocument, password = ''): Promise<void> {
  const node = settingsNode(doc, 'w:documentProtection');
  if (!node) return;
  if (!verifyProtectionPassword(node, password)) throw new Error('Incorrect protection password');
  const settings = doc.xmlSync(doc.parts.settings)!;
  removeChildren(settings.root, 'w:documentProtection');
  doc.touch(doc.parts.settings!);
}

export function checkProtectionPassword(doc: DocxDocument, password: string): boolean {
  const node = settingsNode(doc, 'w:documentProtection');
  return node ? verifyProtectionPassword(node, password) : true;
}

/** "Password to modify" / read-only recommended (w:writeProtection). */
export async function setWriteProtection(doc: DocxDocument, options: { password?: string; recommended?: boolean } | null): Promise<void> {
  const { name, xml } = await ensurePart(doc, 'settings');
  removeChildren(xml.root, 'w:writeProtection');
  if (options) {
    const attrs: Record<string, string> = {};
    if (options.recommended) attrs['w:recommended'] = '1';
    if (options.password) Object.assign(attrs, passwordAttributes(options.password));
    insertOrdered(xml.root, el('w:writeProtection', attrs), SETTINGS_ORDER);
  }
  doc.touch(name);
}

export function readWriteProtection(doc: DocxDocument): { recommended: boolean; hasPassword: boolean } | undefined {
  const node = settingsNode(doc, 'w:writeProtection');
  if (!node) return undefined;
  return { recommended: attr(node, 'w:recommended') === '1' || attr(node, 'w:recommended') === 'true', hasPassword: !!(node.attrs['w:hash'] || node.attrs['w:hashValue']) };
}

export function checkWriteProtectionPassword(doc: DocxDocument, password: string): boolean {
  const node = settingsNode(doc, 'w:writeProtection');
  return node ? verifyProtectionPassword(node, password) : true;
}

export { PASSWORD_ATTRS };
