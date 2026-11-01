/**
 * PDF Standard Security Handler, revision 6 (ISO 32000-2 §7.6.4): AES-256 encryption with an
 * optional user (open) password and an owner password guarding the permission flags.
 *
 * Permission flags are honoured by conforming viewers; only the user password protects content.
 */
import { aesEncryptCbc, aesEncryptEcb, hash, hashConcat, randomBytes } from '../crypto/index.js';
import { name, str, type ObjectCrypt, type PdfDict } from '../pdf/objects.js';
import type { PdfEncryption } from '../pdf/writer.js';
import { utf8Encode } from '../util/text.js';

export interface PdfPermissions {
  /** Printing; 'lowQuality' allows only degraded printing. */
  print?: boolean | 'lowQuality';
  modify?: boolean;
  /** Copying text and graphics. */
  copy?: boolean;
  /** Adding comments and filling forms. */
  annotate?: boolean;
  fillForms?: boolean;
  /** Text extraction for accessibility (screen readers). */
  accessibility?: boolean;
  /** Inserting, rotating and deleting pages. */
  assemble?: boolean;
}

export interface PdfSecurityOptions {
  /** Password needed to open the file. Empty or missing: anyone can open it. */
  userPassword?: string;
  /** Password that lifts the permission restrictions. Random when missing. */
  ownerPassword?: string;
  /** Allowed operations; anything not set to false is allowed. */
  permissions?: PdfPermissions;
}

/** Passwords are UTF-8 (after normalisation), at most 127 bytes. */
function passwordBytes(password: string): Uint8Array {
  return utf8Encode(password.normalize('NFKC')).subarray(0, 127);
}

/** Algorithm 2.B: the revision 6 password hash. */
export function hashR6(password: Uint8Array, salt: Uint8Array, userKey: Uint8Array = new Uint8Array()): Uint8Array {
  let k = hashConcat('SHA-256', password, salt, userKey);
  let e: Uint8Array = new Uint8Array();
  for (let round = 0; round < 64 || e[e.length - 1]! > round - 32; round++) {
    const unit = new Uint8Array(password.length + k.length + userKey.length);
    unit.set(password);
    unit.set(k, password.length);
    unit.set(userKey, password.length + k.length);
    const k1 = new Uint8Array(unit.length * 64);
    for (let i = 0; i < 64; i++) k1.set(unit, i * unit.length);
    e = aesEncryptCbc(k.subarray(0, 16), k.subarray(16, 32), k1, 'none');
    let sum = 0;
    for (let i = 0; i < 16; i++) sum += e[i]!;
    k = hash(sum % 3 === 0 ? 'SHA-256' : sum % 3 === 1 ? 'SHA-384' : 'SHA-512', e);
  }
  return k.subarray(0, 32);
}

/** The /P value: reserved bits set, permission bits per the options. */
export function permissionBits(p: PdfPermissions = {}): number {
  let bits = 0xfffff0c0; // bits 7-8 and 13-32 must be 1
  const allow = (flag: boolean | 'lowQuality' | undefined) => flag !== false;
  if (allow(p.print)) bits |= 1 << 2;
  if (allow(p.modify)) bits |= 1 << 3;
  if (allow(p.copy)) bits |= 1 << 4;
  if (allow(p.annotate)) bits |= 1 << 5;
  if (allow(p.fillForms)) bits |= 1 << 8;
  if (allow(p.accessibility)) bits |= 1 << 9;
  if (allow(p.assemble)) bits |= 1 << 10;
  if (allow(p.print) && p.print !== 'lowQuality') bits |= 1 << 11;
  return bits | 0; // signed 32-bit, as PDF writes it
}

/** Build the security handler for exportPdf({ encryption }). */
export function pdfSecurity(options: PdfSecurityOptions = {}): PdfEncryption {
  const fileKey = randomBytes(32);
  const user = passwordBytes(options.userPassword ?? '');
  const owner = passwordBytes(options.ownerPassword ?? Array.from(randomBytes(24), (b) => b.toString(36)).join(''));

  const uValidation = randomBytes(8);
  const uKeySalt = randomBytes(8);
  const U = new Uint8Array(48);
  U.set(hashR6(user, uValidation));
  U.set(uValidation, 32);
  U.set(uKeySalt, 40);
  const UE = aesEncryptCbc(hashR6(user, uKeySalt), new Uint8Array(16), fileKey, 'none');

  const oValidation = randomBytes(8);
  const oKeySalt = randomBytes(8);
  const O = new Uint8Array(48);
  O.set(hashR6(owner, oValidation, U));
  O.set(oValidation, 32);
  O.set(oKeySalt, 40);
  const OE = aesEncryptCbc(hashR6(owner, oKeySalt, U), new Uint8Array(16), fileKey, 'none');

  const P = permissionBits(options.permissions);
  const perms = new Uint8Array(16);
  new DataView(perms.buffer).setInt32(0, P, true);
  perms.fill(0xff, 4, 8);
  perms.set([0x54 /* T: metadata encrypted */, 0x61, 0x64, 0x62], 8);
  perms.set(randomBytes(4), 12);
  const Perms = aesEncryptEcb(fileKey, perms, 'none');

  const crypt: ObjectCrypt = {
    string: (bytes) => {
      const iv = randomBytes(16);
      const out = aesEncryptCbc(fileKey, iv, bytes, 'pkcs7');
      const result = new Uint8Array(16 + out.length);
      result.set(iv);
      result.set(out, 16);
      return result;
    },
    stream: (bytes) => crypt.string(bytes),
  };

  const dict: PdfDict = {
    Filter: name('Standard'),
    V: 5,
    R: 6,
    Length: 256,
    CF: { StdCF: { Type: name('CryptFilter'), AuthEvent: name('DocOpen'), CFM: name('AESV3'), Length: 32 } },
    StmF: name('StdCF'),
    StrF: name('StdCF'),
    O: str(O, true),
    U: str(U, true),
    OE: str(OE, true),
    UE: str(UE, true),
    P,
    Perms: str(Perms, true),
    EncryptMetadata: true,
  };
  return {
    dict,
    forObject: () => crypt,
    extensions: { ADBE: { BaseVersion: name('1.7'), ExtensionLevel: 8 } },
  };
}
