import { inflateSync } from 'node:zlib';
import { describe, expect, it } from 'vitest';
import { aesDecryptCbc, aesDecryptEcb } from '../src/crypto/index.js';
import { DocxDocument } from '../src/model/index.js';
import { exportPdf } from '../src/pdf/index.js';
import { hashR6, pdfSecurity, permissionBits } from '../src/security/index.js';
import { bytesEqual } from '../src/util/bytes.js';
import { utf8Encode } from '../src/util/text.js';
import { buildDocx, p } from './helpers/docx.js';
import { testRegistry } from './helpers/fonts.js';
import { dictOf, latin1, pages, parsePdf, resolve, type PdfDictObj, type PdfStr } from './helpers/pdf.js';

describe('PDF permission bits', () => {
  it('sets reserved bits and clears denied permissions', () => {
    expect(permissionBits() >>> 0).toBe(0xfffffffc);
    const p = permissionBits({ print: false, copy: false, modify: false }) >>> 0;
    expect(p & 0x4).toBe(0);
    expect(p & 0x10).toBe(0);
    expect(p & 0x8).toBe(0);
    expect(p & 0x20).toBe(0x20);
    expect(p & 0xc0).toBe(0xc0);
    expect(permissionBits({ print: 'lowQuality' }) & 0x800).toBe(0);
  });
});

describe('PDF encryption (R6, AES-256)', () => {
  it('encrypts strings and streams so only the right passwords unlock them', async () => {
    const doc = await DocxDocument.open(await buildDocx({ body: p('Top secret pricing') }));
    const bytes = await exportPdf(doc, {
      fonts: testRegistry(),
      metadata: { title: 'Pricing' },
      encryption: pdfSecurity({ userPassword: 'open-me', ownerPassword: 'owner-pw', permissions: { copy: false, print: 'lowQuality' } }),
    });
    expect(latin1(bytes)).not.toContain('Pricing');
    const pdf = parsePdf(bytes);
    const enc = dictOf(pdf, pdf.trailer.get('Encrypt'));
    expect(enc.get('V')).toBe(5);
    expect(enc.get('R')).toBe(6);
    const U = (enc.get('U') as PdfStr).bytes;
    const O = (enc.get('O') as PdfStr).bytes;
    const UE = (enc.get('UE') as PdfStr).bytes;
    const user = utf8Encode('open-me');

    // Algorithm 11: user password check; Algorithm 12: owner password check.
    expect(bytesEqual(hashR6(user, U.subarray(32, 40)), U.subarray(0, 32))).toBe(true);
    expect(bytesEqual(hashR6(utf8Encode('wrong'), U.subarray(32, 40)), U.subarray(0, 32))).toBe(false);
    expect(bytesEqual(hashR6(utf8Encode('owner-pw'), O.subarray(32, 40), U), O.subarray(0, 32))).toBe(true);

    // File key from UE, then Perms must decrypt to P + "adb".
    const fileKey = aesDecryptCbc(hashR6(user, U.subarray(40, 48)), new Uint8Array(16), UE, 'none');
    const perms = aesDecryptEcb(fileKey, (enc.get('Perms') as PdfStr).bytes, 'none');
    expect(new DataView(perms.buffer).getInt32(0, true)).toBe(enc.get('P'));
    expect(String.fromCharCode(...perms.subarray(9, 12))).toBe('adb');
    expect(perms[8]).toBe(0x54);

    // Decrypt the page content stream and the title.
    const page = pages(pdf)[0]!;
    const content = resolve(pdf, page.get('Contents')) as PdfDictObj;
    const raw = content.stream!;
    const plain = aesDecryptCbc(fileKey, raw.subarray(0, 16), raw.subarray(16), 'pkcs7');
    expect(latin1(new Uint8Array(inflateSync(plain)))).toContain(' TJ ET');
    const title = (dictOf(pdf, pdf.trailer.get('Info')).get('Title') as PdfStr).bytes;
    expect(new TextDecoder().decode(aesDecryptCbc(fileKey, title.subarray(0, 16), title.subarray(16), 'pkcs7'))).toBe('Pricing');

    const root = dictOf(pdf, pdf.trailer.get('Root'));
    expect(dictOf(pdf, dictOf(pdf, root.get('Extensions')).get('ADBE')).get('ExtensionLevel')).toBe(8);
  });
});
