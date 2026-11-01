import { describe, expect, it } from 'vitest';
import { blocksToText, DocxDocument } from '../src/model/index.js';
import { decryptPackage, emptyStorage, encryptPackage, PasswordRequiredError, readCfb, writeCfb, WrongPasswordError } from '../src/security/index.js';
import { bytesEqual } from '../src/util/bytes.js';
import { utf16leDecode } from '../src/util/text.js';
import { buildDocx, p } from './helpers/docx.js';

function filled(n: number, seed: number): Uint8Array {
  const out = new Uint8Array(n);
  let x = seed;
  for (let i = 0; i < n; i++) {
    x = (x * 1103515245 + 12345) >>> 0;
    out[i] = x >>> 24;
  }
  return out;
}

describe('compound file', () => {
  it('round-trips nested storages, mini streams and big streams', () => {
    const root = emptyStorage();
    root.streams.set('small', filled(100, 1));
    root.streams.set('edge', filled(4095, 2));
    root.streams.set('big', filled(70_000, 3));
    root.streams.set('empty', new Uint8Array());
    const sub = emptyStorage();
    sub.streams.set('\u0006Primary', filled(200, 4));
    const deeper = emptyStorage();
    deeper.streams.set('deep', filled(5000, 5));
    sub.storages.set('Deeper', deeper);
    root.storages.set('\u0006DataSpaces', sub);
    for (let i = 0; i < 20; i++) root.streams.set(`s${i}`, filled(10 + i, 10 + i));

    const bytes = writeCfb(root);
    expect(bytes.length % 512).toBe(0);
    const back = readCfb(bytes);
    for (const [name, data] of root.streams) expect(bytesEqual(back.streams.get(name)!, data), name).toBe(true);
    expect(bytesEqual(back.storages.get('\u0006DataSpaces')!.streams.get('\u0006Primary')!, sub.streams.get('\u0006Primary')!)).toBe(true);
    expect(bytesEqual(back.storages.get('\u0006DataSpaces')!.storages.get('Deeper')!.streams.get('deep')!, filled(5000, 5))).toBe(true);
  });

  it('handles files that need DIFAT sectors', () => {
    const root = emptyStorage();
    root.streams.set('huge', filled(8 * 1024 * 1024, 7)); // > 109 FAT sectors
    const back = readCfb(writeCfb(root));
    expect(bytesEqual(back.streams.get('huge')!, root.streams.get('huge')!)).toBe(true);
  });
});

describe('password-to-open (Agile encryption)', () => {
  it('encrypts and decrypts, rejecting wrong passwords and tampering', async () => {
    const zip = await buildDocx({ body: p('Confidential terms: royalty 3.75%') });
    const encrypted = encryptPackage(zip, 'Rahasia#2026', { spinCount: 1000 });
    expect([...encrypted.subarray(0, 4)]).toEqual([0xd0, 0xcf, 0x11, 0xe0]);
    const cfb = readCfb(encrypted);
    expect([...cfb.streams.keys()].sort()).toEqual(['EncryptedPackage', 'EncryptionInfo']);
    const ds = cfb.storages.get('\u0006DataSpaces')!;
    expect([...ds.streams.keys()].sort()).toEqual(['DataSpaceMap', 'Version']);
    const version = ds.streams.get('Version')!;
    expect(utf16leDecode(version.subarray(4, 4 + 60))).toBe('Microsoft.Container.DataSpaces');
    const primary = ds.storages.get('TransformInfo')!.storages.get('StrongEncryptionTransform')!.streams.get('\u0006Primary')!;
    expect(new DataView(primary.buffer, primary.byteOffset).getUint32(0, true)).toBe(88);
    const info = cfb.streams.get('EncryptionInfo')!;
    expect([...info.subarray(0, 8)]).toEqual([4, 0, 4, 0, 0x40, 0, 0, 0]);

    expect(bytesEqual(decryptPackage(encrypted, 'Rahasia#2026'), zip)).toBe(true);
    expect(() => decryptPackage(encrypted, 'rahasia#2026')).toThrow(WrongPasswordError);

    const root = readCfb(encrypted);
    const pkg = root.streams.get('EncryptedPackage')!.slice();
    pkg[100]! ^= 1;
    root.streams.set('EncryptedPackage', pkg);
    expect(() => decryptPackage(writeCfb(root), 'Rahasia#2026')).toThrow(/integrity/);
  });

  it('opens and saves password-protected documents through DocxDocument', async () => {
    const doc = await DocxDocument.open(await buildDocx({ body: p('Secret content') }));
    const encrypted = await doc.save({ password: 'pw' });
    await expect(DocxDocument.open(encrypted)).rejects.toThrow(PasswordRequiredError);
    await expect(DocxDocument.open(encrypted, { password: 'nope' })).rejects.toThrow(WrongPasswordError);
    const opened = await DocxDocument.open(encrypted, { password: 'pw' });
    expect(blocksToText(opened.body().blocks)).toBe('Secret content');
    expect(opened.sourcePassword).toBe('pw');
  });
});
