import { createCipheriv, createHash, createHmac, randomBytes as nodeRandom } from 'node:crypto';
import { describe, expect, it } from 'vitest';
import { aesDecryptCbc, aesDecryptEcb, aesEncryptCbc, aesEncryptEcb, hash, hmac, type HashName } from '../src/crypto/index.js';
import { fromHex, toHex } from '../src/util/bytes.js';
import { utf8Encode } from '../src/util/text.js';

const NODE_NAMES: Record<HashName, string> = { 'SHA-1': 'sha1', 'SHA-256': 'sha256', 'SHA-384': 'sha384', 'SHA-512': 'sha512' };

describe('SHA', () => {
  it('matches FIPS 180 test vectors for "abc"', () => {
    const abc = utf8Encode('abc');
    expect(toHex(hash('SHA-1', abc))).toBe('a9993e364706816aba3e25717850c26c9cd0d89d');
    expect(toHex(hash('SHA-256', abc))).toBe('ba7816bf8f01cfea414140de5dae2223b00361a396177a9cb410ff61f20015ad');
    expect(toHex(hash('SHA-384', abc))).toBe('cb00753f45a35e8bb5a03d699ac65007272c32ab0eded1631a8b605a43ff5bed8086072ba1e7cc2358baeca134c825a7');
    expect(toHex(hash('SHA-512', abc))).toBe(
      'ddaf35a193617abacc417349ae20413112e6fa4e89a97ea20a9eeee64b55d39a2192992a274fc1a836ba3c23a3feebbd454d4423643ce80e2a9ac94fa54ca49f',
    );
  });

  it('matches the platform implementation for many lengths', () => {
    for (const name of Object.keys(NODE_NAMES) as HashName[]) {
      for (const len of [0, 1, 55, 56, 63, 64, 65, 111, 112, 127, 128, 129, 1000, 4097]) {
        const data = nodeRandom(len);
        expect(toHex(hash(name, data)), `${name} len ${len}`).toBe(createHash(NODE_NAMES[name]).update(data).digest('hex'));
      }
    }
  });

  it('HMAC matches RFC 4231 case 2 and the platform', () => {
    const key = utf8Encode('Jefe');
    const data = utf8Encode('what do ya want for nothing?');
    expect(toHex(hmac('SHA-256', key, data))).toBe('5bdcc146bf60754e6a042426089575c75a003f089d2739839dec58b964ec3843');
    for (const name of Object.keys(NODE_NAMES) as HashName[]) {
      const k = nodeRandom(200);
      const d = nodeRandom(300);
      expect(toHex(hmac(name, k, d))).toBe(createHmac(NODE_NAMES[name], k).update(d).digest('hex'));
    }
  });
});

describe('AES', () => {
  it('matches FIPS 197 appendix C', () => {
    const pt = fromHex('00112233445566778899aabbccddeeff');
    const cases: Array<[string, string]> = [
      ['000102030405060708090a0b0c0d0e0f', '69c4e0d86a7b0430d8cdb78070b4c55a'],
      ['000102030405060708090a0b0c0d0e0f1011121314151617', 'dda97ca4864cdfe06eaf70a0ec0d7191'],
      ['000102030405060708090a0b0c0d0e0f101112131415161718191a1b1c1d1e1f', '8ea2b7ca516745bfeafc49904b496089'],
    ];
    for (const [key, ct] of cases) {
      expect(toHex(aesEncryptEcb(fromHex(key), pt))).toBe(ct);
      expect(toHex(aesDecryptEcb(fromHex(key), fromHex(ct)))).toBe(toHex(pt));
    }
  });

  it('matches NIST SP 800-38A CBC example', () => {
    const key = fromHex('2b7e151628aed2a6abf7158809cf4f3c');
    const iv = fromHex('000102030405060708090a0b0c0d0e0f');
    const pt = fromHex('6bc1bee22e409f96e93d7e117393172aae2d8a571e03ac9c9eb76fac45af8e51');
    expect(toHex(aesEncryptCbc(key, iv, pt))).toBe('7649abac8119b246cee98e9b12e9197d5086cb9b507219ee95db113a917678b2');
  });

  it('matches the platform for CBC with and without padding', () => {
    for (const keyLen of [16, 24, 32]) {
      for (const len of [0, 1, 15, 16, 17, 100, 4096]) {
        const key = nodeRandom(keyLen);
        const iv = nodeRandom(16);
        const data = nodeRandom(len);
        const cipher = createCipheriv(`aes-${keyLen * 8}-cbc`, key, iv);
        const expected = Buffer.concat([cipher.update(data), cipher.final()]);
        const ours = aesEncryptCbc(key, iv, data, 'pkcs7');
        expect(toHex(ours)).toBe(expected.toString('hex'));
        expect(toHex(aesDecryptCbc(key, iv, ours, 'pkcs7'))).toBe(toHex(data));
        if (len % 16 === 0) {
          const raw = aesEncryptCbc(key, iv, data, 'none');
          expect(toHex(aesDecryptCbc(key, iv, raw, 'none'))).toBe(toHex(data));
        }
      }
    }
  });

  it('runs 100,000 SHA-512 rounds quickly', () => {
    const start = performance.now();
    let h = hash('SHA-512', utf8Encode('password'));
    const buf = new Uint8Array(68);
    for (let i = 0; i < 100_000; i++) {
      new DataView(buf.buffer).setUint32(0, i, true);
      buf.set(h, 4);
      h = hash('SHA-512', buf);
    }
    expect(performance.now() - start).toBeLessThan(5000);
  });
});
