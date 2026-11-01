export { hash, hashConcat, hashLength, hashBlockSize, hmac, hashNameFromString, type HashName } from './sha.js';
export { Aes, aesEncryptEcb, aesDecryptEcb, aesEncryptCbc, aesDecryptCbc, type Padding } from './aes.js';

/** Cryptographically secure random bytes from the platform. */
export function randomBytes(n: number): Uint8Array {
  const out = new Uint8Array(n);
  globalThis.crypto.getRandomValues(out);
  return out;
}
