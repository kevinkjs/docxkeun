/**
 * Password-to-open encryption for Office files (MS-OFFCRYPTO).
 *
 * Writes Agile encryption (AES-256-CBC, SHA-512, 100,000 rounds, HMAC data integrity) in a
 * compound file with the DataSpaces streams Word requires. Reads Agile and Standard encryption.
 */
import { aesDecryptCbc, aesDecryptEcb, aesEncryptCbc, hash, hashConcat, hashNameFromString, hmac, randomBytes, type HashName } from '../crypto/index.js';
import { ByteReader, ByteWriter, timingSafeEqual } from '../util/bytes.js';
import { base64Decode, base64Encode, utf16leEncode, utf8Decode, utf8Encode } from '../util/text.js';
import { attr, findAll, parseXml, type XmlElement } from '../xml/index.js';
import { emptyStorage, readCfb, writeCfb, type CfbStorage } from './cfb.js';

export class EncryptionError extends Error {
  override name = 'EncryptionError';
}

export class WrongPasswordError extends EncryptionError {
  override name = 'WrongPasswordError';
  constructor() {
    super('The password is incorrect');
  }
}

const BLOCK = {
  verifierInput: Uint8Array.from([0xfe, 0xa7, 0xd2, 0x76, 0x3b, 0x4b, 0x9e, 0x79]),
  verifierValue: Uint8Array.from([0xd7, 0xaa, 0x0f, 0x6d, 0x30, 0x61, 0x34, 0x4e]),
  keyValue: Uint8Array.from([0x14, 0x6e, 0x0b, 0xe7, 0xab, 0xac, 0xd0, 0xd6]),
  hmacKey: Uint8Array.from([0x5f, 0xb2, 0xad, 0x01, 0x0c, 0xb9, 0xe1, 0xf6]),
  hmacValue: Uint8Array.from([0xa0, 0x67, 0x7f, 0x02, 0xb2, 0x2c, 0x84, 0x33]),
};

const SEGMENT = 4096;

function le32(n: number): Uint8Array {
  const b = new Uint8Array(4);
  new DataView(b.buffer).setUint32(0, n, true);
  return b;
}

/** Truncate or pad (with 0x36) to a length, as MS-OFFCRYPTO does for keys and IVs. */
function fit(bytes: Uint8Array, length: number): Uint8Array {
  if (bytes.length >= length) return bytes.slice(0, length);
  const out = new Uint8Array(length).fill(0x36);
  out.set(bytes);
  return out;
}

function padToBlock(data: Uint8Array, block: number): Uint8Array {
  if (data.length % block === 0) return data;
  const out = new Uint8Array(Math.ceil(data.length / block) * block);
  out.set(data);
  return out;
}

/** H(salt || UTF-16LE password), then spinCount rounds of H(LE32 i || H). */
export function agileHash(algorithm: HashName, salt: Uint8Array, password: string, spinCount: number): Uint8Array {
  let h = hashConcat(algorithm, salt, utf16leEncode(password));
  const buf = new Uint8Array(4 + h.length);
  const view = new DataView(buf.buffer);
  for (let i = 0; i < spinCount; i++) {
    view.setUint32(0, i, true);
    buf.set(h, 4);
    h = hash(algorithm, buf);
  }
  return h;
}

interface CipherParams {
  saltValue: Uint8Array;
  blockSize: number;
  keyBits: number;
  hashSize: number;
  hashAlgorithm: HashName;
  cipherAlgorithm: string;
  cipherChaining: string;
}

function readParams(node: XmlElement | undefined): CipherParams {
  if (!node) throw new EncryptionError('Missing encryption parameters');
  const hashAlgorithm = hashNameFromString(attr(node, 'hashAlgorithm') ?? '');
  if (!hashAlgorithm) throw new EncryptionError(`Unsupported hash ${attr(node, 'hashAlgorithm')}`);
  const cipherAlgorithm = attr(node, 'cipherAlgorithm') ?? '';
  const cipherChaining = attr(node, 'cipherChaining') ?? '';
  if (cipherAlgorithm !== 'AES' || cipherChaining !== 'ChainingModeCBC') throw new EncryptionError(`Unsupported cipher ${cipherAlgorithm}/${cipherChaining}`);
  return {
    saltValue: base64Decode(attr(node, 'saltValue') ?? ''),
    blockSize: Number(attr(node, 'blockSize') ?? 16),
    keyBits: Number(attr(node, 'keyBits') ?? 256),
    hashSize: Number(attr(node, 'hashSize') ?? 64),
    hashAlgorithm,
    cipherAlgorithm,
    cipherChaining,
  };
}

// ---------------------------------------------------------------------------------------------
// Decryption

export function isEncryptedPackage(bytes: Uint8Array): boolean {
  try {
    const cfb = readCfb(bytes);
    return cfb.streams.has('EncryptionInfo') && cfb.streams.has('EncryptedPackage');
  } catch {
    return false;
  }
}

/** Decrypt an encrypted Office file to its zip package. Throws WrongPasswordError. */
export function decryptPackage(file: Uint8Array, password: string): Uint8Array {
  const cfb = readCfb(file);
  const info = cfb.streams.get('EncryptionInfo');
  const pkg = cfb.streams.get('EncryptedPackage');
  if (!info || !pkg) throw new EncryptionError('This compound file is not an encrypted Office document');
  const r = new ByteReader(info);
  const major = r.u16le();
  const minor = r.u16le();
  if (major === 4 && minor === 4) return decryptAgile(info.subarray(8), pkg, password);
  if ((major === 3 || major === 4) && minor === 2) return decryptStandard(info, pkg, password);
  throw new EncryptionError(`Unsupported encryption version ${major}.${minor}`);
}

function decryptAgile(xmlBytes: Uint8Array, pkg: Uint8Array, password: string): Uint8Array {
  const xml = parseXml(utf8Decode(xmlBytes), { canonicalize: false });
  const root = xml.root;
  const local = (n: string) => n.replace(/^.*:/, '');
  const child = (parent: typeof root, name: string) => parent.children.find((c) => c.type === 'element' && local(c.name) === name) as typeof root | undefined;
  const keyData = readParams(child(root, 'keyData'));
  const encryptor = findAll(root, 'p:encryptedKey')[0] ?? (child(child(child(root, 'keyEncryptors') ?? root, 'keyEncryptor') ?? root, 'encryptedKey'));
  if (!encryptor) throw new EncryptionError('No password key encryptor (certificate-only encryption is not supported)');
  const p = readParams(encryptor);
  const spinCount = Number(attr(encryptor, 'spinCount') ?? 100000);
  const keyLength = p.keyBits / 8;
  const h = agileHash(p.hashAlgorithm, p.saltValue, password, spinCount);
  const derive = (block: Uint8Array) => fit(hashConcat(p.hashAlgorithm, h, block), keyLength);
  const iv = fit(p.saltValue, p.blockSize);

  const verifierInput = aesDecryptCbc(derive(BLOCK.verifierInput), iv, base64Decode(attr(encryptor, 'encryptedVerifierHashInput') ?? ''));
  const verifierHash = aesDecryptCbc(derive(BLOCK.verifierValue), iv, base64Decode(attr(encryptor, 'encryptedVerifierHashValue') ?? ''));
  const expected = hash(p.hashAlgorithm, verifierInput.subarray(0, p.saltValue.length));
  if (!timingSafeEqual(expected, verifierHash.subarray(0, expected.length))) throw new WrongPasswordError();
  const secretKey = aesDecryptCbc(derive(BLOCK.keyValue), iv, base64Decode(attr(encryptor, 'encryptedKeyValue') ?? '')).subarray(0, keyData.keyBits / 8);

  // Data integrity: HMAC over the whole EncryptedPackage stream, checked before decrypting.
  const integrity = child(root, 'dataIntegrity');
  if (integrity) {
    const ivKey = fit(hashConcat(keyData.hashAlgorithm, keyData.saltValue, BLOCK.hmacKey), keyData.blockSize);
    const ivValue = fit(hashConcat(keyData.hashAlgorithm, keyData.saltValue, BLOCK.hmacValue), keyData.blockSize);
    const hmacKey = aesDecryptCbc(secretKey, ivKey, base64Decode(attr(integrity, 'encryptedHmacKey') ?? '')).subarray(0, keyData.hashSize);
    const hmacValue = aesDecryptCbc(secretKey, ivValue, base64Decode(attr(integrity, 'encryptedHmacValue') ?? '')).subarray(0, keyData.hashSize);
    if (!timingSafeEqual(hmac(keyData.hashAlgorithm, hmacKey, pkg), hmacValue)) throw new EncryptionError('The encrypted package failed its integrity check (file damaged or tampered with)');
  }

  const size = new ByteReader(pkg).u64le();
  const out = new Uint8Array(Math.ceil((pkg.length - 8) / keyData.blockSize) * keyData.blockSize);
  for (let i = 0, off = 8; off < pkg.length; i++, off += SEGMENT) {
    const segment = pkg.subarray(off, Math.min(off + SEGMENT, pkg.length));
    const segIv = fit(hashConcat(keyData.hashAlgorithm, keyData.saltValue, le32(i)), keyData.blockSize);
    out.set(aesDecryptCbc(secretKey, segIv, padToBlock(segment, keyData.blockSize)), off - 8);
  }
  return out.subarray(0, size);
}

/** ECMA-376 Standard encryption (Office 2007): AES-ECB with a SHA-1 key derivation. */
function decryptStandard(info: Uint8Array, pkg: Uint8Array, password: string): Uint8Array {
  const r = new ByteReader(info, 8);
  const headerSize = r.u32le();
  const header = new ByteReader(info, r.pos);
  header.skip(8);
  const algId = header.u32le();
  header.skip(4);
  const keyBits = header.u32le();
  if (algId !== 0x660e && algId !== 0x660f && algId !== 0x6610) throw new EncryptionError('Only AES Standard encryption is supported');
  r.skip(headerSize);
  const saltSize = r.u32le();
  const salt = r.bytesAt(saltSize);
  const encryptedVerifier = r.bytesAt(16);
  r.skip(4);
  const encryptedVerifierHash = r.bytesAt(32);

  let h = hashConcat('SHA-1', salt, utf16leEncode(password));
  for (let i = 0; i < 50000; i++) h = hashConcat('SHA-1', le32(i), h);
  const hFinal = hashConcat('SHA-1', h, le32(0));
  const x = (pad: number) => {
    const buf = new Uint8Array(64).fill(pad);
    for (let i = 0; i < hFinal.length; i++) buf[i]! ^= hFinal[i]!;
    return hash('SHA-1', buf);
  };
  const key = new Uint8Array([...x(0x36), ...x(0x5c)]).subarray(0, keyBits / 8);
  const verifier = aesDecryptEcb(key, encryptedVerifier);
  const verifierHash = aesDecryptEcb(key, encryptedVerifierHash).subarray(0, 20);
  if (!timingSafeEqual(hash('SHA-1', verifier), verifierHash)) throw new WrongPasswordError();
  const size = new ByteReader(pkg).u64le();
  return aesDecryptEcb(key, padToBlock(pkg.subarray(8), 16)).subarray(0, size);
}

// ---------------------------------------------------------------------------------------------
// Encryption

export interface EncryptOptions {
  spinCount?: number;
}

/** Encrypt a .docx (zip bytes) with a password, producing a file Word asks a password for. */
export function encryptPackage(zip: Uint8Array, password: string, options: EncryptOptions = {}): Uint8Array {
  if (!password) throw new EncryptionError('A password is required');
  const spinCount = options.spinCount ?? 100_000;
  const algorithm: HashName = 'SHA-512';
  const keyBytes = 32;
  const blockSize = 16;
  const keyDataSalt = randomBytes(16);
  const passwordSalt = randomBytes(16);
  const secretKey = randomBytes(keyBytes);

  // Encrypted package: size prefix + 4096-byte segments, each with its own IV.
  const w = new ByteWriter(zip.length + 64);
  w.u64le(zip.length);
  for (let i = 0, off = 0; off < zip.length; i++, off += SEGMENT) {
    const segment = padToBlock(zip.subarray(off, Math.min(off + SEGMENT, zip.length)), blockSize);
    const iv = fit(hashConcat(algorithm, keyDataSalt, le32(i)), blockSize);
    w.bytes(aesEncryptCbc(secretKey, iv, segment));
  }
  const encryptedPackage = w.toBytes();

  // Password key encryptor.
  const h = agileHash(algorithm, passwordSalt, password, spinCount);
  const derive = (block: Uint8Array) => fit(hashConcat(algorithm, h, block), keyBytes);
  const iv = fit(passwordSalt, blockSize);
  const verifierInput = randomBytes(16);
  const encryptedVerifierHashInput = aesEncryptCbc(derive(BLOCK.verifierInput), iv, verifierInput);
  const encryptedVerifierHashValue = aesEncryptCbc(derive(BLOCK.verifierValue), iv, padToBlock(hash(algorithm, verifierInput), blockSize));
  const encryptedKeyValue = aesEncryptCbc(derive(BLOCK.keyValue), iv, secretKey);

  // Data integrity.
  const hmacKey = randomBytes(64);
  const hmacValue = hmac(algorithm, hmacKey, encryptedPackage);
  const encryptedHmacKey = aesEncryptCbc(secretKey, fit(hashConcat(algorithm, keyDataSalt, BLOCK.hmacKey), blockSize), hmacKey);
  const encryptedHmacValue = aesEncryptCbc(secretKey, fit(hashConcat(algorithm, keyDataSalt, BLOCK.hmacValue), blockSize), hmacValue);

  const b64 = base64Encode;
  const xml =
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n' +
    '<encryption xmlns="http://schemas.microsoft.com/office/2006/encryption" xmlns:p="http://schemas.microsoft.com/office/2006/keyEncryptor/password" xmlns:c="http://schemas.microsoft.com/office/2006/keyEncryptor/certificate">' +
    `<keyData saltSize="16" blockSize="16" keyBits="256" hashSize="64" cipherAlgorithm="AES" cipherChaining="ChainingModeCBC" hashAlgorithm="SHA512" saltValue="${b64(keyDataSalt)}"/>` +
    `<dataIntegrity encryptedHmacKey="${b64(encryptedHmacKey)}" encryptedHmacValue="${b64(encryptedHmacValue)}"/>` +
    '<keyEncryptors><keyEncryptor uri="http://schemas.microsoft.com/office/2006/keyEncryptor/password">' +
    `<p:encryptedKey spinCount="${spinCount}" saltSize="16" blockSize="16" keyBits="256" hashSize="64" cipherAlgorithm="AES" cipherChaining="ChainingModeCBC" hashAlgorithm="SHA512" saltValue="${b64(passwordSalt)}" encryptedVerifierHashInput="${b64(encryptedVerifierHashInput)}" encryptedVerifierHashValue="${b64(encryptedVerifierHashValue)}" encryptedKeyValue="${b64(encryptedKeyValue)}"/>` +
    '</keyEncryptor></keyEncryptors></encryption>';
  const info = new ByteWriter(xml.length + 8);
  info.u16le(4).u16le(4).u32le(0x40).bytes(utf8Encode(xml));

  const root = emptyStorage();
  root.streams.set('EncryptionInfo', info.toBytes());
  root.streams.set('EncryptedPackage', encryptedPackage);
  root.storages.set('\u0006DataSpaces', dataSpaces());
  return writeCfb(root);
}

/** Length-prefixed UTF-16LE string padded to 4 bytes (UNICODE-LP-P4). */
function unicodeLp4(w: ByteWriter, s: string): void {
  const bytes = utf16leEncode(s);
  w.u32le(bytes.length).bytes(bytes).zeros((4 - (bytes.length % 4)) % 4);
}

function version(w: ByteWriter): void {
  w.u16le(1).u16le(0);
}

/** The \x06DataSpaces storage that marks EncryptedPackage as using the encryption transform. */
function dataSpaces(): CfbStorage {
  const ds = emptyStorage();

  const v = new ByteWriter(128);
  unicodeLp4(v, 'Microsoft.Container.DataSpaces');
  version(v);
  version(v);
  version(v);
  ds.streams.set('Version', v.toBytes());

  const entry = new ByteWriter(128);
  entry.u32le(1); // reference component count
  entry.u32le(0); // component type: stream
  unicodeLp4(entry, 'EncryptedPackage');
  unicodeLp4(entry, 'StrongEncryptionDataSpace');
  const entryBytes = entry.toBytes();
  const map = new ByteWriter(128);
  map.u32le(8).u32le(1).u32le(entryBytes.length + 4).bytes(entryBytes);
  ds.streams.set('DataSpaceMap', map.toBytes());

  const def = new ByteWriter(64);
  def.u32le(8).u32le(1);
  unicodeLp4(def, 'StrongEncryptionTransform');
  const info = emptyStorage();
  info.streams.set('StrongEncryptionDataSpace', def.toBytes());
  ds.storages.set('DataSpaceInfo', info);

  const primary = new ByteWriter(256);
  const idWriter = new ByteWriter(96);
  unicodeLp4(idWriter, '{FF9A3F03-56EF-4613-BDD5-5A41C1D07246}');
  const idBytes = idWriter.toBytes();
  primary.u32le(8 + idBytes.length).u32le(1).bytes(idBytes);
  unicodeLp4(primary, 'Microsoft.Container.EncryptionTransform');
  version(primary);
  version(primary);
  version(primary);
  primary.u32le(0); // EncryptionName (empty UTF-8-LP-P4)
  primary.u32le(0); // EncryptionBlockSize
  primary.u32le(0); // CipherMode
  primary.u32le(4); // Reserved
  const transform = emptyStorage();
  transform.streams.set('\u0006Primary', primary.toBytes());
  const transformInfo = emptyStorage();
  transformInfo.storages.set('StrongEncryptionTransform', transform);
  ds.storages.set('TransformInfo', transformInfo);
  return ds;
}

