export {
  readProtection,
  setProtection,
  removeProtection,
  checkProtectionPassword,
  setWriteProtection,
  readWriteProtection,
  checkWriteProtectionPassword,
  verifyProtectionPassword,
  legacyHash,
  legacyVerifier,
  legacyHashHexReversed,
  iteratedHash,
  type ProtectionMode,
  type ProtectionInfo,
} from './protection.js';
export { AccessPolicy, PermissionError, type Role, type AccessUser, type Capability, type AccessPolicyOptions } from './policy.js';
export { readCfb, writeCfb, emptyStorage, CfbError, type CfbStorage } from './cfb.js';
export { encryptPackage, decryptPackage, isEncryptedPackage, agileHash, EncryptionError, WrongPasswordError, type EncryptOptions } from './encryption.js';

import { decryptionHooks } from '../model/document.js';
import { decryptPackage, encryptPackage, EncryptionError } from './encryption.js';

/** Thrown by DocxDocument.open for an encrypted file opened without a password. */
export class PasswordRequiredError extends EncryptionError {
  override name = 'PasswordRequiredError';
  constructor() {
    super('This document is password protected');
  }
}

// Importing docxkit/security lets DocxDocument.open / save handle password-to-open files.
decryptionHooks.decrypt = async (bytes, password) => {
  if (password === undefined || password === '') throw new PasswordRequiredError();
  return decryptPackage(bytes, password);
};
decryptionHooks.encrypt = async (zip, password) => encryptPackage(zip, password);
export { pdfSecurity, permissionBits, hashR6, type PdfSecurityOptions, type PdfPermissions } from './pdf.js';
