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
