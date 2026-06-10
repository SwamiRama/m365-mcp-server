/**
 * Shared AES-256-GCM encryption for data at rest (sessions, MSAL cache).
 *
 * Each ciphertext carries its own random scrypt salt: a fixed, public salt
 * let every instance with the same SESSION_SECRET derive identical
 * key material and enabled precomputation. New output uses the format
 * `salt:iv:tag:ciphertext`. Legacy entries written before this change used a
 * fixed salt and the 3-part format `iv:tag:ciphertext`; those are still
 * decryptable (LEGACY_SALT fallback) so existing Redis data does not break.
 */

import crypto from 'crypto';
import { config } from './config.js';

const ALGORITHM = 'aes-256-gcm';
const IV_LENGTH = 12;
const SALT_LENGTH = 16;
// Fixed salt used by older releases — retained ONLY to decrypt pre-existing entries.
const LEGACY_SALT = 'm365-mcp-salt';

function deriveKey(secret: string, salt: crypto.BinaryLike): Buffer {
  return crypto.scryptSync(secret, salt, 32);
}

export function encrypt(data: string): string {
  const salt = crypto.randomBytes(SALT_LENGTH);
  const key = deriveKey(config.sessionSecret, salt);
  const iv = crypto.randomBytes(IV_LENGTH);
  const cipher = crypto.createCipheriv(ALGORITHM, key, iv);

  let encrypted = cipher.update(data, 'utf8', 'base64');
  encrypted += cipher.final('base64');

  const tag = cipher.getAuthTag();

  // Format: salt:iv:tag:encrypted
  return `${salt.toString('base64')}:${iv.toString('base64')}:${tag.toString('base64')}:${encrypted}`;
}

export function decrypt(encryptedData: string): string {
  const parts = encryptedData.split(':');

  let salt: crypto.BinaryLike;
  let ivStr: string | undefined;
  let tagStr: string | undefined;
  let encrypted: string | undefined;

  if (parts.length === 4) {
    // New format: salt:iv:tag:encrypted
    salt = Buffer.from(parts[0]!, 'base64');
    [, ivStr, tagStr, encrypted] = parts;
  } else if (parts.length === 3) {
    // Legacy format (older releases): iv:tag:encrypted derived from the fixed salt
    salt = LEGACY_SALT;
    [ivStr, tagStr, encrypted] = parts;
  } else {
    throw new Error('Invalid encrypted data format');
  }

  if (!ivStr || !tagStr || !encrypted) {
    throw new Error('Invalid encrypted data format');
  }

  const key = deriveKey(config.sessionSecret, salt);
  const iv = Buffer.from(ivStr, 'base64');
  const tag = Buffer.from(tagStr, 'base64');

  const decipher = crypto.createDecipheriv(ALGORITHM, key, iv);
  decipher.setAuthTag(tag);

  let decrypted = decipher.update(encrypted, 'base64', 'utf8');
  decrypted += decipher.final('utf8');

  return decrypted;
}
