/**
 * Shared AES-256-GCM encryption for data at rest (sessions, MSAL cache).
 * Key derived from SESSION_SECRET via scrypt - parameters must not change,
 * existing Redis entries would become undecryptable.
 */

import crypto from 'crypto';
import { config } from './config.js';

const ALGORITHM = 'aes-256-gcm';
const IV_LENGTH = 12;

function deriveKey(secret: string): Buffer {
  return crypto.scryptSync(secret, 'm365-mcp-salt', 32);
}

export function encrypt(data: string): string {
  const key = deriveKey(config.sessionSecret);
  const iv = crypto.randomBytes(IV_LENGTH);
  const cipher = crypto.createCipheriv(ALGORITHM, key, iv);

  let encrypted = cipher.update(data, 'utf8', 'base64');
  encrypted += cipher.final('base64');

  const tag = cipher.getAuthTag();

  // Format: iv:tag:encrypted
  return `${iv.toString('base64')}:${tag.toString('base64')}:${encrypted}`;
}

export function decrypt(encryptedData: string): string {
  const parts = encryptedData.split(':');
  if (parts.length !== 3) {
    throw new Error('Invalid encrypted data format');
  }

  const [ivStr, tagStr, encrypted] = parts;
  if (!ivStr || !tagStr || !encrypted) {
    throw new Error('Invalid encrypted data format');
  }

  const key = deriveKey(config.sessionSecret);
  const iv = Buffer.from(ivStr, 'base64');
  const tag = Buffer.from(tagStr, 'base64');

  const decipher = crypto.createDecipheriv(ALGORITHM, key, iv);
  decipher.setAuthTag(tag);

  let decrypted = decipher.update(encrypted, 'base64', 'utf8');
  decrypted += decipher.final('utf8');

  return decrypted;
}
