import { describe, it, expect, vi } from 'vitest';
import crypto from 'node:crypto';

// NOTE: keep this literal in sync with the vi.mock factory below (the factory is
// hoisted and cannot reference module-scope variables).
const SESSION_SECRET = 'test-session-secret-0123456789-abc'; // >= 32 chars

vi.mock('../../src/utils/config.js', () => ({
  config: { sessionSecret: 'test-session-secret-0123456789-abc' },
}));

import { encrypt, decrypt } from '../../src/utils/crypto.js';

describe('crypto (AES-256-GCM, per-ciphertext salt)', () => {
  it('round-trips plaintext', () => {
    const plain = 'sensitive-access-token-value';
    expect(decrypt(encrypt(plain))).toBe(plain);
  });

  it('emits a 4-part salt:iv:tag:ciphertext format', () => {
    expect(encrypt('x').split(':')).toHaveLength(4);
  });

  it('uses a fresh random salt per call (defeats precomputation / shared-key derivation)', () => {
    const a = encrypt('same-plaintext');
    const b = encrypt('same-plaintext');
    // first segment is the salt in the new format
    expect(a.split(':')[0]).not.toBe(b.split(':')[0]);
    expect(decrypt(a)).toBe('same-plaintext');
    expect(decrypt(b)).toBe('same-plaintext');
  });

  it('still decrypts legacy 3-part entries (old fixed salt) for backward-compat', () => {
    // Reproduce a pre-fix ciphertext exactly: fixed salt 'm365-mcp-salt', format iv:tag:enc
    const plain = 'legacy-redis-token';
    const key = crypto.scryptSync(SESSION_SECRET, 'm365-mcp-salt', 32);
    const iv = crypto.randomBytes(12);
    const cipher = crypto.createCipheriv('aes-256-gcm', key, iv);
    let enc = cipher.update(plain, 'utf8', 'base64');
    enc += cipher.final('base64');
    const tag = cipher.getAuthTag();
    const legacy = `${iv.toString('base64')}:${tag.toString('base64')}:${enc}`;

    expect(legacy.split(':')).toHaveLength(3);
    expect(decrypt(legacy)).toBe(plain);
  });
});
