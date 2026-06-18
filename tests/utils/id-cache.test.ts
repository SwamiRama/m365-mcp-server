import { describe, it, expect, beforeEach } from 'vitest';
import { rememberId, resolveId, __clearIdCache } from '../../src/utils/id-cache.js';

describe('id-cache (Graph ID resolution for re-encoded relays)', () => {
  beforeEach(() => __clearIdCache());

  const realId = 'AAMkAGI2x-9_ab3cD=';     // base64url, as Graph returns it
  const reEncoded = 'AAMkAGI2x+9/ab3cD=';  // model turned -/_ into +/ (standard base64)

  it('resolves a re-encoded ID back to the remembered canonical ID', () => {
    rememberId('user-1', realId);
    expect(resolveId('user-1', reEncoded)).toBe(realId);
  });

  it('resolves the exact ID too (no-op when already canonical)', () => {
    rememberId('user-1', realId);
    expect(resolveId('user-1', realId)).toBe(realId);
  });

  it('returns the input unchanged when nothing matches', () => {
    expect(resolveId('user-1', reEncoded)).toBe(reEncoded);
  });

  it('is namespaced per user', () => {
    rememberId('user-1', realId);
    expect(resolveId('user-2', reEncoded)).toBe(reEncoded); // different user: no cross-match
  });

  it('passes undefined through', () => {
    expect(resolveId('user-1', undefined)).toBeUndefined();
  });

  it('does not store IDs that normalize to empty', () => {
    rememberId('user-1', '====');
    expect(resolveId('user-1', '----')).toBe('----');
  });
});
