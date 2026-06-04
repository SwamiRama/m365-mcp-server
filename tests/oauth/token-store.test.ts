import { describe, it, expect, vi, beforeEach, afterEach } from 'vitest';

// Mock redis to return null (uses MemoryTokenStore)
vi.mock('../../src/utils/redis.js', () => ({
  getRedisClient: () => null,
}));

import {
  createRefreshToken,
  validateRefreshToken,
  rotateRefreshToken,
} from '../../src/oauth/token-store.js';

const params = (sessionId: string) => ({
  clientId: 'client-1',
  sessionId,
  userId: 'user-1',
  scope: 'openid',
});

describe('token-store rotation grace period', () => {
  beforeEach(() => {
    vi.useFakeTimers();
  });

  afterEach(() => {
    vi.useRealTimers();
  });

  it('validates a freshly created token', async () => {
    const token = await createRefreshToken(params('s-fresh'));
    const record = await validateRefreshToken(token);
    expect(record).not.toBeNull();
    expect(record?.sessionId).toBe('s-fresh');
  });

  it('keeps the rotated (old) token valid within the grace period', async () => {
    const oldToken = await createRefreshToken(params('s-grace'));
    const newToken = await rotateRefreshToken(oldToken, params('s-grace'));
    expect(newToken).not.toBeNull();

    // 30s later: still within the 60s default grace window
    vi.advanceTimersByTime(30_000);
    const record = await validateRefreshToken(oldToken);
    expect(record).not.toBeNull();
    expect(record?.sessionId).toBe('s-grace');
  });

  it('allows a second rotation of the same old token within grace (concurrent refresh race)', async () => {
    const oldToken = await createRefreshToken(params('s-race'));
    const tokenA = await rotateRefreshToken(oldToken, params('s-race'));
    const tokenB = await rotateRefreshToken(oldToken, params('s-race'));

    expect(tokenA).not.toBeNull();
    expect(tokenB).not.toBeNull();
    expect(await validateRefreshToken(tokenA!)).not.toBeNull();
    expect(await validateRefreshToken(tokenB!)).not.toBeNull();
  });

  it('rejects the old token after the grace period and revokes the whole token family', async () => {
    const oldToken = await createRefreshToken(params('s-reuse'));
    const newToken = await rotateRefreshToken(oldToken, params('s-reuse'));
    expect(newToken).not.toBeNull();

    // 61s later: grace expired, reuse = theft indicator
    vi.advanceTimersByTime(61_000);
    expect(await validateRefreshToken(oldToken)).toBeNull();
    // family revocation: the successor token must be dead too
    expect(await validateRefreshToken(newToken!)).toBeNull();
  });

  it('forgets the rotated token after the retention window without revoking the successor', async () => {
    const oldToken = await createRefreshToken(params('s-retention'));
    const newToken = await rotateRefreshToken(oldToken, params('s-retention'));

    // 12min later (> 60s grace + 600s retention): tombstone evicted,
    // old token unknown -> plain invalid_grant
    vi.advanceTimersByTime(12 * 60_000);
    expect(await validateRefreshToken(oldToken)).toBeNull();
    // successor stays valid (reuse can no longer be distinguished from garbage)
    expect(await validateRefreshToken(newToken!)).not.toBeNull();
  });

  it('detects reuse of a rotated token even after its hard expiry and revokes the family', async () => {
    vi.resetModules();
    process.env['OAUTH_REFRESH_TOKEN_LIFETIME_SECS'] = '1000';
    process.env['SESSION_TTL_SECONDS'] = '1000';
    const fresh = await import('../../src/oauth/token-store.js');

    const oldToken = await fresh.createRefreshToken(params('s-expired-reuse'));
    // rotate shortly before the old token's hard expiry (t=500s of 1000s)
    vi.advanceTimersByTime(500_000);
    const newToken = await fresh.rotateRefreshToken(oldToken, params('s-expired-reuse'));

    // t=1100s: old token past expiresAt (1000s) but tombstone still present
    vi.advanceTimersByTime(600_000);
    expect(await fresh.validateRefreshToken(oldToken)).toBeNull();
    // reuse-after-grace must win over plain expiry: family is revoked
    expect(await fresh.validateRefreshToken(newToken!)).toBeNull();

    delete process.env['OAUTH_REFRESH_TOKEN_LIFETIME_SECS'];
    delete process.env['SESSION_TTL_SECONDS'];
  });

  it('does not extend the grace window on repeated rotation of the same old token', async () => {
    const oldToken = await createRefreshToken(params('s-noextend'));
    await rotateRefreshToken(oldToken, params('s-noextend'));

    vi.advanceTimersByTime(40_000);
    // second racer rotates again at t=40s - must NOT restart the 60s grace clock
    await rotateRefreshToken(oldToken, params('s-noextend'));

    vi.advanceTimersByTime(30_000);
    // t=70s after first rotation: grace (60s) is over even though last rotation was 30s ago
    expect(await validateRefreshToken(oldToken)).toBeNull();
  });
});
