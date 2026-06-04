import { describe, it, expect, vi, beforeEach, afterEach } from 'vitest';

// Mock redis to return null (uses MemorySessionStore)
vi.mock('../../src/utils/redis.js', () => ({
  getRedisClient: () => null,
}));

describe('session TTL configuration', () => {
  beforeEach(() => {
    vi.resetModules();
    vi.useFakeTimers();
  });

  afterEach(() => {
    vi.useRealTimers();
    delete process.env['SESSION_TTL_SECONDS'];
    delete process.env['OAUTH_REFRESH_TOKEN_LIFETIME_SECS'];
  });

  it('exposes sessionTtlSecs with a 24h default', async () => {
    const { config } = await import('../../src/utils/config.js');
    expect(config.sessionTtlSecs).toBe(86400);
  });

  it('reads SESSION_TTL_SECONDS from the environment', async () => {
    process.env['SESSION_TTL_SECONDS'] = '2592000';
    process.env['OAUTH_REFRESH_TOKEN_LIFETIME_SECS'] = '2592000';
    const { config } = await import('../../src/utils/config.js');
    expect(config.sessionTtlSecs).toBe(2592000);
  });

  it('rejects a session TTL shorter than the refresh token lifetime', async () => {
    process.env['SESSION_TTL_SECONDS'] = '3600';
    process.env['OAUTH_REFRESH_TOKEN_LIFETIME_SECS'] = '86400';
    await expect(import('../../src/utils/config.js')).rejects.toThrow(
      /SESSION_TTL_SECONDS/
    );
  });

  it('expires sessions after the configured TTL', async () => {
    process.env['SESSION_TTL_SECONDS'] = '120';
    process.env['OAUTH_REFRESH_TOKEN_LIFETIME_SECS'] = '120';
    const { sessionManager } = await import('../../src/auth/session.js');

    const session = await sessionManager.createSession();
    vi.advanceTimersByTime(121_000);
    expect(await sessionManager.getSession(session.id)).toBeNull();
  });

  it('slides the session TTL on access', async () => {
    process.env['SESSION_TTL_SECONDS'] = '120';
    process.env['OAUTH_REFRESH_TOKEN_LIFETIME_SECS'] = '120';
    const { sessionManager } = await import('../../src/auth/session.js');

    const session = await sessionManager.createSession();
    vi.advanceTimersByTime(100_000);
    expect(await sessionManager.getSession(session.id)).not.toBeNull(); // slides
    vi.advanceTimersByTime(100_000);
    // 200s total, but only 100s since last access -> still alive
    expect(await sessionManager.getSession(session.id)).not.toBeNull();
  });
});
