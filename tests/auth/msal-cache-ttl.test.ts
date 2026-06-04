import { describe, it, expect, vi, beforeEach, afterEach } from 'vitest';

const redisMock = vi.hoisted(() => ({
  get: vi.fn(async () => null),
  setex: vi.fn(async () => 'OK'),
  expire: vi.fn(async () => 1),
}));

vi.mock('../../src/utils/redis.js', () => ({
  getRedisClient: () => redisMock,
}));

describe('MSAL cache TTL follows session TTL config', () => {
  beforeEach(() => {
    vi.resetModules();
    vi.clearAllMocks();
  });

  afterEach(() => {
    delete process.env['SESSION_TTL_SECONDS'];
    delete process.env['OAUTH_REFRESH_TOKEN_LIFETIME_SECS'];
  });

  it('writes cache entries with the configured session TTL', async () => {
    process.env['SESSION_TTL_SECONDS'] = '7777';
    process.env['OAUTH_REFRESH_TOKEN_LIFETIME_SECS'] = '7777';
    const { createMsalCachePlugin } = await import('../../src/auth/msal-cache-plugin.js');

    const plugin = createMsalCachePlugin('part-1');
    await plugin.afterCacheAccess({
      cacheHasChanged: true,
      tokenCache: { serialize: () => 'cache-data', deserialize: vi.fn() },
    } as never);

    // payload is encrypted at rest - this test only cares about the TTL
    expect(redisMock.setex).toHaveBeenCalledWith(
      'm365-mcp:msal-cache:part-1',
      7777,
      expect.any(String)
    );
  });

  it('touches cache entries with the configured session TTL', async () => {
    process.env['SESSION_TTL_SECONDS'] = '7777';
    process.env['OAUTH_REFRESH_TOKEN_LIFETIME_SECS'] = '7777';
    const { touchMsalCache } = await import('../../src/auth/msal-cache-plugin.js');

    await touchMsalCache('part-2');

    expect(redisMock.expire).toHaveBeenCalledWith('m365-mcp:msal-cache:part-2', 7777);
  });
});
