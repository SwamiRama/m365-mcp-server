import { describe, it, expect, vi, beforeEach } from 'vitest';

const redisMock = vi.hoisted(() => ({
  get: vi.fn(async (): Promise<string | null> => null),
  setex: vi.fn(async () => 'OK'),
  expire: vi.fn(async () => 1),
}));

vi.mock('../../src/utils/redis.js', () => ({
  getRedisClient: () => redisMock,
}));

import { createMsalCachePlugin } from '../../src/auth/msal-cache-plugin.js';

const MSAL_BLOB = '{"Account":{"uid.utid-login.windows.net-tenant":{"username":"a@b.com"}}}';

async function writeViaPlugin(partitionKey: string, data: string): Promise<string> {
  const plugin = createMsalCachePlugin(partitionKey);
  await plugin.afterCacheAccess({
    cacheHasChanged: true,
    tokenCache: { serialize: () => data, deserialize: vi.fn() },
  } as never);
  const call = redisMock.setex.mock.calls.at(-1);
  return call![2] as string;
}

describe('MSAL cache encryption at rest', () => {
  beforeEach(() => {
    vi.clearAllMocks();
  });

  it('does not write the MSAL blob as plaintext to redis', async () => {
    const stored = await writeViaPlugin('enc-1', MSAL_BLOB);
    expect(stored).not.toContain('username');
    expect(stored).not.toBe(MSAL_BLOB);
  });

  it('roundtrips the encrypted blob back to the original on read', async () => {
    const stored = await writeViaPlugin('enc-2', MSAL_BLOB);

    redisMock.get.mockResolvedValueOnce(stored);
    const deserialize = vi.fn();
    const plugin = createMsalCachePlugin('enc-2');
    await plugin.beforeCacheAccess({
      tokenCache: { serialize: vi.fn(), deserialize },
    } as never);

    expect(deserialize).toHaveBeenCalledWith(MSAL_BLOB);
  });

  it('still reads legacy plaintext cache entries (migration)', async () => {
    redisMock.get.mockResolvedValueOnce(MSAL_BLOB);
    const deserialize = vi.fn();
    const plugin = createMsalCachePlugin('enc-legacy');
    await plugin.beforeCacheAccess({
      tokenCache: { serialize: vi.fn(), deserialize },
    } as never);

    expect(deserialize).toHaveBeenCalledWith(MSAL_BLOB);
  });
});
