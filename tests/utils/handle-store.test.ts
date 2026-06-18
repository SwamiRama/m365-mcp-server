import { describe, it, expect, vi, beforeEach, afterEach } from 'vitest';
import { config } from '../../src/utils/config.js';

// Re-import the module fresh so we exercise the Memory store (tests run without REDIS_URL).
import { MemoryHandleStore, RedisHandleStore } from '../../src/utils/handle-store.js';

describe('HandleStore', () => {
  describe('MemoryHandleStore', () => {
    let store: MemoryHandleStore;
    beforeEach(() => {
      store = new MemoryHandleStore();
    });

    it('mints a kind-prefixed handle and resolves it back to the payload', async () => {
      const handle = await store.mint('user-1', 'msg', { realId: 'AAMkReal==', mailbox: undefined });
      expect(handle).toMatch(/^m_[0-9a-f]{12}$/);
      const payload = await store.resolve('user-1', handle);
      expect(payload).toEqual({ realId: 'AAMkReal==', mailbox: undefined });
    });

    it('uses the a_ prefix for attachment handles', async () => {
      const handle = await store.mint('user-1', 'att', { realId: 'AttReal==' });
      expect(handle).toMatch(/^a_[0-9a-f]{12}$/);
    });

    it('returns null for an unknown handle', async () => {
      expect(await store.resolve('user-1', 'm_deadbeef0000')).toBeNull();
    });

    it('isolates handles per user (user B cannot resolve user A handle)', async () => {
      const handle = await store.mint('user-A', 'msg', { realId: 'X==' });
      expect(await store.resolve('user-B', handle)).toBeNull();
      expect(await store.resolve('user-A', handle)).toEqual({ realId: 'X==' });
    });

    it('expires a handle after handleTtlSecs', async () => {
      vi.useFakeTimers();
      vi.setSystemTime(0);
      const handle = await store.mint('user-1', 'msg', { realId: 'X==' });
      vi.setSystemTime(config.handleTtlSecs * 1000 + 1);
      expect(await store.resolve('user-1', handle)).toBeNull();
      vi.useRealTimers();
    });
  });

  describe('RedisHandleStore', () => {
    it('mint writes with setex under the namespaced key and resolve reads JSON', async () => {
      let stored: { key: string; ttl: number; val: string } | null = null;
      const fakeRedis = {
        setex: vi.fn(async (key: string, ttl: number, val: string) => {
          stored = { key, ttl, val };
          return 'OK';
        }),
        get: vi.fn(async (key: string) => (stored && stored.key === key ? stored.val : null)),
      } as unknown as import('ioredis').Redis;

      const store = new RedisHandleStore(fakeRedis);
      const handle = await store.mint('user-1', 'msg', { realId: 'AAMkReal==' });

      expect(fakeRedis.setex).toHaveBeenCalledTimes(1);
      expect(stored!.key).toBe(`m365-mcp:handle:user-1:${handle}`);
      expect(stored!.ttl).toBe(config.handleTtlSecs);

      const payload = await store.resolve('user-1', handle);
      expect(payload).toEqual({ realId: 'AAMkReal==' });
    });
  });
});
