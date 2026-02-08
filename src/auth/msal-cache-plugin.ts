/**
 * MSAL Token Cache Plugin for Redis/Memory persistence.
 *
 * Without this plugin, MSAL's token cache is purely in-memory and lost on
 * container restart, causing acquireTokenSilent() to fail because the
 * account and refresh token are gone.
 */

import type { ICachePlugin, TokenCacheContext } from '@azure/msal-common/node';
import { getRedisClient } from '../utils/redis.js';
import { logger } from '../utils/logger.js';

const REDIS_KEY = 'm365-mcp:msal-cache';
const TTL_SECONDS = 86400; // 24 hours

// In-memory fallback for development
let memoryCache: string | null = null;

async function readCache(): Promise<string | null> {
  const redis = getRedisClient();
  if (redis) {
    return redis.get(REDIS_KEY);
  }
  return memoryCache;
}

async function writeCache(data: string): Promise<void> {
  const redis = getRedisClient();
  if (redis) {
    await redis.setex(REDIS_KEY, TTL_SECONDS, data);
  } else {
    memoryCache = data;
  }
}

export function createMsalCachePlugin(): ICachePlugin {
  return {
    async beforeCacheAccess(ctx: TokenCacheContext): Promise<void> {
      try {
        const cached = await readCache();
        if (cached) {
          ctx.tokenCache.deserialize(cached);
        }
      } catch (err) {
        logger.warn(
          { err: err instanceof Error ? { message: err.message } : { message: String(err) } },
          'Failed to read MSAL cache - continuing with empty cache'
        );
      }
    },

    async afterCacheAccess(ctx: TokenCacheContext): Promise<void> {
      if (!ctx.cacheHasChanged) return;

      try {
        const data = ctx.tokenCache.serialize();
        await writeCache(data);
      } catch (err) {
        logger.warn(
          { err: err instanceof Error ? { message: err.message } : { message: String(err) } },
          'Failed to write MSAL cache'
        );
      }
    },
  };
}
