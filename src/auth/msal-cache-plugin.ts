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

const BASE_REDIS_KEY = 'm365-mcp:msal-cache';
const TTL_SECONDS = 86400; // 24 hours

// In-memory fallback for development (map of keys to data)
const memoryCache = new Map<string, string>();

async function readCache(partitionKey: string): Promise<string | null> {
  const redis = getRedisClient();
  const key = `${BASE_REDIS_KEY}:${partitionKey}`;
  if (redis) {
    const data = await redis.get(key);
    if (data) {
      // Refresh TTL on read to prevent cache expiry while session is still active
      redis.expire(key, TTL_SECONDS).catch(() => {});
      return data;
    }
    // Migration fallback: try the old global key for sessions created before per-session cache.
    // Safe because acquireTokenSilent filters by localAccountId per session.
    const legacyData = await redis.get(BASE_REDIS_KEY);
    if (legacyData) {
      logger.info({ partitionKey }, 'Migrating MSAL cache from global to per-session key');
      await redis.setex(key, TTL_SECONDS, legacyData);
      return legacyData;
    }
    return null;
  }
  return memoryCache.get(key) ?? memoryCache.get(BASE_REDIS_KEY) ?? null;
}

async function writeCache(partitionKey: string, data: string): Promise<void> {
  const redis = getRedisClient();
  const key = `${BASE_REDIS_KEY}:${partitionKey}`;
  if (redis) {
    await redis.setex(key, TTL_SECONDS, data);
  } else {
    memoryCache.set(key, data);
  }
}

/**
 * Touch the MSAL cache TTL without reading/deserializing the data.
 * Called on every tool invocation to keep the cache alive even when
 * Azure AD tokens are still valid and no MSAL client is instantiated.
 */
export async function touchMsalCache(partitionKey: string): Promise<void> {
  const redis = getRedisClient();
  if (!redis) return;

  const key = `${BASE_REDIS_KEY}:${partitionKey}`;
  await redis.expire(key, TTL_SECONDS);
}

export function createMsalCachePlugin(partitionKey: string): ICachePlugin {
  return {
    async beforeCacheAccess(ctx: TokenCacheContext): Promise<void> {
      try {
        const cached = await readCache(partitionKey);
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
        await writeCache(partitionKey, data);
      } catch (err) {
        logger.warn(
          { err: err instanceof Error ? { message: err.message } : { message: String(err) } },
          'Failed to write MSAL cache'
        );
      }
    },
  };
}
