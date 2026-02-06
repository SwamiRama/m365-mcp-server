/**
 * Shared Redis client singleton.
 * All stores (session, client, code, token) share a single connection.
 */

import { Redis } from 'ioredis';
import { config } from './config.js';
import { logger } from './logger.js';

let sharedClient: Redis | null = null;

/**
 * Get the shared Redis client instance.
 * Returns null if REDIS_URL is not configured.
 * Creates the client lazily on first call.
 */
export function getRedisClient(): Redis | null {
  if (!config.redisUrl) return null;

  if (!sharedClient) {
    sharedClient = new Redis(config.redisUrl, {
      maxRetriesPerRequest: 3,
      retryStrategy(times: number) {
        const delay = Math.min(times * 100, 3000);
        logger.warn({ attempt: times, delayMs: delay }, 'Redis reconnecting...');
        return delay;
      },
    });

    sharedClient.on('error', (err) => {
      logger.error({ err }, 'Redis connection error');
    });

    sharedClient.on('connect', () => {
      logger.info('Redis connected');
    });

    sharedClient.on('close', () => {
      logger.debug('Redis connection closed');
    });
  }

  return sharedClient;
}

/**
 * Ping Redis to check connectivity and measure latency.
 * Returns { ok: false, latencyMs: 0 } if Redis is not configured.
 */
export async function pingRedis(): Promise<{ ok: boolean; latencyMs: number }> {
  const client = getRedisClient();
  if (!client) {
    return { ok: false, latencyMs: 0 };
  }

  const start = Date.now();
  try {
    await client.ping();
    return { ok: true, latencyMs: Date.now() - start };
  } catch {
    return { ok: false, latencyMs: Date.now() - start };
  }
}

/**
 * Close the shared Redis connection gracefully.
 * Should be called during server shutdown.
 */
export async function closeRedis(): Promise<void> {
  if (sharedClient) {
    try {
      await sharedClient.quit();
      logger.info('Redis connection closed gracefully');
    } catch (err) {
      logger.error({ err }, 'Error closing Redis connection');
      // Force disconnect if quit fails
      sharedClient.disconnect();
    }
    sharedClient = null;
  }
}
