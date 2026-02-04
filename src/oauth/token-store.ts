/**
 * Refresh Token Storage with rotation support
 */

import crypto from 'crypto';
import { Redis } from 'ioredis';
import { config } from '../utils/config.js';
import { logger } from '../utils/logger.js';
import type { RefreshTokenRecord } from './types.js';

// Storage interface
interface TokenStore {
  get(tokenHash: string): Promise<RefreshTokenRecord | null>;
  set(tokenHash: string, record: RefreshTokenRecord, ttlSeconds: number): Promise<void>;
  delete(tokenHash: string): Promise<void>;
  // Index by session for cleanup
  getBySession(sessionId: string): Promise<RefreshTokenRecord[]>;
  deleteBySession(sessionId: string): Promise<void>;
}

// In-memory store
class MemoryTokenStore implements TokenStore {
  private tokens = new Map<string, { record: RefreshTokenRecord; expiresAt: number }>();
  private sessionIndex = new Map<string, Set<string>>(); // sessionId -> tokenHashes

  async get(tokenHash: string): Promise<RefreshTokenRecord | null> {
    const entry = this.tokens.get(tokenHash);
    if (!entry) return null;

    if (Date.now() > entry.expiresAt) {
      this.deleteSync(tokenHash);
      return null;
    }

    return entry.record;
  }

  async set(tokenHash: string, record: RefreshTokenRecord, ttlSeconds: number): Promise<void> {
    this.tokens.set(tokenHash, {
      record,
      expiresAt: Date.now() + ttlSeconds * 1000,
    });

    // Update session index
    let sessionTokens = this.sessionIndex.get(record.sessionId);
    if (!sessionTokens) {
      sessionTokens = new Set();
      this.sessionIndex.set(record.sessionId, sessionTokens);
    }
    sessionTokens.add(tokenHash);
  }

  async delete(tokenHash: string): Promise<void> {
    this.deleteSync(tokenHash);
  }

  private deleteSync(tokenHash: string): void {
    const entry = this.tokens.get(tokenHash);
    if (entry) {
      // Remove from session index
      const sessionTokens = this.sessionIndex.get(entry.record.sessionId);
      if (sessionTokens) {
        sessionTokens.delete(tokenHash);
        if (sessionTokens.size === 0) {
          this.sessionIndex.delete(entry.record.sessionId);
        }
      }
    }
    this.tokens.delete(tokenHash);
  }

  async getBySession(sessionId: string): Promise<RefreshTokenRecord[]> {
    const tokenHashes = this.sessionIndex.get(sessionId);
    if (!tokenHashes) return [];

    const records: RefreshTokenRecord[] = [];
    for (const hash of tokenHashes) {
      const record = await this.get(hash);
      if (record) {
        records.push(record);
      }
    }
    return records;
  }

  async deleteBySession(sessionId: string): Promise<void> {
    const tokenHashes = this.sessionIndex.get(sessionId);
    if (!tokenHashes) return;

    for (const hash of tokenHashes) {
      this.tokens.delete(hash);
    }
    this.sessionIndex.delete(sessionId);
  }
}

// Redis store
class RedisTokenStore implements TokenStore {
  private client: Redis;
  private tokenPrefix = 'm365-mcp:oauth-refresh:';
  private sessionPrefix = 'm365-mcp:oauth-refresh-session:';

  constructor(redisUrl: string) {
    this.client = new Redis(redisUrl);
    this.client.on('error', (err) => {
      logger.error({ err }, 'Redis connection error (token store)');
    });
  }

  async get(tokenHash: string): Promise<RefreshTokenRecord | null> {
    const data = await this.client.get(this.tokenPrefix + tokenHash);
    if (!data) return null;

    try {
      return JSON.parse(data) as RefreshTokenRecord;
    } catch {
      return null;
    }
  }

  async set(tokenHash: string, record: RefreshTokenRecord, ttlSeconds: number): Promise<void> {
    const pipeline = this.client.pipeline();

    // Store token record
    pipeline.setex(this.tokenPrefix + tokenHash, ttlSeconds, JSON.stringify(record));

    // Add to session index
    pipeline.sadd(this.sessionPrefix + record.sessionId, tokenHash);
    pipeline.expire(this.sessionPrefix + record.sessionId, ttlSeconds);

    await pipeline.exec();
  }

  async delete(tokenHash: string): Promise<void> {
    const record = await this.get(tokenHash);
    if (record) {
      await this.client.srem(this.sessionPrefix + record.sessionId, tokenHash);
    }
    await this.client.del(this.tokenPrefix + tokenHash);
  }

  async getBySession(sessionId: string): Promise<RefreshTokenRecord[]> {
    const hashes = await this.client.smembers(this.sessionPrefix + sessionId);
    const records: RefreshTokenRecord[] = [];

    for (const hash of hashes) {
      const record = await this.get(hash);
      if (record) {
        records.push(record);
      }
    }

    return records;
  }

  async deleteBySession(sessionId: string): Promise<void> {
    const hashes = await this.client.smembers(this.sessionPrefix + sessionId);

    if (hashes.length > 0) {
      const pipeline = this.client.pipeline();
      for (const hash of hashes) {
        pipeline.del(this.tokenPrefix + hash);
      }
      pipeline.del(this.sessionPrefix + sessionId);
      await pipeline.exec();
    }
  }
}

// Create store based on config
const store: TokenStore = config.redisUrl
  ? new RedisTokenStore(config.redisUrl)
  : new MemoryTokenStore();

/**
 * Generate a secure opaque refresh token
 */
function generateRefreshToken(): string {
  return crypto.randomBytes(48).toString('base64url');
}

/**
 * Hash a refresh token for storage
 */
function hashToken(token: string): string {
  return crypto.createHash('sha256').update(token).digest('hex');
}

/**
 * Create and store a new refresh token
 */
export async function createRefreshToken(params: {
  clientId: string;
  sessionId: string;
  userId: string;
  scope: string;
  rotatedFrom?: string;
}): Promise<string> {
  const token = generateRefreshToken();
  const tokenHash = hashToken(token);
  const ttl = config.oauthRefreshTokenLifetimeSecs;
  const now = Date.now();

  const record: RefreshTokenRecord = {
    token: tokenHash,
    clientId: params.clientId,
    sessionId: params.sessionId,
    userId: params.userId,
    scope: params.scope,
    expiresAt: now + ttl * 1000,
    rotatedFrom: params.rotatedFrom,
    createdAt: now,
  };

  await store.set(tokenHash, record, ttl);

  logger.debug(
    { clientId: params.clientId, sessionId: params.sessionId },
    'Created refresh token'
  );

  return token;
}

/**
 * Validate and get refresh token record
 */
export async function validateRefreshToken(token: string): Promise<RefreshTokenRecord | null> {
  const tokenHash = hashToken(token);
  const record = await store.get(tokenHash);

  if (!record) {
    logger.debug('Refresh token not found');
    return null;
  }

  // Check expiration
  if (Date.now() > record.expiresAt) {
    logger.debug('Refresh token expired');
    await store.delete(tokenHash);
    return null;
  }

  return record;
}

/**
 * Rotate a refresh token (issue new one, invalidate old)
 */
export async function rotateRefreshToken(
  oldToken: string,
  params: {
    clientId: string;
    sessionId: string;
    userId: string;
    scope: string;
  }
): Promise<string | null> {
  const oldTokenHash = hashToken(oldToken);
  const oldRecord = await store.get(oldTokenHash);

  if (!oldRecord) {
    return null;
  }

  // Create new token with reference to rotated token
  const newToken = await createRefreshToken({
    ...params,
    rotatedFrom: oldTokenHash,
  });

  // Invalidate old token
  await store.delete(oldTokenHash);

  logger.debug(
    { clientId: params.clientId, sessionId: params.sessionId },
    'Rotated refresh token'
  );

  return newToken;
}

/**
 * Revoke a refresh token
 */
export async function revokeRefreshToken(token: string): Promise<void> {
  const tokenHash = hashToken(token);
  await store.delete(tokenHash);
  logger.debug('Revoked refresh token');
}

/**
 * Revoke all refresh tokens for a session
 */
export async function revokeSessionTokens(sessionId: string): Promise<void> {
  await store.deleteBySession(sessionId);
  logger.debug({ sessionId }, 'Revoked all refresh tokens for session');
}

/**
 * Get all refresh tokens for a session (for management/audit)
 */
export async function getSessionTokens(sessionId: string): Promise<RefreshTokenRecord[]> {
  return store.getBySession(sessionId);
}
