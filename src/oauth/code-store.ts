/**
 * Authorization Code Storage with PKCE support
 */

import crypto from 'crypto';
import { Redis } from 'ioredis';
import { config } from '../utils/config.js';
import { logger } from '../utils/logger.js';
import type { AuthorizationCode, PendingAuthorization } from './types.js';

// Storage interface
interface CodeStore {
  get(code: string): Promise<AuthorizationCode | null>;
  set(code: string, authCode: AuthorizationCode, ttlSeconds: number): Promise<void>;
  delete(code: string): Promise<void>;
  // Pending authorizations (during Azure AD flow)
  getPending(sessionId: string): Promise<PendingAuthorization | null>;
  setPending(sessionId: string, pending: PendingAuthorization, ttlSeconds: number): Promise<void>;
  deletePending(sessionId: string): Promise<void>;
}

// In-memory store
class MemoryCodeStore implements CodeStore {
  private codes = new Map<string, { authCode: AuthorizationCode; expiresAt: number }>();
  private pending = new Map<string, { data: PendingAuthorization; expiresAt: number }>();

  async get(code: string): Promise<AuthorizationCode | null> {
    const entry = this.codes.get(code);
    if (!entry) return null;

    if (Date.now() > entry.expiresAt) {
      this.codes.delete(code);
      return null;
    }

    return entry.authCode;
  }

  async set(code: string, authCode: AuthorizationCode, ttlSeconds: number): Promise<void> {
    this.codes.set(code, {
      authCode,
      expiresAt: Date.now() + ttlSeconds * 1000,
    });
  }

  async delete(code: string): Promise<void> {
    this.codes.delete(code);
  }

  async getPending(sessionId: string): Promise<PendingAuthorization | null> {
    const entry = this.pending.get(sessionId);
    if (!entry) return null;

    if (Date.now() > entry.expiresAt) {
      this.pending.delete(sessionId);
      return null;
    }

    return entry.data;
  }

  async setPending(sessionId: string, data: PendingAuthorization, ttlSeconds: number): Promise<void> {
    this.pending.set(sessionId, {
      data,
      expiresAt: Date.now() + ttlSeconds * 1000,
    });
  }

  async deletePending(sessionId: string): Promise<void> {
    this.pending.delete(sessionId);
  }
}

// Redis store
class RedisCodeStore implements CodeStore {
  private client: Redis;
  private codePrefix = 'm365-mcp:oauth-code:';
  private pendingPrefix = 'm365-mcp:oauth-pending:';

  constructor(redisUrl: string) {
    this.client = new Redis(redisUrl);
    this.client.on('error', (err) => {
      logger.error({ err }, 'Redis connection error (code store)');
    });
  }

  async get(code: string): Promise<AuthorizationCode | null> {
    const data = await this.client.get(this.codePrefix + code);
    if (!data) return null;

    try {
      return JSON.parse(data) as AuthorizationCode;
    } catch {
      return null;
    }
  }

  async set(code: string, authCode: AuthorizationCode, ttlSeconds: number): Promise<void> {
    await this.client.setex(this.codePrefix + code, ttlSeconds, JSON.stringify(authCode));
  }

  async delete(code: string): Promise<void> {
    await this.client.del(this.codePrefix + code);
  }

  async getPending(sessionId: string): Promise<PendingAuthorization | null> {
    const data = await this.client.get(this.pendingPrefix + sessionId);
    if (!data) return null;

    try {
      return JSON.parse(data) as PendingAuthorization;
    } catch {
      return null;
    }
  }

  async setPending(sessionId: string, data: PendingAuthorization, ttlSeconds: number): Promise<void> {
    await this.client.setex(this.pendingPrefix + sessionId, ttlSeconds, JSON.stringify(data));
  }

  async deletePending(sessionId: string): Promise<void> {
    await this.client.del(this.pendingPrefix + sessionId);
  }
}

// Create store based on config
const store: CodeStore = config.redisUrl
  ? new RedisCodeStore(config.redisUrl)
  : new MemoryCodeStore();

/**
 * Generate a cryptographically secure authorization code
 */
function generateCode(): string {
  return crypto.randomBytes(32).toString('base64url');
}

/**
 * Verify PKCE code challenge
 */
export function verifyCodeChallenge(codeVerifier: string, codeChallenge: string, method: 'S256'): boolean {
  if (method !== 'S256') {
    return false; // Only S256 supported (OAuth 2.1)
  }

  const computedChallenge = crypto
    .createHash('sha256')
    .update(codeVerifier, 'ascii')
    .digest('base64url');

  // Constant-time comparison
  try {
    return crypto.timingSafeEqual(
      Buffer.from(computedChallenge),
      Buffer.from(codeChallenge)
    );
  } catch {
    return false;
  }
}

/**
 * Store pending authorization (before Azure AD redirect)
 */
export async function storePendingAuthorization(
  sessionId: string,
  params: Omit<PendingAuthorization, 'createdAt' | 'expiresAt'>
): Promise<void> {
  const ttl = config.oauthAuthCodeLifetimeSecs;
  const now = Date.now();

  const pending: PendingAuthorization = {
    ...params,
    createdAt: now,
    expiresAt: now + ttl * 1000,
  };

  await store.setPending(sessionId, pending, ttl);
  logger.debug({ sessionId, clientId: params.clientId }, 'Stored pending authorization');
}

/**
 * Get pending authorization
 */
export async function getPendingAuthorization(sessionId: string): Promise<PendingAuthorization | null> {
  return store.getPending(sessionId);
}

/**
 * Delete pending authorization
 */
export async function deletePendingAuthorization(sessionId: string): Promise<void> {
  await store.deletePending(sessionId);
}

/**
 * Create and store an authorization code
 */
export async function createAuthorizationCode(params: {
  clientId: string;
  redirectUri: string;
  scope: string;
  codeChallenge: string;
  codeChallengeMethod: 'S256';
  sessionId: string;
  userId: string;
  userEmail?: string;
}): Promise<string> {
  const code = generateCode();
  const ttl = config.oauthAuthCodeLifetimeSecs;

  const authCode: AuthorizationCode = {
    code,
    clientId: params.clientId,
    redirectUri: params.redirectUri,
    scope: params.scope,
    codeChallenge: params.codeChallenge,
    codeChallengeMethod: params.codeChallengeMethod,
    sessionId: params.sessionId,
    userId: params.userId,
    userEmail: params.userEmail,
    expiresAt: Date.now() + ttl * 1000,
    used: false,
  };

  await store.set(code, authCode, ttl);

  logger.debug(
    { clientId: params.clientId, sessionId: params.sessionId },
    'Created authorization code'
  );

  return code;
}

/**
 * Consume an authorization code (single-use)
 */
export async function consumeAuthorizationCode(code: string): Promise<AuthorizationCode | null> {
  const authCode = await store.get(code);

  if (!authCode) {
    logger.debug('Authorization code not found');
    return null;
  }

  // Check if already used
  if (authCode.used) {
    logger.warn({ code: code.slice(0, 8) + '...' }, 'Authorization code already used');
    // Delete the code to prevent further attempts
    await store.delete(code);
    return null;
  }

  // Check expiration
  if (Date.now() > authCode.expiresAt) {
    logger.debug('Authorization code expired');
    await store.delete(code);
    return null;
  }

  // Mark as used and delete
  authCode.used = true;
  await store.delete(code);

  logger.debug({ clientId: authCode.clientId }, 'Authorization code consumed');

  return authCode;
}
