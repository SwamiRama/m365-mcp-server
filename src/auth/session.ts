import crypto from 'crypto';
import type { Redis } from 'ioredis';
import { config } from '../utils/config.js';
import { logger } from '../utils/logger.js';
import { getRedisClient } from '../utils/redis.js';

export interface TokenSet {
  accessToken: string;
  refreshToken?: string;
  expiresAt: number; // Unix timestamp in ms
  scope: string;
}

export interface UserSession {
  id: string;
  userId?: string;
  userEmail?: string;
  userDisplayName?: string;
  tokens?: TokenSet;
  pkceVerifier?: string;
  state?: string;
  nonce?: string;
  createdAt: number;
  lastAccessedAt: number;
}

// Session storage interface
interface SessionStore {
  get(sessionId: string): Promise<UserSession | null>;
  set(sessionId: string, session: UserSession, ttlSeconds?: number): Promise<void>;
  delete(sessionId: string): Promise<void>;
}

// In-memory session store (development/single instance)
class MemorySessionStore implements SessionStore {
  private sessions = new Map<string, { session: UserSession; expiresAt: number }>();

  async get(sessionId: string): Promise<UserSession | null> {
    const entry = this.sessions.get(sessionId);
    if (!entry) return null;

    if (Date.now() > entry.expiresAt) {
      this.sessions.delete(sessionId);
      return null;
    }

    return entry.session;
  }

  async set(sessionId: string, session: UserSession, ttlSeconds = 86400): Promise<void> {
    this.sessions.set(sessionId, {
      session,
      expiresAt: Date.now() + ttlSeconds * 1000,
    });
  }

  async delete(sessionId: string): Promise<void> {
    this.sessions.delete(sessionId);
  }
}

// Redis session store (production/distributed)
class RedisSessionStore implements SessionStore {
  private client: Redis;
  private prefix = 'm365-mcp:session:';

  constructor(client: Redis) {
    this.client = client;
  }

  async get(sessionId: string): Promise<UserSession | null> {
    const data = await this.client.get(this.prefix + sessionId);
    if (!data) return null;

    try {
      return JSON.parse(data) as UserSession;
    } catch {
      return null;
    }
  }

  async set(sessionId: string, session: UserSession, ttlSeconds = 86400): Promise<void> {
    await this.client.setex(
      this.prefix + sessionId,
      ttlSeconds,
      JSON.stringify(session)
    );
  }

  async delete(sessionId: string): Promise<void> {
    await this.client.del(this.prefix + sessionId);
  }
}

// Create store based on config
const redisClient = getRedisClient();
const store: SessionStore = redisClient
  ? new RedisSessionStore(redisClient)
  : new MemorySessionStore();

// Encryption for sensitive session data
const ALGORITHM = 'aes-256-gcm';
const IV_LENGTH = 12;

function deriveKey(secret: string): Buffer {
  return crypto.scryptSync(secret, 'm365-mcp-salt', 32);
}

function encrypt(data: string): string {
  const key = deriveKey(config.sessionSecret);
  const iv = crypto.randomBytes(IV_LENGTH);
  const cipher = crypto.createCipheriv(ALGORITHM, key, iv);

  let encrypted = cipher.update(data, 'utf8', 'base64');
  encrypted += cipher.final('base64');

  const tag = cipher.getAuthTag();

  // Format: iv:tag:encrypted
  return `${iv.toString('base64')}:${tag.toString('base64')}:${encrypted}`;
}

function decrypt(encryptedData: string): string {
  const parts = encryptedData.split(':');
  if (parts.length !== 3) {
    throw new Error('Invalid encrypted data format');
  }

  const [ivStr, tagStr, encrypted] = parts;
  if (!ivStr || !tagStr || !encrypted) {
    throw new Error('Invalid encrypted data format');
  }

  const key = deriveKey(config.sessionSecret);
  const iv = Buffer.from(ivStr, 'base64');
  const tag = Buffer.from(tagStr, 'base64');

  const decipher = crypto.createDecipheriv(ALGORITHM, key, iv);
  decipher.setAuthTag(tag);

  let decrypted = decipher.update(encrypted, 'base64', 'utf8');
  decrypted += decipher.final('utf8');

  return decrypted;
}

export class SessionManager {
  /**
   * Generate a new session ID
   */
  generateSessionId(): string {
    return crypto.randomBytes(32).toString('hex');
  }

  /**
   * Create a new session
   */
  async createSession(initialData?: Partial<UserSession>): Promise<UserSession> {
    const now = Date.now();
    const session: UserSession = {
      id: this.generateSessionId(),
      createdAt: now,
      lastAccessedAt: now,
      ...initialData,
    };

    await this.saveSession(session);
    return session;
  }

  /**
   * Get a session by ID
   */
  async getSession(sessionId: string): Promise<UserSession | null> {
    const session = await store.get(sessionId);

    if (session) {
      // Update last accessed time
      session.lastAccessedAt = Date.now();
      await store.set(sessionId, session);
    }

    return session;
  }

  /**
   * Save/update a session
   */
  async saveSession(session: UserSession): Promise<void> {
    // Encrypt tokens before storage
    const sessionToStore = { ...session };

    if (sessionToStore.tokens) {
      sessionToStore.tokens = {
        ...sessionToStore.tokens,
        accessToken: encrypt(sessionToStore.tokens.accessToken),
        refreshToken: sessionToStore.tokens.refreshToken
          ? encrypt(sessionToStore.tokens.refreshToken)
          : undefined,
      };
    }

    if (sessionToStore.pkceVerifier) {
      sessionToStore.pkceVerifier = encrypt(sessionToStore.pkceVerifier);
    }

    await store.set(session.id, sessionToStore);
  }

  /**
   * Get decrypted tokens from session
   */
  getDecryptedTokens(session: UserSession): TokenSet | null {
    if (!session.tokens) return null;

    try {
      return {
        ...session.tokens,
        accessToken: decrypt(session.tokens.accessToken),
        refreshToken: session.tokens.refreshToken
          ? decrypt(session.tokens.refreshToken)
          : undefined,
      };
    } catch (err) {
      logger.error({ err }, 'Failed to decrypt tokens');
      return null;
    }
  }

  /**
   * Get decrypted PKCE verifier
   */
  getDecryptedPkceVerifier(session: UserSession): string | null {
    if (!session.pkceVerifier) return null;

    try {
      return decrypt(session.pkceVerifier);
    } catch (err) {
      logger.error({ err }, 'Failed to decrypt PKCE verifier');
      return null;
    }
  }

  /**
   * Update session tokens
   */
  async updateTokens(sessionId: string, tokens: TokenSet): Promise<void> {
    const session = await this.getSession(sessionId);
    if (!session) {
      throw new Error('Session not found');
    }

    session.tokens = tokens;
    await this.saveSession(session);
  }

  /**
   * Delete a session
   */
  async deleteSession(sessionId: string): Promise<void> {
    await store.delete(sessionId);
  }

  /**
   * Check if tokens need refresh (5 minutes before expiry)
   */
  tokensNeedRefresh(tokens: TokenSet): boolean {
    const bufferMs = 5 * 60 * 1000; // 5 minutes
    return Date.now() >= tokens.expiresAt - bufferMs;
  }
}

export const sessionManager = new SessionManager();
