/**
 * Short, per-user handles for Microsoft Graph item IDs.
 *
 * Graph IDs are long opaque base64url strings. An LLM cannot reliably relay one
 * across tool calls (it truncates, re-encodes, or substitutes a list ordinal),
 * which breaks the "list -> open -> read attachment" flow. So the mail tools hand
 * the model a short handle (e.g. "m_3f9a1c08bd12") and keep the real ID server-side.
 *
 * Security: a handle maps only to an ID *string*, namespaced per user; it grants no
 * access. Every Graph call still uses the requesting user's own token, so even a
 * (vanishingly unlikely) cross-user handle hit could not leak data. Mirrors the
 * token/code/session stores: interface + Memory + Redis, selected by getRedisClient().
 */

import crypto from 'crypto';
import type { Redis } from 'ioredis';
import { config } from './config.js';
import { getRedisClient } from './redis.js';

export type HandleKind = 'msg' | 'att';

export interface HandlePayload {
  realId: string;
  mailbox?: string;
}

export interface HandleStore {
  mint(userKey: string, kind: HandleKind, payload: HandlePayload): Promise<string>;
  resolve(userKey: string, handle: string): Promise<HandlePayload | null>;
}

const KIND_PREFIX: Record<HandleKind, string> = { msg: 'm', att: 'a' };

// 6 random bytes -> 12 hex chars: short enough for an LLM to copy verbatim, wide
// enough (2^48) that collisions within a user's active set are negligible.
function newHandle(kind: HandleKind): string {
  return `${KIND_PREFIX[kind]}_${crypto.randomBytes(6).toString('hex')}`;
}

export class MemoryHandleStore implements HandleStore {
  private map = new Map<string, { payload: HandlePayload; expiresAt: number }>();

  async mint(userKey: string, kind: HandleKind, payload: HandlePayload): Promise<string> {
    const handle = newHandle(kind);
    this.map.set(`${userKey}:${handle}`, {
      payload,
      expiresAt: Date.now() + config.handleTtlSecs * 1000,
    });
    return handle;
  }

  async resolve(userKey: string, handle: string): Promise<HandlePayload | null> {
    const entry = this.map.get(`${userKey}:${handle}`);
    if (!entry) return null;
    if (Date.now() > entry.expiresAt) {
      this.map.delete(`${userKey}:${handle}`);
      return null;
    }
    return { ...entry.payload };
  }

  clear(): void {
    this.map.clear();
  }
}

export class RedisHandleStore implements HandleStore {
  private prefix = 'm365-mcp:handle:';
  private client: Redis;

  constructor(client: Redis) {
    this.client = client;
  }

  private key(userKey: string, handle: string): string {
    return `${this.prefix}${userKey}:${handle}`;
  }

  async mint(userKey: string, kind: HandleKind, payload: HandlePayload): Promise<string> {
    const handle = newHandle(kind);
    await this.client.setex(this.key(userKey, handle), config.handleTtlSecs, JSON.stringify(payload));
    return handle;
  }

  async resolve(userKey: string, handle: string): Promise<HandlePayload | null> {
    const data = await this.client.get(this.key(userKey, handle));
    if (!data) return null;
    try {
      return JSON.parse(data) as HandlePayload;
    } catch {
      return null;
    }
  }
}

const redisClient = getRedisClient();
const memoryFallback = new MemoryHandleStore();
export const handleStore: HandleStore = redisClient
  ? new RedisHandleStore(redisClient)
  : memoryFallback;

/** Test-only: reset the in-memory store between tests (no-op against Redis). */
export function __clearHandleStore(): void {
  memoryFallback.clear();
}
