/**
 * OAuth Client Storage (RFC 7591 - Dynamic Client Registration)
 */

import crypto from 'crypto';
import type { Redis } from 'ioredis';
import { config } from '../utils/config.js';
import { logger } from '../utils/logger.js';
import { getRedisClient } from '../utils/redis.js';
import type { OAuthClient, ClientRegistrationRequest, ClientRegistrationResponse, TokenEndpointAuthMethod } from './types.js';

/**
 * Hash sorted redirect URIs to create a stable lookup key.
 * Sorting ensures URI order doesn't matter.
 */
export function hashRedirectUris(redirectUris: string[]): string {
  const sorted = [...redirectUris].sort();
  return crypto.createHash('sha256').update(sorted.join('\n')).digest('hex');
}

// Storage interface
interface ClientStore {
  get(clientId: string): Promise<OAuthClient | null>;
  getByRedirectUris(redirectUris: string[]): Promise<OAuthClient | null>;
  set(clientId: string, client: OAuthClient): Promise<void>;
  delete(clientId: string): Promise<void>;
}

// In-memory store (development/single instance)
class MemoryClientStore implements ClientStore {
  private clients = new Map<string, OAuthClient>();

  async get(clientId: string): Promise<OAuthClient | null> {
    return this.clients.get(clientId) ?? null;
  }

  async getByRedirectUris(redirectUris: string[]): Promise<OAuthClient | null> {
    const hash = hashRedirectUris(redirectUris);
    for (const client of this.clients.values()) {
      if (hashRedirectUris(client.redirectUris) === hash) {
        return client;
      }
    }
    return null;
  }

  async set(clientId: string, client: OAuthClient): Promise<void> {
    this.clients.set(clientId, client);
  }

  async delete(clientId: string): Promise<void> {
    this.clients.delete(clientId);
  }
}

// Redis store (production/distributed)
class RedisClientStore implements ClientStore {
  private client: Redis;
  private prefix = 'm365-mcp:oauth-client:';
  private uriIndexPrefix = 'm365-mcp:oauth-uris-index:';
  private ttlSeconds = 365 * 24 * 60 * 60; // 365 days

  constructor(client: Redis) {
    this.client = client;
  }

  async get(clientId: string): Promise<OAuthClient | null> {
    const data = await this.client.get(this.prefix + clientId);
    if (!data) return null;

    try {
      return JSON.parse(data) as OAuthClient;
    } catch {
      return null;
    }
  }

  async getByRedirectUris(redirectUris: string[]): Promise<OAuthClient | null> {
    const hash = hashRedirectUris(redirectUris);
    const clientId = await this.client.get(this.uriIndexPrefix + hash);
    if (!clientId) return null;
    return this.get(clientId);
  }

  async set(clientId: string, client: OAuthClient): Promise<void> {
    const hash = hashRedirectUris(client.redirectUris);
    const pipeline = this.client.pipeline();
    pipeline.setex(this.prefix + clientId, this.ttlSeconds, JSON.stringify(client));
    pipeline.setex(this.uriIndexPrefix + hash, this.ttlSeconds, clientId);
    await pipeline.exec();
  }

  async delete(clientId: string): Promise<void> {
    // Look up the client first to clean up the URI index
    const client = await this.get(clientId);
    if (client) {
      const hash = hashRedirectUris(client.redirectUris);
      const pipeline = this.client.pipeline();
      pipeline.del(this.prefix + clientId);
      pipeline.del(this.uriIndexPrefix + hash);
      await pipeline.exec();
    } else {
      await this.client.del(this.prefix + clientId);
    }
  }
}

// Create store based on config
const redisClient = getRedisClient();
const store: ClientStore = redisClient
  ? new RedisClientStore(redisClient)
  : new MemoryClientStore();

// Export for testing
export { store as _store };

/**
 * Generate a secure client ID
 */
function generateClientId(): string {
  return crypto.randomBytes(16).toString('hex');
}

/**
 * Generate a secure client secret
 */
function generateClientSecret(): string {
  return crypto.randomBytes(32).toString('base64url');
}

/**
 * Hash a client secret for storage
 */
function hashClientSecret(secret: string): string {
  return crypto.createHash('sha256').update(secret).digest('hex');
}

/**
 * Verify a client secret against its hash
 */
export function verifyClientSecret(secret: string, hash: string): boolean {
  const inputHash = hashClientSecret(secret);
  return crypto.timingSafeEqual(Buffer.from(inputHash), Buffer.from(hash));
}

/**
 * Validate redirect URI
 */
function validateRedirectUri(uri: string): boolean {
  try {
    const parsed = new URL(uri);

    // Must be HTTPS in production (allow localhost HTTP for development)
    if (config.nodeEnv === 'production') {
      if (parsed.protocol !== 'https:' && parsed.hostname !== 'localhost' && parsed.hostname !== '127.0.0.1') {
        return false;
      }
    }

    // No fragments allowed
    if (parsed.hash) {
      return false;
    }

    return true;
  } catch {
    return false;
  }
}

/**
 * Register a new OAuth client (RFC 7591) - idempotent by redirect_uris.
 * If a client with the same redirect_uris already exists, return it.
 */
export async function registerClient(request: ClientRegistrationRequest): Promise<ClientRegistrationResponse> {
  // Validate redirect URIs
  if (!request.redirect_uris || request.redirect_uris.length === 0) {
    throw new Error('redirect_uris is required');
  }

  for (const uri of request.redirect_uris) {
    if (!validateRedirectUri(uri)) {
      throw new Error(`Invalid redirect URI: ${uri}`);
    }
  }

  // Idempotent: check if a client with these redirect_uris already exists
  const existing = await store.getByRedirectUris(request.redirect_uris);
  if (existing) {
    logger.info(
      { clientId: existing.clientId, clientName: existing.clientName },
      'Idempotent DCR: returning existing client for redirect_uris'
    );

    const response: ClientRegistrationResponse = {
      client_id: existing.clientId,
      client_name: existing.clientName,
      redirect_uris: existing.redirectUris,
      grant_types: existing.grantTypes,
      response_types: existing.responseTypes,
      token_endpoint_auth_method: existing.tokenEndpointAuthMethod,
      client_id_issued_at: Math.floor(existing.createdAt / 1000),
      client_secret_expires_at: 0,
    };

    return response;
  }

  // Generate credentials
  const clientId = generateClientId();

  // Defaults per OAuth 2.1
  const grantTypes = request.grant_types ?? ['authorization_code', 'refresh_token'];
  const responseTypes = request.response_types ?? ['code'];

  // Always enforce public client - PKCE provides sufficient security
  const tokenEndpointAuthMethod: TokenEndpointAuthMethod = 'none';

  const client: OAuthClient = {
    clientId,
    clientSecret: '',
    clientName: request.client_name,
    redirectUris: request.redirect_uris,
    grantTypes,
    responseTypes,
    tokenEndpointAuthMethod,
    scope: request.scope,
    createdAt: Date.now(),
  };

  await store.set(clientId, client);

  logger.info({ clientId, clientName: client.clientName }, 'Registered new public OAuth client');

  const response: ClientRegistrationResponse = {
    client_id: clientId,
    client_name: client.clientName,
    redirect_uris: client.redirectUris,
    grant_types: grantTypes,
    response_types: responseTypes,
    token_endpoint_auth_method: tokenEndpointAuthMethod,
    client_id_issued_at: Math.floor(client.createdAt / 1000),
    client_secret_expires_at: 0,
  };

  return response;
}

/**
 * Get a client by ID
 */
export async function getClient(clientId: string): Promise<OAuthClient | null> {
  return store.get(clientId);
}

/**
 * Authenticate a client (for token endpoint)
 */
export async function authenticateClient(
  clientId: string,
  clientSecret: string
): Promise<OAuthClient | null> {
  const client = await store.get(clientId);

  if (!client) {
    return null;
  }

  if (!verifyClientSecret(clientSecret, client.clientSecret)) {
    return null;
  }

  return client;
}

/**
 * Validate redirect URI against registered client
 */
export function validateClientRedirectUri(client: OAuthClient, redirectUri: string): boolean {
  // Exact match required (no wildcards)
  return client.redirectUris.includes(redirectUri);
}

/**
 * Delete a client (for management)
 */
export async function deleteClient(clientId: string): Promise<void> {
  await store.delete(clientId);
  logger.info({ clientId }, 'Deleted OAuth client');
}
