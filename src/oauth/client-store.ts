/**
 * OAuth Client Storage (RFC 7591 - Dynamic Client Registration)
 */

import crypto from 'crypto';
import type { Redis } from 'ioredis';
import { config } from '../utils/config.js';
import { logger } from '../utils/logger.js';
import { getRedisClient } from '../utils/redis.js';
import type { OAuthClient, ClientRegistrationRequest, ClientRegistrationResponse } from './types.js';

// Storage interface
interface ClientStore {
  get(clientId: string): Promise<OAuthClient | null>;
  set(clientId: string, client: OAuthClient): Promise<void>;
  delete(clientId: string): Promise<void>;
}

// In-memory store (development/single instance)
class MemoryClientStore implements ClientStore {
  private clients = new Map<string, OAuthClient>();

  async get(clientId: string): Promise<OAuthClient | null> {
    return this.clients.get(clientId) ?? null;
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

  async set(clientId: string, client: OAuthClient): Promise<void> {
    // Clients don't expire by default
    await this.client.set(this.prefix + clientId, JSON.stringify(client));
  }

  async delete(clientId: string): Promise<void> {
    await this.client.del(this.prefix + clientId);
  }
}

// Create store based on config
const redisClient = getRedisClient();
const store: ClientStore = redisClient
  ? new RedisClientStore(redisClient)
  : new MemoryClientStore();

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
 * Register a new OAuth client (RFC 7591)
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

  // Generate credentials
  const clientId = generateClientId();

  // Defaults per OAuth 2.1
  const grantTypes = request.grant_types ?? ['authorization_code', 'refresh_token'];
  const responseTypes = request.response_types ?? ['code'];
  const tokenEndpointAuthMethod = request.token_endpoint_auth_method ?? 'none'; // Default to public client

  // Only generate secret for confidential clients
  const isPublicClient = tokenEndpointAuthMethod === 'none';
  const clientSecretPlain = isPublicClient ? '' : generateClientSecret();
  const clientSecretHash = isPublicClient ? '' : hashClientSecret(clientSecretPlain);

  const client: OAuthClient = {
    clientId,
    clientSecret: clientSecretHash,
    clientName: request.client_name,
    redirectUris: request.redirect_uris,
    grantTypes,
    responseTypes,
    tokenEndpointAuthMethod,
    scope: request.scope,
    createdAt: Date.now(),
  };

  await store.set(clientId, client);

  logger.info({ clientId, clientName: client.clientName, isPublicClient }, 'Registered new OAuth client');

  // Build response - only include secret for confidential clients
  const response: ClientRegistrationResponse = {
    client_id: clientId,
    client_secret: clientSecretPlain,
    client_name: client.clientName,
    redirect_uris: client.redirectUris,
    grant_types: grantTypes,
    response_types: responseTypes,
    token_endpoint_auth_method: tokenEndpointAuthMethod,
    client_id_issued_at: Math.floor(client.createdAt / 1000),
    client_secret_expires_at: isPublicClient ? 0 : 0, // 0 = never expires
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
