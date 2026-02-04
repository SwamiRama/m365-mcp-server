/**
 * JWT Signing and Verification with RS256
 */

import crypto from 'crypto';
import fs from 'fs';
import { config } from '../utils/config.js';
import { logger } from '../utils/logger.js';
import type { AccessTokenPayload, JWK, JWKS } from './types.js';

// Key pair storage
let privateKey: crypto.KeyObject | null = null;
let publicKey: crypto.KeyObject | null = null;
let keyId: string | null = null;

/**
 * Initialize or load RSA key pair for JWT signing
 */
export function initializeKeys(): void {
  const privateKeySource = config.oauthSigningKeyPrivate;
  const publicKeySource = config.oauthSigningKeyPublic;

  if (privateKeySource && publicKeySource) {
    // Load from config (PEM string or file path)
    privateKey = loadKeyFromSource(privateKeySource, 'private');
    publicKey = loadKeyFromSource(publicKeySource, 'public');
    logger.info('Loaded OAuth signing keys from configuration');
  } else {
    // Generate ephemeral key pair (development mode)
    const keyPair = crypto.generateKeyPairSync('rsa', {
      modulusLength: 2048,
      publicKeyEncoding: { type: 'spki', format: 'pem' },
      privateKeyEncoding: { type: 'pkcs8', format: 'pem' },
    });

    privateKey = crypto.createPrivateKey(keyPair.privateKey);
    publicKey = crypto.createPublicKey(keyPair.publicKey);

    logger.warn(
      'Generated ephemeral OAuth signing keys - tokens will be invalid after restart. ' +
        'Set OAUTH_SIGNING_KEY_PRIVATE and OAUTH_SIGNING_KEY_PUBLIC for production.'
    );
  }

  // Generate key ID from public key thumbprint
  keyId = generateKeyId(publicKey);
  logger.info({ keyId }, 'OAuth JWT key initialized');
}

/**
 * Load key from PEM string or file path
 */
function loadKeyFromSource(source: string, type: 'private' | 'public'): crypto.KeyObject {
  let pem: string;

  // Check if source is a file path
  if (source.startsWith('/') || source.startsWith('./') || source.includes('.pem')) {
    try {
      pem = fs.readFileSync(source, 'utf-8');
    } catch {
      throw new Error(`Failed to load ${type} key from file: ${source}`);
    }
  } else if (source.includes('-----BEGIN')) {
    // Direct PEM content
    pem = source;
  } else {
    // Base64-encoded PEM
    pem = Buffer.from(source, 'base64').toString('utf-8');
  }

  return type === 'private' ? crypto.createPrivateKey(pem) : crypto.createPublicKey(pem);
}

/**
 * Generate key ID from public key (SHA-256 thumbprint)
 */
function generateKeyId(key: crypto.KeyObject): string {
  const jwk = key.export({ format: 'jwk' });
  const thumbprintData = JSON.stringify({
    e: jwk.e,
    kty: jwk.kty,
    n: jwk.n,
  });

  return crypto.createHash('sha256').update(thumbprintData).digest('base64url').slice(0, 8);
}

/**
 * Ensure keys are initialized
 */
function ensureKeys(): { privateKey: crypto.KeyObject; publicKey: crypto.KeyObject; keyId: string } {
  if (!privateKey || !publicKey || !keyId) {
    initializeKeys();
  }

  if (!privateKey || !publicKey || !keyId) {
    throw new Error('OAuth signing keys not initialized');
  }

  return { privateKey, publicKey, keyId };
}

/**
 * Base64URL encode
 */
function base64url(data: Buffer | string): string {
  const buffer = typeof data === 'string' ? Buffer.from(data) : data;
  return buffer.toString('base64url');
}

/**
 * Sign a JWT with RS256
 */
export function signJwt(payload: AccessTokenPayload): string {
  const { privateKey: key, keyId: kid } = ensureKeys();

  const header = {
    alg: 'RS256',
    typ: 'JWT',
    kid,
  };

  const headerB64 = base64url(JSON.stringify(header));
  const payloadB64 = base64url(JSON.stringify(payload));
  const signingInput = `${headerB64}.${payloadB64}`;

  const signature = crypto.createSign('RSA-SHA256').update(signingInput).sign(key);

  return `${signingInput}.${base64url(signature)}`;
}

/**
 * Verify and decode a JWT
 */
export function verifyJwt(token: string): AccessTokenPayload | null {
  const { publicKey: key } = ensureKeys();

  const parts = token.split('.');
  if (parts.length !== 3) {
    return null;
  }

  const [headerB64, payloadB64, signatureB64] = parts;

  if (!headerB64 || !payloadB64 || !signatureB64) {
    return null;
  }

  try {
    // Verify signature
    const signingInput = `${headerB64}.${payloadB64}`;
    const signature = Buffer.from(signatureB64, 'base64url');

    const isValid = crypto.createVerify('RSA-SHA256').update(signingInput).verify(key, signature);

    if (!isValid) {
      return null;
    }

    // Decode and validate payload
    const payload = JSON.parse(Buffer.from(payloadB64, 'base64url').toString('utf-8')) as AccessTokenPayload;

    // Check expiration
    const now = Math.floor(Date.now() / 1000);
    if (payload.exp && payload.exp < now) {
      return null;
    }

    // Check not before (if present)
    if ('nbf' in payload && typeof payload.nbf === 'number' && payload.nbf > now) {
      return null;
    }

    return payload;
  } catch {
    return null;
  }
}

/**
 * Decode JWT without verification (for debugging/logging)
 */
export function decodeJwt(token: string): { header: unknown; payload: unknown } | null {
  const parts = token.split('.');
  if (parts.length !== 3) {
    return null;
  }

  const [headerB64, payloadB64] = parts;

  if (!headerB64 || !payloadB64) {
    return null;
  }

  try {
    const header = JSON.parse(Buffer.from(headerB64, 'base64url').toString('utf-8'));
    const payload = JSON.parse(Buffer.from(payloadB64, 'base64url').toString('utf-8'));
    return { header, payload };
  } catch {
    return null;
  }
}

/**
 * Generate JWKS (JSON Web Key Set) for public key distribution
 */
export function getJwks(): JWKS {
  const { publicKey: key, keyId: kid } = ensureKeys();

  const jwk = key.export({ format: 'jwk' }) as { n: string; e: string };

  const publicJwk: JWK = {
    kty: 'RSA',
    use: 'sig',
    alg: 'RS256',
    kid,
    n: jwk.n,
    e: jwk.e,
  };

  return { keys: [publicJwk] };
}

/**
 * Create an access token
 */
export function createAccessToken(params: {
  sessionId: string;
  clientId: string;
  userId: string;
  userEmail?: string;
  scope: string;
  lifetimeSecs?: number;
}): string {
  const now = Math.floor(Date.now() / 1000);
  const lifetime = params.lifetimeSecs ?? config.oauthAccessTokenLifetimeSecs;

  const payload: AccessTokenPayload = {
    iss: config.baseUrl,
    sub: params.sessionId,
    aud: params.clientId,
    exp: now + lifetime,
    iat: now,
    jti: crypto.randomUUID(),
    scope: params.scope,
    userId: params.userId,
    userEmail: params.userEmail,
  };

  return signJwt(payload);
}

/**
 * Get key ID for JWKS
 */
export function getKeyId(): string {
  const { keyId: kid } = ensureKeys();
  return kid;
}
