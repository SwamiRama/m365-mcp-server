import crypto from 'crypto';

/**
 * PKCE (Proof Key for Code Exchange) implementation per RFC 7636
 */

export interface PKCEPair {
  codeVerifier: string;
  codeChallenge: string;
  codeChallengeMethod: 'S256';
}

/**
 * Generate a cryptographically random code verifier
 * Per RFC 7636: 43-128 characters from unreserved URI characters
 */
export function generateCodeVerifier(): string {
  // 32 bytes = 43 base64url characters (minimum allowed)
  // Using 48 bytes = 64 characters for extra entropy
  return crypto.randomBytes(48).toString('base64url');
}

/**
 * Generate code challenge from code verifier using SHA-256
 * Per RFC 7636: BASE64URL(SHA256(code_verifier))
 */
export function generateCodeChallenge(codeVerifier: string): string {
  return crypto
    .createHash('sha256')
    .update(codeVerifier, 'ascii')
    .digest('base64url');
}

/**
 * Generate a PKCE pair (verifier + challenge)
 */
export function generatePKCEPair(): PKCEPair {
  const codeVerifier = generateCodeVerifier();
  const codeChallenge = generateCodeChallenge(codeVerifier);

  return {
    codeVerifier,
    codeChallenge,
    codeChallengeMethod: 'S256',
  };
}

/**
 * Generate a cryptographically secure state parameter
 */
export function generateState(): string {
  return crypto.randomBytes(32).toString('hex');
}

/**
 * Generate a cryptographically secure nonce for OIDC
 */
export function generateNonce(): string {
  return crypto.randomBytes(32).toString('base64url');
}
