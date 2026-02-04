/**
 * OAuth 2.1 Authorization Server Module
 *
 * Implements RFC 6749 (OAuth 2.0), RFC 7636 (PKCE), RFC 7591 (Dynamic Client Registration),
 * RFC 8414 (Authorization Server Metadata), with OAuth 2.1 security requirements.
 */

// Types
export type {
  OAuthClient,
  ClientRegistrationRequest,
  ClientRegistrationResponse,
  AuthorizationCode,
  RefreshTokenRecord,
  AccessTokenPayload,
  TokenResponse,
  TokenErrorResponse,
  AuthorizationServerMetadata,
  JWKS,
  JWK,
  GrantType,
  ResponseType,
  TokenEndpointAuthMethod,
  CodeChallengeMethod,
  TokenError,
  AuthorizationError,
  AuthorizationRequest,
  AuthorizationCodeTokenRequest,
  RefreshTokenRequest,
  TokenRequest,
  PendingAuthorization,
} from './types.js';

// JWT utilities
export {
  initializeKeys,
  signJwt,
  verifyJwt,
  decodeJwt,
  getJwks,
  createAccessToken,
  getKeyId,
} from './jwt.js';

// Client management
export {
  registerClient,
  getClient,
  authenticateClient,
  validateClientRedirectUri,
  deleteClient,
  verifyClientSecret,
} from './client-store.js';

// Authorization codes
export {
  createAuthorizationCode,
  consumeAuthorizationCode,
  verifyCodeChallenge,
  storePendingAuthorization,
  getPendingAuthorization,
  deletePendingAuthorization,
} from './code-store.js';

// Refresh tokens
export {
  createRefreshToken,
  validateRefreshToken,
  rotateRefreshToken,
  revokeRefreshToken,
  revokeSessionTokens,
  getSessionTokens,
} from './token-store.js';

// Express router and middleware
export { oauthRouter } from './routes.js';
export {
  bearerAuthMiddleware,
  requireBearerAuth,
  requireAuth,
  requireScope,
  getAuthenticatedUser,
} from './middleware.js';
