/**
 * OAuth 2.1 Authorization Server Routes
 */

import { Router, Request, Response } from 'express';
import rateLimit from 'express-rate-limit';
import { config, GRAPH_SCOPES, getOAuthEndpoints } from '../utils/config.js';
import { sessionManager } from '../auth/session.js';
import { generatePKCEPair, generateState, generateNonce } from '../auth/pkce.js';
import { oauthClient } from '../auth/oauth.js';
import { getJwks, createAccessToken, initializeKeys } from './jwt.js';
import {
  registerClient,
  getClient,
  authenticateClient,
  validateClientRedirectUri,
} from './client-store.js';
import {
  createAuthorizationCode,
  consumeAuthorizationCode,
  peekAuthorizationCode,
  verifyCodeChallenge,
  storePendingAuthorization,
  getPendingAuthorization,
  deletePendingAuthorization,
} from './code-store.js';
import { createRefreshToken, validateRefreshToken, rotateRefreshToken, revokeRefreshToken } from './token-store.js';
import type {
  AuthorizationServerMetadata,
  ClientRegistrationRequest,
  TokenResponse,
  TokenErrorResponse,
} from './types.js';

// Initialize JWT keys on module load
initializeKeys();

export const oauthRouter = Router();

// =============================================================================
// OAuth 2.1 Server Metadata (RFC 8414)
// =============================================================================

oauthRouter.get('/.well-known/oauth-authorization-server', (_req: Request, res: Response) => {
  const metadata: AuthorizationServerMetadata = {
    issuer: config.baseUrl,
    authorization_endpoint: `${config.baseUrl}/authorize`,
    token_endpoint: `${config.baseUrl}/token`,
    registration_endpoint: config.oauthAllowDynamicRegistration
      ? `${config.baseUrl}/register`
      : undefined,
    revocation_endpoint: `${config.baseUrl}/revoke`,
    jwks_uri: `${config.baseUrl}/.well-known/jwks.json`,
    response_types_supported: ['code'],
    response_modes_supported: ['query'],
    grant_types_supported: ['authorization_code', 'refresh_token'],
    token_endpoint_auth_methods_supported: ['none'],
    code_challenge_methods_supported: ['S256'],
    scopes_supported: ['openid', 'offline_access', 'mail.read', 'files.read'],
    service_documentation: 'https://github.com/anthropic/m365-mcp-server',
  };

  res.json(metadata);
});

// =============================================================================
// JWKS Endpoint
// =============================================================================

oauthRouter.get('/.well-known/jwks.json', (_req: Request, res: Response) => {
  res.json(getJwks());
});

// =============================================================================
// Dynamic Client Registration (RFC 7591)
// =============================================================================

// Rate limit for client registration: 20 per hour per IP (relaxed for idempotent DCR)
const dcrRateLimit = rateLimit({
  windowMs: 60 * 60 * 1000, // 1 hour
  max: 20,
  standardHeaders: true,
  legacyHeaders: false,
  message: {
    error: 'too_many_requests',
    error_description: 'Too many registration attempts. Try again later.',
  },
  validate: { trustProxy: false },
});

oauthRouter.post('/register', dcrRateLimit, async (req: Request, res: Response): Promise<void> => {
  if (!config.oauthAllowDynamicRegistration) {
    res.status(403).json({
      error: 'registration_not_supported',
      error_description: 'Dynamic client registration is disabled',
    });
    return;
  }

  try {
    const request = req.body as ClientRegistrationRequest;

    if (!request.client_name) {
      res.status(400).json({
        error: 'invalid_client_metadata',
        error_description: 'client_name is required',
      });
      return;
    }

    if (!request.redirect_uris || !Array.isArray(request.redirect_uris) || request.redirect_uris.length === 0) {
      res.status(400).json({
        error: 'invalid_redirect_uri',
        error_description: 'At least one redirect_uri is required',
      });
      return;
    }

    // Validate redirect URIs against allowed patterns (if configured)
    if (config.oauthAllowedRedirectPatterns) {
      const patterns = config.oauthAllowedRedirectPatterns
        .split(',')
        .map((p) => p.trim())
        .filter(Boolean);
      if (patterns.length > 0) {
        for (const uri of request.redirect_uris) {
          if (!matchesAnyPattern(uri, patterns)) {
            req.log.warn(
              { event: 'dcr.rejected', redirectUri: uri, ip: req.ip },
              'DCR rejected: redirect_uri does not match allowed patterns'
            );
            res.status(400).json({
              error: 'invalid_redirect_uri',
              error_description: `Redirect URI "${uri}" is not allowed by server policy`,
            });
            return;
          }
        }
      }
    }

    const response = await registerClient(request);

    req.log.info(
      {
        event: 'oauth.client_registered',
        clientId: response.client_id,
        clientName: request.client_name,
        redirectUris: request.redirect_uris,
        ip: req.ip,
      },
      'New OAuth client registered'
    );

    res.status(201).json(response);
  } catch (err) {
    const error = err instanceof Error ? err : new Error('Unknown error');
    req.log.error({ err }, 'Client registration failed');

    res.status(400).json({
      error: 'invalid_client_metadata',
      error_description: error.message,
    });
  }
});

// =============================================================================
// Authorization Endpoint (OAuth 2.1 with PKCE)
// =============================================================================

oauthRouter.get('/authorize', async (req: Request, res: Response): Promise<void> => {
  const {
    response_type,
    client_id,
    redirect_uri,
    scope,
    state,
    code_challenge,
    code_challenge_method,
  } = req.query;

  // Validate response_type
  if (response_type !== 'code') {
    res.status(400).json({
      error: 'unsupported_response_type',
      error_description: 'Only "code" response type is supported',
    });
    return;
  }

  // Validate client_id
  if (!client_id || typeof client_id !== 'string') {
    res.status(400).json({
      error: 'invalid_request',
      error_description: 'client_id is required',
    });
    return;
  }

  const client = await getClient(client_id);
  if (!client) {
    res.status(400).json({
      error: 'invalid_client',
      error_description: 'Unknown client',
    });
    return;
  }

  // Validate redirect_uri
  if (!redirect_uri || typeof redirect_uri !== 'string') {
    res.status(400).json({
      error: 'invalid_request',
      error_description: 'redirect_uri is required',
    });
    return;
  }

  if (!validateClientRedirectUri(client, redirect_uri)) {
    res.status(400).json({
      error: 'invalid_redirect_uri',
      error_description: 'Redirect URI not registered for this client',
    });
    return;
  }

  // Validate redirect_uri is a well-formed URL before using it in redirects
  try {
    new URL(redirect_uri);
  } catch {
    res.status(400).json({
      error: 'invalid_redirect_uri',
      error_description: 'redirect_uri is not a valid URL',
    });
    return;
  }

  // From here on, we can redirect errors to the client

  // Validate state (required for CSRF protection)
  if (!state || typeof state !== 'string') {
    const errorUrl = new URL(redirect_uri);
    errorUrl.searchParams.set('error', 'invalid_request');
    errorUrl.searchParams.set('error_description', 'state parameter is required');
    res.redirect(errorUrl.toString());
    return;
  }

  // Validate PKCE (required in OAuth 2.1)
  if (!code_challenge || typeof code_challenge !== 'string') {
    const errorUrl = new URL(redirect_uri);
    errorUrl.searchParams.set('error', 'invalid_request');
    errorUrl.searchParams.set('error_description', 'code_challenge is required (PKCE)');
    errorUrl.searchParams.set('state', state);
    res.redirect(errorUrl.toString());
    return;
  }

  if (code_challenge_method !== 'S256') {
    const errorUrl = new URL(redirect_uri);
    errorUrl.searchParams.set('error', 'invalid_request');
    errorUrl.searchParams.set('error_description', 'code_challenge_method must be S256');
    errorUrl.searchParams.set('state', state);
    res.redirect(errorUrl.toString());
    return;
  }

  // Parse scope (default to our standard scopes)
  const requestedScope = typeof scope === 'string' ? scope : GRAPH_SCOPES.join(' ');

  // Create session for this authorization flow
  const session = await sessionManager.createSession();

  // Store pending authorization data
  await storePendingAuthorization(session.id, {
    clientId: client_id,
    redirectUri: redirect_uri,
    scope: requestedScope,
    state: state,
    codeChallenge: code_challenge,
    codeChallengeMethod: 'S256',
  });

  // Generate PKCE for Azure AD
  const pkce = generatePKCEPair();
  const azureState = generateState();
  const nonce = generateNonce();

  // Store PKCE verifier in session for Azure callback
  session.pkceVerifier = pkce.codeVerifier;
  session.state = azureState;
  session.nonce = nonce;
  await sessionManager.saveSession(session);

  // Set session cookie
  res.cookie('mcp-oauth-session', session.id, {
    httpOnly: true,
    secure: config.nodeEnv === 'production',
    sameSite: 'lax',
    maxAge: 10 * 60 * 1000, // 10 minutes for auth flow
  });

  // Build Azure AD authorization URL
  const endpoints = getOAuthEndpoints(config.azureTenantId);
  const params = new URLSearchParams({
    client_id: config.azureClientId,
    response_type: 'code',
    redirect_uri: `${config.baseUrl}/oauth/callback`,
    response_mode: 'query',
    scope: GRAPH_SCOPES.join(' '),
    state: azureState,
    nonce: nonce,
    code_challenge: pkce.codeChallenge,
    code_challenge_method: 'S256',
    prompt: 'select_account',
  });

  const azureAuthUrl = `${endpoints.authorize}?${params.toString()}`;

  req.log.info({ clientId: client_id, sessionId: session.id }, 'Redirecting to Azure AD for authorization');

  res.redirect(azureAuthUrl);
});

// =============================================================================
// OAuth Callback from Azure AD
// =============================================================================

oauthRouter.get('/oauth/callback', async (req: Request, res: Response): Promise<void> => {
  const { code, state, error, error_description } = req.query;

  // Get session from cookie
  const sessionId = req.cookies?.['mcp-oauth-session'] as string | undefined;
  if (!sessionId) {
    req.log.warn({ event: 'oauth.callback_no_cookie', ip: req.ip }, 'OAuth callback: session cookie missing');
    res.status(400).json({
      error: 'invalid_session',
      error_description: 'OAuth session not found',
    });
    return;
  }

  const session = await sessionManager.getSession(sessionId);
  if (!session) {
    req.log.warn({ event: 'oauth.callback_session_expired', sessionId, ip: req.ip }, 'OAuth callback: session not found in store');
    res.status(400).json({
      error: 'invalid_session',
      error_description: 'Session expired or invalid',
    });
    return;
  }

  // Get pending authorization
  const pending = await getPendingAuthorization(session.id);
  if (!pending) {
    req.log.warn({ event: 'oauth.callback_no_pending', sessionId: session.id, ip: req.ip }, 'OAuth callback: no pending authorization');
    res.status(400).json({
      error: 'invalid_request',
      error_description: 'No pending authorization found',
    });
    return;
  }

  // Handle Azure AD errors
  if (error) {
    req.log.warn({ error, error_description }, 'Azure AD authorization error');

    const errorUrl = new URL(pending.redirectUri);
    errorUrl.searchParams.set('error', 'access_denied');
    errorUrl.searchParams.set('error_description', (error_description as string) || 'Authorization denied');
    errorUrl.searchParams.set('state', pending.state);
    res.redirect(errorUrl.toString());
    return;
  }

  // Validate authorization code
  if (!code || typeof code !== 'string') {
    const errorUrl = new URL(pending.redirectUri);
    errorUrl.searchParams.set('error', 'server_error');
    errorUrl.searchParams.set('error_description', 'Missing authorization code from identity provider');
    errorUrl.searchParams.set('state', pending.state);
    res.redirect(errorUrl.toString());
    return;
  }

  // Validate state
  if (!state || !oauthClient.validateState(session, state as string)) {
    const errorUrl = new URL(pending.redirectUri);
    errorUrl.searchParams.set('error', 'invalid_request');
    errorUrl.searchParams.set('error_description', 'State mismatch');
    errorUrl.searchParams.set('state', pending.state);
    res.redirect(errorUrl.toString());
    return;
  }

  try {
    // Exchange Azure code for tokens
    const redirectUri = `${config.baseUrl}/oauth/callback`;
    const tokenResult = await oauthClient.exchangeCodeForTokens(code, redirectUri, session);

    // Update session with Azure tokens and user info
    session.tokens = tokenResult.tokens;
    session.userId = tokenResult.userId;
    session.userEmail = tokenResult.userEmail;
    session.userDisplayName = tokenResult.userDisplayName;
    session.pkceVerifier = undefined;
    session.state = undefined;
    session.nonce = undefined;
    await sessionManager.saveSession(session);

    // Generate MCP authorization code for the OAuth client
    const mcpCode = await createAuthorizationCode({
      clientId: pending.clientId,
      redirectUri: pending.redirectUri,
      scope: pending.scope,
      codeChallenge: pending.codeChallenge,
      codeChallengeMethod: pending.codeChallengeMethod,
      sessionId: session.id,
      userId: tokenResult.userId || session.id,
      userEmail: tokenResult.userEmail,
    });

    // Clean up pending authorization
    await deletePendingAuthorization(session.id);

    // Clear OAuth session cookie (no longer needed)
    res.clearCookie('mcp-oauth-session');

    // Redirect back to client with authorization code
    const successUrl = new URL(pending.redirectUri);
    successUrl.searchParams.set('code', mcpCode);
    successUrl.searchParams.set('state', pending.state);

    req.log.info(
      { clientId: pending.clientId, userId: tokenResult.userId },
      'Authorization successful, redirecting to client'
    );

    res.redirect(successUrl.toString());
  } catch (err) {
    req.log.error({ err, clientId: pending.clientId, redirectUri: pending.redirectUri }, 'Token exchange failed');

    const errorUrl = new URL(pending.redirectUri);
    errorUrl.searchParams.set('error', 'server_error');
    errorUrl.searchParams.set('error_description', 'Failed to complete authorization');
    errorUrl.searchParams.set('state', pending.state);
    res.redirect(errorUrl.toString());
  }
});

// =============================================================================
// Token Endpoint
// =============================================================================

oauthRouter.post('/token', async (req: Request, res: Response): Promise<void> => {
  const { grant_type, code } = req.body;

  // Extract client credentials
  const credentials = extractClientCredentials(req);
  let clientId = credentials.clientId;
  const clientSecret = credentials.clientSecret;

  // For authorization_code grant, client_id can be extracted from the code
  // This supports clients like Open WebUI that don't send client_id in token request
  if (!clientId && grant_type === 'authorization_code' && code) {
    const authCode = await peekAuthorizationCode(code);
    if (authCode) {
      clientId = authCode.clientId;
      req.log.debug({ clientId }, 'Extracted client_id from authorization code');
    }
  }

  if (!clientId) {
    sendTokenError(res, 'invalid_client', 'client_id is required');
    return;
  }

  // Get client first to check auth method
  const client = await getClient(clientId);

  if (!client) {
    sendTokenError(res, 'invalid_client', 'Unknown client');
    return;
  }

  // Validate client authentication based on token_endpoint_auth_method
  if (client.tokenEndpointAuthMethod === 'none') {
    // Public client - no secret required, PKCE provides security
    // This is common for SPAs and native apps like Open WebUI
  } else if (clientSecret) {
    // Confidential client with secret provided - verify it
    const authenticated = await authenticateClient(clientId, clientSecret);
    if (!authenticated) {
      sendTokenError(res, 'invalid_client', 'Invalid client credentials');
      return;
    }
  } else {
    // Confidential client but no secret provided
    sendTokenError(res, 'invalid_client', 'Client authentication required');
    return;
  }

  switch (grant_type) {
    case 'authorization_code':
      await handleAuthorizationCodeGrant(req, res, client.clientId);
      break;

    case 'refresh_token':
      await handleRefreshTokenGrant(req, res, client.clientId);
      break;

    default:
      sendTokenError(res, 'unsupported_grant_type', `Grant type "${grant_type}" is not supported`);
  }
});

/**
 * Extract client credentials from request (Basic auth or body)
 */
function extractClientCredentials(req: Request): { clientId: string | null; clientSecret: string | null } {
  // Try Basic auth header first
  const authHeader = req.headers.authorization;
  if (authHeader?.startsWith('Basic ')) {
    const credentials = Buffer.from(authHeader.slice(6), 'base64').toString('utf-8');
    const [clientId, clientSecret] = credentials.split(':');
    return { clientId: clientId || null, clientSecret: clientSecret || null };
  }

  // Fall back to body params (client_secret_post)
  const { client_id, client_secret } = req.body;
  return {
    clientId: typeof client_id === 'string' ? client_id : null,
    clientSecret: typeof client_secret === 'string' ? client_secret : null,
  };
}

/**
 * Handle authorization_code grant
 */
async function handleAuthorizationCodeGrant(
  req: Request,
  res: Response,
  clientId: string
): Promise<void> {
  const { code, redirect_uri, code_verifier } = req.body;

  // Validate required parameters
  if (!code || typeof code !== 'string') {
    sendTokenError(res, 'invalid_request', 'Authorization code is required');
    return;
  }

  if (!redirect_uri || typeof redirect_uri !== 'string') {
    sendTokenError(res, 'invalid_request', 'redirect_uri is required');
    return;
  }

  if (!code_verifier || typeof code_verifier !== 'string') {
    sendTokenError(res, 'invalid_request', 'code_verifier is required (PKCE)');
    return;
  }

  // Consume authorization code
  const authCode = await consumeAuthorizationCode(code);
  if (!authCode) {
    sendTokenError(res, 'invalid_grant', 'Authorization code is invalid or expired');
    return;
  }

  // Validate client ID matches
  if (authCode.clientId !== clientId) {
    req.log.warn(
      { event: 'oauth.client_id_mismatch', expected: authCode.clientId, received: clientId, grant: 'authorization_code' },
      'Client ID mismatch in authorization_code grant'
    );
    sendTokenError(res, 'invalid_grant', 'Client ID mismatch');
    return;
  }

  // Validate redirect URI matches
  if (authCode.redirectUri !== redirect_uri) {
    sendTokenError(res, 'invalid_grant', 'Redirect URI mismatch');
    return;
  }

  // Verify PKCE code verifier
  if (!verifyCodeChallenge(code_verifier, authCode.codeChallenge, authCode.codeChallengeMethod)) {
    sendTokenError(res, 'invalid_grant', 'PKCE verification failed');
    return;
  }

  // Generate tokens
  const accessToken = createAccessToken({
    sessionId: authCode.sessionId,
    clientId,
    userId: authCode.userId,
    userEmail: authCode.userEmail,
    scope: authCode.scope,
  });

  const refreshToken = await createRefreshToken({
    clientId,
    sessionId: authCode.sessionId,
    userId: authCode.userId,
    scope: authCode.scope,
  });

  const response: TokenResponse = {
    access_token: accessToken,
    token_type: 'Bearer',
    expires_in: config.oauthAccessTokenLifetimeSecs,
    refresh_token: refreshToken,
    scope: authCode.scope,
  };

  req.log.info({ clientId, userId: authCode.userId }, 'Issued tokens via authorization_code grant');

  res.json(response);
}

/**
 * Handle refresh_token grant
 */
async function handleRefreshTokenGrant(
  req: Request,
  res: Response,
  clientId: string
): Promise<void> {
  const { refresh_token, scope } = req.body;

  if (!refresh_token || typeof refresh_token !== 'string') {
    sendTokenError(res, 'invalid_request', 'refresh_token is required');
    return;
  }

  // Validate refresh token
  const tokenRecord = await validateRefreshToken(refresh_token);
  if (!tokenRecord) {
    sendTokenError(res, 'invalid_grant', 'Refresh token is invalid or expired');
    return;
  }

  // Validate client ID matches
  if (tokenRecord.clientId !== clientId) {
    req.log.warn(
      { event: 'oauth.client_id_mismatch', expected: tokenRecord.clientId, received: clientId, grant: 'refresh_token' },
      'Client ID mismatch in refresh_token grant'
    );
    sendTokenError(res, 'invalid_grant', 'Client ID mismatch');
    return;
  }

  // Check if session still exists
  const session = await sessionManager.getSession(tokenRecord.sessionId);
  if (!session) {
    sendTokenError(res, 'invalid_grant', 'Session no longer valid');
    return;
  }

  // Use requested scope or original scope
  const tokenScope = (typeof scope === 'string' ? scope : tokenRecord.scope);

  // Generate new access token
  const accessToken = createAccessToken({
    sessionId: tokenRecord.sessionId,
    clientId,
    userId: tokenRecord.userId,
    userEmail: session.userEmail,
    scope: tokenScope,
  });

  // Rotate refresh token
  const newRefreshToken = await rotateRefreshToken(refresh_token, {
    clientId,
    sessionId: tokenRecord.sessionId,
    userId: tokenRecord.userId,
    scope: tokenScope,
  });

  if (!newRefreshToken) {
    sendTokenError(res, 'invalid_grant', 'Failed to rotate refresh token');
    return;
  }

  const response: TokenResponse = {
    access_token: accessToken,
    token_type: 'Bearer',
    expires_in: config.oauthAccessTokenLifetimeSecs,
    refresh_token: newRefreshToken,
    scope: tokenScope,
  };

  req.log.info({ clientId, userId: tokenRecord.userId }, 'Issued tokens via refresh_token grant');

  res.json(response);
}

/**
 * Send token error response
 */
function sendTokenError(
  res: Response,
  error: TokenErrorResponse['error'],
  description: string
): void {
  res.status(400).json({
    error,
    error_description: description,
  } as TokenErrorResponse);
}

// =============================================================================
// Token Revocation Endpoint (RFC 7009)
// =============================================================================

oauthRouter.post('/revoke', async (req: Request, res: Response): Promise<void> => {
  const { token, token_type_hint } = req.body;

  if (!token || typeof token !== 'string') {
    // Per RFC 7009, invalid tokens should still return 200
    res.status(200).send();
    return;
  }

  try {
    if (token_type_hint === 'refresh_token' || !token_type_hint) {
      // Try to revoke as refresh token
      await revokeRefreshToken(token);
    }

    // Access tokens are JWTs and are stateless - they cannot be directly revoked
    // They will expire naturally based on their exp claim
    // For stronger revocation, the session can be deleted

    req.log.info({ event: 'oauth.token_revoked', tokenTypeHint: token_type_hint }, 'Token revoked');

    // RFC 7009: Always return 200 OK regardless of whether the token was valid
    res.status(200).send();
  } catch (err) {
    req.log.error({ err }, 'Token revocation failed');
    // Per RFC 7009, errors during revocation should still return 200
    res.status(200).send();
  }
});

// =============================================================================
// Utility functions
// =============================================================================

/**
 * Check if a URL matches any of the given glob-like patterns.
 * Supports `**` for multi-segment wildcard (any characters including `/`)
 * and `*` for single-segment wildcard (any characters except `/`).
 */
function matchesAnyPattern(url: string, patterns: string[]): boolean {
  return patterns.some((pattern) => {
    // Escape regex special chars (except *), then convert glob wildcards
    const regexStr = pattern
      .replace(/[.+?^${}()|[\]\\]/g, '\\$&')
      .replace(/\*\*/g, '@@GLOBSTAR@@')
      .replace(/\*/g, '[^/]*')
      .replace(/@@GLOBSTAR@@/g, '.*');
    try {
      return new RegExp(`^${regexStr}$`).test(url);
    } catch {
      return false;
    }
  });
}
