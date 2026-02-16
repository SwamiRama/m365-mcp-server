/**
 * OAuth 2.1 Bearer Token Middleware
 */

import { Request, Response, NextFunction } from 'express';
import { verifyJwt } from './jwt.js';
import { sessionManager } from '../auth/session.js';
import type { AccessTokenPayload } from './types.js';
import { config } from '../utils/config.js';

/**
 * Build WWW-Authenticate header value pointing to our OAuth metadata.
 * Required by RFC 6750 / MCP spec so clients can discover the authorization server.
 */
function wwwAuthenticateHeader(): string {
  return `Bearer resource_metadata="${config.baseUrl}/.well-known/oauth-authorization-server"`;
}

/**
 * Extended request with OAuth info
 */
declare global {
  namespace Express {
    interface Request {
      oauthToken?: AccessTokenPayload;
    }
  }
}

/**
 * Extract Bearer token from Authorization header
 */
function extractBearerToken(req: Request): string | null {
  const authHeader = req.headers.authorization;

  if (!authHeader) {
    return null;
  }

  if (!authHeader.startsWith('Bearer ')) {
    return null;
  }

  return authHeader.slice(7);
}

/**
 * Bearer token authentication middleware
 *
 * Validates JWT Bearer tokens and loads the associated session.
 * Sets req.session and req.oauthToken on success.
 *
 * Does NOT fail the request if no token is present - allows
 * fallback to session cookie authentication.
 */
export async function bearerAuthMiddleware(
  req: Request,
  res: Response,
  next: NextFunction
): Promise<void> {
  const token = extractBearerToken(req);

  if (!token) {
    // No Bearer token - continue without OAuth authentication
    // Session cookie middleware will handle session-based auth
    next();
    return;
  }

  // Verify JWT
  const payload = verifyJwt(token);

  if (!payload) {
    res.setHeader('WWW-Authenticate', wwwAuthenticateHeader());
    res.status(401).json({
      error: 'invalid_token',
      error_description: 'Bearer token is invalid or expired',
    });
    return;
  }

  // Validate issuer
  // Note: We don't check issuer here since verifyJwt already validates signature
  // which proves the token was issued by us

  // Load session from token subject (session ID)
  const session = await sessionManager.getSession(payload.sub);

  if (!session) {
    res.setHeader('WWW-Authenticate', wwwAuthenticateHeader());
    res.status(401).json({
      error: 'invalid_token',
      error_description: 'Session associated with token no longer valid',
    });
    return;
  }

  // Verify session has tokens (is authenticated with Azure AD)
  if (!session.tokens) {
    res.setHeader('WWW-Authenticate', wwwAuthenticateHeader());
    res.status(401).json({
      error: 'invalid_token',
      error_description: 'Session not authenticated',
    });
    return;
  }

  // Attach to request
  req.session = session;
  req.oauthToken = payload;

  req.log.debug(
    { userId: payload.userId, clientId: payload.aud },
    'Bearer token authenticated'
  );

  next();
}

/**
 * Require Bearer token authentication
 *
 * Use this middleware on routes that REQUIRE OAuth Bearer authentication
 * and should not accept session cookie auth.
 */
export async function requireBearerAuth(
  req: Request,
  res: Response,
  next: NextFunction
): Promise<void> {
  const token = extractBearerToken(req);

  if (!token) {
    res.setHeader('WWW-Authenticate', wwwAuthenticateHeader());
    res.status(401).json({
      error: 'invalid_request',
      error_description: 'Bearer token required',
    });
    return;
  }

  // Delegate to bearerAuthMiddleware
  await bearerAuthMiddleware(req, res, next);
}

/**
 * Require authentication (Bearer token OR session cookie)
 *
 * Use this middleware on routes that require any form of authentication.
 * Prefers Bearer token but accepts session cookie as fallback.
 */
export function requireAuth(
  req: Request,
  res: Response,
  next: NextFunction
): void {
  if (!req.session || !req.session.tokens) {
    res.setHeader('WWW-Authenticate', wwwAuthenticateHeader());
    res.status(401).json({
      error: 'unauthorized',
      error_description: 'Authentication required',
    });
    return;
  }

  next();
}

/**
 * Scope validation middleware factory
 *
 * Returns middleware that checks if the Bearer token has the required scope.
 */
export function requireScope(requiredScope: string): (req: Request, res: Response, next: NextFunction) => void {
  return (req: Request, res: Response, next: NextFunction): void => {
    if (!req.oauthToken) {
      // No OAuth token - skip scope check (session-based auth)
      next();
      return;
    }

    const tokenScopes = req.oauthToken.scope.split(' ');

    if (!tokenScopes.includes(requiredScope)) {
      res.status(403).json({
        error: 'insufficient_scope',
        error_description: `Required scope: ${requiredScope}`,
      });
      return;
    }

    next();
  };
}

/**
 * Get authenticated user info from request
 */
export function getAuthenticatedUser(req: Request): {
  userId: string;
  userEmail?: string;
  sessionId: string;
  clientId?: string; // Only present for OAuth authentication
} | null {
  if (!req.session) {
    return null;
  }

  return {
    userId: req.oauthToken?.userId ?? req.session.userId ?? req.session.id,
    userEmail: req.oauthToken?.userEmail ?? req.session.userEmail,
    sessionId: req.session.id,
    clientId: req.oauthToken?.aud,
  };
}
