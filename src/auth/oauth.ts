import { ConfidentialClientApplication, Configuration } from '@azure/msal-node';
import { config, GRAPH_SCOPES, getOAuthEndpoints } from '../utils/config.js';
import { logger } from '../utils/logger.js';
import { generatePKCEPair, generateState, generateNonce, type PKCEPair } from './pkce.js';
import { sessionManager, type TokenSet, type UserSession } from './session.js';
import { createMsalCachePlugin } from './msal-cache-plugin.js';

const createMsalConfig = (sessionId: string): Configuration => ({
  auth: {
    clientId: config.azureClientId,
    clientSecret: config.azureClientSecret,
    authority: `https://login.microsoftonline.com/${config.azureTenantId}`,
  },
  cache: {
    cachePlugin: createMsalCachePlugin(sessionId),
  },
  system: {
    loggerOptions: {
      loggerCallback: (level, message) => {
        // Map MSAL log levels to pino
        const levelMap: Record<number, 'error' | 'warn' | 'info' | 'debug' | 'trace'> = {
          0: 'error',
          1: 'warn',
          2: 'info',
          3: 'debug',
          4: 'trace',
        };
        const pinoLevel = levelMap[level] ?? 'info';
        logger[pinoLevel]({ source: 'msal', sessionId }, message);
      },
      piiLoggingEnabled: false, // Never log PII
    },
  },
});

export interface AuthorizationUrlResult {
  url: string;
  session: UserSession;
}

export interface TokenExchangeResult {
  tokens: TokenSet;
  userId?: string;
  userEmail?: string;
  userDisplayName?: string;
}

/**
 * Extract the Azure AD refresh token from MSAL's serialized cache.
 * The RT is not exposed on AuthenticationResult, but lives inside
 * the serialized cache under the RefreshToken key.
 */
function extractRefreshTokenFromCache(msalClient: ConfidentialClientApplication): string | undefined {
  try {
    const cacheData = msalClient.getTokenCache().serialize();
    const cacheJson = JSON.parse(cacheData) as {
      RefreshToken?: Record<string, { secret?: string }>;
    };
    const rtEntries = Object.values(cacheJson.RefreshToken ?? {});
    if (rtEntries.length > 0 && rtEntries[0]?.secret) {
      return rtEntries[0].secret;
    }
  } catch (err) {
    logger.debug(
      { err: err instanceof Error ? err.message : String(err) },
      'Could not extract refresh token from MSAL cache'
    );
  }
  return undefined;
}

export class OAuthClient {
  private endpoints = getOAuthEndpoints(config.azureTenantId);

  /**
   * Helper to create an isolated MSAL client for a session
   */
  private getMsalClient(sessionId: string): ConfidentialClientApplication {
    return new ConfidentialClientApplication(createMsalConfig(sessionId));
  }

  /**
   * Generate authorization URL with PKCE
   */
  async getAuthorizationUrl(
    redirectUri: string,
    existingSessionId?: string
  ): Promise<AuthorizationUrlResult> {
    const pkce: PKCEPair = generatePKCEPair();
    const state = generateState();
    const nonce = generateNonce();

    // Create or update session with PKCE data
    let session: UserSession;
    if (existingSessionId) {
      const existing = await sessionManager.getSession(existingSessionId);
      if (existing) {
        existing.pkceVerifier = pkce.codeVerifier;
        existing.state = state;
        existing.nonce = nonce;
        await sessionManager.saveSession(existing);
        session = existing;
      } else {
        session = await sessionManager.createSession({
          pkceVerifier: pkce.codeVerifier,
          state,
          nonce,
        });
      }
    } else {
      session = await sessionManager.createSession({
        pkceVerifier: pkce.codeVerifier,
        state,
        nonce,
      });
    }

    // Build authorization URL
    const params = new URLSearchParams({
      client_id: config.azureClientId,
      response_type: 'code',
      redirect_uri: redirectUri,
      response_mode: 'query',
      scope: GRAPH_SCOPES.join(' '),
      state: state,
      nonce: nonce,
      code_challenge: pkce.codeChallenge,
      code_challenge_method: pkce.codeChallengeMethod,
      prompt: 'select_account', // Allow account selection
    });

    const url = `${this.endpoints.authorize}?${params.toString()}`;

    logger.info({ redirectUri, sessionId: session.id }, 'Generated authorization URL');

    return { url, session };
  }

  /**
   * Exchange authorization code for tokens
   */
  async exchangeCodeForTokens(
    code: string,
    redirectUri: string,
    session: UserSession
  ): Promise<TokenExchangeResult> {
    // Get decrypted PKCE verifier
    const codeVerifier = sessionManager.getDecryptedPkceVerifier(session);
    if (!codeVerifier) {
      throw new Error('PKCE verifier not found in session');
    }

    const msalClient = this.getMsalClient(session.id);

    try {
      const result = await msalClient.acquireTokenByCode({
        code,
        redirectUri,
        scopes: GRAPH_SCOPES,
        codeVerifier,
      });

      if (!result) {
        throw new Error('Token acquisition returned null');
      }

      // Extract refresh token from MSAL cache as fallback for cache loss scenarios
      const azureRefreshToken = extractRefreshTokenFromCache(msalClient);

      const tokens: TokenSet = {
        accessToken: result.accessToken,
        refreshToken: azureRefreshToken,
        expiresAt: result.expiresOn?.getTime() ?? Date.now() + 3600 * 1000,
        scope: result.scopes.join(' '),
      };

      // Extract user info from ID token claims
      const account = result.account;
      const userId = account?.localAccountId;
      const userEmail = account?.username;
      const userDisplayName = account?.name;

      logger.info(
        { userId, sessionId: session.id, hasRefreshToken: !!azureRefreshToken },
        'Successfully exchanged code for tokens'
      );

      return {
        tokens,
        userId,
        userEmail,
        userDisplayName,
      };
    } catch (err) {
      logger.error(
        { err: err instanceof Error ? { message: err.message, name: err.name } : { message: String(err) }, sessionId: session.id },
        'Failed to exchange code for tokens'
      );
      throw err;
    }
  }

  /**
   * Refresh access token using refresh token
   */
  async refreshTokens(session: UserSession): Promise<TokenSet> {
    const currentTokens = sessionManager.getDecryptedTokens(session);

    if (!currentTokens) {
      throw new Error('No tokens found in session');
    }

    const msalClient = this.getMsalClient(session.id);

    try {
      // MSAL caches accounts, try to get silently
      const accounts = await msalClient.getTokenCache().getAllAccounts();
      const account = accounts.find((a) => a.localAccountId === session.userId);

      if (!account) {
        // Fallback: use stored refresh token to recover from MSAL cache loss
        if (currentTokens.refreshToken) {
          logger.info({ sessionId: session.id }, 'MSAL cache miss - using stored refresh token as fallback');

          const fallbackResult = await msalClient.acquireTokenByRefreshToken({
            scopes: GRAPH_SCOPES,
            refreshToken: currentTokens.refreshToken,
            forceCache: true, // Repopulate MSAL cache with the new account/tokens
          });

          if (!fallbackResult) {
            throw new Error('Fallback token refresh returned null - re-authentication required');
          }

          // Azure AD may rotate the refresh token; extract the latest one
          const newRefreshToken = extractRefreshTokenFromCache(msalClient) ?? currentTokens.refreshToken;

          return {
            accessToken: fallbackResult.accessToken,
            refreshToken: newRefreshToken,
            expiresAt: fallbackResult.expiresOn?.getTime() ?? Date.now() + 3600 * 1000,
            scope: fallbackResult.scopes.join(' '),
          };
        }

        throw new Error('Account not found in cache - re-authentication required');
      }

      const result = await msalClient.acquireTokenSilent({
        account,
        scopes: GRAPH_SCOPES,
        forceRefresh: true,
      });

      if (!result) {
        throw new Error('Silent token acquisition returned null');
      }

      // Update stored refresh token in case it was rotated
      const latestRefreshToken = extractRefreshTokenFromCache(msalClient) ?? currentTokens.refreshToken;

      const tokens: TokenSet = {
        accessToken: result.accessToken,
        refreshToken: latestRefreshToken,
        expiresAt: result.expiresOn?.getTime() ?? Date.now() + 3600 * 1000,
        scope: result.scopes.join(' '),
      };

      logger.info({ userId: session.userId, sessionId: session.id }, 'Successfully refreshed tokens');

      return tokens;
    } catch (err) {
      logger.error(
        { err: err instanceof Error ? { message: err.message, name: err.name } : { message: String(err) }, sessionId: session.id },
        'Failed to refresh tokens'
      );
      throw err;
    }
  }

  /**
   * Get logout URL
   */
  getLogoutUrl(postLogoutRedirectUri?: string): string {
    const params = new URLSearchParams();
    if (postLogoutRedirectUri) {
      params.set('post_logout_redirect_uri', postLogoutRedirectUri);
    }

    const queryString = params.toString();
    return queryString
      ? `${this.endpoints.logout}?${queryString}`
      : this.endpoints.logout;
  }

  /**
   * Validate state parameter from callback
   */
  validateState(session: UserSession, state: string): boolean {
    return session.state === state;
  }
}

export const oauthClient = new OAuthClient();
