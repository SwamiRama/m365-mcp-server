import { ConfidentialClientApplication, Configuration } from '@azure/msal-node';
import { config, GRAPH_SCOPES, getOAuthEndpoints } from '../utils/config.js';
import { logger } from '../utils/logger.js';
import { generatePKCEPair, generateState, generateNonce, type PKCEPair } from './pkce.js';
import { sessionManager, type TokenSet, type UserSession } from './session.js';
import { createMsalCachePlugin } from './msal-cache-plugin.js';

const msalConfig: Configuration = {
  auth: {
    clientId: config.azureClientId,
    clientSecret: config.azureClientSecret,
    authority: `https://login.microsoftonline.com/${config.azureTenantId}`,
  },
  cache: {
    cachePlugin: createMsalCachePlugin(),
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
        logger[pinoLevel]({ source: 'msal' }, message);
      },
      piiLoggingEnabled: false, // Never log PII
    },
  },
};

const msalClient = new ConfidentialClientApplication(msalConfig);

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

export class OAuthClient {
  private endpoints = getOAuthEndpoints(config.azureTenantId);

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

    logger.info({ redirectUri }, 'Generated authorization URL');

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

      const tokens: TokenSet = {
        accessToken: result.accessToken,
        refreshToken: undefined, // MSAL handles refresh internally, but we need it for manual control
        expiresAt: result.expiresOn?.getTime() ?? Date.now() + 3600 * 1000,
        scope: result.scopes.join(' '),
      };

      // Extract user info from ID token claims
      const account = result.account;
      const userId = account?.localAccountId;
      const userEmail = account?.username;
      const userDisplayName = account?.name;

      logger.info(
        { userId, hasRefreshToken: false },
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
        { err: err instanceof Error ? { message: err.message, name: err.name } : { message: String(err) } },
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

    try {
      // MSAL caches accounts, try to get silently
      const accounts = await msalClient.getTokenCache().getAllAccounts();
      const account = accounts.find((a) => a.localAccountId === session.userId);

      if (!account) {
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

      const tokens: TokenSet = {
        accessToken: result.accessToken,
        refreshToken: currentTokens.refreshToken, // Keep existing refresh token
        expiresAt: result.expiresOn?.getTime() ?? Date.now() + 3600 * 1000,
        scope: result.scopes.join(' '),
      };

      logger.info({ userId: session.userId }, 'Successfully refreshed tokens');

      return tokens;
    } catch (err) {
      logger.error(
        { err: err instanceof Error ? { message: err.message, name: err.name } : { message: String(err) } },
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
