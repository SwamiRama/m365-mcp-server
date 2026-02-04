/**
 * OAuth 2.1 Authorization Server Types
 */

/**
 * Registered OAuth Client (RFC 7591)
 */
export interface OAuthClient {
  clientId: string;
  clientSecret: string; // Hashed
  clientSecretPlain?: string; // Only returned during registration
  clientName: string;
  redirectUris: string[];
  grantTypes: GrantType[];
  responseTypes: ResponseType[];
  tokenEndpointAuthMethod: TokenEndpointAuthMethod;
  scope?: string;
  createdAt: number;
}

/**
 * Client registration request (RFC 7591)
 */
export interface ClientRegistrationRequest {
  client_name: string;
  redirect_uris: string[];
  grant_types?: GrantType[];
  response_types?: ResponseType[];
  token_endpoint_auth_method?: TokenEndpointAuthMethod;
  scope?: string;
}

/**
 * Client registration response (RFC 7591)
 */
export interface ClientRegistrationResponse {
  client_id: string;
  client_secret: string;
  client_name: string;
  redirect_uris: string[];
  grant_types: GrantType[];
  response_types: ResponseType[];
  token_endpoint_auth_method: TokenEndpointAuthMethod;
  client_id_issued_at: number;
  client_secret_expires_at: number; // 0 = never expires
}

/**
 * Authorization Code with PKCE
 */
export interface AuthorizationCode {
  code: string;
  clientId: string;
  redirectUri: string;
  scope: string;
  codeChallenge: string;
  codeChallengeMethod: 'S256';
  sessionId: string; // MCP session with Azure tokens
  userId: string;
  userEmail?: string;
  expiresAt: number;
  used: boolean;
}

/**
 * Refresh Token Record
 */
export interface RefreshTokenRecord {
  token: string; // Hashed
  clientId: string;
  sessionId: string;
  userId: string;
  scope: string;
  expiresAt: number;
  rotatedFrom?: string; // Previous token (for rotation)
  createdAt: number;
}

/**
 * Access Token JWT Payload
 */
export interface AccessTokenPayload {
  iss: string; // Issuer (MCP server URL)
  sub: string; // Subject (session ID)
  aud: string; // Audience (client ID)
  exp: number; // Expiration time
  iat: number; // Issued at
  jti: string; // JWT ID
  scope: string;
  userId: string;
  userEmail?: string;
}

/**
 * Token Response (RFC 6749)
 */
export interface TokenResponse {
  access_token: string;
  token_type: 'Bearer';
  expires_in: number;
  refresh_token?: string;
  scope: string;
}

/**
 * Token Error Response (RFC 6749)
 */
export interface TokenErrorResponse {
  error: TokenError;
  error_description?: string;
  error_uri?: string;
}

/**
 * OAuth Authorization Server Metadata (RFC 8414)
 */
export interface AuthorizationServerMetadata {
  issuer: string;
  authorization_endpoint: string;
  token_endpoint: string;
  registration_endpoint?: string;
  jwks_uri: string;
  response_types_supported: ResponseType[];
  grant_types_supported: GrantType[];
  token_endpoint_auth_methods_supported: TokenEndpointAuthMethod[];
  code_challenge_methods_supported: CodeChallengeMethod[];
  scopes_supported?: string[];
  service_documentation?: string;
}

/**
 * JWKS (JSON Web Key Set)
 */
export interface JWKS {
  keys: JWK[];
}

/**
 * JSON Web Key
 */
export interface JWK {
  kty: 'RSA';
  use: 'sig';
  alg: 'RS256';
  kid: string;
  n: string; // RSA modulus
  e: string; // RSA exponent
}

// Enums as string literal types

export type GrantType = 'authorization_code' | 'refresh_token';
export type ResponseType = 'code';
export type TokenEndpointAuthMethod = 'client_secret_post' | 'client_secret_basic';
export type CodeChallengeMethod = 'S256';

export type TokenError =
  | 'invalid_request'
  | 'invalid_client'
  | 'invalid_grant'
  | 'unauthorized_client'
  | 'unsupported_grant_type'
  | 'invalid_scope'
  | 'server_error';

export type AuthorizationError =
  | 'invalid_request'
  | 'unauthorized_client'
  | 'access_denied'
  | 'unsupported_response_type'
  | 'invalid_scope'
  | 'server_error'
  | 'temporarily_unavailable';

/**
 * Authorization Request Parameters
 */
export interface AuthorizationRequest {
  response_type: ResponseType;
  client_id: string;
  redirect_uri: string;
  scope?: string;
  state: string;
  code_challenge: string;
  code_challenge_method: CodeChallengeMethod;
}

/**
 * Token Request (authorization_code grant)
 */
export interface AuthorizationCodeTokenRequest {
  grant_type: 'authorization_code';
  code: string;
  redirect_uri: string;
  client_id: string;
  client_secret?: string;
  code_verifier: string;
}

/**
 * Token Request (refresh_token grant)
 */
export interface RefreshTokenRequest {
  grant_type: 'refresh_token';
  refresh_token: string;
  client_id: string;
  client_secret?: string;
  scope?: string;
}

export type TokenRequest = AuthorizationCodeTokenRequest | RefreshTokenRequest;

/**
 * Pending Authorization (stored during Azure AD flow)
 */
export interface PendingAuthorization {
  clientId: string;
  redirectUri: string;
  scope: string;
  state: string;
  codeChallenge: string;
  codeChallengeMethod: CodeChallengeMethod;
  createdAt: number;
  expiresAt: number;
}
