import 'dotenv/config';
import { z } from 'zod';

const configSchema = z.object({
  // Azure AD / Entra ID
  azureClientId: z.string().uuid('AZURE_CLIENT_ID must be a valid UUID'),
  azureClientSecret: z.string().min(1, 'AZURE_CLIENT_SECRET is required'),
  azureTenantId: z.string().min(1, 'AZURE_TENANT_ID is required'),

  // Server
  port: z.coerce.number().int().min(1).max(65535).default(3000),
  baseUrl: z.string().url().default('http://localhost:3000'),
  nodeEnv: z.enum(['development', 'production', 'test']).default('development'),

  // Session
  sessionSecret: z.string().min(32, 'SESSION_SECRET must be at least 32 characters'),

  // Redis (optional)
  redisUrl: z.string().url().optional(),

  // Rate limiting
  rateLimitWindowMs: z.coerce.number().int().positive().default(60000),
  rateLimitMaxRequests: z.coerce.number().int().positive().default(100),

  // Timeouts
  graphApiTimeoutMs: z.coerce.number().int().positive().default(30000),

  // File parsing
  fileParseTimeoutMs: z.coerce.number().int().positive().default(30000),
  fileParseMaxOutputKb: z.coerce.number().int().positive().default(500),

  // Logging
  logLevel: z.enum(['trace', 'debug', 'info', 'warn', 'error', 'fatal']).default('info'),

  // OAuth 2.1 Authorization Server
  oauthSigningKeyPrivate: z.string().optional(), // PEM or file path, auto-generated if not set
  oauthSigningKeyPublic: z.string().optional(), // PEM or file path, auto-generated if not set
  oauthAccessTokenLifetimeSecs: z.coerce.number().int().positive().default(900), // 15 minutes
  oauthRefreshTokenLifetimeSecs: z.coerce.number().int().positive().default(86400), // 24 hours
  oauthAuthCodeLifetimeSecs: z.coerce.number().int().positive().default(600), // 10 minutes
  oauthAllowDynamicRegistration: z
    .string()
    .transform((val) => val === 'true' || val === '1')
    .pipe(z.boolean())
    .or(z.boolean())
    .default(true),

  // DCR security
  oauthAllowedRedirectPatterns: z.string().optional(), // Comma-separated URL patterns
})
  .superRefine((data, ctx) => {
    if (data.nodeEnv === 'production') {
      if (!data.redisUrl) {
        ctx.addIssue({
          code: z.ZodIssueCode.custom,
          path: ['redisUrl'],
          message: 'REDIS_URL is required in production (in-memory storage is not suitable)',
        });
      }
      if (!data.oauthSigningKeyPrivate) {
        ctx.addIssue({
          code: z.ZodIssueCode.custom,
          path: ['oauthSigningKeyPrivate'],
          message: 'OAUTH_SIGNING_KEY_PRIVATE is required in production (ephemeral keys cause token invalidation on restart)',
        });
      }
      if (!data.oauthSigningKeyPublic) {
        ctx.addIssue({
          code: z.ZodIssueCode.custom,
          path: ['oauthSigningKeyPublic'],
          message: 'OAUTH_SIGNING_KEY_PUBLIC is required in production',
        });
      }
      if (!data.baseUrl.startsWith('https://')) {
        ctx.addIssue({
          code: z.ZodIssueCode.custom,
          path: ['baseUrl'],
          message: 'MCP_SERVER_BASE_URL must use HTTPS in production',
        });
      }
    }
  });

export type Config = z.infer<typeof configSchema>;

function loadConfig(): Config {
  const rawConfig = {
    azureClientId: process.env['AZURE_CLIENT_ID'],
    azureClientSecret: process.env['AZURE_CLIENT_SECRET'],
    azureTenantId: process.env['AZURE_TENANT_ID'],
    port: process.env['PORT'] ?? process.env['MCP_SERVER_PORT'],
    baseUrl: process.env['MCP_SERVER_BASE_URL'],
    nodeEnv: process.env['NODE_ENV'],
    sessionSecret: process.env['SESSION_SECRET'],
    redisUrl: process.env['REDIS_URL'],
    rateLimitWindowMs: process.env['RATE_LIMIT_WINDOW_MS'],
    rateLimitMaxRequests: process.env['RATE_LIMIT_MAX_REQUESTS'],
    graphApiTimeoutMs: process.env['GRAPH_API_TIMEOUT_MS'],
    fileParseTimeoutMs: process.env['FILE_PARSE_TIMEOUT_MS'],
    fileParseMaxOutputKb: process.env['FILE_PARSE_MAX_OUTPUT_KB'],
    logLevel: process.env['LOG_LEVEL'],
    // OAuth 2.1 config
    oauthSigningKeyPrivate: process.env['OAUTH_SIGNING_KEY_PRIVATE'],
    oauthSigningKeyPublic: process.env['OAUTH_SIGNING_KEY_PUBLIC'],
    oauthAccessTokenLifetimeSecs: process.env['OAUTH_ACCESS_TOKEN_LIFETIME_SECS'],
    oauthRefreshTokenLifetimeSecs: process.env['OAUTH_REFRESH_TOKEN_LIFETIME_SECS'],
    oauthAuthCodeLifetimeSecs: process.env['OAUTH_AUTH_CODE_LIFETIME_SECS'],
    oauthAllowDynamicRegistration: process.env['OAUTH_ALLOW_DYNAMIC_REGISTRATION'],
    oauthAllowedRedirectPatterns: process.env['OAUTH_ALLOWED_REDIRECT_PATTERNS'],
  };

  const result = configSchema.safeParse(rawConfig);

  if (!result.success) {
    const errors = result.error.issues
      .map((issue) => `  - ${issue.path.join('.')}: ${issue.message}`)
      .join('\n');
    throw new Error(`Configuration validation failed:\n${errors}`);
  }

  return result.data;
}

export const config = loadConfig();

// OAuth scopes for Microsoft Graph
export const GRAPH_SCOPES = [
  'openid',
  'offline_access',
  'https://graph.microsoft.com/User.Read',
  'https://graph.microsoft.com/Mail.Read',
  'https://graph.microsoft.com/Mail.Read.Shared',
  'https://graph.microsoft.com/Files.Read.All',
  'https://graph.microsoft.com/Sites.Read.All',
];

// OAuth endpoints
export const getOAuthEndpoints = (tenantId: string) => ({
  authorize: `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/authorize`,
  token: `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`,
  logout: `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/logout`,
});
