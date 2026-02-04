#!/usr/bin/env node

import express, { Request, Response, NextFunction } from 'express';
import helmet from 'helmet';
import cors from 'cors';
import cookieParser from 'cookie-parser';
import rateLimit from 'express-rate-limit';
import { v4 as uuidv4 } from 'uuid';
import { config } from './utils/config.js';
import { logger, createRequestLogger } from './utils/logger.js';
import { oauthClient } from './auth/oauth.js';
import { sessionManager, type UserSession, type TokenSet } from './auth/session.js';
import { createGraphClient } from './graph/client.js';
import { ToolExecutor, allToolDefinitions } from './tools/index.js';
import { oauthRouter } from './oauth/routes.js';
import { bearerAuthMiddleware } from './oauth/middleware.js';

const app = express();

// =============================================================================
// Middleware
// =============================================================================

// Security headers
app.use(
  helmet({
    contentSecurityPolicy: {
      directives: {
        defaultSrc: ["'self'"],
        scriptSrc: ["'self'"],
        styleSrc: ["'self'", "'unsafe-inline'"],
        imgSrc: ["'self'", 'data:', 'https:'],
        connectSrc: ["'self'", 'https://graph.microsoft.com', 'https://login.microsoftonline.com'],
      },
    },
    hsts: config.nodeEnv === 'production',
  })
);

// CORS configuration
app.use(
  cors({
    origin: config.nodeEnv === 'production' ? config.baseUrl : true,
    credentials: true,
    methods: ['GET', 'POST', 'DELETE', 'OPTIONS'],
    allowedHeaders: ['Content-Type', 'Authorization', 'MCP-Session-Id', 'MCP-Protocol-Version', 'Accept'],
  })
);

// Rate limiting
const limiter = rateLimit({
  windowMs: config.rateLimitWindowMs,
  max: config.rateLimitMaxRequests,
  standardHeaders: true,
  legacyHeaders: false,
  message: { error: 'Too many requests, please try again later' },
});
app.use(limiter);

// Body parsing
app.use(express.json({ limit: '1mb' }));
app.use(cookieParser());

// Correlation ID middleware
app.use((req: Request, res: Response, next: NextFunction) => {
  const correlationId = (req.headers['x-correlation-id'] as string) || uuidv4();
  req.correlationId = correlationId;
  res.setHeader('X-Correlation-Id', correlationId);
  req.log = createRequestLogger(correlationId);
  next();
});

// Request logging
app.use((req: Request, res: Response, next: NextFunction) => {
  req.log.info({ method: req.method, path: req.path }, 'Request received');
  next();
});

// =============================================================================
// Type augmentation
// =============================================================================

declare global {
  namespace Express {
    interface Request {
      correlationId: string;
      log: typeof logger;
      session?: UserSession;
    }
  }
}

// =============================================================================
// Session middleware
// =============================================================================

async function sessionMiddleware(
  req: Request,
  res: Response,
  next: NextFunction
): Promise<void> {
  const sessionId =
    (req.headers['mcp-session-id'] as string | undefined) ??
    (req.cookies?.['mcp-session'] as string | undefined);

  if (sessionId) {
    const session = await sessionManager.getSession(sessionId);
    if (session) {
      req.session = session;
    }
  }

  next();
}

app.use(sessionMiddleware);

// Bearer token authentication (OAuth 2.1)
// This middleware validates Bearer tokens and loads the session
// It runs AFTER session middleware so it can override cookie-based sessions
app.use(bearerAuthMiddleware);

// =============================================================================
// OAuth 2.1 Authorization Server
// =============================================================================

// Mount OAuth routes (/.well-known/*, /authorize, /token, /register, /oauth/callback)
app.use(oauthRouter);

// =============================================================================
// Health check
// =============================================================================

app.get('/health', (_req: Request, res: Response) => {
  res.json({ status: 'healthy', version: '1.0.0' });
});

// =============================================================================
// OAuth routes
// =============================================================================

app.get('/auth/login', async (req: Request, res: Response) => {
  try {
    const redirectUri = `${config.baseUrl}/auth/callback`;
    const result = await oauthClient.getAuthorizationUrl(redirectUri, req.session?.id);

    // Set session cookie
    res.cookie('mcp-session', result.session.id, {
      httpOnly: true,
      secure: config.nodeEnv === 'production',
      sameSite: 'lax',
      maxAge: 24 * 60 * 60 * 1000, // 24 hours
    });

    res.redirect(result.url);
  } catch (err) {
    req.log.error({ err }, 'Failed to initiate login');
    res.status(500).json({ error: 'Failed to initiate login' });
  }
});

app.get('/auth/callback', async (req: Request, res: Response): Promise<void> => {
  try {
    const { code, state, error, error_description } = req.query;

    if (error) {
      req.log.warn({ error, error_description }, 'OAuth error');
      res.status(400).json({
        error: error as string,
        description: error_description as string,
      });
      return;
    }

    if (!code || typeof code !== 'string') {
      res.status(400).json({ error: 'Missing authorization code' });
      return;
    }

    if (!req.session) {
      res.status(400).json({ error: 'Session not found' });
      return;
    }

    // Validate state
    if (!state || !oauthClient.validateState(req.session, state as string)) {
      res.status(400).json({ error: 'Invalid state parameter' });
      return;
    }

    // Exchange code for tokens
    const redirectUri = `${config.baseUrl}/auth/callback`;
    const tokenResult = await oauthClient.exchangeCodeForTokens(
      code,
      redirectUri,
      req.session
    );

    // Update session with tokens and user info
    req.session.tokens = tokenResult.tokens;
    req.session.userId = tokenResult.userId;
    req.session.userEmail = tokenResult.userEmail;
    req.session.userDisplayName = tokenResult.userDisplayName;

    // Clear PKCE data
    req.session.pkceVerifier = undefined;
    req.session.state = undefined;
    req.session.nonce = undefined;

    await sessionManager.saveSession(req.session);

    req.log.info({ userId: tokenResult.userId }, 'Login successful');

    // Return success page or redirect
    res.send(`
      <!DOCTYPE html>
      <html>
      <head>
        <title>Login Successful</title>
        <style>
          body { font-family: system-ui, sans-serif; max-width: 600px; margin: 50px auto; padding: 20px; }
          .success { color: #22c55e; }
        </style>
      </head>
      <body>
        <h1 class="success">Login Successful!</h1>
        <p>You are now authenticated as: <strong>${tokenResult.userDisplayName ?? tokenResult.userEmail ?? 'User'}</strong></p>
        <p>You can close this window and return to your application.</p>
        <p><small>Session ID: ${req.session.id}</small></p>
      </body>
      </html>
    `);
  } catch (err) {
    req.log.error({ err }, 'OAuth callback failed');
    res.status(500).json({ error: 'Authentication failed' });
  }
});

app.get('/auth/logout', async (req: Request, res: Response) => {
  try {
    if (req.session) {
      await sessionManager.deleteSession(req.session.id);
    }

    res.clearCookie('mcp-session');

    const postLogoutRedirect = `${config.baseUrl}/`;
    const logoutUrl = oauthClient.getLogoutUrl(postLogoutRedirect);

    res.redirect(logoutUrl);
  } catch (err) {
    req.log.error({ err }, 'Logout failed');
    res.status(500).json({ error: 'Logout failed' });
  }
});

app.get('/auth/status', (req: Request, res: Response): void => {
  if (!req.session || !req.session.tokens) {
    res.json({
      authenticated: false,
      loginUrl: `${config.baseUrl}/auth/login`,
    });
    return;
  }

  res.json({
    authenticated: true,
    userId: req.session.userId,
    userEmail: req.session.userEmail,
    userDisplayName: req.session.userDisplayName,
    sessionId: req.session.id,
  });
});

// =============================================================================
// MCP Protocol Handler (Streamable HTTP)
// =============================================================================

// MCP endpoint
app.post('/mcp', async (req: Request, res: Response): Promise<void> => {
  const protocolVersion = req.headers['mcp-protocol-version'] as string;

  // Validate protocol version (allow missing for backwards compatibility)
  if (protocolVersion && !['2025-11-25', '2025-03-26', '2024-11-05'].includes(protocolVersion)) {
    res.status(400).json({
      jsonrpc: '2.0',
      error: { code: -32600, message: 'Unsupported protocol version' },
      id: null,
    });
    return;
  }

  // Handle JSON-RPC request
  const jsonRpcRequest = req.body;

  if (!jsonRpcRequest || !jsonRpcRequest.method) {
    res.status(400).json({
      jsonrpc: '2.0',
      error: { code: -32600, message: 'Invalid request' },
      id: jsonRpcRequest?.id ?? null,
    });
    return;
  }

  try {
    const result = await handleMcpRequest(req, jsonRpcRequest);

    if (result === null) {
      // Notification - no response needed
      res.status(202).send();
      return;
    }

    res.json({
      jsonrpc: '2.0',
      result,
      id: jsonRpcRequest.id,
    });
  } catch (err) {
    const error = err instanceof Error ? err : new Error('Unknown error');
    req.log.error({ err, method: jsonRpcRequest.method }, 'MCP request failed');

    res.status(500).json({
      jsonrpc: '2.0',
      error: {
        code: -32603,
        message: error.message,
      },
      id: jsonRpcRequest.id,
    });
  }
});

// MCP GET endpoint for SSE streams
app.get('/mcp', (req: Request, res: Response): void => {
  const accept = req.headers['accept'] as string ?? '';

  if (!accept.includes('text/event-stream')) {
    res.status(405).json({ error: 'Method not allowed' });
    return;
  }

  // Set SSE headers
  res.setHeader('Content-Type', 'text/event-stream');
  res.setHeader('Cache-Control', 'no-cache');
  res.setHeader('Connection', 'keep-alive');

  // Send initial event with event ID for reconnection
  res.write(`id: ${Date.now()}\ndata: \n\n`);

  // Keep connection alive
  const keepAlive = setInterval(() => {
    res.write(': keepalive\n\n');
  }, 30000);

  req.on('close', () => {
    clearInterval(keepAlive);
  });
});

// MCP session termination
app.delete('/mcp', async (req: Request, res: Response) => {
  if (req.session) {
    await sessionManager.deleteSession(req.session.id);
    res.clearCookie('mcp-session');
  }
  res.status(200).send();
});

// =============================================================================
// MCP Request Handler
// =============================================================================

interface JsonRpcRequest {
  jsonrpc: '2.0';
  method: string;
  params?: Record<string, unknown>;
  id?: string | number | null;
}

async function handleMcpRequest(
  req: Request,
  jsonRpc: JsonRpcRequest
): Promise<unknown> {
  const { method, params } = jsonRpc;

  switch (method) {
    case 'initialize':
      return handleInitialize(req, params);

    case 'notifications/initialized':
      // Client notification - no response
      req.log.info('Client initialized');
      return null;

    case 'tools/list':
      return handleToolsList(req);

    case 'tools/call':
      return handleToolsCall(req, params);

    case 'resources/list':
      return { resources: [] }; // No resources exposed

    case 'prompts/list':
      return { prompts: [] }; // No prompts defined

    case 'ping':
      return {};

    default:
      throw new Error(`Unknown method: ${method}`);
  }
}

async function handleInitialize(
  req: Request,
  params?: Record<string, unknown>
): Promise<object> {
  const clientInfo = params?.['clientInfo'] as { name?: string; version?: string } | undefined;

  req.log.info({ clientInfo }, 'MCP initialization requested');

  // Create or get session
  let session = req.session;
  if (!session) {
    session = await sessionManager.createSession();
    req.session = session;
  }

  return {
    protocolVersion: '2025-11-25',
    capabilities: {
      tools: { listChanged: false },
      resources: { subscribe: false, listChanged: false },
      prompts: { listChanged: false },
    },
    serverInfo: {
      name: 'm365-mcp-server',
      version: '1.0.0',
    },
    // Return session ID in header (set by response)
    _sessionId: session.id,
  };
}

async function handleToolsList(req: Request): Promise<object> {
  // Check authentication
  if (!req.session?.tokens) {
    return {
      tools: allToolDefinitions.map((tool) => ({
        ...tool,
        description: `${tool.description} [Requires authentication - visit ${config.baseUrl}/auth/login]`,
      })),
    };
  }

  return { tools: allToolDefinitions };
}

async function handleToolsCall(
  req: Request,
  params?: Record<string, unknown>
): Promise<object> {
  const toolName = params?.['name'] as string;
  const toolArgs = (params?.['arguments'] as Record<string, unknown>) ?? {};

  if (!toolName) {
    throw new Error('Tool name is required');
  }

  // Check authentication
  if (!req.session?.tokens) {
    throw new Error(
      `Authentication required. Please visit ${config.baseUrl}/auth/login to sign in with Microsoft 365.`
    );
  }

  // Check if tokens need refresh
  const tokens = sessionManager.getDecryptedTokens(req.session);
  if (!tokens) {
    throw new Error('Session tokens invalid - please re-authenticate');
  }

  let currentTokens: TokenSet = tokens;

  if (sessionManager.tokensNeedRefresh(tokens)) {
    req.log.info('Refreshing tokens');
    try {
      currentTokens = await oauthClient.refreshTokens(req.session);
      await sessionManager.updateTokens(req.session.id, currentTokens);
    } catch (err) {
      req.log.warn({ err }, 'Token refresh failed - re-authentication required');
      throw new Error(
        `Session expired. Please visit ${config.baseUrl}/auth/login to sign in again.`
      );
    }
  }

  // Create Graph client and tool executor
  const graphClient = createGraphClient(currentTokens);
  const toolExecutor = new ToolExecutor(graphClient);

  req.log.info({ toolName }, 'Executing tool');

  const result = await toolExecutor.execute(toolName, toolArgs);

  return {
    content: [
      {
        type: 'text',
        text: JSON.stringify(result, null, 2),
      },
    ],
    isError: false,
  };
}

// =============================================================================
// Error handling
// =============================================================================

app.use((err: Error, req: Request, res: Response, _next: NextFunction) => {
  req.log.error({ err }, 'Unhandled error');

  res.status(500).json({
    error: 'Internal server error',
    correlationId: req.correlationId,
  });
});

// =============================================================================
// Start server
// =============================================================================

const server = app.listen(config.port, () => {
  logger.info(
    {
      port: config.port,
      baseUrl: config.baseUrl,
      nodeEnv: config.nodeEnv,
    },
    `m365-mcp-server started`
  );
  logger.info(`MCP endpoint: ${config.baseUrl}/mcp`);
  logger.info(`Login URL: ${config.baseUrl}/auth/login`);
  logger.info(`OAuth 2.1 metadata: ${config.baseUrl}/.well-known/oauth-authorization-server`);
  logger.info(`Dynamic registration: ${config.oauthAllowDynamicRegistration ? 'enabled' : 'disabled'}`);
});

// Graceful shutdown
process.on('SIGTERM', () => {
  logger.info('SIGTERM received, shutting down gracefully');
  server.close(() => {
    logger.info('Server closed');
    process.exit(0);
  });
});

process.on('SIGINT', () => {
  logger.info('SIGINT received, shutting down gracefully');
  server.close(() => {
    logger.info('Server closed');
    process.exit(0);
  });
});

export { app };
