import { describe, it, expect, vi, beforeEach } from 'vitest';
import type { Request, Response, NextFunction } from 'express';
import { bearerAuthMiddleware, requireBearerAuth, requireAuth } from '../../src/oauth/middleware.js';

// Mock dependencies
vi.mock('../../src/utils/redis.js', () => ({
  getRedisClient: () => null,
}));

vi.mock('../../src/oauth/jwt.js', () => ({
  verifyJwt: vi.fn(),
  initializeKeys: vi.fn(),
}));

vi.mock('../../src/auth/session.js', () => ({
  sessionManager: {
    getSession: vi.fn(),
    createSession: vi.fn(),
    saveSession: vi.fn(),
    deleteSession: vi.fn(),
  },
}));

import { verifyJwt } from '../../src/oauth/jwt.js';
import { sessionManager } from '../../src/auth/session.js';

const WWW_AUTH_PATTERN = /^Bearer resource_metadata=".*\/\.well-known\/oauth-authorization-server"$/;

function mockReq(overrides: Partial<Request> = {}): Request {
  return {
    headers: {},
    log: { debug: vi.fn(), info: vi.fn(), warn: vi.fn(), error: vi.fn() },
    ...overrides,
  } as unknown as Request;
}

function mockRes(): Response {
  const headers: Record<string, string> = {};
  const res = {
    status: vi.fn().mockReturnThis(),
    json: vi.fn().mockReturnThis(),
    setHeader: vi.fn((key: string, val: string) => { headers[key] = val; }),
    getHeader: (key: string) => headers[key],
    _headers: headers,
  };
  return res as unknown as Response;
}

describe('OAuth Middleware WWW-Authenticate header', () => {
  beforeEach(() => {
    vi.clearAllMocks();
  });

  describe('bearerAuthMiddleware', () => {
    it('should set WWW-Authenticate on invalid JWT (401)', async () => {
      vi.mocked(verifyJwt).mockReturnValue(null);

      const req = mockReq({ headers: { authorization: 'Bearer invalid-token' } });
      const res = mockRes();
      const next = vi.fn();

      await bearerAuthMiddleware(req, res, next);

      expect(res.status).toHaveBeenCalledWith(401);
      expect(res.setHeader).toHaveBeenCalledWith(
        'WWW-Authenticate',
        expect.stringMatching(WWW_AUTH_PATTERN)
      );
      expect(next).not.toHaveBeenCalled();
    });

    it('should set WWW-Authenticate when session not found (401)', async () => {
      vi.mocked(verifyJwt).mockReturnValue({
        iss: 'http://localhost:3000',
        sub: 'session-123',
        aud: 'client-1',
        exp: Math.floor(Date.now() / 1000) + 3600,
        iat: Math.floor(Date.now() / 1000),
        jti: 'jti-1',
        scope: 'mcp:tools',
        userId: 'user-1',
      });
      vi.mocked(sessionManager.getSession).mockResolvedValue(null);

      const req = mockReq({ headers: { authorization: 'Bearer valid-token' } });
      const res = mockRes();
      const next = vi.fn();

      await bearerAuthMiddleware(req, res, next);

      expect(res.status).toHaveBeenCalledWith(401);
      expect(res.setHeader).toHaveBeenCalledWith(
        'WWW-Authenticate',
        expect.stringMatching(WWW_AUTH_PATTERN)
      );
    });

    it('should set WWW-Authenticate when session has no tokens (401)', async () => {
      vi.mocked(verifyJwt).mockReturnValue({
        iss: 'http://localhost:3000',
        sub: 'session-123',
        aud: 'client-1',
        exp: Math.floor(Date.now() / 1000) + 3600,
        iat: Math.floor(Date.now() / 1000),
        jti: 'jti-1',
        scope: 'mcp:tools',
        userId: 'user-1',
      });
      vi.mocked(sessionManager.getSession).mockResolvedValue({
        id: 'session-123',
        tokens: undefined,
        createdAt: Date.now(),
        lastAccessedAt: Date.now(),
      });

      const req = mockReq({ headers: { authorization: 'Bearer valid-token' } });
      const res = mockRes();
      const next = vi.fn();

      await bearerAuthMiddleware(req, res, next);

      expect(res.status).toHaveBeenCalledWith(401);
      expect(res.setHeader).toHaveBeenCalledWith(
        'WWW-Authenticate',
        expect.stringMatching(WWW_AUTH_PATTERN)
      );
    });

    it('should NOT set WWW-Authenticate when no Bearer token (pass-through)', async () => {
      const req = mockReq({ headers: {} });
      const res = mockRes();
      const next = vi.fn();

      await bearerAuthMiddleware(req, res, next);

      expect(res.setHeader).not.toHaveBeenCalled();
      expect(next).toHaveBeenCalled();
    });

    it('should NOT set WWW-Authenticate on successful auth', async () => {
      vi.mocked(verifyJwt).mockReturnValue({
        iss: 'http://localhost:3000',
        sub: 'session-123',
        aud: 'client-1',
        exp: Math.floor(Date.now() / 1000) + 3600,
        iat: Math.floor(Date.now() / 1000),
        jti: 'jti-1',
        scope: 'mcp:tools',
        userId: 'user-1',
      });
      vi.mocked(sessionManager.getSession).mockResolvedValue({
        id: 'session-123',
        tokens: {
          accessToken: 'enc-token',
          scope: 'mcp:tools',
          expiresAt: Date.now() + 3600000,
        },
        createdAt: Date.now(),
        lastAccessedAt: Date.now(),
      });

      const req = mockReq({ headers: { authorization: 'Bearer valid-token' } });
      const res = mockRes();
      const next = vi.fn();

      await bearerAuthMiddleware(req, res, next);

      expect(res.status).not.toHaveBeenCalled();
      expect(next).toHaveBeenCalled();
    });
  });

  describe('requireBearerAuth', () => {
    it('should set WWW-Authenticate when no Bearer token (401)', async () => {
      const req = mockReq({ headers: {} });
      const res = mockRes();
      const next: NextFunction = vi.fn();

      await requireBearerAuth(req, res, next);

      expect(res.status).toHaveBeenCalledWith(401);
      expect(res.setHeader).toHaveBeenCalledWith(
        'WWW-Authenticate',
        expect.stringMatching(WWW_AUTH_PATTERN)
      );
    });
  });

  describe('requireAuth', () => {
    it('should set WWW-Authenticate when no session (401)', () => {
      const req = mockReq();
      const res = mockRes();
      const next: NextFunction = vi.fn();

      requireAuth(req, res, next);

      expect(res.status).toHaveBeenCalledWith(401);
      expect(res.setHeader).toHaveBeenCalledWith(
        'WWW-Authenticate',
        expect.stringMatching(WWW_AUTH_PATTERN)
      );
    });

    it('should set WWW-Authenticate when session has no tokens (401)', () => {
      const req = mockReq();
      (req as Record<string, unknown>).session = { id: 'sess-1', tokens: undefined };
      const res = mockRes();
      const next: NextFunction = vi.fn();

      requireAuth(req, res, next);

      expect(res.status).toHaveBeenCalledWith(401);
      expect(res.setHeader).toHaveBeenCalledWith(
        'WWW-Authenticate',
        expect.stringMatching(WWW_AUTH_PATTERN)
      );
    });

    it('should call next when authenticated', () => {
      const req = mockReq();
      (req as Record<string, unknown>).session = {
        id: 'sess-1',
        tokens: { accessToken: 'tok', scope: 'mcp:tools', expiresAt: Date.now() + 3600000 },
      };
      const res = mockRes();
      const next: NextFunction = vi.fn();

      requireAuth(req, res, next);

      expect(next).toHaveBeenCalled();
      expect(res.status).not.toHaveBeenCalled();
    });
  });
});
