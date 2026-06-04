import { describe, it, expect, vi } from 'vitest';
import type { Request, Response } from 'express';

vi.mock('../../src/utils/redis.js', () => ({
  getRedisClient: () => null,
}));

import { sendTokenError } from '../../src/oauth/routes.js';

function mockReq(): Request {
  return {
    log: { debug: vi.fn(), info: vi.fn(), warn: vi.fn(), error: vi.fn() },
  } as unknown as Request;
}

function mockRes(): Response {
  const res = {
    status: vi.fn().mockReturnThis(),
    json: vi.fn().mockReturnThis(),
  };
  return res as unknown as Response;
}

describe('sendTokenError', () => {
  it('responds with RFC 6749 error JSON', () => {
    const req = mockReq();
    const res = mockRes();

    sendTokenError(req, res, 'invalid_grant', 'Refresh token is invalid or expired');

    expect(res.status).toHaveBeenCalledWith(400);
    expect(res.json).toHaveBeenCalledWith({
      error: 'invalid_grant',
      error_description: 'Refresh token is invalid or expired',
    });
  });

  it('logs the failed grant for observability', () => {
    const req = mockReq();
    const res = mockRes();

    sendTokenError(req, res, 'invalid_grant', 'Refresh token is invalid or expired');

    expect(req.log.warn).toHaveBeenCalledWith(
      expect.objectContaining({
        event: 'oauth.token_error',
        error: 'invalid_grant',
        error_description: 'Refresh token is invalid or expired',
      }),
      expect.any(String)
    );
  });
});
