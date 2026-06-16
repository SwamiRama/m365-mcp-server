import { describe, it, expect, vi, beforeEach } from 'vitest';
import request from 'supertest';

// Mock the MCP auth middleware so DELETE /mcp resolves to a known OAuth session
// without minting a real Bearer JWT. index.ts imports both of these names.
type ReqWithSession = { session?: { id: string } };
vi.mock('../../src/oauth/middleware.js', () => ({
  bearerAuthMiddleware: (req: ReqWithSession, _res: unknown, next: () => void) => {
    req.session = { id: 'session-under-test' };
    next();
  },
  requireMcpAuth: (req: ReqWithSession, _res: unknown, next: () => void) => {
    req.session = { id: 'session-under-test' };
    next();
  },
}));

import { sessionManager } from '../../src/auth/session.js';
import { app } from '../../src/index.js';

describe('DELETE /mcp (MCP transport-session termination)', () => {
  beforeEach(() => {
    vi.restoreAllMocks();
  });

  it('acknowledges termination without deleting the OAuth user session', async () => {
    // Regression guard: terminating the MCP transport session must not revoke
    // the OAuth grant. Deleting the session here orphaned the client's still
    // valid refresh token and surfaced as "Failed to connect to MCP server".
    const deleteSession = vi.spyOn(sessionManager, 'deleteSession').mockResolvedValue();

    const res = await request(app).delete('/mcp');

    expect(res.status).toBe(200);
    expect(deleteSession).not.toHaveBeenCalled();
  });

  it('still deletes the session on explicit logout (/auth/logout)', async () => {
    // The identity teardown path must keep working.
    const deleteSession = vi.spyOn(sessionManager, 'deleteSession').mockResolvedValue();

    await request(app).get('/auth/logout').redirects(0);

    expect(deleteSession).toHaveBeenCalled();
  });
});
