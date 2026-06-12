import { describe, it, expect } from 'vitest';
import { signJwt, verifyJwt, createAccessToken } from '../../src/oauth/jwt.js';
import { config } from '../../src/utils/config.js';
import type { AccessTokenPayload } from '../../src/oauth/types.js';

function basePayload(overrides: Partial<AccessTokenPayload> = {}): AccessTokenPayload {
  const now = Math.floor(Date.now() / 1000);
  return {
    iss: config.baseUrl,
    sub: 'session-1',
    aud: 'client-1',
    exp: now + 3600,
    iat: now,
    jti: 'jti-1',
    scope: 'mcp:tools',
    userId: 'user-1',
    ...overrides,
  };
}

describe('verifyJwt issuer/audience validation', () => {
  it('verifies a token issued by us (correct iss)', () => {
    const token = createAccessToken({
      sessionId: 'session-1',
      clientId: 'client-1',
      userId: 'user-1',
      scope: 'mcp:tools',
    });
    expect(verifyJwt(token)).not.toBeNull();
  });

  it('rejects a validly-signed token with a foreign issuer', () => {
    const token = signJwt(basePayload({ iss: 'https://attacker.example.com' }));
    expect(verifyJwt(token)).toBeNull();
  });

  it('rejects a validly-signed token with no issuer', () => {
    const payload = basePayload();
    // @ts-expect-error - deliberately omit iss to model a forged/legacy token
    delete payload.iss;
    const token = signJwt(payload);
    expect(verifyJwt(token)).toBeNull();
  });

  it('rejects a validly-signed token with no audience', () => {
    const payload = basePayload();
    // @ts-expect-error - deliberately omit aud
    delete payload.aud;
    const token = signJwt(payload);
    expect(verifyJwt(token)).toBeNull();
  });
});
