import { describe, it, expect } from 'vitest';

// Mock redis to return null (uses MemoryCodeStore)
vi.mock('../../src/utils/redis.js', () => ({
  getRedisClient: () => null,
}));

import { vi } from 'vitest';
import { createAuthorizationCode, consumeAuthorizationCode } from '../../src/oauth/code-store.js';

const codeParams = (sessionId: string) => ({
  clientId: 'client-1',
  redirectUri: 'https://chat.catella.de/oauth/clients/mcp:m365/callback',
  scope: 'mcp:tools',
  codeChallenge: 'challenge',
  codeChallengeMethod: 'S256' as const,
  sessionId,
  userId: 'user-1',
});

describe('consumeAuthorizationCode single-use', () => {
  it('returns the code on first consume, null on second', async () => {
    const code = await createAuthorizationCode(codeParams('s-seq'));
    expect(await consumeAuthorizationCode(code)).not.toBeNull();
    expect(await consumeAuthorizationCode(code)).toBeNull();
  });

  it('lets exactly one of two concurrent consumes win (ToCToU)', async () => {
    const code = await createAuthorizationCode(codeParams('s-race'));
    const [a, b] = await Promise.all([
      consumeAuthorizationCode(code),
      consumeAuthorizationCode(code),
    ]);
    const winners = [a, b].filter(Boolean);
    expect(winners).toHaveLength(1);
  });
});
