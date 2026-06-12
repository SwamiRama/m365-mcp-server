import { describe, it, expect } from 'vitest';
import { redactPII } from '../../src/utils/redact.js';

describe('redactPII session handling', () => {
  it('redacts sessionId (acts as a credential, must not be logged in clear)', () => {
    const out = redactPII({ sessionId: 'sess-abc-123' }) as Record<string, unknown>;
    expect(out['sessionId']).toBe('[REDACTED]');
  });

  it('redacts the mcp-session-id header value', () => {
    const out = redactPII({ 'mcp-session-id': 'sess-xyz-789' }) as Record<string, unknown>;
    expect(out['mcp-session-id']).toBe('[REDACTED]');
  });

  it('keeps correlationId so request tracing still works', () => {
    const out = redactPII({ correlationId: 'corr-1' }) as Record<string, unknown>;
    expect(out['correlationId']).toBe('corr-1');
  });

  it('keeps non-sensitive fields like userId', () => {
    const out = redactPII({ userId: 'user-1' }) as Record<string, unknown>;
    expect(out['userId']).toBe('user-1');
  });
});
