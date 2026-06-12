/**
 * PII / credential redaction for structured logging.
 *
 * Kept in its own module so it can be unit-tested without the logger's
 * transport side effects (and without the global logger mock in tests).
 */

// PII patterns to redact from string values
const PII_PATTERNS = [
  /Bearer\s+[A-Za-z0-9\-_]+\.[A-Za-z0-9\-_]+\.[A-Za-z0-9\-_]*/gi, // JWT tokens
  /access_token['":\s]+['"]?[A-Za-z0-9\-_]+\.[A-Za-z0-9\-_]+\.[A-Za-z0-9\-_]*/gi,
  /refresh_token['":\s]+['"]?[A-Za-z0-9\-_]+/gi,
  /client_secret['":\s]+['"]?[A-Za-z0-9\-_]+/gi,
  /password['":\s]+['"]?[^\s'"]+/gi,
  /[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}/g, // Email addresses
];

// Keys whose values are redacted entirely. `session` covers sessionId and the
// mcp-session-id header, which act as bearer credentials and must never be
// logged in clear (request tracing uses correlationId, which is not matched).
const SENSITIVE_KEY = /token|secret|password|authorization|cookie|session/i;

export function redactPII(obj: unknown): unknown {
  if (typeof obj === 'string') {
    let result = obj;
    for (const pattern of PII_PATTERNS) {
      result = result.replace(pattern, '[REDACTED]');
    }
    return result;
  }

  if (Array.isArray(obj)) {
    return obj.map(redactPII);
  }

  if (obj !== null && typeof obj === 'object') {
    const redacted: Record<string, unknown> = {};
    for (const [key, value] of Object.entries(obj)) {
      if (SENSITIVE_KEY.test(key)) {
        redacted[key] = '[REDACTED]';
      } else {
        redacted[key] = redactPII(value);
      }
    }
    return redacted;
  }

  return obj;
}
