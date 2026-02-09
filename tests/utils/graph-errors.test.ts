import { describe, it, expect } from 'vitest';
import { mapGraphError } from '../../src/utils/graph-errors.js';

describe('mapGraphError', () => {
  it('should map ErrorInvalidMailboxItemId to remediation with mailbox guidance', () => {
    const err = Object.assign(new Error("Item doesn't belong to the targeted mailbox"), {
      code: 'ErrorInvalidMailboxItemId',
      statusCode: 404,
    });
    const result = mapGraphError(err);

    expect(result.code).toBe('ErrorInvalidMailboxItemId');
    expect(result.statusCode).toBe(404);
    expect(result.remediation).toContain('mailbox_context');
    expect(result.remediation).toContain('mail_list_messages');
    expect(result.remediation).toContain('mail_get_message');
  });

  it('should map itemNotFound to remediation with sp_list_drives guidance', () => {
    const err = Object.assign(new Error('Item not found'), {
      code: 'itemNotFound',
      statusCode: 404,
    });
    const result = mapGraphError(err);

    expect(result.code).toBe('itemNotFound');
    expect(result.statusCode).toBe(404);
    expect(result.remediation).toContain('sp_list_drives');
    expect(result.remediation).toContain('sp_list_children');
  });

  it('should map ErrorAccessDenied to permission-related remediation', () => {
    const err = Object.assign(new Error('Access denied'), {
      code: 'ErrorAccessDenied',
      statusCode: 403,
    });
    const result = mapGraphError(err);

    expect(result.code).toBe('ErrorAccessDenied');
    expect(result.remediation).toContain('permission');
  });

  it('should map ErrorMailboxNotEnabledForRESTAPI', () => {
    const err = Object.assign(new Error('Mailbox not enabled'), {
      code: 'ErrorMailboxNotEnabledForRESTAPI',
      statusCode: 400,
    });
    const result = mapGraphError(err);

    expect(result.remediation).toContain('group');
  });

  it('should map 401 status code to re-authentication hint', () => {
    const err = Object.assign(new Error('Unauthorized'), {
      statusCode: 401,
    });
    const result = mapGraphError(err);

    expect(result.statusCode).toBe(401);
    expect(result.remediation).toContain('re-authenticate');
  });

  it('should map 429 status code to rate limit hint', () => {
    const err = Object.assign(new Error('Too many requests'), {
      statusCode: 429,
    });
    const result = mapGraphError(err);

    expect(result.statusCode).toBe(429);
    expect(result.remediation).toContain('Rate limited');
  });

  it('should map generic 404 to resource not found hint', () => {
    const err = Object.assign(new Error('Not found'), {
      statusCode: 404,
    });
    const result = mapGraphError(err);

    expect(result.statusCode).toBe(404);
    expect(result.remediation).toContain('Re-run the listing tool');
  });

  it('should handle unknown errors with fallback remediation', () => {
    const err = new Error('Something unexpected');
    const result = mapGraphError(err);

    expect(result.error).toBe('Something unexpected');
    expect(result.remediation).toContain('re-listing');
  });

  it('should handle non-Error objects', () => {
    const result = mapGraphError('string error');

    expect(result.error).toBe('string error');
    expect(result.remediation).toBeTruthy();
  });
});
