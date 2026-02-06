/**
 * Structured audit logging for security-relevant events.
 * All events are logged at info level with a structured `event` field.
 */

import { logger } from './logger.js';

export type AuditEvent =
  | 'auth.login'
  | 'auth.logout'
  | 'auth.callback_success'
  | 'auth.callback_failed'
  | 'oauth.client_registered'
  | 'oauth.token_issued'
  | 'oauth.token_refreshed'
  | 'oauth.token_revoked'
  | 'tool.executed'
  | 'tool.file_accessed';

interface AuditContext {
  event: AuditEvent;
  userId?: string;
  userEmail?: string;
  clientId?: string;
  ip?: string;
  toolName?: string;
  fileName?: string;
  driveId?: string;
  [key: string]: unknown;
}

/**
 * Log a structured audit event.
 */
export function audit(context: AuditContext, message: string): void {
  logger.info(context, `[AUDIT] ${message}`);
}
