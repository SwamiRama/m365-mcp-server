/**
 * Centralized mapping from Microsoft Graph API error codes to LLM-friendly messages.
 * These messages include actionable remediation hints so LLM clients (e.g., Open WebUI)
 * can recover from errors without human intervention.
 */

export interface LlmErrorResponse {
  error: string;
  code: string | undefined;
  statusCode: number | undefined;
  remediation: string;
}

interface GraphLikeError {
  message?: string;
  code?: string;
  statusCode?: number;
}

function extractErrorFields(err: unknown): GraphLikeError {
  if (typeof err !== 'object' || err === null) {
    return { message: String(err) };
  }
  return {
    message: (err as { message?: string }).message ?? String(err),
    code: (err as { code?: string }).code,
    statusCode: (err as { statusCode?: number }).statusCode,
  };
}

const ERROR_MAP: Record<string, (fields: GraphLikeError) => string> = {
  ErrorInvalidMailboxItemId: () =>
    "The message ID does not belong to the specified mailbox. The ID was likely obtained from a different mailbox context. " +
    "Remediation: check the 'mailbox_context' value returned by mail_list_messages and pass the same value as the 'mailbox' parameter to mail_get_message. " +
    "If mailbox_context was 'personal', do NOT pass a mailbox parameter.",

  itemNotFound: () =>
    "The requested item was not found. The drive_id or item_id may be stale or from a previous session. " +
    "Remediation: call sp_list_drives to get a fresh drive_id, then call sp_list_children with that drive_id to get fresh item IDs.",

  ErrorItemNotFound: () =>
    "The requested item was not found. The resource ID may be invalid or the item may have been deleted. " +
    "Remediation: re-run the listing tool (e.g., sp_list_children or mail_list_messages) to get fresh IDs.",

  ErrorAccessDenied: () =>
    "Access denied. The current user does not have permission to access this resource. " +
    "The mailbox or site may not be shared with you, or additional admin consent may be required.",

  AccessDenied: () =>
    "Access denied. The current user does not have permission to access this resource.",

  ErrorMailboxNotEnabledForRESTAPI: () =>
    "This mailbox does not support the REST API. The mailbox parameter may point to a group or resource mailbox. " +
    "Use only user mailboxes or shared mailboxes.",

  AuthenticationError: () =>
    "Authentication has expired. Please re-authenticate to continue.",
};

export function mapGraphError(err: unknown): LlmErrorResponse {
  const fields = extractErrorFields(err);
  const code = fields.code;
  const statusCode = fields.statusCode;
  const message = fields.message ?? 'Unknown error';

  // Check for known error codes
  if (code && ERROR_MAP[code]) {
    return {
      error: message,
      code,
      statusCode,
      remediation: ERROR_MAP[code](fields),
    };
  }

  // Check by status code for unmapped codes
  if (statusCode === 401) {
    return {
      error: message,
      code,
      statusCode,
      remediation: 'Session expired or authentication invalid. Please re-authenticate.',
    };
  }

  if (statusCode === 429) {
    return {
      error: message,
      code,
      statusCode,
      remediation: 'Rate limited by Microsoft Graph API. Wait a moment and retry the same call.',
    };
  }

  if (statusCode === 404) {
    return {
      error: message,
      code,
      statusCode,
      remediation: 'Resource not found. The ID may be invalid or the resource may have been deleted. Re-run the listing tool to get fresh IDs.',
    };
  }

  // Default fallback
  return {
    error: message,
    code,
    statusCode,
    remediation: 'If this error persists, try re-listing the resource to get fresh IDs.',
  };
}
