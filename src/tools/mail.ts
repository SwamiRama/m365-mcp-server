import { z } from 'zod';
import { GraphClient, type GraphMessage, type GraphMailFolder } from '../graph/client.js';
import { logger } from '../utils/logger.js';

/**
 * User context from the authenticated session.
 * Used to include user identity in responses, making them portable across MCP sessions.
 */
export interface UserContext {
  userEmail?: string;
  userId?: string;
}

// Shared mailbox schema — email address or Graph user ID
const mailboxSchema = z.string().min(1).max(256)
  .describe('Email address or user ID of a shared mailbox. Omit to use your personal mailbox.');

// Input schemas for mail tools
export const listMessagesInputSchema = z.object({
  folder: z
    .string()
    .optional()
    .describe(
      'Mail folder ID or well-known name (inbox, drafts, sentitems, deleteditems). Defaults to inbox.'
    ),
  top: z
    .number()
    .int()
    .min(1)
    .max(100)
    .optional()
    .default(25)
    .describe('Maximum number of messages to return (1-100)'),
  search: z
    .string()
    .max(500)
    .optional()
    .describe(
      'KQL search query (PREFERRED for sender/subject/body searches). Examples: "from:user@example.com", "subject:budget", "from:john subject:report". Supports from, to, cc, bcc, subject, body, attachment keywords.'
    ),
  query: z
    .string()
    .max(500)
    .optional()
    .describe(
      'OData $filter query (advanced use only, prefer "search" instead). Example: "hasAttachments eq true". Cannot be combined with "search".'
    ),
  since: z
    .string()
    .datetime()
    .optional()
    .describe('Filter messages received after this ISO 8601 datetime'),
  mailbox: mailboxSchema.optional(),
});

export const getMessageInputSchema = z.object({
  message_id: z.string().min(1).describe('The unique ID of the message'),
  include_body: z
    .boolean()
    .optional()
    .default(false)
    .describe('Whether to include the full message body (HTML/text)'),
  mailbox: mailboxSchema.optional(),
});

export const listFoldersInputSchema = z.object({
  mailbox: mailboxSchema.optional(),
});

export type ListMessagesInput = z.infer<typeof listMessagesInputSchema>;
export type GetMessageInput = z.infer<typeof getMessageInputSchema>;
export type ListFoldersInput = z.infer<typeof listFoldersInputSchema>;

// Well-known folder mappings
const WELL_KNOWN_FOLDERS: Record<string, string> = {
  inbox: 'inbox',
  drafts: 'drafts',
  sent: 'sentitems',
  sentitems: 'sentitems',
  deleted: 'deleteditems',
  deleteditems: 'deleteditems',
  junk: 'junkemail',
  junkemail: 'junkemail',
  archive: 'archive',
};

// Output formatters (to minimize PII exposure)
function formatMessage(message: GraphMessage, includeBody: boolean = false): object {
  const formatted: Record<string, unknown> = {
    id: message.id,
    subject: message.subject,
    preview: message.bodyPreview?.substring(0, 200), // Limit preview length
    from: message.from?.emailAddress?.address,
    fromName: message.from?.emailAddress?.name,
    to: message.toRecipients?.map((r) => r.emailAddress?.address).filter(Boolean),
    receivedAt: message.receivedDateTime,
    sentAt: message.sentDateTime,
    hasAttachments: message.hasAttachments,
    isRead: message.isRead,
    importance: message.importance,
    webLink: message.webLink,
  };

  if (includeBody && message.body) {
    formatted['body'] = {
      contentType: message.body.contentType,
      content: message.body.content,
    };
  }

  return formatted;
}

function formatFolder(folder: GraphMailFolder): object {
  return {
    id: folder.id,
    name: folder.displayName,
    unreadCount: folder.unreadItemCount,
    totalCount: folder.totalItemCount,
    childFolderCount: folder.childFolderCount,
  };
}

// Tool implementations
export class MailTools {
  private graphClient: GraphClient;
  private userContext?: UserContext;

  constructor(graphClient: GraphClient, userContext?: UserContext) {
    this.graphClient = graphClient;
    this.userContext = userContext;
  }

  /**
   * List messages in a mail folder
   */
  async listMessages(input: ListMessagesInput): Promise<object> {
    const validated = listMessagesInputSchema.parse(input);

    // Resolve well-known folder name
    let folderId = validated.folder;
    if (folderId) {
      const lowerFolder = folderId.toLowerCase();
      if (WELL_KNOWN_FOLDERS[lowerFolder]) {
        folderId = WELL_KNOWN_FOLDERS[lowerFolder];
      }
    }

    // Build search/filter — $search and $filter are mutually exclusive in Graph API
    let filter: string | undefined;
    let search: string | undefined;

    if (validated.search) {
      // KQL search mode — cannot combine with $filter or $orderby
      // Results are automatically sorted by receivedDateTime desc
      search = validated.search;
      if (validated.since) {
        // Append received date constraint to KQL search
        search = `${search} received>=${validated.since.split('T')[0]}`;
      }
    } else {
      // OData filter mode
      filter = validated.query;
      if (validated.since) {
        const sinceFilter = `receivedDateTime ge ${validated.since}`;
        filter = filter ? `(${filter}) and ${sinceFilter}` : sinceFilter;
      }
    }

    logger.debug({ folderId, top: validated.top, filter, search }, 'Listing messages');

    const result = await this.graphClient.listMessages({
      folderId,
      top: validated.top,
      filter,
      search,
      orderBy: search ? undefined : 'receivedDateTime desc',
      userId: validated.mailbox,
    });

    // Determine the effective mailbox identifier for this response.
    // When no explicit mailbox was specified, use the authenticated user's email
    // so the LLM can pass it as the 'mailbox' parameter in follow-up calls.
    // This makes message IDs portable across MCP sessions (where /me might resolve differently).
    const effectiveMailbox = validated.mailbox ?? this.userContext?.userEmail;

    return {
      messages: result.messages.map((m) => formatMessage(m)),
      count: result.messages.length,
      hasMore: !!result.nextLink,
      mailbox_context: effectiveMailbox ?? 'personal',
      _note: effectiveMailbox
        ? `These message IDs belong to mailbox '${effectiveMailbox}'. When calling mail_get_message, you MUST pass mailbox='${effectiveMailbox}'.`
        : "These message IDs belong to your personal mailbox. When calling mail_get_message, do NOT pass a 'mailbox' parameter.",
    };
  }

  /**
   * Get a specific message by ID
   */
  async getMessage(input: GetMessageInput): Promise<object> {
    const validated = getMessageInputSchema.parse(input);

    logger.debug(
      { messageId: validated.message_id, includeBody: validated.include_body },
      'Getting message'
    );

    try {
      const message = await this.graphClient.getMessage(
        validated.message_id,
        validated.include_body,
        validated.mailbox
      );

      return formatMessage(message, validated.include_body);
    } catch (err) {
      const code = (err as { code?: string }).code;
      if (code === 'ErrorInvalidMailboxItemId') {
        const hint = validated.mailbox
          ? `The message ID does not belong to mailbox '${validated.mailbox}'. This message was likely listed from your personal mailbox (no mailbox parameter). Retry: call mail_get_message with the same message_id but WITHOUT the mailbox parameter.`
          : `The message ID does not belong to your personal mailbox. It was likely listed from a shared mailbox. Retry: call mail_list_messages to get the correct mailbox_context, then call mail_get_message with that mailbox parameter.`;

        const enrichedError = new Error(hint) as Error & { code?: string; statusCode?: number };
        enrichedError.code = code;
        enrichedError.statusCode = (err as { statusCode?: number }).statusCode;
        throw enrichedError;
      }
      throw err;
    }
  }

  /**
   * List mail folders
   */
  async listFolders(input: ListFoldersInput): Promise<object> {
    const validated = listFoldersInputSchema.parse(input);
    logger.debug({ mailbox: validated.mailbox }, 'Listing mail folders');

    const folders = await this.graphClient.listMailFolders(validated.mailbox);

    return {
      folders: folders.map(formatFolder),
      count: folders.length,
    };
  }
}

// Tool definitions for MCP registration
export const mailToolDefinitions = [
  {
    name: 'mail_list_messages',
    description:
      'List email messages from a Microsoft 365 mailbox. Returns subject, sender, date, and preview. Use "search" (KQL) for sender/subject/body searches (PREFERRED). Use "query" (OData $filter) only for property filters like hasAttachments or isRead. Supports shared mailboxes via the mailbox parameter. Use mail_get_message with a message ID from THIS response to get full content.',
    inputSchema: {
      type: 'object' as const,
      properties: {
        folder: {
          type: 'string',
          description:
            'Mail folder ID or well-known name (inbox, drafts, sentitems, deleteditems). Defaults to inbox.',
        },
        top: {
          type: 'number',
          description: 'Maximum number of messages to return (1-100). Default: 25',
          minimum: 1,
          maximum: 100,
        },
        search: {
          type: 'string',
          description:
            'KQL search query (PREFERRED for sender/subject/body searches). Examples: "from:user@example.com", "subject:budget report", "from:john subject:invoice". Cannot be combined with "query".',
        },
        query: {
          type: 'string',
          description:
            'OData $filter query (advanced, prefer "search" instead). Example: "hasAttachments eq true". Cannot be combined with "search".',
        },
        since: {
          type: 'string',
          format: 'date-time',
          description: 'Filter messages received after this ISO 8601 datetime',
        },
        mailbox: {
          type: 'string',
          description: 'Email address or user ID of a shared mailbox. Omit to use your personal mailbox.',
        },
      },
    },
  },
  {
    name: 'mail_get_message',
    description:
      'Get the full details of a specific email message by ID, including the full body if requested. IMPORTANT: The message_id MUST be an ID returned by a recent mail_list_messages call — do not reuse IDs from previous conversations. You MUST pass the mailbox parameter with the exact mailbox_context value from the mail_list_messages response.',
    inputSchema: {
      type: 'object' as const,
      properties: {
        message_id: {
          type: 'string',
          description: 'The unique message ID from a recent mail_list_messages response. Do not fabricate or reuse stale IDs.',
        },
        include_body: {
          type: 'boolean',
          description: 'Whether to include the full message body (HTML/text). Default: false',
        },
        mailbox: {
          type: 'string',
          description: 'The mailbox_context value from the mail_list_messages response. MUST be provided to ensure the message ID resolves correctly.',
        },
      },
      required: ['message_id'],
    },
  },
  {
    name: 'mail_list_folders',
    description:
      'List all mail folders in a Microsoft 365 mailbox with unread/total message counts. Supports shared mailboxes via the mailbox parameter.',
    inputSchema: {
      type: 'object' as const,
      properties: {
        mailbox: {
          type: 'string',
          description: 'Email address or user ID of a shared mailbox. Omit to use your personal mailbox.',
        },
      },
    },
  },
];
