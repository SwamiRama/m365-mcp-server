import { z } from 'zod';
import { GraphClient, type GraphMessage, type GraphMailFolder } from '../graph/client.js';
import { logger } from '../utils/logger.js';

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
  query: z
    .string()
    .max(500)
    .optional()
    .describe(
      'OData filter query for messages (e.g., "from/emailAddress/address eq \'user@example.com\'")'
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
  constructor(private graphClient: GraphClient) {}

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

    // Build filter
    let filter = validated.query;
    if (validated.since) {
      const sinceFilter = `receivedDateTime ge ${validated.since}`;
      filter = filter ? `(${filter}) and ${sinceFilter}` : sinceFilter;
    }

    logger.debug({ folderId, top: validated.top, filter }, 'Listing messages');

    const result = await this.graphClient.listMessages({
      folderId,
      top: validated.top,
      filter,
      orderBy: 'receivedDateTime desc',
      userId: validated.mailbox,
    });

    return {
      messages: result.messages.map((m) => formatMessage(m)),
      count: result.messages.length,
      hasMore: !!result.nextLink,
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

    const message = await this.graphClient.getMessage(
      validated.message_id,
      validated.include_body,
      validated.mailbox
    );

    return formatMessage(message, validated.include_body);
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
      'List email messages from a Microsoft 365 mailbox. Returns subject, sender, date, and preview. Supports shared mailboxes via the mailbox parameter. Use mail_get_message for full content.',
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
        query: {
          type: 'string',
          description:
            "OData filter query for messages (e.g., \"from/emailAddress/address eq 'user@example.com'\")",
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
      'Get the full details of a specific email message by ID, including the full body if requested. Supports shared mailboxes via the mailbox parameter.',
    inputSchema: {
      type: 'object' as const,
      properties: {
        message_id: {
          type: 'string',
          description: 'The unique ID of the message',
        },
        include_body: {
          type: 'boolean',
          description: 'Whether to include the full message body (HTML/text). Default: false',
        },
        mailbox: {
          type: 'string',
          description: 'Email address or user ID of a shared mailbox. Omit to use your personal mailbox.',
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
