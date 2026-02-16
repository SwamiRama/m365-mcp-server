import { z } from 'zod';
import { GraphClient, type GraphMessage, type GraphMailFolder } from '../graph/client.js';
import { logger } from '../utils/logger.js';
import { stripHtmlTags, isParsableMimeType, parseFileContent } from '../utils/file-parser.js';
import { isTextMimeType, formatFileSize } from '../utils/content-fetcher.js';

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
    .default(true)
    .describe('Include the full message body (default: true). HTML is auto-converted to plain text. Set to false for metadata only.'),
  mailbox: mailboxSchema.optional(),
});

export const listFoldersInputSchema = z.object({
  parent_folder_id: z.string().min(1).optional()
    .describe('Folder ID or well-known name (inbox, sent, drafts, deleted, junk, archive) to list subfolders of. Omit to list top-level folders.'),
  mailbox: mailboxSchema.optional(),
});

export const getAttachmentInputSchema = z.object({
  message_id: z.string().min(1).describe('The message ID containing the attachment'),
  attachment_id: z.string().min(1).describe('The attachment ID from mail_get_message response'),
  mailbox: mailboxSchema.optional(),
});

export type ListMessagesInput = z.infer<typeof listMessagesInputSchema>;
export type GetMessageInput = z.infer<typeof getMessageInputSchema>;
export type ListFoldersInput = z.infer<typeof listFoldersInputSchema>;
export type GetAttachmentInput = z.infer<typeof getAttachmentInputSchema>;

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
    preview: message.bodyPreview?.substring(0, 200),
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

  // CC/BCC — only in getMessage, omitted when empty
  const cc = message.ccRecipients?.map((r) => r.emailAddress?.address).filter(Boolean);
  if (cc && cc.length > 0) formatted['cc'] = cc;
  const bcc = message.bccRecipients?.map((r) => r.emailAddress?.address).filter(Boolean);
  if (bcc && bcc.length > 0) formatted['bcc'] = bcc;

  if (includeBody && message.body) {
    const content = message.body.content;
    if (message.body.contentType?.toLowerCase() === 'html') {
      formatted['body'] = { text: stripHtmlTags(content), html: content };
    } else {
      formatted['body'] = { text: content };
    }
  }

  // Attachment metadata from $expand
  if (message.attachments && message.attachments.length > 0) {
    const regular = message.attachments.filter((a) => !a.isInline);
    const inlineCount = message.attachments.length - regular.length;
    if (regular.length > 0) {
      formatted['attachments'] = regular.map((a) => ({
        id: a.id,
        name: a.name,
        contentType: a.contentType,
        size: a.size,
      }));
    }
    if (inlineCount > 0) {
      formatted['inlineAttachmentCount'] = inlineCount;
    }
  }

  return formatted;
}

function formatFolder(folder: GraphMailFolder): object {
  return {
    id: folder.id,
    name: folder.displayName,
    parentFolderId: folder.parentFolderId,
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
   * List mail folders (top-level or children of a specific folder)
   */
  async listFolders(input: ListFoldersInput): Promise<object> {
    const validated = listFoldersInputSchema.parse(input);
    logger.debug({ mailbox: validated.mailbox, parentFolderId: validated.parent_folder_id }, 'Listing mail folders');

    let folders: GraphMailFolder[];
    if (validated.parent_folder_id) {
      // Resolve well-known folder names
      const lower = validated.parent_folder_id.toLowerCase();
      const resolvedId = WELL_KNOWN_FOLDERS[lower] ?? validated.parent_folder_id;
      folders = await this.graphClient.listChildFolders(resolvedId, validated.mailbox);
    } else {
      folders = await this.graphClient.listMailFolders(validated.mailbox);
    }

    return {
      folders: folders.map(formatFolder),
      count: folders.length,
    };
  }
  /**
   * Get and parse an email attachment
   */
  async getAttachment(input: GetAttachmentInput): Promise<object> {
    const validated = getAttachmentInputSchema.parse(input);

    logger.debug(
      { messageId: validated.message_id, attachmentId: validated.attachment_id },
      'Getting attachment'
    );

    try {
      const attachment = await this.graphClient.getAttachment(
        validated.message_id,
        validated.attachment_id,
        validated.mailbox
      );

      const odataType = attachment['@odata.type'] ?? '';
      const result: Record<string, unknown> = {
        id: attachment.id,
        name: attachment.name,
        contentType: attachment.contentType,
        size: attachment.size,
      };

      // Item attachment (embedded email/event)
      if (odataType.includes('itemAttachment')) {
        result['type'] = 'itemAttachment';
        result['item'] = attachment.item;
        result['content'] = null;
        result['_note'] = 'Item attachments (embedded emails/events) cannot be read as files. The item metadata is shown above.';
        return result;
      }

      // Reference attachment (link to OneDrive/SharePoint file)
      if (odataType.includes('referenceAttachment')) {
        result['type'] = 'referenceAttachment';
        result['sourceUrl'] = attachment.sourceUrl;
        result['content'] = null;
        result['_note'] = 'This is a link to a file in OneDrive/SharePoint. Use sp_get_file or od_get_file to read the actual content.';
        return result;
      }

      // File attachment
      result['type'] = 'fileAttachment';

      if (!attachment.contentBytes) {
        throw new Error('Attachment content is empty — the file may have been deleted or is inaccessible.');
      }

      const buffer = Buffer.from(attachment.contentBytes, 'base64');
      const MAX_SIZE = 10 * 1024 * 1024; // 10MB
      if (buffer.length > MAX_SIZE) {
        throw new Error(`Attachment size (${formatFileSize(buffer.length)}) exceeds the 10 MB limit.`);
      }

      const mimeType = attachment.contentType ?? 'application/octet-stream';

      if (isTextMimeType(mimeType)) {
        result['content'] = buffer.toString('utf-8');
        result['contentType'] = 'text';
      } else if (isParsableMimeType(mimeType)) {
        const parsed = await parseFileContent(buffer, mimeType, attachment.name ?? 'attachment');
        result['content'] = parsed.text;
        result['contentType'] = 'parsed_text';
        result['parsedFormat'] = parsed.format;
        result['truncated'] = parsed.truncated;
      } else {
        result['content'] = null;
        result['contentType'] = 'binary';
        result['_note'] = `Binary file (${mimeType}) cannot be displayed as text. The metadata is shown above.`;
      }

      return result;
    } catch (err) {
      const code = (err as { code?: string }).code;
      if (code === 'ErrorInvalidMailboxItemId') {
        const hint = validated.mailbox
          ? `The message/attachment ID does not belong to mailbox '${validated.mailbox}'. Retry without the mailbox parameter or with the correct mailbox.`
          : `The message/attachment ID does not belong to your personal mailbox. It was likely from a shared mailbox — retry with the correct mailbox parameter.`;

        const enrichedError = new Error(hint) as Error & { code?: string; statusCode?: number };
        enrichedError.code = code;
        enrichedError.statusCode = (err as { statusCode?: number }).statusCode;
        throw enrichedError;
      }
      throw err;
    }
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
      'Get full email details including body (HTML auto-converted to plain text), CC/BCC recipients, and attachment metadata. Body is included by default. IMPORTANT: The message_id MUST be from a recent mail_list_messages call. You MUST pass the mailbox parameter with the exact mailbox_context value from the mail_list_messages response. If attachments are listed, use mail_get_attachment to read their content.',
    inputSchema: {
      type: 'object' as const,
      properties: {
        message_id: {
          type: 'string',
          description: 'The unique message ID from a recent mail_list_messages response. Do not fabricate or reuse stale IDs.',
        },
        include_body: {
          type: 'boolean',
          description: 'Include the full message body (default: true). HTML is auto-converted to plain text.',
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
      'List mail folders in a Microsoft 365 mailbox with unread/total message counts. Supports browsing subfolders via parent_folder_id. Supports shared mailboxes via the mailbox parameter.',
    inputSchema: {
      type: 'object' as const,
      properties: {
        parent_folder_id: {
          type: 'string',
          description: 'Folder ID or well-known name (inbox, sent, drafts, deleted, junk, archive) to list subfolders of. Omit to list top-level folders.',
        },
        mailbox: {
          type: 'string',
          description: 'Email address or user ID of a shared mailbox. Omit to use your personal mailbox.',
        },
      },
    },
  },
  {
    name: 'mail_get_attachment',
    description:
      'Read the content of an email attachment. Automatically parses PDF, Word, Excel, PowerPoint, CSV, and HTML into readable text. Text files are returned as-is. Binary files return metadata only (no base64 dumps). Max 10 MB. Use attachment IDs from the mail_get_message response.',
    inputSchema: {
      type: 'object' as const,
      properties: {
        message_id: {
          type: 'string',
          description: 'The message ID containing the attachment.',
        },
        attachment_id: {
          type: 'string',
          description: 'The attachment ID from the mail_get_message response.',
        },
        mailbox: {
          type: 'string',
          description: 'The mailbox_context value from the original mail_list_messages response.',
        },
      },
      required: ['message_id', 'attachment_id'],
    },
  },
];
