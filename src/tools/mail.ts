import { z } from 'zod';
import { GraphClient, type GraphMessage, type GraphMailFolder } from '../graph/client.js';
import { logger } from '../utils/logger.js';
import { stripHtmlTags, isParsableMimeType, parseFileContent } from '../utils/file-parser.js';
import { isTextMimeType, formatFileSize } from '../utils/content-fetcher.js';
import { rememberId, resolveId } from '../utils/id-cache.js';
import { handleStore, type HandleKind } from '../utils/handle-store.js';

/**
 * User context from the authenticated session.
 * Used to include user identity in responses, making them portable across MCP sessions.
 */
export interface UserContext {
  userEmail?: string;
  userId?: string;
}

// Shared-mailbox selector. The personal mailbox is the default (/me) and needs NO
// value here. Empty/whitespace is coerced to "omitted" so a stray "" never trips Zod
// before the GraphClient mailbox normalization can run.
const mailboxSchema = z
  .string()
  .max(256)
  .transform((v) => (v.trim() ? v.trim() : undefined))
  .describe('Email address of a SHARED mailbox. OMIT for your own mailbox (the default). Never pass "me" or an empty value.');

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
  attachment_id: z
    .string()
    .optional()
    .describe('The attachment ID. Omit it to auto-select when the message has a single attachment.'),
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
function formatMessage(
  message: GraphMessage,
  includeBody: boolean = false,
  refs?: { self?: string; attachments?: Record<string, string> }
): object {
  const formatted: Record<string, unknown> = {
    id: refs?.self ?? message.id,
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
        id: refs?.attachments?.[a.id] ?? a.id,
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

  // Per-user namespace for the Graph ID resolution cache (id-cache fallback).
  private get userKey(): string {
    return this.userContext?.userId ?? this.userContext?.userEmail ?? 'default';
  }

  // Authenticated identity for the handle store. No 'default' fallback: without a
  // real identity we do NOT mint/resolve handles, we fall back to the raw id.
  private get identityKey(): string | null {
    return this.userContext?.userId ?? this.userContext?.userEmail ?? null;
  }

  // Matches a handle we minted (m_/a_ + 12 hex). Real Graph ids never match this.
  static readonly HANDLE_RE = /^[ma]_[0-9a-f]{12}$/;

  // Mint a handle for an id we are handing out. Without identity, return the raw
  // id so the model still gets something the raw-id fallback path can resolve.
  private async mintRef(kind: HandleKind, realId: string, mailbox?: string): Promise<string> {
    const key = this.identityKey;
    if (!key) return realId;
    return handleStore.mint(key, kind, { realId, mailbox });
  }

  // Resolve an incoming message_id/attachment_id. Handle -> stored payload (real id
  // + mailbox). If it looks like our handle but does not resolve, it is stale: tell
  // the model to re-list rather than firing a doomed Graph call. Otherwise fall back
  // to the id-cache (re-encoded long id) then passthrough (a correct raw id).
  private async resolveRef(
    input: string,
    mailbox?: string
  ): Promise<{ id: string; mailbox?: string }> {
    const key = this.identityKey;
    if (key && MailTools.HANDLE_RE.test(input)) {
      const payload = await handleStore.resolve(key, input);
      if (payload) return { id: payload.realId, mailbox: payload.mailbox ?? mailbox };
      throw new Error(
        'This reference is no longer valid (it may have expired). Call mail_list_messages again and use the id from the new results.'
      );
    }
    return { id: resolveId(this.userKey, input) ?? input, mailbox };
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

    // Remember the canonical IDs for the id-cache fallback, and mint a short handle
    // per message so the model never has to relay the long opaque Graph id.
    result.messages.forEach((m) => rememberId(this.userKey, m.id));
    const messages = await Promise.all(
      result.messages.map(async (m) =>
        formatMessage(m, false, { self: await this.mintRef('msg', m.id, validated.mailbox) })
      )
    );

    const sharedMailbox = validated.mailbox;

    return {
      messages,
      count: result.messages.length,
      hasMore: !!result.nextLink,
      mailbox_context: sharedMailbox ?? 'personal',
      _note: sharedMailbox
        ? `These messages are from the shared mailbox '${sharedMailbox}'. To open one with mail_get_message or mail_get_attachment, pass the message's id (you do not need to pass mailbox; the id already carries it).`
        : `These messages are from your own mailbox. To open one with mail_get_message or mail_get_attachment, pass only the message's id and OMIT the mailbox parameter.`,
    };
  }

  /**
   * Get a specific message by ID
   */
  async getMessage(input: GetMessageInput): Promise<object> {
    const validated = getMessageInputSchema.parse(input);
    const incoming = validated.message_id;
    const { id: messageId, mailbox } = await this.resolveRef(incoming, validated.mailbox);

    logger.debug({ messageId, includeBody: validated.include_body }, 'Getting message');

    try {
      const message = await this.graphClient.getMessage(messageId, validated.include_body, mailbox);

      // id-cache fallback for re-encoded relays.
      rememberId(this.userKey, message.id);
      (message.attachments ?? []).forEach((a) => rememberId(this.userKey, a.id));

      // Re-use the incoming handle for the message when the call arrived via one;
      // mint attachment handles so mail_get_attachment uses the same mechanism.
      const self = MailTools.HANDLE_RE.test(incoming)
        ? incoming
        : await this.mintRef('msg', message.id, mailbox);
      const attachments: Record<string, string> = {};
      for (const a of message.attachments ?? []) {
        if (a.isInline) continue;
        attachments[a.id] = await this.mintRef('att', a.id, mailbox);
      }

      return formatMessage(message, validated.include_body, { self, attachments });
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
      if (code === 'ErrorInvalidIdMalformed') {
        const enrichedError = new Error(
          'The message_id is not a valid Microsoft Graph ID. Call mail_list_messages again and pass the exact "id" value from its response verbatim (do not shorten, truncate, or modify it).'
        ) as Error & { code?: string; statusCode?: number };
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
    const { id: messageId, mailbox } = await this.resolveRef(validated.message_id, validated.mailbox);

    // The model often omits the attachment id, or relays a re-encoded/handle one.
    // Resolve a provided id (handle or raw) against what we handed out; when none is
    // given, resolve it from the message: auto-select a lone attachment, else list.
    const requested = validated.attachment_id?.trim();
    let attachmentId: string | undefined;
    if (requested) {
      attachmentId = (await this.resolveRef(requested, mailbox)).id;
    }
    if (!attachmentId) {
      const candidates = (await this.graphClient.listAttachments(messageId, mailbox)).filter(
        (a) => !a.isInline
      );
      candidates.forEach((a) => rememberId(this.userKey, a.id));

      if (candidates.length === 0) {
        throw new Error('This message has no readable (non-inline) attachments.');
      }
      if (candidates.length > 1) {
        return {
          needs_selection: true,
          message: `This message has ${candidates.length} attachments. Call mail_get_attachment again with one of the listed attachment_id values.`,
          attachments: await Promise.all(
            candidates.map(async (a) => ({
              attachment_id: await this.mintRef('att', a.id, mailbox),
              name: a.name,
              contentType: a.contentType,
              size: a.size,
            }))
          ),
        };
      }
      attachmentId = candidates[0]?.id;
      if (!attachmentId) {
        throw new Error('This message has no readable (non-inline) attachments.');
      }
    }

    logger.debug(
      { messageId, attachmentId },
      'Getting attachment'
    );

    try {
      const attachment = await this.graphClient.getAttachment(
        messageId,
        attachmentId,
        mailbox
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
      const MAX_SIZE = 20 * 1024 * 1024; // 20MB
      if (buffer.length > MAX_SIZE) {
        throw new Error(`Attachment size (${formatFileSize(buffer.length)}) exceeds the 20 MB limit.`);
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
          description: 'Email address of a SHARED mailbox. OMIT for your own mailbox (the default).',
        },
      },
    },
  },
  {
    name: 'mail_get_message',
    description:
      'Get full email details including body (HTML auto-converted to plain text), CC/BCC recipients, and attachment metadata. Body is included by default. The message_id MUST be the exact id from a recent mail_list_messages call. For your own mailbox, OMIT the mailbox parameter (it defaults to you); only set mailbox to a shared mailbox email address when the listed messages came from one. To read an attachment, call mail_get_attachment with the same message_id.',
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
          description: 'Email address of a SHARED mailbox. OMIT for your own mailbox (the default). Never pass "me" or an empty string.',
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
          description: 'Email address of a SHARED mailbox. OMIT for your own mailbox (the default).',
        },
      },
    },
  },
  {
    name: 'mail_get_attachment',
    description:
      'Read the content of an email attachment. Automatically parses PDF, Word, Excel, PowerPoint, CSV, and HTML into readable text. Text files are returned as-is. Binary files return metadata only (no base64 dumps). Max 20 MB. Pass attachment_id from a mail_get_message response, or omit it to auto-select when the message has a single attachment (if there are several, the tool returns the list to choose from).',
    inputSchema: {
      type: 'object' as const,
      properties: {
        message_id: {
          type: 'string',
          description: 'The message ID containing the attachment.',
        },
        attachment_id: {
          type: 'string',
          description: 'The attachment ID from a mail_get_message response. Optional: omit it to auto-select when the message has exactly one attachment.',
        },
        mailbox: {
          type: 'string',
          description: 'Email address of a SHARED mailbox. OMIT for your own mailbox (the default).',
        },
      },
      required: ['message_id'],
    },
  },
];
