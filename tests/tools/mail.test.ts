import { describe, it, expect, vi, beforeEach } from 'vitest';
import { MailTools, listMessagesInputSchema, getMessageInputSchema, listFoldersInputSchema, getAttachmentInputSchema } from '../../src/tools/mail.js';
import type { GraphClient, GraphMessage, GraphAttachment } from '../../src/graph/client.js';

describe('Mail Tools', () => {
  let mockGraphClient: GraphClient;
  let mailTools: MailTools;

  const mockMessages: GraphMessage[] = [
    {
      id: 'msg-1',
      subject: 'Test Subject 1',
      bodyPreview: 'This is a preview of the email body...',
      from: { emailAddress: { name: 'John Doe', address: 'john@example.com' } },
      toRecipients: [{ emailAddress: { name: 'Jane Doe', address: 'jane@example.com' } }],
      receivedDateTime: '2026-01-30T10:00:00Z',
      sentDateTime: '2026-01-30T09:59:00Z',
      hasAttachments: false,
      isRead: true,
      importance: 'normal',
      webLink: 'https://outlook.office.com/mail/id/msg-1',
    },
    {
      id: 'msg-2',
      subject: 'Test Subject 2',
      bodyPreview: 'Another preview...',
      from: { emailAddress: { name: 'Alice', address: 'alice@example.com' } },
      toRecipients: [{ emailAddress: { address: 'jane@example.com' } }],
      receivedDateTime: '2026-01-29T15:00:00Z',
      hasAttachments: true,
      isRead: false,
      importance: 'high',
    },
  ];

  beforeEach(() => {
    mockGraphClient = {
      listMessages: vi.fn().mockResolvedValue({ messages: mockMessages }),
      getMessage: vi.fn().mockResolvedValue(mockMessages[0]),
      listMailFolders: vi.fn().mockResolvedValue([
        { id: 'inbox-id', displayName: 'Inbox', parentFolderId: 'root', unreadItemCount: 5, totalItemCount: 100, childFolderCount: 2 },
        { id: 'sent-id', displayName: 'Sent Items', parentFolderId: 'root', unreadItemCount: 0, totalItemCount: 50, childFolderCount: 0 },
      ]),
      listChildFolders: vi.fn().mockResolvedValue([
        { id: 'subfolder-1', displayName: 'Subfolder A', parentFolderId: 'inbox-id', unreadItemCount: 1, totalItemCount: 10, childFolderCount: 0 },
      ]),
      getAttachment: vi.fn(),
    } as unknown as GraphClient;

    mailTools = new MailTools(mockGraphClient);
  });

  describe('listMessages', () => {
    it('should list messages with default parameters', async () => {
      const result = await mailTools.listMessages({});

      expect(mockGraphClient.listMessages).toHaveBeenCalledWith({
        folderId: undefined,
        top: 25,
        filter: undefined,
        orderBy: 'receivedDateTime desc',
      });

      expect(result).toHaveProperty('messages');
      expect(result).toHaveProperty('count', 2);
    });

    it('should include mailbox_context "personal" when no mailbox and no userContext', async () => {
      const result = await mailTools.listMessages({}) as Record<string, unknown>;

      expect(result['mailbox_context']).toBe('personal');
      expect(result['_note']).toContain('do NOT pass');
    });

    it('should include mailbox_context with user email when userContext is provided', async () => {
      const mailToolsWithContext = new MailTools(mockGraphClient, {
        userEmail: 'user@example.com',
        userId: 'user-id-123',
      });
      const result = await mailToolsWithContext.listMessages({}) as Record<string, unknown>;

      expect(result['mailbox_context']).toBe('user@example.com');
      expect(result['_note']).toContain("mailbox='user@example.com'");
    });

    it('should prefer explicit mailbox over userContext email', async () => {
      const mailToolsWithContext = new MailTools(mockGraphClient, {
        userEmail: 'user@example.com',
      });
      const result = await mailToolsWithContext.listMessages({ mailbox: 'shared@example.com' }) as Record<string, unknown>;

      expect(result['mailbox_context']).toBe('shared@example.com');
      expect(result['_note']).toContain("mailbox='shared@example.com'");
    });

    it('should resolve well-known folder names', async () => {
      await mailTools.listMessages({ folder: 'sent' });

      expect(mockGraphClient.listMessages).toHaveBeenCalledWith(
        expect.objectContaining({ folderId: 'sentitems' })
      );
    });

    it('should apply since filter', async () => {
      const since = '2026-01-29T00:00:00Z';
      await mailTools.listMessages({ since });

      expect(mockGraphClient.listMessages).toHaveBeenCalledWith(
        expect.objectContaining({
          filter: `receivedDateTime ge ${since}`,
        })
      );
    });

    it('should combine query and since filters', async () => {
      const query = "from/emailAddress/address eq 'john@example.com'";
      const since = '2026-01-29T00:00:00Z';
      await mailTools.listMessages({ query, since });

      expect(mockGraphClient.listMessages).toHaveBeenCalledWith(
        expect.objectContaining({
          filter: `(${query}) and receivedDateTime ge ${since}`,
        })
      );
    });
  });

  describe('getMessage', () => {
    it('should include body by default (include_body defaults to true)', async () => {
      const mockMessageWithBody = {
        ...mockMessages[0],
        body: { contentType: 'text', content: 'Plain text body' },
      };
      (mockGraphClient.getMessage as ReturnType<typeof vi.fn>).mockResolvedValue(mockMessageWithBody);

      const result = await mailTools.getMessage({ message_id: 'msg-1' });

      expect(mockGraphClient.getMessage).toHaveBeenCalledWith('msg-1', true, undefined);
      expect(result).toHaveProperty('body');
    });

    it('should not include body when include_body is false', async () => {
      const result = await mailTools.getMessage({ message_id: 'msg-1', include_body: false });

      expect(mockGraphClient.getMessage).toHaveBeenCalledWith('msg-1', false, undefined);
      expect(result).not.toHaveProperty('body');
    });

    it('should convert HTML body to plain text and preserve original HTML', async () => {
      const htmlContent = '<html><body><h1>Title</h1><p>Hello <b>world</b></p><script>alert(1)</script></body></html>';
      const mockMessageWithHtml = {
        ...mockMessages[0],
        body: { contentType: 'html', content: htmlContent },
      };
      (mockGraphClient.getMessage as ReturnType<typeof vi.fn>).mockResolvedValue(mockMessageWithHtml);

      const result = await mailTools.getMessage({ message_id: 'msg-1' }) as Record<string, unknown>;
      const body = result['body'] as { text: string; html: string };

      expect(body.text).toContain('Title');
      expect(body.text).toContain('Hello world');
      expect(body.text).not.toContain('<h1>');
      expect(body.text).not.toContain('alert');
      expect(body.html).toBe(htmlContent);
    });

    it('should return text-only body for plain text content', async () => {
      const mockMessageWithText = {
        ...mockMessages[0],
        body: { contentType: 'text', content: 'Just plain text' },
      };
      (mockGraphClient.getMessage as ReturnType<typeof vi.fn>).mockResolvedValue(mockMessageWithText);

      const result = await mailTools.getMessage({ message_id: 'msg-1' }) as Record<string, unknown>;
      const body = result['body'] as { text: string; html?: string };

      expect(body.text).toBe('Just plain text');
      expect(body).not.toHaveProperty('html');
    });

    it('should include CC and BCC when present', async () => {
      const mockMessageWithCC = {
        ...mockMessages[0],
        ccRecipients: [{ emailAddress: { name: 'CC User', address: 'cc@example.com' } }],
        bccRecipients: [{ emailAddress: { name: 'BCC User', address: 'bcc@example.com' } }],
      };
      (mockGraphClient.getMessage as ReturnType<typeof vi.fn>).mockResolvedValue(mockMessageWithCC);

      const result = await mailTools.getMessage({ message_id: 'msg-1', include_body: false }) as Record<string, unknown>;

      expect(result['cc']).toEqual(['cc@example.com']);
      expect(result['bcc']).toEqual(['bcc@example.com']);
    });

    it('should omit CC and BCC when empty', async () => {
      const mockMessageNoCC = {
        ...mockMessages[0],
        ccRecipients: [],
        bccRecipients: [],
      };
      (mockGraphClient.getMessage as ReturnType<typeof vi.fn>).mockResolvedValue(mockMessageNoCC);

      const result = await mailTools.getMessage({ message_id: 'msg-1', include_body: false }) as Record<string, unknown>;

      expect(result).not.toHaveProperty('cc');
      expect(result).not.toHaveProperty('bcc');
    });

    it('should include attachment metadata for non-inline attachments', async () => {
      const mockMessageWithAttachments: GraphMessage = {
        ...mockMessages[0]!,
        attachments: [
          { id: 'att-1', name: 'report.pdf', contentType: 'application/pdf', size: 50000, isInline: false },
          { id: 'att-2', name: 'logo.png', contentType: 'image/png', size: 1024, isInline: true },
        ],
      };
      (mockGraphClient.getMessage as ReturnType<typeof vi.fn>).mockResolvedValue(mockMessageWithAttachments);

      const result = await mailTools.getMessage({ message_id: 'msg-1', include_body: false }) as Record<string, unknown>;

      expect(result['attachments']).toEqual([
        { id: 'att-1', name: 'report.pdf', contentType: 'application/pdf', size: 50000 },
      ]);
      expect(result['inlineAttachmentCount']).toBe(1);
    });

    it('should throw enriched error for ErrorInvalidMailboxItemId with mailbox', async () => {
      const graphError = Object.assign(new Error("Item doesn't belong to the targeted mailbox"), {
        code: 'ErrorInvalidMailboxItemId',
        statusCode: 404,
      });
      (mockGraphClient.getMessage as ReturnType<typeof vi.fn>).mockRejectedValue(graphError);

      await expect(
        mailTools.getMessage({ message_id: 'msg-1', mailbox: 'shared@example.com' })
      ).rejects.toThrow(/does not belong to mailbox 'shared@example.com'/);
    });

    it('should throw enriched error for ErrorInvalidMailboxItemId without mailbox', async () => {
      const graphError = Object.assign(new Error("Item doesn't belong to the targeted mailbox"), {
        code: 'ErrorInvalidMailboxItemId',
        statusCode: 404,
      });
      (mockGraphClient.getMessage as ReturnType<typeof vi.fn>).mockRejectedValue(graphError);

      await expect(
        mailTools.getMessage({ message_id: 'msg-1' })
      ).rejects.toThrow(/does not belong to your personal mailbox/);
    });

    it('should re-throw non-ErrorInvalidMailboxItemId errors unchanged', async () => {
      const graphError = Object.assign(new Error('Server error'), {
        code: 'InternalServerError',
        statusCode: 500,
      });
      (mockGraphClient.getMessage as ReturnType<typeof vi.fn>).mockRejectedValue(graphError);

      await expect(
        mailTools.getMessage({ message_id: 'msg-1' })
      ).rejects.toThrow('Server error');
    });
  });

  describe('listFolders', () => {
    it('should list top-level mail folders when no parent_folder_id', async () => {
      const result = await mailTools.listFolders({});

      expect(mockGraphClient.listMailFolders).toHaveBeenCalledWith(undefined);
      expect(mockGraphClient.listChildFolders).not.toHaveBeenCalled();
      expect(result).toHaveProperty('folders');
      expect(result).toHaveProperty('count', 2);
    });

    it('should include parentFolderId in formatted output', async () => {
      const result = await mailTools.listFolders({}) as { folders: Array<Record<string, unknown>> };

      expect(result.folders[0]).toHaveProperty('parentFolderId', 'root');
    });

    it('should call listChildFolders when parent_folder_id is provided', async () => {
      const result = await mailTools.listFolders({ parent_folder_id: 'inbox-id' });

      expect(mockGraphClient.listChildFolders).toHaveBeenCalledWith('inbox-id', undefined);
      expect(mockGraphClient.listMailFolders).not.toHaveBeenCalled();
      expect(result).toHaveProperty('count', 1);
    });

    it('should resolve well-known folder names for parent_folder_id', async () => {
      await mailTools.listFolders({ parent_folder_id: 'sent' });

      expect(mockGraphClient.listChildFolders).toHaveBeenCalledWith('sentitems', undefined);
    });

    it('should resolve well-known folder names case-insensitively', async () => {
      await mailTools.listFolders({ parent_folder_id: 'INBOX' });

      expect(mockGraphClient.listChildFolders).toHaveBeenCalledWith('inbox', undefined);
    });

    it('should pass mailbox to listChildFolders', async () => {
      await mailTools.listFolders({ parent_folder_id: 'inbox', mailbox: 'shared@example.com' });

      expect(mockGraphClient.listChildFolders).toHaveBeenCalledWith('inbox', 'shared@example.com');
    });

    it('should pass mailbox to listMailFolders', async () => {
      await mailTools.listFolders({ mailbox: 'shared@example.com' });

      expect(mockGraphClient.listMailFolders).toHaveBeenCalledWith('shared@example.com');
    });
  });

  describe('getAttachment', () => {
    it('should return parsed text for a text file attachment', async () => {
      const textAttachment: GraphAttachment = {
        id: 'att-1',
        '@odata.type': '#microsoft.graph.fileAttachment',
        name: 'notes.txt',
        contentType: 'text/plain',
        size: 100,
        isInline: false,
        contentBytes: Buffer.from('Hello from attachment').toString('base64'),
      };
      (mockGraphClient.getAttachment as ReturnType<typeof vi.fn>).mockResolvedValue(textAttachment);

      const result = await mailTools.getAttachment({ message_id: 'msg-1', attachment_id: 'att-1' }) as Record<string, unknown>;

      expect(result['content']).toBe('Hello from attachment');
      expect(result['contentType']).toBe('text');
      expect(result['type']).toBe('fileAttachment');
      expect(result['name']).toBe('notes.txt');
    });

    it('should return parsed text for a CSV attachment', async () => {
      const csvContent = 'Name,Age\nAlice,30';
      const csvAttachment: GraphAttachment = {
        id: 'att-csv',
        '@odata.type': '#microsoft.graph.fileAttachment',
        name: 'data.csv',
        contentType: 'text/csv',
        size: csvContent.length,
        isInline: false,
        contentBytes: Buffer.from(csvContent).toString('base64'),
      };
      (mockGraphClient.getAttachment as ReturnType<typeof vi.fn>).mockResolvedValue(csvAttachment);

      const result = await mailTools.getAttachment({ message_id: 'msg-1', attachment_id: 'att-csv' }) as Record<string, unknown>;

      // text/csv is a text MIME type, so it's returned as text directly
      expect(result['content']).toBe(csvContent);
      expect(result['contentType']).toBe('text');
    });

    it('should return null content for binary attachments', async () => {
      const binaryAttachment: GraphAttachment = {
        id: 'att-bin',
        '@odata.type': '#microsoft.graph.fileAttachment',
        name: 'image.png',
        contentType: 'image/png',
        size: 500,
        isInline: false,
        contentBytes: Buffer.from([0x89, 0x50, 0x4e, 0x47]).toString('base64'),
      };
      (mockGraphClient.getAttachment as ReturnType<typeof vi.fn>).mockResolvedValue(binaryAttachment);

      const result = await mailTools.getAttachment({ message_id: 'msg-1', attachment_id: 'att-bin' }) as Record<string, unknown>;

      expect(result['content']).toBeNull();
      expect(result['contentType']).toBe('binary');
      expect(result['_note']).toContain('Binary file');
    });

    it('should throw for attachments exceeding 10MB', async () => {
      // Create a large base64 string (>10MB decoded)
      const largeBuffer = Buffer.alloc(11 * 1024 * 1024);
      const largeAttachment: GraphAttachment = {
        id: 'att-large',
        '@odata.type': '#microsoft.graph.fileAttachment',
        name: 'huge.bin',
        contentType: 'application/octet-stream',
        size: largeBuffer.length,
        isInline: false,
        contentBytes: largeBuffer.toString('base64'),
      };
      (mockGraphClient.getAttachment as ReturnType<typeof vi.fn>).mockResolvedValue(largeAttachment);

      await expect(
        mailTools.getAttachment({ message_id: 'msg-1', attachment_id: 'att-large' })
      ).rejects.toThrow(/exceeds the 10 MB limit/);
    });

    it('should throw for attachment with missing contentBytes', async () => {
      const emptyAttachment: GraphAttachment = {
        id: 'att-empty',
        '@odata.type': '#microsoft.graph.fileAttachment',
        name: 'missing.txt',
        contentType: 'text/plain',
        size: 0,
        isInline: false,
      };
      (mockGraphClient.getAttachment as ReturnType<typeof vi.fn>).mockResolvedValue(emptyAttachment);

      await expect(
        mailTools.getAttachment({ message_id: 'msg-1', attachment_id: 'att-empty' })
      ).rejects.toThrow(/content is empty/);
    });

    it('should handle reference attachments', async () => {
      const refAttachment: GraphAttachment = {
        id: 'att-ref',
        '@odata.type': '#microsoft.graph.referenceAttachment',
        name: 'shared-doc.docx',
        contentType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
        size: 0,
        isInline: false,
        sourceUrl: 'https://contoso.sharepoint.com/sites/docs/shared-doc.docx',
      };
      (mockGraphClient.getAttachment as ReturnType<typeof vi.fn>).mockResolvedValue(refAttachment);

      const result = await mailTools.getAttachment({ message_id: 'msg-1', attachment_id: 'att-ref' }) as Record<string, unknown>;

      expect(result['type']).toBe('referenceAttachment');
      expect(result['sourceUrl']).toBe('https://contoso.sharepoint.com/sites/docs/shared-doc.docx');
      expect(result['content']).toBeNull();
      expect(result['_note']).toContain('sp_get_file');
    });

    it('should handle item attachments', async () => {
      const itemAttachment: GraphAttachment = {
        id: 'att-item',
        '@odata.type': '#microsoft.graph.itemAttachment',
        name: 'Forwarded Email',
        contentType: 'message/rfc822',
        size: 2048,
        isInline: false,
        item: { subject: 'Original Subject', '@odata.type': '#microsoft.graph.message' },
      };
      (mockGraphClient.getAttachment as ReturnType<typeof vi.fn>).mockResolvedValue(itemAttachment);

      const result = await mailTools.getAttachment({ message_id: 'msg-1', attachment_id: 'att-item' }) as Record<string, unknown>;

      expect(result['type']).toBe('itemAttachment');
      expect(result['item']).toEqual({ subject: 'Original Subject', '@odata.type': '#microsoft.graph.message' });
      expect(result['content']).toBeNull();
      expect(result['_note']).toContain('Item attachments');
    });

    it('should forward mailbox parameter to graphClient', async () => {
      const textAttachment: GraphAttachment = {
        id: 'att-1',
        '@odata.type': '#microsoft.graph.fileAttachment',
        name: 'test.txt',
        contentType: 'text/plain',
        size: 5,
        isInline: false,
        contentBytes: Buffer.from('hello').toString('base64'),
      };
      (mockGraphClient.getAttachment as ReturnType<typeof vi.fn>).mockResolvedValue(textAttachment);

      await mailTools.getAttachment({ message_id: 'msg-1', attachment_id: 'att-1', mailbox: 'shared@example.com' });

      expect(mockGraphClient.getAttachment).toHaveBeenCalledWith('msg-1', 'att-1', 'shared@example.com');
    });

    it('should throw enriched error for ErrorInvalidMailboxItemId', async () => {
      const graphError = Object.assign(new Error('Invalid ID'), {
        code: 'ErrorInvalidMailboxItemId',
        statusCode: 404,
      });
      (mockGraphClient.getAttachment as ReturnType<typeof vi.fn>).mockRejectedValue(graphError);

      await expect(
        mailTools.getAttachment({ message_id: 'msg-1', attachment_id: 'att-1', mailbox: 'shared@example.com' })
      ).rejects.toThrow(/does not belong to mailbox 'shared@example.com'/);
    });
  });
});

describe('Mail Input Schemas', () => {
  describe('listMessagesInputSchema', () => {
    it('should accept valid input', () => {
      const result = listMessagesInputSchema.safeParse({
        folder: 'inbox',
        top: 50,
        query: "isRead eq false",
        since: '2026-01-29T00:00:00Z',
      });

      expect(result.success).toBe(true);
    });

    it('should reject invalid top value', () => {
      const result = listMessagesInputSchema.safeParse({
        top: 200, // Max is 100
      });

      expect(result.success).toBe(false);
    });

    it('should reject invalid datetime', () => {
      const result = listMessagesInputSchema.safeParse({
        since: 'not-a-date',
      });

      expect(result.success).toBe(false);
    });
  });

  describe('getMessageInputSchema', () => {
    it('should require message_id', () => {
      const result = getMessageInputSchema.safeParse({});

      expect(result.success).toBe(false);
    });

    it('should reject empty message_id', () => {
      const result = getMessageInputSchema.safeParse({
        message_id: '',
      });

      expect(result.success).toBe(false);
    });

    it('should default include_body to true', () => {
      const result = getMessageInputSchema.parse({ message_id: 'msg-1' });
      expect(result.include_body).toBe(true);
    });
  });

  describe('listFoldersInputSchema', () => {
    it('should accept empty input', () => {
      const result = listFoldersInputSchema.safeParse({});
      expect(result.success).toBe(true);
    });

    it('should accept parent_folder_id', () => {
      const result = listFoldersInputSchema.safeParse({ parent_folder_id: 'inbox' });
      expect(result.success).toBe(true);
    });

    it('should reject empty parent_folder_id', () => {
      const result = listFoldersInputSchema.safeParse({ parent_folder_id: '' });
      expect(result.success).toBe(false);
    });
  });

  describe('getAttachmentInputSchema', () => {
    it('should require both message_id and attachment_id', () => {
      expect(getAttachmentInputSchema.safeParse({}).success).toBe(false);
      expect(getAttachmentInputSchema.safeParse({ message_id: 'msg-1' }).success).toBe(false);
      expect(getAttachmentInputSchema.safeParse({ attachment_id: 'att-1' }).success).toBe(false);
    });

    it('should accept valid input', () => {
      const result = getAttachmentInputSchema.safeParse({
        message_id: 'msg-1',
        attachment_id: 'att-1',
      });
      expect(result.success).toBe(true);
    });

    it('should accept optional mailbox', () => {
      const result = getAttachmentInputSchema.safeParse({
        message_id: 'msg-1',
        attachment_id: 'att-1',
        mailbox: 'shared@example.com',
      });
      expect(result.success).toBe(true);
    });

    it('should reject empty message_id or attachment_id', () => {
      expect(getAttachmentInputSchema.safeParse({ message_id: '', attachment_id: 'att-1' }).success).toBe(false);
      expect(getAttachmentInputSchema.safeParse({ message_id: 'msg-1', attachment_id: '' }).success).toBe(false);
    });
  });
});
