import { describe, it, expect, vi, beforeEach } from 'vitest';
import { MailTools, listMessagesInputSchema, getMessageInputSchema } from '../../src/tools/mail.js';
import type { GraphClient, GraphMessage } from '../../src/graph/client.js';

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
        { id: 'inbox', displayName: 'Inbox', unreadItemCount: 5, totalItemCount: 100 },
        { id: 'sent', displayName: 'Sent Items', unreadItemCount: 0, totalItemCount: 50 },
      ]),
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
    it('should get message without body by default', async () => {
      const result = await mailTools.getMessage({ message_id: 'msg-1' });

      expect(mockGraphClient.getMessage).toHaveBeenCalledWith('msg-1', false, undefined);
      expect(result).not.toHaveProperty('body');
    });

    it('should get message with body when requested', async () => {
      const mockMessageWithBody = {
        ...mockMessages[0],
        body: { contentType: 'html', content: '<p>Full body</p>' },
      };
      (mockGraphClient.getMessage as ReturnType<typeof vi.fn>).mockResolvedValue(mockMessageWithBody);

      const result = await mailTools.getMessage({ message_id: 'msg-1', include_body: true });

      expect(mockGraphClient.getMessage).toHaveBeenCalledWith('msg-1', true, undefined);
      expect(result).toHaveProperty('body');
    });
  });

  describe('listFolders', () => {
    it('should list mail folders', async () => {
      const result = await mailTools.listFolders({});

      expect(mockGraphClient.listMailFolders).toHaveBeenCalled();
      expect(result).toHaveProperty('folders');
      expect(result).toHaveProperty('count', 2);
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
  });
});
