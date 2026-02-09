import { describe, it, expect, vi, beforeEach } from 'vitest';
import {
  SharePointTools,
  listSitesInputSchema,
  listDrivesInputSchema,
  listChildrenInputSchema,
  getFileInputSchema,
  searchFilesInputSchema,
} from '../../src/tools/sharepoint.js';
import type { GraphClient, GraphSite, GraphDrive, GraphDriveItem, SearchHit } from '../../src/graph/client.js';

describe('SharePoint Tools', () => {
  let mockGraphClient: GraphClient;
  let spTools: SharePointTools;

  const mockSites: GraphSite[] = [
    {
      id: 'contoso.sharepoint.com,guid1,guid2',
      name: 'TeamSite',
      webUrl: 'https://contoso.sharepoint.com/sites/TeamSite',
      description: 'Team Site',
    },
    {
      id: 'contoso.sharepoint.com,guid3,guid4',
      name: 'ProjectSite',
      webUrl: 'https://contoso.sharepoint.com/sites/ProjectSite',
    },
  ];

  const mockDrives: GraphDrive[] = [
    {
      id: 'drive-abc-123',
      name: 'Documents',
      driveType: 'documentLibrary',
      webUrl: 'https://contoso.sharepoint.com/sites/TeamSite/Documents',
    },
    {
      id: 'drive-def-456',
      name: 'Shared Documents',
      driveType: 'documentLibrary',
      webUrl: 'https://contoso.sharepoint.com/sites/TeamSite/Shared',
    },
  ];

  const mockDriveItems: GraphDriveItem[] = [
    {
      id: 'item-1',
      name: 'Report.docx',
      webUrl: 'https://contoso.sharepoint.com/sites/TeamSite/Documents/Report.docx',
      size: 51200,
      createdDateTime: '2026-01-15T10:00:00Z',
      lastModifiedDateTime: '2026-01-20T14:30:00Z',
      file: { mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' },
    },
    {
      id: 'item-2',
      name: 'Archive',
      webUrl: 'https://contoso.sharepoint.com/sites/TeamSite/Documents/Archive',
      createdDateTime: '2025-06-01T08:00:00Z',
      lastModifiedDateTime: '2026-01-10T16:00:00Z',
      folder: { childCount: 5 },
    },
  ];

  beforeEach(() => {
    mockGraphClient = {
      listSites: vi.fn().mockResolvedValue(mockSites),
      listSiteDrives: vi.fn().mockResolvedValue(mockDrives),
      getMyDrive: vi.fn(),
      listDriveItems: vi.fn().mockResolvedValue(mockDriveItems),
      getDriveItem: vi.fn().mockResolvedValue(mockDriveItems[0]),
      getFileContent: vi.fn().mockResolvedValue({
        content: Buffer.from('file content'),
        mimeType: 'text/plain',
        size: 12,
      }),
      searchDriveItems: vi.fn().mockResolvedValue({
        hits: [
          {
            hitId: 'hit-1',
            rank: 1,
            summary: 'Die <c0>Ersthelfer</c0> in <c0>Berlin</c0> sind...',
            resource: {
              id: 'item-search-1',
              name: 'Ersthelfer-Liste.docx',
              webUrl: 'https://contoso.sharepoint.com/sites/TeamSite/Documents/Ersthelfer-Liste.docx',
              size: 25000,
              lastModifiedDateTime: '2026-01-25T09:00:00Z',
              file: { mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' },
              parentReference: { driveId: 'drive-abc-123' },
            },
          },
        ] as SearchHit[],
        total: 1,
        moreResultsAvailable: false,
      }),
    } as unknown as GraphClient;

    spTools = new SharePointTools(mockGraphClient);
  });

  describe('listSites', () => {
    it('should list sites with default wildcard search', async () => {
      const result = await spTools.listSites({});

      expect(mockGraphClient.listSites).toHaveBeenCalledWith({ search: undefined });
      expect(result).toHaveProperty('sites');
      expect(result).toHaveProperty('count', 2);
    });

    it('should filter sites by query', async () => {
      await spTools.listSites({ query: 'Team' });

      expect(mockGraphClient.listSites).toHaveBeenCalledWith({ search: 'Team' });
    });
  });

  describe('listDrives', () => {
    it('should list drives for a specific site', async () => {
      const result = await spTools.listDrives({
        site_id: 'contoso.sharepoint.com,guid1,guid2',
      });

      expect(mockGraphClient.listSiteDrives).toHaveBeenCalledWith(
        'contoso.sharepoint.com,guid1,guid2'
      );
      expect(result).toHaveProperty('drives');
      expect(result).toHaveProperty('count', 2);
      expect(result).toHaveProperty('site_id', 'contoso.sharepoint.com,guid1,guid2');
    });

    it('should require site_id', () => {
      const result = listDrivesInputSchema.safeParse({});
      expect(result.success).toBe(false);
    });

    it('should reject invalid site_id format', () => {
      const result = listDrivesInputSchema.safeParse({
        site_id: 'invalid<>id',
      });
      expect(result.success).toBe(false);
    });
  });

  describe('listChildren', () => {
    it('should list root items of a drive', async () => {
      const result = await spTools.listChildren({ drive_id: 'drive-abc-123' });

      expect(mockGraphClient.listDriveItems).toHaveBeenCalledWith({
        driveId: 'drive-abc-123',
        itemId: undefined,
      });
      expect(result).toHaveProperty('items');
      expect(result).toHaveProperty('count', 2);
    });

    it('should list children of a specific folder', async () => {
      await spTools.listChildren({ drive_id: 'drive-abc-123', item_id: 'item-2' });

      expect(mockGraphClient.listDriveItems).toHaveBeenCalledWith({
        driveId: 'drive-abc-123',
        itemId: 'item-2',
      });
    });

    it('should throw enriched error for 404', async () => {
      const notFoundError = Object.assign(new Error('Item not found'), {
        code: 'itemNotFound',
        statusCode: 404,
      });
      (mockGraphClient.listDriveItems as ReturnType<typeof vi.fn>).mockRejectedValue(notFoundError);

      await expect(
        spTools.listChildren({ drive_id: 'drive-abc-123' })
      ).rejects.toThrow(/Item not found.*call sp_list_drives/);
    });
  });

  describe('getFile', () => {
    it('should get file with text content', async () => {
      const result = await spTools.getFile({
        drive_id: 'drive-abc-123',
        item_id: 'item-1',
      }) as Record<string, unknown>;

      expect(mockGraphClient.getDriveItem).toHaveBeenCalledWith('drive-abc-123', 'item-1');
      expect(result['name']).toBe('Report.docx');
    });

    it('should throw enriched error for 404', async () => {
      const notFoundError = Object.assign(new Error('Item not found'), {
        code: 'itemNotFound',
        statusCode: 404,
      });
      (mockGraphClient.getDriveItem as ReturnType<typeof vi.fn>).mockRejectedValue(notFoundError);

      await expect(
        spTools.getFile({ drive_id: 'drive-abc-123', item_id: 'item-1' })
      ).rejects.toThrow(/File not found.*call sp_list_children/);
    });

    it('should reject non-file items', async () => {
      (mockGraphClient.getDriveItem as ReturnType<typeof vi.fn>).mockResolvedValue(mockDriveItems[1]); // folder

      await expect(
        spTools.getFile({ drive_id: 'drive-abc-123', item_id: 'item-2' })
      ).rejects.toThrow('The specified item is not a file');
    });
  });

  describe('searchFiles', () => {
    it('should search for files and return formatted results', async () => {
      const result = await spTools.searchFiles({ query: 'Ersthelfer Berlin' }) as Record<string, unknown>;

      expect(mockGraphClient.searchDriveItems).toHaveBeenCalledWith({
        query: 'Ersthelfer Berlin',
        size: 10,
      });
      expect(result['total']).toBe(1);

      const results = result['results'] as Record<string, unknown>[];
      expect(results).toHaveLength(1);
      expect(results[0]['name']).toBe('Ersthelfer-Liste.docx');
      expect(results[0]['drive_id']).toBe('drive-abc-123');
      expect(results[0]['item_id']).toBe('item-search-1');
      expect(results[0]['type']).toBe('file');
      // Summary should have highlight tags removed
      expect(results[0]['summary']).toBe('Die Ersthelfer in Berlin sind...');
    });

    it('should use custom size parameter', async () => {
      await spTools.searchFiles({ query: 'test', size: 5 });

      expect(mockGraphClient.searchDriveItems).toHaveBeenCalledWith({
        query: 'test',
        size: 5,
      });
    });

    it('should handle empty search results', async () => {
      (mockGraphClient.searchDriveItems as ReturnType<typeof vi.fn>).mockResolvedValue({
        hits: [],
        total: 0,
        moreResultsAvailable: false,
      });

      const result = await spTools.searchFiles({ query: 'nonexistent' }) as Record<string, unknown>;

      expect(result['total']).toBe(0);
      expect(result['results']).toEqual([]);
    });
  });
});

describe('SharePoint Input Schemas', () => {
  describe('listDrivesInputSchema', () => {
    it('should require site_id', () => {
      const result = listDrivesInputSchema.safeParse({});
      expect(result.success).toBe(false);
    });

    it('should accept valid site_id', () => {
      const result = listDrivesInputSchema.safeParse({
        site_id: 'contoso.sharepoint.com,guid1-abc,guid2-def',
      });
      expect(result.success).toBe(true);
    });
  });

  describe('searchFilesInputSchema', () => {
    it('should require query', () => {
      const result = searchFilesInputSchema.safeParse({});
      expect(result.success).toBe(false);
    });

    it('should reject empty query', () => {
      const result = searchFilesInputSchema.safeParse({ query: '' });
      expect(result.success).toBe(false);
    });

    it('should accept valid search', () => {
      const result = searchFilesInputSchema.safeParse({
        query: 'Ersthelfer Berlin',
        size: 5,
      });
      expect(result.success).toBe(true);
    });

    it('should reject size > 25', () => {
      const result = searchFilesInputSchema.safeParse({
        query: 'test',
        size: 50,
      });
      expect(result.success).toBe(false);
    });
  });

  describe('listChildrenInputSchema', () => {
    it('should require drive_id', () => {
      const result = listChildrenInputSchema.safeParse({});
      expect(result.success).toBe(false);
    });
  });

  describe('getFileInputSchema', () => {
    it('should require both drive_id and item_id', () => {
      expect(getFileInputSchema.safeParse({}).success).toBe(false);
      expect(getFileInputSchema.safeParse({ drive_id: 'abc' }).success).toBe(false);
      expect(getFileInputSchema.safeParse({ item_id: 'abc' }).success).toBe(false);
    });

    it('should accept valid IDs', () => {
      const result = getFileInputSchema.safeParse({
        drive_id: 'b!abc-123',
        item_id: '01ABCDEFGH',
      });
      expect(result.success).toBe(true);
    });
  });
});
