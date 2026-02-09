import { describe, it, expect, vi, beforeEach } from 'vitest';
import {
  SharePointTools,
  listSitesInputSchema,
  listDrivesInputSchema,
  listChildrenInputSchema,
  getFileInputSchema,
  searchFilesInputSchema,
  searchAndReadInputSchema,
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

  const mockSearchHits: SearchHit[] = [
    {
      hitId: 'hit-1',
      rank: 1,
      summary: 'Die <c0>Ersthelfer</c0> in <c0>Berlin</c0> sind...<ddd/>',
      resource: {
        id: 'item-search-1',
        name: 'Ersthelfer-Liste.docx',
        webUrl: 'https://contoso.sharepoint.com/sites/TeamSite/Documents/Ersthelfer-Liste.docx',
        size: 25000,
        lastModifiedDateTime: '2026-01-25T09:00:00Z',
        file: { mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' },
        parentReference: {
          driveId: 'drive-abc-123',
          path: '/drives/drive-abc-123/root:/Documents',
        },
      },
    },
    {
      hitId: 'hit-2',
      rank: 2,
      summary: 'Sicherheitsregeln für <c0>Ersthelfer</c0>...',
      resource: {
        id: 'item-search-2',
        name: 'Sicherheit.pdf',
        webUrl: 'https://contoso.sharepoint.com/sites/TeamSite/Documents/Sicherheit.pdf',
        size: 150000,
        lastModifiedDateTime: '2025-12-01T10:00:00Z',
        file: { mimeType: 'application/pdf' },
        parentReference: {
          driveId: 'drive-abc-123',
          path: '/drives/drive-abc-123/root:/Documents',
        },
      },
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
        content: Buffer.from('file content here'),
        mimeType: 'text/plain',
        size: 17,
      }),
      searchDriveItems: vi.fn().mockResolvedValue({
        hits: mockSearchHits,
        total: 2,
        moreResultsAvailable: false,
      }),
      resolveSiteWebUrl: vi.fn().mockResolvedValue('https://contoso.sharepoint.com/sites/TeamSite'),
    } as unknown as GraphClient;

    spTools = new SharePointTools(mockGraphClient);
  });

  // ===== listSites =====

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

  // ===== listDrives =====

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

  // ===== listChildren =====

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

  // ===== getFile =====

  describe('getFile', () => {
    it('should get file with text content via fetchAndParseContent', async () => {
      const result = await spTools.getFile({
        drive_id: 'drive-abc-123',
        item_id: 'item-1',
      }) as Record<string, unknown>;

      expect(mockGraphClient.getDriveItem).toHaveBeenCalledWith('drive-abc-123', 'item-1');
      expect(mockGraphClient.getFileContent).toHaveBeenCalledWith('drive-abc-123', 'item-1', 10 * 1024 * 1024);
      expect(result['name']).toBe('Report.docx');
      expect(result['content']).toBe('file content here');
      expect(result['contentType']).toBe('text');
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

    it('should handle files that are too large', async () => {
      (mockGraphClient.getDriveItem as ReturnType<typeof vi.fn>).mockResolvedValue({
        ...mockDriveItems[0],
        size: 20 * 1024 * 1024, // 20MB
      });

      const result = await spTools.getFile({
        drive_id: 'drive-abc-123',
        item_id: 'item-1',
      }) as Record<string, unknown>;

      expect(result['content']).toBeNull();
      expect(result['contentError']).toContain('too large');
    });
  });

  // ===== searchFiles =====

  describe('searchFiles', () => {
    it('should search for files and return formatted results', async () => {
      const result = await spTools.searchFiles({ query: 'Ersthelfer Berlin' }) as Record<string, unknown>;

      expect(mockGraphClient.searchDriveItems).toHaveBeenCalledWith({
        query: 'Ersthelfer Berlin',
        size: 10,
        sortBy: 'relevance',
      });
      expect(result['total']).toBe(2);

      const results = result['results'] as Record<string, unknown>[];
      expect(results).toHaveLength(2);
      expect(results[0]!['name']).toBe('Ersthelfer-Liste.docx');
      expect(results[0]!['drive_id']).toBe('drive-abc-123');
      expect(results[0]!['item_id']).toBe('item-search-1');
      expect(results[0]!['type']).toBe('file');
    });

    it('should clean up highlight tags and ddd tags from summary', async () => {
      const result = await spTools.searchFiles({ query: 'Ersthelfer Berlin' }) as Record<string, unknown>;
      const results = result['results'] as Record<string, unknown>[];
      // '<c0>Ersthelfer</c0> in <c0>Berlin</c0> sind...<ddd/>' → 'Ersthelfer in Berlin sind...…'
      expect(results[0]!['summary']).toBe('Die Ersthelfer in Berlin sind...…');
    });

    it('should include action field for anti-hallucination', async () => {
      const result = await spTools.searchFiles({ query: 'test' }) as Record<string, unknown>;
      const results = result['results'] as Record<string, unknown>[];

      expect(results[0]!['action']).toContain('sp_get_file');
      expect(results[0]!['action']).toContain('drive-abc-123');
      expect(results[0]!['action']).toContain('item-search-1');
    });

    it('should include result index numbers', async () => {
      const result = await spTools.searchFiles({ query: 'test' }) as Record<string, unknown>;
      const results = result['results'] as Record<string, unknown>[];

      expect(results[0]!['#']).toBe(1);
      expect(results[1]!['#']).toBe(2);
    });

    it('should include location from parentReference.path', async () => {
      const result = await spTools.searchFiles({ query: 'test' }) as Record<string, unknown>;
      const results = result['results'] as Record<string, unknown>[];

      expect(results[0]!['location']).toBe('/drives/drive-abc-123/root:/Documents');
    });

    it('should include warning note about using exact IDs', async () => {
      const result = await spTools.searchFiles({ query: 'test' }) as Record<string, unknown>;

      expect(result['_note']).toContain('EXACT drive_id and item_id');
      expect(result['_note']).toContain('Do NOT use IDs from earlier messages');
      expect(result['_note']).toContain('sp_search_read');
    });

    it('should use custom size parameter', async () => {
      await spTools.searchFiles({ query: 'test', size: 5 });

      expect(mockGraphClient.searchDriveItems).toHaveBeenCalledWith(
        expect.objectContaining({ size: 5 })
      );
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

    describe('with site_name scoping', () => {
      it('should prepend KQL path filter when site_name is provided', async () => {
        await spTools.searchFiles({ query: 'test', site_name: 'TeamSite' });

        expect(mockGraphClient.resolveSiteWebUrl).toHaveBeenCalledWith('TeamSite');
        expect(mockGraphClient.searchDriveItems).toHaveBeenCalledWith(
          expect.objectContaining({
            query: 'test path:"https://contoso.sharepoint.com/sites/TeamSite"',
          })
        );
      });

      it('should include site_filter in response when site is resolved', async () => {
        const result = await spTools.searchFiles({ query: 'test', site_name: 'TeamSite' }) as Record<string, unknown>;

        expect(result['site_filter']).toBe('https://contoso.sharepoint.com/sites/TeamSite');
      });

      it('should include site_warning when site is not found', async () => {
        (mockGraphClient.resolveSiteWebUrl as ReturnType<typeof vi.fn>).mockResolvedValue(null);

        const result = await spTools.searchFiles({ query: 'test', site_name: 'NonExistent' }) as Record<string, unknown>;

        expect(result['site_warning']).toContain('NonExistent');
        expect(result['site_warning']).toContain('not found');
        // Should still search (without scoping)
        expect(mockGraphClient.searchDriveItems).toHaveBeenCalledWith(
          expect.objectContaining({ query: 'test' })
        );
      });

      it('should work without site_name (backward compatible)', async () => {
        await spTools.searchFiles({ query: 'test' });

        expect(mockGraphClient.resolveSiteWebUrl).not.toHaveBeenCalled();
        expect(mockGraphClient.searchDriveItems).toHaveBeenCalledWith(
          expect.objectContaining({ query: 'test' })
        );
      });
    });

    describe('with sort parameter', () => {
      it('should pass sortBy to searchDriveItems', async () => {
        await spTools.searchFiles({ query: 'test', sort: 'lastModified' });

        expect(mockGraphClient.searchDriveItems).toHaveBeenCalledWith(
          expect.objectContaining({ sortBy: 'lastModified' })
        );
      });

      it('should default to relevance sort', async () => {
        await spTools.searchFiles({ query: 'test' });

        expect(mockGraphClient.searchDriveItems).toHaveBeenCalledWith(
          expect.objectContaining({ sortBy: 'relevance' })
        );
      });
    });
  });

  // ===== searchAndRead =====

  describe('searchAndRead', () => {
    it('should search and return file content in one call', async () => {
      const result = await spTools.searchAndRead({ query: 'Ersthelfer Berlin' }) as Record<string, unknown>;

      expect(result['found']).toBe(true);
      expect(result['name']).toBe('Ersthelfer-Liste.docx');
      expect(result['content']).toBe('file content here');
      expect(result['contentType']).toBe('text');
      expect(result['searchRank']).toBe(1);
      expect(result['totalResults']).toBe(2);
    });

    it('should use IDs directly from search result (no hallucination possible)', async () => {
      await spTools.searchAndRead({ query: 'test' });

      // getFileContent should be called with the EXACT IDs from the search hit
      expect(mockGraphClient.getFileContent).toHaveBeenCalledWith(
        'drive-abc-123',    // from mockSearchHits[0].resource.parentReference.driveId
        'item-search-1',     // from mockSearchHits[0].resource.id
        10 * 1024 * 1024
      );
    });

    it('should return found:false when no results', async () => {
      (mockGraphClient.searchDriveItems as ReturnType<typeof vi.fn>).mockResolvedValue({
        hits: [],
        total: 0,
        moreResultsAvailable: false,
      });

      const result = await spTools.searchAndRead({ query: 'nonexistent' }) as Record<string, unknown>;

      expect(result['found']).toBe(false);
      expect(result['_note']).toContain('No files matched');
    });

    it('should return error when result_index exceeds results', async () => {
      const result = await spTools.searchAndRead({ query: 'test', result_index: 10 }) as Record<string, unknown>;

      expect(result['found']).toBe(false);
      expect(result['availableResults']).toBe(2);
      expect(result['_note']).toContain('out of range');
    });

    it('should use second result when result_index is 1', async () => {
      const result = await spTools.searchAndRead({ query: 'test', result_index: 1 }) as Record<string, unknown>;

      expect(result['found']).toBe(true);
      expect(result['name']).toBe('Sicherheit.pdf');
      expect(result['searchRank']).toBe(2);

      expect(mockGraphClient.getFileContent).toHaveBeenCalledWith(
        'drive-abc-123',
        'item-search-2',
        10 * 1024 * 1024
      );
    });

    it('should handle site_name scoping', async () => {
      await spTools.searchAndRead({ query: 'test', site_name: 'TeamSite' });

      expect(mockGraphClient.resolveSiteWebUrl).toHaveBeenCalledWith('TeamSite');
      expect(mockGraphClient.searchDriveItems).toHaveBeenCalledWith(
        expect.objectContaining({
          query: 'test path:"https://contoso.sharepoint.com/sites/TeamSite"',
        })
      );
    });

    it('should handle folder results gracefully', async () => {
      (mockGraphClient.searchDriveItems as ReturnType<typeof vi.fn>).mockResolvedValue({
        hits: [{
          hitId: 'folder-hit',
          rank: 1,
          resource: {
            id: 'folder-1',
            name: 'MyFolder',
            webUrl: 'https://contoso.sharepoint.com/sites/TeamSite/Documents/MyFolder',
            folder: { childCount: 3 },
            parentReference: { driveId: 'drive-abc-123' },
          },
        }],
        total: 1,
        moreResultsAvailable: false,
      });

      const result = await spTools.searchAndRead({ query: 'folder' }) as Record<string, unknown>;

      expect(result['found']).toBe(true);
      expect(result['content']).toBeNull();
      expect(result['contentError']).toContain('folder, not a file');
    });

    it('should fetch content when search result has no file/folder facet (Graph Search API quirk)', async () => {
      // Graph Search API often omits the file and folder facets from driveItem results
      (mockGraphClient.searchDriveItems as ReturnType<typeof vi.fn>).mockResolvedValue({
        hits: [{
          hitId: 'nofacet-hit',
          rank: 1,
          resource: {
            id: 'item-no-facet',
            name: 'GBAusz_Neuss.pdf',
            size: 289220,
            webUrl: 'https://contoso.sharepoint.com/sites/Test/Documents/GBAusz_Neuss.pdf',
            lastModifiedDateTime: '2023-02-06T10:15:05Z',
            // NOTE: no file and no folder property — this is the real-world scenario
            parentReference: { driveId: 'drive-abc-123' },
          },
        }],
        total: 1,
        moreResultsAvailable: false,
      });

      const result = await spTools.searchAndRead({ query: 'GBAusz Neuss' }) as Record<string, unknown>;

      expect(result['found']).toBe(true);
      expect(result['name']).toBe('GBAusz_Neuss.pdf');
      // Should attempt to fetch content, not reject as folder
      expect(result['contentError'] ?? '').not.toContain('folder');
      // getFileContent should have been called with the IDs from the search hit
      expect(mockGraphClient.getFileContent).toHaveBeenCalledWith('drive-abc-123', 'item-no-facet', expect.any(Number));
    });

    it('should handle file too large gracefully', async () => {
      (mockGraphClient.searchDriveItems as ReturnType<typeof vi.fn>).mockResolvedValue({
        hits: [{
          hitId: 'big-hit',
          rank: 1,
          resource: {
            id: 'big-item',
            name: 'BigFile.zip',
            size: 20 * 1024 * 1024,
            file: { mimeType: 'application/zip' },
            parentReference: { driveId: 'drive-abc-123' },
          },
        }],
        total: 1,
        moreResultsAvailable: false,
      });

      const result = await spTools.searchAndRead({ query: 'big' }) as Record<string, unknown>;

      expect(result['found']).toBe(true);
      expect(result['content']).toBeNull();
      expect(result['contentError']).toContain('too large');
    });

    it('should handle content download failure gracefully', async () => {
      (mockGraphClient.getFileContent as ReturnType<typeof vi.fn>).mockRejectedValue(
        new Error('Network timeout')
      );

      const result = await spTools.searchAndRead({ query: 'test' }) as Record<string, unknown>;

      expect(result['found']).toBe(true);
      expect(result['name']).toBe('Ersthelfer-Liste.docx');
      expect(result['content']).toBeNull();
      expect(result['contentError']).toBe('Network timeout');
    });

    it('should clean up highlight and ddd tags from summary', async () => {
      const result = await spTools.searchAndRead({ query: 'test' }) as Record<string, unknown>;

      expect(result['summary']).toBe('Die Ersthelfer in Berlin sind...…');
    });

    it('should include location from parentReference.path', async () => {
      const result = await spTools.searchAndRead({ query: 'test' }) as Record<string, unknown>;

      expect(result['location']).toBe('/drives/drive-abc-123/root:/Documents');
    });

    it('should handle missing driveId in search result', async () => {
      (mockGraphClient.searchDriveItems as ReturnType<typeof vi.fn>).mockResolvedValue({
        hits: [{
          hitId: 'no-drive-hit',
          rank: 1,
          resource: {
            id: 'item-no-drive',
            name: 'Orphan.txt',
            file: { mimeType: 'text/plain' },
            parentReference: {},
          },
        }],
        total: 1,
        moreResultsAvailable: false,
      });

      const result = await spTools.searchAndRead({ query: 'orphan' }) as Record<string, unknown>;

      expect(result['found']).toBe(true);
      expect(result['content']).toBeNull();
      expect(result['contentError']).toContain('missing drive or item ID');
    });
  });
});

// ===== Schema Validation Tests =====

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

    it('should accept valid search with all optional params', () => {
      const result = searchFilesInputSchema.safeParse({
        query: 'Ersthelfer Berlin',
        site_name: 'IZ - Newsletter',
        sort: 'lastModified',
        size: 5,
      });
      expect(result.success).toBe(true);
    });

    it('should reject invalid sort value', () => {
      const result = searchFilesInputSchema.safeParse({
        query: 'test',
        sort: 'invalid',
      });
      expect(result.success).toBe(false);
    });

    it('should reject size > 25', () => {
      const result = searchFilesInputSchema.safeParse({
        query: 'test',
        size: 50,
      });
      expect(result.success).toBe(false);
    });
  });

  describe('searchAndReadInputSchema', () => {
    it('should require query', () => {
      const result = searchAndReadInputSchema.safeParse({});
      expect(result.success).toBe(false);
    });

    it('should accept valid input with all optional fields', () => {
      const result = searchAndReadInputSchema.safeParse({
        query: 'Ersthelfer Berlin',
        site_name: 'IZ - Newsletter',
        result_index: 2,
      });
      expect(result.success).toBe(true);
    });

    it('should reject result_index > 24', () => {
      const result = searchAndReadInputSchema.safeParse({
        query: 'test',
        result_index: 25,
      });
      expect(result.success).toBe(false);
    });

    it('should reject negative result_index', () => {
      const result = searchAndReadInputSchema.safeParse({
        query: 'test',
        result_index: -1,
      });
      expect(result.success).toBe(false);
    });

    it('should accept result_index 0', () => {
      const result = searchAndReadInputSchema.safeParse({
        query: 'test',
        result_index: 0,
      });
      expect(result.success).toBe(true);
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
