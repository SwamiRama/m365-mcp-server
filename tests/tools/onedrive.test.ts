import { describe, it, expect, vi, beforeEach } from 'vitest';
import {
  OneDriveTools,
  myDriveInputSchema,
  listFilesInputSchema,
  getFileInputSchema,
  searchInputSchema,
  recentInputSchema,
  sharedWithMeInputSchema,
} from '../../src/tools/onedrive.js';
import type { GraphClient, GraphDrive, GraphDriveItem } from '../../src/graph/client.js';

describe('OneDrive Tools', () => {
  let mockGraphClient: GraphClient;
  let odTools: OneDriveTools;

  const mockDrive: GraphDrive = {
    id: 'b!drive-personal-123',
    name: 'OneDrive',
    driveType: 'personal',
    webUrl: 'https://contoso-my.sharepoint.com/personal/user_contoso_com/Documents',
    owner: { user: { displayName: 'Test User' } },
    quota: {
      total: 1099511627776,
      used: 536870912,
      remaining: 1098974756864,
      state: 'normal',
    },
  };

  const mockDriveItems: GraphDriveItem[] = [
    {
      id: 'od-item-1',
      name: 'Notes.txt',
      webUrl: 'https://contoso-my.sharepoint.com/personal/user/Documents/Notes.txt',
      size: 1024,
      createdDateTime: '2026-01-10T08:00:00Z',
      lastModifiedDateTime: '2026-02-01T14:00:00Z',
      file: { mimeType: 'text/plain' },
      parentReference: { driveId: 'b!drive-personal-123', path: '/drive/root:' },
    },
    {
      id: 'od-item-2',
      name: 'Projects',
      webUrl: 'https://contoso-my.sharepoint.com/personal/user/Documents/Projects',
      createdDateTime: '2025-06-01T08:00:00Z',
      lastModifiedDateTime: '2026-01-20T16:00:00Z',
      folder: { childCount: 3 },
      parentReference: { driveId: 'b!drive-personal-123', path: '/drive/root:' },
    },
  ];

  beforeEach(() => {
    mockGraphClient = {
      getMyDrive: vi.fn().mockResolvedValue(mockDrive),
      listDriveItems: vi.fn().mockResolvedValue(mockDriveItems),
      getDriveItem: vi.fn().mockResolvedValue(mockDriveItems[0]),
      getFileContent: vi.fn().mockResolvedValue({
        content: Buffer.from('hello world'),
        mimeType: 'text/plain',
        size: 11,
      }),
      searchMyDrive: vi.fn().mockResolvedValue(mockDriveItems),
      getMyDriveRecent: vi.fn().mockResolvedValue(mockDriveItems),
      getMyDriveSharedWithMe: vi.fn().mockResolvedValue(mockDriveItems),
    } as unknown as GraphClient;

    odTools = new OneDriveTools(mockGraphClient);
  });

  // ===== myDrive =====

  describe('myDrive', () => {
    it('should return drive info with quota', async () => {
      const result = await odTools.myDrive() as Record<string, unknown>;

      expect(mockGraphClient.getMyDrive).toHaveBeenCalled();
      expect(result['id']).toBe('b!drive-personal-123');
      expect(result['name']).toBe('OneDrive');
      expect(result['type']).toBe('personal');
      expect(result['owner']).toBe('Test User');
      expect(result['_note']).toContain('sp_list_children');

      const quota = result['quota'] as Record<string, unknown>;
      expect(quota['state']).toBe('normal');
      expect(quota['total']).toBe('1024.0 GB');
      expect(quota['used']).toBe('512.0 MB');
    });

    it('should handle missing quota gracefully', async () => {
      const driveWithoutQuota = { ...mockDrive, quota: undefined };
      (mockGraphClient.getMyDrive as ReturnType<typeof vi.fn>).mockResolvedValue(driveWithoutQuota);

      const result = await odTools.myDrive() as Record<string, unknown>;

      expect(result['id']).toBe('b!drive-personal-123');
      expect(result['quota']).toBeUndefined();
    });
  });

  // ===== listFiles =====

  describe('listFiles', () => {
    it('should list root folder when no item_id provided', async () => {
      const result = await odTools.listFiles({}) as Record<string, unknown>;

      expect(mockGraphClient.getMyDrive).toHaveBeenCalled();
      expect(mockGraphClient.listDriveItems).toHaveBeenCalledWith({
        driveId: 'b!drive-personal-123',
        itemId: undefined,
        top: 50,
      });
      expect(result['count']).toBe(2);
      expect(result['item_id']).toBe('root');

      const items = result['items'] as Record<string, unknown>[];
      expect(items[0]!['name']).toBe('Notes.txt');
      expect(items[0]!['type']).toBe('file');
      expect(items[1]!['name']).toBe('Projects');
      expect(items[1]!['type']).toBe('folder');
    });

    it('should list subfolder contents', async () => {
      await odTools.listFiles({ item_id: 'od-item-2' });

      expect(mockGraphClient.listDriveItems).toHaveBeenCalledWith({
        driveId: 'b!drive-personal-123',
        itemId: 'od-item-2',
        top: 50,
      });
    });

    it('should respect top parameter', async () => {
      await odTools.listFiles({ top: 10 });

      expect(mockGraphClient.listDriveItems).toHaveBeenCalledWith(
        expect.objectContaining({ top: 10 })
      );
    });

    it('should throw enriched error for 404', async () => {
      const notFoundError = Object.assign(new Error('Not found'), {
        code: 'itemNotFound',
        statusCode: 404,
      });
      (mockGraphClient.listDriveItems as ReturnType<typeof vi.fn>).mockRejectedValue(notFoundError);

      await expect(
        odTools.listFiles({ item_id: 'bad-id' })
      ).rejects.toThrow(/Item not found.*od_list_files without item_id/);
    });
  });

  // ===== getFile =====

  describe('getFile', () => {
    it('should get file with text content', async () => {
      const result = await odTools.getFile({ item_id: 'od-item-1' }) as Record<string, unknown>;

      expect(mockGraphClient.getDriveItem).toHaveBeenCalledWith('b!drive-personal-123', 'od-item-1');
      expect(mockGraphClient.getFileContent).toHaveBeenCalledWith(
        'b!drive-personal-123',
        'od-item-1',
        10 * 1024 * 1024
      );
      expect(result['name']).toBe('Notes.txt');
      expect(result['content']).toBe('hello world');
      expect(result['contentType']).toBe('text');
    });

    it('should throw enriched error for 404', async () => {
      const notFoundError = Object.assign(new Error('Not found'), {
        code: 'itemNotFound',
        statusCode: 404,
      });
      (mockGraphClient.getDriveItem as ReturnType<typeof vi.fn>).mockRejectedValue(notFoundError);

      await expect(
        odTools.getFile({ item_id: 'bad-id' })
      ).rejects.toThrow(/File not found.*od_list_files/);
    });

    it('should reject folders', async () => {
      (mockGraphClient.getDriveItem as ReturnType<typeof vi.fn>).mockResolvedValue(mockDriveItems[1]);

      await expect(
        odTools.getFile({ item_id: 'od-item-2' })
      ).rejects.toThrow(/folder, not a file.*od_list_files/);
    });

    it('should handle files that are too large', async () => {
      (mockGraphClient.getDriveItem as ReturnType<typeof vi.fn>).mockResolvedValue({
        ...mockDriveItems[0],
        size: 20 * 1024 * 1024,
      });

      const result = await odTools.getFile({ item_id: 'od-item-1' }) as Record<string, unknown>;

      expect(result['content']).toBeNull();
      expect(result['contentError']).toContain('too large');
    });
  });

  // ===== search =====

  describe('search', () => {
    it('should search and return formatted results', async () => {
      const result = await odTools.search({ query: 'Notes' }) as Record<string, unknown>;

      expect(mockGraphClient.searchMyDrive).toHaveBeenCalledWith('Notes', 25);
      expect(result['count']).toBe(2);
      expect(result['query']).toBe('Notes');

      const results = result['results'] as Record<string, unknown>[];
      expect(results[0]!['#']).toBe(1);
      expect(results[0]!['name']).toBe('Notes.txt');
      expect(results[0]!['action']).toContain('od_get_file');
      expect(results[0]!['action']).toContain('od-item-1');
    });

    it('should not include action for folders', async () => {
      const result = await odTools.search({ query: 'Projects' }) as Record<string, unknown>;
      const results = result['results'] as Record<string, unknown>[];

      // Second item is a folder — no action
      expect(results[1]!['action']).toBeUndefined();
    });

    it('should handle empty results', async () => {
      (mockGraphClient.searchMyDrive as ReturnType<typeof vi.fn>).mockResolvedValue([]);

      const result = await odTools.search({ query: 'nonexistent' }) as Record<string, unknown>;

      expect(result['count']).toBe(0);
      expect(result['results']).toEqual([]);
    });

    it('should pass top parameter', async () => {
      await odTools.search({ query: 'test', top: 5 });

      expect(mockGraphClient.searchMyDrive).toHaveBeenCalledWith('test', 5);
    });
  });

  // ===== recent =====

  describe('recent', () => {
    it('should return recent items', async () => {
      const result = await odTools.recent({}) as Record<string, unknown>;

      expect(mockGraphClient.getMyDriveRecent).toHaveBeenCalledWith(25);
      expect(result['count']).toBe(2);

      const items = result['items'] as Record<string, unknown>[];
      expect(items[0]!['name']).toBe('Notes.txt');
    });

    it('should handle empty list', async () => {
      (mockGraphClient.getMyDriveRecent as ReturnType<typeof vi.fn>).mockResolvedValue([]);

      const result = await odTools.recent({}) as Record<string, unknown>;
      expect(result['count']).toBe(0);
    });

    it('should pass top parameter', async () => {
      await odTools.recent({ top: 10 });

      expect(mockGraphClient.getMyDriveRecent).toHaveBeenCalledWith(10);
    });
  });

  // ===== sharedWithMe =====

  describe('sharedWithMe', () => {
    it('should return shared items with cross-drive note', async () => {
      const result = await odTools.sharedWithMe({}) as Record<string, unknown>;

      expect(mockGraphClient.getMyDriveSharedWithMe).toHaveBeenCalledWith(25);
      expect(result['count']).toBe(2);
      expect(result['_note']).toContain('other drives');
    });

    it('should handle empty list', async () => {
      (mockGraphClient.getMyDriveSharedWithMe as ReturnType<typeof vi.fn>).mockResolvedValue([]);

      const result = await odTools.sharedWithMe({}) as Record<string, unknown>;
      expect(result['count']).toBe(0);
    });

    it('should pass top parameter', async () => {
      await odTools.sharedWithMe({ top: 10 });

      expect(mockGraphClient.getMyDriveSharedWithMe).toHaveBeenCalledWith(10);
    });
  });

  // ===== Drive ID caching =====

  describe('drive ID caching', () => {
    it('should call getMyDrive only once across multiple method calls', async () => {
      await odTools.listFiles({});
      await odTools.getFile({ item_id: 'od-item-1' });

      expect(mockGraphClient.getMyDrive).toHaveBeenCalledTimes(1);
    });

    it('should cache drive ID from myDrive call', async () => {
      await odTools.myDrive();
      await odTools.listFiles({});

      // myDrive calls getMyDrive, then listFiles should use cached value
      expect(mockGraphClient.getMyDrive).toHaveBeenCalledTimes(1);
    });
  });
});

// ===== Schema Validation Tests =====

describe('OneDrive Input Schemas', () => {
  describe('myDriveInputSchema', () => {
    it('should accept empty object', () => {
      expect(myDriveInputSchema.safeParse({}).success).toBe(true);
    });
  });

  describe('listFilesInputSchema', () => {
    it('should accept empty object (root listing)', () => {
      expect(listFilesInputSchema.safeParse({}).success).toBe(true);
    });

    it('should accept valid item_id', () => {
      const result = listFilesInputSchema.safeParse({ item_id: 'od-item-2' });
      expect(result.success).toBe(true);
    });

    it('should reject invalid item_id format', () => {
      const result = listFilesInputSchema.safeParse({ item_id: 'invalid<>id' });
      expect(result.success).toBe(false);
    });

    it('should reject top > 200', () => {
      const result = listFilesInputSchema.safeParse({ top: 300 });
      expect(result.success).toBe(false);
    });

    it('should reject top < 1', () => {
      const result = listFilesInputSchema.safeParse({ top: 0 });
      expect(result.success).toBe(false);
    });
  });

  describe('getFileInputSchema', () => {
    it('should require item_id', () => {
      expect(getFileInputSchema.safeParse({}).success).toBe(false);
    });

    it('should accept valid item_id', () => {
      expect(getFileInputSchema.safeParse({ item_id: 'b!abc-123' }).success).toBe(true);
    });

    it('should reject invalid item_id format', () => {
      expect(getFileInputSchema.safeParse({ item_id: '' }).success).toBe(false);
      expect(getFileInputSchema.safeParse({ item_id: 'a<b>c' }).success).toBe(false);
    });
  });

  describe('searchInputSchema', () => {
    it('should require query', () => {
      expect(searchInputSchema.safeParse({}).success).toBe(false);
    });

    it('should reject empty query', () => {
      expect(searchInputSchema.safeParse({ query: '' }).success).toBe(false);
    });

    it('should accept valid search', () => {
      expect(searchInputSchema.safeParse({ query: 'test', top: 10 }).success).toBe(true);
    });

    it('should reject top > 50', () => {
      expect(searchInputSchema.safeParse({ query: 'test', top: 100 }).success).toBe(false);
    });
  });

  describe('recentInputSchema', () => {
    it('should accept empty object', () => {
      expect(recentInputSchema.safeParse({}).success).toBe(true);
    });

    it('should reject top > 50', () => {
      expect(recentInputSchema.safeParse({ top: 100 }).success).toBe(false);
    });
  });

  describe('sharedWithMeInputSchema', () => {
    it('should accept empty object', () => {
      expect(sharedWithMeInputSchema.safeParse({}).success).toBe(true);
    });

    it('should reject top > 50', () => {
      expect(sharedWithMeInputSchema.safeParse({ top: 100 }).success).toBe(false);
    });
  });
});
