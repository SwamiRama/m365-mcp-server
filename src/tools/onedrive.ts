import { z } from 'zod';
import { GraphClient, type GraphDrive, type GraphDriveItem } from '../graph/client.js';
import { logger } from '../utils/logger.js';
import { graphIdPattern, graphIdSchema } from '../utils/graph-id.js';
import { formatFileSize, fetchAndParseContent } from '../utils/content-fetcher.js';

// Maximum file size for content retrieval (10MB)
const MAX_FILE_SIZE = 10 * 1024 * 1024;

// Input schemas
export const myDriveInputSchema = z.object({});

export const listFilesInputSchema = z.object({
  item_id: z
    .string()
    .regex(graphIdPattern, 'Invalid item ID format')
    .optional()
    .describe('Folder ID to list contents of. Omit to list root folder.'),
  top: z
    .number()
    .int()
    .min(1)
    .max(200)
    .optional()
    .default(50)
    .describe('Maximum number of items to return (1-200)'),
});

export const getFileInputSchema = z.object({
  item_id: graphIdSchema.describe('The ID of the file to retrieve from your OneDrive'),
});

export const searchInputSchema = z.object({
  query: z
    .string()
    .min(1)
    .max(512)
    .describe('Search query for files in your personal OneDrive'),
  top: z
    .number()
    .int()
    .min(1)
    .max(50)
    .optional()
    .default(25)
    .describe('Maximum number of results (1-50)'),
});

export const recentInputSchema = z.object({
  top: z
    .number()
    .int()
    .min(1)
    .max(50)
    .optional()
    .default(25)
    .describe('Maximum number of recent items (1-50)'),
});

export const sharedWithMeInputSchema = z.object({
  top: z
    .number()
    .int()
    .min(1)
    .max(50)
    .optional()
    .default(25)
    .describe('Maximum number of shared items (1-50)'),
});

export type MyDriveInput = z.infer<typeof myDriveInputSchema>;
export type ListFilesInput = z.infer<typeof listFilesInputSchema>;
export type OdGetFileInput = z.infer<typeof getFileInputSchema>;
export type SearchInput = z.infer<typeof searchInputSchema>;
export type RecentInput = z.infer<typeof recentInputSchema>;
export type SharedWithMeInput = z.infer<typeof sharedWithMeInputSchema>;

// Output formatters
function formatDriveInfo(drive: GraphDrive): Record<string, unknown> {
  const formatted: Record<string, unknown> = {
    id: drive.id,
    name: drive.name,
    type: drive.driveType,
    webUrl: drive.webUrl,
    owner: drive.owner?.user?.displayName,
  };

  if (drive.quota) {
    formatted['quota'] = {
      total: drive.quota.total !== null && drive.quota.total !== undefined ? formatFileSize(drive.quota.total) : undefined,
      used: drive.quota.used !== null && drive.quota.used !== undefined ? formatFileSize(drive.quota.used) : undefined,
      remaining: drive.quota.remaining !== null && drive.quota.remaining !== undefined ? formatFileSize(drive.quota.remaining) : undefined,
      state: drive.quota.state,
    };
  }

  return formatted;
}

function formatDriveItem(item: GraphDriveItem): Record<string, unknown> {
  const formatted: Record<string, unknown> = {
    id: item.id,
    name: item.name,
    webUrl: item.webUrl,
    lastModified: item.lastModifiedDateTime,
    created: item.createdDateTime,
  };

  if (item.file) {
    formatted['type'] = 'file';
    formatted['mimeType'] = item.file.mimeType;
    formatted['size'] = item.size;
  } else if (item.folder) {
    formatted['type'] = 'folder';
    formatted['childCount'] = item.folder.childCount;
  }

  if (item.parentReference?.path) {
    formatted['path'] = item.parentReference.path;
  }

  return formatted;
}

// Tool implementations
export class OneDriveTools {
  private graphClient: GraphClient;
  private cachedDriveId: string | null = null;

  constructor(graphClient: GraphClient) {
    this.graphClient = graphClient;
  }

  private async getDriveId(): Promise<string> {
    if (!this.cachedDriveId) {
      const drive = await this.graphClient.getMyDrive();
      this.cachedDriveId = drive.id;
    }
    return this.cachedDriveId;
  }

  async myDrive(): Promise<object> {
    logger.debug('Getting personal OneDrive info');

    const drive = await this.graphClient.getMyDrive();
    this.cachedDriveId = drive.id;

    return {
      ...formatDriveInfo(drive),
      _note:
        'This drive_id can also be used with sp_list_children and sp_get_file for advanced operations.',
    };
  }

  async listFiles(input: ListFilesInput): Promise<object> {
    const validated = listFilesInputSchema.parse(input);
    const driveId = await this.getDriveId();

    logger.debug({ driveId, itemId: validated.item_id, top: validated.top }, 'Listing OneDrive files');

    try {
      const items = await this.graphClient.listDriveItems({
        driveId,
        itemId: validated.item_id,
        top: validated.top,
      });

      return {
        items: items.map(formatDriveItem),
        count: items.length,
        item_id: validated.item_id ?? 'root',
      };
    } catch (err) {
      const code = (err as { code?: string }).code;
      const statusCode = (err as { statusCode?: number }).statusCode;

      if (code === 'itemNotFound' || (statusCode === 404 && !code)) {
        const target = validated.item_id
          ? `folder '${validated.item_id}'`
          : 'root folder';
        const hint = `Item not found: ${target} in your OneDrive. The item_id may be stale. Remediation: call od_list_files without item_id to browse from the root.`;

        const enrichedError = new Error(hint) as Error & { code?: string; statusCode?: number };
        enrichedError.code = code ?? 'itemNotFound';
        enrichedError.statusCode = statusCode ?? 404;
        throw enrichedError;
      }
      throw err;
    }
  }

  async getFile(input: OdGetFileInput): Promise<object> {
    const validated = getFileInputSchema.parse(input);
    const driveId = await this.getDriveId();

    logger.debug({ driveId, itemId: validated.item_id }, 'Getting OneDrive file');

    let item;
    try {
      item = await this.graphClient.getDriveItem(driveId, validated.item_id);
    } catch (err) {
      const code = (err as { code?: string }).code;
      const statusCode = (err as { statusCode?: number }).statusCode;
      if (code === 'itemNotFound' || (statusCode === 404 && !code)) {
        const hint = `File not found: item '${validated.item_id}' in your OneDrive. The ID may be stale. Remediation: call od_list_files to get fresh item IDs, then retry od_get_file.`;
        const enrichedError = new Error(hint) as Error & { code?: string; statusCode?: number };
        enrichedError.code = code ?? 'itemNotFound';
        enrichedError.statusCode = statusCode ?? 404;
        throw enrichedError;
      }
      throw err;
    }

    if (!item.file) {
      throw new Error(
        'The specified item is a folder, not a file. Use od_list_files with this item_id to browse its contents.'
      );
    }

    const result: Record<string, unknown> = formatDriveItem(item);

    const fileSize = item.size ?? 0;
    if (fileSize > MAX_FILE_SIZE) {
      result['content'] = null;
      result['contentError'] = `File too large (${formatFileSize(fileSize)}). Maximum size: ${formatFileSize(MAX_FILE_SIZE)}`;
      return result;
    }

    await fetchAndParseContent(
      this.graphClient,
      driveId,
      validated.item_id,
      item.name ?? 'unknown',
      MAX_FILE_SIZE,
      result
    );

    return result;
  }

  async search(input: SearchInput): Promise<object> {
    const validated = searchInputSchema.parse(input);

    logger.debug({ query: validated.query, top: validated.top }, 'Searching OneDrive');

    const items = await this.graphClient.searchMyDrive(validated.query, validated.top);

    return {
      results: items.map((item, index) => {
        const formatted = formatDriveItem(item);
        formatted['#'] = index + 1;
        if (item.id && !item.folder) {
          formatted['action'] = `To read this file: od_get_file(item_id="${item.id}")`;
        }
        return formatted;
      }),
      count: items.length,
      query: validated.query,
    };
  }

  async recent(input: RecentInput): Promise<object> {
    const validated = recentInputSchema.parse(input);

    logger.debug({ top: validated.top }, 'Getting recent OneDrive files');

    const items = await this.graphClient.getMyDriveRecent(validated.top);

    return {
      items: items.map(formatDriveItem),
      count: items.length,
    };
  }

  async sharedWithMe(input: SharedWithMeInput): Promise<object> {
    const validated = sharedWithMeInputSchema.parse(input);

    logger.debug({ top: validated.top }, 'Getting files shared with me');

    const items = await this.graphClient.getMyDriveSharedWithMe(validated.top);

    return {
      items: items.map(formatDriveItem),
      count: items.length,
      _note:
        'Shared items may reside on other drives. Use sp_get_file with drive_id from parentReference if you need to read a shared file from another drive.',
    };
  }
}

// Tool definitions for MCP registration
export const oneDriveToolDefinitions = [
  {
    name: 'od_my_drive',
    description:
      'Get your personal OneDrive info including drive ID and storage quota. ' +
      'Use this first to see available space or to get the drive_id for cross-referencing with sp_* tools.',
    inputSchema: {
      type: 'object' as const,
      properties: {},
    },
  },
  {
    name: 'od_list_files',
    description:
      'List files and folders in your personal OneDrive. Omit item_id for root folder, or provide a folder ID to browse subfolders. ' +
      'For SharePoint document libraries, use sp_list_children instead.',
    inputSchema: {
      type: 'object' as const,
      properties: {
        item_id: {
          type: 'string',
          description: 'Folder ID to list contents of. Omit to list root folder.',
        },
        top: {
          type: 'number',
          description: 'Maximum number of items to return (1-200). Default: 50',
          minimum: 1,
          maximum: 200,
        },
      },
    },
  },
  {
    name: 'od_get_file',
    description:
      'Get a file from your personal OneDrive by item_id. PDF, Word (.docx), Excel (.xlsx), PowerPoint (.pptx) are auto-parsed to readable text. Max 10MB. ' +
      'The item_id MUST come from a recent od_list_files, od_search, or od_recent response.',
    inputSchema: {
      type: 'object' as const,
      properties: {
        item_id: {
          type: 'string',
          description: 'The EXACT file ID from a recent od_list_files, od_search, or od_recent response.',
        },
      },
      required: ['item_id'],
    },
  },
  {
    name: 'od_search',
    description:
      'Search for files in your personal OneDrive ONLY. For searching across SharePoint sites and all drives, use sp_search instead. ' +
      'Returns file metadata and IDs. Use od_get_file to read file content.',
    inputSchema: {
      type: 'object' as const,
      properties: {
        query: {
          type: 'string',
          description: 'Search query for files in your personal OneDrive.',
        },
        top: {
          type: 'number',
          description: 'Maximum number of results (1-50). Default: 25',
          minimum: 1,
          maximum: 50,
        },
      },
      required: ['query'],
    },
  },
  {
    name: 'od_recent',
    description:
      'List recently accessed files in your personal OneDrive. Useful for finding files you worked on recently.',
    inputSchema: {
      type: 'object' as const,
      properties: {
        top: {
          type: 'number',
          description: 'Maximum number of recent items (1-50). Default: 25',
          minimum: 1,
          maximum: 50,
        },
      },
    },
  },
  {
    name: 'od_shared_with_me',
    description:
      'List files that others have shared with you via OneDrive. ' +
      'Note: shared items may reside on other users\' drives.',
    inputSchema: {
      type: 'object' as const,
      properties: {
        top: {
          type: 'number',
          description: 'Maximum number of shared items (1-50). Default: 25',
          minimum: 1,
          maximum: 50,
        },
      },
    },
  },
];
