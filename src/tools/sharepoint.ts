import { z } from 'zod';
import {
  GraphClient,
  type GraphSite,
  type GraphDrive,
  type GraphDriveItem,
} from '../graph/client.js';
import { logger } from '../utils/logger.js';

// Input schemas for SharePoint/Files tools
export const listSitesInputSchema = z.object({
  query: z
    .string()
    .optional()
    .describe('Search query to filter sites by name'),
});

export const listDrivesInputSchema = z.object({
  site_id: z
    .string()
    .optional()
    .describe(
      "Site ID to list drives from. If not provided, lists the user's personal OneDrive."
    ),
});

export const listChildrenInputSchema = z.object({
  drive_id: z
    .string()
    .min(1)
    .describe('The ID of the drive to list items from'),
  item_id: z
    .string()
    .optional()
    .describe(
      'The ID of the folder to list children from. If not provided, lists root folder contents.'
    ),
});

export const getFileInputSchema = z.object({
  drive_id: z.string().min(1).describe('The ID of the drive containing the file'),
  item_id: z.string().min(1).describe('The ID of the file to retrieve'),
});

export type ListSitesInput = z.infer<typeof listSitesInputSchema>;
export type ListDrivesInput = z.infer<typeof listDrivesInputSchema>;
export type ListChildrenInput = z.infer<typeof listChildrenInputSchema>;
export type GetFileInput = z.infer<typeof getFileInputSchema>;

// Maximum file size for content retrieval (10MB)
const MAX_FILE_SIZE = 10 * 1024 * 1024;

// Text-based MIME types that can be safely returned as text
const TEXT_MIME_TYPES = new Set([
  'text/plain',
  'text/html',
  'text/css',
  'text/javascript',
  'text/csv',
  'text/xml',
  'text/markdown',
  'application/json',
  'application/xml',
  'application/javascript',
  'application/x-yaml',
  'application/x-sh',
]);

function isTextMimeType(mimeType: string): boolean {
  if (TEXT_MIME_TYPES.has(mimeType)) return true;
  if (mimeType.startsWith('text/')) return true;
  if (mimeType.endsWith('+json')) return true;
  if (mimeType.endsWith('+xml')) return true;
  return false;
}

// Output formatters
function formatSite(site: GraphSite): object {
  return {
    id: site.id,
    name: site.name ?? site.displayName,
    webUrl: site.webUrl,
    description: site.description,
  };
}

function formatDrive(drive: GraphDrive): object {
  return {
    id: drive.id,
    name: drive.name,
    type: drive.driveType,
    webUrl: drive.webUrl,
    owner: drive.owner?.user?.displayName,
  };
}

function formatDriveItem(item: GraphDriveItem): object {
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

  return formatted;
}

function formatFileSize(bytes: number): string {
  if (bytes < 1024) return `${bytes} B`;
  if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(1)} KB`;
  if (bytes < 1024 * 1024 * 1024) return `${(bytes / (1024 * 1024)).toFixed(1)} MB`;
  return `${(bytes / (1024 * 1024 * 1024)).toFixed(1)} GB`;
}

// Tool implementations
export class SharePointTools {
  constructor(private graphClient: GraphClient) {}

  /**
   * List accessible SharePoint sites
   */
  async listSites(input: ListSitesInput): Promise<object> {
    const validated = listSitesInputSchema.parse(input);

    logger.debug({ query: validated.query }, 'Listing sites');

    const sites = await this.graphClient.listSites({
      search: validated.query,
    });

    return {
      sites: sites.map(formatSite),
      count: sites.length,
    };
  }

  /**
   * List drives (OneDrive/SharePoint document libraries)
   */
  async listDrives(input: ListDrivesInput): Promise<object> {
    const validated = listDrivesInputSchema.parse(input);

    logger.debug({ siteId: validated.site_id }, 'Listing drives');

    if (!validated.site_id) {
      // Return user's personal OneDrive
      const drive = await this.graphClient.getMyDrive();
      return {
        drives: [formatDrive(drive)],
        count: 1,
      };
    }

    const drives = await this.graphClient.listDrives(validated.site_id);

    return {
      drives: drives.map(formatDrive),
      count: drives.length,
    };
  }

  /**
   * List children items in a drive folder
   */
  async listChildren(input: ListChildrenInput): Promise<object> {
    const validated = listChildrenInputSchema.parse(input);

    logger.debug(
      { driveId: validated.drive_id, itemId: validated.item_id },
      'Listing drive items'
    );

    const items = await this.graphClient.listDriveItems({
      driveId: validated.drive_id,
      itemId: validated.item_id,
    });

    return {
      items: items.map(formatDriveItem),
      count: items.length,
    };
  }

  /**
   * Get file metadata and optionally content
   */
  async getFile(input: GetFileInput): Promise<object> {
    const validated = getFileInputSchema.parse(input);

    logger.debug(
      { driveId: validated.drive_id, itemId: validated.item_id },
      'Getting file'
    );

    // Get file metadata first
    const item = await this.graphClient.getDriveItem(
      validated.drive_id,
      validated.item_id
    );

    if (!item.file) {
      throw new Error('The specified item is not a file');
    }

    const result: Record<string, unknown> = formatDriveItem(item) as Record<string, unknown>;

    // Check if file is small enough to fetch content
    const fileSize = item.size ?? 0;
    if (fileSize > MAX_FILE_SIZE) {
      result['content'] = null;
      result['contentError'] = `File too large (${formatFileSize(fileSize)}). Maximum size: ${formatFileSize(MAX_FILE_SIZE)}`;
      return result;
    }

    // Fetch content
    try {
      const contentResult = await this.graphClient.getFileContent(
        validated.drive_id,
        validated.item_id,
        MAX_FILE_SIZE
      );

      if (contentResult) {
        const mimeType = contentResult.mimeType;

        if (isTextMimeType(mimeType)) {
          // Return text content directly
          result['content'] = contentResult.content.toString('utf-8');
          result['contentType'] = 'text';
        } else {
          // Return base64 encoded binary content
          result['content'] = contentResult.content.toString('base64');
          result['contentType'] = 'base64';
        }

        result['contentMimeType'] = mimeType;
        result['contentSize'] = contentResult.size;
      }
    } catch (err) {
      const errorMessage = err instanceof Error ? err.message : 'Unknown error';
      logger.warn(
        { err, driveId: validated.drive_id, itemId: validated.item_id },
        'Failed to fetch file content'
      );
      result['content'] = null;
      result['contentError'] = errorMessage;
    }

    return result;
  }
}

// Tool definitions for MCP registration
export const sharePointToolDefinitions = [
  {
    name: 'sp_list_sites',
    description:
      'Search and list SharePoint sites accessible to the user. Use the site ID to list drives within a site.',
    inputSchema: {
      type: 'object' as const,
      properties: {
        query: {
          type: 'string',
          description: 'Search query to filter sites by name',
        },
      },
    },
  },
  {
    name: 'sp_list_drives',
    description:
      "List document libraries/drives. Without a site_id, returns the user's personal OneDrive.",
    inputSchema: {
      type: 'object' as const,
      properties: {
        site_id: {
          type: 'string',
          description:
            "Site ID to list drives from. If not provided, lists the user's personal OneDrive.",
        },
      },
    },
  },
  {
    name: 'sp_list_children',
    description:
      'List files and folders in a drive. Provide item_id to list contents of a specific folder, or omit for root.',
    inputSchema: {
      type: 'object' as const,
      properties: {
        drive_id: {
          type: 'string',
          description: 'The ID of the drive to list items from',
        },
        item_id: {
          type: 'string',
          description:
            'The ID of the folder to list children from. If not provided, lists root folder contents.',
        },
      },
      required: ['drive_id'],
    },
  },
  {
    name: 'sp_get_file',
    description:
      'Get file metadata and content. Text files are returned as text, binary files as base64. Maximum size: 10MB.',
    inputSchema: {
      type: 'object' as const,
      properties: {
        drive_id: {
          type: 'string',
          description: 'The ID of the drive containing the file',
        },
        item_id: {
          type: 'string',
          description: 'The ID of the file to retrieve',
        },
      },
      required: ['drive_id', 'item_id'],
    },
  },
];
