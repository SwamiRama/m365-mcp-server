import { z } from 'zod';
import {
  GraphClient,
  type GraphSite,
  type GraphDrive,
  type GraphDriveItem,
} from '../graph/client.js';
import { logger } from '../utils/logger.js';
import { isParsableMimeType, parseFileContent } from '../utils/file-parser.js';

// Safe pattern for Graph API resource IDs (alphanumeric, hyphens, dots, underscores, commas, colons, exclamation marks)
const graphIdPattern = /^[a-zA-Z0-9\-._,!:]+$/;
const graphIdSchema = z.string().min(1).regex(graphIdPattern, 'Invalid resource ID format');

// Input schemas for SharePoint/Files tools
export const listSitesInputSchema = z.object({
  query: z
    .string()
    .max(256)
    .optional()
    .describe('Search query to filter sites by name'),
});

export const listDrivesInputSchema = z.object({
  site_id: z
    .string()
    .regex(graphIdPattern, 'Invalid site ID format')
    .optional()
    .describe(
      "Site ID to list drives from. If not provided, lists the user's personal OneDrive."
    ),
});

export const listChildrenInputSchema = z.object({
  drive_id: graphIdSchema
    .describe('The ID of the drive to list items from'),
  item_id: z
    .string()
    .regex(graphIdPattern, 'Invalid item ID format')
    .optional()
    .describe(
      'The ID of the folder to list children from. If not provided, lists root folder contents.'
    ),
});

export const getFileInputSchema = z.object({
  drive_id: graphIdSchema.describe('The ID of the drive containing the file'),
  item_id: graphIdSchema.describe('The ID of the file to retrieve'),
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
        source: 'personal_onedrive',
        _note: `This is your personal OneDrive. Use drive_id='${drive.id}' with sp_list_children to browse files.`,
      };
    }

    const drives = await this.graphClient.listDrives(validated.site_id);

    return {
      drives: drives.map(formatDrive),
      count: drives.length,
      source: 'sharepoint_site',
      site_id: validated.site_id,
      _note: 'Use the drive id values from this response with sp_list_children to browse files.',
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

    try {
      const items = await this.graphClient.listDriveItems({
        driveId: validated.drive_id,
        itemId: validated.item_id,
      });

      return {
        items: items.map(formatDriveItem),
        count: items.length,
        drive_id: validated.drive_id,
        item_id: validated.item_id ?? 'root',
      };
    } catch (err) {
      const code = (err as { code?: string }).code;
      const statusCode = (err as { statusCode?: number }).statusCode;

      if (code === 'itemNotFound' || (statusCode === 404 && !code)) {
        const target = validated.item_id
          ? `item '${validated.item_id}' in drive '${validated.drive_id}'`
          : `root of drive '${validated.drive_id}'`;
        const hint = `Item not found: ${target}. The drive_id or item_id may be stale or from a previous session. Remediation: call sp_list_drives to get a fresh drive_id, then call sp_list_children with the new drive_id.`;

        const enrichedError = new Error(hint) as Error & { code?: string; statusCode?: number };
        enrichedError.code = code ?? 'itemNotFound';
        enrichedError.statusCode = statusCode ?? 404;
        throw enrichedError;
      }
      throw err;
    }
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
    let item;
    try {
      item = await this.graphClient.getDriveItem(
        validated.drive_id,
        validated.item_id
      );
    } catch (err) {
      const code = (err as { code?: string }).code;
      const statusCode = (err as { statusCode?: number }).statusCode;
      if (code === 'itemNotFound' || (statusCode === 404 && !code)) {
        const hint = `File not found: item '${validated.item_id}' in drive '${validated.drive_id}'. The IDs may be stale. Remediation: call sp_list_children with the drive_id to get fresh item IDs, then retry sp_get_file.`;
        const enrichedError = new Error(hint) as Error & { code?: string; statusCode?: number };
        enrichedError.code = code ?? 'itemNotFound';
        enrichedError.statusCode = statusCode ?? 404;
        throw enrichedError;
      }
      throw err;
    }

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
        } else if (isParsableMimeType(mimeType)) {
          // Parse document formats (PDF, Office) to extract text
          try {
            const parsed = await parseFileContent(
              contentResult.content,
              mimeType,
              item.name ?? 'unknown'
            );
            result['content'] = parsed.text;
            result['contentType'] = 'parsed_text';
            result['parsedFormat'] = parsed.format;
            result['truncated'] = parsed.truncated;
          } catch (parseErr) {
            // Fallback to base64 if parsing fails
            const parseMessage = parseErr instanceof Error ? parseErr.message : 'Unknown parsing error';
            logger.warn(
              { err: parseErr, mimeType, fileName: item.name },
              'Document parsing failed, falling back to base64'
            );
            result['content'] = contentResult.content.toString('base64');
            result['contentType'] = 'base64';
            result['parseError'] = parseMessage;
          }
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
      'Search and list SharePoint sites accessible to the user. Returns site IDs that MUST be used as-is in sp_list_drives. Always call this first before accessing any SharePoint site content.',
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
      "List document libraries/drives for a SharePoint site. Without a site_id, returns the user's personal OneDrive. IMPORTANT: The site_id MUST be the exact 'id' value returned by sp_list_sites (format: 'hostname,siteCollectionId,siteId'). Do not construct or guess site IDs.",
    inputSchema: {
      type: 'object' as const,
      properties: {
        site_id: {
          type: 'string',
          description:
            "The exact site ID from an sp_list_sites response (e.g., 'contoso.sharepoint.com,guid1,guid2'). Do not guess or construct this value.",
        },
      },
    },
  },
  {
    name: 'sp_list_children',
    description:
      'List files and folders in a drive. Provide item_id to list contents of a specific folder, or omit for root. IMPORTANT: drive_id MUST be the exact ID from an sp_list_drives response. item_id MUST be from a previous sp_list_children response.',
    inputSchema: {
      type: 'object' as const,
      properties: {
        drive_id: {
          type: 'string',
          description: 'The exact drive ID from an sp_list_drives response. Do not guess or construct this value.',
        },
        item_id: {
          type: 'string',
          description:
            'The exact folder ID from a previous sp_list_children response. If not provided, lists root folder contents.',
        },
      },
      required: ['drive_id'],
    },
  },
  {
    name: 'sp_get_file',
    description:
      'Get file metadata and content. Text files are returned as text. PDF, Word (.docx), Excel (.xlsx), and PowerPoint (.pptx) are automatically parsed to extract readable text. Other binary files are returned as base64. Maximum size: 10MB. IMPORTANT: drive_id and item_id MUST come from previous sp_list_drives / sp_list_children responses.',
    inputSchema: {
      type: 'object' as const,
      properties: {
        drive_id: {
          type: 'string',
          description: 'The exact drive ID from an sp_list_drives response.',
        },
        item_id: {
          type: 'string',
          description: 'The exact file ID from an sp_list_children response.',
        },
      },
      required: ['drive_id', 'item_id'],
    },
  },
];
