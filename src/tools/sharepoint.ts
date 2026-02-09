import { z } from 'zod';
import {
  GraphClient,
  type GraphSite,
  type GraphDrive,
  type GraphDriveItem,
  type SearchHit,
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
    .describe(
      "Site ID to list drives from. MUST be the exact 'id' from sp_list_sites. Always call sp_list_sites first."
    ),
});

export const searchFilesInputSchema = z.object({
  query: z
    .string()
    .min(1)
    .max(512)
    .describe(
      'Search query (supports KQL). Examples: "Ersthelfer Berlin", "filename:budget.xlsx", "filetype:pdf quarterly report"'
    ),
  size: z
    .number()
    .int()
    .min(1)
    .max(25)
    .optional()
    .describe('Number of results to return (default: 10, max: 25)'),
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
export type SearchFilesInput = z.infer<typeof searchFilesInputSchema>;

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
   * List drives (document libraries) for a SharePoint site
   */
  async listDrives(input: ListDrivesInput): Promise<object> {
    const validated = listDrivesInputSchema.parse(input);

    logger.debug({ siteId: validated.site_id }, 'Listing drives');

    const drives = await this.graphClient.listSiteDrives(validated.site_id);

    return {
      drives: drives.map(formatDrive),
      count: drives.length,
      site_id: validated.site_id,
      _note: 'Use the drive id values from this response with sp_list_children to browse files, or use sp_get_file to read a file.',
    };
  }

  /**
   * Search for files across all SharePoint sites and OneDrive
   */
  async searchFiles(input: SearchFilesInput): Promise<object> {
    const validated = searchFilesInputSchema.parse(input);
    const size = validated.size ?? 10;

    logger.debug({ query: validated.query, size }, 'Searching files');

    const result = await this.graphClient.searchDriveItems({
      query: validated.query,
      size,
    });

    const items = result.hits.map((hit: SearchHit) => {
      const resource = hit.resource;
      const formatted: Record<string, unknown> = {
        name: resource.name,
        webUrl: resource.webUrl,
        lastModified: resource.lastModifiedDateTime,
        size: resource.size,
      };

      // Include drive/item IDs so LLM can call sp_get_file directly
      if (resource.parentReference?.driveId) {
        formatted['drive_id'] = resource.parentReference.driveId;
      }
      formatted['item_id'] = resource.id;

      if (resource.file) {
        formatted['type'] = 'file';
        formatted['mimeType'] = resource.file.mimeType;
      } else if (resource.folder) {
        formatted['type'] = 'folder';
      }

      // Include search snippet if available
      if (hit.summary) {
        // Clean up highlight tags from summary
        formatted['summary'] = hit.summary.replace(/<\/?c0>/g, '');
      }

      return formatted;
    });

    return {
      results: items,
      total: result.total,
      moreResultsAvailable: result.moreResultsAvailable,
      _note: 'Use drive_id and item_id with sp_get_file to retrieve file content. Use sp_list_children with drive_id to browse folders.',
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
    name: 'sp_search',
    description:
      'Search for files across ALL SharePoint sites and OneDrive. This is the FASTEST way to find documents. Supports KQL syntax. Returns drive_id and item_id that can be passed directly to sp_get_file. Use this FIRST when the user asks about document content (e.g. "Wer sind die Ersthelfer?", "find the budget report"). Examples: "Ersthelfer Berlin", "filename:budget.xlsx", "filetype:pdf quarterly".',
    inputSchema: {
      type: 'object' as const,
      properties: {
        query: {
          type: 'string',
          description: 'Search query. Supports KQL: "Ersthelfer Berlin", "filename:report.docx", "filetype:pdf budget".',
        },
        size: {
          type: 'number',
          description: 'Number of results (default: 10, max: 25)',
        },
      },
      required: ['query'],
    },
  },
  {
    name: 'sp_list_sites',
    description:
      'List SharePoint sites accessible to the user. Returns site IDs needed for sp_list_drives. Use this when the user wants to browse a specific site, NOT when searching for document content (use sp_search instead).',
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
      "List document libraries (drives) for a specific SharePoint site. REQUIRES site_id from sp_list_sites. Use this to browse a site's document libraries before listing files with sp_list_children.",
    inputSchema: {
      type: 'object' as const,
      properties: {
        site_id: {
          type: 'string',
          description:
            "REQUIRED. The exact site ID from sp_list_sites (format: 'hostname,guid1,guid2'). Do not guess this value.",
        },
      },
      required: ['site_id'],
    },
  },
  {
    name: 'sp_list_children',
    description:
      'List files and folders in a drive. Use for browsing folder contents. Provide item_id for a subfolder, or omit for root. drive_id MUST come from sp_list_drives or sp_search.',
    inputSchema: {
      type: 'object' as const,
      properties: {
        drive_id: {
          type: 'string',
          description: 'The exact drive ID from sp_list_drives or sp_search. Do not guess.',
        },
        item_id: {
          type: 'string',
          description:
            'Folder ID from sp_list_children. Omit to list root folder.',
        },
      },
      required: ['drive_id'],
    },
  },
  {
    name: 'sp_get_file',
    description:
      'Get file content. PDF, Word (.docx), Excel (.xlsx), PowerPoint (.pptx) are automatically parsed to readable text. Max 10MB. Use drive_id + item_id from sp_search or sp_list_children.',
    inputSchema: {
      type: 'object' as const,
      properties: {
        drive_id: {
          type: 'string',
          description: 'Drive ID from sp_search, sp_list_drives, or sp_list_children.',
        },
        item_id: {
          type: 'string',
          description: 'File ID from sp_search or sp_list_children.',
        },
      },
      required: ['drive_id', 'item_id'],
    },
  },
];
