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
  site_name: z
    .string()
    .max(256)
    .optional()
    .describe(
      'Optional: SharePoint site name to scope the search. Example: "IZ - Newsletter". When provided, only files from this site are returned.'
    ),
  sort: z
    .enum(['relevance', 'lastModified'])
    .optional()
    .describe('Sort order: "relevance" (default) or "lastModified" (newest first)'),
  size: z
    .number()
    .int()
    .min(1)
    .max(25)
    .optional()
    .describe('Number of results to return (default: 10, max: 25)'),
});

export const searchAndReadInputSchema = z.object({
  query: z
    .string()
    .min(1)
    .max(512)
    .describe(
      'Search query to find the file. KQL supported. Examples: "Ersthelfer Berlin", "filename:budget.xlsx"'
    ),
  site_name: z
    .string()
    .max(256)
    .optional()
    .describe('Optional: SharePoint site name to scope the search.'),
  result_index: z
    .number()
    .int()
    .min(0)
    .max(24)
    .optional()
    .describe(
      'Which search result to read (0-based). Default: 0 (top result). Use this if a previous sp_search showed multiple matches and you want a specific one.'
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
export type SearchFilesInput = z.infer<typeof searchFilesInputSchema>;
export type SearchAndReadInput = z.infer<typeof searchAndReadInputSchema>;

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
   * Shared helper: fetch file content and parse it into the result object.
   * Used by both getFile and searchAndRead to avoid code duplication.
   */
  private async fetchAndParseContent(
    driveId: string,
    itemId: string,
    fileName: string,
    result: Record<string, unknown>
  ): Promise<void> {
    try {
      const contentResult = await this.graphClient.getFileContent(
        driveId,
        itemId,
        MAX_FILE_SIZE
      );

      if (contentResult) {
        const mimeType = contentResult.mimeType;

        if (isTextMimeType(mimeType)) {
          result['content'] = contentResult.content.toString('utf-8');
          result['contentType'] = 'text';
        } else if (isParsableMimeType(mimeType)) {
          try {
            const parsed = await parseFileContent(
              contentResult.content,
              mimeType,
              fileName
            );
            result['content'] = parsed.text;
            result['contentType'] = 'parsed_text';
            result['parsedFormat'] = parsed.format;
            result['truncated'] = parsed.truncated;
          } catch (parseErr) {
            const parseMessage =
              parseErr instanceof Error ? parseErr.message : 'Unknown parsing error';
            logger.warn(
              { err: parseErr, mimeType, fileName },
              'Document parsing failed, falling back to base64'
            );
            result['content'] = contentResult.content.toString('base64');
            result['contentType'] = 'base64';
            result['parseError'] = parseMessage;
          }
        } else {
          result['content'] = contentResult.content.toString('base64');
          result['contentType'] = 'base64';
        }

        result['contentMimeType'] = mimeType;
        result['contentSize'] = contentResult.size;
      }
    } catch (err) {
      const errorMessage = err instanceof Error ? err.message : 'Unknown error';
      logger.warn(
        { err, driveId, itemId },
        'Failed to fetch file content'
      );
      result['content'] = null;
      result['contentError'] = errorMessage;
    }
  }

  /**
   * Build a KQL query with optional site scoping.
   * Returns the final query string and the resolved site URL (if any).
   */
  private async buildScopedQuery(
    query: string,
    siteName?: string
  ): Promise<{ kqlQuery: string; resolvedSite?: string; siteError?: string }> {
    if (!siteName) {
      return { kqlQuery: query };
    }

    const siteUrl = await this.graphClient.resolveSiteWebUrl(siteName);
    if (siteUrl) {
      return {
        kqlQuery: `${query} path:"${siteUrl}"`,
        resolvedSite: siteUrl,
      };
    }

    return {
      kqlQuery: query,
      siteError: `Site "${siteName}" was not found. Search will run across all sites. Use sp_list_sites to see available sites.`,
    };
  }

  /**
   * Search for files across all SharePoint sites and OneDrive
   */
  async searchFiles(input: SearchFilesInput): Promise<object> {
    const validated = searchFilesInputSchema.parse(input);
    const size = validated.size ?? 10;

    logger.debug({ query: validated.query, siteName: validated.site_name, size }, 'Searching files');

    const { kqlQuery, resolvedSite, siteError } = await this.buildScopedQuery(
      validated.query,
      validated.site_name
    );

    const result = await this.graphClient.searchDriveItems({
      query: kqlQuery,
      size,
      sortBy: validated.sort ?? 'relevance',
    });

    const items = result.hits.map((hit: SearchHit, index: number) => {
      const resource = hit.resource;
      const driveId = resource.parentReference?.driveId;
      const itemId = resource.id;

      const formatted: Record<string, unknown> = {
        '#': index + 1,
        name: resource.name,
        webUrl: resource.webUrl,
        lastModified: resource.lastModifiedDateTime,
        size: resource.size,
      };

      if (resource.folder) {
        formatted['type'] = 'folder';
      } else {
        formatted['type'] = 'file';
        if (resource.file?.mimeType) {
          formatted['mimeType'] = resource.file.mimeType;
        }
      }

      // Include search snippet if available
      if (hit.summary) {
        formatted['summary'] = hit.summary.replace(/<\/?c0>/g, '').replace(/<ddd\/>/g, '…');
      }

      // Location context from parentReference
      if (resource.parentReference?.path) {
        formatted['location'] = resource.parentReference.path;
      }

      // Include IDs for sp_get_file
      if (driveId) {
        formatted['drive_id'] = driveId;
      }
      formatted['item_id'] = itemId;

      // Anti-hallucination: embed the exact call the LLM should make
      if (driveId && itemId && !resource.folder) {
        formatted['action'] = `To read this file: sp_get_file(drive_id="${driveId}", item_id="${itemId}")`;
      }

      return formatted;
    });

    return {
      results: items,
      total: result.total,
      moreResultsAvailable: result.moreResultsAvailable,
      ...(resolvedSite ? { site_filter: resolvedSite } : {}),
      ...(siteError ? { site_warning: siteError } : {}),
      _note:
        'IMPORTANT: To read a file, use the EXACT drive_id and item_id from a result above. Do NOT use IDs from earlier messages. Alternatively, use sp_search_read to search and read in one step.',
    };
  }

  /**
   * Search for a file and immediately return its content.
   * Combines search + file retrieval in one step to prevent ID mismatches.
   */
  async searchAndRead(input: SearchAndReadInput): Promise<object> {
    const validated = searchAndReadInputSchema.parse(input);
    const resultIndex = validated.result_index ?? 0;

    logger.debug(
      { query: validated.query, siteName: validated.site_name, resultIndex },
      'Search and read file'
    );

    // Step 1: Build scoped query
    const { kqlQuery, resolvedSite } = await this.buildScopedQuery(
      validated.query,
      validated.site_name
    );

    // Step 2: Search
    const searchResult = await this.graphClient.searchDriveItems({
      query: kqlQuery,
      size: Math.max(resultIndex + 1, 5),
      sortBy: 'relevance',
    });

    if (searchResult.hits.length === 0) {
      return {
        found: false,
        query: validated.query,
        ...(resolvedSite ? { site_filter: resolvedSite } : {}),
        _note: 'No files matched. Try broader search terms or check site_name spelling. Use sp_list_sites to see available sites.',
      };
    }

    if (resultIndex >= searchResult.hits.length) {
      return {
        found: false,
        query: validated.query,
        availableResults: searchResult.hits.length,
        _note: `result_index ${resultIndex} is out of range. Only ${searchResult.hits.length} result(s) found. Use a lower index or try sp_search first to see all results.`,
      };
    }

    // Step 3: Get the target hit (safe: bounds checked above)
    const hit = searchResult.hits[resultIndex]!;
    const resource = hit.resource;
    const driveId = resource.parentReference?.driveId;
    const itemId = resource.id;

    const result: Record<string, unknown> = {
      found: true,
      name: resource.name,
      webUrl: resource.webUrl,
      lastModified: resource.lastModifiedDateTime,
      size: resource.size,
      mimeType: resource.file?.mimeType,
      searchRank: resultIndex + 1,
      totalResults: searchResult.total,
    };

    if (hit.summary) {
      result['summary'] = hit.summary.replace(/<\/?c0>/g, '').replace(/<ddd\/>/g, '…');
    }

    if (resource.parentReference?.path) {
      result['location'] = resource.parentReference.path;
    }

    // Step 4: Fetch content using IDs directly from search hit (no hallucination possible)
    if (!driveId || !itemId) {
      result['content'] = null;
      result['contentError'] = 'Search result missing drive or item ID — cannot retrieve content.';
      return result;
    }

    if (resource.folder) {
      result['content'] = null;
      result['contentError'] = 'Search result is a folder, not a file.';
      return result;
    }

    const fileSize = resource.size ?? 0;
    if (fileSize > MAX_FILE_SIZE) {
      result['content'] = null;
      result['contentError'] = `File too large (${formatFileSize(fileSize)}). Maximum: ${formatFileSize(MAX_FILE_SIZE)}.`;
      return result;
    }

    await this.fetchAndParseContent(driveId, itemId, resource.name ?? 'unknown', result);

    return result;
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

    await this.fetchAndParseContent(
      validated.drive_id,
      validated.item_id,
      item.name ?? 'unknown',
      result
    );

    return result;
  }
}

// Tool definitions for MCP registration
export const sharePointToolDefinitions = [
  {
    name: 'sp_search_read',
    description:
      'Search for a file and immediately return its content in one step. This is the EASIEST and MOST RELIABLE way to find and read a document. ' +
      'Combines search + file retrieval, so there are no ID mismatches. ' +
      'Use this FIRST when the user asks "What does document X say?" or "Show me the content of Y" or any question that requires reading file content. ' +
      'PDF, Word, Excel, PowerPoint are auto-parsed to text. Max 10MB. ' +
      'Examples: query="Ersthelfer Berlin", query="budget report" site_name="Finance Team".',
    inputSchema: {
      type: 'object' as const,
      properties: {
        query: {
          type: 'string',
          description: 'Search query to find the file. KQL supported: "Ersthelfer Berlin", "filename:budget.xlsx".',
        },
        site_name: {
          type: 'string',
          description: 'Optional: restrict search to this SharePoint site. Example: "IZ - Newsletter".',
        },
        result_index: {
          type: 'number',
          description: 'Which search result to read (0 = top result, default). Use if previous sp_search showed multiple matches.',
        },
      },
      required: ['query'],
    },
  },
  {
    name: 'sp_search',
    description:
      'Search for files across ALL SharePoint sites and OneDrive. Returns file metadata and IDs, but NOT file content. ' +
      'If you need file CONTENT, prefer sp_search_read instead. ' +
      'Use this when the user wants to BROWSE or LIST files without reading them. ' +
      'Supports KQL syntax and optional site_name scoping. ' +
      'Examples: "Ersthelfer Berlin", "filename:budget.xlsx", "filetype:pdf quarterly".',
    inputSchema: {
      type: 'object' as const,
      properties: {
        query: {
          type: 'string',
          description: 'Search query. Supports KQL: "Ersthelfer Berlin", "filename:report.docx", "filetype:pdf budget".',
        },
        site_name: {
          type: 'string',
          description: 'Optional: restrict search to this SharePoint site. Example: "IZ - Newsletter".',
        },
        sort: {
          type: 'string',
          enum: ['relevance', 'lastModified'],
          description: 'Sort order: "relevance" (default) or "lastModified" (newest first).',
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
      'List SharePoint sites accessible to the user. Returns site IDs needed for sp_list_drives. Use this when the user wants to browse a specific site, NOT when searching for document content (use sp_search_read instead).',
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
      'Get file content by drive_id + item_id. PDF, Word (.docx), Excel (.xlsx), PowerPoint (.pptx) are auto-parsed to readable text. Max 10MB. ' +
      'IMPORTANT: Use the EXACT drive_id and item_id from the MOST RECENT sp_search or sp_list_children response. Do NOT use IDs from earlier messages.',
    inputSchema: {
      type: 'object' as const,
      properties: {
        drive_id: {
          type: 'string',
          description: 'The EXACT drive ID from the most recent sp_search or sp_list_children response.',
        },
        item_id: {
          type: 'string',
          description: 'The EXACT file ID from the most recent sp_search or sp_list_children response.',
        },
      },
      required: ['drive_id', 'item_id'],
    },
  },
];
