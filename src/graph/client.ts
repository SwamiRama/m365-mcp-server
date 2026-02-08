import { Client } from '@microsoft/microsoft-graph-client';
import { config } from '../utils/config.js';
import { logger } from '../utils/logger.js';
import type { TokenSet } from '../auth/session.js';

// Graph API types (simplified)
export interface GraphUser {
  id: string;
  displayName?: string;
  mail?: string;
  userPrincipalName?: string;
}

export interface GraphMessage {
  id: string;
  subject?: string;
  bodyPreview?: string;
  body?: {
    contentType: string;
    content: string;
  };
  from?: {
    emailAddress?: {
      name?: string;
      address?: string;
    };
  };
  toRecipients?: Array<{
    emailAddress?: {
      name?: string;
      address?: string;
    };
  }>;
  receivedDateTime?: string;
  sentDateTime?: string;
  hasAttachments?: boolean;
  isRead?: boolean;
  importance?: string;
  webLink?: string;
}

export interface GraphMailFolder {
  id: string;
  displayName?: string;
  parentFolderId?: string;
  childFolderCount?: number;
  unreadItemCount?: number;
  totalItemCount?: number;
}

export interface GraphSite {
  id: string;
  name?: string;
  displayName?: string;
  webUrl?: string;
  description?: string;
}

export interface GraphDrive {
  id: string;
  name?: string;
  driveType?: string;
  webUrl?: string;
  owner?: {
    user?: {
      displayName?: string;
    };
  };
}

export interface GraphDriveItem {
  id: string;
  name?: string;
  webUrl?: string;
  size?: number;
  createdDateTime?: string;
  lastModifiedDateTime?: string;
  file?: {
    mimeType?: string;
  };
  folder?: {
    childCount?: number;
  };
  parentReference?: {
    driveId?: string;
    id?: string;
    path?: string;
  };
}

// Rate limiting / retry configuration
const MAX_RETRIES = 3;
const RETRY_DELAY_MS = 1000;

function isGraphError(err: unknown): err is { statusCode: number; code?: string; message?: string } {
  return (
    typeof err === 'object' &&
    err !== null &&
    'statusCode' in err &&
    typeof (err as Record<string, unknown>)['statusCode'] === 'number'
  );
}

async function sleep(ms: number): Promise<void> {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

export class GraphClient {
  private client: Client;
  private accessToken: string;

  constructor(tokens: TokenSet) {
    this.accessToken = tokens.accessToken;

    this.client = Client.init({
      authProvider: (done) => {
        done(null, this.accessToken);
      },
      defaultVersion: 'v1.0',
    });
  }

  /**
   * Execute a Graph API request with retry logic for rate limiting
   */
  private async executeWithRetry<T>(
    operation: () => Promise<T>,
    operationName: string
  ): Promise<T> {
    let lastError: unknown;

    for (let attempt = 1; attempt <= MAX_RETRIES; attempt++) {
      try {
        const controller = new AbortController();
        const timeoutId = setTimeout(
          () => controller.abort(),
          config.graphApiTimeoutMs
        );

        try {
          const result = await operation();
          clearTimeout(timeoutId);
          return result;
        } catch (err) {
          clearTimeout(timeoutId);
          throw err;
        }
      } catch (err) {
        lastError = err;

        if (isGraphError(err)) {
          // Handle rate limiting (429)
          if (err.statusCode === 429) {
            const retryAfterHeader = 'retry-after';
            const retryAfter = (err as Record<string, unknown>)[retryAfterHeader];
            const delayMs =
              typeof retryAfter === 'string'
                ? parseInt(retryAfter, 10) * 1000
                : RETRY_DELAY_MS * Math.pow(2, attempt - 1);

            logger.warn(
              { operationName, attempt, delayMs },
              'Rate limited, retrying...'
            );

            await sleep(delayMs);
            continue;
          }

          // Don't retry client errors (4xx except 429)
          if (err.statusCode >= 400 && err.statusCode < 500) {
            throw err;
          }

          // Retry server errors (5xx)
          if (err.statusCode >= 500) {
            const delayMs = RETRY_DELAY_MS * Math.pow(2, attempt - 1);
            logger.warn(
              { operationName, attempt, statusCode: err.statusCode, delayMs },
              'Server error, retrying...'
            );
            await sleep(delayMs);
            continue;
          }
        }

        throw err;
      }
    }

    throw lastError;
  }

  // === User Operations ===

  async getMe(): Promise<GraphUser> {
    return this.executeWithRetry(
      () => this.client.api('/me').get(),
      'getMe'
    );
  }

  // === Mail Operations ===

  async listMailFolders(userId?: string): Promise<GraphMailFolder[]> {
    const base = userId ? `/users/${userId}` : '/me';
    const result = await this.executeWithRetry(
      () =>
        this.client
          .api(`${base}/mailFolders`)
          .select('id,displayName,parentFolderId,childFolderCount,unreadItemCount,totalItemCount')
          .get(),
      'listMailFolders'
    );

    return result.value ?? [];
  }

  async listMessages(options: {
    folderId?: string;
    top?: number;
    skip?: number;
    filter?: string;
    search?: string;
    select?: string[];
    orderBy?: string;
    userId?: string;
  }): Promise<{ messages: GraphMessage[]; nextLink?: string }> {
    const { folderId, top = 25, skip, filter, search, select, orderBy, userId } = options;

    const base = userId ? `/users/${userId}` : '/me';
    const endpoint = folderId
      ? `${base}/mailFolders/${folderId}/messages`
      : `${base}/messages`;

    let request = this.client.api(endpoint);

    if (top) request = request.top(top);
    if (skip) request = request.skip(skip);

    // $search and $filter/$orderby are mutually exclusive in Graph API
    if (search) {
      request = request.search(`"${search.replace(/"/g, '\\"')}"`);
      // Graph API requires ConsistencyLevel header for $search
      request = request.header('ConsistencyLevel', 'eventual');
    } else {
      if (filter) request = request.filter(filter);
      if (orderBy) request = request.orderby(orderBy);
    }

    const selectFields = select ?? [
      'id',
      'subject',
      'bodyPreview',
      'from',
      'toRecipients',
      'receivedDateTime',
      'sentDateTime',
      'hasAttachments',
      'isRead',
      'importance',
      'webLink',
    ];
    request = request.select(selectFields.join(','));

    const result = await this.executeWithRetry(() => request.get(), 'listMessages');

    return {
      messages: result.value ?? [],
      nextLink: result['@odata.nextLink'],
    };
  }

  async getMessage(
    messageId: string,
    includeBody: boolean = false,
    userId?: string
  ): Promise<GraphMessage> {
    const selectFields = [
      'id',
      'subject',
      'bodyPreview',
      'from',
      'toRecipients',
      'receivedDateTime',
      'sentDateTime',
      'hasAttachments',
      'isRead',
      'importance',
      'webLink',
    ];

    if (includeBody) {
      selectFields.push('body');
    }

    const base = userId ? `/users/${userId}` : '/me';
    return this.executeWithRetry(
      () =>
        this.client
          .api(`${base}/messages/${messageId}`)
          .select(selectFields.join(','))
          .get(),
      'getMessage'
    );
  }

  // === SharePoint/OneDrive Operations ===

  async listSites(options: {
    search?: string;
    top?: number;
  }): Promise<GraphSite[]> {
    const { search, top = 25 } = options;

    // Microsoft Graph requires a search parameter to list sites
    // Use '*' as wildcard to get all accessible sites when no search is provided
    const searchQuery = search || '*';

    const request = this.client
      .api('/sites')
      .query({ search: searchQuery })
      .top(top)
      .select('id,name,displayName,webUrl,description');

    const result = await this.executeWithRetry(() => request.get(), 'listSites');

    return result.value ?? [];
  }

  async listDrives(siteId?: string): Promise<GraphDrive[]> {
    const endpoint = siteId ? `/sites/${siteId}/drives` : '/me/drives';

    const result = await this.executeWithRetry(
      () =>
        this.client
          .api(endpoint)
          .select('id,name,driveType,webUrl,owner')
          .get(),
      'listDrives'
    );

    return result.value ?? [];
  }

  async getMyDrive(): Promise<GraphDrive> {
    return this.executeWithRetry(
      () =>
        this.client
          .api('/me/drive')
          .select('id,name,driveType,webUrl,owner')
          .get(),
      'getMyDrive'
    );
  }

  async listDriveItems(options: {
    driveId: string;
    itemId?: string;
    top?: number;
  }): Promise<GraphDriveItem[]> {
    const { driveId, itemId, top = 50 } = options;

    // If itemId is provided, list children of that item
    // Otherwise, list root items
    const endpoint = itemId
      ? `/drives/${driveId}/items/${itemId}/children`
      : `/drives/${driveId}/root/children`;

    const result = await this.executeWithRetry(
      () =>
        this.client
          .api(endpoint)
          .top(top)
          .select(
            'id,name,webUrl,size,createdDateTime,lastModifiedDateTime,file,folder,parentReference'
          )
          .get(),
      'listDriveItems'
    );

    return result.value ?? [];
  }

  async getDriveItem(driveId: string, itemId: string): Promise<GraphDriveItem> {
    return this.executeWithRetry(
      () =>
        this.client
          .api(`/drives/${driveId}/items/${itemId}`)
          .select(
            'id,name,webUrl,size,createdDateTime,lastModifiedDateTime,file,folder,parentReference'
          )
          .get(),
      'getDriveItem'
    );
  }

  async getFileContent(
    driveId: string,
    itemId: string,
    maxSize: number = 10 * 1024 * 1024 // 10MB default
  ): Promise<{ content: Buffer; mimeType: string; size: number } | null> {
    // First, get the item metadata to check size
    const item = await this.getDriveItem(driveId, itemId);

    if (!item.file) {
      throw new Error('Item is not a file');
    }

    const size = item.size ?? 0;
    if (size > maxSize) {
      throw new Error(
        `File size (${size} bytes) exceeds maximum allowed (${maxSize} bytes)`
      );
    }

    const response = await this.executeWithRetry(
      () =>
        this.client
          .api(`/drives/${driveId}/items/${itemId}/content`)
          .responseType('arraybuffer' as unknown as import('@microsoft/microsoft-graph-client').ResponseType)
          .get(),
      'getFileContent'
    );

    return {
      content: Buffer.from(response as ArrayBuffer),
      mimeType: item.file.mimeType ?? 'application/octet-stream',
      size,
    };
  }
}

/**
 * Create a Graph client with the given tokens
 */
export function createGraphClient(tokens: TokenSet): GraphClient {
  return new GraphClient(tokens);
}
