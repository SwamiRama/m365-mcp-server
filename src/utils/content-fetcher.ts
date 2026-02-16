import type { GraphClient } from '../graph/client.js';
import { isParsableMimeType, parseFileContent } from './file-parser.js';
import { logger } from './logger.js';

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

export function isTextMimeType(mimeType: string): boolean {
  if (TEXT_MIME_TYPES.has(mimeType)) return true;
  if (mimeType.startsWith('text/')) return true;
  if (mimeType.endsWith('+json')) return true;
  if (mimeType.endsWith('+xml')) return true;
  return false;
}

export function formatFileSize(bytes: number): string {
  if (bytes < 1024) return `${bytes} B`;
  if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(1)} KB`;
  if (bytes < 1024 * 1024 * 1024) return `${(bytes / (1024 * 1024)).toFixed(1)} MB`;
  return `${(bytes / (1024 * 1024 * 1024)).toFixed(1)} GB`;
}

/**
 * Fetch file content from a drive and parse it into the result object.
 * Shared between SharePoint and OneDrive tools.
 */
export async function fetchAndParseContent(
  graphClient: GraphClient,
  driveId: string,
  itemId: string,
  fileName: string,
  maxFileSize: number,
  result: Record<string, unknown>
): Promise<void> {
  try {
    const contentResult = await graphClient.getFileContent(
      driveId,
      itemId,
      maxFileSize
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
