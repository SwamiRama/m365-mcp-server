/**
 * File content parser for Office documents and PDFs.
 * Extracts readable text from binary document formats.
 */

import { inflateRawSync } from 'node:zlib';
import { logger } from './logger.js';
import { config } from './config.js';

// MIME types that can be parsed to extract text content
const PARSABLE_MIME_TYPES = new Map<string, string>([
  ['application/pdf', 'pdf'],
  [
    'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
    'docx',
  ],
  ['application/msword', 'doc'],
  [
    'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    'xlsx',
  ],
  ['application/vnd.ms-excel', 'xls'],
  [
    'application/vnd.openxmlformats-officedocument.presentationml.presentation',
    'pptx',
  ],
  ['application/vnd.ms-powerpoint', 'ppt'],
  ['text/csv', 'csv'],
  ['text/html', 'html'],
]);

export interface ParseResult {
  text: string;
  truncated: boolean;
  format: string;
}

/**
 * Check if a MIME type can be parsed to extract text.
 */
export function isParsableMimeType(mimeType: string): boolean {
  return PARSABLE_MIME_TYPES.has(mimeType);
}

/**
 * Parse file content and extract readable text.
 * Applies timeout and output size limits for safety.
 */
export async function parseFileContent(
  buffer: Buffer,
  mimeType: string,
  fileName: string
): Promise<ParseResult> {
  const format = PARSABLE_MIME_TYPES.get(mimeType);
  if (!format) {
    throw new Error(`Unsupported MIME type for parsing: ${mimeType}`);
  }

  const timeoutMs = config.fileParseTimeoutMs;
  const maxOutputBytes = config.fileParseMaxOutputKb * 1024;

  logger.debug(
    { fileName, mimeType, format, bufferSize: buffer.length },
    'Parsing file content'
  );

  let text: string;
  try {
    text = await withTimeout(
      parseByFormat(buffer, format, fileName),
      timeoutMs,
      `File parsing timed out after ${timeoutMs}ms`
    );
  } catch (err) {
    const message = err instanceof Error ? err.message : String(err);
    logger.warn({ err: err instanceof Error ? err : { message: String(err) }, fileName, mimeType }, 'File parsing failed');
    throw new Error(`Failed to parse ${fileName}: ${message}`);
  }

  // Truncate if output exceeds limit
  let truncated = false;
  if (Buffer.byteLength(text, 'utf-8') > maxOutputBytes) {
    text = truncateToByteLimit(text, maxOutputBytes);
    truncated = true;
    logger.debug(
      { fileName, maxOutputBytes },
      'Parsed content truncated due to size limit'
    );
  }

  return { text, truncated, format };
}

/**
 * Route to the appropriate parser based on format.
 */
async function parseByFormat(
  buffer: Buffer,
  format: string,
  fileName: string
): Promise<string> {
  switch (format) {
    case 'pdf':
      return parsePdf(buffer);
    case 'docx':
    case 'doc':
      return parseDocx(buffer, fileName);
    case 'xlsx':
    case 'xls':
      return parseExcel(buffer);
    case 'pptx':
    case 'ppt':
      return parsePptx(buffer, fileName);
    case 'csv':
      return buffer.toString('utf-8');
    case 'html':
      return stripHtmlTags(buffer.toString('utf-8'));
    default:
      throw new Error(`No parser available for format: ${format}`);
  }
}

/**
 * Parse PDF to text using pdf-parse.
 */
async function parsePdf(buffer: Buffer): Promise<string> {
  // pdf-parse uses CommonJS export = syntax — handle both ESM default and CJS shapes
  const pdfParseModule = await import('pdf-parse');
  const pdfParse = ((pdfParseModule as any).default ?? pdfParseModule) as unknown as (
    dataBuffer: Buffer,
    options?: { max?: number }
  ) => Promise<{ text: string }>;
  const result = await pdfParse(buffer, {
    max: 0, // No page limit (0 = all pages)
  });
  return result.text;
}

/**
 * Parse Word documents (.docx) to text using mammoth.
 */
async function parseDocx(buffer: Buffer, fileName: string): Promise<string> {
  const mammoth = await import('mammoth');
  const result = await mammoth.extractRawText({ buffer });

  if (result.messages.length > 0) {
    logger.debug(
      { fileName, messages: result.messages.map((m) => m.message) },
      'Mammoth parsing messages'
    );
  }

  return result.value;
}

/**
 * Parse Excel files (.xlsx/.xls) to text using exceljs.
 * Formats each sheet as a tab-separated table.
 */
async function parseExcel(buffer: Buffer): Promise<string> {
  const ExcelJSModule = await import('exceljs');
  const ExcelJS = ExcelJSModule.default ?? ExcelJSModule;
  const workbook = new ExcelJS.Workbook();
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  await workbook.xlsx.load(buffer as any);

  const sheets: string[] = [];

  workbook.eachSheet((worksheet) => {
    const lines: string[] = [];
    lines.push(`=== Sheet: ${worksheet.name} ===`);

    worksheet.eachRow({ includeEmpty: false }, (row) => {
      const values = row.values as unknown[];
      // row.values is 1-indexed (index 0 is undefined)
      const cells = values
        .slice(1)
        .map((cell) => formatCellValue(cell));
      lines.push(cells.join('\t'));
    });

    if (lines.length > 1) {
      sheets.push(lines.join('\n'));
    }
  });

  return sheets.join('\n\n');
}

/**
 * Format an Excel cell value to a string.
 */
function formatCellValue(value: unknown): string {
  if (value === null || value === undefined) return '';
  if (typeof value === 'object' && value !== null) {
    // Handle rich text, formulas, etc.
    if ('result' in value) return String((value as { result: unknown }).result);
    if ('text' in value) return String((value as { text: unknown }).text);
    if ('richText' in value) {
      const rt = value as { richText: Array<{ text: string }> };
      return rt.richText.map((r) => r.text).join('');
    }
    return String(value);
  }
  return String(value);
}

/**
 * Parse PowerPoint files (.pptx) by extracting text from XML slides.
 * Uses built-in Node.js zlib for ZIP decompression (no external dependencies).
 */
async function parsePptx(buffer: Buffer, fileName: string): Promise<string> {
  try {
    const entries = extractZipEntries(buffer);
    const slideTexts: string[] = [];

    // Sort slide entries by name for correct order
    const slideEntries = entries
      .filter(
        (e) =>
          e.name.startsWith('ppt/slides/slide') && e.name.endsWith('.xml')
      )
      .sort((a, b) => {
        const numA = parseInt(a.name.match(/slide(\d+)/)?.[1] ?? '0');
        const numB = parseInt(b.name.match(/slide(\d+)/)?.[1] ?? '0');
        return numA - numB;
      });

    for (const entry of slideEntries) {
      const xml = entry.data.toString('utf-8');
      const text = extractTextFromPptxXml(xml);
      if (text.trim()) {
        const slideNum = entry.name.match(/slide(\d+)/)?.[1] ?? '?';
        slideTexts.push(`=== Slide ${slideNum} ===\n${text}`);
      }
    }

    return slideTexts.join('\n\n');
  } catch (err) {
    logger.warn(
      { err: err instanceof Error ? err : { message: String(err) }, fileName },
      'PPTX extraction failed'
    );
    throw err;
  }
}

/**
 * Extract text content from PowerPoint slide XML.
 * Parses <a:p> paragraphs containing <a:t> text elements (OOXML format).
 */
function extractTextFromPptxXml(xml: string): string {
  // Group text by paragraphs (<a:p> tags)
  const paragraphs: string[] = [];
  const pRegex = /<a:p\b[^>]*>([\s\S]*?)<\/a:p>/g;
  let pMatch;
  while ((pMatch = pRegex.exec(xml)) !== null) {
    const pContent = pMatch[1]!;
    const pTexts: string[] = [];
    const tRegex = /<a:t>([\s\S]*?)<\/a:t>/g;
    let tMatch;
    while ((tMatch = tRegex.exec(pContent)) !== null) {
      pTexts.push(decodeXmlEntities(tMatch[1]!));
    }
    if (pTexts.length > 0) {
      paragraphs.push(pTexts.join(''));
    }
  }

  if (paragraphs.length > 0) {
    return paragraphs.join('\n');
  }

  // Fallback: extract all <a:t> tags if no <a:p> structure found
  const textParts: string[] = [];
  const regex = /<a:t>([\s\S]*?)<\/a:t>/g;
  let match;
  while ((match = regex.exec(xml)) !== null) {
    textParts.push(decodeXmlEntities(match[1]!));
  }
  return textParts.join(' ');
}

/**
 * Decode common XML entities.
 */
function decodeXmlEntities(text: string): string {
  return text
    .replace(/&amp;/g, '&')
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'");
}

/**
 * Minimal ZIP entry extraction for PPTX/DOCX files.
 * Handles the ZIP format using only built-in Node.js modules.
 */
function extractZipEntries(
  buffer: Buffer
): Array<{ name: string; data: Buffer }> {
  const entries: Array<{ name: string; data: Buffer }> = [];

  // Find End of Central Directory record (search from end of file)
  let eocdOffset = -1;
  for (let i = buffer.length - 22; i >= 0; i--) {
    if (
      buffer[i] === 0x50 &&
      buffer[i + 1] === 0x4b &&
      buffer[i + 2] === 0x05 &&
      buffer[i + 3] === 0x06
    ) {
      eocdOffset = i;
      break;
    }
  }

  if (eocdOffset === -1) {
    throw new Error('Not a valid ZIP file');
  }

  const centralDirOffset = buffer.readUInt32LE(eocdOffset + 16);
  const centralDirEntries = buffer.readUInt16LE(eocdOffset + 10);

  let offset = centralDirOffset;
  for (let i = 0; i < centralDirEntries; i++) {
    // Verify central directory entry signature
    if (
      offset + 46 > buffer.length ||
      buffer[offset] !== 0x50 ||
      buffer[offset + 1] !== 0x4b ||
      buffer[offset + 2] !== 0x01 ||
      buffer[offset + 3] !== 0x02
    ) {
      break;
    }

    const compressionMethod = buffer.readUInt16LE(offset + 10);
    const compressedSize = buffer.readUInt32LE(offset + 20);
    const fileNameLength = buffer.readUInt16LE(offset + 28);
    const extraFieldLength = buffer.readUInt16LE(offset + 30);
    const commentLength = buffer.readUInt16LE(offset + 32);
    const localHeaderOffset = buffer.readUInt32LE(offset + 42);

    const fileName = buffer.toString(
      'utf-8',
      offset + 46,
      offset + 46 + fileNameLength
    );

    // Only process slide XML files to avoid unnecessary decompression
    if (
      fileName.startsWith('ppt/slides/slide') &&
      fileName.endsWith('.xml')
    ) {
      // Read local file header to get actual data offset
      const localNameLength = buffer.readUInt16LE(localHeaderOffset + 26);
      const localExtraLength = buffer.readUInt16LE(localHeaderOffset + 28);
      const dataOffset =
        localHeaderOffset + 30 + localNameLength + localExtraLength;

      const compressedData = buffer.subarray(
        dataOffset,
        dataOffset + compressedSize
      );

      let data: Buffer;
      if (compressionMethod === 0) {
        // Stored (no compression)
        data = Buffer.from(compressedData);
      } else if (compressionMethod === 8) {
        // Deflated - use built-in zlib
        data = Buffer.from(inflateRawSync(compressedData));
      } else {
        // Skip unsupported compression methods
        offset += 46 + fileNameLength + extraFieldLength + commentLength;
        continue;
      }

      entries.push({ name: fileName, data });
    }

    offset += 46 + fileNameLength + extraFieldLength + commentLength;
  }

  return entries;
}

/**
 * Strip HTML tags and return plain text.
 */
function stripHtmlTags(html: string): string {
  return html
    .replace(/<script\b[^>]*>[\s\S]*?<\/script>/gi, '')
    .replace(/<style\b[^>]*>[\s\S]*?<\/style>/gi, '')
    .replace(/<br\s*\/?>/gi, '\n')
    .replace(/<\/p>/gi, '\n\n')
    .replace(/<\/div>/gi, '\n')
    .replace(/<\/li>/gi, '\n')
    .replace(/<\/h[1-6]>/gi, '\n\n')
    .replace(/<\/tr>/gi, '\n')
    .replace(/<td[^>]*>/gi, '\t')
    .replace(/<[^>]+>/g, '')
    .replace(/&nbsp;/g, ' ')
    .replace(/&amp;/g, '&')
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"')
    .replace(/&#(\d+);/g, (_, num: string) =>
      String.fromCharCode(parseInt(num))
    )
    .replace(/\n{3,}/g, '\n\n')
    .trim();
}

/**
 * Wrap a promise with a timeout.
 */
function withTimeout<T>(
  promise: Promise<T>,
  ms: number,
  message: string
): Promise<T> {
  return new Promise((resolve, reject) => {
    const timer = setTimeout(() => reject(new Error(message)), ms);
    promise
      .then((result) => {
        clearTimeout(timer);
        resolve(result);
      })
      .catch((err) => {
        clearTimeout(timer);
        reject(err);
      });
  });
}

/**
 * Truncate a string to fit within a byte limit (UTF-8).
 * Avoids splitting multi-byte characters.
 */
function truncateToByteLimit(text: string, maxBytes: number): string {
  const buf = Buffer.from(text, 'utf-8');
  if (buf.length <= maxBytes) return text;

  // Find a safe truncation point (don't split UTF-8 sequences)
  let end = maxBytes;
  while (end > 0 && (buf[end]! & 0xc0) === 0x80) {
    end--;
  }

  return buf.toString('utf-8', 0, end) + '\n\n[Content truncated due to size limit]';
}
