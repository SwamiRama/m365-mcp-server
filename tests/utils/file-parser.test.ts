import { describe, it, expect, vi } from 'vitest';

// Mock config before importing
vi.mock('../../src/utils/config.js', () => ({
  config: {
    fileParseTimeoutMs: 30000,
    fileParseMaxOutputKb: 500,
    logLevel: 'info',
    nodeEnv: 'test',
  },
}));

// Mock logger
vi.mock('../../src/utils/logger.js', () => ({
  logger: {
    info: vi.fn(),
    warn: vi.fn(),
    error: vi.fn(),
    debug: vi.fn(),
  },
}));

import { isParsableMimeType, parseFileContent } from '../../src/utils/file-parser.js';

describe('file-parser', () => {
  describe('isParsableMimeType', () => {
    it('should recognize PDF MIME type', () => {
      expect(isParsableMimeType('application/pdf')).toBe(true);
    });

    it('should recognize Word MIME type', () => {
      expect(
        isParsableMimeType(
          'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
        )
      ).toBe(true);
    });

    it('should recognize Excel MIME type', () => {
      expect(
        isParsableMimeType(
          'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )
      ).toBe(true);
    });

    it('should recognize PowerPoint MIME type', () => {
      expect(
        isParsableMimeType(
          'application/vnd.openxmlformats-officedocument.presentationml.presentation'
        )
      ).toBe(true);
    });

    it('should recognize CSV MIME type', () => {
      expect(isParsableMimeType('text/csv')).toBe(true);
    });

    it('should recognize HTML MIME type', () => {
      expect(isParsableMimeType('text/html')).toBe(true);
    });

    it('should reject unknown MIME types', () => {
      expect(isParsableMimeType('image/png')).toBe(false);
      expect(isParsableMimeType('application/octet-stream')).toBe(false);
      expect(isParsableMimeType('video/mp4')).toBe(false);
    });

    it('should recognize legacy Office MIME types', () => {
      expect(isParsableMimeType('application/msword')).toBe(true);
      expect(isParsableMimeType('application/vnd.ms-excel')).toBe(true);
      expect(isParsableMimeType('application/vnd.ms-powerpoint')).toBe(true);
    });
  });

  describe('parseFileContent', () => {
    it('should parse CSV content', async () => {
      const csv = 'Name,Age\nAlice,30\nBob,25';
      const buffer = Buffer.from(csv, 'utf-8');

      const result = await parseFileContent(buffer, 'text/csv', 'data.csv');

      expect(result.text).toBe(csv);
      expect(result.truncated).toBe(false);
      expect(result.format).toBe('csv');
    });

    it('should strip HTML tags', async () => {
      const html = '<html><body><h1>Title</h1><p>Hello <b>world</b></p></body></html>';
      const buffer = Buffer.from(html, 'utf-8');

      const result = await parseFileContent(buffer, 'text/html', 'page.html');

      expect(result.text).toContain('Title');
      expect(result.text).toContain('Hello world');
      expect(result.text).not.toContain('<h1>');
      expect(result.text).not.toContain('<b>');
      expect(result.format).toBe('html');
    });

    it('should strip script and style tags from HTML', async () => {
      const html = '<html><head><style>.red{color:red}</style></head><body><script>alert(1)</script><p>Content</p></body></html>';
      const buffer = Buffer.from(html, 'utf-8');

      const result = await parseFileContent(buffer, 'text/html', 'page.html');

      expect(result.text).toContain('Content');
      expect(result.text).not.toContain('alert');
      expect(result.text).not.toContain('.red');
    });

    it('should throw for unsupported MIME types', async () => {
      const buffer = Buffer.from('data');
      await expect(
        parseFileContent(buffer, 'image/png', 'image.png')
      ).rejects.toThrow('Unsupported MIME type');
    });

    it('should truncate large output', async () => {
      // Create a CSV larger than the configured max (500KB)
      const line = 'a'.repeat(1000) + '\n';
      const csv = line.repeat(600); // ~600KB
      const buffer = Buffer.from(csv, 'utf-8');

      const result = await parseFileContent(buffer, 'text/csv', 'large.csv');

      expect(result.truncated).toBe(true);
      expect(result.text).toContain('[Content truncated due to size limit]');
    });

    it('should handle empty files', async () => {
      const buffer = Buffer.from('', 'utf-8');

      const result = await parseFileContent(buffer, 'text/csv', 'empty.csv');

      expect(result.text).toBe('');
      expect(result.truncated).toBe(false);
    });

    it('should decode HTML entities', async () => {
      const html = '<p>5 &gt; 3 &amp; 2 &lt; 4</p>';
      const buffer = Buffer.from(html, 'utf-8');

      const result = await parseFileContent(buffer, 'text/html', 'entities.html');

      expect(result.text).toContain('5 > 3 & 2 < 4');
    });
  });
});
