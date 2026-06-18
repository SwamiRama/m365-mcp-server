import { describe, it, expect, vi } from 'vitest';
import { deflateRawSync } from 'node:zlib';

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

import { isParsableMimeType, parseFileContent, stripHtmlTags, inflateRawCapped } from '../../src/utils/file-parser.js';

// Builds a structurally valid single-page PDF (correct xref offsets) so the real
// pdf-parse (v2 / pdfjs) code path is exercised without mocking. This guards the
// v1->v2 API regression that produced "pdfParse is not a function" in production.
function makeMinimalPdf(text: string): Buffer {
  const stream = `BT /F1 24 Tf 72 700 Td (${text}) Tj ET`;
  const objects = [
    '<< /Type /Catalog /Pages 2 0 R >>',
    '<< /Type /Pages /Kids [3 0 R] /Count 1 >>',
    '<< /Type /Page /Parent 2 0 R /MediaBox [0 0 612 792] /Contents 4 0 R /Resources << /Font << /F1 5 0 R >> >> >>',
    `<< /Length ${stream.length} >>\nstream\n${stream}\nendstream`,
    '<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>',
  ];
  let pdf = '%PDF-1.4\n';
  const offsets: number[] = [];
  objects.forEach((body, i) => {
    offsets.push(Buffer.byteLength(pdf, 'latin1'));
    pdf += `${i + 1} 0 obj\n${body}\nendobj\n`;
  });
  const xrefOffset = Buffer.byteLength(pdf, 'latin1');
  pdf += `xref\n0 ${objects.length + 1}\n0000000000 65535 f \n`;
  offsets.forEach((off) => {
    pdf += `${String(off).padStart(10, '0')} 00000 n \n`;
  });
  pdf += `trailer\n<< /Size ${objects.length + 1} /Root 1 0 R >>\nstartxref\n${xrefOffset}\n%%EOF`;
  return Buffer.from(pdf, 'latin1');
}

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

    it('should extract text from a real PDF (regression: pdf-parse v2 API)', async () => {
      const buffer = makeMinimalPdf('Hello PDF World');

      const result = await parseFileContent(buffer, 'application/pdf', 'doc.pdf');

      expect(result.format).toBe('pdf');
      expect(result.text.length).toBeGreaterThan(0);
      expect(result.text).toContain('Hello');
      expect(result.text).toContain('World');
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

  describe('inflateRawCapped (decompression-bomb guard)', () => {
    it('round-trips normally-sized deflated data', () => {
      const original = Buffer.from('hello world, this is a slide');
      const result = inflateRawCapped(deflateRawSync(original), 1024);
      expect(result.equals(original)).toBe(true);
    });

    it('throws when decompressed output exceeds the cap (zip bomb)', () => {
      // 2MB of zeros compresses to a few hundred bytes (high ratio = bomb shape)
      const bomb = deflateRawSync(Buffer.alloc(2 * 1024 * 1024));
      expect(bomb.length).toBeLessThan(64 * 1024); // sanity: it really is tiny compressed
      expect(() => inflateRawCapped(bomb, 64 * 1024)).toThrow();
    });
  });

  describe('stripHtmlTags (exported)', () => {
    it('should strip basic HTML tags', () => {
      expect(stripHtmlTags('<p>Hello <b>world</b></p>')).toContain('Hello world');
    });

    it('should remove script and style tags with content', () => {
      const html = '<style>.red{color:red}</style><script>alert(1)</script><p>Content</p>';
      const result = stripHtmlTags(html);
      expect(result).toContain('Content');
      expect(result).not.toContain('alert');
      expect(result).not.toContain('.red');
    });

    it('should decode HTML entities', () => {
      expect(stripHtmlTags('5 &gt; 3 &amp; 2 &lt; 4')).toBe('5 > 3 & 2 < 4');
    });

    it('should convert br tags to newlines', () => {
      expect(stripHtmlTags('line1<br>line2<br/>line3')).toBe('line1\nline2\nline3');
    });

    it('should handle empty input', () => {
      expect(stripHtmlTags('')).toBe('');
    });

    it('should decode numeric HTML entities', () => {
      expect(stripHtmlTags('&#65;&#66;&#67;')).toBe('ABC');
    });
  });
});
