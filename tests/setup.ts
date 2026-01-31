import { vi } from 'vitest';

// Mock environment variables for tests
process.env['AZURE_CLIENT_ID'] = '00000000-0000-0000-0000-000000000000';
process.env['AZURE_CLIENT_SECRET'] = 'test-secret-value-for-testing';
process.env['AZURE_TENANT_ID'] = '00000000-0000-0000-0000-000000000001';
process.env['SESSION_SECRET'] = 'test-session-secret-at-least-32-chars-long';
process.env['NODE_ENV'] = 'test';

// Mock logger to reduce noise in tests
vi.mock('../src/utils/logger.js', () => ({
  logger: {
    info: vi.fn(),
    warn: vi.fn(),
    error: vi.fn(),
    debug: vi.fn(),
    trace: vi.fn(),
    child: vi.fn(() => ({
      info: vi.fn(),
      warn: vi.fn(),
      error: vi.fn(),
      debug: vi.fn(),
      trace: vi.fn(),
    })),
  },
  createRequestLogger: vi.fn(() => ({
    info: vi.fn(),
    warn: vi.fn(),
    error: vi.fn(),
    debug: vi.fn(),
    trace: vi.fn(),
  })),
}));
