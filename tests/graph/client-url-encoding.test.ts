import { describe, it, expect, beforeEach, vi } from 'vitest';

vi.mock('../../src/utils/config.js', () => ({
  config: { graphApiTimeoutMs: 30000 },
}));
vi.mock('../../src/utils/logger.js', () => ({
  logger: { info: vi.fn(), warn: vi.fn(), error: vi.fn(), debug: vi.fn() },
}));

import { GraphClient } from '../../src/graph/client.js';

// Records every path passed to client.api(...) and returns a chainable stub whose
// terminal .get()/.getStream() resolve, so executeWithRetry/fetchAllPages complete.
function capturingClient(paths: string[]) {
  const chain: unknown = new Proxy(
    {},
    {
      get(_t, prop) {
        if (typeof prop === 'symbol' || prop === 'then') return undefined;
        if (prop === 'get' || prop === 'getStream') return async () => ({ value: [] });
        return () => chain;
      },
    }
  );
  return {
    api: (path: string) => {
      paths.push(path);
      return chain;
    },
  };
}

describe('GraphClient URL id encoding (regression: ErrorInvalidIdMalformed)', () => {
  let gc: GraphClient;
  let paths: string[];

  // Graph item IDs are base64 and routinely contain the URL-unsafe characters
  // + / = . Interpolated raw into the path, the '/' became a path separator and
  // Graph rejected the request with "Id is malformed". They must be encoded.
  const dirtyId = 'AAMkAGI2x+9/ab3cD=';

  beforeEach(() => {
    paths = [];
    gc = new GraphClient({ accessToken: 'tok' } as never);
    (gc as unknown as { client: unknown }).client = capturingClient(paths);
  });

  it('percent-encodes the message id in getMessage (user mailbox)', async () => {
    await gc.getMessage(dirtyId, false, 'user@crimeu.onmicrosoft.com');
    expect(paths[0]).toBe(
      `/users/user@crimeu.onmicrosoft.com/messages/${encodeURIComponent(dirtyId)}`
    );
    expect(paths[0]).not.toContain(dirtyId); // raw '+' / '/' must not leak into the path
  });

  it('percent-encodes message + attachment ids in getAttachment (personal mailbox)', async () => {
    const attId = 'ATT/x+y=z';
    await gc.getAttachment(dirtyId, attId);
    expect(paths[0]).toBe(
      `/me/messages/${encodeURIComponent(dirtyId)}/attachments/${encodeURIComponent(attId)}`
    );
  });

  it('percent-encodes drive + item ids in getDriveItem', async () => {
    await gc.getDriveItem('drive/+1', 'item/+2=');
    expect(paths[0]).toBe(
      `/drives/${encodeURIComponent('drive/+1')}/items/${encodeURIComponent('item/+2=')}`
    );
  });
});
