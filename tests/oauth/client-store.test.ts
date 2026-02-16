import { describe, it, expect, vi } from 'vitest';
import { registerClient, deleteClient, hashRedirectUris } from '../../src/oauth/client-store.js';

// Mock redis to return null (uses MemoryClientStore)
vi.mock('../../src/utils/redis.js', () => ({
  getRedisClient: () => null,
}));

describe('client-store', () => {
  describe('hashRedirectUris', () => {
    it('should produce the same hash regardless of URI order', () => {
      const hash1 = hashRedirectUris(['https://a.com/cb', 'https://b.com/cb']);
      const hash2 = hashRedirectUris(['https://b.com/cb', 'https://a.com/cb']);
      expect(hash1).toBe(hash2);
    });

    it('should produce different hashes for different URIs', () => {
      const hash1 = hashRedirectUris(['https://a.com/cb']);
      const hash2 = hashRedirectUris(['https://b.com/cb']);
      expect(hash1).not.toBe(hash2);
    });
  });

  describe('registerClient (idempotent DCR)', () => {
    it('should return the same client_id for identical redirect_uris', async () => {
      const uris = [`https://idempotent-${Date.now()}.example.com/callback`];

      const first = await registerClient({
        client_name: 'Test App',
        redirect_uris: uris,
      });

      const second = await registerClient({
        client_name: 'Test App',
        redirect_uris: uris,
      });

      expect(second.client_id).toBe(first.client_id);
    });

    it('should return different client_ids for different redirect_uris', async () => {
      const ts = Date.now();

      const first = await registerClient({
        client_name: 'App A',
        redirect_uris: [`https://a-${ts}.example.com/callback`],
      });

      const second = await registerClient({
        client_name: 'App B',
        redirect_uris: [`https://b-${ts}.example.com/callback`],
      });

      expect(second.client_id).not.toBe(first.client_id);
    });

    it('should treat different URI order as the same client', async () => {
      const ts = Date.now();
      const uri1 = `https://order1-${ts}.example.com/cb`;
      const uri2 = `https://order2-${ts}.example.com/cb`;

      const first = await registerClient({
        client_name: 'Ordered App',
        redirect_uris: [uri1, uri2],
      });

      const second = await registerClient({
        client_name: 'Ordered App',
        redirect_uris: [uri2, uri1],
      });

      expect(second.client_id).toBe(first.client_id);
    });

    it('should return a new client_id after delete + re-register', async () => {
      const uris = [`https://delete-${Date.now()}.example.com/callback`];

      const first = await registerClient({
        client_name: 'Deletable App',
        redirect_uris: uris,
      });

      await deleteClient(first.client_id);

      const second = await registerClient({
        client_name: 'Deletable App',
        redirect_uris: uris,
      });

      expect(second.client_id).not.toBe(first.client_id);
    });

    it('should enforce public client (token_endpoint_auth_method: none, no secret)', async () => {
      const result = await registerClient({
        client_name: 'Public App',
        redirect_uris: [`https://public-${Date.now()}.example.com/callback`],
        token_endpoint_auth_method: 'client_secret_post', // request confidential
      });

      expect(result.token_endpoint_auth_method).toBe('none');
      expect(result.client_secret).toBeUndefined();
    });
  });
});
