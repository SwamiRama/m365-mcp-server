import { describe, it, expect } from 'vitest';
import {
  generateCodeVerifier,
  generateCodeChallenge,
  generatePKCEPair,
  generateState,
  generateNonce,
} from '../../src/auth/pkce.js';
import crypto from 'crypto';

describe('PKCE', () => {
  describe('generateCodeVerifier', () => {
    it('should generate a string of valid length (43-128 chars)', () => {
      const verifier = generateCodeVerifier();
      expect(verifier.length).toBeGreaterThanOrEqual(43);
      expect(verifier.length).toBeLessThanOrEqual(128);
    });

    it('should generate only base64url characters', () => {
      const verifier = generateCodeVerifier();
      // Base64url uses A-Z, a-z, 0-9, -, _
      expect(verifier).toMatch(/^[A-Za-z0-9\-_]+$/);
    });

    it('should generate unique values', () => {
      const verifiers = new Set<string>();
      for (let i = 0; i < 100; i++) {
        verifiers.add(generateCodeVerifier());
      }
      expect(verifiers.size).toBe(100);
    });
  });

  describe('generateCodeChallenge', () => {
    it('should generate a valid S256 challenge', () => {
      const verifier = 'test-verifier-12345678901234567890123456789012';
      const challenge = generateCodeChallenge(verifier);

      // Manually compute expected challenge
      const expected = crypto
        .createHash('sha256')
        .update(verifier, 'ascii')
        .digest('base64url');

      expect(challenge).toBe(expected);
    });

    it('should generate consistent challenges for same verifier', () => {
      const verifier = generateCodeVerifier();
      const challenge1 = generateCodeChallenge(verifier);
      const challenge2 = generateCodeChallenge(verifier);

      expect(challenge1).toBe(challenge2);
    });

    it('should generate different challenges for different verifiers', () => {
      const challenge1 = generateCodeChallenge(generateCodeVerifier());
      const challenge2 = generateCodeChallenge(generateCodeVerifier());

      expect(challenge1).not.toBe(challenge2);
    });
  });

  describe('generatePKCEPair', () => {
    it('should return verifier, challenge, and method', () => {
      const pair = generatePKCEPair();

      expect(pair).toHaveProperty('codeVerifier');
      expect(pair).toHaveProperty('codeChallenge');
      expect(pair).toHaveProperty('codeChallengeMethod', 'S256');
    });

    it('should have matching verifier and challenge', () => {
      const pair = generatePKCEPair();
      const expectedChallenge = generateCodeChallenge(pair.codeVerifier);

      expect(pair.codeChallenge).toBe(expectedChallenge);
    });
  });

  describe('generateState', () => {
    it('should generate a 64-character hex string', () => {
      const state = generateState();
      expect(state).toHaveLength(64);
      expect(state).toMatch(/^[0-9a-f]+$/);
    });

    it('should generate unique values', () => {
      const states = new Set<string>();
      for (let i = 0; i < 100; i++) {
        states.add(generateState());
      }
      expect(states.size).toBe(100);
    });
  });

  describe('generateNonce', () => {
    it('should generate a base64url string', () => {
      const nonce = generateNonce();
      expect(nonce).toMatch(/^[A-Za-z0-9\-_]+$/);
    });

    it('should generate unique values', () => {
      const nonces = new Set<string>();
      for (let i = 0; i < 100; i++) {
        nonces.add(generateNonce());
      }
      expect(nonces.size).toBe(100);
    });
  });
});
