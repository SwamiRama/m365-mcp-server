import { describe, it, expect } from 'vitest';
import { hasRequiredGraphScopes } from '../src/utils/scope-check.js';

describe('hasRequiredGraphScopes', () => {
  it('should return true when all Graph scopes are present', () => {
    const granted = [
      'https://graph.microsoft.com/User.Read',
      'https://graph.microsoft.com/Mail.Read',
      'https://graph.microsoft.com/Mail.Read.Shared',
      'https://graph.microsoft.com/Files.Read.All',
      'https://graph.microsoft.com/Sites.Read.All',
      'https://graph.microsoft.com/Calendars.Read',
    ].join(' ');

    expect(hasRequiredGraphScopes(granted)).toBe(true);
  });

  it('should return false when a Graph scope is missing', () => {
    const granted = [
      'https://graph.microsoft.com/User.Read',
      'https://graph.microsoft.com/Mail.Read',
      // Missing: Mail.Read.Shared, Files.Read.All, Sites.Read.All, Calendars.Read
    ].join(' ');

    expect(hasRequiredGraphScopes(granted)).toBe(false);
  });

  it('should be case-insensitive', () => {
    const granted = [
      'https://graph.microsoft.com/user.read',
      'https://graph.microsoft.com/MAIL.READ',
      'https://graph.microsoft.com/mail.read.shared',
      'https://graph.microsoft.com/FILES.READ.ALL',
      'https://graph.microsoft.com/sites.read.all',
      'https://graph.microsoft.com/calendars.read',
    ].join(' ');

    expect(hasRequiredGraphScopes(granted)).toBe(true);
  });

  it('should ignore openid and offline_access (not checked)', () => {
    // Only Graph scopes matter; openid/offline_access can be absent
    const granted = [
      'https://graph.microsoft.com/User.Read',
      'https://graph.microsoft.com/Mail.Read',
      'https://graph.microsoft.com/Mail.Read.Shared',
      'https://graph.microsoft.com/Files.Read.All',
      'https://graph.microsoft.com/Sites.Read.All',
      'https://graph.microsoft.com/Calendars.Read',
    ].join(' ');

    expect(hasRequiredGraphScopes(granted)).toBe(true);
  });

  it('should handle extra scopes in the granted string', () => {
    const granted = [
      'https://graph.microsoft.com/User.Read',
      'https://graph.microsoft.com/Mail.Read',
      'https://graph.microsoft.com/Mail.Read.Shared',
      'https://graph.microsoft.com/Files.Read.All',
      'https://graph.microsoft.com/Sites.Read.All',
      'https://graph.microsoft.com/Calendars.Read',
      'https://graph.microsoft.com/SomeExtra.Scope',
      'openid',
      'offline_access',
    ].join(' ');

    expect(hasRequiredGraphScopes(granted)).toBe(true);
  });

  it('should return false for empty granted string', () => {
    expect(hasRequiredGraphScopes('')).toBe(false);
  });
});
