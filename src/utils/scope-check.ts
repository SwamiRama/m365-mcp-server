import { GRAPH_SCOPES } from './config.js';

/**
 * Check if granted token scopes include all required Graph API scopes.
 * Only checks https://graph.microsoft.com/* scopes (ignores openid/offline_access).
 */
export function hasRequiredGraphScopes(grantedScope: string): boolean {
  const granted = new Set(grantedScope.toLowerCase().split(/\s+/));
  return GRAPH_SCOPES
    .filter(s => s.startsWith('https://graph.microsoft.com/'))
    .every(s => granted.has(s.toLowerCase()));
}
