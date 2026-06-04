# M365 MCP Server - Project Context

## Projektübersicht
MCP Server für Microsoft 365, der als OAuth 2.1 Authorization Server fungiert und mit Open WebUI kompatibel ist.

## Architektur
```
Open WebUI (OAuth Client) → MCP Server (Auth + Resource Server) → Azure AD (IdP) → Microsoft Graph API
```

## Wichtige Dateien
- `src/oauth/` - OAuth 2.1 Implementation (DCR, PKCE, JWT)
- `src/graph/client.ts` - Microsoft Graph API Client
- `src/tools/mail.ts` - Mail Tools (list, get, folders, attachments; shared mailbox support)
- `src/tools/sharepoint.ts` - SharePoint/OneDrive Tools
- `src/tools/onedrive.ts` - Dedicated OneDrive Tools (personal drive)
- `src/tools/calendar.ts` - Calendar Tools (read-only)
- `src/utils/config.ts` - Konfiguration und Graph Scopes

## Deployment
- Docker Image: `ghcr.io/swamirama/m365-mcp-server:latest`
- Hosting: Azure Container Apps
- CI/CD: GitHub Actions (`.github/workflows/ci.yml`)

## Bekannte Besonderheiten
- Graph API `/sites` Endpoint benötigt `?search=*` Parameter um Sites zu listen
- OAuth Token Endpoint benötigt `express.urlencoded()` Middleware
- Hinter Reverse Proxy: `app.set('trust proxy', true)` erforderlich
- Alle 401-Responses MÜSSEN `WWW-Authenticate: Bearer resource_metadata="..."` Header setzen, damit MCP-Clients (Open WebUI) den OAuth-Server discovern und Re-Auth auslösen können

## Test-Scripts
- `scripts/test-local-oauth.sh` - Voller OAuth Flow Test
- `scripts/test-sharepoint-az.sh` - SharePoint Test mit Azure CLI

## Token-Lifetimes (seit 2026-06-04)
- Refresh-Tokens rotieren bei jedem Refresh-Grant; Lifetime via `OAUTH_REFRESH_TOKEN_LIFETIME_SECS` (Prod: 30d)
- Session-TTL via `SESSION_TTL_SECONDS` (sliding, Prod: 30d) — muss >= Refresh-Token-Lifetime sein, Boot-Validierung erzwingt das
- Rotation-Grace `OAUTH_REFRESH_TOKEN_REUSE_GRACE_SECS` (Default 60s) toleriert konkurrierende Refreshes (Open WebUI 2-6 Replicas); Reuse nach Grace revoked die Token-Familie (Log-Event `oauth.refresh_token_reuse`)
- MSAL-Cache + Sessions liegen AES-256-GCM-verschlüsselt in Redis (Key aus SESSION_SECRET)

## Erledigte Punkte
- [x] Persistente OAuth Signing Keys (Key Vault → OAUTH_SIGNING_KEY_PRIVATE/PUBLIC, via Terraform)
- [x] Redis für persistente Sessions in Produktion (REDIS_URL aus Key Vault)
