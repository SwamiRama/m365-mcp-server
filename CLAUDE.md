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
- `src/tools/mail.ts` - Mail Tools (inkl. Shared Mailbox Support)
- `src/tools/sharepoint.ts` - SharePoint/OneDrive Tools
- `src/utils/config.ts` - Konfiguration und Graph Scopes

## Deployment
- Docker Image: `ghcr.io/swamirama/m365-mcp-server:latest`
- Hosting: Azure Container Apps
- CI/CD: GitHub Actions (`.github/workflows/ci.yml`)

## Bekannte Besonderheiten
- Graph API `/sites` Endpoint benötigt `?search=*` Parameter um Sites zu listen
- OAuth Token Endpoint benötigt `express.urlencoded()` Middleware
- Hinter Reverse Proxy: `app.set('trust proxy', true)` erforderlich

## Test-Scripts
- `scripts/test-local-oauth.sh` - Voller OAuth Flow Test
- `scripts/test-sharepoint-az.sh` - SharePoint Test mit Azure CLI

## Offene Punkte
- [ ] Persistente OAuth Signing Keys konfigurieren (OAUTH_SIGNING_KEY_PRIVATE/PUBLIC)
- [ ] Redis für persistente Sessions in Produktion
- [ ] Container Restart invalidiert alle Tokens (ephemere Keys)
