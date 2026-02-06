# m365-mcp-server

A production-ready **MCP (Model Context Protocol) server** for Microsoft 365, providing secure access to Email, SharePoint, and OneDrive through Azure AD/Entra ID authentication with OAuth 2.1 + PKCE.

## Features

- **Email Access**: List folders, search messages, read email content
- **SharePoint/OneDrive**: Browse sites, drives, folders, and read file content
- **Document Parsing**: Extracts readable text from PDF, Word, Excel, PowerPoint, CSV, and HTML files
- **OAuth 2.1 + PKCE**: Secure authentication via Azure AD/Entra ID
- **Delegated Permissions**: Users access only their authorized content
- **Open WebUI Compatible**: Works with native MCP or MCPO proxy
- **Production Ready**: Docker support, security hardening, structured audit logging
- **Token Revocation**: RFC 7009 compliant token revocation endpoint

## Quick Start

### 1. Azure AD Setup

Follow [docs/entra-app-registration.md](docs/entra-app-registration.md) to create an Azure AD app registration with these permissions:
- `openid`, `offline_access` (OIDC)
- `User.Read`, `Mail.Read`, `Files.Read` (Microsoft Graph)

### 2. Configuration

Create a `.env` file:

```env
# Azure AD / Entra ID (required)
AZURE_CLIENT_ID=your-client-id
AZURE_CLIENT_SECRET=your-client-secret
AZURE_TENANT_ID=your-tenant-id

# Server
MCP_SERVER_PORT=3000
MCP_SERVER_BASE_URL=http://localhost:3000
SESSION_SECRET=$(openssl rand -hex 32)

# Optional
LOG_LEVEL=info
REDIS_URL=redis://localhost:6379

# OAuth signing keys (required in production)
# OAUTH_SIGNING_KEY_PRIVATE=<base64-encoded PEM>
# OAUTH_SIGNING_KEY_PUBLIC=<base64-encoded PEM>
```

### 3. Run Locally

```bash
# Install dependencies
npm install

# Development mode
npm run dev

# Production build
npm run build
npm start
```

### 4. Authenticate

1. Open `http://localhost:3000/auth/login` in a browser
2. Sign in with your Microsoft 365 account
3. Note the session ID returned after login

## Docker Deployment

### Basic

```bash
cd docker
docker-compose up -d m365-mcp-server redis
```

### With Open WebUI

```bash
cd docker
docker-compose --profile with-webui up -d
```

### With MCPO Proxy

```bash
cd docker
docker-compose --profile with-mcpo up -d
```

## Open WebUI Integration

### Option A: Native MCP (Recommended)

1. In Open WebUI, go to **Admin Settings** > **Tools**
2. Add MCP Server:
   ```json
   {
     "url": "http://localhost:3000/mcp",
     "transport": "streamable-http"
   }
   ```
3. Complete OAuth login when prompted

### Option B: Via MCPO Proxy

1. Start MCPO with the provided config:
   ```bash
   mcpo --config docker/mcpo-config.json --port 8000
   ```
2. In Open WebUI, add as OpenAPI Tool:
   ```
   http://localhost:8000/openapi.json
   ```

## MCP Tools

### Email Tools

| Tool | Description |
|------|-------------|
| `mail_list_messages` | List messages with optional filters |
| `mail_get_message` | Get full message details |
| `mail_list_folders` | List mail folders |

### SharePoint/OneDrive Tools

| Tool | Description |
|------|-------------|
| `sp_list_sites` | Search and list SharePoint sites |
| `sp_list_drives` | List drives (OneDrive/document libraries) |
| `sp_list_children` | List folder contents |
| `sp_get_file` | Get file content with automatic document parsing (PDF, Word, Excel, PowerPoint → text). Max 10MB |

## API Endpoints

| Endpoint | Method | Description |
|----------|--------|-------------|
| `/health` | GET | Health check |
| `/auth/login` | GET | Initiate OAuth login |
| `/auth/callback` | GET | OAuth callback |
| `/auth/logout` | GET | Logout and revoke session |
| `/auth/status` | GET | Check authentication status |
| `/revoke` | POST | Token revocation (RFC 7009) |
| `/mcp` | POST | MCP JSON-RPC endpoint |
| `/mcp` | GET | MCP SSE stream endpoint |
| `/mcp` | DELETE | Terminate MCP session |

## Security

- **OAuth 2.1 + PKCE**: Required for all authentication flows
- **Delegated Permissions Only**: No app-only access, read-only Graph scopes
- **Token Encryption**: AES-256-GCM encryption for session tokens at rest
- **PII Redaction**: Sensitive data (tokens, emails, secrets) filtered from logs
- **Structured Audit Logging**: Security events logged with correlation IDs
- **Rate Limiting**: 100 req/min general, 5/hour for client registration
- **Security Headers**: HSTS, CSP (no unsafe-inline), Permissions-Policy, X-Frame-Options via Helmet
- **Input Validation**: Zod schemas + regex validation for all Graph API resource IDs
- **DCR Protection**: Redirect URI pattern whitelist, rate limiting, audit logging
- **Production Enforcement**: Config validation requires Redis, HTTPS, persistent signing keys

See [docs/security/threat-model.md](docs/security/threat-model.md) for full security analysis.

## Architecture

```
┌─────────────────────────────────────────────────────────────┐
│                      Open WebUI / Client                     │
└─────────────────────────────┬───────────────────────────────┘
                              │ MCP Protocol (Streamable HTTP)
                              ▼
┌─────────────────────────────────────────────────────────────┐
│                    m365-mcp-server                           │
│  ┌─────────────┐  ┌─────────────┐  ┌──────────────────────┐ │
│  │ OAuth 2.1   │  │ MCP Handler │  │ Microsoft Graph      │ │
│  │ + PKCE      │  │ (JSON-RPC)  │  │ Client               │ │
│  └──────┬──────┘  └─────────────┘  └──────────┬───────────┘ │
└─────────│────────────────────────────────────│──────────────┘
          │                                    │
          ▼                                    ▼
┌──────────────────────┐            ┌─────────────────────────┐
│  Azure AD / Entra ID │            │   Microsoft Graph API   │
│  (Authorization)     │            │   (Data Access)         │
└──────────────────────┘            └─────────────────────────┘
```

## Environment Variables

| Variable | Required | Default | Description |
|----------|----------|---------|-------------|
| `AZURE_CLIENT_ID` | Yes | - | Azure AD app client ID |
| `AZURE_CLIENT_SECRET` | Yes | - | Azure AD app client secret |
| `AZURE_TENANT_ID` | Yes | - | Azure AD tenant ID |
| `SESSION_SECRET` | Yes | - | Session encryption key (32+ chars) |
| `MCP_SERVER_PORT` | No | 3000 | Server port |
| `MCP_SERVER_BASE_URL` | No | http://localhost:3000 | Public URL (HTTPS required in production) |
| `REDIS_URL` | Prod | - | Redis URL (required in production) |
| `OAUTH_SIGNING_KEY_PRIVATE` | Prod | - | RSA private key PEM (required in production) |
| `OAUTH_SIGNING_KEY_PUBLIC` | Prod | - | RSA public key PEM (required in production) |
| `OAUTH_ALLOWED_REDIRECT_PATTERNS` | No | - | Comma-separated URI patterns for DCR |
| `LOG_LEVEL` | No | info | Log level (trace/debug/info/warn/error) |
| `NODE_ENV` | No | development | Environment mode |
| `FILE_PARSE_TIMEOUT_MS` | No | 30000 | Document parsing timeout |
| `FILE_PARSE_MAX_OUTPUT_KB` | No | 500 | Max parsed text output size |

## Development

```bash
# Install dependencies
npm install

# Run tests
npm test

# Run tests with coverage
npm run test:coverage

# Lint
npm run lint

# Type check
npm run typecheck

# Build
npm run build
```

## MCP Registry

This server is published to the MCP Registry. Add to your MCP client:

```json
{
  "mcpServers": {
    "m365": {
      "command": "npx",
      "args": ["-y", "@anthropic/m365-mcp-server"],
      "env": {
        "AZURE_CLIENT_ID": "your-client-id",
        "AZURE_CLIENT_SECRET": "your-client-secret",
        "AZURE_TENANT_ID": "your-tenant-id",
        "SESSION_SECRET": "your-session-secret"
      }
    }
  }
}
```

## Documentation

- [Azure AD App Registration](docs/entra-app-registration.md)
- [Architecture Decision Record](docs/adr/0001-auth-openwebui.md)
- [Security Threat Model](docs/security/threat-model.md)

## Supported Document Formats

`sp_get_file` automatically extracts readable text from these formats:

| Format | Extensions | Library |
|--------|-----------|---------|
| PDF | `.pdf` | pdf-parse |
| Word | `.docx`, `.doc` | mammoth |
| Excel | `.xlsx`, `.xls` | exceljs |
| PowerPoint | `.pptx`, `.ppt` | Built-in ZIP/XML |
| CSV | `.csv` | Built-in |
| HTML | `.html` | Built-in |

Other binary formats are returned as base64. Parsed text output is limited to 500KB by default.

## Known Limitations

- Maximum file download size: 10MB
- Parsed text output capped at 500KB (configurable via `FILE_PARSE_MAX_OUTPUT_KB`)
- SharePoint site listing requires search query (Graph API limitation)
- Refresh tokens limited to 24 hours for SPA scenarios
- No write operations (read-only by design)
- Access tokens (JWTs) are stateless and cannot be directly revoked (expire naturally)

## License

MIT

## Contributing

1. Fork the repository
2. Create a feature branch
3. Submit a pull request

Please ensure all tests pass and the code follows the existing style.
