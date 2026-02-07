# ADR 0001: Authentication Architecture for Open WebUI Integration

## Status

**Accepted** - 2026-01-31

## Context

We are building an MCP server (`m365-mcp-server`) that provides access to Microsoft 365 resources (Email, SharePoint/OneDrive) through the Model Context Protocol. The server must:

1. Integrate with Open WebUI (v0.6.31+) for LLM-based tool access
2. Use OAuth 2.1 Authorization Code Flow with PKCE via Azure AD/Entra ID
3. Enforce least privilege through delegated permissions only
4. Support containerized deployment

### Research Findings (2026 Best Practices)

#### MCP Authorization/OAuth 2.1 (from MCP Spec 2025-11-25)
- MCP servers can act as **Resource Servers** validating tokens from external IdPs
- **Streamable HTTP transport** is preferred for OAuth scenarios (vs stdio)
- Session management via `MCP-Session-Id` header for stateful connections
- Protocol version header `MCP-Protocol-Version: 2025-11-25` required

**Source**: https://modelcontextprotocol.io/specification/2025-11-25/basic/transports

#### Open WebUI MCP Integration
- **Native MCP support** available since v0.6.31+
- **MCPO** (MCP-to-OpenAPI proxy) available as alternative
- `WEBUI_SECRET_KEY` must be persistent for session continuity
- MCP servers configured via Admin Settings > Tools/Connections

**Source**: https://docs.openwebui.com/features/mcp/

#### Microsoft Identity Platform
- **Authorization Code Flow + PKCE** mandatory for SPAs, recommended for all clients
- **Delegated permissions** enforce user-context access (Graph validates server-side)
- **Refresh tokens**: 90 days for confidential clients, 24h for SPAs
- Endpoints: `https://login.microsoftonline.com/{tenant}/oauth2/v2.0/authorize|token`

**Source**: https://learn.microsoft.com/en-us/entra/identity-platform/v2-oauth2-auth-code-flow

#### MCP Registry Standard
- `server.json` manifest with schema `https://static.modelcontextprotocol.io/schemas/2025-12-11/server.schema.json`
- Namespace format: `io.github.{username}/{server-name}` or `{domain}/{server-name}`
- Publishing via `mcp-publisher` CLI tool

**Source**: https://modelcontextprotocol.io/registry/quickstart

## Decision

### Architecture Choice

```
┌─────────────────────────────────────────────────────────────────────┐
│                         Open WebUI                                   │
│  ┌──────────────┐                                                   │
│  │   LLM Chat   │                                                   │
│  └──────┬───────┘                                                   │
│         │ MCP Protocol (Streamable HTTP)                            │
│         ▼                                                           │
│  ┌──────────────────────────────────────────────────────────────┐   │
│  │                    m365-mcp-server                            │   │
│  │  ┌─────────────┐  ┌─────────────┐  ┌──────────────────────┐  │   │
│  │  │ Auth Module │  │ MCP Handler │  │ Microsoft Graph      │  │   │
│  │  │ (OAuth 2.1) │  │ (JSON-RPC)  │  │ Client               │  │   │
│  │  └──────┬──────┘  └─────────────┘  └──────────┬───────────┘  │   │
│  └─────────│────────────────────────────────────│───────────────┘   │
└────────────│────────────────────────────────────│───────────────────┘
             │                                    │
             │ OAuth 2.1 + PKCE                   │ Delegated Access
             ▼                                    ▼
┌────────────────────────┐              ┌─────────────────────────┐
│  Azure AD / Entra ID   │              │   Microsoft Graph API   │
│  (Authorization Server)│              │   (Resource Server)     │
└────────────────────────┘              └─────────────────────────┘
```

### Key Decisions

#### 1. Transport: Streamable HTTP (Default)

**Rationale**:
- OAuth authentication requires HTTP endpoints for callback handling
- Enables proper session management with `MCP-Session-Id`
- Better suited for containerized deployment
- Supports both synchronous and SSE-streamed responses

**Alternative considered**: stdio transport
- Would require separate HTTP server for OAuth callback
- More complex architecture for auth scenarios
- Still supported for local development

#### 2. Integration Option: Option B (MCPO Proxy) as Default, Option A (Native) Documented

**Default: MCPO Proxy**

**Rationale**:
- Simpler OAuth handling - MCPO can manage token lifecycle
- OpenAPI compatibility provides better tooling
- Easier debugging via standard REST tools
- Works with Open WebUI versions that predate native MCP support

**Alternative: Native MCP**

**Also supported**:
- Direct MCP-over-HTTP integration
- Lower latency (no proxy hop)
- Documented for users preferring native approach

#### 3. OAuth Role Distribution

| Component | Role | Responsibility |
|-----------|------|----------------|
| Azure AD/Entra ID | Authorization Server | User authentication, token issuance |
| m365-mcp-server | Resource Server (partial) | Token validation, session management |
| Microsoft Graph | Resource Server | API access, permission enforcement |

**Key Insight**: We validate tokens for session binding, but Microsoft Graph enforces actual resource permissions. This ensures least privilege without complex local permission logic.

#### 4. Token Handling Strategy

```typescript
interface TokenStrategy {
  // Per-user session tokens (never shared)
  storage: 'encrypted-memory' | 'redis';

  // Token refresh before expiry
  refreshThreshold: '5 minutes before expiry';

  // On refresh failure
  failureAction: 'redirect-to-login';

  // Logging
  tokenLogging: 'NEVER log token values';
}
```

#### 5. App Registration: Single-Tenant Default

**Rationale**:
- Principle of least privilege
- Simpler admin consent model
- Configurable for multi-tenant if needed

```typescript
const tenantConfig = {
  default: 'single-tenant',  // {tenant-id}
  configurable: true,        // Can switch to 'organizations' or 'common'
};
```

### Delegated Permissions

| Permission | Purpose | Type | Admin Consent |
|------------|---------|------|---------------|
| `openid` | OIDC sign-in | OIDC | No |
| `offline_access` | Refresh tokens | OIDC | No |
| `User.Read` | User profile | Delegated | No |
| `Mail.Read` | Read user's mail | Delegated | No |
| `Mail.Read.Shared` | Read shared mailbox mail | Delegated | **Yes** |
| `Files.Read` | Read user's files | Delegated | No |
| `Sites.Read.All` | Read SharePoint sites | Delegated | No |

**Scope justification**:
- `Mail.Read.Shared` enables shared mailbox access — requires admin consent; Graph API enforces mailbox-level delegation server-side
- `Sites.Read.All` required for listing SharePoint sites via `/sites?search=` endpoint
- Not using `Files.Read.All` — `Files.Read` provides access to files user can access (OneDrive + shared)

## Consequences

### Positive

1. **Security**: Delegated permissions ensure users only access their authorized content
2. **Simplicity**: Single OAuth flow for all M365 resources
3. **Compliance**: Token handling follows OAuth 2.1 best practices (PKCE, no implicit flow)
4. **Flexibility**: Both MCPO and native MCP supported

### Negative

1. **Latency**: MCPO proxy adds ~10-20ms per request
2. **Complexity**: Two integration paths to maintain
3. **Dependency**: Requires Azure AD/Entra ID subscription

### Risks

1. **Token expiry during long operations**: Mitigated by proactive refresh
2. **Session hijacking**: Mitigated by secure session IDs and TLS
3. **Scope creep**: Documented scope justification prevents over-permissioning

## Implementation Notes

### Configuration Required

```env
# Azure AD/Entra ID
AZURE_CLIENT_ID=<from-app-registration>
AZURE_CLIENT_SECRET=<from-app-registration>
AZURE_TENANT_ID=<your-tenant-id>

# Server
MCP_SERVER_PORT=3000
MCP_SERVER_BASE_URL=https://your-domain.com
SESSION_SECRET=<cryptographically-secure-random>

# Optional
REDIS_URL=redis://localhost:6379  # For distributed sessions
```

### Security Checklist

- [ ] PKCE enforced on all authorization requests
- [ ] State parameter validated on callback
- [ ] Tokens stored encrypted, never logged
- [ ] TLS required in production
- [ ] Session IDs are cryptographically random
- [ ] CORS restricted to known origins
- [ ] Rate limiting enabled

## References

- [MCP Specification 2025-11-25](https://modelcontextprotocol.io/specification/2025-11-25/)
- [Microsoft Identity Platform OAuth 2.0](https://learn.microsoft.com/en-us/entra/identity-platform/v2-oauth2-auth-code-flow)
- [Open WebUI MCP Documentation](https://docs.openwebui.com/features/mcp/)
- [MCPO GitHub Repository](https://github.com/open-webui/mcpo)
- [MCP Registry Quickstart](https://modelcontextprotocol.io/registry/quickstart)
