# Threat Model: m365-mcp-server

## Overview

This document outlines the security threats, mitigations, and residual risks for the m365-mcp-server application.

## System Boundaries

```
┌─────────────────────────────────────────────────────────────────────────────┐
│                            TRUST BOUNDARY                                    │
│                                                                             │
│  ┌──────────────┐    ┌───────────────────────┐    ┌────────────────────┐   │
│  │   MCP Client │    │   m365-mcp-server     │    │   Redis (optional) │   │
│  │  (Open WebUI)│───▶│   - Auth Handler      │───▶│   - Session Store  │   │
│  └──────────────┘    │   - MCP Protocol      │    └────────────────────┘   │
│                      │   - Graph Client      │                              │
│                      └───────────┬───────────┘                              │
│                                  │                                          │
└──────────────────────────────────│──────────────────────────────────────────┘
                                   │
                                   │ HTTPS
                                   ▼
              ┌────────────────────────────────────────┐
              │           EXTERNAL SERVICES             │
              │  ┌────────────────┐  ┌──────────────┐  │
              │  │  Azure AD/     │  │  Microsoft   │  │
              │  │  Entra ID      │  │  Graph API   │  │
              │  └────────────────┘  └──────────────┘  │
              └────────────────────────────────────────┘
```

## Assets

| Asset | Sensitivity | Description |
|-------|-------------|-------------|
| Access Tokens | HIGH | OAuth tokens for Microsoft Graph API access |
| Refresh Tokens | HIGH | Long-lived tokens for token renewal |
| Session Data | MEDIUM | User session state including encrypted tokens |
| Client Secret | CRITICAL | Azure AD application credential |
| User Email Content | HIGH | Email messages accessed via Graph API |
| User Files | HIGH | OneDrive/SharePoint files accessed via Graph API |

## Threats and Mitigations

### 1. Authentication & Authorization

#### T1.1: Token Theft via Man-in-the-Middle
- **Risk**: HIGH
- **Attack Vector**: Attacker intercepts tokens during transmission
- **Mitigations**:
  - ✅ TLS required for all production traffic
  - ✅ HSTS headers enabled
  - ✅ Tokens never transmitted in URLs
- **Residual Risk**: LOW

#### T1.2: Session Hijacking
- **Risk**: HIGH
- **Attack Vector**: Attacker steals session cookie/ID
- **Mitigations**:
  - ✅ HttpOnly cookies prevent XSS theft
  - ✅ Secure flag requires HTTPS
  - ✅ SameSite=Lax prevents CSRF
  - ✅ Session IDs are cryptographically random (256-bit)
  - ✅ Session expiration (24 hours)
- **Residual Risk**: LOW

#### T1.3: CSRF Attacks
- **Risk**: MEDIUM
- **Attack Vector**: Malicious site triggers unintended actions
- **Mitigations**:
  - ✅ OAuth state parameter validation
  - ✅ SameSite cookies
  - ✅ CORS restricted to known origins
- **Residual Risk**: LOW

#### T1.4: OAuth Authorization Code Interception
- **Risk**: HIGH
- **Attack Vector**: Attacker intercepts authorization code
- **Mitigations**:
  - ✅ PKCE (S256) required for all flows
  - ✅ Code verifier never leaves server
  - ✅ Short-lived authorization codes
- **Residual Risk**: LOW

### 2. Data Protection

#### T2.1: Token Leakage via Logs
- **Risk**: HIGH
- **Attack Vector**: Tokens exposed in log files
- **Mitigations**:
  - ✅ PII redaction in all log output
  - ✅ Token values never logged
  - ✅ Structured logging with sensitive key filtering
- **Residual Risk**: LOW

#### T2.2: Token Exposure in Memory Dumps
- **Risk**: MEDIUM
- **Attack Vector**: Attacker extracts tokens from process memory
- **Mitigations**:
  - ✅ Tokens encrypted at rest in session store
  - ✅ Container runs as non-root user
  - ⚠️ In-memory tokens not encrypted (necessary for API calls)
- **Residual Risk**: MEDIUM (accepted for functionality)

#### T2.3: Sensitive Data in API Responses
- **Risk**: MEDIUM
- **Attack Vector**: Over-exposure of user data to LLM
- **Mitigations**:
  - ✅ Minimal data returned by default
  - ✅ Email body only on explicit request
  - ✅ File size limits (10MB max)
  - ✅ Preview text truncated (200 chars)
- **Residual Risk**: LOW

### 3. Injection Attacks

#### T3.1: OData Injection
- **Risk**: MEDIUM
- **Attack Vector**: Malicious OData filter manipulation
- **Mitigations**:
  - ✅ Input validation with Zod schemas
  - ✅ Microsoft Graph sanitizes OData queries
  - ⚠️ User-provided filters passed to Graph API
- **Residual Risk**: LOW (Graph API validates)

#### T3.2: Command Injection
- **Risk**: LOW
- **Attack Vector**: N/A - no shell execution
- **Mitigations**:
  - ✅ No shell execution in codebase
  - ✅ No dynamic code evaluation
- **Residual Risk**: NONE

### 4. Availability

#### T4.1: Denial of Service (Application)
- **Risk**: MEDIUM
- **Attack Vector**: Resource exhaustion via requests
- **Mitigations**:
  - ✅ Rate limiting (100 req/min default)
  - ✅ Request body size limit (1MB)
  - ✅ Timeout on Graph API calls (30s)
- **Residual Risk**: LOW

#### T4.2: Denial of Service (Graph API)
- **Risk**: MEDIUM
- **Attack Vector**: Trigger Graph API throttling
- **Mitigations**:
  - ✅ Retry logic with exponential backoff
  - ✅ Respect Retry-After headers
  - ✅ Per-user token isolation
- **Residual Risk**: LOW

### 5. Infrastructure

#### T5.1: Container Escape
- **Risk**: MEDIUM
- **Attack Vector**: Exploit container runtime
- **Mitigations**:
  - ✅ Minimal Alpine base image
  - ✅ Non-root user (UID 1001)
  - ✅ No privileged capabilities
  - ⚠️ Runtime security depends on host
- **Residual Risk**: LOW

#### T5.2: Dependency Vulnerabilities
- **Risk**: MEDIUM
- **Attack Vector**: Exploit vulnerable npm packages
- **Mitigations**:
  - ✅ Dependabot alerts enabled
  - ✅ Trivy scanning in CI pipeline
  - ✅ npm audit on install
  - ⚠️ Zero-day vulnerabilities possible
- **Residual Risk**: MEDIUM (accepted with monitoring)

#### T5.3: Secret Exposure
- **Risk**: CRITICAL
- **Attack Vector**: Client secret compromised
- **Mitigations**:
  - ✅ Secrets via environment variables only
  - ✅ .env files in .gitignore
  - ✅ No secrets in container image
  - ⚠️ Production should use Key Vault/Secrets Manager
- **Residual Risk**: LOW (with proper secret management)

### 6. Privilege Escalation

#### T6.1: Delegated to Application Permission Bypass
- **Risk**: HIGH
- **Attack Vector**: Access data beyond user permissions
- **Mitigations**:
  - ✅ Only delegated permissions used
  - ✅ No application-only flows implemented
  - ✅ Graph API enforces user permissions
- **Residual Risk**: NONE (by design)

#### T6.2: Cross-User Data Access
- **Risk**: HIGH
- **Attack Vector**: User A accesses User B's data
- **Mitigations**:
  - ✅ Each session has isolated token
  - ✅ Graph API enforces token audience
  - ✅ No shared token cache between users
- **Residual Risk**: NONE (enforced by Graph)

## Security Controls Summary

### Implemented Controls

| Control | Status | Notes |
|---------|--------|-------|
| PKCE | ✅ Required | S256 method |
| TLS | ✅ Required | Production only |
| Rate Limiting | ✅ Enabled | 100 req/min |
| Input Validation | ✅ Enabled | Zod schemas |
| PII Redaction | ✅ Enabled | Logs, errors |
| Session Encryption | ✅ Enabled | AES-256-GCM |
| Non-root Container | ✅ Enabled | UID 1001 |
| Security Headers | ✅ Enabled | Helmet.js |
| CORS | ✅ Restricted | Same-origin default |

### Recommended Production Controls

| Control | Status | Notes |
|---------|--------|-------|
| Key Vault Integration | ⏳ Documented | For secrets |
| WAF | ⏳ Recommended | Azure WAF/Cloudflare |
| Container Scanning | ⏳ CI Pipeline | Trivy |
| SIEM Integration | ⏳ Documented | Log forwarding |
| Penetration Testing | ⏳ Recommended | Before production |

## Compliance Considerations

### GDPR
- User email/files are personal data
- Data processed only with user consent (OAuth)
- No data stored beyond session lifetime
- Recommend: Data retention policy review

### Microsoft 365 Compliance
- Adheres to Microsoft Graph API terms
- Delegated permissions respect tenant policies
- Conditional Access policies honored
- Recommend: Review with tenant admin

## Incident Response

### Security Incident Indicators
1. Unusual token refresh patterns
2. High rate of 401/403 errors
3. Graph API throttling alerts
4. Session anomalies (multiple IPs)

### Response Actions
1. **Suspected Token Compromise**
   - Revoke user's tokens via Azure AD
   - Invalidate sessions in Redis/memory
   - Rotate client secret if app-level

2. **Suspected Session Hijacking**
   - Force logout affected user
   - Audit session access logs
   - Enable Conditional Access policies

## Review Schedule

| Review Type | Frequency | Owner |
|-------------|-----------|-------|
| Dependency Audit | Weekly (automated) | CI/CD |
| Threat Model Review | Quarterly | Security Team |
| Penetration Test | Annually | External Vendor |
| Access Control Review | Monthly | Admin |

---

*Last Updated: 2026-01-31*
*Version: 1.0.0*
