# Azure AD / Entra ID App Registration Guide

This document provides step-by-step instructions for registering the m365-mcp-server application in Microsoft Entra ID (formerly Azure AD).

## Prerequisites

- Azure account with an active subscription
- **Application Developer** role (minimum) or **Global Administrator**
- Access to [Microsoft Entra admin center](https://entra.microsoft.com)

## Step 1: Create App Registration

1. Sign in to [Microsoft Entra admin center](https://entra.microsoft.com)
2. Navigate to **Identity** → **Applications** → **App registrations**
3. Click **+ New registration**
4. Configure the registration:

| Field | Value |
|-------|-------|
| **Name** | `m365-mcp-server` |
| **Supported account types** | See table below |
| **Redirect URI** | Web: `http://localhost:3000/auth/callback` |

### Supported Account Types

| Option | Tenant Value | Use Case |
|--------|--------------|----------|
| **Accounts in this organizational directory only** (Default) | `{tenant-id}` | Single organization deployment |
| **Accounts in any organizational directory** | `organizations` | Multi-tenant SaaS deployment |
| **Accounts in any organizational directory and personal Microsoft accounts** | `common` | Broadest compatibility |

> **Recommendation**: Start with **single-tenant** for security. Configure multi-tenant only if required.

5. Click **Register**
6. **Record these values** from the Overview page:
   - **Application (client) ID**: `xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx`
   - **Directory (tenant) ID**: `xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx`

## Step 2: Add Redirect URIs

1. Go to **Authentication** in the left menu
2. Under **Platform configurations**, click **Add a platform** → **Web**
3. Add the following redirect URIs:

### Development
```
http://localhost:3000/auth/callback
http://localhost:3000/auth/silent-callback
```

### Production (replace with your domain)
```
https://your-domain.com/auth/callback
https://your-domain.com/auth/silent-callback
```

### Open WebUI Integration (if using MCPO)
```
http://localhost:8080/callback
https://your-openwebui-domain.com/callback
```

4. Under **Implicit grant and hybrid flows**:
   - ☐ Access tokens (unchecked - we use authorization code flow)
   - ☐ ID tokens (unchecked - we use authorization code flow)

5. Under **Advanced settings**:
   - **Allow public client flows**: No (we use confidential client)

6. Click **Save**

## Step 3: Configure Client Secret

1. Go to **Certificates & secrets** in the left menu
2. Under **Client secrets**, click **+ New client secret**
3. Configure:
   - **Description**: `m365-mcp-server-secret`
   - **Expires**: 24 months (recommended) or custom
4. Click **Add**
5. **IMMEDIATELY copy the secret value** - it won't be shown again!

> **Security Note**: Store this secret securely. Never commit to source control.

```bash
# Example: Store in environment variable
export AZURE_CLIENT_SECRET="your-secret-value-here"
```

## Step 4: Configure API Permissions

1. Go to **API permissions** in the left menu
2. Click **+ Add a permission** → **Microsoft Graph** → **Delegated permissions**
3. Add the following permissions:

### Required Permissions

| Permission | Category | Description | Admin Consent |
|------------|----------|-------------|---------------|
| `openid` | OpenID Connect | Sign users in | No |
| `offline_access` | OpenID Connect | Maintain access (refresh tokens) | No |
| `User.Read` | User | Read user profile | No |
| `Mail.Read` | Mail | Read user mail | No |
| `Mail.Read.Shared` | Mail | Read shared mailbox mail | **Yes** |
| `Files.Read` | Files | Read user files | No |
| `Sites.Read.All` | Sites | Read SharePoint sites | No |
| `Calendars.Read` | Calendar | Read user calendars and events | No |

### Adding Permissions

Search and add each permission:
1. Search: `openid` → Check → Add permissions
2. Search: `offline_access` → Check → Add permissions
3. Search: `User.Read` → Check → Add permissions
4. Search: `Mail.Read` → Check → Add permissions
5. Search: `Mail.Read.Shared` → Check → Add permissions
6. Search: `Files.Read` → Check → Add permissions
7. Search: `Sites.Read.All` → Check → Add permissions
8. Search: `Calendars.Read` → Check → Add permissions

### Permission Summary

After adding, your permissions should look like:

| API | Permission | Type | Status |
|-----|------------|------|--------|
| Microsoft Graph | Files.Read | Delegated | ⏳ Not granted |
| Microsoft Graph | Mail.Read | Delegated | ⏳ Not granted |
| Microsoft Graph | Mail.Read.Shared | Delegated | ⏳ Not granted |
| Microsoft Graph | Sites.Read.All | Delegated | ⏳ Not granted |
| Microsoft Graph | Calendars.Read | Delegated | ⏳ Not granted |
| Microsoft Graph | offline_access | Delegated | ⏳ Not granted |
| Microsoft Graph | openid | Delegated | ⏳ Not granted |
| Microsoft Graph | User.Read | Delegated | ✅ Granted |

### Admin Consent (Required)

Admin consent is **required** for `Mail.Read.Shared` (shared mailbox access). For single-tenant deployments, grant admin consent for all users:

1. Click **Grant admin consent for {organization}**
2. Confirm by clicking **Yes**

> **Note**: Without admin consent, `Mail.Read.Shared` will not be available and shared mailbox access will fail. Other permissions can work with individual user consent.

## Step 5: Verify Configuration

### Configuration Checklist

- [ ] Application (client) ID recorded
- [ ] Directory (tenant) ID recorded
- [ ] Client secret created and stored securely
- [ ] Redirect URIs configured for all environments
- [ ] All required API permissions added
- [ ] (Optional) Admin consent granted

### Export Configuration

Create your `.env` file:

```env
# Azure AD / Entra ID Configuration
AZURE_CLIENT_ID=xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx
AZURE_CLIENT_SECRET=your-client-secret
AZURE_TENANT_ID=xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx

# For multi-tenant, use:
# AZURE_TENANT_ID=common
# or
# AZURE_TENANT_ID=organizations
```

## OAuth 2.1 Endpoints

Based on your tenant configuration:

### Single-Tenant

```
Authorization: https://login.microsoftonline.com/{tenant-id}/oauth2/v2.0/authorize
Token:         https://login.microsoftonline.com/{tenant-id}/oauth2/v2.0/token
```

### Multi-Tenant

```
Authorization: https://login.microsoftonline.com/common/oauth2/v2.0/authorize
Token:         https://login.microsoftonline.com/common/oauth2/v2.0/token
```

## Scopes Reference

When requesting authorization, use these scope strings:

```
openid offline_access https://graph.microsoft.com/User.Read https://graph.microsoft.com/Mail.Read https://graph.microsoft.com/Mail.Read.Shared https://graph.microsoft.com/Files.Read https://graph.microsoft.com/Sites.Read.All https://graph.microsoft.com/Calendars.Read
```

Or in URL-encoded format:
```
scope=openid%20offline_access%20https%3A%2F%2Fgraph.microsoft.com%2FUser.Read%20https%3A%2F%2Fgraph.microsoft.com%2FMail.Read%20https%3A%2F%2Fgraph.microsoft.com%2FMail.Read.Shared%20https%3A%2F%2Fgraph.microsoft.com%2FFiles.Read%20https%3A%2F%2Fgraph.microsoft.com%2FSites.Read.All%20https%3A%2F%2Fgraph.microsoft.com%2FCalendars.Read
```

## PKCE Requirements

This application uses **PKCE (Proof Key for Code Exchange)** as required by OAuth 2.1:

| Parameter | Requirement |
|-----------|-------------|
| `code_challenge_method` | `S256` (SHA-256) |
| `code_verifier` | 43-128 character random string |
| `code_challenge` | Base64URL(SHA256(code_verifier)) |

The application handles PKCE automatically. No additional configuration needed.

## Troubleshooting

### Common Errors

| Error | Cause | Solution |
|-------|-------|----------|
| `AADSTS50011` | Invalid redirect URI | Add exact URI to app registration |
| `AADSTS65001` | Consent required | User must consent or admin grants consent |
| `AADSTS7000218` | PKCE required | Ensure code_challenge is included |
| `AADSTS500011` | Resource not found | Check scope format (use full URI) |
| `AADSTS700016` | Application not found | Verify client_id |

### Verify App Registration

Use Microsoft Graph Explorer to test:

1. Go to [Graph Explorer](https://developer.microsoft.com/graph/graph-explorer)
2. Sign in with your account
3. Try: `GET https://graph.microsoft.com/v1.0/me`

## Security Best Practices

1. **Rotate client secrets** every 6-12 months
2. **Use Key Vault** for secret storage in production
3. **Monitor sign-in logs** in Entra ID
4. **Enable Conditional Access** policies as needed
5. **Review permissions** regularly - remove unused permissions
6. **Use certificate credentials** for production (instead of secrets)

## Additional Resources

- [Microsoft identity platform documentation](https://learn.microsoft.com/en-us/entra/identity-platform/)
- [Microsoft Graph permissions reference](https://learn.microsoft.com/en-us/graph/permissions-reference)
- [OAuth 2.0 authorization code flow](https://learn.microsoft.com/en-us/entra/identity-platform/v2-oauth2-auth-code-flow)
- [Entra ID app registration quickstart](https://learn.microsoft.com/en-us/entra/identity-platform/quickstart-register-app)
