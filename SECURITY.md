# Security Policy

## Supported Versions

| Version | Supported          |
| ------- | ------------------ |
| 1.x.x   | Yes                |

## Reporting a Vulnerability

We take security seriously. If you discover a security vulnerability, please report it responsibly.

### How to Report

1. **Do not** open a public GitHub issue for security vulnerabilities
2. Email the maintainers directly or use GitHub's private vulnerability reporting feature
3. Include as much detail as possible:
   - Description of the vulnerability
   - Steps to reproduce
   - Potential impact
   - Suggested fix (if any)

### What to Expect

- Acknowledgment within 48 hours
- Regular updates on progress
- Credit in the security advisory (if desired)

## Security Best Practices for Users

When deploying this MCP server:

1. **Never commit secrets**: Keep `.env` files out of version control
2. **Use HTTPS in production**: Always use TLS for production deployments
3. **Rotate credentials**: Regularly rotate Azure AD client secrets
4. **Principle of least privilege**: Only request necessary Microsoft Graph permissions
5. **Monitor access**: Review Azure AD sign-in logs regularly
6. **Keep dependencies updated**: Run `npm audit` regularly

## Security Features

This server implements several security measures:

- OAuth 2.1 with PKCE (S256) for authentication
- AES-256-GCM encryption for session tokens
- Helmet.js for HTTP security headers
- Rate limiting to prevent abuse
- Input validation with Zod schemas
- No storage of plaintext credentials

## Scope

This security policy covers the m365-mcp-server codebase. Security issues in dependencies should be reported to the respective maintainers.
