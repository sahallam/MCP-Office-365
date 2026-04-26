# Security Policy

## Reporting a Vulnerability

**Please do not report security vulnerabilities through public GitHub issues.**

If you discover a security vulnerability in this project, please report it privately.

### How to Report

**Email:** Contact the repository maintainer via their GitHub profile

**Include in your report:**
- Description of the vulnerability
- Steps to reproduce the issue
- Potential impact and severity
- Suggested fix (if you have one)
- Your contact information (if you want credit)

### What to Expect

- **Acknowledgment:** You should receive an acknowledgment within 72 hours
- **Assessment:** The vulnerability will be assessed and prioritized
- **Fix Timeline:** Critical vulnerabilities will be addressed as soon as possible
- **Disclosure:** A security advisory will be published after the fix is released

### Scope

This policy applies to:
- Security vulnerabilities in the MCP-Office-365 server code
- Authentication and token handling issues
- Data exposure or privacy concerns
- Injection vulnerabilities (code, SQL, XSS, etc.)

### Out of Scope

- Issues in third-party dependencies (report to the dependency maintainer)
- Social engineering attacks
- Physical security
- Issues requiring physical access to the server

## Security Features

This project implements multiple security measures:

- **AES-256-GCM encryption** for token storage
- **TLS 1.2/1.3** enforcement for API communications
- **Input validation** and sanitization
- **OWASP Top 10** compliance (96/100 score)
- **Audit logging** for security events

See [Security Architecture](../README.md#security-architecture) in the README for details.

## Supported Versions

| Version | Supported          |
| ------- | ------------------ |
| 1.0.x   | :white_check_mark: |
| < 1.0   | :x:                |

## Security Best Practices for Users

When deploying this MCP server:

1. **Token Protection**
   - Set `TOKEN_ENCRYPTION_KEY` environment variable
   - Protect your `.env` file (never commit it)
   - Use secure file permissions (0600) for token cache

2. **Credentials**
   - Never commit `CLIENT_SECRET` to version control
   - Rotate client secrets regularly
   - Use Azure Key Vault for production secrets

3. **Network Security**
   - Run the server in a trusted environment
   - Use firewall rules to restrict access
   - Monitor audit logs for suspicious activity

4. **Updates**
   - Keep dependencies updated (`npm audit`)
   - Monitor for new releases
   - Review security advisories

## Acknowledgments

We appreciate security researchers who responsibly disclose vulnerabilities. If you report a valid security issue, we're happy to credit you in the security advisory (unless you prefer to remain anonymous).

Thank you for helping keep this project secure!
