# Security Review Report - Office 365 MCP Server

**Review Date:** 2025-10-27
**Reviewer:** Claude Code Security Analysis
**Version Reviewed:** 1.0.0

## Executive Summary

This security review identified **3 Critical**, **4 High**, **4 Medium**, and **3 Low** severity vulnerabilities in the Office 365 MCP Server codebase. The most critical issues involve **injection vulnerabilities** in search operations and **HTML/script injection** in OneNote and Teams functionality. Immediate remediation is recommended for critical and high-severity findings.

**Overall Security Rating:** ⚠️ **MEDIUM-HIGH RISK** (requires immediate attention)

---

## Critical Vulnerabilities

### 1. 🔴 CRITICAL: OData Query Injection in Search Operations

**Affected Files:**
- `src/tools/onedrive.ts:170`
- `src/tools/outlook.ts:103`
- `src/tools/onenote.ts:143,247`

**Description:**
User-provided search queries are directly interpolated into API calls without sanitization or parameterization, allowing potential OData injection attacks.

**Vulnerable Code Examples:**
```typescript
// onedrive.ts:170
async searchItems(query: string): Promise<DriveItem[]> {
  const result = await this.graphClient
    .api(`${this.getUserPath()}/drive/root/search(q='${query}')`)  // ⚠️ INJECTION
    .select(['id', 'name', 'size', 'webUrl', 'folder', 'file'])
    .get();
  return result.value;
}

// outlook.ts:103
async searchEmails(searchQuery: string, top: number = 10): Promise<EmailMessage[]> {
  const result = await this.graphClient
    .api(`${this.getUserPath()}/messages`)
    .search(`"${searchQuery}"`)  // ⚠️ INJECTION
    .top(top)
    .get();
  return result.value;
}
```

**Risk:**
- Attackers could inject malicious OData query operators
- Bypass access controls or retrieve unauthorized data
- Potential information disclosure

**Recommendation:**
```typescript
// Sanitize user input before using in queries
function sanitizeSearchQuery(query: string): string {
  // Remove or escape OData special characters
  return query.replace(/['"()$&]/g, '');
}

// Or use parameterized queries if supported by the Graph client
async searchItems(query: string): Promise<DriveItem[]> {
  const sanitizedQuery = sanitizeSearchQuery(query);
  const result = await this.graphClient
    .api(`${this.getUserPath()}/drive/root/search`)
    .query({ q: sanitizedQuery })  // Use query parameters
    .select(['id', 'name', 'size', 'webUrl', 'folder', 'file'])
    .get();
  return result.value;
}
```

---

### 2. 🔴 CRITICAL: HTML/Script Injection in OneNote Content

**Affected Files:**
- `src/tools/onenote.ts:175-194` (createPage)
- `src/tools/onenote.ts:199-212` (appendToPage)

**Description:**
User-provided content and titles are directly embedded into HTML without sanitization, enabling Cross-Site Scripting (XSS) and HTML injection attacks.

**Vulnerable Code:**
```typescript
// onenote.ts:175-186
async createPage(sectionId: string, title: string, content: string): Promise<OneNotePage> {
  const htmlContent = `
<!DOCTYPE html>
<html>
  <head>
    <title>${title}</title>  // ⚠️ INJECTION - No sanitization
  </head>
  <body>
    ${content}  // ⚠️ INJECTION - No sanitization
  </body>
</html>`;
```

**Risk:**
- Stored XSS vulnerability
- Malicious scripts could execute when pages are viewed
- Potential session hijacking or data theft
- Could affect all users who view the compromised OneNote page

**Recommendation:**
```typescript
import { escape } from 'html-escaper'; // or similar library

async createPage(sectionId: string, title: string, content: string): Promise<OneNotePage> {
  // Sanitize user input
  const sanitizedTitle = escape(title);
  const sanitizedContent = sanitizeHtml(content, {
    allowedTags: ['p', 'br', 'strong', 'em', 'u', 'h1', 'h2', 'h3', 'ul', 'ol', 'li'],
    allowedAttributes: {}
  });

  const htmlContent = `
<!DOCTYPE html>
<html>
  <head>
    <title>${sanitizedTitle}</title>
  </head>
  <body>
    ${sanitizedContent}
  </body>
</html>`;
  // ... rest of code
}
```

---

### 3. 🔴 CRITICAL: Excel Range Address Injection

**Affected Files:**
- `src/tools/excel.ts:70`
- `src/tools/excel.ts:81`

**Description:**
User-provided range addresses are directly interpolated into API calls without validation.

**Vulnerable Code:**
```typescript
// excel.ts:70
async getRange(itemId: string, worksheetId: string, address: string): Promise<any> {
  const range = await this.graphClient
    .api(`${this.getUserPath()}/drive/items/${itemId}/workbook/worksheets/${worksheetId}/range(address='${address}')`)
    .get();  // ⚠️ Address not validated
  return range;
}
```

**Risk:**
- Injection of malicious Excel formulas
- Potential access to unintended cell ranges
- Formula injection attacks

**Recommendation:**
```typescript
function validateExcelAddress(address: string): boolean {
  // Only allow valid Excel range format (A1, A1:B10, etc.)
  const validPattern = /^[A-Z]+[0-9]+(:[A-Z]+[0-9]+)?$/i;
  return validPattern.test(address);
}

async getRange(itemId: string, worksheetId: string, address: string): Promise<any> {
  if (!validateExcelAddress(address)) {
    throw new Error('Invalid Excel range address format');
  }
  const range = await this.graphClient
    .api(`${this.getUserPath()}/drive/items/${itemId}/workbook/worksheets/${worksheetId}/range(address='${address}')`)
    .get();
  return range;
}
```

---

## High Severity Vulnerabilities

### 4. 🟠 HIGH: Insufficient Token Cache File Permissions

**Affected Files:**
- `src/auth.ts:96`

**Description:**
Token cache files are written to disk without explicit permission validation. While the default location (user home directory) is appropriate, there's no verification that the file has restrictive permissions.

**Vulnerable Code:**
```typescript
// auth.ts:96
private saveCachedTokens(accessToken: string, expiresOn: Date, account?: any): void {
  try {
    const cache: TokenCache = {
      accessToken,
      expiresOn: expiresOn.getTime(),
      userId: this.config.userPrincipalName || this.config.userId,
      account,
    };
    fs.writeFileSync(this.tokenCachePath, JSON.stringify(cache, null, 2));  // ⚠️ No permission check
    console.error(`[AUTH] Saved tokens to cache: ${this.tokenCachePath}`);
  } catch (error) {
    console.error(`[AUTH] Failed to save tokens to cache (${this.tokenCachePath}):`, error);
  }
}
```

**Risk:**
- Token cache files may be readable by other users on multi-user systems
- Leaked access tokens could allow unauthorized access
- Tokens provide full access to user's Office 365 resources

**Recommendation:**
```typescript
import * as fs from 'fs';

private saveCachedTokens(accessToken: string, expiresOn: Date, account?: any): void {
  try {
    const cache: TokenCache = {
      accessToken,
      expiresOn: expiresOn.getTime(),
      userId: this.config.userPrincipalName || this.config.userId,
      account,
    };

    // Write with restrictive permissions (0600 = read/write for owner only)
    const tempPath = this.tokenCachePath + '.tmp';
    fs.writeFileSync(tempPath, JSON.stringify(cache, null, 2), { mode: 0o600 });
    fs.renameSync(tempPath, this.tokenCachePath);

    // Verify permissions
    const stats = fs.statSync(this.tokenCachePath);
    if ((stats.mode & 0o777) !== 0o600) {
      console.error('[AUTH] WARNING: Token cache file has insecure permissions');
    }

    console.error(`[AUTH] Saved tokens to cache: ${this.tokenCachePath}`);
  } catch (error) {
    console.error(`[AUTH] Failed to save tokens to cache (${this.tokenCachePath}):`, error);
  }
}
```

---

### 5. 🟠 HIGH: Information Disclosure Through Error Messages

**Affected Files:**
- `src/index.ts:878-887`

**Description:**
Detailed error messages, including stack traces and internal details, are returned directly to users.

**Vulnerable Code:**
```typescript
// index.ts:878-887
} catch (error) {
  return {
    content: [
      {
        type: 'text',
        text: `Error: ${error instanceof Error ? error.message : String(error)}`,  // ⚠️ Full error exposed
      },
    ],
    isError: true,
  };
}
```

**Risk:**
- Exposes internal implementation details
- May reveal file paths, API endpoints, or system information
- Assists attackers in reconnaissance

**Recommendation:**
```typescript
} catch (error) {
  // Log detailed error internally
  console.error(`[MCP] Tool execution error for ${name}:`, error);

  // Return sanitized error message to user
  const userMessage = error instanceof Error
    ? getPublicErrorMessage(error)
    : 'An unexpected error occurred';

  return {
    content: [
      {
        type: 'text',
        text: `Error: ${userMessage}`,
      },
    ],
    isError: true,
  };
}

function getPublicErrorMessage(error: Error): string {
  // Map known error types to user-friendly messages
  if (error.message.includes('401') || error.message.includes('unauthorized')) {
    return 'Authentication failed. Please check your credentials.';
  }
  if (error.message.includes('403') || error.message.includes('forbidden')) {
    return 'Access denied. You do not have permission for this operation.';
  }
  if (error.message.includes('404')) {
    return 'The requested resource was not found.';
  }
  // Generic message for other errors
  return 'An error occurred while processing your request.';
}
```

---

### 6. 🟠 HIGH: No Input Validation on User-Provided IDs

**Affected Files:**
- All tool files (multiple instances)

**Description:**
Item IDs, message IDs, and other identifiers are not validated before being used in API calls.

**Risk:**
- Path traversal attempts
- Injection of special characters
- Unauthorized access to resources

**Recommendation:**
```typescript
// Add validation utility
function validateResourceId(id: string, resourceType: string): void {
  if (!id || typeof id !== 'string') {
    throw new Error(`Invalid ${resourceType} ID: must be a non-empty string`);
  }

  // Check for path traversal attempts
  if (id.includes('..') || id.includes('/') || id.includes('\\')) {
    throw new Error(`Invalid ${resourceType} ID: contains illegal characters`);
  }

  // Validate format (adjust regex based on actual ID format)
  const validIdPattern = /^[a-zA-Z0-9_-]+$/;
  if (!validIdPattern.test(id)) {
    throw new Error(`Invalid ${resourceType} ID format`);
  }
}

// Use in functions:
async getEmail(messageId: string): Promise<EmailMessage> {
  validateResourceId(messageId, 'message');
  const message = await this.graphClient
    .api(`${this.getUserPath()}/messages/${messageId}`)
    .get();
  return message;
}
```

---

### 7. 🟠 HIGH: HTML Content Injection in Teams Messages

**Affected Files:**
- `src/tools/teams.ts:103-114` (sendChannelMessage)
- `src/tools/teams.ts:119-136` (replyToMessage)
- `src/tools/teams.ts:199-210` (sendChatMessage)

**Description:**
Teams messages accept HTML content type without validation or sanitization.

**Risk:**
- XSS attacks in Teams channels
- Phishing through crafted HTML messages
- Compromised channel security

**Recommendation:**
```typescript
import { sanitizeHtml } from 'sanitize-html'; // Add dependency

async sendChannelMessage(
  teamId: string,
  channelId: string,
  content: string,
  contentType: 'html' | 'text' = 'text'
): Promise<TeamsMessage> {
  let sanitizedContent = content;

  if (contentType === 'html') {
    // Sanitize HTML to prevent XSS
    sanitizedContent = sanitizeHtml(content, {
      allowedTags: ['p', 'br', 'strong', 'em', 'u', 'a', 'ul', 'ol', 'li'],
      allowedAttributes: {
        'a': ['href']
      },
      allowedSchemes: ['http', 'https', 'mailto']
    });
  }

  const message = await this.graphClient
    .api(`/teams/${teamId}/channels/${channelId}/messages`)
    .post({
      body: {
        contentType,
        content: sanitizedContent,
      },
    });
  return message;
}
```

---

## Medium Severity Vulnerabilities

### 8. 🟡 MEDIUM: Insufficient Audit Logging

**Affected Files:**
- All tool files

**Description:**
Security-relevant operations (email sending, file deletion, data modification) lack comprehensive audit logging.

**Risk:**
- Inability to detect unauthorized access
- No forensic trail for incident investigation
- Compliance violations (depending on organization requirements)

**Recommendation:**
```typescript
// Add audit logging utility
class AuditLogger {
  static log(action: string, userId: string, resourceType: string, resourceId: string, success: boolean, details?: any) {
    const logEntry = {
      timestamp: new Date().toISOString(),
      action,
      userId,
      resourceType,
      resourceId,
      success,
      details
    };

    // Log to structured logging system
    console.error(`[AUDIT] ${JSON.stringify(logEntry)}`);

    // In production, send to SIEM or audit logging service
  }
}

// Use in sensitive operations:
async sendEmail(message: EmailMessage): Promise<void> {
  try {
    await this.graphClient.api(`${this.getUserPath()}/sendMail`).post(mailObject);
    AuditLogger.log('send_email', this.userId, 'email', 'new', true, {
      to: message.toRecipients.map(r => r.emailAddress.address),
      subject: message.subject
    });
  } catch (error) {
    AuditLogger.log('send_email', this.userId, 'email', 'new', false, { error: error.message });
    throw error;
  }
}
```

---

### 9. 🟡 MEDIUM: Type Safety Bypass with 'any' Casts

**Affected Files:**
- `src/index.ts` (multiple lines throughout tool handler)

**Description:**
Extensive use of `args as any` bypasses TypeScript's type checking.

**Vulnerable Code:**
```typescript
// index.ts:684-706
case 'outlook_list_emails':
  result = await outlook.listEmails(args as any);  // ⚠️ Type safety bypassed
  break;
case 'outlook_send_email': {
  const { subject, body, bodyType, toRecipients, ccRecipients, importance } = args as any;
  // ...
}
```

**Risk:**
- Runtime errors from invalid data types
- Potential security issues from unexpected input
- Difficult to catch bugs during development

**Recommendation:**
```typescript
// Define proper interfaces for all tool arguments
interface OutlookSendEmailArgs {
  subject: string;
  body: string;
  bodyType?: 'HTML' | 'Text';
  toRecipients: string[];
  ccRecipients?: string[];
  importance?: 'low' | 'normal' | 'high';
}

// Add validation function
function validateArgs<T>(args: unknown, schema: any): T {
  // Use a validation library like zod or joi
  // For now, simple type checking:
  if (typeof args !== 'object' || args === null) {
    throw new Error('Invalid arguments: expected object');
  }
  return args as T;
}

// Use in handler:
case 'outlook_send_email': {
  const validatedArgs = validateArgs<OutlookSendEmailArgs>(args, outlookSendEmailSchema);
  const { subject, body, bodyType, toRecipients, ccRecipients, importance } = validatedArgs;
  // ... rest of code
}
```

---

### 10. 🟡 MEDIUM: No Request Size Limits

**Affected Files:**
- `src/tools/onedrive.ts:71-95` (uploadFile)
- `src/index.ts` (email body, Excel data)

**Description:**
No size limits are enforced on file uploads, email bodies, or data payloads.

**Risk:**
- Denial of service through large uploads
- Excessive API costs
- Memory exhaustion

**Recommendation:**
```typescript
const MAX_FILE_SIZE = 100 * 1024 * 1024; // 100 MB
const MAX_EMAIL_BODY_SIZE = 1024 * 1024; // 1 MB
const MAX_EXCEL_CELLS = 10000;

async uploadFile(fileName: string, content: Buffer | string, parentFolderId?: string): Promise<DriveItem> {
  let uploadContent: Buffer;
  if (typeof content === 'string') {
    uploadContent = Buffer.from(content, 'base64');
  } else {
    uploadContent = content;
  }

  // Check size limit
  if (uploadContent.length > MAX_FILE_SIZE) {
    throw new Error(`File size exceeds maximum allowed size of ${MAX_FILE_SIZE / (1024*1024)} MB`);
  }

  // ... rest of upload logic
}

async sendEmail(message: EmailMessage): Promise<void> {
  if (message.body.content.length > MAX_EMAIL_BODY_SIZE) {
    throw new Error(`Email body exceeds maximum size of ${MAX_EMAIL_BODY_SIZE / 1024} KB`);
  }
  // ... rest of code
}
```

---

### 11. 🟡 MEDIUM: Missing Rate Limiting

**Affected Files:**
- All tool files

**Description:**
No client-side rate limiting is implemented. The README acknowledges this limitation.

**Risk:**
- API throttling by Microsoft Graph
- Service disruption
- Potential abuse

**Recommendation:**
```typescript
import Bottleneck from 'bottleneck'; // Add dependency

class RateLimiter {
  private limiter: Bottleneck;

  constructor() {
    this.limiter = new Bottleneck({
      minTime: 100, // Minimum time between requests (ms)
      maxConcurrent: 10, // Maximum concurrent requests
      reservoir: 100, // Number of requests allowed
      reservoirRefreshAmount: 100,
      reservoirRefreshInterval: 60 * 1000, // Refill every minute
    });
  }

  async schedule<T>(fn: () => Promise<T>): Promise<T> {
    return this.limiter.schedule(fn);
  }
}

// Use in GraphAuthProvider:
export class GraphAuthProvider {
  private rateLimiter: RateLimiter;

  constructor(config: GraphConfig) {
    // ... existing code
    this.rateLimiter = new RateLimiter();
  }

  async getGraphClient(): Promise<Client> {
    const accessToken = await this.getAccessToken();

    return Client.init({
      authProvider: (done) => {
        done(null, accessToken);
      },
      middleware: [
        // Add rate limiting middleware
        {
          execute: async (context, next) => {
            return this.rateLimiter.schedule(() => next());
          }
        }
      ]
    });
  }
}
```

---

## Low Severity Issues

### 12. 🔵 LOW: Credentials in Environment Variables

**Affected Files:**
- `.env.example`
- Documentation

**Description:**
While properly documented, storing client secrets in environment variables can expose them through process listings.

**Recommendation:**
- Document Azure Key Vault integration for production deployments
- Add option to read secrets from secure secret management systems
- Implement automatic secret rotation

```typescript
// Example: Support for Azure Key Vault
import { SecretClient } from "@azure/keyvault-secrets";

async function getClientSecret(): Promise<string> {
  if (process.env.KEY_VAULT_URL) {
    const client = new SecretClient(
      process.env.KEY_VAULT_URL,
      new DefaultAzureCredential()
    );
    const secret = await client.getSecret("office365-client-secret");
    return secret.value!;
  }
  return process.env.CLIENT_SECRET || '';
}
```

---

### 13. 🔵 LOW: No Explicit TLS Configuration

**Affected Files:**
- All files making HTTP requests

**Description:**
No explicit TLS/SSL configuration or certificate validation settings.

**Recommendation:**
```typescript
// Ensure HTTPS is enforced and modern TLS versions are used
import * as https from 'https';

const httpsAgent = new https.Agent({
  minVersion: 'TLSv1.2',
  rejectUnauthorized: true, // Ensure certificate validation
});

// Use in Graph Client if needed
```

---

### 14. 🔵 LOW: Aggressive Token Expiry Buffer

**Affected Files:**
- `src/auth.ts:139,178,216`

**Description:**
5-minute buffer for token expiry might be too aggressive, leading to unnecessary token refreshes.

**Risk:**
- Increased API calls to authentication endpoints
- Potential rate limiting issues

**Recommendation:**
```typescript
// Make buffer configurable
const TOKEN_EXPIRY_BUFFER_MS = parseInt(process.env.TOKEN_EXPIRY_BUFFER_MS || '300000'); // 5 min default

this.tokenExpiry = new Date(response.expiresOn!.getTime() - TOKEN_EXPIRY_BUFFER_MS);
```

---

## Positive Security Practices

The following security best practices were observed:

✅ **Proper .gitignore Configuration**
- `.env`, `.token-cache.json`, and other sensitive files are excluded from version control

✅ **No Known Dependency Vulnerabilities**
- `npm audit` returned clean results with 0 vulnerabilities

✅ **Use of Official Microsoft Libraries**
- MSAL Node for authentication
- Microsoft Graph Client for API access

✅ **Secure Authentication Flows**
- Supports both delegated and app-only authentication
- Proper use of OAuth 2.0 device code flow
- Token caching with expiry handling

✅ **Security Documentation**
- README includes security considerations section
- Proper documentation of authentication requirements
- Clear instructions for permission configuration

✅ **Environment-Based Configuration**
- No hardcoded credentials
- Proper use of environment variables
- Example configuration file provided

✅ **Least Privilege Principle**
- Documentation recommends granting only necessary permissions

---

## Remediation Priority

### Immediate (Critical - Fix within 1 week):
1. ✅ Fix OData query injection in search operations
2. ✅ Sanitize HTML content in OneNote page creation
3. ✅ Validate Excel range addresses

### High Priority (Fix within 2-4 weeks):
4. ✅ Improve token cache file permissions
5. ✅ Sanitize error messages
6. ✅ Add input validation for resource IDs
7. ✅ Sanitize Teams HTML messages

### Medium Priority (Fix within 1-2 months):
8. ✅ Implement audit logging
9. ✅ Replace 'any' casts with proper types
10. ✅ Add request size limits
11. ✅ Implement rate limiting

### Low Priority (Address as time permits):
12. ✅ Document Azure Key Vault integration
13. ✅ Add explicit TLS configuration
14. ✅ Make token expiry buffer configurable

---

## Testing Recommendations

After implementing fixes, perform the following security tests:

1. **Input Validation Testing**
   - Test all search functions with special characters: `' OR 1=1--`, `"; DROP TABLE--`
   - Test OneNote content with `<script>alert('XSS')</script>`
   - Test Excel addresses with invalid formats

2. **Authentication Testing**
   - Verify token expiry and refresh
   - Test with expired tokens
   - Verify token cache file permissions (should be 0600)

3. **Authorization Testing**
   - Attempt to access other users' resources
   - Test with minimal permissions

4. **Error Handling Testing**
   - Trigger various error conditions
   - Verify error messages don't leak sensitive info

5. **Rate Limiting Testing**
   - Send rapid successive requests
   - Verify proper throttling behavior

---

## Compliance Considerations

Depending on your organization's requirements, consider:

- **GDPR**: Implement data retention policies, user data export/deletion
- **HIPAA**: Add encryption at rest for token cache, audit logging
- **SOC 2**: Comprehensive audit logging, access controls
- **ISO 27001**: Document security controls, incident response procedures

---

## Conclusion

The Office 365 MCP Server has a solid foundation with proper use of Microsoft's official libraries and good documentation. However, **critical injection vulnerabilities must be addressed immediately** before production use. The recommended fixes are straightforward to implement and will significantly improve the security posture.

**Recommended Next Steps:**
1. Address all critical vulnerabilities immediately
2. Implement input validation and output sanitization across all tools
3. Add comprehensive audit logging
4. Conduct security testing after remediation
5. Consider a third-party security audit before production deployment

---

**Report Version:** 1.0
**Last Updated:** 2025-10-27
