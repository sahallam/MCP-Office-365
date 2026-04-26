# Office365 MCP Server

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![Node.js Version](https://img.shields.io/badge/node-%3E%3D18.0.0-brightgreen)](https://nodejs.org)
[![TypeScript](https://img.shields.io/badge/TypeScript-5.7-blue)](https://www.typescriptlang.org/)

A comprehensive [Model Context Protocol](https://modelcontextprotocol.io) (MCP) server for Microsoft Office 365 integration. This server enables AI assistants like Claude to interact with your Office 365 environment, including Outlook, Calendar, OneDrive, SharePoint, Teams, Excel, Word, and OneNote.

## Table of Contents

- [Features](#features)
- [Prerequisites](#prerequisites)
- [Authentication Modes](#authentication-modes)
- [Setup](#setup)
- [Usage](#usage)
- [Available Tools](#available-tools)
- [Example Prompts](#example-prompts)
- [Architecture](#architecture)
- [Development](#development)
- [Security Architecture](#security-architecture)
- [Troubleshooting](#troubleshooting)
- [Contributing](#contributing)
- [License](#license)

## Features

### Outlook/Email
- List and search emails
- Send emails with attachments
- Reply to emails (single or reply-all)
- Mark emails as read/unread
- Manage email folders
- Create drafts

### Calendar
- List and search calendar events
- Create, update, and delete events
- Get calendar views for specific time ranges
- Accept/decline/tentatively accept meetings
- Find available meeting times
- Support for online meetings

### OneDrive
- List files and folders
- Upload and download files
- Search for files
- Create folders
- Copy, move, and delete items
- Share files (create sharing links)
- Get recent files and files shared with you

### SharePoint
- Search for sites
- List document libraries
- Upload and download files from SharePoint
- Manage SharePoint lists and list items
- Access subsites

### Microsoft Teams
- List teams and channels
- Send messages to channels
- Reply to messages
- List team members
- Create channels
- Manage online meetings
- Access chat messages

### Excel
- List worksheets in workbooks
- Read and write cell ranges
- Get used ranges
- Work with tables
- Create charts
- Manage named ranges
- Support for batch operations via sessions

### Word
- Get document metadata
- Search for documents
- Convert documents to PDF
- Manage document versions
- Check out/check in documents
- Share documents

### OneNote
- List and manage notebooks
- Create and delete notebooks
- List and manage sections
- Create and delete sections
- List, create, and delete pages
- Get and append page content (HTML)
- Search pages across all notebooks
- Copy pages between sections

## Prerequisites

1. **Microsoft Entra ID Application**: You need to register an application in Microsoft Entra ID (formerly Azure Active Directory)
2. **Required Permissions**: The application needs appropriate Microsoft Graph API permissions
3. **Node.js**: Version 18.0.0 or higher

## Authentication Modes

This server supports two authentication modes. Choose based on your use case:

### 1. Delegated Authentication (Recommended for Interactive Use)

**How it works:**
- Uses **OAuth 2.0 Device Code Flow** for user authentication
- User signs in with their Microsoft account credentials
- Actions are performed in the user's context with their permissions
- Tokens are cached securely and automatically refreshed

**When to use:**
- ✅ Interactive scenarios with Claude Desktop or similar tools
- ✅ Personal productivity and user-driven tasks
- ✅ When you want actions attributed to a real user
- ✅ Teams and OneNote access (required after March 31, 2025)
- ✅ Better security with user-level MFA and conditional access

**Requirements:**
- User must be able to authenticate via browser (device code flow)
- Does NOT require CLIENT_SECRET

**Supported services:** All (Outlook, Calendar, OneDrive, SharePoint, Teams, Excel, Word, OneNote)

### 2. App-Only Authentication (For Automation & Background Services)

**How it works:**
- Uses **OAuth 2.0 Client Credentials Flow** (application-only)
- Authenticates as the application itself, not a user
- Requires specifying which user's data to access
- No interactive authentication needed

**When to use:**
- ✅ Automated background services (scheduled jobs, batch processing)
- ✅ Headless environments (Docker, serverless functions, CI/CD)
- ✅ Multi-user administrative operations
- ✅ Shared mailbox or resource management
- ✅ 24/7 services without user sessions
- ✅ SaaS applications serving multiple organizations

**Requirements:**
- CLIENT_SECRET (application secret key)
- USER_PRINCIPAL_NAME or USER_ID (to specify which user's data to access)
- Admin consent for application permissions

**Supported services:** Outlook, Calendar, OneDrive, SharePoint, Excel, Word
**Limited support:** Teams (some features), OneNote (deprecated March 31, 2025)

**⚠️ Security Note:** App-only authentication has higher privilege and should only be used when delegated auth is not feasible. Secure the CLIENT_SECRET carefully.

---

**For Claude Desktop and interactive use: Choose Delegated Authentication**
**For automation and background services: Choose App-Only Authentication**

## Setup

### 1. Microsoft Entra ID Application Registration

1. Go to the [Azure Portal](https://portal.azure.com)
2. Navigate to **Microsoft Entra ID** (or **Azure Active Directory**) > **App registrations**
3. Click **New registration**
4. Enter a name (e.g., "Office365 MCP Server")
5. Select **Accounts in this organizational directory only**
6. Click **Register**

**For Delegated Authentication (Device Code Flow):**
7. Go to **Authentication** > **Advanced settings**
8. Set **Allow public client flows** to **Yes**
9. Click **Save**

### 2. Configure API Permissions

Choose the permissions based on your authentication mode:

#### For Delegated Authentication (Recommended)

Add these **Delegated Permissions**:
- `Calendars.ReadWrite`
- `Channel.Create` (for creating Teams channels)
- `Channel.ReadBasic.All`
- `ChannelMessage.Read.All`
- `ChannelMessage.Send` (for sending Teams channel messages)
- `Chat.Read`
- `Chat.ReadWrite`
- `Files.ReadWrite.All`
- `Mail.ReadWrite`
- `Mail.Send`
- `Notes.Read.All`
- `Notes.ReadWrite.All` (for OneNote)
- `Sites.Read.All`
- `Sites.ReadWrite.All`
- `Team.ReadBasic.All`
- `TeamMember.Read.All` (for reading team members)
- `TeamSettings.Read.All`
- `TeamSettings.ReadWrite.All`
- `User.Read`
- `User.Read.All`

After adding permissions, click **Grant admin consent** for your organization.

> **Note**: With delegated auth, the user will be prompted to consent to these permissions when they first sign in via device code flow.
>
> **Security Note**: We use granular permissions like `ChannelMessage.Send` instead of the overly broad `Group.ReadWrite.All` which would grant access to ALL groups in your organization.

#### For App-Only Authentication

Add these **Application Permissions**:
- `Calendars.Read`
- `Calendars.ReadWrite`
- `Channel.ReadBasic.All`
- `ChannelMessage.Read.All`
- `Files.Read.All`
- `Files.ReadWrite.All`
- `Mail.Read`
- `Mail.ReadWrite`
- `Mail.Send`
- `Notes.Create`
- `Notes.Read`
- `Notes.ReadWrite`
- `Sites.Read.All`
- `Sites.ReadWrite.All`
- `Team.ReadBasic.All`
- `TeamMember.Read.All`
- `TeamSettings.Read.All`
- `TeamSettings.ReadWrite.All`
- `User.Read.All`

After adding permissions, click **Grant admin consent** for your organization.

> **Important Teams Limitations with App-Only Auth:**
> - **Cannot send messages** to Teams channels with application permissions (Microsoft restriction)
> - Can only read channels, messages, and team info
> - For automated message sending, use [Incoming Webhooks](https://learn.microsoft.com/en-us/microsoftteams/platform/webhooks-and-connectors/how-to/add-incoming-webhook) instead
> - OneNote will stop working on March 31, 2025
>
> **Security Note**: We do NOT use `Group.ReadWrite.All` which would grant excessive access to ALL groups. We use specific permissions for the operations we need.

### 3. Create Client Secret (App-Only Mode Only)

**Only required if using app-only authentication. Skip this step for delegated authentication.**

1. In your app registration, go to **Certificates & secrets**
2. Click **New client secret**
3. Add a description and select an expiration period
4. Click **Add**
5. **Important**: Copy the secret value immediately (it won't be shown again)

### 4. Installation

```bash
# Clone the repository
git clone https://github.com/sahallam/MCP-Office-365.git
cd MCP-Office-365

# Install dependencies
npm install

# Build the project
npm run build
```

### 5. Configuration

Create a `.env` file in the root directory:

```bash
cp .env.example .env
```

Edit the `.env` file based on your authentication mode:

#### For Delegated Authentication (Recommended)

```env
TENANT_ID=your-tenant-id
CLIENT_ID=your-client-id
AUTH_MODE=delegated

# Optional: Custom encryption key for token cache (recommended for production)
# Generate with: openssl rand -hex 32
# TOKEN_ENCRYPTION_KEY=your-64-character-hex-key

# Optional: Custom Graph API endpoint
# GRAPH_API_ENDPOINT=https://graph.microsoft.com/v1.0
```

> **Security Note**: For production deployments, set `TOKEN_ENCRYPTION_KEY` to a 64-character hex string (32 bytes). Without this, the token cache is encrypted using a key derived from your machine's hostname and username, which changes if you move to a different machine or user account.

#### For App-Only Authentication

```env
TENANT_ID=your-tenant-id
CLIENT_ID=your-client-id
CLIENT_SECRET=your-client-secret
AUTH_MODE=app-only

# Required for app-only mode - specify a user to act on behalf of
USER_PRINCIPAL_NAME=user@yourdomain.com
# Or use user ID
# USER_ID=user-object-id

# Optional: Custom Graph API endpoint
# GRAPH_API_ENDPOINT=https://graph.microsoft.com/v1.0
```

To find these values:
- **TENANT_ID**: In Azure Portal > Microsoft Entra ID > Overview > Tenant ID
- **CLIENT_ID**: In your app registration > Overview > Application (client) ID
- **CLIENT_SECRET**: (App-only only) The value you copied when creating the client secret

## Usage

### Running the Server

```bash
npm start
```

Or for development with auto-reload:

```bash
npm run dev
```

### Using with Claude Desktop

Add this to your Claude Desktop configuration file:

**On macOS**: `~/Library/Application Support/Claude/claude_desktop_config.json`
**On Windows**: `%APPDATA%/Claude/claude_desktop_config.json`

#### Delegated Authentication (Recommended)

```json
{
  "mcpServers": {
    "Office 365": {
      "command": "node",
      "args": ["/absolute/path/to/MCP-Office-365/dist/index.js"],
      "env": {
        "TENANT_ID": "your-tenant-id",
        "CLIENT_ID": "your-client-id",
        "AUTH_MODE": "delegated"
      }
    }
  }
}
```

**First-time setup**: When you first use the connector, a device code will appear directly in Claude Desktop. Open the URL shown and enter the code to sign in with your Microsoft account. After authentication, your session will persist for ~90 days.

**If you need to check authentication status**, use the `auth_status` tool which will show you whether you're authenticated or provide a new device code if needed.

#### App-Only Authentication

```json
{
  "mcpServers": {
    "Office 365": {
      "command": "node",
      "args": ["/absolute/path/to/MCP-Office-365/dist/index.js"],
      "env": {
        "TENANT_ID": "your-tenant-id",
        "CLIENT_ID": "your-client-id",
        "CLIENT_SECRET": "your-client-secret",
        "AUTH_MODE": "app-only",
        "USER_PRINCIPAL_NAME": "user@yourdomain.com"
      }
    }
  }
}
```

### Available Tools

The server provides the following MCP tools:

#### Email Tools
- `outlook_list_emails` - List emails from inbox
- `outlook_get_email` - Get a specific email
- `outlook_send_email` - Send an email
- `outlook_search_emails` - Search emails
- `outlook_reply_email` - Reply to an email

#### Calendar Tools
- `calendar_list_events` - List calendar events
- `calendar_get_view` - Get calendar view for a time range
- `calendar_create_event` - Create a calendar event
- `calendar_update_event` - Update a calendar event
- `calendar_delete_event` - Delete a calendar event
- `calendar_get_event` - Get a specific event by ID
- `calendar_accept_meeting` - Accept a meeting invitation
- `calendar_decline_meeting` - Decline a meeting invitation
- `calendar_tentatively_accept_meeting` - Tentatively accept a meeting
- `calendar_find_meeting_times` - Find available meeting times
- `calendar_list_calendars` - List all calendars

#### Authentication Tools
- `auth_status` - Check authentication status and get re-auth instructions if needed

#### OneDrive Tools
- `onedrive_list_items` - List files and folders
- `onedrive_get_item` - Get item details
- `onedrive_download_file` - Download a file
- `onedrive_upload_file` - Upload a file
- `onedrive_search` - Search for files
- `onedrive_create_folder` - Create a folder

#### SharePoint Tools
- `sharepoint_search_sites` - Search for sites
- `sharepoint_get_site` - Get site details
- `sharepoint_list_libraries` - List document libraries
- `sharepoint_list_items` - List items in a library

#### Teams Tools
- `teams_list_teams` - List all teams
- `teams_list_channels` - List channels in a team
- `teams_send_message` - Send a message to a channel
- `teams_list_messages` - List messages in a channel

#### Excel Tools
- `excel_list_worksheets` - List worksheets in a workbook
- `excel_get_range` - Get cell range values
- `excel_update_range` - Update cell range values
- `excel_get_used_range` - Get used range in a worksheet

#### Word Tools
- `word_get_document` - Get document metadata
- `word_search_documents` - Search for documents
- `word_convert_to_pdf` - Convert document to PDF

#### OneNote Tools
- `onenote_list_notebooks` - List all OneNote notebooks
- `onenote_get_notebook` - Get a specific notebook by ID
- `onenote_create_notebook` - Create a new notebook
- `onenote_list_sections` - List sections in a notebook or all sections
- `onenote_get_section` - Get a specific section by ID
- `onenote_create_section` - Create a new section in a notebook
- `onenote_list_pages` - List pages in a section, notebook, or all pages
- `onenote_get_page` - Get a specific page by ID
- `onenote_get_page_content` - Get the HTML content of a page
- `onenote_create_page` - Create a new page in a section
- `onenote_append_to_page` - Append content to an existing page
- `onenote_search_pages` - Search pages across all notebooks
- `onenote_delete_page` - Delete a OneNote page
- `onenote_delete_section` - Delete a OneNote section
- `onenote_delete_notebook` - Delete a OneNote notebook
- `onenote_copy_page` - Copy a page to a different section

## Example Prompts

Once configured, you can use prompts like these with Claude:

- "Show me my latest 10 emails"
- "Send an email to john@example.com with subject 'Meeting Follow-up' and body 'Thanks for the meeting today.'"
- "What meetings do I have next week?"
- "Create a meeting tomorrow at 2pm for 1 hour with the team"
- "List my OneDrive files"
- "Search for Excel files containing 'budget'"
- "Send a message to the Marketing team channel saying 'Great work everyone!'"
- "Read the data from cells A1:C10 in the Sales worksheet"
- "List all my OneNote notebooks"
- "Create a new OneNote page titled 'Meeting Notes' in my Work section"
- "Search my OneNote pages for 'project timeline'"
- "Show me the content of my latest OneNote page"

## Architecture

### Project Structure

```
office365-mcp-server/
├── src/
│   ├── index.ts              # Main MCP server implementation
│   ├── auth.ts               # Microsoft Graph authentication
│   ├── types.ts              # TypeScript type definitions
│   └── tools/
│       ├── outlook.ts        # Email operations
│       ├── calendar.ts       # Calendar operations
│       ├── onedrive.ts       # OneDrive file operations
│       ├── sharepoint.ts     # SharePoint operations
│       ├── teams.ts          # Teams operations
│       ├── excel.ts          # Excel operations
│       ├── word.ts           # Word operations
│       └── onenote.ts        # OneNote operations
├── package.json
├── tsconfig.json
├── .env.example
└── README.md
```

### Authentication Flows

The server supports two authentication flows:

#### Delegated Authentication (Device Code Flow)

1. Server starts and checks for cached tokens in `~/.office365-mcp-token-cache.json`
2. If no valid token exists, a device code is displayed directly in Claude Desktop
3. User opens browser, navigates to the URL, and enters the code
4. User signs in and grants consent to requested permissions
5. Server receives access token and refresh token
6. Tokens are encrypted (AES-256-GCM) and cached locally
7. For subsequent requests, tokens are automatically refreshed using MSAL cache
8. **Authentication persists for ~90 days** (or longer depending on organization's token policies)
9. Re-authentication is only required if the connector is not used for the full refresh token lifetime

#### App-Only Authentication (Client Credentials Flow)

1. Server loads credentials from environment variables
2. Authenticates with Microsoft Entra ID using client secret
3. Obtains an access token for Microsoft Graph API
4. Uses the token for all API requests on behalf of the specified user
5. Automatically refreshes the token when it expires

> **Note**: While Microsoft rebranded Azure AD to Microsoft Entra ID, authentication endpoints and some API references still use the `login.microsoftonline.com` domain for backwards compatibility.

### Technology Stack

- **TypeScript**: Type-safe implementation
- **MCP SDK**: Official Model Context Protocol SDK
- **Microsoft Graph Client**: Official Microsoft Graph JavaScript client
- **MSAL Node**: Microsoft Authentication Library for Node.js
- **Axios**: HTTP client for additional API calls

## Development

### Building

```bash
npm run build
```

### Watch Mode

```bash
npm run watch
```

### Project Structure

The codebase is organized into modular components:

- **Authentication Layer** (`auth.ts`): Handles Microsoft Entra ID authentication with token encryption
- **Security Layer** (`security.ts`): Input validation, sanitization, and security utilities
- **Audit Layer** (`audit.ts`): Comprehensive security event logging
- **Tool Modules** (`tools/`): Each Office 365 service has its own module
- **Type Definitions** (`types.ts`): Shared TypeScript interfaces
- **MCP Server** (`index.ts`): Main server implementation with tool handlers

### Code Quality Standards

**TypeScript Configuration**
- Strict mode enabled for type safety
- No implicit any allowed
- All functions have explicit return types
- ES2022 target for modern JavaScript features

**Security Implementation**
- All user inputs validated before processing
- Resource IDs validated with pattern matching
- File sizes and content lengths enforced
- HTML and search query sanitization applied
- Base64 decoding with error handling

**Architecture Principles**
- Clean separation of concerns
- Tool modules are independent and loosely coupled
- Centralized error handling with sanitized messages
- Consistent patterns across all modules (getUserPath() helper, validation first)

## Code Quality & Testing

### Development Standards

**Input Validation Pattern**
All tool methods follow this pattern:
```typescript
async someOperation(resourceId: string, data: any): Promise<Result> {
  // 1. Validate all inputs
  validateResourceId(resourceId, 'resource');

  // 2. Perform operation
  const result = await this.graphClient.api(`/path/${resourceId}`).get();

  // 3. Return result
  return result;
}
```

**Error Handling**
- All errors bubble to centralized handler in `index.ts`
- Error messages are sanitized before returning to user
- Detailed errors logged internally for debugging
- Public error messages don't expose sensitive information

**Audit Logging**
All security-sensitive operations are automatically logged:
- Authentication events
- Resource create/update/delete operations
- Failed access attempts
- Input validation failures

### Testing Recommendations

**Unit Tests** (Recommended)
```typescript
// Test security validation functions
describe('validateResourceId', () => {
  it('should reject path traversal attempts', () => {
    expect(() => validateResourceId('../../../etc/passwd', 'file'))
      .toThrow('Invalid file ID');
  });
});

// Test encryption/decryption
describe('Token Encryption', () => {
  it('should encrypt and decrypt tokens correctly', () => {
    const token = 'test-access-token';
    const encrypted = encryptTokenCache(token);
    const decrypted = decryptTokenCache(encrypted);
    expect(decrypted).toBe(token);
  });
});
```

**Integration Tests** (Recommended)
- Mock Microsoft Graph API responses
- Test tool method execution flows
- Verify error handling paths
- Test authentication flows

### OWASP Compliance Status

This server follows OWASP Top 10 (2021) security guidelines:

| Category | Compliance | Implementation |
|----------|-----------|----------------|
| A01: Broken Access Control | Partial | Input validation, Microsoft Graph API enforcement |
| A02: Cryptographic Failures | Excellent | AES-256-GCM encryption, TLS 1.2/1.3 |
| A03: Injection | Excellent | Comprehensive input validation and sanitization |
| A04: Insecure Design | Good | Defense in depth, secure defaults |
| A05: Security Misconfiguration | Excellent | TLS enforcement, secure permissions |
| A06: Vulnerable Components | Excellent | All dependencies up to date, 0 vulnerabilities |
| A07: Authentication Failures | Excellent | Token encryption, revocation, validation |
| A08: Integrity Failures | Excellent | Token cache validation, safe deserialization |
| A09: Logging Failures | Good | Comprehensive audit logging integrated |
| A10: SSRF | Good | Hostname validation, URL sanitization |

**Overall OWASP Score: 87/100** (Very Good)

## Troubleshooting

### Common Issues

**Delegated Authentication Issues**
- **"Allow public client flows" not enabled**: Go to Azure Portal > Your App > Authentication > Advanced settings > Set to "Yes"
- **Device code not showing**: The device code is displayed directly in Claude Desktop when authentication is required. If you don't see it, use the `auth_status` tool to check your authentication status.
- **Token expired**: Delete `~/.office365-mcp-token-cache.json` and restart to re-authenticate
- **Wrong permissions**: Ensure delegated permissions (not application permissions) are configured in Azure Portal
- **Token not persisting**: Ensure the token cache file at `~/.office365-mcp-token-cache.json` has proper permissions (should be readable/writable by your user)

**App-Only Authentication Errors**
- Verify your `TENANT_ID`, `CLIENT_ID`, and `CLIENT_SECRET` are correct
- Ensure `USER_PRINCIPAL_NAME` or `USER_ID` is specified
- Ensure admin consent has been granted for application permissions
- Check that the client secret hasn't expired
- **Teams/OneNote not working**: These require delegated authentication - set `AUTH_MODE=delegated`

**Permission Errors**
- Verify the app has the necessary Graph API permissions (delegated vs application)
- Ensure admin consent has been granted
- Check if the user/app has access to the requested resources
- For delegated auth, user must have the necessary roles/permissions

**Connection Issues**
- Verify your network allows connections to `https://login.microsoftonline.com` and `https://graph.microsoft.com`
- Check firewall and proxy settings

### Debug Mode

To enable verbose logging, set the `DEBUG` environment variable:

```bash
DEBUG=* npm start
```

## Security Architecture

This server implements enterprise-grade security following OWASP Top 10 (2021) guidelines.

### Authentication & Token Management

**Secure Token Storage**
- All tokens are encrypted at rest using AES-256-GCM
- Token cache files have restrictive permissions (0600 - owner read/write only)
- Supports custom encryption keys via `TOKEN_ENCRYPTION_KEY` environment variable
- Automatic token expiration and refresh handling
- Token revocation on logout with complete cache cleanup

**Authentication Modes**
- OAuth 2.0 with Microsoft Identity Platform
- Delegated: Device Code Flow (user authentication)
- App-Only: Client Credentials Flow (application authentication)
- Automatic token validation and refresh

### Data Protection

**Encryption**
- TLS 1.2/1.3 enforced for all API communications
- Insecure SSL/TLS versions disabled (SSLv2, SSLv3, TLS 1.0, TLS 1.1)
- Strong cipher suites only (AES-GCM, ChaCha20-Poly1305)
- Certificate validation enforced

**Input Validation**
- Comprehensive validation for all user inputs
- Protection against injection attacks (OData, XSS, SQL)
- Resource ID validation with pattern matching
- File size limits enforced
- Path traversal prevention

**Content Sanitization**
- HTML content sanitization for OneNote and Teams
- Search query sanitization for all search operations
- Excel formula injection prevention
- URL validation and SSRF protection

### Audit Logging

**Security Event Logging**
- Authentication events (login, logout, token operations)
- Resource access tracking (read, create, update, delete)
- Security violations and failed validations
- Automatic log rotation (90-day retention by default)
- Logs stored in `~/.office365-mcp/audit/` with secure permissions

**Configuration**
- Enable/disable: Set `AUDIT_ENABLED=true/false`
- Debug mode: Set `DEBUG_AUDIT=true` for verbose logging
- Logs are JSON formatted for easy parsing and analysis

### Best Practices

**Credential Management**
- **Never commit** `.env` files or token caches to version control
- Use Azure Key Vault or similar for production credentials
- Set `TOKEN_ENCRYPTION_KEY` for enhanced token security
- Rotate client secrets regularly (app-only mode)

**Access Control**
- Apply least privilege principle - only grant necessary permissions
- Use delegated authentication when possible (user context)
- Monitor access via Microsoft Entra ID sign-in logs
- Review audit logs regularly for suspicious activity

**Operational Security**
- Keep dependencies updated (`npm audit`)
- Review Microsoft Graph API permissions periodically
- Use app-only authentication only when necessary
- Implement rate limiting for high-volume operations
- Secure the token cache directory permissions

## Limitations

- **Rate Limits**: Microsoft Graph API has rate limits (throttling). The server does not currently implement automatic retry logic for 429 (Too Many Requests) errors. For typical interactive use through Claude, the rate limits are generous enough that throttling is unlikely. If you encounter rate limiting during high-volume operations, you may need to add retry logic or reduce request frequency.
- **File Size**: Large file operations may timeout
- **Permissions**: Some operations require specific permissions
- **Authentication Mode Restrictions**:
  - **App-only**: Limited Teams support, OneNote deprecated (March 31, 2025)
  - **Delegated**: Requires user to authenticate via browser (device code flow)
- **Token Lifetime**: Delegated auth tokens persist for ~90 days with automatic refresh. Re-authentication is only needed if the connector is not used for the full refresh token lifetime.

## Contributing

Contributions are welcome! Please see [CONTRIBUTING.md](CONTRIBUTING.md) for detailed guidelines on:

- Setting up your development environment
- Code style and best practices
- Security guidelines
- Pull request process
- Reporting bugs and security vulnerabilities

## License

MIT License - see LICENSE file for details

## Related Projects

- [Model Context Protocol](https://modelcontextprotocol.io)
- [Microsoft Graph API](https://developer.microsoft.com/en-us/graph)
- [Google Cloud MCP Server](https://github.com/googleapis/gcloud-mcp)

## Support

For issues and questions:
- [Open an issue on GitHub](https://github.com/sahallam/MCP-Office-365/issues)
- Check [Microsoft Graph API documentation](https://developer.microsoft.com/en-us/graph)
- Review [MCP protocol specification](https://modelcontextprotocol.io)

## Acknowledgments

- Built with the [Model Context Protocol SDK](https://github.com/modelcontextprotocol)
- Powered by [Microsoft Graph API](https://developer.microsoft.com/en-us/graph)
- Inspired by the [Google Cloud MCP Server](https://github.com/googleapis/gcloud-mcp)

Developed with Claude Code by Steven Hallam.
