# Office365 MCP Server

A comprehensive [Model Context Protocol](https://modelcontextprotocol.io) (MCP) server for Microsoft Office 365 integration. This server enables AI assistants like Claude to interact with your Office 365 environment, including Outlook, Calendar, OneDrive, SharePoint, Teams, Excel, Word, and OneNote. 

This code was developed by Claude Code.

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

This server supports two authentication modes:

### 1. Delegated Authentication (Recommended)
- Uses **device code flow** for user authentication
- User signs in with their Microsoft account
- **Required for**: Teams, OneNote (after March 31, 2025)
- **Best for**: Interactive scenarios where a user can authenticate
- **Does NOT require**: CLIENT_SECRET
- Tokens are cached and automatically refreshed

### 2. App-Only Authentication
- Uses **client credentials flow** (application-only)
- No user interaction required
- **Required for**: Automation scenarios
- **Requires**: CLIENT_SECRET, USER_PRINCIPAL_NAME or USER_ID
- Works for Outlook, Calendar, OneDrive, SharePoint
- **Limited support**: Teams (some features), OneNote (deprecated March 31, 2025)

**For most use cases with Claude Desktop, use Delegated Authentication.**

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
- `Channel.ReadBasic.All`
- `ChannelMessage.Read.All`
- `Chat.Read`
- `Chat.ReadWrite`
- `Files.ReadWrite.All`
- `Group.ReadWrite.All` (required for sending Teams channel messages)
- `Mail.ReadWrite`
- `Mail.Send`
- `Notes.Read.All`
- `Notes.ReadWrite.All` (for OneNote)
- `Sites.Read.All`
- `Sites.ReadWrite.All`
- `Team.ReadBasic.All`
- `TeamSettings.Read.All`
- `TeamSettings.ReadWrite.All`
- `User.Read`
- `User.Read.All`

After adding permissions, click **Grant admin consent** for your organization.

> **Note**: With delegated auth, the user will be prompted to consent to these permissions when they first sign in via device code flow.

#### For App-Only Authentication

Add these **Application Permissions**:
- `Calendars.Read`
- `Calendars.ReadWrite`
- `Channel.ReadBasic.All`
- `ChannelMessage.Read.All`
- `ChannelMessage.Send` (for sending Teams messages as the user)
- `Files.Read.All`
- `Files.ReadWrite.All`
- `Group.ReadWrite.All` (required for Teams operations)
- `Mail.Read`
- `Mail.ReadWrite`
- `Mail.Send`
- `Notes.Create`
- `Notes.Read`
- `Notes.ReadWrite`
- `Sites.Read.All`
- `Sites.ReadWrite.All`
- `Team.ReadBasic.All`
- `TeamSettings.Read.All`
- `TeamSettings.ReadWrite.All`
- `User.Read.All`

After adding permissions, click **Grant admin consent** for your organization.

> **Important**: App-only mode has limited Teams support and OneNote will stop working on March 31, 2025.

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
git clone https://github.com/sahallam/office365-mcp-server.git
cd office365-mcp-server

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

# Optional: Custom Graph API endpoint
# GRAPH_API_ENDPOINT=https://graph.microsoft.com/v1.0
```

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

**First-time setup**: When you start Claude Desktop, check the logs for device code authentication instructions. You'll see a URL and code to enter in your browser to sign in.

**Logs location**:
- macOS: `~/Library/Logs/Claude/mcp*.log`
- Windows: `%APPDATA%\Claude\logs\mcp*.log`

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
- `calendar_delete_event` - Delete a calendar event

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

1. Server starts and checks for cached access token
2. If no valid token exists, prompts user with a device code
3. User opens browser, navigates to the URL, and enters the code
4. User signs in and grants consent to requested permissions
5. Server receives access token and refresh token
6. Tokens are cached locally in `.token-cache.json`
7. For subsequent requests, tokens are automatically refreshed using MSAL cache
8. User only needs to authenticate once (until token expires)

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

- **Authentication Layer** (`auth.ts`): Handles Microsoft Entra ID authentication
- **Tool Modules** (`tools/`): Each Office 365 service has its own module
- **Type Definitions** (`types.ts`): Shared TypeScript interfaces
- **MCP Server** (`index.ts`): Main server implementation with tool handlers

## Troubleshooting

### Common Issues

**Delegated Authentication Issues**
- **"Allow public client flows" not enabled**: Go to Azure Portal > Your App > Authentication > Advanced settings > Set to "Yes"
- **Can't see device code**: Check Claude Desktop logs at `~/Library/Logs/Claude/mcp*.log` (macOS) or `%APPDATA%\Claude\logs\` (Windows)
- **Token expired**: Delete `.token-cache.json` and restart to re-authenticate
- **Wrong permissions**: Ensure delegated permissions (not application permissions) are configured in Azure Portal

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

## Security Considerations

- **Never commit** your `.env` file or `.token-cache.json` to version control
- **Token cache**: The `.token-cache.json` file contains access tokens - keep it secure
- **Rotate client secrets** regularly (app-only mode)
- **Use least privilege**: Only grant necessary permissions
- **Monitor access**: Review Microsoft Entra ID sign-in logs regularly
- **Secure storage**: Store credentials securely (use Azure Key Vault in production)
- **Delegated auth**: Tokens are tied to the signed-in user - ensure the user has appropriate access

## Limitations

- **Rate Limits**: Microsoft Graph API has rate limits (throttling). The server does not currently implement automatic retry logic for 429 (Too Many Requests) errors. For typical interactive use through Claude, the rate limits are generous enough that throttling is unlikely. If you encounter rate limiting during high-volume operations, you may need to add retry logic or reduce request frequency.
- **File Size**: Large file operations may timeout
- **Permissions**: Some operations require specific permissions
- **Authentication Mode Restrictions**:
  - **App-only**: Limited Teams support, OneNote deprecated (March 31, 2025)
  - **Delegated**: Requires user to authenticate via browser (device code flow)
- **Token Lifetime**: Delegated auth tokens expire and require re-authentication periodically

## Contributing

Contributions are welcome! Please:

1. Fork the repository
2. Create a feature branch
3. Make your changes
4. Add tests if applicable
5. Submit a pull request

## License

MIT License - see LICENSE file for details

## Related Projects

- [Model Context Protocol](https://modelcontextprotocol.io)
- [Microsoft Graph API](https://developer.microsoft.com/en-us/graph)
- [Google Cloud MCP Server](https://github.com/googleapis/gcloud-mcp)

## Support

For issues and questions:
- Open an issue on GitHub
- Check Microsoft Graph API documentation
- Review MCP protocol specification

## Acknowledgments

- Developed by Claude Code, prompted by Steven Hallam.
- Built with the [Model Context Protocol SDK](https://github.com/modelcontextprotocol)
- Uses [Microsoft Graph API](https://developer.microsoft.com/en-us/graph)
- Inspired by the [Google Cloud MCP Server](https://github.com/googleapis/gcloud-mcp)
