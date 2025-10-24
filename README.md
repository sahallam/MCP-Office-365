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

## Setup

### 1. Microsoft Entra ID Application Registration

1. Go to the [Azure Portal](https://portal.azure.com)
2. Navigate to **Microsoft Entra ID** (or **Azure Active Directory**) > **App registrations**
3. Click **New registration**
4. Enter a name (e.g., "Office365 MCP Server")
5. Select **Accounts in this organizational directory only**
6. Click **Register**

### 2. Configure API Permissions

Add the following Microsoft Graph API permissions:

**Application Permissions** (for app-only access):
- `Mail.Read`
- `Mail.ReadWrite`
- `Mail.Send`
- `Calendars.Read`
- `Calendars.ReadWrite`
- `Files.Read.All`
- `Files.ReadWrite.All`
- `Sites.Read.All`
- `Sites.ReadWrite.All`
- `TeamSettings.Read.All`
- `TeamSettings.ReadWrite.All`
- `ChannelMessage.Read.All`
- `Group.ReadWrite.All` (required for sending Teams channel messages)
- `Notes.Read.All`
- `Notes.ReadWrite.All`
- `User.Read.All`

**Delegated Permissions** (for user context):
- `Mail.Read`
- `Mail.ReadWrite`
- `Mail.Send`
- `Calendars.Read`
- `Calendars.ReadWrite`
- `Files.Read`
- `Files.ReadWrite`
- `Sites.Read.All`
- `Sites.ReadWrite.All`
- `ChannelMessage.Send` (for sending Teams messages as the user)
- `Team.ReadBasic.All`
- `Channel.ReadBasic.All`
- `Notes.Read`
- `Notes.ReadWrite`
- `Notes.Create`

After adding permissions, click **Grant admin consent** for your organization.

> **Important Note on Teams Permissions**:
> - `Group.ReadWrite.All` is a broad permission that allows reading and writing to all Microsoft 365 groups and teams
> - For production use, consider using Resource-Specific Consent (RSC) or Teams Bot Framework for more granular control
> - Delegated permissions with `ChannelMessage.Send` provide better security when running on behalf of a specific user

### 3. Create Client Secret

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

Edit the `.env` file with your Microsoft Entra ID application details:

```env
TENANT_ID=your-tenant-id
CLIENT_ID=your-client-id
CLIENT_SECRET=your-client-secret

# Optional: Specify a user to act on behalf of
USER_PRINCIPAL_NAME=user@yourdomain.com
# Or use user ID
# USER_ID=user-object-id

# Optional: Custom Graph API endpoint
# GRAPH_API_ENDPOINT=https://graph.microsoft.com/v1.0
```

To find these values:
- **TENANT_ID**: In Azure Portal > Microsoft Entra ID > Overview > Tenant ID
- **CLIENT_ID**: In your app registration > Overview > Application (client) ID
- **CLIENT_SECRET**: The value you copied when creating the client secret

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

```json
{
  "mcpServers": {
    "Office 365": {
      "command": "node",
      "args": ["/absolute/path/to/office365-mcp-server/dist/index.js"],
      "env": {
        "TENANT_ID": "your-tenant-id",
        "CLIENT_ID": "your-client-id",
        "CLIENT_SECRET": "your-client-secret",
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

### Authentication Flow

The server uses the **Client Credentials Flow** (app-only authentication):

1. Server loads credentials from environment variables
2. Authenticates with Microsoft Entra ID using MSAL (Microsoft Authentication Library)
3. Obtains an access token for Microsoft Graph API
4. Uses the token for all API requests
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

**Authentication Errors**
- Verify your `TENANT_ID`, `CLIENT_ID`, and `CLIENT_SECRET` are correct
- Ensure admin consent has been granted for the required permissions
- Check that the client secret hasn't expired

**Permission Errors**
- Verify the app has the necessary Graph API permissions
- Ensure admin consent has been granted
- Check if the user/app has access to the requested resources

**Connection Issues**
- Verify your network allows connections to `https://login.microsoftonline.com` and `https://graph.microsoft.com`
- Check firewall and proxy settings

### Debug Mode

To enable verbose logging, set the `DEBUG` environment variable:

```bash
DEBUG=* npm start
```

## Security Considerations

- **Never commit** your `.env` file or expose credentials
- **Rotate client secrets** regularly
- **Use least privilege**: Only grant necessary permissions
- **Monitor access**: Review Microsoft Entra ID sign-in logs regularly
- **Secure storage**: Store credentials securely (use Azure Key Vault in production)

## Limitations

- **Rate Limits**: Microsoft Graph API has rate limits (throttling)
- **File Size**: Large file operations may timeout
- **Permissions**: Some operations require specific permissions
- **Delegated vs Application**: Some features work differently with app-only vs delegated permissions

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
