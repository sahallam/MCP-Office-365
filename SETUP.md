# Office365 MCP Server - Detailed Setup Guide

This guide walks you through the complete setup process for the Office365 MCP Server.

## Table of Contents

1. [Prerequisites](#prerequisites)
2. [Microsoft Entra ID Application Setup](#microsoft-entra-id-application-setup)
3. [Installing the Server](#installing-the-server)
4. [Configuration](#configuration)
5. [Testing the Setup](#testing-the-setup)
6. [Integration with Claude Desktop](#integration-with-claude-desktop)

## Prerequisites

Before you begin, ensure you have:

- **Node.js 18+**: [Download](https://nodejs.org/)
- **npm** or **yarn**: Comes with Node.js
- **Azure Account**: With permissions to create Microsoft Entra ID applications
- **Microsoft 365**: Access to a Microsoft 365 tenant

## Microsoft Entra ID Application Setup

> **Note**: Microsoft Entra ID is the new name for Azure Active Directory. The Azure Portal may show either name depending on when it was last updated.

### Step 1: Create a Microsoft Entra ID Application

1. Navigate to the [Azure Portal](https://portal.azure.com)
2. Go to **Microsoft Entra ID** (you may see **Azure Active Directory** in older portal views)
3. Select **App registrations** from the left menu
4. Click **+ New registration**

   ![New Registration](docs/images/new-registration.png)

5. Fill in the details:
   - **Name**: `Office365 MCP Server` (or your preferred name)
   - **Supported account types**: Select "Accounts in this organizational directory only"
   - **Redirect URI**: Leave blank (not needed for this application)

6. Click **Register**

### Step 2: Copy Application Details

After registration, you'll see the application overview page:

1. Copy the **Application (client) ID** - you'll use this as `CLIENT_ID`
2. Copy the **Directory (tenant) ID** - you'll use this as `TENANT_ID`

### Step 3: Create a Client Secret

1. In your app registration, click **Certificates & secrets** in the left menu
2. Click **+ New client secret**
3. Add a description (e.g., "MCP Server Secret")
4. Choose an expiration period:
   - **Recommended**: 6 months or 1 year
   - **Note**: You'll need to rotate this secret before it expires
5. Click **Add**
6. **IMPORTANT**: Copy the secret **Value** immediately - it won't be shown again
   - This is your `CLIENT_SECRET`

### Step 4: Configure API Permissions

Now we'll grant the application access to Microsoft Graph API:

1. Click **API permissions** in the left menu
2. Click **+ Add a permission**
3. Select **Microsoft Graph**
4. Choose **Application permissions** (for app-only access)

Add these permissions:

**Mail Permissions:**
- Mail.Read
- Mail.ReadWrite
- Mail.Send

**Calendar Permissions:**
- Calendars.Read
- Calendars.ReadWrite

**Files Permissions:**
- Files.Read.All
- Files.ReadWrite.All

**Sites Permissions:**
- Sites.Read.All
- Sites.ReadWrite.All

**Teams Permissions:**
- TeamSettings.Read.All
- TeamSettings.ReadWrite.All
- ChannelMessage.Read.All
- Group.ReadWrite.All (required for sending channel messages)

**User Permissions:**
- User.Read.All

> **Note**: `Group.ReadWrite.All` is required for application permissions to send messages to Teams channels. For delegated permissions (user context), you can use `ChannelMessage.Send` instead.

### Step 5: Grant Admin Consent

**Critical Step**: After adding all permissions:

1. Click **Grant admin consent for [Your Organization]**
2. Click **Yes** to confirm
3. Verify all permissions show "Granted for [Your Organization]"

Without admin consent, the application won't be able to access Microsoft Graph API.

## Installing the Server

### Option 1: From Source

```bash
# Clone the repository
git clone https://github.com/yourusername/office365-mcp-server.git
cd office365-mcp-server

# Install dependencies
npm install

# Build the TypeScript code
npm run build
```

### Option 2: From npm (when published)

```bash
npm install -g @office365/mcp-server
```

## Configuration

### Create Environment File

1. Copy the example environment file:

```bash
cp .env.example .env
```

2. Edit `.env` with your values:

```env
# Required: Azure AD Application Details
TENANT_ID=12345678-1234-1234-1234-123456789abc
CLIENT_ID=87654321-4321-4321-4321-cba987654321
CLIENT_SECRET=your-client-secret-value-here

# Optional: Specify which user's mailbox/calendar to access
USER_PRINCIPAL_NAME=user@yourdomain.com

# Optional: Or use user object ID instead
# USER_ID=a1b2c3d4-e5f6-7890-abcd-ef1234567890

# Optional: Custom Graph API endpoint (rarely needed)
# GRAPH_API_ENDPOINT=https://graph.microsoft.com/v1.0
```

### Configuration Options

**Required Variables:**

- `TENANT_ID`: Your Azure AD tenant ID (from Step 2)
- `CLIENT_ID`: Your application (client) ID (from Step 2)
- `CLIENT_SECRET`: The client secret value (from Step 3)

**Optional Variables:**

- `USER_PRINCIPAL_NAME`: Email address of the user to act on behalf of
  - Example: `john.doe@contoso.com`
  - Use this when you want to access a specific user's mailbox

- `USER_ID`: Azure AD object ID of the user
  - Alternative to `USER_PRINCIPAL_NAME`
  - Find this in Azure AD > Users > [Select User] > Object ID

- `GRAPH_API_ENDPOINT`: Custom Graph API endpoint
  - Default: `https://graph.microsoft.com/v1.0`
  - Only change if using a different environment (e.g., government cloud)

## Testing the Setup

### Test 1: Verify the Build

```bash
npm run build
```

Should complete without errors.

### Test 2: Run the Server

```bash
npm start
```

You should see:
```
Office365 MCP Server running on stdio
```

### Test 3: Manual Authentication Test

Create a test script `test-auth.js`:

```javascript
import { ConfidentialClientApplication } from '@azure/msal-node';
import dotenv from 'dotenv';

dotenv.config();

const msalClient = new ConfidentialClientApplication({
  auth: {
    clientId: process.env.CLIENT_ID,
    authority: `https://login.microsoftonline.com/${process.env.TENANT_ID}`,
    clientSecret: process.env.CLIENT_SECRET,
  },
});

async function testAuth() {
  try {
    const result = await msalClient.acquireTokenByClientCredential({
      scopes: ['https://graph.microsoft.com/.default'],
    });

    console.log('✓ Authentication successful!');
    console.log('Access token received:', result.accessToken.substring(0, 20) + '...');
  } catch (error) {
    console.error('✗ Authentication failed:', error.message);
  }
}

testAuth();
```

Run it:
```bash
node test-auth.js
```

## Integration with Claude Desktop

### Locate Configuration File

**macOS:**
```
~/Library/Application Support/Claude/claude_desktop_config.json
```

**Windows:**
```
%APPDATA%\Claude\claude_desktop_config.json
```

**Linux:**
```
~/.config/Claude/claude_desktop_config.json
```

### Add Server Configuration

Edit the configuration file:

```json
{
  "mcpServers": {
    "office365": {
      "command": "node",
      "args": [
        "/absolute/path/to/office365-mcp-server/dist/index.js"
      ],
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

**Important Notes:**

1. Use absolute paths (not relative paths like `~/`)
2. On Windows, use double backslashes: `C:\\Users\\...\\dist\\index.js`
3. Replace the placeholder values with your actual credentials
4. Restart Claude Desktop after saving

### Verify Integration

1. Restart Claude Desktop
2. Open a new conversation
3. Type: "List my latest emails"
4. Claude should use the Office365 MCP server to fetch your emails

## Common Setup Issues

### Issue: "Authentication failed"

**Causes:**
- Incorrect client ID, tenant ID, or client secret
- Client secret expired
- Missing admin consent

**Solutions:**
1. Double-check all credentials in `.env`
2. Verify the client secret hasn't expired (Azure Portal > Microsoft Entra ID > App registrations > Certificates & secrets)
3. Re-grant admin consent (Azure Portal > Microsoft Entra ID > App registrations > API permissions)

### Issue: "Insufficient privileges"

**Causes:**
- Missing required permissions
- Admin consent not granted

**Solutions:**
1. Verify all required permissions are added (see Step 4)
2. Click "Grant admin consent" in the API permissions page

### Issue: "Server not found" in Claude Desktop

**Causes:**
- Incorrect path to `index.js`
- Node.js not in PATH

**Solutions:**
1. Use absolute path to `dist/index.js`
2. Try using absolute path to node: `/usr/local/bin/node` (macOS/Linux) or `C:\\Program Files\\nodejs\\node.exe` (Windows)

### Issue: "Cannot find module"

**Causes:**
- Dependencies not installed
- Build not completed

**Solutions:**
```bash
npm install
npm run build
```

## Security Best Practices

1. **Never commit `.env`** - It's already in `.gitignore`, but be careful
2. **Rotate secrets regularly** - Set a reminder to rotate client secrets
3. **Use Azure Key Vault** - For production deployments
4. **Monitor access logs** - Review Microsoft Entra ID sign-in logs regularly
5. **Principle of least privilege** - Only grant necessary permissions

## Next Steps

After successful setup:

1. **Explore Available Tools**: See [README.md](README.md#available-tools) for a complete list
2. **Try Example Prompts**: Test different features with Claude
3. **Customize Permissions**: Adjust based on your needs
4. **Monitor Usage**: Check Microsoft Entra ID logs and Graph API usage

## Getting Help

If you encounter issues:

1. Check the [Troubleshooting](README.md#troubleshooting) section in README.md
2. Review Microsoft Entra ID error logs in the Azure Portal
3. Open an issue on GitHub with:
   - Error messages (remove sensitive data)
   - Steps to reproduce
   - Your environment (OS, Node.js version)

## Additional Resources

- [Microsoft Graph API Documentation](https://docs.microsoft.com/en-us/graph/)
- [Microsoft Entra ID App Registration Guide](https://learn.microsoft.com/en-us/entra/identity-platform/quickstart-register-app)
- [Model Context Protocol Specification](https://modelcontextprotocol.io)
- [MSAL Node Documentation](https://github.com/AzureAD/microsoft-authentication-library-for-js/tree/dev/lib/msal-node)
