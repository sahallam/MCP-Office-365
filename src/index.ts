#!/usr/bin/env node

/**
 * Office365 MCP Server
 *
 * A Model Context Protocol server for Microsoft Office 365 integration
 * Provides tools for interacting with Outlook, Calendar, OneDrive, SharePoint, Teams, Excel, Word, and OneNote
 */

import { Server } from '@modelcontextprotocol/sdk/server/index.js';
import { StdioServerTransport } from '@modelcontextprotocol/sdk/server/stdio.js';
import {
  CallToolRequestSchema,
  ListToolsRequestSchema,
  Tool,
} from '@modelcontextprotocol/sdk/types.js';
import * as dotenv from 'dotenv';
import { GraphAuthProvider, AuthenticationError } from './auth.js';
import { OutlookTools } from './tools/outlook.js';
import { CalendarTools } from './tools/calendar.js';
import { OneDriveTools } from './tools/onedrive.js';
import { SharePointTools } from './tools/sharepoint.js';
import { TeamsTools } from './tools/teams.js';
import { ExcelTools } from './tools/excel.js';
import { WordTools } from './tools/word.js';
import { OneNoteTools } from './tools/onenote.js';
import { GraphConfig } from './types.js';
import { getPublicErrorMessage } from './security.js';
import { auditLogger, AuditEventType } from './audit.js';

// Load environment variables
dotenv.config();

// Determine authentication mode
const authMode = (process.env.AUTH_MODE || 'app-only') as 'app-only' | 'delegated';

// Validate required environment variables
const requiredEnvVars = ['TENANT_ID', 'CLIENT_ID'];
for (const envVar of requiredEnvVars) {
  if (!process.env[envVar]) {
    console.error(`Error: ${envVar} environment variable is required`);
    process.exit(1);
  }
}

// For app-only auth, CLIENT_SECRET is required
if (authMode === 'app-only' && !process.env.CLIENT_SECRET) {
  console.error('Error: CLIENT_SECRET is required for app-only authentication');
  console.error('Set AUTH_MODE=delegated to use delegated authentication instead');
  process.exit(1);
}

// For app-only auth, either USER_PRINCIPAL_NAME or USER_ID is required
if (authMode === 'app-only' && !process.env.USER_PRINCIPAL_NAME && !process.env.USER_ID) {
  console.error('Error: Either USER_PRINCIPAL_NAME or USER_ID must be specified for app-only authentication');
  console.error('Example: USER_PRINCIPAL_NAME=user@yourdomain.com');
  console.error('Or set AUTH_MODE=delegated to use delegated authentication');
  process.exit(1);
}

// Configure Graph API
const config: GraphConfig = {
  tenantId: process.env.TENANT_ID!,
  clientId: process.env.CLIENT_ID!,
  clientSecret: process.env.CLIENT_SECRET,
  graphEndpoint: process.env.GRAPH_API_ENDPOINT || 'https://graph.microsoft.com/v1.0',
  userPrincipalName: process.env.USER_PRINCIPAL_NAME,
  userId: process.env.USER_ID,
  authMode,
  tokenCachePath: process.env.TOKEN_CACHE_PATH,
};

// Log authentication mode
// Initialize auth provider silently
const authProvider = new GraphAuthProvider(config);

// Define all available tools
const TOOLS: Tool[] = [
  // Authentication Tool
  {
    name: 'auth_status',
    description: 'Check authentication status and get instructions if re-authentication is needed',
    inputSchema: {
      type: 'object',
      properties: {},
    },
  },

  // Outlook/Email Tools
  {
    name: 'outlook_list_emails',
    description: 'List emails from inbox with optional filtering',
    inputSchema: {
      type: 'object',
      properties: {
        top: { type: 'number', description: 'Number of emails to retrieve (default: 10)' },
        filter: { type: 'string', description: 'OData filter string' },
        orderBy: { type: 'string', description: 'Order by field (default: receivedDateTime DESC)' },
      },
    },
  },
  {
    name: 'outlook_get_email',
    description: 'Get a specific email by ID',
    inputSchema: {
      type: 'object',
      properties: {
        messageId: { type: 'string', description: 'The email message ID' },
      },
      required: ['messageId'],
    },
  },
  {
    name: 'outlook_send_email',
    description: 'Send an email',
    inputSchema: {
      type: 'object',
      properties: {
        subject: { type: 'string', description: 'Email subject' },
        body: { type: 'string', description: 'Email body content' },
        bodyType: { type: 'string', enum: ['HTML', 'Text'], description: 'Body content type (default: Text)' },
        toRecipients: { type: 'array', items: { type: 'string' }, description: 'Array of recipient email addresses' },
        ccRecipients: { type: 'array', items: { type: 'string' }, description: 'Array of CC recipient email addresses' },
        importance: { type: 'string', enum: ['low', 'normal', 'high'], description: 'Email importance' },
      },
      required: ['subject', 'body', 'toRecipients'],
    },
  },
  {
    name: 'outlook_search_emails',
    description: 'Search emails by query string',
    inputSchema: {
      type: 'object',
      properties: {
        query: { type: 'string', description: 'Search query string' },
        top: { type: 'number', description: 'Number of results (default: 10)' },
      },
      required: ['query'],
    },
  },
  {
    name: 'outlook_reply_email',
    description: 'Reply to an email',
    inputSchema: {
      type: 'object',
      properties: {
        messageId: { type: 'string', description: 'The email message ID to reply to' },
        comment: { type: 'string', description: 'Reply message content' },
        replyAll: { type: 'boolean', description: 'Reply to all recipients (default: false)' },
      },
      required: ['messageId', 'comment'],
    },
  },

  // Calendar Tools
  {
    name: 'calendar_list_events',
    description: 'List calendar events',
    inputSchema: {
      type: 'object',
      properties: {
        top: { type: 'number', description: 'Number of events to retrieve (default: 10)' },
        filter: { type: 'string', description: 'OData filter string' },
      },
    },
  },
  {
    name: 'calendar_get_view',
    description: 'Get calendar view for a specific time range',
    inputSchema: {
      type: 'object',
      properties: {
        startDateTime: { type: 'string', description: 'Start date/time (ISO 8601 format)' },
        endDateTime: { type: 'string', description: 'End date/time (ISO 8601 format)' },
      },
      required: ['startDateTime', 'endDateTime'],
    },
  },
  {
    name: 'calendar_create_event',
    description: 'Create a calendar event',
    inputSchema: {
      type: 'object',
      properties: {
        subject: { type: 'string', description: 'Event subject/title' },
        startDateTime: { type: 'string', description: 'Start date/time (ISO 8601 format)' },
        endDateTime: { type: 'string', description: 'End date/time (ISO 8601 format)' },
        timeZone: { type: 'string', description: 'Time zone (default: UTC)' },
        location: { type: 'string', description: 'Event location' },
        body: { type: 'string', description: 'Event description' },
        attendees: { type: 'array', items: { type: 'string' }, description: 'Array of attendee email addresses' },
        isOnlineMeeting: { type: 'boolean', description: 'Create as online meeting (default: false)' },
      },
      required: ['subject', 'startDateTime', 'endDateTime'],
    },
  },
  {
    name: 'calendar_delete_event',
    description: 'Delete a calendar event',
    inputSchema: {
      type: 'object',
      properties: {
        eventId: { type: 'string', description: 'The event ID to delete' },
      },
      required: ['eventId'],
    },
  },
  {
    name: 'calendar_update_event',
    description: 'Update a calendar event - modify subject, time, location, attendees, etc.',
    inputSchema: {
      type: 'object',
      properties: {
        eventId: { type: 'string', description: 'The event ID to update' },
        subject: { type: 'string', description: 'Updated event subject/title' },
        startDateTime: { type: 'string', description: 'Updated start date/time (ISO 8601 format)' },
        endDateTime: { type: 'string', description: 'Updated end date/time (ISO 8601 format)' },
        timeZone: { type: 'string', description: 'Time zone (default: UTC)' },
        location: { type: 'string', description: 'Updated event location' },
        body: { type: 'string', description: 'Updated event description' },
        attendees: { type: 'array', items: { type: 'string' }, description: 'Updated array of attendee email addresses' },
        isOnlineMeeting: { type: 'boolean', description: 'Update online meeting status' },
      },
      required: ['eventId'],
    },
  },
  {
    name: 'calendar_get_event',
    description: 'Get a specific calendar event by ID',
    inputSchema: {
      type: 'object',
      properties: {
        eventId: { type: 'string', description: 'The event ID to retrieve' },
      },
      required: ['eventId'],
    },
  },
  {
    name: 'calendar_accept_meeting',
    description: 'Accept a meeting invitation',
    inputSchema: {
      type: 'object',
      properties: {
        eventId: { type: 'string', description: 'The event ID to accept' },
        comment: { type: 'string', description: 'Optional comment to send with the response' },
      },
      required: ['eventId'],
    },
  },
  {
    name: 'calendar_decline_meeting',
    description: 'Decline a meeting invitation',
    inputSchema: {
      type: 'object',
      properties: {
        eventId: { type: 'string', description: 'The event ID to decline' },
        comment: { type: 'string', description: 'Optional comment to send with the response' },
      },
      required: ['eventId'],
    },
  },
  {
    name: 'calendar_tentatively_accept_meeting',
    description: 'Tentatively accept a meeting invitation',
    inputSchema: {
      type: 'object',
      properties: {
        eventId: { type: 'string', description: 'The event ID to tentatively accept' },
        comment: { type: 'string', description: 'Optional comment to send with the response' },
      },
      required: ['eventId'],
    },
  },
  {
    name: 'calendar_find_meeting_times',
    description: 'Find available meeting times that work for all attendees',
    inputSchema: {
      type: 'object',
      properties: {
        attendees: {
          type: 'array',
          items: { type: 'string' },
          description: 'Array of attendee email addresses'
        },
        startDateTime: {
          type: 'string',
          description: 'Start of time window to search (ISO 8601 format)'
        },
        endDateTime: {
          type: 'string',
          description: 'End of time window to search (ISO 8601 format)'
        },
        timeZone: {
          type: 'string',
          description: 'Time zone for the search (default: UTC)'
        },
        meetingDuration: {
          type: 'string',
          description: 'Duration in ISO 8601 format (e.g., PT1H for 1 hour, PT30M for 30 minutes)'
        },
        maxCandidates: {
          type: 'number',
          description: 'Maximum number of time suggestions to return (default: 5)'
        },
      },
      required: ['attendees', 'startDateTime', 'endDateTime', 'meetingDuration'],
    },
  },
  {
    name: 'calendar_list_calendars',
    description: 'List all calendars for the user',
    inputSchema: {
      type: 'object',
      properties: {},
    },
  },

  // OneDrive Tools
  {
    name: 'onedrive_list_items',
    description: 'List items in OneDrive root or a specific folder',
    inputSchema: {
      type: 'object',
      properties: {
        folderId: { type: 'string', description: 'Folder ID (optional, lists root if not provided)' },
      },
    },
  },
  {
    name: 'onedrive_get_item',
    description: 'Get a OneDrive item by ID',
    inputSchema: {
      type: 'object',
      properties: {
        itemId: { type: 'string', description: 'The item ID' },
      },
      required: ['itemId'],
    },
  },
  {
    name: 'onedrive_download_file',
    description: 'Download a file from OneDrive',
    inputSchema: {
      type: 'object',
      properties: {
        itemId: { type: 'string', description: 'The file item ID' },
      },
      required: ['itemId'],
    },
  },
  {
    name: 'onedrive_upload_file',
    description: 'Upload a file to OneDrive',
    inputSchema: {
      type: 'object',
      properties: {
        fileName: { type: 'string', description: 'Name of the file' },
        content: { type: 'string', description: 'File content (base64 encoded for binary files)' },
        parentFolderId: { type: 'string', description: 'Parent folder ID (optional)' },
      },
      required: ['fileName', 'content'],
    },
  },
  {
    name: 'onedrive_search',
    description: 'Search for files in OneDrive',
    inputSchema: {
      type: 'object',
      properties: {
        query: { type: 'string', description: 'Search query string' },
      },
      required: ['query'],
    },
  },
  {
    name: 'onedrive_create_folder',
    description: 'Create a folder in OneDrive',
    inputSchema: {
      type: 'object',
      properties: {
        folderName: { type: 'string', description: 'Name of the folder' },
        parentFolderId: { type: 'string', description: 'Parent folder ID (optional)' },
      },
      required: ['folderName'],
    },
  },

  // SharePoint Tools
  {
    name: 'sharepoint_search_sites',
    description: 'Search for SharePoint sites',
    inputSchema: {
      type: 'object',
      properties: {
        query: { type: 'string', description: 'Search query string' },
      },
      required: ['query'],
    },
  },
  {
    name: 'sharepoint_get_site',
    description: 'Get a SharePoint site by ID',
    inputSchema: {
      type: 'object',
      properties: {
        siteId: { type: 'string', description: 'The site ID' },
      },
      required: ['siteId'],
    },
  },
  {
    name: 'sharepoint_list_libraries',
    description: 'List document libraries in a SharePoint site',
    inputSchema: {
      type: 'object',
      properties: {
        siteId: { type: 'string', description: 'The site ID' },
      },
      required: ['siteId'],
    },
  },
  {
    name: 'sharepoint_list_items',
    description: 'List items in a SharePoint document library',
    inputSchema: {
      type: 'object',
      properties: {
        siteId: { type: 'string', description: 'The site ID' },
        driveId: { type: 'string', description: 'The document library (drive) ID' },
        folderId: { type: 'string', description: 'Folder ID (optional)' },
      },
      required: ['siteId', 'driveId'],
    },
  },

  // Teams Tools
  {
    name: 'teams_list_teams',
    description: 'List all teams the user is a member of',
    inputSchema: {
      type: 'object',
      properties: {},
    },
  },
  {
    name: 'teams_list_channels',
    description: 'List channels in a team',
    inputSchema: {
      type: 'object',
      properties: {
        teamId: { type: 'string', description: 'The team ID' },
      },
      required: ['teamId'],
    },
  },
  {
    name: 'teams_send_message',
    description: 'Send a message to a Teams channel',
    inputSchema: {
      type: 'object',
      properties: {
        teamId: { type: 'string', description: 'The team ID' },
        channelId: { type: 'string', description: 'The channel ID' },
        content: { type: 'string', description: 'Message content' },
        contentType: { type: 'string', enum: ['text', 'html'], description: 'Content type (default: text)' },
      },
      required: ['teamId', 'channelId', 'content'],
    },
  },
  {
    name: 'teams_list_messages',
    description: 'List messages in a Teams channel',
    inputSchema: {
      type: 'object',
      properties: {
        teamId: { type: 'string', description: 'The team ID' },
        channelId: { type: 'string', description: 'The channel ID' },
        top: { type: 'number', description: 'Number of messages to retrieve (default: 50)' },
      },
      required: ['teamId', 'channelId'],
    },
  },

  // Excel Tools
  {
    name: 'excel_list_worksheets',
    description: 'List worksheets in an Excel workbook',
    inputSchema: {
      type: 'object',
      properties: {
        itemId: { type: 'string', description: 'The Excel file item ID' },
      },
      required: ['itemId'],
    },
  },
  {
    name: 'excel_get_range',
    description: 'Get range values from an Excel worksheet',
    inputSchema: {
      type: 'object',
      properties: {
        itemId: { type: 'string', description: 'The Excel file item ID' },
        worksheetId: { type: 'string', description: 'The worksheet ID or name' },
        address: { type: 'string', description: 'Range address (e.g., "A1:B10")' },
      },
      required: ['itemId', 'worksheetId', 'address'],
    },
  },
  {
    name: 'excel_update_range',
    description: 'Update range values in an Excel worksheet',
    inputSchema: {
      type: 'object',
      properties: {
        itemId: { type: 'string', description: 'The Excel file item ID' },
        worksheetId: { type: 'string', description: 'The worksheet ID or name' },
        address: { type: 'string', description: 'Range address (e.g., "A1:B10")' },
        values: { type: 'array', description: '2D array of values to write' },
      },
      required: ['itemId', 'worksheetId', 'address', 'values'],
    },
  },
  {
    name: 'excel_get_used_range',
    description: 'Get the used range (non-empty cells) from an Excel worksheet',
    inputSchema: {
      type: 'object',
      properties: {
        itemId: { type: 'string', description: 'The Excel file item ID' },
        worksheetId: { type: 'string', description: 'The worksheet ID or name' },
      },
      required: ['itemId', 'worksheetId'],
    },
  },

  // Word Tools
  {
    name: 'word_get_document',
    description: 'Get Word document metadata',
    inputSchema: {
      type: 'object',
      properties: {
        itemId: { type: 'string', description: 'The Word document item ID' },
      },
      required: ['itemId'],
    },
  },
  {
    name: 'word_search_documents',
    description: 'Search for Word documents',
    inputSchema: {
      type: 'object',
      properties: {
        query: { type: 'string', description: 'Search query string' },
      },
      required: ['query'],
    },
  },
  {
    name: 'word_convert_to_pdf',
    description: 'Convert a Word document to PDF',
    inputSchema: {
      type: 'object',
      properties: {
        itemId: { type: 'string', description: 'The Word document item ID' },
      },
      required: ['itemId'],
    },
  },

  // OneNote Tools
  {
    name: 'onenote_list_notebooks',
    description: 'List all OneNote notebooks',
    inputSchema: {
      type: 'object',
      properties: {
        top: { type: 'number', description: 'Number of notebooks to retrieve (default: 50)' },
        orderBy: { type: 'string', description: 'Order by field (default: lastModifiedDateTime DESC)' },
      },
    },
  },
  {
    name: 'onenote_get_notebook',
    description: 'Get a specific OneNote notebook by ID',
    inputSchema: {
      type: 'object',
      properties: {
        notebookId: { type: 'string', description: 'The notebook ID' },
      },
      required: ['notebookId'],
    },
  },
  {
    name: 'onenote_create_notebook',
    description: 'Create a new OneNote notebook',
    inputSchema: {
      type: 'object',
      properties: {
        displayName: { type: 'string', description: 'Name of the notebook' },
      },
      required: ['displayName'],
    },
  },
  {
    name: 'onenote_list_sections',
    description: 'List sections in a notebook or all sections',
    inputSchema: {
      type: 'object',
      properties: {
        notebookId: { type: 'string', description: 'The notebook ID (optional, lists all sections if not provided)' },
        top: { type: 'number', description: 'Number of sections to retrieve (default: 50)' },
      },
    },
  },
  {
    name: 'onenote_get_section',
    description: 'Get a specific OneNote section by ID',
    inputSchema: {
      type: 'object',
      properties: {
        sectionId: { type: 'string', description: 'The section ID' },
      },
      required: ['sectionId'],
    },
  },
  {
    name: 'onenote_create_section',
    description: 'Create a new section in a notebook',
    inputSchema: {
      type: 'object',
      properties: {
        notebookId: { type: 'string', description: 'The notebook ID' },
        displayName: { type: 'string', description: 'Name of the section' },
      },
      required: ['notebookId', 'displayName'],
    },
  },
  {
    name: 'onenote_list_pages',
    description: 'List pages in a section, notebook, or all pages',
    inputSchema: {
      type: 'object',
      properties: {
        sectionId: { type: 'string', description: 'The section ID (optional)' },
        notebookId: { type: 'string', description: 'The notebook ID (optional)' },
        top: { type: 'number', description: 'Number of pages to retrieve (default: 50)' },
        search: { type: 'string', description: 'Search query string (optional)' },
      },
    },
  },
  {
    name: 'onenote_get_page',
    description: 'Get a specific OneNote page by ID',
    inputSchema: {
      type: 'object',
      properties: {
        pageId: { type: 'string', description: 'The page ID' },
      },
      required: ['pageId'],
    },
  },
  {
    name: 'onenote_get_page_content',
    description: 'Get the HTML content of a OneNote page',
    inputSchema: {
      type: 'object',
      properties: {
        pageId: { type: 'string', description: 'The page ID' },
      },
      required: ['pageId'],
    },
  },
  {
    name: 'onenote_create_page',
    description: 'Create a new page in a section',
    inputSchema: {
      type: 'object',
      properties: {
        sectionId: { type: 'string', description: 'The section ID' },
        title: { type: 'string', description: 'Page title' },
        content: { type: 'string', description: 'Page content (HTML)' },
      },
      required: ['sectionId', 'title', 'content'],
    },
  },
  {
    name: 'onenote_append_to_page',
    description: 'Append content to an existing OneNote page',
    inputSchema: {
      type: 'object',
      properties: {
        pageId: { type: 'string', description: 'The page ID' },
        content: { type: 'string', description: 'Content to append (HTML)' },
      },
      required: ['pageId', 'content'],
    },
  },
  {
    name: 'onenote_search_pages',
    description: 'Search pages across all notebooks',
    inputSchema: {
      type: 'object',
      properties: {
        query: { type: 'string', description: 'Search query string' },
        top: { type: 'number', description: 'Number of results (default: 20)' },
      },
      required: ['query'],
    },
  },
  {
    name: 'onenote_delete_page',
    description: 'Delete a OneNote page',
    inputSchema: {
      type: 'object',
      properties: {
        pageId: { type: 'string', description: 'The page ID to delete' },
      },
      required: ['pageId'],
    },
  },
  {
    name: 'onenote_delete_section',
    description: 'Delete a OneNote section',
    inputSchema: {
      type: 'object',
      properties: {
        sectionId: { type: 'string', description: 'The section ID to delete' },
      },
      required: ['sectionId'],
    },
  },
  {
    name: 'onenote_delete_notebook',
    description: 'Delete a OneNote notebook',
    inputSchema: {
      type: 'object',
      properties: {
        notebookId: { type: 'string', description: 'The notebook ID to delete' },
      },
      required: ['notebookId'],
    },
  },
  {
    name: 'onenote_copy_page',
    description: 'Copy a page to a different section',
    inputSchema: {
      type: 'object',
      properties: {
        pageId: { type: 'string', description: 'The page ID to copy' },
        targetSectionId: { type: 'string', description: 'The target section ID' },
      },
      required: ['pageId', 'targetSectionId'],
    },
  },
];

// Initialize MCP server
const server = new Server(
  {
    name: 'office365-mcp-server',
    version: '1.0.0',
  },
  {
    capabilities: {
      tools: {},
    },
  }
);

// Handle tool list request
server.setRequestHandler(ListToolsRequestSchema, async () => {
  return { tools: TOOLS };
});

// Handle tool execution
server.setRequestHandler(CallToolRequestSchema, async (request) => {
  const { name, arguments: args } = request.params;

  try {
    // Get authenticated Graph client
    const graphClient = await authProvider.getGraphClient();
    const userId = authProvider.getUserId();

    // Initialize tool classes
    const outlook = new OutlookTools(graphClient, userId);
    const calendar = new CalendarTools(graphClient, userId);
    const onedrive = new OneDriveTools(graphClient, userId);
    const sharepoint = new SharePointTools(graphClient);
    const teams = new TeamsTools(graphClient, userId);
    const excel = new ExcelTools(graphClient, userId);
    const word = new WordTools(graphClient, userId);
    const onenote = new OneNoteTools(graphClient, userId);

    let result: any;

    switch (name) {
      // Authentication Tool
      case 'auth_status': {
        const isAuthenticated = authProvider.isAuthenticated();
        result = {
          authenticated: isAuthenticated,
          status: isAuthenticated ? 'active' : 'authentication_required',
          message: isAuthenticated
            ? 'Authentication is active and working properly.'
            : 'Authentication required. Please check the server logs for authentication instructions.',
          authMode: config.authMode,
        };
        break;
      }

      // Outlook/Email Tools
      case 'outlook_list_emails':
        result = await outlook.listEmails(args as any);
        break;
      case 'outlook_get_email':
        result = await outlook.getEmail((args as any).messageId);
        break;
      case 'outlook_send_email': {
        const { subject, body, bodyType, toRecipients, ccRecipients, importance } = args as any;
        await outlook.sendEmail({
          subject,
          body: { contentType: bodyType || 'Text', content: body },
          toRecipients: toRecipients.map((email: string) => ({ emailAddress: { address: email } })),
          ccRecipients: ccRecipients?.map((email: string) => ({ emailAddress: { address: email } })),
          importance,
        });
        result = { success: true, message: 'Email sent successfully' };
        break;
      }
      case 'outlook_search_emails':
        result = await outlook.searchEmails((args as any).query, (args as any).top);
        break;
      case 'outlook_reply_email':
        await outlook.replyToEmail((args as any).messageId, (args as any).comment, (args as any).replyAll);
        result = { success: true, message: 'Reply sent successfully' };
        break;

      // Calendar Tools
      case 'calendar_list_events':
        result = await calendar.listEvents(args as any);
        break;
      case 'calendar_get_view':
        result = await calendar.getCalendarView((args as any).startDateTime, (args as any).endDateTime);
        break;
      case 'calendar_create_event': {
        const { subject, startDateTime, endDateTime, timeZone, location, body: eventBody, attendees, isOnlineMeeting } = args as any;
        result = await calendar.createEvent({
          subject,
          start: { dateTime: startDateTime, timeZone: timeZone || 'UTC' },
          end: { dateTime: endDateTime, timeZone: timeZone || 'UTC' },
          location: location ? { displayName: location } : undefined,
          body: eventBody ? { contentType: 'Text', content: eventBody } : undefined,
          attendees: attendees?.map((email: string) => ({ emailAddress: { address: email }, type: 'required' as const })),
          isOnlineMeeting,
        });
        break;
      }
      case 'calendar_delete_event':
        await calendar.deleteEvent((args as any).eventId);
        result = { success: true, message: 'Event deleted successfully' };
        break;
      case 'calendar_update_event': {
        const { eventId, subject, startDateTime, endDateTime, timeZone, location, body: eventBody, attendees, isOnlineMeeting } = args as any;

        // Build the updates object with only provided fields
        const updates: any = {};

        if (subject !== undefined) updates.subject = subject;
        if (startDateTime !== undefined || endDateTime !== undefined || timeZone !== undefined) {
          if (startDateTime) updates.start = { dateTime: startDateTime, timeZone: timeZone || 'UTC' };
          if (endDateTime) updates.end = { dateTime: endDateTime, timeZone: timeZone || 'UTC' };
        }
        if (location !== undefined) updates.location = { displayName: location };
        if (eventBody !== undefined) updates.body = { contentType: 'Text', content: eventBody };
        if (attendees !== undefined) {
          updates.attendees = attendees.map((email: string) => ({
            emailAddress: { address: email },
            type: 'required' as const
          }));
        }
        if (isOnlineMeeting !== undefined) updates.isOnlineMeeting = isOnlineMeeting;

        result = await calendar.updateEvent(eventId, updates);
        break;
      }
      case 'calendar_get_event':
        result = await calendar.getEvent((args as any).eventId);
        break;
      case 'calendar_accept_meeting':
        await calendar.acceptMeeting((args as any).eventId, (args as any).comment);
        result = { success: true, message: 'Meeting accepted successfully' };
        break;
      case 'calendar_decline_meeting':
        await calendar.declineMeeting((args as any).eventId, (args as any).comment);
        result = { success: true, message: 'Meeting declined successfully' };
        break;
      case 'calendar_tentatively_accept_meeting':
        await calendar.tentativelyAcceptMeeting((args as any).eventId, (args as any).comment);
        result = { success: true, message: 'Meeting tentatively accepted successfully' };
        break;
      case 'calendar_find_meeting_times': {
        const { attendees, startDateTime, endDateTime, timeZone, meetingDuration, maxCandidates } = args as any;
        result = await calendar.findMeetingTimes({
          attendees,
          timeConstraint: {
            timeslots: [{
              start: { dateTime: startDateTime, timeZone: timeZone || 'UTC' },
              end: { dateTime: endDateTime, timeZone: timeZone || 'UTC' }
            }]
          },
          meetingDuration,
          maxCandidates,
        });
        break;
      }
      case 'calendar_list_calendars':
        result = await calendar.listCalendars();
        break;

      // OneDrive Tools
      case 'onedrive_list_items':
        result = await onedrive.listItems((args as any).folderId);
        break;
      case 'onedrive_get_item':
        result = await onedrive.getItem((args as any).itemId);
        break;
      case 'onedrive_download_file':
        result = await onedrive.downloadFile((args as any).itemId);
        break;
      case 'onedrive_upload_file':
        result = await onedrive.uploadFile((args as any).fileName, (args as any).content, (args as any).parentFolderId);
        break;
      case 'onedrive_search':
        result = await onedrive.searchItems((args as any).query);
        break;
      case 'onedrive_create_folder':
        result = await onedrive.createFolder((args as any).folderName, (args as any).parentFolderId);
        break;

      // SharePoint Tools
      case 'sharepoint_search_sites':
        result = await sharepoint.searchSites((args as any).query);
        break;
      case 'sharepoint_get_site':
        result = await sharepoint.getSite((args as any).siteId);
        break;
      case 'sharepoint_list_libraries':
        result = await sharepoint.listDocumentLibraries((args as any).siteId);
        break;
      case 'sharepoint_list_items':
        result = await sharepoint.listLibraryItems((args as any).siteId, (args as any).driveId, (args as any).folderId);
        break;

      // Teams Tools
      case 'teams_list_teams':
        result = await teams.listTeams();
        break;
      case 'teams_list_channels':
        result = await teams.listChannels((args as any).teamId);
        break;
      case 'teams_send_message':
        result = await teams.sendChannelMessage(
          (args as any).teamId,
          (args as any).channelId,
          (args as any).content,
          (args as any).contentType || 'text'
        );
        break;
      case 'teams_list_messages':
        result = await teams.listChannelMessages((args as any).teamId, (args as any).channelId, (args as any).top);
        break;

      // Excel Tools
      case 'excel_list_worksheets':
        result = await excel.listWorksheets((args as any).itemId);
        break;
      case 'excel_get_range':
        result = await excel.getRange((args as any).itemId, (args as any).worksheetId, (args as any).address);
        break;
      case 'excel_update_range':
        result = await excel.updateRange((args as any).itemId, (args as any).worksheetId, (args as any).address, (args as any).values);
        break;
      case 'excel_get_used_range':
        result = await excel.getUsedRange((args as any).itemId, (args as any).worksheetId);
        break;

      // Word Tools
      case 'word_get_document':
        result = await word.getDocument((args as any).itemId);
        break;
      case 'word_search_documents':
        result = await word.searchDocuments((args as any).query);
        break;
      case 'word_convert_to_pdf':
        result = await word.convertToPdf((args as any).itemId);
        break;

      // OneNote Tools
      case 'onenote_list_notebooks':
        result = await onenote.listNotebooks(args as any);
        break;
      case 'onenote_get_notebook':
        result = await onenote.getNotebook((args as any).notebookId);
        break;
      case 'onenote_create_notebook':
        result = await onenote.createNotebook((args as any).displayName);
        break;
      case 'onenote_list_sections':
        result = await onenote.listSections(args as any);
        break;
      case 'onenote_get_section':
        result = await onenote.getSection((args as any).sectionId);
        break;
      case 'onenote_create_section':
        result = await onenote.createSection((args as any).notebookId, (args as any).displayName);
        break;
      case 'onenote_list_pages':
        result = await onenote.listPages(args as any);
        break;
      case 'onenote_get_page':
        result = await onenote.getPage((args as any).pageId);
        break;
      case 'onenote_get_page_content':
        result = await onenote.getPageContent((args as any).pageId);
        break;
      case 'onenote_create_page':
        result = await onenote.createPage((args as any).sectionId, (args as any).title, (args as any).content);
        break;
      case 'onenote_append_to_page':
        await onenote.appendToPage((args as any).pageId, (args as any).content);
        result = { success: true, message: 'Content appended successfully' };
        break;
      case 'onenote_search_pages':
        result = await onenote.searchPages((args as any).query, (args as any).top);
        break;
      case 'onenote_delete_page':
        await onenote.deletePage((args as any).pageId);
        result = { success: true, message: 'Page deleted successfully' };
        break;
      case 'onenote_delete_section':
        await onenote.deleteSection((args as any).sectionId);
        result = { success: true, message: 'Section deleted successfully' };
        break;
      case 'onenote_delete_notebook':
        await onenote.deleteNotebook((args as any).notebookId);
        result = { success: true, message: 'Notebook deleted successfully' };
        break;
      case 'onenote_copy_page':
        result = await onenote.copyPage((args as any).pageId, (args as any).targetSectionId);
        break;

      default:
        throw new Error(`Unknown tool: ${name}`);
    }

    // Audit log successful tool execution
    const auditType = name.includes('delete') ? AuditEventType.RESOURCE_DELETE :
                      name.includes('create') || name.includes('send') || name.includes('upload') ? AuditEventType.RESOURCE_CREATE :
                      name.includes('update') || name.includes('accept') || name.includes('decline') ? AuditEventType.RESOURCE_UPDATE :
                      name.includes('search') || name.includes('find') ? AuditEventType.RESOURCE_SEARCH :
                      AuditEventType.RESOURCE_READ;

    auditLogger.logSuccess(
      auditType,
      name,
      `tool:${name}`,
      userId,
      `Args: ${JSON.stringify(args).substring(0, 100)}`
    );

    return {
      content: [
        {
          type: 'text',
          text: JSON.stringify(result, null, 2),
        },
      ],
    };
  } catch (error) {
    // Log detailed error internally for debugging
    console.error(`[MCP] Tool execution error for ${name}:`, error);

    // Audit log failed tool execution
    auditLogger.logFailure(
      AuditEventType.RESOURCE_READ,
      name,
      `tool:${name}`,
      error instanceof Error ? error.message : 'Unknown error',
      authProvider.getUserId(),
      `Args: ${JSON.stringify(args).substring(0, 100)}`
    );

    // Check if this is an authentication error
    if (error instanceof AuthenticationError) {
      // Build authentication error message
      let authMessage = `🔐 AUTHENTICATION REQUIRED\n\n${error.userMessage}\n\n`;

      // Include device code information if available
      if (error.deviceCode && error.verificationUri) {
        const expiresInMinutes = error.expiresIn ? Math.floor(error.expiresIn / 60) : 15;
        authMessage += `Please sign in using your web browser:\n\n`;
        authMessage += `1. Open this URL: ${error.verificationUri}\n`;
        authMessage += `2. Enter this code: ${error.deviceCode}\n`;
        authMessage += `3. Sign in with your Microsoft account\n\n`;
        authMessage += `⏱️  Code expires in ${expiresInMinutes} minutes\n\n`;
        authMessage += `Once authenticated, your session will remain valid for approximately 90 days\n`;
        authMessage += `(or longer depending on your organization's token policies).`;
      } else if (error.requiresReauth) {
        authMessage += 'Your session has expired. Please check the Office 365 MCP server logs for authentication instructions.\n\n';
        authMessage += 'The server logs will display a device code and URL for signing in.\n\n';
        authMessage += 'Once authenticated, your session will remain valid for approximately 90 days\n';
        authMessage += '(or longer depending on your organization\'s policies).';
      } else {
        authMessage += 'Please check the Office 365 MCP server logs for detailed error information and authentication instructions.';
      }

      // Return user-friendly authentication error message
      return {
        content: [
          {
            type: 'text',
            text: authMessage,
          },
        ],
        isError: true,
      };
    }

    // Return sanitized error message for other errors
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
});

// Start the server
async function main() {
  const transport = new StdioServerTransport();
  await server.connect(transport);
}

main().catch((error) => {
  console.error('Fatal error:', error);
  process.exit(1);
});
