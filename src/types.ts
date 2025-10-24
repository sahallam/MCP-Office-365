/**
 * Type definitions for Office365 MCP Server
 */

export interface GraphConfig {
  tenantId: string;
  clientId: string;
  clientSecret?: string;
  graphEndpoint: string;
  userPrincipalName?: string;
  userId?: string;
  authMode?: 'app-only' | 'delegated';
  tokenCachePath?: string;
}

export interface EmailMessage {
  id?: string;
  subject: string;
  body: {
    contentType: 'HTML' | 'Text';
    content: string;
  };
  toRecipients: Array<{ emailAddress: { address: string; name?: string } }>;
  ccRecipients?: Array<{ emailAddress: { address: string; name?: string } }>;
  bccRecipients?: Array<{ emailAddress: { address: string; name?: string } }>;
  from?: { emailAddress: { address: string; name?: string } };
  receivedDateTime?: string;
  hasAttachments?: boolean;
  importance?: 'low' | 'normal' | 'high';
  isRead?: boolean;
}

export interface CalendarEvent {
  id?: string;
  subject: string;
  body?: {
    contentType: 'HTML' | 'Text';
    content: string;
  };
  start: {
    dateTime: string;
    timeZone: string;
  };
  end: {
    dateTime: string;
    timeZone: string;
  };
  location?: {
    displayName: string;
  };
  attendees?: Array<{
    emailAddress: { address: string; name?: string };
    type: 'required' | 'optional' | 'resource';
  }>;
  isOnlineMeeting?: boolean;
  onlineMeetingUrl?: string;
}

export interface DriveItem {
  id: string;
  name: string;
  size?: number;
  createdDateTime: string;
  lastModifiedDateTime: string;
  webUrl: string;
  folder?: { childCount: number };
  file?: { mimeType: string };
  parentReference?: {
    driveId: string;
    id: string;
    path: string;
  };
}

export interface SharePointSite {
  id: string;
  name: string;
  displayName: string;
  webUrl: string;
  description?: string;
}

export interface TeamsChannel {
  id: string;
  displayName: string;
  description?: string;
  webUrl?: string;
}

export interface TeamsMessage {
  id?: string;
  body: {
    contentType: 'html' | 'text';
    content: string;
  };
  from?: {
    user: {
      displayName: string;
      id: string;
    };
  };
  createdDateTime?: string;
}

export interface ExcelWorkbook {
  id: string;
  name: string;
}

export interface ExcelWorksheet {
  id: string;
  name: string;
  position: number;
}

export interface WordDocument {
  id: string;
  name: string;
  webUrl: string;
}
