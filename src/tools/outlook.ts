/**
 * Outlook/Email tools for MCP server
 */

import { Client } from '@microsoft/microsoft-graph-client';
import { EmailMessage } from '../types.js';
import { sanitizeSearchQuery, validateResourceId, validateContentLength, SIZE_LIMITS } from '../security.js';

/**
 * Formats email body text by converting plain text line breaks to HTML.
 * If the text already contains HTML tags, returns it unchanged.
 * @param text - The email body text to format
 * @returns Formatted HTML string
 */
function formatEmailBody(text: string): string {
  // If it already contains HTML tags, use as-is
  if (/<[a-z][\s\S]*>/i.test(text)) {
    return text;
  }

  // Convert plain text with line breaks to HTML
  // Preserve leading/trailing spaces by converting them to &nbsp;
  return text
    .split('\n')
    .map(line => {
      if (line === '') return '&nbsp;'; // Empty lines become non-breaking spaces
      // Convert leading/trailing spaces to &nbsp; to preserve indentation
      return line.replace(/^ +/, match => '&nbsp;'.repeat(match.length))
                 .replace(/ +$/, match => '&nbsp;'.repeat(match.length));
    })
    .join('<br>');
}

export class OutlookTools {
  constructor(private graphClient: Client, private userId: string) {}

  /**
   * Get the correct user path for API endpoints
   * Returns '/me' if userId is 'me', otherwise '/users/{userId}'
   */
  private getUserPath(): string {
    return this.userId === 'me' ? '/me' : `/users/${this.userId}`;
  }

  /**
   * List emails from inbox
   */
  async listEmails(options: {
    top?: number;
    filter?: string;
    orderBy?: string;
    select?: string[];
  } = {}): Promise<EmailMessage[]> {
    const {
      top = 10,
      filter,
      orderBy = 'receivedDateTime DESC',
      select = ['id', 'subject', 'from', 'receivedDateTime', 'hasAttachments', 'isRead', 'importance']
    } = options;

    let query = this.graphClient
      .api(`${this.getUserPath()}/mailFolders/inbox/messages`)
      .top(top)
      .orderby(orderBy)
      .select(select);

    if (filter) {
      query = query.filter(filter);
    }

    const result = await query.get();
    return result.value;
  }

  /**
   * Get a specific email by ID
   */
  async getEmail(messageId: string): Promise<EmailMessage> {
    validateResourceId(messageId, 'message');

    const message = await this.graphClient
      .api(`${this.getUserPath()}/messages/${messageId}`)
      .get();

    return message;
  }

  /**
   * Send an email
   */
  async sendEmail(message: EmailMessage): Promise<void> {
    // Validate email size limits
    validateContentLength(message.subject, SIZE_LIMITS.MAX_EMAIL_SUBJECT_SIZE, 'Email subject');
    validateContentLength(message.body.content, SIZE_LIMITS.MAX_EMAIL_BODY_SIZE, 'Email body');

    // Format the email body (converts plain text line breaks to HTML)
    const formattedContent = formatEmailBody(message.body.content);

    const mailObject = {
      message: {
        subject: message.subject,
        body: {
          contentType: 'HTML',
          content: formattedContent,
        },
        toRecipients: message.toRecipients,
        ccRecipients: message.ccRecipients || [],
        bccRecipients: message.bccRecipients || [],
        importance: message.importance || 'normal',
      },
      saveToSentItems: true,
    };

    await this.graphClient
      .api(`${this.getUserPath()}/sendMail`)
      .post(mailObject);
  }

  /**
   * Reply to an email
   */
  async replyToEmail(messageId: string, comment: string, replyAll: boolean = false): Promise<void> {
    validateResourceId(messageId, 'message');
    validateContentLength(comment, SIZE_LIMITS.MAX_EMAIL_BODY_SIZE, 'Reply comment');

    const endpoint = replyAll ? 'replyAll' : 'reply';

    // Format the comment (converts plain text line breaks to HTML)
    const formattedComment = formatEmailBody(comment);

    await this.graphClient
      .api(`${this.getUserPath()}/messages/${messageId}/${endpoint}`)
      .post({
        comment: formattedComment,
      });
  }

  /**
   * Search emails
   */
  async searchEmails(searchQuery: string, top: number = 10): Promise<EmailMessage[]> {
    // Sanitize search query to prevent OData injection
    const sanitizedQuery = sanitizeSearchQuery(searchQuery);

    const result = await this.graphClient
      .api(`${this.getUserPath()}/messages`)
      .search(`"${sanitizedQuery}"`)
      .top(top)
      .select(['id', 'subject', 'from', 'receivedDateTime', 'bodyPreview', 'hasAttachments'])
      .get();

    return result.value;
  }

  /**
   * Mark email as read/unread
   */
  async markEmailAsRead(messageId: string, isRead: boolean = true): Promise<void> {
    validateResourceId(messageId, 'message');

    await this.graphClient
      .api(`${this.getUserPath()}/messages/${messageId}`)
      .patch({
        isRead,
      });
  }

  /**
   * Delete an email
   */
  async deleteEmail(messageId: string): Promise<void> {
    validateResourceId(messageId, 'message');

    await this.graphClient
      .api(`${this.getUserPath()}/messages/${messageId}`)
      .delete();
  }

  /**
   * Move email to folder
   */
  async moveEmail(messageId: string, destinationFolderId: string): Promise<void> {
    validateResourceId(messageId, 'message');
    validateResourceId(destinationFolderId, 'folder');

    await this.graphClient
      .api(`${this.getUserPath()}/messages/${messageId}/move`)
      .post({
        destinationId: destinationFolderId,
      });
  }

  /**
   * List mail folders
   */
  async listMailFolders(): Promise<any[]> {
    const result = await this.graphClient
      .api(`${this.getUserPath()}/mailFolders`)
      .get();

    return result.value;
  }

  /**
   * Create a draft email
   */
  async createDraft(message: EmailMessage): Promise<EmailMessage> {
    const draft = await this.graphClient
      .api(`${this.getUserPath()}/messages`)
      .post({
        subject: message.subject,
        body: message.body,
        toRecipients: message.toRecipients,
        ccRecipients: message.ccRecipients,
        bccRecipients: message.bccRecipients,
      });

    return draft;
  }
}
