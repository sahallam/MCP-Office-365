/**
 * Outlook/Email tools for MCP server
 */

import { Client } from '@microsoft/microsoft-graph-client';
import { EmailMessage } from '../types.js';

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
    const message = await this.graphClient
      .api(`${this.getUserPath()}/messages/${messageId}`)
      .get();

    return message;
  }

  /**
   * Send an email
   */
  async sendEmail(message: EmailMessage): Promise<void> {
    const mailObject = {
      message: {
        subject: message.subject,
        body: {
          contentType: message.body.contentType,
          content: message.body.content,
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
    const endpoint = replyAll ? 'replyAll' : 'reply';

    await this.graphClient
      .api(`${this.getUserPath()}/messages/${messageId}/${endpoint}`)
      .post({
        comment,
      });
  }

  /**
   * Search emails
   */
  async searchEmails(searchQuery: string, top: number = 10): Promise<EmailMessage[]> {
    const result = await this.graphClient
      .api(`${this.getUserPath()}/messages`)
      .search(`"${searchQuery}"`)
      .top(top)
      .select(['id', 'subject', 'from', 'receivedDateTime', 'bodyPreview', 'hasAttachments'])
      .get();

    return result.value;
  }

  /**
   * Mark email as read/unread
   */
  async markEmailAsRead(messageId: string, isRead: boolean = true): Promise<void> {
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
    await this.graphClient
      .api(`${this.getUserPath()}/messages/${messageId}`)
      .delete();
  }

  /**
   * Move email to folder
   */
  async moveEmail(messageId: string, destinationFolderId: string): Promise<void> {
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
