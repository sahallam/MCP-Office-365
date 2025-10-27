/**
 * Word tools for MCP server
 */

import { Client } from '@microsoft/microsoft-graph-client';
import { WordDocument } from '../types.js';
import { sanitizeSearchQuery, validateResourceId } from '../security.js';

export class WordTools {
  constructor(private graphClient: Client, private userId: string) {}

  /**
   * Get the correct user path for API endpoints
   * Returns '/me' if userId is 'me', otherwise '/users/{userId}'
   */
  private getUserPath(): string {
    return this.userId === 'me' ? '/me' : `/users/${this.userId}`;
  }

  /**
   * Get Word document metadata
   */
  async getDocument(itemId: string): Promise<WordDocument> {
    validateResourceId(itemId, 'document');

    const document = await this.graphClient
      .api(`${this.getUserPath()}/drive/items/${itemId}`)
      .get();

    return {
      id: document.id,
      name: document.name,
      webUrl: document.webUrl,
    };
  }

  /**
   * Get document content as text
   * Note: This downloads the file content. For large files, consider streaming.
   */
  async getDocumentContent(itemId: string): Promise<ArrayBuffer> {
    validateResourceId(itemId, 'document');

    const content = await this.graphClient
      .api(`${this.getUserPath()}/drive/items/${itemId}/content`)
      .get();

    return content;
  }

  /**
   * Convert Word document to PDF
   */
  async convertToPdf(itemId: string): Promise<ArrayBuffer> {
    validateResourceId(itemId, 'document');

    const pdfContent = await this.graphClient
      .api(`${this.getUserPath()}/drive/items/${itemId}/content?format=pdf`)
      .get();

    return pdfContent;
  }

  /**
   * Search for Word documents
   */
  async searchDocuments(query: string): Promise<WordDocument[]> {
    // Sanitize search query to prevent OData injection
    const sanitizedQuery = sanitizeSearchQuery(query);

    const result = await this.graphClient
      .api(`${this.getUserPath()}/drive/root/search(q='${sanitizedQuery}')`)
      .filter("file/mimeType eq 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'")
      .select(['id', 'name', 'webUrl'])
      .get();

    return result.value;
  }

  /**
   * Create a new Word document
   */
  async createDocument(fileName: string, content?: Buffer | string, parentFolderId?: string): Promise<WordDocument> {
    const endpoint = parentFolderId
      ? `${this.getUserPath()}/drive/items/${parentFolderId}:/${fileName}:/content`
      : `${this.getUserPath()}/drive/root:/${fileName}:/content`;

    // Create an empty Word document or with provided content
    const documentContent = content || Buffer.from('');

    const createdFile = await this.graphClient
      .api(endpoint)
      .header('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document')
      .put(documentContent);

    return {
      id: createdFile.id,
      name: createdFile.name,
      webUrl: createdFile.webUrl,
    };
  }

  /**
   * Copy a Word document
   */
  async copyDocument(itemId: string, parentFolderId: string, newName?: string): Promise<void> {
    validateResourceId(itemId, 'document');
    validateResourceId(parentFolderId, 'folder');

    const requestBody: any = {
      parentReference: {
        id: parentFolderId,
      },
    };

    if (newName) {
      requestBody.name = newName;
    }

    await this.graphClient
      .api(`${this.getUserPath()}/drive/items/${itemId}/copy`)
      .post(requestBody);
  }

  /**
   * Get document permissions/sharing info
   */
  async getPermissions(itemId: string): Promise<any[]> {
    validateResourceId(itemId, 'document');

    const result = await this.graphClient
      .api(`${this.getUserPath()}/drive/items/${itemId}/permissions`)
      .get();

    return result.value;
  }

  /**
   * Share a document
   */
  async shareDocument(
    itemId: string,
    type: 'view' | 'edit' = 'view',
    scope: 'anonymous' | 'organization' = 'organization'
  ): Promise<string> {
    validateResourceId(itemId, 'document');

    const result = await this.graphClient
      .api(`${this.getUserPath()}/drive/items/${itemId}/createLink`)
      .post({
        type,
        scope,
      });

    return result.link.webUrl;
  }

  /**
   * Get document versions
   */
  async getVersions(itemId: string): Promise<any[]> {
    const result = await this.graphClient
      .api(`${this.getUserPath()}/drive/items/${itemId}/versions`)
      .get();

    return result.value;
  }

  /**
   * Restore a previous version
   */
  async restoreVersion(itemId: string, versionId: string): Promise<void> {
    await this.graphClient
      .api(`${this.getUserPath()}/drive/items/${itemId}/versions/${versionId}/restoreVersion`)
      .post({});
  }

  /**
   * Get thumbnails for the document
   */
  async getThumbnails(itemId: string): Promise<any[]> {
    const result = await this.graphClient
      .api(`${this.getUserPath()}/drive/items/${itemId}/thumbnails`)
      .get();

    return result.value;
  }

  /**
   * Check out a document (for co-authoring scenarios)
   */
  async checkOut(itemId: string): Promise<void> {
    await this.graphClient
      .api(`${this.getUserPath()}/drive/items/${itemId}/checkout`)
      .post({});
  }

  /**
   * Check in a document
   */
  async checkIn(itemId: string, comment?: string): Promise<void> {
    await this.graphClient
      .api(`${this.getUserPath()}/drive/items/${itemId}/checkin`)
      .post({
        comment: comment || '',
      });
  }

  /**
   * Get recent Word documents
   */
  async getRecentDocuments(top: number = 10): Promise<WordDocument[]> {
    const result = await this.graphClient
      .api(`${this.getUserPath()}/drive/recent`)
      .filter("file/mimeType eq 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'")
      .top(top)
      .select(['id', 'name', 'webUrl'])
      .get();

    return result.value;
  }
}
