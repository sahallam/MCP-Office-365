/**
 * OneDrive tools for MCP server
 */

import { Client } from '@microsoft/microsoft-graph-client';
import { DriveItem } from '../types.js';
import { sanitizeSearchQuery, validateResourceId, validateFileSize, SIZE_LIMITS } from '../security.js';

export class OneDriveTools {
  constructor(private graphClient: Client, private userId: string) {}

  /**
   * Get the correct user path for API endpoints
   * Returns '/me' if userId is 'me', otherwise '/users/{userId}'
   */
  private getUserPath(): string {
    return this.userId === 'me' ? '/me' : `/users/${this.userId}`;
  }

  /**
   * List items in OneDrive root or a specific folder
   */
  async listItems(folderId?: string): Promise<DriveItem[]> {
    const endpoint = folderId
      ? `${this.getUserPath()}/drive/items/${folderId}/children`
      : `${this.getUserPath()}/drive/root/children`;

    const result = await this.graphClient
      .api(endpoint)
      .select(['id', 'name', 'size', 'createdDateTime', 'lastModifiedDateTime', 'webUrl', 'folder', 'file'])
      .get();

    return result.value;
  }

  /**
   * Get item by ID
   */
  async getItem(itemId: string): Promise<DriveItem> {
    validateResourceId(itemId, 'item');

    const item = await this.graphClient
      .api(`${this.getUserPath()}/drive/items/${itemId}`)
      .get();

    return item;
  }

  /**
   * Get item by path
   */
  async getItemByPath(path: string): Promise<DriveItem> {
    const item = await this.graphClient
      .api(`${this.getUserPath()}/drive/root:/${path}`)
      .get();

    return item;
  }

  /**
   * Download file content
   */
  async downloadFile(itemId: string): Promise<ArrayBuffer> {
    validateResourceId(itemId, 'item');

    const content = await this.graphClient
      .api(`${this.getUserPath()}/drive/items/${itemId}/content`)
      .get();

    return content;
  }

  /**
   * Upload file to OneDrive
   */
  async uploadFile(
    fileName: string,
    content: Buffer | string,
    parentFolderId?: string
  ): Promise<DriveItem> {
    // Validate parent folder ID if provided
    if (parentFolderId) {
      validateResourceId(parentFolderId, 'folder');
    }

    // Convert base64 string to Buffer if needed
    let uploadContent: Buffer;
    if (typeof content === 'string') {
      // Assume base64 encoding for string content
      try {
        uploadContent = Buffer.from(content, 'base64');

        // Validate that it's actually valid base64
        if (uploadContent.toString('base64') !== content.replace(/\s/g, '')) {
          throw new Error('Invalid base64 encoding');
        }
      } catch (error) {
        throw new Error('Invalid base64 string provided for file content');
      }
    } else {
      uploadContent = content;
    }

    // Validate file size
    validateFileSize(uploadContent.length, SIZE_LIMITS.MAX_FILE_SIZE, 'File');

    const endpoint = parentFolderId
      ? `${this.getUserPath()}/drive/items/${parentFolderId}:/${fileName}:/content`
      : `${this.getUserPath()}/drive/root:/${fileName}:/content`;

    const uploadedFile = await this.graphClient
      .api(endpoint)
      .header('Content-Type', 'application/octet-stream')
      .put(uploadContent);

    return uploadedFile;
  }

  /**
   * Create a folder
   */
  async createFolder(folderName: string, parentFolderId?: string): Promise<DriveItem> {
    const endpoint = parentFolderId
      ? `${this.getUserPath()}/drive/items/${parentFolderId}/children`
      : `${this.getUserPath()}/drive/root/children`;

    const folder = await this.graphClient
      .api(endpoint)
      .post({
        name: folderName,
        folder: {},
        '@microsoft.graph.conflictBehavior': 'rename',
      });

    return folder;
  }

  /**
   * Delete an item
   */
  async deleteItem(itemId: string): Promise<void> {
    validateResourceId(itemId, 'item');

    await this.graphClient
      .api(`${this.getUserPath()}/drive/items/${itemId}`)
      .delete();
  }

  /**
   * Copy an item
   */
  async copyItem(itemId: string, parentFolderId: string, newName?: string): Promise<void> {
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
   * Move an item
   */
  async moveItem(itemId: string, parentFolderId: string, newName?: string): Promise<DriveItem> {
    const requestBody: any = {
      parentReference: {
        id: parentFolderId,
      },
    };

    if (newName) {
      requestBody.name = newName;
    }

    const movedItem = await this.graphClient
      .api(`${this.getUserPath()}/drive/items/${itemId}`)
      .patch(requestBody);

    return movedItem;
  }

  /**
   * Search for items
   */
  async searchItems(query: string): Promise<DriveItem[]> {
    // Sanitize search query to prevent OData injection
    const sanitizedQuery = sanitizeSearchQuery(query);

    const result = await this.graphClient
      .api(`${this.getUserPath()}/drive/root/search(q='${sanitizedQuery}')`)
      .select(['id', 'name', 'size', 'webUrl', 'folder', 'file'])
      .get();

    return result.value;
  }

  /**
   * Share an item (create sharing link)
   */
  async shareItem(itemId: string, type: 'view' | 'edit' = 'view', scope: 'anonymous' | 'organization' = 'organization'): Promise<string> {
    validateResourceId(itemId, 'item');

    const result = await this.graphClient
      .api(`${this.getUserPath()}/drive/items/${itemId}/createLink`)
      .post({
        type,
        scope,
      });

    return result.link.webUrl;
  }

  /**
   * Get recent files
   */
  async getRecentFiles(top: number = 10): Promise<DriveItem[]> {
    const result = await this.graphClient
      .api(`${this.getUserPath()}/drive/recent`)
      .top(top)
      .get();

    return result.value;
  }

  /**
   * Get shared with me files
   */
  async getSharedWithMe(): Promise<DriveItem[]> {
    const result = await this.graphClient
      .api(`${this.getUserPath()}/drive/sharedWithMe`)
      .get();

    return result.value;
  }
}
