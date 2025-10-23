/**
 * OneDrive tools for MCP server
 */

import { Client } from '@microsoft/microsoft-graph-client';
import { DriveItem } from '../types.js';

export class OneDriveTools {
  constructor(private graphClient: Client, private userId: string) {}

  /**
   * List items in OneDrive root or a specific folder
   */
  async listItems(folderId?: string): Promise<DriveItem[]> {
    const endpoint = folderId
      ? `/users/${this.userId}/drive/items/${folderId}/children`
      : `/users/${this.userId}/drive/root/children`;

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
    const item = await this.graphClient
      .api(`/users/${this.userId}/drive/items/${itemId}`)
      .get();

    return item;
  }

  /**
   * Get item by path
   */
  async getItemByPath(path: string): Promise<DriveItem> {
    const item = await this.graphClient
      .api(`/users/${this.userId}/drive/root:/${path}`)
      .get();

    return item;
  }

  /**
   * Download file content
   */
  async downloadFile(itemId: string): Promise<ArrayBuffer> {
    const content = await this.graphClient
      .api(`/users/${this.userId}/drive/items/${itemId}/content`)
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
    const endpoint = parentFolderId
      ? `/users/${this.userId}/drive/items/${parentFolderId}:/${fileName}:/content`
      : `/users/${this.userId}/drive/root:/${fileName}:/content`;

    const uploadedFile = await this.graphClient
      .api(endpoint)
      .put(content);

    return uploadedFile;
  }

  /**
   * Create a folder
   */
  async createFolder(folderName: string, parentFolderId?: string): Promise<DriveItem> {
    const endpoint = parentFolderId
      ? `/users/${this.userId}/drive/items/${parentFolderId}/children`
      : `/users/${this.userId}/drive/root/children`;

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
    await this.graphClient
      .api(`/users/${this.userId}/drive/items/${itemId}`)
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
      .api(`/users/${this.userId}/drive/items/${itemId}/copy`)
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
      .api(`/users/${this.userId}/drive/items/${itemId}`)
      .patch(requestBody);

    return movedItem;
  }

  /**
   * Search for items
   */
  async searchItems(query: string): Promise<DriveItem[]> {
    const result = await this.graphClient
      .api(`/users/${this.userId}/drive/root/search(q='${query}')`)
      .select(['id', 'name', 'size', 'webUrl', 'folder', 'file'])
      .get();

    return result.value;
  }

  /**
   * Share an item (create sharing link)
   */
  async shareItem(itemId: string, type: 'view' | 'edit' = 'view', scope: 'anonymous' | 'organization' = 'organization'): Promise<string> {
    const result = await this.graphClient
      .api(`/users/${this.userId}/drive/items/${itemId}/createLink`)
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
      .api(`/users/${this.userId}/drive/recent`)
      .top(top)
      .get();

    return result.value;
  }

  /**
   * Get shared with me files
   */
  async getSharedWithMe(): Promise<DriveItem[]> {
    const result = await this.graphClient
      .api(`/users/${this.userId}/drive/sharedWithMe`)
      .get();

    return result.value;
  }
}
