/**
 * SharePoint tools for MCP server
 */

import { Client } from '@microsoft/microsoft-graph-client';
import { SharePointSite, DriveItem } from '../types.js';

export class SharePointTools {
  constructor(private graphClient: Client) {}

  /**
   * Search for SharePoint sites
   */
  async searchSites(query: string): Promise<SharePointSite[]> {
    const result = await this.graphClient
      .api(`/sites?search=${query}`)
      .get();

    return result.value;
  }

  /**
   * Get a site by ID
   */
  async getSite(siteId: string): Promise<SharePointSite> {
    const site = await this.graphClient
      .api(`/sites/${siteId}`)
      .get();

    return site;
  }

  /**
   * Get site by hostname and path
   */
  async getSiteByPath(hostname: string, serverRelativePath: string): Promise<SharePointSite> {
    const site = await this.graphClient
      .api(`/sites/${hostname}:${serverRelativePath}`)
      .get();

    return site;
  }

  /**
   * List document libraries (drives) in a site
   */
  async listDocumentLibraries(siteId: string): Promise<any[]> {
    const result = await this.graphClient
      .api(`/sites/${siteId}/drives`)
      .get();

    return result.value;
  }

  /**
   * List items in a document library
   */
  async listLibraryItems(siteId: string, driveId: string, folderId?: string): Promise<DriveItem[]> {
    const endpoint = folderId
      ? `/sites/${siteId}/drives/${driveId}/items/${folderId}/children`
      : `/sites/${siteId}/drives/${driveId}/root/children`;

    const result = await this.graphClient
      .api(endpoint)
      .select(['id', 'name', 'size', 'createdDateTime', 'lastModifiedDateTime', 'webUrl', 'folder', 'file'])
      .get();

    return result.value;
  }

  /**
   * Upload file to SharePoint library
   */
  async uploadFile(
    siteId: string,
    driveId: string,
    fileName: string,
    content: Buffer | string,
    parentFolderId?: string
  ): Promise<DriveItem> {
    const endpoint = parentFolderId
      ? `/sites/${siteId}/drives/${driveId}/items/${parentFolderId}:/${fileName}:/content`
      : `/sites/${siteId}/drives/${driveId}/root:/${fileName}:/content`;

    // Convert base64 string to Buffer if needed
    let uploadContent: Buffer;
    if (typeof content === 'string') {
      // Assume base64 encoding for string content
      uploadContent = Buffer.from(content, 'base64');
    } else {
      uploadContent = content;
    }

    const uploadedFile = await this.graphClient
      .api(endpoint)
      .header('Content-Type', 'application/octet-stream')
      .put(uploadContent);

    return uploadedFile;
  }

  /**
   * Download file from SharePoint
   */
  async downloadFile(siteId: string, driveId: string, itemId: string): Promise<ArrayBuffer> {
    const content = await this.graphClient
      .api(`/sites/${siteId}/drives/${driveId}/items/${itemId}/content`)
      .get();

    return content;
  }

  /**
   * List SharePoint lists
   */
  async listLists(siteId: string): Promise<any[]> {
    const result = await this.graphClient
      .api(`/sites/${siteId}/lists`)
      .get();

    return result.value;
  }

  /**
   * Get list items
   */
  async getListItems(siteId: string, listId: string): Promise<any[]> {
    const result = await this.graphClient
      .api(`/sites/${siteId}/lists/${listId}/items`)
      .expand('fields')
      .get();

    return result.value;
  }

  /**
   * Create list item
   */
  async createListItem(siteId: string, listId: string, fields: Record<string, any>): Promise<any> {
    const item = await this.graphClient
      .api(`/sites/${siteId}/lists/${listId}/items`)
      .post({
        fields,
      });

    return item;
  }

  /**
   * Update list item
   */
  async updateListItem(siteId: string, listId: string, itemId: string, fields: Record<string, any>): Promise<any> {
    const item = await this.graphClient
      .api(`/sites/${siteId}/lists/${listId}/items/${itemId}/fields`)
      .patch(fields);

    return item;
  }

  /**
   * Delete list item
   */
  async deleteListItem(siteId: string, listId: string, itemId: string): Promise<void> {
    await this.graphClient
      .api(`/sites/${siteId}/lists/${listId}/items/${itemId}`)
      .delete();
  }

  /**
   * Get root site
   */
  async getRootSite(): Promise<SharePointSite> {
    const site = await this.graphClient
      .api('/sites/root')
      .get();

    return site;
  }

  /**
   * List subsites
   */
  async listSubsites(siteId: string): Promise<SharePointSite[]> {
    const result = await this.graphClient
      .api(`/sites/${siteId}/sites`)
      .get();

    return result.value;
  }
}
