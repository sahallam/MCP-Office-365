/**
 * SharePoint tools for MCP server
 */

import { Client } from '@microsoft/microsoft-graph-client';
import { SharePointSite, DriveItem } from '../types.js';
import { sanitizeSearchQuery, validateResourceId, validateFileSize, SIZE_LIMITS } from '../security.js';

export class SharePointTools {
  constructor(private graphClient: Client) {}

  /**
   * Search for SharePoint sites
   */
  async searchSites(query: string): Promise<SharePointSite[]> {
    // Sanitize search query to prevent injection
    const sanitizedQuery = sanitizeSearchQuery(query);

    const result = await this.graphClient
      .api(`/sites?search=${encodeURIComponent(sanitizedQuery)}`)
      .get();

    return result.value;
  }

  /**
   * Get a site by ID
   */
  async getSite(siteId: string): Promise<SharePointSite> {
    validateResourceId(siteId, 'site');

    const site = await this.graphClient
      .api(`/sites/${siteId}`)
      .get();

    return site;
  }

  /**
   * Get site by hostname and path
   */
  async getSiteByPath(hostname: string, serverRelativePath: string): Promise<SharePointSite> {
    // Validate hostname to prevent SSRF
    this.validateSharePointHostname(hostname);
    validateResourceId(serverRelativePath, 'path');

    const site = await this.graphClient
      .api(`/sites/${hostname}:${serverRelativePath}`)
      .get();

    return site;
  }

  /**
   * Validate SharePoint hostname to prevent SSRF
   */
  private validateSharePointHostname(hostname: string): void {
    if (!hostname || typeof hostname !== 'string') {
      throw new Error('Invalid hostname: must be a non-empty string');
    }

    // Only allow SharePoint domains
    const allowedDomains = ['.sharepoint.com', '.sharepoint-df.com'];
    if (!allowedDomains.some(domain => hostname.endsWith(domain))) {
      throw new Error('Invalid SharePoint hostname: must be a sharepoint.com domain');
    }

    // Prevent internal network access
    if (hostname.match(/^(localhost|127\.|10\.|172\.(1[6-9]|2[0-9]|3[01])\.|192\.168\.)/i)) {
      throw new Error('Access to internal network not allowed');
    }

    // Prevent other suspicious patterns
    if (hostname.includes('@') || hostname.includes(':') || hostname.includes('/')) {
      throw new Error('Invalid hostname format');
    }
  }

  /**
   * List document libraries (drives) in a site
   */
  async listDocumentLibraries(siteId: string): Promise<any[]> {
    validateResourceId(siteId, 'site');

    const result = await this.graphClient
      .api(`/sites/${siteId}/drives`)
      .get();

    return result.value;
  }

  /**
   * List items in a document library
   */
  async listLibraryItems(siteId: string, driveId: string, folderId?: string): Promise<DriveItem[]> {
    validateResourceId(siteId, 'site');
    validateResourceId(driveId, 'drive');
    if (folderId) {
      validateResourceId(folderId, 'folder');
    }

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
    validateResourceId(siteId, 'site');
    validateResourceId(driveId, 'drive');
    if (parentFolderId) {
      validateResourceId(parentFolderId, 'folder');
    }

    // Convert base64 string to Buffer if needed
    let uploadContent: Buffer;
    if (typeof content === 'string') {
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
      ? `/sites/${siteId}/drives/${driveId}/items/${parentFolderId}:/${fileName}:/content`
      : `/sites/${siteId}/drives/${driveId}/root:/${fileName}:/content`;

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
    validateResourceId(siteId, 'site');
    validateResourceId(driveId, 'drive');
    validateResourceId(itemId, 'item');

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
    validateResourceId(siteId, 'site');
    validateResourceId(listId, 'list');

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
    validateResourceId(siteId, 'site');
    validateResourceId(listId, 'list');

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
    validateResourceId(siteId, 'site');
    validateResourceId(listId, 'list');
    validateResourceId(itemId, 'item');

    const item = await this.graphClient
      .api(`/sites/${siteId}/lists/${listId}/items/${itemId}/fields`)
      .patch(fields);

    return item;
  }

  /**
   * Delete list item
   */
  async deleteListItem(siteId: string, listId: string, itemId: string): Promise<void> {
    validateResourceId(siteId, 'site');
    validateResourceId(listId, 'list');
    validateResourceId(itemId, 'item');

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
    validateResourceId(siteId, 'site');

    const result = await this.graphClient
      .api(`/sites/${siteId}/sites`)
      .get();

    return result.value;
  }
}
