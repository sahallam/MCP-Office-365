/**
 * OneNote tools for MCP server
 */

import { Client } from '@microsoft/microsoft-graph-client';
import { OneNoteNotebook, OneNoteSection, OneNotePage } from '../types.js';
import { sanitizeSearchQuery, validateResourceId, escapeHtml, sanitizeHtmlContent, validateContentLength, SIZE_LIMITS } from '../security.js';

export class OneNoteTools {
  constructor(private graphClient: Client, private userId: string) {}

  /**
   * List all notebooks
   */
  async listNotebooks(options: {
    top?: number;
    orderBy?: string;
    select?: string[];
  } = {}): Promise<OneNoteNotebook[]> {
    const {
      top = 50,
      orderBy = 'lastModifiedDateTime DESC',
      select = ['id', 'displayName', 'createdDateTime', 'lastModifiedDateTime', 'isDefault', 'isShared']
    } = options;

    let query = this.graphClient
      .api(`/users/${this.userId}/onenote/notebooks`)
      .top(top)
      .orderby(orderBy)
      .select(select);

    const result = await query.get();
    return result.value;
  }

  /**
   * Get a specific notebook by ID
   */
  async getNotebook(notebookId: string): Promise<OneNoteNotebook> {
    validateResourceId(notebookId, 'notebook');

    const notebook = await this.graphClient
      .api(`/users/${this.userId}/onenote/notebooks/${notebookId}`)
      .get();

    return notebook;
  }

  /**
   * Create a new notebook
   */
  async createNotebook(displayName: string): Promise<OneNoteNotebook> {
    const notebook = await this.graphClient
      .api(`/users/${this.userId}/onenote/notebooks`)
      .post({
        displayName
      });

    return notebook;
  }

  /**
   * List sections in a notebook
   */
  async listSections(options: {
    notebookId?: string;
    top?: number;
    orderBy?: string;
  } = {}): Promise<OneNoteSection[]> {
    const {
      notebookId,
      top = 50,
      orderBy = 'lastModifiedDateTime DESC'
    } = options;

    let apiPath = notebookId
      ? `/users/${this.userId}/onenote/notebooks/${notebookId}/sections`
      : `/users/${this.userId}/onenote/sections`;

    let query = this.graphClient
      .api(apiPath)
      .top(top)
      .orderby(orderBy);

    const result = await query.get();
    return result.value;
  }

  /**
   * Get a specific section by ID
   */
  async getSection(sectionId: string): Promise<OneNoteSection> {
    validateResourceId(sectionId, 'section');

    const section = await this.graphClient
      .api(`/users/${this.userId}/onenote/sections/${sectionId}`)
      .get();

    return section;
  }

  /**
   * Create a new section in a notebook
   */
  async createSection(notebookId: string, displayName: string): Promise<OneNoteSection> {
    validateResourceId(notebookId, 'notebook');

    const section = await this.graphClient
      .api(`/users/${this.userId}/onenote/notebooks/${notebookId}/sections`)
      .post({
        displayName
      });

    return section;
  }

  /**
   * List pages in a section or notebook
   */
  async listPages(options: {
    sectionId?: string;
    notebookId?: string;
    top?: number;
    orderBy?: string;
    search?: string;
  } = {}): Promise<OneNotePage[]> {
    const {
      sectionId,
      notebookId,
      top = 50,
      orderBy = 'lastModifiedDateTime DESC',
      search
    } = options;

    let apiPath: string;
    if (sectionId) {
      apiPath = `/users/${this.userId}/onenote/sections/${sectionId}/pages`;
    } else if (notebookId) {
      apiPath = `/users/${this.userId}/onenote/notebooks/${notebookId}/pages`;
    } else {
      apiPath = `/users/${this.userId}/onenote/pages`;
    }

    let query = this.graphClient
      .api(apiPath)
      .top(top)
      .orderby(orderBy);

    if (search) {
      // Sanitize search query to prevent injection
      const sanitizedSearch = sanitizeSearchQuery(search);
      query = query.search(`"${sanitizedSearch}"`);
    }

    const result = await query.get();
    return result.value;
  }

  /**
   * Get a specific page by ID
   */
  async getPage(pageId: string): Promise<OneNotePage> {
    validateResourceId(pageId, 'page');

    const page = await this.graphClient
      .api(`/users/${this.userId}/onenote/pages/${pageId}`)
      .get();

    return page;
  }

  /**
   * Get page content (HTML)
   */
  async getPageContent(pageId: string): Promise<string> {
    validateResourceId(pageId, 'page');

    const content = await this.graphClient
      .api(`/users/${this.userId}/onenote/pages/${pageId}/content`)
      .get();

    return content;
  }

  /**
   * Create a new page in a section
   */
  async createPage(sectionId: string, title: string, content: string): Promise<OneNotePage> {
    validateResourceId(sectionId, 'section');

    // Validate content size
    validateContentLength(content, SIZE_LIMITS.MAX_ONENOTE_PAGE_SIZE, 'OneNote page content');

    // Sanitize title to prevent XSS (escape HTML entities)
    const sanitizedTitle = escapeHtml(title);

    // Sanitize content HTML while preserving safe formatting tags
    const sanitizedContent = sanitizeHtmlContent(content, [
      'p', 'br', 'strong', 'em', 'u', 'h1', 'h2', 'h3', 'h4', 'h5', 'h6',
      'ul', 'ol', 'li', 'blockquote', 'code', 'pre', 'div', 'span', 'a', 'img',
      'table', 'tr', 'td', 'th', 'tbody', 'thead'
    ]);

    // Content must be HTML
    const htmlContent = `
<!DOCTYPE html>
<html>
  <head>
    <title>${sanitizedTitle}</title>
  </head>
  <body>
    ${sanitizedContent}
  </body>
</html>`;

    const page = await this.graphClient
      .api(`/users/${this.userId}/onenote/sections/${sectionId}/pages`)
      .header('Content-Type', 'text/html')
      .post(htmlContent);

    return page;
  }

  /**
   * Update a page by appending content
   */
  async appendToPage(pageId: string, content: string): Promise<void> {
    validateResourceId(pageId, 'page');
    validateContentLength(content, SIZE_LIMITS.MAX_ONENOTE_PAGE_SIZE, 'OneNote page content');

    // Sanitize content HTML to prevent XSS
    const sanitizedContent = sanitizeHtmlContent(content, [
      'p', 'br', 'strong', 'em', 'u', 'h1', 'h2', 'h3', 'h4', 'h5', 'h6',
      'ul', 'ol', 'li', 'blockquote', 'code', 'pre', 'div', 'span', 'a', 'img',
      'table', 'tr', 'td', 'th', 'tbody', 'thead'
    ]);

    // Append content to the page
    const commands = [
      {
        target: 'body',
        action: 'append',
        content: sanitizedContent
      }
    ];

    await this.graphClient
      .api(`/users/${this.userId}/onenote/pages/${pageId}/content`)
      .patch(commands);
  }

  /**
   * Delete a page
   */
  async deletePage(pageId: string): Promise<void> {
    validateResourceId(pageId, 'page');

    await this.graphClient
      .api(`/users/${this.userId}/onenote/pages/${pageId}`)
      .delete();
  }

  /**
   * Delete a section
   */
  async deleteSection(sectionId: string): Promise<void> {
    validateResourceId(sectionId, 'section');

    await this.graphClient
      .api(`/users/${this.userId}/onenote/sections/${sectionId}`)
      .delete();
  }

  /**
   * Delete a notebook
   */
  async deleteNotebook(notebookId: string): Promise<void> {
    validateResourceId(notebookId, 'notebook');

    await this.graphClient
      .api(`/users/${this.userId}/onenote/notebooks/${notebookId}`)
      .delete();
  }

  /**
   * Search pages across all notebooks
   */
  async searchPages(searchQuery: string, top: number = 20): Promise<OneNotePage[]> {
    // Sanitize search query to prevent OData injection
    const sanitizedQuery = sanitizeSearchQuery(searchQuery);

    const result = await this.graphClient
      .api(`/users/${this.userId}/onenote/pages`)
      .search(`"${sanitizedQuery}"`)
      .top(top)
      .select(['id', 'title', 'createdDateTime', 'lastModifiedDateTime'])
      .get();

    return result.value;
  }

  /**
   * Copy page to a section
   */
  async copyPage(pageId: string, targetSectionId: string): Promise<any> {
    validateResourceId(pageId, 'page');
    validateResourceId(targetSectionId, 'section');

    const result = await this.graphClient
      .api(`/users/${this.userId}/onenote/pages/${pageId}/copyToSection`)
      .post({
        id: targetSectionId
      });

    return result;
  }
}
