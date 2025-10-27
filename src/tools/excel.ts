/**
 * Excel tools for MCP server
 */

import { Client } from '@microsoft/microsoft-graph-client';
import { ExcelWorkbook, ExcelWorksheet } from '../types.js';
import { validateResourceId, validateExcelAddress } from '../security.js';

export class ExcelTools {
  constructor(private graphClient: Client, private userId: string) {}

  /**
   * Get the correct user path for API endpoints
   * Returns '/me' if userId is 'me', otherwise '/users/{userId}'
   */
  private getUserPath(): string {
    return this.userId === 'me' ? '/me' : `/users/${this.userId}`;
  }

  /**
   * Get workbook metadata
   */
  async getWorkbook(itemId: string): Promise<ExcelWorkbook> {
    validateResourceId(itemId, 'workbook');

    const workbook = await this.graphClient
      .api(`${this.getUserPath()}/drive/items/${itemId}/workbook`)
      .get();

    return workbook;
  }

  /**
   * List worksheets in a workbook
   */
  async listWorksheets(itemId: string): Promise<ExcelWorksheet[]> {
    validateResourceId(itemId, 'workbook');

    const result = await this.graphClient
      .api(`${this.getUserPath()}/drive/items/${itemId}/workbook/worksheets`)
      .get();

    return result.value;
  }

  /**
   * Get a specific worksheet
   */
  async getWorksheet(itemId: string, worksheetId: string): Promise<ExcelWorksheet> {
    const worksheet = await this.graphClient
      .api(`${this.getUserPath()}/drive/items/${itemId}/workbook/worksheets/${worksheetId}`)
      .get();

    return worksheet;
  }

  /**
   * Create a new worksheet
   */
  async createWorksheet(itemId: string, name: string): Promise<ExcelWorksheet> {
    const worksheet = await this.graphClient
      .api(`${this.getUserPath()}/drive/items/${itemId}/workbook/worksheets/add`)
      .post({
        name,
      });

    return worksheet;
  }

  /**
   * Get range values from a worksheet
   */
  async getRange(itemId: string, worksheetId: string, address: string): Promise<any> {
    validateResourceId(itemId, 'workbook');
    validateResourceId(worksheetId, 'worksheet');
    validateExcelAddress(address);

    const range = await this.graphClient
      .api(`${this.getUserPath()}/drive/items/${itemId}/workbook/worksheets/${worksheetId}/range(address='${address}')`)
      .get();

    return range;
  }

  /**
   * Update range values in a worksheet
   */
  async updateRange(itemId: string, worksheetId: string, address: string, values: any[][]): Promise<any> {
    validateResourceId(itemId, 'workbook');
    validateResourceId(worksheetId, 'worksheet');
    validateExcelAddress(address);

    const range = await this.graphClient
      .api(`${this.getUserPath()}/drive/items/${itemId}/workbook/worksheets/${worksheetId}/range(address='${address}')`)
      .patch({
        values,
      });

    return range;
  }

  /**
   * Get used range (non-empty cells) from a worksheet
   */
  async getUsedRange(itemId: string, worksheetId: string): Promise<any> {
    validateResourceId(itemId, 'workbook');
    validateResourceId(worksheetId, 'worksheet');

    const range = await this.graphClient
      .api(`${this.getUserPath()}/drive/items/${itemId}/workbook/worksheets/${worksheetId}/usedRange`)
      .get();

    return range;
  }

  /**
   * Add a table to a worksheet
   */
  async createTable(itemId: string, worksheetId: string, address: string, hasHeaders: boolean = true): Promise<any> {
    const table = await this.graphClient
      .api(`${this.getUserPath()}/drive/items/${itemId}/workbook/worksheets/${worksheetId}/tables/add`)
      .post({
        address,
        hasHeaders,
      });

    return table;
  }

  /**
   * List tables in a worksheet
   */
  async listTables(itemId: string, worksheetId: string): Promise<any[]> {
    const result = await this.graphClient
      .api(`${this.getUserPath()}/drive/items/${itemId}/workbook/worksheets/${worksheetId}/tables`)
      .get();

    return result.value;
  }

  /**
   * Get table data
   */
  async getTableData(itemId: string, tableId: string): Promise<any> {
    const result = await this.graphClient
      .api(`${this.getUserPath()}/drive/items/${itemId}/workbook/tables/${tableId}/rows`)
      .get();

    return result.value;
  }

  /**
   * Add rows to a table
   */
  async addTableRows(itemId: string, tableId: string, values: any[][]): Promise<any> {
    const result = await this.graphClient
      .api(`${this.getUserPath()}/drive/items/${itemId}/workbook/tables/${tableId}/rows/add`)
      .post({
        values,
      });

    return result;
  }

  /**
   * Get named items (named ranges, etc.)
   */
  async listNamedItems(itemId: string): Promise<any[]> {
    const result = await this.graphClient
      .api(`${this.getUserPath()}/drive/items/${itemId}/workbook/names`)
      .get();

    return result.value;
  }

  /**
   * Get a named range value
   */
  async getNamedRange(itemId: string, name: string): Promise<any> {
    const range = await this.graphClient
      .api(`${this.getUserPath()}/drive/items/${itemId}/workbook/names/${name}/range`)
      .get();

    return range;
  }

  /**
   * Create a chart
   */
  async createChart(
    itemId: string,
    worksheetId: string,
    type: string,
    sourceData: string,
    seriesBy: 'Auto' | 'Columns' | 'Rows' = 'Auto'
  ): Promise<any> {
    const chart = await this.graphClient
      .api(`${this.getUserPath()}/drive/items/${itemId}/workbook/worksheets/${worksheetId}/charts/add`)
      .post({
        type,
        sourceData,
        seriesBy,
      });

    return chart;
  }

  /**
   * List charts in a worksheet
   */
  async listCharts(itemId: string, worksheetId: string): Promise<any[]> {
    const result = await this.graphClient
      .api(`${this.getUserPath()}/drive/items/${itemId}/workbook/worksheets/${worksheetId}/charts`)
      .get();

    return result.value;
  }

  /**
   * Refresh all data connections (for workbooks with external data)
   */
  async refreshData(itemId: string): Promise<void> {
    await this.graphClient
      .api(`${this.getUserPath()}/drive/items/${itemId}/workbook/refreshAllDataConnections`)
      .post({});
  }

  /**
   * Create a session for batch operations
   */
  async createSession(itemId: string, persistChanges: boolean = true): Promise<string> {
    const session = await this.graphClient
      .api(`${this.getUserPath()}/drive/items/${itemId}/workbook/createSession`)
      .post({
        persistChanges,
      });

    return session.id;
  }

  /**
   * Close a session
   */
  async closeSession(itemId: string, sessionId: string): Promise<void> {
    await this.graphClient
      .api(`${this.getUserPath()}/drive/items/${itemId}/workbook/closeSession`)
      .header('workbook-session-id', sessionId)
      .post({});
  }
}
