/**
 * Microsoft Graph API Authentication
 */

import { ConfidentialClientApplication } from '@azure/msal-node';
import { Client } from '@microsoft/microsoft-graph-client';
import { GraphConfig } from './types.js';

export class GraphAuthProvider {
  private msalClient: ConfidentialClientApplication;
  private config: GraphConfig;
  private accessToken: string | null = null;
  private tokenExpiry: Date | null = null;

  constructor(config: GraphConfig) {
    this.config = config;

    this.msalClient = new ConfidentialClientApplication({
      auth: {
        clientId: config.clientId,
        authority: `https://login.microsoftonline.com/${config.tenantId}`,
        clientSecret: config.clientSecret,
      },
    });
  }

  /**
   * Get an access token for Microsoft Graph API
   */
  async getAccessToken(): Promise<string> {
    // Return cached token if still valid
    if (this.accessToken && this.tokenExpiry && this.tokenExpiry > new Date()) {
      return this.accessToken;
    }

    try {
      const authResult = await this.msalClient.acquireTokenByClientCredential({
        scopes: ['https://graph.microsoft.com/.default'],
      });

      if (!authResult || !authResult.accessToken) {
        throw new Error('Failed to acquire access token');
      }

      this.accessToken = authResult.accessToken;
      // Set expiry to 5 minutes before actual expiry to be safe
      this.tokenExpiry = new Date(Date.now() + (authResult.expiresOn!.getTime() - Date.now() - 5 * 60 * 1000));

      return this.accessToken;
    } catch (error) {
      throw new Error(`Authentication failed: ${error instanceof Error ? error.message : String(error)}`);
    }
  }

  /**
   * Create an authenticated Graph API client
   */
  async getGraphClient(): Promise<Client> {
    const accessToken = await this.getAccessToken();

    return Client.init({
      authProvider: (done) => {
        done(null, accessToken);
      },
    });
  }

  /**
   * Get the user ID to use for operations
   * Returns the configured userId or userPrincipalName, or 'me' for current user
   */
  getUserId(): string {
    if (this.config.userId) {
      return this.config.userId;
    }
    if (this.config.userPrincipalName) {
      return this.config.userPrincipalName;
    }
    return 'me';
  }
}
