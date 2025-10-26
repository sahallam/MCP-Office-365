/**
 * Microsoft Graph API Authentication
 */

import { ConfidentialClientApplication, PublicClientApplication, DeviceCodeRequest } from '@azure/msal-node';
import { Client } from '@microsoft/microsoft-graph-client';
import { GraphConfig } from './types.js';
import * as fs from 'fs';
import * as path from 'path';
import * as os from 'os';

interface TokenCache {
  accessToken: string;
  expiresOn: number;
  userId?: string;
  account?: any;
}

export class GraphAuthProvider {
  private msalClient: ConfidentialClientApplication | PublicClientApplication;
  private config: GraphConfig;
  private accessToken: string | null = null;
  private tokenExpiry: Date | null = null;
  private tokenCachePath: string;
  private userAccount: any = null;

  constructor(config: GraphConfig) {
    this.config = config;
    // Use home directory for token cache to avoid permission issues when running from system directories
    this.tokenCachePath = config.tokenCachePath || path.join(os.homedir(), '.office365-mcp-token-cache.json');

    // Determine authentication mode
    const authMode = config.authMode || 'app-only';

    if (authMode === 'delegated') {
      console.error(`[AUTH] Token cache location: ${this.tokenCachePath}`);

      // Use Public Client Application for delegated auth (device code flow)
      this.msalClient = new PublicClientApplication({
        auth: {
          clientId: config.clientId,
          authority: `https://login.microsoftonline.com/${config.tenantId}`,
        },
      });

      // Try to load cached tokens
      this.loadCachedTokens();
    } else {
      // Use Confidential Client Application for app-only auth
      if (!config.clientSecret) {
        throw new Error('CLIENT_SECRET is required for app-only authentication');
      }

      this.msalClient = new ConfidentialClientApplication({
        auth: {
          clientId: config.clientId,
          authority: `https://login.microsoftonline.com/${config.tenantId}`,
          clientSecret: config.clientSecret,
        },
      });
    }
  }

  /**
   * Load cached tokens from file
   */
  private loadCachedTokens(): void {
    try {
      if (fs.existsSync(this.tokenCachePath)) {
        const cache: TokenCache = JSON.parse(fs.readFileSync(this.tokenCachePath, 'utf-8'));
        if (cache.expiresOn > Date.now()) {
          this.accessToken = cache.accessToken;
          this.tokenExpiry = new Date(cache.expiresOn);
          this.userAccount = cache.account;
          console.error('[AUTH] Loaded cached access token');
        } else {
          console.error('[AUTH] Cached token expired, will re-authenticate');
        }
      }
    } catch (error) {
      console.error('[AUTH] Failed to load cached tokens:', error);
    }
  }

  /**
   * Save tokens to cache
   */
  private saveCachedTokens(accessToken: string, expiresOn: Date, account?: any): void {
    try {
      const cache: TokenCache = {
        accessToken,
        expiresOn: expiresOn.getTime(),
        userId: this.config.userPrincipalName || this.config.userId,
        account,
      };
      fs.writeFileSync(this.tokenCachePath, JSON.stringify(cache, null, 2));
      console.error(`[AUTH] Saved tokens to cache: ${this.tokenCachePath}`);
    } catch (error) {
      console.error(`[AUTH] Failed to save tokens to cache (${this.tokenCachePath}):`, error);
    }
  }

  /**
   * Acquire token using device code flow (for delegated auth)
   */
  private async acquireTokenByDeviceCode(): Promise<string> {
    const deviceCodeRequest: DeviceCodeRequest = {
      scopes: [
        'User.Read',
        'Mail.ReadWrite',
        'Mail.Send',
        'Calendars.ReadWrite',
        'Files.ReadWrite.All',
        'Notes.ReadWrite.All',
        'Team.ReadBasic.All',
        'Channel.ReadBasic.All',
        'ChannelMessage.Read.All',
        'Chat.Read',
        'Chat.ReadWrite',
      ],
      deviceCodeCallback: (response) => {
        console.error('\n=======================================================================');
        console.error('AUTHENTICATION REQUIRED');
        console.error('=======================================================================');
        console.error(`\nTo sign in, use a web browser to open the page:\n  ${response.verificationUri}`);
        console.error(`\nAnd enter the code: ${response.userCode}`);
        console.error('\n=======================================================================\n');
      },
    };

    try {
      const response = await (this.msalClient as PublicClientApplication).acquireTokenByDeviceCode(deviceCodeRequest);

      if (!response || !response.accessToken) {
        throw new Error('Failed to acquire access token via device code flow');
      }

      this.accessToken = response.accessToken;
      this.tokenExpiry = new Date(response.expiresOn!.getTime() - 5 * 60 * 1000); // 5 min buffer
      this.userAccount = response.account;

      // Save tokens to cache
      this.saveCachedTokens(this.accessToken, this.tokenExpiry, this.userAccount);

      return this.accessToken;
    } catch (error) {
      throw new Error(`Device code authentication failed: ${error instanceof Error ? error.message : String(error)}`);
    }
  }

  /**
   * Acquire token silently using MSAL cache
   */
  private async acquireTokenSilent(): Promise<string> {
    // Get all accounts from MSAL cache
    const accounts = await (this.msalClient as PublicClientApplication).getTokenCache().getAllAccounts();

    if (accounts.length === 0) {
      // No accounts in cache, need to authenticate
      return this.acquireTokenByDeviceCode();
    }

    // Use the first account (or the cached account if available)
    const account = this.userAccount || accounts[0];

    try {
      const response = await (this.msalClient as PublicClientApplication).acquireTokenSilent({
        account,
        scopes: ['https://graph.microsoft.com/.default'],
      });

      if (!response || !response.accessToken) {
        // Silent acquisition failed, need to re-authenticate
        return this.acquireTokenByDeviceCode();
      }

      this.accessToken = response.accessToken;
      this.tokenExpiry = new Date(response.expiresOn!.getTime() - 5 * 60 * 1000);
      this.userAccount = response.account;

      this.saveCachedTokens(this.accessToken, this.tokenExpiry, this.userAccount);

      return this.accessToken;
    } catch (error) {
      console.error('Silent token acquisition failed, re-authenticating:', error);
      return this.acquireTokenByDeviceCode();
    }
  }

  /**
   * Get an access token for Microsoft Graph API
   */
  async getAccessToken(): Promise<string> {
    // Return cached token if still valid
    if (this.accessToken && this.tokenExpiry && this.tokenExpiry > new Date()) {
      return this.accessToken;
    }

    const authMode = this.config.authMode || 'app-only';

    if (authMode === 'delegated') {
      // Try silent acquisition first, which uses MSAL's token cache and refresh tokens
      return this.acquireTokenSilent();
    } else {
      // App-only authentication (client credentials flow)
      try {
        const authResult = await (this.msalClient as ConfidentialClientApplication).acquireTokenByClientCredential({
          scopes: ['https://graph.microsoft.com/.default'],
        });

        if (!authResult || !authResult.accessToken) {
          throw new Error('Failed to acquire access token');
        }

        this.accessToken = authResult.accessToken;
        this.tokenExpiry = new Date(authResult.expiresOn!.getTime() - 5 * 60 * 1000);

        return this.accessToken;
      } catch (error) {
        throw new Error(`Authentication failed: ${error instanceof Error ? error.message : String(error)}`);
      }
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
   * For delegated auth, returns 'me' (current signed-in user)
   * For app-only auth, returns the configured userId or userPrincipalName
   */
  getUserId(): string {
    const authMode = this.config.authMode || 'app-only';

    if (authMode === 'delegated') {
      // In delegated mode, always use 'me' for the authenticated user
      return 'me';
    } else {
      // In app-only mode, must specify which user to act on behalf of
      if (this.config.userId) {
        return this.config.userId;
      }
      if (this.config.userPrincipalName) {
        return this.config.userPrincipalName;
      }
      throw new Error('USER_PRINCIPAL_NAME or USER_ID must be specified for app-only authentication');
    }
  }
}
