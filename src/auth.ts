/**
 * Microsoft Graph API Authentication
 */

import { ConfidentialClientApplication, PublicClientApplication, DeviceCodeRequest } from '@azure/msal-node';
import { Client } from '@microsoft/microsoft-graph-client';
import { GraphConfig } from './types.js';
import * as fs from 'fs';
import * as path from 'path';
import * as os from 'os';
import * as crypto from 'crypto';
import * as https from 'https';

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
   * Get or generate encryption key for token cache (OWASP A02: Cryptographic Failures)
   */
  private getEncryptionKey(): Buffer {
    // Try to get key from environment variable first (most secure - user-provided)
    if (process.env.TOKEN_ENCRYPTION_KEY) {
      const key = Buffer.from(process.env.TOKEN_ENCRYPTION_KEY, 'hex');
      if (key.length === 32) {
        return key;
      }
      console.error('[AUTH] Invalid TOKEN_ENCRYPTION_KEY length, using derived key');
    }

    // Fall back to deriving key from machine-specific data
    // This is less secure but provides automatic encryption without user configuration
    const machineId = `${os.hostname()}-${os.userInfo().username}-office365-mcp`;
    return crypto.scryptSync(machineId, 'office365-mcp-salt', 32);
  }

  /**
   * Encrypt token cache data (OWASP A02: Cryptographic Failures)
   */
  private encryptTokenCache(data: string): string {
    try {
      const key = this.getEncryptionKey();
      const iv = crypto.randomBytes(16);
      const cipher = crypto.createCipheriv('aes-256-gcm', key, iv);

      let encrypted = cipher.update(data, 'utf8', 'hex');
      encrypted += cipher.final('hex');

      const authTag = cipher.getAuthTag();

      // Return JSON with iv, encrypted data, and auth tag
      return JSON.stringify({
        version: 1,
        iv: iv.toString('hex'),
        data: encrypted,
        authTag: authTag.toString('hex'),
      });
    } catch (error) {
      console.error('[AUTH] Encryption failed:', error);
      throw new Error('Failed to encrypt token cache');
    }
  }

  /**
   * Decrypt token cache data (OWASP A02: Cryptographic Failures)
   */
  private decryptTokenCache(encryptedData: string): string {
    try {
      const key = this.getEncryptionKey();
      const parsed = JSON.parse(encryptedData);

      // Validate encrypted data structure
      if (!parsed.iv || !parsed.data || !parsed.authTag || parsed.version !== 1) {
        throw new Error('Invalid encrypted data format');
      }

      const iv = Buffer.from(parsed.iv, 'hex');
      const encrypted = Buffer.from(parsed.data, 'hex');
      const authTag = Buffer.from(parsed.authTag, 'hex');

      const decipher = crypto.createDecipheriv('aes-256-gcm', key, iv);
      decipher.setAuthTag(authTag);

      let decrypted = decipher.update(encrypted, undefined, 'utf8');
      decrypted += decipher.final('utf8');

      return decrypted;
    } catch (error) {
      console.error('[AUTH] Decryption failed:', error);
      throw new Error('Failed to decrypt token cache');
    }
  }

  /**
   * Validate token cache structure (OWASP A08: Data Integrity)
   */
  private validateTokenCache(data: any): data is TokenCache {
    if (!data || typeof data !== 'object') {
      return false;
    }

    // Validate required fields
    if (typeof data.accessToken !== 'string' || data.accessToken.length === 0) {
      return false;
    }

    if (typeof data.expiresOn !== 'number' || data.expiresOn < 0) {
      return false;
    }

    // Validate optional fields if present
    if (data.userId !== undefined && typeof data.userId !== 'string') {
      return false;
    }

    // Validate token format (should be a valid JWT or opaque token)
    if (data.accessToken.length < 10 || data.accessToken.length > 10000) {
      return false;
    }

    return true;
  }

  /**
   * Load cached tokens from file
   */
  private loadCachedTokens(): void {
    try {
      if (fs.existsSync(this.tokenCachePath)) {
        const fileContent = fs.readFileSync(this.tokenCachePath, 'utf-8');

        let decryptedData: string;
        try {
          // Try to decrypt first (new encrypted format)
          decryptedData = this.decryptTokenCache(fileContent);
        } catch {
          // Fall back to unencrypted format (for backward compatibility)
          console.error('[AUTH] Token cache not encrypted, will re-encrypt on next save');
          decryptedData = fileContent;
        }

        // Parse JSON
        const rawData = JSON.parse(decryptedData);

        // Validate structure before using (OWASP A08: Software and Data Integrity)
        if (!this.validateTokenCache(rawData)) {
          console.error('[AUTH] Invalid token cache structure, ignoring cached tokens');
          // Delete corrupted cache file
          try {
            fs.unlinkSync(this.tokenCachePath);
          } catch {
            // Ignore deletion errors
          }
          return;
        }

        const cache: TokenCache = rawData;

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
      // Delete potentially corrupted cache file
      try {
        if (fs.existsSync(this.tokenCachePath)) {
          fs.unlinkSync(this.tokenCachePath);
        }
      } catch {
        // Ignore deletion errors
      }
    }
  }

  /**
   * Save tokens to cache (encrypted)
   */
  private saveCachedTokens(accessToken: string, expiresOn: Date, account?: any): void {
    try {
      const cache: TokenCache = {
        accessToken,
        expiresOn: expiresOn.getTime(),
        userId: this.config.userPrincipalName || this.config.userId,
        account,
      };

      // Encrypt the token cache (OWASP A02: Cryptographic Failures)
      const plaintext = JSON.stringify(cache, null, 2);
      const encrypted = this.encryptTokenCache(plaintext);

      // Write to temporary file with restrictive permissions first
      const tempPath = this.tokenCachePath + '.tmp';

      // Write encrypted data with restrictive permissions (0600 = read/write for owner only)
      fs.writeFileSync(tempPath, encrypted, { mode: 0o600 });

      // Atomically rename to final location
      fs.renameSync(tempPath, this.tokenCachePath);

      // Verify permissions on the final file
      try {
        const stats = fs.statSync(this.tokenCachePath);
        const permissions = stats.mode & 0o777;
        if (permissions !== 0o600) {
          console.error(`[AUTH] WARNING: Token cache file has insecure permissions (${permissions.toString(8)}). Expected 0600.`);
          // Try to fix permissions
          fs.chmodSync(this.tokenCachePath, 0o600);
        }
      } catch (permError) {
        console.error('[AUTH] Failed to verify/fix token cache permissions:', permError);
      }

      console.error(`[AUTH] Saved encrypted tokens to cache: ${this.tokenCachePath}`);
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
  /**
   * Create a secure HTTPS agent with TLS 1.3 enforcement (OWASP A05: Security Misconfiguration)
   */
  private createSecureHttpsAgent(): https.Agent {
    return new https.Agent({
      // Enforce TLS 1.3 or TLS 1.2 minimum
      minVersion: 'TLSv1.2',
      maxVersion: 'TLSv1.3',
      // Reject unauthorized certificates
      rejectUnauthorized: true,
      // Disable insecure ciphers
      ciphers: [
        'TLS_AES_256_GCM_SHA384',
        'TLS_AES_128_GCM_SHA256',
        'TLS_CHACHA20_POLY1305_SHA256',
        'ECDHE-RSA-AES256-GCM-SHA384',
        'ECDHE-RSA-AES128-GCM-SHA256',
      ].join(':'),
      // Enable secure renegotiation
      secureOptions: crypto.constants.SSL_OP_NO_SSLv2 |
                     crypto.constants.SSL_OP_NO_SSLv3 |
                     crypto.constants.SSL_OP_NO_TLSv1 |
                     crypto.constants.SSL_OP_NO_TLSv1_1,
    });
  }

  async getGraphClient(): Promise<Client> {
    const accessToken = await this.getAccessToken();

    // Create secure HTTPS agent (OWASP A05: Security Misconfiguration)
    const httpsAgent = this.createSecureHttpsAgent();

    return Client.init({
      authProvider: (done) => {
        done(null, accessToken);
      },
      // Use custom fetch with secure HTTPS agent
      fetchOptions: {
        agent: httpsAgent,
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

    console.error(`[AUTH] getUserId called, authMode: ${authMode}`);

    if (authMode === 'delegated') {
      // In delegated mode, always use 'me' for the authenticated user
      console.error(`[AUTH] Returning 'me' for delegated auth`);
      return 'me';
    } else {
      // In app-only mode, must specify which user to act on behalf of
      if (this.config.userId) {
        console.error(`[AUTH] Returning userId: ${this.config.userId}`);
        return this.config.userId;
      }
      if (this.config.userPrincipalName) {
        console.error(`[AUTH] Returning userPrincipalName: ${this.config.userPrincipalName}`);
        return this.config.userPrincipalName;
      }
      throw new Error('USER_PRINCIPAL_NAME or USER_ID must be specified for app-only authentication');
    }
  }

  /**
   * Revoke the current token and clear the cache (OWASP A07: Authentication Failures)
   * This should be called on logout or when the user wants to invalidate their session
   */
  async revokeToken(): Promise<void> {
    console.error('[AUTH] Revoking token and clearing cache');

    try {
      // Clear in-memory token
      this.accessToken = null;
      this.tokenExpiry = null;
      this.userAccount = null;

      // Delete the cached token file
      if (fs.existsSync(this.tokenCachePath)) {
        fs.unlinkSync(this.tokenCachePath);
        console.error('[AUTH] Token cache file deleted');
      }

      // For delegated auth, remove the cached account from MSAL
      if (this.config.authMode === 'delegated' && this.userAccount) {
        try {
          const accounts = await (this.msalClient as PublicClientApplication).getTokenCache().getAllAccounts();
          for (const account of accounts) {
            await (this.msalClient as PublicClientApplication).getTokenCache().removeAccount(account);
          }
          console.error('[AUTH] MSAL account cache cleared');
        } catch (error) {
          console.error('[AUTH] Failed to clear MSAL account cache:', error);
        }
      }

      console.error('[AUTH] Token revocation completed');
    } catch (error) {
      console.error('[AUTH] Error during token revocation:', error);
      throw new Error('Failed to revoke token');
    }
  }

  /**
   * Check if the user is currently authenticated
   */
  isAuthenticated(): boolean {
    return this.accessToken !== null &&
           this.tokenExpiry !== null &&
           this.tokenExpiry > new Date();
  }
}
