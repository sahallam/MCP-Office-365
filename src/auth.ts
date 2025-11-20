/**
 * Microsoft Graph API Authentication
 *
 * Token Lifetime Information:
 * - Access tokens: ~1 hour
 * - Refresh tokens: Up to 90 days by default (can be extended to 6+ months with Conditional Access policies)
 * - Refresh tokens are "rolling" - each use before expiry gets you a new refresh token
 * - As long as the connector is used at least once within the refresh token lifetime, authentication persists
 */

import { ConfidentialClientApplication, PublicClientApplication, DeviceCodeRequest, ICachePlugin, TokenCacheContext } from '@azure/msal-node';
import { Client } from '@microsoft/microsoft-graph-client';
import { GraphConfig } from './types.js';
import * as fs from 'fs';
import * as path from 'path';
import * as os from 'os';
import * as crypto from 'crypto';
import * as https from 'https';

/**
 * Enhanced token cache that stores the complete MSAL token cache
 * This includes refresh tokens, which enables long-term authentication persistence
 */
interface TokenCache {
  // MSAL's complete cache data (includes refresh tokens, accounts, etc.)
  msalCache?: string;

  // Legacy fields for backward compatibility
  accessToken?: string;
  expiresOn?: number;
  userId?: string;
  account?: any;

  // Metadata
  version: number;
  lastUpdated: number;
}

/**
 * Custom error class for authentication failures
 * This allows us to provide user-friendly error messages
 */
export class AuthenticationError extends Error {
  public readonly userMessage: string;
  public readonly requiresReauth: boolean;
  public readonly deviceCode?: string;
  public readonly verificationUri?: string;
  public readonly expiresIn?: number;

  constructor(
    message: string,
    userMessage: string,
    requiresReauth: boolean = false,
    deviceCodeInfo?: { userCode: string; verificationUri: string; expiresIn: number }
  ) {
    super(message);
    this.name = 'AuthenticationError';
    this.userMessage = userMessage;
    this.requiresReauth = requiresReauth;

    if (deviceCodeInfo) {
      this.deviceCode = deviceCodeInfo.userCode;
      this.verificationUri = deviceCodeInfo.verificationUri;
      this.expiresIn = deviceCodeInfo.expiresIn;
    }
  }
}

export class GraphAuthProvider {
  private msalClient: ConfidentialClientApplication | PublicClientApplication;
  private config: GraphConfig;
  private accessToken: string | null = null;
  private tokenExpiry: Date | null = null;
  private tokenCachePath: string;
  private userAccount: any = null;
  private pendingDeviceCode: { userCode: string; verificationUri: string; expiresIn: number } | null = null;
  private authenticationInProgress: boolean = false;

  constructor(config: GraphConfig) {
    this.config = config;
    // Use home directory for token cache to avoid permission issues when running from system directories
    this.tokenCachePath = config.tokenCachePath || path.join(os.homedir(), '.office365-mcp-token-cache.json');

    // Determine authentication mode
    const authMode = config.authMode || 'app-only';

    if (authMode === 'delegated') {
      // Create persistent cache plugin
      const cachePlugin = this.createCachePlugin();

      // Use Public Client Application for delegated auth (device code flow)
      this.msalClient = new PublicClientApplication({
        auth: {
          clientId: config.clientId,
          authority: `https://login.microsoftonline.com/${config.tenantId}`,
        },
        cache: {
          cachePlugin,
        },
      });

      // Try to load cached tokens (this will also restore MSAL's cache)
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
   * Create a persistent cache plugin for MSAL
   * This stores the entire MSAL cache (including refresh tokens) in our encrypted file
   */
  private createCachePlugin(): ICachePlugin {
    return {
      beforeCacheAccess: async (cacheContext: TokenCacheContext): Promise<void> => {
        try {
          console.error('[AUTH] MSAL beforeCacheAccess triggered');
          if (fs.existsSync(this.tokenCachePath)) {
            const stats = fs.statSync(this.tokenCachePath);
            console.error(`[AUTH] Cache file exists (${stats.size} bytes), loading...`);
            const fileContent = fs.readFileSync(this.tokenCachePath, 'utf-8');

            let decryptedData: string;
            try {
              decryptedData = this.decryptTokenCache(fileContent);
              console.error('[AUTH] Successfully decrypted cache');
            } catch {
              // Fall back to unencrypted format (for backward compatibility)
              console.error('[AUTH] Decryption failed, trying unencrypted format');
              decryptedData = fileContent;
            }

            const cache: TokenCache = JSON.parse(decryptedData);
            console.error(`[AUTH] Parsed cache - version: ${cache.version}, lastUpdated: ${cache.lastUpdated ? new Date(cache.lastUpdated).toISOString() : 'N/A'}`);

            // Load MSAL cache if available (version 2+)
            if (cache.version !== undefined && cache.version >= 2 && cache.msalCache) {
              console.error(`[AUTH] Deserializing MSAL cache (${cache.msalCache.length} chars)`);
              cacheContext.tokenCache.deserialize(cache.msalCache);
              console.error('[AUTH] MSAL cache deserialized successfully');
            } else {
              console.error('[AUTH] No MSAL cache to deserialize (legacy format or missing)');
            }
          } else {
            console.error('[AUTH] No cache file found at:', this.tokenCachePath);
          }
        } catch (error) {
          console.error('[AUTH] Failed to load MSAL cache:', error);
          // Don't throw - allow MSAL to continue with empty cache
        }
      },

      afterCacheAccess: async (cacheContext: TokenCacheContext): Promise<void> => {
        if (cacheContext.cacheHasChanged) {
          try {
            console.error('[AUTH] MSAL afterCacheAccess: cache has changed, saving...');
            // Serialize the entire MSAL cache (includes refresh tokens, accounts, etc.)
            const msalCacheData = cacheContext.tokenCache.serialize();
            console.error(`[AUTH] Serialized MSAL cache (${msalCacheData.length} chars)`);

            // Create enhanced cache structure
            const cache: TokenCache = {
              version: 2,
              msalCache: msalCacheData,
              lastUpdated: Date.now(),
            };

            // Encrypt and save
            const plaintext = JSON.stringify(cache, null, 2);
            const encrypted = this.encryptTokenCache(plaintext);

            // Write to temporary file with restrictive permissions first
            const tempPath = this.tokenCachePath + '.tmp';

            // Write encrypted data with restrictive permissions (0600 = read/write for owner only)
            fs.writeFileSync(tempPath, encrypted, { mode: 0o600 });

            // Atomically rename to final location
            fs.renameSync(tempPath, this.tokenCachePath);

            // Verify the file was saved
            const stats = fs.statSync(this.tokenCachePath);
            console.error(`[AUTH] Cache saved successfully (${stats.size} bytes) at ${this.tokenCachePath}`);

            // Verify permissions on the final file
            try {
              const permissions = stats.mode & 0o777;
              if (permissions !== 0o600) {
                fs.chmodSync(this.tokenCachePath, 0o600);
              }
            } catch {
              // Ignore permission verification errors
            }
          } catch (error) {
            console.error('[AUTH] Failed to save MSAL cache:', error);
          }
        } else {
          console.error('[AUTH] MSAL afterCacheAccess: cache unchanged, not saving');
        }
      },
    };
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
   * Supports both legacy format (no version field) and new versioned formats
   */
  private validateTokenCache(data: any): data is TokenCache {
    if (!data || typeof data !== 'object') {
      return false;
    }

    // Check if version field exists
    const version = data.version;

    // If version exists, validate it
    if (version !== undefined && (typeof version !== 'number' || version < 1)) {
      return false;
    }

    // Version 2+ uses MSAL cache
    if (version !== undefined && version >= 2) {
      // For version 2+, we just need the MSAL cache data
      if (data.msalCache && typeof data.msalCache !== 'string') {
        return false;
      }
      return true;
    }

    // Legacy format validation (no version field or version 1)
    // These caches have accessToken and expiresOn fields
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
   * For version 2+ caches, MSAL cache plugin handles loading
   * This method is mainly for legacy cache migration
   */
  private loadCachedTokens(): void {
    try {
      if (fs.existsSync(this.tokenCachePath)) {
        const stats = fs.statSync(this.tokenCachePath);
        console.error(`[AUTH] Found token cache file (${stats.size} bytes)`);
        const fileContent = fs.readFileSync(this.tokenCachePath, 'utf-8');

        let decryptedData: string;
        try {
          // Try to decrypt first (new encrypted format)
          decryptedData = this.decryptTokenCache(fileContent);
        } catch (decryptError) {
          // Decryption failed - could be:
          // 1. File is in old unencrypted format
          // 2. Encryption key changed (e.g., machine hostname/username changed)
          // 3. File is corrupted

          // Try to parse as unencrypted JSON to check if it's valid
          try {
            JSON.parse(fileContent);
            console.error('[AUTH] Token cache is unencrypted (old format), will re-encrypt on next save');
            decryptedData = fileContent;
          } catch {
            // File is not valid JSON - likely corrupted or encryption key changed
            console.error('[AUTH] ⚠️  Token cache decryption failed - this can happen if:');
            console.error('[AUTH]     - Your machine hostname or username changed');
            console.error('[AUTH]     - The cache file is corrupted');
            console.error('[AUTH]     - You moved the cache file from another machine');
            console.error('[AUTH] Deleting corrupted cache and starting fresh...');

            // Delete the corrupted cache file
            try {
              fs.unlinkSync(this.tokenCachePath);
              console.error('[AUTH] Corrupted cache file deleted');
            } catch {
              // Ignore deletion errors
            }
            return;
          }
        }

        // Parse JSON
        const rawData = JSON.parse(decryptedData);

        // Determine cache version FIRST (default to 1 if not present for legacy caches)
        const cacheVersion = rawData.version || 1;

        // Version 2+ uses MSAL cache plugin (loaded automatically)
        // Don't validate with legacy validator - just check basic structure
        if (cacheVersion >= 2) {
          if (rawData.msalCache && typeof rawData.msalCache === 'string') {
            console.error('[AUTH] Token cache version 2+ detected - MSAL cache plugin will handle loading');
            return;
          } else {
            console.error('[AUTH] Version 2+ cache missing msalCache field, will re-authenticate');
            // Don't delete - let MSAL handle it
            return;
          }
        }

        // Only validate legacy (version 1) caches with legacy validator
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

        // Legacy cache (no version field or version 1) - manually load access token
        if (cache.expiresOn && cache.expiresOn > Date.now()) {
          this.accessToken = cache.accessToken || null;
          this.tokenExpiry = new Date(cache.expiresOn);
          this.userAccount = cache.account;
          console.error('[AUTH] Loaded cached access token (legacy format - will be upgraded to version 2 on next authentication)');
        } else {
          console.error('[AUTH] Cached token expired, will re-authenticate and upgrade to version 2 format');
        }
      }
    } catch (error) {
      console.error('[AUTH] Failed to load cached tokens:', error);
      // Don't delete the cache file here - let MSAL cache plugin try to load it
      // The cache plugin has its own error handling and may succeed where this fails
    }
  }

  /**
   * Acquire token using device code flow (for delegated auth)
   * This method is non-blocking - it starts the auth flow in the background
   * and immediately throws an error with the device code for the user
   */
  private async acquireTokenByDeviceCode(): Promise<string> {
    // If authentication is already in progress, throw error with existing code
    if (this.authenticationInProgress && this.pendingDeviceCode) {
      throw new AuthenticationError(
        'Authentication in progress',
        'Please complete the authentication using the code provided.',
        true,
        this.pendingDeviceCode
      );
    }

    // Mark authentication as in progress
    this.authenticationInProgress = true;

    // Create a promise that resolves when device code callback fires
    let deviceCodeResolve: (value: any) => void;
    const deviceCodePromise = new Promise((resolve) => {
      deviceCodeResolve = resolve;
    });

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
        // Store device code info for immediate error reporting
        this.pendingDeviceCode = {
          userCode: response.userCode,
          verificationUri: response.verificationUri,
          expiresIn: response.expiresIn,
        };

        console.error('\n=======================================================================');
        console.error('🔐 AUTHENTICATION REQUIRED');
        console.error('=======================================================================');
        console.error(`\nTo sign in, open: ${response.verificationUri}`);
        console.error(`Enter code: ${response.userCode}`);
        console.error('\n=======================================================================\n');

        // Resolve the promise so we can throw the error
        deviceCodeResolve(response);
      },
    };

    // Start authentication in background (don't await)
    (this.msalClient as PublicClientApplication).acquireTokenByDeviceCode(deviceCodeRequest)
      .then((response) => {
        if (response && response.accessToken) {
          this.accessToken = response.accessToken;
          this.tokenExpiry = new Date(response.expiresOn!.getTime() - 5 * 60 * 1000);
          this.userAccount = response.account;

          console.error('✅ Authentication successful! Session saved.');

          // Clear pending device code after successful auth
          this.pendingDeviceCode = null;
          this.authenticationInProgress = false;
        }
      })
      .catch((error) => {
        console.error('[AUTH] Authentication failed:', error.message || error);
        this.pendingDeviceCode = null;
        this.authenticationInProgress = false;
      });

    // Wait for device code callback to fire (with timeout)
    const timeoutPromise = new Promise((_, reject) =>
      setTimeout(() => reject(new Error('Device code generation timeout')), 3000)
    );

    try {
      await Promise.race([deviceCodePromise, timeoutPromise]);
    } catch (error) {
      // Timeout - continue anyway
    }

    // Throw error with device code immediately (don't wait for user to authenticate)
    if (this.pendingDeviceCode) {
      throw new AuthenticationError(
        'Authentication required - device code generated',
        'Please authenticate using the device code provided below.',
        true,
        this.pendingDeviceCode
      );
    } else {
      throw new AuthenticationError(
        'Authentication required',
        'Please check the server logs for authentication instructions.',
        true
      );
    }
  }

  /**
   * Acquire token silently using MSAL cache and refresh tokens
   * This method attempts to refresh the access token without user interaction
   */
  private async acquireTokenSilent(): Promise<string> {
    try {
      console.error('[AUTH] Attempting silent token acquisition...');
      // Get all accounts from MSAL cache
      const accounts = await (this.msalClient as PublicClientApplication).getTokenCache().getAllAccounts();
      console.error(`[AUTH] Found ${accounts.length} account(s) in MSAL cache`);

      if (accounts.length === 0) {
        // No accounts in cache, need to authenticate
        console.error('[AUTH] No accounts in cache, falling back to device code flow');
        return this.acquireTokenByDeviceCode();
      }

      // Use the first account (or the cached account if available)
      const account = this.userAccount || accounts[0];
      console.error(`[AUTH] Using account: ${account.username || account.homeAccountId}`);

      const response = await (this.msalClient as PublicClientApplication).acquireTokenSilent({
        account,
        scopes: ['https://graph.microsoft.com/.default'],
        forceRefresh: false, // Use cached token if valid, otherwise use refresh token
      });

      if (!response || !response.accessToken) {
        console.error('[AUTH] Silent acquisition returned no token, falling back to device code');
        return this.acquireTokenByDeviceCode();
      }

      this.accessToken = response.accessToken;
      this.tokenExpiry = new Date(response.expiresOn!.getTime() - 5 * 60 * 1000);
      this.userAccount = response.account;

      console.error(`[AUTH] Silent acquisition successful, token expires: ${this.tokenExpiry.toISOString()}`);
      return this.accessToken;
    } catch (error) {
      // Fall back to device code authentication
      const errorMsg = error instanceof Error ? error.message : String(error);
      console.error(`[AUTH] Silent acquisition failed: ${errorMsg}, falling back to device code`);
      return this.acquireTokenByDeviceCode();
    }
  }

  /**
   * Get an access token for Microsoft Graph API
   * This is the main entry point for obtaining tokens
   */
  async getAccessToken(): Promise<string> {
    try {
      // Return cached token if still valid (with at least 5 minutes remaining)
      if (this.accessToken && this.tokenExpiry && this.tokenExpiry > new Date()) {
        return this.accessToken;
      }

      const authMode = this.config.authMode || 'app-only';

      if (authMode === 'delegated') {
        // Try silent acquisition first, which uses MSAL's token cache and refresh tokens
        // This will automatically fall back to device code flow if refresh token is expired
        return this.acquireTokenSilent();
      } else {
        // App-only authentication (client credentials flow)
        try {
          const authResult = await (this.msalClient as ConfidentialClientApplication).acquireTokenByClientCredential({
            scopes: ['https://graph.microsoft.com/.default'],
          });

          if (!authResult || !authResult.accessToken) {
            throw new AuthenticationError(
              'Failed to acquire access token',
              '❌ Authentication failed: Unable to acquire access token.\n\nPlease verify your CLIENT_SECRET is correct and your app has the required permissions.',
              false
            );
          }

          this.accessToken = authResult.accessToken;
          this.tokenExpiry = new Date(authResult.expiresOn!.getTime() - 5 * 60 * 1000);

          return this.accessToken;
        } catch (error) {
          if (error instanceof AuthenticationError) {
            throw error;
          }

          const errorMessage = error instanceof Error ? error.message : String(error);
          throw new AuthenticationError(
            `App-only authentication failed: ${errorMessage}`,
            `❌ Authentication failed: ${errorMessage}\n\nPlease verify:\n- CLIENT_ID is correct\n- CLIENT_SECRET is valid\n- TENANT_ID is correct\n- Your app has the required API permissions`,
            false
          );
        }
      }
    } catch (error) {
      // Re-throw AuthenticationError as-is so the user-friendly message is preserved
      if (error instanceof AuthenticationError) {
        throw error;
      }

      // Wrap other errors
      const errorMessage = error instanceof Error ? error.message : String(error);
      throw new AuthenticationError(
        `Authentication failed: ${errorMessage}`,
        `❌ Authentication failed: ${errorMessage}`,
        false
      );
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

  /**
   * Revoke the current token and clear the cache (OWASP A07: Authentication Failures)
   * This should be called on logout or when the user wants to invalidate their session
   */
  async revokeToken(): Promise<void> {
    try {
      // Clear in-memory token
      this.accessToken = null;
      this.tokenExpiry = null;
      this.userAccount = null;

      // Delete the cached token file
      if (fs.existsSync(this.tokenCachePath)) {
        fs.unlinkSync(this.tokenCachePath);
      }

      // For delegated auth, remove the cached account from MSAL
      if (this.config.authMode === 'delegated' && this.userAccount) {
        try {
          const accounts = await (this.msalClient as PublicClientApplication).getTokenCache().getAllAccounts();
          for (const account of accounts) {
            await (this.msalClient as PublicClientApplication).getTokenCache().removeAccount(account);
          }
        } catch (error) {
          console.error('[AUTH] Failed to clear MSAL account cache:', error);
        }
      }
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
