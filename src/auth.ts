/**
 * Microsoft Graph API Authentication
 *
 * Token Lifetime and Persistence (Updated for 3-Month Authentication):
 * - Access tokens: ~1 hour (automatically refreshed in background)
 * - Refresh tokens: 90 days (3 months) by default, can be extended to 6+ months with Conditional Access policies
 * - Refresh tokens are "rolling" - each use before expiry renews them for another 90 days
 * - As long as the connector is used at least once within 90 days, authentication persists indefinitely
 * - Tokens are proactively refreshed 10 minutes before expiry (improved from 5 minutes)
 * - Retry logic handles transient network failures (3 retries with exponential backoff)
 * - Enhanced diagnostics help identify when refresh tokens expire or are revoked
 *
 * To maintain 3-month authentication persistence:
 * 1. Use the connector at least once every 90 days
 * 2. Each usage automatically extends the refresh token for another 90 days
 * 3. No re-authentication needed as long as the connector is used regularly
 * 4. If refresh token expires (90+ days of inactivity), user must re-authenticate
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

  // Scopes for delegated authentication - must be consistent between acquireTokenByDeviceCode and acquireTokenSilent
  private static readonly DELEGATED_SCOPES = [
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
  ];

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
          if (fs.existsSync(this.tokenCachePath)) {
            const stats = fs.statSync(this.tokenCachePath);
            const fileContent = fs.readFileSync(this.tokenCachePath, 'utf-8');

            let decryptedData: string;
            try {
              decryptedData = this.decryptTokenCache(fileContent);
              console.error('[AUTH] Cache file decrypted successfully');
            } catch {
              // Fall back to unencrypted format (for backward compatibility)
              decryptedData = fileContent;
              console.error('[AUTH] Using unencrypted cache format');
            }

            const cache: TokenCache = JSON.parse(decryptedData);

            // Load MSAL cache if available (version 2+)
            if (cache.version !== undefined && cache.version >= 2 && cache.msalCache) {
              cacheContext.tokenCache.deserialize(cache.msalCache);
              const lastUpdated = cache.lastUpdated ? new Date(cache.lastUpdated).toLocaleString() : 'unknown';

              // Parse MSAL cache to show what's in it
              const msalCache = JSON.parse(cache.msalCache);
              const accountCount = Object.keys(msalCache.Account || {}).length;
              const accessTokenCount = Object.keys(msalCache.AccessToken || {}).length;
              const refreshTokenCount = Object.keys(msalCache.RefreshToken || {}).length;

              console.error(`[AUTH] Loaded cached session (${stats.size} bytes, saved ${lastUpdated})`);
              console.error(`[AUTH] Cache contains: ${accountCount} accounts, ${accessTokenCount} access tokens, ${refreshTokenCount} refresh tokens`);
            } else {
              console.error('[AUTH] Cache file exists but no valid MSAL cache data found');
            }
          } else {
            console.error('[AUTH] No cache file found at:', this.tokenCachePath);
          }
        } catch (error) {
          console.error('[AUTH] Cache load error:', error instanceof Error ? error.message : error);
          // Don't throw - allow MSAL to continue with empty cache
        }
      },

      afterCacheAccess: async (cacheContext: TokenCacheContext): Promise<void> => {
        if (cacheContext.cacheHasChanged) {
          try {
            // Serialize the entire MSAL cache (includes refresh tokens, accounts, etc.)
            const msalCacheData = cacheContext.tokenCache.serialize();

            // Parse to show what we're saving
            const msalCache = JSON.parse(msalCacheData);
            const accountCount = Object.keys(msalCache.Account || {}).length;
            const accessTokenCount = Object.keys(msalCache.AccessToken || {}).length;
            const refreshTokenCount = Object.keys(msalCache.RefreshToken || {}).length;

            console.error(`[AUTH] Saving cache with: ${accountCount} accounts, ${accessTokenCount} access tokens, ${refreshTokenCount} refresh tokens`);

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
            console.error(`[AUTH] Session cached successfully (${stats.size} bytes) at ${this.tokenCachePath}`);

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
            console.error('[AUTH] Cache save error:', error instanceof Error ? error.message : error);
          }
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
            const parsed = JSON.parse(fileContent);

            // Check if this is an encrypted envelope that we couldn't decrypt
            // Encrypted envelopes have: {version: 1, iv: "...", data: "...", authTag: "..."}
            if (parsed.iv && parsed.data && parsed.authTag) {
              console.error('[AUTH] ⚠️  Token cache is encrypted but decryption failed');
              console.error('[AUTH]     This happens when the encryption key changed:');
              console.error('[AUTH]     - Your machine hostname or username changed');
              console.error('[AUTH]     - You moved the cache file from another machine');
              console.error('[AUTH]     - TOKEN_ENCRYPTION_KEY env var changed');
              console.error('[AUTH]     You will need to re-authenticate. The old cache will be overwritten.');
              console.error('[AUTH]     To use this cache, restore the original encryption key.');
              return;
            }

            // Not encrypted - old unencrypted format
            console.error('[AUTH] Token cache is unencrypted (old format), will re-encrypt on next save');
            decryptedData = fileContent;
          } catch {
            // File is not valid JSON - likely corrupted
            console.error('[AUTH] ⚠️  Token cache file is corrupted (not valid JSON)');
            console.error('[AUTH]     Deleting corrupted cache and starting fresh...');

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
      scopes: GraphAuthProvider.DELEGATED_SCOPES,
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
          // Use 2-minute buffer (consistent with other token acquisitions)
          this.tokenExpiry = new Date(response.expiresOn!.getTime() - 2 * 60 * 1000);
          this.userAccount = response.account;

          const expiryTime = new Date(response.expiresOn!).toLocaleString();
          console.error('✅ Authentication successful! Session saved and will persist for 90 days (3 months) with automatic refresh.');
          console.error(`[AUTH] Access token expires at: ${expiryTime}`);
          console.error('[AUTH] Refresh token will be automatically used for seamless re-authentication');

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
   * Implements retry logic for transient failures to improve reliability
   */
  private async acquireTokenSilent(retryCount: number = 0): Promise<string> {
    const MAX_RETRIES = 3;
    const RETRY_DELAY_MS = 2000;

    try {
      // Get all accounts from MSAL cache
      const accounts = await (this.msalClient as PublicClientApplication).getTokenCache().getAllAccounts();

      console.error(`[AUTH] Attempting silent token acquisition, found ${accounts.length} account(s) in cache`);

      if (accounts.length === 0) {
        // No accounts in cache, need to authenticate
        console.error('[AUTH] No cached session found, authentication required');
        return this.acquireTokenByDeviceCode();
      }

      // Use the first account (or the cached account if available)
      const account = this.userAccount || accounts[0];
      console.error(`[AUTH] Using account: ${account.username || account.homeAccountId}`);
      console.error(`[AUTH] Requesting scopes: ${GraphAuthProvider.DELEGATED_SCOPES.join(', ')}`);

      // Proactively refresh if token will expire soon (within 10 minutes)
      // This prevents using tokens that are about to expire
      const shouldForceRefresh = !!(this.tokenExpiry &&
                                    (this.tokenExpiry.getTime() - Date.now()) < 10 * 60 * 1000);

      if (shouldForceRefresh) {
        console.error('[AUTH] Token expiring soon, forcing refresh from refresh token');
      }

      const response = await (this.msalClient as PublicClientApplication).acquireTokenSilent({
        account,
        scopes: GraphAuthProvider.DELEGATED_SCOPES,
        forceRefresh: shouldForceRefresh, // Proactively refresh if expiring soon
      });

      if (!response || !response.accessToken) {
        console.error('[AUTH] Silent token acquisition returned no token');
        return this.acquireTokenByDeviceCode();
      }

      this.accessToken = response.accessToken;
      // Keep token valid for longer - only subtract 2 minutes instead of 5
      // This reduces unnecessary refresh attempts
      this.tokenExpiry = new Date(response.expiresOn!.getTime() - 2 * 60 * 1000);
      this.userAccount = response.account;

      const tokenSource = response.fromCache ? 'cache' : 'refresh token';
      const expiryTime = new Date(response.expiresOn!).toLocaleString();
      console.error(`[AUTH] ✅ Successfully acquired token from ${tokenSource} for ${account.username || 'user'}`);
      console.error(`[AUTH] Token expires at: ${expiryTime} (valid for ${Math.round((response.expiresOn!.getTime() - Date.now()) / 1000 / 60)} minutes)`);

      return this.accessToken;
    } catch (error) {
      const errorMsg = error instanceof Error ? error.message : String(error);
      const errorCode = (error as any)?.errorCode || 'unknown';

      // Enhanced error diagnostics
      console.error(`[AUTH] ⚠️  Silent token acquisition failed (attempt ${retryCount + 1}/${MAX_RETRIES + 1})`);
      console.error(`[AUTH] Error code: ${errorCode}`);
      console.error(`[AUTH] Error message: ${errorMsg}`);

      // Check for specific error types that indicate refresh token issues
      const isRefreshTokenExpired = errorMsg.includes('AADSTS700082') || // Token expired
                                     errorMsg.includes('AADSTS700084') || // Refresh token expired
                                     errorMsg.includes('AADSTS50173') || // Refresh token expired
                                     errorCode === 'invalid_grant';

      const isTransientError = errorMsg.includes('ECONNRESET') ||
                              errorMsg.includes('ETIMEDOUT') ||
                              errorMsg.includes('ENOTFOUND') ||
                              errorCode === 'network_error' ||
                              errorCode === 'service_unavailable';

      if (isRefreshTokenExpired) {
        console.error('[AUTH] ❌ Refresh token has expired or been revoked');
        console.error('[AUTH] This typically happens after 90 days of inactivity or if:');
        console.error('[AUTH]   - Your password was changed');
        console.error('[AUTH]   - An admin revoked the token');
        console.error('[AUTH]   - Conditional access policy changed');
        console.error('[AUTH] You will need to re-authenticate.');
        return this.acquireTokenByDeviceCode();
      }

      if (isTransientError && retryCount < MAX_RETRIES) {
        console.error(`[AUTH] Transient network error detected, retrying in ${RETRY_DELAY_MS}ms...`);
        await new Promise(resolve => setTimeout(resolve, RETRY_DELAY_MS * (retryCount + 1)));
        return this.acquireTokenSilent(retryCount + 1);
      }

      // For other errors, fall back to device code authentication
      console.error(`[AUTH] Falling back to device code authentication`);
      return this.acquireTokenByDeviceCode();
    }
  }

  /**
   * Get an access token for Microsoft Graph API
   * This is the main entry point for obtaining tokens
   * Implements proactive token refresh to maintain 3-month authentication persistence
   */
  async getAccessToken(): Promise<string> {
    try {
      const authMode = this.config.authMode || 'app-only';

      // For delegated auth, check if we should proactively refresh
      // Refresh if token expires in less than 10 minutes (more proactive than before)
      const shouldRefresh = !this.accessToken ||
                           !this.tokenExpiry ||
                           (this.tokenExpiry.getTime() - Date.now()) < 10 * 60 * 1000;

      // Return cached token only if it's still valid for at least 10 minutes
      if (this.accessToken && this.tokenExpiry && !shouldRefresh) {
        const minutesRemaining = Math.round((this.tokenExpiry.getTime() - Date.now()) / 1000 / 60);
        console.error(`[AUTH] Using cached token (expires in ${minutesRemaining} minutes)`);
        return this.accessToken;
      }

      if (authMode === 'delegated') {
        // Try silent acquisition first, which uses MSAL's token cache and refresh tokens
        // This will automatically fall back to device code flow if refresh token is expired
        // The refresh token persists for 90 days (3 months) and is automatically renewed
        // with each use, providing seamless 3-month authentication persistence
        console.error('[AUTH] Refreshing access token using refresh token...');
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
          // Reduce buffer from 5 minutes to 2 minutes for app-only as well
          this.tokenExpiry = new Date(authResult.expiresOn!.getTime() - 2 * 60 * 1000);

          const expiryTime = new Date(authResult.expiresOn!).toLocaleString();
          console.error(`[AUTH] ✅ App-only token acquired, expires at: ${expiryTime}`);

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
