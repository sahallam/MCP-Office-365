/**
 * Audit Logging Module
 * Implements comprehensive audit logging for security-sensitive operations
 * OWASP A09:2021 - Security Logging and Monitoring
 */

import * as fs from 'fs';
import * as path from 'path';
import * as os from 'os';

export interface AuditEvent {
  type: AuditEventType;
  userId?: string;
  action: string;
  resource: string;
  success: boolean;
  ipAddress?: string;
  userAgent?: string;
  details?: string;
  error?: string;
}

export enum AuditEventType {
  // Authentication events
  AUTH_LOGIN = 'AUTH_LOGIN',
  AUTH_LOGOUT = 'AUTH_LOGOUT',
  AUTH_TOKEN_REFRESH = 'AUTH_TOKEN_REFRESH',
  AUTH_TOKEN_REVOKE = 'AUTH_TOKEN_REVOKE',
  AUTH_FAILED = 'AUTH_FAILED',

  // Resource access events
  RESOURCE_READ = 'RESOURCE_READ',
  RESOURCE_CREATE = 'RESOURCE_CREATE',
  RESOURCE_UPDATE = 'RESOURCE_UPDATE',
  RESOURCE_DELETE = 'RESOURCE_DELETE',
  RESOURCE_SEARCH = 'RESOURCE_SEARCH',

  // File operations
  FILE_UPLOAD = 'FILE_UPLOAD',
  FILE_DOWNLOAD = 'FILE_DOWNLOAD',
  FILE_DELETE = 'FILE_DELETE',

  // Email operations
  EMAIL_SEND = 'EMAIL_SEND',
  EMAIL_READ = 'EMAIL_READ',
  EMAIL_DELETE = 'EMAIL_DELETE',

  // Teams operations
  TEAMS_MESSAGE_SEND = 'TEAMS_MESSAGE_SEND',
  TEAMS_MESSAGE_READ = 'TEAMS_MESSAGE_READ',

  // Configuration changes
  CONFIG_CHANGE = 'CONFIG_CHANGE',

  // Security events
  SECURITY_VIOLATION = 'SECURITY_VIOLATION',
  INPUT_VALIDATION_FAILED = 'INPUT_VALIDATION_FAILED',
  PERMISSION_DENIED = 'PERMISSION_DENIED',
}

export class AuditLogger {
  private static instance: AuditLogger;
  private logPath: string;
  private enabled: boolean = true;

  private constructor() {
    // Create audit logs directory in user's home
    const auditDir = path.join(os.homedir(), '.office365-mcp', 'audit');
    if (!fs.existsSync(auditDir)) {
      fs.mkdirSync(auditDir, { recursive: true, mode: 0o700 });
    }

    // Create log file with date
    const logFileName = `audit-${new Date().toISOString().split('T')[0]}.log`;
    this.logPath = path.join(auditDir, logFileName);

    // Disable audit logging if AUDIT_ENABLED env var is set to false
    if (process.env.AUDIT_ENABLED === 'false') {
      this.enabled = false;
    }
  }

  public static getInstance(): AuditLogger {
    if (!AuditLogger.instance) {
      AuditLogger.instance = new AuditLogger();
    }
    return AuditLogger.instance;
  }

  /**
   * Log an audit event
   */
  public log(event: AuditEvent): void {
    if (!this.enabled) {
      return;
    }

    try {
      const logEntry = {
        timestamp: new Date().toISOString(),
        eventType: event.type,
        userId: event.userId || 'unknown',
        action: event.action,
        resource: event.resource,
        success: event.success,
        ipAddress: event.ipAddress,
        userAgent: event.userAgent,
        details: event.details,
        error: event.error,
      };

      // Write to log file (append mode)
      const logLine = JSON.stringify(logEntry) + '\n';
      fs.appendFileSync(this.logPath, logLine, { mode: 0o600 });

      // Also log to console for debugging (if enabled)
      if (process.env.DEBUG_AUDIT === 'true') {
        console.log('[AUDIT]', logLine.trim());
      }
    } catch (error) {
      // Don't throw - we don't want audit logging to break the application
      console.error('[AUDIT] Failed to write audit log:', error);
    }
  }

  /**
   * Log successful operation
   */
  public logSuccess(
    type: AuditEventType,
    action: string,
    resource: string,
    userId?: string,
    details?: string
  ): void {
    this.log({
      type,
      userId,
      action,
      resource,
      success: true,
      details,
    });
  }

  /**
   * Log failed operation
   */
  public logFailure(
    type: AuditEventType,
    action: string,
    resource: string,
    error: string,
    userId?: string,
    details?: string
  ): void {
    this.log({
      type,
      userId,
      action,
      resource,
      success: false,
      error,
      details,
    });
  }

  /**
   * Log security violation
   */
  public logSecurityViolation(
    action: string,
    resource: string,
    reason: string,
    userId?: string
  ): void {
    this.log({
      type: AuditEventType.SECURITY_VIOLATION,
      userId,
      action,
      resource,
      success: false,
      error: reason,
    });
  }

  /**
   * Get audit log path
   */
  public getLogPath(): string {
    return this.logPath;
  }

  /**
   * Rotate log files (clean up old logs)
   */
  public rotateLogFiles(maxAgeDays: number = 90): void {
    if (!this.enabled) {
      return;
    }

    try {
      const auditDir = path.dirname(this.logPath);
      const files = fs.readdirSync(auditDir);
      const now = Date.now();
      const maxAge = maxAgeDays * 24 * 60 * 60 * 1000;

      for (const file of files) {
        if (!file.startsWith('audit-') || !file.endsWith('.log')) {
          continue;
        }

        const filePath = path.join(auditDir, file);
        const stats = fs.statSync(filePath);
        const age = now - stats.mtimeMs;

        if (age > maxAge) {
          fs.unlinkSync(filePath);
          console.log(`[AUDIT] Rotated old log file: ${file}`);
        }
      }
    } catch (error) {
      console.error('[AUDIT] Failed to rotate log files:', error);
    }
  }
}

// Export singleton instance
export const auditLogger = AuditLogger.getInstance();
