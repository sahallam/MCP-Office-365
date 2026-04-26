/**
 * Security utilities for input validation and sanitization
 */

/**
 * Sanitize search queries to prevent OData injection
 */
export function sanitizeSearchQuery(query: string): string {
  if (!query || typeof query !== 'string') {
    throw new Error('Search query must be a non-empty string');
  }

  // Remove or escape OData special characters and operators
  // This prevents injection of operators like $filter, $select, etc.
  return query
    .replace(/['"()$&]/g, '') // Remove special characters
    .replace(/\s+(and|or|not|eq|ne|gt|lt|ge|le)\s+/gi, ' ') // Remove OData operators
    .trim();
}

/**
 * Escape HTML special characters to prevent XSS
 */
export function escapeHtml(text: string): string {
  if (!text || typeof text !== 'string') {
    return '';
  }

  const htmlEscapeMap: { [key: string]: string } = {
    '&': '&amp;',
    '<': '&lt;',
    '>': '&gt;',
    '"': '&quot;',
    "'": '&#x27;',
    '/': '&#x2F;',
  };

  return text.replace(/[&<>"'/]/g, (char) => htmlEscapeMap[char]);
}

/**
 * Sanitize HTML content by escaping all HTML tags and special characters
 *
 * SECURITY WARNING: This function is only suitable for displaying content from
 * TRUSTED sources (like Office 365 emails/documents). It converts HTML to plain
 * text by escaping all HTML tags.
 *
 * For production environments requiring HTML display from untrusted user input,
 * use a comprehensive HTML sanitization library like DOMPurify:
 *
 * Installation: npm install dompurify @types/dompurify
 * Usage: import DOMPurify from 'dompurify';
 *        const clean = DOMPurify.sanitize(dirty);
 *
 * @param html - The HTML string to sanitize
 * @param _allowedTags - Reserved for future use with allowlist-based sanitization
 * @returns Sanitized string with all HTML escaped
 */
export function sanitizeHtmlContent(html: string, _allowedTags: string[] = []): string {
  if (!html || typeof html !== 'string') {
    return '';
  }

  // For security, we escape all HTML rather than trying to filter specific patterns
  // This prevents XSS while preserving the text content for display
  // Note: _allowedTags parameter is available for future enhancement with DOMPurify
  return escapeHtml(html);
}

/**
 * Validate resource IDs to prevent path traversal and injection
 */
export function validateResourceId(id: string, resourceType: string): void {
  if (!id || typeof id !== 'string') {
    throw new Error(`Invalid ${resourceType} ID: must be a non-empty string`);
  }

  // Check for path traversal attempts
  if (id.includes('..') || id.includes('/') || id.includes('\\')) {
    throw new Error(`Invalid ${resourceType} ID: contains illegal characters`);
  }

  // Validate format (Graph API IDs are typically alphanumeric with some special chars)
  // Microsoft Graph IDs can contain: letters, numbers, hyphens, underscores, and some special chars
  const validIdPattern = /^[a-zA-Z0-9_\-=+!]+$/;
  if (!validIdPattern.test(id)) {
    throw new Error(`Invalid ${resourceType} ID format`);
  }

  // Check reasonable length (Graph IDs are typically < 200 chars)
  if (id.length > 200) {
    throw new Error(`Invalid ${resourceType} ID: too long`);
  }
}

/**
 * Validate Excel range address format
 */
export function validateExcelAddress(address: string): void {
  if (!address || typeof address !== 'string') {
    throw new Error('Excel address must be a non-empty string');
  }

  // Valid Excel range patterns:
  // - Single cell: A1, Z99, AA100
  // - Range: A1:B10, A1:Z999
  // - Named range: MyRange (letters, numbers, underscores)
  const singleCellPattern = /^[A-Z]+[0-9]+$/i;
  const rangePattern = /^[A-Z]+[0-9]+:[A-Z]+[0-9]+$/i;
  const namedRangePattern = /^[a-zA-Z_][a-zA-Z0-9_]*$/;

  if (!singleCellPattern.test(address) &&
      !rangePattern.test(address) &&
      !namedRangePattern.test(address)) {
    throw new Error('Invalid Excel range address format');
  }

  // Additional validation: if it's a range, ensure start comes before end
  if (rangePattern.test(address)) {
    const [start, end] = address.split(':');
    // Basic check that both parts are valid
    if (!singleCellPattern.test(start) || !singleCellPattern.test(end)) {
      throw new Error('Invalid Excel range format');
    }
  }
}

/**
 * Validate file size limits
 */
export function validateFileSize(size: number, maxSize: number, resourceType: string): void {
  if (size > maxSize) {
    const maxSizeMB = (maxSize / (1024 * 1024)).toFixed(2);
    throw new Error(`${resourceType} size exceeds maximum allowed size of ${maxSizeMB} MB`);
  }
}

/**
 * Validate content length limits
 */
export function validateContentLength(content: string, maxLength: number, contentType: string): void {
  if (content.length > maxLength) {
    const maxLengthKB = (maxLength / 1024).toFixed(2);
    throw new Error(`${contentType} exceeds maximum size of ${maxLengthKB} KB`);
  }
}

/**
 * Get user-friendly error message from exception
 */
export function getPublicErrorMessage(error: Error): string {
  const errorMsg = error.message.toLowerCase();

  // Map known error patterns to user-friendly messages
  if (errorMsg.includes('401') || errorMsg.includes('unauthorized')) {
    return 'Authentication failed. Please check your credentials.';
  }
  if (errorMsg.includes('403') || errorMsg.includes('forbidden')) {
    return 'Access denied. You do not have permission for this operation.';
  }
  if (errorMsg.includes('404') || errorMsg.includes('not found')) {
    return 'The requested resource was not found.';
  }
  if (errorMsg.includes('429') || errorMsg.includes('too many requests')) {
    return 'Rate limit exceeded. Please try again later.';
  }
  if (errorMsg.includes('500') || errorMsg.includes('internal server')) {
    return 'An internal server error occurred. Please try again later.';
  }
  if (errorMsg.includes('network') || errorMsg.includes('timeout')) {
    return 'Network error. Please check your connection and try again.';
  }

  // For validation errors, return the actual message as it's safe
  if (errorMsg.includes('invalid') || errorMsg.includes('must be')) {
    return error.message;
  }

  // Generic message for unknown errors
  return 'An error occurred while processing your request.';
}

/**
 * Size limits constants
 */
export const SIZE_LIMITS = {
  MAX_FILE_SIZE: 100 * 1024 * 1024, // 100 MB
  MAX_EMAIL_BODY_SIZE: 1024 * 1024, // 1 MB
  MAX_EMAIL_SUBJECT_SIZE: 1024, // 1 KB
  MAX_EXCEL_CELLS: 10000,
  MAX_TEAMS_MESSAGE_SIZE: 28 * 1024, // 28 KB (Teams limit)
  MAX_ONENOTE_PAGE_SIZE: 5 * 1024 * 1024, // 5 MB
};
