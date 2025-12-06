/**
 * Timezone utilities for converting UTC times to local timezone
 */

import { execSync } from 'child_process';

/**
 * Detect the system timezone using various methods depending on the platform
 */
export function detectSystemTimezone(): string {
  try {
    // Try to get timezone from environment variable (works on most systems)
    if (process.env.TZ) {
      return process.env.TZ;
    }

    // For macOS/Linux: Try reading from /etc/localtime symlink
    if (process.platform === 'darwin' || process.platform === 'linux') {
      try {
        const timezone = execSync('readlink /etc/localtime', { encoding: 'utf-8' }).trim();
        // Extract timezone name from path like /var/db/timezone/zoneinfo/America/New_York
        const match = timezone.match(/zoneinfo\/(.+)$/);
        if (match) {
          return match[1];
        }
      } catch {
        // Fall through to next method
      }
    }

    // For macOS: Use systemsetup command
    if (process.platform === 'darwin') {
      try {
        const result = execSync('systemsetup -gettimezone', { encoding: 'utf-8' });
        const match = result.match(/Time Zone: (.+)/);
        if (match) {
          return match[1].trim();
        }
      } catch {
        // Fall through to next method
      }
    }

    // For Linux: Try reading /etc/timezone file
    if (process.platform === 'linux') {
      try {
        const timezone = execSync('cat /etc/timezone', { encoding: 'utf-8' }).trim();
        if (timezone) {
          return timezone;
        }
      } catch {
        // Fall through to next method
      }
    }

    // Last resort: Use JavaScript's Intl API
    const intlTimezone = Intl.DateTimeFormat().resolvedOptions().timeZone;
    if (intlTimezone) {
      return intlTimezone;
    }

    // Default to UTC if all else fails
    console.warn('[Timezone] Could not detect system timezone, defaulting to UTC');
    return 'UTC';
  } catch (error) {
    console.error('[Timezone] Error detecting system timezone:', error);
    return 'UTC';
  }
}

/**
 * Convert a datetime string from one timezone to another
 * @param dateTime ISO 8601 datetime string (without timezone suffix)
 * @param fromTimeZone Source timezone (typically 'UTC')
 * @param toTimeZone Destination timezone
 * @returns ISO 8601 datetime string in the destination timezone
 */
export function convertTimezone(dateTime: string, fromTimeZone: string, toTimeZone: string): string {
  try {
    // Microsoft Graph returns datetime strings without timezone info (e.g., "2025-12-05T04:00:00")
    // We need to interpret this datetime as being in the source timezone

    let dateToConvert: Date;

    // If the source timezone is UTC, append 'Z' to parse as UTC
    // Otherwise, we need to parse it differently
    if (fromTimeZone === 'UTC') {
      // Ensure the datetime string is treated as UTC by adding 'Z' if not present
      const utcDateString = dateTime.endsWith('Z') ? dateTime : dateTime + 'Z';
      dateToConvert = new Date(utcDateString);
    } else {
      // For non-UTC source timezones, we need to be more careful
      // Parse the datetime string and interpret it as being in the source timezone
      // This is tricky because JavaScript Date doesn't support parsing in arbitrary timezones
      // For now, we'll handle the common case (UTC) and fall back to direct parsing
      dateToConvert = new Date(dateTime);
    }

    // Create a formatter for the target timezone
    const formatter = new Intl.DateTimeFormat('en-US', {
      timeZone: toTimeZone,
      year: 'numeric',
      month: '2-digit',
      day: '2-digit',
      hour: '2-digit',
      minute: '2-digit',
      second: '2-digit',
      hour12: false,
    });

    // Format the date in the target timezone
    const parts = formatter.formatToParts(dateToConvert);
    const values: any = {};
    parts.forEach(part => {
      if (part.type !== 'literal') {
        values[part.type] = part.value;
      }
    });

    // Construct ISO 8601 datetime string (without timezone offset)
    const converted = `${values.year}-${values.month}-${values.day}T${values.hour}:${values.minute}:${values.second}`;
    return converted;
  } catch (error) {
    console.error('[Timezone] Error converting timezone:', error);
    return dateTime;
  }
}

/**
 * Convert a calendar event's datetime fields to local timezone
 */
export function convertEventToLocalTimezone(event: any, localTimezone: string): any {
  if (!event) return event;

  const convertedEvent = { ...event };

  // Convert start time
  if (event.start?.dateTime && event.start?.timeZone) {
    convertedEvent.start = {
      dateTime: convertTimezone(event.start.dateTime, event.start.timeZone, localTimezone),
      timeZone: localTimezone,
    };
  }

  // Convert end time
  if (event.end?.dateTime && event.end?.timeZone) {
    convertedEvent.end = {
      dateTime: convertTimezone(event.end.dateTime, event.end.timeZone, localTimezone),
      timeZone: localTimezone,
    };
  }

  return convertedEvent;
}

/**
 * Convert an array of calendar events to local timezone
 */
export function convertEventsToLocalTimezone(events: any[], localTimezone: string): any[] {
  return events.map(event => convertEventToLocalTimezone(event, localTimezone));
}
