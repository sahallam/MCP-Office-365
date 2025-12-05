/**
 * Calendar tools for MCP server
 */

import { Client } from '@microsoft/microsoft-graph-client';
import { CalendarEvent } from '../types.js';
import { validateResourceId } from '../security.js';
import { detectSystemTimezone, convertEventToLocalTimezone, convertEventsToLocalTimezone } from '../utils/timezone.js';

export class CalendarTools {
  private localTimezone: string;

  constructor(private graphClient: Client, private userId: string) {
    this.localTimezone = detectSystemTimezone();
    console.error(`[Calendar] Detected system timezone: ${this.localTimezone}`);
  }

  /**
   * Get the correct user path for API endpoints
   * Returns '/me' if userId is 'me', otherwise '/users/{userId}'
   */
  private getUserPath(): string {
    return this.userId === 'me' ? '/me' : `/users/${this.userId}`;
  }

  /**
   * List calendar events
   */
  async listEvents(options: {
    startDateTime?: string;
    endDateTime?: string;
    top?: number;
    filter?: string;
  } = {}): Promise<CalendarEvent[]> {
    const { top = 10, filter } = options;

    let query = this.graphClient
      .api(`${this.getUserPath()}/calendar/events`)
      .top(top)
      .select(['id', 'subject', 'start', 'end', 'location', 'attendees', 'isOnlineMeeting', 'onlineMeetingUrl'])
      .orderby('start/dateTime');

    if (filter) {
      query = query.filter(filter);
    }

    const result = await query.get();
    return convertEventsToLocalTimezone(result.value, this.localTimezone);
  }

  /**
   * Get calendar view for a specific time range
   */
  async getCalendarView(startDateTime: string, endDateTime: string): Promise<CalendarEvent[]> {
    const result = await this.graphClient
      .api(`${this.getUserPath()}/calendar/calendarView`)
      .query({
        startDateTime,
        endDateTime,
      })
      .select(['id', 'subject', 'start', 'end', 'location', 'attendees', 'isOnlineMeeting', 'onlineMeetingUrl'])
      .orderby('start/dateTime')
      .get();

    return convertEventsToLocalTimezone(result.value, this.localTimezone);
  }

  /**
   * Get a specific event by ID
   */
  async getEvent(eventId: string): Promise<CalendarEvent> {
    validateResourceId(eventId, 'event');

    const event = await this.graphClient
      .api(`${this.getUserPath()}/calendar/events/${eventId}`)
      .get();

    return convertEventToLocalTimezone(event, this.localTimezone);
  }

  /**
   * Create a calendar event
   */
  async createEvent(event: CalendarEvent): Promise<CalendarEvent> {
    const eventObject = {
      subject: event.subject,
      body: event.body,
      start: event.start,
      end: event.end,
      location: event.location,
      attendees: event.attendees,
      isOnlineMeeting: event.isOnlineMeeting || false,
    };

    const createdEvent = await this.graphClient
      .api(`${this.getUserPath()}/calendar/events`)
      .post(eventObject);

    return convertEventToLocalTimezone(createdEvent, this.localTimezone);
  }

  /**
   * Update a calendar event
   */
  async updateEvent(eventId: string, updates: Partial<CalendarEvent>): Promise<CalendarEvent> {
    validateResourceId(eventId, 'event');

    const updatedEvent = await this.graphClient
      .api(`${this.getUserPath()}/calendar/events/${eventId}`)
      .patch(updates);

    return convertEventToLocalTimezone(updatedEvent, this.localTimezone);
  }

  /**
   * Delete a calendar event
   */
  async deleteEvent(eventId: string): Promise<void> {
    validateResourceId(eventId, 'event');

    await this.graphClient
      .api(`${this.getUserPath()}/calendar/events/${eventId}`)
      .delete();
  }

  /**
   * Accept a meeting
   */
  async acceptMeeting(eventId: string, comment?: string): Promise<void> {
    validateResourceId(eventId, 'event');

    await this.graphClient
      .api(`${this.getUserPath()}/events/${eventId}/accept`)
      .post({
        comment: comment || '',
        sendResponse: true,
      });
  }

  /**
   * Decline a meeting
   */
  async declineMeeting(eventId: string, comment?: string): Promise<void> {
    validateResourceId(eventId, 'event');

    await this.graphClient
      .api(`${this.getUserPath()}/events/${eventId}/decline`)
      .post({
        comment: comment || '',
        sendResponse: true,
      });
  }

  /**
   * Tentatively accept a meeting
   */
  async tentativelyAcceptMeeting(eventId: string, comment?: string): Promise<void> {
    validateResourceId(eventId, 'event');

    await this.graphClient
      .api(`${this.getUserPath()}/events/${eventId}/tentativelyAccept`)
      .post({
        comment: comment || '',
        sendResponse: true,
      });
  }

  /**
   * Find meeting times
   */
  async findMeetingTimes(options: {
    attendees: string[];
    timeConstraint: {
      timeslots: Array<{ start: { dateTime: string; timeZone: string }; end: { dateTime: string; timeZone: string } }>;
    };
    meetingDuration: string; // ISO 8601 duration format (e.g., 'PT1H' for 1 hour)
    maxCandidates?: number;
  }): Promise<any> {
    const requestBody = {
      attendees: options.attendees.map(email => ({
        emailAddress: { address: email },
        type: 'required',
      })),
      timeConstraint: options.timeConstraint,
      meetingDuration: options.meetingDuration,
      maxCandidates: options.maxCandidates || 5,
    };

    const result = await this.graphClient
      .api(`${this.getUserPath()}/findMeetingTimes`)
      .post(requestBody);

    return result;
  }

  /**
   * List calendars
   */
  async listCalendars(): Promise<any[]> {
    const result = await this.graphClient
      .api(`${this.getUserPath()}/calendars`)
      .get();

    return result.value;
  }
}
