import { describe, it, expect, vi, beforeEach } from 'vitest';
import { CalendarTools, listEventsInputSchema, getEventInputSchema, listCalendarsInputSchema } from '../../src/tools/calendar.js';
import type { GraphClient, GraphCalendar, GraphCalendarEvent } from '../../src/graph/client.js';

describe('Calendar Tools', () => {
  let mockGraphClient: GraphClient;
  let calendarTools: CalendarTools;

  const mockCalendars: GraphCalendar[] = [
    {
      id: 'cal-1',
      name: 'Calendar',
      color: 'auto',
      isDefaultCalendar: true,
      canEdit: true,
      owner: { name: 'John Doe', address: 'john@example.com' },
    },
    {
      id: 'cal-2',
      name: 'Team Calendar',
      color: 'lightBlue',
      isDefaultCalendar: false,
      canEdit: false,
      owner: { name: 'Team', address: 'team@example.com' },
    },
  ];

  const mockEvents: GraphCalendarEvent[] = [
    {
      id: 'evt-1',
      subject: 'Team Standup',
      bodyPreview: 'Daily standup meeting for the team',
      start: { dateTime: '2026-02-16T09:00:00.0000000', timeZone: 'UTC' },
      end: { dateTime: '2026-02-16T09:30:00.0000000', timeZone: 'UTC' },
      location: { displayName: 'Conference Room A' },
      organizer: { emailAddress: { name: 'John Doe', address: 'john@example.com' } },
      attendees: [
        {
          emailAddress: { name: 'Jane Doe', address: 'jane@example.com' },
          status: { response: 'accepted' },
          type: 'required',
        },
      ],
      isAllDay: false,
      showAs: 'busy',
      importance: 'normal',
      webLink: 'https://outlook.office.com/calendar/item/evt-1',
      onlineMeeting: { joinUrl: 'https://teams.microsoft.com/meet/123' },
      recurrence: null,
    },
    {
      id: 'evt-2',
      subject: 'All Hands',
      bodyPreview: 'Monthly all-hands meeting',
      start: { dateTime: '2026-02-16T14:00:00.0000000', timeZone: 'UTC' },
      end: { dateTime: '2026-02-16T15:00:00.0000000', timeZone: 'UTC' },
      isAllDay: false,
      showAs: 'busy',
      importance: 'high',
      recurrence: { pattern: { type: 'monthly' } },
    },
  ];

  beforeEach(() => {
    mockGraphClient = {
      listCalendars: vi.fn().mockResolvedValue(mockCalendars),
      listEvents: vi.fn().mockResolvedValue(mockEvents),
      listCalendarView: vi.fn().mockResolvedValue(mockEvents),
      getEvent: vi.fn().mockResolvedValue({
        ...mockEvents[0],
        body: { contentType: 'html', content: '<p>Full description</p>' },
      }),
    } as unknown as GraphClient;

    calendarTools = new CalendarTools(mockGraphClient);
  });

  describe('listCalendars', () => {
    it('should list calendars', async () => {
      const result = await calendarTools.listCalendars() as Record<string, unknown>;

      expect(mockGraphClient.listCalendars).toHaveBeenCalled();
      expect(result['count']).toBe(2);
      expect(result['calendars']).toHaveLength(2);
    });

    it('should format calendar fields correctly', async () => {
      const result = await calendarTools.listCalendars() as Record<string, unknown>;
      const calendars = result['calendars'] as Record<string, unknown>[];

      expect(calendars[0]).toEqual({
        id: 'cal-1',
        name: 'Calendar',
        color: 'auto',
        isDefault: true,
        canEdit: true,
        ownerName: 'John Doe',
        ownerEmail: 'john@example.com',
      });
    });

    it('should handle empty calendar list', async () => {
      (mockGraphClient.listCalendars as ReturnType<typeof vi.fn>).mockResolvedValue([]);

      const result = await calendarTools.listCalendars() as Record<string, unknown>;

      expect(result['count']).toBe(0);
      expect(result['calendars']).toEqual([]);
    });
  });

  describe('listEvents', () => {
    it('should list events without date range using listEvents endpoint', async () => {
      const result = await calendarTools.listEvents({}) as Record<string, unknown>;

      expect(mockGraphClient.listEvents).toHaveBeenCalledWith({
        calendarId: undefined,
        top: 25,
        orderBy: 'start/dateTime desc',
      });
      expect(mockGraphClient.listCalendarView).not.toHaveBeenCalled();
      expect(result['count']).toBe(2);
      expect(result['_note']).toContain('NOT expanded');
    });

    it('should use calendarView when date range is provided', async () => {
      const result = await calendarTools.listEvents({
        start_date: '2026-02-16T00:00:00Z',
        end_date: '2026-02-17T00:00:00Z',
      }) as Record<string, unknown>;

      expect(mockGraphClient.listCalendarView).toHaveBeenCalledWith({
        startDateTime: '2026-02-16T00:00:00Z',
        endDateTime: '2026-02-17T00:00:00Z',
        calendarId: undefined,
        top: 25,
        orderBy: 'start/dateTime',
      });
      expect(mockGraphClient.listEvents).not.toHaveBeenCalled();
      expect(result['dateRange']).toEqual({
        start: '2026-02-16T00:00:00Z',
        end: '2026-02-17T00:00:00Z',
      });
      expect(result['_note']).toContain('recurring events are expanded');
    });

    it('should pass calendar_id to calendarView', async () => {
      await calendarTools.listEvents({
        calendar_id: 'cal-2',
        start_date: '2026-02-16T00:00:00Z',
        end_date: '2026-02-17T00:00:00Z',
      });

      expect(mockGraphClient.listCalendarView).toHaveBeenCalledWith(
        expect.objectContaining({ calendarId: 'cal-2' })
      );
    });

    it('should pass calendar_id to listEvents', async () => {
      await calendarTools.listEvents({ calendar_id: 'cal-2' });

      expect(mockGraphClient.listEvents).toHaveBeenCalledWith(
        expect.objectContaining({ calendarId: 'cal-2' })
      );
    });

    it('should respect top parameter', async () => {
      await calendarTools.listEvents({ top: 10 });

      expect(mockGraphClient.listEvents).toHaveBeenCalledWith(
        expect.objectContaining({ top: 10 })
      );
    });

    it('should format event fields correctly', async () => {
      const result = await calendarTools.listEvents({}) as Record<string, unknown>;
      const events = result['events'] as Record<string, unknown>[];

      expect(events[0]).toMatchObject({
        id: 'evt-1',
        subject: 'Team Standup',
        start: '2026-02-16T09:00:00.0000000',
        startTimeZone: 'UTC',
        end: '2026-02-16T09:30:00.0000000',
        endTimeZone: 'UTC',
        isAllDay: false,
        location: 'Conference Room A',
        organizer: 'john@example.com',
        organizerName: 'John Doe',
        showAs: 'busy',
        importance: 'normal',
        joinUrl: 'https://teams.microsoft.com/meet/123',
        isRecurring: false,
      });

      expect(events[0]).toHaveProperty('attendees');
      const attendees = (events[0] as Record<string, unknown>)['attendees'] as Record<string, unknown>[];
      expect(attendees[0]).toEqual({
        email: 'jane@example.com',
        name: 'Jane Doe',
        response: 'accepted',
        type: 'required',
      });
    });

    it('should detect recurring events', async () => {
      const result = await calendarTools.listEvents({}) as Record<string, unknown>;
      const events = result['events'] as Record<string, unknown>[];

      expect(events[0]['isRecurring']).toBe(false);
      expect(events[1]['isRecurring']).toBe(true);
    });

    it('should truncate long previews', async () => {
      const longPreview = 'a'.repeat(300);
      (mockGraphClient.listEvents as ReturnType<typeof vi.fn>).mockResolvedValue([
        { ...mockEvents[0], bodyPreview: longPreview },
      ]);

      const result = await calendarTools.listEvents({}) as Record<string, unknown>;
      const events = result['events'] as Record<string, unknown>[];

      expect((events[0]['preview'] as string).length).toBe(200);
    });
  });

  describe('getEvent', () => {
    it('should get event with body', async () => {
      const result = await calendarTools.getEvent({ event_id: 'evt-1' }) as Record<string, unknown>;

      expect(mockGraphClient.getEvent).toHaveBeenCalledWith('evt-1');
      expect(result).toHaveProperty('body');
      expect((result['body'] as Record<string, unknown>)['content']).toBe('<p>Full description</p>');
    });

    it('should throw enriched error for ErrorItemNotFound', async () => {
      const graphError = Object.assign(new Error('Item not found'), {
        code: 'ErrorItemNotFound',
        statusCode: 404,
      });
      (mockGraphClient.getEvent as ReturnType<typeof vi.fn>).mockRejectedValue(graphError);

      await expect(
        calendarTools.getEvent({ event_id: 'evt-nonexistent' })
      ).rejects.toThrow(/Use cal_list_events/);
    });

    it('should preserve error code on enriched error', async () => {
      const graphError = Object.assign(new Error('Item not found'), {
        code: 'ErrorItemNotFound',
        statusCode: 404,
      });
      (mockGraphClient.getEvent as ReturnType<typeof vi.fn>).mockRejectedValue(graphError);

      try {
        await calendarTools.getEvent({ event_id: 'evt-nonexistent' });
        expect.unreachable('Should have thrown');
      } catch (err) {
        expect((err as { code?: string }).code).toBe('ErrorItemNotFound');
        expect((err as { statusCode?: number }).statusCode).toBe(404);
      }
    });

    it('should re-throw non-ErrorItemNotFound errors unchanged', async () => {
      const graphError = Object.assign(new Error('Server error'), {
        code: 'InternalServerError',
        statusCode: 500,
      });
      (mockGraphClient.getEvent as ReturnType<typeof vi.fn>).mockRejectedValue(graphError);

      await expect(
        calendarTools.getEvent({ event_id: 'evt-1' })
      ).rejects.toThrow('Server error');
    });
  });
});

describe('Calendar Input Schemas', () => {
  describe('listCalendarsInputSchema', () => {
    it('should accept empty object', () => {
      const result = listCalendarsInputSchema.safeParse({});
      expect(result.success).toBe(true);
    });
  });

  describe('listEventsInputSchema', () => {
    it('should accept empty object', () => {
      const result = listEventsInputSchema.safeParse({});
      expect(result.success).toBe(true);
    });

    it('should accept valid date range', () => {
      const result = listEventsInputSchema.safeParse({
        start_date: '2026-02-16T00:00:00Z',
        end_date: '2026-02-17T00:00:00Z',
      });
      expect(result.success).toBe(true);
    });

    it('should reject start_date without end_date', () => {
      const result = listEventsInputSchema.safeParse({
        start_date: '2026-02-16T00:00:00Z',
      });
      expect(result.success).toBe(false);
    });

    it('should reject end_date without start_date', () => {
      const result = listEventsInputSchema.safeParse({
        end_date: '2026-02-17T00:00:00Z',
      });
      expect(result.success).toBe(false);
    });

    it('should reject start_date >= end_date', () => {
      const result = listEventsInputSchema.safeParse({
        start_date: '2026-02-17T00:00:00Z',
        end_date: '2026-02-16T00:00:00Z',
      });
      expect(result.success).toBe(false);
    });

    it('should reject equal start and end dates', () => {
      const result = listEventsInputSchema.safeParse({
        start_date: '2026-02-16T00:00:00Z',
        end_date: '2026-02-16T00:00:00Z',
      });
      expect(result.success).toBe(false);
    });

    it('should accept valid top value', () => {
      const result = listEventsInputSchema.safeParse({ top: 50 });
      expect(result.success).toBe(true);
    });

    it('should reject top > 100', () => {
      const result = listEventsInputSchema.safeParse({ top: 200 });
      expect(result.success).toBe(false);
    });

    it('should reject top < 1', () => {
      const result = listEventsInputSchema.safeParse({ top: 0 });
      expect(result.success).toBe(false);
    });

    it('should reject invalid datetime format', () => {
      const result = listEventsInputSchema.safeParse({
        start_date: 'not-a-date',
        end_date: '2026-02-17T00:00:00Z',
      });
      expect(result.success).toBe(false);
    });

    it('should accept valid calendar_id', () => {
      const result = listEventsInputSchema.safeParse({
        calendar_id: 'AAMkAGI2TG93AAA',
      });
      expect(result.success).toBe(true);
    });

    it('should reject calendar_id with invalid characters', () => {
      const result = listEventsInputSchema.safeParse({
        calendar_id: 'id with spaces',
      });
      expect(result.success).toBe(false);
    });
  });

  describe('getEventInputSchema', () => {
    it('should accept valid event_id', () => {
      const result = getEventInputSchema.safeParse({
        event_id: 'AAMkAGI2TG93AAA',
      });
      expect(result.success).toBe(true);
    });

    it('should require event_id', () => {
      const result = getEventInputSchema.safeParse({});
      expect(result.success).toBe(false);
    });

    it('should reject empty event_id', () => {
      const result = getEventInputSchema.safeParse({ event_id: '' });
      expect(result.success).toBe(false);
    });

    it('should reject event_id with invalid characters', () => {
      const result = getEventInputSchema.safeParse({
        event_id: 'id with spaces',
      });
      expect(result.success).toBe(false);
    });
  });
});
