import { z } from 'zod';
import { GraphClient, type GraphCalendar, type GraphCalendarEvent } from '../graph/client.js';
import { logger } from '../utils/logger.js';

// Safe pattern for Graph API resource IDs
const graphIdPattern = /^[a-zA-Z0-9\-._,!:]+$/;
const graphIdSchema = z.string().min(1).regex(graphIdPattern, 'Invalid resource ID format');

// Input schemas
export const listCalendarsInputSchema = z.object({});

export const listEventsInputSchema = z
  .object({
    calendar_id: graphIdSchema
      .optional()
      .describe('Calendar ID from cal_list_calendars. Omit to use default calendar.'),
    start_date: z
      .string()
      .datetime()
      .optional()
      .describe('Start of date range (ISO 8601). Must be used with end_date. Enables recurring event expansion.'),
    end_date: z
      .string()
      .datetime()
      .optional()
      .describe('End of date range (ISO 8601). Must be used with start_date.'),
    top: z
      .number()
      .int()
      .min(1)
      .max(100)
      .optional()
      .default(25)
      .describe('Maximum number of events to return (1-100)'),
  })
  .refine(
    (data) => {
      const hasStart = data.start_date !== undefined;
      const hasEnd = data.end_date !== undefined;
      return hasStart === hasEnd;
    },
    { message: 'start_date and end_date must both be provided or both omitted' }
  )
  .refine(
    (data) => {
      if (data.start_date && data.end_date) {
        return new Date(data.start_date) < new Date(data.end_date);
      }
      return true;
    },
    { message: 'start_date must be before end_date' }
  );

export const getEventInputSchema = z.object({
  event_id: graphIdSchema.describe('Event ID from a recent cal_list_events response'),
});

export type ListCalendarsInput = z.infer<typeof listCalendarsInputSchema>;
export type ListEventsInput = z.infer<typeof listEventsInputSchema>;
export type GetEventInput = z.infer<typeof getEventInputSchema>;

// Output formatters
function formatCalendar(calendar: GraphCalendar): object {
  return {
    id: calendar.id,
    name: calendar.name,
    color: calendar.color,
    isDefault: calendar.isDefaultCalendar,
    canEdit: calendar.canEdit,
    ownerName: calendar.owner?.name,
    ownerEmail: calendar.owner?.address,
  };
}

function formatEvent(event: GraphCalendarEvent, includeBody: boolean = false): object {
  const formatted: Record<string, unknown> = {
    id: event.id,
    subject: event.subject,
    preview: event.bodyPreview?.substring(0, 200),
    start: event.start?.dateTime,
    startTimeZone: event.start?.timeZone,
    end: event.end?.dateTime,
    endTimeZone: event.end?.timeZone,
    isAllDay: event.isAllDay,
    location: event.location?.displayName,
    organizer: event.organizer?.emailAddress?.address,
    organizerName: event.organizer?.emailAddress?.name,
    attendees: event.attendees?.map((a) => ({
      email: a.emailAddress?.address,
      name: a.emailAddress?.name,
      response: a.status?.response,
      type: a.type,
    })),
    showAs: event.showAs,
    importance: event.importance,
    webLink: event.webLink,
    joinUrl: event.onlineMeeting?.joinUrl,
    isRecurring: event.recurrence !== undefined && event.recurrence !== null,
  };

  if (includeBody && event.body) {
    formatted['body'] = {
      contentType: event.body.contentType,
      content: event.body.content,
    };
  }

  return formatted;
}

// Tool implementations
export class CalendarTools {
  private graphClient: GraphClient;

  constructor(graphClient: GraphClient) {
    this.graphClient = graphClient;
  }

  async listCalendars(): Promise<object> {
    logger.debug('Listing calendars');

    const calendars = await this.graphClient.listCalendars();

    return {
      calendars: calendars.map(formatCalendar),
      count: calendars.length,
    };
  }

  async listEvents(input: ListEventsInput): Promise<object> {
    const validated = listEventsInputSchema.parse(input);

    logger.debug(
      { calendarId: validated.calendar_id, startDate: validated.start_date, endDate: validated.end_date, top: validated.top },
      'Listing events'
    );

    if (validated.start_date && validated.end_date) {
      const events = await this.graphClient.listCalendarView({
        startDateTime: validated.start_date,
        endDateTime: validated.end_date,
        calendarId: validated.calendar_id,
        top: validated.top,
        orderBy: 'start/dateTime',
      });

      return {
        events: events.map((e) => formatEvent(e)),
        count: events.length,
        dateRange: { start: validated.start_date, end: validated.end_date },
        _note: 'Using calendarView: recurring events are expanded into individual occurrences within the date range.',
      };
    }

    const events = await this.graphClient.listEvents({
      calendarId: validated.calendar_id,
      top: validated.top,
      orderBy: 'start/dateTime desc',
    });

    return {
      events: events.map((e) => formatEvent(e)),
      count: events.length,
      _note: 'Without a date range, recurring events are NOT expanded. Provide start_date and end_date to see individual occurrences.',
    };
  }

  async getEvent(input: GetEventInput): Promise<object> {
    const validated = getEventInputSchema.parse(input);

    logger.debug({ eventId: validated.event_id }, 'Getting event');

    try {
      const event = await this.graphClient.getEvent(validated.event_id);
      return formatEvent(event, true);
    } catch (err) {
      const code = (err as { code?: string }).code;
      if (code === 'ErrorItemNotFound') {
        const hint =
          'Event not found. The event_id may be stale or belong to a different calendar. ' +
          'Use cal_list_events to get current event IDs.';

        const enrichedError = new Error(hint) as Error & { code?: string; statusCode?: number };
        enrichedError.code = code;
        enrichedError.statusCode = (err as { statusCode?: number }).statusCode;
        throw enrichedError;
      }
      throw err;
    }
  }
}

// Tool definitions for MCP registration
export const calendarToolDefinitions = [
  {
    name: 'cal_list_calendars',
    description:
      'List all calendars in the user\'s Microsoft 365 account with metadata (name, color, owner). Use this to discover calendar IDs for filtering events.',
    inputSchema: {
      type: 'object' as const,
      properties: {},
    },
  },
  {
    name: 'cal_list_events',
    description:
      'List calendar events. Provide start_date and end_date to expand recurring events into individual occurrences (uses calendarView). Without dates, returns recent events without recurring expansion. Use cal_get_event with an event ID from THIS response to get full details.',
    inputSchema: {
      type: 'object' as const,
      properties: {
        calendar_id: {
          type: 'string',
          description:
            'Calendar ID from cal_list_calendars. Omit to use the default calendar.',
        },
        start_date: {
          type: 'string',
          format: 'date-time',
          description:
            'Start of date range (ISO 8601). Must be used with end_date. Enables recurring event expansion.',
        },
        end_date: {
          type: 'string',
          format: 'date-time',
          description: 'End of date range (ISO 8601). Must be used with start_date.',
        },
        top: {
          type: 'number',
          description: 'Maximum number of events to return (1-100). Default: 25',
          minimum: 1,
          maximum: 100,
        },
      },
    },
  },
  {
    name: 'cal_get_event',
    description:
      'Get the full details of a specific calendar event by ID, including the full body/description. The event_id MUST be from a recent cal_list_events response.',
    inputSchema: {
      type: 'object' as const,
      properties: {
        event_id: {
          type: 'string',
          description: 'Event ID from a recent cal_list_events response.',
        },
      },
      required: ['event_id'],
    },
  },
];
