import { describe, expect, it } from 'vitest';
import {
  isAppOnlyCalendarEnabled,
  isCalendarToolPath,
  rewriteCalendarPath,
} from '../src/lib/app-only-calendar.js';

describe('app-only calendar helpers', () => {
  it('requires both the flag and the client secret', () => {
    expect(isAppOnlyCalendarEnabled({})).toBe(false);
    expect(isAppOnlyCalendarEnabled({ MS365_MCP_CALENDAR_APP_ONLY: '1' })).toBe(false);
    expect(
      isAppOnlyCalendarEnabled({
        MS365_MCP_CALENDAR_APP_ONLY: '1',
        MS365_MCP_APP_CLIENT_SECRET: 'secret',
      })
    ).toBe(true);
  });

  it('detects calendar paths and leaves other tools alone', () => {
    for (const path of [
      '/me/events',
      '/me/events/{event-id}/accept',
      '/me/calendar/getSchedule',
      '/me/findMeetingTimes',
      '/me/calendars/{calendar-id}/events',
      '/me/calendarView',
      '/users/{user-id}/calendarView',
    ]) {
      expect(isCalendarToolPath(path), path).toBe(true);
    }
    for (const path of ['/me/messages', '/me/todo/lists', '/drives/{drive-id}/root/children']) {
      expect(isCalendarToolPath(path), path).toBe(false);
    }
  });

  it('rewrites /me calendar paths to the organizer mailbox', () => {
    expect(rewriteCalendarPath('/me/events', 'sofus@truecollective.dk')).toBe(
      '/users/sofus%40truecollective.dk/events'
    );
    expect(rewriteCalendarPath('/me/findMeetingTimes', 'a b@x.dk')).toBe(
      '/users/a%20b%40x.dk/findMeetingTimes'
    );
    expect(rewriteCalendarPath('/users/{user-id}/calendarView', 'x@y.dk')).toBeUndefined();
  });
});
