import { test } from 'node:test';
import assert from 'node:assert/strict';
import { fakeCalendar, loadCalSync, formatDate } from './fakes.mjs';

const WORK = 'primary';
const PERSONAL = 'me@personal.example';
const pdt = s => `${s}:00-07:00`; // Los Angeles wall-clock time in September
const at = s => Date.parse(pdt(s));
const tagOf = ev => ev.extendedProperties?.private?.workHoldId ?? null;

/** A work calendar plus a personal calendar, with CalSync loaded and pointed at them. */
function setup({ now = '2026-09-02T12:00', lockAvailable = true, readOnly = false } = {}) {
  const cal = fakeCalendar();
  cal.addCalendar(WORK);
  cal.addCalendar(PERSONAL, { readOnly });
  const clock = { now: at(now) };
  const cs = loadCalSync({ calendar: cal, now: () => clock.now, lockAvailable });
  cs.CONFIG.personalCalendarIds = [PERSONAL];
  const personal = (start, end, extra = {}) =>
    cal.addEvent(PERSONAL, { start: { dateTime: pdt(start) }, end: { dateTime: pdt(end) }, ...extra });
  const holds = () => cal.live(WORK)
    .filter(ev => ev.summary === cs.HOLD_TITLE)
    .sort((a, b) => Date.parse(a.start.dateTime) - Date.parse(b.start.dateTime));
  const times = () => holds().map(h => [h.start.dateTime, h.end.dateTime].map(t => new Date(t).toISOString()));
  return { cal, cs, clock, personal, holds, times };
}

test('creates a private, opaque hold clamped to work hours and tags the personal event with it', () => {
  const { cs, personal, holds, times } = setup();
  const ev = personal('2026-09-03T08:30', '2026-09-03T10:15');
  cs.sync();

  assert.deepEqual(times(), [['2026-09-03T16:00:00.000Z', '2026-09-03T17:15:00.000Z']]);
  const [hold] = holds();
  assert.equal(hold.visibility, 'private');
  assert.equal(hold.transparency, 'opaque');
  assert.deepEqual(hold.reminders, { useDefault: false, overrides: [] });
  assert.deepEqual(hold.extendedProperties.private, { sourceEventId: ev.id, sourceCalendarId: PERSONAL });
  assert.equal(tagOf(ev), hold.id);
  assert.deepEqual(cs.logs.log, ['CalSync: 1 created, 0 updated, 0 removed']);
});

test('skips events that do not block work time, and does not tag them', () => {
  const { cs, cal, personal, holds } = setup();
  const skipped = [
    cal.addEvent(PERSONAL, { start: { date: '2026-09-03' }, end: { date: '2026-09-04' } }), // all-day
    personal('2026-09-03T10:00', '2026-09-03T11:00', { transparency: 'transparent' }), // marked free
    personal('2026-09-03T10:00', '2026-09-03T11:00', { status: 'cancelled' }),
    personal('2026-09-03T10:00', '2026-09-03T11:00', { attendees: [{ email: PERSONAL, self: true, responseStatus: 'declined' }] }),
    personal('2026-09-05T10:00', '2026-09-05T11:00'), // Saturday
    personal('2026-09-03T18:00', '2026-09-03T19:00'), // after work
    personal('2026-09-03T07:00', '2026-09-03T09:00'), // ends as work starts
    personal('2026-09-03T09:00', '2026-09-03T17:00'), // fills the whole workday
  ];
  cs.sync();
  assert.equal(holds().length, 0);
  assert.deepEqual(skipped.map(tagOf), skipped.map(() => null));
  assert.deepEqual(cal.writes(), []);
});

test('leaves the work calendar\'s own events alone', () => {
  const { cal, cs, personal, holds } = setup();
  const meeting = cal.addEvent(WORK, { summary: 'Standup', start: { dateTime: pdt('2026-09-03T10:00') }, end: { dateTime: pdt('2026-09-03T10:30') } });
  const series = cal.addEvent(WORK, { summary: 'Weekly', recurrence: ['RRULE:FREQ=WEEKLY'], start: { dateTime: pdt('2026-09-03T14:00') }, end: { dateTime: pdt('2026-09-03T15:00') } });
  personal('2026-09-03T10:00', '2026-09-03T11:00');
  cs.sync();
  cs.uninstall();
  assert.equal(holds().length, 0);
  assert.deepEqual(cal.live(WORK).map(ev => ev.id), [meeting.id, series.id]);
  assert.deepEqual(cs.logs.error, []);
});

test('still holds time for invitations the owner accepted or has not answered', () => {
  const { cs, personal, holds } = setup();
  personal('2026-09-03T10:00', '2026-09-03T11:00', { attendees: [{ email: PERSONAL, self: true, responseStatus: 'accepted' }] });
  personal('2026-09-03T12:00', '2026-09-03T13:00', { attendees: [{ email: PERSONAL, self: true, responseStatus: 'needsAction' }] });
  cs.sync();
  assert.equal(holds().length, 2);
});

test('is idempotent: a second run performs no writes', () => {
  const { cal, cs, personal } = setup();
  personal('2026-09-03T10:00', '2026-09-03T11:00');
  cs.sync();
  cal.calls.length = 0;
  cs.sync();
  assert.deepEqual(cal.writes(), []);
  assert.equal(cs.logs.log.at(-1), 'CalSync: 0 created, 0 updated, 0 removed');
});

test('moves the hold when the personal event moves, and keeps the tag pointing at it', () => {
  const { cs, personal, holds, times } = setup();
  const ev = personal('2026-09-03T10:00', '2026-09-03T11:00');
  cs.sync();
  const holdId = holds()[0].id;

  ev.start.dateTime = pdt('2026-09-03T13:00'); // later the same day: the hold is patched in place
  ev.end.dateTime = pdt('2026-09-03T19:00');
  cs.sync();
  assert.deepEqual(times(), [['2026-09-03T20:00:00.000Z', '2026-09-04T00:00:00.000Z']]);
  assert.equal(holds()[0].id, holdId);
  assert.equal(tagOf(ev), holdId);
  assert.equal(cs.logs.log.at(-1), 'CalSync: 0 created, 1 updated, 0 removed');

  ev.start.dateTime = pdt('2026-09-04T10:00'); // another day: the old hold goes, a new one comes
  ev.end.dateTime = pdt('2026-09-04T11:00');
  cs.sync();
  assert.deepEqual(times(), [['2026-09-04T17:00:00.000Z', '2026-09-04T18:00:00.000Z']]);
  assert.notEqual(holds()[0].id, holdId);
  assert.equal(tagOf(ev), holds()[0].id);
  assert.equal(cs.logs.log.at(-1), 'CalSync: 1 created, 0 updated, 1 removed');
});

test('removes the hold and the tag when the event is deleted, marked free, or covered by Out of Office', () => {
  const { cal, cs, personal, holds } = setup();
  const deleted = personal('2026-09-03T10:00', '2026-09-03T11:00');
  const freed = personal('2026-09-04T10:00', '2026-09-04T11:00');
  const covered = personal('2026-09-07T10:00', '2026-09-07T11:00');
  cs.sync();
  assert.equal(holds().length, 3);

  deleted.status = 'cancelled';
  freed.transparency = 'transparent';
  cal.addEvent(WORK, { eventType: 'outOfOffice', start: { date: '2026-09-07' }, end: { date: '2026-09-08' } });
  cs.sync();
  assert.equal(holds().length, 0);
  assert.equal(tagOf(freed), null);
  assert.equal(tagOf(covered), null);
  assert.equal(cs.logs.log.at(-1), 'CalSync: 0 created, 0 updated, 3 removed');
  assert.deepEqual(cs.logs.error, []); // clearing the tag of the deleted event fails silently
});

test('clears the tag of an event that moved out of the sync window', () => {
  const { cs, personal, holds } = setup();
  const ev = personal('2026-09-03T10:00', '2026-09-03T11:00');
  cs.sync();
  assert.notEqual(tagOf(ev), null);

  ev.start.dateTime = pdt('2026-12-03T10:00');
  ev.end.dateTime = pdt('2026-12-03T11:00');
  cs.sync();
  assert.equal(holds().length, 0);
  assert.equal(tagOf(ev), null);
});

test('points the tag at the earliest hold still inside the window', () => {
  const { cs, clock, personal, holds } = setup({ now: '2026-09-04T09:00' });
  const ev = personal('2026-09-04T14:00', '2026-09-07T12:00'); // Friday afternoon to Monday noon
  cs.sync();
  const [friday, monday] = holds();
  assert.equal(tagOf(ev), friday.id);

  clock.now = at('2026-09-05T18:00'); // the Friday hold is now behind the window
  cs.sync();
  assert.equal(tagOf(ev), monday.id);
  assert.equal(holds().length, 2);
});

test('re-creates a hold that was deleted by hand on the work calendar', () => {
  const { cal, cs, personal, holds } = setup();
  const ev = personal('2026-09-03T10:00', '2026-09-03T11:00');
  cs.sync();
  cal.api.Events.remove(WORK, holds()[0].id);
  cs.sync();
  assert.equal(holds().length, 1);
  assert.equal(tagOf(ev), holds()[0].id);
});

test('applies a changed holdVisibility to existing holds', () => {
  const { cs, personal, holds } = setup();
  personal('2026-09-03T10:00', '2026-09-03T11:00');
  cs.sync();
  cs.CONFIG.holdVisibility = 'default';
  cs.sync();
  assert.equal(holds()[0].visibility, 'default');
  assert.equal(cs.logs.log.at(-1), 'CalSync: 0 created, 1 updated, 0 removed');

  delete holds()[0].visibility; // the API leaves the field out for default visibility
  cs.sync();
  assert.equal(cs.logs.log.at(-1), 'CalSync: 0 created, 0 updated, 0 removed');
});

test('holds each weekday of a multi-day event, leaving fully covered workdays to Out of Office', () => {
  const { cs, personal, holds, times } = setup();
  const trip = personal('2026-09-07T15:00', '2026-09-09T11:00'); // Monday afternoon to Wednesday morning
  personal('2026-09-13T22:00', '2026-09-14T10:00'); // Sunday night to Monday morning
  cs.sync();
  assert.deepEqual(times(), [
    ['2026-09-07T22:00:00.000Z', '2026-09-08T00:00:00.000Z'], // Mon 15:00–17:00
    ['2026-09-09T16:00:00.000Z', '2026-09-09T18:00:00.000Z'], // Wed 09:00–11:00; Tuesday would be a full workday
    ['2026-09-14T16:00:00.000Z', '2026-09-14T17:00:00.000Z'], // Mon 09:00–10:00
  ]);
  assert.equal(tagOf(trip), holds()[0].id); // the tag names the first hold
});

test('never duplicates the holds of an event that keeps overlapping the window', () => {
  const { cs, clock, personal, holds } = setup({ now: '2026-09-04T09:00' });
  personal('2026-09-04T14:00', '2026-09-07T12:00'); // Friday afternoon to Monday noon
  cs.sync();
  assert.equal(holds().length, 2);

  for (const later of ['2026-09-05T18:00', '2026-09-05T18:05', '2026-09-06T10:00', '2026-09-07T09:00', '2026-09-08T09:00']) {
    clock.now = at(later);
    cs.sync();
    assert.equal(holds().length, 2, `hold count changed at ${later}`);
  }
});

test('waits for a hold to enter the window instead of duplicating it at the far edge', () => {
  const { cs, clock, personal, holds } = setup({ now: '2026-09-02T07:00' });
  personal('2026-09-30T06:30', '2026-09-30T09:30'); // starts inside the window, its hold would start outside
  for (const t of ['07:00', '07:05', '08:55', '09:00']) { // at 09:00 the hold would start exactly at the (exclusive) window end
    clock.now = at(`2026-09-02T${t}`);
    cs.sync();
    assert.equal(holds().length, 0, `hold created too early at ${t}`);
  }
  for (const t of ['09:05', '09:10']) {
    clock.now = at(`2026-09-02T${t}`);
    cs.sync();
    assert.equal(holds().length, 1, `wrong hold count at ${t}`);
  }
});

test('cleans up duplicate holds left behind by earlier versions', () => {
  const { cal, cs, personal, holds } = setup();
  const ev = personal('2026-09-03T10:00', '2026-09-03T11:00');
  for (let i = 0; i < 3; i++) {
    cal.addEvent(WORK, {
      summary: cs.HOLD_TITLE, visibility: 'private',
      start: { dateTime: pdt('2026-09-03T10:00') }, end: { dateTime: pdt('2026-09-03T11:00') },
      extendedProperties: { private: { sourceEventId: ev.id, sourceCalendarId: PERSONAL } },
    });
  }
  cs.sync();
  assert.equal(holds().length, 1);
  assert.equal(cs.logs.log.at(-1), 'CalSync: 0 created, 0 updated, 2 removed');
});

test('replaces a hold that was edited into an all-day event', () => {
  const { cal, cs, personal, holds, times } = setup();
  const ev = personal('2026-09-03T10:00', '2026-09-03T11:00');
  cal.addEvent(WORK, {
    summary: cs.HOLD_TITLE, visibility: 'private', start: { date: '2026-09-03' }, end: { date: '2026-09-04' },
    extendedProperties: { private: { sourceEventId: ev.id, sourceCalendarId: PERSONAL } },
  });
  cs.sync();
  assert.deepEqual(times(), [['2026-09-03T17:00:00.000Z', '2026-09-03T18:00:00.000Z']]);
  assert.equal(holds().length, 1);
  assert.equal(cs.logs.log.at(-1), 'CalSync: 1 created, 0 updated, 1 removed');
});

test('skips holds fully covered by Out of Office, timed or all-day, but keeps partially covered ones', () => {
  const { cal, cs, personal, holds } = setup();
  cal.addEvent(WORK, { eventType: 'outOfOffice', start: { dateTime: pdt('2026-09-03T09:00') }, end: { dateTime: pdt('2026-09-03T12:00') } });
  cal.addEvent(WORK, { eventType: 'outOfOffice', start: { date: '2026-09-04' }, end: { date: '2026-09-05' } });
  personal('2026-09-03T10:00', '2026-09-03T11:00'); // inside the timed OOO
  personal('2026-09-04T10:00', '2026-09-04T11:00'); // inside the all-day OOO
  const partial = personal('2026-09-03T11:30', '2026-09-03T13:00'); // straddles the end of the timed OOO
  cs.sync();

  assert.equal(holds().length, 1);
  assert.equal(holds()[0].extendedProperties.private.sourceEventId, partial.id);
});

test('creates a single hold when the same invitation is on two personal calendars', () => {
  const { cal, cs, holds } = setup();
  const other = 'partner@personal.example';
  cal.addCalendar(other);
  cs.CONFIG.personalCalendarIds = [PERSONAL, other];
  const copies = [PERSONAL, other].map(calId =>
    cal.addEvent(calId, { id: 'shared-invite', start: { dateTime: pdt('2026-09-03T10:00') }, end: { dateTime: pdt('2026-09-03T11:00') } }));
  cs.sync();
  assert.equal(holds().length, 1);
  assert.deepEqual(copies.map(tagOf), [holds()[0].id, holds()[0].id]);
  assert.equal(cs.logs.log.at(-1), 'CalSync: 1 created, 0 updated, 0 removed');
});

test('removes the holds of a calendar taken out of CONFIG', () => {
  const { cal, cs, personal, holds } = setup();
  const other = 'old@personal.example';
  cal.addCalendar(other);
  cs.CONFIG.personalCalendarIds = [PERSONAL, other];
  personal('2026-09-03T10:00', '2026-09-03T11:00');
  const ev = cal.addEvent(other, { start: { dateTime: pdt('2026-09-03T14:00') }, end: { dateTime: pdt('2026-09-03T15:00') } });
  cs.sync();
  assert.equal(holds().length, 2);

  cs.CONFIG.personalCalendarIds = [PERSONAL];
  cs.sync();
  assert.equal(holds().length, 1);
  assert.equal(holds()[0].extendedProperties.private.sourceCalendarId, PERSONAL);
  assert.equal(tagOf(ev), null);
});

test('walks every page of results', () => {
  const { cs, personal, holds } = setup(); // the fake serves two events per page
  for (let day = 14; day <= 18; day++) personal(`2026-09-${day}T10:00`, `2026-09-${day}T11:00`);
  cs.sync();
  assert.equal(holds().length, 5);
});

test('works with read-only access to the personal calendar when tagging is off', () => {
  const { cal, cs, personal, holds } = setup({ readOnly: true });
  cs.CONFIG.tagPersonalEvents = false;
  personal('2026-09-03T10:00', '2026-09-03T11:00');
  personal('2026-09-04T10:00', '2026-09-04T11:00');
  cs.sync();
  assert.equal(holds().length, 2);
  assert.ok(!cal.calls.includes('Events.patch'));
});

test('explains the missing permission when tagging is on and the personal calendar is read-only', () => {
  const { cs, personal } = setup({ readOnly: true });
  personal('2026-09-03T10:00', '2026-09-03T11:00');
  assert.throws(() => cs.install(), /Make changes to events.*tagPersonalEvents/);
  assert.equal(cs.triggers.length, 0);
});

test('install syncs first and only then schedules the trigger; uninstall removes everything', () => {
  const { cal, cs, personal, holds } = setup();
  personal('2026-09-03T10:00', '2026-09-03T11:00');
  const oldHold = { // far outside the sync window
    summary: cs.HOLD_TITLE, start: { dateTime: '2025-01-06T17:00:00Z' }, end: { dateTime: '2025-01-06T18:00:00Z' },
    extendedProperties: { private: { sourceEventId: 'old', sourceCalendarId: PERSONAL } },
  };
  cal.addEvent(WORK, oldHold);

  cs.install();
  assert.equal(holds().length, 1);
  assert.deepEqual(cs.triggers.map(t => [t.getHandlerFunction(), t.minutes]), [['sync', 5]]);

  cs.install(); // re-running replaces the trigger instead of stacking a second one
  assert.equal(cs.triggers.length, 1);

  cal.addEvent(WORK, oldHold);
  cs.uninstall();
  assert.equal(holds().length, 0);
  assert.equal(cs.triggers.length, 0);
  assert.equal(cs.logs.log.at(-1), 'CalSync uninstalled: removed 2 holds and the sync trigger');
});

test('install leaves no trigger behind when the first sync fails', () => {
  const { cs } = setup();
  cs.CONFIG.personalCalendarIds = ['not.shared@personal.example'];
  assert.throws(() => cs.install(), /Not Found/);
  assert.equal(cs.triggers.length, 0);
});

test('rejects an unedited or inconsistent CONFIG before touching any calendar', () => {
  const { cal, cs } = setup();
  cs.CONFIG.personalCalendarIds = ['your.personal@email.com'];
  assert.throws(() => cs.sync(), /personalCalendarIds/);
  cs.CONFIG.personalCalendarIds = [];
  assert.throws(() => cs.sync(), /personalCalendarIds/);

  cs.CONFIG.personalCalendarIds = [PERSONAL];
  cs.CONFIG.workStartHour = 17;
  cs.CONFIG.workEndHour = 9;
  assert.throws(() => cs.sync(), /workStartHour/);
  assert.deepEqual(cal.calls, []);
});

test('refuses to overlap with a run that is still in progress', () => {
  const { cal, cs } = setup({ lockAvailable: false });
  assert.throws(() => cs.sync(), /still in progress/);
  assert.throws(() => cs.uninstall(), /still in progress/);
  assert.deepEqual(cal.calls, []);
});

test('workRanges evaluates days and hours in the work calendar time zone', () => {
  const { cs } = setup();
  const ranges = (start, end, tz) => [...cs.workRanges(Date.parse(start), Date.parse(end), tz, [])]
    .map(r => [new Date(r.start).toISOString(), new Date(r.end).toISOString()]);

  assert.deepEqual(ranges('2026-09-04T23:30:00Z', '2026-09-05T01:00:00Z', 'Asia/Tokyo'), []); // Saturday morning in Tokyo
  assert.deepEqual(ranges('2026-09-05T02:00:00Z', '2026-09-05T03:00:00Z', 'America/Los_Angeles'), []); // Friday 7 PM in Los Angeles
  assert.deepEqual(ranges('2026-09-04T00:30:00Z', '2026-09-04T02:00:00Z', 'Asia/Tokyo'), // Friday 09:30–11:00 JST
    [['2026-09-04T00:30:00.000Z', '2026-09-04T02:00:00.000Z']]);
  assert.deepEqual(ranges('2026-09-04T00:30:00Z', '2026-09-04T10:00:00Z', 'Asia/Kolkata'), // Friday 06:00–15:30 IST
    [['2026-09-04T03:30:00.000Z', '2026-09-04T10:00:00.000Z']]);
  assert.deepEqual(ranges('2026-09-04T00:30:00Z', '2026-09-06T02:00:00Z', 'Asia/Kolkata'), []); // Fri 06:00 to Sun 07:30 IST: a full Friday workday, then weekend
});

test('midnightInTz lands on local midnight for every offset, including DST switch days', () => {
  const { cs } = setup();
  const cases = [
    ['2026-09-02', 'UTC'],
    ['2026-09-02', 'America/Los_Angeles'],
    ['2026-03-08', 'America/Los_Angeles'], // DST starts
    ['2026-11-01', 'America/Los_Angeles'], // DST ends
    ['2026-03-29', 'Europe/London'],
    ['2026-03-29', 'Europe/Berlin'],
    ['2026-10-25', 'Europe/Berlin'],
    ['2026-09-02', 'Asia/Kolkata'], // +05:30
    ['2026-10-04', 'Australia/Sydney'], // DST starts
    ['2026-04-05', 'Australia/Sydney'], // DST ends
    ['2026-09-02', 'Pacific/Auckland'], // +12
    ['2026-09-27', 'Pacific/Auckland'], // DST starts, +13
    ['2026-09-27', 'Pacific/Chatham'], // +12:45 to +13:45
    ['2026-09-02', 'Pacific/Tongatapu'], // +13
    ['2026-09-02', 'Pacific/Kiritimati'], // +14
  ];
  for (const [date, tz] of cases) {
    const got = formatDate(new Date(cs.midnightInTz(date, tz)), tz, 'yyyy-MM-dd HH:mm');
    assert.equal(got, `${date} 00:00`, `${tz} ${date}`);
  }
});
