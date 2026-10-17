import { test } from 'node:test';
import assert from 'node:assert/strict';
import { fakeCalendar, loadCalSync, formatDate } from './fakes.mjs';

const WORK = 'primary';
const PERSONAL = 'me@personal.example';
const OTHER = 'partner@personal.example';
const LA = 'America/Los_Angeles';
const pdt = s => `${s}:00-07:00`; // Los Angeles wall-clock time in September
const at = s => Date.parse(pdt(s));
const local = (ms, tz = LA) => formatDate(new Date(ms), tz, 'yyyy-MM-dd HH:mm');
const invite = responseStatus => ({ attendees: [
  { email: 'organizer@example.com', organizer: true, responseStatus: 'accepted' },
  { email: PERSONAL, self: true, responseStatus },
] });

/**
 * A work calendar and a read-only personal calendar, with CalSync loaded and pointed at them.
 * The personal calendar refuses writes, so every test also checks CalSync only reads it.
 * The clock starts on Wednesday 2 September 2026 at noon in Los Angeles.
 */
function setup({ now = '2026-09-02T12:00', workTimeZone = LA, lockAvailable = true } = {}) {
  const cal = fakeCalendar();
  cal.addCalendar(WORK, { timeZone: workTimeZone });
  cal.addCalendar(PERSONAL, { readOnly: true });
  const clock = { now: at(now) };
  const cs = loadCalSync({ calendar: cal, now: () => clock.now, lockAvailable });
  cs.CONFIG.personalCalendarIds = [PERSONAL];
  const personal = (start, end, extra = {}, calId = PERSONAL) =>
    cal.addEvent(calId, { start: { dateTime: pdt(start) }, end: { dateTime: pdt(end) }, ...extra });
  const holds = () => cal.live(WORK)
    .filter(ev => ev.id.startsWith('csdnshold'))
    .sort((a, b) => Date.parse(a.start.dateTime) - Date.parse(b.start.dateTime));
  /** Each hold as [start, end] on the work calendar's wall clock. */
  const times = () => holds().map(h => [h.start, h.end].map(t => local(Date.parse(t.dateTime), workTimeZone)));
  return { cal, cs, clock, personal, holds, times };
}

// ── What gets a hold ────────────────────────────────────────

test('mirrors a personal event as a private, busy hold on its work-hour part', () => {
  const { cs, personal, holds, times } = setup();
  personal('2026-09-03T08:30', '2026-09-03T10:15');
  cs.sync();

  assert.deepEqual(times(), [['2026-09-03 09:00', '2026-09-03 10:15']]);
  const [hold] = holds();
  assert.match(hold.id, /^csdnshold[0-9a-f]{32}$/);
  assert.equal(hold.summary, '[DNS] External Appointment');
  assert.equal(hold.visibility, 'private');
  assert.equal(hold.transparency, 'opaque');
  assert.deepEqual(hold.reminders, { useDefault: false, overrides: [] });
  assert.equal(hold.extendedProperties, undefined);
  assert.deepEqual(cs.logs.log, ['CalSync: 1 created, 0 removed']);
});

test('skips all-day events, events marked free, weekends, time outside work hours, and full workdays', () => {
  const { cal, cs, personal, holds } = setup();
  cal.addEvent(PERSONAL, { start: { date: '2026-09-03' }, end: { date: '2026-09-04' } });
  personal('2026-09-03T10:00', '2026-09-03T11:00', { transparency: 'transparent' });
  personal('2026-09-05T10:00', '2026-09-05T11:00'); // Saturday
  personal('2026-09-03T18:00', '2026-09-03T19:00');
  personal('2026-09-03T07:00', '2026-09-03T09:00'); // ends as work starts
  personal('2026-09-03T09:00', '2026-09-03T17:00'); // Out of Office says this better
  cs.sync();
  assert.equal(holds().length, 0);
  assert.deepEqual(cal.writes(), []);
});

test('mirrors appointments and accepted invitations, not tentative, unanswered, or declined ones', () => {
  const { cs, personal, times } = setup();
  personal('2026-09-03T09:00', '2026-09-03T10:00');
  personal('2026-09-03T11:00', '2026-09-03T12:00', invite('accepted'));
  for (const response of ['tentative', 'needsAction', 'declined', undefined]) {
    personal('2026-09-03T14:00', '2026-09-03T15:00', invite(response));
  }
  cs.sync();
  assert.deepEqual(times(), [['2026-09-03 09:00', '2026-09-03 10:00'], ['2026-09-03 11:00', '2026-09-03 12:00']]);
});

test('holds each weekday of a multi-day event', () => {
  const { cs, personal, times } = setup();
  personal('2026-09-07T15:00', '2026-09-09T11:00'); // Monday afternoon to Wednesday morning
  personal('2026-09-13T22:00', '2026-09-14T10:00'); // Sunday night to Monday morning
  cs.sync();
  assert.deepEqual(times(), [
    ['2026-09-07 15:00', '2026-09-07 17:00'],
    ['2026-09-09 09:00', '2026-09-09 11:00'], // Tuesday would be a full workday
    ['2026-09-14 09:00', '2026-09-14 10:00'],
  ]);
});

test('skips holds inside Out of Office, timed or all-day, and keeps partly covered ones', () => {
  const { cal, cs, personal, times } = setup();
  cal.addEvent(WORK, { eventType: 'outOfOffice', start: { dateTime: pdt('2026-09-03T09:00') }, end: { dateTime: pdt('2026-09-03T12:00') } });
  cal.addEvent(WORK, { eventType: 'outOfOffice', start: { date: '2026-09-04' }, end: { date: '2026-09-05' } });
  personal('2026-09-03T10:00', '2026-09-03T11:00');
  personal('2026-09-04T10:00', '2026-09-04T11:00');
  personal('2026-09-03T11:30', '2026-09-03T13:00'); // runs past the end of the timed block
  cs.sync();
  assert.deepEqual(times(), [['2026-09-03 11:30', '2026-09-03 13:00']]);
});

test('gives busy time one hold, however many events or calendars it comes from', () => {
  const { cal, cs, personal, times } = setup();
  cal.addCalendar(OTHER, { readOnly: true });
  cs.CONFIG.personalCalendarIds = [PERSONAL, OTHER];
  const copies = [PERSONAL, OTHER].map(calId => personal('2026-09-03T10:00', '2026-09-03T11:00', { iCalUID: 'shared@google.com' }, calId));
  personal('2026-09-03T10:00', '2026-09-03T11:00'); // a different event at the same time
  personal('2026-09-03T10:30', '2026-09-03T11:30'); // overlapping time is its own hold
  cs.sync();
  assert.deepEqual(times(), [['2026-09-03 10:00', '2026-09-03 11:00'], ['2026-09-03 10:30', '2026-09-03 11:30']]);

  for (const copy of copies) copy.status = 'cancelled';
  cs.sync();
  assert.equal(times().length, 2);
  assert.equal(cs.logs.log.at(-1), 'CalSync: 0 created, 0 removed');
});

// ── Time zones ──────────────────────────────────────────────

test('reads work hours and weekends on the work calendar\'s clock', () => {
  const { cs } = setup();
  const parts = (start, end, tz) => [...cs.workHours(Date.parse(start), Date.parse(end), tz)].map(p => [local(p.start, tz), local(p.end, tz)]);

  assert.deepEqual(parts('2026-09-04T23:30:00Z', '2026-09-05T01:00:00Z', 'Asia/Tokyo'), []); // Saturday morning in Tokyo
  assert.deepEqual(parts('2026-09-05T02:00:00Z', '2026-09-05T03:00:00Z', LA), []); // Friday evening in Los Angeles
  assert.deepEqual(parts('2026-09-04T00:30:00Z', '2026-09-04T02:00:00Z', 'Asia/Tokyo'), [['2026-09-04 09:30', '2026-09-04 11:00']]);
  assert.deepEqual(parts('2026-09-04T00:30:00Z', '2026-09-06T02:00:00Z', 'Asia/Kolkata'), // Friday 06:00 to Sunday 07:30
    [['2026-09-04 09:00', '2026-09-04 17:00']]);
});

test('keeps work hours on the wall clock on a weekday the clocks change', () => {
  const { cs } = setup();
  const tz = 'Asia/Jerusalem'; // Clocks go forward at 02:00 on Friday 27 March 2026
  const [part] = cs.workHours(Date.parse('2026-03-27T05:00:00Z'), Date.parse('2026-03-27T09:00:00Z'), tz); // 08:00–12:00 local
  assert.deepEqual([local(part.start, tz), local(part.end, tz)], ['2026-03-27 09:00', '2026-03-27 12:00']);
});

test('places holds on the work calendar\'s clock, whatever zone the personal event uses', () => {
  const { cs, personal, times } = setup({ workTimeZone: 'Europe/London' });
  personal('2026-09-03T00:00', '2026-09-03T02:00'); // 08:00–10:00 in London
  cs.sync();
  assert.deepEqual(times(), [['2026-09-03 09:00', '2026-09-03 10:00']]);
});

// ── Keeping holds in step ───────────────────────────────────

test('a second run changes nothing', () => {
  const { cal, cs, personal } = setup();
  personal('2026-09-03T10:00', '2026-09-03T11:00');
  personal('2026-09-07T15:00', '2026-09-09T11:00');
  cal.addEvent(PERSONAL, { id: 'yoga', recurrence: ['RRULE:FREQ=WEEKLY'], start: { dateTime: pdt('2026-09-03T12:00') }, end: { dateTime: pdt('2026-09-03T13:00') } });
  personal('2026-09-03T12:00', '2026-09-03T13:00', { recurringEventId: 'yoga' });
  cs.sync();

  cal.calls.length = 0;
  cs.sync();
  assert.deepEqual(cal.writes(), []);
  assert.equal(cs.logs.log.at(-1), 'CalSync: 0 created, 0 removed');
});

test('replaces a hold when its event moves', () => {
  const { cs, personal, holds, times } = setup();
  const ev = personal('2026-09-03T10:00', '2026-09-03T11:00');
  cs.sync();
  const before = holds()[0].id;

  ev.start.dateTime = pdt('2026-09-03T13:00');
  ev.end.dateTime = pdt('2026-09-03T14:00');
  cs.sync();
  assert.deepEqual(times(), [['2026-09-03 13:00', '2026-09-03 14:00']]);
  assert.notEqual(holds()[0].id, before);
  assert.equal(cs.logs.log.at(-1), 'CalSync: 1 created, 1 removed');
});

test('removes holds whose events are deleted, declined, marked free, or covered by Out of Office', () => {
  const { cal, cs, personal, holds } = setup();
  const deleted = personal('2026-09-03T10:00', '2026-09-03T11:00');
  const declined = personal('2026-09-03T13:00', '2026-09-03T14:00', invite('accepted'));
  const freed = personal('2026-09-04T10:00', '2026-09-04T11:00');
  personal('2026-09-07T10:00', '2026-09-07T11:00');
  cs.sync();
  assert.equal(holds().length, 4);

  deleted.status = 'cancelled';
  declined.attendees[1].responseStatus = 'declined';
  freed.transparency = 'transparent';
  cal.addEvent(WORK, { eventType: 'outOfOffice', start: { date: '2026-09-07' }, end: { date: '2026-09-08' } });
  cs.sync();
  assert.equal(holds().length, 0);
  assert.equal(cs.logs.log.at(-1), 'CalSync: 0 created, 4 removed');
});

test('follows each occurrence of a recurring event', () => {
  const { cal, cs, personal, times } = setup();
  cal.addEvent(PERSONAL, { id: 'yoga', recurrence: ['RRULE:FREQ=WEEKLY'], start: { dateTime: pdt('2026-09-03T12:00') }, end: { dateTime: pdt('2026-09-03T13:00') } });
  const [, cancelled, moved] = ['03', '10', '17'].map(day => personal(`2026-09-${day}T12:00`, `2026-09-${day}T13:00`, { recurringEventId: 'yoga' }));
  cs.sync();
  assert.equal(times().length, 3);

  cancelled.status = 'cancelled';
  moved.start.dateTime = pdt('2026-09-18T12:00');
  moved.end.dateTime = pdt('2026-09-18T13:00');
  cs.sync();
  assert.deepEqual(times(), [['2026-09-03 12:00', '2026-09-03 13:00'], ['2026-09-18 12:00', '2026-09-18 13:00']]);
});

test('undoes hand edits: a hold deleted, moved, made all-day, or copied', () => {
  const { cal, cs, personal, holds, times } = setup();
  for (const day of ['03', '04', '08', '09']) personal(`2026-09-${day}T10:00`, `2026-09-${day}T11:00`);
  cs.sync();
  const expected = times();
  const [deleted, moved, allDay, copied] = holds();

  cal.api.Events.remove(WORK, deleted.id);
  moved.start = { dateTime: pdt('2026-09-04T12:00') };
  moved.end = { dateTime: pdt('2026-09-04T13:00') };
  allDay.start = { date: '2026-09-08' };
  allDay.end = { date: '2026-09-09' };
  cal.addEvent(WORK, { ...JSON.parse(JSON.stringify(copied)), id: 'csdnshold0duplicate' });
  cs.sync();
  assert.deepEqual(times(), expected);
  assert.equal(cs.logs.log.at(-1), 'CalSync: 3 created, 3 removed');
});

test('applies a changed holdVisibility to existing holds', () => {
  const { cs, personal, holds } = setup();
  personal('2026-09-03T10:00', '2026-09-03T11:00');
  cs.sync();
  cs.CONFIG.holdVisibility = 'default';
  cs.sync();
  assert.equal(holds()[0].visibility, undefined); // The API leaves default visibility out
  assert.equal(cs.logs.log.at(-1), 'CalSync: 1 created, 1 removed');

  cs.sync();
  assert.equal(cs.logs.log.at(-1), 'CalSync: 0 created, 0 removed');
});

test('removes the holds of a calendar taken out of CONFIG', () => {
  const { cal, cs, personal, times } = setup();
  cal.addCalendar(OTHER, { readOnly: true });
  cs.CONFIG.personalCalendarIds = [PERSONAL, OTHER];
  personal('2026-09-03T10:00', '2026-09-03T11:00');
  personal('2026-09-03T14:00', '2026-09-03T15:00', {}, OTHER);
  cs.sync();
  assert.equal(times().length, 2);

  cs.CONFIG.personalCalendarIds = [PERSONAL];
  cs.sync();
  assert.deepEqual(times(), [['2026-09-03 10:00', '2026-09-03 11:00']]);
});

test('leaves the work calendar\'s own events alone, even ones that look like holds', () => {
  const { cal, cs, personal } = setup();
  const own = [
    cal.addEvent(WORK, { summary: 'Standup', start: { dateTime: pdt('2026-09-03T10:00') }, end: { dateTime: pdt('2026-09-03T10:30') } }),
    cal.addEvent(WORK, { summary: 'Weekly', recurrence: ['RRULE:FREQ=WEEKLY'], start: { dateTime: pdt('2026-09-03T14:00') }, end: { dateTime: pdt('2026-09-03T15:00') } }),
    cal.addEvent(WORK, { eventType: 'outOfOffice', start: { date: '2026-09-10' }, end: { date: '2026-09-11' } }),
    cal.addEvent(WORK, { summary: '[DNS] External Appointment', start: { dateTime: pdt('2026-09-04T10:00') }, end: { dateTime: pdt('2026-09-04T11:00') } }),
  ];
  personal('2026-09-03T10:00', '2026-09-03T11:00');
  cs.sync();
  cs.uninstall();
  assert.deepEqual(cal.live(WORK).map(ev => ev.id), own.map(ev => ev.id));
  assert.deepEqual(cs.logs.error, []);
});

test('keeps today and the next syncDaysAhead days in step, and leaves earlier days alone', () => {
  const { cs, clock, personal, times } = setup();
  const earlier = personal('2026-09-02T09:00', '2026-09-02T10:00'); // over by noon
  const lunch = personal('2026-09-02T11:00', '2026-09-02T11:30');
  personal('2026-09-30T10:00', '2026-09-30T11:00'); // the last day in the window
  personal('2026-10-01T10:00', '2026-10-01T11:00'); // the day after
  cs.sync();
  assert.deepEqual(times(), [
    ['2026-09-02 09:00', '2026-09-02 10:00'],
    ['2026-09-02 11:00', '2026-09-02 11:30'],
    ['2026-09-30 10:00', '2026-09-30 11:00'],
  ]);

  lunch.status = 'cancelled'; // today is still in step
  cs.sync();
  clock.now = at('2026-09-03T08:00');
  earlier.status = 'cancelled'; // yesterday is history
  cs.sync();
  assert.deepEqual(times(), [
    ['2026-09-02 09:00', '2026-09-02 10:00'],
    ['2026-09-30 10:00', '2026-09-30 11:00'],
    ['2026-10-01 10:00', '2026-10-01 11:00'],
  ]);
});

test('never duplicates the holds of an event that started before today', () => {
  const { cs, clock, personal, times } = setup({ now: '2026-09-04T09:00' });
  personal('2026-09-04T14:00', '2026-09-07T12:00'); // Friday afternoon to Monday noon
  for (const now of ['2026-09-04T09:00', '2026-09-05T18:00', '2026-09-07T09:00', '2026-09-08T09:00']) {
    clock.now = at(now);
    cs.sync();
  }
  assert.deepEqual(times(), [['2026-09-04 14:00', '2026-09-04 17:00'], ['2026-09-07 09:00', '2026-09-07 12:00']]);
});

test('reads every page of results', () => {
  const { cal, cs, personal } = setup(); // The fake serves two events per page
  for (let day = 14; day <= 18; day++) {
    personal(`2026-09-${day}T10:00`, `2026-09-${day}T11:00`);
    cal.addEvent(WORK, { summary: 'Meeting', start: { dateTime: pdt(`2026-09-${day}T15:00`) }, end: { dateTime: pdt(`2026-09-${day}T16:00`) } });
  }
  cs.sync();
  assert.equal(cs.logs.log.at(-1), 'CalSync: 5 created, 0 removed');
  cs.sync();
  assert.equal(cs.logs.log.at(-1), 'CalSync: 0 created, 0 removed');
});

// ── Failures ────────────────────────────────────────────────

test('a failed change does not stop the others, fails the run, and is retried by the next', () => {
  const { cal, cs, personal, times } = setup();
  const gone = personal('2026-09-03T10:00', '2026-09-03T11:00');
  cs.sync();
  gone.status = 'cancelled';
  personal('2026-09-04T10:00', '2026-09-04T11:00');
  personal('2026-09-08T10:00', '2026-09-08T11:00');
  const insert = cal.api.Events.insert;
  cal.api.Events.insert = (resource, calId) => {
    if (Date.parse(resource.start.dateTime) === at('2026-09-04T10:00')) throw new Error('Backend Error');
    return insert(resource, calId);
  };

  assert.throws(() => cs.sync(), /1 hold change\(s\) failed\. Could not create the hold at 2026-09-04T17:00:00\.000Z \(Backend Error\)/);
  assert.deepEqual(times(), [['2026-09-08 10:00', '2026-09-08 11:00']]);
  assert.equal(cs.logs.log.at(-1), 'CalSync: 1 created, 1 removed');

  cal.api.Events.insert = insert;
  cs.sync();
  assert.deepEqual(times(), [['2026-09-04 10:00', '2026-09-04 11:00'], ['2026-09-08 10:00', '2026-09-08 11:00']]);
});

test('a failed read changes nothing', () => {
  const { cal, cs, personal, times } = setup();
  cal.addCalendar(OTHER, { readOnly: true });
  cs.CONFIG.personalCalendarIds = [PERSONAL, OTHER];
  const ev = personal('2026-09-03T10:00', '2026-09-03T11:00');
  cs.sync();
  ev.status = 'cancelled';
  personal('2026-09-04T10:00', '2026-09-04T11:00');
  const list = cal.api.Events.list;
  cal.api.Events.list = (calId, params) => {
    if (calId === OTHER) throw new Error('Backend Error');
    return list(calId, params);
  };

  cal.calls.length = 0;
  assert.throws(() => cs.sync(), /Backend Error/);
  assert.deepEqual(cal.writes(), []);
  assert.deepEqual(times(), [['2026-09-03 10:00', '2026-09-03 11:00']]);
});

test('refuses to overlap with a run that is still in progress', () => {
  const { cal, cs } = setup({ lockAvailable: false });
  assert.throws(() => cs.sync(), /still in progress/);
  assert.throws(() => cs.uninstall(), /still in progress/);
  assert.deepEqual(cal.calls, []);
});

test('rejects an unedited or inconsistent CONFIG before touching any calendar', () => {
  const { cal, cs } = setup();
  cs.CONFIG.personalCalendarIds = ['your.personal@email.com'];
  assert.throws(() => cs.install(), /personalCalendarIds/);
  cs.CONFIG.personalCalendarIds = [];
  assert.throws(() => cs.sync(), /personalCalendarIds/);

  cs.CONFIG.personalCalendarIds = [PERSONAL];
  for (const [start, end] of [[17, 9], [9, 9], [8.5, 17], [-1, 17], [9, 25]]) {
    Object.assign(cs.CONFIG, { workStartHour: start, workEndHour: end });
    assert.throws(() => cs.sync(), /workStartHour/, `${start}–${end}`);
  }
  assert.deepEqual(cal.calls, []);
  assert.equal(cs.triggers.length, 0);
});

// ── Install and uninstall ───────────────────────────────────

test('install syncs, then schedules one trigger, and running it again changes nothing', () => {
  const { cal, cs, personal, holds } = setup();
  personal('2026-09-03T10:00', '2026-09-03T11:00');
  cs.install();
  assert.deepEqual(cs.triggers.map(t => [t.getHandlerFunction(), t.minutes]), [['sync', 5]]);
  const ids = holds().map(h => h.id);

  cal.calls.length = 0;
  cs.install();
  assert.deepEqual(cal.writes(), []);
  assert.deepEqual(holds().map(h => h.id), ids);
  assert.equal(cs.triggers.length, 1);
});

test('install leaves no trigger behind when the first sync fails', () => {
  const { cs } = setup();
  cs.CONFIG.personalCalendarIds = ['not.shared@personal.example'];
  assert.throws(() => cs.install(), /Not Found/);
  assert.equal(cs.triggers.length, 0);
});

test('uninstall removes the trigger and every hold, past and future, and can be rerun', () => {
  const { cal, cs, personal, holds } = setup();
  cal.addEvent(WORK, { id: 'csdnshold0past', summary: '[DNS] External Appointment', start: { dateTime: '2025-01-06T17:00:00Z' }, end: { dateTime: '2025-01-06T18:00:00Z' } });
  personal('2026-09-03T10:00', '2026-09-03T11:00');
  personal('2026-09-04T10:00', '2026-09-04T11:00');
  cs.install();
  assert.equal(holds().length, 3);

  const remove = cal.api.Events.remove;
  cal.api.Events.remove = (calId, id) => {
    if (id === 'csdnshold0past') throw new Error('Backend Error');
    return remove(calId, id);
  };
  assert.throws(() => cs.uninstall(), /1 hold change\(s\) failed/);
  assert.equal(cs.triggers.length, 0);
  assert.deepEqual(holds().map(h => h.id), ['csdnshold0past']);

  cal.api.Events.remove = remove;
  cs.uninstall();
  assert.equal(holds().length, 0);
  assert.equal(cs.logs.log.at(-1), 'CalSync: 0 created, 1 removed');
});
