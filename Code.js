/**
 * CalSync
 *
 * Mirrors the busy time on your personal calendars as "[DNS] External Appointment"
 * holds on your work calendar: the work-hour part of each event, one hold per weekday
 * it touches. Skips all-day events, events marked free, invitations you have not
 * accepted, holds that would fill a whole workday, and holds inside an Out of Office block.
 *
 * Every run works out the holds that today and the coming weeks call for and makes the
 * work calendar match, so nothing is stored outside the calendars and personal calendars
 * are only ever read. A hold is just an interval of busy time: holds covering the same
 * interval are interchangeable, and a hold whose interval changes is replaced, not edited.
 * Every hold's event ID starts with HOLD_ID_PREFIX, which is how CalSync, and anyone with
 * free/busy access to the work calendar, tells holds from real meetings.
 */

// ── Configuration ───────────────────────────────────────────

const CONFIG = {
  // 'primary' is the calendar of the account running the script (usually your work account).
  workCalendarId: 'primary',

  // Personal calendars to mirror. Your work account needs "See all event details" access to each.
  personalCalendarIds: [
    'your.personal@email.com', // ⬅️ Replace with your personal calendar email
  ],

  workStartHour: 9,
  workEndHour: 17,
  syncDaysAhead: 28, // Days after today kept in step; earlier days are left as they were
  maxHoldHours: 8,   // Holds this long or longer are skipped; a full workday is better expressed as Out of Office
  holdVisibility: 'private', // 'private', 'public', or 'default'
};

const HOLD_TITLE = '[DNS] External Appointment';
const HOLD_ID_PREFIX = 'csdnshold'; // Chosen event IDs may only use a–v and 0–9
const HOUR_MS = 60 * 60 * 1000;
const DAY_MS = 24 * HOUR_MS;

// ── Lifecycle ───────────────────────────────────────────────

/** Run once from the editor: syncs, then schedules the sync every 5 minutes. Safe to run again. */
function install() {
  sync(); // First, so a misconfiguration fails here instead of every 5 minutes
  removeTrigger();
  ScriptApp.newTrigger('sync').timeBased().everyMinutes(5).create(); // Apps Script allows 1, 5, 10, 15, or 30
}

/** Run from the editor to stop syncing and remove every hold, past and future. */
function uninstall() {
  withLock(() => {
    removeTrigger();
    reconcile(new Map(), readWorkCalendar({}).holds);
  });
}

// ── Sync ────────────────────────────────────────────────────

function sync() {
  validateConfig();
  withLock(() => {
    const tz = Calendar.Calendars.get(CONFIG.workCalendarId).timeZone;
    const today = dayOf(Date.now(), tz);
    const window = { start: wallClock(today, 0, tz), end: wallClock(addDays(today, CONFIG.syncDaysAhead + 1), 0, tz) };
    const bounds = { timeMin: iso(window.start), timeMax: iso(window.end), singleEvents: true };

    // Every read comes before any write, so a failed read changes nothing and the next run retries.
    const work = readWorkCalendar(bounds);
    const outOfOffice = work.outOfOffice.map(ev => span(ev, tz));
    const events = CONFIG.personalCalendarIds.flatMap(calId => [...listEvents(calId, bounds)]);

    reconcile(wantedHolds(events, outOfOffice, tz, window), work.holds);
  });
}

/**
 * The holds the personal events call for, keyed by interval and visibility so that busy
 * time shared by several events, or by copies of one invitation, gets a single hold.
 */
function wantedHolds(events, outOfOffice, tz, window) {
  const wanted = new Map();
  for (const ev of events) {
    if (!blocksTime(ev)) continue;
    for (const hold of workHours(Date.parse(ev.start.dateTime), Date.parse(ev.end.dateTime), tz)) {
      if (hold.end - hold.start >= CONFIG.maxHoldHours * HOUR_MS) continue;
      if (outOfOffice.some(ooo => ooo.start <= hold.start && hold.end <= ooo.end)) continue;
      // An event can start before today, and holds outside the window are never read back,
      // so creating one would add a duplicate on every run.
      if (hold.end <= window.start || hold.start >= window.end) continue;
      wanted.set(holdKey(hold.start, hold.end, CONFIG.holdVisibility), hold);
    }
  }
  return wanted;
}

function blocksTime(ev) {
  if (!ev.start.dateTime || ev.transparency === 'transparent') return false; // All-day, or marked free
  const self = ev.attendees?.find(a => a.self);
  return !self || self.responseStatus === 'accepted'; // Appointments without an RSVP still block time
}

/** Makes the holds match `wanted`: keeps one hold per wanted key, removes every other, and creates the rest. */
function reconcile(wanted, holds) {
  const counts = { created: 0, removed: 0 };
  const failures = [];
  // Each change stands alone, so one failure cannot hold up the others.
  const attempt = (count, what, change) => {
    try {
      change();
      counts[count]++;
    } catch (e) {
      failures.push(`Could not ${what} (${e.message}).`);
    }
  };

  const kept = new Set();
  for (const hold of holds) {
    // A hold edited into an all-day event has no dateTime, so it matches nothing and is replaced.
    // The API leaves visibility out when it is 'default'.
    const key = holdKey(Date.parse(hold.start.dateTime), Date.parse(hold.end.dateTime), hold.visibility ?? 'default');
    if (wanted.has(key) && !kept.has(key)) kept.add(key);
    else attempt('removed', `remove hold ${hold.id}`, () => Calendar.Events.remove(CONFIG.workCalendarId, hold.id));
  }
  for (const [key, { start, end }] of wanted) {
    if (kept.has(key)) continue;
    const hold = newHold(start, end);
    attempt('created', `create the hold at ${hold.start.dateTime}`, () => Calendar.Events.insert(hold, CONFIG.workCalendarId));
  }

  failures.forEach(failure => console.error(failure));
  console.log(`CalSync: ${counts.created} created, ${counts.removed} removed`);
  if (failures.length) throw new Error(`${failures.length} hold change(s) failed. ${failures[0]}`);
}

function holdKey(start, end, visibility) {
  return `${start}|${end}|${visibility}`;
}

function newHold(start, end) {
  return {
    id: HOLD_ID_PREFIX + Utilities.getUuid().replace(/-/g, ''),
    summary: HOLD_TITLE,
    start: { dateTime: iso(start) },
    end: { dateTime: iso(end) },
    visibility: CONFIG.holdVisibility,
    transparency: 'opaque',
    reminders: { useDefault: false, overrides: [] },
  };
}

// ── Work hours ──────────────────────────────────────────────

/** The parts of [start, end) within work hours, one per weekday, on the work calendar's clock. */
function workHours(start, end, tz) {
  const parts = [];
  for (let day = dayOf(start, tz); wallClock(day, CONFIG.workStartHour, tz) < end; day = addDays(day, 1)) {
    const part = {
      start: Math.max(start, wallClock(day, CONFIG.workStartHour, tz)),
      end: Math.min(end, wallClock(day, CONFIG.workEndHour, tz)),
    };
    if (part.start < part.end && !isWeekend(day)) parts.push(part);
  }
  return parts;
}

// ── Time ────────────────────────────────────────────────────
// Instants are milliseconds since the epoch. Days are 'yyyy-MM-dd' dates, which need no
// time zone until an hour on a day's wall clock becomes an instant.

function dayOf(ms, tz) {
  return Utilities.formatDate(new Date(ms), tz, 'yyyy-MM-dd');
}

/** The instant the wall clock in tz shows `hour` o'clock on `day`. */
function wallClock(day, hour, tz) {
  return Utilities.parseDate(`${day} ${hour}`, tz, 'yyyy-MM-dd H').getTime();
}

/** A 'yyyy-MM-dd' date parses as UTC midnight, and every UTC day has 24 hours. */
function addDays(day, days) {
  return new Date(Date.parse(day) + days * DAY_MS).toISOString().slice(0, 10);
}

function isWeekend(day) {
  const weekday = new Date(day).getUTCDay();
  return weekday === 0 || weekday === 6;
}

/** An event's start and end; an all-day event runs from midnight to midnight in tz. */
function span(ev, tz) {
  const instant = time => (time.dateTime ? Date.parse(time.dateTime) : wallClock(time.date, 0, tz));
  return { start: instant(ev.start), end: instant(ev.end) };
}

function iso(ms) {
  return new Date(ms).toISOString();
}

// ── Calendar ────────────────────────────────────────────────

/**
 * The holds and Out of Office blocks on the work calendar. Only the fields CalSync reads are
 * fetched: whole events from years of history do not fit in memory.
 */
function readWorkCalendar(params) {
  const work = { holds: [], outOfOffice: [] };
  const query = {
    ...params,
    eventTypes: ['default', 'outOfOffice'],
    maxResults: 2500,
    fields: 'nextPageToken,items(id,eventType,start,end,visibility)', // Without nextPageToken, later pages are lost
  };
  for (const ev of listEvents(CONFIG.workCalendarId, query)) {
    if (ev.id.startsWith(HOLD_ID_PREFIX)) work.holds.push(ev);
    else if (ev.eventType === 'outOfOffice') work.outOfOffice.push(ev);
  }
  return work;
}

/** Every event a query matches, across all pages. */
function* listEvents(calId, params) {
  let pageToken;
  do {
    const page = Calendar.Events.list(calId, pageToken ? { ...params, pageToken } : params);
    yield* page.items ?? [];
    pageToken = page.nextPageToken;
  } while (pageToken);
}

// ── Helpers ─────────────────────────────────────────────────

function validateConfig() {
  const ids = CONFIG.personalCalendarIds;
  if (!ids.length || ids.includes('your.personal@email.com')) {
    throw new Error('Set CONFIG.personalCalendarIds to the email address(es) of your personal calendar(s).');
  }
  const { workStartHour: start, workEndHour: end } = CONFIG;
  if (!Number.isInteger(start) || !Number.isInteger(end) || start < 0 || start >= end || end > 24) {
    throw new Error('CONFIG.workStartHour and workEndHour must be whole hours, with 0 ≤ start < end ≤ 24.');
  }
}

/** A slow sync overlapping the next one would create every new hold twice. */
function withLock(fn) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30 * 1000)) throw new Error('Another CalSync run is still in progress.');
  try {
    fn();
  } finally {
    lock.releaseLock();
  }
}

function removeTrigger() {
  for (const trigger of ScriptApp.getProjectTriggers()) {
    if (trigger.getHandlerFunction() === 'sync') ScriptApp.deleteTrigger(trigger);
  }
}
