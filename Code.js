/**
 * CalSync
 *
 * Mirrors personal calendar events as "[DNS] External Appointment" holds on
 * your work calendar, clamped to work hours, one hold per weekday an event
 * touches. Skips all-day events, events marked free, declined invitations,
 * holds that would fill a whole workday, and holds fully covered by an Out of
 * Office block.
 *
 * Every run re-derives the holds for a rolling window from the personal
 * calendars, so nothing is stored outside the calendars themselves: each hold
 * names the personal event it mirrors in private extended properties, which
 * is what lets holds follow, update, and retire with their event. Optionally,
 * each mirrored personal event (a recurring one once, on its series) is tagged
 * with a hold's id so tools that only see the personal calendar can tell holds
 * from real conflicts.
 */

// ── Configuration ───────────────────────────────────────────

const CONFIG = {
  // 'primary' is the calendar of the account running the script (usually your work account).
  workCalendarId: 'primary',

  // Personal calendars to mirror. Your work account needs "Make changes to events" access
  // to each of them, or "See all event details" if tagPersonalEvents is false.
  personalCalendarIds: [
    'your.personal@email.com', // ⬅️ Replace with your personal calendar email
  ],

  workStartHour: 9,  // 9 AM
  workEndHour: 17,   // 5 PM
  syncDaysAhead: 28,
  maxHoldHours: 8,   // Holds this long or longer are skipped; a full workday is better expressed as Out of Office
  holdVisibility: 'private', // 'private', 'public', or 'default'

  // Tag each mirrored personal event, or recurring series, with the id of its earliest work hold
  // in the sync window (private property "workHoldId"), so tools that only see your personal
  // calendar can tell holds from real conflicts. Needs "Make changes to events" access.
  tagPersonalEvents: true,
};

const HOLD_TITLE = '[DNS] External Appointment';
const HOUR_MS = 60 * 60 * 1000;
const DAY_MS = 24 * HOUR_MS;

// ── Lifecycle ───────────────────────────────────────────────

/** Run once from the editor: clears any previous installation, syncs, then schedules the sync. */
function install() {
  uninstall();
  sync(); // Before scheduling, so a misconfiguration fails here instead of every few minutes
  ScriptApp.newTrigger('sync').timeBased().everyMinutes(5).create(); // Apps Script allows 1, 5, 10, 15, or 30
}

/** Run from the editor to stop syncing and remove every hold of the configured calendars. */
function uninstall() {
  withLock(() => {
    for (const trigger of ScriptApp.getProjectTriggers()) {
      if (trigger.getHandlerFunction() === 'sync') ScriptApp.deleteTrigger(trigger);
    }

    const holdIds = [];
    listHolds({}, hold => holdIds.push(hold.id));
    holdIds.forEach(tryDelete);

    console.log(`CalSync uninstalled: removed ${holdIds.length} holds and the sync trigger`);
  });
}

// ── Sync ────────────────────────────────────────────────────

function sync() {
  validateConfig();
  withLock(() => {
    const now = Date.now();
    const windowStart = now - DAY_MS;
    const windowEnd = now + CONFIG.syncDaysAhead * DAY_MS;
    const bounds = { timeMin: new Date(windowStart).toISOString(), timeMax: new Date(windowEnd).toISOString() };
    const tz = Calendar.Calendars.get(CONFIG.workCalendarId).timeZone;
    const ooo = getOOORanges(bounds, tz);
    const stats = { created: 0, updated: 0, removed: 0 };

    // Existing holds in the window, keyed by the personal event and day they mirror.
    const holds = new Map();
    const strays = [];
    listHolds(bounds, hold => {
      // A duplicate, a hold edited into an all-day event, or one from an older version cannot be kept in step.
      const uid = hold.extendedProperties.private.sourceUid;
      const key = uid && hold.start.dateTime && holdKey(uid, Date.parse(hold.start.dateTime), tz);
      if (!key || holds.has(key)) strays.push(hold.id);
      else holds.set(key, hold);
    });
    strays.forEach(tryDelete);
    stats.removed += strays.length;

    // Create or update the holds each personal event needs.
    const wanted = new Set();
    const seen = new Set();
    const tags = new Map(); // One per event or series: patching an occurrence would turn it into an exception
    for (const calId of CONFIG.personalCalendarIds) {
      paginate(calId, { ...bounds, singleEvents: true, orderBy: 'startTime' }, ev => {
        const eventId = ev.recurringEventId ?? ev.id;
        seen.add(eventId);
        const ranges = blocksTime(ev) ? workRanges(Date.parse(ev.start.dateTime), Date.parse(ev.end.dateTime), tz, ooo) : [];
        // Holds outside the window were not indexed above, so touching them would add a
        // duplicate on every run. The event can still overlap the window, e.g. a trip.
        const holdIds = ranges.filter(range => range.end > windowStart && range.start < windowEnd).map(range => {
          const key = holdKey(sourceUid(ev), range.start, tz);
          wanted.add(key);
          let hold = holds.get(key);
          if (hold) {
            if (updateHold(hold, range)) stats.updated++;
          } else {
            hold = createHold(calId, ev, range);
            holds.set(key, hold); // Another copy of the same invitation reuses it
            stats.created++;
          }
          return hold.id;
        });
        const tag = tags.get(`${calId}|${eventId}`)
          ?? { calId, eventId, series: Boolean(ev.recurringEventId), current: ev.extendedProperties?.private?.workHoldId ?? null, holdId: null };
        tag.holdId = tag.holdId ?? holdIds[0] ?? null; // Occurrences arrive in start order, so the earliest hold wins
        tags.set(`${calId}|${eventId}`, tag);
      });
    }
    for (const tag of tags.values()) tagPersonalEvent(tag);

    // Retire holds nothing needs any more.
    const untag = new Map();
    for (const [key, hold] of holds) {
      if (wanted.has(key)) continue;
      tryDelete(hold.id);
      stats.removed++;
      const { sourceEventId, sourceCalendarId } = hold.extendedProperties.private;
      // Events that left the window or were deleted were not seen above, so their tags are still set.
      if (!seen.has(sourceEventId)) untag.set(sourceEventId, sourceCalendarId);
    }
    for (const [eventId, calId] of untag) untagQuietly(calId, eventId);

    console.log(`CalSync: ${stats.created} created, ${stats.updated} updated, ${stats.removed} removed`);
  });
}

function blocksTime(ev) {
  if (!ev.start?.dateTime || ev.status === 'cancelled' || ev.transparency === 'transparent') return false;
  return ev.attendees?.find(a => a.self)?.responseStatus !== 'declined'; // Google Calendar shows declined events as free too
}

/**
 * Identity shared by every copy of an event, whichever calendar or system it
 * arrived through; occurrences of a series are told apart by their original slot.
 */
function sourceUid(ev) {
  const uid = ev.iCalUID ?? ev.id;
  return ev.originalStartTime ? `${uid}@${Date.parse(ev.originalStartTime.dateTime)}` : uid;
}

/** One hold per personal event and day, so a multi-day event gets a hold on each of its weekdays. */
function holdKey(uid, start, tz) {
  return `${uid}|${dayOf(start, tz)}`;
}

// ── Work hours ──────────────────────────────────────────────

/**
 * The work-hour ranges of [start, end) that deserve a hold, one {start, end} per
 * day, in the work calendar's time zone. Weekends, days the event does not reach
 * work hours, holds of maxHoldHours or longer, and holds fully inside an Out of
 * Office block are left out. All instants are ms since the epoch.
 */
function workRanges(start, end, tz, ooo) {
  const ranges = [];
  const maxHoldMs = CONFIG.maxHoldHours * HOUR_MS;
  for (let midnight = midnightInTz(dayOf(start, tz), tz); midnight < end; midnight = nextMidnight(midnight, tz)) {
    const range = {
      start: Math.max(start, midnight + CONFIG.workStartHour * HOUR_MS),
      end: Math.min(end, midnight + CONFIG.workEndHour * HOUR_MS),
    };
    if (range.start >= range.end || range.end - range.start >= maxHoldMs) continue;
    if (isWeekend(range.start, tz)) continue;
    if (ooo.some(block => block.start <= range.start && block.end >= range.end)) continue;
    ranges.push(range);
  }
  return ranges;
}

function getOOORanges(bounds, tz) {
  const ranges = [];
  paginate(CONFIG.workCalendarId, { ...bounds, eventTypes: ['outOfOffice'], singleEvents: true }, ev => {
    ranges.push({
      start: ev.start.dateTime ? Date.parse(ev.start.dateTime) : midnightInTz(ev.start.date, tz),
      end: ev.end.dateTime ? Date.parse(ev.end.dateTime) : midnightInTz(ev.end.date, tz),
    });
  });
  return ranges;
}

// ── Time zones ──────────────────────────────────────────────

/** Calendar day of instant `ms` in `tz`, as yyyy-MM-dd. */
function dayOf(ms, tz) {
  return Utilities.formatDate(new Date(ms), tz, 'yyyy-MM-dd');
}

function isWeekend(ms, tz) {
  return Number(Utilities.formatDate(new Date(ms), tz, 'u')) >= 6; // ISO day of week: 6 = Saturday, 7 = Sunday
}

/** The instant at which the calendar day `dateStr` (yyyy-MM-dd) starts in `tz`. */
function midnightInTz(dateStr, tz) {
  return Utilities.parseDate(dateStr, tz, 'yyyy-MM-dd').getTime();
}

function nextMidnight(midnight, tz) {
  return midnightInTz(dayOf(midnight + 36 * HOUR_MS, tz), tz); // 36h lands in the next day whether it has 23, 24, or 25 hours
}

// ── Helpers ─────────────────────────────────────────────────

function validateConfig() {
  const ids = CONFIG.personalCalendarIds;
  if (!ids.length || ids.includes('your.personal@email.com')) {
    throw new Error('Set CONFIG.personalCalendarIds to the email address(es) of your personal calendar(s).');
  }
  if (CONFIG.workStartHour >= CONFIG.workEndHour) {
    throw new Error('CONFIG.workStartHour must be earlier than CONFIG.workEndHour.');
  }
}

/** A slow sync overlapping the next trigger would create every hold twice. */
function withLock(fn) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30 * 1000)) throw new Error('Another CalSync run is still in progress.');
  try {
    fn();
  } finally {
    lock.releaseLock();
  }
}

function paginate(calId, params, fn) {
  let pageToken;
  do {
    const res = Calendar.Events.list(calId, pageToken ? { ...params, pageToken } : params);
    (res.items || []).forEach(fn);
    pageToken = res.nextPageToken;
  } while (pageToken);
}

/** Holds of the configured calendars, filtered by the API so only holds are ever loaded. */
function listHolds(params, fn) {
  for (const calId of CONFIG.personalCalendarIds) {
    paginate(CONFIG.workCalendarId, { ...params, privateExtendedProperty: `sourceCalendarId=${calId}` }, fn);
  }
}

function createHold(calId, ev, range) {
  return Calendar.Events.insert({
    summary: HOLD_TITLE,
    start: { dateTime: new Date(range.start).toISOString() },
    end: { dateTime: new Date(range.end).toISOString() },
    visibility: CONFIG.holdVisibility,
    transparency: 'opaque',
    reminders: { useDefault: false, overrides: [] },
    extendedProperties: { private: {
      sourceUid: sourceUid(ev),
      sourceEventId: ev.recurringEventId ?? ev.id, // The event, or series, that carries the tag
      sourceCalendarId: calId,
    } },
  }, CONFIG.workCalendarId);
}

/** Returns whether the hold had to change. */
function updateHold(hold, range) {
  const patch = {};
  if (Date.parse(hold.start.dateTime) !== range.start || Date.parse(hold.end.dateTime) !== range.end) {
    patch.start = { dateTime: new Date(range.start).toISOString() };
    patch.end = { dateTime: new Date(range.end).toISOString() };
  }
  if ((hold.visibility || 'default') !== CONFIG.holdVisibility) patch.visibility = CONFIG.holdVisibility; // The API omits 'default'
  if (!Object.keys(patch).length) return false;
  Calendar.Events.patch(patch, CONFIG.workCalendarId, hold.id);
  return true;
}

function tagPersonalEvent({ calId, eventId, series, current, holdId }) {
  if (!CONFIG.tagPersonalEvents) return;
  // Occurrences can carry stale copies of a series' tag, so read the series itself.
  if (series) current = Calendar.Events.get(calId, eventId).extendedProperties?.private?.workHoldId ?? null;
  if (current === holdId) return;
  try {
    writeTag(calId, eventId, holdId);
  } catch (e) {
    throw new Error(`Could not tag event ${eventId} on ${calId} (${e.message}). Tagging needs ` +
      '"Make changes to events" access to that calendar; grant it or set CONFIG.tagPersonalEvents to false.');
  }
}

/** For events that may no longer exist. */
function untagQuietly(calId, eventId) {
  if (!CONFIG.tagPersonalEvents) return;
  try {
    writeTag(calId, eventId, null);
  } catch (e) {
    // Nothing left to untag.
  }
}

function writeTag(calId, eventId, holdId) {
  Calendar.Events.patch({ extendedProperties: { private: { workHoldId: holdId } } }, calId, eventId);
}

function tryDelete(holdId) {
  try {
    Calendar.Events.remove(CONFIG.workCalendarId, holdId);
  } catch (e) {
    console.error(`Failed to delete hold ${holdId}: ${e.message}`);
  }
}
