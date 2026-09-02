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
 * each mirrored personal event is tagged with its hold's id so tools that only
 * see the personal calendar can tell holds from real conflicts.
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

  // Tag each mirrored personal event with the id of its work hold (private property
  // "workHoldId"), so tools that only see your personal calendar can tell holds from
  // real conflicts. Needs "Make changes to events" access to the personal calendars.
  tagPersonalEvents: true,
};

const HOLD_TITLE = '[DNS] External Appointment';
const HOLD_TAG = { key: 'calsync', value: 'hold' }; // Lets every hold be found in one query, whatever CONFIG says now
const SYNC_INTERVAL_MINUTES = 5; // Apps Script allows 1, 5, 10, 15, or 30
const HOUR_MS = 60 * 60 * 1000;
const DAY_MS = 24 * HOUR_MS;

// ── Lifecycle ───────────────────────────────────────────────

/** Run once from the editor: clears any previous installation, syncs, then schedules the sync. */
function install() {
  uninstall();
  sync(); // Before scheduling, so a misconfiguration fails here instead of every few minutes
  ScriptApp.newTrigger('sync').timeBased().everyMinutes(SYNC_INTERVAL_MINUTES).create();
}

/** Run from the editor to stop syncing and remove every hold CalSync created. */
function uninstall() {
  withLock(() => {
    for (const trigger of ScriptApp.getProjectTriggers()) {
      if (trigger.getHandlerFunction() === 'sync') ScriptApp.deleteTrigger(trigger);
    }

    const holdIds = new Set();
    listHolds({}, hold => holdIds.add(hold.id));
    // Holds from versions before HOLD_TAG existed can only be found through their source calendar.
    for (const calId of CONFIG.personalCalendarIds) {
      paginate(CONFIG.workCalendarId, { privateExtendedProperty: `sourceCalendarId=${calId}` }, hold => holdIds.add(hold.id));
    }
    holdIds.forEach(tryDelete);

    console.log(`CalSync uninstalled: removed ${holdIds.size} holds and the sync trigger`);
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
      const sourceEventId = hold.extendedProperties?.private?.sourceEventId;
      if (!sourceEventId) return;
      // A duplicate, or a hold someone edited into an all-day event, cannot be kept in step.
      const key = hold.start.dateTime && holdKey(sourceEventId, hold.start.dateTime, tz);
      if (!key || holds.has(key)) strays.push(hold.id);
      else holds.set(key, hold);
    });
    strays.forEach(tryDelete);
    stats.removed += strays.length;

    // Create or update the holds each personal event needs.
    const wanted = new Set();
    const seen = new Set();
    for (const calId of CONFIG.personalCalendarIds) {
      paginate(calId, { ...bounds, singleEvents: true }, ev => {
        seen.add(ev.id);
        const needed = blocksTime(ev) ? workRanges(new Date(ev.start.dateTime), new Date(ev.end.dateTime), tz, ooo) : [];
        // Holds outside the window were not indexed above, so touching them would add a
        // duplicate on every run. The event can still overlap the window, e.g. a trip.
        const inWindow = needed.filter(range => range.end > windowStart && range.start < windowEnd);
        const holdIds = inWindow.map(range => {
          const key = holdKey(ev.id, range.start, tz);
          wanted.add(key);
          const hold = holds.get(key);
          if (!hold) {
            stats.created++;
            return createHold(calId, ev.id, range);
          }
          if (updateHold(hold, range)) stats.updated++;
          return hold.id;
        });
        // Holds outside the window cannot be checked, so the tag is only cleared
        // once the event needs no hold at all.
        if (holdIds.length) tagPersonalEvent(calId, ev, holdIds[0]);
        else if (!needed.length) tagPersonalEvent(calId, ev, null);
      });
    }

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

/** One hold per personal event and day, so a multi-day event gets a hold on each of its weekdays. */
function holdKey(sourceEventId, start, tz) {
  return `${sourceEventId}|${dayOf(new Date(start), tz)}`;
}

// ── Work hours ──────────────────────────────────────────────

/**
 * The work-hour ranges of [start, end) that deserve a hold, as one {start, end}
 * in ms per day, in the work calendar's time zone. Weekends, days the event does
 * not reach work hours, holds of maxHoldHours or longer, and holds fully inside
 * an Out of Office block are left out.
 */
function workRanges(start, end, tz, ooo) {
  const ranges = [];
  const maxHoldMs = CONFIG.maxHoldHours * HOUR_MS;
  for (let midnight = midnightInTz(dayOf(start, tz), tz); midnight < end.getTime(); midnight = nextMidnight(midnight, tz)) {
    const range = {
      start: Math.max(start.getTime(), midnight + CONFIG.workStartHour * HOUR_MS),
      end: Math.min(end.getTime(), midnight + CONFIG.workEndHour * HOUR_MS),
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

/** Calendar day of `date` in `tz`, as yyyy-MM-dd. */
function dayOf(date, tz) {
  return Utilities.formatDate(date, tz, 'yyyy-MM-dd');
}

function isWeekend(ms, tz) {
  return Number(Utilities.formatDate(new Date(ms), tz, 'u')) >= 6; // ISO day of week: 6 = Saturday, 7 = Sunday
}

/** The instant at which the calendar day `dateStr` (yyyy-MM-dd) starts in `tz`. */
function midnightInTz(dateStr, tz) {
  const utcMidnight = Date.parse(`${dateStr}T00:00:00Z`);
  // A DST switch between UTC midnight and local midnight changes the offset, so re-read it at the estimate.
  const estimate = utcMidnight - tzOffsetMs(utcMidnight, tz);
  return utcMidnight - tzOffsetMs(estimate, tz);
}

function nextMidnight(midnight, tz) {
  return midnightInTz(dayOf(new Date(midnight + 36 * HOUR_MS), tz), tz); // 36h lands in the next day whether it has 23, 24, or 25 hours
}

/** UTC offset of `tz` at instant `ms`, in ms (e.g. -7 hours for PDT). */
function tzOffsetMs(ms, tz) {
  const rfc822 = Utilities.formatDate(new Date(ms), tz, 'Z'); // e.g. "-0700"
  const sign = rfc822[0] === '-' ? -1 : 1;
  return sign * (Number(rfc822.slice(1, 3)) * 60 + Number(rfc822.slice(3, 5))) * 60 * 1000;
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

function listHolds(params, fn) {
  paginate(CONFIG.workCalendarId, { ...params, privateExtendedProperty: `${HOLD_TAG.key}=${HOLD_TAG.value}` }, fn);
}

function createHold(sourceCalendarId, sourceEventId, range) {
  return Calendar.Events.insert({
    summary: HOLD_TITLE,
    start: { dateTime: new Date(range.start).toISOString() },
    end: { dateTime: new Date(range.end).toISOString() },
    visibility: CONFIG.holdVisibility,
    transparency: 'opaque',
    reminders: { useDefault: false, overrides: [] },
    extendedProperties: { private: { [HOLD_TAG.key]: HOLD_TAG.value, sourceEventId, sourceCalendarId } },
  }, CONFIG.workCalendarId).id;
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

function tagPersonalEvent(calId, ev, holdId) {
  if (!CONFIG.tagPersonalEvents || (ev.extendedProperties?.private?.workHoldId ?? null) === holdId) return;
  try {
    writeTag(calId, ev.id, holdId);
  } catch (e) {
    throw new Error(`Could not tag event ${ev.id} on ${calId} (${e.message}). Tagging needs ` +
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
