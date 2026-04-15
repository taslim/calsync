/**
 * CalSync
 * * Mirrors personal calendar events as private "[DNS] External Appointment"
 * holds on your work calendar during work hours. Skips weekends, free
 * personal events, and holds fully covered by OOO.
 *
 * Uses a stateless rolling window to ensure events naturally aging
 * into the window are caught, and infinite recurrences are ignored.
 * Safe for multiple users on the same event via private extended properties.
 */

// ── Configuration ───────────────────────────────────────────

const CONFIG = {
  // 'primary' refers to the calendar of the account running the script (usually your work account)
  workCalendarId: 'primary', 
  
  // Add the email addresses of the personal calendars you want to sync
  // Note: Your work account MUST have at least "See all event details" access to these calendars
  personalCalendarIds: [
    'your.personal@email.com', // ⬅️ Replace with your personal calendar email
  ],
  
  workStartHour: 9,  // 9 AM
  workEndHour: 17,   // 5 PM
  syncDaysAhead: 28, // How many days into the future to sync
  maxHoldHours: 8,   // Skips events that span more than 8 hours (e.g., full workday blocks)
};

const HOLD_TITLE = '[DNS] External Appointment';

// ── Lifecycle ───────────────────────────────────────────────

/**
 * Run this function ONCE to set up the sync.
 * It will clear any old triggers, set up a new 5-minute timer, and run an initial sync.
 */
function install() {
  uninstall();
  ScriptApp.newTrigger('sync').timeBased().everyMinutes(5).create();
  sync();
}

/**
 * Run this function to stop the sync and clean up existing holds.
 */
function uninstall() {
  for (const t of ScriptApp.getProjectTriggers()) {
    if (t.getHandlerFunction() === 'sync') ScriptApp.deleteTrigger(t);
  }
  for (const calId of CONFIG.personalCalendarIds) {
    paginate(CONFIG.workCalendarId, { privateExtendedProperty: `sourceCalendarId=${calId}` }, ev => tryDelete(ev.id));
  }
}

// ── Sync ────────────────────────────────────────────────────

function sync() {
  const now = new Date();
  const sweepStart = new Date(now.getTime() - 86400000);
  const sweepEnd = new Date(now.getTime() + CONFIG.syncDaysAhead * 86400000);
  const tz = Calendar.Calendars.get(CONFIG.workCalendarId).timeZone;
  
  const ooo = getOOORanges(sweepStart, sweepEnd, tz);

  // 1. Index existing holds
  const activeHolds = new Map();

  for (const calId of CONFIG.personalCalendarIds) {
    paginate(CONFIG.workCalendarId, {
      timeMin: sweepStart.toISOString(),
      timeMax: sweepEnd.toISOString(),
      singleEvents: true,
      privateExtendedProperty: `sourceCalendarId=${calId}`
    }, ev => {
      if (ev.status === 'cancelled') return;
      
      const srcId = ev.extendedProperties?.private?.sourceEventId;
      if (srcId) activeHolds.set(srcId, ev);
    });
  }

  const processedSourceIds = new Set();

  // 2. Process personal events
  for (const calId of CONFIG.personalCalendarIds) {
    paginate(calId, {
      timeMin: sweepStart.toISOString(),
      timeMax: sweepEnd.toISOString(),
      singleEvents: true
    }, ev => {      
      if (!ev.start?.dateTime || ev.transparency === 'transparent' || ev.status === 'cancelled') return;

      const clamped = clampToWorkHours(new Date(ev.start.dateTime), new Date(ev.end.dateTime), tz, ooo);
      const existingHold = activeHolds.get(ev.id);      
      
      if (!clamped) return;

      processedSourceIds.add(ev.id);      
      
      if (existingHold) {        
        if (new Date(existingHold.start.dateTime).getTime() !== clamped.start.getTime() ||            
            new Date(existingHold.end.dateTime).getTime() !== clamped.end.getTime()) {
          Calendar.Events.patch({
            start: { dateTime: clamped.start.toISOString() },
            end: { dateTime: clamped.end.toISOString() },          
          }, CONFIG.workCalendarId, existingHold.id);        
        }
        if (ev.extendedProperties?.private?.workHoldId !== existingHold.id) {
          linkPersonalEvent(calId, ev.id, existingHold.id);
        }
      } else {        
        const newHoldId = createHold(calId, ev.id, clamped);        
        linkPersonalEvent(calId, ev.id, newHoldId);
      }
    });
  }

  // 3. Cleanup orphaned holds
  for (const [srcId, hold] of activeHolds.entries()) {
    if (!processedSourceIds.has(srcId)) {      
      tryDelete(hold.id);
      try { 
        linkPersonalEvent(hold.extendedProperties.private.sourceCalendarId, srcId, null); 
      } catch (e) {
      }
    }
  }
}

// ── OOO ─────────────────────────────────────────────────────

function getOOORanges(start, end, tz) {
  const ranges = [];
  paginate(CONFIG.workCalendarId, {
    timeMin: start.toISOString(),
    timeMax: end.toISOString(),
    eventTypes: ['outOfOffice'],
    singleEvents: true,
  }, ev => {
    ranges.push({
      start: ev.start.dateTime ? new Date(ev.start.dateTime).getTime() : midnightInTz(ev.start.date, tz),
      end:   ev.end.dateTime   ? new Date(ev.end.dateTime).getTime()   : midnightInTz(ev.end.date, tz),
    });
  });
  return ranges;
}

function midnightInTz(dateStr, tz) {
  const noon = new Date(dateStr + 'T12:00:00Z');
  const h = Number(Utilities.formatDate(noon, tz, 'H'));
  const m = Number(Utilities.formatDate(noon, tz, 'm'));
  return noon.getTime() - (h * 60 + m) * 60000;
}

// ── Clamping ────────────────────────────────────────────────

function clampToWorkHours(start, end, tz, ooo) {
  const fmt = (d, p) => Number(Utilities.formatDate(d, tz, p));  
  if (fmt(start, 'u') >= 6) return null;  
  
  const startDay = Utilities.formatDate(start, tz, 'yyyyMMdd');  
  const endDay = Utilities.formatDate(end, tz, 'yyyyMMdd');  
  const startMin = fmt(start, 'H') * 60 + fmt(start, 'm');  
  const endMin = startDay === endDay ? fmt(end, 'H') * 60 + fmt(end, 'm') : CONFIG.workEndHour * 60;  
  const clampStart = Math.max(startMin, CONFIG.workStartHour * 60);  
  const clampEnd = Math.min(endMin, CONFIG.workEndHour * 60);  
  
  if (clampStart >= clampEnd) return null;  
  
  const durationHours = (clampEnd - clampStart) / 60;
  if (durationHours >= CONFIG.maxHoldHours) return null;
  
  const result = {
    start: new Date(start.getTime() + (clampStart - startMin) * 60000),    
    end:   new Date(start.getTime() + (clampEnd - startMin) * 60000),
  };  
  
  if (ooo.some(r => r.start <= result.start.getTime() && r.end >= result.end.getTime())) return null;  
  return result;
}

// ── Helpers ─────────────────────────────────────────────────

function paginate(calId, params, fn) {
  let pageToken;
  do {
    const p = pageToken ? { ...params, pageToken } : params;
    const res = Calendar.Events.list(calId, p);
    (res.items || []).forEach(fn);
    pageToken = res.nextPageToken;
  } while (pageToken);
}

function createHold(sourceCalId, sourceEventId, { start, end }) {
  return Calendar.Events.insert({
    summary: HOLD_TITLE,
    start: { dateTime: start.toISOString() },
    end: { dateTime: end.toISOString() },
    visibility: 'private',
    transparency: 'opaque',
    reminders: { useDefault: false, overrides: [] },
    extendedProperties: { private: { sourceEventId, sourceCalendarId: sourceCalId } },
  }, CONFIG.workCalendarId).id;
}

function linkPersonalEvent(calId, eventId, holdId) {  
  Calendar.Events.patch({    
    extendedProperties: { private: { workHoldId: holdId } },  
  }, calId, eventId);
}

function tryDelete(holdId) {
  try {
    Calendar.Events.remove(CONFIG.workCalendarId, holdId);
  } catch (e) {
    console.error(`Failed to delete hold ${holdId}: ${e.message}`);
  }
}
