// In-memory stand-ins for the Apps Script globals Code.js uses (Utilities, Calendar,
// ScriptApp, LockService, console, Date), so the sync logic can run under `node --test`.
import { randomUUID } from 'node:crypto';
import { readFileSync } from 'node:fs';
import vm from 'node:vm';

const SOURCE = readFileSync(new URL('../Code.js', import.meta.url), 'utf8');

/** Utilities.formatDate for the patterns Code.js and the tests use. */
export function formatDate(date, tz, pattern) {
  const p = wallClockParts(date.getTime(), tz);
  switch (pattern) {
    case 'yyyy-MM-dd': return `${p.year}-${p.month}-${p.day}`;
    case 'yyyy-MM-dd HH:mm': return `${p.year}-${p.month}-${p.day} ${p.hour}:${p.minute}`;
    default: throw new Error(`formatDate pattern not faked: ${pattern}`);
  }
}

/** Utilities.parseDate for the one pattern Code.js uses: an hour on a day's wall clock in tz. */
export function parseDate(text, tz, pattern) {
  const match = /^(\d{4})-(\d{2})-(\d{2}) (\d{1,2})$/.exec(text);
  if (pattern !== 'yyyy-MM-dd H' || !match) throw new Error(`parseDate not faked for ${pattern}: ${text}`);
  const [, year, month, day, hour] = match.map(Number);
  const wall = Date.UTC(year, month - 1, day, hour); // Hour 24 rolls over to the next day, as lenient parsing does
  // The offset can change between the wall-clock reading and the instant (a DST switch), so re-read it at the estimate.
  const estimate = wall - offsetMs(wall, tz);
  return new Date(wall - offsetMs(estimate, tz));
}

function wallClockParts(ms, tz) {
  return Object.fromEntries(
    new Intl.DateTimeFormat('en-US', {
      timeZone: tz, hourCycle: 'h23', year: 'numeric', month: '2-digit', day: '2-digit', hour: '2-digit', minute: '2-digit', second: '2-digit',
    }).formatToParts(new Date(ms)).filter(x => x.type !== 'literal').map(x => [x.type, x.value]),
  );
}

function offsetMs(ms, tz) {
  const p = wallClockParts(ms, tz);
  return Date.UTC(p.year, p.month - 1, p.day, p.hour, p.minute, p.second) - ms;
}

const clone = x => JSON.parse(JSON.stringify(x));
const boundary = part => Date.parse(part.dateTime ?? `${part.date}T00:00:00Z`);

/** The item fields a `fields` parameter asks for: 'nextPageToken,items(id,start)' → ['id', 'start']. */
const itemFields = fields => /items\(([^)]*)\)/.exec(fields)[1].split(',');
const pick = (ev, fields) => Object.fromEntries(fields.filter(f => f in ev).map(f => [f, clone(ev[f])]));

/** Calendar advanced service with the list, ID, and field semantics of the real API. */
export function fakeCalendar({ pageSize = 2 } = {}) {
  const calendars = new Map();
  const cursors = new Map();
  const calls = [];
  let seq = 0;

  const get = id => {
    const cal = calendars.get(id);
    if (!cal) throw new Error('Not Found');
    return cal;
  };
  const writable = id => {
    const cal = get(id);
    if (cal.readOnly) throw new Error('Forbidden');
    return cal;
  };

  const api = {
    Calendars: {
      get(id) { calls.push('Calendars.get'); return { timeZone: get(id).timeZone }; },
    },
    Events: {
      list(calId, params = {}) {
        calls.push('Events.list');
        // Page tokens are cursors into the result set as it stood on the first call, so
        // deleting items while paginating cannot skip any, as with the real API.
        let cursor = cursors.get(params.pageToken);
        if (!cursor) {
          const items = [...get(calId).events.values()].filter(ev => {
            if (ev.status === 'cancelled') return false;
            // With singleEvents a series appears as its occurrences, otherwise as the series itself.
            if (params.singleEvents ? ev.recurrence : ev.recurringEventId) return false;
            // timeMin bounds the END (exclusive), timeMax bounds the START (exclusive), as in the real API.
            if (params.timeMin && !(boundary(ev.end) > Date.parse(params.timeMin))) return false;
            if (params.timeMax && !(boundary(ev.start) < Date.parse(params.timeMax))) return false;
            if (params.eventTypes && !params.eventTypes.includes(ev.eventType ?? 'default')) return false;
            return true;
          });
          cursor = { items, offset: 0 };
        }
        const page = cursor.items.slice(cursor.offset, cursor.offset + pageSize).filter(ev => ev.status !== 'cancelled');
        const fields = params.fields && itemFields(params.fields);
        const res = { items: page.map(ev => (fields ? pick(ev, fields) : clone(ev))) };
        // A `fields` parameter that leaves out nextPageToken drops it, and with it every later page.
        const tokenRequested = !params.fields || params.fields.split(',').includes('nextPageToken');
        if (cursor.offset + pageSize < cursor.items.length && tokenRequested) {
          res.nextPageToken = `tok${cursors.size + 1}`;
          cursors.set(res.nextPageToken, { items: cursor.items, offset: cursor.offset + pageSize });
        }
        return res;
      },
      insert(resource, calId) {
        calls.push('Events.insert');
        const cal = writable(calId);
        const ev = { ...clone(resource), id: resource.id ?? `ev${++seq}` };
        if (resource.id && !/^[a-v0-9]{5,1024}$/.test(ev.id)) throw new Error('Invalid resource id value.');
        if (cal.events.has(ev.id)) throw new Error('The requested identifier already exists.'); // Deleted events keep their IDs
        if (ev.visibility === 'default') delete ev.visibility; // The API leaves default visibility out
        cal.events.set(ev.id, ev);
        return clone(ev);
      },
      remove(calId, id) {
        calls.push('Events.remove');
        const ev = writable(calId).events.get(id);
        if (!ev || ev.status === 'cancelled') throw new Error('Not Found');
        ev.status = 'cancelled';
      },
    },
  };

  return {
    api,
    calls,
    addCalendar(id, { timeZone = 'America/Los_Angeles', readOnly = false } = {}) {
      calendars.set(id, { timeZone, readOnly, events: new Map() });
    },
    addEvent(calId, ev) {
      const full = { id: `ev${++seq}`, ...clone(ev) };
      get(calId).events.set(full.id, full);
      return full;
    },
    /** Live (not cancelled) events on a calendar. */
    live: calId => [...get(calId).events.values()].filter(ev => ev.status !== 'cancelled'),
    writes: () => calls.filter(c => !['Events.list', 'Calendars.get'].includes(c)),
  };
}

/** Loads Code.js into a sandbox wired to the fakes and returns its top-level bindings. */
export function loadCalSync({ calendar, now = () => Date.now(), lockAvailable = true }) {
  const triggers = [];
  const logs = { log: [], error: [] };

  class FakeDate extends Date {
    constructor(...args) { super(...(args.length ? args : [FakeDate.now()])); }
    static now() { return now(); }
  }

  const ScriptApp = {
    newTrigger: fn => ({ timeBased: () => ({ everyMinutes: minutes => ({
      create() { triggers.push({ getHandlerFunction: () => fn, minutes }); },
    }) }) }),
    getProjectTriggers: () => [...triggers],
    deleteTrigger(t) { triggers.splice(triggers.indexOf(t), 1); },
  };
  const LockService = {
    getScriptLock: () => ({ tryLock: () => lockAvailable, releaseLock() {} }),
  };
  const console = {
    log: m => logs.log.push(String(m)),
    error: m => logs.error.push(String(m)),
  };

  const Utilities = { formatDate, parseDate, getUuid: randomUUID };
  const sandbox = { Utilities, Calendar: calendar.api, ScriptApp, LockService, console, Date: FakeDate };
  const bindings = vm.runInNewContext(
    `${SOURCE}\n;({ CONFIG, HOLD_TITLE, HOLD_ID_PREFIX, install, uninstall, sync, workHours })`,
    sandbox,
  );
  return { ...bindings, triggers, logs };
}
