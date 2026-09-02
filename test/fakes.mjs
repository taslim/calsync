// In-memory stand-ins for the Apps Script globals Code.js uses (Utilities, Calendar,
// ScriptApp, LockService, console, Date), so the sync logic can run under `node --test`.
import { readFileSync } from 'node:fs';
import vm from 'node:vm';

const SOURCE = readFileSync(new URL('../Code.js', import.meta.url), 'utf8');
const ISO_DAY = { Mon: 1, Tue: 2, Wed: 3, Thu: 4, Fri: 5, Sat: 6, Sun: 7 };

/** Utilities.formatDate for the SimpleDateFormat patterns Code.js relies on. */
export function formatDate(date, tz, pattern) {
  const p = Object.fromEntries(
    new Intl.DateTimeFormat('en-US', {
      timeZone: tz, hourCycle: 'h23', weekday: 'short',
      year: 'numeric', month: '2-digit', day: '2-digit', hour: '2-digit', minute: '2-digit', second: '2-digit',
    }).formatToParts(date).filter(x => x.type !== 'literal').map(x => [x.type, x.value]),
  );
  switch (pattern) {
    case 'u': return String(ISO_DAY[p.weekday]);
    case 'yyyy-MM-dd': return `${p.year}-${p.month}-${p.day}`;
    case 'yyyy-MM-dd HH:mm': return `${p.year}-${p.month}-${p.day} ${p.hour}:${p.minute}`;
    case 'Z': { // RFC 822 offset, e.g. "-0700"
      const wallClockAsUtc = Date.UTC(p.year, p.month - 1, p.day, p.hour, p.minute, p.second);
      const offsetMin = Math.round((wallClockAsUtc - date.getTime()) / 60000);
      const abs = Math.abs(offsetMin);
      return `${offsetMin < 0 ? '-' : '+'}${String(Math.floor(abs / 60)).padStart(2, '0')}${String(abs % 60).padStart(2, '0')}`;
    }
    default: throw new Error(`formatDate pattern not faked: ${pattern}`);
  }
}

const clone = x => JSON.parse(JSON.stringify(x));
const boundary = part => Date.parse(part.dateTime ?? `${part.date}T00:00:00Z`);

/** Calendar advanced service with the list-filter semantics of the real API. */
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
  const live = (cal, id) => {
    const ev = cal.events.get(id);
    if (!ev || ev.status === 'cancelled') throw new Error('Not Found');
    return ev;
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
            // timeMin bounds the END (exclusive), timeMax bounds the START (exclusive), as in the real API.
            if (params.timeMin && !(boundary(ev.end) > Date.parse(params.timeMin))) return false;
            if (params.timeMax && !(boundary(ev.start) < Date.parse(params.timeMax))) return false;
            if (params.eventTypes && !params.eventTypes.includes(ev.eventType ?? 'default')) return false;
            return true;
          });
          cursor = { items, offset: 0 };
        }
        const page = cursor.items.slice(cursor.offset, cursor.offset + pageSize);
        const res = { items: clone(page.filter(ev => ev.status !== 'cancelled')) };
        if (cursor.offset + pageSize < cursor.items.length) {
          res.nextPageToken = `tok${cursors.size + 1}`;
          cursors.set(res.nextPageToken, { items: cursor.items, offset: cursor.offset + pageSize });
        }
        return res;
      },
      insert(resource, calId) {
        calls.push('Events.insert');
        const ev = { ...clone(resource), id: `ev${++seq}` };
        writable(calId).events.set(ev.id, ev);
        return clone(ev);
      },
      patch(resource, calId, id) {
        calls.push('Events.patch');
        const ev = live(writable(calId), id);
        Object.assign(ev, clone(resource));
        return clone(ev);
      },
      remove(calId, id) {
        calls.push('Events.remove');
        live(writable(calId), id).status = 'cancelled';
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
    writes: () => calls.filter(c => c !== 'Events.list' && c !== 'Calendars.get'),
  };
}

/** Loads Code.js into a sandbox wired to the fakes and returns its top-level bindings. */
export function loadCalSync({ calendar, now = () => Date.now(), lockAvailable = true }) {
  const triggers = [];
  const logs = { log: [], warn: [], error: [] };

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
    warn: m => logs.warn.push(String(m)),
    error: m => logs.error.push(String(m)),
  };

  const sandbox = { Utilities: { formatDate }, Calendar: calendar.api, ScriptApp, LockService, console, Date: FakeDate };
  const bindings = vm.runInNewContext(
    `${SOURCE}\n;({ CONFIG, HOLD_TITLE, install, uninstall, sync, workRanges, midnightInTz })`,
    sandbox,
  );
  return { ...bindings, triggers, logs };
}
