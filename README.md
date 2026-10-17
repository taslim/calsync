# CalSync 📅

A Google Apps Script that mirrors the busy time on your personal calendars as `[DNS] External Appointment` holds on your work calendar, so colleagues see that you are busy without seeing why.

## What it does

- Holds the work-hour part of each personal event, in your work calendar's time zone, with one hold per weekday it touches.
- Ignores all-day events, events marked Free, invitations you have not accepted, and weekends.
- Skips holds fully covered by an Out of Office block, and holds that would fill a whole workday (use Out of Office for those).
- Keeps today and the next 28 days in step every 5 minutes: holds appear, move, and disappear with your personal events. Earlier days are left as they were.
- Only reads your personal calendars, and stores nothing outside the two calendars.

## Telling holds from meetings

Every hold's event ID starts with `csdnshold`. Anyone with free/busy access to your work calendar, a person or an agent, gets each event's ID, start, and end when they list its events, so they can set holds aside and treat the rest as work commitments. A free/busy query merges busy time and returns no IDs, so list events instead.

## Setup

1. On your personal account, share the calendar with your work account: Settings › Settings for my calendars › your calendar › Share with specific people. Permission: **See all event details**. Free/busy is not enough.
2. On your work account, open [script.google.com](https://script.google.com/), create a project, and paste `Code.js` into it.
3. Edit `CONFIG` at the top: your personal calendar address(es), work hours, and anything else you want changed.
4. Add the **Google Calendar API** under Services (the `+` in the sidebar).
5. Run `install` from the function dropdown and grant the permissions it asks for (Advanced › Go to project).

Each run logs what it created and removed under Executions. A change that fails marks the run as failed and is retried by the next one.

## Everyday use

- **Change a setting:** edit `CONFIG`, including the list of calendars, and save. The next run applies it to existing holds.
- **Skip one event:** mark it Free, or answer Maybe or No on your personal calendar. Deleting or editing a hold only brings it back on the next run.
- **Upgrade:** paste the new `Code.js` and save. From a version whose holds lack `csdnshold` IDs, run `uninstall` with the old code first.
- **Stop:** run `uninstall`. It removes the trigger and every hold. Rerun it if it reports a failure.

## Development

`node --test` runs `Code.js` against in-memory stand-ins for the Calendar API. Node 20 or newer, no dependencies.
