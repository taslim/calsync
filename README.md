# CalSync 📅

A Google Apps Script that mirrors your personal calendar as `[DNS] External Appointment` holds on your work calendar, so colleagues see that you are busy without seeing why.

## What it does

- Clamps each personal event to your work hours, in your work calendar's time zone, with one hold per weekday it touches.
- Ignores all-day events, events marked Free, declined invitations, and weekends.
- Skips holds fully covered by an Out of Office block, and holds that would fill a whole workday (use Out of Office for those).
- Keeps holds in step with their events: they move, shrink, and disappear as the personal event changes. Runs every 5 minutes and stores nothing outside the two calendars.
- Tags each mirrored personal event with the id of its earliest hold in the window (`workHoldId`, a private extended property), so tools that only see your personal calendar can tell holds from real conflicts. Optional.

## Setup

1. On your personal account, share the calendar with your work account: Settings › Settings for my calendars › your calendar › Share with specific people. Permission: **Make changes to events**, or **See all event details** if you turn `tagPersonalEvents` off. Free/busy is not enough.
2. On your work account, open [script.google.com](https://script.google.com/), create a project, and paste `Code.js` into it.
3. Edit `CONFIG` at the top: your personal calendar address(es), work hours, and anything else you want changed.
4. Add the **Google Calendar API** under Services (the `+` in the sidebar).
5. Run `install` from the function dropdown and grant the permissions it asks for (Advanced › Go to project).

Each run logs what it created, updated, and removed under Executions.

## Everyday use

- **Change a setting:** edit `CONFIG` and save. The next run applies it to existing holds.
- **Skip one event:** mark it Free or decline it on your personal calendar. Deleting the hold only brings it back on the next run.
- **Read the tag:** `workHoldId` is present while the event has a hold within the 28-day window and cleared once it needs none.
- **Upgrade:** paste the new `Code.js` and run `install` once. It rebuilds the holds.
- **Stop:** run `uninstall`. It removes the trigger and every hold.

## Development

`node --test` runs `Code.js` against in-memory stand-ins for the Calendar API. Node 20 or newer, no dependencies.
