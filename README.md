# CalSync 📅

A lightweight, stateless Google Apps Script that automatically mirrors your personal calendar events as private holds on your work calendar.

Protect your time without leaking personal details or dealing with messy third-party calendar sync tools.

## Features

* **Privacy First:** Creates opaque `[DNS] External Appointment` holds on your work calendar. By default, these are set to private so colleagues only see that you are busy (configurable).
* **Smart Work Hours:** Automatically clamps personal events to fit within your designated work hours (e.g., 9 AM to 5 PM).
* **Weekend & OOO Detection:** Ignores events on weekends and events that are fully covered by an existing "Out of Office" block on your work calendar.
* **Duration Limits:** Automatically ignores all-day events or events that span longer than a configurable threshold (defaults to 8 hours).
* **Stateless & Resilient:** Uses a rolling 28-day window and private extended properties. It cleans up orphaned holds naturally, ignores infinite recurrences, and safely supports multiple people using CalSync on shared events.

## Prerequisites

1. A Google Work account (where the script will run).
2. One or more Personal Google accounts.
3. Your Personal calendar(s) must be shared with your Work account with **"Make changes to events"** or **"See all event details"** permissions.
   *(Note: Sharing as "See only free/busy" will not work).*

## Installation

1. **Share your Personal Calendar:**

   * Go to Google Calendar on your *personal* account.
   * Go to Settings > Settings for my calendars > `[Your Calendar]`.
   * Scroll to "Share with specific people or groups", add your work email, and grant permission to "Make changes to events" (or at least "See all event details").

2. **Create the Script:**

   * Log into your *work* account and go to [script.google.com](https://script.google.com/).
   * Click **New Project**.
   * Delete any placeholder code and paste the contents of `Code.js` from this repository.

3. **Configure:**

   * At the top of the script, update the `CONFIG` object.
   * Replace `'your.personal@email.com'` with your actual personal calendar email address(es).
   * Adjust `workStartHour`, `workEndHour`, `maxHoldHours`, and `holdVisibility` if necessary.

4. **Enable Calendar API:**

   * On the left sidebar of the Apps Script editor, click on **Services** (the `+` icon).
   * Scroll down, select **Google Calendar API**, and click **Add**.

5. **Run the Setup:**

   * In the toolbar at the top of the editor, select the `install` function from the dropdown menu.
   * Click **Run**.
   * Google will prompt you to review permissions. Click **Review Permissions**, choose your work account, click **Advanced**, and then click Go to `[Project Name]` (unsafe). Click **Allow**.

That's it! CalSync will now run automatically in the background every 5 minutes.

## Uninstalling

If you ever want to stop syncing and remove the generated holds from your work calendar:

1. Open your Apps Script project.
2. Select the `uninstall` function from the top toolbar dropdown.
3. Click **Run**.
   This will delete the 5-minute trigger and sweep your work calendar to remove any active holds it created.
   