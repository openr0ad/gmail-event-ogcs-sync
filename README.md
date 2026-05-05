# Gmail Event OGCS Sync

A Google Apps Script that automatically duplicates Gmail-created Google Calendar events so they can be synced to Microsoft Outlook via [Outlook Google Calendar Sync (OGCS)](https://github.com/phw198/OutlookGoogleCalendarSync).

## Background

Google automatically creates calendar events from Gmail (flight bookings, restaurant reservations, etc.) but marks them as `fromGmail` type, which are read-only via the API. OGCS cannot sync these events to Outlook. This script works around that limitation by creating a standard editable duplicate that OGCS can sync.

See the upstream OGCS issue: [#1988](https://github.com/phw198/OutlookGoogleCalendarSync/issues/1988)

## Setup

1. Go to [script.google.com](https://script.google.com) and create a new project
2. Replace the default code with the contents of `Code.gs`
3. Click **Services** → add **Google Calendar API**
4. Click **Run** once to grant permissions
5. Go to **Triggers** → **Add Trigger**:
   - Function: `duplicateGmailEvents`
   - Event source: Time-driven
   - Type: Hour timer (every 1 hour recommended)

## Configuration

At the top of the script, you can change:
- `CALENDAR_ID` — defaults to `'primary'`, change if targeting a different calendar
- `colorId` — the color applied to duplicated events (default: `2` = Sage). See [color reference](https://developers.google.com/calendar/api/v3/reference/colors/get)

## How It Works

- Scans your calendar for events with `eventType: fromGmail` in a rolling 3-month window
- Creates an editable duplicate with the same title, time, and location
- Marks the duplicate's description with `[ogcs-ready]` to prevent re-processing
- Applies a distinct calendar color so you can visually distinguish duplicates from originals

## Known Limitations

- As far as I can tell, the original Gmail event cannot be deleted via the API (Google restricts this) and must be removed manually
- If you accidentally delete a duplicate, the script will re-create it on the next run

## License

MIT
