/**
 * Gmail Event OGCS Sync
 * 
 * Automatically duplicates Google Calendar events that were auto-created from Gmail
 * (e.g. restaurant reservations, flight bookings) so that they can be synced to
 * Microsoft Outlook via Outlook Google Calendar Sync (OGCS).
 *
 * Background: Google marks these events as `fromGmail` type, making them read-only
 * via the API. OGCS cannot sync read-only events. This script creates a standard
 * editable duplicate that OGCS can sync.
 *
 * Related OGCS issue: https://github.com/phw198/OutlookGoogleCalendarSync/issues/1988
 *
 * Setup:
 *   1. In Apps Script, click Services → add Google Calendar API
 *   2. Run once manually to grant permissions
 *   3. Set a time-driven trigger on duplicateGmailEvents() — every 1 hour recommended
 *
 * Author: openr0ad
 * License: MIT
 */


// ─── Configuration ────────────────────────────────────────────────────────────

/** The calendar to monitor. 'primary' is your main Google Calendar. */
var CALENDAR_ID = 'primary';

/**
 * How many months ahead to scan for fromGmail events.
 * 3 months is a reasonable default for most use cases.
 */
var MONTHS_AHEAD = 3;

/**
 * Color applied to duplicated events so you can visually distinguish them
 * from the original read-only Gmail events in Google Calendar.
 *
 * Color IDs:
 *   1  Lavender    2  Sage       3  Grape      4  Flamingo
 *   5  Banana      6  Tangerine  7  Peacock    8  Graphite
 *   9  Blueberry   10 Basil      11 Tomato
 */
var DUPLICATE_COLOR_ID = '2'; // Sage

/**
 * Tag written to the description of duplicated events.
 * Prevents the script from re-duplicating events on subsequent runs.
 * Not visible unless you open the event details.
 */
var PROCESSED_TAG = '[ogcs-ready]';


// ─── Main Function ─────────────────────────────────────────────────────────────

/**
 * Scans the calendar for fromGmail events and duplicates any that haven't
 * been processed yet. Set this as your time-driven trigger.
 */
function duplicateGmailEvents() {
  var now = new Date();
  var lookahead = new Date();
  lookahead.setMonth(lookahead.getMonth() + MONTHS_AHEAD);

  var response = Calendar.Events.list(CALENDAR_ID, {
    timeMin: now.toISOString(),
    timeMax: lookahead.toISOString(),
    singleEvents: true,
    maxResults: 250
  });

  var events = response.items || [];
  var duplicated = 0;

  events.forEach(function(event) {
    // Only target events auto-created from Gmail
    if (event.eventType !== 'fromGmail') return;

    // Skip if already processed on a previous run
    if ((event.description || '').includes(PROCESSED_TAG)) return;

    duplicateEvent(event);
    duplicated++;
  });

  Logger.log('Run complete. Duplicated ' + duplicated + ' event(s).');
}


// ─── Helper Functions ──────────────────────────────────────────────────────────

/**
 * Creates an editable duplicate of a fromGmail event that OGCS can sync.
 * Applies a distinct color and tags the description to prevent re-processing.
 *
 * @param {Object} event - A Google Calendar event resource object.
 */
function duplicateEvent(event) {
  var newEvent = {
    summary: event.summary,
    description: (event.description || '') + '\n' + PROCESSED_TAG,
    location: event.location || '',
    start: event.start,
    end: event.end,
    colorId: DUPLICATE_COLOR_ID
  };

  Calendar.Events.insert(newEvent, CALENDAR_ID);
  Logger.log('Duplicated: ' + event.summary + ' on ' + (event.start.dateTime || event.start.date));
}

/**
 * Utility: Deletes all calendar events whose title contains a given string.
 * Useful for cleaning up accidental mass-duplications.
 * Change the SEARCH_TERM below and run this function manually as needed.
 */
function deleteEventsByTitle() {
  var SEARCH_TERM = 'Enter title here'; // ← change this before running

  var now = new Date();
  var lookahead = new Date();
  lookahead.setMonth(lookahead.getMonth() + MONTHS_AHEAD);

  var response = Calendar.Events.list(CALENDAR_ID, {
    timeMin: now.toISOString(),
    timeMax: lookahead.toISOString(),
    singleEvents: true,
    maxResults: 500,
    q: SEARCH_TERM
  });

  var events = response.items || [];
  var deleted = 0;

  events.forEach(function(event) {
    if (event.summary && event.summary.includes(SEARCH_TERM)) {
      try {
        Calendar.Events.remove(CALENDAR_ID, event.id);
        deleted++;
        Logger.log('Deleted #' + deleted + ': ' + event.summary + ' on ' + (event.start.dateTime || event.start.date));
        Utilities.sleep(300); // Pause to avoid hitting API rate limits
      } catch(e) {
        Logger.log('Failed to delete: ' + event.summary + ' — ' + e.message);
      }
    }
  });

  Logger.log('Cleanup complete. Deleted ' + deleted + ' event(s).');
}
