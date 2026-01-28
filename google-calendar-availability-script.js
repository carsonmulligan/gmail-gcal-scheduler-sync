/**
 * Google Calendar → Google Doc Availability Script
 *
 * Reads all events from specified Google Calendars,
 * computes open time slots, and writes a bullet-point
 * list to a Google Doc. Set to run daily via trigger.
 *
 * SETUP:
 * 1. Go to script.google.com → New Project
 * 2. Paste this script
 * 3. Replace DOC_ID and CALENDAR_IDS below
 * 4. Run updateAvailabilityDoc() once to authorize
 * 5. Add a trigger: Triggers → Add → updateAvailabilityDoc → Time-driven → Day timer
 */

// === CONFIGURATION ===
const DOC_ID = 'YOUR_GOOGLE_DOC_ID_HERE'; // from the doc URL: docs.google.com/document/d/THIS_PART/edit
const CALENDAR_IDS = [
  'primary',                          // default calendar
  // 'teresa@example.com',            // add more calendar IDs as needed
  // 'campaign-calendar-id@group.calendar.google.com',
];
const DAYS_AHEAD = 14;
const DAY_START_HOUR = 8;   // 8 AM
const DAY_END_HOUR = 18;    // 6 PM
const SKIP_WEEKENDS = false; // set true to skip Sat/Sun

// === MAIN ===
function updateAvailabilityDoc() {
  const today = new Date();
  today.setHours(0, 0, 0, 0);

  const endDate = new Date(today);
  endDate.setDate(endDate.getDate() + DAYS_AHEAD);

  // Fetch events from all calendars
  const allEvents = [];
  CALENDAR_IDS.forEach(calId => {
    const cal = CalendarApp.getCalendarById(calId);
    if (!cal) {
      Logger.log('Calendar not found: ' + calId);
      return;
    }
    const events = cal.getEvents(today, endDate);
    events.forEach(e => {
      allEvents.push({
        start: e.getStartTime(),
        end: e.getEndTime(),
        allDay: e.isAllDayEvent()
      });
    });
  });

  // Write the doc
  const doc = DocumentApp.openById(DOC_ID);
  const body = doc.getBody();
  body.clear();

  body.appendParagraph('Available Times')
      .setHeading(DocumentApp.ParagraphHeading.HEADING1);
  body.appendParagraph('Last updated: ' + new Date().toLocaleString())
      .setItalic(true);
  body.appendParagraph('');

  for (let d = 0; d < DAYS_AHEAD; d++) {
    const date = new Date(today);
    date.setDate(date.getDate() + d);

    const dow = date.getDay();
    if (SKIP_WEEKENDS && (dow === 0 || dow === 6)) continue;

    const dayStart = new Date(date);
    dayStart.setHours(DAY_START_HOUR, 0, 0, 0);
    const dayEnd = new Date(date);
    dayEnd.setHours(DAY_END_HOUR, 0, 0, 0);

    // Get busy blocks for this day
    const busy = [];
    allEvents.forEach(e => {
      if (e.allDay && isSameDay(e.start, date)) {
        busy.push({ start: dayStart, end: dayEnd });
        return;
      }
      // Event overlaps this day's working hours?
      const s = e.start < dayStart ? dayStart : e.start;
      const f = e.end > dayEnd ? dayEnd : e.end;
      if (s < f && isSameDay(dayStart, date) &&
          (isSameDay(e.start, date) || isSameDay(e.end, date) || (e.start < dayStart && e.end > dayEnd))) {
        busy.push({ start: new Date(s), end: new Date(f) });
      }
    });

    // Merge overlapping busy blocks
    busy.sort((a, b) => a.start - b.start);
    const merged = [];
    busy.forEach(b => {
      if (merged.length > 0 && b.start <= merged[merged.length - 1].end) {
        merged[merged.length - 1].end = new Date(
          Math.max(merged[merged.length - 1].end.getTime(), b.end.getTime())
        );
      } else {
        merged.push({ start: new Date(b.start), end: new Date(b.end) });
      }
    });

    // Compute open slots
    const slots = [];
    let cursor = new Date(dayStart);
    merged.forEach(b => {
      if (cursor < b.start) {
        slots.push({ start: new Date(cursor), end: new Date(b.start) });
      }
      cursor = b.end > cursor ? new Date(b.end) : cursor;
    });
    if (cursor < dayEnd) {
      slots.push({ start: new Date(cursor), end: new Date(dayEnd) });
    }

    // Write day heading
    const dateStr = Utilities.formatDate(date, Session.getScriptTimeZone(), 'EEEE, MMMM d, yyyy');
    body.appendParagraph(dateStr)
        .setHeading(DocumentApp.ParagraphHeading.HEADING2);

    // Write bullet points
    if (slots.length === 0) {
      body.appendListItem('No availability')
          .setGlyphType(DocumentApp.GlyphType.BULLET);
    } else {
      slots.forEach(slot => {
        const startStr = Utilities.formatDate(slot.start, Session.getScriptTimeZone(), 'h:mm a');
        const endStr = Utilities.formatDate(slot.end, Session.getScriptTimeZone(), 'h:mm a');
        body.appendListItem(startStr + ' – ' + endStr)
            .setGlyphType(DocumentApp.GlyphType.BULLET);
      });
    }
  }

  Logger.log('Doc updated successfully');
}

function isSameDay(date1, date2) {
  return date1.getFullYear() === date2.getFullYear() &&
         date1.getMonth() === date2.getMonth() &&
         date1.getDate() === date2.getDate();
}
