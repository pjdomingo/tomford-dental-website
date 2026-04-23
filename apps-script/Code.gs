/**
 * TomFord Dental — Booking System Apps Script v2
 *
 * Sheet tabs required:
 *   Services     → Col A: one service name per row
 *   Config       → Col A: key | Col B: value
 *   Appointments → auto-created on first POST
 *
 * Config keys:
 *   CLINIC_NAME              TomFord Dental
 *   CLINIC_ADDRESS           RB & A BLDG., 166 Lakeview Drive, COR Kawilihan Lane, Pasig Blvd, Pasig, Philippines 1600
 *   CLINIC_PHONE             0995 418 8879
 *   CLINIC_EMAIL             tomford.dental@gmail.com
 *   CLINIC_HOURS             Mon–Sat  9:00 AM – 7:00 PM
 *   CONCIERGE_EMAIL          tomford.dental@gmail.com
 *   BOOKING_TAGLINE          where every tooth matters.
 *   CALENDAR_DURATION_MINS   60
 *   OPEN_TIME                09:00
 *   CLOSE_TIME               19:00
 *   SLOT_DURATION_MINS       30
 *   OPEN_DAYS                1,2,3,4,5,6
 *   MAX_BOOKINGS_PER_SLOT    1
 *   CALENDAR_ID              primary
 *   SLOT_BUFFER_MINS         60
 *   ADVANCE_BOOKING_DAYS     30
 */

function jsonOut(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ── GET ──────────────────────────────────────────────────────
function doGet(e) {
  const action = (e && e.parameter && e.parameter.action) || '';
  var result;
  if      (action === 'getServices') result = getServices();
  else if (action === 'getConfig')   result = getPublicConfig();
  else if (action === 'getSlots')    result = getSlots(e);
  else if (action === 'debug')       result = debugDump(e);
  else                               result = { error: 'Unknown action' };
  return jsonOut(result);
}

// Diagnostic: ?action=debug&date=YYYY-MM-DD
// Returns raw sheet values, their types, and normalized forms so we can
// see exactly where the booked-slot comparison is failing.
function debugDump(e) {
  var dateStr = (e && e.parameter && e.parameter.date) || getTodayPHT();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Appointments');
  var rows = [];
  if (sheet && sheet.getLastRow() > 1) {
    var data = sheet.getRange(2, 1, sheet.getLastRow()-1, 7).getValues();
    data.forEach(function(row, i) {
      rows.push({
        row: i+2,
        rawDate: String(row[5]),
        rawDateType: typeof row[5] + (Object.prototype.toString.call(row[5])==='[object Date]' ? ' (Date)' : ''),
        normalizedDate: normalizeSheetDate(row[5]),
        rawTime: String(row[6]),
        rawTimeType: typeof row[6] + (Object.prototype.toString.call(row[6])==='[object Date]' ? ' (Date)' : ''),
        normalizedTime: normalizeSheetTime(row[6]),
        matchesQueryDate: normalizeSheetDate(row[5]) === dateStr
      });
    });
  }
  var config = getConfigObject();

  // Calendar diagnostic — each step independently so one failure doesn't mask others
  var calDiag = { CALENDAR_ID: config.CALENDAR_ID || '(empty)' };

  // Step 1: who is the script running as?
  try {
    calDiag.scriptOwner = Session.getEffectiveUser().getEmail();
  } catch(err) {
    calDiag.scriptOwnerError = 'Session.getEffectiveUser failed — web app is deployed as "Execute as: User accessing" (anonymous access has no identity). Change Deploy settings → Execute as: Me.';
  }

  // Step 2: can we open the spreadsheet at all?
  try {
    var _ss = SpreadsheetApp.getActiveSpreadsheet();
    calDiag.spreadsheetAccess = _ss ? ('OK — ' + _ss.getName()) : 'FAIL — getActiveSpreadsheet returned null';
  } catch(err) {
    calDiag.spreadsheetAccess = 'FAIL — ' + err;
  }

  // Step 3: resolve calendar by id
  var calId = (config.CALENDAR_ID || '').trim();
  if (!calId) {
    calDiag.status = 'ERROR: CALENDAR_ID is empty — either the Config sheet row is blank OR getConfigObject() returned {} because the spreadsheet could not be read. Check calDiag.spreadsheetAccess above.';
  } else if (calId === 'primary') {
    calDiag.status = 'ERROR: CALENDAR_ID is literal "primary" — set it to the shared clinic calendar id';
  } else {
    try {
      var cal = CalendarApp.getCalendarById(calId);
      if (!cal) {
        calDiag.status = 'ERROR: getCalendarById returned null. Script owner likely lacks "Make changes to events" access on this calendar.';
      } else {
        calDiag.status = 'OK';
        calDiag.calendarName = cal.getName();
        try { calDiag.isOwnedByMe = cal.isOwnedByMe(); } catch(e) { calDiag.isOwnedByMeError = String(e); }
        // Probe write permission by creating + deleting a tiny event.
        try {
          var probe = cal.createEvent('__probe__', new Date(), new Date(Date.now()+60000), { description:'probe' });
          calDiag.canWrite = true;
          probe.deleteEvent();
        } catch (wErr) {
          calDiag.canWrite = false;
          calDiag.writeError = String(wErr);
        }
      }
    } catch(err) {
      calDiag.status = 'EXCEPTION: ' + err;
    }
  }

  return {
    queryDate: dateStr,
    config: {
      OPEN_DAYS_raw: config.OPEN_DAYS,
      OPEN_DAYS_parsed: parseOpenDays(config.OPEN_DAYS || '1,2,3,4,5,6'),
      CALENDAR_ID: config.CALENDAR_ID,
      SLOT_DURATION_MINS: config.SLOT_DURATION_MINS,
      OPEN_TIME: config.OPEN_TIME,
      CLOSE_TIME: config.CLOSE_TIME,
      MAX_BOOKINGS_PER_SLOT: config.MAX_BOOKINGS_PER_SLOT,
      SLOT_BUFFER_MINS: config.SLOT_BUFFER_MINS,
    },
    calendarDiagnostic: calDiag,
    bookedSlotCounts: getBookedSlotCounts(dateStr),
    allRows: rows
  };
}

// ── POST ─────────────────────────────────────────────────────
function doPost(e) {
  var lock = LockService.getScriptLock();
  try {
    // Wait up to 10s for the lock so two simultaneous submits can't
    // both pass the availability check before either writes.
    lock.waitLock(10000);
  } catch (err) {
    return jsonOut({ success: false, error: 'System is busy. Please try again in a moment.' });
  }
  try {
    var data   = e.parameter || {};
    var config = getConfigObject();

    // Required fields
    if (!data.fullName || !data.email || !data.phone || !data.service || !data.date || !data.time) {
      return jsonOut({ success: false, error: 'Please complete all required fields.' });
    }

    // Re-validate the slot is still free (inside the lock — race-safe)
    var check = validateSlotAvailability(data, config);
    if (!check.ok) {
      return jsonOut({ success: false, error: check.error });
    }

    saveAppointment(data);
    createClinicCalendarEvent(data, config); // actual event on clinic calendar
    var calLink = buildCalendarLink(data, config);
    if (data.email) sendPatientEmail(data, config, calLink);
    sendClinicNotification(data, config);
    return jsonOut({ success: true });
  } catch (err) {
    Logger.log('doPost error: ' + err);
    return jsonOut({ success: false, error: err.toString() });
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}

// Server-side availability check: blocks double-booking even if the
// client sent a stale slot from a cached picker.
function validateSlotAvailability(data, config) {
  // 1. Sheet-level check — is this date+time already at capacity?
  var maxPerSlot = parseInt(config.MAX_BOOKINGS_PER_SLOT || '1', 10);
  var booked = getBookedSlotCounts(data.date);
  if ((booked[data.time] || 0) >= maxPerSlot) {
    return { ok: false, error: 'That time slot was just taken. Please pick another.' };
  }

  // 2. Calendar-level check — does the clinic calendar have a conflicting event?
  //    (covers staff-added events and the createClinicCalendarEvent output of
  //    a previous booking that hasn't landed in the sheet yet.)
  try {
    var calendarId = config.CALENDAR_ID || 'primary';
    var cal = CalendarApp.getCalendarById(calendarId);
    if (cal) {
      var slotDur = parseInt(config.SLOT_DURATION_MINS || '30', 10);
      var dp = data.date.split('-').map(Number);
      var tp = data.time.split(':').map(Number);
      var start = new Date(dp[0], dp[1]-1, dp[2], tp[0], tp[1], 0);
      var end   = new Date(start.getTime() + slotDur * 60000);
      var events = cal.getEvents(start, end);
      if (events.length > 0) {
        return { ok: false, error: 'That time conflicts with an existing appointment. Please pick another slot.' };
      }
    }
  } catch (err) {
    // Don't fail the booking on a calendar glitch — sheet check still holds.
    Logger.log('validateSlotAvailability calendar check failed: ' + err);
  }

  // 3. Range sanity — is the time within OPEN_TIME/CLOSE_TIME and on an open day?
  try {
    var openDays = parseOpenDays(config.OPEN_DAYS || '1,2,3,4,5,6');
    var dp2 = data.date.split('-').map(Number);
    var dow = new Date(dp2[0], dp2[1]-1, dp2[2]).getDay();
    if (openDays.indexOf(dow) === -1) {
      return { ok: false, error: 'That day is not a clinic working day.' };
    }
    var tp2 = data.time.split(':').map(Number);
    var tMins = tp2[0]*60 + tp2[1];
    var oArr  = (config.OPEN_TIME || '09:00').split(':').map(Number);
    var cArr  = (config.CLOSE_TIME|| '19:00').split(':').map(Number);
    if (tMins < oArr[0]*60+oArr[1] || tMins >= cArr[0]*60+cArr[1]) {
      return { ok: false, error: 'That time is outside clinic hours.' };
    }
  } catch (err) {
    Logger.log('validateSlotAvailability range check failed: ' + err);
  }

  return { ok: true };
}

// ── getSlots ─────────────────────────────────────────────────
function getSlots(e) {
  try {
    var config   = getConfigObject();
    var fromStr  = (e && e.parameter && e.parameter.from) || getTodayPHT();
    var numDays  = parseInt((e && e.parameter && e.parameter.days) || '14', 10);

    var openTime     = config.OPEN_TIME             || '09:00';
    var closeTime    = config.CLOSE_TIME            || '19:00';
    var slotDur      = parseInt(config.SLOT_DURATION_MINS    || '30', 10);
    var maxPerSlot   = parseInt(config.MAX_BOOKINGS_PER_SLOT || '1',  10);
    var bufferMins   = parseInt(config.SLOT_BUFFER_MINS      || '60', 10);
    var advanceDays  = parseInt(config.ADVANCE_BOOKING_DAYS  || '30', 10);
    var calendarId   = config.CALENDAR_ID || 'primary';
    var openDaysStr  = config.OPEN_DAYS   || '1,2,3,4,5,6';
    var openDays     = parseOpenDays(openDaysStr);

    var nowPHT      = getNowPHT();
    var bufferMs    = bufferMins * 60 * 1000;
    var advanceCutoff = new Date(nowPHT.getTime() + advanceDays * 24 * 60 * 60 * 1000);

    // Parse from date
    var fromParts = fromStr.split('-').map(Number);
    var result    = {};
    var collected = 0;

    for (var offset = 0; offset < numDays + 7; offset++) {
      if (collected >= numDays) break;

      var d = new Date(fromParts[0], fromParts[1]-1, fromParts[2] + offset);
      var dow = d.getDay(); // 0=Sun…6=Sat
      if (openDays.indexOf(dow) === -1) continue; // closed day

      var dateStr = formatDateStr(d);

      // Don't show beyond advance booking window
      if (d > advanceCutoff) { collected++; result[dateStr] = []; continue; }

      // All possible slots this day
      var allSlots = generateSlots(openTime, closeTime, slotDur);

      // Filter past/within-buffer slots
      var cutoff = new Date(nowPHT.getTime() + bufferMs);
      allSlots = allSlots.filter(function(t) {
        var parts = t.split(':').map(Number);
        var slotDt = new Date(d.getFullYear(), d.getMonth(), d.getDate(), parts[0], parts[1], 0);
        return slotDt > cutoff;
      });

      // Booked slots from Appointments sheet
      var booked = getBookedSlotCounts(dateStr);

      // Blocked times from Google Calendar
      var busy = getCalendarBusy(calendarId, d, slotDur);

      var available = allSlots.filter(function(t) {
        // Check sheet bookings
        if ((booked[t] || 0) >= maxPerSlot) return false;

        // Check calendar overlaps
        var parts   = t.split(':').map(Number);
        var slotStart = new Date(d.getFullYear(), d.getMonth(), d.getDate(), parts[0], parts[1], 0);
        var slotEnd   = new Date(slotStart.getTime() + slotDur * 60 * 1000);

        for (var i = 0; i < busy.length; i++) {
          if (slotStart < busy[i].end && slotEnd > busy[i].start) return false;
        }
        return true;
      });

      result[dateStr] = available;
      collected++;
    }

    return { slots: result, slotDuration: slotDur };

  } catch(err) {
    Logger.log('getSlots error: ' + err);
    return { slots: {}, error: err.toString() };
  }
}

// ── helpers: slots ───────────────────────────────────────────
function generateSlots(openTime, closeTime, durationMins) {
  var slots = [];
  var oArr = openTime.split(':').map(Number);
  var cArr = closeTime.split(':').map(Number);
  var cur  = oArr[0]*60 + oArr[1];
  var end  = cArr[0]*60 + cArr[1];
  while (cur < end) {
    slots.push(pad2(Math.floor(cur/60)) + ':' + pad2(cur%60));
    cur += durationMins;
  }
  return slots;
}

function getBookedSlotCounts(dateStr) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Appointments');
  if (!sheet || sheet.getLastRow() <= 1) return {};
  var data = sheet.getRange(2, 1, sheet.getLastRow()-1, 7).getValues();
  var counts = {};
  data.forEach(function(row, i) {
    var rawD = row[5], rawT = row[6];
    var d = normalizeSheetDate(rawD); // F → YYYY-MM-DD
    var t = normalizeSheetTime(rawT); // G → HH:MM
    // Debug: log every read so we can see what Sheets is returning
    Logger.log('[bookings] row '+(i+2)+': rawD=' + JSON.stringify(rawD) + ' (' + typeof rawD + ') → ' + d
             + ' | rawT=' + JSON.stringify(rawT) + ' (' + typeof rawT + ') → ' + t
             + ' | match=' + (d === dateStr));
    if (d === dateStr && t) counts[t] = (counts[t] || 0) + 1;
  });
  Logger.log('[bookings] counts for ' + dateStr + ' = ' + JSON.stringify(counts));
  return counts;
}

// Accept OPEN_DAYS as any of:
//   "1,2,3,4,5,6"          (numeric, 0=Sun .. 6=Sat)
//   "Mon,Tue,Wed,Thu,Fri"  (3-letter names, any case)
//   "Monday,Tuesday,..."   (full names)
// → array of numbers [1,2,3,4,5]
function parseOpenDays(s) {
  var NAMES = {
    sun:0, sunday:0,
    mon:1, monday:1,
    tue:2, tues:2, tuesday:2,
    wed:3, weds:3, wednesday:3,
    thu:4, thur:4, thurs:4, thursday:4,
    fri:5, friday:5,
    sat:6, saturday:6
  };
  return s.split(',').map(function(tok){
    var t = (tok||'').toString().trim().toLowerCase();
    if (!t) return -1;
    if (/^\d+$/.test(t)) return parseInt(t, 10);
    return (t in NAMES) ? NAMES[t] : -1;
  }).filter(function(n){ return n >= 0 && n <= 6; });
}

// Sheets auto-types date/time cells. Normalize any of:
//   Date object, "2026-04-22", "04/22/2026", number (serial), ""
// → "YYYY-MM-DD"
function normalizeSheetDate(v) {
  if (v === null || v === undefined || v === '') return '';
  if (Object.prototype.toString.call(v) === '[object Date]') {
    return v.getFullYear() + '-' + pad2(v.getMonth()+1) + '-' + pad2(v.getDate());
  }
  var s = v.toString().trim();
  if (/^\d{4}-\d{2}-\d{2}$/.test(s)) return s;
  var parsed = new Date(s);
  if (!isNaN(parsed.getTime())) {
    return parsed.getFullYear() + '-' + pad2(parsed.getMonth()+1) + '-' + pad2(parsed.getDate());
  }
  return s;
}

// → "HH:MM" with leading zero.
function normalizeSheetTime(v) {
  if (v === null || v === undefined || v === '') return '';
  if (Object.prototype.toString.call(v) === '[object Date]') {
    return pad2(v.getHours()) + ':' + pad2(v.getMinutes());
  }
  if (typeof v === 'number') {
    // Sheets stores time as fraction of a 24-hour day
    var totalMins = Math.round(v * 24 * 60);
    return pad2(Math.floor(totalMins/60)) + ':' + pad2(totalMins%60);
  }
  var s = v.toString().trim();
  var m = s.match(/^(\d{1,2}):(\d{2})/);
  if (m) return pad2(parseInt(m[1],10)) + ':' + pad2(parseInt(m[2],10));
  return s;
}

function getCalendarBusy(calendarId, date, slotDurMins) {
  try {
    var cal = CalendarApp.getCalendarById(calendarId);
    if (!cal) { Logger.log('Calendar not found: ' + calendarId); return []; }
    var dayStart = new Date(date.getFullYear(), date.getMonth(), date.getDate(), 0, 0, 0);
    var dayEnd   = new Date(date.getFullYear(), date.getMonth(), date.getDate(), 23, 59, 59);
    var events = cal.getEvents(dayStart, dayEnd);
    return events.map(function(ev) {
      return { start: ev.getStartTime(), end: ev.getEndTime() };
    });
  } catch(err) {
    Logger.log('Calendar error: ' + err);
    return [];
  }
}

// ── helpers: date/time ───────────────────────────────────────
function getTodayPHT() {
  var now = new Date();
  var phtOffset = 8 * 60; // UTC+8
  var utc = now.getTime() + now.getTimezoneOffset() * 60000;
  var pht = new Date(utc + phtOffset * 60000);
  return new Date(pht.getFullYear(), pht.getMonth(), pht.getDate(), 0, 0, 0);
}

function getNowPHT() {
  var now = new Date();
  var phtOffset = 8 * 60;
  var utc = now.getTime() + now.getTimezoneOffset() * 60000;
  return new Date(utc + phtOffset * 60000);
}

function formatDateStr(d) {
  return d.getFullYear() + '-' + pad2(d.getMonth()+1) + '-' + pad2(d.getDate());
}

function pad2(n) { return String(n).padStart(2,'0'); }

function formatDisplayDate(dateStr) {
  if (!dateStr) return '—';
  try {
    var p = dateStr.split('-').map(Number);
    var d = new Date(p[0], p[1]-1, p[2]);
    var DAYS   = ['Sunday','Monday','Tuesday','Wednesday','Thursday','Friday','Saturday'];
    var MONTHS = ['January','February','March','April','May','June','July','August','September','October','November','December'];
    return DAYS[d.getDay()] + ', ' + MONTHS[p[1]-1] + ' ' + p[2] + ', ' + p[0];
  } catch(e) { return dateStr; }
}

function formatDisplayTime(timeStr) {
  if (!timeStr) return '—';
  try {
    var p = timeStr.split(':').map(Number);
    var h = p[0], m = p[1];
    return (h%12||12) + ':' + pad2(m) + ' ' + (h>=12?'PM':'AM');
  } catch(e) { return timeStr; }
}

// ── getServices ──────────────────────────────────────────────
function getServices() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('Services');
    if (!sheet || sheet.getLastRow()===0) return { services:[] };
    var rows = sheet.getRange(1,1,sheet.getLastRow(),1).getValues();
    return { services: rows.map(function(r){return (r[0]||'').toString().trim();}).filter(Boolean) };
  } catch(err) { return { services:[], error:err.toString() }; }
}

// ── getPublicConfig ───────────────────────────────────────────
function getPublicConfig() {
  var c = getConfigObject();
  return { config: {
    CLINIC_NAME:    c.CLINIC_NAME    || 'TomFord Dental',
    CLINIC_ADDRESS: c.CLINIC_ADDRESS || '',
    CLINIC_PHONE:   c.CLINIC_PHONE   || '',
    CLINIC_EMAIL:   c.CLINIC_EMAIL   || '',
    CLINIC_HOURS:   c.CLINIC_HOURS   || '',
    BOOKING_TAGLINE:c.BOOKING_TAGLINE|| 'where every tooth matters.',
    SLOT_DURATION_MINS: c.SLOT_DURATION_MINS || '30',
    OPEN_TIME:      c.OPEN_TIME      || '09:00',
    CLOSE_TIME:     c.CLOSE_TIME     || '19:00',
  }};
}

// ── getConfigObject (internal) ────────────────────────────────
function getConfigObject() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('Config');
    if (!sheet || sheet.getLastRow()===0) return {};
    var rows = sheet.getRange(1,1,sheet.getLastRow(),2).getValues();
    var cfg = {};
    rows.forEach(function(row){
      var k = (row[0]||'').toString().trim();
      var v = (row[1]||'').toString().trim();
      if (k) cfg[k] = v;
    });
    return cfg;
  } catch(err) { return {}; }
}

// ── saveAppointment ──────────────────────────────────────────
function saveAppointment(data) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Appointments');
  if (!sheet) {
    sheet = ss.insertSheet('Appointments');
    sheet.appendRow(['Timestamp','Full Name','Email','Phone','Service','Date','Time','Notes']);
    var h = sheet.getRange(1,1,1,8);
    h.setFontWeight('bold'); h.setBackground('#f46709'); h.setFontColor('#ffffff');
  }
  sheet.appendRow([
    new Date().toLocaleString('en-PH',{timeZone:'Asia/Manila'}),
    data.fullName||'', data.email||'', data.phone||'',
    data.service||'',  data.date||'',  data.time||'', data.notes||''
  ]);
  // Force Date (col F) and Time (col G) columns to plain text so Sheets
  // doesn't auto-convert "2026-04-22" → Date obj or "09:00" → time number.
  // This keeps the stored value identical to what getSlots generates and
  // makes the booked-slot comparison reliable.
  var lastRow = sheet.getLastRow();
  sheet.getRange(lastRow, 6, 1, 2).setNumberFormat('@');
  sheet.getRange(lastRow, 6).setValue(data.date || '');
  sheet.getRange(lastRow, 7).setValue(data.time || '');
}

// ── createClinicCalendarEvent ────────────────────────────────
// Creates an actual event on the clinic's Google Calendar so:
//   1. Staff can see bookings in Google Calendar
//   2. getCalendarBusy() will block these times on future slot lookups
//
// IMPORTANT: NO fallback to 'primary'. If CALENDAR_ID is missing or
// the script owner lacks access to the clinic calendar, we log loudly
// and skip — we do NOT want to accidentally pollute someone's personal
// calendar.
function createClinicCalendarEvent(data, config) {
  if (!data.date || !data.time) return;

  var calendarId = (config.CALENDAR_ID || '').trim();
  if (!calendarId) {
    Logger.log('[calendar] CALENDAR_ID is empty in Config — skipping event creation.');
    return;
  }
  if (calendarId === 'primary') {
    Logger.log('[calendar] CALENDAR_ID is "primary" — refusing to write to owner primary. Set it to the shared clinic calendar ID.');
    return;
  }

  try {
    var cal = CalendarApp.getCalendarById(calendarId);
    if (!cal) {
      Logger.log('[calendar] getCalendarById returned null for "' + calendarId + '". '
               + 'The script owner (' + Session.getEffectiveUser().getEmail() + ') likely does not have '
               + '"Make changes to events" access on that calendar. Fix: open Google Calendar → '
               + 'clinic calendar settings → Share with specific people → add that account with '
               + '"Make changes to events" permission. Then re-run the script once to re-authorize.');
      return;
    }

    var dur = parseInt(config.CALENDAR_DURATION_MINS || '60', 10);
    var dp  = data.date.split('-').map(Number);
    var tp  = data.time.split(':').map(Number);
    var start = new Date(dp[0], dp[1]-1, dp[2], tp[0], tp[1], 0);
    var end   = new Date(start.getTime() + dur * 60000);

    var title = (data.service || 'Appointment') + ' — ' + (data.fullName || 'Patient');
    var desc  = 'Patient: ' + (data.fullName || '—')
              + '\nEmail: '  + (data.email    || '—')
              + '\nPhone: '  + (data.phone    || '—')
              + '\nService: '+ (data.service  || '—')
              + '\n\nStatus: Awaiting confirmation'
              + '\nBooked via website at ' + new Date().toLocaleString('en-PH',{timeZone:'Asia/Manila'}) + ' PHT';

    var ev = cal.createEvent(title, start, end, {
      description: desc,
      location:    config.CLINIC_ADDRESS || '',
    });
    Logger.log('[calendar] ✓ Event created on calendar "' + cal.getName() + '" (' + calendarId + '): ' + ev.getId());
  } catch(err) {
    Logger.log('[calendar] createClinicCalendarEvent error: ' + err);
    // Non-fatal — sheet row + emails still land. Log and continue.
  }
}

// ── buildCalendarLink ─────────────────────────────────────────
function buildCalendarLink(data, config) {
  try {
    var name   = config.CLINIC_NAME   || 'TomFord Dental';
    var addr   = config.CLINIC_ADDRESS|| '';
    var phone  = config.CLINIC_PHONE  || '';
    var dur    = parseInt(config.CALENDAR_DURATION_MINS||'60',10);
    if (!data.date||!data.time) return '';
    var dp = data.date.split('-').map(Number);
    var tp = data.time.split(':').map(Number);
    var s  = new Date(dp[0],dp[1]-1,dp[2],tp[0],tp[1],0);
    var en = new Date(s.getTime()+dur*60000);
    function fmt(d){ return d.getFullYear()+pad2(d.getMonth()+1)+pad2(d.getDate())+'T'+pad2(d.getHours())+pad2(d.getMinutes())+'00'; }
    return 'https://calendar.google.com/calendar/render?action=TEMPLATE'
      +'&text='+encodeURIComponent(name+' — '+(data.service||'Appointment'))
      +'&dates='+encodeURIComponent(fmt(s)+'/'+fmt(en))
      +'&details='+encodeURIComponent('Service: '+(data.service||'')+'\nClinic: '+name+'\nAddress: '+addr+'\nPhone: '+phone+'\n\nAwaiting confirmation from clinic.')
      +'&location='+encodeURIComponent(addr)+'&sf=true';
  } catch(e){ return ''; }
}

// ── sendPatientEmail ──────────────────────────────────────────
function sendPatientEmail(data, config, calLink) {
  var name  = config.CLINIC_NAME    ||'TomFord Dental';
  var addr  = config.CLINIC_ADDRESS ||'';
  var phone = config.CLINIC_PHONE   ||'';
  var hours = config.CLINIC_HOURS   ||'';
  var email = config.CLINIC_EMAIL   ||'';
  var tag   = config.BOOKING_TAGLINE||'where every tooth matters.';
  var dd    = formatDisplayDate(data.date);
  var dt    = formatDisplayTime(data.time);
  var subj  = 'Your '+name+' Appointment — '+dd;

  var cal = calLink
    ? '<tr><td align="center" style="padding:0 0 24px;"><a href="'+calLink+'" style="display:inline-block;background:#f46709;color:#fff;text-decoration:none;padding:13px 30px;border-radius:50px;font-size:14px;font-weight:600;font-family:Arial,sans-serif;">&#128197;&nbsp; Add to Google Calendar</a></td></tr>'
    : '';

  var html =
    '<!DOCTYPE html><html><body style="margin:0;padding:0;background:#f2f2f2;font-family:Arial,sans-serif;">'
    +'<table width="100%" cellpadding="0" cellspacing="0" style="background:#f2f2f2;padding:40px 16px;"><tr><td align="center">'
    +'<table width="560" cellpadding="0" cellspacing="0" style="max-width:560px;width:100%;background:#fff;border-radius:16px;overflow:hidden;box-shadow:0 4px 24px rgba(0,0,0,.08);">'
    // header
    +'<tr><td style="background:linear-gradient(135deg,#f46709,#e05800);padding:32px 36px;text-align:center;">'
    +'<p style="margin:0;color:rgba(255,255,255,.7);font-size:11px;letter-spacing:.3em;text-transform:uppercase;">Appointment Request</p>'
    +'<h1 style="margin:7px 0 3px;color:#fff;font-size:24px;font-family:Georgia,serif;">'+name+'</h1>'
    +'<p style="margin:0;color:rgba(255,255,255,.85);font-size:12px;font-style:italic;font-family:Georgia,serif;">'+tag+'</p>'
    +'</td></tr>'
    // body
    +'<tr><td style="padding:32px 36px 24px;">'
    +'<p style="margin:0 0 18px;color:#162b57;font-size:15px;">Hi <strong>'+(data.fullName||'there')+'</strong>,</p>'
    +'<p style="margin:0 0 22px;color:#4a5568;font-size:14px;line-height:1.7;">We\'ve received your appointment request. Our team will confirm your slot shortly.</p>'
    // details card
    +'<table width="100%" cellpadding="0" cellspacing="0" style="background:#f9f9f9;border-radius:10px;border-left:4px solid #f46709;margin-bottom:22px;">'
    +'<tr><td style="padding:20px 24px;">'
    +'<p style="margin:0 0 12px;color:#f46709;font-size:11px;letter-spacing:.25em;text-transform:uppercase;font-weight:700;">Appointment Details</p>'
    +'<table width="100%" cellpadding="5" cellspacing="0">'
    +'<tr><td style="color:#6b7280;font-size:11px;text-transform:uppercase;letter-spacing:.1em;width:90px;">Service</td><td style="color:#162b57;font-size:14px;font-weight:700;">'+(data.service||'—')+'</td></tr>'
    +'<tr><td style="color:#6b7280;font-size:11px;text-transform:uppercase;letter-spacing:.1em;">Date</td><td style="color:#162b57;font-size:14px;font-weight:700;">'+dd+'</td></tr>'
    +'<tr><td style="color:#6b7280;font-size:11px;text-transform:uppercase;letter-spacing:.1em;">Time</td><td style="color:#162b57;font-size:14px;font-weight:700;">'+dt+'</td></tr>'
    +'<tr><td style="color:#6b7280;font-size:11px;text-transform:uppercase;letter-spacing:.1em;">Phone</td><td style="color:#162b57;font-size:13px;">'+(data.phone||'—')+'</td></tr>'
    +'</table></td></tr></table>'
    +'<table width="100%" cellpadding="0" cellspacing="0">'+cal+'</table>'
    +'<div style="background:#fff8f0;border-radius:8px;padding:13px 16px;margin-bottom:20px;">'
    +'<p style="margin:0;color:#f46709;font-size:13px;"><strong>⏳ Note:</strong> This is a <em>request</em> — not a confirmed booking. We\'ll reach out to confirm.</p>'
    +'</div>'
    +'<p style="margin:0;color:#4a5568;font-size:13px;">Questions? <strong style="color:#162b57;">'+phone+'</strong>'+(email?' &nbsp;·&nbsp; <a href="mailto:'+email+'" style="color:#f46709;text-decoration:none;">'+email+'</a>':'')+'</p>'
    +'</td></tr>'
    // footer
    +'<tr><td style="background:#f9f9f9;border-top:1px solid rgba(22,43,87,.08);padding:18px 36px;text-align:center;">'
    +'<p style="margin:0 0 3px;color:#162b57;font-size:13px;font-weight:700;">'+name+'</p>'
    +'<p style="margin:0 0 2px;color:#6b7280;font-size:11px;">'+addr+'</p>'
    +'<p style="margin:0;color:#6b7280;font-size:11px;">'+hours+'</p>'
    +'</td></tr>'
    +'</table>'
    +'<p style="margin:12px 0 0;color:#9ca3af;font-size:11px;text-align:center;">Sent because you submitted a booking at '+name+'.</p>'
    +'</td></tr></table></body></html>';

  GmailApp.sendEmail(data.email, subj,
    'Hi '+(data.fullName||'there')+', your request for '+(data.service||'an appointment')+' on '+dd+' at '+dt+' is received. We\'ll confirm shortly.'+(calLink?' Add to calendar: '+calLink:''),
    { htmlBody:html, name:name, replyTo:email||undefined }
  );
}

// ── sendClinicNotification ────────────────────────────────────
function sendClinicNotification(data, config) {
  var to = config.CONCIERGE_EMAIL || config.CLINIC_EMAIL;
  if (!to) { Logger.log('No CONCIERGE_EMAIL — skipping.'); return; }
  var name = config.CLINIC_NAME||'TomFord Dental';
  var dd   = formatDisplayDate(data.date);
  var dt   = formatDisplayTime(data.time);
  var subj = '📋 New Booking: '+(data.fullName||'?')+' — '+(data.service||'—')+' — '+dd;

  var html =
    '<!DOCTYPE html><html><body style="font-family:Arial,sans-serif;background:#f2f2f2;padding:20px;">'
    +'<table width="100%" cellpadding="0" cellspacing="0" style="max-width:520px;margin:0 auto;background:#fff;border-radius:12px;overflow:hidden;box-shadow:0 2px 12px rgba(0,0,0,.08);">'
    +'<tr><td style="background:linear-gradient(135deg,#162b57,#1a3565);padding:20px 26px;">'
    +'<p style="margin:0;color:rgba(255,255,255,.5);font-size:11px;letter-spacing:.2em;text-transform:uppercase;">New Appointment</p>'
    +'<h2 style="margin:5px 0 0;color:#fff;font-size:18px;">'+name+' Concierge</h2>'
    +'</td></tr>'
    +'<tr><td style="padding:24px 26px;">'
    +'<table width="100%" cellpadding="8" cellspacing="0" style="border-collapse:collapse;">'
    +'<tr style="background:#f9f9f9;"><td style="color:#6b7280;font-size:11px;text-transform:uppercase;letter-spacing:.1em;width:100px;">Patient</td><td style="color:#162b57;font-size:14px;font-weight:700;">'+(data.fullName||'—')+'</td></tr>'
    +'<tr><td style="color:#6b7280;font-size:11px;text-transform:uppercase;letter-spacing:.1em;">Email</td><td><a href="mailto:'+(data.email||'')+'" style="color:#f46709;text-decoration:none;font-size:13px;">'+(data.email||'—')+'</a></td></tr>'
    +'<tr style="background:#f9f9f9;"><td style="color:#6b7280;font-size:11px;text-transform:uppercase;letter-spacing:.1em;">Phone</td><td style="color:#162b57;font-size:13px;">'+(data.phone||'—')+'</td></tr>'
    +'<tr><td style="color:#6b7280;font-size:11px;text-transform:uppercase;letter-spacing:.1em;">Service</td><td><span style="background:#fff8f0;color:#f46709;font-size:13px;font-weight:700;padding:3px 11px;border-radius:20px;border:1px solid rgba(244,103,9,.2);">'+(data.service||'—')+'</span></td></tr>'
    +'<tr style="background:#f9f9f9;"><td style="color:#6b7280;font-size:11px;text-transform:uppercase;letter-spacing:.1em;">Date</td><td style="color:#162b57;font-size:14px;font-weight:700;">'+dd+'</td></tr>'
    +'<tr><td style="color:#6b7280;font-size:11px;text-transform:uppercase;letter-spacing:.1em;">Time</td><td style="color:#162b57;font-size:14px;font-weight:700;">'+dt+'</td></tr>'
    +'</table>'
    +'<div style="margin-top:18px;padding:12px 15px;background:#fff8f0;border-radius:8px;border-left:3px solid #f46709;">'
    +'<p style="margin:0;color:#f46709;font-size:12px;font-weight:700;text-transform:uppercase;letter-spacing:.1em;">Action Required</p>'
    +'<p style="margin:4px 0 0;color:#4a5568;font-size:13px;">Please confirm this slot and contact the patient.</p>'
    +'</div></td></tr>'
    +'<tr><td style="background:#f9f9f9;border-top:1px solid rgba(22,43,87,.06);padding:13px 26px;text-align:center;">'
    +'<p style="margin:0;color:#9ca3af;font-size:11px;">Submitted '+new Date().toLocaleString('en-PH',{timeZone:'Asia/Manila'})+' PHT</p>'
    +'</td></tr>'
    +'</table></body></html>';

  GmailApp.sendEmail(to, subj,
    'New booking from '+(data.fullName||'?')+' for '+(data.service||'—')+' on '+dd+' at '+dt+'. Email: '+(data.email||'—')+'. Phone: '+(data.phone||'—')+'.',
    { htmlBody:html, name:name+' Booking System' }
  );
}
