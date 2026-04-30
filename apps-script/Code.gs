/**
 * TomFord Dental — Booking System Apps Script v3
 *
 * Sheet tabs required:
 *   Services     → Col A: one service name per row
 *   Config       → Col A: key | Col B: value
 *   Appointments → auto-created; 10 columns (see saveAppointment)
 *
 * Config keys (existing + new):
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
 *   ADMIN_TOKEN              (set a secret password here — used by the admin panel)
 *   ADMIN_URL                https://tomforddental.com/admin
 *
 * What changed in v3:
 *   - Bookings save as "Pending" (no calendar event yet)
 *   - Clinic gets a "Review in Admin" email with patient details
 *   - Patient gets a "Request Received" email (not a confirmation)
 *   - Admin GET:  getPending, getAll
 *   - Admin POST: adminAction=approve, adminAction=reject
 *   - On approve → calendar event created + patient confirmation sent
 *   - On reject  → patient rejection email sent
 *   - Slot availability counts Pending + Approved (not Rejected)
 *   - Old rows without a Status value are treated as Approved (legacy)
 *
 * Appointments sheet columns:
 *   A  Timestamp   B  Full Name   C  Email      D  Phone
 *   E  Service     F  Date        G  Time        H  Notes
 *   I  Status      J  Booking ID
 */

// ─────────────────────────────────────────────────────────────
// Helpers
// ─────────────────────────────────────────────────────────────

function jsonOut(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function generateBookingId() {
  return 'TFD-' + new Date().getTime().toString(36).toUpperCase()
       + '-' + Math.random().toString(36).substr(2, 4).toUpperCase();
}

function verifyAdminToken(token) {
  if (!token) return false;
  var cfg = getConfigObject();
  var stored = (cfg.ADMIN_TOKEN || '').trim();
  return stored && token === stored;
}

// ─────────────────────────────────────────────────────────────
// GET router
// ─────────────────────────────────────────────────────────────

function doGet(e) {
  var action = (e && e.parameter && e.parameter.action) || '';
  var token  = (e && e.parameter && e.parameter.token)  || '';
  var result;

  if      (action === 'getServices')   result = getServices();
  else if (action === 'getConfig')     result = getPublicConfig();
  else if (action === 'getSlots')      result = getSlots(e);
  else if (action === 'debug')         result = debugDump(e);
  else if (action === 'getPending') {
    if (!verifyAdminToken(token)) return jsonOut({ error: 'Unauthorized' });
    result = getBookingsByStatus(['Pending']);
  }
  else if (action === 'getAll') {
    if (!verifyAdminToken(token)) return jsonOut({ error: 'Unauthorized' });
    result = getAllBookings();
  }
  else result = { error: 'Unknown action' };

  return jsonOut(result);
}

// ─────────────────────────────────────────────────────────────
// POST router
// ─────────────────────────────────────────────────────────────

function doPost(e) {
  var data = e.parameter || {};

  // ── Admin actions ──────────────────────────────────────────
  if (data.adminAction) {
    if (!verifyAdminToken(data.token)) {
      return jsonOut({ success: false, error: 'Unauthorized' });
    }
    if (data.adminAction === 'approve') return handleApproveBooking(data);
    if (data.adminAction === 'reject')  return handleRejectBooking(data);
    if (data.adminAction === 'update')  return handleUpdateBooking(data);
    return jsonOut({ success: false, error: 'Unknown admin action' });
  }

  // ── Patient booking submission ─────────────────────────────
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
  } catch (err) {
    return jsonOut({ success: false, error: 'System is busy. Please try again in a moment.' });
  }
  try {
    var config = getConfigObject();

    if (!data.fullName || !data.email || !data.phone || !data.service || !data.date || !data.time) {
      return jsonOut({ success: false, error: 'Please complete all required fields.' });
    }

    var check = validateSlotAvailability(data, config);
    if (!check.ok) {
      return jsonOut({ success: false, error: check.error });
    }

    var bookingId = generateBookingId();
    saveAppointment(data, bookingId, 'Pending');
    sendPatientPendingEmail(data, config);
    sendClinicRequestEmail(data, config, bookingId);

    return jsonOut({ success: true, bookingId: bookingId });
  } catch (err) {
    Logger.log('doPost error: ' + err);
    return jsonOut({ success: false, error: err.toString() });
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}

// ─────────────────────────────────────────────────────────────
// Admin: approve a booking
// ─────────────────────────────────────────────────────────────

function handleApproveBooking(data) {
  var bookingId = data.id;
  if (!bookingId) return jsonOut({ success: false, error: 'Missing booking id.' });

  var lock = LockService.getScriptLock();
  try { lock.waitLock(10000); } catch(e) {
    return jsonOut({ success: false, error: 'System busy. Try again.' });
  }
  try {
    var ss    = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('Appointments');
    if (!sheet || sheet.getLastRow() <= 1) return jsonOut({ success: false, error: 'No bookings found.' });

    var numRows = sheet.getLastRow() - 1;
    var rows    = sheet.getRange(2, 1, numRows, 10).getValues();
    var rowIdx  = -1;

    for (var i = 0; i < rows.length; i++) {
      if (rows[i][9] === bookingId) { rowIdx = i; break; }
    }
    if (rowIdx === -1) return jsonOut({ success: false, error: 'Booking not found.' });

    var row    = rows[rowIdx];
    var status = (row[8] || '').toString().trim();
    if (status === 'Approved') return jsonOut({ success: false, error: 'Already approved.' });

    // Update status column (I = col 9)
    sheet.getRange(rowIdx + 2, 9).setValue('Approved');

    // Build data object for downstream functions
    var apptData = {
      fullName: row[1], email: row[2], phone: row[3],
      service:  row[4], date:  normalizeSheetDate(row[5]),
      time:     normalizeSheetTime(row[6]), notes: row[7]
    };

    var config  = getConfigObject();
    createClinicCalendarEvent(apptData, config);
    var calLink = buildCalendarLink(apptData, config);
    if (apptData.email) sendPatientApprovedEmail(apptData, config, calLink);

    return jsonOut({ success: true, message: 'Booking approved.' });
  } catch (err) {
    Logger.log('handleApproveBooking error: ' + err);
    return jsonOut({ success: false, error: err.toString() });
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}

// ─────────────────────────────────────────────────────────────
// Admin: reject a booking
// ─────────────────────────────────────────────────────────────

function handleRejectBooking(data) {
  var bookingId = data.id;
  if (!bookingId) return jsonOut({ success: false, error: 'Missing booking id.' });

  var lock = LockService.getScriptLock();
  try { lock.waitLock(10000); } catch(e) {
    return jsonOut({ success: false, error: 'System busy. Try again.' });
  }
  try {
    var ss    = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('Appointments');
    if (!sheet || sheet.getLastRow() <= 1) return jsonOut({ success: false, error: 'No bookings found.' });

    var numRows = sheet.getLastRow() - 1;
    var rows    = sheet.getRange(2, 1, numRows, 10).getValues();
    var rowIdx  = -1;

    for (var i = 0; i < rows.length; i++) {
      if (rows[i][9] === bookingId) { rowIdx = i; break; }
    }
    if (rowIdx === -1) return jsonOut({ success: false, error: 'Booking not found.' });

    var row    = rows[rowIdx];
    var status = (row[8] || '').toString().trim();
    if (status === 'Rejected') return jsonOut({ success: false, error: 'Already rejected.' });

    sheet.getRange(rowIdx + 2, 9).setValue('Rejected');

    var apptData = {
      fullName: row[1], email: row[2], phone: row[3],
      service:  row[4], date:  normalizeSheetDate(row[5]),
      time:     normalizeSheetTime(row[6]), notes: row[7]
    };

    var config = getConfigObject();
    var reason = (data.reason || '').trim();
    if (apptData.email) sendPatientRejectedEmail(apptData, config, reason);

    return jsonOut({ success: true, message: 'Booking rejected.' });
  } catch (err) {
    Logger.log('handleRejectBooking error: ' + err);
    return jsonOut({ success: false, error: err.toString() });
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}

// ─────────────────────────────────────────────────────────────
// Admin: update (edit) a booking + notify patient
// Expected POST fields:
//   id, fullName, email, phone, service, date, time, notes,
//   message (optional concierge note included in patient email)
// ─────────────────────────────────────────────────────────────

function handleUpdateBooking(data) {
  var bookingId = data.id;
  if (!bookingId) return jsonOut({ success: false, error: 'Missing booking id.' });

  var lock = LockService.getScriptLock();
  try { lock.waitLock(10000); } catch (e) {
    return jsonOut({ success: false, error: 'System busy. Try again.' });
  }
  try {
    var ss    = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('Appointments');
    if (!sheet || sheet.getLastRow() <= 1) return jsonOut({ success: false, error: 'No bookings found.' });

    var numRows = sheet.getLastRow() - 1;
    var rows    = sheet.getRange(2, 1, numRows, 10).getValues();
    var rowIdx  = -1;
    for (var i = 0; i < rows.length; i++) {
      if (rows[i][9] === bookingId) { rowIdx = i; break; }
    }
    if (rowIdx === -1) return jsonOut({ success: false, error: 'Booking not found.' });

    var oldRow = rows[rowIdx];
    var oldData = {
      fullName: oldRow[1] || '',
      email:    oldRow[2] || '',
      phone:    oldRow[3] || '',
      service:  oldRow[4] || '',
      date:     normalizeSheetDate(oldRow[5]),
      time:     normalizeSheetTime(oldRow[6]),
      notes:    oldRow[7] || '',
      status:   (oldRow[8] || 'Approved').toString().trim()
    };

    // Build the new data — fall back to old values if not supplied
    var newData = {
      fullName: (data.fullName || oldData.fullName).toString().trim(),
      email:    (data.email    || oldData.email).toString().trim(),
      phone:    (data.phone    || oldData.phone).toString().trim(),
      service:  (data.service  || oldData.service).toString().trim(),
      date:     (data.date     || oldData.date).toString().trim(),
      time:     (data.time     || oldData.time).toString().trim(),
      notes:    (data.notes !== undefined ? data.notes : oldData.notes).toString().trim(),
      status:   oldData.status
    };

    // Persist updated cells. Sheet rows are 1-indexed and we have a header row.
    var sheetRow = rowIdx + 2;
    sheet.getRange(sheetRow, 2).setValue(newData.fullName);
    sheet.getRange(sheetRow, 3).setValue(newData.email);
    sheet.getRange(sheetRow, 4).setValue(newData.phone);
    sheet.getRange(sheetRow, 5).setValue(newData.service);
    // Cols 6-7 are text-formatted to avoid Sheets auto-formatting dates/times
    sheet.getRange(sheetRow, 6, 1, 2).setNumberFormat('@');
    sheet.getRange(sheetRow, 6).setValue(newData.date);
    sheet.getRange(sheetRow, 7).setValue(newData.time);
    sheet.getRange(sheetRow, 8).setValue(newData.notes);

    // If this was already Approved, sync the calendar event
    var dateOrTimeOrServiceChanged =
      oldData.date !== newData.date ||
      oldData.time !== newData.time ||
      oldData.service !== newData.service ||
      oldData.fullName !== newData.fullName;

    var config = getConfigObject();
    if (oldData.status === 'Approved' && dateOrTimeOrServiceChanged) {
      try { removeCalendarEventFor(oldData, config); } catch (err) { Logger.log('removeCalendarEventFor failed: ' + err); }
      try { createClinicCalendarEvent(newData, config); } catch (err) { Logger.log('createClinicCalendarEvent failed: ' + err); }
    }

    // Notify patient if we have an email on file
    if (newData.email) {
      var concMsg = (data.message || '').toString().trim();
      sendPatientUpdatedEmail(oldData, newData, config, concMsg);
    }

    return jsonOut({ success: true, message: 'Booking updated.' });
  } catch (err) {
    Logger.log('handleUpdateBooking error: ' + err);
    return jsonOut({ success: false, error: err.toString() });
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}

// ─────────────────────────────────────────────────────────────
// Calendar: remove existing event for a booking (used on update)
// Matches by date + time + patient name on the configured calendar.
// ─────────────────────────────────────────────────────────────

function removeCalendarEventFor(data, config) {
  var calendarId = (config.CALENDAR_ID || '').trim();
  if (!calendarId || calendarId === 'primary') return;
  if (!data.date || !data.time) return;

  var cal = CalendarApp.getCalendarById(calendarId);
  if (!cal) return;

  var dur   = parseInt(config.CALENDAR_DURATION_MINS || '60', 10);
  var dp    = data.date.split('-').map(Number);
  var tp    = data.time.split(':').map(Number);
  var start = new Date(dp[0], dp[1]-1, dp[2], tp[0], tp[1], 0);
  var end   = new Date(start.getTime() + dur * 60000);

  // Search a small window around the slot to be safe
  var windowStart = new Date(start.getTime() - 5 * 60000);
  var windowEnd   = new Date(end.getTime()   + 5 * 60000);
  var events = cal.getEvents(windowStart, windowEnd);

  var nameNeedle = (data.fullName || '').toLowerCase();
  events.forEach(function(ev) {
    var title = (ev.getTitle() || '').toLowerCase();
    if (nameNeedle && title.indexOf(nameNeedle) === -1) return;
    try { ev.deleteEvent(); Logger.log('[calendar] removed event: ' + ev.getTitle()); }
    catch (err) { Logger.log('[calendar] delete failed: ' + err); }
  });
}

// ─────────────────────────────────────────────────────────────
// Patient email: "Your booking has been updated"
// ─────────────────────────────────────────────────────────────

function sendPatientUpdatedEmail(oldData, newData, config, concMsg) {
  var clinicName = config.CLINIC_NAME    || 'TomFord Dental';
  var phone      = config.CLINIC_PHONE   || '';
  var email      = config.CLINIC_EMAIL   || '';
  var hours      = config.CLINIC_HOURS   || '';
  var addr       = config.CLINIC_ADDRESS || '';
  var tag        = config.BOOKING_TAGLINE|| 'where every tooth matters.';

  var oldDd = formatDisplayDate(oldData.date);
  var oldDt = formatDisplayTime(oldData.time);
  var newDd = formatDisplayDate(newData.date);
  var newDt = formatDisplayTime(newData.time);

  var changeLines = [];
  if (oldData.service  !== newData.service)  changeLines.push({ label: 'Service', oldVal: oldData.service,  newVal: newData.service  });
  if (oldData.date     !== newData.date)     changeLines.push({ label: 'Date',    oldVal: oldDd,            newVal: newDd            });
  if (oldData.time     !== newData.time)     changeLines.push({ label: 'Time',    oldVal: oldDt,            newVal: newDt            });
  if (oldData.fullName !== newData.fullName) changeLines.push({ label: 'Name',    oldVal: oldData.fullName, newVal: newData.fullName });

  var changesHtml = '';
  changeLines.forEach(function(c) {
    changesHtml +=
      '<tr><td style="color:#9ca3af;font-size:11px;text-transform:uppercase;letter-spacing:.1em;width:90px;vertical-align:top;padding:6px 0;">'+c.label+'</td>'
      +'<td style="padding:6px 0;color:#6b7280;font-size:13px;text-decoration:line-through;">'+c.oldVal+'</td></tr>'
      +'<tr><td></td><td style="padding:0 0 8px;color:#111827;font-size:14px;font-weight:700;">'+c.newVal+'</td></tr>';
  });
  if (!changesHtml) {
    changesHtml = '<tr><td colspan="2" style="padding:8px 0;color:#4b5563;font-size:13px;">Your contact details or notes were updated. Time and date are unchanged.</td></tr>';
  }

  var concBlock = concMsg
    ? '<div style="background:#fff8f3;border-left:4px solid #f46709;border-radius:8px;padding:14px 18px;margin:18px 0;">'
      +'<p style="margin:0 0 4px;color:#7c2d04;font-size:11px;font-weight:700;letter-spacing:.1em;text-transform:uppercase;">Note from the clinic</p>'
      +'<p style="margin:0;color:#4b5563;font-size:13px;line-height:1.7;">'+concMsg+'</p></div>'
    : '';

  var subj = 'Booking Updated - ' + clinicName + ' - ' + newDd;

  var html =
    '<!DOCTYPE html><html><body style="margin:0;padding:0;background:#f5f5f5;font-family:Arial,sans-serif;">'
    +'<table width="100%" cellpadding="0" cellspacing="0" style="background:#f5f5f5;padding:40px 16px;"><tr><td align="center">'
    +'<table width="560" cellpadding="0" cellspacing="0" style="max-width:560px;width:100%;background:#ffffff;border-radius:16px;overflow:hidden;box-shadow:0 4px 24px rgba(0,0,0,.07);">'
    +'<tr><td style="background:linear-gradient(135deg,#f46709,#e05800);padding:32px 36px;text-align:center;">'
    +'<p style="margin:0;color:rgba(255,255,255,.7);font-size:11px;letter-spacing:.3em;text-transform:uppercase;">Booking Updated</p>'
    +'<h1 style="margin:8px 0 4px;color:#fff;font-size:24px;font-family:Georgia,serif;">'+clinicName+'</h1>'
    +'<p style="margin:0;color:rgba(255,255,255,.8);font-size:12px;font-style:italic;font-family:Georgia,serif;">'+tag+'</p>'
    +'</td></tr>'
    +'<tr><td style="padding:32px 36px 24px;">'
    +'<p style="margin:0 0 6px;color:#111827;font-size:16px;font-weight:700;">Hi '+(newData.fullName||'there')+',</p>'
    +'<p style="margin:0 0 22px;color:#4b5563;font-size:14px;line-height:1.7;">Our team has updated the details of your booking. Here is what changed:</p>'
    +'<table width="100%" cellpadding="0" cellspacing="0" style="background:#fff8f3;border-radius:12px;border-left:4px solid #f46709;margin-bottom:18px;">'
    +'<tr><td style="padding:18px 22px;">'
    +'<p style="margin:0 0 14px;color:#f46709;font-size:10px;letter-spacing:.25em;text-transform:uppercase;font-weight:700;">Updated Details</p>'
    +'<table width="100%" cellpadding="0" cellspacing="0">'+changesHtml+'</table>'
    +'</td></tr></table>'
    + concBlock
    +'<div style="background:#fef3f2;border-radius:10px;padding:14px 18px;margin-bottom:22px;">'
    +'<p style="margin:0;color:#991b1b;font-size:13px;"><strong>Are you okay with this change?</strong></p>'
    +'<p style="margin:6px 0 0;color:#6b7280;font-size:12px;line-height:1.6;">Please reply to this email or call us at <strong>'+phone+'</strong> to confirm. If we do not hear from you, we will assume the new slot works.</p>'
    +'</div>'
    +'<p style="margin:0 0 6px;color:#111827;font-size:14px;font-weight:700;">New appointment summary</p>'
    +'<table width="100%" cellpadding="0" cellspacing="0" style="background:#f9fafb;border-radius:10px;margin-bottom:18px;">'
    +'<tr><td style="padding:14px 18px;">'
    +'<table width="100%" cellpadding="4" cellspacing="0">'
    +'<tr><td style="color:#9ca3af;font-size:11px;text-transform:uppercase;letter-spacing:.1em;width:80px;">Service</td><td style="color:#111827;font-size:13px;font-weight:600;">'+(newData.service||'-')+'</td></tr>'
    +'<tr><td style="color:#9ca3af;font-size:11px;text-transform:uppercase;letter-spacing:.1em;">Date</td><td style="color:#111827;font-size:13px;font-weight:600;">'+newDd+'</td></tr>'
    +'<tr><td style="color:#9ca3af;font-size:11px;text-transform:uppercase;letter-spacing:.1em;">Time</td><td style="color:#111827;font-size:13px;font-weight:600;">'+newDt+'</td></tr>'
    +'</table>'
    +'</td></tr></table>'
    +'<p style="margin:0;color:#6b7280;font-size:13px;">Questions? Call <strong style="color:#111827;">'+phone+'</strong>'+(email?' or email <a href="mailto:'+email+'" style="color:#f46709;text-decoration:none;">'+email+'</a>':'')+'</p>'
    +'</td></tr>'
    +'<tr><td style="background:#fafafa;border-top:1px solid #f3f4f6;padding:18px 36px;text-align:center;">'
    +'<p style="margin:0 0 3px;color:#111827;font-size:13px;font-weight:700;">'+clinicName+'</p>'
    +'<p style="margin:0 0 2px;color:#9ca3af;font-size:11px;">'+addr+'</p>'
    +'<p style="margin:0;color:#9ca3af;font-size:11px;">'+hours+'</p>'
    +'</td></tr>'
    +'</table>'
    +'</td></tr></table></body></html>';

  GmailApp.sendEmail(newData.email, subj,
    'Hi '+(newData.fullName||'there')+', we updated your booking. Please reply to this email or call '+phone+' to confirm the new details.',
    { htmlBody: html, name: clinicName, replyTo: email || undefined }
  );
}

// ─────────────────────────────────────────────────────────────
// Admin: get bookings data
// ─────────────────────────────────────────────────────────────

function getAllBookings() {
  return getBookingsByStatus(null); // null = all statuses
}

function getBookingsByStatus(statuses) {
  try {
    var ss    = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('Appointments');
    if (!sheet || sheet.getLastRow() <= 1) return { bookings: [] };

    var numRows = sheet.getLastRow() - 1;
    var rows    = sheet.getRange(2, 1, numRows, 10).getValues();
    var result  = [];

    rows.forEach(function(row) {
      var status = (row[8] || 'Approved').toString().trim(); // legacy = Approved
      if (statuses && statuses.indexOf(status) === -1) return;
      result.push({
        timestamp: row[0] ? row[0].toString() : '',
        fullName:  row[1] || '',
        email:     row[2] || '',
        phone:     row[3] || '',
        service:   row[4] || '',
        date:      normalizeSheetDate(row[5]),
        time:      normalizeSheetTime(row[6]),
        notes:     row[7] || '',
        status:    status,
        id:        row[9] || ''
      });
    });

    // Newest first
    result.sort(function(a, b) {
      return new Date(b.timestamp) - new Date(a.timestamp);
    });

    return { bookings: result };
  } catch (err) {
    return { bookings: [], error: err.toString() };
  }
}

// ─────────────────────────────────────────────────────────────
// Slot availability (counts Pending + Approved, not Rejected)
// ─────────────────────────────────────────────────────────────

function validateSlotAvailability(data, config) {
  var maxPerSlot = parseInt(config.MAX_BOOKINGS_PER_SLOT || '1', 10);
  var booked = getBookedSlotCounts(data.date);
  if ((booked[data.time] || 0) >= maxPerSlot) {
    return { ok: false, error: 'That time slot is no longer available. Please pick another.' };
  }

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
    Logger.log('validateSlotAvailability calendar check failed: ' + err);
  }

  try {
    var openDays = parseOpenDays(config.OPEN_DAYS || '1,2,3,4,5,6');
    var dp2 = data.date.split('-').map(Number);
    var dow = new Date(dp2[0], dp2[1]-1, dp2[2]).getDay();
    if (openDays.indexOf(dow) === -1) {
      return { ok: false, error: 'That day is not a clinic working day.' };
    }
    var tp2 = data.time.split(':').map(Number);
    var tMins = tp2[0]*60 + tp2[1];
    var oArr  = (config.OPEN_TIME  || '09:00').split(':').map(Number);
    var cArr  = (config.CLOSE_TIME || '19:00').split(':').map(Number);
    if (tMins < oArr[0]*60+oArr[1] || tMins >= cArr[0]*60+cArr[1]) {
      return { ok: false, error: 'That time is outside clinic hours.' };
    }
  } catch (err) {
    Logger.log('validateSlotAvailability range check failed: ' + err);
  }

  return { ok: true };
}

function getBookedSlotCounts(dateStr) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Appointments');
  if (!sheet || sheet.getLastRow() <= 1) return {};
  var numCols = Math.min(10, sheet.getLastColumn());
  var data = sheet.getRange(2, 1, sheet.getLastRow()-1, numCols).getValues();
  var counts = {};
  data.forEach(function(row) {
    // Col I (index 8) = Status. Old rows without status default to Approved.
    var status = numCols >= 9 ? (row[8] || 'Approved').toString().trim() : 'Approved';
    if (status === 'Rejected') return; // rejected slots are free again
    var d = normalizeSheetDate(row[5]);
    var t = normalizeSheetTime(row[6]);
    if (d === dateStr && t) counts[t] = (counts[t] || 0) + 1;
  });
  return counts;
}

// ─────────────────────────────────────────────────────────────
// getSlots (unchanged logic, uses updated getBookedSlotCounts)
// ─────────────────────────────────────────────────────────────

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

    var nowPHT       = getNowPHT();
    var bufferMs     = bufferMins * 60 * 1000;
    var advanceCutoff = new Date(nowPHT.getTime() + advanceDays * 24 * 60 * 60 * 1000);

    var fromParts = fromStr.split('-').map(Number);
    var result    = {};
    var collected = 0;

    for (var offset = 0; offset < numDays + 7; offset++) {
      if (collected >= numDays) break;

      var d   = new Date(fromParts[0], fromParts[1]-1, fromParts[2] + offset);
      var dow = d.getDay();
      if (openDays.indexOf(dow) === -1) continue;

      var dateStr = formatDateStr(d);
      if (d > advanceCutoff) { collected++; result[dateStr] = []; continue; }

      var allSlots = generateSlots(openTime, closeTime, slotDur);
      var cutoff   = new Date(nowPHT.getTime() + bufferMs);
      allSlots = allSlots.filter(function(t) {
        var parts  = t.split(':').map(Number);
        var slotDt = new Date(d.getFullYear(), d.getMonth(), d.getDate(), parts[0], parts[1], 0);
        return slotDt > cutoff;
      });

      var booked = getBookedSlotCounts(dateStr);
      var busy   = getCalendarBusy(calendarId, d, slotDur);

      var available = allSlots.filter(function(t) {
        if ((booked[t] || 0) >= maxPerSlot) return false;
        var parts    = t.split(':').map(Number);
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

// ─────────────────────────────────────────────────────────────
// saveAppointment (v3: adds Status + Booking ID columns)
// ─────────────────────────────────────────────────────────────

function saveAppointment(data, bookingId, status) {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Appointments');
  if (!sheet) {
    sheet = ss.insertSheet('Appointments');
    sheet.appendRow(['Timestamp','Full Name','Email','Phone','Service','Date','Time','Notes','Status','Booking ID']);
    var h = sheet.getRange(1, 1, 1, 10);
    h.setFontWeight('bold'); h.setBackground('#f46709'); h.setFontColor('#ffffff');
  }

  // If sheet exists but has old 8-column headers, add new headers
  if (sheet.getLastColumn() < 10) {
    sheet.getRange(1, 9).setValue('Status').setFontWeight('bold').setBackground('#f46709').setFontColor('#ffffff');
    sheet.getRange(1, 10).setValue('Booking ID').setFontWeight('bold').setBackground('#f46709').setFontColor('#ffffff');
  }

  sheet.appendRow([
    new Date().toLocaleString('en-PH', { timeZone: 'Asia/Manila' }),
    data.fullName || '', data.email || '', data.phone || '',
    data.service  || '', data.date  || '', data.time  || '', data.notes || '',
    status    || 'Pending',
    bookingId || ''
  ]);

  var lastRow = sheet.getLastRow();
  sheet.getRange(lastRow, 6, 1, 2).setNumberFormat('@');
  sheet.getRange(lastRow, 6).setValue(data.date || '');
  sheet.getRange(lastRow, 7).setValue(data.time || '');
}

// ─────────────────────────────────────────────────────────────
// Emails
// ─────────────────────────────────────────────────────────────

// 1. To patient: "We received your request, pending confirmation"
function sendPatientPendingEmail(data, config) {
  var clinicName = config.CLINIC_NAME    || 'TomFord Dental';
  var phone      = config.CLINIC_PHONE   || '';
  var email      = config.CLINIC_EMAIL   || '';
  var hours      = config.CLINIC_HOURS   || '';
  var addr       = config.CLINIC_ADDRESS || '';
  var tag        = config.BOOKING_TAGLINE|| 'where every tooth matters.';
  var dd         = formatDisplayDate(data.date);
  var dt         = formatDisplayTime(data.time);
  var subj       = 'Request Received — ' + clinicName + ' · ' + dd;

  var html =
    '<!DOCTYPE html><html><body style="margin:0;padding:0;background:#f5f5f5;font-family:Arial,sans-serif;">'
    +'<table width="100%" cellpadding="0" cellspacing="0" style="background:#f5f5f5;padding:40px 16px;"><tr><td align="center">'
    +'<table width="560" cellpadding="0" cellspacing="0" style="max-width:560px;width:100%;background:#ffffff;border-radius:16px;overflow:hidden;box-shadow:0 4px 24px rgba(0,0,0,.07);">'
    // Header
    +'<tr><td style="background:linear-gradient(135deg,#f46709,#e05800);padding:32px 36px;text-align:center;">'
    +'<p style="margin:0;color:rgba(255,255,255,.7);font-size:11px;letter-spacing:.3em;text-transform:uppercase;">Booking Request</p>'
    +'<h1 style="margin:8px 0 4px;color:#fff;font-size:24px;font-family:Georgia,serif;">'+clinicName+'</h1>'
    +'<p style="margin:0;color:rgba(255,255,255,.8);font-size:12px;font-style:italic;font-family:Georgia,serif;">'+tag+'</p>'
    +'</td></tr>'
    // Body
    +'<tr><td style="padding:32px 36px 24px;">'
    +'<p style="margin:0 0 6px;color:#111827;font-size:16px;font-weight:700;">Hi '+( data.fullName || 'there')+'!</p>'
    +'<p style="margin:0 0 24px;color:#4b5563;font-size:14px;line-height:1.7;">We\'ve received your appointment request. Our team will review it and confirm your slot shortly — usually within a few hours during clinic hours.</p>'
    // Details card
    +'<table width="100%" cellpadding="0" cellspacing="0" style="background:#fff8f3;border-radius:12px;border-left:4px solid #f46709;margin-bottom:24px;">'
    +'<tr><td style="padding:20px 24px;">'
    +'<p style="margin:0 0 14px;color:#f46709;font-size:10px;letter-spacing:.25em;text-transform:uppercase;font-weight:700;">Your Request</p>'
    +'<table width="100%" cellpadding="6" cellspacing="0">'
    +'<tr><td style="color:#9ca3af;font-size:11px;text-transform:uppercase;letter-spacing:.1em;width:80px;vertical-align:top;">Service</td><td style="color:#111827;font-size:14px;font-weight:700;">'+( data.service||'—')+'</td></tr>'
    +'<tr><td style="color:#9ca3af;font-size:11px;text-transform:uppercase;letter-spacing:.1em;vertical-align:top;">Date</td><td style="color:#111827;font-size:14px;font-weight:700;">'+dd+'</td></tr>'
    +'<tr><td style="color:#9ca3af;font-size:11px;text-transform:uppercase;letter-spacing:.1em;vertical-align:top;">Time</td><td style="color:#111827;font-size:14px;font-weight:700;">'+dt+'</td></tr>'
    +'</table>'
    +'</td></tr></table>'
    // Status pill
    +'<table width="100%" cellpadding="0" cellspacing="0" style="margin-bottom:24px;">'
    +'<tr><td style="background:#fff3cd;border-radius:10px;padding:14px 18px;">'
    +'<p style="margin:0;color:#92400e;font-size:13px;"><strong>Status: Pending Review</strong></p>'
    +'<p style="margin:4px 0 0;color:#78350f;font-size:12px;line-height:1.6;">This is not yet a confirmed appointment. We\'ll send you another email once confirmed.</p>'
    +'</td></tr></table>'
    +'<p style="margin:0;color:#6b7280;font-size:13px;">Questions? Call us at <strong style="color:#111827;">'+phone+'</strong>'+(email?' or email <a href="mailto:'+email+'" style="color:#f46709;text-decoration:none;">'+email+'</a>':'')+'</p>'
    +'</td></tr>'
    // Footer
    +'<tr><td style="background:#fafafa;border-top:1px solid #f3f4f6;padding:18px 36px;text-align:center;">'
    +'<p style="margin:0 0 3px;color:#111827;font-size:13px;font-weight:700;">'+clinicName+'</p>'
    +'<p style="margin:0 0 2px;color:#9ca3af;font-size:11px;">'+addr+'</p>'
    +'<p style="margin:0;color:#9ca3af;font-size:11px;">'+hours+'</p>'
    +'</td></tr>'
    +'</table>'
    +'<p style="margin:12px 0 0;color:#d1d5db;font-size:11px;text-align:center;">You received this because you submitted a booking request at '+clinicName+'.</p>'
    +'</td></tr></table></body></html>';

  GmailApp.sendEmail(data.email, subj,
    'Hi '+(data.fullName||'there')+', we received your request for '+(data.service||'an appointment')+' on '+dd+' at '+dt+'. Our team will review and confirm shortly.',
    { htmlBody: html, name: clinicName, replyTo: email || undefined }
  );
}

// 2. To clinic: "New booking request — review in admin"
function sendClinicRequestEmail(data, config, bookingId) {
  var to         = config.CONCIERGE_EMAIL || config.CLINIC_EMAIL;
  if (!to) { Logger.log('No CONCIERGE_EMAIL — skipping clinic notification.'); return; }

  var clinicName = config.CLINIC_NAME    || 'TomFord Dental';
  var adminUrl   = (config.ADMIN_URL     || 'https://tomforddental.com/admin').replace(/\/$/, '');
  var dd         = formatDisplayDate(data.date);
  var dt         = formatDisplayTime(data.time);
  var subj       = 'New Request: ' + (data.fullName||'?') + ' - ' + (data.service||'-') + ' - ' + dd;
  var reviewLink = adminUrl + '?highlight=' + encodeURIComponent(bookingId);

  var html =
    '<!DOCTYPE html><html><body style="margin:0;padding:0;background:#f5f5f5;font-family:Arial,sans-serif;">'
    +'<table width="100%" cellpadding="0" cellspacing="0" style="background:#f5f5f5;padding:40px 16px;"><tr><td align="center">'
    +'<table width="520" cellpadding="0" cellspacing="0" style="max-width:520px;width:100%;background:#ffffff;border-radius:16px;overflow:hidden;box-shadow:0 4px 24px rgba(0,0,0,.07);">'
    // Header
    +'<tr><td style="background:#111827;padding:24px 30px;">'
    +'<p style="margin:0;color:rgba(255,255,255,.45);font-size:10px;letter-spacing:.3em;text-transform:uppercase;">New Booking Request</p>'
    +'<h2 style="margin:6px 0 0;color:#ffffff;font-size:20px;font-family:Georgia,serif;">'+clinicName+' Admin</h2>'
    +'</td></tr>'
    // Body
    +'<tr><td style="padding:28px 30px 24px;">'
    +'<p style="margin:0 0 20px;color:#4b5563;font-size:13px;line-height:1.6;">A patient has requested an appointment. Please review and approve or reject it from the admin panel.</p>'
    // Patient details
    +'<table width="100%" cellpadding="0" cellspacing="0" style="border:1px solid #f3f4f6;border-radius:10px;margin-bottom:24px;overflow:hidden;">'
    +'<tr style="background:#f9fafb;"><td style="padding:10px 16px;color:#9ca3af;font-size:11px;text-transform:uppercase;letter-spacing:.1em;width:90px;">Patient</td>'
    +'<td style="padding:10px 16px;color:#111827;font-size:14px;font-weight:700;">'+(data.fullName||'—')+'</td></tr>'
    +'<tr><td style="padding:10px 16px;color:#9ca3af;font-size:11px;text-transform:uppercase;letter-spacing:.1em;">Service</td>'
    +'<td style="padding:10px 16px;"><span style="background:#fff3e8;color:#f46709;font-size:12px;font-weight:700;padding:3px 12px;border-radius:20px;">'+(data.service||'—')+'</span></td></tr>'
    +'<tr style="background:#f9fafb;"><td style="padding:10px 16px;color:#9ca3af;font-size:11px;text-transform:uppercase;letter-spacing:.1em;">Date</td>'
    +'<td style="padding:10px 16px;color:#111827;font-size:14px;font-weight:700;">'+dd+'</td></tr>'
    +'<tr><td style="padding:10px 16px;color:#9ca3af;font-size:11px;text-transform:uppercase;letter-spacing:.1em;">Time</td>'
    +'<td style="padding:10px 16px;color:#111827;font-size:14px;font-weight:700;">'+dt+'</td></tr>'
    +'<tr style="background:#f9fafb;"><td style="padding:10px 16px;color:#9ca3af;font-size:11px;text-transform:uppercase;letter-spacing:.1em;">Phone</td>'
    +'<td style="padding:10px 16px;color:#111827;font-size:13px;">'+(data.phone||'—')+'</td></tr>'
    +'<tr><td style="padding:10px 16px;color:#9ca3af;font-size:11px;text-transform:uppercase;letter-spacing:.1em;">Email</td>'
    +'<td style="padding:10px 16px;"><a href="mailto:'+(data.email||'')+'" style="color:#f46709;text-decoration:none;font-size:13px;">'+(data.email||'—')+'</a></td></tr>'
    +(data.notes ? '<tr style="background:#f9fafb;"><td style="padding:10px 16px;color:#9ca3af;font-size:11px;text-transform:uppercase;letter-spacing:.1em;vertical-align:top;">Notes</td><td style="padding:10px 16px;color:#4b5563;font-size:13px;">'+data.notes+'</td></tr>' : '')
    +'</table>'
    // CTA
    +'<table width="100%" cellpadding="0" cellspacing="0"><tr><td align="center" style="padding-bottom:8px;">'
    +'<a href="'+reviewLink+'" style="display:inline-block;background:#f46709;color:#ffffff;text-decoration:none;padding:14px 36px;border-radius:50px;font-size:14px;font-weight:700;font-family:Arial,sans-serif;">Review in Admin Panel →</a>'
    +'</td></tr></table>'
    +'<p style="margin:16px 0 0;color:#d1d5db;font-size:11px;text-align:center;">Booking ID: '+bookingId+'</p>'
    +'</td></tr>'
    // Footer
    +'<tr><td style="background:#f9fafb;border-top:1px solid #f3f4f6;padding:14px 30px;text-align:center;">'
    +'<p style="margin:0;color:#9ca3af;font-size:11px;">Sent '+new Date().toLocaleString('en-PH',{timeZone:'Asia/Manila'})+' PHT · '+clinicName+' Booking System</p>'
    +'</td></tr>'
    +'</table>'
    +'</td></tr></table></body></html>';

  GmailApp.sendEmail(to, subj,
    'New booking request from '+(data.fullName||'?')+' for '+(data.service||'—')+' on '+dd+' at '+dt+'. Review at: '+reviewLink,
    { htmlBody: html, name: clinicName + ' Booking System' }
  );
}

// 3. To patient: "Your appointment is confirmed!"
function sendPatientApprovedEmail(data, config, calLink) {
  var clinicName = config.CLINIC_NAME    || 'TomFord Dental';
  var addr       = config.CLINIC_ADDRESS || '';
  var phone      = config.CLINIC_PHONE   || '';
  var hours      = config.CLINIC_HOURS   || '';
  var email      = config.CLINIC_EMAIL   || '';
  var tag        = config.BOOKING_TAGLINE|| 'where every tooth matters.';
  var dd         = formatDisplayDate(data.date);
  var dt         = formatDisplayTime(data.time);
  var subj       = 'Confirmed: Your ' + clinicName + ' Appointment - ' + dd;

  var calBtn = calLink
    ? '<tr><td align="center" style="padding-bottom:24px;"><a href="'+calLink+'" style="display:inline-block;background:#f46709;color:#fff;text-decoration:none;padding:13px 30px;border-radius:50px;font-size:14px;font-weight:600;font-family:Arial,sans-serif;">Add to Google Calendar</a></td></tr>'
    : '';

  var html =
    '<!DOCTYPE html><html><body style="margin:0;padding:0;background:#f5f5f5;font-family:Arial,sans-serif;">'
    +'<table width="100%" cellpadding="0" cellspacing="0" style="background:#f5f5f5;padding:40px 16px;"><tr><td align="center">'
    +'<table width="560" cellpadding="0" cellspacing="0" style="max-width:560px;width:100%;background:#ffffff;border-radius:16px;overflow:hidden;box-shadow:0 4px 24px rgba(0,0,0,.07);">'
    +'<tr><td style="background:linear-gradient(135deg,#f46709,#e05800);padding:32px 36px;text-align:center;">'
    +'<p style="margin:0;color:rgba(255,255,255,.7);font-size:11px;letter-spacing:.3em;text-transform:uppercase;">Appointment Confirmed</p>'
    +'<h1 style="margin:8px 0 4px;color:#fff;font-size:24px;font-family:Georgia,serif;">'+clinicName+'</h1>'
    +'<p style="margin:0;color:rgba(255,255,255,.8);font-size:12px;font-style:italic;font-family:Georgia,serif;">'+tag+'</p>'
    +'</td></tr>'
    +'<tr><td style="padding:32px 36px 24px;">'
    +'<p style="margin:0 0 6px;color:#111827;font-size:16px;font-weight:700;">You\'re all set, '+(data.fullName||'there')+'!</p>'
    +'<p style="margin:0 0 24px;color:#4b5563;font-size:14px;line-height:1.7;">Your appointment has been confirmed. We look forward to seeing you!</p>'
    +'<table width="100%" cellpadding="0" cellspacing="0" style="background:#fff8f3;border-radius:12px;border-left:4px solid #f46709;margin-bottom:24px;">'
    +'<tr><td style="padding:20px 24px;">'
    +'<p style="margin:0 0 14px;color:#f46709;font-size:10px;letter-spacing:.25em;text-transform:uppercase;font-weight:700;">Appointment Details</p>'
    +'<table width="100%" cellpadding="6" cellspacing="0">'
    +'<tr><td style="color:#9ca3af;font-size:11px;text-transform:uppercase;letter-spacing:.1em;width:80px;">Service</td><td style="color:#111827;font-size:14px;font-weight:700;">'+(data.service||'—')+'</td></tr>'
    +'<tr><td style="color:#9ca3af;font-size:11px;text-transform:uppercase;letter-spacing:.1em;">Date</td><td style="color:#111827;font-size:14px;font-weight:700;">'+dd+'</td></tr>'
    +'<tr><td style="color:#9ca3af;font-size:11px;text-transform:uppercase;letter-spacing:.1em;">Time</td><td style="color:#111827;font-size:14px;font-weight:700;">'+dt+'</td></tr>'
    +'<tr><td style="color:#9ca3af;font-size:11px;text-transform:uppercase;letter-spacing:.1em;">Address</td><td style="color:#4b5563;font-size:13px;">'+addr+'</td></tr>'
    +'</table>'
    +'</td></tr></table>'
    +'<table width="100%" cellpadding="0" cellspacing="0">'+calBtn+'</table>'
    +'<p style="margin:0;color:#6b7280;font-size:13px;">Need to reschedule? Call <strong style="color:#111827;">'+phone+'</strong>'+(email?' or email <a href="mailto:'+email+'" style="color:#f46709;text-decoration:none;">'+email+'</a>':'')+'</p>'
    +'</td></tr>'
    +'<tr><td style="background:#fafafa;border-top:1px solid #f3f4f6;padding:18px 36px;text-align:center;">'
    +'<p style="margin:0 0 3px;color:#111827;font-size:13px;font-weight:700;">'+clinicName+'</p>'
    +'<p style="margin:0 0 2px;color:#9ca3af;font-size:11px;">'+addr+'</p>'
    +'<p style="margin:0;color:#9ca3af;font-size:11px;">'+hours+'</p>'
    +'</td></tr>'
    +'</table>'
    +'</td></tr></table></body></html>';

  GmailApp.sendEmail(data.email, subj,
    'Great news, '+(data.fullName||'there')+'! Your '+( data.service||'appointment')+' on '+dd+' at '+dt+' is confirmed. We look forward to seeing you!'+(calLink?' Add to calendar: '+calLink:''),
    { htmlBody: html, name: clinicName, replyTo: email || undefined }
  );
}

// 4. To patient: "We couldn't accommodate your request"
function sendPatientRejectedEmail(data, config, reason) {
  var clinicName = config.CLINIC_NAME    || 'TomFord Dental';
  var phone      = config.CLINIC_PHONE   || '';
  var email      = config.CLINIC_EMAIL   || '';
  var hours      = config.CLINIC_HOURS   || '';
  var addr       = config.CLINIC_ADDRESS || '';
  var tag        = config.BOOKING_TAGLINE|| 'where every tooth matters.';
  var dd         = formatDisplayDate(data.date);
  var dt         = formatDisplayTime(data.time);
  var subj       = 'Update on Your ' + clinicName + ' Booking Request';

  var reasonBlock = reason
    ? '<p style="margin:0 0 4px;color:#374151;font-size:13px;font-weight:700;">Note from the clinic:</p><p style="margin:0;color:#4b5563;font-size:13px;line-height:1.6;">'+reason+'</p>'
    : '<p style="margin:0;color:#4b5563;font-size:13px;line-height:1.6;">Unfortunately we\'re unable to accommodate your requested slot. This may be due to a scheduling conflict.</p>';

  var html =
    '<!DOCTYPE html><html><body style="margin:0;padding:0;background:#f5f5f5;font-family:Arial,sans-serif;">'
    +'<table width="100%" cellpadding="0" cellspacing="0" style="background:#f5f5f5;padding:40px 16px;"><tr><td align="center">'
    +'<table width="560" cellpadding="0" cellspacing="0" style="max-width:560px;width:100%;background:#ffffff;border-radius:16px;overflow:hidden;box-shadow:0 4px 24px rgba(0,0,0,.07);">'
    +'<tr><td style="background:linear-gradient(135deg,#f46709,#e05800);padding:32px 36px;text-align:center;">'
    +'<p style="margin:0;color:rgba(255,255,255,.7);font-size:11px;letter-spacing:.3em;text-transform:uppercase;">Booking Update</p>'
    +'<h1 style="margin:8px 0 4px;color:#fff;font-size:24px;font-family:Georgia,serif;">'+clinicName+'</h1>'
    +'<p style="margin:0;color:rgba(255,255,255,.8);font-size:12px;font-style:italic;font-family:Georgia,serif;">'+tag+'</p>'
    +'</td></tr>'
    +'<tr><td style="padding:32px 36px 24px;">'
    +'<p style="margin:0 0 6px;color:#111827;font-size:16px;font-weight:700;">Hi '+(data.fullName||'there')+',</p>'
    +'<p style="margin:0 0 24px;color:#4b5563;font-size:14px;line-height:1.7;">Thank you for reaching out to us. Unfortunately, we\'re unable to confirm your request for the following slot:</p>'
    +'<table width="100%" cellpadding="0" cellspacing="0" style="background:#f9fafb;border-radius:12px;border-left:4px solid #d1d5db;margin-bottom:24px;">'
    +'<tr><td style="padding:18px 22px;">'
    +'<table width="100%" cellpadding="6" cellspacing="0">'
    +'<tr><td style="color:#9ca3af;font-size:11px;text-transform:uppercase;letter-spacing:.1em;width:80px;">Service</td><td style="color:#6b7280;font-size:14px;text-decoration:line-through;">'+(data.service||'—')+'</td></tr>'
    +'<tr><td style="color:#9ca3af;font-size:11px;text-transform:uppercase;letter-spacing:.1em;">Date</td><td style="color:#6b7280;font-size:14px;text-decoration:line-through;">'+dd+'</td></tr>'
    +'<tr><td style="color:#9ca3af;font-size:11px;text-transform:uppercase;letter-spacing:.1em;">Time</td><td style="color:#6b7280;font-size:14px;text-decoration:line-through;">'+dt+'</td></tr>'
    +'</table>'
    +'</td></tr></table>'
    +'<div style="background:#fef3f2;border-radius:10px;padding:16px 20px;margin-bottom:24px;">'+reasonBlock+'</div>'
    +'<p style="margin:0 0 8px;color:#111827;font-size:14px;font-weight:700;">Want to book a different slot?</p>'
    +'<p style="margin:0;color:#6b7280;font-size:13px;">Call us at <strong style="color:#111827;">'+phone+'</strong>'+(email?' or email <a href="mailto:'+email+'" style="color:#f46709;text-decoration:none;">'+email+'</a>':'')+'<br>'+hours+'</p>'
    +'</td></tr>'
    +'<tr><td style="background:#fafafa;border-top:1px solid #f3f4f6;padding:18px 36px;text-align:center;">'
    +'<p style="margin:0 0 3px;color:#111827;font-size:13px;font-weight:700;">'+clinicName+'</p>'
    +'<p style="margin:0 0 2px;color:#9ca3af;font-size:11px;">'+addr+'</p>'
    +'<p style="margin:0;color:#9ca3af;font-size:11px;">'+hours+'</p>'
    +'</td></tr>'
    +'</table>'
    +'</td></tr></table></body></html>';

  GmailApp.sendEmail(data.email, subj,
    'Hi '+(data.fullName||'there')+', unfortunately we\'re unable to confirm your request for '+(data.service||'an appointment')+' on '+dd+' at '+dt+'. Please contact us at '+phone+' to rebook.',
    { htmlBody: html, name: clinicName, replyTo: email || undefined }
  );
}

// ─────────────────────────────────────────────────────────────
// Calendar helpers (unchanged)
// ─────────────────────────────────────────────────────────────

function createClinicCalendarEvent(data, config) {
  if (!data.date || !data.time) return;

  var calendarId = (config.CALENDAR_ID || '').trim();
  if (!calendarId || calendarId === 'primary') {
    Logger.log('[calendar] CALENDAR_ID not set or is "primary" — skipping.');
    return;
  }

  try {
    var cal = CalendarApp.getCalendarById(calendarId);
    if (!cal) { Logger.log('[calendar] Calendar not found: ' + calendarId); return; }

    var dur   = parseInt(config.CALENDAR_DURATION_MINS || '60', 10);
    var dp    = data.date.split('-').map(Number);
    var tp    = data.time.split(':').map(Number);
    var start = new Date(dp[0], dp[1]-1, dp[2], tp[0], tp[1], 0);
    var end   = new Date(start.getTime() + dur * 60000);

    var title = (data.service || 'Appointment') + ' — ' + (data.fullName || 'Patient');
    var desc  = 'Patient: ' + (data.fullName || '—')
              + '\nEmail: '  + (data.email    || '—')
              + '\nPhone: '  + (data.phone    || '—')
              + '\nService: '+ (data.service  || '—')
              + '\n\nStatus: Confirmed'
              + '\nBooked via website · Confirmed ' + new Date().toLocaleString('en-PH', { timeZone: 'Asia/Manila' }) + ' PHT';

    cal.createEvent(title, start, end, {
      description: desc,
      location: config.CLINIC_ADDRESS || ''
    });
    Logger.log('[calendar] ✓ Event created: ' + title);
  } catch (err) {
    Logger.log('[calendar] createClinicCalendarEvent error: ' + err);
  }
}

function buildCalendarLink(data, config) {
  try {
    var name  = config.CLINIC_NAME    || 'TomFord Dental';
    var addr  = config.CLINIC_ADDRESS || '';
    var phone = config.CLINIC_PHONE   || '';
    var dur   = parseInt(config.CALENDAR_DURATION_MINS || '60', 10);
    if (!data.date || !data.time) return '';
    var dp = data.date.split('-').map(Number);
    var tp = data.time.split(':').map(Number);
    var s  = new Date(dp[0], dp[1]-1, dp[2], tp[0], tp[1], 0);
    var en = new Date(s.getTime() + dur * 60000);
    function fmt(d) { return d.getFullYear()+pad2(d.getMonth()+1)+pad2(d.getDate())+'T'+pad2(d.getHours())+pad2(d.getMinutes())+'00'; }
    return 'https://calendar.google.com/calendar/render?action=TEMPLATE'
      +'&text='+encodeURIComponent(name+' — '+(data.service||'Appointment'))
      +'&dates='+encodeURIComponent(fmt(s)+'/'+fmt(en))
      +'&details='+encodeURIComponent('Service: '+(data.service||'')+'\nClinic: '+name+'\nAddress: '+addr+'\nPhone: '+phone)
      +'&location='+encodeURIComponent(addr)+'&sf=true';
  } catch (e) { return ''; }
}

function getCalendarBusy(calendarId, date, slotDurMins) {
  try {
    var cal = CalendarApp.getCalendarById(calendarId);
    if (!cal) { Logger.log('Calendar not found: ' + calendarId); return []; }
    var dayStart = new Date(date.getFullYear(), date.getMonth(), date.getDate(), 0, 0, 0);
    var dayEnd   = new Date(date.getFullYear(), date.getMonth(), date.getDate(), 23, 59, 59);
    var events   = cal.getEvents(dayStart, dayEnd);
    return events.map(function(ev) {
      return { start: ev.getStartTime(), end: ev.getEndTime() };
    });
  } catch (err) {
    Logger.log('Calendar error: ' + err);
    return [];
  }
}

// ─────────────────────────────────────────────────────────────
// Config + Services (unchanged)
// ─────────────────────────────────────────────────────────────

function getServices() {
  try {
    var ss    = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('Services');
    if (!sheet || sheet.getLastRow() === 0) return { services: [] };
    var rows = sheet.getRange(1, 1, sheet.getLastRow(), 1).getValues();
    return { services: rows.map(function(r){ return (r[0]||'').toString().trim(); }).filter(Boolean) };
  } catch (err) { return { services: [], error: err.toString() }; }
}

function getPublicConfig() {
  var c = getConfigObject();
  return { config: {
    CLINIC_NAME:        c.CLINIC_NAME        || 'TomFord Dental',
    CLINIC_ADDRESS:     c.CLINIC_ADDRESS     || '',
    CLINIC_PHONE:       c.CLINIC_PHONE       || '',
    CLINIC_EMAIL:       c.CLINIC_EMAIL       || '',
    CLINIC_HOURS:       c.CLINIC_HOURS       || '',
    BOOKING_TAGLINE:    c.BOOKING_TAGLINE    || 'where every tooth matters.',
    SLOT_DURATION_MINS: c.SLOT_DURATION_MINS || '30',
    OPEN_TIME:          c.OPEN_TIME          || '09:00',
    CLOSE_TIME:         c.CLOSE_TIME         || '19:00'
  }};
}

function getConfigObject() {
  try {
    var ss    = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('Config');
    if (!sheet || sheet.getLastRow() === 0) return {};
    var rows = sheet.getRange(1, 1, sheet.getLastRow(), 2).getValues();
    var cfg  = {};
    rows.forEach(function(row) {
      var k = (row[0]||'').toString().trim();
      var v = (row[1]||'').toString().trim();
      if (k) cfg[k] = v;
    });
    return cfg;
  } catch (err) { return {}; }
}

// ─────────────────────────────────────────────────────────────
// Date / time helpers (unchanged)
// ─────────────────────────────────────────────────────────────

function generateSlots(openTime, closeTime, durationMins) {
  var slots = [];
  var oArr  = openTime.split(':').map(Number);
  var cArr  = closeTime.split(':').map(Number);
  var cur   = oArr[0]*60 + oArr[1];
  var end   = cArr[0]*60 + cArr[1];
  while (cur < end) {
    slots.push(pad2(Math.floor(cur/60)) + ':' + pad2(cur%60));
    cur += durationMins;
  }
  return slots;
}

function parseOpenDays(s) {
  var NAMES = {
    sun:0,sunday:0, mon:1,monday:1, tue:2,tues:2,tuesday:2,
    wed:3,weds:3,wednesday:3, thu:4,thur:4,thurs:4,thursday:4,
    fri:5,friday:5, sat:6,saturday:6
  };
  return s.split(',').map(function(tok) {
    var t = (tok||'').toString().trim().toLowerCase();
    if (!t) return -1;
    if (/^\d+$/.test(t)) return parseInt(t, 10);
    return (t in NAMES) ? NAMES[t] : -1;
  }).filter(function(n){ return n >= 0 && n <= 6; });
}

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

function normalizeSheetTime(v) {
  if (v === null || v === undefined || v === '') return '';
  if (Object.prototype.toString.call(v) === '[object Date]') {
    return pad2(v.getHours()) + ':' + pad2(v.getMinutes());
  }
  if (typeof v === 'number') {
    var totalMins = Math.round(v * 24 * 60);
    return pad2(Math.floor(totalMins/60)) + ':' + pad2(totalMins%60);
  }
  var s = v.toString().trim();
  var m = s.match(/^(\d{1,2}):(\d{2})/);
  if (m) return pad2(parseInt(m[1],10)) + ':' + pad2(parseInt(m[2],10));
  return s;
}

function getTodayPHT() {
  var now = new Date();
  var utc = now.getTime() + now.getTimezoneOffset() * 60000;
  var pht = new Date(utc + 8 * 60 * 60000);
  return pht.getFullYear() + '-' + pad2(pht.getMonth()+1) + '-' + pad2(pht.getDate());
}

function getNowPHT() {
  var now = new Date();
  var utc = now.getTime() + now.getTimezoneOffset() * 60000;
  return new Date(utc + 8 * 60 * 60000);
}

function formatDateStr(d) {
  return d.getFullYear() + '-' + pad2(d.getMonth()+1) + '-' + pad2(d.getDate());
}

function pad2(n) { return String(n).padStart(2, '0'); }

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

// ─────────────────────────────────────────────────────────────
// Debug dump (unchanged)
// ─────────────────────────────────────────────────────────────

function debugDump(e) {
  var dateStr = (e && e.parameter && e.parameter.date) || getTodayPHT();
  var ss      = SpreadsheetApp.getActiveSpreadsheet();
  var sheet   = ss.getSheetByName('Appointments');
  var rows    = [];
  if (sheet && sheet.getLastRow() > 1) {
    var numCols = Math.min(10, sheet.getLastColumn());
    var data = sheet.getRange(2, 1, sheet.getLastRow()-1, numCols).getValues();
    data.forEach(function(row, i) {
      rows.push({
        row: i+2,
        rawDate: String(row[5]),
        normalizedDate: normalizeSheetDate(row[5]),
        rawTime: String(row[6]),
        normalizedTime: normalizeSheetTime(row[6]),
        status: row[8] || '(legacy→Approved)',
        bookingId: row[9] || '',
        matchesQueryDate: normalizeSheetDate(row[5]) === dateStr
      });
    });
  }
  return { queryDate: dateStr, allRows: rows };
}
