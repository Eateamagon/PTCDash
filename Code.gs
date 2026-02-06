/**
 * PTC Sign-Up Dashboard — Google Apps Script Web App
 *
 * Data lives in the "SIGN-UPS" sheet (or first sheet).
 * Status tracking lives in a "StatusTracking" sheet tab (auto-created).
 *
 * Admin: etruslow@waynesboro.k12.va.us
 * Everyone else: read-only
 */

var SPREADSHEET_ID = '1FhnS8B4GKz3vA3COT0RGqJpKz4AdDf28Tq-zfvDV8sc';
var ADMIN_EMAIL = 'etruslow@waynesboro.k12.va.us';
var STATUS_SHEET_NAME = 'StatusTracking';

// ── Web App Entry ──────────────────────────────────────────────────

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('PTC Sign-Up Dashboard')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// ── Auth ───────────────────────────────────────────────────────────

function getCurrentUser() {
  var email = Session.getActiveUser().getEmail();
  // If email is empty (common in "anyone" deployments), fall back to effective user
  if (!email) {
    email = Session.getEffectiveUser().getEmail();
  }
  var role = (email.toLowerCase() === ADMIN_EMAIL.toLowerCase()) ? 'admin' : 'readonly';
  return { email: email, role: role };
}

// ── Data Access ────────────────────────────────────────────────────

/**
 * Get all sign-up data merged with status tracking.
 * @param {string|null} gradeFilter - e.g. "6th Grade" or null for all
 * @return {Object} { signups: [...], summary: {...} }
 */
function getSignups(gradeFilter) {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheets()[0]; // First sheet has the sign-up data
  var data = sheet.getDataRange().getValues();

  if (data.length < 2) return { signups: [], summary: buildEmptySummary() };

  var headers = data[0];
  var colMap = buildColumnMap(headers);

  // Load status tracking
  var statusMap = loadStatusMap(ss);

  var signups = [];
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var email = valAt(row, colMap, 'email');
    var firstName = valAt(row, colMap, 'first_name');

    // Skip empty rows
    if (!email && !firstName) continue;

    var startDt = valAt(row, colMap, 'start_datetime');
    var endDt = valAt(row, colMap, 'end_datetime');
    var item = valAt(row, colMap, 'item');

    // Build unique key for status tracking
    var key = buildKey(startDt, endDt, item, email);
    var status = statusMap[key] || 'none';

    var grade = item;
    if (gradeFilter && grade !== gradeFilter) continue;

    signups.push({
      rowIndex: i + 1, // 1-based sheet row
      key: key,
      sign_up: valAt(row, colMap, 'sign_up'),
      start_datetime: formatDateVal(startDt),
      end_datetime: formatDateVal(endDt),
      location: valAt(row, colMap, 'location'),
      qty: valAt(row, colMap, 'qty') || 1,
      item: grade,
      first_name: firstName,
      last_name: valAt(row, colMap, 'last_name'),
      email: email,
      signup_comment: valAt(row, colMap, 'signup_comment'),
      sign_up_coleader: valAt(row, colMap, 'sign_up_coleader'),
      signup_timestamp: formatDateVal(valAt(row, colMap, 'signup_timestamp')),
      status: status
    });
  }

  // Sort by start time then by student name (signup_comment)
  signups.sort(function(a, b) {
    if (a.start_datetime < b.start_datetime) return -1;
    if (a.start_datetime > b.start_datetime) return 1;
    if (a.item < b.item) return -1;
    if (a.item > b.item) return 1;
    var nameA = (a.signup_comment || '').toLowerCase();
    var nameB = (b.signup_comment || '').toLowerCase();
    return nameA < nameB ? -1 : nameA > nameB ? 1 : 0;
  });

  var summary = buildSummary(signups);
  return { signups: signups, summary: summary };
}

/**
 * Update a student's status. Admin only.
 * @param {string} key - unique row key
 * @param {string} status - 'none', 'in_building', 'late', 'cancel'
 */
function updateStatus(key, status) {
  var user = getCurrentUser();
  if (user.role !== 'admin') {
    throw new Error('Admin access required');
  }

  var validStatuses = ['none', 'in_building', 'late', 'cancel'];
  if (validStatuses.indexOf(status) === -1) {
    throw new Error('Invalid status: ' + status);
  }

  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var statusSheet = getOrCreateStatusSheet(ss);
  var data = statusSheet.getDataRange().getValues();

  // Find existing row for this key
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === key) {
      if (status === 'none') {
        // Remove the row
        statusSheet.deleteRow(i + 1);
      } else {
        statusSheet.getRange(i + 1, 2).setValue(status);
        statusSheet.getRange(i + 1, 3).setValue(new Date());
      }
      return { ok: true };
    }
  }

  // Not found — add new row (only if not 'none')
  if (status !== 'none') {
    statusSheet.appendRow([key, status, new Date()]);
  }

  return { ok: true };
}

/**
 * Manually refresh data — just returns fresh data (sheet is always live).
 * Kept as explicit action so admin knows data is current.
 */
function refreshData(gradeFilter) {
  SpreadsheetApp.flush();
  return getSignups(gradeFilter);
}

// ── Status Sheet Helpers ───────────────────────────────────────────

function getOrCreateStatusSheet(ss) {
  var sheet = ss.getSheetByName(STATUS_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(STATUS_SHEET_NAME);
    sheet.appendRow(['Key', 'Status', 'Updated']);
    sheet.setColumnWidth(1, 400);
  }
  return sheet;
}

function loadStatusMap(ss) {
  var sheet = ss.getSheetByName(STATUS_SHEET_NAME);
  var map = {};
  if (!sheet) return map;

  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0]) {
      map[data[i][0]] = data[i][1];
    }
  }
  return map;
}

function buildKey(startDt, endDt, item, email) {
  return [startDt, endDt, item, email].join('|||');
}

// ── Column Mapping ─────────────────────────────────────────────────

function buildColumnMap(headers) {
  var map = {};
  for (var i = 0; i < headers.length; i++) {
    var h = String(headers[i]).trim().toLowerCase();
    if (h === 'sign up' || h === 'sign-up') map['sign_up'] = i;
    else if (h.indexOf('start date') !== -1) map['start_datetime'] = i;
    else if (h.indexOf('end date') !== -1) map['end_datetime'] = i;
    else if (h === 'location') map['location'] = i;
    else if (h === 'qty') map['qty'] = i;
    else if (h === 'item') map['item'] = i;
    else if (h === 'first name') map['first_name'] = i;
    else if (h === 'last name') map['last_name'] = i;
    else if (h === 'email') map['email'] = i;
    else if (h.indexOf('sign up comment') !== -1 || h.indexOf('signup comment') !== -1) map['signup_comment'] = i;
    else if (h.indexOf('coleader') !== -1 || h.indexOf('co-leader') !== -1) map['sign_up_coleader'] = i;
    else if (h.indexOf('timestamp') !== -1) map['signup_timestamp'] = i;
  }
  return map;
}

function valAt(row, colMap, field) {
  if (colMap[field] === undefined) return '';
  var val = row[colMap[field]];
  if (val === null || val === undefined) return '';
  return val;
}

function formatDateVal(val) {
  if (!val) return '';
  if (val instanceof Date) {
    var m = val.getMonth() + 1;
    var d = val.getDate();
    var y = val.getFullYear();
    var hh = val.getHours();
    var mm = val.getMinutes();
    return m + '/' + d + '/' + y + ' ' + hh + ':' + (mm < 10 ? '0' + mm : mm);
  }
  return String(val);
}

// ── Summary ────────────────────────────────────────────────────────

function buildSummary(signups) {
  var grades = {};
  var totals = { total: 0, in_building: 0, late: 0, cancelled: 0, pending: 0 };

  for (var i = 0; i < signups.length; i++) {
    var s = signups[i];
    var g = s.item;

    if (!grades[g]) {
      grades[g] = { grade: g, total: 0, in_building: 0, late: 0, cancelled: 0, pending: 0 };
    }

    grades[g].total++;
    totals.total++;

    if (s.status === 'in_building') { grades[g].in_building++; totals.in_building++; }
    else if (s.status === 'late') { grades[g].late++; totals.late++; }
    else if (s.status === 'cancel') { grades[g].cancelled++; totals.cancelled++; }
    else { grades[g].pending++; totals.pending++; }
  }

  // Sort grade keys naturally
  var gradeList = Object.keys(grades).sort();
  var gradeArray = gradeList.map(function(k) { return grades[k]; });

  return { grades: gradeArray, totals: totals };
}

function buildEmptySummary() {
  return { grades: [], totals: { total: 0, in_building: 0, late: 0, cancelled: 0, pending: 0 } };
}
