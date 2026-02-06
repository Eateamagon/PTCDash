/**
 * PTC Sign-Up Dashboard — Google Apps Script Web App
 *
 * Data lives in the main data sheet (first sheet).
 * Status tracking lives in a "StatusTracking" sheet tab (auto-created).
 * CSV uploads merge into the data sheet: fill blanks, never overwrite.
 *
 * Structure: 2 slots per grade (6th, 7th, 8th) per 15-min block.
 *
 * Admin: etruslow@waynesboro.k12.va.us
 * Everyone else: read-only
 */

var SPREADSHEET_ID = '1FhnS8B4GKz3vA3COT0RGqJpKz4AdDf28Tq-zfvDV8sc';
var ADMIN_EMAIL = 'etruslow@waynesboro.k12.va.us';
var STATUS_SHEET_NAME = 'StatusTracking';
var DATA_SHEET_NAME = null; // null = use first sheet

// Canonical header order for the data sheet
var HEADERS = [
  'Sign Up',
  'Start Date/Time (mm/dd/yyyy)',
  'End Date/Time (mm/dd/yyyy)',
  'Location',
  'Qty',
  'Item',
  'First Name',
  'Last Name',
  'Email',
  'Sign Up Comment',
  'Sign Up Coleader',
  'Sign Up Timestamp'
];

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
  if (!email) {
    email = Session.getEffectiveUser().getEmail();
  }
  var role = (email.toLowerCase() === ADMIN_EMAIL.toLowerCase()) ? 'admin' : 'readonly';
  return { email: email, role: role };
}

// ── Data Access ────────────────────────────────────────────────────

function getDataSheet(ss) {
  if (DATA_SHEET_NAME) {
    return ss.getSheetByName(DATA_SHEET_NAME) || ss.getSheets()[0];
  }
  return ss.getSheets()[0];
}

/**
 * Get all sign-up data merged with status tracking.
 * @param {string|null} gradeFilter - e.g. "6th Grade" or null for all
 * @return {Object} { signups: [...], summary: {...} }
 */
function getSignups(gradeFilter) {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = getDataSheet(ss);
  var data = sheet.getDataRange().getValues();

  if (data.length < 2) return { signups: [], summary: buildEmptySummary() };

  var headers = data[0];
  var colMap = buildColumnMap(headers);
  var statusMap = loadStatusMap(ss);

  // Track slot positions for building unique keys
  var groupCounts = {};

  var signups = [];
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var email = str(valAt(row, colMap, 'email'));
    var firstName = str(valAt(row, colMap, 'first_name'));
    var item = str(valAt(row, colMap, 'item'));
    var startDt = valAt(row, colMap, 'start_datetime');
    var endDt = valAt(row, colMap, 'end_datetime');

    // Skip completely empty rows (no item means not a real slot)
    if (!item) continue;

    var startStr = formatDateVal(startDt);
    var endStr = formatDateVal(endDt);

    // Build unique key using slot position within (start, end, item) group
    var groupKey = startStr + '|||' + endStr + '|||' + item;
    if (!groupCounts[groupKey]) groupCounts[groupKey] = 0;
    groupCounts[groupKey]++;
    var key = groupKey + '|||slot_' + groupCounts[groupKey];
    var status = statusMap[key] || 'none';

    if (gradeFilter && item !== gradeFilter) continue;

    signups.push({
      rowIndex: i + 1,
      key: key,
      sign_up: str(valAt(row, colMap, 'sign_up')),
      start_datetime: startStr,
      end_datetime: endStr,
      location: str(valAt(row, colMap, 'location')),
      qty: valAt(row, colMap, 'qty') || 1,
      item: item,
      first_name: firstName,
      last_name: str(valAt(row, colMap, 'last_name')),
      email: email,
      signup_comment: str(valAt(row, colMap, 'signup_comment')),
      sign_up_coleader: str(valAt(row, colMap, 'sign_up_coleader')),
      signup_timestamp: formatDateVal(valAt(row, colMap, 'signup_timestamp')),
      status: status,
      is_empty_slot: !email && !firstName
    });
  }

  // Sort: start time -> grade -> slot (filled first, empty last)
  signups.sort(function(a, b) {
    if (a.start_datetime < b.start_datetime) return -1;
    if (a.start_datetime > b.start_datetime) return 1;
    if (a.item < b.item) return -1;
    if (a.item > b.item) return 1;
    // Filled slots before empty
    if (a.is_empty_slot !== b.is_empty_slot) return a.is_empty_slot ? 1 : -1;
    var nameA = (a.signup_comment || '').toLowerCase();
    var nameB = (b.signup_comment || '').toLowerCase();
    return nameA < nameB ? -1 : nameA > nameB ? 1 : 0;
  });

  var summary = buildSummary(signups);
  return { signups: signups, summary: summary };
}

// ── CSV Upload & Merge ─────────────────────────────────────────────

/**
 * Build a slot-position index for rows.
 * Each row in a (start, end, item) group is numbered sequentially (slot 1, slot 2, etc.).
 * This ensures that two empty slots for the same grade/time don't collide.
 *
 * @param {Array[]} data - 2D array of sheet values (includes header row)
 * @param {Object} colMap - column mapping from buildColumnMap
 * @return {Object} { positionKey -> dataIndex (0-based) }
 */
function buildSlotIndex(data, colMap) {
  var groupCounts = {}; // "start|||end|||item" -> count seen so far
  var index = {};       // "start|||end|||item|||slotNum" -> data row index

  for (var i = 1; i < data.length; i++) {
    var r = data[i];
    var rStart = formatDateVal(valAt(r, colMap, 'start_datetime'));
    var rEnd = formatDateVal(valAt(r, colMap, 'end_datetime'));
    var rItem = str(valAt(r, colMap, 'item'));

    if (!rItem) continue;

    var groupKey = rStart + '|||' + rEnd + '|||' + rItem;
    if (!groupCounts[groupKey]) groupCounts[groupKey] = 0;
    groupCounts[groupKey]++;

    var posKey = groupKey + '|||' + groupCounts[groupKey];
    index[posKey] = i; // 0-based data index
  }

  return index;
}

/**
 * Upload CSV content and merge into data sheet.
 *
 * Matching strategy: rows are matched by SLOT POSITION within each
 * (start_datetime, end_datetime, item) group. E.g. the 1st "8th Grade"
 * row at 12:00-12:15 in the CSV matches the 1st "8th Grade" row at
 * 12:00-12:15 in the sheet, the 2nd matches the 2nd, etc.
 *
 * This handles empty slots correctly — even if both slots for a grade
 * are blank (no email/name), they each match to their positional
 * counterpart in the sheet.
 *
 * For matched rows: only blank cells get filled. Existing data is never
 * overwritten, so you can upload partial exports without losing anything.
 *
 * @param {string} csvContent - raw CSV text from the uploaded file
 * @return {Object} { inserted, updated, skipped }
 */
function uploadCSV(csvContent) {
  var user = getCurrentUser();
  if (user.role !== 'admin') {
    throw new Error('Admin access required');
  }

  var rows = parseCSV(csvContent);
  if (rows.length === 0) {
    return { inserted: 0, updated: 0, skipped: 0 };
  }

  var csvHeaders = rows[0];
  var csvColMap = buildColumnMap(csvHeaders);

  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = getDataSheet(ss);
  var sheetData = sheet.getDataRange().getValues();

  // If the sheet is empty, write headers first
  if (sheetData.length === 0) {
    sheet.appendRow(HEADERS);
    sheetData = [HEADERS];
  }

  var sheetHeaders = sheetData[0];
  var sheetColMap = buildColumnMap(sheetHeaders);

  // Build slot-position index for existing sheet rows
  var slotIndex = buildSlotIndex(sheetData, sheetColMap);

  // Track CSV group counts to assign slot positions
  var csvGroupCounts = {};

  var inserted = 0, updated = 0, skipped = 0;

  for (var c = 1; c < rows.length; c++) {
    var csvRow = rows[c];

    var cSignUp = csvVal(csvRow, csvColMap, 'sign_up');
    var cStart = csvVal(csvRow, csvColMap, 'start_datetime');
    var cEnd = csvVal(csvRow, csvColMap, 'end_datetime');
    var cLocation = csvVal(csvRow, csvColMap, 'location');
    var cQty = csvVal(csvRow, csvColMap, 'qty');
    var cItem = csvVal(csvRow, csvColMap, 'item');
    var cFirst = csvVal(csvRow, csvColMap, 'first_name');
    var cLast = csvVal(csvRow, csvColMap, 'last_name');
    var cEmail = csvVal(csvRow, csvColMap, 'email');
    var cComment = csvVal(csvRow, csvColMap, 'signup_comment');
    var cColeader = csvVal(csvRow, csvColMap, 'sign_up_coleader');
    var cTimestamp = csvVal(csvRow, csvColMap, 'signup_timestamp');

    // Skip rows with no item (not a real slot)
    if (!cItem) { skipped++; continue; }

    // Determine this CSV row's slot position within its group
    var groupKey = cStart + '|||' + cEnd + '|||' + cItem;
    if (!csvGroupCounts[groupKey]) csvGroupCounts[groupKey] = 0;
    csvGroupCounts[groupKey]++;
    var slotNum = csvGroupCounts[groupKey];

    var posKey = groupKey + '|||' + slotNum;

    if (slotIndex.hasOwnProperty(posKey)) {
      // Matched to existing sheet row by slot position — fill blanks only
      var dataIdx = slotIndex[posKey];
      var sheetRow = sheetData[dataIdx];
      var rowNum = dataIdx + 1; // 1-based sheet row
      var changed = false;

      changed = mergeCell(sheet, sheetRow, sheetColMap, 'sign_up', cSignUp, rowNum) || changed;
      changed = mergeCell(sheet, sheetRow, sheetColMap, 'location', cLocation, rowNum) || changed;
      changed = mergeCell(sheet, sheetRow, sheetColMap, 'qty', cQty, rowNum) || changed;
      changed = mergeCell(sheet, sheetRow, sheetColMap, 'first_name', cFirst, rowNum) || changed;
      changed = mergeCell(sheet, sheetRow, sheetColMap, 'last_name', cLast, rowNum) || changed;
      changed = mergeCell(sheet, sheetRow, sheetColMap, 'email', cEmail, rowNum) || changed;
      changed = mergeCell(sheet, sheetRow, sheetColMap, 'signup_comment', cComment, rowNum) || changed;
      changed = mergeCell(sheet, sheetRow, sheetColMap, 'sign_up_coleader', cColeader, rowNum) || changed;
      changed = mergeCell(sheet, sheetRow, sheetColMap, 'signup_timestamp', cTimestamp, rowNum) || changed;

      if (changed) updated++;
      else skipped++;
    } else {
      // No matching slot position in sheet — append as new row
      var newRow = buildNewRow(sheetHeaders, sheetColMap, {
        sign_up: cSignUp,
        start_datetime: cStart,
        end_datetime: cEnd,
        location: cLocation,
        qty: cQty,
        item: cItem,
        first_name: cFirst,
        last_name: cLast,
        email: cEmail,
        signup_comment: cComment,
        sign_up_coleader: cColeader,
        signup_timestamp: cTimestamp
      });
      sheet.appendRow(newRow);
      inserted++;
    }
  }

  SpreadsheetApp.flush();
  return { inserted: inserted, updated: updated, skipped: skipped };
}

/**
 * Merge a single cell: only write if the sheet cell is blank and the CSV value is not.
 */
function mergeCell(sheet, sheetRow, colMap, field, csvValue, rowNum) {
  if (colMap[field] === undefined) return false;
  var colIdx = colMap[field];
  var existing = str(sheetRow[colIdx]);
  if (!existing && csvValue) {
    sheet.getRange(rowNum, colIdx + 1).setValue(csvValue);
    return true;
  }
  return false;
}

/**
 * Build a new row array matching the sheet's header order.
 */
function buildNewRow(headers, colMap, values) {
  var row = [];
  for (var i = 0; i < headers.length; i++) {
    row.push('');
  }
  var fieldMap = {
    'sign_up': values.sign_up,
    'start_datetime': values.start_datetime,
    'end_datetime': values.end_datetime,
    'location': values.location,
    'qty': values.qty,
    'item': values.item,
    'first_name': values.first_name,
    'last_name': values.last_name,
    'email': values.email,
    'signup_comment': values.signup_comment,
    'sign_up_coleader': values.sign_up_coleader,
    'signup_timestamp': values.signup_timestamp
  };
  for (var field in fieldMap) {
    if (colMap[field] !== undefined) {
      row[colMap[field]] = fieldMap[field] || '';
    }
  }
  return row;
}

// ── CSV Parser ─────────────────────────────────────────────────────

/**
 * Simple CSV parser that handles quoted fields with commas/newlines.
 */
function parseCSV(text) {
  var rows = [];
  var row = [];
  var field = '';
  var inQuotes = false;
  var i = 0;

  while (i < text.length) {
    var ch = text[i];

    if (inQuotes) {
      if (ch === '"') {
        if (i + 1 < text.length && text[i + 1] === '"') {
          field += '"';
          i += 2;
        } else {
          inQuotes = false;
          i++;
        }
      } else {
        field += ch;
        i++;
      }
    } else {
      if (ch === '"') {
        inQuotes = true;
        i++;
      } else if (ch === ',') {
        row.push(field.trim());
        field = '';
        i++;
      } else if (ch === '\r' || ch === '\n') {
        row.push(field.trim());
        field = '';
        if (ch === '\r' && i + 1 < text.length && text[i + 1] === '\n') {
          i++;
        }
        if (row.length > 1 || (row.length === 1 && row[0] !== '')) {
          rows.push(row);
        }
        row = [];
        i++;
      } else {
        field += ch;
        i++;
      }
    }
  }

  // Last field/row
  if (field || row.length > 0) {
    row.push(field.trim());
    if (row.length > 1 || (row.length === 1 && row[0] !== '')) {
      rows.push(row);
    }
  }

  return rows;
}

function csvVal(row, colMap, field) {
  if (colMap[field] === undefined) return '';
  if (colMap[field] >= row.length) return '';
  return (row[colMap[field]] || '').toString().trim();
}

// ── Status ─────────────────────────────────────────────────────────

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

  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === key) {
      if (status === 'none') {
        statusSheet.deleteRow(i + 1);
      } else {
        statusSheet.getRange(i + 1, 2).setValue(status);
        statusSheet.getRange(i + 1, 3).setValue(new Date());
      }
      return { ok: true };
    }
  }

  if (status !== 'none') {
    statusSheet.appendRow([key, status, new Date()]);
  }

  return { ok: true };
}

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
    if (h === 'sign up' || h === 'sign-up' || h === 'sign ups' || h === 'sign-ups') map['sign_up'] = i;
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

function str(val) {
  if (val === null || val === undefined) return '';
  if (val instanceof Date) return formatDateVal(val);
  return String(val).trim();
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
  return String(val).trim();
}

// ── Summary ────────────────────────────────────────────────────────

function buildSummary(signups) {
  var grades = {};
  var totals = { total: 0, in_building: 0, late: 0, cancelled: 0, pending: 0, open_slots: 0 };

  for (var i = 0; i < signups.length; i++) {
    var s = signups[i];
    var g = s.item;

    if (!grades[g]) {
      grades[g] = { grade: g, total: 0, in_building: 0, late: 0, cancelled: 0, pending: 0, open_slots: 0 };
    }

    if (s.is_empty_slot) {
      grades[g].open_slots++;
      totals.open_slots++;
      continue;
    }

    grades[g].total++;
    totals.total++;

    if (s.status === 'in_building') { grades[g].in_building++; totals.in_building++; }
    else if (s.status === 'late') { grades[g].late++; totals.late++; }
    else if (s.status === 'cancel') { grades[g].cancelled++; totals.cancelled++; }
    else { grades[g].pending++; totals.pending++; }
  }

  var gradeList = Object.keys(grades).sort();
  var gradeArray = gradeList.map(function(k) { return grades[k]; });

  return { grades: gradeArray, totals: totals };
}

function buildEmptySummary() {
  return { grades: [], totals: { total: 0, in_building: 0, late: 0, cancelled: 0, pending: 0, open_slots: 0 } };
}
