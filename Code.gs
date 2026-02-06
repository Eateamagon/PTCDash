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
 * Normalize any date value to a consistent string for comparison.
 * Handles: Date objects, "2/16/2026 12:00", "2/16/2026 12:00:00", etc.
 */
function normDate(val) {
  if (!val) return '';
  if (val instanceof Date) {
    var m = val.getMonth() + 1;
    var d = val.getDate();
    var y = val.getFullYear();
    var hh = val.getHours();
    var mm = val.getMinutes();
    return m + '/' + d + '/' + y + ' ' + hh + ':' + (mm < 10 ? '0' + mm : mm);
  }
  // String: strip seconds if present, trim
  var s = String(val).trim();
  // "2/16/2026 12:00:00" -> "2/16/2026 12:00"
  var match = s.match(/^(\d{1,2}\/\d{1,2}\/\d{4}\s+\d{1,2}:\d{2})(:\d{2})?$/);
  if (match) return match[1];
  return s;
}

/**
 * Get all sign-up data merged with status tracking.
 */
function getSignups(gradeFilter) {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = getDataSheet(ss);
  var data = sheet.getDataRange().getValues();

  if (data.length < 2) return { signups: [], summary: buildEmptySummary() };

  var headers = data[0];
  var colMap = buildColumnMap(headers);
  var statusMap = loadStatusMap(ss);

  var groupCounts = {};
  var signups = [];

  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var email = str(valAt(row, colMap, 'email'));
    var firstName = str(valAt(row, colMap, 'first_name'));
    var item = str(valAt(row, colMap, 'item'));
    var startDt = valAt(row, colMap, 'start_datetime');
    var endDt = valAt(row, colMap, 'end_datetime');

    if (!item) continue;

    var startStr = normDate(startDt);
    var endStr = normDate(endDt);

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
      signup_timestamp: normDate(valAt(row, colMap, 'signup_timestamp')),
      status: status,
      is_empty_slot: !email && !firstName
    });
  }

  signups.sort(function(a, b) {
    if (a.start_datetime < b.start_datetime) return -1;
    if (a.start_datetime > b.start_datetime) return 1;
    if (a.item < b.item) return -1;
    if (a.item > b.item) return 1;
    if (a.is_empty_slot !== b.is_empty_slot) return a.is_empty_slot ? 1 : -1;
    var nameA = (a.signup_comment || '').toLowerCase();
    var nameB = (b.signup_comment || '').toLowerCase();
    return nameA < nameB ? -1 : nameA > nameB ? 1 : 0;
  });

  return { signups: signups, summary: buildSummary(signups) };
}

// ── CSV Upload & Merge ─────────────────────────────────────────────

/**
 * Build slot-position index using normalized dates.
 */
function buildSlotIndex(data, colMap) {
  var groupCounts = {};
  var index = {};

  for (var i = 1; i < data.length; i++) {
    var r = data[i];
    var rStart = normDate(valAt(r, colMap, 'start_datetime'));
    var rEnd = normDate(valAt(r, colMap, 'end_datetime'));
    var rItem = str(valAt(r, colMap, 'item'));

    if (!rItem) continue;

    var groupKey = rStart + '|||' + rEnd + '|||' + rItem;
    if (!groupCounts[groupKey]) groupCounts[groupKey] = 0;
    groupCounts[groupKey]++;

    var posKey = groupKey + '|||' + groupCounts[groupKey];
    index[posKey] = i;
  }

  return index;
}

/**
 * Upload CSV content and merge into data sheet.
 * Uses batch writes (setValues) instead of individual setValue calls for speed.
 *
 * Matching: by slot position within (start, end, item) groups.
 * Merge: only fills blank cells, never overwrites.
 */
function uploadCSV(csvContent) {
  var user = getCurrentUser();
  if (user.role !== 'admin') {
    throw new Error('Admin access required');
  }

  if (!csvContent || !csvContent.trim()) {
    throw new Error('CSV file is empty');
  }

  var rows = parseCSV(csvContent);
  if (rows.length < 2) {
    throw new Error('CSV has no data rows (only ' + rows.length + ' row found)');
  }

  var csvHeaders = rows[0];
  var csvColMap = buildColumnMap(csvHeaders);

  // Verify we can at least find the Item column
  if (csvColMap['item'] === undefined) {
    throw new Error('Could not find "Item" column in CSV. Found headers: ' + csvHeaders.join(', '));
  }

  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = getDataSheet(ss);
  var sheetData = sheet.getDataRange().getValues();

  var isNewSheet = (sheetData.length === 0);

  // If sheet is empty, write all CSV data directly as a batch
  if (isNewSheet || sheetData.length < 2) {
    return writeFullCSV(sheet, rows, csvColMap, isNewSheet ? null : sheetData[0]);
  }

  var sheetHeaders = sheetData[0];
  var sheetColMap = buildColumnMap(sheetHeaders);
  var numCols = sheetHeaders.length;

  // Build slot-position index for existing sheet rows
  var slotIndex = buildSlotIndex(sheetData, sheetColMap);

  // Track which sheet rows need updating (collect changes, write in batch)
  var csvGroupCounts = {};
  var rowsToAppend = [];
  var inserted = 0, updated = 0, skipped = 0;

  // Fields we'll try to merge (excluding start/end/item which are the fixed slot identifiers)
  var mergeFields = ['sign_up', 'location', 'qty', 'first_name', 'last_name',
                     'email', 'signup_comment', 'sign_up_coleader', 'signup_timestamp'];

  for (var c = 1; c < rows.length; c++) {
    var csvRow = rows[c];
    var cItem = csvVal(csvRow, csvColMap, 'item');
    var cStart = csvVal(csvRow, csvColMap, 'start_datetime');
    var cEnd = csvVal(csvRow, csvColMap, 'end_datetime');

    if (!cItem) { skipped++; continue; }

    var groupKey = normDate(cStart) + '|||' + normDate(cEnd) + '|||' + cItem;
    if (!csvGroupCounts[groupKey]) csvGroupCounts[groupKey] = 0;
    csvGroupCounts[groupKey]++;
    var slotNum = csvGroupCounts[groupKey];
    var posKey = groupKey + '|||' + slotNum;

    // Extract all CSV values for this row
    var csvValues = {};
    csvValues['sign_up'] = csvVal(csvRow, csvColMap, 'sign_up');
    csvValues['start_datetime'] = cStart;
    csvValues['end_datetime'] = cEnd;
    csvValues['location'] = csvVal(csvRow, csvColMap, 'location');
    csvValues['qty'] = csvVal(csvRow, csvColMap, 'qty');
    csvValues['item'] = cItem;
    csvValues['first_name'] = csvVal(csvRow, csvColMap, 'first_name');
    csvValues['last_name'] = csvVal(csvRow, csvColMap, 'last_name');
    csvValues['email'] = csvVal(csvRow, csvColMap, 'email');
    csvValues['signup_comment'] = csvVal(csvRow, csvColMap, 'signup_comment');
    csvValues['sign_up_coleader'] = csvVal(csvRow, csvColMap, 'sign_up_coleader');
    csvValues['signup_timestamp'] = csvVal(csvRow, csvColMap, 'signup_timestamp');

    if (slotIndex.hasOwnProperty(posKey)) {
      // Matched existing row — check for blanks to fill
      var dataIdx = slotIndex[posKey];
      var sheetRow = sheetData[dataIdx];
      var rowNum = dataIdx + 1;
      var changed = false;

      for (var f = 0; f < mergeFields.length; f++) {
        var field = mergeFields[f];
        if (sheetColMap[field] === undefined) continue;
        var colIdx = sheetColMap[field];
        var existing = str(sheetRow[colIdx]);
        var csvV = csvValues[field];
        if (!existing && csvV) {
          // Update in the in-memory array too so we track what's changed
          sheetData[dataIdx][colIdx] = csvV;
          changed = true;
        }
      }

      if (changed) updated++;
      else skipped++;
    } else {
      // New row — build and queue for append
      var newRow = [];
      for (var n = 0; n < numCols; n++) newRow.push('');
      for (var field2 in csvValues) {
        if (sheetColMap[field2] !== undefined) {
          newRow[sheetColMap[field2]] = csvValues[field2] || '';
        }
      }
      rowsToAppend.push(newRow);
      inserted++;
    }
  }

  // Batch write: update existing rows (write entire data block at once)
  if (updated > 0) {
    var range = sheet.getRange(1, 1, sheetData.length, numCols);
    range.setValues(sheetData);
  }

  // Batch write: append new rows
  if (rowsToAppend.length > 0) {
    var startRow = sheet.getLastRow() + 1;
    var appendRange = sheet.getRange(startRow, 1, rowsToAppend.length, numCols);
    appendRange.setValues(rowsToAppend);
  }

  SpreadsheetApp.flush();
  return { inserted: inserted, updated: updated, skipped: skipped };
}

/**
 * Write an entire CSV to an empty or header-only sheet as a batch.
 */
function writeFullCSV(sheet, csvRows, csvColMap, existingHeaderRow) {
  var headers = existingHeaderRow || HEADERS;

  // Build column map for the target headers
  var targetColMap = buildColumnMap(headers);
  var numCols = headers.length;

  var outputRows = [headers];

  for (var i = 1; i < csvRows.length; i++) {
    var csvRow = csvRows[i];
    var cItem = csvVal(csvRow, csvColMap, 'item');
    if (!cItem) continue;

    var newRow = [];
    for (var n = 0; n < numCols; n++) newRow.push('');

    var allFields = ['sign_up', 'start_datetime', 'end_datetime', 'location', 'qty',
                     'item', 'first_name', 'last_name', 'email', 'signup_comment',
                     'sign_up_coleader', 'signup_timestamp'];
    for (var f = 0; f < allFields.length; f++) {
      var field = allFields[f];
      if (targetColMap[field] !== undefined) {
        newRow[targetColMap[field]] = csvVal(csvRow, csvColMap, field);
      }
    }
    outputRows.push(newRow);
  }

  // Clear and write all at once
  sheet.clear();
  if (outputRows.length > 0) {
    sheet.getRange(1, 1, outputRows.length, numCols).setValues(outputRows);
  }

  SpreadsheetApp.flush();
  return { inserted: outputRows.length - 1, updated: 0, skipped: 0 };
}

// ── CSV Parser ─────────────────────────────────────────────────────

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
  if (val instanceof Date) return normDate(val);
  return String(val).trim();
}

function formatDateVal(val) {
  return normDate(val);
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
