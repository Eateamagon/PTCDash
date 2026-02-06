const Database = require('better-sqlite3');
const path = require('path');
const fs = require('fs');

const dataDir = path.join(__dirname, 'data');
fs.mkdirSync(dataDir, { recursive: true });

const dbPath = path.join(dataDir, 'ptcdash.db');
const db = new Database(dbPath);
db.pragma('journal_mode = WAL');

// Signups table holds data synced from Google Sheet + manual CSV uploads
// Status is tracked locally (not in the sheet)
db.exec(`
  CREATE TABLE IF NOT EXISTS signups (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    sign_up TEXT,
    start_datetime TEXT,
    end_datetime TEXT,
    location TEXT,
    qty INTEGER DEFAULT 1,
    item TEXT,
    first_name TEXT,
    last_name TEXT,
    email TEXT,
    signup_comment TEXT,
    sign_up_coleader TEXT,
    signup_timestamp TEXT,
    status TEXT DEFAULT 'none',
    created_at TEXT DEFAULT (datetime('now')),
    updated_at TEXT DEFAULT (datetime('now')),
    UNIQUE(start_datetime, end_datetime, item, email)
  )
`);

// Upsert: fills in blanks without overwriting existing data or status
const upsertStmt = db.prepare(`
  INSERT INTO signups (sign_up, start_datetime, end_datetime, location, qty, item,
    first_name, last_name, email, signup_comment, sign_up_coleader, signup_timestamp)
  VALUES (@sign_up, @start_datetime, @end_datetime, @location, @qty, @item,
    @first_name, @last_name, @email, @signup_comment, @sign_up_coleader, @signup_timestamp)
  ON CONFLICT(start_datetime, end_datetime, item, email) DO UPDATE SET
    sign_up = CASE WHEN excluded.sign_up != '' AND excluded.sign_up IS NOT NULL THEN excluded.sign_up ELSE signups.sign_up END,
    location = CASE WHEN excluded.location != '' AND excluded.location IS NOT NULL THEN excluded.location ELSE signups.location END,
    qty = CASE WHEN excluded.qty IS NOT NULL THEN excluded.qty ELSE signups.qty END,
    first_name = CASE WHEN excluded.first_name != '' AND excluded.first_name IS NOT NULL THEN excluded.first_name ELSE signups.first_name END,
    last_name = CASE WHEN excluded.last_name != '' AND excluded.last_name IS NOT NULL THEN excluded.last_name ELSE signups.last_name END,
    signup_comment = CASE WHEN excluded.signup_comment != '' AND excluded.signup_comment IS NOT NULL THEN excluded.signup_comment ELSE signups.signup_comment END,
    sign_up_coleader = CASE WHEN excluded.sign_up_coleader != '' AND excluded.sign_up_coleader IS NOT NULL THEN excluded.sign_up_coleader ELSE signups.sign_up_coleader END,
    signup_timestamp = CASE WHEN excluded.signup_timestamp != '' AND excluded.signup_timestamp IS NOT NULL THEN excluded.signup_timestamp ELSE signups.signup_timestamp END,
    updated_at = datetime('now')
`);

const findExisting = db.prepare(
  'SELECT id FROM signups WHERE start_datetime = @start_datetime AND end_datetime = @end_datetime AND item = @item AND email = @email'
);

function normalizeRow(row) {
  return {
    sign_up: row['Sign Up'] || row['sign_up'] || '',
    start_datetime: row['Start Date/Time (mm/dd/yyyy)'] || row['Start Date/Time'] || row['start_datetime'] || '',
    end_datetime: row['End Date/Time (mm/dd/yyyy)'] || row['End Date/Time'] || row['end_datetime'] || '',
    location: row['Location'] || row['location'] || '',
    qty: parseInt(row['Qty'] || row['qty'] || '1', 10) || 1,
    item: row['Item'] || row['item'] || '',
    first_name: row['First Name'] || row['first_name'] || '',
    last_name: row['Last Name'] || row['last_name'] || '',
    email: row['Email'] || row['email'] || '',
    signup_comment: row['Sign Up Comment'] || row['Signup Comment'] || row['signup_comment'] || '',
    sign_up_coleader: row['Sign Up Coleader'] || row['sign_up_coleader'] || '',
    signup_timestamp: row['Sign Up Timestamp'] || row['Signup Timestamp'] || row['signup_timestamp'] || ''
  };
}

function upsertSignup(row) {
  const normalized = normalizeRow(row);
  // Skip rows with no email (empty slot rows from sheet)
  if (!normalized.email && !normalized.first_name) return 'skipped';

  const existing = findExisting.get({
    start_datetime: normalized.start_datetime,
    end_datetime: normalized.end_datetime,
    item: normalized.item,
    email: normalized.email
  });

  upsertStmt.run(normalized);
  return existing ? 'updated' : 'inserted';
}

function getSignups(grade) {
  if (grade) {
    return db.prepare('SELECT * FROM signups WHERE item = ? ORDER BY start_datetime, signup_comment').all(grade);
  }
  return db.prepare('SELECT * FROM signups ORDER BY start_datetime, item, signup_comment').all();
}

function getSummary() {
  const grades = db.prepare(`
    SELECT item as grade,
      COUNT(*) as total,
      SUM(CASE WHEN status = 'in_building' THEN 1 ELSE 0 END) as in_building,
      SUM(CASE WHEN status = 'late' THEN 1 ELSE 0 END) as late,
      SUM(CASE WHEN status = 'cancel' THEN 1 ELSE 0 END) as cancelled,
      SUM(CASE WHEN status = 'none' THEN 1 ELSE 0 END) as pending
    FROM signups
    GROUP BY item
    ORDER BY item
  `).all();

  const totals = db.prepare(`
    SELECT
      COUNT(*) as total,
      SUM(CASE WHEN status = 'in_building' THEN 1 ELSE 0 END) as in_building,
      SUM(CASE WHEN status = 'late' THEN 1 ELSE 0 END) as late,
      SUM(CASE WHEN status = 'cancel' THEN 1 ELSE 0 END) as cancelled,
      SUM(CASE WHEN status = 'none' THEN 1 ELSE 0 END) as pending
    FROM signups
  `).get();

  return { grades, totals };
}

function updateStatus(id, status) {
  db.prepare('UPDATE signups SET status = ?, updated_at = datetime(?) WHERE id = ?')
    .run(status, new Date().toISOString(), id);
}

module.exports = { upsertSignup, getSignups, getSummary, updateStatus };
