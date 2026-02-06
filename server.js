const express = require('express');
const session = require('express-session');
const multer = require('multer');
const { parse } = require('csv-parse/sync');
const path = require('path');
const https = require('https');
const db = require('./db');

const app = express();
const upload = multer({ storage: multer.memoryStorage() });

app.use(express.json());
app.use(express.urlencoded({ extended: true }));
app.use(express.static(path.join(__dirname, 'public')));
app.use(session({
  secret: 'ptc-dash-secret-key-change-in-production',
  resave: false,
  saveUninitialized: false,
  cookie: { maxAge: 8 * 60 * 60 * 1000 }
}));

const ADMIN_EMAIL = 'etruslow@waynesboro.k12.va.us';
const GOOGLE_SHEET_ID = '1FhnS8B4GKz3vA3COT0RGqJpKz4AdDf28Tq-zfvDV8sc';

// Auth middleware
function requireAuth(req, res, next) {
  if (!req.session.user) return res.status(401).json({ error: 'Not authenticated' });
  next();
}

function requireAdmin(req, res, next) {
  if (!req.session.user || req.session.user.role !== 'admin') {
    return res.status(403).json({ error: 'Admin access required' });
  }
  next();
}

// Login
app.post('/api/login', (req, res) => {
  const { email } = req.body;
  if (!email) return res.status(400).json({ error: 'Email required' });
  const normalizedEmail = email.trim().toLowerCase();
  const role = normalizedEmail === ADMIN_EMAIL.toLowerCase() ? 'admin' : 'readonly';
  req.session.user = { email: normalizedEmail, role };
  res.json({ email: normalizedEmail, role });
});

app.post('/api/logout', (req, res) => {
  req.session.destroy();
  res.json({ ok: true });
});

app.get('/api/me', (req, res) => {
  if (!req.session.user) return res.status(401).json({ error: 'Not authenticated' });
  res.json(req.session.user);
});

// Fetch CSV from Google Sheets
function fetchGoogleSheet() {
  const url = `https://docs.google.com/spreadsheets/d/${GOOGLE_SHEET_ID}/export?format=csv`;
  return new Promise((resolve, reject) => {
    function doRequest(reqUrl, redirects) {
      if (redirects > 5) return reject(new Error('Too many redirects'));
      const mod = reqUrl.startsWith('https') ? https : require('http');
      mod.get(reqUrl, (resp) => {
        if (resp.statusCode >= 300 && resp.statusCode < 400 && resp.headers.location) {
          return doRequest(resp.headers.location, redirects + 1);
        }
        if (resp.statusCode !== 200) {
          return reject(new Error(`Google Sheets returned status ${resp.statusCode}. Make sure the sheet is shared as "Anyone with the link can view".`));
        }
        let data = '';
        resp.on('data', chunk => data += chunk);
        resp.on('end', () => resolve(data));
        resp.on('error', reject);
      }).on('error', reject);
    }
    doRequest(url, 0);
  });
}

// Sync from Google Sheet - admin only
app.post('/api/sync', requireAdmin, async (req, res) => {
  try {
    const csvData = await fetchGoogleSheet();
    const records = parse(csvData, {
      columns: true,
      skip_empty_lines: true,
      trim: true,
      relax_column_count: true
    });

    let inserted = 0, updated = 0, skipped = 0;
    for (const row of records) {
      const result = db.upsertSignup(row);
      if (result === 'inserted') inserted++;
      else if (result === 'updated') updated++;
      else skipped++;
    }

    res.json({ success: true, inserted, updated, skipped, total: records.length });
  } catch (err) {
    console.error('Google Sheet sync error:', err);
    res.status(500).json({ error: 'Failed to sync from Google Sheet: ' + err.message });
  }
});

// CSV upload - admin only (fallback / manual upload)
app.post('/api/upload', requireAdmin, upload.single('csv'), (req, res) => {
  if (!req.file) return res.status(400).json({ error: 'No file uploaded' });

  try {
    const content = req.file.buffer.toString('utf-8');
    const records = parse(content, {
      columns: true,
      skip_empty_lines: true,
      trim: true,
      relax_column_count: true
    });

    let inserted = 0, updated = 0, skipped = 0;
    for (const row of records) {
      const result = db.upsertSignup(row);
      if (result === 'inserted') inserted++;
      else if (result === 'updated') updated++;
      else skipped++;
    }

    res.json({ success: true, inserted, updated, skipped, total: records.length });
  } catch (err) {
    console.error('CSV parse error:', err);
    res.status(400).json({ error: 'Failed to parse CSV: ' + err.message });
  }
});

// Get signups with optional grade filter
app.get('/api/signups', requireAuth, (req, res) => {
  const { grade } = req.query;
  res.json(db.getSignups(grade || null));
});

// Get summary counts
app.get('/api/summary', requireAuth, (req, res) => {
  res.json(db.getSummary());
});

// Update student status - admin only
app.put('/api/signups/:id/status', requireAdmin, (req, res) => {
  const { status } = req.body;
  const validStatuses = ['none', 'in_building', 'late', 'cancel'];
  if (!validStatuses.includes(status)) {
    return res.status(400).json({ error: 'Invalid status' });
  }
  db.updateStatus(req.params.id, status);
  res.json({ ok: true });
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
  console.log(`PTC Dashboard running on http://localhost:${PORT}`);
});
