'use strict';

const express = require('express');
const { WebSocketServer } = require('ws');
const http = require('http');
const fs = require('fs');
const path = require('path');
const crypto = require('crypto');
const XLSX = require('xlsx');

const app = express();
const server = http.createServer(app);
const wss = new WebSocketServer({ server, path: '/ws' });

const DATA_DIR = path.join(__dirname, 'data');
const DATA_PATH = path.join(DATA_DIR, 'scs_data.json');

// ── OneDrive / SharePoint Excel source ───────────────
const DEFAULT_EXCEL_URL = 'https://1drv.ms/x/c/a5c8ae2f1213eda1/IQBsx93bcfE1TK8-08MrXkCZAbfvyfIgOBXqW-zf1oNGlTw?e=FA3lDr';
const CONFIG_PATH = path.join(DATA_DIR, 'config.json');
const POLL_INTERVAL = parseInt(process.env.POLL_INTERVAL, 10) || 5 * 60 * 1000; // 5 minutes

// Mutable — can be changed at runtime via admin page
let excelUrl = process.env.EXCEL_URL || DEFAULT_EXCEL_URL;
let lastDataHash = null;
let lastFetchStatus = { ok: true, time: null, message: 'Not fetched yet' };
let refreshInProgress = false;
let refreshPromise = null;

function startBackgroundRefresh() {
  if (refreshInProgress) return refreshPromise;
  refreshInProgress = true;
  refreshPromise = fetchExcelAndUpdate()
    .catch(err => {
      console.error(`[${new Date().toISOString()}] Background refresh error: ${err.message}`);
      lastFetchStatus = { ok: false, time: new Date().toISOString(), message: err.message };
    })
    .finally(() => {
      refreshInProgress = false;
      refreshPromise = null;
    });
  return refreshPromise;
}

// Load saved config (persists URL across restarts)
function loadConfig() {
  try {
    if (fs.existsSync(CONFIG_PATH)) {
      const cfg = JSON.parse(fs.readFileSync(CONFIG_PATH, 'utf8'));
      if (cfg.excelUrl) excelUrl = cfg.excelUrl;
    }
  } catch (_) { /* use default */ }
}
function saveConfig() {
  fs.writeFileSync(CONFIG_PATH, JSON.stringify({ excelUrl }, null, 2), 'utf8');
}

// ── OneDrive cookie-aware download ──────────────────
// OneDrive personal sharing links require a two-phase approach:
//   Phase 1: Follow the sharing link redirects, collecting cookies (FedAuth)
//            and extracting the sourcedoc UniqueId from the redirect URL.
//   Phase 2: Use download.aspx?UniqueId=... with the FedAuth cookie to get the file.

async function fetchOneDriveBuffer(shareUrl) {
  const UA = 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36';
  const cookies = {};
  let url = shareUrl;
  let baseOrigin = '';
  let basePath = '';
  let uniqueId = '';

  // Phase 1: Follow redirects, collect cookies & extract UniqueId
  for (let i = 0; i < 10; i++) {
    const cookieStr = Object.entries(cookies).map(([k, v]) => k + '=' + v).join('; ');
    const r = await fetch(url, {
      redirect: 'manual',
      headers: { 'User-Agent': UA, 'Cookie': cookieStr }
    });
    // Collect Set-Cookie headers
    const sc = r.headers.getSetCookie ? r.headers.getSetCookie() : [];
    sc.forEach(c => { const m = c.match(/^([^=]+)=([^;]*)/); if (m) cookies[m[1]] = m[2]; });

    const loc = r.headers.get('location');
    if (loc && r.status >= 300 && r.status < 400) {
      url = new URL(loc, url).href;
      // Extract base and UniqueId from redirect URL parameters
      try {
        const u = new URL(url);
        const personalMatch = u.pathname.match(/\/(personal\/[^/]+)/);
        if (personalMatch) {
          baseOrigin = u.origin;
          basePath = personalMatch[1];
        }
        const sd = u.searchParams.get('sourcedoc');
        if (sd) uniqueId = sd.replace(/[{}]/g, '');
      } catch (_) {}
    } else {
      // Reached the viewer page — also try extracting UniqueId from HTML
      if (!uniqueId) {
        const html = await r.text();
        const fileGetMatch = html.match(/"FileGetUrl":"[^"]*UniqueId=([a-f0-9-]+)/i);
        if (fileGetMatch) uniqueId = fileGetMatch[1];
      } else {
        // Consume the body to free resources
        await r.text();
      }
      break;
    }
  }

  if (!uniqueId || !baseOrigin || !basePath) {
    throw new Error('Could not extract download parameters from OneDrive sharing link');
  }

  // Phase 2: Download the file using UniqueId + cookies
  const dlUrl = `${baseOrigin}/${basePath}/_layouts/15/download.aspx?UniqueId=${uniqueId}`;
  const cookieStr = Object.entries(cookies).map(([k, v]) => k + '=' + v).join('; ');
  console.log(`[${new Date().toISOString()}] OneDrive download: UniqueId=${uniqueId.slice(0, 8)}… cookies=[${Object.keys(cookies).join(',')}]`);

  const r = await fetch(dlUrl, {
    redirect: 'follow',
    headers: { 'User-Agent': UA, 'Cookie': cookieStr }
  });

  if (!r.ok) throw new Error(`OneDrive download failed: HTTP ${r.status}`);

  const ct = r.headers.get('content-type') || '';
  const buf = Buffer.from(await r.arrayBuffer());

  // Verify we got an actual file, not HTML
  if (ct.includes('text/html') || (buf[0] !== 0x50 || buf[1] !== 0x4B)) {
    throw new Error('OneDrive returned an HTML page instead of a file. Ensure the sharing link is set to "Anyone with the link".');
  }
  return buf;
}

// Pick the best sheet from the workbook: prefer "Suppliers", then the first sheet with data rows
function pickDataSheet(workbook) {
  // Prefer a sheet named "Suppliers" (case-insensitive)
  const suppliersSheet = workbook.SheetNames.find(n => /^suppliers$/i.test(n));
  if (suppliersSheet) return suppliersSheet;

  // Otherwise find the first sheet with >0 data rows
  for (const name of workbook.SheetNames) {
    const data = XLSX.utils.sheet_to_json(workbook.Sheets[name], { defval: '' });
    if (data.length > 0) return name;
  }
  return workbook.SheetNames[0];
}

// Parse an Excel buffer into JSON rows from the correct sheet
function parseExcelBuffer(buffer) {
  const workbook = XLSX.read(buffer, { type: 'buffer', cellDates: true });
  const sheetName = pickDataSheet(workbook);
  const sheet = workbook.Sheets[sheetName];
  const jsonData = XLSX.utils.sheet_to_json(sheet, { defval: '' });

  if (!Array.isArray(jsonData) || jsonData.length === 0) {
    throw new Error(`Excel sheet "${sheetName}" has no data rows`);
  }

  // Normalize date fields to YYYY-MM-DD strings
  jsonData.forEach(row => {
    ['Last Assessment Date', 'Next Assessment Due Date'].forEach(key => {
      if (row[key] instanceof Date) {
        row[key] = row[key].toISOString().split('T')[0];
      }
    });
  });

  console.log(`[${new Date().toISOString()}] Parsed sheet "${sheetName}": ${jsonData.length} rows`);
  return jsonData;
}

// Fetch Excel from OneDrive, parse, and update local JSON
async function fetchExcelAndUpdate() {
  console.log(`[${new Date().toISOString()}] Fetching Excel from: ${excelUrl.slice(0, 80)}…`);

  let buffer;
  // OneDrive sharing links need the cookie-aware approach
  if (/1drv\.ms|onedrive\.live\.com|sharepoint\.com/i.test(excelUrl)) {
    buffer = await fetchOneDriveBuffer(excelUrl);
  } else {
    // Direct URL — simple fetch
    const r = await fetch(excelUrl, {
      redirect: 'follow',
      headers: { 'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64)' }
    });
    if (!r.ok) throw new Error(`HTTP ${r.status}: ${r.statusText}`);
    buffer = Buffer.from(await r.arrayBuffer());
  }

  const jsonData = parseExcelBuffer(buffer);

  // Check if data actually changed
  const newHash = crypto.createHash('md5').update(JSON.stringify(jsonData)).digest('hex');
  if (newHash === lastDataHash) {
    console.log(`[${new Date().toISOString()}] Excel data unchanged (${jsonData.length} rows, hash: ${newHash.slice(0, 8)})`);
    lastFetchStatus = { ok: true, time: new Date().toISOString(), message: `No changes (${jsonData.length} rows)` };
    return { changed: false, count: jsonData.length };
  }

  // Save to disk
  fs.writeFileSync(DATA_PATH, JSON.stringify(jsonData, null, 2), 'utf8');
  lastDataHash = newHash;

  // Broadcast to all WebSocket clients
  const msg = JSON.stringify({ type: 'data-updated', timestamp: Date.now(), count: jsonData.length });
  wss.clients.forEach(client => {
    if (client.readyState === 1) client.send(msg);
  });

  console.log(`[${new Date().toISOString()}] Excel data updated: ${jsonData.length} rows (hash: ${newHash.slice(0, 8)})`);
  lastFetchStatus = { ok: true, time: new Date().toISOString(), message: `Updated — ${jsonData.length} rows` };
  return { changed: true, count: jsonData.length };
}

// Wrapper that catches errors for polling
async function pollExcel() {
  try {
    await fetchExcelAndUpdate();
  } catch (err) {
    console.error(`[${new Date().toISOString()}] Excel poll error: ${err.message}`);
    lastFetchStatus = { ok: false, time: new Date().toISOString(), message: err.message };
  }
}

// Ensure data directory exists
if (!fs.existsSync(DATA_DIR)) {
  fs.mkdirSync(DATA_DIR, { recursive: true });
}

// Load saved config (URL etc.)
loadConfig();

// Seed with default data if missing
if (!fs.existsSync(DATA_PATH)) {
  const seed = path.join(__dirname, 'public', 'scs_data.json');
  if (fs.existsSync(seed)) {
    fs.copyFileSync(seed, DATA_PATH);
  } else {
    fs.writeFileSync(DATA_PATH, '[]', 'utf8');
  }
}

// --- Middleware ---
app.use(express.json({ limit: '10mb' }));
app.use(express.raw({ type: 'application/octet-stream', limit: '50mb' }));

// Serve scs_data.json from /data with Last-Modified header
app.get('/scs_data.json', (req, res) => {
  if (!fs.existsSync(DATA_PATH)) {
    return res.status(404).json({ error: 'scs_data.json not found' });
  }
  const stat = fs.statSync(DATA_PATH);
  res.set('Last-Modified', stat.mtime.toUTCString());
  res.set('Cache-Control', 'no-cache, no-store, must-revalidate');
  res.set('Content-Type', 'application/json; charset=utf-8');
  fs.createReadStream(DATA_PATH).pipe(res);
});

// Serve static files (HTML, CSS, JS)
app.use(express.static(path.join(__dirname, 'public')));

// Serve dashboard at root path
app.get('/', (req, res) => {
  res.sendFile(path.join(__dirname, 'public', 'scs_dashboard.html'));
});

// --- API: Manual refresh — fetch Excel NOW ---
app.post('/api/refresh', async (req, res) => {
  try {
    const result = await fetchExcelAndUpdate();
    res.json({ ok: true, changed: result.changed, count: result.count, timestamp: new Date().toISOString() });
  } catch (err) {
    res.status(502).json({ ok: false, error: err.message, timestamp: new Date().toISOString() });
  }
});

app.post('/api/refresh-async', async (req, res) => {
  if (refreshInProgress) {
    return res.status(202).json({ ok: true, status: 'already-running' });
  }
  startBackgroundRefresh();
  res.status(202).json({ ok: true, status: 'queued' });
});

// --- API: Refresh status (for the UI to show last poll result) ---
app.get('/api/refresh-status', (req, res) => {
  res.json({ ...lastFetchStatus, refreshInProgress });
});

// --- API: Get current config (admin page) ---
app.get('/api/config', (req, res) => {
  res.json({ excelUrl, pollInterval: POLL_INTERVAL, lastFetchStatus });
});

// --- API: Update Excel URL and trigger fetch ---
app.post('/api/config', async (req, res) => {
  const { excelUrl: newUrl } = req.body;
  if (!newUrl || typeof newUrl !== 'string' || !newUrl.startsWith('https://')) {
    return res.status(400).json({ ok: false, error: 'A valid HTTPS URL is required' });
  }
  excelUrl = newUrl.trim();
  saveConfig();
  console.log(`[${new Date().toISOString()}] Excel URL updated via admin page`);

  // Immediately fetch with the new URL
  try {
    lastDataHash = null; // force re-process even if same data
    const result = await fetchExcelAndUpdate();
    res.json({ ok: true, changed: result.changed, count: result.count, timestamp: new Date().toISOString() });
  } catch (err) {
    res.status(502).json({ ok: false, error: err.message, timestamp: new Date().toISOString() });
  }
});

// --- API: Upload Excel file directly (binary) ---
app.post('/api/upload-excel', (req, res) => {
  try {
    if (!req.body || req.body.length === 0) {
      return res.status(400).json({ ok: false, error: 'No file data received' });
    }

    const buffer = Buffer.from(req.body);
    const jsonData = parseExcelBuffer(buffer);

    // Save to disk
    fs.writeFileSync(DATA_PATH, JSON.stringify(jsonData, null, 2), 'utf8');
    lastDataHash = crypto.createHash('md5').update(JSON.stringify(jsonData)).digest('hex');

    // Broadcast to all WebSocket clients
    const msg = JSON.stringify({ type: 'data-updated', timestamp: Date.now(), count: jsonData.length });
    wss.clients.forEach(client => {
      if (client.readyState === 1) client.send(msg);
    });

    console.log(`[${new Date().toISOString()}] Excel uploaded via admin: ${jsonData.length} suppliers`);
    lastFetchStatus = { ok: true, time: new Date().toISOString(), message: `Uploaded — ${jsonData.length} suppliers` };
    res.json({ ok: true, count: jsonData.length, timestamp: new Date().toISOString() });
  } catch (err) {
    res.status(400).json({ ok: false, error: 'Failed to parse Excel: ' + err.message });
  }
});

// --- API: Upload new data (POST JSON array) ---
app.post('/api/data', (req, res) => {
  if (!Array.isArray(req.body)) {
    return res.status(400).json({ error: 'Request body must be a JSON array' });
  }
  fs.writeFileSync(DATA_PATH, JSON.stringify(req.body, null, 2), 'utf8');

  const msg = JSON.stringify({ type: 'data-updated', timestamp: Date.now(), count: req.body.length });
  wss.clients.forEach(client => {
    if (client.readyState === 1) client.send(msg);
  });

  console.log(`[${new Date().toISOString()}] Data updated: ${req.body.length} suppliers`);
  res.json({ ok: true, count: req.body.length, timestamp: new Date().toISOString() });
});

// --- API: Get current data (GET) ---
app.get('/api/data', (req, res) => {
  if (!fs.existsSync(DATA_PATH)) {
    return res.json([]);
  }
  const raw = fs.readFileSync(DATA_PATH, 'utf8');
  res.set('Cache-Control', 'no-cache');
  res.set('Content-Type', 'application/json; charset=utf-8');
  res.send(raw);
});

// --- Health check for OpenShift/K8s probes ---
app.get('/health', (req, res) => {
  res.json({ status: 'ok', uptime: process.uptime() });
});

app.get('/ready', (req, res) => {
  res.json({ status: 'ready' });
});

// --- WebSocket heartbeat to keep connections alive ---
const HEARTBEAT_INTERVAL = 30000;
const heartbeat = setInterval(() => {
  wss.clients.forEach(client => {
    if (client.readyState === 1) {
      client.send(JSON.stringify({ type: 'ping' }));
    }
  });
}, HEARTBEAT_INTERVAL);

wss.on('connection', (ws, req) => {
  console.log(`[${new Date().toISOString()}] WebSocket connected (clients: ${wss.clients.size})`);
  ws.on('close', () => {
    console.log(`[${new Date().toISOString()}] WebSocket disconnected (clients: ${wss.clients.size})`);
  });
});

wss.on('close', () => {
  clearInterval(heartbeat);
});

// --- Start server ---
const PORT = parseInt(process.env.PORT, 10) || 8080;
server.listen(PORT, '0.0.0.0', () => {
  console.log(`SCS Dashboard server running on http://0.0.0.0:${PORT}`);
  console.log(`  Dashboard:  http://localhost:${PORT}/scs_dashboard.html`);
  console.log(`  Excel src:  ${excelUrl.slice(0, 80)}…`);  console.log(`  Admin:      http://localhost:${PORT}/admin.html`);
  console.log(`  Poll every: ${POLL_INTERVAL / 1000}s`);
  console.log(`  Refresh:    POST http://localhost:${PORT}/api/refresh`);
  console.log(`  Health:     http://localhost:${PORT}/health`);
  console.log(`  WebSocket:  ws://localhost:${PORT}/ws`);

  // Initial fetch from Excel on startup
  pollExcel().then(() => {
    // Start polling every POLL_INTERVAL
    setInterval(pollExcel, POLL_INTERVAL);
    console.log(`[${new Date().toISOString()}] Excel polling started (every ${POLL_INTERVAL / 1000}s)`);
  });
});
