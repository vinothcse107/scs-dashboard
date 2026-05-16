'use strict';

const express = require('express');
const { WebSocketServer } = require('ws');
const http = require('http');
const fs = require('fs');
const path = require('path');

const app = express();
const server = http.createServer(app);
const wss = new WebSocketServer({ server, path: '/ws' });

const DATA_DIR = path.join(__dirname, 'data');
const DATA_PATH = path.join(DATA_DIR, 'scs_data.json');

// Ensure data directory exists
if (!fs.existsSync(DATA_DIR)) {
  fs.mkdirSync(DATA_DIR, { recursive: true });
}

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

// --- API: Upload new data (POST JSON array) ---
app.post('/api/data', (req, res) => {
  if (!Array.isArray(req.body)) {
    return res.status(400).json({ error: 'Request body must be a JSON array' });
  }
  fs.writeFileSync(DATA_PATH, JSON.stringify(req.body, null, 2), 'utf8');

  // Broadcast to all connected WebSocket clients
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
  console.log(`  Dashboard: http://localhost:${PORT}/scs_dashboard.html`);
  console.log(`  API:       POST http://localhost:${PORT}/api/data`);
  console.log(`  Health:    http://localhost:${PORT}/health`);
  console.log(`  WebSocket: ws://localhost:${PORT}/ws`);
});
