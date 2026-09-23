import express from 'express';
import path from 'path';
import { fileURLToPath } from 'url';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const app = express();
const PORT = 3000;
const HOST = '0.0.0.0';

const APPS_SCRIPT_URL = process.env.APPS_SCRIPT_URL || 'https://script.google.com/macros/s/AKfycbxOKN15Szd7A3OGEM1-Xg_Jnwhg4QUibhQ4zAaFHuQBmJiUhtoyeAHFzUtEwdDrUtEFeg/exec';

app.use(express.json({ limit: '50mb' }));
app.use(express.urlencoded({ extended: true, limit: '50mb' }));

// Proxy endpoint to seamlessly bridge requests to Google Apps Script backend
app.post('/api/appscript', async (req, res) => {
  try {
    const upstreamRes = await fetch(APPS_SCRIPT_URL, {
      method: 'POST',
      headers: {
        'Content-Type': 'text/plain;charset=utf-8'
      },
      body: JSON.stringify(req.body || {}),
      redirect: 'follow'
    });

    const text = await upstreamRes.text();
    try {
      const json = JSON.parse(text);
      res.json(json);
    } catch {
      res.type('application/json').send(text);
    }
  } catch (err) {
    console.error('Apps Script Proxy Error:', err);
    res.status(502).json({
      status: 'error',
      message: 'Proxy error connecting to Google Apps Script: ' + err.message
    });
  }
});

app.get('/api/appscript', async (req, res) => {
  try {
    const upstreamRes = await fetch(APPS_SCRIPT_URL, {
      method: 'GET',
      redirect: 'follow'
    });
    const text = await upstreamRes.text();
    try {
      const json = JSON.parse(text);
      res.json(json);
    } catch {
      res.type('application/json').send(text);
    }
  } catch (err) {
    console.error('Apps Script GET Proxy Error:', err);
    res.status(502).json({
      status: 'error',
      message: 'Proxy error connecting to Google Apps Script: ' + err.message
    });
  }
});

// Serve static assets from project root
app.use(express.static(__dirname));

// Fallback all routes to index.html for SPA behavior
app.get('*', (req, res) => {
  res.sendFile(path.join(__dirname, 'index.html'));
});

app.listen(PORT, HOST, () => {
  console.log(`Server is running on http://${HOST}:${PORT}`);
});
