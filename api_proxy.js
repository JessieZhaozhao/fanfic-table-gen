/**
 * Local API Proxy - bypass browser CORS
 *
 * Usage: node api_proxy.js
 * Proxy:  http://localhost:8787/proxy
 *
 * Frontend request format:
 *   POST http://localhost:8787/proxy
 *   Header: X-Target-URL: https://api.deepseek.com/v1/chat/completions
 *   Header: X-API-Key: sk-xxx
 *   Body: { model, messages, ... }
 */

const http = require('http');
const https = require('https');

const PORT = 8787;

const server = http.createServer((req, res) => {
  // CORS preflight
  if (req.method === 'OPTIONS') {
    setCORS(res);
    res.writeHead(204);
    res.end();
    return;
  }

  // Only allow /proxy
  if (req.url !== '/proxy' || req.method !== 'POST') {
    res.writeHead(404, { 'Content-Type': 'text/plain' });
    res.end('Not found. Use POST /proxy');
    return;
  }

  const targetURL = req.headers['x-target-url'];
  const apiKey = req.headers['x-api-key'];

  if (!targetURL) {
    res.writeHead(400, { 'Content-Type': 'application/json' });
    res.end(JSON.stringify({ error: 'Missing X-Target-URL header' }));
    return;
  }

  // Read request body
  let body = '';
  req.on('data', chunk => { body += chunk; });
  req.on('end', () => {
    let parsed;
    try {
      parsed = new URL(targetURL);
    } catch (e) {
      res.writeHead(400, { 'Content-Type': 'application/json' });
      res.end(JSON.stringify({ error: 'Invalid target URL: ' + targetURL }));
      return;
    }

    const options = {
      hostname: parsed.hostname,
      port: parsed.port || (parsed.protocol === 'https:' ? 443 : 80),
      path: parsed.pathname + parsed.search,
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'Authorization': 'Bearer ' + (apiKey || ''),
        'Content-Length': Buffer.byteLength(body),
      },
    };

    const lib = parsed.protocol === 'https:' ? https : http;
    const proxyReq = lib.request(options, (proxyRes) => {
      setCORS(res);
      res.writeHead(proxyRes.statusCode, proxyRes.headers);
      proxyRes.pipe(res);
    });

    proxyReq.on('error', (e) => {
      res.writeHead(502, { 'Content-Type': 'application/json' });
      res.end(JSON.stringify({ error: 'Proxy error: ' + e.message }));
    });

    proxyReq.write(body);
    proxyReq.end();
  });
});

function setCORS(res) {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type, X-Target-URL, X-API-Key');
}

server.listen(PORT, () => {
  console.log('\n  API proxy started -> http://localhost:' + PORT + '/proxy');
  console.log('  Check "local proxy" in HTML tool to use\n');
});
