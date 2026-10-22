#!/usr/bin/env node
'use strict';

/*
 * server.js - serves the static page (static/index.html) and nothing else. No database, no API, no extra packages.
 * It only exists so hosts that need a start command (e.g. a Render "Web Service": `npm start`) can serve the page.
 * A "Static Site" host does not need it: just publish the `static` folder.
 *
 *   PORT  port to listen on (default 5500; hosts such as Render set it)
 *   HOST  address to listen on (default 0.0.0.0 when PORT is set by a host, otherwise 127.0.0.1)
 */

const fs = require('fs');
const path = require('path');
const http = require('http');

const PORT = Number(process.env.PORT) || 5500;
const HOST = process.env.HOST || (process.env.PORT ? '0.0.0.0' : '127.0.0.1');
const ROOT = path.join(__dirname, 'static');
const TYPES = {
  '.html': 'text/html; charset=utf-8',
  '.js': 'text/javascript; charset=utf-8',
  '.css': 'text/css; charset=utf-8',
  '.csv': 'text/csv; charset=utf-8',
  '.json': 'application/json; charset=utf-8',
  '.ico': 'image/x-icon',
};

const server = http.createServer((req, res) => {
  if (req.method !== 'GET' && req.method !== 'HEAD') {
    res.writeHead(405, { Allow: 'GET, HEAD' });
    return res.end();
  }
  let pathname;
  try { pathname = decodeURIComponent(new URL(req.url, 'http://localhost').pathname); } catch { pathname = ''; }
  if (pathname === '/' || pathname === '') pathname = '/index.html';

  // Only files inside static/ are served (no ../ tricks).
  const file = path.join(ROOT, path.normalize(pathname));
  if (!file.startsWith(ROOT + path.sep) || !fs.existsSync(file) || !fs.statSync(file).isFile()) {
    res.writeHead(404, { 'Content-Type': 'text/plain; charset=utf-8' });
    return res.end('Not found');
  }
  res.writeHead(200, {
    'Content-Type': TYPES[path.extname(file).toLowerCase()] || 'application/octet-stream',
    'Cache-Control': 'no-store',
    'X-Content-Type-Options': 'nosniff',
  });
  return req.method === 'HEAD' ? res.end() : fs.createReadStream(file).pipe(res);
});

server.listen(PORT, HOST, () => {
  console.log(`Registration automation page: http://${HOST === '0.0.0.0' ? 'localhost' : HOST}:${PORT}`);
});
