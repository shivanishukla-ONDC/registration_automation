#!/usr/bin/env node
'use strict';

/*
 * server.js - serves the page (static/index.html) and runs script.js for each uploaded CSV.
 * Everything is configured with environment variables (or the .env file); nothing else needs changing.
 *
 *   TARGET_BASE_URL  empty -> MOCK mode (fake server + fake lookup, nothing real is touched)
 *                    set   -> REAL mode: create-token is called on that server
 *   DB_HOST, DB_PORT, DB_USER, DB_PASSWORD, DB_DATABASE (DB_SSL)   database with environment_access
 *   ALLOW_REAL_API=yes   required for real mode (creates real credentials)
 *   ACCESS_KEY       key users type into the page; REQUIRED in real mode
 *   PORT (3000), HOST (127.0.0.1; use 0.0.0.0 so other machines can open the page)
 *   MAX_ROWS         rows per file (default 20 in real mode, 100 in mock mode)
 *   LOOKUP_STAGE     optional stage_id filter for the environment_access lookup
 */

require('./env');
const fs = require('fs');
const os = require('os');
const path = require('path');
const http = require('http');
const crypto = require('crypto');
const { spawn } = require('child_process');
const ExcelJS = require('exceljs');

const PORT = Number(process.env.PORT) || 3000;
const HOST = process.env.HOST || '127.0.0.1';
const ACCESS_KEY = process.env.ACCESS_KEY || '';
const TARGET = (process.env.TARGET_BASE_URL || '').trim();
const REAL = Boolean(TARGET);
const MAX_ROWS = Number(process.env.MAX_ROWS) || (REAL ? 20 : 100);
const MAX_UPLOAD_BYTES = 2 * 1024 * 1024;
const RUN_TIMEOUT_MS = 15 * 60 * 1000;
const PAGE = path.join(__dirname, 'static', 'index.html');

let busy = false; // one run at a time: the API rate limit is shared and runs must not overlap

function sendJson(res, code, obj) {
  res.writeHead(code, { 'Content-Type': 'application/json; charset=utf-8', 'Cache-Control': 'no-store' });
  res.end(JSON.stringify(obj));
}

function isAuthorized(req) {
  if (!ACCESS_KEY) return true;
  const given = Buffer.from(String(req.headers['x-access-key'] || ''));
  const want = Buffer.from(ACCESS_KEY);
  return given.length === want.length && crypto.timingSafeEqual(given, want);
}

function readBody(req) {
  return new Promise((resolve, reject) => {
    const chunks = [];
    let size = 0;
    req.on('data', (c) => {
      size += c.length;
      if (size > MAX_UPLOAD_BYTES) { reject(Object.assign(new Error('File is too large (max 2 MB).'), { status: 413 })); req.destroy(); return; }
      chunks.push(c);
    });
    req.on('end', () => resolve(Buffer.concat(chunks).toString('utf8')));
    req.on('error', reject);
  });
}

function hostOf(url) {
  try { return new URL(url).host; } catch { return ''; }
}

function runScript(args) {
  return new Promise((resolve) => {
    const child = spawn(process.execPath, [path.join(__dirname, 'script.js'), ...args], { env: process.env });
    let out = '';
    const timer = setTimeout(() => child.kill('SIGKILL'), RUN_TIMEOUT_MS);
    child.stdout.on('data', (d) => { out += d; });
    child.stderr.on('data', (d) => { out += d; });
    child.on('close', () => { clearTimeout(timer); resolve(out); });
  });
}

async function readResultRows(file) {
  const wb = new ExcelJS.Workbook();
  await wb.csv.readFile(file, { map: (v) => v });
  const ws = wb.worksheets[0];
  const cols = {};
  ws.getRow(1).eachCell({ includeEmpty: false }, (cell, n) => { cols[cell.text.trim().replace(/^﻿/, '').toLowerCase()] = n; });
  const get = (r, name) => (cols[name] ? ws.getCell(r, cols[name]).text.trim() : '');
  const rows = [];
  for (let r = 2; r <= ws.rowCount; r += 1) {
    const row = { row: r, subscriber_id: get(r, 'subscriber_id'), igm_version: get(r, 'igm_version'), org_id: get(r, 'org_id'), status: get(r, 'status'), message: get(r, 'message') };
    if (row.subscriber_id || row.igm_version || row.org_id || row.status) rows.push(row);
  }
  return rows;
}

async function handleRun(req, res, url) {
  if (REAL && !ACCESS_KEY) return sendJson(res, 503, { error: 'Real mode is disabled: the server has no ACCESS_KEY set.' });
  if (busy) return sendJson(res, 429, { error: 'Another run is in progress. Please wait for it to finish and try again.' });

  const dryRun = url.searchParams.get('dryRun') === 'yes';
  if (REAL && !dryRun && url.searchParams.get('confirm') !== 'yes') {
    return sendJson(res, 400, { error: 'Please tick the box confirming that real credentials will be created.' });
  }
  const stageParam = url.searchParams.get('stage') || process.env.LOOKUP_STAGE || '';
  if (stageParam !== '' && !/^\d+$/.test(stageParam)) return sendJson(res, 400, { error: 'Stage must be a whole number.' });

  busy = true;
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'regauto-'));
  try {
    const upload = await readBody(req);
    if (!upload.trim() || upload.includes('\u0000')) return sendJson(res, 400, { error: 'That does not look like a CSV file.' });
    const inFile = path.join(dir, 'input.csv');
    const outFile = path.join(dir, 'result.csv');
    fs.writeFileSync(inFile, upload);

    const args = ['--file', inFile, '--out', outFile, '--max-rows', String(MAX_ROWS)];
    if (stageParam !== '') args.push('--stage', stageParam);
    if (dryRun) args.push('--dry-run');
    else if (REAL) args.push('--yes-create-real-credentials');
    const log = await runScript(args);

    if (!fs.existsSync(outFile)) {
      const m = log.match(/ERROR: ([\s\S]*?)\n\s*\n/) || log.match(/Unexpected error: (.*)/);
      return sendJson(res, 400, { error: m ? m[1].trim() : 'The run failed before any row was processed.' });
    }
    const num = (re) => Number((log.match(re) || [])[1] || 0);
    return sendJson(res, 200, {
      mode: REAL ? 'real' : 'mock',
      dryRun,
      summary: {
        completed: num(/Completed: (\d+)/),
        failed: num(/Failed: (\d+)/),
        notProcessed: num(/Left untouched: (\d+)/),
        skipped: num(/Skipped \(already Completed\): (\d+)/),
        dryRun: num(/Dry run OK: (\d+)/),
      },
      halted: (log.match(/RUN STOPPED EARLY: (.*)/) || [])[1] || '',
      rows: await readResultRows(outFile),
      csv: fs.readFileSync(outFile, 'utf8'),
    });
  } catch (err) {
    return sendJson(res, err.status || 500, { error: err.message || 'Unexpected error.' });
  } finally {
    busy = false;
    fs.rmSync(dir, { recursive: true, force: true });
  }
}

const server = http.createServer(async (req, res) => {
  const url = new URL(req.url, 'http://localhost');
  try {
    if (req.method === 'GET' && url.pathname === '/') {
      res.writeHead(200, { 'Content-Type': 'text/html; charset=utf-8', 'Cache-Control': 'no-store', 'X-Content-Type-Options': 'nosniff' });
      return res.end(fs.readFileSync(PAGE));
    }
    if (req.method === 'GET' && url.pathname === '/api/info') {
      return sendJson(res, 200, {
        mode: REAL ? 'real' : 'mock',
        target: REAL ? hostOf(TARGET) : 'local mock (nothing real is touched)',
        maxRows: MAX_ROWS,
        needsKey: Boolean(ACCESS_KEY),
        disabled: REAL && !ACCESS_KEY,
      });
    }
    if (req.method === 'POST' && url.pathname === '/api/run') {
      if (!isAuthorized(req)) return sendJson(res, 401, { error: 'Wrong or missing access key.' });
      return await handleRun(req, res, url);
    }
    res.writeHead(404);
    return res.end('Not found');
  } catch (err) {
    return sendJson(res, 500, { error: 'Unexpected error.' });
  }
});

server.listen(PORT, HOST, () => {
  console.log(`Registration automation: http://${HOST === '0.0.0.0' ? 'localhost' : HOST}:${PORT}`);
  console.log(REAL ? `MODE: REAL -> ${hostOf(TARGET)}${ACCESS_KEY ? '' : '  (disabled: set ACCESS_KEY)'}` : 'MODE: MOCK (local fake server, nothing real is touched)');
  if (HOST === '0.0.0.0' && !ACCESS_KEY) console.log('Warning: listening on all interfaces without an ACCESS_KEY.');
});
