#!/usr/bin/env node
'use strict';

/*
 * registration_automation / script.js
 *
 * Reads rows (subscriber_id, IGM_version) from an Excel sheet, finds the organisation for each
 * subscriber_id in the environment_access table, calls the create-token API once per row, and
 * writes the resolved org_id, "Completed" or "Failed", and a message into that row.
 * A row that already has an org_id is used as is (no database lookup).
 *
 * HOW THE ORGANISATION IS FOUND (real mode)
 *   SELECT organisation_id, status FROM environment_access WHERE subscriber_id matches.
 *   - Matching ignores spaces and upper/lower case.
 *   - No stage filter: the environment is whichever database you point the script at
 *     (DB_HOST, DB_PORT, DB_USER, DB_PASSWORD, DB_DATABASE; optional DB_SSL=true|no-verify).
 *     Use --stage <n> to restrict to one stage_id (0 Sandbox, 1 Staging, 2 Pre-Prod, 3 Beta-Prod).
 *   - Only rows with status Active (1), Whitelisted (4) or Subscribed (7) count as registered.
 *   - Exactly one distinct organisation must match; zero or several => the row is Failed.
 *   - The database session is read-only. Nothing is ever written to the database by this script.
 *
 * WHAT create-token REALLY DOES
 *   POST /api/plan/create-token creates API credentials (client id, client secret, token) that
 *   link an organisation to an external grievance platform (IGM v1 = platform 2, IGM v2 = 3).
 *   It does NOT register a user or an organisation. The organisation must already exist.
 *
 * SAFETY (read this)
 *   - The default mode is MOCK. The script starts a fake create-token server on 127.0.0.1 and
 *     talks only to that. No real API, database or data is touched.
 *   - By default the result is written to a NEW file next to the input, so the original sheet is
 *     never changed. Mock results go to *.mock-result.xlsx, real results to *.result.xlsx.
 *   - REAL mode creates real credentials in a shared database. It needs all three of:
 *         --mode real --base-url https://...   AND   --yes-create-real-credentials
 *         AND the environment variable ALLOW_REAL_API=yes
 *     and still waits 5 seconds so you can press Ctrl+C.
 *   - Tokens, client ids and secrets are never printed and never written to the Excel file.
 *
 * LOCAL (MOCK) MODE vs HOSTED (REAL) MODE - what is called and what happens
 *   The mode is chosen only by TARGET_BASE_URL: empty = local mock, set = real.
 *   The same code runs in both; only the lookup and the create-token target differ.
 *   (Hosted = the Node server, `npm start` on a server, reading its settings from environment variables.)
 *
 *                      LOCAL / MOCK (TARGET_BASE_URL empty)   HOSTED / REAL (TARGET_BASE_URL set)
 *   Org lookup         fake SAMPLE_ACCESS table in this file   real SELECT on environment_access
 *                      (ids 70001-70011); no database          using DB_HOST/DB_PORT/DB_USER/...
 *   create-token       POST to a fake server this script       POST to ${TARGET_BASE_URL}/api/plan/create-token
 *                      starts on 127.0.0.1                     on the real server
 *   Credentials        fake (MOCK-TOKEN-<org> ...)             real client id / secret / token, stored
 *                                                              in the real database
 *   Safeguards         none needed                             ALLOW_REAL_API=yes, ACCESS_KEY (page),
 *                                                              --yes-create-real-credentials (CLI),
 *                                                              5 second Ctrl+C window (CLI)
 *   Dry run            n/a                                     looks up organisations only; sends no
 *                                                              create-token request
 *   Result file        *.mock-result.xlsx                      *.result.xlsx
 *   Static page only   static/index.html alone is always a browser demo (own fake data, no network);
 *   (no Node server)   it can never reach the real API or database.
 *
 * MOCK RULES (so you can see both outcomes without any real system)
 *   org_id divisible by 5  -> 400 "Organisation not found"   -> Failed
 *   org_id divisible by 7  -> 200 "Token already exists"     -> Completed
 *   org_id 999999          -> 429 plain text (rate limit)    -> run halts
 *   anything else          -> 200 token created              -> Completed
 *
 * MOCK SUBSCRIBER ID RULES (mock mode has no database; it fakes the environment_access lookup)
 *   Uses the SAMPLE_ACCESS table below, the same one the static page uses: ids 70001-70011.
 *   Any other id -> "not found" -> Failed.
 */

const fs = require('fs');
const path = require('path');
const http = require('http');
const ExcelJS = require('exceljs');

// Settings for this script live in registration_automation/.env (copy .env.example to .env).
require('./env');

/*
 * TARGET ENDPOINT
 * Leave empty to use the local mock server (safe default).
 * To run against another server (e.g. pre-production), put its base URL here, for example
 *     const TARGET_BASE_URL = 'https://preprod.mycompany.com';
 * or set the environment variable TARGET_BASE_URL, or pass --base-url on the command line.
 * The script appends /api/plan/create-token itself. Setting a URL switches the script to REAL mode,
 * which still needs --yes-create-real-credentials and ALLOW_REAL_API=yes.
 */
const TARGET_BASE_URL = process.env.TARGET_BASE_URL || '';

const DEFAULT_FILE = path.join(__dirname, 'registration_automation.xlsx');
const REQUEST_TIMEOUT_MS = 30000;
const STATUS_DONE = 'Completed';
const STATUS_FAIL = 'Failed';
// environment_access.status values that count as "registered".
const REGISTERED_STATUSES = [1, 4, 7];
const STATUS_NAMES = {
  0: 'Inactive', 1: 'Active', 2: 'Pending', 3: 'Rejected', 4: 'Whitelisted',
  5: 'Not Whitelisted', 6: 'Not subscribed', 7: 'Subscribed', 8: 'Whitelisted Old',
};
const DB_TIMEOUT_MS = 15000;
const HALT_KINDS = new Set(['rate-limit', 'timeout', 'network']);
const HALT_MESSAGES = {
  'rate-limit': 'Server rate limit reached (the API allows 20 requests per hour per IP).',
  timeout: 'A request timed out, so its outcome is unknown. Check that organisation before retrying.',
  network: 'Could not reach the server.',
};
const MOCK_RATE_LIMIT_ORG = 999999;
// Fake environment_access rows for mock mode (keep in sync with SAMPLE_ACCESS in static/index.html).
const SAMPLE_ACCESS = [
  { org: 6533, id: '70001', status: 1, stage: 2 },
  { org: 6533, id: '70001', status: 7, stage: 0 },
  { org: 7952, id: '70002', status: 7, stage: 2 },
  { org: 4900, id: '70003', status: 1, stage: 2 },
  { org: 2338, id: '70004', status: 4, stage: 2 },
  { org: 1504, id: '70005', status: 1, stage: 2 },
  { org: 9217, id: '70006', status: 1, stage: 2 },
  { org: 1001, id: '70007', status: 2, stage: 2 },
  { org: 1002, id: '70008', status: 3, stage: 2 },
  { org: 1003, id: '70009', status: 1, stage: 2 },
  { org: 1004, id: '70009', status: 1, stage: 2 },
  { org: 6129, id: '70010', status: 7, stage: 2 },
  { org: 999999, id: '70011', status: 1, stage: 2 },
];

const HELP = `
Usage: node script.js [options]

  --file <path>        .xlsx or .csv file to read (default: registration_automation.xlsx next to this script)
  --sheet <name>       Sheet name for .xlsx files (default: Registrations; .csv uses its only sheet)
  --out <path>         Where to write the result (default: <file>.mock-result.xlsx or <file>.result.xlsx)
  --dry-run            Only look up the organisations (database, read-only). Sends NO create-token
                       request and creates nothing. Status is written as "Dry run OK" or "Failed".
  --in-place           Write the statuses back into the input file itself (close it in Excel first)
  --max-rows <n>       Max API calls per run (default: 20, the server's per-hour limit)
  --stage <n>          Real mode only: only match environment_access rows with this stage_id
                       (0 Sandbox, 1 Staging, 2 Pre-Prod, 3 Beta-Prod). Default: any stage.
  --mode <mock|real>   mock (default) = local fake server, real = the actual API (see safety below)
  --base-url <url>     Real mode only: backend base URL (https://... or http://...), also set by TARGET_BASE_URL
  --yes-create-real-credentials
                       Real mode only: confirms you understand real credentials will be created
  --help               Show this help

Sheet columns: subscriber_id and IGM_version (or org_id instead of subscriber_id). The script
adds org_id, status and message columns if they are missing.
Real mode also requires the environment variable ALLOW_REAL_API=yes, and the database variables
DB_HOST, DB_PORT, DB_USER, DB_PASSWORD, DB_DATABASE when a row needs a subscriber_id lookup.
Rows already marked Completed are skipped; empty and Failed rows are (re)tried.
`;

function die(message, code = 2) {
  console.error(`\nERROR: ${message}\n`);
  process.exit(code);
}

function parseArgs(argv) {
  const o = {
    file: DEFAULT_FILE,
    sheet: 'Registrations',
    out: '',
    inPlace: false,
    maxRows: 20,
    mode: TARGET_BASE_URL ? 'real' : 'mock',
    baseUrl: TARGET_BASE_URL.replace(/\/+$/, ''),
    confirmReal: false,
    dryRun: false,
    stage: null,
    help: false,
  };
  let modeGiven = false;
  for (let i = 0; i < argv.length; i += 1) {
    const a = argv[i];
    const next = () => {
      if (i + 1 >= argv.length) die(`Missing value for ${a}`);
      i += 1;
      return argv[i];
    };
    switch (a) {
      case '--file': o.file = path.resolve(next()); break;
      case '--sheet': o.sheet = next(); break;
      case '--out': o.out = path.resolve(next()); break;
      case '--in-place': o.inPlace = true; break;
      case '--max-rows': o.maxRows = Number(next()); break;
      case '--stage': o.stage = Number(next()); break;
      case '--mode': o.mode = next(); modeGiven = true; break;
      case '--base-url': o.baseUrl = next().replace(/\/+$/, ''); if (!modeGiven) o.mode = 'real'; break;
      case '--yes-create-real-credentials': o.confirmReal = true; break;
      case '--dry-run': o.dryRun = true; break;
      case '--help': case '-h': o.help = true; break;
      default: die(`Unknown option: ${a}\n${HELP}`);
    }
  }
  if (!['mock', 'real'].includes(o.mode)) die('--mode must be "mock" or "real".');
  if (!Number.isInteger(o.maxRows) || o.maxRows < 1) die('--max-rows must be a positive whole number.');
  if (o.stage !== null && !Number.isInteger(o.stage)) die('--stage must be a whole number (stage_id).');
  return o;
}

/* ------------------------------------------------------------------ safety guards */

const DB_VARS = ['DB_HOST', 'DB_PORT', 'DB_USER', 'DB_PASSWORD', 'DB_DATABASE'];

function assertRealAllowed(o, pendingCount, needsDb) {
  const problems = [];
  if (needsDb) {
    const missing = DB_VARS.filter((v) => !process.env[v]);
    if (missing.length) problems.push(`database variables missing for the subscriber_id lookup: ${missing.join(', ')}`);
  }
  if (!o.dryRun) {
    if (!o.baseUrl) problems.push('TARGET_BASE_URL (or --base-url <https://...>) is required in real mode');
    else if (!/^https?:\/\//i.test(o.baseUrl)) problems.push('the server URL must start with https:// or http://');
    if (!o.confirmReal) problems.push('--yes-create-real-credentials is required');
    if (process.env.ALLOW_REAL_API !== 'yes') problems.push('environment variable ALLOW_REAL_API=yes is required');
  }
  if (problems.length) {
    die(`Real mode refused. No request was sent.\n  - ${problems.join('\n  - ')}\n\nRun without --mode real to use the safe local mock.`);
  }
  if (o.dryRun) {
    console.log(`\nDRY RUN: organisations are looked up (read-only) in ${process.env.DB_DATABASE} on ${process.env.DB_HOST}:${process.env.DB_PORT}. No create-token request is sent.\n`);
    return Promise.resolve();
  }
  console.log('\n!!! REAL MODE !!!');
  console.log(`!!! ${pendingCount} request(s) will be sent to ${o.baseUrl} and may create REAL credentials.`);
  if (needsDb) {
    console.log(`!!! Organisations are looked up (read-only) in ${process.env.DB_DATABASE} on ${process.env.DB_HOST}:${process.env.DB_PORT}`
      + `${o.stage !== null ? `, stage_id ${o.stage}` : ', any stage'}.`);
    console.log('!!! Make sure the database and the API above belong to the SAME environment.');
  }
  console.log('!!! Press Ctrl+C within 5 seconds to cancel.\n');
  return new Promise((resolve) => { setTimeout(resolve, 5000); });
}

/** Excel keeps a hidden "~$name.xlsx" file while a workbook is open. Writing then would fail or clash. */
function findExcelLockFile(filePath) {
  const dir = path.dirname(filePath);
  const base = path.basename(filePath);
  return [`~$${base}`, `~$${base.slice(2)}`]
    .map((n) => path.join(dir, n))
    .find((p) => fs.existsSync(p));
}

/* ------------------------------------------------------------------ mock server */

function startMockServer() {
  return new Promise((resolve) => {
    const server = http.createServer((req, res) => {
      let body = '';
      req.on('data', (c) => { body += c; });
      req.on('end', () => {
        const send = (code, type, out) => { res.writeHead(code, { 'Content-Type': type }); res.end(out); };
        const json = (code, obj) => send(code, 'application/json', JSON.stringify(obj));
        const error = (code, msg) => json(code, { status: 'error', code, error: { description: 'Validation Error', message: { err_msg: msg } } });

        if (req.method !== 'POST' || req.url !== '/api/plan/create-token') return error(404, 'Not found');
        let p;
        try { p = JSON.parse(body); } catch { return error(400, 'Invalid JSON'); }

        // Same input rules as the real backend schema (organisation_id and platform_id integers).
        if (!Number.isInteger(p.organisation_id)) return error(400, '"Organisation_id" must be an integer');
        if (!Number.isInteger(p.platform_id)) return error(400, '"Platform_id" must be an integer');
        if (![2, 3].includes(p.platform_id)) return error(400, 'Mock only accepts platform_id 2 or 3');

        const org = p.organisation_id;
        if (org === MOCK_RATE_LIMIT_ORG) return send(429, 'text/plain', 'Too many requests, please try again later.');
        if (org % 5 === 0) return error(400, 'Organisation not found');

        const data = {
          token: `MOCK-TOKEN-${org}`,
          client_id: `MOCK-CLIENT-${org}`,
          client_secret: `MOCK-SECRET-${org}`,
          organisation_id: org,
          organisation_name: `Mock Organisation ${org}`,
          platform_id: p.platform_id,
        };
        if (org % 7 === 0) data.message = 'Token already exists for this organisation and platform';
        return json(200, { status: 'success', data });
      });
    });
    server.listen(0, '127.0.0.1', () => {
      resolve({ server, baseUrl: `http://127.0.0.1:${server.address().port}` });
    });
  });
}

/* ------------------------------------------------------------------ API call */

function readableError(body) {
  const msg = body && (body.error && body.error.message !== undefined ? body.error.message : (body.message !== undefined ? body.message : body.error));
  if (!msg) return 'Something went wrong.';
  if (typeof msg === 'string') return msg;
  if (msg.err_msg) return msg.err_msg;
  return Object.values(msg).filter((v) => typeof v === 'string').join(', ') || 'Something went wrong.';
}

/** Always resolves to { ok, kind, message }. Never throws. Credentials in the reply are discarded on purpose. */
async function createToken(baseUrl, organisationId, platformId) {
  const controller = new AbortController();
  const timer = setTimeout(() => controller.abort(), REQUEST_TIMEOUT_MS);
  try {
    const res = await fetch(`${baseUrl}/api/plan/create-token`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ organisation_id: organisationId, platform_id: platformId }),
      signal: controller.signal,
    });
    const text = await res.text();
    let body = null;
    try { body = JSON.parse(text); } catch { /* not JSON */ }

    if (res.status === 429) return { ok: false, kind: 'rate-limit', message: 'Rate limited (HTTP 429)' };
    if (!body || typeof body !== 'object') return { ok: false, kind: 'bad-response', message: `Unexpected non-JSON response (HTTP ${res.status})` };
    if (body.status === 'success') {
      const existing = /already exists/i.test((body.data && body.data.message) || '');
      return { ok: true, kind: existing ? 'existing' : 'created', message: existing ? 'Token already existed' : 'Token created' };
    }
    return { ok: false, kind: 'error', message: readableError(body) };
  } catch (err) {
    if (err.name === 'AbortError') return { ok: false, kind: 'timeout', message: 'Request timed out' };
    return { ok: false, kind: 'network', message: `Could not reach ${baseUrl}` };
  } finally {
    clearTimeout(timer);
  }
}

/* ------------------------------------------------------------------ organisation lookup */

/** Same normalisation is used on the sheet value and (in SQL) on the stored value. */
function normalizeId(value) {
  return String(value === null || value === undefined ? '' : value).trim().toLowerCase();
}

/** Fake environment_access lookup for mock mode. Returns rows like the real query: [{organisation_id, status}]. */
function mockLookup(normId, stage) {
  return SAMPLE_ACCESS
    .filter((x) => x.id === normId && (stage === null || x.stage === stage))
    .map((x) => ({ organisation_id: x.org, status: x.status }));
}

/** Read-only connection to the database that holds environment_access. */
async function openDb() {
  const { Client } = require('pg');
  let ssl;
  if (process.env.DB_SSL === 'true') ssl = true;
  else if (process.env.DB_SSL === 'no-verify') ssl = { rejectUnauthorized: false };
  const client = new Client({
    host: process.env.DB_HOST,
    port: Number(process.env.DB_PORT) || 5432,
    user: process.env.DB_USER,
    password: process.env.DB_PASSWORD,
    database: process.env.DB_DATABASE,
    ssl,
    connectionTimeoutMillis: DB_TIMEOUT_MS,
    statement_timeout: DB_TIMEOUT_MS,
  });
  await client.connect();
  await client.query('SET default_transaction_read_only = on');
  return client;
}

/** Returns { lookup(normUrl) -> rows, close() }. Mock mode never touches a database. */
async function createLookup(o) {
  if (o.mode === 'mock') return { lookup: async (u) => mockLookup(u, o.stage), close: async () => {} };
  const client = await openDb();
  const sql = 'SELECT DISTINCT organisation_id, status FROM environment_access '
    + 'WHERE lower(btrim(subscriber_id::text)) = $1'
    + `${o.stage !== null ? ' AND stage_id = $2' : ''}`;
  return {
    lookup: async (u) => (await client.query(sql, o.stage !== null ? [u, o.stage] : [u])).rows,
    close: () => client.end(),
  };
}

/** Turns lookup rows into { orgId } or { error }. */
function pickOrganisation(rows) {
  if (!rows.length) return { error: 'Subscriber ID not found in environment_access' };
  const registered = new Set(rows.filter((x) => REGISTERED_STATUSES.includes(Number(x.status))).map((x) => Number(x.organisation_id)));
  if (registered.size === 1) return { orgId: [...registered][0] };
  if (registered.size > 1) return { error: `Ambiguous: subscriber ID matches several organisations (${[...registered].sort((a, b) => a - b).join(', ')})` };
  const seen = [...new Set(rows.map((x) => STATUS_NAMES[Number(x.status)] || `status ${x.status}`))].join(', ');
  return { error: `Subscriber ID found, but not registered (status: ${seen})` };
}

/* ------------------------------------------------------------------ row helpers */

function toPlatformId(value) {
  const s = String(value === null || value === undefined ? '' : value)
    .trim().toLowerCase().replace(/[\s_-]+/g, '');
  if (s === 'igmv1' || s === 'v1' || s === '2') return 2;
  if (s === 'igmv2' || s === 'v2' || s === '3') return 3;
  return null;
}

function toOrgId(value) {
  const s = String(value === null || value === undefined ? '' : value).trim();
  return /^[1-9]\d*$/.test(s) ? Number(s) : null;
}

/* ------------------------------------------------------------------ main */

async function main() {
  const o = parseArgs(process.argv.slice(2));
  if (o.help) { console.log(HELP); return; }

  if (path.basename(o.file).startsWith('~$')) die('That is an Excel lock file, not a workbook. Use the real .xlsx or .csv file.');
  if (!fs.existsSync(o.file)) die(`File not found: ${o.file}`);
  if (o.mode === 'mock' && o.baseUrl) die('--base-url cannot be used in mock mode. Mock mode only talks to its own local server.');
  if (o.maxRows > 20) console.warn('Warning: the real API allows only 20 requests per hour per IP; a higher --max-rows will probably hit the rate limit.');

  const isCsv = /\.csv$/i.test(o.file);
  const outPath = o.inPlace
    ? o.file
    : (o.out || o.file.replace(/\.(xlsx|csv)$/i, '') + (o.mode === 'mock' ? '.mock-result' : '.result') + (isCsv ? '.csv' : '.xlsx'));

  const lock = findExcelLockFile(outPath);
  if (lock) {
    die(`"${path.basename(outPath)}" looks open in Excel (found ${path.basename(lock)}).\nClose it in Excel and run again, or choose another output with --out.`);
  }

  const wb = new ExcelJS.Workbook();
  // For CSV, keep every value as text (so ids with leading zeros stay as they are).
  if (isCsv) await wb.csv.readFile(o.file, { map: (v) => v }); else await wb.xlsx.readFile(o.file);
  const ws = isCsv ? wb.worksheets[0] : wb.getWorksheet(o.sheet);
  if (!ws) die(`Sheet "${o.sheet}" not found. Sheets in this file: ${wb.worksheets.map((w) => w.name).join(', ')}`);

  // Find columns by header name (case-insensitive).
  const cols = {};
  let lastCol = 0;
  ws.getRow(1).eachCell({ includeEmpty: false }, (cell, colNumber) => {
    cols[cell.text.trim().replace(/^\uFEFF/, '').toLowerCase()] = colNumber;
    lastCol = Math.max(lastCol, colNumber);
  });
  if (!cols.igm_version || (!cols.subscriber_id && !cols.org_id)) {
    die(`Row 1 must contain the headers subscriber_id (or org_id) and IGM_version. Found: ${Object.keys(cols).join(', ') || '(nothing)'}`);
  }
  // Add the output columns that are missing (org_id, status, message), copying the header style.
  const styleFrom = lastCol;
  [['org_id', 12], ['status', 18], ['message', 60]].forEach(([name, width]) => {
    if (cols[name]) return;
    lastCol += 1;
    cols[name] = lastCol;
    const header = ws.getCell(1, lastCol);
    header.value = name;
    header.style = { ...ws.getCell(1, styleFrom).style };
    ws.getColumn(lastCol).width = width;
  });

  // Collect the rows to work on.
  const rows = [];
  for (let r = 2; r <= ws.rowCount; r += 1) {
    const idText = cols.subscriber_id ? ws.getCell(r, cols.subscriber_id).text.trim() : '';
    const orgText = ws.getCell(r, cols.org_id).text.trim();
    const verText = ws.getCell(r, cols.igm_version).text.trim();
    const statusText = ws.getCell(r, cols.status).text.trim();
    if (!idText && !orgText && !verText && !statusText) continue; // blank row
    rows.push({ r, idText, orgText, verText, statusText });
  }
  const pending = rows.filter((x) => x.statusText.toLowerCase() !== STATUS_DONE.toLowerCase());
  const alreadyDone = rows.length - pending.length;
  // Rows without an org_id but with a subscriber_id need the environment_access lookup.
  const needsDb = o.mode === 'real' && pending.some((x) => !x.orgText && x.idText);

  console.log('============================================================');
  console.log(o.mode === 'mock'
    ? ' MODE: MOCK  (local fake server - no real API, database or data)'
    : ` MODE: REAL  (actual API)${o.dryRun ? '  - DRY RUN: nothing will be created' : ''}`);
  console.log(` Input : ${o.file}`);
  console.log(` Output: ${outPath}${o.inPlace ? '  (in place)' : ''}`);
  console.log(` Rows  : ${rows.length} with data, ${alreadyDone} already Completed (skipped), ${pending.length} to process`);
  console.log('============================================================');

  let baseUrl;
  let mock = null;
  if (o.mode === 'mock') {
    if (!o.dryRun) {
      mock = await startMockServer();
      baseUrl = mock.baseUrl;
      if (!/^http:\/\/127\.0\.0\.1:\d+$/.test(baseUrl)) die('Internal safety check failed: mock URL is not local.');
    }
  } else {
    const toCall = Math.min(pending.length, o.maxRows);
    await assertRealAllowed(o, toCall, needsDb);
    baseUrl = o.baseUrl;
  }

  // Database access is opened only when a row actually needs the subscriber_id lookup.
  let db = null;
  const cache = new Map();
  const lookupOrg = async (url) => {
    const key = normalizeId(url);
    if (!cache.has(key)) {
      if (!db) db = await createLookup(o);
      cache.set(key, pickOrganisation(await db.lookup(key)));
    }
    return cache.get(key);
  };

  const setCell = (r, col, value) => {
    const cell = ws.getCell(r, col);
    cell.value = value;
    if (!cell.font || !cell.font.name) cell.style = { ...ws.getCell(r, cols.igm_version).style };
  };
  const setStatus = (r, value, message) => {
    setCell(r, cols.status, value);
    setCell(r, cols.message, message);
  };

  const counts = { completed: 0, failed: 0, notAttempted: 0, dryRun: 0 };
  let attempted = 0;
  let halted = '';

  for (let i = 0; i < pending.length; i += 1) {
    const row = pending[i];
    const who = row.orgText || !row.idText ? `org_id ${(row.orgText || '-').padEnd(8)}` : `subscriber_id ${row.idText}`;
    const label = `Row ${String(row.r).padStart(3)} | ${who} | ${(row.verText || '-').padEnd(6)}`;
    const fail = (reason) => {
      setStatus(row.r, STATUS_FAIL, reason);
      counts.failed += 1;
      console.log(`${label} -> ${STATUS_FAIL}  (${reason}, no request sent)`);
    };

    let orgId = toOrgId(row.orgText);
    const platformId = toPlatformId(row.verText);
    let fromUrl = false;
    if (!row.orgText && row.idText) {
      let found;
      try {
        // eslint-disable-next-line no-await-in-loop
        found = await lookupOrg(row.idText);
      } catch (err) {
        halted = `Could not read environment_access (${err && err.message ? err.message : err}). Nothing was changed; check the database settings and run again.`;
        counts.notAttempted += pending.length - i;
        break;
      }
      if (found.error) { fail(found.error); continue; }
      orgId = found.orgId;
      fromUrl = true;
    }
    if (orgId === null || platformId === null) {
      const noOrgReason = !row.orgText && !row.idText && cols.subscriber_id ? 'subscriber_id is empty' : 'invalid org_id';
      fail(orgId === null ? noOrgReason : 'invalid IGM_version (use IGMv1 or IGMv2)');
      continue;
    }
    if (o.dryRun) {
      // Look-up only: show what would happen, create nothing. "Dry run OK" is not "Completed", so a real run still processes the row.
      if (fromUrl) setCell(row.r, cols.org_id, orgId);
      setStatus(row.r, 'Dry run OK', `Would create token for org ${orgId}, platform ${platformId}`);
      counts.dryRun += 1;
      console.log(`${label} -> Dry run OK  (org_id ${orgId}, platform ${platformId}; nothing created)`);
      continue;
    }
    if (attempted >= o.maxRows) {
      counts.notAttempted += pending.length - i;
      console.log(`\nReached --max-rows (${o.maxRows}); ${pending.length - i} row(s) left untouched for the next run.`);
      break;
    }

    attempted += 1;
    if (fromUrl) setCell(row.r, cols.org_id, orgId);
    // eslint-disable-next-line no-await-in-loop
    const result = await createToken(baseUrl, orgId, platformId);
    setStatus(row.r, result.ok ? STATUS_DONE : STATUS_FAIL, result.message);
    if (result.ok) counts.completed += 1; else counts.failed += 1;
    console.log(`${label} -> ${result.ok ? STATUS_DONE : STATUS_FAIL}  (${fromUrl ? `org_id ${orgId}, ` : ''}platform ${platformId}: ${result.message})`);

    if (HALT_KINDS.has(result.kind)) {
      halted = HALT_MESSAGES[result.kind];
      counts.notAttempted += pending.length - i - 1;
      break;
    }
  }

  if (db) await db.close();
  if (mock) { mock.server.closeAllConnections(); mock.server.close(); }

  // Write atomically: temp file first, then rename over the target.
  const tmp = `${outPath}.tmp-${process.pid}`;
  if (isCsv) await wb.csv.writeFile(tmp); else await wb.xlsx.writeFile(tmp);
  fs.renameSync(tmp, outPath);

  console.log('\n------------------------------------------------------------');
  console.log(` Completed: ${counts.completed}   Failed: ${counts.failed}   Left untouched: ${counts.notAttempted}   Skipped (already Completed): ${alreadyDone}`);
  if (o.dryRun) console.log(` Dry run OK: ${counts.dryRun}  (nothing was created)`);
  if (halted) console.log(` RUN STOPPED EARLY: ${halted}\n Rows after the stop were left empty so the next run picks them up.`);
  console.log(` Saved: ${outPath}`);
  console.log('------------------------------------------------------------');
  if (halted) process.exitCode = 1;
}

main().catch((err) => {
  console.error(`\nUnexpected error: ${err && err.message ? err.message : err}\n`);
  process.exit(1);
});
