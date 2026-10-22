#!/usr/bin/env node
'use strict';

/*
 * registration_automation / script.js
 *
 * Reads a .csv or .xlsx file row by row and calls the onboarding API once per row:
 *
 *   POST <UAT_BASE_URL or PROD_BASE_URL>/api/plan/create-token-by-subscriber   { subscriber_id, platform_id }
 *
 * The API finds the organisation for the subscriber id and creates the token. This script does everything
 * else: reads the file, checks each row, turns IGMv1 / IGMv2 into platform_id 2 / 3, skips finished rows,
 * stops cleanly on a rate limit or network problem, and writes org_id, status, message and environment
 * back to the file. It has no database access.
 *
 * ENVIRONMENTS
 *   --env uat | prod picks the API address from UAT_BASE_URL / PROD_BASE_URL (in .env). Without --env the
 *   script runs in MOCK mode against a fake API on 127.0.0.1 and nothing real is touched.
 *
 * SAFETY
 *   - A real run creates real credentials. It needs --yes-create-real-credentials AND ALLOW_REAL_API=yes,
 *     and waits before sending (5 seconds for UAT, 10 for PROD) so you can press Ctrl+C.
 *   - Credentials in the API reply are discarded: never printed, never written to the file.
 *   - The result goes to a NEW file next to the input (*.mock-result.*, *.uat-result.* or *.prod-result.*)
 *     unless --in-place.
 *   - A row completed in UAT is not skipped in PROD (and the other way round): each row records its environment.
 */

const fs = require('fs');
const path = require('path');
const http = require('http');
const ExcelJS = require('exceljs');

// Settings live in registration_automation/.env (copy .env.example to .env).
require('./env');

const ROUTE = '/api/plan/create-token-by-subscriber';
const DEFAULT_FILE = path.join(__dirname, 'registration_automation.xlsx');
const REQUEST_TIMEOUT_MS = 30000;
const STATUS_DONE = 'Completed';
const STATUS_FAIL = 'Failed';
const MOCK_RATE_LIMIT_ID = 'ratelimit-test';
const ENV_URLS = {
  uat: (process.env.UAT_BASE_URL || '').trim().replace(/\/+$/, ''),
  prod: (process.env.PROD_BASE_URL || '').trim().replace(/\/+$/, ''),
};

const HELP = `
Usage: node script.js [options]

  --file <path>        .csv or .xlsx file to read (default: registration_automation.xlsx next to this script)
  --sheet <name>       Sheet name for .xlsx files (default: Registrations; .csv uses its only sheet)
  --out <path>         Where to write the result (default: next to the input, *.<env>-result.<ext>)
  --in-place           Write the results into the input file itself (close it in Excel first)
  --max-rows <n>       Stop after n API calls (default: no limit)
  --env <uat|prod>     Real run against UAT_BASE_URL or PROD_BASE_URL from .env. Without it: safe local mock.
  --yes-create-real-credentials
                       Real run: confirms that real credentials will be created
  --help               Show this help

File columns: subscriber_id and IGM_version (IGMv1 or IGMv2). org_id, status, message and environment are filled
in by the script. A row with an org_id and no subscriber_id is sent with its org_id. A row completed in the same
environment is skipped; empty, Failed and other-environment rows are (re)tried.
A real run also needs the environment variable ALLOW_REAL_API=yes.
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
    maxRows: Infinity,
    env: '',
    confirmReal: false,
    help: false,
  };
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
      case '--env': o.env = next().toLowerCase(); break;
      case '--yes-create-real-credentials': o.confirmReal = true; break;
      case '--help': case '-h': o.help = true; break;
      default: die(`Unknown option: ${a}\n${HELP}`);
    }
  }
  if (o.env && !['uat', 'prod'].includes(o.env)) die('--env must be "uat" or "prod".');
  if (o.maxRows !== Infinity && (!Number.isInteger(o.maxRows) || o.maxRows < 1)) die('--max-rows must be a positive whole number.');
  o.baseUrl = o.env ? ENV_URLS[o.env] : '';
  return o;
}

/* ------------------------------------------------------------------ safety guards */

function assertRealAllowed(o) {
  const name = o.env.toUpperCase();
  const problems = [];
  if (!o.baseUrl) problems.push(`${name}_BASE_URL is not set in .env`);
  else if (!/^https?:\/\//i.test(o.baseUrl)) problems.push(`${name}_BASE_URL must start with https:// or http://`);
  if (!o.confirmReal) problems.push('--yes-create-real-credentials is required');
  if (process.env.ALLOW_REAL_API !== 'yes') problems.push('environment variable ALLOW_REAL_API=yes is required');
  if (problems.length) {
    die(`Real run refused. No request was sent.\n  - ${problems.join('\n  - ')}\n\nRun without --env to use the safe local mock.`);
  }
  const seconds = o.env === 'prod' ? 10 : 5;
  console.log(`\n!!! REAL RUN IN ${name} !!!`);
  console.log(`!!! Requests will be sent to ${o.baseUrl} and REAL credentials may be created.`);
  console.log(`!!! Press Ctrl+C within ${seconds} seconds to cancel.\n`);
  return new Promise((resolve) => { setTimeout(resolve, seconds * 1000); });
}

/** Excel keeps a hidden "~$name.xlsx" file while a workbook is open. Writing then would fail or clash. */
function findExcelLockFile(filePath) {
  const dir = path.dirname(filePath);
  const base = path.basename(filePath);
  return [`~$${base}`, `~$${base.slice(2)}`]
    .map((n) => path.join(dir, n))
    .find((p) => fs.existsSync(p));
}

/* ------------------------------------------------------------------ mock API (same answers as the real one) */

// Fake environment_access data for mock mode (keep in sync with SAMPLE_ACCESS in static/index.html).
// id = subscriber_id, org = organisation_id, status = environment_access.status.
const MOCK_ACCESS = [
  { id: '70001', org: 6533, status: 1 },
  { id: '70001', org: 6533, status: 7 },
  { id: '70002', org: 7952, status: 7 },
  { id: '70003', org: 4900, status: 1 },
  { id: '70004', org: 2338, status: 4 },
  { id: '70005', org: 1504, status: 1 },
  { id: '70006', org: 9217, status: 1 },
  { id: '70007', org: 1001, status: 2 },
  { id: '70008', org: 1002, status: 3 },
  { id: '70009', org: 1003, status: 1 },
  { id: '70009', org: 1004, status: 1 },
  { id: '70010', org: 6129, status: 7 },
];
const MOCK_STATUS_NAMES = {
  0: 'Inactive', 1: 'Active', 2: 'Pending', 3: 'Rejected', 4: 'Whitelisted', 7: 'Subscribed',
};

/** Returns { organisationId } or { error } like the API's subscriber id lookup. */
function mockFindOrganisation(subscriberId) {
  const id = String(subscriberId).trim().toLowerCase();
  const rows = MOCK_ACCESS.filter((x) => x.id === id);
  if (!rows.length) return { error: 'Subscriber ID not found' };
  const orgs = [...new Set(rows.filter((x) => [1, 4, 7].includes(x.status)).map((x) => x.org))];
  if (orgs.length === 1) return { organisationId: orgs[0] };
  if (orgs.length > 1) return { error: `Subscriber ID matches several organisations (${orgs.join(', ')})` };
  const seen = [...new Set(rows.map((x) => MOCK_STATUS_NAMES[x.status]))].join(', ');
  return { error: `Subscriber ID found, but it is not registered (status: ${seen})` };
}

function startMockServer() {
  return new Promise((resolve) => {
    const server = http.createServer((req, res) => {
      let body = '';
      req.on('data', (c) => { body += c; });
      req.on('end', () => {
        const send = (code, type, out) => { res.writeHead(code, { 'Content-Type': type }); res.end(out); };
        const json = (code, obj) => send(code, 'application/json', JSON.stringify(obj));
        const error = (code, msg) => json(code, { status: 'error', code, error: { description: 'Validation Error', message: { err_msg: msg } } });

        if (req.method !== 'POST' || req.url !== ROUTE) return error(404, 'Not found');
        let p;
        try { p = JSON.parse(body); } catch { return error(400, 'Invalid JSON'); }
        if (![2, 3].includes(p.platform_id)) return error(400, '"Platform_id" must be one of [2, 3]');
        if (p.subscriber_id === undefined && p.organisation_id === undefined) return error(400, 'Send subscriber_id or organisation_id');
        if (String(p.subscriber_id).trim().toLowerCase() === MOCK_RATE_LIMIT_ID) {
          return send(429, 'text/plain', 'Too many requests, please try again later.');
        }

        let org = p.organisation_id;
        if (p.subscriber_id !== undefined) {
          const found = mockFindOrganisation(p.subscriber_id);
          if (found.error) return error(400, found.error);
          if (org !== undefined && org !== found.organisationId) return error(400, `Subscriber ID belongs to organisation ${found.organisationId}, not ${org}`);
          org = found.organisationId;
        }
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

const HALT_MESSAGES = {
  'rate-limit': 'The API rate limit was reached (HTTP 429). Try again later.',
  timeout: 'A request timed out, so its outcome is unknown. Check that organisation before retrying.',
  network: 'Could not reach the API.',
  'bad-endpoint': 'The API has no create-token-by-subscriber endpoint at this address (HTTP 404). Check the address.',
};
const HALT_KINDS = new Set(Object.keys(HALT_MESSAGES));

/** Always resolves to { ok, kind, message, organisationId }. Never throws. Credentials in the reply are discarded. */
async function createToken(baseUrl, payload) {
  const controller = new AbortController();
  const timer = setTimeout(() => controller.abort(), REQUEST_TIMEOUT_MS);
  try {
    const res = await fetch(`${baseUrl}${ROUTE}`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(payload),
      signal: controller.signal,
    });
    const text = await res.text();
    let body = null;
    try { body = JSON.parse(text); } catch { /* not JSON */ }

    if (res.status === 429) return { ok: false, kind: 'rate-limit', message: HALT_MESSAGES['rate-limit'] };
    if (res.status === 404 && !body) return { ok: false, kind: 'bad-endpoint', message: HALT_MESSAGES['bad-endpoint'] };
    if (!body || typeof body !== 'object') return { ok: false, kind: 'bad-response', message: `Unexpected non-JSON response (HTTP ${res.status})` };
    if (body.status === 'success') {
      const data = body.data || {};
      const existing = /already exists/i.test(data.message || '');
      return {
        ok: true,
        kind: existing ? 'existing' : 'created',
        message: existing ? 'Token already existed' : 'Token created',
        organisationId: data.organisation_id,
      };
    }
    return { ok: false, kind: 'error', message: readableError(body) };
  } catch (err) {
    if (err.name === 'AbortError') return { ok: false, kind: 'timeout', message: HALT_MESSAGES.timeout };
    return { ok: false, kind: 'network', message: `Could not reach ${baseUrl}` };
  } finally {
    clearTimeout(timer);
  }
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

  const envName = o.env ? o.env.toUpperCase() : 'MOCK';
  const isCsv = /\.csv$/i.test(o.file);
  const outPath = o.inPlace
    ? o.file
    : (o.out || o.file.replace(/\.(xlsx|csv)$/i, '') + `.${o.env || 'mock'}-result` + (isCsv ? '.csv' : '.xlsx'));

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
    cols[cell.text.trim().replace(/^﻿/, '').toLowerCase()] = colNumber;
    lastCol = Math.max(lastCol, colNumber);
  });
  if (!cols.igm_version || (!cols.subscriber_id && !cols.org_id)) {
    die(`Row 1 must contain the headers subscriber_id (or org_id) and IGM_version. Found: ${Object.keys(cols).join(', ') || '(nothing)'}`);
  }
  // Add the output columns that are missing, copying the header style.
  const styleFrom = lastCol;
  [['org_id', 12], ['status', 18], ['message', 60], ['environment', 14]].forEach(([name, width]) => {
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
    const envText = ws.getCell(r, cols.environment).text.trim();
    if (!idText && !orgText && !verText && !statusText) continue; // blank row
    rows.push({
      r, idText, orgText, verText, statusText, envText,
    });
  }
  // Completed in this environment (or with no environment recorded) means done. Completed in the OTHER
  // environment does not: the same file can be run in UAT first and in PROD afterwards.
  const isDone = (x) => x.statusText.toLowerCase() === STATUS_DONE.toLowerCase()
    && (!x.envText || x.envText.toUpperCase() === envName);
  const pending = rows.filter((x) => !isDone(x));
  const alreadyDone = rows.length - pending.length;

  console.log('============================================================');
  console.log(o.env
    ? ` ENVIRONMENT: ${envName}  (${o.baseUrl || 'address not set'})`
    : ' ENVIRONMENT: MOCK  (local fake API - nothing real is touched)');
  console.log(` Input : ${o.file}`);
  console.log(` Output: ${outPath}${o.inPlace ? '  (in place)' : ''}`);
  console.log(` Rows  : ${rows.length} with data, ${alreadyDone} already Completed (skipped), ${pending.length} to process`);
  console.log('============================================================');

  let baseUrl;
  let mock = null;
  if (!o.env) {
    mock = await startMockServer();
    baseUrl = mock.baseUrl;
    if (!/^http:\/\/127\.0\.0\.1:\d+$/.test(baseUrl)) die('Internal safety check failed: mock URL is not local.');
  } else {
    await assertRealAllowed(o);
    baseUrl = o.baseUrl;
  }

  const setCell = (r, col, value) => {
    const cell = ws.getCell(r, col);
    cell.value = value;
    if (!cell.font || !cell.font.name) cell.style = { ...ws.getCell(r, cols.igm_version).style };
  };
  const setStatus = (row, status, message, orgId) => {
    setCell(row.r, cols.status, status);
    setCell(row.r, cols.message, message);
    setCell(row.r, cols.environment, envName);
    // The org id of a previous run belongs to another environment's database: replace it, never trust it.
    if (orgId) setCell(row.r, cols.org_id, orgId);
    else if (row.idText) setCell(row.r, cols.org_id, '');
  };

  const counts = { completed: 0, failed: 0, notAttempted: 0 };
  let attempted = 0;
  let halted = '';

  for (let i = 0; i < pending.length; i += 1) {
    const row = pending[i];
    const who = row.idText ? `subscriber_id ${row.idText}` : `org_id ${row.orgText || '-'}`;
    const label = `Row ${String(row.r).padStart(3)} | ${who} | ${row.verText || '-'}`;
    const fail = (reason) => {
      setStatus(row, STATUS_FAIL, reason, null);
      counts.failed += 1;
      console.log(`${label} -> ${STATUS_FAIL}  (${reason}, no request sent)`);
    };

    // Checks the API does not do: what to send, and whether the row is usable at all.
    const platformId = toPlatformId(row.verText);
    const orgId = toOrgId(row.orgText);
    const payload = {};
    if (row.idText) payload.subscriber_id = row.idText;
    else if (row.orgText) {
      if (orgId === null) { fail('invalid org_id'); continue; }
      payload.organisation_id = orgId;
    } else { fail(cols.subscriber_id ? 'subscriber_id is empty' : 'org_id is empty'); continue; }
    if (platformId === null) { fail('invalid IGM_version (use IGMv1 or IGMv2)'); continue; }
    payload.platform_id = platformId;

    if (attempted >= o.maxRows) {
      counts.notAttempted += pending.length - i;
      console.log(`\nReached --max-rows (${o.maxRows}); ${pending.length - i} row(s) left untouched for the next run.`);
      break;
    }

    attempted += 1;
    // eslint-disable-next-line no-await-in-loop
    const result = await createToken(baseUrl, payload);

    if (HALT_KINDS.has(result.kind)) {
      // The outcome of this row is unknown (or the API refuses everything): leave it empty for the next run.
      halted = result.message;
      counts.notAttempted += pending.length - i;
      break;
    }
    if (result.ok) {
      setStatus(row, STATUS_DONE, result.message, result.organisationId);
      counts.completed += 1;
    } else {
      setStatus(row, STATUS_FAIL, result.message, null);
      counts.failed += 1;
    }
    console.log(`${label} -> ${result.ok ? STATUS_DONE : STATUS_FAIL}  (platform ${platformId}: ${result.message})`);
  }

  if (mock) { mock.server.closeAllConnections(); mock.server.close(); }

  // Write atomically: temp file first, then rename over the target.
  const tmp = `${outPath}.tmp-${process.pid}`;
  if (isCsv) await wb.csv.writeFile(tmp); else await wb.xlsx.writeFile(tmp);
  fs.renameSync(tmp, outPath);

  console.log('\n------------------------------------------------------------');
  console.log(` Completed: ${counts.completed}   Failed: ${counts.failed}   Left untouched: ${counts.notAttempted}   Skipped (already Completed): ${alreadyDone}`);
  if (halted) console.log(` RUN STOPPED EARLY: ${halted}\n Rows after the stop were left empty so the next run picks them up.`);
  console.log(` Saved: ${outPath}`);
  console.log('------------------------------------------------------------');
  if (halted) process.exitCode = 1;
}

main().catch((err) => {
  console.error(`\nUnexpected error: ${err && err.message ? err.message : err}\n`);
  process.exit(1);
});
