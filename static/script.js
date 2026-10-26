/*
 * Registration automation - the page's script. Loaded by index.html (after the UAT_URL / PROD_URL settings there).
 *   1. RegCore: reads the CSV, checks every row, calls the API row by row, builds the result CSV.
 *   2. The page wiring: UAT | PROD switch, confirmations, file upload, Run button, results table.
 */

/*
 * The logic of the page. The API does the lookup (subscriber id -> organisation) and creates the token.
 * This page does everything else: reads the CSV, checks each row, turns IGMv1 / IGMv2 into platform_id 2 / 3,
 * calls the API once per row, skips finished rows, stops cleanly on a rate limit or network problem,
 * and fills org_id, status, message and environment into the CSV you download.
 */
var RegCore = (function () {
  const STATUS_DONE = 'Completed';
  const STATUS_FAIL = 'Failed';
  const REQUEST_TIMEOUT_MS = 30000;
  const ROUTE = '/api/plan/create-token-by-subscriber';

  const text = (v) => String(v === null || v === undefined ? '' : v).trim();

  function toPlatformId(value) {
    const s = text(value).toLowerCase().replace(/[\s_-]+/g, '');
    if (s === 'igmv1' || s === 'v1' || s === '2') return 2;
    if (s === 'igmv2' || s === 'v2' || s === '3') return 3;
    return null;
  }

  function toOrgId(value) {
    const s = text(value);
    return /^[1-9]\d*$/.test(s) ? Number(s) : null;
  }

  /* ---------------------------------------------------------------- CSV */

  /** Parses CSV text into an array of rows (arrays of strings). Handles quotes, commas, CRLF and a BOM. */
  function parseCsv(input) {
    const src = String(input).replace(/^﻿/, '');
    const rows = [];
    let row = [];
    let field = '';
    let quoted = false;
    for (let i = 0; i < src.length; i += 1) {
      const ch = src[i];
      if (quoted) {
        if (ch === '"') {
          if (src[i + 1] === '"') { field += '"'; i += 1; } else quoted = false;
        } else field += ch;
      } else if (ch === '"') quoted = true;
      else if (ch === ',') { row.push(field); field = ''; }
      else if (ch === '\n' || ch === '\r') {
        if (ch === '\r' && src[i + 1] === '\n') i += 1;
        row.push(field); field = '';
        rows.push(row); row = [];
      } else field += ch;
    }
    if (field !== '' || row.length) { row.push(field); rows.push(row); }
    return rows;
  }

  /** A value starting with = + - @ would be run as a formula by Excel; prefix it so it stays text. */
  function safeCell(value) {
    const s = String(value === null || value === undefined ? '' : value);
    return /^[=+\-@\t\r]/.test(s) ? `'${s}` : s;
  }

  function toCsv(rows) {
    return rows.map((r) => r.map((v) => {
      const s = safeCell(v);
      return /[",\r\n]/.test(s) ? `"${s.replace(/"/g, '""')}"` : s;
    }).join(',')).join('\r\n') + '\r\n';
  }

  /* ---------------------------------------------------------------- the API: one create-token-by-subscriber call */

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
    network: 'Could not reach the API. Check the address, that the API allows this page (CORS), and that http / https match.',
    'bad-endpoint': 'The API has no create-token-by-subscriber endpoint at this address (HTTP 404). Check the address.',
  };
  const HALT_KINDS = ['rate-limit', 'timeout', 'network', 'bad-endpoint'];

  /**
   * API backend for one environment address. Resolves { ok:true, kind, message, organisationId } or
   * { ok:false, kind, message }. Credentials in the reply are ignored on purpose.
   */
  function apiBackend(baseUrl) {
    const base = String(baseUrl).replace(/\/+$/, '');
    return {
      createToken(payload) {
        const controller = new AbortController();
        const timer = setTimeout(() => controller.abort(), REQUEST_TIMEOUT_MS);
        return fetch(`${base}${ROUTE}`, {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify(payload),
          signal: controller.signal,
        }).then((res) => res.text().then((raw) => {
          let body = null;
          try { body = JSON.parse(raw); } catch (e) { /* not JSON */ }
          if (res.status === 429) return { ok: false, kind: 'rate-limit', message: HALT_MESSAGES['rate-limit'] };
          if (res.status === 404 && !body) return { ok: false, kind: 'bad-endpoint', message: HALT_MESSAGES['bad-endpoint'] };
          if (!body || typeof body !== 'object') return { ok: false, kind: 'bad-response', message: `Unexpected response from the API (HTTP ${res.status})` };
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
        })).catch((err) => (err && err.name === 'AbortError'
          ? { ok: false, kind: 'timeout', message: HALT_MESSAGES.timeout }
          : { ok: false, kind: 'network', message: HALT_MESSAGES.network }))
          .then((r) => { clearTimeout(timer); return r; });
      },
    };
  }

  /* ---------------------------------------------------------------- processing */

  /**
   * Processes the CSV text row by row for one environment. Returns { csv, rows, counts, halted }.
   * onRow(rowResult, counts) is called after each row so the page can show live progress.
   */
  async function processCustomers(csvText, backend, environment, onRow) {
    const table = parseCsv(csvText);
    if (!table.length) throw new Error('The CSV file is empty.');

    const header = table[0].map((h) => text(h));
    const cols = {};
    header.forEach((h, i) => { if (h) cols[h.toLowerCase()] = i; });
    if (cols.igm_version === undefined || (cols.subscriber_id === undefined && cols.org_id === undefined)) {
      throw new Error(`Row 1 must contain the headers subscriber_id and IGM_version. Found: ${header.filter(Boolean).join(', ') || '(nothing)'}`);
    }
    ['org_id', 'status', 'message', 'environment'].forEach((name) => {
      if (cols[name] !== undefined) return;
      cols[name] = header.length;
      header.push(name);
      table[0][cols[name]] = name;
    });

    const get = (r, name) => (cols[name] === undefined ? '' : text(table[r][cols[name]]));
    const set = (r, name, value) => { while (table[r].length <= cols[name]) table[r].push(''); table[r][cols[name]] = value; };

    const rows = [];
    for (let r = 1; r < table.length; r += 1) {
      const idText = get(r, 'subscriber_id');
      const orgText = get(r, 'org_id');
      const verText = get(r, 'igm_version');
      const statusText = get(r, 'status');
      if (!idText && !orgText && !verText && !statusText) continue; // blank line
      rows.push({
        r, idText, orgText, verText, statusText, envText: get(r, 'environment'),
      });
    }
    // Completed in this environment (or with no environment recorded) means done. Completed in the OTHER
    // environment does not: the same file can be run in UAT first and in PROD afterwards.
    const isDone = (x) => x.statusText.toLowerCase() === STATUS_DONE.toLowerCase()
      && (!x.envText || x.envText.toUpperCase() === environment.toUpperCase());
    const pending = rows.filter((x) => !isDone(x));

    const results = [];
    const counts = {
      completed: 0, failed: 0, notProcessed: 0, skipped: rows.length - pending.length,
    };
    let halted = '';

    const record = (row, status, message, orgId) => {
      set(row.r, 'status', status);
      set(row.r, 'message', message);
      set(row.r, 'environment', environment);
      // The org id of a previous run belongs to another environment's database: replace it, never trust it.
      if (orgId) set(row.r, 'org_id', String(orgId));
      else if (!row.idText) { /* org_id-only row: keep the org id the user gave */ } else set(row.r, 'org_id', '');
      if (status === STATUS_DONE) counts.completed += 1; else counts.failed += 1;
      const res = {
        row: row.r + 1,
        subscriber_id: row.idText,
        igm_version: row.verText,
        org_id: orgId ? String(orgId) : (row.idText ? '' : row.orgText),
        status,
        message,
      };
      results.push(res);
      if (onRow) onRow(res, counts);
    };

    for (let i = 0; i < pending.length; i += 1) {
      const row = pending[i];

      // Checks the API does not do: what to send, and whether the row is usable at all.
      const platformId = toPlatformId(row.verText);
      const orgId = toOrgId(row.orgText);
      const payload = {};
      if (row.idText) payload.subscriber_id = row.idText;
      else if (row.orgText) {
        if (orgId === null) { record(row, STATUS_FAIL, 'invalid org_id', null); continue; }
        payload.organisation_id = orgId;
      } else { record(row, STATUS_FAIL, cols.subscriber_id !== undefined ? 'subscriber_id is empty' : 'org_id is empty', null); continue; }
      if (platformId === null) { record(row, STATUS_FAIL, 'invalid IGM_version (use IGMv1 or IGMv2)', null); continue; }
      payload.platform_id = platformId;

      // eslint-disable-next-line no-await-in-loop
      const result = await backend.createToken(payload);
      if (result.ok) {
        record(row, STATUS_DONE, result.message, result.organisationId);
      } else if (HALT_KINDS.includes(result.kind)) {
        // The outcome of this row is unknown (or the API refuses everything): leave it empty for the next run.
        halted = result.message;
        counts.notProcessed += pending.length - i;
        break;
      } else {
        record(row, STATUS_FAIL, result.message, null);
      }
    }
    return { csv: toCsv(table), rows: results, counts, halted };
  }

  // Header row only: no data. Users fill in their own subscriber ids.
  const TEMPLATE_CSV = 'subscriber_id,IGM_version\n';

  return { TEMPLATE_CSV, apiBackend, processCustomers };
}());

(function () {
  var $ = function (id) { return document.getElementById(id); };
  var cfg = window.REGISTRATION_CONFIG || {};
  var urls = { UAT: cfg.UAT_URL, PROD: cfg.PROD_URL };
  var isSet = function (name) { return typeof urls[name] === 'string'; };
  var configured = isSet('UAT') || isSet('PROD');
  var env = isSet('UAT') ? 'UAT' : (isSet('PROD') ? 'PROD' : null); // starts on UAT, never on PROD by itself
  var running = false;
  var cus = null;
  var blobUrl = null;

  function showError(msg) { $('error').textContent = msg; $('error').hidden = !msg; }

  function refresh() {
    var ok = configured && !!cus && !running && $('confirm').checked;
    if (env === 'PROD') ok = ok && $('prodType').value.trim() === 'PROD';
    $('run').disabled = !ok;
    $('envUat').disabled = running || !isSet('UAT');
    $('envProd').disabled = running || !isSet('PROD');
  }

  /* ---------- environment switch ---------- */
  function applyEnv() {
    var prod = env === 'PROD';
    document.body.className = prod ? 'prod' : '';
    $('envUat').setAttribute('aria-checked', String(env === 'UAT'));
    $('envProd').setAttribute('aria-checked', String(prod));
    var shown = urls[env] === '' ? window.location.host + ' (this server)' : urls[env];
    $('modeBadge').textContent = env + (prod ? ' - REAL CREDENTIALS' : '');
    $('modeBadge').className = 'badge' + (prod ? ' prod' : '');
    $('confirmLabel').className = 'check' + (prod ? ' prod' : '');
    $('confirmText').textContent = 'I understand this creates REAL credentials in ' + env + ' (' + shown + ').';
    $('prodTypeRow').hidden = !prod;
    $('confirm').checked = false;
    $('prodType').value = '';
    $('results').hidden = true;
    showError('');
    refresh();
  }

  function pick(next) {
    if (running || next === env || !isSet(next)) return;
    env = next;
    applyEnv();
  }

  if (configured) {
    $('envUat').title = isSet('UAT') ? '' : 'UAT address is not set';
    $('envProd').title = isSet('PROD') ? '' : 'PROD address is not set';
    $('envUat').addEventListener('click', function () { pick('UAT'); });
    $('envProd').addEventListener('click', function () { pick('PROD'); });
    $('confirm').addEventListener('change', refresh);
    $('prodType').addEventListener('input', refresh);
    applyEnv();
  } else {
    $('notConfigured').hidden = false;
    $('modeBadge').textContent = 'NOT SET UP';
    $('modeBadge').className = 'badge off';
    $('confirmText').textContent = 'I understand this creates REAL credentials.';
    refresh();
  }

  /* ---------- file handling ---------- */
  function setFile(name, content) {
    showError('');
    cus = { name: name, text: content };
    $('txtCus').textContent = name;
    $('dropCus').classList.add('ready');
    refresh();
  }

  function readFile(f) {
    if (!f) return;
    if (!/\.csv$/i.test(f.name)) { showError('Please choose a .csv file.'); return; }
    if (f.size > 5 * 1024 * 1024) { showError('File is too large (max 5 MB).'); return; }
    var fr = new FileReader();
    fr.onload = function () { setFile(f.name, fr.result); };
    fr.readAsText(f);
  }

  var drop = $('dropCus'), input = $('fileCus');
  drop.addEventListener('click', function () { input.click(); });
  drop.addEventListener('keydown', function (e) { if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); input.click(); } });
  drop.addEventListener('dragover', function (e) { e.preventDefault(); drop.classList.add('over'); });
  drop.addEventListener('dragleave', function () { drop.classList.remove('over'); });
  drop.addEventListener('drop', function (e) { e.preventDefault(); drop.classList.remove('over'); readFile(e.dataTransfer.files[0]); });
  input.addEventListener('change', function () { readFile(input.files[0]); });
  $('templateDownload').addEventListener('click', function () {
    var a = document.createElement('a');
    a.href = URL.createObjectURL(new Blob([RegCore.TEMPLATE_CSV], { type: 'text/csv' }));
    a.download = 'registration_template.csv';
    document.body.appendChild(a);
    a.click();
    setTimeout(function () { URL.revokeObjectURL(a.href); a.remove(); }, 1000);
  });

  /* ---------- results ---------- */
  function cell(tr, value, cls) {
    var td = document.createElement('td');
    if (cls) td.className = cls;
    td.textContent = value || '';
    tr.appendChild(td);
    return td;
  }

  function addRow(r, counts) {
    var tr = document.createElement('tr');
    cell(tr, String(r.row));
    cell(tr, r.subscriber_id, 'url');
    cell(tr, r.igm_version);
    cell(tr, r.org_id);
    var pill = document.createElement('span');
    pill.className = 'pill ' + r.status;
    pill.textContent = r.status;
    cell(tr, '').appendChild(pill);
    cell(tr, r.message);
    $('tbody').appendChild(tr);
    $('nOk').textContent = counts.completed;
    $('nBad').textContent = counts.failed;
    $('busy').textContent = 'Processed ' + (counts.completed + counts.failed) + ' rows in ' + env + '... please do not close this page.';
  }

  $('run').addEventListener('click', function () {
    showError('');
    running = true;
    refresh();
    $('results').hidden = false;
    $('download').hidden = true;
    $('tbody').textContent = '';
    $('halted').hidden = true;
    ['nOk', 'nBad', 'nLeft', 'nSkip'].forEach(function (id) { $(id).textContent = '0'; });
    $('busy').hidden = false;
    $('busy').textContent = 'Processing in ' + env + '... please do not close this page.';
    var runEnv = env; // the environment cannot change while a run is going (the switch is locked)

    RegCore.processCustomers(cus.text, RegCore.apiBackend(urls[runEnv]), runEnv, addRow)
      .then(function (res) {
        $('nOk').textContent = res.counts.completed;
        $('nBad').textContent = res.counts.failed;
        $('nLeft').textContent = res.counts.notProcessed;
        $('nSkip').textContent = res.counts.skipped;
        if (res.halted) {
          $('halted').textContent = 'Stopped early: ' + res.halted + ' Rows after the stop were left empty; run again later to pick them up.';
          $('halted').hidden = false;
        }
        if (blobUrl) URL.revokeObjectURL(blobUrl);
        // BOM first so Excel reads the file as UTF-8.
        blobUrl = URL.createObjectURL(new Blob(['﻿' + res.csv], { type: 'text/csv;charset=utf-8' }));
        $('download').href = blobUrl;
        $('download').download = cus.name.replace(/\.csv$/i, '') + '.result.' + runEnv.toLowerCase() + '.csv';
        $('download').hidden = false;
      })
      .catch(function (err) { $('results').hidden = true; showError(err && err.message ? err.message : 'Something went wrong.'); })
      .then(function () { running = false; $('busy').hidden = true; refresh(); });
  });
})();
