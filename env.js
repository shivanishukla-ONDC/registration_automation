'use strict';

/*
 * Loads registration_automation/.env into process.env (real environment variables always win).
 * Parsed by hand so no extra package is needed. Shared by script.js and server.js.
 */
const fs = require('fs');
const path = require('path');

const envPath = path.join(__dirname, '.env');
if (fs.existsSync(envPath)) {
  fs.readFileSync(envPath, 'utf8').split(/\r?\n/).forEach((line) => {
    const m = line.match(/^\s*([A-Za-z_][A-Za-z0-9_]*)\s*=\s*(.*?)\s*$/);
    if (!m || m[0].trim().startsWith('#')) return;
    let value = m[2];
    if (/^(['"]).*\1$/.test(value)) value = value.slice(1, -1);
    if (value !== '' && process.env[m[1]] === undefined) process.env[m[1]] = value;
  });
}
