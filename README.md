# Registration automation (hosted page)

Upload a CSV with `Share your pre production subscriber ID/IDs` and `Share the IGM version` columns. Each row is sent to the onboarding API's
`POST /api/plan/create-token-by-subscriber` in the selected environment (UAT or PROD). If the row has an `org_id` it is used first; otherwise the API finds the organisation
for the subscriber URL (put the URL in that column) and creates the IGM token. You download the CSV back with `org_id`, `status`, `message` and
`environment` filled in.

This folder has no database access and no secrets. It only needs the address of each environment's API.

## Files
- `static/index.html` - the page: layout, styling and the two address settings (`UAT_URL`, `PROD_URL`) at the top.
- `static/script.js` - the page's code: reads the CSV, calls the API row by row, UAT | PROD switch, results.
  Host these two files **together** (same folder).
- `server.js` - serves `static/` (no packages needed). Only for hosts that need a start command.
- `package.json` - `npm start` runs `server.js`.

## Setup: two values
At the top of `static/index.html` (the only place to edit per environment; `script.js` does not change):
```js
var REGISTRATION_CONFIG = { UAT_URL: null, PROD_URL: null };
```
| Value | Meaning |
|---|---|
| `null` | this environment is not set up (its side of the switch is disabled) |
| `''` | the API on the same server that serves the page |
| `'https://host.example'` | the API at that address (no path) |

If both are `null` the page shows "NOT SET UP" and Run stays disabled.
The page is served over https, so the API addresses must be https too. The API route allows calls from other websites (CORS).

## Hosting (e.g. Render)
- **Static Site:** publish the `static` folder (no build command).
- **Web Service:** build `npm install` (or none), start `npm start`. It listens on the `PORT` the host gives it.

## CSV format
Sample file: `static/mock.csv` (made-up values only; replace them with your own). Keep the first line exactly as it is.

| Column | Required? | What to enter | Example |
|---|---|---|---|
| `Share your pre production subscriber ID/IDs` | One of this or `org_id` | The **subscriber URL, written without `http://` or `https://`** (and without a trailing `/`) | `seller-app.example.com` |
| `Share the IGM version` | Yes | `IGMv1` or `IGMv2` (also accepted: `v1`, `v2`, `2`, `3`) | `IGMv1` |
| `org_id` | One of this or the subscriber URL | The organisation id (a whole number). It is used first; leave it empty to use the subscriber URL | `1001` |

`status`, `message` and `environment` are added by the page; do not fill them in.

The four rows in `mock.csv` show the cases: subscriber URL only (two rows), `org_id` only
(first column empty) and both (the `org_id` is used and the URL is ignored). A row with neither value fails.
The format notes are here and not inside the CSV, because every line of a CSV is read as a row.

## How it behaves
- **UAT | PROD switch:** starts on UAT; locked while a run is going. PROD turns the page red and needs the word `PROD` typed.
- Every run creates real credentials in the selected environment, after you tick the confirmation box.
- Rows are sent one at a time. A row already `Completed` in the selected environment is skipped; a row completed in the
  other environment is processed again. Each row records its `environment`.
- If the API's rate limit (HTTP 429), a timeout, or an unreachable API stops the run, the remaining rows are left empty;
  upload the result file again later to continue.
- Credentials in the API's reply are ignored: never shown, never saved.
