# Registration automation

Upload a CSV with `subscriber_id` and `IGM_version` columns. Each row is sent to the onboarding API's
`POST /api/plan/create-token-by-subscriber` (UAT or PROD). The API finds the organisation for the subscriber id and
creates the token. You get the file back with `org_id`, `status`, `message` and `environment` filled in.

This folder has **no database access**. It only needs the address of each environment's API.

## Pieces
- `static/index.html` - the web page (one file). Host it anywhere.
- `script.js` - the same logic on the command line (`.csv` or `.xlsx`).
- `env.js` - loads `.env` for `script.js`.

## Web page
Open `static/index.html` (or `npm start` and visit http://127.0.0.1:5500).

At the top of the file are the only settings:
```js
var REGISTRATION_CONFIG = { UAT_URL: null, PROD_URL: null };
```
| Value | Meaning |
|---|---|
| `null` | this environment is not set up (its side of the switch is disabled) |
| `''` | the API on the same server that serves the page |
| `'https://host.example'` | the API at that address |

The **UAT | PROD** switch decides which address is used. It starts on UAT and is locked while a run is going.
PROD turns the page red and asks you to type `PROD` before Run is enabled. If both are `null` the page is a **demo**
(built-in sample data, fake API, nothing real). The API route allows calls from other websites (CORS), so the page can be
hosted on a different server than the API; use `https` for both, or `http` for both.

## Hosting (e.g. Render)
Either of these works:
- **Static Site:** publish the `static` folder (no build command, no start command).
- **Web Service:** build command `npm install` (or empty), start command `npm start`. `server.js` only serves the page
  (no database, no extra packages) and listens on the `PORT` the host gives it.

After deploying, set `UAT_URL` / `PROD_URL` in `static/index.html` (and push) to point the page at the real environments.

## Command line
```
node script.js --file customers.csv                       # mock: fake API, nothing real
ALLOW_REAL_API=yes node script.js --file customers.csv --env uat --yes-create-real-credentials
```
Put `UAT_BASE_URL` and `PROD_BASE_URL` in `.env` (copy `.env.example`). `--help` lists all options.

## Rules worth knowing
- Every real run creates real credentials in the selected environment.
- A row completed in UAT is **not** skipped in PROD (and the other way round): each row records its `environment`,
  and `org_id` is replaced from the environment you run. A row completed in the same environment is skipped, so a
  stopped run can simply be run again.
- If the API's rate limit (HTTP 429) is hit, or it cannot be reached, the run stops and the remaining rows are left empty.
- Credentials in the API reply are ignored: never shown, never saved.
- Subscriber ids are matched ignoring case and surrounding spaces.
