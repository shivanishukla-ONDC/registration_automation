# Registration automation

Upload a CSV with `subscriber_id` and `IGM_version`. For each row the organisation is found in the
`environment_access` table (by subscriber_id), IGM credentials are created through `create-token`,
and you get the CSV back with `org_id`, `status` and `message` filled in.

## Run on a server - environment variables only
Set these (in `.env`, copied from `.env.example`, or as real environment variables), then `npm install && npm start`:

| Variable | Meaning |
|---|---|
| `TARGET_BASE_URL` | Server that creates the tokens, e.g. `https://preprod.example.com`. Empty = safe local mock. |
| `DB_HOST` `DB_PORT` `DB_USER` `DB_PASSWORD` `DB_DATABASE` (`DB_SSL`) | Database with `environment_access` (use a read-only user, same environment as the URL). |
| `ALLOW_REAL_API=yes` | Required for real mode (it creates real credentials). |
| `ACCESS_KEY` | Required for real mode; users type it into the page. |
| `PORT`, `HOST` | Default `3000` / `127.0.0.1`. Use `HOST=0.0.0.0` to let others open the page. |

Then open `http://<server>:<PORT>`. The page shows the mode and the target server. In real mode tick
**Dry run** first: it looks up every organisation but creates nothing.

## Pieces
- `server.js` - serves the page and runs `script.js` for each upload (`npm start`).
- `script.js` - the same logic on the command line (`npm run cli -- --file x.csv`, `--help`, `--dry-run`).
- `static/index.html` - the page. Served by `server.js`, or open/host it alone as a browser-only demo
  (built-in sample data, fake create-token, ids 70001-70011).
- `env.js` - loads `.env`.
