# Static page (local demo)

`index.html` is the whole page: one file, no server, no libraries, no network. Open it by double-clicking it,
or run `npm run static` from `registration_automation/` and visit http://127.0.0.1:5500.

- Upload a CSV with the columns `subscriber_id` and `IGM_version` ("Download the sample" gives an example).
- The `environment_access` table is built into the page (`SAMPLE_ACCESS`), and create-token is a fake
  (same rules as `script.js` mock mode). Nothing real is touched.
- Lookup rules match `script.js`: subscriber_id compared ignoring case and spaces; statuses Active (1),
  Whitelisted (4) and Subscribed (7) count as registered; exactly one organisation must match.
- The result CSV has `org_id`, `status` and `message` added and is saved to your browser's Downloads folder.
