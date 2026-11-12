Material Request Portal — Hybrid Search update

Changed:
- Code.gs: Search Item reads migrated rows from Cloudflare D1 table `search_items`
  and appends only source rows not yet represented in D1 from the original
  `Search Item` Google Sheet.
- If D1 read fails, Search Item falls back to the full original Sheet.
- Stock values continue to come from `Automatic Spare Tracker`.
- Existing duplicate Item Master ID handling and filters are retained.

Unchanged companion files:
- index.html
- netlify-api-bridge.js
- functions/api/gas-api.js
- functions/api/pm-checklist.js
- .gitignore

Important:
- This bundle does not modify the PM Checklist spreadsheet or its D1 data.
- Do not delete the original Search Item sheet or backup tab.
- The D1 table must remain `search_items`; D1 records should contain `data.sourceRow`.
- Replace/deploy only after taking a copy of the current project.
