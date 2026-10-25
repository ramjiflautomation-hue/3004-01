# PM Checklist D1 Integration — candidate package

This package adds a dedicated Cloudflare Pages route at `functions/api/pm-checklist.js`; it does **not** change `gas-api.js`.

## What it does
- Reads PM checklist rows from D1 `DB` binding/table `pm_checklist_rows`.
- Checks the existing Apps Script session and `maintenance.view` / `maintenance.edit` permissions before responding.
- For saves, calls the existing Apps Script `savePMChecklistRow` first, preserving its Google Sheet write and configured notification behavior; only after that succeeds does it update the matching D1 row.
- Updates the frontend PM Checklist load/save calls to use `/api/pm-checklist`; other portal calls remain on their existing `google.script.run` path.
- Adds `getCurrentPortalPermissions` and `savePMChecklistRow` to the Apps Script `doPost` allowlist. These functions already exist in the supplied Code.gs; this is dispatcher registration only.

## Required deployment details
1. Keep the Cloudflare Pages `DB` binding pointed at `portal-db` and `GAS_WEB_APP_URL` set as before.
2. Put `functions/api/pm-checklist.js` at the Pages project root under `functions/api/`.
3. Deploy the included `index.html` as the Pages root page.
4. Apply the included `Code.gs` change to the Apps Script project and deploy a new web-app version. This allowlist change is required for permission validation and PM saves through the gateway.
5. Do not rerun the initial bulk import just to deploy the route. D1 table must contain the existing 7,305 rows and have the schema in `pm_checklist_d1_schema.sql`.

## Important behavior / limits
- The source Google Sheet remains authoritative for saves. D1 is updated only after Apps Script confirms the source write.
- Existing notifications are sent by Apps Script as before; this route does not send a second notification.
- If the source spreadsheet changes independently outside this portal, D1 will not automatically refresh. Re-import/sync needs a separate scheduled sync job.
- This package has not been deployed to your Cloudflare/GAS accounts from this chat. Test in a preview deployment first, especially permissions, date filters, and a single test-row save.
