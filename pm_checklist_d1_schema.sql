-- Cloudflare D1 schema: PM Checklist
-- Database: portal-db
-- Safe to run once. This creates application tables only; it does not delete or modify Google Sheets data.
-- Source fields are based on getPMChecklist(sessionId, requestedMonth) in the supplied Code.gs.

CREATE TABLE IF NOT EXISTS pm_checklist_rows (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  source_sheet_id TEXT NOT NULL,
  source_sheet_name TEXT NOT NULL,
  source_row_number INTEGER NOT NULL,
  due_date TEXT NOT NULL,              -- YYYY-MM-DD
  module TEXT NOT NULL DEFAULT '',
  component TEXT NOT NULL DEFAULT '',
  checkpoint TEXT NOT NULL DEFAULT '',
  frequency TEXT NOT NULL DEFAULT '',
  condition TEXT NOT NULL DEFAULT '',
  status TEXT NOT NULL DEFAULT 'Open',
  issue_description TEXT NOT NULL DEFAULT '',
  correction_method TEXT NOT NULL DEFAULT '',
  eta TEXT NOT NULL DEFAULT '',
  eta_iso TEXT NOT NULL DEFAULT '',
  imported_at TEXT NOT NULL DEFAULT (strftime('%Y-%m-%dT%H:%M:%fZ','now')),
  updated_at TEXT NOT NULL DEFAULT (strftime('%Y-%m-%dT%H:%M:%fZ','now')),
  UNIQUE (source_sheet_id, source_row_number)
);

CREATE INDEX IF NOT EXISTS idx_pm_checklist_due_date
  ON pm_checklist_rows (due_date);

CREATE INDEX IF NOT EXISTS idx_pm_checklist_module
  ON pm_checklist_rows (module);

CREATE INDEX IF NOT EXISTS idx_pm_checklist_status
  ON pm_checklist_rows (status);

CREATE INDEX IF NOT EXISTS idx_pm_checklist_due_status
  ON pm_checklist_rows (due_date, status);

-- Check the created schema after running:
-- SELECT name, type FROM sqlite_master
-- WHERE type IN ('table', 'index') AND name LIKE 'pm_checklist%'
-- ORDER BY type, name;
