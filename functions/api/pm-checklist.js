
const json = (body, status = 200) => new Response(JSON.stringify(body), {
  status,
  headers: {
    "Content-Type": "application/json; charset=utf-8",
    "Cache-Control": "no-store"
  }
});

async function callGas(env, fn, args) {
  if (!env.GAS_WEB_APP_URL) {
    throw new Error("GAS_WEB_APP_URL is not configured.");
  }

  const response = await fetch(env.GAS_WEB_APP_URL, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ fn, args }),
    redirect: "follow"
  });

  const raw = await response.text();

  if (!response.ok) {
    throw new Error(`Apps Script HTTP ${response.status}: ${raw.slice(0, 300)}`);
  }

  let parsed;
  try {
    parsed = JSON.parse(raw);
  } catch {
    throw new Error("Apps Script returned non-JSON output.");
  }

  return parsed;
}

// Convert common spreadsheet date formats to YYYY-MM-DD.
function toISODate(value) {
  if (value == null || value === "") return "";

  if (value instanceof Date && !Number.isNaN(value.getTime())) {
    return [
      value.getFullYear(),
      String(value.getMonth() + 1).padStart(2, "0"),
      String(value.getDate()).padStart(2, "0")
    ].join("-");
  }

  const s = String(value).trim();
  if (!s) return "";

  let m = s.match(/^(\d{4})[-/](\d{1,2})[-/](\d{1,2})/);

  if (m) {
    const y = Number(m[1]);
    const mo = Number(m[2]);
    const d = Number(m[3]);

    if (validDate(y, mo, d)) {
      return `${y}-${String(mo).padStart(2, "0")}-${String(d).padStart(2, "0")}`;
    }
  }

  m = s.match(/^(\d{1,2})[-/.](\d{1,2})[-/.](\d{4})/);

  if (m) {
    const d = Number(m[1]);
    const mo = Number(m[2]);
    const y = Number(m[3]);

    if (validDate(y, mo, d)) {
      return `${y}-${String(mo).padStart(2, "0")}-${String(d).padStart(2, "0")}`;
    }
  }

  const parsed = new Date(s);

  if (!Number.isNaN(parsed.getTime())) {
    return [
      parsed.getFullYear(),
      String(parsed.getMonth() + 1).padStart(2, "0"),
      String(parsed.getDate()).padStart(2, "0")
    ].join("-");
  }

  return "";
}

function validDate(y, m, d) {
  const dt = new Date(y, m - 1, d);

  return (
    dt.getFullYear() === y &&
    dt.getMonth() === m - 1 &&
    dt.getDate() === d
  );
}

// Convert D1 database row into portal row format.
function toRow(r) {
  const condition = String(r.condition || "").trim();

  const dueDate = toISODate(
    r.due_date || r.dueDate || ""
  );

  const etaISO = toISODate(
    r.eta_iso || r.etaISO || r.eta || ""
  );

  return {
    id: Number(r.id) || 0,

    sheetId: Number(
      r.source_sheet_id || r.sheetId
    ) || 0,

    sheetName:
      r.source_sheet_name ||
      r.sheetName ||
      "",

    rowNumber: Number(
      r.source_row_number || r.rowNumber
    ) || 0,

    dueDate,
    dueDates: dueDate ? [dueDate] : [],

    module: String(r.module || "").trim(),
    component: String(r.component || "").trim(),
    checkpoint: String(r.checkpoint || "").trim(),
    frequency: String(r.frequency || "").trim(),

    condition,

    status:
      condition.toUpperCase() === "OK"
        ? "Completed"
        : "Open",

    issueDescription:
      r.issue_description ||
      r.issueDescription ||
      "",

    correctionMethod:
      r.correction_method ||
      r.correctionMethod ||
      "",

    eta: r.eta || etaISO || "",
    etaISO
  };
}

// Exclude blank and navigation rows from actual task counts.
function isRealTask(row) {
  const searchable = [
    row.module,
    row.component,
    row.checkpoint
  ].join(" ").toLowerCase();

  const hasTaskDescription = !!(
    String(row.module || "").trim() ||
    String(row.component || "").trim() ||
    String(row.checkpoint || "").trim()
  );

  const isNavigationRow =
    searchable.includes("go to main sheet") ||
    searchable.includes("main sheet");

  return (
    !!row.dueDate &&
    hasTaskDescription &&
    !isNavigationRow
  );
}

export async function onRequestOptions() {
  return new Response(null, {
    status: 204,
    headers: {
      "Access-Control-Allow-Origin": "*",
      "Access-Control-Allow-Methods": "POST, OPTIONS",
      "Access-Control-Allow-Headers": "Content-Type"
    }
  });
}

export async function onRequestPost({ request, env }) {
  if (!env.DB) {
    return json({
      success: false,
      message: "Cloudflare D1 binding DB is not configured."
    }, 500);
  }

  let body;

  try {
    body = await request.json();
  } catch {
    return json({
      success: false,
      message: "Request body must be valid JSON."
    }, 400);
  }

  const fn = String(body?.fn || "");
  const args = Array.isArray(body?.args) ? body.args : [];
  const sessionId = String(args[0] || "");

  if (!sessionId) {
    return json({
      success: false,
      message: "Session missing. Please log in again."
    }, 401);
  }

  try {
    // Validate session and portal permissions using Apps Script.
    const perms = await callGas(
      env,
      "getCurrentPortalPermissions",
      [sessionId]
    );

    if (!perms || perms.success === false) {
      return json(
        perms || {
          success: false,
          message: "Session validation failed."
        },
        401
      );
    }

    // GET PM CHECKLIST
    if (fn === "getPMChecklist") {
      if (!perms.maintenance?.view) {
        return json({
          success: false,
          message: "Access denied: maintenance view permission is not enabled."
        }, 403);
      }

      const month = String(args[1] || "all").trim() || "all";

      if (
        month !== "all" &&
        !/^\d{4}-\d{2}$/.test(month)
      ) {
        return json({
          success: false,
          message: "Month must be YYYY-MM or all."
        }, 400);
      }

      // Read D1 records and normalize dates before month filtering.
      const result = await env.DB.prepare(
        "SELECT * FROM pm_checklist_rows ORDER BY source_row_number"
      ).all();

      const allRows = (result.results || [])
        .map(toRow)
        .filter(isRealTask);

      const rows = month === "all"
        ? allRows
        : allRows.filter(
            row => row.dueDate.slice(0, 7) === month
          );

      const dates = new Set(
        rows.map(row => row.dueDate).filter(Boolean)
      );

      return json({
        success: true,
        rows,
        today: new Date().toISOString().slice(0, 10),
        dateTabsFound: dates.size,
        total: rows.length,
        message: `Loaded ${rows.length} valid PM Checklist points from D1.`
      });
    }

    // SAVE PM CHECKLIST ROW
    if (fn === "savePMChecklistRow") {
      if (!perms.maintenance?.edit) {
        return json({
          success: false,
          message: "Access denied: maintenance edit permission is not enabled."
        }, 403);
      }

      const [
        ,
        sheetId,
        rowNumber,
        conditionRaw,
        etaRaw
      ] = args;

      const sid = String(sheetId ?? "");
      const rn = Number(rowNumber);
      const condition = String(conditionRaw || "").trim();
      const eta = String(etaRaw || "").trim();

      if (
        !sid ||
        !Number.isInteger(rn) ||
        rn < 2
      ) {
        return json({
          success: false,
          message: "Invalid PM checklist row."
        }, 400);
      }

      if (!["", "OK", "Not OK"].includes(condition)) {
        return json({
          success: false,
          message: "Condition must be OK or Not OK."
        }, 400);
      }

      if (condition !== "OK" && !eta) {
        return json({
          success: false,
          message: "ETA / next planned date is required while the task is open."
        }, 400);
      }

      const exists = await env.DB.prepare(
        `SELECT id
         FROM pm_checklist_rows
         WHERE source_sheet_id = ?
           AND source_row_number = ?`
      ).bind(sid, rn).first();

      if (!exists) {
        return json({
          success: false,
          message: "This row is not present in D1. Re-import the PM checklist before saving."
        }, 404);
      }

      // Google Sheets remains the source-of-truth writer.
      const saved = await callGas(
        env,
        "savePMChecklistRow",
        [
          sessionId,
          Number(sheetId),
          rn,
          condition,
          eta
        ]
      );

      if (!saved || saved.success !== true) {
        return json(
          saved || {
            success: false,
            message: "Google Sheets save failed."
          },
          502
        );
      }

      const etaISO = toISODate(eta);

      const status =
        condition === "OK"
          ? "Completed"
          : "Open";

      await env.DB.prepare(
        `UPDATE pm_checklist_rows
         SET condition = ?,
             status = ?,
             eta = ?,
             eta_iso = ?,
             updated_at = strftime('%Y-%m-%dT%H:%M:%fZ','now')
         WHERE source_sheet_id = ?
           AND source_row_number = ?`
      ).bind(
        condition,
        status,
        eta,
        etaISO,
        sid,
        rn
      ).run();

      return json({
        success: true,
        message: "PM checklist updated in Google Sheets and D1."
      });
    }

    return json({
      success: false,
      message: "Unsupported PM checklist operation."
    }, 400);

  } catch (e) {
    return json({
      success: false,
      message: e?.message || "PM checklist API failed."
    }, 500);
  }
}
