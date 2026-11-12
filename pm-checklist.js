const json = (body, status = 200) => new Response(JSON.stringify(body), {
  status, headers: { "Content-Type": "application/json; charset=utf-8", "Cache-Control": "no-store" }
});

async function callGas(env, fn, args) {
  if (!env.GAS_WEB_APP_URL) throw new Error("GAS_WEB_APP_URL is not configured.");
  const response = await fetch(env.GAS_WEB_APP_URL, {
    method: "POST", headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ fn, args }), redirect: "follow"
  });
  const raw = await response.text();
  if (!response.ok) throw new Error(`Apps Script HTTP ${response.status}: ${raw.slice(0, 300)}`);
  let parsed;
  try { parsed = JSON.parse(raw); } catch { throw new Error("Apps Script returned non-JSON output."); }
  return parsed;
}

function toRow(r) {
  const condition = String(r.condition || "").trim();
  return {
    sheetId: Number(r.source_sheet_id), sheetName: r.source_sheet_name || "",
    rowNumber: Number(r.source_row_number), dueDate: r.due_date || "",
    dueDates: r.due_date ? [r.due_date] : [], module: r.module || "",
    component: r.component || "", checkpoint: r.checkpoint || "",
    frequency: r.frequency || "", condition, status: condition.toUpperCase() === "OK" ? "Completed" : "Open",
    issueDescription: r.issue_description || "", correctionMethod: r.correction_method || "",
    eta: r.eta || "", etaISO: r.eta_iso || ""
  };
}

export async function onRequestOptions() { return new Response(null, { status: 204 }); }
export async function onRequestPost({ request, env }) {
  if (!env.DB) return json({ success: false, message: "Cloudflare D1 binding DB is not configured." }, 500);
  let body;
  try { body = await request.json(); } catch { return json({ success: false, message: "Request body must be valid JSON." }, 400); }
  const fn = String(body?.fn || "");
  const args = Array.isArray(body?.args) ? body.args : [];
  const sessionId = String(args[0] || "");
  if (!sessionId) return json({ success: false, message: "Session missing. Please log in again." }, 401);
  try {
    // Always validate the session and effective portal permission against Apps Script.
    const perms = await callGas(env, "getCurrentPortalPermissions", [sessionId]);
    if (!perms || perms.success === false) return json(perms || { success: false, message: "Session validation failed." }, 401);

    if (fn === "getPMChecklist") {
      if (!perms.maintenance?.view) return json({ success: false, message: "Access denied: maintenance view permission is not enabled." }, 403);
      const month = String(args[1] || "").trim();
      let result;
      if (/^\d{4}-\d{2}$/.test(month)) {
        result = await env.DB.prepare("SELECT * FROM pm_checklist_rows WHERE substr(due_date,1,7)=? ORDER BY due_date, source_row_number").bind(month).all();
      } else {
        result = await env.DB.prepare("SELECT * FROM pm_checklist_rows ORDER BY due_date, source_row_number").all();
      }
      const rows = (result.results || []).map(toRow);
      return json({ success: true, rows, today: new Date().toISOString().slice(0,10), dateTabsFound: new Set((result.results || []).map(r => r.source_sheet_name)).size, message: `Loaded ${rows.length} checklist points from D1.` });
    }

    if (fn === "savePMChecklistRow") {
      if (!perms.maintenance?.edit) return json({ success: false, message: "Access denied: maintenance edit permission is not enabled." }, 403);
      const [, sheetId, rowNumber, conditionRaw, etaRaw] = args;
      const sid = String(sheetId ?? ""), rn = Number(rowNumber);
      const condition = String(conditionRaw || "").trim(), eta = String(etaRaw || "").trim();
      if (!sid || !Number.isInteger(rn) || rn < 2) return json({ success: false, message: "Invalid PM checklist row." }, 400);
      if (!["", "OK", "Not OK"].includes(condition)) return json({ success: false, message: "Condition must be OK or Not OK." }, 400);
      if (condition !== "OK" && !eta) return json({ success: false, message: "ETA / next planned date is required while the task is open." }, 400);
      const exists = await env.DB.prepare("SELECT id FROM pm_checklist_rows WHERE source_sheet_id=? AND source_row_number=?").bind(sid, rn).first();
      if (!exists) return json({ success: false, message: "This row is not present in D1. Re-import the PM checklist before saving." }, 404);
      // Apps Script remains the source-of-truth writer for Google Sheets and existing notices.
      const saved = await callGas(env, "savePMChecklistRow", [sessionId, Number(sheetId), rn, condition, eta]);
      if (!saved || saved.success !== true) return json(saved || { success: false, message: "Google Sheets save failed." }, 502);
      const status = condition === "OK" ? "Completed" : "Open";
      await env.DB.prepare("UPDATE pm_checklist_rows SET condition=?, status=?, eta=?, eta_iso=?, updated_at=strftime('%Y-%m-%dT%H:%M:%fZ','now') WHERE source_sheet_id=? AND source_row_number=?")
        .bind(condition, status, eta, /^\d{4}-\d{2}-\d{2}$/.test(eta) ? eta : "", sid, rn).run();
      return json({ success: true, message: "PM checklist updated in Google Sheets and D1." });
    }
    return json({ success: false, message: "Unsupported PM checklist operation." }, 400);
  } catch (e) {
    return json({ success: false, message: e?.message || "PM checklist API failed." }, 500);
  }
}
