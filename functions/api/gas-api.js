function json(body, status = 200) {
  return new Response(JSON.stringify(body), {
    status,
    headers: {
      "Content-Type": "application/json; charset=utf-8",
      "Cache-Control": "no-store",
      "Access-Control-Allow-Origin": "*",
      "Access-Control-Allow-Methods": "POST, OPTIONS",
      "Access-Control-Allow-Headers": "Content-Type"
    }
  });
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
  const started = Date.now();
  const gasUrl = env.GAS_WEB_APP_URL;

  // Validate request method
  if (request.method !== "POST") {
    return json({ success: false, message: "POST requests only." }, 405);
  }

  // Check GAS configuration
  if (!gasUrl) {
    return json({
      success: false,
      message: "GAS_WEB_APP_URL environment variable is missing in Cloudflare."
    }, 500);
  }

  // Parse payload
  let payload;
  try {
    payload = await request.json();
  } catch {
    return json({ success: false, message: "Invalid JSON request body." }, 400);
  }

  if (!payload || typeof payload.fn !== "string" || !Array.isArray(payload.args)) {
    return json({ success: false, message: "Expected fn and args in request." }, 400);
  }

  const functionName = payload.fn.trim();

  // Agar request getPMChecklist ya savePMChecklistRow hai aur D1 configure hai, to pm-checklist logic chalega
  if ((functionName === "getPMChecklist" || functionName === "savePMChecklistRow") && env.DB) {
    // Ye direct aapke D1 handler par pass ho sakta hai
  }

  // Baaki sabhi requests (jaise loginUser, getDashboardData, etc.) Apps Script ko forward hongi
  try {
    const upstream = await fetch(gasUrl, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        fn: functionName,
        args: payload.args
      }),
      redirect: "follow",
      signal: AbortSignal.timeout(55000)
    });

    const responseText = await upstream.text();
    if (!upstream.ok) {
      return json({
        success: false,
        message: "Google Apps Script returned an error (" + upstream.status + ").",
        details: responseText.slice(0, 500)
      }, 502);
    }

    try {
      const parsed = JSON.parse(responseText);
      return json(parsed, 200);
    } catch {
      return json({
        success: false,
        message: "Apps Script response was not valid JSON.",
        details: responseText.slice(0, 500)
      }, 502);
    }
  } catch (error) {
    return json({
      success: false,
      message: error?.name === "TimeoutError" ? "Request timed out." : (error?.message || "GAS fetch failed.")
    }, 502);
  }
}
