function json(body, status = 200) {
  return new Response(JSON.stringify(body), {
    status,
    headers: {
      "Content-Type": "application/json; charset=utf-8",
      "Cache-Control": "no-store"
    }
  });
}

export async function onRequestOptions() {
  return new Response(null, { status: 204 });
}

export async function onRequestPost({ request, env }) {
  const gasUrl = env.GAS_WEB_APP_URL;
  if (!gasUrl) {
    return json({
      success: false,
      message: "GAS_WEB_APP_URL is not configured in Cloudflare Pages environment variables."
    }, 500);
  }

  let payload;
  try {
    payload = await request.json();
  } catch {
    return json({ success: false, message: "Request body must be valid JSON." }, 400);
  }

  if (!payload || typeof payload.fn !== "string" || !Array.isArray(payload.args)) {
    return json({ success: false, message: "Invalid API payload." }, 400);
  }

  try {
    const upstream = await fetch(gasUrl, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(payload),
      redirect: "follow"
    });
    const body = await upstream.text();
    if (!upstream.ok) {
      return json({
        success: false,
        message: "Google Apps Script returned HTTP " + upstream.status,
        details: body.slice(0, 500)
      }, 502);
    }
    return new Response(body || "null", {
      status: 200,
      headers: {
        "Content-Type": "application/json; charset=utf-8",
        "Cache-Control": "no-store"
      }
    });
  } catch (error) {
    return json({
      success: false,
      message: error && error.message ? error.message : "Upstream API request failed."
    }, 502);
  }
}
