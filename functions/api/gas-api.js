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
      stage: "configuration",
      message: "GAS_WEB_APP_URL is missing."
    }, 500);
  }

  let payload;

  try {
    payload = await request.json();
  } catch {
    return json({
      success: false,
      stage: "request",
      message: "Invalid JSON request body."
    }, 400);
  }

  if (!payload || typeof payload.fn !== "string" ||
      !Array.isArray(payload.args)) {
    return json({
      success: false,
      stage: "request",
      message: "Expected fn and args in request."
    }, 400);
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
        stage: "google-apps-script",
        function: payload.fn,
        upstreamStatus: upstream.status,
        message: "Google Apps Script returned a non-success status.",
        details: body.slice(0, 2000)
      }, 502);
    }

    try {
      JSON.parse(body);
    } catch {
      return json({
        success: false,
        stage: "response",
        function: payload.fn,
        message: "Apps Script response was not valid JSON.",
        details: body.slice(0, 2000)
      }, 502);
    }

    return new Response(body, {
      status: 200,
      headers: {
        "Content-Type": "application/json; charset=utf-8",
        "Cache-Control": "no-store"
      }
    });
  } catch (error) {
    return json({
      success: false,
      stage: "fetch",
      function: payload.fn,
      message: error?.message || "Request to Apps Script failed."
    }, 502);
  }
}
