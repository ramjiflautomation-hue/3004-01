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

  if (
    !payload ||
    typeof payload.fn !== "string" ||
    !Array.isArray(payload.args)
  ) {
    return json({
      success: false,
      stage: "request",
      message: "Expected fn and args in request."
    }, 400);
  }

  const functionName = payload.fn;

  try {
    const upstream = await fetch(gasUrl, {
      method: "POST",
      headers: {
        "Content-Type": "application/json"
      },
      body: JSON.stringify(payload),
      redirect: "follow"
    });

    const body = await upstream.text();
    const elapsedMs = Date.now() - started;

    console.log(JSON.stringify({
      event: "gas_api_response",
      function: functionName,
      upstreamStatus: upstream.status,
      elapsedMs: elapsedMs,
      responseBytes: body.length
    }));

    if (!upstream.ok) {
      return json({
        success: false,
        stage: "google-apps-script",
        function: functionName,
        upstreamStatus: upstream.status,
        elapsedMs: elapsedMs,
        message:
          upstream.status === 524
            ? "Apps Script timed out. Optimize the Apps Script operation."
            : "Google Apps Script returned a non-success status.",
        details: body.slice(0, 1500)
      }, 502);
    }

    let parsed;

    try {
      parsed = JSON.parse(body);
    } catch {
      return json({
        success: false,
        stage: "response",
        function: functionName,
        elapsedMs: elapsedMs,
        message: "Apps Script response was not valid JSON.",
        details: body.slice(0, 1500)
      }, 502);
    }

    return json(parsed, 200);

  } catch (error) {
    const elapsedMs = Date.now() - started;

    console.error(JSON.stringify({
      event: "gas_api_fetch_error",
      function: functionName,
      elapsedMs: elapsedMs,
      message: error?.message || String(error)
    }));

    return json({
      success: false,
      stage: "fetch",
      function: functionName,
      elapsedMs: elapsedMs,
      message: error?.message || "Request to Apps Script failed."
    }, 502);
  }
}
