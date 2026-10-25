```javascript
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

  // Validate request method.
  if (request.method !== "POST") {
    return json({
      success: false,
      stage: "request",
      message: "POST requests only."
    }, 405);
  }

  // Check upstream configuration.
  if (!gasUrl) {
    console.error("GAS_WEB_APP_URL is missing.");

    return json({
      success: false,
      stage: "configuration",
      message: "GAS_WEB_APP_URL is missing."
    }, 500);
  }

  // Parse request body.
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

  // Validate expected payload.
  if (
    !payload ||
    typeof payload.fn !== "string" ||
    !payload.fn.trim() ||
    !Array.isArray(payload.args)
  ) {
    return json({
      success: false,
      stage: "request",
      message: "Expected fn and args in request."
    }, 400);
  }

  const functionName = payload.fn.trim();

  try {
    const upstream = await fetch(gasUrl, {
      method: "POST",
      headers: {
        "Content-Type": "application/json"
      },
      body: JSON.stringify({
        fn: functionName,
        args: payload.args
      }),
      redirect: "follow",
      signal: AbortSignal.timeout(55000)
    });

    const responseText = await upstream.text();
    const elapsedMs = Date.now() - started;

    console.log(JSON.stringify({
      event: "gas_api_response",
      function: functionName,
      upstreamStatus: upstream.status,
      elapsedMs,
      responseBytes: responseText.length
    }));

    if (!upstream.ok) {
      return json({
        success: false,
        stage: "google-apps-script",
        function: functionName,
        upstreamStatus: upstream.status,
        elapsedMs,
        message:
          upstream.status === 524
            ? "Apps Script timed out. Optimize the Apps Script operation."
            : "Google Apps Script returned a non-success status.",
        details: responseText.slice(0, 1000)
      }, 502);
    }

    let parsed;

    try {
      parsed = JSON.parse(responseText);
    } catch {
      return json({
        success: false,
        stage: "response",
        function: functionName,
        elapsedMs,
        message: "Apps Script response was not valid JSON.",
        details: responseText.slice(0, 1000)
      }, 502);
    }

    return json(parsed, 200);

  } catch (error) {
    const elapsedMs = Date.now() - started;
    const timedOut = error?.name === "TimeoutError";

    console.error(JSON.stringify({
      event: "gas_api_fetch_error",
      function: functionName,
      elapsedMs,
      message: error?.message || String(error)
    }));

    return json({
      success: false,
      stage: timedOut ? "timeout" : "fetch",
      function: functionName,
      elapsedMs,
      message: timedOut
        ? "The upstream request timed out."
        : "Request to Apps Script failed."
    }, 502);
  }
}
```
