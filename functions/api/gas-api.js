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

  // 1. Verify Method
  if (request.method !== "POST") {
    return json({
      success: false,
      stage: "request",
      message: "POST requests only."
    }, 405);
  }

  // 2. Verify GAS URL configuration in Cloudflare Pages
  if (!gasUrl) {
    return json({
      success: false,
      stage: "configuration",
      message: "GAS_WEB_APP_URL is missing in Cloudflare environment variables."
    }, 500);
  }

  // 3. Parse JSON Body
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

  // 4. Validate function name and arguments
  if (
    !payload ||
    typeof payload.fn !== "string" ||
    !payload.fn.trim() ||
    !Array.isArray(payload.args)
  ) {
    return json({
      success: false,
      stage: "request",
      message: "Expected 'fn' and 'args' in payload."
    }, 400);
  }

  const functionName = payload.fn.trim();

  // 5. Forward request to Google Apps Script
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

    if (!upstream.ok) {
      return json({
        success: false,
        stage: "google-apps-script",
        function: functionName,
        upstreamStatus: upstream.status,
        elapsedMs,
        message: upstream.status === 524
          ? "Apps Script timed out."
          : "Google Apps Script returned an error status.",
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

    return json({
      success: false,
      stage: timedOut ? "timeout" : "fetch",
      function: functionName,
      elapsedMs,
      message: timedOut
        ? "The upstream Apps Script request timed out."
        : (error?.message || "Failed to connect to Apps Script.")
    }, 502);
  }
}
