(function () {
  "use strict";

  var gasEndpoint = "/api/gas-api";
  var pmEndpoint = "/api/pm-checklist";
  var REQUEST_TIMEOUT_MS = 55000;

  var googleShim = window.google = window.google || {};
  googleShim.script = googleShim.script || {};

  // Give every google.script.run chain its own handlers.
  // This prevents simultaneous API calls from sharing success/failure callbacks.
  function createRunner() {
    var successHandler = function () {};
    var failureHandler = function () {};

    var runner = {
      withSuccessHandler: function (fn) {
        successHandler = typeof fn === "function" ? fn : function () {};
        return proxy;
      },
      withFailureHandler: function (fn) {
        failureHandler = typeof fn === "function" ? fn : function () {};
        return proxy;
      }
    };

    var proxy = new Proxy(runner, {
      get: function (target, property) {
        if (property in target) return target[property];
        if (typeof property !== "string" || property.charAt(0) === "_") {
          return undefined;
        }

        return function () {
          var args = Array.prototype.slice.call(arguments);
          var onSuccess = successHandler;
          var onFailure = failureHandler;
          var endpoint =
            property === "getPMChecklist" || property === "savePMChecklistRow"
              ? pmEndpoint
              : gasEndpoint;

          var controller = new AbortController();
          var timeoutId = setTimeout(function () {
            controller.abort();
          }, REQUEST_TIMEOUT_MS);

          fetch(endpoint, {
            method: "POST",
            headers: {
              "Content-Type": "application/json",
              "Accept": "application/json"
            },
            cache: "no-store",
            credentials: "same-origin",
            body: JSON.stringify({ fn: property, args: args }),
            signal: controller.signal
          })
            .then(function (response) {
              return response.text().then(function (text) {
                var result;
                try {
                  result = JSON.parse(text);
                } catch (parseError) {
                  throw new Error(
                    "API returned invalid JSON (HTTP " + response.status + "): " +
                    text.slice(0, 500)
                  );
                }

                if (!response.ok) {
                  throw new Error(
                    "API HTTP " + response.status + ": " +
                    (result && (result.message || result.details) || text)
                  );
                }
                return result;
              });
            })
            .then(function (result) {
              onSuccess(result);
            })
            .catch(function (error) {
              var message = error && error.name === "AbortError"
                ? "API request timed out after " + Math.round(REQUEST_TIMEOUT_MS / 1000) + " seconds."
                : (error && error.message ? error.message : String(error));
              onFailure({ message: message });
            })
            .finally(function () {
              clearTimeout(timeoutId);
            });
        };
      }
    });

    return proxy;
  }

  Object.defineProperty(googleShim.script, "run", {
    configurable: true,
    enumerable: true,
    get: createRunner
  });
})();
