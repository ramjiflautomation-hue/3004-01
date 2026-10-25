/* Bridges google.script.run calls to the Netlify serverless proxy. */
(function () {
  "use strict";
  var endpoint = "/api/gas-api";
  var googleShim = window.google = window.google || {};
  googleShim.script = googleShim.script || {};
  var runner = {
    _success: function () {},
    _failure: function () {},
    withSuccessHandler: function (fn) { this._success = typeof fn === "function" ? fn : function () {}; return this; },
    withFailureHandler: function (fn) { this._failure = typeof fn === "function" ? fn : function () {}; return this; }
  };
  googleShim.script.run = new Proxy(runner, {
    get: function (target, property) {
      if (property in target) return target[property];
      if (typeof property !== "string" || property.charAt(0) === "_") return undefined;
      return function () {
        var args = Array.prototype.slice.call(arguments);
        var success = target._success;
        var failure = target._failure;
        target._success = function () {};
        target._failure = function () {};
        fetch(endpoint, {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({ fn: property, args: args })
        }).then(function (response) {
          if (!response.ok) throw new Error("API request failed (HTTP " + response.status + ")");
          return response.json();
        }).then(function (result) {
          if (typeof success === "function") success(result);
        }).catch(function (error) {
          if (typeof failure === "function") failure({ message: error && error.message ? error.message : String(error) });
          else console.error("Portal API error:", error);
        });
      };
    }
  });
})();
