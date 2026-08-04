(function () {
  "use strict";

  const GLOBAL_NAME = "HeraFirestoreInflightReadCoalescer";
  const WRAPPED = "__heraInflightReadCoalescerWrapped";
  const ORIGINAL = "__heraInflightReadCoalescerOriginal";
  const TARGETS = new Set(["personale", "mezzi", "platformUsers", "appConfig"]);
  const RETRY_MS = 50;
  const RETRY_LIMIT = 400;

  if (window[GLOBAL_NAME]?.installed) return;

  const inFlight = new Map();
  let attempts = 0;
  let timer = null;
  const stats = {
    installedAt: "",
    installAttempts: 0,
    eligibleCalls: 0,
    networkRequestsStarted: 0,
    duplicateCallsShared: 0,
    completed: 0,
    rejected: 0,
    bypassed: 0
  };

  function canonicalPath(value) {
    if (!value) return "";
    if (typeof value === "string") return value;
    if (typeof value.canonicalString === "function") {
      try { return value.canonicalString(); } catch (_) {}
    }
    if (typeof value.toArray === "function") {
      try {
        const parts = value.toArray();
        if (Array.isArray(parts)) return parts.join("/");
      } catch (_) {}
    }
    if (Array.isArray(value.segments)) return value.segments.join("/");
    if (Array.isArray(value._segments)) return value._segments.join("/");
    return "";
  }

  function queryPath(query) {
    const values = [
      query?.path,
      query?._query?.path,
      query?._query?._path,
      query?.Ae?.path,
      query?.je?.path,
      query?._delegate?._query?.path,
      query?._delegate?._query?._path
    ];
    for (const value of values) {
      const path = canonicalPath(value).replace(/^\/+|\/+$/g, "");
      if (path) return path;
    }
    return "";
  }

  function canonicalQuery(query) {
    const candidates = [query?._query, query?.Ae, query?.je, query?._delegate?._query].filter(Boolean);
    for (const internal of candidates) {
      for (const method of ["canonicalId", "canonicalString"]) {
        if (typeof internal?.[method] !== "function") continue;
        try {
          const value = String(internal[method]() || "");
          if (value) return value;
        } catch (_) {}
      }
    }
    return "";
  }

  function stableValue(value) {
    if (value == null) return "";
    if (typeof value !== "object") return String(value);
    const selected = {};
    ["source", "timeoutMs", "retries"].forEach((key) => {
      if (Object.prototype.hasOwnProperty.call(value, key)) selected[key] = value[key];
    });
    return JSON.stringify(selected);
  }

  function requestInfo(query, options) {
    const path = queryPath(query);
    const target = path.split("/")[0] || "";
    if (!TARGETS.has(target)) return null;
    const canonical = canonicalQuery(query);
    if (!canonical && !query?.path) return null;
    return {
      target,
      key: `${path}|${canonical || `collection:${path}`}|${stableValue(options)}`
    };
  }

  function install() {
    attempts += 1;
    stats.installAttempts = attempts;
    const current = window.runFirestoreGetWithRetry;

    if (typeof current !== "function") {
      if (attempts < RETRY_LIMIT && !timer) {
        timer = setTimeout(() => {
          timer = null;
          install();
        }, RETRY_MS);
      }
      return false;
    }

    if (current[WRAPPED]) {
      api.installed = true;
      return true;
    }

    const original = current[ORIGINAL] || current;
    const wrapped = function coalescedFirestoreGet(query, options) {
      const info = requestInfo(query, options);
      if (!info) {
        stats.bypassed += 1;
        return original.apply(this, arguments);
      }

      stats.eligibleCalls += 1;
      const pending = inFlight.get(info.key);
      if (pending) {
        stats.duplicateCallsShared += 1;
        return pending;
      }

      stats.networkRequestsStarted += 1;
      let request;
      try {
        request = Promise.resolve(original.apply(this, arguments));
      } catch (error) {
        stats.rejected += 1;
        throw error;
      }

      const tracked = request.then(
        (snapshot) => {
          stats.completed += 1;
          return snapshot;
        },
        (error) => {
          stats.rejected += 1;
          throw error;
        }
      ).finally(() => {
        if (inFlight.get(info.key) === tracked) inFlight.delete(info.key);
      });

      inFlight.set(info.key, tracked);
      return tracked;
    };

    Object.defineProperty(wrapped, WRAPPED, { value: true });
    Object.defineProperty(wrapped, ORIGINAL, { value: original });
    window.runFirestoreGetWithRetry = wrapped;
    stats.installedAt = new Date().toISOString();
    api.installed = true;
    return true;
  }

  const api = {
    installed: false,
    targets: Object.freeze(Array.from(TARGETS)),
    stats,
    refreshInstallation: install,
    getState() {
      return {
        installed: api.installed,
        targets: Array.from(TARGETS),
        inFlight: inFlight.size,
        stats: { ...stats }
      };
    }
  };

  window[GLOBAL_NAME] = api;
  install();
  window.addEventListener?.("load", install, { once: true });
})();
