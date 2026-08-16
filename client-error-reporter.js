(() => {
  "use strict";

  if (window.HeraClientErrorReporter?.installed) return;

  const REGION = "europe-west1";
  const FUNCTION_NAME = "reportClientError";
  const QUEUE_KEY = "hera_client_error_queue_v1";
  const DEDUPE_KEY = "hera_client_error_dedupe_v1";
  const DEDUPE_MS = 30 * 60 * 1000;
  const QUEUE_MAX = 10;
  const QUEUE_MAX_AGE_MS = 24 * 60 * 60 * 1000;
  let flushing = false;
  let authFlushBound = false;

  function truncate(value, max = 1000) {
    return String(value || "")
      .replace(/Bearer\s+[A-Za-z0-9._~+/=-]+/gi, "Bearer [REDACTED]")
      .replace(/([?&](?:token|access_token|id_token|apikey|api_key|key|password|secret)=)[^&#\s]+/gi, "$1[REDACTED]")
      .replace(/\b(AIza[0-9A-Za-z_-]{20,})\b/g, "[REDACTED_API_KEY]")
      .slice(0, max);
  }

  function safeJsonRead(key, fallback) {
    try {
      const value = JSON.parse(localStorage.getItem(key) || "null");
      return value ?? fallback;
    } catch (_) {
      return fallback;
    }
  }

  function safeJsonWrite(key, value) {
    try { localStorage.setItem(key, JSON.stringify(value)); } catch (_) {}
  }

  function hash(value) {
    let result = 2166136261;
    const text = String(value || "");
    for (let index = 0; index < text.length; index += 1) {
      result ^= text.charCodeAt(index);
      result = Math.imul(result, 16777619);
    }
    return (result >>> 0).toString(36);
  }

  function newReportId() {
    try { return crypto.randomUUID(); } catch (_) {}
    return `err-${Date.now().toString(36)}-${Math.random().toString(36).slice(2, 10)}`;
  }

  function currentView() {
    const candidates = Array.from(document.querySelectorAll("section[id], main[id], .page[id], [data-page][id]"));
    const visible = candidates.find((element) => {
      if (element.classList.contains("hidden")) return false;
      const style = window.getComputedStyle?.(element);
      return style?.display !== "none" && style?.visibility !== "hidden";
    });
    return visible?.id || document.body?.dataset?.page || "";
  }

  function appVersion() {
    return truncate(
      window.VARGA_BUILD_VERSION
      || window.HERA_BUILD_VERSION
      || document.querySelector('meta[name="app-version"]')?.content
      || "",
      100
    );
  }

  function connectionSummary() {
    const connection = navigator.connection || navigator.mozConnection || navigator.webkitConnection;
    if (!connection) return "";
    return [connection.effectiveType, Number.isFinite(connection.downlink) ? `${connection.downlink}Mbps` : "", Number.isFinite(connection.rtt) ? `${connection.rtt}ms` : ""].filter(Boolean).join(" · ");
  }

  function normalizeError(errorLike) {
    if (errorLike instanceof Error) return errorLike;
    if (errorLike && typeof errorLike === "object") {
      const error = new Error(truncate(errorLike.message || errorLike.reason || JSON.stringify(errorLike), 1200));
      if (errorLike.stack) error.stack = truncate(errorLike.stack, 7000);
      return error;
    }
    return new Error(truncate(errorLike, 1200) || "Errore senza dettagli");
  }

  function buildReport(errorLike, context = {}) {
    const error = normalizeError(errorLike);
    const kind = truncate(context.kind || error.name || "runtime-error", 80);
    const message = truncate(context.message || error.message, 1200);
    const stack = truncate(context.stack || error.stack, 7000);
    const source = truncate(context.source || "", 700);
    const fingerprint = hash(`${kind}|${message}|${stack.split("\n").slice(0, 2).join("|")}|${source}`);
    return {
      reportId: newReportId(),
      fingerprint,
      kind,
      message,
      stack,
      source,
      line: Number(context.line) || null,
      column: Number(context.column) || null,
      occurredAt: new Date().toISOString(),
      page: `${location.pathname}${location.hash || ""}`.slice(0, 300),
      activeView: currentView(),
      online: navigator.onLine !== false,
      visibility: document.visibilityState || "",
      userAgent: truncate(navigator.userAgent, 900),
      platform: truncate(navigator.userAgentData?.platform || navigator.platform || "", 160),
      language: truncate(navigator.language, 40),
      screen: `${window.innerWidth || 0}x${window.innerHeight || 0} · DPR ${window.devicePixelRatio || 1}`,
      connection: truncate(connectionSummary(), 160),
      appVersion: appVersion(),
      queuedAt: Date.now()
    };
  }

  function recentFingerprints() {
    const now = Date.now();
    const data = safeJsonRead(DEDUPE_KEY, {});
    const next = {};
    Object.entries(data && typeof data === "object" ? data : {}).forEach(([fingerprint, timestamp]) => {
      if (now - Number(timestamp) < DEDUPE_MS) next[fingerprint] = Number(timestamp);
    });
    safeJsonWrite(DEDUPE_KEY, next);
    return next;
  }

  function isDuplicate(fingerprint) {
    if (!fingerprint) return false;
    return Boolean(recentFingerprints()[fingerprint]);
  }

  function markSent(fingerprint) {
    if (!fingerprint) return;
    const data = recentFingerprints();
    data[fingerprint] = Date.now();
    const trimmed = Object.fromEntries(Object.entries(data).sort((a, b) => b[1] - a[1]).slice(0, 40));
    safeJsonWrite(DEDUPE_KEY, trimmed);
  }

  function readQueue() {
    const now = Date.now();
    const queue = safeJsonRead(QUEUE_KEY, []);
    return (Array.isArray(queue) ? queue : [])
      .filter((item) => item?.reportId && now - Number(item.queuedAt || 0) < QUEUE_MAX_AGE_MS)
      .slice(-QUEUE_MAX);
  }

  function writeQueue(queue) {
    safeJsonWrite(QUEUE_KEY, (Array.isArray(queue) ? queue : []).slice(-QUEUE_MAX));
  }

  function enqueue(report) {
    if (!report || isDuplicate(report.fingerprint)) return;
    const queue = readQueue();
    if (queue.some((item) => item.fingerprint && item.fingerprint === report.fingerprint)) return;
    queue.push(report);
    writeQueue(queue);
  }

  function callable() {
    if (!window.firebase?.apps?.length || !window.firebase?.functions) return null;
    try { return window.firebase.app().functions(REGION).httpsCallable(FUNCTION_NAME); } catch (_) { return null; }
  }

  function currentUser() {
    try { return window.firebase?.auth?.()?.currentUser || null; } catch (_) { return null; }
  }

  async function send(report) {
    if (!report || isDuplicate(report.fingerprint)) return true;
    if (navigator.onLine === false || !currentUser()) {
      enqueue(report);
      return false;
    }
    const invoke = callable();
    if (!invoke) {
      enqueue(report);
      return false;
    }
    try {
      const result = await invoke(report);
      if (result?.data?.sent || result?.data?.rateLimited || result?.data?.monthlyLimited) {
        markSent(report.fingerprint);
        return true;
      }
    } catch (_) {
      enqueue(report);
    }
    return false;
  }

  async function flushQueue() {
    if (flushing || navigator.onLine === false || !currentUser()) return;
    flushing = true;
    try {
      const queue = readQueue();
      const remaining = [];
      for (const report of queue) {
        if (isDuplicate(report.fingerprint)) continue;
        const sent = await send(report);
        if (!sent) remaining.push(report);
      }
      writeQueue(remaining);
    } finally {
      flushing = false;
    }
  }

  function report(errorLike, context = {}) {
    try {
      const item = buildReport(errorLike, context);
      if (isDuplicate(item.fingerprint)) return Promise.resolve(false);
      return send(item);
    } catch (_) {
      return Promise.resolve(false);
    }
  }

  function bindAuthFlush() {
    if (authFlushBound) return;
    try {
      const auth = window.firebase?.auth?.();
      if (!auth) return;
      authFlushBound = true;
      auth.onAuthStateChanged((user) => {
        if (user) void flushQueue();
      });
    } catch (_) {}
  }

  window.addEventListener("error", (event) => {
    try {
      const target = event.target;
      if (target && target !== window && target !== document) {
        if (!target.matches?.("script,link")) return;
        const source = target.src || target.href || "";
        if (source && new URL(source, document.baseURI).origin !== location.origin) return;
        void report(new Error(`Risorsa non caricata: ${source || target.tagName}`), { kind: "resource-error", source });
        return;
      }
      void report(event.error || new Error(event.message || "Errore JavaScript"), {
        kind: event.error?.name || "javascript-error",
        message: event.message,
        source: event.filename,
        line: event.lineno,
        column: event.colno
      });
    } catch (_) {}
  }, true);

  window.addEventListener("unhandledrejection", (event) => {
    try { void report(event.reason, { kind: "unhandled-rejection" }); } catch (_) {}
  });

  window.addEventListener("online", () => { void flushQueue(); }, { passive: true });
  document.addEventListener("visibilitychange", () => {
    if (document.visibilityState === "visible") void flushQueue();
  }, { passive: true });
  window.addEventListener("load", () => {
    bindAuthFlush();
    void flushQueue();
  }, { once: true });

  bindAuthFlush();
  window.HeraClientErrorReporter = Object.freeze({ installed: true, version: "1.1.0", report, flushQueue });
  window.setTimeout(() => { void flushQueue(); }, 0);
})();
