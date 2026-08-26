(() => {
  "use strict";

  if (window.HeraAppErrorMonitor?.installed) return;

  const VERSION = "1.1.1";
  const REGION = "europe-west1";
  const FUNCTION_NAME = "recordClientErrorGroup";
  const QUEUE_KEY = "hera_error_center_queue_v1";
  const DEDUPE_KEY = "hera_error_center_dedupe_v1";
  const HEALTH_KEY = "hera_error_center_health_v1";
  const QUEUE_MAX = 30;
  const QUEUE_MAX_AGE_MS = 7 * 24 * 60 * 60 * 1000;
  const DEDUPE_MS = 10 * 60 * 1000;
  const MANUAL_DEDUPE_MS = 60 * 1000;
  const BREADCRUMB_MAX = 15;
  const LONG_TASK_MIN_MS = 1000;
  const SLOW_INTERACTION_MIN_MS = 900;
  const REPEATED_TAP_WINDOW_MS = 1800;
  const SENSITIVE_KEY = /password|passcode|pin|token|secret|cookie|authorization|credential|api.?key|session|jwt/i;

  const breadcrumbs = [];
  let flushing = false;
  let authBound = false;
  let longTaskBound = false;
  let consoleErrorBound = false;
  let lastTap = { key: "", at: 0, count: 0 };

  function redactText(value, max = 1800) {
    return String(value ?? "")
      .replace(/Bearer\s+[A-Za-z0-9._~+/=-]+/gi, "Bearer [RIMOSSO]")
      .replace(/([?&](?:token|access_token|id_token|apikey|api_key|key|password|secret|pin)=)[^&#\s]+/gi, "$1[RIMOSSO]")
      .replace(/\b(AIza[0-9A-Za-z_-]{20,})\b/g, "[CHIAVE_API_RIMOSSA]")
      .replace(/\beyJ[A-Za-z0-9_-]{20,}\.[A-Za-z0-9_-]{10,}\.[A-Za-z0-9_-]{10,}\b/g, "[TOKEN_RIMOSSO]")
      .replace(/\b([A-Z0-9._%+-]{2})[A-Z0-9._%+-]*(@[A-Z0-9.-]+\.[A-Z]{2,})\b/gi, "$1***$2")
      .replace(/\b(password|passcode|pin|secret|token)\s*[:=]\s*[^\s,;]+/gi, "$1=[RIMOSSO]")
      .slice(0, max);
  }

  function sanitizeValue(value, depth = 0) {
    if (depth > 3 || value == null) return value == null ? null : "[DATI_RIDOTTI]";
    if (typeof value === "string") return redactText(value, 1000);
    if (typeof value === "number" || typeof value === "boolean") return value;
    if (value instanceof Date) return value.toISOString();
    if (Array.isArray(value)) return value.slice(0, 15).map((item) => sanitizeValue(item, depth + 1));
    if (typeof value === "object") {
      const output = {};
      Object.entries(value).slice(0, 30).forEach(([key, item]) => {
        if (SENSITIVE_KEY.test(key)) output[key] = "[RIMOSSO]";
        else output[redactText(key, 80)] = sanitizeValue(item, depth + 1);
      });
      return output;
    }
    return redactText(value, 300);
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

  function updateHealth(patch = {}) {
    const previous = safeJsonRead(HEALTH_KEY, {});
    safeJsonWrite(HEALTH_KEY, {
      ...(previous && typeof previous === "object" ? previous : {}),
      ...sanitizeValue(patch),
      updatedAt: new Date().toISOString()
    });
  }

  function healthSnapshot() {
    const saved = safeJsonRead(HEALTH_KEY, {});
    return {
      monitorInstalled: true,
      monitorVersion: VERSION,
      authenticated: Boolean(currentUser()),
      online: navigator.onLine !== false,
      queuedReports: readQueue().length,
      lastSuccessfulSendAt: saved?.lastSuccessfulSendAt || "",
      lastFailureAt: saved?.lastFailureAt || "",
      lastFailure: redactText(saved?.lastFailure || "", 500)
    };
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

  function newId() {
    try { return crypto.randomUUID(); } catch (_) {}
    return `bug-${Date.now().toString(36)}-${Math.random().toString(36).slice(2, 9)}`;
  }

  function currentView() {
    try {
      const candidates = Array.from(document.querySelectorAll("section[id], main[id], .page[id], [data-page][id]"));
      const visible = candidates.find((element) => {
        if (element.hidden || element.classList?.contains("hidden") || element.getAttribute?.("aria-hidden") === "true") return false;
        const style = window.getComputedStyle?.(element);
        return style?.display !== "none" && style?.visibility !== "hidden";
      });
      return redactText(visible?.id || document.body?.dataset?.page || "", 120);
    } catch (_) {
      return "";
    }
  }

  function selectedContext() {
    const readGlobal = (name) => {
      try { return window[name]; } catch (_) { return null; }
    };
    const commessaId = readGlobal("selectedCommessaId") || readGlobal("currentCommessaId") || "";
    const impiantoId = readGlobal("selectedImpiantoId") || readGlobal("currentImpiantoId") || "";
    let commessaName = "";
    try {
      commessaName = document.getElementById("impianti-page-title")?.textContent
        || document.getElementById("commessa-focus-label")?.textContent
        || "";
    } catch (_) {}
    return {
      commessaId: redactText(commessaId, 160),
      commessaName: redactText(commessaName, 180),
      impiantoId: redactText(impiantoId, 160)
    };
  }

  function appVersion() {
    try {
      return redactText(
        window.VARGA_BUILD_VERSION
        || window.HERA_BUILD_VERSION
        || document.querySelector('meta[name="app-version"]')?.content
        || "",
        100
      );
    } catch (_) {
      return "";
    }
  }

  function platformKind() {
    try {
      if (window.Capacitor?.isNativePlatform?.()) return `native-${window.Capacitor.getPlatform?.() || "mobile"}`;
    } catch (_) {}
    if (/iPad|iPhone|iPod/i.test(navigator.userAgent || "")) return "pwa-ios";
    if (/Android/i.test(navigator.userAgent || "")) return "pwa-android";
    return "web";
  }

  function connectionSummary() {
    const connection = navigator.connection || navigator.mozConnection || navigator.webkitConnection;
    if (!connection) return "";
    return [
      connection.effectiveType,
      Number.isFinite(connection.downlink) ? `${connection.downlink}Mbps` : "",
      Number.isFinite(connection.rtt) ? `${connection.rtt}ms` : ""
    ].filter(Boolean).join(" · ");
  }

  function normalizeError(errorLike) {
    if (errorLike instanceof Error) return errorLike;
    if (errorLike && typeof errorLike === "object") {
      let raw = "";
      try { raw = errorLike.message || errorLike.reason || JSON.stringify(errorLike); } catch (_) { raw = String(errorLike); }
      const error = new Error(redactText(raw, 1600));
      if (errorLike.stack) error.stack = redactText(errorLike.stack, 7000);
      return error;
    }
    return new Error(redactText(errorLike, 1600) || "Errore senza dettagli");
  }

  function severityFor(kind, context = {}) {
    const requested = String(context.severity || "").toLowerCase();
    if (["critical", "high", "medium", "low", "info"].includes(requested)) return requested;
    const duration = Number(context.durationMs || 0);
    if (kind === "ui-freeze" && duration >= 5000) return "critical";
    if (duration >= 3000) return "high";
    if (/javascript|unhandled|resource-error|repeated-tap/.test(kind)) return "high";
    if (/long-task|slow-interaction|network/.test(kind)) return "medium";
    return "medium";
  }

  function inferFeature(context = {}) {
    if (context.feature) return redactText(context.feature, 120);
    const view = currentView();
    if (/impiant/i.test(view)) return "impianti";
    if (/squadr/i.test(view)) return "squadre";
    if (/ore|hours/i.test(view)) return "gestione-ore";
    if (/map|mappa/i.test(view)) return "mappa";
    if (/login|auth/i.test(view)) return "accesso";
    return view || "app";
  }

  function addBreadcrumb(type, label, metadata = {}) {
    breadcrumbs.push({
      at: new Date().toISOString(),
      type: redactText(type, 50),
      label: redactText(label, 180),
      metadata: sanitizeValue(metadata)
    });
    if (breadcrumbs.length > BREADCRUMB_MAX) breadcrumbs.splice(0, breadcrumbs.length - BREADCRUMB_MAX);
  }

  function buildReport(errorLike, context = {}) {
    const error = normalizeError(errorLike);
    const kind = redactText(context.kind || error.name || "runtime-error", 80);
    const message = redactText(context.message || error.message, 1600);
    const stack = redactText(context.stack || error.stack || "", 7000);
    const feature = inferFeature(context);
    const source = redactText(context.source || "", 700);
    const selected = selectedContext();
    const fingerprint = redactText(
      context.fingerprint || hash(`${kind}|${feature}|${message}|${stack.split("\n").slice(0, 2).join("|")}|${source}`),
      120
    );
    return {
      reportId: newId(),
      fingerprint,
      kind,
      severity: severityFor(kind, context),
      feature,
      message,
      stack,
      source,
      line: Number(context.line) || null,
      column: Number(context.column) || null,
      durationMs: Math.max(0, Math.round(Number(context.durationMs || 0))),
      tapCount: Math.max(0, Math.round(Number(context.tapCount || 0))),
      manual: Boolean(context.manual),
      occurredAt: new Date().toISOString(),
      page: redactText(`${location.pathname || ""}${location.hash || ""}`, 300),
      activeView: currentView(),
      online: navigator.onLine !== false,
      visibility: redactText(document.visibilityState || "", 40),
      userAgent: redactText(navigator.userAgent || "", 900),
      platform: platformKind(),
      language: redactText(navigator.language || "", 40),
      screen: `${window.innerWidth || 0}x${window.innerHeight || 0} · DPR ${window.devicePixelRatio || 1}`,
      connection: redactText(connectionSummary(), 160),
      appVersion: appVersion(),
      commessaId: selected.commessaId,
      commessaName: selected.commessaName,
      impiantoId: selected.impiantoId,
      metadata: sanitizeValue(context.metadata || {}),
      breadcrumbs: breadcrumbs.slice(-BREADCRUMB_MAX).map((item) => sanitizeValue(item)),
      queuedAt: Date.now()
    };
  }

  function recentFingerprints() {
    const now = Date.now();
    const data = safeJsonRead(DEDUPE_KEY, {});
    const next = {};
    Object.entries(data && typeof data === "object" ? data : {}).forEach(([fingerprint, entry]) => {
      const timestamp = typeof entry === "object" ? Number(entry.at) : Number(entry);
      const ttl = typeof entry === "object" && entry.manual ? MANUAL_DEDUPE_MS : DEDUPE_MS;
      if (timestamp && now - timestamp < ttl) next[fingerprint] = typeof entry === "object" ? entry : { at: timestamp, manual: false };
    });
    safeJsonWrite(DEDUPE_KEY, next);
    return next;
  }

  function isDuplicate(report) {
    if (!report?.fingerprint) return false;
    return Boolean(recentFingerprints()[report.fingerprint]);
  }

  function markSent(report) {
    if (!report?.fingerprint) return;
    const data = recentFingerprints();
    data[report.fingerprint] = { at: Date.now(), manual: Boolean(report.manual) };
    const trimmed = Object.fromEntries(
      Object.entries(data)
        .sort((a, b) => Number(b[1]?.at || 0) - Number(a[1]?.at || 0))
        .slice(0, 80)
    );
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
    if (!report || (!report.manual && isDuplicate(report))) return;
    const queue = readQueue();
    if (queue.some((item) => item.reportId === report.reportId)) return;
    queue.push(report);
    writeQueue(queue);
  }

  function currentUser() {
    try { return window.firebase?.auth?.()?.currentUser || null; } catch (_) { return null; }
  }

  function callable() {
    if (!window.firebase?.apps?.length || !window.firebase?.functions) return null;
    try { return window.firebase.app().functions(REGION).httpsCallable(FUNCTION_NAME); } catch (_) { return null; }
  }

  async function send(report) {
    if (!report) return false;
    if (!report.manual && isDuplicate(report)) return true;
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
      const response = await invoke(report);
      if (response?.data?.recorded || response?.data?.rateLimited) {
        markSent(report);
        updateHealth({
          lastSuccessfulSendAt: new Date().toISOString(),
          lastFailureAt: "",
          lastFailure: ""
        });
        return true;
      }
      updateHealth({
        lastFailureAt: new Date().toISOString(),
        lastFailure: "Risposta non valida dal servizio di registrazione errori."
      });
    } catch (error) {
      updateHealth({
        lastFailureAt: new Date().toISOString(),
        lastFailure: error?.message || error?.code || "Invio diagnostica non riuscito."
      });
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
        if (!report.manual && isDuplicate(report)) continue;
        const sent = await send(report);
        if (!sent) remaining.push(report);
      }
      writeQueue(remaining);
    } finally {
      flushing = false;
    }
  }

  function capture(errorLike, context = {}) {
    try {
      const report = buildReport(errorLike, context);
      if (!report.manual && isDuplicate(report)) return Promise.resolve({ sent: false, duplicate: true, queued: false });
      return send(report).then((sent) => ({ sent, queued: !sent, duplicate: false, reportId: report.reportId }));
    } catch (_) {
      return Promise.resolve({ sent: false, queued: false, duplicate: false });
    }
  }

  function reportManual(input = {}) {
    const title = redactText(input.title || "Problema segnalato dall'utente", 180);
    const description = redactText(input.description || "", 3000);
    const expected = redactText(input.expected || "", 1800);
    const steps = redactText(input.steps || "", 2200);
    return capture(new Error(title), {
      kind: "manual-bug-report",
      severity: input.severity || "medium",
      feature: input.feature || currentView() || "app",
      manual: true,
      message: title,
      source: "segnalazione-manuale",
      metadata: { description, expected, steps }
    });
  }

  function targetLabel(target) {
    if (!target || typeof target.closest !== "function") return { key: "elemento", label: "Elemento" };
    const actionable = target.closest("button, a, [role='button'], input, select, textarea, [data-action]") || target;
    const tag = String(actionable.tagName || "elemento").toLowerCase();
    const id = redactText(actionable.id || actionable.dataset?.action || actionable.getAttribute?.("data-action-key") || "", 100);
    let label = "";
    if (/^(input|select|textarea)$/.test(tag)) {
      label = actionable.getAttribute?.("aria-label") || actionable.name || actionable.id || `${tag} ${actionable.type || ""}`;
    } else {
      label = actionable.getAttribute?.("aria-label") || actionable.title || actionable.textContent || id || tag;
    }
    label = redactText(label, 120).replace(/\s+/g, " ").trim() || tag;
    return { key: `${tag}|${id}|${label}`, label };
  }

  function monitorInteraction(event) {
    if (event.isTrusted === false) return;
    const target = event.target;
    const info = targetLabel(target);
    const now = Date.now();
    if (lastTap.key === info.key && now - lastTap.at <= REPEATED_TAP_WINDOW_MS) lastTap.count += 1;
    else lastTap = { key: info.key, at: now, count: 1 };
    lastTap.at = now;

    addBreadcrumb("tap", info.label, { view: currentView(), count: lastTap.count });

    if (lastTap.count === 3) {
      void capture(new Error(`Il comando ${info.label} ha richiesto più tocchi`), {
        kind: "repeated-tap",
        severity: "high",
        feature: currentView() || "interfaccia",
        tapCount: lastTap.count,
        metadata: { control: info.label }
      });
    }

    const startedAt = performance.now?.() || Date.now();
    window.setTimeout(() => {
      if (document.visibilityState === "hidden") return;
      const endedAt = performance.now?.() || Date.now();
      const delay = Math.max(0, endedAt - startedAt);
      if (delay < SLOW_INTERACTION_MIN_MS) return;
      void capture(new Error(`Risposta lenta dopo il comando ${info.label}`), {
        kind: delay >= 5000 ? "ui-freeze" : "slow-interaction",
        severity: delay >= 5000 ? "critical" : delay >= 2500 ? "high" : "medium",
        feature: currentView() || "interfaccia",
        durationMs: delay,
        metadata: { control: info.label }
      });
    }, 0);
  }

  function installLongTaskObserver() {
    try {
      if (longTaskBound) return;
      if (!window.PerformanceObserver?.supportedEntryTypes?.includes("longtask")) return;
      longTaskBound = true;
      const observer = new PerformanceObserver((list) => {
        if (document.visibilityState === "hidden") return;
        const entries = list.getEntries()
          .filter((entry) => Number(entry.duration || 0) >= LONG_TASK_MIN_MS);
        if (!entries.length) return;

        const longest = entries.reduce((worst, entry) =>
          Number(entry.duration || 0) > Number(worst.duration || 0) ? entry : worst
        );
        const longestDuration = Number(longest.duration || 0);
        const totalDuration = entries.reduce((sum, entry) => sum + Number(entry.duration || 0), 0);

        void capture(new Error("Operazione lunga sul thread principale"), {
          kind: longestDuration >= 5000 ? "ui-freeze" : "main-thread-long-task",
          severity: longestDuration >= 5000 ? "critical" : longestDuration >= 2500 ? "high" : "medium",
          feature: currentView() || "app",
          durationMs: longestDuration,
          metadata: {
            entryType: longest.entryType,
            name: longest.name || "longtask",
            taskCount: entries.length,
            totalDurationMs: Math.round(totalDuration),
            longestStartTimeMs: Math.round(Number(longest.startTime || 0)),
            navigationAgeMs: Math.round(performance.now?.() || 0)
          }
        });
      });
      observer.observe({ type: "longtask", buffered: false });
    } catch (_) {}
  }

  function installConsoleErrorCapture() {
    if (consoleErrorBound || !window.console?.error) return;
    try {
      consoleErrorBound = true;
      const original = window.console.error.bind(window.console);
      const wrapped = (...args) => {
        original(...args);
        try {
          const firstError = args.find((item) => item instanceof Error);
          const message = args.map((item) => {
            if (item instanceof Error) return item.message;
            if (typeof item === "string") return item;
            if (item && typeof item === "object") {
              return [item.name, item.code, item.message].filter(Boolean).map((value) => redactText(value, 500)).join(" · ");
            }
            return redactText(item, 300);
          }).filter(Boolean).join(" ");
          if (!message || /Push Centro errori non inviata/i.test(message)) return;
          void capture(firstError || new Error(message), {
            kind: "handled-console-error",
            severity: "high",
            message,
            source: "console.error"
          });
        } catch (_) {}
      };
      wrapped.__heraErrorMonitor = true;
      window.console.error = wrapped;
    } catch (_) {
      consoleErrorBound = false;
    }
  }

  function bindAuthFlush() {
    if (authBound) return;
    try {
      const auth = window.firebase?.auth?.();
      if (!auth) return;
      authBound = true;
      auth.onAuthStateChanged((user) => {
        if (user) void flushQueue();
      });
    } catch (_) {}
  }

  window.addEventListener("error", (event) => {
    try {
      const target = event.target;
      if (target && target !== window && target !== document) {
        const tag = String(target.tagName || "").toUpperCase();
        if (!/^(SCRIPT|LINK|IMG|SOURCE)$/.test(tag)) return;
        const source = target.src || target.href || "";
        void capture(new Error(`Risorsa non caricata: ${source || tag}`), {
          kind: "resource-error",
          severity: "high",
          source
        });
        return;
      }
      void capture(event.error || new Error(event.message || "Errore JavaScript"), {
        kind: event.error?.name || "javascript-error",
        severity: "high",
        message: event.message,
        source: event.filename,
        line: event.lineno,
        column: event.colno
      });
    } catch (_) {}
  }, true);

  window.addEventListener("unhandledrejection", (event) => {
    try { void capture(event.reason, { kind: "unhandled-rejection", severity: "high" }); } catch (_) {}
  });

  document.addEventListener("pointerup", monitorInteraction, { capture: true, passive: true });
  window.addEventListener("hashchange", () => addBreadcrumb("navigation", location.hash || "home"), { passive: true });
  window.addEventListener("popstate", () => addBreadcrumb("navigation", `${location.pathname}${location.hash || ""}`), { passive: true });
  window.addEventListener("online", () => {
    addBreadcrumb("network", "Connessione ripristinata");
    void flushQueue();
  }, { passive: true });
  window.addEventListener("offline", () => addBreadcrumb("network", "Dispositivo offline"), { passive: true });
  document.addEventListener("visibilitychange", () => {
    if (document.visibilityState === "visible") void flushQueue();
  }, { passive: true });
  window.addEventListener("load", () => {
    bindAuthFlush();
    installLongTaskObserver();
    void flushQueue();
  }, { once: true });

  bindAuthFlush();
  installLongTaskObserver();
  installConsoleErrorCapture();
  addBreadcrumb("session", "Monitor errori avviato", { platform: platformKind(), version: appVersion() });

  window.HeraAppErrorMonitor = Object.freeze({
    installed: true,
    version: VERSION,
    capture,
    reportManual,
    flushQueue,
    breadcrumb: addBreadcrumb,
    getQueueLength: () => readQueue().length,
    getHealth: healthSnapshot
  });

  window.setTimeout(() => { void flushQueue(); }, 0);
})();
