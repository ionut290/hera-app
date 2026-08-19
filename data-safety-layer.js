(() => {
  "use strict";

  if (window.HeraDataSafety?.installed) return;

  const VERSION = "1.0.0";
  const DEFAULT_SCHEMA_VERSION = 1;
  const JOURNAL_KEY = "hera_data_safety_journal_v1";
  const COMPLETED_KEY = "hera_data_safety_completed_v1";
  const MAX_JOURNAL_ENTRIES = 120;
  const MAX_COMPLETED_ENTRIES = 250;
  const JOURNAL_TTL_MS = 24 * 60 * 60 * 1000;
  const COMPLETED_TTL_MS = 7 * 24 * 60 * 60 * 1000;
  const MAX_PAYLOAD_BYTES = 1_500_000;
  const TRANSIENT_CODES = new Set([
    "aborted",
    "deadline-exceeded",
    "internal",
    "network-request-failed",
    "resource-exhausted",
    "unavailable",
    "unknown"
  ]);

  const inflight = new Map();
  const schemas = new Map();
  let originalOfflineEnqueue = null;
  let originalOfflineSync = null;

  function now() {
    return Date.now();
  }

  function createOperationId(prefix = "op") {
    try {
      if (crypto?.randomUUID) return `${prefix}:${crypto.randomUUID()}`;
    } catch (_) {}
    return `${prefix}:${now().toString(36)}:${Math.random().toString(36).slice(2, 12)}`;
  }

  function safeJsonParse(value, fallback) {
    try {
      const parsed = JSON.parse(String(value || ""));
      return parsed ?? fallback;
    } catch (_) {
      return fallback;
    }
  }

  function safeClone(value) {
    try {
      if (typeof structuredClone === "function") return structuredClone(value);
    } catch (_) {}
    try {
      return JSON.parse(JSON.stringify(value));
    } catch (_) {
      return value;
    }
  }

  function payloadBytes(payload) {
    try {
      return new Blob([JSON.stringify(payload)]).size;
    } catch (_) {
      try { return JSON.stringify(payload).length * 2; }
      catch (_) { return MAX_PAYLOAD_BYTES + 1; }
    }
  }

  function normalizeCode(error) {
    const raw = String(error?.code || "").trim().toLowerCase();
    return raw.replace(/^firestore\//, "").replace(/^auth\//, "");
  }

  function errorMessage(error) {
    return String(error?.message || error || "Errore sconosciuto").slice(0, 500);
  }

  function isTransientError(error) {
    const code = normalizeCode(error);
    if (TRANSIENT_CODES.has(code)) return true;
    const text = `${code} ${errorMessage(error)}`.toLowerCase();
    return /(network|offline|internet|timeout|timed out|temporar|unavailable|connection|connessione|failed to fetch)/i.test(text);
  }

  function dispatch(name, detail) {
    try {
      window.dispatchEvent(new CustomEvent(name, { detail }));
    } catch (_) {}
  }

  function readArray(key) {
    try {
      const parsed = safeJsonParse(localStorage.getItem(key), []);
      return Array.isArray(parsed) ? parsed : [];
    } catch (_) {
      return [];
    }
  }

  function writeArray(key, rows) {
    try {
      localStorage.setItem(key, JSON.stringify(rows));
      return true;
    } catch (_) {
      return false;
    }
  }

  function pruneRows(rows, ttlMs, maxEntries) {
    const cutoff = now() - ttlMs;
    return (Array.isArray(rows) ? rows : [])
      .filter((row) => Number(row?.updatedAt || row?.completedAt || row?.createdAt || 0) >= cutoff)
      .sort((a, b) => Number(b?.updatedAt || b?.completedAt || b?.createdAt || 0) - Number(a?.updatedAt || a?.completedAt || a?.createdAt || 0))
      .slice(0, maxEntries);
  }

  function journalOperation(operation, status, extra = {}) {
    const rows = pruneRows(readArray(JOURNAL_KEY), JOURNAL_TTL_MS, MAX_JOURNAL_ENTRIES);
    const filtered = rows.filter((row) => row?.operationId !== operation.operationId);
    filtered.unshift({
      operationId: operation.operationId,
      type: operation.type,
      schemaVersion: operation.schemaVersion,
      status: String(status || "pending"),
      createdAt: operation.createdAt,
      updatedAt: now(),
      payloadBytes: Number(operation.payloadBytes || 0),
      queueKind: String(extra.queueKind || operation.meta?.queueKind || ""),
      dedupeKey: String(extra.dedupeKey || operation.meta?.dedupeKey || "").slice(0, 180),
      errorCode: String(extra.errorCode || "").slice(0, 80),
      errorMessage: String(extra.errorMessage || "").slice(0, 240)
    });
    writeArray(JOURNAL_KEY, pruneRows(filtered, JOURNAL_TTL_MS, MAX_JOURNAL_ENTRIES));
  }

  function removeJournalOperation(operationId) {
    const rows = pruneRows(readArray(JOURNAL_KEY), JOURNAL_TTL_MS, MAX_JOURNAL_ENTRIES)
      .filter((row) => row?.operationId !== operationId);
    writeArray(JOURNAL_KEY, rows);
  }

  function readCompleted() {
    return pruneRows(readArray(COMPLETED_KEY), COMPLETED_TTL_MS, MAX_COMPLETED_ENTRIES);
  }

  function isCompleted(operationId) {
    if (!operationId) return false;
    return readCompleted().some((row) => row?.operationId === operationId);
  }

  function markCompleted(operation) {
    const rows = readCompleted().filter((row) => row?.operationId !== operation.operationId);
    rows.unshift({
      operationId: operation.operationId,
      type: operation.type,
      completedAt: now(),
      updatedAt: now()
    });
    writeArray(COMPLETED_KEY, pruneRows(rows, COMPLETED_TTL_MS, MAX_COMPLETED_ENTRIES));
    removeJournalOperation(operation.operationId);
  }

  async function snapshot(reason) {
    try {
      if (window.HeraDataDurability?.snapshot) {
        return await window.HeraDataDurability.snapshot(String(reason || "data-safety"));
      }
    } catch (error) {
      console.warn("[DATA SAFETY] Backup locale non completato:", error);
    }
    return null;
  }

  function baseValidate(type, payload) {
    const errors = [];
    const safeType = String(type || "").trim();
    if (!safeType) errors.push("Tipo operazione mancante");
    if (safeType.length > 100) errors.push("Tipo operazione troppo lungo");
    if (payload === undefined) errors.push("Dati operazione mancanti");

    const bytes = payloadBytes(payload);
    if (bytes > MAX_PAYLOAD_BYTES) errors.push("Dati operazione troppo grandi");

    if (payload && typeof payload === "object" && !Array.isArray(payload)) {
      const latitude = payload.latitude ?? payload.lat ?? payload.gpsLat;
      const longitude = payload.longitude ?? payload.lng ?? payload.lon ?? payload.gpsLng;
      if (latitude !== undefined && latitude !== null && latitude !== "") {
        const value = Number(latitude);
        if (!Number.isFinite(value) || value < -90 || value > 90) errors.push("Latitudine non valida");
      }
      if (longitude !== undefined && longitude !== null && longitude !== "") {
        const value = Number(longitude);
        if (!Number.isFinite(value) || value < -180 || value > 180) errors.push("Longitudine non valida");
      }
      if (payload.date !== undefined && payload.date !== null && payload.date !== "") {
        if (!/^\d{4}-\d{2}-\d{2}$/.test(String(payload.date))) errors.push("Data non valida: usare YYYY-MM-DD");
      }
    }

    return { ok: errors.length === 0, errors, bytes };
  }

  function registerSchema(type, config = {}) {
    const key = String(type || "").trim();
    if (!key) throw new Error("Nome schema mancante");
    const version = Math.max(1, Number(config.version || DEFAULT_SCHEMA_VERSION));
    schemas.set(key, {
      version,
      validate: typeof config.validate === "function" ? config.validate : null,
      migrate: typeof config.migrate === "function" ? config.migrate : null
    });
    return version;
  }

  function validate(type, payload, options = {}) {
    const base = baseValidate(type, payload);
    const schema = schemas.get(String(type || "").trim());
    if (!schema?.validate) return { ...base, schemaVersion: schema?.version || Number(options.schemaVersion || DEFAULT_SCHEMA_VERSION) };

    try {
      const custom = schema.validate(payload, options);
      const customErrors = Array.isArray(custom)
        ? custom.map((item) => String(item))
        : custom === false
          ? ["Validazione schema non superata"]
          : [];
      const errors = [...base.errors, ...customErrors];
      return { ok: errors.length === 0, errors, bytes: base.bytes, schemaVersion: schema.version };
    } catch (error) {
      const errors = [...base.errors, `Errore validazione schema: ${errorMessage(error)}`];
      return { ok: false, errors, bytes: base.bytes, schemaVersion: schema.version };
    }
  }

  function migrate(type, payload, fromVersion = DEFAULT_SCHEMA_VERSION) {
    const schema = schemas.get(String(type || "").trim());
    if (!schema || Number(fromVersion || 1) >= schema.version || !schema.migrate) {
      return { payload: safeClone(payload), schemaVersion: schema?.version || Number(fromVersion || DEFAULT_SCHEMA_VERSION) };
    }
    const migrated = schema.migrate(safeClone(payload), Number(fromVersion || DEFAULT_SCHEMA_VERSION), schema.version);
    return { payload: migrated, schemaVersion: schema.version };
  }

  function createOperation(type, payload, options = {}) {
    const migrated = migrate(type, payload, options.schemaVersion || DEFAULT_SCHEMA_VERSION);
    const operationId = String(options.operationId || "").trim() || createOperationId(String(type || "write").replace(/\s+/g, "-").slice(0, 32));
    const validation = validate(type, migrated.payload, { ...options, schemaVersion: migrated.schemaVersion });
    return {
      operationId,
      type: String(type || "write").trim() || "write",
      schemaVersion: migrated.schemaVersion,
      payload: migrated.payload,
      payloadBytes: validation.bytes,
      validation,
      createdAt: now(),
      meta: {
        queueKind: String(options.queueKind || ""),
        dedupeKey: String(options.dedupeKey || "")
      }
    };
  }

  function validationError(operation) {
    const error = new Error(operation.validation.errors.join("; ") || "Dati non validi");
    error.name = "HeraDataValidationError";
    error.code = "hera/invalid-data";
    error.validationErrors = operation.validation.errors.slice();
    return error;
  }

  function getOfflineEnqueue() {
    return originalOfflineEnqueue || (typeof window.enqueueOfflineMutation === "function" ? window.enqueueOfflineMutation : null);
  }

  function queueOperation(operation, queueKind) {
    const enqueue = getOfflineEnqueue();
    if (typeof enqueue !== "function") throw new Error("Coda offline non disponibile");
    const kind = String(queueKind || operation.meta?.queueKind || operation.type || "write");
    const result = enqueue.call(window, kind, operation.payload);
    journalOperation(operation, "queued", { queueKind: kind });
    dispatch("hera:data-safety-queued", { operationId: operation.operationId, type: operation.type, queueKind: kind });
    void snapshot(`queued:${operation.type}`);
    return result;
  }

  async function run(type, payload, perform, options = {}) {
    if (typeof perform !== "function") throw new TypeError("perform deve essere una funzione");

    const operation = createOperation(type, payload, options);
    if (!operation.validation.ok) throw validationError(operation);

    if (isCompleted(operation.operationId)) {
      return { status: "already-completed", operationId: operation.operationId, result: options.completedResult };
    }

    const dedupeKey = String(options.dedupeKey || "").trim();
    if (dedupeKey && inflight.has(dedupeKey)) return inflight.get(dedupeKey);

    const task = (async () => {
      journalOperation(operation, "starting", { dedupeKey });
      dispatch("hera:data-safety-start", { operationId: operation.operationId, type: operation.type });
      await snapshot(`before-write:${operation.type}`);

      if (options.queueKind && navigator.onLine === false) {
        queueOperation(operation, options.queueKind);
        return { status: "queued", operationId: operation.operationId };
      }

      try {
        const result = await perform(operation);
        markCompleted(operation);
        dispatch("hera:data-safety-complete", { operationId: operation.operationId, type: operation.type });
        void snapshot(`after-write:${operation.type}`);
        return { status: "completed", operationId: operation.operationId, result };
      } catch (error) {
        const canQueue = Boolean(options.queueKind) && isTransientError(error);
        if (canQueue) {
          queueOperation(operation, options.queueKind);
          journalOperation(operation, "queued-after-error", {
            queueKind: options.queueKind,
            errorCode: normalizeCode(error),
            errorMessage: errorMessage(error)
          });
          return { status: "queued", operationId: operation.operationId, error };
        }
        journalOperation(operation, "failed", {
          errorCode: normalizeCode(error),
          errorMessage: errorMessage(error)
        });
        dispatch("hera:data-safety-error", {
          operationId: operation.operationId,
          type: operation.type,
          code: normalizeCode(error),
          message: errorMessage(error)
        });
        throw error;
      }
    })().finally(() => {
      if (dedupeKey) inflight.delete(dedupeKey);
    });

    if (dedupeKey) inflight.set(dedupeKey, task);
    return task;
  }

  function installOfflineQueueObserver() {
    const current = window.enqueueOfflineMutation;
    if (typeof current !== "function" || current.__heraDataSafetyObserved) return false;

    originalOfflineEnqueue = current;
    function observedOfflineEnqueue(kind, payload) {
      const operation = createOperation(`offline:${String(kind || "mutation")}`, payload, {
        queueKind: String(kind || "mutation")
      });
      journalOperation(operation, "existing-queue", { queueKind: String(kind || "mutation") });
      try {
        const result = originalOfflineEnqueue.apply(this, arguments);
        dispatch("hera:data-safety-queue-observed", {
          operationId: operation.operationId,
          type: operation.type,
          queueKind: String(kind || "mutation")
        });
        void snapshot(`existing-queue:${String(kind || "mutation")}`);
        return result;
      } catch (error) {
        journalOperation(operation, "queue-failed", {
          queueKind: String(kind || "mutation"),
          errorCode: normalizeCode(error),
          errorMessage: errorMessage(error)
        });
        throw error;
      }
    }
    observedOfflineEnqueue.__heraDataSafetyObserved = true;
    observedOfflineEnqueue.__heraOriginal = current;
    window.enqueueOfflineMutation = observedOfflineEnqueue;
    return true;
  }

  function installOfflineSyncObserver() {
    const current = window.syncPendingOfflineMutations;
    if (typeof current !== "function" || current.__heraDataSafetyObserved) return false;

    originalOfflineSync = current;
    async function observedOfflineSync() {
      void snapshot("before-offline-sync");
      try {
        const result = await originalOfflineSync.apply(this, arguments);
        void snapshot("after-offline-sync");
        dispatch("hera:data-safety-sync-complete", { at: now() });
        return result;
      } catch (error) {
        try { window.HeraDataDurability?.setSyncError?.(error); } catch (_) {}
        dispatch("hera:data-safety-sync-error", { code: normalizeCode(error), message: errorMessage(error) });
        throw error;
      }
    }
    observedOfflineSync.__heraDataSafetyObserved = true;
    observedOfflineSync.__heraOriginal = current;
    window.syncPendingOfflineMutations = observedOfflineSync;
    return true;
  }

  function installExistingQueueBridge() {
    const queueInstalled = installOfflineQueueObserver();
    const syncInstalled = installOfflineSyncObserver();
    return { queueInstalled, syncInstalled };
  }

  function getState() {
    return {
      installed: true,
      version: VERSION,
      schemaVersion: DEFAULT_SCHEMA_VERSION,
      journal: pruneRows(readArray(JOURNAL_KEY), JOURNAL_TTL_MS, MAX_JOURNAL_ENTRIES),
      completedCount: readCompleted().length,
      inflightCount: inflight.size,
      schemas: Array.from(schemas.entries()).map(([type, schema]) => ({ type, version: schema.version })),
      queueBridge: {
        enqueueObserved: Boolean(window.enqueueOfflineMutation?.__heraDataSafetyObserved),
        syncObserved: Boolean(window.syncPendingOfflineMutations?.__heraDataSafetyObserved)
      }
    };
  }

  registerSchema("hoursReport", {
    version: 1,
    validate(payload) {
      const errors = [];
      if (payload?.entries !== undefined && !Array.isArray(payload.entries)) errors.push("entries deve essere un array");
      return errors;
    }
  });

  registerSchema("impianto", { version: 1 });
  registerSchema("commessa", { version: 1 });
  registerSchema("commessaNote", { version: 1 });
  registerSchema("squadra", { version: 1 });

  window.HeraDataSafety = {
    installed: true,
    version: VERSION,
    schemaVersion: DEFAULT_SCHEMA_VERSION,
    createOperationId,
    createOperation,
    validate,
    migrate,
    registerSchema,
    isTransientError,
    isCompleted,
    run,
    queue: (type, payload, options = {}) => {
      const operation = createOperation(type, payload, options);
      if (!operation.validation.ok) throw validationError(operation);
      return queueOperation(operation, options.queueKind || type);
    },
    snapshot,
    installExistingQueueBridge,
    getState
  };

  installExistingQueueBridge();
  window.setTimeout(installExistingQueueBridge, 1000);
  window.setTimeout(installExistingQueueBridge, 4000);

  window.addEventListener("offline", () => void snapshot("network-offline"));
  window.addEventListener("pagehide", () => void snapshot("pagehide"));

  dispatch("hera:data-safety-ready", { version: VERSION, schemaVersion: DEFAULT_SCHEMA_VERSION });
})();
