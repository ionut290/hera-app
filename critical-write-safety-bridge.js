(() => {
  "use strict";

  if (window.HeraCriticalWriteSafetyBridge?.installed) return;

  const VERSION = "1.1.0";
  const RETRY_MS = 500;
  const MAX_ATTEMPTS = 40;
  const wrapped = new Set();
  const state = {
    installed: true,
    version: VERSION,
    attempts: 0,
    wrapped: [],
    lastError: ""
  };

  const TARGETS = [
    // FATTO e RESET hanno gia una coda persistente e un retry dedicati.
    // Non avvolgere le funzioni composte: lo snapshot IndexedDB eseguito prima
    // della chiamata puo restare sospeso su iOS e bloccare il pulsante.
    { name: "deleteImpianto", type: "impianto:delete" },
    { name: "saveImpianto", type: "impianto:save" },
    { name: "updateImpianto", type: "impianto:update" },
    { name: "saveCommessa", type: "commessa:save" },
    { name: "saveCommessaNote", type: "commessaNote:save" },
    { name: "saveHoursReport", type: "hoursReport:save" },
    { name: "saveSquadra", type: "squadra:save" }
  ];

  const text = (value) => String(value ?? "").trim();

  function readOptions(args) {
    for (let index = args.length - 1; index >= 0; index -= 1) {
      const value = args[index];
      if (value && typeof value === "object" && !Array.isArray(value)) return value;
    }
    return {};
  }

  function firstEntity(args) {
    const value = args[0];
    return value && typeof value === "object" && !Array.isArray(value) ? value : {};
  }

  function getEntityId(entity, args) {
    const direct = text(
      entity?.id
      || entity?.impiantoId
      || entity?.physicalPlantId
      || entity?.idSap
      || entity?.sap
      || entity?.commessaId
      || entity?.reportId
    );
    if (direct) return direct;
    const scalar = args.find((value) => typeof value === "string" || typeof value === "number");
    return text(scalar).slice(0, 160);
  }

  function getCommessaId(entity, options) {
    return text(
      options?.commessaId
      || entity?.commessaId
      || entity?.parentCommessaId
      || window.selectedCommessaId
      || ""
    ).slice(0, 160);
  }

  function buildMeta(target, args) {
    const entity = firstEntity(args);
    const options = readOptions(args);
    const entityId = getEntityId(entity, args);
    const commessaId = getCommessaId(entity, options);
    return {
      action: target.type,
      entityId,
      commessaId,
      source: text(options?.source || "runtime-bridge").slice(0, 80)
    };
  }

  function getOperationId(target, args) {
    const entity = firstEntity(args);
    const options = readOptions(args);
    return text(options?.operationId || entity?.operationId || "")
      || window.HeraDataSafety?.createOperationId?.(target.type.replace(/[^a-z0-9]+/gi, "-"));
  }

  function isAsyncFunction(fn) {
    try {
      return fn?.constructor?.name === "AsyncFunction" || /^\s*async\b/.test(Function.prototype.toString.call(fn));
    } catch (_) {
      return false;
    }
  }

  function validateMeta(target, meta) {
    const safety = window.HeraDataSafety;
    if (!safety?.validate) return;
    const result = safety.validate(target.type, meta);
    if (result?.ok !== false) return;
    const error = new Error((result.errors || []).join("; ") || `Dati ${target.type} non validi`);
    error.code = "hera/critical-write-invalid";
    throw error;
  }

  function reportError(error) {
    state.lastError = text(error?.message || error).slice(0, 300);
    try { window.HeraDataDurability?.setSyncError?.(error); } catch (_) {}
  }

  function makeAsyncWrapper(target, original) {
    const wrappedFunction = async function criticalWriteSafetyAsyncWrapper(...args) {
      const safety = window.HeraDataSafety;
      if (!safety?.run) return original.apply(this, args);

      const meta = buildMeta(target, args);
      const operationId = getOperationId(target, args);
      const dedupeIdentity = meta.entityId || operationId;
      // forceMoveImpiantoToFatti compone markImpiantoDone: le due funzioni
      // possono essere attive nello stesso momento per la stessa scheda. Se la
      // chiave non distingue il punto d'ingresso, la chiamata interna riceve la
      // Promise della chiamata esterna e le due operazioni si attendono a
      // vicenda senza mai arrivare alla scrittura Firestore.
      const dedupeKey = `${target.name}:${target.type}:${meta.commessaId}:${dedupeIdentity}`;
      try {
        const outcome = await safety.run(
          target.type,
          meta,
          () => original.apply(this, args),
          { operationId, dedupeKey }
        );
        return outcome?.result;
      } catch (error) {
        reportError(error);
        throw error;
      }
    };
    wrappedFunction.__heraCriticalWriteSafetyWrapped = true;
    wrappedFunction.__heraCriticalWriteSafetyOriginal = original;
    return wrappedFunction;
  }

  function makeSyncWrapper(target, original) {
    const wrappedFunction = function criticalWriteSafetySyncWrapper(...args) {
      const safety = window.HeraDataSafety;
      if (!safety) return original.apply(this, args);

      const meta = buildMeta(target, args);
      try {
        validateMeta(target, meta);
        void safety.snapshot?.(`before-write:${target.type}`);
        const result = original.apply(this, args);
        if (result && typeof result.then === "function") {
          return Promise.resolve(result).then(
            (value) => {
              void safety.snapshot?.(`after-write:${target.type}`);
              return value;
            },
            (error) => {
              reportError(error);
              throw error;
            }
          );
        }
        void safety.snapshot?.(`after-write:${target.type}`);
        return result;
      } catch (error) {
        reportError(error);
        throw error;
      }
    };
    wrappedFunction.__heraCriticalWriteSafetyWrapped = true;
    wrappedFunction.__heraCriticalWriteSafetyOriginal = original;
    return wrappedFunction;
  }

  function wrapTarget(target) {
    const current = window[target.name];
    if (typeof current !== "function") return false;
    if (current.__heraCriticalWriteSafetyWrapped) {
      wrapped.add(target.name);
      return true;
    }

    window[target.name] = isAsyncFunction(current)
      ? makeAsyncWrapper(target, current)
      : makeSyncWrapper(target, current);
    wrapped.add(target.name);
    state.wrapped = Array.from(wrapped).sort();
    window.dispatchEvent(new CustomEvent("hera:critical-write-safety-wrapped", {
      detail: { name: target.name, type: target.type, version: VERSION }
    }));
    return true;
  }

  function install() {
    state.attempts += 1;
    if (!window.HeraDataSafety?.installed) return false;
    TARGETS.forEach(wrapTarget);
    return wrapped.size > 0;
  }

  function scheduleInstall() {
    install();
    const timer = window.setInterval(() => {
      install();
      if (state.attempts >= MAX_ATTEMPTS || wrapped.size === TARGETS.length) {
        window.clearInterval(timer);
      }
    }, RETRY_MS);
  }

  window.HeraCriticalWriteSafetyBridge = {
    installed: true,
    version: VERSION,
    install,
    getState: () => ({ ...state, wrapped: Array.from(wrapped).sort() })
  };

  scheduleInstall();
})();
