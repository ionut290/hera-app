(function () {
  "use strict";

  if (window.HeraFirestoreRegistryOptimizer?.installed) return;

  const firestoreApi = window.firebase?.firestore;
  const QueryProto = firestoreApi?.Query?.prototype;
  const DocumentReferenceProto = firestoreApi?.DocumentReference?.prototype;
  const CollectionReferenceProto = firestoreApi?.CollectionReference?.prototype;

  if (!QueryProto || typeof QueryProto.get !== "function" || typeof QueryProto.onSnapshot !== "function") {
    console.warn("Ottimizzatore letture personale/mezzi non installato: API Firestore non disponibile.");
    return;
  }

  const TARGET_TTL_MS = Object.freeze({
    personale: 30 * 1000,
    mezzi: 15 * 1000
  });
  const inFlight = new Map();
  const recent = new Map();
  const generations = new Map(Object.keys(TARGET_TTL_MS).map((name) => [name, 0]));
  const stats = {
    installedAt: new Date().toISOString(),
    networkGets: 0,
    reusedInFlight: 0,
    reusedRecent: 0,
    listenerSnapshots: 0,
    invalidations: 0
  };

  function canonicalPath(value) {
    if (!value) return "";
    if (typeof value === "string") return value;
    if (typeof value.canonicalString === "function") return value.canonicalString();
    if (Array.isArray(value.segments)) return value.segments.join("/");
    const text = String(value);
    return text === "[object Object]" ? "" : text;
  }

  function queryInfo(query) {
    const internal = query?._query || query?._delegate?._query || null;
    const path = String(
      query?.path ||
      canonicalPath(internal?.path) ||
      canonicalPath(internal?._path) ||
      ""
    ).replace(/^\/+|\/+$/g, "");
    const target = Object.prototype.hasOwnProperty.call(TARGET_TTL_MS, path) ? path : "";
    if (!target) return null;

    let canonical = "";
    try {
      canonical = typeof internal?.canonicalId === "function" ? internal.canonicalId() : "";
    } catch (_) {}

    // CollectionReference espone `path`; per le query composte richiediamo
    // invece l'identificatore canonico, così due filtri diversi non condividono dati.
    if (!query?.path && !canonical) return null;

    return {
      target,
      key: `${target}|${canonical || `collection:${path}`}`
    };
  }

  function documentTarget(ref) {
    const path = String(ref?.path || canonicalPath(ref?._key?.path) || "").replace(/^\/+|\/+$/g, "");
    const first = path.split("/")[0] || "";
    return Object.prototype.hasOwnProperty.call(TARGET_TTL_MS, first) ? first : "";
  }

  function invalidate(target) {
    if (!target) return;
    generations.set(target, (generations.get(target) || 0) + 1);
    for (const key of recent.keys()) {
      if (key.startsWith(`${target}|`)) recent.delete(key);
    }
    for (const key of inFlight.keys()) {
      if (key.startsWith(`${target}|`)) inFlight.delete(key);
    }
    stats.invalidations += 1;
  }

  function remember(info, snapshot, source) {
    if (!info || !snapshot) return;
    recent.set(info.key, {
      snapshot,
      savedAt: Date.now(),
      generation: generations.get(info.target) || 0,
      source
    });
    if (source === "listener") stats.listenerSnapshots += 1;
  }

  function getFresh(info) {
    const entry = recent.get(info.key);
    if (!entry) return null;
    const currentGeneration = generations.get(info.target) || 0;
    if (entry.generation !== currentGeneration || Date.now() - entry.savedAt > TARGET_TTL_MS[info.target]) {
      recent.delete(info.key);
      return null;
    }
    return entry.snapshot;
  }

  const originalGet = QueryProto.get;
  QueryProto.get = function optimizedRegistryGet(options) {
    const info = queryInfo(this);
    if (!info || options?.source) {
      return originalGet.apply(this, arguments);
    }

    const sourceKey = "default";
    const key = `${info.key}|source:${sourceKey}`;
    const fresh = getFresh(info);
    if (fresh) {
      stats.reusedRecent += 1;
      return Promise.resolve(fresh);
    }

    const pending = inFlight.get(key);
    if (pending) {
      stats.reusedInFlight += 1;
      return pending;
    }

    const generationAtStart = generations.get(info.target) || 0;
    stats.networkGets += 1;
    const request = Promise.resolve(originalGet.apply(this, arguments))
      .then((snapshot) => {
        if ((generations.get(info.target) || 0) === generationAtStart) {
          remember(info, snapshot, "get");
        }
        return snapshot;
      })
      .finally(() => {
        if (inFlight.get(key) === request) inFlight.delete(key);
      });

    inFlight.set(key, request);
    return request;
  };

  function wrapSnapshotHandler(info, handler) {
    if (typeof handler === "function") {
      return function optimizedSnapshotHandler(snapshot) {
        remember(info, snapshot, "listener");
        return handler.apply(this, arguments);
      };
    }
    if (handler && typeof handler === "object" && typeof handler.next === "function") {
      return Object.assign({}, handler, {
        next(snapshot) {
          remember(info, snapshot, "listener");
          return handler.next.call(handler, snapshot);
        }
      });
    }
    return handler;
  }

  const originalOnSnapshot = QueryProto.onSnapshot;
  QueryProto.onSnapshot = function optimizedRegistryOnSnapshot() {
    const info = queryInfo(this);
    if (!info) return originalOnSnapshot.apply(this, arguments);

    const args = Array.from(arguments);
    const firstIsOptions = args[0] && typeof args[0] === "object" &&
      typeof args[0].next !== "function" &&
      Object.prototype.hasOwnProperty.call(args[0], "includeMetadataChanges");
    const handlerIndex = firstIsOptions ? 1 : 0;
    args[handlerIndex] = wrapSnapshotHandler(info, args[handlerIndex]);
    return originalOnSnapshot.apply(this, args);
  };

  function wrapMutation(proto, method) {
    if (!proto || typeof proto[method] !== "function") return;
    const original = proto[method];
    proto[method] = function optimizedRegistryMutation() {
      const target = documentTarget(this);
      if (target) invalidate(target);
      return original.apply(this, arguments);
    };
  }

  wrapMutation(DocumentReferenceProto, "set");
  wrapMutation(DocumentReferenceProto, "update");
  wrapMutation(DocumentReferenceProto, "delete");
  wrapMutation(CollectionReferenceProto, "add");

  window.HeraFirestoreRegistryOptimizer = {
    installed: true,
    stats,
    invalidate,
    getState() {
      return {
        stats: { ...stats },
        inFlight: inFlight.size,
        recent: recent.size
      };
    }
  };
})();
