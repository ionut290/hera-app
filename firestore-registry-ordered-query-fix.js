(function () {
  "use strict";

  const GLOBAL_NAME = "HeraFirestoreRegistryOrderedQueryFix";
  const WRAPPED_ORDER_BY = "__heraRegistryOrderedQueryOrderBy";
  const WRAPPED_GET = "__heraRegistryOrderedQueryGet";
  const WRAPPED_SNAPSHOT = "__heraRegistryOrderedQuerySnapshot";
  const INSTALL_RETRY_MS = 100;
  const INSTALL_RETRY_LIMIT = 150;
  const TARGETS = new Set(["personale", "mezzi"]);
  const TARGET_TTL_MS = Object.freeze({ personale: 30 * 1000, mezzi: 15 * 1000 });
  const DEVICE_CACHE_MAX_AGE_MS = Object.freeze({
    personale: 6 * 60 * 60 * 1000,
    mezzi: 12 * 60 * 60 * 1000
  });

  if (window[GLOBAL_NAME]?.installed) return;

  const safeQueries = typeof WeakMap === "function" ? new WeakMap() : null;
  const inFlight = new Map();
  const recent = new Map();
  let installAttempts = 0;
  let retryTimer = null;

  const stats = {
    installedAt: "",
    installAttempts: 0,
    markedQueries: 0,
    interceptedGets: 0,
    interceptedListeners: 0,
    reusedInFlight: 0,
    reusedRecent: 0,
    reusedDeviceCache: 0,
    sourceServerCacheHits: 0,
    networkGets: 0,
    networkFallbacks: 0,
    listenerSnapshots: 0,
    deviceCacheWrites: 0,
    invalidations: 0
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
    const text = String(value);
    return text === "[object Object]" ? "" : text;
  }

  function queryPath(query) {
    const direct = canonicalPath(query?.path);
    if (direct) return direct.replace(/^\/+|\/+$/g, "");

    const candidates = [
      query?._query,
      query?.Ae,
      query?.je,
      query?._delegate?._query,
      query?._delegate?.Ae,
      query?._delegate?.je
    ].filter(Boolean);

    for (const internal of candidates) {
      const values = [internal?.path, internal?._path, internal?.query?.path, internal?._query?.path];
      for (const value of values) {
        const path = canonicalPath(value);
        if (path) return path.replace(/^\/+|\/+$/g, "");
      }
    }
    return "";
  }

  function targetName(query) {
    const path = queryPath(query);
    return TARGETS.has(path) ? path : "";
  }

  function fieldPathName(value) {
    if (typeof value === "string") return value;
    if (value && typeof value.toString === "function") {
      const text = String(value.toString() || "");
      if (text && text !== "[object Object]") return text;
    }
    return "";
  }

  function isRawTargetCollection(query) {
    return Boolean(
      targetName(query) &&
      typeof query?.doc === "function" &&
      typeof query?.add === "function"
    );
  }

  function orderedQueryInfo(query) {
    return safeQueries?.get(query) || null;
  }

  function optimizerStats() {
    return window.HeraFirestoreRegistryOptimizer?.stats || null;
  }

  function bumpOptimizerStat(name, amount = 1) {
    const target = optimizerStats();
    if (!target || typeof target[name] !== "number") return;
    target[name] += amount;
  }

  function cachedRecordId(record) {
    return String(record?.__heraDocId || record?.id || record?.docId || record?._id || "").trim();
  }

  function cachedRecordData(record) {
    const data = { ...(record || {}) };
    delete data.__heraDocId;
    delete data.id;
    delete data.docId;
    delete data._id;
    return data;
  }

  function readNestedValue(data, fieldPath) {
    if (typeof fieldPath !== "string" || !fieldPath) return undefined;
    return fieldPath.split(".").reduce((value, part) => value == null ? undefined : value[part], data);
  }

  function timestampMillis(value) {
    if (value == null) return Number.POSITIVE_INFINITY;
    if (value instanceof Date) return value.getTime();
    if (typeof value?.toMillis === "function") {
      try {
        const millis = Number(value.toMillis());
        if (Number.isFinite(millis)) return millis;
      } catch (_) {}
    }
    const seconds = Number(value?.seconds ?? value?._seconds);
    const nanoseconds = Number(value?.nanoseconds ?? value?._nanoseconds ?? 0);
    if (Number.isFinite(seconds)) {
      return seconds * 1000 + (Number.isFinite(nanoseconds) ? nanoseconds / 1e6 : 0);
    }
    if (typeof value === "number" && Number.isFinite(value)) return value;
    if (typeof value === "string") {
      const millis = Date.parse(value);
      if (Number.isFinite(millis)) return millis;
    }
    return Number.POSITIVE_INFINITY;
  }

  function sortCreatedAtAscending(records) {
    return records
      .map((record, index) => ({ record, index }))
      .sort((left, right) => {
        const delta = timestampMillis(left.record?.createdAt) - timestampMillis(right.record?.createdAt);
        if (Number.isFinite(delta) && delta !== 0) return delta;
        const idDelta = cachedRecordId(left.record).localeCompare(cachedRecordId(right.record));
        return idDelta || left.index - right.index;
      })
      .map((entry) => entry.record);
  }

  function buildDeviceCachedSnapshot(query, info, records) {
    if (!info || !Array.isArray(records) || !records.length) return null;
    const firestore = query?.firestore || query?._delegate?.firestore;
    if (!firestore || typeof firestore.collection !== "function") return null;

    const metadata = Object.freeze({ fromCache: true, hasPendingWrites: false });
    const docs = [];
    for (const record of sortCreatedAtAscending(records)) {
      const id = cachedRecordId(record);
      if (!id) return null;
      const data = cachedRecordData(record);
      const ref = firestore.collection(info.target).doc(id);
      docs.push(Object.freeze({
        id,
        ref,
        exists: true,
        metadata,
        data() { return { ...data }; },
        get(fieldPath) { return readNestedValue(data, fieldPath); },
        isEqual(other) {
          return Boolean(other && other.id === id && other.ref?.path === ref.path);
        }
      }));
    }

    const snapshot = {
      query,
      docs: Object.freeze(docs),
      size: docs.length,
      empty: docs.length === 0,
      metadata,
      forEach(callback, thisArg) {
        docs.forEach((document) => callback.call(thisArg, document));
      },
      docChanges() {
        return docs.map((document, index) => ({
          type: "added",
          doc: document,
          oldIndex: -1,
          newIndex: index
        }));
      },
      isEqual(other) { return other === snapshot; }
    };
    return Object.freeze(snapshot);
  }

  function extractSnapshotRecords(snapshot) {
    if (!Array.isArray(snapshot?.docs)) return null;
    const records = [];
    for (const document of snapshot.docs) {
      const id = String(document?.id || "").trim();
      if (!id || typeof document?.data !== "function") return null;
      const data = document.data();
      if (!data || typeof data !== "object" || Array.isArray(data)) return null;
      records.push({ id, ...data });
    }
    return records;
  }

  function persistSnapshot(info, snapshot) {
    const cache = window.HeraRegistryDeviceCache;
    if (!cache || typeof cache.writeIfChanged !== "function") return;
    const records = extractSnapshotRecords(snapshot);
    if (!records?.length) return;
    Promise.resolve(cache.writeIfChanged(info.target, records))
      .then((changed) => {
        if (changed) {
          stats.deviceCacheWrites += 1;
          bumpOptimizerStat("deviceCacheWrites");
        }
      })
      .catch((error) => {
        console.warn(`Cache ordinata ${info.target} non aggiornata:`, error);
      });
  }

  function remember(info, snapshot, source) {
    recent.set(info.key, { snapshot, savedAt: Date.now() });
    if (source === "listener") {
      stats.listenerSnapshots += 1;
      bumpOptimizerStat("listenerSnapshots");
    }
    if (source !== "device-cache") persistSnapshot(info, snapshot);
  }

  function freshSnapshot(info) {
    const entry = recent.get(info.key);
    if (!entry) return null;
    if (Date.now() - entry.savedAt > TARGET_TTL_MS[info.target]) {
      recent.delete(info.key);
      return null;
    }
    return entry.snapshot;
  }

  async function deviceCachedSnapshot(query, info) {
    const cache = window.HeraRegistryDeviceCache;
    if (!cache || typeof cache.readFresh !== "function") return null;
    try {
      const cached = await cache.readFresh(info.target, DEVICE_CACHE_MAX_AGE_MS[info.target]);
      if (!cached?.records?.length) return null;
      return buildDeviceCachedSnapshot(query, info, cached.records);
    } catch (error) {
      console.warn(`Cache ordinata ${info.target} non leggibile:`, error);
      return null;
    }
  }

  function sourceMode(options) {
    if (options == null) return "default";
    if (typeof options !== "object" || Array.isArray(options)) return "unsupported";
    const keys = Object.keys(options);
    if (!keys.length) return "default";
    if (keys.some((key) => key !== "source")) return "unsupported";
    if (options.source === "server") return "server";
    if (options.source === "default" || options.source == null) return "default";
    if (options.source === "cache") return "cache";
    return "unsupported";
  }

  function invalidate(target) {
    if (!TARGETS.has(target)) return false;
    recent.delete(`${target}|createdAt:asc`);
    inFlight.delete(`${target}|createdAt:asc`);
    stats.invalidations += 1;
    return true;
  }

  function unwrap(method) {
    return typeof method?.__heraOriginal === "function" ? method.__heraOriginal : method;
  }

  function handleGet(query, originalGet, argsLike) {
    const info = orderedQueryInfo(query);
    if (!info) return originalGet.apply(query, argsLike);

    const args = Array.from(argsLike);
    const mode = sourceMode(args[0]);
    if (mode === "cache" || mode === "unsupported") {
      return originalGet.apply(query, args);
    }

    stats.interceptedGets += 1;
    bumpOptimizerStat("interceptedGets");

    const fresh = freshSnapshot(info);
    if (fresh) {
      stats.reusedRecent += 1;
      bumpOptimizerStat("reusedRecent");
      if (mode === "server") {
        stats.sourceServerCacheHits += 1;
        bumpOptimizerStat("sourceServerCacheHits");
      }
      return Promise.resolve(fresh);
    }

    const pending = inFlight.get(info.key);
    if (pending) {
      stats.reusedInFlight += 1;
      bumpOptimizerStat("reusedInFlight");
      if (mode === "server") {
        stats.sourceServerCacheHits += 1;
        bumpOptimizerStat("sourceServerCacheHits");
      }
      return pending;
    }

    const nativeGet = unwrap(originalGet);
    let request;
    request = Promise.resolve()
      .then(() => deviceCachedSnapshot(query, info))
      .then((cachedSnapshot) => {
        if (cachedSnapshot) {
          stats.reusedDeviceCache += 1;
          bumpOptimizerStat("reusedDeviceCache");
          if (mode === "server") {
            stats.sourceServerCacheHits += 1;
            bumpOptimizerStat("sourceServerCacheHits");
          }
          remember(info, cachedSnapshot, "device-cache");
          return cachedSnapshot;
        }

        stats.networkGets += 1;
        stats.networkFallbacks += 1;
        bumpOptimizerStat("networkGets");
        bumpOptimizerStat("networkFallbacks");
        return Promise.resolve(nativeGet.apply(query, args)).then((snapshot) => {
          remember(info, snapshot, "get");
          return snapshot;
        });
      })
      .finally(() => {
        if (inFlight.get(info.key) === request) inFlight.delete(info.key);
      });

    inFlight.set(info.key, request);
    return request;
  }

  function wrapSnapshotHandler(info, handler) {
    if (typeof handler === "function") {
      return function orderedRegistrySnapshotHandler(snapshot) {
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

  function handleOnSnapshot(query, originalOnSnapshot, argsLike) {
    const info = orderedQueryInfo(query);
    if (!info) return originalOnSnapshot.apply(query, argsLike);

    stats.interceptedListeners += 1;
    bumpOptimizerStat("interceptedListeners");

    const args = Array.from(argsLike);
    const first = args[0];
    const firstIsOptions = first && typeof first === "object" &&
      typeof first.next !== "function" &&
      (Object.prototype.hasOwnProperty.call(first, "includeMetadataChanges") || Object.keys(first).length === 0);
    const handlerIndex = firstIsOptions ? 1 : 0;
    args[handlerIndex] = wrapSnapshotHandler(info, args[handlerIndex]);
    return unwrap(originalOnSnapshot).apply(query, args);
  }

  function markOrderedQuery(source, result, fieldPath, direction) {
    if (!safeQueries || !result || typeof result !== "object" || !isRawTargetCollection(source)) return;
    const field = fieldPathName(fieldPath);
    const normalizedDirection = String(direction || "asc").toLowerCase();
    if (field !== "createdAt" || normalizedDirection !== "asc") return;

    const target = targetName(source);
    safeQueries.set(result, {
      target,
      path: target,
      field: "createdAt",
      direction: "asc",
      key: `${target}|createdAt:asc`
    });
    stats.markedQueries += 1;
  }

  function patchQueryPrototype(proto) {
    if (!proto) return false;
    let changed = false;

    if (typeof proto.orderBy === "function" && !proto.orderBy[WRAPPED_ORDER_BY]) {
      const originalOrderBy = proto.orderBy;
      const wrappedOrderBy = function orderedRegistryOrderBy(fieldPath, direction) {
        const result = originalOrderBy.apply(this, arguments);
        markOrderedQuery(this, result, fieldPath, direction);
        return result;
      };
      Object.defineProperty(wrappedOrderBy, WRAPPED_ORDER_BY, { value: true });
      Object.defineProperty(wrappedOrderBy, "__heraOriginal", { value: unwrap(originalOrderBy) });
      proto.orderBy = wrappedOrderBy;
      changed = true;
    }

    if (typeof proto.get === "function" && !proto.get[WRAPPED_GET]) {
      const originalGet = proto.get;
      const wrappedGet = function orderedRegistryGet() {
        return handleGet(this, originalGet, arguments);
      };
      Object.defineProperty(wrappedGet, WRAPPED_GET, { value: true });
      Object.defineProperty(wrappedGet, "__heraOriginal", { value: unwrap(originalGet) });
      proto.get = wrappedGet;
      changed = true;
    }

    if (typeof proto.onSnapshot === "function" && !proto.onSnapshot[WRAPPED_SNAPSHOT]) {
      const originalOnSnapshot = proto.onSnapshot;
      const wrappedOnSnapshot = function orderedRegistryOnSnapshot() {
        return handleOnSnapshot(this, originalOnSnapshot, arguments);
      };
      Object.defineProperty(wrappedOnSnapshot, WRAPPED_SNAPSHOT, { value: true });
      Object.defineProperty(wrappedOnSnapshot, "__heraOriginal", { value: unwrap(originalOnSnapshot) });
      proto.onSnapshot = wrappedOnSnapshot;
      changed = true;
    }

    return changed;
  }

  function attachOptimizerState() {
    const optimizer = window.HeraFirestoreRegistryOptimizer;
    if (!optimizer || optimizer.__heraOrderedQueryStateAttached || typeof optimizer.getState !== "function") return;
    const originalGetState = optimizer.getState.bind(optimizer);
    optimizer.getState = function getStateWithOrderedQueryFix() {
      return {
        ...originalGetState(),
        orderedQueryFix: api.getState()
      };
    };
    Object.defineProperty(optimizer, "__heraOrderedQueryStateAttached", { value: true });
  }

  function install() {
    installAttempts += 1;
    stats.installAttempts = installAttempts;

    const proto = window.firebase?.firestore?.Query?.prototype;
    const installed = patchQueryPrototype(proto);
    attachOptimizerState();

    api.installed = Boolean(installed || proto?.get?.[WRAPPED_GET]);
    api.pending = !api.installed;
    if (api.installed && !stats.installedAt) stats.installedAt = new Date().toISOString();

    if (!api.installed && installAttempts < INSTALL_RETRY_LIMIT && !retryTimer) {
      retryTimer = setTimeout(() => {
        retryTimer = null;
        install();
      }, INSTALL_RETRY_MS);
    }
    return api.installed;
  }

  const api = {
    installed: false,
    pending: true,
    stats,
    install,
    invalidate,
    getState() {
      return {
        installed: api.installed,
        pending: api.pending,
        stats: { ...stats },
        recent: recent.size,
        inFlight: inFlight.size
      };
    }
  };

  window[GLOBAL_NAME] = api;
  install();

  if (typeof window.addEventListener === "function") {
    window.addEventListener("load", install, { once: true });
    window.addEventListener("hera:firebase-ready", install);
  }
})();
