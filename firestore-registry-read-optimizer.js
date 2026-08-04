(function () {
  "use strict";

  const GLOBAL_NAME = "HeraFirestoreRegistryOptimizer";
  const WRAPPED_GET = "__heraRegistryOptimizedGet";
  const WRAPPED_SNAPSHOT = "__heraRegistryOptimizedSnapshot";
  const WRAPPED_MUTATION = "__heraRegistryOptimizedMutation";
  const INSTALL_RETRY_MS = 100;
  const INSTALL_RETRY_LIMIT = 150;
  const TARGET_TTL_MS = Object.freeze({
    personale: 30 * 1000,
    mezzi: 15 * 1000
  });
  const DEVICE_CACHE_MAX_AGE_MS = Object.freeze({
    personale: 6 * 60 * 60 * 1000,
    mezzi: 12 * 60 * 60 * 1000
  });
  const PROFILE_WRITE_DEDUPE_MS = 60 * 1000;
  const PROFILE_FIELDS = new Set(["name", "email", "photoURL", "updatedAt"]);
  const TARGETS = new Set(Object.keys(TARGET_TTL_MS));

  const existing = window[GLOBAL_NAME];
  if (existing && typeof existing.refreshInstallation === "function") {
    existing.refreshInstallation();
    return;
  }

  const inFlight = new Map();
  const recent = new Map();
  const recentProfileWrites = new Map();
  const generations = new Map(Array.from(TARGETS, (name) => [name, 0]));
  const patchedPrototypes = new Set();
  let installAttempts = 0;
  let retryTimer = null;
  let objectSequence = 0;
  const objectKeys = typeof WeakMap === "function" ? new WeakMap() : null;

  const stats = {
    installedAt: "",
    installAttempts: 0,
    patchedPrototypeCount: 0,
    networkGets: 0,
    interceptedGets: 0,
    interceptedListeners: 0,
    reusedInFlight: 0,
    reusedRecent: 0,
    reusedDeviceCache: 0,
    sourceServerCacheHits: 0,
    sourceCacheBypasses: 0,
    unsupportedOptionBypasses: 0,
    queryInfoMisses: 0,
    filteredQueryBypasses: 0,
    deviceCacheWrites: 0,
    deviceCacheReadErrors: 0,
    networkFallbacks: 0,
    listenerSnapshots: 0,
    invalidations: 0,
    profileWritesPassed: 0,
    profileWritesSkipped: 0
  };

  function targetName(value) {
    const normalized = String(value || "").replace(/^\/+|\/+$/g, "");
    return TARGETS.has(normalized) ? normalized : "";
  }

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

  function internalCandidates(query) {
    const candidates = [
      query?._query,
      query?.Ae,
      query?.je,
      query?._delegate?._query,
      query?._delegate?.Ae,
      query?._delegate?.je
    ];
    return candidates.filter((value, index) => value && candidates.indexOf(value) === index);
  }

  function firstInternalPath(query) {
    const direct = canonicalPath(query?.path);
    if (direct) return direct;

    for (const internal of internalCandidates(query)) {
      const values = [
        internal?.path,
        internal?._path,
        internal?.query?.path,
        internal?._query?.path
      ];
      for (const value of values) {
        const path = canonicalPath(value);
        if (path) return path;
      }
    }
    return "";
  }

  function canonicalQueryId(query) {
    for (const internal of internalCandidates(query)) {
      const methods = ["canonicalId", "canonicalString"];
      for (const name of methods) {
        if (typeof internal?.[name] !== "function") continue;
        try {
          const value = String(internal[name]() || "");
          if (value) return value;
        } catch (_) {}
      }
    }
    return "";
  }

  function objectKey(value) {
    if (!value || !objectKeys) return "";
    let key = objectKeys.get(value);
    if (!key) {
      objectSequence += 1;
      key = `object:${objectSequence}`;
      objectKeys.set(value, key);
    }
    return key;
  }

  function isCollectionReference(query, path) {
    if (!query || !path) return false;
    if (typeof query.doc === "function" && typeof query.add === "function") return true;

    const constructorName = String(query?.constructor?.name || "").toLowerCase();
    if (constructorName.includes("collectionreference")) return true;

    const collectionProto = window.firebase?.firestore?.CollectionReference?.prototype;
    if (collectionProto && typeof collectionProto.isPrototypeOf === "function") {
      try {
        if (collectionProto.isPrototypeOf(query)) return true;
      } catch (_) {}
    }

    return false;
  }

  function queryInfo(query) {
    const path = String(firstInternalPath(query) || "").replace(/^\/+|\/+$/g, "");
    const target = targetName(path);
    if (!target) {
      stats.queryInfoMisses += 1;
      return null;
    }

    const fullCollection = isCollectionReference(query, path);
    if (!fullCollection) {
      stats.filteredQueryBypasses += 1;
      return null;
    }

    const canonical = canonicalQueryId(query);
    return {
      target,
      path,
      isFullCollection: true,
      key: `${target}|${canonical || `collection:${path}` || objectKey(query)}`
    };
  }

  function documentPath(ref) {
    const values = [
      ref?.path,
      ref?._key?.path,
      ref?._delegate?._key?.path,
      ref?._delegate?.path
    ];
    for (const value of values) {
      const path = canonicalPath(value);
      if (path) return String(path).replace(/^\/+|\/+$/g, "");
    }
    return "";
  }

  function documentTarget(ref) {
    return targetName(documentPath(ref).split("/")[0] || "");
  }

  function invalidate(target) {
    if (!TARGETS.has(target)) return;
    generations.set(target, (generations.get(target) || 0) + 1);
    for (const key of Array.from(recent.keys())) {
      if (key.startsWith(`${target}|`)) recent.delete(key);
    }
    for (const key of Array.from(inFlight.keys())) {
      if (key.startsWith(`${target}|`)) inFlight.delete(key);
    }
    stats.invalidations += 1;
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

  function persistSnapshotToDevice(info, snapshot) {
    if (!info?.isFullCollection) return;
    const cache = window.HeraRegistryDeviceCache;
    if (!cache || typeof cache.writeIfChanged !== "function") return;
    const records = extractSnapshotRecords(snapshot);
    if (!records?.length) return;
    Promise.resolve(cache.writeIfChanged(info.target, records))
      .then((changed) => {
        if (changed) stats.deviceCacheWrites += 1;
      })
      .catch((error) => {
        console.warn(`Cache locale ${info.target} non aggiornata:`, error);
      });
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
    if (source !== "device-cache") persistSnapshotToDevice(info, snapshot);
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

  function buildDeviceCachedSnapshot(query, info, records) {
    if (!info?.isFullCollection || !Array.isArray(records) || !records.length) return null;
    const firestore = query?.firestore || query?._delegate?.firestore;
    if (!firestore || typeof firestore.collection !== "function") return null;

    const metadata = Object.freeze({ fromCache: true, hasPendingWrites: false });
    const docs = [];
    for (const record of records) {
      const id = cachedRecordId(record);
      if (!id) return null;
      const data = cachedRecordData(record);
      const ref = firestore.collection(info.target).doc(id);
      const document = {
        id,
        ref,
        exists: true,
        metadata,
        data() {
          return { ...data };
        },
        get(fieldPath) {
          return readNestedValue(data, fieldPath);
        },
        isEqual(other) {
          return Boolean(other && other.id === id && other.ref?.path === ref.path);
        }
      };
      docs.push(Object.freeze(document));
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
      isEqual(other) {
        return other === snapshot;
      }
    };
    return Object.freeze(snapshot);
  }

  async function getDeviceCachedSnapshot(query, info) {
    if (!info?.isFullCollection) return null;
    const cache = window.HeraRegistryDeviceCache;
    if (!cache || typeof cache.readFresh !== "function") return null;
    try {
      const cached = await cache.readFresh(info.target, DEVICE_CACHE_MAX_AGE_MS[info.target]);
      if (!cached?.records?.length) return null;
      return buildDeviceCachedSnapshot(query, info, cached.records);
    } catch (error) {
      stats.deviceCacheReadErrors += 1;
      console.warn(`Cache locale ${info.target} non leggibile:`, error);
      return null;
    }
  }

  function getSourceMode(options) {
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

  function handleGet(query, originalGet, argsLike) {
    const args = Array.from(argsLike);
    const options = args[0];
    const info = queryInfo(query);
    if (!info) return originalGet.apply(query, args);

    const sourceMode = getSourceMode(options);
    if (sourceMode === "cache") {
      stats.sourceCacheBypasses += 1;
      return originalGet.apply(query, args);
    }
    if (sourceMode === "unsupported") {
      stats.unsupportedOptionBypasses += 1;
      return originalGet.apply(query, args);
    }

    stats.interceptedGets += 1;
    const key = `${info.key}|cacheable`;
    const fresh = getFresh(info);
    if (fresh) {
      stats.reusedRecent += 1;
      if (sourceMode === "server") stats.sourceServerCacheHits += 1;
      return Promise.resolve(fresh);
    }

    const pending = inFlight.get(key);
    if (pending) {
      stats.reusedInFlight += 1;
      if (sourceMode === "server") stats.sourceServerCacheHits += 1;
      return pending;
    }

    const generationAtStart = generations.get(info.target) || 0;
    let request;
    request = Promise.resolve()
      .then(() => getDeviceCachedSnapshot(query, info))
      .then((cachedSnapshot) => {
        if (cachedSnapshot) {
          stats.reusedDeviceCache += 1;
          if (sourceMode === "server") stats.sourceServerCacheHits += 1;
          remember(info, cachedSnapshot, "device-cache");
          return cachedSnapshot;
        }

        stats.networkGets += 1;
        stats.networkFallbacks += 1;
        return Promise.resolve(originalGet.apply(query, args)).then((snapshot) => {
          if ((generations.get(info.target) || 0) === generationAtStart) {
            remember(info, snapshot, "get");
          }
          return snapshot;
        });
      })
      .finally(() => {
        if (inFlight.get(key) === request) inFlight.delete(key);
      });

    inFlight.set(key, request);
    return request;
  }

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

  function handleOnSnapshot(query, originalOnSnapshot, argsLike) {
    const info = queryInfo(query);
    if (!info) return originalOnSnapshot.apply(query, argsLike);

    stats.interceptedListeners += 1;
    const args = Array.from(argsLike);
    const first = args[0];
    const firstIsOptions = first && typeof first === "object" &&
      typeof first.next !== "function" &&
      (
        Object.prototype.hasOwnProperty.call(first, "includeMetadataChanges") ||
        Object.keys(first).length === 0
      );
    const handlerIndex = firstIsOptions ? 1 : 0;
    args[handlerIndex] = wrapSnapshotHandler(info, args[handlerIndex]);
    return originalOnSnapshot.apply(query, args);
  }

  function patchQueryPrototype(proto) {
    if (!proto || patchedPrototypes.has(proto)) return false;
    let changed = false;

    if (Object.prototype.hasOwnProperty.call(proto, "get") &&
        typeof proto.get === "function" &&
        !proto.get[WRAPPED_GET]) {
      const originalGet = proto.get;
      const wrappedGet = function optimizedRegistryGet() {
        return handleGet(this, originalGet, arguments);
      };
      Object.defineProperty(wrappedGet, WRAPPED_GET, { value: true });
      Object.defineProperty(wrappedGet, "__heraOriginal", { value: originalGet });
      proto.get = wrappedGet;
      changed = true;
    }

    if (Object.prototype.hasOwnProperty.call(proto, "onSnapshot") &&
        typeof proto.onSnapshot === "function" &&
        !proto.onSnapshot[WRAPPED_SNAPSHOT]) {
      const originalOnSnapshot = proto.onSnapshot;
      const wrappedOnSnapshot = function optimizedRegistryOnSnapshot() {
        return handleOnSnapshot(this, originalOnSnapshot, arguments);
      };
      Object.defineProperty(wrappedOnSnapshot, WRAPPED_SNAPSHOT, { value: true });
      Object.defineProperty(wrappedOnSnapshot, "__heraOriginal", { value: originalOnSnapshot });
      proto.onSnapshot = wrappedOnSnapshot;
      changed = true;
    }

    if (changed) {
      patchedPrototypes.add(proto);
      stats.patchedPrototypeCount = patchedPrototypes.size;
    }
    return changed;
  }

  function profilePatchPayload(data) {
    if (!data || typeof data !== "object" || Array.isArray(data)) return null;
    const keys = Object.keys(data);
    if (!keys.length || keys.some((key) => !PROFILE_FIELDS.has(key))) return null;

    const patch = {};
    if (Object.prototype.hasOwnProperty.call(data, "name")) patch.name = String(data.name || "").trim();
    if (Object.prototype.hasOwnProperty.call(data, "email")) patch.email = String(data.email || "").trim();
    if (Object.prototype.hasOwnProperty.call(data, "photoURL")) patch.photoURL = String(data.photoURL || "").trim();
    return Object.keys(patch).length ? patch : null;
  }

  function profileSyncMutation(ref, method, args) {
    if (method !== "set" || documentTarget(ref) !== "personale") return null;
    const options = args[1];
    if (!options || options.merge !== true) return null;
    const patch = profilePatchPayload(args[0]);
    if (!patch) return null;
    return {
      path: documentPath(ref),
      signature: JSON.stringify(patch),
      patch
    };
  }

  function updateProfileDeviceCache(profileMutation) {
    const cache = window.HeraRegistryDeviceCache;
    if (!cache || typeof cache.patchRecord !== "function") return;
    const id = profileMutation.path.split("/")[1] || "";
    if (!id) return;
    Promise.resolve(cache.patchRecord("personale", id, profileMutation.patch))
      .then((changed) => {
        if (changed) stats.deviceCacheWrites += 1;
      })
      .catch((error) => {
        console.warn("Aggiornamento profilo nella cache locale non riuscito:", error);
      });
  }

  function wrapMutation(proto, method) {
    if (!proto || !Object.prototype.hasOwnProperty.call(proto, method) ||
        typeof proto[method] !== "function" || proto[method][WRAPPED_MUTATION]) return false;
    const original = proto[method];
    const wrapped = function optimizedRegistryMutation() {
      const ref = this;
      const args = Array.from(arguments);
      const target = documentTarget(ref);
      const profileMutation = profileSyncMutation(ref, method, args);

      if (profileMutation) {
        const previous = recentProfileWrites.get(profileMutation.path);
        if (previous && previous.signature === profileMutation.signature &&
            Date.now() - previous.savedAt <= PROFILE_WRITE_DEDUPE_MS) {
          stats.profileWritesSkipped += 1;
          return previous.promise || Promise.resolve();
        }

        stats.profileWritesPassed += 1;
        const request = Promise.resolve(original.apply(ref, args))
          .then((value) => {
            updateProfileDeviceCache(profileMutation);
            return value;
          })
          .catch((error) => {
            const current = recentProfileWrites.get(profileMutation.path);
            if (current?.signature === profileMutation.signature) {
              recentProfileWrites.delete(profileMutation.path);
            }
            throw error;
          });

        recentProfileWrites.set(profileMutation.path, {
          signature: profileMutation.signature,
          savedAt: Date.now(),
          promise: request
        });
        return request;
      }

      const result = original.apply(ref, args);
      if (!target) return result;
      return Promise.resolve(result).then((value) => {
        invalidate(target);
        return value;
      });
    };
    Object.defineProperty(wrapped, WRAPPED_MUTATION, { value: true });
    Object.defineProperty(wrapped, "__heraOriginal", { value: original });
    proto[method] = wrapped;
    return true;
  }

  function collectFirestorePrototypes() {
    const firestoreApi = window.firebase?.firestore;
    if (!firestoreApi) return [];
    const prototypes = [];
    const add = (proto) => {
      if (proto && !prototypes.includes(proto)) prototypes.push(proto);
    };
    add(firestoreApi.Query?.prototype);
    add(firestoreApi.CollectionReference?.prototype);
    return prototypes;
  }

  function installAvailablePrototypes() {
    installAttempts += 1;
    stats.installAttempts = installAttempts;

    const firestoreApi = window.firebase?.firestore;
    const queryPrototypes = collectFirestorePrototypes();
    let queryPatched = false;
    for (const proto of queryPrototypes) {
      queryPatched = patchQueryPrototype(proto) || queryPatched;
    }

    const documentProto = firestoreApi?.DocumentReference?.prototype;
    const collectionProto = firestoreApi?.CollectionReference?.prototype;
    const mutationPatched = [
      wrapMutation(documentProto, "set"),
      wrapMutation(documentProto, "update"),
      wrapMutation(documentProto, "delete"),
      wrapMutation(collectionProto, "add")
    ].some(Boolean);

    const installed = queryPatched || patchedPrototypes.size > 0;
    if (installed && !stats.installedAt) stats.installedAt = new Date().toISOString();

    api.installed = installed;
    api.pending = !installed;
    api.lastInstallAttemptAt = new Date().toISOString();

    if (!installed && installAttempts < INSTALL_RETRY_LIMIT && !retryTimer) {
      retryTimer = setTimeout(() => {
        retryTimer = null;
        installAvailablePrototypes();
      }, INSTALL_RETRY_MS);
    }

    return installed || mutationPatched;
  }

  const api = {
    installed: false,
    pending: true,
    stats,
    invalidate,
    refreshInstallation: installAvailablePrototypes,
    clearDeviceCache(type) {
      const cache = window.HeraRegistryDeviceCache;
      if (cache && typeof cache.remove === "function") return cache.remove(type);
      return Promise.resolve(false);
    },
    getState() {
      return {
        installed: api.installed,
        pending: api.pending,
        lastInstallAttemptAt: api.lastInstallAttemptAt || "",
        stats: { ...stats },
        inFlight: inFlight.size,
        recent: recent.size,
        profileWriteGuards: recentProfileWrites.size,
        patchedPrototypeCount: patchedPrototypes.size
      };
    }
  };

  window[GLOBAL_NAME] = api;
  installAvailablePrototypes();

  if (typeof window.addEventListener === "function") {
    window.addEventListener("load", installAvailablePrototypes, { once: true });
    window.addEventListener("hera:firebase-ready", installAvailablePrototypes);
  }
})();