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
  const DEVICE_CACHE_MAX_AGE_MS = Object.freeze({
    personale: 24 * 60 * 60 * 1000,
    mezzi: 12 * 60 * 60 * 1000
  });
  const PROFILE_WRITE_DEDUPE_MS = 60 * 1000;
  const PROFILE_FIELDS = new Set(["name", "email", "photoURL", "updatedAt"]);
  const inFlight = new Map();
  const recent = new Map();
  const recentProfileWrites = new Map();
  const generations = new Map(Object.keys(TARGET_TTL_MS).map((name) => [name, 0]));
  const stats = {
    installedAt: new Date().toISOString(),
    networkGets: 0,
    reusedInFlight: 0,
    reusedRecent: 0,
    reusedDeviceCache: 0,
    deviceCacheWrites: 0,
    listenerSnapshots: 0,
    invalidations: 0,
    profileWritesPassed: 0,
    profileWritesSkipped: 0
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

    if (!query?.path && !canonical) return null;

    return {
      target,
      path,
      isFullCollection: Boolean(query?.path),
      key: `${target}|${canonical || `collection:${path}`}`
    };
  }

  function documentPath(ref) {
    return String(ref?.path || canonicalPath(ref?._key?.path) || "").replace(/^\/+|\/+$/g, "");
  }

  function documentTarget(ref) {
    const first = documentPath(ref).split("/")[0] || "";
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

  function extractSnapshotRecords(snapshot) {
    if (!Array.isArray(snapshot?.docs)) return [];
    const records = [];
    for (const document of snapshot.docs) {
      const id = String(document?.id || "").trim();
      if (!id || typeof document?.data !== "function") return [];
      const data = document.data() || {};
      records.push({ id, ...data });
    }
    return records;
  }

  function persistSnapshotToDevice(info, snapshot) {
    if (!info?.isFullCollection) return;
    const cache = window.HeraRegistryDeviceCache;
    if (!cache || typeof cache.writeIfChanged !== "function") return;
    const records = extractSnapshotRecords(snapshot);
    if (!records.length) return;
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
    return String(record?.id || record?.docId || record?._id || "").trim();
  }

  function cachedRecordData(record) {
    const data = { ...(record || {}) };
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
    const firestore = query?.firestore;
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
      console.warn(`Cache locale ${info.target} non leggibile:`, error);
      return null;
    }
  }

  const originalGet = QueryProto.get;
  QueryProto.get = function optimizedRegistryGet(options) {
    const query = this;
    const args = arguments;
    const info = queryInfo(query);
    if (!info || options?.source) {
      return originalGet.apply(query, args);
    }

    const key = `${info.key}|source:default`;
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
    const request = Promise.resolve()
      .then(() => getDeviceCachedSnapshot(query, info))
      .then((cachedSnapshot) => {
        if (cachedSnapshot) {
          stats.reusedDeviceCache += 1;
          remember(info, cachedSnapshot, "device-cache");
          return cachedSnapshot;
        }

        stats.networkGets += 1;
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
    if (!proto || typeof proto[method] !== "function") return;
    const original = proto[method];
    proto[method] = function optimizedRegistryMutation() {
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

        // Nome, email e foto vengono già propagati dal listener `personale`.
        // Non svuotiamo la cache dell'intera collezione per questa sola patch:
        // il prossimo snapshot sostituirà automaticamente la copia recente.
        return request;
      }

      const result = original.apply(ref, args);
      if (!target) return result;
      return Promise.resolve(result).then((value) => {
        invalidate(target);
        return value;
      });
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
    clearDeviceCache(type) {
      const cache = window.HeraRegistryDeviceCache;
      if (cache && typeof cache.remove === "function") return cache.remove(type);
      return Promise.resolve(false);
    },
    getState() {
      return {
        stats: { ...stats },
        inFlight: inFlight.size,
        recent: recent.size,
        profileWriteGuards: recentProfileWrites.size
      };
    }
  };
})();
