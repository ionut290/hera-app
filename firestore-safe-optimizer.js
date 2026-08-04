(() => {
  "use strict";

  const GLOBAL_NAME = "VargaFirestoreSafeOptimizer";
  const WRAPPED_FLAG = "__vargaSharedListenerOptimizerWrapped";
  const ORIGINAL_FLAG = "__vargaSharedListenerOptimizerOriginal";
  const VERSION = "3.0.0";
  const RETRY_MS = 25;
  const RETRY_LIMIT = 800;
  const RELEASE_GRACE_MS = 2500;

  // Solo raccolte per le quali il report diagnostico ha mostrato aperture
  // duplicate o una chiusura/riapertura immediata. Le query diverse non
  // vengono mai unite: la chiave comprende anche filtri, ordinamenti e limiti.
  const TARGET_COLLECTIONS = new Set([
    "commesse",
    "squadreStorico",
    "userAlerts"
  ]);

  if (window[GLOBAL_NAME]?.version === VERSION && window[GLOBAL_NAME]?.installed) return;

  const groups = new Map();
  let installAttempts = 0;
  let installTimer = null;
  let registryLoaderStarted = false;

  const stats = {
    installedAt: "",
    installAttempts: 0,
    logicalSubscriptions: 0,
    physicalListenersStarted: 0,
    preventedListenerStarts: 0,
    gracePeriodReuses: 0,
    physicalListenersClosed: 0,
    snapshotsDelivered: 0,
    errorsDelivered: 0,
    bypassedSubscriptions: 0
  };

  function canonicalPath(value) {
    if (!value) return "";
    if (typeof value === "string") return value;
    if (typeof value.canonicalString === "function") {
      try { return String(value.canonicalString() || ""); } catch (_) {}
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
    const candidates = [
      query?.path,
      query?._query?.path,
      query?._query?._path,
      query?.Ae?.path,
      query?.je?.path,
      query?._delegate?._query?.path,
      query?._delegate?._query?._path
    ];
    for (const candidate of candidates) {
      const path = canonicalPath(candidate).replace(/^\/+|\/+$/g, "");
      if (path) return path;
    }
    return "";
  }

  function canonicalQuery(query) {
    const candidates = [
      query?._query,
      query?.Ae,
      query?.je,
      query?._delegate?._query
    ].filter(Boolean);

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

  function firstCollection(path) {
    return String(path || "").split("/")[0] || "";
  }

  function isSnapshotOptions(value) {
    return Boolean(
      value &&
      typeof value === "object" &&
      typeof value.next !== "function" &&
      Object.prototype.hasOwnProperty.call(value, "includeMetadataChanges")
    );
  }

  function parseSnapshotArguments(argsLike) {
    const args = Array.from(argsLike || []);
    let options = null;
    let index = 0;

    if (isSnapshotOptions(args[0])) {
      options = args[0];
      index = 1;
    }

    const candidate = args[index];
    let observer;
    if (typeof candidate === "function") {
      observer = {
        next: candidate,
        error: typeof args[index + 1] === "function" ? args[index + 1] : null,
        complete: typeof args[index + 2] === "function" ? args[index + 2] : null
      };
    } else if (candidate && typeof candidate === "object") {
      observer = {
        next: typeof candidate.next === "function" ? candidate.next : null,
        error: typeof candidate.error === "function" ? candidate.error : null,
        complete: typeof candidate.complete === "function" ? candidate.complete : null,
        context: candidate
      };
    } else {
      return null;
    }

    if (!observer.next && !observer.error && !observer.complete) return null;
    return { options, observer };
  }

  function stableOptions(options) {
    return options?.includeMetadataChanges === true ? "metadata:1" : "metadata:0";
  }

  function queryInfo(query, options) {
    const path = queryPath(query);
    const collection = firstCollection(path);
    if (!TARGET_COLLECTIONS.has(collection)) return null;

    const canonical = canonicalQuery(query);
    // Una CollectionReference semplice ha `path`. Per una Query filtrata,
    // invece, pretendiamo sempre un identificatore canonico: così query con
    // filtri o limiti diversi non possono mai essere unite per errore.
    const isPlainCollection = Boolean(query?.path && path === collection);
    if (!canonical && !isPlainCollection) return null;

    return {
      collection,
      path,
      key: `${path}|${canonical || `collection:${path}`}|${stableOptions(options)}`
    };
  }

  function reportCallbackError(error) {
    window.setTimeout(() => { throw error; }, 0);
  }

  function invokeSubscriber(subscriber, type, value) {
    if (!subscriber || subscriber.closed) return;
    const callback = subscriber.observer?.[type];
    if (typeof callback !== "function") return;
    try {
      callback.call(subscriber.observer.context || undefined, value);
    } catch (error) {
      reportCallbackError(error);
    }
  }

  function fanOut(group, type, value) {
    const subscribers = Array.from(group.subscribers);
    for (const subscriber of subscribers) invokeSubscriber(subscriber, type, value);
  }

  function closePhysicalGroup(group, reason) {
    if (!group || group.closed) return;
    group.closed = true;
    if (group.releaseTimer) {
      clearTimeout(group.releaseTimer);
      group.releaseTimer = null;
    }
    groups.delete(group.key);
    try {
      if (typeof group.unsubscribeNative === "function") group.unsubscribeNative();
    } catch (error) {
      console.warn("Chiusura listener Firestore condiviso non riuscita:", error);
    }
    stats.physicalListenersClosed += 1;
    if (reason) group.closedReason = reason;
  }

  function schedulePhysicalRelease(group) {
    if (!group || group.closed || group.subscribers.size || group.releaseTimer) return;
    group.releaseTimer = window.setTimeout(() => {
      group.releaseTimer = null;
      if (!group.subscribers.size) closePhysicalGroup(group, "grace-expired");
    }, RELEASE_GRACE_MS);
  }

  function createPhysicalGroup(query, info, parsed, originalOnSnapshot) {
    const group = {
      key: info.key,
      info,
      query,
      options: parsed.options,
      subscribers: new Set(),
      unsubscribeNative: null,
      releaseTimer: null,
      latestSnapshot: null,
      hasSnapshot: false,
      closed: false,
      closedReason: ""
    };

    const nativeObserver = {
      next(snapshot) {
        if (group.closed) return;
        group.latestSnapshot = snapshot;
        group.hasSnapshot = true;
        stats.snapshotsDelivered += 1;
        fanOut(group, "next", snapshot);
      },
      error(error) {
        if (group.closed) return;
        stats.errorsDelivered += 1;
        fanOut(group, "error", error);
        closePhysicalGroup(group, "native-error");
      },
      complete() {
        if (group.closed) return;
        fanOut(group, "complete");
        closePhysicalGroup(group, "native-complete");
      }
    };

    groups.set(group.key, group);
    try {
      group.unsubscribeNative = group.options
        ? originalOnSnapshot.call(query, group.options, nativeObserver)
        : originalOnSnapshot.call(query, nativeObserver);
      stats.physicalListenersStarted += 1;
      return group;
    } catch (error) {
      groups.delete(group.key);
      group.closed = true;
      throw error;
    }
  }

  function subscribeToGroup(group, parsed) {
    if (group.releaseTimer) {
      clearTimeout(group.releaseTimer);
      group.releaseTimer = null;
      stats.gracePeriodReuses += 1;
    }

    const subscriber = {
      observer: parsed.observer,
      closed: false
    };
    group.subscribers.add(subscriber);
    stats.logicalSubscriptions += 1;

    if (group.hasSnapshot) {
      const replay = () => invokeSubscriber(subscriber, "next", group.latestSnapshot);
      if (typeof queueMicrotask === "function") queueMicrotask(replay);
      else Promise.resolve().then(replay);
    }

    return function unsubscribeSharedListener() {
      if (subscriber.closed) return;
      subscriber.closed = true;
      group.subscribers.delete(subscriber);
      schedulePhysicalRelease(group);
    };
  }

  function loadScriptOnce(src, marker, readyCheck, onLoad) {
    if (typeof readyCheck === "function" && readyCheck()) {
      if (typeof onLoad === "function") onLoad();
      return;
    }
    const selector = `script[data-${marker}]`;
    const existing = document.querySelector(selector);
    if (existing) {
      if (typeof onLoad === "function") existing.addEventListener("load", onLoad, { once: true });
      return;
    }
    const script = document.createElement("script");
    script.src = src;
    script.setAttribute(`data-${marker}`, "1");
    if (typeof onLoad === "function") script.addEventListener("load", onLoad, { once: true });
    script.addEventListener("error", () => {
      console.warn(`Modulo Firestore non caricato: ${src}`);
    }, { once: true });
    document.head.appendChild(script);
  }

  function ensureRegistryOptimizerLoaded() {
    if (registryLoaderStarted) return;
    registryLoaderStarted = true;

    const loadOptimizer = () => loadScriptOnce(
      "./firestore-registry-read-optimizer.js?v=20260805a",
      "hera-registry-read-optimizer",
      () => Boolean(window.HeraFirestoreRegistryOptimizer?.installed)
    );

    loadScriptOnce(
      "./registry-device-cache.js?v=20260805a",
      "hera-registry-device-cache",
      () => Boolean(window.HeraRegistryDeviceCache),
      loadOptimizer
    );
    window.setTimeout(loadOptimizer, 500);
  }

  function install() {
    installAttempts += 1;
    stats.installAttempts = installAttempts;

    const firestoreApi = window.firebase?.firestore;
    const QueryProto = firestoreApi?.Query?.prototype;
    if (!QueryProto || typeof QueryProto.onSnapshot !== "function") {
      if (installAttempts < RETRY_LIMIT && !installTimer) {
        installTimer = window.setTimeout(() => {
          installTimer = null;
          install();
        }, RETRY_MS);
      }
      return false;
    }

    const current = QueryProto.onSnapshot;
    if (current[WRAPPED_FLAG]) {
      api.installed = true;
      ensureRegistryOptimizerLoaded();
      return true;
    }

    const originalOnSnapshot = current[ORIGINAL_FLAG] || current;
    const wrapped = function sharedOptimizedOnSnapshot() {
      const parsed = parseSnapshotArguments(arguments);
      const info = parsed ? queryInfo(this, parsed.options) : null;
      if (!parsed || !info) {
        stats.bypassedSubscriptions += 1;
        return current.apply(this, arguments);
      }

      let group = groups.get(info.key);
      if (!group || group.closed) {
        group = createPhysicalGroup(this, info, parsed, current);
      } else {
        stats.preventedListenerStarts += 1;
      }
      return subscribeToGroup(group, parsed);
    };

    Object.defineProperty(wrapped, WRAPPED_FLAG, { value: true });
    Object.defineProperty(wrapped, ORIGINAL_FLAG, { value: originalOnSnapshot });
    QueryProto.onSnapshot = wrapped;

    stats.installedAt = new Date().toISOString();
    api.installed = true;
    ensureRegistryOptimizerLoaded();
    console.info("Ottimizzatore Firestore sicuro attivo:", Array.from(TARGET_COLLECTIONS));
    return true;
  }

  function clear() {
    for (const group of Array.from(groups.values())) closePhysicalGroup(group, "manual-clear");
  }

  const api = {
    version: VERSION,
    enabled: true,
    installed: false,
    targets: Object.freeze(Array.from(TARGET_COLLECTIONS)),
    stats,
    refreshInstallation: install,
    pendingCount: () => groups.size,
    clear,
    getState() {
      return {
        version: VERSION,
        enabled: true,
        installed: api.installed,
        targets: Array.from(TARGET_COLLECTIONS),
        activeGroups: groups.size,
        groups: Array.from(groups.values()).map((group) => ({
          collection: group.info.collection,
          path: group.info.path,
          subscribers: group.subscribers.size,
          hasSnapshot: group.hasSnapshot,
          waitingForRelease: Boolean(group.releaseTimer)
        })),
        stats: { ...stats }
      };
    }
  };

  window.__vargaFirestoreSafeOptimizer = true;
  window[GLOBAL_NAME] = api;
  install();
  window.addEventListener?.("load", install, { once: true });
})();
