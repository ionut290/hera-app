(() => {
  "use strict";

  const GLOBAL_NAME = "VargaFirestoreNestedListenerOptimizer";
  const WRAPPED_FLAG = "__vargaNestedListenerOptimizerWrapped";
  const ORIGINAL_FLAG = "__vargaNestedListenerOptimizerOriginal";
  const VERSION = "1.0.0";
  const RETRY_MS = 25;
  const RETRY_LIMIT = 800;
  const RELEASE_GRACE_MS = 2500;
  const TARGET_PATH = /^commesse\/[^/]+\/impianti$/;

  if (window[GLOBAL_NAME]?.version === VERSION && window[GLOBAL_NAME]?.installed) return;

  const groups = new Map();
  let installAttempts = 0;
  let installTimer = null;

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

  function scheduleInstall() {
    if (installAttempts >= RETRY_LIMIT || installTimer) return;
    installTimer = window.setTimeout(() => {
      installTimer = null;
      install();
    }, RETRY_MS);
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
        complete: typeof args[index + 2] === "function" ? args[index + 2] : null,
        context: undefined
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

  function queryInfo(query, options) {
    const path = String(query?.path || "").replace(/^\/+|\/+$/g, "");
    if (!TARGET_PATH.test(path)) return null;

    // Condividiamo esclusivamente CollectionReference pure degli impianti.
    // Query con where/orderBy/limit non espongono `path` nel SDK compat e
    // mantengono quindi il comportamento Firestore originale.
    return {
      path,
      key: `${path}|metadata:${options?.includeMetadataChanges === true ? 1 : 0}`
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
      callback.call(subscriber.observer.context, value);
    } catch (error) {
      reportCallbackError(error);
    }
  }

  function fanOut(group, type, value) {
    for (const subscriber of Array.from(group.subscribers)) {
      invokeSubscriber(subscriber, type, value);
    }
  }

  function closePhysicalGroup(group, reason) {
    if (!group || group.closed) return;
    group.closed = true;
    if (group.releaseTimer) {
      window.clearTimeout(group.releaseTimer);
      group.releaseTimer = null;
    }
    groups.delete(group.key);
    try {
      if (typeof group.unsubscribeNative === "function") group.unsubscribeNative();
    } catch (error) {
      console.warn("Chiusura listener impianti condiviso non riuscita:", error);
    }
    stats.physicalListenersClosed += 1;
    group.closedReason = reason || "closed";
  }

  function schedulePhysicalRelease(group) {
    if (!group || group.closed || group.subscribers.size || group.releaseTimer) return;
    group.releaseTimer = window.setTimeout(() => {
      group.releaseTimer = null;
      if (!group.subscribers.size) closePhysicalGroup(group, "grace-expired");
    }, RELEASE_GRACE_MS);
  }

  function createPhysicalGroup(query, info, parsed, currentOnSnapshot) {
    const group = {
      key: info.key,
      info,
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
      group.unsubscribeNative = parsed.options
        ? currentOnSnapshot.call(query, parsed.options, nativeObserver)
        : currentOnSnapshot.call(query, nativeObserver);
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
      window.clearTimeout(group.releaseTimer);
      group.releaseTimer = null;
      stats.gracePeriodReuses += 1;
    }

    const subscriber = { observer: parsed.observer, closed: false };
    group.subscribers.add(subscriber);
    stats.logicalSubscriptions += 1;

    if (group.hasSnapshot) {
      const replay = () => invokeSubscriber(subscriber, "next", group.latestSnapshot);
      if (typeof queueMicrotask === "function") queueMicrotask(replay);
      else Promise.resolve().then(replay);
    }

    return function unsubscribeSharedImpiantiListener() {
      if (subscriber.closed) return;
      subscriber.closed = true;
      group.subscribers.delete(subscriber);
      schedulePhysicalRelease(group);
    };
  }

  function install() {
    installAttempts += 1;
    stats.installAttempts = installAttempts;

    const firestoreApi = window.firebase?.firestore;
    const QueryProto = firestoreApi?.Query?.prototype;
    const safeOptimizerReady = Boolean(window.VargaFirestoreSafeOptimizer?.installed);
    if (!QueryProto || typeof QueryProto.onSnapshot !== "function" || !safeOptimizerReady) {
      scheduleInstall();
      return false;
    }

    const current = QueryProto.onSnapshot;
    if (current[WRAPPED_FLAG]) {
      api.installed = true;
      return true;
    }

    const originalOnSnapshot = current[ORIGINAL_FLAG] || current;
    const wrapped = function optimizedNestedOnSnapshot() {
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
    console.info("Ottimizzatore listener impianti condivisi attivo.");
    return true;
  }

  function clear() {
    for (const group of Array.from(groups.values())) {
      closePhysicalGroup(group, "manual-clear");
    }
  }

  const api = {
    version: VERSION,
    installed: false,
    enabled: true,
    target: "commesse/{id}/impianti",
    stats,
    refreshInstallation: install,
    pendingCount: () => groups.size,
    clear,
    getState() {
      return {
        version: VERSION,
        installed: api.installed,
        enabled: true,
        target: api.target,
        activeGroups: groups.size,
        groups: Array.from(groups.values()).map((group) => ({
          path: group.info.path,
          subscribers: group.subscribers.size,
          hasSnapshot: group.hasSnapshot,
          waitingForRelease: Boolean(group.releaseTimer)
        })),
        stats: { ...stats }
      };
    }
  };

  window[GLOBAL_NAME] = api;
  install();
  window.addEventListener?.("load", install, { once: true });
})();
