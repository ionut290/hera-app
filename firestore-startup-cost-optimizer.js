(() => {
  "use strict";

  const GLOBAL = "HeraFirestoreStartupCostOptimizer";
  const VERSION = "1.0.0";
  if (window[GLOBAL]?.installed) return;

  const PROFILE_WRITE_TTL_MS = 5 * 60 * 1000;
  const STARTUP_LOG_TTL_MS = 5 * 60 * 1000;
  const ACTION_LOG_TTL_MS = 2500;
  const SQUADRE_FALLBACK_MS = 6500;
  const PROFILE_STORAGE_PREFIX = "hera_profile_write_guard_v1:";
  const ACTIVITY_STORAGE_PREFIX = "hera_activity_log_guard_v1:";

  const state = {
    installed: false,
    version: VERSION,
    chatEnabled: false,
    resourcesEnabled: false,
    legacyAlertsEnabled: false,
    blockedChatStarts: 0,
    blockedResourceStarts: 0,
    blockedLegacyAlertStarts: 0,
    staticSquadreStarts: 0,
    staticSquadreSnapshots: 0,
    staticSquadreFallbacks: 0,
    profileWritesSkipped: 0,
    profileWritesPassed: 0,
    activityWritesSkipped: 0,
    activityWritesPassed: 0
  };

  const profilePending = new Map();
  const actionPending = new Map();

  function safeStorageRead(key) {
    try { return JSON.parse(localStorage.getItem(key) || "null"); } catch (_) { return null; }
  }

  function safeStorageWrite(key, value) {
    try { localStorage.setItem(key, JSON.stringify(value)); } catch (_) {}
  }

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

  function documentPath(ref) {
    return String(ref?.path || canonicalPath(ref?._key?.path) || "").replace(/^\/+|\/+$/g, "");
  }

  function collectionName(path) {
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
    let index = 0;
    if (isSnapshotOptions(args[0])) index = 1;
    const candidate = args[index];
    if (typeof candidate === "function") {
      return {
        next: candidate,
        error: typeof args[index + 1] === "function" ? args[index + 1] : null,
        complete: typeof args[index + 2] === "function" ? args[index + 2] : null,
        context: null
      };
    }
    if (candidate && typeof candidate === "object") {
      return {
        next: typeof candidate.next === "function" ? candidate.next : null,
        error: typeof candidate.error === "function" ? candidate.error : null,
        complete: typeof candidate.complete === "function" ? candidate.complete : null,
        context: candidate
      };
    }
    return null;
  }

  function invoke(observer, type, value) {
    const callback = observer?.[type];
    if (typeof callback !== "function") return;
    try { callback.call(observer.context || undefined, value); }
    catch (error) { setTimeout(() => { throw error; }, 0); }
  }

  function todayKey() {
    const now = new Date();
    return `${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, "0")}-${String(now.getDate()).padStart(2, "0")}`;
  }

  function activeSquadreDateKeys() {
    const today = todayKey();
    let selected = today;
    try {
      if (typeof window.getActiveSquadreDateKey === "function") {
        selected = String(window.getActiveSquadreDateKey() || today).slice(0, 10);
      }
    } catch (_) {}
    return [...new Set([selected, today].filter((value) => /^\d{4}-\d{2}-\d{2}$/.test(value)))];
  }

  function groupStaticSquadreRows(views) {
    const groups = new Map();
    for (const view of views.values()) {
      const payload = view?.payload || {};
      const dateKey = String(payload.date || view?.key || todayKey()).slice(0, 10);
      const rows = Array.isArray(payload.squadre) ? payload.squadre : [];
      rows.forEach((row) => {
        if (!row || typeof row !== "object") return;
        const commessaId = String(row.commessaId || row.commessa || row.commessaKey || "").trim();
        if (!commessaId) return;
        const key = `${dateKey}__${commessaId}`;
        if (Array.isArray(row.squadre)) {
          groups.set(key, {
            ...row,
            dateKey: String(row.dateKey || dateKey),
            commessaId
          });
          return;
        }
        const current = groups.get(key) || {
          dateKey,
          commessaId,
          commessaNome: row.commessaNome || row.commessaName || row.nomeCommessa || "",
          squadre: []
        };
        current.squadre.push({ ...row, dateKey, commessaId });
        if (!current.commessaNome) current.commessaNome = row.commessaNome || row.commessaName || row.nomeCommessa || "";
        groups.set(key, current);
      });
    }
    return groups;
  }

  function nestedValue(data, fieldPath) {
    return String(fieldPath || "").split(".").reduce((value, part) => value == null ? undefined : value[part], data);
  }

  function buildStaticSquadreSnapshot(query, views) {
    const groups = groupStaticSquadreRows(views);
    const metadata = Object.freeze({ fromCache: false, hasPendingWrites: false });
    const docs = Array.from(groups.entries()).map(([id, data]) => {
      const ref = query?.firestore?.collection?.("squadreStorico")?.doc?.(id) || null;
      return Object.freeze({
        id,
        ref,
        exists: true,
        metadata,
        data: () => ({ ...data, squadre: Array.isArray(data.squadre) ? data.squadre.map((row) => ({ ...row })) : [] }),
        get: (fieldPath) => nestedValue(data, fieldPath),
        isEqual: (other) => Boolean(other && other.id === id)
      });
    });
    const snapshot = {
      query,
      docs: Object.freeze(docs),
      size: docs.length,
      empty: docs.length === 0,
      metadata,
      forEach(callback, thisArg) { docs.forEach((doc) => callback.call(thisArg, doc)); },
      docChanges() {
        return docs.map((doc, index) => ({ type: "added", doc, oldIndex: -1, newIndex: index }));
      },
      isEqual(other) { return other === snapshot; }
    };
    return Object.freeze(snapshot);
  }

  function subscribeSquadreStatic(query, args, originalOnSnapshot) {
    const observer = parseSnapshotArguments(args);
    const api = window.HeraSharedStaticViews;
    if (!observer?.next || !api?.subscribe) return originalOnSnapshot.apply(query, args);
    try {
      if (typeof window.isSnowServiceContext === "function" && window.isSnowServiceContext()) {
        return originalOnSnapshot.apply(query, args);
      }
    } catch (_) {}

    const dates = activeSquadreDateKeys();
    const views = new Map();
    const unsubs = [];
    let nativeUnsubscribe = null;
    let closed = false;
    let fallbackStarted = false;
    let fallbackTimer = null;
    state.staticSquadreStarts += 1;

    const deliver = () => {
      if (closed || fallbackStarted || !dates.every((date) => views.has(date))) return;
      const snapshot = buildStaticSquadreSnapshot(query, views);
      if (fallbackTimer) {
        clearTimeout(fallbackTimer);
        fallbackTimer = null;
      }
      state.staticSquadreSnapshots += 1;
      invoke(observer, "next", snapshot);
    };

    const startFallback = (reason) => {
      if (closed || fallbackStarted) return;
      fallbackStarted = true;
      state.staticSquadreFallbacks += 1;
      unsubs.splice(0).forEach((unsubscribe) => {
        try { unsubscribe?.(); } catch (_) {}
      });
      console.warn("[COST OPTIMIZER] fallback squadreStorico", reason || "vista condivisa non disponibile");
      nativeUnsubscribe = originalOnSnapshot.apply(query, args);
    };

    dates.forEach((date) => {
      try {
        const unsubscribe = api.subscribe("squadre", date, (view) => {
          if (!view?.payload || !Array.isArray(view.payload.squadre)) return;
          views.set(date, view);
          deliver();
        });
        unsubs.push(unsubscribe);
      } catch (error) {
        startFallback(error?.message || "errore sottoscrizione");
      }
    });

    if (!dates.every((date) => views.has(date))) {
      fallbackTimer = setTimeout(() => startFallback("timeout"), SQUADRE_FALLBACK_MS);
    }
    return () => {
      if (closed) return;
      closed = true;
      clearTimeout(fallbackTimer);
      unsubs.splice(0).forEach((unsubscribe) => {
        try { unsubscribe?.(); } catch (_) {}
      });
      try { nativeUnsubscribe?.(); } catch (_) {}
    };
  }

  function installReadGates() {
    const QueryProto = window.firebase?.firestore?.Query?.prototype;
    if (!QueryProto || typeof QueryProto.onSnapshot !== "function") return false;
    if (QueryProto.onSnapshot.__heraStartupCostOptimizerWrapped) return true;

    const originalOnSnapshot = QueryProto.onSnapshot;
    const wrapped = function optimizedStartupOnSnapshot() {
      const path = queryPath(this);
      const collection = collectionName(path);
      const args = Array.from(arguments);

      if (collection === "chatMessages" && !state.chatEnabled) {
        state.blockedChatStarts += 1;
        return () => {};
      }
      if (collection === "commessaResources" && !state.resourcesEnabled) {
        state.blockedResourceStarts += 1;
        return () => {};
      }
      if (collection === "userAlerts" && !state.legacyAlertsEnabled) {
        const stack = String(new Error().stack || "");
        const fromCentralCenter = /notification-center\.js/i.test(stack);
        if (!fromCentralCenter && /subscribeUserAlerts|app\.js/i.test(stack)) {
          state.blockedLegacyAlertStarts += 1;
          return () => {};
        }
      }
      if (collection === "squadreStorico") {
        return subscribeSquadreStatic(this, args, originalOnSnapshot);
      }
      return originalOnSnapshot.apply(this, args);
    };
    Object.defineProperty(wrapped, "__heraStartupCostOptimizerWrapped", { value: true });
    Object.defineProperty(wrapped, "__heraStartupCostOptimizerOriginal", { value: originalOnSnapshot });
    QueryProto.onSnapshot = wrapped;
    return true;
  }

  function normalizeStable(value, seen = new WeakSet()) {
    if (value == null || typeof value === "string" || typeof value === "number" || typeof value === "boolean") return value;
    if (typeof value !== "object") return String(value);
    if (seen.has(value)) return "[circular]";
    seen.add(value);
    if (Array.isArray(value)) return value.map((item) => normalizeStable(item, seen));
    if (typeof value.toDate === "function" || typeof value.isEqual === "function") return "[firestore-value]";
    const result = {};
    Object.keys(value).sort().forEach((key) => {
      if (["lastSeenAt", "updatedAt", "pushTokenUpdatedAt"].includes(key)) return;
      result[key] = normalizeStable(value[key], seen);
    });
    return result;
  }

  function profileSignature(data) {
    return JSON.stringify(normalizeStable(data));
  }

  function installWriteGuards() {
    const firestoreApi = window.firebase?.firestore;
    const DocumentProto = firestoreApi?.DocumentReference?.prototype;
    const CollectionProto = firestoreApi?.CollectionReference?.prototype;
    if (!DocumentProto || !CollectionProto) return false;

    if (typeof DocumentProto.set === "function" && !DocumentProto.set.__heraProfileWriteGuardWrapped) {
      const originalSet = DocumentProto.set;
      const wrappedSet = function guardedProfileSet(data, options) {
        const path = documentPath(this);
        const isPresencePatch = path.startsWith("platformUsers/")
          && options?.merge === true
          && data && typeof data === "object"
          && Object.prototype.hasOwnProperty.call(data, "lastSeenAt");
        if (!isPresencePatch) return originalSet.apply(this, arguments);

        const signature = profileSignature(data);
        const storageKey = `${PROFILE_STORAGE_PREFIX}${path}`;
        const recent = profilePending.get(path) || safeStorageRead(storageKey);
        if (recent && recent.signature === signature && Date.now() - Number(recent.savedAt || 0) < PROFILE_WRITE_TTL_MS) {
          state.profileWritesSkipped += 1;
          return recent.promise || Promise.resolve();
        }

        state.profileWritesPassed += 1;
        const request = Promise.resolve(originalSet.apply(this, arguments))
          .then((result) => {
            const saved = { signature, savedAt: Date.now() };
            safeStorageWrite(storageKey, saved);
            profilePending.set(path, saved);
            return result;
          })
          .catch((error) => {
            profilePending.delete(path);
            throw error;
          });
        profilePending.set(path, { signature, savedAt: Date.now(), promise: request });
        return request;
      };
      Object.defineProperty(wrappedSet, "__heraProfileWriteGuardWrapped", { value: true });
      DocumentProto.set = wrappedSet;
    }

    if (typeof CollectionProto.add === "function" && !CollectionProto.add.__heraActivityWriteGuardWrapped) {
      const originalAdd = CollectionProto.add;
      const wrappedAdd = function guardedActivityAdd(data) {
        if (collectionName(documentPath(this)) !== "activityLogs" || !data || typeof data !== "object") {
          return originalAdd.apply(this, arguments);
        }
        const type = String(data.actionType || "azione").trim();
        const user = String(data.userId || data.userEmail || "anonymous").trim();
        if (type === "login_app") {
          state.activityWritesSkipped += 1;
          return Promise.resolve(null);
        }

        const isStartup = type === "apertura_app";
        const ttl = isStartup ? STARTUP_LOG_TTL_MS : ACTION_LOG_TTL_MS;
        const signature = isStartup
          ? `${user}|startup`
          : `${user}|${type}|${String(data.actionDescription || "")}|${String(data.commessaId || "")}|${String(data.impiantoId || "")}`;
        const storageKey = `${ACTIVITY_STORAGE_PREFIX}${encodeURIComponent(signature)}`;
        const recent = actionPending.get(signature) || safeStorageRead(storageKey);
        if (recent && Date.now() - Number(recent.savedAt || 0) < ttl) {
          state.activityWritesSkipped += 1;
          return recent.promise || Promise.resolve(null);
        }

        state.activityWritesPassed += 1;
        const request = Promise.resolve(originalAdd.apply(this, arguments))
          .then((result) => {
            const saved = { savedAt: Date.now() };
            actionPending.set(signature, saved);
            safeStorageWrite(storageKey, saved);
            return result;
          })
          .catch((error) => {
            actionPending.delete(signature);
            throw error;
          });
        actionPending.set(signature, { savedAt: Date.now(), promise: request });
        return request;
      };
      Object.defineProperty(wrappedAdd, "__heraActivityWriteGuardWrapped", { value: true });
      CollectionProto.add = wrappedAdd;
    }
    return true;
  }

  function callSubscription(name) {
    const callback = window[name];
    if (typeof callback !== "function") return;
    try { callback(); } catch (error) { console.warn(`[COST OPTIMIZER] ${name}`, error); }
  }

  function enable(kind) {
    if (kind === "chat") {
      if (state.chatEnabled) return;
      state.chatEnabled = true;
      callSubscription("subscribeChat");
      return;
    }
    if (kind === "resources") {
      if (state.resourcesEnabled) return;
      state.resourcesEnabled = true;
      callSubscription("subscribeResources");
      return;
    }
    if (kind === "legacyAlerts") {
      if (state.legacyAlertsEnabled) return;
      state.legacyAlertsEnabled = true;
      callSubscription("subscribeUserAlerts");
    }
  }

  function installLazyTriggers() {
    document.addEventListener("click", (event) => {
      const target = event.target instanceof Element ? event.target : null;
      if (!target) return;
      if (target.closest("#chat-open-btn,[data-open-chat]")) enable("chat");
      if (target.closest("#open-panel-info-utili")) enable("resources");
      if (target.closest("#open-panel-notifiche")) enable("legacyAlerts");
    }, true);
  }

  function install() {
    const reads = installReadGates();
    const writes = installWriteGuards();
    if (!reads || !writes) return false;
    state.installed = true;
    return true;
  }

  const api = {
    installed: false,
    version: VERSION,
    enableChat: () => enable("chat"),
    enableResources: () => enable("resources"),
    enableLegacyAlerts: () => enable("legacyAlerts"),
    refreshInstallation: install,
    getState: () => ({ ...state })
  };
  window[GLOBAL] = api;

  installLazyTriggers();
  let attempts = 0;
  const timer = setInterval(() => {
    attempts += 1;
    if (install() || attempts >= 200) clearInterval(timer);
    api.installed = state.installed;
  }, 25);
  install();
  api.installed = state.installed;
})();
