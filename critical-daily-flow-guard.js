(() => {
  "use strict";

  if (window.__heraCriticalDailyFlowGuardInstalled) return;
  window.__heraCriticalDailyFlowGuardInstalled = true;

  const VERSION = "1.0.0";
  const QUEUE_KEY = "heraPendingOfflineMutations";
  const CONFLICT_KEY = "heraOfflineMutationConflicts:v1";
  const SNOW_MODE_KEY = "heraCriticalOperationalMode:v1";
  const AUTH_RECOVERY_KEY = "heraCriticalAuthRecovery:v1";
  const DUPLICATE_MS = 10 * 60 * 1000;
  const APPROVAL_TIMEOUT_MS = 15_000;
  const MAX_PATCH_ATTEMPTS = 240;
  const state = {
    installed: false,
    deduplicated: 0,
    conflicts: 0,
    deferredSyncs: 0,
    staleActionsBlocked: 0,
    authRecoveries: 0,
    lastError: ""
  };

  let patchAttempts = 0;
  let syncInFlight = null;
  let noteBaseMs = 0;
  let approvalGeneration = 0;
  let expectedAuthUid = "";
  const navigationStack = [];

  const text = (value) => String(value ?? "").trim();
  const email = (value) => text(value).toLowerCase();

  function readJson(key, fallback) {
    try {
      const raw = localStorage.getItem(key);
      return raw ? JSON.parse(raw) : fallback;
    } catch (error) {
      state.lastError = text(error?.message || error);
      return fallback;
    }
  }

  function writeJson(key, value) {
    try {
      localStorage.setItem(key, JSON.stringify(value));
      return true;
    } catch (error) {
      state.lastError = text(error?.message || error);
      return false;
    }
  }

  function toMillis(value) {
    if (!value) return 0;
    try {
      if (typeof value.toMillis === "function") return Number(value.toMillis()) || 0;
      if (typeof value.toDate === "function") return value.toDate().getTime() || 0;
      if (Number.isFinite(Number(value.seconds))) return (Number(value.seconds) * 1000) + Math.floor(Number(value.nanoseconds || 0) / 1e6);
      const parsed = new Date(value).getTime();
      return Number.isFinite(parsed) ? parsed : 0;
    } catch (_) {
      return 0;
    }
  }

  function firebaseUser() {
    try { return window.firebase?.auth?.()?.currentUser || null; } catch (_) { return null; }
  }

  function appUser() {
    try { return typeof currentUser !== "undefined" ? currentUser : null; } catch (_) { return null; }
  }

  function userSnapshot() {
    const user = firebaseUser() || appUser() || {};
    return {
      uid: text(user.uid),
      userEmail: email(user.email),
      operatorName: text(user.displayName || user.email) || "Operatore"
    };
  }

  function commessaSnapshot(payload = {}) {
    let id = text(payload.commessaId);
    let name = text(payload.commessaName);
    try {
      if (!id && typeof selectedCommessaId !== "undefined") id = text(selectedCommessaId);
      if (!name && typeof selectedCommessaName !== "undefined") name = text(selectedCommessaName);
    } catch (_) {}
    return { id, name };
  }

  function stableValue(value, seen = new WeakSet()) {
    if (value === null || value === undefined) return null;
    if (["string", "number", "boolean"].includes(typeof value)) return value;
    if (value instanceof Date) return value.toISOString();
    if (Array.isArray(value)) return value.map((entry) => stableValue(entry, seen));
    if (typeof value !== "object") return String(value);
    if (seen.has(value)) return "[Circular]";
    seen.add(value);
    const output = {};
    Object.keys(value).sort().forEach((key) => {
      if (["createdAt", "updatedAt", "offlineCreatedAt", "status", "syncedFromOfflineQueue"].includes(key)) return;
      output[key] = stableValue(value[key], seen);
    });
    seen.delete(value);
    return output;
  }

  function fingerprint(type, payload, uid) {
    const source = JSON.stringify({ type: text(type), uid: text(uid), payload: stableValue(payload) });
    let hash = 2166136261;
    for (let index = 0; index < source.length; index += 1) {
      hash ^= source.charCodeAt(index);
      hash = Math.imul(hash, 16777619);
    }
    return `${text(type)}:${(hash >>> 0).toString(36)}:${source.length}`;
  }

  function wrapOfflineEnqueue() {
    if (window.enqueueOfflineMutation?.__heraCriticalWrapped) return true;
    if (typeof window.enqueueOfflineMutation !== "function") return false;
    const original = window.enqueueOfflineMutation;
    const wrapped = function enqueueOfflineMutationGuarded(type, payload) {
      const user = userSnapshot();
      const itemFingerprint = fingerprint(type, payload, user.uid);
      const queue = readJson(QUEUE_KEY, []);
      const now = Date.now();
      const duplicate = Array.isArray(queue) && queue.find((item) => {
        const itemTime = toMillis(item.createdAt);
        const savedFingerprint = item.guard?.fingerprint || fingerprint(item.type, item.payload, item.userId || user.uid);
        return item.status !== "synced"
          && (!item.userId || item.userId === user.uid)
          && savedFingerprint === itemFingerprint
          && itemTime > 0
          && now - itemTime <= DUPLICATE_MS;
      });
      if (duplicate) {
        state.deduplicated += 1;
        return duplicate;
      }

      const item = original.call(this, type, payload);
      if (!item?.id) return item;
      const latestQueue = readJson(QUEUE_KEY, []);
      const stored = Array.isArray(latestQueue) ? latestQueue.find((entry) => entry.id === item.id) : null;
      const target = stored || item;
      target.guard = {
        ...(target.guard || {}),
        version: VERSION,
        fingerprint: itemFingerprint,
        baseUpdatedAtMs: type === "commessaNote" && payload?.noteId ? noteBaseMs : 0,
        context: { userId: user.uid, userEmail: user.userEmail, ...commessaSnapshot(payload) }
      };
      if (stored) writeJson(QUEUE_KEY, latestQueue);
      return target;
    };
    Object.defineProperty(wrapped, "__heraCriticalWrapped", { value: true });
    window.enqueueOfflineMutation = wrapped;
    return true;
  }

  function saveConflict(item, serverUpdatedAtMs) {
    const conflicts = readJson(CONFLICT_KEY, []);
    const next = Array.isArray(conflicts) ? conflicts : [];
    next.unshift({
      ...item,
      conflictReason: "server-note-newer-than-offline-edit",
      serverUpdatedAtMs,
      conflictDetectedAt: new Date().toISOString()
    });
    writeJson(CONFLICT_KEY, next.slice(0, 50));
    state.conflicts += 1;
  }

  async function prepareQueuedNotes() {
    const queue = readJson(QUEUE_KEY, []);
    if (!Array.isArray(queue) || !queue.length) return true;
    const user = userSnapshot();
    const candidates = queue.filter((item) => item.type === "commessaNote"
      && item.payload?.noteId
      && item.payload?.commessaId
      && (!item.userId || item.userId === user.uid));
    if (!candidates.length) return true;

    let database;
    try { database = window.firebase?.firestore?.(); } catch (_) { database = null; }
    if (!database) return false;

    const conflicts = new Set();
    try {
      for (const item of candidates) {
        const baseMs = Number(item.guard?.baseUpdatedAtMs || 0);
        if (!baseMs) continue;
        const snapshot = await database.collection("commesse")
          .doc(item.payload.commessaId)
          .collection("noteCommessa")
          .doc(item.payload.noteId)
          .get({ source: "server" });
        const serverMs = snapshot.exists ? toMillis(snapshot.data()?.updatedAt || snapshot.data()?.createdAt) : 0;
        if (serverMs > baseMs) {
          conflicts.add(item.id);
          saveConflict(item, serverMs);
        }
      }
    } catch (error) {
      state.lastError = text(error?.message || error);
      return false;
    }

    if (conflicts.size) writeJson(QUEUE_KEY, queue.filter((item) => !conflicts.has(item.id)));
    return true;
  }

  function wrapOfflineSync() {
    if (window.syncPendingOfflineMutations?.__heraCriticalWrapped) return true;
    if (typeof window.syncPendingOfflineMutations !== "function") return false;
    const original = window.syncPendingOfflineMutations;
    const wrapped = function syncPendingOfflineMutationsGuarded(...args) {
      if (syncInFlight) return syncInFlight;
      syncInFlight = Promise.resolve()
        .then(prepareQueuedNotes)
        .then((ready) => {
          if (!ready) {
            state.deferredSyncs += 1;
            return { deferred: true };
          }
          return original.apply(this, args);
        })
        .finally(() => { syncInFlight = null; });
      return syncInFlight;
    };
    Object.defineProperty(wrapped, "__heraCriticalWrapped", { value: true });
    window.syncPendingOfflineMutations = wrapped;
    return true;
  }

  function wrapNoteEditor() {
    if (window.openCommessaNoteForm?.__heraCriticalWrapped) return true;
    if (typeof window.openCommessaNoteForm !== "function") return false;
    const original = window.openCommessaNoteForm;
    const wrapped = function openCommessaNoteFormGuarded(note, ...args) {
      noteBaseMs = note ? toMillis(note.updatedAt || note.createdAt) : 0;
      return original.call(this, note, ...args);
    };
    Object.defineProperty(wrapped, "__heraCriticalWrapped", { value: true });
    window.openCommessaNoteForm = wrapped;
    return true;
  }

  function activeNavigation() {
    return navigationStack[navigationStack.length - 1] || null;
  }

  function navigationStillOwned(context) {
    const liveUid = text(firebaseUser()?.uid || appUser()?.uid);
    if (!context?.user.uid || context.user.uid === liveUid) return true;
    state.staleActionsBlocked += 1;
    return false;
  }

  function navigationText(context) {
    const current = context.impiantoName || "Impianto";
    const area = context.comune || "zona non specificata";
    return `🧭 ${context.user.operatorName} naviga verso ${current}. La squadra è al lavoro nella zona ${area}.`;
  }

  function wrapNavigationSideEffects() {
    if (typeof window.setImpiantoNavigated === "function" && !window.setImpiantoNavigated.__heraCriticalWrapped) {
      const original = window.setImpiantoNavigated;
      const wrapped = function setImpiantoNavigatedGuarded(commessaId, ids, date, operatorName, ...args) {
        const context = activeNavigation();
        if (!context) return original.call(this, commessaId, ids, date, operatorName, ...args);
        if (!navigationStillOwned(context)) return Promise.resolve(false);
        return original.call(this, context.commessa.id || commessaId, ids, date, context.user.operatorName, ...args);
      };
      Object.defineProperty(wrapped, "__heraCriticalWrapped", { value: true });
      window.setImpiantoNavigated = wrapped;
    }

    if (typeof window.sendChatMessage === "function" && !window.sendChatMessage.__heraCriticalWrapped) {
      const original = window.sendChatMessage;
      const wrapped = function sendChatMessageGuarded(payload, ...args) {
        const context = activeNavigation();
        if (!context || payload?.metadata?.type !== "impianto_navigate") return original.call(this, payload, ...args);
        if (!navigationStillOwned(context)) return Promise.resolve(null);
        return original.call(this, {
          ...payload,
          text: navigationText(context),
          metadata: {
            ...(payload.metadata || {}),
            commessaId: context.commessa.id,
            commessaName: context.commessa.name,
            impiantoName: context.impiantoName,
            comune: context.comune
          }
        }, ...args);
      };
      Object.defineProperty(wrapped, "__heraCriticalWrapped", { value: true });
      window.sendChatMessage = wrapped;
    }

    if (typeof window.publishGlobalNotificationEvent === "function" && !window.publishGlobalNotificationEvent.__heraCriticalWrapped) {
      const original = window.publishGlobalNotificationEvent;
      const wrapped = function publishGlobalNotificationEventGuarded(type, payload, ...args) {
        const context = activeNavigation();
        if (!context || type !== "impianto-navigate") return original.call(this, type, payload, ...args);
        if (!navigationStillOwned(context)) return Promise.resolve(null);
        return original.call(this, type, {
          ...(payload || {}),
          body: navigationText(context),
          commessaId: context.commessa.id,
          commessaName: context.commessa.name,
          impiantoName: context.impiantoName
        }, ...args);
      };
      Object.defineProperty(wrapped, "__heraCriticalWrapped", { value: true });
      window.publishGlobalNotificationEvent = wrapped;
    }
  }

  function wrapNavigation() {
    wrapNavigationSideEffects();
    if (window.navigateToImpianto?.__heraCriticalWrapped) return true;
    if (typeof window.navigateToImpianto !== "function") return false;
    const original = window.navigateToImpianto;
    const wrapped = function navigateToImpiantoGuarded(impianto, ...args) {
      const context = {
        token: `${Date.now()}:${Math.random()}`,
        user: userSnapshot(),
        commessa: commessaSnapshot(),
        impiantoName: text(impianto?.denominazione || impianto?.nome) || "Impianto",
        comune: text(impianto?.comune || impianto?.competenza || impianto?.zona || impianto?.indirizzo) || "zona non specificata"
      };
      navigationStack.push(context);
      return Promise.resolve(original.call(this, impianto, ...args)).finally(() => {
        const index = navigationStack.findIndex((item) => item.token === context.token);
        if (index >= 0) navigationStack.splice(index, 1);
      });
    };
    Object.defineProperty(wrapped, "__heraCriticalWrapped", { value: true });
    window.navigateToImpianto = wrapped;
    return true;
  }

  function timeout(promise, ms) {
    return Promise.race([
      promise,
      new Promise((resolve) => setTimeout(() => resolve({ allowed: false, timeout: true, status: "verification-timeout" }), ms))
    ]);
  }

  function patchApproval() {
    const api = window.HeraAccessApproval;
    if (api?.__heraCriticalWrapped) return true;
    if (!api || typeof api.verify !== "function") return false;
    const original = api.verify.bind(api);
    api.verify = function verifyGuarded(user) {
      const generation = ++approvalGeneration;
      const uid = text(user?.uid);
      return timeout(Promise.resolve().then(() => original(user)), APPROVAL_TIMEOUT_MS).then((result) => {
        const liveUid = text(firebaseUser()?.uid || appUser()?.uid);
        if (generation !== approvalGeneration || (uid && liveUid && uid !== liveUid)) {
          return { allowed: false, stale: true, status: "stale-auth-event" };
        }
        return result;
      });
    };
    Object.defineProperty(api, "__heraCriticalWrapped", { value: true });
    return true;
  }

  function restoreSnowMode() {
    const saved = readJson(SNOW_MODE_KEY, null);
    if (!saved || saved.mode !== "snow" || Date.now() - Number(saved.at || 0) > 30 * 86400000) return;
    const apply = () => {
      document.body?.classList.add("snow-management-context");
      if (!location.hash && /^#[^\s]{1,180}$/.test(saved.hash || "")) location.hash = saved.hash;
    };
    if (document.body) apply();
    else document.addEventListener("DOMContentLoaded", apply, { once: true });
  }

  function saveSnowMode() {
    if (!document.body?.classList.contains("snow-management-context")) return;
    writeJson(SNOW_MODE_KEY, { mode: "snow", uid: userSnapshot().uid, hash: location.hash || "#servizio-neve", at: Date.now() });
  }

  function watchSnowMode() {
    const update = () => {
      if (document.body?.classList.contains("snow-management-context")) saveSnowMode();
      else localStorage.removeItem(SNOW_MODE_KEY);
    };
    addEventListener("hashchange", () => setTimeout(update, 0));
    addEventListener("pagehide", saveSnowMode);
    if (document.body && typeof MutationObserver === "function") {
      new MutationObserver(update).observe(document.body, { attributes: true, attributeFilter: ["class"] });
    }
  }

  function recoverAuthMismatch(expected, actual) {
    try {
      const previous = readJson(AUTH_RECOVERY_KEY, null);
      if (previous && previous.expected === expected && previous.actual === actual && Date.now() - previous.at < 30000) return;
      sessionStorage.setItem(AUTH_RECOVERY_KEY, JSON.stringify({ expected, actual, at: Date.now() }));
      state.authRecoveries += 1;
      const url = new URL(location.href);
      url.searchParams.set("authRecovery", String(Date.now()));
      location.replace(url.toString());
    } catch (_) {}
  }

  function watchAuth() {
    try {
      const auth = window.firebase?.auth?.();
      if (!auth || auth.__heraCriticalObserver) return;
      auth.onAuthStateChanged((user) => {
        expectedAuthUid = text(user?.uid);
        approvalGeneration += 1;
        const savedSnow = readJson(SNOW_MODE_KEY, null);
        if (savedSnow?.uid && expectedAuthUid && savedSnow.uid !== expectedAuthUid) {
          localStorage.removeItem(SNOW_MODE_KEY);
          document.body?.classList.remove("snow-management-context");
          if (/^#(?:servizio-neve|fuel=)/.test(location.hash)) location.hash = "";
        }
        [500, 1500, 4000].forEach((delay) => setTimeout(() => {
          let resolved = true;
          try { if (typeof authStateResolved !== "undefined") resolved = Boolean(authStateResolved); } catch (_) {}
          if (!resolved) return;
          const actual = text(appUser()?.uid);
          if (actual !== expectedAuthUid && !(appUser()?.persistedOnly && expectedAuthUid)) recoverAuthMismatch(expectedAuthUid, actual);
        }, delay));
      });
      Object.defineProperty(auth, "__heraCriticalObserver", { value: true });
    } catch (_) {}
  }

  function install() {
    patchAttempts += 1;
    const results = [wrapOfflineEnqueue(), wrapOfflineSync(), wrapNoteEditor(), wrapNavigation(), patchApproval()];
    state.installed = results.slice(0, 4).every(Boolean);
    if (!state.installed && patchAttempts < MAX_PATCH_ATTEMPTS) setTimeout(install, 50);
    return state.installed;
  }

  restoreSnowMode();
  watchAuth();
  const start = () => { watchSnowMode(); install(); };
  if (document.readyState === "loading") document.addEventListener("DOMContentLoaded", start, { once: true });
  else start();
  addEventListener("load", install, { once: true });

  window.HeraCriticalDailyFlowGuard = {
    version: VERSION,
    install,
    prepareQueuedNotes,
    getState: () => ({ ...state, activeNavigations: navigationStack.length }),
    getConflicts: () => readJson(CONFLICT_KEY, [])
  };
})();
