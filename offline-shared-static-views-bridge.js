(() => {
  "use strict";

  const GLOBAL = "HeraOfflineSharedStaticViewsBridge";
  const VERSION = "1.0.0";
  const STORAGE_PREFIX = "hera-offline-shared-view-v1:";
  const COLLECTION = "sharedStaticViews";

  if (window[GLOBAL]?.installed) return;

  const state = {
    installed: false,
    cachedDelivered: 0,
    cachedSaved: 0,
    listenersSkippedOffline: 0,
    errors: []
  };

  function safeRead(key) {
    try { return JSON.parse(localStorage.getItem(key) || "null"); } catch (_) { return null; }
  }

  function safeWrite(key, value) {
    try {
      localStorage.setItem(key, JSON.stringify(value));
      return true;
    } catch (_) {
      return false;
    }
  }

  function safeClone(value) {
    try {
      return JSON.parse(JSON.stringify(value, (_key, item) => {
        if (item && typeof item.toDate === "function") {
          try { return item.toDate().toISOString(); } catch (_) {}
        }
        return item;
      }));
    } catch (_) {
      return null;
    }
  }

  function currentUid() {
    try {
      const uid = window.firebase?.auth?.()?.currentUser?.uid;
      if (uid) return String(uid);
    } catch (_) {}
    try {
      const session = JSON.parse(localStorage.getItem("heraPersistedUserSession") || "null");
      return String(session?.uid || session?.user?.uid || "").trim();
    } catch (_) {
      return "";
    }
  }

  function storageKey(path) {
    const uid = currentUid() || "anonymous";
    return `${STORAGE_PREFIX}${uid}:${path}`;
  }

  function documentPath(ref) {
    return String(ref?.path || "").replace(/^\/+|\/+$/g, "");
  }

  function isSharedViewPath(path) {
    return path.startsWith(`${COLLECTION}/`) && path.split("/").length === 2;
  }

  function isOffline() {
    try {
      if (window.HeraOperationalOfflineCache?.isOffline instanceof Function) {
        return Boolean(window.HeraOperationalOfflineCache.isOffline());
      }
    } catch (_) {}
    return navigator.onLine === false;
  }

  function isOptions(value) {
    return Boolean(value && typeof value === "object"
      && typeof value.next !== "function"
      && Object.prototype.hasOwnProperty.call(value, "includeMetadataChanges"));
  }

  function parseObserver(argsLike) {
    const args = Array.from(argsLike || []);
    const index = isOptions(args[0]) ? 1 : 0;
    const candidate = args[index];
    if (typeof candidate === "function") {
      return {
        index,
        next: candidate,
        error: typeof args[index + 1] === "function" ? args[index + 1] : null,
        complete: typeof args[index + 2] === "function" ? args[index + 2] : null,
        context: null,
        args
      };
    }
    if (candidate && typeof candidate === "object") {
      return {
        index,
        next: typeof candidate.next === "function" ? candidate.next : null,
        error: typeof candidate.error === "function" ? candidate.error : null,
        complete: typeof candidate.complete === "function" ? candidate.complete : null,
        context: candidate,
        args
      };
    }
    return { index, next: null, error: null, complete: null, context: null, args };
  }

  function invoke(observer, type, value) {
    const callback = observer?.[type];
    if (typeof callback !== "function") return;
    try { callback.call(observer.context || undefined, value); }
    catch (error) { setTimeout(() => { throw error; }, 0); }
  }

  function buildSnapshot(ref, cached) {
    const exists = cached?.exists !== false && cached?.data && typeof cached.data === "object";
    const data = exists ? cached.data : null;
    const metadata = Object.freeze({ fromCache: true, hasPendingWrites: false });
    return Object.freeze({
      id: String(ref?.id || cached?.id || ""),
      ref,
      exists,
      metadata,
      data: () => exists ? (safeClone(data) || { ...data }) : undefined,
      get: (fieldPath) => {
        if (!exists) return undefined;
        return String(fieldPath || "").split(".").reduce((value, part) => value == null ? undefined : value[part], data);
      },
      isEqual(other) { return Boolean(other && other.id === String(ref?.id || cached?.id || "")); }
    });
  }

  function saveSnapshot(path, snapshot) {
    const payload = {
      id: String(snapshot?.id || path.split("/").pop() || ""),
      exists: Boolean(snapshot?.exists),
      savedAt: new Date().toISOString(),
      data: snapshot?.exists ? (safeClone(snapshot.data?.() || {}) || {}) : null
    };
    if (safeWrite(storageKey(path), payload)) state.cachedSaved += 1;
  }

  function readSnapshot(path) {
    return safeRead(storageKey(path));
  }

  function install() {
    const DocumentPrototype = window.firebase?.firestore?.DocumentReference?.prototype;
    if (!DocumentPrototype || typeof DocumentPrototype.onSnapshot !== "function") return false;
    if (DocumentPrototype.onSnapshot.__heraOfflineSharedStaticViewsWrapped) return true;

    const originalOnSnapshot = DocumentPrototype.onSnapshot;
    const wrapped = function offlineSharedStaticViewOnSnapshot() {
      const path = documentPath(this);
      if (!isSharedViewPath(path)) return originalOnSnapshot.apply(this, arguments);

      const ref = this;
      const observer = parseObserver(arguments);
      const cached = readSnapshot(path);
      let closed = false;
      let unsubscribe = () => {};

      if (cached && observer.next) {
        queueMicrotask(() => {
          if (closed) return;
          invoke(observer, "next", buildSnapshot(ref, cached));
          state.cachedDelivered += 1;
        });
      }

      if (isOffline()) {
        state.listenersSkippedOffline += 1;
        if (!cached && observer.next) {
          queueMicrotask(() => {
            if (!closed) invoke(observer, "next", buildSnapshot(ref, { id: ref.id, exists: false, data: null }));
          });
        }
        return () => { closed = true; };
      }

      const args = Array.from(arguments);
      const originalNext = observer.next;
      if (originalNext) {
        if (typeof args[observer.index] === "function") {
          args[observer.index] = (snapshot) => {
            saveSnapshot(path, snapshot);
            originalNext(snapshot);
          };
        } else {
          const originalObject = args[observer.index] || {};
          args[observer.index] = {
            ...originalObject,
            next(snapshot) {
              saveSnapshot(path, snapshot);
              originalNext.call(originalObject, snapshot);
            }
          };
        }
      }

      try {
        unsubscribe = originalOnSnapshot.apply(ref, args);
      } catch (error) {
        state.errors.push(String(error?.message || error || "listener-error"));
        if (!cached) invoke(observer, "error", error);
      }

      return () => {
        closed = true;
        try { unsubscribe?.(); } catch (_) {}
      };
    };

    Object.defineProperty(wrapped, "__heraOfflineSharedStaticViewsWrapped", { value: true });
    Object.defineProperty(wrapped, "__heraOfflineSharedStaticViewsOriginal", { value: originalOnSnapshot });
    DocumentPrototype.onSnapshot = wrapped;
    state.installed = true;
    return true;
  }

  window[GLOBAL] = {
    installed: false,
    version: VERSION,
    getState: () => ({ ...state, errors: state.errors.slice() })
  };

  if (!install()) {
    let attempts = 0;
    const timer = setInterval(() => {
      attempts += 1;
      if (install() || attempts >= 120) clearInterval(timer);
      window[GLOBAL].installed = state.installed;
    }, 25);
  }
  window[GLOBAL].installed = state.installed;
})();
