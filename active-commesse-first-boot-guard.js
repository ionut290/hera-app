(() => {
  "use strict";

  const GLOBAL = "HeraActiveCommesseFirstBootGuard";
  const VERSION = "1.1.0";
  const INDEX_PATH = "appConfig/activeCommesse";
  const MAX_ACTIVE_IDS_PER_QUERY = 30;
  const INDEX_READ_TIMEOUT_MS = 5000;

  if (window[GLOBAL]?.installed) return;

  const state = {
    installed: false,
    version: VERSION,
    authWaits: 0,
    deferredIndexReads: 0,
    filteredRootListeners: 0,
    emptyRootListeners: 0,
    failOpenRootListeners: 0,
    indexTimeouts: 0,
    blockedLegacyAlertListeners: 0,
    lastError: ""
  };

  let inFlightIndexRead = null;
  let fallbackStatePromise = null;
  const bypassQueries = new WeakSet();

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
      query?._delegate?._query?.path,
      query?._delegate?._query?._path,
      query?.Ae?.path,
      query?.je?.path
    ];
    for (const candidate of candidates) {
      const path = canonicalPath(candidate).replace(/^\/+|\/+$/g, "");
      if (path) return path;
    }
    return "";
  }

  function documentPath(ref) {
    return String(ref?.path || canonicalPath(ref?._key?.path) || "")
      .replace(/^\/+|\/+$/g, "");
  }

  function normalizeIds(value) {
    return [...new Set((Array.isArray(value) ? value : [])
      .map((id) => String(id || "").trim())
      .filter(Boolean))];
  }

  function waitForAuthenticatedUser() {
    let auth;
    try {
      auth = window.firebase?.auth?.();
    } catch (error) {
      return Promise.reject(error);
    }

    if (!auth || typeof auth.onAuthStateChanged !== "function") {
      return Promise.reject(new Error("Firebase Auth non disponibile"));
    }
    if (auth.currentUser) return Promise.resolve(auth.currentUser);

    state.authWaits += 1;
    return new Promise((resolve, reject) => {
      let unsubscribe = () => {};
      const finish = (callback, value) => {
        try { unsubscribe(); } catch (_) {}
        callback(value);
      };
      unsubscribe = auth.onAuthStateChanged(
        (user) => {
          if (!user) return;
          finish(resolve, user);
        },
        (error) => finish(reject, error)
      );
    });
  }

  function installIndexReadGuard() {
    const DocumentPrototype = window.firebase?.firestore?.DocumentReference?.prototype;
    if (!DocumentPrototype || typeof DocumentPrototype.get !== "function") return false;
    if (DocumentPrototype.get.__heraActiveCommesseAuthGuardWrapped) return true;

    const originalGet = DocumentPrototype.get;
    const wrappedGet = function activeCommesseAuthAwareGet(options) {
      if (documentPath(this) !== INDEX_PATH || options?.source === "server") {
        return originalGet.apply(this, arguments);
      }

      if (window.firebase?.auth?.().currentUser) {
        return originalGet.apply(this, arguments);
      }

      if (inFlightIndexRead) return inFlightIndexRead;

      const reference = this;
      const args = Array.from(arguments);
      state.deferredIndexReads += 1;
      const request = waitForAuthenticatedUser()
        .then(() => originalGet.apply(reference, args))
        .catch((error) => {
          state.lastError = String(error?.message || error || "");
          throw error;
        });
      inFlightIndexRead = request;
      request.then(
        () => { if (inFlightIndexRead === request) inFlightIndexRead = null; },
        () => { if (inFlightIndexRead === request) inFlightIndexRead = null; }
      );
      return request;
    };

    Object.defineProperty(wrappedGet, "__heraActiveCommesseAuthGuardWrapped", { value: true });
    Object.defineProperty(wrappedGet, "__heraActiveCommesseAuthGuardOriginal", { value: originalGet });
    DocumentPrototype.get = wrappedGet;
    return true;
  }

      function currentGuardHasRootFiltering() {
    const api = window.HeraActiveCommesse;
    return Boolean(api?.installed && typeof api.loadAllForManager === "function");
  }

  function loadFallbackState(query) {
    if (fallbackStatePromise) return fallbackStatePromise;
    const firestore = query?.firestore;
    if (!firestore?.collection) {
      return Promise.resolve({ explicit: false, ids: [] });
    }

    const indexRead = firestore.collection("appConfig").doc("activeCommesse").get();
    const timeout = new Promise((resolve) => window.setTimeout(() => {
      state.indexTimeouts += 1;
      resolve({ __timeout: true });
    }, INDEX_READ_TIMEOUT_MS));

    fallbackStatePromise = Promise.race([indexRead, timeout])
      .then((snapshot) => {
        if (snapshot?.__timeout) return { explicit: false, ids: [] };
        if (!snapshot?.exists) return { explicit: false, ids: [] };
        const data = snapshot.data?.() || {};
        if (!Array.isArray(data.ids)) return { explicit: false, ids: [] };
        return { explicit: true, ids: normalizeIds(data.ids) };
      })
      .catch((error) => {
        state.lastError = String(error?.message || error || "");
        console.warn("[ACTIVE COMMESSE FIRST BOOT] indice non disponibile; uso la query originale.", error);
        return { explicit: false, ids: [] };
      });
    return fallbackStatePromise;
  }

  function shouldBlockLegacyUserAlerts(path) {
    if (path !== "userAlerts") return false;
    const stack = String(new Error().stack || "");
    const fromCentralCenter = /notification-center\.js/i.test(stack);
    const fromLegacyApp = /subscribeUserAlerts|app\.js/i.test(stack);
    return fromLegacyApp && !fromCentralCenter;
  }

  function installListenerGuard() {
    const QueryPrototype = window.firebase?.firestore?.Query?.prototype;
    if (!QueryPrototype || typeof QueryPrototype.onSnapshot !== "function") return false;
    if (QueryPrototype.onSnapshot.__heraActiveCommesseFirstBootWrapped) return true;

    const originalOnSnapshot = QueryPrototype.onSnapshot;
    const wrappedOnSnapshot = function activeCommesseFirstBootOnSnapshot() {
      const path = queryPath(this);
      const args = Array.from(arguments);

      if (shouldBlockLegacyUserAlerts(path)) {
        state.blockedLegacyAlertListeners += 1;
        return () => {};
      }

      if (path !== "commesse" || bypassQueries.has(this) || currentGuardHasRootFiltering()) {
        return originalOnSnapshot.apply(this, args);
      }

      const originalQuery = this;
      let closed = false;
      let unsubscribe = () => {};

      loadFallbackState(originalQuery).then(({ explicit, ids }) => {
        if (closed) return;
        if (!explicit) {
          unsubscribe = originalOnSnapshot.apply(originalQuery, args);
          return;
        }
        if (!ids.length) {
          // Fail open: un indice vuoto può essere temporaneo o non ancora sincronizzato.
          // La raccolta originale è la fonte autorevole e impedisce una Home vuota.
          state.emptyRootListeners += 1;
          state.failOpenRootListeners += 1;
          console.warn("[ACTIVE COMMESSE FIRST BOOT] indice vuoto; uso la raccolta commesse originale.");
          unsubscribe = originalOnSnapshot.apply(originalQuery, args);
          return;
        }
        if (ids.length > MAX_ACTIVE_IDS_PER_QUERY) {
          console.warn(
            `[ACTIVE COMMESSE FIRST BOOT] ${ids.length} ID attivi superano il limite ${MAX_ACTIVE_IDS_PER_QUERY}; uso la query originale.`
          );
          unsubscribe = originalOnSnapshot.apply(originalQuery, args);
          return;
        }

        const filteredQuery = originalQuery.firestore.collection("commesse").where(
          window.firebase.firestore.FieldPath.documentId(),
          "in",
          ids
        );
        bypassQueries.add(filteredQuery);
        state.filteredRootListeners += 1;
        unsubscribe = originalOnSnapshot.apply(filteredQuery, args);
      });

      return () => {
        closed = true;
        try { unsubscribe?.(); } catch (_) {}
      };
    };

    Object.defineProperty(wrappedOnSnapshot, "__heraActiveCommesseFirstBootWrapped", { value: true });
    Object.defineProperty(wrappedOnSnapshot, "__heraActiveCommesseFirstBootOriginal", { value: originalOnSnapshot });
    QueryPrototype.onSnapshot = wrappedOnSnapshot;
    return true;
  }

  function install() {
    const readReady = installIndexReadGuard();
    const listenerReady = installListenerGuard();
    state.installed = readReady && listenerReady;
    return state.installed;
  }

  window[GLOBAL] = {
    installed: false,
    version: VERSION,
    install,
    getState: () => ({ ...state })
  };

  if (!install()) {
    let attempts = 0;
    const timer = window.setInterval(() => {
      attempts += 1;
      if (install() || attempts >= 100) window.clearInterval(timer);
      window[GLOBAL].installed = state.installed;
    }, 25);
  }
  window[GLOBAL].installed = state.installed;
})();
