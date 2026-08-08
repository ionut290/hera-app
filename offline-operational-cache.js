(() => {
  "use strict";

  const GLOBAL = "HeraOperationalOfflineCache";
  const VERSION = "1.1.0";
  const COMMESSE_CACHE_PREFIX = "hera-offline-root-commesse-v1:";
  const IMPIANTI_CACHE_PREFIX = "hera-offline-impianti-v1:";
  const WEATHER_PREFIX = "hera-offline-weather-v1:";
  const SESSION_KEY = "heraPersistedUserSession";
  const PROBE_TIMEOUT_MS = 2200;

  if (window[GLOBAL]?.installed) return;

  const state = {
    installed: false,
    verified: false,
    online: navigator.onLine !== false,
    probing: false,
    lastProbeAt: null,
    cachedCommesseDelivered: 0,
    cachedCommesseSaved: 0,
    cachedImpiantiDelivered: 0,
    cachedImpiantiSaved: 0,
    firestoreListenersSkippedOffline: 0,
    errors: []
  };

  let probePromise = null;
  let connectionObserver = null;
  let weatherObserver = null;

  function safeRead(key) {
    try { return JSON.parse(localStorage.getItem(key) || "null"); } catch (_) { return null; }
  }

  function safeWrite(key, value) {
    try { localStorage.setItem(key, JSON.stringify(value)); return true; }
    catch (_) { return false; }
  }

  function setText(node, value) {
    if (node && node.textContent !== value) node.textContent = value;
  }

  function persistedUid() {
    try {
      const current = window.firebase?.auth?.()?.currentUser?.uid;
      if (current) return String(current);
    } catch (_) {}
    const saved = safeRead(SESSION_KEY);
    return String(saved?.uid || saved?.user?.uid || "").trim();
  }

  function userCacheKey(prefix, suffix = "") {
    const uid = persistedUid();
    if (!uid) return "";
    return `${prefix}${uid}${suffix ? `:${suffix}` : ""}`;
  }

  function safeClone(value) {
    try {
      return JSON.parse(JSON.stringify(value, (_key, item) => {
        if (item && typeof item.toDate === "function") {
          try { return item.toDate().toISOString(); } catch (_) {}
        }
        return item;
      }));
    } catch (_) { return null; }
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
    const candidates = [query?.path, query?._query?.path, query?._query?._path,
      query?._delegate?._query?.path, query?._delegate?._query?._path, query?.Ae?.path, query?.je?.path];
    for (const candidate of candidates) {
      const path = canonicalPath(candidate).replace(/^\/+|\/+$/g, "");
      if (path) return path;
    }
    return "";
  }

  function impiantiCommessaId(path) {
    const match = String(path || "").match(/^commesse\/([^/]+)\/impianti$/);
    return match ? match[1] : "";
  }

  function snapshotOptions(value) {
    return Boolean(value && typeof value === "object" && typeof value.next !== "function"
      && Object.prototype.hasOwnProperty.call(value, "includeMetadataChanges"));
  }

  function wrapNext(argsLike, wrapper) {
    const args = Array.from(argsLike || []);
    const index = snapshotOptions(args[0]) ? 1 : 0;
    const candidate = args[index];
    if (typeof candidate === "function") {
      const originalNext = candidate;
      args[index] = (snapshot) => wrapper(snapshot, originalNext, null);
      return args;
    }
    if (candidate && typeof candidate === "object" && typeof candidate.next === "function") {
      const observer = candidate;
      args[index] = { ...observer, next(snapshot) { return wrapper(snapshot, observer.next, observer); } };
    }
    return args;
  }

  function invokeNext(next, context, snapshot) {
    if (typeof next === "function") next.call(context || undefined, snapshot);
  }

  function nestedValue(data, fieldPath) {
    return String(fieldPath || "").split(".").reduce((value, part) => value == null ? undefined : value[part], data);
  }

  function documentRefForPath(query, path, id) {
    try {
      const firestore = query?.firestore;
      if (!firestore?.collection) return null;
      if (path === "commesse") return firestore.collection("commesse").doc(id);
      const commessaId = impiantiCommessaId(path);
      if (commessaId) return firestore.collection("commesse").doc(commessaId).collection("impianti").doc(id);
    } catch (_) {}
    return null;
  }

  function buildCachedSnapshot(query, path, stored) {
    const metadata = Object.freeze({ fromCache: true, hasPendingWrites: false });
    const rows = Array.isArray(stored?.docs) ? stored.docs : [];
    const docs = rows.map((row) => {
      const data = row?.data && typeof row.data === "object" ? row.data : {};
      const id = String(row?.id || "");
      const ref = documentRefForPath(query, path, id);
      return Object.freeze({
        id, ref, exists: true, metadata,
        data: () => safeClone(data) || { ...data },
        get: (fieldPath) => nestedValue(data, fieldPath),
        isEqual: (other) => Boolean(other && other.id === id)
      });
    });
    const snapshot = {
      query, docs: Object.freeze(docs), size: docs.length, empty: docs.length === 0, metadata,
      forEach(callback, thisArg) { docs.forEach((doc) => callback.call(thisArg, doc)); },
      docChanges() { return docs.map((doc, index) => ({ type: "added", doc, oldIndex: -1, newIndex: index })); },
      isEqual(other) { return other === snapshot; }
    };
    return Object.freeze(snapshot);
  }

  function cacheDescriptor(path) {
    if (path === "commesse") {
      return {
        key: userCacheKey(COMMESSE_CACHE_PREFIX),
        kind: "commesse"
      };
    }
    const commessaId = impiantiCommessaId(path);
    if (commessaId) {
      return {
        key: userCacheKey(IMPIANTI_CACHE_PREFIX, commessaId),
        kind: "impianti",
        commessaId
      };
    }
    return null;
  }

  function readCachedSnapshot(path) {
    const descriptor = cacheDescriptor(path);
    return descriptor?.key ? safeRead(descriptor.key) : null;
  }

  function saveSnapshot(path, snapshot) {
    const descriptor = cacheDescriptor(path);
    if (!descriptor?.key || !snapshot?.docs) return;
    if (snapshot.empty && snapshot.metadata?.fromCache) return;
    const docs = snapshot.docs
      .map((doc) => ({ id: String(doc.id || ""), data: safeClone(doc.data?.() || {}) || {} }))
      .filter((row) => row.id);
    const payload = { savedAt: new Date().toISOString(), path, docs };
    if (!safeWrite(descriptor.key, payload)) return;
    if (descriptor.kind === "commesse") state.cachedCommesseSaved += 1;
    else state.cachedImpiantiSaved += 1;
  }

  function renderConnectionUi() {
    const indicator = document.getElementById("connection-indicator");
    const offline = document.getElementById("offline-mode-indicator");
    if (!indicator && !offline) return;
    if (state.verified && !state.online) {
      setText(indicator, "🔴 Offline");
      if (offline) {
        setText(offline, "📦 Modalità offline attiva");
        offline.classList.remove("hidden");
      }
    } else if (state.verified && state.online) {
      setText(indicator, "🟢 Online • Connessione disponibile");
      offline?.classList.add("hidden");
    } else {
      setText(indicator, "🟡 Connessione in verifica…");
    }
  }

  function restoreWeatherOffline() {
    if (!state.verified || state.online) return;
    const summary = document.getElementById("weather-summary");
    if (!summary) return;
    const key = userCacheKey(WEATHER_PREFIX);
    const saved = key ? safeRead(key) : null;
    if (saved?.summary) {
      const when = saved.savedAt ? new Date(saved.savedAt).toLocaleString("it-IT", { dateStyle: "short", timeStyle: "short" }) : "precedentemente";
      setText(summary, `📦 Offline. Ultimo dato online (${when}): ${saved.summary} — dato non aggiornato, non usare per decisioni di sicurezza.`);
    } else {
      setText(summary, "📦 Offline. Meteo live non disponibile.");
    }
  }

  function rememberWeatherIfUseful() {
    if (!state.verified || !state.online) return;
    const text = String(document.getElementById("weather-summary")?.textContent || "").trim();
    if (!text || /caricamento|non disponibile|offline/i.test(text)) return;
    const key = userCacheKey(WEATHER_PREFIX);
    if (key) safeWrite(key, { savedAt: new Date().toISOString(), summary: text });
  }

  function emitConnectivity() {
    renderConnectionUi();
    if (!state.online) restoreWeatherOffline();
    window.dispatchEvent(new CustomEvent("hera:verified-connectivity", {
      detail: { online: state.online, verified: state.verified, at: state.lastProbeAt }
    }));
    window.dispatchEvent(new CustomEvent("hera:offline-mode", { detail: { offline: !state.online, verified: true } }));
  }

  async function probeConnectivity(force = false) {
    if (probePromise && !force) return probePromise;
    if (navigator.onLine === false) {
      state.verified = true;
      state.online = false;
      state.lastProbeAt = new Date().toISOString();
      emitConnectivity();
      return false;
    }
    state.probing = true;
    renderConnectionUi();
    const request = (async () => {
      const controller = new AbortController();
      const timeout = setTimeout(() => controller.abort(), PROBE_TIMEOUT_MS);
      try {
        const url = new URL("./manifest.webmanifest", window.location.href);
        url.searchParams.set("hera_connectivity", String(Date.now()));
        const response = await fetch(url.href, { method: "GET", cache: "no-store", credentials: "same-origin", signal: controller.signal });
        state.online = Boolean(response?.ok);
      } catch (_) {
        state.online = false;
      } finally {
        clearTimeout(timeout);
        state.verified = true;
        state.probing = false;
        state.lastProbeAt = new Date().toISOString();
        emitConnectivity();
      }
      return state.online;
    })();
    probePromise = request;
    request.finally(() => { if (probePromise === request) probePromise = null; });
    return request;
  }

  function installOperationalReadCache() {
    const QueryPrototype = window.firebase?.firestore?.Query?.prototype;
    if (!QueryPrototype || typeof QueryPrototype.onSnapshot !== "function") return false;
    if (QueryPrototype.onSnapshot.__heraOperationalOfflineCacheWrapped) return true;
    const originalOnSnapshot = QueryPrototype.onSnapshot;

    const wrapped = function offlineAwareOnSnapshot() {
      const path = queryPath(this);
      const descriptor = cacheDescriptor(path);
      if (!descriptor) return originalOnSnapshot.apply(this, arguments);

      const query = this;
      const originalArgs = Array.from(arguments);
      const cached = readCachedSnapshot(path);
      let cancelled = false;
      let unsubscribe = () => {};
      let cachedDelivered = false;

      const deliver = (stored) => {
        const wrappedArgs = wrapNext(originalArgs, (snapshot, next, context) => invokeNext(next, context, snapshot));
        const index = snapshotOptions(wrappedArgs[0]) ? 1 : 0;
        const candidate = wrappedArgs[index];
        const snapshot = buildCachedSnapshot(query, path, stored);
        if (typeof candidate === "function") candidate(snapshot);
        else candidate?.next?.(snapshot);
      };

      if (cached?.docs?.length) {
        deliver(cached);
        cachedDelivered = true;
        if (descriptor.kind === "commesse") state.cachedCommesseDelivered += 1;
        else state.cachedImpiantiDelivered += 1;
      }

      Promise.resolve(probeConnectivity()).then((online) => {
        if (cancelled) return;
        if (!online) {
          state.firestoreListenersSkippedOffline += 1;
          if (!cachedDelivered) deliver({ docs: [] });
          return;
        }
        const args = wrapNext(originalArgs, (snapshot, next, context) => {
          saveSnapshot(path, snapshot);
          invokeNext(next, context, snapshot);
        });
        unsubscribe = originalOnSnapshot.apply(query, args);
      }).catch((error) => state.errors.push(String(error?.message || error || "probe-error")));

      return () => { cancelled = true; try { unsubscribe?.(); } catch (_) {} };
    };

    Object.defineProperty(wrapped, "__heraOperationalOfflineCacheWrapped", { value: true });
    Object.defineProperty(wrapped, "__heraOperationalOfflineCacheOriginal", { value: originalOnSnapshot });
    QueryPrototype.onSnapshot = wrapped;
    return true;
  }

  function installUiObservers() {
    const bind = () => {
      const indicator = document.getElementById("connection-indicator");
      if (indicator && !connectionObserver) {
        connectionObserver = new MutationObserver(() => { if (state.verified) renderConnectionUi(); });
        connectionObserver.observe(indicator, { childList: true, characterData: true, subtree: true });
      }
      const summary = document.getElementById("weather-summary");
      if (summary && !weatherObserver) {
        weatherObserver = new MutationObserver(() => {
          if (state.online) rememberWeatherIfUseful();
          else restoreWeatherOffline();
        });
        weatherObserver.observe(summary, { childList: true, characterData: true, subtree: true });
      }
      renderConnectionUi();
      if (!state.online) restoreWeatherOffline();
    };
    if (document.readyState === "loading") document.addEventListener("DOMContentLoaded", bind, { once: true });
    else bind();
  }

  function install() {
    const cacheReady = installOperationalReadCache();
    installUiObservers();
    state.installed = cacheReady;
    return state.installed;
  }

  window.addEventListener("offline", () => {
    state.verified = true;
    state.online = false;
    state.lastProbeAt = new Date().toISOString();
    emitConnectivity();
  });
  window.addEventListener("online", () => { void probeConnectivity(true); });
  document.addEventListener("visibilitychange", () => {
    if (document.visibilityState === "visible") void probeConnectivity(true);
  });

  window[GLOBAL] = {
    installed: false,
    version: VERSION,
    probe: probeConnectivity,
    isOffline: () => state.verified ? !state.online : navigator.onLine === false,
    getCachedCommesse: () => readCachedSnapshot("commesse"),
    getCachedImpianti: (commessaId) => readCachedSnapshot(`commesse/${String(commessaId || "")}/impianti`),
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
  void probeConnectivity();
})();