(function installPersistentOfflineAuth() {
  "use strict";

  const AUTHZ_PREFIX = "hera-offline-authz:";
  const APP_SESSION_KEY = "heraPersistedUserSession";
  const ACTIVE_STATUSES = new Set(["attivo", "active", "approved", "autorizzato", "abilitato"]);
  let verifiedConnectivityOnline = null;

  function normalizeStatus(value) {
    return String(value || "").trim().toLowerCase();
  }

  function normalizeIdentifier(value) {
    return String(value || "")
      .trim()
      .toLowerCase()
      .normalize("NFD")
      .replace(/[\u0300-\u036f]/g, "")
      .replace(/[^a-z0-9@._-]+/g, ".")
      .replace(/^\.+|\.+$/g, "")
      .replace(/\.{2,}/g, ".");
  }

  function cacheKey(uid) {
    return `${AUTHZ_PREFIX}${String(uid || "").trim()}`;
  }

  function readPersistedAppSession() {
    try {
      const raw = localStorage.getItem(APP_SESSION_KEY);
      if (!raw) return null;
      const value = JSON.parse(raw);
      if (!value || !String(value.uid || value.user?.uid || "").trim()) return null;
      const normalized = value.uid ? value : { ...value, uid: value.user?.uid };
      if (normalized.banned === true || normalized.user?.banned === true) return null;
      return normalized;
    } catch (_) {
      return null;
    }
  }

  function rememberAllowed(user, result) {
    if (!user?.uid || !result?.allowed) return;
    const status = normalizeStatus(result.status || result.profile?.statoAccount || result.profile?.accountStatus || "attivo");
    if (!ACTIVE_STATUSES.has(status)) return;
    try {
      localStorage.setItem(cacheKey(user.uid), JSON.stringify({
        uid: user.uid,
        email: String(user.email || result.profile?.email || ""),
        displayName: String(user.displayName || result.profile?.displayName || result.profile?.nomeCompleto || "Utente"),
        status,
        verifiedAt: Date.now()
      }));
    } catch (_) {}
  }

  function forgetAllowed(uid) {
    if (!uid) return;
    try { localStorage.removeItem(cacheKey(uid)); } catch (_) {}
  }

  function readAllowed(uid) {
    if (!uid) return null;
    try {
      const raw = localStorage.getItem(cacheKey(uid));
      const value = raw ? JSON.parse(raw) : null;
      return value && value.uid === uid && ACTIVE_STATUSES.has(normalizeStatus(value.status)) ? value : null;
    } catch (_) {
      return null;
    }
  }

  function isNetworkError(error) {
    const code = String(error?.code || "").toLowerCase();
    const message = String(error?.message || error || "").toLowerCase();
    return code.includes("network-request-failed")
      || code.includes("unavailable")
      || /connessione non disponibile|network error|failed to fetch|offline|internet/.test(message);
  }

  function isEffectivelyOffline() {
    if (navigator.onLine === false) return true;
    if (verifiedConnectivityOnline === false) return true;
    try {
      if (window.HeraOperationalOfflineCache?.isOffline?.() === true) return true;
    } catch (_) {}
    return false;
  }

  function hideApprovalGate() {
    document.body?.classList.remove("access-approval-locked");
    document.getElementById("access-approval-screen")?.classList.add("hidden");
  }

  function releaseOfflineStartupGate(user, result, options = {}) {
    const force = options.force === true;
    if ((!force && !isEffectivelyOffline()) || !user?.uid || !result?.allowed) return false;
    hideApprovalGate();
    document.body?.classList.remove("auth-pending");
    document.body?.classList.add("offline-session-active");
    document.getElementById("auth-gate")?.classList.add("hidden");
    const startup = document.getElementById("app-startup-loading");
    if (startup) {
      startup.classList.add("hidden");
      startup.style.display = "none";
      startup.setAttribute("aria-hidden", "true");
    }
    const home = document.getElementById("home-page");
    if (home) home.classList.remove("hidden");

    window.__heraOfflineBootSession = {
      ...(result.profile || {}),
      uid: user.uid,
      email: user.email || result.profile?.email || "",
      displayName: user.displayName || result.profile?.displayName || result.profile?.nomeCompleto || "Utente",
      offline: true,
      source: result.source || "device-session"
    };

    document.dispatchEvent(new CustomEvent("hera:offline-session-ready", {
      detail: {
        uid: user.uid,
        email: user.email || "",
        source: result.source || "device-session",
        forcedByNetworkFailure: force
      }
    }));
    window.dispatchEvent(new CustomEvent("hera:offline-mode", {
      detail: { offline: true, verified: true, source: result.source || "device-session" }
    }));
    return true;
  }

  function sessionApproved(session, cachedAuthorization) {
    if (!session?.uid) return false;
    if (session.banned === true || session.user?.banned === true) return false;
    if (cachedAuthorization) return true;
    const status = normalizeStatus(
      session.statoAccount || session.accountStatus || session.status || session.user?.statoAccount || session.user?.accountStatus || ""
    );
    return session.accessApproved === true || ACTIVE_STATUSES.has(status);
  }

  function identifierMatchesSession(identifier, session) {
    const requested = normalizeIdentifier(identifier);
    if (!requested) return true;

    try {
      const currentUser = window.firebase?.auth?.()?.currentUser;
      if (currentUser?.uid && String(currentUser.uid) === String(session.uid)) return true;
    } catch (_) {}

    const values = new Set();
    const add = (value) => {
      const normalized = normalizeIdentifier(value);
      if (normalized) values.add(normalized);
    };

    add(session.email);
    add(session.username);
    add(session.operatorUsername);
    add(session.userName);
    add(session.displayName);
    add(session.nomeCompleto);
    add(session.user?.email);
    add(session.user?.username);
    add(session.user?.displayName);

    for (const value of Array.from(values)) {
      if (value.includes("@")) add(value.split("@")[0]);
      if (value.includes(".")) add(value.replace(/\./g, ""));
    }

    return values.has(requested) || values.has(requested.replace(/\./g, ""));
  }

  function bootstrapFromSavedSession(options = {}) {
    const force = options.force === true;
    if (!force && !isEffectivelyOffline()) return false;
    const session = readPersistedAppSession();
    if (!session?.uid) return false;

    const cachedAuthorization = readAllowed(session.uid);
    if (!sessionApproved(session, cachedAuthorization)) return false;
    if (!identifierMatchesSession(options.identifier || "", session)) return false;

    const user = {
      uid: String(session.uid),
      email: String(session.email || session.user?.email || cachedAuthorization?.email || ""),
      displayName: String(
        session.displayName || session.userName || session.nomeCompleto || session.user?.displayName || cachedAuthorization?.displayName || "Utente"
      )
    };
    const result = {
      allowed: true,
      offline: true,
      source: cachedAuthorization ? "device-authorization" : "persisted-app-session",
      status: cachedAuthorization?.status || normalizeStatus(session.statoAccount || session.accountStatus || "attivo") || "attivo",
      profile: session
    };
    return releaseOfflineStartupGate(user, result, { force });
  }

  async function verifyOfflineFromCache(firebaseUser, options = {}) {
    if (!firebaseUser?.uid) return null;

    if (window.firebase && typeof firebase.firestore === "function") {
      try {
        const snapshot = await firebase.firestore().collection("platformUsers").doc(firebaseUser.uid).get({ source: "cache" });
        if (snapshot.exists) {
          const profile = { id: snapshot.id, ...(snapshot.data() || {}) };
          const status = normalizeStatus(profile.banned === true ? "bloccato" : (profile.statoAccount || profile.accountStatus || "attivo"));
          if (ACTIVE_STATUSES.has(status)) {
            const result = { allowed: true, profile, status, offline: true, source: "firestore-cache" };
            rememberAllowed(firebaseUser, result);
            releaseOfflineStartupGate(firebaseUser, result, options);
            return result;
          }
          forgetAllowed(firebaseUser.uid);
          return { allowed: false, profile, status, offline: true, source: "firestore-cache" };
        }
      } catch (_) {}
    }

    const saved = readAllowed(firebaseUser.uid);
    if (!saved) return null;
    const result = {
      allowed: true,
      status: saved.status,
      offline: true,
      source: "device-session",
      profile: {
        uid: saved.uid,
        email: saved.email,
        displayName: saved.displayName,
        nomeCompleto: saved.displayName,
        statoAccount: saved.status,
        accountStatus: saved.status
      }
    };
    releaseOfflineStartupGate(firebaseUser, result, options);
    return result;
  }

  function wrapApprovalApi(api) {
    if (!api || api.__heraPersistentOfflineWrapped || typeof api.verify !== "function") return api;
    const originalVerify = api.verify.bind(api);

    api.verify = async function verifyWithOfflineSession(firebaseUser) {
      if (isEffectivelyOffline()) {
        const cached = await verifyOfflineFromCache(firebaseUser);
        if (cached) return cached;
      }
      try {
        const result = await originalVerify(firebaseUser);
        if (result?.allowed) rememberAllowed(firebaseUser, result);
        else if (result && !result.error) forgetAllowed(firebaseUser?.uid);
        return result;
      } catch (error) {
        if (isEffectivelyOffline() || isNetworkError(error)) {
          const cached = await verifyOfflineFromCache(firebaseUser, { force: isNetworkError(error) });
          if (cached) return cached;
          if (bootstrapFromSavedSession({ force: isNetworkError(error) })) {
            return {
              allowed: true,
              offline: true,
              source: "persisted-app-session-network-fallback",
              profile: readPersistedAppSession()
            };
          }
        }
        throw error;
      }
    };

    Object.defineProperty(api, "__heraPersistentOfflineWrapped", {
      value: true,
      configurable: false,
      enumerable: false
    });
    return api;
  }

  let approvalApi = window.HeraAccessApproval;
  if (approvalApi) {
    window.HeraAccessApproval = wrapApprovalApi(approvalApi);
  } else {
    try {
      Object.defineProperty(window, "HeraAccessApproval", {
        configurable: true,
        enumerable: true,
        get() { return approvalApi; },
        set(value) { approvalApi = wrapApprovalApi(value); }
      });
    } catch (_) {}
  }

  function ensureLocalAuthPersistence() {
    try {
      if (!window.firebase || typeof firebase.auth !== "function" || !firebase.auth.Auth?.Persistence?.LOCAL) return false;
      const auth = firebase.auth();
      if (!auth) return false;
      Promise.resolve(auth.setPersistence(firebase.auth.Auth.Persistence.LOCAL)).catch((error) => {
        console.warn("Persistenza login locale non impostata:", error);
      });
      return true;
    } catch (_) {
      return false;
    }
  }

  function tryReleasePersistedOfflineUser(options = {}) {
    if (!options.force && !isEffectivelyOffline()) return false;
    if (bootstrapFromSavedSession(options)) return true;
    try {
      if (!window.firebase || typeof firebase.auth !== "function") return false;
      const currentUser = firebase.auth().currentUser;
      if (currentUser?.uid) {
        void verifyOfflineFromCache(currentUser, options);
        return true;
      }
    } catch (_) {}
    return false;
  }

  function installLoginNetworkFailureFallback() {
    const bind = () => {
      const feedback = document.getElementById("auth-email-feedback");
      if (!feedback || feedback.__heraOfflineFallbackObserved) return;
      feedback.__heraOfflineFallbackObserved = true;
      const observer = new MutationObserver(() => {
        const text = String(feedback.textContent || "");
        if (!/connessione non disponibile|network request failed|failed to fetch|offline/i.test(text)) return;
        const identifier = document.getElementById("auth-email-input")?.value || "";
        const resumed = bootstrapFromSavedSession({ force: true, identifier });
        if (resumed) feedback.textContent = "Accesso offline con sessione salvata.";
      });
      observer.observe(feedback, { childList: true, characterData: true, subtree: true });
    };
    if (document.readyState === "loading") document.addEventListener("DOMContentLoaded", bind, { once: true });
    else bind();
  }

  let attempts = 0;
  const persistenceTimer = window.setInterval(() => {
    attempts += 1;
    ensureLocalAuthPersistence();
    if (isEffectivelyOffline()) tryReleasePersistedOfflineUser();
    if (attempts >= 40 || (navigator.onLine !== false && ensureLocalAuthPersistence())) window.clearInterval(persistenceTimer);
  }, 250);

  window.addEventListener("hera:verified-connectivity", (event) => {
    verifiedConnectivityOnline = event?.detail?.online !== false;
    if (verifiedConnectivityOnline === false) tryReleasePersistedOfflineUser({ force: true });
  });
  window.addEventListener("online", () => {
    verifiedConnectivityOnline = null;
    document.body?.classList.remove("offline-session-active");
    try {
      const currentUser = firebase.auth().currentUser;
      if (!currentUser?.uid || !window.HeraAccessApproval?.verify) return;
      void window.HeraAccessApproval.verify(currentUser);
    } catch (_) {}
  });
  window.addEventListener("offline", () => {
    verifiedConnectivityOnline = false;
    tryReleasePersistedOfflineUser({ force: true });
  });

  installLoginNetworkFailureFallback();

  if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", () => tryReleasePersistedOfflineUser(), { once: true });
  } else {
    tryReleasePersistedOfflineUser();
  }
  window.setTimeout(() => tryReleasePersistedOfflineUser(), 50);
  window.setTimeout(() => tryReleasePersistedOfflineUser(), 500);
  window.setTimeout(() => tryReleasePersistedOfflineUser(), 1500);

  window.HeraPersistentOfflineAuth = {
    installed: true,
    readAllowed,
    forgetAllowed,
    readPersistedAppSession,
    bootstrapFromSavedSession,
    verifyOfflineFromCache,
    releaseOfflineStartupGate,
    isEffectivelyOffline,
    resumeFromSavedSession(identifier = "") {
      return bootstrapFromSavedSession({ force: true, identifier });
    }
  };
})();