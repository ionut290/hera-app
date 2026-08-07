(function installPersistentOfflineAuth() {
  "use strict";

  const AUTHZ_PREFIX = "hera-offline-authz:";
  const ACTIVE_STATUSES = new Set(["attivo", "active", "approved", "autorizzato", "abilitato"]);

  function normalizeStatus(value) {
    return String(value || "").trim().toLowerCase();
  }

  function cacheKey(uid) {
    return `${AUTHZ_PREFIX}${String(uid || "").trim()}`;
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

  function hideApprovalGate() {
    document.body?.classList.remove("access-approval-locked");
    document.getElementById("access-approval-screen")?.classList.add("hidden");
  }

  function releaseOfflineStartupGate(firebaseUser, result) {
    if (navigator.onLine || !firebaseUser?.uid || !result?.allowed) return false;
    hideApprovalGate();
    document.body?.classList.remove("auth-pending");
    document.getElementById("auth-gate")?.classList.add("hidden");
    const startup = document.getElementById("app-startup-loading");
    if (startup) {
      startup.classList.add("hidden");
      startup.setAttribute("aria-hidden", "true");
    }
    document.dispatchEvent(new CustomEvent("hera:offline-session-ready", {
      detail: { uid: firebaseUser.uid, source: result.source || "device-session" }
    }));
    return true;
  }

  async function verifyOfflineFromCache(firebaseUser) {
    if (!firebaseUser?.uid) return null;

    if (window.firebase && typeof firebase.firestore === "function") {
      try {
        const snapshot = await firebase.firestore()
          .collection("platformUsers")
          .doc(firebaseUser.uid)
          .get({ source: "cache" });
        if (snapshot.exists) {
          const profile = { id: snapshot.id, ...(snapshot.data() || {}) };
          const status = normalizeStatus(profile.banned === true ? "bloccato" : (profile.statoAccount || profile.accountStatus || "attivo"));
          if (ACTIVE_STATUSES.has(status)) {
            const result = { allowed: true, profile, status, offline: true, source: "firestore-cache" };
            rememberAllowed(firebaseUser, result);
            hideApprovalGate();
            releaseOfflineStartupGate(firebaseUser, result);
            return result;
          }
          forgetAllowed(firebaseUser.uid);
          return { allowed: false, profile, status, offline: true, source: "firestore-cache" };
        }
      } catch (_) {}
    }

    const saved = readAllowed(firebaseUser.uid);
    if (!saved) return null;
    hideApprovalGate();
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
    releaseOfflineStartupGate(firebaseUser, result);
    return result;
  }

  function wrapApprovalApi(api) {
    if (!api || api.__heraPersistentOfflineWrapped || typeof api.verify !== "function") return api;
    const originalVerify = api.verify.bind(api);

    api.verify = async function verifyWithOfflineSession(firebaseUser) {
      if (!navigator.onLine) {
        const cached = await verifyOfflineFromCache(firebaseUser);
        if (cached) return cached;
      }

      try {
        const result = await originalVerify(firebaseUser);
        if (result?.allowed) rememberAllowed(firebaseUser, result);
        else if (result && !result.error) forgetAllowed(firebaseUser?.uid);
        return result;
      } catch (error) {
        if (!navigator.onLine) {
          const cached = await verifyOfflineFromCache(firebaseUser);
          if (cached) return cached;
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

  function tryReleasePersistedOfflineUser() {
    if (navigator.onLine || !window.firebase || typeof firebase.auth !== "function") return;
    const currentUser = firebase.auth().currentUser;
    if (!currentUser?.uid) return;
    void verifyOfflineFromCache(currentUser);
  }

  let attempts = 0;
  const persistenceTimer = window.setInterval(() => {
    attempts += 1;
    if (ensureLocalAuthPersistence() || attempts >= 40) window.clearInterval(persistenceTimer);
    if (!navigator.onLine) tryReleasePersistedOfflineUser();
  }, 250);

  window.addEventListener("online", () => {
    try {
      const currentUser = firebase.auth().currentUser;
      if (!currentUser?.uid || !window.HeraAccessApproval?.verify) return;
      void window.HeraAccessApproval.verify(currentUser);
    } catch (_) {}
  });

  document.addEventListener("DOMContentLoaded", tryReleasePersistedOfflineUser, { once: true });

  window.HeraPersistentOfflineAuth = {
    installed: true,
    readAllowed,
    forgetAllowed,
    verifyOfflineFromCache,
    releaseOfflineStartupGate
  };
})();
