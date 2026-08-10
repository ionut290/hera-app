(function installPwaLoginForceUpdate() {
  "use strict";

  const CACHE_PREFIX = "varga-cantieri-";
  const APP_VERSION = "v120";
  const SERVICE_WORKER_URL = "./sw.js";
  const GOOGLE_LOGIN_BUTTON_ID = "auth-gate-login-btn";
  const GOOGLE_LOGIN_PROVIDER_KEY = "heraLastAuthProvider";
  let updating = false;
  let googleLoginBusy = false;

  function mountVersionBadges() {
    if (!document.getElementById("pwa-version-styles")) {
      const style = document.createElement("style");
      style.id = "pwa-version-styles";
      style.textContent = ".pwa-version-badge{display:flex;width:max-content;max-width:100%;align-items:center;justify-content:center;margin:.3rem auto .15rem;padding:.18rem .5rem;border:1px solid rgba(7,91,73,.18);border-radius:999px;background:rgba(237,249,245,.88);color:#45645d;font-size:.66rem;font-weight:800;line-height:1.15;white-space:nowrap}.pwa-version-badge.is-update{border-color:#d89b18;background:#fff7da;color:#805b00;cursor:pointer}#home-pwa-version{opacity:.82}";
      document.head.appendChild(style);
    }
    const add = (id, parent) => {
      if (!parent || document.getElementById(id)) return;
      const badge = document.createElement("button");
      badge.id = id;
      badge.type = "button";
      badge.className = "pwa-version-badge";
      badge.textContent = `PWA ${APP_VERSION} • Aggiornata`;
      badge.title = "Versione PWA installata";
      badge.addEventListener("click", () => badge.classList.contains("is-update") && forcePwaUpdate());
      parent.appendChild(badge);
    };
    add("auth-pwa-version", document.querySelector("#auth-gate .auth-gate-card"));
    add("home-pwa-version", document.getElementById("home-screen") || document.querySelector("main") || document.getElementById("app"));
  }

  function setUpdateAvailable() {
    ["auth-pwa-version", "home-pwa-version"].forEach((id) => {
      const badge = document.getElementById(id);
      if (!badge) return;
      badge.textContent = "⚠️ Aggiornamento PWA disponibile • Aggiorna";
      badge.classList.add("is-update");
      badge.title = "Premi per aggiornare la PWA";
    });
  }

  async function ensureServiceWorkerRegistration() {
    if (!("serviceWorker" in navigator)) return null;
    try {
      const existing = await navigator.serviceWorker.getRegistration();
      if (existing) return existing;
      const registration = await navigator.serviceWorker.register(SERVICE_WORKER_URL);
      console.log("PWA SERVICE WORKER REGISTERED", registration.scope);
      return registration;
    } catch (error) {
      console.warn("Registrazione Service Worker PWA non riuscita:", error);
      return null;
    }
  }

  async function watchForUpdates() {
    if (!("serviceWorker" in navigator)) return;
    try {
      const registration = await ensureServiceWorkerRegistration();
      if (!registration) return;
      if (registration.waiting) setUpdateAvailable();
      registration.addEventListener("updatefound", () => {
        const worker = registration.installing;
        worker?.addEventListener("statechange", () => {
          if (worker.state === "installed" && navigator.serviceWorker.controller) setUpdateAvailable();
        });
      });
      if (navigator.onLine) await registration.update();
    } catch (error) {
      console.warn("Controllo aggiornamento PWA non riuscito:", error);
    }
  }

  function setFeedback(message) {
    const feedback = document.getElementById("auth-email-feedback");
    if (feedback) feedback.textContent = message;
  }

  function setGoogleButtonBusy(busy, text) {
    const button = document.getElementById(GOOGLE_LOGIN_BUTTON_ID);
    if (!button) return;
    button.disabled = Boolean(busy);
    button.textContent = text || (busy ? "Accesso Google..." : "Login con Google");
    button.setAttribute("aria-busy", busy ? "true" : "false");
  }

  function isAndroidWebViewRuntime() {
    const capacitorPlatform = window.Capacitor && typeof window.Capacitor.getPlatform === "function"
      ? String(window.Capacitor.getPlatform() || "").toLowerCase()
      : "";
    const ua = String(navigator.userAgent || "");
    return capacitorPlatform === "android"
      || (/Android/i.test(ua) && /; wv\)|\bwv\b|Version\/\d+\.\d+.*Chrome\//i.test(ua));
  }

  function canUseFirebaseGoogleAuth() {
    return Boolean(
      window.firebase
      && typeof firebase.auth === "function"
      && firebase.auth.GoogleAuthProvider
    );
  }

  function getGoogleLoginErrorMessage(error, androidRuntime) {
    const code = String(error?.code || "").toLowerCase();
    if (code === "auth/popup-closed-by-user" || code === "auth/cancelled-popup-request") {
      return "Accesso Google annullato. Puoi riprovare oppure usare email e password.";
    }
    if (code === "auth/network-request-failed") {
      return "Rete troppo debole per Google. Se hai già effettuato l’accesso su questo dispositivo, riapri l’app; altrimenti usa email e password quando torna la connessione.";
    }
    if (code === "auth/unauthorized-domain") {
      return "Google non autorizza questo indirizzo dell’app. Contatta l’amministratore.";
    }
    if (code === "auth/account-exists-with-different-credential") {
      return "Questo indirizzo è già collegato con un altro metodo di accesso. Entra con email e password.";
    }
    if (code === "auth/popup-blocked" || code === "auth/operation-not-supported-in-this-environment") {
      return androidRuntime
        ? "Il login Google non può aprire la finestra su questo Android. Usa email e password come accesso principale."
        : "Il browser ha bloccato la finestra Google. Provo il metodo alternativo...";
    }
    return "Accesso Google non riuscito. Puoi riprovare oppure usare email e password.";
  }

  function shouldFallbackToGoogleRedirect(error) {
    const code = String(error?.code || "").toLowerCase();
    return code === "auth/popup-blocked"
      || code === "auth/operation-not-supported-in-this-environment";
  }

  async function persistGoogleAuthPreference(auth) {
    if (!auth || typeof auth.setPersistence !== "function") return;
    const localPersistence = firebase.auth.Auth?.Persistence?.LOCAL
      || firebase.auth.browserLocalPersistence;
    if (!localPersistence) return;
    try {
      await auth.setPersistence(localPersistence);
    } catch (error) {
      const code = String(error?.code || "").toLowerCase();
      if (code !== "auth/unsupported-persistence-type" && code !== "auth/invalid-persistence-type") throw error;
      console.warn("Persistenza Google permanente non supportata; continuo con la sessione disponibile.", error);
    }
  }

  async function completeGoogleLoginSuccess(result) {
    const user = result?.user || firebase.auth().currentUser;
    if (!user) throw new Error("Google non ha restituito un utente autenticato.");
    try {
      localStorage.setItem(GOOGLE_LOGIN_PROVIDER_KEY, "google");
    } catch (_) {}
    setFeedback("Accesso Google completato. Sessione salvata sul dispositivo.");
    console.log("GOOGLE BACKUP LOGIN OK", { uid: user.uid, email: user.email || "" });
    return user;
  }

  async function loginWithGoogleBackup() {
    if (googleLoginBusy) return;
    if (!canUseFirebaseGoogleAuth()) {
      setFeedback("Login Google non disponibile in questo momento. Usa email e password.");
      return;
    }
    if (navigator.onLine === false) {
      setFeedback("Senza Internet non è possibile avviare un nuovo login Google. Se avevi già effettuato l’accesso, riapri l’app per usare la sessione salvata.");
      return;
    }

    googleLoginBusy = true;
    setGoogleButtonBusy(true);
    setFeedback("Apro Google per l’accesso di riserva...");

    const auth = firebase.auth();
    const provider = new firebase.auth.GoogleAuthProvider();
    provider.addScope("https://www.googleapis.com/auth/userinfo.email");
    provider.setCustomParameters({ prompt: "select_account" });
    const androidRuntime = isAndroidWebViewRuntime();

    try {
      await persistGoogleAuthPreference(auth);
      const result = await auth.signInWithPopup(provider);
      await completeGoogleLoginSuccess(result);
    } catch (error) {
      console.warn("Google backup login popup non riuscito:", error);
      if (!androidRuntime && shouldFallbackToGoogleRedirect(error) && typeof auth.signInWithRedirect === "function") {
        try {
          setFeedback("Popup Google bloccato. Passo al metodo alternativo...");
          await persistGoogleAuthPreference(auth);
          await auth.signInWithRedirect(provider);
          return;
        } catch (redirectError) {
          console.error("Google backup login redirect non riuscito:", redirectError);
          setFeedback(getGoogleLoginErrorMessage(redirectError, androidRuntime));
        }
      } else {
        setFeedback(getGoogleLoginErrorMessage(error, androidRuntime));
      }
    } finally {
      googleLoginBusy = false;
      setGoogleButtonBusy(false);
    }
  }

  async function consumeGoogleRedirectResult() {
    if (!canUseFirebaseGoogleAuth()) return;
    const auth = firebase.auth();
    if (typeof auth.getRedirectResult !== "function") return;
    try {
      await persistGoogleAuthPreference(auth);
      const result = await auth.getRedirectResult();
      if (result?.user) await completeGoogleLoginSuccess(result);
    } catch (error) {
      console.warn("Esito redirect Google non riuscito:", error);
      setFeedback(getGoogleLoginErrorMessage(error, false));
    }
  }

  function handleGoogleLoginClick(event) {
    const button = event.target?.closest?.(`#${GOOGLE_LOGIN_BUTTON_ID}`);
    if (!button) return;
    event.preventDefault();
    event.stopPropagation();
    event.stopImmediatePropagation();
    void loginWithGoogleBackup();
  }

  async function forcePwaUpdate() {
    if (updating) return;
    const button = document.getElementById("auth-force-update-btn");

    if (navigator.onLine === false) {
      setFeedback("Per aggiornare l’app è necessaria una connessione Internet.");
      return;
    }

    updating = true;
    if (button) {
      button.disabled = true;
      button.textContent = "AGGIORNAMENTO...";
    }
    setFeedback("Aggiornamento dell’app in corso. Attendi...");

    try {
      if ("serviceWorker" in navigator) {
        await ensureServiceWorkerRegistration();
        const registrations = await navigator.serviceWorker.getRegistrations();
        await Promise.allSettled(registrations.map(async (registration) => {
          await registration.update();
          registration.waiting?.postMessage({ type: "SKIP_WAITING" });
        }));
      }

      if ("caches" in window) {
        const keys = await caches.keys();
        await Promise.all(
          keys
            .filter((key) => key.startsWith(CACHE_PREFIX))
            .map((key) => caches.delete(key))
        );
      }

      const target = new URL(window.location.href);
      target.searchParams.set("pwa-refresh", String(Date.now()));
      setFeedback("Aggiornamento completato. Riapertura dell’app...");
      window.setTimeout(() => window.location.replace(target.href), 250);
    } catch (error) {
      console.error("Aggiornamento PWA non riuscito:", error);
      updating = false;
      if (button) {
        button.disabled = false;
        button.textContent = "🔄 AGGIORNA APP";
      }
      setFeedback("Aggiornamento non riuscito. Controlla Internet e riprova.");
    }
  }

  function initialize() {
    document.addEventListener("click", handleGoogleLoginClick, true);
    document.getElementById("auth-force-update-btn")
      ?.addEventListener("click", forcePwaUpdate);
    mountVersionBadges();
    new MutationObserver(mountVersionBadges).observe(document.body, { childList: true, subtree: true });
    watchForUpdates();
    void consumeGoogleRedirectResult();
  }

  if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", initialize, { once: true });
  } else {
    initialize();
  }
})();
