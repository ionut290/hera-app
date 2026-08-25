(() => {
  "use strict";

  const ROTATION_MS = 2500;
  const SLOW_NOTICE_MS = 5000;
  const LOGIN_FAILSAFE_MS = 4500;
  const TRUSTED_DEVICE_KEY = "heraTrustedDeviceSessionV1";
  const AUTH_WATCH_MAX_MS = 20000;
  const messages = [
    ["🧑‍🌾", "Sto preparando il cantiere…"],
    ["🗺️", "Metto in ordine commesse e squadre…"],
    ["🤔", "Sto pensando… quasi fatto!"],
    ["🛠️", "Controllo che sia tutto al posto giusto…"],
    ["🌱", "Ancora un attimo, ci siamo…"]
  ];
  const controllers = new Map();
  let authWatchStarted = false;
  let authStateResolved = false;
  let authenticatedUser = null;

  function readTrustedDevice() {
    try {
      const raw = localStorage.getItem(TRUSTED_DEVICE_KEY);
      if (!raw) return null;
      const parsed = JSON.parse(raw);
      return parsed && parsed.v === 1 ? parsed : null;
    } catch (_) {
      return null;
    }
  }

  function rememberTrustedDevice(user) {
    if (!user?.uid) return;
    try {
      localStorage.setItem(TRUSTED_DEVICE_KEY, JSON.stringify({
        v: 1,
        uid: String(user.uid),
        email: String(user.email || "").toLowerCase(),
        lastAuthenticatedAt: Date.now()
      }));
    } catch (_) {}
  }

  function forgetTrustedDevice() {
    try { localStorage.removeItem(TRUSTED_DEVICE_KEY); } catch (_) {}
  }

  function isUsableUser(user) {
    if (!user) return false;
    if (user.email && user.emailVerified === false) return false;
    return true;
  }

  function getFirebaseAuth() {
    try {
      return window.firebase && typeof firebase.auth === "function" ? firebase.auth() : null;
    } catch (_) {
      return null;
    }
  }

  function hideLoginGate() {
    const gate = document.getElementById("auth-gate");
    if (!gate) return;
    gate.hidden = true;
    gate.classList.add("hidden");
    gate.style.setProperty("display", "none", "important");
    gate.setAttribute("aria-hidden", "true");
  }

  function hideStartupLoader() {
    const loader = document.getElementById("app-startup-loading");
    if (!loader) return;
    loader.classList.add("hidden");
    loader.hidden = true;
    loader.setAttribute("aria-hidden", "true");
  }

  function setStartupMessage(text, emoji = "🔐") {
    const loader = document.getElementById("app-startup-loading");
    if (!loader) return;
    const emojiNode = loader.querySelector("[data-loading-humor-emoji]");
    const messageNode = loader.querySelector("[data-loading-humor-message]");
    if (emojiNode) emojiNode.textContent = emoji;
    if (messageNode) messageNode.textContent = text;
  }

  function revealLoginGate(message = "Accesso pronto. Inserisci email e password.") {
    const gate = document.getElementById("auth-gate");
    if (!gate) return;
    hideStartupLoader();
    gate.hidden = false;
    gate.classList.remove("hidden");
    gate.style.removeProperty("display");
    gate.removeAttribute("aria-hidden");
    const feedback = document.getElementById("auth-email-feedback");
    if (feedback && !feedback.textContent.trim()) feedback.textContent = message;
    window.__heraStartupLoginFailsafeUsed = true;
  }

  function onAuthenticated(user) {
    authenticatedUser = user;
    authStateResolved = true;
    rememberTrustedDevice(user);
    hideLoginGate();
  }

  function onUnauthenticated() {
    authenticatedUser = null;
    authStateResolved = true;
    if (readTrustedDevice()) {
      revealLoginGate("La sessione salvata su questo dispositivo non è più valida. Accedi di nuovo una sola volta.");
    }
  }

  function startAuthWatch() {
    if (authWatchStarted) return;
    authWatchStarted = true;
    const startedAt = Date.now();

    const attach = () => {
      const auth = getFirebaseAuth();
      if (!auth) {
        if (Date.now() - startedAt < AUTH_WATCH_MAX_MS) {
          window.setTimeout(attach, 100);
        }
        return;
      }

      try {
        const current = auth.currentUser;
        if (isUsableUser(current)) onAuthenticated(current);

        if (typeof auth.onIdTokenChanged === "function") {
          auth.onIdTokenChanged((user) => {
            if (isUsableUser(user)) onAuthenticated(user);
            else onUnauthenticated();
          }, () => {
            authStateResolved = true;
          });
        } else if (typeof auth.onAuthStateChanged === "function") {
          auth.onAuthStateChanged((user) => {
            if (isUsableUser(user)) onAuthenticated(user);
            else onUnauthenticated();
          }, () => {
            authStateResolved = true;
          });
        }
      } catch (_) {
        authStateResolved = true;
      }
    };

    attach();
  }

  function isVisible(surface) {
    return surface.closest(".hidden, [hidden], [aria-hidden='true']") === null;
  }

  function setMessage(surface, index) {
    const emoji = surface.querySelector("[data-loading-humor-emoji]");
    const message = surface.querySelector("[data-loading-humor-message]");
    const next = messages[index % messages.length];
    if (emoji) emoji.textContent = next[0];
    if (message) message.textContent = next[1];
  }

  function stop(surface) {
    const controller = controllers.get(surface);
    if (!controller) return;
    window.clearInterval(controller.rotationTimer);
    window.clearTimeout(controller.slowTimer);
    controllers.delete(surface);
  }

  function start(surface) {
    if (controllers.has(surface) || !isVisible(surface)) return;
    let index = 0;
    const slowNotice = surface.querySelector("[data-loading-humor-slow]");
    setMessage(surface, index);
    slowNotice?.classList.add("hidden");
    const rotationTimer = window.setInterval(() => {
      index = (index + 1) % messages.length;
      setMessage(surface, index);
    }, ROTATION_MS);
    const slowTimer = window.setTimeout(() => {
      if (!isVisible(surface)) return;
      window.clearInterval(rotationTimer);
      const emoji = surface.querySelector("[data-loading-humor-emoji]");
      const message = surface.querySelector("[data-loading-humor-message]");
      if (emoji) emoji.textContent = "🐌";
      if (message) message.textContent = "Connessione lenta… controllo l’accesso salvato.";
      slowNotice?.classList.remove("hidden");
    }, SLOW_NOTICE_MS);
    controllers.set(surface, { rotationTimer, slowTimer });
  }

  function sync(surface) {
    if (isVisible(surface)) start(surface);
    else stop(surface);
  }

  function revealLoginFailsafe() {
    const auth = getFirebaseAuth();
    const user = authenticatedUser || auth?.currentUser || null;
    if (isUsableUser(user)) {
      onAuthenticated(user);
      return;
    }

    const trusted = readTrustedDevice();
    if (trusted && !authStateResolved) {
      hideLoginGate();
      setStartupMessage("Ripristino accesso automatico…", "🔐");
      window.setTimeout(() => {
        const latestAuth = getFirebaseAuth();
        const latestUser = authenticatedUser || latestAuth?.currentUser || null;
        if (isUsableUser(latestUser)) {
          onAuthenticated(latestUser);
          return;
        }
        if (authStateResolved) {
          revealLoginGate("La sessione salvata non è più valida. Accedi di nuovo una sola volta.");
        }
      }, Math.max(1000, AUTH_WATCH_MAX_MS - LOGIN_FAILSAFE_MS));
      return;
    }

    if (!trusted || authStateResolved) revealLoginGate();
  }

  function install() {
    startAuthWatch();

    document.addEventListener("click", (event) => {
      const logout = event.target?.closest?.("#logout-btn,#access-approval-logout,[data-logout]");
      if (logout) forgetTrustedDevice();
    }, true);

    document.querySelectorAll("[data-loading-humor-surface]").forEach((surface) => {
      surface.querySelector("[data-loading-humor-retry]")?.addEventListener("click", () => window.location.reload());
      const visibilityObserver = new MutationObserver(() => sync(surface));
      let observedNode = surface;
      while (observedNode && observedNode !== document.body) {
        visibilityObserver.observe(observedNode, {
          attributes: true,
          attributeFilter: ["class", "hidden", "aria-hidden"]
        });
        observedNode = observedNode.parentElement;
      }
      sync(surface);
    });
    window.setTimeout(revealLoginFailsafe, LOGIN_FAILSAFE_MS);
  }

  window.HeraTrustedDeviceSession = {
    installed: true,
    hasTrustedDevice: () => Boolean(readTrustedDevice()),
    forget: forgetTrustedDevice
  };

  if (document.readyState === "loading") document.addEventListener("DOMContentLoaded", install, { once: true });
  else install();
})();

(() => {
  "use strict";

  let pendingCredential = null;
  let authUnsubscribe = null;

  function normalizeEmail(value) {
    return String(value || "").trim().toLowerCase();
  }

  function isNativeAndroid() {
    try {
      return Boolean(window.Capacitor?.isNativePlatform?.() && window.Capacitor?.getPlatform?.() === "android");
    } catch (_) {
      return false;
    }
  }

  async function storeCredential(email, password) {
    const normalizedEmail = normalizeEmail(email);
    if (!normalizedEmail || !password) return;

    try { localStorage.setItem("heraSavedLoginEmailV1", normalizedEmail); } catch (_) {}

    if (isNativeAndroid()) {
      try {
        const vault = window.Capacitor?.Plugins?.HeraCredentialVault;
        if (vault?.storeCredential) {
          await vault.storeCredential({ email: normalizedEmail, password: String(password) });
          return;
        }
      } catch (error) {
        console.warn("Salvataggio credenziale Android non riuscito:", error);
      }
    }

    try {
      if (window.PasswordCredential && navigator.credentials?.store) {
        const credential = new PasswordCredential({ id: normalizedEmail, name: normalizedEmail, password: String(password) });
        await navigator.credentials.store(credential);
      }
    } catch (error) {
      console.info("Password manager browser non disponibile; resta attivo autocomplete.", error);
    }
  }

  function attachAuthWatcher() {
    if (authUnsubscribe) return;
    try {
      const auth = window.firebase && typeof firebase.auth === "function" ? firebase.auth() : null;
      if (!auth?.onAuthStateChanged) return;
      authUnsubscribe = auth.onAuthStateChanged((user) => {
        const pending = pendingCredential;
        if (!pending || !user?.email) return;
        if (normalizeEmail(user.email) !== pending.email) return;
        if (Date.now() - pending.capturedAt > 120000) {
          pendingCredential = null;
          return;
        }
        pendingCredential = null;
        void storeCredential(pending.email, pending.password);
      });
    } catch (_) {}
  }

  function captureLogin(event) {
    const form = event.target;
    if (!(form instanceof HTMLFormElement) || form.id !== "auth-email-form") return;
    const email = normalizeEmail(document.getElementById("auth-email-input")?.value);
    const password = String(document.getElementById("auth-password-input")?.value || "");
    if (!email || !password) return;
    pendingCredential = { email, password, capturedAt: Date.now() };
    attachAuthWatcher();
  }

  function prefillSavedEmail() {
    const emailInput = document.getElementById("auth-email-input");
    if (!emailInput || emailInput.value) return;
    try {
      const savedEmail = localStorage.getItem("heraSavedLoginEmailV1") || "";
      if (savedEmail) emailInput.value = savedEmail;
    } catch (_) {}
  }

  function install() {
    document.addEventListener("submit", captureLogin, true);
    prefillSavedEmail();
    document.getElementById("auth-email-input")?.setAttribute("autocomplete", "username");
    document.getElementById("auth-password-input")?.setAttribute("autocomplete", "current-password");
    document.getElementById("auth-email-form")?.setAttribute("autocomplete", "on");
  }

  if (document.readyState === "loading") document.addEventListener("DOMContentLoaded", install, { once: true });
  else install();
})();

(() => {
  "use strict";

  function addStyle(href, dataName) {
    if (document.querySelector(`link[data-${dataName}]`)) return;
    const link = document.createElement("link");
    link.rel = "stylesheet";
    link.href = href;
    link.setAttribute(`data-${dataName}`, "1");
    document.head.appendChild(link);
  }

  function addScript(src, dataName, defer = true) {
    if (document.querySelector(`script[data-${dataName}]`)) return;
    const script = document.createElement("script");
    script.src = src;
    script.defer = defer;
    script.setAttribute(`data-${dataName}`, "1");
    document.head.appendChild(script);
  }

  function loadOptionalRuntime() {
    addStyle("./impianti-zebra-fatto-guard.css?v=20260817b", "impianti-zebra-style");
    addStyle("./desktop-fullscreen.css?v=20260817a", "desktop-fullscreen-style");
    addStyle("./squadre-commessa-themes.css?v=20260817b", "squadre-commessa-themes");
    addStyle("./squadre-commessa-themes-visible-fix.css?v=20260817d", "squadre-commessa-themes-visible-fix");
    addScript("./fatto-scroll-guard.js?v=20260817a", "fatto-scroll-guard");
    addScript("./client-error-reporter.js?v=20260816a", "client-error-reporter");
    addScript("./app-error-monitor.js?v=20260824b", "app-error-monitor");
    addScript("./admin-error-center.js?v=20260824b", "admin-error-center");
    addStyle("./admin-error-center.css?v=20260824b", "admin-error-center-style");
    addScript("./squadre-commessa-themes.js?v=20260817b", "squadre-commessa-themes-script");
    addScript("./shared-pdf-attachments.js?v=20260821v2", "shared-whazzup-pdf");

    const isNative = Boolean(window.Capacitor?.isNativePlatform?.());
    if (!isNative) addScript("./pwa-whazzup-continuous-camera.js?v=20260817a", "pwa-whazzup-camera");
  }

  function scheduleOptionalRuntime() {
    const run = () => {
      if ("requestIdleCallback" in window) window.requestIdleCallback(loadOptionalRuntime, { timeout: 2500 });
      else window.setTimeout(loadOptionalRuntime, 1200);
    };
    if (document.readyState === "complete") run();
    else window.addEventListener("load", run, { once: true });
  }

  scheduleOptionalRuntime();
})();

(() => {
  "use strict";
  const REGION = "europe-west1";
  let attempts = 0;

  function attach() {
    const dialog = document.getElementById("hera-error-center-dialog");
    const actions = dialog?.querySelector?.(".hera-error-head-actions");
    if (!dialog || !actions) {
      if (attempts < 30) {
        attempts += 1;
        window.setTimeout(attach, 500);
      }
      return;
    }
    if (actions.querySelector("[data-error-reset]")) return;
    const button = document.createElement("button");
    button.type = "button";
    button.className = "btn";
    button.dataset.errorReset = "1";
    button.textContent = "AZZERA";
    button.style.background = "#b91c1c";
    button.style.borderColor = "#b91c1c";
    button.style.color = "#fff";
    const closeButton = actions.querySelector("[data-error-close]");
    actions.insertBefore(button, closeButton || null);
    button.addEventListener("click", async () => {
      if (button.disabled) return;
      if (!window.confirm("Azzerare definitivamente tutti gli errori e i contatori del Centro errori? Questa operazione cancella solo la diagnostica del Centro errori e non modifica commesse, impianti, FATTO o WHAZZUP.")) return;
      const originalText = button.textContent;
      button.disabled = true;
      button.textContent = "AZZERAMENTO…";
      try {
        if (!window.firebase?.apps?.length || !window.firebase?.functions) throw new Error("Firebase non disponibile.");
        const callable = window.firebase.app().functions(REGION).httpsCallable("resetErrorCenter");
        const response = await callable({});
        const deleted = Math.max(0, Number(response?.data?.deletedGroups || 0));
        await window.HeraAdminErrorCenter?.refresh?.();
        window.alert(`✅ Centro errori azzerato. Eliminati ${deleted} gruppi di errore.`);
      } catch (error) {
        console.error("Azzeramento Centro errori fallito:", error);
        window.alert(`⚠️ Azzeramento non riuscito: ${error?.message || "errore sconosciuto"}`);
      } finally {
        button.disabled = false;
        button.textContent = originalText;
      }
    });
  }

  window.addEventListener("load", () => window.setTimeout(attach, 1800), { once: true });
})();