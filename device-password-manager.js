(() => {
  "use strict";

  if (window.HeraDevicePasswordManager?.installed) return;

  const LEGACY_STORAGE_KEY = "heraSavedLoginAccountsV1";
  let pendingCredential = null;
  let authHookInstalled = false;
  let formHookInstalled = false;

  function normalizeEmail(value) {
    return String(value || "").trim().toLowerCase();
  }

  function isNativeAndroid() {
    try {
      return Boolean(
        window.Capacitor?.isNativePlatform?.()
        && window.Capacitor?.getPlatform?.() === "android"
      );
    } catch (_) {
      return false;
    }
  }

  function getAndroidVault() {
    try {
      return window.Capacitor?.Plugins?.HeraCredentialVault || null;
    } catch (_) {
      return null;
    }
  }

  function removeLegacyPlaintextPasswords() {
    try {
      if (localStorage.getItem(LEGACY_STORAGE_KEY) !== null) {
        localStorage.removeItem(LEGACY_STORAGE_KEY);
        console.info("[device-password-manager] Rimosso archivio password legacy da localStorage.");
      }
    } catch (_) {}
  }

  function configureLoginFields() {
    const form = document.getElementById("auth-email-form");
    const email = document.getElementById("auth-email-input");
    const password = document.getElementById("auth-password-input");
    if (!form || !email || !password) return false;

    form.setAttribute("autocomplete", "on");
    email.setAttribute("name", "username");
    email.setAttribute("autocomplete", "username");
    email.setAttribute("autocapitalize", "none");
    email.setAttribute("spellcheck", "false");
    email.setAttribute("inputmode", "email");
    password.setAttribute("name", "password");
    password.setAttribute("autocomplete", "current-password");
    password.setAttribute("autocapitalize", "none");
    password.setAttribute("spellcheck", "false");

    if (!formHookInstalled) {
      formHookInstalled = true;
      form.addEventListener("submit", () => {
        const emailValue = normalizeEmail(email.value);
        const passwordValue = String(password.value || "");
        pendingCredential = emailValue && passwordValue
          ? { email: emailValue, password: passwordValue, capturedAt: Date.now() }
          : null;
      }, true);
    }
    return true;
  }

  async function storeWithBrowserPasswordManager(email, password) {
    if (!window.PasswordCredential || !navigator.credentials?.store) return false;
    try {
      const credential = new PasswordCredential({ id: email, name: email, password });
      await navigator.credentials.store(credential);
      console.info("[device-password-manager] Credenziale affidata al password manager del dispositivo/browser.");
      return true;
    } catch (error) {
      console.info("[device-password-manager] Salvataggio gestito dal browser tramite autocomplete.", error);
      return false;
    }
  }

  async function storeWithAndroidVault(email, password) {
    if (!isNativeAndroid()) return false;
    const vault = getAndroidVault();
    if (!vault?.storeCredential) return false;
    try {
      await vault.storeCredential({ email, password });
      console.info("[device-password-manager] Credenziale salvata nel vault sicuro Android.");
      return true;
    } catch (error) {
      console.warn("[device-password-manager] Vault Android non disponibile:", error);
      return false;
    }
  }

  async function saveAfterSuccessfulLogin(user) {
    const pending = pendingCredential;
    pendingCredential = null;
    if (!pending || !user?.email) return;
    if (Date.now() - pending.capturedAt > 120000) return;
    if (normalizeEmail(user.email) !== pending.email) return;

    if (await storeWithAndroidVault(pending.email, pending.password)) return;
    await storeWithBrowserPasswordManager(pending.email, pending.password);
  }

  function installAuthHook() {
    if (authHookInstalled) return true;
    try {
      if (!window.firebase || typeof firebase.auth !== "function") return false;
      const auth = firebase.auth();
      if (!auth?.onAuthStateChanged) return false;
      authHookInstalled = true;
      auth.onAuthStateChanged((user) => {
        if (user) void saveAfterSuccessfulLogin(user);
      });
      return true;
    } catch (_) {
      return false;
    }
  }

  function install() {
    removeLegacyPlaintextPasswords();
    configureLoginFields();
    installAuthHook();

    const observer = new MutationObserver(() => {
      configureLoginFields();
      installAuthHook();
    });
    observer.observe(document.documentElement, { childList: true, subtree: true });

    let attempts = 0;
    const timer = window.setInterval(() => {
      attempts += 1;
      configureLoginFields();
      if ((installAuthHook() && formHookInstalled) || attempts >= 80) window.clearInterval(timer);
    }, 250);
  }

  window.HeraDevicePasswordManager = {
    installed: true,
    configureLoginFields,
    removeLegacyPlaintextPasswords
  };

  if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", install, { once: true });
  } else {
    install();
  }
})();
