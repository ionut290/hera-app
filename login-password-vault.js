(() => {
  "use strict";

  const STORAGE_KEY = "heraSavedLoginAccountsV1";
  const NATIVE_UNLOCK_MS = 60000;
  let pendingCredential = null;
  let nativeAccountsCache = [];
  let nativeUnlockedUntil = 0;
  let authHookInstalled = false;

  function isNativeAndroid() {
    try { return Boolean(window.Capacitor?.isNativePlatform?.() && window.Capacitor?.getPlatform?.() === "android"); }
    catch (_) { return false; }
  }

  function vaultPlugin() {
    try { return window.Capacitor?.Plugins?.HeraCredentialVault || null; }
    catch (_) { return null; }
  }

  function normalizeEmail(value) {
    return String(value || "").trim().toLowerCase();
  }

  function readLocalAccounts() {
    try {
      const data = JSON.parse(localStorage.getItem(STORAGE_KEY) || "[]");
      return Array.isArray(data) ? data.filter((item) => item?.email && item?.password) : [];
    } catch (_) { return []; }
  }

  function writeLocalAccounts(items) {
    try { localStorage.setItem(STORAGE_KEY, JSON.stringify(items)); } catch (_) {}
  }

  function upsertLocalAccount(email, password) {
    email = normalizeEmail(email);
    password = String(password || "");
    if (!email || !password) return;
    const next = readLocalAccounts().filter((item) => normalizeEmail(item.email) !== email);
    next.unshift({ email, password, savedAt: Date.now() });
    writeLocalAccounts(next.slice(0, 10));
  }

  async function storeAccount(email, password) {
    email = normalizeEmail(email);
    password = String(password || "");
    if (!email || !password) return false;

    if (isNativeAndroid()) {
      const plugin = vaultPlugin();
      if (!plugin?.storeCredential) throw new Error("Vault sicuro Android non disponibile. Aggiorna l'app.");
      await plugin.storeCredential({ email, password });
      nativeAccountsCache = nativeAccountsCache.filter((item) => normalizeEmail(item.email) !== email);
      nativeAccountsCache.unshift({ email, password, savedAt: Date.now() });
      return true;
    }

    upsertLocalAccount(email, password);
    return true;
  }

  async function readAccountsSecure() {
    if (isNativeAndroid()) {
      if (Date.now() < nativeUnlockedUntil && nativeAccountsCache.length) return nativeAccountsCache;
      const plugin = vaultPlugin();
      if (!plugin?.listCredentials) throw new Error("Vault sicuro Android non disponibile. Aggiorna l'app.");
      const result = await plugin.listCredentials({ title: "Password salvate", subtitle: "Conferma la tua identità" });
      nativeAccountsCache = Array.isArray(result?.accounts) ? result.accounts : [];
      nativeUnlockedUntil = Date.now() + NATIVE_UNLOCK_MS;
      return nativeAccountsCache;
    }

    const ok = window.confirm("Le password sono salvate solo su questo browser/dispositivo. Continuare?");
    if (!ok) throw new Error("Operazione annullata.");
    return readLocalAccounts();
  }

  async function deleteAccount(email) {
    email = normalizeEmail(email);
    if (!email) return;
    if (isNativeAndroid()) {
      const plugin = vaultPlugin();
      if (!plugin?.deleteCredential) throw new Error("Vault sicuro Android non disponibile.");
      await plugin.deleteCredential({ email });
      nativeAccountsCache = nativeAccountsCache.filter((item) => normalizeEmail(item.email) !== email);
      return;
    }
    writeLocalAccounts(readLocalAccounts().filter((saved) => normalizeEmail(saved.email) !== email));
  }

  async function migrateLegacyAndroidAccounts() {
    if (!isNativeAndroid()) return;
    const plugin = vaultPlugin();
    if (!plugin?.storeCredential) return;
    const legacy = readLocalAccounts();
    if (!legacy.length) return;
    try {
      for (const item of legacy) {
        await plugin.storeCredential({ email: normalizeEmail(item.email), password: String(item.password || "") });
      }
      localStorage.removeItem(STORAGE_KEY);
      console.info(`[credential-vault] Migrate ${legacy.length} credenziali legacy nel vault Android cifrato.`);
    } catch (error) {
      console.warn("Migrazione credenziali legacy Android non completata:", error);
    }
  }

  function rememberEnabled() {
    const checkbox = document.getElementById("saved-password-remember");
    return !checkbox || checkbox.checked;
  }

  function capturePendingCredential(options = {}) {
    if (!rememberEnabled()) {
      pendingCredential = null;
      return false;
    }
    const hasProvidedCredential = options
      && typeof options === "object"
      && Object.prototype.hasOwnProperty.call(options, "email");
    const email = normalizeEmail(hasProvidedCredential
      ? options.email
      : document.getElementById("auth-email-input")?.value || "");
    const password = String(hasProvidedCredential
      ? options.password || ""
      : document.getElementById("auth-password-input")?.value || "");
    pendingCredential = email && password ? { email, password, capturedAt: Date.now() } : null;
    return Boolean(pendingCredential);
  }

  async function savePendingAfterSuccessfulLogin(user) {
    if (!pendingCredential || !user?.email) return;
    const userEmail = normalizeEmail(user.email);
    if (!userEmail || userEmail !== pendingCredential.email) return;
    if (Date.now() - pendingCredential.capturedAt > 120000) {
      pendingCredential = null;
      return;
    }

    const credential = pendingCredential;
    pendingCredential = null;
    try {
      await storeAccount(credential.email, credential.password);
      console.info("[credential-vault] Password aggiornata nel vault locale dopo login riuscito.");
    } catch (error) {
      console.warn("Salvataggio password nel vault locale non riuscito:", error);
    }
  }

  function installFirebaseSuccessHook() {
    if (authHookInstalled) return true;
    try {
      if (!window.firebase || typeof firebase.auth !== "function") return false;
      const auth = firebase.auth();
      if (!auth?.onAuthStateChanged) return false;
      authHookInstalled = true;
      auth.onAuthStateChanged((user) => {
        if (user) void savePendingAfterSuccessfulLogin(user);
      });
      return true;
    } catch (_) {
      return false;
    }
  }

  function escapeHtml(value) {
    return String(value || "").replace(/[&<>"']/g, (char) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[char]));
  }

  function ensureStyles() {
    if (document.getElementById("saved-password-vault-style")) return;
    const style = document.createElement("style");
    style.id = "saved-password-vault-style";
    style.textContent = ".saved-password-tools{display:grid;gap:8px;margin:10px 0}.saved-password-remember{display:flex;gap:8px;align-items:center;font-size:.95rem}.saved-password-vault-list{display:grid;gap:10px;margin-top:12px}.saved-password-item{border:1px solid rgba(127,127,127,.25);border-radius:12px;padding:12px}.saved-password-value{font-family:monospace;word-break:break-all;margin:8px 0}.saved-password-actions{display:flex;flex-wrap:wrap;gap:7px}";
    document.head.appendChild(style);
  }

  function ensureUi() {
    const form = document.getElementById("auth-email-form");
    if (!form || document.getElementById("saved-password-open-btn")) return;
    ensureStyles();
    const wrap = document.createElement("div");
    wrap.className = "saved-password-tools";
    wrap.innerHTML = '<label class="saved-password-remember"><input id="saved-password-remember" type="checkbox" checked> Memorizza password su questo dispositivo</label><button id="saved-password-open-btn" class="btn" type="button">🔐 PASSWORD SALVATE</button>';
    const feedback = document.getElementById("auth-email-feedback");
    form.insertBefore(wrap, feedback || null);
    wrap.querySelector("#saved-password-open-btn")?.addEventListener("click", () => void openVault());
    form.addEventListener("submit", capturePendingCredential, true);
  }

  function ensureDialog() {
    let dialog = document.getElementById("saved-password-vault-dialog");
    if (dialog) return dialog;
    dialog = document.createElement("dialog");
    dialog.id = "saved-password-vault-dialog";
    dialog.className = "biometric-dialog";
    dialog.innerHTML = '<form method="dialog"><h2>🔐 Password salvate</h2><p class="muted">Salvate solo su questo dispositivo.</p><div id="saved-password-vault-list" class="saved-password-vault-list"></div><p id="saved-password-vault-feedback" class="muted"></p><div class="actions-row"><button class="btn" value="cancel">CHIUDI</button></div></form>';
    document.body.appendChild(dialog);
    return dialog;
  }

  function setFeedback(message) {
    const node = document.getElementById("saved-password-vault-feedback");
    if (node) node.textContent = String(message || "");
  }

  async function openVault() {
    const dialog = ensureDialog();
    if (!dialog.open) dialog.showModal();
    setFeedback(isNativeAndroid() ? "Verifica biometrica..." : "Apertura password salvate...");
    try {
      const items = await readAccountsSecure();
      renderVault(items);
      setFeedback("");
    } catch (error) {
      setFeedback(error?.message || "Impossibile aprire le password salvate.");
    }
  }

  function renderVault(items) {
    const list = document.getElementById("saved-password-vault-list");
    if (!list) return;
    items = Array.isArray(items) ? items : [];
    if (!items.length) {
      list.innerHTML = '<p class="muted">Nessuna password salvata.</p>';
      return;
    }
    list.innerHTML = "";
    items.forEach((item) => {
      const row = document.createElement("div");
      row.className = "saved-password-item";
      row.innerHTML = `<strong>${escapeHtml(item.email)}</strong><div class="saved-password-value" data-password>••••••••</div><div class="saved-password-actions"><button class="btn" type="button" data-action="show">👁 MOSTRA</button><button class="btn" type="button" data-action="copy">COPIA</button><button class="btn btn-primary" type="button" data-action="use">USA PER ACCEDERE</button><button class="btn" type="button" data-action="delete">ELIMINA</button></div>`;
      row.addEventListener("click", (event) => void handleAction(event, item, row));
      list.appendChild(row);
    });
  }

  async function handleAction(event, item, row) {
    const action = event.target?.closest?.("[data-action]")?.dataset?.action;
    if (!action) return;
    try {
      if (isNativeAndroid() && Date.now() >= nativeUnlockedUntil) {
        const refreshed = await readAccountsSecure();
        const fresh = refreshed.find((saved) => normalizeEmail(saved.email) === normalizeEmail(item.email));
        if (fresh) item = fresh;
      }
      if (action === "show") {
        const node = row.querySelector("[data-password]");
        node.textContent = node.textContent === item.password ? "••••••••" : item.password;
      } else if (action === "copy") {
        await navigator.clipboard.writeText(item.password);
        setFeedback("Password copiata.");
      } else if (action === "use") {
        const email = document.getElementById("auth-email-input");
        const password = document.getElementById("auth-password-input");
        if (email) email.value = item.email;
        if (password) password.value = item.password;
        const remember = document.getElementById("saved-password-remember");
        if (remember) remember.checked = true;
        document.getElementById("saved-password-vault-dialog")?.close();
        password?.focus();
      } else if (action === "delete") {
        await deleteAccount(item.email);
        const items = isNativeAndroid() ? nativeAccountsCache : readLocalAccounts();
        renderVault(items);
        setFeedback("Password eliminata da questo dispositivo.");
      }
    } catch (error) {
      setFeedback(error?.message || "Operazione non riuscita.");
    }
  }

  function install() {
    ensureUi();
    void migrateLegacyAndroidAccounts();
    installFirebaseSuccessHook();
    const observer = new MutationObserver(() => {
      ensureUi();
      installFirebaseSuccessHook();
    });
    observer.observe(document.body, { childList: true, subtree: true });
    let attempts = 0;
    const authTimer = window.setInterval(() => {
      attempts += 1;
      if (installFirebaseSuccessHook() || attempts >= 60) window.clearInterval(authTimer);
    }, 250);
  }

  window.HeraLoginCredentialVault = {
    installed: true,
    storeAccount,
    capturePendingCredential,
    migrateLegacyAndroidAccounts
  };

  if (document.readyState === "loading") document.addEventListener("DOMContentLoaded", install, { once: true });
  else install();
})();
