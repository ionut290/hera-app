(() => {
  "use strict";

  const STORAGE_KEY = "heraSavedLoginAccountsV1";
  let unlockedUntil = 0;

  function isNativeAndroid() {
    try { return Boolean(window.Capacitor?.isNativePlatform?.() && window.Capacitor?.getPlatform?.() === "android"); }
    catch (_) { return false; }
  }

  function biometricPlugin() {
    try { return window.Capacitor?.Plugins?.HeraBiometric || null; }
    catch (_) { return null; }
  }

  function readAccounts() {
    try {
      const data = JSON.parse(localStorage.getItem(STORAGE_KEY) || "[]");
      return Array.isArray(data) ? data.filter((item) => item?.email && item?.password) : [];
    } catch (_) { return []; }
  }

  function writeAccounts(items) {
    localStorage.setItem(STORAGE_KEY, JSON.stringify(items));
  }

  function upsertAccount(email, password) {
    email = String(email || "").trim().toLowerCase();
    password = String(password || "");
    if (!email || !password) return;
    const next = readAccounts().filter((item) => String(item.email).toLowerCase() !== email);
    next.unshift({ email, password, savedAt: Date.now() });
    writeAccounts(next.slice(0, 10));
  }

  async function requireBiometric() {
    if (!isNativeAndroid()) {
      const ok = window.confirm("Le password sono salvate solo su questo browser/dispositivo. Continuare?");
      if (!ok) throw new Error("Operazione annullata.");
      unlockedUntil = Date.now() + 60000;
      return;
    }
    if (Date.now() < unlockedUntil) return;
    const plugin = biometricPlugin();
    if (!plugin) throw new Error("Protezione biometrica Android non disponibile.");
    let status = null;
    try { status = await plugin.status(); } catch (_) {}
    if (!status?.available) throw new Error("Configura impronta o riconoscimento biometrico sul telefono.");
    if (!status?.enabled) {
      await plugin.enable({ title: "Proteggi password salvate", subtitle: "Conferma la tua identità" });
    } else {
      await plugin.authenticate({ title: "Password salvate", subtitle: "Conferma la tua identità" });
    }
    unlockedUntil = Date.now() + 60000;
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
    wrap.innerHTML = '<label class="saved-password-remember"><input id="saved-password-remember" type="checkbox"> Memorizza password su questo dispositivo</label><button id="saved-password-open-btn" class="btn" type="button">🔐 PASSWORD SALVATE</button>';
    const feedback = document.getElementById("auth-email-feedback");
    form.insertBefore(wrap, feedback || null);
    wrap.querySelector("#saved-password-open-btn")?.addEventListener("click", () => void openVault());

    form.addEventListener("submit", () => {
      if (!document.getElementById("saved-password-remember")?.checked) return;
      const email = document.getElementById("auth-email-input")?.value || "";
      const password = document.getElementById("auth-password-input")?.value || "";
      if (email && password) upsertAccount(email, password);
    }, true);
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
    setFeedback("Verifica identità...");
    try {
      await requireBiometric();
      renderVault();
      setFeedback("");
    } catch (error) {
      setFeedback(error?.message || "Impossibile aprire le password salvate.");
    }
  }

  function renderVault() {
    const list = document.getElementById("saved-password-vault-list");
    if (!list) return;
    const items = readAccounts();
    if (!items.length) {
      list.innerHTML = '<p class="muted">Nessuna password salvata.</p>';
      return;
    }
    list.innerHTML = "";
    items.forEach((item) => {
      const row = document.createElement("div");
      row.className = "saved-password-item";
      row.innerHTML = `<strong>${escapeHtml(item.email)}</strong><div class="saved-password-value" data-password>••••••••</div><div class="saved-password-actions"><button class="btn" type="button" data-action="show">👁 MOSTRA</button><button class="btn" type="button" data-action="copy">COPIA</button><button class="btn btn-primary" type="button" data-action="use">USA PER ACCEDERE</button><button class="btn" type="button" data-action="delete">ELIMINA</button></div>`;
      row.addEventListener("click", (event) => void handleAction(event, item));
      list.appendChild(row);
    });
  }

  async function handleAction(event, item) {
    const action = event.target?.closest?.("[data-action]")?.dataset?.action;
    if (!action) return;
    try { await requireBiometric(); }
    catch (error) { setFeedback(error?.message || "Verifica non riuscita."); return; }
    const row = event.currentTarget;
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
      document.getElementById("saved-password-vault-dialog")?.close();
      password?.focus();
    } else if (action === "delete") {
      writeAccounts(readAccounts().filter((saved) => saved.email !== item.email));
      renderVault();
      setFeedback("Password eliminata da questo dispositivo.");
    }
  }

  function install() {
    ensureUi();
    const observer = new MutationObserver(ensureUi);
    observer.observe(document.body, { childList: true, subtree: true });
  }

  if (document.readyState === "loading") document.addEventListener("DOMContentLoaded", install, { once: true });
  else install();
})();
