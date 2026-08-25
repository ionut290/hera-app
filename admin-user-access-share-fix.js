(() => {
  "use strict";

  const VERSION = "1.1.0";
  const SEARCH_LIST_IDS = [
    "pending-users-list",
    "admin-users-list",
    "user-ban-list",
    "user-permissions-list"
  ];
  const TEMP_LOGIN_HASH_KEY = "temp-login";
  const ANDROID_DEEP_LINK_SCHEME = "vargacantieri";
  let observer = null;
  let scheduled = false;

  function normalize(value) {
    return String(value || "")
      .trim()
      .toLowerCase()
      .normalize("NFD")
      .replace(/[\u0300-\u036f]/g, "")
      .replace(/\s+/g, " ");
  }

  function isManager() {
    try {
      return typeof canManageData === "function" && canManageData();
    } catch (_) {
      return false;
    }
  }

  function encodePayload(payload) {
    const json = JSON.stringify(payload);
    const bytes = new TextEncoder().encode(json);
    let binary = "";
    bytes.forEach((byte) => { binary += String.fromCharCode(byte); });
    return window.btoa(binary).replace(/\+/g, "-").replace(/\//g, "_").replace(/=+$/g, "");
  }

  function decodePayload(value) {
    const normalized = String(value || "").replace(/-/g, "+").replace(/_/g, "/");
    const padded = normalized + "=".repeat((4 - (normalized.length % 4)) % 4);
    const binary = window.atob(padded);
    const bytes = Uint8Array.from(binary, (char) => char.charCodeAt(0));
    return JSON.parse(new TextDecoder().decode(bytes));
  }

  function buildTemporaryLoginPayload(email, password) {
    return encodePayload({
      email: String(email || "").trim().toLowerCase(),
      password: String(password || "")
    });
  }

  function buildTemporaryLoginLink(email, password) {
    const url = new URL(window.location.href);
    url.search = "";
    url.hash = `${TEMP_LOGIN_HASH_KEY}=${encodeURIComponent(buildTemporaryLoginPayload(email, password))}`;
    return url.toString();
  }

  function buildAndroidLoginLink(email, password) {
    const payload = buildTemporaryLoginPayload(email, password);
    return `${ANDROID_DEEP_LINK_SCHEME}://login?${TEMP_LOGIN_HASH_KEY}=${encodeURIComponent(payload)}`;
  }

  function buildCopyPasswordLink(password) {
    const url = new URL("copy-password.html", window.location.href);
    url.searchParams.set("password", String(password || ""));
    return url.toString();
  }

  function readTemporaryLoginFromHash() {
    const hash = String(window.location.hash || "").replace(/^#/, "");
    if (!hash) return null;
    const params = new URLSearchParams(hash);
    const encoded = params.get(TEMP_LOGIN_HASH_KEY);
    if (!encoded) return null;
    try {
      const payload = decodePayload(encoded);
      if (!payload?.email || !payload?.password) return null;
      return {
        email: String(payload.email).trim().toLowerCase(),
        password: String(payload.password)
      };
    } catch (error) {
      console.warn("Link password temporanea non valido:", error);
      return null;
    }
  }

  function clearTemporaryLoginHash() {
    if (!String(window.location.hash || "").includes(TEMP_LOGIN_HASH_KEY)) return;
    try {
      window.history.replaceState(null, document.title, `${window.location.pathname}${window.location.search}`);
    } catch (_) {
      window.location.hash = "";
    }
  }

  function applyTemporaryLoginLink() {
    const payload = readTemporaryLoginFromHash();
    if (!payload) return;

    const apply = () => {
      const emailInput = document.getElementById("auth-email-input");
      const passwordInput = document.getElementById("auth-password-input");
      const feedback = document.getElementById("auth-email-feedback");
      if (!emailInput || !passwordInput) return false;

      emailInput.value = payload.email;
      passwordInput.value = payload.password;
      emailInput.dispatchEvent(new Event("input", { bubbles: true }));
      passwordInput.dispatchEvent(new Event("input", { bubbles: true }));
      if (feedback) feedback.textContent = "Credenziali temporanee inserite. Premi Entra per accedere.";
      clearTemporaryLoginHash();
      window.setTimeout(() => passwordInput.focus({ preventScroll: true }), 0);
      return true;
    };

    if (apply()) return;
    let attempts = 0;
    const timer = window.setInterval(() => {
      attempts += 1;
      if (apply() || attempts >= 40) window.clearInterval(timer);
    }, 100);
  }

  function getSearchText(item) {
    if (!(item instanceof HTMLElement)) return "";
    const datasetText = Object.values(item.dataset || {}).join(" ");
    const nestedDatasetText = Array.from(item.querySelectorAll("[data-user-id],[data-uid],[data-user],[data-ban-user],[data-user-ban],[data-user-permission]"))
      .map((node) => Object.values(node.dataset || {}).join(" "))
      .join(" ");
    return normalize(`${item.textContent || ""} ${datasetText} ${nestedDatasetText}`);
  }

  function forceItemVisibility(item, visible) {
    if (!(item instanceof HTMLElement)) return;
    if (visible) {
      item.hidden = false;
      item.removeAttribute("aria-hidden");
      item.style.removeProperty("display");
      item.style.removeProperty("visibility");
    } else {
      item.hidden = true;
      item.setAttribute("aria-hidden", "true");
      item.style.setProperty("display", "none", "important");
    }
  }

  function applyFixedSearch() {
    if (!isManager()) return;
    const input = document.getElementById("user-management-search-input");
    const feedback = document.getElementById("user-management-search-feedback");
    if (!input) return;

    const query = normalize(input.value);
    let visible = 0;

    SEARCH_LIST_IDS.forEach((id) => {
      const list = document.getElementById(id);
      if (!list) return;
      let visibleInList = 0;
      Array.from(list.children).forEach((item) => {
        if (!(item instanceof HTMLElement)) return;
        const matches = !query || getSearchText(item).includes(query);
        forceItemVisibility(item, matches);
        if (matches) {
          visible += 1;
          visibleInList += 1;
        }
      });

      if (query && list.children.length > 0 && visibleInList === 0) {
        list.setAttribute("data-search-empty", "true");
      } else {
        list.removeAttribute("data-search-empty");
      }
    });

    if (!feedback) return;
    if (!query) feedback.textContent = "";
    else if (visible === 1) feedback.textContent = "1 risultato trovato.";
    else if (visible > 1) feedback.textContent = `${visible} risultati trovati.`;
    else feedback.textContent = "Nessun utente trovato.";
  }

  function parseDialogProfile() {
    const label = String(document.getElementById("admin-password-private-user")?.textContent || "").trim();
    const parts = label.split("•").map((part) => part.trim()).filter(Boolean);
    const email = String(parts.at(-1) || "").toLowerCase();
    const displayName = parts.length > 1 ? parts.slice(0, -1).join(" • ") : "Utente";
    return { displayName, email };
  }

  function temporaryPasswordValue() {
    return String(document.getElementById("admin-password-private-value")?.value || "");
  }

  function ensureShareActions() {
    const result = document.getElementById("admin-password-private-result");
    if (!result || result.querySelector("#admin-password-private-whatsapp")) return;

    const actions = document.createElement("div");
    actions.className = "actions-row admin-password-share-actions";
    actions.innerHTML = `
      <button id="admin-password-private-whatsapp" class="btn btn-primary" type="button">INVIA SU WHATSAPP</button>
      <button id="admin-password-private-copy-link" class="btn" type="button">COPIA LINK PWA</button>`;
    result.appendChild(actions);

    actions.querySelector("#admin-password-private-whatsapp")?.addEventListener("click", shareTemporaryPasswordWhatsApp);
    actions.querySelector("#admin-password-private-copy-link")?.addEventListener("click", copyTemporaryLoginLink);
  }

  function setDialogFeedback(message) {
    const node = document.getElementById("admin-password-private-feedback");
    if (node) node.textContent = String(message || "");
  }

  function createShareMessage() {
    const password = temporaryPasswordValue();
    const { displayName, email } = parseDialogProfile();
    if (!password || !email) throw new Error("Genera prima la nuova password temporanea.");

    const pwaLink = buildTemporaryLoginLink(email, password);
    const androidLink = buildAndroidLoginLink(email, password);
    const copyPasswordLink = buildCopyPasswordLink(password);
    const greetingName = displayName && displayName !== "Utente" ? ` ${displayName}` : "";
    const text = [
      `Ciao${greetingName},`,
      "ti è stata generata una nuova password temporanea per Varga Cantieri.",
      "",
      `Email: ${email}`,
      `Password temporanea: ${password}`,
      "",
      "📋 COPIA PASSWORD:",
      copyPasswordLink,
      "",
      "🌐 ACCEDI DA PWA:",
      pwaLink,
      "",
      "📱 ACCEDI DA ANDROID:",
      androidLink,
      "",
      "I link di accesso contengono già email e password temporanea.",
      "Se usi COPIA PASSWORD, premi il pulsante Copia password e poi incollala nell'app.",
      "Al primo accesso dovrai scegliere una nuova password personale.",
      "Per sicurezza non inoltrare questo messaggio ad altre persone."
    ].join("\n");
    return { text, link: pwaLink, pwaLink, androidLink, copyPasswordLink };
  }

  function shareTemporaryPasswordWhatsApp() {
    try {
      const { text } = createShareMessage();
      const whatsappUrl = `https://wa.me/?text=${encodeURIComponent(text)}`;
      const opened = window.open(whatsappUrl, "_blank", "noopener,noreferrer");
      if (!opened) window.location.href = whatsappUrl;
      setDialogFeedback("Messaggio WhatsApp preparato con password, copia password, link PWA e link Android.");
    } catch (error) {
      setDialogFeedback(error?.message || "Impossibile preparare il messaggio WhatsApp.");
    }
  }

  async function copyTemporaryLoginLink() {
    try {
      const { pwaLink } = createShareMessage();
      await navigator.clipboard.writeText(pwaLink);
      setDialogFeedback("Link PWA copiato. Contiene la password temporanea: condividilo solo con l'utente interessato.");
    } catch (error) {
      setDialogFeedback(error?.message || "Impossibile copiare il link di accesso.");
    }
  }

  function enhance() {
    ensureShareActions();
    applyFixedSearch();
  }

  function scheduleEnhance() {
    if (scheduled) return;
    scheduled = true;
    window.setTimeout(() => {
      scheduled = false;
      enhance();
    }, 50);
  }

  function initialize() {
    applyTemporaryLoginLink();

    document.addEventListener("input", (event) => {
      if (event.target?.id !== "user-management-search-input") return;
      window.setTimeout(applyFixedSearch, 0);
    }, true);

    document.addEventListener("click", (event) => {
      if (event.target?.closest?.("#user-management-search-clear")) {
        window.setTimeout(applyFixedSearch, 0);
      }
      if (event.target?.closest?.("#open-panel-utenti,[data-admin-private-password],[data-admin-manage-password]")) {
        window.setTimeout(enhance, 80);
      }
    }, true);

    observer = new MutationObserver(scheduleEnhance);
    observer.observe(document.body, { childList: true, subtree: true });
    enhance();
  }

  window.HeraAdminUserAccessShareFix = {
    installed: true,
    version: VERSION,
    refresh: enhance,
    buildTemporaryLoginLink,
    buildAndroidLoginLink,
    buildCopyPasswordLink,
    applyTemporaryLoginLink
  };

  if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", initialize, { once: true });
  } else {
    initialize();
  }
})();

(() => {
  if (window.__heraLoginPasswordVaultLoaderInstalled) return;
  window.__heraLoginPasswordVaultLoaderInstalled = true;
  const script = document.createElement("script");
  script.src = "login-password-vault.js?v=20260825a";
  script.defer = true;
  script.onerror = () => console.warn("Password vault non caricato.");
  document.head.appendChild(script);
})();
