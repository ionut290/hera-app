(function installAppUpdateButton() {
  "use strict";

  const PLAY_STORE_URL = "https://play.google.com/store/apps/details?id=it.vargacantieri.hera";
  const APP_CACHE_PREFIXES = ["varga-cantieri-shell-", "hera-app-shell-"];
  const DATA_DURABILITY_SRC = "data-durability-runtime.js?v=20260818a";
  const DATA_SAFETY_SRC = "data-safety-layer.js?v=20260819a";
  const CRITICAL_WRITE_SAFETY_SRC = "critical-write-safety-bridge.js?v=20260824-oneclick1";
  const SYNC_BADGE_PENDING_FIX_SRC = "sync-badge-pending-fix.js?v=20260821a";
  const UPDATE_NOTICE_DISMISSED_KEY = "heraUpdateNoticeDismissed";
  const NATIVE_CHECK_INTERVAL_MS = 15 * 60 * 1000;
  const watchedPwaRegistrations = new WeakSet();
  const hadPwaControllerAtStartup = Boolean(navigator.serviceWorker?.controller);
  let lastNativeCheckAt = 0;
  let pendingPwaRegistration = null;

  function ensureDataDurabilityRuntime() {
    if (window.HeraDataDurability) return;
    const existing = Array.from(document.scripts || []).find((script) => {
      try { return new URL(script.src, document.baseURI).pathname.endsWith("/data-durability-runtime.js"); }
      catch (_) { return false; }
    });
    if (existing) return;
    if (document.readyState === "loading") {
      document.write(`<script src="${DATA_DURABILITY_SRC}" data-hera-data-durability="1"><\/script>`);
      return;
    }
    const script = document.createElement("script");
    script.src = DATA_DURABILITY_SRC;
    script.async = false;
    script.setAttribute("data-hera-data-durability", "1");
    document.head.appendChild(script);
  }

  function ensureDataSafetyLayer() {
    if (window.HeraDataSafety) return;
    const existing = Array.from(document.scripts || []).find((script) => {
      try { return new URL(script.src, document.baseURI).pathname.endsWith("/data-safety-layer.js"); }
      catch (_) { return false; }
    });
    if (existing) return;
    if (document.readyState === "loading") {
      document.write(`<script src="${DATA_SAFETY_SRC}" data-hera-data-safety="1"><\/script>`);
      return;
    }
    const script = document.createElement("script");
    script.src = DATA_SAFETY_SRC;
    script.async = false;
    script.setAttribute("data-hera-data-safety", "1");
    document.head.appendChild(script);
  }

  function ensureCriticalWriteSafetyBridge() {
    if (window.HeraCriticalWriteSafetyBridge) return;
    const existing = Array.from(document.scripts || []).find((script) => {
      try { return new URL(script.src, document.baseURI).pathname.endsWith("/critical-write-safety-bridge.js"); }
      catch (_) { return false; }
    });
    if (existing) return;
    if (document.readyState === "loading") {
      document.write(`<script src="${CRITICAL_WRITE_SAFETY_SRC}" data-hera-critical-write-safety="1"><\/script>`);
      return;
    }
    const script = document.createElement("script");
    script.src = CRITICAL_WRITE_SAFETY_SRC;
    script.async = false;
    script.setAttribute("data-hera-critical-write-safety", "1");
    document.head.appendChild(script);
  }

  function ensureSyncBadgePendingFix() {
    if (window.HeraSyncBadgePendingFix) return;
    const existing = Array.from(document.scripts || []).find((script) => {
      try { return new URL(script.src, document.baseURI).pathname.endsWith("/sync-badge-pending-fix.js"); }
      catch (_) { return false; }
    });
    if (existing) return;
    if (document.readyState === "loading") {
      document.write(`<script src="${SYNC_BADGE_PENDING_FIX_SRC}" data-hera-sync-badge-fix="1"><\/script>`);
      return;
    }
    const script = document.createElement("script");
    script.src = SYNC_BADGE_PENDING_FIX_SRC;
    script.async = false;
    script.setAttribute("data-hera-sync-badge-fix", "1");
    document.head.appendChild(script);
  }

  ensureDataDurabilityRuntime();
  ensureDataSafetyLayer();
  ensureCriticalWriteSafetyBridge();
  ensureSyncBadgePendingFix();

  function isNativeAndroid() {
    return Boolean(
      window.Capacitor?.isNativePlatform?.()
      && window.Capacitor?.getPlatform?.() === "android"
    );
  }

  function isAppShellCache(name) {
    return APP_CACHE_PREFIXES.some((prefix) => String(name || "").startsWith(prefix));
  }

  function nativeUpdatePlugin() {
    if (!isNativeAndroid()) return null;
    return window.Capacitor?.Plugins?.HeraAppUpdate
      || window.Capacitor?.registerPlugin?.("HeraAppUpdate")
      || null;
  }

  function dismissedNoticeId() {
    try { return sessionStorage.getItem(UPDATE_NOTICE_DISMISSED_KEY) || ""; }
    catch (_) { return ""; }
  }

  function hideUpdateNotice(noticeId = "") {
    const notice = document.getElementById("hera-update-notice");
    notice?.classList.add("hidden");
    if (!noticeId) return;
    try { sessionStorage.setItem(UPDATE_NOTICE_DISMISSED_KEY, noticeId); }
    catch (_) {}
  }

  function ensureUpdateNotice() {
    let notice = document.getElementById("hera-update-notice");
    if (notice) return notice;
    notice = document.createElement("aside");
    notice.id = "hera-update-notice";
    notice.className = "hera-update-notice hidden";
    notice.setAttribute("role", "status");
    notice.setAttribute("aria-live", "polite");
    notice.innerHTML = `
      <div class="hera-update-notice-copy">
        <strong id="hera-update-notice-title">Nuova versione disponibile</strong>
        <span id="hera-update-notice-detail">Aggiorna per usare le ultime novita.</span>
      </div>
      <div class="hera-update-notice-actions">
        <button id="hera-update-notice-action" type="button">AGGIORNA ORA</button>
        <button id="hera-update-notice-dismiss" type="button" aria-label="Ricordamelo piu tardi">PIU TARDI</button>
      </div>`;
    document.body.appendChild(notice);
    document.getElementById("hera-update-notice-dismiss")?.addEventListener("click", () => {
      hideUpdateNotice(notice.dataset.noticeId || "");
    });
    document.getElementById("hera-update-notice-action")?.addEventListener("click", () => {
      if (notice.dataset.updateKind === "android") void startNativeUpdate();
      else void applyPendingPwaUpdate();
    });
    return notice;
  }

  function showUpdateNotice(kind, noticeId, detail) {
    if (!noticeId || dismissedNoticeId() === noticeId) return;
    const notice = ensureUpdateNotice();
    notice.dataset.updateKind = kind;
    notice.dataset.noticeId = noticeId;
    const title = document.getElementById("hera-update-notice-title");
    const description = document.getElementById("hera-update-notice-detail");
    if (title) title.textContent = kind === "android"
      ? "Nuova versione Android disponibile"
      : "Nuova versione dell'app disponibile";
    if (description) description.textContent = detail;
    notice.classList.remove("hidden");
  }

  function openPlayStore() {
    const storeWindow = window.open(PLAY_STORE_URL, "_system", "noopener,noreferrer");
    if (!storeWindow) window.location.href = PLAY_STORE_URL;
  }

  async function startNativeUpdate() {
    const action = document.getElementById("hera-update-notice-action");
    action?.setAttribute("disabled", "");
    try {
      const plugin = nativeUpdatePlugin();
      const result = plugin?.startUpdate ? await plugin.startUpdate() : null;
      if (!result?.started) openPlayStore();
    } catch (error) {
      console.warn("Aggiornamento Android integrato non disponibile; apro Google Play.", error);
      openPlayStore();
    } finally {
      action?.removeAttribute("disabled");
    }
  }

  async function checkNativeUpdate() {
    if (!isNativeAndroid() || !navigator.onLine) return;
    const now = Date.now();
    if (now - lastNativeCheckAt < NATIVE_CHECK_INTERVAL_MS) return;
    lastNativeCheckAt = now;
    try {
      const details = await nativeUpdatePlugin()?.checkForUpdate?.();
      if (!details?.available) return;
      const version = Number(details.availableVersionCode || 0);
      showUpdateNotice(
        "android",
        `android:${version || "available"}`,
        "Tocca Aggiorna ora per installarla in modo sicuro da Google Play."
      );
    } catch (error) {
      console.warn("Controllo aggiornamento Android non disponibile.", error);
    }
  }

  async function applyPendingPwaUpdate() {
    const action = document.getElementById("hera-update-notice-action");
    action?.setAttribute("disabled", "");
    try {
      await protectDataBeforeReload("pwa-update-notice");
      const registration = pendingPwaRegistration || await navigator.serviceWorker?.getRegistration?.();
      registration?.waiting?.postMessage({ type: "SKIP_WAITING" });
      await clearWebAppCaches();
      reloadWithoutCache();
    } catch (error) {
      console.warn("Applicazione aggiornamento PWA non riuscita.", error);
      action?.removeAttribute("disabled");
    }
  }

  function notifyPwaUpdate(registration, versionHint = "") {
    pendingPwaRegistration = registration || pendingPwaRegistration;
    const workerUrl = registration?.waiting?.scriptURL || versionHint || "ready";
    showUpdateNotice(
      "pwa",
      `pwa:${workerUrl}`,
      "Tocca Aggiorna ora: accesso e dati salvati resteranno invariati."
    );
  }

  function watchPwaRegistration(registration) {
    if (!registration || watchedPwaRegistrations.has(registration)) return;
    watchedPwaRegistrations.add(registration);
    if (registration.waiting) notifyPwaUpdate(registration);
    registration.addEventListener("updatefound", () => {
      const worker = registration.installing;
      worker?.addEventListener("statechange", () => {
        if (worker.state === "installed" && navigator.serviceWorker.controller) {
          notifyPwaUpdate(registration, worker.scriptURL);
        }
      });
    });
  }

  async function protectDataBeforeReload(reason) {
    try {
      if (window.HeraDataDurability?.prepareForUpdate) {
        await window.HeraDataDurability.prepareForUpdate(reason || "pwa-refresh");
      }
    } catch (error) {
      console.warn("Backup dati pre-aggiornamento non completato; evito cancellazioni dei dati locali.", error);
    }
  }

  async function clearWebAppCaches() {
    if (!("caches" in window)) return;
    try {
      const cacheNames = await caches.keys();
      const appShellCaches = cacheNames.filter(isAppShellCache);
      await Promise.all(appShellCaches.map((name) => caches.delete(name)));
    } catch (error) {
      console.warn("Pulizia cache web non riuscita; proseguo con l'aggiornamento.", error);
    }
  }

  function reloadWithoutCache() {
    const refreshUrl = new URL(window.location.href);
    refreshUrl.searchParams.set("appRefresh", String(Date.now()));
    window.location.replace(refreshUrl.toString());
  }

  async function requestPwaUpdate({ reload = false } = {}) {
    try {
      // IMPORTANTE: Refresh elimina solo la cache dell'app shell.
      // Non deve mai eseguire signOut, localStorage.clear(), sessionStorage.clear()
      // o cancellare IndexedDB: Firebase Auth, dati offline e preferenze restano intatti.
      if (reload) {
        await protectDataBeforeReload("manual-pwa-update");
        await clearWebAppCaches();
      }

      if ("serviceWorker" in navigator) {
        const registrations = await navigator.serviceWorker.getRegistrations();
        await Promise.all(registrations.map(async (registration) => {
          await registration.update().catch(() => null);
          if (reload) registration.waiting?.postMessage({ type: "SKIP_WAITING" });
        }));
      }

      if (reload) reloadWithoutCache();
      return true;
    } catch (error) {
      console.warn("Controllo aggiornamento web non riuscito.", error);
      if (reload) {
        await protectDataBeforeReload("manual-pwa-update-fallback");
        reloadWithoutCache();
      }
      return false;
    }
  }

  async function updateApplication(button) {
    if (!button || button.dataset.updateBusy === "1") return;
    button.dataset.updateBusy = "1";
    button.disabled = true;
    button.setAttribute("aria-busy", "true");

    try {
      const updated = await requestPwaUpdate({ reload: true });
      if (!updated && document.visibilityState === "visible") reloadWithoutCache();
    } finally {
      if (document.visibilityState === "visible") {
        button.disabled = false;
        button.removeAttribute("aria-busy");
        delete button.dataset.updateBusy;
      }
    }
  }

  function installAutomaticPwaUpdate() {
    if (isNativeAndroid()) {
      const checkNative = () => void checkNativeUpdate();
      window.addEventListener("online", checkNative);
      window.addEventListener("pageshow", checkNative);
      document.addEventListener("visibilitychange", () => {
        if (document.visibilityState === "visible") checkNative();
      });
      window.setTimeout(checkNative, 1800);
      return;
    }
    if (!("serviceWorker" in navigator)) return;

    // Gli aggiornamenti automatici vengono solo scaricati. Non ricarichiamo più
    // la pagina da soli: evitiamo di interrompere moduli o dati in compilazione.
    navigator.serviceWorker.addEventListener("controllerchange", () => {
      window.dispatchEvent(new CustomEvent("hera:update-controller-changed"));
    });
    navigator.serviceWorker.addEventListener("message", (event) => {
      if (hadPwaControllerAtStartup && event.data?.type === "HERA_SW_UPDATE_READY") {
        notifyPwaUpdate(null, String(event.data.version || "ready"));
      }
    });
    navigator.serviceWorker.ready.then((registration) => {
      watchPwaRegistration(registration);
      if (registration.waiting) notifyPwaUpdate(registration);
    }).catch(() => null);

    const check = async () => {
      const registration = await navigator.serviceWorker.getRegistration().catch(() => null);
      watchPwaRegistration(registration);
      await requestPwaUpdate();
      if (registration?.waiting) notifyPwaUpdate(registration);
    };
    window.addEventListener("online", check);
    window.addEventListener("pageshow", check);
    document.addEventListener("visibilitychange", () => {
      if (document.visibilityState === "visible") check();
    });
    window.setTimeout(check, 1500);
  }

  function ensureStyles() {
    if (document.getElementById("pwa-update-compact-style")) return;
    const style = document.createElement("style");
    style.id = "pwa-update-compact-style";
    style.textContent = `
      #update-app-btn,
      #auth-update-pwa-btn {
        width: 30px;
        height: 30px;
        min-width: 30px;
        min-height: 30px;
        padding: 0;
        display: inline-flex;
        align-items: center;
        justify-content: center;
        border: 1px solid #cfe0da;
        border-radius: 9px;
        background: linear-gradient(180deg, #fff 0%, #f3f8f6 100%);
        color: #184c3d;
        box-shadow: 0 4px 10px rgba(15, 23, 42, 0.08);
        font-size: 0.95rem;
        font-weight: 800;
        line-height: 1;
        cursor: pointer;
        flex: 0 0 auto;
      }
      #update-app-btn span:last-child {
        position: absolute;
        width: 1px;
        height: 1px;
        padding: 0;
        margin: -1px;
        overflow: hidden;
        clip: rect(0, 0, 0, 0);
        white-space: nowrap;
        border: 0;
      }
      #auth-update-pwa-wrap {
        display: flex;
        justify-content: center;
        align-items: center;
        min-height: 30px;
        margin-top: 8px;
      }
      #auth-update-pwa-btn {
        width: auto;
        min-width: 30px;
        padding: 0 8px;
        gap: 5px;
        font-size: 0.75rem;
        font-weight: 700;
        white-space: nowrap;
      }
      #auth-update-pwa-btn .pwa-update-icon {
        font-size: 0.95rem;
        line-height: 1;
      }
      #update-app-btn:focus-visible,
      #auth-update-pwa-btn:focus-visible {
        outline: 3px solid rgba(37, 99, 235, 0.24);
        outline-offset: 2px;
      }
      .hera-update-notice {
        position: fixed;
        left: max(12px, env(safe-area-inset-left));
        right: max(12px, env(safe-area-inset-right));
        bottom: max(12px, env(safe-area-inset-bottom));
        z-index: 30000;
        width: min(720px, calc(100% - 24px));
        margin: 0 auto;
        padding: 14px;
        display: flex;
        align-items: center;
        justify-content: space-between;
        gap: 14px;
        border: 1px solid #86efac;
        border-radius: 16px;
        background: #f0fdf4;
        color: #14532d;
        box-shadow: 0 16px 40px rgba(15, 23, 42, 0.24);
      }
      .hera-update-notice.hidden { display: none; }
      .hera-update-notice-copy { display: grid; gap: 3px; }
      .hera-update-notice-copy span { font-size: 0.86rem; }
      .hera-update-notice-actions { display: flex; gap: 8px; flex: 0 0 auto; }
      .hera-update-notice-actions button {
        min-height: 38px;
        padding: 0 12px;
        border: 0;
        border-radius: 10px;
        background: #166534;
        color: #fff;
        font-weight: 800;
        cursor: pointer;
      }
      #hera-update-notice-dismiss { background: #dcfce7; color: #14532d; }
      @media (max-width: 420px) {
        #auth-update-pwa-btn {
          padding: 0 7px;
          font-size: 0.72rem;
        }
        .hera-update-notice { align-items: stretch; flex-direction: column; }
        .hera-update-notice-actions { width: 100%; }
        .hera-update-notice-actions button { flex: 1; }
      }
    `;
    document.head.appendChild(style);
  }

  function bindUpdateButton(button) {
    if (!button || button.dataset.pwaUpdateBound === "1") return;
    button.dataset.pwaUpdateBound = "1";
    button.addEventListener("click", (event) => {
      event.preventDefault();
      event.stopImmediatePropagation();
      if (isNativeAndroid()) void startNativeUpdate();
      else void updateApplication(button);
    });
  }

  function installHomeButton() {
    const userButton = document.getElementById("user-toggle-btn");
    if (!userButton) return;

    let button = document.getElementById("update-app-btn");
    if (!button) {
      button = document.createElement("button");
      button.id = "update-app-btn";
      button.className = "update-app-btn";
      button.type = "button";
      button.innerHTML = '<span aria-hidden="true">↻</span><span>Aggiorna app</span>';
      userButton.insertAdjacentElement("afterend", button);
    }

    button.title = "Refresh: aggiorna l’app mantenendo accesso e dati";
    button.setAttribute("aria-label", "Refresh: aggiorna l’app mantenendo accesso e dati");
    bindUpdateButton(button);
  }

  function installLoginButton() {
    if (isNativeAndroid()) return;
    const card = document.querySelector("#auth-gate .auth-gate-card");
    if (!card || document.getElementById("auth-update-pwa-btn")) return;

    const wrap = document.createElement("div");
    wrap.id = "auth-update-pwa-wrap";

    const button = document.createElement("button");
    button.id = "auth-update-pwa-btn";
    button.type = "button";
    button.title = "Aggiorna la versione PWA mantenendo accesso e dati";
    button.setAttribute("aria-label", "Aggiorna la versione PWA mantenendo accesso e dati");
    button.innerHTML = '<span class="pwa-update-icon" aria-hidden="true">↻</span><span>Aggiorna PWA</span>';

    wrap.appendChild(button);
    card.appendChild(wrap);
    bindUpdateButton(button);
  }

  function install() {
    ensureStyles();
    installHomeButton();
    installLoginButton();
  }

  installAutomaticPwaUpdate();
  if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", install, { once: true });
  } else {
    install();
  }
})();
