"use strict";
(function exposeVargaPureUtils(global) {
  const api = {};
  function toggleOperatorPositionsVisibility() {
    return;
  }
  api.toggleOperatorPositionsVisibility = toggleOperatorPositionsVisibility;
  function createEmptyCorsoState() {
    return { possiede: false };
  }
  api.createEmptyCorsoState = createEmptyCorsoState;
  function getHomeWorklimateButtonLabel() {
    return "worklimate";
  }
  api.getHomeWorklimateButtonLabel = getHomeWorklimateButtonLabel;
  function normalizeEmail(email) {
    return String(email || "").trim().toLowerCase();
  }
  api.normalizeEmail = normalizeEmail;
  Object.assign(global, api);
  global.VargaPureUtils = Object.freeze({ ...(global.VargaPureUtils || {}), ...api });
})(window);

// Phase 27: pure utilities extracted from app.js
(function exposeMoreVargaPureUtils(global) {
  const api = {};
  function sanitizePhone(value) {
    return String(value || "").replace(/[^0-9+]/g, "");
  }
  api.sanitizePhone = sanitizePhone;
  function normalizeMezzoNId(value) {
    return String(value || "").trim().toLowerCase();
  }
  api.normalizeMezzoNId = normalizeMezzoNId;
  function formatPhoneHref(phone = "") {
    return String(phone || "").replace(/[^+\d]/g, "");
  }
  api.formatPhoneHref = formatPhoneHref;
  function getMinutesSinceMidnight(date) {
    return date.getHours() * 60 + date.getMinutes();
  }
  api.getMinutesSinceMidnight = getMinutesSinceMidnight;
  Object.assign(global, api);
  global.VargaPureUtils = Object.freeze({ ...(global.VargaPureUtils || {}), ...api });
})(window);

// Phase 28: pure utilities extracted from app.js
(function exposePhase28PureUtils(global) {
  const api = {};
  function encodeHoursLockPart(value) {
    return encodeURIComponent(String(value || "").trim());
  }
  api.encodeHoursLockPart = encodeHoursLockPart;
  function getCentralDriveNotConfiguredMessage() {
    return "Cloud amministratore non configurato";
  }
  api.getCentralDriveNotConfiguredMessage = getCentralDriveNotConfiguredMessage;
  function normalizeAtexSearchValue(value) {
    return String(value || "").trim().toLocaleUpperCase("it-IT");
  }
  api.normalizeAtexSearchValue = normalizeAtexSearchValue;
  function setUsedActionButtonState(btn, used) {
    btn.disabled = used;
    btn.classList.toggle("is-used", used);
  }
  api.setUsedActionButtonState = setUsedActionButtonState;
  Object.assign(global, api);
  global.VargaPureUtils = Object.freeze({ ...(global.VargaPureUtils || {}), ...api });
})(window);

// Performance runtime: piccolo, senza dipendenze e senza letture/scritture remote.
(function loadPerformanceRuntime() {
  if (document.querySelector("script[data-hera-performance-runtime]")) return;
  const script = document.createElement("script");
  script.src = "./performance-runtime.js?v=20260903a";
  script.async = true;
  script.setAttribute("data-hera-performance-runtime", "1");
  document.head.appendChild(script);
})();

// Loader non critico. I moduli Verde Bologna/Levato non vengono più parsati
// durante il bootstrap: si caricano al primo ingresso pertinente. Un fallback
// idle mantiene compatibilità con i flussi storici che possono invocarli senza click.
(function installDeferredOptionalRuntimes() {
  const state = {
    verdeRequested: false,
    licenseRequested: false,
    loadedScripts: new Set()
  };

  function addScript(src, dataName) {
    if (document.querySelector(`script[data-${dataName}]`)) return Promise.resolve(true);
    const existing = Array.from(document.scripts || []).find((node) => {
      try { return new URL(node.src, document.baseURI).pathname.endsWith(new URL(src, document.baseURI).pathname); }
      catch (_) { return false; }
    });
    if (existing) return Promise.resolve(true);

    return new Promise((resolve, reject) => {
      const script = document.createElement("script");
      script.src = src;
      script.async = false;
      script.setAttribute(`data-${dataName}`, "1");
      script.addEventListener("load", () => {
        state.loadedScripts.add(src);
        resolve(true);
      }, { once: true });
      script.addEventListener("error", () => reject(new Error(`Caricamento ${src} non riuscito`)), { once: true });
      document.head.appendChild(script);
    });
  }

  async function loadDataLicenseView() {
    if (state.licenseRequested) return;
    state.licenseRequested = true;
    try {
      await addScript("./data-license-view.js?v=20260831-license1", "hera-data-license-view");
    } catch (error) {
      state.licenseRequested = false;
      console.warn("Vista licenze non caricata:", error);
    }
  }

  async function loadVerdeRuntimes() {
    if (state.verdeRequested) return;
    state.verdeRequested = true;
    try {
      await addScript("./verde-bologna.js?v=20260901-cobo-sfalcio1", "hera-verde-bologna");
      await addScript("./verde-bologna-operativo.js?v=20260901-catasto-open2", "hera-verde-bologna-operativo");
      await addScript("./verde-bologna-parchi-mobile.js?v=20260901-cobo-sfalcio1", "hera-verde-bologna-parchi-mobile");
      await addScript("./verde-levato.js?v=20260902-verde-levato2", "hera-verde-levato");
    } catch (error) {
      state.verdeRequested = false;
      console.warn("Moduli Verde differiti non caricati:", error);
    }
  }

  function isVerdeEntry(target) {
    const button = target?.closest?.("button, [role='button'], a");
    if (!button) return false;
    const id = String(button.id || "").toLowerCase();
    const action = String(button.dataset?.action || "").toLowerCase();
    const text = String(button.textContent || "").toLowerCase();
    return id.includes("green")
      || id.includes("verde")
      || id.includes("tree-search")
      || action.includes("green")
      || action.includes("verde")
      || /verde bologna|verde levato|catasto alberi|aree verdi/.test(text);
  }

  document.addEventListener("pointerdown", (event) => {
    if (isVerdeEntry(event.target)) void loadVerdeRuntimes();
  }, { capture: true, passive: true });

  document.addEventListener("focusin", (event) => {
    if (isVerdeEntry(event.target)) void loadVerdeRuntimes();
  }, true);

  const scheduleIdle = (job, timeout, fallbackDelay) => {
    if ("requestIdleCallback" in window) window.requestIdleCallback(() => void job(), { timeout });
    else window.setTimeout(() => void job(), fallbackDelay);
  };

  const afterLoad = () => {
    scheduleIdle(loadDataLicenseView, 4500, 2500);
    scheduleIdle(loadVerdeRuntimes, 9000, 6500);
  };
  if (document.readyState === "complete") afterLoad();
  else window.addEventListener("load", afterLoad, { once: true });

  window.HeraDeferredOptionalRuntimes = {
    installed: true,
    version: "1.0.0",
    loadVerde: loadVerdeRuntimes,
    loadDataLicense: loadDataLicenseView,
    getState: () => ({
      verdeRequested: state.verdeRequested,
      licenseRequested: state.licenseRequested,
      loadedScripts: Array.from(state.loadedScripts)
    })
  };
})();
