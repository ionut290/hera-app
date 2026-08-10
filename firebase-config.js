window.firebaseConfig = {
  apiKey: "AIzaSyBHZG9D1H5YOT9QUzG-cdSdftlreDJNa_k",
  authDomain: "hera-app-6cd2b.firebaseapp.com",
  projectId: "hera-app-6cd2b",
  storageBucket: "hera-app-6cd2b.firebasestorage.app",
  messagingSenderId: "645390631375",
  appId: "1:645390631375:web:df3659a23812560e4012ba"
};

// Moduli necessari prima di app.js: protezioni letture, viste condivise e guardia
// del primo caricamento delle commesse. Tutto il resto viene caricato dopo che
// la pagina è pronta, così l'avvio non resta bloccato da strumenti diagnostici
// o funzioni amministrative non necessarie alla home.
const HERA_FIRESTORE_OPERATION_DIAGNOSTICS_SRC = "firestore-operation-diagnostics.js?v=20260806a";
const HERA_FIRESTORE_DIAGNOSTICS_V4_CLEANUP_SRC = "firestore-diagnostics-v4-session-cleanup.js?v=20260805a";
const HERA_FIRESTORE_DIAGNOSTICS_V4_SRC = "firestore-diagnostics-dashboard-v4.js?v=20260805b";
const HERA_FIRESTORE_SAFE_OPTIMIZER_SRC = "firestore-safe-optimizer.js?v=20260805b";
const HERA_FIRESTORE_INFLIGHT_COALESCER_SRC = "firestore-inflight-read-coalescer.js?v=20260805a";
const HERA_FIRESTORE_DIAGNOSTICS_OPTIMIZER_SRC = "firestore-diagnostics-optimizer-extension.js?v=20260804b";
const HERA_SHARED_STATIC_VIEWS_SRC = "shared-static-views.js?v=20260804a";
const HERA_ACTIVE_COMMESSE_FIRST_BOOT_GUARD_SRC = "active-commesse-first-boot-guard.js?v=20260806a";
const HERA_FIRESTORE_STARTUP_COST_OPTIMIZER_SRC = "firestore-startup-cost-optimizer.js?v=20260805a";
const HERA_SHARED_STATIC_VIEWS_UI_SRC = "shared-static-views-ui.js?v=20260804b";
const HERA_ADMIN_PASSWORD_MANAGER_SRC = "admin-password-manager.js?v=20260810c";
const HERA_USER_MANAGEMENT_SEARCH_FIX_SRC = "user-management-search-input-fix.js?v=20260810c";

function loadOnce(src, dataName, ready, onLoad) {
  if (ready?.()) {
    onLoad?.();
    return;
  }

  const selector = `script[data-${dataName}="true"], script[data-${dataName}="1"]`;
  const existing = document.querySelector(selector);
  if (existing) {
    if (onLoad) {
      if (existing.dataset.loaded === "1") onLoad();
      else existing.addEventListener("load", onLoad, { once: true });
    }
    return;
  }

  const script = document.createElement("script");
  script.src = src;
  script.async = true;
  script.setAttribute(`data-${dataName}`, "true");
  script.addEventListener("load", () => {
    script.dataset.loaded = "1";
    onLoad?.();
  }, { once: true });
  script.addEventListener("error", () => {
    console.warn("Modulo opzionale non caricato:", src);
  }, { once: true });
  document.head.appendChild(script);
}

function scheduleDeferredStartupModules() {
  if (window.__heraDeferredStartupModulesScheduled) return;
  window.__heraDeferredStartupModulesScheduled = true;

  const loadDeferredModules = () => {
    // V4 non viene più caricato automaticamente.
    // V3 resta l'unico osservatore Firestore durante l'uso normale dell'app:
    // in questo modo il report attribuisce le letture al vero chiamante invece
    // di vedere il wrapper V4 come sorgente, senza cambiare query o listener reali.

    loadOnce(
      HERA_SHARED_STATIC_VIEWS_UI_SRC,
      "hera-shared-static-views-ui",
      () => window.__heraSharedStaticViewsUiInstalled
    );
    loadOnce(
      HERA_ADMIN_PASSWORD_MANAGER_SRC,
      "hera-admin-password-manager",
      () => window.HeraAdminPasswordManager?.installed
    );
    loadOnce(
      HERA_USER_MANAGEMENT_SEARCH_FIX_SRC,
      "hera-user-management-search-fix",
      () => window.HeraUserManagementSearchFix?.version === "2.0.0"
    );
    loadOnce("notification-session-enhancements.js?v=20260727b", "hera-notification-session", () => false);
    loadOnce("update-app-feature.js?v=20260727a", "hera-app-update", () => false);
    loadOnce("google-sheet-two-way-sync.js?v=20260729b", "hera-google-sheet-sync", () => false);
    loadOnce("personnel-training-manager.js?v=20260803a", "personnel-training-manager", () => false);
  };

  const scheduleIdle = () => {
    if (typeof window.requestIdleCallback === "function") {
      window.requestIdleCallback(loadDeferredModules, { timeout: 2500 });
    } else {
      window.setTimeout(loadDeferredModules, 700);
    }
  };

  if (document.readyState === "complete") scheduleIdle();
  else window.addEventListener("load", scheduleIdle, { once: true });
}

if (document.readyState === "loading") {
  // Mantiene l'ordine necessario prima di app.js, ma elimina dal percorso
  // critico diagnostica grafica, UI accessorie, Google Sheet e formazione.
  document.write(`<script src="${HERA_FIRESTORE_OPERATION_DIAGNOSTICS_SRC}" data-firestore-operation-diagnostics="1"><\/script>`);
  document.write(`<script src="${HERA_FIRESTORE_SAFE_OPTIMIZER_SRC}" data-firestore-safe-optimizer="1"><\/script>`);
  document.write(`<script src="${HERA_FIRESTORE_INFLIGHT_COALESCER_SRC}" data-hera-firestore-inflight-coalescer="1"><\/script>`);
  document.write(`<script src="${HERA_SHARED_STATIC_VIEWS_SRC}" data-hera-shared-static-views="1"><\/script>`);
  document.write(`<script src="${HERA_ACTIVE_COMMESSE_FIRST_BOOT_GUARD_SRC}" data-active-commesse-first-boot-guard="1"><\/script>`);
  document.write(`<script src="${HERA_FIRESTORE_STARTUP_COST_OPTIMIZER_SRC}" data-firestore-startup-cost-optimizer="1"><\/script>`);
  scheduleDeferredStartupModules();
} else {
  const loadCriticalModules = () => {
    loadOnce(
      HERA_FIRESTORE_OPERATION_DIAGNOSTICS_SRC,
      "firestore-operation-diagnostics",
      () => window.__vargaFsDiagV3,
      () => loadOnce(
        HERA_FIRESTORE_SAFE_OPTIMIZER_SRC,
        "firestore-safe-optimizer",
        () => window.VargaFirestoreSafeOptimizer?.installed
      )
    );

    loadOnce(
      HERA_FIRESTORE_INFLIGHT_COALESCER_SRC,
      "hera-firestore-inflight-coalescer",
      () => window.HeraFirestoreInflightReadCoalescer?.installed
    );
    loadOnce(
      HERA_SHARED_STATIC_VIEWS_SRC,
      "hera-shared-static-views",
      () => window.HeraSharedStaticViews?.installed
    );
    loadOnce(
      HERA_ACTIVE_COMMESSE_FIRST_BOOT_GUARD_SRC,
      "active-commesse-first-boot-guard",
      () => window.HeraActiveCommesseFirstBootGuard?.installed
    );
    loadOnce(
      HERA_FIRESTORE_STARTUP_COST_OPTIMIZER_SRC,
      "firestore-startup-cost-optimizer",
      () => window.HeraFirestoreStartupCostOptimizer?.installed
    );
  };

  loadCriticalModules();
  scheduleDeferredStartupModules();
}
