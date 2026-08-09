window.firebaseConfig = {
  apiKey: "AIzaSyBHZG9D1H5YOT9QUzG-cdSdftlreDJNa_k",
  authDomain: "hera-app-6cd2b.firebaseapp.com",
  projectId: "hera-app-6cd2b",
  storageBucket: "hera-app-6cd2b.firebasestorage.app",
  messagingSenderId: "645390631375",
  appId: "1:645390631375:web:df3659a23812560e4012ba"
};

const HERA_FIRESTORE_OPERATION_DIAGNOSTICS_SRC = "firestore-operation-diagnostics.js?v=20260806a";
const HERA_FIRESTORE_DIAGNOSTICS_V4_CLEANUP_SRC = "firestore-diagnostics-v4-session-cleanup.js?v=20260805a";
const HERA_FIRESTORE_DIAGNOSTICS_V4_SRC = "firestore-diagnostics-dashboard-v4.js?v=20260805b";
const HERA_FIRESTORE_SAFE_OPTIMIZER_SRC = "firestore-safe-optimizer.js?v=20260805b";
const HERA_FIRESTORE_INFLIGHT_COALESCER_SRC = "firestore-inflight-read-coalescer.js?v=20260805a";
const HERA_FIRESTORE_DIAGNOSTICS_OPTIMIZER_SRC = "firestore-diagnostics-optimizer-extension.js?v=20260804b";
const HERA_OFFLINE_SHARED_STATIC_VIEWS_BRIDGE_SRC = "offline-shared-static-views-bridge.js?v=20260808a";
const HERA_SHARED_STATIC_VIEWS_SRC = "shared-static-views.js?v=20260804a";
const HERA_ACTIVE_COMMESSE_FIRST_BOOT_GUARD_SRC = "active-commesse-first-boot-guard.js?v=20260806a";
const HERA_OPERATIONAL_OFFLINE_CACHE_SRC = "offline-operational-cache.js?v=20260808b";
const HERA_FIRESTORE_STARTUP_COST_OPTIMIZER_SRC = "firestore-startup-cost-optimizer.js?v=20260805a";
const HERA_SHARED_STATIC_VIEWS_UI_SRC = "shared-static-views-ui.js?v=20260804b";
const HERA_PERSISTENT_OFFLINE_AUTH_SRC = "persistent-offline-auth.js?v=20260809a";
const HERA_ACCESS_REQUEST_LOGIN_SRC = "access-request-login.js?v=20260809c";
const HERA_ACCESS_REQUEST_ADMIN_SRC = "access-request-admin.js?v=20260809b";

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
    loadOnce(
      HERA_SHARED_STATIC_VIEWS_UI_SRC,
      "hera-shared-static-views-ui",
      () => window.__heraSharedStaticViewsUiInstalled
    );
    loadOnce("notification-session-enhancements.js?v=20260727b", "hera-notification-session", () => false);
    loadOnce("update-app-feature.js?v=20260727a", "hera-app-update", () => false);
    loadOnce("google-sheet-two-way-sync.js?v=20260729b", "hera-google-sheet-sync", () => false);
    loadOnce("personnel-training-manager.js?v=20260803a", "personnel-training-manager", () => false);
    loadOnce("offline-first-runtime.js?v=20260807a", "hera-offline-first-runtime", () => window.HeraOfflineFirstRuntime?.installed === true);
    loadOnce(
      "operator-account-admin.js?v=20260807a",
      "hera-operator-account-admin",
      () => window.__heraOperatorAccountAdminInstalled === true
    );
    loadOnce(
      "personnel-app-access.js?v=20260807a",
      "hera-personnel-app-access",
      () => window.HeraPersonnelAppAccess?.installed === true
    );
    loadOnce(
      HERA_ACCESS_REQUEST_ADMIN_SRC,
      "hera-access-request-admin",
      () => window.HeraAccessRequestAdmin?.installed === true
    );
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
  document.write(`<script src="${HERA_FIRESTORE_OPERATION_DIAGNOSTICS_SRC}" data-firestore-operation-diagnostics="1"><\/script>`);
  document.write(`<script src="${HERA_FIRESTORE_SAFE_OPTIMIZER_SRC}" data-firestore-safe-optimizer="1"><\/script>`);
  document.write(`<script src="${HERA_FIRESTORE_INFLIGHT_COALESCER_SRC}" data-hera-firestore-inflight-coalescer="1"><\/script>`);
  document.write(`<script src="${HERA_OFFLINE_SHARED_STATIC_VIEWS_BRIDGE_SRC}" data-hera-offline-shared-static-views-bridge="1"><\/script>`);
  document.write(`<script src="${HERA_SHARED_STATIC_VIEWS_SRC}" data-hera-shared-static-views="1"><\/script>`);
  document.write(`<script src="${HERA_ACTIVE_COMMESSE_FIRST_BOOT_GUARD_SRC}" data-active-commesse-first-boot-guard="1"><\/script>`);
  document.write(`<script src="${HERA_OPERATIONAL_OFFLINE_CACHE_SRC}" data-hera-operational-offline-cache="1"><\/script>`);
  document.write(`<script src="${HERA_FIRESTORE_STARTUP_COST_OPTIMIZER_SRC}" data-firestore-startup-cost-optimizer="1"><\/script>`);
  document.write(`<script src="${HERA_PERSISTENT_OFFLINE_AUTH_SRC}" data-hera-persistent-offline-auth="1"><\/script>`);
  document.write(`<script src="${HERA_ACCESS_REQUEST_LOGIN_SRC}" data-hera-access-request-login="1"><\/script>`);
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
      HERA_OFFLINE_SHARED_STATIC_VIEWS_BRIDGE_SRC,
      "hera-offline-shared-static-views-bridge",
      () => window.HeraOfflineSharedStaticViewsBridge?.installed === true,
      () => loadOnce(
        HERA_SHARED_STATIC_VIEWS_SRC,
        "hera-shared-static-views",
        () => window.HeraSharedStaticViews?.installed
      )
    );
    loadOnce(
      HERA_ACTIVE_COMMESSE_FIRST_BOOT_GUARD_SRC,
      "active-commesse-first-boot-guard",
      () => window.HeraActiveCommesseFirstBootGuard?.installed
    );
    loadOnce(
      HERA_OPERATIONAL_OFFLINE_CACHE_SRC,
      "hera-operational-offline-cache",
      () => window.HeraOperationalOfflineCache?.installed === true
    );
    loadOnce(
      HERA_FIRESTORE_STARTUP_COST_OPTIMIZER_SRC,
      "firestore-startup-cost-optimizer",
      () => window.HeraFirestoreStartupCostOptimizer?.installed
    );
    loadOnce(
      HERA_PERSISTENT_OFFLINE_AUTH_SRC,
      "hera-persistent-offline-auth",
      () => window.HeraPersistentOfflineAuth?.installed === true
    );
    loadOnce(
      HERA_ACCESS_REQUEST_LOGIN_SRC,
      "hera-access-request-login",
      () => window.HeraAccessRequestLogin?.installed === true
    );
  };

  loadCriticalModules();
  scheduleDeferredStartupModules();
}