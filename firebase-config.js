window.firebaseConfig = {
  apiKey: "AIzaSyBHZG9D1H5YOT9QUzG-cdSdftlreDJNa_k",
  authDomain: "hera-app-6cd2b.firebaseapp.com",
  projectId: "hera-app-6cd2b",
  storageBucket: "hera-app-6cd2b.firebasestorage.app",
  messagingSenderId: "645390631375",
  appId: "1:645390631375:web:df3659a23812560e4012ba"
};

// Carica diagnostica, protezioni Firestore e bridge Android prima di app.js.
// La diagnostica viene installata per prima: l'ottimizzatore la richiama solo
// quando apre un listener fisico, così il report non conta due volte gli
// abbonati logici che condividono la stessa query.
const HERA_FIRESTORE_OPERATION_DIAGNOSTICS_SRC = "firestore-operation-diagnostics.js?v=20260805a";
const HERA_FIRESTORE_DIAGNOSTICS_V4_CLEANUP_SRC = "firestore-diagnostics-v4-session-cleanup.js?v=20260805a";
const HERA_FIRESTORE_DIAGNOSTICS_V4_SRC = "firestore-diagnostics-dashboard-v4.js?v=20260805b";
const HERA_FIRESTORE_SAFE_OPTIMIZER_SRC = "firestore-safe-optimizer.js?v=20260805b";
const HERA_NATIVE_RUNTIME_SRC = "native-android-runtime.js?v=20260803-whatsapp-early2";
const HERA_FIRESTORE_INFLIGHT_COALESCER_SRC = "firestore-inflight-read-coalescer.js?v=20260805a";
const HERA_FIRESTORE_DIAGNOSTICS_OPTIMIZER_SRC = "firestore-diagnostics-optimizer-extension.js?v=20260804b";
const HERA_SHARED_STATIC_VIEWS_SRC = "shared-static-views.js?v=20260804a";
const HERA_ACTIVE_COMMESSE_FIRST_BOOT_GUARD_SRC = "active-commesse-first-boot-guard.js?v=20260806a";
const HERA_FIRESTORE_STARTUP_COST_OPTIMIZER_SRC = "firestore-startup-cost-optimizer.js?v=20260805a";
const HERA_SHARED_STATIC_VIEWS_UI_SRC = "shared-static-views-ui.js?v=20260804b";
const HERA_FATTO_EMBEDDED_OPTIMIZER_SRC = "fatto-visual-evidence-embedded.js?v=20260806a";

if (document.readyState === "loading") {
  document.write(`<script src="${HERA_FIRESTORE_OPERATION_DIAGNOSTICS_SRC}" data-firestore-operation-diagnostics="1"><\/script>`);
  document.write(`<script src="${HERA_FIRESTORE_DIAGNOSTICS_V4_CLEANUP_SRC}" data-firestore-diagnostics-v4-cleanup="1"><\/script>`);
  document.write(`<script src="${HERA_FIRESTORE_DIAGNOSTICS_V4_SRC}" data-firestore-diagnostics-v4="1"><\/script>`);
  document.write(`<script src="${HERA_FIRESTORE_SAFE_OPTIMIZER_SRC}" data-firestore-safe-optimizer="1"><\/script>`);
  document.write(`<script src="${HERA_FIRESTORE_INFLIGHT_COALESCER_SRC}"><\/script>`);
  document.write(`<script src="${HERA_FIRESTORE_DIAGNOSTICS_OPTIMIZER_SRC}"><\/script>`);
  document.write(`<script src="${HERA_SHARED_STATIC_VIEWS_SRC}"><\/script>`);
  document.write(`<script src="${HERA_ACTIVE_COMMESSE_FIRST_BOOT_GUARD_SRC}" data-active-commesse-first-boot-guard="1"><\/script>`);
  document.write(`<script src="${HERA_FIRESTORE_STARTUP_COST_OPTIMIZER_SRC}" data-firestore-startup-cost-optimizer="1"><\/script>`);
  document.write(`<script src="${HERA_SHARED_STATIC_VIEWS_UI_SRC}"><\/script>`);
  document.write(`<script src="${HERA_NATIVE_RUNTIME_SRC}"><\/script>`);
  document.write(`<script src="${HERA_FATTO_EMBEDDED_OPTIMIZER_SRC}" data-fatto-embedded-optimizer="1"><\/script>`);
  document.write('<script src="notification-session-enhancements.js?v=20260727b"><\/script>');
  document.write('<script src="update-app-feature.js?v=20260727a"><\/script>');
  document.write('<script src="google-sheet-two-way-sync.js?v=20260729b"><\/script>');
  document.write('<script src="personnel-training-manager.js?v=20260803a"><\/script>');
} else {
  function loadOnce(src, dataName, ready, onLoad) {
    if (ready?.()) {
      onLoad?.();
      return;
    }
    const existing = document.querySelector(`script[data-${dataName}="true"], script[data-${dataName}="1"]`);
    if (existing) {
      if (onLoad) existing.addEventListener("load", onLoad, { once: true });
      return;
    }
    const script = document.createElement("script");
    script.src = src;
    script.setAttribute(`data-${dataName}`, "true");
    if (onLoad) script.addEventListener("load", onLoad, { once: true });
    document.head.appendChild(script);
  }

  const loadSafeOptimizer = () => loadOnce(
    HERA_FIRESTORE_SAFE_OPTIMIZER_SRC,
    "firestore-safe-optimizer",
    () => window.VargaFirestoreSafeOptimizer?.installed
  );

  const loadDiagnosticsV4 = () => loadOnce(
    HERA_FIRESTORE_DIAGNOSTICS_V4_SRC,
    "firestore-diagnostics-v4",
    () => window.VargaFirestoreDiagnosticsV4?.installed,
    loadSafeOptimizer
  );

  const loadDiagnosticsV4Cleanup = () => loadOnce(
    HERA_FIRESTORE_DIAGNOSTICS_V4_CLEANUP_SRC,
    "firestore-diagnostics-v4-cleanup",
    () => false,
    loadDiagnosticsV4
  );

  loadOnce(
    HERA_FIRESTORE_OPERATION_DIAGNOSTICS_SRC,
    "firestore-operation-diagnostics",
    () => window.__vargaFsDiagV3,
    loadDiagnosticsV4Cleanup
  );
  window.setTimeout(loadDiagnosticsV4Cleanup, 75);
  window.setTimeout(loadDiagnosticsV4, 150);
  window.setTimeout(loadSafeOptimizer, 300);

  loadOnce(HERA_FIRESTORE_INFLIGHT_COALESCER_SRC, "hera-firestore-inflight-coalescer", () => window.HeraFirestoreInflightReadCoalescer?.installed);
  loadOnce(HERA_FIRESTORE_DIAGNOSTICS_OPTIMIZER_SRC, "hera-firestore-diagnostics-optimizer", () => window.__vargaFsOptimizerDiagnosticsExtension);
  loadOnce(HERA_SHARED_STATIC_VIEWS_SRC, "hera-shared-static-views", () => window.HeraSharedStaticViews?.installed);
  loadOnce(HERA_ACTIVE_COMMESSE_FIRST_BOOT_GUARD_SRC, "active-commesse-first-boot-guard", () => window.HeraActiveCommesseFirstBootGuard?.installed);
  loadOnce(HERA_FIRESTORE_STARTUP_COST_OPTIMIZER_SRC, "firestore-startup-cost-optimizer", () => window.HeraFirestoreStartupCostOptimizer?.installed);
  loadOnce(HERA_SHARED_STATIC_VIEWS_UI_SRC, "hera-shared-static-views-ui", () => window.__heraSharedStaticViewsUiInstalled);
  loadOnce(HERA_NATIVE_RUNTIME_SRC, "hera-native-runtime", () => false);
  loadOnce(HERA_FATTO_EMBEDDED_OPTIMIZER_SRC, "fatto-embedded-optimizer", () => window.HeraFattoEmbeddedOptimizer?.installed);
  loadOnce("notification-session-enhancements.js?v=20260727b", "hera-notification-session", () => false);
  loadOnce("update-app-feature.js?v=20260727a", "hera-app-update", () => false);
  loadOnce("google-sheet-two-way-sync.js?v=20260729b", "hera-google-sheet-sync", () => false);
  loadOnce("personnel-training-manager.js?v=20260803a", "personnel-training-manager", () => false);
}
