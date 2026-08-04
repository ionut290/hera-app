window.firebaseConfig = {
  apiKey: "AIzaSyBHZG9D1H5YOT9QUzG-cdSdftlreDJNa_k",
  authDomain: "hera-app-6cd2b.firebaseapp.com",
  projectId: "hera-app-6cd2b",
  storageBucket: "hera-app-6cd2b.firebasestorage.app",
  messagingSenderId: "645390631375",
  appId: "1:645390631375:web:df3659a23812560e4012ba"
};

// Carica il bridge Android prima di app.js, così la mappa usa subito
// Capacitor Geolocation e la PWA installa subito il blocco WhatsApp Web.
const HERA_NATIVE_RUNTIME_SRC = "native-android-runtime.js?v=20260803-whatsapp-early2";
const HERA_FIRESTORE_INFLIGHT_COALESCER_SRC = "firestore-inflight-read-coalescer.js?v=20260804b";
const HERA_FIRESTORE_DIAGNOSTICS_OPTIMIZER_SRC = "firestore-diagnostics-optimizer-extension.js?v=20260804b";
const HERA_SHARED_STATIC_VIEWS_SRC = "shared-static-views.js?v=20260804a";
const HERA_SHARED_STATIC_VIEWS_UI_SRC = "shared-static-views-ui.js?v=20260804a";

if (document.readyState === "loading") {
  document.write(`<script src="${HERA_FIRESTORE_INFLIGHT_COALESCER_SRC}"><\/script>`);
  document.write(`<script src="${HERA_FIRESTORE_DIAGNOSTICS_OPTIMIZER_SRC}"><\/script>`);
  document.write(`<script src="${HERA_SHARED_STATIC_VIEWS_SRC}"><\/script>`);
  document.write(`<script src="${HERA_SHARED_STATIC_VIEWS_UI_SRC}"><\/script>`);
  document.write(`<script src="${HERA_NATIVE_RUNTIME_SRC}"><\/script>`);
  document.write('<script src="notification-session-enhancements.js?v=20260727b"><\/script>');
  document.write('<script src="update-app-feature.js?v=20260727a"><\/script>');
  document.write('<script src="google-sheet-two-way-sync.js?v=20260729b"><\/script>');
  document.write('<script src="personnel-training-manager.js?v=20260803a"><\/script>');
} else {
  function loadOnce(src, dataName, ready) {
    if (ready?.() || document.querySelector(`script[data-${dataName}="true"]`)) return;
    const script = document.createElement("script");
    script.src = src;
    script.setAttribute(`data-${dataName}`, "true");
    document.head.appendChild(script);
  }

  loadOnce(HERA_FIRESTORE_INFLIGHT_COALESCER_SRC, "hera-firestore-inflight-coalescer", () => window.HeraFirestoreInflightReadCoalescer?.installed);
  loadOnce(HERA_FIRESTORE_DIAGNOSTICS_OPTIMIZER_SRC, "hera-firestore-diagnostics-optimizer", () => window.__vargaFsOptimizerDiagnosticsExtension);
  loadOnce(HERA_SHARED_STATIC_VIEWS_SRC, "hera-shared-static-views", () => window.HeraSharedStaticViews?.installed);
  loadOnce(HERA_SHARED_STATIC_VIEWS_UI_SRC, "hera-shared-static-views-ui", () => window.__heraSharedStaticViewsUiInstalled);
  loadOnce(HERA_NATIVE_RUNTIME_SRC, "hera-native-runtime", () => false);
  loadOnce("notification-session-enhancements.js?v=20260727b", "hera-notification-session", () => false);
  loadOnce("update-app-feature.js?v=20260727a", "hera-app-update", () => false);
  loadOnce("google-sheet-two-way-sync.js?v=20260729b", "hera-google-sheet-sync", () => false);
  loadOnce("personnel-training-manager.js?v=20260803a", "personnel-training-manager", () => false);
}
