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

if (document.readyState === "loading") {
  // Condivide soltanto richieste identiche ancora in corso e restituisce
  // sempre lo stesso QuerySnapshot originale ricevuto da Firestore.
  document.write(`<script src="${HERA_FIRESTORE_INFLIGHT_COALESCER_SRC}"><\/script>`);
  // Mantiene disponibili i contatori diagnostici del coalescer.
  document.write(`<script src="${HERA_FIRESTORE_DIAGNOSTICS_OPTIMIZER_SRC}"><\/script>`);
  // Prepara le viste condivise senza sostituire i caricamenti operativi esistenti.
  document.write(`<script src="${HERA_SHARED_STATIC_VIEWS_SRC}"><\/script>`);
  document.write(`<script src="${HERA_NATIVE_RUNTIME_SRC}"><\/script>`);
  document.write('<script src="notification-session-enhancements.js?v=20260727b"><\/script>');
  document.write('<script src="update-app-feature.js?v=20260727a"><\/script>');
  document.write('<script src="google-sheet-two-way-sync.js?v=20260729b"><\/script>');
  document.write('<script src="personnel-training-manager.js?v=20260803a"><\/script>');
} else {
  if (!window.HeraFirestoreInflightReadCoalescer?.installed &&
      !document.querySelector('script[data-hera-firestore-inflight-coalescer="true"]')) {
    const script = document.createElement("script");
    script.src = HERA_FIRESTORE_INFLIGHT_COALESCER_SRC;
    script.dataset.heraFirestoreInflightCoalescer = "true";
    document.head.appendChild(script);
  }

  if (!window.__vargaFsOptimizerDiagnosticsExtension &&
      !document.querySelector('script[data-hera-firestore-diagnostics-optimizer="true"]')) {
    const diagnosticsOptimizerScript = document.createElement("script");
    diagnosticsOptimizerScript.src = HERA_FIRESTORE_DIAGNOSTICS_OPTIMIZER_SRC;
    diagnosticsOptimizerScript.dataset.heraFirestoreDiagnosticsOptimizer = "true";
    document.head.appendChild(diagnosticsOptimizerScript);
  }

  if (!window.HeraSharedStaticViews?.installed &&
      !document.querySelector('script[data-hera-shared-static-views="true"]')) {
    const sharedViewsScript = document.createElement("script");
    sharedViewsScript.src = HERA_SHARED_STATIC_VIEWS_SRC;
    sharedViewsScript.dataset.heraSharedStaticViews = "true";
    document.head.appendChild(sharedViewsScript);
  }

  if (!document.querySelector('script[data-hera-native-runtime="true"]')) {
    const nativeRuntimeScript = document.createElement("script");
    nativeRuntimeScript.src = HERA_NATIVE_RUNTIME_SRC;
    nativeRuntimeScript.dataset.heraNativeRuntime = "true";
    document.head.appendChild(nativeRuntimeScript);
  }
  if (!document.querySelector('script[data-hera-notification-session="true"]')) {
    const enhancementScript = document.createElement("script");
    enhancementScript.src = "notification-session-enhancements.js?v=20260727b";
    enhancementScript.dataset.heraNotificationSession = "true";
    document.head.appendChild(enhancementScript);
  }
  if (!document.querySelector('script[data-hera-app-update="true"]')) {
    const updateScript = document.createElement("script");
    updateScript.src = "update-app-feature.js?v=20260727a";
    updateScript.dataset.heraAppUpdate = "true";
    document.head.appendChild(updateScript);
  }
  if (!document.querySelector('script[data-hera-google-sheet-sync="true"]')) {
    const sheetSyncScript = document.createElement("script");
    sheetSyncScript.src = "google-sheet-two-way-sync.js?v=20260729b";
    sheetSyncScript.dataset.heraGoogleSheetSync = "true";
    document.head.appendChild(sheetSyncScript);
  }
  if (!document.querySelector('script[data-personnel-training-manager="true"]')) {
    const trainingScript = document.createElement("script");
    trainingScript.src = "personnel-training-manager.js?v=20260803a";
    trainingScript.dataset.personnelTrainingManager = "true";
    document.head.appendChild(trainingScript);
  }
}
