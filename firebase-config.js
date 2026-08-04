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
const HERA_FIRESTORE_REGISTRY_OPTIMIZER_SRC = "firestore-registry-read-optimizer.js?v=20260804b";
const HERA_REGISTRY_DEVICE_CACHE_SRC = "registry-device-cache.js?v=20260804c";

if (document.readyState === "loading") {
  // IndexedDB deve essere disponibile prima che app.js chieda personale e mezzi.
  document.write(`<script src="${HERA_REGISTRY_DEVICE_CACHE_SRC}"><\/script>`);
  // Intercetta soltanto le query duplicate delle collezioni personale e mezzi.
  document.write(`<script src="${HERA_FIRESTORE_REGISTRY_OPTIMIZER_SRC}"><\/script>`);
  document.write(`<script src="${HERA_NATIVE_RUNTIME_SRC}"><\/script>`);
  document.write('<script src="notification-session-enhancements.js?v=20260727b"><\/script>');
  document.write('<script src="update-app-feature.js?v=20260727a"><\/script>');
  document.write('<script src="google-sheet-two-way-sync.js?v=20260729b"><\/script>');
  document.write('<script src="personnel-training-manager.js?v=20260803a"><\/script>');
} else {
  function loadRegistryOptimizer() {
    if (window.HeraFirestoreRegistryOptimizer?.installed ||
        document.querySelector('script[data-hera-firestore-registry-optimizer="true"]')) return;
    const optimizerScript = document.createElement("script");
    optimizerScript.src = HERA_FIRESTORE_REGISTRY_OPTIMIZER_SRC;
    optimizerScript.dataset.heraFirestoreRegistryOptimizer = "true";
    document.head.appendChild(optimizerScript);
  }

  if (!window.HeraRegistryDeviceCache &&
      !document.querySelector('script[data-hera-registry-device-cache="true"]')) {
    const cacheScript = document.createElement("script");
    cacheScript.src = HERA_REGISTRY_DEVICE_CACHE_SRC;
    cacheScript.dataset.heraRegistryDeviceCache = "true";
    cacheScript.addEventListener("load", loadRegistryOptimizer, { once: true });
    document.head.appendChild(cacheScript);
  } else {
    loadRegistryOptimizer();
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
