window.firebaseConfig = {
  apiKey: "AIzaSyBHZG9D1H5YOT9QUzG-cdSdftlreDJNa_k",
  authDomain: "hera-app-6cd2b.firebaseapp.com",
  projectId: "hera-app-6cd2b",
  storageBucket: "hera-app-6cd2b.firebasestorage.app",
  messagingSenderId: "645390631375",
  appId: "1:645390631375:web:df3659a23812560e4012ba"
};

// Carica il bridge Android prima di app.js, così la mappa usa subito
// Capacitor Geolocation e non il rilevamento posizione del browser Chrome.
if (document.readyState === "loading") {
  document.write('<script src="native-android-runtime.js?v=20260725c"><\/script>');
  document.write('<script src="notification-session-enhancements.js?v=20260727b"><\/script>');
  document.write('<script src="update-app-feature.js?v=20260727a"><\/script>');
  document.write('<script src="google-sheet-two-way-sync.js?v=20260729b"><\/script>');
  document.write('<script src="multi-organization-menu-runtime.js?v=20260802a"><\/script>');
} else {
  if (!document.querySelector('script[data-hera-native-runtime="true"]')) {
    const nativeRuntimeScript = document.createElement("script");
    nativeRuntimeScript.src = "native-android-runtime.js?v=20260725c";
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
  if (!document.querySelector('script[data-hera-multi-organization-menu="true"]')) {
    const organizationMenuScript = document.createElement("script");
    organizationMenuScript.src = "multi-organization-menu-runtime.js?v=20260802a";
    organizationMenuScript.dataset.heraMultiOrganizationMenu = "true";
    document.head.appendChild(organizationMenuScript);
  }
}
