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
} else if (!document.querySelector('script[data-hera-native-runtime="true"]')) {
  const nativeRuntimeScript = document.createElement("script");
  nativeRuntimeScript.src = "native-android-runtime.js?v=20260725c";
  nativeRuntimeScript.dataset.heraNativeRuntime = "true";
  document.head.appendChild(nativeRuntimeScript);
}
