(function () {
  "use strict";

  const match = String(window.location.hash || "").match(/(?:^#|&)vargaSso=([^&]+)/);
  if (!match) return;
  const token = decodeURIComponent(match[1] || "").trim();
  window.history.replaceState(null, "", `${window.location.pathname}${window.location.search}`);
  if (!token || !window.firebase?.auth) return;

  window.__vargaSessionTransfer = firebase.auth()
    .setPersistence(firebase.auth.Auth.Persistence.LOCAL)
    .then(() => firebase.auth().signInWithCustomToken(token))
    .catch((error) => {
      console.warn("Accesso automatico da Varga Gestionale non riuscito:", error);
      return null;
    });
})();
