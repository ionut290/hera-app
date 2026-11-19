(function installIsolatedAccountDeletionDemo() {
  "use strict";

  // This file is intentionally named apple-review-* and excluded from ALL
  // ordinary native web assets by scripts/prepare-capacitor-web.js.
  // Only a separately approved internal demo build may opt in.
  if (window.__VARGA_INTERNAL_DEMO_DELETION_ENABLED !== true) return;
  if (window.Capacitor?.getPlatform?.() !== "ios" ||
      window.Capacitor?.isNativePlatform?.() !== true) return;

  let latestCheck = 0;
  let deleting = false;
  let button = null;
  const TAG = "varga-internal-demo-account-delete";

  function userAuth() {
    if (!window.firebase?.apps?.length) return null;
    return window.firebase.auth();
  }

  function removeButton() {
    if (button) button.remove();
    button = null;
  }

  async function validDemo(user) {
    if (!user || typeof user.getIdTokenResult !== "function") return false;
    try {
      const claims = (await user.getIdTokenResult(true))?.claims;
      return claims?.appReviewDemo === true && claims?.apple_review === true;
    } catch (_) { return false; }
  }

  async function runDemoDeletion() {
    if (deleting) return;
    const firebaseAuth = userAuth();
    const user = firebaseAuth?.currentUser;
    if (!user || !await validDemo(user) || firebaseAuth.currentUser !== user) {
      removeButton();
      return;
    }
    // Demo Auth is password-only. Apple credential revocation is tested
    // separately in the real client; never claim this preview tests it.
    if (user.providerData?.some(p => p.providerId === "apple.com")) {
      window.alert("Questa demo non elimina gli account Apple collegati.");
      return;
    }
    if (!window.confirm("Eliminare DEFINITIVAMENTE soltanto questo account di prova?")) return;
    if (window.prompt("Scrivi ELIMINA per confermare:") !== "ELIMINA") return;
    deleting = true;
    if (button) button.disabled = true;
    try {
      // A fresh signed token check precedes even the non-mutating preflight.
      if (!await validDemo(user) || firebaseAuth.currentUser !== user) {
        throw new Error("Permessi del demo non più validi.");
      }
      const callable = window.firebase.app().functions("europe-west1")
        .httpsCallable("deleteOwnVargaAccount");
      const preflight = await callable({ confirmation: "ELIMINA", mode: "preflight" });
      if (preflight?.data?.ready !== true) throw new Error("Preflight non confermato.");
      const result = await callable({ confirmation: "ELIMINA" });
      if (result?.data?.deleted !== true) throw new Error("Server non ha confermato.");
      await firebaseAuth.signOut();
      window.alert("Account demo eliminato. Verificare Firebase Auth e i dati di test.");
      window.location.reload();
    } catch (_) {
      // No false success. The server may have partially processed a request:
      // verify its state before retrying; do not automatically retry deletes.
      window.alert("Eliminazione demo non confermata. Controllare il backend prima di riprovare.");
    } finally {
      deleting = false;
      if (button) button.disabled = false;
    }
  }

  async function refresh(user) {
    const version = ++latestCheck;
    removeButton();
    if (!await validDemo(user)) return;
    if (version !== latestCheck || userAuth()?.currentUser !== user) return;
    const parent = document.getElementById("user-details-panel");
    if (!parent) return;
    button = document.createElement("button");
    button.id = TAG;
    button.type = "button";
    button.className = "btn";
    button.textContent = "Elimina account DEMO (test interno)";
    button.addEventListener("click", () => { void runDemoDeletion(); });
    parent.appendChild(button);
  }

  function connect() {
    const auth = userAuth();
    if (auth && typeof auth.onAuthStateChanged === "function") {
      auth.onAuthStateChanged((user) => { void refresh(user); });
    }
    // Fail closed if Firebase is not ready; do not create a button from
    // guessed email, localStorage, or a client-supplied user ID.
  }

  if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", connect, { once: true });
  } else {
    connect();
  }
})();
