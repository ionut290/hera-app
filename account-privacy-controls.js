(function installAccountPrivacyControls() {
  "use strict";
  if (window.__vargaAccountPrivacyInstalled) return;
  window.__vargaAccountPrivacyInstalled = true;

  const REGION = "europe-west1";
  let deleting = false;
  let signingIn = false;

  function nativePlatform() {
    const cap = window.Capacitor;
    return cap?.isNativePlatform?.() ? cap.getPlatform?.() : "";
  }

  function auth() {
    if (!window.firebase || typeof firebase.auth !== "function" || !firebase.apps?.length) {
      throw new Error("Firebase Authentication non è ancora pronto. Riprova.");
    }
    return firebase.auth();
  }

  function makeButton(id, text, className) {
    const button = document.createElement("button");
    button.id = id;
    button.type = "button";
    button.textContent = text;
    button.className = className;
    return button;
  }

  function errorText(error) {
    const code = String(error?.code || "");
    if (code.includes("popup-closed") || code.includes("cancelled")) return "Accesso annullato.";
    if (code.includes("account-exists-with-different-credential")) {
      return "Esiste già un account con questa email. Accedi con il metodo originale e contatta l'amministratore per collegare Apple.";
    }
    if (code.includes("operation-not-allowed")) {
      return "Accedi con Apple non è ancora attivato in Firebase. Contatta l'amministratore.";
    }
    return String(error?.message || "Operazione non riuscita. Riprova.");
  }

  async function signInWithApple() {
    if (signingIn) return;
    signingIn = true;
    const button = document.getElementById("auth-apple-login-btn");
    if (button) button.disabled = true;
    try {
      const currentAuth = auth();
      const provider = new firebase.auth.OAuthProvider("apple.com");
      if (nativePlatform() === "ios") {
        const plugin = window.Capacitor?.Plugins?.FirebaseAuthentication;
        if (!plugin || typeof plugin.signInWithApple !== "function") {
          throw new Error("Login Apple nativo non disponibile in questa versione dell'app.");
        }
        const result = await plugin.signInWithApple({ skipNativeAuth: true });
        const idToken = result?.credential?.idToken;
        const rawNonce = result?.credential?.nonce; // Capacitor FirebaseAuthentication: credential.nonce is the raw nonce
        if (!idToken || !rawNonce) {
          throw new Error("Apple non ha restituito i dati di accesso necessari. Riprova.");
        }
        const credential = provider.credential({ idToken, rawNonce });
        await currentAuth.signInWithCredential(credential);
      } else {
        await currentAuth.signInWithPopup(provider);
      }
    } catch (error) {
      console.warn("[varga-auth] Apple login:", error?.code || "failed");
      alert(errorText(error));
    } finally {
      signingIn = false;
      if (button) button.disabled = false;
    }
  }

  async function deleteCurrentAccount() {
    if (deleting) return;
    let user;
    try { user = auth().currentUser; } catch (error) { alert(errorText(error)); return; }
    if (!user) { alert("Per eliminare l'account devi prima accedere."); return; }
    const warning = "Eliminare definitivamente l'account " + (user.email || "") +
      " e i relativi dati personali?\n\nI documenti e gli interventi aziendali condivisi potranno essere conservati quando necessario. Questa operazione non si può annullare.";
    if (!window.confirm(warning)) return;
    const confirmation = window.prompt('Per confermare scrivi ELIMINA (tutto maiuscolo):');
    if (confirmation !== "ELIMINA") {
      if (confirmation !== null) alert("Conferma non valida. Account non eliminato.");
      return;
    }
    // Apple requires credential revocation when deleting an Apple-linked account.
    // Request a fresh Apple code, reauthenticate the same Firebase account, then revoke.
    if (nativePlatform() === "ios" &&
        user.providerData?.some((provider) => provider.providerId === "apple.com")) {
      const plugin = window.Capacitor?.Plugins?.FirebaseAuthentication;
      if (plugin?.signInWithApple) {
        try {
          const fresh = await plugin.signInWithApple({ skipNativeAuth: true });
          const appleProvider = new firebase.auth.OAuthProvider("apple.com");
          const idToken = fresh?.credential?.idToken;
          const rawNonce = fresh?.credential?.nonce;
          const code = fresh?.credential?.authorizationCode;
          if (!idToken || !rawNonce || !code || typeof plugin.revokeAccessToken !== "function") {
            throw new Error("Impossibile ottenere le credenziali per la revoca dell'accesso Apple.");
          }
          const credential = appleProvider.credential({ idToken, rawNonce });
          await user.reauthenticateWithCredential(credential);
          await plugin.revokeAccessToken({ token: code });
        } catch (error) {
          alert("Verifica Apple non riuscita: " + errorText(error) +
            "\nL'account non è stato eliminato.");
          return;
        }
      }
    }
    deleting = true;
    const buttons = document.querySelectorAll("[data-delete-varga-account]");
    buttons.forEach((button) => { button.disabled = true; button.textContent = "Eliminazione in corso..."; });
    try {
      const callable = firebase.app().functions(REGION).httpsCallable("deleteOwnVargaAccount");
      const result = await callable({ confirmation: "ELIMINA" });
      if (result?.data?.deleted !== true) throw new Error("Il server non ha confermato l'eliminazione.");
      try { await auth().signOut(); } catch (_) {}
      try { await window.Capacitor?.Plugins?.FirebaseAuthentication?.signOut?.(); } catch (_) {}
      alert("Account eliminato. I dati personali previsti dalla procedura sono stati cancellati.");
      window.location.reload();
    } catch (error) {
      console.warn("[varga-privacy] Account deletion:", error?.code || "failed");
      const message = String(error?.code || "").includes("failed-precondition")
        ? "Per sicurezza esci e accedi nuovamente, poi ripeti l'eliminazione. " + errorText(error)
        : "La cancellazione non è stata confermata dal server. " + errorText(error);
      alert(message);
    } finally {
      deleting = false;
      buttons.forEach((button) => { button.disabled = false; button.textContent = "Elimina il mio account"; });
    }
  }

  function mount() {
    const gateGoogle = document.getElementById("auth-gate-login-btn");
    if (gateGoogle && !document.getElementById("auth-apple-login-btn") && nativePlatform() !== "android") {
      const apple = makeButton("auth-apple-login-btn", "Accedi con Apple", "btn");
      apple.setAttribute("aria-label", "Accedi con Apple");
      gateGoogle.insertAdjacentElement("afterend", apple);
      apple.addEventListener("click", signInWithApple);
    }

    // Provide deletion in the profile, normal menu, AND pending-user view.
    const locations = [
      ["user-details-panel", "varga-account-delete-profile", "btn"],
      ["access-approval-logout", "varga-account-delete-pending", "btn"]
    ];
    const profile = document.getElementById(locations[0][0]);
    const menu = document.querySelector("#side-menu .side-menu-body .menu-section:nth-child(2)");
    const pending = document.getElementById(locations[1][0]);

    function attach(parent, id, className, after) {
      if (!parent || document.getElementById(id)) return;
      const button = makeButton(id, "Elimina il mio account", className);
      button.dataset.deleteVargaAccount = "1";
      button.addEventListener("click", deleteCurrentAccount);
      if (after) after.insertAdjacentElement("afterend", button);
      else parent.appendChild(button);
    }
    attach(profile, "varga-account-delete-profile", "btn");
    attach(menu, "varga-account-delete-menu", "btn menu-title-btn");
    attach(pending?.parentElement, "varga-account-delete-pending", "btn", pending);
    // An unverified newly-created email account can remain on the login gate:
    // make account deletion discoverable there as well.
    const gateAnchor = document.getElementById("auth-apple-login-btn") || gateGoogle;
    attach(gateAnchor?.parentElement, "varga-account-delete-gate", "btn", gateAnchor);
  }

  if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", mount, { once: true });
  } else {
    mount();
  }
})();