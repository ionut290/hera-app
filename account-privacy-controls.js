(function installAccountPrivacyControls() {
  "use strict";
  if (window.__vargaAccountPrivacyInstalled) return;
  window.__vargaAccountPrivacyInstalled = true;

  const REGION = "europe-west1";
  // Fail closed until the deletion backend and privacy release are actually live.
  // This does NOT gate Sign in with Apple; only the destructive account-deletion UI.
  const deletionEnabled = window.__VARGA_ACCOUNT_DELETION_ENABLED === true;
  // Only enable after the callable and fulfillment process are authorized and deployed.
  // Require the independently approved fulfillment SLA, never invent a default.
  const requestSlaDays = window.__VARGA_ACCOUNT_DELETION_REQUEST_SLA_DAYS;
  const validRequestSla = Number.isSafeInteger(requestSlaDays) &&
    requestSlaDays >= 1 && requestSlaDays <= 90;
  const requestEnabled = window.__VARGA_ACCOUNT_DELETION_REQUEST_ENABLED === true &&
    !deletionEnabled && validRequestSla;
  let deleting = false;
  let requestingDeletion = false;
  let checkingDeletionRequestStatus = false;
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
    if (code.includes("unauthorized-domain")) {
      return "Il dominio di questa web-app non è autorizzato per Apple in Firebase Authentication. Contatta l'amministratore.";
    }
    if (code.includes("popup-blocked")) {
      return "Il browser ha bloccato la finestra di accesso Apple. Consenti i popup per VARGA CANTIERI e riprova.";
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
      if (nativePlatform() === "ios" || nativePlatform() === "android") {
        const plugin = window.Capacitor?.Plugins?.FirebaseAuthentication;
        if (!plugin || typeof plugin.signInWithApple !== "function") {
          throw new Error("Login Apple nativo non disponibile in questa versione dell'app. Aggiorna VARGA CANTIERI.");
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
    // Lock immediately, including preflight and Apple revocation. A second tap
    // must not start a concurrent deletion or revoke the same Apple code twice.
    deleting = true;
    // This flag distinguishes an already-revoked Apple grant from a confirmed
    // account deletion. If Firebase cleanup later fails, the user needs a
    // specific recovery message rather than a misleading generic retry.
    let appleAuthorizationRevoked = false;
    // Preflight first: do not revoke Apple access if the server already knows
    // that personal files cannot be safely removed.
    let callable;
    try {
      callable = firebase.app().functions(REGION).httpsCallable("deleteOwnVargaAccount");
      const check = await callable({ confirmation: "ELIMINA", mode: "preflight" });
      if (check?.data?.ready !== true) throw new Error("Verifica cancellazione non confermata.");
    } catch (error) {
      alert("Impossibile procedere con la cancellazione. " + errorText(error));
      deleting = false;
      return;
    }
    // Apple-linked identities require Apple OAuth revocation before account
    // deletion. The Android SDK revokes the returned OAuth access token,
    // while iOS revokes a fresh Apple authorization code. On both platforms
    // failure or missing token blocks deletion. Web remains fail-closed.
    const appleLinked = user.providerData?.some((provider) => provider.providerId === "apple.com");
    const platform = nativePlatform();
    if (appleLinked && platform !== "ios" && platform !== "android") {
      alert("La cancellazione diretta di un account collegato ad Apple non è ancora disponibile nel browser. Usa l'app per iPhone o Android aggiornata, oppure contatta l'assistenza. Nessun account è stato eliminato.");
      deleting = false;
      return;
    }
    // Reauthenticate the same Firebase account before revoking Apple's grant.
    // iOS revokes a fresh authorization code; Android uses the OAuth access
    // token from the native sign-in result. Never fall back between token types.
    if (appleLinked) {
      const plugin = window.Capacitor?.Plugins?.FirebaseAuthentication;
      if (!plugin || typeof plugin.signInWithApple !== "function" ||
          typeof plugin.revokeAccessToken !== "function") {
        alert("Impossibile revocare l'accesso Apple su questo dispositivo. Aggiorna l'app e riprova: l'account non è stato eliminato.");
        deleting = false;
        return;
      }
      {
        try {
          const fresh = await plugin.signInWithApple({ skipNativeAuth: true });
          const appleProvider = new firebase.auth.OAuthProvider("apple.com");
          const idToken = fresh?.credential?.idToken;
          const rawNonce = fresh?.credential?.nonce;
          const token = platform === "android"
            ? fresh?.credential?.accessToken
            : fresh?.credential?.authorizationCode;
          if (!idToken || !rawNonce || !token) {
            throw new Error("Impossibile ottenere il token Apple necessario per la revoca su questo dispositivo.");
          }
          const credential = appleProvider.credential({ idToken, rawNonce });
          await user.reauthenticateWithCredential(credential);
          await plugin.revokeAccessToken({ token });
          appleAuthorizationRevoked = true;
        } catch (error) {
          alert("Verifica Apple non riuscita: " + errorText(error) +
            "\nL'account non è stato eliminato.");
          deleting = false;
          return;
        }
      }
    }
    const buttons = document.querySelectorAll("[data-delete-varga-account]");
    buttons.forEach((button) => { button.disabled = true; button.textContent = "Eliminazione in corso..."; });
    try {
      const result = await callable({ confirmation: "ELIMINA" });
      if (result?.data?.deleted !== true) throw new Error("Il server non ha confermato l'eliminazione.");
      try { await auth().signOut(); } catch (_) {}
      try { await window.Capacitor?.Plugins?.FirebaseAuthentication?.signOut?.(); } catch (_) {}
      alert("Account eliminato. I dati personali previsti dalla procedura sono stati cancellati.");
      window.location.reload();
    } catch (error) {
      console.warn("[varga-privacy] Account deletion:", error?.code || "failed");
      const message = appleAuthorizationRevoked
        ? "L’autorizzazione Apple è stata revocata, ma la cancellazione dell’account NON è stata confermata dal server. Non considerare l’account eliminato: contatta l’assistenza per completare la procedura in sicurezza. " + errorText(error)
        : String(error?.code || "").includes("failed-precondition")
          ? "Per sicurezza esci e accedi nuovamente, poi ripeti l'eliminazione. " + errorText(error)
          : "La cancellazione non è stata confermata dal server. " + errorText(error);
      alert(message);
    } finally {
      deleting = false;
      buttons.forEach((button) => { button.disabled = false; button.textContent = "Elimina il mio account"; });
    }
  }

  // Request initiation is NON-DESTRUCTIVE and separate from immediate deletion.
  // Never show this UI until the reviewed callable exists in production.
  async function requestDeletion() {
    if (requestingDeletion) return;
    let user;
    try { user = auth().currentUser; } catch (error) { alert(errorText(error)); return; }
    if (!user) { alert("Per richiedere la cancellazione devi prima accedere."); return; }
    if (!window.confirm("Richiedere l'eliminazione del tuo account VARGA CANTIERI? Il completamento è previsto entro " + requestSlaDays + " giorni secondo la procedura approvata. Alcuni dati aziendali possono essere conservati quando necessario.")) return;
    if (window.prompt("Per confermare scrivi RICHIEDO (tutto maiuscolo):") !== "RICHIEDO") return;
    requestingDeletion = true;
    const buttons = document.querySelectorAll("[data-request-varga-deletion]");
    buttons.forEach((button) => { button.disabled = true; });
    try {
      const callable = firebase.app().functions(REGION).httpsCallable("requestOwnVargaAccountDeletion");
      const result = await callable({ confirmation: "RICHIEDO ELIMINAZIONE" });
      if (result?.data?.requested !== true ||
          !["pending", "in_review", "completed", "rejected"].includes(result?.data?.status)) {
        throw new Error("Il server non ha confermato la richiesta.");
      }
      if (result.data.status === "pending" || result.data.status === "in_review") {
        alert("Richiesta di cancellazione registrata. Il completamento è previsto entro " + requestSlaDays + " giorni secondo la procedura approvata; riceverai conferma quando l'eliminazione sarà completata. L'account NON è ancora eliminato.");
      } else {
        alert("Esiste già una richiesta con stato " + result.data.status + ". Contatta l'assistenza per i dettagli: nessuna nuova cancellazione è stata eseguita.");
      }
    } catch (error) {
      alert("La richiesta non è stata confermata dal server. " + errorText(error));
    } finally {
      requestingDeletion = false;
      buttons.forEach((button) => { button.disabled = false; });
    }
  }

  // Explicit, user-triggered status fetch. No background collection reads,
  // employee data, or status guess from local cache. Requires signed-in UID.
  async function checkDeletionRequestStatus() {
    if (checkingDeletionRequestStatus) return;
    let user;
    try { user = auth().currentUser; } catch (error) { alert(errorText(error)); return; }
    if (!user) {
      alert("Accedi per verificare lo stato della tua richiesta.");
      return;
    }
    checkingDeletionRequestStatus = true;
    const buttons = document.querySelectorAll("[data-check-varga-deletion-status]");
    buttons.forEach((button) => { button.disabled = true; });
    try {
      const callable = firebase.app().functions(REGION).httpsCallable("getOwnVargaAccountDeletionStatus");
      const result = await callable({});
      const state = result?.data;
      if (state?.requested === false) {
        alert("Nessuna richiesta di cancellazione registrata per il tuo account.");
      } else if (state?.requested === true) {
        if (state.status === "pending") {
          alert("Richiesta ricevuta e in attesa di verifica. L'account NON è ancora eliminato.");
        } else if (state.status === "in_review") {
          alert("La richiesta è in revisione. L'account NON è ancora eliminato.");
        } else if (state.status === "completed") {
          alert("La richiesta risulta completata. Conserva la conferma ufficiale e contatta l'assistenza per eventuali chiarimenti.");
        } else if (state.status === "rejected") {
          // Display only a server-approved reason code. Never print operator
          // notes or arbitrary backend text that could disclose personal data.
          const explanations = {
            identity_not_verified: "Non è stato possibile verificare l'identità del richiedente.",
            retention_review_required: "Occorre verificare gli obblighi di conservazione applicabili.",
            request_needs_clarification: "Servono ulteriori chiarimenti per esaminare la richiesta."
          };
          const reason = Object.prototype.hasOwnProperty.call(explanations, state.reasonCode)
            ? explanations[state.reasonCode]
            : "La motivazione non è disponibile nell'app.";
          alert("La richiesta è stata chiusa senza cancellazione. " + reason +
            " Se desideri ripresentarla, accedi nuovamente e usa 'Richiedi cancellazione account'. Per chiarimenti contatta l'assistenza.");
        } else {
          throw new Error("Stato richiesta non riconosciuto.");
        }
      } else {
        throw new Error("Risposta di stato non valida.");
      }
    } catch (error) {
      alert("Impossibile verificare lo stato della richiesta. " + errorText(error));
    } finally {
      checkingDeletionRequestStatus = false;
      buttons.forEach((button) => { button.disabled = false; });
    }
  }

  function mount() {
    const gateGoogle = document.getElementById("auth-gate-login-btn");
    // Apple OAuth is a web popup only on the external WEB site; Android and
    // iOS use the Capacitor native FirebaseAuthentication plugin. Never open
    // a Firebase browser popup inside a native WebView.
    if (gateGoogle && !document.getElementById("auth-apple-login-btn")) {
      const apple = makeButton("auth-apple-login-btn", "Accedi con Apple", "btn");
      apple.setAttribute("aria-label", "Accedi con Apple");
      gateGoogle.insertAdjacentElement("afterend", apple);
      apple.addEventListener("click", signInWithApple);
    }

    // Keep account privacy options discoverable even while the production
    // deletion/request backend is intentionally disabled. This is INFO ONLY:
    // it never calls Firebase, promises a deadline, revokes Apple, signs out,
    // or pretends a request has been submitted. Do not confuse UI visibility
    // with completion of Apple's in-app account-deletion requirement.
    if (!requestEnabled && !deletionEnabled) {
      const locations = [
        [document.getElementById("user-details-panel"), "varga-deletion-info-profile"],
        [document.querySelector("#side-menu .side-menu-body .menu-section:nth-child(2)"), "varga-deletion-info-menu"],
        [document.getElementById("access-approval-logout")?.parentElement, "varga-deletion-info-pending"],
        [(document.getElementById("auth-apple-login-btn") || gateGoogle)?.parentElement,
          "varga-deletion-info-gate"]
      ];
      for (const [parent, id] of locations) {
        if (!parent || document.getElementById(id)) continue;
        // The published /delete-account.html already documents the actual
        // assistance route and clearly says this is NOT in-app deletion.
        // A real hyperlink lets users open that information even if scripts
        // or popups are blocked, without calling a destructive backend.
        const link = document.createElement("a");
        link.id = id;
        link.href = "./delete-account.html";
        link.textContent = "Elimina account (informazioni e assistenza)";
        link.className = "btn";
        parent.appendChild(link);
      }
    }

    if (requestEnabled) {
      // User-initiated request available to authenticated normal, pending and
      // newly registered accounts once the backend is actually operational.
      const profileNode = document.getElementById("user-details-panel");
      const menuNode = document.querySelector("#side-menu .side-menu-body .menu-section:nth-child(2)");
      const pendingNode = document.getElementById("access-approval-logout");
      const gateNode = document.getElementById("auth-apple-login-btn") || gateGoogle;
      for (const [parent, id, after] of [
        [profileNode, "varga-deletion-request-profile", null],
        [menuNode, "varga-deletion-request-menu", null],
        [pendingNode?.parentElement, "varga-deletion-request-pending", pendingNode],
        [gateNode?.parentElement, "varga-deletion-request-gate", gateNode]
      ]) {
        if (!parent || document.getElementById(id)) continue;
        const button = makeButton(id, "Richiedi cancellazione account", "btn");
        button.dataset.requestVargaDeletion = "1";
        button.addEventListener("click", requestDeletion);
        if (after) after.insertAdjacentElement("afterend", button);
        else parent.appendChild(button);
        // The matching status control never creates or duplicates a request.
        const statusId = id + "-status";
        if (!document.getElementById(statusId)) {
          const statusButton = makeButton(statusId, "Verifica stato cancellazione", "btn");
          statusButton.dataset.checkVargaDeletionStatus = "1";
          statusButton.addEventListener("click", checkDeletionRequestStatus);
          button.insertAdjacentElement("afterend", statusButton);
        }
      }
    }

    if (!deletionEnabled) return;
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