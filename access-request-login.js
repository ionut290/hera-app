(() => {
  "use strict";

  if (window.HeraAccessRequestLogin?.installed) return;

  const ENDPOINT = "https://europe-west1-hera-app-6cd2b.cloudfunctions.net/registerTester";
  let pendingCredential = null;
  let busy = false;

  function text(value) {
    return String(value ?? "").trim();
  }

  function ensureUi() {
    const form = document.getElementById("auth-email-form");
    if (!form) return false;

    let button = document.getElementById("auth-request-access-btn");
    if (!button) {
      button = document.createElement("button");
      button.id = "auth-request-access-btn";
      button.type = "button";
      button.className = "btn";
      button.textContent = "RICHIEDI ACCESSO";
      const forgot = document.getElementById("auth-request-password-btn");
      if (forgot?.parentNode) forgot.parentNode.insertBefore(button, forgot.nextSibling);
      else form.appendChild(button);
      button.addEventListener("click", openDialog);
    }

    if (!document.getElementById("access-request-login-dialog")) {
      const dialog = document.createElement("dialog");
      dialog.id = "access-request-login-dialog";
      dialog.className = "biometric-dialog registration-dialog";
      dialog.innerHTML = `
        <form id="access-request-login-form">
          <h2>Richiedi accesso</h2>
          <p>Inserisci nome e cognome. Se sei già presente nel personale, ti mostreremo lo username corretto. Se non risulti presente, la richiesta verrà inviata all’amministratore.</p>
          <label for="access-request-first-name">Nome</label>
          <input id="access-request-first-name" type="text" autocomplete="given-name" maxlength="80" required>
          <label for="access-request-last-name">Cognome</label>
          <input id="access-request-last-name" type="text" autocomplete="family-name" maxlength="80" required>
          <p id="access-request-feedback" class="muted" role="status" aria-live="polite"></p>
          <div class="actions-row">
            <button id="access-request-cancel-btn" class="btn" type="button">ANNULLA</button>
            <button id="access-request-submit-btn" class="btn btn-primary" type="submit">CONTINUA</button>
          </div>
        </form>
      `;
      document.body.appendChild(dialog);
      dialog.querySelector("#access-request-cancel-btn")?.addEventListener("click", () => dialog.close());
      dialog.querySelector("#access-request-login-form")?.addEventListener("submit", submitRequest);
    }
    return true;
  }

  function openDialog() {
    const dialog = document.getElementById("access-request-login-dialog");
    if (!dialog) return;
    const feedback = dialog.querySelector("#access-request-feedback");
    const submit = dialog.querySelector("#access-request-submit-btn");
    if (feedback) feedback.textContent = "";
    if (submit) submit.textContent = "CONTINUA";
    if (!dialog.open) dialog.showModal();
    window.setTimeout(() => dialog.querySelector("#access-request-first-name")?.focus(), 0);
  }

  async function callEndpoint(payload, authenticated = false) {
    const headers = { "Content-Type": "application/json" };
    if (authenticated) {
      const user = window.firebase?.auth?.()?.currentUser;
      if (user) headers.Authorization = `Bearer ${await user.getIdToken(true)}`;
    }

    let response;
    try {
      response = await fetch(ENDPOINT, {
        method: "POST",
        headers,
        body: JSON.stringify({ data: payload })
      });
    } catch (networkError) {
      const offline = navigator.onLine === false;
      const error = new Error(offline
        ? "Connessione assente. Impossibile verificare il personale in questo momento."
        : "Servizio accessi temporaneamente non disponibile. Riprova tra qualche secondo.");
      error.code = offline ? "offline" : "network-error";
      error.cause = networkError;
      throw error;
    }

    const body = await response.json().catch(() => ({}));
    if (!response.ok || body.error) {
      const error = new Error(body?.error?.message || `Operazione non riuscita (${response.status}).`);
      error.code = body?.error?.status || "request-failed";
      throw error;
    }
    return body.result || body.data || body;
  }

  async function submitRequest(event) {
    event.preventDefault();
    if (busy) return;
    const dialog = document.getElementById("access-request-login-dialog");
    const firstName = text(dialog?.querySelector("#access-request-first-name")?.value);
    const lastName = text(dialog?.querySelector("#access-request-last-name")?.value);
    const feedback = dialog?.querySelector("#access-request-feedback");
    const submit = dialog?.querySelector("#access-request-submit-btn");
    if (!firstName || !lastName) {
      if (feedback) feedback.textContent = "Inserisci nome e cognome.";
      return;
    }

    busy = true;
    if (submit) submit.disabled = true;
    if (feedback) feedback.textContent = "Controllo nel personale...";
    try {
      const result = await callEndpoint({ action: "requestAccessLookup", firstName, lastName });
      if (result?.status === "found" && result?.username) {
        const usernameInput = document.getElementById("auth-email-input");
        if (usernameInput) {
          usernameInput.type = "text";
          usernameInput.value = result.username;
          usernameInput.dispatchEvent(new Event("input", { bubbles: true }));
        }
        if (dialog?.open) dialog.close();
        const loginFeedback = document.getElementById("auth-email-feedback");
        if (loginFeedback) loginFeedback.textContent = `✅ Operatore trovato. Username: ${result.username}. Inserisci la password.`;
        document.getElementById("auth-password-input")?.focus();
        return;
      }
      if (feedback) {
        feedback.textContent = result?.existingPersonnel
          ? "✅ Sei presente nel personale, ma l’account deve essere preparato. Richiesta inviata all’amministratore."
          : "✅ Richiesta inviata all’amministratore. Riceverai username e password dopo l’approvazione.";
      }
      if (submit) submit.textContent = "RICHIESTA INVIATA";
    } catch (error) {
      console.error("Richiesta accesso fallita:", error);
      if (feedback) feedback.textContent = error?.message || "Richiesta non riuscita. Riprova.";
    } finally {
      busy = false;
      if (submit) submit.disabled = false;
    }
  }

  function captureCredentials(event) {
    const form = event.target;
    if (!(form instanceof HTMLFormElement) || form.id !== "auth-email-form") return;
    const username = text(document.getElementById("auth-email-input")?.value);
    const password = String(document.getElementById("auth-password-input")?.value || "");
    if (!username || password.length < 6) return;
    pendingCredential = { username, password };
  }

  async function saveCredentialSecurely() {
    if (!pendingCredential) return;
    const credential = pendingCredential;
    pendingCredential = null;
    try {
      if (window.PasswordCredential && navigator.credentials?.store) {
        await navigator.credentials.store(new PasswordCredential({
          id: credential.username,
          password: credential.password,
          name: credential.username
        }));
      }
    } catch (error) {
      console.warn("Salvataggio nel gestore password del dispositivo non disponibile:", error);
    }
  }

  function installAuthSuccessListener() {
    const firebaseAuth = window.firebase?.auth?.();
    if (!firebaseAuth || firebaseAuth.__heraAccessRequestCredentialListener) return;
    firebaseAuth.onAuthStateChanged((user) => {
      if (user && pendingCredential) void saveCredentialSecurely();
    });
    try {
      Object.defineProperty(firebaseAuth, "__heraAccessRequestCredentialListener", { value: true });
    } catch (_) {}
  }

  function install() {
    ensureUi();
    document.addEventListener("submit", captureCredentials, true);
    installAuthSuccessListener();
    let attempts = 0;
    const timer = window.setInterval(() => {
      attempts += 1;
      const uiReady = ensureUi();
      installAuthSuccessListener();
      if ((uiReady && window.firebase?.auth) || attempts >= 80) window.clearInterval(timer);
    }, 250);
  }

  if (document.readyState === "loading") document.addEventListener("DOMContentLoaded", install, { once: true });
  else install();

  window.HeraAccessRequestLogin = { installed: true, open: openDialog };
})();
