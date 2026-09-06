(function installSharedPasswordRecoveryCode() {
  "use strict";

  const CODE_PREFIX = "REC-";
  const MIN_CODE_LENGTH = 16;
  const MIN_PASSWORD_LENGTH = 10;
  const DIALOG_ID = "password-recovery-code-dialog";
  let challenge = null;
  let completing = false;

  function firebaseApp() {
    if (!window.firebase || !Array.isArray(firebase.apps) || !firebase.apps.length) return null;
    return firebase.apps.find((app) => app.name === "vargaGestionale") || firebase.app();
  }

  function auth() {
    try { return firebaseApp()?.auth?.() || null; } catch (_) { return null; }
  }

  function callable(name) {
    const app = firebaseApp();
    if (!app || typeof app.functions !== "function") throw new Error("Servizio recupero password non disponibile.");
    return app.functions("europe-west1").httpsCallable(name);
  }

  function isRecoveryCodeCandidate(value) {
    const code = String(value || "").trim();
    return code.startsWith(CODE_PREFIX) && code.length >= MIN_CODE_LENGTH && !/\s/.test(code);
  }

  function friendlyError(error) {
    const code = String(error?.code || "").replace(/^functions\//, "");
    if (code === "resource-exhausted") return "Troppi tentativi. Attendi 30 minuti e riprova.";
    if (code === "unavailable") return "Connessione non disponibile. Controlla Internet e riprova.";
    if (code === "permission-denied" || code === "invalid-argument") return error?.message || "Codice di recupero non valido.";
    return error?.message || "Recupero password non riuscito.";
  }

  function addStyle() {
    if (document.getElementById("password-recovery-code-style")) return;
    const style = document.createElement("style");
    style.id = "password-recovery-code-style";
    style.textContent = `
      .password-recovery-code-dialog{border:0;border-radius:18px;padding:0;width:min(520px,calc(100vw - 24px));box-shadow:0 24px 70px rgba(0,0,0,.28)}
      .password-recovery-code-dialog::backdrop{background:rgba(5,24,17,.72)}
      .password-recovery-code-card{padding:22px;display:grid;gap:14px}
      .password-recovery-code-card h2,.password-recovery-code-admin h3{margin:0}
      .password-recovery-code-card label,.password-recovery-code-admin label{display:grid;gap:6px;font-weight:700}
      .password-recovery-code-card input,.password-recovery-code-admin input{min-height:44px;padding:10px 12px;border:1px solid #b8c8bf;border-radius:9px;font:inherit}
      .password-recovery-code-confirmation{display:flex!important;grid-template-columns:auto 1fr!important;align-items:flex-start;gap:10px!important;font-weight:600!important}
      .password-recovery-code-confirmation input{min-height:0;width:20px;height:20px;margin-top:2px;padding:0}
      .password-recovery-code-actions{display:flex;gap:9px;flex-wrap:wrap;justify-content:flex-end}
      .password-recovery-code-feedback{min-height:22px;margin:0}
      .password-recovery-code-admin{border-top:1px solid #d9e4dd;margin-top:20px;padding-top:18px;display:grid;gap:12px}
      .password-recovery-code-admin form{display:grid;gap:10px}
    `;
    document.head.appendChild(style);
  }

  function ensureDialog() {
    let dialog = document.getElementById(DIALOG_ID);
    if (dialog) return dialog;
    dialog = document.createElement("dialog");
    dialog.id = DIALOG_ID;
    dialog.className = "password-recovery-code-dialog";
    dialog.innerHTML = `<form class="password-recovery-code-card" method="dialog">
      <div><small>RECUPERO SICURO</small><h2>Imposta la nuova password</h2></div>
      <p id="password-recovery-code-account" class="muted"></p>
      <label>Nuova password<input id="password-recovery-code-new" type="password" minlength="10" maxlength="128" autocomplete="new-password" required></label>
      <label>Conferma password<input id="password-recovery-code-confirm" type="password" minlength="10" maxlength="128" autocomplete="new-password" required></label>
      <label class="password-recovery-code-confirmation"><input id="password-recovery-code-accept" type="checkbox" required><span>Confermo di sostituire la vecchia password con quella nuova.</span></label>
      <p id="password-recovery-code-feedback" class="password-recovery-code-feedback muted" role="status" aria-live="polite"></p>
      <div class="password-recovery-code-actions"><button id="password-recovery-code-cancel" class="btn ghost" type="button">ANNULLA</button><button id="password-recovery-code-save" class="btn primary" type="submit">CONFERMA E SOSTITUISCI</button></div>
    </form>`;
    document.body.appendChild(dialog);
    dialog.addEventListener("cancel", closeDialog);
    dialog.querySelector("#password-recovery-code-cancel")?.addEventListener("click", closeDialog);
    dialog.querySelector("form")?.addEventListener("submit", completeRecovery);
    return dialog;
  }

  function closeDialog(event) {
    event?.preventDefault?.();
    const dialog = document.getElementById(DIALOG_ID);
    dialog?.querySelector("form")?.reset();
    if (dialog?.open) dialog.close();
    challenge = null;
  }

  async function handleLoginCode({ email, code, feedback } = {}) {
    const normalizedEmail = String(email || "").trim().toLowerCase();
    const recoveryCode = String(code || "").trim();
    if (!isRecoveryCodeCandidate(recoveryCode)) return false;
    if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(normalizedEmail)) {
      throw new Error("Inserisci prima un indirizzo email valido.");
    }
    if (feedback) feedback.textContent = "Verifica del codice di recupero...";
    try {
      const response = await callable("startPasswordRecoveryWithCode")({ email: normalizedEmail, code: recoveryCode });
      if (!response?.data?.allowed || !response.data.challengeToken) throw new Error("Codice di recupero non valido.");
      challenge = { email: normalizedEmail, token: response.data.challengeToken };
      const dialog = ensureDialog();
      dialog.querySelector("#password-recovery-code-account").textContent = `Account: ${normalizedEmail}`;
      dialog.querySelector("#password-recovery-code-feedback").textContent = "Inserisci due volte la nuova password. Deve contenere almeno 10 caratteri.";
      if (!dialog.open) dialog.showModal();
      window.setTimeout(() => dialog.querySelector("#password-recovery-code-new")?.focus(), 0);
      return true;
    } catch (error) {
      throw new Error(friendlyError(error));
    }
  }

  async function completeRecovery(event) {
    event.preventDefault();
    if (completing || !challenge) return;
    const dialog = ensureDialog();
    const password = String(dialog.querySelector("#password-recovery-code-new")?.value || "");
    const confirmation = String(dialog.querySelector("#password-recovery-code-confirm")?.value || "");
    const accepted = Boolean(dialog.querySelector("#password-recovery-code-accept")?.checked);
    const feedback = dialog.querySelector("#password-recovery-code-feedback");
    const button = dialog.querySelector("#password-recovery-code-save");
    if (password.length < MIN_PASSWORD_LENGTH) return void (feedback.textContent = "La password deve contenere almeno 10 caratteri.");
    if (password !== confirmation) return void (feedback.textContent = "Le due password non coincidono.");
    if (!accepted) return void (feedback.textContent = "Devi confermare la sostituzione della vecchia password.");
    completing = true;
    button.disabled = true;
    feedback.textContent = "Cambio password in corso...";
    const activeChallenge = { ...challenge };
    try {
      await callable("completePasswordRecoveryWithCode")({ challengeToken: activeChallenge.token, newPassword: password });
      try {
        const authInstance = auth();
        if (!authInstance) throw new Error("Accesso Firebase non disponibile.");
        window.HeraLoginCredentialVault?.capturePendingCredential?.({ email: activeChallenge.email, password });
        await authInstance.signInWithEmailAndPassword(activeChallenge.email, password);
        feedback.textContent = "Password cambiata. Accesso in corso...";
      } catch (loginError) {
        console.warn("Password cambiata; accesso automatico non riuscito", loginError);
        feedback.textContent = "Password cambiata correttamente. Chiudi e accedi con la nuova password.";
      }
      window.setTimeout(closeDialog, 1400);
    } catch (error) {
      feedback.textContent = friendlyError(error);
    } finally {
      completing = false;
      button.disabled = false;
    }
  }

  function ensureAdminPanel() {
    const host = document.getElementById("panel-utenti");
    if (!host || document.getElementById("password-recovery-code-admin")) return;
    const section = document.createElement("section");
    section.id = "password-recovery-code-admin";
    section.className = "password-recovery-code-admin";
    section.innerHTML = `<h3>Codice unico recupero password</h3>
      <p class="muted">Funziona in Varga Cantieri e Varga Gestionale. Non viene mai salvato in chiaro. Usa un codice di almeno 16 caratteri che inizi con <strong>REC-</strong>.</p>
      <form id="password-recovery-code-admin-form">
        <label>Nuovo codice<input id="password-recovery-code-admin-value" type="password" minlength="16" maxlength="128" autocomplete="new-password" placeholder="Es. REC-Giardino-8427!" required></label>
        <label>Conferma codice<input id="password-recovery-code-admin-confirm" type="password" minlength="16" maxlength="128" autocomplete="new-password" required></label>
        <div class="actions-row"><button class="btn btn-primary" type="submit">SALVA CODICE UNICO</button></div>
      </form>
      <p id="password-recovery-code-admin-feedback" class="muted" role="status" aria-live="polite">Stato non verificato.</p>`;
    const adminList = host.querySelector("#admin-users-list");
    if (adminList) adminList.insertAdjacentElement("afterend", section);
    else host.prepend(section);
    section.querySelector("form")?.addEventListener("submit", saveAdminCode);
  }

  async function loadAdminStatus() {
    const feedback = document.getElementById("password-recovery-code-admin-feedback");
    if (!feedback) return;
    feedback.textContent = "Verifica configurazione...";
    try {
      const response = await callable("getPasswordRecoveryCodeStatus")({});
      feedback.textContent = response?.data?.configured
        ? `✅ Codice configurato${response.data.updatedByEmail ? ` da ${response.data.updatedByEmail}` : ""}. Per sicurezza non può essere visualizzato.`
        : "Nessun codice configurato.";
    } catch (error) {
      feedback.textContent = friendlyError(error);
    }
  }

  async function saveAdminCode(event) {
    event.preventDefault();
    const form = event.currentTarget;
    const code = String(form.querySelector("#password-recovery-code-admin-value")?.value || "").trim();
    const confirmation = String(form.querySelector("#password-recovery-code-admin-confirm")?.value || "").trim();
    const feedback = form.parentElement.querySelector("#password-recovery-code-admin-feedback");
    const button = form.querySelector("button[type=submit]");
    if (!isRecoveryCodeCandidate(code)) return void (feedback.textContent = "Il codice deve iniziare con REC-, avere almeno 16 caratteri e non contenere spazi.");
    if (code !== confirmation) return void (feedback.textContent = "I due codici non coincidono.");
    button.disabled = true;
    feedback.textContent = "Salvataggio sicuro del codice...";
    try {
      await callable("setPasswordRecoveryCode")({ code });
      form.reset();
      feedback.textContent = "✅ Codice unico salvato. Conservalo in un luogo sicuro: non potrà essere visualizzato dall’app.";
    } catch (error) {
      feedback.textContent = friendlyError(error);
    } finally {
      button.disabled = false;
    }
  }

  function install() {
    addStyle();
    ensureDialog();
    ensureAdminPanel();
    document.getElementById("open-panel-utenti")?.addEventListener("click", () => window.setTimeout(loadAdminStatus, 0));
  }

  window.VargaPasswordRecovery = { isRecoveryCodeCandidate, handleLoginCode, loadAdminStatus };
  if (document.readyState === "loading") document.addEventListener("DOMContentLoaded", install, { once: true });
  else install();
})();
