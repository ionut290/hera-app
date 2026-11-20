(function installFirstLoginPasswordChange() {
  "use strict";

  const MIN_PASSWORD_LENGTH = 10;
  const PASSWORD_RESET_COOLDOWN_MS = 5 * 60 * 1000;
  const PASSWORD_RESET_STORAGE_KEY = "vargaCantieriPasswordResetLastRequest";
  let pendingUser = null;
  let saving = false;
  let passwordResetPending = false;

  function elements() {
    return {
      dialog: document.getElementById("first-password-dialog"),
      form: document.getElementById("first-password-form"),
      password: document.getElementById("first-password-input"),
      confirm: document.getElementById("first-password-confirm-input"),
      feedback: document.getElementById("first-password-feedback"),
      save: document.getElementById("first-password-save-btn")
    };
  }

  function setFeedback(message) {
    const { feedback } = elements();
    if (feedback) feedback.textContent = message;
  }

  function closeDialog() {
    const { dialog, form } = elements();
    if (form) form.reset();
    if (dialog && dialog.open) dialog.close();
    document.documentElement.classList.remove("password-change-required");
    pendingUser = null;
  }

  function openDialog(user) {
    const { dialog, password } = elements();
    if (!dialog) return;
    pendingUser = user;
    document.documentElement.classList.add("password-change-required");
    if (!dialog.open) dialog.showModal();
    window.setTimeout(() => password?.focus(), 0);
  }

  async function requiresPasswordChange(user) {
    if (!user || !window.firebase || typeof firebase.firestore !== "function") return false;
    const snapshot = await firebase.firestore().collection("platformUsers").doc(user.uid).get();
    return snapshot.exists && snapshot.data()?.mustChangePassword === true;
  }

  async function handleAuthenticatedUser(user) {
    if (!user) {
      closeDialog();
      return;
    }
    try {
      if (await requiresPasswordChange(user)) openDialog(user);
      else closeDialog();
    } catch (error) {
      console.error("Controllo cambio password iniziale fallito:", error);
      setFeedback("Impossibile verificare il profilo. Controlla la connessione e riprova.");
    }
  }

  async function clearForcedPasswordFlag(user) {
    await user.reload();
    await user.getIdToken(true);
    await firebase.firestore().collection("platformUsers").doc(user.uid).set({
      mustChangePassword: false,
      passwordChangedAt: firebase.firestore.FieldValue.serverTimestamp(),
      updatedAt: firebase.firestore.FieldValue.serverTimestamp()
    }, { merge: true });
  }

  async function saveNewPassword(event) {
    event.preventDefault();
    if (saving || !pendingUser) return;

    const { password, confirm, save } = elements();
    const nextPassword = String(password?.value || "");
    const repeatedPassword = String(confirm?.value || "");

    if (nextPassword.length < MIN_PASSWORD_LENGTH) {
      setFeedback(`La password deve contenere almeno ${MIN_PASSWORD_LENGTH} caratteri.`);
      return;
    }
    if (nextPassword !== repeatedPassword) {
      setFeedback("Le due password non coincidono.");
      return;
    }

    saving = true;
    if (save) save.disabled = true;
    setFeedback("Salvataggio della nuova password...");
    try {
      await pendingUser.updatePassword(nextPassword);

      try {
        await clearForcedPasswordFlag(pendingUser);
      } catch (profileError) {
        console.error("Password aggiornata, ma aggiornamento profilo Firestore fallito:", profileError);
        if (profileError?.code === "permission-denied" || profileError?.code === "firestore/permission-denied") {
          setFeedback("Password salvata. Aggiornamento del profilo non riuscito: riprova ad accedere.");
        } else {
          setFeedback("Password salvata. Alcuni dati del profilo non sono stati aggiornati.");
        }
      }

      if (!pendingUser.emailVerified) {
        await pendingUser.sendEmailVerification();
        setFeedback("Password salvata. Controlla la tua email e verifica l’indirizzo prima di accedere.");
        window.setTimeout(async () => {
          await firebase.auth().signOut();
          closeDialog();
        }, 2500);
      } else {
        closeDialog();
      }
    } catch (error) {
      console.error("Cambio password iniziale fallito:", error);
      if (error?.code === "auth/requires-recent-login") {
        setFeedback("Sessione scaduta. Esci, accedi di nuovo con la password temporanea e riprova.");
      } else if (error?.code === "auth/weak-password") {
        setFeedback("Scegli una password più sicura.");
      } else {
        setFeedback(error?.message || "Cambio password non riuscito.");
      }
    } finally {
      saving = false;
      if (save) save.disabled = false;
    }
  }

  // The standard Firebase Authentication email handler is intentionally used
  // instead of app-specific continue URLs. The signed one-time link must be
  // opened by the mailbox owner; no password or reset token is processed here.
  async function sendPasswordResetInstructions(email) {
    const authInstance = firebase.auth();
    authInstance.languageCode = "it";
    await authInstance.sendPasswordResetEmail(email);
  }

  function resetDialogElements() {
    return {
      dialog: document.getElementById("auth-password-reset-dialog"),
      form: document.getElementById("auth-password-reset-form"),
      email: document.getElementById("auth-password-reset-email"),
      feedback: document.getElementById("auth-password-reset-feedback"),
      submit: document.getElementById("auth-password-reset-submit")
    };
  }

  function closeResetDialog() {
    if (passwordResetPending) return;
    const { dialog, form } = resetDialogElements();
    if (dialog?.open) dialog.close();
    form?.reset();
  }

  function openResetDialog() {
    const { dialog, email, feedback } = resetDialogElements();
    const loginFeedback = document.getElementById("auth-email-feedback");
    if (!dialog || typeof dialog.showModal !== "function") {
      if (loginFeedback) loginFeedback.textContent =
        "Il recupero password non è disponibile in questa versione. Aggiorna l'app e riprova.";
      return;
    }
    if (passwordResetPending) return;
    if (feedback) feedback.textContent = "";
    const loginEmail = String(document.getElementById("auth-email-input")?.value || "").trim();
    if (email) email.value = loginEmail;
    if (!dialog.open) dialog.showModal();
    window.setTimeout(() => email?.focus(), 0);
  }

  function getPasswordResetRemainingMs() {
    try {
      const lastRequest = Number(window.localStorage.getItem(PASSWORD_RESET_STORAGE_KEY) || 0);
      return Math.max(0, PASSWORD_RESET_COOLDOWN_MS - (Date.now() - lastRequest));
    } catch (_error) {
      return 0;
    }
  }

  function rememberPasswordResetRequest() {
    try {
      window.localStorage.setItem(PASSWORD_RESET_STORAGE_KEY, String(Date.now()));
    } catch (_error) {
      // localStorage non disponibile: il blocco Firebase contro gli abusi resta comunque attivo.
    }
  }

  async function requestPasswordReset(event) {
    event?.preventDefault?.();
    if (passwordResetPending) return;
    const { email: emailInput, feedback, submit: requestButton } = resetDialogElements();
    const email = String(emailInput?.value || "").trim().toLowerCase();
    const genericMessage =
      "Se l'indirizzo è associato a un account con password, riceverai un'email con il link sicuro. " +
      "Apri il link, scegli una nuova password e poi torna qui per accedere. " +
      "Controlla anche Spam. Non comunicare mai il link o il codice a nessuno.";

    if (!email || email.length > 254 ||
        !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) {
      if (feedback) feedback.textContent = "Inserisci un indirizzo email valido.";
      emailInput?.focus();
      return;
    }

    const remainingMs = getPasswordResetRemainingMs();
    if (remainingMs > 0) {
      const remainingMinutes = Math.max(1, Math.ceil(remainingMs / 60000));
      if (feedback) feedback.textContent =
        "Hai già richiesto un link. Attendi circa " + remainingMinutes +
        " minuti prima di riprovare.";
      return;
    }

    passwordResetPending = true;
    if (requestButton) {
      requestButton.disabled = true;
      requestButton.textContent = "Invio in corso...";
    }
    if (feedback) feedback.textContent = "Invio del link di recupero...";

    try {
      // This runs in the existing Firebase Auth session. It never asks for
      // the old password, saves credentials, or creates/reset other accounts.
      await sendPasswordResetInstructions(email);
      rememberPasswordResetRequest();
      if (feedback) feedback.textContent = genericMessage +
        " Per sicurezza puoi richiedere un nuovo link tra 5 minuti.";
    } catch (error) {
      // Do not disclose whether the mailbox belongs to a registered user.
      // User-not-found and invalid-credential errors produce the same text
      // as success. Firebase handles server-side abuse/rate limits.
      const code = String(error?.code || "").toLowerCase();
      if (code === "auth/network-request-failed") {
        if (feedback) feedback.textContent =
          "Connessione non disponibile. Controlla Internet e riprova.";
      } else if (code === "auth/too-many-requests" || code === "auth/quota-exceeded") {
        rememberPasswordResetRequest();
        if (feedback) feedback.textContent =
          "Sono state inviate troppe richieste. Riprova più tardi.";
      } else if (code === "auth/operation-not-allowed") {
        if (feedback) feedback.textContent =
          "Recupero via email temporaneamente non disponibile. Contatta l'assistenza.";
      } else {
        rememberPasswordResetRequest();
        if (feedback) feedback.textContent = genericMessage +
          " Se l'email non arriva, controlla la casella corretta e chiedi assistenza.";
      }
    } finally {
      passwordResetPending = false;
      if (requestButton) {
        requestButton.disabled = false;
        requestButton.textContent = "Invia link di recupero";
      }
    }
  }

  function initialize() {
    const { dialog, form } = elements();
    if (!dialog || !form || !window.firebase || typeof firebase.auth !== "function") return;
    dialog.addEventListener("cancel", (event) => event.preventDefault());
    form.addEventListener("submit", saveNewPassword);
    document.getElementById("auth-request-password-btn")
      ?.addEventListener("click", openResetDialog);
    const reset = resetDialogElements();
    reset.form?.addEventListener("submit", requestPasswordReset);
    document.getElementById("auth-password-reset-close")
      ?.addEventListener("click", closeResetDialog);
    reset.dialog?.addEventListener("cancel", (event) => {
      if (passwordResetPending) event.preventDefault();
    });
    firebase.auth().onAuthStateChanged(handleAuthenticatedUser);
  }

  if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", initialize, { once: true });
  } else {
    initialize();
  }
})();