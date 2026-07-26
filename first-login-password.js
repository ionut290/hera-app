(function installFirstLoginPasswordChange() {
  "use strict";

  const MIN_PASSWORD_LENGTH = 10;
  let pendingUser = null;
  let saving = false;

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
      await firebase.firestore().collection("platformUsers").doc(pendingUser.uid).set({
        mustChangePassword: false,
        passwordChangedAt: firebase.firestore.FieldValue.serverTimestamp(),
        updatedAt: firebase.firestore.FieldValue.serverTimestamp()
      }, { merge: true });
      closeDialog();
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

  async function requestTemporaryPassword() {
    const emailInput = document.getElementById("auth-email-input");
    const requestButton = document.getElementById("auth-request-password-btn");
    const email = String(emailInput?.value || "").trim().toLowerCase();
    const genericMessage = "Se l’indirizzo è autorizzato, riceverai un’email per impostare la password.";

    if (!email || !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) {
      const loginFeedback = document.getElementById("auth-email-feedback");
      if (loginFeedback) loginFeedback.textContent = "Inserisci prima un indirizzo email valido.";
      emailInput?.focus();
      return;
    }

    if (requestButton) requestButton.disabled = true;
    const loginFeedback = document.getElementById("auth-email-feedback");
    if (loginFeedback) loginFeedback.textContent = "Invio delle istruzioni...";
    try {
      await firebase.auth().sendPasswordResetEmail(email);
      if (loginFeedback) loginFeedback.textContent = genericMessage;
    } catch (error) {
      console.error("Richiesta password fallita:", error);
      // Messaggio intenzionalmente generico: non rivela se un indirizzo è registrato.
      if (loginFeedback) loginFeedback.textContent = genericMessage;
    } finally {
      if (requestButton) requestButton.disabled = false;
    }
  }

  function initialize() {
    const { dialog, form } = elements();
    if (!dialog || !form || !window.firebase || typeof firebase.auth !== "function") return;
    dialog.addEventListener("cancel", (event) => event.preventDefault());
    form.addEventListener("submit", saveNewPassword);
    document.getElementById("auth-request-password-btn")
      ?.addEventListener("click", requestTemporaryPassword);
    firebase.auth().onAuthStateChanged(handleAuthenticatedUser);
  }

  if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", initialize, { once: true });
  } else {
    initialize();
  }
})();
