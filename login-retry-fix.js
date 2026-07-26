(function installLoginRetryFix() {
  "use strict";

  let registrationPending = false;

  function friendlyLoginError(error) {
    const code = String(error?.code || "").toLowerCase();
    const message = String(error?.message || error || "");
    if (
      ["auth/invalid-credential", "auth/wrong-password", "auth/user-not-found"].includes(code)
      || /INVALID_LOGIN_CREDENTIALS|INVALID_PASSWORD|EMAIL_NOT_FOUND/i.test(message)
    ) {
      return "Email o password non corretta. Controlla i dati e riprova.";
    }
    if (code === "auth/too-many-requests") {
      return "Troppi tentativi. Attendi qualche minuto e riprova.";
    }
    if (code === "auth/network-request-failed") {
      return "Connessione non disponibile. Controlla internet e riprova.";
    }
    return "Accesso non riuscito. Controlla email e password e riprova.";
  }

  function registrationElements() {
    return {
      dialog: document.getElementById("registration-dialog"),
      form: document.getElementById("registration-form"),
      email: document.getElementById("registration-email"),
      firstName: document.getElementById("registration-first-name"),
      lastName: document.getElementById("registration-last-name"),
      password: document.getElementById("registration-password"),
      passwordConfirm: document.getElementById("registration-password-confirm"),
      feedback: document.getElementById("registration-feedback"),
      submit: document.getElementById("registration-submit-btn"),
      cancel: document.getElementById("registration-cancel-btn")
    };
  }

  function closeRegistrationDialog() {
    const { dialog, form } = registrationElements();
    if (dialog?.open) dialog.close();
    form?.reset();
    registrationPending = false;
  }

  function openRegistrationDialog(email, password) {
    const elements = registrationElements();
    if (!elements.dialog || !elements.form) {
      return Promise.reject(new Error("Finestra di registrazione non disponibile."));
    }

    elements.form.reset();
    elements.email.value = email;
    elements.password.value = password;
    elements.feedback.textContent = "";
    if (!elements.dialog.open) elements.dialog.showModal();
    window.setTimeout(() => elements.firstName?.focus(), 0);

    return new Promise((resolve, reject) => {
      registrationPending = true;

      const cleanup = () => {
        elements.form.removeEventListener("submit", submit);
        elements.cancel.removeEventListener("click", cancel);
        elements.dialog.removeEventListener("cancel", preventDialogCancel);
      };

      const cancel = () => {
        cleanup();
        closeRegistrationDialog();
        reject(new Error("Creazione account annullata."));
      };

      const preventDialogCancel = (event) => {
        event.preventDefault();
        cancel();
      };

      const submit = async (event) => {
        event.preventDefault();
        if (!registrationPending) return;

        const firstName = String(elements.firstName.value || "").trim();
        const lastName = String(elements.lastName.value || "").trim();
        const chosenPassword = String(elements.password.value || "");
        const confirmation = String(elements.passwordConfirm.value || "");

        if (!firstName || !lastName) {
          elements.feedback.textContent = "Inserisci nome e cognome.";
          return;
        }
        if (!chosenPassword) {
          elements.feedback.textContent = "Inserisci la password.";
          return;
        }
        if (chosenPassword !== confirmation) {
          elements.feedback.textContent = "Le due password non coincidono.";
          return;
        }

        registrationPending = false;
        elements.submit.disabled = true;
        elements.feedback.textContent = "Creazione account in corso...";

        try {
          const registerTester = firebase.app().functions("europe-west1").httpsCallable("registerTester");
          await registerTester({
            email,
            temporaryPassword: chosenPassword,
            firstName,
            lastName
          });
          cleanup();
          closeRegistrationDialog();
          resolve({ password: chosenPassword });
        } catch (error) {
          registrationPending = true;
          const code = String(error?.code || "").toLowerCase();
          if (code.includes("already-exists")) {
            elements.feedback.textContent = "Questa email ha già un account. Torna al login e controlla la password.";
          } else if (code.includes("permission-denied")) {
            elements.feedback.textContent = "Password di registrazione non valida. Chiedi la password temporanea all’amministratore.";
          } else {
            elements.feedback.textContent = error?.message || "Creazione account non riuscita.";
          }
        } finally {
          elements.submit.disabled = false;
        }
      };

      elements.form.addEventListener("submit", submit);
      elements.cancel.addEventListener("click", cancel);
      elements.dialog.addEventListener("cancel", preventDialogCancel);
    });
  }

  async function startRegistrationFromLogin() {
    const emailInput = document.getElementById("auth-email-input");
    const passwordInput = document.getElementById("auth-password-input");
    const createButton = document.getElementById("auth-create-account-btn");
    const feedback = document.getElementById("auth-email-feedback");
    const email = String(emailInput?.value || "").trim().toLowerCase();
    const password = String(passwordInput?.value || "");

    if (!email || !/^[^\\s@]+@[^\\s@]+\\.[^\\s@]+$/.test(email)) {
      if (feedback) feedback.textContent = "Inserisci prima un indirizzo email valido.";
      emailInput?.focus();
      return;
    }

    if (createButton) createButton.disabled = true;
    if (feedback) feedback.textContent = "Compila i dati per creare il nuovo account.";

    try {
      const auth = firebase.auth();
      await auth.setPersistence(firebase.auth.Auth.Persistence.LOCAL);
      const registration = await openRegistrationDialog(email, password);
      await auth.signInWithEmailAndPassword(email, registration.password);
      if (passwordInput) passwordInput.value = "";
      if (feedback) feedback.textContent = "Account creato. Accesso completato.";
    } catch (error) {
      if (feedback) {
        feedback.textContent = error?.message === "Creazione account annullata."
          ? "Creazione account annullata."
          : friendlyLoginError(error);
      }
    } finally {
      if (createButton) createButton.disabled = false;
    }
  }

  async function handleLogin(event) {
    const form = event.target;
    if (!(form instanceof HTMLFormElement) || form.id !== "auth-email-form") return;
    event.preventDefault();
    event.stopPropagation();
    event.stopImmediatePropagation();

    const emailInput = document.getElementById("auth-email-input");
    const passwordInput = document.getElementById("auth-password-input");
    const loginButton = document.getElementById("auth-email-login-btn");
    const feedback = document.getElementById("auth-email-feedback");
    const email = String(emailInput?.value || "").trim().toLowerCase();
    const password = String(passwordInput?.value || "");

    if (!email || !password) {
      if (feedback) feedback.textContent = "Inserisci email e password.";
      return;
    }

    if (loginButton) {
      loginButton.disabled = true;
      loginButton.textContent = "Accesso...";
    }
    if (feedback) feedback.textContent = "Accesso in corso...";

    try {
      const auth = firebase.auth();
      await auth.setPersistence(firebase.auth.Auth.Persistence.LOCAL);
      let loginPassword = password;
      try {
        await auth.signInWithEmailAndPassword(email, loginPassword);
      } catch (loginError) {
        const code = String(loginError?.code || "").toLowerCase();
        if (!["auth/invalid-credential", "auth/user-not-found"].includes(code)) throw loginError;
        if (feedback) feedback.textContent = "Account non trovato. Completa la creazione del nuovo account.";
        const registration = await openRegistrationDialog(email, password);
        loginPassword = registration.password;
        await auth.signInWithEmailAndPassword(email, loginPassword);
      }
      if (passwordInput) passwordInput.value = "";
      if (feedback) feedback.textContent = "Login completato.";
    } catch (error) {
      console.error("Errore login email/password:", error);
      if (feedback) {
        feedback.textContent = error?.message === "Creazione account annullata."
          ? "Creazione account annullata. Puoi riprovare quando vuoi."
          : friendlyLoginError(error);
      }
      if (passwordInput) {
        passwordInput.value = "";
        passwordInput.focus({ preventScroll: true });
      }
      window.scrollTo({ left: 0, top: 0, behavior: "auto" });
    } finally {
      if (loginButton) {
        loginButton.disabled = false;
        loginButton.textContent = "Entra";
      }
    }
  }

  function initialize() {
    document.addEventListener("submit", handleLogin, true);
    document.getElementById("auth-create-account-btn")
      ?.addEventListener("click", startRegistrationFromLogin);
  }

  if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", initialize, { once: true });
  } else {
    initialize();
  }
})();
