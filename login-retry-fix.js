(function installLoginRetryFix() {
  "use strict";

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
    const email = String(emailInput?.value || "").trim();
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
      try {
        await auth.signInWithEmailAndPassword(email, password);
      } catch (loginError) {
        const code = String(loginError?.code || "").toLowerCase();
        if (!["auth/invalid-credential", "auth/user-not-found"].includes(code)) throw loginError;
        const registerTester = firebase.app().functions("europe-west1").httpsCallable("registerTester");
        await registerTester({ email, temporaryPassword: password });
        await auth.signInWithEmailAndPassword(email, password);
      }
      if (passwordInput) passwordInput.value = "";
      if (feedback) feedback.textContent = "Login completato.";
    } catch (error) {
      console.error("Errore login email/password:", error);
      if (feedback) feedback.textContent = friendlyLoginError(error);
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
  }

  if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", initialize, { once: true });
  } else {
    initialize();
  }
})();
