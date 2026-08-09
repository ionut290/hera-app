(function installLoginRetryFix() {
  "use strict";

  const MIN_REGISTRATION_PASSWORD_LENGTH = 10;
  const REMEMBER_LOGIN_KEY = "heraRememberLogin";
  const PERSISTED_SESSION_KEY = "heraPersistedUserSession";
  const PERSISTED_AUTH_WAIT_MS = 2500;
  let registrationPending = false;

  function readRememberLoginPreference() {
    try {
      const saved = localStorage.getItem(REMEMBER_LOGIN_KEY);
      return saved === null ? true : saved === "true";
    } catch (_) {
      return true;
    }
  }

  function saveRememberLoginPreference(remember) {
    try {
      localStorage.setItem(REMEMBER_LOGIN_KEY, remember ? "true" : "false");
    } catch (error) {
      console.warn("Preferenza Ricordami non memorizzabile:", error);
    }
  }

  function readPersistedUserSession() {
    try {
      const raw = localStorage.getItem(PERSISTED_SESSION_KEY);
      if (!raw) return null;
      const session = JSON.parse(raw);
      if (!session || session.banned === true || session.accessApproved === false) return null;
      if (!String(session.uid || "").trim()) return null;
      if (!String(session.email || "").includes("@")) return null;
      return session;
    } catch (error) {
      console.warn("Sessione locale non leggibile:", error);
      return null;
    }
  }

  function normalizeEmail(value) {
    return String(value || "").trim().toLowerCase();
  }

  function isMatchingRememberedUser(user, email) {
    if (!user?.uid) return false;
    if (normalizeEmail(user.email) !== normalizeEmail(email)) return false;
    const session = readPersistedUserSession();
    if (!session) return true;
    return String(session.uid) === String(user.uid)
      && normalizeEmail(session.email) === normalizeEmail(email);
  }

  function waitForPersistedAuthUser(auth, email, timeoutMs = PERSISTED_AUTH_WAIT_MS) {
    if (!auth) return Promise.resolve(null);
    if (isMatchingRememberedUser(auth.currentUser, email)) {
      return Promise.resolve(auth.currentUser);
    }

    return new Promise((resolve) => {
      let settled = false;
      let unsubscribe = null;
      const finish = (user) => {
        if (settled) return;
        settled = true;
        if (typeof unsubscribe === "function") unsubscribe();
        resolve(user || null);
      };
      const timer = window.setTimeout(() => finish(null), timeoutMs);
      unsubscribe = auth.onAuthStateChanged((user) => {
        if (!isMatchingRememberedUser(user, email)) return;
        window.clearTimeout(timer);
        finish(user);
      }, () => {
        window.clearTimeout(timer);
        finish(null);
      });
    });
  }

  async function tryRememberedLogin(auth, email, feedback) {
    if (!readRememberLoginPreference()) return null;
    const rememberedSession = readPersistedUserSession();
    if (rememberedSession && normalizeEmail(rememberedSession.email) !== normalizeEmail(email)) {
      return null;
    }

    const user = await waitForPersistedAuthUser(auth, email);
    if (!user) return null;
    if (feedback) {
      feedback.textContent = navigator.onLine === false
        ? "Accesso con sessione salvata. Modalità offline attiva."
        : "Sessione salvata ripristinata.";
    }
    console.log("LOGIN: sessione Firebase locale riutilizzata", {
      uid: user.uid,
      email: user.email,
      online: navigator.onLine !== false
    });
    return user;
  }

  function installPasswordVisibilityToggle() {
    const passwordInput = document.getElementById("auth-password-input");
    const toggle = document.getElementById("auth-password-toggle-btn");
    if (!passwordInput || !toggle || toggle.dataset.installed === "1") return;
    toggle.dataset.installed = "1";
    toggle.addEventListener("click", () => {
      const show = passwordInput.type === "password";
      passwordInput.type = show ? "text" : "password";
      toggle.textContent = show ? "🙈" : "👁️";
      toggle.setAttribute("aria-label", show ? "Nascondi password" : "Mostra password");
      toggle.setAttribute("aria-pressed", show ? "true" : "false");
      passwordInput.focus({ preventScroll: true });
      passwordInput.setSelectionRange(passwordInput.value.length, passwordInput.value.length);
    });
  }

  function initializeRememberLogin() {
    const checkbox = document.getElementById("auth-remember-login");
    if (!checkbox) return;
    checkbox.checked = readRememberLoginPreference();
    checkbox.addEventListener("change", () => saveRememberLoginPreference(checkbox.checked));
  }

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
      const rememberedSession = readPersistedUserSession();
      if (rememberedSession) {
        return "Rete troppo debole per verificare di nuovo l’account. Se questo dispositivo ha già effettuato l’accesso, riapri l’app: verrà usata automaticamente la sessione salvata.";
      }
      return "Connessione non disponibile. Il primo accesso su questo dispositivo richiede internet.";
    }
    if (code === "auth/email-not-verified") {
      return "Email non ancora verificata. Apri il messaggio ricevuto da Firebase e conferma l’indirizzo.";
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

  async function createSelfRegisteredProfile(user, details) {
    if (!user?.uid || typeof firebase.firestore !== "function") {
      throw new Error("Profilo Firebase non disponibile.");
    }

    const displayName = [details.firstName, details.lastName].filter(Boolean).join(" ");
    await firebase.firestore().collection("platformUsers").doc(user.uid).set({
      uid: user.uid,
      email: details.email,
      displayName,
      nome: details.firstName,
      cognome: details.lastName,
      nomeCompleto: displayName,
      firstName: details.firstName,
      lastName: details.lastName,
      teamId: "",
      role: "user",
      ruolo: "user",
      isAdmin: false,
      admin: false,
      statoAccount: "in_attesa",
      accountStatus: "in_attesa",
      primoAccessoAt: firebase.firestore.FieldValue.serverTimestamp(),
      firstLoginAt: firebase.firestore.FieldValue.serverTimestamp(),
      approvatoAt: null,
      approvedAt: null,
      approvatoDa: null,
      approvedBy: null,
      numeroRichieste: 0,
      requestCount: 0,
      permissions: {},
      mustChangePassword: false,
      selfRegistered: true,
      authProviders: ["password"],
      createdAt: firebase.firestore.FieldValue.serverTimestamp(),
      updatedAt: firebase.firestore.FieldValue.serverTimestamp(),
      lastSeenAt: firebase.firestore.FieldValue.serverTimestamp()
    }, { merge: true });
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
        if (chosenPassword.length < MIN_REGISTRATION_PASSWORD_LENGTH) {
          elements.feedback.textContent = `La password deve contenere almeno ${MIN_REGISTRATION_PASSWORD_LENGTH} caratteri.`;
          return;
        }
        if (chosenPassword !== confirmation) {
          elements.feedback.textContent = "Le due password non coincidono.";
          return;
        }

        registrationPending = false;
        elements.submit.disabled = true;
        elements.feedback.textContent = "Creazione account in corso...";

        let createdUser = null;
        try {
          const auth = firebase.auth();
          const credential = await auth.createUserWithEmailAndPassword(email, chosenPassword);
          createdUser = credential.user;
          await createdUser.updateProfile({
            displayName: [firstName, lastName].filter(Boolean).join(" ")
          });
          await createdUser.sendEmailVerification();
          await createSelfRegisteredProfile(createdUser, { email, firstName, lastName });
          try {
            await auth.signOut();
          } catch (signOutError) {
            console.warn("Uscita dopo registrazione non riuscita:", signOutError);
          }
          cleanup();
          closeRegistrationDialog();
          resolve({ verificationRequired: true });
        } catch (error) {
          registrationPending = true;
          if (createdUser?.uid && firebase.auth().currentUser?.uid === createdUser.uid) {
            try {
              await createdUser.delete();
            } catch (cleanupError) {
              console.error("Pulizia account incompleto fallita:", cleanupError);
              try {
                await firebase.auth().signOut();
              } catch (signOutError) {
                console.error("Uscita dopo registrazione fallita:", signOutError);
              }
            }
          }

          const code = String(error?.code || "").toLowerCase();
          const message = String(error?.message || "");
          if (code.includes("email-already-in-use")) {
            elements.feedback.textContent = "Questa email ha già un account. Torna al login e controlla la password.";
          } else if (code.includes("weak-password")) {
            elements.feedback.textContent = `La password deve contenere almeno ${MIN_REGISTRATION_PASSWORD_LENGTH} caratteri.`;
          } else if (code.includes("operation-not-allowed")) {
            elements.feedback.textContent = "La registrazione con email non è abilitata. Contatta l’amministratore.";
          } else if (code.includes("network-request-failed")) {
            elements.feedback.textContent = "Connessione non disponibile. Controlla internet e riprova.";
          } else {
            elements.feedback.textContent = message && message.toLowerCase() !== "internal"
              ? message
              : "Creazione account non riuscita. Riprova tra poco.";
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

    if (!email || !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) {
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
      if (passwordInput) passwordInput.value = "";
      if (registration.verificationRequired && feedback) {
        feedback.textContent = "Account creato. Controlla la tua email, conferma l’indirizzo e poi accedi.";
      }
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
      const rememberLogin = document.getElementById("auth-remember-login")?.checked !== false;
      saveRememberLoginPreference(rememberLogin);
      await auth.setPersistence(
        rememberLogin
          ? firebase.auth.Auth.Persistence.LOCAL
          : firebase.auth.Auth.Persistence.SESSION
      );

      if (rememberLogin) {
        const rememberedUser = await tryRememberedLogin(auth, email, feedback);
        if (rememberedUser) {
          if (passwordInput) passwordInput.value = "";
          return;
        }
      }

      if (navigator.onLine === false) {
        const offlineError = new Error("Connessione non disponibile.");
        offlineError.code = "auth/network-request-failed";
        throw offlineError;
      }

      try {
        const credential = await auth.signInWithEmailAndPassword(email, password);
        if (credential.user && credential.user.emailVerified === false) {
          await auth.signOut();
          const verificationError = new Error("Email non verificata.");
          verificationError.code = "auth/email-not-verified";
          throw verificationError;
        }
      } catch (loginError) {
        const code = String(loginError?.code || "").toLowerCase();
        if (code === "auth/network-request-failed" && rememberLogin) {
          const rememberedUser = await tryRememberedLogin(auth, email, feedback);
          if (rememberedUser) {
            if (passwordInput) passwordInput.value = "";
            return;
          }
        }
        if (!["auth/invalid-credential", "auth/user-not-found"].includes(code)) throw loginError;
        if (feedback) feedback.textContent = "Account non trovato. Completa la creazione del nuovo account.";
        const registration = await openRegistrationDialog(email, password);
        if (passwordInput) passwordInput.value = "";
        if (registration.verificationRequired && feedback) {
          feedback.textContent = "Account creato. Controlla la tua email, conferma l’indirizzo e poi accedi.";
        }
        return;
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
    installPasswordVisibilityToggle();
    initializeRememberLogin();
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
