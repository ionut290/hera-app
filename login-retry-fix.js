(function installLoginRetryFix() {
  "use strict";

  const MIN_REGISTRATION_PASSWORD_LENGTH = 10;
  const LOGIN_STABILITY_CONTRACT = "LOGIN_STABILITY_CONTRACT_V1";
  const POST_LOGIN_VISIBILITY_CHECK_DELAYS_MS = [0, 1500, 4000, 8000];
  let registrationPending = false;

  function isElementPresented(element) {
    return Boolean(element && !element.hidden && !element.classList.contains("hidden"));
  }

  function hasVisibleAuthenticationSurface() {
    const body = document.body;
    const loaderVisible = isElementPresented(document.getElementById("app-startup-loading"));
    const gateVisible = isElementPresented(document.getElementById("auth-gate"));
    const approvalVisible = isElementPresented(document.getElementById("access-approval-screen"));
    const homeVisible = Boolean(
      isElementPresented(document.getElementById("home-page"))
      && body
      && !body.classList.contains("auth-pending")
      && !body.classList.contains("auth-required")
      && !body.classList.contains("auth-banned")
    );
    return loaderVisible || gateVisible || approvalVisible || homeVisible;
  }

  function recoverVisibleAuthenticationSurface() {
    if (hasVisibleAuthenticationSurface()) return true;
    const loader = document.getElementById("app-startup-loading");
    document.body?.classList.add("auth-pending");
    document.body?.classList.remove("auth-required", "auth-banned");
    if (loader) {
      loader.classList.remove("hidden");
      loader.hidden = false;
      loader.removeAttribute("aria-hidden");
    }
    console.warn(`[${LOGIN_STABILITY_CONTRACT}] Recuperata una transizione login senza superficie visibile.`);
    return false;
  }

  function schedulePostLoginVisibilityChecks() {
    POST_LOGIN_VISIBILITY_CHECK_DELAYS_MS.forEach((delay) => {
      window.setTimeout(recoverVisibleAuthenticationSurface, delay);
    });
  }

  function friendlyLoginError(error) {
    const code = String(error?.code || "").toLowerCase();
    const message = String(error?.message || error || "");
    if (
      ["auth/invalid-credential", "auth/wrong-password", "auth/user-not-found"].includes(code)
      || /INVALID_LOGIN_CREDENTIALS|INVALID_PASSWORD|EMAIL_NOT_FOUND/i.test(message)
    ) {
      return "Email o password non corretta. Se l’account è già registrato, usa PASSWORD DIMENTICATA?.";
    }
    if (code === "auth/too-many-requests") return "Troppi tentativi. Attendi qualche minuto e riprova.";
    if (code === "auth/network-request-failed") return "Connessione non disponibile. Controlla internet e riprova.";
    if (code === "auth/email-not-verified") return "Email non ancora verificata. Apri il messaggio ricevuto da Firebase e conferma il tuo indirizzo.";
    return "Accesso non riuscito. Controlla email e password e riprova.";
  }

  async function ensureLocalPersistenceWithoutBlocking(auth) {
    try {
      await auth.setPersistence(firebase.auth.Auth.Persistence.LOCAL);
      return true;
    } catch (error) {
      console.warn("Persistenza Firebase LOCAL non disponibile: continuo comunque con il login.", error);
      return false;
    }
  }

  function completeSuccessfulLoginUi() {
    try {
      // L'autenticazione Firebase e la verifica autorizzazioni sono due fasi
      // distinte. Mantieni visibile lo splash durante la seconda fase: togliere
      // insieme gate e loader mentre `body.auth-required` nasconde la Home
      // produce una schermata completamente bianca.
      document.body?.classList.add("auth-pending");
      document.body?.classList.remove("auth-required", "auth-banned");
      document.documentElement?.classList.add("auth-pending");
      const loader = document.getElementById("app-startup-loading");
      if (loader) {
        loader.classList.remove("hidden");
        loader.hidden = false;
        loader.removeAttribute("aria-hidden");
      }
      const gate = document.getElementById("auth-gate");
      if (gate) {
        gate.hidden = true;
        gate.classList.add("hidden");
        gate.style.setProperty("display", "none", "important");
        gate.setAttribute("aria-hidden", "true");
      }
      window.HeraAuthStartupController?.reconcile?.();
      window.dispatchEvent(new CustomEvent("hera-login-ui-ready", { detail: { at: Date.now() } }));
      schedulePostLoginVisibilityChecks();
    } catch (error) {
      console.warn("Ripristino UI dopo login non riuscito:", error);
    }
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
    if (!user?.uid || typeof firebase.firestore !== "function") throw new Error("Profilo Firebase non disponibile.");
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
    if (!elements.dialog || !elements.form) return Promise.reject(new Error("Finestra di registrazione non disponibile."));

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
          await createdUser.updateProfile({ displayName: [firstName, lastName].filter(Boolean).join(" ") });
          await createdUser.sendEmailVerification();
          await createSelfRegisteredProfile(createdUser, { email, firstName, lastName });
          try { await auth.signOut(); } catch (signOutError) { console.warn("Uscita dopo registrazione non riuscita:", signOutError); }
          cleanup();
          closeRegistrationDialog();
          resolve({ verificationRequired: true });
        } catch (error) {
          registrationPending = true;
          if (createdUser?.uid && firebase.auth().currentUser?.uid === createdUser.uid) {
            try { await createdUser.delete(); }
            catch (cleanupError) {
              console.error("Pulizia account incompleto fallita:", cleanupError);
              try { await firebase.auth().signOut(); } catch (signOutError) { console.error("Uscita dopo registrazione fallita:", signOutError); }
            }
          }
          const code = String(error?.code || "").toLowerCase();
          const message = String(error?.message || "");
          if (code.includes("email-already-in-use")) elements.feedback.textContent = "Questa email ha già un account. Torna al login e controlla la password.";
          else if (code.includes("weak-password")) elements.feedback.textContent = `La password deve contenere almeno ${MIN_REGISTRATION_PASSWORD_LENGTH} caratteri.`;
          else if (code.includes("operation-not-allowed")) elements.feedback.textContent = "La registrazione con email non è abilitata. Contatta l’amministratore.";
          else if (code.includes("network-request-failed")) elements.feedback.textContent = "Connessione non disponibile. Controlla internet e riprova.";
          else elements.feedback.textContent = message && message.toLowerCase() !== "internal" ? message : "Creazione account non riuscita. Riprova tra poco.";
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
      await ensureLocalPersistenceWithoutBlocking(auth);
      const registration = await openRegistrationDialog(email, password);
      if (passwordInput) passwordInput.value = "";
      if (registration.verificationRequired && feedback) feedback.textContent = "Account creato. Controlla la tua email, conferma l’indirizzo e poi accedi.";
    } catch (error) {
      if (feedback) feedback.textContent = error?.message === "Creazione account annullata." ? "Creazione account annullata." : friendlyLoginError(error);
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
    if (window.VargaPasswordRecovery?.isRecoveryCodeCandidate?.(password)) {
      if (loginButton) {
        loginButton.disabled = true;
        loginButton.textContent = "Verifica codice...";
      }
      try {
        await window.VargaPasswordRecovery.handleLoginCode({ email, code: password, feedback });
        if (passwordInput) passwordInput.value = "";
      } catch (error) {
        if (feedback) feedback.textContent = error?.message || "Codice di recupero non valido.";
      } finally {
        if (loginButton) {
          loginButton.disabled = false;
          loginButton.textContent = "Entra";
        }
      }
      return;
    }
    try {
      window.HeraLoginCredentialVault?.capturePendingCredential?.({ email, password });
    } catch (error) {
      console.warn("Preparazione salvataggio credenziale non riuscita:", error);
    }
    if (loginButton) {
      loginButton.disabled = true;
      loginButton.textContent = "Accesso...";
    }
    if (feedback) feedback.textContent = "Accesso in corso...";

    try {
      const auth = firebase.auth();
      await ensureLocalPersistenceWithoutBlocking(auth);
      await auth.signInWithEmailAndPassword(email, password);
      completeSuccessfulLoginUi();
      if (passwordInput) passwordInput.value = "";
      if (feedback) feedback.textContent = "Login completato.";
    } catch (error) {
      console.error("Errore login email/password:", error);
      if (feedback) feedback.textContent = error?.message === "Creazione account annullata." ? "Creazione account annullata. Puoi riprovare quando vuoi." : friendlyLoginError(error);
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
    document.getElementById("auth-create-account-btn")?.addEventListener("click", startRegistrationFromLogin);
  }

  if (document.readyState === "loading") document.addEventListener("DOMContentLoaded", initialize, { once: true });
  else initialize();
})();

(function installNonBlockingAuthStateBridge() {
  "use strict";

  if (window.__heraNonBlockingAuthStateBridgeInstalled) return;

  function findFirebasePrototypeMethod(instance, methodName) {
    let prototype = Object.getPrototypeOf(instance);
    while (prototype) {
      if (Object.prototype.hasOwnProperty.call(prototype, methodName) && typeof prototype[methodName] === "function") {
        return prototype[methodName];
      }
      prototype = Object.getPrototypeOf(prototype);
    }
    return null;
  }

  function showEmailVerificationRequired() {
    const message = "Prima di accedere ai dati, apri l’email di verifica e conferma il tuo indirizzo.";
    const gateMessage = document.getElementById("auth-gate-message");
    const feedback = document.getElementById("auth-email-feedback");
    if (gateMessage) gateMessage.textContent = message;
    if (feedback) feedback.textContent = message;
  }

  function install() {
    if (!window.firebase || typeof firebase.auth !== "function") return false;

    let auth;
    try {
      auth = firebase.auth();
    } catch (_) {
      return false;
    }
    if (!auth || auth.__heraNonBlockingAuthStateBridgeInstalled) return true;

    const firebaseOnAuthStateChanged = findFirebasePrototypeMethod(auth, "onAuthStateChanged");
    if (typeof firebaseOnAuthStateChanged !== "function") {
      console.warn("[login-retry-fix] Metodo Firebase originale onAuthStateChanged non trovato.");
      return false;
    }

    auth.onAuthStateChanged = function onAuthStateChangedWithoutProfileBlocking(nextOrObserver, error, completed) {
      const wrapCallback = (callback) => (user) => {
        const verificationRequired = Boolean(user?.email && user.emailVerified === false);
        const effectiveUser = verificationRequired ? null : user;

        window.__heraEmailVerificationRequired = verificationRequired;
        let result;
        try {
          result = typeof callback === "function" ? callback(effectiveUser) : undefined;
        } finally {
          if (verificationRequired) queueMicrotask(showEmailVerificationRequired);
        }
        return result;
      };

      if (typeof nextOrObserver === "function") {
        return firebaseOnAuthStateChanged.call(auth, wrapCallback(nextOrObserver), error, completed);
      }

      const observer = nextOrObserver || {};
      return firebaseOnAuthStateChanged.call(auth, {
        next: wrapCallback(observer.next),
        error: observer.error,
        complete: observer.complete
      });
    };

    Object.defineProperty(auth, "__heraNonBlockingAuthStateBridgeInstalled", {
      value: true,
      configurable: false,
      enumerable: false
    });
    window.__heraNonBlockingAuthStateBridgeInstalled = true;
    return true;
  }

  if (!install()) {
    let attempts = 0;
    const timer = window.setInterval(() => {
      attempts += 1;
      if (install() || attempts >= 50) window.clearInterval(timer);
    }, 100);
  }
})();
