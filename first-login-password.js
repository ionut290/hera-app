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

  function getPasswordResetContinueUrl() {
    const nativeAndroid = Boolean(
      window.Capacitor
      && typeof window.Capacitor.isNativePlatform === "function"
      && window.Capacitor.isNativePlatform()
    );
    const baseUrl = nativeAndroid
      ? "https://creative-syrniki-dddbae.netlify.app/"
      : window.location.href;
    const continueUrl = new URL(baseUrl);
    continueUrl.search = "";
    continueUrl.hash = "";
    continueUrl.searchParams.set("passwordReset", "complete");
    return continueUrl.toString();
  }

  async function sendPasswordResetInstructions(email) {
    const auth = firebase.auth();
    auth.languageCode = "it";
    try {
      await auth.sendPasswordResetEmail(email, {
        url: getPasswordResetContinueUrl(),
        handleCodeInApp: false
      });
    } catch (error) {
      const code = String(error?.code || "").toLowerCase();
      if (!["auth/unauthorized-continue-uri", "auth/invalid-continue-uri"].includes(code)) {
        throw error;
      }
      console.warn("URL di ritorno reset non autorizzato, invio senza continue URL:", error);
      await auth.sendPasswordResetEmail(email);
    }
  }

  function showPasswordResetReturnNotice() {
    let currentUrl = null;
    try {
      currentUrl = new URL(window.location.href);
    } catch (_error) {
      return;
    }
    if (currentUrl.searchParams.get("passwordReset") !== "complete") return;

    const loginFeedback = document.getElementById("auth-email-feedback");
    if (loginFeedback) {
      loginFeedback.textContent = "Password aggiornata. Inserisci email e nuova password per accedere.";
    }
    currentUrl.searchParams.delete("passwordReset");
    const cleanUrl = `${currentUrl.pathname}${currentUrl.search}${currentUrl.hash}`;
    if (window.history && typeof window.history.replaceState === "function") {
      window.history.replaceState({}, document.title, cleanUrl || "/");
    }
  }

  function maskEmail(email) {
    const [localPart, domain] = String(email || "").split("@");
    if (!localPart || !domain) return email;
    const visibleStart = localPart.slice(0, Math.min(2, localPart.length));
    return `${visibleStart}${"•".repeat(Math.max(3, localPart.length - visibleStart.length))}@${domain}`;
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

  async function requestPasswordReset() {
    const emailInput = document.getElementById("auth-email-input");
    const requestButton = document.getElementById("auth-request-password-btn");
    const loginFeedback = document.getElementById("auth-email-feedback");
    const email = String(emailInput?.value || "").trim().toLowerCase();
    const genericMessage = "Se esiste un account con questa email, riceverai il link per reimpostare la password. Controlla anche Spam.";

    if (passwordResetPending) return;
    if (!email || !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) {
      if (loginFeedback) loginFeedback.textContent = "Inserisci prima un indirizzo email valido.";
      emailInput?.focus();
      return;
    }

    const remainingMs = getPasswordResetRemainingMs();
    if (remainingMs > 0) {
      const remainingMinutes = Math.max(1, Math.ceil(remainingMs / 60000));
      if (loginFeedback) loginFeedback.textContent = `Hai già richiesto un link. Attendi circa ${remainingMinutes} minut${remainingMinutes === 1 ? "o" : "i"} prima di riprovare.`;
      return;
    }

    const confirmed = window.confirm(
      `Vuoi davvero ricevere un’email per cambiare la password dell’account ${maskEmail(email)}?\n\nSe non sei tu a richiederlo, premi Annulla.`
    );
    if (!confirmed) {
      if (loginFeedback) loginFeedback.textContent = "Richiesta annullata. Nessuna email è stata inviata.";
      return;
    }

    passwordResetPending = true;
    const originalButtonText = requestButton?.textContent || "PASSWORD DIMENTICATA?";
    if (requestButton) {
      requestButton.disabled = true;
      requestButton.textContent = "INVIO EMAIL...";
    }
    if (loginFeedback) loginFeedback.textContent = "Invio del link per reimpostare la password...";

    try {
      await sendPasswordResetInstructions(email);
      rememberPasswordResetRequest();
      if (loginFeedback) loginFeedback.textContent = `${genericMessage} Per sicurezza, una nuova richiesta sarà disponibile tra 5 minuti.`;
    } catch (error) {
      console.error("Richiesta reset password fallita:", error);
      const code = String(error?.code || "").toLowerCase();
      if (code === "auth/network-request-failed") {
        if (loginFeedback) loginFeedback.textContent = "Connessione non disponibile. Controlla Internet e riprova.";
      } else if (code === "auth/too-many-requests") {
        if (loginFeedback) loginFeedback.textContent = "Troppe richieste. Attendi qualche minuto e riprova.";
      } else {
        if (loginFeedback) loginFeedback.textContent = genericMessage;
      }
    } finally {
      passwordResetPending = false;
      if (requestButton) {
        requestButton.disabled = false;
        requestButton.textContent = originalButtonText;
      }
    }
  }

  function initialize() {
    const { dialog, form } = elements();
    if (!dialog || !form || !window.firebase || typeof firebase.auth !== "function") return;
    dialog.addEventListener("cancel", (event) => event.preventDefault());
    form.addEventListener("submit", saveNewPassword);
    showPasswordResetReturnNotice();
    document.getElementById("auth-request-password-btn")
      ?.addEventListener("click", requestPasswordReset);
    firebase.auth().onAuthStateChanged(handleAuthenticatedUser);
  }

  if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", initialize, { once: true });
  } else {
    initialize();
  }
})();