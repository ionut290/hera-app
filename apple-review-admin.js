(() => {
  "use strict";
  const $ = (id) => document.getElementById(id);
  const status = $("review-admin-status");
  const form = $("review-admin-create");
  const result = $("review-admin-result");
  const createButton = $("create-review-btn");
  const emailInput = $("new-review-email");
  const userField = $("result-review-email");
  const passField = $("result-review-password");

  try {
    firebase.initializeApp(window.VargaReviewFirebase);
    firebase.auth().onAuthStateChanged((user) => {
      if (!user) {
        form.hidden = true;
        status.textContent = "Non risulti autenticato. Accedi prima a VARGA CANTIERI come amministratore e torna a questa pagina.";
      } else {
        form.hidden = false;
        status.textContent = "Sessione Firebase rilevata. L'autorizzazione amministrativa verrà verificata dal server.";
      }
    });
  } catch (_) {
    status.textContent = "Impossibile collegarsi a Firebase. Riprovare dopo avere eseguito l'accesso all'app.";
  }
  form.addEventListener("submit", async (event) => {
    event.preventDefault();
    if (!firebase.auth().currentUser) {
      status.textContent = "Prima accedi come amministratore.";
      return;
    }
    createButton.disabled = true;
    status.textContent = "Creazione sicura dell'account...";
    try {
      const email = emailInput.value.trim().toLowerCase();
      const fn = firebase.app().functions("europe-west1").httpsCallable("createAppleReviewAccount");
      const response = await fn({ email });
      if (!response?.data?.created || !response.data.password) throw new Error("Risposta incompleta.");
      userField.value = response.data.email;
      passField.value = response.data.password;
      result.hidden = false;
      form.hidden = true;
      status.textContent = "Account Apple Review creato. Salva le credenziali direttamente su App Store Connect.";
    } catch (error) {
      status.textContent = "Impossibile creare l'account: " + (error?.message || "verifica autorizzazioni e disponibilità");
    } finally {
      createButton.disabled = false;
    }
  });
  $("toggle-review-password").addEventListener("click", () => {
    passField.type = passField.type === "password" ? "text" : "password";
  });
  $("copy-review-credentials").addEventListener("click", () => {
    const credentials = "Username: " + userField.value + "\nPassword: " + passField.value;
    if (navigator.clipboard?.writeText) {
      navigator.clipboard.writeText(credentials).then(() => {
        status.textContent = "Credenziali copiate: incollale in App Store Connect, non in chat.";
      }).catch(() => {
        status.textContent = "Usa il pulsante Mostra e copia manualmente i dati.";
      });
    } else {
      status.textContent = "Usa il pulsante Mostra e copia manualmente i dati.";
    }
  });
})();