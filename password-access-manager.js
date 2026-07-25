(function installPasswordAccessManager() {
  "use strict";

  const BUTTON_ID = "password-access-manager-btn";
  const PANEL_ID = "password-access-manager-panel";

  function getAuth() {
    return window.firebase && typeof firebase.auth === "function" ? firebase.auth() : null;
  }

  function hasProvider(user, providerId) {
    return Array.isArray(user?.providerData)
      && user.providerData.some((provider) => provider && provider.providerId === providerId);
  }

  function formatError(error) {
    const code = String(error?.code || "");
    if (code === "auth/requires-recent-login") return "Per sicurezza, esci e rientra con Google, poi riprova.";
    if (code === "auth/weak-password") return "La password deve avere almeno 6 caratteri.";
    if (code === "auth/email-already-in-use" || code === "auth/credential-already-in-use") {
      return "Questa email è già collegata a un altro account Firebase. Contatta l’amministratore.";
    }
    if (code === "auth/provider-already-linked") return "L’accesso con password è già attivo.";
    return String(error?.message || "Operazione non riuscita.");
  }

  function createPanel() {
    if (document.getElementById(PANEL_ID)) return document.getElementById(PANEL_ID);

    const panel = document.createElement("section");
    panel.id = PANEL_ID;
    panel.className = "card hidden";
    panel.setAttribute("aria-live", "polite");
    panel.innerHTML = `
      <h2>Accesso con password</h2>
      <p id="password-access-status" class="muted">Verifica account...</p>
      <form id="password-access-form">
        <label for="password-access-input">Nuova password</label>
        <input id="password-access-input" type="password" minlength="6" autocomplete="new-password" placeholder="Almeno 6 caratteri" required>
        <label for="password-access-confirm-input">Conferma password</label>
        <input id="password-access-confirm-input" type="password" minlength="6" autocomplete="new-password" placeholder="Ripeti la password" required>
        <div class="actions-row">
          <button id="password-access-save-btn" class="btn btn-primary" type="submit">ATTIVA ACCESSO CON PASSWORD</button>
          <button id="password-access-close-btn" class="btn" type="button">CHIUDI</button>
        </div>
        <p id="password-access-feedback" class="muted" role="status"></p>
      </form>
    `;

    const home = document.getElementById("home-page") || document.querySelector("main") || document.body;
    home.appendChild(panel);

    panel.querySelector("#password-access-close-btn")?.addEventListener("click", () => panel.classList.add("hidden"));
    panel.querySelector("#password-access-form")?.addEventListener("submit", handleSubmit);
    return panel;
  }

  function createMenuButton() {
    if (document.getElementById(BUTTON_ID)) return;
    const menu = document.querySelector(".side-menu-body") || document.getElementById("side-menu");
    if (!menu) return;

    const button = document.createElement("button");
    button.id = BUTTON_ID;
    button.type = "button";
    button.className = "btn menu-title-btn";
    button.textContent = "🔐 Accesso con password";
    button.addEventListener("click", openPanel);
    menu.appendChild(button);
  }

  function updatePanelState() {
    const user = getAuth()?.currentUser;
    const panel = createPanel();
    const status = panel.querySelector("#password-access-status");
    const form = panel.querySelector("#password-access-form");
    const button = panel.querySelector("#password-access-save-btn");

    if (!user) {
      status.textContent = "Devi essere collegato all’app.";
      form.classList.add("hidden");
      return;
    }

    form.classList.remove("hidden");
    const passwordLinked = hasProvider(user, "password");
    if (passwordLinked) {
      status.textContent = `Accesso con password già attivo per ${user.email || "questo account"}.`;
      button.disabled = true;
      button.textContent = "ACCESSO CON PASSWORD GIÀ ATTIVO";
      return;
    }

    button.disabled = false;
    button.textContent = "ATTIVA ACCESSO CON PASSWORD";
    status.textContent = `Imposta una password per ${user.email || "questo account"}. Lo stesso account e gli stessi dati resteranno invariati.`;
  }

  function openPanel() {
    const panel = createPanel();
    updatePanelState();
    panel.classList.remove("hidden");
    panel.scrollIntoView({ behavior: "smooth", block: "start" });
  }

  async function handleSubmit(event) {
    event.preventDefault();
    const user = getAuth()?.currentUser;
    const panel = createPanel();
    const password = panel.querySelector("#password-access-input")?.value || "";
    const confirmPassword = panel.querySelector("#password-access-confirm-input")?.value || "";
    const feedback = panel.querySelector("#password-access-feedback");
    const saveButton = panel.querySelector("#password-access-save-btn");

    if (!user?.email) {
      feedback.textContent = "Account non disponibile. Esci e rientra con Google.";
      return;
    }
    if (password.length < 6) {
      feedback.textContent = "La password deve avere almeno 6 caratteri.";
      return;
    }
    if (password !== confirmPassword) {
      feedback.textContent = "Le due password non coincidono.";
      return;
    }

    saveButton.disabled = true;
    feedback.textContent = "Attivazione in corso...";

    try {
      const credential = firebase.auth.EmailAuthProvider.credential(user.email, password);
      await user.linkWithCredential(credential);
      await user.reload();
      feedback.textContent = "✅ Accesso con password attivato. Ora puoi entrare anche con email e password mantenendo lo stesso account.";
      panel.querySelector("#password-access-input").value = "";
      panel.querySelector("#password-access-confirm-input").value = "";
      updatePanelState();
    } catch (error) {
      console.error("Errore collegamento accesso con password:", error);
      feedback.textContent = formatError(error);
      saveButton.disabled = false;
    }
  }

  function install() {
    createMenuButton();
    const auth = getAuth();
    if (auth) auth.onAuthStateChanged(() => updatePanelState());
  }

  if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", install, { once: true });
  } else {
    install();
  }
})();
