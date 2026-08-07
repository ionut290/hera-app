(function installOperatorAccountAdmin() {
  "use strict";

  const BUTTON_ID = "operator-account-admin-btn";
  const PANEL_ID = "operator-account-admin-panel";
  const ADMIN_EMAIL_FALLBACK = "ionut29019@gmail.com";
  const CALLABLE_URL = "https://us-central1-hera-app-6cd2b.cloudfunctions.net/createTesterAccounts";
  let busy = false;

  function normalizeEmail(value) {
    return String(value || "").trim().toLowerCase();
  }

  function getAuth() {
    if (!window.firebase || typeof firebase.auth !== "function") return null;
    try {
      return firebase.auth();
    } catch (_error) {
      return null;
    }
  }

  async function currentUserIsAdmin() {
    const user = getAuth()?.currentUser;
    if (!user) return false;
    const email = normalizeEmail(user.email);
    if (email === ADMIN_EMAIL_FALLBACK) return true;
    if (!window.firebase || typeof firebase.firestore !== "function") return false;

    try {
      const snapshot = await firebase.firestore().collection("platformUsers").doc(user.uid).get();
      if (!snapshot.exists) return false;
      const profile = snapshot.data() || {};
      return profile.isAdmin === true || profile.admin === true || String(profile.role || profile.ruolo || "").toLowerCase() === "admin";
    } catch (error) {
      console.warn("Controllo admin gestore account non riuscito:", error);
      return false;
    }
  }

  function createPanel() {
    let panel = document.getElementById(PANEL_ID);
    if (panel) return panel;

    panel = document.createElement("section");
    panel.id = PANEL_ID;
    panel.className = "card hidden";
    panel.setAttribute("aria-live", "polite");
    panel.innerHTML = `
      <h2>Account operatori</h2>
      <p class="muted">Crea un accesso email/password oppure reimposta la password di un account esistente. Gli ID e i dati dell’operatore non vengono cancellati.</p>
      <form id="operator-account-admin-form">
        <label for="operator-account-email">Email di accesso</label>
        <input id="operator-account-email" type="email" autocomplete="off" placeholder="operatore01@example.com" required>
        <div class="actions-row" style="margin-top:10px;display:flex;flex-wrap:wrap;gap:8px">
          <button id="operator-account-create-btn" class="btn btn-primary" type="submit">CREA / PREPARA ACCOUNT</button>
          <button id="operator-account-close-btn" class="btn" type="button">CHIUDI</button>
        </div>
        <p id="operator-account-feedback" class="muted" role="status"></p>
        <div id="operator-account-result" class="hidden" style="margin-top:12px;padding:12px;border:1px solid rgba(127,127,127,.35);border-radius:10px">
          <strong>Credenziali temporanee</strong>
          <p id="operator-account-result-email" style="overflow-wrap:anywhere"></p>
          <p id="operator-account-result-password" style="font-family:monospace;font-size:1.05rem;overflow-wrap:anywhere"></p>
          <p class="muted">La password temporanea viene mostrata solo ora. L’operatore dovrà cambiarla al primo accesso.</p>
          <button id="operator-account-copy-btn" class="btn" type="button">COPIA CREDENZIALI</button>
        </div>
      </form>
    `;

    const home = document.getElementById("home-page") || document.querySelector("main") || document.body;
    home.appendChild(panel);
    panel.querySelector("#operator-account-close-btn")?.addEventListener("click", () => panel.classList.add("hidden"));
    panel.querySelector("#operator-account-admin-form")?.addEventListener("submit", handleProvision);
    panel.querySelector("#operator-account-copy-btn")?.addEventListener("click", copyCredentials);
    return panel;
  }

  async function createMenuButton() {
    if (document.getElementById(BUTTON_ID)) return;
    if (!(await currentUserIsAdmin())) return;

    const menu = document.querySelector(".side-menu-body") || document.getElementById("side-menu");
    if (!menu) return;

    const button = document.createElement("button");
    button.id = BUTTON_ID;
    button.type = "button";
    button.className = "btn menu-title-btn";
    button.textContent = "🔐 Account operatori";
    button.addEventListener("click", openPanel);
    menu.appendChild(button);
  }

  async function openPanel() {
    if (!(await currentUserIsAdmin())) {
      alert("Questa funzione è disponibile solo all’amministratore.");
      return;
    }
    const panel = createPanel();
    panel.classList.remove("hidden");
    panel.querySelector("#operator-account-email")?.focus();
    panel.scrollIntoView({ behavior: "smooth", block: "start" });
  }

  async function callCreateTesterAccounts(email) {
    const user = getAuth()?.currentUser;
    if (!user) throw new Error("Sessione non disponibile. Esci e rientra nell’app.");
    const token = await user.getIdToken(true);
    const response = await fetch(CALLABLE_URL, {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        "Authorization": `Bearer ${token}`
      },
      body: JSON.stringify({ data: { emails: [email] } })
    });
    const payload = await response.json().catch(() => ({}));
    if (!response.ok || payload.error) {
      const message = payload?.error?.message || `Operazione non riuscita (${response.status}).`;
      throw new Error(message);
    }
    return payload.result || payload.data || payload;
  }

  async function handleProvision(event) {
    event.preventDefault();
    if (busy) return;
    const panel = createPanel();
    const emailInput = panel.querySelector("#operator-account-email");
    const feedback = panel.querySelector("#operator-account-feedback");
    const submit = panel.querySelector("#operator-account-create-btn");
    const resultBox = panel.querySelector("#operator-account-result");
    const email = normalizeEmail(emailInput?.value);

    resultBox?.classList.add("hidden");
    if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) {
      if (feedback) feedback.textContent = "Inserisci un indirizzo email valido.";
      emailInput?.focus();
      return;
    }
    if (!(await currentUserIsAdmin())) {
      if (feedback) feedback.textContent = "Permesso negato: funzione riservata all’amministratore.";
      return;
    }

    const confirmed = window.confirm(
      `Preparare l’accesso per ${email}?\n\nSe l’account esiste già, verrà mantenuto lo stesso UID e verrà soltanto impostata una nuova password temporanea.`
    );
    if (!confirmed) return;

    busy = true;
    if (submit) submit.disabled = true;
    if (feedback) feedback.textContent = "Preparazione account in corso...";
    try {
      const result = await callCreateTesterAccounts(email);
      const credential = Array.isArray(result?.credentials) ? result.credentials[0] : null;
      if (!credential?.temporaryPassword) throw new Error("Il server non ha restituito la password temporanea.");

      panel.querySelector("#operator-account-result-email").textContent = `Email: ${credential.email || email}`;
      panel.querySelector("#operator-account-result-password").textContent = `Password temporanea: ${credential.temporaryPassword}`;
      resultBox?.classList.remove("hidden");
      if (feedback) {
        feedback.textContent = credential.created
          ? "✅ Account creato. Nessun dato esistente è stato eliminato."
          : "✅ Account esistente mantenuto. Password temporanea reimpostata sullo stesso account.";
      }
    } catch (error) {
      console.error("Gestione account operatore fallita:", error);
      if (feedback) feedback.textContent = error?.message || "Operazione non riuscita.";
    } finally {
      busy = false;
      if (submit) submit.disabled = false;
    }
  }

  async function copyCredentials() {
    const panel = createPanel();
    const email = panel.querySelector("#operator-account-result-email")?.textContent || "";
    const password = panel.querySelector("#operator-account-result-password")?.textContent || "";
    const text = `${email}\n${password}`.trim();
    if (!text) return;
    try {
      await navigator.clipboard.writeText(text);
      const feedback = panel.querySelector("#operator-account-feedback");
      if (feedback) feedback.textContent = "✅ Credenziali copiate. Consegnale solo all’operatore interessato.";
    } catch (_error) {
      window.prompt("Copia queste credenziali:", text);
    }
  }

  async function install() {
    await createMenuButton();
    const auth = getAuth();
    if (auth) {
      auth.onAuthStateChanged(async () => {
        const existing = document.getElementById(BUTTON_ID);
        if (existing) existing.remove();
        await createMenuButton();
      });
    }
    window.__heraOperatorAccountAdminInstalled = true;
  }

  if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", install, { once: true });
  } else {
    install();
  }
})();
