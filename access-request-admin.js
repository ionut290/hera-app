(() => {
  "use strict";

  if (window.HeraAccessRequestAdmin?.installed) return;

  const ENDPOINT = "https://europe-west1-hera-app-6cd2b.cloudfunctions.net/registerTester";
  const BUTTON_ID = "access-request-admin-btn";
  const PANEL_ID = "access-request-admin-panel";
  const ADMIN_EMAIL_FALLBACK = "ionut29019@gmail.com";
  let busy = false;

  function text(value) {
    return String(value ?? "").trim();
  }

  function normalizeEmail(value) {
    return text(value).toLowerCase();
  }

  function auth() {
    try { return window.firebase?.auth?.() || null; } catch (_) { return null; }
  }

  async function currentUserIsAdmin() {
    const user = auth()?.currentUser;
    if (!user) return false;
    if (normalizeEmail(user.email) === ADMIN_EMAIL_FALLBACK) return true;
    try {
      const snap = await firebase.firestore().collection("platformUsers").doc(user.uid).get();
      const profile = snap.exists ? (snap.data() || {}) : {};
      return profile.isAdmin === true || profile.admin === true || ["admin", "administrator", "amministratore"].includes(text(profile.role || profile.ruolo).toLowerCase());
    } catch (_) {
      return false;
    }
  }

  async function callEndpoint(payload) {
    const user = auth()?.currentUser;
    if (!user) throw new Error("Sessione amministratore non disponibile.");
    const token = await user.getIdToken(true);
    let response;
    try {
      response = await fetch(ENDPOINT, {
        method: "POST",
        headers: { "Content-Type": "application/json", Authorization: `Bearer ${token}` },
        body: JSON.stringify({ data: payload })
      });
    } catch (error) {
      const friendly = new Error("Connessione al servizio accessi non disponibile. Controlla internet e riprova.");
      friendly.code = "network-error";
      friendly.cause = error;
      throw friendly;
    }
    const body = await response.json().catch(() => ({}));
    if (!response.ok || body.error) throw new Error(body?.error?.message || `Operazione non riuscita (${response.status}).`);
    return body.result || body.data || body;
  }

  function ensurePanel() {
    let panel = document.getElementById(PANEL_ID);
    if (panel) return panel;
    panel = document.createElement("section");
    panel.id = PANEL_ID;
    panel.className = "card hidden";
    panel.innerHTML = `
      <h2>Richieste accesso</h2>
      <p class="muted">Qui trovi gli operatori che non riescono ad accedere. Puoi accettare o rifiutare. Se accetti, l’account viene preparato e ti vengono mostrati username e password temporanea da consegnare all’operatore.</p>
      <div class="actions-row" style="display:flex;gap:8px;flex-wrap:wrap">
        <button id="access-request-admin-refresh" class="btn btn-primary" type="button">AGGIORNA RICHIESTE</button>
        <button id="access-request-admin-close" class="btn" type="button">CHIUDI</button>
      </div>
      <p id="access-request-admin-feedback" class="muted" role="status"></p>
      <div id="access-request-admin-list"></div>
      <div id="access-request-admin-credentials" class="hidden" style="margin-top:12px;padding:12px;border:1px solid rgba(127,127,127,.35);border-radius:10px">
        <strong>Credenziali da consegnare all’operatore</strong>
        <p id="access-request-admin-username" style="font-family:monospace;overflow-wrap:anywhere"></p>
        <p id="access-request-admin-password" style="font-family:monospace;overflow-wrap:anywhere"></p>
        <button id="access-request-admin-copy" class="btn" type="button">COPIA CREDENZIALI</button>
      </div>
    `;
    (document.getElementById("home-page") || document.querySelector("main") || document.body).appendChild(panel);
    panel.querySelector("#access-request-admin-refresh")?.addEventListener("click", refreshRequests);
    panel.querySelector("#access-request-admin-close")?.addEventListener("click", () => panel.classList.add("hidden"));
    panel.querySelector("#access-request-admin-copy")?.addEventListener("click", copyCredentials);
    return panel;
  }

  async function ensureMenuButton() {
    if (document.getElementById(BUTTON_ID)) return;
    if (!(await currentUserIsAdmin())) return;
    const menu = document.querySelector(".side-menu-body") || document.getElementById("side-menu");
    if (!menu) return;
    const button = document.createElement("button");
    button.id = BUTTON_ID;
    button.type = "button";
    button.className = "btn menu-title-btn";
    button.innerHTML = `📨 Richieste accesso <span id="access-request-admin-badge" class="pending-users-badge hidden">0</span>`;
    button.addEventListener("click", async () => {
      const panel = ensurePanel();
      panel.classList.remove("hidden");
      await refreshRequests();
      panel.scrollIntoView({ behavior: "smooth", block: "start" });
    });
    const targetSection = document.getElementById("open-panel-utenti")?.parentElement || menu;
    targetSection.appendChild(button);
    void refreshBadge();
  }

  function renderRequests(requests) {
    const panel = ensurePanel();
    const list = panel.querySelector("#access-request-admin-list");
    if (!list) return;
    list.innerHTML = "";
    if (!requests.length) {
      list.innerHTML = `<p class="muted">Nessuna richiesta in attesa.</p>`;
      return;
    }
    for (const request of requests) {
      const card = document.createElement("div");
      card.style.cssText = "margin-top:10px;padding:12px;border:1px solid rgba(127,127,127,.35);border-radius:10px";
      card.innerHTML = `
        <strong>${text(request.displayName || `${request.firstName || ""} ${request.lastName || ""}`)}</strong>
        <p class="muted" style="margin:.35rem 0">${request.personnelId ? "Già presente nel personale" : "Non trovato nel personale"}</p>
        <div style="display:flex;gap:8px;flex-wrap:wrap">
          <button class="btn btn-primary" data-access-action="approve" data-request-id="${text(request.id)}" type="button">ACCETTA</button>
          <button class="btn" data-access-action="reject" data-request-id="${text(request.id)}" type="button">RIFIUTA</button>
        </div>
      `;
      list.appendChild(card);
    }
    list.querySelectorAll("[data-access-action]").forEach((button) => button.addEventListener("click", handleAction));
  }

  async function refreshRequests() {
    if (busy || !(await currentUserIsAdmin())) return;
    busy = true;
    const panel = ensurePanel();
    const feedback = panel.querySelector("#access-request-admin-feedback");
    if (feedback) feedback.textContent = "Caricamento richieste...";
    try {
      const result = await callEndpoint({ action: "listAccessRequests" });
      const requests = Array.isArray(result?.requests) ? result.requests : [];
      renderRequests(requests);
      updateBadge(requests.length);
      if (feedback) feedback.textContent = requests.length ? `${requests.length} richiesta/e in attesa.` : "Nessuna richiesta in attesa.";
    } catch (error) {
      console.error("Caricamento richieste accesso fallito:", error);
      if (feedback) feedback.textContent = error?.message || "Impossibile caricare le richieste.";
    } finally {
      busy = false;
    }
  }

  async function refreshBadge() {
    if (!(await currentUserIsAdmin())) return;
    try {
      const result = await callEndpoint({ action: "listAccessRequests" });
      updateBadge(Array.isArray(result?.requests) ? result.requests.length : 0);
    } catch (_) {}
  }

  function updateBadge(count) {
    const badge = document.getElementById("access-request-admin-badge");
    if (!badge) return;
    badge.textContent = String(count || 0);
    badge.classList.toggle("hidden", !count);
  }

  async function handleAction(event) {
    const button = event.currentTarget;
    const requestId = text(button?.dataset?.requestId);
    const action = text(button?.dataset?.accessAction);
    if (!requestId || busy) return;
    if (action === "reject" && !window.confirm("Rifiutare questa richiesta di accesso?")) return;
    if (action === "approve" && !window.confirm("Accettare questa richiesta e preparare le credenziali dell’operatore?")) return;

    busy = true;
    const panel = ensurePanel();
    const feedback = panel.querySelector("#access-request-admin-feedback");
    button.disabled = true;
    try {
      if (action === "approve") {
        if (feedback) feedback.textContent = "Preparazione account in corso...";
        const result = await callEndpoint({ action: "approveAccessRequest", requestId });
        const credentials = result?.credentials || {};
        if (!credentials.username || !credentials.temporaryPassword) throw new Error("Credenziali operatore non restituite dal server.");
        panel.querySelector("#access-request-admin-username").textContent = `Username: ${credentials.username}`;
        panel.querySelector("#access-request-admin-password").textContent = `Password: ${credentials.temporaryPassword}`;
        panel.querySelector("#access-request-admin-credentials")?.classList.remove("hidden");
        if (feedback) feedback.textContent = "✅ Richiesta accettata. Copia e consegna le credenziali all’operatore.";
      } else {
        if (feedback) feedback.textContent = "Rifiuto richiesta...";
        await callEndpoint({ action: "rejectAccessRequest", requestId });
        if (feedback) feedback.textContent = "Richiesta rifiutata.";
      }
      busy = false;
      await refreshRequests();
    } catch (error) {
      console.error("Gestione richiesta accesso fallita:", error);
      if (feedback) feedback.textContent = error?.message || "Operazione non riuscita.";
      busy = false;
    } finally {
      button.disabled = false;
    }
  }

  async function copyCredentials() {
    const panel = ensurePanel();
    const username = text(panel.querySelector("#access-request-admin-username")?.textContent);
    const password = text(panel.querySelector("#access-request-admin-password")?.textContent);
    const value = `${username}\n${password}`.trim();
    if (!value) return;
    try { await navigator.clipboard.writeText(value); }
    catch (_) { window.prompt("Copia le credenziali:", value); }
    const feedback = panel.querySelector("#access-request-admin-feedback");
    if (feedback) feedback.textContent = "✅ Credenziali copiate.";
  }

  async function install() {
    await ensureMenuButton();
    const firebaseAuth = auth();
    firebaseAuth?.onAuthStateChanged(async () => {
      document.getElementById(BUTTON_ID)?.remove();
      await ensureMenuButton();
    });
  }

  if (document.readyState === "loading") document.addEventListener("DOMContentLoaded", install, { once: true });
  else void install();

  window.HeraAccessRequestAdmin = { installed: true, refresh: refreshRequests };
})();