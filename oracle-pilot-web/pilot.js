"use strict";

firebase.initializeApp({
  apiKey: "AIzaSyBHZG9D1H5YOT9QUzG-cdSdftlreDJNa_k",
  authDomain: "hera-app-6cd2b.firebaseapp.com",
  projectId: "hera-app-6cd2b",
  appId: "1:645390631375:web:df3659a23812560e4012ba"
});

const auth = firebase.auth();
const ui = Object.fromEntries(["serviceStatus", "identity", "login", "logout", "message", "data", "commesse", "commesseCount", "impianti", "impiantiCount", "mapPoints", "mapCount"].map((id) => [id, document.getElementById(id)]));

function message(text, error = false) {
  ui.message.textContent = text;
  ui.message.classList.toggle("error", error);
}

function escapeHtml(value) {
  return String(value ?? "").replace(/[&<>'"]/g, (char) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", "'": "&#39;", '"': "&quot;" })[char]);
}

async function api(path, token) {
  const response = await fetch(`/api${path}`, { headers: { Authorization: `Bearer ${token}` } });
  const body = await response.json().catch(() => ({}));
  if (!response.ok) throw new Error(body.error || `HTTP ${response.status}`);
  return body;
}

function renderRows(target, items, kind) {
  target.innerHTML = items.map((item) => {
    if (kind === "map") return `<div class="row"><strong>${escapeHtml(item.document_id)}</strong><small>${escapeHtml(item.latitudine)}, ${escapeHtml(item.longitudine)}</small></div>`;
    const title = item.denominazione || item.codice || item.document_id;
    const detail = [item.codice, item.comune, item.indirizzo, item.distretto, item.tipologia_intervento].filter(Boolean).join(" · ");
    return `<div class="row"><strong>${escapeHtml(title)}</strong><small>${escapeHtml(detail || item.document_id)}</small></div>`;
  }).join("") || '<div class="row">Nessun elemento</div>';
}

async function loadPilot(user) {
  message("Verifica autorizzazione e caricamento dati…");
  const token = await user.getIdToken(true);
  const [identity, commesse, impianti, points] = await Promise.all([
    api("/v1/me", token),
    api("/v1/catalogo?collection=commesse&limit=100", token),
    api("/v1/catalogo?collection=impianti&limit=100", token),
    api("/v1/mappa?collection=impianti&limit=100", token)
  ]);
  ui.identity.textContent = `${identity.email} · ${identity.role}`;
  renderRows(ui.commesse, commesse.items, "catalog");
  renderRows(ui.impianti, impianti.items, "catalog");
  renderRows(ui.mapPoints, points.items, "map");
  ui.commesseCount.textContent = commesse.page.returned;
  ui.impiantiCount.textContent = impianti.page.returned;
  ui.mapCount.textContent = points.page.returned;
  ui.data.hidden = false;
  message("Dati caricati dall’API Oracle in sola lettura.");
}

ui.login.addEventListener("click", async () => {
  ui.login.disabled = true;
  try { await auth.signInWithPopup(new firebase.auth.GoogleAuthProvider()); }
  catch (error) { message(`Accesso non riuscito: ${error.message}`, true); }
  finally { ui.login.disabled = false; }
});
ui.logout.addEventListener("click", () => auth.signOut());

auth.onAuthStateChanged(async (user) => {
  ui.login.hidden = Boolean(user);
  ui.logout.hidden = !user;
  if (!user) {
    ui.identity.textContent = "Non autenticato";
    ui.data.hidden = true;
    message("Accedi con l’email autorizzata per caricare i dati pilota.");
    return;
  }
  try { await loadPilot(user); }
  catch (error) { ui.data.hidden = true; message(`Accesso pilota negato o servizio non disponibile: ${error.message}`, true); }
});

fetch("/api/healthz").then((response) => {
  if (!response.ok) throw new Error();
  ui.serviceStatus.textContent = "API Oracle disponibile";
  ui.serviceStatus.className = "badge ok";
}).catch(() => {
  ui.serviceStatus.textContent = "API Oracle non disponibile";
  ui.serviceStatus.className = "badge error";
});
