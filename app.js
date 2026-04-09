firebase.initializeApp(firebaseConfig);

const auth = firebase.auth();
const db = firebase.firestore();

const ui = {
  loginBtn: document.getElementById("login-btn"),
  logoutBtn: document.getElementById("logout-btn"),
  user: document.getElementById("user"),
  commessaForm: document.getElementById("commessa-form"),
  commessaName: document.getElementById("commessa-name"),
  commesseLista: document.getElementById("commesse-lista"),
  commessaAttiva: document.getElementById("commessa-attiva"),
  excelFile: document.getElementById("excel-file"),
  importBtn: document.getElementById("import-btn"),
  importFeedback: document.getElementById("import-feedback"),
  impiantiLista: document.getElementById("impianti-lista"),
  gpsStatus: document.getElementById("gps-status")
};

let pendingRows = [];
let selectedCommessaId = "";
let selectedCommessaName = "";
let unsubscribeCommesse = null;
let unsubscribeImpianti = null;
let currentUserPos = null;
let currentImpianti = [];
let currentUser = null;
const ADMIN_EMAIL = "ionut29019@gmail.com";

const map = L.map("map");
L.tileLayer("https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png", {
  maxZoom: 19,
  attribution: "&copy; OpenStreetMap"
}).addTo(map);
map.setView([44.4949, 11.3426], 11);

const markerLayer = L.layerGroup().addTo(map);

ui.loginBtn.addEventListener("click", loginWithGoogle);
ui.logoutBtn.addEventListener("click", logout);
ui.commessaForm.addEventListener("submit", createCommessa);
ui.excelFile.addEventListener("change", onExcelSelected);
ui.importBtn.addEventListener("click", importPendingRows);

initGeolocation();

auth.onAuthStateChanged((user) => {
  currentUser = user || null;
  const loggedIn = Boolean(user);

  ui.loginBtn.disabled = loggedIn;
  ui.logoutBtn.disabled = !loggedIn;
  ui.user.textContent = loggedIn
    ? `Loggato: ${user.displayName || "Utente"} (${user.email || "email non disponibile"})`
    : "Non loggato";

  ui.importBtn.disabled = !loggedIn || !selectedCommessaId || pendingRows.length === 0 || !canManageData();
  updateAdminControls();

  stopCommesseSubscription();
  stopImpiantiSubscription();
  selectedCommessaId = "";
  selectedCommessaName = "";
  ui.commesseLista.innerHTML = "";
  ui.impiantiLista.innerHTML = loggedIn
    ? "<p class='muted'>Seleziona una commessa.</p>"
    : "<p class='muted'>Fai login per vedere le commesse.</p>";
  clearMap();

  if (loggedIn) subscribeCommesse();
});

function updateAdminControls() {
  const canManage = canManageData();
  ui.commessaName.disabled = !canManage;
  const submitBtn = ui.commessaForm.querySelector("button[type='submit']");
  if (submitBtn) submitBtn.disabled = !canManage;
}

function loginWithGoogle() {
  const provider = new firebase.auth.GoogleAuthProvider();
  auth.signInWithPopup(provider).catch((error) => {
    if (error.code === "auth/popup-blocked" || error.code === "auth/cancelled-popup-request") {
      return auth.signInWithRedirect(provider);
    }
    alert("Errore login: " + error.message);
  });
}

function logout() {
  auth.signOut();
}

async function createCommessa(event) {
  event.preventDefault();

  const user = auth.currentUser;
  if (!user) {
    alert("Devi fare login.");
    return;
  }

  const nome = ui.commessaName.value.trim();
  if (!nome) return;
  if (!canManageData()) {
    alert("Solo ionut29019@gmail.com può aggiungere commesse.");
    return;
  }

  await db.collection("commesse").add({
    nome,
    creatoDa: user.email || "",
    createdAt: firebase.firestore.FieldValue.serverTimestamp()
  });

  ui.commessaForm.reset();
}

function subscribeCommesse() {
  unsubscribeCommesse = db
    .collection("commesse")
    .orderBy("createdAt", "desc")
    .onSnapshot((snapshot) => {
      ui.commesseLista.innerHTML = "";

      if (snapshot.empty) {
        ui.commesseLista.innerHTML = "<p class='muted'>Nessuna commessa disponibile.</p>";
        return;
      }

      snapshot.forEach((doc) => {
        const commessa = doc.data();
        const row = document.createElement("div");
        row.className = "commessa-row";

        const btn = document.createElement("button");
        btn.type = "button";
        btn.className = "btn commessa-btn" + (doc.id === selectedCommessaId ? " active" : "");
        btn.dataset.commessaId = doc.id;
        btn.textContent = commessa.nome || "Commessa senza nome";
        btn.addEventListener("click", () => selectCommessa(doc.id, commessa.nome || "Commessa"));

        const deleteBtn = createButton("Elimina", () => deleteCommessa(doc.id, commessa.nome || "Commessa"));
        deleteBtn.disabled = !canManageData();

        row.appendChild(btn);
        row.appendChild(deleteBtn);
        ui.commesseLista.appendChild(row);
      });
    }, (error) => {
      console.error(error);
      ui.commesseLista.innerHTML = "<p class='muted'>Errore caricamento commesse.</p>";
    });
}

function stopCommesseSubscription() {
  if (unsubscribeCommesse) {
    unsubscribeCommesse();
    unsubscribeCommesse = null;
  }
}

function selectCommessa(id, nome) {
  selectedCommessaId = id;
  selectedCommessaName = nome;
  ui.commessaAttiva.textContent = `Commessa selezionata: ${nome}`;
  ui.importBtn.disabled = !auth.currentUser || pendingRows.length === 0 || !canManageData();
  updateCommessaButtonsActive();

  stopImpiantiSubscription();
  subscribeImpianti();
}

function updateCommessaButtonsActive() {
  const buttons = ui.commesseLista.querySelectorAll(".commessa-btn");
  buttons.forEach((btn) => {
    const isActive = btn.dataset.commessaId === selectedCommessaId;
    btn.classList.toggle("active", isActive);
  });
}

function subscribeImpianti() {
  if (!selectedCommessaId) return;

  unsubscribeImpianti = db
    .collection("commesse")
    .doc(selectedCommessaId)
    .collection("impianti")
    .onSnapshot((snapshot) => {
      currentImpianti = snapshot.docs.map((doc) => ({ id: doc.id, ...doc.data() }));
      renderImpianti();
      renderMap();
    }, (error) => {
      console.error(error);
      ui.impiantiLista.innerHTML = "<p class='muted'>Errore caricamento impianti.</p>";
    });
}

function stopImpiantiSubscription() {
  if (unsubscribeImpianti) {
    unsubscribeImpianti();
    unsubscribeImpianti = null;
  }
  currentImpianti = [];
  clearMap();
}

function onExcelSelected(event) {
  pendingRows = [];
  ui.importBtn.disabled = true;

  const file = event.target.files && event.target.files[0];
  if (!file) {
    ui.importFeedback.textContent = "Nessun file selezionato.";
    return;
  }

  const reader = new FileReader();
  reader.onload = (e) => {
    try {
      const workbook = XLSX.read(e.target.result, { type: "array" });
      const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
      const rows = XLSX.utils.sheet_to_json(firstSheet, { defval: "", raw: false });

      pendingRows = rows.map(normalizeRow).filter((row) => row.denominazione);
      ui.importFeedback.textContent = `Righe valide trovate: ${pendingRows.length}`;
      ui.importBtn.disabled = !auth.currentUser || !selectedCommessaId || pendingRows.length === 0;
    } catch (error) {
      console.error(error);
      ui.importFeedback.textContent = "Errore lettura Excel.";
    }
  };

  reader.readAsArrayBuffer(file);
}

async function importPendingRows() {
  const user = auth.currentUser;

  if (!user || !selectedCommessaId) {
    alert("Fai login e seleziona una commessa.");
    return;
  }

  if (!pendingRows.length) {
    alert("Nessuna riga da importare.");
    return;
  }
  if (!canManageData()) {
    alert("Solo ionut29019@gmail.com può aggiungere impianti.");
    return;
  }

  const ref = db.collection("commesse").doc(selectedCommessaId).collection("impianti");

  for (let i = 0; i < pendingRows.length; i += 450) {
    const chunk = pendingRows.slice(i, i + 450);
    const batch = db.batch();

    chunk.forEach((row) => {
      const docRef = ref.doc();
      batch.set(docRef, {
        ...row,
        hasOrdinario: hasOrdinario(row.codicePrezzo),
        hasStraordinario: hasStraordinario(row.codicePrezzo),
        tipoManutenzione: classifyTipoManutenzione(row.codicePrezzo),
        done: false,
        doneAt: null,
        doneBy: "",
        createdAt: firebase.firestore.FieldValue.serverTimestamp()
      });
    });

    await batch.commit();
  }

  pendingRows = [];
  ui.excelFile.value = "";
  ui.importBtn.disabled = true;
  ui.importFeedback.textContent = "Import completato.";
}

function normalizeRow(row) {
  const keys = normalizeKeys(row);

  return {
    distretto: getValue(keys, ["distretto"]),
    idSap: getValue(keys, ["idsap"]),
    denominazione: getValue(keys, ["denominazioneimpianto", "denominazione", "impianto"]),
    comune: getValue(keys, ["comuneubicazioneimpianto", "comune"]),
    indirizzo: getValue(keys, ["viaecivicodiubicazioneimpianto", "indirizzo", "via"]),
    voceRiferimento: getValue(keys, ["vocediriferimentoelencoprezzi"]),
    codicePrezzo: getValue(keys, ["vocediriferimentoelencoprezzi", "codiceprezzo"]),
    sfalci: getValue(keys, ["sfalciareeverdimqpotaturasiepim"]),
    frequenzaAnnua: getValue(keys, ["frequenzaannuaminimasfalcieopotaturasiepin"]),
    tipologiaIntervento: getValue(keys, ["tipologiadisfalciointervento", "lavorazionirichieste"]),
    lavorazioniRichieste: getValue(keys, ["lavorazionirichieste", "tipologiadisfalciointervento"]),
    gpsY: parseCoordinate(getValue(keys, ["coordinategpsy"])),
    gpsX: parseCoordinate(getValue(keys, ["coordinategpsx"]))
  };
}

function normalizeKeys(obj) {
  return Object.entries(obj).reduce((acc, [key, value]) => {
    const normalizedKey = String(key || "")
      .toLowerCase()
      .normalize("NFD")
      .replace(/[\u0300-\u036f]/g, "")
      .replace(/[^a-z0-9]/g, "");

    acc[normalizedKey] = String(value || "").trim();
    return acc;
  }, {});
}

function getValue(obj, aliases) {
  for (const alias of aliases) {
    if (obj[alias]) return obj[alias];
  }
  return "";
}

function parseCoordinate(value) {
  const normalized = String(value || "").replace(",", ".").trim();
  const parsed = Number.parseFloat(normalized);
  return Number.isFinite(parsed) ? parsed : null;
}

function classifyTipoManutenzione(codicePrezzo) {
  const ordinario = hasOrdinario(codicePrezzo);
  const straordinario = hasStraordinario(codicePrezzo);
  if (ordinario && straordinario) return "Ordinaria + Straordinaria";
  if (ordinario) return "Ordinaria";
  if (straordinario) return "Straordinaria";
  return "Non specificata";
}

function splitCodes(codicePrezzo) {
  return String(codicePrezzo || "")
    .toUpperCase()
    .split(/[^A-Z0-9]+/)
    .map((c) => c.trim())
    .filter(Boolean);
}

function hasOrdinario(codicePrezzo) {
  const codes = splitCodes(codicePrezzo);
  return codes.includes("A11") || codes.includes("A12");
}

function hasStraordinario(codicePrezzo) {
  const codes = splitCodes(codicePrezzo);
  if (codes.length === 0) return false;
  return codes.some((code) => code !== "A11" && code !== "A12");
}

function renderImpianti() {
  ui.impiantiLista.innerHTML = "";

  if (!currentImpianti.length) {
    ui.impiantiLista.innerHTML = "<p class='muted'>Nessun impianto in questa commessa.</p>";
    return;
  }

  const sorted = [...currentImpianti].sort((a, b) => distanceFromUser(a) - distanceFromUser(b));

  sorted.forEach((impianto) => {
    const article = document.createElement("article");
    article.className = "impianto-item" + (impianto.done ? " done" : "");

    const distance = formatDistance(distanceFromUser(impianto));
    const tipo = impianto.tipoManutenzione || classifyTipoManutenzione(impianto.codicePrezzo);
    const hasStraordinariaFlag = impianto.hasStraordinario ?? hasStraordinario(impianto.codicePrezzo);
    article.innerHTML = `
      <strong>${escapeHTML(impianto.denominazione || "(senza nome)")}</strong>
      <p><b>Comune:</b> ${escapeHTML(impianto.comune || "-")}</p>
      <p><b>Indirizzo:</b> ${escapeHTML(impianto.indirizzo || "-")}</p>
      <p><b>Codice prezzo:</b> ${escapeHTML(impianto.codicePrezzo || impianto.voceRiferimento || "-")}</p>
      <p><b>Tipo:</b> <span class="badge ${hasStraordinariaFlag ? "badge-straordinaria" : "badge-ordinaria"}">${escapeHTML(tipo)}</span></p>
      <p><b>Lavorazioni richieste:</b> ${escapeHTML(impianto.lavorazioniRichieste || impianto.tipologiaIntervento || "-")}</p>
      <p><b>Distanza:</b> ${distance}</p>
      <p><b>Stato:</b> ${impianto.done ? "Fatto" : "Da fare"}</p>
    `;

    const actions = document.createElement("div");
    actions.className = "item-actions";

    const doneBtn = createButton("Fatto", () => markImpiantoDone(impianto));
    const navigateBtn = createButton("Naviga", () => navigateToImpianto(impianto));
    const resetBtn = createButton("Reset", () => resetImpianto(impianto));
    resetBtn.disabled = !canManageData();
    const waBtn = createButton("WhatsApp", () => openWhatsApp(impianto));
    const deleteBtn = createButton("Elimina", () => deleteImpianto(impianto));
    deleteBtn.disabled = !canManageData();

    actions.appendChild(doneBtn);
    actions.appendChild(navigateBtn);
    actions.appendChild(resetBtn);
    actions.appendChild(waBtn);
    actions.appendChild(deleteBtn);
    article.appendChild(actions);

    ui.impiantiLista.appendChild(article);
  });
}

function createButton(label, onClick) {
  const btn = document.createElement("button");
  btn.type = "button";
  btn.className = "btn";
  btn.textContent = label;
  btn.addEventListener("click", onClick);
  return btn;
}

async function navigateToImpianto(impianto) {
  if (!selectedCommessaId || !impianto.id) return;

  const lat = impianto.gpsY;
  const lng = impianto.gpsX;

  if (lat == null || lng == null) {
    alert("Coordinate mancanti per questo impianto.");
    return;
  }

  const url = `https://www.google.com/maps/dir/?api=1&destination=${lat},${lng}`;
  window.open(url, "_blank");
}

async function markImpiantoDone(impianto) {
  if (!selectedCommessaId || !impianto.id) return;
  updateImpiantoLocalState(impianto.id, { done: true });
  await setImpiantoDone(impianto.id, true);
}

async function resetImpianto(impianto) {
  if (!selectedCommessaId || !impianto.id) return;
  if (!canManageData()) {
    alert("Solo ionut29019@gmail.com può usare reset.");
    return;
  }
  updateImpiantoLocalState(impianto.id, { done: false });
  await setImpiantoDone(impianto.id, false);
}

async function deleteImpianto(impianto) {
  if (!selectedCommessaId || !impianto.id) return;
  if (!canManageData()) {
    alert("Solo ionut29019@gmail.com può eliminare impianti.");
    return;
  }

  const ok = window.confirm(`Eliminare impianto ${impianto.denominazione || ""}?`);
  if (!ok) return;

  await db.collection("commesse").doc(selectedCommessaId).collection("impianti").doc(impianto.id).delete();
}

async function deleteCommessa(commessaId, nome) {
  if (!canManageData()) {
    alert("Solo ionut29019@gmail.com può eliminare commesse.");
    return;
  }

  const ok = window.confirm(`Eliminare commessa ${nome}? Prima elimina gli impianti associati.`);
  if (!ok) return;

  await db.collection("commesse").doc(commessaId).delete();

  if (selectedCommessaId === commessaId) {
    selectedCommessaId = "";
    selectedCommessaName = "";
    stopImpiantiSubscription();
    ui.impiantiLista.innerHTML = "<p class='muted'>Seleziona una commessa.</p>";
    ui.commessaAttiva.textContent = "Seleziona una commessa.";
  }
}

async function setImpiantoDone(impiantoId, done) {
  const user = auth.currentUser;
  if (!user) return;

  const ref = db.collection("commesse").doc(selectedCommessaId).collection("impianti").doc(impiantoId);

  await ref.update({
    done,
    doneAt: done ? firebase.firestore.FieldValue.serverTimestamp() : null,
    doneBy: done ? (user.displayName || user.email || "Operatore") : ""
  });
}

function openWhatsApp(impianto) {
  const user = auth.currentUser;
  if (!user) {
    alert("Devi fare login.");
    return;
  }

  const now = new Date();
  const message = [
    "Manutenzione eseguita",
    `Commessa: ${selectedCommessaName || "-"}`,
    `Impianto: ${impianto.denominazione || "-"}`,
    `Operatore: ${user.displayName || user.email || "-"}`,
    `Data/Ora: ${now.toLocaleString("it-IT")}`
  ].join("\n");

  const url = `https://wa.me/?text=${encodeURIComponent(message)}`;
  window.open(url, "_blank");
}

function renderMap() {
  clearMap();

  const bounds = [];

  currentImpianti.forEach((impianto) => {
    if (impianto.gpsY == null || impianto.gpsX == null) return;

    const markerClass = getMarkerClass(impianto);
    const marker = L.marker([impianto.gpsY, impianto.gpsX], {
      icon: L.divIcon({
        className: "",
        html: `<div class="marker-pin ${markerClass}"></div>`,
        iconSize: [14, 14],
        iconAnchor: [7, 7]
      })
    });

    const tipo = impianto.tipoManutenzione || classifyTipoManutenzione(impianto.codicePrezzo);
    marker.bindPopup(`<b>${escapeHTML(impianto.denominazione || "Impianto")}</b><br>${escapeHTML(impianto.comune || "")}<br>${escapeHTML(tipo)}`);
    marker.addTo(markerLayer);
    bounds.push([impianto.gpsY, impianto.gpsX]);
  });

  if (currentUserPos) {
    L.circleMarker([currentUserPos.lat, currentUserPos.lng], {
      radius: 7,
      color: "#111827",
      fillColor: "#111827",
      fillOpacity: 0.85
    }).addTo(markerLayer).bindPopup("La tua posizione");
    bounds.push([currentUserPos.lat, currentUserPos.lng]);
  }

  if (bounds.length > 0) {
    map.fitBounds(bounds, { padding: [24, 24] });
  }
}

function getMarkerClass(impianto) {
  const done = Boolean(impianto.done);
  const straordinario = impianto.hasStraordinario ?? hasStraordinario(impianto.codicePrezzo);
  if (done) return "done";
  if (straordinario) return "straordinario";
  return "todo";
}

function updateImpiantoLocalState(impiantoId, patch) {
  currentImpianti = currentImpianti.map((item) => (
    item.id === impiantoId ? { ...item, ...patch } : item
  ));
  renderImpianti();
  renderMap();
}

function canManageData() {
  const email = (currentUser && currentUser.email) ? currentUser.email.toLowerCase() : "";
  return email === ADMIN_EMAIL;
}

function clearMap() {
  markerLayer.clearLayers();
}

function initGeolocation() {
  if (!navigator.geolocation) {
    ui.gpsStatus.textContent = "Geolocalizzazione non supportata dal browser.";
    return;
  }

  navigator.geolocation.getCurrentPosition((pos) => {
    currentUserPos = {
      lat: pos.coords.latitude,
      lng: pos.coords.longitude
    };
    ui.gpsStatus.textContent = "Posizione attiva per ordinare gli impianti per distanza.";
    renderImpianti();
    renderMap();
  }, () => {
    ui.gpsStatus.textContent = "Posizione non disponibile. Elenco non ordinato per distanza reale.";
  }, {
    enableHighAccuracy: true,
    timeout: 8000
  });
}

function distanceFromUser(impianto) {
  if (!currentUserPos || impianto.gpsY == null || impianto.gpsX == null) return Number.MAX_SAFE_INTEGER;
  return haversine(currentUserPos.lat, currentUserPos.lng, impianto.gpsY, impianto.gpsX);
}

function haversine(lat1, lon1, lat2, lon2) {
  const toRad = (deg) => (deg * Math.PI) / 180;
  const R = 6371;
  const dLat = toRad(lat2 - lat1);
  const dLon = toRad(lon2 - lon1);
  const a = Math.sin(dLat / 2) ** 2
    + Math.cos(toRad(lat1)) * Math.cos(toRad(lat2)) * Math.sin(dLon / 2) ** 2;
  return 2 * R * Math.asin(Math.sqrt(a));
}

function formatDistance(km) {
  if (!Number.isFinite(km) || km > 1e10) return "N/D";
  if (km < 1) return `${Math.round(km * 1000)} m`;
  return `${km.toFixed(2)} km`;
}

function escapeHTML(value) {
  return String(value || "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#039;");
}
