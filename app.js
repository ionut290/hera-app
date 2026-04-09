firebase.initializeApp(firebaseConfig);

const auth = firebase.auth();
const db = firebase.firestore();

const ui = {
  menuToggleBtn: document.getElementById("menu-toggle-btn"),
  menuCloseBtn: document.getElementById("menu-close-btn"),
  sideMenu: document.getElementById("side-menu"),
  menuOverlay: document.getElementById("menu-overlay"),
  loginBtn: document.getElementById("login-btn"),
  logoutBtn: document.getElementById("logout-btn"),
  driveConnectBtn: document.getElementById("drive-connect-btn"),
  user: document.getElementById("user"),
  userName: document.getElementById("user-name"),
  driveStatus: document.getElementById("drive-status"),
  commessaForm: document.getElementById("commessa-form"),
  commessaName: document.getElementById("commessa-name"),
  commesseLista: document.getElementById("commesse-lista"),
  commessaAttiva: document.getElementById("commessa-attiva"),
  excelFile: document.getElementById("excel-file"),
  importBtn: document.getElementById("import-btn"),
  importFeedback: document.getElementById("import-feedback"),
  impiantiLista: document.getElementById("impianti-lista"),
  gpsStatus: document.getElementById("gps-status"),
  chatOpenBtn: document.getElementById("chat-open-btn"),
  chatCounter: document.getElementById("chat-counter"),
  chatModal: document.getElementById("chat-modal"),
  chatCloseBtn: document.getElementById("chat-close-btn"),
  chatFullList: document.getElementById("chat-full-list"),
  chatSendForm: document.getElementById("chat-send-form"),
  chatRecipient: document.getElementById("chat-recipient"),
  chatText: document.getElementById("chat-text"),
  chatSendBtn: document.getElementById("chat-send-btn"),
  chatMediaInput: document.getElementById("chat-media-input"),
  chatVoiceBtn: document.getElementById("chat-voice-btn"),
  chatFeedback: document.getElementById("chat-feedback"),
  homePage: document.getElementById("home-page"),
  impiantiPage: document.getElementById("impianti-page"),
  backToHomeBtn: document.getElementById("back-to-home-btn"),
  exportCurrentCommessaBtn: document.getElementById("export-current-commessa-btn"),
  impiantiPageTitle: document.getElementById("impianti-page-title"),
  impiantoSearch: document.getElementById("impianto-search"),
  viewDoneBtn: document.getElementById("view-done-btn"),
  viewTodoBtn: document.getElementById("view-todo-btn"),
  personaleForm: document.getElementById("personale-form"),
  personaleNome: document.getElementById("personale-nome"),
  personaleLista: document.getElementById("personale-lista"),
  mezziForm: document.getElementById("mezzi-form"),
  mezzoNome: document.getElementById("mezzo-nome"),
  mezziLista: document.getElementById("mezzi-lista"),
  squadraForm: document.getElementById("squadra-form"),
  squadraCommessa: document.getElementById("squadra-commessa"),
  squadra1: document.getElementById("squadra-1"),
  squadra2: document.getElementById("squadra-2"),
  squadra3: document.getElementById("squadra-3"),
  squadra1Mezzi: document.getElementById("squadra-1-mezzi"),
  squadra2Mezzi: document.getElementById("squadra-2-mezzi"),
  squadra3Mezzi: document.getElementById("squadra-3-mezzi"),
  squadraRiferimento: document.getElementById("squadra-riferimento"),
  squadraHint: document.getElementById("squadra-hint"),
  squadreLista: document.getElementById("squadre-lista"),
  personaleExcelFile: document.getElementById("personale-excel-file"),
  personaleImportBtn: document.getElementById("personale-import-btn"),
  mezziExcelFile: document.getElementById("mezzi-excel-file"),
  mezziImportBtn: document.getElementById("mezzi-import-btn"),
  openPanelCommesse: document.getElementById("open-panel-commesse"),
  openPanelSquadre: document.getElementById("open-panel-squadre"),
  openPanelPersonale: document.getElementById("open-panel-personale"),
  openPanelMezzi: document.getElementById("open-panel-mezzi"),
  managementPage: document.getElementById("management-page"),
  managementTitle: document.getElementById("management-title"),
  managementCloseBtn: document.getElementById("management-close-btn"),
  panelCommesse: document.getElementById("panel-commesse"),
  panelSquadre: document.getElementById("panel-squadre"),
  panelPersonale: document.getElementById("panel-personale"),
  panelMezzi: document.getElementById("panel-mezzi"),
  weatherCard: document.getElementById("weather-card"),
  weatherSummary: document.getElementById("weather-summary"),
  weatherModal: document.getElementById("weather-modal"),
  weatherCloseBtn: document.getElementById("weather-close-btn"),
  weatherDetails: document.getElementById("weather-details")
};

let pendingRows = [];
let selectedCommessaId = "";
let selectedCommessaName = "";
let unsubscribeCommesse = null;
let unsubscribeImpianti = null;
let currentUserPos = null;
let currentImpianti = [];
let currentUser = null;
let unsubscribeChat = null;
let unsubscribeDriveBridge = null;
let unsubscribePersonale = null;
let unsubscribeMezzi = null;
let unsubscribeSquadre = null;
let unsubscribeUsers = null;
let chatMessages = [];
let platformUsers = [];
let mediaRecorder = null;
let mediaChunks = [];
let isRecording = false;
let lastReadChatAt = null;
let driveAccessToken = "";
let driveRootFolderId = "";
let driveChatFolderId = "";
let driveReportsFolderId = "";
const commessaSheetCache = new Map();
let commesseById = new Map();
let personaleRecords = [];
let mezziRecords = [];
let squadreByCommessa = new Map();
let highlightedImpiantoKey = "";
let impiantiSearchTerm = "";
let impiantiViewMode = "done";
let pendingSheetExports = [];
let sheetRetryTimer = null;
let isProcessingAdminSheetQueue = false;
const commessaSheetSyncTimers = new Map();

const DRIVE_CHAT_MEDIA_MAX_MB = 512;
const ADMIN_EMAIL = "ionut29019@gmail.com";
const PENDING_SHEET_EXPORTS_KEY = "heraPendingSheetExports";
const SHEET_RETRY_MS = 30 * 1000;
window.googleDriveAccessToken = localStorage.getItem("googleDriveAccessToken") || null;
driveAccessToken = window.googleDriveAccessToken || "";

const map = L.map("map");
L.tileLayer("https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png", {
  maxZoom: 19,
  attribution: "&copy; OpenStreetMap"
}).addTo(map);
map.setView([44.4949, 11.3426], 11);

const markerLayer = L.layerGroup().addTo(map);

ui.loginBtn.addEventListener("click", loginWithGoogle);
ui.menuToggleBtn.addEventListener("click", openSideMenu);
ui.menuCloseBtn.addEventListener("click", closeSideMenu);
ui.menuOverlay.addEventListener("click", closeSideMenu);
ui.logoutBtn.addEventListener("click", logout);
ui.driveConnectBtn.addEventListener("click", connectGoogleDrive);
ui.commessaForm.addEventListener("submit", createCommessa);
ui.excelFile.addEventListener("change", onExcelSelected);
ui.importBtn.addEventListener("click", importPendingRows);
ui.chatOpenBtn.addEventListener("click", openChatModal);
ui.chatCloseBtn.addEventListener("click", closeChatModal);
ui.chatSendForm.addEventListener("submit", sendTextMessage);
ui.chatMediaInput.addEventListener("change", sendMediaMessage);
ui.chatVoiceBtn.addEventListener("click", toggleVoiceRecording);
ui.backToHomeBtn.addEventListener("click", closeImpiantiPage);
ui.exportCurrentCommessaBtn.addEventListener("click", () => exportCommessaSummary(selectedCommessaId, selectedCommessaName));
ui.impiantoSearch.addEventListener("input", onImpiantoSearchInput);
ui.viewDoneBtn.addEventListener("click", () => setImpiantiViewMode("done"));
ui.viewTodoBtn.addEventListener("click", () => setImpiantiViewMode("todo"));
ui.personaleForm.addEventListener("submit", addPersonale);
ui.mezziForm.addEventListener("submit", addMezzo);
ui.squadraForm.addEventListener("submit", saveSquadraComposition);
ui.squadraCommessa.addEventListener("change", autofillSquadraForm);
ui.personaleImportBtn.addEventListener("click", importPersonaleFromExcel);
ui.mezziImportBtn.addEventListener("click", importMezziFromExcel);
ui.openPanelCommesse.addEventListener("click", () => openManagementPanel("commesse"));
ui.openPanelSquadre.addEventListener("click", () => openManagementPanel("squadre"));
ui.openPanelPersonale.addEventListener("click", () => openManagementPanel("personale"));
ui.openPanelMezzi.addEventListener("click", () => openManagementPanel("mezzi"));
ui.managementCloseBtn.addEventListener("click", closeManagementPanel);
ui.weatherCard.addEventListener("click", openWeatherModal);
ui.weatherCard.addEventListener("keydown", (event) => {
  if (event.key === "Enter" || event.key === " ") {
    event.preventDefault();
    openWeatherModal();
  }
});
ui.weatherCloseBtn.addEventListener("click", closeWeatherModal);

initGeolocation();
applyRoute();
window.addEventListener("hashchange", applyRoute);
loadPendingSheetExports();
startSheetRetryLoop();

auth.onAuthStateChanged((user) => {
  currentUser = user || null;
  const loggedIn = Boolean(user);
  const canManage = canManageData();

  ui.loginBtn.disabled = loggedIn;
  ui.logoutBtn.disabled = !loggedIn;
  ui.driveConnectBtn.disabled = !loggedIn || !canManage;
  ui.driveConnectBtn.classList.toggle("hidden", !loggedIn || !canManage);
  ui.user.textContent = loggedIn
    ? `Loggato: ${user.email || "email non disponibile"}`
    : "Non loggato";
  ui.userName.textContent = loggedIn
    ? `Nome utente: ${user.displayName || "Nome non disponibile"}`
    : "Nome utente: -";

  ui.importBtn.disabled = !loggedIn || !selectedCommessaId || pendingRows.length === 0 || !canManageData();
  ui.exportCurrentCommessaBtn.disabled = !loggedIn || !selectedCommessaId;
  updateAdminControls();

  stopCommesseSubscription();
  stopImpiantiSubscription();
  stopChatSubscription();
  stopDriveBridgeSubscription();
  stopPersonaleSubscription();
  stopMezziSubscription();
  stopSquadreSubscription();
  stopUsersSubscription();
  selectedCommessaId = "";
  selectedCommessaName = "";
  window.location.hash = "";
  ui.commesseLista.innerHTML = "";
  ui.squadraCommessa.innerHTML = "<option value=''>Seleziona commessa</option>";
  ui.squadreLista.innerHTML = "";
  squadreByCommessa = new Map();
  commesseById = new Map();
  ui.impiantiLista.innerHTML = loggedIn
    ? "<p class='muted'>Seleziona una commessa.</p>"
    : "<p class='muted'>Fai login per vedere le commesse.</p>";
  clearMap();
  lastReadChatAt = null;
  resetDriveState();
  if (loggedIn && canManage) {
    const storedToken = getStoredDriveToken();
    if (storedToken) {
      driveAccessToken = storedToken;
      window.googleDriveAccessToken = storedToken;
    }
  }
  renderChat([]);
  applyRoute();

  if (loggedIn) {
    upsertCurrentPlatformUser();
    subscribeCommesse();
    subscribeChat();
    subscribeUsers();
    subscribeDriveBridge();
    subscribePersonale();
    subscribeMezzi();
    subscribeSquadre();
    processPendingSheetExports();
  }
  fetchWeather();
});

function updateAdminControls() {
  const canManage = canManageData();
  ui.commessaName.disabled = !canManage;
  const submitBtn = ui.commessaForm.querySelector("button[type='submit']");
  if (submitBtn) submitBtn.disabled = !canManage;
  ui.personaleNome.disabled = !canManage;
  ui.mezzoNome.disabled = !canManage;
  if (ui.personaleForm.querySelector("button[type='submit']")) ui.personaleForm.querySelector("button[type='submit']").disabled = !canManage;
  if (ui.mezziForm.querySelector("button[type='submit']")) ui.mezziForm.querySelector("button[type='submit']").disabled = !canManage;
  ui.squadraCommessa.disabled = !canManage;
  ui.squadra1.disabled = !canManage;
  ui.squadra2.disabled = !canManage;
  ui.squadra3.disabled = !canManage;
  ui.squadra1Mezzi.disabled = !canManage;
  ui.squadra2Mezzi.disabled = !canManage;
  ui.squadra3Mezzi.disabled = !canManage;
  ui.squadraRiferimento.disabled = !canManage;
  if (ui.squadraForm.querySelector("button[type='submit']")) ui.squadraForm.querySelector("button[type='submit']").disabled = !canManage;
  ui.squadraHint.textContent = canManage
    ? "Suggerimento: usa i nomi in Personale e i mezzi in Mezzi per compilare le squadre."
    : "Solo l'admin può modificare personale, mezzi e composizione squadre.";
}

function openSideMenu() {
  ui.sideMenu.classList.remove("hidden");
  ui.menuOverlay.classList.remove("hidden");
  ui.sideMenu.setAttribute("aria-hidden", "false");
}

function closeSideMenu() {
  ui.sideMenu.classList.add("hidden");
  ui.menuOverlay.classList.add("hidden");
  ui.sideMenu.setAttribute("aria-hidden", "true");
}

function openManagementPanel(panel) {
  const panelMap = {
    commesse: { el: ui.panelCommesse, title: "Aggiungi commesse" },
    squadre: { el: ui.panelSquadre, title: "Composizione squadre" },
    personale: { el: ui.panelPersonale, title: "Personale" },
    mezzi: { el: ui.panelMezzi, title: "Mezzi" }
  };
  const target = panelMap[panel];
  if (!target) return;
  [ui.panelCommesse, ui.panelSquadre, ui.panelPersonale, ui.panelMezzi].forEach((el) => el.classList.add("hidden"));
  target.el.classList.remove("hidden");
  ui.managementTitle.textContent = target.title;
  ui.managementPage.classList.remove("hidden");
  ui.managementPage.setAttribute("aria-hidden", "false");
  closeSideMenu();
}

function closeManagementPanel() {
  ui.managementPage.classList.add("hidden");
  ui.managementPage.setAttribute("aria-hidden", "true");
}

function applyRoute() {
  const hash = window.location.hash || "";
  const match = hash.match(/^#commessa=([a-zA-Z0-9_-]+)$/);
  const commessaIdFromHash = match ? match[1] : "";
  const showImpianti = Boolean(commessaIdFromHash && selectedCommessaId === commessaIdFromHash);
  ui.homePage.classList.toggle("hidden", showImpianti);
  ui.impiantiPage.classList.toggle("hidden", !showImpianti);
  if (showImpianti) {
    ui.impiantiPageTitle.textContent = `Impianti commessa: ${selectedCommessaName || "Commessa"}`;
    setTimeout(() => map.invalidateSize(), 50);
  }
}

function openImpiantiPage() {
  if (!selectedCommessaId) return;
  window.location.hash = `commessa=${selectedCommessaId}`;
  applyRoute();
}

function closeImpiantiPage() {
  window.location.hash = "";
  ui.exportCurrentCommessaBtn.disabled = true;
  applyRoute();
}

function loadPendingSheetExports() {
  try {
    const raw = localStorage.getItem(PENDING_SHEET_EXPORTS_KEY);
    pendingSheetExports = raw ? JSON.parse(raw) : [];
    if (!Array.isArray(pendingSheetExports)) pendingSheetExports = [];
  } catch (error) {
    pendingSheetExports = [];
  }
}

function savePendingSheetExports() {
  localStorage.setItem(PENDING_SHEET_EXPORTS_KEY, JSON.stringify(pendingSheetExports));
}

function startSheetRetryLoop() {
  if (sheetRetryTimer) clearInterval(sheetRetryTimer);
  sheetRetryTimer = setInterval(() => {
    processPendingSheetExports();
    processAdminSheetExportQueue();
  }, SHEET_RETRY_MS);
}

function queuePendingSheetExport(payload) {
  pendingSheetExports.push({
    ...payload,
    attempts: Number(payload.attempts || 0),
    nextRetryAt: payload.nextRetryAt || Date.now()
  });
  savePendingSheetExports();
}

function getStoredDriveToken() {
  return localStorage.getItem("googleDriveAccessToken") || "";
}

function scheduleCommessaSheetSync(commessaId, commessaName = "", delayMs = 900) {
  if (!commessaId) return;
  if (!canManageData() || !driveAccessToken) return;
  const existingTimer = commessaSheetSyncTimers.get(commessaId);
  if (existingTimer) clearTimeout(existingTimer);
  const timer = setTimeout(async () => {
    try {
      await syncCommessaDoneImpiantiToDriveSheet(commessaId, commessaName || selectedCommessaName || "Commessa");
    } catch (error) {
      console.error("Sync foglio commessa fallita:", error);
    } finally {
      commessaSheetSyncTimers.delete(commessaId);
    }
  }, delayMs);
  commessaSheetSyncTimers.set(commessaId, timer);
}

async function queueSheetExportForAdmin(payload) {
  const createdBy = (currentUser && currentUser.email) ? currentUser.email : "";
  await db.collection("sheetExportQueue").add({
    status: "pending",
    attempts: 0,
    nextRetryAt: firebase.firestore.FieldValue.serverTimestamp(),
    createdAt: firebase.firestore.FieldValue.serverTimestamp(),
    createdBy,
    commessaId: payload.commessaId || "",
    commessaName: payload.commessaName || "",
    impianto: payload.impianto || {}
  });
}

function parseGoogleSheetId(value) {
  const input = String(value || "").trim();
  if (!input) return "";

  const urlMatch = input.match(/\/spreadsheets\/d\/([a-zA-Z0-9-_]+)/);
  if (urlMatch && urlMatch[1]) return urlMatch[1];

  const idOnly = input.match(/^[a-zA-Z0-9-_]{20,}$/);
  return idOnly ? input : "";
}

function loginWithGoogle() {
  const provider = new firebase.auth.GoogleAuthProvider();
  provider.addScope("https://www.googleapis.com/auth/userinfo.email");
  provider.addScope("https://www.googleapis.com/auth/drive.file");

  auth.signInWithPopup(provider).then((result) => {
    // Gli scope extra (es. Drive) arrivano nel credential del login, non nel profilo utente Firebase.
    const accessToken = extractGoogleAccessToken(result);
    if (accessToken) {
      driveAccessToken = accessToken;
    }
  }).catch((error) => {
    if (error.code === "auth/popup-blocked" || error.code === "auth/cancelled-popup-request") {
      return auth.signInWithRedirect(provider);
    }
    alert("Errore login: " + error.message);
  });
}

function logout() {
  resetDriveState();
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

function subscribeDriveBridge() {
  unsubscribeDriveBridge = db.collection("appConfig").doc("driveBridge").onSnapshot((doc) => {
    const data = doc.exists ? doc.data() : null;
    if (!data || !data.accessToken) {
      if (canManageData()) {
        const storedToken = getStoredDriveToken();
        if (storedToken) {
          driveAccessToken = storedToken;
          window.googleDriveAccessToken = storedToken;
        }
      }
      updateDriveStatus(Boolean(driveAccessToken));
      return;
    }

    const owner = data.ownerEmail || ADMIN_EMAIL;
    if (canManageData()) {
      driveAccessToken = data.accessToken;
      driveRootFolderId = data.rootFolderId || "";
      driveChatFolderId = data.chatFolderId || "";
      driveReportsFolderId = data.reportsFolderId || "";
      ui.driveStatus.textContent = `Drive centralizzato attivo (${owner}).`;
      processPendingSheetExports();
      processAdminSheetExportQueue();
      return;
    }

    driveAccessToken = "";
    driveRootFolderId = "";
    driveChatFolderId = "";
    driveReportsFolderId = "";
    ui.driveStatus.textContent = `Report foglio gestito dall'admin (${owner}).`;
  }, (error) => {
    console.error(error);
    ui.driveStatus.textContent = "Errore lettura configurazione Drive centralizzato.";
  });
}

function stopDriveBridgeSubscription() {
  if (unsubscribeDriveBridge) {
    unsubscribeDriveBridge();
    unsubscribeDriveBridge = null;
  }
}

function subscribeCommesse() {
  unsubscribeCommesse = db
    .collection("commesse")
    .orderBy("createdAt", "desc")
    .onSnapshot((snapshot) => {
      ui.commesseLista.innerHTML = "";
      commesseById = new Map();
      ui.squadraCommessa.innerHTML = "<option value=''>Seleziona commessa</option>";

      if (snapshot.empty) {
        ui.commesseLista.innerHTML = "<p class='muted'>Nessuna commessa disponibile.</p>";
        return;
      }

      snapshot.forEach((doc) => {
        const commessa = doc.data();
        commesseById.set(doc.id, { id: doc.id, ...commessa });
        const row = document.createElement("div");
        row.className = "commessa-row";

        const btn = document.createElement("button");
        btn.type = "button";
        btn.className = "btn commessa-btn" + (doc.id === selectedCommessaId ? " active" : "");
        btn.dataset.commessaId = doc.id;
        btn.textContent = commessa.nome || "Commessa senza nome";
        btn.addEventListener("click", () => selectCommessa(doc.id, commessa.nome || "Commessa"));

        row.appendChild(btn);
        ui.commesseLista.appendChild(row);

        const option = document.createElement("option");
        option.value = doc.id;
        option.textContent = commessa.nome || "Commessa senza nome";
        ui.squadraCommessa.appendChild(option);
      });

      renderSquadre();
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
  ui.exportCurrentCommessaBtn.disabled = !auth.currentUser;
  updateCommessaButtonsActive();

  stopImpiantiSubscription();
  subscribeImpianti();
  openImpiantiPage();
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
      const rawImpianti = snapshot.docs.map((doc) => ({ id: doc.id, ...doc.data() }));
      currentImpianti = combineImpiantiForView(rawImpianti);
      renderImpianti();
      renderMap();
      scheduleCommessaSheetSync(selectedCommessaId, selectedCommessaName, 1200);
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

      const normalizedRows = rows.map(normalizeRow).filter((row) => row.denominazione);
      pendingRows = mergeRowsByImpianto(normalizedRows);
      ui.importFeedback.textContent = `Righe valide: ${normalizedRows.length}. Impianti unici: ${pendingRows.length}`;
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

function mergeRowsByImpianto(rows) {
  const grouped = new Map();

  rows.forEach((row) => {
    const key = buildImpiantoKey(row);
    const existing = grouped.get(key);
    if (!existing) {
      grouped.set(key, { ...row });
      return;
    }

    existing.codicePrezzo = mergeMultiValue(existing.codicePrezzo, row.codicePrezzo);
    existing.voceRiferimento = mergeMultiValue(existing.voceRiferimento, row.voceRiferimento);
    existing.tipologiaIntervento = mergeMultiValue(existing.tipologiaIntervento, row.tipologiaIntervento);
    existing.lavorazioniRichieste = mergeMultiValue(existing.lavorazioniRichieste, row.lavorazioniRichieste);
    existing.frequenzaAnnua = mergeMultiValue(existing.frequenzaAnnua, row.frequenzaAnnua);

    if (!existing.distretto && row.distretto) existing.distretto = row.distretto;
    if (!existing.comune && row.comune) existing.comune = row.comune;
    if (!existing.indirizzo && row.indirizzo) existing.indirizzo = row.indirizzo;
    if (!existing.idSap && row.idSap) existing.idSap = row.idSap;
    if (existing.gpsY == null && row.gpsY != null) existing.gpsY = row.gpsY;
    if (existing.gpsX == null && row.gpsX != null) existing.gpsX = row.gpsX;
  });

  return Array.from(grouped.values());
}

function combineImpiantiForView(impianti) {
  const grouped = new Map();

  impianti.forEach((item) => {
    const key = buildImpiantoKey(item);
    const existing = grouped.get(key);
    if (!existing) {
      grouped.set(key, {
        ...item,
        done: Boolean(item.done),
        doneAt: item.doneAt || null,
        doneBy: item.doneBy || "",
        sourceIds: [item.id]
      });
      return;
    }

    existing.sourceIds.push(item.id);
    existing.codicePrezzo = mergeMultiValue(existing.codicePrezzo, item.codicePrezzo);
    existing.voceRiferimento = mergeMultiValue(existing.voceRiferimento, item.voceRiferimento);
    existing.tipologiaIntervento = mergeMultiValue(existing.tipologiaIntervento, item.tipologiaIntervento);
    existing.lavorazioniRichieste = mergeMultiValue(existing.lavorazioniRichieste, item.lavorazioniRichieste);
    existing.frequenzaAnnua = mergeMultiValue(existing.frequenzaAnnua, item.frequenzaAnnua);

    existing.hasOrdinario = hasOrdinario(existing.codicePrezzo);
    existing.hasStraordinario = hasStraordinario(existing.codicePrezzo);
    existing.tipoManutenzione = classifyTipoManutenzione(existing.codicePrezzo);
    const itemDone = Boolean(item.done);
    const existingDoneAtMs = firestoreDateToMillis(existing.doneAt);
    const itemDoneAtMs = firestoreDateToMillis(item.doneAt);
    existing.done = Boolean(existing.done || itemDone);

    if (itemDone && (!existing.doneBy || itemDoneAtMs >= existingDoneAtMs)) {
      existing.doneBy = item.doneBy || existing.doneBy || "";
    }
    if (itemDoneAtMs >= existingDoneAtMs) {
      existing.doneAt = item.doneAt || existing.doneAt || null;
    }

    if (!existing.idSap && item.idSap) existing.idSap = item.idSap;
    if (!existing.comune && item.comune) existing.comune = item.comune;
    if (!existing.indirizzo && item.indirizzo) existing.indirizzo = item.indirizzo;
    if (existing.gpsY == null && item.gpsY != null) existing.gpsY = item.gpsY;
    if (existing.gpsX == null && item.gpsX != null) existing.gpsX = item.gpsX;
  });

  return Array.from(grouped.values());
}

function buildImpiantoKey(row) {
  const idSap = String(row.idSap || "").trim().toLowerCase();
  if (idSap) return `sap:${idSap}`;

  const normalizedName = String(row.denominazione || "").trim().toLowerCase();
  const normalizedComune = String(row.comune || "").trim().toLowerCase();
  const normalizedAddress = String(row.indirizzo || "").trim().toLowerCase();
  return `name:${normalizedName}|comune:${normalizedComune}|indirizzo:${normalizedAddress}`;
}

function mergeMultiValue(oldValue, newValue) {
  const values = [oldValue, newValue]
    .flatMap((value) => String(value || "").split("|"))
    .map((value) => value.trim())
    .filter(Boolean);

  return Array.from(new Set(values)).join(" | ");
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

function firestoreDateToMillis(value) {
  if (!value) return 0;
  if (typeof value.toDate === "function") {
    const date = value.toDate();
    return date instanceof Date ? date.getTime() : 0;
  }
  if (value instanceof Date) return value.getTime();
  const parsed = Date.parse(value);
  return Number.isFinite(parsed) ? parsed : 0;
}

function formatDoneDateTime(doneAt) {
  const millis = firestoreDateToMillis(doneAt);
  if (!millis) return { date: "-", time: "-" };
  const date = new Date(millis);
  return {
    date: date.toLocaleDateString("it-IT"),
    time: date.toLocaleTimeString("it-IT", { hour: "2-digit", minute: "2-digit", hour12: false })
  };
}

async function exportCommessaSummary(commessaId, commessaName) {
  if (!auth.currentUser) {
    alert("Devi fare login per esportare il riepilogo.");
    return;
  }

  try {
    const snapshot = await db
      .collection("commesse")
      .doc(commessaId)
      .collection("impianti")
      .get();

    const rawImpianti = snapshot.docs.map((doc) => ({ id: doc.id, ...doc.data() }));
    const doneImpianti = combineImpiantiForView(rawImpianti).filter((impianto) => impianto.done);

    if (!doneImpianti.length) {
      alert(`Nessun impianto FATTO da esportare per la commessa "${commessaName}".`);
      return;
    }

    const rows = doneImpianti.flatMap((impianto) => buildRowsForEachCodicePrezzo(impianto)).map((impianto) => {
      const doneInfo = formatDoneDateTime(impianto.doneAt);
      return {
        Commessa: commessaName || "",
        Cantiere: impianto.cantiereRiga || "",
        Distretto: impianto.distretto || "",
        "ID SAP": impianto.idSap || "",
        Denominazione: impianto.denominazione || "",
        Comune: impianto.comune || "",
        Indirizzo: impianto.indirizzo || "",
        "Voce riferimento": impianto.voceRiferimento || "",
        "Codice prezzo": impianto.codicePrezzoSingolo || impianto.codicePrezzo || "",
        Sfalci: impianto.sfalci || "",
        "Frequenza annua": impianto.frequenzaAnnua || "",
        "Tipologia intervento": impianto.tipologiaIntervento || "",
        "Lavorazioni richieste": impianto.lavorazioniRichieste || "",
        "GPS Y": impianto.gpsY ?? "",
        "GPS X": impianto.gpsX ?? "",
        "Tipo manutenzione": impianto.tipoManutenzione || classifyTipoManutenzione(impianto.codicePrezzo),
        Stato: "Fatto",
        "Eseguito da": impianto.doneBy || "-",
        "Data esecuzione": doneInfo.date,
        "Ora esecuzione (hh:mm)": doneInfo.time
      };
    });

    const worksheet = XLSX.utils.json_to_sheet(rows);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, "Riepilogo impianti");

    const safeName = String(commessaName || "commessa")
      .trim()
      .replace(/[\\/:*?"<>|]/g, "_")
      .replace(/\s+/g, "_");
    const timestamp = new Date().toISOString().slice(0, 19).replace(/[:T]/g, "-");
    XLSX.writeFile(workbook, `riepilogo_impianti_${safeName}_${timestamp}.xlsx`);
  } catch (error) {
    console.error(error);
    alert("Errore durante l'esportazione del riepilogo in Excel.");
  }
}

function splitCodes(codicePrezzo) {
  return String(codicePrezzo || "")
    .toUpperCase()
    .split(/[^A-Z0-9]+/)
    .map((c) => c.trim())
    .filter(Boolean);
}

function buildRowsForEachCodicePrezzo(impianto) {
  const rawCodes = String(impianto.codicePrezzo || impianto.voceRiferimento || "")
    .split("|")
    .map((value) => value.trim())
    .filter(Boolean);
  const codes = rawCodes.length ? rawCodes : [""];
  return codes.map((code) => ({
    ...impianto,
    codicePrezzoSingolo: code,
    cantiereRiga: `${impianto.denominazione || "Impianto"}`
  }));
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

function onImpiantoSearchInput(event) {
  impiantiSearchTerm = String(event.target.value || "").trim().toLowerCase();
  renderImpianti();
}

function setImpiantiViewMode(mode) {
  impiantiViewMode = mode === "todo" ? "todo" : "done";
  ui.viewDoneBtn.classList.toggle("btn-primary", impiantiViewMode === "done");
  ui.viewTodoBtn.classList.toggle("btn-primary", impiantiViewMode === "todo");
  renderImpianti();
}

function matchesImpiantoSearch(impianto) {
  if (!impiantiSearchTerm) return true;
  const haystack = [
    impianto.denominazione,
    impianto.comune,
    impianto.indirizzo,
    impianto.codicePrezzo,
    impianto.voceRiferimento,
    impianto.idSap
  ].map((value) => String(value || "").toLowerCase()).join(" ");
  return haystack.includes(impiantiSearchTerm);
}

function renderImpianti() {
  ui.impiantiLista.innerHTML = "";

  if (!currentImpianti.length) {
    ui.impiantiLista.innerHTML = "<p class='muted'>Nessun impianto in questa commessa.</p>";
    return;
  }

  const filtered = currentImpianti.filter((impianto) => {
    const viewMatch = impiantiViewMode === "done" ? Boolean(impianto.done) : !impianto.done;
    return viewMatch && matchesImpiantoSearch(impianto);
  });
  const sorted = [...filtered].sort((a, b) => distanceFromUser(a) - distanceFromUser(b));

  if (!sorted.length) {
    ui.impiantiLista.innerHTML = "<p class='muted'>Nessun impianto trovato con i filtri correnti.</p>";
    return;
  }

  sorted.forEach((impianto) => {
    const article = document.createElement("article");
    article.className = "impianto-item" + (impianto.done ? " done" : "");
    const impiantoKey = buildImpiantoKey(impianto);
    article.dataset.impiantoKey = impiantoKey;
    if (highlightedImpiantoKey === impiantoKey) article.classList.add("highlight");

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
      <p><b>Eseguito da:</b> ${escapeHTML(impianto.doneBy || "-")}</p>
    `;

    const actions = document.createElement("div");
    actions.className = "item-actions";

    const doneBtn = createButton("Fatto", () => markImpiantoDone(impianto));
    doneBtn.disabled = Boolean(impianto.done);
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
  const ids = getImpiantoDocIds(impianto);
  if (!selectedCommessaId || !ids.length) return;
  const exportPayload = {
    commessaId: selectedCommessaId,
    commessaName: selectedCommessaName || "Commessa",
    impianto
  };
  try {
    updateImpiantoLocalState(ids, { done: true });
    await setImpiantoDone(ids, true);
  } catch (error) {
    console.error("Aggiornamento stato FATTO non completato al primo tentativo:", error);
    retrySetImpiantoDone(ids, true);
  }

  if (!canManageData()) {
    try {
      await queueSheetExportForAdmin(exportPayload);
    } catch (error) {
      console.error("Impianto FATTO ma coda admin non salvata:", error);
    }
    return;
  }

  try {
    await syncCommessaDoneImpiantiToDriveSheet(exportPayload.commessaId, exportPayload.commessaName);
  } catch (error) {
    console.error(error);
    const retried = await retrySheetExport(exportPayload, 2);
    if (!retried) {
      queuePendingSheetExport(exportPayload);
      try {
        await queueSheetExportForAdmin(exportPayload);
      } catch (queueError) {
        console.error("Impianto FATTO ma coda condivisa non salvata:", queueError);
      }
      console.warn("Export foglio in coda: sarà ritentato automaticamente.");
    }
  }
}

async function retrySetImpiantoDone(impiantoIds, done, retries = 3) {
  for (let i = 0; i < retries; i += 1) {
    try {
      await setImpiantoDone(impiantoIds, done);
      return true;
    } catch (error) {
      console.warn(`Tentativo aggiornamento stato FATTO fallito (${i + 1}/${retries})`, error);
      await new Promise((resolve) => setTimeout(resolve, 1000 * (i + 1)));
    }
  }
  return false;
}

async function retrySheetExport(payload, retries = 2) {
  for (let i = 0; i < retries; i += 1) {
    try {
      await syncCommessaDoneImpiantiToDriveSheet(payload.commessaId, payload.commessaName);
      return true;
    } catch (error) {
      console.warn(`Tentativo export foglio fallito (${i + 1}/${retries})`, error);
    }
  }
  return false;
}

async function processPendingSheetExports() {
  if (!canManageData()) return;
  if (!pendingSheetExports.length) return;
  const now = Date.now();
  const remaining = [];

  for (const item of pendingSheetExports) {
    if ((item.nextRetryAt || 0) > now) {
      remaining.push(item);
      continue;
    }
    try {
      await syncCommessaDoneImpiantiToDriveSheet(item.commessaId, item.commessaName);
    } catch (error) {
      const attempts = Number(item.attempts || 0) + 1;
      if (attempts < 20) {
        remaining.push({
          ...item,
          attempts,
          nextRetryAt: Date.now() + Math.min(SHEET_RETRY_MS * attempts, 10 * 60 * 1000)
        });
      } else {
        console.error("Export foglio fallito definitivamente dopo più tentativi:", error);
      }
    }
  }

  pendingSheetExports = remaining;
  savePendingSheetExports();
}

async function processAdminSheetExportQueue() {
  if (!canManageData()) return;
  if (!driveAccessToken) return;
  if (isProcessingAdminSheetQueue) return;
  isProcessingAdminSheetQueue = true;
  try {
    const now = new Date();
    const snapshot = await db
      .collection("sheetExportQueue")
      .where("status", "==", "pending")
      .where("nextRetryAt", "<=", now)
      .orderBy("nextRetryAt", "asc")
      .limit(20)
      .get();

    for (const doc of snapshot.docs) {
      const data = doc.data() || {};
      try {
        await syncCommessaDoneImpiantiToDriveSheet(data.commessaId || "", data.commessaName || "Commessa");
        await doc.ref.set({
          status: "done",
          doneAt: firebase.firestore.FieldValue.serverTimestamp(),
          lastError: firebase.firestore.FieldValue.delete()
        }, { merge: true });
      } catch (error) {
        const attempts = Number(data.attempts || 0) + 1;
        const retryMs = Math.min(SHEET_RETRY_MS * attempts, 10 * 60 * 1000);
        await doc.ref.set({
          status: attempts >= 20 ? "failed" : "pending",
          attempts,
          nextRetryAt: new Date(Date.now() + retryMs),
          lastError: String(error && error.message ? error.message : error).slice(0, 500),
          updatedAt: firebase.firestore.FieldValue.serverTimestamp()
        }, { merge: true });
      }
    }
  } catch (error) {
    console.error("Errore processamento coda export foglio (admin):", error);
  } finally {
    isProcessingAdminSheetQueue = false;
  }
}

async function resetImpianto(impianto) {
  const ids = getImpiantoDocIds(impianto);
  if (!selectedCommessaId || !ids.length) return;
  if (!canManageData()) {
    alert("Solo ionut29019@gmail.com può usare reset.");
    return;
  }
  updateImpiantoLocalState(ids, { done: false });
  await setImpiantoDone(ids, false);
}

async function deleteImpianto(impianto) {
  const ids = getImpiantoDocIds(impianto);
  if (!selectedCommessaId || !ids.length) return;
  if (!canManageData()) {
    alert("Solo ionut29019@gmail.com può eliminare impianti.");
    return;
  }

  const ok = window.confirm(`Eliminare impianto ${impianto.denominazione || ""}?`);
  if (!ok) return;

  const ref = db.collection("commesse").doc(selectedCommessaId).collection("impianti");
  await Promise.all(ids.map((id) => ref.doc(id).delete()));
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
    window.location.hash = "";
    stopImpiantiSubscription();
    ui.impiantiLista.innerHTML = "<p class='muted'>Seleziona una commessa.</p>";
    ui.commessaAttiva.textContent = "Seleziona una commessa.";
    applyRoute();
  }
}

async function addPersonale(event) {
  event.preventDefault();
  if (!canManageData()) {
    alert("Solo ionut29019@gmail.com può gestire il personale.");
    return;
  }
  const nome = ui.personaleNome.value.trim();
  if (!nome) return;
  await db.collection("personale").add({
    nome,
    createdAt: firebase.firestore.FieldValue.serverTimestamp()
  });
  ui.personaleForm.reset();
}

async function addMezzo(event) {
  event.preventDefault();
  if (!canManageData()) {
    alert("Solo ionut29019@gmail.com può gestire i mezzi.");
    return;
  }
  const nome = ui.mezzoNome.value.trim();
  if (!nome) return;
  await db.collection("mezzi").add({
    nome,
    createdAt: firebase.firestore.FieldValue.serverTimestamp()
  });
  ui.mezziForm.reset();
}

async function importPersonaleFromExcel() {
  if (!canManageData()) {
    alert("Solo ionut29019@gmail.com può importare il personale.");
    return;
  }
  await importSimpleRegistryFromExcel(ui.personaleExcelFile, "personale");
}

async function importMezziFromExcel() {
  if (!canManageData()) {
    alert("Solo ionut29019@gmail.com può importare i mezzi.");
    return;
  }
  await importSimpleRegistryFromExcel(ui.mezziExcelFile, "mezzi");
}

async function importSimpleRegistryFromExcel(inputEl, collectionName) {
  const file = inputEl.files && inputEl.files[0];
  if (!file) {
    alert("Seleziona un file Excel.");
    return;
  }

  const rows = await parseSimpleExcelRows(file);
  const uniqueNames = [...new Set(rows.filter(Boolean).map((v) => v.trim()).filter(Boolean))];
  if (!uniqueNames.length) {
    alert("Il file Excel non contiene nomi validi.");
    return;
  }

  const batch = db.batch();
  uniqueNames.forEach((nome) => {
    const ref = db.collection(collectionName).doc();
    batch.set(ref, {
      nome,
      createdAt: firebase.firestore.FieldValue.serverTimestamp()
    });
  });
  await batch.commit();
  inputEl.value = "";
  alert(`Import completato (${uniqueNames.length} elementi).`);
}

async function parseSimpleExcelRows(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = (event) => {
      try {
        const data = new Uint8Array(event.target.result);
        const workbook = XLSX.read(data, { type: "array" });
        const sheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[sheetName];
        const jsonRows = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
        const names = [];
        jsonRows.forEach((row) => {
          if (!Array.isArray(row) || !row.length) return;
          const firstCell = String(row[0] || "").trim();
          if (firstCell && firstCell.toLowerCase() !== "nome") names.push(firstCell);
        });
        resolve(names);
      } catch (error) {
        reject(error);
      }
    };
    reader.onerror = () => reject(reader.error || new Error("Errore lettura file"));
    reader.readAsArrayBuffer(file);
  });
}

function subscribePersonale() {
  unsubscribePersonale = db.collection("personale").orderBy("createdAt", "asc").onSnapshot((snapshot) => {
    personaleRecords = snapshot.docs.map((doc) => ({ id: doc.id, ...doc.data() }));
    renderSimpleList(ui.personaleLista, personaleRecords, deletePersonale);
    updateSquadraHintFromSources();
  }, (error) => {
    console.error(error);
    ui.personaleLista.innerHTML = "<p class='muted'>Errore caricamento personale.</p>";
  });
}

function subscribeMezzi() {
  unsubscribeMezzi = db.collection("mezzi").orderBy("createdAt", "asc").onSnapshot((snapshot) => {
    mezziRecords = snapshot.docs.map((doc) => ({ id: doc.id, ...doc.data() }));
    renderSimpleList(ui.mezziLista, mezziRecords, deleteMezzo);
    updateSquadraHintFromSources();
  }, (error) => {
    console.error(error);
    ui.mezziLista.innerHTML = "<p class='muted'>Errore caricamento mezzi.</p>";
  });
}

function subscribeSquadre() {
  unsubscribeSquadre = db.collection("squadreCommesse").onSnapshot((snapshot) => {
    squadreByCommessa = new Map();
    snapshot.forEach((doc) => {
      squadreByCommessa.set(doc.id, { id: doc.id, ...doc.data() });
    });
    renderSquadre();
    autofillSquadraForm();
  }, (error) => {
    console.error(error);
    ui.squadreLista.innerHTML = "<p class='muted'>Errore caricamento squadre.</p>";
  });
}

function stopPersonaleSubscription() {
  if (unsubscribePersonale) {
    unsubscribePersonale();
    unsubscribePersonale = null;
  }
}

function stopMezziSubscription() {
  if (unsubscribeMezzi) {
    unsubscribeMezzi();
    unsubscribeMezzi = null;
  }
}

function stopSquadreSubscription() {
  if (unsubscribeSquadre) {
    unsubscribeSquadre();
    unsubscribeSquadre = null;
  }
}

function renderSimpleList(container, items, onDelete) {
  container.innerHTML = "";
  if (!items.length) {
    container.innerHTML = "<p class='muted'>Nessun elemento.</p>";
    return;
  }
  items.forEach((item) => {
    const row = document.createElement("div");
    row.className = "simple-list-item";
    const label = document.createElement("span");
    label.textContent = item.nome || "-";
    row.appendChild(label);
    const deleteBtn = createButton("Elimina", () => onDelete(item.id, item.nome || "elemento"));
    deleteBtn.disabled = !canManageData();
    row.appendChild(deleteBtn);
    container.appendChild(row);
  });
}

async function deletePersonale(id, nome) {
  if (!canManageData()) return;
  if (!window.confirm(`Eliminare ${nome} dal personale?`)) return;
  await db.collection("personale").doc(id).delete();
}

async function deleteMezzo(id, nome) {
  if (!canManageData()) return;
  if (!window.confirm(`Eliminare ${nome} dai mezzi?`)) return;
  await db.collection("mezzi").doc(id).delete();
}

function autofillSquadraForm() {
  const commessaId = ui.squadraCommessa.value;
  if (!commessaId) {
    ui.squadra1.value = "";
    ui.squadra2.value = "";
    ui.squadra3.value = "";
    ui.squadra1Mezzi.value = "";
    ui.squadra2Mezzi.value = "";
    ui.squadra3Mezzi.value = "";
    ui.squadraRiferimento.value = "";
    return;
  }

  const data = squadreByCommessa.get(commessaId) || {};
  ui.squadra1.value = data.squadra1 || "";
  ui.squadra2.value = data.squadra2 || "";
  ui.squadra3.value = data.squadra3 || "";
  ui.squadra1Mezzi.value = data.squadra1Mezzi || "";
  ui.squadra2Mezzi.value = data.squadra2Mezzi || "";
  ui.squadra3Mezzi.value = data.squadra3Mezzi || "";
  ui.squadraRiferimento.value = data.riferimentoData || "";
}

async function saveSquadraComposition(event) {
  event.preventDefault();
  if (!canManageData()) {
    alert("Solo ionut29019@gmail.com può compilare la composizione squadre.");
    return;
  }
  const commessaId = ui.squadraCommessa.value;
  if (!commessaId) {
    alert("Seleziona una commessa.");
    return;
  }
  await db.collection("squadreCommesse").doc(commessaId).set({
    commessaId,
    commessaNome: (commesseById.get(commessaId) || {}).nome || "Commessa",
    riferimentoData: ui.squadraRiferimento.value || "",
    squadra1: ui.squadra1.value.trim(),
    squadra1Mezzi: ui.squadra1Mezzi.value.trim(),
    squadra2: ui.squadra2.value.trim(),
    squadra2Mezzi: ui.squadra2Mezzi.value.trim(),
    squadra3: ui.squadra3.value.trim(),
    squadra3Mezzi: ui.squadra3Mezzi.value.trim(),
    updatedAt: firebase.firestore.FieldValue.serverTimestamp(),
    updatedBy: (currentUser && currentUser.email) ? currentUser.email : ""
  }, { merge: true });
}

function renderSquadre() {
  ui.squadreLista.innerHTML = "";

  const commesse = Array.from(commesseById.values());
  const commesseConSquadre = commesse.filter((commessa) => {
    const squad = squadreByCommessa.get(commessa.id) || {};
    return Boolean(
      squad.squadra1 || squad.squadra2 || squad.squadra3
      || squad.squadra1Mezzi || squad.squadra2Mezzi || squad.squadra3Mezzi
    );
  });
  if (!commesseConSquadre.length) {
    ui.squadreLista.innerHTML = "<p class='muted'>Nessuna commessa disponibile.</p>";
    return;
  }

  commesseConSquadre.forEach((commessa) => {
    const item = document.createElement("article");
    item.className = "squadra-item";
    const squad = squadreByCommessa.get(commessa.id) || {};
    const riferimento = squad.riferimentoData
      ? new Date(`${squad.riferimentoData}T00:00:00`).toLocaleDateString("it-IT")
      : "-";
    item.innerHTML = `
      <strong>📁 ${escapeHTML(commessa.nome || "Commessa senza nome")}</strong>
      <p><b>📅 Giorno:</b> ${escapeHTML(riferimento)}</p>
      <p><b>👥 Squadra 1:</b> ${escapeHTML(squad.squadra1 || "-")}</p>
      <p><b>🚚 Mezzi 1:</b> ${escapeHTML(squad.squadra1Mezzi || "-")}</p>
      <p><b>👥 Squadra 2:</b> ${escapeHTML(squad.squadra2 || "-")}</p>
      <p><b>🚚 Mezzi 2:</b> ${escapeHTML(squad.squadra2Mezzi || "-")}</p>
      <p><b>👥 Squadra 3:</b> ${escapeHTML(squad.squadra3 || "-")}</p>
      <p><b>🚚 Mezzi 3:</b> ${escapeHTML(squad.squadra3Mezzi || "-")}</p>
    `;
    const askBtn = createButton("WhatsApp al tecnico", () => openSquadraWhatsApp(squad, commessa));
    item.appendChild(askBtn);
    ui.squadreLista.appendChild(item);
  });
}

function updateSquadraHintFromSources() {
  if (!canManageData()) return;
  const personale = personaleRecords.map((p) => p.nome).filter(Boolean).join(", ") || "Nessuno";
  const mezzi = mezziRecords.map((m) => m.nome).filter(Boolean).join(", ") || "Nessuno";
  ui.squadraHint.textContent = `Personale disponibile (assegnalo a Squadra 1/2/3): ${personale}. Mezzi disponibili (assegnali per squadra): ${mezzi}.`;
}

async function setImpiantoDone(impiantoIds, done) {
  const user = auth.currentUser;
  if (!user) return;

  const ref = db.collection("commesse").doc(selectedCommessaId).collection("impianti");
  await Promise.all(impiantoIds.map((impiantoId) => ref.doc(impiantoId).update({
    done,
    doneAt: done ? firebase.firestore.FieldValue.serverTimestamp() : null,
    doneBy: done ? (user.displayName || user.email || "Operatore") : ""
  })));
}

function openWhatsApp(impianto) {
  const user = auth.currentUser;
  if (!user) {
    alert("Devi fare login.");
    return;
  }

  const now = new Date();
  const isOnlyOrdinaria = hasOrdinario(impianto.codicePrezzo) && !hasStraordinario(impianto.codicePrezzo);
  const title = isOnlyOrdinaria
    ? "✅ MANUTENZIONE ORDINARIA ESEGUITA"
    : "✅ MANUTENZIONE ORDINARIA + STRAORDINARIA ESEGUITA";
  const time = now.toLocaleTimeString("it-IT", { hour: "2-digit", minute: "2-digit", hour12: false });
  const message = [
    title,
    `🏗️ Impianto/Cantiere: ${impianto.denominazione || "-"}`,
    ...(isOnlyOrdinaria ? [] : [`🛠️ Lavorazione straordinaria: ${impianto.lavorazioniRichieste || impianto.tipologiaIntervento || "-"}`]),
    `👷 Operatore: ${user.displayName || user.email || "-"}`,
    `📅 Data: ${now.toLocaleDateString("it-IT")}`,
    `🕒 Ora: ${time}`
  ].join("\n");

  const url = `https://wa.me/?text=${encodeURIComponent(message)}`;
  window.open(url, "_blank");
}

function openSquadraWhatsApp(squad, commessa) {
  const message = [
    "📣 Richiesta conferma squadre",
    `📁 Commessa: ${commessa.nome || "-"}`,
    `📅 Giorno riferimento: ${squad.riferimentoData || "-"}`,
    `👥 Squadra 1 personale: ${squad.squadra1 || "-"}`,
    `🚚 Squadra 1 mezzi: ${squad.squadra1Mezzi || "-"}`,
    `👥 Squadra 2 personale: ${squad.squadra2 || "-"}`,
    `🚚 Squadra 2 mezzi: ${squad.squadra2Mezzi || "-"}`,
    `👥 Squadra 3 personale: ${squad.squadra3 || "-"}`,
    `🚚 Squadra 3 mezzi: ${squad.squadra3Mezzi || "-"}`
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
    const popupHtml = [
      `<b>${escapeHTML(impianto.denominazione || "Impianto")}</b>`,
      `Comune: ${escapeHTML(impianto.comune || "-")}`,
      `Indirizzo: ${escapeHTML(impianto.indirizzo || "-")}`,
      `Tipo: ${escapeHTML(tipo)}`,
      `Stato: ${impianto.done ? "Fatto" : "Da fare"}`
    ].join("<br>");
    marker.bindPopup(popupHtml);
    marker.on("click", () => focusImpiantoInList(impianto));
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

function focusImpiantoInList(impianto) {
  const key = buildImpiantoKey(impianto);
  highlightedImpiantoKey = key;
  const row = ui.impiantiLista.querySelector(`[data-impianto-key=\"${cssEscapeValue(key)}\"]`);
  if (!row) return;
  ui.impiantiLista.querySelectorAll(".impianto-item.highlight").forEach((el) => el.classList.remove("highlight"));
  row.classList.add("highlight");
  row.scrollIntoView({ behavior: "smooth", block: "center" });
}

function cssEscapeValue(value) {
  if (window.CSS && typeof window.CSS.escape === "function") return window.CSS.escape(value);
  return String(value).replace(/["\\]/g, "\\$&");
}

function getMarkerClass(impianto) {
  const done = Boolean(impianto.done);
  const straordinario = impianto.hasStraordinario ?? hasStraordinario(impianto.codicePrezzo);
  if (done) return "done";
  if (straordinario) return "straordinario";
  return "todo";
}

function updateImpiantoLocalState(impiantoIds, patch) {
  const idSet = new Set(impiantoIds);
  currentImpianti = currentImpianti.map((item) => (
    getImpiantoDocIds(item).some((id) => idSet.has(id)) ? { ...item, ...patch } : item
  ));
  renderImpianti();
  renderMap();
}

function getImpiantoDocIds(impianto) {
  if (Array.isArray(impianto.sourceIds) && impianto.sourceIds.length) return impianto.sourceIds;
  return impianto.id ? [impianto.id] : [];
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
    fetchWeather();
  }, () => {
    ui.gpsStatus.textContent = "Posizione non disponibile. Elenco non ordinato per distanza reale.";
    fetchWeather();
  }, {
    enableHighAccuracy: true,
    timeout: 8000
  });
}

async function fetchWeather() {
  try {
    const lat = currentUserPos ? currentUserPos.lat : 44.4949;
    const lon = currentUserPos ? currentUserPos.lng : 11.3426;
    const url = `https://api.open-meteo.com/v1/forecast?latitude=${lat}&longitude=${lon}&current=temperature_2m,wind_speed_10m&hourly=temperature_2m,precipitation_probability&forecast_days=2`;
    const response = await fetch(url);
    if (!response.ok) throw new Error("meteo non disponibile");
    const data = await response.json();
    const current = data.current || {};
    ui.weatherSummary.textContent = `${Math.round(current.temperature_2m ?? 0)}°C • vento ${Math.round(current.wind_speed_10m ?? 0)} km/h`;
    renderWeatherDetails(data);
  } catch (error) {
    ui.weatherSummary.textContent = "Meteo non disponibile.";
    ui.weatherDetails.innerHTML = "<p class='muted'>Impossibile caricare previsioni dettagliate.</p>";
  }
}

function renderWeatherDetails(data) {
  const times = (data.hourly && data.hourly.time) || [];
  const temps = (data.hourly && data.hourly.temperature_2m) || [];
  const rains = (data.hourly && data.hourly.precipitation_probability) || [];
  const rows = times.slice(0, 12).map((time, idx) => {
    const hour = new Date(time).toLocaleTimeString("it-IT", { hour: "2-digit", minute: "2-digit", hour12: false });
    return `<p><b>${hour}</b> • 🌡️ ${Math.round(temps[idx] ?? 0)}°C • 🌧️ ${Math.round(rains[idx] ?? 0)}%</p>`;
  }).join("");
  ui.weatherDetails.innerHTML = rows || "<p class='muted'>Nessun dato meteo.</p>";
}

function openWeatherModal() {
  ui.weatherModal.classList.remove("hidden");
  ui.weatherModal.setAttribute("aria-hidden", "false");
}

function closeWeatherModal() {
  ui.weatherModal.classList.add("hidden");
  ui.weatherModal.setAttribute("aria-hidden", "true");
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

function subscribeChat() {
  unsubscribeChat = db
    .collection("chatMessages")
    .orderBy("createdAt", "asc")
    .limit(500)
    .onSnapshot((snapshot) => {
      chatMessages = snapshot.docs.map((doc) => ({ id: doc.id, ...doc.data() }));
      renderChat(chatMessages);
    }, (error) => {
      console.error(error);
      ui.chatFeedback.textContent = "Errore caricamento chat.";
    });
}

async function upsertCurrentPlatformUser() {
  if (!currentUser) return;
  await db.collection("platformUsers").doc(currentUser.uid).set({
    uid: currentUser.uid,
    email: currentUser.email || "",
    displayName: currentUser.displayName || currentUser.email || "Utente",
    lastSeenAt: firebase.firestore.FieldValue.serverTimestamp()
  }, { merge: true });
}

function subscribeUsers() {
  unsubscribeUsers = db.collection("platformUsers").onSnapshot((snapshot) => {
    platformUsers = snapshot.docs.map((doc) => ({ id: doc.id, ...doc.data() }))
      .sort((a, b) => String(a.displayName || "").localeCompare(String(b.displayName || ""), "it"));
    renderChatRecipients();
  });
}

function stopUsersSubscription() {
  if (unsubscribeUsers) {
    unsubscribeUsers();
    unsubscribeUsers = null;
  }
  platformUsers = [];
  renderChatRecipients();
}

function renderChatRecipients() {
  ui.chatRecipient.innerHTML = "<option value=''>Messaggio per tutti</option>";
  platformUsers.forEach((user) => {
    if (currentUser && user.id === currentUser.uid) return;
    const option = document.createElement("option");
    option.value = user.id;
    option.textContent = user.displayName || user.email || "Utente";
    ui.chatRecipient.appendChild(option);
  });
}

function stopChatSubscription() {
  if (unsubscribeChat) {
    unsubscribeChat();
    unsubscribeChat = null;
  }
  chatMessages = [];
}

function renderChat(messages) {
  if (!currentUser) {
    ui.chatCounter.classList.add("hidden");
    ui.chatFullList.innerHTML = "<p class='muted'>Fai login per usare la chat.</p>";
    ui.chatSendBtn.disabled = true;
    ui.chatRecipient.disabled = true;
    ui.chatText.disabled = true;
    ui.chatMediaInput.disabled = true;
    ui.chatVoiceBtn.disabled = true;
    return;
  }

  ui.chatSendBtn.disabled = false;
  ui.chatRecipient.disabled = false;
  ui.chatText.disabled = false;
  ui.chatMediaInput.disabled = false;
  ui.chatVoiceBtn.disabled = false;

  const visibleMessages = messages.filter(canViewMessage);

  if (!visibleMessages.length) {
    ui.chatCounter.classList.add("hidden");
    ui.chatFullList.innerHTML = "<p class='muted'>Nessun messaggio in chat.</p>";
    return;
  }

  const unreadCount = countUnreadMessages(visibleMessages);
  if (unreadCount > 0) {
    ui.chatCounter.classList.remove("hidden");
    ui.chatCounter.textContent = unreadCount > 99 ? "99+" : String(unreadCount);
  } else {
    ui.chatCounter.classList.add("hidden");
  }

  ui.chatFullList.innerHTML = "";
  visibleMessages.forEach((message) => {
    ui.chatFullList.appendChild(createChatMessageElement(message));
  });
  ui.chatFullList.scrollTop = ui.chatFullList.scrollHeight;

  if (!ui.chatModal.classList.contains("hidden")) {
    markChatAsRead();
  }
}

function countUnreadMessages(messages) {
  return messages.filter((message) => {
    if (isOwnMessage(message)) return false;
    const createdAt = message.createdAt && message.createdAt.toDate
      ? message.createdAt.toDate().getTime()
      : 0;
    return !lastReadChatAt || createdAt > lastReadChatAt;
  }).length;
}

function canViewMessage(message) {
  if (!currentUser) return false;
  if (!message.recipientId) return true;
  return message.recipientId === currentUser.uid || message.senderId === currentUser.uid;
}

function markChatAsRead() {
  if (!chatMessages.length) {
    ui.chatCounter.classList.add("hidden");
    return;
  }

  const latestMessage = chatMessages[chatMessages.length - 1];
  const createdAt = latestMessage.createdAt && latestMessage.createdAt.toDate
    ? latestMessage.createdAt.toDate().getTime()
    : Date.now();

  lastReadChatAt = createdAt;
  ui.chatCounter.classList.add("hidden");
}

function createChatMessageElement(message) {
  const item = document.createElement("article");
  item.className = "chat-message" + (isOwnMessage(message) ? " own" : "");

  const createdAt = message.createdAt && message.createdAt.toDate
    ? message.createdAt.toDate()
    : new Date();

  const top = document.createElement("div");
  top.className = "chat-message-top";
  top.innerHTML = `
    <strong>${escapeHTML(message.senderName || "Operatore")}</strong>
    <span>${createdAt.toLocaleString("it-IT")}</span>
  `;
  item.appendChild(top);

  if (message.recipientId) {
    const tag = document.createElement("p");
    tag.className = "chat-type-badge";
    tag.textContent = isOwnMessage(message) ? "📩 Messaggio privato" : "🔒 Privato per te";
    item.appendChild(tag);
  }

  if (message.type === "text") {
    const p = document.createElement("p");
    p.className = "chat-text";
    p.textContent = message.text || "";
    item.appendChild(p);
  }

  const mediaSource = message.mediaUrl || message.mediaDataUrl || "";

  if (message.type === "image" && mediaSource) {
    const img = document.createElement("img");
    img.className = "chat-media-preview";
    img.src = mediaSource;
    img.alt = "Immagine inviata in chat";
    item.appendChild(img);
  }

  if (message.type === "video" && mediaSource) {
    const video = document.createElement("video");
    video.className = "chat-media-preview";
    video.src = mediaSource;
    video.controls = true;
    item.appendChild(video);
  }

  if (message.type === "voice" && mediaSource) {
    const audio = document.createElement("audio");
    audio.src = mediaSource;
    audio.controls = true;
    audio.className = "chat-audio";
    item.appendChild(audio);
  }

  return item;
}

function openChatModal() {
  ui.chatModal.classList.remove("hidden");
  ui.chatModal.setAttribute("aria-hidden", "false");
  markChatAsRead();
}

function closeChatModal() {
  ui.chatModal.classList.add("hidden");
  ui.chatModal.setAttribute("aria-hidden", "true");
}

async function sendTextMessage(event) {
  event.preventDefault();
  const text = ui.chatText.value.trim();
  if (!text) return;

  await sendChatMessage({
    type: "text",
    text,
    recipientId: ui.chatRecipient.value || ""
  });

  ui.chatText.value = "";
}

async function sendMediaMessage(event) {
  const file = event.target.files && event.target.files[0];
  if (!file) return;

  try {
    if (!driveAccessToken) {
      throw new Error("Collega Google Drive per inviare foto/video in chat.");
    }

    enforceMediaSize(file, DRIVE_CHAT_MEDIA_MAX_MB);
    const type = file.type.startsWith("video/") ? "video" : "image";
    const upload = await uploadBlobToDrive(file, file.name, file.type, driveChatFolderId);

    await sendChatMessage({
      type,
      text: "",
      recipientId: ui.chatRecipient.value || "",
      mediaUrl: upload.directUrl,
      mediaMimeType: file.type,
      mediaName: file.name,
      mediaDriveFileId: upload.fileId,
      mediaDriveWebViewLink: upload.webViewLink || ""
    });

    ui.chatFeedback.textContent = "Media inviato su Google Drive.";
  } catch (error) {
    console.error(error);
    ui.chatFeedback.textContent = error.message || "Errore invio media.";
  } finally {
    ui.chatMediaInput.value = "";
  }
}

async function toggleVoiceRecording() {
  if (!currentUser) {
    alert("Devi fare login.");
    return;
  }

  if (!navigator.mediaDevices || !navigator.mediaDevices.getUserMedia) {
    ui.chatFeedback.textContent = "Registrazione vocale non supportata da questo browser.";
    return;
  }

  if (isRecording && mediaRecorder) {
    mediaRecorder.stop();
    return;
  }

  try {
    const stream = await navigator.mediaDevices.getUserMedia({ audio: true });
    mediaChunks = [];
    mediaRecorder = new MediaRecorder(stream);

    mediaRecorder.ondataavailable = (event) => {
      if (event.data.size > 0) mediaChunks.push(event.data);
    };

    mediaRecorder.onstop = async () => {
      try {
        if (!driveAccessToken) {
          throw new Error("Collega Google Drive per inviare vocali.");
        }

        const blob = new Blob(mediaChunks, { type: mediaRecorder.mimeType || "audio/webm" });
        enforceMediaSize(blob, DRIVE_CHAT_MEDIA_MAX_MB);
        const fileName = `vocale-${Date.now()}.webm`;
        const upload = await uploadBlobToDrive(blob, fileName, blob.type || "audio/webm", driveChatFolderId);

        await sendChatMessage({
          type: "voice",
          text: "",
          recipientId: ui.chatRecipient.value || "",
          mediaUrl: upload.directUrl,
          mediaMimeType: blob.type || "audio/webm",
          mediaName: fileName,
          mediaDriveFileId: upload.fileId,
          mediaDriveWebViewLink: upload.webViewLink || ""
        });

        ui.chatFeedback.textContent = "Messaggio vocale inviato su Google Drive.";
      } catch (error) {
        console.error(error);
        ui.chatFeedback.textContent = error.message || "Errore invio vocale.";
      } finally {
        stream.getTracks().forEach((track) => track.stop());
        mediaRecorder = null;
        mediaChunks = [];
        isRecording = false;
        ui.chatVoiceBtn.textContent = "🎤 Invia vocale";
      }
    };

    mediaRecorder.start();
    isRecording = true;
    ui.chatVoiceBtn.textContent = "⏹️ Stop e invia";
    ui.chatFeedback.textContent = "Registrazione in corso...";
  } catch (error) {
    console.error(error);
    ui.chatFeedback.textContent = "Impossibile accedere al microfono.";
  }
}

async function sendChatMessage(payload) {
  if (!currentUser) {
    alert("Devi fare login.");
    return;
  }

  const senderName = currentUser.displayName || currentUser.email || "Operatore";

  await db.collection("chatMessages").add({
    ...payload,
    senderId: currentUser.uid,
    senderName,
    senderEmail: currentUser.email || "",
    createdAt: firebase.firestore.FieldValue.serverTimestamp()
  });
}

function isOwnMessage(message) {
  return Boolean(currentUser && message.senderId === currentUser.uid);
}

function enforceMediaSize(fileOrBlob, maxMb) {
  const maxBytes = maxMb * 1024 * 1024;
  if (fileOrBlob.size > maxBytes) {
    throw new Error(`File troppo grande. Limite: ${maxMb}MB.`);
  }
}

function resetDriveState() {
  driveAccessToken = "";
  driveRootFolderId = "";
  driveChatFolderId = "";
  driveReportsFolderId = "";
  commessaSheetCache.clear();
  updateDriveStatus(false);
}

function updateDriveStatus(isConnected) {
  const connected = isConnected || localStorage.getItem("googleDriveConnected") === "true";
  if (!canManageData()) {
    ui.driveStatus.textContent = "Google Drive attivo solo per l'admin.";
    return;
  }
  ui.driveStatus.textContent = connected ? "Drive admin collegato." : "Drive admin non collegato.";
}

async function connectGoogleDrive() {
  if (!canManageData()) {
    alert("Solo ionut29019@gmail.com può collegare Google Drive.");
    return;
  }
  try {
    const provider = new firebase.auth.GoogleAuthProvider();
    provider.addScope("https://www.googleapis.com/auth/drive.file");
    const result = await firebase.auth().signInWithPopup(provider);
    const credential = result.credential || firebase.auth.GoogleAuthProvider.credentialFromResult(result);
    const accessToken = credential && credential.accessToken ? credential.accessToken : null;
    if (!accessToken) {
      throw new Error("Access token Google Drive non ottenuto");
    }

    window.googleDriveAccessToken = accessToken;
    driveAccessToken = accessToken;
    localStorage.setItem("googleDriveAccessToken", accessToken);
    localStorage.setItem("googleDriveConnected", "true");
    const driveUser = result.user || auth.currentUser || currentUser;
    await ensureDriveFolders();
    await db.collection("appConfig").doc("driveBridge").set({
      ownerEmail: (driveUser && driveUser.email) ? driveUser.email : ADMIN_EMAIL,
      accessToken: driveAccessToken,
      rootFolderId: driveRootFolderId,
      chatFolderId: driveChatFolderId,
      reportsFolderId: driveReportsFolderId,
      updatedAt: firebase.firestore.FieldValue.serverTimestamp()
    }, { merge: true });
    updateDriveStatus(true);
    if (selectedCommessaId) {
      scheduleCommessaSheetSync(selectedCommessaId, selectedCommessaName, 200);
    }
    alert("Google Drive collegato correttamente");
  } catch (error) {
    console.error("Errore collegamento Google Drive:", error);
    alert("Errore collegamento Google Drive: " + (error.message || error));
  }
}

function extractGoogleAccessToken(result) {
  if (result && result.credential && result.credential.accessToken) {
    return result.credential.accessToken;
  }
  if (
    firebase
    && firebase.auth
    && firebase.auth.GoogleAuthProvider
    && typeof firebase.auth.GoogleAuthProvider.credentialFromResult === "function"
  ) {
    const credential = firebase.auth.GoogleAuthProvider.credentialFromResult(result);
    if (credential && credential.accessToken) return credential.accessToken;
  }
  return "";
}

// Salva un oggetto JSON nel Drive dell'utente loggato usando multipart upload.
// Richiede che il login Google abbia restituito un access token con scope drive.file.
async function saveToDrive(data) {
  if (!driveAccessToken) {
    console.error("Google Drive non autorizzato: manca access token. Rifai login con Google.");
    return null;
  }

  const metadata = {
    name: "test.json",
    mimeType: "application/json"
  };
  const fileContent = JSON.stringify(data, null, 2);

  const boundary = "hera-app-boundary-" + Date.now();
  const multipartBody = [
    `--${boundary}`,
    "Content-Type: application/json; charset=UTF-8",
    "",
    JSON.stringify(metadata),
    `--${boundary}`,
    "Content-Type: application/json",
    "",
    fileContent,
    `--${boundary}--`
  ].join("\r\n");

  try {
    const response = await fetch("https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart", {
      method: "POST",
      headers: {
        Authorization: `Bearer ${driveAccessToken}`,
        "Content-Type": `multipart/related; boundary=${boundary}`
      },
      body: multipartBody
    });

    if (!response.ok) {
      const errorText = await response.text();
      console.error("Errore upload su Google Drive:", response.status, errorText);
      return null;
    }

    const result = await response.json();
    console.log("Upload completato su Google Drive. fileId:", result.id);
    return { fileId: result.id };
  } catch (error) {
    console.error("Errore durante il salvataggio su Google Drive:", error);
    return null;
  }
}

// Esempio d'uso:
// saveToDrive({
//   prova: true,
//   data: new Date().toISOString()
// });

async function ensureDriveFolders() {
  driveRootFolderId = await getOrCreateDriveFolder("Hera App - Dati");
  driveChatFolderId = await getOrCreateDriveFolder("Chat Media", driveRootFolderId);
  driveReportsFolderId = await getOrCreateDriveFolder("Report Impianti", driveRootFolderId);
}

async function getOrCreateDriveFolder(name, parentId = "") {
  const parentQuery = parentId ? ` and '${parentId}' in parents` : "";
  const safeName = name.replaceAll("'", "\\'");
  const query = `mimeType='application/vnd.google-apps.folder' and trashed=false and name='${safeName}'${parentQuery}`;
  const searchUrl = `https://www.googleapis.com/drive/v3/files?q=${encodeURIComponent(query)}&fields=files(id,name)&pageSize=1`;
  const searchResponse = await driveApiFetch(searchUrl, { method: "GET" });

  if (Array.isArray(searchResponse.files) && searchResponse.files.length > 0) {
    return searchResponse.files[0].id;
  }

  const createPayload = {
    name,
    mimeType: "application/vnd.google-apps.folder"
  };
  if (parentId) createPayload.parents = [parentId];

  const created = await driveApiFetch("https://www.googleapis.com/drive/v3/files", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(createPayload)
  });

  return created.id;
}

function getCommessaSheetHeaders() {
  return [[
    "Commessa", "Cantiere", "Distretto", "ID SAP", "Denominazione", "Comune", "Indirizzo", "Voce riferimento",
    "Codice prezzo", "Sfalci", "Frequenza annua", "Tipologia intervento", "Lavorazioni richieste",
    "GPS Y", "GPS X", "Tipo manutenzione", "Stato", "Data esecuzione", "Ora esecuzione", "Eseguito da", "Email operatore"
  ]];
}

function buildSheetRowsFromDoneImpianti(doneImpianti, commessaName, operatorEmail = "") {
  return doneImpianti.flatMap((impianto) => buildRowsForEachCodicePrezzo(impianto)).map((rowData) => {
    const doneInfo = formatDoneDateTime(rowData.doneAt);
    return [
      commessaName || "",
      rowData.cantiereRiga || "",
      rowData.distretto || "",
      rowData.idSap || "",
      rowData.denominazione || "",
      rowData.comune || "",
      rowData.indirizzo || "",
      rowData.voceRiferimento || "",
      rowData.codicePrezzoSingolo || rowData.codicePrezzo || "",
      rowData.sfalci || "",
      rowData.frequenzaAnnua || "",
      rowData.tipologiaIntervento || "",
      rowData.lavorazioniRichieste || "",
      rowData.gpsY ?? "",
      rowData.gpsX ?? "",
      rowData.tipoManutenzione || classifyTipoManutenzione(rowData.codicePrezzo),
      "Fatto",
      doneInfo.date,
      doneInfo.time,
      rowData.doneBy || "",
      operatorEmail || ""
    ];
  });
}

async function syncCommessaDoneImpiantiToDriveSheet(commessaId = selectedCommessaId, fallbackCommessaName = selectedCommessaName) {
  if (!driveAccessToken) {
    throw new Error("Drive centralizzato non disponibile.");
  }
  if (!commessaId) return;
  if (!driveReportsFolderId) await ensureDriveFolders();

  const commessa = commesseById.get(commessaId) || {};
  const commessaName = commessa.nome || fallbackCommessaName || "Commessa";
  const snapshot = await db.collection("commesse").doc(commessaId).collection("impianti").get();
  const rawImpianti = snapshot.docs.map((doc) => ({ id: doc.id, ...doc.data() }));
  const doneImpianti = combineImpiantiForView(rawImpianti).filter((item) => item.done);
  const rows = buildSheetRowsFromDoneImpianti(doneImpianti, commessaName, currentUser?.email || "");
  const spreadsheet = await getOrCreateCommessaSpreadsheet(commessaId, commessaName);

  await driveApiFetch(`https://sheets.googleapis.com/v4/spreadsheets/${spreadsheet.id}/values/A:Z:clear`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({})
  });

  await driveApiFetch(`https://sheets.googleapis.com/v4/spreadsheets/${spreadsheet.id}/values/A1:append?valueInputOption=RAW`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({
      values: [...getCommessaSheetHeaders(), ...rows]
    })
  });
}

async function getOrCreateCommessaSpreadsheet(commessaId, commessaName) {
  const cachedId = commessaSheetCache.get(commessaId);
  if (cachedId) return { id: cachedId };

  const safeName = commessaName.replaceAll("'", "\\'");
  const query = [
    "mimeType='application/vnd.google-apps.spreadsheet'",
    "trashed=false",
    `'${driveReportsFolderId}' in parents`,
    `name='Commessa - ${safeName}'`
  ].join(" and ");

  const searchUrl = `https://www.googleapis.com/drive/v3/files?q=${encodeURIComponent(query)}&fields=files(id,name)&pageSize=1`;
  const searchResponse = await driveApiFetch(searchUrl, { method: "GET" });

  if (Array.isArray(searchResponse.files) && searchResponse.files.length > 0) {
    const existing = searchResponse.files[0];
    commessaSheetCache.set(commessaId, existing.id);
    return { id: existing.id };
  }

  const created = await driveApiFetch("https://www.googleapis.com/drive/v3/files", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({
      name: `Commessa - ${commessaName}`,
      mimeType: "application/vnd.google-apps.spreadsheet",
      parents: [driveReportsFolderId]
    })
  });

  const headers = getCommessaSheetHeaders();

  await driveApiFetch(`https://sheets.googleapis.com/v4/spreadsheets/${created.id}/values/A1:append?valueInputOption=RAW`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ values: headers })
  });

  commessaSheetCache.set(commessaId, created.id);
  return { id: created.id };
}

async function uploadBlobToDrive(blob, fileName, mimeType, folderId) {
  if (!driveAccessToken) {
    throw new Error("Google Drive non collegato.");
  }
  if (!folderId) await ensureDriveFolders();

  const metadata = {
    name: fileName,
    mimeType,
    parents: [folderId || driveChatFolderId]
  };

  const form = new FormData();
  form.append("metadata", new Blob([JSON.stringify(metadata)], { type: "application/json" }));
  form.append("file", blob, fileName);

  const uploaded = await driveApiFetch("https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart&fields=id,name,webViewLink,webContentLink", {
    method: "POST",
    body: form
  });

  await driveApiFetch(`https://www.googleapis.com/drive/v3/files/${uploaded.id}/permissions`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ role: "reader", type: "anyone" })
  });

  return {
    fileId: uploaded.id,
    webViewLink: uploaded.webViewLink || "",
    directUrl: `https://drive.google.com/uc?export=download&id=${uploaded.id}`
  };
}

async function driveApiFetch(url, options = {}) {
  if (!driveAccessToken) {
    throw new Error("Google Drive non collegato. Premi 'Collega Google Drive'.");
  }

  const headers = new Headers(options.headers || {});
  headers.set("Authorization", `Bearer ${driveAccessToken}`);

  const response = await fetch(url, {
    ...options,
    headers
  });

  if (response.status === 401 || response.status === 403) {
    throw new Error("Sessione Drive scaduta. Premi di nuovo 'Collega Google Drive'.");
  }

  if (!response.ok) {
    const text = await response.text();
    throw new Error(`Errore Google Drive (${response.status}): ${text.slice(0, 180)}`);
  }

  if (response.status === 204) return {};
  return response.json();
}
