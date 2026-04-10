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
  mapFullscreenBtn: document.getElementById("map-fullscreen-btn"),
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
  squadraRows: document.getElementById("squadra-rows"),
  addSquadraRowBtn: document.getElementById("add-squadra-row-btn"),
  squadraRiferimento: document.getElementById("squadra-riferimento"),
  squadraCalendarDate: document.getElementById("squadra-calendar-date"),
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
  openSegnalazioniBtn: document.getElementById("open-segnalazioni-btn"),
  managementPage: document.getElementById("management-page"),
  managementTitle: document.getElementById("management-title"),
  managementCloseBtn: document.getElementById("management-close-btn"),
  panelCommesse: document.getElementById("panel-commesse"),
  panelSquadre: document.getElementById("panel-squadre"),
  panelPersonale: document.getElementById("panel-personale"),
  panelMezzi: document.getElementById("panel-mezzi"),
  personaleOptions: document.getElementById("personale-options"),
  mezziOptions: document.getElementById("mezzi-options"),
  weatherCard: document.getElementById("weather-card"),
  userToggleBtn: document.getElementById("user-toggle-btn"),
  userDetailsPanel: document.getElementById("user-details-panel"),
  weatherSummary: document.getElementById("weather-summary"),
  weatherModal: document.getElementById("weather-modal"),
  weatherCloseBtn: document.getElementById("weather-close-btn"),
  weatherDetails: document.getElementById("weather-details"),
  fuelPage: document.getElementById("fuel-page"),
  backFromFuelBtn: document.getElementById("back-from-fuel-btn"),
  fuelPageTitle: document.getElementById("fuel-page-title"),
  fuelMap: document.getElementById("fuel-map"),
  fuelStationsList: document.getElementById("fuel-stations-list"),
  fuelMezzoDetailsBtn: document.getElementById("fuel-mezzo-details-btn"),
  fuelMezzoDetailsCard: document.getElementById("fuel-mezzo-details-card"),
  fuelMezzoDetails: document.getElementById("fuel-mezzo-details"),
  segnalazioniPage: document.getElementById("segnalazioni-page"),
  backFromSegnalazioniBtn: document.getElementById("back-from-segnalazioni-btn"),
  segnalazioneForm: document.getElementById("segnalazione-form"),
  segnalazionePreposto: document.getElementById("segnalazione-preposto"),
  segnalazioneData: document.getElementById("segnalazione-data"),
  segnalazioneDataFooter: document.getElementById("segnalazione-data-footer"),
  segnalazioneOra: document.getElementById("segnalazione-ora"),
  segnalazioneCantiere: document.getElementById("segnalazione-cantiere"),
  segnalazioneDescrizione: document.getElementById("segnalazione-descrizione"),
  segnalazionePresaVisione: document.getElementById("segnalazione-presa-visione"),
  segnalazioneFirmaTec: document.getElementById("segnalazione-firma-tec"),
  segnalazioneFirmaPreposto: document.getElementById("segnalazione-firma-preposto"),
  segnalazioneShareWhatsappBtn: document.getElementById("segnalazione-share-whatsapp-btn"),
  segnalazioneShareEmailBtn: document.getElementById("segnalazione-share-email-btn"),
  segnalazioneFeedback: document.getElementById("segnalazione-feedback")
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
let unsubscribeSquadreHistory = null;
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
let driveSquadreFolderId = "";
const commessaSheetCache = new Map();
let commesseById = new Map();
let personaleRecords = [];
let mezziRecords = [];
let squadreByCommessa = new Map();
let squadreHistoryByDate = new Map();
let highlightedImpiantoKey = "";
let expandedImpiantoKey = "";
let impiantiSearchTerm = "";
let impiantiViewMode = "done";
let pendingSheetExports = [];
let sheetRetryTimer = null;
let isProcessingAdminSheetQueue = false;
const commessaSheetSyncTimers = new Map();
let fuelMapInstance = null;
let fuelStationsLayer = null;
let selectedFuelMezzo = null;
let lastSegnalazionePdfBlob = null;
let lastSegnalazionePdfName = "";

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
document.addEventListener("fullscreenchange", () => {
  setTimeout(() => map.invalidateSize(), 60);
});

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
ui.mapFullscreenBtn.addEventListener("click", toggleMapFullscreen);
ui.impiantoSearch.addEventListener("input", onImpiantoSearchInput);
ui.viewDoneBtn.addEventListener("click", () => setImpiantiViewMode("done"));
ui.viewTodoBtn.addEventListener("click", () => setImpiantiViewMode("todo"));
ui.personaleForm.addEventListener("submit", addPersonale);
ui.mezziForm.addEventListener("submit", addMezzo);
ui.squadraForm.addEventListener("submit", saveSquadraComposition);
ui.squadraCommessa.addEventListener("change", autofillSquadraForm);
ui.squadraCalendarDate.addEventListener("change", renderSquadre);
ui.addSquadraRowBtn.addEventListener("click", () => addSquadraRow());
ui.personaleImportBtn.addEventListener("click", importPersonaleFromExcel);
ui.mezziImportBtn.addEventListener("click", importMezziFromExcel);
ui.openPanelCommesse.addEventListener("click", () => openManagementPanel("commesse"));
ui.openPanelSquadre.addEventListener("click", () => openManagementPanel("squadre"));
ui.openPanelPersonale.addEventListener("click", () => openManagementPanel("personale"));
ui.openPanelMezzi.addEventListener("click", () => openManagementPanel("mezzi"));
ui.openSegnalazioniBtn.addEventListener("click", openSegnalazioniPage);
ui.managementCloseBtn.addEventListener("click", closeManagementPanel);
ui.userToggleBtn.addEventListener("click", toggleUserDetailsPanel);
ui.weatherCloseBtn.addEventListener("click", closeWeatherModal);
ui.backFromFuelBtn.addEventListener("click", closeFuelPage);
ui.fuelMezzoDetailsBtn.addEventListener("click", toggleFuelMezzoDetails);
ui.backFromSegnalazioniBtn.addEventListener("click", closeSegnalazioniPage);
ui.segnalazioneForm.addEventListener("submit", generateSegnalazionePdf);
ui.segnalazionePreposto.addEventListener("input", syncSegnalazioneFirmaPreposto);
ui.segnalazioneShareWhatsappBtn.addEventListener("click", () => shareSegnalazione("whatsapp"));
ui.segnalazioneShareEmailBtn.addEventListener("click", () => shareSegnalazione("email"));

addSquadraRow();
initGeolocation();
prefillSegnalazioneDateTime();
applyRoute();
window.addEventListener("hashchange", applyRoute);
loadPendingSheetExports();
startSheetRetryLoop();

function toggleUserDetailsPanel() {
  const isHidden = ui.userDetailsPanel.classList.contains("hidden");
  ui.userDetailsPanel.classList.toggle("hidden", !isHidden);
  ui.userToggleBtn.setAttribute("aria-expanded", String(isHidden));
}

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
  prefillSegnalazioneDateTime();
  syncSegnalazioneFirmaPreposto();

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
  squadreHistoryByDate = new Map();
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
  ui.personaleImportBtn.disabled = !canManage;
  ui.mezziImportBtn.disabled = !canManage;
  ui.importBtn.disabled = !canManage || !auth.currentUser || !selectedCommessaId || pendingRows.length === 0;
  ui.squadraCommessa.disabled = !canManage;
  ui.squadraRiferimento.disabled = !canManage;
  ui.addSquadraRowBtn.disabled = !canManage;
  ui.squadraRows.querySelectorAll("input,button").forEach((el) => { el.disabled = !canManage; });
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

function toggleMapFullscreen() {
  const mapContainer = document.getElementById("map");
  if (!mapContainer) return;
  if (!document.fullscreenElement) {
    mapContainer.requestFullscreen().catch((error) => {
      console.error("Fullscreen mappa non disponibile:", error);
    });
    return;
  }
  document.exitFullscreen().catch((error) => {
    console.error("Uscita fullscreen non disponibile:", error);
  });
}

function applyRoute() {
  const hash = window.location.hash || "";
  const match = hash.match(/^#commessa=([a-zA-Z0-9_-]+)$/);
  const fuelMatch = hash.match(/^#fuel=(.+)$/);
  const showSegnalazioni = hash === "#segnalazioni";
  const commessaIdFromHash = match ? match[1] : "";
  const showFuel = Boolean(fuelMatch);
  const showImpianti = Boolean(commessaIdFromHash && selectedCommessaId === commessaIdFromHash);
  ui.homePage.classList.toggle("hidden", showImpianti || showFuel || showSegnalazioni);
  ui.impiantiPage.classList.toggle("hidden", !showImpianti);
  ui.fuelPage.classList.toggle("hidden", !showFuel);
  ui.segnalazioniPage.classList.toggle("hidden", !showSegnalazioni);
  if (showImpianti) {
    ui.impiantiPageTitle.textContent = `Impianti commessa: ${selectedCommessaName || "Commessa"}`;
    setTimeout(() => map.invalidateSize(), 50);
  }
  if (showFuel) {
    setTimeout(() => {
      if (fuelMapInstance) fuelMapInstance.invalidateSize();
    }, 50);
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

function closeFuelPage() {
  window.location.hash = "";
  applyRoute();
}

function openSegnalazioniPage() {
  prefillSegnalazioneDateTime();
  syncSegnalazioneFirmaPreposto();
  window.location.hash = "segnalazioni";
  applyRoute();
  closeSideMenu();
}

function closeSegnalazioniPage() {
  window.location.hash = "";
  applyRoute();
}

function prefillSegnalazioneDateTime() {
  const now = new Date();
  const dateValue = now.toLocaleDateString("it-IT");
  ui.segnalazioneData.value = dateValue;
  ui.segnalazioneDataFooter.value = dateValue;
  ui.segnalazioneOra.value = now.toLocaleTimeString("it-IT", { hour: "2-digit", minute: "2-digit", hour12: false });
}

function syncSegnalazioneFirmaPreposto() {
  const nameFromInput = (ui.segnalazionePreposto.value || "").trim();
  const fallback = currentUser?.displayName || currentUser?.email || "";
  ui.segnalazioneFirmaPreposto.value = nameFromInput || fallback;
}

function getSegnalazioneData() {
  const selectedTypes = Array.from(document.querySelectorAll("input[name='segnalazione-tipo']:checked"))
    .map((input) => input.value);
  return {
    preposto: (ui.segnalazionePreposto.value || "").trim(),
    data: (ui.segnalazioneData.value || "").trim(),
    ora: (ui.segnalazioneOra.value || "").trim(),
    cantiere: (ui.segnalazioneCantiere.value || "").trim(),
    tipi: selectedTypes,
    descrizione: (ui.segnalazioneDescrizione.value || "").trim(),
    presaVisioneTec: "",
    firmaTec: "",
    firmaPreposto: (ui.segnalazioneFirmaPreposto.value || "").trim()
  };
}

function validateSegnalazioneData(data) {
  const requiredValues = [
    data.preposto,
    data.data,
    data.ora,
    data.cantiere,
    data.descrizione,
    data.firmaPreposto
  ];
  if (requiredValues.some((value) => !value)) return "Compila tutti i campi obbligatori.";
  if (!data.tipi.length) return "Seleziona almeno una voce in 'Segnalazione di'.";
  return "";
}

async function generateSegnalazionePdf(event) {
  event.preventDefault();
  prefillSegnalazioneDateTime();
  syncSegnalazioneFirmaPreposto();
  const data = getSegnalazioneData();
  const validationError = validateSegnalazioneData(data);
  if (validationError) {
    ui.segnalazioneFeedback.textContent = validationError;
    ui.segnalazioneShareWhatsappBtn.disabled = true;
    ui.segnalazioneShareEmailBtn.disabled = true;
    return;
  }
  if (!window.jspdf || !window.jspdf.jsPDF) {
    ui.segnalazioneFeedback.textContent = "Generatore PDF non disponibile.";
    return;
  }
  if (!window.html2canvas) {
    ui.segnalazioneFeedback.textContent = "Motore di acquisizione modulo non disponibile.";
    return;
  }

  const { jsPDF } = window.jspdf;
  const sheetNode = document.querySelector(".segnalazione-sheet");
  if (!sheetNode) {
    ui.segnalazioneFeedback.textContent = "Modulo segnalazione non trovato.";
    return;
  }
  const canvas = await window.html2canvas(sheetNode, {
    scale: 2,
    useCORS: true,
    backgroundColor: "#ffffff"
  });
  const imageData = canvas.toDataURL("image/png");
  const doc = new jsPDF({ orientation: "portrait", unit: "mm", format: "a4" });
  const pageWidth = doc.internal.pageSize.getWidth();
  const pageHeight = doc.internal.pageSize.getHeight();
  const margin = 5;
  const maxWidth = pageWidth - margin * 2;
  const maxHeight = pageHeight - margin * 2;
  const imageRatio = canvas.width / canvas.height;
  let renderWidth = maxWidth;
  let renderHeight = renderWidth / imageRatio;
  if (renderHeight > maxHeight) {
    renderHeight = maxHeight;
    renderWidth = renderHeight * imageRatio;
  }
  const x = (pageWidth - renderWidth) / 2;
  const y = (pageHeight - renderHeight) / 2;
  doc.addImage(imageData, "PNG", x, y, renderWidth, renderHeight, undefined, "FAST");

  const safeDate = data.data.replace(/[^\d]/g, "-");
  lastSegnalazionePdfName = `scheda-segnalazione-${safeDate || "oggi"}.pdf`;
  lastSegnalazionePdfBlob = doc.output("blob");
  doc.save(lastSegnalazionePdfName);

  ui.segnalazioneShareWhatsappBtn.disabled = false;
  ui.segnalazioneShareEmailBtn.disabled = false;
  ui.segnalazioneFeedback.textContent = "PDF creato. Ora puoi condividerlo con WhatsApp o Email.";
}

async function shareSegnalazione(channel) {
  if (!lastSegnalazionePdfBlob) {
    ui.segnalazioneFeedback.textContent = "Prima genera il PDF.";
    return;
  }
  const file = new File([lastSegnalazionePdfBlob], lastSegnalazionePdfName || "scheda-segnalazione.pdf", {
    type: "application/pdf"
  });
  const shareMessage = "Invio scheda segnalazione in PDF.";
  if (navigator.share && navigator.canShare && navigator.canShare({ files: [file] })) {
    try {
      await navigator.share({
        title: "Scheda segnalazione",
        text: shareMessage,
        files: [file]
      });
      ui.segnalazioneFeedback.textContent = "Condivisione completata.";
      return;
    } catch (error) {
      if (error?.name !== "AbortError") {
        ui.segnalazioneFeedback.textContent = "Condivisione annullata o non disponibile su questo dispositivo.";
      }
      return;
    }
  }

  if (channel === "whatsapp") {
    window.open(`https://wa.me/?text=${encodeURIComponent(`${shareMessage} Ho generato il PDF: ${lastSegnalazionePdfName}`)}`, "_blank");
    ui.segnalazioneFeedback.textContent = "WhatsApp aperto. Allega il PDF scaricato prima di inviare.";
    return;
  }

  window.location.href = `mailto:?subject=${encodeURIComponent("Scheda segnalazione PDF")}&body=${encodeURIComponent(`Buongiorno,\n\nin allegato la scheda segnalazione (${lastSegnalazionePdfName}).`)}`;
  ui.segnalazioneFeedback.textContent = "Email aperta. Allega il PDF scaricato prima di inviare.";
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
      driveSquadreFolderId = data.squadreFolderId || "";
      ui.driveStatus.textContent = `Drive centralizzato attivo (${owner}).`;
      processPendingSheetExports();
      processAdminSheetExportQueue();
      return;
    }

    driveAccessToken = "";
    driveRootFolderId = "";
    driveChatFolderId = "";
    driveReportsFolderId = "";
    driveSquadreFolderId = "";
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
  let previousDoneDocIds = null;

  unsubscribeImpianti = db
    .collection("commesse")
    .doc(selectedCommessaId)
    .collection("impianti")
    .onSnapshot((snapshot) => {
      const rawImpianti = snapshot.docs.map((doc) => ({ id: doc.id, ...doc.data() }));
      currentImpianti = combineImpiantiForView(rawImpianti);
      renderImpianti();
      renderMap();

      const currentDoneDocIds = new Set(
        rawImpianti
          .filter((impianto) => Boolean(impianto.done))
          .map((impianto) => impianto.id)
      );

      const hasNewDoneImpianto = previousDoneDocIds !== null
        && Array.from(currentDoneDocIds).some((docId) => !previousDoneDocIds.has(docId));

      if (hasNewDoneImpianto) {
        scheduleCommessaSheetSync(selectedCommessaId, selectedCommessaName, 1200);
      }

      previousDoneDocIds = currentDoneDocIds;
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
    const detailsVisible = expandedImpiantoKey === impiantoKey;
    article.dataset.impiantoKey = impiantoKey;
    article.classList.toggle("is-expanded", detailsVisible);
    if (highlightedImpiantoKey === impiantoKey) article.classList.add("highlight");

    const distance = formatDistance(distanceFromUser(impianto));
    const tipo = impianto.tipoManutenzione || classifyTipoManutenzione(impianto.codicePrezzo);
    const hasStraordinariaFlag = impianto.hasStraordinario ?? hasStraordinario(impianto.codicePrezzo);
    const header = document.createElement("button");
    header.type = "button";
    header.className = "impianto-summary-btn";
    header.innerHTML = `
      <strong>${escapeHTML(impianto.denominazione || "(senza nome)")}</strong>
      <span class="badge ${hasStraordinariaFlag ? "badge-straordinaria" : "badge-ordinaria"}">${escapeHTML(tipo)}</span>
      <small>${distance}</small>
    `;
    header.setAttribute("aria-expanded", detailsVisible ? "true" : "false");
    header.addEventListener("click", () => {
      expandedImpiantoKey = expandedImpiantoKey === impiantoKey ? "" : impiantoKey;
      renderImpianti();
    });
    article.appendChild(header);

    const details = document.createElement("div");
    details.className = "impianto-details";
    details.innerHTML = `
      <p><b>Comune:</b> ${escapeHTML(impianto.comune || "-")}</p>
      <p><b>Indirizzo:</b> ${escapeHTML(impianto.indirizzo || "-")}</p>
      <p><b>Codice prezzo:</b> ${escapeHTML(impianto.codicePrezzo || impianto.voceRiferimento || "-")}</p>
      <p><b>Lavorazioni richieste:</b> ${escapeHTML(impianto.lavorazioniRichieste || impianto.tipologiaIntervento || "-")}</p>
      <p><b>Stato:</b> ${impianto.done ? "Fatto" : "Da fare"}</p>
      <p><b>Eseguito da:</b> ${escapeHTML(impianto.doneBy || "-")}</p>
    `;
    article.appendChild(details);

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
  if (String(label).toLowerCase().includes("whatsapp")) btn.classList.add("btn-whatsapp");
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
  if (!canManageData()) {
    if (!currentUserPos) {
      alert("Per segnare FATTO devi attivare la posizione GPS.");
      return;
    }
    const distanceKm = distanceFromUser(impianto);
    if (!Number.isFinite(distanceKm) || distanceKm > 4) {
      alert("Puoi segnare FATTO solo entro 4 km dall'impianto.");
      return;
    }
  }
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
  if (nome.split(/\s+/).length < 2) {
    alert("Inserisci Nome e Cognome del personale.");
    return;
  }
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
  const file = ui.mezziExcelFile.files && ui.mezziExcelFile.files[0];
  if (!file) {
    alert("Seleziona un file Excel mezzi.");
    return;
  }
  const rows = await parseMezziExcelRows(file);
  if (!rows.length) {
    alert("Nessun mezzo valido trovato nel file.");
    return;
  }
  const batch = db.batch();
  rows.forEach((mezzo) => {
    const ref = db.collection("mezzi").doc();
    batch.set(ref, {
      nome: mezzo.nId || "",
      nId: mezzo.nId || "",
      marca: mezzo.marca || "",
      portataCarico: mezzo.portataCarico || "",
      massaComplessivaKg: mezzo.massaComplessivaKg || "",
      alimentazione: mezzo.alimentazione || "",
      createdAt: firebase.firestore.FieldValue.serverTimestamp()
    });
  });
  await batch.commit();
  ui.mezziExcelFile.value = "";
  alert(`Import mezzi completato (${rows.length} elementi).`);
}

async function parseMezziExcelRows(file) {
  const matrix = await parseSimpleExcelRows(file, true);
  if (!matrix.length) return [];
  const header = matrix[0].map((cell) => normalizeHeaderKey(cell));
  return matrix.slice(1).map((row) => {
    const get = (aliases) => {
      const index = header.findIndex((h) => aliases.includes(h));
      return index >= 0 ? String(row[index] || "").trim() : "";
    };
    return {
      nId: get(["nid", "nid.", "n.id", "n.id.", "id"]),
      marca: get(["marca"]),
      portataCarico: get(["portatacarico", "portata"]),
      massaComplessivaKg: get(["massacomplessivapesodelcamioncaricokg", "massacomplessivakg", "massa"]),
      alimentazione: get(["alimentazione"])
    };
  }).filter((row) => row.nId);
}

function normalizeHeaderKey(value) {
  return String(value || "")
    .toLowerCase()
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/[^a-z0-9]/g, "");
}

async function importSimpleRegistryFromExcel(inputEl, collectionName) {
  const file = inputEl.files && inputEl.files[0];
  if (!file) {
    alert("Seleziona un file Excel.");
    return;
  }

  const rows = await parseSimpleExcelRows(file);
  const uniqueNames = [...new Set(rows.filter(Boolean).map((v) => v.trim()).filter(Boolean))]
    .filter((name) => (collectionName === "personale" ? name.split(/\s+/).length >= 2 : true));
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

async function parseSimpleExcelRows(file, rawRows = false) {
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
          if (rawRows) {
            names.push(row);
            return;
          }
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
    updateSuggestionLists();
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
    updateSuggestionLists();
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

  unsubscribeSquadreHistory = db.collection("squadreStorico").onSnapshot((snapshot) => {
    squadreHistoryByDate = new Map();
    snapshot.forEach((doc) => {
      const data = doc.data() || {};
      if (!data.dateKey) return;
      if (!squadreHistoryByDate.has(data.dateKey)) {
        squadreHistoryByDate.set(data.dateKey, new Map());
      }
      squadreHistoryByDate.get(data.dateKey).set(data.commessaId, { id: doc.id, ...data });
    });
    renderSquadre();
  }, (error) => {
    console.error(error);
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
  if (unsubscribeSquadreHistory) {
    unsubscribeSquadreHistory();
    unsubscribeSquadreHistory = null;
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
    ui.squadraRows.innerHTML = "";
    ui.squadraRiferimento.value = "";
    addSquadraRow();
    return;
  }

  const data = squadreByCommessa.get(commessaId) || {};
  ui.squadraRiferimento.value = data.riferimentoData || "";
  setSquadraRowsFromData(data);
}

function addSquadraRow(rowData = { personale: "", mezzi: "" }) {
  const index = ui.squadraRows.children.length + 1;
  const row = document.createElement("div");
  row.className = "squadra-row";
  row.innerHTML = `
    <div class="squadra-row-head">
      <strong>Squadra ${index}</strong>
      <button type="button" class="btn remove-squadra-btn">Rimuovi</button>
    </div>
    <input type="text" class="squadra-personale" list="personale-options" placeholder="Personale squadra" value="${escapeHTML(rowData.personale || "")}">
    <input type="text" class="squadra-mezzi" list="mezzi-options" placeholder="Mezzi squadra" value="${escapeHTML(rowData.mezzi || "")}">
  `;
  row.querySelector(".remove-squadra-btn").addEventListener("click", () => {
    row.remove();
    renumberSquadraRows();
    if (!ui.squadraRows.children.length) addSquadraRow();
  });
  const personaleInput = row.querySelector(".squadra-personale");
  const mezziInput = row.querySelector(".squadra-mezzi");
  personaleInput.addEventListener("blur", () => {
    personaleInput.value = resolveSuggestionValue(personaleInput.value, personaleRecords.map((p) => p.nome));
  });
  mezziInput.addEventListener("blur", () => {
    const suggestion = resolveSuggestionValue(mezziInput.value, mezziRecords.map((m) => m.nId || m.nome));
    if (suggestion) mezziInput.value = suggestion;
  });
  ui.squadraRows.appendChild(row);
  updateAdminControls();
}

function resolveSuggestionValue(rawValue, sourceValues) {
  const value = String(rawValue || "").trim();
  if (!value) return "";
  const exact = sourceValues.find((item) => String(item || "").toLowerCase() === value.toLowerCase());
  if (exact) return exact;
  const matches = sourceValues.filter((item) => String(item || "").toLowerCase().includes(value.toLowerCase()));
  if (matches.length === 1) return matches[0];
  return value;
}

function renumberSquadraRows() {
  Array.from(ui.squadraRows.children).forEach((row, idx) => {
    const title = row.querySelector(".squadra-row-head strong");
    if (title) title.textContent = `Squadra ${idx + 1}`;
  });
}

function readSquadraRows() {
  return Array.from(ui.squadraRows.querySelectorAll(".squadra-row")).map((row) => ({
    personale: String((row.querySelector(".squadra-personale") || {}).value || "").trim(),
    mezzi: String((row.querySelector(".squadra-mezzi") || {}).value || "").trim()
  })).filter((row) => row.personale || row.mezzi);
}

function getLegacySquadreRows(data) {
  const rows = [];
  if (data.squadra1 || data.squadra1Mezzi) rows.push({ personale: data.squadra1 || "", mezzi: data.squadra1Mezzi || "" });
  if (data.squadra2 || data.squadra2Mezzi) rows.push({ personale: data.squadra2 || "", mezzi: data.squadra2Mezzi || "" });
  if (data.squadra3 || data.squadra3Mezzi) rows.push({ personale: data.squadra3 || "", mezzi: data.squadra3Mezzi || "" });
  return rows;
}

function setSquadraRowsFromData(data) {
  ui.squadraRows.innerHTML = "";
  const rows = Array.isArray(data.squadre) ? data.squadre : getLegacySquadreRows(data);
  if (!rows.length) {
    addSquadraRow();
    return;
  }
  rows.forEach((row) => addSquadraRow(row));
  renumberSquadraRows();
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
  const dateKey = ui.squadraRiferimento.value || new Date().toISOString().slice(0, 10);
  const squadreRows = readSquadraRows();
  const payload = {
    commessaId,
    commessaNome: (commesseById.get(commessaId) || {}).nome || "Commessa",
    riferimentoData: dateKey,
    squadre: squadreRows,
    updatedAt: firebase.firestore.FieldValue.serverTimestamp(),
    updatedBy: (currentUser && currentUser.email) ? currentUser.email : ""
  };
  await db.collection("squadreCommesse").doc(commessaId).set(payload, { merge: true });
  await db.collection("squadreStorico").doc(`${dateKey}__${commessaId}`).set({
    ...payload,
    dateKey
  }, { merge: true });
  await backupSquadreSnapshotToDrive(dateKey, payload);
  ui.squadraCalendarDate.value = dateKey;
}

function renderSquadre() {
  ui.squadreLista.innerHTML = "";
  const selectedDateKey = ui.squadraCalendarDate.value || "";
  const storicoDelGiorno = selectedDateKey ? (squadreHistoryByDate.get(selectedDateKey) || new Map()) : null;

  const commesse = Array.from(commesseById.values());
  const commesseConSquadre = commesse.filter((commessa) => {
    const squad = storicoDelGiorno ? (storicoDelGiorno.get(commessa.id) || {}) : (squadreByCommessa.get(commessa.id) || {});
    const rows = Array.isArray(squad.squadre) ? squad.squadre : getLegacySquadreRows(squad);
    return rows.some((row) => row.personale || row.mezzi);
  });
  if (!commesseConSquadre.length) {
    ui.squadreLista.innerHTML = selectedDateKey
      ? "<p class='muted'>Nessuna composizione trovata per la data selezionata.</p>"
      : "<p class='muted'>Nessuna commessa disponibile.</p>";
    return;
  }

  commesseConSquadre.forEach((commessa) => {
    const item = document.createElement("article");
    item.className = "squadra-item";
    const squad = storicoDelGiorno ? (storicoDelGiorno.get(commessa.id) || {}) : (squadreByCommessa.get(commessa.id) || {});
    const squadRows = Array.isArray(squad.squadre) ? squad.squadre : getLegacySquadreRows(squad);
    const riferimento = squad.riferimentoData
      ? new Date(`${squad.riferimentoData}T00:00:00`).toLocaleDateString("it-IT")
      : "-";
    const rowsHtml = squadRows.map((row, idx) => (
      `<p><b>👥 Squadra ${idx + 1}:</b> ${escapeHTML(row.personale || "-")}<br><b>🚚 Mezzi ${idx + 1}:</b> ${renderMezziButtonsMarkup(row.mezzi)}</p>`
    )).join("");
    item.innerHTML = `
      <strong>📁 ${escapeHTML(commessa.nome || "Commessa senza nome")}</strong>
      <p><b>📅 Giorno:</b> ${escapeHTML(riferimento)}</p>
      ${rowsHtml}
    `;
    const askBtn = createButton("WhatsApp al tecnico", () => openSquadraWhatsApp(squad, commessa));
    item.appendChild(askBtn);
    item.querySelectorAll(".mezzo-chip-btn").forEach((btn) => {
      btn.addEventListener("click", () => openFuelPage(btn.dataset.mezzo || ""));
    });
    ui.squadreLista.appendChild(item);
  });
}

function renderMezziButtonsMarkup(rawValue) {
  const parts = String(rawValue || "")
    .split(/[\s,;|]+/)
    .map((value) => value.trim())
    .filter(Boolean);
  if (!parts.length) return "-";
  return parts.map((mezzo) => `<button type=\"button\" class=\"mezzo-chip-btn\" data-mezzo=\"${escapeHTML(mezzo)}\">${escapeHTML(mezzo)}</button>`).join(" ");
}

function updateSquadraHintFromSources() {
  if (!canManageData()) return;
  ui.squadraHint.textContent = "Scrivi nel campo Personale/Mezzi per ottenere suggerimenti automatici. Usa “Aggiungi squadra” per creare tutte le squadre che vuoi.";
}

function updateSuggestionLists() {
  ui.personaleOptions.innerHTML = "";
  personaleRecords.forEach((person) => {
    const option = document.createElement("option");
    option.value = person.nome || "";
    ui.personaleOptions.appendChild(option);
  });
  ui.mezziOptions.innerHTML = "";
  mezziRecords.forEach((mezzo) => {
    const option = document.createElement("option");
    option.value = mezzo.nId || mezzo.nome || "";
    ui.mezziOptions.appendChild(option);
  });
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
    `${title} - Report operativo`,
    `🏗️ Impianto/Cantiere: ${impianto.denominazione || "-"}`,
    `🆔 ID SAP: ${impianto.idSap || "-"}`,
    ...(isOnlyOrdinaria ? [] : [`🛠️ Lavorazione straordinaria: ${impianto.lavorazioniRichieste || impianto.tipologiaIntervento || "-"}`]),
    `👷 Operatore: ${user.displayName || user.email || "-"}`,
    `📅 Data: ${now.toLocaleDateString("it-IT")}`,
    `🕒 Ora: ${time}`,
    "Messaggio generato automaticamente da Hera App."
  ].join("\n");

  const url = `https://wa.me/?text=${encodeURIComponent(message)}`;
  window.open(url, "_blank");
}

function openSquadraWhatsApp(squad, commessa) {
  const squadRows = Array.isArray(squad.squadre) ? squad.squadre : getLegacySquadreRows(squad);
  const rowsMessage = squadRows.map((row, idx) => (
    `👥 Squadra ${idx + 1} personale: ${row.personale || "-"}\n🚚 Squadra ${idx + 1} mezzi: ${row.mezzi || "-"}`
  )).join("\n");
  const message = [
    "📣 Richiesta di conferma composizione squadre",
    "Gentile tecnico, di seguito la composizione registrata.",
    `📁 Commessa: ${commessa.nome || "-"}`,
    `📅 Giorno riferimento: ${squad.riferimentoData || "-"}`,
    rowsMessage || "Nessuna squadra compilata.",
    "Grazie per la verifica."
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
  expandedImpiantoKey = key;
  renderImpianti();
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

function getMezzoByLabel(label) {
  const normalized = String(label || "").trim().toLowerCase();
  if (!normalized) return null;
  const exact = mezziRecords.find((mezzo) => {
    const nId = String(mezzo.nId || "").toLowerCase();
    const nome = String(mezzo.nome || "").toLowerCase();
    return nId === normalized || nome === normalized;
  });
  if (exact) return exact;
  const byNIdContains = mezziRecords.find((mezzo) => {
    const nId = String(mezzo.nId || "").toLowerCase();
    return nId && (nId.includes(normalized) || normalized.includes(nId));
  });
  return byNIdContains || null;
}

async function openFuelPage(mezzoLabel) {
  selectedFuelMezzo = getMezzoByLabel(mezzoLabel) || { nId: mezzoLabel, nome: mezzoLabel };
  ui.fuelPageTitle.textContent = `Distributori Q8/ENI • ${selectedFuelMezzo.nId || selectedFuelMezzo.nome || "Mezzo"}`;
  ui.fuelMezzoDetailsCard.classList.add("hidden");
  renderFuelMezzoDetails();
  window.location.hash = `fuel=${encodeURIComponent(selectedFuelMezzo.nId || selectedFuelMezzo.nome || "mezzo")}`;
  applyRoute();
  await loadNearbyFuelStations();
}

function toggleFuelMezzoDetails() {
  ui.fuelMezzoDetailsCard.classList.toggle("hidden");
  if (!ui.fuelMezzoDetailsCard.classList.contains("hidden")) {
    ui.fuelMezzoDetailsCard.scrollIntoView({ behavior: "smooth", block: "start" });
  }
}

function renderFuelMezzoDetails() {
  if (!selectedFuelMezzo) {
    ui.fuelMezzoDetails.innerHTML = "<p class='muted'>Nessun mezzo selezionato.</p>";
    return;
  }
  ui.fuelMezzoDetails.innerHTML = `
    <p><b>N. ID:</b> ${escapeHTML(selectedFuelMezzo.nId || selectedFuelMezzo.nome || "-")}</p>
    <p><b>Marca:</b> ${escapeHTML(selectedFuelMezzo.marca || "-")}</p>
    <p><b>Portata (carico):</b> ${escapeHTML(selectedFuelMezzo.portataCarico || "-")}</p>
    <p><b>Massa complessiva (kg):</b> ${escapeHTML(selectedFuelMezzo.massaComplessivaKg || "-")}</p>
    <p><b>Alimentazione:</b> ${escapeHTML(selectedFuelMezzo.alimentazione || "-")}</p>
  `;
}

async function loadNearbyFuelStations() {
  if (!currentUserPos) {
    ui.fuelStationsList.innerHTML = "<p class='muted'>Posizione non disponibile. Attiva GPS per vedere i distributori vicini.</p>";
    return;
  }
  ui.fuelStationsList.innerHTML = "<p class='muted'>Caricamento distributori...</p>";
  try {
    const data = await fetchFuelStationsFromOverpass(currentUserPos.lat, currentUserPos.lng);
    const stations = (data.elements || []).map((item) => {
      const lat = item.lat || (item.center && item.center.lat);
      const lon = item.lon || (item.center && item.center.lon);
      const brand = String((item.tags && (item.tags.brand || item.tags.name)) || "").toLowerCase();
      if (!lat || !lon) return null;
      if (!brand.includes("eni") && !brand.includes("q8")) return null;
      return {
        id: item.id,
        name: item.tags.name || item.tags.brand || "Distributore",
        brand: item.tags.brand || "-",
        lat,
        lon,
        distance: haversine(currentUserPos.lat, currentUserPos.lng, lat, lon)
      };
    }).filter(Boolean).sort((a, b) => a.distance - b.distance).slice(0, 20);
    renderFuelStations(stations);
  } catch (error) {
    console.error("Errore caricamento distributori:", error);
    const retryBtn = createButton("Riprova", () => loadNearbyFuelStations());
    ui.fuelStationsList.innerHTML = "<p class='muted'>Errore caricamento distributori. Riprova tra pochi secondi.</p>";
    ui.fuelStationsList.appendChild(retryBtn);
    if (fuelStationsLayer) fuelStationsLayer.clearLayers();
  }
}

async function fetchFuelStationsFromOverpass(lat, lng) {
  const query = `
    [out:json][timeout:25];
    (
      node[\"amenity\"=\"fuel\"](around:12000,${lat},${lng});
      way[\"amenity\"=\"fuel\"](around:12000,${lat},${lng});
    );
    out center 40;
  `;
  const endpoints = [
    "https://overpass-api.de/api/interpreter",
    "https://overpass.kumi.systems/api/interpreter"
  ];

  let lastError = null;
  for (const endpoint of endpoints) {
    try {
      const controller = new AbortController();
      const timeout = setTimeout(() => controller.abort(), 12000);
      const response = await fetch(endpoint, {
        method: "POST",
        body: query,
        signal: controller.signal
      });
      clearTimeout(timeout);
      if (!response.ok) throw new Error(`Overpass ${response.status}`);
      return await response.json();
    } catch (error) {
      lastError = error;
    }
  }
  throw lastError || new Error("Overpass non disponibile");
}

function ensureFuelMap() {
  if (fuelMapInstance) return;
  fuelMapInstance = L.map("fuel-map");
  L.tileLayer("https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png", {
    maxZoom: 19,
    attribution: "&copy; OpenStreetMap"
  }).addTo(fuelMapInstance);
  fuelStationsLayer = L.layerGroup().addTo(fuelMapInstance);
}

function renderFuelStations(stations) {
  ensureFuelMap();
  fuelStationsLayer.clearLayers();
  ui.fuelStationsList.innerHTML = "";
  if (!stations.length) {
    ui.fuelStationsList.innerHTML = "<p class='muted'>Nessun distributore Q8/ENI trovato vicino a te.</p>";
    return;
  }
  const bounds = [];
  stations.forEach((station) => {
    const marker = L.marker([station.lat, station.lon]).addTo(fuelStationsLayer);
    marker.bindPopup(`<b>${escapeHTML(station.name)}</b><br>${escapeHTML(station.brand)}<br>${formatDistance(station.distance)}`);
    const navBtn = createButton("Naviga", () => window.open(`https://www.google.com/maps/dir/?api=1&destination=${station.lat},${station.lon}`, "_blank"));
    const row = document.createElement("div");
    row.className = "simple-list-item";
    row.innerHTML = `<span><b>${escapeHTML(station.name)}</b><br><small>${escapeHTML(station.brand)} • ${formatDistance(station.distance)}</small></span>`;
    row.appendChild(navBtn);
    ui.fuelStationsList.appendChild(row);
    marker.on("click", () => navBtn.focus());
    bounds.push([station.lat, station.lon]);
  });
  if (currentUserPos) bounds.push([currentUserPos.lat, currentUserPos.lng]);
  fuelMapInstance.fitBounds(bounds, { padding: [24, 24] });
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
  driveSquadreFolderId = "";
  commessaSheetCache.clear();
  updateDriveStatus(false);
}

function updateDriveStatus(isConnected) {
  const connected = isConnected || localStorage.getItem("googleDriveConnected") === "true";
  ui.driveStatus.classList.toggle("status-chip-drive", connected);
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
      squadreFolderId: driveSquadreFolderId,
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
  driveSquadreFolderId = await getOrCreateDriveFolder("Storico Squadre", driveRootFolderId);
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

async function backupSquadreSnapshotToDrive(dateKey, squadraPayload) {
  if (!driveAccessToken) return;
  if (!driveSquadreFolderId) await ensureDriveFolders();
  const exportData = await buildAppBackupPayload(dateKey, squadraPayload);
  const blob = new Blob([JSON.stringify(exportData, null, 2)], { type: "application/json" });
  const commessaLabel = String(squadraPayload.commessaNome || "Commessa").replace(/[^\w\-]+/g, "_");
  const fileName = `squadre_${dateKey}_${commessaLabel}.json`;
  await uploadBlobToDrive(blob, fileName, "application/json", driveSquadreFolderId);
}

async function buildAppBackupPayload(dateKey, squadraPayload) {
  const [commesseSnapshot, personaleSnapshot, mezziSnapshot, squadreCorrentiSnapshot, squadreStoricoSnapshot] = await Promise.all([
    db.collection("commesse").get(),
    db.collection("personale").get(),
    db.collection("mezzi").get(),
    db.collection("squadreCommesse").get(),
    db.collection("squadreStorico").where("dateKey", "==", dateKey).get()
  ]);
  return {
    exportedAt: new Date().toISOString(),
    exportedBy: (currentUser && currentUser.email) ? currentUser.email : "",
    selectedDate: dateKey,
    savedComposition: squadraPayload,
    commesse: commesseSnapshot.docs.map((doc) => ({ id: doc.id, ...doc.data() })),
    personale: personaleSnapshot.docs.map((doc) => ({ id: doc.id, ...doc.data() })),
    mezzi: mezziSnapshot.docs.map((doc) => ({ id: doc.id, ...doc.data() })),
    squadreCorrenti: squadreCorrentiSnapshot.docs.map((doc) => ({ id: doc.id, ...doc.data() })),
    squadreStoricoGiorno: squadreStoricoSnapshot.docs.map((doc) => ({ id: doc.id, ...doc.data() }))
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
