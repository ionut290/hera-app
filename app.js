firebase.initializeApp(firebaseConfig);

const auth = firebase.auth();
const db = firebase.firestore();
const DEFAULT_PUSH_PUBLIC_VAPID_KEY = "BLWYWSC_rEbfAoOnOaO6JYhaYVBCa7IDZaN-2cGMt6uqUYLWwl6mKq8hng9V5B5GPVUOlgjLPLhqz2KvdsuJUoAA";
let firebaseMessaging = null;

db.enablePersistence({ synchronizeTabs: true }).catch((error) => {
  console.warn("Persistenza offline Firestore non disponibile:", error && error.code ? error.code : error);
});

if (firebase.messaging && typeof firebase.messaging === "function") {
  try {
    firebaseMessaging = firebase.messaging();
  } catch (error) {
    console.warn("Firebase Messaging non inizializzato:", error);
  }
}

const ui = {
  menuToggleBtn: document.getElementById("menu-toggle-btn"),
  menuCloseBtn: document.getElementById("menu-close-btn"),
  sideMenu: document.getElementById("side-menu"),
  menuOverlay: document.getElementById("menu-overlay"),
  loginBtn: document.getElementById("login-btn"),
  switchAccountBtn: document.getElementById("switch-account-btn"),
  logoutBtn: document.getElementById("logout-btn"),
  driveConnectBtn: document.getElementById("drive-connect-btn"),
  user: document.getElementById("user"),
  userName: document.getElementById("user-name"),
  driveStatus: document.getElementById("drive-status"),
  pwaNotificationStatus: document.getElementById("pwa-notification-status"),
  enableNotificationsBtn: document.getElementById("enable-notifications-btn"),
  testNotificationBtn: document.getElementById("test-notification-btn"),
  commessaForm: document.getElementById("commessa-form"),
  commessaName: document.getElementById("commessa-name"),
  commesseLista: document.getElementById("commesse-lista"),
  commessaAttiva: document.getElementById("commessa-attiva"),
  commesseNextAction: document.getElementById("commesse-next-action"),
  commessaTargetSelect: document.getElementById("commessa-target-select"),
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
  commessaFocusLabel: document.getElementById("commessa-focus-label"),
  backToHomeBtn: document.getElementById("back-to-home-btn"),
  showNextActionBtn: document.getElementById("show-next-action-btn"),
  impiantiNextAction: document.getElementById("impianti-next-action"),
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
  mezzoNId: document.getElementById("mezzo-n-id"),
  mezzoMarca: document.getElementById("mezzo-marca"),
  mezzoModello: document.getElementById("mezzo-modello"),
  mezzoPortataCarico: document.getElementById("mezzo-portata-carico"),
  mezzoMassaComplessivaKg: document.getElementById("mezzo-massa-complessiva-kg"),
  mezzoAlimentazione: document.getElementById("mezzo-alimentazione"),
  mezziLista: document.getElementById("mezzi-lista"),
  squadraForm: document.getElementById("squadra-form"),
  squadraCommessa: document.getElementById("squadra-commessa"),
  squadraRows: document.getElementById("squadra-rows"),
  addSquadraRowBtn: document.getElementById("add-squadra-row-btn"),
  squadraRiferimento: document.getElementById("squadra-riferimento"),
  squadraCalendarDate: document.getElementById("squadra-calendar-date"),
  squadraHint: document.getElementById("squadra-hint"),
  squadreNextAction: document.getElementById("squadre-next-action"),
  squadreLista: document.getElementById("squadre-lista"),
  squadreWhatsappAllBtn: document.getElementById("squadre-whatsapp-all-btn"),
  squadreFilterControls: document.getElementById("squadre-filter-controls"),
  squadreFilterDate: document.getElementById("squadre-filter-date"),
  squadreFilterClearBtn: document.getElementById("squadre-filter-clear-btn"),
  squadreFilterStatus: document.getElementById("squadre-filter-status"),
  personaleExcelFile: document.getElementById("personale-excel-file"),
  personaleImportBtn: document.getElementById("personale-import-btn"),
  mezziExcelFile: document.getElementById("mezzi-excel-file"),
  mezziImportBtn: document.getElementById("mezzi-import-btn"),
  openPanelCommesse: document.getElementById("open-panel-commesse"),
  openPanelSquadre: document.getElementById("open-panel-squadre"),
  openPanelPersonale: document.getElementById("open-panel-personale"),
  openPanelMezzi: document.getElementById("open-panel-mezzi"),
  openPanelUtenti: document.getElementById("open-panel-utenti"),
  openPanelInfoUtili: document.getElementById("open-panel-info-utili"),
  openPrivateDocsBtn: document.getElementById("open-private-docs-btn"),
  openPersonalServicesBtn: document.getElementById("open-personal-services-btn"),
  openHoursBtn: document.getElementById("open-hours-btn"),
  openSegnalazioniBtn: document.getElementById("open-segnalazioni-btn"),
  openHowtoBtn: document.getElementById("open-howto-btn"),
  managementPage: document.getElementById("management-page"),
  managementTitle: document.getElementById("management-title"),
  managementCloseBtn: document.getElementById("management-close-btn"),
  panelCommesse: document.getElementById("panel-commesse"),
  panelSquadre: document.getElementById("panel-squadre"),
  panelPersonale: document.getElementById("panel-personale"),
  panelMezzi: document.getElementById("panel-mezzi"),
  panelUtenti: document.getElementById("panel-utenti"),
  panelInfoUtili: document.getElementById("panel-info-utili"),
  commesseManageList: document.getElementById("commesse-manage-list"),
  adminUserForm: document.getElementById("admin-user-form"),
  adminUserEmail: document.getElementById("admin-user-email"),
  adminUsersList: document.getElementById("admin-users-list"),
  userPermissionsList: document.getElementById("user-permissions-list"),
  externalAppForm: document.getElementById("external-app-form"),
  externalAppName: document.getElementById("external-app-name"),
  externalAppUrl: document.getElementById("external-app-url"),
  externalAppsList: document.getElementById("external-apps-list"),
  gpsRequestsList: document.getElementById("gps-requests-list"),
  resourceForm: document.getElementById("resource-form"),
  resourceType: document.getElementById("resource-type"),
  resourceTitle: document.getElementById("resource-title"),
  resourceValue: document.getElementById("resource-value"),
  resourceCommesse: document.getElementById("resource-commesse"),
  resourceSubmit: document.getElementById("resource-submit"),
  resourcesList: document.getElementById("resources-list"),
  commessaResourceButtons: document.getElementById("commessa-resource-buttons"),
  commessaResourceViewer: document.getElementById("commessa-resource-viewer"),
  commessaResourceViewerTitle: document.getElementById("commessa-resource-viewer-title"),
  commessaResourceViewerCloseBtn: document.getElementById("commessa-resource-viewer-close-btn"),
  commessaResourceViewerList: document.getElementById("commessa-resource-viewer-list"),
  personaleOptions: document.getElementById("personale-options"),
  mezziOptions: document.getElementById("mezzi-options"),
  weatherCard: document.getElementById("weather-card"),
  activeUsersSummary: document.getElementById("active-users-summary"),
  lastImpiantoActionSummary: document.getElementById("last-impianto-action-summary"),
  nextActionSummary: document.getElementById("next-action-summary"),
  weatherRisks: document.getElementById("weather-risks"),
  userCard: document.getElementById("user-card"),
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
  personalServicesPage: document.getElementById("personal-services-page"),
  backFromPersonalServicesBtn: document.getElementById("back-from-personal-services-btn"),
  personalServicesMap: document.getElementById("personal-services-map"),
  personalServicesPageTitle: document.getElementById("personal-services-page-title"),
  personalServicesListTitle: document.getElementById("personal-services-list-title"),
  personalServicesFeedback: document.getElementById("personal-services-feedback"),
  personalServicesList: document.getElementById("personal-services-list"),
  personalServicesRadius: document.getElementById("personal-services-radius"),
  personalServicesCategories: document.getElementById("personal-services-categories"),
  segnalazioniPage: document.getElementById("segnalazioni-page"),
  backFromSegnalazioniBtn: document.getElementById("back-from-segnalazioni-btn"),
  howtoPage: document.getElementById("howto-page"),
  backFromHowtoBtn: document.getElementById("back-from-howto-btn"),
  howtoFaqList: document.getElementById("howto-faq-list"),
  privateDocsPage: document.getElementById("private-docs-page"),
  backFromPrivateDocsBtn: document.getElementById("back-from-private-docs-btn"),
  hoursPage: document.getElementById("hours-page"),
  backFromHoursBtn: document.getElementById("back-from-hours-btn"),
  hoursForm: document.getElementById("hours-form"),
  hoursDate: document.getElementById("hours-date"),
  hoursCommesseList: document.getElementById("hours-commesse-list"),
  addHoursCommessaBtn: document.getElementById("add-hours-commessa-btn"),
  hoursFinalizeBtn: document.getElementById("hours-finalize-btn"),
  hoursFeedback: document.getElementById("hours-feedback"),
  hoursSummary: document.getElementById("hours-summary"),
  viewHoursBtn: document.getElementById("view-hours-btn"),
  hoursStatsMonth: document.getElementById("hours-stats-month"),
  hoursSavedList: document.getElementById("hours-saved-list"),
  hoursOperatoriOptions: document.getElementById("hours-operatori-options"),
  hoursViewModal: document.getElementById("hours-view-modal"),
  hoursViewCloseBtn: document.getElementById("hours-view-close-btn"),
  hoursTableMonth: document.getElementById("hours-table-month"),
  hoursTableCommessaSelect: document.getElementById("hours-table-commessa-select"),
  hoursTotalOperatorBtn: document.getElementById("hours-total-operator-btn"),
  hoursTotalOperatorCommessaBtn: document.getElementById("hours-total-operator-commessa-btn"),
  hoursTableExportBtn: document.getElementById("hours-table-export-btn"),
  hoursTableFeedback: document.getElementById("hours-table-feedback"),
  hoursTableContainer: document.getElementById("hours-table-container"),
  privateDocsPresetPinBtn: document.getElementById("private-docs-preset-pin-btn"),
  privateDocsPresetTesseraBtn: document.getElementById("private-docs-preset-tessera-btn"),
  privateDocsForm: document.getElementById("private-docs-form"),
  privateDocsName: document.getElementById("private-docs-name"),
  privateDocsNote: document.getElementById("private-docs-note"),
  privateDocsFile: document.getElementById("private-docs-file"),
  privateDocsCamera: document.getElementById("private-docs-camera"),
  privateDocsSaveBtn: document.getElementById("private-docs-save-btn"),
  privateDocsDriveOnly: document.getElementById("private-docs-drive-only"),
  privateDocsFeedback: document.getElementById("private-docs-feedback"),
  privateDocsList: document.getElementById("private-docs-list"),
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
  segnalazioneFeedback: document.getElementById("segnalazione-feedback"),
  manualImpiantoForm: document.getElementById("manual-impianto-form"),
  manualImpiantoDenominazione: document.getElementById("manual-impianto-denominazione"),
  manualImpiantoComune: document.getElementById("manual-impianto-comune"),
  manualImpiantoIndirizzo: document.getElementById("manual-impianto-indirizzo"),
  manualImpiantoCodice: document.getElementById("manual-impianto-codice"),
  manualImpiantoSubmit: document.getElementById("manual-impianto-submit"),
  manualImpiantoFeedback: document.getElementById("manual-impianto-feedback"),
  impiantoEditModal: document.getElementById("impianto-edit-modal"),
  impiantoEditCloseBtn: document.getElementById("impianto-edit-close-btn"),
  impiantoEditForm: document.getElementById("impianto-edit-form"),
  impiantoEditFeedback: document.getElementById("impianto-edit-feedback"),
  impiantoReportModal: document.getElementById("impianto-report-modal"),
  impiantoReportCloseBtn: document.getElementById("impianto-report-close-btn"),
  impiantoReportForm: document.getElementById("impianto-report-form"),
  impiantoReportTitle: document.getElementById("impianto-report-title"),
  impiantoReportText: document.getElementById("impianto-report-text"),
  impiantoReportFeedback: document.getElementById("impianto-report-feedback"),
  editDistretto: document.getElementById("edit-distretto"),
  editIdSap: document.getElementById("edit-id-sap"),
  editDenominazione: document.getElementById("edit-denominazione"),
  editComune: document.getElementById("edit-comune"),
  editIndirizzo: document.getElementById("edit-indirizzo"),
  editVoceRiferimento: document.getElementById("edit-voce-riferimento"),
  editCodicePrezzo: document.getElementById("edit-codice-prezzo"),
  editFrequenzaAnnua: document.getElementById("edit-frequenza-annua"),
  editTipologiaIntervento: document.getElementById("edit-tipologia-intervento"),
  editLavorazioniRichieste: document.getElementById("edit-lavorazioni-richieste"),
  editSfalci: document.getElementById("edit-sfalci"),
  editGpsY: document.getElementById("edit-gps-y"),
  editGpsX: document.getElementById("edit-gps-x")
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
let unsubscribeAdminUsers = null;
let unsubscribeResources = null;
let unsubscribePrivateDocs = null;
let unsubscribeGpsRequests = null;
let unsubscribeGlobalNotifications = null;
let presenceHeartbeatTimer = null;
let chatMessages = [];
let platformUsers = [];
let deniedImpiantoActions = new Set();
const usedActionKeys = new Set();
let mediaRecorder = null;
let mediaChunks = [];
let isRecording = false;
let lastReadChatAt = null;
let driveAccessToken = "";
let driveRootFolderId = "";
let driveChatFolderId = "";
let driveReportsFolderId = "";
let driveSquadreFolderId = "";
let driveHelpCenterFolderId = "";
let driveTokenRefreshPromise = null;
const commessaSheetCache = new Map();
let commesseById = new Map();
let personaleRecords = [];
let mezziRecords = [];
let squadreByCommessa = new Map();
let squadreHistoryByDate = new Map();
let manualSquadreFilterDateKey = "";
let highlightedImpiantoKey = "";
let expandedImpiantoKey = "";
const expandedImpiantoManagementKeys = new Set();
let impiantiSearchTerm = "";
let impiantiViewMode = "todo";
let pendingSheetExports = [];
let sheetRetryTimer = null;
let isProcessingAdminSheetQueue = false;
const commessaSheetSyncTimers = new Map();
const localSheetMutationAt = new Map();
let fuelMapInstance = null;
let fuelStationsLayer = null;
let selectedFuelMezzo = null;
let personalServicesMapInstance = null;
let personalServicesLayer = null;
let personalServicesResults = [];
let expandedPersonalServiceId = "";
let activePersonalServiceCategory = "";
let lastSegnalazionePdfBlob = null;
let lastSegnalazionePdfName = "";
let resourceRecords = [];
let privateDocsRecords = [];
let hoursDraftEntries = [];
let hoursFinalizeLocked = false;
let hoursTableRowsMap = new Map();
let hoursTableContext = null;
let hoursApprovalRequests = [];
let gpsUpdateRequests = [];
let activeResourceTypeForViewer = "";
let activeResourceManageFilter = "";
let editingImpiantoIds = [];
let reportingImpianto = null;
let chatRetentionTimer = null;
let geolocationWatchId = null;
let activeNearbyImpiantoContext = null;
let globalNotificationsInitialized = false;
const CHAT_RETENTION_MS = 24 * 60 * 60 * 1000;
const PROXIMITY_NEAR_KM = 0.25;
const PROXIMITY_AWAY_KM = 0.7;
const GPS_APPROVAL_PHONE = "3892352575";
const HOURS_WHATSAPP_PHONE = "3892352575";
const HOWTO_UPDATED_AT = "2026-04-11";
const PERSONAL_SERVICE_CATEGORIES = {
  breakfast: {
    title: "Colazione (bar e caffetterie)",
    icon: "☕",
    query: "node[\"amenity\"~\"^(cafe|bar|pub)$\"](around:{radius},{lat},{lng});way[\"amenity\"~\"^(cafe|bar|pub)$\"](around:{radius},{lat},{lng});relation[\"amenity\"~\"^(cafe|bar|pub)$\"](around:{radius},{lat},{lng});",
    detailFields: ["opening_hours", "cuisine", "takeaway", "delivery", "contact:phone", "website", "outdoor_seating"]
  },
  lunch: {
    title: "Pranzo (ristoranti, mense, circoli ARCI)",
    icon: "🍽️",
    query: "node[\"amenity\"~\"^(restaurant|fast_food|food_court|canteen|biergarten|pub)$\"](around:{radius},{lat},{lng});way[\"amenity\"~\"^(restaurant|fast_food|food_court|canteen|biergarten|pub)$\"](around:{radius},{lat},{lng});relation[\"amenity\"~\"^(restaurant|fast_food|food_court|canteen|biergarten|pub)$\"](around:{radius},{lat},{lng});node[\"club\"=\"social\"](around:{radius},{lat},{lng});way[\"club\"=\"social\"](around:{radius},{lat},{lng});relation[\"club\"=\"social\"](around:{radius},{lat},{lng});node[\"social_facility\"=\"canteen\"](around:{radius},{lat},{lng});way[\"social_facility\"=\"canteen\"](around:{radius},{lat},{lng});",
    detailFields: ["cuisine", "opening_hours", "opening_hours:covid19", "payment:meal_voucher", "payment:sodexo", "payment:edenred", "payment:ticket_restaurant", "payment:cash", "payment:credit_cards", "diet:vegetarian", "diet:vegan", "takeaway", "delivery", "contact:phone", "website", "addr:street", "addr:housenumber", "addr:city"]
  },
  supermarket: {
    title: "Supermarket",
    icon: "🛒",
    query: "node[\"shop\"~\"supermarket|convenience\"](around:{radius},{lat},{lng});way[\"shop\"~\"supermarket|convenience\"](around:{radius},{lat},{lng});",
    detailFields: ["opening_hours", "brand", "contact:phone", "website"]
  },
  tobacco: {
    title: "Tabaccherie",
    icon: "🚬",
    query: "node[\"shop\"=\"tobacco\"](around:{radius},{lat},{lng});way[\"shop\"=\"tobacco\"](around:{radius},{lat},{lng});",
    detailFields: ["opening_hours", "contact:phone", "website"]
  },
  wc: {
    title: "WC pubblici",
    icon: "🚻",
    query: "node[\"amenity\"=\"toilets\"](around:{radius},{lat},{lng});way[\"amenity\"=\"toilets\"](around:{radius},{lat},{lng});",
    detailFields: ["fee", "wheelchair", "opening_hours"]
  },
  atm: {
    title: "Bancomat / ATM",
    icon: "🏧",
    query: "node[\"amenity\"=\"atm\"](around:{radius},{lat},{lng});way[\"amenity\"=\"atm\"](around:{radius},{lat},{lng});",
    detailFields: ["opening_hours", "operator", "cash_in", "contactless", "currency:EUR"]
  },
  pharmacy: {
    title: "Farmacie",
    icon: "💊",
    query: "node[\"amenity\"=\"pharmacy\"](around:{radius},{lat},{lng});way[\"amenity\"=\"pharmacy\"](around:{radius},{lat},{lng});",
    detailFields: ["opening_hours", "dispensing", "contact:phone", "website"]
  },
  parking: {
    title: "Parcheggi",
    icon: "🅿️",
    query: "node[\"amenity\"=\"parking\"](around:{radius},{lat},{lng});way[\"amenity\"=\"parking\"](around:{radius},{lat},{lng});",
    detailFields: ["access", "fee", "capacity", "opening_hours"]
  }
};
const PUSH_PUBLIC_VAPID_KEY = resolvePushPublicVapidKey();
const AUTO_ENABLE_NOTIFICATIONS_KEY = "heraAutoEnableNotifications";
let serviceWorkerRegistration = null;
let hasTriedAutoEnableNotifications = false;

function resolvePushPublicVapidKey() {
  const sources = [
    window?.HERA_PUSH_PUBLIC_VAPID_KEY,
    document.querySelector('meta[name="hera-push-vapid-key"]')?.content,
    localStorage.getItem("heraPushPublicVapidKey"),
    DEFAULT_PUSH_PUBLIC_VAPID_KEY
  ];
  for (const value of sources) {
    if (typeof value === "string" && value.trim()) return value.trim();
  }
  return "";
}

function isAutoNotificationEnabled() {
  const value = localStorage.getItem(AUTO_ENABLE_NOTIFICATIONS_KEY);
  if (value === null) {
    localStorage.setItem(AUTO_ENABLE_NOTIFICATIONS_KEY, "true");
    return true;
  }
  return value === "true";
}

function setAutoNotificationEnabled(enabled) {
  localStorage.setItem(AUTO_ENABLE_NOTIFICATIONS_KEY, enabled ? "true" : "false");
}

async function persistNotificationAutoPreference(enabled) {
  setAutoNotificationEnabled(enabled);
  if (!currentUser) return;
  try {
    await db.collection("platformUsers").doc(currentUser.uid).set({
      notificationsAutoEnabled: enabled,
      updatedAt: firebase.firestore.FieldValue.serverTimestamp(),
      updatedBy: currentUser.email || ""
    }, { merge: true });
  } catch (error) {
    console.warn("Salvataggio preferenza notifiche non riuscito:", error);
  }
}

function syncNotificationAutoPreferenceFromProfile() {
  if (!currentUser) return;
  const row = platformUsers.find((user) => user.id === currentUser.uid);
  if (!row || typeof row.notificationsAutoEnabled !== "boolean") return;
  setAutoNotificationEnabled(row.notificationsAutoEnabled);
}
const MENU_HOWTO_CONTENT = {
  "open-panel-commesse": {
    rispostaBreve: "Da qui gestisci commesse e impianti (aggiunta, import Excel e gestione lista).",
    passi: [
      "Apri il menu (⋮) e premi “Aggiungi commesse”.",
      "Inserisci il nome commessa oppure seleziona una commessa per import/aggiunte impianto.",
      "Usa i form della pagina per completare l'operazione."
    ],
    tags: ["commesse", "impianti", "excel", "admin"]
  },
  "open-panel-squadre": {
    rispostaBreve: "Serve per creare e salvare la composizione giornaliera delle squadre.",
    passi: [
      "Apri il menu (⋮) e premi “Composizione squadre”.",
      "Scegli commessa e data, poi aggiungi le righe squadra.",
      "Premi “Salva composizione” e verifica lo storico sotto al form."
    ],
    tags: ["squadre", "operativo", "personale", "mezzi"]
  },
  "open-panel-personale": {
    rispostaBreve: "Da qui inserisci o importi l'anagrafica personale.",
    passi: [
      "Apri il menu (⋮) e premi “Personale”.",
      "Aggiungi un nominativo singolo o importa il file Excel.",
      "Controlla la lista aggiornata subito sotto."
    ],
    tags: ["personale", "anagrafica", "excel"]
  },
  "open-panel-mezzi": {
    rispostaBreve: "Da qui inserisci o importi l'elenco mezzi disponibili.",
    passi: [
      "Apri il menu (⋮) e premi “Mezzi”.",
      "Aggiungi un mezzo manualmente o importa da Excel.",
      "Controlla che il mezzo compaia in elenco."
    ],
    tags: ["mezzi", "flotta", "excel"]
  },
  "open-panel-utenti": {
    rispostaBreve: "Permette la gestione admin, permessi utente e app collegate.",
    passi: [
      "Apri il menu (⋮) e premi “Gestione utenti”.",
      "Aggiungi/rimuovi admin oppure aggiorna i permessi azione per utente.",
      "Salva le modifiche e verifica l'elenco utenti."
    ],
    tags: ["utenti", "permessi", "admin"]
  },
  "open-panel-info-utili": {
    rispostaBreve: "Consente di pubblicare risorse utili (contatti, note, documenti) per commessa.",
    passi: [
      "Apri il menu (⋮) e premi “Informazioni utili”.",
      "Seleziona tipo risorsa, titolo e contenuto/link.",
      "Salva e verifica che la risorsa sia disponibile nella commessa."
    ],
    tags: ["risorse", "contatti", "note", "documenti"]
  },
  "open-private-docs-btn": {
    rispostaBreve: "Area personale per caricare e consultare documenti individuali.",
    passi: [
      "Apri il menu (⋮) e premi “Documenti personali”.",
      "Compila nome/note e allega file o foto.",
      "Salva e verifica la presenza del documento nell'elenco."
    ],
    tags: ["documenti", "personale", "drive"]
  },
  "open-personal-services-btn": {
    rispostaBreve: "Trovi servizi vicini (colazione, pranzo, market, tabacchi, WC, bancomat e altri) con mappa e navigazione.",
    passi: [
      "Apri il menu (⋮) e premi “Servizi personali”.",
      "Scegli una categoria (es. Colazione o Pranzo).",
      "Apri un luogo dall'elenco o dalla mappa e usa “Naviga” o “Dettagli”."
    ],
    tags: ["servizi", "mappa", "navigazione", "personale"]
  },
  "open-hours-btn": {
    rispostaBreve: "Compili ore per commessa e operatore, salvi il resoconto e invii WhatsApp.",
    passi: [
      "Apri il menu (⋮) e premi “Gestione ore”.",
      "Aggiungi una o più commesse, poi operatori con ore e note.",
      "Premi “Fine: salva e invia” per salvare su Drive e aprire WhatsApp."
    ],
    tags: ["ore", "commesse", "operatori", "whatsapp", "drive"]
  },
  "open-segnalazioni-btn": {
    rispostaBreve: "Compili la segnalazione sicurezza e generi il PDF da condividere.",
    passi: [
      "Apri il menu (⋮) e premi “Segnalazioni”.",
      "Compila i campi obbligatori e scegli il tipo di segnalazione.",
      "Genera il PDF e condividilo via WhatsApp o email."
    ],
    tags: ["segnalazioni", "pdf", "sicurezza"]
  }
};
const STATIC_HOWTO_ITEMS = [
  {
    id: "login-google",
    domanda: "Come faccio il login con Google?",
    rispostaBreve: "Apri il pannello utente e premi “Login con Google”.",
    passi: [
      "Nella home premi l'icona 👤 in alto.",
      "Tocca “Login con Google” e scegli l'account aziendale.",
      "Controlla che compaia “Loggato” con email e nome utente."
    ],
    tags: ["login", "google", "accesso"],
    updatedAt: HOWTO_UPDATED_AT
  },
  {
    id: "chat-operatori",
    domanda: "Come uso la chat operatori?",
    rispostaBreve: "Apri la chat dal pulsante 💬, scrivi e invia il messaggio al destinatario.",
    passi: [
      "Premi il pulsante 💬 in basso a destra.",
      "Scegli un destinatario o lascia “Messaggio per tutti”.",
      "Scrivi il testo (o allega media/vocale) e premi invio."
    ],
    tags: ["chat", "messaggi", "operatori"],
    updatedAt: HOWTO_UPDATED_AT
  },
  {
    id: "google-drive",
    domanda: "Come collego Google Drive?",
    rispostaBreve: "Dopo il login, usa “Collega Google Drive” nel pannello utente.",
    passi: [
      "Esegui login con Google con un account autorizzato.",
      "Apri il pannello utente e premi “Collega Google Drive”.",
      "Conferma i permessi richiesti e verifica lo stato collegato."
    ],
    tags: ["drive", "google", "integrazione"],
    updatedAt: HOWTO_UPDATED_AT
  }
];

const DRIVE_CHAT_MEDIA_MAX_MB = 512;
const ADMIN_EMAIL = "ionut29019@gmail.com";
const IMPIANTO_ACTIONS = ["done", "navigate", "reset", "whatsapp", "problem-report", "gps-update", "edit", "delete"];
let adminEmails = new Set([ADMIN_EMAIL]);
const PENDING_SHEET_EXPORTS_KEY = "heraPendingSheetExports";
const LAST_SELECTED_COMMESSA_KEY = "heraLastSelectedCommessaId";
const LAST_OPENED_COMMESSA_KEY = "heraLastOpenedCommessaId";
const USER_WORKFLOW_STEP_KEY = "heraUserWorkflowStep";
const SHEET_RETRY_MS = 30 * 1000;
const HELP_CENTER_CONFIG_PATH = { collection: "appConfig", doc: "helpCenter" };
const IMPIANTO_NEXT_ACTION_FLOW = ["navigate", "done", "whatsapp"];
const HELP_CENTER_FAQ_FALLBACK = {
  version: 1,
  updatedAt: null,
  updatedBy: "",
  items: [
    {
      id: "faq-import-impianti",
      domanda: "Come importo un file Excel impianti?",
      risposta: "Apri il pannello commesse, seleziona la commessa target, carica il file Excel e conferma l'importazione.",
      passi: ["Apri Gestione commesse", "Seleziona commessa", "Carica file Excel", "Premi Importa"]
    },
    {
      id: "faq-drive-connessione",
      domanda: "Come collego Google Drive?",
      risposta: "Solo l'admin deve cliccare su Collega Google Drive: il bridge viene poi condiviso per tutti gli utenti loggati.",
      passi: ["Login admin", "Premi Collega Google Drive", "Concedi autorizzazioni", "Verifica stato Drive attivo"]
    }
  ]
};
let faqDataset = HELP_CENTER_FAQ_FALLBACK;
let currentWorkflowStepId = localStorage.getItem(USER_WORKFLOW_STEP_KEY) || "";
let impiantoNextActionIndex = 0;
let impiantoNextActionHighlightEnabled = false;
window.googleDriveAccessToken = localStorage.getItem("googleDriveAccessToken") || null;
driveAccessToken = window.googleDriveAccessToken || "";

const map = L.map("map");
L.tileLayer("https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png", {
  maxZoom: 19,
  attribution: "&copy; OpenStreetMap"
}).addTo(map);
map.setView([44.4949, 11.3426], 11);
map.doubleClickZoom.disable();

const markerLayer = L.layerGroup().addTo(map);
const mapElement = document.getElementById("map");
if (mapElement) {
  let lastTapAt = 0;
  mapElement.addEventListener("touchend", () => {
    const now = Date.now();
    if (now - lastTapAt < 320) {
      toggleMapFullscreen();
      lastTapAt = 0;
      return;
    }
    lastTapAt = now;
  }, { passive: true });
  mapElement.addEventListener("dblclick", (event) => {
    event.preventDefault();
    toggleMapFullscreen();
  });
}
document.addEventListener("fullscreenchange", () => {
  syncMapFullscreenUiState();
  setTimeout(() => map.invalidateSize(), 60);
});
document.addEventListener("webkitfullscreenchange", () => {
  syncMapFullscreenUiState();
  setTimeout(() => map.invalidateSize(), 60);
});

ui.loginBtn.addEventListener("click", loginWithGoogle);
ui.switchAccountBtn.addEventListener("click", switchGoogleAccount);
ui.menuToggleBtn.addEventListener("click", openSideMenu);
ui.menuCloseBtn.addEventListener("click", closeSideMenu);
ui.menuOverlay.addEventListener("click", closeSideMenu);
ui.logoutBtn.addEventListener("click", logout);
ui.driveConnectBtn.addEventListener("click", connectGoogleDrive);
ui.commessaForm.addEventListener("submit", createCommessa);
ui.excelFile.addEventListener("change", onExcelSelected);
ui.importBtn.addEventListener("click", importPendingRows);
ui.commessaTargetSelect.addEventListener("change", onCommessaTargetChanged);
ui.chatOpenBtn.addEventListener("click", openChatModal);
ui.chatCloseBtn.addEventListener("click", closeChatModal);
ui.chatSendForm.addEventListener("submit", sendTextMessage);
ui.chatMediaInput.addEventListener("change", sendMediaMessage);
ui.chatVoiceBtn.addEventListener("click", toggleVoiceRecording);
ui.backToHomeBtn.addEventListener("click", closeImpiantiPage);
ui.showNextActionBtn?.addEventListener("click", toggleImpiantoNextActionHighlight);
ui.exportCurrentCommessaBtn.addEventListener("click", () => exportCommessaSummary(selectedCommessaId, selectedCommessaName));
ui.mapFullscreenBtn.addEventListener("click", toggleMapFullscreen);
ui.squadreWhatsappAllBtn?.addEventListener("click", shareAllSquadreToWhatsApp);
ui.impiantoSearch.addEventListener("input", onImpiantoSearchInput);
ui.viewDoneBtn.addEventListener("click", () => setImpiantiViewMode("done"));
ui.viewTodoBtn.addEventListener("click", () => setImpiantiViewMode("todo"));
ui.personaleForm.addEventListener("submit", addPersonale);
ui.mezziForm.addEventListener("submit", addMezzo);
ui.squadraForm.addEventListener("submit", saveSquadraComposition);
ui.squadraCommessa.addEventListener("change", autofillSquadraForm);
ui.squadraCalendarDate.addEventListener("change", () => {
  manualSquadreFilterDateKey = ui.squadraCalendarDate.value || "";
  if (ui.squadreFilterDate) ui.squadreFilterDate.value = manualSquadreFilterDateKey;
  renderSquadre();
});
ui.squadreFilterDate?.addEventListener("change", onSquadreFilterDateChange);
ui.squadreFilterClearBtn?.addEventListener("click", clearManualSquadreFilterDate);
ui.addSquadraRowBtn.addEventListener("click", () => addSquadraRow());
ui.personaleImportBtn.addEventListener("click", importPersonaleFromExcel);
ui.mezziImportBtn.addEventListener("click", importMezziFromExcel);
ui.openPanelCommesse.addEventListener("click", () => openManagementPanel("commesse"));
ui.openPanelSquadre.addEventListener("click", () => openManagementPanel("squadre"));
ui.openPanelPersonale.addEventListener("click", () => openManagementPanel("personale"));
ui.openPanelMezzi.addEventListener("click", () => openManagementPanel("mezzi"));
ui.openPanelUtenti.addEventListener("click", () => openManagementPanel("utenti"));
ui.openPanelInfoUtili.addEventListener("click", () => openManagementPanel("infoUtili"));
ui.openPrivateDocsBtn.addEventListener("click", openPrivateDocsPage);
ui.openPersonalServicesBtn.addEventListener("click", openPersonalServicesPage);
ui.openHoursBtn.addEventListener("click", openHoursPage);
ui.openSegnalazioniBtn.addEventListener("click", openSegnalazioniPage);
ui.openHowtoBtn.addEventListener("click", openHowtoPage);
ui.managementCloseBtn.addEventListener("click", closeManagementPanel);
ui.userToggleBtn.addEventListener("click", toggleUserDetailsPanel);
ui.weatherCloseBtn.addEventListener("click", closeWeatherModal);
ui.backFromFuelBtn.addEventListener("click", closeFuelPage);
ui.fuelMezzoDetailsBtn.addEventListener("click", toggleFuelMezzoDetails);
ui.backFromPersonalServicesBtn.addEventListener("click", closePersonalServicesPage);
ui.backFromSegnalazioniBtn.addEventListener("click", closeSegnalazioniPage);
ui.backFromHowtoBtn.addEventListener("click", closeHowtoPage);
ui.backFromPrivateDocsBtn.addEventListener("click", closePrivateDocsPage);
ui.backFromHoursBtn.addEventListener("click", closeHoursPage);
ui.hoursForm.addEventListener("submit", finalizeHoursReport);
ui.addHoursCommessaBtn.addEventListener("click", () => {
  unlockHoursFinalizeButton();
  addHoursCommessaBlock();
});
ui.hoursDate?.addEventListener("input", unlockHoursFinalizeButton);
ui.viewHoursBtn.addEventListener("click", openHoursViewModal);
ui.hoursViewCloseBtn?.addEventListener("click", closeHoursViewModal);
ui.hoursViewModal?.addEventListener("click", (event) => {
  if (event.target === ui.hoursViewModal) closeHoursViewModal();
});
ui.hoursStatsMonth?.addEventListener("change", () => {
  if (ui.hoursTableMonth) ui.hoursTableMonth.value = ui.hoursStatsMonth.value || "";
});
ui.hoursTableMonth?.addEventListener("change", loadHoursMonthlyTable);
ui.hoursTableCommessaSelect?.addEventListener("change", loadHoursMonthlyTable);
ui.hoursTotalOperatorBtn?.addEventListener("click", loadHoursTotalByOperator);
ui.hoursTotalOperatorCommessaBtn?.addEventListener("click", loadHoursTotalByOperatorAndCommessa);
ui.hoursTableExportBtn?.addEventListener("click", exportHoursMonthlyTable);
ui.privateDocsPresetPinBtn.addEventListener("click", () => applyPrivateDocPreset("pin"));
ui.privateDocsPresetTesseraBtn.addEventListener("click", () => applyPrivateDocPreset("tessera"));
ui.privateDocsForm.addEventListener("submit", savePrivateDocument);
ui.personalServicesCategories?.addEventListener("click", onPersonalServiceCategoryClick);
ui.personalServicesRadius?.addEventListener("change", () => {
  if (activePersonalServiceCategory) loadPersonalServicesByCategory(activePersonalServiceCategory);
});
ui.segnalazioneForm.addEventListener("submit", generateSegnalazionePdf);
ui.segnalazionePreposto.addEventListener("input", syncSegnalazioneFirmaPreposto);
ui.segnalazioneShareWhatsappBtn.addEventListener("click", () => shareSegnalazione("whatsapp"));
ui.segnalazioneShareEmailBtn.addEventListener("click", () => shareSegnalazione("email"));
ui.manualImpiantoForm.addEventListener("submit", addManualImpianto);
ui.adminUserForm.addEventListener("submit", addAdminUserByEmail);
ui.externalAppForm.addEventListener("submit", saveExternalAppForCurrentUser);
ui.resourceForm.addEventListener("submit", addResourceItem);
ui.resourceType.addEventListener("change", updateResourceFormByType);
ui.impiantoEditCloseBtn.addEventListener("click", closeImpiantoEditor);
ui.impiantoEditForm.addEventListener("submit", saveImpiantoEdits);
ui.impiantoReportCloseBtn.addEventListener("click", closeImpiantoReportModal);
ui.impiantoReportForm.addEventListener("submit", submitImpiantoReport);
ui.enableNotificationsBtn?.addEventListener("click", async () => {
  await persistNotificationAutoPreference(true);
  await enablePushNotifications({ auto: false });
});
ui.testNotificationBtn?.addEventListener("click", sendTestNotification);
window.addEventListener("online", updateConnectivityStatus);
window.addEventListener("offline", updateConnectivityStatus);
ui.commessaResourceViewerCloseBtn.addEventListener("click", closeCommessaResourceViewer);
document.querySelectorAll(".resource-filter-btn").forEach((btn) => {
  btn.addEventListener("click", () => {
    activeResourceManageFilter = btn.dataset.resourceFilter || "";
    renderResourceManageFilters();
    renderResourcesList();
  });
});

addSquadraRow();
initHoursPage();
initGeolocation();
prefillSegnalazioneDateTime();
renderHowtoFaq();
applyRoute();
window.addEventListener("hashchange", applyRoute);
loadPendingSheetExports();
startSheetRetryLoop();
initHelpCenterFaq();
renderResourceManageFilters();
updateResourceFormByType();
updateConnectivityStatus();
initPwaCapabilities();

function toggleUserDetailsPanel() {
  const isHidden = ui.userDetailsPanel.classList.contains("hidden");
  ui.userDetailsPanel.classList.toggle("hidden", !isHidden);
  ui.userToggleBtn.setAttribute("aria-expanded", String(isHidden));
}

function updateNotificationUi(message, canTest = false) {
  if (ui.pwaNotificationStatus) ui.pwaNotificationStatus.textContent = message;
  if (ui.testNotificationBtn) ui.testNotificationBtn.disabled = !canTest;
}

async function initPwaCapabilities() {
  if (!("serviceWorker" in navigator)) {
    updateNotificationUi("Notifiche: browser non supportato.");
    if (ui.enableNotificationsBtn) ui.enableNotificationsBtn.disabled = true;
    return;
  }
  try {
    serviceWorkerRegistration = await navigator.serviceWorker.ready;
  } catch (error) {
    console.warn("Service Worker non pronto per notifiche:", error);
  }
  if (!("Notification" in window)) {
    updateNotificationUi("Notifiche: API non disponibile su questo dispositivo.");
    if (ui.enableNotificationsBtn) ui.enableNotificationsBtn.disabled = true;
    return;
  }
  if (Notification.permission === "granted") {
    updateNotificationUi("Notifiche attive.");
    await ensurePushSubscription();
    return;
  }
  if (Notification.permission === "denied") {
    updateNotificationUi("Notifiche bloccate. Sbloccale dalle impostazioni browser.");
    if (ui.enableNotificationsBtn) ui.enableNotificationsBtn.disabled = true;
    return;
  }
  updateNotificationUi("Notifiche disattive. Premi 'Attiva notifiche'.");
  await maybeAutoEnableNotifications();
}

async function maybeAutoEnableNotifications() {
  if (hasTriedAutoEnableNotifications) return;
  if (!isAutoNotificationEnabled()) return;
  if (!("Notification" in window)) return;
  if (Notification.permission !== "default") return;
  hasTriedAutoEnableNotifications = true;
  updateNotificationUi("Attivazione automatica notifiche in corso...");
  await enablePushNotifications({ auto: true });
}

async function enablePushNotifications(options = {}) {
  const { auto = false } = options;
  if (!("Notification" in window)) return;
  let permission = "default";
  try {
    permission = await Notification.requestPermission();
  } catch (error) {
    console.warn("Richiesta permesso notifiche non riuscita:", error);
    updateNotificationUi("Impossibile richiedere i permessi notifiche su questo browser.");
    return;
  }
  if (permission !== "granted") {
    if (permission === "default" && auto) {
      updateNotificationUi("Attivazione automatica bloccata dal browser. Premi 'Attiva notifiche'.");
      return;
    }
    updateNotificationUi("Notifiche non autorizzate.");
    return;
  }
  await attivaNotifiche();
}

async function attivaNotifiche() {
  try {
    if (!firebaseMessaging) {
      updateNotificationUi("Notifiche locali attive (push cloud non disponibile).", true);
      return;
    }
    if (!PUSH_PUBLIC_VAPID_KEY) {
      updateNotificationUi("Notifiche locali attive (chiave VAPID assente).", true);
      return;
    }
    if (!serviceWorkerRegistration && "serviceWorker" in navigator) {
      serviceWorkerRegistration = await navigator.serviceWorker.ready;
    }
    const token = await firebaseMessaging.getToken({
      vapidKey: PUSH_PUBLIC_VAPID_KEY,
      serviceWorkerRegistration
    });
    if (token) {
      localStorage.setItem("heraPushFcmToken", token);
      console.log("Token push:", token);
      if (currentUser) {
        await db.collection("platformUsers").doc(currentUser.uid).set({
          pushToken: token,
          pushTokenUpdatedAt: firebase.firestore.FieldValue.serverTimestamp()
        }, { merge: true });
      }
      updateNotificationUi("Notifiche push attive.", true);
      return;
    }
    updateNotificationUi("Notifiche locali attive (token push non disponibile).", true);
  } catch (error) {
    console.error("Errore notifiche:", error);
    updateNotificationUi("Notifiche locali attive (push cloud non disponibile).", true);
  }
}

function urlBase64ToUint8Array(base64String) {
  const padding = "=".repeat((4 - (base64String.length % 4)) % 4);
  const base64 = (base64String + padding).replace(/-/g, "+").replace(/_/g, "/");
  const rawData = atob(base64);
  return Uint8Array.from([...rawData].map((char) => char.charCodeAt(0)));
}

async function ensurePushSubscription() {
  if (!firebaseMessaging) {
    updateNotificationUi("Notifiche attive (Firebase Messaging non disponibile).", true);
    return;
  }
  const existingToken = localStorage.getItem("heraPushFcmToken");
  if (existingToken) {
    updateNotificationUi("Notifiche push attive.", true);
    return;
  }
  await attivaNotifiche();
}

async function sendTestNotification() {
  if (Notification.permission !== "granted") {
    updateNotificationUi("Abilita prima i permessi notifiche.");
    return;
  }
  const title = "Hera App";
  const options = {
    body: "Test notifiche completato con successo.",
    icon: "./icons/hera-icon.svg",
    badge: "./icons/hera-icon.svg",
    tag: "hera-test-notification",
    data: { url: "./index.html" }
  };
  if (serviceWorkerRegistration) {
    await serviceWorkerRegistration.showNotification(title, options);
    if ("sync" in serviceWorkerRegistration) {
      try {
        await serviceWorkerRegistration.sync.register("hera-app-background-check");
      } catch (error) {
        console.warn("Background sync non disponibile:", error);
      }
    }
    return;
  }
  new Notification(title, options);
}

async function showLocalNotification(title, options = {}) {
  if (!("Notification" in window) || Notification.permission !== "granted") return false;
  const payload = {
    icon: "./icons/hera-icon.svg",
    badge: "./icons/hera-icon.svg",
    ...options
  };
  if (serviceWorkerRegistration) {
    await serviceWorkerRegistration.showNotification(title, payload);
    return true;
  }
  new Notification(title, payload);
  return true;
}

async function publishGlobalNotificationEvent(eventType, payload = {}) {
  if (!currentUser) return;
  try {
    await db.collection("appNotifications").add({
      eventType,
      title: payload.title || "Hera App",
      body: payload.body || "Nuovo aggiornamento operativo.",
      commessaId: payload.commessaId || "",
      commessaName: payload.commessaName || "",
      impiantoName: payload.impiantoName || "",
      impiantoKey: payload.impiantoKey || "",
      createdByUid: currentUser.uid || "",
      createdByName: currentUser.displayName || currentUser.email || "Operatore",
      createdByEmail: currentUser.email || "",
      createdAt: firebase.firestore.FieldValue.serverTimestamp()
    });
  } catch (error) {
    console.warn("Invio evento notifica globale non riuscito:", error);
  }
}

function subscribeGlobalNotifications() {
  stopGlobalNotificationsSubscription();
  globalNotificationsInitialized = false;
  unsubscribeGlobalNotifications = db.collection("appNotifications")
    .orderBy("createdAt", "desc")
    .limit(40)
    .onSnapshot(async (snapshot) => {
      if (!globalNotificationsInitialized) {
        globalNotificationsInitialized = true;
        return;
      }
      const added = snapshot.docChanges().filter((change) => change.type === "added");
      for (const change of added) {
        const data = change.doc.data() || {};
        if (String(data.createdByUid || "") === String(currentUser?.uid || "")) continue;
        await showLocalNotification(data.title || "Hera App", {
          body: data.body || "Nuovo aggiornamento operativo.",
          tag: `hera-event-${change.doc.id}`,
          data: { url: "./index.html" }
        });
      }
    }, (error) => {
      console.warn("Sottoscrizione notifiche globali non disponibile:", error);
    });
}

function stopGlobalNotificationsSubscription() {
  if (unsubscribeGlobalNotifications) {
    unsubscribeGlobalNotifications();
    unsubscribeGlobalNotifications = null;
  }
  globalNotificationsInitialized = false;
}

function weatherCodeLabel(weatherCode) {
  const code = Number(weatherCode);
  const weatherMap = {
    0: "☀️ Sereno",
    1: "⛅ Poco nuvoloso",
    2: "☁️ Parzialmente nuvoloso",
    3: "☁️ Coperto",
    45: "🌫️ Nebbia",
    48: "🌫️ Nebbia con brina",
    51: "🌦️ Pioviggine",
    53: "🌦️ Pioviggine moderata",
    55: "🌧️ Pioviggine intensa",
    56: "🌨️ Pioviggine gelata",
    57: "🌨️ Pioggia gelata",
    61: "🌧️ Pioggia debole",
    63: "🌧️ Pioggia moderata",
    65: "⛈️ Pioggia forte",
    66: "🧊 Pioggia gelata debole",
    67: "🧊 Pioggia gelata forte",
    71: "🌨️ Neve debole",
    73: "🌨️ Neve moderata",
    75: "❄️ Neve intensa",
    77: "🌨️ Nevischio",
    80: "🌧️ Rovesci deboli",
    81: "🌧️ Rovesci moderati",
    82: "⛈️ Rovesci forti",
    85: "🌨️ Rovesci di neve",
    86: "❄️ Rovesci di neve forti",
    95: "⛈️ Temporale",
    96: "⛈️ Temporale con grandine",
    99: "⛈️ Temporale forte con grandine"
  };
  return weatherMap[code] || "ℹ️ Condizioni variabili";
}

auth.onAuthStateChanged((user) => {
  currentUser = user || null;
  const loggedIn = Boolean(user);

  ui.loginBtn.disabled = loggedIn;
  ui.switchAccountBtn.classList.toggle("hidden", !loggedIn);
  ui.switchAccountBtn.disabled = !loggedIn;
  ui.logoutBtn.disabled = !loggedIn;
  ui.driveConnectBtn.disabled = !loggedIn;
  ui.driveConnectBtn.classList.toggle("hidden", !loggedIn);
  ui.user.textContent = loggedIn
    ? `Loggato: ${user.email || "email non disponibile"}`
    : "Non loggato";
  ui.userName.textContent = loggedIn
    ? `Nome utente: ${user.displayName || "Nome non disponibile"}`
    : "Nome utente: -";
  prefillSegnalazioneDateTime();
  syncSegnalazioneFirmaPreposto();

  ui.importBtn.disabled = !loggedIn || !selectedCommessaId || pendingRows.length === 0 || !canManageData();
  ui.exportCurrentCommessaBtn.disabled = !loggedIn || !selectedCommessaId || !canManageData();
  updateAdminControls();

  stopCommesseSubscription();
  stopImpiantiSubscription();
  stopChatSubscription();
  stopDriveBridgeSubscription();
  stopPersonaleSubscription();
  stopMezziSubscription();
  stopSquadreSubscription();
  stopUsersSubscription();
  stopAdminUsersSubscription();
  stopResourcesSubscription();
  stopPrivateDocsSubscription();
  stopGpsRequestsSubscription();
  stopGlobalNotificationsSubscription();
  stopChatRetentionLoop();
  selectedCommessaId = "";
  selectedCommessaName = "";
  updateCommessaContextUI();
  window.location.hash = "";
  ui.commesseLista.innerHTML = "";
  ui.squadraCommessa.innerHTML = "<option value=''>Seleziona commessa</option>";
  ui.squadreLista.innerHTML = "";
  if (ui.squadreFilterDate) ui.squadreFilterDate.value = "";
  manualSquadreFilterDateKey = "";
  squadreByCommessa = new Map();
  squadreHistoryByDate = new Map();
  commesseById = new Map();
  resourceRecords = [];
  privateDocsRecords = [];
  gpsUpdateRequests = [];
  hoursApprovalRequests = [];
  renderPrivateDocsList();
  renderResourceButtonsForCommessa();
  closeCommessaResourceViewer();
  ui.impiantiLista.innerHTML = loggedIn
    ? "<p class='muted'>Seleziona una commessa.</p>"
    : "<p class='muted'>Fai login per vedere le commesse.</p>";
  clearMap();
  lastReadChatAt = null;
  resetDriveState();
  if (loggedIn) {
    const storedToken = getStoredDriveToken();
    if (storedToken) {
      driveAccessToken = storedToken;
      window.googleDriveAccessToken = storedToken;
      autoConnectDriveBridge({ notifyOnError: false });
    }
  }
  renderChat([]);
  applyRoute();

  if (loggedIn) {
    startPresenceHeartbeat();
    upsertCurrentPlatformUser();
    subscribeCommesse();
    subscribeChat();
    subscribeUsers();
    subscribeAdminUsers();
    subscribeDriveBridge();
    subscribePersonale();
    subscribeMezzi();
    subscribeSquadre();
    subscribeResources();
    subscribePrivateDocs();
    subscribeGpsRequests();
    subscribeGlobalNotifications();
    processPendingSheetExports();
    startChatRetentionLoop();
  } else {
    stopPresenceHeartbeat();
  }
  renderHeaderActivitySummary();
  renderExternalApps();
  fetchWeather();
  renderNextActionCard();
});

function updateAdminControls() {
  const canManage = canManageData();
  [ui.openPanelCommesse, ui.openPanelSquadre, ui.openPanelPersonale, ui.openPanelMezzi, ui.openPanelUtenti, ui.openPanelInfoUtili]
    .forEach((button) => button.classList.toggle("hidden", !canManage));
  ui.commessaName.disabled = !canManage;
  const submitBtn = ui.commessaForm.querySelector("button[type='submit']");
  if (submitBtn) submitBtn.disabled = !canManage;
  ui.personaleNome.disabled = !canManage;
  ui.mezzoNId.disabled = !canManage;
  ui.mezzoMarca.disabled = !canManage;
  ui.mezzoModello.disabled = !canManage;
  ui.mezzoPortataCarico.disabled = !canManage;
  ui.mezzoMassaComplessivaKg.disabled = !canManage;
  ui.mezzoAlimentazione.disabled = !canManage;
  if (ui.personaleForm.querySelector("button[type='submit']")) ui.personaleForm.querySelector("button[type='submit']").disabled = !canManage;
  if (ui.mezziForm.querySelector("button[type='submit']")) ui.mezziForm.querySelector("button[type='submit']").disabled = !canManage;
  ui.personaleImportBtn.disabled = !canManage;
  ui.mezziImportBtn.disabled = !canManage;
  ui.importBtn.disabled = !canManage || !auth.currentUser || !selectedCommessaId || pendingRows.length === 0;
  ui.manualImpiantoDenominazione.disabled = !canManage;
  ui.manualImpiantoComune.disabled = !canManage;
  ui.manualImpiantoIndirizzo.disabled = !canManage;
  ui.manualImpiantoCodice.disabled = !canManage;
  ui.manualImpiantoSubmit.disabled = !canManage;
  ui.adminUserEmail.disabled = !canManage;
  if (ui.adminUserForm.querySelector("button[type='submit']")) ui.adminUserForm.querySelector("button[type='submit']").disabled = !canManage;
  ui.resourceType.disabled = !canManage;
  ui.resourceTitle.disabled = !canManage;
  ui.resourceValue.disabled = !canManage;
  ui.resourceCommesse.disabled = !canManage;
  ui.resourceSubmit.disabled = !canManage;
  if (ui.externalAppName) ui.externalAppName.disabled = !auth.currentUser;
  if (ui.externalAppUrl) ui.externalAppUrl.disabled = !auth.currentUser;
  if (ui.externalAppForm && ui.externalAppForm.querySelector("button[type='submit']")) {
    ui.externalAppForm.querySelector("button[type='submit']").disabled = !auth.currentUser;
  }
  ui.squadraCommessa.disabled = !canManage;
  ui.squadreWhatsappAllBtn?.classList.toggle("hidden", !canManage);
  ui.squadreFilterControls?.classList.toggle("hidden", !canManage);
  if (ui.squadreFilterDate) ui.squadreFilterDate.disabled = !canManage;
  if (ui.squadreFilterClearBtn) ui.squadreFilterClearBtn.disabled = !canManage;
  ui.exportCurrentCommessaBtn?.classList.toggle("hidden", !canManage);
  ui.exportCurrentCommessaBtn.disabled = !canManage || !auth.currentUser || !selectedCommessaId;
  if (ui.gpsRequestsList && !canManage) {
    ui.gpsRequestsList.innerHTML = "<p class='muted'>Solo gli admin possono gestire le richieste GPS.</p>";
  } else if (ui.gpsRequestsList && canManage) {
    renderGpsRequests();
  }
  ui.squadraRiferimento.disabled = !canManage;
  ui.addSquadraRowBtn.disabled = !canManage;
  ui.squadraRows.querySelectorAll("input,button").forEach((el) => { el.disabled = !canManage; });
  if (ui.squadraForm.querySelector("button[type='submit']")) ui.squadraForm.querySelector("button[type='submit']").disabled = !canManage;
  ui.squadraHint.textContent = canManage
    ? "Suggerimento: usa i nomi in Personale e i mezzi in Mezzi per compilare le squadre."
    : "Solo l'admin può modificare personale, mezzi e composizione squadre.";
  updateResourceFormByType();
  renderUserPermissionList();
  renderExternalApps();
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
  if (!canManageData()) {
    closeSideMenu();
    return;
  }
  const panelMap = {
    commesse: { el: ui.panelCommesse, title: "Aggiungi commesse" },
    squadre: { el: ui.panelSquadre, title: "Composizione squadre" },
    personale: { el: ui.panelPersonale, title: "Personale" },
    mezzi: { el: ui.panelMezzi, title: "Mezzi" },
    utenti: { el: ui.panelUtenti, title: "Gestione utenti" },
    infoUtili: { el: ui.panelInfoUtili, title: "Informazioni utili" }
  };
  const target = panelMap[panel];
  if (!target) return;
  [ui.panelCommesse, ui.panelSquadre, ui.panelPersonale, ui.panelMezzi, ui.panelUtenti, ui.panelInfoUtili].forEach((el) => el.classList.add("hidden"));
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
  const canUseNativeFullscreen = typeof mapContainer.requestFullscreen === "function";
  if (!canUseNativeFullscreen) {
    toggleMapCssFullscreen(mapContainer);
    return;
  }
  if (!document.fullscreenElement) {
    mapContainer.requestFullscreen().catch((error) => {
      console.error("Fullscreen mappa non disponibile, uso fallback CSS:", error);
      toggleMapCssFullscreen(mapContainer);
    });
    return;
  }
  document.exitFullscreen().catch((error) => {
    console.error("Uscita fullscreen non disponibile:", error);
  });
}

function toggleMapCssFullscreen(mapContainer) {
  mapContainer.classList.toggle("map-fallback-fullscreen");
  document.body.classList.toggle("map-fallback-fullscreen-open", mapContainer.classList.contains("map-fallback-fullscreen"));
  syncMapFullscreenUiState();
  setTimeout(() => map.invalidateSize(), 60);
}

function syncMapFullscreenUiState() {
  if (!ui.mapFullscreenBtn) return;
  const mapContainer = document.getElementById("map");
  const isCssFullscreen = mapContainer ? mapContainer.classList.contains("map-fallback-fullscreen") : false;
  const isNativeFullscreen = Boolean(document.fullscreenElement);
  const isFullscreenOpen = isCssFullscreen || isNativeFullscreen;
  ui.mapFullscreenBtn.textContent = isFullscreenOpen ? "✕ Chiudi schermo intero" : "⤢ Mappa schermo intero";
}

function applyRoute() {
  const hash = window.location.hash || "";
  const match = hash.match(/^#commessa=([a-zA-Z0-9_-]+)(?:&resource=(phone|document|note))?$/);
  const fuelMatch = hash.match(/^#fuel=(.+)$/);
  const showSegnalazioni = hash === "#segnalazioni";
  const showHowto = hash === "#howto";
  const showPrivateDocs = hash === "#documenti";
  const personalServiceMatch = hash.match(/^#servizi-personali(?:=([a-z]+))?$/);
  const showHours = hash === "#ore";
  const commessaIdFromHash = match ? match[1] : "";
  const resourceTypeFromHash = match ? (match[2] || "") : "";
  const showFuel = Boolean(fuelMatch);
  const showPersonalServices = Boolean(personalServiceMatch);
  const showImpianti = Boolean(commessaIdFromHash && selectedCommessaId === commessaIdFromHash);
  const showResourceViewer = Boolean(showImpianti && resourceTypeFromHash);
  ui.homePage.classList.toggle("hidden", showImpianti || showFuel || showSegnalazioni || showHowto || showPrivateDocs || showHours || showPersonalServices);
  ui.impiantiPage.classList.toggle("hidden", !showImpianti);
  ui.fuelPage.classList.toggle("hidden", !showFuel);
  ui.personalServicesPage.classList.toggle("hidden", !showPersonalServices);
  ui.segnalazioniPage.classList.toggle("hidden", !showSegnalazioni);
  ui.howtoPage.classList.toggle("hidden", !showHowto);
  ui.privateDocsPage.classList.toggle("hidden", !showPrivateDocs);
  ui.hoursPage.classList.toggle("hidden", !showHours);
  document.body.classList.toggle("resource-view-open", showResourceViewer);
  ui.mapFullscreenBtn.classList.toggle("hidden", showResourceViewer);
  const mapElement = document.getElementById("map");
  if (mapElement) mapElement.classList.toggle("hidden", showResourceViewer);
  if (ui.gpsStatus) ui.gpsStatus.classList.toggle("hidden", showResourceViewer);
  const impiantiCard = ui.impiantiLista?.closest(".card");
  if (impiantiCard) impiantiCard.classList.toggle("hidden", showResourceViewer);
  if (showImpianti) {
    ui.impiantiPageTitle.textContent = `Impianti commessa: ${selectedCommessaName || "Commessa"}`;
    if (showResourceViewer) {
      activeResourceTypeForViewer = resourceTypeFromHash;
      renderCommessaResourceViewer();
      ui.commessaResourceViewer.classList.remove("hidden");
      ui.commessaResourceViewer.classList.add("page-mode");
      ui.commessaResourceViewerCloseBtn.textContent = "← Torna alla commessa";
    } else {
      closeCommessaResourceViewer();
    }
    setTimeout(() => map.invalidateSize(), 50);
  }
  if (showHowto) renderHowtoFaq();
  if (showPrivateDocs) renderPrivateDocsList();
  if (showFuel) {
    setTimeout(() => {
      if (fuelMapInstance) fuelMapInstance.invalidateSize();
    }, 50);
  }
  if (showPersonalServices) {
    const categoryFromHash = personalServiceMatch && personalServiceMatch[1] ? personalServiceMatch[1] : "";
    if (categoryFromHash && categoryFromHash !== activePersonalServiceCategory) {
      loadPersonalServicesByCategory(categoryFromHash);
    }
    setTimeout(() => {
      if (personalServicesMapInstance) personalServicesMapInstance.invalidateSize();
    }, 50);
  }
  renderNextActionCard();
}

function openImpiantiPage() {
  if (!selectedCommessaId) return;
  localStorage.setItem(LAST_OPENED_COMMESSA_KEY, selectedCommessaId);
  window.location.hash = `commessa=${selectedCommessaId}`;
  applyRoute();
}

function closeImpiantiPage() {
  localStorage.removeItem(LAST_OPENED_COMMESSA_KEY);
  window.location.hash = "";
  ui.exportCurrentCommessaBtn.disabled = true;
  document.body.classList.remove("resource-view-open");
  closeCommessaResourceViewer();
  applyRoute();
}

function closeFuelPage() {
  window.location.hash = "";
  applyRoute();
}

function openPersonalServicesPage() {
  window.location.hash = "servizi-personali";
  applyRoute();
}

function closePersonalServicesPage() {
  window.location.hash = "";
  applyRoute();
}

function setCurrentWorkflowStep(stepId) {
  currentWorkflowStepId = String(stepId || "").trim();
  if (!currentWorkflowStepId) {
    localStorage.removeItem(USER_WORKFLOW_STEP_KEY);
  } else {
    localStorage.setItem(USER_WORKFLOW_STEP_KEY, currentWorkflowStepId);
  }
  renderNextActionCard();
}

function getWorkflowSteps() {
  const routeHash = window.location.hash || "";
  const hasSelectedCommessa = Boolean(selectedCommessaId);
  const todoCount = currentImpianti.filter((item) => !item.done).length;
  const doneCount = currentImpianti.filter((item) => Boolean(item.done)).length;
  const hasOpenCommessaRoute = hasSelectedCommessa && routeHash === `#commessa=${selectedCommessaId}`;
  const isLoggedIn = Boolean(currentUser);
  return [
    {
      id: "login",
      label: "Login con Google",
      description: "Accedi con il tuo account per sbloccare commesse e strumenti.",
      available: !isLoggedIn,
      done: isLoggedIn,
      action: () => loginWithGoogle()
    },
    {
      id: "select-commessa",
      label: "Seleziona commessa",
      description: "Scegli una commessa dalla home per iniziare il turno operativo.",
      available: isLoggedIn && !hasSelectedCommessa,
      done: hasSelectedCommessa,
      action: () => {
        window.location.hash = "";
        applyRoute();
      }
    },
    {
      id: "open-commessa",
      label: "Apri impianti commessa",
      description: "Apri la commessa selezionata per lavorare sugli impianti.",
      available: isLoggedIn && hasSelectedCommessa && !hasOpenCommessaRoute,
      done: hasOpenCommessaRoute,
      action: () => openImpiantiPage()
    },
    {
      id: "mark-next-impianto",
      label: "Completa prossimo impianto",
      description: todoCount > 0
        ? `Hai ${todoCount} impianti da fare: apri il primo e premi FATTO.`
        : "Nessun impianto da completare in questa commessa.",
      available: isLoggedIn && hasOpenCommessaRoute && todoCount > 0,
      done: hasOpenCommessaRoute && todoCount === 0,
      action: () => setImpiantiViewMode("todo")
    },
    {
      id: "review-completed",
      label: "Controlla impianti fatti",
      description: doneCount > 0
        ? `Hai ${doneCount} impianti completati: verifica riepilogo e note finali.`
        : "Ancora nessun impianto completato da verificare.",
      available: isLoggedIn && hasOpenCommessaRoute && doneCount > 0,
      done: false,
      action: () => setImpiantiViewMode("done")
    }
  ];
}

function renderNextActionCard() {
  if (!ui.nextActionSummary) return;
  const steps = getWorkflowSteps();
  const availableSteps = steps.filter((step) => step.available);
  const stepMap = new Map(steps.map((step) => [step.id, step]));
  let primary = stepMap.get(currentWorkflowStepId);
  if (!primary || !primary.available) primary = availableSteps[0] || steps[steps.length - 1];

  if (primary?.id !== currentWorkflowStepId) {
    currentWorkflowStepId = primary?.id || "";
    if (currentWorkflowStepId) localStorage.setItem(USER_WORKFLOW_STEP_KEY, currentWorkflowStepId);
    else localStorage.removeItem(USER_WORKFLOW_STEP_KEY);
  }

  if (!primary) {
    ui.nextActionSummary.textContent = "Prossima azione consigliata: nessuna azione disponibile al momento.";
    return;
  }

  ui.nextActionSummary.textContent = `Prossima azione consigliata: ${primary.label}.`;
  if (ui.commesseNextAction) {
    ui.commesseNextAction.textContent = `Prossima azione consigliata: ${primary.label}.`;
  }
  if (ui.squadreNextAction) {
    ui.squadreNextAction.textContent = "Prossima azione consigliata: premi sul tuo mezzo per trovare il distributore.";
  }
  renderImpiantoNextActionUI();
}

function getCurrentImpiantoNextAction() {
  return IMPIANTO_NEXT_ACTION_FLOW[impiantoNextActionIndex] || IMPIANTO_NEXT_ACTION_FLOW[0];
}

function impiantoNextActionLabel(actionKey) {
  if (actionKey === "navigate") return "Naviga verso l'impianto";
  if (actionKey === "done") return "Fatto per aggiornare lo stato";
  return "Invia messaggio WhatsApp";
}

function impiantoNextActionIcon(actionKey) {
  if (actionKey === "navigate") return "🗺️";
  if (actionKey === "done") return "✅";
  return "✉️";
}

function buildInlineActionButton(label, actionKey, compact = false) {
  const icon = impiantoNextActionIcon(actionKey);
  const compactClass = compact ? " inline-action-preview--compact" : "";
  const iconHtml = `<span class="inline-action-preview__icon" aria-hidden="true">${icon}</span>`;
  if (compact) {
    return `<span class="inline-action-preview${compactClass}" data-action-key="${escapeHTML(actionKey)}" role="img" aria-label="${escapeHTML(label)}">${iconHtml}</span>`;
  }
  return `<span class="inline-action-preview${compactClass}" data-action-key="${escapeHTML(actionKey)}" aria-hidden="true">${iconHtml}${escapeHTML(label)}</span>`;
}

function renderImpiantoNextActionUI() {
  if (!ui.impiantiNextAction && !ui.showNextActionBtn) return;
  const actionKey = getCurrentImpiantoNextAction();
  const label = impiantoNextActionLabel(actionKey);
  const actionIcon = impiantoNextActionIcon(actionKey);
  const showButtonPreview = buildInlineActionButton(`Mostra pulsante ${label}`, actionKey, true);
  const targetButtonPreview = buildInlineActionButton(label, actionKey, true);
  if (ui.showNextActionBtn) {
    ui.showNextActionBtn.innerHTML = `Mostra pulsante <span class="inline-action-preview inline-action-preview--compact" data-action-key="${escapeHTML(actionKey)}" aria-hidden="true"><span class="inline-action-preview__icon" aria-hidden="true">${actionIcon}</span></span>`;
    ui.showNextActionBtn.setAttribute("aria-label", `Mostra pulsante ${label}`);
    ui.showNextActionBtn.classList.toggle("btn-primary", impiantoNextActionHighlightEnabled);
  }
  if (ui.impiantiNextAction) {
    ui.impiantiNextAction.innerHTML = impiantoNextActionHighlightEnabled
      ? `Passaggio consigliato: premi questo pulsante ${targetButtonPreview}.`
      : `Prossima azione consigliata: premi prima ${showButtonPreview}.`;
  }
}

function toggleImpiantoNextActionHighlight() {
  impiantoNextActionHighlightEnabled = !impiantoNextActionHighlightEnabled;
  renderImpiantoNextActionUI();
  renderImpianti();
}

function registerImpiantoSessionAction(actionKey) {
  const expectedAction = getCurrentImpiantoNextAction();
  if (actionKey !== expectedAction) return;
  impiantoNextActionIndex = (impiantoNextActionIndex + 1) % IMPIANTO_NEXT_ACTION_FLOW.length;
  impiantoNextActionHighlightEnabled = false;
  renderImpiantoNextActionUI();
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

function openHowtoPage() {
  window.location.hash = "howto";
  renderHowtoFaq();
  applyRoute();
  closeSideMenu();
}

function closeHowtoPage() {
  window.location.hash = "";
  applyRoute();
}

function buildHowtoFaqItems() {
  const menuButtons = Array.from(document.querySelectorAll("#side-menu .menu-title-btn"));
  const menuFaqItems = menuButtons.map((button, index) => {
    const buttonId = button.id || `menu-item-${index + 1}`;
    const menuTitle = (button.textContent || "").trim() || "Voce menu";
    const config = MENU_HOWTO_CONTENT[buttonId] || {};
    const fallbackPassi = [
      "Apri il menu (⋮) nella home.",
      `Premi “${menuTitle}”.`,
      "Segui i campi/pulsanti del pannello e conferma l'azione."
    ];
    return {
      id: `menu-${buttonId}`,
      domanda: `Come si usa “${menuTitle}”?`,
      rispostaBreve: config.rispostaBreve || `Questa voce apre “${menuTitle}” con tutte le azioni disponibili.`,
      passi: config.passi || fallbackPassi,
      tags: config.tags || ["menu", "funzione"],
      updatedAt: HOWTO_UPDATED_AT
    };
  });
  return [...menuFaqItems, ...STATIC_HOWTO_ITEMS];
}

function openPrivateDocsPage() {
  if (!currentUser) {
    alert("Devi fare login per usare i documenti personali.");
    return;
  }
  window.location.hash = "documenti";
  applyRoute();
  closeSideMenu();
}

function closePrivateDocsPage() {
  window.location.hash = "";
  applyRoute();
}

function initHoursPage() {
  if (ui.hoursDate) ui.hoursDate.value = new Date().toISOString().slice(0, 10);
  if (ui.hoursTableMonth) ui.hoursTableMonth.value = new Date().toISOString().slice(0, 7);
  if (ui.hoursStatsMonth) ui.hoursStatsMonth.value = new Date().toISOString().slice(0, 7);
  if (!ui.hoursCommesseList) return;
  if (!ui.hoursCommesseList.children.length) addHoursCommessaBlock();
  renderHoursOperatoriOptions();
  renderHoursCommessaSelectOptions();
  renderHoursTableCommessaOptions();
  renderHoursSummary();
  setHoursFinalizeLocked(false);
  renderSavedHoursReports([]);
  if (ui.hoursTableExportBtn) ui.hoursTableExportBtn.disabled = true;
}

function openHoursPage() {
  if (!currentUser) {
    alert("Devi fare login per compilare la gestione ore.");
    return;
  }
  if (!ui.hoursDate.value) ui.hoursDate.value = new Date().toISOString().slice(0, 10);
  if (!ui.hoursStatsMonth?.value) ui.hoursStatsMonth.value = new Date().toISOString().slice(0, 7);
  if (!ui.hoursCommesseList.children.length) addHoursCommessaBlock();
  renderHoursTableCommessaOptions();
  window.location.hash = "ore";
  applyRoute();
  closeSideMenu();
}

function closeHoursPage() {
  window.location.hash = "";
  applyRoute();
}

function renderSavedHoursReports(records = []) {
  if (!ui.hoursSavedList) return;
  if (!records.length) {
    ui.hoursSavedList.innerHTML = "<p class='muted'>Nessun report ore disponibile.</p>";
    return;
  }
  ui.hoursSavedList.innerHTML = records.map((report) => {
    const dateLabel = report.date ? new Date(`${report.date}T00:00:00`).toLocaleDateString("it-IT") : "-";
    const author = report.createdByName || report.createdByEmail || "Operatore";
    const commesseHtml = (Array.isArray(report.entries) ? report.entries : []).map((entry) => {
      const rows = (Array.isArray(entry.rows) ? entry.rows : [])
        .map((row) => `<li>${escapeHTML(row.operatore || "-")}: <b>${escapeHTML(String(row.ore || 0))}h</b></li>`)
        .join("");
      return `
        <div class="item-card">
          <p><b>Commessa:</b> ${escapeHTML(entry.commessaName || "-")}</p>
          <ul>${rows || "<li>Nessun operatore</li>"}</ul>
          ${entry.note ? `<p><b>Nota:</b> ${escapeHTML(entry.note)}</p>` : ""}
        </div>
      `;
    }).join("");
    return `
      <article class="item-card">
        <h3>${escapeHTML(dateLabel)}</h3>
        <p class="muted">Compilato da: ${escapeHTML(author)}</p>
        ${commesseHtml || "<p class='muted'>Nessuna commessa nel report.</p>"}
      </article>
    `;
  }).join("");
}

async function loadSavedHoursReports() {
  if (!currentUser) {
    renderSavedHoursReports([]);
    return;
  }
  if (ui.viewHoursBtn) ui.viewHoursBtn.disabled = true;
  if (ui.hoursSavedList) ui.hoursSavedList.innerHTML = "<p class='muted'>Caricamento ore salvate...</p>";
  try {
    const baseQuery = db.collection("oreReports");
    const snapshot = await baseQuery.orderBy("createdAt", "desc").limit(100).get();
    const reports = snapshot.docs.map((doc) => ({ id: doc.id, ...doc.data() }));
    renderSavedHoursReports(reports);
  } catch (error) {
    console.error("Errore caricamento report ore:", error);
    if (ui.hoursSavedList) ui.hoursSavedList.innerHTML = "<p class='muted'>Errore caricamento ore salvate.</p>";
  } finally {
    if (ui.viewHoursBtn) ui.viewHoursBtn.disabled = false;
  }
}

function openHoursViewModal() {
  if (!currentUser) {
    alert("Devi fare login per visualizzare le ore.");
    return;
  }
  renderHoursTableCommessaOptions();
  if (ui.hoursTableMonth) {
    ui.hoursTableMonth.value = ui.hoursStatsMonth?.value || ui.hoursTableMonth.value || new Date().toISOString().slice(0, 7);
  }
  ui.hoursViewModal?.classList.remove("hidden");
  ui.hoursViewModal?.setAttribute("aria-hidden", "false");
  loadHoursMonthlyTable();
}

function closeHoursViewModal() {
  ui.hoursViewModal?.classList.add("hidden");
  ui.hoursViewModal?.setAttribute("aria-hidden", "true");
}

function getMonthMeta(monthValue) {
  if (!/^\d{4}-\d{2}$/.test(monthValue || "")) return null;
  const [yearText, monthText] = monthValue.split("-");
  const year = Number(yearText);
  const month = Number(monthText);
  if (!year || !month || month < 1 || month > 12) return null;
  const daysInMonth = new Date(year, month, 0).getDate();
  return { year, month, daysInMonth };
}

function resolveHoursStatsMonth() {
  const monthValue = String(ui.hoursStatsMonth?.value || ui.hoursTableMonth?.value || "").trim();
  return { monthValue, monthMeta: getMonthMeta(monthValue) };
}

async function fetchHoursReportsForMonth(monthValue, monthMeta) {
  if (!monthMeta) return [];
  const fromDate = `${monthValue}-01`;
  const toDate = `${monthValue}-${String(monthMeta.daysInMonth).padStart(2, "0")}`;
  const snapshot = await db.collection("oreReports")
    .where("date", ">=", fromDate)
    .where("date", "<=", toDate)
    .orderBy("date", "asc")
    .get();
  const reports = snapshot.docs.map((doc) => ({ id: doc.id, ...doc.data() }));
  return reports;
}

function ensureHoursViewModalOpen() {
  if (!ui.hoursViewModal || !ui.hoursViewModal.classList.contains("hidden")) return;
  ui.hoursViewModal.classList.remove("hidden");
  ui.hoursViewModal.setAttribute("aria-hidden", "false");
}

async function loadHoursMonthlyTable() {
  if (!ui.hoursTableFeedback || !ui.hoursTableContainer) return;
  hoursTableContext = null;
  if (ui.hoursTableCommessaSelect) ui.hoursTableCommessaSelect.disabled = false;
  if (ui.hoursTableExportBtn) ui.hoursTableExportBtn.disabled = true;
  const monthValue = String(ui.hoursTableMonth?.value || "").trim();
  const commessaId = String(ui.hoursTableCommessaSelect?.value || "").trim();
  const monthMeta = getMonthMeta(monthValue);
  if (!monthMeta) {
    ui.hoursTableFeedback.textContent = "Seleziona un mese valido.";
    ui.hoursTableContainer.innerHTML = "";
    return;
  }
  if (!commessaId) {
    ui.hoursTableFeedback.textContent = "Seleziona una commessa per vedere la tabella.";
    ui.hoursTableContainer.innerHTML = "";
    return;
  }
  if (ui.hoursStatsMonth) ui.hoursStatsMonth.value = monthValue;
  ui.hoursTableFeedback.textContent = "Caricamento tabella ore...";
  ui.hoursTableContainer.innerHTML = "";
  try {
    const reports = await fetchHoursReportsForMonth(monthValue, monthMeta);
    renderHoursMonthlyTable(reports, commessaId, monthMeta);
  } catch (error) {
    console.error("Errore caricamento tabella mensile ore:", error);
    ui.hoursTableFeedback.textContent = "Errore caricamento tabella ore.";
    ui.hoursTableContainer.innerHTML = "";
  }
}

async function loadHoursTotalByOperator() {
  if (!ui.hoursTableFeedback || !ui.hoursTableContainer) return;
  hoursTableContext = null;
  if (!currentUser) {
    ui.hoursTableFeedback.textContent = "Devi fare login per visualizzare i totali.";
    ui.hoursTableContainer.innerHTML = "";
    return;
  }
  if (ui.hoursTableExportBtn) ui.hoursTableExportBtn.disabled = true;
  const { monthValue, monthMeta } = resolveHoursStatsMonth();
  if (!monthMeta) {
    ui.hoursTableFeedback.textContent = "Seleziona un mese valido per calcolare i totali.";
    ui.hoursTableContainer.innerHTML = "";
    return;
  }
  if (ui.hoursTableMonth) ui.hoursTableMonth.value = monthValue;
  if (ui.hoursTableCommessaSelect) ui.hoursTableCommessaSelect.disabled = true;
  ensureHoursViewModalOpen();
  ui.hoursTableFeedback.textContent = "Caricamento totale ore per operatore...";
  ui.hoursTableContainer.innerHTML = "";
  try {
    const snapshot = await fetchHoursReportsForMonth(monthValue, monthMeta);
    const operatorTotals = new Map();
    snapshot.forEach((report) => {
      const entries = Array.isArray(report.entries) ? report.entries : [];
      entries.forEach((entry) => {
        (Array.isArray(entry.rows) ? entry.rows : []).forEach((row) => {
          const displayName = String(row.operatore || "").trim();
          const ore = Number(row.ore || 0);
          if (!displayName || ore <= 0) return;
          const normalizedKey = displayName.toLocaleLowerCase("it-IT").replace(/\s+/g, " ").trim();
          if (!operatorTotals.has(normalizedKey)) {
            operatorTotals.set(normalizedKey, { name: displayName, total: 0 });
          }
          operatorTotals.get(normalizedKey).total += ore;
        });
      });
    });
    const rows = Array.from(operatorTotals.values())
      .sort((a, b) => a.name.localeCompare(b.name, "it"))
      .map((item) => {
        const totalLabel = Number.isInteger(item.total)
          ? String(item.total)
          : item.total.toFixed(2).replace(".", ",");
        return `<tr><th scope="row">${escapeHTML(item.name)}</th><td><b>${escapeHTML(totalLabel)}h</b></td></tr>`;
      });
    if (!rows.length) {
      ui.hoursTableFeedback.textContent = "Nessuna ora trovata per calcolare i totali.";
      ui.hoursTableContainer.innerHTML = "";
      return;
    }
    ui.hoursTableContainer.innerHTML = `
      <table class="hours-month-table">
        <thead>
          <tr>
            <th>Operatore</th>
            <th>Totale ore</th>
          </tr>
        </thead>
        <tbody>${rows.join("")}</tbody>
      </table>
    `;
    ui.hoursTableFeedback.textContent = `Totale ore per operatore calcolato per ${monthValue}.`;
    hoursTableContext = {
      mode: "tot_operator",
      monthValue,
      rows: Array.from(operatorTotals.values()).sort((a, b) => a.name.localeCompare(b.name, "it"))
    };
    if (ui.hoursTableExportBtn) ui.hoursTableExportBtn.disabled = false;
  } catch (error) {
    console.error("Errore caricamento totale ore per operatore:", error);
    ui.hoursTableFeedback.textContent = "Errore caricamento totale ore per operatore.";
    ui.hoursTableContainer.innerHTML = "";
  }
}

async function loadHoursTotalByOperatorAndCommessa() {
  if (!ui.hoursTableFeedback || !ui.hoursTableContainer) return;
  hoursTableContext = null;
  if (!currentUser) {
    ui.hoursTableFeedback.textContent = "Devi fare login per visualizzare i totali.";
    ui.hoursTableContainer.innerHTML = "";
    return;
  }
  if (ui.hoursTableExportBtn) ui.hoursTableExportBtn.disabled = true;
  const { monthValue, monthMeta } = resolveHoursStatsMonth();
  if (!monthMeta) {
    ui.hoursTableFeedback.textContent = "Seleziona un mese valido per calcolare i totali per commessa.";
    ui.hoursTableContainer.innerHTML = "";
    return;
  }
  if (ui.hoursTableMonth) ui.hoursTableMonth.value = monthValue;
  if (ui.hoursTableCommessaSelect) ui.hoursTableCommessaSelect.disabled = true;
  ensureHoursViewModalOpen();
  ui.hoursTableFeedback.textContent = "Caricamento totale ore operatore per commessa...";
  ui.hoursTableContainer.innerHTML = "";
  try {
    const reports = await fetchHoursReportsForMonth(monthValue, monthMeta);
    const totals = new Map();
    reports.forEach((report) => {
      const entries = Array.isArray(report.entries) ? report.entries : [];
      entries.forEach((entry) => {
        const commessaId = String(entry.commessaId || "").trim();
        const commessaName = String(entry.commessaName || commesseById.get(commessaId)?.nome || "Commessa").trim();
        (Array.isArray(entry.rows) ? entry.rows : []).forEach((row) => {
          const operatore = String(row.operatore || "").trim();
          const ore = Number(row.ore || 0);
          if (!operatore || ore <= 0 || !commessaId) return;
          const key = `${operatore}__${commessaId}`;
          if (!totals.has(key)) totals.set(key, { operatore, commessaName, total: 0 });
          totals.get(key).total += ore;
        });
      });
    });
    const rows = Array.from(totals.values())
      .sort((a, b) => {
        const commessaCmp = a.commessaName.localeCompare(b.commessaName, "it");
        return commessaCmp || a.operatore.localeCompare(b.operatore, "it");
      })
      .map((item) => {
        const totalLabel = Number.isInteger(item.total) ? String(item.total) : item.total.toFixed(2).replace(".", ",");
        return `<tr><th scope="row">${escapeHTML(item.commessaName)}</th><td>${escapeHTML(item.operatore)}</td><td><b>${escapeHTML(totalLabel)}h</b></td></tr>`;
      });
    if (!rows.length) {
      ui.hoursTableFeedback.textContent = "Nessuna ora trovata per calcolare i totali per commessa.";
      ui.hoursTableContainer.innerHTML = "";
      return;
    }
    ui.hoursTableContainer.innerHTML = `
      <table class="hours-month-table">
        <thead>
          <tr>
            <th>Commessa</th>
            <th>Operatore</th>
            <th>Totale ore</th>
          </tr>
        </thead>
        <tbody>${rows.join("")}</tbody>
      </table>
    `;
    ui.hoursTableFeedback.textContent = `Totale ore operatore per commessa calcolato per ${monthValue}.`;
    hoursTableContext = {
      mode: "tot_operator_commessa",
      monthValue,
      rows: Array.from(totals.values()).sort((a, b) => {
        const c = a.commessaName.localeCompare(b.commessaName, "it");
        return c || a.operatore.localeCompare(b.operatore, "it");
      })
    };
    if (ui.hoursTableExportBtn) ui.hoursTableExportBtn.disabled = false;
  } catch (error) {
    console.error("Errore caricamento totale ore operatore per commessa:", error);
    ui.hoursTableFeedback.textContent = "Errore caricamento totale ore operatore per commessa.";
    ui.hoursTableContainer.innerHTML = "";
  }
}

function renderHoursMonthlyTable(reports, commessaId, monthMeta) {
  if (!ui.hoursTableFeedback || !ui.hoursTableContainer) return;
  const operatorsMap = new Map();
  hoursTableRowsMap = new Map();
  const formatHoursValue = (value) => (Number.isInteger(value) ? String(value) : Number(value || 0).toFixed(2).replace(".", ","));
  (Array.isArray(reports) ? reports : []).forEach((report) => {
    const reportDate = String(report.date || "").trim();
    const day = Number(reportDate.split("-")[2] || 0);
    if (!day || day < 1 || day > monthMeta.daysInMonth) return;
    const entries = Array.isArray(report.entries) ? report.entries : [];
    entries.forEach((entry, entryIndex) => {
      if (String(entry.commessaId || "") !== commessaId) return;
      (Array.isArray(entry.rows) ? entry.rows : []).forEach((row, rowIndex) => {
        const operatore = String(row.operatore || "").trim();
        const ore = Number(row.ore || 0);
        if (!operatore || ore <= 0) return;
        if (!operatorsMap.has(operatore)) {
          operatorsMap.set(operatore, Array.from({ length: monthMeta.daysInMonth }, () => []));
        }
        operatorsMap.get(operatore)[day - 1].push(ore);
        const key = `${operatore}__${day}`;
        if (!hoursTableRowsMap.has(key)) hoursTableRowsMap.set(key, []);
        hoursTableRowsMap.get(key).push({
          reportId: report.id,
          reportDate,
          entryCommessaId: entry.commessaId,
          entryIndex,
          rowIndex,
          operatore,
          ore
        });
      });
    });
  });

  const operators = Array.from(operatorsMap.keys()).sort((a, b) => a.localeCompare(b, "it"));
  const commessaName = commesseById.get(commessaId)?.nome || "Commessa";
  const daysHeader = Array.from({ length: monthMeta.daysInMonth }, (_, idx) => `<th>${idx + 1}</th>`).join("");
  const bodyRowsReal = operators.map((operatorName) => {
    const dayValues = operatorsMap.get(operatorName);
    const total = dayValues.reduce((sum, dayItems) => sum + dayItems.reduce((daySum, ore) => daySum + Number(ore || 0), 0), 0);
    const cells = dayValues.map((dayItems, idx) => {
      const day = idx + 1;
      if (!dayItems.length) return "<td>-</td>";
      const key = `${operatorName}__${day}`;
      const sources = hoursTableRowsMap.get(key) || [];
      const canManage = canManageData() && sources.length;
      const dayTotal = dayItems.reduce((sum, value) => sum + Number(value || 0), 0);
      const hasDuplicates = dayItems.length > 1;
      const valueLabel = hasDuplicates ? `⚠ ${dayItems.length} inserimenti` : `${formatHoursValue(dayTotal)}h`;
      const mergedDetails = hasDuplicates
        ? `Duplicato non valido: stesso operatore/commessa/giorno inserito più volte. Elimina le righe duplicate.`
        : "";
      const title = canManage
        ? hasDuplicates
          ? `${mergedDetails}`
          : `Modifica o elimina ${sources.length} registrazione/i di ${operatorName} del giorno ${day}. Ore salvate: ${valueLabel}.`
        : `${operatorName} - giorno ${day}`;
      return `<td><button type="button" class="hours-value-btn" data-hours-key="${escapeHTML(key)}" ${canManage ? "" : "disabled"} title="${escapeHTML(title)}">${escapeHTML(valueLabel)}</button></td>`;
    }).join("");
    const totalLabel = formatHoursValue(total);
    return `<tr><th scope="row">${escapeHTML(operatorName)}</th>${cells}<td><b>${escapeHTML(totalLabel)}h</b></td></tr>`;
  });
  const emptyRowsNeeded = Math.max(0, 10 - bodyRowsReal.length);
  const emptyCells = Array.from({ length: monthMeta.daysInMonth }, () => "<td>-</td>").join("");
  const emptyRows = Array.from({ length: emptyRowsNeeded }, () => (
    `<tr><th scope="row" class="muted">—</th>${emptyCells}<td class="muted">0h</td></tr>`
  ));
  const bodyRows = [...bodyRowsReal, ...emptyRows].join("");

  ui.hoursTableContainer.innerHTML = `
    <table class="hours-month-table">
      <thead>
        <tr>
          <th>Operatore</th>
          ${daysHeader}
          <th>Totale</th>
        </tr>
      </thead>
      <tbody>${bodyRows}</tbody>
    </table>
  `;
  const monthLabel = `${String(monthMeta.month).padStart(2, "0")}/${monthMeta.year}`;
  hoursTableContext = {
    mode: "monthly",
    monthLabel,
    commessaName,
    monthMeta,
    operators: operators.map((name) => ({
      name,
      dayValues: (operatorsMap.get(name) || []).map((items) => items.reduce((sum, ore) => sum + Number(ore || 0), 0))
    }))
  };

  if (ui.hoursTableExportBtn) ui.hoursTableExportBtn.disabled = false;
  if (!operators.length) {
    ui.hoursTableFeedback.textContent = "Nessuna ora trovata: mostro una tabella vuota (minimo 10 righe).";
  } else {
    ui.hoursTableFeedback.textContent = canManageData()
      ? "Clicca un valore per modificare o eliminare la registrazione ore."
      : "Vista sola lettura: solo l'amministratore può modificare o eliminare le ore.";
  }

  ui.hoursTableContainer.querySelectorAll(".hours-value-btn").forEach((btn) => {
    btn.addEventListener("click", () => handleHoursValueAction(btn.dataset.hoursKey || ""));
  });
  const today = new Date();
  if (today.getFullYear() === monthMeta.year && (today.getMonth() + 1) === monthMeta.month) {
    const todayDay = today.getDate();
    const todayHeaderCell = ui.hoursTableContainer.querySelector(`thead th:nth-child(${todayDay + 1})`);
    if (todayHeaderCell && typeof todayHeaderCell.offsetLeft === "number") {
      const left = Math.max(0, todayHeaderCell.offsetLeft - 220);
      ui.hoursTableContainer.scrollLeft = left;
    }
  } else {
    ui.hoursTableContainer.scrollLeft = 0;
  }
}

async function handleHoursValueAction(cellKey) {
  if (!canManageData()) return;
  let sources = hoursTableRowsMap.get(String(cellKey || ""));
  if (!sources || !sources.length) return;
  const action = window.prompt("Admin: scrivi M per modificare oppure E per eliminare.", "M");
  const normalizedAction = String(action || "").trim().toUpperCase();
  if (!normalizedAction) return;
  if (!["M", "E"].includes(normalizedAction)) {
    if (ui.hoursTableFeedback) ui.hoursTableFeedback.textContent = "Azione annullata: usa M (modifica) o E (elimina).";
    return;
  }
  if (sources.length > 1) {
    const details = sources.map((source, idx) => `${idx + 1}) ${source.reportDate || "-"} • ${source.operatore || "-"} • ${Number(source.ore || 0)}h`).join("\n");
    const choice = window.prompt(
      `Ci sono ${sources.length} registrazioni in questa cella:\n${details}\n\nScrivi il numero da aggiornare/eliminare oppure A per tutte.`,
      "A"
    );
    const normalizedChoice = String(choice || "").trim().toUpperCase();
    if (!normalizedChoice) return;
    if (normalizedChoice !== "A") {
      const selectedIndex = Number(normalizedChoice);
      if (!Number.isInteger(selectedIndex) || selectedIndex < 1 || selectedIndex > sources.length) {
        if (ui.hoursTableFeedback) ui.hoursTableFeedback.textContent = "Azione annullata: selezione registrazione non valida.";
        return;
      }
      sources = [sources[selectedIndex - 1]];
    }
  }
  let nextHoursValue = null;
  if (normalizedAction === "M") {
    const rawValue = window.prompt("Nuovo valore ore (esempio: 4 oppure 7.5).");
    const parsedValue = Number(String(rawValue || "").replace(",", "."));
    if (!Number.isFinite(parsedValue) || parsedValue <= 0) {
      if (ui.hoursTableFeedback) ui.hoursTableFeedback.textContent = "Modifica annullata: valore ore non valido.";
      return;
    }
    nextHoursValue = parsedValue;
  } else {
    const confirmed = window.confirm(`Confermi eliminazione di ${sources.length} registrazione/i?`);
    if (!confirmed) return;
  }
  if (ui.hoursTableFeedback) ui.hoursTableFeedback.textContent = normalizedAction === "M"
    ? "Modifica ore in corso..."
    : "Eliminazione ore in corso...";
  try {
    const groupedByReport = new Map();
    sources.forEach((source) => {
      if (!groupedByReport.has(source.reportId)) groupedByReport.set(source.reportId, []);
      groupedByReport.get(source.reportId).push(source);
    });
    for (const [reportId, reportSources] of groupedByReport.entries()) {
      const docRef = db.collection("oreReports").doc(reportId);
      const docSnap = await docRef.get();
      if (!docSnap.exists) continue;
      const data = docSnap.data() || {};
      const nextEntries = (Array.isArray(data.entries) ? data.entries : []).map((entry, entryIndex) => {
        const targetRows = reportSources
          .filter((source) => Number(source.entryIndex) === entryIndex)
          .map((source) => Number(source.rowIndex))
          .filter(Number.isInteger);
        if (!targetRows.length) return entry;
        const nextRows = (Array.isArray(entry.rows) ? entry.rows : []).map((row, rowIndex) => {
          if (!targetRows.includes(rowIndex)) return row;
          if (normalizedAction === "M") return { ...row, ore: nextHoursValue };
          return null;
        }).filter((row) => row && Number(row.ore || 0) > 0);
        return { ...entry, rows: nextRows };
      }).filter((entry) => Array.isArray(entry.rows) && entry.rows.length);
      if (nextEntries.length) {
        await docRef.update({ entries: nextEntries });
      } else {
        await docRef.delete();
      }
    }
    await loadSavedHoursReports();
    await loadHoursMonthlyTable();
  } catch (error) {
    console.error("Errore aggiornamento ore:", error);
    if (ui.hoursTableFeedback) ui.hoursTableFeedback.textContent = "Errore modifica/eliminazione ore.";
  }
}

async function exportHoursMonthlyTable() {
  const mode = String(hoursTableContext?.mode || "monthly");
  if (mode === "tot_operator") {
    const monthValue = String(hoursTableContext?.monthValue || ui.hoursStatsMonth?.value || "").trim();
    const rows = Array.isArray(hoursTableContext?.rows) ? hoursTableContext.rows : [];
    if (!rows.length) {
      alert("Nessun totale operatore da esportare.");
      return;
    }
    const aoa = [["Mese", monthValue], [], ["Operatore", "Totale ore"]];
    rows.forEach((row) => aoa.push([row.name, row.total]));
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(aoa), "Totale operatori");
    XLSX.writeFile(wb, `totale_operatori_${monthValue}.xlsx`);
    return;
  }
  if (mode === "tot_operator_commessa") {
    const monthValue = String(hoursTableContext?.monthValue || ui.hoursStatsMonth?.value || "").trim();
    const rows = Array.isArray(hoursTableContext?.rows) ? hoursTableContext.rows : [];
    if (!rows.length) {
      alert("Nessun totale operatore per commessa da esportare.");
      return;
    }
    const aoa = [["Mese", monthValue], [], ["Commessa", "Operatore", "Totale ore"]];
    rows.forEach((row) => aoa.push([row.commessaName, row.operatore, row.total]));
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(aoa), "Operatore x commessa");
    XLSX.writeFile(wb, `totale_operatori_commesse_${monthValue}.xlsx`);
    return;
  }

  const monthValue = String(ui.hoursTableMonth?.value || ui.hoursStatsMonth?.value || "").trim();
  const monthMeta = getMonthMeta(monthValue);
  const commessaId = String(ui.hoursTableCommessaSelect?.value || "").trim();
  if (!monthMeta || !commessaId) {
    alert("Seleziona mese e commessa prima di esportare Excel.");
    return;
  }
  const commessaName = commesseById.get(commessaId)?.nome || "Commessa";
  const monthLabel = `${String(monthMeta.month).padStart(2, "0")}/${monthMeta.year}`;
  const reports = await fetchHoursReportsForMonth(monthValue, monthMeta);

  const operatorDayMap = new Map();
  const operatorTotals = new Map();
  const operatorCommessaTotals = new Map();
  reports.forEach((report) => {
    const day = Number(String(report.date || "").split("-")[2] || 0);
    const entries = Array.isArray(report.entries) ? report.entries : [];
    entries.forEach((entry) => {
      const entryCommessaId = String(entry.commessaId || "").trim();
      const entryCommessaName = String(entry.commessaName || commesseById.get(entryCommessaId)?.nome || "Commessa").trim();
      (Array.isArray(entry.rows) ? entry.rows : []).forEach((row) => {
        const operatore = String(row.operatore || "").trim();
        const ore = Number(row.ore || 0);
        if (!operatore || ore <= 0) return;
        const operatorNorm = operatore.toLocaleLowerCase("it-IT").replace(/\s+/g, " ").trim();
        if (!operatorTotals.has(operatorNorm)) operatorTotals.set(operatorNorm, { name: operatore, total: 0 });
        operatorTotals.get(operatorNorm).total += ore;
        const byCommessaKey = `${operatorNorm}__${entryCommessaId}`;
        if (!operatorCommessaTotals.has(byCommessaKey)) operatorCommessaTotals.set(byCommessaKey, { operatore, commessaName: entryCommessaName, total: 0 });
        operatorCommessaTotals.get(byCommessaKey).total += ore;
        if (entryCommessaId !== commessaId || !day) return;
        if (!operatorDayMap.has(operatore)) operatorDayMap.set(operatore, Array(monthMeta.daysInMonth).fill(0));
        operatorDayMap.get(operatore)[day - 1] += ore;
      });
    });
  });

  const headerRow = ["Operatore"];
  for (let day = 1; day <= monthMeta.daysInMonth; day += 1) headerRow.push(String(day));
  headerRow.push("Totale");
  const aoa = [
    [`Commessa: ${commessaName}`],
    [`Mese: ${monthLabel}`],
    [],
    headerRow
  ];
  const selectedCommessaOperators = Array.from(operatorDayMap.keys()).sort((a, b) => a.localeCompare(b, "it"));
  selectedCommessaOperators.forEach((name) => {
    const values = operatorDayMap.get(name) || [];
    const total = values.reduce((sum, value) => sum + value, 0);
    aoa.push([name, ...values, total]);
  });
  while (aoa.length < 14) {
    aoa.push(["", ...Array(monthMeta.daysInMonth).fill(""), ""]);
  }
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, XLSX.utils.aoa_to_sheet(aoa), "Ore mensili");

  const operatorTotalsRows = [["Operatore", "Totale ore"]];
  Array.from(operatorTotals.values())
    .sort((a, b) => a.name.localeCompare(b.name, "it"))
    .forEach((item) => operatorTotalsRows.push([item.name, item.total]));
  XLSX.utils.book_append_sheet(workbook, XLSX.utils.aoa_to_sheet(operatorTotalsRows), "Totale operatori");

  const operatorCommessaRows = [["Commessa", "Operatore", "Totale ore"]];
  Array.from(operatorCommessaTotals.values())
    .sort((a, b) => {
      const c = a.commessaName.localeCompare(b.commessaName, "it");
      return c || a.operatore.localeCompare(b.operatore, "it");
    })
    .forEach((item) => operatorCommessaRows.push([item.commessaName, item.operatore, item.total]));
  XLSX.utils.book_append_sheet(workbook, XLSX.utils.aoa_to_sheet(operatorCommessaRows), "Operatore x commessa");

  const safeCommessa = String(commessaName || "commessa").replace(/[\\/:*?"<>|]/g, "_").replace(/\s+/g, "_");
  const safeMonth = monthLabel.replace("/", "-");
  XLSX.writeFile(workbook, `ore_${safeCommessa}_${safeMonth}.xlsx`);
}

function renderHoursOperatoriOptions() {
  if (!ui.hoursOperatoriOptions) return;
  ui.hoursOperatoriOptions.innerHTML = "";
  personaleRecords.forEach((person) => {
    const option = document.createElement("option");
    option.value = getPersonaleDisplayName(person);
    ui.hoursOperatoriOptions.appendChild(option);
  });
}

function renderHoursCommessaSelectOptions() {
  const selects = ui.hoursCommesseList ? ui.hoursCommesseList.querySelectorAll(".hours-commessa-select") : [];
  const commesse = Array.from(commesseById.values());
  renderHoursTableCommessaOptions(commesse);
  if (!selects.length) return;
  selects.forEach((select) => {
    const selectedValue = select.value;
    select.innerHTML = "<option value=''>Seleziona commessa</option>";
    commesse.forEach((commessa) => {
      const option = document.createElement("option");
      option.value = commessa.id;
      option.textContent = commessa.nome || "Commessa senza nome";
      select.appendChild(option);
    });
    if (selectedValue && commesse.some((commessa) => commessa.id === selectedValue)) {
      select.value = selectedValue;
    }
  });
}

function renderHoursTableCommessaOptions(commesseInput = null) {
  if (!ui.hoursTableCommessaSelect) return;
  const commesse = Array.isArray(commesseInput) ? commesseInput : Array.from(commesseById.values());
  const selectedValue = ui.hoursTableCommessaSelect.value;
  ui.hoursTableCommessaSelect.innerHTML = "<option value=''>Seleziona commessa</option>";
  commesse.forEach((commessa) => {
    const option = document.createElement("option");
    option.value = commessa.id;
    option.textContent = commessa.nome || "Commessa senza nome";
    ui.hoursTableCommessaSelect.appendChild(option);
  });
  if (selectedValue && commesse.some((commessa) => commessa.id === selectedValue)) {
    ui.hoursTableCommessaSelect.value = selectedValue;
  }
}

function addHoursOperatoreRow(container, rowData = { operatore: "", ore: "" }) {
  const row = document.createElement("div");
  row.className = "hours-operator-row";
  row.innerHTML = `
    <input type="text" class="hours-operatore" list="hours-operatori-options" placeholder="Operatore" value="${escapeHTML(rowData.operatore || "")}">
    <input type="number" class="hours-ore" min="0" max="24" step="0.25" placeholder="Ore" value="${escapeHTML(rowData.ore || "")}">
    <button type="button" class="btn hours-remove-operator-btn">Rimuovi</button>
  `;
  const removeBtn = row.querySelector(".hours-remove-operator-btn");
  removeBtn.addEventListener("click", () => {
    unlockHoursFinalizeButton();
    row.remove();
    renderHoursSummary();
  });
  row.querySelectorAll("input").forEach((input) => input.addEventListener("input", () => {
    unlockHoursFinalizeButton();
    renderHoursSummary();
  }));
  container.appendChild(row);
}

function addHoursCommessaBlock(blockData = null) {
  const card = document.createElement("article");
  card.className = "hours-commessa-card";
  card.innerHTML = `
    <div class="hours-commessa-head">
      <h3>Commessa</h3>
      <button type="button" class="btn hours-remove-commessa-btn">Rimuovi commessa</button>
    </div>
    <select class="hours-commessa-select" required>
      <option value="">Seleziona commessa</option>
    </select>
    <div class="hours-operator-list"></div>
    <div class="item-actions">
      <button type="button" class="btn hours-add-operator-btn">+ Aggiungi operatore</button>
    </div>
    <textarea class="hours-note" placeholder="Nota (opzionale)"></textarea>
  `;
  ui.hoursCommesseList.appendChild(card);
  renderHoursCommessaSelectOptions();

  const removeBtn = card.querySelector(".hours-remove-commessa-btn");
  removeBtn.addEventListener("click", () => {
    unlockHoursFinalizeButton();
    card.remove();
    if (!ui.hoursCommesseList.children.length) addHoursCommessaBlock();
    renderHoursSummary();
  });
  const operatorList = card.querySelector(".hours-operator-list");
  card.querySelector(".hours-add-operator-btn").addEventListener("click", () => {
    unlockHoursFinalizeButton();
    addHoursOperatoreRow(operatorList);
    renderHoursSummary();
  });
  card.querySelector(".hours-commessa-select").addEventListener("change", () => {
    unlockHoursFinalizeButton();
    renderHoursSummary();
  });
  card.querySelector(".hours-note").addEventListener("input", () => {
    unlockHoursFinalizeButton();
    renderHoursSummary();
  });

  if (blockData) {
    if (blockData.commessaId) card.querySelector(".hours-commessa-select").value = blockData.commessaId;
    card.querySelector(".hours-note").value = blockData.note || "";
    const rows = Array.isArray(blockData.rows) && blockData.rows.length ? blockData.rows : [{}];
    rows.forEach((row) => addHoursOperatoreRow(operatorList, row));
  } else {
    addHoursOperatoreRow(operatorList);
  }
  renderHoursSummary();
}

function collectHoursEntries() {
  const cards = Array.from(ui.hoursCommesseList.querySelectorAll(".hours-commessa-card"));
  return cards.map((card) => {
    const commessaId = String(card.querySelector(".hours-commessa-select")?.value || "").trim();
    const note = String(card.querySelector(".hours-note")?.value || "").trim();
    const rows = Array.from(card.querySelectorAll(".hours-operator-row")).map((row) => ({
      operatore: String(row.querySelector(".hours-operatore")?.value || "").trim(),
      ore: Number(row.querySelector(".hours-ore")?.value || 0)
    })).filter((row) => row.operatore && row.ore > 0);
    const commessaName = commesseById.get(commessaId)?.nome || "";
    return { commessaId, commessaName, note, rows };
  }).filter((entry) => entry.commessaId || entry.rows.length || entry.note);
}

function normalizeHoursOperatorName(value) {
  return String(value || "").toLocaleLowerCase("it-IT").replace(/\s+/g, " ").trim();
}

function findDuplicateHoursInDraft(entries) {
  const map = new Map();
  const duplicates = [];
  entries.forEach((entry) => {
    const commessaId = String(entry?.commessaId || "").trim();
    if (!commessaId) return;
    (entry.rows || []).forEach((row) => {
      const operatore = String(row?.operatore || "").trim();
      const normalizedOperator = normalizeHoursOperatorName(operatore);
      const ore = Number(row?.ore || 0);
      if (!normalizedOperator || ore <= 0) return;
      const key = `${commessaId}__${normalizedOperator}`;
      if (map.has(key)) {
        duplicates.push({
          commessaId,
          commessaName: entry.commessaName || commesseById.get(commessaId)?.nome || "Commessa",
          operatore
        });
        return;
      }
      map.set(key, true);
    });
  });
  return duplicates;
}

async function findExistingHoursConflicts(dateValue, entries, options = {}) {
  const skipApprovalRequestId = String(options?.skipApprovalRequestId || "").trim();
  const requestedKeys = new Map();
  entries.forEach((entry) => {
    const commessaId = String(entry?.commessaId || "").trim();
    if (!commessaId) return;
    (entry.rows || []).forEach((row) => {
      const operatore = String(row?.operatore || "").trim();
      const normalizedOperator = normalizeHoursOperatorName(operatore);
      const ore = Number(row?.ore || 0);
      if (!normalizedOperator || ore <= 0) return;
      requestedKeys.set(`${commessaId}__${normalizedOperator}`, {
        commessaId,
        commessaName: entry.commessaName || commesseById.get(commessaId)?.nome || "Commessa",
        operatore
      });
    });
  });
  if (!requestedKeys.size) return [];

  const [reportsSnapshot, approvalsSnapshot] = await Promise.all([
    db.collection("oreReports").where("date", "==", dateValue).get(),
    db.collection("oreApprovalRequests").where("date", "==", dateValue).get()
  ]);

  const conflicts = [];
  const collectConflicts = (reportEntries = []) => {
    reportEntries.forEach((entry) => {
      const commessaId = String(entry?.commessaId || "").trim();
      if (!commessaId) return;
      (entry.rows || []).forEach((row) => {
        const operatore = String(row?.operatore || "").trim();
        const normalizedOperator = normalizeHoursOperatorName(operatore);
        const ore = Number(row?.ore || 0);
        if (!normalizedOperator || ore <= 0) return;
        const match = requestedKeys.get(`${commessaId}__${normalizedOperator}`);
        if (!match) return;
        conflicts.push({
          commessaName: entry.commessaName || match.commessaName || commesseById.get(commessaId)?.nome || "Commessa",
          operatore: match.operatore || operatore
        });
      });
    });
  };
  reportsSnapshot.forEach((doc) => {
    const report = doc.data() || {};
    collectConflicts(Array.isArray(report.entries) ? report.entries : []);
  });
  approvalsSnapshot.forEach((doc) => {
    const request = doc.data() || {};
    if (skipApprovalRequestId && doc.id === skipApprovalRequestId) return;
    if (String(request.status || "").trim() === "rejected") return;
    collectConflicts(Array.isArray(request.entries) ? request.entries : []);
  });
  return conflicts;
}

function renderHoursSummary(forcedEntries = null) {
  const entries = forcedEntries || collectHoursEntries();
  hoursDraftEntries = entries;
  if (!ui.hoursSummary) return;
  if (!entries.length) {
    ui.hoursSummary.innerHTML = "<p class='muted'>Resoconto: aggiungi almeno una commessa.</p>";
    return;
  }
  const html = entries.map((entry, idx) => {
    const rows = entry.rows.length
      ? entry.rows.map((row) => `<li>${escapeHTML(row.operatore || "-")}: <b>${escapeHTML(String(row.ore || 0))}h</b></li>`).join("")
      : "<li>Nessun operatore indicato.</li>";
    return `
      <article class="item-card">
        <h3>${idx + 1}. ${escapeHTML(entry.commessaName || "Commessa non selezionata")}</h3>
        <ul>${rows}</ul>
        ${entry.note ? `<p><b>Nota:</b> ${escapeHTML(entry.note)}</p>` : ""}
      </article>
    `;
  }).join("");
  ui.hoursSummary.innerHTML = html;
}

function setHoursFinalizeLocked(locked) {
  hoursFinalizeLocked = Boolean(locked);
  if (ui.hoursFinalizeBtn) {
    ui.hoursFinalizeBtn.disabled = hoursFinalizeLocked;
  }
}

function unlockHoursFinalizeButton() {
  if (!hoursFinalizeLocked) return;
  setHoursFinalizeLocked(false);
  if (ui.hoursFeedback?.textContent && ui.hoursFeedback.textContent.includes("Richiesta inviata")) {
    ui.hoursFeedback.textContent = "Modifiche rilevate. Puoi premere di nuovo “Fine: salva”.";
  }
}

async function finalizeHoursReport(event) {
  event.preventDefault();
  if (hoursFinalizeLocked) {
    ui.hoursFeedback.textContent = "Richiesta già inviata. Modifica commesse/operatori/ore per riattivare “Fine: salva”.";
    return;
  }
  if (!currentUser) {
    ui.hoursFeedback.textContent = "Devi fare login prima di salvare.";
    return;
  }
  const dateValue = String(ui.hoursDate.value || "").trim();
  if (!dateValue) {
    ui.hoursFeedback.textContent = "Seleziona una data.";
    return;
  }
  const entries = collectHoursEntries();
  const hasValidEntry = entries.some((entry) => entry.commessaId && entry.rows.some((row) => row.operatore && row.ore > 0));
  if (!hasValidEntry) {
    ui.hoursFeedback.textContent = "Inserisci almeno una commessa con un operatore e ore > 0.";
    return;
  }
  const payload = {
    date: dateValue,
    entries: entries.map((entry) => ({
      commessaId: entry.commessaId,
      commessaName: entry.commessaName,
      note: entry.note,
      rows: entry.rows.filter((row) => row.operatore && row.ore > 0)
    })).filter((entry) => entry.commessaId && entry.rows.length),
    createdByUid: currentUser.uid,
    createdByName: currentUser.displayName || currentUser.email || "Operatore",
    createdByEmail: currentUser.email || "",
    createdAt: firebase.firestore.FieldValue.serverTimestamp()
  };
  ui.hoursFinalizeBtn.disabled = true;
  ui.hoursFeedback.textContent = "Invio richiesta approvazione in corso...";
  try {
    const duplicateDraft = findDuplicateHoursInDraft(payload.entries);
    if (duplicateDraft.length) {
      const first = duplicateDraft[0];
      ui.hoursFeedback.textContent = `Duplicato bloccato: ${first.operatore || "Operatore"} risulta già inserito nella commessa ${first.commessaName || "Commessa"} per questa data.`;
      return;
    }

    const existingConflicts = await findExistingHoursConflicts(dateValue, payload.entries);
    if (existingConflicts.length) {
      const first = existingConflicts[0];
      ui.hoursFeedback.textContent = `Inserimento bloccato: ${first.operatore || "Operatore"} ha già ore salvate nella commessa ${first.commessaName || "Commessa"} in data ${dateValue}. Per reinserire, elimina prima la riga dal visualizzatore ore.`;
      return;
    }

    const approvalRef = await db.collection("oreApprovalRequests").add({
      ...payload,
      status: "pending_level1",
      level1ApprovedBy: "",
      level1ApprovedAt: null,
      level2ApprovedBy: "",
      level2ApprovedAt: null,
      rejectedBy: "",
      rejectedAt: null,
      rejectionReason: "",
      finalizedReportId: ""
    });
    await notifyLevel1ForHoursApproval(approvalRef.id, payload);

    ui.hoursFeedback.textContent = `Richiesta inviata (ID ${approvalRef.id}). In attesa primo OK, poi conferma admin finale.`;
    ui.hoursCommesseList.innerHTML = "";
    addHoursCommessaBlock();
    setHoursFinalizeLocked(true);
    renderHoursSummary();
    loadSavedHoursReports();
  } catch (error) {
    console.error("Salvataggio gestione ore non riuscito:", error);
    ui.hoursFeedback.textContent = "Errore durante il salvataggio del resoconto ore.";
  } finally {
    if (!hoursFinalizeLocked) ui.hoursFinalizeBtn.disabled = false;
  }
}

async function notifyLevel1ForHoursApproval(requestId, payload) {
  const adminUsers = platformUsers.filter((user) => adminEmails.has(normalizeEmail(user.email)));
  const recipientsMap = new Map(adminUsers.map((user) => [user.id, user]));
  const requesterUser = platformUsers.find((user) => String(user.uid || "") === String(payload?.createdByUid || ""));
  if (requesterUser?.id) recipientsMap.set(requesterUser.id, requesterUser);
  const recipients = Array.from(recipientsMap.values());
  if (!recipients.length) return;
  const dateLabel = payload?.date
    ? new Date(`${payload.date}T00:00:00`).toLocaleDateString("it-IT")
    : "-";
  const author = payload?.createdByName || payload?.createdByEmail || "Operatore";
  const entriesCount = Array.isArray(payload?.entries) ? payload.entries.length : 0;
  const text = `🕒 Richiesta ore ${requestId} da ${author} (${dateLabel}). Commesse: ${entriesCount}. Serve primo OK.`;
  await Promise.all(recipients.map((recipient) => sendPrivateChatNotification({
    recipientId: recipient.id,
    text,
    senderName: currentUser.displayName || currentUser.email || "Operatore",
    metadata: {
      type: "hours_approval",
      approvalRequestId: requestId,
      action: "level1_ok"
    }
  })));
}

async function notifyAdminsForFinalHoursApproval(requestId, payload, approverName) {
  const adminUsers = platformUsers.filter((user) => adminEmails.has(normalizeEmail(user.email)));
  if (!adminUsers.length) return;
  const author = payload?.createdByName || payload?.createdByEmail || "Operatore";
  const text = `🕒 Richiesta ore ${requestId}: primo OK da ${approverName || "utente autorizzato"}. Attesa approvazione admin finale per ${author}.`;
  await Promise.all(adminUsers.map((adminUser) => sendPrivateChatNotification({
    recipientId: adminUser.id,
    text,
    senderName: currentUser.displayName || currentUser.email || "Sistema",
    metadata: {
      type: "hours_approval",
      approvalRequestId: requestId,
      action: "admin_final_ok"
    }
  })));
}

async function sendPrivateChatNotification({ recipientId, text, senderName = "Sistema", metadata = null }) {
  if (!recipientId || !text) return;
  await db.collection("chatMessages").add({
    type: "text",
    text,
    recipientId,
    senderId: currentUser?.uid || "system",
    senderName,
    senderEmail: currentUser?.email || "",
    kind: "system",
    metadata: metadata && typeof metadata === "object" ? metadata : null,
    createdAt: firebase.firestore.FieldValue.serverTimestamp()
  });
}

function renderHowtoFaq() {
  if (!ui.howtoFaqList) return;
  ui.howtoFaqList.innerHTML = "";
  const howtoFaqItems = buildHowtoFaqItems();
  howtoFaqItems.forEach((faq) => {
    const item = document.createElement("article");
    item.className = "howto-faq-item";
    item.dataset.faqId = faq.id;

    const toggleBtn = document.createElement("button");
    toggleBtn.type = "button";
    toggleBtn.className = "howto-faq-question";
    toggleBtn.setAttribute("aria-expanded", "false");
    toggleBtn.innerHTML = `
      <span>${escapeHTML(faq.domanda)}</span>
      <span class="howto-faq-meta">Agg. ${escapeHTML(faq.updatedAt)}</span>
    `;

    const answer = document.createElement("div");
    answer.className = "howto-faq-answer hidden";
    answer.innerHTML = `
      <p class="howto-faq-brief">${escapeHTML(faq.rispostaBreve)}</p>
      <ol class="howto-faq-steps">
        ${faq.passi.map((step) => `<li>${escapeHTML(step)}</li>`).join("")}
      </ol>
      <p class="howto-faq-tags">${faq.tags.map((tag) => `#${escapeHTML(tag)}`).join(" ")}</p>
    `;

    toggleBtn.addEventListener("click", () => {
      const isOpen = toggleBtn.getAttribute("aria-expanded") === "true";
      toggleBtn.setAttribute("aria-expanded", String(!isOpen));
      answer.classList.toggle("hidden", isOpen);
      item.classList.toggle("is-open", !isOpen);
    });

    item.append(toggleBtn, answer);
    ui.howtoFaqList.appendChild(item);
  });
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

function trackLocalSheetMutation(commessaId) {
  if (!commessaId) return;
  localSheetMutationAt.set(commessaId, Date.now());
}

function hasRecentLocalSheetMutation(commessaId, windowMs = 3500) {
  const ts = Number(localSheetMutationAt.get(commessaId) || 0);
  if (!ts) return false;
  return (Date.now() - ts) <= windowMs;
}

function scheduleCommessaSheetSync(commessaId, commessaName = "", delayMs = 900) {
  if (!commessaId) return;
  if (!driveAccessToken) return;
  const existingTimer = commessaSheetSyncTimers.get(commessaId);
  if (existingTimer) clearTimeout(existingTimer);
  const timer = setTimeout(async () => {
    try {
      await syncCommessaDoneImpiantiToDriveSheet(commessaId, commessaName || selectedCommessaName || "Commessa");
    } catch (error) {
      console.error("Sync foglio commessa fallita:", error);
      queuePendingSheetExport({ commessaId, commessaName });
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

function isAndroidWebViewRuntime() {
  // Flusso Android/WebView/Capacitor: qui evitiamo il redirect Firebase
  // perché in WebView può fallire con errore "missing initial state".
  const capacitorPlatform = (window.Capacitor && typeof window.Capacitor.getPlatform === "function")
    ? window.Capacitor.getPlatform()
    : "";
  const isCapacitorNative = Boolean(
    window.Capacitor
    && typeof window.Capacitor.isNativePlatform === "function"
    && window.Capacitor.isNativePlatform()
  );
  const ua = navigator.userAgent || "";
  const isAndroidUa = /Android/i.test(ua);
  const isWebViewUa = /; wv\)/i.test(ua) || /\bVersion\/[\d.]+ Chrome\/[\d.]+ Mobile\b/i.test(ua);
  return capacitorPlatform === "android" || isCapacitorNative || (isAndroidUa && isWebViewUa);
}

function loginWithGoogle(forceAccountSelection = false) {
  const provider = new firebase.auth.GoogleAuthProvider();
  provider.addScope("https://www.googleapis.com/auth/userinfo.email");
  provider.addScope("https://www.googleapis.com/auth/drive.file");
  if (forceAccountSelection) provider.setCustomParameters({ prompt: "select_account" });

  // Flusso dedicato Android (APK/WebView/Capacitor): NO redirect fallback.
  if (isAndroidWebViewRuntime()) {
    auth.signInWithPopup(provider).then((result) => {
      const accessToken = extractGoogleAccessToken(result);
      if (accessToken) {
        persistDriveAccessToken(accessToken);
        autoConnectDriveBridge({ notifyOnError: false });
      }
    }).catch((error) => {
      console.error("Login Google Android/WebView fallito:", error);
      alert("Errore login Android/WebView: " + error.message);
    });
    return;
  }

  // Flusso web desktop/browser standard: manteniamo comportamento esistente.
  auth.signInWithPopup(provider).then((result) => {
    // Gli scope extra (es. Drive) arrivano nel credential del login, non nel profilo utente Firebase.
    const accessToken = extractGoogleAccessToken(result);
    if (accessToken) {
      persistDriveAccessToken(accessToken);
      autoConnectDriveBridge({ notifyOnError: false });
    }
  }).catch((error) => {
    if (error.code === "auth/popup-blocked" || error.code === "auth/cancelled-popup-request") {
      return auth.signInWithRedirect(provider);
    }
    alert("Errore login: " + error.message);
  });
}

async function switchGoogleAccount() {
  try {
    await auth.signOut();
  } catch (error) {
    console.warn("Logout durante cambio account non riuscito:", error);
  }
  loginWithGoogle(true);
}

function persistDriveAccessToken(accessToken) {
  if (!accessToken) return;
  driveAccessToken = accessToken;
  window.googleDriveAccessToken = accessToken;
  localStorage.setItem("googleDriveAccessToken", accessToken);
  localStorage.setItem("googleDriveConnected", "true");
}

async function autoConnectDriveBridge(options = {}) {
  if (!driveAccessToken) return;
  const { notifyOnError = false } = options;
  try {
    await ensureDriveFolders();
    const driveUser = auth.currentUser || currentUser;
    await db.collection("appConfig").doc("driveBridge").set({
      ownerEmail: (driveUser && driveUser.email) ? driveUser.email : ADMIN_EMAIL,
      accessToken: driveAccessToken,
      rootFolderId: driveRootFolderId,
      chatFolderId: driveChatFolderId,
      reportsFolderId: driveReportsFolderId,
      squadreFolderId: driveSquadreFolderId,
      helpCenterFolderId: driveHelpCenterFolderId,
      updatedAt: firebase.firestore.FieldValue.serverTimestamp()
    }, { merge: true });
    updateDriveStatus(true);
  } catch (error) {
    console.error("Connessione automatica Drive non riuscita:", error);
    if (notifyOnError) {
      throw error;
    }
  }
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
    alert("Solo un admin può aggiungere commesse.");
    return;
  }

  const commessaRef = await db.collection("commesse").add({
    nome,
    creatoDa: user.email || "",
    createdAt: firebase.firestore.FieldValue.serverTimestamp()
  });

  if (driveAccessToken) {
    try {
      if (!driveReportsFolderId) await ensureDriveFolders();
      await getOrCreateCommessaSpreadsheet(commessaRef.id, nome);
    } catch (error) {
      console.error("Commessa creata ma foglio Google non inizializzato subito:", error);
    }
  }

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
      driveHelpCenterFolderId = data.helpCenterFolderId || "";
      ui.driveStatus.textContent = `Drive centralizzato attivo (${owner}).`;
      processPendingSheetExports();
      processAdminSheetExportQueue();
      return;
    }

    const storedToken = getStoredDriveToken();
    if (storedToken) {
      driveAccessToken = storedToken;
      window.googleDriveAccessToken = storedToken;
    }
    driveRootFolderId = data.rootFolderId || driveRootFolderId || "";
    driveChatFolderId = data.chatFolderId || driveChatFolderId || "";
    driveReportsFolderId = data.reportsFolderId || driveReportsFolderId || "";
    driveSquadreFolderId = data.squadreFolderId || driveSquadreFolderId || "";
    driveHelpCenterFolderId = data.helpCenterFolderId || driveHelpCenterFolderId || "";
    if (driveAccessToken) {
      ui.driveStatus.textContent = `Drive collegato (account utente) • report centralizzati (${owner}).`;
    } else {
      ui.driveStatus.textContent = `Report foglio gestito dall'admin (${owner}).`;
    }
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
      ui.commessaTargetSelect.innerHTML = "<option value=''>Usa commessa selezionata in home</option>";
      ui.resourceCommesse.innerHTML = "";

      if (snapshot.empty) {
        ui.commesseLista.innerHTML = "<p class='muted'>Nessuna commessa disponibile.</p>";
        renderHoursCommessaSelectOptions();
        updateCommessaContextUI();
        renderNextActionCard();
        return;
      }

      const activeStoredId = localStorage.getItem(LAST_OPENED_COMMESSA_KEY) || localStorage.getItem(LAST_SELECTED_COMMESSA_KEY) || "";
      let shouldRestoreOpenCommessa = false;
      snapshot.forEach((doc, idx) => {
        const commessa = doc.data();
        commesseById.set(doc.id, { id: doc.id, ...commessa });
        const row = document.createElement("div");
        row.className = "commessa-row";

        const btn = document.createElement("button");
        btn.type = "button";
        btn.className = "btn commessa-btn" + (doc.id === selectedCommessaId ? " active" : "");
        btn.dataset.commessaId = doc.id;
        btn.style.setProperty("--commessa-accent", getCommessaAccentColor(doc.id, idx));
        btn.textContent = commessa.nome || "Commessa senza nome";
        btn.addEventListener("click", () => selectCommessa(doc.id, commessa.nome || "Commessa"));

        row.appendChild(btn);
        ui.commesseLista.appendChild(row);

        const option = document.createElement("option");
        option.value = doc.id;
        option.textContent = commessa.nome || "Commessa senza nome";
        ui.squadraCommessa.appendChild(option);
        ui.commessaTargetSelect.appendChild(option.cloneNode(true));
        ui.resourceCommesse.appendChild(option.cloneNode(true));
        if (!selectedCommessaId && activeStoredId && activeStoredId === doc.id) shouldRestoreOpenCommessa = true;
      });

      renderCommesseManagementList();
      renderHoursCommessaSelectOptions();
      renderSquadre();
      renderResourcesList();
      renderResourceButtonsForCommessa();
      syncBannerFormFromSelection();
      updateCommessaContextUI();
      if (!selectedCommessaId && shouldRestoreOpenCommessa) {
        const restored = commesseById.get(activeStoredId);
        if (restored) selectCommessa(restored.id, restored.nome || "Commessa");
      }
      renderNextActionCard();
    }, (error) => {
      console.error(error);
      ui.commesseLista.innerHTML = "<p class='muted'>Errore caricamento commesse.</p>";
      renderNextActionCard();
    });
}

function stopCommesseSubscription() {
  if (unsubscribeCommesse) {
    unsubscribeCommesse();
    unsubscribeCommesse = null;
  }
}

function subscribeResources() {
  unsubscribeResources = db
    .collection("commessaResources")
    .orderBy("createdAt", "desc")
    .onSnapshot((snapshot) => {
      resourceRecords = snapshot.docs.map((doc) => ({ id: doc.id, ...doc.data() }));
      renderResourcesList();
      renderResourceButtonsForCommessa();
      renderCommessaResourceViewer();
    }, (error) => {
      console.error("Errore caricamento informazioni utili:", error);
      ui.resourcesList.innerHTML = "<p class='muted'>Errore caricamento informazioni utili.</p>";
    });
}

function stopResourcesSubscription() {
  if (unsubscribeResources) {
    unsubscribeResources();
    unsubscribeResources = null;
  }
}

function subscribePrivateDocs() {
  if (!currentUser) return;
  unsubscribePrivateDocs = db
    .collection("privateDocuments")
    .doc(currentUser.uid)
    .collection("items")
    .orderBy("createdAt", "desc")
    .onSnapshot((snapshot) => {
      privateDocsRecords = snapshot.docs.map((doc) => ({ id: doc.id, ...doc.data() }));
      renderPrivateDocsList();
    }, (error) => {
      console.error("Errore caricamento documenti personali:", error);
      ui.privateDocsFeedback.textContent = "Errore caricamento documenti personali.";
    });
}

function stopPrivateDocsSubscription() {
  if (unsubscribePrivateDocs) {
    unsubscribePrivateDocs();
    unsubscribePrivateDocs = null;
  }
}

function applyPrivateDocPreset(type) {
  if (!ui.privateDocsName || !ui.privateDocsNote) return;
  if (type === "pin") {
    ui.privateDocsName.value = "PIN carburante";
    ui.privateDocsNote.value = "Inserisci qui il PIN della carta carburante.";
    return;
  }
  if (type === "tessera") {
    ui.privateDocsName.value = "Tessera di riconoscimento";
    ui.privateDocsNote.value = "Documento personale di riconoscimento.";
  }
}

async function readFileAsDataUrl(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(String(reader.result || ""));
    reader.onerror = () => reject(new Error("Lettura file non riuscita"));
    reader.readAsDataURL(file);
  });
}

function getPrivateDocsDriveToken() {
  return String(localStorage.getItem("googleDriveAccessToken") || "").trim();
}

async function driveApiFetchWithToken(token, url, options = {}) {
  const headers = new Headers(options.headers || {});
  headers.set("Authorization", `Bearer ${token}`);
  const response = await fetch(url, { ...options, headers });
  if (!response.ok) {
    const text = await response.text();
    throw new Error(`Google Drive API ${response.status}: ${text}`);
  }
  return response.status === 204 ? null : response.json();
}

async function getOrCreatePrivateDocsFolder(token, uid) {
  const query = [
    "name='Hera App - Documenti privati'",
    "mimeType='application/vnd.google-apps.folder'",
    "trashed=false"
  ].join(" and ");
  const rootSearch = await driveApiFetchWithToken(token, `https://www.googleapis.com/drive/v3/files?q=${encodeURIComponent(query)}&fields=files(id,name)&pageSize=1`, { method: "GET" });
  let rootFolderId = rootSearch?.files?.[0]?.id || "";
  if (!rootFolderId) {
    const createdRoot = await driveApiFetchWithToken(token, "https://www.googleapis.com/drive/v3/files", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ name: "Hera App - Documenti privati", mimeType: "application/vnd.google-apps.folder" })
    });
    rootFolderId = createdRoot.id;
  }

  const userQuery = [
    `name='${String(uid || "").replace(/'/g, "\\'")}'`,
    "mimeType='application/vnd.google-apps.folder'",
    "trashed=false",
    `'${rootFolderId}' in parents`
  ].join(" and ");
  const userSearch = await driveApiFetchWithToken(token, `https://www.googleapis.com/drive/v3/files?q=${encodeURIComponent(userQuery)}&fields=files(id,name)&pageSize=1`, { method: "GET" });
  const existingUserFolder = userSearch?.files?.[0]?.id || "";
  if (existingUserFolder) return existingUserFolder;

  const createdUserFolder = await driveApiFetchWithToken(token, "https://www.googleapis.com/drive/v3/files", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ name: uid, mimeType: "application/vnd.google-apps.folder", parents: [rootFolderId] })
  });
  return createdUserFolder.id;
}

async function uploadPrivateDocumentToDrive(file, uid) {
  const token = getPrivateDocsDriveToken();
  if (!token) {
    throw new Error("Google Drive non autorizzato. Rifai il login Google prima di usare il salvataggio Drive.");
  }
  const folderId = await getOrCreatePrivateDocsFolder(token, uid);
  const metadata = {
    name: file.name || "documento",
    parents: [folderId]
  };
  const boundary = "hera-private-doc-upload";
  const body = [
    `--${boundary}`,
    "Content-Type: application/json; charset=UTF-8",
    "",
    JSON.stringify(metadata),
    `--${boundary}`,
    `Content-Type: ${file.type || "application/octet-stream"}`,
    "",
    file,
    `--${boundary}--`
  ];
  const payload = new Blob(body, { type: `multipart/related; boundary=${boundary}` });
  const uploaded = await driveApiFetchWithToken(token, "https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart&fields=id,name,webViewLink", {
    method: "POST",
    headers: { "Content-Type": `multipart/related; boundary=${boundary}` },
    body: payload
  });
  return {
    driveFileId: uploaded.id || "",
    driveFileName: uploaded.name || file.name || "documento",
    driveWebViewLink: uploaded.webViewLink || `https://drive.google.com/file/d/${uploaded.id}/view`
  };
}

async function savePrivateDocument(event) {
  event.preventDefault();
  if (!currentUser) return;
  try {
    const name = String(ui.privateDocsName.value || "").trim();
    const note = String(ui.privateDocsNote.value || "").trim();
    const file = ui.privateDocsFile.files?.[0] || ui.privateDocsCamera.files?.[0] || null;
    if (!name) {
      ui.privateDocsFeedback.textContent = "La denominazione è obbligatoria.";
      return;
    }
    let fileDataUrl = "";
    let fileName = "";
    let fileType = "";
    let fileSize = 0;
    let driveFileId = "";
    let driveWebViewLink = "";
    const useDriveUpload = Boolean(ui.privateDocsDriveOnly?.checked);
    if (file) {
      fileSize = Number(file.size || 0);
      if (!useDriveUpload && fileSize > 700 * 1024) {
        ui.privateDocsFeedback.textContent = "File troppo grande: usa il salvataggio Drive o file sotto 700KB.";
        return;
      }
      fileName = file.name || "documento";
      fileType = file.type || "application/octet-stream";
      if (useDriveUpload) {
        ui.privateDocsFeedback.textContent = "Caricamento su Google Drive personale...";
        const upload = await uploadPrivateDocumentToDrive(file, currentUser.uid);
        driveFileId = upload.driveFileId;
        driveWebViewLink = upload.driveWebViewLink;
        fileName = upload.driveFileName || fileName;
      } else {
        fileDataUrl = await readFileAsDataUrl(file);
      }
    }
    await db.collection("privateDocuments").doc(currentUser.uid).collection("items").add({
      name,
      note,
      fileName,
      fileType,
      fileSize,
      fileDataUrl,
      driveFileId,
      driveWebViewLink,
      storageMode: driveFileId ? "drive" : "firestore",
      createdAt: firebase.firestore.FieldValue.serverTimestamp(),
      updatedAt: firebase.firestore.FieldValue.serverTimestamp()
    });
    ui.privateDocsForm.reset();
    ui.privateDocsFeedback.textContent = "Documento personale salvato.";
  } catch (error) {
    console.error("Salvataggio documento personale non riuscito:", error);
    ui.privateDocsFeedback.textContent = error?.message || "Errore durante il salvataggio del documento.";
  }
}

async function deletePrivateDocument(docId) {
  if (!currentUser || !docId) return;
  const ok = window.confirm("Eliminare questo documento personale?");
  if (!ok) return;
  await db.collection("privateDocuments").doc(currentUser.uid).collection("items").doc(docId).delete();
}

function renderPrivateDocsList() {
  if (!ui.privateDocsList) return;
  if (!currentUser) {
    ui.privateDocsList.innerHTML = "<p class='muted'>Fai login per usare i documenti personali.</p>";
    return;
  }
  if (!privateDocsRecords.length) {
    ui.privateDocsList.innerHTML = "<p class='muted'>Nessun documento personale salvato.</p>";
    return;
  }
  ui.privateDocsList.innerHTML = "";
  privateDocsRecords.forEach((item) => {
    const row = document.createElement("div");
    row.className = "simple-list-item stacked";
    const createdAt = item.createdAt?.toDate ? item.createdAt.toDate().toLocaleString("it-IT") : "-";
    row.innerHTML = `
      <div>
        <strong>${escapeHTML(item.name || "Documento")}</strong>
        <p class="muted">${escapeHTML(item.note || "-")}</p>
        <p class="muted">Data inserimento: ${escapeHTML(createdAt)}</p>
      </div>
    `;
    const actions = document.createElement("div");
    actions.className = "actions-row";
    if (item.driveWebViewLink) {
      actions.appendChild(createButton("Apri su Drive", () => window.open(item.driveWebViewLink, "_blank")));
    } else if (item.fileDataUrl) {
      actions.appendChild(createButton("Apri allegato", () => window.open(item.fileDataUrl, "_blank")));
    }
    actions.appendChild(createButton("Elimina", () => deletePrivateDocument(item.id)));
    row.appendChild(actions);
    ui.privateDocsList.appendChild(row);
  });
}

async function addResourceItem(event) {
  event.preventDefault();
  if (!canManageData()) {
    alert("Solo l'admin può inserire informazioni utili.");
    return;
  }
  const type = String(ui.resourceType.value || "").trim();
  const title = String(ui.resourceTitle.value || "").trim();
  const value = String(ui.resourceValue.value || "").trim();
  const commessaIds = Array.from(ui.resourceCommesse.selectedOptions || []).map((opt) => opt.value).filter(Boolean);
  if (!type || !title || !value || !commessaIds.length) {
    alert("Compila tutti i campi e seleziona almeno una commessa.");
    return;
  }
  await db.collection("commessaResources").add({
    type,
    title,
    value,
    commessaIds,
    createdAt: firebase.firestore.FieldValue.serverTimestamp(),
    createdBy: currentUser?.email || ""
  });
  ui.resourceForm.reset();
}

async function deleteResourceItem(resourceId) {
  if (!canManageData()) return;
  const ok = window.confirm("Eliminare questa informazione utile?");
  if (!ok) return;
  await db.collection("commessaResources").doc(resourceId).delete();
}

function renderResourcesList() {
  if (!ui.resourcesList) return;
  if (!activeResourceManageFilter) {
    ui.resourcesList.innerHTML = "<p class='muted'>Seleziona una categoria (📞 / 📄 / 📝) per vedere l'archivio.</p>";
    return;
  }
  const visibleResources = resourceRecords.filter((item) => item.type === activeResourceManageFilter);
  if (!visibleResources.length) {
    ui.resourcesList.innerHTML = "<p class='muted'>Nessuna informazione utile caricata.</p>";
    return;
  }
  ui.resourcesList.innerHTML = "";
  visibleResources.forEach((item) => {
    const row = document.createElement("div");
    row.className = "simple-list-item";
    const commesseNames = (item.commessaIds || [])
      .map((id) => (commesseById.get(id) || {}).nome || "Commessa")
      .join(", ");
    row.innerHTML = `
      <div>
        <strong>${resourceTypeLabel(item.type)} · ${escapeHTML(item.title || "-")}</strong>
        <p class="muted">${escapeHTML(item.value || "-")}</p>
        <p class="muted">Commesse: ${escapeHTML(commesseNames || "-")}</p>
      </div>
    `;
    if (canManageData()) {
      row.appendChild(createButton("Elimina", () => deleteResourceItem(item.id)));
    }
    ui.resourcesList.appendChild(row);
  });
}

function renderResourceManageFilters() {
  document.querySelectorAll(".resource-filter-btn").forEach((btn) => {
    const isActive = (btn.dataset.resourceFilter || "") === activeResourceManageFilter;
    btn.classList.toggle("btn-primary", isActive);
  });
}

function resourceTypeLabel(type) {
  if (type === "phone") return "📞";
  if (type === "document") return "📄";
  if (type === "note") return "📝";
  return "Info";
}

function getResourcesByCommessa(commessaId, type = "") {
  if (!commessaId) return [];
  return resourceRecords.filter((item) => {
    const linked = Array.isArray(item.commessaIds) && item.commessaIds.includes(commessaId);
    if (!linked) return false;
    return type ? item.type === type : true;
  });
}

function renderResourceButtonsForCommessa() {
  if (!ui.commessaResourceButtons) return;
  ui.commessaResourceButtons.innerHTML = "";
  if (!selectedCommessaId) return;
  const types = ["phone", "document", "note"];
  types.forEach((type) => {
    const count = getResourcesByCommessa(selectedCommessaId, type).length;
    if (!count) return;
    const label = `${resourceTypeLabel(type)} ${count}`;
    const btn = createButton(label, () => openCommessaResourceWindow(type));
    btn.title = resourceTypeLongLabel(type);
    btn.setAttribute("aria-label", `${resourceTypeLongLabel(type)} (${count})`);
    ui.commessaResourceButtons.appendChild(btn);
  });
}

function openCommessaResourceWindow(type) {
  activeResourceTypeForViewer = type;
  if (!selectedCommessaId) return;
  window.location.hash = `commessa=${selectedCommessaId}&resource=${type}`;
  applyRoute();
}

function closeCommessaResourceViewer() {
  if (window.location.hash.includes("&resource=") && selectedCommessaId) {
    window.location.hash = `commessa=${selectedCommessaId}`;
  }
  activeResourceTypeForViewer = "";
  ui.commessaResourceViewer.classList.remove("page-mode");
  ui.commessaResourceViewer.classList.add("hidden");
  ui.commessaResourceViewerCloseBtn.textContent = "Chiudi";
}

function renderCommessaResourceViewer() {
  if (!selectedCommessaId || !activeResourceTypeForViewer) return;
  const items = getResourcesByCommessa(selectedCommessaId, activeResourceTypeForViewer);
  ui.commessaResourceViewerTitle.textContent = `${resourceTypeLongLabel(activeResourceTypeForViewer)} • ${selectedCommessaName || "Commessa"}`;
  ui.commessaResourceViewerList.innerHTML = "";
  if (!items.length) {
    ui.commessaResourceViewerList.innerHTML = "<p class='muted'>Nessun contenuto disponibile.</p>";
    return;
  }
  items.forEach((item) => {
    const row = document.createElement("div");
    row.className = "simple-list-item";
    const info = document.createElement("div");
    info.innerHTML = `
      <strong>${escapeHTML(item.title || "-")}</strong>
      <p class="muted">${escapeHTML(item.value || "-")}</p>
    `;
    row.appendChild(info);

    const actions = document.createElement("div");
    actions.className = "actions-row";
    if (activeResourceTypeForViewer === "phone") {
      actions.appendChild(createButton("Chiama", () => window.open(`tel:${sanitizePhone(item.value)}`, "_self")));
      actions.appendChild(createButton("SMS", () => window.open(`sms:${sanitizePhone(item.value)}`, "_self")));
      actions.appendChild(createButton("WhatsApp", () => openPhoneOnWhatsApp(item.value)));
      actions.appendChild(createButton("Salva contatto", () => downloadVCard(item.title, item.value)));
    } else if (activeResourceTypeForViewer === "document") {
      actions.appendChild(createButton("Apri documento", () => openDocumentLink(item.value)));
    } else {
      actions.appendChild(createButton("Copia nota", () => navigator.clipboard?.writeText(String(item.value || ""))));
    }
    row.appendChild(actions);
    ui.commessaResourceViewerList.appendChild(row);
  });
}

function resourceTypeLongLabel(type) {
  if (type === "phone") return "Agenda";
  if (type === "document") return "Documenti";
  if (type === "note") return "Note";
  return "Informazioni";
}

function updateResourceFormByType() {
  const hasCategory = Boolean(String(ui.resourceType?.value || "").trim());
  [ui.resourceTitle, ui.resourceValue, ui.resourceCommesse, ui.resourceSubmit].forEach((el) => {
    if (!el) return;
    el.classList.toggle("hidden", !hasCategory);
    if (!hasCategory && el.tagName === "SELECT" && el.hasAttribute("multiple")) {
      Array.from(el.options || []).forEach((opt) => { opt.selected = false; });
    }
  });
}

function sanitizePhone(value) {
  return String(value || "").replace(/[^0-9+]/g, "");
}

function openPhoneOnWhatsApp(value) {
  const raw = sanitizePhone(value).replace("+", "");
  if (!raw) return;
  window.open(`https://wa.me/${raw}`, "_blank");
}

function openDocumentLink(value) {
  const raw = String(value || "").trim();
  if (!raw) return;
  const normalized = /^https?:\/\//i.test(raw) ? raw : `https://${raw}`;
  window.open(normalized, "_blank");
}

function downloadVCard(name, phone) {
  const contactName = String(name || "Contatto").replace(/\n/g, " ").trim();
  const cleanPhone = sanitizePhone(phone);
  const vcf = [
    "BEGIN:VCARD",
    "VERSION:3.0",
    `FN:${contactName}`,
    `TEL;TYPE=CELL:${cleanPhone}`,
    "END:VCARD"
  ].join("\n");
  const blob = new Blob([vcf], { type: "text/vcard;charset=utf-8" });
  const link = document.createElement("a");
  link.href = URL.createObjectURL(blob);
  link.download = `${contactName.replace(/[^\w\-]+/g, "_") || "contatto"}.vcf`;
  document.body.appendChild(link);
  link.click();
  link.remove();
  URL.revokeObjectURL(link.href);
}

function selectCommessa(id, nome) {
  selectedCommessaId = id;
  selectedCommessaName = nome;
  activeNearbyImpiantoContext = null;
  localStorage.setItem(LAST_SELECTED_COMMESSA_KEY, id);
  setImpiantiViewMode("todo");
  if (!ui.commessaTargetSelect.value) {
    ui.commessaTargetSelect.value = id;
  }
  ui.commessaAttiva.textContent = `Commessa selezionata: ${nome}`;
  updateCommessaContextUI();
  ui.importBtn.disabled = !auth.currentUser || pendingRows.length === 0 || !getTargetCommessaId() || !canManageData();
  ui.exportCurrentCommessaBtn.disabled = !auth.currentUser || !canManageData();
  updateCommessaButtonsActive();
  renderResourceButtonsForCommessa();
  closeCommessaResourceViewer();

  stopImpiantiSubscription();
  subscribeImpianti();
  setCurrentWorkflowStep("open-commessa");
  openImpiantiPage();
}

function updateCommessaContextUI() {
  if (ui.commessaFocusLabel) {
    ui.commessaFocusLabel.textContent = (selectedCommessaName || "Commessa").toUpperCase();
  }
}

function updateCommessaButtonsActive() {
  const buttons = ui.commesseLista.querySelectorAll(".commessa-btn");
  buttons.forEach((btn) => {
    const isActive = btn.dataset.commessaId === selectedCommessaId;
    btn.classList.toggle("active", isActive);
  });
}

function onCommessaTargetChanged() {
  const targetId = getTargetCommessaId();
  const targetName = getTargetCommessaName();
  ui.importBtn.disabled = !auth.currentUser || pendingRows.length === 0 || !targetId || !canManageData();
  if (targetId) {
    ui.commessaAttiva.textContent = `Commessa selezionata: ${targetName}`;
  }
}

function getTargetCommessaId() {
  return ui.commessaTargetSelect.value || selectedCommessaId || "";
}

function getTargetCommessaName() {
  const targetId = getTargetCommessaId();
  if (!targetId) return "";
  return (commesseById.get(targetId) || {}).nome || selectedCommessaName || "Commessa";
}

function subscribeImpianti() {
  if (!selectedCommessaId) return;
  let previousDoneSignature = null;

  unsubscribeImpianti = db
    .collection("commesse")
    .doc(selectedCommessaId)
    .collection("impianti")
    .onSnapshot((snapshot) => {
      const rawImpianti = snapshot.docs.map((doc) => ({ id: doc.id, ...doc.data() }));
      currentImpianti = combineImpiantiForView(rawImpianti);
      renderHeaderActivitySummary();
      renderImpianti();
      renderMap();
      evaluateImpiantoProximityAlerts();

      const currentDoneSignature = rawImpianti
        .filter((impianto) => Boolean(impianto.done))
        .map((impianto) => `${impianto.id}__${firestoreDateToMillis(impianto.doneAt)}`)
        .sort()
        .join("|");
      const doneStateChanged = previousDoneSignature !== null && currentDoneSignature !== previousDoneSignature;

      if (doneStateChanged && !hasRecentLocalSheetMutation(selectedCommessaId)) {
        scheduleCommessaSheetSync(selectedCommessaId, selectedCommessaName, 700);
      }

      previousDoneSignature = currentDoneSignature;
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
  activeNearbyImpiantoContext = null;
  renderHeaderActivitySummary();
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
      ui.importBtn.disabled = !auth.currentUser || !getTargetCommessaId() || pendingRows.length === 0 || !canManageData();
    } catch (error) {
      console.error(error);
      ui.importFeedback.textContent = "Errore lettura Excel.";
    }
  };

  reader.readAsArrayBuffer(file);
}

async function importPendingRows() {
  const user = auth.currentUser;
  const targetCommessaId = getTargetCommessaId();
  const targetCommessaName = getTargetCommessaName();

  if (!user || !targetCommessaId) {
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

  const totalPending = pendingRows.length;
  const ref = db.collection("commesse").doc(targetCommessaId).collection("impianti");
  const existingSnapshot = await ref.get();
  const existingByKey = new Map();
  existingSnapshot.forEach((doc) => {
    const data = doc.data() || {};
    const key = buildImpiantoKey(data);
    if (!existingByKey.has(key)) {
      existingByKey.set(key, { id: doc.id, ...data });
      return;
    }
    const merged = existingByKey.get(key);
    merged.codicePrezzo = mergeMultiValue(merged.codicePrezzo, data.codicePrezzo);
    merged.voceRiferimento = mergeMultiValue(merged.voceRiferimento, data.voceRiferimento);
    merged.tipologiaIntervento = mergeMultiValue(merged.tipologiaIntervento, data.tipologiaIntervento);
    merged.lavorazioniRichieste = mergeMultiValue(merged.lavorazioniRichieste, data.lavorazioniRichieste);
    merged.frequenzaAnnua = mergeMultiValue(merged.frequenzaAnnua, data.frequenzaAnnua);
  });

  const rowsToCreate = [];
  const rowsToUpdate = [];
  pendingRows.forEach((row) => {
    const key = buildImpiantoKey(row);
    const existing = existingByKey.get(key);
    if (!existing) {
      rowsToCreate.push(row);
      return;
    }
    const mergedCodicePrezzo = mergeMultiValue(existing.codicePrezzo, row.codicePrezzo);
    const mergedVoceRiferimento = mergeMultiValue(existing.voceRiferimento, row.voceRiferimento);
    const mergedTipologiaIntervento = mergeMultiValue(existing.tipologiaIntervento, row.tipologiaIntervento);
    const mergedLavorazioniRichieste = mergeMultiValue(existing.lavorazioniRichieste, row.lavorazioniRichieste);
    const mergedFrequenzaAnnua = mergeMultiValue(existing.frequenzaAnnua, row.frequenzaAnnua);
    const changed = mergedCodicePrezzo !== String(existing.codicePrezzo || "")
      || mergedVoceRiferimento !== String(existing.voceRiferimento || "")
      || mergedTipologiaIntervento !== String(existing.tipologiaIntervento || "")
      || mergedLavorazioniRichieste !== String(existing.lavorazioniRichieste || "")
      || mergedFrequenzaAnnua !== String(existing.frequenzaAnnua || "");
    if (!changed) return;
    rowsToUpdate.push({
      id: existing.id,
      codicePrezzo: mergedCodicePrezzo,
      voceRiferimento: mergedVoceRiferimento,
      tipologiaIntervento: mergedTipologiaIntervento,
      lavorazioniRichieste: mergedLavorazioniRichieste,
      frequenzaAnnua: mergedFrequenzaAnnua
    });
    existing.codicePrezzo = mergedCodicePrezzo;
    existing.voceRiferimento = mergedVoceRiferimento;
    existing.tipologiaIntervento = mergedTipologiaIntervento;
    existing.lavorazioniRichieste = mergedLavorazioniRichieste;
    existing.frequenzaAnnua = mergedFrequenzaAnnua;
  });

  const operations = [
    ...rowsToCreate.map((row) => ({ type: "create", row })),
    ...rowsToUpdate.map((row) => ({ type: "update", row }))
  ];
  for (let i = 0; i < operations.length; i += 450) {
    const chunk = operations.slice(i, i + 450);
    const batch = db.batch();
    chunk.forEach((operation) => {
      if (operation.type === "create") {
        const row = operation.row;
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
      } else {
        const row = operation.row;
        batch.update(ref.doc(row.id), {
          codicePrezzo: row.codicePrezzo,
          voceRiferimento: row.voceRiferimento,
          tipologiaIntervento: row.tipologiaIntervento,
          lavorazioniRichieste: row.lavorazioniRichieste,
          frequenzaAnnua: row.frequenzaAnnua,
          hasOrdinario: hasOrdinario(row.codicePrezzo),
          hasStraordinario: hasStraordinario(row.codicePrezzo),
          tipoManutenzione: classifyTipoManutenzione(row.codicePrezzo)
        });
      }
    });
    await batch.commit();
  }

  pendingRows = [];
  ui.excelFile.value = "";
  ui.importBtn.disabled = true;
  const skippedCount = Math.max(0, totalPending - rowsToCreate.length - rowsToUpdate.length);
  ui.importFeedback.textContent = `Import completato su "${targetCommessaName}": nuovi ${rowsToCreate.length}, aggiornati ${rowsToUpdate.length}, invariati ${skippedCount}.`;
}

async function addManualImpianto(event) {
  event.preventDefault();
  const targetCommessaId = getTargetCommessaId();
  const targetCommessaName = getTargetCommessaName();
  if (!auth.currentUser || !targetCommessaId) {
    ui.manualImpiantoFeedback.textContent = "Seleziona prima una commessa.";
    return;
  }
  if (!canManageData()) {
    ui.manualImpiantoFeedback.textContent = "Solo l'admin può aggiungere impianti.";
    return;
  }

  const denominazione = String(ui.manualImpiantoDenominazione.value || "").trim();
  if (!denominazione) {
    ui.manualImpiantoFeedback.textContent = "Inserisci almeno la denominazione impianto.";
    return;
  }

  const row = {
    distretto: "",
    idSap: "",
    denominazione,
    comune: String(ui.manualImpiantoComune.value || "").trim(),
    indirizzo: String(ui.manualImpiantoIndirizzo.value || "").trim(),
    voceRiferimento: "",
    codicePrezzo: String(ui.manualImpiantoCodice.value || "").trim(),
    sfalci: "",
    frequenzaAnnua: "",
    tipologiaIntervento: "",
    lavorazioniRichieste: "",
    gpsY: null,
    gpsX: null
  };

  const ref = db.collection("commesse").doc(targetCommessaId).collection("impianti");
  const existingSnapshot = await ref.get();
  const existingDoc = existingSnapshot.docs.find((doc) => buildImpiantoKey(doc.data() || {}) === buildImpiantoKey(row));

  if (existingDoc) {
    const existingData = existingDoc.data() || {};
    const mergedCodicePrezzo = mergeMultiValue(existingData.codicePrezzo, row.codicePrezzo);
    await ref.doc(existingDoc.id).update({
      codicePrezzo: mergedCodicePrezzo,
      voceRiferimento: mergeMultiValue(existingData.voceRiferimento, row.voceRiferimento),
      tipologiaIntervento: mergeMultiValue(existingData.tipologiaIntervento, row.tipologiaIntervento),
      lavorazioniRichieste: mergeMultiValue(existingData.lavorazioniRichieste, row.lavorazioniRichieste),
      frequenzaAnnua: mergeMultiValue(existingData.frequenzaAnnua, row.frequenzaAnnua),
      hasOrdinario: hasOrdinario(mergedCodicePrezzo),
      hasStraordinario: hasStraordinario(mergedCodicePrezzo),
      tipoManutenzione: classifyTipoManutenzione(mergedCodicePrezzo)
    });
    ui.manualImpiantoForm.reset();
    ui.manualImpiantoFeedback.textContent = `Impianto già presente in "${targetCommessaName}": codice prezzo aggiornato senza modificare i precedenti.`;
    return;
  }

  await ref.add({
    ...row,
    hasOrdinario: hasOrdinario(row.codicePrezzo),
    hasStraordinario: hasStraordinario(row.codicePrezzo),
    tipoManutenzione: classifyTipoManutenzione(row.codicePrezzo),
    done: false,
    doneAt: null,
    doneBy: "",
    createdAt: firebase.firestore.FieldValue.serverTimestamp()
  });

  ui.manualImpiantoForm.reset();
  ui.manualImpiantoFeedback.textContent = `Impianto aggiunto in "${targetCommessaName}": i precedenti sono stati mantenuti.`;
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

function renderHeaderActivitySummary() {
  if (ui.activeUsersSummary) {
    const activeUsers = platformUsers.filter((user) => {
      const lastSeenMs = firestoreDateToMillis(user.lastSeenAt);
      return lastSeenMs > 0 && (Date.now() - lastSeenMs) <= 10 * 60 * 1000;
    });
    ui.activeUsersSummary.textContent = `Utenti attivi: ${activeUsers.length}`;
  }

  if (ui.lastImpiantoActionSummary) {
    let latestImpiantoAction = null;
    currentImpianti.forEach((impianto) => {
      const doneAtMs = firestoreDateToMillis(impianto.doneAt);
      if (!doneAtMs) return;
      if (!latestImpiantoAction || doneAtMs > latestImpiantoAction.doneAtMs) {
        latestImpiantoAction = {
          doneAtMs,
          doneBy: impianto.doneBy || "Operatore",
          impiantoName: impianto.denominazione || "Impianto"
        };
      }
    });

    if (!latestImpiantoAction) {
      ui.lastImpiantoActionSummary.textContent = "Ultima azione impianto: -";
      return;
    }

    const when = new Date(latestImpiantoAction.doneAtMs).toLocaleString("it-IT", {
      day: "2-digit",
      month: "2-digit",
      year: "numeric",
      hour: "2-digit",
      minute: "2-digit",
      hour12: false
    });
    ui.lastImpiantoActionSummary.textContent = `Ultima azione impianto: ${latestImpiantoAction.doneBy} ha premuto FATTO su ${latestImpiantoAction.impiantoName} (${when})`;
  }
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
        "Data esecuzione": doneInfo.date,
        "Ora esecuzione": doneInfo.time,
        "Eseguito da": impianto.doneBy || "-",
        "Email operatore": auth.currentUser?.email || ""
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
    renderNextActionCard();
    return;
  }

  const filtered = currentImpianti.filter((impianto) => {
    const viewMatch = impiantiViewMode === "done" ? Boolean(impianto.done) : !impianto.done;
    return viewMatch && matchesImpiantoSearch(impianto);
  });
  const sorted = [...filtered].sort((a, b) => {
    if (Boolean(a.done) !== Boolean(b.done)) return a.done ? 1 : -1;
    return distanceFromUser(a) - distanceFromUser(b);
  });

  if (!sorted.length) {
    ui.impiantiLista.innerHTML = "<p class='muted'>Nessun impianto trovato con i filtri correnti.</p>";
    renderNextActionCard();
    return;
  }

  sorted.forEach((impianto) => {
    const article = document.createElement("article");
    article.className = `impianto-item ${impianto.done ? "done" : "todo"}`;
    const impiantoKey = buildImpiantoKey(impianto);
    const detailsVisible = expandedImpiantoKey === impiantoKey;
    article.dataset.impiantoKey = impiantoKey;
    article.classList.toggle("is-expanded", detailsVisible);
    if (highlightedImpiantoKey === impiantoKey) article.classList.add("highlight");

    const distanceKm = distanceFromUser(impianto);
    const distance = formatDistance(distanceKm);
    const travelMeta = estimateTravelMeta(distanceKm);
    const tipo = impianto.tipoManutenzione || classifyTipoManutenzione(impianto.codicePrezzo);
    const hasStraordinariaFlag = impianto.hasStraordinario ?? hasStraordinario(impianto.codicePrezzo);
    const header = document.createElement("button");
    header.type = "button";
    header.className = "impianto-summary-btn";
    header.innerHTML = `
      <strong>${escapeHTML(impianto.denominazione || "(senza nome)")}</strong>
      <span class="badge ${hasStraordinariaFlag ? "badge-straordinaria" : "badge-ordinaria"}">${escapeHTML(tipo)}</span>
      <small class="impianto-travel-meta">
        ${distance} • Traffico <span class="traffic-level traffic-${travelMeta.intensityKey}">${travelMeta.intensityLabel}</span> • ETA ${travelMeta.etaLabel}
      </small>
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

    if (!impianto.done) {
      clearActionUsed(`${selectedCommessaId}:${impiantoKey}:navigate`);
      clearActionUsed(`${selectedCommessaId}:${impiantoKey}:done`);
      clearActionUsed(`${selectedCommessaId}:${impiantoKey}:whatsapp`);
    }

    const actions = document.createElement("div");
    actions.className = "item-actions";
    const managementActions = document.createElement("div");
    managementActions.className = "item-actions item-actions-gestione";
    const isManagementExpanded = expandedImpiantoManagementKeys.has(impiantoKey);

    const addAction = (
      actionKey,
      icon,
      title,
      callback,
      forceDisabled = false,
      trackAsUsed = true,
      targetContainer = actions
    ) => {
      if (isImpiantoActionDenied(actionKey)) return;
      const actionId = `${selectedCommessaId}:${impiantoKey}:${actionKey}`;
      const btn = createActionIconButton(icon, title, async () => {
        await callback();
        registerImpiantoSessionAction(actionKey);
        if (trackAsUsed) markActionAsUsed(actionId);
      });
      btn.dataset.actionKey = actionKey;
      if (impiantoNextActionHighlightEnabled && actionKey === getCurrentImpiantoNextAction()) {
        btn.classList.add("next-action-target");
      }
      if (forceDisabled || (trackAsUsed && isActionUsed(actionId))) setUsedActionButtonState(btn, true);
      targetContainer.appendChild(btn);
    };

    addAction("navigate", "🗺️", "Naviga", () => navigateToImpianto(impianto), false, false, actions);
    addAction(
      "done",
      "✅",
      "Fatto",
      async () => {
        const doneMarked = await markImpiantoDone(impianto);
        if (doneMarked) triggerImpiantoWhatsAppAction(impianto);
      },
      Boolean(impianto.done),
      true,
      actions
    );
    addAction("whatsapp", "✉️", "Invia messaggio", () => triggerImpiantoWhatsAppAction(impianto), false, false, actions);
    addAction("problem-report", "🚨", "Segnala problema", () => openImpiantoReportModal(impianto), false, false, managementActions);
    addAction("gps-update", "📍", "Aggiorna GPS", () => requestGpsUpdate(impianto), false, true, managementActions);
    if (canManageData()) addAction("reset", "♻️", "Reset", () => resetImpianto(impianto), false, false, managementActions);
    if (canManageData()) addAction("edit", "✏️", "Modifica", () => openImpiantoEditor(impianto), false, true, managementActions);
    if (canManageData()) addAction("delete", "🗑️", "Elimina", () => deleteImpianto(impianto), false, true, managementActions);
    if (managementActions.childElementCount > 0) {
      const manageBtn = createButton("⚙️ Gestione", () => {
        if (expandedImpiantoManagementKeys.has(impiantoKey)) expandedImpiantoManagementKeys.delete(impiantoKey);
        else expandedImpiantoManagementKeys.add(impiantoKey);
        renderImpianti();
      });
      manageBtn.classList.add("gestione-toggle-btn");
      manageBtn.setAttribute("aria-expanded", isManagementExpanded ? "true" : "false");
      actions.appendChild(manageBtn);
      managementActions.classList.toggle("hidden", !isManagementExpanded);
    }

    if (actions.childElementCount > 0) article.appendChild(actions);
    if (managementActions.childElementCount > 0) article.appendChild(managementActions);

    ui.impiantiLista.appendChild(article);
  });
  renderNextActionCard();
}

function openImpiantoEditor(impianto) {
  if (!canManageData()) return;
  editingImpiantoIds = getImpiantoDocIds(impianto);
  ui.editDistretto.value = impianto.distretto || "";
  ui.editIdSap.value = impianto.idSap || "";
  ui.editDenominazione.value = impianto.denominazione || "";
  ui.editComune.value = impianto.comune || "";
  ui.editIndirizzo.value = impianto.indirizzo || "";
  ui.editVoceRiferimento.value = impianto.voceRiferimento || "";
  ui.editCodicePrezzo.value = impianto.codicePrezzo || "";
  ui.editFrequenzaAnnua.value = impianto.frequenzaAnnua || "";
  ui.editTipologiaIntervento.value = impianto.tipologiaIntervento || "";
  ui.editLavorazioniRichieste.value = impianto.lavorazioniRichieste || "";
  ui.editSfalci.value = impianto.sfalci || "";
  ui.editGpsY.value = impianto.gpsY == null ? "" : String(impianto.gpsY);
  ui.editGpsX.value = impianto.gpsX == null ? "" : String(impianto.gpsX);
  ui.impiantoEditFeedback.textContent = "";
  ui.impiantoEditModal.classList.remove("hidden");
}

function closeImpiantoEditor() {
  editingImpiantoIds = [];
  ui.impiantoEditForm.reset();
  ui.impiantoEditFeedback.textContent = "";
  ui.impiantoEditModal.classList.add("hidden");
}

async function saveImpiantoEdits(event) {
  event.preventDefault();
  if (!selectedCommessaId || !editingImpiantoIds.length || !canManageData()) return;
  const gpsYValue = String(ui.editGpsY.value || "").trim();
  const gpsXValue = String(ui.editGpsX.value || "").trim();
  const gpsY = gpsYValue ? parseCoordinate(gpsYValue) : null;
  const gpsX = gpsXValue ? parseCoordinate(gpsXValue) : null;
  if ((gpsYValue && gpsY == null) || (gpsXValue && gpsX == null)) {
    ui.impiantoEditFeedback.textContent = "Coordinate GPS non valide. Usa numeri decimali.";
    return;
  }

  const patch = {
    distretto: String(ui.editDistretto.value || "").trim(),
    idSap: String(ui.editIdSap.value || "").trim(),
    denominazione: String(ui.editDenominazione.value || "").trim(),
    comune: String(ui.editComune.value || "").trim(),
    indirizzo: String(ui.editIndirizzo.value || "").trim(),
    voceRiferimento: String(ui.editVoceRiferimento.value || "").trim(),
    codicePrezzo: String(ui.editCodicePrezzo.value || "").trim(),
    frequenzaAnnua: String(ui.editFrequenzaAnnua.value || "").trim(),
    tipologiaIntervento: String(ui.editTipologiaIntervento.value || "").trim(),
    lavorazioniRichieste: String(ui.editLavorazioniRichieste.value || "").trim(),
    sfalci: String(ui.editSfalci.value || "").trim(),
    gpsY,
    gpsX,
    hasOrdinario: hasOrdinario(ui.editCodicePrezzo.value || ""),
    hasStraordinario: hasStraordinario(ui.editCodicePrezzo.value || ""),
    tipoManutenzione: classifyTipoManutenzione(ui.editCodicePrezzo.value || "")
  };

  const ref = db.collection("commesse").doc(selectedCommessaId).collection("impianti");
  trackLocalSheetMutation(selectedCommessaId);
  await Promise.all(editingImpiantoIds.map((id) => ref.doc(id).set(patch, { merge: true })));
  ui.impiantoEditFeedback.textContent = "Modifiche salvate. Sincronizzazione per tutti gli utenti in corso...";
  setTimeout(closeImpiantoEditor, 500);
}

function updateConnectivityStatus() {
  if (!ui.gpsStatus) return;
  if (navigator.onLine) {
    ui.gpsStatus.textContent = "Online: modifiche sincronizzate con il cloud (cache offline attiva).";
  } else {
    ui.gpsStatus.textContent = "Offline: l'app continua in locale, i dati verranno sincronizzati appena torna la rete.";
  }
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

function createActionIconButton(icon, title, onClick) {
  const btn = document.createElement("button");
  btn.type = "button";
  btn.className = "btn action-icon-btn";
  btn.textContent = icon || "";
  btn.title = title;
  btn.setAttribute("aria-label", title);
  btn.addEventListener("click", async () => {
    if (btn.disabled) return;
    await onClick();
  });
  return btn;
}

function setUsedActionButtonState(btn, used) {
  btn.disabled = used;
  btn.classList.toggle("is-used", used);
}

function isActionUsed(actionId) {
  if (!actionId) return false;
  if (usedActionKeys.has(actionId)) return true;
  return localStorage.getItem(`usedAction:${actionId}`) === "1";
}

function markActionAsUsed(actionId) {
  if (!actionId) return;
  usedActionKeys.add(actionId);
  localStorage.setItem(`usedAction:${actionId}`, "1");
  renderImpianti();
}

function clearActionUsed(actionId) {
  if (!actionId) return;
  usedActionKeys.delete(actionId);
  localStorage.removeItem(`usedAction:${actionId}`);
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
  if (!selectedCommessaId || !ids.length) return false;
  if (!canManageData()) {
    if (!currentUserPos) {
      alert("Per segnare FATTO devi attivare la posizione GPS.");
      return false;
    }
    const distanceKm = distanceFromUser(impianto);
    if (!Number.isFinite(distanceKm) || distanceKm > 4) {
      alert("Puoi segnare FATTO solo entro 4 km dall'impianto.");
      return false;
    }
  }
  const exportPayload = {
    commessaId: selectedCommessaId,
    commessaName: selectedCommessaName || "Commessa",
    impianto
  };
  const doneAtLocal = new Date();
  const doneByLocal = auth.currentUser?.displayName || auth.currentUser?.email || "Operatore";
  trackLocalSheetMutation(selectedCommessaId);
  try {
    updateImpiantoLocalState(ids, { done: true, doneAt: doneAtLocal, doneBy: doneByLocal });
    await setImpiantoDone(selectedCommessaId, ids, true);
  } catch (error) {
    console.error("Aggiornamento stato FATTO non completato al primo tentativo:", error);
    retrySetImpiantoDone(selectedCommessaId, ids, true);
  }

  if (!canManageData()) {
    try {
      await queueSheetExportForAdmin(exportPayload);
    } catch (error) {
      console.error("Impianto FATTO ma coda admin non salvata:", error);
    }
    await publishGlobalNotificationEvent("impianto-done", {
      title: "Impianto completato",
      body: `${doneByLocal} ha premuto FATTO su ${impianto.denominazione || "Impianto"} (${selectedCommessaName || "Commessa"}).`,
      commessaId: selectedCommessaId,
      commessaName: selectedCommessaName || "Commessa",
      impiantoName: impianto.denominazione || "Impianto",
      impiantoKey: buildImpiantoKey(impianto)
    });
    return true;
  }

  scheduleCommessaSheetSync(exportPayload.commessaId, exportPayload.commessaName, 200);
  await publishGlobalNotificationEvent("impianto-done", {
    title: "Impianto completato",
    body: `${doneByLocal} ha premuto FATTO su ${impianto.denominazione || "Impianto"} (${selectedCommessaName || "Commessa"}).`,
    commessaId: selectedCommessaId,
    commessaName: selectedCommessaName || "Commessa",
    impiantoName: impianto.denominazione || "Impianto",
    impiantoKey: buildImpiantoKey(impianto)
  });
  return true;
}

async function retrySetImpiantoDone(commessaId, impiantoIds, done, retries = 3) {
  for (let i = 0; i < retries; i += 1) {
    try {
      await setImpiantoDone(commessaId, impiantoIds, done);
      return true;
    } catch (error) {
      console.warn(`Tentativo aggiornamento stato FATTO fallito (${i + 1}/${retries})`, error);
      await new Promise((resolve) => setTimeout(resolve, 1000 * (i + 1)));
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
    alert("Solo un admin può usare reset.");
    return;
  }
  trackLocalSheetMutation(selectedCommessaId);
  updateImpiantoLocalState(ids, { done: false, doneAt: null, doneBy: "" });
  await setImpiantoDone(selectedCommessaId, ids, false);
  const impiantoKey = buildImpiantoKey(impianto);
  clearActionUsed(`${selectedCommessaId}:${impiantoKey}:navigate`);
  clearActionUsed(`${selectedCommessaId}:${impiantoKey}:done`);
  clearActionUsed(`${selectedCommessaId}:${impiantoKey}:whatsapp`);
  clearActionUsed(`${selectedCommessaId}:${impiantoKey}:reset`);
  scheduleCommessaSheetSync(selectedCommessaId, selectedCommessaName, 250);
}

async function deleteImpianto(impianto) {
  const ids = getImpiantoDocIds(impianto);
  if (!selectedCommessaId || !ids.length) return;
  if (!canManageData()) {
    alert("Solo un admin può eliminare impianti.");
    return;
  }

  const ok = window.confirm(`Eliminare impianto ${impianto.denominazione || ""}?`);
  if (!ok) return;

  const ref = db.collection("commesse").doc(selectedCommessaId).collection("impianti");
  trackLocalSheetMutation(selectedCommessaId);
  await Promise.all(ids.map((id) => ref.doc(id).delete()));
  scheduleCommessaSheetSync(selectedCommessaId, selectedCommessaName, 250);
}

async function deleteCommessa(commessaId, nome) {
  if (!canManageData()) {
    alert("Solo un admin può eliminare commesse.");
    return;
  }

  const ok = window.confirm(`Eliminare definitivamente la commessa "${nome}" e tutti i suoi impianti?`);
  if (!ok) return;

  const impiantiRef = db.collection("commesse").doc(commessaId).collection("impianti");
  await deleteCollectionDocs(impiantiRef);
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
    alert("Solo un admin può gestire il personale.");
    return;
  }
  const fullName = ui.personaleNome.value.trim().replace(/\s+/g, " ");
  if (!fullName) return;
  const [cognome, ...nomeParts] = fullName.split(" ");
  const nome = nomeParts.join(" ").trim();
  if (!nome || !cognome) {
    alert("Inserisci Cognome e Nome del personale.");
    return;
  }
  await db.collection("personale").add({
    nome,
    cognome,
    fullName: `${cognome} ${nome}`.trim(),
    createdAt: firebase.firestore.FieldValue.serverTimestamp()
  });
  ui.personaleForm.reset();
}

async function addMezzo(event) {
  event.preventDefault();
  if (!canManageData()) {
    alert("Solo un admin può gestire i mezzi.");
    return;
  }
  const mezzo = {
    nId: String(ui.mezzoNId?.value || "").trim(),
    marca: String(ui.mezzoMarca?.value || "").trim(),
    modello: String(ui.mezzoModello?.value || "").trim(),
    portataCarico: String(ui.mezzoPortataCarico?.value || "").trim(),
    massaComplessivaKg: String(ui.mezzoMassaComplessivaKg?.value || "").trim(),
    alimentazione: String(ui.mezzoAlimentazione?.value || "").trim()
  };
  if (!mezzo.nId) return;

  const existing = findExistingMezzoByNId(mezzo.nId);
  if (existing) {
    await db.collection("mezzi").doc(existing.id).set(buildMezzoPatch(existing, mezzo), { merge: true });
  } else {
    await db.collection("mezzi").add({
      nome: mezzo.nId,
      ...mezzo,
      createdAt: firebase.firestore.FieldValue.serverTimestamp()
    });
  }
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
  const existingByKey = new Map();
  mezziRecords.forEach((item) => {
    const key = normalizeMezzoNId(item.nId || item.nome);
    if (key) existingByKey.set(key, item);
  });
  const importByKey = new Map();
  rows.forEach((row) => {
    const key = normalizeMezzoNId(row.nId);
    if (!key) return;
    const previous = importByKey.get(key) || {};
    importByKey.set(key, mergeMezzoData(previous, row));
  });

  const batch = db.batch();
  importByKey.forEach((mezzo, key) => {
    const existing = existingByKey.get(key);
    if (existing) {
      batch.set(db.collection("mezzi").doc(existing.id), buildMezzoPatch(existing, mezzo), { merge: true });
      return;
    }
    const ref = db.collection("mezzi").doc();
    batch.set(ref, {
      nome: mezzo.nId || "",
      ...mezzo,
      createdAt: firebase.firestore.FieldValue.serverTimestamp()
    }, { merge: true });
  });
  await batch.commit();
  ui.mezziExcelFile.value = "";
  alert(`Import mezzi completato (${importByKey.size} elementi unici).`);
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
      modello: get(["modello", "model"]),
      portataCarico: get(["portatacarico", "portatacaricokg", "portata", "portatakg"]),
      massaComplessivaKg: get(["massacomplessivapesodelcamioncaricokg", "massacomplessivakg", "massa"]),
      alimentazione: get(["alimentazione"])
    };
  }).filter((row) => row.nId);
}

function normalizeMezzoNId(value) {
  return String(value || "").trim().toLowerCase();
}

function mergeMezzoData(base, incoming) {
  const result = { ...base };
  ["nId", "marca", "modello", "portataCarico", "massaComplessivaKg", "alimentazione"].forEach((field) => {
    if (incoming[field]) result[field] = incoming[field];
  });
  return result;
}

function buildMezzoPatch(existing, incoming) {
  const patch = {};
  const merged = mergeMezzoData({
    nId: existing.nId || existing.nome || "",
    marca: existing.marca || "",
    modello: existing.modello || "",
    portataCarico: existing.portataCarico || "",
    massaComplessivaKg: existing.massaComplessivaKg || "",
    alimentazione: existing.alimentazione || ""
  }, incoming);
  patch.nome = merged.nId || existing.nome || "";
  Object.assign(patch, merged);
  return patch;
}

function findExistingMezzoByNId(nId) {
  const normalized = normalizeMezzoNId(nId);
  if (!normalized) return null;
  return mezziRecords.find((item) => normalizeMezzoNId(item.nId || item.nome) === normalized) || null;
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

  const rows = collectionName === "personale"
    ? extractPersonnelNamesFromRawRows(await parseSimpleExcelRows(file, true))
    : await parseSimpleExcelRows(file);
  const uniqueNames = [...new Set(rows.filter(Boolean).map((v) => v.trim()).filter(Boolean))]
    .filter((name) => (collectionName === "personale" ? name.split(/\s+/).length >= 2 : true));
  if (!uniqueNames.length) {
    alert("Il file Excel non contiene nomi validi.");
    return;
  }

  const batch = db.batch();
  uniqueNames.forEach((nomeCompleto) => {
    const ref = db.collection(collectionName).doc();
    const normalized = String(nomeCompleto || "").trim().replace(/\s+/g, " ");
    const [cognome, ...nomeParts] = normalized.split(" ");
    const nome = nomeParts.join(" ").trim();
    batch.set(ref, {
      nome: collectionName === "personale" ? nome : normalized,
      cognome: collectionName === "personale" ? cognome : "",
      fullName: collectionName === "personale" ? `${cognome} ${nome}`.trim() : "",
      createdAt: firebase.firestore.FieldValue.serverTimestamp()
    });
  });
  await batch.commit();
  inputEl.value = "";
  alert(`Import completato (${uniqueNames.length} elementi).`);
}

function extractPersonnelNamesFromRawRows(rawRows) {
  if (!Array.isArray(rawRows) || !rawRows.length) return [];

  const firstNonEmptyRow = rawRows.find((row) => Array.isArray(row) && row.some((cell) => String(cell || "").trim()));
  if (!firstNonEmptyRow) return [];

  const headerKeys = firstNonEmptyRow.map((cell) => normalizeHeaderKey(cell));
  const hasHeader = headerKeys.some((key) => ["nome", "cognome", "nominativo", "nomecognome", "operatore"].includes(key));
  const nomeIndex = headerKeys.findIndex((key) => ["nome", "firstname"].includes(key));
  const cognomeIndex = headerKeys.findIndex((key) => ["cognome", "lastname", "surname"].includes(key));
  const fullNameIndex = headerKeys.findIndex((key) => ["nominativo", "nomecognome", "operatore", "fullName", "fullname"].includes(key));
  const startIndex = hasHeader ? 1 : 0;

  const names = [];
  for (let i = startIndex; i < rawRows.length; i += 1) {
    const row = rawRows[i];
    if (!Array.isArray(row) || !row.length) continue;

    let value = "";
    if (fullNameIndex >= 0) {
      value = String(row[fullNameIndex] || "").trim();
    } else if (nomeIndex >= 0 || cognomeIndex >= 0) {
      const nome = nomeIndex >= 0 ? String(row[nomeIndex] || "").trim() : "";
      const cognome = cognomeIndex >= 0 ? String(row[cognomeIndex] || "").trim() : "";
      value = `${cognome} ${nome}`.trim();
    } else {
      value = String(row[0] || "").trim();
    }

    if (value) names.push(value.replace(/\s+/g, " "));
  }
  return names;
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
    renderHoursOperatoriOptions();
  }, (error) => {
    console.error(error);
    ui.personaleLista.innerHTML = "<p class='muted'>Errore caricamento personale.</p>";
  });
}

function subscribeMezzi() {
  unsubscribeMezzi = db.collection("mezzi").orderBy("createdAt", "asc").onSnapshot((snapshot) => {
    mezziRecords = snapshot.docs.map((doc) => ({ id: doc.id, ...doc.data() }));
    renderMezziList(ui.mezziLista, mezziRecords, deleteMezzo);
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
    const displayLabel = getPersonaleDisplayName(item);
    const label = document.createElement("span");
    label.textContent = displayLabel;
    row.appendChild(label);
    const deleteBtn = createButton("Elimina", () => onDelete(item.id, displayLabel || "elemento"));
    deleteBtn.disabled = !canManageData();
    row.appendChild(deleteBtn);
    container.appendChild(row);
  });
}

function renderMezziList(container, items, onDelete) {
  container.innerHTML = "";
  if (!items.length) {
    container.innerHTML = "<p class='muted'>Nessun elemento.</p>";
    return;
  }
  items.forEach((item) => {
    const row = document.createElement("div");
    row.className = "simple-list-item";
    const title = item.nId || item.nome || "-";
    const portataLabel = item.portataCarico || item.portataCaricoKg || item.portata || "";
    const massaLabel = item.massaComplessivaKg || item.massaComplessiva || item.massa || "";
    const details = [
      item.marca ? `Marca: ${item.marca}` : "",
      item.modello ? `Modello: ${item.modello}` : "",
      portataLabel ? `Portata: ${portataLabel}` : "",
      massaLabel ? `Massa complessiva: ${massaLabel}` : "",
      item.alimentazione ? `Alimentazione: ${item.alimentazione}` : ""
    ].filter(Boolean).join(" • ");
    const label = document.createElement("span");
    label.textContent = details ? `${title} — ${details}` : title;
    row.appendChild(label);
    const deleteBtn = createButton("Elimina", () => onDelete(item.id, title || "elemento"));
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
  const personaleValues = parseMultiEntryValue(rowData.personale);
  const mezziValues = parseMultiEntryValue(rowData.mezzi);
  const row = document.createElement("div");
  row.className = "squadra-row";
  row.innerHTML = `
    <div class="squadra-row-head">
      <strong>Squadra ${index}</strong>
      <button type="button" class="btn remove-squadra-btn">Rimuovi</button>
    </div>
    <div class="squadra-multi-field">
      <div class="squadra-multi-field-head"><strong>👥 Personale</strong></div>
      <div class="squadra-personale-list"></div>
      <button type="button" class="btn btn-small add-personale-input-btn">+ Persona</button>
    </div>
    <div class="squadra-multi-field">
      <div class="squadra-multi-field-head"><strong>🚚 Mezzi</strong></div>
      <div class="squadra-mezzi-list"></div>
      <button type="button" class="btn btn-small add-mezzo-input-btn">+ Mezzo</button>
    </div>
  `;
  row.querySelector(".remove-squadra-btn").addEventListener("click", () => {
    row.remove();
    renumberSquadraRows();
    if (!ui.squadraRows.children.length) addSquadraRow();
  });
  const personaleList = row.querySelector(".squadra-personale-list");
  const mezziList = row.querySelector(".squadra-mezzi-list");
  const addPersonaleBtn = row.querySelector(".add-personale-input-btn");
  const addMezzoBtn = row.querySelector(".add-mezzo-input-btn");

  const addPersonaleInput = (value = "") => addMultiEntryInput({
    container: personaleList,
    listId: "personale-options",
    placeholder: "Personale squadra",
    value,
    sourceValues: personaleRecords.map((p) => getPersonaleDisplayName(p))
  });
  const addMezzoInput = (value = "") => addMultiEntryInput({
    container: mezziList,
    listId: "mezzi-options",
    placeholder: "Mezzo squadra",
    value,
    sourceValues: mezziRecords.map((m) => m.nId || m.nome)
  });

  (personaleValues.length ? personaleValues : [""]).forEach((value) => addPersonaleInput(value));
  (mezziValues.length ? mezziValues : [""]).forEach((value) => addMezzoInput(value));

  addPersonaleBtn.addEventListener("click", () => addPersonaleInput(""));
  addMezzoBtn.addEventListener("click", () => addMezzoInput(""));
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

function parseMultiEntryValue(rawValue) {
  return String(rawValue || "")
    .split(/[;,\n|]+/)
    .map((part) => String(part || "").trim())
    .filter(Boolean);
}

function addMultiEntryInput({ container, listId, placeholder, value, sourceValues }) {
  if (!container) return;
  const wrap = document.createElement("div");
  wrap.className = "squadra-multi-entry-row";
  wrap.innerHTML = `
    <input type="text" class="squadra-multi-entry-input" list="${escapeHTML(listId)}" placeholder="${escapeHTML(placeholder)}" value="${escapeHTML(value || "")}">
    <button type="button" class="btn btn-small remove-squadra-entry-btn" title="Rimuovi elemento">−</button>
  `;
  const input = wrap.querySelector(".squadra-multi-entry-input");
  const removeBtn = wrap.querySelector(".remove-squadra-entry-btn");
  input.addEventListener("blur", () => {
    input.value = resolveSuggestionValue(input.value, sourceValues);
  });
  removeBtn.addEventListener("click", () => {
    wrap.remove();
    if (!container.children.length) {
      addMultiEntryInput({ container, listId, placeholder, value: "", sourceValues });
    }
  });
  container.appendChild(wrap);
}

function renumberSquadraRows() {
  Array.from(ui.squadraRows.children).forEach((row, idx) => {
    const title = row.querySelector(".squadra-row-head strong");
    if (title) title.textContent = `Squadra ${idx + 1}`;
  });
}

function readSquadraRows() {
  return Array.from(ui.squadraRows.querySelectorAll(".squadra-row")).map((row) => ({
    personale: Array.from(row.querySelectorAll(".squadra-personale-list .squadra-multi-entry-input"))
      .map((input) => String(input.value || "").trim())
      .filter(Boolean)
      .join(", "),
    mezzi: Array.from(row.querySelectorAll(".squadra-mezzi-list .squadra-multi-entry-input"))
      .map((input) => String(input.value || "").trim())
      .filter(Boolean)
      .join(", ")
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
  const tomorrow = new Date();
  tomorrow.setDate(tomorrow.getDate() + 1);
  const tomorrowKey = getDateKeyFromLocalDate(tomorrow);
  if (dateKey === tomorrowKey) {
    await publishGlobalNotificationEvent("next-day-squadre-published", {
      title: "Nuove squadre giorno successivo",
      body: `${currentUser?.displayName || currentUser?.email || "Operatore"} ha pubblicato le squadre del ${new Date(`${dateKey}T00:00:00`).toLocaleDateString("it-IT")} per ${payload.commessaNome}.`,
      commessaId,
      commessaName: payload.commessaNome
    });
  }
  ui.squadraCalendarDate.value = dateKey;
}

function getDateKeyFromLocalDate(date) {
  const year = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, "0");
  const day = String(date.getDate()).padStart(2, "0");
  return `${year}-${month}-${day}`;
}

function getAutomaticSquadreDateKey(now = new Date()) {
  const base = new Date(now);
  if (base.getHours() > 17 || (base.getHours() === 17 && base.getMinutes() >= 30)) {
    base.setDate(base.getDate() + 1);
  }
  return getDateKeyFromLocalDate(base);
}

function getActiveSquadreDateKey() {
  return manualSquadreFilterDateKey || getAutomaticSquadreDateKey();
}

function onSquadreFilterDateChange() {
  manualSquadreFilterDateKey = ui.squadreFilterDate?.value || "";
  if (ui.squadraCalendarDate) ui.squadraCalendarDate.value = manualSquadreFilterDateKey;
  renderSquadre();
}

function clearManualSquadreFilterDate() {
  manualSquadreFilterDateKey = "";
  if (ui.squadreFilterDate) ui.squadreFilterDate.value = "";
  if (ui.squadraCalendarDate) ui.squadraCalendarDate.value = "";
  renderSquadre();
}

function renderSquadre() {
  ui.squadreLista.innerHTML = "";
  const selectedDateKey = getActiveSquadreDateKey();
  const storicoDelGiorno = squadreHistoryByDate.get(selectedDateKey) || new Map();
  if (ui.squadreFilterStatus) {
    const selectedDateLabel = new Date(`${selectedDateKey}T00:00:00`).toLocaleDateString("it-IT");
    ui.squadreFilterStatus.textContent = manualSquadreFilterDateKey
      ? `Filtro admin attivo: mostro le squadre del ${selectedDateLabel}.`
      : `Filtro automatico attivo: mostro le squadre del ${selectedDateLabel} (dopo le 17:30 passa al giorno successivo).`;
  }

  const commesse = Array.from(commesseById.values());
  const commesseConSquadre = commesse.filter((commessa) => {
    const squad = storicoDelGiorno.get(commessa.id) || {};
    const rows = Array.isArray(squad.squadre) ? squad.squadre : getLegacySquadreRows(squad);
    return rows.some((row) => row.personale || row.mezzi);
  });
  if (!commesseConSquadre.length) {
    ui.squadreLista.innerHTML = "<p class='muted'>Nessuna composizione trovata per la data selezionata.</p>";
    return;
  }

  commesseConSquadre.forEach((commessa) => {
    const item = document.createElement("article");
    item.className = "squadra-item";
    const squad = storicoDelGiorno.get(commessa.id) || {};
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
  ui.squadraHint.textContent = "Usa “+ Persona” e “+ Mezzo”: il nuovo campo resta sulla stessa riga del precedente finché c'è spazio.";
}

function updateSuggestionLists() {
  ui.personaleOptions.innerHTML = "";
  personaleRecords.forEach((person) => {
    const option = document.createElement("option");
    option.value = getPersonaleDisplayName(person);
    ui.personaleOptions.appendChild(option);
  });
  ui.mezziOptions.innerHTML = "";
  mezziRecords.forEach((mezzo) => {
    const option = document.createElement("option");
    option.value = mezzo.nId || mezzo.nome || "";
    ui.mezziOptions.appendChild(option);
  });
}

function getPersonaleDisplayName(person) {
  if (!person) return "";
  const cognome = String(person.cognome || "").trim();
  const nome = String(person.nome || "").trim();
  const composed = `${cognome} ${nome}`.trim();
  return composed || String(person.fullName || "").trim();
}

async function setImpiantoDone(commessaId, impiantoIds, done) {
  const user = auth.currentUser;
  if (!user) return;
  const doneAt = done ? firebase.firestore.Timestamp.fromDate(new Date()) : null;

  if (!commessaId) throw new Error("Commessa non selezionata per aggiornamento stato impianto.");
  const ref = db.collection("commesse").doc(commessaId).collection("impianti");
  await Promise.all(impiantoIds.map((impiantoId) => ref.doc(impiantoId).update({
    done,
    doneAt,
    doneBy: done ? (user.displayName || user.email || "Operatore") : ""
  })));
}

function canTriggerImpiantoWhatsApp(impianto, notify = true) {
  if (canManageData()) return true;
  if (!currentUserPos) {
    if (notify) alert("Per inviare WhatsApp devi attivare la posizione GPS.");
    return false;
  }
  const distanceKm = distanceFromUser(impianto);
  if (!Number.isFinite(distanceKm) || distanceKm > 4) {
    if (notify) alert("Puoi inviare WhatsApp solo entro 4 km dall'impianto.");
    return false;
  }
  return true;
}

function triggerImpiantoWhatsAppAction(impianto) {
  if (!canTriggerImpiantoWhatsApp(impianto, true)) return false;
  openWhatsApp(impianto);
  return true;
}

function openWhatsApp(impianto, options = {}) {
  const user = auth.currentUser;
  if (!user) {
    alert("Devi fare login.");
    return;
  }

  const isOnlyOrdinaria = hasOrdinario(impianto.codicePrezzo) && !hasStraordinario(impianto.codicePrezzo);
  const title = isOnlyOrdinaria
    ? "✅ MANUTENZIONE ORDINARIA ESEGUITA"
    : "✅ MANUTENZIONE ORDINARIA + STRAORDINARIA ESEGUITA";
  const doneInfo = formatDoneDateTime(impianto.doneAt);
  const date = doneInfo.date === "-" ? new Date().toLocaleDateString("it-IT") : doneInfo.date;
  const time = doneInfo.time === "-" ? new Date().toLocaleTimeString("it-IT", { hour: "2-digit", minute: "2-digit", hour12: false }) : doneInfo.time;
  const message = [
    `${title} - Report operativo`,
    `🏗️ Impianto: ${impianto.denominazione || "-"}`,
    `📍 Comune: ${impianto.comune || "-"}`,
    `🛣️ Via: ${impianto.indirizzo || "-"}`,
    `🆔 ID SAP: ${impianto.idSap || "-"}`,
    ...(isOnlyOrdinaria ? [] : [`🛠️ Lavorazione straordinaria: ${impianto.lavorazioniRichieste || impianto.tipologiaIntervento || "-"}`]),
    `👷 Operatore: ${user.displayName || user.email || "-"}`,
    `📅 Data: ${date}`,
    `🕒 Ora: ${time}`
  ].join("\n");

  const encodedMessage = encodeURIComponent(message);
  const appUrl = `whatsapp://send?text=${encodedMessage}`;
  const useContactPicker = options?.preferContactPicker !== false;
  const webUrl = useContactPicker
    ? `https://wa.me/?text=${encodedMessage}`
    : `https://api.whatsapp.com/send?text=${encodedMessage}`;
  const targetWindow = options?.targetWindow;
  const disableWebFallback = Boolean(options?.disableWebFallback);
  if (targetWindow && !targetWindow.closed) {
    targetWindow.location.replace(appUrl);
    if (!disableWebFallback) {
      setTimeout(() => {
        if (!targetWindow.closed) targetWindow.location.replace(webUrl);
      }, 700);
    }
    return;
  }
  const target = options?.target || "_blank";
  const popup = window.open(appUrl, target, "noopener");
  if (!popup) {
    window.location.href = disableWebFallback ? appUrl : webUrl;
    return;
  }
  if (target === "_self") {
    if (!disableWebFallback) {
      setTimeout(() => {
        window.location.href = webUrl;
      }, 700);
    }
    return;
  }
  if (!disableWebFallback) {
    setTimeout(() => {
      if (!popup.closed) popup.location.replace(webUrl);
    }, 700);
  }
}

function openImpiantoReportModal(impianto) {
  reportingImpianto = impianto || null;
  ui.impiantoReportForm.reset();
  ui.impiantoReportFeedback.textContent = "";
  ui.impiantoReportModal.classList.remove("hidden");
}

function closeImpiantoReportModal() {
  reportingImpianto = null;
  ui.impiantoReportForm.reset();
  ui.impiantoReportFeedback.textContent = "";
  ui.impiantoReportModal.classList.add("hidden");
}

async function submitImpiantoReport(event) {
  event.preventDefault();
  if (!reportingImpianto) {
    ui.impiantoReportFeedback.textContent = "Impianto non disponibile per la segnalazione.";
    return;
  }
  const user = auth.currentUser;
  if (!user) {
    ui.impiantoReportFeedback.textContent = "Devi fare login prima di inviare una segnalazione.";
    return;
  }
  const titolo = String(ui.impiantoReportTitle.value || "").trim();
  const testo = String(ui.impiantoReportText.value || "").trim();
  if (!titolo || !testo) {
    ui.impiantoReportFeedback.textContent = "Compila titolo e testo della segnalazione.";
    return;
  }
  const now = new Date();
  const message = [
    "⚠️ SEGNALAZIONE PROBLEMA IMPIANTO - Report operativo",
    `🏗️ Impianto: ${reportingImpianto.denominazione || "-"}`,
    `📍 Comune: ${reportingImpianto.comune || "-"}`,
    `🛣️ Via: ${reportingImpianto.indirizzo || "-"}`,
    `🆔 ID SAP: ${reportingImpianto.idSap || "-"}`,
    `📝 Oggetto segnalazione: ${titolo}`,
    `📋 Dettaglio problema segnalato: ${testo}`,
    `👷 Operatore segnalante: ${user.displayName || user.email || "-"}`,
    `📅 Data segnalazione: ${now.toLocaleDateString("it-IT")}`,
    `🕒 Ora segnalazione: ${now.toLocaleTimeString("it-IT", { hour: "2-digit", minute: "2-digit", hour12: false })}`,
    "✅ Conferma: stiamo segnalando al cliente il problema riscontrato e il relativo intervento richiesto."
  ].join("\n");
  window.open(`https://wa.me/?text=${encodeURIComponent(message)}`, "_blank");
  ui.impiantoReportFeedback.textContent = "WhatsApp aperto con la segnalazione pronta da inviare.";
  setTimeout(closeImpiantoReportModal, 200);
}

function getCurrentPositionOnce() {
  return new Promise((resolve, reject) => {
    if (!navigator.geolocation) {
      reject(new Error("Geolocalizzazione non supportata."));
      return;
    }
    navigator.geolocation.getCurrentPosition(
      (pos) => {
        resolve({
          lat: Number(pos.coords.latitude),
          lng: Number(pos.coords.longitude)
        });
      },
      (error) => reject(error),
      { enableHighAccuracy: true, timeout: 15000, maximumAge: 0 }
    );
  });
}

async function requestGpsUpdate(impianto) {
  if (!currentUser || !selectedCommessaId) {
    alert("Seleziona una commessa ed esegui il login.");
    return;
  }
  const confirmed = window.confirm("Vuoi aggiornare la posizione di questo impianto? Verrà inviata richiesta WhatsApp all'amministratore.");
  if (!confirmed) return;

  try {
    const pos = await getCurrentPositionOnce();
    const impiantoIds = getImpiantoDocIds(impianto);
    const requestRef = await db.collection("gpsUpdateRequests").add({
      commessaId: selectedCommessaId,
      commessaName: selectedCommessaName || "",
      impiantoKey: buildImpiantoKey(impianto),
      impiantoIds,
      impiantoDenominazione: impianto.denominazione || "",
      impiantoIdSap: impianto.idSap || "",
      impiantoComune: impianto.comune || "",
      impiantoIndirizzo: impianto.indirizzo || "",
      operatorId: currentUser.uid,
      operatorName: currentUser.displayName || currentUser.email || "Operatore",
      operatorEmail: currentUser.email || "",
      operatorLat: pos.lat,
      operatorLng: pos.lng,
      status: "pending",
      createdAt: firebase.firestore.FieldValue.serverTimestamp()
    });

    const mapsUrl = `https://maps.google.com/?q=${pos.lat},${pos.lng}`;
    const waText = [
      "📍 Richiesta aggiornamento GPS impianto",
      `ID richiesta: ${requestRef.id}`,
      `Commessa: ${selectedCommessaName || "-"}`,
      `Impianto: ${impianto.denominazione || "-"}`,
      `ID SAP: ${impianto.idSap || "-"}`,
      `Operatore: ${currentUser.displayName || currentUser.email || "-"}`,
      `Coordinate operatore: ${pos.lat}, ${pos.lng}`,
      `Mappa: ${mapsUrl}`
    ].join("\n");
    window.open(`https://wa.me/${GPS_APPROVAL_PHONE}?text=${encodeURIComponent(waText)}`, "_blank");

    await notifyAdminsForGpsRequest(requestRef.id, impianto, pos);
    alert("Richiesta inviata. In attesa approvazione admin.");
  } catch (error) {
    console.error("Errore richiesta aggiornamento GPS:", error);
    alert("Impossibile inviare la richiesta GPS.");
  }
}

async function notifyAdminsForGpsRequest(requestId, impianto, pos) {
  const adminUsers = platformUsers.filter((user) => adminEmails.has(normalizeEmail(user.email)));
  if (!adminUsers.length) return;
  const text = `📍 Richiesta GPS ${requestId} per ${impianto.denominazione || "impianto"} (${pos.lat.toFixed(6)}, ${pos.lng.toFixed(6)}). Apri Gestione > Utenti per accettare/rifiutare.`;
  await Promise.all(adminUsers.map((adminUser) => db.collection("chatMessages").add({
    type: "text",
    text,
    recipientId: adminUser.id,
    senderId: currentUser.uid,
    senderName: currentUser.displayName || currentUser.email || "Operatore",
    senderEmail: currentUser.email || "",
    createdAt: firebase.firestore.FieldValue.serverTimestamp()
  })));
}

function openSquadraWhatsApp(squad, commessa) {
  const squadRows = Array.isArray(squad.squadre) ? squad.squadre : getLegacySquadreRows(squad);
  const rowsMessage = squadRows.map((row, idx) => ([
    `👥 SQUADRA ${idx + 1}`,
    `   • Personale: ${row.personale || "-"}`,
    `   • Mezzi: ${row.mezzi || "-"}`,
    "────────────────────"
  ].join("\n"))).join("\n");
  const message = [
    "📣 Richiesta di conferma composizione squadre",
    "Gentile tecnico, di seguito la composizione registrata.",
    "────────────────────",
    `📁 Commessa: ${commessa.nome || "-"}`,
    `📅 Giorno riferimento: ${squad.riferimentoData || "-"}`,
    "────────────────────",
    rowsMessage || "Nessuna squadra compilata.\n────────────────────",
    "Grazie per la verifica."
  ].join("\n");

  const url = `https://wa.me/?text=${encodeURIComponent(message)}`;
  window.open(url, "_blank");
}

function getSquadrePackageEntries() {
  const selectedDateKey = getActiveSquadreDateKey();
  const storicoDelGiorno = squadreHistoryByDate.get(selectedDateKey) || new Map();
  const commesse = Array.from(commesseById.values());
  return commesse.map((commessa) => {
    const squad = storicoDelGiorno.get(commessa.id) || {};
    const squadRows = Array.isArray(squad.squadre) ? squad.squadre : getLegacySquadreRows(squad);
    const hasRows = squadRows.some((row) => row.personale || row.mezzi);
    return {
      commessa,
      squad,
      squadRows,
      hasRows
    };
  }).filter((entry) => entry.hasRows);
}

async function buildSquadrePackagePdfBlob(entries) {
  if (!window.jspdf || !window.jspdf.jsPDF) throw new Error("Libreria PDF non disponibile.");
  const { jsPDF } = window.jspdf;
  const doc = new jsPDF({ orientation: "portrait", unit: "mm", format: "a4" });
  const pageWidth = 210;
  const pageHeight = 297;
  const margin = 12;
  const contentWidth = pageWidth - (margin * 2);
  const maxY = pageHeight - margin;

  const drawHeader = (entry, idx) => {
    doc.setFillColor(99, 102, 241);
    doc.roundedRect(margin, margin, contentWidth, 24, 4, 4, "F");
    doc.setTextColor(255, 255, 255);
    doc.setFont("helvetica", "bold");
    doc.setFontSize(14);
    doc.text(`Squadre per commessa • ${idx + 1}/${entries.length}`, margin + 6, margin + 9);
    doc.setFontSize(11);
    doc.text(entry.commessa.nome || "Commessa senza nome", margin + 6, margin + 16);
    doc.setFont("helvetica", "normal");
    doc.setFontSize(9.5);
    doc.text(`Giorno: ${entry.squad.riferimentoData || "-"}`, margin + 6, margin + 21);
    doc.text(`Export: ${new Date().toLocaleString("it-IT")}`, pageWidth - margin - 46, margin + 21);
  };

  const drawSquadraCard = (row, rowIdx, yStart) => {
    let y = yStart;
    doc.setFillColor(251, 253, 255);
    doc.setDrawColor(226, 232, 240);
    doc.roundedRect(margin, y, contentWidth, 34, 4, 4, "FD");

    doc.setFillColor(236, 253, 243);
    doc.setDrawColor(187, 247, 208);
    doc.roundedRect(margin + 4, y + 4, 30, 7, 3, 3, "FD");
    doc.setTextColor(22, 101, 52);
    doc.setFont("helvetica", "bold");
    doc.setFontSize(9.5);
    doc.text(`Squadra ${rowIdx + 1}`, margin + 7, y + 8.8);

    const personaleLabel = "👥 Personale:";
    const mezziLabel = "🚚 Mezzi:";
    const personnelLines = doc.splitTextToSize(String(row.personale || "-"), contentWidth - 44);
    const mezziLines = doc.splitTextToSize(String(row.mezzi || "-"), contentWidth - 44);

    doc.setTextColor(17, 24, 39);
    doc.setFont("helvetica", "bold");
    doc.setFontSize(9.5);
    doc.text(personaleLabel, margin + 4, y + 16);
    doc.text(mezziLabel, margin + 4, y + 25);

    doc.setFont("helvetica", "normal");
    doc.setFontSize(9.3);
    doc.text(personnelLines.slice(0, 2), margin + 34, y + 16);
    doc.text(mezziLines.slice(0, 2), margin + 34, y + 25);

    const rowsUsed = Math.max(personnelLines.length, mezziLines.length, 1);
    return y + Math.max(34, 24 + (rowsUsed * 4.3));
  };

  entries.forEach((entry, idx) => {
    if (idx > 0) doc.addPage();
    drawHeader(entry, idx);
    let y = margin + 30;

    if (!entry.squadRows.length) {
      doc.setTextColor(75, 85, 99);
      doc.setFontSize(10);
      doc.text("Nessuna squadra compilata per questa commessa.", margin, y + 8);
      return;
    }

    entry.squadRows.forEach((row, rowIdx) => {
      if (y > maxY - 40) {
        doc.addPage();
        drawHeader(entry, idx);
        y = margin + 30;
      }
      y = drawSquadraCard(row, rowIdx, y + 2) + 4;
    });
  });
  return doc.output("blob");
}

async function shareAllSquadreToWhatsApp() {
  const entries = getSquadrePackageEntries();
  if (!entries.length) {
    alert("Nessuna composizione squadre disponibile da inviare.");
    return;
  }

  const sortedEntries = [...entries].sort((a, b) => String(a.commessa.nome || "").localeCompare(String(b.commessa.nome || ""), "it"));
  const groupedLines = sortedEntries.map((entry, entryIdx) => {
    const dateLabel = entry.squad.riferimentoData
      ? new Date(`${entry.squad.riferimentoData}T00:00:00`).toLocaleDateString("it-IT")
      : "-";
    const squadLines = entry.squadRows.map((row, rowIdx) => ([
      `   👥 SQUADRA ${rowIdx + 1}`,
      `      • Personale: ${row.personale || "-"}`,
      `      • Mezzi: ${row.mezzi || "-"}`,
      "   ────────────────"
    ].join("\n"))).join("\n");
    return [
      "════════════════════",
      `📁 COMMESSA ${entryIdx + 1}: ${String(entry.commessa.nome || "Commessa").toUpperCase()}`,
      `📅 Giorno programmato: ${dateLabel}`,
      squadLines || "   - Nessuna squadra assegnata -"
    ].join("\n");
  });

  const message = [
    "📣 PROPOSTA SQUADRE OPERATIVE",
    "Buongiorno, condivido la proposta squadre per la pianificazione operativa.",
    "",
    groupedLines.join("\n\n"),
    "",
    "✅ Per favore confermare o segnalare eventuali modifiche."
  ].join("\n");
  window.open(`https://wa.me/?text=${encodeURIComponent(message)}`, "_blank");
}

function getCommessaAccentColor(commessaId, index) {
  const palette = ["#2563eb", "#7c3aed", "#0f766e", "#d97706", "#db2777", "#0891b2", "#4f46e5", "#ca8a04"];
  const source = String(commessaId || index || "");
  let hash = 0;
  for (let i = 0; i < source.length; i += 1) {
    hash = ((hash << 5) - hash) + source.charCodeAt(i);
    hash |= 0;
  }
  return palette[Math.abs(hash) % palette.length];
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
    const popupHtml = buildImpiantoMapPopup(impianto, tipo);
    marker.bindPopup(popupHtml);
    marker.on("click", () => focusImpiantoInList(impianto, false));
    marker.addTo(markerLayer);
    bounds.push([impianto.gpsY, impianto.gpsX]);
  });

  if (currentUserPos) {
    L.marker([currentUserPos.lat, currentUserPos.lng], {
      icon: L.divIcon({
        className: "",
        html: "<div class='marker-operator'>👷</div>",
        iconSize: [24, 24],
        iconAnchor: [12, 12]
      })
    }).addTo(markerLayer).bindPopup("Operatore al lavoro");
    bounds.push([currentUserPos.lat, currentUserPos.lng]);
  }

  if (bounds.length > 0) {
    map.fitBounds(bounds, { padding: [24, 24], maxZoom: 14 });
  }
}

function buildImpiantoMapPopup(impianto, tipo) {
  const impiantoKey = buildImpiantoKey(impianto);
  return `
    <div class="map-popup-card" data-impianto-key="${escapeHTML(impiantoKey)}">
      <p><b>${escapeHTML(impianto.denominazione || "Impianto")}</b></p>
      <p>Comune: ${escapeHTML(impianto.comune || "-")}</p>
      <p>Indirizzo: ${escapeHTML(impianto.indirizzo || "-")}</p>
      <p>Tipo: ${escapeHTML(tipo)}</p>
      <p>Stato: ${impianto.done ? "Fatto" : "Da fare"}</p>
      <div class="map-popup-actions">
        <button type="button" class="btn btn-small" data-map-popup-action="navigate" data-impianto-key="${escapeHTML(impiantoKey)}">Naviga</button>
        <button type="button" class="btn btn-small" data-map-popup-action="close">Chiudi</button>
      </div>
    </div>
  `;
}

map.on("popupopen", (event) => {
  const popupElement = event.popup?.getElement();
  if (!popupElement) return;
  popupElement.querySelectorAll("[data-map-popup-action='navigate']").forEach((button) => {
    button.addEventListener("click", async () => {
      const key = button.getAttribute("data-impianto-key");
      const impianto = currentImpianti.find((item) => buildImpiantoKey(item) === key);
      if (!impianto) return;
      await navigateToImpianto(impianto);
    });
  });
  popupElement.querySelectorAll("[data-map-popup-action='close']").forEach((button) => {
    button.addEventListener("click", () => {
      map.closePopup(event.popup);
      renderMap();
    });
  });
});

function focusImpiantoInList(impianto, scroll = true) {
  const key = buildImpiantoKey(impianto);
  highlightedImpiantoKey = key;
  expandedImpiantoKey = key;
  renderImpianti();
  if (!scroll) return;
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
  return adminEmails.has(email);
}

function normalizeEmail(email) {
  return String(email || "").trim().toLowerCase();
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
  const portataLabel = selectedFuelMezzo.portataCarico || selectedFuelMezzo.portataCaricoKg || selectedFuelMezzo.portata || "-";
  const massaLabel = selectedFuelMezzo.massaComplessivaKg || selectedFuelMezzo.massaComplessiva || selectedFuelMezzo.massa || "-";
  ui.fuelMezzoDetails.innerHTML = `
    <p><b>N. ID:</b> ${escapeHTML(selectedFuelMezzo.nId || selectedFuelMezzo.nome || "-")}</p>
    <p><b>Marca:</b> ${escapeHTML(selectedFuelMezzo.marca || "-")}</p>
    <p><b>Modello:</b> ${escapeHTML(selectedFuelMezzo.modello || "-")}</p>
    <p><b>Portata (carico):</b> ${escapeHTML(portataLabel)}</p>
    <p><b>Massa complessiva (kg):</b> ${escapeHTML(massaLabel)}</p>
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
      if (!lat || !lon) return null;
      const brandLabel = detectFuelBrand(item.tags || {});
      if (!brandLabel) return null;
      return {
        id: item.id,
        name: item.tags.name || item.tags.brand || "Distributore",
        brand: item.tags.brand || item.tags.operator || brandLabel,
        brandLabel,
        lat,
        lon,
        distance: haversine(currentUserPos.lat, currentUserPos.lng, lat, lon)
      };
    }).filter(Boolean).sort((a, b) => a.distance - b.distance);
    renderFuelStations(stations);
  } catch (error) {
    console.error("Errore caricamento distributori:", error);
    const retryBtn = createButton("Riprova", () => loadNearbyFuelStations());
    ui.fuelStationsList.innerHTML = "<p class='muted'>Errore caricamento distributori. Riprova tra pochi secondi.</p>";
    ui.fuelStationsList.appendChild(retryBtn);
    if (fuelStationsLayer) fuelStationsLayer.clearLayers();
  }
}

function detectFuelBrand(tags) {
  const brandText = [
    tags.brand,
    tags.name,
    tags.operator,
    tags["brand:it"]
  ].filter(Boolean).join(" ").toLowerCase();
  if (brandText.includes("q8")) return "Q8";
  if (brandText.includes("eni") || brandText.includes("agip")) return "ENI";
  return "";
}

async function fetchFuelStationsFromOverpass(lat, lng) {
  const query = `
    [out:json][timeout:25];
    (
      node[\"amenity\"=\"fuel\"](around:12000,${lat},${lng});
      way[\"amenity\"=\"fuel\"](around:12000,${lat},${lng});
    );
    out center;
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
    const marker = L.marker([station.lat, station.lon], {
      icon: createFuelMarkerIcon(station.brandLabel)
    }).addTo(fuelStationsLayer);
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

function createFuelMarkerIcon(brandLabel) {
  const markerClass = getFuelMarkerClass(brandLabel);
  return L.divIcon({
    className: "fuel-marker-wrap",
    html: `<span class="fuel-marker-label ${markerClass}">${escapeHTML(brandLabel || "FUEL")}</span>`,
    iconSize: [44, 24],
    iconAnchor: [22, 12],
    popupAnchor: [0, -10]
  });
}

function getFuelMarkerClass(brandLabel) {
  const normalized = String(brandLabel || "").toLowerCase();
  if (normalized.includes("q8")) return "fuel-marker-q8";
  if (normalized.includes("eni")) return "fuel-marker-eni";
  return "fuel-marker-default";
}

function onPersonalServiceCategoryClick(event) {
  const btn = event.target.closest(".personal-service-category-btn");
  if (!btn) return;
  const category = btn.dataset.serviceCategory || "";
  if (!category) return;
  window.location.hash = `servizi-personali=${category}`;
  applyRoute();
}

async function loadPersonalServicesByCategory(category) {
  if (!PERSONAL_SERVICE_CATEGORIES[category]) return;
  activePersonalServiceCategory = category;
  expandedPersonalServiceId = "";
  personalServicesResults = [];
  ui.personalServicesCategories?.querySelectorAll(".personal-service-category-btn").forEach((btn) => {
    btn.classList.toggle("is-active", btn.dataset.serviceCategory === category);
  });
  const cfg = PERSONAL_SERVICE_CATEGORIES[category];
  const radiusMeters = getSelectedPersonalServicesRadius();
  ui.personalServicesPageTitle.textContent = `${cfg.icon} ${cfg.title}`;
  ui.personalServicesListTitle.textContent = `Più vicini a te • ${cfg.title} • raggio ${Math.round(radiusMeters / 1000)} km`;
  if (!currentUserPos) {
    ui.personalServicesFeedback.textContent = "Posizione non disponibile. Attiva GPS per usare i servizi personali.";
    ui.personalServicesList.innerHTML = "";
    clearPersonalServicesMap();
    return;
  }
  ui.personalServicesFeedback.textContent = "Caricamento luoghi in corso...";
  ui.personalServicesList.innerHTML = "";
  try {
    const data = await fetchPersonalServicesFromOverpass(category, currentUserPos.lat, currentUserPos.lng, radiusMeters);
    const places = normalizePersonalServices(data.elements || [], category);
    personalServicesResults = places;
    renderPersonalServicesList();
    renderPersonalServicesMap();
    if (!places.length) {
      ui.personalServicesFeedback.textContent = "Nessun risultato trovato nella zona.";
    } else if (category === "lunch") {
      const acceptedCount = places.filter((place) => isMealVoucherAccepted(place.tags)).length;
      ui.personalServicesFeedback.textContent = `Trovati ${places.length} luoghi (${acceptedCount} con buoni pasto).`;
    } else {
      ui.personalServicesFeedback.textContent = `Trovati ${places.length} luoghi vicino a te.`;
    }
  } catch (error) {
    console.error("Errore caricamento servizi personali:", error);
    ui.personalServicesFeedback.textContent = "Errore durante il caricamento. Riprova.";
    ui.personalServicesList.innerHTML = "";
    clearPersonalServicesMap();
  }
}

function normalizePersonalServices(items, category) {
  const seen = new Set();
  return items.map((item) => {
    const lat = item.lat || (item.center && item.center.lat);
    const lon = item.lon || (item.center && item.center.lon);
    if (!lat || !lon) return null;
    const tags = item.tags || {};
    const name = tags.name || tags.brand || defaultPersonalServiceName(category);
    const key = `${name.toLowerCase()}-${Math.round(lat * 10000)}-${Math.round(lon * 10000)}`;
    if (seen.has(key)) return null;
    seen.add(key);
    return {
      id: item.id || key,
      category,
      name,
      lat,
      lon,
      tags,
      distance: haversine(currentUserPos.lat, currentUserPos.lng, lat, lon)
    };
  }).filter(Boolean).sort((a, b) => a.distance - b.distance);
}

function defaultPersonalServiceName(category) {
  const cfg = PERSONAL_SERVICE_CATEGORIES[category];
  return cfg ? cfg.title : "Servizio";
}

function ensurePersonalServicesMap() {
  if (personalServicesMapInstance) return;
  personalServicesMapInstance = L.map("personal-services-map");
  L.tileLayer("https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png", {
    maxZoom: 19,
    attribution: "&copy; OpenStreetMap"
  }).addTo(personalServicesMapInstance);
  personalServicesLayer = L.layerGroup().addTo(personalServicesMapInstance);
}

function clearPersonalServicesMap() {
  if (!personalServicesLayer) return;
  personalServicesLayer.clearLayers();
}

function renderPersonalServicesMap() {
  ensurePersonalServicesMap();
  clearPersonalServicesMap();
  if (!personalServicesResults.length) return;
  const bounds = [];
  personalServicesResults.forEach((place) => {
    const marker = L.marker([place.lat, place.lon], {
      icon: createPersonalServiceMarkerIcon(place.category)
    }).addTo(personalServicesLayer);
    marker.bindPopup(`<b>${escapeHTML(place.name)}</b><br>${formatDistance(place.distance)}`);
    marker.on("click", () => selectPersonalService(place.id));
    bounds.push([place.lat, place.lon]);
  });
  if (currentUserPos) bounds.push([currentUserPos.lat, currentUserPos.lng]);
  personalServicesMapInstance.fitBounds(bounds, { padding: [24, 24] });
}

function createPersonalServiceMarkerIcon(category) {
  const cfg = PERSONAL_SERVICE_CATEGORIES[category] || {};
  return L.divIcon({
    className: "fuel-marker-wrap",
    html: `<span class="fuel-marker-label ${getPersonalServiceMarkerClass(category)}">${escapeHTML(cfg.icon || "📍")}</span>`,
    iconSize: [44, 24],
    iconAnchor: [22, 12],
    popupAnchor: [0, -10]
  });
}

function getPersonalServiceMarkerClass(category) {
  const palette = {
    breakfast: "ps-marker-breakfast",
    lunch: "ps-marker-lunch",
    supermarket: "ps-marker-supermarket",
    tobacco: "ps-marker-tobacco",
    wc: "ps-marker-wc",
    atm: "ps-marker-atm",
    pharmacy: "ps-marker-pharmacy",
    parking: "ps-marker-parking"
  };
  return palette[category] || "ps-marker-default";
}

function renderPersonalServicesList() {
  ui.personalServicesList.innerHTML = "";
  if (!personalServicesResults.length) return;
  if (activePersonalServiceCategory === "lunch") {
    renderLunchGroupedList();
    return;
  }
  personalServicesResults.forEach((place) => {
    ui.personalServicesList.appendChild(buildPersonalServiceRow(place));
  });
}

function renderLunchGroupedList() {
  const accepted = personalServicesResults.filter((place) => isMealVoucherAccepted(place.tags));
  const other = personalServicesResults.filter((place) => !isMealVoucherAccepted(place.tags));
  const acceptedCard = document.createElement("div");
  acceptedCard.className = "simple-list-item stacked";
  acceptedCard.innerHTML = `<strong>✅ Accettano buoni pasto (${accepted.length})</strong>`;
  const acceptedList = document.createElement("div");
  acceptedList.className = "simple-list";
  accepted.forEach((place) => acceptedList.appendChild(buildPersonalServiceRow(place)));
  acceptedCard.appendChild(acceptedList);
  ui.personalServicesList.appendChild(acceptedCard);

  const otherCard = document.createElement("div");
  otherCard.className = "simple-list-item stacked";
  otherCard.innerHTML = `<strong>ℹ️ Non accettano o non indicato (${other.length})</strong>`;
  const otherList = document.createElement("div");
  otherList.className = "simple-list";
  other.forEach((place) => otherList.appendChild(buildPersonalServiceRow(place)));
  otherCard.appendChild(otherList);
  ui.personalServicesList.appendChild(otherCard);
}

function buildPersonalServiceRow(place) {
  const row = document.createElement("div");
  row.className = "simple-list-item stacked";
  row.dataset.placeId = String(place.id);
  const head = document.createElement("div");
  head.className = "personal-service-row-head";
  const iconBtn = createButton(PERSONAL_SERVICE_CATEGORIES[place.category]?.icon || "📍", () => selectPersonalService(place.id));
  iconBtn.classList.add("action-icon-btn");
  const nameBtn = createButton(place.name, () => selectPersonalService(place.id));
  nameBtn.classList.add("personal-service-name-btn");
  const meta = document.createElement("small");
  meta.className = "muted";
  meta.textContent = formatDistance(place.distance);
  const textWrap = document.createElement("div");
  textWrap.className = "personal-service-row-main";
  textWrap.appendChild(nameBtn);
  textWrap.appendChild(meta);
  head.appendChild(iconBtn);
  head.appendChild(textWrap);
  row.appendChild(head);
  if (expandedPersonalServiceId === String(place.id)) {
    row.classList.add("is-selected");
    row.appendChild(buildExpandedPersonalServiceDetails(place));
  }
  return row;
}

function selectPersonalService(placeId) {
  expandedPersonalServiceId = expandedPersonalServiceId === String(placeId) ? "" : String(placeId);
  renderPersonalServicesList();
}

function buildExpandedPersonalServiceDetails(place) {
  const wrap = document.createElement("div");
  wrap.className = "personal-service-expanded";
  const tags = place.tags || {};
  const navBtn = createButton("Naviga", () => {
    window.open(`https://www.google.com/maps/dir/?api=1&destination=${place.lat},${place.lon}`, "_blank");
  });
  navBtn.classList.add("btn-primary");
  const googleDetailsBtn = createButton("Google dettagli", () => {
    const query = encodeURIComponent(`${place.name} ${place.lat},${place.lon}`);
    window.open(`https://www.google.com/maps/search/?api=1&query=${query}`, "_blank");
  });
  const closeBtn = createButton("Chiudi dettagli", () => selectPersonalService(place.id));
  const actions = document.createElement("div");
  actions.className = "item-actions";
  actions.appendChild(navBtn);
  actions.appendChild(googleDetailsBtn);
  actions.appendChild(closeBtn);
  wrap.innerHTML = `
    <p><b>Nome:</b> ${escapeHTML(place.name)}</p>
    <p><b>Distanza:</b> ${escapeHTML(formatDistance(place.distance))}</p>
    <p><b>Indirizzo:</b> ${escapeHTML(formatAddress(tags))}</p>
  `;
  wrap.appendChild(actions);
  const details = document.createElement("div");
  details.className = "simple-list";
  details.innerHTML = renderExtendedPersonalServiceDetails(place);
  wrap.appendChild(details);
  return wrap;
}

function renderExtendedPersonalServiceDetails(place) {
  const tags = place.tags || {};
  const cfg = PERSONAL_SERVICE_CATEGORIES[place.category];
  const rows = [];
  if (place.category === "lunch") {
    rows.push(`<p><b>Buoni pasto:</b> ${escapeHTML(formatMealVoucherStatus(tags))}</p>`);
  }
  const detailFields = Array.isArray(cfg?.detailFields) ? cfg.detailFields : [];
  detailFields.forEach((field) => {
    const rawValue = tags[field];
    if (rawValue == null || rawValue === "") return;
    rows.push(`<p><b>${escapeHTML(formatDetailFieldLabel(field))}:</b> ${escapeHTML(String(rawValue))}</p>`);
  });
  const allTagRows = Object.entries(tags)
    .filter(([key, value]) => value != null && String(value).trim() !== "")
    .sort(([a], [b]) => a.localeCompare(b))
    .map(([key, value]) => `<p><b>${escapeHTML(formatDetailFieldLabel(key))}:</b> ${escapeHTML(String(value))}</p>`);
  if (allTagRows.length) {
    rows.push("<hr>");
    rows.push("<p><b>Tutti i dati disponibili:</b></p>");
    rows.push(...allTagRows);
  }
  if (!rows.length) rows.push("<p class='muted'>Nessun dettaglio aggiuntivo disponibile.</p>");
  return rows.join("");
}

function formatDetailFieldLabel(field) {
  const labels = {
    opening_hours: "Orari",
    cuisine: "Tipo cucina",
    takeaway: "Take-away",
    delivery: "Consegna",
    "contact:phone": "Telefono",
    website: "Sito web",
    "payment:meal_voucher": "Buoni pasto",
    "payment:sodexo": "Sodexo",
    "payment:edenred": "Edenred",
    "payment:ticket_restaurant": "Ticket Restaurant",
    "diet:vegetarian": "Opzioni vegetariane",
    fee: "A pagamento",
    wheelchair: "Accessibilità carrozzina",
    operator: "Operatore",
    cash_in: "Versamento contanti",
    contactless: "Contactless",
    "currency:EUR": "Euro",
    dispensing: "Dispensazione",
    access: "Accesso",
    capacity: "Capacità",
    brand: "Marchio"
  };
  return labels[field] || field;
}

function formatAddress(tags) {
  const parts = [
    tags["addr:street"],
    tags["addr:housenumber"],
    tags["addr:city"]
  ].filter(Boolean);
  return parts.length ? parts.join(", ") : "Non disponibile";
}

function isMealVoucherAccepted(tags) {
  if (!tags) return false;
  const positiveFields = ["payment:meal_voucher", "payment:sodexo", "payment:edenred", "payment:ticket_restaurant"];
  return positiveFields.some((field) => String(tags[field] || "").toLowerCase() === "yes");
}

function formatMealVoucherStatus(tags) {
  if (isMealVoucherAccepted(tags)) return "Accettati";
  const fields = ["payment:meal_voucher", "payment:sodexo", "payment:edenred", "payment:ticket_restaurant"];
  if (fields.some((field) => String(tags[field] || "").toLowerCase() === "no")) return "Non accettati";
  return "Non specificato";
}

function getSelectedPersonalServicesRadius() {
  const value = Number(ui.personalServicesRadius?.value || 3000);
  if (!Number.isFinite(value) || value < 500) return 3000;
  return value;
}

async function fetchPersonalServicesFromOverpass(category, lat, lng, radiusMeters = 3000) {
  const cfg = PERSONAL_SERVICE_CATEGORIES[category];
  if (!cfg) return { elements: [] };
  const fragment = cfg.query
    .replaceAll("{lat}", String(lat))
    .replaceAll("{lng}", String(lng))
    .replaceAll("{radius}", String(radiusMeters));
  const query = `
    [out:json][timeout:25];
    (
      ${fragment}
    );
    out center tags;
  `;
  const firstResult = await fetchOverpassWithFallback(query);
  if (category !== "lunch" || (firstResult.elements || []).length) return firstResult;
  const broadLunchQuery = `
    [out:json][timeout:25];
    (
      node["amenity"~"restaurant|fast_food|food_court|canteen|cafe|bar"](around:${Math.max(20000, radiusMeters)},${lat},${lng});
      way["amenity"~"restaurant|fast_food|food_court|canteen|cafe|bar"](around:${Math.max(20000, radiusMeters)},${lat},${lng});
      relation["amenity"~"restaurant|fast_food|food_court|canteen|cafe|bar"](around:${Math.max(20000, radiusMeters)},${lat},${lng});
    );
    out center tags;
  `;
  return fetchOverpassWithFallback(broadLunchQuery);
}

async function fetchOverpassWithFallback(query) {
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

function initGeolocation() {
  if (!navigator.geolocation) {
    ui.gpsStatus.textContent = "Geolocalizzazione non supportata dal browser.";
    return;
  }
  if (geolocationWatchId != null) navigator.geolocation.clearWatch(geolocationWatchId);

  const onPosition = (pos) => {
    currentUserPos = {
      lat: pos.coords.latitude,
      lng: pos.coords.longitude
    };
    ui.gpsStatus.textContent = "Posizione attiva per ordinare gli impianti per distanza.";
    renderImpianti();
    renderMap();
    evaluateImpiantoProximityAlerts();
  };

  navigator.geolocation.getCurrentPosition((pos) => {
    onPosition(pos);
    fetchWeather();
  }, () => {
    ui.gpsStatus.textContent = "Posizione non disponibile. Elenco non ordinato per distanza reale.";
    fetchWeather();
  }, {
    enableHighAccuracy: true,
    timeout: 8000
  });

  geolocationWatchId = navigator.geolocation.watchPosition(onPosition, () => {
    ui.gpsStatus.textContent = "Posizione non disponibile. Elenco non ordinato per distanza reale.";
  }, {
    enableHighAccuracy: true,
    maximumAge: 10000,
    timeout: 10000
  });
}

function buildImpiantoEventLocalKey(eventType, commessaId, impiantoKey) {
  return `heraNotified:${eventType}:${commessaId || ""}:${impiantoKey || ""}`;
}

function hasLocalImpiantoEvent(eventType, commessaId, impiantoKey) {
  return localStorage.getItem(buildImpiantoEventLocalKey(eventType, commessaId, impiantoKey)) === "1";
}

function markLocalImpiantoEvent(eventType, commessaId, impiantoKey) {
  localStorage.setItem(buildImpiantoEventLocalKey(eventType, commessaId, impiantoKey), "1");
}

function evaluateImpiantoProximityAlerts() {
  if (!selectedCommessaId || !currentUserPos || !Array.isArray(currentImpianti) || !currentImpianti.length) return;
  const todoSorted = combineImpiantiForView(currentImpianti)
    .filter((impianto) => !impianto.done)
    .sort((a, b) => distanceFromUser(a) - distanceFromUser(b));
  const nearest = todoSorted[0];
  if (!nearest) {
    activeNearbyImpiantoContext = null;
    return;
  }
  const nearestKey = buildImpiantoKey(nearest);
  const distanceKm = distanceFromUser(nearest);
  if (!Number.isFinite(distanceKm)) return;

  if (!activeNearbyImpiantoContext && distanceKm <= PROXIMITY_NEAR_KM) {
    activeNearbyImpiantoContext = { commessaId: selectedCommessaId, impiantoKey: nearestKey };
    if (!hasLocalImpiantoEvent("near", selectedCommessaId, nearestKey)) {
      markLocalImpiantoEvent("near", selectedCommessaId, nearestKey);
      const commessaName = selectedCommessaName || "Commessa";
      const impiantoName = nearest.denominazione || "Impianto";
      publishGlobalNotificationEvent("impianto-near", {
        title: "Operatore vicino impianto",
        body: `${currentUser?.displayName || currentUser?.email || "Operatore"} è vicino a ${impiantoName} (${commessaName}).`,
        commessaId: selectedCommessaId,
        commessaName,
        impiantoName,
        impiantoKey: nearestKey
      });
    }
    return;
  }

  if (!activeNearbyImpiantoContext) return;
  if (activeNearbyImpiantoContext.commessaId !== selectedCommessaId || activeNearbyImpiantoContext.impiantoKey !== nearestKey) return;
  if (distanceKm < PROXIMITY_AWAY_KM) return;
  if (nearest.done) {
    activeNearbyImpiantoContext = null;
    return;
  }
  if (!hasLocalImpiantoEvent("away-without-done", selectedCommessaId, nearestKey)) {
    markLocalImpiantoEvent("away-without-done", selectedCommessaId, nearestKey);
    const commessaName = selectedCommessaName || "Commessa";
    const impiantoName = nearest.denominazione || "Impianto";
    publishGlobalNotificationEvent("impianto-away-without-done", {
      title: "Allontanamento senza FATTO",
      body: `${currentUser?.displayName || currentUser?.email || "Operatore"} si è allontanato da ${impiantoName} senza premere FATTO.`,
      commessaId: selectedCommessaId,
      commessaName,
      impiantoName,
      impiantoKey: nearestKey
    });
  }
  activeNearbyImpiantoContext = null;
}

async function fetchWeather() {
  try {
    const lat = currentUserPos ? currentUserPos.lat : 44.4949;
    const lon = currentUserPos ? currentUserPos.lng : 11.3426;
    const url = `https://api.open-meteo.com/v1/forecast?latitude=${lat}&longitude=${lon}&current=temperature_2m,wind_speed_10m,weather_code&hourly=temperature_2m,precipitation_probability,snowfall,visibility,weather_code&forecast_days=2`;
    const response = await fetch(url);
    if (!response.ok) throw new Error("meteo non disponibile");
    const data = await response.json();
    const current = data.current || {};
    const weatherLabel = weatherCodeLabel(current.weather_code);
    ui.weatherSummary.textContent = `${weatherLabel} • ${Math.round(current.temperature_2m ?? 0)}°C • vento ${Math.round(current.wind_speed_10m ?? 0)} km/h`;
    renderWeatherDetails(data);
  } catch (error) {
    ui.weatherSummary.textContent = "Meteo non disponibile.";
    ui.weatherRisks.innerHTML = "<span class='weather-risk-chip'>⚠️ Nessun dato rischio disponibile</span>";
    ui.weatherDetails.innerHTML = "<p class='muted'>Impossibile caricare previsioni dettagliate.</p>";
  }
}

function renderWeatherDetails(data) {
  const times = (data.hourly && data.hourly.time) || [];
  const temps = (data.hourly && data.hourly.temperature_2m) || [];
  const rains = (data.hourly && data.hourly.precipitation_probability) || [];
  const snows = (data.hourly && data.hourly.snowfall) || [];
  const visibilities = (data.hourly && data.hourly.visibility) || [];
  const codes = (data.hourly && data.hourly.weather_code) || [];
  const maxRain = Math.max(...rains.slice(0, 12).map((value) => Number(value) || 0), 0);
  const snowSum = snows.slice(0, 12).reduce((acc, value) => acc + (Number(value) || 0), 0);
  const minVisibility = Math.min(...visibilities.slice(0, 12).map((value) => Number(value) || Number.MAX_SAFE_INTEGER));
  const hasFogCode = codes.slice(0, 12).some((value) => Number(value) === 45 || Number(value) === 48);
  const riskIce = temps.slice(0, 12).some((value, idx) => Number(value) <= 1 && Number(rains[idx] || 0) >= 40);

  const risks = [];
  risks.push(maxRain >= 60 ? "🌧️ Rischio pioggia alta" : "🌧️ Rischio pioggia bassa");
  if (snowSum > 0) risks.push("❄️ Possibile neve");
  if (hasFogCode || minVisibility < 1200) risks.push("🌫️ Possibile nebbia");
  if (riskIce) risks.push("🧊 Possibile ghiaccio");
  ui.weatherRisks.innerHTML = risks.map((risk) => `<span class='weather-risk-chip'>${risk}</span>`).join("");

  const rows = times.slice(0, 12).map((time, idx) => {
    const hour = new Date(time).toLocaleTimeString("it-IT", { hour: "2-digit", minute: "2-digit", hour12: false });
    const visKm = ((Number(visibilities[idx]) || 0) / 1000).toFixed(1);
    const label = weatherCodeLabel(codes[idx]);
    return `<p><b>${hour}</b> • ${label} • 🌡️ ${Math.round(temps[idx] ?? 0)}°C • 🌧️ ${Math.round(rains[idx] ?? 0)}% • ❄️ ${Number(snows[idx] || 0).toFixed(1)} mm • 👁️ ${visKm} km</p>`;
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

function getTrafficIntensityByHour(date = new Date()) {
  const hour = date.getHours();
  const day = date.getDay();
  const isWeekend = day === 0 || day === 6;

  if (!isWeekend && ((hour >= 7 && hour <= 9) || (hour >= 17 && hour <= 19))) return "intenso";
  if (hour >= 22 || hour <= 5) return "leggero";
  if (isWeekend && hour >= 12 && hour <= 14) return "moderato";
  return "moderato";
}

function getDistanceIntensityOffset(distanceKm) {
  if (!Number.isFinite(distanceKm)) return 0;
  if (distanceKm < 1.2) return -1;
  if (distanceKm < 6) return 0;
  if (distanceKm < 20) return 1;
  return 0;
}

function getRouteVarianceOffset(distanceKm) {
  if (!Number.isFinite(distanceKm)) return 0;
  const fingerprint = Math.floor(distanceKm * 1000) % 5;
  if (fingerprint === 0) return -1;
  if (fingerprint === 4) return 1;
  return 0;
}

function normalizeTrafficIntensity(baseIntensity, distanceKm) {
  const levels = ["leggero", "moderato", "intenso"];
  const baseIndex = levels.indexOf(baseIntensity);
  if (baseIndex === -1) return "moderato";

  const distanceOffset = getDistanceIntensityOffset(distanceKm);
  const varianceOffset = getRouteVarianceOffset(distanceKm);
  const normalizedIndex = Math.max(0, Math.min(levels.length - 1, baseIndex + distanceOffset + varianceOffset));
  return levels[normalizedIndex];
}

function estimateTravelMeta(distanceKm) {
  if (!Number.isFinite(distanceKm) || distanceKm > 1e10) {
    return { intensityKey: "na", intensityLabel: "N/D", etaLabel: "N/D" };
  }
  const baseIntensity = getTrafficIntensityByHour();
  const intensityKey = normalizeTrafficIntensity(baseIntensity, distanceKm);
  const intensityLabel = intensityKey.charAt(0).toUpperCase() + intensityKey.slice(1);
  const speedByIntensity = {
    intenso: 25,
    moderato: 40,
    leggero: 60
  };
  const avgSpeed = speedByIntensity[intensityKey] || 35;
  const etaMinutes = Math.max(1, Math.round((Math.max(distanceKm, 0) / avgSpeed) * 60));
  return {
    intensityKey,
    intensityLabel,
    etaLabel: `${etaMinutes} min`
  };
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

function isChatMessageFresh(message) {
  const createdAtMs = message?.createdAt && typeof message.createdAt.toDate === "function"
    ? message.createdAt.toDate().getTime()
    : 0;
  if (!createdAtMs) return true;
  return createdAtMs >= (Date.now() - CHAT_RETENTION_MS);
}

function startChatRetentionLoop() {
  stopChatRetentionLoop();
  purgeOldChatMessages();
  chatRetentionTimer = setInterval(() => {
    purgeOldChatMessages();
  }, 60 * 60 * 1000);
}

function stopChatRetentionLoop() {
  if (chatRetentionTimer) {
    clearInterval(chatRetentionTimer);
    chatRetentionTimer = null;
  }
}

async function purgeOldChatMessages() {
  if (!canManageData()) return;
  const cutoffDate = new Date(Date.now() - CHAT_RETENTION_MS);
  try {
    const snapshot = await db
      .collection("chatMessages")
      .where("createdAt", "<=", firebase.firestore.Timestamp.fromDate(cutoffDate))
      .limit(200)
      .get();
    if (snapshot.empty) return;
    const batch = db.batch();
    snapshot.docs.forEach((doc) => batch.delete(doc.ref));
    await batch.commit();
  } catch (error) {
    console.warn("Pulizia chat 24h non completata:", error);
  }
}

async function upsertCurrentPlatformUser() {
  if (!currentUser) return;
  await db.collection("platformUsers").doc(currentUser.uid).set({
    uid: currentUser.uid,
    email: currentUser.email || "",
    displayName: currentUser.displayName || currentUser.email || "Utente",
    isAdmin: canManageData(),
    notificationsAutoEnabled: isAutoNotificationEnabled(),
    lastSeenAt: firebase.firestore.FieldValue.serverTimestamp()
  }, { merge: true });
}

function subscribeAdminUsers() {
  unsubscribeAdminUsers = db.collection("appConfig").doc("adminUsers").onSnapshot((doc) => {
    const data = doc.exists ? doc.data() : {};
    const rawList = Array.isArray(data.emails) ? data.emails : [];
    const normalized = rawList
      .map((email) => normalizeEmail(email))
      .filter(Boolean);
    adminEmails = new Set([ADMIN_EMAIL, ...normalized]);
    updateAdminControls();
    renderCommesseManagementList();
    renderAdminUsers();
  }, (error) => {
    console.error("Errore caricamento admin users:", error);
    adminEmails = new Set([ADMIN_EMAIL]);
    updateAdminControls();
    renderCommesseManagementList();
    renderAdminUsers();
  });
}

function stopAdminUsersSubscription() {
  if (unsubscribeAdminUsers) {
    unsubscribeAdminUsers();
    unsubscribeAdminUsers = null;
  }
  adminEmails = new Set([ADMIN_EMAIL]);
  renderAdminUsers();
}

function subscribeUsers() {
  unsubscribeUsers = db.collection("platformUsers").onSnapshot((snapshot) => {
    platformUsers = snapshot.docs.map((doc) => ({ id: doc.id, ...doc.data() }))
      .sort((a, b) => String(a.displayName || "").localeCompare(String(b.displayName || ""), "it"));
    syncNotificationAutoPreferenceFromProfile();
    maybeAutoEnableNotifications();
    deniedImpiantoActions = getDeniedActionsForCurrentUser();
    renderChatRecipients();
    renderUserPermissionList();
    renderHeaderActivitySummary();
    renderExternalApps();
  });
}

function stopUsersSubscription() {
  if (unsubscribeUsers) {
    unsubscribeUsers();
    unsubscribeUsers = null;
  }
  platformUsers = [];
  deniedImpiantoActions = new Set();
  renderChatRecipients();
  renderUserPermissionList();
  renderHeaderActivitySummary();
  renderExternalApps();
}

function canApproveHoursLevel1(request) {
  if (!currentUser || !request) return false;
  if (canManageData()) return true;
  return String(request.createdByUid || "") === String(currentUser.uid || "");
}

function renderHoursApprovalRequests() {
  if (!ui.hoursApprovalsList || !ui.hoursApprovalsFeedback) return;
  if (!currentUser) {
    ui.hoursApprovalsFeedback.textContent = "Fai login per vedere le richieste ore.";
    ui.hoursApprovalsList.innerHTML = "";
    return;
  }
  const visible = hoursApprovalRequests.filter((request) => canManageData() || String(request.createdByUid || "") === String(currentUser.uid || ""));
  if (!visible.length) {
    ui.hoursApprovalsFeedback.textContent = "Nessuna richiesta ore in approvazione.";
    ui.hoursApprovalsList.innerHTML = "";
    return;
  }
  ui.hoursApprovalsFeedback.textContent = `Richieste trovate: ${visible.length}.`;
  ui.hoursApprovalsList.innerHTML = "";
  visible.forEach((request) => {
    const card = document.createElement("article");
    card.className = "item-card";
    const dateLabel = request.date ? new Date(`${request.date}T00:00:00`).toLocaleDateString("it-IT") : "-";
    const statusMap = {
      pending_level1: "In attesa primo OK",
      pending_admin: "In attesa OK admin finale",
      approved: `Approvata ✅ (report ${request.finalizedReportId || "-"})`,
      rejected: "Rifiutata ❌"
    };
    const statusText = statusMap[request.status] || request.status || "-";
    const author = request.createdByName || request.createdByEmail || "Operatore";
    const summary = (Array.isArray(request.entries) ? request.entries : []).map((entry) => {
      const tot = (entry.rows || []).reduce((sum, row) => sum + (Number(row.ore || 0) || 0), 0);
      return `<li>${escapeHTML(entry.commessaName || "Commessa")}: ${escapeHTML(String(tot))}h</li>`;
    }).join("");
    card.innerHTML = `
      <p><b>ID:</b> ${escapeHTML(request.id || "-")}</p>
      <p><b>Data:</b> ${escapeHTML(dateLabel)} • <b>Creato da:</b> ${escapeHTML(author)}</p>
      <p><b>Stato:</b> ${escapeHTML(statusText)}</p>
      <ul>${summary || "<li>Nessuna commessa</li>"}</ul>
      ${request.rejectionReason ? `<p><b>Motivo rifiuto:</b> ${escapeHTML(request.rejectionReason)}</p>` : ""}
    `;
    const actions = document.createElement("div");
    actions.className = "item-actions";
    if (request.status === "pending_level1" && canApproveHoursLevel1(request)) {
      actions.appendChild(createButton("Accetta livello 1", () => approveHoursRequestLevel1(request)));
      actions.appendChild(createButton("Rifiuta", () => rejectHoursRequest(request)));
    }
    if (request.status === "pending_admin" && canManageData()) {
      actions.appendChild(createButton("Accetta admin finale", () => approveHoursRequestLevel2(request)));
      actions.appendChild(createButton("Rifiuta", () => rejectHoursRequest(request)));
    }
    if (actions.children.length) card.appendChild(actions);
    ui.hoursApprovalsList.appendChild(card);
  });
}



async function approveHoursRequestLevel1(request) {
  if (!canApproveHoursLevel1(request)) {
    alert("Non autorizzato al primo livello di approvazione.");
    return;
  }
  await db.collection("oreApprovalRequests").doc(request.id).set({
    status: "pending_admin",
    level1ApprovedBy: currentUser.email || currentUser.displayName || "utente",
    level1ApprovedAt: firebase.firestore.FieldValue.serverTimestamp(),
    updatedAt: firebase.firestore.FieldValue.serverTimestamp()
  }, { merge: true });
  await notifyAdminsForFinalHoursApproval(request.id, request, currentUser.displayName || currentUser.email || "utente");
  const requester = platformUsers.find((user) => String(user.uid || "") === String(request.createdByUid || ""));
  if (requester?.id) {
    await sendPrivateChatNotification({
      recipientId: requester.id,
      text: `✅ Richiesta ore ${request.id}: primo livello approvato. In attesa conferma admin finale.`,
      senderName: currentUser.displayName || currentUser.email || "Sistema"
    });
  }
}

async function approveHoursRequestLevel2(request) {
  if (!canManageData()) {
    alert("Solo admin può confermare il livello finale.");
    return;
  }
  if (request.status !== "pending_admin") {
    alert("Questa richiesta non è più in attesa dell'approvazione admin.");
    return;
  }
  const conflicts = await findExistingHoursConflicts(
    String(request.date || "").trim(),
    Array.isArray(request.entries) ? request.entries : [],
    { skipApprovalRequestId: request.id }
  );
  if (conflicts.length) {
    const first = conflicts[0];
    alert(`Conferma bloccata: ${first.operatore || "Operatore"} ha già ore inserite nella commessa ${first.commessaName || "Commessa"} per questa data. Elimina prima la riga esistente.`);
    return;
  }
  const reportPayload = {
    date: request.date || "",
    entries: Array.isArray(request.entries) ? request.entries : [],
    createdByUid: request.createdByUid || "",
    createdByName: request.createdByName || request.createdByEmail || "Operatore",
    createdByEmail: request.createdByEmail || "",
    createdAt: firebase.firestore.FieldValue.serverTimestamp()
  };
  const docRef = await db.collection("oreReports").add(reportPayload);
  let driveLink = "";
  if (driveAccessToken) {
    if (!driveReportsFolderId) await ensureDriveFolders();
    const blob = new Blob([JSON.stringify({ id: docRef.id, ...reportPayload, createdAtIso: new Date().toISOString() }, null, 2)], { type: "application/json" });
    const fileName = `ore_${reportPayload.date}_${docRef.id}.json`;
    const upload = await uploadBlobToDrive(blob, fileName, "application/json", driveReportsFolderId);
    driveLink = upload?.webViewLink || "";
  }
  await db.collection("oreApprovalRequests").doc(request.id).set({
    status: "approved",
    level2ApprovedBy: currentUser.email || currentUser.displayName || "admin",
    level2ApprovedAt: firebase.firestore.FieldValue.serverTimestamp(),
    finalizedReportId: docRef.id,
    driveBackupLink: driveLink,
    updatedAt: firebase.firestore.FieldValue.serverTimestamp()
  }, { merge: true });
  const targetUser = platformUsers.find((user) => String(user.uid || "") === String(request.createdByUid || ""));
  if (targetUser?.id) {
    await sendPrivateChatNotification({
      recipientId: targetUser.id,
      text: `✅ Richiesta ore ${request.id} approvata definitivamente. Report salvato: ${docRef.id}.`,
      senderName: currentUser.displayName || currentUser.email || "Admin"
    });
  }
  loadSavedHoursReports();
}

async function approveHoursRequestFromChat(requestId) {
  if (!canManageData()) {
    throw new Error("Solo admin può confermare le ore da chat.");
  }
  if (!requestId) {
    throw new Error("ID richiesta non valido.");
  }
  const request = await getHoursApprovalRequestById(requestId);
  if (!request) {
    throw new Error("Richiesta ore non trovata.");
  }
  if (String(request.status || "") !== "pending_admin") {
    throw new Error(`Idempotenza: richiesta non pending_admin (stato: ${String(request.status || "sconosciuto")}).`);
  }
  await approveHoursRequestLevel2(request);
}

async function rejectHoursRequestFromChat(requestId) {
  if (!canManageData()) {
    throw new Error("Solo admin può rifiutare le ore da chat.");
  }
  if (!requestId) {
    throw new Error("ID richiesta non valido.");
  }
  const request = await getHoursApprovalRequestById(requestId);
  if (!request) {
    throw new Error("Richiesta ore non trovata.");
  }
  if (String(request.status || "") !== "pending_admin") {
    throw new Error(`Idempotenza: richiesta non pending_admin (stato: ${String(request.status || "sconosciuto")}).`);
  }
  await db.collection("oreApprovalRequests").doc(request.id).set({
    status: "rejected",
    rejectedBy: currentUser.email || currentUser.displayName || "admin",
    rejectedAt: firebase.firestore.FieldValue.serverTimestamp(),
    rejectionReason: "Rifiutata da chat admin",
    updatedAt: firebase.firestore.FieldValue.serverTimestamp()
  }, { merge: true });
  const targetUser = platformUsers.find((user) => String(user.uid || "") === String(request.createdByUid || ""));
  if (targetUser?.id) {
    await sendPrivateChatNotification({
      recipientId: targetUser.id,
      text: `❌ Richiesta ore ${request.id} rifiutata da chat admin.`,
      senderName: currentUser.displayName || currentUser.email || "Admin"
    });
  }
}

async function rejectHoursRequest(request) {
  const canReject = (request.status === "pending_level1" && canApproveHoursLevel1(request))
    || (request.status === "pending_admin" && canManageData());
  if (!canReject) {
    alert("Non autorizzato a rifiutare questa richiesta.");
    return;
  }
  const reason = window.prompt("Motivo del rifiuto (opzionale):", "") || "";
  await db.collection("oreApprovalRequests").doc(request.id).set({
    status: "rejected",
    rejectedBy: currentUser.email || currentUser.displayName || "utente",
    rejectedAt: firebase.firestore.FieldValue.serverTimestamp(),
    rejectionReason: String(reason).trim(),
    updatedAt: firebase.firestore.FieldValue.serverTimestamp()
  }, { merge: true });
  const targetUser = platformUsers.find((user) => String(user.uid || "") === String(request.createdByUid || ""));
  if (targetUser?.id) {
    await sendPrivateChatNotification({
      recipientId: targetUser.id,
      text: `❌ Richiesta ore ${request.id} rifiutata.${reason ? ` Motivo: ${String(reason).trim()}` : ""}`,
      senderName: currentUser.displayName || currentUser.email || "Sistema"
    });
  }
}

function startPresenceHeartbeat() {
  stopPresenceHeartbeat();
  if (!currentUser) return;
  presenceHeartbeatTimer = setInterval(() => {
    upsertCurrentPlatformUser();
  }, 60 * 1000);
}

function stopPresenceHeartbeat() {
  if (presenceHeartbeatTimer) {
    clearInterval(presenceHeartbeatTimer);
    presenceHeartbeatTimer = null;
  }
}

function subscribeGpsRequests() {
  if (unsubscribeGpsRequests) unsubscribeGpsRequests();
  unsubscribeGpsRequests = db
    .collection("gpsUpdateRequests")
    .orderBy("createdAt", "desc")
    .limit(200)
    .onSnapshot((snapshot) => {
      gpsUpdateRequests = snapshot.docs.map((doc) => ({ id: doc.id, ...doc.data() }));
      renderGpsRequests();
    }, (error) => {
      console.error("Errore caricamento richieste GPS:", error);
      if (ui.gpsRequestsList) ui.gpsRequestsList.innerHTML = "<p class='muted'>Errore caricamento richieste GPS.</p>";
    });
}

function stopGpsRequestsSubscription() {
  if (unsubscribeGpsRequests) {
    unsubscribeGpsRequests();
    unsubscribeGpsRequests = null;
  }
  gpsUpdateRequests = [];
  renderGpsRequests();
}

function renderGpsRequests() {
  if (!ui.gpsRequestsList) return;
  if (!currentUser) {
    ui.gpsRequestsList.innerHTML = "<p class='muted'>Fai login per visualizzare le richieste GPS.</p>";
    return;
  }
  if (!canManageData()) {
    ui.gpsRequestsList.innerHTML = "<p class='muted'>Solo gli admin possono gestire le richieste GPS.</p>";
    return;
  }
  if (!gpsUpdateRequests.length) {
    ui.gpsRequestsList.innerHTML = "<p class='muted'>Nessuna richiesta GPS.</p>";
    return;
  }
  ui.gpsRequestsList.innerHTML = "";
  gpsUpdateRequests.forEach((request) => {
    const row = document.createElement("div");
    row.className = "simple-list-item";
    const when = request.createdAt && typeof request.createdAt.toDate === "function"
      ? request.createdAt.toDate().toLocaleString("it-IT")
      : "-";
    const status = String(request.status || "pending").toUpperCase();
    const info = document.createElement("span");
    info.innerHTML = `${escapeHTML(request.impiantoDenominazione || "Impianto")} • ${escapeHTML(request.operatorName || "Operatore")}<br><small>${escapeHTML(request.operatorLat)}, ${escapeHTML(request.operatorLng)} • ${when} • ${status}</small>`;
    row.appendChild(info);

    const actions = document.createElement("div");
    actions.className = "item-actions";
    const canDecide = String(request.status || "pending") === "pending";
    actions.appendChild(createButton("Accetta", () => approveGpsRequest(request), !canDecide));
    actions.appendChild(createButton("Rifiuta", () => rejectGpsRequest(request), !canDecide));
    row.appendChild(actions);
    ui.gpsRequestsList.appendChild(row);
  });
}

async function approveGpsRequest(request) {
  if (!canManageData()) return;
  const impiantoIds = Array.isArray(request.impiantoIds) ? request.impiantoIds.filter(Boolean) : [];
  if (!request.commessaId || !impiantoIds.length) {
    alert("Richiesta non valida: impianto non trovato.");
    return;
  }
  await updateImpiantoCoordinates(request.commessaId, impiantoIds, request.operatorLat, request.operatorLng);
  await db.collection("gpsUpdateRequests").doc(request.id).set({
    status: "approved",
    approvedAt: firebase.firestore.FieldValue.serverTimestamp(),
    approvedBy: currentUser?.email || ""
  }, { merge: true });
  await notifyGpsDecision(request, true);
}

async function rejectGpsRequest(request) {
  if (!canManageData()) return;
  await db.collection("gpsUpdateRequests").doc(request.id).set({
    status: "rejected",
    rejectedAt: firebase.firestore.FieldValue.serverTimestamp(),
    rejectedBy: currentUser?.email || ""
  }, { merge: true });
  await notifyGpsDecision(request, false);
}

async function notifyGpsDecision(request, approved) {
  if (!request.operatorId) return;
  await db.collection("chatMessages").add({
    type: "text",
    text: approved
      ? `✅ Richiesta GPS accettata per ${request.impiantoDenominazione || "impianto"}. Coordinate aggiornate.`
      : `❌ Richiesta GPS rifiutata per ${request.impiantoDenominazione || "impianto"}.`,
    recipientId: request.operatorId,
    senderId: currentUser?.uid || "",
    senderName: currentUser?.displayName || currentUser?.email || "Admin",
    senderEmail: currentUser?.email || "",
    createdAt: firebase.firestore.FieldValue.serverTimestamp()
  });
}

async function updateImpiantoCoordinates(commessaId, impiantoIds, lat, lng) {
  const ref = db.collection("commesse").doc(commessaId).collection("impianti");
  await Promise.all(impiantoIds.map((impiantoId) => ref.doc(impiantoId).update({
    gpsY: Number(lat),
    gpsX: Number(lng),
    updatedAt: firebase.firestore.FieldValue.serverTimestamp(),
    updatedBy: currentUser?.email || "admin"
  })));
}

function getDeniedActionsForCurrentUser() {
  if (!currentUser) return new Set();
  const row = platformUsers.find((user) => user.id === currentUser.uid);
  const denied = Array.isArray(row?.deniedImpiantoActions) ? row.deniedImpiantoActions : [];
  return new Set(denied.filter((action) => IMPIANTO_ACTIONS.includes(action)));
}

function isImpiantoActionDenied(action) {
  return deniedImpiantoActions.has(action);
}

function actionLabel(action) {
  const map = {
    done: "✅ Fatto",
    navigate: "🧭 Naviga",
    reset: "♻️ Reset",
    whatsapp: "💬 WhatsApp",
    "problem-report": "🚨 Segnala problema",
    "gps-update": "📍 Aggiorna GPS",
    edit: "✏️ Modifica",
    delete: "🗑️ Elimina"
  };
  return map[action] || action;
}

function renderUserPermissionList() {
  if (!ui.userPermissionsList) return;
  if (!currentUser) {
    ui.userPermissionsList.innerHTML = "<p class='muted'>Fai login per gestire i permessi.</p>";
    return;
  }
  if (!canManageData()) {
    ui.userPermissionsList.innerHTML = "<p class='muted'>Solo gli admin possono cambiare i permessi azione.</p>";
    return;
  }
  const users = platformUsers.filter((user) => !adminEmails.has(normalizeEmail(user.email)));
  if (!users.length) {
    ui.userPermissionsList.innerHTML = "<p class='muted'>Nessun utente disponibile.</p>";
    return;
  }
  ui.userPermissionsList.innerHTML = "";
  users.forEach((user) => {
    const row = document.createElement("div");
    row.className = "simple-list-item stacked";
    const title = document.createElement("strong");
    title.textContent = user.displayName || user.email || user.id;
    row.appendChild(title);
    const actionBox = document.createElement("div");
    actionBox.className = "actions-row";
    const denied = new Set(Array.isArray(user.deniedImpiantoActions) ? user.deniedImpiantoActions : []);
    IMPIANTO_ACTIONS.forEach((action) => {
      const btn = createButton(actionLabel(action), async () => {
        const next = new Set(denied);
        if (next.has(action)) next.delete(action);
        else next.add(action);
        await db.collection("platformUsers").doc(user.id).set({
          deniedImpiantoActions: Array.from(next),
          updatedAt: firebase.firestore.FieldValue.serverTimestamp(),
          updatedBy: currentUser?.email || ""
        }, { merge: true });
      });
      btn.classList.add("btn-small");
      if (denied.has(action)) btn.classList.add("btn-primary");
      actionBox.appendChild(btn);
    });
    row.appendChild(actionBox);
    ui.userPermissionsList.appendChild(row);
  });
}

function renderExternalApps() {
  if (!ui.externalAppsList) return;
  if (!currentUser) {
    ui.externalAppsList.innerHTML = "<p class='muted'>Fai login per collegare app esterne.</p>";
    return;
  }
  const row = platformUsers.find((user) => user.id === currentUser.uid);
  const apps = Array.isArray(row?.externalApps) ? row.externalApps : [];
  if (!apps.length) {
    ui.externalAppsList.innerHTML = "<p class='muted'>Nessuna app esterna collegata.</p>";
    return;
  }
  ui.externalAppsList.innerHTML = "";
  apps.forEach((app, index) => {
    const item = document.createElement("div");
    item.className = "simple-list-item";
    const label = document.createElement("a");
    label.href = app.url;
    label.target = "_blank";
    label.rel = "noopener noreferrer";
    label.textContent = `🔗 ${app.name || app.url}`;
    item.appendChild(label);
    const removeBtn = createButton("Rimuovi", () => removeExternalApp(index));
    item.appendChild(removeBtn);
    ui.externalAppsList.appendChild(item);
  });
}

async function saveExternalAppForCurrentUser(event) {
  event.preventDefault();
  if (!currentUser) return;
  const name = String(ui.externalAppName.value || "").trim();
  const rawUrl = String(ui.externalAppUrl.value || "").trim();
  if (!name || !rawUrl) return;
  const url = /^https?:\/\//i.test(rawUrl) ? rawUrl : `https://${rawUrl}`;
  const row = platformUsers.find((user) => user.id === currentUser.uid);
  const apps = Array.isArray(row?.externalApps) ? row.externalApps : [];
  const next = [...apps, { name, url }];
  await db.collection("platformUsers").doc(currentUser.uid).set({
    externalApps: next,
    updatedAt: firebase.firestore.FieldValue.serverTimestamp(),
    updatedBy: currentUser.email || ""
  }, { merge: true });
  ui.externalAppForm.reset();
}

async function removeExternalApp(index) {
  if (!currentUser) return;
  const row = platformUsers.find((user) => user.id === currentUser.uid);
  const apps = Array.isArray(row?.externalApps) ? row.externalApps : [];
  if (index < 0 || index >= apps.length) return;
  const next = apps.filter((_, i) => i !== index);
  await db.collection("platformUsers").doc(currentUser.uid).set({
    externalApps: next,
    updatedAt: firebase.firestore.FieldValue.serverTimestamp(),
    updatedBy: currentUser.email || ""
  }, { merge: true });
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

function renderCommesseManagementList() {
  if (!ui.commesseManageList) return;
  if (!canManageData()) {
    ui.commesseManageList.innerHTML = "<p class='muted'>Solo gli admin possono rinominare, svuotare o eliminare commesse.</p>";
    return;
  }
  const commesse = Array.from(commesseById.values());
  if (!commesse.length) {
    ui.commesseManageList.innerHTML = "<p class='muted'>Nessuna commessa disponibile.</p>";
    return;
  }
  ui.commesseManageList.innerHTML = "";
  commesse.forEach((commessa) => {
    const row = document.createElement("div");
    row.className = "simple-list-item";
    const title = document.createElement("span");
    title.textContent = commessa.nome || "Commessa senza nome";

    const actions = document.createElement("div");
    actions.className = "item-actions";
    actions.appendChild(createButton("Rinomina", () => renameCommessa(commessa.id, commessa.nome || "Commessa")));
    actions.appendChild(createButton("Svuota", () => clearCommessaImpianti(commessa.id, commessa.nome || "Commessa")));
    actions.appendChild(createButton("Elimina", () => deleteCommessa(commessa.id, commessa.nome || "Commessa")));

    row.appendChild(title);
    row.appendChild(actions);
    ui.commesseManageList.appendChild(row);
  });
}

async function renameCommessa(commessaId, currentName) {
  if (!canManageData()) {
    alert("Solo un admin può rinominare commesse.");
    return;
  }
  const nextName = window.prompt("Nuovo nome commessa:", currentName || "");
  if (nextName == null) return;
  const normalized = nextName.trim();
  if (!normalized) return;
  await db.collection("commesse").doc(commessaId).set({ nome: normalized }, { merge: true });
  if (selectedCommessaId === commessaId) {
    selectedCommessaName = normalized;
    ui.commessaAttiva.textContent = `Commessa selezionata: ${normalized}`;
  }
}

async function clearCommessaImpianti(commessaId, nome) {
  if (!canManageData()) {
    alert("Solo un admin può svuotare commesse.");
    return;
  }
  const ok = window.confirm(
    `ATTENZIONE: stai per svuotare la commessa "${nome}" ed eliminare tutti gli impianti.\n\nPremi OK per confermare, Annulla per tornare indietro.`
  );
  if (!ok) return;
  const impiantiRef = db.collection("commesse").doc(commessaId).collection("impianti");
  await deleteCollectionDocs(impiantiRef);
}

async function deleteCollectionDocs(collectionRef, batchSize = 200) {
  while (true) {
    const snapshot = await collectionRef.limit(batchSize).get();
    if (snapshot.empty) break;
    const batch = db.batch();
    snapshot.docs.forEach((doc) => batch.delete(doc.ref));
    await batch.commit();
  }
}

function renderAdminUsers() {
  if (!ui.adminUsersList) return;
  if (!canManageData()) {
    ui.adminUsersList.innerHTML = "<p class='muted'>Solo un admin può gestire i permessi admin.</p>";
    return;
  }
  const emails = Array.from(adminEmails).sort((a, b) => a.localeCompare(b, "it"));
  ui.adminUsersList.innerHTML = "";
  emails.forEach((email) => {
    const row = document.createElement("div");
    row.className = "simple-list-item";
    const label = document.createElement("span");
    label.textContent = email;
    row.appendChild(label);
    if (email !== ADMIN_EMAIL) {
      const revokeBtn = createButton("Rimuovi", () => removeAdminEmail(email));
      row.appendChild(revokeBtn);
    }
    ui.adminUsersList.appendChild(row);
  });
}

async function addAdminUserByEmail(event) {
  event.preventDefault();
  if (!canManageData()) {
    alert("Solo un admin può aggiungere altri admin.");
    return;
  }
  const email = normalizeEmail(ui.adminUserEmail.value);
  if (!email || !email.includes("@")) {
    alert("Inserisci un'email valida.");
    return;
  }
  const next = Array.from(new Set([...adminEmails, email]));
  await db.collection("appConfig").doc("adminUsers").set({
    emails: next,
    updatedAt: firebase.firestore.FieldValue.serverTimestamp(),
    updatedBy: currentUser?.email || ""
  }, { merge: true });
  ui.adminUserForm.reset();
}

async function removeAdminEmail(email) {
  if (!canManageData()) {
    alert("Solo un admin può rimuovere admin.");
    return;
  }
  const normalized = normalizeEmail(email);
  if (!normalized || normalized === ADMIN_EMAIL) return;
  const ok = window.confirm(`Rimuovere i permessi admin per ${normalized}?`);
  if (!ok) return;
  const next = Array.from(adminEmails).filter((item) => item !== normalized);
  await db.collection("appConfig").doc("adminUsers").set({
    emails: next,
    updatedAt: firebase.firestore.FieldValue.serverTimestamp(),
    updatedBy: currentUser?.email || ""
  }, { merge: true });
}

function stopChatSubscription() {
  if (unsubscribeChat) {
    unsubscribeChat();
    unsubscribeChat = null;
  }
  chatMessages = [];
}

async function renderChat(messages) {
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

  const visibleMessages = messages.filter((message) => canViewMessage(message) && isChatMessageFresh(message));

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
  const messageElements = await Promise.all(visibleMessages.map((message) => createChatMessageElement(message)));
  messageElements.forEach((element) => {
    if (element) ui.chatFullList.appendChild(element);
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

function canConfirmHoursFromChat(message) {
  if (String(message?.kind || "") !== "system") return false;
  const metadata = message?.metadata && typeof message.metadata === "object" ? message.metadata : null;
  if (!metadata) return false;
  if (metadata.type !== "hours_approval" || !metadata.approvalRequestId) return false;
  if (metadata.action === "level1_ok") return true;
  if (metadata.action === "admin_final_ok") return canManageData();
  return false;
}

async function getHoursApprovalRequestById(requestId) {
  const normalizedRequestId = String(requestId || "").trim();
  if (!normalizedRequestId) return null;
  const docSnap = await db.collection("oreApprovalRequests").doc(normalizedRequestId).get();
  if (!docSnap.exists) return null;
  return { id: docSnap.id, ...docSnap.data() };
}

function setChatHoursActionButtonsState(acceptButton, rejectButton, state) {
  const bothButtons = [acceptButton, rejectButton].filter(Boolean);
  bothButtons.forEach((button) => {
    button.disabled = state;
  });
}

async function buildChatMessageActions(message) {
  if (!canConfirmHoursFromChat(message)) return null;
  const actionType = String(message?.metadata?.action || "").trim();
  const approvalRequestId = String(message?.metadata?.approvalRequestId || "").trim();
  const resolvedRequest = await getHoursApprovalRequestById(approvalRequestId);
  const actions = document.createElement("div");
  actions.className = "item-actions";
  if (!resolvedRequest) {
    actions.appendChild(createButton("Richiesta non trovata", () => {}, true));
    return actions;
  }

  const expectedStatus = actionType === "level1_ok" ? "pending_level1" : "pending_admin";
  const canAct = actionType === "level1_ok"
    ? canApproveHoursLevel1(resolvedRequest)
    : canManageData();

  if (resolvedRequest.status !== expectedStatus) {
    const statusMap = {
      pending_level1: "In attesa primo OK",
      pending_admin: "In attesa admin finale",
      approved: "Già approvata",
      rejected: "Già rifiutata"
    };
    actions.appendChild(createButton(statusMap[resolvedRequest.status] || "Già gestita", () => {}, true));
    return actions;
  }
  if (!canAct) return null;

  const acceptButton = createButton("Accetta", async () => {
    setChatHoursActionButtonsState(acceptButton, rejectButton, true);
    try {
      if (actionType === "level1_ok") {
        await approveHoursRequestLevel1(resolvedRequest);
        ui.chatFeedback.textContent = "Primo OK registrato.";
      } else {
        await approveHoursRequestFromChat(approvalRequestId);
        ui.chatFeedback.textContent = "Ore risultate confermate.";
      }
    } catch (error) {
      console.error("Errore conferma ore da chat:", error);
      const latestState = await getHoursApprovalRequestById(approvalRequestId);
      if (!latestState || latestState.status !== expectedStatus) {
        ui.chatFeedback.textContent = "Richiesta già gestita.";
        setChatHoursActionButtonsState(acceptButton, rejectButton, true);
        return;
      }
      ui.chatFeedback.textContent = error?.message || "Errore durante la conferma ore dalla chat.";
      setChatHoursActionButtonsState(acceptButton, rejectButton, false);
    }
  });

  const rejectButton = createButton("Rifiuta", async () => {
    setChatHoursActionButtonsState(acceptButton, rejectButton, true);
    try {
      if (actionType === "level1_ok") {
        await rejectHoursRequest(resolvedRequest);
      } else {
        await rejectHoursRequestFromChat(approvalRequestId);
      }
      ui.chatFeedback.textContent = "Richiesta ore rifiutata.";
    } catch (error) {
      console.error("Errore rifiuto ore da chat:", error);
      const latestState = await getHoursApprovalRequestById(approvalRequestId);
      if (!latestState || latestState.status !== expectedStatus) {
        ui.chatFeedback.textContent = "Richiesta già gestita.";
        setChatHoursActionButtonsState(acceptButton, rejectButton, true);
        return;
      }
      ui.chatFeedback.textContent = error?.message || "Errore durante il rifiuto ore dalla chat.";
      setChatHoursActionButtonsState(acceptButton, rejectButton, false);
    }
  });
  actions.appendChild(acceptButton);
  actions.appendChild(rejectButton);
  return actions;
}

async function createChatMessageElement(message) {
  const item = document.createElement("article");
  item.className = "chat-message" + (isOwnMessage(message) ? " own" : "");
  const isIncomingPrivate = Boolean(message.recipientId && !isOwnMessage(message));

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

  const contentWrap = document.createElement("div");
  contentWrap.className = isIncomingPrivate ? "chat-private-content" : "chat-message-content";

  const messageText = typeof message.text === "string" && message.text.trim()
    ? message.text
    : typeof message.message === "string" && message.message.trim()
      ? message.message
      : typeof message.body === "string" && message.body.trim()
        ? message.body
        : typeof message.content === "string" && message.content.trim()
          ? message.content
          : "";

  if ((message.type === "text" || (!message.type && messageText)) && messageText) {
    const p = document.createElement("p");
    p.className = "chat-text";
    p.textContent = messageText;
    contentWrap.appendChild(p);
  }

  const mediaSource = message.mediaUrl || message.mediaDataUrl || "";
  const mediaMimeType = String(message.mediaMimeType || "").toLowerCase();
  const hasImageMedia = message.type === "image" || mediaMimeType.startsWith("image/");
  const hasVideoMedia = message.type === "video" || mediaMimeType.startsWith("video/");
  const hasVoiceMedia = message.type === "voice" || mediaMimeType.startsWith("audio/");

  if (hasImageMedia && mediaSource) {
    const img = document.createElement("img");
    img.className = "chat-media-preview";
    img.src = mediaSource;
    img.alt = "Immagine inviata in chat";
    contentWrap.appendChild(img);
  }

  if (hasVideoMedia && mediaSource) {
    const video = document.createElement("video");
    video.className = "chat-media-preview";
    video.src = mediaSource;
    video.controls = true;
    contentWrap.appendChild(video);
  }

  if (hasVoiceMedia && mediaSource) {
    const audio = document.createElement("audio");
    audio.src = mediaSource;
    audio.controls = true;
    audio.className = "chat-audio";
    contentWrap.appendChild(audio);
  }

  const webViewLink = String(message.mediaDriveWebViewLink || "").trim();
  if (!mediaSource && webViewLink) {
    const openLink = document.createElement("a");
    openLink.className = "btn";
    openLink.href = webViewLink;
    openLink.target = "_blank";
    openLink.rel = "noopener noreferrer";
    openLink.textContent = "Apri allegato";
    contentWrap.appendChild(openLink);
  }

  const hasContent = contentWrap.childElementCount > 0;
  if (isIncomingPrivate) {
    const toggleBtn = document.createElement("button");
    toggleBtn.type = "button";
    toggleBtn.className = "btn btn-chat-private-toggle";
    toggleBtn.textContent = "Apri messaggio";
    toggleBtn.addEventListener("click", () => {
      const opened = contentWrap.classList.toggle("is-open");
      toggleBtn.textContent = opened ? "Chiudi messaggio" : "Apri messaggio";
    });
    item.appendChild(toggleBtn);
    if (hasContent) {
      item.appendChild(contentWrap);
    } else {
      const fallback = document.createElement("p");
      fallback.className = "muted";
      fallback.textContent = "Contenuto non disponibile.";
      item.appendChild(fallback);
    }
    const actions = await buildChatMessageActions(message);
    if (actions) item.appendChild(actions);
    return item;
  }

  if (hasContent) {
    item.appendChild(contentWrap);
  }
  const actions = await buildChatMessageActions(message);
  if (actions) item.appendChild(actions);

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
  driveHelpCenterFolderId = "";
  commessaSheetCache.clear();
  updateDriveStatus(false);
}

function updateDriveStatus(isConnected) {
  const connected = isConnected || localStorage.getItem("googleDriveConnected") === "true";
  ui.driveStatus.classList.toggle("status-chip-drive", connected);
  ui.driveStatus.textContent = connected ? "Drive collegato." : "Drive non collegato.";
}

async function connectGoogleDrive() {
  try {
    const provider = new firebase.auth.GoogleAuthProvider();
    provider.addScope("https://www.googleapis.com/auth/drive.file");
    const result = await firebase.auth().signInWithPopup(provider);
    const credential = result.credential || firebase.auth.GoogleAuthProvider.credentialFromResult(result);
    const accessToken = credential && credential.accessToken ? credential.accessToken : null;
    if (!accessToken) {
      throw new Error("Access token Google Drive non ottenuto");
    }

    persistDriveAccessToken(accessToken);
    await autoConnectDriveBridge({ notifyOnError: true });
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
  driveHelpCenterFolderId = await getOrCreateDriveFolder("Help Center", driveReportsFolderId);
}

function normalizeFaqData(data) {
  const payload = data && typeof data === "object" ? data : {};
  const rawItems = Array.isArray(payload.items) ? payload.items : [];
  const normalizedItems = rawItems.map((item, index) => ({
    id: String(item.id || `faq-${index + 1}`),
    domanda: String(item.domanda || item.question || "").trim(),
    risposta: String(item.risposta || item.answer || "").trim(),
    passi: Array.isArray(item.passi) ? item.passi.map((step) => String(step || "").trim()).filter(Boolean) : []
  })).filter((item) => item.domanda && item.risposta);

  return {
    version: Number(payload.version) > 0 ? Number(payload.version) : HELP_CENTER_FAQ_FALLBACK.version,
    updatedAt: payload.updatedAt || null,
    updatedBy: String(payload.updatedBy || ""),
    items: normalizedItems.length > 0 ? normalizedItems : HELP_CENTER_FAQ_FALLBACK.items
  };
}

function escapeHtml(value) {
  return String(value || "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll("\"", "&quot;")
    .replaceAll("'", "&#39;");
}

function toIsoDate(value) {
  if (!value) return "";
  if (typeof value === "string") return value;
  if (value instanceof Date) return value.toISOString();
  if (value && typeof value.toDate === "function") return value.toDate().toISOString();
  return "";
}

async function loadFaqFromFirestore() {
  try {
    const doc = await db.collection(HELP_CENTER_CONFIG_PATH.collection).doc(HELP_CENTER_CONFIG_PATH.doc).get();
    if (!doc.exists) return null;
    const data = normalizeFaqData(doc.data());
    faqDataset = data;
    return data;
  } catch (error) {
    console.warn("Help Center Firestore non disponibile, uso fallback locale:", error);
    return null;
  }
}

async function saveFaqToFirestore(faqData) {
  if (!canManageData()) {
    throw new Error("Solo un admin può aggiornare le FAQ.");
  }
  const normalized = normalizeFaqData(faqData);
  const existing = await loadFaqFromFirestore();
  const nextVersion = (existing && Number(existing.version) > 0 ? Number(existing.version) : 0) + 1;
  const payload = {
    ...normalized,
    version: nextVersion,
    updatedAt: firebase.firestore.FieldValue.serverTimestamp(),
    updatedBy: currentUser?.email || ""
  };
  await db.collection(HELP_CENTER_CONFIG_PATH.collection).doc(HELP_CENTER_CONFIG_PATH.doc).set(payload, { merge: true });
  faqDataset = { ...normalized, version: nextVersion, updatedAt: new Date().toISOString(), updatedBy: currentUser?.email || "" };
  const snapshot = await exportFaqSnapshotToDrive(faqDataset);
  return { ...faqDataset, snapshot };
}

function renderFaqHelpCenter(faqData) {
  const normalized = normalizeFaqData(faqData);
  faqDataset = normalized;
  window.appHelpFaqData = normalized;

  const list = document.getElementById("help-faq-list");
  if (!list) return;
  list.innerHTML = normalized.items.map((item) => {
    const steps = item.passi.length
      ? `<ol>${item.passi.map((step) => `<li>${escapeHtml(step)}</li>`).join("")}</ol>`
      : "";
    return `<article class="faq-item"><h3>${escapeHtml(item.domanda)}</h3><p>${escapeHtml(item.risposta)}</p>${steps}</article>`;
  }).join("");
}

async function initHelpCenterFaq() {
  const remoteFaq = await loadFaqFromFirestore();
  renderFaqHelpCenter(remoteFaq || HELP_CENTER_FAQ_FALLBACK);
}

window.loadFaqFromFirestore = loadFaqFromFirestore;
window.saveFaqToFirestore = saveFaqToFirestore;
window.exportFaqSnapshotToDrive = exportFaqSnapshotToDrive;

async function exportFaqSnapshotToDrive(faqData = faqDataset) {
  if (!canManageData()) {
    throw new Error("Solo un admin può esportare snapshot FAQ.");
  }
  if (!driveAccessToken) {
    console.warn("Drive non collegato: salto export snapshot FAQ.");
    return null;
  }

  if (!driveHelpCenterFolderId) await ensureDriveFolders();
  const normalized = normalizeFaqData(faqData);
  const metadata = {
    version: Number(normalized.version) || 1,
    updatedAt: toIsoDate(normalized.updatedAt) || new Date().toISOString(),
    updatedBy: normalized.updatedBy || currentUser?.email || ""
  };
  const payload = { ...metadata, items: normalized.items };
  const blob = new Blob([JSON.stringify(payload, null, 2)], { type: "application/json" });
  const fileName = `help-center-faq-v${metadata.version}.json`;
  const uploaded = await uploadBlobToDrive(blob, fileName, "application/json", driveHelpCenterFolderId);
  await db.collection(HELP_CENTER_CONFIG_PATH.collection).doc(HELP_CENTER_CONFIG_PATH.doc).set({
    latestSnapshot: {
      fileId: uploaded.fileId,
      url: uploaded.webViewLink || uploaded.directUrl || "",
      exportedAt: firebase.firestore.FieldValue.serverTimestamp(),
      exportedBy: metadata.updatedBy
    }
  }, { merge: true });
  return uploaded;
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

  const commessaData = commesseById.get(commessaId) || {};
  const configuredSheetId = String(commessaData.sheetSpreadsheetId || "").trim();
  if (configuredSheetId) {
    try {
      const existingSheet = await driveApiFetch(`https://www.googleapis.com/drive/v3/files/${configuredSheetId}?fields=id,name,mimeType`, { method: "GET" });
      if (existingSheet && existingSheet.id) {
        commessaSheetCache.set(commessaId, configuredSheetId);
        return { id: configuredSheetId };
      }
    } catch (error) {
      console.warn("Foglio configurato non più disponibile, provo ricreazione automatica:", error);
      await db.collection("commesse").doc(commessaId).set({
        sheetSpreadsheetId: firebase.firestore.FieldValue.delete(),
        sheetUpdatedAt: firebase.firestore.FieldValue.serverTimestamp()
      }, { merge: true });
    }
  }

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
    await db.collection("commesse").doc(commessaId).set({
      sheetSpreadsheetId: existing.id,
      sheetUpdatedAt: firebase.firestore.FieldValue.serverTimestamp()
    }, { merge: true });
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
  await db.collection("commesse").doc(commessaId).set({
    sheetSpreadsheetId: created.id,
    sheetUpdatedAt: firebase.firestore.FieldValue.serverTimestamp()
  }, { merge: true });
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

async function refreshDriveAccessToken() {
  if (driveTokenRefreshPromise) return driveTokenRefreshPromise;
  if (!auth.currentUser) return false;
  driveTokenRefreshPromise = (async () => {
    try {
      const provider = new firebase.auth.GoogleAuthProvider();
      provider.addScope("https://www.googleapis.com/auth/drive.file");
      provider.setCustomParameters({ prompt: "none" });
      const result = await auth.signInWithPopup(provider);
      const accessToken = extractGoogleAccessToken(result);
      if (!accessToken) return false;
      persistDriveAccessToken(accessToken);
      await autoConnectDriveBridge({ notifyOnError: false });
      return true;
    } catch (error) {
      console.warn("Refresh automatico token Drive non riuscito:", error);
      return false;
    } finally {
      driveTokenRefreshPromise = null;
    }
  })();
  return driveTokenRefreshPromise;
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
    const refreshed = await refreshDriveAccessToken();
    if (refreshed) {
      return driveApiFetch(url, options);
    }
    localStorage.removeItem("googleDriveAccessToken");
    updateDriveStatus(false);
    throw new Error("Sessione Drive scaduta. Premi di nuovo 'Collega Google Drive'.");
  }

  if (!response.ok) {
    const text = await response.text();
    throw new Error(`Errore Google Drive (${response.status}): ${text.slice(0, 180)}`);
  }

  if (response.status === 204) return {};
  return response.json();
}
