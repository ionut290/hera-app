firebase.initializeApp(firebaseConfig);

const auth = firebase.auth();
const db = firebase.firestore();
const functions = firebase.functions ? firebase.functions() : null;
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

const errorFeedbackAudio = {
  context: null,
  lastAt: 0
};

function triggerErrorFeedback() {
  const now = Date.now();
  if (now - errorFeedbackAudio.lastAt < 250) return;
  errorFeedbackAudio.lastAt = now;

  if (navigator?.vibrate) {
    navigator.vibrate([120, 60, 120]);
  }

  const AudioContextCtor = window.AudioContext || window.webkitAudioContext;
  if (!AudioContextCtor) return;

  try {
    if (!errorFeedbackAudio.context) {
      errorFeedbackAudio.context = new AudioContextCtor();
    }

    const context = errorFeedbackAudio.context;
    if (context.state === "suspended") {
      context.resume().catch(() => {});
    }

    const oscillator = context.createOscillator();
    const gain = context.createGain();
    const startTime = context.currentTime;

    oscillator.type = "triangle";
    oscillator.frequency.setValueAtTime(220, startTime);
    oscillator.frequency.exponentialRampToValueAtTime(140, startTime + 0.2);

    gain.gain.setValueAtTime(0.0001, startTime);
    gain.gain.exponentialRampToValueAtTime(0.12, startTime + 0.02);
    gain.gain.exponentialRampToValueAtTime(0.0001, startTime + 0.22);

    oscillator.connect(gain);
    gain.connect(context.destination);
    oscillator.start(startTime);
    oscillator.stop(startTime + 0.24);
  } catch (error) {
    console.warn("Feedback sonoro errore non disponibile:", error);
  }
}

function shouldPlayErrorFeedback(message) {
  if (typeof message !== "string") return true;
  const normalized = message.trim().toLowerCase();
  if (!normalized) return true;

  const nonErrorAlertPatterns = [
    /collegato correttamente/,
    /^import (mezzi )?completato/,
    /^richiesta inviata\./,
    /in attesa approvazione/
  ];
  if (nonErrorAlertPatterns.some((pattern) => pattern.test(normalized))) {
    return false;
  }

  return true;
}

const nativeAlert = window.alert.bind(window);
window.alert = (message) => {
  if (shouldPlayErrorFeedback(String(message || ""))) {
    triggerErrorFeedback();
  }
  nativeAlert(message);
};

const ui = {
  refreshAppBtn: document.getElementById("refresh-app-btn"),
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
  commessaCode: document.getElementById("commessa-code"),
  commessaType: document.getElementById("commessa-type"),
  commessaParent: document.getElementById("commessa-parent"),
  commesseHomeCard: document.getElementById("commesse-home-card"),
  commesseLista: document.getElementById("commesse-lista"),
  commessaAttiva: document.getElementById("commessa-attiva"),
  commesseNextAction: document.getElementById("commesse-next-action"),
  commessaTargetSelect: document.getElementById("commessa-target-select"),
  openOrganizeCommesseBtn: document.getElementById("open-organize-commesse-btn"),
  closeOrganizeCommesseBtn: document.getElementById("close-organize-commesse-btn"),
  organizeCommesseScreen: document.getElementById("organize-commesse-screen"),
  parentCommessaForm: document.getElementById("parent-commessa-form"),
  parentCommessaName: document.getElementById("parent-commessa-name"),
  parentCommessaCode: document.getElementById("parent-commessa-code"),
  moveSubcommesseForm: document.getElementById("move-subcommesse-form"),
  moveParentCommessaSelect: document.getElementById("move-parent-commessa-select"),
  moveSubcommesseList: document.getElementById("move-subcommesse-list"),
  organizeCommesseFeedback: document.getElementById("organize-commesse-feedback"),
  excelFile: document.getElementById("excel-file"),
  importBtn: document.getElementById("import-btn"),
  sheetUrl: document.getElementById("sheet-url"),
  sheetUrlImportBtn: document.getElementById("sheet-url-import-btn"),
  importFeedback: document.getElementById("import-feedback"),
  impiantiLista: document.getElementById("impianti-lista"),
  gpsStatus: document.getElementById("gps-status"),
  chatOpenBtn: document.getElementById("chat-open-btn"),
  chatCounter: document.getElementById("chat-counter"),
  chatModal: document.getElementById("chat-modal"),
  chatCloseBtn: document.getElementById("chat-close-btn"),
  chatClearBtn: document.getElementById("chat-clear-btn"),
  chatClearConfirmModal: document.getElementById("chat-clear-confirm-modal"),
  chatClearCancelBtn: document.getElementById("chat-clear-cancel-btn"),
  chatClearConfirmBtn: document.getElementById("chat-clear-confirm-btn"),
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
  impiantoWeatherDetailPage: document.getElementById("impianto-weather-detail-page"),
  impiantoWeatherDetailSubtitle: document.getElementById("impianto-weather-detail-subtitle"),
  impiantoWeatherDetailBackBtn: document.getElementById("impianto-weather-detail-back-btn"),
  impiantoWeatherDetailRefreshBtn: document.getElementById("impianto-weather-detail-refresh-btn"),
  impiantoWeatherDetailFeedback: document.getElementById("impianto-weather-detail-feedback"),
  impiantoWeatherDetailContent: document.getElementById("impianto-weather-detail-content"),
  atexProcedurePage: document.getElementById("atex-procedure-page"),
  atexProcedureBackBtn: document.getElementById("atex-procedure-back-btn"),
  atexProcedureSubtitle: document.getElementById("atex-procedure-subtitle"),
  atexProcedureContent: document.getElementById("atex-procedure-content"),
  impiantoSafetyPage: document.getElementById("impianto-safety-page"),
  impiantoSafetyBackBtn: document.getElementById("impianto-safety-back-btn"),
  impiantoSafetySubtitle: document.getElementById("impianto-safety-subtitle"),
  impiantoSafetyContent: document.getElementById("impianto-safety-content"),
  commessaFocusLabel: document.getElementById("commessa-focus-label"),
  commessaFocusCode: document.getElementById("commessa-focus-code"),
  commessaHomeBtn: document.getElementById("commessa-home-btn"),
  commessaStatImpianti: document.getElementById("commessa-stat-impianti"),
  commessaStatSegnalazioni: document.getElementById("commessa-stat-segnalazioni"),
  commessaStatAvanzamento: document.getElementById("commessa-stat-avanzamento"),
  commessaStatAvanzamentoDetail: document.getElementById("commessa-stat-avanzamento-detail"),
  commessaStatOre: document.getElementById("commessa-stat-ore"),
  commessaStatGiorni: document.getElementById("commessa-stat-giorni"),
  commessaActiveSquadreCount: document.getElementById("commessa-active-squadre-count"),
  commessaSquadreDetailsBtn: document.getElementById("commessa-squadre-details-btn"),
  backToHomeBtn: document.getElementById("back-to-home-btn"),
  showNextActionBtn: document.getElementById("show-next-action-btn"),
  impiantiNextAction: document.getElementById("impianti-next-action"),
  exportCurrentCommessaBtn: document.getElementById("export-current-commessa-btn"),
  parentCommessaOverview: document.getElementById("parent-commessa-overview"),
  parentCommessaSummary: document.getElementById("parent-commessa-summary"),
  parentSubcommesseTitle: document.getElementById("parent-subcommesse-title"),
  parentSubcommesseList: document.getElementById("parent-subcommesse-list"),
  commessaOperationalCard: document.getElementById("commessa-operational-card"),
  impiantiCard: document.getElementById("impianti-card"),
  mapFullscreenBtn: document.getElementById("map-fullscreen-btn"),
  mapInlineFullscreenBtn: document.getElementById("map-inline-fullscreen-btn"),
  operatorPositionsToggleBtn: document.getElementById("operator-positions-toggle-btn"),
  commessaNotesToggleBtn: document.getElementById("commessa-notes-toggle-btn"),
  commessaWeatherRefreshBtn: document.getElementById("commessa-weather-refresh-btn"),
  commessaWeatherRefreshStatus: document.getElementById("commessa-weather-refresh-status"),
  commessaCallBtn: document.getElementById("commessa-call-btn"),
  commessaNotesPage: document.getElementById("commessa-notes-page"),
  commessaNotesBackBtn: document.getElementById("commessa-notes-back-btn"),
  commessaNotesCard: document.getElementById("commessa-notes-card"),
  commessaNotesTitle: document.getElementById("commessa-notes-title"),
  commessaNotesCounter: document.getElementById("commessa-notes-counter"),
  commessaNoteNewBtn: document.getElementById("commessa-note-new-btn"),
  commessaNotesFormWrap: document.getElementById("commessa-notes-form-wrap"),
  commessaNoteForm: document.getElementById("commessa-note-form"),
  commessaNoteId: document.getElementById("commessa-note-id"),
  commessaNoteDate: document.getElementById("commessa-note-date"),
  commessaNoteTitle: document.getElementById("commessa-note-title"),
  commessaNoteText: document.getElementById("commessa-note-text"),
  commessaNoteDriveLinks: document.getElementById("commessa-note-drive-links"),
  commessaNoteImpiantoKey: document.getElementById("commessa-note-impianto-key"),
  commessaNoteImpiantoSearch: document.getElementById("commessa-note-impianto-search"),
  commessaNoteImpiantoClearBtn: document.getElementById("commessa-note-impianto-clear-btn"),
  commessaNoteImpiantoSuggestions: document.getElementById("commessa-note-impianto-suggestions"),
  commessaNoteImpiantoSelected: document.getElementById("commessa-note-impianto-selected"),
  commessaNoteSubmitBtn: document.getElementById("commessa-note-submit-btn"),
  commessaNoteCancelBtn: document.getElementById("commessa-note-cancel-btn"),
  commessaNotesList: document.getElementById("commessa-notes-list"),
  commessaNoteDetail: document.getElementById("commessa-note-detail"),
  mapFullscreenPage: document.getElementById("map-fullscreen-page"),
  mapFullscreenBackBtn: document.getElementById("map-fullscreen-back-btn"),
  mapSatelliteToggleBtn: document.getElementById("map-satellite-toggle-btn"),
  mapRadarToggleBtn: document.getElementById("map-radar-toggle-btn"),
  mapDrawAreaBtn: document.getElementById("map-draw-area-btn"),
  mapDrawUndoBtn: document.getElementById("map-draw-undo-btn"),
  mapDrawRedoBtn: document.getElementById("map-draw-redo-btn"),
  mapDrawClearBtn: document.getElementById("map-draw-clear-btn"),
  mapNumberSearchForm: document.getElementById("map-number-search-form"),
  mapNumberSearchInput: document.getElementById("map-number-search-input"),
  mapFullscreenNumberSearchForm: document.getElementById("map-fullscreen-number-search-form"),
  mapFullscreenNumberSearchInput: document.getElementById("map-fullscreen-number-search-input"),
  mapShareAreaWhatsappBtn: document.getElementById("map-share-area-whatsapp-btn"),
  mapFullscreenFeedbackBanner: document.getElementById("map-fullscreen-feedback-banner"),
  mapFullscreenFeedback: document.getElementById("map-fullscreen-feedback"),
  mapFullscreenFeedbackClose: document.getElementById("map-fullscreen-feedback-close"),
  mainMapImpiantoDetailPanel: document.getElementById("main-map-impianto-detail-panel"),
  mainMapImpiantoDetailBody: document.getElementById("main-map-impianto-detail-body"),
  pendingWhatsappCard: document.getElementById("pending-whatsapp-card"),
  pendingWhatsappSummary: document.getElementById("pending-whatsapp-summary"),
  pendingWhatsappBadge: document.getElementById("pending-whatsapp-badge"),
  pendingWhatsappList: document.getElementById("pending-whatsapp-list"),
  mapImpiantoDetailPanel: document.getElementById("map-impianto-detail-panel"),
  mapImpiantoDetailBody: document.getElementById("map-impianto-detail-body"),
  impiantiPageTitle: document.getElementById("impianti-page-title"),
  impiantoSearch: document.getElementById("impianto-search"),
  viewDoneBtn: document.getElementById("view-done-btn"),
  viewTodoBtn: document.getElementById("view-todo-btn"),
  viewAlertsBtn: document.getElementById("view-alerts-btn"),
  personaleForm: document.getElementById("personale-form"),
  personaleNome: document.getElementById("personale-nome"),
  personaleLista: document.getElementById("personale-lista"),
  personaleSearchInput: document.getElementById("personale-search-input"),
  personaleSearchSuggestions: document.getElementById("personale-search-suggestions"),
  personaleShowAllBtn: document.getElementById("personale-show-all-btn"),
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
  toggleCommesseHomeBtn: document.getElementById("toggle-commesse-home-btn"),
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
  openPanelGlobal: document.getElementById("open-panel-global"),
  openPanelBanner: document.getElementById("open-panel-banner"),
  openPanelInfoUtili: document.getElementById("open-panel-info-utili"),
  openPanelNotifiche: document.getElementById("open-panel-notifiche"),
  openPanelProgrammazione: document.getElementById("open-panel-programmazione"),
  openPanelBannerGestione: document.getElementById("open-panel-banner-gestione"),
  openPrivateDocsBtn: document.getElementById("open-private-docs-btn"),
  openPrivateDocsUploadBtn: document.getElementById("open-private-docs-upload-btn"),
  openPersonalServicesBtn: document.getElementById("open-personal-services-btn"),
  openHoursBtn: document.getElementById("open-hours-btn"),
  openPosBtn: document.getElementById("open-pos-btn"),
  openSegnalazioniBtn: document.getElementById("open-segnalazioni-btn"),
  openHowtoBtn: document.getElementById("open-howto-btn"),
  openBookPdfBtn: document.getElementById("open-book-pdf-btn"),
  managementPage: document.getElementById("management-page"),
  managementTitle: document.getElementById("management-title"),
  managementCloseBtn: document.getElementById("management-close-btn"),
  panelCommesse: document.getElementById("panel-commesse"),
  panelSquadre: document.getElementById("panel-squadre"),
  panelPersonale: document.getElementById("panel-personale"),
  panelMezzi: document.getElementById("panel-mezzi"),
  panelUtenti: document.getElementById("panel-utenti"),
  panelGlobal: document.getElementById("panel-global"),
  panelBanner: document.getElementById("panel-banner"),
  panelInfoUtili: document.getElementById("panel-info-utili"),
  panelNotifiche: document.getElementById("panel-notifiche"),
  panelProgrammazione: document.getElementById("panel-programmazione"),
  programmazioneAddBtn: document.getElementById("programmazione-add-btn"),
  programmazioneFilter: document.getElementById("programmazione-filter"),
  programmazioneList: document.getElementById("programmazione-list"),
  programmazioneDialog: document.getElementById("programmazione-dialog"),
  programmazioneForm: document.getElementById("programmazione-form"),
  programmazioneCancelBtn: document.getElementById("programmazione-cancel-btn"),
  programmazioneDeleteBtn: document.getElementById("programmazione-delete-btn"),
  programmaId: document.getElementById("programma-id"),
  programmaCommessa: document.getElementById("programma-commessa"),
  programmaOperatoriAutocomplete: document.getElementById("programma-operatori-autocomplete"),
  programmaMezziAutocomplete: document.getElementById("programma-mezzi-autocomplete"),
  programmazioniHomeCard: document.getElementById("programmazioni-home-card"),
  programmazioniHomeList: document.getElementById("programmazioni-home-list"),
  panelBanner: document.getElementById("panel-banner"),
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
  workBannerHome: document.getElementById("work-banner-home"),
  workBannerText: document.getElementById("work-banner-text"),
  bannerConfigForm: document.getElementById("banner-config-form"),
  bannerTextInput: document.getElementById("banner-text-input"),
  bannerNoteDate: document.getElementById("banner-note-date"),
  bannerNoteInput: document.getElementById("banner-note-input"),
  bannerAddNoteBtn: document.getElementById("banner-add-note-btn"),
  bannerNotesList: document.getElementById("banner-notes-list"),
  bannerEnabledToggle: document.getElementById("banner-enabled-toggle"),
  bannerSpeedInput: document.getElementById("banner-speed-input"),
  bannerDisableBtn: document.getElementById("banner-disable-btn"),
  bannerFeedback: document.getElementById("banner-feedback"),
  weatherRisks: document.getElementById("weather-risks"),
  userCard: document.getElementById("user-card"),
  userToggleBtn: document.getElementById("user-toggle-btn"),
  userDetailsPanel: document.getElementById("user-details-panel"),
  weatherSummary: document.getElementById("weather-summary"),
  weatherModal: document.getElementById("weather-modal"),
  weatherCloseBtn: document.getElementById("weather-close-btn"),
  weatherDetails: document.getElementById("weather-details"),
  navigationWeatherWarningModal: document.getElementById("navigation-weather-warning-modal"),
  navigationWeatherWarningList: document.getElementById("navigation-weather-warning-list"),
  navigationWeatherContinueBtn: document.getElementById("navigation-weather-continue-btn"),
  navigationWeatherCancelBtn: document.getElementById("navigation-weather-cancel-btn"),
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
  posPage: document.getElementById("pos-page"),
  backFromPosBtn: document.getElementById("back-from-pos-btn"),
  posAdminCard: document.getElementById("pos-admin-card"),
  posAddToggleBtn: document.getElementById("pos-add-toggle-btn"),
  posDocumentForm: document.getElementById("pos-document-form"),
  posDocumentId: document.getElementById("pos-document-id"),
  posTitle: document.getElementById("pos-title"),
  posDescription: document.getElementById("pos-description"),
  posDriveUrl: document.getElementById("pos-drive-url"),
  posCategory: document.getElementById("pos-category"),
  posOrder: document.getElementById("pos-order"),
  posActive: document.getElementById("pos-active"),
  posSaveBtn: document.getElementById("pos-save-btn"),
  posCancelBtn: document.getElementById("pos-cancel-btn"),
  posFeedback: document.getElementById("pos-feedback"),
  posSearch: document.getElementById("pos-search"),
  posDocumentsList: document.getElementById("pos-documents-list"),
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
  hoursTableCommesseButtons: document.getElementById("hours-table-commesse-buttons"),
  hoursTotalOperatorBtn: document.getElementById("hours-total-operator-btn"),
  hoursTotalOperatorCommessaBtn: document.getElementById("hours-total-operator-commessa-btn"),
  hoursTableExportBtn: document.getElementById("hours-table-export-btn"),
  hoursTableExportGlobalBtn: document.getElementById("hours-table-export-global-btn"),
  hoursTableFeedback: document.getElementById("hours-table-feedback"),
  hoursConfirmVisibleBtn: document.getElementById("hours-confirm-visible-btn"),
  hoursTableContainer: document.getElementById("hours-table-container"),
  hoursConfirmModal: document.getElementById("hours-confirm-modal"),
  hoursConfirmTitle: document.getElementById("hours-confirm-title"),
  hoursConfirmText: document.getElementById("hours-confirm-text"),
  hoursConfirmCancelBtn: document.getElementById("hours-confirm-cancel-btn"),
  hoursConfirmOkBtn: document.getElementById("hours-confirm-ok-btn"),
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
  editGpsX: document.getElementById("edit-gps-x"),
  globalCommessaForm: document.getElementById("global-commessa-form"),
  globalCommessaName: document.getElementById("global-commessa-name"),
  globalCommesseLista: document.getElementById("global-commesse-lista"),
  globalCommessaSelect: document.getElementById("global-commessa-select"),
  globalExcelFile: document.getElementById("global-excel-file"),
  globalImportBtn: document.getElementById("global-import-btn"),
  globalUpdateBtn: document.getElementById("global-update-btn"),
  globalSheetUrl: document.getElementById("global-sheet-url"),
  globalSheetUrlImportBtn: document.getElementById("global-sheet-url-import-btn"),
  globalImportFeedback: document.getElementById("global-import-feedback"),
  globalImpiantoSearchForm: document.getElementById("global-impianto-search-form"),
  globalImpiantoSearchBtn: document.getElementById("global-impianto-search-btn"),
  globalImpiantoSearch: document.getElementById("global-impianto-search"),
  globalOpenReportBtn: document.getElementById("global-open-report-btn"),
  globalImpiantiLista: document.getElementById("global-impianti-lista"),
  globalMapFeedback: document.getElementById("global-map-feedback"),
  globalImpiantoDetails: document.getElementById("global-impianto-details"),
  globalImpiantoDetailsBody: document.getElementById("global-impianto-details-body"),
  globalImpiantoDetailsCloseBtn: document.getElementById("global-impianto-details-close-btn"),
  globalImpiantoNavigateBtn: document.getElementById("global-impianto-navigate-btn"),
  globalImpiantoWhatsappBtn: document.getElementById("global-impianto-whatsapp-btn"),
  globalReportModal: document.getElementById("global-report-modal"),
  globalReportCloseBtn: document.getElementById("global-report-close-btn"),
  globalReportForm: document.getElementById("global-report-form"),
  globalReportImpiantoSelect: document.getElementById("global-report-impianto-select"),
  globalReportIdSap: document.getElementById("global-report-id-sap"),
  globalReportDenominazione: document.getElementById("global-report-denominazione"),
  globalReportComune: document.getElementById("global-report-comune"),
  globalReportVia: document.getElementById("global-report-via"),
  globalReportCoordinate: document.getElementById("global-report-coordinate"),
  globalReportDitta: document.getElementById("global-report-ditta"),
  globalReportText: document.getElementById("global-report-text"),
  globalReportFeedback: document.getElementById("global-report-feedback"),
  notificationForm: document.getElementById("notification-form"),
  notificationTitle: document.getElementById("notification-title"),
  notificationDate: document.getElementById("notification-date"),
  notificationSendAllToggle: document.getElementById("notification-send-all-toggle"),
  notificationUserSelect: document.getElementById("notification-user-select"),
  notificationMessage: document.getElementById("notification-message"),
  notificationAttachments: document.getElementById("notification-attachments"),
  notificationSubmit: document.getElementById("notification-submit"),
  notificationCancelUploadBtn: document.getElementById("notification-cancel-upload-btn"),
  notificationFeedback: document.getElementById("notification-feedback"),
  notificationsList: document.getElementById("notifications-list"),
  notificationMainView: document.getElementById("notification-main-view"),
  notificationOpenCalendarBtn: document.getElementById("notification-open-calendar-btn"),
  notificationCalendarView: document.getElementById("notification-calendar-view"),
  notificationCalendarBackBtn: document.getElementById("notification-calendar-back-btn"),
  notificationCalendarPrevBtn: document.getElementById("notification-calendar-prev-btn"),
  notificationCalendarNextBtn: document.getElementById("notification-calendar-next-btn"),
  notificationCalendarMonthLabel: document.getElementById("notification-calendar-month-label"),
  notificationCalendarGrid: document.getElementById("notification-calendar-grid"),
  notificationDayDetail: document.getElementById("notification-day-detail"),
  userAlertModal: document.getElementById("user-alert-modal"),
  userAlertText: document.getElementById("user-alert-text"),
  userAlertAttachments: document.getElementById("user-alert-attachments"),
  userAlertOkBtn: document.getElementById("user-alert-ok-btn"),
  userAlertLaterBtn: document.getElementById("user-alert-later-btn"),
  notificationDocViewerModal: document.getElementById("notification-doc-viewer-modal"),
  notificationDocViewerTitle: document.getElementById("notification-doc-viewer-title"),
  notificationDocViewerCloseBtn: document.getElementById("notification-doc-viewer-close-btn"),
  notificationDocViewerFrame: document.getElementById("notification-doc-viewer-frame")
};

let pendingRows = [];
let selectedCommessaId = "";
let selectedCommessaName = "";
let unsubscribeCommesse = null;
let unsubscribeImpianti = null;
let unsubscribeCommessaNotes = null;
const unsubscribeCommessaStats = new Map();
let unsubscribeHoursStats = null;
let unsubscribeHoursApprovals = null;
let currentUserPos = null;
let currentWeatherTarget = { lat: 44.4949, lon: 11.3426 };
let currentCivilProtectionAlert = { level: "green", label: "Nessuna allerta", url: "" };
let currentImpianti = [];
let currentCommessaNotes = [];
let commessaNoteImpiantoSearchTerm = "";
let currentUser = null;
let unsubscribeChat = null;
let unsubscribeDriveBridge = null;
let driveBridgeState = { configured: false, ownerEmail: "", rootFolderId: "" };
let unsubscribePersonale = null;
let unsubscribeMezzi = null;
let unsubscribeSquadre = null;
let unsubscribeSquadreHistory = null;
let unsubscribeSquadreViewConfig = null;
let unsubscribeUsers = null;
let unsubscribeOperatorPositions = null;
let unsubscribeAdminUsers = null;
let unsubscribeResources = null;
let unsubscribePrivateDocs = null;
let unsubscribeGpsRequests = null;
let unsubscribeGlobalNotifications = null;
let unsubscribeWorkBanner = null;
let unsubscribeUserAlerts = null;
let currentWorkBannerConfig = { text: "", enabled: false, speed: null, notes: [] };
let workBannerResizeObserver = null;
let presenceHeartbeatTimer = null;
let chatMessages = [];
let chatNotificationsInitialized = false;
let platformUsers = [];
let programmazioni = [];
let programmazioneOperatorAutocomplete = null;
let programmazioneMezziAutocomplete = null;
let unsubscribeProgrammazioni = null;
let operatorPositions = [];
let operatorPositionsVisible = true;
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
let commesseLoadState = { status: "idle", message: "" };
let isCommesseHomeCardVisible = false;
let commessaStatsById = new Map();
let commessaHoursById = new Map();
let commessaWorkSummariesById = new Map();
let allHoursReports = [];
let allHoursApprovalRequests = [];
let hoursReportsLoaded = false;
let hoursApprovalsLoaded = false;
let personaleRecords = [];
let personaleSearchQuery = "";
let personaleExpandedId = "";
let personaleShowAll = false;
const PERSONALE_RECENT_KEY = "hera_personale_recent_ids";
let mezziRecords = [];
let squadreByCommessa = new Map();
let squadreHistoryByDate = new Map();
let squadreLoadState = { status: "idle", message: "" };
let manualSquadreFilterDateKey = "";
let sharedSquadreDateKey = "";
let automaticSquadreDateKey = "";
let startupAssignedCommessaAutoOpenDone = false;
let sharedSquadreViewConfigLoaded = false;
let highlightedImpiantoKey = "";
let expandedImpiantoKey = "";
const expandedImpiantoManagementKeys = new Set();
let impiantiSearchTerm = "";
let impiantiViewMode = "todo";
const whazzupSafetyByImpianto = new Map();
const WHAZZUP_PENDING_DONE_KEY = "heraWhazzupPendingDone";
let pendingSheetExports = [];
let pendingImpiantoActions = [];
let pendingWhatsappAlertShownForSyncIds = new Set();
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
let hoursConfirmModalResolve = null;
let loadingOre = false;
let hoursTableLoadPromise = null;
let hoursTableLoadRequestId = 0;
let hoursSubmitInFlight = false;
let hoursFinalizeStatusTimer = null;
let hoursDuplicateCleanupPromise = null;
let hoursApprovalRequests = [];
let gpsUpdateRequests = [];
let activeResourceTypeForViewer = "";
let activeResourceManageFilter = "";
let editingImpiantoIds = [];
let reportingImpianto = null;
let chatRetentionTimer = null;
let hoursDeadlineAlertTimer = null;
let quickSquadraWindowTimer = null;
let geolocationWatchId = null;
let lastPositionPublishAt = 0;
let lastPublishedUserPos = null;
let latestGeolocationCoords = null;
let radarPaneInitialized = false;
let radarActive = false;
let radarFrames = [];
let radarFrameIndex = 0;
let radarLayer = null;
let radarControlsEl = null;
let weatherLegendEl = null;
let weatherLayerSelectorEl = null;
let radarPlayTimer = null;
let radarPlaying = true;
let radarLoading = false;
let activeWeatherLayerId = "rain";
let weatherFramesBySource = {};
let weatherLayerLoadToken = 0;
let activeNearbyImpiantoContext = null;
let globalNotificationsInitialized = false;
let unsubscribeGlobalCommesse = null;
let unsubscribeGlobalImpianti = null;
let pendingGlobalRows = [];
let selectedGlobalCommessaId = "";
let globalCommesseById = new Map();
let globalImpianti = [];
let globalImpiantoSearchTerm = "";
let selectedGlobalImpiantoKey = "";
let selectedGlobalImpianto = null;
let selectedGlobalSegnalazioneKey = "";
let mainMapViewState = { center: [44.4949, 11.3426], zoom: 11, hasUserMoved: false };
let globalMapViewState = { center: [44.4949, 11.3426], zoom: 6, hasUserMoved: false };
let isMapFullscreenPageOpen = false;
let fullscreenMapMode = "standard";
let selectedFullscreenImpiantoId = "";
let selectedImpiantoId = "";
let selectedImpiantoData = null;
const impiantoWeatherStatusCache = new Map();
const impiantoWeatherCoordinateCache = new Map();
const impiantoWeatherPendingKeys = new Set();
const impiantoWeatherFeedbackByKey = new Map();
let impiantoWeatherPersistentCacheLoaded = false;
let impiantoWeatherRefreshTimer = null;
let impiantoWeatherRenderTimer = null;
let commessaWeatherManualRefreshInFlight = false;
let drawAreaModeActive = false;
let drawnAreaPoints = [];
let drawnAreaRedoStack = [];
let isDrawingStrokeActive = false;
let globalImpiantiFiltered = [];
let userAlerts = [];
let activeUserAlert = null;
let notificationUploadAbortController = null;
let notificationUploadInProgress = false;
let notificationCalendarCursor = new Date();
let selectedNotificationCalendarDateKey = "";
const impiantoMarkerByKey = new Map();
const fullscreenImpiantoMarkerByKey = new Map();
const whazzupProcessingByImpianto = new Set();
let mapMarkerSequenceByKey = new Map();
const CHAT_RETENTION_MS = 24 * 60 * 60 * 1000;
const HOURS_DEADLINE_ALERT_RETENTION_MS = 7 * 24 * 60 * 60 * 1000;
const HOURS_DEADLINE_ALERT_HOUR = 19;
const NETWORK_DEFAULT_TIMEOUT_MS = 12000;
const NETWORK_RETRYABLE_STATUS_CODES = new Set([408, 425, 429, 500, 502, 503, 504]);
const PROXIMITY_NEAR_KM = 0.25;
const PROXIMITY_AWAY_KM = 0.7;
const TIMBRATURA_TARGET_LAT = 44.4949;
const TIMBRATURA_TARGET_LNG = 11.3426;
const TIMBRATURA_RADIUS_M = 200;
const TIMBRATURA_ENTRATA_START_MIN = 6 * 60 + 15;
const TIMBRATURA_ENTRATA_END_MIN = 7 * 60 + 30;
const TIMBRATURA_USCITA_START_MIN = 15 * 60 + 30;
const TIMBRATURA_USCITA_END_MIN = 17 * 60;
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
  "open-pos-btn": {
    rispostaBreve: "Archivio documenti sicurezza: POS, PMS, schede lavorazioni, schede macchine e modulistica.",
    passi: [
      "Apri il menu (⋮) e premi “📄 POS”.",
      "Cerca il documento per titolo, descrizione o categoria.",
      "Premi “Apri documento” per consultare il link Google Drive in una nuova scheda."
    ],
    tags: ["pos", "documenti", "sicurezza", "drive"]
  },
  "open-segnalazioni-btn": {
    rispostaBreve: "Compili la segnalazione sicurezza e generi il PDF da condividere.",
    passi: [
      "Apri il menu (⋮) e premi “Segnalazioni”.",
      "Compila i campi obbligatori e scegli il tipo di segnalazione.",
      "Genera il PDF e condividilo via WhatsApp o email."
    ],
    tags: ["segnalazioni", "pdf", "sicurezza"]
  },
  "open-book-pdf-btn": {
    rispostaBreve: "Apre il manuale completo dell'app in formato PDF in una nuova scheda.",
    passi: [
      "Apri il menu (⋮) e premi “Libro PDF”.",
      "Attendi l'apertura del file PDF in una nuova scheda del browser.",
      "Se il popup è bloccato, abilita i popup oppure scarica il file dal link diretto."
    ],
    tags: ["manuale", "pdf", "guida"]
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
    rispostaBreve: "Solo l’amministratore collega Google Drive; per gli utenti il cloud è automatico.",
    passi: [
      "Esegui login con Google con un account autorizzato.",
      "Solo l’amministratore apre il pannello utente e preme “Collega Google Drive”.",
      "Gli utenti normali non devono autenticare Drive: vedono solo lo stato del cloud centralizzato."
    ],
    tags: ["drive", "google", "integrazione"],
    updatedAt: HOWTO_UPDATED_AT
  }
];

const DRIVE_CHAT_MEDIA_MAX_MB = 512;
const CENTRAL_DRIVE_ROOT_FOLDER_ID = "1s6qmv2SsiTUbCjqFX4yIk4VoPQayFrU0";
const CENTRAL_DRIVE_ROOT_FOLDER_NAME = "Varga Cantieri";
const CENTRAL_DRIVE_DEFAULT_COMMESSA = "Generale";
const CENTRAL_DRIVE_LEGACY_FOLDER_NAME = "VECCHI DATI";
const LEGACY_DRIVE_ROOT_FOLDER_NAMES = ["Hera App - Dati"];
const LEGACY_DRIVE_MIGRATION_KEY = "heraLegacyDriveMigrationDone";
const ADMIN_EMAIL = "ionut29019@gmail.com";
const POS_DEFAULT_CATEGORIES = ["POS", "PMS", "Schede lavorazioni", "Schede macchine e attrezzature", "Sicurezza", "Modulistica", "Altro"];
const IMPIANTO_ACTIONS = ["done", "navigate", "reset", "whatsapp", "problem-report", "gps-update", "edit", "delete"];
let adminEmails = new Set([ADMIN_EMAIL]);
let posDocuments = [];
let unsubscribePosDocuments = null;
const PENDING_SHEET_EXPORTS_KEY = "heraPendingSheetExports";
const PENDING_IMPIANTO_ACTIONS_KEY = "heraPendingImpiantoActions";
const COMMESSE_LOCAL_CACHE_KEY = "heraCommesseCache";
const LAST_SELECTED_COMMESSA_KEY = "heraLastSelectedCommessaId";
const LAST_OPENED_COMMESSA_KEY = "heraLastOpenedCommessaId";
const USER_WORKFLOW_STEP_KEY = "heraUserWorkflowStep";
const IMPIANTO_WEATHER_LOCAL_CACHE_KEY = "heraImpiantoWeatherCache:v1";
const SHEET_RETRY_MS = 30 * 1000;
const HELP_CENTER_CONFIG_PATH = { collection: "appConfig", doc: "helpCenter" };
const WORK_BANNER_CONFIG_PATH = { collection: "appConfig", doc: "workBanner" };
const WORK_BANNER_DEFAULT_DURATION_SEC = 35;
const WORK_BANNER_NEXT_NOTE_PREVIEW_HOUR = 15;
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
      risposta: "Solo l'admin deve collegare Google Drive: il cloud centralizzato viene poi usato automaticamente da tutti gli utenti loggati.",
      passi: ["Login admin", "Premi Collega Google Drive", "Concedi autorizzazioni", "Verifica Cloud centralizzato attivo"]
    }
  ]
};
let faqDataset = HELP_CENTER_FAQ_FALLBACK;
let currentWorkflowStepId = localStorage.getItem(USER_WORKFLOW_STEP_KEY) || "";
let impiantoNextActionIndex = 0;
let impiantoNextActionHighlightEnabled = false;
window.googleDriveAccessToken = localStorage.getItem("googleDriveAccessToken") || null;
driveAccessToken = "";

const STANDARD_TILE_URL = "https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png";
const STANDARD_TILE_OPTIONS = {
  maxNativeZoom: 19,
  maxZoom: 20,
  updateWhenIdle: true,
  keepBuffer: 2,
  attribution: "&copy; OpenStreetMap"
};
const SATELLITE_TILE_URL = "https://server.arcgisonline.com/ArcGIS/rest/services/World_Imagery/MapServer/tile/{z}/{y}/{x}";
const SATELLITE_TILE_OPTIONS = {
  maxNativeZoom: 19,
  maxZoom: 20,
  updateWhenIdle: true,
  keepBuffer: 2,
  attribution: "Tiles &copy; Esri"
};
const HYBRID_LABEL_TILE_URL = "https://services.arcgisonline.com/ArcGIS/rest/services/Reference/World_Boundaries_and_Places/MapServer/tile/{z}/{y}/{x}";
const HYBRID_LABEL_TILE_OPTIONS = {
  maxNativeZoom: 19,
  maxZoom: 20,
  updateWhenIdle: true,
  keepBuffer: 2,
  attribution: "Labels &copy; Esri"
};

const MAP_INTERACTION_OPTIONS = {
  markerZoomAnimation: true,
  zoomSnap: 0.25,
  zoomDelta: 0.5,
  wheelPxPerZoomLevel: 96,
  maxZoom: 20
};

const map = L.map("map", MAP_INTERACTION_OPTIONS);
L.tileLayer(STANDARD_TILE_URL, STANDARD_TILE_OPTIONS).addTo(map);
map.setView(mainMapViewState.center, mainMapViewState.zoom);
map.doubleClickZoom.disable();

const markerLayer = L.layerGroup().addTo(map);
const fullscreenMap = L.map("map-fullscreen-view", {
  ...MAP_INTERACTION_OPTIONS,
  closePopupOnClick: false,
  zoomAnimationThreshold: 4,
  inertia: true,
  worldCopyJump: false
});
const fullscreenStandardTileLayer = L.tileLayer(STANDARD_TILE_URL, STANDARD_TILE_OPTIONS).addTo(fullscreenMap);
const fullscreenSatelliteTileLayer = L.tileLayer(SATELLITE_TILE_URL, SATELLITE_TILE_OPTIONS);
const fullscreenHybridTileLayer = L.layerGroup([
  L.tileLayer(SATELLITE_TILE_URL, SATELLITE_TILE_OPTIONS),
  L.tileLayer(HYBRID_LABEL_TILE_URL, HYBRID_LABEL_TILE_OPTIONS)
]);
fullscreenMap.setView(mainMapViewState.center, mainMapViewState.zoom);
const fullscreenMarkerLayer = L.layerGroup().addTo(fullscreenMap);
const fullscreenDrawLayer = L.layerGroup().addTo(fullscreenMap);
const fullscreenBaseLayers = {
  "Mappa standard": fullscreenStandardTileLayer,
  "Satellite": fullscreenSatelliteTileLayer,
  "Ibrida": fullscreenHybridTileLayer
};
L.control.layers(fullscreenBaseLayers, null, { position: "topright" }).addTo(fullscreenMap);

const globalMap = L.map("global-map", MAP_INTERACTION_OPTIONS);
L.tileLayer(STANDARD_TILE_URL, { ...STANDARD_TILE_OPTIONS, attribution: "&copy; OpenStreetMap contributors" }).addTo(globalMap);
globalMap.setView(globalMapViewState.center, globalMapViewState.zoom);
const globalMarkerLayer = L.layerGroup().addTo(globalMap);
map.on("moveend zoomend", () => {
  const center = map.getCenter();
  mainMapViewState = {
    center: [center.lat, center.lng],
    zoom: map.getZoom(),
    hasUserMoved: true
  };
  if (isMapFullscreenPageOpen && !drawAreaModeActive) {
    fullscreenMap.setView(mainMapViewState.center, mainMapViewState.zoom, { animate: false });
  }
});
fullscreenMap.on("moveend zoomend", () => {
  if (!isMapFullscreenPageOpen || drawAreaModeActive) return;
  const center = fullscreenMap.getCenter();
  mainMapViewState = {
    center: [center.lat, center.lng],
    zoom: fullscreenMap.getZoom(),
    hasUserMoved: true
  };
  map.setView(mainMapViewState.center, mainMapViewState.zoom, { animate: false });
});
globalMap.on("moveend zoomend", () => {
  const center = globalMap.getCenter();
  globalMapViewState = {
    center: [center.lat, center.lng],
    zoom: globalMap.getZoom(),
    hasUserMoved: true
  };
});
fullscreenMap.on("baselayerchange", (event) => {
  const layerName = String(event.name || "").toLowerCase();
  if (layerName.includes("satellite")) fullscreenMapMode = "satellite";
  else if (layerName.includes("ibrida")) fullscreenMapMode = "hybrid";
  else fullscreenMapMode = "standard";
  updateFullscreenMapModeButton();
});
const fullscreenMapContainer = fullscreenMap.getContainer();
fullscreenMapContainer.addEventListener("pointerdown", onFullscreenMapPointerDown);
fullscreenMapContainer.addEventListener("pointermove", onFullscreenMapPointerMove);
fullscreenMapContainer.addEventListener("pointerup", onFullscreenMapPointerUp);
fullscreenMapContainer.addEventListener("pointercancel", onFullscreenMapPointerUp);
window.addEventListener("resize", () => {
  if (isMapFullscreenPageOpen) refreshFullscreenMapLayout();
  updateWorkBannerAnimationDuration();
});

ui.loginBtn.addEventListener("click", loginWithGoogle);
ui.switchAccountBtn.addEventListener("click", switchGoogleAccount);
ui.refreshAppBtn?.addEventListener("click", refreshApplicationData);
ui.menuToggleBtn.addEventListener("click", openSideMenu);
ui.menuCloseBtn.addEventListener("click", closeSideMenu);
ui.menuOverlay.addEventListener("click", closeSideMenu);
ui.logoutBtn.addEventListener("click", logout);
ui.driveConnectBtn.addEventListener("click", connectGoogleDrive);
ui.commessaForm.addEventListener("submit", createCommessa);
ui.commessaType?.addEventListener("change", updateCommessaParentField);
ui.openOrganizeCommesseBtn?.addEventListener("click", () => toggleOrganizeCommesseScreen(true));
ui.closeOrganizeCommesseBtn?.addEventListener("click", () => toggleOrganizeCommesseScreen(false));
ui.parentCommessaForm?.addEventListener("submit", createParentCommessa);
ui.moveSubcommesseForm?.addEventListener("submit", moveSelectedCommesseUnderParent);
ui.moveParentCommessaSelect?.addEventListener("change", renderMoveSubcommesseList);
ui.excelFile.addEventListener("change", onExcelSelected);
ui.importBtn.addEventListener("click", importPendingRows);
ui.sheetUrlImportBtn?.addEventListener("click", importFromGoogleSheetUrl);
ui.commessaTargetSelect.addEventListener("change", onCommessaTargetChanged);
ui.chatOpenBtn.addEventListener("click", openChatModal);
ui.chatCloseBtn.addEventListener("click", closeChatModal);
ui.chatClearBtn?.addEventListener("click", openChatClearConfirmModal);
ui.chatClearCancelBtn?.addEventListener("click", closeChatClearConfirmModal);
ui.chatClearConfirmBtn?.addEventListener("click", clearCurrentChatMessages);
ui.chatClearConfirmModal?.addEventListener("click", (event) => {
  if (event.target === ui.chatClearConfirmModal) closeChatClearConfirmModal();
});
ui.chatSendForm.addEventListener("submit", sendTextMessage);
ui.chatMediaInput.addEventListener("change", sendMediaMessage);
ui.chatVoiceBtn.addEventListener("click", toggleVoiceRecording);
ui.backToHomeBtn.addEventListener("click", closeImpiantiPage);
ui.impiantoWeatherDetailBackBtn?.addEventListener("click", closeDettaglioMeteoImpianto);
ui.impiantoWeatherDetailRefreshBtn?.addEventListener("click", refreshDettaglioMeteoImpianto);
ui.atexProcedureBackBtn?.addEventListener("click", closeAtexProcedurePage);
ui.atexProcedureContent?.addEventListener("click", handleAtexProcedureContentClick);
ui.atexProcedureContent?.addEventListener("submit", saveAtexProcedureForm);
ui.impiantoSafetyBackBtn?.addEventListener("click", closeImpiantoSafetyPage);
ui.impiantoSafetyContent?.addEventListener("click", handleImpiantoSafetyContentClick);
ui.impiantoSafetyContent?.addEventListener("submit", saveImpiantoSafetyContactForm);
ui.commessaHomeBtn?.addEventListener("click", closeImpiantiPage);
ui.showNextActionBtn?.addEventListener("click", toggleImpiantoNextActionHighlight);
ui.exportCurrentCommessaBtn.addEventListener("click", () => exportCommessaSummary(selectedCommessaId, selectedCommessaName));
ui.mapFullscreenBtn.addEventListener("click", openMapFullscreenPage);
ui.mapInlineFullscreenBtn?.addEventListener("click", openMapFullscreenPage);
ui.mapNumberSearchForm?.addEventListener("submit", (event) => { event.preventDefault(); focusImpiantoByMapNumber(ui.mapNumberSearchInput?.value, map); });
ui.mapFullscreenNumberSearchForm?.addEventListener("submit", (event) => { event.preventDefault(); focusImpiantoByMapNumber(ui.mapFullscreenNumberSearchInput?.value, fullscreenMap); });
ui.operatorPositionsToggleBtn?.addEventListener("click", toggleOperatorPositionsVisibility);
ui.commessaNotesToggleBtn?.addEventListener("click", openCommessaNotesPage);
ui.commessaWeatherRefreshBtn?.addEventListener("click", refreshSelectedCommessaWeather);
document.addEventListener("click", handleImpiantoWeatherRetryClick);
document.addEventListener("click", handleAtexProcedureButtonClick);
document.addEventListener("click", handleImpiantoSafetyButtonClick);
ui.commessaCallBtn?.addEventListener("click", openCommessaPhoneResources);
ui.commessaSquadreDetailsBtn?.addEventListener("click", scrollToHomeSquadreSection);
ui.commessaNotesBackBtn?.addEventListener("click", openImpiantiPage);
ui.commessaNoteNewBtn?.addEventListener("click", () => openCommessaNoteForm());
ui.commessaNoteForm?.addEventListener("submit", saveCommessaNote);
ui.commessaNoteCancelBtn?.addEventListener("click", closeCommessaNoteForm);
ui.commessaNoteImpiantoSearch?.addEventListener("input", onCommessaNoteImpiantoSearchInput);
ui.commessaNoteImpiantoSearch?.addEventListener("focus", () => renderCommessaNoteImpiantoSuggestions());
ui.commessaNoteImpiantoSearch?.addEventListener("blur", () => setTimeout(() => {
  ui.commessaNoteImpiantoSuggestions?.classList.add("hidden");
  ui.commessaNoteImpiantoSearch?.setAttribute("aria-expanded", "false");
}, 120));
ui.commessaNoteImpiantoClearBtn?.addEventListener("click", clearCommessaNoteImpiantoSelection);
ui.mapFullscreenBackBtn?.addEventListener("click", closeMapFullscreenPage);
ui.mapSatelliteToggleBtn?.addEventListener("click", toggleFullscreenSatelliteMode);
ui.mapRadarToggleBtn?.addEventListener("click", toggleWeatherRadar);
ui.mapDrawAreaBtn?.addEventListener("click", toggleDrawAreaMode);
ui.mapDrawUndoBtn?.addEventListener("click", undoDrawnArea);
ui.mapDrawRedoBtn?.addEventListener("click", redoDrawnArea);
ui.mapDrawClearBtn?.addEventListener("click", clearDrawnArea);
ui.mapShareAreaWhatsappBtn?.addEventListener("click", shareDrawnAreaViaWhatsapp);
ui.mapFullscreenFeedbackClose?.addEventListener("click", () => ui.mapFullscreenFeedbackBanner?.classList.add("hidden"));
ui.toggleCommesseHomeBtn?.addEventListener("click", toggleCommesseHomeCard);
ui.impiantoSearch.addEventListener("input", onImpiantoSearchInput);
ui.viewDoneBtn.addEventListener("click", () => setImpiantiViewMode("done"));
ui.viewTodoBtn.addEventListener("click", () => setImpiantiViewMode("todo"));
ui.viewAlertsBtn?.addEventListener("click", () => setImpiantiViewMode("alerts"));
document.querySelectorAll(".commessa-stat-item[data-stat-action]").forEach((item) => {
  item.addEventListener("click", () => handleCommessaStatAction(item.dataset.statAction || ""));
  item.addEventListener("keydown", (event) => {
    if (event.key !== "Enter" && event.key !== " ") return;
    event.preventDefault();
    handleCommessaStatAction(item.dataset.statAction || "");
  });
});
ui.personaleForm.addEventListener("submit", addPersonale);
ui.personaleSearchInput?.addEventListener("input", (event) => {
  personaleSearchQuery = String(event.target.value || "");
  personaleShowAll = false;
  renderPersonaleList(ui.personaleLista, personaleRecords, deletePersonale);
});
ui.personaleShowAllBtn?.addEventListener("click", () => {
  personaleShowAll = !personaleShowAll;
  personaleExpandedId = personaleShowAll ? "" : personaleExpandedId;
  ui.personaleShowAllBtn.textContent = personaleShowAll ? "Nascondi elenco" : "Mostra tutto personale";
  renderPersonaleList(ui.personaleLista, personaleRecords, deletePersonale);
});
ui.mezziForm.addEventListener("submit", addMezzo);
ui.squadraForm.addEventListener("submit", saveSquadraComposition);
ui.squadraCommessa.addEventListener("change", autofillSquadraForm);
ui.squadraCalendarDate.addEventListener("change", () => {
  setSquadreDateOverride(ui.squadraCalendarDate.value || "");
});
ui.squadreFilterDate?.addEventListener("change", onSquadreFilterDateChange);
ui.squadreFilterClearBtn?.addEventListener("click", clearManualSquadreFilterDate);

function syncCommesseHomeToggle() {
  const isVisible = Boolean(isCommesseHomeCardVisible);
  ui.commesseHomeCard?.classList.toggle("hidden", !isVisible);
  ui.commesseHomeCard?.setAttribute("aria-hidden", isVisible ? "false" : "true");
  if (ui.toggleCommesseHomeBtn) {
    ui.toggleCommesseHomeBtn.setAttribute("aria-expanded", isVisible ? "true" : "false");
    ui.toggleCommesseHomeBtn.classList.toggle("active", isVisible);
    ui.toggleCommesseHomeBtn.textContent = isVisible ? "Nascondi" : "Tutte le commesse";
    ui.toggleCommesseHomeBtn.setAttribute(
      "aria-label",
      isVisible ? "Nascondi elenco completo commesse" : "Mostra elenco completo commesse"
    );
  }
}

function showCommesseHomeCard() {
  if (ui.homePage?.classList.contains("hidden")) {
    if (isMapFullscreenPageOpen) closeMapFullscreenPage();
    window.location.hash = "";
    applyRoute();
  }
  isCommesseHomeCardVisible = true;
  syncCommesseHomeToggle();
  renderCommesseHomeList();
  ui.commesseHomeCard?.scrollIntoView({ behavior: "smooth", block: "start" });
}

function toggleCommesseHomeCard() {
  if (!isCommesseHomeCardVisible) {
    showCommesseHomeCard();
    return;
  }
  isCommesseHomeCardVisible = false;
  syncCommesseHomeToggle();
}
ui.addSquadraRowBtn.addEventListener("click", () => addSquadraRow());
ui.personaleImportBtn.addEventListener("click", importPersonaleFromExcel);
ui.mezziImportBtn.addEventListener("click", importMezziFromExcel);
ui.openPanelCommesse.addEventListener("click", () => openManagementPanel("commesse"));
ui.openPanelSquadre.addEventListener("click", () => openManagementPanel("squadre"));
ui.openPanelPersonale.addEventListener("click", () => openManagementPanel("personale"));
ui.openPanelMezzi.addEventListener("click", () => openManagementPanel("mezzi"));
ui.openPanelUtenti.addEventListener("click", () => openManagementPanel("utenti"));
ui.openPanelGlobal.addEventListener("click", () => openManagementPanel("global"));
ui.openPanelBanner.addEventListener("click", () => openManagementPanel("banner"));
ui.openPanelInfoUtili.addEventListener("click", () => openManagementPanel("infoUtili"));
ui.openPanelNotifiche?.addEventListener("click", () => openManagementPanel("notifiche"));
ui.openPanelProgrammazione?.addEventListener("click", () => openManagementPanel("programmazione"));
ui.programmazioneAddBtn?.addEventListener("click", () => {
  if (!canManageData()) return;
  ui.programmaId.value = "";
  ui.programmazioneDeleteBtn?.classList.add("hidden");
  ui.programmazioneForm?.reset();
  populateProgrammazioneFormOptions();
  programmazioneOperatorAutocomplete = buildProgrammazioneAutocomplete(ui.programmaOperatoriAutocomplete, "Operatori coinvolti", getProgrammazioneOperatorOptions(), []);
  programmazioneMezziAutocomplete = buildProgrammazioneAutocomplete(ui.programmaMezziAutocomplete, "Mezzi / attrezzature", getProgrammazioneMezziOptions(), []);
  ui.programmazioneDialog?.showModal();
});
ui.programmazioneCancelBtn?.addEventListener("click", () => ui.programmazioneDialog?.close());
ui.programmazioneDeleteBtn?.addEventListener("click", deleteProgrammazioneFromForm);
ui.programmazioneFilter?.addEventListener("change", () => renderProgrammazioni());
ui.programmazioneForm?.addEventListener("submit", saveProgrammazione);
ui.openPanelBannerGestione?.addEventListener("click", () => openManagementPanel("banner"));
ui.openPrivateDocsBtn.addEventListener("click", openPrivateDocsPage);
ui.openPrivateDocsUploadBtn?.addEventListener("click", openPrivateDocsUploadPage);
ui.openPersonalServicesBtn.addEventListener("click", openPersonalServicesPage);
ui.openHoursBtn.addEventListener("click", openHoursPage);
ui.openPosBtn?.addEventListener("click", openPosPage);
ui.openSegnalazioniBtn.addEventListener("click", openSegnalazioniPage);
ui.openHowtoBtn.addEventListener("click", openHowtoPage);
ui.openBookPdfBtn?.addEventListener("click", openBookPdf);
ui.managementCloseBtn.addEventListener("click", closeManagementPanel);
ui.userToggleBtn.addEventListener("click", toggleUserDetailsPanel);
ui.weatherCloseBtn.addEventListener("click", closeWeatherModal);
ui.weatherCard?.addEventListener("click", openWeatherExternalDetail);
ui.weatherCard?.addEventListener("keydown", (event) => {
  if (event.key === "Enter" || event.key === " ") {
    event.preventDefault();
    openWeatherExternalDetail();
  }
});
ui.backFromFuelBtn.addEventListener("click", closeFuelPage);
ui.fuelMezzoDetailsBtn.addEventListener("click", toggleFuelMezzoDetails);
ui.backFromPersonalServicesBtn.addEventListener("click", closePersonalServicesPage);
ui.backFromSegnalazioniBtn.addEventListener("click", closeSegnalazioniPage);
ui.backFromHowtoBtn.addEventListener("click", closeHowtoPage);
ui.backFromPrivateDocsBtn.addEventListener("click", closePrivateDocsPage);
ui.backFromHoursBtn.addEventListener("click", closeHoursPage);
ui.backFromPosBtn?.addEventListener("click", closePosPage);
ui.posAddToggleBtn?.addEventListener("click", () => openPosDocumentForm());
ui.posCancelBtn?.addEventListener("click", closePosDocumentForm);
ui.posDocumentForm?.addEventListener("submit", savePosDocument);
ui.posSearch?.addEventListener("input", renderPosDocuments);
ui.hoursForm.addEventListener("submit", finalizeHoursReport);
ui.addHoursCommessaBtn.addEventListener("click", () => {
  unlockHoursFinalizeButton();
  addHoursCommessaBlock();
});
ui.hoursDate?.addEventListener("input", () => {
  unlockHoursFinalizeButton();
  Array.from(ui.hoursCommesseList?.querySelectorAll(".hours-commessa-card") || []).forEach((card) => {
    applyHoursSuggestedOperators(card, { force: true });
  });
});
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
ui.hoursTableExportGlobalBtn?.addEventListener("click", exportHoursGlobalMonthlyTable);
ui.hoursConfirmVisibleBtn?.addEventListener("click", handleConfirmVisiblePendingHours);
ui.hoursConfirmCancelBtn?.addEventListener("click", () => closeHoursConfirmModal(false));
ui.hoursConfirmOkBtn?.addEventListener("click", () => closeHoursConfirmModal(true));
ui.hoursConfirmModal?.addEventListener("click", (event) => {
  if (event.target === ui.hoursConfirmModal) closeHoursConfirmModal(false);
});
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
ui.globalCommessaForm?.addEventListener("submit", createGlobalCommessa);
ui.globalExcelFile?.addEventListener("change", onGlobalExcelSelected);
ui.globalImportBtn?.addEventListener("click", importPendingGlobalRows);
ui.globalUpdateBtn?.addEventListener("click", updateExistingGlobalRowsOnly);
ui.globalSheetUrlImportBtn?.addEventListener("click", importGlobalFromGoogleSheetUrl);
ui.globalCommessaSelect?.addEventListener("change", onGlobalCommessaSelectionChanged);
ui.globalImpiantoSearch?.addEventListener("input", onGlobalImpiantoSearchInput);
ui.globalImpiantoSearchForm?.addEventListener("submit", onGlobalImpiantoSearchSubmit);
ui.globalImpiantoSearch?.addEventListener("focus", renderGlobalImpianti);
ui.globalImpiantoDetailsCloseBtn?.addEventListener("click", closeGlobalImpiantoModal);
ui.globalCommesseLista?.addEventListener("click", onGlobalCommesseListClick);
ui.globalOpenReportBtn?.addEventListener("click", () => handleOpenGlobalSegnalazioneClick());
ui.globalImpiantoWhatsappBtn?.addEventListener("click", () => handleOpenGlobalSegnalazioneClick());
ui.globalReportCloseBtn?.addEventListener("click", closeGlobalSegnalazioneModal);
ui.globalReportForm?.addEventListener("submit", submitGlobalSegnalazioneWhatsapp);
ui.globalReportImpiantoSelect?.addEventListener("change", onGlobalSegnalazioneImpiantoChange);
ui.globalReportModal?.addEventListener("click", (event) => {
  if (event.target === ui.globalReportModal) closeGlobalSegnalazioneModal();
});
ui.adminUserForm.addEventListener("submit", addAdminUserByEmail);
ui.externalAppForm.addEventListener("submit", saveExternalAppForCurrentUser);
ui.resourceForm.addEventListener("submit", addResourceItem);
ui.notificationForm?.addEventListener("submit", createUserNotification);
ui.notificationCancelUploadBtn?.addEventListener("click", cancelNotificationUpload);
ui.notificationOpenCalendarBtn?.addEventListener("click", (event) => {
  event.preventDefault();
  openNotificationCalendarView();
});
ui.notificationCalendarBackBtn?.addEventListener("click", closeNotificationCalendarView);
ui.notificationCalendarPrevBtn?.addEventListener("click", () => moveNotificationCalendarMonth(-1));
ui.notificationCalendarNextBtn?.addEventListener("click", () => moveNotificationCalendarMonth(1));
ui.notificationSendAllToggle?.addEventListener("change", onNotificationSendAllChange);
ui.bannerConfigForm?.addEventListener("submit", saveWorkBannerConfig);
ui.bannerDisableBtn?.addEventListener("click", disableWorkBanner);
ui.bannerAddNoteBtn?.addEventListener("click", saveWorkBannerNoteForDate);
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
ui.userAlertOkBtn?.addEventListener("click", acknowledgeActiveUserAlert);
ui.userAlertLaterBtn?.addEventListener("click", postponeActiveUserAlert);
ui.notificationDocViewerCloseBtn?.addEventListener("click", closeNotificationDocumentViewer);
ui.notificationDocViewerModal?.addEventListener("click", (event) => {
  if (event.target === ui.notificationDocViewerModal) closeNotificationDocumentViewer();
});
window.addEventListener("online", updateConnectivityStatus);
window.addEventListener("offline", updateConnectivityStatus);
window.addEventListener("pagehide", markCurrentOperatorOffline);
ui.commessaResourceViewerCloseBtn.addEventListener("click", closeCommessaResourceViewer);
document.querySelectorAll(".resource-filter-btn").forEach((btn) => {
  btn.addEventListener("click", () => {
    activeResourceManageFilter = btn.dataset.resourceFilter || "";
    renderResourceManageFilters();
    renderResourcesList();
  });
});

startQuickSquadraWindowTicker();
addSquadraRow();
initHoursPage();
initGeolocation();
prefillSegnalazioneDateTime();
renderHowtoFaq();
applyRoute();
window.addEventListener("hashchange", applyRoute);
window.addEventListener("popstate", applyRoute);
loadPendingSheetExports();
startSheetRetryLoop();
initHelpCenterFaq();
renderResourceManageFilters();
updateResourceFormByType();
updateConnectivityStatus();
initPwaCapabilities();
initNativeGeofenceBridge();
initWorkBannerObservers();


function getNativeHeraGeofencePlugin() {
  const plugins = window.Capacitor && window.Capacitor.Plugins ? window.Capacitor.Plugins : null;
  if (!plugins || !plugins.HeraGeofence) return null;
  return plugins.HeraGeofence;
}

async function initNativeGeofenceBridge() {
  const plugin = getNativeHeraGeofencePlugin();
  if (!plugin) {
    return;
  }

  try {
    const status = await plugin.status();
    const active = Boolean(status && status.active);
    if (ui.gpsStatus) {
      ui.gpsStatus.textContent = active
        ? "Geofence nativo Android attivo (trigger anche ad app chiusa)."
        : "Geofence nativo Android disponibile ma non attivo.";
    }
  } catch (error) {
    console.warn("Status geofence nativo non disponibile:", error);
  }

  window.heraNativeGeofence = {
    activate: async () => plugin.activate(),
    deactivate: async () => plugin.deactivate(),
    status: async () => plugin.status()
  };
}

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

function normalizeWorkBannerConfig(payload = {}) {
  const rawText = typeof payload.text === "string" ? payload.text : "";
  const text = rawText.trim();
  const enabled = Boolean(payload.enabled);
  const speedNumber = Number(payload.speed);
  const speed = Number.isFinite(speedNumber) && speedNumber >= 5 && speedNumber <= 800
    ? Math.round(speedNumber)
    : null;
  const notes = Array.isArray(payload.notes)
    ? payload.notes
      .map((entry) => {
        const dateKey = String(entry?.dateKey || "").trim();
        const note = String(entry?.note || "").trim();
        if (!/^\d{4}-\d{2}-\d{2}$/.test(dateKey) || !note) return null;
        return { dateKey, note };
      })
      .filter(Boolean)
      .sort((a, b) => a.dateKey.localeCompare(b.dateKey))
    : [];
  return { text, enabled, speed, notes };
}

function loadWorkBannerForm(config = {}) {
  if (ui.bannerTextInput) ui.bannerTextInput.value = config.text || "";
  if (ui.bannerEnabledToggle) ui.bannerEnabledToggle.checked = Boolean(config.enabled);
  if (ui.bannerSpeedInput) ui.bannerSpeedInput.value = Number.isFinite(Number(config.speed)) ? String(config.speed) : "";
  if (ui.bannerNoteDate && !ui.bannerNoteDate.value) ui.bannerNoteDate.value = getDateKeyFromLocalDate(new Date());
  renderWorkBannerNotesList(config.notes || []);
}

function renderWorkBannerNotesList(notes = []) {
  if (!ui.bannerNotesList) return;
  const safeNotes = Array.isArray(notes) ? notes : [];
  if (!safeNotes.length) {
    ui.bannerNotesList.innerHTML = "<p class='muted'>Nessuna nota programmata.</p>";
    return;
  }
  ui.bannerNotesList.innerHTML = "";
  safeNotes.forEach((entry) => {
    const row = document.createElement("div");
    row.className = "list-item";
    const dateLabel = new Date(`${entry.dateKey}T00:00:00`).toLocaleDateString("it-IT");
    row.innerHTML = `
      <div>
        <strong>${escapeHTML(dateLabel)}</strong>
        <p class="muted">${escapeHTML(entry.note)}</p>
      </div>
    `;
    if (canManageData()) {
      const actions = document.createElement("div");
      actions.className = "item-actions";
      const deleteBtn = createButton("Rimuovi", async () => {
        await deleteWorkBannerNote(entry.dateKey);
      });
      actions.appendChild(deleteBtn);
      row.appendChild(actions);
    }
    ui.bannerNotesList.appendChild(row);
  });
}

function getActiveWorkBannerMessage(config = {}) {
  const notes = Array.isArray(config.notes) ? [...config.notes] : [];
  if (notes.length) {
    const now = new Date();
    const todayKey = getDateKeyFromLocalDate(now);
    if (now.getHours() >= WORK_BANNER_NEXT_NOTE_PREVIEW_HOUR) {
      const tomorrow = new Date(now);
      tomorrow.setDate(tomorrow.getDate() + 1);
      const tomorrowKey = getDateKeyFromLocalDate(tomorrow);
      const tomorrowNote = notes.find((entry) => entry.dateKey === tomorrowKey);
      if (tomorrowNote) {
        const tomorrowLabel = tomorrow.toLocaleDateString("it-IT");
        return { text: `📅 ${tomorrowLabel} · ${tomorrowNote.note}`, isScheduled: true };
      }
    }
    const todayNote = notes.find((entry) => entry.dateKey === todayKey);
    if (todayNote) return { text: todayNote.note, isScheduled: false };
    const nextNote = notes.find((entry) => entry.dateKey >= todayKey) || notes[0];
    const dateLabel = new Date(`${nextNote.dateKey}T00:00:00`).toLocaleDateString("it-IT");
    return { text: `📅 ${dateLabel} · ${nextNote.note}`, isScheduled: true };
  }
  return { text: String(config.text || "").trim(), isScheduled: false };
}

function initWorkBannerObservers() {
  if (workBannerResizeObserver || typeof ResizeObserver !== "function") return;
  if (!ui.workBannerHome || !ui.workBannerText) return;
  workBannerResizeObserver = new ResizeObserver(() => {
    updateWorkBannerAnimationDuration();
  });
  workBannerResizeObserver.observe(ui.workBannerHome);
  workBannerResizeObserver.observe(ui.workBannerText);
  window.addEventListener("orientationchange", updateWorkBannerAnimationDuration);
  document.addEventListener("visibilitychange", () => {
    if (!document.hidden) updateWorkBannerAnimationDuration();
  });
}

function updateWorkBannerAnimationDuration() {
  if (!ui.workBannerHome || !ui.workBannerText) return;
  const speedSetting = Number.isFinite(Number(currentWorkBannerConfig.speed))
    ? Number(currentWorkBannerConfig.speed)
    : WORK_BANNER_DEFAULT_DURATION_SEC;
  const durationSec = Math.min(Math.max(speedSetting, 5), 800);
  ui.workBannerHome.style.setProperty("--banner-scroll-duration", `${durationSec.toFixed(2)}s`);
}

function applyWorkBannerConfig(config = {}) {
  if (!ui.workBannerHome || !ui.workBannerText) return;
  const normalized = normalizeWorkBannerConfig(config);
  currentWorkBannerConfig = normalized;
  const activeMessage = getActiveWorkBannerMessage(normalized);
  const shouldShow = normalized.enabled && Boolean(activeMessage.text);
  ui.workBannerHome.classList.toggle("hidden", !shouldShow);
  if (!shouldShow) {
    ui.workBannerText.textContent = "";
    return;
  }
  ui.workBannerText.textContent = `${activeMessage.text}   •   ${activeMessage.text}   •   ${activeMessage.text}`;
  window.requestAnimationFrame(updateWorkBannerAnimationDuration);
}

function subscribeWorkBanner() {
  stopWorkBannerSubscription();
  unsubscribeWorkBanner = db.collection(WORK_BANNER_CONFIG_PATH.collection).doc(WORK_BANNER_CONFIG_PATH.doc)
    .onSnapshot((doc) => {
      const config = normalizeWorkBannerConfig(doc.exists ? (doc.data() || {}) : {});
      applyWorkBannerConfig(config);
      loadWorkBannerForm(config);
      if (ui.bannerFeedback) ui.bannerFeedback.textContent = "";
    }, (error) => {
      console.warn("Sottoscrizione banner home non disponibile:", error);
      if (ui.bannerFeedback && canManageData()) {
        ui.bannerFeedback.textContent = "Errore lettura banner. Riprova più tardi.";
      }
    });
}

function stopWorkBannerSubscription() {
  if (unsubscribeWorkBanner) {
    unsubscribeWorkBanner();
    unsubscribeWorkBanner = null;
  }
}

async function saveWorkBannerConfig(event) {
  event.preventDefault();
  if (!currentUser) {
    if (ui.bannerFeedback) ui.bannerFeedback.textContent = "Esegui il login per gestire il banner.";
    return;
  }
  if (!canManageData()) {
    if (ui.bannerFeedback) ui.bannerFeedback.textContent = "Solo gli admin possono salvare il banner.";
    return;
  }
  const text = String(ui.bannerTextInput?.value || "").trim();
  const notes = Array.isArray(currentWorkBannerConfig.notes) ? currentWorkBannerConfig.notes : [];
  if (!text && !notes.length) {
    if (ui.bannerFeedback) ui.bannerFeedback.textContent = "Inserisci un fallback o almeno una nota calendario.";
    return;
  }
  const enabled = Boolean(ui.bannerEnabledToggle?.checked);
  const speedRaw = String(ui.bannerSpeedInput?.value || "").trim();
  const speedNum = Number(speedRaw);
  const speed = speedRaw && Number.isFinite(speedNum) && speedNum >= 5 && speedNum <= 800 ? Math.round(speedNum) : null;
  try {
    await db.collection(WORK_BANNER_CONFIG_PATH.collection).doc(WORK_BANNER_CONFIG_PATH.doc).set({
      text,
      notes,
      enabled,
      speed,
      updatedAt: firebase.firestore.FieldValue.serverTimestamp(),
      updatedBy: currentUser?.email || ""
    }, { merge: true });
    if (ui.bannerFeedback) ui.bannerFeedback.textContent = "Banner salvato correttamente.";
  } catch (error) {
    console.error("Salvataggio banner non riuscito:", error);
    if (ui.bannerFeedback) ui.bannerFeedback.textContent = "Errore durante il salvataggio del banner.";
  }
}

async function saveWorkBannerNoteForDate() {
  if (!currentUser) {
    if (ui.bannerFeedback) ui.bannerFeedback.textContent = "Esegui il login per gestire il banner.";
    return;
  }
  if (!canManageData()) {
    if (ui.bannerFeedback) ui.bannerFeedback.textContent = "Solo gli admin possono gestire le note banner.";
    return;
  }
  const dateKey = String(ui.bannerNoteDate?.value || "").trim();
  const note = String(ui.bannerNoteInput?.value || "").trim();
  if (!/^\d{4}-\d{2}-\d{2}$/.test(dateKey)) {
    if (ui.bannerFeedback) ui.bannerFeedback.textContent = "Seleziona una data valida.";
    return;
  }
  if (!note) {
    if (ui.bannerFeedback) ui.bannerFeedback.textContent = "Inserisci la nota da associare al giorno selezionato.";
    return;
  }
  const currentNotes = Array.isArray(currentWorkBannerConfig.notes) ? [...currentWorkBannerConfig.notes] : [];
  const withoutDate = currentNotes.filter((entry) => entry.dateKey !== dateKey);
  const nextNotes = [...withoutDate, { dateKey, note }].sort((a, b) => a.dateKey.localeCompare(b.dateKey));
  try {
    await db.collection(WORK_BANNER_CONFIG_PATH.collection).doc(WORK_BANNER_CONFIG_PATH.doc).set({
      notes: nextNotes,
      updatedAt: firebase.firestore.FieldValue.serverTimestamp(),
      updatedBy: currentUser?.email || ""
    }, { merge: true });
    if (ui.bannerFeedback) ui.bannerFeedback.textContent = "Nota calendario salvata.";
    if (ui.bannerNoteInput) ui.bannerNoteInput.value = "";
  } catch (error) {
    console.error("Salvataggio nota banner non riuscito:", error);
    if (ui.bannerFeedback) ui.bannerFeedback.textContent = "Errore durante il salvataggio della nota.";
  }
}

async function deleteWorkBannerNote(dateKey) {
  if (!currentUser || !canManageData()) return;
  const currentNotes = Array.isArray(currentWorkBannerConfig.notes) ? [...currentWorkBannerConfig.notes] : [];
  const nextNotes = currentNotes.filter((entry) => entry.dateKey !== dateKey);
  try {
    await db.collection(WORK_BANNER_CONFIG_PATH.collection).doc(WORK_BANNER_CONFIG_PATH.doc).set({
      notes: nextNotes,
      updatedAt: firebase.firestore.FieldValue.serverTimestamp(),
      updatedBy: currentUser?.email || ""
    }, { merge: true });
    if (ui.bannerFeedback) ui.bannerFeedback.textContent = "Nota calendario rimossa.";
  } catch (error) {
    console.error("Rimozione nota banner non riuscita:", error);
    if (ui.bannerFeedback) ui.bannerFeedback.textContent = "Errore durante la rimozione della nota.";
  }
}

async function disableWorkBanner() {
  if (!currentUser) {
    if (ui.bannerFeedback) ui.bannerFeedback.textContent = "Esegui il login per gestire il banner.";
    return;
  }
  if (!canManageData()) {
    if (ui.bannerFeedback) ui.bannerFeedback.textContent = "Solo gli admin possono disattivare il banner.";
    return;
  }
  try {
    await db.collection(WORK_BANNER_CONFIG_PATH.collection).doc(WORK_BANNER_CONFIG_PATH.doc).set({
      enabled: false,
      updatedAt: firebase.firestore.FieldValue.serverTimestamp(),
      updatedBy: currentUser?.email || ""
    }, { merge: true });
    if (ui.bannerEnabledToggle) ui.bannerEnabledToggle.checked = false;
    if (ui.bannerFeedback) ui.bannerFeedback.textContent = "Banner disattivato.";
  } catch (error) {
    console.error("Disattivazione banner non riuscita:", error);
    if (ui.bannerFeedback) ui.bannerFeedback.textContent = "Errore durante la disattivazione del banner.";
  }
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

pendingImpiantoActions = loadPendingImpiantoActions();
renderPendingWhatsappList();

window.addEventListener("online", () => {
  syncPendingImpiantoActions();
  runWhazzupPendingDoneSafetyCheck();
});
window.addEventListener("offline", () => {
  renderPendingWhatsappList();
});
document.addEventListener("visibilitychange", () => {
  if (!document.hidden) runWhazzupPendingDoneSafetyCheck();
});

auth.onAuthStateChanged((user) => {
  currentUser = user || null;
  const loggedIn = Boolean(user);

  ui.loginBtn.disabled = loggedIn;
  ui.switchAccountBtn.classList.toggle("hidden", !loggedIn);
  ui.switchAccountBtn.disabled = !loggedIn;
  ui.logoutBtn.disabled = !loggedIn;
  updateDriveConnectVisibility();
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
  stopCommessaNotesSubscription();
  stopChatSubscription();
  stopDriveBridgeSubscription();
  stopPersonaleSubscription();
  stopMezziSubscription();
  stopSquadreSubscription();
  stopUsersSubscription();
  stopOperatorPositionsSubscription();
  stopAdminUsersSubscription();
  stopResourcesSubscription();
  stopGlobalCommesseSubscription();
  stopGlobalImpiantiSubscription();
  stopPrivateDocsSubscription();
  stopPosDocumentsSubscription();
  stopGpsRequestsSubscription();
  stopGlobalNotificationsSubscription();
  stopWorkBannerSubscription();
  stopUserAlertsSubscription();
  stopChatRetentionLoop();
  stopHoursDeadlineAlertLoop();
  selectedCommessaId = "";
  selectedCommessaName = "";
  updateCommessaContextUI();
  window.location.hash = "";
  commesseLoadState = { status: "idle", message: "" };
  isCommesseHomeCardVisible = false;
  syncCommesseHomeToggle();
  ui.commesseLista.innerHTML = "";
  ui.squadraCommessa.innerHTML = "<option value=''>Seleziona commessa</option>";
  ui.squadreLista.innerHTML = "";
  squadreLoadState = { status: "loading", message: "Caricamento squadre..." };
  manualSquadreFilterDateKey = "";
  sharedSquadreDateKey = "";
  startupAssignedCommessaAutoOpenDone = false;
  sharedSquadreViewConfigLoaded = false;
  squadreByCommessa = new Map();
  squadreHistoryByDate = new Map();
  commesseById = new Map();
  initializeAutomaticSquadreDate();
  globalCommesseById = new Map();
  globalImpianti = [];
  pendingGlobalRows = [];
  selectedGlobalCommessaId = "";
  resourceRecords = [];
  privateDocsRecords = [];
  posDocuments = [];
  gpsUpdateRequests = [];
  operatorPositions = [];
  hoursApprovalRequests = [];
  lastPublishedUserPos = null;
  lastPositionPublishAt = 0;
  renderPrivateDocsList();
  renderPosDocuments();
  renderResourceButtonsForCommessa();
  closeCommessaResourceViewer();
  renderParentCommessaOverview();
  ui.impiantiLista.innerHTML = loggedIn
    ? "<p class='muted'>Seleziona una commessa.</p>"
    : "<p class='muted'>Fai login per vedere le commesse.</p>";
  clearMap();
  lastReadChatAt = null;
  resetDriveState();
  renderChat([]);
  applyRoute();
  subscribeWorkBanner();

  if (loggedIn) {
    startPresenceHeartbeat();
    upsertCurrentPlatformUser();
    initGeolocation({ forcePublishCurrent: true });
    subscribeCommesse();
    subscribeChat();
    subscribeAdminUsers();
    subscribeUsers();
    subscribeOperatorPositions();
    subscribeDriveBridge();
    subscribePersonale();
    subscribeMezzi();
    subscribeSquadre();
    subscribeResources();
    subscribeGlobalCommesse();
    subscribePrivateDocs();
    subscribePosDocuments();
    subscribeGpsRequests();
    subscribeGlobalNotifications();
    subscribeWorkBanner();
    subscribeUserAlerts();
    processPendingSheetExports();
    startChatRetentionLoop();
    startHoursDeadlineAlertLoop();
    repairDuplicateHours().catch((error) => {
      console.error("Riparazione automatica duplicati ore all'avvio non riuscita:", error);
    });
  } else {
    markCurrentOperatorOffline();
    subscribeCommesse();
    subscribeSquadre();
    subscribePosDocuments();
    stopPresenceHeartbeat();
    applyWorkBannerConfig({ text: "", enabled: false, speed: null });
    closeUserAlertModal();
  }
  renderHeaderActivitySummary();
  renderExternalApps();
  renderPendingWhatsappList();
  syncPendingImpiantoActions();
  fetchWeather();
  renderNextActionCard();
});

function updateAdminControls() {
  const canManage = canManageData();
  updateDriveConnectVisibility();
  ui.openPosBtn?.classList.remove("hidden");
  if (ui.openPosBtn) ui.openPosBtn.disabled = false;
  ui.operatorPositionsToggleBtn?.classList.add("hidden");
  if (ui.operatorPositionsToggleBtn) ui.operatorPositionsToggleBtn.disabled = true;
  ui.chatClearBtn?.classList.toggle("hidden", !canManage);
  if (ui.chatClearBtn) ui.chatClearBtn.disabled = !canManage;
  ui.posAdminCard?.classList.toggle("hidden", !canManage);
  if (ui.posAddToggleBtn) ui.posAddToggleBtn.disabled = !canManage;
  ui.posDocumentForm?.querySelectorAll("input, textarea, select, button").forEach((el) => { el.disabled = !canManage; });
  [ui.openPanelCommesse, ui.openPanelSquadre, ui.openPanelPersonale, ui.openPanelMezzi, ui.openPanelUtenti, ui.openPanelGlobal, ui.openPanelBanner, ui.openPanelBannerGestione, ui.openPanelInfoUtili, ui.openPanelNotifiche, ui.openPanelProgrammazione]
    .forEach((button) => button.classList.toggle("hidden", !canManage));
  ui.programmazioneAddBtn?.classList.toggle("hidden", !canManage);
  ui.openPanelBanner?.classList.toggle("hidden", !auth.currentUser);
  ui.openPanelBannerGestione?.classList.toggle("hidden", !auth.currentUser);
  ui.commessaName.disabled = !canManage;
  if (ui.openOrganizeCommesseBtn) ui.openOrganizeCommesseBtn.disabled = !canManage;
  if (ui.parentCommessaName) ui.parentCommessaName.disabled = !canManage;
  if (ui.parentCommessaCode) ui.parentCommessaCode.disabled = !canManage;
  if (ui.moveParentCommessaSelect) ui.moveParentCommessaSelect.disabled = !canManage;
  ui.parentCommessaForm?.querySelector("button[type='submit']")?.toggleAttribute("disabled", !canManage);
  ui.moveSubcommesseForm?.querySelector("button[type='submit']")?.toggleAttribute("disabled", !canManage);
  const submitBtn = ui.commessaForm.querySelector("button[type='submit']");
  if (submitBtn) submitBtn.disabled = !canManage;
  ui.personaleNome.disabled = !canManage;
  ui.mezzoNId.disabled = !canManage;
  ui.mezzoMarca.disabled = !canManage;
  ui.mezzoModello.disabled = !canManage;
  if (ui.globalCommessaName) ui.globalCommessaName.disabled = !canManage;
  if (ui.globalCommessaSelect) ui.globalCommessaSelect.disabled = !canManage;
  if (ui.globalExcelFile) ui.globalExcelFile.disabled = !canManage;
  refreshGlobalImportButtons();
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
  if (ui.bannerTextInput) ui.bannerTextInput.disabled = !canManage;
  if (ui.bannerNoteDate) ui.bannerNoteDate.disabled = !canManage;
  if (ui.bannerNoteInput) ui.bannerNoteInput.disabled = !canManage;
  if (ui.bannerAddNoteBtn) ui.bannerAddNoteBtn.disabled = !canManage;
  if (ui.bannerEnabledToggle) ui.bannerEnabledToggle.disabled = !canManage;
  if (ui.bannerSpeedInput) ui.bannerSpeedInput.disabled = !canManage;
  if (ui.bannerDisableBtn) ui.bannerDisableBtn.disabled = !canManage;
  if (ui.bannerConfigForm && ui.bannerConfigForm.querySelector("button[type='submit']")) ui.bannerConfigForm.querySelector("button[type='submit']").disabled = !canManage;
  if (ui.notificationTitle) ui.notificationTitle.disabled = !canManage;
  if (ui.notificationDate) ui.notificationDate.disabled = !canManage;
  if (ui.notificationSendAllToggle) ui.notificationSendAllToggle.disabled = !canManage;
  if (ui.notificationUserSelect) ui.notificationUserSelect.disabled = !canManage || Boolean(ui.notificationSendAllToggle?.checked);
  if (ui.notificationMessage) ui.notificationMessage.disabled = !canManage;
  if (ui.notificationAttachments) ui.notificationAttachments.disabled = !canManage;
  if (ui.notificationSubmit) ui.notificationSubmit.disabled = !canManage;
  if (ui.notificationCancelUploadBtn) ui.notificationCancelUploadBtn.disabled = !canManage || !notificationUploadInProgress;
  if (ui.notificationOpenCalendarBtn) ui.notificationOpenCalendarBtn.disabled = !canManage;
  renderWorkBannerNotesList(currentWorkBannerConfig.notes || []);
  if (ui.externalAppName) ui.externalAppName.disabled = !auth.currentUser;
  if (ui.externalAppUrl) ui.externalAppUrl.disabled = !auth.currentUser;
  if (ui.externalAppForm && ui.externalAppForm.querySelector("button[type='submit']")) {
    ui.externalAppForm.querySelector("button[type='submit']").disabled = !auth.currentUser;
  }
  ui.squadraCommessa.disabled = !canManage;
  syncCommesseHomeToggle();
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
  ui.squadraRows.querySelectorAll("input,textarea,select,button").forEach((el) => { el.disabled = !canManage; });
  if (ui.squadraForm.querySelector("button[type='submit']")) ui.squadraForm.querySelector("button[type='submit']").disabled = !canManage;
  ui.squadraHint.textContent = canManage
    ? "Suggerimento: usa i nomi in Personale e i mezzi in Mezzi per compilare le squadre."
    : "Solo l'admin può modificare personale, mezzi e composizione squadre.";
  updateResourceFormByType();
  renderUserPermissionList();
  renderNotificationTargetUsers();
  renderNotificationsList();
  renderExternalApps();
  renderCommesseHomeList();
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

function refreshApplicationData() {
  closeSideMenu();
  if (ui.refreshAppBtn) {
    ui.refreshAppBtn.disabled = true;
    ui.refreshAppBtn.classList.add("is-reloading");
  }
  if (ui.commesseNextAction) {
    ui.commesseNextAction.textContent = "Aggiornamento app in corso...";
  }
  const refreshUrl = new URL(window.location.href);
  refreshUrl.searchParams.set("refreshTs", String(Date.now()));
  window.location.replace(refreshUrl.toString());
}

function openManagementPanel(panel) {
  if (panel !== "banner" && !canManageData()) {
    closeSideMenu();
    return;
  }
  const panelMap = {
    commesse: { el: ui.panelCommesse, title: "Aggiungi commesse" },
    squadre: { el: ui.panelSquadre, title: "Composizione squadre" },
    personale: { el: ui.panelPersonale, title: "Personale" },
    mezzi: { el: ui.panelMezzi, title: "Mezzi" },
    utenti: { el: ui.panelUtenti, title: "Gestione utenti" },
    global: { el: ui.panelGlobal, title: "Global" },
    banner: { el: ui.panelBanner, title: "Banner home" },
    infoUtili: { el: ui.panelInfoUtili, title: "Informazioni utili" },
    notifiche: { el: ui.panelNotifiche, title: "Gestione notifiche" },
    programmazione: { el: ui.panelProgrammazione, title: "📅 Programmazione" }
  };
  const target = panelMap[panel];
  if (!target) return;
  [ui.panelCommesse, ui.panelSquadre, ui.panelPersonale, ui.panelMezzi, ui.panelUtenti, ui.panelGlobal, ui.panelBanner, ui.panelInfoUtili, ui.panelNotifiche, ui.panelProgrammazione].forEach((el) => el.classList.add("hidden"));
  target.el.classList.remove("hidden");
  ui.managementTitle.textContent = target.title;
  ui.managementPage.classList.remove("hidden");
  ui.managementPage.setAttribute("aria-hidden", "false");
  if (panel === "squadre") setDefaultSquadraCompositionDate({ force: true });
  if (panel === "global") setTimeout(() => globalMap.invalidateSize(), 60);
  if (panel === "notifiche") closeNotificationCalendarView();
  closeSideMenu();
}

function closeManagementPanel() {
  ui.managementPage.classList.add("hidden");
  ui.managementPage.setAttribute("aria-hidden", "true");
}

function openMapFullscreenPage() {
  if (!ui.mapFullscreenPage) return;
  isMapFullscreenPageOpen = true;
  drawAreaModeActive = false;
  drawnAreaPoints = [];
  drawnAreaRedoStack = [];
  isDrawingStrokeActive = false;
  renderDrawnArea();
  setFullscreenMapInteractivity(true);
  ui.impiantiPage.classList.add("hidden");
  ui.mapFullscreenPage.classList.remove("hidden");
  ui.mapFullscreenBtn.textContent = "⤢ Mappa a schermo intero";
  ui.mapDrawAreaBtn.textContent = "✏️ Disegna";
  syncDrawAreaToolbarState();
  setFullscreenFeedback("Usa “Disegna” per definire il perimetro di lavoro.");
  setTimeout(() => {
    fullscreenMap.setView(mainMapViewState.center, mainMapViewState.zoom, { animate: false });
    refreshFullscreenMapLayout();
    renderMap();
  }, 60);
  setTimeout(() => {
    if (fullscreenMap) fullscreenMap.invalidateSize({ pan: false, animate: false });
  }, 300);
}

function closeMapFullscreenPage() {
  if (!ui.mapFullscreenPage) return;
  destroyWeatherRadar();
  closeSelectedImpiantoDetail({ closePopup: true });
  isMapFullscreenPageOpen = false;
  drawAreaModeActive = false;
  isDrawingStrokeActive = false;
  setFullscreenMapInteractivity(true);
  ui.mapFullscreenPage.classList.add("hidden");
  ui.impiantiPage.classList.remove("hidden");
  ui.mapDrawAreaBtn.textContent = "✏️ Disegna";
  syncDrawAreaToolbarState();
  setFullscreenFeedback("Usa “Disegna” per definire il perimetro di lavoro.");
  setTimeout(() => map.invalidateSize(), 60);
}



const WEATHER_RADAR_MAX_ZOOM = 20;
const WEATHER_PROVIDER_DEFAULT_MAX_NATIVE_ZOOM = 18;
const RAINVIEWER_MAX_NATIVE_ZOOM = 10;
const TRANSPARENT_TILE_URL = "data:image/gif;base64,R0lGODlhAQABAIAAAAAAAP///ywAAAAAAQABAAACAUwAOw==";
const WEATHER_UNAVAILABLE_MESSAGE = "Dato non disponibile";
const OPENWEATHER_API_KEY_PLACEHOLDER = "%VITE_OPENWEATHER_API_KEY%";
const OPENWEATHER_TILE_BASE_URL = "https://maps.openweathermap.org/maps/2.0/weather";
const RAINVIEWER_API_URL = "https://api.rainviewer.com/public/weather-maps.json";
const WEATHER_TILE_PREVIEW_TIMEOUT_MS = 4500;
const WEATHER_LAYER_DEFINITIONS = {
  rain: {
    id: "rain",
    button: "🌧 Pioggia",
    title: "Pioggia / precipitazioni",
    label: "Pioggia",
    opacity: 0.6,
    providerLayers: { openweather: "PR0", rainviewer: "radar" },
    legend: ["🟦 Debole", "🟩 Moderata", "🟨 Forte", "🟥 Molto forte"],
    description: "Precipitazioni OpenWeatherMap Weather Maps 2.0; fallback RainViewer solo se il provider principale non è configurato."
  },
  clouds: {
    id: "clouds",
    button: "☁️ Nuvole",
    title: "Nuvolosità",
    label: "Nuvole",
    opacity: 0.52,
    providerLayers: { openweather: "CL", rainviewer: "satellite" },
    legend: ["⬛ Sereno", "⬜ Nubi basse", "☁️ Nubi dense", "🌩 Celle compatte"],
    description: "Copertura nuvolosa da OpenWeatherMap Weather Maps 2.0."
  },
  temperature: {
    id: "temperature",
    button: "🌡 Temperatura",
    title: "Temperatura",
    label: "Temperatura",
    opacity: 0.46,
    providerLayers: { openweather: "TA2" },
    legend: ["🟦 Freddo", "🟩 Mite", "🟨 Caldo", "🟥 Molto caldo"],
    description: "Temperatura a 2 metri da OpenWeatherMap Weather Maps 2.0."
  },
  wind: {
    id: "wind",
    button: "💨 Vento",
    title: "Vento",
    label: "Vento",
    opacity: 0.44,
    providerLayers: { openweather: "WND" },
    openWeatherParams: { use_norm: "true", arrow_step: "32" },
    legend: ["🟦 Brezza", "🟩 Moderato", "🟨 Forte", "🟥 Raffiche"],
    description: "Velocità e direzione del vento da OpenWeatherMap Weather Maps 2.0."
  },
  storms: {
    id: "storms",
    button: "⚡ Temporali",
    title: "Temporali",
    label: "Temporali",
    opacity: 0.58,
    providerLayers: { openweather: "PAC0", rainviewer: "radar" },
    usesFallbackLayerMessage: true,
    legend: ["🟦 Rovesci", "🟩 Pioggia", "🟨 Celle intense", "🟥 Possibile temporale"],
    description: "Temporali stimati dal layer di precipitazione convettiva OpenWeatherMap; se non disponibile usa il radar precipitazioni come fallback operativo."
  },
  alerts: {
    id: "alerts",
    button: "⚠️ Allerte",
    title: "Allerte meteo",
    label: "Allerte",
    opacity: 0.5,
    providerLayers: {},
    unavailable: true,
    legend: ["⚠️ Dato non disponibile"],
    description: "Dato non disponibile: le allerte ufficiali non sono esposte come tile meteo in questo provider/piano. La mappa resta navigabile."
  }
};

const WEATHER_PROVIDERS = {
  openweather: {
    id: "openweather",
    label: "OpenWeatherMap",
    priority: 1,
    attribution: "Meteo © OpenWeatherMap",
    sourceUrl: "https://openweathermap.org/api/weather-map-2",
    maxNativeZoom: WEATHER_PROVIDER_DEFAULT_MAX_NATIVE_ZOOM,
    maxZoom: WEATHER_RADAR_MAX_ZOOM,
    async loadFrames(definition) {
      const apiKey = getOpenWeatherApiKey();
      const layerCode = definition.providerLayers?.openweather;
      if (!apiKey) throw new Error("VITE_OPENWEATHER_API_KEY non configurata");
      if (!layerCode) throw new Error(`${definition.label}: ${WEATHER_UNAVAILABLE_MESSAGE}`);
      const params = new URLSearchParams({
        appid: apiKey,
        fill_bound: "true",
        opacity: String(definition.opacity ?? 0.8),
        ...(definition.openWeatherParams || {})
      });
      const tileUrl = `${OPENWEATHER_TILE_BASE_URL}/${encodeURIComponent(layerCode)}/{z}/{x}/{y}?${params.toString()}`;
      await verifyWeatherTileTemplate(tileUrl, this.label);
      return [{
        providerId: this.id,
        providerLabel: this.label,
        sourceUrl: this.sourceUrl,
        tileUrl,
        time: Math.floor(Date.now() / 1000),
        maxNativeZoom: this.maxNativeZoom,
        maxZoom: this.maxZoom,
        attribution: this.attribution
      }];
    }
  },
  rainviewer: {
    id: "rainviewer",
    label: "RainViewer fallback",
    priority: 2,
    attribution: "Meteo © RainViewer",
    sourceUrl: "https://www.rainviewer.com/",
    maxNativeZoom: RAINVIEWER_MAX_NATIVE_ZOOM,
    maxZoom: WEATHER_RADAR_MAX_ZOOM,
    async loadFrames(definition) {
      const layerType = definition.providerLayers?.rainviewer;
      if (!layerType) throw new Error(`${definition.label}: ${WEATHER_UNAVAILABLE_MESSAGE}`);
      const response = await fetch(RAINVIEWER_API_URL, { cache: "no-store" });
      if (!response.ok) throw new Error(`RainViewer ${response.status}`);
      const data = await response.json();
      const host = data.host || "https://tilecache.rainviewer.com";
      const sourceFrames = layerType === "satellite"
        ? [...(data.satellite?.infrared || []), ...(data.satellite?.past || []), ...(data.satellite?.nowcast || [])]
        : [...(data.radar?.past || []), ...(data.radar?.nowcast || [])];
      return normalizeRainViewerFrames(sourceFrames, host, definition, layerType, this);
    }
  }
};

function ensureRadarPane() {
  if (!fullscreenMap) return;
  const pane = fullscreenMap.getPane("radarPane") || fullscreenMap.createPane("radarPane");
  pane.style.zIndex = "350";
  pane.style.pointerEvents = "none";
  const markerPane = fullscreenMap.getPane("markerPane");
  const popupPane = fullscreenMap.getPane("popupPane");
  if (markerPane) markerPane.style.zIndex = "650";
  if (popupPane) popupPane.style.zIndex = "750";
  radarPaneInitialized = true;
}

function getRuntimeEnvValue(key) {
  const sources = [
    globalThis.__HERA_ENV__,
    globalThis.__HERA_CONFIG__,
    globalThis.__APP_CONFIG__,
    globalThis
  ];
  for (const source of sources) {
    const value = source?.[key];
    if (typeof value === "string" && value.trim()) return value.trim();
  }
  const metaValue = document.querySelector(`meta[name="${key}"]`)?.getAttribute("content");
  if (typeof metaValue === "string" && metaValue.trim()) return metaValue.trim();
  try {
    const storedValue = localStorage.getItem(key);
    if (typeof storedValue === "string" && storedValue.trim()) return storedValue.trim();
  } catch (error) {
    console.warn(`Configurazione ${key} non leggibile da localStorage:`, error);
  }
  return "";
}

function getOpenWeatherApiKey() {
  const key = getRuntimeEnvValue("VITE_OPENWEATHER_API_KEY") || getRuntimeEnvValue("OPENWEATHER_API_KEY") || OPENWEATHER_API_KEY_PLACEHOLDER;
  if (!key || key === OPENWEATHER_API_KEY_PLACEHOLDER || /^(undefined|null)$/i.test(key)) return "";
  return key;
}

function materializeWeatherTileUrl(template, sample = { z: 2, x: 2, y: 1 }) {
  return String(template || "")
    .replace(/\{z\}/g, String(sample.z))
    .replace(/\{x\}/g, String(sample.x))
    .replace(/\{y\}/g, String(sample.y));
}

function verifyWeatherTileTemplate(tileTemplate, providerLabel = "Provider meteo") {
  if (typeof Image !== "function") return Promise.resolve(true);
  const previewUrl = materializeWeatherTileUrl(tileTemplate);
  return new Promise((resolve, reject) => {
    const image = new Image();
    const timeoutId = setTimeout(() => {
      image.onload = null;
      image.onerror = null;
      reject(new Error(`${providerLabel}: anteprima tile scaduta`));
    }, WEATHER_TILE_PREVIEW_TIMEOUT_MS);
    image.onload = () => {
      clearTimeout(timeoutId);
      if (image.naturalWidth > 0 && image.naturalHeight > 0) {
        resolve(true);
      } else {
        reject(new Error(`${providerLabel}: tile vuoto`));
      }
    };
    image.onerror = () => {
      clearTimeout(timeoutId);
      reject(new Error(`${providerLabel}: tile non caricabile`));
    };
    image.referrerPolicy = "no-referrer-when-downgrade";
    image.src = previewUrl;
  });
}


function formatRadarFrameTime(frame) {
  const timestamp = Number(frame?.time || 0) * 1000;
  if (!timestamp) return "attuale";
  return new Date(timestamp).toLocaleTimeString("it-IT", { hour: "2-digit", minute: "2-digit" });
}

function getActiveWeatherLayerDefinition() {
  return WEATHER_LAYER_DEFINITIONS[activeWeatherLayerId] || WEATHER_LAYER_DEFINITIONS.rain;
}

function getWeatherFramesForActiveLayer() {
  return weatherFramesBySource[activeWeatherLayerId] || [];
}

function getWeatherProviderSourceLabel(frame = radarFrames[radarFrameIndex]) {
  if (frame?.providerLabel) return frame.providerLabel;
  const providerId = frame?.providerId;
  return WEATHER_PROVIDERS[providerId]?.label || "OpenWeatherMap";
}

function getWeatherProviderSourceUrl(frame = radarFrames[radarFrameIndex]) {
  return frame?.sourceUrl || WEATHER_PROVIDERS[frame?.providerId]?.sourceUrl || WEATHER_PROVIDERS.openweather.sourceUrl;
}

function getRainViewerFramePath(frame) {
  const path = String(frame?.path || "").trim();
  if (!path) return "";
  return path.startsWith("/") ? path : `/${path}`;
}

function buildRainViewerTileUrl(frame, definition, layerType) {
  const host = String(frame?.host || "https://tilecache.rainviewer.com").replace(/\/$/, "");
  const path = getRainViewerFramePath(frame);
  if (layerType === "satellite" && path) return `${host}${path}/256/{z}/{x}/{y}/0/0_0.png`;
  const colorScheme = encodeURIComponent(definition.colorScheme || frame?.colorScheme || "2");
  const smooth = encodeURIComponent(definition.smooth || "1_1");
  return `${host}${path}/256/{z}/{x}/{y}/${colorScheme}/${smooth}.png`;
}

function normalizeRainViewerFrames(frames, host, definition, layerType, provider) {
  return (frames || [])
    .filter((frame) => frame && frame.path && frame.time)
    .map((frame) => ({
      ...frame,
      host,
      providerId: provider.id,
      providerLabel: provider.label,
      sourceUrl: provider.sourceUrl,
      tileUrl: buildRainViewerTileUrl({ ...frame, host }, definition, layerType),
      maxNativeZoom: provider.maxNativeZoom,
      maxZoom: provider.maxZoom,
      attribution: provider.attribution
    }));
}

function buildRadarTileLayer(frame, definition = getActiveWeatherLayerDefinition()) {
  return L.tileLayer(frame.tileUrl || TRANSPARENT_TILE_URL, {
    pane: "radarPane",
    opacity: 0,
    minNativeZoom: 0,
    maxNativeZoom: Number(frame.maxNativeZoom || WEATHER_PROVIDER_DEFAULT_MAX_NATIVE_ZOOM),
    maxZoom: Number(frame.maxZoom || WEATHER_RADAR_MAX_ZOOM),
    updateWhenIdle: true,
    updateWhenZooming: false,
    keepBuffer: 2,
    reuseTiles: true,
    detectRetina: true,
    noWrap: false,
    crossOrigin: false,
    errorTileUrl: TRANSPARENT_TILE_URL,
    className: `weather-radar-tile weather-radar-tile--${definition.id}`,
    attribution: frame.attribution || WEATHER_PROVIDERS.openweather.attribution
  });
}

function updateRadarButtonState() {
  if (!ui.mapRadarToggleBtn) return;
  ui.mapRadarToggleBtn.classList.toggle("active", radarActive);
  ui.mapRadarToggleBtn.setAttribute("aria-pressed", String(radarActive));
  ui.mapRadarToggleBtn.disabled = radarLoading;
}

function stopRadarPlayback() {
  if (radarPlayTimer) {
    clearInterval(radarPlayTimer);
    radarPlayTimer = null;
  }
}

function updateWeatherSourceLink() {
  const sourceLink = radarControlsEl?.querySelector("[data-weather-source]");
  if (!sourceLink) return;
  const frame = radarFrames[radarFrameIndex] || null;
  sourceLink.textContent = getWeatherProviderSourceLabel(frame);
  sourceLink.href = getWeatherProviderSourceUrl(frame);
}

function syncRadarControls() {
  if (!radarControlsEl) return;
  const frame = radarFrames[radarFrameIndex] || null;
  const definition = getActiveWeatherLayerDefinition();
  const playBtn = radarControlsEl.querySelector("[data-radar-play]");
  const slider = radarControlsEl.querySelector("[data-radar-slider]");
  const timeLabel = radarControlsEl.querySelector("[data-radar-time]");
  const info = radarControlsEl.querySelector("[data-weather-layer-info]");
  if (playBtn) {
    playBtn.textContent = radarPlaying ? "⏸" : "▶";
    playBtn.setAttribute("aria-label", radarPlaying ? "Pausa radar meteo" : "Avvia radar meteo");
    playBtn.disabled = radarFrames.length < 2;
  }
  if (slider) {
    slider.max = String(Math.max(radarFrames.length - 1, 0));
    slider.value = String(radarFrameIndex);
    slider.disabled = radarFrames.length < 2;
  }
  if (timeLabel) {
    timeLabel.textContent = frame
      ? `${definition.label} ${formatRadarFrameTime(frame)}`
      : `${definition.label} — ${WEATHER_UNAVAILABLE_MESSAGE}`;
  }
  if (info) {
    const providerLabel = frame ? `Fonte: ${getWeatherProviderSourceLabel(frame)}. ` : "";
    const fallbackMessage = definition.usesFallbackLayerMessage ? `${WEATHER_UNAVAILABLE_MESSAGE} come layer dedicato: uso precipitazioni intense. ` : "";
    info.textContent = `${providerLabel}${fallbackMessage}${definition.description}`;
  }
  radarControlsEl.querySelectorAll("[data-weather-layer]").forEach((button) => {
    const selected = button.dataset.weatherLayer === definition.id;
    button.classList.toggle("active", selected);
    button.setAttribute("aria-pressed", String(selected));
  });
  updateWeatherSourceLink();
  syncWeatherLegend();
}

function syncWeatherLegend(message = "") {
  if (!weatherLegendEl) return;
  const definition = getActiveWeatherLayerDefinition();
  const unavailable = message || (!radarFrames.length ? WEATHER_UNAVAILABLE_MESSAGE : "");
  weatherLegendEl.innerHTML = `
    <div class="weather-radar-legend-title">${escapeHTML(definition.title)}</div>
    ${unavailable ? `<div class="weather-radar-unavailable">${escapeHTML(unavailable)}</div>` : ""}
    <div class="weather-radar-legend-items">
      ${definition.legend.map((item) => `<span>${escapeHTML(item)}</span>`).join("")}
    </div>
  `;
}

function showRadarFrame(index, options = {}) {
  if (!radarActive || !radarFrames.length) return;
  radarFrameIndex = Math.max(0, Math.min(index, radarFrames.length - 1));
  const frame = radarFrames[radarFrameIndex];
  const definition = getActiveWeatherLayerDefinition();
  const nextLayer = buildRadarTileLayer(frame, definition);
  const targetOpacity = definition.opacity || 0.58;
  nextLayer.addTo(fullscreenMap);
  const finalizeLayer = () => {
    if (!fullscreenMap.hasLayer(nextLayer)) return;
    nextLayer.setOpacity(targetOpacity);
    if (radarLayer && radarLayer !== nextLayer && fullscreenMap.hasLayer(radarLayer)) fullscreenMap.removeLayer(radarLayer);
    radarLayer = nextLayer;
  };
  nextLayer.once("load", finalizeLayer);
  nextLayer.on("tileerror", (event) => {
    if (event?.tile) event.tile.src = TRANSPARENT_TILE_URL;
  });
  setTimeout(finalizeLayer, options.immediate ? 120 : 700);
  syncRadarControls();
}

function startRadarPlayback() {
  stopRadarPlayback();
  if (!radarActive || !radarPlaying || radarFrames.length < 2) return;
  radarPlayTimer = setInterval(() => {
    showRadarFrame((radarFrameIndex + 1) % radarFrames.length);
  }, 1500);
}

function createRadarControls() {
  destroyRadarControlsOnly();
  const wrap = document.querySelector(".map-fullscreen-map-wrap");
  if (!wrap) return;
  weatherLayerSelectorEl = document.createElement("div");
  weatherLayerSelectorEl.className = "weather-layer-selector";
  weatherLayerSelectorEl.innerHTML = Object.values(WEATHER_LAYER_DEFINITIONS)
    .map((definition) => `<button type="button" data-weather-layer="${definition.id}" aria-pressed="false">${definition.button}</button>`)
    .join("");
  weatherLayerSelectorEl.querySelectorAll("[data-weather-layer]").forEach((button) => {
    button.addEventListener("click", () => switchWeatherLayer(button.dataset.weatherLayer));
  });

  weatherLegendEl = document.createElement("div");
  weatherLegendEl.className = "weather-radar-legend";

  radarControlsEl = document.createElement("div");
  radarControlsEl.className = "weather-radar-controls";
  radarControlsEl.innerHTML = `
    <button class="weather-radar-play" type="button" data-radar-play aria-label="Pausa radar meteo">⏸</button>
    <div class="weather-radar-timeline">
      <span class="weather-radar-time" data-radar-time>Meteo --:--</span>
      <input type="range" min="0" max="0" value="0" step="1" data-radar-slider aria-label="Timeline radar meteo">
      <small data-weather-layer-info></small>
    </div>
    <a class="weather-radar-source" data-weather-source href="https://openweathermap.org/api/weather-map-2" target="_blank" rel="noopener noreferrer">OpenWeatherMap</a>
  `;
  radarControlsEl.querySelector("[data-radar-play]")?.addEventListener("click", () => {
    radarPlaying = !radarPlaying;
    syncRadarControls();
    startRadarPlayback();
  });
  radarControlsEl.querySelector("[data-radar-slider]")?.addEventListener("input", (event) => {
    radarPlaying = false;
    stopRadarPlayback();
    showRadarFrame(Number(event.target.value || 0), { immediate: true });
  });
  wrap.appendChild(weatherLayerSelectorEl);
  wrap.appendChild(weatherLegendEl);
  wrap.appendChild(radarControlsEl);
  syncRadarControls();
}

function destroyRadarControlsOnly() {
  [radarControlsEl, weatherLegendEl, weatherLayerSelectorEl].forEach((el) => el?.remove());
  radarControlsEl = null;
  weatherLegendEl = null;
  weatherLayerSelectorEl = null;
}

async function loadWeatherFramesForLayer(definition = getActiveWeatherLayerDefinition()) {
  if (definition.unavailable) {
    weatherFramesBySource[definition.id] = [];
    return [];
  }

  const orderedProviders = Object.values(WEATHER_PROVIDERS).sort((a, b) => a.priority - b.priority);
  const errors = [];
  for (const provider of orderedProviders) {
    try {
      const frames = await provider.loadFrames(definition);
      if (frames.length) {
        weatherFramesBySource[definition.id] = frames;
        return frames;
      }
      errors.push(`${provider.label}: ${WEATHER_UNAVAILABLE_MESSAGE}`);
    } catch (error) {
      errors.push(`${provider.label}: ${error?.message || WEATHER_UNAVAILABLE_MESSAGE}`);
      console.warn(`Provider meteo ${provider.label} non disponibile:`, error);
    }
  }
  console.warn("Nessun provider meteo disponibile:", errors.join(" | "));
  weatherFramesBySource[definition.id] = [];
  return [];
}

async function switchWeatherLayer(layerId) {
  if (!WEATHER_LAYER_DEFINITIONS[layerId] || activeWeatherLayerId === layerId) return;
  const loadToken = ++weatherLayerLoadToken;
  activeWeatherLayerId = layerId;
  stopRadarPlayback();
  if (radarLayer && fullscreenMap?.hasLayer(radarLayer)) fullscreenMap.removeLayer(radarLayer);
  radarLayer = null;
  radarFrames = getWeatherFramesForActiveLayer();
  syncRadarControls();
  syncWeatherLegend(radarFrames.length ? "" : "Caricamento...");
  if (!radarFrames.length) radarFrames = await loadWeatherFramesForLayer();
  if (loadToken !== weatherLayerLoadToken || activeWeatherLayerId !== layerId) return;
  radarFrameIndex = Math.max(0, Math.min(radarFrameIndex, radarFrames.length - 1));
  if (radarFrames.length) {
    showRadarFrame(radarFrameIndex, { immediate: true });
    syncWeatherLegend();
  } else {
    syncRadarControls();
    syncWeatherLegend(WEATHER_UNAVAILABLE_MESSAGE);
    setFullscreenFeedback(`${getActiveWeatherLayerDefinition().label}: ${WEATHER_UNAVAILABLE_MESSAGE}.`);
  }
  startRadarPlayback();
}

async function enableWeatherRadar() {
  if (!isMapFullscreenPageOpen || radarActive || radarLoading) return;
  const loadToken = ++weatherLayerLoadToken;
  radarLoading = true;
  updateRadarButtonState();
  try {
    ensureRadarPane();
    activeWeatherLayerId = "rain";
    radarActive = true;
    radarPlaying = true;
    createRadarControls();
    syncWeatherLegend("Caricamento...");
    radarFrames = await loadWeatherFramesForLayer();
    if (loadToken !== weatherLayerLoadToken) return;
    if (!radarFrames.length) throw new Error("Nessun layer meteo disponibile");
    radarFrameIndex = Math.max(0, radarFrames.length - 1);
    showRadarFrame(radarFrameIndex, { immediate: true });
    startRadarPlayback();
  } catch (error) {
    console.error("Errore layer meteo:", error);
    setFullscreenFeedback("Layer meteo non disponibile al momento.");
    destroyWeatherRadar();
  } finally {
    radarLoading = false;
    updateRadarButtonState();
  }
}

function destroyWeatherRadar() {
  weatherLayerLoadToken += 1;
  stopRadarPlayback();
  if (radarLayer && fullscreenMap?.hasLayer(radarLayer)) fullscreenMap.removeLayer(radarLayer);
  radarLayer = null;
  radarFrames = [];
  weatherFramesBySource = {};
  radarFrameIndex = 0;
  activeWeatherLayerId = "rain";
  radarActive = false;
  radarPlaying = true;
  radarLoading = false;
  destroyRadarControlsOnly();
  updateRadarButtonState();
}

function toggleWeatherRadar() {
  if (!isMapFullscreenPageOpen) return;
  if (radarActive || radarLoading) {
    destroyWeatherRadar();
    return;
  }
  enableWeatherRadar();
}

function applyFullscreenMapMode(mode) {
  const nextMode = ["standard", "satellite", "hybrid"].includes(mode) ? mode : "standard";
  const nextLayer = nextMode === "satellite"
    ? fullscreenSatelliteTileLayer
    : nextMode === "hybrid"
      ? fullscreenHybridTileLayer
      : fullscreenStandardTileLayer;
  Object.values(fullscreenBaseLayers).forEach((layer) => {
    if (fullscreenMap.hasLayer(layer) && layer !== nextLayer) fullscreenMap.removeLayer(layer);
  });
  if (!fullscreenMap.hasLayer(nextLayer)) nextLayer.addTo(fullscreenMap);
  fullscreenMapMode = nextMode;
  updateFullscreenMapModeButton();
}

function toggleFullscreenSatelliteMode() {
  applyFullscreenMapMode(fullscreenMapMode === "satellite" ? "standard" : "satellite");
  refreshFullscreenMapLayout();
}

function updateFullscreenMapModeButton() {
  if (!ui.mapSatelliteToggleBtn) return;
  const isSatellite = fullscreenMapMode === "satellite";
  ui.mapSatelliteToggleBtn.textContent = isSatellite ? "🗺 Standard" : "🛰 Satellite";
  ui.mapSatelliteToggleBtn.setAttribute("aria-pressed", isSatellite ? "true" : "false");
  ui.mapSatelliteToggleBtn.classList.toggle("is-active", isSatellite);
}

function refreshFullscreenMapLayout() {
  fullscreenMap.invalidateSize({ pan: false, animate: false });
  requestAnimationFrame(() => fullscreenMap.invalidateSize({ pan: false, animate: false }));
  setTimeout(() => fullscreenMap.invalidateSize({ pan: false, animate: false }), 220);
}

function setFullscreenFeedback(message) {
  if (ui.mapFullscreenFeedback) ui.mapFullscreenFeedback.textContent = message;
  if (ui.mapFullscreenFeedbackBanner) ui.mapFullscreenFeedbackBanner.classList.remove("hidden");
}

function toggleDrawAreaMode() {
  drawAreaModeActive = !drawAreaModeActive;
  if (drawAreaModeActive) {
    drawnAreaPoints = [];
    drawnAreaRedoStack = [];
    isDrawingStrokeActive = false;
    setFullscreenMapInteractivity(false);
    renderDrawnArea();
    ui.mapDrawAreaBtn.textContent = "✅ Termina";
    syncDrawAreaToolbarState();
    setFullscreenFeedback("Modalità disegno attiva: trascina il dito per tracciare l'area.");
    return;
  }
  setFullscreenMapInteractivity(true);
  ui.mapDrawAreaBtn.textContent = "✏️ Disegna";
  isDrawingStrokeActive = false;
  syncDrawAreaToolbarState();
  if (drawnAreaPoints.length < 3) {
    setFullscreenFeedback("Area non valida: servono almeno 3 punti.");
    return;
  }
  setFullscreenFeedback(`Area pronta (${drawnAreaPoints.length} punti). Puoi inoltrarla su WhatsApp.`);
  renderDrawnArea();
}

function setFullscreenMapInteractivity(enabled) {
  const actions = [
    fullscreenMap.dragging,
    fullscreenMap.touchZoom,
    fullscreenMap.doubleClickZoom,
    fullscreenMap.scrollWheelZoom,
    fullscreenMap.boxZoom,
    fullscreenMap.keyboard,
    fullscreenMap.tap
  ];
  actions.forEach((action) => {
    if (!action) return;
    if (enabled) action.enable();
    else action.disable();
  });
  const container = fullscreenMap.getContainer();
  if (!container) return;
  container.style.touchAction = enabled ? "pan-x pan-y" : "none";
  container.classList.toggle("map-fullscreen-view--drawing", !enabled);
}

function mapPointerEventToLatLng(event) {
  const rect = fullscreenMap.getContainer().getBoundingClientRect();
  const point = L.point(event.clientX - rect.left, event.clientY - rect.top);
  return fullscreenMap.containerPointToLatLng(point);
}

function onFullscreenMapPointerDown(event) {
  if (!drawAreaModeActive) return;
  event.preventDefault();
  if (event.pointerId !== undefined && event.currentTarget?.setPointerCapture) {
    event.currentTarget.setPointerCapture(event.pointerId);
  }
  isDrawingStrokeActive = true;
  drawnAreaPoints = [];
  drawnAreaRedoStack = [];
  const latLng = mapPointerEventToLatLng(event);
  if (latLng) drawnAreaPoints.push([latLng.lat, latLng.lng]);
  syncDrawAreaToolbarState();
  renderDrawnArea();
}

function onFullscreenMapPointerMove(event) {
  if (!drawAreaModeActive || !isDrawingStrokeActive) return;
  event.preventDefault();
  const latLng = mapPointerEventToLatLng(event);
  if (!latLng) return;
  const lastPoint = drawnAreaPoints[drawnAreaPoints.length - 1];
  if (lastPoint) {
    const distance = fullscreenMap.distance(L.latLng(lastPoint[0], lastPoint[1]), latLng);
    if (distance < 2) return;
  }
  drawnAreaPoints.push([latLng.lat, latLng.lng]);
  renderDrawnArea();
}

function onFullscreenMapPointerUp(event) {
  if (!drawAreaModeActive || !isDrawingStrokeActive) return;
  event.preventDefault();
  if (event.pointerId !== undefined && event.currentTarget?.releasePointerCapture) {
    try {
      event.currentTarget.releasePointerCapture(event.pointerId);
    } catch (error) {
      console.warn("Pointer capture già rilasciato", error);
    }
  }
  isDrawingStrokeActive = false;
  if (drawnAreaPoints.length >= 3) {
    const first = drawnAreaPoints[0];
    const last = drawnAreaPoints[drawnAreaPoints.length - 1];
    const closingDistance = fullscreenMap.distance(L.latLng(first[0], first[1]), L.latLng(last[0], last[1]));
    if (closingDistance > 1) drawnAreaPoints.push([first[0], first[1]]);
  }
  syncDrawAreaToolbarState();
  renderDrawnArea();
}

function syncDrawAreaToolbarState() {
  if (ui.mapShareAreaWhatsappBtn) ui.mapShareAreaWhatsappBtn.disabled = drawnAreaPoints.length < 3;
  if (ui.mapDrawUndoBtn) ui.mapDrawUndoBtn.disabled = drawnAreaPoints.length < 2;
  if (ui.mapDrawRedoBtn) ui.mapDrawRedoBtn.disabled = drawnAreaRedoStack.length < 2;
  if (ui.mapDrawClearBtn) ui.mapDrawClearBtn.disabled = drawnAreaPoints.length === 0;
}

function undoDrawnArea() {
  if (drawnAreaPoints.length < 2) return;
  drawnAreaRedoStack = [...drawnAreaPoints];
  drawnAreaPoints = [];
  syncDrawAreaToolbarState();
  renderDrawnArea();
  setFullscreenFeedback("Disegno annullato. Premi “Rifai” per ripristinarlo.");
}

function redoDrawnArea() {
  if (drawnAreaRedoStack.length < 2) return;
  drawnAreaPoints = [...drawnAreaRedoStack];
  drawnAreaRedoStack = [];
  syncDrawAreaToolbarState();
  renderDrawnArea();
  setFullscreenFeedback("Disegno ripristinato.");
}

function clearDrawnArea() {
  if (!drawnAreaPoints.length) return;
  drawnAreaPoints = [];
  drawnAreaRedoStack = [];
  syncDrawAreaToolbarState();
  renderDrawnArea();
  setFullscreenFeedback("Disegno cancellato.");
}

function renderDrawnArea() {
  fullscreenDrawLayer.clearLayers();
  if (!drawnAreaPoints.length) return;
  if (drawnAreaPoints.length >= 2) {
    L.polyline(drawnAreaPoints, { color: "#dc2626", weight: 4, lineCap: "round", lineJoin: "round" }).addTo(fullscreenDrawLayer);
  }
  if (drawnAreaPoints.length >= 3) {
    L.polygon(drawnAreaPoints, { color: "#b91c1c", fillColor: "#ef4444", fillOpacity: 0.12, weight: 3 }).addTo(fullscreenDrawLayer);
  }
}

function shareDrawnAreaViaWhatsapp() {
  if (drawnAreaPoints.length < 3) {
    alert("Disegna almeno 3 punti per creare un'area condivisibile.");
    return;
  }
  const vertices = drawnAreaPoints.map((point, idx) => `• Punto ${idx + 1}: ${point[0].toFixed(6)}, ${point[1].toFixed(6)}`).join("\n");
  const areaCoords = drawnAreaPoints.map((point) => `${point[0].toFixed(6)},${point[1].toFixed(6)}`).join(" | ");
  const message = [
    "🗺️ *Area lavoro commessa*",
    `Commessa: ${selectedCommessaName || "-"}`,
    `Perimetro: ${drawnAreaPoints.length} punti`,
    "",
    vertices,
    "",
    `Tracciato compatto: ${areaCoords}`,
    "",
    "Apri da Google Maps (primo punto):",
    `https://www.google.com/maps/search/?api=1&query=${drawnAreaPoints[0][0]},${drawnAreaPoints[0][1]}`
  ].join("\n");
  if (!safeOpenWhatsAppMessage(message)) alert("Impossibile aprire WhatsApp su questo dispositivo.");
}

function parseCommessaHash(hash = window.location.hash || "") {
  const rawHash = String(hash || "").replace(/^#/, "");
  if (!rawHash.startsWith("commessa=")) return { id: "", resource: "", notes: false, impianto: "", meteo: "", atex: "", safety: "" };
  const params = new URLSearchParams(rawHash);
  return {
    id: params.get("commessa") || "",
    resource: params.get("resource") || "",
    notes: params.has("notes"),
    impianto: params.get("impianto") || "",
    meteo: params.get("meteo") || "",
    atex: params.get("atex") || "",
    safety: params.get("safety") || ""
  };
}

function setCommessaHash(suffix = "") {
  if (!selectedCommessaId) return;
  window.location.hash = `commessa=${encodeURIComponent(selectedCommessaId)}${suffix}`;
}

function focusSharedImpiantoFromRoute(impiantoKey) {
  const key = String(impiantoKey || "").trim();
  if (!key) return;
  const impianto = findCurrentImpiantoByKey(key);
  if (!impianto) return;
  focusImpiantoInList(impianto, true);
  selectImpiantoForMapDetail(impianto);
}

function applyRoute() {
  const hash = window.location.hash || "";
  const commessaRoute = parseCommessaHash(hash);
  const fuelMatch = hash.match(/^#fuel=(.+)$/);
  const showSegnalazioni = hash === "#segnalazioni";
  const showHowto = hash === "#howto";
  const showPrivateDocs = hash === "#documenti";
  const showPos = hash === "#pos" || (window.location.pathname === "/pos" && !hash);
  const personalServiceMatch = hash.match(/^#servizi-personali(?:=([a-z]+))?$/);
  const showHours = hash === "#ore";
  const commessaIdFromHash = commessaRoute.id;
  const resourceTypeFromHash = commessaRoute.resource;
  const showFuel = Boolean(fuelMatch);
  const showPersonalServices = Boolean(personalServiceMatch);
  const showNotesPage = Boolean(commessaRoute.notes && selectedCommessaId === commessaIdFromHash);
  const showWeatherDetail = Boolean(commessaRoute.meteo && selectedCommessaId === commessaIdFromHash && !showNotesPage);
  const showAtexProcedure = Boolean(commessaRoute.atex && selectedCommessaId === commessaIdFromHash && !showNotesPage && !showWeatherDetail);
  const showImpiantoSafety = Boolean(commessaRoute.safety && selectedCommessaId === commessaIdFromHash && !showNotesPage && !showWeatherDetail && !showAtexProcedure);
  const showImpianti = Boolean(commessaIdFromHash && selectedCommessaId === commessaIdFromHash && !showNotesPage && !showWeatherDetail && !showAtexProcedure && !showImpiantoSafety);
  const showResourceViewer = Boolean(showImpianti && resourceTypeFromHash);
  ui.homePage.classList.toggle("hidden", showImpianti || showNotesPage || showWeatherDetail || showAtexProcedure || showImpiantoSafety || showFuel || showSegnalazioni || showHowto || showPrivateDocs || showPos || showHours || showPersonalServices);
  ui.impiantiPage.classList.toggle("hidden", !showImpianti || isMapFullscreenPageOpen);
  ui.impiantoWeatherDetailPage?.classList.toggle("hidden", !showWeatherDetail);
  ui.atexProcedurePage?.classList.toggle("hidden", !showAtexProcedure);
  ui.impiantoSafetyPage?.classList.toggle("hidden", !showImpiantoSafety);
  ui.commessaNotesPage?.classList.toggle("hidden", !showNotesPage);
  ui.mapFullscreenPage?.classList.toggle("hidden", !isMapFullscreenPageOpen);
  ui.fuelPage.classList.toggle("hidden", !showFuel);
  ui.personalServicesPage.classList.toggle("hidden", !showPersonalServices);
  ui.segnalazioniPage.classList.toggle("hidden", !showSegnalazioni);
  ui.howtoPage.classList.toggle("hidden", !showHowto);
  ui.privateDocsPage.classList.toggle("hidden", !showPrivateDocs);
  ui.posPage?.classList.toggle("hidden", !showPos);
  ui.hoursPage.classList.toggle("hidden", !showHours);
  document.body.classList.toggle("resource-view-open", showResourceViewer);
  ui.mapFullscreenBtn.classList.toggle("hidden", showResourceViewer);
  ui.commessaNotesToggleBtn?.classList.toggle("hidden", showResourceViewer);
  ui.commessaWeatherRefreshBtn?.classList.toggle("hidden", showResourceViewer);
  ui.commessaWeatherRefreshStatus?.classList.toggle("hidden", showResourceViewer);
  ui.commessaNotesCard?.classList.toggle("hidden", showResourceViewer);
  const mapElement = document.getElementById("map");
  if (mapElement) mapElement.classList.toggle("hidden", showResourceViewer);
  if (ui.gpsStatus) ui.gpsStatus.classList.toggle("hidden", showResourceViewer);
  const impiantiCard = ui.impiantiLista?.closest(".card");
  if (impiantiCard) impiantiCard.classList.toggle("hidden", showResourceViewer);
  if (showNotesPage) {
    renderCommessaNotes();
  }
  if (showWeatherDetail) {
    renderDettaglioMeteoImpianto(commessaRoute.meteo);
  }
  if (showAtexProcedure) {
    renderAtexProcedurePage(commessaRoute.atex);
  }
  if (showImpiantoSafety) {
    renderImpiantoSafetyPage(commessaRoute.safety);
  }
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
    setTimeout(() => {
      map.invalidateSize();
      if (commessaRoute.impianto) focusSharedImpiantoFromRoute(commessaRoute.impianto);
    }, 50);
  }
  if (showHowto) renderHowtoFaq();
  if (showPrivateDocs) renderPrivateDocsList();
  if (showPos) renderPosDocuments();
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

function openImpiantiPage(suffix = "") {
  if (!selectedCommessaId) return;
  localStorage.setItem(LAST_OPENED_COMMESSA_KEY, selectedCommessaId);
  setCommessaHash(suffix);
  applyRoute();
}

function openCommessaNotesPage() {
  if (!selectedCommessaId) return;
  closeCommessaResourceViewer();
  setCommessaHash("&notes");
  renderCommessaNotes();
  applyRoute();
}

function openCommessaNotesPage() {
  if (!selectedCommessaId) return;
  closeCommessaResourceViewer();
  window.location.hash = `commessa=${selectedCommessaId}&notes`;
  renderCommessaNotes();
  applyRoute();
}

function closeImpiantiPage() {
  closeMapFullscreenPage();
  localStorage.removeItem(LAST_OPENED_COMMESSA_KEY);
  window.location.hash = "";
  ui.exportCurrentCommessaBtn.disabled = true;
  setCommessaWeatherRefreshStatus("");
  updateCommessaWeatherRefreshButtonState();
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
  const commessaRoute = parseCommessaHash(routeHash);
  const hasOpenCommessaRoute = hasSelectedCommessa && commessaRoute.id === selectedCommessaId && !commessaRoute.notes && !commessaRoute.resource;
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

function openBookPdf() {
  closeSideMenu();
  const bookUrl = "./docs/Libro_Completo_Hera_App.pdf";
  const opened = window.open(bookUrl, "_blank", "noopener,noreferrer");
  if (!opened) {
    window.location.href = bookUrl;
  }
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

function openPrivateDocsUploadPage() {
  openPrivateDocsPage();
  applyPrivateDocPreset("pin");
  setTimeout(() => {
    ui.privateDocsForm?.scrollIntoView({ behavior: "smooth", block: "start" });
    ui.privateDocsName?.focus();
  }, 50);
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
  if (ui.hoursDate) ui.hoursDate.value = new Date().toISOString().slice(0, 10);
  if (!ui.hoursStatsMonth?.value) ui.hoursStatsMonth.value = new Date().toISOString().slice(0, 7);
  if (!ui.hoursCommesseList.children.length) addHoursCommessaBlock();
  Array.from(ui.hoursCommesseList.querySelectorAll(".hours-commessa-card")).forEach((card) => {
    applyHoursSuggestedOperators(card, { force: true });
  });
  renderHoursTableCommessaOptions();
  window.location.hash = "ore";
  applyRoute();
  closeSideMenu();
}

function closeHoursPage() {
  window.location.hash = "";
  applyRoute();
}

function openPosPage() {
  if (window.location.pathname !== "/pos" || window.location.hash) {
    window.history.pushState({}, "", "/pos");
  }
  applyRoute();
  closeSideMenu();
}

function closePosPage() {
  if (window.location.pathname === "/pos") {
    window.history.pushState({}, "", "/");
  } else {
    window.location.hash = "";
  }
  applyRoute();
}

function stopPosDocumentsSubscription() {
  if (unsubscribePosDocuments) {
    unsubscribePosDocuments();
    unsubscribePosDocuments = null;
  }
  posDocuments = [];
}

function subscribePosDocuments() {
  stopPosDocumentsSubscription();
  const query = canManageData()
    ? db.collection("posDocuments")
    : db.collection("posDocuments").where("active", "==", true);
  unsubscribePosDocuments = query.onSnapshot((snapshot) => {
    posDocuments = snapshot.docs.map((doc) => ({ id: doc.id, ...doc.data() }));
    renderPosDocuments();
  }, (error) => {
    console.error("Errore caricamento documenti POS:", error);
    if (ui.posDocumentsList) ui.posDocumentsList.innerHTML = "<p class='muted'>Impossibile caricare i documenti POS.</p>";
  });
}

function getFilteredPosDocuments() {
  const canManage = canManageData();
  const search = String(ui.posSearch?.value || "").trim().toLowerCase();
  return posDocuments
    .filter((doc) => canManage || doc.active === true)
    .filter((doc) => {
      if (!search) return true;
      return [doc.title, doc.description, doc.category]
        .some((value) => String(value || "").toLowerCase().includes(search));
    })
    .sort((a, b) => {
      const categoryCompare = String(a.category || "Altro").localeCompare(String(b.category || "Altro"), "it");
      if (categoryCompare !== 0) return categoryCompare;
      const orderCompare = Number(a.order || 0) - Number(b.order || 0);
      if (orderCompare !== 0) return orderCompare;
      return String(a.title || "").localeCompare(String(b.title || ""), "it");
    });
}

function renderPosDocuments() {
  if (!ui.posDocumentsList) return;
  const canManage = canManageData();
  updateDriveConnectVisibility();
  ui.openPosBtn?.classList.remove("hidden");
  if (ui.openPosBtn) ui.openPosBtn.disabled = false;
  ui.posAdminCard?.classList.toggle("hidden", !canManage);
  const documents = getFilteredPosDocuments();
  if (!documents.length) {
    ui.posDocumentsList.innerHTML = "<p class='muted'>Nessun documento disponibile.</p>";
    return;
  }
  ui.posDocumentsList.innerHTML = "";
  const grouped = new Map();
  documents.forEach((doc) => {
    const category = String(doc.category || "Altro").trim() || "Altro";
    if (!grouped.has(category)) grouped.set(category, []);
    grouped.get(category).push(doc);
  });
  grouped.forEach((items, category) => {
    const group = document.createElement("section");
    group.className = "pos-category-group";
    group.innerHTML = `<h3>📁 ${escapeHTML(category)}</h3>`;
    const grid = document.createElement("div");
    grid.className = "pos-document-grid";
    items.forEach((doc) => grid.appendChild(createPosDocumentCard(doc, canManage)));
    group.appendChild(grid);
    ui.posDocumentsList.appendChild(group);
  });
}

function createPosDocumentCard(doc, canManage) {
  const card = document.createElement("article");
  card.className = "pos-document-card";
  if (doc.active === false) card.classList.add("is-inactive");
  const title = document.createElement("h4");
  title.textContent = doc.title || "Documento senza titolo";
  const description = document.createElement("p");
  description.className = "muted";
  description.textContent = doc.description || "Nessuna descrizione.";
  const actions = document.createElement("div");
  actions.className = "item-actions pos-document-actions";
  const driveUrl = String(doc.driveUrl || "").trim();
  if (driveUrl) {
    const link = document.createElement("a");
    link.className = "btn pos-open-link";
    link.href = driveUrl;
    link.target = "_blank";
    link.rel = "noopener noreferrer";
    link.textContent = "Apri documento";
    actions.appendChild(link);
  } else {
    const unavailable = document.createElement("p");
    unavailable.className = "muted pos-unavailable";
    unavailable.textContent = "Documento non disponibile.";
    actions.appendChild(unavailable);
  }
  if (canManage) {
    const editBtn = createButton("Modifica", () => openPosDocumentForm(doc));
    const deleteBtn = createButton("Elimina", () => deletePosDocument(doc));
    actions.appendChild(editBtn);
    actions.appendChild(deleteBtn);
    const meta = document.createElement("p");
    meta.className = "muted pos-admin-meta";
    meta.textContent = `Ordine: ${Number(doc.order || 0)} • ${doc.active === false ? "Non attivo" : "Attivo"}`;
    card.append(title, description, meta, actions);
    return card;
  }
  card.append(title, description, actions);
  return card;
}

function openPosDocumentForm(doc = null) {
  if (!canManageData()) return;
  ui.posDocumentForm?.classList.remove("hidden");
  if (ui.posAddToggleBtn) ui.posAddToggleBtn.textContent = doc ? "Modifica documento" : "➕ Aggiungi documento";
  if (ui.posDocumentId) ui.posDocumentId.value = doc?.id || "";
  if (ui.posTitle) ui.posTitle.value = doc?.title || "";
  if (ui.posDescription) ui.posDescription.value = doc?.description || "";
  if (ui.posDriveUrl) ui.posDriveUrl.value = doc?.driveUrl || "";
  if (ui.posCategory) ui.posCategory.value = doc?.category || POS_DEFAULT_CATEGORIES[0];
  if (ui.posOrder) ui.posOrder.value = Number(doc?.order || 0);
  if (ui.posActive) ui.posActive.checked = doc?.active !== false;
  ui.posTitle?.focus();
}

function closePosDocumentForm() {
  ui.posDocumentForm?.reset();
  if (ui.posDocumentId) ui.posDocumentId.value = "";
  if (ui.posActive) ui.posActive.checked = true;
  if (ui.posAddToggleBtn) ui.posAddToggleBtn.textContent = "➕ Aggiungi documento";
  ui.posDocumentForm?.classList.add("hidden");
  if (ui.posFeedback) ui.posFeedback.textContent = "";
}

async function savePosDocument(event) {
  event.preventDefault();
  if (!canManageData()) {
    alert("Solo l'admin può salvare documenti POS.");
    return;
  }
  const id = String(ui.posDocumentId?.value || "").trim();
  const now = firebase.firestore.FieldValue.serverTimestamp();
  const payload = {
    title: String(ui.posTitle?.value || "").trim(),
    description: String(ui.posDescription?.value || "").trim(),
    driveUrl: String(ui.posDriveUrl?.value || "").trim(),
    category: String(ui.posCategory?.value || "").trim() || "Altro",
    order: Number(ui.posOrder?.value || 0),
    active: Boolean(ui.posActive?.checked),
    updatedAt: now
  };
  if (!payload.title) {
    alert("Inserisci il titolo documento.");
    return;
  }
  if (id) {
    await db.collection("posDocuments").doc(id).set(payload, { merge: true });
  } else {
    await db.collection("posDocuments").add({
      ...payload,
      createdAt: now,
      createdBy: currentUser?.email || ""
    });
  }
  if (ui.posFeedback) ui.posFeedback.textContent = "Documento salvato.";
  closePosDocumentForm();
}

async function deletePosDocument(doc) {
  if (!canManageData()) {
    alert("Solo l'admin può eliminare documenti POS.");
    return;
  }
  const ok = window.confirm(`Eliminare il documento "${doc.title || "senza titolo"}"?`);
  if (!ok) return;
  await db.collection("posDocuments").doc(doc.id).delete();
}

function renderSavedHoursReports(records = []) {
  if (!ui.hoursSavedList) return;
  if (!records.length) {
    ui.hoursSavedList.innerHTML = "<p class='muted'>Nessun report ore salvato per i filtri correnti.</p>";
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
    await ensureHoursReportsDeduplicated();
    const baseQuery = db.collection("oreReports");
    const snapshot = await baseQuery.orderBy("createdAt", "desc").limit(100).get();
    const reports = deduplicateHoursRecordsForDisplay(snapshot.docs.map((doc) => ({ id: doc.id, ...doc.data() })));
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

function setHoursConfirmVisibleButtonState(show, disabled = false) {
  if (!ui.hoursConfirmVisibleBtn) return;
  ui.hoursConfirmVisibleBtn.classList.toggle("hidden", !show);
  ui.hoursConfirmVisibleBtn.disabled = Boolean(disabled);
}

function openHoursConfirmModal({ title = "Confermare ore?", text = "Vuoi confermare le ore?", confirmLabel = "Conferma ore" } = {}) {
  if (!ui.hoursConfirmModal) return Promise.resolve(window.confirm(text));
  if (ui.hoursConfirmTitle) ui.hoursConfirmTitle.textContent = title;
  if (ui.hoursConfirmText) ui.hoursConfirmText.textContent = text;
  if (ui.hoursConfirmOkBtn) ui.hoursConfirmOkBtn.textContent = confirmLabel;
  ui.hoursConfirmModal.classList.remove("hidden");
  ui.hoursConfirmModal.setAttribute("aria-hidden", "false");
  ui.hoursConfirmOkBtn?.focus();
  return new Promise((resolve) => {
    hoursConfirmModalResolve = resolve;
  });
}

function closeHoursConfirmModal(confirmed) {
  if (!ui.hoursConfirmModal || ui.hoursConfirmModal.classList.contains("hidden")) return;
  ui.hoursConfirmModal.classList.add("hidden");
  ui.hoursConfirmModal.setAttribute("aria-hidden", "true");
  const resolve = hoursConfirmModalResolve;
  hoursConfirmModalResolve = null;
  if (resolve) resolve(Boolean(confirmed));
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

async function fetchHoursReportsForMonth(monthValue, monthMeta, options = {}) {
  if (!monthMeta) return [];
  const includePendingApprovals = options?.includePendingApprovals === true;
  const fromDate = `${monthValue}-01`;
  const toDate = `${monthValue}-${String(monthMeta.daysInMonth).padStart(2, "0")}`;
  const reportsQuery = db.collection("oreReports")
    .where("date", ">=", fromDate)
    .where("date", "<=", toDate)
    .orderBy("date", "asc")
    .get();
  const approvalsQuery = includePendingApprovals
    ? db.collection("oreApprovalRequests")
      .where("date", ">=", fromDate)
      .where("date", "<=", toDate)
      .orderBy("date", "asc")
      .get()
    : Promise.resolve(null);
  const [reportsSnapshot, approvalsSnapshot] = await Promise.all([reportsQuery, approvalsQuery]);
  const reports = reportsSnapshot.docs.map((doc) => ({
    id: doc.id,
    sourceCollection: "oreReports",
    approvalStatus: "approved",
    ...doc.data()
  }));
  const pendingApprovals = approvalsSnapshot
    ? approvalsSnapshot.docs
      .map((doc) => ({
        id: doc.id,
        sourceCollection: "oreApprovalRequests",
        ...doc.data()
      }))
      .filter((request) => !["approved", "rejected"].includes(String(request.status || "").trim()))
    : [];
  return deduplicateHoursRecordsForDisplay([...reports, ...pendingApprovals])
    .sort((a, b) => String(a.date || "").localeCompare(String(b.date || "")));
}

function ensureHoursViewModalOpen() {
  if (!ui.hoursViewModal || !ui.hoursViewModal.classList.contains("hidden")) return;
  ui.hoursViewModal.classList.remove("hidden");
  ui.hoursViewModal.setAttribute("aria-hidden", "false");
}

function logHoursDebug(label, value) {
  console.log(`[ORE] ${label}:`, value);
}

function getSelectedHoursCommessaInfo(commessaId) {
  const commessa = commesseById.get(String(commessaId || ""));
  return {
    id: String(commessaId || "").trim(),
    nome: String(commessa?.nome || "").trim(),
    codice: String(commessa?.codice || "").trim()
  };
}

function normalizeHoursCommessaMatchValue(value) {
  return String(value || "")
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .toLocaleLowerCase("it-IT")
    .replace(/[^a-z0-9]+/g, " ")
    .replace(/\s+/g, " ")
    .trim();
}

function resolveHoursEntryCommessa(entry = {}) {
  const directId = String(entry?.commessaId || "").trim();
  if (directId && commesseById.has(directId)) {
    const commessa = commesseById.get(directId) || {};
    return { id: directId, nome: String(commessa.nome || entry.commessaName || "Commessa").trim(), codice: String(commessa.codice || entry.commessaCode || entry.codice || "").trim(), key: directId };
  }
  const candidates = [entry?.commessaId, entry?.commessaCode, entry?.codice, entry?.commessaName]
    .map(normalizeHoursCommessaMatchValue)
    .filter(Boolean);
  const matched = Array.from(commesseById.values()).find((commessa) => {
    const values = [commessa.id, commessa.codice, commessa.nome, getCommessaDisplayName(commessa)]
      .map(normalizeHoursCommessaMatchValue)
      .filter(Boolean);
    return candidates.some((candidate) => values.includes(candidate));
  });
  if (matched?.id) {
    return { id: matched.id, nome: String(matched.nome || entry.commessaName || "Commessa").trim(), codice: String(matched.codice || entry.commessaCode || entry.codice || "").trim(), key: matched.id };
  }
  const fallbackKey = String(entry?.commessaId || entry?.commessaCode || entry?.codice || entry?.commessaName || "").trim();
  return {
    id: directId,
    nome: String(entry?.commessaName || fallbackKey || "Commessa").trim(),
    codice: String(entry?.commessaCode || entry?.codice || "").trim(),
    key: fallbackKey
  };
}

function doesHoursEntryMatchCommessa(entry, selectedCommessaId) {
  const selected = getSelectedHoursCommessaInfo(selectedCommessaId);
  const resolved = resolveHoursEntryCommessa(entry);
  if (selected.id && resolved.id && selected.id === resolved.id) return true;
  const selectedValues = [selected.id, selected.codice, selected.nome]
    .map(normalizeHoursCommessaMatchValue)
    .filter(Boolean);
  const entryValues = [resolved.id, resolved.key, resolved.codice, resolved.nome, entry?.commessaId, entry?.commessaCode, entry?.codice, entry?.commessaName]
    .map(normalizeHoursCommessaMatchValue)
    .filter(Boolean);
  return selectedValues.some((value) => entryValues.includes(value));
}
function setHoursExportButtonsLoading(isLoading) {
  const monthlyWithoutRows = hoursTableContext?.mode === "monthly" && !hoursTableContext?.operators?.length;
  if (ui.hoursTableExportBtn) ui.hoursTableExportBtn.disabled = isLoading || !hoursTableContext || monthlyWithoutRows;
  if (ui.hoursTableExportGlobalBtn) ui.hoursTableExportGlobalBtn.disabled = isLoading;
}

function buildHoursMonthlyExportData(reports, commessaId, monthMeta) {
  const operatorDayMap = new Map();
  const operatorTotals = new Map();
  const operatorCommessaTotals = new Map();
  (Array.isArray(reports) ? reports : []).forEach((report) => {
    const day = Number(String(report.date || "").split("-")[2] || 0);
    const entries = Array.isArray(report.entries) ? report.entries : [];
    entries.forEach((entry) => {
      const entryCommessaInfo = resolveHoursEntryCommessa(entry);
      const entryCommessaId = String(entryCommessaInfo.id || entryCommessaInfo.key || "").trim();
      const entryCommessaName = String(entryCommessaInfo.nome || commesseById.get(entryCommessaId)?.nome || "Commessa").trim();
      (Array.isArray(entry.rows) ? entry.rows : []).forEach((row) => {
        const operatore = String(row.operatore || "").trim();
        const ore = Number(row.ore || 0);
        if (!operatore || ore <= 0) return;
        const operatorNorm = operatore.toLocaleLowerCase("it-IT").replace(/\s+/g, " ").trim();
        if (!operatorTotals.has(operatorNorm)) operatorTotals.set(operatorNorm, { name: operatore, total: 0 });
        operatorTotals.get(operatorNorm).total += ore;
        const byCommessaKey = `${operatorNorm}__${entryCommessaId || entryCommessaName}`;
        if (!operatorCommessaTotals.has(byCommessaKey)) operatorCommessaTotals.set(byCommessaKey, { operatore, commessaName: entryCommessaName, total: 0 });
        operatorCommessaTotals.get(byCommessaKey).total += ore;
        if (!doesHoursEntryMatchCommessa(entry, commessaId) || !day || day < 1 || day > monthMeta.daysInMonth) return;
        if (!operatorDayMap.has(operatore)) operatorDayMap.set(operatore, Array(monthMeta.daysInMonth).fill(0));
        operatorDayMap.get(operatore)[day - 1] += ore;
      });
    });
  });
  return { operatorDayMap, operatorTotals, operatorCommessaTotals };
}

async function loadHoursMonthlyTable() {
  if (!ui.hoursTableFeedback || !ui.hoursTableContainer) return null;
  const requestId = hoursTableLoadRequestId + 1;
  hoursTableLoadRequestId = requestId;
  hoursTableContext = null;
  setHoursConfirmVisibleButtonState(false);
  loadingOre = true;
  if (ui.hoursTableCommessaSelect) ui.hoursTableCommessaSelect.disabled = false;
  renderHoursTableCommessaButtons();
  setHoursExportButtonsLoading(true);
  const monthValue = String(ui.hoursTableMonth?.value || "").trim();
  const commessaId = String(ui.hoursTableCommessaSelect?.value || "").trim();
  const commessaInfo = getSelectedHoursCommessaInfo(commessaId);
  const monthMeta = getMonthMeta(monthValue);
  logHoursDebug("mese selezionato", monthValue);
  logHoursDebug("anno", monthMeta?.year || "non valido");
  logHoursDebug("mese numerico", monthMeta?.month || "non valido");
  logHoursDebug("commessa selezionata", commessaInfo.codice || commessaInfo.nome || commessaInfo.id || "nessuna");
  if (!monthMeta) {
    ui.hoursTableFeedback.textContent = "Seleziona un mese valido.";
    ui.hoursTableContainer.innerHTML = "";
    loadingOre = false;
    setHoursExportButtonsLoading(false);
    return null;
  }
  if (!commessaId) {
    ui.hoursTableFeedback.textContent = "Seleziona una commessa per vedere la tabella.";
    ui.hoursTableContainer.innerHTML = "";
    loadingOre = false;
    setHoursExportButtonsLoading(false);
    return null;
  }
  if (ui.hoursStatsMonth) ui.hoursStatsMonth.value = monthValue;
  ui.hoursTableFeedback.textContent = "Caricamento tabella ore...";
  ui.hoursTableContainer.innerHTML = "";
  const loadPromise = (async () => {
    try {
      const reports = await fetchHoursReportsForMonth(monthValue, monthMeta, { includePendingApprovals: true });
      if (requestId !== hoursTableLoadRequestId) return null;
      logHoursDebug("record ore trovati", Array.isArray(reports) ? reports.length : 0);
      const context = renderHoursMonthlyTable(reports, commessaId, monthMeta, { monthValue });
      logHoursDebug("dati usati per tabella", context);
      return context;
    } catch (error) {
      if (requestId === hoursTableLoadRequestId) {
        console.error("Errore caricamento tabella mensile ore:", error);
        ui.hoursTableFeedback.textContent = "Errore caricamento ore. Controlla i dati o riprova.";
        ui.hoursTableContainer.innerHTML = "";
      }
      return null;
    } finally {
      loadingOre = false;
      if (requestId === hoursTableLoadRequestId) setHoursExportButtonsLoading(false);
    }
  })();
  hoursTableLoadPromise = loadPromise;
  return loadPromise;
}

async function loadHoursTotalByOperator() {
  if (!ui.hoursTableFeedback || !ui.hoursTableContainer) return;
  hoursTableContext = null;
  setHoursConfirmVisibleButtonState(false);
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
  renderHoursTableCommessaButtons();
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
  setHoursConfirmVisibleButtonState(false);
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
  renderHoursTableCommessaButtons();
  ensureHoursViewModalOpen();
  ui.hoursTableFeedback.textContent = "Caricamento totale ore operatore per commessa...";
  ui.hoursTableContainer.innerHTML = "";
  try {
    const reports = await fetchHoursReportsForMonth(monthValue, monthMeta);
    const totals = new Map();
    reports.forEach((report) => {
      const entries = Array.isArray(report.entries) ? report.entries : [];
      entries.forEach((entry) => {
        const entryCommessaInfo = resolveHoursEntryCommessa(entry);
        const commessaId = String(entryCommessaInfo.id || entryCommessaInfo.key || "").trim();
        const commessaName = String(entryCommessaInfo.nome || entry.commessaName || commesseById.get(commessaId)?.nome || "Commessa").trim();
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

function renderHoursMonthlyTable(reports, commessaId, monthMeta, options = {}) {
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
      const entryCommessaInfo = resolveHoursEntryCommessa(entry);
      if (!doesHoursEntryMatchCommessa(entry, commessaId)) return;
      (Array.isArray(entry.rows) ? entry.rows : []).forEach((row, rowIndex) => {
        const operatore = String(row.operatore || "").trim();
        const ore = Number(row.ore || 0);
        if (!operatore || ore <= 0) return;
        if (!operatorsMap.has(operatore)) {
          operatorsMap.set(operatore, Array.from({ length: monthMeta.daysInMonth }, () => []));
        }
        const isPendingApproval = String(report.sourceCollection || "oreReports") === "oreApprovalRequests";
        operatorsMap.get(operatore)[day - 1].push({ ore, isPendingApproval });
        const key = `${operatore}__${day}`;
        if (!hoursTableRowsMap.has(key)) hoursTableRowsMap.set(key, []);
        hoursTableRowsMap.get(key).push({
          recordId: report.id,
          reportId: report.id,
          sourceCollection: report.sourceCollection || "oreReports",
          approvalStatus: report.status || report.approvalStatus || "approved",
          reportDate,
          monthValue: `${monthMeta.year}-${String(monthMeta.month).padStart(2, "0")}`,
          year: monthMeta.year,
          month: monthMeta.month,
          entryCommessaId: entryCommessaInfo.id || entryCommessaInfo.key || entry.commessaId,
          entryCommessaName: entryCommessaInfo.nome || entry.commessaName || commesseById.get(entryCommessaInfo.id || entryCommessaInfo.key)?.nome || "Commessa",
          cellKey: key,
          rowUniqueKey: row.uniqueKey || buildHoursUniqueKey(reportDate, entryCommessaInfo.id || entryCommessaInfo.key || entry.commessaId, row),
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
    const getDayItemHours = (item) => Number(typeof item === "object" ? item.ore : item || 0);
    const total = dayValues.reduce((sum, dayItems) => sum + dayItems.reduce((daySum, item) => daySum + getDayItemHours(item), 0), 0);
    const cells = dayValues.map((dayItems, idx) => {
      const day = idx + 1;
      if (!dayItems.length) return "<td>-</td>";
      const key = `${operatorName}__${day}`;
      const sources = hoursTableRowsMap.get(key) || [];
      const pendingSources = sources.filter((source) => String(source.sourceCollection || "oreReports") === "oreApprovalRequests");
      const hasPendingApproval = pendingSources.length > 0;
      const canManage = canManageData() && sources.length;
      const dayTotal = dayItems.reduce((sum, value) => sum + getDayItemHours(value), 0);
      const hasDuplicates = dayItems.length > 1;
      const hasDataError = sources.some((source) => !source.reportDate || !source.entryCommessaId || !source.operatore || Number(source.ore || 0) <= 0);
      let valueLabel = `✅ ${formatHoursValue(dayTotal)}h · ore inserite`;
      let statusClass = "hours-value-ok";
      if (hasDataError) {
        valueLabel = `❌ ${formatHoursValue(dayTotal)}h · errore dati`;
        statusClass = "hours-value-error";
      } else if (hasDuplicates) {
        valueLabel = `⚠️ ${formatHoursValue(dayTotal)}h · duplicato da controllare`;
        statusClass = "hours-value-warning";
      } else if (hasPendingApproval) {
        valueLabel = `⚠️ ${formatHoursValue(dayTotal)}h · da confermare`;
        statusClass = "hours-value-warning hours-value-pending-approval";
      }
      const mergedDetails = hasDuplicates
        ? `Duplicato non valido: stesso operatore/commessa/giorno inserito più volte. La pulizia automatica mantiene una sola registrazione valida.`
        : "";
      const title = hasPendingApproval
        ? canManage
          ? `Conferma le ore di ${operatorName} del giorno ${day}. Totale mostrato: ${formatHoursValue(dayTotal)}h.`
          : `${operatorName} - giorno ${day}: ${formatHoursValue(dayTotal)}h da confermare.`
        : canManage
          ? hasDuplicates
            ? `${mergedDetails} Totale mostrato: ${formatHoursValue(dayTotal)}h.`
            : `Modifica o elimina la registrazione ore di ${operatorName} del giorno ${day}. Ore salvate correttamente: ${formatHoursValue(dayTotal)}h.`
          : `${operatorName} - giorno ${day}: ${formatHoursValue(dayTotal)}h inserite.`;
      return `<td><button type="button" class="hours-value-btn ${statusClass}" data-hours-key="${escapeHTML(key)}" data-hours-pending="${hasPendingApproval && !hasDataError && !hasDuplicates ? "1" : "0"}" ${canManage ? "" : "disabled"} title="${escapeHTML(title)}">${escapeHTML(valueLabel)}</button></td>`;
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

  if (!operators.length) {
    hoursTableContext = {
      mode: "monthly",
      monthValue: String(options.monthValue || ""),
      monthLabel: `${String(monthMeta.month).padStart(2, "0")}/${monthMeta.year}`,
      commessaId,
      commessaName,
      monthMeta,
      operators: []
    };
    ui.hoursTableContainer.innerHTML = `<p class="muted hours-empty-message">Nessuna ora registrata per questa commessa nel mese selezionato.</p>`;
    ui.hoursTableFeedback.textContent = "Nessuna ora registrata per questa commessa nel mese selezionato.";
    if (ui.hoursTableExportBtn) ui.hoursTableExportBtn.disabled = true;
    setHoursConfirmVisibleButtonState(false);
    return hoursTableContext;
  }

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
    monthValue: String(options.monthValue || ""),
    monthLabel,
    commessaId,
    commessaName,
    monthMeta,
    operators: operators.map((name) => ({
      name,
      dayValues: (operatorsMap.get(name) || []).map((items) => items.reduce((sum, item) => sum + Number(typeof item === "object" ? item.ore : item || 0), 0))
    })),
    pendingVisibleKeys: []
  };

  const pendingVisibleKeys = Array.from(ui.hoursTableContainer.querySelectorAll(".hours-value-btn[data-hours-pending='1']"))
    .map((btn) => String(btn.dataset.hoursKey || ""))
    .filter(Boolean);
  if (hoursTableContext) hoursTableContext.pendingVisibleKeys = pendingVisibleKeys;
  setHoursConfirmVisibleButtonState(canManageData() && pendingVisibleKeys.length > 0, false);

  if (ui.hoursTableExportBtn) ui.hoursTableExportBtn.disabled = false;
  if (!operators.length) {
    ui.hoursTableFeedback.textContent = "Nessuna ora trovata: mostro una tabella vuota (minimo 10 righe).";
  } else {
    const hasPendingApprovals = (Array.isArray(reports) ? reports : [])
      .some((report) => String(report.sourceCollection || "oreReports") === "oreApprovalRequests");
    ui.hoursTableFeedback.textContent = hasPendingApprovals
      ? "Mostro anche le ore da confermare: sono evidenziate in giallo finché l'admin non le approva."
      : canManageData()
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
  return hoursTableContext;
}


function getPendingHoursSourcesForKeys(keys = []) {
  const selected = [];
  const seen = new Set();
  (Array.isArray(keys) ? keys : []).forEach((key) => {
    const sources = hoursTableRowsMap.get(String(key || "")) || [];
    sources.forEach((source) => {
      if (String(source.sourceCollection || "oreReports") !== "oreApprovalRequests") return;
      const sourceKey = `${source.reportId}__${source.entryIndex}__${source.rowIndex}`;
      if (seen.has(sourceKey)) return;
      seen.add(sourceKey);
      selected.push(source);
    });
  });
  return selected;
}

function getHoursSourceDayLabel(source) {
  const dateValue = String(source?.reportDate || "").trim();
  if (!dateValue) return "selezionato";
  const [year, month, day] = dateValue.split("-");
  if (year && month && day) return `${day}/${month}/${year}`;
  return dateValue;
}

async function approvePendingHoursSourcesFromTable(sources = []) {
  if (!canManageData()) throw new Error("Solo admin può confermare le ore.");
  const pendingSources = (Array.isArray(sources) ? sources : [])
    .filter((source) => String(source.sourceCollection || "oreReports") === "oreApprovalRequests" && source.reportId);
  if (!pendingSources.length) return [];
  const groupedByRequest = new Map();
  pendingSources.forEach((source) => {
    const requestId = String(source.reportId || "").trim();
    if (!requestId) return;
    if (!groupedByRequest.has(requestId)) groupedByRequest.set(requestId, []);
    groupedByRequest.get(requestId).push(source);
  });
  const results = [];
  for (const [requestId, requestSources] of groupedByRequest.entries()) {
    try {
      const request = await getHoursApprovalRequestById(requestId);
      if (!request) throw new Error("Richiesta ore non trovata.");
      const result = await saveApprovedHoursRequest(request, { sources: requestSources, fallbackDate: requestSources[0]?.reportDate || "" });
      results.push({ ok: true, requestId, reportId: result.reportId, sources: requestSources });
    } catch (error) {
      console.error("Errore conferma ore:", error);
      results.push({ ok: false, requestId, error, sources: requestSources });
    }
  }
  return results;
}

function markConfirmedHoursCells(keys = []) {
  const keySet = new Set((Array.isArray(keys) ? keys : []).map((key) => String(key || "")).filter(Boolean));
  if (!keySet.size || !ui.hoursTableContainer) return;
  keySet.forEach((key) => {
    const sources = hoursTableRowsMap.get(key) || [];
    sources.forEach((source) => {
      if (String(source.sourceCollection || "oreReports") !== "oreApprovalRequests") return;
      source.sourceCollection = "oreReports";
      source.approvalStatus = "approved";
    });
  });
  ui.hoursTableContainer.querySelectorAll(".hours-value-btn[data-hours-key]").forEach((btn) => {
    const key = String(btn.dataset.hoursKey || "");
    if (!keySet.has(key)) return;
    const valueText = String(btn.textContent || "").replace(/^⚠️\s*/, "✅ ").replace(" · da confermare", " · ore inserite");
    btn.textContent = valueText;
    btn.classList.remove("hours-value-warning", "hours-value-pending-approval");
    btn.classList.add("hours-value-ok");
    btn.dataset.hoursPending = "0";
    btn.title = btn.title.replace("Conferma le ore", "Ore confermate").replace(" da confermare", " confermate");
  });
  const pendingVisibleKeys = Array.from(ui.hoursTableContainer.querySelectorAll(".hours-value-btn[data-hours-pending='1']"))
    .map((btn) => String(btn.dataset.hoursKey || ""))
    .filter(Boolean);
  if (hoursTableContext) hoursTableContext.pendingVisibleKeys = pendingVisibleKeys;
  setHoursConfirmVisibleButtonState(canManageData() && pendingVisibleKeys.length > 0, false);
}

async function confirmPendingHoursFromTable(sources, options = {}) {
  const pendingSources = (Array.isArray(sources) ? sources : [])
    .filter((source) => String(source.sourceCollection || "oreReports") === "oreApprovalRequests");
  if (!pendingSources.length) return;
  const confirmed = await openHoursConfirmModal({
    title: "Confermare ore?",
    text: options.text || "Vuoi confermare le ore?",
    confirmLabel: options.confirmLabel || "Conferma ore"
  });
  if (!confirmed) return;
  if (ui.hoursTableFeedback) ui.hoursTableFeedback.textContent = "Conferma ore in corso...";
  setHoursConfirmVisibleButtonState(canManageData() && Boolean(hoursTableContext?.pendingVisibleKeys?.length), true);
  const results = await approvePendingHoursSourcesFromTable(pendingSources);
  const successfulResults = results.filter((result) => result.ok);
  const failedResults = results.filter((result) => !result.ok);
  const successfulKeys = Array.from(new Set(successfulResults.flatMap((result) =>
    (Array.isArray(result.sources) ? result.sources : []).map((source) => source.cellKey || `${source.operatore}__${Number(String(source.reportDate || "").split("-")[2] || 0)}`)
  ).filter(Boolean)));
  if (successfulKeys.length) markConfirmedHoursCells(successfulKeys);
  if (successfulResults.length) await loadSavedHoursReports();
  if (failedResults.length) {
    const firstError = failedResults[0]?.error;
    if (options.allVisible) {
      if (ui.hoursTableFeedback) ui.hoursTableFeedback.textContent = successfulResults.length
        ? "Alcune ore non sono state confermate."
        : (firstError?.message || "Alcune ore non sono state confermate.");
    } else if (ui.hoursTableFeedback) {
      ui.hoursTableFeedback.textContent = firstError?.message || "Errore: ore non confermate. Riprova.";
    }
    setHoursConfirmVisibleButtonState(canManageData() && Boolean(hoursTableContext?.pendingVisibleKeys?.length), false);
    return;
  }
  if (ui.hoursTableFeedback) ui.hoursTableFeedback.textContent = "Ore confermate correttamente.";
  setHoursConfirmVisibleButtonState(canManageData() && Boolean(hoursTableContext?.pendingVisibleKeys?.length), false);
}

async function handleConfirmVisiblePendingHours() {
  if (!canManageData()) return;
  const visibleKeys = Array.from(ui.hoursTableContainer?.querySelectorAll(".hours-value-btn[data-hours-pending='1']") || [])
    .map((btn) => String(btn.dataset.hoursKey || ""))
    .filter(Boolean);
  const pendingSources = getPendingHoursSourcesForKeys(visibleKeys);
  if (!pendingSources.length) {
    setHoursConfirmVisibleButtonState(false);
    return;
  }
  await confirmPendingHoursFromTable(pendingSources, {
    text: "Vuoi confermare tutte le ore visibili in questa tabella?",
    confirmLabel: "Conferma ore",
    allVisible: true
  });
}

async function handleHoursValueAction(cellKey) {
  if (!canManageData()) return;
  let sources = hoursTableRowsMap.get(String(cellKey || ""));
  if (!sources || !sources.length) return;
  const pendingSources = sources.filter((source) => String(source.sourceCollection || "oreReports") === "oreApprovalRequests");
  if (pendingSources.length && pendingSources.length === sources.length) {
    const firstSource = pendingSources[0] || {};
    const operatorLabel = String(firstSource.operatore || "OPERATORE").trim() || "OPERATORE";
    const dayLabel = getHoursSourceDayLabel(firstSource);
    await confirmPendingHoursFromTable(pendingSources, {
      text: `Vuoi confermare le ore di ${operatorLabel} per il giorno ${dayLabel}?`,
      confirmLabel: "Conferma ore"
    });
    return;
  }
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
      const collectionName = String(source.sourceCollection || "oreReports") === "oreApprovalRequests"
        ? "oreApprovalRequests"
        : "oreReports";
      const groupKey = `${collectionName}::${source.reportId}`;
      if (!groupedByReport.has(groupKey)) groupedByReport.set(groupKey, { collectionName, reportId: source.reportId, sources: [] });
      groupedByReport.get(groupKey).sources.push(source);
    });
    for (const reportGroup of groupedByReport.values()) {
      const { collectionName, reportId, sources: reportSources } = reportGroup;
      const docRef = db.collection(collectionName).doc(reportId);
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
        await docRef.update({
          entries: nextEntries,
          updatedAt: firebase.firestore.FieldValue.serverTimestamp()
        });
      } else {
        await docRef.delete();
      }
      if (normalizedAction !== "M") {
        const deletedLockEntries = reportSources.map((source) => ({
          commessaId: source.entryCommessaId || "",
          commessaName: commesseById.get(source.entryCommessaId)?.nome || "Commessa",
          rows: [{ operatore: source.operatore || "", ore: source.ore || 1 }]
        }));
        await updateHoursLocksForEntries(reportSources[0]?.reportDate || data.date || "", deletedLockEntries, {
          status: "deleted",
          reportId,
          sourceCollection: collectionName
        });
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
  try {
    if (loadingOre && hoursTableLoadPromise) await hoursTableLoadPromise;
    const mode = String(hoursTableContext?.mode || "monthly");
    if (mode === "tot_operator") {
      const monthValue = String(hoursTableContext?.monthValue || ui.hoursStatsMonth?.value || "").trim();
      const rows = Array.isArray(hoursTableContext?.rows) ? hoursTableContext.rows : [];
      logHoursDebug("dati usati per export", { mode, monthValue, rows });
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
      logHoursDebug("dati usati per export", { mode, monthValue, rows });
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
    const commessaInfo = getSelectedHoursCommessaInfo(commessaId);
    logHoursDebug("mese selezionato", monthValue);
    logHoursDebug("anno", monthMeta?.year || "non valido");
    logHoursDebug("mese numerico", monthMeta?.month || "non valido");
    logHoursDebug("commessa selezionata", commessaInfo.codice || commessaInfo.nome || commessaInfo.id || "nessuna");
    if (!monthMeta || !commessaId) {
      alert("Seleziona mese e commessa prima di esportare Excel.");
      return;
    }
    const contextMatchesSelection = hoursTableContext?.mode === "monthly"
      && hoursTableContext.monthValue === monthValue
      && String(hoursTableContext.commessaId || "") === commessaId;
    if (!contextMatchesSelection) {
      const loadedContext = await loadHoursMonthlyTable();
      if (!loadedContext?.operators?.length) {
        alert("Nessuna ora registrata per questa commessa nel mese selezionato.");
        return;
      }
    }
    if (hoursTableContext?.mode === "monthly" && !hoursTableContext?.operators?.length) {
      alert("Nessuna ora registrata per questa commessa nel mese selezionato.");
      return;
    }
    const commessaName = commesseById.get(commessaId)?.nome || "Commessa";
    const monthLabel = `${String(monthMeta.month).padStart(2, "0")}/${monthMeta.year}`;
    const reports = await fetchHoursReportsForMonth(monthValue, monthMeta, { includePendingApprovals: true });
    const { operatorDayMap, operatorTotals, operatorCommessaTotals } = buildHoursMonthlyExportData(reports, commessaId, monthMeta);
    logHoursDebug("record trovati", Array.isArray(reports) ? reports.length : 0);
    logHoursDebug("dati usati per export", { monthValue, commessa: commessaInfo, operatoriCommessa: Array.from(operatorDayMap.entries()) });
    if (!operatorDayMap.size) {
      alert("Nessuna ora registrata per questa commessa nel mese selezionato.");
      return;
    }

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
  } catch (error) {
    console.error("Errore export Excel ore:", error);
    if (ui.hoursTableFeedback) ui.hoursTableFeedback.textContent = "Errore export Excel ore. Controlla i dati o riprova.";
    alert("Errore export Excel ore. Controlla i dati o riprova.");
  }
}

async function exportHoursGlobalMonthlyTable() {
  const monthValue = String(ui.hoursTableMonth?.value || ui.hoursStatsMonth?.value || "").trim();
  const monthMeta = getMonthMeta(monthValue);
  if (!monthMeta) {
    alert("Seleziona un mese valido prima di esportare il file globale.");
    return;
  }
  if (!window.ExcelJS?.Workbook) {
    alert("Libreria Excel non disponibile. Ricarica la pagina e riprova.");
    return;
  }

  try {
    logHoursDebug("mese selezionato", monthValue);
    logHoursDebug("anno", monthMeta.year);
    logHoursDebug("mese numerico", monthMeta.month);
    const reports = await fetchHoursReportsForMonth(monthValue, monthMeta, { includePendingApprovals: true });
    logHoursDebug("record trovati", Array.isArray(reports) ? reports.length : 0);
    const commessaMap = new Map();
    let totalValidGlobalRows = 0;
  reports.forEach((report) => {
    const day = Number(String(report.date || "").split("-")[2] || 0);
    if (!day || day < 1 || day > monthMeta.daysInMonth) return;
    const entries = Array.isArray(report.entries) ? report.entries : [];
    entries.forEach((entry) => {
      const entryCommessaInfo = resolveHoursEntryCommessa(entry);
      const commessaId = String(entryCommessaInfo.id || entryCommessaInfo.key || "").trim();
      if (!commessaId) return;
      const commessaName = String(entryCommessaInfo.nome || entry.commessaName || commesseById.get(entryCommessaInfo.id)?.nome || "Commessa").trim() || "Commessa";
      const commessaCode = String(entryCommessaInfo.codice || commesseById.get(entryCommessaInfo.id)?.codice || "").trim();
      if (!commessaMap.has(commessaId)) {
        commessaMap.set(commessaId, { commessaName, commessaCode, operatorsMap: new Map() });
      }
      const commessaBucket = commessaMap.get(commessaId);
      (Array.isArray(entry.rows) ? entry.rows : []).forEach((row) => {
        const operatore = String(row.operatore || "").trim();
        const ore = Number(row.ore || 0);
        if (!operatore || ore <= 0) return;
        totalValidGlobalRows += 1;
        const operatorNorm = operatore.toLocaleLowerCase("it-IT").replace(/\s+/g, " ").trim();
        if (!commessaBucket.operatorsMap.has(operatorNorm)) {
          commessaBucket.operatorsMap.set(operatorNorm, {
            displayName: operatore,
            days: Array.from({ length: monthMeta.daysInMonth }, () => 0)
          });
        }
        commessaBucket.operatorsMap.get(operatorNorm).days[day - 1] += ore;
      });
    });
  });

  logHoursDebug("dati usati per export", { mode: "global", monthValue, commesse: Array.from(commessaMap.values()).map((item) => ({
    commessaName: item.commessaName,
    commessaCode: item.commessaCode,
    operatori: Array.from(item.operatorsMap.values())
  })) });
  if (!commessaMap.size || totalValidGlobalRows <= 0) {
    alert("Nessuna ora registrata nel mese selezionato per l'export globale.");
    if (ui.hoursTableFeedback) ui.hoursTableFeedback.textContent = "Nessuna ora registrata nel mese selezionato per l'export globale.";
    return;
  }

  const monthNameIt = [
    "GENNAIO", "FEBBRAIO", "MARZO", "APRILE", "MAGGIO", "GIUGNO",
    "LUGLIO", "AGOSTO", "SETTEMBRE", "OTTOBRE", "NOVEMBRE", "DICEMBRE"
  ][monthMeta.month - 1] || monthValue;
  const monthLabelIt = `${monthNameIt} ${monthMeta.year}`;

  const workbook = new ExcelJS.Workbook();
  workbook.creator = "Hera App";
  workbook.created = new Date();
  const worksheet = workbook.addWorksheet("Export globale", {
    views: [{ state: "frozen", xSplit: 1, ySplit: 8 }]
  });

  const dayStartColumn = 2;
  const totalColumn = monthMeta.daysInMonth + 2;
  const workedDaysColumn = monthMeta.daysInMonth + 3;
  const avgHoursColumn = monthMeta.daysInMonth + 4;
  const lastColumn = avgHoursColumn;
  const dayHeaderFill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFFFFF00" } };
  const totalColumnFill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFD9EAF7" } };
  const hoursFilledCell = { type: "pattern", pattern: "solid", fgColor: { argb: "FFEDE9DD" } };
  const weekendFill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFE5E5E5" } };
  const errorFill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFFFC7CE" } };
  const whiteFill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFFFFFFF" } };
  const thinBorder = {
    top: { style: "thin", color: { argb: "FF000000" } },
    left: { style: "thin", color: { argb: "FF000000" } },
    bottom: { style: "thin", color: { argb: "FF000000" } },
    right: { style: "thin", color: { argb: "FF000000" } }
  };
  const mediumSide = { style: "medium", color: { argb: "FF000000" } };
  const thickSide = { style: "thick", color: { argb: "FF000000" } };
  const isWeekendDay = (dayNumber) => {
    const dayDate = new Date(monthMeta.year, monthMeta.month - 1, dayNumber);
    const weekday = dayDate.getDay();
    return weekday === 0 || weekday === 6;
  };
  const getExcelNumberFormat = (value) => {
    const num = Number(value || 0);
    if (!Number.isFinite(num) || num <= 0) return null;
    return Number.isInteger(num) ? "0" : "0.##";
  };
  const setThinBorder = (cell) => {
    cell.border = thinBorder;
  };
  const setOuterBlockBorder = (startRow, endRow) => {
    for (let row = startRow; row <= endRow; row += 1) {
      for (let col = 1; col <= lastColumn; col += 1) {
        const cell = worksheet.getCell(row, col);
        const border = { ...(cell.border || {}) };
        if (row === startRow) border.top = mediumSide;
        if (row === endRow) border.bottom = mediumSide;
        if (col === 1) border.left = mediumSide;
        if (col === lastColumn) border.right = mediumSide;
        cell.border = border;
      }
    }
  };
  const addWeekSeparatorBorders = (rowIndex) => {
    for (let day = 1; day <= monthMeta.daysInMonth; day += 1) {
      const date = new Date(monthMeta.year, monthMeta.month - 1, day);
      const isSunday = date.getDay() === 0;
      if (!isSunday || day === monthMeta.daysInMonth) continue;
      const dayCol = dayStartColumn + day - 1;
      const cell = worksheet.getCell(rowIndex, dayCol);
      const border = { ...(cell.border || {}) };
      border.right = thickSide;
      cell.border = border;
      const nextCol = dayCol + 1;
      if (nextCol <= dayStartColumn + monthMeta.daysInMonth - 1) {
        const nextCell = worksheet.getCell(rowIndex, nextCol);
        const nextBorder = { ...(nextCell.border || {}) };
        nextBorder.left = thickSide;
        nextCell.border = nextBorder;
      }
    }
  };

  let rowPointer = 1;
  const commesseSorted = Array.from(commessaMap.values())
    .sort((a, b) => a.commessaName.localeCompare(b.commessaName, "it"));

  const totalCommesse = commesseSorted.length;
  const totalOperatorsMonth = commesseSorted.reduce((acc, commessaBlock) => (
    acc + commessaBlock.operatorsMap.size
  ), 0);
  const totalHoursMonth = commesseSorted.reduce((acc, commessaBlock) => (
    acc + Array.from(commessaBlock.operatorsMap.values()).reduce((hoursAcc, operator) => (
      hoursAcc + operator.days.reduce((dayAcc, value) => dayAcc + Number(value || 0), 0)
    ), 0)
  ), 0);

  const summaryStartRow = rowPointer;
  worksheet.mergeCells(summaryStartRow, 1, summaryStartRow, lastColumn);
  const summaryTitleCell = worksheet.getCell(summaryStartRow, 1);
  summaryTitleCell.value = "RIEPILOGO GESTIONE ORE GLOBAL";
  summaryTitleCell.font = { bold: true, size: 14, color: { argb: "FF000000" } };
  summaryTitleCell.alignment = { horizontal: "center", vertical: "middle" };
  summaryTitleCell.fill = whiteFill;
  rowPointer += 1;

  const summaryRows = [
    ["MESE DI RIFERIMENTO", monthNameIt],
    ["ANNO", String(monthMeta.year)],
    ["DATA ESPORTAZIONE", new Date().toLocaleDateString("it-IT")],
    ["TOTALE ORE MESE", totalHoursMonth > 0 ? totalHoursMonth : ""],
    ["TOTALE OPERATORI", totalOperatorsMonth],
    ["TOTALE COMMESSE", totalCommesse]
  ];
  summaryRows.forEach(([label, value]) => {
    const row = worksheet.getRow(rowPointer);
    row.getCell(1).value = label;
    row.getCell(2).value = value;
    if (typeof value === "number" && value > 0) {
      row.getCell(2).numFmt = getExcelNumberFormat(value) || "0";
    }
    worksheet.mergeCells(rowPointer, 2, rowPointer, lastColumn);
    rowPointer += 1;
  });
  rowPointer += 1;

  commesseSorted.forEach((commessaBlock, idx) => {
    const operatorRows = Array.from(commessaBlock.operatorsMap.values())
      .sort((a, b) => a.displayName.localeCompare(b.displayName, "it"));
    const operators = operatorRows.length ? operatorRows : [];

    const startRow = rowPointer;
    const commessaRow = worksheet.getRow(rowPointer);
    commessaRow.getCell(1).value = "COMMESSA";
    commessaRow.getCell(2).value = commessaBlock.commessaName;
    worksheet.mergeCells(rowPointer, 2, rowPointer, lastColumn);
    commessaRow.font = { bold: true, size: 14, color: { argb: "FF0B1F44" } };

    rowPointer += 1;
    const codeRow = worksheet.getRow(rowPointer);
    codeRow.getCell(1).value = "CODICE COMMESSA";
    codeRow.getCell(2).value = commessaBlock.commessaCode || "N/D";
    worksheet.mergeCells(rowPointer, 2, rowPointer, lastColumn);
    codeRow.font = { bold: true, size: 12, color: { argb: "FF0B1F44" } };

    rowPointer += 1;
    const meseRow = worksheet.getRow(rowPointer);
    meseRow.getCell(1).value = "MESE RIF.";
    meseRow.getCell(2).value = monthLabelIt;
    worksheet.mergeCells(rowPointer, 2, rowPointer, lastColumn);

    rowPointer += 1;
    const headerRow = worksheet.getRow(rowPointer);
    headerRow.getCell(1).value = "OPERATORE";
    headerRow.getCell(1).fill = dayHeaderFill;
    headerRow.getCell(1).font = { bold: true, color: { argb: "FF000000" } };
    for (let day = 1; day <= monthMeta.daysInMonth; day += 1) {
      headerRow.getCell(day + 1).value = day;
      headerRow.getCell(day + 1).fill = dayHeaderFill;
      headerRow.getCell(day + 1).font = { bold: true, color: { argb: "FF000000" } };
    }
    headerRow.getCell(totalColumn).value = "TOTALE";
    headerRow.getCell(totalColumn).fill = totalColumnFill;
    headerRow.getCell(totalColumn).font = { bold: true, color: { argb: "FF000000" } };
    headerRow.getCell(workedDaysColumn).value = "GIORNI LAVORATI";
    headerRow.getCell(workedDaysColumn).fill = dayHeaderFill;
    headerRow.getCell(workedDaysColumn).font = { bold: true, color: { argb: "FF000000" } };
    headerRow.getCell(avgHoursColumn).value = "MEDIA ORE/GIORNO";
    headerRow.getCell(avgHoursColumn).fill = dayHeaderFill;
    headerRow.getCell(avgHoursColumn).font = { bold: true, color: { argb: "FF000000" } };
    headerRow.height = 24;

    rowPointer += 1;
    let commessaTotal = 0;

    operators.forEach((operatorData, operatorIdx) => {
      const row = worksheet.getRow(rowPointer + operatorIdx);
      row.getCell(1).value = operatorData.displayName || "";
      let total = 0;
      let workedDays = 0;
      for (let dayIdx = 0; dayIdx < monthMeta.daysInMonth; dayIdx += 1) {
        const value = Number(operatorData.days[dayIdx] || 0);
        const cell = row.getCell(dayIdx + 2);
        if (isWeekendDay(dayIdx + 1)) {
          cell.fill = weekendFill;
        } else {
          cell.fill = whiteFill;
        }
        if (value > 0) {
          cell.value = value;
          cell.fill = value > 12 ? errorFill : hoursFilledCell;
          const numFmt = getExcelNumberFormat(value);
          if (numFmt) cell.numFmt = numFmt;
          total += value;
          workedDays += 1;
        } else {
          cell.value = null;
        }
      }
      row.getCell(totalColumn).value = total > 0 ? total : null;
      row.getCell(totalColumn).fill = totalColumnFill;
      if (total > 0) row.getCell(totalColumn).numFmt = getExcelNumberFormat(total);
      row.getCell(workedDaysColumn).value = workedDays > 0 ? workedDays : null;
      row.getCell(workedDaysColumn).numFmt = "0";
      row.getCell(avgHoursColumn).value = workedDays > 0 ? total / workedDays : null;
      if (workedDays > 0) row.getCell(avgHoursColumn).numFmt = "0.##";
      commessaTotal += total;
      row.height = 21;
    });

    const totalCommessaRowIndex = rowPointer + operators.length;
    const totalCommessaRow = worksheet.getRow(totalCommessaRowIndex);
    totalCommessaRow.getCell(1).value = "TOTALE COMMESSA";
    worksheet.mergeCells(totalCommessaRowIndex, 1, totalCommessaRowIndex, totalColumn - 1);
    totalCommessaRow.getCell(totalColumn).value = commessaTotal > 0 ? commessaTotal : null;
    totalCommessaRow.getCell(totalColumn).fill = totalColumnFill;
    if (commessaTotal > 0) totalCommessaRow.getCell(totalColumn).numFmt = getExcelNumberFormat(commessaTotal);
    totalCommessaRow.getCell(1).font = { bold: true, color: { argb: "FF000000" } };
    totalCommessaRow.getCell(totalColumn).font = { bold: true, color: { argb: "FF000000" } };
    totalCommessaRow.height = 22;

    const endRow = totalCommessaRowIndex;
    for (let row = startRow; row <= endRow; row += 1) {
      for (let col = 1; col <= lastColumn; col += 1) {
        const cell = worksheet.getCell(row, col);
        if (!cell.fill) cell.fill = whiteFill;
        setThinBorder(cell);
        if (row >= startRow + 2 && row < endRow) {
          const isOperatorName = col === 1;
          cell.alignment = {
            vertical: "middle",
            horizontal: isOperatorName ? "left" : "center"
          };
        } else if (row === endRow) {
          cell.alignment = {
            vertical: "middle",
            horizontal: col === 1 ? "left" : "center"
          };
        } else {
          cell.alignment = {
            vertical: "middle",
            horizontal: col === 1 ? "left" : "center"
          };
        }
        const isTitleLabel = col === 1 && row <= startRow + 2;
        const isHeaderRow = row === startRow + 2;
        if (isTitleLabel || isHeaderRow) {
          cell.font = { ...(cell.font || {}), bold: true, color: { argb: "FF000000" } };
        }
      }
      if (row <= startRow + 1) worksheet.getRow(row).height = 22;
      addWeekSeparatorBorders(row);
    }

    setOuterBlockBorder(startRow, endRow);

    rowPointer = endRow + 2;
    if (idx < commesseSorted.length - 1) {
      worksheet.getRow(rowPointer - 1).height = 10;
    }
  });

  for (let row = summaryStartRow; row <= summaryStartRow + summaryRows.length; row += 1) {
    for (let col = 1; col <= lastColumn; col += 1) {
      const cell = worksheet.getCell(row, col);
      setThinBorder(cell);
      if (!cell.fill) cell.fill = whiteFill;
      cell.alignment = {
        vertical: "middle",
        horizontal: col === 1 ? "left" : "center"
      };
      if (col === 1 || row === summaryStartRow) {
        cell.font = { ...(cell.font || {}), bold: true, color: { argb: "FF000000" } };
      }
    }
    worksheet.getRow(row).height = row === summaryStartRow ? 28 : 21;
  }
  setOuterBlockBorder(summaryStartRow, summaryStartRow + summaryRows.length);

  for (let col = 1; col <= lastColumn; col += 1) {
    if (col === 1) {
      worksheet.getColumn(col).width = 28;
      continue;
    }
    if (col >= dayStartColumn && col <= totalColumn - 1) {
      worksheet.getColumn(col).width = 4.2;
      continue;
    }
    if (col === totalColumn) {
      worksheet.getColumn(col).width = 11;
      continue;
    }
    if (col === workedDaysColumn) {
      worksheet.getColumn(col).width = 16;
      continue;
    }
    worksheet.getColumn(col).width = 18;
  }

  worksheet.autoFilter = {
    from: { row: summaryStartRow + 7, column: 1 },
    to: { row: summaryStartRow + 7, column: 1 }
  };

  for (let row = 1; row <= worksheet.rowCount; row += 1) {
    const currentRow = worksheet.getRow(row);
    if (!currentRow.height) currentRow.height = 21;
  }

  for (let row = summaryStartRow; row <= summaryStartRow + 6; row += 1) {
    for (let col = 1; col <= lastColumn; col += 1) {
      const cell = worksheet.getCell(row, col);
      setThinBorder(cell);
      if (!cell.fill) cell.fill = whiteFill;
      cell.alignment = {
        vertical: "middle",
        horizontal: col === 1 ? "left" : "center"
      };
      if (col === 1 || row === summaryStartRow) {
        cell.font = { ...(cell.font || {}), bold: true, color: { argb: "FF000000" } };
      }
    }
    worksheet.getRow(row).height = row === summaryStartRow ? 28 : 21;
  }
  setOuterBlockBorder(summaryStartRow, summaryStartRow + 6);

  for (let col = 1; col <= lastColumn; col += 1) {
    if (col === 1) {
      worksheet.getColumn(col).width = 28;
      continue;
    }
    if (col >= dayStartColumn && col <= totalColumn - 1) {
      worksheet.getColumn(col).width = 4.2;
      continue;
    }
    if (col === totalColumn) {
      worksheet.getColumn(col).width = 11;
      continue;
    }
    if (col === workedDaysColumn) {
      worksheet.getColumn(col).width = 16;
      continue;
    }
    worksheet.getColumn(col).width = 18;
  }

  worksheet.autoFilter = {
    from: { row: summaryStartRow + 7, column: 1 },
    to: { row: summaryStartRow + 7, column: 1 }
  };

  for (let row = 1; row <= worksheet.rowCount; row += 1) {
    const currentRow = worksheet.getRow(row);
    if (!currentRow.height) currentRow.height = 21;
  }

  const safeMonth = monthValue.replace("/", "-");
  const buffer = await workbook.xlsx.writeBuffer();
  const blob = new Blob([buffer], {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
  });
  const fileName = `ore_global_${safeMonth}.xlsx`;
  if (window.navigator?.msSaveOrOpenBlob) {
    window.navigator.msSaveOrOpenBlob(blob, fileName);
    return;
  }
  const url = window.URL.createObjectURL(blob);
  const anchor = document.createElement("a");
  anchor.href = url;
  anchor.download = fileName;
  document.body.appendChild(anchor);
  anchor.click();
  anchor.remove();
  window.URL.revokeObjectURL(url);
  } catch (error) {
    console.error("Errore export Excel Global ore:", error);
    if (ui.hoursTableFeedback) ui.hoursTableFeedback.textContent = "Errore export Excel Global. Controlla i dati o riprova.";
    alert("Errore export Excel Global. Controlla i dati o riprova.");
  }
}

function renderHoursOperatoriOptions() {
  if (!ui.hoursOperatoriOptions) return;
  ui.hoursOperatoriOptions.innerHTML = "";
  personaleRecords.forEach((person) => {
    const option = document.createElement("option");
    option.value = getPersonaleDisplayName(person);
    ui.hoursOperatoriOptions.appendChild(option);
  });
  Array.from(ui.hoursCommesseList?.querySelectorAll(".hours-commessa-card") || []).forEach((card) => {
    renderHoursOperatorSuggestions(card);
  });
}

function renderHoursCommessaSelectOptions() {
  const selects = ui.hoursCommesseList ? ui.hoursCommesseList.querySelectorAll(".hours-commessa-select") : [];
  const commesse = sortCommesseByCreatedAtDesc(Array.from(commesseById.values()));
  renderHoursTableCommessaOptions(commesse);
  if (!selects.length) return;
  selects.forEach((select) => {
    const selectedValue = select.value;
    select.innerHTML = "<option value=''>Seleziona commessa</option>";
    commesse.forEach((commessa) => {
      select.appendChild(createCommessaOption(commessa, { includeHierarchy: true }));
    });
    if (selectedValue && commesse.some((commessa) => commessa.id === selectedValue)) {
      select.value = selectedValue;
    }
    renderHoursCardCommessaButtons(select.closest(".hours-commessa-card"), commesse);
  });
}

function renderHoursTableCommessaOptions(commesseInput = null) {
  if (!ui.hoursTableCommessaSelect) return;
  const commesse = Array.isArray(commesseInput) ? commesseInput : sortCommesseByCreatedAtDesc(Array.from(commesseById.values()));
  const selectedValue = ui.hoursTableCommessaSelect.value;
  ui.hoursTableCommessaSelect.innerHTML = "<option value=''>Seleziona commessa</option>";
  commesse.forEach((commessa) => {
    ui.hoursTableCommessaSelect.appendChild(createCommessaOption(commessa, { includeHierarchy: true }));
  });
  if (selectedValue && commesse.some((commessa) => commessa.id === selectedValue)) {
    ui.hoursTableCommessaSelect.value = selectedValue;
  }
  renderHoursTableCommessaButtons(commesse);
}

function normalizeHoursCommessaSearch(value) {
  return String(value || "")
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .toLocaleLowerCase("it-IT")
    .replace(/[^a-z0-9]+/g, " ")
    .replace(/\s+/g, " ")
    .trim();
}

function focusHoursCommessaSearch(container) {
  const search = container?.querySelector(".hours-commessa-picker-search");
  if (!(search instanceof HTMLInputElement)) return;
  search.focus({ preventScroll: true });
  search.setSelectionRange(search.value.length, search.value.length);
}

function closeOtherHoursCommessaPickers(activeContainer = null) {
  document.querySelectorAll(".hours-commessa-picker.is-open").forEach((picker) => {
    if (activeContainer && picker === activeContainer) return;
    picker.classList.remove("is-open");
    picker.querySelector(".hours-commessa-picker-menu")?.classList.add("hidden");
    picker.querySelector(".hours-commessa-picker-toggle")?.setAttribute("aria-expanded", "false");
  });
}

function renderHoursCommessaPicker(container, select, commesseInput = null, options = {}) {
  if (!container || !select) return;
  const commesse = Array.isArray(commesseInput) ? commesseInput : Array.from(commesseById.values());
  const selectedValue = String(select.value || "").trim();
  const selectedIndex = commesse.findIndex((commessa) => commessa.id === selectedValue);
  const selectedCommessa = selectedIndex >= 0 ? commesse[selectedIndex] : null;
  const isOpen = options.keepOpen === true;
  const query = String(options.query || "").trim();
  const normalizedQuery = normalizeHoursCommessaSearch(query);
  const disabled = select.disabled || options.disabled === true;
  const selectedColor = selectedCommessa ? getCommessaAccentColor(selectedCommessa.id, selectedIndex) : "#64748b";
  const filteredCommesse = normalizedQuery
    ? commesse.filter((commessa) => normalizeHoursCommessaSearch(getCommessaDisplayName(commessa)).includes(normalizedQuery))
    : commesse;

  if (!commesse.length) {
    container.innerHTML = "<p class='muted hours-commessa-empty'>Nessuna commessa disponibile.</p>";
    return;
  }

  container.classList.toggle("is-open", isOpen);
  container.innerHTML = `
    <button type="button" class="hours-commessa-picker-toggle" style="--commessa-accent:${escapeHTML(selectedColor)}" aria-expanded="${isOpen ? "true" : "false"}" ${disabled ? "disabled" : ""}>
      <span class="hours-commessa-picker-label">${escapeHTML(selectedCommessa ? getCommessaDisplayName(selectedCommessa) : "Seleziona commessa")}</span>
      <span class="hours-commessa-picker-arrow" aria-hidden="true">▼</span>
    </button>
    <div class="hours-commessa-picker-menu ${isOpen ? "" : "hidden"}">
      <input class="hours-commessa-picker-search" type="search" placeholder="Cerca commessa…" value="${escapeHTML(query)}" aria-label="Cerca commessa">
      <div class="hours-commessa-picker-list" role="listbox" aria-label="Elenco commesse">
        ${filteredCommesse.length ? filteredCommesse.map((commessa, idx) => {
          const originalIndex = commesse.findIndex((item) => item.id === commessa.id);
          const color = getCommessaAccentColor(commessa.id, originalIndex >= 0 ? originalIndex : idx);
          const active = selectedValue === commessa.id;
          return `<button type="button" class="hours-commessa-picker-option ${active ? "is-active" : ""}" data-hours-commessa-option="${escapeHTML(commessa.id)}" style="--commessa-accent:${escapeHTML(color)}" role="option" aria-selected="${active ? "true" : "false"}">${escapeHTML(getCommessaDisplayName(commessa))}</button>`;
        }).join("") : "<p class='muted hours-commessa-empty'>Nessuna commessa trovata.</p>"}
      </div>
    </div>
  `;

  const toggle = container.querySelector(".hours-commessa-picker-toggle");
  toggle?.addEventListener("click", (event) => {
    event.preventDefault();
    event.stopPropagation();
    const nextOpenState = !isOpen;
    if (nextOpenState) closeOtherHoursCommessaPickers(container);
    renderHoursCommessaPicker(container, select, commesse, { ...options, keepOpen: nextOpenState, query });
    if (nextOpenState) focusHoursCommessaSearch(container);
  });

  const search = container.querySelector(".hours-commessa-picker-search");
  search?.addEventListener("click", (event) => {
    event.stopPropagation();
  });
  search?.addEventListener("input", () => {
    renderHoursCommessaPicker(container, select, commesse, { ...options, keepOpen: true, query: search.value });
    requestAnimationFrame(() => focusHoursCommessaSearch(container));
  });

  container.querySelectorAll("[data-hours-commessa-option]").forEach((btn) => {
    btn.addEventListener("click", (event) => {
      event.stopPropagation();
      const commessaId = String(btn.dataset.hoursCommessaOption || "").trim();
      if (!commessaId) return;
      const changed = select.value !== commessaId;
      select.value = commessaId;
      if (changed) {
        select.dispatchEvent(new Event("change", { bubbles: true }));
      }
      renderHoursCommessaPicker(container, select, commesse, { ...options, keepOpen: false, query: "" });
      if (changed && typeof options.onChange === "function") options.onChange(commessaId);
    });
  });
}

function renderHoursCardCommessaButtons(card, commesseInput = null) {
  if (!card) return;
  const pickerWrap = card.querySelector(".hours-commesse-buttons");
  const select = card.querySelector(".hours-commessa-select");
  renderHoursCommessaPicker(pickerWrap, select, commesseInput, {
    onChange: () => {
      unlockHoursFinalizeButton();
      applyHoursSuggestedOperators(card, { force: true });
    }
  });
}

function renderHoursTableCommessaButtons(commesseInput = null) {
  if (!ui.hoursTableCommesseButtons || !ui.hoursTableCommessaSelect) return;
  renderHoursCommessaPicker(ui.hoursTableCommesseButtons, ui.hoursTableCommessaSelect, commesseInput, {
    disabled: ui.hoursTableCommessaSelect.disabled,
    onChange: () => loadHoursMonthlyTable()
  });
}


document.addEventListener("click", (event) => {
  const target = event.target;
  if (!(target instanceof Element) || target.closest(".hours-commessa-picker")) return;
  closeOtherHoursCommessaPickers();
});

function normalizeSquadraMemberIdentity(value) {
  return String(value || "")
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .toLocaleLowerCase("it-IT")
    .replace(/[^a-z0-9]+/g, " ")
    .replace(/\s+/g, " ")
    .trim();
}

function getCurrentUserIdentityParts() {
  if (!currentUser) return [];
  const parts = [currentUser.displayName, currentUser.email];
  const emailLocal = String(currentUser.email || "").split("@")[0] || "";
  if (emailLocal) parts.push(emailLocal, emailLocal.replace(/[._-]+/g, " "));
  return [...new Set(parts.map(normalizeSquadraMemberIdentity).filter(Boolean))];
}

function doSquadraMemberAndUserMatch(memberName, identityParts = getCurrentUserIdentityParts()) {
  const member = normalizeSquadraMemberIdentity(memberName);
  if (!member || !identityParts.length) return false;
  return identityParts.some((part) => {
    if (!part) return false;
    if (member === part) return true;
    const memberTokens = member.split(" ").filter((token) => token.length > 1);
    const partTokens = part.split(" ").filter((token) => token.length > 1);
    return partTokens.length >= 2 && partTokens.every((token) => memberTokens.includes(token));
  });
}

function getSquadraDataForCommessaDate(commessaId, dateValue = "") {
  const id = String(commessaId || "").trim();
  if (!id) return null;
  const dateKey = String(dateValue || "").trim() || getActiveSquadreDateKey() || new Date().toISOString().slice(0, 10);
  const storicoDelGiorno = squadreHistoryByDate.get(dateKey) || new Map();
  return storicoDelGiorno.get(id) || squadreByCommessa.get(id) || null;
}

function getCurrentUserSquadraAssignment(commessaId, dateValue = "") {
  if (!currentUser) return null;
  const squadData = getSquadraDataForCommessaDate(commessaId, dateValue);
  const squadRows = Array.isArray(squadData?.squadre) ? squadData.squadre : getLegacySquadreRows(squadData || {});
  const identities = getCurrentUserIdentityParts();
  for (let index = 0; index < squadRows.length; index += 1) {
    const personale = parseMultiEntryValue(squadRows[index]?.personale || "");
    const matchedName = personale.find((name) => doSquadraMemberAndUserMatch(name, identities));
    if (matchedName) {
      return {
        squadraIndex: index + 1,
        squadraLabel: `Squadra ${index + 1}`,
        matchedName,
        squadData,
        row: squadRows[index]
      };
    }
  }
  return null;
}

function getCurrentUserAssignedCommesseForDate(dateKey = getActiveSquadreDateKey()) {
  if (!currentUser || !dateKey) return [];
  const storicoDelGiorno = squadreHistoryByDate.get(dateKey) || new Map();
  const identities = getCurrentUserIdentityParts();
  const matches = [];

  Array.from(commesseById.values()).forEach((commessa) => {
    const squadData = storicoDelGiorno.get(commessa.id) || null;
    if (!squadData) return;
    const squadRows = Array.isArray(squadData.squadre) ? squadData.squadre : getLegacySquadreRows(squadData || {});
    const matchedRows = [];

    squadRows.forEach((row, index) => {
      const rowMembers = [
        ...parseMultiEntryValue(row?.personale || ""),
        ...parseMultiEntryValue(row?.caposquadra || "")
      ];
      const matchedName = rowMembers.find((name) => doSquadraMemberAndUserMatch(name, identities));
      if (matchedName) {
        matchedRows.push({
          squadraIndex: index + 1,
          squadraLabel: `Squadra ${index + 1}`,
          matchedName,
          row
        });
      }
    });

    if (matchedRows.length) {
      matches.push({
        commessa,
        commessaId: commessa.id,
        commessaName: commessa.nome || "Commessa",
        squadData,
        matchedRows
      });
    }
  });

  return matches;
}

function tryAutoOpenAssignedCommessaAtStartup() {
  if (startupAssignedCommessaAutoOpenDone || !currentUser) return;
  if (parseCommessaHash().id) {
    startupAssignedCommessaAutoOpenDone = true;
    return;
  }
  if (!sharedSquadreViewConfigLoaded) return;
  if (commesseLoadState.status !== "loaded" && commesseLoadState.status !== "empty") return;
  if (squadreLoadState.status !== "loaded") return;

  const dateKey = getActiveSquadreDateKey();
  const assignments = getCurrentUserAssignedCommesseForDate(dateKey);
  startupAssignedCommessaAutoOpenDone = true;
  if (!assignments.length) return;

  const primary = assignments[0];
  selectCommessa(primary.commessaId, primary.commessaName, primary.commessa.codice || "");

  const otherAssignments = assignments.slice(1);
  if (otherAssignments.length) {
    const dateLabel = new Date(`${dateKey}T00:00:00`).toLocaleDateString("it-IT");
    const otherNames = otherAssignments.map((assignment) => `• ${assignment.commessaName}`).join("\n");
    alert(`Sei assegnato a più commesse per il ${dateLabel}. Ho aperto ${primary.commessaName}. Altre commesse trovate:\n${otherNames}`);
  }
}

function canCurrentUserInsertHoursForCommessa(commessaId, dateValue = "") {
  if (!currentUser) return false;
  if (canManageData()) return true;
  return Boolean(getCurrentUserSquadraAssignment(commessaId, dateValue));
}

function getHoursOperatorForCurrentUser(commessaId, dateValue = "") {
  const assignment = getCurrentUserSquadraAssignment(commessaId, dateValue);
  return assignment?.matchedName || currentUser?.displayName || currentUser?.email || "Operatore";
}

function getHoursRowsForCommessaSquadra(commessaId, dateValue = "") {
  const assignment = getCurrentUserSquadraAssignment(commessaId, dateValue);
  if (assignment) {
    return parseMultiEntryValue(assignment.row?.personale || "").map((name) => ({
      operatore: name,
      ore: "",
      squadraIndex: assignment.squadraIndex,
      squadraLabel: assignment.squadraLabel
    }));
  }
  if (canManageData()) {
    const squadData = getSquadraDataForCommessaDate(commessaId, dateValue);
    const squadRows = Array.isArray(squadData?.squadre) ? squadData.squadre : getLegacySquadreRows(squadData || {});
    return squadRows.flatMap((row, index) => parseMultiEntryValue(row?.personale || "").map((name) => ({
      operatore: name,
      ore: "",
      squadraIndex: index + 1,
      squadraLabel: `Squadra ${index + 1}`
    })));
  }
  return [];
}

function getHoursEntrySquadraIndexes(entry) {
  const indexes = new Set();
  const entryIndex = String(entry?.squadraIndex || "").trim();
  if (entryIndex) indexes.add(entryIndex);
  (Array.isArray(entry?.rows) ? entry.rows : []).forEach((row) => {
    const rowIndex = String(row?.squadraIndex || "").trim();
    if (rowIndex) indexes.add(rowIndex);
  });
  return indexes;
}

function doesHoursEntryMatchSquadra(entry, squadraIndex = "") {
  const targetIndex = String(squadraIndex || "").trim();
  if (!targetIndex) return true;
  const entryIndexes = getHoursEntrySquadraIndexes(entry);
  return !entryIndexes.size || entryIndexes.has(targetIndex);
}

function getHoursParticipantId(row = {}, entry = {}, options = {}) {
  const allowSquadraFallback = options.allowSquadraFallback !== false;
  const savedParticipantId = String(row.participantId || "").trim();
  if (savedParticipantId) return savedParticipantId;
  const directId = String(
    row.utenteId
    || row.userId
    || row.uid
    || row.operatoreId
    || row.personaleId
    || ""
  ).trim();
  if (directId) return `utente:${directId}`;
  const normalizedOperator = normalizeHoursOperatorName(row.operatore || row.nome || row.name || "");
  if (normalizedOperator) return `utente:${normalizedOperator}`;
  const squadraId = String(
    row.squadraId
    || entry.squadraId
    || row.squadraIndex
    || entry.squadraIndex
    || ""
  ).trim();
  if (allowSquadraFallback && squadraId) return `squadra:${squadraId}`;
  return "";
}

function buildHoursFullKey(commessaId, dateValue, participantId) {
  const id = String(commessaId || "").trim();
  const dateKey = String(dateValue || "").trim();
  const participant = String(participantId || "").trim();
  if (!id || !dateKey || !participant) return "";
  return `${id}__${dateKey}__${participant}`;
}

function getRequiredHoursParticipantsForCommessaDate(commessaId, dateValue) {
  const squadData = getSquadraDataForCommessaDate(commessaId, dateValue);
  const squadRows = Array.isArray(squadData?.squadre) ? squadData.squadre : getLegacySquadreRows(squadData || {});
  const participants = new Map();
  squadRows.forEach((row, index) => {
    if (!isSquadraRowFilled(row)) return;
    const squadraIndex = index + 1;
    const squadraLabel = `Squadra ${squadraIndex}`;
    const names = parseMultiEntryValue(row?.personale || "");
    if (names.length) {
      names.forEach((name) => {
        const participantId = getHoursParticipantId(
          { operatore: name, squadraIndex, squadraLabel },
          { squadraIndex, squadraLabel },
          { allowSquadraFallback: false }
        );
        const key = buildHoursFullKey(commessaId, dateValue, participantId);
        if (!key) return;
        participants.set(key, {
          key,
          participantId,
          operatore: name,
          squadraIndex,
          squadraLabel
        });
      });
      return;
    }
    const participantId = getHoursParticipantId(
      { squadraIndex, squadraLabel },
      { squadraIndex, squadraLabel },
      { allowSquadraFallback: true }
    );
    const key = buildHoursFullKey(commessaId, dateValue, participantId);
    if (!key) return;
    participants.set(key, {
      key,
      participantId,
      operatore: squadraLabel,
      squadraIndex,
      squadraLabel
    });
  });
  return participants;
}

function getCompletedHoursParticipantsForCommessaDate(commessaId, dateValue) {
  const id = String(commessaId || "").trim();
  const dateKey = String(dateValue || "").trim();
  const completed = new Map();
  if (!id || !dateKey) return completed;
  const sources = [
    ...allHoursReports,
    ...allHoursApprovalRequests.filter((request) => String(request.status || "").trim() !== "rejected")
  ];
  sources.forEach((record) => {
    if (String(record?.date || "").trim() !== dateKey) return;
    (Array.isArray(record?.entries) ? record.entries : []).forEach((entry) => {
      if (String(entry?.commessaId || "").trim() !== id) return;
      (Array.isArray(entry?.rows) ? entry.rows : []).forEach((row) => {
        if (Number(row?.ore || 0) <= 0) return;
        const participantId = getHoursParticipantId(row, entry, { allowSquadraFallback: true });
        const key = buildHoursFullKey(id, dateKey, participantId);
        if (!key) return;
        completed.set(key, {
          key,
          participantId,
          operatore: row?.operatore || "",
          squadraIndex: row?.squadraIndex || entry?.squadraIndex || "",
          squadraLabel: row?.squadraLabel || entry?.squadraLabel || ""
        });
      });
    });
  });
  return completed;
}

function getMissingHoursParticipantsForCommessaDate(commessaId, dateValue) {
  const required = getRequiredHoursParticipantsForCommessaDate(commessaId, dateValue);
  const completed = getCompletedHoursParticipantsForCommessaDate(commessaId, dateValue);
  return Array.from(required.values()).filter((participant) => !completed.has(participant.key));
}

function areAllHoursParticipantsCompleteForCommessaDate(commessaId, dateValue) {
  const required = getRequiredHoursParticipantsForCommessaDate(commessaId, dateValue);
  if (!required.size) return false;
  const completed = getCompletedHoursParticipantsForCommessaDate(commessaId, dateValue);
  return Array.from(required.keys()).every((key) => completed.has(key));
}

function hasHoursRecordForCommessaDateSquadra(commessaId, dateValue, squadraIndex = "") {
  const id = String(commessaId || "").trim();
  const dateKey = String(dateValue || "").trim();
  const targetIndex = String(squadraIndex || "").trim();
  if (!id || !dateKey) return false;
  const completed = getCompletedHoursParticipantsForCommessaDate(id, dateKey);
  return Array.from(completed.values()).some((participant) => {
    if (!targetIndex) return true;
    return String(participant.squadraIndex || "").trim() === targetIndex;
  });
}

function getQuickHoursContextForCommessa(commessaId, dateValue = "") {
  const dateKey = String(dateValue || "").trim() || getActiveSquadreDateKey();
  if (!hoursReportsLoaded || !hoursApprovalsLoaded) return null;
  const squadData = getSquadraDataForCommessaDate(commessaId, dateKey);
  const squadRows = Array.isArray(squadData?.squadre) ? squadData.squadre : getLegacySquadreRows(squadData || {});
  const hasAssignedSquadra = squadRows.some(isSquadraRowFilled);
  if (!dateKey || !hasAssignedSquadra) return null;
  const assignment = getCurrentUserSquadraAssignment(commessaId, dateKey);
  const squadraIndex = assignment?.squadraIndex || "";
  if (!canManageData() && !assignment) return null;
  if (hasHoursRecordForCommessaDateSquadra(commessaId, dateKey, squadraIndex)) return null;
  return { dateKey, assignment, squadData, squadRows, squadraIndex };
}

function createAddHoursButton(commessa, dateValue = "") {
  const button = document.createElement("button");
  button.type = "button";
  button.className = "btn add-hours-quick-btn";
  button.textContent = "+ ORE";
  button.dataset.addHoursCommessaId = commessa.id || "";
  button.setAttribute("aria-label", `Inserisci ore per ${commessa.nome || "commessa"}`);
  button.addEventListener("click", (event) => {
    event.preventDefault();
    event.stopPropagation();
    openHoursPageForCommessa(commessa.id, dateValue);
  });
  return button;
}

function appendAddHoursButtonIfAllowed(container, commessa, dateValue = "") {
  if (!container || !commessa?.id) return;
  const context = getQuickHoursContextForCommessa(commessa.id, dateValue);
  if (!context) return;
  container.appendChild(createAddHoursButton(commessa, context.dateKey));
}

function normalizeCommessaNavigationName(value) {
  return String(value || "").trim().toLocaleLowerCase("it-IT");
}

function getCommessaNavigationTarget(commessa = {}) {
  const id = String(commessa.id || "").trim();
  const name = String(commessa.nome || commessa.commessaNome || commessa.name || "").trim();
  if (id) {
    return {
      id,
      nome: commessa.nome || name || "Commessa",
      codice: commessa.codice || ""
    };
  }
  if (!name) return null;
  const normalizedName = normalizeCommessaNavigationName(name);
  const matchedCommessa = Array.from(commesseById.values()).find((item) => (
    normalizeCommessaNavigationName(item?.nome) === normalizedName
  ));
  if (matchedCommessa) {
    return {
      id: matchedCommessa.id,
      nome: matchedCommessa.nome || name,
      codice: matchedCommessa.codice || ""
    };
  }
  return {
    id: name,
    nome: name,
    codice: commessa.codice || ""
  };
}

function openCommessaFromSquadre(commessa = {}) {
  const target = getCommessaNavigationTarget(commessa);
  if (!target?.id) {
    alert("Commessa non disponibile.");
    return;
  }
  selectCommessa(target.id, target.nome || "Commessa", target.codice || "");
}

function openHoursPageForCommessa(commessaId, dateValue = "") {
  if (!currentUser) {
    alert("Devi fare login per inserire le ore.");
    return;
  }
  const id = String(commessaId || "").trim();
  if (!id || !commesseById.has(id)) {
    alert("Commessa non disponibile per l'inserimento ore.");
    return;
  }
  const targetDateValue = String(dateValue || "").trim() || getActiveSquadreDateKey() || new Date().toISOString().slice(0, 10);
  if (!canCurrentUserInsertHoursForCommessa(id, targetDateValue)) {
    alert("Permesso negato: puoi inserire ore solo per le commesse dove sei assegnato in squadra.");
    return;
  }
  const assignment = getCurrentUserSquadraAssignment(id, targetDateValue);
  const squadraIndex = assignment?.squadraIndex || "";
  if (hasHoursRecordForCommessaDateSquadra(id, targetDateValue, squadraIndex)) {
    alert("Le ore per oggi sono già state inserite e sono visibili in Visualizza ore.");
    renderSquadre();
    return;
  }

  openHoursPage();
  if (ui.hoursDate) ui.hoursDate.value = targetDateValue;
  if (ui.hoursCommesseList) {
    ui.hoursCommesseList.innerHTML = "";
    const card = addHoursCommessaBlock({ commessaId: id });
    applyHoursSuggestedOperators(card, { force: true });
  }
  if (ui.hoursFeedback) {
    const commessaName = commesseById.get(id)?.nome || "Commessa";
    ui.hoursFeedback.textContent = `Compila le ore per ${commessaName}: il pulsante +ORE apre il form standard Gestione ore.`;
  }
  setTimeout(() => ui.hoursForm?.scrollIntoView({ behavior: "smooth", block: "start" }), 50);
}

function getSuggestedHoursOperators(commessaId, dateValue) {
  const dateKey = String(dateValue || "").trim();
  if (!commessaId || !dateKey) return [];
  const storicoByDate = squadreHistoryByDate.get(dateKey) || new Map();
  const squadData = storicoByDate.get(commessaId) || {};
  const squadRows = Array.isArray(squadData.squadre) ? squadData.squadre : getLegacySquadreRows(squadData);
  const names = squadRows
    .flatMap((row) => parseMultiEntryValue(row.personale || ""))
    .map((name) => String(name || "").trim())
    .filter(Boolean);
  return [...new Set(names)];
}

function renderHoursOperatorSuggestions(card) {
  if (!card) return;
  const suggestionBox = card.querySelector(".hours-operator-suggestions");
  if (!suggestionBox) return;
  const operatorList = card.querySelector(".hours-operator-list");
  const existing = new Set(Array.from(operatorList.querySelectorAll(".hours-operatore"))
    .map((input) => String(input.value || "").trim().toLowerCase())
    .filter(Boolean));
  const source = personaleRecords.map((person) => getPersonaleDisplayName(person)).filter(Boolean);
  const matches = source
    .filter((name) => !existing.has(name.toLowerCase()))
    .slice(0, 8);
  suggestionBox.innerHTML = matches.map((name) => (
    `<button type="button" class="btn btn-small" data-hours-suggested-operator="${escapeHTML(name)}">${escapeHTML(name)}</button>`
  )).join("");
  suggestionBox.classList.toggle("hidden", matches.length === 0);
  suggestionBox.querySelectorAll("[data-hours-suggested-operator]").forEach((btn) => {
    btn.addEventListener("click", () => {
      unlockHoursFinalizeButton();
      addHoursOperatoreRow(operatorList, { operatore: btn.dataset.hoursSuggestedOperator || "", ore: "" }, card);
      renderHoursOperatorSuggestions(card);
      renderHoursSummary();
    });
  });
}

function addHoursOperatoreRow(container, rowData = { operatore: "", ore: "" }, card = null) {
  const row = document.createElement("div");
  row.className = "hours-operator-row";
  if (rowData.squadraIndex) row.dataset.squadraIndex = String(rowData.squadraIndex || "");
  if (rowData.squadraLabel) row.dataset.squadraLabel = String(rowData.squadraLabel || "");
  row.innerHTML = `
    <input type="text" class="hours-operatore" list="hours-operatori-options" placeholder="Operatore" value="${escapeHTML(rowData.operatore || "")}" autocomplete="off">
    <input type="number" class="hours-ore" min="0" max="24" step="0.25" placeholder="Ore" value="${escapeHTML(rowData.ore || "")}">
    <button type="button" class="btn btn-small hours-remove-operator-btn" aria-label="Rimuovi operatore">✕</button>
  `;
  const removeBtn = row.querySelector(".hours-remove-operator-btn");
  removeBtn.addEventListener("click", () => {
    unlockHoursFinalizeButton();
    row.remove();
    if (!container.children.length) addHoursOperatoreRow(container, {}, card);
    renderHoursOperatorSuggestions(card);
    renderHoursSummary();
  });
  row.querySelectorAll("input").forEach((input) => input.addEventListener("input", () => {
    unlockHoursFinalizeButton();
    renderHoursOperatorSuggestions(card);
    renderHoursSummary();
  }));
  container.appendChild(row);
}

function applyHoursSuggestedOperators(card, options = {}) {
  if (!card) return;
  const dateValue = String(ui.hoursDate?.value || "").trim();
  const commessaId = String(card.querySelector(".hours-commessa-select")?.value || "").trim();
  const operatorList = card.querySelector(".hours-operator-list");
  const suggestions = getSuggestedHoursOperators(commessaId, dateValue);
  const sourceTag = `${commessaId}__${dateValue}`;
  const force = options.force === true;
  const alreadyApplied = String(card.dataset.suggestedSource || "") === sourceTag;
  if (!force && alreadyApplied) return;
  operatorList.innerHTML = "";
  if (suggestions.length) {
    suggestions.forEach((name) => addHoursOperatoreRow(operatorList, { operatore: name, ore: "" }, card));
  } else {
    addHoursOperatoreRow(operatorList, {}, card);
  }
  card.dataset.suggestedSource = sourceTag;
  renderHoursOperatorSuggestions(card);
  renderHoursSummary();
}

function addHoursCommessaBlock(blockData = null) {
  const card = document.createElement("article");
  card.className = "hours-commessa-card";
  card.innerHTML = `
    <div class="hours-commessa-head">
      <h3>Commessa</h3>
      <div class="item-actions">
        <button type="button" class="btn hours-compact-pill hours-export-global-btn">📊 Excel</button>
        <button type="button" class="btn hours-compact-pill hours-remove-commessa-btn">🗑 Rimuovi</button>
      </div>
    </div>
    <select class="hours-commessa-select" required>
      <option value="">Seleziona commessa</option>
    </select>
    <p class="hours-locked-commessa-label muted hidden"></p>
    <div class="hours-commesse-buttons hours-commessa-picker" aria-label="Seleziona commessa"></div>
    <p class="hours-team-label muted hidden"></p>
    <p class="hours-inserted-by-label muted hidden"></p>
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
  const exportGlobalBtn = card.querySelector(".hours-export-global-btn");
  exportGlobalBtn?.addEventListener("click", exportHoursGlobalMonthlyTable);
  const operatorList = card.querySelector(".hours-operator-list");
  card.querySelector(".hours-add-operator-btn").addEventListener("click", () => {
    unlockHoursFinalizeButton();
    addHoursOperatoreRow(operatorList, {}, card);
    renderHoursOperatorSuggestions(card);
    renderHoursSummary();
  });
  const commessaSelect = card.querySelector(".hours-commessa-select");
  commessaSelect.addEventListener("change", () => {
    unlockHoursFinalizeButton();
    renderHoursCardCommessaButtons(card);
    applyHoursSuggestedOperators(card, { force: true });
  });
  card.querySelector(".hours-note").addEventListener("input", () => {
    unlockHoursFinalizeButton();
    renderHoursSummary();
  });

  if (blockData) {
    if (blockData.commessaId) {
      commessaSelect.value = blockData.commessaId;
    }
    if (blockData.lockedCommessa) {
      card.dataset.lockedCommessaId = blockData.commessaId || "";
      card.dataset.lockedCommessa = "true";
      commessaSelect.classList.add("hidden");
      commessaSelect.disabled = true;
      const lockedLabel = card.querySelector(".hours-locked-commessa-label");
      if (lockedLabel) {
        lockedLabel.textContent = `Commessa: ${commesseById.get(blockData.commessaId)?.nome || "Commessa"}`;
        lockedLabel.classList.remove("hidden");
      }
      const buttons = card.querySelector(".hours-commesse-buttons");
      buttons?.classList.add("hidden");
    }
    if (blockData.squadraLabel) {
      card.dataset.squadraIndex = String(blockData.squadraIndex || "");
      card.dataset.squadraLabel = String(blockData.squadraLabel || "");
      const teamLabel = card.querySelector(".hours-team-label");
      if (teamLabel) {
        teamLabel.textContent = `Squadra: ${blockData.squadraLabel}`;
        teamLabel.classList.remove("hidden");
      }
    }
    if (blockData.insertedBy) {
      const insertedByLabel = card.querySelector(".hours-inserted-by-label");
      if (insertedByLabel) {
        insertedByLabel.textContent = `Inserito da: ${blockData.insertedBy}`;
        insertedByLabel.classList.remove("hidden");
      }
    }
    card.querySelector(".hours-note").value = blockData.note || "";
    const rows = Array.isArray(blockData.rows) && blockData.rows.length ? blockData.rows : [{}];
    rows.forEach((row) => addHoursOperatoreRow(operatorList, row, card));
    renderHoursOperatorSuggestions(card);
  } else {
    applyHoursSuggestedOperators(card, { force: true });
  }
  renderHoursCardCommessaButtons(card);
  renderHoursSummary();
  return card;
}

function collectHoursEntries() {
  const cards = Array.from(ui.hoursCommesseList.querySelectorAll(".hours-commessa-card"));
  return cards.map((card) => {
    const commessaId = String(card.dataset.lockedCommessaId || card.querySelector(".hours-commessa-select")?.value || "").trim();
    const note = String(card.querySelector(".hours-note")?.value || "").trim();
    const squadraIndex = String(card.dataset.squadraIndex || "").trim();
    const squadraLabel = String(card.dataset.squadraLabel || "").trim();
    const rows = Array.from(card.querySelectorAll(".hours-operator-row")).map((row) => ({
      operatore: String(row.querySelector(".hours-operatore")?.value || "").trim(),
      ore: Number(row.querySelector(".hours-ore")?.value || 0),
      operatoreId: resolveHoursOperatorId(row.querySelector(".hours-operatore")?.value || ""),
      squadraIndex: String(row.dataset.squadraIndex || squadraIndex || "").trim(),
      squadraLabel: String(row.dataset.squadraLabel || squadraLabel || "").trim()
    })).filter((row) => row.operatore && row.ore > 0);
    const commessaName = commesseById.get(commessaId)?.nome || "";
    return { commessaId, commessaName, note, squadraIndex, squadraLabel, rows };
  }).filter((entry) => entry.commessaId || entry.rows.length || entry.note);
}

function normalizeHoursOperatorName(value) {
  return String(value || "").toLocaleLowerCase("it-IT").replace(/\s+/g, " ").trim();
}

function resolveHoursOperatorId(operatore) {
  const normalizedOperator = normalizeHoursOperatorName(operatore);
  if (!normalizedOperator) return "";
  const person = personaleRecords.find((item) => normalizeHoursOperatorName(getPersonaleDisplayName(item)) === normalizedOperator);
  return String(person?.id || normalizedOperator).trim();
}

function getHoursOperatorUniquePart(row = {}) {
  const operatore = String(row?.operatore || row?.nome || row?.name || "").trim();
  const normalizedOperator = normalizeHoursOperatorName(operatore);
  return String(row?.operatoreId || row?.participantId || row?.personaleId || row?.userId || row?.uid || resolveHoursOperatorId(operatore) || normalizedOperator).trim();
}

function encodeHoursLockPart(value) {
  return encodeURIComponent(String(value || "").trim());
}

function getHoursLockDocId(dateValue, commessaId, operatore) {
  const normalizedOperator = normalizeHoursOperatorName(operatore);
  if (!dateValue || !commessaId || !normalizedOperator) return "";
  return [dateValue, commessaId, normalizedOperator].map(encodeHoursLockPart).join("__");
}

function getHoursUniqueLocks(dateValue, entries) {
  const locks = new Map();
  const dateKey = String(dateValue || "").trim();
  if (!dateKey) return [];
  (Array.isArray(entries) ? entries : []).forEach((entry) => {
    const commessaInfo = resolveHoursEntryCommessa(entry);
    const commessaId = String(commessaInfo.id || commessaInfo.key || "").trim();
    if (!commessaId) return;
    (Array.isArray(entry?.rows) ? entry.rows : []).forEach((row) => {
      const operatore = String(row?.operatore || "").trim();
      const normalizedOperator = normalizeHoursOperatorName(operatore);
      const operatoreId = String(row?.operatoreId || row?.participantId || resolveHoursOperatorId(operatore) || normalizedOperator).trim();
      const operatorIdentity = getHoursOperatorIdentity(row);
      const ore = Number(row?.ore || 0);
      const lockId = getHoursLockDocId(dateKey, commessaId, operatorIdentity || operatoreId || operatore);
      if (!lockId || !normalizedOperator || !operatorIdentity || ore <= 0) return;
      if (!locks.has(lockId)) {
        locks.set(lockId, {
          id: lockId,
          date: dateKey,
          commessaId,
          commessaName: commessaInfo.nome || entry.commessaName || commesseById.get(commessaId)?.nome || "Commessa",
          operatore,
          operatoreId,
          uniqueKey: row?.uniqueKey || buildHoursUniqueKey(dateKey, commessaId, row) || `${commessaId}_${dateKey}_${operatoreId || normalizedOperator}`,
          normalizedOperator
        });
      }
    });
  });
  return Array.from(locks.values());
}

function isActiveHoursLock(lockData = {}, skipApprovalRequestId = "") {
  const status = String(lockData.status || "").trim();
  if (["rejected", "deleted", "void"].includes(status)) return false;
  if (skipApprovalRequestId && String(lockData.approvalRequestId || "") === skipApprovalRequestId) return false;
  return true;
}

async function reserveHoursApprovalRequestWithLocks(payload) {
  payload.entries = addHoursUniqueKeysToEntries(payload?.date, payload?.entries);
  const locks = getHoursUniqueLocks(payload?.date, payload?.entries);
  const approvalRef = db.collection("oreApprovalRequests").doc();
  await db.runTransaction(async (transaction) => {
    const lockRefs = locks.map((lock) => ({
      lock,
      ref: db.collection("oreLocks").doc(lock.id)
    }));
    const lockSnapshots = await Promise.all(lockRefs.map(({ ref }) => transaction.get(ref)));
    const conflicts = [];
    lockSnapshots.forEach((snap, index) => {
      if (!snap.exists) return;
      const data = snap.data() || {};
      if (!isActiveHoursLock(data)) return;
      const fallback = lockRefs[index].lock;
      conflicts.push({
        commessaName: data.commessaName || fallback.commessaName || "Commessa",
        operatore: data.operatore || fallback.operatore || "Operatore"
      });
    });
    if (conflicts.length) {
      const error = new Error("Le ore per questo giorno, commessa e operatore sono già state inserite.");
      error.code = "hours-duplicate-lock";
      error.conflicts = conflicts;
      throw error;
    }

    transaction.set(approvalRef, {
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
    lockRefs.forEach(({ lock, ref }) => {
      transaction.set(ref, {
        ...lock,
        approvalRequestId: approvalRef.id,
        reportId: "",
        status: "pending_level1",
        createdByUid: payload.createdByUid || "",
        createdByEmail: payload.createdByEmail || "",
        createdByName: payload.createdByName || "Operatore",
        createdAt: firebase.firestore.FieldValue.serverTimestamp(),
        updatedAt: firebase.firestore.FieldValue.serverTimestamp()
      }, { merge: true });
    });
  });
  return approvalRef;
}

async function updateHoursLocksForEntries(dateValue, entries, updates = {}) {
  const locks = getHoursUniqueLocks(dateValue, entries);
  if (!locks.length) return;
  const batch = db.batch();
  locks.forEach((lock) => {
    batch.set(db.collection("oreLocks").doc(lock.id), {
      ...lock,
      ...updates,
      updatedAt: firebase.firestore.FieldValue.serverTimestamp()
    }, { merge: true });
  });
  await batch.commit();
}

function getHoursTimestampMillis(value) {
  if (!value) return 0;
  if (typeof value.toMillis === "function") return value.toMillis();
  if (typeof value === "number" && Number.isFinite(value)) return value;
  if (typeof value === "object" && Number.isFinite(Number(value.seconds))) return Number(value.seconds) * 1000;
  const parsed = new Date(value).getTime();
  return Number.isFinite(parsed) ? parsed : 0;
}

function getHoursRecordTimestampMillis(record = {}) {
  return Math.max(
    getHoursTimestampMillis(record.updatedAt),
    getHoursTimestampMillis(record.level2ApprovedAt),
    getHoursTimestampMillis(record.approvedAt),
    getHoursTimestampMillis(record.createdAt),
    getHoursTimestampMillis(record.submittedAt)
  );
}

function getHoursOperatorIdentity(row = {}) {
  const email = String(row?.operatoreEmail || row?.email || row?.createdByEmail || "").trim().toLocaleLowerCase("it-IT");
  if (email) return email;
  const name = String(row?.operatoreNome || row?.operatore || row?.nome || row?.name || "").trim();
  return getHoursOperatorUniquePart({ ...row, operatore: name }) || normalizeHoursOperatorName(name);
}

function buildHoursUniqueKey(dateValue, commessaId, row = {}) {
  const dateKey = String(dateValue || "").trim();
  const commessaKey = String(commessaId || "").trim();
  const operatorKey = getHoursOperatorIdentity(row);
  if (!dateKey || !commessaKey || !operatorKey) return "";
  return `${commessaKey}_${dateKey}_${operatorKey}`;
}

function addHoursUniqueKeysToEntries(dateValue, entries = []) {
  return (Array.isArray(entries) ? entries : []).map((entry) => {
    const commessaId = String(resolveHoursEntryCommessa(entry).id || resolveHoursEntryCommessa(entry).key || "").trim();
    return {
      ...entry,
      rows: (Array.isArray(entry?.rows) ? entry.rows : []).map((row) => ({
        ...row,
        uniqueKey: row?.uniqueKey || buildHoursUniqueKey(dateValue, commessaId, row)
      }))
    };
  });
}


function isAdminConfirmedHoursRecord(record = {}, collectionName = "") {
  return String(collectionName || record.sourceCollection || "") === "oreReports"
    || String(record.status || record.approvalStatus || "").trim() === "approved"
    || Boolean(record.level2ApprovedAt || record.finalizedReportId);
}

function compareHoursRowPriority(a, b) {
  const confirmedDiff = Number(isAdminConfirmedHoursRecord(b.record, b.collectionName)) - Number(isAdminConfirmedHoursRecord(a.record, a.collectionName));
  if (confirmedDiff) return confirmedDiff;
  const reportDiff = Number(b.collectionName === "oreReports") - Number(a.collectionName === "oreReports");
  if (reportDiff) return reportDiff;
  const bTs = getHoursRecordTimestampMillis(b.record);
  const aTs = getHoursRecordTimestampMillis(a.record);
  const validDiff = Number(bTs > 0) - Number(aTs > 0);
  if (validDiff) return validDiff;
  if (bTs !== aTs) return bTs - aTs;
  return String(b.docId || "").localeCompare(String(a.docId || ""));
}

function collectHoursRowRefs(records = []) {
  const refs = [];
  (Array.isArray(records) ? records : []).forEach((recordWrapper) => {
    const record = recordWrapper.data || recordWrapper;
    const dateValue = String(record.date || "").trim();
    const collectionName = String(recordWrapper.collectionName || record.sourceCollection || "oreReports");
    const docId = String(recordWrapper.id || record.id || "").trim();
    const entries = Array.isArray(record.entries) ? record.entries : [];
    entries.forEach((entry, entryIndex) => {
      const commessaInfo = resolveHoursEntryCommessa(entry);
      const commessaId = String(commessaInfo.id || commessaInfo.key || "").trim();
      if (!commessaId) return;
      (Array.isArray(entry?.rows) ? entry.rows : []).forEach((row, rowIndex) => {
        const ore = Number(row?.ore || 0);
        const operatorIdentity = getHoursOperatorIdentity(row);
        const uniqueKey = buildHoursUniqueKey(dateValue, commessaId, row);
        if (!uniqueKey || !operatorIdentity || ore <= 0) return;
        refs.push({
          collectionName,
          docId,
          ref: recordWrapper.ref || null,
          record,
          dateValue,
          entry,
          entryIndex,
          row,
          rowIndex,
          ore,
          uniqueKey,
          operatorIdentity
        });
      });
    });
  });
  return refs;
}

function pickUniqueHoursRows(rowRefs = []) {
  const grouped = new Map();
  rowRefs.forEach((rowRef) => {
    if (!grouped.has(rowRef.uniqueKey)) grouped.set(rowRef.uniqueKey, []);
    grouped.get(rowRef.uniqueKey).push(rowRef);
  });
  const keepRefs = new Set();
  const duplicateRefs = [];
  grouped.forEach((items) => {
    const [keeper, ...duplicates] = [...items].sort(compareHoursRowPriority);
    if (keeper) keepRefs.add(keeper);
    duplicateRefs.push(...duplicates);
  });
  return { keepRefs, duplicateRefs, grouped };
}

function deduplicateHoursRecordsForDisplay(records = []) {
  const wrappers = (Array.isArray(records) ? records : []).map((record) => ({
    ...record,
    data: record,
    collectionName: record.sourceCollection || "oreReports"
  }));
  const rowRefs = collectHoursRowRefs(wrappers);
  const { keepRefs } = pickUniqueHoursRows(rowRefs);
  return wrappers.map((wrapper) => {
    const record = wrapper.data || wrapper;
    const nextEntries = (Array.isArray(record.entries) ? record.entries : []).map((entry, entryIndex) => {
      const rows = (Array.isArray(entry.rows) ? entry.rows : [])
        .map((row, rowIndex) => {
          const match = rowRefs.find((rowRef) => rowRef.docId === String(record.id || "")
            && rowRef.collectionName === String(record.sourceCollection || "oreReports")
            && rowRef.entryIndex === entryIndex
            && rowRef.rowIndex === rowIndex);
          if (!match || !keepRefs.has(match)) return null;
          return { ...row, uniqueKey: match.uniqueKey };
        })
        .filter(Boolean);
      return { ...entry, rows };
    }).filter((entry) => Array.isArray(entry.rows) && entry.rows.length);
    return { ...record, entries: nextEntries };
  }).filter((record) => Array.isArray(record.entries) && record.entries.length);
}

async function syncHoursRepairToRealtimeDatabase(payload = {}) {
  if (!firebase.database || typeof firebase.database !== "function") return;
  const realtimeDb = firebase.database();
  if (!realtimeDb?.ref) return;
  await realtimeDb.ref("hoursDuplicateRepair").set({
    ...payload,
    updatedAt: Date.now()
  });
}

async function repairDuplicateHours(options = {}) {
  if (!currentUser || !canManageData()) return { changed: false, repaired: 0, deleted: 0, duplicates: 0 };
  if (hoursDuplicateCleanupPromise && options.force !== true) return hoursDuplicateCleanupPromise;
  hoursDuplicateCleanupPromise = (async () => {
    const [reportsSnapshot, approvalsSnapshot] = await Promise.all([
      db.collection("oreReports").get(),
      db.collection("oreApprovalRequests").get()
    ]);
    const docs = [
      ...reportsSnapshot.docs.map((doc) => ({ id: doc.id, ref: doc.ref, collectionName: "oreReports", data: doc.data() || {} })),
      ...approvalsSnapshot.docs
        .map((doc) => ({ id: doc.id, ref: doc.ref, collectionName: "oreApprovalRequests", data: doc.data() || {} }))
        .filter((doc) => !["rejected", "deleted", "void"].includes(String(doc.data.status || "").trim()))
    ];
    const rowRefs = collectHoursRowRefs(docs);
    const { keepRefs, duplicateRefs } = pickUniqueHoursRows(rowRefs);
    const changedDocs = new Map();
    const getChange = (doc) => {
      const key = `${doc.collectionName}::${doc.id}`;
      if (!changedDocs.has(key)) changedDocs.set(key, { doc, changed: false, entries: [] });
      return changedDocs.get(key);
    };

    docs.forEach((doc) => {
      const entries = Array.isArray(doc.data.entries) ? doc.data.entries : [];
      const nextEntries = entries.map((entry, entryIndex) => {
        const rows = Array.isArray(entry.rows) ? entry.rows : [];
        const nextRows = rows.map((row, rowIndex) => {
          const match = rowRefs.find((rowRef) => rowRef.collectionName === doc.collectionName
            && rowRef.docId === doc.id
            && rowRef.entryIndex === entryIndex
            && rowRef.rowIndex === rowIndex);
          if (!match) return null;
          if (!keepRefs.has(match)) return null;
          const cleanRow = { ...row, uniqueKey: match.uniqueKey };
          if (row.uniqueKey !== match.uniqueKey) getChange(doc).changed = true;
          return cleanRow;
        }).filter(Boolean);
        if (nextRows.length !== rows.length) getChange(doc).changed = true;
        return { ...entry, rows: nextRows };
      }).filter((entry) => Array.isArray(entry.rows) && entry.rows.length);
      if (nextEntries.length !== entries.length) getChange(doc).changed = true;
      const change = getChange(doc);
      change.entries = nextEntries;
    });

    const writes = Array.from(changedDocs.values()).filter((change) => change.changed);
    for (let index = 0; index < writes.length; index += 450) {
      const batch = db.batch();
      writes.slice(index, index + 450).forEach(({ doc, entries }) => {
        if (!entries.length) {
          batch.delete(doc.ref);
          return;
        }
        batch.update(doc.ref, {
          entries,
          duplicateHoursRepairedAt: firebase.firestore.FieldValue.serverTimestamp(),
          updatedAt: firebase.firestore.FieldValue.serverTimestamp()
        });
      });
      await batch.commit();
    }

    const keptByDoc = new Map();
    rowRefs.forEach((rowRef) => {
      if (!keepRefs.has(rowRef)) return;
      const key = `${rowRef.collectionName}::${rowRef.docId}`;
      if (!keptByDoc.has(key)) keptByDoc.set(key, []);
      const entryCopy = { ...rowRef.entry, rows: [{ ...rowRef.row, uniqueKey: rowRef.uniqueKey }] };
      keptByDoc.get(key).push({ rowRef, entry: entryCopy });
    });

    const lockWrites = [];
    keptByDoc.forEach((items) => {
      items.forEach(({ rowRef, entry }) => {
        getHoursUniqueLocks(rowRef.dateValue, [entry]).forEach((lock) => {
          lockWrites.push({ type: "set", lock, rowRef });
        });
      });
    });
    for (let index = 0; index < lockWrites.length; index += 150) {
      const batch = db.batch();
      lockWrites.slice(index, index + 150).forEach(({ type, lock, rowRef }) => {
        const ref = db.collection("oreLocks").doc(lock.id);
        batch.set(ref, {
          ...lock,
          uniqueKey: rowRef.uniqueKey,
          approvalRequestId: rowRef.collectionName === "oreApprovalRequests" ? rowRef.docId : "",
          reportId: rowRef.collectionName === "oreReports" ? rowRef.docId : (rowRef.record.finalizedReportId || ""),
          status: isAdminConfirmedHoursRecord(rowRef.record, rowRef.collectionName) ? "approved" : String(rowRef.record.status || "pending_level1"),
          repairedAt: firebase.firestore.FieldValue.serverTimestamp(),
          updatedAt: firebase.firestore.FieldValue.serverTimestamp()
        }, { merge: true });
      });
      await batch.commit();
    }

    const result = { changed: Boolean(writes.length || duplicateRefs.length), repaired: writes.length, deleted: writes.filter((write) => !write.entries.length).length, duplicates: duplicateRefs.length };
    await syncHoursRepairToRealtimeDatabase(result);
    renderHoursSummary();
    return result;
  })().catch((error) => {
    hoursDuplicateCleanupPromise = null;
    throw error;
  }).finally(() => {
    hoursDuplicateCleanupPromise = null;
  });
  return hoursDuplicateCleanupPromise;
}

async function repairDuplicateHoursReports(options = {}) {
  return repairDuplicateHours(options);
}

async function ensureHoursReportsDeduplicated() {
  if (!currentUser || !canManageData()) return;
  try {
    const result = await repairDuplicateHours();
    if (result?.changed) {
      console.info(`Pulizia ore duplicate completata: ${result.repaired} documenti aggiornati, ${result.deleted} eliminati, ${result.duplicates} duplicati rimossi.`);
    }
  } catch (error) {
    console.error("Pulizia automatica duplicati ore non riuscita:", error);
  }
}

window.repairDuplicateHours = repairDuplicateHours;

async function createApprovedHoursReportWithLocks(request, reportPayload) {
  const reportRef = db.collection("oreReports").doc();
  reportPayload.entries = addHoursUniqueKeysToEntries(reportPayload.date, reportPayload.entries);
  const locks = getHoursUniqueLocks(reportPayload.date, reportPayload.entries);
  await db.runTransaction(async (transaction) => {
    const lockRefs = locks.map((lock) => ({ lock, ref: db.collection("oreLocks").doc(lock.id) }));
    const lockSnapshots = await Promise.all(lockRefs.map(({ ref }) => transaction.get(ref)));
    const conflicts = [];
    lockSnapshots.forEach((snap, index) => {
      if (!snap.exists) return;
      const data = snap.data() || {};
      if (!isActiveHoursLock(data, request.id || "")) return;
      const fallback = lockRefs[index].lock;
      conflicts.push({
        commessaName: data.commessaName || fallback.commessaName || "Commessa",
        operatore: data.operatore || fallback.operatore || "Operatore"
      });
    });
    if (conflicts.length) {
      const error = new Error("Le ore per questo giorno, commessa e operatore sono già state inserite.");
      error.code = "hours-duplicate-lock";
      error.conflicts = conflicts;
      throw error;
    }
    transaction.set(reportRef, reportPayload);
    lockRefs.forEach(({ lock, ref }) => {
      transaction.set(ref, {
        ...lock,
        approvalRequestId: request.id || "",
        reportId: reportRef.id,
        status: "approved",
        createdByUid: reportPayload.createdByUid || "",
        createdByEmail: reportPayload.createdByEmail || "",
        createdByName: reportPayload.createdByName || "Operatore",
        updatedAt: firebase.firestore.FieldValue.serverTimestamp()
      }, { merge: true });
    });
  });
  return reportRef;
}


function buildHoursEntriesFromApprovalSources(request, sources = []) {
  const targetsByEntry = new Map();
  (Array.isArray(sources) ? sources : []).forEach((source) => {
    const entryIndex = Number(source.entryIndex);
    const rowIndex = Number(source.rowIndex);
    if (!Number.isInteger(entryIndex) || !Number.isInteger(rowIndex)) return;
    if (!targetsByEntry.has(entryIndex)) targetsByEntry.set(entryIndex, new Set());
    targetsByEntry.get(entryIndex).add(rowIndex);
  });
  const approvedEntries = [];
  const remainingEntries = [];
  (Array.isArray(request?.entries) ? request.entries : []).forEach((entry, entryIndex) => {
    const targetRows = targetsByEntry.get(entryIndex);
    const approvedRows = [];
    const remainingRows = [];
    (Array.isArray(entry.rows) ? entry.rows : []).forEach((row, rowIndex) => {
      if (!targetsByEntry.size || targetRows?.has(rowIndex)) approvedRows.push(row);
      else remainingRows.push(row);
    });
    if (approvedRows.length) approvedEntries.push({ ...entry, rows: approvedRows });
    if (targetsByEntry.size && remainingRows.length) remainingEntries.push({ ...entry, rows: remainingRows });
  });
  return {
    approvedEntries,
    remainingEntries,
    isPartial: targetsByEntry.size > 0
  };
}

async function saveApprovedHoursRequest(request, options = {}) {
  if (!canManageData()) {
    throw new Error("Solo admin può confermare il livello finale.");
  }
  if (!request?.id) {
    throw new Error("ID richiesta ore non valido.");
  }
  if (String(request.status || "") !== "pending_admin") {
    throw new Error("Questa richiesta non è più in attesa dell'approvazione admin.");
  }
  const { approvedEntries, remainingEntries, isPartial } = buildHoursEntriesFromApprovalSources(request, options.sources || []);
  if (!approvedEntries.length) {
    throw new Error("Nessuna ora da confermare trovata.");
  }
  const duplicateDraft = findDuplicateHoursInDraft(approvedEntries);
  if (duplicateDraft.length) {
    const error = new Error(formatHoursDuplicateMessage(duplicateDraft, { admin: true }));
    error.code = "hours-duplicate-draft";
    error.conflicts = duplicateDraft;
    throw error;
  }
  const dateValue = String(request.date || options.fallbackDate || "").trim();
  const conflicts = await findExistingHoursConflicts(dateValue, approvedEntries, { skipApprovalRequestId: request.id });
  if (conflicts.length) {
    const error = new Error(formatHoursDuplicateMessage(conflicts, { admin: true }));
    error.code = "hours-duplicate-lock";
    error.conflicts = conflicts;
    throw error;
  }
  const reportPayload = {
    date: dateValue,
    entries: approvedEntries,
    createdByUid: request.createdByUid || "",
    createdByName: request.createdByName || request.createdByEmail || "Operatore",
    createdByEmail: request.createdByEmail || "",
    createdAt: firebase.firestore.FieldValue.serverTimestamp()
  };
  let docRef;
  try {
    docRef = await createApprovedHoursReportWithLocks(request, reportPayload);
  } catch (error) {
    if (error?.code === "hours-duplicate-lock") {
      error.message = formatHoursDuplicateMessage(error.conflicts, { admin: true });
    }
    throw error;
  }
  let driveLink = "";
  if (!isPartial && driveAccessToken) {
    if (!driveReportsFolderId) await ensureDriveFolders();
    const blob = new Blob([JSON.stringify({ id: docRef.id, ...reportPayload, createdAtIso: new Date().toISOString() }, null, 2)], { type: "application/json" });
    const fileName = `ore_${reportPayload.date}_${docRef.id}.json`;
    const upload = await uploadBlobToDrive(blob, fileName, "application/json", driveReportsFolderId, { driveType: "ORE", commessaName: "ORE" });
    driveLink = upload?.webViewLink || "";
  }
  const requestUpdate = isPartial && remainingEntries.length
    ? {
        entries: remainingEntries,
        partialFinalizedReportIds: firebase.firestore.FieldValue.arrayUnion(docRef.id),
        level2ApprovedBy: currentUser.email || currentUser.displayName || "admin",
        level2ApprovedAt: firebase.firestore.FieldValue.serverTimestamp(),
        updatedAt: firebase.firestore.FieldValue.serverTimestamp()
      }
    : {
        entries: [],
        status: "approved",
        level2ApprovedBy: currentUser.email || currentUser.displayName || "admin",
        level2ApprovedAt: firebase.firestore.FieldValue.serverTimestamp(),
        finalizedReportId: docRef.id,
        driveBackupLink: driveLink,
        updatedAt: firebase.firestore.FieldValue.serverTimestamp()
      };
  await db.collection("oreApprovalRequests").doc(request.id).set(requestUpdate, { merge: true });
  await updateHoursLocksForEntries(dateValue, approvedEntries, {
    status: "approved",
    approvalRequestId: request.id || "",
    reportId: docRef.id
  });
  return { reportId: docRef.id, approvedEntries, remainingEntries };
}

function formatHoursDuplicateMessage(conflicts = [], options = {}) {
  const unique = [];
  const seen = new Set();
  (Array.isArray(conflicts) ? conflicts : []).forEach((conflict) => {
    const label = `${conflict.commessaName || "Commessa"}: ${conflict.operatore || "Operatore"}`;
    const key = normalizeHoursOperatorName(label);
    if (seen.has(key)) return;
    seen.add(key);
    unique.push(label);
  });
  const details = unique.length ? ` (${unique.slice(0, 3).join("; ")}${unique.length > 3 ? "; ..." : ""})` : "";
  return options?.admin
    ? `Duplicato rilevato.${details}`
    : `Ore già inserite per questo operatore.${details}`;
}


function findDuplicateHoursInDraft(entries) {
  const map = new Map();
  const duplicates = [];
  entries.forEach((entry) => {
    const commessaInfo = resolveHoursEntryCommessa(entry);
    const commessaId = String(commessaInfo.id || commessaInfo.key || "").trim();
    if (!commessaId) return;
    (entry.rows || []).forEach((row) => {
      const operatore = String(row?.operatore || "").trim();
      const normalizedOperator = normalizeHoursOperatorName(operatore);
      const operatorKey = getHoursOperatorIdentity(row);
      const ore = Number(row?.ore || 0);
      if (!normalizedOperator || !operatorKey || ore <= 0) return;
      const key = `${commessaId}__${operatorKey}`;
      if (map.has(key)) {
        duplicates.push({
          commessaId,
          commessaName: commessaInfo.nome || entry.commessaName || commesseById.get(commessaId)?.nome || "Commessa",
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
    const commessaInfo = resolveHoursEntryCommessa(entry);
    const commessaId = String(commessaInfo.id || commessaInfo.key || "").trim();
    if (!commessaId) return;
    (entry.rows || []).forEach((row) => {
      const operatore = String(row?.operatore || "").trim();
      const normalizedOperator = normalizeHoursOperatorName(operatore);
      const operatorKey = getHoursOperatorIdentity(row);
      const ore = Number(row?.ore || 0);
      if (!normalizedOperator || !operatorKey || ore <= 0) return;
      const requestedValue = {
        commessaId,
        commessaName: entry.commessaName || commesseById.get(commessaId)?.nome || "Commessa",
        operatore
      };
      requestedKeys.set(`${commessaId}__${operatorKey}`, requestedValue);
      requestedKeys.set(`${commessaId}__${normalizedOperator}`, requestedValue);
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
      const commessaInfo = resolveHoursEntryCommessa(entry);
      const commessaId = String(commessaInfo.id || commessaInfo.key || "").trim();
      if (!commessaId) return;
      (entry.rows || []).forEach((row) => {
        const operatore = String(row?.operatore || "").trim();
        const normalizedOperator = normalizeHoursOperatorName(operatore);
        const operatorKey = getHoursOperatorIdentity(row);
        const ore = Number(row?.ore || 0);
        if (!normalizedOperator || !operatorKey || ore <= 0) return;
        const match = requestedKeys.get(`${commessaId}__${operatorKey}`) || requestedKeys.get(`${commessaId}__${normalizedOperator}`);
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

function getHoursEntryTotal(entry) {
  return (entry?.rows || []).reduce((sum, row) => sum + Number(row?.ore || 0), 0);
}

function buildHoursInsertedChatText(payload) {
  const author = payload?.createdByName || payload?.createdByEmail || "Operatore";
  const dateLabel = payload?.date
    ? new Date(`${payload.date}T00:00:00`).toLocaleDateString("it-IT")
    : "-";
  const details = (Array.isArray(payload?.entries) ? payload.entries : [])
    .map((entry) => {
      const commessaName = String(entry?.commessaName || "Commessa").trim() || "Commessa";
      const totalHours = getHoursEntryTotal(entry);
      return `${formatHoursNumber(totalHours)} ore in ${commessaName}`;
    })
    .filter(Boolean)
    .join("; ");
  return `🕒 ${author} ha inserito le ore del ${dateLabel}: ${details || "nessun dettaglio ore"}.`;
}

async function notifyHoursInsertedToChat(requestId, payload) {
  const text = buildHoursInsertedChatText(payload);
  const firstEntry = Array.isArray(payload?.entries) && payload.entries.length ? payload.entries[0] : null;
  await sendChatMessage({
    type: "text",
    text,
    recipientId: "",
    kind: "system",
    metadata: {
      type: "hours_inserted",
      approvalRequestId: requestId || "",
      date: payload?.date || "",
      entries: (Array.isArray(payload?.entries) ? payload.entries : []).map((entry) => ({
        commessaId: entry?.commessaId || "",
        commessaName: entry?.commessaName || "",
        totalHours: getHoursEntryTotal(entry)
      }))
    }
  });
  await publishGlobalNotificationEvent("hours-inserted", {
    title: "Ore inserite",
    body: text,
    commessaId: firstEntry?.commessaId || "",
    commessaName: firstEntry?.commessaName || ""
  });
}

function setHoursFinalizeButtonText(state = "idle") {
  const button = ui.hoursFinalizeBtn;
  if (!button) return;
  if (hoursFinalizeStatusTimer) {
    clearTimeout(hoursFinalizeStatusTimer);
    hoursFinalizeStatusTimer = null;
  }
  if (state === "loading") {
    button.textContent = "⏳ Salvataggio…";
    button.setAttribute("aria-busy", "true");
    return;
  }
  button.removeAttribute("aria-busy");
  if (state === "saved") {
    button.textContent = "✅ Salvato";
    hoursFinalizeStatusTimer = setTimeout(() => {
      if (!hoursSubmitInFlight && !hoursFinalizeLocked) setHoursFinalizeButtonText("idle");
    }, 1800);
    return;
  }
  button.textContent = "✓ Fine: salva";
}

function setHoursFinalizeLocked(locked) {
  hoursFinalizeLocked = Boolean(locked);
  if (ui.hoursFinalizeBtn) {
    ui.hoursFinalizeBtn.disabled = hoursFinalizeLocked;
  }
}

function unlockHoursFinalizeButton() {
  hoursSubmitInFlight = false;
  setHoursFinalizeButtonText("idle");
  if (!hoursFinalizeLocked) return;
  setHoursFinalizeLocked(false);
  if (ui.hoursFeedback?.textContent && ui.hoursFeedback.textContent.includes("Richiesta inviata")) {
    ui.hoursFeedback.textContent = "Modifiche rilevate. Puoi premere di nuovo “Fine: salva”.";
  }
}

async function finalizeHoursReport(event) {
  event.preventDefault();
  if (hoursFinalizeLocked || hoursSubmitInFlight) {
    ui.hoursFeedback.textContent = hoursSubmitInFlight
      ? "Invio già in corso: attendi il completamento prima di premere di nuovo."
      : "Richiesta già inviata. Modifica commesse/operatori/ore per riattivare “Fine: salva”.";
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
      rows: entry.rows
        .filter((row) => row.operatore && row.ore > 0)
        .map((row) => ({
          operatore: row.operatore,
          operatoreId: row.operatoreId || resolveHoursOperatorId(row.operatore),
          ore: row.ore
        }))
    })).filter((entry) => entry.commessaId && entry.rows.length),
    createdByUid: currentUser.uid,
    createdByName: currentUser.displayName || currentUser.email || "Operatore",
    createdByEmail: currentUser.email || "",
    createdAt: firebase.firestore.FieldValue.serverTimestamp()
  };
  if (!canManageData()) {
    const unauthorizedEntry = payload.entries.find((entry) => !canCurrentUserInsertHoursForCommessa(entry.commessaId, dateValue));
    if (unauthorizedEntry) {
      ui.hoursFeedback.textContent = `Permesso negato: puoi inserire ore solo per le commesse dove sei assegnato in squadra (${unauthorizedEntry.commessaName || "Commessa"}).`;
      return;
    }
  }
  hoursSubmitInFlight = true;
  ui.hoursFinalizeBtn.disabled = true;
  setHoursFinalizeButtonText("loading");
  ui.hoursFeedback.textContent = "Invio richiesta approvazione in corso...";
  try {
    const duplicateDraft = findDuplicateHoursInDraft(payload.entries);
    if (duplicateDraft.length) {
      ui.hoursFeedback.textContent = formatHoursDuplicateMessage(duplicateDraft, { admin: false });
      setHoursFinalizeButtonText("idle");
      return;
    }

    const existingConflicts = await findExistingHoursConflicts(dateValue, payload.entries);
    if (existingConflicts.length) {
      ui.hoursFeedback.textContent = formatHoursDuplicateMessage(existingConflicts, { admin: false });
      setHoursFinalizeButtonText("idle");
      return;
    }

    const approvalRef = await reserveHoursApprovalRequestWithLocks(payload);
    await notifyHoursInsertedToChat(approvalRef.id, payload);
    await notifyLevel1ForHoursApproval(approvalRef.id, payload);

    setHoursFinalizeButtonText("saved");
    ui.hoursFeedback.textContent = `Richiesta inviata (ID ${approvalRef.id}). In attesa primo OK, poi conferma admin finale.`;
    ui.hoursCommesseList.innerHTML = "";
    addHoursCommessaBlock();
    setHoursFinalizeLocked(true);
    renderHoursSummary();
    loadSavedHoursReports();
  } catch (error) {
    console.error("Salvataggio gestione ore non riuscito:", error);
    setHoursFinalizeButtonText("idle");
    if (error?.code === "hours-duplicate-lock") {
      ui.hoursFeedback.textContent = formatHoursDuplicateMessage(error.conflicts, { admin: false });
    } else {
      ui.hoursFeedback.textContent = "Errore durante il salvataggio del resoconto ore.";
    }
  } finally {
    hoursSubmitInFlight = false;
    if (!hoursFinalizeLocked) {
      ui.hoursFinalizeBtn.disabled = false;
      setHoursFinalizeButtonText("idle");
    }
  }
}

async function sendHoursApprovalChatNotification({ recipients = [], text, senderName, metadata }) {
  const normalizedRecipients = (Array.isArray(recipients) ? recipients : [])
    .map((recipient) => String(recipient?.id || recipient?.uid || "").trim())
    .filter(Boolean);
  if (normalizedRecipients.length) {
    await Promise.all(normalizedRecipients.map((recipientId) => sendPrivateChatNotification({
      recipientId,
      text,
      senderName,
      metadata
    })));
    return;
  }
  await sendChatMessage({
    type: "text",
    text,
    recipientId: "",
    kind: "system",
    metadata
  });
}

function getKnownAdminChatRecipients() {
  return platformUsers.filter((user) => adminEmails.has(normalizeEmail(user.email)));
}

async function notifyLevel1ForHoursApproval(requestId, payload) {
  const adminUsers = getKnownAdminChatRecipients();
  const recipientsMap = new Map(adminUsers.map((user) => [user.id, user]));
  const requesterUser = platformUsers.find((user) => String(user.uid || "") === String(payload?.createdByUid || ""));
  if (adminUsers.length && requesterUser?.id) recipientsMap.set(requesterUser.id, requesterUser);
  const dateLabel = payload?.date
    ? new Date(`${payload.date}T00:00:00`).toLocaleDateString("it-IT")
    : "-";
  const author = payload?.createdByName || payload?.createdByEmail || "Operatore";
  const entriesCount = Array.isArray(payload?.entries) ? payload.entries.length : 0;
  const text = `🕒 Richiesta ore ${requestId} da ${author} (${dateLabel}). Commesse: ${entriesCount}. Serve primo OK.`;
  await sendHoursApprovalChatNotification({
    recipients: Array.from(recipientsMap.values()),
    text,
    senderName: currentUser.displayName || currentUser.email || "Operatore",
    metadata: {
      type: "hours_approval",
      approvalRequestId: requestId,
      action: "level1_ok"
    }
  });
}

async function notifyAdminsForFinalHoursApproval(requestId, payload, approverName) {
  const adminUsers = getKnownAdminChatRecipients();
  const author = payload?.createdByName || payload?.createdByEmail || "Operatore";
  const text = `🕒 Richiesta ore ${requestId}: primo OK da ${approverName || "utente autorizzato"}. Attesa approvazione admin finale per ${author}.`;
  await sendHoursApprovalChatNotification({
    recipients: adminUsers,
    text,
    senderName: currentUser.displayName || currentUser.email || "Sistema",
    metadata: {
      type: "hours_approval",
      approvalRequestId: requestId,
      action: "admin_final_ok"
    }
  });
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
    const opened = safeOpenWhatsAppMessage(`${shareMessage} Ho generato il PDF: ${lastSegnalazionePdfName}`);
    ui.segnalazioneFeedback.textContent = opened
      ? "WhatsApp aperto. Allega il PDF scaricato prima di inviare."
      : "Impossibile aprire WhatsApp automaticamente. Condividi il PDF manualmente.";
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
  return canManageData() ? (localStorage.getItem("googleDriveAccessToken") || "") : "";
}

function updateDriveConnectVisibility() {
  const showConnect = Boolean(currentUser && canManageData());
  if (!ui.driveConnectBtn) return;
  ui.driveConnectBtn.classList.toggle("hidden", !showConnect);
  ui.driveConnectBtn.disabled = !showConnect;
}

function isCentralDriveConfigured() {
  return Boolean(driveBridgeState.configured || driveAccessToken);
}

function getCentralDriveNotConfiguredMessage() {
  return "Cloud amministratore non configurato";
}

function normalizeDriveFolderName(value, fallback = "Generale") {
  return String(value || fallback).trim().replace(/[\\/:*?"<>|]+/g, "-").slice(0, 120) || fallback;
}

function escapeDriveQueryValue(value) {
  return String(value || "").replace(/\\/g, "\\\\").replace(/'/g, "\\'");
}

function inferCentralDriveType(folderId = "", options = {}) {
  const explicit = String(options.driveType || options.type || "").trim().toUpperCase();
  if (explicit) return normalizeDriveFolderName(explicit, "EXPORT").toUpperCase();
  if (folderId && folderId === driveChatFolderId) return "FOTO";
  if (folderId && folderId === driveReportsFolderId) return "SEGNALAZIONI";
  if (folderId && folderId === driveSquadreFolderId) return "EXPORT";
  if (folderId && folderId === driveHelpCenterFolderId) return "EXPORT";
  return "EXPORT";
}

function getCurrentDriveCommessaName(options = {}) {
  return normalizeDriveFolderName(options.commessaName || selectedCommessaName || "", CENTRAL_DRIVE_DEFAULT_COMMESSA);
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
  if (forceAccountSelection) provider.setCustomParameters({ prompt: "select_account" });

  // Flusso dedicato Android (APK/WebView/Capacitor): NO redirect fallback.
  if (isAndroidWebViewRuntime()) {
    auth.signInWithPopup(provider).then((result) => {
      void result;
    }).catch((error) => {
      console.error("Login Google Android/WebView fallito:", error);
      alert("Errore login Android/WebView: " + error.message);
    });
    return;
  }

  // Flusso web desktop/browser standard: manteniamo comportamento esistente.
  auth.signInWithPopup(provider).then((result) => {
    void result;
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
  if (!driveAccessToken || !canManageData()) return;
  const { notifyOnError = false } = options;
  try {
    await ensureDriveFolders();
    const driveUser = auth.currentUser || currentUser;
    const ownerEmail = (driveUser && driveUser.email) ? driveUser.email : ADMIN_EMAIL;
    await db.collection("appConfig").doc("driveAdminSecret").set({
      ownerEmail,
      accessToken: driveAccessToken,
      refreshToken: "",
      rootFolderId: CENTRAL_DRIVE_ROOT_FOLDER_ID,
      chatFolderId: driveChatFolderId,
      reportsFolderId: driveReportsFolderId,
      squadreFolderId: driveSquadreFolderId,
      helpCenterFolderId: driveHelpCenterFolderId,
      updatedAt: firebase.firestore.FieldValue.serverTimestamp()
    }, { merge: true });
    await db.collection("appConfig").doc("driveBridge").set({
      ownerEmail,
      configured: true,
      rootFolderId: CENTRAL_DRIVE_ROOT_FOLDER_ID,
      rootFolderName: CENTRAL_DRIVE_ROOT_FOLDER_NAME,
      updatedAt: firebase.firestore.FieldValue.serverTimestamp()
    }, { merge: true });
    await migrateLegacyDriveDataToCentralRoot({ force: true });
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


function isSubcommessa(commessa = {}) {
  return Boolean(String(commessa.parentCommessaId || "").trim());
}

function getMainCommesse() {
  return Array.from(commesseById.values()).filter((commessa) => !isSubcommessa(commessa));
}

function getSubcommesse(parentCommessaId) {
  return Array.from(commesseById.values()).filter((commessa) => String(commessa.parentCommessaId || "") === String(parentCommessaId || ""));
}

function validateFirebaseConfigForCommesse() {
  const requiredKeys = ["apiKey", "projectId", "appId"];
  const missingKeys = requiredKeys.filter((key) => !String(firebaseConfig?.[key] || "").trim());
  if (missingKeys.length) {
    console.error("Configurazione Firebase incompleta per il caricamento commesse. Variabili equivalenti richieste: VITE_FIREBASE_API_KEY, VITE_FIREBASE_PROJECT_ID, VITE_FIREBASE_APP_ID. Campi mancanti:", missingKeys);
    return false;
  }
  if (!db || typeof db.collection !== "function") {
    console.error("Firestore non inizializzato correttamente: db.collection non disponibile.");
    return false;
  }
  return true;
}

function getCommesseErrorMessage() {
  return "Impossibile caricare le commesse online. Mostro dati salvati localmente.";
}

function parseCachedCommesse() {
  try {
    const parsed = JSON.parse(localStorage.getItem(COMMESSE_LOCAL_CACHE_KEY) || "[]");
    return Array.isArray(parsed) ? parsed.filter((item) => item && item.id) : [];
  } catch (error) {
    console.error("Errore lettura cache locale commesse:", error);
    return [];
  }
}

function saveCommesseLocalCache(commesse) {
  if (!Array.isArray(commesse)) return;
  try {
    localStorage.setItem(COMMESSE_LOCAL_CACHE_KEY, JSON.stringify(commesse));
  } catch (error) {
    console.error("Errore salvataggio cache locale commesse:", error);
  }
}

function loadCommesseFromLocalCache() {
  const cachedCommesse = parseCachedCommesse();
  commesseById = new Map(cachedCommesse.map((commessa) => [commessa.id, commessa]));
  if (cachedCommesse.length) {
    console.log("Commesse ricevute dalla cache locale:", cachedCommesse);
  }
  return cachedCommesse;
}

function refreshCommesseDependentUI(includeRemoteStats = true) {
  if (includeRemoteStats) {
    subscribeStatsForCommesse();
    subscribeHoursStats();
  }
  renderCommesseHomeList();
  renderCommessaSelects();
  renderOrganizeCommesseControls();
  renderCommesseManagementList();
  renderHoursCommessaSelectOptions();
  renderSquadre();
  renderResourcesList();
  renderResourceButtonsForCommessa();
  syncBannerFormFromSelection();
  updateCommessaContextUI();
  renderParentCommessaOverview();
  renderNextActionCard();
}


function getEmptyCommessaStats() {
  return {
    total: 0,
    done: 0,
    openAlerts: 0,
    firstDoneAtMs: 0,
    firstDoneDateKey: "",
    lastResetAtMs: 0,
    lastResetDateKey: "",
    totaleMqPrevisti: 0,
    totaleMqEseguiti: 0,
    mqRimanenti: 0,
    avanzamentoMq: 0
  };
}

function dateKeyFromMillis(millis) {
  if (!Number.isFinite(millis) || millis <= 0) return "";
  const date = new Date(millis);
  const year = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, "0");
  const day = String(date.getDate()).padStart(2, "0");
  return `${year}-${month}-${day}`;
}

function calculateImpiantiStats(rawImpianti = []) {
  const combined = combineImpiantiForView(rawImpianti);
  const total = combined.length;
  const doneImpianti = combined.filter((impianto) => Boolean(impianto.done));
  const done = doneImpianti.length;
  const openAlerts = combined.filter((impianto) => hasOpenImpiantoAlert(impianto)).length;
  const mqStats = calculateImpiantiMqProgress(combined);
  const firstDoneAtMs = doneImpianti.reduce((earliest, impianto) => {
    const doneAtMs = firestoreDateToMillis(impianto.doneAt);
    if (!doneAtMs) return earliest;
    return earliest ? Math.min(earliest, doneAtMs) : doneAtMs;
  }, 0);
  const lastResetAtMs = combined.reduce((latest, impianto) => {
    const resetAtMs = firestoreDateToMillis(impianto.resetAt);
    return resetAtMs > latest ? resetAtMs : latest;
  }, 0);
  return {
    total,
    done,
    openAlerts,
    firstDoneAtMs,
    firstDoneDateKey: dateKeyFromMillis(firstDoneAtMs),
    lastResetAtMs,
    lastResetDateKey: dateKeyFromMillis(lastResetAtMs),
    ...mqStats
  };
}

function hasOpenImpiantoAlert(impianto = {}) {
  const fields = [impianto.segnalazioneStatus, impianto.reportStatus, impianto.alertStatus, impianto.problemStatus, impianto.status];
  const text = fields.map((value) => String(value || "").toLowerCase()).join(" ");
  return Boolean(
    impianto.hasSegnalazione ||
    impianto.segnalazioneAperta ||
    impianto.problemOpen ||
    impianto.alertOpen ||
    text.includes("apert") ||
    text.includes("pending")
  ) && !text.includes("chius") && !text.includes("risolt") && !text.includes("done") && !text.includes("fatto");
}

function getCommessaStats(commessaId) {
  return commessaStatsById.get(commessaId) || getEmptyCommessaStats();
}

function getEmptyCommessaWorkSummary() {
  return { totalHours: 0, workedDays: 0, averageHoursPerDay: 0, firstDoneAtMs: 0, firstDoneDateKey: "", workedDateKeys: new Set() };
}

function getCommessaWorkSummary(commessaId) {
  return commessaWorkSummariesById.get(commessaId) || getEmptyCommessaWorkSummary();
}

function getCommessaHoursTotal(commessaId) {
  return Number(getCommessaWorkSummary(commessaId).totalHours || 0);
}

function getParentCommessaAggregate(parentCommessaId) {
  const subcommesse = getSubcommesse(parentCommessaId);
  return subcommesse.reduce((acc, sub) => {
    const stats = getCommessaStats(sub.id);
    const workSummary = getCommessaWorkSummary(sub.id);
    acc.subCount += 1;
    acc.total += stats.total;
    acc.done += stats.done;
    acc.openAlerts += stats.openAlerts;
    acc.totaleMqPrevisti += Number(stats.totaleMqPrevisti || 0);
    acc.totaleMqEseguiti += Number(stats.totaleMqEseguiti || 0);
    acc.mqRimanenti += Number(stats.mqRimanenti || 0);
    acc.hours += Number(workSummary.totalHours || 0);
    (workSummary.workedDateKeys || new Set()).forEach((dateKey) => acc.workedDateKeys.add(dateKey));
    const firstDoneAtMs = Number(workSummary.firstDoneAtMs || stats.firstDoneAtMs || 0);
    if (firstDoneAtMs) acc.firstDoneAtMs = acc.firstDoneAtMs ? Math.min(acc.firstDoneAtMs, firstDoneAtMs) : firstDoneAtMs;
    return acc;
  }, { subCount: 0, total: 0, done: 0, openAlerts: 0, totaleMqPrevisti: 0, totaleMqEseguiti: 0, mqRimanenti: 0, hours: 0, workedDateKeys: new Set(), firstDoneAtMs: 0 });
}

function formatProgress(done, total) {
  if (!total) return "0%";
  return `${Math.round((Number(done || 0) / total) * 100)}%`;
}

function formatMqNumber(value) {
  return formatAreaMqValue(value);
}

function formatMqProgressDetails(stats = {}) {
  return `MQ eseguiti: ${formatMqNumber(stats.totaleMqEseguiti || 0)} • MQ rimanenti: ${formatMqNumber(stats.mqRimanenti || 0)} • Totale MQ: ${formatMqNumber(stats.totaleMqPrevisti || 0)}`;
}

function formatHoursNumber(value) {
  return Number(value || 0).toLocaleString("it-IT", { maximumFractionDigits: 2 });
}

function formatWorkSummaryParts(workSummary) {
  const totalHours = Number(workSummary.totalHours || workSummary.hours || 0);
  const workedDays = Number(workSummary.workedDays ?? workSummary.workedDateKeys?.size ?? 0);
  const averageHoursPerDay = workedDays > 0 ? totalHours / workedDays : 0;
  return `Ore ${formatHoursNumber(totalHours)} • Giorni lavorati ${workedDays} • Media ore/giorno ${formatHoursNumber(averageHoursPerDay)}`;
}

function formatParentCommessaSummary(aggregate) {
  const workedDays = aggregate.workedDateKeys?.size || 0;
  return `${aggregate.subCount} subcommesse • ${aggregate.total} impianti • ${aggregate.openAlerts} segnalazioni aperte • Avanzamento ${formatProgress(aggregate.done, aggregate.total)} • ${formatWorkSummaryParts({ totalHours: aggregate.hours, workedDays })}`;
}

function formatSingleCommessaSummary(commessaId) {
  const stats = getCommessaStats(commessaId);
  const workSummary = getCommessaWorkSummary(commessaId);
  return `${stats.total} impianti • ${stats.openAlerts} segnalazioni aperte • Avanzamento ${stats.avanzamentoMq || 0}% • ${formatMqProgressDetails(stats)} • ${formatWorkSummaryParts(workSummary)}`;
}

function getSelectedCommessaDashboardStats() {
  const selectedSubcommesse = selectedCommessaId ? getSubcommesse(selectedCommessaId) : [];
  if (selectedSubcommesse.length) {
    const aggregate = getParentCommessaAggregate(selectedCommessaId);
    const activeDateKey = getActiveSquadreDateKey();
    const squadreCount = selectedSubcommesse.reduce((sum, subcommessa) => sum + getSquadreCountForCommessaDate(subcommessa.id, activeDateKey), 0);
    const avanzamentoMq = aggregate.totaleMqPrevisti > 0
      ? Math.round((aggregate.totaleMqEseguiti / aggregate.totaleMqPrevisti) * 100)
      : Number.parseInt(formatProgress(aggregate.done, aggregate.total), 10) || 0;
    return {
      total: aggregate.total,
      done: aggregate.done,
      segnalazioni: aggregate.openAlerts,
      avanzamento: `${avanzamentoMq}%`,
      totaleMqPrevisti: aggregate.totaleMqPrevisti,
      totaleMqEseguiti: aggregate.totaleMqEseguiti,
      mqRimanenti: aggregate.mqRimanenti,
      ore: Number(aggregate.hours || 0),
      giorni: aggregate.workedDateKeys?.size || 0,
      squadreCount
    };
  }

  const liveTotal = currentImpianti.length;
  const liveDone = currentImpianti.filter((impianto) => Boolean(impianto.done)).length;
  const storedStats = getCommessaStats(selectedCommessaId);
  const total = liveTotal || Number(storedStats.total || 0);
  const done = liveTotal ? liveDone : Number(storedStats.done || 0);
  const linkedNoteCount = currentCommessaNotes.filter((note) => String(note?.impiantoKey || note?.impiantoId || "").trim()).length;
  const segnalazioni = linkedNoteCount || Number(storedStats.openAlerts || 0) || currentCommessaNotes.length;
  const workSummary = getCommessaWorkSummary(selectedCommessaId);
  const liveMqStats = liveTotal ? calculateImpiantiMqProgress(currentImpianti) : null;
  const mqStats = liveMqStats || {
    totaleMqPrevisti: Number(storedStats.totaleMqPrevisti || 0),
    totaleMqEseguiti: Number(storedStats.totaleMqEseguiti || 0),
    mqRimanenti: Number(storedStats.mqRimanenti || 0),
    avanzamentoMq: Number(storedStats.avanzamentoMq || 0)
  };
  const activeDateKey = getActiveSquadreDateKey();
  const squadreCount = getSquadreCountForCommessaDate(selectedCommessaId, activeDateKey);
  return {
    total,
    done,
    segnalazioni,
    avanzamento: `${mqStats.avanzamentoMq || 0}%`,
    ...mqStats,
    ore: Number(workSummary.totalHours || 0),
    giorni: Number(workSummary.workedDays ?? workSummary.workedDateKeys?.size ?? 0),
    squadreCount
  };
}

function updateCommessaDashboard() {
  if (!selectedCommessaId) return;
  const stats = getSelectedCommessaDashboardStats();
  if (ui.commessaStatImpianti) ui.commessaStatImpianti.textContent = String(stats.total || 0);
  if (ui.commessaStatSegnalazioni) ui.commessaStatSegnalazioni.textContent = String(stats.segnalazioni || 0);
  if (ui.commessaStatAvanzamento) {
    ui.commessaStatAvanzamento.textContent = stats.avanzamento;
    const progressDetails = formatMqProgressDetails(stats);
    if (ui.commessaStatAvanzamentoDetail) {
      const mqDetailRows = [
        ["MQ eseguiti", stats.totaleMqEseguiti],
        ["MQ rimanenti", stats.mqRimanenti],
        ["Totale MQ", stats.totaleMqPrevisti]
      ];
      ui.commessaStatAvanzamentoDetail.innerHTML = mqDetailRows
        .map(([label, value]) => `<span>${label}: ${escapeHTML(formatMqNumber(value || 0))}</span>`)
        .join("");
    }
    const progressItem = ui.commessaStatAvanzamento.closest?.(".commessa-stat-item");
    if (progressItem) {
      progressItem.title = progressDetails;
      progressItem.setAttribute("aria-label", `Vai agli impianti fatti. ${progressDetails}`);
    }
  }
  if (ui.commessaStatOre) ui.commessaStatOre.textContent = `${formatHoursNumber(stats.ore)} h`;
  if (ui.commessaStatGiorni) ui.commessaStatGiorni.textContent = `${stats.giorni || 0} gg`;
  if (ui.commessaActiveSquadreCount) {
    ui.commessaActiveSquadreCount.textContent = `${stats.squadreCount || 0} squadr${stats.squadreCount === 1 ? "a attiva" : "e attive"}`;
  }
  updateCommessaWeatherRefreshButtonState();
  if (ui.commessaCallBtn) {
    const hasPhoneResources = getResourcesByCommessa(selectedCommessaId, "phone").length > 0;
    ui.commessaCallBtn.classList.toggle("commessa-action-btn--disabled", !hasPhoneResources);
    ui.commessaCallBtn.setAttribute("aria-disabled", String(!hasPhoneResources));
  }
}

function scrollToImpiantiCard() {
  ui.impiantiCard?.scrollIntoView({ behavior: "smooth", block: "start" });
}

function handleCommessaStatAction(action) {
  if (!selectedCommessaId) return;
  if (action === "impianti") {
    setImpiantiViewMode("todo");
    scrollToImpiantiCard();
    return;
  }
  if (action === "segnalazioni") {
    setImpiantiViewMode("alerts");
    scrollToImpiantiCard();
    return;
  }
  if (action === "avanzamento") {
    setImpiantiViewMode("done");
    scrollToImpiantiCard();
    return;
  }
  if (action === "ore" || action === "giorni") {
    openHoursPage();
  }
}


function toggleOrganizeCommesseScreen(show) {
  ui.organizeCommesseScreen?.classList.toggle("hidden", !show);
  ui.organizeCommesseScreen?.setAttribute("aria-hidden", show ? "false" : "true");
  if (show) {
    renderOrganizeCommesseControls();
    if (ui.organizeCommesseFeedback) ui.organizeCommesseFeedback.textContent = "";
  }
}

function renderOrganizeCommesseControls() {
  renderMoveParentCommessaSelect();
  renderMoveSubcommesseList();
}

function renderMoveParentCommessaSelect() {
  if (!ui.moveParentCommessaSelect) return;
  const previousValue = ui.moveParentCommessaSelect.value;
  ui.moveParentCommessaSelect.innerHTML = "<option value=''>Seleziona commessa padre</option>";
  sortCommesseByCreatedAtDesc(getMainCommesse()).forEach((commessa) => {
    ui.moveParentCommessaSelect.appendChild(createCommessaOption(commessa));
  });
  if (previousValue && commesseById.has(previousValue) && !isSubcommessa(commesseById.get(previousValue))) {
    ui.moveParentCommessaSelect.value = previousValue;
  }
}

function renderMoveSubcommesseList() {
  if (!ui.moveSubcommesseList) return;
  const parentId = String(ui.moveParentCommessaSelect?.value || "").trim();
  const candidates = sortCommesseByCreatedAtDesc(Array.from(commesseById.values()))
    .filter((commessa) => commessa.id !== parentId && String(commessa.parentCommessaId || "") !== parentId && getSubcommesse(commessa.id).length === 0);
  if (!candidates.length) {
    ui.moveSubcommesseList.innerHTML = "<p class='muted'>Nessuna commessa spostabile.</p>";
    return;
  }
  ui.moveSubcommesseList.innerHTML = "";
  candidates.forEach((commessa) => {
    const row = document.createElement("label");
    row.className = "organize-commessa-choice";
    const parent = commessa.parentCommessaId ? commesseById.get(commessa.parentCommessaId) : null;
    row.innerHTML = `
      <input type="checkbox" value="${escapeHTML(commessa.id)}">
      <span><strong>${escapeHTML(commessa.nome || "Commessa senza nome")}</strong>${commessa.codice ? ` • Cod. ${escapeHTML(commessa.codice)}` : ""}${parent ? `<br><small>Ora sotto: ${escapeHTML(parent.nome || "commessa padre")}</small>` : ""}</span>
    `;
    ui.moveSubcommesseList.appendChild(row);
  });
}

async function createParentCommessa(event) {
  event.preventDefault();
  if (!currentUser) {
    alert("Devi fare login.");
    return;
  }
  if (!canManageData()) {
    alert("Solo un admin può creare commesse padre.");
    return;
  }
  const nome = String(ui.parentCommessaName?.value || "").trim();
  const codice = String(ui.parentCommessaCode?.value || "").trim();
  if (!nome) return;
  await db.collection("commesse").add({
    nome,
    codice,
    parentCommessaId: null,
    creatoDa: currentUser.email || "",
    createdAt: firebase.firestore.FieldValue.serverTimestamp()
  });
  ui.parentCommessaForm?.reset();
  if (ui.organizeCommesseFeedback) ui.organizeCommesseFeedback.textContent = `Commessa padre "${nome}" creata.`;
}

async function moveSelectedCommesseUnderParent(event) {
  event.preventDefault();
  if (!canManageData()) {
    alert("Solo un admin può organizzare commesse.");
    return;
  }
  const parentId = String(ui.moveParentCommessaSelect?.value || "").trim();
  if (!parentId) {
    alert("Seleziona una commessa padre.");
    return;
  }
  const selectedIds = Array.from(ui.moveSubcommesseList?.querySelectorAll("input[type='checkbox']:checked") || [])
    .map((input) => input.value)
    .filter((id) => id && id !== parentId);
  if (!selectedIds.length) {
    alert("Seleziona almeno una commessa da spostare.");
    return;
  }
  const batch = db.batch();
  selectedIds.forEach((id) => {
    batch.set(db.collection("commesse").doc(id), {
      parentCommessaId: parentId,
      updatedAt: firebase.firestore.FieldValue.serverTimestamp(),
      updatedBy: currentUser?.email || ""
    }, { merge: true });
  });
  await batch.commit();
  if (ui.organizeCommesseFeedback) ui.organizeCommesseFeedback.textContent = `${selectedIds.length} commesse spostate sotto la commessa padre selezionata.`;
}

async function moveSubcommessaToMain(commessaId) {
  if (!canManageData()) {
    alert("Solo un admin può spostare subcommesse nella vista principale.");
    return;
  }
  await db.collection("commesse").doc(commessaId).set({
    parentCommessaId: null,
    updatedAt: firebase.firestore.FieldValue.serverTimestamp(),
    updatedBy: currentUser?.email || ""
  }, { merge: true });
}


function normalizeHoursReportDateKey(value) {
  const text = String(value || "").trim();
  if (/^\d{4}-\d{2}-\d{2}$/.test(text)) return text;
  const millis = firestoreDateToMillis(value);
  return dateKeyFromMillis(millis);
}

function normalizeCommessaNameForRules(value) {
  return String(value || "")
    .trim()
    .toUpperCase()
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/\s+/g, " ");
}

function isHeraDiscaricheCommessa(commessa = {}) {
  return normalizeCommessaNameForRules(commessa.nome).includes("HERA DISCARICHE");
}

function isUnderHeraDiscaricheParent(commessaId) {
  let parentId = String((commesseById.get(commessaId) || {}).parentCommessaId || "").trim();
  const visited = new Set();
  while (parentId && !visited.has(parentId)) {
    visited.add(parentId);
    const parent = commesseById.get(parentId);
    if (!parent) return false;
    if (isHeraDiscaricheCommessa(parent)) return true;
    parentId = String(parent.parentCommessaId || "").trim();
  }
  return false;
}

function sumPositiveHoursRows(rows = []) {
  return (Array.isArray(rows) ? rows : []).reduce((sum, row) => {
    const hours = Number(row.ore || 0);
    return Number.isFinite(hours) && hours > 0 ? sum + hours : sum;
  }, 0);
}

function getFirstWorkedDateKeyForCommessa(commessaId, options = {}) {
  const minExclusiveDateKey = String(options.minExclusiveDateKey || "");
  const maxDateKey = String(options.maxDateKey || "");
  let firstWorkedDateKey = "";
  allHoursReports.forEach((report) => {
    const reportDateKey = normalizeHoursReportDateKey(report.date);
    if (!reportDateKey || (minExclusiveDateKey && reportDateKey <= minExclusiveDateKey) || (maxDateKey && reportDateKey > maxDateKey)) return;
    const hasCommessaHours = (Array.isArray(report.entries) ? report.entries : []).some((entry) => (
      String(entry.commessaId || "").trim() === String(commessaId) && sumPositiveHoursRows(entry.rows) > 0
    ));
    if (!hasCommessaHours) return;
    if (!firstWorkedDateKey || reportDateKey < firstWorkedDateKey) firstWorkedDateKey = reportDateKey;
  });
  return firstWorkedDateKey;
}

function getCommessaWorkRange(commessaId, stats) {
  if (isUnderHeraDiscaricheParent(commessaId)) {
    const resetDateKey = String(stats.lastResetDateKey || "");
    const doneDateKey = String(stats.firstDoneDateKey || "");
    const effectiveDoneDateKey = doneDateKey && (!resetDateKey || doneDateKey > resetDateKey) ? doneDateKey : "";
    return {
      startDateKey: getFirstWorkedDateKeyForCommessa(commessaId, {
        minExclusiveDateKey: resetDateKey,
        maxDateKey: effectiveDoneDateKey
      }),
      endDateKey: effectiveDoneDateKey,
      resetDateKey,
      startMode: "hera_discariche"
    };
  }
  const firstDoneAtMs = Number(stats.firstDoneAtMs || 0);
  return {
    startDateKey: Number(stats.done || 0) > 0 ? String(stats.firstDoneDateKey || dateKeyFromMillis(firstDoneAtMs) || "") : "",
    endDateKey: "",
    resetDateKey: String(stats.lastResetDateKey || ""),
    startMode: "done"
  };
}

function recalculateCommessaWorkSummaries() {
  const summaries = new Map();
  commesseById.forEach((_commessa, commessaId) => {
    const stats = getCommessaStats(commessaId);
    const workRange = getCommessaWorkRange(commessaId, stats);
    const startDateKey = String(workRange.startDateKey || "");
    const endDateKey = String(workRange.endDateKey || "");
    if (!startDateKey) {
      summaries.set(commessaId, getEmptyCommessaWorkSummary());
      return;
    }

    let totalHours = 0;
    const workedDateKeys = new Set();
    allHoursReports.forEach((report) => {
      const reportDateKey = normalizeHoursReportDateKey(report.date);
      if (!reportDateKey || reportDateKey < startDateKey || (endDateKey && reportDateKey > endDateKey)) return;
      (Array.isArray(report.entries) ? report.entries : []).forEach((entry) => {
        if (String(entry.commessaId || "").trim() !== String(commessaId)) return;
        const entryHours = sumPositiveHoursRows(entry.rows);
        if (entryHours <= 0) return;
        totalHours += entryHours;
        workedDateKeys.add(reportDateKey);
      });
    });

    const workedDays = workedDateKeys.size;
    summaries.set(commessaId, {
      totalHours,
      workedDays,
      averageHoursPerDay: workedDays > 0 ? totalHours / workedDays : 0,
      firstDoneAtMs: Number(stats.firstDoneAtMs || 0),
      firstDoneDateKey: String(stats.firstDoneDateKey || ""),
      lastResetAtMs: Number(stats.lastResetAtMs || 0),
      lastResetDateKey: String(stats.lastResetDateKey || ""),
      startDateKey,
      endDateKey,
      resetDateKey: String(workRange.resetDateKey || ""),
      startMode: workRange.startMode,
      workedDateKeys
    });
  });
  commessaWorkSummariesById = summaries;
  commessaHoursById = new Map(Array.from(summaries.entries()).map(([commessaId, summary]) => [commessaId, Number(summary.totalHours || 0)]));
}

function subscribeStatsForCommesse() {
  const activeIds = new Set(commesseById.keys());
  Array.from(unsubscribeCommessaStats.keys()).forEach((commessaId) => {
    if (!activeIds.has(commessaId)) {
      unsubscribeCommessaStats.get(commessaId)?.();
      unsubscribeCommessaStats.delete(commessaId);
      commessaStatsById.delete(commessaId);
      commessaWorkSummariesById.delete(commessaId);
      commessaHoursById.delete(commessaId);
    }
  });
  activeIds.forEach((commessaId) => {
    if (unsubscribeCommessaStats.has(commessaId)) return;
    const unsubscribe = db.collection("commesse").doc(commessaId).collection("impianti").onSnapshot((snapshot) => {
      const rawImpianti = snapshot.docs.map((doc) => ({ id: doc.id, ...doc.data() }));
      commessaStatsById.set(commessaId, calculateImpiantiStats(rawImpianti));
      recalculateCommessaWorkSummaries();
      renderCommesseHomeList();
      renderParentCommessaOverview();
      updateCommessaDashboard();
    }, (error) => console.error("Errore stats commessa:", error));
    unsubscribeCommessaStats.set(commessaId, unsubscribe);
  });
}

function subscribeHoursStats() {
  if (!unsubscribeHoursStats) {
    unsubscribeHoursStats = db.collection("oreReports").onSnapshot((snapshot) => {
      allHoursReports = snapshot.docs.map((doc) => ({ id: doc.id, ...doc.data() }));
      hoursReportsLoaded = true;
      recalculateCommessaWorkSummaries();
      renderParentCommessaOverview();
      renderSquadre();
      updateCommessaDashboard();
    }, (error) => console.error("Errore stats ore commesse:", error));
  }
  if (!unsubscribeHoursApprovals) {
    unsubscribeHoursApprovals = db.collection("oreApprovalRequests").onSnapshot((snapshot) => {
      allHoursApprovalRequests = snapshot.docs.map((doc) => ({ id: doc.id, ...doc.data() }));
      hoursApprovalsLoaded = true;
      hoursApprovalRequests = allHoursApprovalRequests;
      renderHoursApprovalRequests();
      renderSquadre();
    }, (error) => console.error("Errore richieste ore commesse:", error));
  }
}

function stopCommessaStatsSubscriptions() {
  unsubscribeCommessaStats.forEach((unsubscribe) => unsubscribe?.());
  unsubscribeCommessaStats.clear();
  commessaStatsById = new Map();
  if (unsubscribeHoursStats) {
    unsubscribeHoursStats();
    unsubscribeHoursStats = null;
  }
  if (unsubscribeHoursApprovals) {
    unsubscribeHoursApprovals();
    unsubscribeHoursApprovals = null;
  }
  commessaHoursById = new Map();
  commessaWorkSummariesById = new Map();
  allHoursReports = [];
  allHoursApprovalRequests = [];
  hoursReportsLoaded = false;
  hoursApprovalsLoaded = false;
  hoursApprovalRequests = [];
}

function sortCommesseByCreatedAtDesc(commesse) {
  const safeCommesse = Array.isArray(commesse) ? commesse : [];
  return [...safeCommesse].sort((a, b) => firestoreDateToMillis(b.createdAt) - firestoreDateToMillis(a.createdAt));
}

function getCommessaDisplayName(commessa = {}) {
  const parent = commessa.parentCommessaId ? commesseById.get(commessa.parentCommessaId) : null;
  const name = commessa.nome || "Commessa senza nome";
  return parent ? `${parent.nome || "Commessa padre"} / ${name}` : name;
}

function createCommessaOption(commessa, options = {}) {
  const option = document.createElement("option");
  option.value = commessa.id;
  const prefix = options.includeHierarchy && isSubcommessa(commessa) ? "↳ " : "";
  option.textContent = `${prefix}${getCommessaDisplayName(commessa)}`;
  return option;
}

function updateCommessaParentField() {
  if (!ui.commessaType || !ui.commessaParent) return;
  const isSub = ui.commessaType.value === "sub";
  ui.commessaParent.classList.toggle("hidden", !isSub);
  ui.commessaParent.required = isSub;
  ui.commessaParent.disabled = !isSub;
  if (!isSub) ui.commessaParent.value = "";
}

function populateCommessaParentSelect() {
  if (!ui.commessaParent) return;
  const previousValue = ui.commessaParent.value;
  ui.commessaParent.innerHTML = "<option value=''>Commessa padre</option>";
  sortCommesseByCreatedAtDesc(getMainCommesse()).forEach((commessa) => {
    ui.commessaParent.appendChild(createCommessaOption(commessa));
  });
  if (previousValue && commesseById.has(previousValue) && !isSubcommessa(commesseById.get(previousValue))) {
    ui.commessaParent.value = previousValue;
  }
  updateCommessaParentField();
}


function getTomorrowSquadraDateKey() {
  const tomorrow = new Date();
  tomorrow.setDate(tomorrow.getDate() + 1);
  return getDateKeyFromLocalDate(tomorrow);
}

function isQuickSquadraWindowOpen(now = new Date()) {
  const hour = now.getHours();
  return hour >= 12 && hour < 19;
}

function isValidActiveCommessaForQuickSquadra(commessa) {
  return Boolean(commessa?.id && String(commessa.nome || "").trim());
}

function canShowQuickSquadraButton(commessa) {
  return Boolean(currentUser && canManageData() && isQuickSquadraWindowOpen() && isValidActiveCommessaForQuickSquadra(commessa));
}

function getSquadreCountForCommessaDate(commessaId, dateKey) {
  const storicoDelGiorno = squadreHistoryByDate.get(dateKey) || new Map();
  const squad = storicoDelGiorno.get(commessaId) || squadreByCommessa.get(commessaId) || {};
  const rows = Array.isArray(squad.squadre) ? squad.squadre : getLegacySquadreRows(squad);
  return rows.filter((row) => String(row?.personale || "").trim() || String(row?.caposquadra || "").trim() || String(row?.mezzi || "").trim()).length;
}

function createQuickSquadraControls(commessa) {
  const wrap = document.createElement("span");
  wrap.className = "quick-squadra-controls";
  const button = document.createElement("button");
  button.type = "button";
  button.className = "btn quick-squadra-btn";
  button.textContent = "+ SQ";
  button.setAttribute("aria-label", `Crea squadra per ${commessa.nome || "commessa"}`);
  button.addEventListener("click", (event) => {
    event.preventDefault();
    event.stopPropagation();
    openQuickSquadraForm(commessa.id);
  });
  wrap.appendChild(button);
  const count = getSquadreCountForCommessaDate(commessa.id, getTomorrowSquadraDateKey());
  if (count > 0) {
    const badge = document.createElement("span");
    badge.className = "quick-squadra-badge";
    badge.textContent = String(count);
    badge.title = `${count} squadre già assegnate per domani`;
    wrap.appendChild(badge);
  }
  return wrap;
}

function openQuickSquadraForm(commessaId) {
  if (!currentUser || !canManageData()) {
    alert("Solo gli amministratori possono creare squadre dalla scheda commesse.");
    return;
  }
  if (!isQuickSquadraWindowOpen()) {
    alert("Le squadre dalla scheda commesse si possono creare solo tra le 12:00 e le 19:00.");
    renderCommesseHomeList();
    return;
  }
  const commessa = commesseById.get(commessaId);
  if (!isValidActiveCommessaForQuickSquadra(commessa)) {
    alert("Commessa non valida per la creazione squadra.");
    return;
  }
  openManagementPanel("squadre");
  const dateKey = getTomorrowSquadraDateKey();
  if (ui.squadraCommessa) ui.squadraCommessa.value = commessaId;
  if (ui.squadraRiferimento) ui.squadraRiferimento.value = dateKey;
  const storicoDelGiorno = squadreHistoryByDate.get(dateKey) || new Map();
  const data = storicoDelGiorno.get(commessaId) || {};
  const existingRows = Array.isArray(data.squadre) ? data.squadre : getLegacySquadreRows(data);
  if (existingRows.length) {
    setSquadraRowsFromData(data);
  } else {
    ui.squadraRows.innerHTML = "";
  }
  addSquadraRow({ quickNew: true });
  setTimeout(() => ui.squadraForm?.scrollIntoView({ behavior: "smooth", block: "start" }), 50);
}

function startQuickSquadraWindowTicker() {
  if (quickSquadraWindowTimer) return;
  quickSquadraWindowTimer = setInterval(() => {
    renderCommesseHomeList();
  }, 60 * 1000);
}

function renderCommessaHomeButton(commessa, index) {
  const row = document.createElement("div");
  row.className = "commessa-row";

  const btn = document.createElement("button");
  btn.type = "button";
  btn.className = "btn commessa-btn" + (commessa.id === selectedCommessaId ? " active" : "");
  btn.dataset.commessaId = commessa.id;
  btn.style.setProperty("--commessa-accent", getCommessaAccentColor(commessa.id, index));
  const codiceCommessa = String(commessa.codice || "").trim();
  const hasSubcommesse = getSubcommesse(commessa.id).length > 0;
  btn.classList.toggle("commessa-btn-parent", hasSubcommesse);
  btn.innerHTML = `
    <span class="commessa-home-main">
      <span>${escapeHTML(commessa.nome || "Commessa senza nome")}</span>
      ${hasSubcommesse ? `<span class="commessa-parent-indicator" title="Contiene subcommesse" aria-label="Contiene subcommesse">📂</span>` : ""}
    </span>
    ${codiceCommessa ? `<small class="muted">Cod. ${escapeHTML(codiceCommessa)}</small>` : ""}
  `;
  btn.addEventListener("click", () => selectCommessa(commessa.id, commessa.nome || "Commessa", commessa.codice || ""));

  row.appendChild(btn);
  if (canShowQuickSquadraButton(commessa)) {
    row.appendChild(createQuickSquadraControls(commessa));
  }
  return row;
}

function renderCommesseHomeList() {
  if (!ui.commesseLista) return;
  ui.commesseLista.innerHTML = "";
  if (!auth.currentUser) {
    ui.commesseLista.innerHTML = "<p class='muted'>Effettua login per visualizzare le commesse</p>";
    return;
  }
  if (commesseLoadState.status === "loading") {
    ui.commesseLista.innerHTML = "<p class='muted'>Caricamento commesse...</p>";
    return;
  }
  const commesse = sortCommesseByCreatedAtDesc(getMainCommesse());
  if (!commesse.length) {
    ui.commesseLista.innerHTML = `<p class='muted'>${escapeHTML(commesseLoadState.message || "Nessuna commessa disponibile")}</p>`;
    return;
  }
  if (commesseLoadState.status === "error" && commesseLoadState.message) {
    const warning = document.createElement("p");
    warning.className = "muted";
    warning.textContent = `${commesseLoadState.message}. Mostro le commesse salvate localmente.`;
    ui.commesseLista.appendChild(warning);
  }
  commesse.forEach((commessa, idx) => {
    ui.commesseLista.appendChild(renderCommessaHomeButton(commessa, idx));
  });
}

function renderParentCommessaOverview() {
  if (!ui.parentCommessaOverview) return;
  const selected = commesseById.get(selectedCommessaId) || {};
  const subcommesse = selectedCommessaId ? sortCommesseByCreatedAtDesc(getSubcommesse(selectedCommessaId)) : [];
  const isParentCommessa = subcommesse.length > 0;

  ui.parentCommessaOverview.classList.toggle("hidden", !isParentCommessa);
  ui.parentCommessaOverview.setAttribute("aria-hidden", isParentCommessa ? "false" : "true");
  ui.commessaOperationalCard?.classList.toggle("hidden", isParentCommessa);
  ui.impiantiCard?.classList.toggle("hidden", isParentCommessa);

  if (!isParentCommessa) {
    if (ui.parentSubcommesseList) ui.parentSubcommesseList.innerHTML = "";
    return;
  }

  const selectedName = selected.nome || selectedCommessaName || "Commessa";
  const aggregate = getParentCommessaAggregate(selectedCommessaId);
  if (ui.parentCommessaSummary) {
    ui.parentCommessaSummary.textContent = `${selectedName}: ${formatParentCommessaSummary(aggregate)}`;
  }
  if (ui.parentSubcommesseTitle) {
    ui.parentSubcommesseTitle.textContent = `Elenco subcommesse di ${selectedName}`;
  }
  if (!ui.parentSubcommesseList) return;
  ui.parentSubcommesseList.innerHTML = "";

  subcommesse.forEach((subcommessa) => {
    const row = document.createElement("article");
    row.className = "subcommessa-compact-card";
    const info = document.createElement("div");
    const codeText = String(subcommessa.codice || "").trim();
    info.innerHTML = `
      <h3>${escapeHTML(subcommessa.nome || "Subcommessa senza nome")}</h3>
      <p class="muted">${escapeHTML(formatSingleCommessaSummary(subcommessa.id))}</p>
      ${codeText ? `<p class="muted">Cod. ${escapeHTML(codeText)}</p>` : ""}
    `;
    const actions = document.createElement("div");
    actions.className = "item-actions";
    const openBtn = createButton("Apri subcommessa", () => selectCommessa(subcommessa.id, subcommessa.nome || "Subcommessa", subcommessa.codice || ""));
    openBtn.classList.add("subcommessa-open-btn");
    actions.appendChild(openBtn);
    if (canManageData()) {
      actions.appendChild(createButton("Sposta in principale", () => moveSubcommessaToMain(subcommessa.id)));
    }
    row.appendChild(info);
    row.appendChild(actions);
    ui.parentSubcommesseList.appendChild(row);
  });
}

function renderCommessaSelects() {
  const orderedCommesse = sortCommesseByCreatedAtDesc(Array.from(commesseById.values()));
  if (ui.squadraCommessa) ui.squadraCommessa.innerHTML = "<option value=''>Seleziona commessa</option>";
  if (ui.commessaTargetSelect) ui.commessaTargetSelect.innerHTML = "<option value=''>Usa commessa selezionata in home</option>";
  if (ui.resourceCommesse) ui.resourceCommesse.innerHTML = "";
  orderedCommesse.forEach((commessa) => {
    ui.squadraCommessa?.appendChild(createCommessaOption(commessa, { includeHierarchy: true }));
    ui.commessaTargetSelect?.appendChild(createCommessaOption(commessa, { includeHierarchy: true }));
    ui.resourceCommesse?.appendChild(createCommessaOption(commessa, { includeHierarchy: true }));
  });
  populateCommessaParentSelect();
}

async function createCommessa(event) {
  event.preventDefault();

  const user = auth.currentUser;
  if (!user) {
    alert("Devi fare login.");
    return;
  }

  const nome = ui.commessaName.value.trim();
  const codice = String(ui.commessaCode?.value || "").trim();
  const tipoCommessa = ui.commessaType?.value === "sub" ? "sub" : "main";
  const parentCommessaId = tipoCommessa === "sub" ? String(ui.commessaParent?.value || "").trim() : "";
  if (!nome) return;
  if (tipoCommessa === "sub" && !parentCommessaId) {
    alert("Seleziona la commessa padre per creare una subcommessa.");
    return;
  }
  if (!canManageData()) {
    alert("Solo un admin può aggiungere commesse.");
    return;
  }

  const commessaRef = await db.collection("commesse").add({
    nome,
    codice,
    parentCommessaId: parentCommessaId || null,
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
  updateCommessaParentField();
}

function subscribeDriveBridge() {
  unsubscribeDriveBridge = db.collection("appConfig").doc("driveBridge").onSnapshot(async (doc) => {
    const data = doc.exists ? doc.data() : null;
    const owner = data?.ownerEmail || ADMIN_EMAIL;
    driveBridgeState = {
      configured: Boolean(data?.configured || data?.rootFolderId),
      ownerEmail: owner,
      rootFolderId: CENTRAL_DRIVE_ROOT_FOLDER_ID
    };

    if (!driveBridgeState.configured) {
      updateDriveStatus(false);
      return;
    }

    driveRootFolderId = CENTRAL_DRIVE_ROOT_FOLDER_ID;
    if (canManageData()) {
      const storedToken = getStoredDriveToken();
      if (storedToken) {
        driveAccessToken = storedToken;
        window.googleDriveAccessToken = storedToken;
      }
      try {
        const secretDoc = await db.collection("appConfig").doc("driveAdminSecret").get();
        const secret = secretDoc.exists ? secretDoc.data() : {};
        driveAccessToken = secret.accessToken || driveAccessToken;
        window.googleDriveAccessToken = driveAccessToken || null;
        driveRootFolderId = CENTRAL_DRIVE_ROOT_FOLDER_ID;
        driveChatFolderId = secret.chatFolderId || driveChatFolderId || "";
        driveReportsFolderId = secret.reportsFolderId || driveReportsFolderId || "";
        driveSquadreFolderId = secret.squadreFolderId || driveSquadreFolderId || "";
        driveHelpCenterFolderId = secret.helpCenterFolderId || driveHelpCenterFolderId || "";
      } catch (error) {
        console.warn("Token Drive admin non leggibile, uso token locale se presente:", error);
      }
      updateDriveStatus(Boolean(driveAccessToken || driveRootFolderId));
      migrateLegacyDriveDataToCentralRoot().catch((error) => console.warn("Migrazione dati Drive vecchi non completata:", error));
      processPendingSheetExports();
      processAdminSheetExportQueue();
      return;
    }

    driveAccessToken = "";
    window.googleDriveAccessToken = null;
    updateDriveStatus(true);
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
  commesseLoadState = { status: "loading", message: "Caricamento commesse..." };
  renderCommesseHomeList();

  if (!validateFirebaseConfigForCommesse()) {
    commesseLoadState = { status: "error", message: "Impossibile connettersi al database" };
    loadCommesseFromLocalCache();
    refreshCommesseDependentUI(false);
    return;
  }

  unsubscribeCommesse = db
    .collection("commesse")
    .orderBy("createdAt", "desc")
    .onSnapshot((snapshot) => {
      const receivedCommesse = [];
      commesseById = new Map();

      snapshot.forEach((doc) => {
        const commessa = { id: doc.id, ...(doc.data() || {}) };
        receivedCommesse.push(commessa);
        commesseById.set(doc.id, commessa);
      });

      console.log("Commesse ricevute:", receivedCommesse);
      saveCommesseLocalCache(receivedCommesse);
      commesseLoadState = receivedCommesse.length
        ? { status: "loaded", message: "" }
        : { status: "empty", message: "Nessuna commessa disponibile" };

      const routeCommessaId = parseCommessaHash().id;
      const activeStoredId = routeCommessaId || localStorage.getItem(LAST_OPENED_COMMESSA_KEY) || localStorage.getItem(LAST_SELECTED_COMMESSA_KEY) || "";
      const shouldRestoreOpenCommessa = Boolean(!selectedCommessaId && activeStoredId && commesseById.has(activeStoredId));
      refreshCommesseDependentUI(Boolean(currentUser));
      if (!selectedCommessaId && shouldRestoreOpenCommessa) {
        const restored = commesseById.get(activeStoredId);
        if (restored) selectCommessa(restored.id, restored.nome || "Commessa", restored.codice || "");
      }
      tryAutoOpenAssignedCommessaAtStartup();
      renderNextActionCard();
    }, (error) => {
      console.error("Errore caricamento commesse:", error);
      loadCommesseFromLocalCache();
      commesseLoadState = {
        status: "error",
        message: getCommesseErrorMessage(error)
      };
      refreshCommesseDependentUI(false);
    });
}

function stopCommesseSubscription() {
  if (unsubscribeCommesse) {
    unsubscribeCommesse();
    unsubscribeCommesse = null;
  }
  stopCommessaStatsSubscriptions();
}

async function createGlobalCommessa(event) {
  event.preventDefault();
  const user = auth.currentUser;
  if (!user) return;
  if (!canManageData()) {
    alert("Solo un admin può gestire il Global.");
    return;
  }
  const nome = String(ui.globalCommessaName?.value || "").trim();
  if (!nome) return;
  await db.collection("globalCommesse").add({
    nome,
    creatoDa: user.email || "",
    createdAt: firebase.firestore.FieldValue.serverTimestamp()
  });
  ui.globalCommessaForm.reset();
}

function subscribeGlobalCommesse() {
  unsubscribeGlobalCommesse = db.collection("globalCommesse").orderBy("createdAt", "desc").onSnapshot((snapshot) => {
    globalCommesseById = new Map();
    if (ui.globalCommessaSelect) ui.globalCommessaSelect.innerHTML = "<option value=''>Seleziona commessa Global</option>";
    snapshot.forEach((doc) => {
      const item = { id: doc.id, ...doc.data() };
      globalCommesseById.set(doc.id, item);
      const option = document.createElement("option");
      option.value = doc.id;
      option.textContent = item.nome || "Commessa Global";
      ui.globalCommessaSelect?.appendChild(option);
    });
    renderGlobalCommesseManagementList();

    if (selectedGlobalCommessaId && !globalCommesseById.has(selectedGlobalCommessaId)) {
      selectedGlobalCommessaId = "";
      globalImpianti = [];
      renderGlobalImpianti();
      renderGlobalMap();
    }
    if (selectedGlobalCommessaId && ui.globalCommessaSelect) ui.globalCommessaSelect.value = selectedGlobalCommessaId;
    onGlobalCommessaSelectionChanged();
  }, (error) => {
    console.error("Errore caricamento commesse Global:", error);
    if (ui.globalImportFeedback) ui.globalImportFeedback.textContent = "Errore caricamento commesse Global.";
  });
}

function renderGlobalCommesseManagementList() {
  if (!ui.globalCommesseLista) return;
  if (!globalCommesseById.size) {
    ui.globalCommesseLista.innerHTML = "<p class='muted'>Nessuna commessa Global disponibile.</p>";
    return;
  }
  const canManage = canManageData();
  ui.globalCommesseLista.innerHTML = Array.from(globalCommesseById.values()).map((commessa) => {
    const isSelected = commessa.id === selectedGlobalCommessaId;
    return `
      <div class="simple-list-item">
        <span><b>${escapeHTML(commessa.nome || "Commessa Global")}</b>${isSelected ? " <small>(attiva)</small>" : ""}</span>
        <button type="button" class="btn btn-danger btn-small" data-delete-global-commessa="${escapeHTML(commessa.id)}" ${canManage ? "" : "disabled"}>
          ELIMINA COMMESSA
        </button>
      </div>
    `;
  }).join("");
}

async function onGlobalCommesseListClick(event) {
  const btn = event.target.closest("[data-delete-global-commessa]");
  if (!btn) return;
  if (!canManageData()) {
    alert("Solo un admin può eliminare commesse Global.");
    return;
  }
  const commessaId = String(btn.getAttribute("data-delete-global-commessa") || "").trim();
  if (!commessaId) return;
  const commessa = globalCommesseById.get(commessaId);
  const nomeCommessa = commessa?.nome || "Commessa Global";
  const confirmed = window.confirm(`Confermi eliminazione di "${nomeCommessa}" e di tutti gli impianti collegati?`);
  if (!confirmed) return;

  btn.disabled = true;
  if (ui.globalImportFeedback) ui.globalImportFeedback.textContent = "Eliminazione commessa Global in corso...";
  try {
    await deleteGlobalCommessaCascade(commessaId);
    if (ui.globalImportFeedback) ui.globalImportFeedback.textContent = `Commessa "${nomeCommessa}" eliminata con tutti i relativi impianti.`;
  } catch (error) {
    console.error("Errore eliminazione commessa Global:", error);
    if (ui.globalImportFeedback) ui.globalImportFeedback.textContent = "Errore durante l'eliminazione della commessa Global.";
  } finally {
    btn.disabled = false;
  }
}

async function deleteGlobalCommessaCascade(commessaId) {
  const impiantiRef = db.collection("globalCommesse").doc(commessaId).collection("impianti");
  const snapshot = await impiantiRef.get();
  for (let i = 0; i < snapshot.docs.length; i += 400) {
    const chunk = snapshot.docs.slice(i, i + 400);
    const batch = db.batch();
    chunk.forEach((doc) => batch.delete(doc.ref));
    await batch.commit();
  }
  await db.collection("globalCommesse").doc(commessaId).delete();
}

function stopGlobalCommesseSubscription() {
  if (unsubscribeGlobalCommesse) {
    unsubscribeGlobalCommesse();
    unsubscribeGlobalCommesse = null;
  }
}

function onGlobalCommessaSelectionChanged() {
  selectedGlobalCommessaId = String(ui.globalCommessaSelect?.value || "").trim();
  globalMapViewState.hasUserMoved = false;
  selectedGlobalImpianto = null;
  selectedGlobalImpiantoKey = "";
  closeGlobalImpiantoModal();
  closeGlobalSegnalazioneModal();
  stopGlobalImpiantiSubscription();
  if (selectedGlobalCommessaId) subscribeGlobalImpianti();
  refreshGlobalImportButtons();
  renderGlobalCommesseManagementList();
  if (!selectedGlobalCommessaId) {
    globalImpianti = [];
    globalImpiantoSearchTerm = "";
    if (ui.globalImpiantoSearch) ui.globalImpiantoSearch.value = "";
    renderGlobalSegnalazioneImpiantiOptions();
    renderGlobalImpianti();
    renderGlobalMap();
  }
}

function subscribeGlobalImpianti() {
  if (!selectedGlobalCommessaId) return;
  unsubscribeGlobalImpianti = db
    .collection("globalCommesse")
    .doc(selectedGlobalCommessaId)
    .collection("impianti")
    .onSnapshot((snapshot) => {
      const rawImpianti = snapshot.docs.map((doc) => ({ id: doc.id, ...doc.data() }));
      globalImpianti = combineImpiantiForView(rawImpianti);
      renderGlobalSegnalazioneImpiantiOptions();
      renderGlobalImpianti();
      renderGlobalMap();
      preloadImpiantiWeather(globalImpianti);
    }, (error) => {
      console.error("Errore caricamento impianti Global:", error);
      ui.globalImpiantiLista.innerHTML = "<p class='muted'>Errore caricamento impianti Global.</p>";
    });
}

function stopGlobalImpiantiSubscription() {
  if (unsubscribeGlobalImpianti) {
    unsubscribeGlobalImpianti();
    unsubscribeGlobalImpianti = null;
  }
}

function onGlobalExcelSelected(event) {
  pendingGlobalRows = [];
  refreshGlobalImportButtons();
  const file = event.target.files && event.target.files[0];
  if (!file) {
    ui.globalImportFeedback.textContent = "Nessun file selezionato.";
    return;
  }
  const reader = new FileReader();
  reader.onload = (e) => {
    try {
      const rows = rowsFromWorkbookBuffer(e.target.result);
      const prepared = prepareGlobalRowsForImport(rows);
      pendingGlobalRows = prepared.rows;
      ui.globalImportFeedback.textContent = `File Global letto: ${prepared.validRows} righe valide, ${pendingGlobalRows.length} impianti unici. Scartate ${prepared.invalidRows} righe (campi mancanti o non validi).`;
      refreshGlobalImportButtons();
    } catch (error) {
      console.error("Errore lettura Excel Global:", error);
      ui.globalImportFeedback.textContent = "Errore lettura file Global.";
    }
  };
  reader.readAsArrayBuffer(file);
}

async function importGlobalFromGoogleSheetUrl() {
  pendingGlobalRows = [];
  refreshGlobalImportButtons();
  const value = String(ui.globalSheetUrl?.value || "").trim();
  if (!value) {
    ui.globalImportFeedback.textContent = "Inserisci URL Google Sheet.";
    return;
  }
  try {
    const rows = await fetchGoogleSheetRows(value);
    const prepared = prepareGlobalRowsForImport(rows);
    pendingGlobalRows = prepared.rows;
    ui.globalImportFeedback.textContent = `Google Sheet Global letto: ${prepared.validRows} righe valide, ${pendingGlobalRows.length} impianti unici. Scartate ${prepared.invalidRows} righe (campi mancanti o non validi).`;
    refreshGlobalImportButtons();
  } catch (error) {
    console.error("Errore import Google Sheet Global:", error);
    ui.globalImportFeedback.textContent = "Errore lettura Google Sheet Global. Verifica link/condivisione.";
  }
}

async function importPendingGlobalRows() {
  if (!canManageData()) {
    alert("Solo un admin può importare nel Global.");
    return;
  }
  if (!selectedGlobalCommessaId) {
    alert("Seleziona una commessa Global.");
    return;
  }
  if (!pendingGlobalRows.length) {
    alert("Nessuna riga da importare.");
    return;
  }

  const ref = db.collection("globalCommesse").doc(selectedGlobalCommessaId).collection("impianti");
  const existingSnapshot = await ref.get();
  const existingByKey = new Map();
  existingSnapshot.forEach((doc) => {
    const data = doc.data() || {};
    const key = buildImpiantoKey(data);
    if (!existingByKey.has(key)) existingByKey.set(key, { id: doc.id, ...data });
  });

  let created = 0;
  let updated = 0;
  for (const row of pendingGlobalRows) {
    const key = buildImpiantoKey(row);
    const existing = existingByKey.get(key);
    if (!existing) {
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
      created += 1;
      continue;
    }
    const mergedCodicePrezzo = mergeMultiValue(existing.codicePrezzo, row.codicePrezzo);
    const mergedExtraFields = mergeExtraFields(existing.extraFields, row.extraFields);
    const extraFieldsChanged = JSON.stringify(mergedExtraFields || {}) !== JSON.stringify(existing.extraFields || {});
    const basePatch = {};
    if (row.area && !existing.area) basePatch.area = row.area;
    if (row.competenza && !existing.competenza) basePatch.competenza = row.competenza;
    if (row.descrizioneVia && !existing.descrizioneVia) basePatch.descrizioneVia = row.descrizioneVia;
    if (row.indirizzo && !existing.indirizzo) basePatch.indirizzo = row.indirizzo;
    if (row.tipologiaImpianto && !existing.tipologiaImpianto) basePatch.tipologiaImpianto = row.tipologiaImpianto;
    if (row.dittaEsecutrice && !existing.dittaEsecutrice) basePatch.dittaEsecutrice = row.dittaEsecutrice;
    if (row.areaMq != null && existing.areaMq == null) basePatch.areaMq = row.areaMq;
    if (row.gpsY != null && existing.gpsY == null) basePatch.gpsY = row.gpsY;
    if (row.gpsX != null && existing.gpsX == null) basePatch.gpsX = row.gpsX;

    if (mergedCodicePrezzo !== String(existing.codicePrezzo || "")) {
      await ref.doc(existing.id).set({
        codicePrezzo: mergedCodicePrezzo,
        voceRiferimento: mergeMultiValue(existing.voceRiferimento, row.voceRiferimento),
        tipologiaIntervento: mergeMultiValue(existing.tipologiaIntervento, row.tipologiaIntervento),
        lavorazioniRichieste: mergeMultiValue(existing.lavorazioniRichieste, row.lavorazioniRichieste),
        frequenzaAnnua: mergeMultiValue(existing.frequenzaAnnua, row.frequenzaAnnua),
        areaMq: row.areaMq ?? existing.areaMq ?? null,
        extraFields: mergedExtraFields,
        hasOrdinario: hasOrdinario(mergedCodicePrezzo),
        hasStraordinario: hasStraordinario(mergedCodicePrezzo),
        tipoManutenzione: classifyTipoManutenzione(mergedCodicePrezzo),
        ...basePatch
      }, { merge: true });
      updated += 1;
    } else if (extraFieldsChanged) {
      await ref.doc(existing.id).set({ extraFields: mergedExtraFields, ...basePatch }, { merge: true });
      updated += 1;
    } else if (Object.keys(basePatch).length) {
      await ref.doc(existing.id).set(basePatch, { merge: true });
      updated += 1;
    }
  }

  const skipped = Math.max(0, pendingGlobalRows.length - created - updated);
  pendingGlobalRows = [];
  if (ui.globalExcelFile) ui.globalExcelFile.value = "";
  refreshGlobalImportButtons();
  ui.globalImportFeedback.textContent = `Import Global completato: nuovi ${created}, aggiornati ${updated}, ignorati ${skipped}.`;
}

async function updateExistingGlobalRowsOnly() {
  if (!canManageData()) {
    alert("Solo un admin può aggiornare impianti nel Global.");
    return;
  }
  if (!selectedGlobalCommessaId) {
    alert("Seleziona una commessa Global.");
    return;
  }
  if (!pendingGlobalRows.length) {
    alert("Nessuna riga da aggiornare.");
    return;
  }

  const ref = db.collection("globalCommesse").doc(selectedGlobalCommessaId).collection("impianti");
  const existingSnapshot = await ref.get();
  const existingRows = existingSnapshot.docs.map((doc) => ({ id: doc.id, ...(doc.data() || {}) }));
  let updated = 0;
  let unmatched = 0;

  for (const row of pendingGlobalRows) {
    const match = findBestGlobalImpiantoMatch(row, existingRows);
    if (!match || match.score < 5) {
      unmatched += 1;
      continue;
    }

    const patch = {};
    if (match.record.gpsY == null && row.gpsY != null) patch.gpsY = row.gpsY;
    if (match.record.gpsX == null && row.gpsX != null) patch.gpsX = row.gpsX;
    if (row.comune && !match.record.comune) patch.comune = row.comune;
    if (row.indirizzo && !match.record.indirizzo) patch.indirizzo = row.indirizzo;
    if (row.descrizioneVia && !match.record.descrizioneVia) patch.descrizioneVia = row.descrizioneVia;
    if (row.area && !match.record.area) patch.area = row.area;
    if (row.tipologiaImpianto && !match.record.tipologiaImpianto) patch.tipologiaImpianto = row.tipologiaImpianto;
    if (row.dittaEsecutrice && !match.record.dittaEsecutrice) patch.dittaEsecutrice = row.dittaEsecutrice;
    if (row.areaMq != null && match.record.areaMq == null) patch.areaMq = row.areaMq;
    if (row.idSap && !match.record.idSap) patch.idSap = row.idSap;
    if (!Object.keys(patch).length) continue;
    await ref.doc(match.record.id).set(patch, { merge: true });
    updated += 1;
  }

  pendingGlobalRows = [];
  if (ui.globalExcelFile) ui.globalExcelFile.value = "";
  refreshGlobalImportButtons();
  ui.globalImportFeedback.textContent = `Aggiorna impianti completato: aggiornati ${updated}, non abbinati ${unmatched}. Nessun nuovo impianto creato.`;
}

function findBestGlobalImpiantoMatch(row, existingRows) {
  let best = null;
  for (const record of existingRows) {
    const score = scoreGlobalImpiantoMatch(row, record);
    if (!best || score > best.score) best = { record, score };
  }
  return best;
}

function scoreGlobalImpiantoMatch(source, target) {
  const sourceSap = normalizeText(source.idSap);
  const targetSap = normalizeText(target.idSap);
  if (sourceSap && targetSap && sourceSap === targetSap) return 100;

  let score = 0;
  const sourceName = normalizeText(source.denominazione);
  const targetName = normalizeText(target.denominazione);
  if (sourceName && targetName) {
    if (sourceName === targetName) score += 5;
    else if (sourceName.includes(targetName) || targetName.includes(sourceName)) score += 4;
    else if (tokenJaccard(sourceName, targetName) >= 0.6) score += 3;
  }
  const sourceComune = normalizeText(source.comune);
  const targetComune = normalizeText(target.comune);
  if (sourceComune && targetComune && sourceComune === targetComune) score += 2;

  const sourceAddress = normalizeAddress(source.indirizzo);
  const targetAddress = normalizeAddress(target.indirizzo);
  if (sourceAddress && targetAddress) {
    if (sourceAddress === targetAddress) score += 3;
    else if (sourceAddress.includes(targetAddress) || targetAddress.includes(sourceAddress)) score += 2;
    else if (tokenJaccard(sourceAddress, targetAddress) >= 0.5) score += 1;
  }
  return score;
}

function normalizeAddress(value) {
  return normalizeText(value)
    .replace(/\b(via|viale|piazza|strada|loc|localita)\b/g, " ")
    .replace(/\s+/g, " ")
    .trim();
}

function normalizeText(value) {
  return String(value || "")
    .toLowerCase()
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/[^a-z0-9\s]/g, " ")
    .replace(/\s+/g, " ")
    .trim();
}

function tokenJaccard(a, b) {
  const aSet = new Set(String(a || "").split(" ").filter(Boolean));
  const bSet = new Set(String(b || "").split(" ").filter(Boolean));
  if (!aSet.size || !bSet.size) return 0;
  let intersect = 0;
  aSet.forEach((token) => {
    if (bSet.has(token)) intersect += 1;
  });
  return intersect / (aSet.size + bSet.size - intersect);
}

function refreshGlobalImportButtons() {
  const disabled = !canManageData() || !selectedGlobalCommessaId || pendingGlobalRows.length === 0;
  if (ui.globalImportBtn) ui.globalImportBtn.disabled = disabled;
  if (ui.globalUpdateBtn) ui.globalUpdateBtn.disabled = disabled;
}

function onGlobalImpiantoSearchInput(event) {
  globalImpiantoSearchTerm = String(event.target.value || "").trim();
  renderGlobalImpianti();
  renderGlobalMap();
}

function onGlobalImpiantoSearchSubmit(event) {
  event.preventDefault();
  globalImpiantoSearchTerm = String(ui.globalImpiantoSearch?.value || "").trim();
  renderGlobalImpianti({ prioritizeTopResult: true });
  renderGlobalMap();
}

function getFilteredGlobalImpianti() {
  const normalizedSearch = globalImpiantoSearchTerm.toLowerCase();
  const filtered = globalImpianti.filter((impianto) => {
    if (!normalizedSearch) return true;
    const haystack = [impianto.idSap, impianto.denominazione, impianto.comune, impianto.area, impianto.competenza, impianto.descrizioneVia, impianto.indirizzo, impianto.codicePrezzo]
      .map((value) => String(value || "").toLowerCase())
      .join(" ");
    return haystack.includes(normalizedSearch);
  });
  return filtered.sort((a, b) => {
    if (!normalizedSearch) return 0;
    const aName = String(a.denominazione || "").toLowerCase();
    const bName = String(b.denominazione || "").toLowerCase();
    const aStarts = aName.startsWith(normalizedSearch);
    const bStarts = bName.startsWith(normalizedSearch);
    if (aStarts !== bStarts) return aStarts ? -1 : 1;
    return aName.localeCompare(bName);
  });
}

function renderGlobalImpianti(options = {}) {
  if (!ui.globalImpiantiLista) return;
  if (!selectedGlobalCommessaId) {
    ui.globalImpiantiLista.innerHTML = "<p class='muted'>Seleziona una commessa Global.</p>";
    globalImpiantiFiltered = [];
    return;
  }
  globalImpiantiFiltered = getFilteredGlobalImpianti();
  if (!globalImpiantiFiltered.length) {
    ui.globalImpiantiLista.innerHTML = "<p class='muted'>Nessun impianto trovato.</p>";
    return;
  }
  if (options.prioritizeTopResult && globalImpiantiFiltered.length === 1) {
    openGlobalImpiantoDetails(globalImpiantiFiltered[0], { focusOnMap: true, updateSearch: true });
  }
  const suggestions = globalImpiantiFiltered.slice(0, globalImpiantoSearchTerm ? 30 : 14);
  ui.globalImpiantiLista.innerHTML = suggestions.map((impianto) => {
    const key = buildImpiantoKey(impianto);
    const isSelected = key === selectedGlobalImpiantoKey;
    return `
      <button type="button" class="simple-list-item global-suggestion-item ${isSelected ? "is-selected" : ""}" data-global-impianto-key="${escapeHTML(key)}">
        <span>
          <b>${escapeHTML(impianto.denominazione || "Impianto")}</b> ${buildImpiantoWeatherBadgeMarkup(impianto)}<br>
          <small>${escapeHTML(impianto.idSap || "-")} • ${escapeHTML(impianto.comune || "-")} • ${escapeHTML(impianto.area || impianto.competenza || "-")}</small>
        </span>
      </button>
    `;
  }).join("");
  ui.globalImpiantiLista.querySelectorAll("[data-global-impianto-key]").forEach((item) => {
    item.addEventListener("click", () => {
      const key = item.getAttribute("data-global-impianto-key");
      const impianto = globalImpianti.find((row) => buildImpiantoKey(row) === key);
      if (!impianto) return;
      openGlobalImpiantoDetails(impianto, { focusOnMap: true });
    });
  });
}

function renderGlobalMap() {
  globalMarkerLayer.clearLayers();
  if (!selectedGlobalCommessaId) {
    ui.globalMapFeedback.textContent = "Seleziona una commessa Global.";
    globalMapViewState = { center: [44.4949, 11.3426], zoom: 6, hasUserMoved: false };
    globalMap.setView(globalMapViewState.center, globalMapViewState.zoom, { animate: false });
    return;
  }
  const source = globalImpiantiFiltered.length || globalImpiantoSearchTerm ? globalImpiantiFiltered : globalImpianti;
  const withGps = source.filter((item) => hasValidGlobalCoordinates(item));
  const dittaColorMap = buildGlobalDittaColorMap(withGps);
  if (!withGps.length) {
    ui.globalMapFeedback.textContent = "Nessuna coordinata GPS disponibile per questa commessa Global.";
    globalMap.setView(globalMapViewState.center, globalMapViewState.zoom, { animate: false });
    return;
  }
  ui.globalMapFeedback.textContent = `Mappa impianti Global: ${withGps.length} con coordinate GPS.`;
  const bounds = [];
  withGps.forEach((impianto) => {
    const impiantoKey = buildImpiantoKey(impianto);
    const isSelected = impiantoKey === selectedGlobalImpiantoKey;
    const ditta = getGlobalDittaLabel(impianto);
    const markerColor = dittaColorMap.get(ditta) || "#2563eb";
    const markerDetails = [
      `<b>${escapeHTML(impianto.denominazione || "Impianto")}</b> ${buildImpiantoWeatherBadgeMarkup(impianto)}`,
      `Codice Hera: ${escapeHTML(impianto.idSap || impianto.codiceHera || "-")}`,
      `Comune: ${escapeHTML(impianto.comune || "-")}`,
      `Area: ${escapeHTML(impianto.area || impianto.competenza || "-")}`,
      `Ditta esecutrice: ${escapeHTML(ditta)}`,
      `<button type="button" class="btn btn-small" data-global-marker-details="${escapeHTML(impiantoKey)}">Dettagli</button>`
    ].join("<br>");
    const marker = L.marker([impianto.gpsY, impianto.gpsX], {
      icon: L.divIcon({
        className: "global-map-pin-wrapper",
        html: `<span class="global-map-pin ${isSelected ? "is-selected" : ""}" style="--global-marker-color:${escapeHTML(markerColor)}">${escapeHTML(ditta)}</span>`,
        iconSize: [82, 24],
        iconAnchor: [41, 12]
      })
    }).addTo(globalMarkerLayer)
      .bindPopup(markerDetails);
    marker.on("click", () => openGlobalImpiantoDetails(impianto, { updateSearch: true }));
    marker.on("popupopen", (event) => {
      const popupElement = event.popup?.getElement();
      if (!popupElement) return;
      popupElement.querySelectorAll("[data-global-marker-details]").forEach((btn) => {
        btn.addEventListener("click", () => openGlobalImpiantoDetails(impianto, { updateSearch: true }));
      });
    });
    bounds.push([impianto.gpsY, impianto.gpsX]);
  });
  if (!globalMapViewState.hasUserMoved) {
    globalMap.fitBounds(bounds, { padding: [20, 20], maxZoom: 14 });
    const center = globalMap.getCenter();
    globalMapViewState.center = [center.lat, center.lng];
    globalMapViewState.zoom = globalMap.getZoom();
    globalMapViewState.hasUserMoved = true;
  } else {
    globalMap.setView(globalMapViewState.center, globalMapViewState.zoom, { animate: false });
  }
  preloadImpiantiWeather(getVisibleMapImpianti(globalMap, source));
}

function getGlobalDittaLabel(impianto) {
  return String(impianto?.dittaEsecutrice || "").trim() || "SENZA DITTA";
}

function buildGlobalDittaColorMap(impianti) {
  const palette = ["#2563eb", "#dc2626", "#059669", "#7c3aed", "#ea580c", "#0891b2", "#be185d", "#4f46e5", "#16a34a", "#b45309"];
  const colorMap = new Map();
  let nextIndex = 0;
  impianti.forEach((impianto) => {
    const ditta = getGlobalDittaLabel(impianto);
    if (colorMap.has(ditta)) return;
    colorMap.set(ditta, palette[nextIndex % palette.length]);
    nextIndex += 1;
  });
  return colorMap;
}

function openGlobalImpiantoDetails(impianto, options = {}) {
  if (!impianto || !ui.globalImpiantoDetailsBody) return;
  selectedGlobalImpianto = impianto;
  selectedGlobalImpiantoKey = buildImpiantoKey(impianto);
  if (options.updateSearch && ui.globalImpiantoSearch) {
    ui.globalImpiantoSearch.value = impianto.denominazione || impianto.idSap || "";
    globalImpiantoSearchTerm = String(ui.globalImpiantoSearch.value || "").trim();
  }
  const details = [
    ["ID-SAP", impianto.idSap || impianto.codiceHera || "-"],
    ["Denominazione impianto", impianto.denominazione || "-"],
    ["Tipologia impianto", impianto.tipologiaImpianto || impianto.tipologiaIntervento || "-"],
    ["Comune", impianto.comune || "-"],
    ["Area", impianto.area || impianto.competenza || "-"],
    ["Descrizione via", impianto.descrizioneVia || impianto.indirizzo || "-"],
    ["Coordinate GPS", formatGpsLabel(impianto)],
    ["Ditta esecutrice", impianto.dittaEsecutrice || "-"]
  ];
  Object.entries(impianto.extraFields || {}).forEach(([key, value]) => details.push([formatExtraFieldLabel(key), value]));
  const weatherBadge = `<p><b>Meteo:</b> ${buildImpiantoWeatherBadgeMarkup(impianto)}</p>`;
  ui.globalImpiantoDetailsBody.innerHTML = weatherBadge + details
    .map(([label, value]) => `<p><b>${escapeHTML(label)}:</b> ${escapeHTML(String(value == null || value === "" ? "-" : value))}</p>`)
    .join("");
  if (ui.globalImpiantoNavigateBtn) {
    ui.globalImpiantoNavigateBtn.onclick = async () => {
      if (!hasValidGlobalCoordinates(impianto)) {
        alert("Coordinate mancanti");
        return;
      }
      const canContinueNavigation = await confirmNavigationWeatherIfNeeded(impianto);
      if (!canContinueNavigation) return;
      window.open(`https://www.google.com/maps?q=${impianto.gpsY},${impianto.gpsX}`, "_blank");
    };
  }
  if (ui.globalImpiantoWhatsappBtn) {
    ui.globalImpiantoWhatsappBtn.onclick = () => handleOpenGlobalSegnalazioneClick(impianto);
  }
  ui.globalImpiantoDetails?.classList.remove("hidden");
  if (options.focusOnMap && hasValidGlobalCoordinates(impianto)) {
    globalMap.setView([impianto.gpsY, impianto.gpsX], Math.max(globalMap.getZoom(), 14), { animate: true });
  }
  renderGlobalMap();
  renderGlobalImpianti();
}

function closeGlobalImpiantoModal() {
  ui.globalImpiantoDetails?.classList.add("hidden");
  selectedGlobalImpianto = null;
  selectedGlobalImpiantoKey = "";
  renderGlobalMap();
  renderGlobalImpianti();
}

function hasValidGlobalCoordinates(impianto) {
  return isValidLatLon(Number(impianto?.gpsY), Number(impianto?.gpsX));
}

function formatGpsLabel(impianto) {
  if (!hasValidGlobalCoordinates(impianto)) return "Coordinate GPS non disponibili";
  return `${Number(impianto.gpsY).toFixed(6)}, ${Number(impianto.gpsX).toFixed(6)}`;
}

function formatExtraFieldLabel(key) {
  return String(key || "")
    .replace(/[_-]+/g, " ")
    .replace(/\s+/g, " ")
    .trim()
    .replace(/\b\w/g, (char) => char.toUpperCase());
}

function shareGlobalImpiantoViaWhatsapp(impianto) {
  handleOpenGlobalSegnalazioneClick(impianto);
}

function handleOpenGlobalSegnalazioneClick(impiantoFromAction = null) {
  const impianto = impiantoFromAction || selectedGlobalImpianto || null;
  console.log("Invio segnalazione manutenzione verde via WhatsApp", impianto);
  if (!impianto) {
    alert("Seleziona prima un impianto per creare la segnalazione.");
    return;
  }
  const opened = sendGlobalSegnalazioneViaWhatsappDirect(impianto);
  if (!opened) {
    alert("Impossibile aprire WhatsApp su questo dispositivo.");
  }
}

function sendGlobalSegnalazioneViaWhatsappDirect(impianto) {
  const message = buildGlobalWhatsappSegnalazioneMessage(impianto, "__________");
  const opened = safeOpenWhatsAppMessage(message)
    || openExternalUrl(`https://wa.me/?text=${encodeURIComponent(message)}`);
  if (ui.globalReportFeedback) {
    ui.globalReportFeedback.textContent = opened
      ? "WhatsApp aperto con segnalazione precompilata."
      : "Impossibile aprire WhatsApp su questo dispositivo.";
  }
  return opened;
}

function renderGlobalSegnalazioneImpiantiOptions() {
  if (!ui.globalReportImpiantoSelect) return;
  const current = selectedGlobalSegnalazioneKey;
  const options = globalImpianti
    .slice()
    .sort((a, b) => String(a.denominazione || "").localeCompare(String(b.denominazione || ""), "it"))
    .map((impianto) => {
      const key = buildImpiantoKey(impianto);
      return `<option value="${escapeHTML(key)}">${escapeHTML(impianto.denominazione || "Impianto")} • ${escapeHTML(impianto.idSap || "-")}</option>`;
    });
  ui.globalReportImpiantoSelect.innerHTML = `<option value="">Seleziona impianto Global</option>${options.join("")}`;
  if (current) ui.globalReportImpiantoSelect.value = current;
}

function openGlobalSegnalazioneModal(impianto = null) {
  renderGlobalSegnalazioneImpiantiOptions();
  if (ui.globalReportText) ui.globalReportText.value = "";
  if (ui.globalReportFeedback && !ui.globalReportFeedback.textContent) {
    ui.globalReportFeedback.textContent = globalImpianti.length ? "" : "Nessun impianto disponibile nella commessa Global selezionata.";
  }
  ui.globalReportModal?.classList.remove("hidden");
  if (ui.globalReportModal) ui.globalReportModal.style.display = "flex";
  ui.globalReportModal?.setAttribute("aria-hidden", "false");

  const preferred = impianto || selectedGlobalImpianto || null;
  if (preferred) {
    const key = buildImpiantoKey(preferred);
    if (ui.globalReportImpiantoSelect) ui.globalReportImpiantoSelect.value = key;
    applyGlobalSegnalazioneImpianto(preferred);
  } else {
    applyGlobalSegnalazioneImpianto(null);
  }
}

function closeGlobalSegnalazioneModal() {
  ui.globalReportModal?.classList.add("hidden");
  if (ui.globalReportModal) ui.globalReportModal.style.display = "";
  ui.globalReportModal?.setAttribute("aria-hidden", "true");
  selectedGlobalSegnalazioneKey = "";
}

function onGlobalSegnalazioneImpiantoChange(event) {
  const key = String(event.target?.value || "").trim();
  const impianto = globalImpianti.find((row) => buildImpiantoKey(row) === key) || null;
  applyGlobalSegnalazioneImpianto(impianto);
}

function applyGlobalSegnalazioneImpianto(impianto) {
  selectedGlobalSegnalazioneKey = impianto ? buildImpiantoKey(impianto) : "";
  const via = impianto?.descrizioneVia || impianto?.indirizzo || "-";
  const coordinateText = impianto ? formatGpsLabel(impianto) : "-";
  if (ui.globalReportIdSap) ui.globalReportIdSap.value = impianto?.idSap || impianto?.codiceHera || "";
  if (ui.globalReportDenominazione) ui.globalReportDenominazione.value = impianto?.denominazione || "";
  if (ui.globalReportComune) ui.globalReportComune.value = impianto?.comune || "";
  if (ui.globalReportVia) ui.globalReportVia.value = via === "-" ? "" : via;
  if (ui.globalReportCoordinate) ui.globalReportCoordinate.value = coordinateText === "Coordinate GPS non disponibili" ? "" : coordinateText;
  if (ui.globalReportDitta) ui.globalReportDitta.value = impianto?.dittaEsecutrice || "";
}

async function submitGlobalSegnalazioneWhatsapp(event) {
  event.preventDefault();
  const key = String(ui.globalReportImpiantoSelect?.value || "").trim();
  const impianto = globalImpianti.find((row) => buildImpiantoKey(row) === key);
  const testoSegnalazione = String(ui.globalReportText?.value || "").trim();
  if (!impianto) {
    if (ui.globalReportFeedback) ui.globalReportFeedback.textContent = "Seleziona prima un impianto Global.";
    return;
  }
  if (!testoSegnalazione) {
    if (ui.globalReportFeedback) ui.globalReportFeedback.textContent = "Inserisci il testo segnalazione.";
    return;
  }
  const message = buildGlobalWhatsappSegnalazioneMessage(impianto, testoSegnalazione);
  const opened = safeOpenWhatsAppMessage(message)
    || openExternalUrl(`https://wa.me/?text=${encodeURIComponent(message)}`);
  if (!opened) {
    if (ui.globalReportFeedback) ui.globalReportFeedback.textContent = "Impossibile aprire WhatsApp su questo dispositivo.";
    return;
  }
  if (ui.globalReportFeedback) ui.globalReportFeedback.textContent = "WhatsApp aperto con il messaggio precompilato.";
}

function buildGlobalWhatsappSegnalazioneMessage(impianto, testoSegnalazione) {
  const hasCoords = hasValidGlobalCoordinates(impianto);
  const coordinateRaw = hasCoords ? `${impianto.gpsY},${impianto.gpsX}` : "Coordinate GPS non disponibili";
  const mapsUrl = hasCoords ? `https://www.google.com/maps?q=${impianto.gpsY},${impianto.gpsX}` : "";
  const coordinateLine = hasCoords
    ? `🌐 Coordinate: ${coordinateRaw} ➡️ 🧭 Naviga verso l’impianto`
    : `🌐 Coordinate: ${coordinateRaw}`;
  const lines = [
    "🟢 SEGNALAZIONE MANUTENZIONE VERDE 🌿",
    "",
    "Si segnala che all’impianto:",
    "",
    `🆔 ID-SAP: ${impianto.idSap || impianto.codiceHera || "-"}`,
    `🏭 Impianto: ${impianto.denominazione || "-"}`,
    `🏙️ Comune: ${impianto.comune || "-"}`,
    `📍 Via: ${impianto.descrizioneVia || impianto.indirizzo || "-"}`,
    coordinateLine,
    `🏢 Ditta esecutrice: ${impianto.dittaEsecutrice || "-"}`,
    "",
    "📢 Segnalazione:",
    testoSegnalazione,
    "",
    "🙏 Si chiede gentilmente un riscontro.",
    "",
    "Grazie."
  ];
  if (mapsUrl) lines.push("", mapsUrl);
  return lines.join("\n");
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
  const response = await fetchWithTimeoutAndRetry(url, { ...options, headers }, {
    timeoutMs: NETWORK_DEFAULT_TIMEOUT_MS,
    retries: 2
  });
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
    const useDriveUpload = Boolean(file);
    if (file) {
      fileSize = Number(file.size || 0);
      fileName = file.name || "documento";
      fileType = file.type || "application/octet-stream";
      if (useDriveUpload) {
        ui.privateDocsFeedback.textContent = "Caricamento sul cloud centralizzato...";
        const upload = await uploadBlobToDrive(file, fileName, fileType, driveReportsFolderId, { driveType: "DOCUMENTI", commessaName: "Documenti personali" });
        driveFileId = upload.fileId;
        driveWebViewLink = upload.webViewLink;
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
  updateCommessaDashboard();
}

function openCommessaPhoneResources() {
  const phones = getResourcesByCommessa(selectedCommessaId, "phone");
  if (!phones.length) {
    alert("Nessun contatto telefonico disponibile per questa commessa.");
    return;
  }
  openCommessaResourceWindow("phone");
}

function scrollToHomeSquadreSection() {
  closeImpiantiPage();
  setTimeout(() => ui.squadreLista?.closest(".squadre-card")?.scrollIntoView({ behavior: "smooth", block: "start" }), 80);
}

function openCommessaResourceWindow(type) {
  activeResourceTypeForViewer = type;
  if (!selectedCommessaId) return;
  setCommessaHash(`&resource=${encodeURIComponent(type)}`);
  applyRoute();
}

function closeCommessaResourceViewer() {
  if (parseCommessaHash().resource && selectedCommessaId) {
    setCommessaHash();
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

function openExternalUrl(url, options = {}) {
  if (!url) return false;
  const target = options?.target || "_blank";
  const features = options?.features || "noopener";
  try {
    const openedWindow = window.open(url, target, features);
    if (!openedWindow) {
      if (options?.allowSameWindowFallback !== false) {
        window.location.href = url;
        return true;
      }
      return false;
    }
    return true;
  } catch (error) {
    console.error("Errore apertura URL esterna:", error);
    if (options?.allowSameWindowFallback !== false) {
      try {
        window.location.href = url;
        return true;
      } catch (fallbackError) {
        console.error("Errore fallback apertura URL esterna:", fallbackError);
      }
    }
    return false;
  }
}

function buildWhatsAppWebUrl(encodedMessage, phone = "") {
  const normalizedPhone = sanitizePhone(phone).replace(/\+/g, "");
  if (normalizedPhone) return `https://wa.me/${normalizedPhone}?text=${encodedMessage}`;
  return `https://wa.me/?text=${encodedMessage}`;
}

function safeOpenWhatsAppMessage(message, options = {}) {
  const encodedMessage = encodeURIComponent(String(message || ""));
  const appUrl = options?.appUrl || `whatsapp://send?text=${encodedMessage}`;
  const webUrl = options?.webUrl || buildWhatsAppWebUrl(encodedMessage, options?.phone);
  const target = options?.target || "_blank";
  const usePopup = options?.usePopup !== false;
  if (!usePopup || target === "_self") {
    try {
      window.location.href = webUrl;
      return true;
    } catch (error) {
      console.error("Errore apertura WhatsApp in stessa finestra:", error);
      return false;
    }
  }
  return openExternalUrl(webUrl, {
    target,
    features: "noopener",
    allowSameWindowFallback: true
  }) || openExternalUrl(appUrl, {
    target,
    features: "noopener",
    allowSameWindowFallback: false
  });
}

function openPhoneOnWhatsApp(value) {
  const raw = sanitizePhone(value).replace(/\+/g, "");
  if (!raw) return;
  const opened = safeOpenWhatsAppMessage("", {
    appUrl: `whatsapp://send?phone=${raw}`,
    webUrl: `https://wa.me/${raw}`,
    phone: raw
  });
  if (!opened) alert("Impossibile aprire WhatsApp su questo dispositivo.");
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

function selectCommessa(id, nome, codice = "") {
  selectedCommessaId = id;
  selectedCommessaName = nome;
  setCommessaWeatherRefreshStatus("");
  mainMapViewState.hasUserMoved = false;
  activeNearbyImpiantoContext = null;
  localStorage.setItem(LAST_SELECTED_COMMESSA_KEY, id);
  ensureImpiantoWeatherPersistentCacheLoaded();
  setImpiantiViewMode("todo");
  if (!ui.commessaTargetSelect.value) {
    ui.commessaTargetSelect.value = id;
  }
  const codeText = String(codice || "").trim();
  const selectedCommessa = commesseById.get(id) || {};
  const parentCommessa = selectedCommessa.parentCommessaId ? commesseById.get(selectedCommessa.parentCommessaId) : null;
  const hierarchyText = parentCommessa ? `Subcommessa di ${parentCommessa.nome || "commessa padre"}: ${nome}` : `Commessa selezionata: ${nome}`;
  ui.commessaAttiva.textContent = codeText ? `${hierarchyText} • Cod. commessa: ${codeText}` : hierarchyText;
  updateCommessaContextUI();
  updateCommessaDashboard();
  ui.importBtn.disabled = !auth.currentUser || pendingRows.length === 0 || !getTargetCommessaId() || !canManageData();
  ui.exportCurrentCommessaBtn.disabled = !auth.currentUser || !canManageData();
  updateCommessaButtonsActive();
  renderResourceButtonsForCommessa();
  closeCommessaResourceViewer();

  stopImpiantiSubscription();
  stopCommessaNotesSubscription();
  const hasSubcommesse = getSubcommesse(id).length > 0;
  renderParentCommessaOverview();
  updateCommessaDashboard();
  if (!hasSubcommesse) {
    subscribeImpianti();
    subscribeCommessaNotes();
  }
  setCurrentWorkflowStep("open-commessa");
  const commessaRoute = parseCommessaHash();
  if (!hasSubcommesse && commessaRoute.notes) openCommessaNotesPage();
  else if (!hasSubcommesse && commessaRoute.atex) openImpiantiPage(`&atex=${encodeURIComponent(commessaRoute.atex)}`);
  else if (!hasSubcommesse) openImpiantiPage(commessaRoute.impianto ? `&impianto=${encodeURIComponent(commessaRoute.impianto)}` : "");
  else openImpiantiPage("");
}

function updateCommessaContextUI() {
  const selected = commesseById.get(selectedCommessaId) || {};
  const rawName = selectedCommessaName || "Commessa";
  let displayName = rawName;
  let codeText = String(selected.codice || "").trim();
  if (!codeText) {
    const splitMatch = String(rawName).trim().match(/^(.+?)\s+(\d{3,})$/);
    if (splitMatch) {
      displayName = splitMatch[1];
      codeText = splitMatch[2];
    }
  }
  if (ui.commessaFocusLabel) {
    ui.commessaFocusLabel.textContent = displayName.toUpperCase();
  }
  if (ui.commessaFocusCode) {
    const parent = selected.parentCommessaId ? commesseById.get(selected.parentCommessaId) : null;
    ui.commessaFocusCode.textContent = codeText;
    ui.commessaFocusCode.title = parent ? `Subcommessa di: ${parent.nome || "commessa padre"}` : "";
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

function isNetworkOffline() {
  return typeof navigator !== "undefined" && navigator.onLine === false;
}

function loadPendingImpiantoActions() {
  try {
    const raw = localStorage.getItem(PENDING_IMPIANTO_ACTIONS_KEY);
    const parsed = raw ? JSON.parse(raw) : [];
    return Array.isArray(parsed) ? parsed.filter((item) => item && item.id && item.type === "done") : [];
  } catch (error) {
    console.warn("Azioni impianto pending non leggibili:", error);
    return [];
  }
}

function savePendingImpiantoActions() {
  try {
    localStorage.setItem(PENDING_IMPIANTO_ACTIONS_KEY, JSON.stringify(pendingImpiantoActions));
  } catch (error) {
    console.warn("Azioni impianto pending non salvate:", error);
  }
}

function loadWhazzupPendingDoneEntries() {
  try {
    const raw = localStorage.getItem(WHAZZUP_PENDING_DONE_KEY);
    const parsed = raw ? JSON.parse(raw) : [];
    return Array.isArray(parsed) ? parsed.filter((entry) => entry?.commessaId && entry?.impiantoKey) : [];
  } catch (error) {
    console.warn("Flag Whazzup pending non leggibili:", error);
    return [];
  }
}

function saveWhazzupPendingDoneEntries(entries) {
  try {
    localStorage.setItem(WHAZZUP_PENDING_DONE_KEY, JSON.stringify(Array.isArray(entries) ? entries : []));
  } catch (error) {
    console.warn("Flag Whazzup pending non salvati:", error);
  }
}

function upsertWhazzupPendingDoneEntry(impianto, pressedAt = new Date()) {
  const commessaId = String(selectedCommessaId || "").trim();
  const impiantoKey = buildImpiantoKey(impianto);
  if (!commessaId || !impiantoKey) return;
  const doneBy = auth.currentUser?.displayName || auth.currentUser?.email || "Operatore";
  const entries = loadWhazzupPendingDoneEntries();
  const existingIndex = entries.findIndex((entry) => entry.commessaId === commessaId && entry.impiantoKey === impiantoKey);
  const payload = {
    commessaId,
    commessaName: selectedCommessaName || "Commessa",
    impiantoKey,
    impiantoIds: getImpiantoDocIds(impianto).filter(Boolean),
    impiantoName: impianto?.denominazione || "Impianto",
    userId: currentUser?.uid || "",
    userEmail: currentUser?.email || "",
    doneBy,
    pendingAt: pressedAt.toISOString()
  };
  if (existingIndex >= 0) entries.splice(existingIndex, 1, { ...entries[existingIndex], ...payload });
  else entries.push(payload);
  saveWhazzupPendingDoneEntries(entries);
}

function clearWhazzupPendingDoneEntry(impianto) {
  const commessaId = String(selectedCommessaId || "").trim();
  const impiantoKey = buildImpiantoKey(impianto);
  if (!commessaId || !impiantoKey) return;
  const nextEntries = loadWhazzupPendingDoneEntries().filter((entry) => !(entry.commessaId === commessaId && entry.impiantoKey === impiantoKey));
  saveWhazzupPendingDoneEntries(nextEntries);
}

function getCurrentUserPendingActions() {
  const uid = currentUser?.uid || "";
  if (!uid) return [];
  return pendingImpiantoActions.filter((action) => !action.userId || action.userId === uid);
}

function isPendingWhatsappAction(action) {
  return action
    && action.type === "done"
    && action.whatsappStatus !== "sent";
}

function getPendingWhatsappActions() {
  return getCurrentUserPendingActions()
    .filter(isPendingWhatsappAction)
    .sort((a, b) => String(b.doneAt || "").localeCompare(String(a.doneAt || "")));
}

function isActionWaitingForSync(action) {
  return ["pending", "syncing", "syncFailed"].includes(String(action?.status || "pending"));
}

function doesPendingActionMatchImpianto(action, commessaId, impianto) {
  if (!action || action.commessaId !== commessaId) return false;
  const impiantoKey = buildImpiantoKey(impianto);
  if (action.impiantoKey && impiantoKey && action.impiantoKey === impiantoKey) return true;
  const actionIds = new Set(Array.isArray(action.impiantoIds) ? action.impiantoIds : []);
  return getImpiantoDocIds(impianto).some((id) => actionIds.has(id));
}

function getPendingActionForImpianto(commessaId, impianto) {
  return getCurrentUserPendingActions().find((action) => isPendingWhatsappAction(action) && doesPendingActionMatchImpianto(action, commessaId, impianto));
}

function clearPendingImpiantoFlags(impianto) {
  if (!impianto?.pendingActionId && !impianto?.pendingActionStatus && !impianto?.pendingWhatsappStatus) return impianto;
  const { pendingActionId, pendingActionStatus, pendingWhatsappStatus, ...clean } = impianto;
  return clean;
}

function applyPendingActionsToImpianti(impianti, commessaId) {
  const activeActions = getCurrentUserPendingActions()
    .filter((action) => action.commessaId === commessaId && isActionWaitingForSync(action));
  return impianti.map((impianto) => {
    const cleanImpianto = clearPendingImpiantoFlags(impianto);
    const action = activeActions.find((entry) => doesPendingActionMatchImpianto(entry, commessaId, cleanImpianto));
    if (!action) return cleanImpianto;
    return {
      ...cleanImpianto,
      done: true,
      doneAt: action.doneAt ? new Date(action.doneAt) : cleanImpianto.doneAt,
      doneBy: action.doneBy || cleanImpianto.doneBy || "Operatore",
      pendingActionId: action.id,
      pendingActionStatus: action.status || "pending",
      pendingWhatsappStatus: action.whatsappStatus || "pending"
    };
  });
}

function getSerializableImpiantoSnapshot(impianto) {
  const keys = [
    "denominazione", "comune", "indirizzo", "idSap", "codicePrezzo", "voceRiferimento",
    "lavorazioniRichieste", "tipologiaIntervento", "gpsY", "gpsX", "hasStraordinario",
    "tipoManutenzione"
  ];
  return keys.reduce((acc, key) => {
    if (impianto[key] !== undefined) acc[key] = impianto[key];
    return acc;
  }, {});
}

function upsertPendingDoneAction(impianto, impiantoIds, doneAtLocal, doneByLocal) {
  const impiantoKey = buildImpiantoKey(impianto);
  const existingIndex = pendingImpiantoActions.findIndex((action) => (
    action.type === "done"
    && action.userId === currentUser?.uid
    && action.commessaId === selectedCommessaId
    && action.impiantoKey === impiantoKey
    && action.whatsappStatus !== "sent"
  ));
  const actionId = existingIndex >= 0 ? pendingImpiantoActions[existingIndex].id : `${currentUser?.uid || "user"}:${selectedCommessaId}:${impiantoKey || impiantoIds[0] || "impianto"}:${Date.now()}`;
  const payload = buildImpiantoWhatsAppPayload({ ...impianto, doneAt: doneAtLocal, doneBy: doneByLocal }, {
    doneAt: doneAtLocal,
    operatorName: doneByLocal
  });
  const action = {
    ...(existingIndex >= 0 ? pendingImpiantoActions[existingIndex] : {}),
    id: actionId,
    type: "done",
    status: "pending",
    whatsappStatus: "pending",
    userId: currentUser?.uid || "",
    userEmail: currentUser?.email || "",
    commessaId: selectedCommessaId,
    commessaName: selectedCommessaName || "Commessa",
    impiantoIds,
    impiantoKey,
    impiantoName: impianto.denominazione || "Impianto",
    impianto: getSerializableImpiantoSnapshot(impianto),
    doneAt: doneAtLocal.toISOString(),
    doneBy: doneByLocal,
    whatsappMessage: payload.message,
    createdAt: new Date().toISOString(),
    updatedAt: new Date().toISOString()
  };
  if (existingIndex >= 0) pendingImpiantoActions.splice(existingIndex, 1, action);
  else pendingImpiantoActions.push(action);
  savePendingImpiantoActions();
  renderPendingWhatsappList();
  return action;
}

function markPendingActionStatus(actionId, patch) {
  const index = pendingImpiantoActions.findIndex((action) => action.id === actionId);
  if (index < 0) return null;
  pendingImpiantoActions[index] = {
    ...pendingImpiantoActions[index],
    ...patch,
    updatedAt: new Date().toISOString()
  };
  savePendingImpiantoActions();
  renderPendingWhatsappList();
  if (selectedCommessaId) {
    currentImpianti = applyPendingActionsToImpianti(currentImpianti, selectedCommessaId);
    renderImpianti();
    renderMap();
  }
  return pendingImpiantoActions[index];
}

function buildPendingActionImpianto(action) {
  return {
    ...(action.impianto || {}),
    id: Array.isArray(action.impiantoIds) ? action.impiantoIds[0] : "",
    sourceIds: Array.isArray(action.impiantoIds) ? action.impiantoIds : [],
    done: true,
    doneAt: action.doneAt ? new Date(action.doneAt) : new Date(),
    doneBy: action.doneBy || "Operatore"
  };
}

async function syncPendingImpiantoActions() {
  if (isNetworkOffline() || !currentUser) {
    renderPendingWhatsappList();
    return;
  }
  const actionsToSync = getCurrentUserPendingActions().filter((action) => isActionWaitingForSync(action));
  if (!actionsToSync.length) {
    renderPendingWhatsappList();
    return;
  }
  const syncedWhatsappActions = [];
  for (const action of actionsToSync) {
    const impiantoIds = Array.isArray(action.impiantoIds) ? action.impiantoIds.filter(Boolean) : [];
    if (!action.commessaId || !impiantoIds.length) continue;
    markPendingActionStatus(action.id, { status: "syncing", lastError: "" });
    try {
      const doneAtDate = action.doneAt ? new Date(action.doneAt) : new Date();
      await setImpiantoDone(action.commessaId, impiantoIds, true, {
        doneAt: doneAtDate,
        doneBy: action.doneBy || currentUser.displayName || currentUser.email || "Operatore"
      });
      const exportPayload = {
        commessaId: action.commessaId,
        commessaName: action.commessaName || "Commessa",
        impianto: buildPendingActionImpianto(action)
      };
      if (!canManageData()) await queueSheetExportForAdmin(exportPayload);
      else scheduleCommessaSheetSync(action.commessaId, action.commessaName || "Commessa", 200);
      await publishGlobalNotificationEvent("impianto-done", {
        title: "Impianto completato",
        body: `${action.doneBy || "Operatore"} ha premuto FATTO su ${action.impiantoName || "Impianto"} (${action.commessaName || "Commessa"}).`,
        commessaId: action.commessaId,
        commessaName: action.commessaName || "Commessa",
        impiantoName: action.impiantoName || "Impianto",
        impiantoKey: action.impiantoKey || ""
      });
      const updatedAction = markPendingActionStatus(action.id, {
        status: "synced",
        syncedAt: new Date().toISOString(),
        whatsappStatus: action.whatsappStatus === "sent" ? "sent" : "pending",
        lastError: ""
      });
      if (updatedAction && updatedAction.whatsappStatus !== "sent") syncedWhatsappActions.push(updatedAction);
    } catch (error) {
      console.error("Sincronizzazione pendingAction FATTO non riuscita:", error);
      markPendingActionStatus(action.id, {
        status: "syncFailed",
        lastError: String(error && error.message ? error.message : error).slice(0, 500)
      });
    }
  }
  if (selectedCommessaId) {
    currentImpianti = applyPendingActionsToImpianti(currentImpianti, selectedCommessaId);
    renderImpianti();
    renderMap();
  }
  const alertable = syncedWhatsappActions.filter((action) => !pendingWhatsappAlertShownForSyncIds.has(action.id));
  if (alertable.length) {
    alertable.forEach((action) => pendingWhatsappAlertShownForSyncIds.add(action.id));
    alert("Ci sono messaggi WhatsApp da inviare");
  }
  renderPendingWhatsappList();
}

function openPendingWhatsApp(actionId) {
  if (isNetworkOffline()) {
    alert("WhatsApp non può essere inviato automaticamente offline. Torna online e premi di nuovo Invia WhatsApp.");
    return;
  }
  const action = pendingImpiantoActions.find((item) => item.id === actionId);
  if (!action) return;
  const message = action.whatsappMessage || buildImpiantoWhatsAppPayload(buildPendingActionImpianto(action), {
    doneAt: action.doneAt ? new Date(action.doneAt) : new Date(),
    operatorName: action.doneBy || "Operatore"
  }).message;
  const opened = safeOpenWhatsAppMessage(message);
  if (!opened) {
    alert("Impossibile aprire WhatsApp automaticamente su questo dispositivo.");
    return;
  }
  markPendingActionStatus(actionId, {
    whatsappStatus: "whatsappOpened",
    whatsappOpenedAt: new Date().toISOString()
  });
}

function markPendingWhatsAppSent(actionId) {
  const currentStatus = pendingImpiantoActions.find((item) => item.id === actionId)?.status || "pending";
  markPendingActionStatus(actionId, {
    whatsappStatus: "sent",
    sentAt: new Date().toISOString(),
    status: isActionWaitingForSync({ status: currentStatus }) ? currentStatus : "complete"
  });
}

function renderPendingWhatsappList() {
  if (!ui.pendingWhatsappCard || !ui.pendingWhatsappList || !ui.pendingWhatsappSummary) return;
  const actions = getPendingWhatsappActions();
  const count = actions.length;
  ui.pendingWhatsappCard.classList.toggle("hidden", count === 0);
  ui.pendingWhatsappBadge?.classList.toggle("hidden", count === 0);
  ui.pendingWhatsappSummary.textContent = count
    ? `${count} messagg${count === 1 ? "io" : "i"} WhatsApp da inviare manualmente appena sei online.`
    : "Nessun messaggio in attesa.";
  ui.pendingWhatsappList.innerHTML = "";
  actions.forEach((action) => {
    const row = document.createElement("div");
    row.className = "simple-list-item stacked pending-whatsapp-item";
    const doneLabel = action.doneAt ? new Date(action.doneAt).toLocaleString("it-IT") : "-";
    const whatsappLabel = action.whatsappStatus === "whatsappOpened" ? "WhatsApp aperto" : "WhatsApp in attesa";
    row.innerHTML = `
      <div>
        <strong>${escapeHTML(action.impiantoName || "Impianto")}</strong>
        <p class="muted">${escapeHTML(action.commessaName || "Commessa")} • ${escapeHTML(doneLabel)} • ${escapeHTML(whatsappLabel)}</p>
        ${action.lastError ? `<p class="muted pending-error">Ultimo errore sync: ${escapeHTML(action.lastError)}</p>` : ""}
      </div>
    `;
    const actionsBox = document.createElement("div");
    actionsBox.className = "item-actions";
    const sendBtn = createButton("Invia WhatsApp", () => openPendingWhatsApp(action.id));
    sendBtn.disabled = isNetworkOffline();
    actionsBox.appendChild(sendBtn);
    actionsBox.appendChild(createButton("Segna come inviato", () => markPendingWhatsAppSent(action.id)));
    row.appendChild(actionsBox);
    ui.pendingWhatsappList.appendChild(row);
  });
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
      currentImpianti = applyPendingActionsToImpianti(combineImpiantiForView(rawImpianti), selectedCommessaId);
      renderHeaderActivitySummary();
      updateCommessaDashboard();
      renderImpianti();
      renderMap();
      runWhazzupPendingDoneSafetyCheck();
      preloadCommessaWeatherForVisibleImpianti();
      evaluateImpiantoProximityAlerts();
      if (!currentUserPos) fetchWeather();

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

function subscribeCommessaNotes() {
  if (!selectedCommessaId) return;
  unsubscribeCommessaNotes = db
    .collection("commesse")
    .doc(selectedCommessaId)
    .collection("noteCommessa")
    .orderBy("noteDate", "desc")
    .onSnapshot((snapshot) => {
      currentCommessaNotes = snapshot.docs
        .map((doc) => ({ id: doc.id, ...doc.data() }))
        .sort((a, b) => {
          const dateCompare = String(b.noteDate || "").localeCompare(String(a.noteDate || ""));
          if (dateCompare) return dateCompare;
          return firestoreDateToMillis(b.createdAt) - firestoreDateToMillis(a.createdAt);
        });
      renderCommessaNotes();
      updateCommessaDashboard();
      if (selectedCommessaId && ui.impiantiLista && !ui.impiantiPage?.classList.contains("hidden")) renderImpianti();
    }, (error) => {
      console.error(error);
      if (ui.commessaNotesList) ui.commessaNotesList.innerHTML = "<p class='muted'>Errore caricamento note commessa.</p>";
    });
}

function stopCommessaNotesSubscription() {
  if (unsubscribeCommessaNotes) {
    unsubscribeCommessaNotes();
    unsubscribeCommessaNotes = null;
  }
  currentCommessaNotes = [];
  renderCommessaNotes();
}

function getTodayDateKey() {
  return new Date().toISOString().slice(0, 10);
}

function toggleCommessaNoteForm() {
  openCommessaNotesPage();
}

function getCommessaNoteTitle(note) {
  const title = String(note?.title || "").trim();
  if (title) return title;
  const legacyType = String(note?.type || "").trim();
  if (legacyType) return legacyType;
  return String(note?.text || "Nota commessa").trim().split(/\n+/)[0].slice(0, 80) || "Nota commessa";
}

function getImpiantoDisplayLabel(impianto) {
  return String(impianto?.denominazione || impianto?.idSap || impianto?.codiceHera || impianto?.codicePrezzo || "Impianto").trim();
}

function getImpiantoSearchLabel(impianto) {
  return [
    getImpiantoDisplayLabel(impianto),
    impianto?.comune,
    impianto?.idSap,
    impianto?.indirizzo || impianto?.descrizioneVia
  ].map((value) => String(value || "").trim()).filter(Boolean).join(" • ");
}

function getImpiantoSearchText(impianto) {
  return [
    impianto?.denominazione,
    impianto?.comune,
    impianto?.idSap,
    impianto?.indirizzo,
    impianto?.descrizioneVia,
    impianto?.codiceHera,
    impianto?.codicePrezzo,
    impianto?.voceRiferimento
  ].map((value) => String(value || "").toLowerCase()).join(" ");
}

function getCommessaNoteLinkedNotes(impianto) {
  const impiantoKey = buildImpiantoKey(impianto);
  if (!impiantoKey) return [];
  return currentCommessaNotes.filter((note) => note.impiantoKey && note.impiantoKey === impiantoKey);
}

function updateCommessaNoteImpiantoSelection(selectedKey = "", options = {}) {
  const key = String(selectedKey || "").trim();
  const impianto = key ? currentImpianti.find((item) => buildImpiantoKey(item) === key) : null;
  const label = impianto ? getImpiantoSearchLabel(impianto) : "";
  if (ui.commessaNoteImpiantoKey) ui.commessaNoteImpiantoKey.value = impianto ? key : "";
  if (ui.commessaNoteImpiantoSearch && options.updateSearch !== false) ui.commessaNoteImpiantoSearch.value = label;
  if (ui.commessaNoteImpiantoSelected) {
    ui.commessaNoteImpiantoSelected.textContent = impianto
      ? `Impianto selezionato: ${label}`
      : "Nessun impianto collegato.";
  }
  renderCommessaNoteImpiantoSuggestions();
}

function clearCommessaNoteImpiantoSelection() {
  commessaNoteImpiantoSearchTerm = "";
  if (ui.commessaNoteImpiantoSearch) ui.commessaNoteImpiantoSearch.value = "";
  updateCommessaNoteImpiantoSelection("");
  ui.commessaNoteImpiantoSearch?.focus();
}

function getFilteredCommessaNoteImpianti() {
  const term = commessaNoteImpiantoSearchTerm.trim().toLowerCase();
  return currentImpianti
    .filter((impianto) => {
      const key = buildImpiantoKey(impianto);
      if (!key) return false;
      if (!term) return true;
      return getImpiantoSearchText(impianto).includes(term);
    })
    .sort((a, b) => {
      const aName = getImpiantoDisplayLabel(a).toLowerCase();
      const bName = getImpiantoDisplayLabel(b).toLowerCase();
      const aStarts = term && aName.startsWith(term);
      const bStarts = term && bName.startsWith(term);
      if (aStarts !== bStarts) return aStarts ? -1 : 1;
      return getImpiantoDisplayLabel(a).localeCompare(getImpiantoDisplayLabel(b), "it");
    })
    .slice(0, 12);
}

function renderCommessaNoteImpiantoSuggestions() {
  if (!ui.commessaNoteImpiantoSuggestions) return;
  const suggestions = getFilteredCommessaNoteImpianti();
  ui.commessaNoteImpiantoSuggestions.innerHTML = "";
  const isFocused = document.activeElement === ui.commessaNoteImpiantoSearch;
  if (!isFocused && !commessaNoteImpiantoSearchTerm) {
    ui.commessaNoteImpiantoSuggestions.classList.add("hidden");
    ui.commessaNoteImpiantoSearch?.setAttribute("aria-expanded", "false");
    return;
  }
  if (!suggestions.length) {
    ui.commessaNoteImpiantoSuggestions.innerHTML = "<p class='muted'>Nessun impianto della commessa trovato.</p>";
    ui.commessaNoteImpiantoSuggestions.classList.remove("hidden");
    ui.commessaNoteImpiantoSearch?.setAttribute("aria-expanded", "true");
    return;
  }
  suggestions.forEach((impianto) => {
    const key = buildImpiantoKey(impianto);
    const btn = document.createElement("button");
    btn.type = "button";
    btn.className = "commessa-note-suggestion";
    btn.setAttribute("role", "option");
    btn.innerHTML = `
      <strong>${escapeHTML(getImpiantoDisplayLabel(impianto))}</strong>
      <small>${escapeHTML(getImpiantoSearchLabel(impianto))}</small>
    `;
    btn.addEventListener("mousedown", (event) => event.preventDefault());
    btn.addEventListener("click", () => {
      commessaNoteImpiantoSearchTerm = "";
      updateCommessaNoteImpiantoSelection(key);
      ui.commessaNoteImpiantoSuggestions?.classList.add("hidden");
      ui.commessaNoteImpiantoSearch?.setAttribute("aria-expanded", "false");
    });
    ui.commessaNoteImpiantoSuggestions.appendChild(btn);
  });
  ui.commessaNoteImpiantoSuggestions.classList.remove("hidden");
  ui.commessaNoteImpiantoSearch?.setAttribute("aria-expanded", "true");
}

function onCommessaNoteImpiantoSearchInput(event) {
  commessaNoteImpiantoSearchTerm = String(event.target.value || "").trim();
  if (ui.commessaNoteImpiantoKey) ui.commessaNoteImpiantoKey.value = "";
  if (ui.commessaNoteImpiantoSelected) ui.commessaNoteImpiantoSelected.textContent = "Seleziona un suggerimento oppure lascia nessun impianto collegato.";
  renderCommessaNoteImpiantoSuggestions();
}

function openCommessaNoteForm(note = null) {
  if (!selectedCommessaId) return;
  ui.commessaNoteDetail?.classList.add("hidden");
  if (ui.commessaNoteDetail) ui.commessaNoteDetail.innerHTML = "";
  ui.commessaNotesFormWrap?.classList.remove("hidden");
  if (ui.commessaNotesToggleBtn) ui.commessaNotesToggleBtn.setAttribute("aria-expanded", "true");
  ui.commessaNoteId.value = note?.id || "";
  ui.commessaNoteDate.value = note?.noteDate || getTodayDateKey();
  if (ui.commessaNoteTitle) ui.commessaNoteTitle.value = getCommessaNoteTitle(note);
  ui.commessaNoteText.value = note?.text || "";
  ui.commessaNoteDriveLinks.value = Array.isArray(note?.driveLinks) ? note.driveLinks.join("\n") : "";
  commessaNoteImpiantoSearchTerm = "";
  updateCommessaNoteImpiantoSelection(note?.impiantoKey || "");
  ui.commessaNoteSubmitBtn.textContent = note?.id ? "Aggiorna nota" : "Salva nota";
  setTimeout(() => (note?.id ? ui.commessaNoteTitle : ui.commessaNoteTitle || ui.commessaNoteText)?.focus(), 30);
}

function closeCommessaNoteForm() {
  ui.commessaNotesFormWrap?.classList.add("hidden");
  if (ui.commessaNotesToggleBtn) ui.commessaNotesToggleBtn.setAttribute("aria-expanded", "false");
  ui.commessaNoteForm?.reset();
  if (ui.commessaNoteId) ui.commessaNoteId.value = "";
  if (ui.commessaNoteSubmitBtn) ui.commessaNoteSubmitBtn.textContent = "Salva nota";
}

function parseDriveLinks(value) {
  return String(value || "")
    .split(/\n+/)
    .map((item) => item.trim())
    .filter(Boolean);
}

function getDriveLinkLabel(url) {
  const normalized = String(url || "").toLowerCase();
  if (/\.(png|jpe?g|webp|gif)(\?|#|$)/.test(normalized) || normalized.includes("photo") || normalized.includes("foto")) return "📷 Apri foto";
  if (/\.(pdf|docx?|xlsx?|pptx?)(\?|#|$)/.test(normalized) || normalized.includes("document")) return "📄 Apri documento";
  return "🔗 Apri allegato";
}

function formatCommessaNoteDate(noteDate) {
  if (!noteDate) return "-";
  const date = new Date(`${noteDate}T00:00:00`);
  if (Number.isNaN(date.getTime())) return noteDate;
  return date.toLocaleDateString("it-IT");
}

async function saveCommessaNote(event) {
  event.preventDefault();
  if (!selectedCommessaId) return;
  const noteId = String(ui.commessaNoteId.value || "").trim();
  const impiantoKey = String(ui.commessaNoteImpiantoKey?.value || "").trim();
  const impianto = impiantoKey ? currentImpianti.find((item) => buildImpiantoKey(item) === impiantoKey) : null;
  const payload = {
    commessaId: selectedCommessaId,
    commessaName: selectedCommessaName || "",
    noteDate: ui.commessaNoteDate.value || getTodayDateKey(),
    title: String(ui.commessaNoteTitle?.value || "").trim(),
    text: String(ui.commessaNoteText.value || "").trim(),
    driveLinks: parseDriveLinks(ui.commessaNoteDriveLinks.value),
    impiantoKey,
    impiantoLabel: impianto ? getImpiantoDisplayLabel(impianto) : "",
    updatedAt: firebase.firestore.FieldValue.serverTimestamp(),
    updatedBy: currentUser?.email || ""
  };
  if (!payload.title) {
    alert("Inserisci il titolo della nota.");
    return;
  }
  if (!payload.text) {
    alert("Inserisci il testo della nota.");
    return;
  }
  const notesRef = db.collection("commesse").doc(selectedCommessaId).collection("noteCommessa");
  if (noteId) {
    await notesRef.doc(noteId).set(payload, { merge: true });
  } else {
    await notesRef.add({
      ...payload,
      createdAt: firebase.firestore.FieldValue.serverTimestamp(),
      createdBy: currentUser?.email || ""
    });
  }
  closeCommessaNoteForm();
}

async function updateCommessaNoteStatus(noteId, status) {
  if (!selectedCommessaId || !noteId) return;
  await db.collection("commesse").doc(selectedCommessaId).collection("noteCommessa").doc(noteId).set({
    status,
    updatedAt: firebase.firestore.FieldValue.serverTimestamp(),
    updatedBy: currentUser?.email || ""
  }, { merge: true });
}

async function deleteCommessaNote(noteId) {
  if (!selectedCommessaId || !noteId) return;
  if (!window.confirm("Eliminare questa nota commessa?")) return;
  await db.collection("commesse").doc(selectedCommessaId).collection("noteCommessa").doc(noteId).delete();
}

function openCommessaNoteDetail(note) {
  if (!note || !ui.commessaNoteDetail) return;
  closeCommessaNoteForm();
  const driveLinks = Array.isArray(note.driveLinks) ? note.driveLinks : [];
  ui.commessaNoteDetail.innerHTML = `
    <div class="section-head">
      <div>
        <h3>${escapeHTML(getCommessaNoteTitle(note))}</h3>
        <p class="commessa-note-meta">
          <span>📅 ${escapeHTML(formatCommessaNoteDate(note.noteDate))}</span>
          ${note.impiantoLabel ? `<span>🏭 ${escapeHTML(note.impiantoLabel)}</span>` : ""}
        </p>
      </div>
      <button class="btn" type="button" data-note-close>Chiudi dettaglio</button>
    </div>
    <p class="commessa-note-text">${escapeHTML(note.text || "-")}</p>
    <div class="commessa-note-attachments">
      ${driveLinks.length ? driveLinks.map((link) => `<a class="btn btn-small commessa-note-attachment" href="${escapeHTML(link)}" target="_blank" rel="noopener noreferrer">${escapeHTML(getDriveLinkLabel(link))}</a>`).join("") : "<span class='muted'>Nessun link Google Drive.</span>"}
    </div>
    <div class="commessa-note-actions">
      <button class="btn" type="button" data-note-edit>Modifica</button>
      <button class="btn" type="button" data-note-delete>Elimina</button>
    </div>
  `;
  ui.commessaNoteDetail.querySelector("[data-note-close]")?.addEventListener("click", () => ui.commessaNoteDetail.classList.add("hidden"));
  ui.commessaNoteDetail.querySelector("[data-note-edit]")?.addEventListener("click", () => openCommessaNoteForm(note));
  ui.commessaNoteDetail.querySelector("[data-note-delete]")?.addEventListener("click", () => deleteCommessaNote(note.id));
  ui.commessaNoteDetail.classList.remove("hidden");
  ui.commessaNoteDetail.scrollIntoView({ behavior: "smooth", block: "start" });
}

function renderCommessaNotes() {
  if (!ui.commessaNotesList || !ui.commessaNotesCounter) return;
  const total = currentCommessaNotes.length;
  ui.commessaNotesCounter.textContent = `📝 Note: ${total}`;
  if (ui.commessaNotesTitle) ui.commessaNotesTitle.textContent = selectedCommessaName ? `Commessa: ${selectedCommessaName}` : "";

  if (!selectedCommessaId) {
    ui.commessaNotesList.innerHTML = "<p class='muted'>Seleziona una commessa per vedere le note.</p>";
    return;
  }
  if (!currentCommessaNotes.length) {
    ui.commessaNotesList.innerHTML = "<p class='muted'>Nessuna nota inserita per questa commessa.</p>";
    return;
  }

  ui.commessaNotesList.innerHTML = "";
  currentCommessaNotes.forEach((note) => {
    const article = document.createElement("article");
    article.className = "commessa-note-item";
    const titleButton = document.createElement("button");
    titleButton.type = "button";
    titleButton.className = "commessa-note-title-btn";
    titleButton.textContent = getCommessaNoteTitle(note);
    titleButton.addEventListener("click", () => openCommessaNoteDetail(note));
    const meta = document.createElement("div");
    meta.className = "commessa-note-meta";
    meta.innerHTML = `
      <span>📅 ${escapeHTML(formatCommessaNoteDate(note.noteDate))}</span>
      ${note.impiantoLabel ? `<span>🏭 ${escapeHTML(note.impiantoLabel)}</span>` : ""}
    `;
    article.appendChild(titleButton);
    article.appendChild(meta);
    ui.commessaNotesList.appendChild(article);
  });
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
      const rows = rowsFromWorkbookBuffer(e.target.result);
      const normalizedRows = rows.map(normalizeRow).filter((row) => row.denominazione || row.idSap || row.indirizzo);
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

async function importFromGoogleSheetUrl() {
  pendingRows = [];
  ui.importBtn.disabled = true;
  const value = String(ui.sheetUrl?.value || "").trim();
  if (!value) {
    ui.importFeedback.textContent = "Inserisci URL Google Sheet.";
    return;
  }
  try {
    const rows = await fetchGoogleSheetRows(value);
    const normalizedRows = rows.map(normalizeRow).filter((row) => row.denominazione || row.idSap || row.indirizzo);
    pendingRows = mergeRowsByImpianto(normalizedRows);
    ui.importFeedback.textContent = `Google Sheet letto: ${normalizedRows.length} righe. Impianti unici: ${pendingRows.length}`;
    ui.importBtn.disabled = !auth.currentUser || !getTargetCommessaId() || pendingRows.length === 0 || !canManageData();
  } catch (error) {
    console.error("Errore import Google Sheet:", error);
    ui.importFeedback.textContent = "Errore lettura Google Sheet. Verifica link/condivisione.";
  }
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
    if (merged.areaMq == null && data.areaMq != null) merged.areaMq = data.areaMq;
    if (merged.sfalciMq == null && data.sfalciMq != null) merged.sfalciMq = data.sfalciMq;
    if (merged.sfalciVerdiMq == null && data.sfalciVerdiMq != null) merged.sfalciVerdiMq = data.sfalciVerdiMq;
    if (merged.potaturaSiepiM == null && data.potaturaSiepiM != null) merged.potaturaSiepiM = data.potaturaSiepiM;
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
    const mergedAreaMq = row.areaMq != null ? row.areaMq : (existing.areaMq ?? null);
    const mergedSfalciMq = row.sfalciMq != null ? row.sfalciMq : (existing.sfalciMq ?? null);
    const mergedSfalciVerdiMq = row.sfalciVerdiMq != null ? row.sfalciVerdiMq : (existing.sfalciVerdiMq ?? null);
    const mergedPotaturaSiepiM = row.potaturaSiepiM != null ? row.potaturaSiepiM : (existing.potaturaSiepiM ?? null);
    const mergedExtraFields = mergeExtraFields(existing.extraFields, row.extraFields);
    const extraFieldsChanged = JSON.stringify(mergedExtraFields || {}) !== JSON.stringify(existing.extraFields || {});
    const changed = mergedCodicePrezzo !== String(existing.codicePrezzo || "")
      || mergedVoceRiferimento !== String(existing.voceRiferimento || "")
      || mergedTipologiaIntervento !== String(existing.tipologiaIntervento || "")
      || mergedLavorazioniRichieste !== String(existing.lavorazioniRichieste || "")
      || mergedFrequenzaAnnua !== String(existing.frequenzaAnnua || "")
      || mergedAreaMq !== (existing.areaMq ?? null)
      || mergedSfalciMq !== (existing.sfalciMq ?? null)
      || mergedSfalciVerdiMq !== (existing.sfalciVerdiMq ?? null)
      || mergedPotaturaSiepiM !== (existing.potaturaSiepiM ?? null)
      || extraFieldsChanged;
    if (!changed) return;
    rowsToUpdate.push({
      id: existing.id,
      codicePrezzo: mergedCodicePrezzo,
      voceRiferimento: mergedVoceRiferimento,
      tipologiaIntervento: mergedTipologiaIntervento,
      lavorazioniRichieste: mergedLavorazioniRichieste,
      frequenzaAnnua: mergedFrequenzaAnnua,
      areaMq: mergedAreaMq,
      sfalciMq: mergedSfalciMq,
      sfalciVerdiMq: mergedSfalciVerdiMq,
      potaturaSiepiM: mergedPotaturaSiepiM,
      extraFields: mergedExtraFields
    });
    existing.codicePrezzo = mergedCodicePrezzo;
    existing.voceRiferimento = mergedVoceRiferimento;
    existing.tipologiaIntervento = mergedTipologiaIntervento;
    existing.lavorazioniRichieste = mergedLavorazioniRichieste;
    existing.frequenzaAnnua = mergedFrequenzaAnnua;
    existing.areaMq = mergedAreaMq;
    existing.sfalciMq = mergedSfalciMq;
    existing.sfalciVerdiMq = mergedSfalciVerdiMq;
    existing.potaturaSiepiM = mergedPotaturaSiepiM;
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
          areaMq: row.areaMq ?? null,
          sfalciMq: row.sfalciMq ?? null,
          sfalciVerdiMq: row.sfalciVerdiMq ?? null,
          potaturaSiepiM: row.potaturaSiepiM ?? null,
          extraFields: row.extraFields || {},
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
    areaMq: null,
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
  const consumedKeys = new Set();

  const idSapEntry = getValueWithMatchedKey(keys, ["idsap", "id", "codicehera", "codiceimpianto", "codicecliente", "code"]);
  if (idSapEntry.key) consumedKeys.add(idSapEntry.key);
  const denominazioneEntry = getValueWithMatchedKey(keys, ["denominazioneimpianto", "denominazione", "impianto", "impiantounico", "sito", "sitonome", "nomeimpianto"]);
  if (denominazioneEntry.key) consumedKeys.add(denominazioneEntry.key);
  const comuneEntry = getValueWithMatchedKey(keys, ["comuneubicazioneimpianto", "comuneubicazione", "comune", "citta", "city"]);
  if (comuneEntry.key) consumedKeys.add(comuneEntry.key);
  const indirizzoEntry = getValueWithMatchedKey(keys, ["descrizionevia", "viaecivicodiubicazioneimpianto", "indirizzoubicazione", "indirizzo", "via", "address"]);
  if (indirizzoEntry.key) consumedKeys.add(indirizzoEntry.key);
  const tipologiaImpiantoEntry = getValueWithMatchedKey(keys, ["tipologiaimpianto", "tipoimpianto"]);
  if (tipologiaImpiantoEntry.key) consumedKeys.add(tipologiaImpiantoEntry.key);
  const sfalciEntry = getExactValueWithMatchedKey(keys, [
    "sfalciverdimq",
    "sfalciareeverdimq",
    "sfalciareeverdimqpotaturasiepim"
  ]);
  const potaturaSiepiEntry = getExactValueWithMatchedKey(keys, [
    "potaturasiepim",
    "potaturasiepi"
  ]);
  const fallbackMqEntry = getExactValueWithMatchedKey(keys, [
    "superficiemq",
    "mq",
    "superficie"
  ]);
  const mqEntry = sfalciEntry.value ? sfalciEntry : (potaturaSiepiEntry.value ? potaturaSiepiEntry : fallbackMqEntry);
  const parsedSfalciMq = parseAreaMqValue(sfalciEntry.value);
  const parsedPotaturaSiepiM = parseAreaMqValue(potaturaSiepiEntry.value);
  const parsedAreaMq = parseAreaMqValue(mqEntry.value);
  if (sfalciEntry.key && parsedSfalciMq != null) consumedKeys.add(sfalciEntry.key);
  if (potaturaSiepiEntry.key && parsedPotaturaSiepiM != null) consumedKeys.add(potaturaSiepiEntry.key);
  if (fallbackMqEntry.key && parsedAreaMq != null) consumedKeys.add(fallbackMqEntry.key);
  const areaEntry = getValueWithMatchedKey(keys, ["area", "competenza"]);
  if (areaEntry.key && (areaEntry.key !== mqEntry.key || parsedAreaMq == null)) consumedKeys.add(areaEntry.key);
  const dittaEsecutriceEntry = getValueWithMatchedKey(keys, ["dittaesecutrice", "ditaesecutrice", "dittaappaltatrice", "ditta"]);
  if (dittaEsecutriceEntry.key) consumedKeys.add(dittaEsecutriceEntry.key);
  const voceEntry = getValueWithMatchedKey(keys, ["vocediriferimentoelencoprezzi", "voce", "riferimento", "codiceintervento"]);
  if (voceEntry.key) consumedKeys.add(voceEntry.key);
  const codicePrezzoEntry = getValueWithMatchedKey(keys, ["vocediriferimentoelencoprezzi", "codiceprezzo", "prezzo", "codice"]);
  if (codicePrezzoEntry.key) consumedKeys.add(codicePrezzoEntry.key);
  const gpsYEntry = getValueWithMatchedKey(keys, ["coordinategpsy", "gpslat", "latitudine", "latitude"]);
  if (gpsYEntry.key) consumedKeys.add(gpsYEntry.key);
  const gpsXEntry = getValueWithMatchedKey(keys, ["coordinategpsx", "gpslong", "gpslng", "longitudine", "longitude"]);
  if (gpsXEntry.key) consumedKeys.add(gpsXEntry.key);
  const coordinatesEntry = getValueWithMatchedKey(keys, ["coordinate", "coord"]);
  if (coordinatesEntry.key) consumedKeys.add(coordinatesEntry.key);
  const coordinatePair = parseCoordinatePair(coordinatesEntry.value);

  const extraFields = {};
  Object.entries(keys).forEach(([key, value]) => {
    if (!value || consumedKeys.has(key)) return;
    extraFields[key] = value;
  });

  return {
    distretto: getValue(keys, ["distretto"]),
    idSap: idSapEntry.value,
    denominazione: denominazioneEntry.value,
    comune: comuneEntry.value,
    indirizzo: indirizzoEntry.value,
    descrizioneVia: indirizzoEntry.value,
    tipologiaImpianto: tipologiaImpiantoEntry.value,
    area: areaEntry.key === mqEntry.key && parsedAreaMq != null ? "" : areaEntry.value,
    competenza: areaEntry.key === mqEntry.key && parsedAreaMq != null ? "" : areaEntry.value,
    areaMq: parsedAreaMq,
    sfalciMq: parsedSfalciMq,
    sfalciVerdiMq: parsedSfalciMq,
    potaturaSiepiM: parsedPotaturaSiepiM,
    dittaEsecutrice: dittaEsecutriceEntry.value,
    voceRiferimento: voceEntry.value,
    codicePrezzo: codicePrezzoEntry.value,
    sfalci: parsedSfalciMq == null ? "" : sfalciEntry.value,
    potaturaSiepi: parsedPotaturaSiepiM == null ? "" : potaturaSiepiEntry.value,
    frequenzaAnnua: getValue(keys, ["frequenzaannuaminimasfalcieopotaturasiepin", "frequenzaindicativanvolteanno"]),
    tipologiaIntervento: getValue(keys, ["tipologiadisfalciointervento", "lavorazionirichieste", "tipoimpianto"]) || tipologiaImpiantoEntry.value,
    lavorazioniRichieste: getValue(keys, ["lavorazionirichieste", "tipologiadisfalciointervento"]),
    gpsY: parseCoordinate(gpsYEntry.value) ?? coordinatePair.lat,
    gpsX: parseCoordinate(gpsXEntry.value) ?? coordinatePair.lon,
    extraFields
  };
}

function prepareGlobalRowsForImport(rawRows) {
  const normalizedRows = rawRows.map(normalizeRow);
  const validRows = normalizedRows.filter((row) => {
    const hasIdentity = Boolean(row.denominazione || row.idSap || row.indirizzo || row.descrizioneVia);
    const hasDitta = Boolean(String(row.dittaEsecutrice || "").trim());
    return hasIdentity && hasDitta;
  });
  return {
    rows: mergeRowsByImpianto(validRows),
    validRows: validRows.length,
    invalidRows: Math.max(0, normalizedRows.length - validRows.length)
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
    if (existing.areaMq == null && row.areaMq != null) existing.areaMq = row.areaMq;

    if (!existing.distretto && row.distretto) existing.distretto = row.distretto;
    if (!existing.comune && row.comune) existing.comune = row.comune;
    if (!existing.indirizzo && row.indirizzo) existing.indirizzo = row.indirizzo;
    if (!existing.descrizioneVia && row.descrizioneVia) existing.descrizioneVia = row.descrizioneVia;
    if (!existing.area && row.area) existing.area = row.area;
    if (!existing.competenza && row.competenza) existing.competenza = row.competenza;
    if (!existing.tipologiaImpianto && row.tipologiaImpianto) existing.tipologiaImpianto = row.tipologiaImpianto;
    if (!existing.dittaEsecutrice && row.dittaEsecutrice) existing.dittaEsecutrice = row.dittaEsecutrice;
    if (!existing.idSap && row.idSap) existing.idSap = row.idSap;
    if (existing.gpsY == null && row.gpsY != null) existing.gpsY = row.gpsY;
    if (existing.gpsX == null && row.gpsX != null) existing.gpsX = row.gpsX;
    existing.extraFields = mergeExtraFields(existing.extraFields, row.extraFields);
  });

  return Array.from(grouped.values());
}


function isImpiantoDoneState(impianto) {
  if (!impianto || typeof impianto !== "object") return false;
  if (Boolean(impianto.done) || Boolean(impianto.fatto) || Boolean(impianto.completed)) return true;
  const stato = String(impianto.stato || impianto.status || "").trim().toLowerCase();
  if (["fatto", "done", "completed", "completato"].includes(stato)) return true;
  const doneValue = String(impianto.done || "").trim().toLowerCase();
  return ["true", "1", "fatto", "done", "completed"].includes(doneValue);
}

function combineImpiantiForView(impianti) {
  const grouped = new Map();

  impianti.forEach((item) => {
    const key = buildImpiantoKey(item);
    const existing = grouped.get(key);
    if (!existing) {
      const doneAtMs = firestoreDateToMillis(item.doneAt);
      const resetAtMs = firestoreDateToMillis(item.resetAt);
      grouped.set(key, {
        ...item,
        done: doneAtMs >= resetAtMs && isImpiantoDoneState(item),
        doneAt: item.doneAt || null,
        doneBy: doneAtMs >= resetAtMs ? (item.doneBy || "") : "",
        resetAt: item.resetAt || null,
        resetBy: item.resetBy || "",
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
    if (existing.areaMq == null && item.areaMq != null) existing.areaMq = item.areaMq;
    if (existing.sfalciMq == null && item.sfalciMq != null) existing.sfalciMq = item.sfalciMq;
    if (existing.sfalciVerdiMq == null && item.sfalciVerdiMq != null) existing.sfalciVerdiMq = item.sfalciVerdiMq;
    if (existing.potaturaSiepiM == null && item.potaturaSiepiM != null) existing.potaturaSiepiM = item.potaturaSiepiM;

    existing.hasOrdinario = hasOrdinario(existing.codicePrezzo);
    existing.hasStraordinario = hasStraordinario(existing.codicePrezzo);
    existing.tipoManutenzione = classifyTipoManutenzione(existing.codicePrezzo);
    const itemDone = isImpiantoDoneState(item);
    const existingDoneAtMs = firestoreDateToMillis(existing.doneAt);
    const itemDoneAtMs = firestoreDateToMillis(item.doneAt);
    const existingResetAtMs = firestoreDateToMillis(existing.resetAt);
    const itemResetAtMs = firestoreDateToMillis(item.resetAt);

    if (itemDone && (!existing.doneBy || itemDoneAtMs >= existingDoneAtMs)) {
      existing.doneBy = item.doneBy || existing.doneBy || "";
    }
    if (itemDoneAtMs >= existingDoneAtMs) {
      existing.doneAt = item.doneAt || existing.doneAt || null;
    }
    if (itemResetAtMs >= existingResetAtMs) {
      existing.resetAt = item.resetAt || existing.resetAt || null;
      existing.resetBy = item.resetBy || existing.resetBy || "";
    }
    const latestDoneMs = firestoreDateToMillis(existing.doneAt);
    const latestResetMs = firestoreDateToMillis(existing.resetAt);
    existing.done = latestDoneMs >= latestResetMs;
    if (!existing.done) {
      existing.doneBy = "";
    }
    if (!existing.idSap && item.idSap) existing.idSap = item.idSap;
    if (!existing.comune && item.comune) existing.comune = item.comune;
    if (!existing.indirizzo && item.indirizzo) existing.indirizzo = item.indirizzo;
    if (!existing.descrizioneVia && item.descrizioneVia) existing.descrizioneVia = item.descrizioneVia;
    if (!existing.area && item.area) existing.area = item.area;
    if (!existing.competenza && item.competenza) existing.competenza = item.competenza;
    if (!existing.tipologiaImpianto && item.tipologiaImpianto) existing.tipologiaImpianto = item.tipologiaImpianto;
    if (!existing.dittaEsecutrice && item.dittaEsecutrice) existing.dittaEsecutrice = item.dittaEsecutrice;
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
  return getValueWithMatchedKey(obj, aliases).value;
}

function getValueWithMatchedKey(obj, aliases) {
  for (const alias of aliases) {
    if (obj[alias]) return { value: obj[alias], key: alias };
  }
  const keys = Object.keys(obj);
  for (const alias of aliases) {
    const matchedKey = keys.find((key) => key.includes(alias) && obj[key]);
    if (matchedKey) return { value: obj[matchedKey], key: matchedKey };
  }
  return { value: "", key: "" };
}

function getExactValueWithMatchedKey(obj, aliases) {
  for (const alias of aliases) {
    if (obj[alias]) return { value: obj[alias], key: alias };
  }
  return { value: "", key: "" };
}

function mergeExtraFields(oldFields, newFields) {
  const current = oldFields && typeof oldFields === "object" ? oldFields : {};
  const incoming = newFields && typeof newFields === "object" ? newFields : {};
  const merged = { ...current };
  Object.entries(incoming).forEach(([key, value]) => {
    const oldValue = String(merged[key] || "").trim();
    const newValue = String(value || "").trim();
    if (!newValue) return;
    if (!oldValue) {
      merged[key] = newValue;
      return;
    }
    merged[key] = mergeMultiValue(oldValue, newValue);
  });
  return merged;
}

function rowsFromWorkbookBuffer(arrayBuffer) {
  const workbook = XLSX.read(arrayBuffer, { type: "array" });
  const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
  return XLSX.utils.sheet_to_json(firstSheet, { defval: "", raw: false });
}

async function fetchGoogleSheetRows(sheetUrl) {
  const exportUrl = buildGoogleSheetCsvUrl(sheetUrl);
  const response = await fetchWithTimeoutAndRetry(exportUrl, {}, {
    timeoutMs: NETWORK_DEFAULT_TIMEOUT_MS,
    retries: 2
  });
  if (!response.ok) throw new Error(`HTTP ${response.status}`);
  const csvText = await response.text();
  const workbook = XLSX.read(csvText, { type: "string" });
  const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
  return XLSX.utils.sheet_to_json(firstSheet, { defval: "", raw: false });
}

function buildGoogleSheetCsvUrl(inputUrl) {
  const url = new URL(inputUrl);
  if (!/docs\.google\.com$/i.test(url.hostname)) return inputUrl;
  const match = url.pathname.match(/\/spreadsheets\/d\/([^/]+)/i);
  if (!match || !match[1]) return inputUrl;
  const sheetId = match[1];
  const gid = url.searchParams.get("gid") || "0";
  return `https://docs.google.com/spreadsheets/d/${sheetId}/export?format=csv&gid=${gid}`;
}

function parseCoordinate(value) {
  const normalized = String(value || "").replace(",", ".").trim();
  const parsed = Number.parseFloat(normalized);
  return Number.isFinite(parsed) ? parsed : null;
}

function parseAreaMqValue(value) {
  const raw = String(value ?? "").trim();
  if (!raw) return null;
  let normalized = raw
    .replace(/\s+/g, "")
    .replace(/[^0-9,.-]/g, "")
    .replace(/(?!^)-/g, "");
  if (!/[0-9]/.test(normalized)) return null;
  const hasComma = normalized.includes(",");
  const hasDot = normalized.includes(".");
  if (hasComma && hasDot) {
    const decimalSeparator = normalized.lastIndexOf(",") > normalized.lastIndexOf(".") ? "," : ".";
    const thousandsSeparator = decimalSeparator === "," ? "." : ",";
    normalized = normalized.split(thousandsSeparator).join("").replace(decimalSeparator, ".");
  } else if (hasComma) {
    const commaParts = normalized.split(",");
    const lastPart = commaParts[commaParts.length - 1] || "";
    normalized = commaParts.length > 2 || lastPart.length === 3
      ? commaParts.join("")
      : normalized.replace(",", ".");
  } else if (hasDot) {
    const dotParts = normalized.split(".");
    const lastPart = dotParts[dotParts.length - 1] || "";
    normalized = dotParts.length > 2 || lastPart.length === 3
      ? dotParts.join("")
      : normalized;
  }
  const parsed = Number.parseFloat(normalized);
  return Number.isFinite(parsed) ? parsed : null;
}

function formatAreaMqValue(value) {
  const parsed = typeof value === "number" ? value : parseAreaMqValue(value);
  if (!Number.isFinite(parsed)) return "--";
  const rounded = Number(parsed.toFixed(2));
  const sign = rounded < 0 ? "-" : "";
  const [integerPart, decimalPart = ""] = String(Math.abs(rounded)).split(".");
  const formattedInteger = integerPart.replace(/\B(?=(\d{3})+(?!\d))/g, ".");
  const formattedDecimal = decimalPart.replace(/0+$/, "");
  return formattedDecimal ? `${sign}${formattedInteger},${formattedDecimal}` : `${sign}${formattedInteger}`;
}

function parseCoordinatePair(value) {
  const raw = String(value || "").trim();
  if (!raw) return { lat: null, lon: null };
  const parts = raw.split(/[;,/\s]+/).map((part) => part.trim()).filter(Boolean);
  if (parts.length < 2) return { lat: null, lon: null };
  const lat = parseCoordinate(parts[0]);
  const lon = parseCoordinate(parts[1]);
  if (!isValidLatLon(lat, lon)) return { lat: null, lon: null };
  return { lat, lon };
}

function isValidLatLon(lat, lon) {
  return Number.isFinite(lat) && Number.isFinite(lon) && Math.abs(lat) <= 90 && Math.abs(lon) <= 180;
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
  if (Number.isFinite(value?.seconds)) return (value.seconds * 1000) + Math.floor(Number(value.nanoseconds || 0) / 1000000);
  if (Number.isFinite(value?._seconds)) return (value._seconds * 1000) + Math.floor(Number(value._nanoseconds || 0) / 1000000);
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

async function fetchDoneImpiantiRowsForExport(commessaId, commessaName, parentName = "") {
  const snapshot = await db
    .collection("commesse")
    .doc(commessaId)
    .collection("impianti")
    .get();

  const rawImpianti = snapshot.docs.map((doc) => ({ id: doc.id, ...doc.data() }));
  const doneImpianti = combineImpiantiForView(rawImpianti).filter((impianto) => impianto.done);

  return doneImpianti.flatMap((impianto) => buildRowsForEachCodicePrezzo(impianto)).map((impianto) => {
    const doneInfo = formatDoneDateTime(impianto.doneAt);
    return {
      "Commessa padre": parentName || "",
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
}

async function exportCommessaSummary(commessaId, commessaName) {
  if (!auth.currentUser) {
    alert("Devi fare login per esportare il riepilogo.");
    return;
  }

  try {
    const selected = commesseById.get(commessaId) || { id: commessaId, nome: commessaName };
    const subcommesse = getSubcommesse(commessaId);
    const includeSubcommesse = subcommesse.length > 0 && window.confirm(
      `La commessa "${commessaName}" ha ${subcommesse.length} subcommesse.\n\nPremi OK per esportare una vista raggruppata con commessa padre e subcommesse, oppure Annulla per esportare solo la commessa padre.`
    );
    const exportCommesse = includeSubcommesse ? [selected, ...sortCommesseByCreatedAtDesc(subcommesse)] : [selected];
    const parentName = includeSubcommesse ? (selected.nome || commessaName || "Commessa padre") : "";
    const rowsGroups = await Promise.all(exportCommesse.map((commessa) => (
      fetchDoneImpiantiRowsForExport(
        commessa.id,
        commessa.nome || (commessa.id === commessaId ? commessaName : "Commessa"),
        includeSubcommesse && commessa.id !== selected.id ? parentName : ""
      )
    )));
    const rows = rowsGroups.flat();

    if (!rows.length) {
      const message = includeSubcommesse
        ? `Nessun impianto FATTO da esportare per la commessa "${commessaName}" e le sue subcommesse.`
        : `Nessun impianto FATTO da esportare per la commessa "${commessaName}".`;
      alert(message);
      return;
    }

    const worksheet = XLSX.utils.json_to_sheet(rows);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, includeSubcommesse ? "Riepilogo raggruppato" : "Riepilogo impianti");

    const safeName = String(commessaName || "commessa")
      .trim()
      .replace(/[\\/:*?"<>|]/g, "_")
      .replace(/\s+/g, "_");
    const suffix = includeSubcommesse ? "_raggruppato" : "";
    const timestamp = new Date().toISOString().slice(0, 19).replace(/[:T]/g, "-");
    XLSX.writeFile(workbook, `riepilogo_impianti_${safeName}${suffix}_${timestamp}.xlsx`);
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
  impiantiViewMode = ["todo", "alerts"].includes(mode) ? mode : "done";
  ui.viewDoneBtn.classList.toggle("btn-primary", impiantiViewMode === "done");
  ui.viewTodoBtn.classList.toggle("btn-primary", impiantiViewMode === "todo");
  ui.viewAlertsBtn?.classList.toggle("btn-primary", impiantiViewMode === "alerts");
  renderImpianti();
}

function getFirstParsedAreaValue(values = []) {
  for (const value of values) {
    const parsed = parseAreaMqValue(value);
    if (parsed != null) return parsed;
  }
  return null;
}

function getPlantSfalciMq(plant = {}) {
  const importedSfalciMq = getFirstParsedAreaValue([plant.sfalciMq, plant.sfalciVerdiMq, plant.sfalciVerdi, plant.sfalci]);
  if (importedSfalciMq != null) return importedSfalciMq;
  const hasPotaturaValue = getFirstParsedAreaValue([plant.potaturaSiepiM, plant.potaturaSiepi]) != null;
  return hasPotaturaValue ? null : getFirstParsedAreaValue([plant.areaMq]);
}

function getPlantMq(plant = {}) {
  return getFirstParsedAreaValue([
    plant.sfalciMq,
    plant.sfalciVerdiMq,
    plant.sfalciVerdi,
    plant.potaturaSiepiM,
    plant.potaturaSiepi,
    plant.areaMq,
    plant.sfalci,
    plant.mq,
    plant.superficieMq,
    plant.superficie_mq,
    plant.area_mq
  ]);
}

function calculateImpiantiMqProgress(impianti = []) {
  const totaleMqPrevisti = impianti.reduce((sum, impianto) => sum + Number(getPlantSfalciMq(impianto) || 0), 0);
  const totaleMqEseguiti = impianti.reduce((sum, impianto) => {
    if (!impianto.done) return sum;
    return sum + Number(getPlantSfalciMq(impianto) || 0);
  }, 0);
  const avanzamentoMq = totaleMqPrevisti > 0
    ? Math.round((totaleMqEseguiti / totaleMqPrevisti) * 100)
    : 0;
  const mqRimanenti = Math.max(totaleMqPrevisti - totaleMqEseguiti, 0);
  return { totaleMqPrevisti, totaleMqEseguiti, mqRimanenti, avanzamentoMq };
}

function createPlantMqBox(plant) {
  const mqValue = getPlantMq(plant);
  if (mqValue == null) return null;
  const box = document.createElement("div");
  box.className = "area-mq-box mq-chip";
  box.title = "Superficie impianto";

  const icon = document.createElement("span");
  icon.className = "area-mq-icon";
  icon.setAttribute("aria-hidden", "true");
  icon.textContent = "📐";

  const value = document.createElement("span");
  value.className = "area-mq-value";
  value.textContent = `${formatAreaMqValue(mqValue)} mq`;

  box.appendChild(icon);
  box.appendChild(value);
  return box;
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
    const linkedNotes = getCommessaNoteLinkedNotes(impianto);
    const viewMatch = impiantiViewMode === "done"
      ? Boolean(impianto.done)
      : !impianto.done && (impiantiViewMode === "alerts" ? linkedNotes.length > 0 : true);
    return viewMatch && matchesImpiantoSearch(impianto);
  });
  const sorted = [...filtered].sort((a, b) => {
    if (Boolean(a.done) !== Boolean(b.done)) return a.done ? 1 : -1;
    return distanceFromUser(a) - distanceFromUser(b);
  });

  if (!sorted.length) {
    ui.impiantiLista.innerHTML = impiantiViewMode === "alerts"
      ? "<p class='muted'>Nessun impianto da fare con segnalazione collegata.</p>"
      : "<p class='muted'>Nessun impianto trovato con i filtri correnti.</p>";
    renderNextActionCard();
    return;
  }

  sorted.forEach((impianto) => {
    const article = document.createElement("article");
    article.className = `impianto-item card-impianto ${impianto.done ? "done" : "todo"}`;
    const impiantoKey = buildImpiantoKey(impianto);
    const detailsVisible = expandedImpiantoKey === impiantoKey;
    const pendingAction = getPendingActionForImpianto(selectedCommessaId, impianto);
    const whazzupSafetyState = getWhazzupSafetyState(impianto);
    const showWhazzupRecovery = !impianto.done && Boolean(whazzupSafetyState?.needsManualMove);
    const waitingForSync = isActionWaitingForSync(pendingAction);
    article.dataset.impiantoKey = impiantoKey;
    article.classList.toggle("is-expanded", detailsVisible);
    if (highlightedImpiantoKey === impiantoKey) article.classList.add("highlight");

    const linkedNotes = getCommessaNoteLinkedNotes(impianto);
    article.classList.toggle("has-segnalazione", linkedNotes.length > 0);
    article.classList.toggle("whazzup-recovery-warning", showWhazzupRecovery);
    const distanceKm = distanceFromUser(impianto);
    const distance = formatDistance(distanceKm);
    const travelMeta = estimateTravelMeta(distanceKm);
    const tipo = impianto.tipoManutenzione || classifyTipoManutenzione(impianto.codicePrezzo);
    const hasStraordinariaFlag = impianto.hasStraordinario ?? hasStraordinario(impianto.codicePrezzo);
    const badgeTipo = hasStraordinariaFlag ? (tipo || "Straordinaria") : "Ordinaria";
    const mainColumn = document.createElement("div");
    mainColumn.className = "impianto-main-column impianto-left";
    const weatherColumn = document.createElement("div");
    weatherColumn.className = "impianto-weather-column weather-compact";
    weatherColumn.innerHTML = buildImpiantoWeatherBadgeMarkup(impianto);
    weatherColumn.addEventListener("click", (event) => {
      if (event.target?.closest?.("[data-weather-retry], [data-atex-procedure], [data-impianto-safety]")) return;
      openDettaglioMeteoImpianto(impianto);
    });
    weatherColumn.addEventListener("keydown", (event) => {
      if (event.target?.closest?.("[data-weather-retry], [data-atex-procedure], [data-impianto-safety]")) return;
      if (event.key !== "Enter" && event.key !== " ") return;
      event.preventDefault();
      openDettaglioMeteoImpianto(impianto);
    });

    const markerNumber = getMapMarkerNumberForImpianto(impianto);
    const markerChipMarkup = Number.isFinite(markerNumber)
      ? `<span class="impianto-marker-chip ${escapeHTML(getMarkerClass(impianto))}" aria-label="Impianto numero ${escapeHTML(String(markerNumber))}">${escapeHTML(String(markerNumber))}</span>`
      : "";
    const header = document.createElement("button");
    header.type = "button";
    header.className = "impianto-summary-btn";
    header.innerHTML = `
      <span class="impianto-summary-topline">
        <span class="impianto-summary-title-wrap">${markerChipMarkup}<strong>${escapeHTML(impianto.denominazione || "(senza nome)")}</strong></span>
      </span>
      <small class="impianto-summary-meta">
        <span class="badge ${hasStraordinariaFlag ? "badge-straordinaria" : "badge-ordinaria"}">${escapeHTML(badgeTipo)}</span>
        ${linkedNotes.length ? `<span class="badge badge-segnalazione">⚠️ Segnalazione</span>` : ""}
        ${pendingAction ? `<span class="badge badge-whatsapp-pending">WhatsApp in attesa</span>` : ""}
        <span>${distance}</span><span aria-hidden="true">•</span><span class="traffic-level traffic-${travelMeta.intensityKey}">${travelMeta.intensityLabel}</span><span aria-hidden="true">•</span><span>ETA ${travelMeta.etaLabel}</span>
      </small>
    `;
    header.setAttribute("aria-expanded", detailsVisible ? "true" : "false");
    header.addEventListener("click", () => {
      expandedImpiantoKey = expandedImpiantoKey === impiantoKey ? "" : impiantoKey;
      renderImpianti();
    });
    const summaryRow = document.createElement("div");
    summaryRow.className = "impianto-summary-row";
    summaryRow.appendChild(header);
    summaryRow.appendChild(weatherColumn);
    mainColumn.appendChild(summaryRow);

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
    if (linkedNotes.length) {
      const notesBox = document.createElement("div");
      notesBox.className = "impianto-linked-notes";
      notesBox.innerHTML = "<p><b>⚠️ Segnalazione presente:</b></p>";
      linkedNotes.forEach((note) => {
        const noteBtn = document.createElement("button");
        noteBtn.type = "button";
        noteBtn.className = "commessa-note-title-btn impianto-note-title-btn";
        noteBtn.textContent = getCommessaNoteTitle(note);
        noteBtn.addEventListener("click", () => {
          openCommessaNotesPage();
          setTimeout(() => openCommessaNoteDetail(note), 50);
        });
        notesBox.appendChild(noteBtn);
      });
      details.appendChild(notesBox);
    }
    if (showWhazzupRecovery) {
      const warningBox = document.createElement("div");
      warningBox.className = "impianto-whazzup-recovery";
      warningBox.innerHTML = `
        <p><b>⚠️ Da confermare</b> — WhatsApp aperto, ma il salvataggio FATTO non è stato confermato.</p>
        <button type="button" class="btn btn-small">Conferma FATTO</button>
      `;
      const moveBtn = warningBox.querySelector("button");
      moveBtn?.addEventListener("click", async () => {
        moveBtn.disabled = true;
        await forceMoveImpiantoToFatti(impianto);
      });
      details.appendChild(warningBox);
    }
    mainColumn.appendChild(details);

    if (!impianto.done) {
      clearActionUsed(`${selectedCommessaId}:${impiantoKey}:navigate`);
      clearActionUsed(`${selectedCommessaId}:${impiantoKey}:done`);
      clearActionUsed(`${selectedCommessaId}:${impiantoKey}:whatsapp`);
    }

    const actions = document.createElement("div");
    actions.className = "item-actions impianto-actions";
    const primaryActionsRow = document.createElement("div");
    primaryActionsRow.className = "impianto-primary-actions";
    const secondaryActionsRow = document.createElement("div");
    secondaryActionsRow.className = "impianto-secondary-actions";
    const managementStack = document.createElement("div");
    managementStack.className = "impianto-management-stack";
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
      targetContainer = primaryActionsRow
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

    addAction("navigate", "🗺️", "Naviga", () => navigateToImpianto(impianto), false, false, primaryActionsRow);
    const safetyQuickBtn = document.createElement("button");
    safetyQuickBtn.type = "button";
    safetyQuickBtn.className = "impianto-safety-quick-btn";
    safetyQuickBtn.setAttribute("aria-label", "Apri sicurezza impianto");
    safetyQuickBtn.innerHTML = "<span aria-hidden='true'>🦺</span>";
    safetyQuickBtn.addEventListener("click", (event) => {
      event.preventDefault();
      event.stopPropagation();
      openImpiantoSafetyProcedureModal(impianto);
    });
    secondaryActionsRow.appendChild(safetyQuickBtn);
    addAction(
      "whatsapp",
      "✉️",
      "Whazzup / Fatto",
      async () => {
        if (isNetworkOffline() && impianto.done) {
          alert("Sei offline: WhatsApp non può essere aperto. Il messaggio resta in attesa finché torna internet.");
          return;
        }
        const whatsappOpened = triggerImpiantoWhatsAppAction(impianto);
        if (!whatsappOpened) return;
        if (!impianto.done) markImpiantoDoneVisualFallback(impianto);
        try {
          const doneMarked = await markImpiantoDone(impianto, { source: "whatsapp" });
          if (!doneMarked) markImpiantoDoneVisualFallback(impianto);
        } catch (error) {
          console.warn("Salvataggio FATTO da WHAZZUP non riuscito:", error);
        }
      },
      false,
      false,
      primaryActionsRow
    );
    addAction("problem-report", "🚨", "Segnala problema", () => openImpiantoReportModal(impianto), false, false, managementActions);
    addAction("gps-update", "📍", "Aggiorna GPS", () => requestGpsUpdate(impianto), false, true, managementActions);
    if (canManageData()) addAction("reset", "♻️", "Reset", () => resetImpianto(impianto), false, false, managementActions);
    if (canManageData()) addAction("edit", "✏️", "Modifica", () => openImpiantoEditor(impianto), false, true, managementActions);
    if (canManageData()) addAction("delete", "🗑️", "Elimina", () => deleteImpianto(impianto), false, true, managementActions);
    if (canManageData()) {
      const uploadPdfBtn = createButton("Inserisci PDF richiesta", () => setImpiantoRequestDriveLink(impianto));
      uploadPdfBtn.classList.add("pdf-request-btn");
      managementActions.appendChild(uploadPdfBtn);
    }
    if (managementActions.childElementCount > 0) {
      const manageBtn = createButton("⚙️", () => {
        if (expandedImpiantoManagementKeys.has(impiantoKey)) expandedImpiantoManagementKeys.delete(impiantoKey);
        else expandedImpiantoManagementKeys.add(impiantoKey);
        renderImpianti();
      });
      manageBtn.classList.add("gestione-toggle-btn");
      manageBtn.setAttribute("aria-label", "Gestione impianto");
      manageBtn.title = "Gestione";
      manageBtn.setAttribute("aria-expanded", isManagementExpanded ? "true" : "false");
      managementStack.appendChild(manageBtn);
      managementActions.classList.toggle("hidden", !isManagementExpanded);
    }
    if (managementStack.childElementCount > 0) secondaryActionsRow.appendChild(managementStack);
    const mqBox = createPlantMqBox(impianto);
    if (mqBox) secondaryActionsRow.appendChild(mqBox);
    if (primaryActionsRow.childElementCount > 0) actions.appendChild(primaryActionsRow);
    if (secondaryActionsRow.childElementCount > 0) actions.appendChild(secondaryActionsRow);
    const richiestaPdfUrl = String(
      impianto.linkRichiestaDrive
      || impianto.linkDocumentoRichiesta
      || impianto.richiestaPdfDriveWebViewLink
      || impianto.richiestaPdfDriveDirectUrl
      || impianto.richiestaPdfDataUrl
      || ""
    ).trim();
    if (detailsVisible && richiestaPdfUrl) {
      const viewPdfBtn = createButton("Visualizza richiesta", () => window.open(richiestaPdfUrl, "_blank"));
      viewPdfBtn.classList.add("pdf-request-view-btn");
      actions.appendChild(viewPdfBtn);
    } else if (detailsVisible) {
      const noPdfLabel = document.createElement("small");
      noPdfLabel.className = "muted";
      noPdfLabel.textContent = "Nessuna richiesta caricata";
      actions.appendChild(noPdfLabel);
    }

    if (actions.childElementCount > 0) mainColumn.appendChild(actions);
    if (managementActions.childElementCount > 0) mainColumn.appendChild(managementActions);

    article.appendChild(mainColumn);
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
  ui.editSfalci.value = impianto.sfalci || (impianto.areaMq == null ? "" : formatAreaMqValue(impianto.areaMq));
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
    areaMq: parseAreaMqValue(ui.editSfalci.value),
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

async function setImpiantoRequestDriveLink(impianto) {
  if (!selectedCommessaId || !canManageData() || !impianto) return;
  const ids = getImpiantoDocIds(impianto);
  if (!ids.length) {
    alert("Impianto non valido per salvare il link richiesta.");
    return;
  }
  const currentLink = String(impianto.linkRichiestaDrive || "").trim();
  const linkRichiestaDrive = window.prompt("Link documento richiesta (Google Drive):", currentLink);
  if (linkRichiestaDrive == null) return;
  const normalizedLink = String(linkRichiestaDrive || "").trim();
  if (normalizedLink && !/^https?:\/\//i.test(normalizedLink)) {
    alert("Inserisci un link valido (http/https).");
    return;
  }
  try {
    const ref = db.collection("commesse").doc(selectedCommessaId).collection("impianti");
    await Promise.all(ids.map((impiantoId) => ref.doc(impiantoId).set({
      linkRichiestaDrive: normalizedLink,
      richiestaPdfUpdatedAt: firebase.firestore.FieldValue.serverTimestamp(),
      richiestaPdfUpdatedBy: currentUser?.email || ""
    }, { merge: true })));
    alert(normalizedLink ? "Link richiesta salvato." : "Link richiesta rimosso.");
  } catch (error) {
    console.error("Errore salvataggio link richiesta:", error);
    alert("Impossibile salvare il link richiesta.");
  }
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

function getImpiantoNavigationCoordinates(impianto) {
  const lat = Number(impianto?.gpsY);
  const lon = Number(impianto?.gpsX);
  if (!Number.isFinite(lat) || !Number.isFinite(lon)) return null;
  if (lat < -90 || lat > 90 || lon < -180 || lon > 180) return null;
  return { lat, lon };
}

function buildImpiantoMapsUrl(impianto) {
  const coordinates = getImpiantoNavigationCoordinates(impianto);
  if (!coordinates) return "";
  return `https://www.google.com/maps/dir/?api=1&destination=${coordinates.lat},${coordinates.lon}`;
}


function getImpiantoWeatherCacheKey(impianto) {
  const coordinates = getImpiantoNavigationCoordinates(impianto);
  const coordinateKey = getImpiantoWeatherCoordinateKey(coordinates);
  const impiantoKey = buildImpiantoKey(impianto);
  if (impiantoKey && coordinateKey) return `${impiantoKey}@${coordinateKey}`;
  if (impiantoKey) return impiantoKey;
  if (!coordinates) return "";
  return coordinateKey;
}

function getPlantWeatherCacheStorageKey(plant) {
  const plantId = String(plant?.id || buildImpiantoKey(plant) || "").trim();
  return plantId ? `weather_cache_${plantId}` : "";
}

function readPlantWeatherLocalCache(plant) {
  const storageKey = getPlantWeatherCacheStorageKey(plant);
  if (!storageKey) return null;
  try {
    const raw = localStorage.getItem(storageKey);
    if (!raw) return null;
    const parsed = JSON.parse(raw);
    if (!parsed || !Number.isFinite(Number(parsed.updatedAt))) return null;
    return normalizeStoredImpiantoWeatherStatus({ ...parsed, impiantoKey: getImpiantoWeatherCacheKey(plant), impiantoId: getImpiantoWeatherCacheKey(plant) });
  } catch (error) {
    console.warn("Cache meteo impianto non leggibile:", error);
    return null;
  }
}

function writePlantWeatherLocalCache(plant, status) {
  const storageKey = getPlantWeatherCacheStorageKey(plant);
  if (!storageKey || !status) return;
  try {
    localStorage.setItem(storageKey, JSON.stringify(status));
  } catch (error) {
    console.warn("Cache meteo impianto non salvata:", error);
  }
}

function getImpiantoWeatherCoordinateKey(coordinates) {
  if (!coordinates) return "";
  return `${Number(coordinates.lat).toFixed(IMPIANTO_WEATHER_COORDINATE_PRECISION)},${Number(coordinates.lon).toFixed(IMPIANTO_WEATHER_COORDINATE_PRECISION)}`;
}

function isImpiantoWeatherCacheFresh(entry) {
  return Boolean(entry && Date.now() - Number(entry.updatedAt || 0) < IMPIANTO_WEATHER_CACHE_TTL_MS);
}

function ensureImpiantoWeatherPersistentCacheLoaded() {
  if (impiantoWeatherPersistentCacheLoaded) return;
  impiantoWeatherPersistentCacheLoaded = true;
  try {
    const raw = localStorage.getItem(IMPIANTO_WEATHER_LOCAL_CACHE_KEY);
    const parsed = raw ? JSON.parse(raw) : [];
    if (!Array.isArray(parsed)) return;
    parsed.forEach((entry) => {
      if (!(entry?.impiantoId || entry?.impiantoKey) || !Number.isFinite(Number(entry.updatedAt))) return;
      const normalized = normalizeStoredImpiantoWeatherStatus(entry);
      impiantoWeatherStatusCache.set(normalized.impiantoKey, normalized);
      if (normalized.coordinateKey) impiantoWeatherCoordinateCache.set(normalized.coordinateKey, normalized);
    });
  } catch (error) {
    console.warn("Cache meteo impianti non leggibile:", error);
  }
}

function normalizeStoredImpiantoWeatherStatus(entry) {
  const lat = Number(entry.lat);
  const lon = Number(entry.lon);
  const coordinates = Number.isFinite(lat) && Number.isFinite(lon) ? { lat, lon } : null;
  const syntheticState = ["ok", "pioggia", "vento", "temporale", "allerta"].includes(entry.syntheticState) ? entry.syntheticState : "ok";
  const description = /Meteo non disponibile/i.test(entry.description || "") ? "Meteo temporaneamente non disponibile" : (entry.description || "Meteo temporaneamente non disponibile");
  const riskLevel = getRiskLevelFromSyntheticImpiantoWeatherState(syntheticState, description);
  return {
    impiantoKey: String(entry.impiantoId || entry.impiantoKey || ""),
    lat: coordinates?.lat ?? null,
    lon: coordinates?.lon ?? null,
    coordinateKey: getImpiantoWeatherCoordinateKey(coordinates),
    syntheticState,
    statusLabel: entry.statusLabel || description,
    iconType: entry.iconType || syntheticState,
    weatherState: description,
    temperature: Number.isFinite(Number(entry.temperature)) ? Number(entry.temperature) : null,
    apparentTemperature: Number.isFinite(Number(entry.apparentTemperature)) ? Number(entry.apparentTemperature) : null,
    rainProbability: Number.isFinite(Number(entry.rainProbability ?? entry.precipitationProbability)) ? Number(entry.rainProbability ?? entry.precipitationProbability) : null,
    currentRainProbability: Number.isFinite(Number(entry.currentRainProbability ?? entry.rainProbability ?? entry.precipitationProbability)) ? Number(entry.currentRainProbability ?? entry.rainProbability ?? entry.precipitationProbability) : null,
    precipitationProbability: Number.isFinite(Number(entry.precipitationProbability ?? entry.rainProbability)) ? Number(entry.precipitationProbability ?? entry.rainProbability) : null,
    precipitationIntensity: isPresentFiniteNumber(entry.precipitationIntensity ?? entry.rainAmount) ? Number(entry.precipitationIntensity ?? entry.rainAmount) : null,
    description,
    hasCurrentRain: syntheticState === "pioggia",
    hasNextHourRain: syntheticState === "pioggia",
    civilProtectionAlert: null,
    riskLevel,
    alertText: description,
    badgeLabel: getBadgeLabelFromSyntheticImpiantoWeatherState(syntheticState, description),
    rainExpected: Boolean(entry.rainExpected || entry.rainWindow),
    rainStartTime: isPresentFiniteNumber(entry.rainStartTime ?? entry.rainWindow?.start) ? Number(entry.rainStartTime ?? entry.rainWindow?.start) : null,
    rainEndTime: isPresentFiniteNumber(entry.rainEndTime ?? entry.rainWindow?.end) ? Number(entry.rainEndTime ?? entry.rainWindow?.end) : null,
    rainWindow: entry.rainWindow || null,
    rainIntensity: entry.rainIntensity || "",
    rainAmount: isPresentFiniteNumber(entry.rainAmount) ? Number(entry.rainAmount) : null,
    windSpeed: isPresentFiniteNumber(entry.windSpeed ?? entry.windSpeedKmh ?? entry.importantWindKmh) ? Number(entry.windSpeed ?? entry.windSpeedKmh ?? entry.importantWindKmh) : null,
    windSpeedKmh: isPresentFiniteNumber(entry.windSpeedKmh ?? entry.windSpeed ?? entry.importantWindKmh) ? Number(entry.windSpeedKmh ?? entry.windSpeed ?? entry.importantWindKmh) : null,
    windDirectionDegrees: isPresentFiniteNumber(entry.windDirectionDegrees) ? Number(entry.windDirectionDegrees) : null,
    windDirection: entry.windDirection || entry.windDirectionLabel || getWindDirectionLabel(entry.windDirectionDegrees),
    windDirectionLabel: entry.windDirectionLabel || entry.windDirection || getWindDirectionLabel(entry.windDirectionDegrees),
    windGust: isPresentFiniteNumber(entry.windGust ?? entry.windGustKmh) ? Number(entry.windGust ?? entry.windGustKmh) : null,
    windGustKmh: isPresentFiniteNumber(entry.windGustKmh ?? entry.windGust) ? Number(entry.windGustKmh ?? entry.windGust) : null,
    importantWindKmh: isPresentFiniteNumber(entry.importantWindKmh) ? Number(entry.importantWindKmh) : null,
    riskMessage: entry.riskMessage || entry.operationMessage || getImpiantoWeatherOperationMessage({ riskLevel, hasCurrentRain: syntheticState === "pioggia", hasNextHourRain: syntheticState === "pioggia", importantWindKmh: entry.importantWindKmh }),
    operationMessage: entry.operationMessage || entry.riskMessage || getImpiantoWeatherOperationMessage({ riskLevel, hasCurrentRain: syntheticState === "pioggia", hasNextHourRain: syntheticState === "pioggia", importantWindKmh: entry.importantWindKmh }),
    stale: Boolean(entry.stale),
    weatherPartial: Boolean(entry.weatherPartial),
    messages: Array.isArray(entry.messages) ? entry.messages : (syntheticState === "ok" ? [] : [description]),
    currentWeather: entry.currentWeather || null,
    forecastSlots: Array.isArray(entry.forecastSlots) ? entry.forecastSlots : [],
    updatedAt: Number(entry.updatedAt) || Date.now()
  };
}

function getRiskLevelFromSyntheticImpiantoWeatherState(syntheticState, description = "") {
  if (!description || /non disponibile/i.test(description)) return "unavailable";
  if (["allerta", "temporale", "vento"].includes(syntheticState)) return "red";
  if (syntheticState === "pioggia") return "yellow";
  return "green";
}

function getBadgeLabelFromSyntheticImpiantoWeatherState(syntheticState, description = "") {
  if (!description || /non disponibile/i.test(description)) return "Meteo temporaneamente non disponibile";
  return {
    allerta: "Allerta meteo",
    temporale: "Temporale",
    vento: "Vento",
    pioggia: "Pioggia",
    ok: "Meteo OK"
  }[syntheticState] || "Meteo temporaneamente non disponibile";
}

function persistImpiantoWeatherCache() {
  try {
    const entries = Array.from(impiantoWeatherStatusCache.values())
      .filter((entry) => entry?.impiantoKey)
      .sort((a, b) => Number(b.updatedAt || 0) - Number(a.updatedAt || 0))
      .slice(0, IMPIANTO_WEATHER_LOCAL_CACHE_MAX_ENTRIES)
      .map((entry) => ({
        impiantoId: entry.impiantoKey,
        lat: entry.lat ?? null,
        lon: entry.lon ?? null,
        syntheticState: entry.syntheticState || "ok",
        statusLabel: entry.statusLabel || entry.weatherState || entry.description || "",
        iconType: entry.iconType || entry.syntheticState || "cloud",
        temperature: entry.temperature ?? null,
        apparentTemperature: entry.apparentTemperature ?? null,
        rainProbability: entry.rainProbability ?? entry.precipitationProbability ?? null,
        currentRainProbability: entry.currentRainProbability ?? entry.rainProbability ?? entry.precipitationProbability ?? null,
        precipitationProbability: entry.precipitationProbability ?? entry.rainProbability ?? null,
        precipitationIntensity: entry.precipitationIntensity ?? entry.rainAmount ?? null,
        rainExpected: Boolean(entry.rainExpected || entry.rainWindow),
        rainStartTime: entry.rainStartTime ?? entry.rainWindow?.start ?? null,
        rainEndTime: entry.rainEndTime ?? entry.rainWindow?.end ?? null,
        rainWindow: entry.rainWindow || null,
        rainIntensity: entry.rainIntensity || "",
        rainAmount: isPresentFiniteNumber(entry.rainAmount) ? Number(entry.rainAmount) : null,
        windSpeed: isPresentFiniteNumber(entry.windSpeed ?? entry.windSpeedKmh) ? Number(entry.windSpeed ?? entry.windSpeedKmh) : null,
        windSpeedKmh: isPresentFiniteNumber(entry.windSpeedKmh ?? entry.windSpeed) ? Number(entry.windSpeedKmh ?? entry.windSpeed) : null,
        windDirectionDegrees: isPresentFiniteNumber(entry.windDirectionDegrees) ? Number(entry.windDirectionDegrees) : null,
        windDirection: entry.windDirection || entry.windDirectionLabel || getWindDirectionLabel(entry.windDirectionDegrees),
        windDirectionLabel: entry.windDirectionLabel || entry.windDirection || getWindDirectionLabel(entry.windDirectionDegrees),
        windGust: isPresentFiniteNumber(entry.windGust ?? entry.windGustKmh) ? Number(entry.windGust ?? entry.windGustKmh) : null,
        windGustKmh: isPresentFiniteNumber(entry.windGustKmh ?? entry.windGust) ? Number(entry.windGustKmh ?? entry.windGust) : null,
        importantWindKmh: isPresentFiniteNumber(entry.importantWindKmh) ? Number(entry.importantWindKmh) : null,
        currentWeather: entry.currentWeather || null,
        forecastSlots: Array.isArray(entry.forecastSlots) ? entry.forecastSlots.slice(0, 12) : [],
        riskMessage: entry.riskMessage || entry.operationMessage || "",
        operationMessage: entry.operationMessage || entry.riskMessage || "",
        stale: Boolean(entry.stale),
        weatherPartial: Boolean(entry.weatherPartial),
        description: entry.description || entry.badgeLabel || "Meteo temporaneamente non disponibile",
        updatedAt: Number(entry.updatedAt) || Date.now()
      }));
    localStorage.setItem(IMPIANTO_WEATHER_LOCAL_CACHE_KEY, JSON.stringify(entries));
  } catch (error) {
    console.warn("Cache meteo impianti non salvata:", error);
  }
}

function getCachedImpiantoWeatherStatus(impianto) {
  ensureImpiantoWeatherPersistentCacheLoaded();
  const key = getImpiantoWeatherCacheKey(impianto);
  if (key && impiantoWeatherStatusCache.has(key)) return impiantoWeatherStatusCache.get(key) || null;
  const plantCached = readPlantWeatherLocalCache(impianto);
  if (plantCached) {
    impiantoWeatherStatusCache.set(plantCached.impiantoKey, plantCached);
    if (plantCached.coordinateKey) impiantoWeatherCoordinateCache.set(plantCached.coordinateKey, plantCached);
    return plantCached;
  }
  const coordinateKey = getImpiantoWeatherCoordinateKey(getImpiantoNavigationCoordinates(impianto));
  return coordinateKey ? impiantoWeatherCoordinateCache.get(coordinateKey) || null : null;
}

function isImpiantoWeatherUpdating(impianto) {
  const key = getImpiantoWeatherCacheKey(impianto);
  const coordinateKey = getImpiantoWeatherCoordinateKey(getImpiantoNavigationCoordinates(impianto));
  return Boolean((key && impiantoWeatherPendingKeys.has(key)) || (coordinateKey && impiantoWeatherPendingKeys.has(coordinateKey)));
}

function getImpiantoWeatherFeedback(impianto) {
  const key = getImpiantoWeatherCacheKey(impianto);
  const coordinateKey = getImpiantoWeatherCoordinateKey(getImpiantoNavigationCoordinates(impianto));
  return (key && impiantoWeatherFeedbackByKey.get(key))
    || (coordinateKey && impiantoWeatherFeedbackByKey.get(coordinateKey))
    || null;
}

function setImpiantoWeatherFeedback(impianto, message = "") {
  const key = getImpiantoWeatherCacheKey(impianto);
  const coordinateKey = getImpiantoWeatherCoordinateKey(getImpiantoNavigationCoordinates(impianto));
  [key, coordinateKey].filter(Boolean).forEach((cacheKey) => {
    if (message) impiantoWeatherFeedbackByKey.set(cacheKey, message);
    else impiantoWeatherFeedbackByKey.delete(cacheKey);
  });
}


function normalizeAtexSearchValue(value) {
  return String(value || "").trim().toLocaleUpperCase("it-IT");
}


const DEFAULT_IMPIANTO_SAFETY_CONTACTS = [
  { id: "default-112", name: "Emergenze", role: "Numero unico emergenze", phone: "112", type: "emergenza", whatsappEnabled: true, isDefault: true },
  { id: "default-115", name: "Vigili del Fuoco", role: "Soccorso tecnico urgente", phone: "115", type: "emergenza", whatsappEnabled: true, isDefault: true },
  { id: "default-118", name: "Ambulanza", role: "Emergenza sanitaria", phone: "118", type: "emergenza", whatsappEnabled: true, isDefault: true },
  { id: "default-hera-bo-acqua", name: "Pronto Intervento Hera Bologna", role: "Acqua, fognature e depurazione • 24h", phone: "800713900", type: "pronto intervento", whatsappEnabled: true, isDefault: true },
  { id: "default-varga-ionel", name: "Varga Ionel", role: "Capo squadra", phone: "3892352575", type: "capo squadra", whatsappEnabled: true, isDefault: true },
  { id: "default-alessandro-minarini", name: "Alessandro Minarini", role: "Responsabile commessa", phone: "+393356815371", type: "responsabile", whatsappEnabled: true, isDefault: true }
];

const IMPIANTO_SAFETY_SECTIONS = [
  {
    cls: "access",
    title: "🚧 ACCESSO SEDI TECNICHE",
    eyebrow: "⚠️ ACCESSO REGOLAMENTATO",
    intro: "Le squadre manutenzione verde operano solo dove autorizzate dal referente Hera.",
    icons: ["🚪 Cancello", "⛔ Divieto accesso", "🛣️ Percorso autorizzato"],
    groups: [
      { title: "Operare esclusivamente", items: ["nelle aree verdi autorizzate", "nei percorsi consentiti", "nelle aree assegnate dal referente Hera"] },
      { title: "NON devono", danger: true, items: ["entrare in vasche", "accedere ai locali impianto", "aprire tombini o pozzetti", "ostacolare accessi tecnici", "lasciare materiali davanti a cancelli o quadri"] }
    ]
  },
  {
    cls: "info",
    title: "🚛 CONTINUITÀ OPERATIVA IMPIANTO",
    intro: "Le attività di sfalcio e manutenzione verde devono garantire continuità, accessibilità e sicurezza operativa dell’impianto.",
    groups: [
      { title: "Garantire sempre", items: ["accessibilità impianto", "sicurezza operativa", "passaggio mezzi Hera", "accesso squadre emergenza", "visibilità della segnaletica"] },
      { title: "Controllare sempre", items: ["cancelli liberi", "strade interne sgombre", "assenza rami sui passaggi", "assenza materiali vicino impianti"] }
    ]
  },
  {
    cls: "equipment",
    title: "🚜 CONTROLLO ATTREZZATURE",
    intro: "Checklist obbligatoria prima dell’avvio di decespugliatori, soffiatori, rasaerba e attrezzature a motore.",
    checklist: ["protezione decespugliatore", "stato lama/testina", "perdite carburante", "acceleratore funzionante", "dispositivi sicurezza presenti", "rumorosità anomala", "fissaggio imbragature", "livello carburante"],
    footer: "❌ Se l’attrezzatura non è sicura: NON UTILIZZARE"
  },
  {
    cls: "fire",
    title: "🔥 RISCHIO INCENDIO",
    eyebrow: "⚠️ ATTENZIONE RISCHIO INCENDIO",
    groups: [
      { title: "In presenza di", items: ["erba secca", "caldo intenso", "vento forte"] },
      { title: "Prestare attenzione a", items: ["scarichi motori caldi", "scintille", "attriti metallici", "carburante", "mozziconi"] },
      { title: "Vietato", danger: true, items: ["fumare vicino vegetazione secca", "lasciare motori accesi inutilmente"] }
    ]
  },
  {
    cls: "plant",
    title: "⚡ SICUREZZA IMPIANTI",
    eyebrow: "⚡ ATTENZIONE IMPIANTI TECNICI",
    intro: "Mantenere distanza di sicurezza da quadri, tubazioni, sensori e apparecchiature.",
    groups: [
      { title: "Vietato", danger: true, items: ["appoggiare attrezzature agli impianti", "lavorare vicino a quadri elettrici", "dirigere getti o materiali verso apparecchiature", "urtare tubazioni o sensori"] }
    ]
  },
  {
    cls: "traffic",
    title: "🚚 TRAFFICO INTERNO IMPIANTO",
    eyebrow: "⚠️ ATTENZIONE MEZZI OPERATIVI",
    groups: [
      { title: "Possibile presenza di", items: ["autospurghi", "camion", "pale meccaniche", "mezzi Hera", "manutentori"] },
      { title: "Obblighi", items: ["usare gilet alta visibilità", "mantenere contatto visivo coi conducenti", "non lavorare dietro ai mezzi", "attenzione nelle curve e strade strette"] }
    ]
  },
  {
    cls: "biological",
    title: "☣️ RISCHIO BIOLOGICO",
    intro: "Nelle aree di depurazione e discarica può essere presente rischio biologico da reflui, fanghi, aerosol contaminati e superfici contaminate.",
    groups: [
      { title: "Fonti di esposizione", items: ["reflui", "fanghi", "aerosol contaminati", "superfici contaminate", "rifiuti organici o materiali sospetti"] },
      { title: "Norme operative", items: ["non mangiare durante il lavoro", "lavare mani obbligatoriamente", "disinfettare ferite", "cambiare guanti sporchi", "evitare contatto viso/bocca", "segnalare liquidi, odori forti o materiali contaminati"] }
    ]
  },
  {
    cls: "hygiene",
    title: "🧴 IGIENE OPERATIVA",
    intro: "A fine lavoro ripristinare l’area e ridurre il rischio di contaminazione della squadra e dei mezzi.",
    items: ["lavare mani", "pulire attrezzi", "disinfettare parti contaminate", "cambiare DPI sporchi", "non lasciare rifiuti nell’impianto"]
  },
  {
    cls: "dpi",
    title: "🦺 DPI OBBLIGATORI",
    items: ["Casco ove richiesto dal sito", "Scarpe antinfortunistiche S3", "Guanti da lavoro", "Occhiali o visiera", "Cuffie antirumore", "Gilet alta visibilità"]
  },
  {
    cls: "emergency",
    title: "🚨 IN CASO DI EMERGENZA",
    items: ["Allontanarsi dalla zona", "Avvisare il capo squadra", "Avvisare il responsabile commessa", "Chiamare i soccorsi se necessario", "Non intervenire su impianti o strutture tecniche"]
  }
];

const DISCARICHE_SPECIFIC_RISK_SECTIONS = [
  {
    cls: "discariche-risk",
    title: "⚠️ RISCHI SPECIFICI DISCARICHE",
    intro: "Avvisi operativi dedicati alle commesse discariche. Non bloccano l’inserimento lavoro/squadra ma richiedono attenzione costante.",
    groups: [
      {
        title: "⚠️ LAVORO IN PENDENZA",
        danger: true,
        items: [
          "rischio ribaltamento trattore/trincia",
          "rischio scivolamento operatore",
          "caduta su scarpate",
          "perdita controllo decespugliatore",
          "terreno instabile, fango, ghiaia, erba bagnata",
          "obbligo valutazione pendenza prima di iniziare",
          "vietato lavorare con mezzi su pendenze pericolose",
          "procedere lentamente",
          "usare DPI: casco, scarpe antiscivolo, guanti, occhiali, cuffie, alta visibilità",
          "sospendere lavoro con pioggia o terreno instabile"
        ]
      },
      {
        title: "⛽ TUBAZIONI GAS A TERRA NASCOSTE DALL’ERBA",
        danger: true,
        items: [
          "rischio urto/taglio tubo con trattore, trincia o decespugliatore",
          "rischio tritatura tubo gas",
          "rischio fuga gas",
          "rischio incendio/esplosione",
          "rischio ustioni e danni ai mezzi",
          "obbligo sopralluogo visivo prima di lavorare",
          "vietato passare sopra tubi con mezzi",
          "lavorare manualmente nelle zone critiche",
          "mantenere distanza di sicurezza",
          "in caso di odore gas o tubo danneggiato: fermare tutto, spegnere motori, allontanarsi, chiamare emergenza"
        ]
      },
      {
        title: "🕳️ POZZETTI, TOMBINI E CANALETTE NASCOSTE",
        items: [
          "rischio caduta operatore",
          "rischio ribaltamento mezzo",
          "rischio danneggiamento trattore/trincia",
          "rischio inciampo, distorsioni, fratture",
          "rischio caduta in pozzetti aperti o danneggiati",
          "obbligo controllo area prima del taglio",
          "procedere lentamente con erba alta",
          "segnalare subito tombini aperti o danneggiati",
          "delimitare area pericolosa",
          "usare decespugliatore manuale nelle zone non sicure"
        ]
      },
      {
        title: "Rischi collegati",
        items: [
          "🔥 rischio incendio con erba secca, scintille, pietre/metallo",
          "🧫 rischio biologico da rifiuti, insetti, animali, materiali contaminati",
          "🌫️ rischio polveri e scarsa visibilità",
          "🔊 rischio rumore e vibrazioni",
          "🚜 rischio investimento tra mezzi e operatori"
        ]
      }
    ]
  },
  {
    cls: "gas-emergency",
    title: "🚨 EMERGENZA GAS",
    eyebrow: "RIQUADRO ROSSO",
    items: [
      "NON accendere motori",
      "NON fumare",
      "NON usare fiamme libere",
      "spegnere mezzi",
      "allontanare operatori",
      "chiamare 112 / 115 / responsabile"
    ]
  }
];

const IMPIANTO_SAFETY_MANDATORY_CHECKLIST = [
  "DPI indossati",
  "Area controllata",
  "Meteo verificato",
  "Mezzi controllati",
  "Attrezzature controllate",
  "Accessi liberi",
  "Nessuna anomalia visibile",
  "Squadra pronta"
];

const IMPIANTO_SAFETY_ANOMALY_CATEGORIES = ["sicurezza", "impianto", "biologico", "incendio", "mezzi", "ostacoli", "altro"];

const IMPIANTO_SAFETY_WEATHER_ALERTS = [
  { key: "wet-grass", icon: "🌧", label: "Erba bagnata", risk: "medio", when: (analysis) => analysis.wetGround, explanation: "Erba, rampe e camminamenti bagnati aumentano scivolamenti e perdita di controllo delle attrezzature.", advice: "Ridurre velocità, evitare pendenze critiche, aumentare distanza tra operatori e usare calzature S3 in buono stato." },
  { key: "strong-wind", icon: "💨", label: "Vento forte", risk: "medio/alto", when: (analysis) => analysis.maxWind >= 25 || analysis.maxGust >= 35, explanation: "Il vento può deviare materiale proiettato dal decespugliatore e rendere instabili rami o vegetazione.", advice: "Orientare il lavoro lontano da persone, mezzi e vetrate; sospendere se le raffiche rendono insicura l’attività." },
  { key: "thunder", icon: "⚡", label: "Temporali", risk: "alto", when: (analysis) => analysis.thunderWithinWindow, explanation: "Temporali e fulmini espongono la squadra a rischio elettrico, scarsa visibilità e terreno rapidamente scivoloso.", advice: "Sospendere attività, allontanarsi da alberi/strutture metalliche e attendere indicazioni del responsabile." },
  { key: "heat", icon: "🌡", label: "Caldo intenso", risk: "medio/alto", when: (analysis) => analysis.hotModerate || analysis.intenseHeat, explanation: "Temperature elevate aumentano affaticamento, disidratazione e rischio di colpi di calore.", advice: "Aumentare pause, bere acqua, alternare gli operatori e preferire zone d’ombra nelle ore più calde." },
  { key: "fire", icon: "🔥", label: "Rischio incendi alto", risk: "alto", when: (analysis) => analysis.fireRiskHigh || analysis.intenseHeat, explanation: "Erba secca, caldo e motori caldi possono innescare principi di incendio durante sfalcio o decespugliamento.", advice: "Controllare marmitte e carburante, evitare scintille, non fumare e segnalare subito fumo o odore di bruciato." },
  { key: "mud", icon: "🚜", label: "Terreno fangoso", risk: "medio", when: (analysis) => analysis.wetGround && Number(analysis.rainProbability) >= 35, explanation: "Il terreno fangoso riduce aderenza di mezzi e operatori e può bloccare passaggi tecnici.", advice: "Valutare portanza, evitare manovre vicino a fossi/pendenze e mantenere liberi accessi per mezzi Hera." }
];

function getCurrentCommessaSafetyKind() {
  return getCommessaSafetyKind(selectedCommessaId);
}

function getCommessaSafetyKind(commessaId) {
  const commessa = commesseById.get(commessaId) || {};
  const values = [];
  let cursorId = commessaId;
  const visited = new Set();
  while (cursorId && !visited.has(cursorId)) {
    visited.add(cursorId);
    const item = commesseById.get(cursorId) || {};
    values.push(item.nome, item.name, item.codice, item.code, item.cliente, item.customer, item.category, item.categoria, item.tipologia, item.tipo);
    cursorId = String(item.parentCommessaId || "").trim();
  }
  values.push(commessaId, selectedCommessaName);
  const normalized = normalizeCommessaNameForRules(values.filter(Boolean).join(" "));
  if (normalized.includes("DISCARIC")) return "discariche";
  if (normalized.includes("DEPUR")) return "depurazione";
  return "";
}

function isCurrentCommessaDepurazioneOrDiscariche() {
  return Boolean(getCurrentCommessaSafetyKind());
}

function shouldShowImpiantoSafetyButtonForImpianto() {
  return isCurrentCommessaDepurazioneOrDiscariche();
}

function getImpiantoSafetyImpiantoByKey(key) {
  return getDettaglioMeteoImpiantoByKey(key);
}

function openImpiantoSafetyPage(impianto) {
  if (!selectedCommessaId || !impianto || !isCurrentCommessaDepurazioneOrDiscariche()) return;
  const key = buildImpiantoKey(impianto);
  window.location.hash = `commessa=${encodeURIComponent(selectedCommessaId)}&safety=${encodeURIComponent(key)}`;
  applyRoute();
}

function closeImpiantoSafetyPage() {
  openImpiantiPage();
}

function handleImpiantoSafetyButtonClick(event) {
  const button = event.target?.closest?.("[data-impianto-safety]");
  if (!button) return;
  event.preventDefault();
  event.stopPropagation();
  const key = button.getAttribute("data-impianto-safety") || button.closest("[data-weather-card]")?.getAttribute("data-weather-card") || "";
  const impianto = findImpiantoByWeatherKey(key) || getImpiantoSafetyImpiantoByKey(key);
  if (!impianto) return;
  openImpiantoSafetyPage(impianto);
}

function getCurrentImpiantoSafetyContext() {
  const route = parseCommessaHash();
  const impiantoKey = route.safety || "";
  const impianto = getImpiantoSafetyImpiantoByKey(impiantoKey) || {};
  return { impiantoKey, impianto };
}

function getImpiantoDisplayName(impianto = {}) {
  return String(impianto.denominazione || impianto.nome || impianto.impianto || impianto.descrizione || impianto.codice || "").trim();
}

function getImpiantoComune(impianto = {}) {
  return String(impianto.comune || impianto.localita || impianto.citta || impianto.city || "").trim();
}

function buildImpiantoSafetyWhatsappText(impianto = {}) {
  const operator = currentUser?.displayName || currentUser?.email || "Operatore";
  return [
    "⚠️ SEGNALAZIONE SICUREZZA IMPIANTO",
    `Commessa: ${selectedCommessaName || "Commessa"}`,
    `Impianto: ${getImpiantoDisplayName(impianto) || "non disponibile"}`,
    `Comune: ${getImpiantoComune(impianto) || "non disponibile"}`,
    `Operatore: ${operator}`,
    "Problema rilevato:",
    "Richiedo supporto urgente."
  ].join("\n");
}

function formatPhoneHref(phone = "") {
  return String(phone || "").replace(/[^+\d]/g, "");
}

function formatWhatsappPhone(phone = "") {
  const cleaned = formatPhoneHref(phone);
  if (cleaned.startsWith("+")) return cleaned.slice(1);
  if (/^3\d{8,}$/.test(cleaned)) return `39${cleaned}`;
  return cleaned;
}

function buildImpiantoSafetyList(items = [], className = "") {
  return `<ul${className ? ` class="${escapeHTML(className)}"` : ""}>${items.map((item) => `<li>${escapeHTML(item)}</li>`).join("")}</ul>`;
}

function buildImpiantoSafetySection(section) {
  const iconsMarkup = Array.isArray(section.icons) && section.icons.length
    ? `<div class="impianto-safety-icons">${section.icons.map((item) => `<span>${escapeHTML(item)}</span>`).join("")}</div>`
    : "";
  const checklistMarkup = Array.isArray(section.checklist)
    ? `<div class="impianto-safety-mini-checklist">${section.checklist.map((item) => `<span>☑ ${escapeHTML(item)}</span>`).join("")}</div>`
    : "";
  const groupsMarkup = Array.isArray(section.groups)
    ? `<div class="impianto-safety-groups">${section.groups.map((group) => `
      <div class="impianto-safety-group${group.danger ? " is-danger" : ""}">
        <h4>${escapeHTML(group.title)}</h4>
        ${buildImpiantoSafetyList(group.items || [])}
      </div>`).join("")}</div>`
    : "";
  const itemsMarkup = Array.isArray(section.items) ? buildImpiantoSafetyList(section.items) : "";
  return `<article class="impianto-safety-section is-${escapeHTML(section.cls)}">
    <h3>${escapeHTML(section.title)}</h3>
    ${section.eyebrow ? `<p class="impianto-safety-eyebrow">${escapeHTML(section.eyebrow)}</p>` : ""}
    ${section.intro ? `<p>${escapeHTML(section.intro)}</p>` : ""}
    ${iconsMarkup}${groupsMarkup}${checklistMarkup}${itemsMarkup}
    ${section.footer ? `<p class="impianto-safety-stop">${escapeHTML(section.footer)}</p>` : ""}
  </article>`;
}

function getImpiantoSafetyOperatorName() {
  return currentUser?.displayName || currentUser?.email || "Operatore";
}

function buildImpiantoSafetyChecklistSection() {
  return `
    <article class="impianto-safety-section is-checklist">
      <h3>📋 CHECKLIST OBBLIGATORIA</h3>
      <p>Conferma operativa prima di iniziare. La registrazione viene salvata su Firebase con data, ora, operatore, commessa e impianto.</p>
      <form id="impianto-safety-checklist-form" class="impianto-safety-checklist-form">
        <div class="impianto-safety-checklist-grid">
          ${IMPIANTO_SAFETY_MANDATORY_CHECKLIST.map((item, index) => `
            <label class="impianto-safety-check-row">
              <input type="checkbox" name="checklist" value="${escapeHTML(item)}" data-safety-check-index="${index}" required>
              <span>☑ ${escapeHTML(item)}</span>
            </label>`).join("")}
        </div>
        ${getCurrentCommessaSafetyKind() === "discariche" ? `
          <label class="impianto-safety-check-row impianto-safety-check-row-wide">
            <input type="checkbox" name="discaricheReadConfirm" value="true" required>
            <span>☑️ Confermo di aver letto le norme di sicurezza discariche</span>
          </label>` : ""}
        <button class="btn btn-primary impianto-safety-confirm-btn" type="submit">✅ CONFERMA CONTROLLO SICUREZZA</button>
        <p class="muted" data-safety-checklist-feedback role="status" aria-live="polite"></p>
      </form>
    </article>`;
}

function buildImpiantoSafetyAnomalySection(impianto = {}) {
  const whatsappText = buildImpiantoSafetyWhatsappText(impianto);
  return `
    <article class="impianto-safety-section is-anomaly">
      <h3>📸 SEGNALAZIONE ANOMALIA</h3>
      <p>Usa il pulsante rosso per segnalare subito criticità sicurezza, impianto, biologiche, incendio, mezzi o ostacoli.</p>
      <button class="btn btn-danger impianto-safety-anomaly-toggle" type="button" data-safety-toggle-anomaly>⚠️ SEGNALA ANOMALIA</button>
      <form id="impianto-safety-anomaly-form" class="impianto-safety-anomaly-form hidden">
        <label>Categoria<select name="category" required>${IMPIANTO_SAFETY_ANOMALY_CATEGORIES.map((item) => `<option value="${escapeHTML(item)}">${escapeHTML(item)}</option>`).join("")}</select></label>
        <label>Foto<input name="photo" type="file" accept="image/*" capture="environment"></label>
        <label>Nota<textarea name="note" rows="4" placeholder="Descrivi anomalia, punto impianto, rischio e azione richiesta" required></textarea></label>
        <div class="impianto-safety-form-actions">
          <button class="btn" type="button" data-safety-get-gps>📍 POSIZIONE GPS</button>
          <a class="btn impianto-safety-whatsapp" href="https://wa.me/?text=${encodeURIComponent(whatsappText)}" target="_blank" rel="noopener" data-safety-whatsapp-anomaly>💬 WHAZZUP</a>
        </div>
        <input type="hidden" name="lat"><input type="hidden" name="lon">
        <p class="muted" data-safety-gps-feedback>GPS non acquisito.</p>
        <button class="btn btn-danger" type="submit">Invia responsabile</button>
        <p class="muted" data-safety-anomaly-feedback role="status" aria-live="polite"></p>
      </form>
    </article>`;
}

function buildImpiantoSafetyWeatherSection(impianto = {}) {
  const status = getCachedImpiantoWeatherStatus(impianto) || {};
  const analysis = buildDettaglioMeteoRiskAnalysis(status);
  const alerts = IMPIANTO_SAFETY_WEATHER_ALERTS.map((alert) => ({ ...alert, active: Boolean(alert.when(analysis, status)) }));
  const hasActive = alerts.some((alert) => alert.active);
  return `
    <article class="impianto-safety-section is-weather-ops">
      <h3>🌧 METEO OPERATIVO</h3>
      <p>${hasActive ? "Avvisi automatici rilevati per la squadra." : "Nessun avviso critico automatico rilevato: mantenere comunque i controlli ordinari."}</p>
      <div class="impianto-safety-weather-alerts">
        ${alerts.map((alert) => `
          <button type="button" class="impianto-safety-weather-alert${alert.active ? " is-active" : ""}" data-safety-weather-alert="${escapeHTML(alert.key)}" data-explanation="${escapeHTML(alert.explanation)}" data-advice="${escapeHTML(alert.advice)}" data-risk="${escapeHTML(alert.risk)}">
            <span>${escapeHTML(alert.icon)}</span><strong>${escapeHTML(alert.label)}</strong><small>${alert.active ? "Avviso attivo" : "Controllo"}</small>
          </button>`).join("")}
      </div>
      <div class="impianto-safety-weather-detail hidden" data-safety-weather-detail></div>
    </article>`;
}

function buildImpiantoSafetyAdminSection(procedureConfig = {}) {
  if (!canManageData()) return "";
  return `
    <article class="impianto-safety-section is-admin">
      <h3>🔐 ADMIN SICUREZZA</h3>
      <p>Solo admin: aggiornamento procedure, norme sicurezza e numeri di emergenza della scheda.</p>
      <form id="impianto-safety-admin-form" class="impianto-safety-contact-form">
        <label>Procedure operative<textarea name="procedure" rows="4" placeholder="Modifica procedure operative per depurazione/discariche">${escapeHTML(procedureConfig.procedure || "")}</textarea></label>
        <label>Norme sicurezza<textarea name="norme" rows="4" placeholder="Aggiorna norme sicurezza da capitolato Hera / INRETE manutenzione verde">${escapeHTML(procedureConfig.norme || "")}</textarea></label>
        <button class="btn btn-primary" type="submit">Aggiorna norme sicurezza</button>
        <p class="muted" data-safety-admin-feedback role="status" aria-live="polite"></p>
      </form>
    </article>`;
}

async function loadImpiantoSafetyProcedureConfig() {
  if (!canManageData()) return {};
  const kind = getCurrentCommessaSafetyKind() || "depurazione_discariche";
  try {
    const doc = await db.collection("impiantoSafetyProcedures").doc(kind).get();
    return doc.exists ? doc.data() || {} : {};
  } catch (error) {
    console.warn("Norme sicurezza impianto non caricate:", error);
    return {};
  }
}

function isSafetyContactVisibleForCommessa(contact = {}, kind = getCurrentCommessaSafetyKind()) {
  const scope = String(contact.scope || "depurazione_discariche").trim();
  if (scope === "commessa") return String(contact.commessaId || "") === String(selectedCommessaId || "");
  if (scope === "depurazione") return kind === "depurazione";
  if (scope === "discariche") return kind === "discariche";
  return scope === "depurazione_discariche" || !scope;
}

async function loadImpiantoSafetyContacts() {
  let customContacts = [];
  try {
    const snapshot = await db.collection("safetyContacts").get();
    customContacts = snapshot.docs.map((doc) => ({ id: doc.id, ...doc.data(), isDefault: false }));
  } catch (error) {
    console.warn("Contatti sicurezza impianto non caricati:", error);
  }
  const kind = getCurrentCommessaSafetyKind();
  return [
    ...DEFAULT_IMPIANTO_SAFETY_CONTACTS,
    ...customContacts.filter((contact) => isSafetyContactVisibleForCommessa(contact, kind))
  ];
}

function buildImpiantoSafetyContactCard(contact, whatsappText) {
  const phoneHref = formatPhoneHref(contact.phone);
  const whatsappPhone = formatWhatsappPhone(contact.phone);
  const whatsappUrl = whatsappPhone ? `https://wa.me/${encodeURIComponent(whatsappPhone)}?text=${encodeURIComponent(whatsappText)}` : "";
  const canEdit = canManageData() && !contact.isDefault;
  const whatsappMarkup = contact.whatsappEnabled === false || !whatsappUrl
    ? `<button class="btn impianto-safety-whatsapp" type="button" disabled>💬 WHAZZUP</button>`
    : `<a class="btn impianto-safety-whatsapp" href="${whatsappUrl}" target="_blank" rel="noopener">💬 WHAZZUP</a>`;
  const adminActions = canEdit ? `<div class="impianto-safety-contact-admin"><button class="btn" type="button" data-safety-edit-contact="${escapeHTML(contact.id)}">Modifica</button><button class="btn btn-danger" type="button" data-safety-delete-contact="${escapeHTML(contact.id)}">Elimina</button></div>` : "";
  return `
    <article class="impianto-safety-contact" data-safety-contact-id="${escapeHTML(contact.id || "")}">
      <div><strong>${escapeHTML(contact.name || "Contatto")}</strong><small>${escapeHTML(contact.role || contact.type || "")}</small><small>${escapeHTML(contact.phone || "")}</small></div>
      <div class="impianto-safety-contact-actions">
        <a class="btn btn-primary" href="tel:${escapeHTML(phoneHref)}">📞 CHIAMA</a>
        ${whatsappMarkup}
      </div>
      ${adminActions}
    </article>`;
}

function buildImpiantoSafetyContactsSection(contacts = [], impianto = {}) {
  const whatsappText = buildImpiantoSafetyWhatsappText(impianto);
  return `
    <article class="impianto-safety-section is-contacts">
      <div class="impianto-safety-section-head">
        <h3>📞 NUMERI DI EMERGENZA E SUPPORTO</h3>
        ${canManageData() ? `<button class="btn btn-primary" type="button" data-safety-add-contact>➕ AGGIUNGI NUMERO</button>` : ""}
      </div>
      <p class="muted">Per Hera Bologna il numero 800713900 è indicato come pronto intervento acqua, fognature e depurazione, attivo 24 ore su 24.</p>
      <div class="impianto-safety-contacts-grid">${contacts.map((contact) => buildImpiantoSafetyContactCard(contact, whatsappText)).join("")}</div>
      <div id="impianto-safety-contact-form-wrap" class="impianto-safety-contact-form-wrap hidden"></div>
    </article>`;
}

function buildImpiantoSafetyContactForm(contact = {}) {
  const isEdit = Boolean(contact.id);
  const scope = String(contact.scope || "commessa");
  const type = String(contact.type || "altro");
  const whatsappEnabled = contact.whatsappEnabled !== false;
  const option = (value, label, current) => `<option value="${escapeHTML(value)}"${current === value ? " selected" : ""}>${escapeHTML(label)}</option>`;
  return `
    <form id="impianto-safety-contact-form" class="impianto-safety-contact-form" data-contact-id="${escapeHTML(contact.id || "")}">
      <h4>${isEdit ? "Modifica numero" : "Aggiungi numero"}</h4>
      <label>Nome contatto<input name="name" type="text" value="${escapeHTML(contact.name || "")}" required></label>
      <label>Ruolo<input name="role" type="text" value="${escapeHTML(contact.role || "")}" required></label>
      <label>Numero telefono<input name="phone" type="tel" value="${escapeHTML(contact.phone || "")}" required></label>
      <label>Tipo numero<select name="type">${["emergenza", "pronto intervento", "capo squadra", "responsabile", "cliente", "sicurezza", "altro"].map((item) => option(item, item, type)).join("")}</select></label>
      <label>Visibilità<select name="scope">
        ${option("commessa", "solo questa commessa", scope)}
        ${option("depurazione", "tutte le commesse depurazione", scope)}
        ${option("discariche", "tutte le commesse discariche", scope)}
        ${option("depurazione_discariche", "tutte le commesse depurazione e discariche", scope)}
      </select></label>
      <label>Attiva WhatsApp<select name="whatsappEnabled">${option("true", "sì", String(whatsappEnabled))}${option("false", "no", String(whatsappEnabled))}</select></label>
      <div class="impianto-safety-form-actions"><button class="btn btn-primary" type="submit">Salva numero</button><button class="btn" type="button" data-safety-cancel-form>Annulla</button></div>
      <p class="muted" data-safety-form-feedback role="status" aria-live="polite"></p>
    </form>`;
}

async function renderImpiantoSafetyPage(impiantoKey) {
  if (!ui.impiantoSafetyContent) return;
  const impianto = getImpiantoSafetyImpiantoByKey(impiantoKey) || {};
  const impiantoName = getImpiantoDisplayName(impianto) || "Impianto";
  if (ui.impiantoSafetySubtitle) ui.impiantoSafetySubtitle.textContent = `${impiantoName} • ${selectedCommessaName || "Commessa"}`;
  const contacts = await loadImpiantoSafetyContacts();
  const procedureConfig = await loadImpiantoSafetyProcedureConfig();
  const commessaSafetyKind = getCurrentCommessaSafetyKind();
  const extraDiscaricheSections = commessaSafetyKind === "discariche"
    ? DISCARICHE_SPECIFIC_RISK_SECTIONS.map(buildImpiantoSafetySection).join("")
    : "";
  ui.impiantoSafetyContent.innerHTML = `
    <article class="impianto-safety-hero">
      <div class="impianto-safety-hero-badge">Sicurezza industriale Hera • manutenzione verde</div>
      <h3>🦺 SICUREZZA IMPIANTO<br>DEPURAZIONE / DISCARICHE</h3>
      <p>Schermata operativa per squadre manutenzione verde: accessi regolamentati, continuità impianto, attrezzature, meteo operativo, anomalie ed emergenze.</p>
      <div class="impianto-safety-hero-meta">
        <span>👷 ${escapeHTML(getImpiantoSafetyOperatorName())}</span>
        <span>🏭 ${escapeHTML(impiantoName)}</span>
        <span>📍 ${escapeHTML(getImpiantoComune(impianto) || "Comune non indicato")}</span>
      </div>
    </article>
    <div class="impianto-safety-grid">${IMPIANTO_SAFETY_SECTIONS.map(buildImpiantoSafetySection).join("")}${extraDiscaricheSections}</div>
    ${buildImpiantoSafetyChecklistSection()}
    ${buildImpiantoSafetyAnomalySection(impianto)}
    ${buildImpiantoSafetyWeatherSection(impianto)}
    ${buildImpiantoSafetyContactsSection(contacts, impianto)}
    ${buildImpiantoSafetyAdminSection(procedureConfig)}
  `;
}

async function getSafetyContactById(contactId) {
  if (!contactId) return null;
  const doc = await db.collection("safetyContacts").doc(contactId).get();
  return doc.exists ? { id: doc.id, ...doc.data() } : null;
}

async function handleImpiantoSafetyContentClick(event) {
  const addBtn = event.target?.closest?.("[data-safety-add-contact]");
  const editBtn = event.target?.closest?.("[data-safety-edit-contact]");
  const deleteBtn = event.target?.closest?.("[data-safety-delete-contact]");
  const cancelBtn = event.target?.closest?.("[data-safety-cancel-form]");
  const anomalyToggle = event.target?.closest?.("[data-safety-toggle-anomaly]");
  const gpsBtn = event.target?.closest?.("[data-safety-get-gps]");
  const weatherBtn = event.target?.closest?.("[data-safety-weather-alert]");
  if (anomalyToggle) {
    event.preventDefault();
    const form = ui.impiantoSafetyContent?.querySelector("#impianto-safety-anomaly-form");
    form?.classList.toggle("hidden");
    form?.querySelector("select, textarea, input")?.focus?.();
    return;
  }
  if (gpsBtn) {
    event.preventDefault();
    acquireImpiantoSafetyGps(gpsBtn.closest("form"));
    return;
  }
  if (weatherBtn) {
    event.preventDefault();
    const detail = ui.impiantoSafetyContent?.querySelector("[data-safety-weather-detail]");
    if (!detail) return;
    detail.innerHTML = `
      <h4>${escapeHTML(weatherBtn.textContent.trim())}</h4>
      <p><strong>Spiegazione:</strong> ${escapeHTML(weatherBtn.dataset.explanation || "")}</p>
      <p><strong>Comportamento consigliato:</strong> ${escapeHTML(weatherBtn.dataset.advice || "")}</p>
      <p><strong>Livello rischio:</strong> ${escapeHTML(weatherBtn.dataset.risk || "")}</p>`;
    detail.classList.remove("hidden");
    detail.scrollIntoView({ behavior: "smooth", block: "nearest" });
    return;
  }
  if (addBtn || editBtn) {
    if (!canManageData()) return;
    const wrap = ui.impiantoSafetyContent?.querySelector("#impianto-safety-contact-form-wrap");
    if (!wrap) return;
    let contact = {};
    if (editBtn) contact = await getSafetyContactById(editBtn.getAttribute("data-safety-edit-contact"));
    wrap.innerHTML = buildImpiantoSafetyContactForm(contact || {});
    wrap.classList.remove("hidden");
    wrap.scrollIntoView({ behavior: "smooth", block: "center" });
    return;
  }
  if (deleteBtn) {
    if (!canManageData()) return;
    const contactId = deleteBtn.getAttribute("data-safety-delete-contact") || "";
    if (!contactId || !window.confirm("Eliminare questo numero di sicurezza?")) return;
    await db.collection("safetyContacts").doc(contactId).delete();
    const { impiantoKey } = getCurrentImpiantoSafetyContext();
    renderImpiantoSafetyPage(impiantoKey);
    return;
  }
  if (cancelBtn) {
    const wrap = ui.impiantoSafetyContent?.querySelector("#impianto-safety-contact-form-wrap");
    if (wrap) wrap.classList.add("hidden");
  }
}

function acquireImpiantoSafetyGps(form) {
  const feedback = form?.querySelector("[data-safety-gps-feedback]");
  if (!form || !navigator.geolocation) {
    if (feedback) feedback.textContent = "GPS non disponibile su questo dispositivo.";
    return;
  }
  if (feedback) feedback.textContent = "Acquisizione GPS…";
  navigator.geolocation.getCurrentPosition((position) => {
    form.elements.lat.value = String(position.coords.latitude || "");
    form.elements.lon.value = String(position.coords.longitude || "");
    if (feedback) feedback.textContent = `GPS acquisito: ${position.coords.latitude.toFixed(5)}, ${position.coords.longitude.toFixed(5)}`;
  }, () => {
    if (feedback) feedback.textContent = "Impossibile acquisire GPS. Verifica autorizzazioni posizione.";
  }, { enableHighAccuracy: true, timeout: 10000, maximumAge: 60000 });
}

function readSafetyPhotoAsDataUrl(file) {
  return new Promise((resolve) => {
    if (!file) {
      resolve("");
      return;
    }
    const reader = new FileReader();
    reader.onload = () => resolve(String(reader.result || ""));
    reader.onerror = () => resolve("");
    reader.readAsDataURL(file);
  });
}

function getImpiantoSafetyPayloadBase(impiantoKey, impianto) {
  const now = new Date();
  return {
    commessaId: selectedCommessaId || "",
    commessaName: selectedCommessaName || "",
    impiantoKey: impiantoKey || "",
    impiantoName: getImpiantoDisplayName(impianto) || "",
    impiantoComune: getImpiantoComune(impianto) || "",
    operatore: getImpiantoSafetyOperatorName(),
    date: formatAtexDateValue(now),
    time: formatAtexTimeValue(now),
    createdByUid: auth.currentUser?.uid || "",
    createdByEmail: auth.currentUser?.email || "",
    createdAt: firebase.firestore.FieldValue.serverTimestamp(),
    updatedAt: firebase.firestore.FieldValue.serverTimestamp()
  };
}

async function saveImpiantoSafetyChecklistForm(event) {
  if (event.target?.id !== "impianto-safety-checklist-form") return false;
  event.preventDefault();
  const form = event.target;
  const feedback = form.querySelector("[data-safety-checklist-feedback]");
  const checked = Array.from(form.querySelectorAll("input[name='checklist']:checked")).map((input) => input.value);
  if (checked.length !== IMPIANTO_SAFETY_MANDATORY_CHECKLIST.length) {
    if (feedback) feedback.textContent = "Conferma tutti i controlli obbligatori prima di procedere.";
    return true;
  }
  const { impiantoKey, impianto } = getCurrentImpiantoSafetyContext();
  const payload = {
    ...getImpiantoSafetyPayloadBase(impiantoKey, impianto),
    checklist: checked,
    checklistComplete: true,
    safetyKind: getCurrentCommessaSafetyKind(),
    safetyType: getCurrentCommessaSafetyKind() === "discariche" ? "DISCARICHE" : "",
    discaricheReadConfirm: getCurrentCommessaSafetyKind() === "discariche"
  };
  try {
    if (feedback) feedback.textContent = "Salvataggio controllo sicurezza…";
    await db.collection("impiantoSafetyChecks").add(payload);
    if (feedback) feedback.textContent = "✅ Controllo sicurezza salvato su Firebase.";
    form.querySelector("button[type='submit']").disabled = true;
  } catch (error) {
    console.error("Controllo sicurezza non salvato:", error);
    if (feedback) feedback.textContent = "Errore salvataggio controllo sicurezza. Riprova.";
  }
  return true;
}

async function saveImpiantoSafetyAnomalyForm(event) {
  if (event.target?.id !== "impianto-safety-anomaly-form") return false;
  event.preventDefault();
  const form = event.target;
  const feedback = form.querySelector("[data-safety-anomaly-feedback]");
  const data = Object.fromEntries(new FormData(form).entries());
  const photoFile = form.elements.photo?.files?.[0] || null;
  const { impiantoKey, impianto } = getCurrentImpiantoSafetyContext();
  if (!String(data.note || "").trim()) {
    if (feedback) feedback.textContent = "Inserisci una nota per inviare la segnalazione.";
    return true;
  }
  try {
    if (feedback) feedback.textContent = "Invio segnalazione anomalia…";
    const photoDataUrl = await readSafetyPhotoAsDataUrl(photoFile);
    const payload = {
      ...getImpiantoSafetyPayloadBase(impiantoKey, impianto),
      category: String(data.category || "altro"),
      note: String(data.note || "").trim(),
      lat: String(data.lat || ""),
      lon: String(data.lon || ""),
      photoName: photoFile?.name || "",
      photoType: photoFile?.type || "",
      photoDataUrl: photoDataUrl.slice(0, 900000),
      hasPhoto: Boolean(photoDataUrl),
      status: "inviata"
    };
    await db.collection("impiantoSafetyAnomalies").add(payload);
    const whatsapp = form.querySelector("[data-safety-whatsapp-anomaly]");
    if (whatsapp) {
      const text = [buildImpiantoSafetyWhatsappText(impianto), `Categoria: ${payload.category}`, `Nota: ${payload.note}`, payload.lat && payload.lon ? `GPS: ${payload.lat}, ${payload.lon}` : "GPS: non acquisito"].join("\n");
      whatsapp.href = `https://wa.me/?text=${encodeURIComponent(text)}`;
    }
    if (feedback) feedback.textContent = "✅ Segnalazione inviata e pronta per condivisione WhatsApp.";
    form.reset();
  } catch (error) {
    console.error("Segnalazione anomalia non salvata:", error);
    if (feedback) feedback.textContent = "Errore invio segnalazione. Riprova.";
  }
  return true;
}

async function saveImpiantoSafetyAdminForm(event) {
  if (event.target?.id !== "impianto-safety-admin-form") return false;
  event.preventDefault();
  if (!canManageData()) return true;
  const form = event.target;
  const feedback = form.querySelector("[data-safety-admin-feedback]");
  const data = Object.fromEntries(new FormData(form).entries());
  try {
    if (feedback) feedback.textContent = "Aggiornamento norme…";
    await db.collection("impiantoSafetyProcedures").doc(getCurrentCommessaSafetyKind() || "depurazione_discariche").set({
      procedure: String(data.procedure || "").trim(),
      norme: String(data.norme || "").trim(),
      updatedByUid: auth.currentUser?.uid || "",
      updatedByEmail: auth.currentUser?.email || "",
      updatedAt: firebase.firestore.FieldValue.serverTimestamp()
    }, { merge: true });
    if (feedback) feedback.textContent = "✅ Procedure e norme sicurezza aggiornate.";
  } catch (error) {
    console.error("Norme sicurezza non aggiornate:", error);
    if (feedback) feedback.textContent = "Errore aggiornamento norme sicurezza.";
  }
  return true;
}

async function saveImpiantoSafetyContactForm(event) {
  if (await saveImpiantoSafetyChecklistForm(event)) return;
  if (await saveImpiantoSafetyAnomalyForm(event)) return;
  if (await saveImpiantoSafetyAdminForm(event)) return;
  if (event.target?.id !== "impianto-safety-contact-form") return;
  event.preventDefault();
  if (!canManageData()) return;
  const form = event.target;
  const feedback = form.querySelector("[data-safety-form-feedback]");
  const contactId = form.getAttribute("data-contact-id") || "";
  const data = Object.fromEntries(new FormData(form).entries());
  const payload = {
    name: String(data.name || "").trim(),
    role: String(data.role || "").trim(),
    phone: String(data.phone || "").trim(),
    type: String(data.type || "altro").trim(),
    whatsappEnabled: String(data.whatsappEnabled) === "true",
    scope: String(data.scope || "commessa").trim(),
    commessaId: String(data.scope || "") === "commessa" ? selectedCommessaId : null,
    updatedAt: firebase.firestore.FieldValue.serverTimestamp()
  };
  if (!payload.name || !payload.role || !payload.phone) {
    if (feedback) feedback.textContent = "Compila nome, ruolo e telefono.";
    return;
  }
  try {
    if (feedback) feedback.textContent = "Salvataggio in corso…";
    if (contactId) {
      await db.collection("safetyContacts").doc(contactId).set(payload, { merge: true });
    } else {
      await db.collection("safetyContacts").add({
        ...payload,
        createdBy: currentUser?.uid || "",
        createdAt: firebase.firestore.FieldValue.serverTimestamp()
      });
    }
    const { impiantoKey } = getCurrentImpiantoSafetyContext();
    renderImpiantoSafetyPage(impiantoKey);
  } catch (error) {
    console.error("Numero sicurezza impianto non salvato:", error);
    if (feedback) feedback.textContent = "Errore durante il salvataggio. Riprova.";
  }
}

function valueContainsAtex(value) {
  return /\bATEX\b/.test(normalizeAtexSearchValue(value));
}

function isTruthyAtexFlag(value) {
  if (value === true) return true;
  if (value === false || value === null || value === undefined) return false;
  const normalized = normalizeAtexSearchValue(value);
  return ["ATEX", "TRUE", "SI", "SÌ", "YES", "1", "EX"].includes(normalized) || valueContainsAtex(normalized);
}

function isInreteCommessa(commessa = {}) {
  return [
    commessa.nome,
    commessa.name,
    commessa.codice,
    commessa.code,
    commessa.id,
    commessa.cliente,
    commessa.customer,
    commessa.category,
    commessa.categoria,
    commessa.tipologia
  ].some((value) => normalizeAtexSearchValue(value).includes("INRETE"));
}

function isCurrentCommessaInrete() {
  const selected = commesseById.get(selectedCommessaId) || {};
  return isInreteCommessa({
    id: selectedCommessaId,
    nome: selectedCommessaName,
    ...selected
  });
}

function hasImpiantoAtexFlag(impianto = {}) {
  const explicitFlags = [
    impianto.atex,
    impianto.isAtex,
    impianto.flagAtex,
    impianto.atexFlag,
    impianto.areaAtex,
    impianto.zonaAtex,
    impianto.ex
  ];
  if (explicitFlags.some(isTruthyAtexFlag)) return true;
  return [
    impianto.category,
    impianto.categoria,
    impianto.flag,
    impianto.flags,
    impianto.tipo,
    impianto.tipologia,
    impianto.tipologiaImpianto,
    impianto.lavorazioniRichieste,
    impianto.note,
    impianto.codicePrezzo,
    impianto.voceRiferimento,
    impianto.denominazione
  ].some(valueContainsAtex);
}

function shouldShowAtexButtonForImpianto(impianto = {}) {
  if (isCurrentCommessaDepurazioneOrDiscariche()) return false;
  return Boolean(isCurrentCommessaInrete() || hasImpiantoAtexFlag(impianto));
}

function getAtexProcedureImpiantoByKey(key) {
  return getDettaglioMeteoImpiantoByKey(key);
}

function openAtexProcedurePage(impianto) {
  if (!selectedCommessaId || !impianto) return;
  const key = buildImpiantoKey(impianto);
  window.location.hash = `commessa=${encodeURIComponent(selectedCommessaId)}&atex=${encodeURIComponent(key)}`;
  applyRoute();
}

function closeAtexProcedurePage() {
  openImpiantiPage();
}

function handleAtexProcedureButtonClick(event) {
  const button = event.target?.closest?.("[data-atex-procedure]");
  if (!button) return;
  event.preventDefault();
  event.stopPropagation();
  const key = button.getAttribute("data-atex-procedure") || button.closest("[data-weather-card]")?.getAttribute("data-weather-card") || "";
  const impianto = findImpiantoByWeatherKey(key) || getAtexProcedureImpiantoByKey(key);
  if (!impianto) return;
  openAtexProcedurePage(impianto);
}

function getCurrentAtexProcedureContext() {
  const route = parseCommessaHash();
  const impiantoKey = route.atex || "";
  const impianto = getAtexProcedureImpiantoByKey(impiantoKey) || {};
  return { impiantoKey, impianto };
}

function formatAtexDateValue(date = new Date()) {
  const year = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, "0");
  const day = String(date.getDate()).padStart(2, "0");
  return `${year}-${month}-${day}`;
}

function formatAtexTimeValue(date = new Date()) {
  const hours = String(date.getHours()).padStart(2, "0");
  const minutes = String(date.getMinutes()).padStart(2, "0");
  return `${hours}:${minutes}`;
}

function buildAtexChecklist(items = []) {
  return items.map((item, index) => `
    <label class="atex-check-row">
      <input type="checkbox" name="check_${index}" value="${escapeHTML(item)}">
      <span>${escapeHTML(item)}</span>
    </label>
  `).join("");
}

function buildAtexList(items = []) {
  return `<ul>${items.map((item) => `<li>${escapeHTML(item)}</li>`).join("")}</ul>`;
}

function normalizeAtexCommessaMatchValue(value) {
  return normalizeAtexSearchValue(value).normalize("NFD").replace(/[\u0300-\u036f]/g, "");
}

function getSelectedAtexCommessaName() {
  const selected = commesseById.get(selectedCommessaId) || {};
  return selectedCommessaName || selected.nome || selected.name || selected.codice || selected.code || selectedCommessaId || "Commessa non indicata";
}

function getAtexClientContactsForCommessa() {
  const clients = [
    { area: "Bologna", name: "Carani Claudio", phone: "347 7614277" },
    { area: "Modena", name: "Montagnana Giorgio", phone: "320 4791013" },
    { area: "Ferrara", name: "Mateo Gardelini", phone: "348 0900290" }
  ];
  const selected = commesseById.get(selectedCommessaId) || {};
  const haystack = [
    selectedCommessaName,
    selectedCommessaId,
    selected.nome,
    selected.name,
    selected.codice,
    selected.code,
    selected.cliente,
    selected.customer,
    selected.category,
    selected.categoria,
    selected.tipologia
  ].map(normalizeAtexCommessaMatchValue).join(" ");
  const matched = clients.find((client) => haystack.includes(normalizeAtexCommessaMatchValue(client.area)));
  if (matched) return [{ ...matched, role: `Cliente INRETE ${matched.area}`, callFirst: true }];
  return clients.map((client) => ({
    ...client,
    role: `Cliente INRETE ${client.area}`,
    note: "cliente di riferimento da verificare",
    callFirst: true
  }));
}

function sanitizePhoneHref(phone = "") {
  const value = String(phone || "").trim();
  return value.startsWith("+") ? `+${value.replace(/\D/g, "")}` : value.replace(/\D/g, "");
}

function sanitizeWhatsappPhone(phone = "") {
  return String(phone || "").replace(/\D/g, "");
}

function getAtexWhatsappText(impianto = {}) {
  const operator = currentUser?.displayName || currentUser?.email || "Operatore non indicato";
  return [
    "⚠️ PROBLEMA ATEX RILEVATO",
    `Commessa: ${getSelectedAtexCommessaName()}`,
    `Impianto: ${impianto.denominazione || impianto.nome || "Impianto non indicato"}`,
    `Comune: ${impianto.comune || "Comune non indicato"}`,
    `Operatore: ${operator}`,
    "Serve intervento/verifica urgente."
  ].join("\n");
}

function buildAtexContactCard(contact, whatsappText) {
  const phoneHref = sanitizePhoneHref(contact.phone);
  const whatsappPhone = sanitizeWhatsappPhone(contact.phone);
  const whatsappUrl = `https://wa.me/${encodeURIComponent(whatsappPhone)}?text=${encodeURIComponent(whatsappText)}`;
  const noteMarkup = contact.note ? `<span class="atex-contact-note">${escapeHTML(contact.note)}</span>` : "";
  const badgeMarkup = contact.callFirst ? `<span class="atex-call-first-badge">CHIAMARE PRIMA</span>` : "";
  return `
    <section class="atex-contact-card${contact.callFirst ? " is-primary" : ""}">
      <div class="atex-contact-head">
        <span class="atex-contact-role">${escapeHTML(contact.role)}</span>
        ${badgeMarkup}
      </div>
      ${noteMarkup}
      <strong class="atex-contact-name">${escapeHTML(contact.name)}</strong>
      <span class="atex-contact-phone">Tel: ${escapeHTML(contact.phone)}</span>
      <div class="atex-contact-actions">
        <a class="atex-contact-action call" href="tel:${escapeHTML(phoneHref)}">📞 CHIAMA</a>
        <a class="atex-contact-action whatsapp" href="${escapeHTML(whatsappUrl)}" target="_blank" rel="noopener noreferrer">💬 WHAZZUP</a>
      </div>
    </section>
  `;
}

function buildAtexEmergencyContactsSection(impianto = {}) {
  const contacts = [
    ...getAtexClientContactsForCommessa(),
    { role: "Capo squadra", name: "Varga Ionel", phone: "0039 389 2352575" },
    { role: "Responsabile commessa", name: "Alessandro Minarini", phone: "+39 335 6815371" },
    { role: "Numero unico emergenze", name: "Numero unico emergenze", phone: "112" },
    { role: "Vigili del Fuoco", name: "Vigili del Fuoco", phone: "115" },
    { role: "Emergenza sanitaria / ambulanza", name: "Emergenza sanitaria / ambulanza", phone: "118" },
    { role: "Carabinieri", name: "Carabinieri", phone: "112" },
    { role: "Polizia", name: "Polizia", phone: "113" }
  ];
  const whatsappText = getAtexWhatsappText(impianto);
  return `
    <article class="atex-procedure-section atex-emergency-section">
      <h3>6. CONTATTI EMERGENZA ATEX</h3>
      <div class="atex-contacts-grid">${contacts.map((contact) => buildAtexContactCard(contact, whatsappText)).join("")}</div>
    </article>
  `;
}


function getAtexIllustrationSvg(type = "safety") {
  const svgMap = {
    safety: `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 420 240" role="img"><defs><linearGradient id="sky" x1="0" y1="0" x2="1" y2="1"><stop stop-color="#fed7aa"/><stop offset="1" stop-color="#dcfce7"/></linearGradient></defs><rect width="420" height="240" rx="28" fill="url(#sky)"/><path d="M0 174c60-26 118-22 178 2s119 26 242-12v76H0z" fill="#86efac"/><rect x="48" y="71" width="118" height="86" rx="14" fill="#fff7ed" stroke="#9a3412" stroke-width="6"/><path d="M78 124h59M78 101h59" stroke="#9a3412" stroke-width="10" stroke-linecap="round"/><circle cx="292" cy="96" r="46" fill="#fb923c"/><path d="M292 57l36 70h-72z" fill="#fff7ed" stroke="#7c2d12" stroke-width="7"/><path d="M292 78v24" stroke="#7c2d12" stroke-width="8" stroke-linecap="round"/><circle cx="292" cy="114" r="5" fill="#7c2d12"/><path d="M208 163c26-30 65-30 91 0" fill="none" stroke="#166534" stroke-width="10" stroke-linecap="round"/></svg>`,
    checklist: `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 320 210" role="img"><rect width="320" height="210" rx="24" fill="#ffedd5"/><rect x="76" y="30" width="168" height="150" rx="18" fill="#fffaf3" stroke="#fdba74" stroke-width="6"/><rect x="119" y="19" width="82" height="28" rx="12" fill="#9a3412"/><g fill="none" stroke-linecap="round" stroke-linejoin="round"><path d="M105 78l13 13 29-31" stroke="#16a34a" stroke-width="10"/><path d="M162 80h48" stroke="#7c2d12" stroke-width="8"/><path d="M105 124l13 13 29-31" stroke="#16a34a" stroke-width="10"/><path d="M162 126h48" stroke="#7c2d12" stroke-width="8"/></g></svg>`,
    altair: `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 320 240" role="img"><rect width="320" height="240" rx="28" fill="#fff7ed"/><rect x="96" y="26" width="128" height="188" rx="30" fill="#2f2f2f" stroke="#111827" stroke-width="8"/><rect x="118" y="52" width="84" height="50" rx="10" fill="#93c5fd"/><text x="160" y="83" text-anchor="middle" font-family="Arial" font-size="20" font-weight="900" fill="#0f172a">4XR</text><circle cx="160" cy="139" r="28" fill="#f97316"/><circle cx="160" cy="139" r="13" fill="#ffedd5"/><circle cx="124" cy="179" r="10" fill="#fb923c"/><circle cx="160" cy="179" r="10" fill="#22c55e"/><circle cx="196" cy="179" r="10" fill="#ef4444"/><path d="M125 28c4-16 66-16 70 0" fill="none" stroke="#6b7280" stroke-width="10" stroke-linecap="round"/></svg>`,
    forbidden: `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 220 150" role="img"><rect width="220" height="150" rx="20" fill="#fef2f2"/><circle cx="110" cy="75" r="47" fill="#fff" stroke="#dc2626" stroke-width="14"/><path d="M78 107l64-64" stroke="#dc2626" stroke-width="14" stroke-linecap="round"/><path d="M107 43c17 17-17 22 0 39 16 17-15 24 6 39" fill="none" stroke="#7f1d1d" stroke-width="8" stroke-linecap="round"/></svg>`,
    dpi: `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 220 150" role="img"><rect width="220" height="150" rx="20" fill="#fff7ed"/><path d="M63 103c8-31 28-50 47-50s39 19 47 50z" fill="#fb923c" stroke="#9a3412" stroke-width="6"/><path d="M72 69c5-23 23-39 38-39s33 16 38 39" fill="#fdba74"/><rect x="54" y="101" width="112" height="18" rx="9" fill="#7c2d12"/><path d="M85 76h50" stroke="#fff7ed" stroke-width="8" stroke-linecap="round"/></svg>`,
    alarm: `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 220 150" role="img"><rect width="220" height="150" rx="20" fill="#fffbeb"/><path d="M110 24l70 102H40z" fill="#facc15" stroke="#92400e" stroke-width="8" stroke-linejoin="round"/><path d="M110 61v32" stroke="#7c2d12" stroke-width="10" stroke-linecap="round"/><circle cx="110" cy="111" r="6" fill="#7c2d12"/></svg>`,
    danger: `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 220 150" role="img"><rect width="220" height="150" rx="20" fill="#fff7ed"/><path d="M44 100c42-45 91-45 132 0" fill="none" stroke="#16a34a" stroke-width="12" stroke-linecap="round"/><rect x="73" y="45" width="74" height="56" rx="12" fill="#431407"/><circle cx="110" cy="73" r="22" fill="#fb923c"/><path d="M110 55v36M92 73h36" stroke="#fff7ed" stroke-width="7" stroke-linecap="round"/></svg>`
  };
  return svgMap[type] || svgMap.safety;
}

function buildAtexImageCard(type, alt, extraClass = "") {
  const src = `data:image/svg+xml;charset=UTF-8,${encodeURIComponent(getAtexIllustrationSvg(type))}`;
  return `<figure class="atex-image-card ${escapeHTML(extraClass)}"><img src="${src}" alt="${escapeHTML(alt)}" loading="lazy"></figure>`;
}

function handleAtexProcedureContentClick(event) {
  const checklistToggle = event.target?.closest?.("[data-atex-checklist-toggle]");
  if (checklistToggle) {
    event.preventDefault();
    const targetId = checklistToggle.getAttribute("aria-controls");
    const panel = targetId ? document.getElementById(targetId) : null;
    const arrow = checklistToggle.querySelector("[data-atex-checklist-arrow]");
    const isOpen = checklistToggle.getAttribute("aria-expanded") === "true";
    checklistToggle.setAttribute("aria-expanded", String(!isOpen));
    if (panel) panel.hidden = isOpen;
    if (arrow) arrow.textContent = isOpen ? "▼" : "▲";
    return;
  }

  const openFormButton = event.target?.closest?.("[data-open-atex-form]");
  if (!openFormButton) return;
  event.preventDefault();
  const form = ui.atexProcedureContent?.querySelector?.("#atex-module-form");
  if (!form) return;
  form.classList.remove("hidden");
  openFormButton.classList.add("hidden");
  form.querySelector("input, select, textarea")?.focus?.();
}

async function saveAtexProcedureForm(event) {
  const form = event.target?.closest?.("#atex-module-form");
  if (!form) return;
  event.preventDefault();
  const feedback = form.querySelector("[data-atex-form-feedback]");
  const submitButton = form.querySelector("button[type='submit']");
  const { impiantoKey, impianto } = getCurrentAtexProcedureContext();
  const formData = new FormData(form);
  const payload = {
    commessaId: selectedCommessaId || "",
    commessaName: selectedCommessaName || "",
    impiantoKey,
    impiantoName: impianto.denominazione || formData.get("impianto") || "",
    impiantoComune: impianto.comune || "",
    operatore: String(formData.get("operatore") || "").trim(),
    squadra: String(formData.get("squadra") || "").trim(),
    impianto: String(formData.get("impianto") || "").trim(),
    data: String(formData.get("data") || "").trim(),
    ora: String(formData.get("ora") || "").trim(),
    presenzaGas: String(formData.get("presenzaGas") || "").trim(),
    altairVerificato: String(formData.get("altairVerificato") || "").trim(),
    dpiVerificati: String(formData.get("dpiVerificati") || "").trim(),
    noteOperative: String(formData.get("noteOperative") || "").trim(),
    firma: String(formData.get("firma") || "").trim(),
    checklist: Array.from(form.querySelectorAll(".atex-check-row input:checked")).map((input) => input.value),
    createdByUid: auth.currentUser?.uid || "",
    createdByEmail: auth.currentUser?.email || "",
    createdAt: firebase.firestore.FieldValue.serverTimestamp(),
    updatedAt: firebase.firestore.FieldValue.serverTimestamp()
  };
  if (!payload.operatore || !payload.squadra || !payload.impianto || !payload.data || !payload.ora || !payload.firma) {
    if (feedback) feedback.textContent = "Compila operatore, squadra, impianto, data, ora e firma.";
    return;
  }
  try {
    if (submitButton) submitButton.disabled = true;
    if (feedback) feedback.textContent = "Salvataggio modulo ATEX in corso…";
    await db.collection("commesse").doc(selectedCommessaId).collection("atexModules").add(payload);
    form.reset();
    form.querySelector("[name='impianto']").value = payload.impianto;
    form.querySelector("[name='data']").value = formatAtexDateValue();
    form.querySelector("[name='ora']").value = formatAtexTimeValue();
    if (feedback) feedback.textContent = "Modulo ATEX salvato e collegato a commessa e impianto.";
  } catch (error) {
    console.error("Modulo ATEX non salvato:", error);
    if (feedback) feedback.textContent = "Errore durante il salvataggio del modulo ATEX. Riprova.";
  } finally {
    if (submitButton) submitButton.disabled = false;
  }
}

function renderAtexProcedurePage(impiantoKey) {
  if (!ui.atexProcedureContent) return;
  const impianto = getAtexProcedureImpiantoByKey(impiantoKey) || {};
  const impiantoName = impianto.denominazione || "Impianto";
  if (ui.atexProcedureSubtitle) {
    ui.atexProcedureSubtitle.textContent = `${impiantoName} • ${selectedCommessaName || "Commessa"}`;
  }
  const mandatoryChecks = [
    "Verifica presenza gas",
    "Controllo zona classificata",
    "Controllo DPI",
    "Verifica estintore",
    "Verifica vie di fuga",
    "Controllo vento e ventilazione area",
    "Controllo autorizzazione lavoro",
    "Compilazione modulo ATEX",
    "Verifica comunicazione squadra",
    "Controllo assenza inneschi/scintille"
  ];
  const altairChecks = [
    "Accendere il dispositivo prima di entrare o avvicinarsi all’area",
    "Attendere autotest completo",
    "Verificare batteria sufficiente",
    "Verificare sensori attivi",
    "Eseguire bump test se previsto dalla procedura aziendale",
    "Non entrare in zona se il dispositivo segnala errore",
    "Tenere il dispositivo vicino alla respirazione",
    "Controllare continuamente eventuali allarmi acustici, visivi o vibrazione"
  ];
  const statusIndicators = ["Vento", "Meteo", "Pioggia", "Segnalazioni", "Sensori", "Checklist"];
  ui.atexProcedureContent.innerHTML = `
    <article class="atex-procedure-alert">
      <strong>1. AVVISO IMPORTANTE</strong>
      <p>Leggere attentamente e completare tutti i controlli prima di iniziare ogni attività.</p>
    </article>

    <article class="atex-procedure-section atex-hero-section">
      ${buildAtexImageCard("safety", "Area verde esterna con segnalazione sicurezza ATEX", "atex-hero-image")}
      <div>
        <h3>SICUREZZA ATEX IN AREA VERDE ESTERNA</h3>
        <p>Procedura operativa sintetica per lavorare solo nelle aree verdi autorizzate, mantenendo controlli, DPI e comunicazioni sempre attivi.</p>
      </div>
    </article>

    <article class="atex-procedure-section atex-access-limits">
      <h3>2. LIMITAZIONE ACCESSI</h3>
      <p><strong>È vietato entrare all’interno delle strutture.</strong></p>
      <p>L’accesso è consentito solo nel cortile e nelle aree verdi esterne autorizzate.</p>
    </article>

    <article class="atex-procedure-section atex-team-warning">
      <strong>ATTENZIONE</strong>
      <p>Le squadre di manutenzione verde NON devono entrare nelle strutture, locali tecnici, vasche, pozzetti o ambienti confinati.</p>
      <p>È consentito operare solo nel cortile e nelle aree verdi esterne autorizzate.</p>
      <p>In caso di dubbio fermarsi e contattare il responsabile.</p>
    </article>

    <article class="atex-procedure-section atex-collapsible-section">
      <button class="atex-section-toggle" type="button" data-atex-checklist-toggle aria-expanded="false" aria-controls="atex-mandatory-checklist">
        <span>3. CONTROLLI OBBLIGATORI PRIMA ATTIVITÀ</span>
        <span class="atex-section-arrow" data-atex-checklist-arrow aria-hidden="true">▼</span>
      </button>
      <div id="atex-mandatory-checklist" class="atex-collapsible-panel" hidden>
        ${buildAtexImageCard("checklist", "Checklist sicurezza prima attività", "atex-checklist-image")}
        <div class="atex-checklist-grid">${buildAtexChecklist(mandatoryChecks)}</div>
      </div>
    </article>

    <article class="atex-procedure-section atex-altair-box">
      <h3>4. UTILIZZO DISPOSITIVO ALTAIR 4XR</h3>
      <div class="atex-altair-layout">
        ${buildAtexImageCard("altair", "Dispositivo rilevatore multigas Altair 4XR", "atex-altair-image")}
        <div class="atex-device-card">
          <p><strong>Nome dispositivo:</strong> Altair 4XR</p>
          <p><strong>Uso:</strong> rilevatore multigas personale</p>
          <p><strong>Obbligo:</strong> accendere prima di avvicinarsi alla zona ATEX</p>
          <p><strong>Posizione:</strong> tenere vicino alla zona di respirazione</p>
        </div>
      </div>
      ${buildAtexList(altairChecks)}
      <div class="atex-alarm-box">
        <strong>In caso di allarme:</strong>
        <ul>
          <li>interrompere immediatamente l’attività</li>
          <li>allontanarsi dalla zona in sicurezza</li>
          <li>avvisare squadra e responsabile</li>
          <li>non rientrare senza autorizzazione</li>
          <li>vietato riavviare attività senza verifica</li>
        </ul>
      </div>
    </article>

    <article class="atex-procedure-section">
      <h3>5. NORME OPERATIVE ATEX</h3>
      <div class="atex-rules-grid">
        <section class="atex-rule-box atex-rule-forbidden">${buildAtexImageCard("forbidden", "Divieti in zona ATEX", "atex-rule-image")}<h4>🔥 DIVIETI</h4>${buildAtexList(["Vietato fumare", "Vietato usare fiamme libere", "Vietato produrre scintille", "Vietato utilizzare utensili non certificati", "Vietato usare telefoni non autorizzati in area ATEX"])}</section>
        <section class="atex-rule-box">${buildAtexImageCard("dpi", "DPI obbligatori", "atex-rule-image")}<h4>🦺 DPI OBBLIGATORI</h4>${buildAtexList(["Scarpe antistatiche", "Guanti idonei", "Visiera/protezione occhi", "Alta visibilità", "DPI previsti dalla commessa"])}</section>
        <section class="atex-rule-box">${buildAtexImageCard("alarm", "Allarme sicurezza", "atex-rule-image")}<h4>⚠️ COMPORTAMENTO OPERATIVO</h4>${buildAtexList(["Lavorare sempre in squadra", "Mantenere distanza sicurezza", "Controllare costantemente atmosfera", "Fermare attività in caso di dubbio", "Segnalare immediatamente anomalie"])}</section>
        <section class="atex-rule-box">${buildAtexImageCard("danger", "Zona pericolosa esterna", "atex-rule-image")}<h4>🌬 GAS E VENTILAZIONE</h4>${buildAtexList(["Controllare direzione vento", "Evitare ristagni gas", "Non entrare in pozzetti senza verifica", "Aerare zona prima attività"])}</section>
      </div>
    </article>

    ${buildAtexEmergencyContactsSection(impianto)}

    <article class="atex-procedure-section atex-status-section">
      <h3>7. STATO SICUREZZA</h3>
      <div class="atex-status-grid">
        <div class="atex-status-card is-green">🟢 <strong>SICUREZZA OPERATIVA</strong></div>
        <div class="atex-status-card is-yellow">🟡 <strong>ATTENZIONE</strong></div>
        <div class="atex-status-card is-red">🔴 <strong>RISCHIO ALTO</strong></div>
      </div>
      <div class="atex-indicators">${statusIndicators.map((item) => `<span>${escapeHTML(item)}</span>`).join("")}</div>
    </article>

    <article class="atex-procedure-section atex-form-section">
      <h3>8. MODULO ATEX</h3>
      <button class="btn btn-primary atex-open-form-btn" type="button" data-open-atex-form>📄 COMPILA MODULO ATEX</button>
      <form id="atex-module-form" class="atex-module-form hidden">
        <div class="atex-form-grid">
          <label>Operatore<input name="operatore" type="text" autocomplete="name" required></label>
          <label>Squadra<input name="squadra" type="text" required></label>
          <label>Impianto<input name="impianto" type="text" value="${escapeHTML(impiantoName)}" required></label>
          <label>Data<input name="data" type="date" value="${formatAtexDateValue()}" required></label>
          <label>Ora<input name="ora" type="time" value="${formatAtexTimeValue()}" required></label>
          <label>Presenza gas<select name="presenzaGas" required><option value="NO">NO</option><option value="SI">SI</option></select></label>
          <label>Altair verificato<select name="altairVerificato" required><option value="SI">SI</option><option value="NO">NO</option></select></label>
          <label>DPI verificati<select name="dpiVerificati" required><option value="SI">SI</option><option value="NO">NO</option></select></label>
          <label class="atex-form-wide">Note operative<textarea name="noteOperative" rows="4" placeholder="Annotazioni, anomalie, valori o comunicazioni operative"></textarea></label>
          <label class="atex-form-wide">Firma semplice<input name="firma" type="text" placeholder="Nome e cognome" required></label>
        </div>
        <button class="btn btn-primary" type="submit">Salva modulo ATEX</button>
        <p class="muted" data-atex-form-feedback role="status" aria-live="polite"></p>
      </form>
    </article>
  `;
}

function getImpiantoWeatherBadgeState(impianto) {
  const key = getImpiantoWeatherCacheKey(impianto);
  const coordinates = getImpiantoNavigationCoordinates(impianto);
  const entry = getCachedImpiantoWeatherStatus(impianto);
  const updating = isImpiantoWeatherUpdating(impianto);
  const feedback = getImpiantoWeatherFeedback(impianto);
  if (!coordinates) {
    return {
      level: "unavailable",
      label: "Coordinate mancanti",
      text: "Coordinate mancanti",
      updating: false,
      feedback: "",
      iconType: "cloud",
      display: { description: "Coordinate mancanti", temperature: "--°", wind: "", rain: "" },
      lines: ["Coordinate mancanti"],
      compact: true,
      canRetry: false,
      retryKey: key,
      showAtex: shouldShowAtexButtonForImpianto(impianto),
      showSafety: shouldShowImpiantoSafetyButtonForImpianto(impianto)
    };
  }
  if (!entry) {
    return {
      level: "unavailable",
      label: updating ? "Aggiornamento meteo…" : "Meteo temporaneamente non disponibile",
      text: updating ? "Aggiornamento meteo…" : "Meteo temporaneamente non disponibile",
      updating,
      feedback,
      iconType: "cloud",
      display: { description: updating ? "Aggiornamento meteo…" : "Meteo non disponibile", temperature: "--°", wind: "", rain: "" },
      lines: [updating ? "Aggiornamento meteo…" : "Meteo temporaneamente non disponibile"],
      compact: true,
      canRetry: !updating,
      retryKey: key,
      showAtex: shouldShowAtexButtonForImpianto(impianto),
      showSafety: shouldShowImpiantoSafetyButtonForImpianto(impianto)
    };
  }
  const level = entry.riskLevel || "unavailable";
  const label = getImpiantoWeatherPrimaryLabel(entry);
  const emoji = getImpiantoWeatherLevelEmoji(level);
  const updatedLabel = entry.updatedAt ? ` • aggiornato ${new Date(entry.updatedAt).toLocaleTimeString("it-IT", { hour: "2-digit", minute: "2-digit" })}` : "";
  const lines = buildImpiantoWeatherCardLines(entry, label, emoji);
  const iconType = getImpiantoWeatherIconType(entry, entry.iconType || entry.syntheticState || "cloud");
  return {
    level,
    label: `${label}${updatedLabel}${entry.weatherPartial ? " • Dati meteo parziali" : ""}`,
    text: `${emoji} ${label}`,
    updating,
    feedback,
    iconType,
    display: buildImpiantoWeatherDisplay(entry, label),
    lines,
    compact: lines.length <= 1,
    canRetry: level === "unavailable" && !updating && entry.canRetry !== false,
    retryKey: key,
    showAtex: shouldShowAtexButtonForImpianto(impianto),
      showSafety: shouldShowImpiantoSafetyButtonForImpianto(impianto)
  };
}

function getImpiantoWeatherLevelEmoji(level) {
  if (level === "red") return "🔴";
  if (level === "yellow") return "🟡";
  if (level === "green") return "🟢";
  return "⚪";
}

function getImpiantoWeatherPrimaryLabel(entry = {}) {
  if (entry.riskLevel === "red") return "Rischio";
  if (entry.riskLevel === "yellow") return entry.hasCurrentRain || entry.hasNextHourRain ? "Pioggia prevista" : "Rischio";
  if (entry.riskLevel === "green") return "Meteo ok";
  return entry.description || entry.badgeLabel || "Meteo temporaneamente non disponibile";
}

function formatImpiantoWeatherTimeRange(start, end) {
  if (!Number.isFinite(Number(start))) return "";
  const startLabel = formatWeatherSlotTime(start);
  if (!Number.isFinite(Number(end)) || Number(end) <= Number(start)) return startLabel;
  return `${startLabel}–${formatWeatherSlotTime(end)}`;
}

function formatCompactImpiantoWeatherStatus(entry = {}, label = getImpiantoWeatherPrimaryLabel(entry)) {
  const temperature = isPresentFiniteNumber(entry.temperature) ? ` · ${Math.round(Number(entry.temperature))}°C` : "";
  const weatherState = String(entry.weatherState || entry.description || label || "Meteo").replace(/^[^A-Za-zÀ-ÿ]+\s*/u, "");
  const intensity = entry.rainIntensity ? `${entry.rainIntensity}` : "";
  if (entry.hasCurrentRain) return `Ora ${intensity || weatherState || "Pioggia"}${temperature}`;
  return `Ora ${weatherState || label}${temperature}`;
}

function formatCompactImpiantoWeatherRiskLine(message = "") {
  const normalized = String(message || "").trim().toLowerCase();
  if (!normalized) return "";
  if (normalized.includes("terreno")) return "⚠️ Terreno";
  if (normalized.includes("sfalcio")) return "⚠️ Sfalcio";
  if (normalized.includes("rinvio")) return "⚠️ Rinvio";
  if (normalized.includes("allerta")) return "⚠️ Allerta";
  if (normalized.includes("tempor")) return "⚠️ Temporali";
  if (normalized.includes("ok")) return "✓ Lavoro";
  return `⚠️ ${String(message).replace(/^⚠️?\s*/u, "").split(/[.;]/)[0].trim().slice(0, 18)}`;
}

function truncateImpiantoWeatherLine(line = "", maxLength = 24) {
  const value = String(line || "").trim();
  if (value.length <= maxLength) return value;
  return `${value.slice(0, Math.max(0, maxLength - 1)).trim()}…`;
}

function buildImpiantoWeatherCardLines(entry = {}, label = getImpiantoWeatherPrimaryLabel(entry), emoji = getImpiantoWeatherLevelEmoji(entry.riskLevel)) {
  if (entry.riskLevel === "unavailable") return [entry.description || entry.badgeLabel || "Meteo temporaneamente non disponibile"];
  const lines = [formatCompactImpiantoWeatherStatus(entry, label)];

  const range = formatImpiantoWeatherTimeRange(entry.rainStartTime ?? entry.rainWindow?.start, entry.rainEndTime ?? entry.rainWindow?.end);
  if (range && (entry.hasNextHourRain || entry.rainWindow || entry.rainExpected)) lines.push(`Prossima pioggia ${range}`);

  const realWindSpeed = Number(entry.windSpeed ?? entry.windSpeedKmh);
  if (isPresentFiniteNumber(realWindSpeed) && realWindSpeed > 0.4) {
    const direction = entry.windDirection || entry.windDirectionLabel || "";
    lines.push(`Vento ${Math.round(realWindSpeed)} km/h${direction ? ` ${direction}` : ""}`);
  }

  return lines.filter(Boolean).map((line) => truncateImpiantoWeatherLine(line, 32)).slice(0, 4);
}

function getImpiantoWeatherIconType(entry = {}, fallback = "cloud") {
  const text = `${entry.weatherState || ""} ${entry.description || ""} ${entry.badgeLabel || ""}`.toLowerCase();
  const state = String(entry.syntheticState || fallback || "cloud").toLowerCase();
  if (state.includes("temporale") || /tempor|fulmin|thunder/.test(text)) return "storm";
  if (state.includes("vento") || /vento forte|raffica/.test(text)) return "wind";
  if (state.includes("pioggia") || /piogg|rovesc|rain/.test(text)) return /moderata|forte|intens/.test(text) ? "rain" : "rain-light";
  if (/nebb|fog|foschia/.test(text)) return "fog";
  if (/copert|overcast/.test(text)) return "overcast";
  if (/nubi sparse|parzial|poco nuvol|partly|scattered/.test(text)) return "partly";
  if (/seren|clear|sole/.test(text) || state === "ok") return "sun";
  return fallback || "cloud";
}

function buildCompactSfalcioWarnings(entry = {}) {
  const warnings = [];
  const add = (text) => { if (text && !warnings.includes(text)) warnings.push(text); };
  if (entry.hasCurrentRain || entry.hasNextHourRain || entry.rainWindow) add("Sfalcio: terreno bagnato");
  if (isPresentFiniteNumber(entry.windSpeedKmh) && Number(entry.windSpeedKmh) >= 20) add("Sfalcio: vento, cautela");
  if (entry.riskLevel === "red") add("Sfalcio: valuta rinvio");
  if (isPresentFiniteNumber(entry.temperature) && Number(entry.temperature) > 32) add("Sfalcio: caldo, pause");
  if (!warnings.length && entry.riskLevel === "green") add("Sfalcio regolare");
  return warnings.slice(0, 2);
}


function getImpiantoCompactCondition(entry = {}) {
  const weatherCode = Number(entry.currentWeather?.weather_code ?? entry.weatherCode);
  const weatherText = String(entry.weatherState || entry.description || "").toLowerCase();
  const temp = Number(entry.temperature ?? entry.apparentTemperature);
  const wind = Number(entry.windSpeed ?? entry.windSpeedKmh ?? entry.importantWindKmh);

  if (Number.isFinite(temp) && temp >= 35) return "🔥 Caldo intenso";
  if (Number.isFinite(wind) && wind >= NAVIGATION_WEATHER_STRONG_WIND_KMH) return "💨 Vento forte";
  if (isThunderWeatherCode(weatherCode) || /tempor|thunder|fulmin/.test(weatherText)) return "⛈️ Temporale";
  if ((weatherCode >= 71 && weatherCode <= 77) || weatherCode === 85 || weatherCode === 86 || /neve|snow/.test(weatherText)) return "❄️ Neve";
  if (weatherCode === 45 || weatherCode === 48 || /nebb|fog|foschia/.test(weatherText)) return "🌫️ Nebbia";
  if (NAVIGATION_WEATHER_RAIN_CODES.has(weatherCode) || /pioggi|rovesc|rain/.test(weatherText)) return "🌧️ Pioggia";
  if (weatherCode === 0 || /seren|sole|clear/.test(weatherText)) return "☀️ Sole";
  return "☁️ Nuvoloso";
}

function buildImpiantoWeatherDisplay(entry = {}, label = getImpiantoWeatherPrimaryLabel(entry)) {
  const description = getImpiantoCompactCondition(entry);
  const temperature = isPresentFiniteNumber(entry.temperature) ? `${Math.round(Number(entry.temperature))}°` : "--°";
  const windSpeed = Number(entry.windSpeed ?? entry.windSpeedKmh);
  const windDirection = entry.windDirection || entry.windDirectionLabel || "";
  const wind = isPresentFiniteNumber(windSpeed) && windSpeed > 0.4 ? `${Math.round(windSpeed)} km/h${windDirection ? ` ${windDirection}` : ""}` : "";
  const rainRange = formatImpiantoWeatherTimeRange(entry.rainStartTime ?? entry.rainWindow?.start, entry.rainEndTime ?? entry.rainWindow?.end);
  const rain = rainRange && (entry.hasNextHourRain || entry.rainWindow || entry.rainExpected) ? `Prossima pioggia ${rainRange}` : "";
  return {
    description: truncateImpiantoWeatherLine(description || label || "Meteo", 34),
    temperature,
    wind: truncateImpiantoWeatherLine(wind, 24),
    rain: truncateImpiantoWeatherLine(rain, 30),
    alerts: buildCompactSfalcioWarnings(entry).map((line) => truncateImpiantoWeatherLine(line, 26))
  };
}

function buildImpiantoWeatherIconSvg(type = "cloud") {
  const normalized = String(type || "cloud").toLowerCase();
  const cloud = `<g filter="url(#softShadow)"><path d="M41 74h48c10 0 18-7 18-17 0-9-7-16-16-17-4-12-15-20-29-20-16 0-29 10-32 25-9 1-16 7-16 15 0 8 7 14 27 14Z" fill="url(#cloudGrad)"/></g>`;
  const defs = `<defs><filter id="softShadow" x="-20%" y="-20%" width="140%" height="150%"><feDropShadow dx="0" dy="5" stdDeviation="5" flood-color="#0f172a" flood-opacity=".18"/></filter><radialGradient id="sunGrad" cx="35%" cy="30%" r="70%"><stop offset="0" stop-color="#fff7ad"/><stop offset=".55" stop-color="#fbbf24"/><stop offset="1" stop-color="#f97316"/></radialGradient><linearGradient id="cloudGrad" x1="0" y1="0" x2="1" y2="1"><stop offset="0" stop-color="#ffffff"/><stop offset="1" stop-color="#cbd5e1"/></linearGradient><linearGradient id="darkCloudGrad" x1="0" y1="0" x2="1" y2="1"><stop offset="0" stop-color="#94a3b8"/><stop offset="1" stop-color="#475569"/></linearGradient><linearGradient id="dropGrad" x1="0" y1="0" x2="0" y2="1"><stop offset="0" stop-color="#60a5fa"/><stop offset="1" stop-color="#0ea5e9"/></linearGradient></defs>`;
  const sun = `<g filter="url(#softShadow)"><circle cx="55" cy="50" r="25" fill="url(#sunGrad)"/><g stroke="#fbbf24" stroke-width="7" stroke-linecap="round" opacity=".75"><path d="M55 12v10M55 78v10M17 50h10M83 50h10M28 23l7 7M75 70l7 7M82 23l-7 7M35 70l-7 7"/></g></g>`;
  const rainDrops = `<g fill="url(#dropGrad)" filter="url(#softShadow)"><path d="M43 82c5 7 5 13 0 15-5-2-5-8 0-15Z"/><path d="M64 82c5 7 5 13 0 15-5-2-5-8 0-15Z"/><path d="M85 82c5 7 5 13 0 15-5-2-5-8 0-15Z"/></g>`;
  const rainMore = `<g fill="url(#dropGrad)" filter="url(#softShadow)"><path d="M34 82c5 7 5 13 0 15-5-2-5-8 0-15Z"/><path d="M51 86c5 7 5 13 0 15-5-2-5-8 0-15Z"/><path d="M68 82c5 7 5 13 0 15-5-2-5-8 0-15Z"/><path d="M86 86c5 7 5 13 0 15-5-2-5-8 0-15Z"/></g>`;
  const darkCloud = cloud.replace('fill="url(#cloudGrad)"', 'fill="url(#darkCloudGrad)"');
  let body = cloud;
  if (normalized === "sun") body = sun;
  else if (normalized === "partly") body = `${sun.replace('cx="55" cy="50"', 'cx="45" cy="42"')}${cloud}`;
  else if (normalized === "overcast" || normalized === "cloud") body = `${cloud}<path d="M23 65h60c8 0 14-5 14-12" fill="none" stroke="#e2e8f0" stroke-width="8" stroke-linecap="round" opacity=".8"/>`;
  else if (normalized === "rain-light") body = `${cloud}${rainDrops}`;
  else if (normalized === "rain") body = `${darkCloud}${rainMore}`;
  else if (normalized === "storm" || normalized === "allerta" || normalized === "temporale") body = `${darkCloud}${rainMore}<path d="M61 67 48 96h17l-6 23 24-38H66l10-14Z" fill="#facc15" filter="url(#softShadow)"/>`;
  else if (normalized === "fog") body = `<g stroke="#94a3b8" stroke-width="9" stroke-linecap="round" filter="url(#softShadow)"><path d="M20 43h88"/><path d="M35 61h73"/><path d="M20 79h70"/><path d="M39 97h69"/></g>`;
  else if (normalized === "wind" || normalized === "vento") body = `<g fill="none" stroke="#2563eb" stroke-width="9" stroke-linecap="round" filter="url(#softShadow)"><path d="M18 43h62c13 0 13-20 0-20-7 0-11 4-13 9"/><path d="M18 65h84c14 0 14 22 0 22-8 0-13-5-15-11"/><path d="M18 87h49"/></g>`;
  return `<svg class="impianto-weather-svg" viewBox="0 0 128 128" aria-hidden="true">${defs}${body}</svg>`;
}

function formatImpiantoRainLine(entry = {}) {
  if (!entry.hasCurrentRain && !entry.hasNextHourRain && !entry.rainWindow) return "";
  const intensity = entry.rainIntensity ? ` • ${entry.rainIntensity}` : "";
  if (entry.rainWindow?.label) return `${entry.rainWindow.label}${intensity}`;
  return entry.rainIntensity ? `Adesso • ${entry.rainIntensity}` : "Adesso";
}

function getImpiantoWeatherAlertLine(entry = {}) {
  if (entry.civilProtectionAlert) return "⚠ Allerta meteo";
  if (Array.isArray(entry.messages) && entry.messages.some((message) => /temporale/i.test(message))) return "⚠ Temporali";
  return "";
}

function getImpiantoWeatherLineClass(line = "", index = 0) {
  const normalized = String(line || "").toLowerCase();
  if (normalized.includes("rischio") || normalized.includes("allerta") || normalized.includes("rinvio") || normalized.includes("attenzione") || normalized.includes("temporali")) return " is-risk";
  if (normalized.includes("pioggia") || normalized.includes("probabilità") || normalized.includes("prob.")) return " is-rain";
  if (normalized.includes("vento") || normalized.startsWith("💨") || normalized.includes("dato vento")) return " is-wind";
  if (index === 0) return " is-primary";
  return "";
}

function handleImpiantoWeatherRetryClick(event) {
  const button = event.target?.closest?.("[data-weather-retry]");
  if (!button) return;
  event.preventDefault();
  event.stopPropagation();
  const key = button.getAttribute("data-weather-retry") || button.closest("[data-weather-card]")?.getAttribute("data-weather-card") || "";
  const impianto = findImpiantoByWeatherKey(key);
  if (!impianto) return;
  button.disabled = true;
  refreshImpiantoWeatherStatus(impianto, { force: true }).finally(() => {
    button.disabled = false;
  });
}

function buildImpiantoWeatherCardInnerMarkup(state) {
  const level = escapeHTML(state.level);
  const iconType = state.iconType || "cloud";
  const display = state.display || { description: state.text || "Meteo", temperature: "--°", wind: "", rain: "" };
  const retryMarkup = state.canRetry && state.retryKey
    ? `<button type="button" class="impianto-weather-retry-btn" data-weather-retry="${escapeHTML(state.retryKey)}">Riprova</button>`
    : "";
  const windMarkup = display.wind
    ? `<span class="impianto-weather-wind"><svg viewBox="0 0 24 24" aria-hidden="true"><path d="M3 8h11a3 3 0 1 0-2.6-4.5"/><path d="M3 13h16a3 3 0 1 1-2.6 4.5"/><path d="M3 18h8"/></svg>${escapeHTML(display.wind)}</span>`
    : "";
  const rainMarkup = display.rain ? `<span class="impianto-weather-rain">${escapeHTML(display.rain)}</span>` : "";
  const alertMarkup = Array.isArray(display.alerts) ? display.alerts.slice(0, 2).map((line) => `<span class="impianto-weather-sfalcio-alert">${escapeHTML(line)}</span>`).join("") : "";
  const atexMarkup = state.showAtex && state.retryKey
    ? `<button type="button" class="impianto-weather-atex-btn" data-atex-procedure="${escapeHTML(state.retryKey)}" aria-label="Apri procedura sicurezza ATEX"><span aria-hidden="true">⚠️</span><span>ATEX</span></button>`
    : "";
  const safetyMarkup = "";
  const detailHint = state.canRetry ? "" : `<span class="impianto-weather-detail-hint">Tocca per dettaglio <span aria-hidden="true">›</span></span>`;
  return `<span class="impianto-weather-badge impianto-weather-card weather-${level}${state.compact ? " is-compact" : ""}" title="${escapeHTML(state.label)}" role="button" tabindex="0" aria-label="Apri dettaglio meteo impianto"><span class="impianto-weather-icon-shell weather-icon-${escapeHTML(iconType)}">${buildImpiantoWeatherIconSvg(iconType)}</span><span class="impianto-weather-copy"><span class="impianto-weather-description">${escapeHTML(display.description)}</span><span class="impianto-weather-temp">${escapeHTML(display.temperature)}</span>${windMarkup}${rainMarkup}${alertMarkup}${detailHint}${atexMarkup}${safetyMarkup}${retryMarkup}</span></span>`;
}

function buildImpiantoWeatherBadgeMarkup(impianto) {
  const key = getImpiantoWeatherCacheKey(impianto);
  const state = getImpiantoWeatherBadgeState(impianto);
  const feedbackText = state.feedback || "";
  return `<span class="impianto-weather-wrap" data-weather-card="${escapeHTML(key)}">${buildImpiantoWeatherCardInnerMarkup(state)}<small class="impianto-weather-updating${feedbackText ? "" : " hidden"}" data-weather-updating="${escapeHTML(key)}">${escapeHTML(feedbackText)}</small></span>`;
}


function openDettaglioMeteoImpianto(impianto) {
  if (!selectedCommessaId || !impianto) return;
  const key = buildImpiantoKey(impianto);
  window.location.hash = `commessa=${encodeURIComponent(selectedCommessaId)}&meteo=${encodeURIComponent(key)}`;
  applyRoute();
}

function closeDettaglioMeteoImpianto() {
  openImpiantiPage();
}

function getDettaglioMeteoImpiantoByKey(key) {
  const normalized = String(key || "").trim();
  return currentImpianti.find((item) => buildImpiantoKey(item) === normalized)
    || globalImpianti.find((item) => buildImpiantoKey(item) === normalized)
    || null;
}

function formatWeatherDetailValue(value, suffix = "") {
  if (!isPresentFiniteNumber(value)) return "-";
  return `${Math.round(Number(value))}${suffix}`;
}

function formatWeatherDetailDirection(slot = {}) {
  return slot.windDirectionLabel || slot.windDirection || getWindDirectionLabel(slot.wind_direction_10m ?? slot.windDirectionDegrees);
}

function getWeatherSlotWindKmh(slot = {}) {
  const raw = slot.wind_speed_10m ?? slot.windSpeedKmh ?? slot.windSpeed;
  return convertWindSpeedToKmh(raw, slot.wind_speed_unit || "km/h");
}

function getWeatherSlotGustKmh(slot = {}) {
  const raw = slot.wind_gusts_10m ?? slot.windGustKmh ?? slot.windGust;
  return convertWindSpeedToKmh(raw, slot.wind_gust_unit || "km/h");
}

function normalizeWeatherRiskLevel(level) {
  if (level === "alto" || level === "red") return "alto";
  if (level === "medio" || level === "yellow" || level === "orange") return "medio";
  return "basso";
}

function getWeatherRiskMeta(level) {
  return {
    basso: {
      className: "low",
      colorClass: "green",
      icon: "🟢",
      label: "BASSO",
      title: "RISCHIO OPERATIVO BASSO",
      finalTitle: "Lavoro regolare",
      finalText: "Condizioni compatibili con lavoro regolare. Mantenere i controlli ordinari su mezzi, area e operatori."
    },
    medio: {
      className: "medium",
      colorClass: "yellow",
      icon: "🟡",
      label: "MEDIO",
      title: "RISCHIO OPERATIVO MEDIO",
      finalTitle: "Procedere con cautela",
      finalText: "Procedere con cautela. Verificare terreno, vento e pioggia."
    },
    alto: {
      className: "high",
      colorClass: "red",
      icon: "🔴",
      label: "ALTO",
      title: "RISCHIO OPERATIVO ALTO",
      finalTitle: "Valutare sospensione attività",
      finalText: "Rischio alto. Valutare sospensione o rinvio finché temporali, raffiche, gelo o caldo intenso non sono terminati."
    }
  }[normalizeWeatherRiskLevel(level)];
}

function getWeatherSlotTemperature(slot = {}) {
  return Number(slot.temperature_2m ?? slot.temperature ?? slot.apparent_temperature);
}

function isThunderWeatherCode(code) {
  return NAVIGATION_WEATHER_THUNDER_CODES.has(Number(code));
}

function isFogOrLowVisibility(slot = {}) {
  const code = Number(slot.weather_code);
  return [45, 48].includes(code) || (Number(slot.visibility) > 0 && Number(slot.visibility) < 1000);
}

function isWetWeatherSlot(slot = {}) {
  return getPrecipitationAmount(slot) > 0 || Number(slot.precipitation_probability) >= 35 || NAVIGATION_WEATHER_RAIN_CODES.has(Number(slot.weather_code));
}

function hasFireRiskWeather(status = {}, slots = []) {
  const current = status.currentWeather || {};
  const temp = Math.max(Number(status.temperature) || -99, Number(current.temperature_2m) || -99, ...slots.slice(0, 6).map((slot) => getWeatherSlotTemperature(slot) || -99));
  const humidity = Number(current.relative_humidity_2m);
  const rainProbability = Number(status.rainProbability ?? status.currentRainProbability) || 0;
  return temp >= 35 && rainProbability < 20 && (!Number.isFinite(humidity) || humidity <= 35);
}

function getWeatherDetailRiskForSlot(slot = {}, status = {}) {
  const wind = getWeatherSlotWindKmh(slot) || 0;
  const gust = getWeatherSlotGustKmh(slot) || 0;
  const code = Number(slot.weather_code);
  const prob = Number(slot.precipitation_probability) || 0;
  const temp = getWeatherSlotTemperature(slot);
  const hasThunder = isThunderWeatherCode(code);
  const hasFrost = Number.isFinite(temp) && temp <= 0;
  const hasIntenseHeat = Number.isFinite(temp) && temp >= 35;
  const hasAlert = Boolean(status.civilProtectionAlert);
  if (prob > 65 || hasThunder || wind > 40 || gust > 55 || hasIntenseHeat || hasFrost || hasAlert) return "alto";
  if ((prob >= 35 && prob <= 65) || (wind >= 25 && wind <= 40) || (gust >= 35 && gust <= 55) || isWetWeatherSlot(slot) || (Number.isFinite(temp) && temp >= 30)) return "medio";
  return "basso";
}

function buildDettaglioMeteoRiskAnalysis(status = {}) {
  const slots = Array.isArray(status.forecastSlots) ? status.forecastSlots : [];
  const current = status.currentWeather || {};
  const nearSlots = slots.filter((slot) => Number(slot.timestamp) >= Date.now() - 15 * 60 * 1000 && Number(slot.timestamp) <= Date.now() + 3 * 60 * 60 * 1000);
  const currentWind = Number(status.windSpeedKmh ?? current.wind_speed_10m) || 0;
  const currentGust = Number(status.windGustKmh ?? current.wind_gusts_10m) || 0;
  const maxWind = Math.max(currentWind, ...nearSlots.map((slot) => getWeatherSlotWindKmh(slot) || 0));
  const maxGust = Math.max(currentGust, ...nearSlots.map((slot) => getWeatherSlotGustKmh(slot) || 0));
  const rainProbability = Math.max(Number(status.rainProbability ?? status.currentRainProbability) || 0, ...nearSlots.map((slot) => Number(slot.precipitation_probability) || 0));
  const thunderWithinWindow = nearSlots.some((slot) => isThunderWeatherCode(slot.weather_code)) || isThunderWeatherCode(current.weather_code);
  const wetGround = Boolean(status.hasCurrentRain || status.hasNextHourRain || Number(status.rainAmount) > 0 || nearSlots.some(isWetWeatherSlot));
  const temperatures = [Number(status.temperature), Number(current.temperature_2m), ...nearSlots.map(getWeatherSlotTemperature)].filter(Number.isFinite);
  const minTemperature = temperatures.length ? Math.min(...temperatures) : null;
  const maxTemperature = temperatures.length ? Math.max(...temperatures) : null;
  const frostRisk = Number.isFinite(minTemperature) && minTemperature <= 0;
  const hotModerate = Number.isFinite(maxTemperature) && maxTemperature >= 30;
  const intenseHeat = Number.isFinite(maxTemperature) && maxTemperature >= 35;
  const lowVisibility = nearSlots.some(isFogOrLowVisibility) || isFogOrLowVisibility(current);
  const fireRiskHigh = hasFireRiskWeather(status, nearSlots);
  const hasCivilProtectionAlert = Boolean(status.civilProtectionAlert);
  const slotRisks = nearSlots.map((slot) => getWeatherDetailRiskForSlot(slot, status));
  let level = "basso";
  if (rainProbability > 65 || thunderWithinWindow || maxWind > 40 || maxGust > 55 || intenseHeat || frostRisk || fireRiskHigh || hasCivilProtectionAlert || slotRisks.includes("alto")) level = "alto";
  else if ((rainProbability >= 35 && rainProbability <= 65) || (maxWind >= 25 && maxWind <= 40) || (maxGust >= 35 && maxGust <= 55) || wetGround || hotModerate || slotRisks.includes("medio")) level = "medio";
  return { level, rainProbability, thunderWithinWindow, wetGround, hotModerate, intenseHeat, frostRisk, lowVisibility, fireRiskHigh, hasCivilProtectionAlert, maxWind, maxGust, minTemperature, maxTemperature };
}

function buildSfalcioOperationalIndications(status = {}, slot = null) {
  const scopedStatus = slot ? { ...status, forecastSlots: [slot], rainProbability: slot.precipitation_probability ?? status.rainProbability, windSpeedKmh: getWeatherSlotWindKmh(slot) ?? status.windSpeedKmh, windGustKmh: getWeatherSlotGustKmh(slot) ?? status.windGustKmh, temperature: getWeatherSlotTemperature(slot) ?? status.temperature, currentWeather: { ...(status.currentWeather || {}), ...slot } } : status;
  const analysis = buildDettaglioMeteoRiskAnalysis(scopedStatus);
  const slots = slot ? [slot] : (Array.isArray(status.forecastSlots) ? status.forecastSlots : []);
  const current = scopedStatus.currentWeather || {};
  const messages = [];
  const add = (key, text, priority = 50) => {
    if (!messages.some((item) => item.key === key)) messages.push({ key, text, priority });
  };
  const rainProbability = Math.max(Number(scopedStatus.rainProbability ?? scopedStatus.currentRainProbability) || 0, ...slots.map((item) => Number(item.precipitation_probability) || 0));
  const wet = analysis.wetGround || rainProbability >= 35 || slots.some(isWetWeatherSlot);
  const maxWind = Math.max(Number(scopedStatus.windSpeedKmh) || 0, ...slots.map((item) => getWeatherSlotWindKmh(item) || 0));
  const maxGust = Math.max(Number(scopedStatus.windGustKmh) || 0, ...slots.map((item) => getWeatherSlotGustKmh(item) || 0));
  const thunder = analysis.thunderWithinWindow || slots.some((item) => isThunderWeatherCode(item.weather_code));
  const temp = Math.max(Number(scopedStatus.temperature) || -99, Number(current.temperature_2m) || -99, ...slots.map((item) => getWeatherSlotTemperature(item) || -99));
  if (thunder) {
    add("temporali-stop", "Temporali vicini: sospendere attività all’aperto e rientrare in zona sicura o mezzo aziendale.", 1);
    add("temporali-metallo", "Non usare attrezzi metallici in campo aperto; allontanarsi da alberi isolati, pali e recinzioni metalliche.", 2);
  }
  if (rainProbability > 65 || wet) {
    add("erba-bagnata", "Erba bagnata: ridurre velocità di lavoro e usare DPI antiscivolo.", 5);
    add("pendenze", "Evitare pendenze e scarpate se il terreno è scivoloso.", 6);
    add("rinvio-pioggia", "Valutare rinvio dello sfalcio se la pioggia è imminente; completare prima zone urgenti o vicine agli accessi.", 7);
  }
  if (wet && (rainProbability >= 50 || Number(scopedStatus.rainAmount) >= 1 || slots.some((item) => getPrecipitationAmount(item) >= 1))) {
    add("terreno-umido", "Terreno molto umido: rischio impantanamento mezzi. Evitare zone morbide e controllare accessi prima di entrare.", 10);
    add("solchi", "Usare trattore, rasaerba o mezzi solo se il terreno regge e non creare solchi profondi nelle aree verdi.", 11);
  }
  if (maxWind >= 25 || maxGust >= 35) {
    add("vento-decespugliatore", "Vento forte: attenzione uso decespugliatore e orientare il lavoro evitando proiezioni verso colleghi.", 15);
    add("raffiche-rami", "Raffiche: attenzione caduta rami. Non lavorare sotto alberature instabili o vicino a rami secchi.", 16);
    add("occhi-viso", "Proteggere occhi e viso da polvere, erba e materiale proiettato.", 17);
  }
  if (temp >= 30) {
    add("caldo", "Caldo intenso: aumentare pause e idratazione, usare cappello, acqua e pause all’ombra.", 20);
    add("ore-calde", "Evitare lavori pesanti nelle ore più calde e controllare affaticamento della squadra.", 21);
  }
  if (analysis.frostRisk || temp <= 2) {
    add("gelo", "Gelo o temperatura bassa: attenzione scivolamento su scale, tombini e rampe; usare guanti adeguati.", 25);
  }
  if (analysis.lowVisibility || slots.some(isFogOrLowVisibility)) {
    add("visibilita", "Nebbia o scarsa visibilità: aumentare distanza tra operatori, usare giubbino alta visibilità e lampeggianti mezzi.", 30);
  }
  if (analysis.fireRiskHigh) {
    add("incendi", "Rischio incendi alto: evitare scintille, controllare motori caldi e marmitte, segnalare subito fumo.", 35);
  }
  add("mezzi", "Prima di entrare verificare accessi e terreno; usare mezzi con cautela in pendenza e vicino a fossi.", 60);
  add("decespugliatore-dpi", "Controllare distanza di sicurezza tra operatori; usare visiera, cuffie, guanti e scarpe antinfortunistiche.", 61);
  add("materiale-proiettato", "Attenzione a sassi, vetri e materiale proiettato: non lavorare troppo vicino a persone, auto o vetrate.", 62);
  if (analysis.level === "basso") add("regolare", "Condizioni favorevoli allo sfalcio: procedere mantenendo controlli ordinari su terreno, mezzi e area di lancio materiali.", 40);
  return messages.sort((a, b) => a.priority - b.priority);
}

function buildWeatherDetailMetric(label, value, icon = "") {
  if (value === null || value === undefined || value === "") return "";
  return `<div class="weather-detail-metric weather-info-box"><span class="weather-detail-metric-label weather-info-label">${escapeHTML(icon)} ${escapeHTML(label)}</span><strong class="weather-info-value">${escapeHTML(String(value))}</strong></div>`;
}

function buildWeatherRadarUrl(impianto, coordinates) {
  if (coordinates && isPresentFiniteNumber(coordinates.lat) && isPresentFiniteNumber(coordinates.lon)) return `${METEO_3B_BASE_URL}?lat=${encodeURIComponent(coordinates.lat)}&lon=${encodeURIComponent(coordinates.lon)}`;
  const query = [impianto?.comune, impianto?.provincia].filter(Boolean).join(" ").trim() || "italia";
  return `${METEO_3B_BASE_URL}/${encodeURIComponent(query.toLowerCase().replace(/\s+/g, "-"))}`;
}

function renderDettaglioMeteoImpianto(impiantoKey) {
  const impianto = getDettaglioMeteoImpiantoByKey(impiantoKey);
  if (!impianto || !ui.impiantoWeatherDetailContent) return;
  const coordinates = getImpiantoNavigationCoordinates(impianto);
  const status = getCachedImpiantoWeatherStatus(impianto) || buildUnavailableImpiantoWeatherStatus(impianto, coordinates, null);
  const current = status.currentWeather || {};
  const slots = Array.isArray(status.forecastSlots) ? status.forecastSlots : [];
  const analysis = buildDettaglioMeteoRiskAnalysis(status);
  const riskMeta = getWeatherRiskMeta(analysis.level);
  ui.impiantoWeatherDetailSubtitle.textContent = `${impianto.denominazione || "Impianto"} • ${impianto.comune || "Comune non indicato"}`;
  const windValue = isPresentFiniteNumber(status.windSpeedKmh) ? `${Math.round(Number(status.windSpeedKmh))} km/h${status.windDirectionLabel ? ` ${status.windDirectionLabel}` : ""}` : "";
  const gustValue = isPresentFiniteNumber(status.windGustKmh ?? current.wind_gusts_10m) ? `${Math.round(Number(status.windGustKmh ?? current.wind_gusts_10m))} km/h` : "";
  const rainProbabilityValue = isPresentFiniteNumber(status.rainProbability) ? `${Math.round(Number(status.rainProbability))}%` : "";
  const rainWindowText = status.rainWindow?.label ? `Pioggia prevista ${status.rainWindow.label}` : "Pioggia non prevista a breve";
  const currentMetrics = [
    buildWeatherDetailMetric("Meteo", status.weatherState || status.description, "🌤️"),
    buildWeatherDetailMetric("Percepita", isPresentFiniteNumber(status.apparentTemperature) ? `${Math.round(Number(status.apparentTemperature))}°` : "", "🤚"),
    buildWeatherDetailMetric("Vento", windValue, "🌬️"),
    buildWeatherDetailMetric("Pioggia", rainProbabilityValue, "🌧️"),
    buildWeatherDetailMetric("Temperatura", isPresentFiniteNumber(status.temperature) ? `${Math.round(Number(status.temperature))}°` : "", "🌡️"),
    buildWeatherDetailMetric("Umidità", isPresentFiniteNumber(current.relative_humidity_2m) ? `${Math.round(Number(current.relative_humidity_2m))}%` : "", "💧"),
    buildWeatherDetailMetric("Raffiche", gustValue, "💨"),
    status.rainWindow?.label ? buildWeatherDetailMetric("Prossima pioggia", status.rainWindow.label, "⏱️") : ""
  ].filter(Boolean).join("");
  const forecastRows = slots.slice(0, 12).map((slot, index) => {
    const risk = getWeatherDetailRiskForSlot(slot, status);
    const meta = getWeatherRiskMeta(risk);
    const wind = getWeatherSlotWindKmh(slot);
    const gust = getWeatherSlotGustKmh(slot);
    const probability = isPresentFiniteNumber(slot.precipitation_probability) ? `${Math.round(Number(slot.precipitation_probability))}%` : "n/d";
    const slotWind = isPresentFiniteNumber(wind) ? `${Math.round(Number(wind))} km/h${formatWeatherDetailDirection(slot) ? ` ${formatWeatherDetailDirection(slot)}` : ""}` : "n/d";
    const slotGust = isPresentFiniteNumber(gust) ? `${Math.round(Number(gust))} km/h` : "n/d";
    const slotIndications = buildSfalcioOperationalIndications(status, slot).slice(0, 4).map((item) => `<li>${escapeHTML(item.text)}</li>`).join("");
    return `<details class="weather-detail-hour-card hourly-card risk-${meta.className}">
      <summary><span class="weather-detail-hour-time">${escapeHTML(formatWeatherSlotTime(slot.timestamp))}</span><span class="weather-detail-hour-icon">${escapeHTML(weatherCodeLabel(slot.weather_code).split(" ")[0] || "☁️")}</span><span class="weather-detail-hour-rain">🌧️ ${escapeHTML(probability)}</span>${isPresentFiniteNumber(wind) ? `<span class="weather-detail-hour-wind">🌬️ ${Math.round(Number(wind))}</span>` : ""}<span class="weather-detail-risk-badge risk-${meta.className}">${meta.label}</span></summary>
      <div class="weather-detail-hour-panel">
        <p><b>Meteo previsto:</b> ${escapeHTML(weatherCodeLabel(slot.weather_code))}${isPresentFiniteNumber(slot.temperature_2m) ? ` • ${Math.round(Number(slot.temperature_2m))}°` : ""}</p>
        <p><b>Probabilità pioggia:</b> ${escapeHTML(probability)} • <b>Vento:</b> ${escapeHTML(slotWind)} • <b>Raffiche:</b> ${escapeHTML(slotGust)}</p>
        <p><b>Rischio operativo:</b> ${meta.icon} ${meta.label}</p>
        <ul>${slotIndications}</ul>
      </div>
    </details>`;
  }).join("") || "<p class='muted'>Previsioni prossime ore non disponibili.</p>";
  const indications = buildSfalcioOperationalIndications(status);
  const indicationItems = indications.map((item) => `<li>${escapeHTML(item.text)}</li>`).join("");
  const radarUrl = buildWeatherRadarUrl(impianto, coordinates);
  const atexButtonKey = buildImpiantoKey(impianto) || impiantoKey;
  const atexActionMarkup = isCurrentCommessaInrete() && !isCurrentCommessaDepurazioneOrDiscariche() ? `
    <button type="button" class="weather-detail-atex-action" data-atex-procedure="${escapeHTML(atexButtonKey)}" aria-label="Apri istruzioni ATEX per questo impianto">
      <span class="weather-detail-atex-icon" aria-hidden="true">⚠️</span>
      <span><strong>ATTENZIONE ATEX</strong><small>Premi per istruzioni</small></span>
    </button>
  ` : "";
  const safetyActionMarkup = isCurrentCommessaDepurazioneOrDiscariche() ? `
    <button type="button" class="weather-detail-safety-action" data-impianto-safety="${escapeHTML(atexButtonKey)}" aria-label="Apri sicurezza impianto per questo impianto">
      <span class="weather-detail-atex-icon" aria-hidden="true">🦺</span>
      <span><strong>SICUREZZA IMPIANTO</strong><small>Premi per istruzioni</small></span>
    </button>
  ` : "";
  ui.impiantoWeatherDetailContent.innerHTML = `
    <article class="weather-detail-risk-summary risk-card risk-${riskMeta.className}">
      <strong>${riskMeta.icon} ${riskMeta.title}</strong>
      <div class="weather-detail-risk-grid">
        <span class="risk-row">${escapeHTML(rainWindowText)}</span>
        ${rainProbabilityValue ? `<span class="risk-row">Pioggia ${escapeHTML(rainProbabilityValue)}</span>` : ""}
        ${windValue ? `<span class="risk-row">Vento ${escapeHTML(windValue)}</span>` : ""}
        ${gustValue ? `<span class="risk-row">Raffiche ${escapeHTML(gustValue)}</span>` : ""}
      </div>
    </article>
    ${atexActionMarkup}${safetyActionMarkup}
    <article class="weather-detail-section"><h3>Meteo attuale</h3><div class="weather-detail-current-grid current-weather-grid">${currentMetrics || "<p class='muted'>Meteo attuale non disponibile.</p>"}</div></article>
    <article class="weather-detail-section"><h3>Previsioni prossime ore</h3><div class="weather-detail-timeline hourly-forecast-list" aria-label="Timeline previsioni prossime ore">${forecastRows}</div></article>
    <a class="weather-detail-radar-card" href="${escapeHTML(radarUrl)}" target="_blank" rel="noopener noreferrer" aria-label="Apri radar meteo">
      <span><b>🌦️ Radar meteo</b><small>Visualizza evoluzione precipitazioni</small></span><span aria-hidden="true">›</span>
    </a>
    <article class="weather-detail-section weather-detail-indications-card">
      <input class="weather-detail-indications-toggle" id="weather-detail-indications-toggle" type="checkbox">
      <div class="weather-detail-indications-head"><h3>Indicazioni operative</h3><label class="weather-detail-show-all" for="weather-detail-indications-toggle"><span class="show-more">Mostra tutte</span><span class="show-less">Mostra meno</span></label></div>
      <ul class="weather-detail-indications">${indicationItems}</ul>
    </article>
    <article class="weather-detail-risk-box risk-${riskMeta.colorClass}"><strong>${riskMeta.icon} ATTENZIONE</strong><p><b>${riskMeta.finalTitle}.</b> ${riskMeta.finalText}</p></article>
  `;
}

async function refreshDettaglioMeteoImpianto() {
  const route = parseCommessaHash();
  const impianto = getDettaglioMeteoImpiantoByKey(route.meteo);
  if (!impianto) return;
  ui.impiantoWeatherDetailFeedback.textContent = "Aggiornamento meteo in corso…";
  ui.impiantoWeatherDetailRefreshBtn.disabled = true;
  await refreshImpiantoWeatherStatus(impianto, { force: true });
  renderDettaglioMeteoImpianto(route.meteo);
  ui.impiantoWeatherDetailFeedback.textContent = "Meteo aggiornato.";
  ui.impiantoWeatherDetailRefreshBtn.disabled = false;
}

function findImpiantoByWeatherKey(key) {
  return currentImpianti.find((item) => getImpiantoWeatherCacheKey(item) === key)
    || globalImpianti.find((item) => getImpiantoWeatherCacheKey(item) === key)
    || (selectedGlobalImpianto && getImpiantoWeatherCacheKey(selectedGlobalImpianto) === key ? selectedGlobalImpianto : null)
    || (selectedImpiantoData && getImpiantoWeatherCacheKey(selectedImpiantoData) === key ? selectedImpiantoData : null);
}

function updateImpiantoWeatherBadgesInPlace() {
  document.querySelectorAll("[data-weather-card]").forEach((wrap) => {
    const key = wrap.getAttribute("data-weather-card") || "";
    const impianto = findImpiantoByWeatherKey(key);
    if (!impianto) return;
    const state = getImpiantoWeatherBadgeState(impianto);
    const feedbackText = state.feedback || "";
    wrap.innerHTML = `${buildImpiantoWeatherCardInnerMarkup(state)}<small class="impianto-weather-updating${feedbackText ? "" : " hidden"}" data-weather-updating="${escapeHTML(key)}">${escapeHTML(feedbackText)}</small>`;
  });
}

function scheduleImpiantoWeatherBadgeRender() {
  clearTimeout(impiantoWeatherRenderTimer);
  impiantoWeatherRenderTimer = setTimeout(() => {
    updateImpiantoWeatherBadgesInPlace();
  }, 120);
}

function formatWeatherAmount(value) {
  const amount = Number(value);
  if (!Number.isFinite(amount)) return "";
  return amount >= 10 ? String(Math.round(amount)) : amount.toFixed(1);
}

function isPresentFiniteNumber(value) {
  return value !== null && value !== undefined && value !== "" && Number.isFinite(Number(value));
}

function convertWindSpeedToKmh(value, unit = "km/h") {
  if (!isPresentFiniteNumber(value)) return null;
  const speed = Number(value);
  const normalizedUnit = String(unit || "km/h").toLowerCase();
  if (["km/h", "kmh", "kilometres per hour", "kilometers per hour"].includes(normalizedUnit)) return speed;
  if (["m/s", "ms", "meter/s", "metre/s"].includes(normalizedUnit)) return speed * 3.6;
  if (["mph"].includes(normalizedUnit)) return speed * 1.609344;
  if (["kn", "kt", "knot", "knots"].includes(normalizedUnit)) return speed * 1.852;
  return speed;
}

function getWindDirectionLabel(degrees) {
  if (!isPresentFiniteNumber(degrees)) return "";
  const directions = ["N", "NE", "E", "SE", "S", "SW", "W", "NW"];
  const normalized = ((Number(degrees) % 360) + 360) % 360;
  return directions[Math.round(normalized / 45) % directions.length];
}

function getWeatherSeriesUnit(data, section, field, fallback = "") {
  return data?.[`${section}_units`]?.[field] || fallback;
}

function formatWeatherSlotTime(timestamp) {
  if (!Number.isFinite(Number(timestamp))) return "";
  return new Date(Number(timestamp)).toLocaleTimeString("it-IT", { hour: "2-digit", minute: "2-digit" });
}

function getWeatherSlotIntervalMinutes(slots = []) {
  const timestamps = slots.map((slot) => Number(slot.timestamp)).filter(Number.isFinite).sort((a, b) => a - b);
  if (timestamps.length >= 2) return Math.max(15, Math.round((timestamps[1] - timestamps[0]) / 60000));
  return 60;
}

function buildRainWindowLabel(rainSlots = [], allSlots = []) {
  if (!rainSlots.length) return null;
  const sortedRainSlots = [...rainSlots].sort((a, b) => Number(a.timestamp) - Number(b.timestamp));
  const intervalMs = getWeatherSlotIntervalMinutes(allSlots.length ? allSlots : sortedRainSlots) * 60000;
  const firstWindow = [];
  sortedRainSlots.forEach((slot) => {
    if (!firstWindow.length) {
      firstWindow.push(slot);
      return;
    }
    const previous = firstWindow[firstWindow.length - 1];
    if (Number(slot.timestamp) - Number(previous.timestamp) <= intervalMs * 1.5) firstWindow.push(slot);
  });
  const start = Number(firstWindow[0]?.timestamp);
  const end = firstWindow.length > 1 ? Number(firstWindow[firstWindow.length - 1].timestamp) + intervalMs : start;
  if (!Number.isFinite(start) || !Number.isFinite(end)) return null;
  return {
    start,
    end,
    label: start === end ? formatWeatherSlotTime(start) : `${formatWeatherSlotTime(start)}–${formatWeatherSlotTime(end)}`
  };
}

function getRainIntensity(amount, probability, code) {
  const numericAmount = Number(amount) || 0;
  const numericProbability = Number(probability) || 0;
  const weatherCode = Number(code);
  if (numericAmount >= 5 || [65, 67, 82, 96, 99].includes(weatherCode)) return "Forte";
  if (numericAmount >= 2 || numericProbability >= 70 || [55, 63, 81].includes(weatherCode)) return "Moderata";
  if (numericAmount > 0 || numericProbability > 0 || NAVIGATION_WEATHER_RAIN_CODES.has(weatherCode)) return "Debole";
  return "";
}

function getImpiantoWeatherOperationMessage({ riskLevel, hasCurrentRain, hasNextHourRain, rainIntensity = "", importantWindKmh = null }) {
  if (riskLevel === "red") return "Valuta rinvio";
  if (Number(importantWindKmh) >= NAVIGATION_WEATHER_STRONG_WIND_KMH) return "Valuta rinvio";
  if (hasCurrentRain || hasNextHourRain) {
    if (/forte/i.test(rainIntensity)) return "Evita sfalcio";
    return "Attenzione terreno";
  }
  return "Ok lavoro";
}

function isRainWeatherSlot(slot) {
  const probability = Number(slot?.precipitation_probability) || 0;
  return getPrecipitationAmount(slot) > 0 || probability >= NAVIGATION_WEATHER_NEXT_HOUR_PROBABILITY || NAVIGATION_WEATHER_RAIN_CODES.has(Number(slot?.weather_code));
}

function getImpiantoWeatherForecastSlots(data) {
  const now = Date.now();
  const maxForecast = now + 12 * 60 * 60 * 1000;
  const hourlySlots = buildNavigationWeatherSlots(data?.hourly || {}, data, "hourly")
    .filter((slot) => slot.timestamp >= now - 60 * 60 * 1000 && slot.timestamp <= maxForecast);
  const minutelySlots = buildNavigationWeatherSlots(data?.minutely_15 || {}, data, "minutely_15")
    .filter((slot) => slot.timestamp >= now - 15 * 60 * 1000 && slot.timestamp <= maxForecast)
    .map((slot) => enrichWeatherSlotWithHourlyProbability(slot, hourlySlots));
  return minutelySlots.length ? minutelySlots : hourlySlots;
}

function enrichWeatherSlotWithHourlyProbability(slot, hourlySlots = []) {
  if (isPresentFiniteNumber(slot?.precipitation_probability) || !hourlySlots.length) return slot;
  const nearestHourly = hourlySlots
    .filter((hourlySlot) => isPresentFiniteNumber(hourlySlot.precipitation_probability))
    .sort((a, b) => Math.abs(Number(a.timestamp) - Number(slot.timestamp)) - Math.abs(Number(b.timestamp) - Number(slot.timestamp)))[0];
  if (!nearestHourly || Math.abs(Number(nearestHourly.timestamp) - Number(slot.timestamp)) > 60 * 60 * 1000) return slot;
  return { ...slot, precipitation_probability: nearestHourly.precipitation_probability };
}

function getCurrentWindDetails(data) {
  const current = data?.current || {};
  const currentUnit = getWeatherSeriesUnit(data, "current", "wind_speed_10m", "km/h");
  const speedKmh = convertWindSpeedToKmh(current.wind_speed_10m, currentUnit);
  const firstForecastWind = getImpiantoWeatherForecastSlots(data).find((slot) => isPresentFiniteNumber(slot.wind_speed_10m) || isPresentFiniteNumber(slot.wind_direction_10m));
  const directionDegrees = isPresentFiniteNumber(current.wind_direction_10m)
    ? Number(current.wind_direction_10m)
    : (isPresentFiniteNumber(firstForecastWind?.wind_direction_10m) ? Number(firstForecastWind.wind_direction_10m) : null);
  if (isPresentFiniteNumber(speedKmh)) {
    return {
      speedKmh,
      directionDegrees,
      directionLabel: getWindDirectionLabel(directionDegrees)
    };
  }
  if (!firstForecastWind || !isPresentFiniteNumber(firstForecastWind.wind_speed_10m)) return { speedKmh: null, directionDegrees: null, directionLabel: "" };
  const forecastSpeed = convertWindSpeedToKmh(firstForecastWind.wind_speed_10m, firstForecastWind.wind_speed_unit || getWeatherSeriesUnit(data, "hourly", "wind_speed_10m", "km/h"));
  const forecastDirection = isPresentFiniteNumber(firstForecastWind.wind_direction_10m) ? Number(firstForecastWind.wind_direction_10m) : null;
  return {
    speedKmh: forecastSpeed,
    directionDegrees: forecastDirection,
    directionLabel: getWindDirectionLabel(forecastDirection)
  };
}

function buildImpiantoWeatherOperationalDetails(weatherData, { hasCurrentRain, hasNextHourRain, riskLevel }) {
  const current = weatherData?.current || {};
  const allSlots = getImpiantoWeatherForecastSlots(weatherData);
  const rainSlots = allSlots.filter(isRainWeatherSlot);
  const rainWindow = buildRainWindowLabel(rainSlots.length ? rainSlots : (hasCurrentRain ? [{ ...current, timestamp: Date.now() }] : []), allSlots);
  const windowRainSlots = rainWindow
    ? rainSlots.filter((slot) => Number(slot.timestamp) >= rainWindow.start && Number(slot.timestamp) <= (rainWindow.end || rainWindow.start))
    : rainSlots;
  const relevantRainSlots = windowRainSlots.length ? windowRainSlots : rainSlots;
  const rainAmount = relevantRainSlots.reduce((sum, slot) => sum + getPrecipitationAmount(slot), hasCurrentRain ? getPrecipitationAmount(current) : 0);
  const rainProbability = relevantRainSlots.reduce((max, slot) => Math.max(max, Number(slot.precipitation_probability) || 0), 0);
  const strongestRainSlot = [...relevantRainSlots, ...(hasCurrentRain ? [current] : [])]
    .sort((a, b) => getPrecipitationAmount(b) - getPrecipitationAmount(a))[0] || null;
  const rainIntensity = hasCurrentRain || hasNextHourRain
    ? getRainIntensity(Math.max(rainAmount, getPrecipitationAmount(strongestRainSlot)), rainProbability, strongestRainSlot?.weather_code ?? current.weather_code)
    : "";
  const wind = getCurrentWindDetails(weatherData);
  const importantWindKmh = isPresentFiniteNumber(wind.speedKmh) && Number(wind.speedKmh) >= NAVIGATION_WEATHER_STRONG_WIND_KMH ? Number(wind.speedKmh) : null;
  const weatherPartial = (hasCurrentRain || hasNextHourRain || riskLevel !== "green") && ((!rainSlots.length && !hasCurrentRain) || (!isPresentFiniteNumber(rainProbability) && !isPresentFiniteNumber(rainAmount)));
  return {
    rainWindow,
    rainIntensity,
    rainAmount: isPresentFiniteNumber(rainAmount) ? rainAmount : null,
    rainProbability: isPresentFiniteNumber(rainProbability) && rainProbability > 0 ? rainProbability : null,
    windSpeedKmh: wind.speedKmh,
    windDirectionDegrees: wind.directionDegrees,
    windDirectionLabel: wind.directionLabel,
    importantWindKmh,
    operationMessage: getImpiantoWeatherOperationMessage({ riskLevel, hasCurrentRain, hasNextHourRain, rainIntensity, importantWindKmh }),
    weatherPartial: Boolean(weatherPartial)
  };
}

function buildImpiantoWeatherStatus(impianto, weatherData, civilProtectionAlert = null) {
  const current = weatherData?.current || {};
  const nextHourSlots = getNavigationWeatherNextHourSlots(weatherData);
  const messages = buildNavigationWeatherMessages(weatherData);
  const coordinates = getImpiantoNavigationCoordinates(impianto);
  const currentCode = Number(current.weather_code);
  const currentPrecipitation = getPrecipitationAmount(current);
  const hasCurrentRain = currentPrecipitation > 0 || NAVIGATION_WEATHER_RAIN_CODES.has(currentCode);
  const hasNextHourRain = nextHourSlots.some((slot) => {
    const probability = Number(slot.precipitation_probability) || 0;
    return getPrecipitationAmount(slot) > 0 || probability >= NAVIGATION_WEATHER_NEXT_HOUR_PROBABILITY || NAVIGATION_WEATHER_RAIN_CODES.has(Number(slot.weather_code));
  });
  const civilLevel = civilProtectionAlert?.level || "green";
  const hasCivilProtectionAlert = ALERT_LEVEL_META[civilLevel]?.rank > 0;
  let riskLevel = "green";
  if (hasCivilProtectionAlert && ALERT_LEVEL_META[civilLevel]?.rank >= ALERT_LEVEL_META.orange.rank) riskLevel = "red";
  else if (messages.some((message) => /rischio meteo rilevante|temporale|vento forte/i.test(message))) riskLevel = "red";
  else if (messages.length || hasCivilProtectionAlert || hasCurrentRain || hasNextHourRain) riskLevel = "yellow";

  const alertMessages = [];
  if (hasCivilProtectionAlert) alertMessages.push(formatCivilProtectionNavigationMessage(civilProtectionAlert));
  alertMessages.push(...messages);
  const shortAlertText = [...new Set(alertMessages)].slice(0, 5).join("; ") || "Meteo OK";
  const badgeLabel = riskLevel === "red"
    ? (hasCivilProtectionAlert ? "Allerta meteo" : "Rischio")
    : riskLevel === "yellow"
      ? (hasNextHourRain ? "Pioggia prevista" : hasCurrentRain ? "Pioggia in corso" : "Attenzione")
      : "Meteo OK";
  const weatherDetails = buildImpiantoWeatherOperationalDetails(weatherData, { hasCurrentRain, hasNextHourRain, riskLevel });
  const currentRainProbability = getCurrentNavigationRainProbability(weatherData);
  const rainProbability = weatherDetails.rainProbability ?? currentRainProbability ?? getMaxNavigationRainProbability(weatherData);
  const currentGustKmh = convertWindSpeedToKmh(current.wind_gusts_10m, getWeatherSeriesUnit(weatherData, "current", "wind_gusts_10m", "km/h"));
  const syntheticState = getSyntheticImpiantoWeatherState({
    riskLevel,
    hasCivilProtectionAlert,
    currentCode,
    hasCurrentRain,
    hasNextHourRain,
    currentWind: weatherDetails.windSpeedKmh ?? 0,
    currentGust: currentGustKmh ?? 0
  });
  const description = badgeLabel === "Meteo OK" ? weatherCodeLabel(current.weather_code).replace(/^[^A-Za-zÀ-ÿ]+\s*/, "") : shortAlertText;

  return {
    impiantoKey: getImpiantoWeatherCacheKey(impianto),
    lat: coordinates?.lat ?? null,
    lon: coordinates?.lon ?? null,
    coordinateKey: getImpiantoWeatherCoordinateKey(coordinates),
    syntheticState,
    weatherState: weatherCodeLabel(current.weather_code),
    temperature: Number.isFinite(Number(current.temperature_2m)) ? Number(current.temperature_2m) : null,
    apparentTemperature: Number.isFinite(Number(current.apparent_temperature)) ? Number(current.apparent_temperature) : null,
    rainProbability,
    currentRainProbability,
    description,
    hasCurrentRain,
    hasNextHourRain,
    civilProtectionAlert: hasCivilProtectionAlert ? civilProtectionAlert : null,
    riskLevel,
    alertText: shortAlertText,
    badgeLabel,
    rainWindow: weatherDetails.rainWindow,
    rainIntensity: weatherDetails.rainIntensity,
    rainAmount: weatherDetails.rainAmount,
    windSpeedKmh: weatherDetails.windSpeedKmh,
    windDirectionDegrees: weatherDetails.windDirectionDegrees,
    windDirectionLabel: weatherDetails.windDirectionLabel,
    windGustKmh: currentGustKmh,
    importantWindKmh: weatherDetails.importantWindKmh,
    operationMessage: weatherDetails.operationMessage,
    weatherPartial: weatherDetails.weatherPartial,
    messages: [...new Set(alertMessages)].slice(0, 5),
    currentWeather: current,
    forecastSlots: getImpiantoWeatherForecastSlots(weatherData).slice(0, 12),
    weatherDataUpdatedAt: Date.now(),
    updatedAt: Date.now()
  };
}

function getMaxNavigationRainProbability(data) {
  const values = Array.isArray(data?.hourly?.precipitation_probability) ? data.hourly.precipitation_probability : [];
  const numeric = values.map((value) => Number(value)).filter(Number.isFinite);
  return numeric.length ? Math.max(...numeric) : null;
}

function getCurrentNavigationRainProbability(data) {
  const current = data?.current || {};
  if (isPresentFiniteNumber(current.precipitation_probability)) return Number(current.precipitation_probability);
  const now = Date.now();
  const nearestSlot = buildNavigationWeatherSlots(data?.hourly || {}, data, "hourly")
    .filter((slot) => isPresentFiniteNumber(slot.precipitation_probability))
    .sort((a, b) => Math.abs(Number(a.timestamp) - now) - Math.abs(Number(b.timestamp) - now))[0];
  if (!nearestSlot || Math.abs(Number(nearestSlot.timestamp) - now) > 90 * 60 * 1000) return null;
  return Number(nearestSlot.precipitation_probability);
}

function getSyntheticImpiantoWeatherState({ riskLevel, hasCivilProtectionAlert, currentCode, hasCurrentRain, hasNextHourRain, currentWind, currentGust }) {
  if (riskLevel === "red" || hasCivilProtectionAlert) return "allerta";
  if (NAVIGATION_WEATHER_THUNDER_CODES.has(currentCode)) return "temporale";
  if (currentWind >= NAVIGATION_WEATHER_STRONG_WIND_KMH || currentGust >= NAVIGATION_WEATHER_STRONG_GUST_KMH) return "vento";
  if (hasCurrentRain || hasNextHourRain) return "pioggia";
  return "ok";
}


async function getWeatherForPlantFromMainSource(plant) {
  const coordinates = getImpiantoNavigationCoordinates(plant);
  if (!coordinates) throw new Error("Coordinate impianto non disponibili");
  const payload = await fetchWeatherForecast(coordinates, { operational: true, cache: "no-store" });
  return buildImpiantoWeatherStatus(plant, payload, null);
}

async function getPlantWeather(plant) {
  const coordinates = getImpiantoNavigationCoordinates(plant);
  if (!coordinates) {
    return {
      impiantoKey: getImpiantoWeatherCacheKey(plant),
      lat: null,
      lon: null,
      coordinateKey: "",
      syntheticState: "ok",
      weatherState: "Coordinate mancanti",
      description: "Coordinate mancanti",
      riskLevel: "unavailable",
      alertText: "Coordinate mancanti",
      badgeLabel: "Coordinate mancanti",
      iconType: "cloud",
      weatherPartial: true,
      canRetry: false,
      messages: [],
      updatedAt: Date.now()
    };
  }

  try {
    return await getWeatherForPlantFromMainSource(plant);
  } catch (error) {
    console.error("Meteo impianto non disponibile dalla fonte principale:", error);
    throw error;
  }
}

function buildUnavailableImpiantoWeatherStatus(impianto, coordinates, error = null) {
  return {
    impiantoKey: getImpiantoWeatherCacheKey(impianto),
    lat: coordinates?.lat ?? null,
    lon: coordinates?.lon ?? null,
    coordinateKey: getImpiantoWeatherCoordinateKey(coordinates),
    syntheticState: "ok",
    weatherState: "temporaneamente non disponibile",
    temperature: null,
    apparentTemperature: null,
    rainProbability: null,
    currentRainProbability: null,
    description: "Meteo temporaneamente non disponibile",
    hasCurrentRain: false,
    hasNextHourRain: false,
    civilProtectionAlert: null,
    riskLevel: "unavailable",
    alertText: "Meteo temporaneamente non disponibile",
    badgeLabel: "Meteo temporaneamente non disponibile",
    iconType: "cloud",
    rainWindow: null,
    rainIntensity: "",
    rainAmount: null,
    windSpeedKmh: null,
    windDirectionDegrees: null,
    windDirectionLabel: "",
    importantWindKmh: null,
    operationMessage: "",
    weatherPartial: true,
    canRetry: true,
    messages: [],
    errorMessage: error?.message || String(error || ""),
    updatedAt: Date.now()
  };
}

async function refreshImpiantoWeatherStatus(impianto, { force = false } = {}) {
  ensureImpiantoWeatherPersistentCacheLoaded();
  const key = getImpiantoWeatherCacheKey(impianto);
  const coordinates = getImpiantoNavigationCoordinates(impianto);
  const coordinateKey = getImpiantoWeatherCoordinateKey(coordinates);
  if (!key || !coordinates) return null;
  const cached = getCachedImpiantoWeatherStatus(impianto);
  const coordinateCached = coordinateKey ? impiantoWeatherCoordinateCache.get(coordinateKey) : null;
  if (!force && isImpiantoWeatherCacheFresh(cached)) return cached;
  if (!force && isImpiantoWeatherCacheFresh(coordinateCached)) {
    const cloned = cloneImpiantoWeatherStatusForImpianto(coordinateCached, impianto, coordinates);
    saveImpiantoWeatherStatus(cloned, impianto);
    return cloned;
  }
  if (impiantoWeatherPendingKeys.has(key) || (coordinateKey && impiantoWeatherPendingKeys.has(coordinateKey))) return cached || coordinateCached || null;
  impiantoWeatherPendingKeys.add(key);
  if (coordinateKey) impiantoWeatherPendingKeys.add(coordinateKey);
  scheduleImpiantoWeatherBadgeRender();
  try {
    const status = await getPlantWeather(impianto);
    setImpiantoWeatherFeedback(impianto, "");
    saveImpiantoWeatherStatus(status, impianto);
    return status;
  } catch (error) {
    const fallback = buildUnavailableImpiantoWeatherStatus(impianto, coordinates, error);
    setImpiantoWeatherFeedback(impianto, "");
    saveImpiantoWeatherStatus(fallback, impianto);
    return fallback;
  } finally {
    impiantoWeatherPendingKeys.delete(key);
    if (coordinateKey) impiantoWeatherPendingKeys.delete(coordinateKey);
    scheduleImpiantoWeatherBadgeRender();
  }
}

function cloneImpiantoWeatherStatusForImpianto(entry, impianto, coordinates = getImpiantoNavigationCoordinates(impianto)) {
  return {
    ...entry,
    impiantoKey: getImpiantoWeatherCacheKey(impianto),
    lat: coordinates?.lat ?? entry.lat ?? null,
    lon: coordinates?.lon ?? entry.lon ?? null,
    coordinateKey: getImpiantoWeatherCoordinateKey(coordinates) || entry.coordinateKey || ""
  };
}

function saveImpiantoWeatherStatus(status, impianto = null) {
  if (!status?.impiantoKey) return;
  impiantoWeatherStatusCache.set(status.impiantoKey, status);
  if (status.coordinateKey) impiantoWeatherCoordinateCache.set(status.coordinateKey, status);
  if (impianto) writePlantWeatherLocalCache(impianto, status);
  persistImpiantoWeatherCache();
}

function preloadImpiantiWeather(impianti = [], { limit = IMPIANTO_WEATHER_REFRESH_LIMIT, preferNearest = true } = {}) {
  ensureImpiantoWeatherPersistentCacheLoaded();
  const byCoordinate = new Map();
  impianti
    .filter((impianto) => getImpiantoNavigationCoordinates(impianto))
    .filter((impianto) => !isImpiantoWeatherCacheFresh(getCachedImpiantoWeatherStatus(impianto)))
    .sort((a, b) => (preferNearest ? distanceFromUser(a) - distanceFromUser(b) : 0))
    .forEach((impianto) => {
      const coordinateKey = getImpiantoWeatherCoordinateKey(getImpiantoNavigationCoordinates(impianto));
      if (!coordinateKey || byCoordinate.has(coordinateKey)) return;
      byCoordinate.set(coordinateKey, impianto);
    });
  const candidates = Array.from(byCoordinate.values()).slice(0, limit);
  if (!candidates.length) return;
  clearTimeout(impiantoWeatherRefreshTimer);
  impiantoWeatherRefreshTimer = setTimeout(async () => {
    for (let index = 0; index < candidates.length; index += IMPIANTO_WEATHER_BATCH_SIZE) {
      const batch = candidates.slice(index, index + IMPIANTO_WEATHER_BATCH_SIZE);
      await Promise.allSettled(batch.map((impianto, batchIndex) => new Promise((resolve) => {
        setTimeout(() => resolve(refreshImpiantoWeatherStatus(impianto)), batchIndex * 250);
      })));
      await new Promise((resolve) => setTimeout(resolve, 500));
    }
  }, 250);
}

function preloadCommessaWeatherForVisibleImpianti() {
  if (!selectedCommessaId || !currentImpianti.length) return;
  preloadImpiantiWeather(currentImpianti, { limit: currentImpianti.length, preferNearest: true });
}


function setCommessaWeatherRefreshStatus(message, state = "") {
  if (!ui.commessaWeatherRefreshStatus) return;
  ui.commessaWeatherRefreshStatus.textContent = message || "";
  ui.commessaWeatherRefreshStatus.dataset.state = state || "";
}

function updateCommessaWeatherRefreshButtonState() {
  if (!ui.commessaWeatherRefreshBtn) return;
  const disabled = commessaWeatherManualRefreshInFlight || !selectedCommessaId || !currentImpianti.length;
  ui.commessaWeatherRefreshBtn.disabled = disabled;
  ui.commessaWeatherRefreshBtn.classList.toggle("is-loading", commessaWeatherManualRefreshInFlight);
  ui.commessaWeatherRefreshBtn.setAttribute("aria-busy", String(commessaWeatherManualRefreshInFlight));
}

async function refreshSelectedCommessaWeather() {
  if (commessaWeatherManualRefreshInFlight || !selectedCommessaId) return;
  const commessaId = selectedCommessaId;
  const impianti = [...currentImpianti];
  const candidates = impianti.filter((impianto) => getImpiantoNavigationCoordinates(impianto));

  if (!candidates.length) {
    setCommessaWeatherRefreshStatus("Nessun dato meteo disponibile", "empty");
    return;
  }

  commessaWeatherManualRefreshInFlight = true;
  updateCommessaWeatherRefreshButtonState();
  setCommessaWeatherRefreshStatus("Aggiornamento meteo…", "loading");

  try {
    clearTimeout(impiantoWeatherRefreshTimer);
    const results = [];
    for (let index = 0; index < candidates.length; index += IMPIANTO_WEATHER_BATCH_SIZE) {
      if (selectedCommessaId !== commessaId) return;
      const batch = candidates.slice(index, index + IMPIANTO_WEATHER_BATCH_SIZE);
      const settled = await Promise.allSettled(batch.map((impianto, batchIndex) => new Promise((resolve) => {
        setTimeout(() => resolve(refreshImpiantoWeatherStatus(impianto, { force: true })), batchIndex * 250);
      })));
      results.push(...settled);
      if (index + IMPIANTO_WEATHER_BATCH_SIZE < candidates.length) {
        await new Promise((resolve) => setTimeout(resolve, 500));
      }
    }

    if (selectedCommessaId !== commessaId) return;
    const statuses = results
      .filter((result) => result.status === "fulfilled" && result.value)
      .map((result) => result.value);
    const hasAvailableWeather = statuses.some((status) => status.riskLevel !== "unavailable");

    if (!statuses.length || !hasAvailableWeather) {
      setCommessaWeatherRefreshStatus("Nessun dato meteo disponibile", "empty");
    } else {
      setCommessaWeatherRefreshStatus("Meteo commessa aggiornato", "success");
    }
    scheduleImpiantoWeatherBadgeRender();
  } catch (error) {
    console.error("Errore aggiornamento meteo commessa:", error);
    if (selectedCommessaId === commessaId) setCommessaWeatherRefreshStatus("Errore aggiornamento meteo", "error");
  } finally {
    commessaWeatherManualRefreshInFlight = false;
    if (selectedCommessaId === commessaId) updateCommessaWeatherRefreshButtonState();
  }
}

function getVisibleMapImpianti(targetMap = map, source = currentImpianti) {
  const bounds = targetMap?.getBounds?.();
  if (!bounds) return source;
  return source.filter((impianto) => {
    const coordinates = getImpiantoNavigationCoordinates(impianto);
    return coordinates && bounds.contains([coordinates.lat, coordinates.lon]);
  });
}

async function confirmNavigationWeatherIfNeeded(impianto) {
  const coordinates = getImpiantoNavigationCoordinates(impianto);
  if (!coordinates) return true;

  const cachedStatus = getCachedImpiantoWeatherStatus(impianto);
  if (!cachedStatus || cachedStatus.riskLevel === "unavailable") return true;
  if (!["yellow", "red"].includes(cachedStatus.riskLevel)) return true;
  const messages = cachedStatus.messages?.length ? cachedStatus.messages : [cachedStatus.alertText || cachedStatus.badgeLabel || "attenzione meteo"];
  return await showNavigationWeatherWarning(messages);
}

async function getNavigationWeatherWarning(impianto, coordinates) {
  const weatherData = await fetchImpiantoNavigationWeather(coordinates);
  const messages = buildNavigationWeatherMessages(weatherData);
  const civilProtectionAlert = await getOfficialCivilProtectionAlertForNavigation(coordinates).catch((error) => {
    console.warn("Allerta Protezione Civile non disponibile per navigazione:", error);
    return null;
  });

  if (civilProtectionAlert && ALERT_LEVEL_META[civilProtectionAlert.level || "green"]?.rank > 0) {
    messages.unshift(formatCivilProtectionNavigationMessage(civilProtectionAlert));
  }

  return {
    impiantoKey: buildImpiantoKey(impianto),
    coordinates,
    messages: [...new Set(messages)].slice(0, 5)
  };
}

async function fetchImpiantoNavigationWeather({ lat, lon }) {
  return fetchWeatherForecast({ lat, lon }, { operational: true, cache: "no-store" });
}

function validateImpiantoWeatherPayload(data, provider = "meteo") {
  if (!data || typeof data !== "object") throw new Error(`${provider}: risposta non valida`);
  const hasCurrentWeather = isPresentFiniteNumber(data.current?.weather_code) || isPresentFiniteNumber(data.current?.precipitation);
  const hasHourlyProbability = Array.isArray(data.hourly?.precipitation_probability) && data.hourly.precipitation_probability.some(isPresentFiniteNumber);
  const hasHourlyWeather = Array.isArray(data.hourly?.weather_code) && data.hourly.weather_code.some(isPresentFiniteNumber);
  if (!hasCurrentWeather && !hasHourlyProbability && !hasHourlyWeather) throw new Error(`${provider}: campi meteo assenti`);
}

function buildNavigationWeatherMessages(data) {
  const messages = [];
  const current = data?.current || {};
  const currentPrecipitation = getPrecipitationAmount(current);
  const currentCode = Number(current.weather_code);
  const currentWind = convertWindSpeedToKmh(current.wind_speed_10m, getWeatherSeriesUnit(data, "current", "wind_speed_10m", "km/h")) ?? 0;
  const currentGust = convertWindSpeedToKmh(current.wind_gusts_10m, getWeatherSeriesUnit(data, "current", "wind_gusts_10m", "km/h")) ?? 0;
  const nextHourSlots = getNavigationWeatherNextHourSlots(data);

  if (NAVIGATION_WEATHER_THUNDER_CODES.has(currentCode)) {
    messages.push("Temporale attivo nella zona");
  }
  if (NAVIGATION_WEATHER_THUNDER_CODES.has(getFirstWeatherCode(nextHourSlots, NAVIGATION_WEATHER_THUNDER_CODES))) {
    messages.push(`Temporale previsto ${formatNavigationWeatherEta(nextHourSlots, NAVIGATION_WEATHER_THUNDER_CODES)}`);
  }

  if ((currentPrecipitation > 0 && currentPrecipitation < NAVIGATION_WEATHER_LIGHT_RAIN_MAX_MM && NAVIGATION_WEATHER_RAIN_CODES.has(currentCode)) || [51, 53, 61, 80].includes(currentCode)) {
    messages.push("Pioggia debole in corso");
  }

  const firstRainSlot = nextHourSlots.find((slot) => {
    const precipitation = getPrecipitationAmount(slot);
    const probability = Number(slot.precipitation_probability) || 0;
    const code = Number(slot.weather_code);
    return precipitation > 0 || probability >= NAVIGATION_WEATHER_NEXT_HOUR_PROBABILITY || NAVIGATION_WEATHER_RAIN_CODES.has(code);
  });
  if (firstRainSlot) {
    messages.push(`Pioggia prevista ${formatNavigationWeatherSlotEta(firstRainSlot)}`);
  }

  if (currentWind >= NAVIGATION_WEATHER_STRONG_WIND_KMH || currentGust >= NAVIGATION_WEATHER_STRONG_GUST_KMH) {
    messages.push("Vento forte attivo nella zona");
  }
  const firstWindSlot = nextHourSlots.find((slot) => (convertWindSpeedToKmh(slot.wind_speed_10m, slot.wind_speed_unit) ?? 0) >= NAVIGATION_WEATHER_STRONG_WIND_KMH || (convertWindSpeedToKmh(slot.wind_gusts_10m, slot.wind_speed_unit) ?? 0) >= NAVIGATION_WEATHER_STRONG_GUST_KMH);
  if (firstWindSlot) {
    messages.push(`Vento forte previsto ${formatNavigationWeatherSlotEta(firstWindSlot)}`);
  }

  const relevantRisk = getRelevantNavigationWeatherRisk(current, nextHourSlots);
  if (relevantRisk) messages.push(relevantRisk);

  return messages;
}

function getNavigationWeatherNextHourSlots(data) {
  const now = Date.now();
  const nextHour = now + 60 * 60 * 1000;
  const minutelySlots = buildNavigationWeatherSlots(data?.minutely_15 || {}, data, "minutely_15").filter((slot) => slot.timestamp >= now - 15 * 60 * 1000 && slot.timestamp <= nextHour);
  if (minutelySlots.length) return minutelySlots;
  return buildNavigationWeatherSlots(data?.hourly || {}, data, "hourly").filter((slot) => slot.timestamp >= now - 60 * 60 * 1000 && slot.timestamp <= nextHour);
}

function buildNavigationWeatherSlots(series, data = null, section = "hourly") {
  const times = Array.isArray(series?.time) ? series.time : [];
  const windSpeedUnit = data ? getWeatherSeriesUnit(data, section, "wind_speed_10m", "km/h") : "km/h";
  return times.map((time, idx) => ({
    time,
    timestamp: new Date(time).getTime(),
    precipitation: series.precipitation?.[idx],
    precipitation_probability: series.precipitation_probability?.[idx],
    rain: series.rain?.[idx],
    showers: series.showers?.[idx],
    weather_code: series.weather_code?.[idx],
    wind_speed_10m: series.wind_speed_10m?.[idx],
    wind_speed_unit: windSpeedUnit,
    wind_direction_10m: series.wind_direction_10m?.[idx],
    wind_gusts_10m: series.wind_gusts_10m?.[idx],
    visibility: series.visibility?.[idx]
  })).filter((slot) => Number.isFinite(slot.timestamp));
}

function getPrecipitationAmount(values) {
  return Math.max(Number(values?.precipitation) || 0, Number(values?.rain) || 0, Number(values?.showers) || 0);
}

function getFirstWeatherCode(slots, codeSet) {
  const slot = slots.find((item) => codeSet.has(Number(item.weather_code)));
  return slot ? Number(slot.weather_code) : NaN;
}

function formatNavigationWeatherEta(slots, codeSet) {
  const slot = slots.find((item) => codeSet.has(Number(item.weather_code)));
  return slot ? formatNavigationWeatherSlotEta(slot) : "nella prossima ora";
}

function formatNavigationWeatherSlotEta(slot) {
  const minutes = Math.max(0, Math.round((slot.timestamp - Date.now()) / 60000));
  if (minutes <= 5) return "entro pochi minuti";
  if (minutes >= 55) return "entro circa 1 ora";
  return `entro ${minutes} minuti`;
}

function getRelevantNavigationWeatherRisk(current, nextHourSlots) {
  const currentPrecipitation = getPrecipitationAmount(current);
  const currentCode = Number(current?.weather_code);
  if (currentPrecipitation >= NAVIGATION_WEATHER_RELEVANT_RAIN_MM || [65, 82, 96, 99].includes(currentCode)) {
    return "Rischio meteo rilevante: precipitazioni intense attive";
  }
  const severeSlot = nextHourSlots.find((slot) => getPrecipitationAmount(slot) >= NAVIGATION_WEATHER_RELEVANT_RAIN_MM || [65, 82, 96, 99].includes(Number(slot.weather_code)));
  if (severeSlot) return `Rischio meteo rilevante previsto ${formatNavigationWeatherSlotEta(severeSlot)}`;
  const lowVisibilitySlot = nextHourSlots.find((slot) => Number(slot.visibility) > 0 && Number(slot.visibility) < 500);
  if (lowVisibilitySlot) return `Rischio meteo rilevante: visibilità ridotta ${formatNavigationWeatherSlotEta(lowVisibilitySlot)}`;
  return "";
}

async function getOfficialCivilProtectionAlertForNavigation(coordinates) {
  const target = { lat: coordinates.lat, lon: coordinates.lon, source: "impianto" };
  const region = await reverseGeocodeRegion(target).catch(() => "");
  const officialText = await fetchCivilProtectionOfficialText().catch(() => "");
  const officialAlert = parseCivilProtectionAlertText(officialText, region);
  return {
    ...officialAlert,
    region,
    url: CIVIL_PROTECTION_ALERT_PAGE,
    label: officialAlert.label || "Nessuna allerta"
  };
}

function formatCivilProtectionNavigationMessage(alert) {
  const levelText = {
    yellow: "gialla",
    orange: "arancione",
    red: "rossa"
  }[alert?.level] || "attiva";
  const phenomenon = alert?.phenomenon && alert.phenomenon !== "Protezione Civile" ? ` per ${alert.phenomenon}` : "";
  return `Allerta ${levelText} Protezione Civile${phenomenon}`;
}

function showNavigationWeatherWarning(messages) {
  const summary = messages.join("; ");
  if (!ui.navigationWeatherWarningModal || !ui.navigationWeatherWarningList) {
    return Promise.resolve(window.confirm(`⚠️ Nell’area dell’impianto è previsto: ${summary}.
Vuoi continuare la navigazione?`));
  }

  const summaryEl = document.getElementById("navigation-weather-warning-summary");
  if (summaryEl) summaryEl.textContent = summary;
  ui.navigationWeatherWarningList.innerHTML = messages.map((message) => `<li>${escapeHTML(message)}</li>`).join("");
  ui.navigationWeatherWarningModal.classList.remove("hidden");
  ui.navigationWeatherWarningModal.setAttribute("aria-hidden", "false");
  ui.navigationWeatherContinueBtn?.focus();

  return new Promise((resolve) => {
    const cleanup = (confirmed) => {
      ui.navigationWeatherWarningModal.classList.add("hidden");
      ui.navigationWeatherWarningModal.setAttribute("aria-hidden", "true");
      ui.navigationWeatherContinueBtn?.removeEventListener("click", onContinue);
      ui.navigationWeatherCancelBtn?.removeEventListener("click", onCancel);
      ui.navigationWeatherWarningModal?.removeEventListener("click", onBackdrop);
      document.removeEventListener("keydown", onKeydown);
      resolve(confirmed);
    };
    const onContinue = () => cleanup(true);
    const onCancel = () => cleanup(false);
    const onBackdrop = (event) => {
      if (event.target === ui.navigationWeatherWarningModal) cleanup(false);
    };
    const onKeydown = (event) => {
      if (event.key === "Escape") cleanup(false);
    };

    ui.navigationWeatherContinueBtn?.addEventListener("click", onContinue);
    ui.navigationWeatherCancelBtn?.addEventListener("click", onCancel);
    ui.navigationWeatherWarningModal?.addEventListener("click", onBackdrop);
    document.addEventListener("keydown", onKeydown);
  });
}

async function navigateToImpianto(impianto) {
  if (!selectedCommessaId || !impianto.id) return;

  const url = buildImpiantoMapsUrl(impianto);

  if (!url) {
    alert("Coordinate mancanti per questo impianto.");
    return;
  }

  const canContinueNavigation = await confirmNavigationWeatherIfNeeded(impianto);
  if (!canContinueNavigation) return;

  window.open(url, "_blank");

  const operatorName = currentUser?.displayName || currentUser?.email || "Operatore";
  const navigateAtLocal = new Date();
  const ids = getImpiantoDocIds(impianto);
  if (ids.length && !firestoreDateToMillis(impianto.navigateAt)) {
    updateImpiantoLocalState(ids, { navigateAt: navigateAtLocal, navigatedBy: operatorName });
    setImpiantoNavigated(selectedCommessaId, ids, navigateAtLocal, operatorName)
      .catch((error) => console.error("Errore salvataggio navigazione impianto:", error));
  }

  const impiantoName = impianto.denominazione || "Impianto";
  const areaLabel = [
    impianto.comune,
    impianto.competenza,
    impianto.zona,
    impianto.indirizzo
  ].find((value) => String(value || "").trim()) || "zona non specificata";
  const comuneLabel = String(impianto.comune || "").trim() || areaLabel;
  const chatText = `🧭 ${operatorName} naviga verso ${impiantoName}. La squadra è al lavoro nella zona ${comuneLabel}.`;

  try {
    await sendChatMessage({
      type: "text",
      text: chatText,
      recipientId: "",
      kind: "system",
      metadata: {
        type: "impianto_navigate",
        commessaId: selectedCommessaId,
        commessaName: selectedCommessaName || "Commessa",
        impiantoName,
        impiantoKey: buildImpiantoKey(impianto),
        comune: comuneLabel,
        area: areaLabel
      }
    });
    await publishGlobalNotificationEvent("impianto-navigate", {
      title: "Operatore in navigazione",
      body: chatText,
      commessaId: selectedCommessaId,
      commessaName: selectedCommessaName || "Commessa",
      impiantoName,
      impiantoKey: buildImpiantoKey(impianto)
    });
  } catch (error) {
    console.error("Errore invio messaggio chat navigazione impianto:", error);
  }
}

async function markImpiantoDone(impianto, options = {}) {
  const ids = getImpiantoDocIds(impianto);
  if (!selectedCommessaId || !ids.length) return false;
  const source = String(options?.source || "").trim().toLowerCase();
  if (!isNetworkOffline()) {
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

  if (isNetworkOffline()) {
    const pendingAction = upsertPendingDoneAction(impianto, ids, doneAtLocal, doneByLocal);
    expandedImpiantoKey = buildImpiantoKey(impianto);
    updateImpiantoLocalState(ids, {
      done: true,
      doneAt: doneAtLocal,
      doneBy: doneByLocal,
      pendingActionId: pendingAction.id,
      pendingActionStatus: "pending",
      pendingWhatsappStatus: "pending"
    });
    setImpiantiViewMode("done");
    alert("Sei offline: FATTO salvato localmente. WhatsApp resta in attesa e sarà disponibile quando torna internet.");
    return true;
  }

  try {
    updateImpiantoLocalState(ids, { done: true, doneAt: doneAtLocal, doneBy: doneByLocal });
    await setImpiantoDone(selectedCommessaId, ids, true, {
      doneAt: doneAtLocal,
      doneBy: doneByLocal,
      doneByUid: auth.currentUser?.uid || "",
      doneByEmail: auth.currentUser?.email || ""
    });
    expandedImpiantoKey = buildImpiantoKey(impianto);
    setImpiantiViewMode("done");
  } catch (error) {
    console.error("Aggiornamento stato FATTO non completato al primo tentativo:", error);
    if (isNetworkOffline()) {
      const pendingAction = upsertPendingDoneAction(impianto, ids, doneAtLocal, doneByLocal);
      updateImpiantoLocalState(ids, { pendingActionId: pendingAction.id, pendingActionStatus: "pending", pendingWhatsappStatus: "pending" });
      return true;
    }
    const retrySucceeded = await retrySetImpiantoDone(selectedCommessaId, ids, true);
    if (!retrySucceeded) {
      console.error("Aggiornamento stato FATTO fallito anche dopo i tentativi di retry.", { commessaId: selectedCommessaId, impiantoIds: ids });
      return false;
    }
  }

  if (!canManageData()) {
    try {
      await queueSheetExportForAdmin(exportPayload);
    } catch (error) {
      console.error("Impianto FATTO ma coda admin non salvata:", error);
    }
    await publishGlobalNotificationEvent("impianto-done", {
      title: "Impianto completato",
      body: `${doneByLocal} ha premuto ${source === "whatsapp" ? "WHAZZUP" : "FATTO"} su ${impianto.denominazione || "Impianto"} (${selectedCommessaName || "Commessa"}).`,
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
    body: `${doneByLocal} ha premuto ${source === "whatsapp" ? "WHAZZUP" : "FATTO"} su ${impianto.denominazione || "Impianto"} (${selectedCommessaName || "Commessa"}).`,
    commessaId: selectedCommessaId,
    commessaName: selectedCommessaName || "Commessa",
    impiantoName: impianto.denominazione || "Impianto",
    impiantoKey: buildImpiantoKey(impianto)
  });
  return true;
}

function markImpiantoDoneVisualFallback(impianto) {
  const ids = getImpiantoDocIds(impianto);
  if (!ids.length) return;
  const doneAtLocal = new Date();
  const doneByLocal = auth.currentUser?.displayName || auth.currentUser?.email || "Operatore";
  expandedImpiantoKey = buildImpiantoKey(impianto);
  updateImpiantoLocalState(ids, { done: true, doneAt: doneAtLocal, doneBy: doneByLocal });
  setImpiantiViewMode("done");
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
  const resetAtLocal = new Date();
  const resetByLocal = currentUser?.displayName || currentUser?.email || "Operatore";
  clearImpiantoWhazzupProcessing(impianto, selectedCommessaId);
  const safetyState = getWhazzupSafetyState(impianto);
  if (safetyState) {
    safetyState.whazzupPremuto = false;
    safetyState.needsManualMove = false;
  }
  clearWhazzupPendingDoneEntry(impianto);
  trackLocalSheetMutation(selectedCommessaId);
  updateImpiantoLocalState(ids, {
    done: false,
    doneAt: null,
    doneBy: "",
    resetAt: resetAtLocal,
    resetBy: resetByLocal,
    navigateAt: null,
    navigatedBy: "",
    pendingActionId: "",
    pendingActionStatus: "",
    pendingWhatsappStatus: ""
  });
  await setImpiantoDone(selectedCommessaId, ids, false, { resetAt: resetAtLocal, resetBy: resetByLocal });
  const impiantoKey = buildImpiantoKey(impianto);
  clearActionUsed(`${selectedCommessaId}:${impiantoKey}:navigate`);
  clearActionUsed(`${selectedCommessaId}:${impiantoKey}:done`);
  clearActionUsed(`${selectedCommessaId}:${impiantoKey}:whatsapp`);
  clearActionUsed(`${selectedCommessaId}:${impiantoKey}:reset`);
  updateConnectivityStatus();
  renderImpianti();
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
  const subcommesse = getSubcommesse(commessaId);
  if (subcommesse.length) {
    alert(`Non puoi eliminare la commessa "${nome}" perché contiene ${subcommesse.length} subcommesse. Elimina o sposta prima le subcommesse.`);
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
    stopCommessaNotesSubscription();
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
    telefono: "",
    email: "",
    mansione: "",
    note: "",
    abilitatoTutteCommesse: false,
    commesseAbilitate: [],
    corsi: normalizePersonCourses({}),
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
    ? extractPersonnelFromRawRows(await parseSimpleExcelRows(file, true))
    : await parseSimpleExcelRows(file);
  const uniqueNames = collectionName === "personale"
    ? rows
    : [...new Set(rows.filter(Boolean).map((v) => v.trim()).filter(Boolean))]
      .filter((name) => name.split(/\s+/).length >= 2);
  if (!uniqueNames.length) {
    alert("Il file Excel non contiene nomi validi.");
    return;
  }

  const batch = db.batch();
  uniqueNames.forEach((entry) => {
    const ref = db.collection(collectionName).doc();
    const normalized = String(collectionName === "personale" ? entry.fullName : entry || "").trim().replace(/\s+/g, " ");
    const [cognome, ...nomeParts] = normalized.split(" ");
    const nome = nomeParts.join(" ").trim();
    batch.set(ref, {
      nome: collectionName === "personale" ? nome : normalized,
      cognome: collectionName === "personale" ? cognome : "",
      fullName: collectionName === "personale" ? `${cognome} ${nome}`.trim() : "",
      telefono: collectionName === "personale" ? String(entry.telefono || "").trim() : "",
      email: collectionName === "personale" ? String(entry.email || "").trim() : "",
      mansione: collectionName === "personale" ? String(entry.mansione || "").trim() : "",
      note: collectionName === "personale" ? String(entry.note || "").trim() : "",
      commesseAbilitate: collectionName === "personale" ? parseMultiEntryValue(entry.commesseAbilitate || "") : [],
      corsi: collectionName === "personale" ? buildCoursesFromExcelEntry(entry) : [],
      createdAt: firebase.firestore.FieldValue.serverTimestamp()
    });
  });
  await batch.commit();
  inputEl.value = "";
  alert(`Import completato (${uniqueNames.length} elementi).`);
}

function extractPersonnelFromRawRows(rawRows) {
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

    const normalized = value ? value.replace(/\s+/g, " ") : "";
    if (!normalized) continue;
    const getByHeader = (aliases) => {
      const idx = headerKeys.findIndex((h) => aliases.includes(h));
      return idx >= 0 ? String(row[idx] || "").trim() : "";
    };
    names.push({
      fullName: normalized,
      telefono: getByHeader(["telefono", "cellulare"]),
      email: getByHeader(["email", "mail"]),
      mansione: getByHeader(["mansione", "ruolo"]),
      commesseAbilitate: getByHeader(["commesseabilitate", "commesse"]),
      corsi: getByHeader(["corsi"]),
      scadenzaCorsi: getByHeader(["scadenzacorsi", "scadenze"]),
      note: getByHeader(["note"])
    });
  }
  return names;
}

function buildCoursesFromExcelEntry(entry) {
  const names = parseMultiEntryValue(entry?.corsi || "");
  return names.map((nome) => ({ nome })).filter((c) => c.nome);
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
    renderPersonaleList(ui.personaleLista, personaleRecords, deletePersonale);
    updateSquadraHintFromSources();
    updateSuggestionLists();
    renderHoursOperatoriOptions();
  }, (error) => {
    console.error(error);
    ui.personaleLista.innerHTML = "<p class='muted'>Errore caricamento personale.</p>";
  });
}

const DEFAULT_COMMESSE_ABILITAZIONI = [
  "HERA Depurazione", "HERA Discariche", "INRETE Gas Bologna", "INRETE Gas Modena", "INRETE Gas Ferrara", "WTE", "Altro"
];
const DEFAULT_CORSI = [
  "Sicurezza generale", "Sicurezza specifica", "Primo soccorso", "Antincendio", "Preposto", "DPI 3ª categoria", "Lavori in quota", "PLE", "Carrello elevatore", "Decespugliatore / attrezzature verde", "ATEX", "Rischio biologico", "Spazi confinati"
];

const PRIMARY_CORSI = ["Primo soccorso", "Antincendio", "Preposto", "ATEX"];

function getRecentPersonaleIds() {
  try { return JSON.parse(localStorage.getItem(PERSONALE_RECENT_KEY) || "[]"); } catch { return []; }
}
function setRecentPersonaleId(id) {
  if (!id) return;
  const next = [id, ...getRecentPersonaleIds().filter((item) => item !== id)].slice(0, 3);
  localStorage.setItem(PERSONALE_RECENT_KEY, JSON.stringify(next));
}
function findPersonaleById(id) {
  return personaleRecords.find((item) => item.id === id) || null;
}
function renderPersonaleSuggestions() {
  if (!ui.personaleSearchSuggestions) return;
  const query = personaleSearchQuery.trim().toLowerCase();
  let source = [];
  if (query) {
    source = personaleRecords.filter((item) => getPersonaleDisplayName(item).toLowerCase().includes(query)).slice(0, 5);
  } else if (!personaleShowAll) {
    source = getRecentPersonaleIds().map(findPersonaleById).filter(Boolean);
  }
  ui.personaleSearchSuggestions.innerHTML = source.map((item) => {
    const commesse = Array.isArray(item.commesseAbilitate) ? item.commesseAbilitate : [];
    const isAllCommesse = Boolean(item.abilitatoTutteCommesse);
    const corsiCount = toCourseList(normalizePersonCourses(item)).length;
    return `<button type="button" class="personale-suggestion-item" data-person-id="${escapeHTML(item.id)}"><strong>${escapeHTML(getPersonaleDisplayName(item) || "Senza nome")}</strong><small>${escapeHTML(item.telefono || "-")} • ${isAllCommesse ? "✅ Tutte le commesse" : (commesse.length ? "✅ Abilitato parziale" : "❌ Non abilitato")} • Corsi ${corsiCount}/${PRIMARY_CORSI.length}</small></button>`;
  }).join("") || (query ? "<p class='muted'>Nessun risultato.</p>" : "<p class='muted'>Cerca personale per aprire una scheda.</p>");
  Array.from(ui.personaleSearchSuggestions.querySelectorAll("[data-person-id]")).forEach((btn) => {
    btn.addEventListener("click", () => {
      personaleExpandedId = btn.dataset.personId || "";
      setRecentPersonaleId(personaleExpandedId);
      renderPersonaleList(ui.personaleLista, personaleRecords, deletePersonale);
      renderPersonaleSuggestions();
    });
  });
}

function createEmptyCorsoState() {
  return { possiede: false };
}

function normalizePersonCourses(person = {}) {
  const base = Object.fromEntries(PRIMARY_CORSI.map((nome) => [nome.toLowerCase(), createEmptyCorsoState()]));
  const raw = person?.corsi;
  if (Array.isArray(raw)) {
    raw.forEach((corso) => {
      const key = String(corso?.nome || "").toLowerCase();
      if (!base[key]) return;
      base[key] = {
        possiede: true
      };
    });
    return base;
  }
  if (raw && typeof raw === "object") {
    Object.keys(base).forEach((key) => {
      const source = raw[key] || {};
      base[key] = {
        possiede: Boolean(source?.possiede)
      };
    });
  }
  return base;
}

function toCourseList(corsiObj = {}) {
  return PRIMARY_CORSI.map((nome) => {
    const key = nome.toLowerCase();
    const corso = corsiObj[key] || createEmptyCorsoState();
    if (!corso.possiede) return null;
    return { nome, possiede: true };
  }).filter(Boolean);
}

function renderPersonaleList(container, items, onDelete) {
  container.innerHTML = "";
  renderPersonaleSuggestions();
  if (!items.length || !personaleExpandedId) {
    container.innerHTML = "<p class='muted'>Apri una scheda dal campo ricerca.</p>";
    return;
  }
  items.filter((it) => it.id === personaleExpandedId || personaleShowAll).forEach((item) => {
    const card = document.createElement("div");
    card.className = "personale-card";
    const fullName = getPersonaleDisplayName(item) || "";
    const commesseAbilitate = Array.isArray(item.commesseAbilitate) ? item.commesseAbilitate : [];
    const corsi = toCourseList(normalizePersonCourses(item));
    const coursesOwned = corsi.length;
    const commesseBadge = item.abilitatoTutteCommesse
      ? "✅ Abilitato a tutte le commesse"
      : (commesseAbilitate.length ? "✅ Abilitato parziale" : "❌ Non abilitato");
    card.innerHTML = `
      <div class="personale-card-main">
        <h4>${escapeHTML(fullName || "Senza nome")}</h4>
        <p class="muted">${escapeHTML(item.telefono || "-")} • ${escapeHTML(item.email || "-")}</p>
        <div class="personale-badges">
          <span class="personale-badge ${item.abilitatoTutteCommesse || commesseAbilitate.length ? "ok" : "no"}">${commesseBadge}</span>
          <span class="personale-badge ok">Corsi: ${coursesOwned}/${PRIMARY_CORSI.length}</span>
        </div>
      </div>
      <div class="item-actions">
        <button class="btn btn-small personale-edit-btn" type="button">Modifica</button>
        <button class="btn btn-small personale-sheet-btn" type="button">Scheda</button>
        <button class="btn btn-small personale-delete-btn" type="button">Elimina</button>
      </div>
      <div class="personale-details hidden"></div>
    `;
    card.querySelector(".personale-delete-btn").addEventListener("click", () => onDelete(item.id, fullName || "elemento"));
    card.querySelector(".personale-delete-btn").disabled = !canManageData();
    card.querySelector(".personale-edit-btn").addEventListener("click", () => openPersonaleDetail(card, item, true));
    card.querySelector(".personale-sheet-btn").addEventListener("click", () => openPersonaleDetail(card, item, false));
    container.appendChild(card);
  });
}

function openPersonaleDetail(card, person, editMode = false) {
  const details = card.querySelector(".personale-details");
  const commesseOptions = [...new Set([...DEFAULT_COMMESSE_ABILITAZIONI, ...Array.from(commesseById.values()).map((c) => c.nome).filter(Boolean)])];
  const selectedCommesse = new Set(Array.isArray(person.commesseAbilitate) ? person.commesseAbilitate : []);
  const allCommesseEnabled = Boolean(person.abilitatoTutteCommesse);
  const normalizedCourses = normalizePersonCourses(person);
  const primary = PRIMARY_CORSI.map((nome) => {
    const existing = normalizedCourses[nome.toLowerCase()] || createEmptyCorsoState();
    return {
      nome,
      possiede: Boolean(existing?.possiede)
    };
  });
  details.classList.remove("hidden");
  details.innerHTML = `
    <div class="personale-fields-grid">
      <input class="personale-edit-cognome-nome" ${editMode ? "" : "disabled"} value="${escapeHTML(getPersonaleDisplayName(person) || "")}" placeholder="Cognome e nome">
      <input class="personale-edit-tel" ${editMode ? "" : "disabled"} value="${escapeHTML(person.telefono || "")}" placeholder="Telefono">
      <input class="personale-edit-email" ${editMode ? "" : "disabled"} value="${escapeHTML(person.email || "")}" placeholder="Email">
      <input class="personale-edit-ruolo" ${editMode ? "" : "disabled"} value="${escapeHTML(person.mansione || person.ruolo || "")}" placeholder="Mansione / ruolo">
      <textarea class="personale-edit-note" ${editMode ? "" : "disabled"} placeholder="Note">${escapeHTML(person.note || "")}</textarea>
    </div>
    <h5>Abilitazione commesse</h5>
    <label class="personale-commessa-check"><input class="personale-all-commesse-toggle" type="checkbox" ${editMode ? "" : "disabled"} ${allCommesseEnabled ? "checked" : ""}> Abilitato a tutte le commesse</label>
    <div class="personale-commesse-panel ${allCommesseEnabled ? "hidden" : ""}">
      <button type="button" class="btn btn-small personale-toggle-commesse-btn">Seleziona commesse</button>
      <div class="personale-badges hidden personale-commesse-list">${commesseOptions.map((nome) => `<label class="personale-commessa-check"><input type="checkbox" ${editMode ? "" : "disabled"} value="${escapeHTML(nome)}" ${selectedCommesse.has(nome) ? "checked" : ""}> ${escapeHTML(nome)}</label>`).join("")}</div>
    </div>
    <h5>Corsi principali</h5>
    <div class="personale-corsi-list">${primary.map((corso, idx) => {
      return `<div class="personale-corso-row personale-corso-main" data-course-key="${idx}">
        <label class="personale-course-toggle-label"><input class="personale-course-has" type="checkbox" ${editMode ? "" : "disabled"} ${corso.possiede ? "checked" : ""}> ${escapeHTML(corso.nome)}</label>
      </div>`;
    }).join("")}</div>
    ${editMode ? '<button type="button" class="btn btn-primary personale-save-btn">Salva scheda</button>' : ""}
  `;
  details.querySelector(".personale-toggle-commesse-btn")?.addEventListener("click", () => details.querySelector(".personale-commesse-list")?.classList.toggle("hidden"));
  const autosave = () => (editMode ? savePersonaleDetail(person.id, details) : Promise.resolve());
  details.querySelector(".personale-all-commesse-toggle")?.addEventListener("change", (event) => {
    details.querySelector(".personale-commesse-panel")?.classList.toggle("hidden", event.target.checked);
    autosave();
  });
  details.querySelectorAll(".personale-course-has").forEach((cb) => cb.addEventListener("change", autosave));
  details.querySelectorAll(".personale-edit-cognome-nome, .personale-edit-tel, .personale-edit-email, .personale-edit-ruolo, .personale-edit-note, .personale-commesse-list input[type='checkbox']").forEach((el) => {
    el.addEventListener("change", autosave);
    el.addEventListener("input", autosave);
  });
  if (editMode) details.querySelector(".personale-save-btn").addEventListener("click", async () => savePersonaleDetail(person.id, details));
}

async function savePersonaleDetail(personId, root) {
  if (!canManageData()) return;
  const fullName = String(root.querySelector(".personale-edit-cognome-nome")?.value || "").trim().replace(/\s+/g, " ");
  const [cognome, ...nomeParts] = fullName.split(" ");
  const nome = nomeParts.join(" ").trim();
  const allCommesseEnabled = Boolean(root.querySelector(".personale-all-commesse-toggle")?.checked);
  const commesseAbilitate = Array.from(root.querySelectorAll(".personale-commesse-list input[type='checkbox']:checked")).map((el) => el.value);
  const corsi = PRIMARY_CORSI.reduce((acc, nomeCorso, idx) => {
    const hasCourse = Boolean(root.querySelector(`.personale-corso-main[data-course-key="${idx}"] .personale-course-has`)?.checked);
    acc[nomeCorso.toLowerCase()] = {
      possiede: hasCourse
    };
    return acc;
  }, {});
  await db.collection("personale").doc(personId).set({
    nome, cognome, fullName: `${cognome} ${nome}`.trim(),
    telefono: String(root.querySelector(".personale-edit-tel")?.value || "").trim(),
    email: String(root.querySelector(".personale-edit-email")?.value || "").trim(),
    mansione: String(root.querySelector(".personale-edit-ruolo")?.value || "").trim(),
    note: String(root.querySelector(".personale-edit-note")?.value || "").trim(),
    abilitatoTutteCommesse: allCommesseEnabled,
    commesseAbilitate,
    corsi
  }, { merge: true });
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
  squadreLoadState = { status: "loading", message: "Caricamento squadre..." };
  renderSquadre();

  unsubscribeSquadre = db.collection("squadreCommesse").onSnapshot((snapshot) => {
    squadreByCommessa = new Map();
    snapshot.forEach((doc) => {
      squadreByCommessa.set(doc.id, { id: doc.id, ...doc.data() });
    });
    renderSquadre();
    updateCommessaDashboard();
    renderCommesseHomeList();
    autofillSquadraForm();
    Array.from(ui.hoursCommesseList?.querySelectorAll(".hours-commessa-card") || []).forEach((card) => {
      applyHoursSuggestedOperators(card, { force: true });
    });
    checkAndSendHoursDeadlineAlerts();
  }, (error) => {
    console.error("Errore caricamento squadre:", error);
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
    squadreLoadState = { status: "loaded", message: "" };
    renderSquadre();
    updateCommessaDashboard();
    renderCommesseHomeList();
    tryAutoOpenAssignedCommessaAtStartup();
    Array.from(ui.hoursCommesseList?.querySelectorAll(".hours-commessa-card") || []).forEach((card) => {
      applyHoursSuggestedOperators(card, { force: true });
    });
  }, (error) => {
    console.error("Errore caricamento squadre:", error);
    squadreLoadState = { status: "error", message: "Errore caricamento squadre da Firebase." };
    renderSquadre();
  });

  unsubscribeSquadreViewConfig = db.collection("appConfig").doc("squadreView").onSnapshot((doc) => {
    const data = doc.exists ? doc.data() || {} : {};
    sharedSquadreDateKey = String(data.selectedDateKey || "").trim();
    manualSquadreFilterDateKey = sharedSquadreDateKey;
    sharedSquadreViewConfigLoaded = true;
    syncSquadreDateInputs();
    renderSquadre();
    updateCommessaDashboard();
    tryAutoOpenAssignedCommessaAtStartup();
  }, (error) => {
    console.error("Errore caricamento giorno squadre condiviso:", error);
    sharedSquadreViewConfigLoaded = true;
    tryAutoOpenAssignedCommessaAtStartup();
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
  if (unsubscribeSquadreViewConfig) {
    unsubscribeSquadreViewConfig();
    unsubscribeSquadreViewConfig = null;
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
    setDefaultSquadraCompositionDate();
    addSquadraRow();
    return;
  }

  const data = squadreByCommessa.get(commessaId) || {};
  setDefaultSquadraCompositionDate();
  setSquadraRowsFromData(data);
}

function addSquadraRow(rowData = { caposquadra: "", personale: "", mezzi: "", note: "", orario: "", impianti: "" }) {
  const index = ui.squadraRows.children.length + 1;
  const personaleValues = parseMultiEntryValue(rowData.personale);
  const mezziValues = parseMultiEntryValue(rowData.mezzi);
  const impiantiValues = parseMultiEntryValue(rowData.impianti || rowData.impiantiAssegnati || "");
  const row = document.createElement("div");
  row.className = "squadra-row";
  row.innerHTML = `
    <div class="squadra-row-head">
      <strong>Squadra ${index}</strong>
    </div>
    <label class="squadra-simple-field">Caposquadra
      <input type="text" class="squadra-caposquadra-input" list="personale-options" placeholder="Caposquadra" value="${escapeHTML(rowData.caposquadra || "")}">
    </label>
    <label class="squadra-simple-field">Orario
      <input type="time" class="squadra-orario-input" value="${escapeHTML(rowData.orario || "")}">
    </label>
    <div class="squadra-multi-field">
      <div class="squadra-multi-field-head"><strong>👥 Operatori</strong></div>
      <div class="squadra-personale-list"></div>
      <button type="button" class="btn btn-small add-personale-input-btn">+ Operatore</button>
    </div>
    <div class="squadra-multi-field">
      <div class="squadra-multi-field-head"><strong>🚚 Mezzi</strong></div>
      <div class="squadra-mezzi-list"></div>
      <button type="button" class="btn btn-small add-mezzo-input-btn">+ Mezzo</button>
    </div>
    <div class="squadra-multi-field">
      <div class="squadra-multi-field-head"><strong>📍 Impianti</strong></div>
      <div class="squadra-impianti-list"></div>
      <button type="button" class="btn btn-small add-impianto-input-btn">+ Impianto</button>
    </div>
    <label class="squadra-note-field">Note
      <textarea class="squadra-note-input" rows="2" placeholder="Note squadra">${escapeHTML(rowData.note || "")}</textarea>
    </label>
  `;
  const personaleList = row.querySelector(".squadra-personale-list");
  const mezziList = row.querySelector(".squadra-mezzi-list");
  const impiantiList = row.querySelector(".squadra-impianti-list");
  const addPersonaleBtn = row.querySelector(".add-personale-input-btn");
  const addMezzoBtn = row.querySelector(".add-mezzo-input-btn");
  const addImpiantoBtn = row.querySelector(".add-impianto-input-btn");

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
  const addImpiantoInput = (value = "") => addMultiEntryInput({
    container: impiantiList,
    listId: "",
    placeholder: "Impianto assegnato",
    value,
    sourceValues: []
  });

  (personaleValues.length ? personaleValues : [""]).forEach((value) => addPersonaleInput(value));
  (mezziValues.length ? mezziValues : [""]).forEach((value) => addMezzoInput(value));
  (impiantiValues.length ? impiantiValues : [""]).forEach((value) => addImpiantoInput(value));

  addPersonaleBtn.addEventListener("click", () => addPersonaleInput(""));
  addMezzoBtn.addEventListener("click", () => addMezzoInput(""));
  addImpiantoBtn.addEventListener("click", () => addImpiantoInput(""));
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

function updateEmptySquadraRowsHint() {
  if (!ui.squadraHint || !canManageData()) return;
  if (!ui.squadraRows.children.length) {
    ui.squadraHint.textContent = "Nessuna squadra nel modulo: salva per eliminare la composizione della commessa in questa data, oppure premi “Aggiungi squadra”.";
  }
}

function isSquadraRowFilled(row) {
  return Boolean(row?.caposquadra || row?.personale || row?.mezzi || row?.impianti || row?.note || row?.orario);
}

function readSquadraRows() {
  return Array.from(ui.squadraRows.querySelectorAll(".squadra-row")).map((row) => ({
    caposquadra: String(row.querySelector(".squadra-caposquadra-input")?.value || "").trim(),
    personale: Array.from(row.querySelectorAll(".squadra-personale-list .squadra-multi-entry-input"))
      .map((input) => String(input.value || "").trim())
      .filter(Boolean)
      .join(", "),
    mezzi: Array.from(row.querySelectorAll(".squadra-mezzi-list .squadra-multi-entry-input"))
      .map((input) => String(input.value || "").trim())
      .filter(Boolean)
      .join(", "),
    impianti: Array.from(row.querySelectorAll(".squadra-impianti-list .squadra-multi-entry-input"))
      .map((input) => String(input.value || "").trim())
      .filter(Boolean)
      .join(", "),
    note: String(row.querySelector(".squadra-note-input")?.value || "").trim(),
    orario: String(row.querySelector(".squadra-orario-input")?.value || "").trim()
  })).filter(isSquadraRowFilled);
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


function normalizeSquadraDuplicatePart(value) {
  return String(value || "")
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .toLocaleLowerCase("it-IT")
    .replace(/[^a-z0-9]+/g, " ")
    .replace(/\s+/g, " ")
    .trim();
}

function buildSquadraDuplicateSignature(row) {
  const caposquadra = normalizeSquadraDuplicatePart(row?.caposquadra || "");
  const operatori = parseMultiEntryValue(row?.personale || "")
    .map(normalizeSquadraDuplicatePart)
    .filter(Boolean)
    .sort()
    .join("|");
  return `${caposquadra}__${operatori}`;
}

function findDuplicateSquadraRows(rows) {
  const seen = new Map();
  const duplicates = [];
  rows.forEach((row, index) => {
    const signature = buildSquadraDuplicateSignature(row);
    if (!signature.replace(/[_|]/g, "")) return;
    if (seen.has(signature)) {
      duplicates.push({ firstIndex: seen.get(signature), duplicateIndex: index });
      return;
    }
    seen.set(signature, index);
  });
  return duplicates;
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
  if (!dateKey) {
    alert("Seleziona una data per la composizione squadre.");
    return;
  }
  const squadreRows = readSquadraRows();
  if (!squadreRows.length) {
    await deleteSquadraCompositionForDate(commessaId, dateKey);
    return;
  }
  const duplicateRows = findDuplicateSquadraRows(squadreRows);
  if (duplicateRows.length) {
    const dup = duplicateRows[0];
    alert(`Duplicato bloccato: Squadra ${dup.duplicateIndex + 1} è identica alla Squadra ${dup.firstIndex + 1} per caposquadra e operatori.`);
    return;
  }
  const payload = {
    commessaId,
    commessaNome: (commesseById.get(commessaId) || {}).nome || "Commessa",
    riferimentoData: dateKey,
    squadre: squadreRows,
    updatedAt: firebase.firestore.FieldValue.serverTimestamp(),
    updatedBy: (currentUser && currentUser.email) ? currentUser.email : ""
  };
  const currentRef = db.collection("squadreCommesse").doc(commessaId);
  const historyRef = db.collection("squadreStorico").doc(`${dateKey}__${commessaId}`);
  try {
    await db.runTransaction(async (transaction) => {
      const historySnap = await transaction.get(historyRef);
      const currentRows = historySnap.exists
        ? (Array.isArray(historySnap.data()?.squadre) ? historySnap.data().squadre : getLegacySquadreRows(historySnap.data() || {}))
        : [];
      const mergedRows = squadreRows;
      const transactionDuplicates = findDuplicateSquadraRows(mergedRows);
      if (transactionDuplicates.length) {
        throw new Error("DUPLICATE_SQUADRA");
      }
      const nextPayload = {
        ...payload,
        squadre: mergedRows,
        existingSquadreCountBeforeSave: currentRows.length
      };
      transaction.set(currentRef, nextPayload, { merge: true });
      transaction.set(historyRef, {
        ...nextPayload,
        dateKey
      }, { merge: true });
    });
  } catch (error) {
    if (error?.message === "DUPLICATE_SQUADRA") {
      alert("Duplicato bloccato: esiste già una squadra con stessa commessa, stessa data, stesso caposquadra e stessi operatori.");
      return;
    }
    throw error;
  }
  renderCommesseHomeList();
  renderSquadre();
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

async function deleteSquadraCompositionForDate(commessaId, dateKey) {
  if (!canManageData()) return;
  const commessaNome = (commesseById.get(commessaId) || {}).nome || "Commessa";
  const dateLabel = new Date(`${dateKey}T00:00:00`).toLocaleDateString("it-IT");
  if (!window.confirm(`Eliminare tutte le squadre di ${commessaNome} per il ${dateLabel}?`)) return;

  const currentRef = db.collection("squadreCommesse").doc(commessaId);
  const historyRef = db.collection("squadreStorico").doc(`${dateKey}__${commessaId}`);
  await db.runTransaction(async (transaction) => {
    const currentSnap = await transaction.get(currentRef);
    transaction.delete(historyRef);
    if (!currentSnap.exists || currentSnap.data()?.riferimentoData === dateKey) {
      transaction.delete(currentRef);
    }
  });
  ui.squadraRows.innerHTML = "";
  addSquadraRow();
  ui.squadraCalendarDate.value = dateKey;
  renderCommesseHomeList();
  renderSquadre();
}

function getDateKeyFromLocalDate(date) {
  const year = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, "0");
  const day = String(date.getDate()).padStart(2, "0");
  return `${year}-${month}-${day}`;
}

function getNextDayDateKey(now = new Date()) {
  const nextDay = new Date(now);
  nextDay.setDate(nextDay.getDate() + 1);
  return getDateKeyFromLocalDate(nextDay);
}

function setDefaultSquadraCompositionDate({ force = false } = {}) {
  if (!ui.squadraRiferimento) return;
  if (force || !ui.squadraRiferimento.value) {
    ui.squadraRiferimento.value = getNextDayDateKey();
  }
}

function getAutomaticSquadreDateKey(now = new Date()) {
  const base = new Date(now);
  if (base.getHours() > 17 || (base.getHours() === 17 && base.getMinutes() >= 30)) {
    base.setDate(base.getDate() + 1);
  }
  return getDateKeyFromLocalDate(base);
}

function initializeAutomaticSquadreDate() {
  automaticSquadreDateKey = getAutomaticSquadreDateKey();
  syncSquadreDateInputs();
  renderSquadre();
}

function getActiveSquadreDateKey() {
  if (manualSquadreFilterDateKey) return manualSquadreFilterDateKey;
  if (sharedSquadreDateKey) return sharedSquadreDateKey;
  if (!automaticSquadreDateKey) automaticSquadreDateKey = getAutomaticSquadreDateKey();
  return automaticSquadreDateKey;
}

function syncSquadreDateInputs() {
  const activeDateKey = getActiveSquadreDateKey();
  if (ui.squadreFilterDate) ui.squadreFilterDate.value = activeDateKey;
  if (ui.squadraCalendarDate) ui.squadraCalendarDate.value = activeDateKey;
}

async function persistSharedSquadreDate(dateKey) {
  if (!canManageData()) return;
  await db.collection("appConfig").doc("squadreView").set({
    selectedDateKey: dateKey || "",
    updatedAt: firebase.firestore.FieldValue.serverTimestamp(),
    updatedBy: (currentUser && currentUser.email) ? currentUser.email : ""
  }, { merge: true });
}

function setSquadreDateOverride(dateKey) {
  const selectedDateKey = String(dateKey || "").trim();
  manualSquadreFilterDateKey = selectedDateKey;
  sharedSquadreDateKey = selectedDateKey;
  syncSquadreDateInputs();
  renderSquadre();
  persistSharedSquadreDate(manualSquadreFilterDateKey).catch((error) => {
    console.error("Errore salvataggio giorno squadre condiviso:", error);
  });
}

function onSquadreFilterDateChange() {
  setSquadreDateOverride(ui.squadreFilterDate?.value || "");
}

function clearManualSquadreFilterDate() {
  manualSquadreFilterDateKey = "";
  sharedSquadreDateKey = "";
  persistSharedSquadreDate("").catch((error) => {
    console.error("Errore reset giorno squadre condiviso:", error);
  });
  initializeAutomaticSquadreDate();
}


const INRETE_COMMESSE_REQUIRED = new Set([
  "inrete gas bologna",
  "inrete gas modena",
  "inrete gas ferrara",
  "inrete ferrara",
  "inrete bologna",
  "inrete modena"
]);

function normalizeSafetyKey(value) {
  return String(value || "")
    .normalize("NFD")
    .replace(/[̀-ͯ]/g, "")
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, " ")
    .trim();
}

function getPersonByDisplayName(name) {
  const target = normalizeSafetyKey(name);
  if (!target) return null;
  return personaleRecords.find((person) => normalizeSafetyKey(getPersonaleDisplayName(person)) === target) || null;
}

function isPersonAbilitataForCommessa(person, commessaName) {
  if (!person) return false;
  if (Boolean(person.abilitatoTutteCommesse)) return true;
  const target = normalizeSafetyKey(commessaName);
  const enabled = Array.isArray(person.commesseAbilitate) ? person.commesseAbilitate : [];
  return enabled.some((item) => normalizeSafetyKey(item) === target);
}

function hasRequiredCourse(person, courseKey) {
  const corsi = normalizePersonCourses(person);
  return Boolean(corsi?.[courseKey]?.possiede);
}

function buildSquadraWarningDetails(commessa, squadRows) {
  const commessaName = String(commessa?.nome || "").trim();
  const isInrete = INRETE_COMMESSE_REQUIRED.has(normalizeSafetyKey(commessaName));
  const requiredCourses = isInrete
    ? ["primo soccorso", "antincendio", "preposto", "atex"]
    : ["primo soccorso", "antincendio", "preposto"];
  const missingLabelPrefix = isInrete ? "Requisiti INRETE mancanti" : "Requisiti sicurezza mancanti";

  const issues = [];
  const abilitati = [];
  squadRows.forEach((row) => {
    parseMultiEntryValue(row?.personale || "").forEach((name) => {
      const person = getPersonByDisplayName(name);
      if (!person) {
        issues.push(`⚠️ ${name} non è presente in anagrafica personale`);
        return;
      }
      const isAbilitata = isPersonAbilitataForCommessa(person, commessaName);
      if (!isAbilitata) {
        issues.push(`⚠️ ${getPersonaleDisplayName(person) || name} non è abilitato per questa commessa`);
        return;
      }
      abilitati.push(person);
    });
  });

  const missing = requiredCourses.filter((course) => !abilitati.some((person) => hasRequiredCourse(person, course)));
  if (missing.length) {
    issues.push(`⚠️ ${missingLabelPrefix}:`);
    missing.forEach((course) => issues.push(`- manca ${course.toUpperCase() === "ATEX" ? "ATEX" : course.replace(/\b\w/g, (ch) => ch.toUpperCase())}`));
  }

  return issues;
}

function renderSquadre() {
  if (!ui.squadreLista) return;
  ui.squadreLista.innerHTML = "";
  const selectedDateKey = getActiveSquadreDateKey();
  if (!selectedDateKey) return;
  const storicoDelGiorno = squadreHistoryByDate.get(selectedDateKey) || new Map();
  if (squadreLoadState.status === "loading") {
    ui.squadreLista.innerHTML = "<p class='muted'>Caricamento squadre...</p>";
    return;
  }
  if (squadreLoadState.status === "error") {
    ui.squadreLista.innerHTML = "<p class='muted'>Errore caricamento squadre da Firebase.</p>";
    return;
  }

  const commesse = Array.from(commesseById.values());
  const commesseConSquadre = commesse.filter((commessa) => {
    const squad = storicoDelGiorno.get(commessa.id) || {};
    const rows = Array.isArray(squad.squadre) ? squad.squadre : getLegacySquadreRows(squad);
    return rows.some(isSquadraRowFilled);
  });
  if (!commesseConSquadre.length) {
    ui.squadreLista.innerHTML = "<p class='muted'>Nessuna squadra inserita per questo giorno.</p>";
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
    const rowsHtml = squadRows.map((row, idx) => {
      const details = [
        row.caposquadra ? `<br><b>🧑‍✈️ Caposquadra:</b> ${escapeHTML(row.caposquadra)}` : "",
        row.orario ? `<br><b>🕒 Orario:</b> ${escapeHTML(row.orario)}` : "",
        row.impianti ? `<br><b>📍 Impianti:</b> ${escapeHTML(row.impianti)}` : "",
        row.note ? `<br><b>📝 Note:</b> ${escapeHTML(row.note)}` : ""
      ].join("");
      return `<div class="squadra-saved-row"><p><b>👥 Squadra ${idx + 1}:</b> ${escapeHTML(row.personale || "-")}${details}<br><b>🚚 Mezzi ${idx + 1}:</b> ${renderMezziButtonsMarkup(row.mezzi)}</p></div>`;
    }).join("");
    const warningIssues = buildSquadraWarningDetails(commessa, squadRows);
    const warningMarkup = warningIssues.length
      ? `<div class="squadra-warning-wrap"><button type="button" class="squadra-warning-toggle" aria-expanded="false" aria-label="Mostra controllo squadra">⚠️</button><div class="squadra-warning-details hidden"><p><b>⚠️ Controllo squadra</b></p><ul>${warningIssues.map((issue) => `<li>${escapeHTML(issue.replace(/^⚠️\s*/, ""))}</li>`).join("")}</ul></div></div>`
      : "";
    item.innerHTML = `
      <div class="squadra-item-head squadra-commessa-link" role="button" tabindex="0" aria-label="Apri dettaglio commessa ${escapeHTML(commessa.nome || "Commessa senza nome")}">
        <strong>📁 ${escapeHTML(commessa.nome || "Commessa senza nome")}</strong>
        ${warningMarkup}
      </div>
      <p><b>📅 Giorno:</b> ${escapeHTML(riferimento)}</p>
      ${rowsHtml}
    `;
    const head = item.querySelector(".squadra-item-head");
    appendAddHoursButtonIfAllowed(head, commessa, selectedDateKey);
    head?.addEventListener("click", (event) => {
      if (event.target.closest("button, a, input, select, textarea")) return;
      openCommessaFromSquadre(commessa);
    });
    head?.addEventListener("keydown", (event) => {
      if (event.key !== "Enter" && event.key !== " ") return;
      if (event.target.closest("button, a, input, select, textarea")) return;
      event.preventDefault();
      openCommessaFromSquadre(commessa);
    });
    item.querySelectorAll(".mezzo-chip-btn").forEach((btn) => {
      btn.addEventListener("click", () => openFuelPage(btn.dataset.mezzo || ""));
    });
    const warningToggle = item.querySelector(".squadra-warning-toggle");
    warningToggle?.addEventListener("click", (event) => {
      event.stopPropagation();
      const details = item.querySelector(".squadra-warning-details");
      const isHidden = details?.classList.contains("hidden");
      details?.classList.toggle("hidden", !isHidden);
      warningToggle.setAttribute("aria-expanded", isHidden ? "true" : "false");
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

async function setImpiantoDone(commessaId, impiantoIds, done, options = {}) {
  const user = auth.currentUser;
  if (!user) return;
  const doneAtDate = options.doneAt instanceof Date ? options.doneAt : new Date();
  const doneAt = done ? firebase.firestore.Timestamp.fromDate(doneAtDate) : null;

  if (!commessaId) throw new Error("Commessa non selezionata per aggiornamento stato impianto.");
  const ref = db.collection("commesse").doc(commessaId).collection("impianti");
  await Promise.all(impiantoIds.map((impiantoId) => {
    const payload = {
      done,
      doneAt,
      doneBy: done ? (options.doneBy || user.displayName || user.email || "Operatore") : "",
      doneByUid: done ? String(options.doneByUid || user.uid || "") : "",
      doneByEmail: done ? String(options.doneByEmail || user.email || "") : ""
    };
    if (done) {
      payload.resetAt = null;
      payload.resetBy = "";
    }
    if (!done) {
      const resetAtDate = options.resetAt instanceof Date ? options.resetAt : new Date();
      payload.resetAt = firebase.firestore.Timestamp.fromDate(resetAtDate);
      payload.resetBy = options.resetBy || user.displayName || user.email || "Operatore";
      payload.navigateAt = null;
      payload.navigatedBy = "";
    }
    return ref.doc(impiantoId).set(payload, { merge: true });
  }));
}

function canTriggerImpiantoWhatsApp(impianto, notify = true) {
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

function triggerImpiantoWhatsAppAction(impianto, options = {}) {
  if (!options.force && !canTriggerImpiantoWhatsApp(impianto, true)) return false;
  return openWhatsApp(impianto);
}

async function handleImpiantoWhatsAppClick(impianto) {
  if (!impianto) return;

  const doneAt = new Date();
  const doneBy = auth.currentUser?.displayName || auth.currentUser?.email || "Operatore";
  const doneByEmail = auth.currentUser?.email || "";
  const doneByUid = auth.currentUser?.uid || "";
  const doneIds = getImpiantoDocIds(impianto);

  markWhazzupSafetyPressed(impianto, doneAt);
  upsertWhazzupPendingDoneEntry(impianto, doneAt);
  updateImpiantoLocalState(doneIds, { done: true, doneAt, doneBy, doneByEmail, doneByUid });
  setImpiantiViewMode("done");
  updateConnectivityStatus();
  renderImpianti();
  openWhatsApp(impianto, { doneAt, operatorName: doneBy });

  const auditLogId = await auditLogWhazzupClick(impianto, { clickedAt: doneAt, fattoEsito: "pending", fattoConfermato: false })
    .catch((error) => {
      console.error("Errore avvio audit log Whazzup:", error);
      return null;
    });

  try {
    let doneMarked = false;
    console.debug("[WHAZZUP->FATTO] Avvio salvataggio", { commessaId: selectedCommessaId, impiantoKey: buildImpiantoKey(impianto) });
    try {
      doneMarked = await markImpiantoDone(impianto, { source: "whatsapp" });
    } catch (error) {
      console.error("Errore salvataggio impianto FATTO da WhatsApp:", error);
      await updateAuditLogWhazzupClick(auditLogId, {
        fattoEsito: "save_exception",
        fattoConfermato: false,
        errorMessage: String(error?.message || error || "Errore sconosciuto")
      });
    }
    if (!doneMarked) {
      await updateAuditLogWhazzupClick(auditLogId, { fattoEsito: "save_failed", fattoConfermato: false });
      alert("Errore salvataggio. Riprova.");
      return;
    }
    const persisted = await verifyImpiantoDoneBackground(impianto);
    await updateAuditLogWhazzupClick(auditLogId, {
      fattoEsito: persisted ? "persisted" : "verify_failed",
      fattoConfermato: Boolean(persisted)
    });
    if (!persisted) {
      alert("Errore salvataggio. Riprova.");
      return;
    }
    updateConnectivityStatus();
    renderImpianti();
  } catch (error) {
    console.error("Errore processo FATTO:", error);
    alert("Errore salvataggio. Riprova.");
  }
}


async function auditLogWhazzupClick(impianto, options = {}) {
  if (!db) return null;
  const now = options.clickedAt instanceof Date ? options.clickedAt : new Date();
  const commessaId = String(options.commessaId || selectedCommessaId || "").trim();
  const commessaName = String(options.commessaName || selectedCommessaName || "").trim() || "Commessa";
  const user = auth.currentUser || currentUser || null;
  const payload = {
    eventType: "whazzup_click",
    createdAt: firebase.firestore.Timestamp.fromDate(now),
    clickedAtIso: now.toISOString(),
    commessaId,
    commessaName,
    impiantoKey: buildImpiantoKey(impianto),
    impiantoIdSap: String(impianto?.idSap || "").trim(),
    impiantoNome: String(impianto?.denominazione || "").trim() || "Impianto",
    impiantoComune: String(impianto?.comune || "").trim(),
    impiantoDocIds: getImpiantoDocIds(impianto).filter(Boolean),
    userUid: String(user?.uid || "").trim(),
    userEmail: String(user?.email || "").trim(),
    userName: String(user?.displayName || "").trim() || String(user?.email || "").trim() || "Operatore",
    fattoEsito: String(options.fattoEsito || "pending"),
    fattoConfermato: Boolean(options.fattoConfermato)
  };
  if (options.errorMessage) payload.errorMessage = String(options.errorMessage).slice(0, 500);
  try {
    const ref = await db.collection("auditLogsWhazzup").add(payload);
    return ref.id;
  } catch (error) {
    console.error("Errore salvataggio audit log Whazzup:", error);
    return null;
  }
}

async function updateAuditLogWhazzupClick(logId, patch = {}) {
  if (!logId || !db) return;
  try {
    await db.collection("auditLogsWhazzup").doc(logId).set({
      ...patch,
      updatedAt: firebase.firestore.FieldValue.serverTimestamp()
    }, { merge: true });
  } catch (error) {
    console.error("Errore aggiornamento audit log Whazzup:", error);
  }
}

function notifyImpiantoBackgroundSyncPending() {
  if (!ui.gpsStatus) return;
  ui.gpsStatus.textContent = "Attendere, salvataggio in corso…";
}

function getWhazzupProcessingKey(impianto, commessaId = selectedCommessaId) {
  const commessaKey = String(commessaId || "").trim();
  const impiantoKey = buildImpiantoKey(impianto);
  if (!commessaKey || !impiantoKey) return "";
  return `${commessaKey}:${impiantoKey}`;
}

function isImpiantoWhazzupProcessing(impianto, commessaId = selectedCommessaId) {
  const key = getWhazzupProcessingKey(impianto, commessaId);
  return key ? whazzupProcessingByImpianto.has(key) : false;
}

function clearImpiantoWhazzupProcessing(impianto, commessaId = selectedCommessaId) {
  const key = getWhazzupProcessingKey(impianto, commessaId);
  if (!key) return;
  whazzupProcessingByImpianto.delete(key);
}

async function verifyImpiantoDoneBackground(impianto) {
  await new Promise((resolve) => setTimeout(resolve, 3200));
  const persisted = await isImpiantoPersistedAsDone(impianto);
  console.debug("[WHAZZUP->FATTO] Esito verifica persistenza", { commessaId: selectedCommessaId, impiantoKey: buildImpiantoKey(impianto), persisted });
  updateWhazzupSafetyAfterBackgroundCheck(impianto, persisted);
  if (persisted) return true;
  await handleImpiantoDoneSaveFailure(impianto, "Verifica post-salvataggio negativa: impianto non presente nei FATTI.");
  return false;
}

function getWhazzupSafetyState(impianto) {
  const commessaId = String(selectedCommessaId || "").trim();
  const impiantoKey = buildImpiantoKey(impianto);
  if (!commessaId || !impiantoKey) return null;
  return whazzupSafetyByImpianto.get(`${commessaId}:${impiantoKey}`) || null;
}

function markWhazzupSafetyPressed(impianto, pressedAt = new Date()) {
  const commessaId = String(selectedCommessaId || "").trim();
  const impiantoKey = buildImpiantoKey(impianto);
  if (!commessaId || !impiantoKey) return;
  whazzupSafetyByImpianto.set(`${commessaId}:${impiantoKey}`, {
    whazzupPremuto: true,
    whazzupPremutoAlle: pressedAt.toISOString(),
    needsManualMove: false
  });
}

function updateWhazzupSafetyAfterBackgroundCheck(impianto, isDonePersisted) {
  const state = getWhazzupSafetyState(impianto);
  if (!state) return;
  state.needsManualMove = !isDonePersisted;
  if (isDonePersisted) {
    state.whazzupPremuto = false;
    clearWhazzupPendingDoneEntry(impianto);
  }
  renderImpianti();
}

async function forceMoveImpiantoToFatti(impianto) {
  const moved = await markImpiantoDone(impianto, { source: "whatsapp" });
  if (!moved) return;
  const state = getWhazzupSafetyState(impianto);
  if (state) {
    state.needsManualMove = false;
    state.whazzupPremuto = false;
  }
  clearWhazzupPendingDoneEntry(impianto);
  setImpiantiViewMode("done");
  renderImpianti();
}

async function runWhazzupPendingDoneSafetyCheck() {
  const commessaId = String(selectedCommessaId || "").trim();
  const uid = currentUser?.uid || "";
  if (!commessaId || !uid) return;
  const allEntries = loadWhazzupPendingDoneEntries();
  const pendingEntries = allEntries.filter((entry) => entry.commessaId === commessaId && (!entry.userId || entry.userId === uid));
  if (!pendingEntries.length) return;
  const remaining = [];
  for (const entry of pendingEntries) {
    const impianto = currentImpianti.find((item) => buildImpiantoKey(item) === entry.impiantoKey)
      || { denominazione: entry.impiantoName, sourceIds: entry.impiantoIds || [], id: entry.impiantoIds?.[0] || "" };
    const persisted = await isImpiantoPersistedAsDone(impianto);
    if (persisted) {
      const state = getWhazzupSafetyState(impianto);
      if (state) {
        state.whazzupPremuto = false;
        state.needsManualMove = false;
      }
      continue;
    }
    markWhazzupSafetyPressed(impianto, entry.pendingAt ? new Date(entry.pendingAt) : new Date());
    const state = getWhazzupSafetyState(impianto);
    if (state) state.needsManualMove = true;
    remaining.push(entry);
  }
  const untouched = allEntries.filter((entry) => !(entry.commessaId === commessaId && (!entry.userId || entry.userId === uid)));
  saveWhazzupPendingDoneEntries([...untouched, ...remaining]);
  renderImpianti();
}

async function isImpiantoPersistedAsDone(impianto) {
  const commessaId = String(selectedCommessaId || "").trim();
  const impiantoIds = getImpiantoDocIds(impianto).filter(Boolean);
  if (!commessaId || !impiantoIds.length) return false;
  try {
    const ref = db.collection("commesse").doc(commessaId).collection("impianti");
    const snapshots = await Promise.all(impiantoIds.map((impiantoId) => ref.doc(impiantoId).get()));
    return snapshots.some((snap) => snap.exists && isImpiantoDoneState(snap.data() || {}));
  } catch (error) {
    console.error("Errore verifica persistenza FATTO:", error);
    return false;
  }
}

async function handleImpiantoDoneSaveFailure(impianto, reason = "") {
  alert("Errore salvataggio. Riprova.");
  try {
    await notifyAdminsForImpiantoDoneSaveError(impianto, reason);
  } catch (error) {
    console.error("Errore notifica admin salvataggio FATTO:", error);
  }
}

async function notifyAdminsForImpiantoDoneSaveError(impianto, reason = "") {
  const adminUsers = platformUsers.filter((user) => adminEmails.has(normalizeEmail(user.email)));
  if (!adminUsers.length) return;
  const operatorName = currentUser?.displayName || currentUser?.email || "Operatore";
  const now = new Date();
  const text = [
    "⚠️ ERRORE SALVATAGGIO FATTO",
    "L’impianto è stato inviato su Whazzup, ma potrebbe non essere passato nella lista FATTI.",
    "Verificare manualmente.",
    "",
    `Impianto: ${impianto?.denominazione || "Impianto"}`,
    `ID SAP: ${impianto?.idSap || "-"}`,
    `Comune: ${impianto?.comune || "-"}`,
    `Operatore: ${operatorName}`,
    `Data e ora: ${now.toLocaleString("it-IT")}`,
    reason ? `Errore rilevato: ${reason}` : "Errore rilevato: verifica FATTI non confermata."
  ].join("\n");

  await Promise.all(adminUsers.map((adminUser) => db.collection("chatMessages").add({
    type: "text",
    text,
    recipientId: adminUser.id,
    senderId: currentUser?.uid || "",
    senderName: operatorName,
    senderEmail: currentUser?.email || "",
    kind: "system",
    metadata: {
      type: "impianto_done_save_error",
      commessaId: selectedCommessaId || "",
      commessaName: selectedCommessaName || "Commessa",
      impiantoName: impianto?.denominazione || "Impianto",
      impiantoKey: buildImpiantoKey(impianto),
      impiantoIdSap: impianto?.idSap || "-",
      impiantoComune: impianto?.comune || "-",
      operatorName,
      detectedAt: firebase.firestore.FieldValue.serverTimestamp(),
      reason: reason || "not_confirmed_in_done_list"
    },
    createdAt: firebase.firestore.FieldValue.serverTimestamp()
  })));
}

function buildImpiantoWhatsAppPayload(impianto, options = {}) {
  const user = auth.currentUser;
  const isOnlyOrdinaria = hasOrdinario(impianto.codicePrezzo) && !hasStraordinario(impianto.codicePrezzo);
  const title = isOnlyOrdinaria
    ? "✅ MANUTENZIONE ORDINARIA ESEGUITA"
    : "✅ MANUTENZIONE ORDINARIA + STRAORDINARIA ESEGUITA";
  const doneAt = options.doneAt || impianto.doneAt || new Date();
  const doneInfo = formatDoneDateTime(doneAt);
  const date = doneInfo.date === "-" ? new Date().toLocaleDateString("it-IT") : doneInfo.date;
  const time = doneInfo.time === "-" ? new Date().toLocaleTimeString("it-IT", { hour: "2-digit", minute: "2-digit", hour12: false }) : doneInfo.time;
  const linkedNotes = getCommessaNoteLinkedNotes(impianto);
  const segnalazioniLines = linkedNotes.length
    ? [
        "",
        "⚠️ A questo impianto è stata segnalata una criticità:",
        ...linkedNotes.map((note) => `${getCommessaNoteTitle(note)}\n${note.text || "-"}`)
      ]
    : [];
  const operatorName = options.operatorName || user?.displayName || user?.email || impianto.doneBy || "-";
  const message = [
    `${title} - Report operativo`,
    `🏗️ Impianto: ${impianto.denominazione || "-"}`,
    `📍 Comune: ${impianto.comune || "-"}`,
    `🛣️ Via: ${impianto.indirizzo || "-"}`,
    `🆔 ID SAP: ${impianto.idSap || "-"}`,
    ...(isOnlyOrdinaria ? [] : [`🛠️ Lavorazione straordinaria: ${impianto.lavorazioniRichieste || impianto.tipologiaIntervento || "-"}`]),
    `👷 Operatore: ${operatorName}`,
    `📅 Data: ${date}`,
    `🕒 Ora: ${time}`,
    ...segnalazioniLines
  ].join("\n");

  const encodedMessage = encodeURIComponent(message);
  return {
    message,
    appUrl: `whatsapp://send?text=${encodedMessage}`,
    webUrl: options.preferContactPicker === false
      ? `https://api.whatsapp.com/send?text=${encodedMessage}`
      : `https://wa.me/?text=${encodedMessage}`
  };
}

function openWhatsApp(impianto, options = {}) {
  const user = auth.currentUser;
  if (!user) {
    alert("Devi fare login.");
    return false;
  }

  const { message, appUrl, webUrl } = buildImpiantoWhatsAppPayload(impianto, options);
  const targetWindow = options?.targetWindow;
  const disableWebFallback = Boolean(options?.disableWebFallback);
  if (targetWindow && !targetWindow.closed) {
    try {
      targetWindow.location.replace(appUrl);
    } catch (error) {
      console.error("Errore apertura WhatsApp nella finestra target:", error);
    }
    if (!disableWebFallback) {
      setTimeout(() => {
        try {
          if (!targetWindow.closed) targetWindow.location.replace(webUrl);
        } catch (error) {
          console.error("Errore fallback WhatsApp nella finestra target:", error);
        }
      }, 700);
    }
    return true;
  }
  const target = options?.target || "_blank";
  const opened = safeOpenWhatsAppMessage(message, {
    appUrl,
    webUrl,
    target,
    usePopup: target !== "_self"
  });
  if (!opened && !disableWebFallback) {
    openExternalUrl(webUrl, { target: "_blank", features: "noopener", allowSameWindowFallback: true });
  }
  return Boolean(opened);
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
  const opened = safeOpenWhatsAppMessage(message);
  ui.impiantoReportFeedback.textContent = opened
    ? "WhatsApp aperto con la segnalazione pronta da inviare."
    : "Impossibile aprire WhatsApp automaticamente su questo dispositivo.";
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
    const opened = safeOpenWhatsAppMessage(waText, { phone: GPS_APPROVAL_PHONE });
    if (!opened) alert("Richiesta creata, ma non è stato possibile aprire WhatsApp automaticamente.");

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

async function notifyAdminsForImpiantoDoneRecovery(impianto, reason = "") {
  const adminUsers = platformUsers.filter((user) => adminEmails.has(normalizeEmail(user.email)));
  if (!adminUsers.length) return;
  const commessaId = String(selectedCommessaId || "").trim();
  const impiantoIds = getImpiantoDocIds(impianto).filter(Boolean);
  if (!commessaId || !impiantoIds.length) return;
  const operatorName = currentUser?.displayName || currentUser?.email || "Operatore";
  const text = [
    "⚠️ Recupero stato impianto richiesto",
    `Operatore: ${operatorName}`,
    `Commessa: ${selectedCommessaName || "Commessa"}`,
    `Impianto: ${impianto.denominazione || "Impianto"}`,
    reason ? `Dettaglio: ${reason}` : "Dettaglio: passaggio automatico ai FATTI non riuscito.",
    "Premi il pulsante per spostare l'impianto nei FATTI."
  ].join("\n");
  await Promise.all(adminUsers.map((adminUser) => db.collection("chatMessages").add({
    type: "text",
    text,
    recipientId: adminUser.id,
    senderId: currentUser?.uid || "",
    senderName: operatorName,
    senderEmail: currentUser?.email || "",
    kind: "system",
    metadata: {
      type: "impianto_done_recovery",
      action: "move_done",
      commessaId,
      commessaName: selectedCommessaName || "Commessa",
      impiantoIds,
      impiantoName: impianto.denominazione || "Impianto",
      impiantoKey: buildImpiantoKey(impianto)
    },
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

  if (!safeOpenWhatsAppMessage(message)) alert("Impossibile aprire WhatsApp su questo dispositivo.");
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
  if (!safeOpenWhatsAppMessage(message)) alert("Impossibile aprire WhatsApp su questo dispositivo.");
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

function getTimestampDate(value) {
  if (!value) return null;
  if (typeof value.toDate === "function") return value.toDate();
  if (typeof value.seconds === "number") return new Date(value.seconds * 1000);
  const parsed = new Date(value);
  return Number.isNaN(parsed.getTime()) ? null : parsed;
}

function addOperatorPositionMarkerToLayer(position, layer) {
  return L.marker([position.lat, position.lng], {
    icon: L.divIcon({
      className: "",
      html: "<div class='marker-operator' aria-hidden='true'>🦺</div>",
      iconSize: [16, 16],
      iconAnchor: [8, 8]
    })
  }).addTo(layer);
}

function getOperatorPositionsForMap() {
  if (!currentUserPos || !currentUser) return [];
  const byId = new Map();
  if (currentUser && currentUserPos) {
    const currentAssignment = getCurrentOperatorPositionAssignment();
    byId.set(currentUser.uid, {
      ...(byId.get(currentUser.uid) || {}),
      id: currentUser.uid,
      uid: currentUser.uid,
      email: currentUser.email || "",
      displayName: currentUser.displayName || currentUser.email || "Utente",
      operatorName: currentAssignment.operatorName || currentUser.displayName || currentUser.email || "Utente",
      ...currentAssignment,
      lat: currentUserPos.lat,
      lng: currentUserPos.lng,
      accuracy: currentUserPos.accuracy || 0,
      updatedAt: new Date()
    });
  }
  return Array.from(byId.values()).slice(0, 1).filter((position) => {
    const lat = Number(position.lat);
    const lng = Number(position.lng);
    if (!Number.isFinite(lat) || !Number.isFinite(lng)) return false;
    position.lat = lat;
    position.lng = lng;
    position.accuracy = Number(position.accuracy || 0);
    return true;
  });
}

function renderOperatorPositionMarkers(bounds) {
  getOperatorPositionsForMap().forEach((position) => {
    addOperatorPositionMarkerToLayer(position, markerLayer);
    addOperatorPositionMarkerToLayer(position, fullscreenMarkerLayer);
    bounds.push([position.lat, position.lng]);
  });
}

function toggleOperatorPositionsVisibility() {
  return;
}


function buildMapMarkerSequence(impianti = []) {
  return impianti
    .map((impianto) => ({ impianto, key: buildImpiantoKey(impianto), lat: Number(impianto.gpsY), lng: Number(impianto.gpsX) }))
    .filter((row) => row.key && Number.isFinite(row.lat) && Number.isFinite(row.lng))
    .sort((a, b) => {
      if (Math.abs(a.lat - b.lat) > 0.000001) return b.lat - a.lat;
      if (Math.abs(a.lng - b.lng) > 0.000001) return a.lng - b.lng;
      return String(a.key).localeCompare(String(b.key), "it");
    })
    .reduce((acc, row, index) => acc.set(row.key, index + 1), new Map());
}

function getMapMarkerNumberForImpianto(impianto) {
  if (!impianto) return null;
  return mapMarkerSequenceByKey.get(buildImpiantoKey(impianto)) || null;
}

function buildImpiantoMarkerBadge(impianto) {
  const markerClass = getMarkerClass(impianto);
  const markerNumber = getMapMarkerNumberForImpianto(impianto);
  if (!Number.isFinite(markerNumber)) return "";
  return `<span class="marker-pin-badge" aria-hidden="true"><span class="marker-pin ${markerClass}"><span class="marker-pin-number">${markerNumber}</span></span></span>`;
}

function focusImpiantoByMapNumber(rawNumber, targetMap = map) {
  const markerNumber = Number.parseInt(String(rawNumber || "").trim(), 10);
  if (!Number.isFinite(markerNumber) || markerNumber < 1) {
    alert("Numero impianto non trovato");
    return;
  }
  const match = currentImpianti.find((impianto) => getMapMarkerNumberForImpianto(impianto) === markerNumber);
  if (!match || match.gpsY == null || match.gpsX == null) {
    alert("Numero impianto non trovato");
    return;
  }
  const key = buildImpiantoKey(match);
  const markerMap = targetMap === fullscreenMap ? fullscreenImpiantoMarkerByKey : impiantoMarkerByKey;
  const marker = markerMap.get(key);
  targetMap.setView([match.gpsY, match.gpsX], Math.max(targetMap.getZoom(), 15), { animate: true });
  if (marker?.openPopup) marker.openPopup();
  selectImpiantoForMapDetail(match);
}

function renderMap() {
  clearMap();

  const bounds = [];
  mapMarkerSequenceByKey = buildMapMarkerSequence(currentImpianti);
  let markerForActiveFullscreenPopup = null;

  currentImpianti.forEach((impianto) => {
    const impiantoKey = buildImpiantoKey(impianto);
    const marker = addImpiantoMarkerToMapLayer(impianto, markerLayer, map);
    if (marker) impiantoMarkerByKey.set(impiantoKey, marker);
    const fullscreenMarker = addImpiantoMarkerToMapLayer(impianto, fullscreenMarkerLayer, fullscreenMap);
    if (fullscreenMarker) {
      if (impiantoKey) fullscreenImpiantoMarkerByKey.set(impiantoKey, fullscreenMarker);
      bounds.push([impianto.gpsY, impianto.gpsX]);
      if (impiantoKey && impiantoKey === selectedFullscreenImpiantoId) markerForActiveFullscreenPopup = fullscreenMarker;
    }
  });

  renderOperatorPositionMarkers(bounds);

  if (bounds.length > 0 && !mainMapViewState.hasUserMoved) {
    map.fitBounds(bounds, { padding: [24, 24], maxZoom: 14 });
    const center = map.getCenter();
    mainMapViewState.center = [center.lat, center.lng];
    mainMapViewState.zoom = map.getZoom();
    mainMapViewState.hasUserMoved = true;
  } else {
    map.setView(mainMapViewState.center, mainMapViewState.zoom, { animate: false });
  }
  fullscreenMap.setView(mainMapViewState.center, mainMapViewState.zoom, { animate: false });
  syncSelectedImpiantoDetailAfterRefresh(markerForActiveFullscreenPopup);
  preloadCommessaWeatherForVisibleImpianti();
}

function syncSelectedImpiantoDetailAfterRefresh(markerForSelectedImpianto) {
  if (!selectedImpiantoId) return;
  const latestImpianto = findCurrentImpiantoByKey(selectedImpiantoId);
  if (!latestImpianto) {
    closeSelectedImpiantoDetail({ closePopup: true });
    return;
  }
  selectedImpiantoData = { ...latestImpianto };
  selectedFullscreenImpiantoId = selectedImpiantoId;
  renderSelectedImpiantoDetailPanel();
  keepSelectedFullscreenPopupOpen(markerForSelectedImpianto);
}

function keepSelectedFullscreenPopupOpen(markerForSelectedImpianto) {
  if (!selectedFullscreenImpiantoId || !markerForSelectedImpianto) return;
  const reopenPopup = () => {
    if (!selectedFullscreenImpiantoId || !fullscreenMarkerLayer.hasLayer(markerForSelectedImpianto) || !markerForSelectedImpianto.getPopup?.()) return;
    markerForSelectedImpianto.openPopup();
  };
  requestAnimationFrame(reopenPopup);
  setTimeout(reopenPopup, 80);
}

function addImpiantoMarkerToMapLayer(impianto, targetLayer, targetMap = map) {
  if (impianto.gpsY == null || impianto.gpsX == null) return null;

  const markerClass = getMarkerClass(impianto);
  const markerSequence = mapMarkerSequenceByKey.get(buildImpiantoKey(impianto));
  const marker = L.marker([impianto.gpsY, impianto.gpsX], {
    icon: L.divIcon({
      className: "",
      html: `<div class="marker-pin ${markerClass}">${Number.isFinite(markerSequence) ? `<span class="marker-pin-number">${markerSequence}</span>` : ""}</div>`,
      iconSize: [18, 18],
      iconAnchor: [9, 9]
    })
  });
  if (targetMap !== fullscreenMap) {
    const tipo = impianto.tipoManutenzione || classifyTipoManutenzione(impianto.codicePrezzo);
    marker.bindPopup(buildImpiantoMapPopup(impianto, tipo), {
      autoClose: false,
      closeOnClick: false,
      closeButton: false,
      keepInView: true,
      maxWidth: 340,
      minWidth: 220,
      className: "impianto-map-popup"
    });
  }
  marker.on("click", () => {
    const nextImpiantoId = buildImpiantoKey(impianto);
    if (targetMap === fullscreenMap && selectedImpiantoId && selectedImpiantoId !== nextImpiantoId) fullscreenMap.closePopup();
    selectImpiantoForMapDetail(impianto);
    focusImpiantoInList(impianto, false);
  });
  marker.addTo(targetLayer);
  return marker;
}

function getImpiantoPopupData(impianto, tipo = "") {
  const doneInfo = formatDoneDateTime(impianto.doneAt);
  const idSap = impianto.idSap || impianto.codiceSap || "-";
  const via = impianto.indirizzo || impianto.descrizioneVia || impianto.via || "-";
  const tipologia = impianto.tipologiaImpianto || impianto.tipoImpianto || impianto.tipologiaIntervento || impianto.lavorazioniRichieste || tipo || "-";
  const operatore = impianto.doneBy || impianto.operatore || impianto.operator || impianto.navigatedBy || "-";
  const squadra = impianto.squadra || impianto.squadraAssegnata || impianto.team || "";
  return {
    idSap,
    via,
    tipologia,
    stato: impianto.done ? "Fatto" : "Da fare",
    dataFatto: doneInfo.date === "-" ? "-" : `${doneInfo.date} ${doneInfo.time}`,
    operatoreSquadra: [operatore, squadra].filter((value) => value && value !== "-").join(" • ") || "-",
    coordinates: impianto.gpsY != null && impianto.gpsX != null
      ? `${Number(impianto.gpsY).toFixed(6)}, ${Number(impianto.gpsX).toFixed(6)}`
      : "-"
  };
}

function buildImpiantoPositionUrl(impianto) {
  if (impianto.gpsY == null || impianto.gpsX == null) return "";
  return `https://www.google.com/maps/search/?api=1&query=${encodeURIComponent(`${impianto.gpsY},${impianto.gpsX}`)}`;
}

function buildImpiantoAppUrl(impianto) {
  const impiantoKey = buildImpiantoKey(impianto);
  const params = new URLSearchParams();
  if (selectedCommessaId) params.set("commessa", selectedCommessaId);
  if (impiantoKey) params.set("impianto", impiantoKey);
  const hash = params.toString() || "home";
  return `${window.location.origin}${window.location.pathname}#${hash}`;
}

function buildFullscreenImpiantoWhatsAppMessage(impianto) {
  const tipo = impianto.tipoManutenzione || classifyTipoManutenzione(impianto.codicePrezzo);
  const data = getImpiantoPopupData(impianto, tipo);
  const linkedNotes = getCommessaNoteLinkedNotes(impianto);
  const linkedNotesLines = linkedNotes.length
    ? linkedNotes.flatMap((note) => [
        `- ${getCommessaNoteTitle(note)}`,
        ...(String(note.text || "").trim() ? [String(note.text || "").trim()] : [])
      ])
    : ["-"];
  const positionUrl = buildImpiantoPositionUrl(impianto);
  const appUrl = buildImpiantoAppUrl(impianto);
  return [
    "Ti inoltro i dettagli dell’impianto:",
    "",
    `Nome impianto: ${impianto.denominazione || "-"}`,
    `ID SAP: ${data.idSap}`,
    `Comune: ${impianto.comune || "-"}`,
    `Indirizzo / Via: ${data.via}`,
    `Tipologia impianto: ${data.tipologia}`,
    `Stato lavoro: ${data.stato}`,
    `Data fatto: ${data.dataFatto}`,
    `Operatore / Squadra: ${data.operatoreSquadra}`,
    "Segnalazione collegata:",
    ...linkedNotesLines,
    "",
    `Posizione impianto: ${data.coordinates}`,
    positionUrl ? `[📍 Apri posizione impianto](${positionUrl})` : "📍 Apri posizione impianto: posizione non disponibile",
    `[🔗 Apri impianto nell’app](${appUrl})`
  ].join("\n");
}

function openFullscreenImpiantoWhatsApp(impianto) {
  const message = buildFullscreenImpiantoWhatsAppMessage(impianto);
  if (!safeOpenWhatsAppMessage(message)) alert("Impossibile aprire WhatsApp su questo dispositivo.");
}

function buildImpiantoMapPopup(impianto, tipo, options = {}) {
  const impiantoKey = buildImpiantoKey(impianto);
  const linkedNotes = getCommessaNoteLinkedNotes(impianto);
  const popupData = getImpiantoPopupData(impianto, tipo);
  const noteImpianto = getImpiantoPopupNotes(impianto);
  const showSapInHeader = Boolean(options?.showSapInHeader);
  const whatsappAction = options?.fullscreenWhatsApp ? "fullscreen-whatsapp" : "whatsapp";
  const linkedNotesMarkup = linkedNotes.length
    ? linkedNotes.map((note) => `
        <button type="button" class="map-popup-note-btn" data-map-popup-action="note" data-note-id="${escapeHTML(note.id || "")}">
          ${escapeHTML(getCommessaNoteTitle(note))}
        </button>
      `).join("")
    : "<span>-</span>";
  const headerSubtitle = showSapInHeader
    ? `<p class="map-popup-subtitle">ID SAP: ${escapeHTML(popupData.idSap)}</p>`
    : "";
  const idSapDetail = showSapInHeader
    ? ""
    : `<div><dt>ID SAP</dt><dd>${escapeHTML(popupData.idSap)}</dd></div>`;

  return `
    <div class="map-popup-card" data-impianto-key="${escapeHTML(impiantoKey)}">
      <div class="map-popup-header">
        <div class="map-popup-title">
          <h3>${buildImpiantoMarkerBadge(impianto)}${escapeHTML(impianto.denominazione || "Impianto")}</h3>
          ${buildImpiantoWeatherBadgeMarkup(impianto)}
          ${headerSubtitle}
        </div>
        <button type="button" class="map-popup-close-btn" data-map-popup-action="close" aria-label="Chiudi popup dettaglio impianto" title="Chiudi">×</button>
      </div>
      <div class="map-popup-scroll">
        <dl class="map-popup-details">
          ${idSapDetail}
          <div><dt>Comune</dt><dd>${escapeHTML(impianto.comune || "-")}</dd></div>
          <div><dt>Indirizzo / via</dt><dd>${escapeHTML(popupData.via)}</dd></div>
          <div><dt>Tipologia impianto</dt><dd>${escapeHTML(popupData.tipologia)}</dd></div>
          <div><dt>Stato lavoro</dt><dd>${escapeHTML(popupData.stato)}</dd></div>
          <div><dt>Data fatto</dt><dd>${escapeHTML(popupData.dataFatto)}</dd></div>
          <div><dt>Operatore / squadra</dt><dd>${escapeHTML(popupData.operatoreSquadra)}</dd></div>
          <div><dt>Segnalazione collegata</dt><dd class="map-popup-notes-list">${linkedNotesMarkup}</dd></div>
          <div><dt>Note impianto</dt><dd>${escapeHTML(noteImpianto || "-")}</dd></div>
          <div><dt>Coordinate GPS</dt><dd>${escapeHTML(popupData.coordinates)}</dd></div>
        </dl>
      </div>
      <div class="map-popup-actions">
        <button type="button" class="btn btn-small btn-primary" data-map-popup-action="navigate" data-impianto-key="${escapeHTML(impiantoKey)}">NAVIGA</button>
        <button type="button" class="btn btn-small btn-whatsapp" data-map-popup-action="${escapeHTML(whatsappAction)}" data-impianto-key="${escapeHTML(impiantoKey)}">WHATSAPP</button>
        <button type="button" class="btn btn-small" data-map-popup-action="detail" data-impianto-key="${escapeHTML(impiantoKey)}">DETTAGLIO IMPIANTO</button>
      </div>
    </div>
  `;
}

function selectImpiantoForMapDetail(impianto) {
  const key = buildImpiantoKey(impianto);
  if (!key) return;
  selectedImpiantoId = key;
  selectedFullscreenImpiantoId = key;
  selectedImpiantoData = { ...impianto };
  renderSelectedImpiantoDetailPanel();
}

function closeSelectedImpiantoDetail({ closePopup = false } = {}) {
  selectedImpiantoId = "";
  selectedImpiantoData = null;
  selectedFullscreenImpiantoId = "";
  ui.mainMapImpiantoDetailPanel?.classList.add("hidden");
  ui.mapImpiantoDetailPanel?.classList.add("hidden");
  if (ui.mainMapImpiantoDetailBody) ui.mainMapImpiantoDetailBody.innerHTML = "";
  if (ui.mapImpiantoDetailBody) ui.mapImpiantoDetailBody.innerHTML = "";
  if (closePopup) fullscreenMap.closePopup();
}

function renderSelectedImpiantoDetailPanel() {
  const panels = [
    { panel: ui.mainMapImpiantoDetailPanel, body: ui.mainMapImpiantoDetailBody },
    { panel: ui.mapImpiantoDetailPanel, body: ui.mapImpiantoDetailBody }
  ].filter((entry) => entry.panel && entry.body);
  if (!panels.length) return;
  if (!selectedImpiantoId || !selectedImpiantoData) {
    panels.forEach(({ panel, body }) => {
      panel.classList.add("hidden");
      body.innerHTML = "";
    });
    return;
  }
  const tipo = selectedImpiantoData.tipoManutenzione || classifyTipoManutenzione(selectedImpiantoData.codicePrezzo);
  panels.forEach(({ panel, body }) => {
    const isFullscreenPanel = panel === ui.mapImpiantoDetailPanel;
    body.innerHTML = buildImpiantoMapPopup(selectedImpiantoData, tipo, {
      showSapInHeader: isFullscreenPanel,
      fullscreenWhatsApp: isFullscreenPanel
    });
    panel.classList.remove("hidden");
  });
  bindPersistentImpiantoDetailActions();
}

function bindPersistentImpiantoDetailActions() {
  [ui.mainMapImpiantoDetailPanel, ui.mapImpiantoDetailPanel].filter(Boolean).forEach((panel) => {
    panel.querySelectorAll("[data-map-popup-action='navigate']").forEach((button) => {
      button.addEventListener("click", async () => {
        const key = button.getAttribute("data-impianto-key") || selectedImpiantoId;
        const impianto = findCurrentImpiantoByKey(key) || selectedImpiantoData;
        if (!impianto) return;
        await navigateToImpianto(impianto);
      });
    });
    panel.querySelectorAll("[data-map-popup-action='whatsapp']").forEach((button) => {
      button.addEventListener("click", async () => {
        const key = button.getAttribute("data-impianto-key") || selectedImpiantoId;
        const impianto = findCurrentImpiantoByKey(key) || selectedImpiantoData;
        if (!impianto) return;
        triggerImpiantoWhatsAppAction(impianto);
      });
    });
    panel.querySelectorAll("[data-map-popup-action='fullscreen-whatsapp']").forEach((button) => {
      button.addEventListener("click", () => {
        const key = button.getAttribute("data-impianto-key") || selectedImpiantoId;
        const impianto = findCurrentImpiantoByKey(key) || selectedImpiantoData;
        if (!impianto) return;
        openFullscreenImpiantoWhatsApp(impianto);
      });
    });
    panel.querySelectorAll("[data-map-popup-action='detail']").forEach((button) => {
      button.addEventListener("click", () => {
        const key = button.getAttribute("data-impianto-key") || selectedImpiantoId;
        const impianto = findCurrentImpiantoByKey(key) || selectedImpiantoData;
        if (!impianto) return;
        focusImpiantoInList(impianto, true);
        closeMapFullscreenPage();
      });
    });
    panel.querySelectorAll("[data-map-popup-action='note']").forEach((button) => {
      button.addEventListener("click", () => {
        const noteId = button.getAttribute("data-note-id") || "";
        const note = currentCommessaNotes.find((item) => item.id === noteId);
        if (!note) return;
        openCommessaNotesPage();
        setTimeout(() => openCommessaNoteDetail(note), 50);
      });
    });
    panel.querySelectorAll("[data-map-popup-action='close']").forEach((button) => {
      button.addEventListener("click", (clickEvent) => {
        clickEvent.preventDefault();
        clickEvent.stopPropagation();
        closeSelectedImpiantoDetail({ closePopup: true });
      });
    });
  });
}

function getImpiantoPopupNotes(impianto) {
  return [
    impianto.noteImpianto,
    impianto.note,
    impianto.notes,
    impianto.annotazioni,
    impianto.descrizione,
    impianto.descrizioneImpianto
  ].map((value) => String(value || "").trim()).find(Boolean) || "";
}

function findCurrentImpiantoByKey(key) {
  return currentImpianti.find((item) => buildImpiantoKey(item) === key);
}

function bindImpiantoMapPopupActions(event, popupMap) {
  const popupElement = event.popup?.getElement();
  if (!popupElement) return;
  const card = popupElement.querySelector(".map-popup-card[data-impianto-key]");
  const popupKey = card?.getAttribute("data-impianto-key") || "";

  popupElement.querySelectorAll("[data-map-popup-action='navigate']").forEach((button) => {
    button.addEventListener("click", async () => {
      const key = button.getAttribute("data-impianto-key") || popupKey;
      const impianto = findCurrentImpiantoByKey(key);
      if (!impianto) return;
      await navigateToImpianto(impianto);
    });
  });
  popupElement.querySelectorAll("[data-map-popup-action='whatsapp']").forEach((button) => {
    button.addEventListener("click", async () => {
      const key = button.getAttribute("data-impianto-key") || popupKey;
      const impianto = findCurrentImpiantoByKey(key);
      if (!impianto) return;
      triggerImpiantoWhatsAppAction(impianto);
    });
  });
  popupElement.querySelectorAll("[data-map-popup-action='fullscreen-whatsapp']").forEach((button) => {
    button.addEventListener("click", () => {
      const key = button.getAttribute("data-impianto-key") || popupKey;
      const impianto = findCurrentImpiantoByKey(key);
      if (!impianto) return;
      openFullscreenImpiantoWhatsApp(impianto);
    });
  });
  popupElement.querySelectorAll("[data-map-popup-action='detail']").forEach((button) => {
    button.addEventListener("click", () => {
      const key = button.getAttribute("data-impianto-key") || popupKey;
      const impianto = findCurrentImpiantoByKey(key);
      if (!impianto) return;
      selectedFullscreenImpiantoId = popupMap === fullscreenMap ? key : selectedFullscreenImpiantoId;
      focusImpiantoInList(impianto, true);
      if (popupMap === fullscreenMap) closeMapFullscreenPage();
    });
  });
  popupElement.querySelectorAll("[data-map-popup-action='note']").forEach((button) => {
    button.addEventListener("click", () => {
      const noteId = button.getAttribute("data-note-id") || "";
      const note = currentCommessaNotes.find((item) => item.id === noteId);
      if (!note) return;
      openCommessaNotesPage();
      setTimeout(() => openCommessaNoteDetail(note), 50);
    });
  });
  popupElement.querySelectorAll("[data-map-popup-action='close']").forEach((button) => {
    button.addEventListener("click", (clickEvent) => {
      clickEvent.preventDefault();
      clickEvent.stopPropagation();
      if (popupMap === fullscreenMap && popupKey === selectedImpiantoId) closeSelectedImpiantoDetail({ closePopup: false });
      popupMap.closePopup(event.popup);
    });
  });
}

map.on("popupopen", (event) => bindImpiantoMapPopupActions(event, map));
fullscreenMap.on("popupopen", (event) => bindImpiantoMapPopupActions(event, fullscreenMap));
map.on("moveend zoomend", () => preloadCommessaWeatherForVisibleImpianti());
fullscreenMap.on("moveend zoomend", () => preloadImpiantiWeather(getVisibleMapImpianti(fullscreenMap, currentImpianti), { limit: IMPIANTO_WEATHER_REFRESH_LIMIT, preferNearest: true }));
globalMap.on("moveend zoomend", () => preloadImpiantiWeather(getVisibleMapImpianti(globalMap, globalImpianti)));

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
  impiantoMarkerByKey.clear();
  fullscreenImpiantoMarkerByKey.clear();
  markerLayer.clearLayers();
  fullscreenMarkerLayer.clearLayers();
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
    <div class="item-actions">
      <button id="fuel-open-pin-doc-btn" class="btn" type="button">📌 PIN carburante</button>
    </div>
  `;
  const openPinBtn = document.getElementById("fuel-open-pin-doc-btn");
  openPinBtn?.addEventListener("click", () => {
    openPrivateDocsPage();
    applyPrivateDocPreset("pin");
    setTimeout(() => {
      ui.privateDocsForm?.scrollIntoView({ behavior: "smooth", block: "start" });
      ui.privateDocsName?.focus();
    }, 50);
  });
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

function shouldPublishOperatorPosition(coords, options = {}) {
  if (!currentUser || !coords) return false;
  if (options.force) return true;
  const now = Date.now();
  if (!lastPublishedUserPos) return true;
  const movedMeters = haversine(
    lastPublishedUserPos.lat,
    lastPublishedUserPos.lng,
    coords.latitude,
    coords.longitude
  ) * 1000;
  return movedMeters >= 20 || now - lastPositionPublishAt >= 15 * 1000;
}

function getCurrentOperatorPositionAssignment() {
  const dateKey = getActiveSquadreDateKey();
  const assignment = getCurrentUserAssignedCommesseForDate(dateKey)[0];
  const matchedRow = assignment?.matchedRows?.[0] || null;
  return {
    operatorName: matchedRow?.matchedName || currentUser?.displayName || currentUser?.email || "Operatore",
    squadraIndex: matchedRow?.squadraIndex || "",
    squadraLabel: matchedRow?.squadraLabel || "",
    squadraName: matchedRow?.row?.nome || matchedRow?.row?.name || "",
    commessaId: assignment?.commessaId || selectedCommessaId || "",
    commessaName: assignment?.commessaName || selectedCommessaName || "",
    riferimentoData: dateKey || ""
  };
}

async function publishCurrentOperatorPosition(coords, options = {}) {
  if (!shouldPublishOperatorPosition(coords, options)) return;
  lastPositionPublishAt = Date.now();
  lastPublishedUserPos = { lat: coords.latitude, lng: coords.longitude };
  try {
    const assignment = getCurrentOperatorPositionAssignment();
    await db.collection("operatorPositions").doc(currentUser.uid).set({
      userId: currentUser.uid,
      uid: currentUser.uid,
      email: currentUser.email || "",
      displayName: currentUser.displayName || currentUser.email || "Utente",
      nomeOperatore: assignment.operatorName || currentUser.displayName || currentUser.email || "Operatore",
      operatorName: assignment.operatorName || currentUser.displayName || currentUser.email || "Operatore",
      squadra: assignment.squadraLabel || assignment.squadraName || "",
      squadraIndex: assignment.squadraIndex || "",
      squadraLabel: assignment.squadraLabel || "",
      squadraName: assignment.squadraName || "",
      commessaAttiva: assignment.commessaName || "",
      commessaId: assignment.commessaId || "",
      commessaName: assignment.commessaName || "",
      riferimentoData: assignment.riferimentoData || "",
      lat: Number(coords.latitude),
      lng: Number(coords.longitude),
      latitude: Number(coords.latitude),
      longitude: Number(coords.longitude),
      accuracy: Number(coords.accuracy || 0),
      stato: "online",
      status: "online",
      updatedAt: firebase.firestore.FieldValue.serverTimestamp(),
      lastUpdatedAt: firebase.firestore.FieldValue.serverTimestamp()
    }, { merge: true });
    await db.collection("platformUsers").doc(currentUser.uid).set({
      uid: currentUser.uid,
      email: currentUser.email || "",
      displayName: currentUser.displayName || currentUser.email || "Utente",
      lastSeenAt: firebase.firestore.FieldValue.serverTimestamp()
    }, { merge: true });
  } catch (error) {
    console.warn("Aggiornamento posizione operatore non riuscito:", error);
  }
}

function markCurrentOperatorOffline() {
  if (!currentUser) return;
  db.collection("operatorPositions").doc(currentUser.uid).set({
    userId: currentUser.uid,
    uid: currentUser.uid,
    stato: "offline",
    status: "offline",
    updatedAt: firebase.firestore.FieldValue.serverTimestamp(),
    lastUpdatedAt: firebase.firestore.FieldValue.serverTimestamp()
  }, { merge: true }).catch((error) => console.warn("Aggiornamento stato operatore non riuscito:", error));
}

function initGeolocation(options = {}) {
  if (!navigator.geolocation) {
    ui.gpsStatus.textContent = "Geolocalizzazione non supportata dal browser.";
    return;
  }
  if (options.forcePublishCurrent && latestGeolocationCoords && currentUser) {
    publishCurrentOperatorPosition(latestGeolocationCoords, { force: true });
    options.forcePublishCurrent = false;
  }
  if (geolocationWatchId != null) navigator.geolocation.clearWatch(geolocationWatchId);

  const onPosition = (pos) => {
    latestGeolocationCoords = pos.coords;
    currentUserPos = {
      lat: pos.coords.latitude,
      lng: pos.coords.longitude,
      accuracy: pos.coords.accuracy || 0
    };
    publishCurrentOperatorPosition(pos.coords, { force: Boolean(options.forcePublishCurrent) });
    options.forcePublishCurrent = false;
    evaluateTimbraturaReminders();
    ui.gpsStatus.textContent = "Posizione attiva: impianti ordinati per distanza.";
    renderImpianti();
    renderMap();
    evaluateImpiantoProximityAlerts();
  };

  navigator.geolocation.getCurrentPosition((pos) => {
    onPosition(pos);
    fetchWeather();
  }, () => {
    ui.gpsStatus.textContent = "Posizione non disponibile";
    fetchWeather();
  }, {
    enableHighAccuracy: true,
    timeout: 8000
  });

  geolocationWatchId = navigator.geolocation.watchPosition(onPosition, () => {
    ui.gpsStatus.textContent = "Posizione non disponibile";
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

function buildTimbraturaReminderLocalKey(dateKey, fascia) {
  return `heraTimbraturaReminder:${dateKey}:${fascia}`;
}

function getLocalDateKey(date) {
  const year = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, "0");
  const day = String(date.getDate()).padStart(2, "0");
  return `${year}-${month}-${day}`;
}

function getMinutesSinceMidnight(date) {
  return date.getHours() * 60 + date.getMinutes();
}

function evaluateTimbraturaReminders(now = new Date()) {
  if (!currentUserPos) return;
  const distanceMeters = haversine(currentUserPos.lat, currentUserPos.lng, TIMBRATURA_TARGET_LAT, TIMBRATURA_TARGET_LNG) * 1000;
  if (!Number.isFinite(distanceMeters) || distanceMeters > TIMBRATURA_RADIUS_M) return;

  const minutes = getMinutesSinceMidnight(now);
  let fascia = "";
  let notificationText = "";
  if (minutes >= TIMBRATURA_ENTRATA_START_MIN && minutes <= TIMBRATURA_ENTRATA_END_MIN) {
    fascia = "entrata";
    notificationText = "RICORDATI DI TIMBRARE L ENTRATA";
  } else if (minutes >= TIMBRATURA_USCITA_START_MIN && minutes <= TIMBRATURA_USCITA_END_MIN) {
    fascia = "uscita";
    notificationText = "RICORDATI DI TIMBRARE L USCITA";
  } else {
    return;
  }

  const dateKey = getLocalDateKey(now);
  const reminderKey = buildTimbraturaReminderLocalKey(dateKey, fascia);
  if (localStorage.getItem(reminderKey) === "1") return;
  localStorage.setItem(reminderKey, "1");
  showLocalNotification("Hera App", {
    body: notificationText,
    tag: `hera-timbratura-${dateKey}-${fascia}`,
    renotify: false,
    data: { url: "./index.html" }
  }).catch((error) => {
    console.warn("Invio notifica timbratura non riuscito:", error);
  });
}

const CIVIL_PROTECTION_ALERT_PAGE = "https://mappe.protezionecivile.gov.it/it/mappe-rischi/bollettino-di-criticita/";
const CIVIL_PROTECTION_GITHUB_API = "https://api.github.com/repos/pcm-dpc/DPC-Bollettini-Criticita-Idrogeologica-Idraulica/contents/files/xml?ref=master";
const METEO_3B_BASE_URL = "https://www.3bmeteo.com/meteo/italia";
const ALERT_LEVEL_META = {
  green: { rank: 0, emoji: "🟢", className: "alert-green", label: "Nessuna allerta" },
  yellow: { rank: 1, emoji: "🟡", className: "alert-yellow", label: "Allerta Protezione Civile" },
  orange: { rank: 2, emoji: "🟠", className: "alert-orange", label: "Allerta Protezione Civile" },
  red: { rank: 3, emoji: "🔴", className: "alert-red", label: "Allerta Protezione Civile" }
};
const ALERT_KEYWORDS = [
  { key: "temporali", label: "Temporali", patterns: ["temporali", "temporale"] },
  { key: "vento", label: "vento", patterns: ["vento", "venti", "burrasca"] },
  { key: "neve", label: "neve", patterns: ["neve", "nevicate"] },
  { key: "ghiaccio", label: "ghiaccio", patterns: ["ghiaccio", "gelate"] },
  { key: "alluvione", label: "alluvione", patterns: ["idraulico", "idrogeologico", "alluvione", "allagamenti"] },
  { key: "nebbia", label: "nebbia", patterns: ["nebbia", "nebbie"] },
  { key: "caldo", label: "caldo estremo", patterns: ["caldo", "ondate di calore", "temperature elevate"] }
];
const NAVIGATION_WEATHER_STRONG_WIND_KMH = 50;
const NAVIGATION_WEATHER_STRONG_GUST_KMH = 70;
const NAVIGATION_WEATHER_RELEVANT_RAIN_MM = 5;
const NAVIGATION_WEATHER_LIGHT_RAIN_MAX_MM = 2;
const NAVIGATION_WEATHER_NEXT_HOUR_PROBABILITY = 40;
const NAVIGATION_WEATHER_THUNDER_CODES = new Set([95, 96, 99]);
const NAVIGATION_WEATHER_RAIN_CODES = new Set([51, 53, 55, 56, 57, 61, 63, 65, 66, 67, 80, 81, 82]);
const IMPIANTO_WEATHER_CACHE_TTL_MS = 5 * 60 * 1000;
const IMPIANTO_WEATHER_REFRESH_LIMIT = 40;
const IMPIANTO_WEATHER_BATCH_SIZE = 2;
const IMPIANTO_WEATHER_COORDINATE_PRECISION = 5;
const IMPIANTO_WEATHER_LOCAL_CACHE_MAX_ENTRIES = 400;


function buildWeatherForecastRequestParams(target, { operational = false } = {}) {
  const baseParams = {
    latitude: String(target.lat),
    longitude: String(target.lon),
    current: "temperature_2m,wind_speed_10m,weather_code",
    hourly: "temperature_2m,precipitation_probability,snowfall,visibility,weather_code,wind_speed_10m",
    forecast_days: "2"
  };

  if (!operational) return baseParams;

  return {
    ...baseParams,
    current: "temperature_2m,apparent_temperature,relative_humidity_2m,precipitation,rain,showers,weather_code,wind_speed_10m,wind_direction_10m,wind_gusts_10m",
    minutely_15: "precipitation,weather_code",
    hourly: "temperature_2m,precipitation_probability,precipitation,rain,showers,snowfall,visibility,weather_code,apparent_temperature,wind_speed_10m,wind_direction_10m,wind_gusts_10m",
    forecast_hours: "12",
    forecast_minutely_15: "48",
    forecast_days: "1",
    timezone: "auto"
  };
}

async function fetchWeatherForecast(target, options = {}) {
  const params = new URLSearchParams(buildWeatherForecastRequestParams(target, options));
  const response = await fetch(`https://api.open-meteo.com/v1/forecast?${params.toString()}`, { cache: options.cache || "no-store" });
  if (!response.ok) throw new Error("meteo non disponibile");
  const data = await response.json();
  if (options.operational) validateImpiantoWeatherPayload(data, "Open-Meteo");
  return { ...data, provider: "Open-Meteo" };
}

async function fetchWeather() {
  const target = getWeatherTargetCoordinates();
  currentWeatherTarget = target;
  renderCivilProtectionAlert({ level: "green", label: "Verifica Protezione Civile...", url: CIVIL_PROTECTION_ALERT_PAGE, loading: true });

  try {
    const data = await fetchWeatherForecast(target);
    const current = data.current || {};
    const weatherLabel = weatherCodeLabel(current.weather_code);
    ui.weatherSummary.textContent = `${weatherLabel} • ${Math.round(current.temperature_2m ?? 0)}°C • vento ${Math.round(current.wind_speed_10m ?? 0)} km/h`;
    await renderWeatherDetails(data, target);
  } catch (error) {
    ui.weatherSummary.textContent = "Meteo non disponibile.";
    ui.weatherRisks.innerHTML = "<span class='weather-risk-chip'>⚠️ Nessun dato rischio disponibile</span>";
    renderCivilProtectionAlert({ level: "green", label: "Protezione Civile non disponibile", url: CIVIL_PROTECTION_ALERT_PAGE });
    ui.weatherDetails.innerHTML = "<p class='muted'>Impossibile caricare previsioni dettagliate.</p>";
  }
}

function getWeatherTargetCoordinates() {
  if (currentUserPos) return { lat: Number(currentUserPos.lat), lon: Number(currentUserPos.lng), source: "gps" };
  const gpsImpianti = currentImpianti
    .map((impianto) => ({ lat: Number(impianto.gpsY), lon: Number(impianto.gpsX) }))
    .filter((pos) => Number.isFinite(pos.lat) && Number.isFinite(pos.lon));
  if (gpsImpianti.length) {
    const sum = gpsImpianti.reduce((acc, pos) => ({ lat: acc.lat + pos.lat, lon: acc.lon + pos.lon }), { lat: 0, lon: 0 });
    return { lat: sum.lat / gpsImpianti.length, lon: sum.lon / gpsImpianti.length, source: "commessa" };
  }
  return { lat: 44.4949, lon: 11.3426, source: "fallback" };
}

async function renderWeatherDetails(data, target) {
  const times = (data.hourly && data.hourly.time) || [];
  const temps = (data.hourly && data.hourly.temperature_2m) || [];
  const rains = (data.hourly && data.hourly.precipitation_probability) || [];
  const snows = (data.hourly && data.hourly.snowfall) || [];
  const visibilities = (data.hourly && data.hourly.visibility) || [];
  const codes = (data.hourly && data.hourly.weather_code) || [];
  const winds = (data.hourly && data.hourly.wind_speed_10m) || [];
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

  const alert = await getCivilProtectionAlert(target, { temps, winds, snows, visibilities, codes });
  const riskChips = risks.map((risk) => `<span class='weather-risk-chip'>${escapeHTML(risk)}</span>`).join("");
  ui.weatherRisks.innerHTML = `${riskChips}${buildCivilProtectionAlertChip(alert)}`;
  renderCivilProtectionAlert(alert);

  const rows = times.slice(0, 12).map((time, idx) => {
    const hour = new Date(time).toLocaleTimeString("it-IT", { hour: "2-digit", minute: "2-digit", hour12: false });
    const visKm = ((Number(visibilities[idx]) || 0) / 1000).toFixed(1);
    const label = weatherCodeLabel(codes[idx]);
    return `<p><b>${hour}</b> • ${label} • 🌡️ ${Math.round(temps[idx] ?? 0)}°C • 🌧️ ${Math.round(rains[idx] ?? 0)}% • ❄️ ${Number(snows[idx] || 0).toFixed(1)} mm • 👁️ ${visKm} km</p>`;
  }).join("");
  ui.weatherDetails.innerHTML = rows || "<p class='muted'>Nessun dato meteo.</p>";
}

async function getCivilProtectionAlert(target, forecast = {}) {
  const region = await reverseGeocodeRegion(target).catch(() => "");
  const officialText = await fetchCivilProtectionOfficialText().catch(() => "");
  const officialAlert = parseCivilProtectionAlertText(officialText, region);
  const forecastAlert = buildOperationalForecastAlert(forecast);
  const alert = pickHighestAlert([officialAlert, forecastAlert]);
  return {
    ...alert,
    region,
    url: CIVIL_PROTECTION_ALERT_PAGE,
    label: alert.label || "Nessuna allerta"
  };
}

async function reverseGeocodeRegion(target) {
  const key = `heraWeatherRegion:${target.lat.toFixed(2)}:${target.lon.toFixed(2)}`;
  const cached = localStorage.getItem(key);
  if (cached) return cached;
  const url = `https://nominatim.openstreetmap.org/reverse?format=jsonv2&lat=${encodeURIComponent(target.lat)}&lon=${encodeURIComponent(target.lon)}&zoom=8&addressdetails=1`;
  const response = await fetch(url, { headers: { Accept: "application/json" } });
  if (!response.ok) return "";
  const data = await response.json();
  const region = normalizeItalianRegionName(data?.address?.state || data?.address?.region || "");
  if (region) localStorage.setItem(key, region);
  return region;
}

function normalizeItalianRegionName(region) {
  return String(region || "")
    .replace(/^regione\s+/i, "")
    .replace(/emilia-romagna/i, "Emilia Romagna")
    .replace(/trentino-alto adige\/südtirol/i, "Trentino Alto Adige")
    .trim();
}

async function fetchCivilProtectionOfficialText() {
  const cached = getCachedCivilProtectionText();
  if (cached) return cached;
  const githubText = await fetchLatestCivilProtectionXmlText().catch(() => "");
  if (githubText) {
    cacheCivilProtectionText(githubText);
    return githubText;
  }
  const response = await fetch(CIVIL_PROTECTION_ALERT_PAGE, { cache: "no-store" });
  if (!response.ok) return "";
  const html = await response.text();
  const doc = new DOMParser().parseFromString(html, "text/html");
  const text = doc.body?.textContent || html;
  cacheCivilProtectionText(text);
  return text;
}

async function fetchLatestCivilProtectionXmlText() {
  const response = await fetch(CIVIL_PROTECTION_GITHUB_API, { headers: { Accept: "application/vnd.github+json" } });
  if (!response.ok) return "";
  const files = await response.json();
  const latest = (Array.isArray(files) ? files : [])
    .filter((file) => String(file.name || "").toLowerCase().endsWith(".xml") && file.download_url)
    .sort((a, b) => String(b.name).localeCompare(String(a.name)))
    .at(0);
  if (!latest) return "";
  const xmlResponse = await fetch(latest.download_url, { cache: "no-store" });
  return xmlResponse.ok ? xmlResponse.text() : "";
}

function getCachedCivilProtectionText() {
  try {
    const cached = JSON.parse(localStorage.getItem("heraCivilProtectionAlertText") || "null");
    if (cached && Date.now() - Number(cached.savedAt || 0) < 60 * 60 * 1000) return String(cached.text || "");
  } catch (error) {
    console.warn("Cache Protezione Civile non leggibile:", error);
  }
  return "";
}

function cacheCivilProtectionText(text) {
  try {
    localStorage.setItem("heraCivilProtectionAlertText", JSON.stringify({ text, savedAt: Date.now() }));
  } catch (error) {
    console.warn("Cache Protezione Civile non salvata:", error);
  }
}

function parseCivilProtectionAlertText(text, region) {
  const normalizedText = String(text || "").replace(/<[^>]+>/g, "\n");
  const lines = normalizedText.split(/\r?\n/).map((line) => line.replace(/\s+/g, " ").trim()).filter(Boolean);
  const regionNeedle = normalizeForSearch(region);
  let currentLevel = "green";
  let currentPhenomenon = "Protezione Civile";
  const alerts = [];

  lines.forEach((line) => {
    const upper = line.toUpperCase();
    const level = levelFromText(upper);
    if (upper.includes("ALLERTA") || upper.includes("CRITICITA")) {
      currentLevel = level || currentLevel;
      currentPhenomenon = phenomenonFromText(line) || currentPhenomenon;
    }
    const lineHasRegion = regionNeedle && normalizeForSearch(line).includes(regionNeedle);
    const isNationalNoAlert = upper.includes("NESSUNA ALLERTA") || upper.includes("ASSENZA DI FENOMENI SIGNIFICATIVI");
    if (lineHasRegion && ALERT_LEVEL_META[currentLevel]?.rank > 0) {
      alerts.push({ level: currentLevel, label: buildAlertLabel(currentLevel, currentPhenomenon), phenomenon: currentPhenomenon });
    } else if (!regionNeedle && isNationalNoAlert) {
      alerts.push({ level: "green", label: "Nessuna allerta", phenomenon: "" });
    }
  });

  return pickHighestAlert(alerts.length ? alerts : [{ level: "green", label: "Nessuna allerta", phenomenon: "" }]);
}

function normalizeForSearch(value) {
  return String(value || "").normalize("NFD").replace(/[\u0300-\u036f]/g, "").toLowerCase();
}

function levelFromText(text) {
  if (text.includes("ROSSA") || text.includes("ELEVATA")) return "red";
  if (text.includes("ARANCIONE") || text.includes("MODERATA")) return "orange";
  if (text.includes("GIALLA") || text.includes("ORDINARIA")) return "yellow";
  if (text.includes("VERDE") || text.includes("NESSUNA ALLERTA")) return "green";
  return "";
}

function phenomenonFromText(text) {
  const normalized = normalizeForSearch(text);
  const match = ALERT_KEYWORDS.find((item) => item.patterns.some((pattern) => normalized.includes(normalizeForSearch(pattern))));
  return match?.label || "Protezione Civile";
}

function buildOperationalForecastAlert({ temps = [], winds = [], snows = [], visibilities = [], codes = [] } = {}) {
  const maxWind = Math.max(...winds.slice(0, 12).map((value) => Number(value) || 0), 0);
  const snowSum = snows.slice(0, 12).reduce((acc, value) => acc + (Number(value) || 0), 0);
  const minTemp = Math.min(...temps.slice(0, 12).map((value) => Number(value) || Number.MAX_SAFE_INTEGER));
  const maxTemp = Math.max(...temps.slice(0, 12).map((value) => Number(value) || -100), -100);
  const minVisibility = Math.min(...visibilities.slice(0, 12).map((value) => Number(value) || Number.MAX_SAFE_INTEGER));
  const hasStormCode = codes.slice(0, 12).some((value) => [95, 96, 99].includes(Number(value)));
  if (maxWind >= 75) return { level: "orange", label: buildAlertLabel("orange", "vento"), phenomenon: "vento" };
  if (snowSum >= 20) return { level: "orange", label: buildAlertLabel("orange", "neve"), phenomenon: "neve" };
  if (hasStormCode) return { level: "yellow", label: buildAlertLabel("yellow", "Temporali"), phenomenon: "Temporali" };
  if (minTemp <= -2) return { level: "yellow", label: buildAlertLabel("yellow", "ghiaccio"), phenomenon: "ghiaccio" };
  if (minVisibility < 500) return { level: "yellow", label: buildAlertLabel("yellow", "nebbia"), phenomenon: "nebbia" };
  if (maxTemp >= 38) return { level: "yellow", label: buildAlertLabel("yellow", "caldo estremo"), phenomenon: "caldo estremo" };
  return { level: "green", label: "Nessuna allerta", phenomenon: "" };
}

function pickHighestAlert(alerts) {
  return alerts.reduce((best, alert) => {
    const level = alert?.level || "green";
    return ALERT_LEVEL_META[level].rank > ALERT_LEVEL_META[best.level].rank ? alert : best;
  }, { level: "green", label: "Nessuna allerta", phenomenon: "" });
}

function buildAlertLabel(level, phenomenon) {
  if (level === "green") return "Nessuna allerta";
  return phenomenon && phenomenon !== "Protezione Civile" ? `Allerta ${phenomenon}` : "Allerta Protezione Civile";
}

function buildCivilProtectionAlertChip(alert) {
  const level = alert?.level || "green";
  const meta = ALERT_LEVEL_META[level] || ALERT_LEVEL_META.green;
  const text = alert?.loading ? "Verifica Protezione Civile..." : `${meta.emoji} ${alert?.label || meta.label}`;
  return `<span class='weather-risk-chip ${meta.className}' title='Avviso Protezione Civile${alert?.region ? ` • ${escapeHTML(alert.region)}` : ""}'>${escapeHTML(text)}</span>`;
}

function renderCivilProtectionAlert(alert) {
  currentCivilProtectionAlert = { ...currentCivilProtectionAlert, ...(alert || {}) };
  const level = currentCivilProtectionAlert.level || "green";
  const showBanner = ALERT_LEVEL_META[level]?.rank >= ALERT_LEVEL_META.orange.rank;
  ui.weatherAlertBanner?.classList.toggle("hidden", !showBanner);
}

function openWeatherExternalDetail() {
  const target = currentWeatherTarget || getWeatherTargetCoordinates();
  const hasCivilProtectionAlert = ALERT_LEVEL_META[currentCivilProtectionAlert?.level || "green"].rank > 0;
  const url = hasCivilProtectionAlert
    ? (currentCivilProtectionAlert.url || CIVIL_PROTECTION_ALERT_PAGE)
    : `${METEO_3B_BASE_URL}?lat=${encodeURIComponent(target.lat)}&lon=${encodeURIComponent(target.lon)}`;
  window.open(url, "_blank", "noopener,noreferrer");
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
  chatNotificationsInitialized = false;
  unsubscribeChat = db
    .collection("chatMessages")
    .orderBy("createdAt", "desc")
    .limit(500)
    .onSnapshot((snapshot) => {
      chatMessages = snapshot.docs
        .map((doc) => ({ id: doc.id, ...doc.data() }))
        .reverse();
      notifyForNewChatMessages(snapshot.docChanges());
      renderChat(chatMessages);
    }, (error) => {
      console.error(error);
      ui.chatFeedback.textContent = "Errore caricamento chat.";
    });
}

async function notifyForNewChatMessages(changes = []) {
  if (!chatNotificationsInitialized) {
    chatNotificationsInitialized = true;
    return;
  }
  const addedMessages = changes
    .filter((change) => change.type === "added")
    .map((change) => ({ id: change.doc.id, ...change.doc.data() }))
    .filter((message) => canNotifyForChatMessage(message));

  for (const message of addedMessages) {
    const senderName = String(message.senderName || "Operatore").trim();
    const body = getChatNotificationBody(message);
    try {
      await showLocalNotification(`Nuovo messaggio da ${senderName}`, {
        body,
        tag: `hera-chat-${message.id}`,
        data: { url: "./index.html#chat" }
      });
    } catch (error) {
      console.warn("Invio notifica chat non riuscito:", error);
    }
  }
}

function canNotifyForChatMessage(message) {
  if (!currentUser || isOwnMessage(message)) return false;
  if (!canViewMessage(message) || !isChatMessageFresh(message)) return false;
  if (!document.hidden && ui.chatModal && !ui.chatModal.classList.contains("hidden")) return false;
  return true;
}

function getChatNotificationBody(message) {
  const text = String(message.text || message.message || message.body || message.content || "").trim();
  if (text) return text.length > 180 ? `${text.slice(0, 177)}...` : text;
  const type = String(message.type || "").toLowerCase();
  if (type === "image") return "Ha inviato una foto.";
  if (type === "video") return "Ha inviato un video.";
  if (type === "voice") return "Ha inviato un messaggio vocale.";
  return "Hai un nuovo messaggio in chat.";
}

function isChatMessageFresh(message) {
  const expiresAtMs = getChatMessageExpiryMs(message);
  if (!expiresAtMs) return true;
  return expiresAtMs >= Date.now();
}

function getChatMessageCreatedAtMs(message) {
  if (message?.createdAt && typeof message.createdAt.toDate === "function") {
    return message.createdAt.toDate().getTime();
  }
  return 0;
}

function getChatMessageExpiryMs(message) {
  if (message?.expiresAt && typeof message.expiresAt.toDate === "function") {
    return message.expiresAt.toDate().getTime();
  }
  const createdAtMs = getChatMessageCreatedAtMs(message);
  if (!createdAtMs) return 0;
  return createdAtMs + CHAT_RETENTION_MS;
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
  const nowDate = new Date();
  try {
    const [legacySnapshot, expiresSnapshot] = await Promise.all([
      db
        .collection("chatMessages")
        .where("createdAt", "<=", firebase.firestore.Timestamp.fromDate(cutoffDate))
        .limit(200)
        .get(),
      db
        .collection("chatMessages")
        .where("expiresAt", "<=", firebase.firestore.Timestamp.fromDate(nowDate))
        .limit(200)
        .get()
    ]);
    const docsToDelete = new Map();
    legacySnapshot.docs.forEach((doc) => {
      const data = doc.data() || {};
      if (data.expiresAt && typeof data.expiresAt.toDate === "function") return;
      docsToDelete.set(doc.id, doc);
    });
    expiresSnapshot.docs.forEach((doc) => docsToDelete.set(doc.id, doc));
    if (!docsToDelete.size) return;
    const batch = db.batch();
    docsToDelete.forEach((doc) => batch.delete(doc.ref));
    await batch.commit();
  } catch (error) {
    console.warn("Pulizia chat 24h non completata:", error);
  }
}

function startHoursDeadlineAlertLoop() {
  stopHoursDeadlineAlertLoop();
  checkAndSendHoursDeadlineAlerts();
  hoursDeadlineAlertTimer = setInterval(() => {
    checkAndSendHoursDeadlineAlerts();
  }, 15 * 60 * 1000);
}

function stopHoursDeadlineAlertLoop() {
  if (hoursDeadlineAlertTimer) {
    clearInterval(hoursDeadlineAlertTimer);
    hoursDeadlineAlertTimer = null;
  }
}

function hasSquadreRowsForData(squadData) {
  const rows = Array.isArray(squadData?.squadre) ? squadData.squadre : getLegacySquadreRows(squadData || {});
  return rows.some((row) => String(row?.personale || "").trim() || String(row?.mezzi || "").trim());
}

function hasHoursForCommessaInEntries(entries, commessaId) {
  if (!Array.isArray(entries) || !commessaId) return false;
  return entries.some((entry) => {
    if (String(entry?.commessaId || "") !== String(commessaId)) return false;
    return Array.isArray(entry.rows) && entry.rows.some((row) => Number(row?.ore || 0) > 0);
  });
}

async function checkAndSendHoursDeadlineAlerts() {
  if (!currentUser || !canManageData()) return;
  const now = new Date();
  if (now.getHours() < HOURS_DEADLINE_ALERT_HOUR) return;
  const dateKey = getDateKeyFromLocalDate(now);
  const storicoDelGiorno = squadreHistoryByDate.get(dateKey) || new Map();
  const commesseConSquadra = Array.from(storicoDelGiorno.values())
    .filter((squadData) => hasSquadreRowsForData(squadData))
    .map((squadData) => ({
      commessaId: String(squadData.commessaId || "").trim(),
      commessaName: String(squadData.commessaNome || "").trim() || (commesseById.get(String(squadData.commessaId || "").trim()) || {}).nome || "Commessa"
    }))
    .filter((row) => row.commessaId);
  if (!commesseConSquadra.length) return;

  const [reportsSnapshot, approvalSnapshot] = await Promise.all([
    db.collection("oreReports").where("date", "==", dateKey).get(),
    db.collection("oreApprovalRequests").where("date", "==", dateKey).get()
  ]);

  const reports = reportsSnapshot.docs.map((doc) => ({ id: doc.id, ...doc.data() }));
  const approvals = approvalSnapshot.docs
    .map((doc) => ({ id: doc.id, ...doc.data() }))
    .filter((request) => String(request.status || "").trim() !== "rejected");

  for (const commessa of commesseConSquadra) {
    const hasHoursSaved = reports.some((report) => hasHoursForCommessaInEntries(report.entries, commessa.commessaId));
    const hasHoursPending = approvals.some((request) => hasHoursForCommessaInEntries(request.entries, commessa.commessaId));
    if (hasHoursSaved || hasHoursPending) continue;
    await sendHoursDeadlineAlertIfMissing({
      commessaId: commessa.commessaId,
      commessaName: commessa.commessaName,
      dateKey
    });
  }
}

async function sendHoursDeadlineAlertIfMissing({ commessaId, commessaName, dateKey }) {
  if (!commessaId || !dateKey) return;
  const alertId = `${dateKey}__${commessaId}__all`;
  const alertRef = db.collection("hoursDeadlineAlerts").doc(alertId);
  const existing = await alertRef.get();
  if (existing.exists) return;

  const dateLabel = new Date(`${dateKey}T00:00:00`).toLocaleDateString("it-IT");
  const text = `⚠️ Avviso ore mancanti: per la commessa ${commessaName || "Commessa"} (${dateLabel}) non risultano ore inserite entro le 19:00.`;
  const expiresAt = new Date(Date.now() + HOURS_DEADLINE_ALERT_RETENTION_MS);
  const chatDocRef = await db.collection("chatMessages").add({
    type: "text",
    text,
    recipientId: "",
    senderId: "system",
    senderName: "Sistema ore",
    senderEmail: "",
    kind: "system",
    metadata: {
      type: "hours_deadline_alert",
      action: "open_hours",
      commessaId,
      commessaName: commessaName || "",
      date: dateKey
    },
    expiresAt: firebase.firestore.Timestamp.fromDate(expiresAt),
    createdAt: firebase.firestore.FieldValue.serverTimestamp()
  });

  await db.collection("appNotifications").add({
    eventType: "hours-deadline-missing",
    title: "Ore mancanti entro le 19",
    body: text,
    commessaId,
    commessaName: commessaName || "",
    impiantoName: "",
    impiantoKey: "",
    createdByUid: "system",
    createdByName: "Sistema ore",
    createdByEmail: "",
    createdAt: firebase.firestore.FieldValue.serverTimestamp()
  });

  await alertRef.set({
    alertId,
    recipient: "all",
    commessaId,
    commessaName: commessaName || "",
    date: dateKey,
    chatMessageId: chatDocRef.id,
    createdAt: firebase.firestore.FieldValue.serverTimestamp()
  }, { merge: true });
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
    subscribePosDocuments();
    renderCommesseManagementList();
    renderAdminUsers();
    if (currentUser) {
      subscribeUsers();
      subscribeOperatorPositions();
      subscribeProgrammazioni();
    }
  }, (error) => {
    console.error("Errore caricamento admin users:", error);
    adminEmails = new Set([ADMIN_EMAIL]);
    updateAdminControls();
    subscribePosDocuments();
    renderCommesseManagementList();
    renderAdminUsers();
    if (currentUser) {
      subscribeUsers();
      subscribeOperatorPositions();
      subscribeProgrammazioni();
    }
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
  stopUsersSubscription();
  if (!currentUser) return;
  const source = canManageData()
    ? db.collection("platformUsers")
    : db.collection("platformUsers").where(firebase.firestore.FieldPath.documentId(), "==", currentUser.uid);
  unsubscribeUsers = source.onSnapshot((snapshot) => {
    platformUsers = snapshot.docs.map((doc) => ({ id: doc.id, ...doc.data() }))
      .sort((a, b) => String(a.displayName || "").localeCompare(String(b.displayName || ""), "it"));
    syncNotificationAutoPreferenceFromProfile();
    maybeAutoEnableNotifications();
    deniedImpiantoActions = getDeniedActionsForCurrentUser();
    renderChatRecipients();
    renderUserPermissionList();
    renderNotificationTargetUsers();
    renderHeaderActivitySummary();
    renderExternalApps();
    renderMap();
    checkAndSendHoursDeadlineAlerts();
  }, (error) => {
    console.error("Errore caricamento utenti:", error);
    platformUsers = [];
    renderChatRecipients();
    renderHeaderActivitySummary();
  });
}

function subscribeProgrammazioni() {
  if (unsubscribeProgrammazioni) unsubscribeProgrammazioni();
  unsubscribeProgrammazioni = db.collection("programmazioni").onSnapshot((snapshot) => {
    programmazioni = snapshot.docs.map((doc) => ({ id: doc.id, ...doc.data() }));
    renderProgrammazioni();
  }, () => {
    programmazioni = [];
    renderProgrammazioni();
  });
}

function populateProgrammazioneFormOptions() {
  if (ui.programmaCommessa) {
    const previous = String(ui.programmaCommessa.value || "");
    const commesse = sortCommesseByCreatedAtDesc(Array.from(commesseById.values()));
    ui.programmaCommessa.innerHTML = "<option value=''>Commessa</option>";
    commesse.forEach((commessa) => {
      const option = document.createElement("option");
      option.value = String(commessa.nome || "").trim();
      option.textContent = String(commessa.nome || "Commessa");
      ui.programmaCommessa.appendChild(option);
    });
    if (previous) ui.programmaCommessa.value = previous;
  }
  if (ui.programmaOperatori) {
    const selected = new Set(Array.from(ui.programmaOperatori.selectedOptions || []).map((opt) => String(opt.value || "").trim()));
    ui.programmaOperatori.innerHTML = "";
    personaleRecords
      .map((person) => getPersonaleDisplayName(person))
      .filter(Boolean)
      .sort((a, b) => String(a).localeCompare(String(b), "it"))
      .forEach((operatorName) => {
        const value = String(operatorName || "").trim();
        if (!value) return;
        const option = document.createElement("option");
        option.value = value;
        option.textContent = value;
        if (selected.has(value)) option.selected = true;
        ui.programmaOperatori.appendChild(option);
      });
  }
  if (ui.programmaMezzi) {
    const selected = new Set(Array.from(ui.programmaMezzi.selectedOptions || []).map((opt) => String(opt.value || "")));
    ui.programmaMezzi.innerHTML = "";
    mezziRecords.forEach((mezzo) => {
      const label = String(mezzo.nId || mezzo.nome || mezzo.modello || "").trim();
      if (!label) return;
      const option = document.createElement("option");
      option.value = label;
      option.textContent = label;
      if (selected.has(label)) option.selected = true;
      ui.programmaMezzi.appendChild(option);
    });
  }
}

function getProgrammazioneOperatorOptions() {
  return personaleRecords
    .map((person) => ({ value: getPersonaleDisplayName(person), avatar: String(person.avatarUrl || "").trim() }))
    .filter((item) => item.value)
    .sort((a, b) => String(a.value).localeCompare(String(b.value), "it"));
}

function getProgrammazioneMezziOptions() {
  return mezziRecords
    .map((mezzo) => String(mezzo.nId || mezzo.nome || mezzo.modello || "").trim())
    .filter(Boolean)
    .sort((a, b) => a.localeCompare(b, "it"))
    .map((value) => ({ value, avatar: "" }));
}

function buildProgrammazioneAutocomplete(root, label, options, selectedValues = []) {
  if (!root) return { getValues: () => [] };
  const selected = new Map();
  selectedValues.forEach((v) => selected.set(String(v).toLowerCase(), { value: v, avatar: "" }));
  root.innerHTML = `<label>${label}</label><input type="text" placeholder="Cerca..." autocomplete="off"><div class="autocomplete-list hidden"></div><div class="autocomplete-chips"></div>`;
  const input = root.querySelector("input");
  const list = root.querySelector(".autocomplete-list");
  const chips = root.querySelector(".autocomplete-chips");
  function renderChips() {
    chips.innerHTML = Array.from(selected.values()).map((item) => `<span class="autocomplete-chip">${item.avatar ? `<img src="${escapeHTML(item.avatar)}" alt="" width="18" height="18">` : ""}${escapeHTML(item.value)}<button type="button" data-remove="${escapeHTML(item.value)}">✕</button></span>`).join("");
  }
  function renderList() {
    const q = String(input.value || "").trim().toLowerCase();
    const filtered = options.filter((item) => item.value.toLowerCase().includes(q) && !selected.has(item.value.toLowerCase())).slice(0, 8);
    list.innerHTML = filtered.map((item) => `<div class="autocomplete-item" data-value="${escapeHTML(item.value)}">${escapeHTML(item.value)}</div>`).join("") || "<div class='autocomplete-item muted'>Nessun risultato</div>";
    list.classList.remove("hidden");
  }
  input.addEventListener("input", renderList);
  input.addEventListener("focus", renderList);
  document.addEventListener("click", (event) => { if (!root.contains(event.target)) list.classList.add("hidden"); });
  list.addEventListener("click", (event) => {
    const row = event.target.closest("[data-value]");
    if (!row) return;
    const value = String(row.getAttribute("data-value") || "");
    const found = options.find((item) => item.value === value) || { value, avatar: "" };
    selected.set(value.toLowerCase(), found);
    input.value = "";
    renderChips();
    renderList();
  });
  chips.addEventListener("click", (event) => {
    const btn = event.target.closest("[data-remove]");
    if (!btn) return;
    selected.delete(String(btn.getAttribute("data-remove") || "").toLowerCase());
    renderChips();
  });
  renderChips();
  return { getValues: () => Array.from(selected.values()).map((item) => item.value) };
}

function subscribeOperatorPositions() {
  stopOperatorPositionsSubscription();
  if (!currentUser || !canManageData()) {
    operatorPositions = [];
    renderMap();
    return;
  }
  unsubscribeOperatorPositions = db.collection("operatorPositions").onSnapshot((snapshot) => {
    operatorPositions = snapshot.docs.map((doc) => ({ id: doc.id, ...doc.data() }));
    renderMap();
  }, (error) => {
    console.error("Errore caricamento posizioni squadre:", error);
    operatorPositions = [];
    renderMap();
  });
}

function stopOperatorPositionsSubscription() {
  if (unsubscribeOperatorPositions) {
    unsubscribeOperatorPositions();
    unsubscribeOperatorPositions = null;
  }
  operatorPositions = [];
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
  renderNotificationTargetUsers();
  renderHeaderActivitySummary();
  renderExternalApps();
}

function canApproveHoursLevel1(request) {
  if (!currentUser || !request) return false;
  if (canManageData()) return true;
  return String(request.createdByUid || "") === String(currentUser.uid || "");
}

function hasValidHoursRows(record = {}) {
  return (Array.isArray(record.entries) ? record.entries : []).some((entry) => (
    String(entry?.commessaId || "").trim()
    && (Array.isArray(entry?.rows) ? entry.rows : []).some((row) => (
      String(row?.operatore || "").trim() && Number(row?.ore || 0) > 0
    ))
  ));
}

function isBlockingHoursApprovalWithoutVisibleRecord(request = {}) {
  const status = String(request.status || "").trim();
  if (status === "rejected" || status === "approved") return false;
  return !hasValidHoursRows(request);
}

async function unblockInvalidHoursRequest(request) {
  if (!canManageData()) {
    alert("Solo l'admin può sbloccare una richiesta ore incompleta.");
    return;
  }
  if (!request?.id) return;
  const ok = window.confirm(`Sbloccare la richiesta ore ${request.id}? Verrà marcata come rifiutata perché non contiene un record ore completo.`);
  if (!ok) return;
  await db.collection("oreApprovalRequests").doc(request.id).set({
    status: "rejected",
    rejectedBy: currentUser.email || currentUser.displayName || "admin",
    rejectedAt: firebase.firestore.FieldValue.serverTimestamp(),
    rejectionReason: "Sblocco admin: blocco ore senza record ore completo visibile.",
    updatedAt: firebase.firestore.FieldValue.serverTimestamp()
  }, { merge: true });
  await updateHoursLocksForEntries(request.date || "", request.entries || [], {
    status: "rejected",
    approvalRequestId: request.id || "",
    rejectedBy: currentUser.email || currentUser.displayName || "admin"
  });
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
    const hasInvalidBlock = isBlockingHoursApprovalWithoutVisibleRecord(request);
    card.innerHTML = `
      <p><b>ID:</b> ${escapeHTML(request.id || "-")}</p>
      <p><b>Data:</b> ${escapeHTML(dateLabel)} • <b>Creato da:</b> ${escapeHTML(author)}</p>
      <p><b>Stato:</b> ${escapeHTML(statusText)}</p>
      <ul>${summary || "<li>Nessuna commessa</li>"}</ul>
      ${hasInvalidBlock ? `<p class="warning"><b>Errore:</b> risulta un blocco ore ma il record ore non è stato trovato. Contattare amministratore.</p>` : ""}
      ${request.rejectionReason ? `<p><b>Motivo rifiuto:</b> ${escapeHTML(request.rejectionReason)}</p>` : ""}
    `;
    const actions = document.createElement("div");
    actions.className = "item-actions";
    if (hasInvalidBlock && canManageData()) {
      actions.appendChild(createButton("Sblocca blocco ore", () => unblockInvalidHoursRequest(request)));
    }
    if (!hasInvalidBlock && request.status === "pending_level1" && canApproveHoursLevel1(request)) {
      actions.appendChild(createButton("Accetta livello 1", () => approveHoursRequestLevel1(request)));
      actions.appendChild(createButton("Rifiuta", () => rejectHoursRequest(request)));
    }
    if (!hasInvalidBlock && request.status === "pending_admin" && canManageData()) {
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
  await updateHoursLocksForEntries(request.date || "", request.entries || [], {
    status: "pending_admin",
    approvalRequestId: request.id || ""
  });
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
  try {
    const result = await saveApprovedHoursRequest(request);
    const targetUser = platformUsers.find((user) => String(user.uid || "") === String(request.createdByUid || ""));
    if (targetUser?.id) {
      await sendPrivateChatNotification({
        recipientId: targetUser.id,
        text: `✅ Richiesta ore ${request.id} approvata definitivamente. Report salvato: ${result.reportId}.`,
        senderName: currentUser.displayName || currentUser.email || "Admin"
      });
    }
    loadSavedHoursReports();
  } catch (error) {
    if (error?.code === "hours-duplicate-lock" || error?.code === "hours-duplicate-draft") {
      alert(error.message || formatHoursDuplicateMessage(error.conflicts, { admin: true }));
      return;
    }
    alert(error?.message || "Errore durante la conferma ore.");
    throw error;
  }
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
  await updateHoursLocksForEntries(request.date || "", request.entries || [], {
    status: "rejected",
    approvalRequestId: request.id || "",
    rejectedBy: currentUser.email || currentUser.displayName || "admin"
  });
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
  await updateHoursLocksForEntries(request.date || "", request.entries || [], {
    status: "rejected",
    approvalRequestId: request.id || "",
    rejectedBy: currentUser.email || currentUser.displayName || "utente"
  });
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
  if (action === "done" || action === "whatsapp") return false;
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


function getPlatformUserLabel(user) {
  if (!user) return "Utente sconosciuto";
  return String(user.displayName || user.email || user.id || "Utente").trim() || "Utente";
}

function renderNotificationTargetUsers() {
  if (!ui.notificationUserSelect) return;
  const previous = Array.from(ui.notificationUserSelect.selectedOptions || []).map((opt) => opt.value);
  ui.notificationUserSelect.innerHTML = "";
  const recipients = platformUsers.filter((user) => !adminEmails.has(normalizeEmail(user.email)));
  recipients.forEach((user) => {
    const opt = document.createElement("option");
    opt.value = user.id;
    opt.textContent = `${getPlatformUserLabel(user)}${user.email ? ` (${user.email})` : ""}`;
    if (previous.includes(user.id)) opt.selected = true;
    ui.notificationUserSelect.appendChild(opt);
  });
  onNotificationSendAllChange();
}

function onNotificationSendAllChange() {
  const sendAll = Boolean(ui.notificationSendAllToggle?.checked);
  if (ui.notificationUserSelect) ui.notificationUserSelect.disabled = sendAll;
}

function renderNotificationsList() {
  if (!ui.notificationsList) return;
  if (!currentUser) {
    ui.notificationsList.innerHTML = "<p class='muted'>Fai login per vedere le notifiche.</p>";
    return;
  }
  if (!canManageData()) {
    ui.notificationsList.innerHTML = "<p class='muted'>Solo gli admin possono gestire le notifiche.</p>";
    return;
  }
  if (!userAlerts.length) {
    ui.notificationsList.innerHTML = "<p class='muted'>Nessuna notifica disponibile.</p>";
    return;
  }
  ui.notificationsList.innerHTML = "";
  userAlerts.forEach((item) => {
    const userLabel = formatNotificationRecipientsLabel(item);
    const row = document.createElement("article");
    row.className = "simple-list-item stacked";
    const status = item.scheduledDateKey && item.scheduledDateKey > getDateKeyFromLocalDate(new Date()) ? "Programm. futura" : "Attiva";
    const title = String(item.title || "Notifica").trim();
    const attachments = Array.isArray(item.attachments) ? item.attachments : [];
    row.innerHTML = `<strong>${escapeHTML(title)}</strong><p>${escapeHTML(item.message || "")}</p><small>${escapeHTML(userLabel)} • ${escapeHTML(status)}</small>`;
    const attachmentsBox = document.createElement("div");
    if (!attachments.length) {
      attachmentsBox.innerHTML = "<small>Nessun allegato.</small>";
    } else {
      const list = document.createElement("ul");
      attachments.forEach((att) => {
        const li = document.createElement("li");
        const link = document.createElement("a");
        link.href = "#";
        link.textContent = `📎 ${att.name || "Documento"}`;
        link.addEventListener("click", (event) => {
          event.preventDefault();
          openNotificationDocumentViewer(att.url || "", att.name || "Documento");
        });
        li.appendChild(link);
        list.appendChild(li);
      });
      attachmentsBox.appendChild(list);
    }
    row.appendChild(attachmentsBox);
    const actions = document.createElement("div");
    actions.className = "item-actions";
    actions.appendChild(createButton("Elimina", () => deleteUserNotification(item.id)));
    actions.appendChild(createButton("Dettaglio giorno", () => openNotificationDayDetail(getNotificationPrimaryDateKey(item))));
    row.appendChild(actions);
    ui.notificationsList.appendChild(row);
  });
}

function getNotificationRecipientUserIds(alertItem) {
  if (!alertItem) return [];
  if (Array.isArray(alertItem.targetUserIds)) return alertItem.targetUserIds.map((id) => String(id || "").trim()).filter(Boolean);
  if (alertItem.targetUserId) return [String(alertItem.targetUserId).trim()].filter(Boolean);
  return [];
}

function formatNotificationRecipientsLabel(alertItem) {
  if (alertItem?.sendToAllRegistered) return "Destinatari: tutti gli utenti registrati";
  const ids = getNotificationRecipientUserIds(alertItem);
  if (!ids.length) return "Destinatari non impostati";
  const labels = ids.map((id) => getPlatformUserLabel(platformUsers.find((u) => u.id === id)));
  return `Destinatari: ${labels.join(", ")}`;
}

function getNotificationPrimaryDateKey(item) {
  return String(item?.scheduledDateKey || item?.createdDateKey || "").trim();
}

function renderActiveUserAlertAttachments() {
  if (!ui.userAlertAttachments) return;
  const attachments = Array.isArray(activeUserAlert?.attachments) ? activeUserAlert.attachments : [];
  if (!attachments.length) {
    ui.userAlertAttachments.innerHTML = "<p class='muted'>Nessun documento allegato.</p>";
    return;
  }
  ui.userAlertAttachments.innerHTML = "";
  attachments.forEach((attachment) => {
    const row = document.createElement("div");
    row.className = "simple-list-item";
    const label = document.createElement("span");
    label.textContent = `📎 ${attachment.name || "Documento"}`;
    row.appendChild(label);
    const openBtn = createButton("Apri", () => openNotificationDocumentViewer(attachment.url || "", attachment.name || "Documento"));
    row.appendChild(openBtn);
    ui.userAlertAttachments.appendChild(row);
  });
}

function buildNotificationEmailPayload(targetUser, message, notificationId) {
  const recipientEmail = String(targetUser?.email || "").trim();
  const appUrl = `${window.location.origin}${window.location.pathname}#home`;
  return {
    to: recipientEmail,
    subject: "Nuova notifica da Hera App",
    text: `Ciao ${targetUser?.displayName || "utente"},\n\nl'amministratore ti ha inviato una nuova notifica:\n\"${message}\"\n\nApri l'app per visualizzarla: ${appUrl}\n\nID notifica: ${notificationId}`,
    html: `
      <p>Ciao ${escapeHTML(targetUser?.displayName || "utente")},</p>
      <p>L'amministratore ti ha inviato una nuova notifica su Hera App.</p>
      <p><b>Messaggio:</b> ${escapeHTML(message)}</p>
      <p><a href="${appUrl}" style="display:inline-block;padding:10px 14px;background:#1d4ed8;color:#fff;text-decoration:none;border-radius:8px;">Clicca per aprire l'app</a></p>
      <p style="font-size:12px;color:#64748b;">ID notifica: ${escapeHTML(notificationId)}</p>
    `,
    appUrl,
    notificationId
  };
}

async function queueNotificationEmail(targetUser, message, notificationId) {
  if (!targetUser?.email) return { queued: false, reason: "missing_email" };
  const payload = buildNotificationEmailPayload(targetUser, message, notificationId);
  await db.collection("notificationEmailQueue").add({
    ...payload,
    status: "pending",
    createdAt: firebase.firestore.FieldValue.serverTimestamp(),
    createdBy: currentUser?.email || currentUser?.uid || "admin"
  });
  const webhookUrl = String(window?.HERA_NOTIFICATION_EMAIL_WEBHOOK || localStorage.getItem("heraNotificationEmailWebhook") || "").trim();
  if (!webhookUrl) return { queued: true, delivered: false };
  try {
    await fetchWithTimeoutAndRetry(webhookUrl, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(payload)
    }, {
      timeoutMs: NETWORK_DEFAULT_TIMEOUT_MS,
      retries: 1
    });
    return { queued: true, delivered: true };
  } catch (error) {
    console.warn("Webhook email notifica non raggiungibile:", error);
    return { queued: true, delivered: false, error };
  }
}

function mapFileNameForNotificationAttachment(fileName = "") {
  const safeName = String(fileName || "allegato").replace(/\s+/g, "_").replace(/[^\w.\-]+/g, "_");
  return `notifiche/${Date.now()}_${safeName}`;
}

function setNotificationUploadState(inProgress) {
  notificationUploadInProgress = Boolean(inProgress);
  if (ui.notificationCancelUploadBtn) ui.notificationCancelUploadBtn.disabled = !notificationUploadInProgress || !canManageData();
  if (ui.notificationSubmit) ui.notificationSubmit.disabled = notificationUploadInProgress || !canManageData();
}

function cancelNotificationUpload() {
  if (!notificationUploadAbortController) return;
  notificationUploadAbortController.abort();
  notificationUploadAbortController = null;
  setNotificationUploadState(false);
  if (ui.notificationFeedback) ui.notificationFeedback.textContent = "Caricamento allegati annullato.";
}

async function uploadNotificationAttachments(files = [], options = {}) {
  if (!files.length) return [];
  if (!isCentralDriveConfigured() && !driveAccessToken) {
    throw new Error("Cloud amministratore non configurato. Configura Drive admin per allegare documenti alle notifiche.");
  }
  if (!driveReportsFolderId) await ensureDriveFolders();
  const { signal = null, onProgress = null } = options;
  let completed = 0;
  const uploads = await Promise.all(files.map(async (file) => {
    if (signal?.aborted) throw new DOMException("Upload annullato", "AbortError");
    const upload = await uploadBlobToDrive(
      file,
      mapFileNameForNotificationAttachment(file.name || "allegato"),
      file.type || "application/octet-stream",
      driveReportsFolderId,
      { signal }
    );
    completed += 1;
    if (typeof onProgress === "function") onProgress(completed, files.length, file.name || "allegato");
    return {
      name: file.name || "allegato",
      type: file.type || "application/octet-stream",
      size: Number(file.size || 0),
      url: upload.webViewLink || upload.directUrl || "",
      fileId: upload.fileId || ""
    };
  }));
  return uploads;
}

function buildDocumentViewerUrl(rawUrl = "") {
  const url = String(rawUrl || "").trim();
  if (!url) return "";
  if (/docs\.google\.com\/spreadsheets/i.test(url)) return url;
  if (/drive\.google\.com/i.test(url)) {
    return `https://docs.google.com/gview?embedded=1&url=${encodeURIComponent(url)}`;
  }
  return `https://docs.google.com/gview?embedded=1&url=${encodeURIComponent(url)}`;
}

function openNotificationDocumentViewer(rawUrl, title = "Documento") {
  const viewerUrl = buildDocumentViewerUrl(rawUrl);
  if (!viewerUrl) return;
  if (ui.notificationDocViewerTitle) ui.notificationDocViewerTitle.textContent = title;
  if (ui.notificationDocViewerFrame) ui.notificationDocViewerFrame.src = viewerUrl;
  ui.notificationDocViewerModal?.classList.remove("hidden");
  ui.notificationDocViewerModal?.setAttribute("aria-hidden", "false");
}

function closeNotificationDocumentViewer() {
  if (ui.notificationDocViewerFrame) ui.notificationDocViewerFrame.src = "";
  ui.notificationDocViewerModal?.classList.add("hidden");
  ui.notificationDocViewerModal?.setAttribute("aria-hidden", "true");
}

async function createUserNotification(event) {
  event.preventDefault();
  if (notificationUploadInProgress) {
    if (ui.notificationFeedback) ui.notificationFeedback.textContent = "Caricamento già in corso...";
    return;
  }
  if (!currentUser || !canManageData()) {
    if (ui.notificationFeedback) ui.notificationFeedback.textContent = "Solo gli admin possono inviare avvisi.";
    return;
  }
  const sendToAllRegistered = Boolean(ui.notificationSendAllToggle?.checked);
  const targetUserIds = sendToAllRegistered
    ? []
    : Array.from(ui.notificationUserSelect?.selectedOptions || []).map((opt) => String(opt.value || "").trim()).filter(Boolean);
  const title = String(ui.notificationTitle?.value || "").trim();
  const message = String(ui.notificationMessage?.value || "").trim();
  const scheduledDateKey = String(ui.notificationDate?.value || "").trim();
  const files = Array.from(ui.notificationAttachments?.files || []);
  const totalSizeMb = files.reduce((sum, file) => sum + Number(file.size || 0), 0) / (1024 * 1024);
  if (totalSizeMb > 80) {
    if (ui.notificationFeedback) ui.notificationFeedback.textContent = "Allegati troppo grandi (>80MB). Riduci il peso per velocizzare il caricamento.";
    return;
  }
  if (!title || !message) {
    if (ui.notificationFeedback) ui.notificationFeedback.textContent = "Inserisci titolo e testo notifica.";
    return;
  }
  if (!sendToAllRegistered && !targetUserIds.length) {
    if (ui.notificationFeedback) ui.notificationFeedback.textContent = "Seleziona almeno un utente o scegli invio a tutti.";
    return;
  }
  notificationUploadAbortController = new AbortController();
  setNotificationUploadState(true);
  try {
    if (ui.notificationFeedback) ui.notificationFeedback.textContent = files.length ? `Caricamento allegati (0/${files.length})...` : "Invio notifica...";
    const attachments = await uploadNotificationAttachments(files, {
      signal: notificationUploadAbortController.signal,
      onProgress: (completed, total) => {
        if (ui.notificationFeedback) ui.notificationFeedback.textContent = `Caricamento allegati (${completed}/${total})...`;
      }
    });
    if (ui.notificationFeedback) ui.notificationFeedback.textContent = "Salvataggio notifica...";
    const createdDateKey = getDateKeyFromLocalDate(new Date());
    const notificationRef = await db.collection("userAlerts").add({
      title,
      sendToAllRegistered,
      targetUserIds,
      message,
      attachments,
      scheduledDateKey: scheduledDateKey || "",
      createdDateKey,
      status: scheduledDateKey && scheduledDateKey > createdDateKey ? "scheduled" : "active",
      createdAt: firebase.firestore.FieldValue.serverTimestamp(),
      createdBy: currentUser.email || currentUser.uid || "admin",
      acknowledgedUsers: 0
    });
    const targetUsers = sendToAllRegistered
      ? platformUsers.filter((user) => !adminEmails.has(normalizeEmail(user.email)))
      : platformUsers.filter((user) => targetUserIds.includes(user.id));
    const emailResults = await Promise.all(targetUsers.map((targetUser) => queueNotificationEmail(targetUser, message, notificationRef.id)));
    const hasQueued = emailResults.some((result) => result?.queued);
    if (ui.notificationForm) ui.notificationForm.reset();
    onNotificationSendAllChange();
    if (ui.notificationFeedback) {
      ui.notificationFeedback.textContent = hasQueued
        ? "Notifica salvata. Email messe in coda per i destinatari con email."
        : "Notifica salvata. Nessuna email inviata (destinatari senza email).";
    }
  } catch (error) {
    if (error?.name === "AbortError") {
      if (ui.notificationFeedback) ui.notificationFeedback.textContent = "Caricamento annullato.";
      return;
    }
    console.error("Errore invio avviso utente:", error);
    if (ui.notificationFeedback) ui.notificationFeedback.textContent = "Errore durante il salvataggio dell'avviso.";
  } finally {
    notificationUploadAbortController = null;
    setNotificationUploadState(false);
  }
}

async function deleteUserNotification(notificationId) {
  if (!currentUser || !canManageData() || !notificationId) return;
  const confirmed = window.confirm("Eliminare questa notifica?");
  if (!confirmed) return;
  await db.collection("userAlerts").doc(notificationId).delete();
}

function subscribeUserAlerts() {
  if (unsubscribeUserAlerts) unsubscribeUserAlerts();
  if (!currentUser) return;
  if (canManageData()) {
    unsubscribeUserAlerts = db.collection("userAlerts")
      .orderBy("createdAt", "desc")
      .limit(200)
      .onSnapshot((snapshot) => {
        userAlerts = snapshot.docs.map((doc) => ({ id: doc.id, ...doc.data() }));
        renderNotificationsList();
        renderNotificationCalendar();
      }, (error) => {
        console.error("Errore caricamento notifiche admin:", error);
      });
    return;
  }
  unsubscribeUserAlerts = db.collection("userAlerts")
    .limit(30)
    .onSnapshot((snapshot) => {
      const todayKey = getDateKeyFromLocalDate(new Date());
      userAlerts = snapshot.docs
        .map((doc) => ({ id: doc.id, ...doc.data() }))
        .filter((item) => {
          const userTargets = getNotificationRecipientUserIds(item);
          const belongsToUser = item.sendToAllRegistered || userTargets.includes(currentUser.uid);
          if (!belongsToUser) return false;
          const dateKey = String(item.scheduledDateKey || "").trim();
          return !dateKey || dateKey <= todayKey;
        })
        .sort((a, b) => {
          const aMs = a?.createdAt && typeof a.createdAt.toMillis === "function" ? a.createdAt.toMillis() : 0;
          const bMs = b?.createdAt && typeof b.createdAt.toMillis === "function" ? b.createdAt.toMillis() : 0;
          return bMs - aMs;
        });
      maybeShowUserAlert();
    }, (error) => {
      console.error("Errore caricamento notifiche utente:", error);
    });
}

function stopUserAlertsSubscription() {
  if (unsubscribeUserAlerts) {
    unsubscribeUserAlerts();
    unsubscribeUserAlerts = null;
  }
  userAlerts = [];
  activeUserAlert = null;
  renderNotificationsList();
}

function maybeShowUserAlert() {
  if (!currentUser || canManageData()) return;
  const firstPending = userAlerts.find((alertItem) => !Boolean(alertItem?.ackByUserIds?.[currentUser.uid]));
  if (!firstPending) {
    closeUserAlertModal();
    return;
  }
  closeSideMenu();
  closeManagementPanel();
  activeUserAlert = firstPending;
  if (ui.userAlertText) ui.userAlertText.textContent = `${firstPending.title ? `${firstPending.title}\n\n` : ""}${firstPending.message || ""}`;
  renderActiveUserAlertAttachments();
  ui.userAlertModal?.classList.remove("hidden");
  ui.userAlertModal?.setAttribute("aria-hidden", "false");
}

function closeUserAlertModal() {
  if (ui.userAlertAttachments) ui.userAlertAttachments.innerHTML = "";
  ui.userAlertModal?.classList.add("hidden");
  ui.userAlertModal?.setAttribute("aria-hidden", "true");
}

async function acknowledgeActiveUserAlert() {
  if (!currentUser || !activeUserAlert?.id) return;
  const acknowledgementId = `${activeUserAlert.id}__${currentUser.uid}`;
  const acknowledgementRef = db.collection("userAlertAcknowledgements").doc(acknowledgementId);
  const existing = await acknowledgementRef.get();
  if (existing.exists) {
    activeUserAlert = null;
    closeUserAlertModal();
    return;
  }
  const now = new Date();
  await acknowledgementRef.set({
    notificationId: activeUserAlert.id,
    userId: currentUser.uid || "",
    userName: currentUser.displayName || currentUser.email || "Utente",
    acknowledgedAt: firebase.firestore.FieldValue.serverTimestamp(),
    acknowledgedDateKey: getDateKeyFromLocalDate(now)
  }, { merge: true });
  await db.collection("userAlerts").doc(activeUserAlert.id).set({
    [`ackByUserIds.${currentUser.uid}`]: firebase.firestore.FieldValue.serverTimestamp(),
    lastAcknowledgedAt: firebase.firestore.FieldValue.serverTimestamp()
  }, { merge: true });
  await sendNotificationAckToAdmins(activeUserAlert, now);
  activeUserAlert = null;
  closeUserAlertModal();
}

async function sendNotificationAckToAdmins(alertItem, ackDate = new Date()) {
  const adminUsers = platformUsers.filter((user) => adminEmails.has(normalizeEmail(user.email)));
  if (!adminUsers.length) return;
  const whenLabel = ackDate.toLocaleString("it-IT");
  const text = `✅ NOTIFICA CONFERMATA\nL’utente ${currentUser?.displayName || currentUser?.email || "Utente"} ha premuto “OK, HO CAPITO”\nNotifica: ${alertItem?.title || "Notifica"}\nData/Ora: ${whenLabel}`;
  await Promise.all(adminUsers.map((adminUser) => db.collection("chatMessages").add({
    type: "text",
    text,
    recipientId: adminUser.id,
    senderId: currentUser?.uid || "system",
    senderName: "Sistema notifiche",
    senderEmail: currentUser?.email || "",
    kind: "system",
    metadata: {
      type: "notification_ack",
      notificationId: alertItem?.id || "",
      notificationTitle: alertItem?.title || "",
      acknowledgedByUserId: currentUser?.uid || "",
      acknowledgedByUserName: currentUser?.displayName || currentUser?.email || "Utente",
      acknowledgedAt: ackDate.toISOString()
    },
    createdAt: firebase.firestore.FieldValue.serverTimestamp()
  })));
}

function postponeActiveUserAlert() {
  activeUserAlert = null;
  closeUserAlertModal();
}

function openNotificationCalendarView() {
  if (!ui.notificationCalendarView || !ui.notificationMainView) return;
  ui.notificationMainView.classList.add("hidden");
  ui.notificationCalendarView.classList.remove("hidden");
  renderNotificationCalendar();
}

function closeNotificationCalendarView() {
  if (!ui.notificationCalendarView || !ui.notificationMainView) return;
  ui.notificationCalendarView.classList.add("hidden");
  ui.notificationMainView.classList.remove("hidden");
}

function moveNotificationCalendarMonth(offset) {
  notificationCalendarCursor = new Date(notificationCalendarCursor.getFullYear(), notificationCalendarCursor.getMonth() + offset, 1);
  renderNotificationCalendar();
}

function renderNotificationCalendar() {
  if (!ui.notificationCalendarGrid || !ui.notificationCalendarMonthLabel) return;
  const monthStart = new Date(notificationCalendarCursor.getFullYear(), notificationCalendarCursor.getMonth(), 1);
  const monthLabel = monthStart.toLocaleDateString("it-IT", { month: "long", year: "numeric" });
  ui.notificationCalendarMonthLabel.textContent = monthLabel.charAt(0).toUpperCase() + monthLabel.slice(1);
  const weekdayLabels = ["Lun", "Mar", "Mer", "Gio", "Ven", "Sab", "Dom"];
  const dayMap = new Map();
  userAlerts.forEach((item) => {
    const key = getNotificationPrimaryDateKey(item);
    if (!key) return;
    if (!dayMap.has(key)) dayMap.set(key, []);
    dayMap.get(key).push(item);
  });
  const firstWeekday = (monthStart.getDay() + 6) % 7;
  const daysInMonth = new Date(monthStart.getFullYear(), monthStart.getMonth() + 1, 0).getDate();
  ui.notificationCalendarGrid.innerHTML = "";
  weekdayLabels.forEach((label) => {
    const header = document.createElement("div");
    header.className = "notification-calendar-cell notification-calendar-weekday";
    header.textContent = label;
    ui.notificationCalendarGrid.appendChild(header);
  });
  for (let i = 0; i < firstWeekday; i += 1) {
    const empty = document.createElement("div");
    empty.className = "notification-calendar-cell is-empty";
    ui.notificationCalendarGrid.appendChild(empty);
  }
  for (let day = 1; day <= daysInMonth; day += 1) {
    const date = new Date(monthStart.getFullYear(), monthStart.getMonth(), day);
    const dateKey = getDateKeyFromLocalDate(date);
    const btn = document.createElement("button");
    btn.type = "button";
    btn.className = "notification-calendar-cell notification-calendar-day";
    btn.textContent = String(day);
    if (dayMap.has(dateKey)) btn.classList.add("has-notification");
    if (selectedNotificationCalendarDateKey === dateKey) btn.classList.add("is-selected");
    btn.addEventListener("click", () => openNotificationDayDetail(dateKey));
    ui.notificationCalendarGrid.appendChild(btn);
  }
}

async function openNotificationDayDetail(dateKey) {
  selectedNotificationCalendarDateKey = String(dateKey || "").trim();
  renderNotificationCalendar();
  if (!ui.notificationDayDetail) return;
  if (!selectedNotificationCalendarDateKey) {
    ui.notificationDayDetail.innerHTML = "";
    return;
  }
  const dateLabel = new Date(`${selectedNotificationCalendarDateKey}T00:00:00`).toLocaleDateString("it-IT");
  const dayItems = userAlerts.filter((item) => getNotificationPrimaryDateKey(item) === selectedNotificationCalendarDateKey);
  ui.notificationDayDetail.innerHTML = `<h4>Dettaglio ${escapeHTML(dateLabel)}</h4>`;
  const addButton = createButton("➕ Programma notifica per questo giorno", () => openScheduledNotificationForm(selectedNotificationCalendarDateKey));
  ui.notificationDayDetail.appendChild(addButton);
  if (!dayItems.length) {
    const empty = document.createElement("p");
    empty.className = "muted";
    empty.textContent = "Nessuna notifica registrata per questo giorno.";
    ui.notificationDayDetail.appendChild(empty);
    return;
  }
  const ackSnapshot = await db.collection("userAlertAcknowledgements").where("acknowledgedDateKey", "==", selectedNotificationCalendarDateKey).get();
  const acknowledgements = ackSnapshot.docs.map((doc) => ({ id: doc.id, ...doc.data() }));
  dayItems.forEach((item) => {
    const card = document.createElement("article");
    card.className = "simple-list-item stacked";
    const recipientsIds = item.sendToAllRegistered
      ? platformUsers.filter((user) => !adminEmails.has(normalizeEmail(user.email))).map((user) => user.id)
      : getNotificationRecipientUserIds(item);
    const recipientLabels = recipientsIds.map((id) => getPlatformUserLabel(platformUsers.find((u) => u.id === id)));
    const itemAcks = acknowledgements.filter((ack) => String(ack.notificationId || "") === String(item.id || ""));
    const ackedIds = new Set(itemAcks.map((ack) => String(ack.userId || "")));
    const pendingLabels = recipientsIds.filter((id) => !ackedIds.has(id)).map((id) => getPlatformUserLabel(platformUsers.find((u) => u.id === id)));
    const ackRows = itemAcks.map((ack) => {
      const when = ack.acknowledgedAt?.toDate ? ack.acknowledgedAt.toDate().toLocaleString("it-IT") : "-";
      return `<li>${escapeHTML(ack.userName || ack.userId || "Utente")} • ${escapeHTML(when)}</li>`;
    }).join("");
    card.innerHTML = `
      <strong>${escapeHTML(item.title || "Notifica")}</strong>
      <p>${escapeHTML(item.message || "")}</p>
      <small>${escapeHTML(`Destinatari: ${recipientLabels.join(", ") || "-"}`)}</small>
      <small>Confermati: ${itemAcks.length}</small>
      <small>Non confermati: ${pendingLabels.length ? escapeHTML(pendingLabels.join(", ")) : "Nessuno"}</small>
      <ul>${ackRows || "<li>Nessuna conferma</li>"}</ul>
    `;
    ui.notificationDayDetail.appendChild(card);
  });
}

function openScheduledNotificationForm(dateKey) {
  closeNotificationCalendarView();
  if (ui.notificationDate) ui.notificationDate.value = dateKey;
  ui.notificationTitle?.focus();
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
    row.className = "simple-list-item commessa-manage-item";
    const info = document.createElement("div");
    info.className = "commessa-manage-info";
    const title = document.createElement("strong");
    const codiceCommessa = String(commessa.codice || "").trim();
    const hasSubcommesse = getSubcommesse(commessa.id).length > 0;
    title.innerHTML = `${escapeHTML(commessa.nome || "Commessa senza nome")}${codiceCommessa ? ` • Cod. ${escapeHTML(codiceCommessa)}` : ""}${hasSubcommesse ? ` <span class="commessa-parent-indicator" title="Contiene subcommesse" aria-label="Contiene subcommesse">📂</span>` : ""}`;
    info.appendChild(title);

    const actions = document.createElement("div");
    actions.className = "item-actions";
    actions.appendChild(createButton("Modifica", () => renameCommessa(commessa.id, commessa.nome || "Commessa", commessa.codice || "")));
    actions.appendChild(createButton("Svuota", () => clearCommessaImpianti(commessa.id, commessa.nome || "Commessa")));
    actions.appendChild(createButton("Elimina", () => deleteCommessa(commessa.id, commessa.nome || "Commessa")));

    row.appendChild(info);
    row.appendChild(actions);
    ui.commesseManageList.appendChild(row);
  });
}

async function renameCommessa(commessaId, currentName, currentCode = "") {
  if (!canManageData()) {
    alert("Solo un admin può rinominare commesse.");
    return;
  }
  const nextName = window.prompt("Nuovo nome commessa:", currentName || "");
  if (nextName == null) return;
  const nextCode = window.prompt("Codice commessa:", currentCode || "");
  if (nextCode == null) return;
  const normalized = nextName.trim();
  const normalizedCode = String(nextCode || "").trim();
  if (!normalized) return;
  await db.collection("commesse").doc(commessaId).set({ nome: normalized, codice: normalizedCode }, { merge: true });
  if (selectedCommessaId === commessaId) {
    selectedCommessaName = normalized;
    ui.commessaAttiva.textContent = normalizedCode ? `Commessa selezionata: ${normalized} • Cod. commessa: ${normalizedCode}` : `Commessa selezionata: ${normalized}`;
    updateCommessaContextUI();
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
  chatNotificationsInitialized = false;
}

async function renderChat(messages) {
  if (!currentUser) {
    ui.chatCounter.classList.add("hidden");
    renderChatEmptyState("Fai login per usare la chat.");
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
    renderChatEmptyState();
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

function renderChatEmptyState(message = "Nessun messaggio presente") {
  if (!ui.chatFullList) return;
  ui.chatFullList.innerHTML = `
    <div class="chat-empty-state">
      <div class="chat-empty-icon" aria-hidden="true">💬</div>
      <p>${escapeHTML(message)}</p>
    </div>
  `;
}

function openChatClearConfirmModal() {
  if (!canManageData()) return;
  ui.chatClearConfirmModal?.classList.remove("hidden");
  ui.chatClearConfirmModal?.setAttribute("aria-hidden", "false");
  ui.chatClearConfirmBtn?.focus();
}

function closeChatClearConfirmModal() {
  ui.chatClearConfirmModal?.classList.add("hidden");
  ui.chatClearConfirmModal?.setAttribute("aria-hidden", "true");
}

async function clearCurrentChatMessages() {
  if (!canManageData()) {
    alert("Solo un admin può svuotare la chat.");
    return;
  }

  closeChatClearConfirmModal();
  const visibleMessageIds = chatMessages
    .filter((message) => canViewMessage(message))
    .map((message) => String(message.id || "").trim())
    .filter(Boolean);

  ui.chatClearBtn.disabled = true;
  ui.chatClearConfirmBtn.disabled = true;
  ui.chatFeedback.textContent = visibleMessageIds.length ? "Svuotamento chat in corso..." : "Chat già vuota.";

  try {
    await animateChatDeletion();
    chatMessages = chatMessages.filter((message) => !visibleMessageIds.includes(String(message.id || "")));
    await renderChat(chatMessages);

    if (visibleMessageIds.length) {
      await deleteChatMessagesByIds(visibleMessageIds);
      ui.chatFeedback.textContent = "Chat svuotata.";
    } else {
      renderChatEmptyState();
    }
  } catch (error) {
    console.error("Errore svuotamento chat:", error);
    ui.chatFeedback.textContent = error?.message || "Impossibile svuotare la chat.";
  } finally {
    ui.chatClearConfirmBtn.disabled = false;
    ui.chatClearBtn.disabled = !canManageData();
  }
}

function animateChatDeletion() {
  const messageNodes = Array.from(ui.chatFullList?.querySelectorAll(".chat-message") || []);
  if (!messageNodes.length) return Promise.resolve();
  messageNodes.forEach((node, index) => {
    node.style.setProperty("--chat-delete-delay", `${Math.min(index * 24, 180)}ms`);
    node.classList.add("is-deleting");
  });
  return new Promise((resolve) => window.setTimeout(resolve, 360));
}

async function deleteChatMessagesByIds(messageIds) {
  const uniqueIds = Array.from(new Set(messageIds));
  for (let index = 0; index < uniqueIds.length; index += 450) {
    const batch = db.batch();
    uniqueIds.slice(index, index + 450).forEach((messageId) => {
      batch.delete(db.collection("chatMessages").doc(messageId));
    });
    await batch.commit();
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
  const metadataType = String(message?.metadata?.type || "");
  if (metadataType === "notification_ack" && !canManageData()) return false;
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

function canOpenHoursFromChatAlert(message) {
  if (String(message?.kind || "") !== "system") return false;
  if (!canManageData()) return false;
  const metadata = message?.metadata && typeof message.metadata === "object" ? message.metadata : null;
  if (!metadata) return false;
  return metadata.type === "hours_deadline_alert" && metadata.action === "open_hours";
}

function canMoveImpiantoDoneFromChat(message) {
  if (String(message?.kind || "") !== "system") return false;
  if (!canManageData()) return false;
  const metadata = message?.metadata && typeof message.metadata === "object" ? message.metadata : null;
  if (!metadata) return false;
  if (metadata.type !== "impianto_done_recovery" || metadata.action !== "move_done") return false;
  const commessaId = String(metadata.commessaId || "").trim();
  const impiantoIds = Array.isArray(metadata.impiantoIds) ? metadata.impiantoIds.filter(Boolean) : [];
  return Boolean(commessaId && impiantoIds.length);
}

async function moveImpiantoToDoneFromChat(message) {
  const metadata = message?.metadata && typeof message.metadata === "object" ? message.metadata : null;
  if (!metadata) throw new Error("Messaggio chat senza metadati validi.");
  const commessaId = String(metadata.commessaId || "").trim();
  const commessaName = String(metadata.commessaName || "").trim() || "Commessa";
  const impiantoIds = Array.isArray(metadata.impiantoIds) ? metadata.impiantoIds.map((id) => String(id || "").trim()).filter(Boolean) : [];
  if (!commessaId || !impiantoIds.length) throw new Error("Dati impianto insufficienti per lo spostamento nei FATTI.");
  await setImpiantoDone(commessaId, impiantoIds, true);
  if (selectedCommessaId === commessaId) {
    updateImpiantoLocalState(impiantoIds, {
      done: true,
      doneAt: new Date(),
      doneBy: currentUser?.displayName || currentUser?.email || "Admin"
    });
  }
  scheduleCommessaSheetSync(commessaId, commessaName, 250);
}

async function deleteChatMessageById(messageId) {
  const normalized = String(messageId || "").trim();
  if (!normalized) return;
  await db.collection("chatMessages").doc(normalized).delete();
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
  if (canMoveImpiantoDoneFromChat(message)) {
    const actions = document.createElement("div");
    actions.className = "item-actions";
    const moveDoneBtn = createButton("Sposta nei FATTI", async () => {
      moveDoneBtn.disabled = true;
      try {
        await moveImpiantoToDoneFromChat(message);
        ui.chatFeedback.textContent = "Impianto spostato nei FATTI.";
      } catch (error) {
        console.error("Errore spostamento impianto nei FATTI da chat:", error);
        ui.chatFeedback.textContent = error?.message || "Impossibile spostare l'impianto nei FATTI.";
        moveDoneBtn.disabled = false;
      }
    });
    actions.appendChild(moveDoneBtn);
    return actions;
  }

  if (canOpenHoursFromChatAlert(message)) {
    const actions = document.createElement("div");
    actions.className = "item-actions";
    const openHoursBtn = createButton("Inserisci ore", () => {
      if (!currentUser) return;
      openHoursPage();
      if (ui.hoursDate && message?.metadata?.date) ui.hoursDate.value = message.metadata.date;
      if (ui.hoursFeedback) {
        const commessaName = String(message?.metadata?.commessaName || "").trim();
        ui.hoursFeedback.textContent = commessaName
          ? `Avviso ore: compila la commessa ${commessaName}.`
          : "Avviso ore: compila le ore mancanti.";
      }
    });
    actions.appendChild(openHoursBtn);
    if (canManageData()) {
      const deleteBtn = createButton("Elimina", async () => {
        try {
          await deleteChatMessageById(message.id);
          ui.chatFeedback.textContent = "Messaggio avviso eliminato.";
        } catch (error) {
          console.error("Errore eliminazione messaggio avviso:", error);
          ui.chatFeedback.textContent = "Impossibile eliminare il messaggio avviso.";
        }
      });
      actions.appendChild(deleteBtn);
    }
    return actions;
  }

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
    if (!isCentralDriveConfigured() && !driveAccessToken) {
      throw new Error("Cloud amministratore non configurato. Contatta un admin per inviare foto/video.");
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
        if (!isCentralDriveConfigured() && !driveAccessToken) {
          throw new Error("Cloud amministratore non configurato. Contatta un admin per inviare vocali.");
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
  const connected = Boolean(isConnected || driveBridgeState.configured);
  ui.driveStatus.classList.toggle("status-chip-drive", connected);
  if (connected) {
    ui.driveStatus.textContent = canManageData()
      ? `Cloud centralizzato attivo${driveBridgeState.ownerEmail ? ` (${driveBridgeState.ownerEmail})` : ""}.`
      : "Archiviazione cloud attiva • Cloud centralizzato attivo";
  } else {
    ui.driveStatus.textContent = getCentralDriveNotConfiguredMessage();
  }
}

async function connectGoogleDrive() {
  if (!canManageData()) {
    alert("Solo admin può configurare Google Drive.");
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
  driveRootFolderId = CENTRAL_DRIVE_ROOT_FOLDER_ID;
  driveChatFolderId = "FOTO";
  driveReportsFolderId = "SEGNALAZIONI";
  driveSquadreFolderId = "EXPORT";
  driveHelpCenterFolderId = "EXPORT";
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

  const list = document.getElementById("help-faq-list") || ui.howtoFaqList;
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
    console.warn("Cloud amministratore non configurato: salto export snapshot FAQ.");
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
  const safeName = escapeDriveQueryValue(name);
  const query = `mimeType='application/vnd.google-apps.folder' and trashed=false and name='${safeName}'${parentQuery}`;
  const searchUrl = `https://www.googleapis.com/drive/v3/files?q=${encodeURIComponent(query)}&fields=files(id,name,createdTime)&orderBy=createdTime&pageSize=1`;
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

async function findDriveFoldersByName(name, parentId = "") {
  const parentQuery = parentId ? ` and '${parentId}' in parents` : "";
  const safeName = escapeDriveQueryValue(name);
  const query = `mimeType='application/vnd.google-apps.folder' and trashed=false and name='${safeName}'${parentQuery}`;
  const searchUrl = `https://www.googleapis.com/drive/v3/files?q=${encodeURIComponent(query)}&fields=files(id,name,parents,createdTime)&orderBy=createdTime&pageSize=100`;
  const searchResponse = await driveApiFetch(searchUrl, { method: "GET" });
  return Array.isArray(searchResponse.files) ? searchResponse.files : [];
}

async function moveDriveFileToFolder(fileId, targetParentId, currentParents = []) {
  if (!fileId || !targetParentId) return;
  const removeParents = currentParents.filter((parentId) => parentId && parentId !== targetParentId).join(",");
  const params = new URLSearchParams({ addParents: targetParentId, fields: "id,parents" });
  if (removeParents) params.set("removeParents", removeParents);
  await driveApiFetch(`https://www.googleapis.com/drive/v3/files/${fileId}?${params.toString()}`, {
    method: "PATCH",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({})
  });
}

async function migrateLegacyDriveDataToCentralRoot(options = {}) {
  const { force = false } = options;
  if (!canManageData() || !driveAccessToken) return;
  const migrationKey = `${LEGACY_DRIVE_MIGRATION_KEY}:${CENTRAL_DRIVE_ROOT_FOLDER_ID}`;
  if (!force && sessionStorage.getItem(migrationKey) === "true") return;
  try {
    await ensureDriveFolders();
    const legacyContainerId = await getOrCreateDriveFolder(CENTRAL_DRIVE_LEGACY_FOLDER_NAME, CENTRAL_DRIVE_ROOT_FOLDER_ID);
    for (const legacyName of LEGACY_DRIVE_ROOT_FOLDER_NAMES) {
      const legacyFolders = await findDriveFoldersByName(legacyName);
      for (const legacyFolder of legacyFolders) {
        if (!legacyFolder.id || legacyFolder.id === CENTRAL_DRIVE_ROOT_FOLDER_ID || legacyFolder.id === legacyContainerId) continue;
        const parents = Array.isArray(legacyFolder.parents) ? legacyFolder.parents : [];
        if (parents.includes(legacyContainerId)) continue;
        await moveDriveFileToFolder(legacyFolder.id, legacyContainerId, parents);
      }
    }
    sessionStorage.setItem(migrationKey, "true");
  } catch (error) {
    console.warn("Migrazione dati Drive vecchi non completata:", error);
  }
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

async function getOrCreateCentralDriveTypeFolder(commessaName, driveType) {
  await ensureDriveFolders();
  const commessaFolderId = await getOrCreateDriveFolder(normalizeDriveFolderName(commessaName, CENTRAL_DRIVE_DEFAULT_COMMESSA), driveRootFolderId);
  return getOrCreateDriveFolder(normalizeDriveFolderName(driveType, "EXPORT").toUpperCase(), commessaFolderId);
}

async function getOrCreateCommessaSpreadsheet(commessaId, commessaName) {
  const cachedId = commessaSheetCache.get(commessaId);
  if (cachedId) return { id: cachedId };

  const spreadsheetFolderId = await getOrCreateCentralDriveTypeFolder(commessaName, "EXPORT");
  const commessaData = commesseById.get(commessaId) || {};
  const configuredSheetId = String(commessaData.sheetSpreadsheetId || "").trim();
  if (configuredSheetId) {
    try {
      const existingSheet = await driveApiFetch(`https://www.googleapis.com/drive/v3/files/${configuredSheetId}?fields=id,name,mimeType,parents`, { method: "GET" });
      if (existingSheet && existingSheet.id) {
        const parents = Array.isArray(existingSheet.parents) ? existingSheet.parents : [];
        if (!parents.includes(spreadsheetFolderId)) {
          const removeParents = parents.filter(Boolean).join(",");
          const moveParams = new URLSearchParams({ addParents: spreadsheetFolderId, fields: "id,parents" });
          if (removeParents) moveParams.set("removeParents", removeParents);
          await driveApiFetch(`https://www.googleapis.com/drive/v3/files/${configuredSheetId}?${moveParams.toString()}`, {
            method: "PATCH",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify({})
          });
        }
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

  const safeName = escapeDriveQueryValue(commessaName);
  const query = [
    "mimeType='application/vnd.google-apps.spreadsheet'",
    "trashed=false",
    `'${spreadsheetFolderId}' in parents`,
    `name='Commessa - ${safeName}'`
  ].join(" and ");

  const searchUrl = `https://www.googleapis.com/drive/v3/files?q=${encodeURIComponent(query)}&fields=files(id,name,createdTime)&orderBy=createdTime&pageSize=1`;
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
      parents: [spreadsheetFolderId]
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

function readBlobAsBase64(blob) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => {
      const result = String(reader.result || "");
      resolve(result.includes(",") ? result.split(",").pop() : result);
    };
    reader.onerror = () => reject(reader.error || new Error("Lettura file non riuscita."));
    reader.readAsDataURL(blob);
  });
}

async function uploadBlobThroughCentralBackend(blob, fileName, mimeType, folderId, options = {}) {
  if (!isCentralDriveConfigured()) {
    throw new Error(getCentralDriveNotConfiguredMessage());
  }
  if (!functions || typeof functions.httpsCallable !== "function") {
    throw new Error("Backend Firebase per upload centralizzato non disponibile.");
  }
  const base64 = await readBlobAsBase64(blob);
  const uploadCentralDriveFile = functions.httpsCallable("uploadCentralDriveFile");
  const result = await uploadCentralDriveFile({
    fileName: normalizeDriveFolderName(fileName, "file"),
    mimeType: mimeType || "application/octet-stream",
    base64,
    commessaName: getCurrentDriveCommessaName(options),
    driveType: inferCentralDriveType(folderId, options)
  });
  const data = result?.data || {};
  return {
    fileId: data.fileId || "",
    webViewLink: data.webViewLink || "",
    directUrl: ""
  };
}

async function uploadBlobDirectToAdminDrive(blob, fileName, mimeType, folderId, options = {}) {
  if (!driveAccessToken) {
    throw new Error(getCentralDriveNotConfiguredMessage());
  }
  await ensureDriveFolders();
  const { signal = null } = options;
  const commessaFolderId = await getOrCreateDriveFolder(getCurrentDriveCommessaName(options), CENTRAL_DRIVE_ROOT_FOLDER_ID);
  const typeFolderId = await getOrCreateDriveFolder(inferCentralDriveType(folderId, options), commessaFolderId);
  const metadata = {
    name: normalizeDriveFolderName(fileName, "file"),
    mimeType: mimeType || "application/octet-stream",
    parents: [typeFolderId]
  };
  const form = new FormData();
  form.append("metadata", new Blob([JSON.stringify(metadata)], { type: "application/json" }));
  form.append("file", blob, metadata.name);
  const uploaded = await driveApiFetch("https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart&fields=id,name,webViewLink,webContentLink", {
    method: "POST",
    body: form,
    signal
  });
  return {
    fileId: uploaded.id || "",
    webViewLink: uploaded.webViewLink || "",
    directUrl: ""
  };
}

async function uploadBlobToDrive(blob, fileName, mimeType, folderId, options = {}) {
  if (!isCentralDriveConfigured() && !driveAccessToken) {
    throw new Error(getCentralDriveNotConfiguredMessage());
  }
  if (canManageData() && driveAccessToken) {
    try {
      return await uploadBlobDirectToAdminDrive(blob, fileName, mimeType, folderId, options);
    } catch (error) {
      if (!functions || typeof functions.httpsCallable !== "function") throw error;
      console.warn("Upload diretto admin non riuscito, provo backend centralizzato:", error);
    }
  }
  return uploadBlobThroughCentralBackend(blob, fileName, mimeType, folderId, options);
}

async function backupSquadreSnapshotToDrive(dateKey, squadraPayload) {
  if (!driveAccessToken) return;
  if (!driveSquadreFolderId) await ensureDriveFolders();
  const exportData = await buildAppBackupPayload(dateKey, squadraPayload);
  const blob = new Blob([JSON.stringify(exportData, null, 2)], { type: "application/json" });
  const commessaLabel = String(squadraPayload.commessaNome || "Commessa").replace(/[^\w\-]+/g, "_");
  const fileName = `squadre_${dateKey}_${commessaLabel}.json`;
  await uploadBlobToDrive(blob, fileName, "application/json", driveSquadreFolderId, { driveType: "EXPORT", commessaName: squadraPayload.commessaNome || "Squadre" });
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
    throw new Error(getCentralDriveNotConfiguredMessage());
  }

  const headers = new Headers(options.headers || {});
  headers.set("Authorization", `Bearer ${driveAccessToken}`);

  const response = await fetchWithTimeoutAndRetry(url, {
    ...options,
    headers
  }, {
    timeoutMs: NETWORK_DEFAULT_TIMEOUT_MS,
    retries: 2
  });

  if (response.status === 401 || response.status === 403) {
    const refreshed = await refreshDriveAccessToken();
    if (refreshed) {
      return driveApiFetch(url, options);
    }
    localStorage.removeItem("googleDriveAccessToken");
    updateDriveStatus(false);
    throw new Error("Sessione Drive amministratore scaduta. Solo admin deve ricollegare Google Drive.");
  }

  if (!response.ok) {
    const text = await response.text();
    throw new Error(`Errore Google Drive (${response.status}): ${text.slice(0, 180)}`);
  }

  if (response.status === 204) return {};
  return response.json();
}

function isRetryableNetworkError(error) {
  if (!error) return false;
  if (error.name === "AbortError") return true;
  if (error instanceof TypeError) return true;
  return /network|fetch|failed|timeout/i.test(String(error.message || ""));
}

function shouldRetryHttpStatus(status) {
  return NETWORK_RETRYABLE_STATUS_CODES.has(Number(status));
}

function wait(ms) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

async function fetchWithTimeoutAndRetry(url, options = {}, config = {}) {
  const retries = Number.isFinite(config.retries) ? Math.max(0, config.retries) : 0;
  const timeoutMs = Number.isFinite(config.timeoutMs) ? Math.max(1, config.timeoutMs) : NETWORK_DEFAULT_TIMEOUT_MS;
  const baseDelayMs = Number.isFinite(config.baseDelayMs) ? Math.max(0, config.baseDelayMs) : 600;

  let attempt = 0;
  while (attempt <= retries) {
    const controller = new AbortController();
    const externalSignal = options?.signal || null;
    if (externalSignal?.aborted) {
      throw new DOMException("Operazione annullata", "AbortError");
    }
    const onExternalAbort = () => controller.abort();
    if (externalSignal) externalSignal.addEventListener("abort", onExternalAbort, { once: true });
    const timeoutId = setTimeout(() => controller.abort(), timeoutMs);
    try {
      const mergedOptions = { ...options, signal: controller.signal };
      const response = await fetch(url, mergedOptions);
      clearTimeout(timeoutId);
      if (externalSignal) externalSignal.removeEventListener("abort", onExternalAbort);
      if (attempt < retries && shouldRetryHttpStatus(response.status)) {
        await wait(baseDelayMs * (attempt + 1));
        attempt += 1;
        continue;
      }
      return response;
    } catch (error) {
      clearTimeout(timeoutId);
      if (externalSignal) externalSignal.removeEventListener("abort", onExternalAbort);
      if (attempt >= retries || !isRetryableNetworkError(error)) {
        throw error;
      }
      await wait(baseDelayMs * (attempt + 1));
      attempt += 1;
    }
  }
}

function isProgrammazioneVisibleToCurrentUser(item) {
  if (canManageData()) return true;
  const email = normalizeEmail(currentUser?.email || "");
  const displayName = String(currentUser?.displayName || "").trim().toLowerCase();
  const operators = Array.isArray(item?.operatoriCoinvolti) ? item.operatoriCoinvolti : [];
  const involved = operators.some((entry) => {
    const raw = String(entry || "").trim();
    if (!raw) return false;
    const normalizedEmail = normalizeEmail(raw);
    const normalizedName = raw.toLowerCase();
    return (email && normalizedEmail === email) || (displayName && normalizedName === displayName);
  });
  if (!involved) return false;
  const dateKey = String(item?.data || "");
  if (!dateKey) return false;
  const now = new Date();
  const target = new Date(`${dateKey}T00:00:00`);
  const today = new Date(now.getFullYear(), now.getMonth(), now.getDate());
  const diff = Math.floor((target - today) / 86400000);
  if (target.getDay() === 1) return diff <= 3 && diff >= 0;
  return diff <= 1 && diff >= 0;
}

function programmazioneReminderBadge(dateKey) {
  const today = new Date();
  const target = new Date(`${dateKey}T00:00:00`);
  const diff = Math.floor((target - new Date(today.getFullYear(), today.getMonth(), today.getDate())) / 86400000);
  if (diff === 0) return "📅 Oggi";
  if (diff === 1) return "📌 Domani";
  const day = target.getDay();
  if (day === 1 && diff >= 3 && diff <= 5) return "📌 Programmazione lunedì";
  return "";
}

function renderProgrammazioni() {
  const visible = programmazioni.filter(isProgrammazioneVisibleToCurrentUser);
  const filter = String(ui.programmazioneFilter?.value || "all");
  const today = new Date().toISOString().slice(0, 10);
  const tomorrow = new Date(Date.now() + 86400000).toISOString().slice(0, 10);
  const filtered = visible.filter((item) => {
    if (filter === "oggi") return item.data === today;
    if (filter === "domani") return item.data === tomorrow;
    if (filter === "programmato") return String(item.stato || "") === "Programmato";
    if (filter === "urgente") return String(item.priorita || "") === "Urgente" || String(item.tipo || "") === "urgente";
    if (filter === "fatto") return String(item.stato || "") === "Fatto";
    if (filter === "annullato") return String(item.stato || "") === "Annullato";
    return true;
  });
  if (ui.programmazioneList) {
    ui.programmazioneList.innerHTML = filtered.map((item) => `<article class="simple-list-item ${String(item.stato||"")==="Fatto"?"programmazione-done":""}"><strong>${escapeHTML(item.ora||"--:--")} - ${escapeHTML(item.oraFine||"--:--")} ${escapeHTML(item.tipoLabel||item.tipo||"")}</strong><p>${escapeHTML(item.commessa||"")} • ${escapeHTML(item.zona||"")}</p><p>${escapeHTML((item.note||"").slice(0,80))}</p><p>${escapeHTML(item.stato||"")} ${escapeHTML(programmazioneReminderBadge(item.data)||"")}</p>${canManageData()?`<div class='item-actions'><button type='button' class='btn' data-edit-programmazione='${escapeHTML(item.id||"")}'>Modifica</button><button type='button' class='btn btn-danger' data-delete-programmazione='${escapeHTML(item.id||"")}'>Elimina</button></div>`:""}</article>`).join("") || "<p class='muted'>Nessuna programmazione visibile.</p>";
    ui.programmazioneList.querySelectorAll("[data-edit-programmazione]").forEach((btn) => btn.addEventListener("click", () => openEditProgrammazione(btn.getAttribute("data-edit-programmazione"))));
    ui.programmazioneList.querySelectorAll("[data-delete-programmazione]").forEach((btn) => btn.addEventListener("click", () => deleteProgrammazioneById(btn.getAttribute("data-delete-programmazione"))));
  }
  if (ui.programmazioniHomeCard && ui.programmazioniHomeList) {
    const homeItems = visible.filter((item) => Boolean(programmazioneReminderBadge(item.data)));
    ui.programmazioniHomeCard.classList.toggle("hidden", !homeItems.length);
    ui.programmazioniHomeCard.setAttribute("aria-hidden", homeItems.length ? "false" : "true");
    ui.programmazioniHomeList.innerHTML = homeItems.map((item) => `<article class="simple-list-item"><strong>${escapeHTML(programmazioneReminderBadge(item.data))}</strong><p>${escapeHTML(item.ora||"")} • ${escapeHTML(item.tipoLabel||item.tipo||"")} • ${escapeHTML(item.commessa||"")}</p></article>`).join("");
  }
}

function openEditProgrammazione(id) {
  if (!canManageData()) return;
  const item = programmazioni.find((row) => row.id === id);
  if (!item) return;
  populateProgrammazioneFormOptions();
  ui.programmaId.value = item.id || "";
  document.getElementById("programma-data").value = item.data || "";
  document.getElementById("programma-ora").value = item.ora || "";
  document.getElementById("programma-ora-fine").value = item.oraFine || "";
  document.getElementById("programma-zona").value = item.zona || "";
  document.getElementById("programma-tipo").value = item.tipo || "sfalcio";
  document.getElementById("programma-priorita").value = item.priorita || "normale";
  document.getElementById("programma-note").value = item.note || "";
  document.getElementById("programma-commessa").value = item.commessa || "";
  programmazioneOperatorAutocomplete = buildProgrammazioneAutocomplete(ui.programmaOperatoriAutocomplete, "Operatori coinvolti", getProgrammazioneOperatorOptions(), item.operatoriCoinvolti || []);
  programmazioneMezziAutocomplete = buildProgrammazioneAutocomplete(ui.programmaMezziAutocomplete, "Mezzi / attrezzature", getProgrammazioneMezziOptions(), item.mezziAssegnati || []);
  ui.programmazioneDeleteBtn?.classList.remove("hidden");
  ui.programmazioneDialog?.showModal();
}

async function saveProgrammazione(event) {
  event.preventDefault();
  if (!canManageData()) return;
  const payload = {
    updatedAt: firebase.firestore.FieldValue.serverTimestamp(),
    data: document.getElementById("programma-data")?.value || "",
    ora: document.getElementById("programma-ora")?.value || "",
    oraFine: document.getElementById("programma-ora-fine")?.value || "",
    tipo: document.getElementById("programma-tipo")?.value || "",
    tipoLabel: document.getElementById("programma-tipo")?.selectedOptions?.[0]?.textContent || "",
    commessa: document.getElementById("programma-commessa")?.value || "",
    zona: document.getElementById("programma-zona")?.value || "",
    operatoriCoinvolti: programmazioneOperatorAutocomplete?.getValues?.() || [],
    priorita: document.getElementById("programma-priorita")?.value || "normale",
    note: document.getElementById("programma-note")?.value || "",
    stato: "Programmato",
    mezziAssegnati: programmazioneMezziAutocomplete?.getValues?.() || [],
    createdBy: currentUser?.email || ""
  };
  const id = String(ui.programmaId?.value || "").trim();
  if (id) {
    await db.collection("programmazioni").doc(id).set(payload, { merge: true });
    alert("Programmazione aggiornata correttamente");
    await syncProgrammazioneToSquadra(id, payload, { remove: false });
  } else {
    payload.createdAt = firebase.firestore.FieldValue.serverTimestamp();
    const docRef = await db.collection("programmazioni").add(payload);
    await syncProgrammazioneToSquadra(docRef.id, payload, { remove: false });
  }
  ui.programmazioneDialog?.close();
  ui.programmazioneForm?.reset();
}

async function syncProgrammazioneToSquadra(programmazioneId, payload, { remove = false } = {}) {
  const commessa = Array.from(commesseById.values()).find((row) => String(row.nome || "").trim() === String(payload.commessa || "").trim());
  if (!commessa?.id || !payload?.data) return;
  const historyRef = db.collection("squadreStorico").doc(`${payload.data}__${commessa.id}`);
  if (remove) {
    await historyRef.set({ autogeneratedFromProgrammazione: firebase.firestore.FieldValue.delete(), autoProgrammazioneId: firebase.firestore.FieldValue.delete() }, { merge: true });
    return;
  }
  await historyRef.set({
    commessaId: commessa.id,
    commessaNome: commessa.nome || "Commessa",
    riferimentoData: payload.data,
    dateKey: payload.data,
    autoProgrammazioneId: programmazioneId,
    autogeneratedFromProgrammazione: true,
    squadre: [{ personale: (payload.operatoriCoinvolti || []).join(", "), mezzi: (payload.mezziAssegnati || []).join(", "), impianti: payload.zona || "", note: payload.note || "", orario: `${payload.ora || ""}-${payload.oraFine || ""}` }]
  }, { merge: true });
}

async function deleteProgrammazioneById(id) {
  if (!canManageData() || !id) return;
  const item = programmazioni.find((row) => row.id === id);
  if (!item) return;
  if (!window.confirm("Sei sicuro di voler eliminare questa programmazione?")) return;
  await syncProgrammazioneToSquadra(id, item, { remove: true });
  await db.collection("programmazioni").doc(id).delete();
}
async function deleteProgrammazioneFromForm() { return deleteProgrammazioneById(ui.programmaId?.value || ""); }
