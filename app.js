console.log("APP START");

const firebaseConfig = window.firebaseConfig || {};
const requiredFirebaseConfigKeys = ["apiKey", "authDomain", "projectId", "appId"];
const missingFirebaseConfigKeys = requiredFirebaseConfigKeys.filter((key) => !String(firebaseConfig?.[key] || "").trim());
let firebaseApp = null;
let auth = null;
let db = null;
let functions = null;
let firebaseInitError = null;

try {
  if (!window.firebase || typeof firebase.initializeApp !== "function") {
    throw new Error("Firebase SDK non caricato o versione non compatibile.");
  }
  if (missingFirebaseConfigKeys.length) {
    throw new Error(`Configurazione Firebase incompleta. Campi mancanti: ${missingFirebaseConfigKeys.join(", ")}`);
  }

  firebaseApp = firebase.apps && firebase.apps.length
    ? firebase.app()
    : firebase.initializeApp(firebaseConfig);
  auth = typeof firebase.auth === "function" ? firebase.auth() : null;
  db = typeof firebase.firestore === "function" ? firebase.firestore() : null;
  functions = typeof firebase.functions === "function" ? firebase.functions() : null;

  if (!auth || !db) {
    throw new Error("Firebase Auth o Firestore non disponibili nel SDK caricato.");
  }

  console.log("FIREBASE INIT OK", {
    appName: firebaseApp.name,
    projectId: firebaseConfig?.projectId || "non impostato",
    sdkVersion: firebase.SDK_VERSION || "non disponibile"
  });
  console.log("FIREBASE READY");
} catch (error) {
  firebaseInitError = error;
  console.error("FIREBASE INIT ERROR", error);
}

const PERSISTED_SESSION_KEY = "heraPersistedUserSession";
const PERSISTED_SESSION_VERSION = 1;

function getCapacitorPreferencesPlugin() {
  return window.Capacitor
    && window.Capacitor.Plugins
    && window.Capacitor.Plugins.Preferences
    && typeof window.Capacitor.Plugins.Preferences.get === "function"
    ? window.Capacitor.Plugins.Preferences
    : null;
}

function buildPersistedSession(user, overrides = {}) {
  const previousSession = readLocalPersistedSession();
  const samePreviousUser = String(previousSession?.uid || "") === String(overrides.uid ?? user?.uid ?? "");
  const email = String(overrides.email ?? user?.email ?? "");
  const displayName = String(overrides.displayName ?? user?.displayName ?? (email || "Utente"));
  const isAdmin = Boolean(overrides.isAdmin ?? canManageData());
  const role = String(overrides.role || overrides.ruolo || (isAdmin ? "admin" : "user"));
  const teamId = String(overrides.teamId || user?.teamId || "").trim();
  const authorizationStatus = String(overrides.statoAccount || overrides.accountStatus || "").trim();
  const accessApproved = authorizationStatus
    ? authorizationStatus === "attivo"
    : Boolean(overrides.accessApproved ?? (samePreviousUser && previousSession?.accessApproved));
  return {
    version: PERSISTED_SESSION_VERSION,
    uid: String(overrides.uid ?? user?.uid ?? ""),
    email,
    displayName,
    userName: displayName,
    role,
    ruolo: role,
    isAdmin,
    admin: isAdmin,
    teamId: teamId || null,
    accessApproved,
    banned: Boolean(overrides.banned),
    bannedReason: overrides.bannedReason || null,
    bannedAt: overrides.bannedAt || null,
    bannedBy: overrides.bannedBy || null,
    lastLoginAt: overrides.lastLoginAt || new Date().toISOString(),
    updatedAt: new Date().toISOString()
  };
}

function isValidPersistedSession(session) {
  return Boolean(
    session
    && Number(session.version) === PERSISTED_SESSION_VERSION
    && String(session.uid || "").trim()
    && String(session.email || "").includes("@")
    && String(session.lastLoginAt || "").trim()
  );
}

function isPersistedApprovalValid(session, user) {
  return Boolean(
    isValidPersistedSession(session)
    // Le sessioni create prima dell'introduzione dell'approvazione non hanno
    // questo campo: erano gia utenti ammessi e vengono migrate al primo avvio.
    && session.accessApproved !== false
    && session.banned !== true
    && String(session.uid) === String(user?.uid || "")
    && normalizeEmail(session.email) === normalizeEmail(user?.email)
  );
}

function readLocalPersistedSession() {
  try {
    const raw = localStorage.getItem(PERSISTED_SESSION_KEY);
    if (!raw) return null;
    const parsed = JSON.parse(raw);
    return isValidPersistedSession(parsed) ? parsed : null;
  } catch (error) {
    console.warn("Sessione locale corrotta: verr√† ignorata.", error);
    return null;
  }
}

async function readPersistedSession() {
  const localSession = readLocalPersistedSession();
  const preferences = getCapacitorPreferencesPlugin();
  if (!preferences) return localSession;
  try {
    const result = await preferences.get({ key: PERSISTED_SESSION_KEY });
    if (!result?.value) return localSession;
    const nativeSession = JSON.parse(result.value);
    return isValidPersistedSession(nativeSession) ? nativeSession : localSession;
  } catch (error) {
    console.warn("Sessione Capacitor Preferences non leggibile: uso localStorage se disponibile.", error);
    return localSession;
  }
}

async function savePersistedSession(user, overrides = {}) {
  if (!user?.uid) return null;
  const session = buildPersistedSession(user, overrides);
  const serialized = JSON.stringify(session);
  try {
    localStorage.setItem(PERSISTED_SESSION_KEY, serialized);
  } catch (error) {
    console.warn("Salvataggio sessione localStorage non riuscito:", error);
  }
  const preferences = getCapacitorPreferencesPlugin();
  if (preferences && typeof preferences.set === "function") {
    try {
      await preferences.set({ key: PERSISTED_SESSION_KEY, value: serialized });
    } catch (error) {
      console.warn("Salvataggio sessione Capacitor Preferences non riuscito:", error);
    }
  }
  return session;
}

async function clearPersistedSession() {
  try {
    localStorage.removeItem(PERSISTED_SESSION_KEY);
  } catch (error) {
    console.warn("Cancellazione sessione localStorage non riuscita:", error);
  }
  const preferences = getCapacitorPreferencesPlugin();
  if (preferences && typeof preferences.remove === "function") {
    try {
      await preferences.remove({ key: PERSISTED_SESSION_KEY });
    } catch (error) {
      console.warn("Cancellazione sessione Capacitor Preferences non riuscita:", error);
    }
  }
}

function applyPersistedSessionPreview(session) {
  if (!isValidPersistedSession(session) || session.banned) return false;
  currentUser = {
    uid: session.uid,
    email: session.email,
    displayName: session.displayName || session.userName || session.email,
    teamId: session.teamId || "",
    persistedOnly: true
  };
  setAuthenticationGateState("checking", "Sessione salvata trovata. Ripristino accesso in corso...");
  if (ui.user) ui.user.textContent = `Sessione salvata: ${session.email}`;
  if (ui.userName) ui.userName.textContent = `Nome utente: ${session.displayName || session.userName || "Nome non disponibile"}`;
  return true;
}

async function verifyPersistedSessionAgainstDatabase(user, savedSession) {
  if (!user?.uid || !savedSession || !db) return { valid: true, profile: null };
  const doc = await db.collection("platformUsers").doc(user.uid).get();
  if (!doc.exists) return { valid: false, profile: null };
  const profile = doc.data() || {};
  const normalizedEmail = normalizeEmail(profile.email || user.email);
  const isAdminUser = isBuiltInSuperAdminEmail(normalizedEmail) || adminEmails.has(normalizedEmail);
  await savePersistedSession(user, {
    ...profile,
    uid: user.uid,
    email: profile.email || user.email || "",
    displayName: profile.displayName || user.displayName || user.email || "Utente",
    isAdmin: isAdminUser,
    role: profile.role || profile.ruolo || (isAdminUser ? "admin" : "user"),
    teamId: profile.teamId || "",
    lastLoginAt: savedSession.lastLoginAt || new Date().toISOString()
  });
  return { valid: true, profile, banned: Boolean(profile.banned) };
}

function buildBannedWhatsAppUrl(profile = currentUserBanProfile) {
  const now = new Date().toLocaleString("it-IT");
  const name = profile?.displayName || currentUser?.displayName || currentUser?.email || "Utente";
  const email = profile?.email || currentUser?.email || "";
  const text = `Ciao Admin, ti chiedo di riattivare il mio accesso alla Hera App.

Nome utente: ${name}
Email: ${email}
Data richiesta: ${now}

Grazie.`;
  return `https://wa.me/393892352575?text=${encodeURIComponent(text)}`;
}

function openBannedAccessRequest() {
  window.open(buildBannedWhatsAppUrl(), "_blank", "noopener,noreferrer");
}

function isCurrentUserBanned() {
  return Boolean(currentUserBanProfile?.banned);
}

const DEFAULT_PUSH_PUBLIC_VAPID_KEY = "BLWYWSC_rEbfAoOnOaO6JYhaYVBCa7IDZaN-2cGMt6uqUYLWwl6mKq8hng9V5B5GPVUOlgjLPLhqz2KvdsuJUoAA";
const FIRESTORE_PERSISTENCE_RECOVERY_KEY = "heraFirestorePersistenceRecoveryAttempted";
let firebaseMessaging = null;
let authLocalPersistencePromise = null;
let authStateResolved = false;

function getFirebaseLocalAuthPersistence() {
  return (firebase && firebase.auth && firebase.auth.browserLocalPersistence)
    || (firebase
      && firebase.auth
      && firebase.auth.Auth
      && firebase.auth.Auth.Persistence
      && firebase.auth.Auth.Persistence.LOCAL);
}

function ensureAuthLocalPersistence() {
  if (!auth || typeof auth.setPersistence !== "function") return Promise.resolve(false);
  if (authLocalPersistencePromise) return authLocalPersistencePromise;

  const localPersistence = getFirebaseLocalAuthPersistence();
  if (!localPersistence) {
    console.warn("Persistenza Firebase Auth LOCAL non disponibile nel SDK caricato.");
    return Promise.resolve(false);
  }

  authLocalPersistencePromise = auth.setPersistence(localPersistence)
    .then(() => {
      console.log("AUTH PERSISTENCE LOCAL READY");
      return true;
    })
    .catch((error) => {
      console.error("Errore impostazione persistenza Firebase Auth LOCAL:", error);
      return false;
    });

  return authLocalPersistencePromise;
}

function getSessionStorageValue(key) {
  try {
    return sessionStorage.getItem(key);
  } catch (error) {
    console.warn("SessionStorage non disponibile per la recovery Firestore:", error);
    return null;
  }
}

function setSessionStorageValue(key, value) {
  try {
    sessionStorage.setItem(key, value);
    return true;
  } catch (error) {
    console.warn("SessionStorage non disponibile per la recovery Firestore:", error);
    return false;
  }
}

function getFirebaseErrorMessage(error) {
  return String(error?.message || error || "");
}

function getFirebaseErrorCode(error) {
  return String(error?.code || "").toLowerCase();
}

function isFirestoreInternalAssertionError(error) {
  const message = getFirebaseErrorMessage(error);
  return /FIRESTORE/i.test(message) && /INTERNAL ASSERTION FAILED|Unexpected state/i.test(message);
}

function isFirestoreInternalError(error) {
  const code = getFirebaseErrorCode(error);
  const message = getFirebaseErrorMessage(error);
  return code === "internal"
    || code === "firestore/internal"
    || /INTERNAL ASSERTION FAILED|Unexpected state/i.test(message)
    || isFirestoreInternalAssertionError(error);
}

function formatLoginError(error) {
  const message = getFirebaseErrorMessage(error);
  if (isFirestoreInternalAssertionError(error) || isFirestoreInternalError(error)) {
    return "Abbiamo rilevato un problema temporaneo con la cache locale dei dati. Chiudi eventuali altre schede dell'app e riapri l'app; se il problema continua, svuota i dati del sito e riprova il login.";
  }
  if (/FIRESTORE/i.test(message)) {
    return "Non riesco a collegarmi ai dati dell'app in questo momento. Controlla la connessione, chiudi eventuali altre schede aperte e riprova il login.";
  }
  return message || "Errore sconosciuto durante il login. Riprova tra qualche istante.";
}

function markFirestorePersistenceRecovery(reason) {
  const currentValue = getSessionStorageValue(FIRESTORE_PERSISTENCE_RECOVERY_KEY) || "";
  const attempts = new Set(currentValue.split(",").filter(Boolean));
  if (attempts.has(reason)) return false;
  attempts.add(reason);
  setSessionStorageValue(FIRESTORE_PERSISTENCE_RECOVERY_KEY, Array.from(attempts).join(","));
  return true;
}

function reloadAfterFirestorePersistenceRecovery(reason) {
  const recoveryUrl = new URL(window.location.href);
  recoveryUrl.searchParams.set("firestoreRecovery", reason);
  recoveryUrl.searchParams.set("firestoreRecoveryTs", String(Date.now()));
  window.location.replace(recoveryUrl.toString());
}

function recoverFirestorePersistence(error) {
  const code = getFirebaseErrorCode(error);

  if (code === "failed-precondition") {
    if (markFirestorePersistenceRecovery("failed-precondition")) {
      console.warn("Persistenza Firestore disattivata: l'app sembra aperta in pi√π schede.", error);
    }
    return;
  }

  if (code === "unimplemented") {
    if (markFirestorePersistenceRecovery("unimplemented")) {
      console.warn("Persistenza Firestore non supportata da questo browser/dispositivo.", error);
    }
    return;
  }

  if (!isFirestoreInternalError(error)) return;

  if (!markFirestorePersistenceRecovery("internal")) {
    console.warn("Recovery Firestore interna gi√† tentata in questa sessione, evito un nuovo reload.", error);
    return;
  }

  console.warn("Errore interno Firestore rilevato: non cancello automaticamente IndexedDB/cache locale all'avvio. Ricarico una sola volta senza clearPersistence.", error);
  reloadAfterFirestorePersistenceRecovery("internal");
}

function enableFirestorePersistence() {
  if (!db || typeof db.enablePersistence !== "function") return;
  db.enablePersistence({ synchronizeTabs: true }).catch((error) => {
    const code = getFirebaseErrorCode(error) || "unknown";
    console.warn("Persistenza offline Firestore non disponibile:", code, error);
    recoverFirestorePersistence(error);
  });
}

enableFirestorePersistence();

if (!firebaseInitError && window.firebase && firebase.messaging && typeof firebase.messaging === "function") {
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
  updateAppBtn: document.getElementById("update-app-btn"),
  menuToggleBtn: document.getElementById("menu-toggle-btn"),
  menuCloseBtn: document.getElementById("menu-close-btn"),
  installAppBtn: document.getElementById("install-app-btn"),
  sideMenu: document.getElementById("side-menu"),
  menuOverlay: document.getElementById("menu-overlay"),
  whazzupPreparingFeedback: document.getElementById("whazzup-preparing-feedback"),
  authGate: document.getElementById("auth-gate"),
  authGateLoginBtn: document.getElementById("auth-gate-login-btn"),
  bannedRequestAccessBtn: document.getElementById("banned-request-access-btn"),
  authGateMessage: document.getElementById("auth-gate-message"),
  authEmailForm: document.getElementById("auth-email-form"),
  authEmailInput: document.getElementById("auth-email-input"),
  authPasswordInput: document.getElementById("auth-password-input"),
  authEmailLoginBtn: document.getElementById("auth-email-login-btn"),
  authEmailFeedback: document.getElementById("auth-email-feedback"),
  biometricOfferDialog: document.getElementById("biometric-offer-dialog"),
  biometricOfferFeedback: document.getElementById("biometric-offer-feedback"),
  biometricEnableBtn: document.getElementById("biometric-enable-btn"),
  biometricNotNowBtn: document.getElementById("biometric-not-now-btn"),
  biometricSecuritySettings: document.getElementById("biometric-security-settings"),
  biometricToggle: document.getElementById("biometric-toggle"),
  biometricSettingsFeedback: document.getElementById("biometric-settings-feedback"),
  biometricVerifyBtn: document.getElementById("biometric-verify-btn"),
  biometricDisableBtn: document.getElementById("biometric-disable-btn"),
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
  weatherAlertSafetyPage: document.getElementById("weather-alert-safety-page"),
  weatherAlertSafetyBackBtn: document.getElementById("weather-alert-safety-back-btn"),
  weatherAlertSafetySubtitle: document.getElementById("weather-alert-safety-subtitle"),
  weatherAlertSafetyContent: document.getElementById("weather-alert-safety-content"),
  weatherAlertSafetyConfirmBtn: document.getElementById("weather-alert-safety-confirm-btn"),
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
  impiantoSafetyTitle: document.getElementById("impianto-safety-title"),
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
  mapLocationWarning: document.getElementById("map-location-warning"),
  mapEnableLocationBtn: document.getElementById("map-enable-location-btn"),
  mapRetryLocationBtn: document.getElementById("map-retry-location-btn"),
  mapLocationHelpBtn: document.getElementById("map-location-help-btn"),
  mapLocationHelpPanel: document.getElementById("map-location-help-panel"),
  mapLocationWarningTitle: document.getElementById("map-location-warning-title"),
  mapLocationWarningText: document.getElementById("map-location-warning-text"),
  mapLocationWarningPlatform: document.getElementById("map-location-warning-platform"),
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
  operatorGreeting: document.getElementById("operator-greeting"),
  connectionIndicator: document.getElementById("connection-indicator"),
  offlineModeIndicator: document.getElementById("offline-mode-indicator"),
  syncProgressOverlay: document.getElementById("sync-progress-overlay"),
  syncProgressTitle: document.getElementById("sync-progress-title"),
  syncProgressDetail: document.getElementById("sync-progress-detail"),
  syncProgressList: document.getElementById("sync-progress-list"),
  mapImpiantoDetailPanel: document.getElementById("map-impianto-detail-panel"),
  mapImpiantoDetailBody: document.getElementById("map-impianto-detail-body"),
  biogasMapPage: document.getElementById("biogas-map-page"),
  biogasMapBackBtn: document.getElementById("biogas-map-back-btn"),
  biogasDistanceIndicator: document.getElementById("biogas-distance-indicator"),
  biogasMapSettingsBtn: document.getElementById("biogas-map-settings-btn"),
  biogasMapControls: document.getElementById("biogas-map-controls"),
  biogasMapLayerBtn: document.getElementById("biogas-map-layer-btn"),
  biogasMapToggleBtn: document.getElementById("biogas-map-toggle-btn"),
  biogasMapRefreshBtn: document.getElementById("biogas-map-refresh-btn"),
  biogasMapDeleteBtn: document.getElementById("biogas-map-delete-btn"),
  biogasMapAddPipesBtn: document.getElementById("biogas-map-add-pipes-btn"),
  biogasMapFileInput: document.getElementById("biogas-map-file-input"),
  biogasMapSearch: document.getElementById("biogas-map-search"),
  biogasMapStatus: document.getElementById("biogas-map-status"),
  biogasMapLastUpdate: document.getElementById("biogas-map-last-update"),
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
  mezzoPosti: document.getElementById("mezzo-posti"),
  mezziLista: document.getElementById("mezzi-lista"),
  squadraForm: document.getElementById("squadra-form"),
  squadraCommessa: document.getElementById("squadra-commessa"),
  squadraRows: document.getElementById("squadra-rows"),
  addSquadraRowBtn: document.getElementById("add-squadra-row-btn"),
  squadraRiferimento: document.getElementById("squadra-riferimento"),
  squadraCalendarDate: document.getElementById("squadra-calendar-date"),
  squadraHint: document.getElementById("squadra-hint"),
  squadraFeedback: document.getElementById("squadra-feedback"),
  squadreNextAction: document.getElementById("squadre-next-action"),
  squadreLista: document.getElementById("squadre-lista"),
  squadreImpiantoDetail: document.getElementById("squadre-impianto-detail"),
  squadreImpiantoDetailTitle: document.getElementById("squadre-impianto-detail-title"),
  squadreImpiantoDetailCommessa: document.getElementById("squadre-impianto-detail-commessa"),
  squadreImpiantoDetailBody: document.getElementById("squadre-impianto-detail-body"),
  squadreImpiantoPositionFeedback: document.getElementById("squadre-impianto-position-feedback"),
  squadreImpiantoNavigateBtn: document.getElementById("squadre-impianto-navigate-btn"),
  squadreImpiantoBackBtn: document.getElementById("squadre-impianto-back-btn"),
  toggleCommesseHomeBtn: document.getElementById("toggle-commesse-home-btn"),
  squadreFilterControls: document.getElementById("squadre-filter-controls"),
  squadreFilterDate: document.getElementById("squadre-filter-date"),
  squadreFilterClearBtn: document.getElementById("squadre-filter-clear-btn"),
  snowSquadreFilterControls: document.getElementById("snow-squadre-filter-controls"),
  snowSquadreFilterDate: document.getElementById("snow-squadre-filter-date"),
  snowSquadreFilterClearBtn: document.getElementById("snow-squadre-filter-clear-btn"),
  snowServiceBtn: document.getElementById("snow-service-btn"),
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
  openControlCenterBtn: document.getElementById("open-control-center-btn"),
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
  ferieForm: document.getElementById("ferie-form"),
  ferieCommessa: document.getElementById("ferie-commessa"),
  ferieOperatore: document.getElementById("ferie-operatore"),
  ferieInizio: document.getElementById("ferie-inizio"),
  ferieFine: document.getElementById("ferie-fine"),
  ferieNote: document.getElementById("ferie-note"),
  ferieList: document.getElementById("ferie-list"),
  ferieCheckCommessa: document.getElementById("ferie-check-commessa"),
  ferieCheckStart: document.getElementById("ferie-check-start"),
  ferieCheckEnd: document.getElementById("ferie-check-end"),
  ferieCheckBtn: document.getElementById("ferie-check-btn"),
  ferieCalendarResult: document.getElementById("ferie-calendar-result"),
  programmazioniHomeCard: document.getElementById("programmazioni-home-card"),
  programmazioniHomeList: document.getElementById("programmazioni-home-list"),
  panelBanner: document.getElementById("panel-banner"),
  commesseManageList: document.getElementById("commesse-manage-list"),
  adminUserForm: document.getElementById("admin-user-form"),
  adminUserEmail: document.getElementById("admin-user-email"),
  adminUsersList: document.getElementById("admin-users-list"),
  userBanList: document.getElementById("user-ban-list"),
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
  activeUsersDetailPage: document.getElementById("active-users-detail-page"),
  controlCenterPage: document.getElementById("control-center-page"),
  controlCenterContent: document.getElementById("control-center-content"),
  controlCenterResults: document.getElementById("control-center-results"),
  runControlCheckBtn: document.getElementById("run-control-check-btn"),
  backFromControlCenterBtn: document.getElementById("back-from-control-center-btn"),
  activeUsersBackBtn: document.getElementById("active-users-back-btn"),
  activeUsersAccessMessage: document.getElementById("active-users-access-message"),
  activeUsersAdminConsole: document.getElementById("active-users-admin-console"),
  activeUsersPeriodSelect: document.getElementById("active-users-period-select"),
  activeUsersStartDate: document.getElementById("active-users-start-date"),
  activeUsersEndDate: document.getElementById("active-users-end-date"),
  activeUsersRefreshBtn: document.getElementById("active-users-refresh-btn"),
  activeUsersSearchUser: document.getElementById("active-users-search-user"),
  activeUsersSearchCommessa: document.getElementById("active-users-search-commessa"),
  activeUsersSearchImpianto: document.getElementById("active-users-search-impianto"),
  activeUsersFilterOperator: document.getElementById("active-users-filter-operator"),
  activeUsersFilterAction: document.getElementById("active-users-filter-action"),
  activeUsersErrorsOnly: document.getElementById("active-users-errors-only"),
  activeUsersTopSummary: document.getElementById("active-users-top-summary"),
  activeUsersFilterToggle: document.getElementById("active-users-filter-toggle"),
  activeUsersFilterPanel: document.getElementById("active-users-filter-panel"),
  activeUsersDashboard: document.getElementById("active-users-dashboard"),
  activeUsersCardDetail: document.getElementById("active-users-card-detail"),
  activeUsersFullToggle: document.getElementById("active-users-full-toggle"),
  activeUsersLogToggle: document.getElementById("active-users-log-toggle"),
  activeUsersNowList: document.getElementById("active-users-now-list"),
  activeUsersFullList: document.getElementById("active-users-full-list"),
  activeUsersUserDetail: document.getElementById("active-users-user-detail"),
  activeUsersLogList: document.getElementById("active-users-log-list"),
  userActivityPage: document.getElementById("user-activity-page"),
  userActivityBackBtn: document.getElementById("user-activity-back-btn"),
  userActivityAccessMessage: document.getElementById("user-activity-access-message"),
  userActivityAdminContent: document.getElementById("user-activity-admin-content"),
  userActivitySummary: document.getElementById("user-activity-summary"),
  userActivityDate: document.getElementById("user-activity-date"),
  userActivityCount: document.getElementById("user-activity-count"),
  userActivityTimeline: document.getElementById("user-activity-timeline"),
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
  homeCalendarBtn: document.getElementById("home-calendar-btn"),
  homeSegnalazioniBtn: document.getElementById("home-segnalazioni-btn"),
  todayCommesseBtn: document.getElementById("today-commesse-btn"),
  todayHoursBtn: document.getElementById("today-hours-btn"),
  todayMezziBtn: document.getElementById("today-mezzi-btn"),
  todayAlertsBtn: document.getElementById("today-alerts-btn"),
  todayCommesseCount: document.getElementById("today-commesse-count"),
  todayHoursCount: document.getElementById("today-hours-count"),
  todayMezziCount: document.getElementById("today-mezzi-count"),
  todayAlertsCount: document.getElementById("today-alerts-count"),
  todaySquadsSection: document.getElementById("today-squads-section"),
  userCard: document.getElementById("user-card"),
  userConnectionBar: document.getElementById("user-connection-bar"),
  profileSummaryDetails: document.getElementById("profile-summary-details"),
  userToggleBtn: document.getElementById("user-toggle-btn"),
  userDetailsPanel: document.getElementById("user-details-panel"),
  weatherSummary: document.getElementById("weather-summary"),
  weatherDiagnostics: document.getElementById("weather-diagnostics"),
  weatherExpandedContent: document.getElementById("weather-expanded-content"),
  weatherExternalDetailBtn: document.getElementById("weather-external-detail-btn"),
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
  fuelRadius: document.getElementById("fuel-radius"),
  fuelSearchBtn: document.getElementById("fuel-search-btn"),
  fuelFilterSummary: document.getElementById("fuel-filter-summary"),
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
  calendarPage: document.getElementById("calendar-page"),
  calendarChoiceCard: document.getElementById("calendar-choice-card"),
  calendarChoiceBackBtn: document.getElementById("calendar-choice-back-btn"),
  calendarChoiceHoursBtn: document.getElementById("calendar-choice-hours-btn"),
  calendarChoiceSharedBtn: document.getElementById("calendar-choice-shared-btn"),
  calendarHeroCard: document.getElementById("calendar-hero-card"),
  calendarMainCard: document.getElementById("calendar-main-card"),
  calendarDayCard: document.getElementById("calendar-day-card"),
  calendarPageHeading: document.getElementById("calendar-page-heading"),
  calendarPageDescription: document.getElementById("calendar-page-description"),
  calendarHoursTab: document.getElementById("calendar-hours-tab"),
  calendarSharedTab: document.getElementById("calendar-shared-tab"),
  backFromCalendarBtn: document.getElementById("back-from-calendar-btn"),
  calendarNewEventBtn: document.getElementById("calendar-new-event-btn"),
  calendarPrevBtn: document.getElementById("calendar-prev-btn"),
  calendarTodayBtn: document.getElementById("calendar-today-btn"),
  calendarNextBtn: document.getElementById("calendar-next-btn"),
  calendarMonthTitle: document.getElementById("calendar-month-title"),
  calendarGrid: document.getElementById("calendar-grid"),
  calendarFeedback: document.getElementById("calendar-feedback"),
  calendarSelectedDayTitle: document.getElementById("calendar-selected-day-title"),
  calendarSelectedDaySummary: document.getElementById("calendar-selected-day-summary"),
  calendarAddSelectedDayBtn: document.getElementById("calendar-add-selected-day-btn"),
  calendarDayEvents: document.getElementById("calendar-day-events"),
  calendarEventDialog: document.getElementById("calendar-event-dialog"),
  calendarEventForm: document.getElementById("calendar-event-form"),
  calendarEventFormTitle: document.getElementById("calendar-event-form-title"),
  calendarEventId: document.getElementById("calendar-event-id"),
  calendarEventType: document.getElementById("calendar-event-type"),
  calendarEventTitle: document.getElementById("calendar-event-title"),
  calendarEventStartDate: document.getElementById("calendar-event-start-date"),
  calendarEventEndDate: document.getElementById("calendar-event-end-date"),
  calendarEventAllDay: document.getElementById("calendar-event-all-day"),
  calendarEventTimeFields: document.getElementById("calendar-event-time-fields"),
  calendarEventStartTime: document.getElementById("calendar-event-start-time"),
  calendarEventEndTime: document.getElementById("calendar-event-end-time"),
  calendarEventCommessa: document.getElementById("calendar-event-commessa"),
  calendarEventCustomCommessaField: document.getElementById("calendar-event-custom-commessa-field"),
  calendarEventCustomCommessa: document.getElementById("calendar-event-custom-commessa"),
  calendarEventImpianto: document.getElementById("calendar-event-impianto"),
  calendarEventCustomImpiantoField: document.getElementById("calendar-event-custom-impianto-field"),
  calendarEventCustomImpianto: document.getElementById("calendar-event-custom-impianto"),
  calendarEventLocation: document.getElementById("calendar-event-location"),
  calendarParticipantsPicker: document.getElementById("calendar-participants-picker"),
  calendarParticipantsChips: document.getElementById("calendar-participants-chips"),
  calendarParticipantsSearch: document.getElementById("calendar-participants-search"),
  calendarParticipantsSuggestions: document.getElementById("calendar-participants-suggestions"),
  calendarEventParticipants: document.getElementById("calendar-event-participants"),
  calendarEventDescription: document.getElementById("calendar-event-description"),
  calendarEventLink: document.getElementById("calendar-event-link"),
  calendarEventFormFeedback: document.getElementById("calendar-event-form-feedback"),
  calendarEventCloseBtn: document.getElementById("calendar-event-close-btn"),
  calendarEventCancelBtn: document.getElementById("calendar-event-cancel-btn"),
  calendarEventSaveBtn: document.getElementById("calendar-event-save-btn"),
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
  segnalazioneGeneraPdfBtn: document.getElementById("segnalazione-genera-pdf-btn"),
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
  impiantoAlertModal: document.getElementById("impianto-alert-modal"),
  impiantoAlertTitle: document.getElementById("impianto-alert-title"),
  impiantoAlertBody: document.getElementById("impianto-alert-body"),
  impiantoAlertContinueBtn: document.getElementById("impianto-alert-continue-btn"),
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
  editNoteImpianto: document.getElementById("edit-note-impianto"),
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
  globalImpiantoAddToCommessaBtn: document.getElementById("global-impianto-add-to-commessa-btn"),
  globalImpiantoUsage: document.getElementById("global-impianto-usage"),
  globalAddModal: document.getElementById("global-add-to-commessa-modal"),
  globalAddForm: document.getElementById("global-add-form"),
  globalAddCloseBtn: document.getElementById("global-add-close-btn"),
  globalAddCancelBtn: document.getElementById("global-add-cancel-btn"),
  globalAddCommessaSearch: document.getElementById("global-add-commessa-search"),
  globalAddCommesseOptions: document.getElementById("global-add-commesse-options"),
  globalAddLavorazione: document.getElementById("global-add-lavorazione"),
  globalAddNota: document.getElementById("global-add-nota"),
  globalAddFeedback: document.getElementById("global-add-feedback"),
  globalAddDuplicate: document.getElementById("global-add-duplicate"),
  globalAddSuccess: document.getElementById("global-add-success"),
  globalAddSubmitBtn: document.getElementById("global-add-submit-btn"),
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
let unsubscribeFattoVisualEvidence = null;
const fattoVisualEvidenceByImpianto = new Map();
let unsubscribeCommessaNotes = null;
const unsubscribeCommessaStats = new Map();
let unsubscribeHoursStats = null;
let unsubscribeHoursApprovals = null;
let currentUserPos = null;
const FATTO_POSITION_MAX_AGE_MS = 60 * 1000;
let currentWeatherTarget = { lat: 44.4949, lon: 11.3426 };
let selectedWeatherLocation = null;
let currentHomeWeatherForecast = null;
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
let unsubscribeCalendarEvents = null;
let calendarEvents = [];
let calendarSelectedParticipants = [];
const confirmedSquadraAbsenceAssignments = new Set();
const calendarAbsenceCache = new Map();
let calendarVisibleMonth = new Date(new Date().getFullYear(), new Date().getMonth(), 1);
let calendarSelectedDate = formatCalendarDateKey(new Date());
let calendarMode = "choice";
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
let currentUserBanProfile = null;
let programmazioni = [];
let programmazioneOperatorAutocomplete = null;
let programmazioneMezziAutocomplete = null;
let unsubscribeProgrammazioni = null;
let programmazioniLoadedAt = 0;
let programmazioniLoadPromise = null;
const PROGRAMMAZIONI_CACHE_TTL_MS = 5 * 60 * 1000;
let operatorPositions = [];
let operatorPositionsVisible = true;
let activeUsersLogs = [];
let activeUsersLoaded = false;
let activeUsersFilterOpen = false;
let activeUsersFullListOpen = false;
let activeUsersLogListOpen = false;
let selectedActiveUsersCard = "";
let selectedActiveUsersUserId = "";
let selectedUserActivityLogs = [];
let selectedUserActivityUser = null;
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
let deferredInstallPrompt = null;
let driveTokenRefreshPromise = null;
const commessaSheetCache = new Map();
let commesseById = new Map();
let commesseLoadState = { status: "idle", message: "" };
let commesseLoadTimeout = null;
let isCommesseHomeCardVisible = false;
let commessaStatsById = new Map();
const impiantiByCommessaId = new Map();
let commessaHoursById = new Map();
let commessaWorkSummariesById = new Map();
let allHoursReports = [];
let allHoursApprovalRequests = [];
let hoursReportsLoaded = false;
let hoursApprovalsLoaded = false;
let personaleRecords = [];
let personaleLoadState = { status: "idle", message: "" };
let personaleSearchQuery = "";
let personaleExpandedId = "";
let personaleShowAll = false;
const PERSONALE_RECENT_KEY = "hera_personale_recent_ids";
let mezziRecords = [];
let mezziLoadState = { status: "idle", message: "" };
let startupCoreCollectionsLoadState = { status: "idle", message: "" };
let squadreByCommessa = new Map();
let squadreHistoryByDate = new Map();
let latestSquadraAutofillRequestId = 0;
let squadreLoadState = { status: "idle", message: "" };
let squadreLoadTimeout = null;
const weatherAlertsByDate = new Map();
let weatherAlertsDateLoaded = "";
const worklimateRiskByCommessaId = new Map();
let worklimateRiskCacheLoaded = false;
let worklimateRiskCacheLoading = false;
let selectedWeatherAlertContext = null;
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
const whazzupPhotoFilesByImpianto = new Map();
const whazzupPhotoNotesByImpianto = new Map();
const whazzupPhotoSavedAtByImpianto = new Map();
const WHAZZUP_PHOTO_DB_NAME = "heraWhazzupPhotoAttachments";
const WHAZZUP_PHOTO_STORE_NAME = "attachments";
const WHAZZUP_PHOTO_MAX_AGE_MS = 10 * 60 * 60 * 1000;
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
let fuelStationsLoadPromise = null;
let fuelStationsAbortController = null;
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
let mapAutoFitSignature = "";
let userLocationMapMarker = null;
let userLocationFullscreenMarker = null;
let latestGeolocationCoords = null;
let locationPermission = "prompt";
let isLocationEnabled = false;
let showLocationWarning = false;
let locationWarningMode = "prompt";
let locationClientInfo = { browser: "Altro browser", os: "Altro sistema" };
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
let biogasMapInstance = null;
let biogasLayerGroup = null;
let biogasTileLayer = null;
let biogasLabelTileLayer = null;
let biogasUserMarker = null;
let biogasWatchId = null;
let biogasFeatures = [];
let biogasVisible = true;
let biogasDistanceAlertLevel = "";
let biogasBaseLayerMode = localStorage.getItem("hera_biogas_base_layer") || "standard";
let biogasHighlightedLayer = null;
let biogasPipeRenderLayers = [];
let biogasNearestSnapshot = null;
let biogasMapContentType = "rete_biogas";
const TOMBINI_ENABLED_COMMESSE = [
  "DISCARICA RETI BENTIVOGLIO",
  "DISCARICA RETI FERRARA CA",
  "DISCARICA RETI FERRARA 2B"
];
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
const impiantoWhatsAppTemplateCache = new Map();
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
    icon: "‚òï",
    query: "node[\"amenity\"~\"^(cafe|bar|pub)$\"](around:{radius},{lat},{lng});way[\"amenity\"~\"^(cafe|bar|pub)$\"](around:{radius},{lat},{lng});relation[\"amenity\"~\"^(cafe|bar|pub)$\"](around:{radius},{lat},{lng});",
    detailFields: ["opening_hours", "cuisine", "takeaway", "delivery", "contact:phone", "website", "outdoor_seating"]
  },
  lunch: {
    title: "Pranzo (ristoranti, mense, circoli ARCI)",
    icon: "üçΩÔ∏è",
    query: "node[\"amenity\"~\"^(restaurant|fast_food|food_court|canteen|biergarten|pub)$\"](around:{radius},{lat},{lng});way[\"amenity\"~\"^(restaurant|fast_food|food_court|canteen|biergarten|pub)$\"](around:{radius},{lat},{lng});relation[\"amenity\"~\"^(restaurant|fast_food|food_court|canteen|biergarten|pub)$\"](around:{radius},{lat},{lng});node[\"club\"=\"social\"](around:{radius},{lat},{lng});way[\"club\"=\"social\"](around:{radius},{lat},{lng});relation[\"club\"=\"social\"](around:{radius},{lat},{lng});node[\"social_facility\"=\"canteen\"](around:{radius},{lat},{lng});way[\"social_facility\"=\"canteen\"](around:{radius},{lat},{lng});",
    detailFields: ["cuisine", "opening_hours", "opening_hours:covid19", "payment:meal_voucher", "payment:sodexo", "payment:edenred", "payment:ticket_restaurant", "payment:cash", "payment:credit_cards", "diet:vegetarian", "diet:vegan", "takeaway", "delivery", "contact:phone", "website", "addr:street", "addr:housenumber", "addr:city"]
  },
  supermarket: {
    title: "Supermarket",
    icon: "üõí",
    query: "node[\"shop\"~\"supermarket|convenience\"](around:{radius},{lat},{lng});way[\"shop\"~\"supermarket|convenience\"](around:{radius},{lat},{lng});",
    detailFields: ["opening_hours", "brand", "contact:phone", "website"]
  },
  tobacco: {
    title: "Tabaccherie",
    icon: "üö¨",
    query: "node[\"shop\"=\"tobacco\"](around:{radius},{lat},{lng});way[\"shop\"=\"tobacco\"](around:{radius},{lat},{lng});",
    detailFields: ["opening_hours", "contact:phone", "website"]
  },
  wc: {
    title: "WC pubblici",
    icon: "üöª",
    query: "node[\"amenity\"=\"toilets\"](around:{radius},{lat},{lng});way[\"amenity\"=\"toilets\"](around:{radius},{lat},{lng});",
    detailFields: ["fee", "wheelchair", "opening_hours"]
  },
  atm: {
    title: "Bancomat / ATM",
    icon: "üèß",
    query: "node[\"amenity\"=\"atm\"](around:{radius},{lat},{lng});way[\"amenity\"=\"atm\"](around:{radius},{lat},{lng});",
    detailFields: ["opening_hours", "operator", "cash_in", "contactless", "currency:EUR"]
  },
  pharmacy: {
    title: "Farmacie",
    icon: "üíä",
    query: "node[\"amenity\"=\"pharmacy\"](around:{radius},{lat},{lng});way[\"amenity\"=\"pharmacy\"](around:{radius},{lat},{lng});",
    detailFields: ["opening_hours", "dispensing", "contact:phone", "website"]
  },
  parking: {
    title: "Parcheggi",
    icon: "üÖøÔ∏è",
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
      "Apri il menu (‚ãÆ) e premi ‚ÄúGestione commesse‚Äù.",
      "Inserisci il nome commessa oppure seleziona una commessa per import/aggiunte impianto.",
      "Usa i form della pagina per completare l'operazione."
    ],
    tags: ["commesse", "impianti", "excel", "admin"]
  },
  "open-panel-squadre": {
    rispostaBreve: "Serve per creare e salvare la composizione giornaliera delle squadre.",
    passi: [
      "Apri il menu (‚ãÆ) e premi ‚ÄúComposizione squadre‚Äù.",
      "Scegli commessa e data, poi aggiungi le righe squadra.",
      "Premi ‚ÄúSalva composizione‚Äù e verifica lo storico sotto al form."
    ],
    tags: ["squadre", "operativo", "personale", "mezzi"]
  },

  "open-panel-personale": {
    rispostaBreve: "Da qui inserisci o importi l'anagrafica personale.",
    passi: [
      "Apri il menu (‚ãÆ) e premi ‚ÄúPersonale‚Äù.",
      "Aggiungi un nominativo singolo o importa il file Excel.",
      "Controlla la lista aggiornata subito sotto."
    ],
    tags: ["personale", "anagrafica", "excel"]
  },
  "open-panel-mezzi": {
    rispostaBreve: "Da qui inserisci o importi l'elenco mezzi disponibili.",
    passi: [
      "Apri il menu (‚ãÆ) e premi ‚ÄúMezzi‚Äù.",
      "Aggiungi un mezzo manualmente o importa da Excel.",
      "Controlla che il mezzo compaia in elenco."
    ],
    tags: ["mezzi", "flotta", "excel"]
  },
  "open-panel-utenti": {
    rispostaBreve: "Permette la gestione admin, permessi utente e app collegate.",
    passi: [
      "Apri il menu (‚ãÆ) e premi ‚ÄúGestione utenti‚Äù.",
      "Aggiungi/rimuovi admin oppure aggiorna i permessi azione per utente.",
      "Salva le modifiche e verifica l'elenco utenti."
    ],
    tags: ["utenti", "permessi", "admin"]
  },
  "open-panel-info-utili": {
    rispostaBreve: "Consente di pubblicare risorse utili (contatti, note, documenti) per commessa.",
    passi: [
      "Apri il menu (‚ãÆ) e premi ‚ÄúInformazioni utili‚Äù.",
      "Seleziona tipo risorsa, titolo e contenuto/link.",
      "Salva e verifica che la risorsa sia disponibile nella commessa."
    ],
    tags: ["risorse", "contatti", "note", "documenti"]
  },
  "open-private-docs-btn": {
    rispostaBreve: "Area personale per caricare e consultare documenti individuali.",
    passi: [
      "Apri il menu (‚ãÆ) e premi ‚ÄúDocumenti‚Äù.",
      "Compila nome/note e allega file o foto.",
      "Salva e verifica la presenza del documento nell'elenco."
    ],
    tags: ["documenti", "personale", "drive"]
  },
  "open-personal-services-btn": {
    rispostaBreve: "Trovi servizi vicini (colazione, pranzo, market, tabacchi, WC, bancomat e altri) con mappa e navigazione.",
    passi: [
      "Apri il menu (‚ãÆ) e premi ‚ÄúServizi personali‚Äù.",
      "Scegli una categoria (es. Colazione o Pranzo).",
      "Apri un luogo dall'elenco o dalla mappa e usa ‚ÄúNaviga‚Äù o ‚ÄúDettagli‚Äù."
    ],
    tags: ["servizi", "mappa", "navigazione", "personale"]
  },
  "open-hours-btn": {
    rispostaBreve: "Compili ore per commessa e operatore, salvi il resoconto e invii WhatsApp.",
    passi: [
      "Apri il menu (‚ãÆ) e premi ‚ÄúGestione ore‚Äù.",
      "Aggiungi una o pi√π commesse, poi operatori con ore e note.",
      "Premi ‚ÄúFine: salva e invia‚Äù per salvare su Drive e aprire WhatsApp."
    ],
    tags: ["ore", "commesse", "operatori", "whatsapp", "drive"]
  },
  "open-pos-btn": {
    rispostaBreve: "Archivio documenti sicurezza: POS, PMS, schede lavorazioni, schede macchine e modulistica.",
    passi: [
      "Apri il menu (‚ãÆ) e premi ‚ÄúüìÑ POS‚Äù.",
      "Cerca il documento per titolo, descrizione o categoria.",
      "Premi ‚ÄúApri documento‚Äù per consultare il link Google Drive in una nuova scheda."
    ],
    tags: ["pos", "documenti", "sicurezza", "drive"]
  },
  "open-segnalazioni-btn": {
    rispostaBreve: "Compili la segnalazione sicurezza e generi il PDF da condividere.",
    passi: [
      "Apri il menu (‚ãÆ) e premi ‚ÄúSegnalazioni‚Äù.",
      "Compila i campi obbligatori e scegli il tipo di segnalazione.",
      "Genera il PDF e condividilo via WhatsApp o email."
    ],
    tags: ["segnalazioni", "pdf", "sicurezza"]
  },
  "open-book-pdf-btn": {
    rispostaBreve: "Apre il manuale completo dell'app in formato PDF in una nuova scheda.",
    passi: [
      "Apri il menu (‚ãÆ) e premi ‚ÄúLibro PDF‚Äù.",
      "Attendi l'apertura del file PDF in una nuova scheda del browser.",
      "Se il popup √® bloccato, abilita i popup oppure scarica il file dal link diretto."
    ],
    tags: ["manuale", "pdf", "guida"]
  }
};
const STATIC_HOWTO_ITEMS = [
  {
    id: "login-google",
    domanda: "Come faccio il login con Google?",
    rispostaBreve: "Apri il pannello utente e premi ‚ÄúLogin con Google‚Äù.",
    passi: [
      "Nella home premi l'icona üë§ in alto.",
      "Tocca ‚ÄúLogin con Google‚Äù e scegli l'account aziendale.",
      "Controlla che compaia ‚ÄúLoggato‚Äù con email e nome utente."
    ],
    tags: ["login", "google", "accesso"],
    updatedAt: HOWTO_UPDATED_AT
  },
  {
    id: "chat-operatori",
    domanda: "Come uso la chat operatori?",
    rispostaBreve: "Apri la chat dal pulsante üí¨, scrivi e invia il messaggio al destinatario.",
    passi: [
      "Premi il pulsante üí¨ in basso a destra.",
      "Scegli un destinatario o lascia ‚ÄúMessaggio per tutti‚Äù.",
      "Scrivi il testo (o allega media/vocale) e premi invio."
    ],
    tags: ["chat", "messaggi", "operatori"],
    updatedAt: HOWTO_UPDATED_AT
  },
  {
    id: "google-drive",
    domanda: "Come collego Google Drive?",
    rispostaBreve: "Solo l‚Äôamministratore collega Google Drive; per gli utenti il cloud √® automatico.",
    passi: [
      "Esegui login con Google con un account autorizzato.",
      "Solo l‚Äôamministratore apre il pannello utente e preme ‚ÄúCollega Google Drive‚Äù.",
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
const BUILT_IN_SUPER_ADMIN_EMAILS = [ADMIN_EMAIL, "Ionut29019@gmail.com"];
const POS_DEFAULT_CATEGORIES = ["POS", "PMS", "Schede lavorazioni", "Schede macchine e attrezzature", "Sicurezza", "Modulistica", "Altro"];
const IMPIANTO_ACTIONS = ["done", "navigate", "reset", "whatsapp", "problem-report", "gps-update", "edit", "delete"];
const ADMIN_ONLY_IMPIANTO_ACTIONS = ["reset", "edit", "delete"];
let adminEmails = new Set(BUILT_IN_SUPER_ADMIN_EMAILS.map((email) => normalizeEmail(email)));
let posDocuments = [];
let unsubscribePosDocuments = null;
const PENDING_SHEET_EXPORTS_KEY = "heraPendingSheetExports";
const PENDING_IMPIANTO_ACTIONS_KEY = "heraPendingImpiantoActions";
const PENDING_OFFLINE_MUTATIONS_KEY = "heraPendingOfflineMutations";
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
    },
    {
      id: "faq-install-app",
      domanda: "Come installare l‚Äôapp sul telefono?",
      risposta: "Dal menu laterale premi Installa app. Su iPhone o Safari usa il menu Condividi del browser e scegli Aggiungi alla schermata Home.",
      passi: ["Apri il menu laterale", "Premi Installa app", "Se compare il prompt, conferma l‚Äôinstallazione", "Su iPhone/Safari: Apri Condividi ‚Üí Aggiungi alla schermata Home"]
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
const snowRoadLayer = L.layerGroup().addTo(map);
const fullscreenSnowRoadLayer = L.layerGroup().addTo(fullscreenMap);
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

const UserLocationControl = L.Control.extend({
  options: { position: "topright" },
  onAdd(targetMap) {
    const button = L.DomUtil.create("button", "map-geolocate-btn");
    button.type = "button";
    button.title = "Vai alla mia posizione";
    button.setAttribute("aria-label", "Centra sulla mia posizione");
    button.innerHTML = "üìç";
    L.DomEvent.disableClickPropagation(button);
    L.DomEvent.on(button, "click", (event) => {
      L.DomEvent.stop(event);
      centerMapOnUserLocation(targetMap);
    });
    return button;
  }
});
map.addControl(new UserLocationControl());
fullscreenMap.addControl(new UserLocationControl());

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

ui.loginBtn?.addEventListener("click", loginWithGoogle);
ui.authGateLoginBtn?.addEventListener("click", loginWithGoogle);
ui.authEmailForm?.addEventListener("submit", (event) => {
  event.preventDefault();
  void loginWithEmailPassword();
});
ui.switchAccountBtn?.addEventListener("click", switchGoogleAccount);
ui.refreshAppBtn?.addEventListener("click", refreshApplicationData);
ui.updateAppBtn?.addEventListener("click", openApplicationUpdate);
ui.menuToggleBtn?.addEventListener("click", openSideMenu);
ui.menuCloseBtn?.addEventListener("click", closeSideMenu);
ui.installAppBtn?.addEventListener("click", handleInstallAppClick);
ui.menuOverlay?.addEventListener("click", closeSideMenu);
ui.logoutBtn?.addEventListener("click", logout);
ui.biometricEnableBtn?.addEventListener("click", () => void enableBiometricAccess());
ui.biometricToggle?.addEventListener("change", () => void (ui.biometricToggle.checked ? enableBiometricAccess() : disableBiometricAccess()));
ui.biometricVerifyBtn?.addEventListener("click", () => void verifyBiometricFromSettings());
ui.biometricDisableBtn?.addEventListener("click", () => void disableBiometricAccess());
ui.driveConnectBtn?.addEventListener("click", connectGoogleDrive);
ui.commessaForm?.addEventListener("submit", createCommessa);
document.getElementById("open-new-commessa-btn")?.addEventListener("click", openNewCommessaDialog);
document.getElementById("close-new-commessa-btn")?.addEventListener("click", closeNewCommessaDialog);
document.getElementById("cancel-new-commessa-btn")?.addEventListener("click", closeNewCommessaDialog);
document.getElementById("commesse-management-search")?.addEventListener("input", renderCommesseManagementList);
document.getElementById("close-impianti-management-btn")?.addEventListener("click", closeImpiantiManagement);
document.getElementById("download-excel-template-btn")?.addEventListener("click", downloadOfficialImpiantiTemplate);
document.getElementById("export-all-impianti-btn")?.addEventListener("click", exportAllImpiantiStatus);
document.getElementById("add-management-impianto-btn")?.addEventListener("click", addManagementImpianto);
document.getElementById("toggle-import-impianti-btn")?.addEventListener("click", () => document.getElementById("impianti-import-card")?.classList.toggle("hidden"));
document.getElementById("clear-impianti-management-search")?.addEventListener("click", () => { const input = document.getElementById("impianti-management-search"); if (input) input.value = ""; managementPage = 1; renderImpiantiManagementTable(); });
document.getElementById("impianti-management-search")?.addEventListener("input", () => { clearTimeout(managementSearchTimer); managementSearchTimer = setTimeout(() => { managementPage = 1; renderImpiantiManagementTable(); }, 250); });
["impianti-status-filter", "impianti-comune-filter", "impianti-tipologia-filter", "impianti-operatore-filter"].forEach((id) => document.getElementById(id)?.addEventListener("change", () => { managementPage = 1; renderImpiantiManagementTable(); }));
document.getElementById("impianti-management-thead")?.addEventListener("click", (event) => { const header = event.target.closest("[data-sort]"); if (header?.dataset.sort) { managementSort = managementSort.field === header.dataset.sort ? { field: header.dataset.sort, direction: -managementSort.direction } : { field: header.dataset.sort, direction: 1 }; renderImpiantiManagementTable(); } });
document.getElementById("impianti-management-thead")?.addEventListener("change", (event) => { if (!event.target.matches("[data-select-all]")) return; document.querySelectorAll("#impianti-management-tbody [data-plant-id]").forEach((row) => event.target.checked ? managementSelectedIds.add(row.dataset.plantId) : managementSelectedIds.delete(row.dataset.plantId)); renderImpiantiManagementTable(); });
document.getElementById("impianti-management-tbody")?.addEventListener("change", (event) => { if (!event.target.matches("[data-select-row]")) return; const id = event.target.closest("[data-plant-id]")?.dataset.plantId; if (event.target.checked) managementSelectedIds.add(id); else managementSelectedIds.delete(id); updateManagementBulkBar(); });
document.getElementById("impianti-management-tbody")?.addEventListener("click", (event) => { const button = event.target.closest("[data-row-action]"); if (!button) return; const row = button.closest("[data-plant-id]"); if (button.dataset.rowAction === "edit") { managementEditingId = row.dataset.plantId; renderImpiantiManagementTable(); } else if (button.dataset.rowAction === "cancel") { managementEditingId = ""; renderImpiantiManagementTable(); } else if (button.dataset.rowAction === "save") void saveManagementPlantRow(row); });
document.getElementById("impianti-management-pagination")?.addEventListener("click", (event) => { const button = event.target.closest("[data-page]"); if (button && !button.disabled) { managementPage = Number(button.dataset.page); renderImpiantiManagementTable(); } });
document.getElementById("impianti-bulk-actions")?.addEventListener("click", (event) => { const button = event.target.closest("[data-bulk]"); if (button) void runManagementBulkAction(button.dataset.bulk); });
document.getElementById("snow-roads-form")?.addEventListener("submit", addSnowRoadsToSelectedCommessa);
ui.commessaType?.addEventListener("change", updateCommessaParentField);
ui.openOrganizeCommesseBtn?.addEventListener("click", () => toggleOrganizeCommesseScreen(true));
ui.closeOrganizeCommesseBtn?.addEventListener("click", () => toggleOrganizeCommesseScreen(false));
ui.parentCommessaForm?.addEventListener("submit", createParentCommessa);
ui.moveSubcommesseForm?.addEventListener("submit", moveSelectedCommesseUnderParent);
ui.moveParentCommessaSelect?.addEventListener("change", renderMoveSubcommesseList);
ui.excelFile?.addEventListener("change", onExcelSelected);
ui.importBtn?.addEventListener("click", importPendingRows);
ui.sheetUrlImportBtn?.addEventListener("click", importFromGoogleSheetUrl);
ui.commessaTargetSelect?.addEventListener("change", onCommessaTargetChanged);
ui.chatOpenBtn?.addEventListener("click", openChatModal);
ui.chatCloseBtn?.addEventListener("click", closeChatModal);
ui.chatClearBtn?.addEventListener("click", openChatClearConfirmModal);
ui.chatClearCancelBtn?.addEventListener("click", closeChatClearConfirmModal);
ui.chatClearConfirmBtn?.addEventListener("click", clearCurrentChatMessages);
ui.chatClearConfirmModal?.addEventListener("click", (event) => {
  if (event.target === ui.chatClearConfirmModal) closeChatClearConfirmModal();
});
ui.chatSendForm?.addEventListener("submit", sendTextMessage);
ui.chatMediaInput?.addEventListener("change", sendMediaMessage);
ui.chatVoiceBtn?.addEventListener("click", toggleVoiceRecording);
ui.backToHomeBtn?.addEventListener("click", closeImpiantiPage);
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
ui.exportCurrentCommessaBtn?.addEventListener("click", () => exportCommessaSummary(selectedCommessaId, selectedCommessaName));
ui.mapFullscreenBtn?.addEventListener("click", openMapFullscreenPage);
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
ui.biogasMapBackBtn?.addEventListener("click", closeBiogasMapPage);
ui.biogasMapSettingsBtn?.addEventListener("click", toggleBiogasMapControls);
ui.biogasMapLayerBtn?.addEventListener("click", toggleBiogasBaseLayerMode);
ui.biogasMapToggleBtn?.addEventListener("click", toggleBiogasNetworkVisibility);
ui.biogasMapSearch?.addEventListener("input", onBiogasSearchInput);
ui.biogasMapPage?.addEventListener("click", (event) => {
  const btn = event.target?.closest?.("[data-tombino-code]");
  if (!btn) return;
  toggleTombinoStatus(btn.dataset.tombinoCode).catch((error) => {
    console.error("Errore aggiornamento stato pozzetto:", error);
    if (ui.biogasMapStatus) ui.biogasMapStatus.textContent = "Errore aggiornamento stato pozzetto. Riprova.";
  });
});
ui.biogasMapRefreshBtn?.addEventListener("click", () => loadBiogasNetworkForCurrentCommessa({ forceRefresh: true, type: biogasMapContentType }));
ui.biogasMapDeleteBtn?.addEventListener("click", deleteBiogasNetworkForCurrentCommessa);
ui.biogasMapAddPipesBtn?.addEventListener("click", () => ui.biogasMapFileInput?.click());
ui.biogasMapFileInput?.addEventListener("change", onBiogasFileSelected);
ui.biogasDistanceIndicator?.addEventListener("click", () => {
  if (!biogasMapInstance || !biogasNearestSnapshot || !Number.isFinite(biogasNearestSnapshot.dist)) return;
  const cfg = getBiogasMapConfig();
  const nearest = biogasNearestSnapshot.nearest || [0, 0];
  const html = `<div><b>Codice ${escapeHTML(cfg.itemLabel)}:</b> ${escapeHTML(biogasNearestSnapshot.code)}<br><b>Distanza:</b> ${Math.round(biogasNearestSnapshot.dist)} m<br><b>Coordinate:</b> ${nearest[0].toFixed(6)}, ${nearest[1].toFixed(6)}<br><button type="button" id="biogas-center-map-btn" class="btn" style="margin-top:6px">Centra sulla mappa</button></div>`;
  L.popup().setLatLng(biogasUserMarker?.getLatLng?.() || biogasMapInstance.getCenter()).setContent(html).openOn(biogasMapInstance);
});

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
document.getElementById("map-save-snow-road-btn")?.addEventListener("click", saveDrawnSnowRoadPath);
ui.mapFullscreenFeedbackClose?.addEventListener("click", () => ui.mapFullscreenFeedbackBanner?.classList.add("hidden"));
ui.mapEnableLocationBtn?.addEventListener("click", () => { void requestLocationEnableFlow(); });
ui.mapRetryLocationBtn?.addEventListener("click", () => { void requestLocationEnableFlow({ forceRetry: true }); });
ui.mapLocationHelpBtn?.addEventListener("click", () => toggleLocationHelpPanel());
ui.toggleCommesseHomeBtn?.addEventListener("click", toggleCommesseHomeCard);
ui.impiantoSearch?.addEventListener("input", onImpiantoSearchInput);
ui.viewDoneBtn?.addEventListener("click", () => setImpiantiViewMode("done"));
ui.viewTodoBtn?.addEventListener("click", () => setImpiantiViewMode("todo"));
ui.viewAlertsBtn?.addEventListener("click", () => setImpiantiViewMode("alerts"));
document.querySelectorAll(".commessa-stat-item[data-stat-action]").forEach((item) => {
  item.addEventListener("click", () => handleCommessaStatAction(item.dataset.statAction || ""));
  item.addEventListener("keydown", (event) => {
    if (event.key !== "Enter" && event.key !== " ") return;
    event.preventDefault();
    handleCommessaStatAction(item.dataset.statAction || "");
  });
});
ui.personaleForm?.addEventListener("submit", addPersonale);
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
ui.mezziForm?.addEventListener("submit", addMezzo);
ui.squadraForm?.addEventListener("submit", saveSquadraComposition);
ui.squadraCommessa?.addEventListener("change", autofillSquadraForm);
ui.squadraRiferimento?.addEventListener("change", autofillSquadraForm);
ui.squadraCalendarDate?.addEventListener("change", () => {
  setSquadreDateOverride(ui.squadraCalendarDate.value || "");
});
ui.squadreFilterDate?.addEventListener("change", onSquadreFilterDateChange);
ui.weatherAlertSafetyBackBtn?.addEventListener("click", () => setCommessaHash());
ui.weatherAlertSafetyConfirmBtn?.addEventListener("click", confirmWeatherAlertRead);
ui.squadreFilterClearBtn?.addEventListener("click", () => clearManualSquadreFilterDate());
ui.snowSquadreFilterDate?.addEventListener("change", onSnowSquadreFilterDateChange);
ui.snowSquadreFilterClearBtn?.addEventListener("click", () => clearManualSquadreFilterDate({ snow: true }));

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
ui.addSquadraRowBtn?.addEventListener("click", () => addSquadraRow());
ui.personaleImportBtn?.addEventListener("click", importPersonaleFromExcel);
ui.mezziImportBtn?.addEventListener("click", importMezziFromExcel);
ui.openPanelCommesse?.addEventListener("click", () => openManagementPanel("commesse"));
ui.openPanelSquadre?.addEventListener("click", () => openManagementPanel("squadre"));
ui.openPanelPersonale?.addEventListener("click", () => openManagementPanel("personale"));
ui.openPanelMezzi?.addEventListener("click", () => openManagementPanel("mezzi"));
ui.openPanelUtenti?.addEventListener("click", () => openManagementPanel("utenti"));
ui.openPanelGlobal?.addEventListener("click", () => openManagementPanel("global"));
ui.openPanelBanner?.addEventListener("click", () => openManagementPanel("banner"));
ui.openPanelInfoUtili?.addEventListener("click", () => openManagementPanel("infoUtili"));
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
ui.ferieForm?.addEventListener("submit", saveFerieCollega);
ui.ferieCheckBtn?.addEventListener("click", renderFerieDisponibilitaCalendar);
ui.openPanelBannerGestione?.addEventListener("click", () => openManagementPanel("banner"));
ui.openPrivateDocsBtn?.addEventListener("click", openPrivateDocsPage);
ui.homeCalendarBtn?.addEventListener("click", openCalendarPage);
ui.openPrivateDocsUploadBtn?.addEventListener("click", openPrivateDocsUploadPage);
ui.openPersonalServicesBtn?.addEventListener("click", openPersonalServicesPage);
ui.openHoursBtn?.addEventListener("click", openHoursPage);
ui.openPosBtn?.addEventListener("click", openPosPage);
ui.openSegnalazioniBtn?.addEventListener("click", openSegnalazioniPage);
ui.homeSegnalazioniBtn?.addEventListener("click", openSegnalazioniPage);
ui.openHowtoBtn?.addEventListener("click", openHowtoPage);
ui.openControlCenterBtn?.addEventListener("click", openControlCenterPage);
ui.runControlCheckBtn?.addEventListener("click", runControlCenterCheck);
ui.backFromControlCenterBtn?.addEventListener("click", closeControlCenterPage);
ui.openBookPdfBtn?.addEventListener("click", openBookPdf);
ui.managementCloseBtn?.addEventListener("click", closeManagementPanel);
ui.userToggleBtn?.addEventListener("click", toggleUserDetailsPanel);
ui.userConnectionBar?.addEventListener("click", toggleProfileSummary);
ui.userConnectionBar?.addEventListener("keydown", (event) => {
  if (event.key !== "Enter" && event.key !== " ") return;
  event.preventDefault();
  toggleProfileSummary();
});
ui.weatherCloseBtn?.addEventListener("click", closeWeatherModal);
ui.weatherCard?.addEventListener("click", toggleWeatherCard);
ui.weatherCard?.addEventListener("keydown", (event) => {
  if (event.key === "Enter" || event.key === " ") {
    event.preventDefault();
    toggleWeatherCard();
  }
});
ui.weatherExternalDetailBtn?.addEventListener("click", (event) => {
  event.stopPropagation();
  openWeatherExternalDetail();
});
ui.todayCommesseBtn?.addEventListener("click", () => ui.todaySquadsSection?.scrollIntoView({ behavior: "smooth", block: "start" }));
ui.todayHoursBtn?.addEventListener("click", openHoursPage);
ui.todayMezziBtn?.addEventListener("click", () => ui.todaySquadsSection?.scrollIntoView({ behavior: "smooth", block: "start" }));
ui.todayAlertsBtn?.addEventListener("click", () => ui.todaySquadsSection?.scrollIntoView({ behavior: "smooth", block: "start" }));
ui.backFromFuelBtn?.addEventListener("click", closeFuelPage);
ui.fuelMezzoDetailsBtn?.addEventListener("click", toggleFuelMezzoDetails);
ui.fuelSearchBtn?.addEventListener("click", () => loadNearbyFuelStations({ force: true }));
ui.fuelRadius?.addEventListener("change", () => {
  if (!ui.fuelPage?.classList.contains("hidden")) loadNearbyFuelStations({ force: true });
});
ui.backFromPersonalServicesBtn?.addEventListener("click", closePersonalServicesPage);
ui.backFromSegnalazioniBtn?.addEventListener("click", closeSegnalazioniPage);
ui.backFromHowtoBtn?.addEventListener("click", closeHowtoPage);
ui.backFromPrivateDocsBtn?.addEventListener("click", closePrivateDocsPage);
ui.backFromCalendarBtn?.addEventListener("click", closeCalendarPage);
ui.calendarChoiceBackBtn?.addEventListener("click", closeCalendarPage);
ui.calendarChoiceHoursBtn?.addEventListener("click", () => setCalendarMode("hours"));
ui.calendarChoiceSharedBtn?.addEventListener("click", () => setCalendarMode("shared"));
ui.calendarHoursTab?.addEventListener("click", () => setCalendarMode("hours"));
ui.calendarSharedTab?.addEventListener("click", () => setCalendarMode("shared"));
ui.calendarNewEventBtn?.addEventListener("click", () => openCalendarEventForm(calendarSelectedDate));
ui.calendarAddSelectedDayBtn?.addEventListener("click", () => openCalendarEventForm(calendarSelectedDate));
ui.calendarPrevBtn?.addEventListener("click", () => changeCalendarMonth(-1));
ui.calendarTodayBtn?.addEventListener("click", showCalendarToday);
ui.calendarNextBtn?.addEventListener("click", () => changeCalendarMonth(1));
ui.calendarEventAllDay?.addEventListener("change", syncCalendarTimeFields);
ui.calendarEventCommessa?.addEventListener("change", handleCalendarCommessaChange);
ui.calendarEventImpianto?.addEventListener("change", handleCalendarImpiantoChange);
ui.calendarParticipantsSearch?.addEventListener("input", renderCalendarParticipantSuggestions);
ui.calendarParticipantsSearch?.addEventListener("focus", renderCalendarParticipantSuggestions);
ui.calendarParticipantsSearch?.addEventListener("keydown", handleCalendarParticipantSearchKeydown);
ui.calendarEventCloseBtn?.addEventListener("click", closeCalendarEventForm);
ui.calendarEventCancelBtn?.addEventListener("click", closeCalendarEventForm);
ui.calendarEventForm?.addEventListener("submit", saveCalendarEvent);
ui.backFromHoursBtn?.addEventListener("click", closeHoursPage);
ui.backFromPosBtn?.addEventListener("click", closePosPage);
ui.posAddToggleBtn?.addEventListener("click", () => openPosDocumentForm());
ui.posCancelBtn?.addEventListener("click", closePosDocumentForm);
ui.posDocumentForm?.addEventListener("submit", savePosDocument);
ui.posSearch?.addEventListener("input", renderPosDocuments);
ui.hoursForm?.addEventListener("submit", finalizeHoursReport);
ui.addHoursCommessaBtn?.addEventListener("click", () => {
  unlockHoursFinalizeButton();
  addHoursCommessaBlock();
});
ui.hoursDate?.addEventListener("input", () => {
  unlockHoursFinalizeButton();
  Array.from(ui.hoursCommesseList?.querySelectorAll(".hours-commessa-card") || []).forEach((card) => {
    applyHoursSuggestedOperators(card, { force: true });
  });
});
ui.viewHoursBtn?.addEventListener("click", openHoursViewModal);
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
ui.privateDocsPresetPinBtn?.addEventListener("click", () => applyPrivateDocPreset("pin"));
ui.privateDocsPresetTesseraBtn?.addEventListener("click", () => applyPrivateDocPreset("tessera"));
ui.privateDocsForm?.addEventListener("submit", savePrivateDocument);
ui.personalServicesCategories?.addEventListener("click", onPersonalServiceCategoryClick);
ui.personalServicesRadius?.addEventListener("change", () => {
  if (activePersonalServiceCategory) loadPersonalServicesByCategory(activePersonalServiceCategory);
});
ui.segnalazioneForm?.addEventListener("submit", generateSegnalazionePdf);
ui.segnalazionePreposto?.addEventListener("input", syncSegnalazioneFirmaPreposto);
ui.segnalazioneShareWhatsappBtn?.addEventListener("click", () => shareSegnalazione("whatsapp"));
ui.segnalazioneShareEmailBtn?.addEventListener("click", () => shareSegnalazione("email"));
ui.manualImpiantoForm?.addEventListener("submit", addManualImpianto);
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
ui.globalImpiantoAddToCommessaBtn?.addEventListener("click", openGlobalAddToCommessaModal);
ui.globalAddCloseBtn?.addEventListener("click", closeGlobalAddToCommessaModal);
ui.globalAddCancelBtn?.addEventListener("click", closeGlobalAddToCommessaModal);
ui.globalAddForm?.addEventListener("submit", onGlobalAddToCommessaSubmit);
ui.globalCommesseLista?.addEventListener("click", onGlobalCommesseListClick);
ui.globalOpenReportBtn?.addEventListener("click", () => handleOpenGlobalSegnalazioneClick());
ui.globalImpiantoWhatsappBtn?.addEventListener("click", () => handleOpenGlobalSegnalazioneClick());
ui.globalReportCloseBtn?.addEventListener("click", closeGlobalSegnalazioneModal);
ui.globalReportForm?.addEventListener("submit", submitGlobalSegnalazioneWhatsapp);
ui.globalReportImpiantoSelect?.addEventListener("change", onGlobalSegnalazioneImpiantoChange);
ui.globalReportModal?.addEventListener("click", (event) => {
  if (event.target === ui.globalReportModal) closeGlobalSegnalazioneModal();
});
ui.adminUserForm?.addEventListener("submit", addAdminUserByEmail);
ui.externalAppForm?.addEventListener("submit", saveExternalAppForCurrentUser);
ui.resourceForm?.addEventListener("submit", addResourceItem);
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
ui.resourceType?.addEventListener("change", updateResourceFormByType);
ui.impiantoEditCloseBtn?.addEventListener("click", closeImpiantoEditor);
ui.impiantoEditForm?.addEventListener("submit", saveImpiantoEdits);
ui.impiantoReportCloseBtn?.addEventListener("click", closeImpiantoReportModal);
ui.impiantoReportForm?.addEventListener("submit", submitImpiantoReport);
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
window.addEventListener("online", () => {
  setFirestoreConnectionState("Online", "");
  syncPendingImpiantoActions();
  syncPendingOfflineMutations();
});
window.addEventListener("offline", () => { setFirestoreConnectionState("Offline", "Dati caricati da cache"); });
window.addEventListener("hera:native-location", (event) => {
  const nativePosition = event?.detail;
  if (!nativePosition?.coords) return;
  latestGeolocationCoords = nativePosition.coords;
  updateCurrentUserPosition(nativePosition.coords, nativePosition.timestamp);
  if (ui.gpsStatus) ui.gpsStatus.textContent = "Posizione Android aggiornata.";
});
ui.commessaResourceViewerCloseBtn?.addEventListener("click", closeCommessaResourceViewer);
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
if (window.location.hash) {
  window.location.hash = "";
}
applyRoute();
window.addEventListener("hashchange", applyRoute);
window.addEventListener("popstate", applyRoute);
loadPendingSheetExports();
startSheetRetryLoop();
renderFaqHelpCenter(HELP_CENTER_FAQ_FALLBACK);
renderResourceManageFilters();
updateResourceFormByType();
initPwaCapabilities();
initNativeGeofenceBridge();
initWorkBannerObservers();

function hideStartupLoading() {
  document.getElementById("app-startup-loading")?.classList.add("hidden");
}

function setAuthenticationGateState(state, message = "") {
  const isChecking = state === "checking";
  const isRequired = state === "required";
  const isAuthenticated = state === "authenticated";
  const isBanned = state === "banned";

  document.body.classList.toggle("auth-pending", isChecking);
  document.body.classList.toggle("auth-required", isRequired || isBanned);
  document.body.classList.toggle("auth-banned", isBanned);
  ui.authGate?.classList.toggle("hidden", !(isRequired || isBanned));
  if (ui.authGateMessage) {
    ui.authGateMessage.textContent = isBanned
      ? "Ti √® stato negato l‚Äôaccesso. Richiedi l‚Äôaccesso all‚Äôamministratore."
      : (message || "Accedi con il tuo account Google per utilizzare l'app.");
  }
  if (ui.authGateLoginBtn) {
    ui.authGateLoginBtn.disabled = isChecking;
    ui.authGateLoginBtn.classList.toggle("hidden", isBanned);
  }
  ui.authEmailForm?.classList.toggle("hidden", isBanned);
  ui.bannedRequestAccessBtn?.classList.toggle("hidden", !isBanned);

  if (isRequired || isBanned) {
    ui.sideMenu?.classList.add("hidden");
    ui.sideMenu?.setAttribute("aria-hidden", "true");
    ui.menuOverlay?.classList.add("hidden");
  }
  if (isAuthenticated) {
    ui.authGate?.classList.add("hidden");
    document.body.classList.remove("auth-banned");
  }
}

function runAfterFirstRender(callback) {
  const runner = () => {
    if (typeof window.requestIdleCallback === "function") {
      window.requestIdleCallback(callback, { timeout: 2500 });
    } else {
      setTimeout(callback, 0);
    }
  };
  window.requestAnimationFrame(() => setTimeout(runner, 0));
}

function runDeferredStartupTasks(tasks = []) {
  runAfterFirstRender(() => {
    tasks.forEach((task, index) => {
      setTimeout(() => {
        try {
          task();
        } catch (error) {
          console.error("Errore task avvio differito:", error);
        }
      }, index * 75);
    });
  });
}

function withTimeout(promise, ms, timeoutMessage) {
  let timeoutId = null;
  const timeoutPromise = new Promise((_, reject) => {
    timeoutId = setTimeout(() => reject(new Error(timeoutMessage || "Timeout caricamento dati")), ms);
  });
  return Promise.race([promise, timeoutPromise]).finally(() => clearTimeout(timeoutId));
}

const FIRESTORE_QUERY_TIMEOUT_MS = 9000;
const FIRESTORE_RETRY_DELAYS_MS = [700, 1600];
let firestoreNetworkState = navigator.onLine ? "Online" : "Offline";
let firestoreCacheState = "";
let firestoreSlowTimer = null;
let lastConnectionMbps = null;
updateConnectivityStatus();

function sleep(ms) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

function setFirestoreConnectionState(state, detail = "") {
  firestoreNetworkState = state || firestoreNetworkState;
  firestoreCacheState = detail || firestoreCacheState;
  console.log(`FIRESTORE ${String(firestoreNetworkState).toUpperCase()}`, detail || "");
  updateConnectivityStatus();
}

const browserConnection = navigator.connection || navigator.mozConnection || navigator.webkitConnection;
if (browserConnection && typeof browserConnection.addEventListener === "function") {
  browserConnection.addEventListener("change", updateConnectivityStatus);
}

function startFirestoreSlowWatch(label) {
  clearFirestoreSlowWatch();
  firestoreSlowTimer = setTimeout(() => {
    setFirestoreConnectionState("Connessione lenta", label || "Firestore non risponde rapidamente");
  }, 4000);
}

function clearFirestoreSlowWatch() {
  if (!firestoreSlowTimer) return;
  clearTimeout(firestoreSlowTimer);
  firestoreSlowTimer = null;
}

function getReadableFirestoreError(error, fallback = "Errore caricamento dati") {
  const code = getFirebaseErrorCode(error);
  if (code.includes("permission-denied")) {
    return "Permesso negato: il tuo utente non √® autorizzato a leggere questi dati. Contatta un amministratore.";
  }
  if (code.includes("unavailable")) {
    return "Firestore non raggiungibile. Controlla la connessione e riprova.";
  }
  if (/timeout/i.test(getFirebaseErrorMessage(error))) {
    return "Connessione lenta: Firestore non ha risposto entro pochi secondi. Puoi riprovare.";
  }
  return fallback;
}

function logFirestoreError(label, error, extra = {}) {
  const firebaseCode = getFirebaseErrorCode(error);
  const firebaseMessage = getFirebaseErrorMessage(error);
  const firestoreFailureType = firebaseCode.includes("permission-denied")
    ? "permesso negato"
    : firebaseCode.includes("unavailable") || /network|offline|timeout/i.test(firebaseMessage)
      ? "rete"
      : /missing|not-found|undefined|null/i.test(firebaseMessage)
        ? "dato mancante"
        : "errore tecnico";
  console.error(`${label} ERROR`, {
    code: error?.code || "",
    message: firebaseMessage,
    stack: error?.stack || "",
    ...extra
  }, error);
  logActivity("errore_firestore", "Errori Firestore", {
    detail: `${label}: ${firebaseMessage}`,
    technicalError: `${firebaseCode || "firestore-error"}: ${firebaseMessage}${error?.stack ? `\n${error.stack}` : ""}`,
    errorCode: firebaseCode,
    firestoreCollection: extra.collection || extra.collectionPath || extra.path || "Da verificare nel codice",
    firestoreOperation: extra.operation || label || "Operazione Firestore",
    firestoreFailureType,
    unsavedData: extra.unsavedData || extra.payload || extra.data || "Possibili dati non salvati: verificare operazione e payload nel punto errore.",
    firestoreRuleHint: extra.ruleHint || "Controllare allow read/write della collection per ruolo utente, uid, commessaId e campi obbligatori.",
    possibleCause: firestoreFailureType === "permesso negato" ? "Le regole Firestore potrebbero bloccare l'utente o il ruolo corrente." : firestoreFailureType === "rete" ? "Connessione instabile, Firestore non raggiungibile o timeout." : "Dato assente/non valido oppure errore applicativo durante la richiesta.",
    resolutionHint: "Copiare l'errore per Codex, controllare regole Firestore, collection, operazione e dati inviati; poi riprovare il salvataggio."
  });
  if (firebaseCode.includes("permission-denied")) {
    console.error("Verifica regole Firestore: l'utente autenticato deve poter leggere squadre, commesse, impianti, mezzi e utenti/platformUsers.");
  }
}

async function runFirestoreGetWithRetry(query, options = {}) {
  const {
    label = "FIRESTORE QUERY",
    timeoutMs = FIRESTORE_QUERY_TIMEOUT_MS,
    retries = 2,
    retryDelaysMs = FIRESTORE_RETRY_DELAYS_MS,
    onRetry = null
  } = options;
  let lastError = null;
  for (let attempt = 0; attempt <= retries; attempt += 1) {
    try {
      startFirestoreSlowWatch(label);
      const snapshot = await withTimeout(query.get(), timeoutMs, `Timeout ${label}`);
      clearFirestoreSlowWatch();
      if (snapshot?.metadata?.fromCache) {
        setFirestoreConnectionState(navigator.onLine ? "Online" : "Offline", "Dati caricati da cache");
      } else {
        setFirestoreConnectionState(navigator.onLine ? "Online" : "Offline", "");
      }
      return snapshot;
    } catch (error) {
      clearFirestoreSlowWatch();
      lastError = error;
      logFirestoreError(label, error, { attempt: attempt + 1, maxAttempts: retries + 1 });
      if (attempt >= retries) break;
      if (typeof onRetry === "function") onRetry(attempt + 1, error);
      await sleep(retryDelaysMs[attempt] || 1200);
    }
  }
  throw lastError;
}

function normalizeCommessaDocument(doc) {
  const data = doc.data() || {};
  return {
    id: doc.id,
    nome: String(data.nome || data.name || "Commessa senza nome"),
    codice: String(data.codice || data.code || ""),
    parentCommessaId: String(data.parentCommessaId || ""),
    createdAt: data.createdAt || null,
    ...data
  };
}

function normalizeSquadraStoricoDocument(doc, fallbackDateKey = "") {
  const data = doc.data() || {};
  const squadre = Array.isArray(data.squadre) ? data.squadre : [];
  return {
    id: doc.id,
    commessaId: String(data.commessaId || "").trim(),
    dateKey: String(data.dateKey || fallbackDateKey || "").trim(),
    squadre,
    personale: String(data.personale || ""),
    mezzi: String(data.mezzi || ""),
    impianti: String(data.impianti || ""),
    note: String(data.note || ""),
    ...data
  };
}

function normalizePersonaleDocument(doc) {
  const data = doc.data() || {};
  const fullName = String(data.fullName || `${data.cognome || ""} ${data.nome || ""}`.trim() || "Senza nome");
  return {
    id: doc.id,
    nome: String(data.nome || ""),
    cognome: String(data.cognome || ""),
    fullName,
    telefono: String(data.telefono || ""),
    email: normalizeEmail(data.email || ""),
    mansione: String(data.mansione || data.ruolo || ""),
    commesseAbilitate: Array.isArray(data.commesseAbilitate) ? data.commesseAbilitate : [],
    corsi: data.corsi && typeof data.corsi === "object" ? data.corsi : {},
    ...data
  };
}

function getPersonaleByLoginEmail(email = currentUser?.email) {
  const normalizedEmail = normalizeEmail(email);
  if (!normalizedEmail) return null;
  return personaleRecords.find((person) => normalizeEmail(person?.email) === normalizedEmail) || null;
}

function getCurrentUserResolvedName(fallback = "Operatore") {
  const person = getPersonaleByLoginEmail();
  const personnelName = String(person ? getPersonaleDisplayName(person) : "").trim();
  return personnelName
    || String(currentUser?.displayName || "").trim()
    || String(currentUser?.email || "").trim()
    || fallback;
}

function refreshResolvedUserIdentity() {
  if (!currentUser) return;
  const resolvedName = getCurrentUserResolvedName("Nome non disponibile");
  if (ui.userName) ui.userName.textContent = `Nome utente: ${resolvedName}`;
  if (ui.operatorGreeting) ui.operatorGreeting.textContent = `üëã Ciao, ${resolvedName}`;
}

function normalizeMezzoDocument(doc) {
  const data = doc.data() || {};
  return {
    id: doc.id,
    nId: String(data.nId || data.numero || ""),
    marca: String(data.marca || ""),
    modello: String(data.modello || ""),
    targa: String(data.targa || ""),
    ...data,
    posti: normalizeMezzoPosti(data.posti || data.numeroPosti || data.postiTrasporto || "")
  };
}


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

function setHomeAccordionState(trigger, content, expanded, labels = {}) {
  if (!trigger || !content) return;
  trigger.setAttribute("aria-expanded", String(expanded));
  trigger.setAttribute("aria-label", expanded ? (labels.close || "Chiudi dettagli") : (labels.open || "Apri dettagli"));
  content.classList.toggle("hidden", !expanded);
  content.setAttribute("aria-hidden", String(!expanded));
}

function toggleProfileSummary() {
  const expanded = ui.userConnectionBar?.getAttribute("aria-expanded") !== "true";
  setHomeAccordionState(ui.userConnectionBar, ui.profileSummaryDetails, expanded, {
    open: "Apri riepilogo utente",
    close: "Chiudi riepilogo utente"
  });
}

function toggleWeatherCard() {
  const expanded = ui.weatherCard?.getAttribute("aria-expanded") !== "true";
  setHomeAccordionState(ui.weatherCard, ui.weatherExpandedContent, expanded, {
    open: "Apri meteo operativo",
    close: "Chiudi meteo operativo"
  });
}

function updateNotificationUi(message, canTest = false) {
  if (ui.pwaNotificationStatus) ui.pwaNotificationStatus.textContent = message;
  if (ui.testNotificationBtn) ui.testNotificationBtn.disabled = !canTest;
}

function isAppInstalled() {
  return Boolean(
    window.matchMedia?.("(display-mode: standalone)")?.matches
    || window.matchMedia?.("(display-mode: fullscreen)")?.matches
    || window.navigator?.standalone
  );
}

function isIosOrSafariInstallFlow() {
  const userAgent = String(navigator.userAgent || "");
  const vendor = String(navigator.vendor || "");
  const isIos = /iPad|iPhone|iPod/.test(userAgent) || (navigator.platform === "MacIntel" && navigator.maxTouchPoints > 1);
  const isSafari = /Safari/i.test(userAgent) && /Apple/i.test(vendor) && !/CriOS|FxiOS|EdgiOS|Chrome|Chromium|Android/i.test(userAgent);
  return isIos || isSafari;
}

async function handleInstallAppClick() {
  if (isAppInstalled()) {
    alert("L'app risulta gi√† installata sul dispositivo.");
    closeSideMenu();
    return;
  }

  if (deferredInstallPrompt && typeof deferredInstallPrompt.prompt === "function") {
    const promptEvent = deferredInstallPrompt;
    deferredInstallPrompt = null;
    promptEvent.prompt();
    try {
      const choice = await promptEvent.userChoice;
      if (choice?.outcome === "accepted") {
        alert("Installazione app avviata.");
      } else {
        alert("Installazione annullata. Puoi riprovare dal menu laterale quando il prompt sar√† disponibile.");
      }
    } catch (error) {
      console.warn("Risultato prompt installazione non disponibile:", error);
    }
    closeSideMenu();
    return;
  }

  if (isIosOrSafariInstallFlow()) {
    alert("Per installare l'app: Apri Condividi ‚Üí Aggiungi alla schermata Home");
  } else {
    alert("Installazione non disponibile in questo momento: l'app potrebbe essere gi√† installata oppure il browser non ha ancora reso disponibile il prompt.");
  }
  closeSideMenu();
}

window.addEventListener("beforeinstallprompt", (event) => {
  event.preventDefault();
  deferredInstallPrompt = event;
});

window.addEventListener("appinstalled", () => {
  deferredInstallPrompt = null;
});

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
  const centralTypes = {
    "impianto-done": { type: "IMPIANTO_COMPLETATO", priority: "NORMALE", title: "‚úÖ Impianto completato", actionType: "impianto" },
    "impianto-navigate": { type: "NAVIGAZIONE", priority: "NORMALE", title: "üìç Navigazione avviata", actionType: "impianto" },
    "hours-inserted": { type: "ORE", priority: "NORMALE", title: "üïí Ore inserite", actionType: "ore" }
  };
  const central = centralTypes[eventType];
  if (central && window.HeraNotificationCenter) {
    // Fire-and-forget: una notifica non deve mai bloccare FATTO, NAVIGA o il salvataggio ore.
    void window.HeraNotificationCenter.create({
      ...central,
      preview: payload.body || "Nuovo aggiornamento operativo.", message: payload.body || "Nuovo aggiornamento operativo.",
      actorId: currentUser.uid || "", actorName: currentUser.displayName || currentUser.email || "Operatore",
      scopeType: payload.commessaId ? "COMMESSA" : "PERSONALE", recipientUserIds: [currentUser.uid].filter(Boolean),
      commessaId: payload.commessaId || "", commessaName: payload.commessaName || "",
      impiantoId: payload.impiantoKey || "", impiantoName: payload.impiantoName || "", actionTarget: payload.impiantoKey || payload.commessaId || "",
      dedupeKey: `${eventType}:${currentUser.uid}:${payload.impiantoKey || payload.commessaId || "self"}:${eventType === "impianto-navigate" ? Math.floor(Date.now() / 600000) : Date.now()}`,
      metadata: { sourceEvent: eventType }
    });
  }
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

function syncBannerFormFromSelection() {
  if (ui.bannerNoteDate && !ui.bannerNoteDate.value) {
    ui.bannerNoteDate.value = getDateKeyFromLocalDate(new Date());
  }
  renderWorkBannerNotesList(currentWorkBannerConfig.notes || []);
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
        return { text: `üìÖ ${tomorrowLabel} ¬∑ ${tomorrowNote.note}`, isScheduled: true };
      }
    }
    const todayNote = notes.find((entry) => entry.dateKey === todayKey);
    if (todayNote) return { text: todayNote.note, isScheduled: false };
    const nextNote = notes.find((entry) => entry.dateKey >= todayKey) || notes[0];
    const dateLabel = new Date(`${nextNote.dateKey}T00:00:00`).toLocaleDateString("it-IT");
    return { text: `üìÖ ${dateLabel} ¬∑ ${nextNote.note}`, isScheduled: true };
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
  ui.workBannerText.textContent = `${activeMessage.text}   ‚Ä¢   ${activeMessage.text}   ‚Ä¢   ${activeMessage.text}`;
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
        ui.bannerFeedback.textContent = "Errore lettura banner. Riprova pi√π tardi.";
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
    0: "‚òÄÔ∏è Sereno",
    1: "‚õÖ Poco nuvoloso",
    2: "‚òÅÔ∏è Parzialmente nuvoloso",
    3: "‚òÅÔ∏è Coperto",
    45: "üå´Ô∏è Nebbia",
    48: "üå´Ô∏è Nebbia con brina",
    51: "üå¶Ô∏è Pioviggine",
    53: "üå¶Ô∏è Pioviggine moderata",
    55: "üåßÔ∏è Pioviggine intensa",
    56: "üå®Ô∏è Pioviggine gelata",
    57: "üå®Ô∏è Pioggia gelata",
    61: "üåßÔ∏è Pioggia debole",
    63: "üåßÔ∏è Pioggia moderata",
    65: "‚õàÔ∏è Pioggia forte",
    66: "üßä Pioggia gelata debole",
    67: "üßä Pioggia gelata forte",
    71: "üå®Ô∏è Neve debole",
    73: "üå®Ô∏è Neve moderata",
    75: "‚ùÑÔ∏è Neve intensa",
    77: "üå®Ô∏è Nevischio",
    80: "üåßÔ∏è Rovesci deboli",
    81: "üåßÔ∏è Rovesci moderati",
    82: "‚õàÔ∏è Rovesci forti",
    85: "üå®Ô∏è Rovesci di neve",
    86: "‚ùÑÔ∏è Rovesci di neve forti",
    95: "‚õàÔ∏è Temporale",
    96: "‚õàÔ∏è Temporale con grandine",
    99: "‚õàÔ∏è Temporale forte con grandine"
  };
  return weatherMap[code] || "‚ÑπÔ∏è Condizioni variabili";
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

console.log("AUTH CHECK START");

if (!auth || firebaseInitError) {
  console.error("Errore verifica login Firebase:", firebaseInitError || "Auth non disponibile");
  setAuthenticationGateState("required", "Login non disponibile: configurazione Firebase non caricata correttamente.");
  authStateResolved = true;
  currentUser = null;
  squadreLoadState = { status: "error", message: "Errore caricamento dati" };
  if (typeof renderSquadre === "function") renderSquadre();
  hideStartupLoading();
} else {
  setAuthenticationGateState("checking");
  let savedStartupSession = null;
  void readPersistedSession().then((session) => {
    savedStartupSession = session;
    if (!authStateResolved && session) applyPersistedSessionPreview(session);
  });
  if (ui.loginBtn) ui.loginBtn.disabled = true;
  if (ui.user) ui.user.textContent = "Verifica sessione in corso...";
  const authCheckWatchdog = setTimeout(() => {
    if (authStateResolved) return;
    console.warn("Timeout verifica sessione Firebase: mostro la schermata di login invece del caricamento infinito.");
    authStateResolved = true;
    currentUser = null;
    setAuthenticationGateState("required", "Verifica sessione lenta o non disponibile. Riprova il login.");
    squadreLoadState = { status: "auth-required", message: "Fai login per caricare le squadre." };
    if (typeof renderSquadre === "function") renderSquadre();
    hideStartupLoading();
  }, 8000);
  withTimeout(ensureAuthLocalPersistence(), 3500, "Timeout persistenza auth locale")
    .catch((error) => {
      console.warn("Persistenza auth locale non pronta, continuo comunque la verifica sessione:", error);
      return false;
    })
    .finally(() => {
    auth.onAuthStateChanged(async (user) => {
  clearTimeout(authCheckWatchdog);
  console.log("AUTH READY");
  authStateResolved = true;
  currentUser = user || null;
  currentUserBanProfile = null;
  const loggedIn = Boolean(user);
  console.log(loggedIn ? "USER LOGGED" : "USER NOT LOGGED", {
    email: user?.email || "",
    uid: user?.uid || ""
  });
  if (loggedIn) {
    try {
      const savedSession = savedStartupSession || await readPersistedSession();
      const hasSavedApproval = isPersistedApprovalValid(savedSession, user);
      if (!hasSavedApproval) {
        const authorization = await window.HeraAccessApproval.verify(user);
        if (!authorization.allowed) {
          stopCommesseSubscription(); stopImpiantiSubscription(); stopChatSubscription();
          stopPersonaleSubscription(); stopMezziSubscription(); stopSquadreSubscription(); stopUsersSubscription();
          window.location.hash = "";
          hideStartupLoading();
          return;
        }
        await savePersistedSession(user, { ...authorization.profile, accessApproved: true });
      }
      const databaseCheck = hasSavedApproval
        ? { valid: true, profile: { ...savedSession, accessApproved: true }, banned: false }
        : await verifyPersistedSessionAgainstDatabase(user, savedSession);
      if (databaseCheck.banned) {
        currentUserBanProfile = databaseCheck.profile || { email: user.email, displayName: user.displayName };
        await savePersistedSession(user, currentUserBanProfile);
        stopCommesseSubscription();
        stopImpiantiSubscription();
        stopCommessaNotesSubscription();
        stopChatSubscription();
        stopPersonaleSubscription();
        stopMezziSubscription();
        stopSquadreSubscription();
        stopUsersSubscription();
        stopOperatorPositionsSubscription();
        stopGlobalNotificationsSubscription();
        stopUserAlertsSubscription();
        setAuthenticationGateState("banned");
        hideStartupLoading();
        return;
      }
      currentUserBanProfile = null;
      if (!databaseCheck.valid) {
        console.warn("Sessione salvata non valida: utente non presente in platformUsers.");
        await clearPersistedSession();
        await auth.signOut();
        currentUser = null;
        setAuthenticationGateState("required", "Sessione non pi√π valida. Effettua di nuovo il login.");
        hideStartupLoading();
        return;
      }
      await savePersistedSession(user, databaseCheck.profile || {});
    } catch (error) {
      console.error("Errore verifica sessione salvata:", error);
      if (String(error?.code || "").startsWith("auth/")) {
        await clearPersistedSession();
      }
    }
    if (!(await requireBiometricAtStartup())) return;
    console.log("USER UID", user.uid);
    logActivity("login_app", "Login app");
    logActivity("apertura_app", "Apertura app");
  } else {
    savedStartupSession = null;
  }
  setAuthenticationGateState(loggedIn ? "authenticated" : "required");
  if (loggedIn) setTimeout(() => {
    void refreshBiometricSettings();
    void offerBiometricsAfterGoogleLogin();
  }, 0);

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
  stopCalendarEventsSubscription();
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
  if (!loggedIn) window.location.hash = "";
  commesseLoadState = { status: "idle", message: "" };
  isCommesseHomeCardVisible = false;
  syncCommesseHomeToggle();
  ui.commesseLista.innerHTML = "";
  ui.squadraCommessa.innerHTML = "<option value=''>Seleziona commessa</option>";
  ui.squadreLista.innerHTML = "";
  squadreLoadState = { status: "loading", message: "Caricamento squadre..." };
  personaleLoadState = { status: "idle", message: "" };
  mezziLoadState = { status: "idle", message: "" };
  startupCoreCollectionsLoadState = { status: "idle", message: "" };
  manualSquadreFilterDateKey = "";
  sharedSquadreDateKey = "";
  startupAssignedCommessaAutoOpenDone = false;
  sharedSquadreViewConfigLoaded = false;
  squadreByCommessa = new Map();
  squadreHistoryByDate = new Map();
  commesseById = new Map();
  personaleRecords = [];
  mezziRecords = [];
  initializeAutomaticSquadreDate();
  globalCommesseById = new Map();
  globalImpianti = [];
  pendingGlobalRows = [];
  selectedGlobalCommessaId = "";
  resourceRecords = [];
  privateDocsRecords = [];
  calendarEvents = [];
  calendarAbsenceCache.clear();
  confirmedSquadraAbsenceAssignments.clear();
  posDocuments = [];
  gpsUpdateRequests = [];
  operatorPositions = [];
  hoursApprovalRequests = [];
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

  if (loggedIn) {
    hideStartupLoading();
    const loadInitialData = isSnowServiceRoute() && canManageData()
      ? (() => {
          document.body.classList.add("snow-management-context");
          return loadSnowModeData();
        })
      : loadStartupCoreCollections;
    loadInitialData()
      .catch((error) => {
        console.error("Caricamento iniziale collezioni principali non completato:", error);
      })
      .finally(() => {
        renderHeaderActivitySummary();
        renderExternalApps();
        renderPendingWhatsappList();
        fetchWeather();
        if (!isSnowServiceContext()) {
          syncPendingImpiantoActions();
          syncPendingOfflineMutations();
        }
        renderNextActionCard();
      });
    runDeferredStartupTasks([
      () => startPresenceHeartbeat(),
      () => upsertCurrentPlatformUser(),
      () => initGeolocation(),
      () => subscribeUsers(),
      () => subscribeAdminUsers(),
      () => subscribeChat(),
      () => subscribeOperatorPositions(),
      () => subscribeDriveBridge(),
      () => subscribeResources(),
      () => subscribeGlobalCommesse(),
      () => subscribePrivateDocs(),
      () => subscribePosDocuments(),
      () => subscribeGpsRequests(),
      () => subscribeGlobalNotifications(),
      () => subscribeWorkBanner(),
      () => subscribeUserAlerts(),
      () => initHelpCenterFaq(),
      () => processPendingSheetExports(),
      () => startChatRetentionLoop(),
      () => startHoursDeadlineAlertLoop(),
      () => loadWeatherAlertsForActiveDate().catch((error) => console.error("Caricamento allerte meteo non riuscito:", error)),
      () => loadWorklimateRiskCacheBackground(),
      () => repairDuplicateHours().catch((error) => {
        console.error("Riparazione automatica duplicati ore all'avvio non riuscita:", error);
      })
    ]);
  } else {
    squadreLoadState = { status: "auth-required", message: "Fai login per caricare le squadre." };
    renderSquadre();
    if (ui.commesseLista) ui.commesseLista.innerHTML = "<p class='muted'>Fai login per vedere le commesse.</p>";
    if (ui.squadraCommessa) ui.squadraCommessa.innerHTML = "<option value=''>Login richiesto</option>";
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
  console.log("APP READY");
  hideStartupLoading();
}, async (error) => {
  clearTimeout(authCheckWatchdog);
  console.error("Errore verifica login Firebase:", error);
  if (String(error?.code || "").startsWith("auth/")) await clearPersistedSession();
  authStateResolved = true;
  currentUser = null;
  setAuthenticationGateState("required", "Non riesco a verificare la sessione. Riprova il login.");
  squadreLoadState = { status: "error", message: "Errore caricamento dati" };
  renderSquadre();
  hideStartupLoading();
});
  });
}


function stopNormalDataSubscriptionsForSnowMode() {
  stopCommesseSubscription();
  stopSquadreSubscription();
  stopPersonaleSubscription();
  stopMezziSubscription();
}

function clearSnowModeRuntimeData() {
  snowServiceState.clients = [];
  snowServiceState.routes = [];
  snowServiceState.vehicles = [];
  snowServiceState.operators = [];
  snowServiceState.reports = [];
  commesseById = new Map();
  squadreByCommessa = new Map();
  squadreHistoryByDate = new Map();
  personaleRecords = [];
  mezziRecords = [];
  commesseLoadState = { status: "idle", message: "" };
  squadreLoadState = { status: "idle", message: "" };
  personaleLoadState = { status: "idle", message: "" };
  mezziLoadState = { status: "idle", message: "" };
}

function stopSnowServiceCollections() {
  snowServiceUnsubscribers.forEach((unsubscribe) => unsubscribe && unsubscribe());
  snowServiceUnsubscribers = [];
}

function stopSnowModeData() {
  stopSnowServiceCollections();
  stopCommesseSubscription();
  stopSquadreSubscription();
  stopPersonaleSubscription();
  stopMezziSubscription();
  clearSnowModeRuntimeData();
}

function loadSnowModeData() {
  if (!currentUser || !isSnowServiceContext()) return Promise.resolve(false);
  return Promise.all([
    subscribeCommesse(),
    subscribeSquadre(),
    subscribePersonale(),
    subscribeMezzi(),
    subscribeSnowServiceCollections()
  ]);
}

function reloadNormalModeData() {
  if (!currentUser) return Promise.resolve(false);
  commesseById = new Map();
  squadreByCommessa = new Map();
  squadreHistoryByDate = new Map();
  personaleRecords = [];
  mezziRecords = [];
  renderCommesseHomeList();
  renderSquadre();
  return Promise.all([
    subscribeCommesse(),
    subscribeSquadre(),
    subscribePersonale(),
    subscribeMezzi()
  ]);
}

async function loadStartupCoreCollections() {
  if (!currentUser) return;
  startupCoreCollectionsLoadState = { status: "loading", message: "Caricamento dati iniziali..." };
  commesseLoadState = { status: "loading", message: "Caricamento commesse..." };
  personaleLoadState = { status: "loading", message: "Caricamento anagrafica personale..." };
  mezziLoadState = { status: "loading", message: "Caricamento mezzi..." };
  squadreLoadState = { status: "loading", message: "Caricamento squadre..." };
  renderCommesseHomeList();
  renderSquadre();

  try {
    const personalePromise = subscribePersonale();
    await Promise.all([
      personalePromise, // anagrafiche personale
      personalePromise, // qualifiche/corsi salvati sulle anagrafiche
      personalePromise, // sicurezza salvata sulle anagrafiche
      subscribeSquadre(),
      subscribeCommesse(),
      subscribeMezzi()
    ]);
    startupCoreCollectionsLoadState = { status: "loaded", message: "" };
  } catch (error) {
    startupCoreCollectionsLoadState = { status: "error", message: getReadableFirestoreError(error, "Errore caricamento dati iniziali") };
    throw error;
  } finally {
    renderCommesseHomeList();
    renderSquadre();
  }
}

function areStartupCoreCollectionsLoading() {
  return startupCoreCollectionsLoadState.status === "loading";
}

function isPersonaleReadyForSquadraValidation() {
  return personaleLoadState.status === "loaded" && personaleRecords.length > 0;
}

function updateAdminControls() {
  const canManage = canManageData();
  updateDriveConnectVisibility();
  ui.openPosBtn?.classList.remove("hidden");
  if (ui.openPosBtn) ui.openPosBtn.disabled = false;
  ui.operatorPositionsToggleBtn?.classList.add("hidden");
  if (ui.operatorPositionsToggleBtn) ui.operatorPositionsToggleBtn.disabled = true;
  ui.chatClearBtn?.classList.toggle("hidden", !canManage);
  ui.snowServiceBtn?.classList.toggle("hidden", !canManage);
  if (ui.snowServiceBtn) ui.snowServiceBtn.disabled = !canManage;
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
  ui.snowSquadreFilterControls?.classList.toggle("hidden", !canManage);
  if (ui.squadreFilterDate) ui.squadreFilterDate.disabled = !canManage;
  if (ui.squadreFilterClearBtn) ui.squadreFilterClearBtn.disabled = !canManage;
  if (ui.snowSquadreFilterDate) ui.snowSquadreFilterDate.disabled = !canManage;
  if (ui.snowSquadreFilterClearBtn) ui.snowSquadreFilterClearBtn.disabled = !canManage;
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
    : "Solo l'admin pu√≤ modificare personale, mezzi e composizione squadre.";
  updateResourceFormByType();
  renderUserPermissionList();
  renderNotificationTargetUsers();
  renderNotificationsList();
  renderExternalApps();
  renderCommesseHomeList();
}

function openSideMenu() {
  configureSnowSideMenu(isSnowServiceRoute());
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

const ANDROID_PLAY_STORE_URL = "https://play.google.com/store/apps/details?id=it.vargacantieri.hera";
const NETLIFY_APP_URL = "https://creative-syrniki-dddbae.netlify.app/";

async function openApplicationUpdate() {
  const isAndroid = Boolean(
    window.Capacitor?.isNativePlatform?.()
    && window.Capacitor?.getPlatform?.() === "android"
  );

  if (ui.updateAppBtn) {
    ui.updateAppBtn.disabled = true;
    ui.updateAppBtn.setAttribute("aria-label", isAndroid
      ? "Apertura aggiornamento nel Play Store"
      : "Apertura aggiornamento da Netlify");
  }

  if (isAndroid) {
    // _system keeps the Play Store outside the Capacitor WebView, preventing the
    // app from replacing its own page and becoming stuck on the splash screen.
    const storeWindow = window.open(ANDROID_PLAY_STORE_URL, "_system", "noopener,noreferrer");
    if (!storeWindow) window.location.href = ANDROID_PLAY_STORE_URL;
    return;
  }

  try {
    const registration = await navigator.serviceWorker?.getRegistration?.();
    await registration?.update?.();
    if (ui.updateAppBtn) {
      ui.updateAppBtn.disabled = false;
      ui.updateAppBtn.title = "Aggiornamento controllato. Le novit√† saranno attive alla prossima apertura.";
      ui.updateAppBtn.setAttribute("aria-label", "App aggiornata; riapri l'app per applicare le novit√†");
    }
  } catch (error) {
    console.warn("Controllo aggiornamenti non riuscito:", error);
    if (ui.updateAppBtn) ui.updateAppBtn.disabled = false;
  }
}

function openManagementPanel(panel) {
  if (panel !== "banner" && !canManageData()) {
    closeSideMenu();
    return;
  }
  const panelMap = {
    commesse: { el: ui.panelCommesse, title: "Gestione commesse" },
    squadre: { el: ui.panelSquadre, title: "Composizione squadre" },
    personale: { el: ui.panelPersonale, title: "Personale" },
    mezzi: { el: ui.panelMezzi, title: "Mezzi" },
    utenti: { el: ui.panelUtenti, title: "Gestione utenti" },
    global: { el: ui.panelGlobal, title: "Global" },
    banner: { el: ui.panelBanner, title: "Banner home" },
    infoUtili: { el: ui.panelInfoUtili, title: "Informazioni utili" },
    notifiche: { el: ui.panelNotifiche, title: "Gestione notifiche" },
    programmazione: { el: ui.panelProgrammazione, title: "üìÖ Programmazione" }
  };
  const target = panelMap[panel];
  if (!target) return;
  [ui.panelCommesse, ui.panelSquadre, ui.panelPersonale, ui.panelMezzi, ui.panelUtenti, ui.panelGlobal, ui.panelBanner, ui.panelInfoUtili, ui.panelNotifiche, ui.panelProgrammazione].forEach((el) => el?.classList.add("hidden"));
  target.el.classList.remove("hidden");
  ui.managementTitle.textContent = target.title;
  ui.managementPage.classList.remove("hidden");
  ui.managementPage.setAttribute("aria-hidden", "false");
  if (panel === "squadre") setDefaultSquadraCompositionDate({ force: true });
  if (panel === "global") setTimeout(() => globalMap.invalidateSize(), 60);
  if (panel === "notifiche") closeNotificationCalendarView();
  if (panel === "programmazione") void subscribeProgrammazioni();
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
  ui.mapFullscreenBtn.textContent = "‚§¢ Mappa a schermo intero";
  ui.mapDrawAreaBtn.textContent = "‚úèÔ∏è Disegna";
  const saveSnowRoadBtn = document.getElementById("map-save-snow-road-btn");
  saveSnowRoadBtn?.classList.toggle("hidden", !isSnowServiceContext());
  syncDrawAreaToolbarState();
  setFullscreenFeedback(isSnowServiceContext() ? "Seleziona una via neve, premi Disegna e salva il tracciato della strada." : "Usa ‚ÄúDisegna‚Äù per definire il perimetro di lavoro.");
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
  ui.mapDrawAreaBtn.textContent = "‚úèÔ∏è Disegna";
  document.getElementById("map-save-snow-road-btn")?.classList.add("hidden");
  syncDrawAreaToolbarState();
  setFullscreenFeedback("Usa ‚ÄúDisegna‚Äù per definire il perimetro di lavoro.");
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
    button: "üåß Pioggia",
    title: "Pioggia / precipitazioni",
    label: "Pioggia",
    opacity: 0.6,
    providerLayers: { openweather: "PR0", rainviewer: "radar" },
    legend: ["üü¶ Debole", "üü© Moderata", "üü® Forte", "üü• Molto forte"],
    description: "Precipitazioni OpenWeatherMap Weather Maps 2.0; fallback RainViewer solo se il provider principale non √® configurato."
  },
  clouds: {
    id: "clouds",
    button: "‚òÅÔ∏è Nuvole",
    title: "Nuvolosit√†",
    label: "Nuvole",
    opacity: 0.52,
    providerLayers: { openweather: "CL", rainviewer: "satellite" },
    legend: ["‚¨õ Sereno", "‚¨ú Nubi basse", "‚òÅÔ∏è Nubi dense", "üå© Celle compatte"],
    description: "Copertura nuvolosa da OpenWeatherMap Weather Maps 2.0."
  },
  temperature: {
    id: "temperature",
    button: "üå° Temperatura",
    title: "Temperatura",
    label: "Temperatura",
    opacity: 0.46,
    providerLayers: { openweather: "TA2" },
    legend: ["üü¶ Freddo", "üü© Mite", "üü® Caldo", "üü• Molto caldo"],
    description: "Temperatura a 2 metri da OpenWeatherMap Weather Maps 2.0."
  },
  wind: {
    id: "wind",
    button: "üí® Vento",
    title: "Vento",
    label: "Vento",
    opacity: 0.44,
    providerLayers: { openweather: "WND" },
    openWeatherParams: { use_norm: "true", arrow_step: "32" },
    legend: ["üü¶ Brezza", "üü© Moderato", "üü® Forte", "üü• Raffiche"],
    description: "Velocit√† e direzione del vento da OpenWeatherMap Weather Maps 2.0."
  },
  storms: {
    id: "storms",
    button: "‚ö° Temporali",
    title: "Temporali",
    label: "Temporali",
    opacity: 0.58,
    providerLayers: { openweather: "PAC0", rainviewer: "radar" },
    usesFallbackLayerMessage: true,
    legend: ["üü¶ Rovesci", "üü© Pioggia", "üü® Celle intense", "üü• Possibile temporale"],
    description: "Temporali stimati dal layer di precipitazione convettiva OpenWeatherMap; se non disponibile usa il radar precipitazioni come fallback operativo."
  },
  alerts: {
    id: "alerts",
    button: "‚ö†Ô∏è Allerte",
    title: "Allerte meteo",
    label: "Allerte",
    opacity: 0.5,
    providerLayers: {},
    unavailable: true,
    legend: ["‚ö†Ô∏è Dato non disponibile"],
    description: "Dato non disponibile: le allerte ufficiali non sono esposte come tile meteo in questo provider/piano. La mappa resta navigabile."
  }
};

const WEATHER_PROVIDERS = {
  openweather: {
    id: "openweather",
    label: "OpenWeatherMap",
    priority: 1,
    attribution: "Meteo ¬© OpenWeatherMap",
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
    attribution: "Meteo ¬© RainViewer",
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
    playBtn.textContent = radarPlaying ? "‚è∏" : "‚ñ∂";
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
      : `${definition.label} ‚Äî ${WEATHER_UNAVAILABLE_MESSAGE}`;
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
    <button class="weather-radar-play" type="button" data-radar-play aria-label="Pausa radar meteo">‚è∏</button>
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
  ui.mapSatelliteToggleBtn.textContent = isSatellite ? "üó∫ Standard" : "üõ∞ Satellite";
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
    ui.mapDrawAreaBtn.textContent = "‚úÖ Termina";
    syncDrawAreaToolbarState();
    setFullscreenFeedback("Modalit√† disegno attiva: trascina il dito per tracciare l'area.");
    return;
  }
  setFullscreenMapInteractivity(true);
  ui.mapDrawAreaBtn.textContent = "‚úèÔ∏è Disegna";
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
      console.warn("Pointer capture gi√† rilasciato", error);
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
  const saveSnowRoadBtn = document.getElementById("map-save-snow-road-btn");
  if (saveSnowRoadBtn) saveSnowRoadBtn.disabled = !isSnowServiceContext() || !selectedImpiantoData?.snowRoad || drawnAreaPoints.length < 2;
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
  setFullscreenFeedback("Disegno annullato. Premi ‚ÄúRifai‚Äù per ripristinarlo.");
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

async function saveDrawnSnowRoadPath() {
  if (!isSnowServiceContext() || !selectedCommessaId || !selectedImpiantoData?.snowRoad) {
    alert("Seleziona prima una via neve nella mappa.");
    return;
  }
  if (drawnAreaPoints.length < 2) {
    alert("Disegna almeno due punti sulla strada.");
    return;
  }
  const path = drawnAreaPoints.map((point) => ({ lat: Number(point[0]), lng: Number(point[1]) }));
  const impiantoId = selectedImpiantoData.id;
  if (!impiantoId) return;
  await db.collection("neve_commesse").doc(selectedCommessaId).collection("impianti").doc(impiantoId).set({
    routePath: path,
    updatedAt: firebase.firestore.FieldValue.serverTimestamp(),
    updatedBy: currentUser?.email || ""
  }, { merge: true });
  selectedImpiantoData = { ...selectedImpiantoData, routePath: path };
  setFullscreenFeedback("Tracciato via neve salvato: la linea diventer√† verde quando l‚Äôoperatore passa sulla strada.");
  renderMap();
}

function shareDrawnAreaViaWhatsapp() {
  if (drawnAreaPoints.length < 3) {
    alert("Disegna almeno 3 punti per creare un'area condivisibile.");
    return;
  }
  const vertices = drawnAreaPoints.map((point, idx) => `‚Ä¢ Punto ${idx + 1}: ${point[0].toFixed(6)}, ${point[1].toFixed(6)}`).join("\n");
  const areaCoords = drawnAreaPoints.map((point) => `${point[0].toFixed(6)},${point[1].toFixed(6)}`).join(" | ");
  const message = [
    "üó∫Ô∏è *Area lavoro commessa*",
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
  if (!rawHash.startsWith("commessa=")) return { id: "", resource: "", notes: false, impianto: "", meteo: "", alert: "", atex: "", safety: "", capitolato: "", biogas: false, tombini: false };
  const params = new URLSearchParams(rawHash);
  return {
    id: params.get("commessa") || "",
    resource: params.get("resource") || "",
    notes: params.has("notes"),
    impianto: params.get("impianto") || "",
    meteo: params.get("meteo") || "",
    alert: params.get("alert") || "",
    atex: params.get("atex") || "",
    safety: params.get("safety") || "",
    capitolato: params.get("capitolato") || "",
    biogas: params.has("biogas"),
    tombini: params.has("tombini")
  };
}

function setCommessaHash(suffix = "") {
  if (!selectedCommessaId) return;
  window.location.hash = `commessa=${encodeURIComponent(selectedCommessaId)}${suffix}`;
}

function focusSharedImpiantoFromRoute(impiantoKey) {
  const key = String(impiantoKey || "").trim();
  if (!key) return false;
  const impianto = findCurrentImpiantoByKey(key);
  if (!impianto) return false;
  const targetViewMode = impianto.done ? "done" : "todo";
  if (impiantiViewMode !== targetViewMode) setImpiantiViewMode(targetViewMode);
  if (impiantiSearchTerm) {
    impiantiSearchTerm = "";
    if (ui.impiantoSearch) ui.impiantoSearch.value = "";
  }
  focusImpiantoInList(impianto, true);
  selectImpiantoForMapDetail(impianto);
  return true;
}

function applyRoute() {
  if (authStateResolved && !currentUser) {
    if (window.location.hash) {
      window.location.hash = "";
      return;
    }
  }
  const hash = window.location.hash || "";
  const commessaRoute = parseCommessaHash(hash);
  const fuelMatch = hash.match(/^#fuel=(.+)$/);
  const showSegnalazioni = hash === "#segnalazioni";
  const showHowto = hash === "#howto";
  const showControlCenter = hash === "#centro-controllo";
  const showPrivateDocs = hash === "#documenti";
  const showCalendar = hash === "#calendario";
  const showPos = hash === "#pos" || (window.location.pathname === "/pos" && !hash);
  const personalServiceMatch = hash.match(/^#servizi-personali(?:=([a-z]+))?$/);
  const showHours = hash === "#ore";
  const showActiveUsersDetail = hash === "#dettaglio-utenti-attivi";
  const userActivityMatch = hash.match(/^#attivita-utente=([^&]+)$/);
  const showUserActivity = Boolean(userActivityMatch);
  const showSnowService = hash === "#servizio-neve" && canManageData();
  if (hash === "#servizio-neve" && !canManageData()) {
    closeSnowServicePage();
    return;
  }
  const commessaIdFromHash = commessaRoute.id;
  const resourceTypeFromHash = commessaRoute.resource;
  const showFuel = Boolean(fuelMatch);
  const showPersonalServices = Boolean(personalServiceMatch);
  const showNotesPage = Boolean(commessaRoute.notes && selectedCommessaId === commessaIdFromHash);
  const showWeatherDetail = Boolean(commessaRoute.meteo && selectedCommessaId === commessaIdFromHash && !showNotesPage);
  const showWeatherAlertSafety = Boolean(commessaRoute.alert && selectedCommessaId === commessaIdFromHash && !showNotesPage && !showWeatherDetail);
  const showAtexProcedure = Boolean(commessaRoute.atex && selectedCommessaId === commessaIdFromHash && !showNotesPage && !showWeatherDetail && !showWeatherAlertSafety);
  const showCapitolatoOperativo = Boolean(commessaRoute.capitolato && selectedCommessaId === commessaIdFromHash && !showNotesPage && !showWeatherDetail && !showAtexProcedure);
  const showImpiantoSafety = Boolean(commessaRoute.safety && selectedCommessaId === commessaIdFromHash && !showNotesPage && !showWeatherDetail && !showAtexProcedure && !showCapitolatoOperativo);
  const showBiogasMap = Boolean(commessaRoute.biogas && selectedCommessaId === commessaIdFromHash && isBiogasEnabledForCurrentCommessa());
  const showTombiniMap = Boolean(commessaRoute.tombini && selectedCommessaId === commessaIdFromHash && isTombiniEnabledForCurrentCommessa());
  const showImpianti = Boolean(commessaIdFromHash && selectedCommessaId === commessaIdFromHash && !showNotesPage && !showWeatherDetail && !showWeatherAlertSafety && !showAtexProcedure && !showImpiantoSafety && !showCapitolatoOperativo && !showBiogasMap && !showTombiniMap);
  const showResourceViewer = Boolean(showImpianti && resourceTypeFromHash);
  ui.homePage.classList.toggle("hidden", showActiveUsersDetail || showUserActivity || showImpianti || showNotesPage || showWeatherDetail || showWeatherAlertSafety || showAtexProcedure || showImpiantoSafety || showCapitolatoOperativo || showBiogasMap || showTombiniMap || showFuel || showSegnalazioni || showHowto || showControlCenter || showPrivateDocs || showCalendar || showPos || showHours || showPersonalServices || showSnowService);
  ui.impiantiPage.classList.toggle("hidden", !showImpianti || isMapFullscreenPageOpen);
  ui.weatherAlertSafetyPage?.classList.toggle("hidden", !showWeatherAlertSafety);
  ui.impiantoWeatherDetailPage?.classList.toggle("hidden", !showWeatherDetail);
  ui.atexProcedurePage?.classList.toggle("hidden", !showAtexProcedure);
  ui.impiantoSafetyPage?.classList.toggle("hidden", !(showImpiantoSafety || showCapitolatoOperativo));
  ui.commessaNotesPage?.classList.toggle("hidden", !showNotesPage);
  ui.mapFullscreenPage?.classList.toggle("hidden", !isMapFullscreenPageOpen);
  ui.biogasMapPage?.classList.toggle("hidden", !(showBiogasMap || showTombiniMap));
  ui.fuelPage.classList.toggle("hidden", !showFuel);
  ui.personalServicesPage.classList.toggle("hidden", !showPersonalServices);
  ui.segnalazioniPage.classList.toggle("hidden", !showSegnalazioni);
  ui.howtoPage.classList.toggle("hidden", !showHowto);
  ui.controlCenterPage?.classList.toggle("hidden", !showControlCenter);
  ui.privateDocsPage.classList.toggle("hidden", !showPrivateDocs);
  ui.calendarPage?.classList.toggle("hidden", !showCalendar);
  ui.posPage?.classList.toggle("hidden", !showPos);
  ui.hoursPage.classList.toggle("hidden", !showHours);
  ui.activeUsersDetailPage?.classList.toggle("hidden", !showActiveUsersDetail);
  ui.userActivityPage?.classList.toggle("hidden", !showUserActivity);
  if (showActiveUsersDetail) openActiveUsersDetailView(); else activeUsersLoaded = false;
  if (showUserActivity) openUserActivityView(decodeURIComponent(userActivityMatch[1]));
  document.getElementById("snow-service-page")?.classList.toggle("hidden", !showSnowService);
  document.getElementById("snow-service-page")?.setAttribute("aria-hidden", String(!showSnowService));
  if (showSnowService) renderSnowService();
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
  if (showWeatherAlertSafety) {
    renderWeatherAlertSafetyPage(commessaRoute.alert);
  }
  if (showAtexProcedure) {
    renderAtexProcedurePage(commessaRoute.atex);
  }
  if (showImpiantoSafety) {
    renderImpiantoSafetyPage(commessaRoute.safety);
  }
  if (showBiogasMap) loadBiogasNetworkForCurrentCommessa({ type: "rete_biogas" });
  else if (showTombiniMap) loadBiogasNetworkForCurrentCommessa({ type: "tombini" });
  else teardownBiogasMapPage();
  if (showCapitolatoOperativo) {
    renderCapitolatoOperativoPage(commessaRoute.capitolato);
  }
  if (showImpianti) {
    ui.impiantiPageTitle.textContent = `Impianti commessa: ${selectedCommessaName || "Commessa"}`;
    if (showResourceViewer) {
      activeResourceTypeForViewer = resourceTypeFromHash;
      renderCommessaResourceViewer();
      ui.commessaResourceViewer.classList.remove("hidden");
      ui.commessaResourceViewer.classList.add("page-mode");
      ui.commessaResourceViewerCloseBtn.textContent = "‚Üê Torna alla commessa";
    } else {
      closeCommessaResourceViewer();
    }
    setTimeout(() => {
      map.invalidateSize();
      if (commessaRoute.impianto) focusSharedImpiantoFromRoute(commessaRoute.impianto);
    }, 50);
  }
  if (showHowto) renderHowtoFaq();
  if (showControlCenter) renderControlCenter();
  if (showPrivateDocs) renderPrivateDocsList();
  if (showCalendar) {
    subscribeCalendarEvents();
    renderCalendar();
  }
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
  if (actionKey === "navigate") return "üó∫Ô∏è";
  if (actionKey === "done") return "‚úÖ";
  return "‚úâÔ∏è";
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


function formatControlCenterDate(value) {
  if (!value) return "Non disponibile";
  const date = value?.toDate ? value.toDate() : new Date(value);
  if (Number.isNaN(date.getTime())) return "Non disponibile";
  return date.toLocaleString("it-IT", { dateStyle: "medium", timeStyle: "short" });
}

function getControlCenterRoleLabel() {
  return canManageData() ? "Amministratore" : "Operatore";
}

function getConnectionQuality() {
  if (typeof navigator !== "undefined" && navigator.onLine === false) return { label: "Assente", color: "red", status: "Offline" };
  const connection = navigator.connection || navigator.mozConnection || navigator.webkitConnection;
  const mbps = Number(connection?.downlink || 0);
  if (mbps && mbps < 1.5) return { label: "Debole", color: "yellow", status: "Online" };
  if (mbps && mbps < 5) return { label: "Buona", color: "yellow", status: "Online" };
  return { label: "Ottima", color: "green", status: "Online" };
}

function getNetworkTypeLabel() {
  const connection = navigator.connection || navigator.mozConnection || navigator.webkitConnection;
  const type = String(connection?.effectiveType || connection?.type || "").toLowerCase();
  if (type.includes("wifi")) return "Wi‚ÄëFi";
  if (type.includes("ethernet")) return "Ethernet";
  if (type.includes("5g")) return "5G";
  if (type.includes("4g")) return "4G";
  return type ? type.toUpperCase() : "Rete non rilevata";
}

function getControlCenterPendingItems() {
  const offline = loadPendingOfflineMutations().filter((item) => !item.userId || item.userId === currentUser?.uid || canManageData());
  const done = (canManageData() ? pendingImpiantoActions : getCurrentUserPendingActions()).filter(isActionWaitingForSync);
  return [...done.map((item) => ({ ...item, controlType: "Impianto FATTO offline", when: item.createdAt || item.doneAt, operator: item.doneBy || item.userEmail })), ...offline.map((item) => ({ ...item, controlType: getOfflineMutationLabel(item), when: item.createdAt, operator: item.payload?.operatorName || item.userEmail }))]
    .sort((a, b) => String(a.when || "").localeCompare(String(b.when || "")));
}

function buildControlCenterCard(title, rows, options = {}) {
  const detail = rows.map((row) => `<div class="control-center-row"><span>${escapeHTML(row[0])}</span><strong>${escapeHTML(String(row[1] ?? "-"))}</strong></div>`).join("");
  return `<details class="card control-center-card" open><summary><span class="control-dot ${options.color || "green"}"></span>${escapeHTML(title)}</summary>${detail}${options.extra || ""}</details>`;
}

function renderControlCenter() {
  if (!ui.controlCenterContent) return;
  const isAdmin = canManageData();
  const quality = getConnectionQuality();
  const installedVersion = document.querySelector('meta[name="app-version"]')?.content || "1.0.0";
  const pending = getControlCenterPendingItems();
  const impiantiCount = Array.from(impiantiByCommessaId.values()).reduce((sum, list) => sum + (Array.isArray(list) ? list.length : 0), currentImpianti.length || 0);
  const todayKey = new Date().toISOString().slice(0, 10);
  const appRows = [
    ["Stato app", firebaseInitError ? "Attenzione" : "Operativa"], ["Versione installata", installedVersion], ["Ultimo aggiornamento pubblicato", "Verifica disponibile nella sezione aggiornamenti"], ["Ultimo avvio", formatControlCenterDate(performance?.timeOrigin || Date.now())], ["Dispositivo", navigator.userAgent || "Non disponibile"], ["Sistema operativo", navigator.platform || "Non disponibile"], ["Operatore", currentUser?.displayName || currentUser?.email || "Non collegato"], ["UID utente", currentUser?.uid || "-"], ["Ruolo", getControlCenterRoleLabel()], ["Stato login", currentUser ? "Attivo" : "Scaduto"]
  ];
  const cloudRows = [["Firebase Authentication", auth ? "Operativo" : "Errore"], ["Cloud Firestore", db ? "Operativo" : "Errore"], ["Firebase Realtime Database", firebase?.database ? "Operativo" : "Non configurato"], ["Firebase Storage", firebase?.storage ? "Operativo" : "Non configurato"], ["Hosting", "Operativo"], ["Google Drive", driveBridgeState.configured || driveRootFolderId ? "Operativo" : "Non collegato"], ["Servizio notifiche", firebaseMessaging ? "Operativo" : "Non configurato"]];
  const dataRows = ["Commesse", "Impianti", "Squadre", "Ore lavorate", "Segnalazioni", "Note commessa", "Documenti POS", "Mezzi", "Utenti", "Notifiche", "Posizioni operatori"].map((name) => [name, "Ultimo aggiornamento: dati caricati nella sessione corrente"]);
  const pendingExtra = `<div class="control-center-actions"><button class="btn btn-primary" type="button" onclick="syncPendingImpiantoActions(); syncPendingOfflineMutations(); renderControlCenter();">SINCRONIZZA TUTTO</button><button class="btn" type="button" onclick="syncPendingImpiantoActions(); renderControlCenter();">RIPROVA ERRORI</button><button class="btn" type="button">VISUALIZZA DETTAGLI</button>${isAdmin ? '<button class="btn" type="button">ELIMINA OPERAZIONE</button>' : ''}</div><ol class="control-center-list">${pending.map((item) => `<li><strong>${escapeHTML(item.controlType)}</strong><br><span>${escapeHTML(formatControlCenterDate(item.when))} ‚Ä¢ ${escapeHTML(item.operator || "Operatore")}</span><br><em>${escapeHTML(item.status || "In attesa")}</em></li>`).join("") || "<li>Nessuna operazione in attesa.</li>"}</ol>`;
  const usageRows = [["Utenti registrati", platformUsers.length], ["Utenti attivi oggi", platformUsers.filter((u) => String(u.lastSeenAt || u.lastLoginAt || "").includes(todayKey)).length], ["Utenti online ora", platformUsers.filter((u) => Date.now() - firestoreDateToMillis(u.lastSeenAt) <= 10 * 60 * 1000).length], ["Dispositivi collegati", platformUsers.length], ["Numero commesse", commesseById.size], ["Numero impianti", impiantiCount], ["Impianti fatti oggi", currentImpianti.filter((i) => String(i.doneAt || "").includes(todayKey)).length], ["Ore inserite oggi", allHoursReports.filter((r) => String(r.date || r.createdAt || "").includes(todayKey)).length], ["Segnalazioni aperte", "Verifica da archivio segnalazioni"], ["Notifiche non confermate", "Verifica da notifiche"], ["Dati offline in attesa", pending.length]];
  const operatorCards = [buildControlCenterCard("Stato generale dell‚Äôapp", appRows, { color: firebaseInitError ? "yellow" : "green" }), buildControlCenterCard("Stato connessione", [["Stato", quality.status], ["Tipo rete", getNetworkTypeLabel()], ["Velocit√† indicativa", `${navigator.connection?.downlink || "n/d"} Mbps`], ["Qualit√†", quality.label], ["Tempo risposta server", db ? "In verifica" : "Non disponibile"], ["Ultimo online", localStorage.getItem("heraLastOnlineAt") || "Sessione corrente"]], { color: quality.color }), buildControlCenterCard("Dati da sincronizzare", [["Operazioni totali in attesa", pending.length], ["Impianti FATTO offline", pending.filter((i) => i.controlType.includes("Impianto")).length], ["Ore inserite offline", pending.filter((i) => i.controlType.includes("Ore")).length], ["Note salvate offline", pending.filter((i) => i.controlType.includes("Nota")).length], ["Foto da caricare", 0], ["WhatsApp da preparare", pending.filter((i) => i.whatsappStatus !== "sent").length]], { color: pending.length ? "blue" : "green", extra: pendingExtra }), buildControlCenterCard("Controllo aggiornamenti", [["Versione installata", installedVersion], ["Versione disponibile", installedVersion], ["Ultima pubblicazione", "Non configurata"], ["Tipo aggiornamento", "Facoltativo"], ["Note", "L‚Äôapp risulta allineata alla versione configurata"]], { color: "green", extra: '<div class="control-center-actions"><button class="btn" type="button">AGGIORNA APP</button></div>' })];
  const adminCards = isAdmin ? [buildControlCenterCard("Stato cloud", cloudRows, { color: db && auth ? "green" : "red" }), buildControlCenterCard("Ultimo aggiornamento dati", dataRows, { color: "yellow", extra: '<p class="control-center-warning">Attenzione: questi dati non vengono aggiornati da pi√π di 24 ore se la relativa sincronizzazione resta ferma.</p>' }), buildControlCenterCard("Utilizzo dell‚Äôapp", usageRows, { color: "green" }), buildControlCenterCard("Utenti e dispositivi", platformUsers.slice(0, 12).map((u) => [u.displayName || u.email || u.id, `${u.email || "-"} ‚Ä¢ ${adminEmails.has(normalizeEmail(u.email)) ? "Amministratore" : "Operatore"} ‚Ä¢ ${Date.now() - firestoreDateToMillis(u.lastSeenAt) <= 10 * 60 * 1000 ? "Online" : "Offline"}`]), { color: "green" }), buildControlCenterCard("Errori e segnalazioni tecniche", [["Errori salvataggio / Firestore / login / sync", firebaseInitError?.message || "Nessun errore critico registrato"], ["Livelli", "Informazione, Attenzione, Errore, Errore grave"]], { color: firebaseInitError ? "red" : "green", extra: '<div class="control-center-actions"><button class="btn" type="button">RIPROVA</button><button class="btn" type="button">SEGNA COME RISOLTO</button><button class="btn" type="button">COPIA ERRORE</button><button class="btn" type="button">INVIA ASSISTENZA</button><button class="btn" type="button">CANCELLA REGISTRO</button></div>' }), buildControlCenterCard("Controllo sicurezza", [["Tentativi accesso falliti", "Registro non configurato"], ["Utenti bannati", platformUsers.filter((u) => u.banned).length], ["Utenti in attesa", platformUsers.filter((u) => u.pendingApproval).length], ["Sessioni attive", platformUsers.filter((u) => Date.now() - firestoreDateToMillis(u.lastSeenAt) <= 10 * 60 * 1000).length], ["Ultimo backup", "Non configurato"]], { color: "yellow", extra: '<div class="control-center-actions"><button class="btn">GESTISCI UTENTI</button><button class="btn">UTENTI BANNATI</button><button class="btn">RICHIESTE DI ACCESSO</button><button class="btn">SESSIONI ATTIVE</button><button class="btn">REGISTRO ATTIVIT√Ä</button></div>' }), buildControlCenterCard("Backup dati", [["Ultimo backup", "Non configurato"], ["Stato", "Da configurare"], ["Dimensione dati", "n/d"], ["Record salvati", commesseById.size + impiantiCount], ["Destinazione", "Cloud amministratore"], ["Errori", "Nessuno"]], { color: "gray", extra: '<div class="control-center-actions"><button class="btn">ESEGUI BACKUP</button><button class="btn">SCARICA BACKUP</button><button class="btn">RIPRISTINA BACKUP</button><button class="btn">VISUALIZZA BACKUP PRECEDENTI</button></div>' })] : [];
  ui.controlCenterContent.innerHTML = [...operatorCards, ...adminCards].join("");
}

function openControlCenterPage() {
  window.location.hash = "centro-controllo";
  renderControlCenter();
  applyRoute();
  closeSideMenu();
}

function closeControlCenterPage() {
  window.location.hash = "";
  applyRoute();
}

function runControlCenterCheck() {
  const problems = [];
  if (navigator.onLine === false) problems.push("Connessione Internet assente: verifica Wi‚ÄëFi o rete mobile.");
  if (!auth) problems.push("Firebase Authentication non disponibile: controlla configurazione Firebase.");
  if (!db) problems.push("Cloud Firestore non disponibile: controlla SDK e regole.");
  if (firebaseInitError) problems.push(`Errore Firebase: ${firebaseInitError.message}`);
  if (getControlCenterPendingItems().length) problems.push("Sono presenti dati offline: premi SINCRONIZZA TUTTO.");
  const title = problems.length ? "Sono presenti alcuni problemi" : "Tutto funziona correttamente";
  if (ui.controlCenterResults) {
    ui.controlCenterResults.classList.remove("hidden");
    ui.controlCenterResults.innerHTML = `<strong>${escapeHTML(title)}</strong>${problems.length ? `<ul>${problems.map((p) => `<li>${escapeHTML(p)}</li>`).join("")}</ul>` : ""}`;
  }
  renderControlCenter();
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
      "Apri il menu (‚ãÆ) nella home.",
      `Premi ‚Äú${menuTitle}‚Äù.`,
      "Segui i campi/pulsanti del pannello e conferma l'azione."
    ];
    return {
      id: `menu-${buttonId}`,
      domanda: `Come si usa ‚Äú${menuTitle}‚Äù?`,
      rispostaBreve: config.rispostaBreve || `Questa voce apre ‚Äú${menuTitle}‚Äù con tutte le azioni disponibili.`,
      passi: config.passi || fallbackPassi,
      tags: config.tags || ["menu", "funzione"],
      updatedAt: HOWTO_UPDATED_AT
    };
  });
  return [...menuFaqItems, ...STATIC_HOWTO_ITEMS];
}

function openPrivateDocsPage() {
  if (!currentUser) {
    alert("Devi fare login per usare i documenti.");
    return;
  }
  window.location.hash = "documenti";
  applyRoute();
  window.HeraDocuments?.activate?.();
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
  window.HeraDocuments?.deactivate?.();
  window.location.hash = "";
  applyRoute();
}

const CALENDAR_EVENT_TYPES = {
  ferie: { label: "Ferie", icon: "üèñÔ∏è" },
  permesso: { label: "Permesso", icon: "üïí" },
  malattia: { label: "Malattia", icon: "ü§í" },
  intervento: { label: "Programmazione intervento", icon: "üõ†Ô∏è" },
  riunione: { label: "Riunione", icon: "üë•" },
  formazione: { label: "Formazione", icon: "üéì" },
  scadenza: { label: "Scadenza", icon: "‚è∞" },
  altro: { label: "Altro", icon: "üìå" }
};

function formatCalendarDateKey(date) {
  const value = date instanceof Date ? date : new Date(date);
  if (Number.isNaN(value.getTime())) return "";
  const year = value.getFullYear();
  const month = String(value.getMonth() + 1).padStart(2, "0");
  const day = String(value.getDate()).padStart(2, "0");
  return `${year}-${month}-${day}`;
}

function parseCalendarDateKey(dateKey) {
  const match = String(dateKey || "").match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (!match) return null;
  const date = new Date(Number(match[1]), Number(match[2]) - 1, Number(match[3]), 12);
  return Number.isNaN(date.getTime()) ? null : date;
}

function formatCalendarLongDate(dateKey) {
  const date = parseCalendarDateKey(dateKey);
  return date
    ? new Intl.DateTimeFormat("it-IT", { weekday: "long", day: "numeric", month: "long", year: "numeric" }).format(date)
    : "Data non disponibile";
}

function openCalendarPage() {
  if (!currentUser) {
    alert("Devi fare login per aprire il calendario.");
    return;
  }
  calendarMode = "choice";
  window.location.hash = "calendario";
  applyRoute();
  closeSideMenu();
}

function closeCalendarPage() {
  closeCalendarEventForm();
  calendarMode = "choice";
  window.location.hash = "";
  applyRoute();
}

function setCalendarMode(mode) {
  if (mode !== "hours" && mode !== "shared") return;
  calendarMode = mode;
  renderCalendarMode();
  renderCalendar();
}

function renderCalendarMode() {
  const isChoice = calendarMode === "choice";
  const isHours = calendarMode === "hours";
  ui.calendarChoiceCard?.classList.toggle("hidden", !isChoice);
  ui.calendarHeroCard?.classList.toggle("hidden", isChoice);
  ui.calendarMainCard?.classList.toggle("hidden", isChoice);
  ui.calendarDayCard?.classList.toggle("hidden", isChoice);
  ui.calendarNewEventBtn?.classList.toggle("hidden", isHours);
  ui.calendarAddSelectedDayBtn?.classList.toggle("hidden", isHours);
  ui.calendarHoursTab?.classList.toggle("is-active", isHours);
  ui.calendarSharedTab?.classList.toggle("is-active", calendarMode === "shared");
  ui.calendarHoursTab?.setAttribute("aria-selected", String(isHours));
  ui.calendarSharedTab?.setAttribute("aria-selected", String(calendarMode === "shared"));
  if (ui.calendarPageHeading) ui.calendarPageHeading.textContent = isHours ? "üïí Le mie ore" : "üóìÔ∏è Calendario condiviso";
  if (ui.calendarPageDescription) {
    ui.calendarPageDescription.textContent = isHours
      ? "Ore lavorate personali recuperate dalla Gestione ore."
      : "Ferie, permessi, malattie, interventi e altre informazioni visibili a tutti gli utenti.";
  }
  if (ui.calendarGrid) ui.calendarGrid.setAttribute("aria-label", isHours ? "Calendario mensile delle mie ore" : "Calendario mensile condiviso");
}

function getPersonalHoursRowsForDate(dateKey) {
  if (!currentUser || !dateKey) return [];
  const identity = getCurrentUserSquadraIdentity();
  const rows = [];
  allHoursReports.forEach((report) => {
    if (normalizeHoursReportDateKey(report.date) !== dateKey) return;
    (Array.isArray(report.entries) ? report.entries : []).forEach((entry) => {
      (Array.isArray(entry.rows) ? entry.rows : []).forEach((row) => {
        const operatorId = String(row.operatoreId || row.personaleId || "").replace(/^utente:/, "");
        const matches = doesSquadraMemberMatchCurrentUser({
          id: operatorId,
          uid: row.uid || row.userId || "",
          email: row.email || "",
          name: row.operatore || row.nome || row.name || ""
        }, identity);
        const hours = Number(row.ore || 0);
        if (matches && Number.isFinite(hours) && hours > 0) rows.push({ report, entry, row, hours });
      });
    });
  });
  return rows;
}

function formatPersonalHours(hours) {
  const minutes = Math.round(Number(hours || 0) * 60);
  return `${String(Math.floor(minutes / 60)).padStart(2, "0")}:${String(minutes % 60).padStart(2, "0")}`;
}

function subscribeCalendarEvents() {
  if (!currentUser || !db || unsubscribeCalendarEvents) return;
  if (ui.calendarFeedback) ui.calendarFeedback.textContent = "Caricamento eventi...";
  const visibleEvents = new Map();
  const subscriptions = [];
  const applySnapshot = (snapshot) => {
    snapshot.docChanges().forEach((change) => {
      if (change.type === "removed") visibleEvents.delete(change.doc.id);
      else visibleEvents.set(change.doc.id, { id: change.doc.id, ...change.doc.data() });
    });
    calendarEvents = Array.from(visibleEvents.values());
    calendarAbsenceCache.clear();
    if (ui.calendarFeedback) {
      ui.calendarFeedback.textContent = calendarEvents.length
        ? `${calendarEvents.length} ${calendarEvents.length === 1 ? "evento condiviso" : "eventi condivisi"}`
        : "Nessun evento inserito.";
    }
    renderCalendar();
  };
  const handleError = (error) => {
    console.error("Errore caricamento calendario condiviso:", error);
    if (ui.calendarFeedback) ui.calendarFeedback.textContent = "Impossibile caricare gli eventi. Verifica la connessione e i permessi.";
  };
  // Keep private document expirations out of the broad shared-calendar query.
  // Firestore then validates each privacy-scoped query against the same rules used
  // for direct reads, so an administrator cannot enumerate personal expirations.
  const events = db.collection("calendarEvents");
  subscriptions.push(events.where("type", "!=", "SCADENZA_DOCUMENTO").onSnapshot(applySnapshot, handleError));
  subscriptions.push(events.where("ownerUserId", "==", currentUser.uid).onSnapshot(applySnapshot, handleError));
  subscriptions.push(events.where("authorizedUserIds", "array-contains", currentUser.uid).onSnapshot(applySnapshot, handleError));
  subscriptions.push(events.where("sharedToAll", "==", true).onSnapshot(applySnapshot, handleError));
  unsubscribeCalendarEvents = () => subscriptions.forEach((unsubscribe) => unsubscribe());
}

function stopCalendarEventsSubscription() {
  if (unsubscribeCalendarEvents) {
    unsubscribeCalendarEvents();
    unsubscribeCalendarEvents = null;
  }
}

function calendarEventIncludesDate(event, dateKey) {
  const start = String(event.startDate || "");
  const end = String(event.endDate || start);
  return Boolean(dateKey && start && start <= dateKey && dateKey <= end);
}

function getCalendarEventsForDate(dateKey) {
  return calendarEvents
    .filter((event) => calendarEventIncludesDate(event, dateKey))
    .sort((a, b) => {
      const aTime = a.allDay === false ? String(a.startTime || "23:59") : "00:00";
      const bTime = b.allDay === false ? String(b.startTime || "23:59") : "00:00";
      return aTime.localeCompare(bTime) || String(a.title || "").localeCompare(String(b.title || ""), "it");
    });
}

function renderCalendar() {
  renderCalendarMode();
  if (calendarMode === "choice") return;
  if (calendarMode === "hours") {
    renderPersonalHoursCalendar();
    return;
  }
  if (!ui.calendarGrid || !ui.calendarMonthTitle) return;
  const year = calendarVisibleMonth.getFullYear();
  const month = calendarVisibleMonth.getMonth();
  ui.calendarMonthTitle.textContent = new Intl.DateTimeFormat("it-IT", { month: "long", year: "numeric" }).format(calendarVisibleMonth);
  const firstDay = new Date(year, month, 1, 12);
  const mondayOffset = (firstDay.getDay() + 6) % 7;
  const gridStart = new Date(year, month, 1 - mondayOffset, 12);
  const todayKey = formatCalendarDateKey(new Date());
  const cells = [];

  for (let index = 0; index < 42; index += 1) {
    const date = new Date(gridStart);
    date.setDate(gridStart.getDate() + index);
    const dateKey = formatCalendarDateKey(date);
    const events = getCalendarEventsForDate(dateKey);
    const typeDots = [...new Set(events.map((event) => String(event.type || "altro")))]
      .slice(0, 3)
      .map((type) => `<span class="calendar-type-dot calendar-type-${escapeHTML(type)}"></span>`)
      .join("");
    const classes = [
      "calendar-day",
      date.getMonth() === month ? "" : "is-outside",
      dateKey === todayKey ? "is-today" : "",
      dateKey === calendarSelectedDate ? "is-selected" : "",
      events.length ? "has-events" : ""
    ].filter(Boolean).join(" ");
    const eventLabel = events.length ? `${events.length} ${events.length === 1 ? "evento" : "eventi"}` : "nessun evento";
    cells.push(`
      <button class="${classes}" type="button" role="gridcell" data-calendar-date="${dateKey}" aria-label="${escapeHTML(formatCalendarLongDate(dateKey))}, ${eventLabel}">
        <span class="calendar-day-number">${date.getDate()}</span>
        ${events.length ? `<span class="calendar-event-count">${events.length}</span>` : ""}
        <span class="calendar-type-dots">${typeDots}</span>
      </button>
    `);
  }

  ui.calendarGrid.innerHTML = cells.join("");
  ui.calendarGrid.querySelectorAll("[data-calendar-date]").forEach((button) => {
    button.addEventListener("click", () => selectCalendarDate(button.dataset.calendarDate || ""));
  });
  renderCalendarSelectedDay();
}

function renderPersonalHoursCalendar() {
  if (!ui.calendarGrid || !ui.calendarMonthTitle) return;
  const year = calendarVisibleMonth.getFullYear();
  const month = calendarVisibleMonth.getMonth();
  ui.calendarMonthTitle.textContent = new Intl.DateTimeFormat("it-IT", { month: "long", year: "numeric" }).format(calendarVisibleMonth);
  const firstDay = new Date(year, month, 1, 12);
  const mondayOffset = (firstDay.getDay() + 6) % 7;
  const gridStart = new Date(year, month, 1 - mondayOffset, 12);
  const todayKey = formatCalendarDateKey(new Date());
  const cells = [];
  for (let index = 0; index < 42; index += 1) {
    const date = new Date(gridStart);
    date.setDate(gridStart.getDate() + index);
    const dateKey = formatCalendarDateKey(date);
    const total = getPersonalHoursRowsForDate(dateKey).reduce((sum, item) => sum + item.hours, 0);
    const classes = [
      "calendar-day", "personal-hours-day",
      date.getMonth() === month ? "" : "is-outside",
      dateKey === todayKey ? "is-today" : "",
      dateKey === calendarSelectedDate ? "is-selected" : "",
      total > 0 ? "has-hours" : ""
    ].filter(Boolean).join(" ");
    const hoursLabel = total > 0 ? `${formatPersonalHours(total)} ore` : "nessuna ora";
    cells.push(`
      <button class="${classes}" type="button" role="gridcell" data-calendar-date="${dateKey}" aria-label="${escapeHTML(formatCalendarLongDate(dateKey))}, ${hoursLabel}">
        <span class="calendar-day-number">${date.getDate()}</span>
        ${total > 0 ? `<span class="calendar-personal-hours">${formatPersonalHours(total)} <small>ore</small></span>` : ""}
      </button>
    `);
  }
  ui.calendarGrid.innerHTML = cells.join("");
  ui.calendarGrid.querySelectorAll("[data-calendar-date]").forEach((button) => {
    button.addEventListener("click", () => selectCalendarDate(button.dataset.calendarDate || ""));
  });
  if (ui.calendarFeedback) {
    ui.calendarFeedback.textContent = hoursReportsLoaded
      ? "Sono mostrate esclusivamente le ore dell‚Äôoperatore connesso."
      : "Caricamento delle ore personali...";
  }
  renderPersonalHoursSelectedDay();
}

function renderPersonalHoursSelectedDay() {
  if (!ui.calendarDayEvents) return;
  const rows = getPersonalHoursRowsForDate(calendarSelectedDate);
  const total = rows.reduce((sum, item) => sum + item.hours, 0);
  if (ui.calendarSelectedDayTitle) ui.calendarSelectedDayTitle.textContent = formatCalendarLongDate(calendarSelectedDate);
  if (ui.calendarSelectedDaySummary) {
    ui.calendarSelectedDaySummary.textContent = total > 0
      ? `Totale personale: ${formatPersonalHours(total)} ore`
      : "Nessuna ora personale inserita";
  }
  if (!rows.length) {
    ui.calendarDayEvents.innerHTML = "<div class='calendar-empty-day'><span>üïí</span><p>Non risultano ore personali per questa giornata.</p></div>";
    return;
  }
  ui.calendarDayEvents.innerHTML = rows.map(({ entry, row, hours }) => `
    <article class="calendar-event-card personal-hours-detail">
      <div class="calendar-event-heading">
        <span class="calendar-event-icon" aria-hidden="true">üïí</span>
        <div><h3>${escapeHTML(entry.commessaName || commesseById.get(entry.commessaId)?.nome || "Ore lavorate")}</h3>
        <p>${formatPersonalHours(hours)} ore</p></div>
      </div>
      ${entry.note ? `<p class="calendar-event-description">${escapeHTML(entry.note)}</p>` : ""}
    </article>
  `).join("");
}

function selectCalendarDate(dateKey) {
  const date = parseCalendarDateKey(dateKey);
  if (!date) return;
  calendarSelectedDate = dateKey;
  if (date.getMonth() !== calendarVisibleMonth.getMonth() || date.getFullYear() !== calendarVisibleMonth.getFullYear()) {
    calendarVisibleMonth = new Date(date.getFullYear(), date.getMonth(), 1, 12);
  }
  renderCalendar();
}

function changeCalendarMonth(offset) {
  calendarVisibleMonth = new Date(calendarVisibleMonth.getFullYear(), calendarVisibleMonth.getMonth() + Number(offset || 0), 1, 12);
  calendarSelectedDate = formatCalendarDateKey(calendarVisibleMonth);
  renderCalendar();
}

function showCalendarToday() {
  const today = new Date();
  calendarVisibleMonth = new Date(today.getFullYear(), today.getMonth(), 1, 12);
  calendarSelectedDate = formatCalendarDateKey(today);
  renderCalendar();
}

function canModifyCalendarEvent(event) {
  if (event?.type === "SCADENZA_DOCUMENTO") return Boolean(currentUser && String(event?.ownerUserId || "") === String(currentUser.uid || ""));
  return Boolean(currentUser && (canManageData() || String(event?.createdByUid || "") === String(currentUser.uid || "")));
}

function formatCalendarEventPeriod(event) {
  const allDay = event.allDay !== false;
  const startDate = String(event.startDate || "");
  const endDate = String(event.endDate || startDate);
  if (allDay) return startDate === endDate ? "Tutto il giorno" : `Dal ${startDate} al ${endDate}`;
  const time = [event.startTime, event.endTime].filter(Boolean).join("‚Äì");
  return startDate === endDate ? (time || "Orario da definire") : `Dal ${startDate} al ${endDate}${time ? ` ‚Ä¢ ${time}` : ""}`;
}

function renderCalendarSelectedDay() {
  if (!ui.calendarDayEvents) return;
  const events = getCalendarEventsForDate(calendarSelectedDate);
  if (ui.calendarSelectedDayTitle) ui.calendarSelectedDayTitle.textContent = formatCalendarLongDate(calendarSelectedDate);
  if (ui.calendarSelectedDaySummary) {
    ui.calendarSelectedDaySummary.textContent = events.length
      ? `${events.length} ${events.length === 1 ? "evento programmato" : "eventi programmati"}`
      : "Nessun evento in questo giorno";
  }
  if (!events.length) {
    ui.calendarDayEvents.innerHTML = "<div class='calendar-empty-day'><span>üóìÔ∏è</span><p>Nessun evento. Premi ‚ÄúAggiungi‚Äù per inserirne uno.</p></div>";
    return;
  }
  ui.calendarDayEvents.innerHTML = events.map((event) => {
    const isDocumentExpiration = event.type === "SCADENZA_DOCUMENTO";
    const type = isDocumentExpiration ? { icon: "üìÑ", label: "Scadenza documento" } : (CALENDAR_EVENT_TYPES[event.type] || CALENDAR_EVENT_TYPES.altro);
    const mayModify = canModifyCalendarEvent(event);
    const safeLink = /^https?:\/\//i.test(String(event.link || "")) ? String(event.link) : "";
    const detailRows = [
      event.worksite ? `<p><strong>Commessa / impianto:</strong> ${escapeHTML(event.worksite)}</p>` : "",
      event.location ? `<p><strong>Luogo:</strong> ${escapeHTML(event.location)}</p>` : "",
      event.participants ? `<p><strong>Persone:</strong> ${escapeHTML(event.participants)}</p>` : "",
      event.description ? `<p class="calendar-event-description">${escapeHTML(event.description)}</p>` : "",
      safeLink ? `<p><a class="btn calendar-link-btn" href="${escapeHTML(safeLink)}" target="_blank" rel="noopener noreferrer">üîó Apri link</a></p>` : ""
    ].join("");
    return `
      <article class="calendar-event-card calendar-type-border-${escapeHTML(event.type || "altro")}">
        <div class="calendar-event-heading">
          <span class="calendar-event-icon" aria-hidden="true">${type.icon}</span>
          <div>
            <span class="calendar-event-type">${escapeHTML(type.label)}</span>
            <h3>${escapeHTML(isDocumentExpiration ? (event.compactTitle || event.title || "Documento") : (event.title || "Evento"))}</h3>
            <p class="calendar-event-period">${escapeHTML(formatCalendarEventPeriod(event))}</p>
          </div>
        </div>
        <div class="calendar-event-details">${detailRows}</div>
        <div class="calendar-event-footer">
          <span>Inserito da <strong>${escapeHTML(event.createdByName || event.createdByEmail || "Utente")}</strong></span>
          ${mayModify ? `
            <span class="calendar-event-actions">
              ${isDocumentExpiration ? `<button class="btn" type="button" data-calendar-document="${escapeHTML(event.documentId || "")}">Apri documento</button>` : `<button class="btn" type="button" data-calendar-edit="${escapeHTML(event.id)}">Modifica</button>`}
              ${isDocumentExpiration ? "" : `<button class="btn btn-danger" type="button" data-calendar-delete="${escapeHTML(event.id)}">Elimina</button>`}
            </span>
          ` : ""}
        </div>
      </article>
    `;
  }).join("");
  ui.calendarDayEvents.querySelectorAll("[data-calendar-edit]").forEach((button) => {
    button.addEventListener("click", () => {
      const event = calendarEvents.find((item) => item.id === button.dataset.calendarEdit);
      if (event) openCalendarEventForm(event.startDate, event);
    });
  });
  ui.calendarDayEvents.querySelectorAll("[data-calendar-document]").forEach((button) => {
    button.addEventListener("click", () => window.HeraDocuments?.open({ visibility: "personal" }));
  });
  ui.calendarDayEvents.querySelectorAll("[data-calendar-delete]").forEach((button) => {
    button.addEventListener("click", () => deleteCalendarEvent(button.dataset.calendarDelete || ""));
  });
}

function syncCalendarTimeFields() {
  const allDay = Boolean(ui.calendarEventAllDay?.checked);
  ui.calendarEventTimeFields?.classList.toggle("hidden", allDay);
  if (ui.calendarEventStartTime) ui.calendarEventStartTime.required = !allDay;
  if (ui.calendarEventEndTime) ui.calendarEventEndTime.required = false;
}

function getCalendarParticipantSnapshot(person = null, freeName = "") {
  const name = String(person ? getPersonaleDisplayName(person) : freeName).trim();
  return {
    id: String(person?.id || "").trim(),
    name,
    email: String(person?.email || "").trim(),
    freeText: !person
  };
}

function addCalendarParticipant(person = null, freeName = "") {
  const participant = getCalendarParticipantSnapshot(person, freeName);
  if (!participant.name) return;
  const key = normalizeSquadraMemberIdentity(participant.id || participant.email || participant.name);
  if (calendarSelectedParticipants.some((item) => normalizeSquadraMemberIdentity(item.id || item.email || item.name) === key)) return;
  calendarSelectedParticipants.push(participant);
  renderCalendarParticipantPicker();
}

function removeCalendarParticipant(index) {
  calendarSelectedParticipants.splice(index, 1);
  renderCalendarParticipantPicker();
}

function renderCalendarParticipantPicker() {
  if (!ui.calendarParticipantsChips || !ui.calendarEventParticipants) return;
  ui.calendarParticipantsChips.innerHTML = calendarSelectedParticipants.map((participant, index) => `
    <span class="calendar-participant-chip">
      <span>${escapeHTML(participant.name)}</span>
      <button type="button" data-calendar-participant-remove="${index}" aria-label="Rimuovi ${escapeHTML(participant.name)}">√ó</button>
    </span>
  `).join("");
  ui.calendarParticipantsChips.querySelectorAll("[data-calendar-participant-remove]").forEach((button) => {
    button.addEventListener("click", () => removeCalendarParticipant(Number(button.dataset.calendarParticipantRemove)));
  });
  ui.calendarEventParticipants.value = calendarSelectedParticipants.map((participant) => participant.name).join(", ");
}

function renderCalendarParticipantSuggestions() {
  if (!ui.calendarParticipantsSuggestions || !ui.calendarParticipantsSearch) return;
  const query = normalizeSquadraMemberIdentity(ui.calendarParticipantsSearch.value);
  const selectedKeys = new Set(calendarSelectedParticipants.map((item) => normalizeSquadraMemberIdentity(item.id || item.email || item.name)));
  const matches = personaleRecords
    .filter((person) => {
      const key = normalizeSquadraMemberIdentity(person.id || person.email || getPersonaleDisplayName(person));
      if (selectedKeys.has(key)) return false;
      return !query || normalizeSquadraMemberIdentity(`${getPersonaleDisplayName(person)} ${person.email || ""}`).includes(query);
    })
    .slice(0, 8);
  const freeValue = String(ui.calendarParticipantsSearch.value || "").trim();
  ui.calendarParticipantsSuggestions.innerHTML = [
    ...matches.map((person) => `<button type="button" role="option" data-calendar-person-id="${escapeHTML(person.id)}"><strong>${escapeHTML(getPersonaleDisplayName(person))}</strong>${person.email ? `<small>${escapeHTML(person.email)}</small>` : ""}</button>`),
    freeValue && !matches.some((person) => normalizeSquadraMemberIdentity(getPersonaleDisplayName(person)) === normalizeSquadraMemberIdentity(freeValue))
      ? `<button type="button" role="option" data-calendar-free-person="${escapeHTML(freeValue)}">Ôºã Usa nome libero: <strong>${escapeHTML(freeValue)}</strong></button>`
      : ""
  ].join("");
  ui.calendarParticipantsSuggestions.classList.toggle("hidden", !matches.length && !freeValue);
  ui.calendarParticipantsSuggestions.querySelectorAll("[data-calendar-person-id]").forEach((button) => {
    button.addEventListener("click", () => {
      const person = personaleRecords.find((item) => item.id === button.dataset.calendarPersonId);
      if (person) addCalendarParticipant(person);
      ui.calendarParticipantsSearch.value = "";
      renderCalendarParticipantSuggestions();
      ui.calendarParticipantsSearch.focus();
    });
  });
  ui.calendarParticipantsSuggestions.querySelectorAll("[data-calendar-free-person]").forEach((button) => {
    button.addEventListener("click", () => {
      addCalendarParticipant(null, button.dataset.calendarFreePerson || "");
      ui.calendarParticipantsSearch.value = "";
      renderCalendarParticipantSuggestions();
      ui.calendarParticipantsSearch.focus();
    });
  });
}

function handleCalendarParticipantSearchKeydown(event) {
  if (event.key === "Escape") {
    ui.calendarParticipantsSuggestions?.classList.add("hidden");
    return;
  }
  if (event.key !== "Enter" && event.key !== ",") return;
  event.preventDefault();
  const freeValue = String(ui.calendarParticipantsSearch?.value || "").trim();
  if (!freeValue) return;
  const exactPerson = personaleRecords.find((person) => normalizeSquadraMemberIdentity(getPersonaleDisplayName(person)) === normalizeSquadraMemberIdentity(freeValue));
  addCalendarParticipant(exactPerson || null, exactPerson ? "" : freeValue);
  ui.calendarParticipantsSearch.value = "";
  renderCalendarParticipantSuggestions();
}

function populateCalendarCommesse(selectedId = "", customName = "") {
  if (!ui.calendarEventCommessa) return;
  const commesse = sortCommesseByCreatedAtDesc(Array.from(commesseById.values()));
  ui.calendarEventCommessa.innerHTML = [
    '<option value="">Nessuna commessa</option>',
    ...commesse.map((commessa) => `<option value="${escapeHTML(commessa.id)}">${escapeHTML(getCommessaDisplayName(commessa))}</option>`),
    '<option value="__custom">Ôºã Scrivi una commessa non presente</option>'
  ].join("");
  if (selectedId && commesseById.has(selectedId)) ui.calendarEventCommessa.value = selectedId;
  else if (customName) ui.calendarEventCommessa.value = "__custom";
  else ui.calendarEventCommessa.value = "";
  ui.calendarEventCustomCommessa.value = customName || "";
  ui.calendarEventCustomCommessaField?.classList.toggle("hidden", ui.calendarEventCommessa.value !== "__custom");
}

function getCalendarImpiantoDisplayName(impianto = {}) {
  return String(impianto.denominazione || impianto.nome || impianto.idSap || "Impianto").trim();
}

function getCalendarImpiantoLocation(impianto = {}) {
  return [impianto.indirizzo || impianto.descrizioneVia, impianto.comune].map((value) => String(value || "").trim()).filter(Boolean).join(", ");
}

function populateCalendarImpianti(commessaId = "", selectedId = "", customName = "") {
  if (!ui.calendarEventImpianto) return;
  const impianti = commessaId ? getCommessaCachedImpianti(commessaId) : [];
  ui.calendarEventImpianto.innerHTML = [
    '<option value="">Nessun impianto</option>',
    ...impianti.map((impianto) => `<option value="${escapeHTML(getSquadraImpiantoId(impianto))}">${escapeHTML(getCalendarImpiantoDisplayName(impianto))}</option>`),
    '<option value="__custom">Ôºã Scrivi un impianto non presente</option>'
  ].join("");
  const hasCustomCommessa = ui.calendarEventCommessa?.value === "__custom";
  ui.calendarEventImpianto.disabled = !commessaId && !customName && !hasCustomCommessa;
  if (selectedId && impianti.some((impianto) => getSquadraImpiantoId(impianto) === selectedId)) ui.calendarEventImpianto.value = selectedId;
  else if (customName) ui.calendarEventImpianto.value = "__custom";
  else ui.calendarEventImpianto.value = "";
  ui.calendarEventCustomImpianto.value = customName || "";
  ui.calendarEventCustomImpiantoField?.classList.toggle("hidden", ui.calendarEventImpianto.value !== "__custom");
}

function handleCalendarCommessaChange() {
  const commessaId = String(ui.calendarEventCommessa?.value || "");
  const isCustom = commessaId === "__custom";
  ui.calendarEventCustomCommessaField?.classList.toggle("hidden", !isCustom);
  if (!isCustom) ui.calendarEventCustomCommessa.value = "";
  populateCalendarImpianti(isCustom ? "" : commessaId);
}

function handleCalendarImpiantoChange() {
  const commessaId = String(ui.calendarEventCommessa?.value || "");
  const impiantoId = String(ui.calendarEventImpianto?.value || "");
  const isCustom = impiantoId === "__custom";
  ui.calendarEventCustomImpiantoField?.classList.toggle("hidden", !isCustom);
  if (!isCustom) ui.calendarEventCustomImpianto.value = "";
  const impianto = getCommessaCachedImpianti(commessaId).find((item) => getSquadraImpiantoId(item) === impiantoId);
  if (impianto && ui.calendarEventLocation) {
    const location = getCalendarImpiantoLocation(impianto);
    if (location) ui.calendarEventLocation.value = location;
  }
}

function openCalendarEventForm(dateKey = calendarSelectedDate, event = null) {
  if (!currentUser) return;
  const fallbackDate = parseCalendarDateKey(dateKey) ? dateKey : formatCalendarDateKey(new Date());
  ui.calendarEventForm?.reset();
  ui.calendarEventId.value = event?.id || "";
  ui.calendarEventFormTitle.textContent = event ? "Modifica evento" : "Nuovo evento";
  ui.calendarEventType.value = event?.type || "ferie";
  ui.calendarEventTitle.value = event?.title || "";
  ui.calendarEventStartDate.value = event?.startDate || fallbackDate;
  ui.calendarEventEndDate.value = event?.endDate || event?.startDate || fallbackDate;
  ui.calendarEventAllDay.checked = event?.allDay !== false;
  ui.calendarEventStartTime.value = event?.startTime || "";
  ui.calendarEventEndTime.value = event?.endTime || "";
  const initialCommessaId = event?.commessaId || (!event ? selectedCommessaId : "");
  populateCalendarCommesse(initialCommessaId, event?.customCommessa || (!event?.commessaId ? event?.commessaName || event?.worksite || "" : ""));
  populateCalendarImpianti(initialCommessaId, event?.impiantoId || "", event?.customImpianto || (!event?.impiantoId ? event?.impiantoName || "" : ""));
  ui.calendarEventLocation.value = event?.location || "";
  calendarSelectedParticipants = Array.isArray(event?.participantSnapshots)
    ? event.participantSnapshots.map((participant) => ({
      id: String(participant?.id || ""),
      name: String(participant?.name || ""),
      email: String(participant?.email || ""),
      freeText: Boolean(participant?.freeText)
    })).filter((participant) => participant.name)
    : parseMultiEntryValue(event?.participants || "").map((name) => getCalendarParticipantSnapshot(null, name));
  if (!event && !calendarSelectedParticipants.length) {
    const currentPerson = getPersonaleByLoginEmail();
    addCalendarParticipant(currentPerson, currentPerson ? "" : getCurrentUserResolvedName("Utente"));
  } else {
    renderCalendarParticipantPicker();
  }
  if (ui.calendarParticipantsSearch) ui.calendarParticipantsSearch.value = "";
  ui.calendarParticipantsSuggestions?.classList.add("hidden");
  ui.calendarEventDescription.value = event?.description || "";
  ui.calendarEventLink.value = event?.link || "";
  ui.calendarEventFormFeedback.textContent = "";
  syncCalendarTimeFields();
  if (typeof ui.calendarEventDialog.showModal === "function") ui.calendarEventDialog.showModal();
  else ui.calendarEventDialog.setAttribute("open", "");
  setTimeout(() => ui.calendarEventTitle?.focus(), 50);
}

function closeCalendarEventForm() {
  if (!ui.calendarEventDialog) return;
  if (typeof ui.calendarEventDialog.close === "function" && ui.calendarEventDialog.open) ui.calendarEventDialog.close();
  else ui.calendarEventDialog.removeAttribute("open");
}

function getCalendarAuthorName() {
  const profile = platformUsers.find((user) => String(user.id || user.uid || "") === String(currentUser?.uid || ""));
  return String(profile?.displayName || profile?.nome || currentUser?.displayName || currentUser?.email || "Utente").trim();
}

async function saveCalendarEvent(event) {
  event.preventDefault();
  if (!currentUser || !db) return;
  const eventId = String(ui.calendarEventId.value || "").trim();
  const existing = calendarEvents.find((item) => item.id === eventId);
  if (existing && !canModifyCalendarEvent(existing)) {
    ui.calendarEventFormFeedback.textContent = "Non puoi modificare un evento inserito da un altro utente.";
    return;
  }
  const startDate = String(ui.calendarEventStartDate.value || "");
  const endDate = String(ui.calendarEventEndDate.value || startDate);
  const allDay = Boolean(ui.calendarEventAllDay.checked);
  const startTime = allDay ? "" : String(ui.calendarEventStartTime.value || "");
  const endTime = allDay ? "" : String(ui.calendarEventEndTime.value || "");
  if (!startDate || !endDate || endDate < startDate) {
    ui.calendarEventFormFeedback.textContent = "Controlla le date: la data finale non pu√≤ precedere quella iniziale.";
    return;
  }
  if (!allDay && !startTime) {
    ui.calendarEventFormFeedback.textContent = "Inserisci almeno l'ora di inizio.";
    return;
  }
  if (!allDay && startDate === endDate && endTime && endTime < startTime) {
    ui.calendarEventFormFeedback.textContent = "L'ora finale non pu√≤ precedere quella iniziale.";
    return;
  }
  const commessaSelection = String(ui.calendarEventCommessa.value || "");
  const customCommessa = commessaSelection === "__custom" ? String(ui.calendarEventCustomCommessa.value || "").trim() : "";
  const commessa = commessaSelection && commessaSelection !== "__custom" ? commesseById.get(commessaSelection) : null;
  const impiantoSelection = String(ui.calendarEventImpianto.value || "");
  const customImpianto = impiantoSelection === "__custom" ? String(ui.calendarEventCustomImpianto.value || "").trim() : "";
  const impianto = commessa
    ? getCommessaCachedImpianti(commessa.id).find((item) => getSquadraImpiantoId(item) === impiantoSelection)
    : null;
  const eventType = String(ui.calendarEventType.value || "altro");
  const generatedTitle = `${CALENDAR_EVENT_TYPES[eventType]?.label || "Evento"}${calendarSelectedParticipants[0]?.name ? ` ‚Ä¢ ${calendarSelectedParticipants[0].name}` : ""}`;
  const payload = {
    type: eventType,
    title: String(ui.calendarEventTitle.value || "").trim() || generatedTitle,
    titleWasGenerated: !String(ui.calendarEventTitle.value || "").trim(),
    startDate,
    endDate,
    allDay,
    startTime,
    endTime,
    commessaId: commessa?.id || "",
    commessaName: String(commessa?.nome || customCommessa || "").trim(),
    customCommessa,
    impiantoId: impianto ? getSquadraImpiantoId(impianto) : "",
    impiantoName: String(impianto ? getCalendarImpiantoDisplayName(impianto) : customImpianto).trim(),
    customImpianto,
    worksite: [commessa?.nome || customCommessa, impianto ? getCalendarImpiantoDisplayName(impianto) : customImpianto].filter(Boolean).join(" ‚Ä¢ "),
    location: String(ui.calendarEventLocation.value || "").trim(),
    participants: calendarSelectedParticipants.map((participant) => participant.name).join(", "),
    participantIds: calendarSelectedParticipants.map((participant) => participant.id).filter(Boolean),
    participantEmails: calendarSelectedParticipants.map((participant) => participant.email).filter(Boolean),
    participantSnapshots: calendarSelectedParticipants,
    description: String(ui.calendarEventDescription.value || "").trim(),
    link: String(ui.calendarEventLink.value || "").trim(),
    updatedAt: firebase.firestore.FieldValue.serverTimestamp(),
    updatedByUid: currentUser.uid || "",
    updatedByEmail: currentUser.email || ""
  };
  ui.calendarEventSaveBtn.disabled = true;
  ui.calendarEventFormFeedback.textContent = "Salvataggio...";
  try {
    if (eventId) {
      await db.collection("calendarEvents").doc(eventId).set(payload, { merge: true });
    } else {
      await db.collection("calendarEvents").add({
        ...payload,
        createdByUid: currentUser.uid || "",
        createdByEmail: currentUser.email || "",
        createdByName: getCalendarAuthorName(),
        createdAt: firebase.firestore.FieldValue.serverTimestamp()
      });
    }
    calendarSelectedDate = startDate;
    const savedDate = parseCalendarDateKey(startDate);
    if (savedDate) calendarVisibleMonth = new Date(savedDate.getFullYear(), savedDate.getMonth(), 1, 12);
    closeCalendarEventForm();
  } catch (error) {
    console.error("Salvataggio evento calendario non riuscito:", error);
    ui.calendarEventFormFeedback.textContent = error?.message || "Errore durante il salvataggio dell'evento.";
  } finally {
    ui.calendarEventSaveBtn.disabled = false;
  }
}

async function deleteCalendarEvent(eventId) {
  const event = calendarEvents.find((item) => item.id === eventId);
  if (!event || !canModifyCalendarEvent(event)) {
    alert("Pu√≤ eliminare questo evento solo chi lo ha inserito o un amministratore.");
    return;
  }
  if (!window.confirm(`Eliminare l'evento ‚Äú${event.title || "Evento"}‚Äù?`)) return;
  try {
    await db.collection("calendarEvents").doc(eventId).delete();
  } catch (error) {
    console.error("Eliminazione evento calendario non riuscita:", error);
    alert(error?.message || "Impossibile eliminare l'evento.");
  }
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
  ui.addHoursCommessaBtn?.classList.remove("hidden");
  ui.hoursFinalizeBtn?.classList.remove("hidden");
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
    group.innerHTML = `<h3>üìÅ ${escapeHTML(category)}</h3>`;
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
    meta.textContent = `Ordine: ${Number(doc.order || 0)} ‚Ä¢ ${doc.active === false ? "Non attivo" : "Attivo"}`;
    card.append(title, description, meta, actions);
    return card;
  }
  card.append(title, description, actions);
  return card;
}

function openPosDocumentForm(doc = null) {
  if (!canManageData()) return;
  ui.posDocumentForm?.classList.remove("hidden");
  if (ui.posAddToggleBtn) ui.posAddToggleBtn.textContent = doc ? "Modifica documento" : "‚ûï Aggiungi documento";
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
  if (ui.posAddToggleBtn) ui.posAddToggleBtn.textContent = "‚ûï Aggiungi documento";
  ui.posDocumentForm?.classList.add("hidden");
  if (ui.posFeedback) ui.posFeedback.textContent = "";
}

async function savePosDocument(event) {
  event.preventDefault();
  if (!canManageData()) {
    alert("Solo l'admin pu√≤ salvare documenti POS.");
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
    alert("Solo l'admin pu√≤ eliminare documenti POS.");
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
    const baseQuery = db.collection(getOreReportsCollectionName());
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
  const reportsQuery = db.collection(getOreReportsCollectionName())
    .where("date", ">=", fromDate)
    .where("date", "<=", toDate)
    .orderBy("date", "asc")
    .get();
  const approvalsQuery = includePendingApprovals
    ? db.collection(getOreApprovalRequestsCollectionName())
      .where("date", ">=", fromDate)
      .where("date", "<=", toDate)
      .orderBy("date", "asc")
      .get()
    : Promise.resolve(null);
  const [reportsSnapshot, approvalsSnapshot] = await Promise.all([reportsQuery, approvalsQuery]);
  const reports = reportsSnapshot.docs.map((doc) => ({
    id: doc.id,
    sourceCollection: getOreReportsCollectionName(),
    approvalStatus: "approved",
    ...doc.data()
  }));
  const pendingApprovals = approvalsSnapshot
    ? approvalsSnapshot.docs
      .map((doc) => ({
        id: doc.id,
        sourceCollection: getOreApprovalRequestsCollectionName(),
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
        const isPendingApproval = String(report.sourceCollection || getOreReportsCollectionName()) === getOreApprovalRequestsCollectionName();
        operatorsMap.get(operatore)[day - 1].push({ ore, isPendingApproval });
        const key = `${operatore}__${day}`;
        if (!hoursTableRowsMap.has(key)) hoursTableRowsMap.set(key, []);
        hoursTableRowsMap.get(key).push({
          recordId: report.id,
          reportId: report.id,
          sourceCollection: report.sourceCollection || getOreReportsCollectionName(),
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
      const pendingSources = sources.filter((source) => String(source.sourceCollection || getOreReportsCollectionName()) === getOreApprovalRequestsCollectionName());
      const hasPendingApproval = pendingSources.length > 0;
      const canManage = canManageData() && sources.length;
      const dayTotal = dayItems.reduce((sum, value) => sum + getDayItemHours(value), 0);
      const hasDuplicates = dayItems.length > 1;
      const hasDataError = sources.some((source) => !source.reportDate || !source.entryCommessaId || !source.operatore || Number(source.ore || 0) <= 0);
      let valueLabel = `‚úÖ ${formatHoursValue(dayTotal)}h ¬∑ ore inserite`;
      let statusClass = "hours-value-ok";
      if (hasDataError) {
        valueLabel = `‚ùå ${formatHoursValue(dayTotal)}h ¬∑ errore dati`;
        statusClass = "hours-value-error";
      } else if (hasDuplicates) {
        valueLabel = `‚ö†Ô∏è ${formatHoursValue(dayTotal)}h ¬∑ duplicato da controllare`;
        statusClass = "hours-value-warning";
      } else if (hasPendingApproval) {
        valueLabel = `‚ö†Ô∏è ${formatHoursValue(dayTotal)}h ¬∑ da confermare`;
        statusClass = "hours-value-warning hours-value-pending-approval";
      }
      const mergedDetails = hasDuplicates
        ? `Duplicato non valido: stesso operatore/commessa/giorno inserito pi√π volte. La pulizia automatica mantiene una sola registrazione valida.`
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
    `<tr><th scope="row" class="muted">‚Äî</th>${emptyCells}<td class="muted">0h</td></tr>`
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
      .some((report) => String(report.sourceCollection || getOreReportsCollectionName()) === getOreApprovalRequestsCollectionName());
    ui.hoursTableFeedback.textContent = hasPendingApprovals
      ? "Mostro anche vecchie richieste da confermare: sono evidenziate in giallo finch√© l'admin non le approva. Le nuove ore vengono salvate subito."
      : canManageData()
        ? "Clicca un valore per modificare o eliminare la registrazione ore."
        : "Ore salvate automaticamente: solo l'amministratore pu√≤ modificare o eliminare le ore.";
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
      if (String(source.sourceCollection || getOreReportsCollectionName()) !== getOreApprovalRequestsCollectionName()) return;
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
  if (!canManageData()) throw new Error("Solo admin pu√≤ confermare le ore.");
  const pendingSources = (Array.isArray(sources) ? sources : [])
    .filter((source) => String(source.sourceCollection || getOreReportsCollectionName()) === getOreApprovalRequestsCollectionName() && source.reportId);
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
      if (String(source.sourceCollection || getOreReportsCollectionName()) !== getOreApprovalRequestsCollectionName()) return;
      source.sourceCollection = getOreReportsCollectionName();
      source.approvalStatus = "approved";
    });
  });
  ui.hoursTableContainer.querySelectorAll(".hours-value-btn[data-hours-key]").forEach((btn) => {
    const key = String(btn.dataset.hoursKey || "");
    if (!keySet.has(key)) return;
    const valueText = String(btn.textContent || "").replace(/^‚ö†Ô∏è\s*/, "‚úÖ ").replace(" ¬∑ da confermare", " ¬∑ ore inserite");
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
    .filter((source) => String(source.sourceCollection || getOreReportsCollectionName()) === getOreApprovalRequestsCollectionName());
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
  const pendingSources = sources.filter((source) => String(source.sourceCollection || getOreReportsCollectionName()) === getOreApprovalRequestsCollectionName());
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
    const details = sources.map((source, idx) => `${idx + 1}) ${source.reportDate || "-"} ‚Ä¢ ${source.operatore || "-"} ‚Ä¢ ${Number(source.ore || 0)}h`).join("\n");
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
      const collectionName = String(source.sourceCollection || getOreReportsCollectionName()) === getOreApprovalRequestsCollectionName()
        ? getOreApprovalRequestsCollectionName()
        : getOreReportsCollectionName();
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
    await exportHoursGlobalMonthlyTable({
      onlyCommessaId: commessaId,
      emptyMessage: "Nessuna ora registrata per questa commessa nel mese selezionato.",
      fileNamePrefix: `ore_${String(commessaName || "commessa").replace(/[\\/:*?"<>|]/g, "_").replace(/\s+/g, "_")}`
    });
  } catch (error) {
    console.error("Errore export Excel ore:", error);
    if (ui.hoursTableFeedback) ui.hoursTableFeedback.textContent = "Errore export Excel ore. Controlla i dati o riprova.";
    alert("Errore export Excel ore. Controlla i dati o riprova.");
  }
}

async function exportHoursGlobalMonthlyTable(options = {}) {
  const onlyCommessaId = String(options.onlyCommessaId || "").trim();
  const emptyMessage = String(options.emptyMessage || "Nessuna ora registrata nel mese selezionato per l'export globale.");
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
    const globalOperatorDayMap = new Map();
    let totalValidGlobalRows = 0;
    reports.forEach((report) => {
      const day = Number(String(report.date || "").split("-")[2] || 0);
      if (!day || day < 1 || day > monthMeta.daysInMonth) return;
      const entries = Array.isArray(report.entries) ? report.entries : [];
      entries.forEach((entry) => {
        const entryCommessaInfo = resolveHoursEntryCommessa(entry);
        const commessaId = String(entryCommessaInfo.id || entryCommessaInfo.key || "").trim();
        if (!commessaId) return;
        if (onlyCommessaId && commessaId !== onlyCommessaId) return;
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
          if (!globalOperatorDayMap.has(operatorNorm)) {
            globalOperatorDayMap.set(operatorNorm, Array.from({ length: monthMeta.daysInMonth }, () => 0));
          }
          commessaBucket.operatorsMap.get(operatorNorm).days[day - 1] += ore;
          globalOperatorDayMap.get(operatorNorm)[day - 1] += ore;
        });
      });
    });

  logHoursDebug("dati usati per export", { mode: "global", monthValue, commesse: Array.from(commessaMap.values()).map((item) => ({
    commessaName: item.commessaName,
    commessaCode: item.commessaCode,
    operatori: Array.from(item.operatorsMap.values())
  })) });
  if (!commessaMap.size || totalValidGlobalRows <= 0) {
    alert(emptyMessage);
    if (ui.hoursTableFeedback) ui.hoursTableFeedback.textContent = emptyMessage;
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
  const ordinaryHoursColumn = monthMeta.daysInMonth + 3;
  const overtimeHoursColumn = monthMeta.daysInMonth + 4;
  const lastColumn = overtimeHoursColumn;
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
  const getEasterSunday = (year) => {
    const a = year % 19;
    const b = Math.floor(year / 100);
    const c = year % 100;
    const d = Math.floor(b / 4);
    const e = b % 4;
    const f = Math.floor((b + 8) / 25);
    const g = Math.floor((b - f + 1) / 3);
    const h = (19 * a + b - d - g + 15) % 30;
    const i = Math.floor(c / 4);
    const k = c % 4;
    const l = (32 + 2 * e + 2 * i - h - k) % 7;
    const m = Math.floor((a + 11 * h + 22 * l) / 451);
    const month = Math.floor((h + l - 7 * m + 114) / 31);
    const day = ((h + l - 7 * m + 114) % 31) + 1;
    return new Date(year, month - 1, day);
  };
  const easterMonday = getEasterSunday(monthMeta.year);
  easterMonday.setDate(easterMonday.getDate() + 1);
  const easterMondayKey = `${String(easterMonday.getMonth() + 1).padStart(2, "0")}-${String(easterMonday.getDate()).padStart(2, "0")}`;
  const italianHolidayKeys = new Set([
    "01-01",
    "01-06",
    easterMondayKey,
    "04-25",
    "05-01",
    "06-02",
    "08-15",
    "11-01",
    "12-08",
    "12-25",
    "12-26"
  ]);
  const isHolidayDay = (dayNumber) => {
    const holidayKey = `${String(monthMeta.month).padStart(2, "0")}-${String(dayNumber).padStart(2, "0")}`;
    return italianHolidayKeys.has(holidayKey);
  };
  const isWeekendDay = (dayNumber) => {
    const dayDate = new Date(monthMeta.year, monthMeta.month - 1, dayNumber);
    const weekday = dayDate.getDay();
    return weekday === 0 || weekday === 6;
  };
  const getOrdinaryHoursLimit = (dayNumber) => {
    if (isWeekendDay(dayNumber) || isHolidayDay(dayNumber)) return 0;
    const weekday = new Date(monthMeta.year, monthMeta.month - 1, dayNumber).getDay();
    if (weekday >= 1 && weekday <= 4) return 8;
    if (weekday === 5) return 7;
    return 0;
  };
  const splitOrdinaryAndOvertimeHours = (hours, dayNumber) => {
    const dailyHours = Number(hours || 0);
    if (!Number.isFinite(dailyHours) || dailyHours <= 0) return { ordinary: 0, overtime: 0 };
    const ordinaryLimit = getOrdinaryHoursLimit(dayNumber);
    const ordinary = Math.min(dailyHours, ordinaryLimit);
    return {
      ordinary,
      overtime: Math.max(dailyHours - ordinaryLimit, 0)
    };
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
  const totalOperatorsUnique = globalOperatorDayMap.size;
  const totalOperatorsActive = commesseSorted.reduce((acc, commessaBlock) => (
    acc + Array.from(commessaBlock.operatorsMap.values()).filter((operator) => (
      operator.days.some((value) => Number(value || 0) > 0)
    )).length
  ), 0);
  const monthlyHourTotals = Array.from(globalOperatorDayMap.values()).reduce((acc, days) => {
    days.forEach((value, dayIndex) => {
      const dailyHours = Number(value || 0);
      if (dailyHours <= 0) return;
      const dailyBreakdown = splitOrdinaryAndOvertimeHours(dailyHours, dayIndex + 1);
      acc.ordinary += dailyBreakdown.ordinary;
      acc.overtime += dailyBreakdown.overtime;
      acc.total += dailyHours;
    });
    return acc;
  }, { ordinary: 0, overtime: 0, total: 0 });

  const summaryStartRow = rowPointer;
  worksheet.mergeCells(summaryStartRow, 1, summaryStartRow, lastColumn);
  const summaryTitleCell = worksheet.getCell(summaryStartRow, 1);
  summaryTitleCell.value = " VARGA CANTIERI   RIEPILOGO GESTIONE ORE GLOBAL";
  summaryTitleCell.font = { bold: true, size: 14, color: { argb: "FF000000" } };
  summaryTitleCell.alignment = { horizontal: "center", vertical: "middle" };
  summaryTitleCell.fill = whiteFill;
  rowPointer += 1;

  const formatSummaryValue = (value) => {
    if (typeof value !== "number") return value;
    return Number.isInteger(value) ? value : Number(value.toFixed(2));
  };
  const summaryCardRows = [
    [
      ["MESE DI RIFERIMENTO", monthNameIt],
      ["ANNO", String(monthMeta.year)],
      ["DATA ESPORTAZIONE", new Date().toLocaleDateString("it-IT")]
    ],
    [
      ["TOTALE ORE", formatSummaryValue(monthlyHourTotals.total)],
      ["ORE ORDINARIE", formatSummaryValue(monthlyHourTotals.ordinary)],
      ["ORE STRAORDINARIE", formatSummaryValue(monthlyHourTotals.overtime)]
    ],
    [
      ["TOTALE OPERATORI", totalOperatorsActive],
      ["TOTALE OPERATORI UNICI", totalOperatorsUnique],
      ["TOTALE COMMESSE", totalCommesse]
    ]
  ];
  const summaryColumnGroups = [
    [1, Math.floor(lastColumn / 3)],
    [Math.floor(lastColumn / 3) + 1, Math.floor((lastColumn * 2) / 3)],
    [Math.floor((lastColumn * 2) / 3) + 1, lastColumn]
  ];
  summaryCardRows.forEach((cards) => {
    const labelRowIndex = rowPointer;
    const valueRowIndex = rowPointer + 1;
    cards.forEach(([label, value], cardIndex) => {
      const [startCol, endCol] = summaryColumnGroups[cardIndex];
      worksheet.mergeCells(labelRowIndex, startCol, labelRowIndex, endCol);
      worksheet.mergeCells(valueRowIndex, startCol, valueRowIndex, endCol);
      const labelCell = worksheet.getCell(labelRowIndex, startCol);
      const valueCell = worksheet.getCell(valueRowIndex, startCol);
      labelCell.value = label;
      valueCell.value = value;
      labelCell.font = { bold: true, size: 8, color: { argb: "FF4B5563" } };
      valueCell.font = { bold: true, size: 13, color: { argb: "FF000000" } };
      labelCell.alignment = { horizontal: "center", vertical: "bottom" };
      valueCell.alignment = { horizontal: "center", vertical: "top" };
      if (typeof value === "number") valueCell.numFmt = getExcelNumberFormat(value) || "0";
      for (let row = labelRowIndex; row <= valueRowIndex; row += 1) {
        for (let col = startCol; col <= endCol; col += 1) {
          const cell = worksheet.getCell(row, col);
          cell.fill = whiteFill;
          setThinBorder(cell);
        }
      }
    });
    worksheet.getRow(labelRowIndex).height = 13;
    worksheet.getRow(valueRowIndex).height = 19;
    rowPointer += 2;
  });
  const summaryEndRow = rowPointer - 1;
  rowPointer += 1;
  const firstCommessaStartRow = rowPointer;

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
    headerRow.getCell(ordinaryHoursColumn).value = "ORE ORDINARIE";
    headerRow.getCell(ordinaryHoursColumn).fill = dayHeaderFill;
    headerRow.getCell(ordinaryHoursColumn).font = { bold: true, color: { argb: "FF000000" } };
    headerRow.getCell(overtimeHoursColumn).value = "ORE STRAORDINARIE";
    headerRow.getCell(overtimeHoursColumn).fill = dayHeaderFill;
    headerRow.getCell(overtimeHoursColumn).font = { bold: true, color: { argb: "FF000000" } };
    headerRow.height = 24;

    rowPointer += 1;
    let commessaTotal = 0;

    operators.forEach((operatorData, operatorIdx) => {
      const row = worksheet.getRow(rowPointer + operatorIdx);
      row.getCell(1).value = operatorData.displayName || "";
      let total = 0;
      let ordinaryHours = 0;
      let overtimeHours = 0;
      for (let dayIdx = 0; dayIdx < monthMeta.daysInMonth; dayIdx += 1) {
        const value = Number(operatorData.days[dayIdx] || 0);
        const cell = row.getCell(dayIdx + 2);
        if (isWeekendDay(dayIdx + 1) || isHolidayDay(dayIdx + 1)) {
          cell.fill = weekendFill;
        } else {
          cell.fill = whiteFill;
        }
        if (value > 0) {
          cell.value = value;
          cell.fill = value > 12 ? errorFill : hoursFilledCell;
          const numFmt = getExcelNumberFormat(value);
          if (numFmt) cell.numFmt = numFmt;
          const dailyBreakdown = splitOrdinaryAndOvertimeHours(value, dayIdx + 1);
          ordinaryHours += dailyBreakdown.ordinary;
          overtimeHours += dailyBreakdown.overtime;
          total += value;
        } else {
          cell.value = null;
        }
      }
      row.getCell(totalColumn).value = total > 0 ? total : null;
      row.getCell(totalColumn).fill = totalColumnFill;
      if (total > 0) row.getCell(totalColumn).numFmt = getExcelNumberFormat(total);
      row.getCell(ordinaryHoursColumn).value = ordinaryHours > 0 ? ordinaryHours : null;
      if (ordinaryHours > 0) row.getCell(ordinaryHoursColumn).numFmt = getExcelNumberFormat(ordinaryHours);
      row.getCell(overtimeHoursColumn).value = overtimeHours > 0 ? overtimeHours : null;
      if (overtimeHours > 0) row.getCell(overtimeHoursColumn).numFmt = getExcelNumberFormat(overtimeHours);
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

  for (let col = 1; col <= lastColumn; col += 1) {
    const cell = worksheet.getCell(summaryStartRow, col);
    cell.fill = whiteFill;
    setThinBorder(cell);
  }
  worksheet.getRow(summaryStartRow).height = 24;
  setOuterBlockBorder(summaryStartRow, summaryEndRow);

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
    if (col === ordinaryHoursColumn) {
      worksheet.getColumn(col).width = 16;
      continue;
    }
    worksheet.getColumn(col).width = 18;
  }

  worksheet.autoFilter = {
    from: { row: firstCommessaStartRow + 3, column: 1 },
    to: { row: firstCommessaStartRow + 3, column: 1 }
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
  const safePrefix = String(options.fileNamePrefix || "ore_global").replace(/[\\/:*?"<>|]/g, "_").replace(/\s+/g, "_");
  const fileName = `${safePrefix}_${safeMonth}.xlsx`;
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
      <span class="hours-commessa-picker-arrow" aria-hidden="true">‚ñº</span>
    </button>
    <div class="hours-commessa-picker-menu ${isOpen ? "" : "hidden"}">
      <input class="hours-commessa-picker-search" type="search" placeholder="Cerca commessa‚Ä¶" value="${escapeHTML(query)}" aria-label="Cerca commessa">
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
  const currentProfile = platformUsers.find((user) => String(user.id || user.uid || "") === String(currentUser.uid || ""));
  return getPlatformUserIdentityParts(currentProfile || currentUser);
}

function getSquadraNameVariants(value) {
  const normalized = normalizeSquadraMemberIdentity(value);
  if (!normalized) return [];
  const words = normalized.split(" ").filter(Boolean);
  const reversed = words.length > 1 ? [...words].reverse().join(" ") : normalized;
  return [...new Set([normalized, reversed])];
}

function getCurrentUserSquadraIdentity() {
  if (!currentUser) return { uids: new Set(), personaleIds: new Set(), emails: new Set(), names: new Set() };
  const currentProfile = platformUsers.find((user) => String(user.id || user.uid || "") === String(currentUser.uid || "")) || {};
  const currentEmail = normalizeEmail(currentUser.email || currentProfile.email || "");
  const linkedPerson = personaleRecords.find((person) =>
    [currentUser.personaleId, currentProfile.personaleId, currentProfile.personId].filter(Boolean).some((id) => String(person.id) === String(id))
    || (currentEmail && normalizeEmail(person.email) === currentEmail)
  );
  const values = (...items) => new Set(items.map((item) => String(item || "").trim()).filter(Boolean));
  const names = new Set();
  [
    linkedPerson && getPersonaleDisplayName(linkedPerson),
    linkedPerson?.fullName,
    currentUser.displayName,
    currentProfile.displayName,
    currentProfile.fullName,
    currentProfile.nome,
    currentProfile.nome && currentProfile.cognome ? `${currentProfile.nome} ${currentProfile.cognome}` : ""
  ].forEach((name) => getSquadraNameVariants(name).forEach((variant) => names.add(variant)));
  return {
    uids: values(currentUser.uid, currentProfile.uid),
    personaleIds: values(currentUser.personaleId, currentProfile.personaleId, currentProfile.personId, linkedPerson?.id),
    emails: new Set([currentEmail, normalizeEmail(linkedPerson?.email)].filter(Boolean)),
    names
  };
}

function getSquadraMemberIdentifiers(member) {
  if (Array.isArray(member)) {
    return member.reduce((all, item) => {
      const next = getSquadraMemberIdentifiers(item);
      Object.keys(all).forEach((key) => all[key].push(...next[key]));
      return all;
    }, { uids: [], personaleIds: [], emails: [], names: [] });
  }
  if (member && typeof member === "object") {
    return {
      uids: [member.uid, member.firebaseUid, member.userId, member.utenteId].filter(Boolean).map(String),
      personaleIds: [member.personaleId, member.personId, member.id].filter(Boolean).map(String),
      emails: [member.email, member.mail].filter(Boolean).map(normalizeEmail),
      names: [member.displayName, member.nomeCompleto, member.fullName, member.name,
        member.nome && member.cognome ? `${member.nome} ${member.cognome}` : ""].filter(Boolean).flatMap(getSquadraNameVariants)
    };
  }
  const parts = parseMultiEntryValue(member || "");
  return {
    // A legacy scalar can represent any supported identifier. Exact matching keeps
    // UID/personnel-ID comparisons safe while names are normalized separately.
    uids: parts,
    personaleIds: parts,
    emails: parts.map(normalizeEmail),
    names: parts.flatMap(getSquadraNameVariants)
  };
}

function doesSquadraMemberMatchCurrentUser(member, identity = getCurrentUserSquadraIdentity()) {
  const candidate = getSquadraMemberIdentifiers(member);
  // Keep this order intentional: stable identifiers win over mutable labels.
  if (candidate.uids.some((value) => identity.uids.has(String(value).trim()))) return true;
  if (candidate.personaleIds.some((value) => identity.personaleIds.has(String(value).trim()))) return true;
  if (candidate.emails.some((value) => identity.emails.has(normalizeEmail(value)))) return true;
  return candidate.names.some((value) => identity.names.has(normalizeSquadraMemberIdentity(value)));
}

function getSquadraMemberIdentityValues(member) {
  if (Array.isArray(member)) return member.flatMap(getSquadraMemberIdentityValues);
  if (member && typeof member === "object") {
    return [
      member.uid, member.userId, member.utenteId, member.id, member.email,
      member.displayName, member.nomeCompleto, member.name,
      member.nome && member.cognome ? `${member.nome} ${member.cognome}` : ""
    ].map((value) => String(value || "").trim()).filter(Boolean);
  }
  return parseMultiEntryValue(member || "");
}

function getSquadraRowMembers(row = {}) {
  return [row.personale, row.operatori, row.caposquadra].flatMap(getSquadraMemberIdentityValues);
}

function getPlatformUserIdentityParts(user) {
  if (!user) return [];
  const userEmail = String(user.email || "").trim();
  const linkedPerson = userEmail
    ? personaleRecords.find((person) => normalizeEmail(person?.email) === normalizeEmail(userEmail))
    : null;
  const parts = [
    user.id,
    user.uid,
    userEmail,
    linkedPerson?.email,
    linkedPerson ? getPersonaleDisplayName(linkedPerson) : "",
    user.displayName,
    user.firstName && user.lastName ? `${user.firstName} ${user.lastName}` : "",
    user.userName,
    user.nome
  ];
  const emailLocal = userEmail.split("@")[0] || "";
  if (emailLocal) parts.push(emailLocal, emailLocal.replace(/[._-]+/g, " "));
  return [...new Set(parts.map(normalizeSquadraMemberIdentity).filter(Boolean))];
}

function doSquadraMemberAndUserMatch(memberName, identityParts = null) {
  if (!identityParts) return doesSquadraMemberMatchCurrentUser(memberName);
  const members = getSquadraMemberIdentityValues(memberName).map(normalizeSquadraMemberIdentity).filter(Boolean);
  if (!members.length || !identityParts.length) return false;
  return members.some((member) => identityParts.some((part) => {
    if (!part) return false;
    if (member === part) return true;
    const memberTokens = member.split(" ").filter((token) => token.length > 1);
    const partTokens = part.split(" ").filter((token) => token.length > 1);
    return partTokens.length >= 2 && partTokens.every((token) => memberTokens.includes(token));
  }));
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
  const identity = getCurrentUserSquadraIdentity();
  for (let index = 0; index < squadRows.length; index += 1) {
    const personale = getSquadraRowMembers(squadRows[index]);
    const matchedName = personale.find((member) => doesSquadraMemberMatchCurrentUser(member, identity));
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

function getSquadrePerCommessaForDate(dateKey = getActiveSquadreDateKey()) {
  if (!dateKey) return [];
  const storicoDelGiorno = squadreHistoryByDate.get(dateKey) || new Map();
  return Array.from(commesseById.values()).flatMap((commessa) => {
    const squadData = storicoDelGiorno.get(commessa.id) || null;
    if (!squadData) return [];
    const squadRows = Array.isArray(squadData.squadre) ? squadData.squadre : getLegacySquadreRows(squadData || {});
    if (!squadRows.some(isSquadraRowFilled)) return [];
    return [{ commessa, squadData, squadRows }];
  });
}

function findCurrentUserSquadreForDate(dateKey = getActiveSquadreDateKey()) {
  if (!currentUser || !dateKey) return [];
  const identity = getCurrentUserSquadraIdentity();
  return getSquadrePerCommessaForDate(dateKey).flatMap(({ commessa, squadData, squadRows }) => {
    const matchedRows = [];

    squadRows.forEach((row, index) => {
      const rowMembers = getSquadraRowMembers(row);
      const matchedName = rowMembers.find((member) => doesSquadraMemberMatchCurrentUser(member, identity));
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
      return [{
        commessa,
        commessaId: commessa.id,
        commessaName: commessa.nome || "Commessa",
        squadData,
        matchedRows
      }];
    }
    return [];
  });
}

function getCurrentUserAssignedCommesseForDate(dateKey = getActiveSquadreDateKey()) {
  return findCurrentUserSquadreForDate(dateKey);
}

function tryAutoOpenAssignedCommessaAtStartup() {
  // All'avvio l'app deve restare sempre sulla home principale: non ripristina
  // commesse salvate e non apre automaticamente le commesse assegnate.
  startupAssignedCommessaAutoOpenDone = true;
}

function canCurrentUserInsertHoursForCommessa(commessaId, dateValue = "") {
  return Boolean(currentUser && String(commessaId || "").trim());
}

function getHoursOperatorForCurrentUser(commessaId, dateValue = "") {
  const assignment = getCurrentUserSquadraAssignment(commessaId, dateValue);
  return assignment?.matchedName || getCurrentUserResolvedName();
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
        const operatoreId = resolveHoursOperatorId(name);
        const participantId = getHoursParticipantId(
          {
            operatore: name,
            operatoreId,
            squadraIndex,
            squadraLabel
          },
          { squadraIndex, squadraLabel },
          { allowSquadraFallback: false }
        );
        const key = buildHoursFullKey(commessaId, dateValue, participantId);
        if (!key) return;
        participants.set(key, {
          key,
          participantId,
          operatoreId,
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
  if (areAllHoursParticipantsCompleteForCommessaDate(commessaId, dateKey)) return null;
  const assignment = getCurrentUserSquadraAssignment(commessaId, dateKey);
  const squadraIndex = assignment?.squadraIndex || "";
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
    handleQuickAddHoursClick(commessa, dateValue);
  });
  return button;
}

function getTodayDateKey() {
  const parts = new Intl.DateTimeFormat("it-IT", {
    timeZone: "Europe/Rome", year: "numeric", month: "2-digit", day: "2-digit"
  }).formatToParts(new Date());
  const value = Object.fromEntries(parts.map(({ type, value: partValue }) => [type, partValue]));
  return `${value.year}-${value.month}-${value.day}`;
}

function buildSquadraDateCorrectionWhatsappUrl({ commessaName, squadraDate, todayDate, operatorName }) {
  const message = [
    "‚ö†Ô∏è RICHIESTA CORREZIONE GIORNO SQUADRA",
    "",
    "Ciao, sto provando a inserire le ore, ma la squadra risulta programmata per un giorno diverso da oggi.",
    "",
    `Commessa: ${commessaName || "Commessa"}`,
    `Data squadra: ${formatDateKeyForDisplay(squadraDate)}`,
    `Data attuale: ${formatDateKeyForDisplay(todayDate)}`,
    `Operatore: ${operatorName || "Operatore"}`,
    "",
    "Puoi verificare e correggere il giorno della squadra?",
    "",
    "Grazie."
  ].join("\n");
  return `https://wa.me/393892352575?text=${encodeURIComponent(message)}`;
}

function openSquadraDifferentDayInfoPopup({ commessa, squadraDate, todayDate, onContinue }) {
  const commessaName = commessa?.nome || "Commessa";
  const operatorName = currentUser?.displayName || currentUser?.email || "Operatore";
  const overlay = document.createElement("div");
  overlay.className = "confirm-modal";
  overlay.innerHTML = `
    <div class="confirm-modal-card" role="dialog" aria-modal="true" aria-labelledby="squadra-day-warning-title">
      <h2 id="squadra-day-warning-title">‚ö†Ô∏è Giorno diverso da oggi</h2>
      <p>Stai inserendo le ore per una squadra programmata in un giorno diverso da oggi.</p>
      <p>Se il giorno √® sbagliato, avvisa l‚Äôamministratore per correggere la data della squadra.<br>Se invece vuoi continuare comunque, puoi inserire le ore per il giorno selezionato.</p>
      <div class="confirm-modal-actions">
        <button type="button" class="btn btn-primary" data-hours-day-continue>Inserisci lo stesso</button>
        <button type="button" class="btn" data-hours-day-correction>Richiedi correzione</button>
      </div>
    </div>`;
  const close = () => overlay.remove();
  overlay.querySelector("[data-hours-day-continue]")?.addEventListener("click", () => {
    close();
    onContinue?.();
  });
  overlay.querySelector("[data-hours-day-correction]")?.addEventListener("click", () => {
    const url = buildSquadraDateCorrectionWhatsappUrl({ commessaName, squadraDate, todayDate, operatorName });
    window.open(url, "_blank", "noopener,noreferrer");
    close();
  });
  overlay.addEventListener("click", (event) => {
    if (event.target === overlay) close();
  });
  document.body.appendChild(overlay);
  overlay.querySelector("[data-hours-day-continue]")?.focus();
}

function handleQuickAddHoursClick(commessa, dateValue = "") {
  const squadra˜~˚ﬂ¶ÚµÎ(ö+my÷‚ì∞¢ñbÇ6ˆ÷÷W76ñB«¬ñ◊ñÁFÙñG2Ê∆VÊwFÇí&WGW&‚f«6S∞¢G'í∞¢6ˆÁ7B&Vb“F"Ê6ˆ∆∆V7Fñˆ‚ÜvWD6ˆ÷÷W76T6ˆ∆∆V7Fñˆ‰Ê÷RÇííÊFˆ2Ü6ˆ÷÷W76ñBíÊ6ˆ∆∆V7Fñˆ‚Ç&ñ◊ñÁFí"ì∞¢6ˆÁ7B6Ê6Ü˜G2“vóB&ˆ÷ó6RÊ∆¬Üñ◊ñÁFÙñG2Ê÷ÇÜñ◊ñÁFÙñBí”‚&VbÊFˆ2Üñ◊ñÁFÙñBíÊvWBÇííì∞¢&WGW&‚6Ê6Ü˜G2ÊWfW'íÇá6Êí”‚6ÊÊWÜó7G2bbó4ñ◊ñÁFÙFˆÊU7FFRá6ÊÊFFÇí«¬∑“íì∞¢“6F6ÇÜW'&˜"í∞¢6ˆÁ6ˆ∆RÊW'&˜"Ç$W'&˜&RfW&ñfñ6W'6ó7FVÁ¶dEDÛ¢"¬W'&˜"ì∞¢&WGW&‚f«6S∞¢–ß–†¶7ñÊ2gVÊ7Fñˆ‚ÜÊF∆Tñ◊ñÁFÙFˆÊU6fTfñ«W&RÜñ◊ñÁFÚ¬&V6ˆ‚“""í∞¢∆W'BÇ$W'&˜&R6«fFvvñÚ‚&ó&˜f‚"ì∞¢G'í∞¢vóBÊ˜FñgîF÷ñÁ4f˜$ñ◊ñÁFÙFˆÊU6fTW'&˜"Üñ◊ñÁFÚ¬&V6ˆ‚ì∞¢“6F6ÇÜW'&˜"í∞¢6ˆÁ6ˆ∆RÊW'&˜"Ç$W'&˜&RÊ˜Fñfñ6F÷ñ‚6«fFvvñÚdEDÛ¢"¬W'&˜"ì∞¢–ß–†¶7ñÊ2gVÊ7Fñˆ‚Ê˜FñgîF÷ñÁ4f˜$ñ◊ñÁFÙFˆÊU6fTW'&˜"Üñ◊ñÁFÚ¬&V6ˆ‚“""í∞¢6ˆÁ7BF÷ñÂW6W'2“∆Ff˜&’W6W'2Êfñ«FW"ÇáW6W"í”‚F÷ñ‰V÷ñ«2ÊÜ2ÜÊ˜&÷∆ó¶TV÷ñ¬áW6W"ÊV÷ñ¬ííì∞¢ñbÇF÷ñÂW6W'2Ê∆VÊwFÇí&WGW&„∞¢6ˆÁ7B˜W&F˜$Ê÷R“7W'&VÁEW6W#ÚÊFó7∆îÊ÷R«¬7W'&VÁEW6W#ÚÊV÷ñ¬«¬$˜W&F˜&R#∞¢6ˆÁ7BÊ˜r“ÊWrFFRÇì∞¢6ˆÁ7BFWáB“∞¢.)™˚àÚU%$ı$R4≈dDttîÚdEDÚ"¿¢$Œ(	ññ◊ñÁFÚ:Ç7FFÚñÁfñFÚ7RvÜßßW¬÷˜G&V&&RÊˆ‚W76W&R76FÚÊV∆∆∆ó7FdEDí‚"¿¢%fW&ñfñ6&R÷ÁV∆÷VÁFR‚"¿¢""¿¢ñ◊ñÁFÛ¢G∂ñ◊ñÁFÛÚÊFVÊˆ÷ñÊ¶ñˆÊR«¬$ñ◊ñÁFÚ'÷¿¢îB4¢G∂ñ◊ñÁFÛÚÊñE6«¬"“'÷¿¢6ˆ◊VÊS¢G∂ñ◊ñÁFÛÚÊ6ˆ◊VÊR«¬"“'÷¿¢˜W&F˜&S¢G∂˜W&F˜$Ê÷W÷¿¢FFR˜&¢G∂Ê˜rÁFÙ∆ˆ6∆U7G&ñÊrÇ&óB‘ïB"ó÷¿¢&V6ˆ‚ÚW'&˜&R&ñ∆WfFÛ¢G∑&V6ˆÁ÷¢$W'&˜&R&ñ∆WfFÛ¢fW&ñfñ6dEDíÊˆ‚6ˆÊfW&÷F‚ ¢“Ê¶ˆñ‚Ç%∆‚"ì∞†¢vóB&ˆ÷ó6RÊ∆¬ÜF÷ñÂW6W'2Ê÷ÇÜF÷ñÂW6W"í”‚F"Ê6ˆ∆∆V7Fñˆ‚Ç&6ÜD÷W76vW2"íÊFBá∞¢GóS¢'FWáB"¿¢FWáB¿¢&V6óñVÁDñC¢F÷ñÂW6W"ÊñB¿¢6VÊFW$ñC¢7W'&VÁEW6W#ÚÁVñB«¬""¿¢6VÊFW$Ê÷S¢˜W&F˜$Ê÷R¿¢6VÊFW$V÷ñ√¢7W'&VÁEW6W#ÚÊV÷ñ¬«¬""¿¢∂ñÊC¢'7ó7FV“"¿¢÷WFFF¢∞¢GóS¢&ñ◊ñÁFıˆFˆÊU˜6fUˆW'&˜""¿¢6ˆ÷÷W76ñC¢6V∆V7FVD6ˆ÷÷W76ñB«¬""¿¢6ˆ÷÷W76Ê÷S¢6V∆V7FVD6ˆ÷÷W76Ê÷R«¬$6ˆ÷÷W76"¿¢ñ◊ñÁFÙÊ÷S¢ñ◊ñÁFÛÚÊFVÊˆ÷ñÊ¶ñˆÊR«¬$ñ◊ñÁFÚ"¿¢ñ◊ñÁFÙ∂Wì¢'Vñ∆Dñ◊ñÁFÙ∂WíÜñ◊ñÁFÚí¿¢ñ◊ñÁFÙñE6¢ñ◊ñÁFÛÚÊñE6«¬"“"¿¢ñ◊ñÁFÙ6ˆ◊VÊS¢ñ◊ñÁFÛÚÊ6ˆ◊VÊR«¬"“"¿¢˜W&F˜$Ê÷R¿¢FWFV7FVDC¢fó&V&6RÊfó&W7F˜&R‰fñV∆Ef«VRÁ6W'fW%Fñ÷W7F◊Çí¿¢&V6ˆ„¢&V6ˆ‚«¬&Ê˜Eˆ6ˆÊfó&÷VEˆñÂˆFˆÊUˆ∆ó7B ¢“¿¢7&VFVDC¢fó&V&6RÊfó&W7F˜&R‰fñV∆Ef«VRÁ6W'fW%Fñ÷W7F◊Çê¢“ííì∞ß–†¶gVÊ7Fñˆ‚vWDñ◊ñÁFıvÜG4FV◊∆FT66ÜT∂WíÜñ◊ñÁFÚ¬6ˆ÷÷W76ñB“6V∆V7FVD6ˆ÷÷W76ñBí∞¢6ˆÁ7Bñ◊ñÁFÙ∂Wí“'Vñ∆Dñ◊ñÁFÙ∂WíÜñ◊ñÁFÚì∞¢&WGW&‚6ˆ÷÷W76ñBbbñ◊ñÁFÙ∂WíÚG∂6ˆ÷÷W76ñG”¢G∂ñ◊ñÁFÙ∂Wó÷¢"#∞ß–†¶gVÊ7Fñˆ‚vWD7W'&VÁEvÜG4˜W&F˜$Ê÷RÇí∞¢6ˆÁ7BW6W"“WFÇÊ7W'&VÁEW6W"«¬7W'&VÁEW6W"«¬ÁV∆√∞¢&WGW&‚7G&ñÊráW6W#ÚÊFó7∆îÊ÷R«¬W6W#ÚÊV÷ñ¬«¬$˜W&F˜&R"íÁG&ñ“Çí«¬$˜W&F˜&R#∞ß–†¶gVÊ7Fñˆ‚vWDFWfñ6UvÜG4FFT∆&V¬ÜFFR“ÊWrFFRÇíí∞¢6ˆÁ7Bf«VR“FFRñÁ7FÊ6VˆbFFRÚFFR¢ÊWrFFRÜFFRì∞¢6ˆÁ7Bf∆ñDFFR“ÁV÷&W"Êó4Ê‚áf«VRÊvWEFñ÷RÇííÚÊWrFFRÇí¢f«VS∞¢&WGW&‚f∆ñDFFRÁFÙ∆ˆ6∆TFFU7G&ñÊrÇ&óB‘ïB"ì∞ß–†¶gVÊ7Fñˆ‚vWDñ◊ñÁFıvÜG4FV◊∆FU6ñvÊGW&RÜñ◊ñÁFÚí∞¢6ˆÁ7B∆ñÊ∂VDÊ˜FW2“vWD6ˆ÷÷W76Ê˜FT∆ñÊ∂VDÊ˜FW2Üñ◊ñÁFÚì∞¢&WGW&‚•4Ù‚Á7G&ñÊvñgíá∞¢6ˆ÷÷W76ñC¢6V∆V7FVD6ˆ÷÷W76ñB«¬""¿¢˜W&F˜$Ê÷S¢vWD7W'&VÁEvÜG4˜W&F˜$Ê÷RÇí¿¢FWfñ6TFFS¢vWDFWfñ6UvÜG4FFT∆&V¬Çí¿¢ñE6¢ñ◊ñÁFÛÚÊñE6«¬""¿¢FVÊˆ÷ñÊ¶ñˆÊS¢ñ◊ñÁFÛÚÊFVÊˆ÷ñÊ¶ñˆÊR«¬""¿¢6ˆ◊VÊS¢ñ◊ñÁFÛÚÊ6ˆ◊VÊR«¬""¿¢ñÊFó&óß¶Û¢ñ◊ñÁFÛÚÊñÊFó&óß¶Ú«¬""¿¢6ˆFñ6U&Wß¶Û¢ñ◊ñÁFÛÚÊ6ˆFñ6U&Wß¶Ú«¬""¿¢Fóˆ∆ˆvñ¢ñ◊ñÁFÛÚÁFóˆ∆ˆvññ◊ñÁFÚ«¬ñ◊ñÁFÛÚÁFóÙñ◊ñÁFÚ«¬ñ◊ñÁFÛÚÁFóˆ∆ˆvññÁFW'fVÁFÚ«¬""¿¢6ˆ÷÷W76¢6V∆V7FVD6ˆ÷÷W76Ê÷R«¬""¿¢∆f˜&¶ñˆÊï&ñ6ÜñW7FS¢ñ◊ñÁFÛÚÊ∆f˜&¶ñˆÊï&ñ6ÜñW7FR«¬""¿¢Fóˆ∆ˆvññÁFW'fVÁFÛ¢ñ◊ñÁFÛÚÁFóˆ∆ˆvññÁFW'fVÁFÚ«¬""¿¢Ê˜FTñ◊ñÁFÛ¢ñ◊ñÁFÛÚÊÊ˜FTñ◊ñÁFÚ«¬""¿¢∆ñÊ∂VDÊ˜FW3¢∆ñÊ∂VDÊ˜FW2Ê÷ÇÜÊ˜FRí”‚á≤ñC¢Ê˜FRÊñB«¬""¬FóF∆S¢vWD6ˆ÷÷W76Ê˜FUFóF∆RÜÊ˜FRí¬FWáC¢Ê˜FRÁFWáB«¬""“íê¢“ì∞ß–†¶gVÊ7Fñˆ‚'Vñ∆Dñ◊ñÁFıvÜG4FV◊∆FRÜñ◊ñÁFÚí∞¢6ˆÁ7Bó4ˆÊ«î˜&FñÊ&ñ“Ü4˜&FñÊ&ñÚÜñ◊ñÁFÚÊ6ˆFñ6U&Wß¶ÚíbbÜ57G&˜&FñÊ&ñÚÜñ◊ñÁFÚÊ6ˆFñ6U&Wß¶Úì∞¢6ˆÁ7B˜W&F˜$Ê÷R“vWD7W'&VÁEvÜG4˜W&F˜$Ê÷RÇì∞¢6ˆÁ7BFFR“vWDFWfñ6UvÜG4FFT∆&V¬Çì∞¢6ˆÁ7BFóF∆R“/	˘˙"î’îÂDÚdEDÚ#∞¢6ˆÁ7Bv˜&¥∆&V¬“ó4ˆÊ«î˜&FñÊ&ñ¢Ú$÷ÁWFVÁ¶ñˆÊR˜&FñÊ&ñW6VwVóF ¢¢$÷ÁWFVÁ¶ñˆÊR˜&FñÊ&ñ≤7G&˜&FñÊ&ñW6VwVóF#∞¢6ˆÁ7BFóˆ∆ˆvñ“ñ◊ñÁFÚÁFóˆ∆ˆvññ◊ñÁFÚ«¬ñ◊ñÁFÚÁFóÙñ◊ñÁFÚ«¬ñ◊ñÁFÚÁFóˆ∆ˆvññÁFW'fVÁFÚ«¬"“#∞¢6ˆÁ7B∆ñÊ∂VDÊ˜FW2“vWD6ˆ÷÷W76Ê˜FT∆ñÊ∂VDÊ˜FW2Üñ◊ñÁFÚì∞¢6ˆÁ7BÊ˜FT∆ñÊW2“∞¢ñ◊ñÁFÚÊÊ˜FTñ◊ñÁFÚÚ	˘9“Ê˜FRñ◊ñÁFÛ¢G∂ñ◊ñÁFÚÊÊ˜FTñ◊ñÁF˜÷¢""¿¢‚‚‚Ü∆ñÊ∂VDÊ˜FW2Ê∆VÊwFÇÚ∞¢.)™˚àÚVW7FÚñ◊ñÁFÚ:Ç7FF6VvÊ∆FVÊ7&óFñ6óL:¢"¿¢‚‚Ê∆ñÊ∂VDÊ˜FW2Ê÷ÇÜÊ˜FRí”‚G∂vWD6ˆ÷÷W76Ê˜FUFóF∆RÜÊ˜FRó’∆‚G∂Ê˜FRÁFWáB«¬"“'÷ê¢“¢µ“ê¢“Êfñ«FW"Ñ&ˆˆ∆V‚ì∞¢&WGW&‚∞¢FóF∆R¿¢)»RGFófóL:¢G∑v˜&¥∆&V«÷¿¢	¯iBîB4¢G∂ñ◊ñÁFÚÊñE6«¬"“'÷¿¢	¯˘~˚àÚñ◊ñÁFÛ¢G∂ñ◊ñÁFÚÊFVÊˆ÷ñÊ¶ñˆÊR«¬"“'÷¿¢	˘8“6ˆ◊VÊS¢G∂ñ◊ñÁFÚÊ6ˆ◊VÊR«¬"“'÷¿¢	˘∫>˚àÚfñ¢G∂ñ◊ñÁFÚÊñÊFó&óß¶Ú«¬"“'÷¿¢	˘9“Ê˜FS¢G∑Fóˆ∆ˆvñ÷¿¢‚‚‚Üó4ˆÊ«î˜&FñÊ&ñÚµ“¢∂	˘∫˚àÚ∆f˜&¶ñˆÊR7G&˜&FñÊ&ñ¢G∂ñ◊ñÁFÚÊ∆f˜&¶ñˆÊï&ñ6ÜñW7FR«¬ñ◊ñÁFÚÁFóˆ∆ˆvññÁFW'fVÁFÚ«¬"“'÷“í¿¢‚‚ÊÊ˜FT∆ñÊW2¿¢	˘r˜W&F˜&S¢G∂˜W&F˜$Ê÷W÷¿¢	˘8RFF¢G∂FFW÷ ¢“Ê¶ˆñ‚Ç%∆‚"ì∞ß–†¶gVÊ7Fñˆ‚&W&Tñ◊ñÁFıvÜG4FV◊∆FRÜñ◊ñÁFÚí∞¢6ˆÁ7B66ÜT∂Wí“vWDñ◊ñÁFıvÜG4FV◊∆FT66ÜT∂WíÜñ◊ñÁFÚì∞¢ñbÇ66ÜT∂Wíí&WGW&‚ÁV∆√∞¢6ˆÁ7B6ñvÊGW&R“vWDñ◊ñÁFıvÜG4FV◊∆FU6ñvÊGW&RÜñ◊ñÁFÚì∞¢6ˆÁ7B66ÜVB“ñ◊ñÁFıvÜG4FV◊∆FT66ÜRÊvWBÜ66ÜT∂Wíì∞¢ñbÜ66ÜVCÚÁ6ñvÊGW&R””“6ñvÊGW&Rí&WGW&‚66ÜVC∞¢6ˆÁ7B&W&VB“≤6ñvÊGW&R¬FV◊∆FS¢'Vñ∆Dñ◊ñÁFıvÜG4FV◊∆FRÜñ◊ñÁFÚí¬WFFVDC¢FFRÊÊ˜rÇí”∞¢ñ◊ñÁFıvÜG4FV◊∆FT66ÜRÁ6WBÜ66ÜT∂Wí¬&W&VBì∞¢&WGW&‚&W&VC∞ß–†¶gVÊ7Fñˆ‚&Vg&W6Ññ◊ñÁFıvÜG4FV◊∆FT66ÜRÜñ◊ñÁFí“µ“í∞¢6ˆÁ7B7FófT∂Wó2“ÊWr6WBÇì∞¢ñ◊ñÁFíÊf˜$V6ÇÇÜñ◊ñÁFÚí”‚∞¢6ˆÁ7B&W&VB“&W&Tñ◊ñÁFıvÜG4FV◊∆FRÜñ◊ñÁFÚì∞¢6ˆÁ7B66ÜT∂Wí“vWDñ◊ñÁFıvÜG4FV◊∆FT66ÜT∂WíÜñ◊ñÁFÚì∞¢ñbá&W&VBbb66ÜT∂Wíí7FófT∂Wó2ÊFBÜ66ÜT∂Wíì∞¢“ì∞¢'&íÊg&ˆ“Üñ◊ñÁFıvÜG4FV◊∆FT66ÜRÊ∂Wó2ÇííÊf˜$V6ÇÇÜ∂Wíí”‚∞¢ñbÜ∂WíÁ7F'G5vóFÇÜG∑6V∆V7FVD6ˆ÷÷W76ñG”¶íbb7FófT∂Wó2ÊÜ2Ü∂Wíííñ◊ñÁFıvÜG4FV◊∆FT66ÜRÊFV∆WFRÜ∂Wíì∞¢“ì∞ß–†¶gVÊ7Fñˆ‚ñÁf∆ñFFTñ◊ñÁFıvÜG4FV◊∆FRÜñ◊ñÁFÙ˜$ñG2í∞¢6ˆÁ7BñG2“'&íÊó4'&íÜñ◊ñÁFÙ˜$ñG2íÚñ◊ñÁFÙ˜$ñG2¢vWDñ◊ñÁFÙFˆ4ñG2Üñ◊ñÁFÙ˜$ñG2ì∞¢6ˆÁ7BñE6WB“ÊWr6WBÜñG2Êfñ«FW"Ñ&ˆˆ∆V‚íì∞¢6ˆÁ7B÷F6ÜñÊr“7W'&VÁDñ◊ñÁFíÊfñ«FW"ÇÜóFV“í”‚vWDñ◊ñÁFÙFˆ4ñG2ÜóFV“íÁ6ˆ÷RÇÜñBí”‚ñE6WBÊÜ2ÜñBííì∞¢÷F6ÜñÊrÊf˜$V6ÇÇÜñ◊ñÁFÚí”‚ñ◊ñÁFıvÜG4FV◊∆FT66ÜRÊFV∆WFRÜvWDñ◊ñÁFıvÜG4FV◊∆FT66ÜT∂WíÜñ◊ñÁFÚííì∞ß–†¶gVÊ7Fñˆ‚6∆V$ñ◊ñÁFıvÜG4FV◊∆FT66ÜRÇí∞¢ñ◊ñÁFıvÜG4FV◊∆FT66ÜRÊ6∆V"Çì∞ß–†¶gVÊ7Fñˆ‚'Vñ∆Dñ◊ñÁFıvÜG4ñ∆ˆBÜñ◊ñÁFÚ¬˜FñˆÁ2“∑“í∞¢6ˆÁ7BFˆÊTB“˜FñˆÁ2ÊFˆÊTB«¬ñ◊ñÁFÚÊFˆÊTB«¬ÊWrFFRÇì∞¢6ˆÁ7BFˆÊTñÊfÚ“f˜&÷DFˆÊTFFUFñ÷RÜFˆÊTBì∞¢6ˆÁ7BFñ÷R“FˆÊTñÊfÚÁFñ÷R””“"“"ÚÊWrFFRÇíÁFÙ∆ˆ6∆UFñ÷U7G&ñÊrÇ&óB‘ïB"¬≤Ü˜W#¢#"÷FñvóB"¬÷ñÁWFS¢#"÷FñvóB"¬Ü˜W##¢f«6R“í¢FˆÊTñÊfÚÁFñ÷S∞¢6ˆÁ7BWÜV7WFñˆ‰Ê˜FR“7G&ñÊrÜ˜FñˆÁ2ÊÊ˜FR«¬˜FñˆÁ2ÊWÜV7WFñˆ‰Ê˜FR«¬ñ◊ñÁFÚÊWÜV7WFñˆ‰Ê˜FR«¬""íÁG&ñ“Çì∞¢6ˆÁ7B&W&VEFV◊∆FR“&W&Tñ◊ñÁFıvÜG4FV◊∆FRÜñ◊ñÁFÚìÚÁFV◊∆FR«¬'Vñ∆Dñ◊ñÁFıvÜG4FV◊∆FRÜñ◊ñÁFÚì∞¢6ˆÁ7B÷W76vR“∞¢&W&VEFV◊∆FR¿¢	˘Y"˜&¢G∑Fñ÷W÷¿¢WÜV7WFñˆ‰Ê˜FRÚ	˘9“Ê˜FW6V7W¶ñˆÊS¢G∂WÜV7WFñˆ‰Ê˜FW÷¢" ¢“Êfñ«FW"Ñ&ˆˆ∆V‚íÊ¶ˆñ‚Ç%∆‚"ì∞†¢6ˆÁ7BVÊ6ˆFVD÷W76vR“VÊ6ˆFUU$î6ˆ◊ˆÊVÁBÜ÷W76vRì∞¢&WGW&‚∞¢÷W76vR¿¢W&√¢vÜG6¢Ú˜6VÊC˜FWáC“G∂VÊ6ˆFVD÷W76vW÷¿¢vV%W&√¢áGG3¢Ú˜vÊ÷RÛ˜FWáC“G∂VÊ6ˆFVD÷W76vW÷ ¢”∞ß–†¶gVÊ7Fñˆ‚vWEvÜßßWÜ˜FÙ∂WíÜñ◊ñÁFÚí∞¢&WGW&‚Gµ7G&ñÊrá6V∆V7FVD6ˆ÷÷W76ñB«¬""ó”£¢G∂'Vñ∆Dñ◊ñÁFÙ∂WíÜñ◊ñÁFÚó÷∞ß–†¶gVÊ7Fñˆ‚˜VÂvÜßßWÜ˜FÙFF&6RÇí∞¢ñbÇvñÊF˜rÊñÊFWÜVDD"í&WGW&‚&ˆ÷ó6RÁ&V¶V7BÜÊWrW'&˜"Ç$ñÊFWÜVDD"Êˆ‚Fó7ˆÊñ&ñ∆R"íì∞¢&WGW&‚ÊWr&ˆ÷ó6RÇá&W6ˆ«fR¬&V¶V7Bí”‚∞¢6ˆÁ7B&WVW7B“vñÊF˜rÊñÊFWÜVDD"Ê˜V‚ÖtÑ••UıÑıDıÙD%Ù‰‘R¬ì∞¢&WVW7BÊˆÁWw&FVÊVVFVB“Çí”‚∞¢6ˆÁ7BFF&6R“&WVW7BÁ&W7V«C∞¢ñbÇFF&6RÊˆ&¶V7E7F˜&TÊ÷W2Ê6ˆÁFñÁ2ÖtÑ••UıÑıDıı5Dı$UÙ‰‘Ríí∞¢FF&6RÊ7&VFTˆ&¶V7E7F˜&RÖtÑ••UıÑıDıı5Dı$UÙ‰‘R¬≤∂WïFÉ¢&∂Wí"“ì∞¢–¢”∞¢&WVW7BÊˆÁ7V66W72“Çí”‚&W6ˆ«fRá&WVW7BÁ&W7V«Bì∞¢&WVW7BÊˆÊW'&˜"“Çí”‚&V¶V7Bá&WVW7BÊW'&˜"«¬ÊWrW'&˜"Ç$&6ÜófñÚf˜FÚÊˆ‚66W76ñ&ñ∆R"íì∞¢“ì∞ß–†¶gVÊ7Fñˆ‚Ê˜&÷∆ó¶UvÜßßWÜ˜FÙÊ˜FW2ÜÊ˜FW2¬Ü˜FÙ6˜VÁBí∞¢6ˆÁ7BÊ˜FT∆ó7B“'&íÊg&ˆ“ÜÊ˜FW2«¬µ“ì∞¢&WGW&‚'&íÊg&ˆ“á≤∆VÊwFÉ¢÷FÇÊ÷ÇÉ¬ÁV÷&W"áÜ˜FÙ6˜VÁB«¬íí“¬ÖÚ¬ñÊFWÇí”‚Ä¢7G&ñÊrÜÊ˜FT∆ó7E∂ñÊFWÖ“«¬""íÁG&ñ“ÇíÁ6∆ñ6RÉ¬Éê¢íì∞ß–†¶7ñÊ2gVÊ7Fñˆ‚W'6ó7EvÜßßWÜ˜F˜2Ü∂Wí¬fñ∆W2¬6fVDB“FFRÊÊ˜rÇí¬Ê˜FW2“µ“í∞¢6ˆÁ7BFF&6R“vóB˜VÂvÜßßWÜ˜FÙFF&6RÇì∞¢G'í∞¢vóBÊWr&ˆ÷ó6RÇá&W6ˆ«fR¬&V¶V7Bí”‚∞¢6ˆÁ7BG&Á67Fñˆ‚“FF&6RÁG&Á67Fñˆ‚ÖtÑ••UıÑıDıı5Dı$UÙ‰‘R¬'&VGw&óFR"ì∞¢G&Á67Fñˆ‚Êˆ&¶V7E7F˜&RÖtÑ••UıÑıDıı5Dı$UÙ‰‘RíÁWBá∞¢∂Wí¿¢fñ∆W2¿¢6fVDB¿¢Ê˜FW3¢Ê˜&÷∆ó¶UvÜßßWÜ˜FÙÊ˜FW2ÜÊ˜FW2¬fñ∆W2Ê∆VÊwFÇê¢“ì∞¢G&Á67Fñˆ‚ÊˆÊ6ˆ◊∆WFR“Çí”‚&W6ˆ«fRÇì∞¢G&Á67Fñˆ‚ÊˆÊW'&˜"“Çí”‚&V¶V7BáG&Á67Fñˆ‚ÊW'&˜"«¬ÊWrW'&˜"Ç%6«fFvvñÚf˜FÚÊˆ‚&óW66óFÚ"íì∞¢G&Á67Fñˆ‚ÊˆÊ&˜'B“Çí”‚&V¶V7BáG&Á67Fñˆ‚ÊW'&˜"«¬ÊWrW'&˜"Ç%6«fFvvñÚf˜FÚñÁFW'&˜GFÚ"íì∞¢“ì∞¢“fñÊ∆«í∞¢FF&6RÊ6∆˜6RÇì∞¢–ß–†¶7ñÊ2gVÊ7Fñˆ‚FV∆WFUW'6ó7FVEvÜßßWÜ˜F˜2Ü∂Wíí∞¢vÜßßWÜ˜FÙfñ∆W4'îñ◊ñÁFÚÊFV∆WFRÜ∂Wíì∞¢vÜßßWÜ˜FÙÊ˜FW4'îñ◊ñÁFÚÊFV∆WFRÜ∂Wíì∞¢vÜßßWÜ˜Fı6fVDD'îñ◊ñÁFÚÊFV∆WFRÜ∂Wíì∞¢G'í∞¢6ˆÁ7BFF&6R“vóB˜VÂvÜßßWÜ˜FÙFF&6RÇì∞¢vóBÊWr&ˆ÷ó6RÇá&W6ˆ«fR¬&V¶V7Bí”‚∞¢6ˆÁ7BG&Á67Fñˆ‚“FF&6RÁG&Á67Fñˆ‚ÖtÑ••UıÑıDıı5Dı$UÙ‰‘R¬'&VGw&óFR"ì∞¢G&Á67Fñˆ‚Êˆ&¶V7E7F˜&RÖtÑ••UıÑıDıı5Dı$UÙ‰‘RíÊFV∆WFRÜ∂Wíì∞¢G&Á67Fñˆ‚ÊˆÊ6ˆ◊∆WFR“Çí”‚&W6ˆ«fRÇì∞¢G&Á67Fñˆ‚ÊˆÊW'&˜"“Çí”‚&V¶V7BáG&Á67Fñˆ‚ÊW'&˜"«¬ÊWrW'&˜"Ç$V∆ñ÷ñÊ¶ñˆÊRf˜FÚÊˆ‚&óW66óF"íì∞¢“ì∞¢FF&6RÊ6∆˜6RÇì∞¢“6F6ÇÜW'&˜"í∞¢6ˆÁ6ˆ∆RÁv&‚Ç%V∆ó¶ñf˜FÚvÜßßW∆ˆ6∆RÊˆ‚&óW66óF¢"¬W'&˜"ì∞¢–ß–†¶7ñÊ2gVÊ7Fñˆ‚&W7F˜&UW'6ó7FVEvÜßßWÜ˜F˜2Çí∞¢G'í∞¢6ˆÁ7BFF&6R“vóB˜VÂvÜßßWÜ˜FÙFF&6RÇì∞¢6ˆÁ7B&V6˜&G2“vóBÊWr&ˆ÷ó6RÇá&W6ˆ«fR¬&V¶V7Bí”‚∞¢6ˆÁ7BG&Á67Fñˆ‚“FF&6RÁG&Á67Fñˆ‚ÖtÑ••UıÑıDıı5Dı$UÙ‰‘R¬'&VFˆÊ«í"ì∞¢6ˆÁ7B&WVW7B“G&Á67Fñˆ‚Êˆ&¶V7E7F˜&RÖtÑ••UıÑıDıı5Dı$UÙ‰‘RíÊvWD∆¬Çì∞¢&WVW7BÊˆÁ7V66W72“Çí”‚&W6ˆ«fRá&WVW7BÁ&W7V«B«¬µ“ì∞¢&WVW7BÊˆÊW'&˜"“Çí”‚&V¶V7Bá&WVW7BÊW'&˜"«¬ÊWrW'&˜"Ç$∆WGGW&f˜FÚÊˆ‚&óW66óF"íì∞¢“ì∞¢FF&6RÊ6∆˜6RÇì∞¢6ˆÁ7BÊ˜r“FFRÊÊ˜rÇì∞¢6ˆÁ7BWáó&VD∂Wó2“µ”∞¢&V6˜&G2Êf˜$V6ÇÇá&V6˜&Bí”‚∞¢6ˆÁ7B∂Wí“7G&ñÊrá&V6˜&CÚÊ∂Wí«¬""ì∞¢6ˆÁ7B6fVDB“ÁV÷&W"á&V6˜&CÚÁ6fVDB«¬ì∞¢6ˆÁ7Bfñ∆W2“'&íÊg&ˆ“á&V6˜&CÚÊfñ∆W2«¬µ“íÊfñ«FW"ÇÜfñ∆Rí”‚fñ∆RñÁ7FÊ6Vˆb&∆ˆ"ì∞¢ñbÇ∂Wí«¬6fVDB«¬Ê˜r“6fVDB„“tÑ••UıÑıDıÙ‘ÖÙtUÙ’2«¬fñ∆W2Ê∆VÊwFÇí∞¢ñbÜ∂WííWáó&VD∂Wó2ÁW6ÇÜ∂Wíì∞¢&WGW&„∞¢–¢vÜßßWÜ˜FÙfñ∆W4'îñ◊ñÁFÚÁ6WBÜ∂Wí¬fñ∆W2ì∞¢vÜßßWÜ˜FÙÊ˜FW4'îñ◊ñÁFÚÁ6WBÜ∂Wí¬Ê˜&÷∆ó¶UvÜßßWÜ˜FÙÊ˜FW2á&V6˜&CÚÊÊ˜FW2¬fñ∆W2Ê∆VÊwFÇíì∞¢vÜßßWÜ˜Fı6fVDD'îñ◊ñÁFÚÁ6WBÜ∂Wí¬6fVDBì∞¢“ì∞¢vóB&ˆ÷ó6RÊ∆¬ÜWáó&VD∂Wó2Ê÷ÇÜ∂Wíí”‚FV∆WFUW'6ó7FVEvÜßßWÜ˜F˜2Ü∂Wíííì∞¢ñbá6V∆V7FVD6ˆ÷÷W76ñBbbGóVˆb&VÊFW$ñ◊ñÁFí””“&gVÊ7Fñˆ‚"í&VÊFW$ñ◊ñÁFíÇì∞¢“6F6ÇÜW'&˜"í∞¢6ˆÁ6ˆ∆RÁv&‚Ç%&ó&ó7FñÊÚf˜FÚvÜßßW∆ˆ6∆RÊˆ‚Fó7ˆÊñ&ñ∆S¢"¬W'&˜"ì∞¢–ß–†¢ÚÚñÊFWÜVDD"Êˆ‚fñVÊR6Ê6V∆∆FÚF¬&Vg&W6ÇFV∆¬v¢&ó&ó7FñÊv∆í∆∆VvFê¢ÚÚ∆ˆ6∆íÊ6˜&f∆ñFíR&ñ◊V˜fRWFˆ÷Fñ6÷VÁFRVV∆∆íú;ífV66ÜíFí˜&R‡ßfˆñB&W7F˜&UW'6ó7FVEvÜßßWÜ˜F˜2Çì∞†¶gVÊ7Fñˆ‚vWEvÜßßWÜ˜F˜2Üñ◊ñÁFÚí∞¢6ˆÁ7B∂Wí“vWEvÜßßWÜ˜FÙ∂WíÜñ◊ñÁFÚì∞¢6ˆÁ7B6fVDB“ÁV÷&W"ávÜßßWÜ˜Fı6fVDD'îñ◊ñÁFÚÊvWBÜ∂Wíí«¬ì∞¢ñbá6fVDBbbFFRÊÊ˜rÇí“6fVDB„“tÑ••UıÑıDıÙ‘ÖÙtUÙ’2í∞¢fˆñBFV∆WFUW'6ó7FVEvÜßßWÜ˜F˜2Ü∂Wíì∞¢&WGW&‚µ”∞¢–¢&WGW&‚vÜßßWÜ˜FÙfñ∆W4'îñ◊ñÁFÚÊvWBÜ∂Wíí«¬µ”∞ß–†¶gVÊ7Fñˆ‚vWEvÜßßWÜ˜FÙÊ˜FW2Üñ◊ñÁFÚí∞¢6ˆÁ7Bfñ∆W2“vWEvÜßßWÜ˜F˜2Üñ◊ñÁFÚì∞¢&WGW&‚Ê˜&÷∆ó¶UvÜßßWÜ˜FÙÊ˜FW2ávÜßßWÜ˜FÙÊ˜FW4'îñ◊ñÁFÚÊvWBÜvWEvÜßßWÜ˜FÙ∂WíÜñ◊ñÁFÚíí¬fñ∆W2Ê∆VÊwFÇì∞ß–†¶gVÊ7Fñˆ‚WFFUvÜßßWGF6Ü÷VÁD'WGFˆ‚Ü'WGFˆ‚¬ñ◊ñÁFÚí∞¢ñbÇ'WGFˆ‚í&WGW&„∞¢6ˆÁ7B6˜VÁB“vWEvÜßßWÜ˜F˜2Üñ◊ñÁFÚíÊ∆VÊwFÉ∞¢'WGFˆ‚ÁFWáD6ˆÁFVÁB“6˜VÁBÚ	˘8‚G∂6˜VÁG÷¢/	˘8‚#∞¢'WGFˆ‚ÊFF6WBÊGF6Ü÷VÁD6˜VÁB“6˜VÁBÚ7G&ñÊrÜ6˜VÁBí¢"#∞¢'WGFˆ‚ÁFóF∆R“6˜VÁBÚG∂6˜VÁG“f˜FÚ6V∆W¶ñˆÊFR‚&V÷íW"÷ˆFñfñ6&∆V¢$∆∆VvVÊÚú;íf˜FÚ¬÷W76vvñÚvÜßßW#∞¢'WGFˆ‚Á6WDGG&ñ'WFRÇ&&ñ÷∆&V¬"¬'WGFˆ‚ÁFóF∆Rì∞¢'WGFˆ‚Ê6∆74∆ó7BÁFˆvv∆RÇ&Ü2÷GF6Ü÷VÁG2"¬6˜VÁB‚ì∞ß–†¶7ñÊ2gVÊ7Fñˆ‚6fUvÜßßWÜ˜Fı6V∆V7Fñˆ‚Üñ◊ñÁFÚ¬fñ∆W2¬Ê˜FW2“µ“í∞¢6ˆÁ7B∂Wí“vWEvÜßßWÜ˜FÙ∂WíÜñ◊ñÁFÚì∞¢6ˆÁ7Bf∆ñDfñ∆W2“'&íÊg&ˆ“Üfñ∆W2«¬µ“íÊfñ«FW"ÇÜfñ∆Rí”‚fñ∆RñÁ7FÊ6Vˆb&∆ˆ"bb7G&ñÊrÜfñ∆RÁGóR«¬""íÁ7F'G5vóFÇÇ&ñ÷vRÚ"íì∞¢ñbÇf∆ñDfñ∆W2Ê∆VÊwFÇí∞¢vóBFV∆WFUW'6ó7FVEvÜßßWÜ˜F˜2Ü∂Wíì∞¢&WGW&‚µ”∞¢–¢6ˆÁ7B6fVDB“FFRÊÊ˜rÇì∞¢6ˆÁ7BÊ˜&÷∆ó¶VDÊ˜FW2“Ê˜&÷∆ó¶UvÜßßWÜ˜FÙÊ˜FW2ÜÊ˜FW2¬f∆ñDfñ∆W2Ê∆VÊwFÇì∞¢vÜßßWÜ˜FÙfñ∆W4'îñ◊ñÁFÚÁ6WBÜ∂Wí¬f∆ñDfñ∆W2ì∞¢vÜßßWÜ˜FÙÊ˜FW4'îñ◊ñÁFÚÁ6WBÜ∂Wí¬Ê˜&÷∆ó¶VDÊ˜FW2ì∞¢vÜßßWÜ˜Fı6fVDD'îñ◊ñÁFÚÁ6WBÜ∂Wí¬6fVDBì∞¢vóBW'6ó7EvÜßßWÜ˜F˜2Ü∂Wí¬f∆ñDfñ∆W2¬6fVDB¬Ê˜&÷∆ó¶VDÊ˜FW2ì∞¢&WGW&‚f∆ñDfñ∆W3∞ß–†¶7ñÊ2gVÊ7Fñˆ‚6fUvÜßßWÜ˜FÙÊ˜FW2Üñ◊ñÁFÚ¬Ê˜FW2í∞¢6ˆÁ7B∂Wí“vWEvÜßßWÜ˜FÙ∂WíÜñ◊ñÁFÚì∞¢6ˆÁ7Bfñ∆W2“vWEvÜßßWÜ˜F˜2Üñ◊ñÁFÚì∞¢ñbÇfñ∆W2Ê∆VÊwFÇí&WGW&‚µ”∞¢6ˆÁ7BÊ˜&÷∆ó¶VDÊ˜FW2“Ê˜&÷∆ó¶UvÜßßWÜ˜FÙÊ˜FW2ÜÊ˜FW2¬fñ∆W2Ê∆VÊwFÇì∞¢6ˆÁ7B6fVDB“ÁV÷&W"ávÜßßWÜ˜Fı6fVDD'îñ◊ñÁFÚÊvWBÜ∂Wíí«¬FFRÊÊ˜rÇíì∞¢vÜßßWÜ˜FÙÊ˜FW4'îñ◊ñÁFÚÁ6WBÜ∂Wí¬Ê˜&÷∆ó¶VDÊ˜FW2ì∞¢vóBW'6ó7EvÜßßWÜ˜F˜2Ü∂Wí¬fñ∆W2¬6fVDB¬Ê˜&÷∆ó¶VDÊ˜FW2ì∞¢&WGW&‚Ê˜&÷∆ó¶VDÊ˜FW3∞ß–†¶gVÊ7Fñˆ‚˜VÂvÜßßWÜ˜Fı&WfñWrÜfñ∆Rí∞¢6ˆÁ7Bˆ&¶V7EW&¬“U$¬Ê7&VFTˆ&¶V7EU$¬Üfñ∆Rì∞¢6ˆÁ7B˜fW&∆í“Fˆ7V÷VÁBÊ7&VFTV∆V÷VÁBÇ&Fób"ì∞¢˜fW&∆íÊ6∆74Ê÷R“'vÜßßW◊Ü˜FÚ◊&WfñWr#∞¢˜fW&∆íÁ6WDGG&ñ'WFRÇ'&ˆ∆R"¬&Fñ∆ˆr"ì∞¢˜fW&∆íÁ6WDGG&ñ'WFRÇ&&ñ÷÷ˆF¬"¬'G'VR"ì∞¢˜fW&∆íÁ6WDGG&ñ'WFRÇ&&ñ÷∆&V¬"¬$ÁFW&ñ÷f˜FÚ∆∆VvF"ì∞¢˜fW&∆íÊñÊÊW$ÖD‘¬“ ¢∆'WGFˆ‚GóS“&'WGFˆ‚"6∆73“'vÜßßW◊Ü˜FÚ◊&WfñWr÷6∆˜6R"&ñ÷∆&V√“$6ÜóVFíÁFW&ñ÷#Ï9s¬ˆ'WGFˆ„‡¢∆ñ÷r7&3“"G∂ˆ&¶V7EW&«“"«C“$f˜FÚ∆∆VvF¬÷W76vvñÚvÜßßW#‡¢∞¢6ˆÁ7B6∆˜6R“Çí”‚∞¢U$¬Á&Wfˆ∂Tˆ&¶V7EU$¬Üˆ&¶V7EW&¬ì∞¢˜fW&∆íÁ&V÷˜fRÇì∞¢Fˆ7V÷VÁBÁ&V÷˜fTWfVÁD∆ó7FVÊW"Ç&∂WñF˜v‚"¬ˆ‰∂WîF˜v‚¬G'VRì∞¢”∞¢6ˆÁ7Bˆ‰∂WîF˜v‚“ÜWfVÁBí”‚∞¢ñbÜWfVÁBÊ∂Wí””“$W66R"í∞¢WfVÁBÁ7F˜ñ÷÷VFñFU&˜vFñˆ‚Çì∞¢6∆˜6RÇì∞¢–¢”∞¢˜fW&∆íÊFDWfVÁD∆ó7FVÊW"Ç&6∆ñ6≤"¬ÜWfVÁBí”‚∞¢ñbÜWfVÁBÁF&vWB””“˜fW&∆í«¬WfVÁBÁF&vWBÊ6∆˜6W7BÇ"ÁvÜßßW◊Ü˜FÚ◊&WfñWr÷6∆˜6R"íí6∆˜6RÇì∞¢“ì∞¢Fˆ7V÷VÁBÊFDWfVÁD∆ó7FVÊW"Ç&∂WñF˜v‚"¬ˆ‰∂WîF˜v‚¬G'VRì∞¢Fˆ7V÷VÁBÊ&ˆGíÊVÊD6Üñ∆BÜ˜fW&∆íì∞¢˜fW&∆íÁVW'ï6V∆V7F˜"Ç"ÁvÜßßW◊Ü˜FÚ◊&WfñWr÷6∆˜6R"ìÚÊfˆ7W2Çì∞ß–†¶gVÊ7Fñˆ‚ñ6µvÜßßWÜ˜F˜2Üñ◊ñÁFÚ¬'WGFˆ‚¬˜FñˆÁ2“∑“í∞¢6ˆÁ7BñÁWB“Fˆ7V÷VÁBÊ7&VFTV∆V÷VÁBÇ&ñÁWB"ì∞¢ñÁWBÁGóR“&fñ∆R#∞¢ñÁWBÊ66WB“&ñ÷vRÚ¢#∞¢ñÁWBÊ◊V«Fó∆R“˜FñˆÁ2Ê÷ˆFR”“'&W∆6R÷ˆÊR#∞¢ñÁWBÊÜñFFV‚“G'VS∞¢ñÁWBÊFDWfVÁD∆ó7FVÊW"Ç&6ÜÊvR"¬7ñÊ2Çí”‚∞¢6ˆÁ7Bfñ∆W2“'&íÊg&ˆ“ÜñÁWBÊfñ∆W2«¬µ“íÊfñ«FW"ÇÜfñ∆Rí”‚fñ∆RÁGóRÁ7F'G5vóFÇÇ&ñ÷vRÚ"íì∞¢ñÁWBÁ&V÷˜fRÇì∞¢ñbÇfñ∆W2Ê∆VÊwFÇí&WGW&„∞¢6ˆÁ7B7W'&VÁB“vWEvÜßßWÜ˜F˜2Üñ◊ñÁFÚíÁ6∆ñ6RÇì∞¢6ˆÁ7B7W'&VÁDÊ˜FW2“vWEvÜßßWÜ˜FÙÊ˜FW2Üñ◊ñÁFÚì∞¢∆WBÊWáDfñ∆W2“fñ∆W3∞¢∆WBÊWáDÊ˜FW2“fñ∆W2Ê÷ÇÇí”‚""ì∞¢ñbÜ˜FñˆÁ2Ê÷ˆFR””“&VÊB"íÊWáDfñ∆W2“≤‚‚Ê7W'&VÁB¬‚‚Êfñ∆W5”∞¢ñbÜ˜FñˆÁ2Ê÷ˆFR””“&VÊB"íÊWáDÊ˜FW2“≤‚‚Ê7W'&VÁDÊ˜FW2¬‚‚Êfñ∆W2Ê÷ÇÇí”‚""ï”∞¢ñbÜ˜FñˆÁ2Ê÷ˆFR””“'&W∆6R÷ˆÊR"í∞¢ÊWáDfñ∆W2“7W'&VÁC∞¢ÊWáDfñ∆W2Á7∆ñ6RÑÁV÷&W"Ü˜FñˆÁ2ÊñÊFWÇ«¬í¬¬fñ∆W5≥“ì∞¢ÊWáDÊ˜FW2“7W'&VÁDÊ˜FW3∞¢–¢G'í∞¢vóB6fUvÜßßWÜ˜Fı6V∆V7Fñˆ‚Üñ◊ñÁFÚ¬ÊWáDfñ∆W2¬ÊWáDÊ˜FW2ì∞¢WFFUvÜßßWGF6Ü÷VÁD'WGFˆ‚Ü'WGFˆ‚¬ñ◊ñÁFÚì∞¢ñbÜ˜FñˆÁ2Á&V˜V‰÷ÊvW"”“f«6Rí˜VÂvÜßßWÜ˜FÙ÷ÊvW"Üñ◊ñÁFÚ¬'WGFˆ‚ì∞¢“6F6ÇÜW'&˜"í∞¢6ˆÁ6ˆ∆RÊW'&˜"Ç%6«fFvvñÚ∆ˆ6∆Rf˜FÚvÜßßWÊˆ‚&óW66óFÛ¢"¬W'&˜"ì∞¢∆W'BÇ$Êˆ‚:Ç7FFÚ˜76ñ&ñ∆R6ˆÁ6W'f&R∆Rf˜FÚ7V¬Fó7˜6óFófÚ‚&ó&˜f˜W&R∆ñ&W&7¶ñÚ7V¬FV∆VfˆÊÚ‚"ì∞¢–¢“¬≤ˆÊ6S¢G'VR“ì∞¢ñÁWBÊFDWfVÁD∆ó7FVÊW"Ç&6Ê6V¬"¬Çí”‚ñÁWBÁ&V÷˜fRÇí¬≤ˆÊ6S¢G'VR“ì∞¢Fˆ7V÷VÁBÊ&ˆGíÊVÊD6Üñ∆BÜñÁWBì∞¢ñÁWBÊ6∆ñ6≤Çì∞ß–†¶gVÊ7Fñˆ‚˜VÂvÜßßWÜ˜FÙ÷ÊvW"Üñ◊ñÁFÚ¬'WGFˆ‚í∞¢6ˆÁ7Bfñ∆W2“vWEvÜßßWÜ˜F˜2Üñ◊ñÁFÚì∞¢ñbÇfñ∆W2Ê∆VÊwFÇí∞¢ñ6µvÜßßWÜ˜F˜2Üñ◊ñÁFÚ¬'WGFˆ‚¬≤÷ˆFS¢'&W∆6R÷∆¬"“ì∞¢&WGW&„∞¢–¢6ˆÁ7B˜fW&∆í“Fˆ7V÷VÁBÊ7&VFTV∆V÷VÁBÇ&Fób"ì∞¢˜fW&∆íÊ6∆74Ê÷R“'vÜßßW◊Ü˜FÚ÷÷ÊvW"#∞¢˜fW&∆íÁ6WDGG&ñ'WFRÇ'&ˆ∆R"¬&Fñ∆ˆr"ì∞¢˜fW&∆íÁ6WDGG&ñ'WFRÇ&&ñ÷÷ˆF¬"¬'G'VR"ì∞¢˜fW&∆íÁ6WDGG&ñ'WFRÇ&&ñ÷∆&V∆∆VF'í"¬'vÜßßW◊Ü˜FÚ÷÷ÊvW"◊FóF∆R"ì∞¢6ˆÁ7B6&B“Fˆ7V÷VÁBÊ7&VFTV∆V÷VÁBÇ'6V7Fñˆ‚"ì∞¢6&BÊ6∆74Ê÷R“'vÜßßW◊Ü˜FÚ÷÷ÊvW"÷6&B#∞¢˜fW&∆íÊVÊD6Üñ∆BÜ6&Bì∞¢∆WBˆ&¶V7EW&«2“µ”∞¢∆WBÊ˜FU6fUFñ÷W"“ÁV∆√∞†¢6ˆÁ7B&VEfó6ñ&∆TÊ˜FW2“Çí”‚∞¢6ˆÁ7B7W'&VÁDÊ˜FW2“vWEvÜßßWÜ˜FÙÊ˜FW2Üñ◊ñÁFÚì∞¢6&BÁVW'ï6V∆V7F˜$∆¬Ç%∂FF◊Ü˜FÚ÷Ê˜FR÷ñÊFWÖ“"íÊf˜$V6ÇÇÜñÁWBí”‚∞¢6ˆÁ7BñÊFWÇ“ÁV÷&W"ÜñÁWBÊFF6WBÁÜ˜FÙÊ˜FTñÊFWÇì∞¢ñbÑÁV÷&W"Êó4ñÁFVvW"ÜñÊFWÇíbbñÊFWÇ„“bbñÊFWÇ¬7W'&VÁDÊ˜FW2Ê∆VÊwFÇí∞¢7W'&VÁDÊ˜FW5∂ñÊFWÖ““7G&ñÊrÜñÁWBÁf«VR«¬""íÁG&ñ“ÇíÁ6∆ñ6RÉ¬Éì∞¢–¢“ì∞¢&WGW&‚7W'&VÁDÊ˜FW3∞¢”∞†¢6ˆÁ7Bf«W6Öfó6ñ&∆TÊ˜FW2“7ñÊ2Çí”‚∞¢ñbÜÊ˜FU6fUFñ÷W"ívñÊF˜rÊ6∆V%Fñ÷V˜WBÜÊ˜FU6fUFñ÷W"ì∞¢Ê˜FU6fUFñ÷W"“ÁV∆√∞¢&WGW&‚6fUvÜßßWÜ˜FÙÊ˜FW2Üñ◊ñÁFÚ¬&VEfó6ñ&∆TÊ˜FW2Çíì∞¢”∞†¢6ˆÁ7B6∆˜6R“Çí”‚∞¢ñbÜÊ˜FU6fUFñ÷W"í∞¢vñÊF˜rÊ6∆V%Fñ÷V˜WBÜÊ˜FU6fUFñ÷W"ì∞¢Ê˜FU6fUFñ÷W"“ÁV∆√∞¢fˆñB6fUvÜßßWÜ˜FÙÊ˜FW2Üñ◊ñÁFÚ¬&VEfó6ñ&∆TÊ˜FW2Çíì∞¢–¢ˆ&¶V7EW&«2Êf˜$V6ÇÇáW&¬í”‚U$¬Á&Wfˆ∂Tˆ&¶V7EU$¬áW&¬íì∞¢ˆ&¶V7EW&«2“µ”∞¢˜fW&∆íÁ&V÷˜fRÇì∞¢Fˆ7V÷VÁBÁ&V÷˜fTWfVÁD∆ó7FVÊW"Ç&∂WñF˜v‚"¬ˆ‰∂WîF˜v‚ì∞¢”∞¢6ˆÁ7Bˆ‰∂WîF˜v‚“ÜWfVÁBí”‚∞¢ñbÜWfVÁBÊ∂Wí””“$W66R"í6∆˜6RÇì∞¢”∞†¢6ˆÁ7B&VÊFW"“Çí”‚∞¢ˆ&¶V7EW&«2Êf˜$V6ÇÇáW&¬í”‚U$¬Á&Wfˆ∂Tˆ&¶V7EU$¬áW&¬íì∞¢ˆ&¶V7EW&«2“µ”∞¢6ˆÁ7B7W'&VÁDfñ∆W2“vWEvÜßßWÜ˜F˜2Üñ◊ñÁFÚì∞¢6ˆÁ7B7W'&VÁDÊ˜FW2“vWEvÜßßWÜ˜FÙÊ˜FW2Üñ◊ñÁFÚì∞¢ñbÇ7W'&VÁDfñ∆W2Ê∆VÊwFÇí∞¢6∆˜6RÇì∞¢WFFUvÜßßWGF6Ü÷VÁD'WGFˆ‚Ü'WGFˆ‚¬ñ◊ñÁFÚì∞¢&WGW&„∞¢–¢6ˆÁ7BÜ˜FÙ6&G2“7W'&VÁDfñ∆W2Ê÷ÇÜfñ∆R¬ñÊFWÇí”‚∞¢6ˆÁ7BW&¬“U$¬Ê7&VFTˆ&¶V7EU$¬Üfñ∆Rì∞¢ˆ&¶V7EW&«2ÁW6ÇáW&¬ì∞¢&WGW&‚ ¢∆'Fñ6∆R6∆73“'vÜßßW◊Ü˜FÚ÷óFV“"FF◊Ü˜FÚ÷ñÊFWÉ“"G∂ñÊFWá“#‡¢∆'WGFˆ‚GóS“&'WGFˆ‚"6∆73“'vÜßßW◊Ü˜FÚ÷˜V‚"FF◊Ü˜FÚ÷7Fñˆ„“'fñWr"&ñ÷∆&V√“%fó7V∆óß¶f˜FÚG∂ñÊFWÇ≤“#‡¢∆ñ÷r7&3“"G∑W&«“"«C“$f˜FÚ∆∆VvFG∂ñÊFWÇ≤“#‡¢«7‚6∆73“'vÜßßW◊Ü˜FÚ÷˜&FW"#‰f˜FÚG∂ñÊFWÇ≤”¬˜7„‡¢«7‚6∆73“'vÜßßW◊Ü˜FÚ◊fñWr÷∆&V¬#Âfó7V∆óß¶¬˜7„‡¢¬ˆ'WGFˆ„‡¢∆∆&V¬6∆73“'vÜßßW◊Ü˜FÚ÷Ê˜FR#‡¢«7„‰Ê˜FFV∆∆f˜FÚ«6÷∆√‚Ü˜¶ñˆÊ∆Rì¬˜6÷∆√„¬˜7„‡¢«FWáF&VFF◊Ü˜FÚ÷Ê˜FR÷ñÊFWÉ“"G∂ñÊFWá“"÷Ü∆VÊwFÉ“#É"&˜w3“#""∆6VÜˆ∆FW#“$W2‚∆&W&Ú6GWFÚ÷'¶&˜GFÚ#‚G∂W66TÖD‘¬Ü7W'&VÁDÊ˜FW5∂ñÊFWÖ“ó”¬˜FWáF&V‡¢¬ˆ∆&V√‡¢∆Fób6∆73“'vÜßßW◊Ü˜FÚ÷óFV“÷7FñˆÁ2#‡¢∆'WGFˆ‚GóS“&'WGFˆ‚"6∆73“&'F‚"FF◊Ü˜FÚ÷7Fñˆ„“'&W∆6R#Â6˜7FóGVó66ì¬ˆ'WGFˆ„‡¢∆'WGFˆ‚GóS“&'WGFˆ‚"6∆73“&'F‚'F‚÷FÊvW""FF◊Ü˜FÚ÷7Fñˆ„“&FV∆WFR#‰V∆ñ÷ñÊ¬ˆ'WGFˆ„‡¢¬ˆFóc‡¢¬ˆ'Fñ6∆S‡¢∞¢“íÊ¶ˆñ‚Ç""ì∞¢6ˆÁ7B7V&÷óD∆&V¬“ñ◊ñÁFÚÊFˆÊRÚ$îÂdîdıDÚ5RtÑ••U"¢$dEDÚRîÂdîtÑ••U#∞¢6&BÊñÊÊW$ÖD‘¬“ ¢∆ÜVFW"6∆73“'vÜßßW◊Ü˜FÚ÷÷ÊvW"÷ÜVB#‡¢∆Fóc‡¢«6∆73“&÷ÊvV÷VÁB÷WñV'&˜r#‰ƒƒTtDítÑ••U¬˜‡¢∆É"ñC“'vÜßßW◊Ü˜FÚ÷÷ÊvW"◊FóF∆R#‰f˜FÚFV∆Œ(	ññ◊ñÁFÛ¬ˆÉ#‡¢«‚G∂7W'&VÁDfñ∆W2Ê∆VÊwFá“G∂7W'&VÁDfñ∆W2Ê∆VÊwFÇ””“Ú&f˜FÚ∆∆VvF"¢&f˜FÚ∆∆VvFR'“+r6ˆÁ6W'fFRW"÷76ñ÷Ú˜&S¬˜‡¢¬ˆFóc‡¢∆'WGFˆ‚GóS“&'WGFˆ‚"6∆73“'vÜßßW◊Ü˜FÚ÷÷ÊvW"÷6∆˜6R"FF÷÷ÊvW"÷7Fñˆ„“&6∆˜6R"&ñ÷∆&V√“$6ÜóVFí#Ï9s¬ˆ'WGFˆ„‡¢¬ˆÜVFW#‡¢∆Fób6∆73“'vÜßßW◊Ü˜FÚ÷w&ñB#‚G∑Ü˜FÙ6&G7”¬ˆFóc‡¢∆fˆ˜FW"6∆73“'vÜßßW◊Ü˜FÚ÷÷ÊvW"÷7FñˆÁ2#‡¢∆'WGFˆ‚GóS“&'WGFˆ‚"6∆73“&'F‚"FF÷÷ÊvW"÷7Fñˆ„“&FB#Ó˚»≤vvóVÊvíf˜FÛ¬ˆ'WGFˆ„‡¢∆'WGFˆ‚GóS“&'WGFˆ‚"6∆73“&'F‚"FF÷÷ÊvW"÷7Fñˆ„“'&W∆6R÷∆¬#Â6˜7FóGVó66íGWGFS¬ˆ'WGFˆ„‡¢∆'WGFˆ‚GóS“&'WGFˆ‚"6∆73“&'F‚'F‚÷FÊvW""FF÷÷ÊvW"÷7Fñˆ„“&FV∆WFR÷∆¬#‰V∆ñ÷ñÊGWGFS¬ˆ'WGFˆ„‡¢∆'WGFˆ‚GóS“&'WGFˆ‚"6∆73“&'F‚vÜßßW◊Ü˜FÚ÷÷ÊvW"◊7V&÷óB"FF÷÷ÊvW"÷7Fñˆ„“&FˆÊR#Ó)»RG∑7V&÷óD∆&V«”¬ˆ'WGFˆ„‡¢¬ˆfˆ˜FW#‡¢∞¢”∞†¢6&BÊFDWfVÁD∆ó7FVÊW"Ç&6∆ñ6≤"¬7ñÊ2ÜWfVÁBí”‚∞¢6ˆÁ7B÷ÊvW$7Fñˆ‚“WfVÁBÁF&vWBÊ6∆˜6W7BÇ%∂FF÷÷ÊvW"÷7FñˆÂ“"ìÚÊFF6WBÊ÷ÊvW$7Fñˆ„∞¢ñbÜ÷ÊvW$7Fñˆ‚””“&6∆˜6R"í&WGW&‚6∆˜6RÇì∞¢ñbÜ÷ÊvW$7Fñˆ‚””“&FˆÊR"í∞¢6ˆÁ7B7V&÷óD'WGFˆ‚“WfVÁBÁF&vWBÊ6∆˜6W7BÇ%∂FF÷÷ÊvW"÷7Fñˆ„“vFˆÊRu“"ì∞¢ñbÇ7V&÷óD'WGFˆ‚«¬7V&÷óD'WGFˆ‚ÊFó6&∆VB«¬ó4ñ◊ñÁFıvÜßßW&ˆ6W76ñÊrÜñ◊ñÁFÚíí&WGW&„∞¢7V&÷óD'WGFˆ‚ÊFó6&∆VB“G'VS∞¢7V&÷óD'WGFˆ‚ÁFWáD6ˆÁFVÁB“ñ◊ñÁFÚÊFˆÊRÚ$W'GW&vÜßßW(
b"¢%6«fFvvñ˛(
b#∞¢G'í∞¢vóBf«W6Öfó6ñ&∆TÊ˜FW2Çì∞¢“6F6ÇÜW'&˜"í∞¢6ˆÁ6ˆ∆RÁv&‚Ç$6ˆÁ6W'f¶ñˆÊR∆ˆ6∆RÊ˜FRf˜FÚvÜßßWÊˆ‚&óW66óF¢"¬W'&˜"ì∞¢–¢6∆˜6RÇì∞¢ñbÜñ◊ñÁFÚÊFˆÊRívóBÜÊF∆T6ˆ◊∆WFVDñ◊ñÁFıvÜG46∆ñ6≤Üñ◊ñÁFÚì∞¢V«6RvóBÜÊF∆Tñ◊ñÁFıvÜG46∆ñ6≤Üñ◊ñÁFÚì∞¢&WGW&„∞¢–¢ñbÜ÷ÊvW$7Fñˆ‚””“&FB"«¬÷ÊvW$7Fñˆ‚””“'&W∆6R÷∆¬"í∞¢6∆˜6RÇì∞¢ñ6µvÜßßWÜ˜F˜2Üñ◊ñÁFÚ¬'WGFˆ‚¬≤÷ˆFS¢÷ÊvW$7Fñˆ‚””“&FB"Ú&VÊB"¢'&W∆6R÷∆¬"“ì∞¢&WGW&„∞¢–¢ñbÜ÷ÊvW$7Fñˆ‚””“&FV∆WFR÷∆¬"í∞¢ñbÇ6ˆÊfó&“Ç$V∆ñ÷ñÊ&RGWGFR∆Rf˜FÚ∆∆VvFRVW7FÚñ◊ñÁFÛÚ"íí&WGW&„∞¢vóB6fUvÜßßWÜ˜Fı6V∆V7Fñˆ‚Üñ◊ñÁFÚ¬µ“ì∞¢WFFUvÜßßWGF6Ü÷VÁD'WGFˆ‚Ü'WGFˆ‚¬ñ◊ñÁFÚì∞¢6∆˜6RÇì∞¢&WGW&„∞¢–¢6ˆÁ7BóFV““WfVÁBÁF&vWBÊ6∆˜6W7BÇ%∂FF◊Ü˜FÚ÷ñÊFWÖ“"ì∞¢6ˆÁ7BÜ˜FÙ7Fñˆ‚“WfVÁBÁF&vWBÊ6∆˜6W7BÇ%∂FF◊Ü˜FÚ÷7FñˆÂ“"ìÚÊFF6WBÁÜ˜FÙ7Fñˆ„∞¢ñbÇóFV“«¬Ü˜FÙ7Fñˆ‚í&WGW&„∞¢6ˆÁ7BñÊFWÇ“ÁV÷&W"ÜóFV“ÊFF6WBÁÜ˜FÙñÊFWÇì∞¢6ˆÁ7B7W'&VÁDfñ∆W2“vWEvÜßßWÜ˜F˜2Üñ◊ñÁFÚì∞¢ñbÇ7W'&VÁDfñ∆W5∂ñÊFWÖ“í&WGW&„∞¢ñbáÜ˜FÙ7Fñˆ‚””“'fñWr"í&WGW&‚˜VÂvÜßßWÜ˜Fı&WfñWrÜ7W'&VÁDfñ∆W5∂ñÊFWÖ“ì∞¢ñbáÜ˜FÙ7Fñˆ‚””“'&W∆6R"í∞¢6∆˜6RÇì∞¢ñ6µvÜßßWÜ˜F˜2Üñ◊ñÁFÚ¬'WGFˆ‚¬≤÷ˆFS¢'&W∆6R÷ˆÊR"¬ñÊFWÇ“ì∞¢&WGW&„∞¢–¢ñbáÜ˜FÙ7Fñˆ‚””“&FV∆WFR"í∞¢6ˆÁ7B&V÷ñÊñÊr“7W'&VÁDfñ∆W2Êfñ«FW"ÇÖÚ¬Ü˜FÙñÊFWÇí”‚Ü˜FÙñÊFWÇ”“ñÊFWÇì∞¢6ˆÁ7B&V÷ñÊñÊtÊ˜FW2“vWEvÜßßWÜ˜FÙÊ˜FW2Üñ◊ñÁFÚíÊfñ«FW"ÇÖÚ¬Ü˜FÙñÊFWÇí”‚Ü˜FÙñÊFWÇ”“ñÊFWÇì∞¢vóB6fUvÜßßWÜ˜Fı6V∆V7Fñˆ‚Üñ◊ñÁFÚ¬&V÷ñÊñÊr¬&V÷ñÊñÊtÊ˜FW2ì∞¢WFFUvÜßßWGF6Ü÷VÁD'WGFˆ‚Ü'WGFˆ‚¬ñ◊ñÁFÚì∞¢&VÊFW"Çì∞¢–¢“ì∞¢6&BÊFDWfVÁD∆ó7FVÊW"Ç&ñÁWB"¬ÜWfVÁBí”‚∞¢6ˆÁ7BÊ˜FTñÁWB“WfVÁBÁF&vWBÊ6∆˜6W7BÇ%∂FF◊Ü˜FÚ÷Ê˜FR÷ñÊFWÖ“"ì∞¢ñbÇÊ˜FTñÁWBí&WGW&„∞¢6ˆÁ7B∂Wí“vWEvÜßßWÜ˜FÙ∂WíÜñ◊ñÁFÚì∞¢vÜßßWÜ˜FÙÊ˜FW4'îñ◊ñÁFÚÁ6WBÜ∂Wí¬&VEfó6ñ&∆TÊ˜FW2Çíì∞¢ñbÜÊ˜FU6fUFñ÷W"ívñÊF˜rÊ6∆V%Fñ÷V˜WBÜÊ˜FU6fUFñ÷W"ì∞¢Ê˜FU6fUFñ÷W"“vñÊF˜rÁ6WEFñ÷V˜WBÇÇí”‚∞¢Ê˜FU6fUFñ÷W"“ÁV∆√∞¢fˆñB6fUvÜßßWÜ˜FÙÊ˜FW2Üñ◊ñÁFÚ¬&VEfó6ñ&∆TÊ˜FW2Çíì∞¢“¬CSì∞¢“ì∞¢˜fW&∆íÊFDWfVÁD∆ó7FVÊW"Ç&6∆ñ6≤"¬ÜWfVÁBí”‚∞¢ñbÜWfVÁBÁF&vWB””“˜fW&∆íí6∆˜6RÇì∞¢“ì∞¢Fˆ7V÷VÁBÊFDWfVÁD∆ó7FVÊW"Ç&∂WñF˜v‚"¬ˆ‰∂WîF˜v‚ì∞¢Fˆ7V÷VÁBÊ&ˆGíÊVÊD6Üñ∆BÜ˜fW&∆íì∞¢&VÊFW"Çì∞¢6&BÁVW'ï6V∆V7F˜"Ç"ÁvÜßßW◊Ü˜FÚ÷÷ÊvW"÷6∆˜6R"ìÚÊfˆ7W2Çì∞ß–†¶gVÊ7Fñˆ‚6Üˆ˜6UvÜßßWÜ˜F˜2Üñ◊ñÁFÚ¬'WGFˆ‚í∞¢ñbÜvWEvÜßßWÜ˜F˜2Üñ◊ñÁFÚíÊ∆VÊwFÇí˜VÂvÜßßWÜ˜FÙ÷ÊvW"Üñ◊ñÁFÚ¬'WGFˆ‚ì∞¢V«6Rñ6µvÜßßWÜ˜F˜2Üñ◊ñÁFÚ¬'WGFˆ‚¬≤÷ˆFS¢'&W∆6R÷∆¬"“ì∞ß–†¶gVÊ7Fñˆ‚'Vñ∆D˜&FW&VEvÜßßW6Ü&Tfñ∆W2Üfñ∆W2í∞¢&WGW&‚'&íÊg&ˆ“Üfñ∆W2«¬µ“íÊ÷ÇÜfñ∆R¬ñÊFWÇí”‚∞¢6ˆÁ7B˜&ñvñÊƒÊ÷R“7G&ñÊrÜfñ∆SÚÊÊ÷R«¬""ì∞¢6ˆÁ7BWáFVÁ6ñˆ‰÷F6Ç“˜&ñvñÊƒÊ÷RÊ÷F6ÇÇı¬‚Ö∂◊§’£”ï◊≥"√W“íBÚì∞¢6ˆÁ7B÷ñ÷TWáFVÁ6ñˆ‚“7G&ñÊrÜfñ∆SÚÁGóR«¬""íÁ7∆óBÇ"Ú"ï≥”ÚÁ&W∆6RÇ&ßVr"¬&ßr"í«¬&ßr#∞¢6ˆÁ7BWáFVÁ6ñˆ‚“WáFVÁ6ñˆ‰÷F6ÉÚÂ≥”ÚÁFÙ∆˜vW$66RÇí«¬÷ñ÷TWáFVÁ6ñˆ„∞¢6ˆÁ7B˜&FW&VDÊ÷R“f˜FÚ“Gµ7G&ñÊrÜñÊFWÇ≤íÁE7F'BÉ"¬#"ó“‚G∂WáFVÁ6ñˆÁ÷∞¢G'í∞¢&WGW&‚ÊWrfñ∆RÖ∂fñ∆U“¬˜&FW&VDÊ÷R¬∞¢GóS¢fñ∆RÁGóR«¬&ñ÷vRˆßVr"¿¢∆7D÷ˆFñfñVC¢ÁV÷&W"Üfñ∆RÊ∆7D÷ˆFñfñVB«¬FFRÊÊ˜rÇíê¢“ì∞¢“6F6ÇÜW'&˜"í∞¢6ˆÁ6ˆ∆RÁv&‚Ç$Êˆ÷R&ˆw&W76ófÚf˜FÚÊˆ‚∆ñ6&ñ∆S¢÷ÁFVÊvÚñ¬fñ∆R˜&ñvñÊ∆R‚"¬W'&˜"ì∞¢&WGW&‚fñ∆S∞¢–¢“ì∞ß–†¶gVÊ7Fñˆ‚∆ˆEvÜßßWÜ˜FÙf˜$ÊÊ˜FFñˆ‚Üfñ∆Rí∞¢&WGW&‚ÊWr&ˆ÷ó6RÇá&W6ˆ«fR¬&V¶V7Bí”‚∞¢6ˆÁ7Bˆ&¶V7EW&¬“U$¬Ê7&VFTˆ&¶V7EU$¬Üfñ∆Rì∞¢6ˆÁ7Bñ÷vR“ÊWrñ÷vRÇì∞¢ñ÷vRÊˆÊ∆ˆB“Çí”‚&W6ˆ«fRá≤ñ÷vR¬ˆ&¶V7EW&¬“ì∞¢ñ÷vRÊˆÊW'&˜"“Çí”‚∞¢U$¬Á&Wfˆ∂Tˆ&¶V7EU$¬Üˆ&¶V7EW&¬ì∞¢&V¶V7BÜÊWrW'&˜"Ç$∆f˜FÚÊˆ‚\;"W76W&R&W&F6ˆ‚∆Ê˜F"íì∞¢”∞¢ñ÷vRÁ7&2“ˆ&¶V7EW&√∞¢“ì∞ß–†¶gVÊ7Fñˆ‚w&vÜßßWÜ˜FÙÊ˜FRÜ6ˆÁFWáB¬FWáB¬÷ÖvñGFÇ¬÷Ñ∆ñÊW2í∞¢6ˆÁ7Bv˜&G2“7G&ñÊráFWáB«¬""íÁG&ñ“ÇíÁ7∆óBÇı«2≤ÚíÊfñ«FW"Ñ&ˆˆ∆V‚íÊf∆D÷Çáv˜&Bí”‚∞¢ñbÜ6ˆÁFWáBÊ÷V7W&UFWáBáv˜&BíÁvñGFÇ√“÷ÖvñGFÇí&WGW&‚∑v˜&E”∞¢6ˆÁ7B6áVÊ∑2“µ”∞¢∆WB6áVÊ≤“"#∞¢'&íÊg&ˆ“áv˜&BíÊf˜$V6ÇÇÜ6Ü&7FW"í”‚∞¢ñbÜ6áVÊ≤bb6ˆÁFWáBÊ÷V7W&UFWáBÜG∂6áVÊ∑“G∂6Ü&7FW'÷íÁvñGFÇ‚÷ÖvñGFÇí∞¢6áVÊ∑2ÁW6ÇÜ6áVÊ≤ì∞¢6áVÊ≤“6Ü&7FW#∞¢“V«6R∞¢6áVÊ≤≥“6Ü&7FW#∞¢–¢“ì∞¢ñbÜ6áVÊ≤í6áVÊ∑2ÁW6ÇÜ6áVÊ≤ì∞¢&WGW&‚6áVÊ∑3∞¢“ì∞¢6ˆÁ7B∆ñÊW2“µ”∞¢∆WB7W'&VÁD∆ñÊR“"#∞¢v˜&G2Êf˜$V6ÇÇáv˜&Bí”‚∞¢6ˆÁ7B6ÊFñFFR“7W'&VÁD∆ñÊRÚG∂7W'&VÁD∆ñÊW“G∑v˜&G÷¢v˜&C∞¢ñbÇ7W'&VÁD∆ñÊR«¬6ˆÁFWáBÊ÷V7W&UFWáBÜ6ÊFñFFRíÁvñGFÇ√“÷ÖvñGFÇí∞¢7W'&VÁD∆ñÊR“6ÊFñFFS∞¢&WGW&„∞¢–¢∆ñÊW2ÁW6ÇÜ7W'&VÁD∆ñÊRì∞¢7W'&VÁD∆ñÊR“v˜&C∞¢“ì∞¢ñbÜ7W'&VÁD∆ñÊRí∆ñÊW2ÁW6ÇÜ7W'&VÁD∆ñÊRì∞¢ñbÜ∆ñÊW2Ê∆VÊwFÇ√“÷Ñ∆ñÊW2í&WGW&‚∆ñÊW3∞¢6ˆÁ7Bfó6ñ&∆T∆ñÊW2“∆ñÊW2Á6∆ñ6RÉ¬÷Ñ∆ñÊW2ì∞¢∆WB∆7D∆ñÊR“fó6ñ&∆T∆ñÊW5∂÷Ñ∆ñÊW2“”∞¢vÜñ∆RÜ∆7D∆ñÊRbb6ˆÁFWáBÊ÷V7W&UFWáBÜG∂∆7D∆ñÊWﬁ(
fíÁvñGFÇ‚÷ÖvñGFÇí∞¢∆7D∆ñÊR“∆7D∆ñÊRÁ6∆ñ6RÉ¬”íÁG&ñ‘VÊBÇì∞¢–¢fó6ñ&∆T∆ñÊW5∂÷Ñ∆ñÊW2“““G∂∆7D∆ñÊWﬁ(
f∞¢&WGW&‚fó6ñ&∆T∆ñÊW3∞ß–†¶7ñÊ2gVÊ7Fñˆ‚FEvÜßßWÊ˜FUFıÜ˜FÚÜfñ∆R¬Ê˜FRí∞¢6ˆÁ7BÊ˜&÷∆ó¶VDÊ˜FR“7G&ñÊrÜÊ˜FR«¬""íÁG&ñ“ÇíÁ6∆ñ6RÉ¬Éì∞¢ñbÇÊ˜&÷∆ó¶VDÊ˜FRí&WGW&‚fñ∆S∞†¢6ˆÁ7B≤ñ÷vR¬ˆ&¶V7EW&¬““vóB∆ˆEvÜßßWÜ˜FÙf˜$ÊÊ˜FFñˆ‚Üfñ∆Rì∞¢G'í∞¢6ˆÁ7B÷ÑFñ÷VÁ6ñˆ‚“Cìc∞¢6ˆÁ7B66∆R“÷FÇÊ÷ñ‚É¬÷ÑFñ÷VÁ6ñˆ‚Ú÷FÇÊ÷ÇÜñ÷vRÊÊGW&≈vñGFÇ¬ñ÷vRÊÊGW&ƒÜVñváBíì∞¢6ˆÁ7B6Áf2“Fˆ7V÷VÁBÊ7&VFTV∆V÷VÁBÇ&6Áf2"ì∞¢6Áf2ÁvñGFÇ“÷FÇÊ÷ÇÉ¬÷FÇÁ&˜VÊBÜñ÷vRÊÊGW&≈vñGFÇ¢66∆Ríì∞¢6Áf2ÊÜVñváB“÷FÇÊ÷ÇÉ¬÷FÇÁ&˜VÊBÜñ÷vRÊÊGW&ƒÜVñváB¢66∆Ríì∞¢6ˆÁ7B6ˆÁFWáB“6Áf2ÊvWD6ˆÁFWáBÇ#&B"¬≤«Ü¢f«6R“ì∞¢ñbÇ6ˆÁFWáBíFá&˜rÊWrW'&˜"Ç$V∆&˜&¶ñˆÊRf˜FÚÊˆ‚Fó7ˆÊñ&ñ∆R"ì∞¢6ˆÁFWáBÊG&tñ÷vRÜñ÷vR¬¬¬6Áf2ÁvñGFÇ¬6Áf2ÊÜVñváBì∞†¢6ˆÁ7BFFñÊr“÷FÇÊ÷ÇÉ#"¬÷FÇÁ&˜VÊBÜ6Áf2ÁvñGFÇ¢„3Ríì∞¢6ˆÁ7B66VÁEvñGFÇ“÷FÇÊ÷ÇÉÇ¬÷FÇÁ&˜VÊBÜ6Áf2ÁvñGFÇ¢„ííì∞¢6ˆÁ7BFóF∆U6ó¶R“÷FÇÊ÷ÇÉ"¬÷FÇÊ÷ñ‚ÉC"¬÷FÇÁ&˜VÊBÜ6Áf2ÁvñGFÇ¢„3"í¬÷FÇÁ&˜VÊBÜ6Áf2ÊÜVñváB¢„S"ííì∞¢6ˆÁ7BFWáE6ó¶R“÷FÇÊ÷ÇÉb¬÷FÇÊ÷ñ‚ÉSÇ¬÷FÇÁ&˜VÊBÜ6Áf2ÁvñGFÇ¢„CÇí¬÷FÇÁ&˜VÊBÜ6Áf2ÊÜVñváB¢„rííì∞¢6ˆÁ7B∆ñÊTÜVñváB“÷FÇÁ&˜VÊBáFWáE6ó¶R¢„#Bì∞¢6ˆÁFWáBÊfˆÁB“sG∑FWáE6ó¶W◊Ç6Á2◊6W&ñf∞¢6ˆÁ7B∆ñÊW2“w&vÜßßWÜ˜FÙÊ˜FRÜ6ˆÁFWáB¬Ê˜&÷∆ó¶VDÊ˜FR¬6Áf2ÁvñGFÇ“áFFñÊr¢"í“66VÁEvñGFÇ¬"ì∞¢6ˆÁ7BÊVƒÜVñváB“FFñÊr≤FóF∆U6ó¶R≤÷FÇÁ&˜VÊBáFóF∆U6ó¶R¢„SRí≤Ü∆ñÊW2Ê∆VÊwFÇ¢∆ñÊTÜVñváBí≤FFñÊs∞¢6ˆÁ7BÊV≈F˜“÷FÇÊ÷ÇÉ¬6Áf2ÊÜVñváB“ÊVƒÜVñváBì∞†¢6ˆÁFWáBÊfñ∆≈7Gñ∆R“'&v&ÉB¬3R¬3¬„ÉÇí#∞¢6ˆÁFWáBÊfñ∆≈&V7BÉ¬ÊV≈F˜¬6Áf2ÁvñGFÇ¬6Áf2ÊÜVñváB“ÊV≈F˜ì∞¢6ˆÁFWáBÊfñ∆≈7Gñ∆R“"6cV3SC"#∞¢6ˆÁFWáBÊfñ∆≈&V7BÉ¬ÊV≈F˜¬66VÁEvñGFÇ¬6Áf2ÊÜVñváB“ÊV≈F˜ì∞¢6ˆÁFWáBÊfñ∆≈7Gñ∆R“"3vfc3R#∞¢6ˆÁFWáBÊfˆÁB“ÉG∑FóF∆U6ó¶W◊Ç6Á2◊6W&ñf∞¢6ˆÁFWáBÊfñ∆≈FWáBÇ$‰ıDdıDÚ"¬FFñÊr≤66VÁEvñGFÇ¬ÊV≈F˜≤FFñÊr≤FóF∆U6ó¶Rì∞¢6ˆÁFWáBÊfñ∆≈7Gñ∆R“"6fffffb#∞¢6ˆÁFWáBÊfˆÁB“sG∑FWáE6ó¶W◊Ç6Á2◊6W&ñf∞¢6ˆÁ7Bfó'7EFWáD&6V∆ñÊR“ÊV≈F˜≤FFñÊr≤FóF∆U6ó¶R≤÷FÇÁ&˜VÊBáFóF∆U6ó¶R¢„SRí≤FWáE6ó¶S∞¢∆ñÊW2Êf˜$V6ÇÇÜ∆ñÊR¬ñÊFWÇí”‚∞¢6ˆÁFWáBÊfñ∆≈FWáBÜ∆ñÊR¬FFñÊr≤66VÁEvñGFÇ¬fó'7EFWáD&6V∆ñÊR≤ÜñÊFWÇ¢∆ñÊTÜVñváBíì∞¢“ì∞†¢6ˆÁ7B˜WGWEGóR“7G&ñÊrÜfñ∆RÁGóR«¬""íÁFÙ∆˜vW$66RÇí””“&ñ÷vR˜Êr"Ú&ñ÷vR˜Êr"¢&ñ÷vRˆßVr#∞¢6ˆÁ7B&∆ˆ"“vóBÊWr&ˆ÷ó6RÇá&W6ˆ«fR¬&V¶V7Bí”‚∞¢6Áf2ÁFÙ&∆ˆ"Çá&W7V«Bí”‚&W7V«BÚ&W6ˆ«fRá&W7V«Bí¢&V¶V7BÜÊWrW'&˜"Ç$7&V¶ñˆÊRf˜FÚ6ˆ‚Ê˜FÊˆ‚&óW66óF"íí¬˜WGWEGóR¬„ì"ì∞¢“ì∞¢6ˆÁ7B˜WGWDWáFVÁ6ñˆ‚“˜WGWEGóR””“&ñ÷vR˜Êr"Ú'Êr"¢&ßr#∞¢6ˆÁ7B˜&ñvñÊƒ&6TÊ÷R“7G&ñÊrÜfñ∆RÊÊ÷R«¬&f˜FÚ÷Ê˜F"íÁ&W∆6RÇı¬Âµ‚Â“≤BÚ¬""í«¬&f˜FÚ÷Ê˜F#∞¢&WGW&‚ÊWrfñ∆RÖ∂&∆ˆ%“¬G∂˜&ñvñÊƒ&6TÊ÷W“‚G∂˜WGWDWáFVÁ6ñˆÁ÷¬∞¢GóS¢˜WGWEGóR¿¢∆7D÷ˆFñfñVC¢ÁV÷&W"Üfñ∆RÊ∆7D÷ˆFñfñVB«¬FFRÊÊ˜rÇíê¢“ì∞¢“fñÊ∆«í∞¢U$¬Á&Wfˆ∂Tˆ&¶V7EU$¬Üˆ&¶V7EW&¬ì∞¢–ß–†¶7ñÊ2gVÊ7Fñˆ‚'Vñ∆EvÜßßW6Ü&Tfñ∆W5vóFÑÊ˜FW2Üñ◊ñÁFÚ¬fñ∆W2í∞¢6ˆÁ7BÊ˜FW2“vWEvÜßßWÜ˜FÙÊ˜FW2Üñ◊ñÁFÚì∞¢6ˆÁ7BÊÊ˜FFVDfñ∆W2“µ”∞¢f˜"Ü∆WBñÊFWÇ“≤ñÊFWÇ¬fñ∆W2Ê∆VÊwFÉ≤ñÊFWÇ≥“í∞¢ÊÊ˜FFVDfñ∆W2ÁW6ÇÜvóBFEvÜßßWÊ˜FUFıÜ˜FÚÜfñ∆W5∂ñÊFWÖ“¬Ê˜FW5∂ñÊFWÖ“íì∞¢–¢&WGW&‚'Vñ∆D˜&FW&VEvÜßßW6Ü&Tfñ∆W2ÜÊÊ˜FFVDfñ∆W2ì∞ß–†¶gVÊ7Fñˆ‚vWDÊFófTÊG&ˆñEvÜßßW6Ü&U«VvñÁ2Çí∞¢6ˆÁ7B66óF˜"“vñÊF˜r‰66óF˜#∞¢6ˆÁ7Bó4ÊFófTÊG&ˆñB“&ˆˆ∆V‚Ä¢66óF˜#ÚÊó4ÊFófU∆Ff˜&”Ú‚Çê¢bb66óF˜#ÚÊvWE∆Ff˜&”Ú‚Çí””“&ÊG&ˆñB ¢ì∞¢ñbÇó4ÊFófTÊG&ˆñBí&WGW&‚ÁV∆√∞†¢6ˆÁ7B&Vvó7FW%«Vvñ‚“GóVˆb66óF˜#ÚÁ&Vvó7FW%«Vvñ‚””“&gVÊ7Fñˆ‚ ¢Ú66óF˜"Á&Vvó7FW%«Vvñ‚Ê&ñÊBÜ66óF˜"ê¢¢ÁV∆√∞¢6ˆÁ7Bfñ∆W7ó7FV““66óF˜"Â«VvñÁ3Ú‰fñ∆W7ó7FV“«¬&Vvó7FW%«Vvñ„Ú‚Ç$fñ∆W7ó7FV“"ì∞¢6ˆÁ7B6Ü&R“66óF˜"Â«VvñÁ3ÚÂ6Ü&R«¬&Vvó7FW%«Vvñ„Ú‚Ç%6Ü&R"ì∞¢&WGW&‚fñ∆W7ó7FV“bb6Ü&RÚ≤fñ∆W7ó7FV“¬6Ü&R“¢ÁV∆√∞ß–†¶gVÊ7Fñˆ‚vWDFVFñ6FVDÊG&ˆñEvÜßßWÜ˜Fı«Vvñ‚Çí∞¢6ˆÁ7B66óF˜"“vñÊF˜r‰66óF˜#∞¢6ˆÁ7Bó4ÊFófTÊG&ˆñB“&ˆˆ∆V‚Ä¢66óF˜#ÚÊó4ÊFófU∆Ff˜&”Ú‚Çê¢bb66óF˜#ÚÊvWE∆Ff˜&”Ú‚Çí””“&ÊG&ˆñB ¢ì∞¢ñbÇó4ÊFófTÊG&ˆñBí&WGW&‚ÁV∆√∞†¢6ˆÁ7B&Vvó7FW%«Vvñ‚“GóVˆb66óF˜#ÚÁ&Vvó7FW%«Vvñ‚””“&gVÊ7Fñˆ‚ ¢Ú66óF˜"Á&Vvó7FW%«Vvñ‚Ê&ñÊBÜ66óF˜"ê¢¢ÁV∆√∞¢6ˆÁ7B«Vvñ‚“66óF˜"Â«VvñÁ3Ú‰ÜW&vÜßßWÜ˜F˜2«¬&Vvó7FW%«Vvñ„Ú‚Ç$ÜW&vÜßßWÜ˜F˜2"ì∞¢&WGW&‚«Vvñ‡¢bbGóVˆb«Vvñ‚Ê&Vvñ‚””“&gVÊ7Fñˆ‚ ¢bbGóVˆb«Vvñ‚ÊFEÜ˜FÚ””“&gVÊ7Fñˆ‚ ¢bbGóVˆb«Vvñ‚Á6Ü&R””“&gVÊ7Fñˆ‚ ¢Ú«Vvñ‡¢¢ÁV∆√∞ß–†¶gVÊ7Fñˆ‚&VEvÜßßWÜ˜FÙ4&6ScBÜfñ∆Rí∞¢&WGW&‚ÊWr&ˆ÷ó6RÇá&W6ˆ«fR¬&V¶V7Bí”‚∞¢6ˆÁ7B&VFW"“ÊWrfñ∆U&VFW"Çì∞¢&VFW"ÊˆÊ∆ˆB“Çí”‚∞¢6ˆÁ7B&W7V«B“7G&ñÊrá&VFW"Á&W7V«B«¬""ì∞¢6ˆÁ7B6W&F˜$ñÊFWÇ“&W7V«BÊñÊFWÑˆbÇ"¬"ì∞¢ñbá6W&F˜$ñÊFWÇ¬í&WGW&‚&V¶V7BÜÊWrW'&˜"Ç$f˜&÷FÚf˜FÚÊˆ‚f∆ñFÚ"íì∞¢&W6ˆ«fRá&W7V«BÁ6∆ñ6Rá6W&F˜$ñÊFWÇ≤íì∞¢”∞¢&VFW"ÊˆÊW'&˜"“Çí”‚&V¶V7Bá&VFW"ÊW'&˜"«¬ÊWrW'&˜"Ç$∆WGGW&f˜FÚÊˆ‚&óW66óF"íì∞¢&VFW"Á&VD4FFU$¬Üfñ∆Rì∞¢“ì∞ß–†¶7ñÊ2gVÊ7Fñˆ‚&V÷˜fTÊFófUvÜßßW6Ü&Tfˆ∆FW"Üfñ∆W7ó7FV“¬fˆ∆FW%FÇí∞¢G'í∞¢vóBfñ∆W7ó7FV“Á&÷Fó"á∞¢FÉ¢fˆ∆FW%FÇ¿¢Fó&V7F˜'ì¢$44ÑR"¿¢&V7W'6ófS¢G'VP¢“ì∞¢“6F6ÇÜW'&˜"í∞¢6ˆÁ6ˆ∆RÁv&‚Ç%V∆ó¶ñf˜FÚFV◊˜&ÊVRvÜßßWÊˆ‚&óW66óF¢"¬W'&˜"ì∞¢–ß–†¶gVÊ7Fñˆ‚66ÜVGV∆TÊFófUvÜßßW6Ü&T6∆VÁWÜfñ∆W7ó7FV“¬fˆ∆FW%FÇí∞¢ÚÚvÜG4FWfRfW&Rñ¬FV◊ÚFí∆VvvW&Rv∆íU$í&ñ6WgWFíF¬ÊÊV∆∆ÚÊG&ˆñB‡¢vñÊF˜rÁ6WEFñ÷V˜WBÇÇí”‚∞¢fˆñB&V÷˜fTÊFófUvÜßßW6Ü&Tfˆ∆FW"Üfñ∆W7ó7FV“¬fˆ∆FW%FÇì∞¢“¬R¢c¢ì∞ß–†¶7ñÊ2gVÊ7Fñˆ‚6Ü&UvÜßßWÜ˜F˜4FVFñ6FVDÊG&ˆñBá«Vvñ‚¬˜&FW&VDfñ∆W2¬÷W76vRí∞¢∆WB6W76ñˆ‰ñB“"#∞¢G'í∞¢6ˆÁ7B6W76ñˆ‚“vóB«Vvñ‚Ê&Vvñ‚Çì∞¢6W76ñˆ‰ñB“7G&ñÊrá6W76ñˆ„ÚÁ6W76ñˆ‰ñB«¬""ì∞¢ñbÇ6W76ñˆ‰ñBíFá&˜rÊWrW'&˜"Ç%6W76ñˆÊRf˜FÚÊG&ˆñBÊˆ‚Fó7ˆÊñ&ñ∆R"ì∞†¢f˜"Ü6ˆÁ7Bfñ∆Rˆb˜&FW&VDfñ∆W2í∞¢vóB«Vvñ‚ÊFEÜ˜FÚá∞¢6W76ñˆ‰ñB¿¢fñ∆TÊ÷S¢7G&ñÊrÜfñ∆SÚÊÊ÷R«¬&f˜FÚÊßr"í¿¢÷ñ÷UGóS¢7G&ñÊrÜfñ∆SÚÁGóR«¬&ñ÷vRˆßVr"í¿¢FF¢vóB&VEvÜßßWÜ˜FÙ4&6ScBÜfñ∆Rê¢“ì∞¢–†¢&WGW&‚vóB«Vvñ‚Á6Ü&Rá∞¢6W76ñˆ‰ñB¿¢FWáC¢÷W76vP¢“ì∞¢“6F6ÇÜW'&˜"í∞¢ñbá6W76ñˆ‰ñBbbGóVˆb«Vvñ‚ÊFó66&B””“&gVÊ7Fñˆ‚"í∞¢G'í∞¢vóB«Vvñ‚ÊFó66&Bá≤6W76ñˆ‰ñB“ì∞¢“6F6ÇÖÚí∑–¢–¢Fá&˜rW'&˜#∞¢–ß–†¶7ñÊ2gVÊ7Fñˆ‚6Ü&UvÜßßWÜ˜F˜4ÊFófTÊG&ˆñBÜ˜&FW&VDfñ∆W2¬÷W76vRí∞¢6ˆÁ7BFVFñ6FVE«Vvñ‚“vWDFVFñ6FVDÊG&ˆñEvÜßßWÜ˜Fı«Vvñ‚Çì∞¢ñbÜFVFñ6FVE«Vvñ‚í∞¢&WGW&‚6Ü&UvÜßßWÜ˜F˜4FVFñ6FVDÊG&ˆñBÜFVFñ6FVE«Vvñ‚¬˜&FW&VDfñ∆W2¬÷W76vRì∞¢–†¢6ˆÁ7B«VvñÁ2“vWDÊFófTÊG&ˆñEvÜßßW6Ü&U«VvñÁ2Çì∞¢ñbÇ«VvñÁ2í&WGW&‚ÁV∆√∞¢6ˆÁ7Bfˆ∆FW%FÇ“ÜW&◊vÜßßW◊6Ü&RÚG¥FFRÊÊ˜rÇó““G¥÷FÇÁ&ÊFˆ“ÇíÁFı7G&ñÊrÉ3bíÁ6∆ñ6RÉ"¬ó÷∞¢6ˆÁ7Bfñ∆UW&ó2“µ”∞¢G'í∞¢f˜"Ü6ˆÁ7Bfñ∆Rˆb˜&FW&VDfñ∆W2í∞¢6ˆÁ7Bfñ∆TÊ÷R“7G&ñÊrÜfñ∆SÚÊÊ÷R«¬f˜FÚ“Gµ7G&ñÊrÜfñ∆UW&ó2Ê∆VÊwFÇ≤íÁE7F'BÉ"¬#"ó“Êßvì∞¢6ˆÁ7B&W7V«B“vóB«VvñÁ2Êfñ∆W7ó7FV“Áw&óFTfñ∆Rá∞¢FÉ¢G∂fˆ∆FW%Fá“ÚG∂fñ∆TÊ÷W÷¿¢FF¢vóB&VEvÜßßWÜ˜FÙ4&6ScBÜfñ∆Rí¿¢Fó&V7F˜'ì¢$44ÑR"¿¢&V7W'6ófS¢G'VP¢“ì∞¢fñ∆UW&ó2ÁW6Çá&W7V«BÁW&íì∞¢–¢6ˆÁ7B6Ü&U&W7V«B“vóB«VvñÁ2Á6Ü&RÁ6Ü&Rá∞¢fñ∆W3¢fñ∆UW&ó2¿¢FWáC¢÷W76vR¿¢FóF∆S¢$ñ◊ñÁFÚfGFÚ"¿¢Fñ∆ˆuFóF∆S¢$6ˆÊFófñFíf˜FÚR÷W76vvñÚvÜßßW ¢“ì∞¢66ÜVGV∆TÊFófUvÜßßW6Ü&T6∆VÁWá«VvñÁ2Êfñ∆W7ó7FV“¬fˆ∆FW%FÇì∞¢&WGW&‚6Ü&U&W7V«B«¬G'VS∞¢“6F6ÇÜW'&˜"í∞¢vóB&V÷˜fTÊFófUvÜßßW6Ü&Tfˆ∆FW"á«VvñÁ2Êfñ∆W7ó7FV“¬fˆ∆FW%FÇì∞¢Fá&˜rW'&˜#∞¢–ß–†¶gVÊ7Fñˆ‚ó5vÜßßW6Ü&T6Ê6V∆∆Fñˆ‚ÜW'&˜"í∞¢&WGW&‚W'&˜#ÚÊÊ÷R””“$&˜'DW'&˜""«¬ˆ6Ê6V«∆ÊÁV∆¬ˆíÁFW7BÖ7G&ñÊrÜW'&˜#ÚÊ÷W76vR«¬""íì∞ß–†¶7ñÊ2gVÊ7Fñˆ‚6Ü&UvÜßßWvóFÖÜ˜F˜2Üñ◊ñÁFÚ¬˜FñˆÁ2“∑“í∞¢6ˆÁ7Bfñ∆W2“vWEvÜßßWÜ˜F˜2Üñ◊ñÁFÚì∞¢ñbÇfñ∆W2Ê∆VÊwFÇí&WGW&‚ÁV∆√∞¢6ˆÁ7B≤÷W76vR““'Vñ∆Dñ◊ñÁFıvÜG4ñ∆ˆBÜñ◊ñÁFÚ¬˜FñˆÁ2ì∞¢G'í∞¢6ˆÁ7B˜&FW&VDfñ∆W2“vóB'Vñ∆EvÜßßW6Ü&Tfñ∆W5vóFÑÊ˜FW2Üñ◊ñÁFÚ¬fñ∆W2ì∞¢6ˆÁ7BÊFófU6Ü&U&W7V«B“vóB6Ü&UvÜßßWÜ˜F˜4ÊFófTÊG&ˆñBÜ˜&FW&VDfñ∆W2¬÷W76vRì∞¢ñbÇÊFófU6Ü&U&W7V«Bí∞¢ñbÇÊfñvF˜"Á6Ü&R«¬ÜÊfñvF˜"Ê6Â6Ü&RbbÊfñvF˜"Ê6Â6Ü&Rá≤fñ∆W3¢˜&FW&VDfñ∆W2“ííí∞¢∆W'BÇ$ñ¬Fó7˜6óFófÚÊˆ‚W&÷WGFR∆6ˆÊFófó6ñˆÊRFó&WGFFV∆∆Rf˜FÚ‚&ÚvÜßßW6ˆ‚ñ¬FW7FÛ¢∆∆Vv∆Rf˜FÚF∆∆w&ffWGFFívÜG4‚"ì∞¢&WGW&‚˜VÂvÜG4Üñ◊ñÁFÚ¬˜FñˆÁ2ì∞¢–¢vóBÊfñvF˜"Á6Ü&Rá∞¢fñ∆W3¢˜&FW&VDfñ∆W2¿¢FWáC¢÷W76vR¿¢FóF∆S¢$ñ◊ñÁFÚfGFÚ"¿¢“ì∞¢–¢vóBFV∆WFUW'6ó7FVEvÜßßWÜ˜F˜2ÜvWEvÜßßWÜ˜FÙ∂WíÜñ◊ñÁFÚíì∞¢ñbáGóVˆb&VÊFW$ñ◊ñÁFí””“&gVÊ7Fñˆ‚"í&VÊFW$ñ◊ñÁFíÇì∞¢&WGW&‚G'VS∞¢“6F6ÇÜW'&˜"í∞¢ñbÜó5vÜßßW6Ü&T6Ê6V∆∆Fñˆ‚ÜW'&˜"íí&WGW&‚f«6S∞¢6ˆÁ6ˆ∆RÊW'&˜"Ç$6ˆÊFófó6ñˆÊRf˜FÚvÜßßWÊˆ‚&óW66óF¢"¬W'&˜"ì∞¢∆W'BÇ$Êˆ‚:Ç7FFÚ˜76ñ&ñ∆R6ˆÊFófñFW&R∆Rf˜FÚ‚&ÚvÜßßW6ˆ‚ñ¬6ˆ∆ÚFW7FÚ‚"ì∞¢&WGW&‚˜VÂvÜG4Üñ◊ñÁFÚ¬˜FñˆÁ2ì∞¢–ß–†¶gVÊ7Fñˆ‚˜VÂvÜG4Üñ◊ñÁFÚ¬˜FñˆÁ2“∑“í∞¢6ˆÁ7BW6W"“WFÇÊ7W'&VÁEW6W#∞¢ñbÇW6W"í∞¢∆W'BÇ$FWfíf&R∆ˆvñ‚‚"ì∞¢&WGW&‚f«6S∞¢–†¢6ˆÁ7B≤÷W76vR¬W&¬““'Vñ∆Dñ◊ñÁFıvÜG4ñ∆ˆBÜñ◊ñÁFÚ¬˜FñˆÁ2ì∞¢6ˆÁ7BF&vWEvñÊF˜r“˜FñˆÁ3ÚÁF&vWEvñÊF˜s∞¢ñbáF&vWEvñÊF˜rbbF&vWEvñÊF˜rÊ6∆˜6VBí∞¢G'í∞¢F&vWEvñÊF˜rÊ∆ˆ6Fñˆ‚Á&W∆6RÜW&¬ì∞¢“6F6ÇÜW'&˜"í∞¢6ˆÁ6ˆ∆RÊW'&˜"Ç$W'&˜&RW'GW&vÜG4ÊV∆∆fñÊW7G&F&vWC¢"¬W'&˜"ì∞¢–¢&WGW&‚G'VS∞¢–¢6ˆÁ7BF&vWB“˜FñˆÁ3ÚÁF&vWB«¬%ˆ&∆Ê≤#∞¢6ˆÁ7B˜VÊVB“6fT˜VÂvÜG4÷W76vRÜ÷W76vR¬∞¢W&¬¿¢F&vWB¿¢W6U˜W¢F&vWB”“%˜6V∆b ¢“ì∞¢&WGW&‚&ˆˆ∆V‚Ü˜VÊVBì∞ß–†¶gVÊ7Fñˆ‚˜V‰ñ◊ñÁFı&W˜'D÷ˆF¬Üñ◊ñÁFÚí∞¢&W˜'FñÊtñ◊ñÁFÚ“ñ◊ñÁFÚ«¬ÁV∆√∞¢VíÊñ◊ñÁFı&W˜'Df˜&“Á&W6WBÇì∞¢VíÊñ◊ñÁFı&W˜'DfVVF&6≤ÁFWáD6ˆÁFVÁB“"#∞¢VíÊñ◊ñÁFı&W˜'D÷ˆF¬Ê6∆74∆ó7BÁ&V÷˜fRÇ&ÜñFFV‚"ì∞ß–†¶gVÊ7Fñˆ‚6∆˜6Tñ◊ñÁFı&W˜'D÷ˆF¬Çí∞¢&W˜'FñÊtñ◊ñÁFÚ“ÁV∆√∞¢VíÊñ◊ñÁFı&W˜'Df˜&“Á&W6WBÇì∞¢VíÊñ◊ñÁFı&W˜'DfVVF&6≤ÁFWáD6ˆÁFVÁB“"#∞¢VíÊñ◊ñÁFı&W˜'D÷ˆF¬Ê6∆74∆ó7BÊFBÇ&ÜñFFV‚"ì∞ß–†¶7ñÊ2gVÊ7Fñˆ‚7V&÷óDñ◊ñÁFı&W˜'BÜWfVÁBí∞¢WfVÁBÁ&WfVÁDFVfV«BÇì∞¢ñbÇ&W˜'FñÊtñ◊ñÁFÚí∞¢VíÊñ◊ñÁFı&W˜'DfVVF&6≤ÁFWáD6ˆÁFVÁB“$ñ◊ñÁFÚÊˆ‚Fó7ˆÊñ&ñ∆RW"∆6VvÊ∆¶ñˆÊR‚#∞¢&WGW&„∞¢–¢6ˆÁ7BW6W"“WFÇÊ7W'&VÁEW6W#∞¢ñbÇW6W"í∞¢VíÊñ◊ñÁFı&W˜'DfVVF&6≤ÁFWáD6ˆÁFVÁB“$FWfíf&R∆ˆvñ‚&ñ÷FíñÁfñ&RVÊ6VvÊ∆¶ñˆÊR‚#∞¢&WGW&„∞¢–¢6ˆÁ7BFóFˆ∆Ú“7G&ñÊráVíÊñ◊ñÁFı&W˜'EFóF∆RÁf«VR«¬""íÁG&ñ“Çì∞¢6ˆÁ7BFW7FÚ“7G&ñÊráVíÊñ◊ñÁFı&W˜'EFWáBÁf«VR«¬""íÁG&ñ“Çì∞¢ñbÇFóFˆ∆Ú«¬FW7FÚí∞¢VíÊñ◊ñÁFı&W˜'DfVVF&6≤ÁFWáD6ˆÁFVÁB“$6ˆ◊ñ∆FóFˆ∆ÚRFW7FÚFV∆∆6VvÊ∆¶ñˆÊR‚#∞¢&WGW&„∞¢–¢6ˆÁ7BÊ˜r“ÊWrFFRÇì∞¢6ˆÁ7B÷W76vR“∞¢.)™˚àÚ4Tt‰ƒ§îÙ‰R$Ù$ƒT‘î’îÂDÚ“&W˜'B˜W&FófÚ"¿¢	¯˘~˚àÚñ◊ñÁFÛ¢G∑&W˜'FñÊtñ◊ñÁFÚÊFVÊˆ÷ñÊ¶ñˆÊR«¬"“'÷¿¢	˘8“6ˆ◊VÊS¢G∑&W˜'FñÊtñ◊ñÁFÚÊ6ˆ◊VÊR«¬"“'÷¿¢	˘∫>˚àÚfñ¢G∑&W˜'FñÊtñ◊ñÁFÚÊñÊFó&óß¶Ú«¬"“'÷¿¢	¯iBîB4¢G∑&W˜'FñÊtñ◊ñÁFÚÊñE6«¬"“'÷¿¢	˘9“ˆvvWGFÚ6VvÊ∆¶ñˆÊS¢G∑FóFˆ∆˜÷¿¢	˘8≤FWGFv∆ñÚ&ˆ&∆V÷6VvÊ∆FÛ¢G∑FW7F˜÷¿¢	˘r˜W&F˜&R6VvÊ∆ÁFS¢G∑W6W"ÊFó7∆îÊ÷R«¬W6W"ÊV÷ñ¬«¬"“'÷¿¢	˘8RFF6VvÊ∆¶ñˆÊS¢G∂Ê˜rÁFÙ∆ˆ6∆TFFU7G&ñÊrÇ&óB‘ïB"ó÷¿¢	˘Y"˜&6VvÊ∆¶ñˆÊS¢G∂Ê˜rÁFÙ∆ˆ6∆UFñ÷U7G&ñÊrÇ&óB‘ïB"¬≤Ü˜W#¢#"÷FñvóB"¬÷ñÁWFS¢#"÷FñvóB"¬Ü˜W##¢f«6R“ó÷¿¢.)»R6ˆÊfW&÷¢7Fñ÷Ú6VvÊ∆ÊFÚ¬6∆ñVÁFRñ¬&ˆ&∆V÷&ó66ˆÁG&FÚRñ¬&V∆FófÚñÁFW'fVÁFÚ&ñ6ÜñW7FÚ‚ ¢“Ê¶ˆñ‚Ç%∆‚"ì∞¢6ˆÁ7B˜VÊVB“6fT˜VÂvÜG4÷W76vRÜ÷W76vRì∞¢VíÊñ◊ñÁFı&W˜'DfVVF&6≤ÁFWáD6ˆÁFVÁB“˜VÊV@¢Ú%vÜG4W'FÚ6ˆ‚∆6VvÊ∆¶ñˆÊR&ˆÁFFñÁfñ&R‚ ¢¢$ñ◊˜76ñ&ñ∆R&ó&RvÜG4WFˆ÷Fñ6÷VÁFR7RVW7FÚFó7˜6óFófÚ‚#∞¢6WEFñ÷V˜WBÜ6∆˜6Tñ◊ñÁFı&W˜'D÷ˆF¬¬#ì∞ß–†¶gVÊ7Fñˆ‚vWD7W'&VÁE˜6óFñˆ‰ˆÊ6RÇí∞¢&WGW&‚ÊWr&ˆ÷ó6RÇá&W6ˆ«fR¬&V¶V7Bí”‚∞¢ñbÇÊfñvF˜"ÊvVˆ∆ˆ6Fñˆ‚í∞¢&V¶V7BÜÊWrW'&˜"Ç$vVˆ∆ˆ6∆óß¶¶ñˆÊRÊˆ‚7W˜'FF‚"íì∞¢&WGW&„∞¢–¢ÊfñvF˜"ÊvVˆ∆ˆ6Fñˆ‚ÊvWD7W'&VÁE˜6óFñˆ‚Ä¢á˜2í”‚∞¢&W6ˆ«fRá∞¢∆C¢ÁV÷&W"á˜2Ê6ˆ˜&G2Ê∆FóGVFRí¿¢∆Ês¢ÁV÷&W"á˜2Ê6ˆ˜&G2Ê∆ˆÊvóGVFRí¿¢67W&7ì¢ÁV÷&W"á˜2Ê6ˆ˜&G2Ê67W&7í«¬í¿¢Fñ÷W7F◊¢ÁV÷&W"á˜2ÁFñ÷W7F◊«¬FFRÊÊ˜rÇíê¢“ì∞¢“¿¢ÜW'&˜"í”‚&V¶V7BÜW'&˜"í¿¢≤VÊ&∆TÜñvÑ67W&7ì¢G'VR¬Fñ÷V˜WC¢S¬÷Üñ◊V‘vS¢–¢ì∞¢“ì∞ß–†¶7ñÊ2gVÊ7Fñˆ‚&WVW7Dw5WFFRÜñ◊ñÁFÚí∞¢ñbÇ7W'&VÁEW6W"«¬6V∆V7FVD6ˆ÷÷W76ñBí∞¢∆W'BÇ%6V∆W¶ñˆÊVÊ6ˆ÷÷W76VBW6VwVíñ¬∆ˆvñ‚‚"ì∞¢&WGW&„∞¢–¢6ˆÁ7B6ˆÊfó&÷VB“vñÊF˜rÊ6ˆÊfó&“Ç%gVˆívvñ˜&Ê&R∆˜6ó¶ñˆÊRFíVW7FÚñ◊ñÁFÛÚfW',:ñÁfñF&ñ6ÜñW7FvÜG4∆¬v÷÷ñÊó7G&F˜&R‚"ì∞¢ñbÇ6ˆÊfó&÷VBí&WGW&„∞†¢G'í∞¢6ˆÁ7B˜2“vóBvWD7W'&VÁE˜6óFñˆ‰ˆÊ6RÇì∞¢6ˆÁ7Bñ◊ñÁFÙñG2“vWDñ◊ñÁFÙFˆ4ñG2Üñ◊ñÁFÚì∞¢6ˆÁ7B&WVW7E&Vb“vóBF"Ê6ˆ∆∆V7Fñˆ‚Ç&w5WFFU&WVW7G2"íÊFBá∞¢6ˆ÷÷W76ñC¢6V∆V7FVD6ˆ÷÷W76ñB¿¢6ˆ÷÷W76Ê÷S¢6V∆V7FVD6ˆ÷÷W76Ê÷R«¬""¿¢ñ◊ñÁFÙ∂Wì¢'Vñ∆Dñ◊ñÁFÙ∂WíÜñ◊ñÁFÚí¿¢ñ◊ñÁFÙñG2¿¢ñ◊ñÁFÙFVÊˆ÷ñÊ¶ñˆÊS¢ñ◊ñÁFÚÊFVÊˆ÷ñÊ¶ñˆÊR«¬""¿¢ñ◊ñÁFÙñE6¢ñ◊ñÁFÚÊñE6«¬""¿¢ñ◊ñÁFÙ6ˆ◊VÊS¢ñ◊ñÁFÚÊ6ˆ◊VÊR«¬""¿¢ñ◊ñÁFÙñÊFó&óß¶Û¢ñ◊ñÁFÚÊñÊFó&óß¶Ú«¬""¿¢˜W&F˜$ñC¢7W'&VÁEW6W"ÁVñB¿¢˜W&F˜$Ê÷S¢7W'&VÁEW6W"ÊFó7∆îÊ÷R«¬7W'&VÁEW6W"ÊV÷ñ¬«¬$˜W&F˜&R"¿¢˜W&F˜$V÷ñ√¢7W'&VÁEW6W"ÊV÷ñ¬«¬""¿¢˜W&F˜$∆C¢˜2Ê∆B¿¢˜W&F˜$∆Ês¢˜2Ê∆Êr¿¢7FGW3¢'VÊFñÊr"¿¢7&VFVDC¢fó&V&6RÊfó&W7F˜&R‰fñV∆Ef«VRÁ6W'fW%Fñ÷W7F◊Çê¢“ì∞†¢6ˆÁ7B÷5W&¬“áGG3¢Úˆ÷2Êvˆˆv∆RÊ6ˆ“Û˜“G∑˜2Ê∆G“¬G∑˜2Ê∆Êw÷∞¢6ˆÁ7BvFWáB“∞¢/	˘8“&ñ6ÜñW7Fvvñ˜&Ê÷VÁFÚu2ñ◊ñÁFÚ"¿¢îB&ñ6ÜñW7F¢G∑&WVW7E&VbÊñG÷¿¢6ˆ÷÷W76¢G∑6V∆V7FVD6ˆ÷÷W76Ê÷R«¬"“'÷¿¢ñ◊ñÁFÛ¢G∂ñ◊ñÁFÚÊFVÊˆ÷ñÊ¶ñˆÊR«¬"“'÷¿¢îB4¢G∂ñ◊ñÁFÚÊñE6«¬"“'÷¿¢˜W&F˜&S¢G∂7W'&VÁEW6W"ÊFó7∆îÊ÷R«¬7W'&VÁEW6W"ÊV÷ñ¬«¬"“'÷¿¢6ˆ˜&FñÊFR˜W&F˜&S¢G∑˜2Ê∆G“¬G∑˜2Ê∆Êw÷¿¢÷¢G∂÷5W&«÷ ¢“Ê¶ˆñ‚Ç%∆‚"ì∞¢6ˆÁ7B˜VÊVB“6fT˜VÂvÜG4÷W76vRávFWáB¬≤ÜˆÊS¢u5Ù$ıd≈ıÑÙ‰R“ì∞¢ñbÇ˜VÊVBí∆W'BÇ%&ñ6ÜñW7F7&VF¬÷Êˆ‚:Ç7FFÚ˜76ñ&ñ∆R&ó&RvÜG4WFˆ÷Fñ6÷VÁFR‚"ì∞†¢vóBÊ˜FñgîF÷ñÁ4f˜$w5&WVW7Bá&WVW7E&VbÊñB¬ñ◊ñÁFÚ¬˜2ì∞¢∆W'BÇ%&ñ6ÜñW7FñÁfñF‚ñ‚GFW6&˜f¶ñˆÊRF÷ñ‚‚"ì∞¢“6F6ÇÜW'&˜"í∞¢6ˆÁ6ˆ∆RÊW'&˜"Ç$W'&˜&R&ñ6ÜñW7Fvvñ˜&Ê÷VÁFÚu3¢"¬W'&˜"ì∞¢∆W'BÇ$ñ◊˜76ñ&ñ∆RñÁfñ&R∆&ñ6ÜñW7Fu2‚"ì∞¢–ß–†¶7ñÊ2gVÊ7Fñˆ‚Ê˜FñgîF÷ñÁ4f˜$w5&WVW7Bá&WVW7DñB¬ñ◊ñÁFÚ¬˜2í∞¢6ˆÁ7BF÷ñÂW6W'2“∆Ff˜&’W6W'2Êfñ«FW"ÇáW6W"í”‚F÷ñ‰V÷ñ«2ÊÜ2ÜÊ˜&÷∆ó¶TV÷ñ¬áW6W"ÊV÷ñ¬ííì∞¢ñbÇF÷ñÂW6W'2Ê∆VÊwFÇí&WGW&„∞¢6ˆÁ7BFWáB“	˘8“&ñ6ÜñW7Fu2G∑&WVW7DñG“W"G∂ñ◊ñÁFÚÊFVÊˆ÷ñÊ¶ñˆÊR«¬&ñ◊ñÁFÚ'“ÇG∑˜2Ê∆BÁFÙfóÜVBÉbó“¬G∑˜2Ê∆ÊrÁFÙfóÜVBÉbó“í‚&ívW7FñˆÊR‚WFVÁFíW"66WGF&R˜&ñfóWF&RÊ∞¢vóB&ˆ÷ó6RÊ∆¬ÜF÷ñÂW6W'2Ê÷ÇÜF÷ñÂW6W"í”‚F"Ê6ˆ∆∆V7Fñˆ‚Ç&6ÜD÷W76vW2"íÊFBá∞¢GóS¢'FWáB"¿¢FWáB¿¢&V6óñVÁDñC¢F÷ñÂW6W"ÊñB¿¢6VÊFW$ñC¢7W'&VÁEW6W"ÁVñB¿¢6VÊFW$Ê÷S¢7W'&VÁEW6W"ÊFó7∆îÊ÷R«¬7W'&VÁEW6W"ÊV÷ñ¬«¬$˜W&F˜&R"¿¢6VÊFW$V÷ñ√¢7W'&VÁEW6W"ÊV÷ñ¬«¬""¿¢7&VFVDC¢fó&V&6RÊfó&W7F˜&R‰fñV∆Ef«VRÁ6W'fW%Fñ÷W7F◊Çê¢“ííì∞ß–†¶7ñÊ2gVÊ7Fñˆ‚Ê˜FñgîF÷ñÁ4f˜$ñ◊ñÁFÙFˆÊU&V6˜fW'íÜñ◊ñÁFÚ¬&V6ˆ‚“""í∞¢6ˆÁ7BF÷ñÂW6W'2“∆Ff˜&’W6W'2Êfñ«FW"ÇáW6W"í”‚F÷ñ‰V÷ñ«2ÊÜ2ÜÊ˜&÷∆ó¶TV÷ñ¬áW6W"ÊV÷ñ¬ííì∞¢ñbÇF÷ñÂW6W'2Ê∆VÊwFÇí&WGW&„∞¢6ˆÁ7B6ˆ÷÷W76ñB“7G&ñÊrá6V∆V7FVD6ˆ÷÷W76ñB«¬""íÁG&ñ“Çì∞¢6ˆÁ7Bñ◊ñÁFÙñG2“vWDñ◊ñÁFÙFˆ4ñG2Üñ◊ñÁFÚíÊfñ«FW"Ñ&ˆˆ∆V‚ì∞¢ñbÇ6ˆ÷÷W76ñB«¬ñ◊ñÁFÙñG2Ê∆VÊwFÇí&WGW&„∞¢6ˆÁ7B˜W&F˜$Ê÷R“7W'&VÁEW6W#ÚÊFó7∆îÊ÷R«¬7W'&VÁEW6W#ÚÊV÷ñ¬«¬$˜W&F˜&R#∞¢6ˆÁ7BFWáB“∞¢.)™˚àÚ&V7WW&Ú7FFÚñ◊ñÁFÚ&ñ6ÜñW7FÚ"¿¢˜W&F˜&S¢G∂˜W&F˜$Ê÷W÷¿¢6ˆ÷÷W76¢G∑6V∆V7FVD6ˆ÷÷W76Ê÷R«¬$6ˆ÷÷W76'÷¿¢ñ◊ñÁFÛ¢G∂ñ◊ñÁFÚÊFVÊˆ÷ñÊ¶ñˆÊR«¬$ñ◊ñÁFÚ'÷¿¢&V6ˆ‚ÚFWGFv∆ñÛ¢G∑&V6ˆÁ÷¢$FWGFv∆ñÛ¢76vvñÚWFˆ÷Fñ6ÚídEDíÊˆ‚&óW66óFÚ‚"¿¢%&V÷íñ¬V«6ÁFRW"7˜7F&R¬vñ◊ñÁFÚÊVídEDí‚ ¢“Ê¶ˆñ‚Ç%∆‚"ì∞¢vóB&ˆ÷ó6RÊ∆¬ÜF÷ñÂW6W'2Ê÷ÇÜF÷ñÂW6W"í”‚F"Ê6ˆ∆∆V7Fñˆ‚Ç&6ÜD÷W76vW2"íÊFBá∞¢GóS¢'FWáB"¿¢FWáB¿¢&V6óñVÁDñC¢F÷ñÂW6W"ÊñB¿¢6VÊFW$ñC¢7W'&VÁEW6W#ÚÁVñB«¬""¿¢6VÊFW$Ê÷S¢˜W&F˜$Ê÷R¿¢6VÊFW$V÷ñ√¢7W'&VÁEW6W#ÚÊV÷ñ¬«¬""¿¢∂ñÊC¢'7ó7FV“"¿¢÷WFFF¢∞¢GóS¢&ñ◊ñÁFıˆFˆÊU˜&V6˜fW'í"¿¢7Fñˆ„¢&÷˜fUˆFˆÊR"¿¢6ˆ÷÷W76ñB¿¢6ˆ÷÷W76Ê÷S¢6V∆V7FVD6ˆ÷÷W76Ê÷R«¬$6ˆ÷÷W76"¿¢ñ◊ñÁFÙñG2¿¢ñ◊ñÁFÙÊ÷S¢ñ◊ñÁFÚÊFVÊˆ÷ñÊ¶ñˆÊR«¬$ñ◊ñÁFÚ"¿¢ñ◊ñÁFÙ∂Wì¢'Vñ∆Dñ◊ñÁFÙ∂WíÜñ◊ñÁFÚê¢“¿¢7&VFVDC¢fó&V&6RÊfó&W7F˜&R‰fñV∆Ef«VRÁ6W'fW%Fñ÷W7F◊Çê¢“ííì∞ß–†¶gVÊ7Fñˆ‚˜VÂ7VG&vÜG4á7VB¬6ˆ÷÷W76í∞¢6ˆÁ7B7VE&˜w2“'&íÊó4'&íá7VBÁ7VG&RíÚ7VBÁ7VG&R¢vWD∆Vv7ï7VG&U&˜w2á7VBì∞¢6ˆÁ7B&˜w4÷W76vR“7VE&˜w2Ê÷Çá&˜r¬ñGÇí”‚Ö∞¢	˘R5TE$G∂ñGÇ≤÷¿¢(
"W'6ˆÊ∆S¢G∑&˜rÁW'6ˆÊ∆R«¬"“'÷¿¢(
"÷Wß¶ì¢G∑&˜rÊ÷Wß¶í«¬"“'÷¿¢.)H)H)H)H)H)H)H)H)H)H)H)H)H)H)H)H)H)H)H)H ¢“Ê¶ˆñ‚Ç%∆‚"íííÊ¶ˆñ‚Ç%∆‚"ì∞¢6ˆÁ7B÷W76vR“∞¢/	˘:2&ñ6ÜñW7FFí6ˆÊfW&÷6ˆ◊˜6ó¶ñˆÊR7VG&R"¿¢$vVÁFñ∆RFV6Êñ6Ú¬Fí6VwVóFÚ∆6ˆ◊˜6ó¶ñˆÊR&Vvó7G&F‚"¿¢.)H)H)H)H)H)H)H)H)H)H)H)H)H)H)H)H)H)H)H)H"¿¢	˘86ˆ÷÷W76¢G∂6ˆ÷÷W76ÊÊˆ÷R«¬"“'÷¿¢	˘8Rvñ˜&ÊÚ&ñfW&ñ÷VÁFÛ¢G∑7VBÁ&ñfW&ñ÷VÁFÙFF«¬"“'÷¿¢.)H)H)H)H)H)H)H)H)H)H)H)H)H)H)H)H)H)H)H)H"¿¢&˜w4÷W76vR«¬$ÊW77VÊ7VG&6ˆ◊ñ∆FÂ∆Ó)H)H)H)H)H)H)H)H)H)H)H)H)H)H)H)H)H)H)H)H"¿¢$w&¶ñRW"∆fW&ñfñ6‚ ¢“Ê¶ˆñ‚Ç%∆‚"ì∞†¢ñbÇ6fT˜VÂvÜG4÷W76vRÜ÷W76vRíí∆W'BÇ$ñ◊˜76ñ&ñ∆R&ó&RvÜG47RVW7FÚFó7˜6óFófÚ‚"ì∞ß–†¶gVÊ7Fñˆ‚vWE7VG&U6∂vTVÁG&ñW2Çí∞¢6ˆÁ7B6V∆V7FVDFFT∂Wí“vWD7FófU7VG&TFFT∂WíÇì∞¢6ˆÁ7B7F˜&ñ6ÙFVƒvñ˜&ÊÚ“7VG&TÜó7F˜'î'îFFRÊvWBá6V∆V7FVDFFT∂Wíí«¬ÊWr÷Çì∞¢6ˆÁ7B6ˆ÷÷W76R“'&íÊg&ˆ“Ü6ˆ÷÷W76T'îñBÁf«VW2Çíì∞¢&WGW&‚6ˆ÷÷W76RÊ÷ÇÜ6ˆ÷÷W76í”‚∞¢6ˆÁ7B7VB“7F˜&ñ6ÙFVƒvñ˜&ÊÚÊvWBÜ6ˆ÷÷W76ÊñBí«¬∑”∞¢6ˆÁ7B7VE&˜w2“'&íÊó4'&íá7VBÁ7VG&RíÚ7VBÁ7VG&R¢vWD∆Vv7ï7VG&U&˜w2á7VBì∞¢6ˆÁ7BÜ5&˜w2“7VE&˜w2Á6ˆ÷RÇá&˜rí”‚&˜rÁW'6ˆÊ∆R«¬&˜rÊ÷Wß¶íì∞¢&WGW&‚∞¢6ˆ÷÷W76¿¢7VB¿¢7VE&˜w2¿¢Ü5&˜w0¢”∞¢“íÊfñ«FW"ÇÜVÁG'íí”‚VÁG'íÊÜ5&˜w2ì∞ß–†¶7ñÊ2gVÊ7Fñˆ‚'Vñ∆E7VG&U6∂vUFd&∆ˆ"ÜVÁG&ñW2í∞¢ñbÇvñÊF˜rÊß7Fb«¬vñÊF˜rÊß7FbÊß5DbíFá&˜rÊWrW'&˜"Ç$∆ñ'&W&ñDbÊˆ‚Fó7ˆÊñ&ñ∆R‚"ì∞¢6ˆÁ7B≤ß5Db““vñÊF˜rÊß7Fc∞¢6ˆÁ7BFˆ2“ÊWrß5Dbá≤˜&ñVÁFFñˆ„¢'˜'G&óB"¬VÊóC¢&÷“"¬f˜&÷C¢&B"“ì∞¢6ˆÁ7BvUvñGFÇ“#∞¢6ˆÁ7BvTÜVñváB“#ìs∞¢6ˆÁ7B÷&vñ‚“#∞¢6ˆÁ7B6ˆÁFVÁEvñGFÇ“vUvñGFÇ“Ü÷&vñ‚¢"ì∞¢6ˆÁ7B÷Öí“vTÜVñváB“÷&vñ„∞†¢6ˆÁ7BG&tÜVFW"“ÜVÁG'í¬ñGÇí”‚∞¢Fˆ2Á6WDfñ∆ƒ6ˆ∆˜"Éìí¬"¬#Cì∞¢Fˆ2Á&˜VÊFVE&V7BÜ÷&vñ‚¬÷&vñ‚¬6ˆÁFVÁEvñGFÇ¬#B¬B¬B¬$b"ì∞¢Fˆ2Á6WEFWáD6ˆ∆˜"É#SR¬#SR¬#SRì∞¢Fˆ2Á6WDfˆÁBÇ&ÜV«fWFñ6"¬&&ˆ∆B"ì∞¢Fˆ2Á6WDfˆÁE6ó¶RÉBì∞¢Fˆ2ÁFWáBÜ7VG&RW"6ˆ÷÷W76(
"G∂ñGÇ≤“ÚG∂VÁG&ñW2Ê∆VÊwFá÷¬÷&vñ‚≤b¬÷&vñ‚≤íì∞¢Fˆ2Á6WDfˆÁE6ó¶RÉì∞¢Fˆ2ÁFWáBÜVÁG'íÊ6ˆ÷÷W76ÊÊˆ÷R«¬$6ˆ÷÷W766VÁ¶Êˆ÷R"¬÷&vñ‚≤b¬÷&vñ‚≤bì∞¢Fˆ2Á6WDfˆÁBÇ&ÜV«fWFñ6"¬&Ê˜&÷¬"ì∞¢Fˆ2Á6WDfˆÁE6ó¶RÉí„Rì∞¢Fˆ2ÁFWáBÜvñ˜&ÊÛ¢G∂VÁG'íÁ7VBÁ&ñfW&ñ÷VÁFÙFF«¬"“'÷¬÷&vñ‚≤b¬÷&vñ‚≤#ì∞¢Fˆ2ÁFWáBÜWá˜'C¢G∂ÊWrFFRÇíÁFÙ∆ˆ6∆U7G&ñÊrÇ&óB‘ïB"ó÷¬vUvñGFÇ“÷&vñ‚“Cb¬÷&vñ‚≤#ì∞¢”∞†¢6ˆÁ7BG&u7VG&6&B“á&˜r¬&˜tñGÇ¬ï7F'Bí”‚∞¢∆WBí“ï7F'C∞¢Fˆ2Á6WDfñ∆ƒ6ˆ∆˜"É#S¬#S2¬#SRì∞¢Fˆ2Á6WDG&t6ˆ∆˜"É##b¬#3"¬#Cì∞¢Fˆ2Á&˜VÊFVE&V7BÜ÷&vñ‚¬í¬6ˆÁFVÁEvñGFÇ¬3B¬B¬B¬$dB"ì∞†¢Fˆ2Á6WDfñ∆ƒ6ˆ∆˜"É#3b¬#S2¬#C2ì∞¢Fˆ2Á6WDG&t6ˆ∆˜"ÉÉr¬#Cr¬#Çì∞¢Fˆ2Á&˜VÊFVE&V7BÜ÷&vñ‚≤B¬í≤B¬3¬r¬2¬2¬$dB"ì∞¢Fˆ2Á6WEFWáD6ˆ∆˜"É#"¬¬S"ì∞¢Fˆ2Á6WDfˆÁBÇ&ÜV«fWFñ6"¬&&ˆ∆B"ì∞¢Fˆ2Á6WDfˆÁE6ó¶RÉí„Rì∞¢Fˆ2ÁFWáBÜ7VG&G∑&˜tñGÇ≤÷¬÷&vñ‚≤r¬í≤Ç„Çì∞†¢6ˆÁ7BW'6ˆÊ∆T∆&V¬“/	˘RW'6ˆÊ∆S¢#∞¢6ˆÁ7B÷Wß¶î∆&V¬“/	˘©¢÷Wß¶ì¢#∞¢6ˆÁ7BW'6ˆÊÊVƒ∆ñÊW2“Fˆ2Á7∆óEFWáEFı6ó¶RÖ7G&ñÊrá&˜rÁW'6ˆÊ∆R«¬"“"í¬6ˆÁFVÁEvñGFÇ“CBì∞¢6ˆÁ7B÷Wß¶î∆ñÊW2“Fˆ2Á7∆óEFWáEFı6ó¶RÖ7G&ñÊrá&˜rÊ÷Wß¶í«¬"“"í¬6ˆÁFVÁEvñGFÇ“CBì∞†¢Fˆ2Á6WEFWáD6ˆ∆˜"Ér¬#B¬3íì∞¢Fˆ2Á6WDfˆÁBÇ&ÜV«fWFñ6"¬&&ˆ∆B"ì∞¢Fˆ2Á6WDfˆÁE6ó¶RÉí„Rì∞¢Fˆ2ÁFWáBáW'6ˆÊ∆T∆&V¬¬÷&vñ‚≤B¬í≤bì∞¢Fˆ2ÁFWáBÜ÷Wß¶î∆&V¬¬÷&vñ‚≤B¬í≤#Rì∞†¢Fˆ2Á6WDfˆÁBÇ&ÜV«fWFñ6"¬&Ê˜&÷¬"ì∞¢Fˆ2Á6WDfˆÁE6ó¶RÉí„2ì∞¢Fˆ2ÁFWáBáW'6ˆÊÊVƒ∆ñÊW2Á6∆ñ6RÉ¬"í¬÷&vñ‚≤3B¬í≤bì∞¢Fˆ2ÁFWáBÜ÷Wß¶î∆ñÊW2Á6∆ñ6RÉ¬"í¬÷&vñ‚≤3B¬í≤#Rì∞†¢6ˆÁ7B&˜w5W6VB“÷FÇÊ÷ÇáW'6ˆÊÊVƒ∆ñÊW2Ê∆VÊwFÇ¬÷Wß¶î∆ñÊW2Ê∆VÊwFÇ¬ì∞¢&WGW&‚í≤÷FÇÊ÷ÇÉ3B¬#B≤á&˜w5W6VB¢B„2íì∞¢”∞†¢VÁG&ñW2Êf˜$V6ÇÇÜVÁG'í¬ñGÇí”‚∞¢ñbÜñGÇ‚íFˆ2ÊFEvRÇì∞¢G&tÜVFW"ÜVÁG'í¬ñGÇì∞¢∆WBí“÷&vñ‚≤3∞†¢ñbÇVÁG'íÁ7VE&˜w2Ê∆VÊwFÇí∞¢Fˆ2Á6WEFWáD6ˆ∆˜"ÉsR¬ÉR¬ìíì∞¢Fˆ2Á6WDfˆÁE6ó¶RÉì∞¢Fˆ2ÁFWáBÇ$ÊW77VÊ7VG&6ˆ◊ñ∆FW"VW7F6ˆ÷÷W76‚"¬÷&vñ‚¬í≤Çì∞¢&WGW&„∞¢–†¢VÁG'íÁ7VE&˜w2Êf˜$V6ÇÇá&˜r¬&˜tñGÇí”‚∞¢ñbáí‚÷Öí“Cí∞¢Fˆ2ÊFEvRÇì∞¢G&tÜVFW"ÜVÁG'í¬ñGÇì∞¢í“÷&vñ‚≤3∞¢–¢í“G&u7VG&6&Bá&˜r¬&˜tñGÇ¬í≤"í≤C∞¢“ì∞¢“ì∞¢&WGW&‚Fˆ2Ê˜WGWBÇ&&∆ˆ""ì∞ß–†¶7ñÊ2gVÊ7Fñˆ‚6Ü&T∆≈7VG&UFıvÜG4Çí∞¢6ˆÁ7BVÁG&ñW2“vWE7VG&U6∂vTVÁG&ñW2Çì∞¢ñbÇVÁG&ñW2Ê∆VÊwFÇí∞¢∆W'BÇ$ÊW77VÊ6ˆ◊˜6ó¶ñˆÊR7VG&RFó7ˆÊñ&ñ∆RFñÁfñ&R‚"ì∞¢&WGW&„∞¢–†¢6ˆÁ7B6˜'FVDVÁG&ñW2“≤‚‚ÊVÁG&ñW5“Á6˜'BÇÜ¬"í”‚7G&ñÊrÜÊ6ˆ÷÷W76ÊÊˆ÷R«¬""íÊ∆ˆ6∆T6ˆ◊&RÖ7G&ñÊrÜ"Ê6ˆ÷÷W76ÊÊˆ÷R«¬""í¬&óB"íì∞¢6ˆÁ7Bw&˜WVD∆ñÊW2“6˜'FVDVÁG&ñW2Ê÷ÇÜVÁG'í¬VÁG'îñGÇí”‚∞¢6ˆÁ7BFFT∆&V¬“VÁG'íÁ7VBÁ&ñfW&ñ÷VÁFÙFF¢ÚÊWrFFRÜG∂VÁG'íÁ7VBÁ&ñfW&ñ÷VÁFÙFF’C££íÁFÙ∆ˆ6∆TFFU7G&ñÊrÇ&óB‘ïB"ê¢¢"“#∞¢6ˆÁ7B7VD∆ñÊW2“VÁG'íÁ7VE&˜w2Ê÷Çá&˜r¬&˜tñGÇí”‚Ö∞¢	˘R5TE$G∑&˜tñGÇ≤÷¿¢(
"W'6ˆÊ∆S¢G∑&˜rÁW'6ˆÊ∆R«¬"“'÷¿¢(
"÷Wß¶ì¢G∑&˜rÊ÷Wß¶í«¬"“'÷¿¢")H)H)H)H)H)H)H)H)H)H)H)H)H)H)H)H ¢“Ê¶ˆñ‚Ç%∆‚"íííÊ¶ˆñ‚Ç%∆‚"ì∞¢&WGW&‚∞¢.)Y)Y)Y)Y)Y)Y)Y)Y)Y)Y)Y)Y)Y)Y)Y)Y)Y)Y)Y)Y"¿¢	˘84Ù‘‘U54G∂VÁG'îñGÇ≤”¢Gµ7G&ñÊrÜVÁG'íÊ6ˆ÷÷W76ÊÊˆ÷R«¬$6ˆ÷÷W76"íÁFıWW$66RÇó÷¿¢	˘8Rvñ˜&ÊÚ&ˆw&÷÷FÛ¢G∂FFT∆&V«÷¿¢7VD∆ñÊW2«¬"“ÊW77VÊ7VG&76VvÊF“ ¢“Ê¶ˆñ‚Ç%∆‚"ì∞¢“ì∞†¢6ˆÁ7B÷W76vR“∞¢/	˘:2$ıı5D5TE$RıU$DïdR"¿¢$'VˆÊvñ˜&ÊÚ¬6ˆÊFófñFÚ∆&˜˜7F7VG&RW"∆ñÊñfñ6¶ñˆÊR˜W&Fóf‚"¿¢""¿¢w&˜WVD∆ñÊW2Ê¶ˆñ‚Ç%∆Â∆‚"í¿¢""¿¢.)»RW"ff˜&R6ˆÊfW&÷&RÚ6VvÊ∆&RWfVÁGV∆í÷ˆFñfñ6ÜR‚ ¢“Ê¶ˆñ‚Ç%∆‚"ì∞¢ñbÇ6fT˜VÂvÜG4÷W76vRÜ÷W76vRíí∆W'BÇ$ñ◊˜76ñ&ñ∆R&ó&RvÜG47RVW7FÚFó7˜6óFófÚ‚"ì∞ß–†¶gVÊ7Fñˆ‚vWD6ˆ÷÷W7666VÁD6ˆ∆˜"Ü6ˆ÷÷W76ñB¬ñÊFWÇí∞¢6ˆÁ7B∆WGFR“≤"3#Sc6V""¬"3v36VB"¬"3cscfR"¬"6Cìssb"¬"6F##ssr"¬"3Éì#""¬"3FcCfSR"¬"66ÜB%”∞¢6ˆÁ7B6˜W&6R“7G&ñÊrÜ6ˆ÷÷W76ñB«¬ñÊFWÇ«¬""ì∞¢∆WBÜ6Ç“∞¢f˜"Ü∆WBí“≤í¬6˜W&6RÊ∆VÊwFÉ≤í≥“í∞¢Ü6Ç“ÇÜÜ6Ç√¬Rí“Ü6Çí≤6˜W&6RÊ6Ü$6ˆFTBÜíì∞¢Ü6Ç√“∞¢–¢&WGW&‚∆WGFU¥÷FÇÊ'2ÜÜ6ÇíR∆WGFRÊ∆VÊwFÖ”∞ß–†¶gVÊ7Fñˆ‚vWEFñ÷W7F◊FFRáf«VRí∞¢ñbÇf«VRí&WGW&‚ÁV∆√∞¢ñbáGóVˆbf«VRÁFÙFFR””“&gVÊ7Fñˆ‚"í&WGW&‚f«VRÁFÙFFRÇì∞¢ñbáGóVˆbf«VRÁ6V6ˆÊG2””“&ÁV÷&W""í&WGW&‚ÊWrFFRáf«VRÁ6V6ˆÊG2¢ì∞¢6ˆÁ7B'6VB“ÊWrFFRáf«VRì∞¢&WGW&‚ÁV÷&W"Êó4Ê‚á'6VBÊvWEFñ÷RÇííÚÁV∆¬¢'6VC∞ß–†¶gVÊ7Fñˆ‚FD˜W&F˜%˜6óFñˆ‰÷&∂W%FÙ∆ñW"á˜6óFñˆ‚¬∆ñW"í∞¢&WGW&‚¬Ê÷&∂W"Ö∑˜6óFñˆ‚Ê∆B¬˜6óFñˆ‚Ê∆Êu“¬∞¢ñ6ˆ„¢¬ÊFódñ6ˆ‚á∞¢6∆74Ê÷S¢""¿¢áF÷√¢#∆Fób6∆73“v÷&∂W"÷˜W&F˜"r&ñ÷ÜñFFV„“wG'VRsÔ	˙k£¬ˆFóc‚"¿¢ñ6ˆÂ6ó¶S¢≥b¬e“¿¢ñ6ˆ‰Ê6Ü˜#¢≥Ç¬Ö–¢“ê¢“íÊFEFÚÜ∆ñW"ì∞ß–†¶gVÊ7Fñˆ‚vWD˜W&F˜%˜6óFñˆÁ4f˜$÷Çí∞¢ñbÇ7W'&VÁEW6W%˜2«¬7W'&VÁEW6W"í&WGW&‚µ”∞¢6ˆÁ7B'îñB“ÊWr÷Çì∞¢ñbÜ7W'&VÁEW6W"bb7W'&VÁEW6W%˜2í∞¢6ˆÁ7B7W'&VÁD76ñvÊ÷VÁB“vWD7W'&VÁD˜W&F˜%˜6óFñˆ‰76ñvÊ÷VÁBÇì∞¢'îñBÁ6WBÜ7W'&VÁEW6W"ÁVñB¬∞¢‚‚‚Ü'îñBÊvWBÜ7W'&VÁEW6W"ÁVñBí«¬∑“í¿¢ñC¢7W'&VÁEW6W"ÁVñB¿¢VñC¢7W'&VÁEW6W"ÁVñB¿¢V÷ñ√¢7W'&VÁEW6W"ÊV÷ñ¬«¬""¿¢Fó7∆îÊ÷S¢7W'&VÁEW6W"ÊFó7∆îÊ÷R«¬7W'&VÁEW6W"ÊV÷ñ¬«¬%WFVÁFR"¿¢˜W&F˜$Ê÷S¢7W'&VÁD76ñvÊ÷VÁBÊ˜W&F˜$Ê÷R«¬7W'&VÁEW6W"ÊFó7∆îÊ÷R«¬7W'&VÁEW6W"ÊV÷ñ¬«¬%WFVÁFR"¿¢‚‚Ê7W'&VÁD76ñvÊ÷VÁB¿¢∆C¢7W'&VÁEW6W%˜2Ê∆B¿¢∆Ês¢7W'&VÁEW6W%˜2Ê∆Êr¿¢67W&7ì¢7W'&VÁEW6W%˜2Ê67W&7í«¬¿¢WFFVDC¢ÊWrFFRÇê¢“ì∞¢–¢&WGW&‚'&íÊg&ˆ“Ü'îñBÁf«VW2ÇííÁ6∆ñ6RÉ¬íÊfñ«FW"Çá˜6óFñˆ‚í”‚∞¢6ˆÁ7B∆B“ÁV÷&W"á˜6óFñˆ‚Ê∆Bì∞¢6ˆÁ7B∆Êr“ÁV÷&W"á˜6óFñˆ‚Ê∆Êrì∞¢ñbÇÁV÷&W"Êó4fñÊóFRÜ∆Bí«¬ÁV÷&W"Êó4fñÊóFRÜ∆Êríí&WGW&‚f«6S∞¢˜6óFñˆ‚Ê∆B“∆C∞¢˜6óFñˆ‚Ê∆Êr“∆Ês∞¢˜6óFñˆ‚Ê67W&7í“ÁV÷&W"á˜6óFñˆ‚Ê67W&7í«¬ì∞¢&WGW&‚G'VS∞¢“ì∞ß–†¶gVÊ7Fñˆ‚&VÊFW$˜W&F˜%˜6óFñˆ‰÷&∂W'2Ü&˜VÊG2í∞¢vWD˜W&F˜%˜6óFñˆÁ4f˜$÷ÇíÊf˜$V6ÇÇá˜6óFñˆ‚í”‚∞¢FD˜W&F˜%˜6óFñˆ‰÷&∂W%FÙ∆ñW"á˜6óFñˆ‚¬÷&∂W$∆ñW"ì∞¢FD˜W&F˜%˜6óFñˆ‰÷&∂W%FÙ∆ñW"á˜6óFñˆ‚¬gV∆«67&VV‰÷&∂W$∆ñW"ì∞¢&˜VÊG2ÁW6ÇÖ∑˜6óFñˆ‚Ê∆B¬˜6óFñˆ‚Ê∆Êu“ì∞¢“ì∞ß–†¶gVÊ7Fñˆ‚Fˆvv∆T˜W&F˜%˜6óFñˆÁ5fó6ñ&ñ∆óGíÇí∞¢&WGW&„∞ß–††¶gVÊ7Fñˆ‚'Vñ∆D÷÷&∂W%6WVVÊ6RÜñ◊ñÁFí“µ“í∞¢&WGW&‚ñ◊ñÁFê¢Ê÷ÇÜñ◊ñÁFÚí”‚á≤ñ◊ñÁFÚ¬∂Wì¢'Vñ∆Dñ◊ñÁFÙ∂WíÜñ◊ñÁFÚí¬∆C¢ÁV÷&W"Üñ◊ñÁFÚÊw5íí¬∆Ês¢ÁV÷&W"Üñ◊ñÁFÚÊw5Çí“íê¢Êfñ«FW"Çá&˜rí”‚&˜rÊ∂WíbbÁV÷&W"Êó4fñÊóFRá&˜rÊ∆BíbbÁV÷&W"Êó4fñÊóFRá&˜rÊ∆Êríê¢Á6˜'BÇÜ¬"í”‚∞¢ñbÑ÷FÇÊ'2ÜÊ∆B“"Ê∆Bí‚„í&WGW&‚"Ê∆B“Ê∆C∞¢ñbÑ÷FÇÊ'2ÜÊ∆Êr“"Ê∆Êrí‚„í&WGW&‚Ê∆Êr“"Ê∆Ês∞¢&WGW&‚7G&ñÊrÜÊ∂WííÊ∆ˆ6∆T6ˆ◊&RÖ7G&ñÊrÜ"Ê∂Wíí¬&óB"ì∞¢“ê¢Á&VGV6RÇÜ62¬&˜r¬ñÊFWÇí”‚62Á6WBá&˜rÊ∂Wí¬ñÊFWÇ≤í¬ÊWr÷Çíì∞ß–†¶gVÊ7Fñˆ‚vWD÷÷&∂W$ÁV÷&W$f˜$ñ◊ñÁFÚÜñ◊ñÁFÚí∞¢ñbÇñ◊ñÁFÚí&WGW&‚ÁV∆√∞¢&WGW&‚÷÷&∂W%6WVVÊ6T'î∂WíÊvWBÜ'Vñ∆Dñ◊ñÁFÙ∂WíÜñ◊ñÁFÚíí«¬ÁV∆√∞ß–†¶gVÊ7Fñˆ‚'Vñ∆Dñ◊ñÁFÙ÷&∂W$&FvRÜñ◊ñÁFÚí∞¢6ˆÁ7B÷&∂W$6∆72“vWD÷&∂W$6∆72Üñ◊ñÁFÚì∞¢6ˆÁ7B÷&∂W$ÁV÷&W"“vWD÷÷&∂W$ÁV÷&W$f˜$ñ◊ñÁFÚÜñ◊ñÁFÚì∞¢ñbÇÁV÷&W"Êó4fñÊóFRÜ÷&∂W$ÁV÷&W"íí&WGW&‚"#∞¢&WGW&‚«7‚6∆73“&÷&∂W"◊ñ‚÷&FvR"&ñ÷ÜñFFV„“'G'VR#„«7‚6∆73“&÷&∂W"◊ñ‚G∂÷&∂W$6∆77“#„«7‚6∆73“&÷&∂W"◊ñ‚÷ÁV÷&W"#‚G∂÷&∂W$ÁV÷&W'”¬˜7„„¬˜7„„¬˜7„Ê∞ß–†¶gVÊ7Fñˆ‚'6Tñ◊ñÁFÙ÷6ˆ˜&FñÊFRáf«VR¬÷ñ‚¬÷Çí∞¢ñbáf«VR”“ÁV∆¬«¬7G&ñÊráf«VRíÁG&ñ“Çí””“""í&WGW&‚ÁV∆√∞¢6ˆÁ7B6ˆ˜&FñÊFR“ÁV÷&W"Ö7G&ñÊráf«VRíÁG&ñ“ÇíÁ&W∆6RÇ"¬"¬"‚"íì∞¢ñbÇÁV÷&W"Êó4fñÊóFRÜ6ˆ˜&FñÊFRí«¬6ˆ˜&FñÊFR¬÷ñ‚«¬6ˆ˜&FñÊFR‚÷Ç«¬6ˆ˜&FñÊFR””“í&WGW&‚ÁV∆√∞¢&WGW&‚6ˆ˜&FñÊFS∞ß–†¶gVÊ7Fñˆ‚vWDñ◊ñÁFÙ÷6ˆ˜&FñÊFW2Üñ◊ñÁFÚí∞¢6ˆÁ7B∆B“'6Tñ◊ñÁFÙ÷6ˆ˜&FñÊFRÜñ◊ñÁFÛÚÊw5í¬”ì¬ìì∞¢6ˆÁ7B∆Êr“'6Tñ◊ñÁFÙ÷6ˆ˜&FñÊFRÜñ◊ñÁFÛÚÊw5Ç¬”É¬Éì∞¢&WGW&‚∆B”“ÁV∆¬«¬∆Êr”“ÁV∆¬ÚÁV∆¬¢∂∆B¬∆Êu”∞ß–†¶gVÊ7Fñˆ‚fˆ7W4ñ◊ñÁFÙ'î÷ÁV÷&W"á&tÁV÷&W"¬F&vWD÷“÷í∞¢6ˆÁ7B÷&∂W$ÁV÷&W"“ÁV÷&W"Á'6TñÁBÖ7G&ñÊrá&tÁV÷&W"«¬""íÁG&ñ“Çí¬ì∞¢ñbÇÁV÷&W"Êó4fñÊóFRÜ÷&∂W$ÁV÷&W"í«¬÷&∂W$ÁV÷&W"¬í∞¢∆W'BÇ$ÁV÷W&Úñ◊ñÁFÚÊˆ‚G&˜fFÚ"ì∞¢&WGW&„∞¢–¢6ˆÁ7B÷F6Ç“7W'&VÁDñ◊ñÁFíÊfñÊBÇÜñ◊ñÁFÚí”‚vWD÷÷&∂W$ÁV÷&W$f˜$ñ◊ñÁFÚÜñ◊ñÁFÚí””“÷&∂W$ÁV÷&W"ì∞¢6ˆÁ7B6ˆ˜&FñÊFW2“vWDñ◊ñÁFÙ÷6ˆ˜&FñÊFW2Ü÷F6Çì∞¢ñbÇ÷F6Ç«¬6ˆ˜&FñÊFW2í∞¢∆W'BÇ$ÁV÷W&Úñ◊ñÁFÚÊˆ‚G&˜fFÚ"ì∞¢&WGW&„∞¢–¢6ˆÁ7B∂Wí“'Vñ∆Dñ◊ñÁFÙ∂WíÜ÷F6Çì∞¢6ˆÁ7B÷&∂W$÷“F&vWD÷””“gV∆«67&VV‰÷ÚgV∆«67&VV‰ñ◊ñÁFÙ÷&∂W$'î∂Wí¢ñ◊ñÁFÙ÷&∂W$'î∂Wì∞¢6ˆÁ7B÷&∂W"“÷&∂W$÷ÊvWBÜ∂Wíì∞¢F&vWD÷Á6WEfñWrÜ6ˆ˜&FñÊFW2¬÷FÇÊ÷ÇáF&vWD÷ÊvWE¶ˆˆ“Çí¬Rí¬≤Êñ÷FS¢G'VR“ì∞¢ñbÜ÷&∂W#ÚÊ˜VÂ˜Wí÷&∂W"Ê˜VÂ˜WÇì∞¢6V∆V7Dñ◊ñÁFÙf˜$÷FWFñ¬Ü÷F6Çì∞ß–†¶gVÊ7Fñˆ‚&VÊFW$÷Çí∞¢6∆V$÷Çì∞†¢6ˆÁ7B&˜VÊG2“µ”∞¢6ˆÁ7Bñ◊ñÁFî&˜VÊG2“µ”∞¢÷÷&∂W%6WVVÊ6T'î∂Wí“'Vñ∆D÷÷&∂W%6WVVÊ6RÜ7W'&VÁDñ◊ñÁFíì∞¢6ˆÁ7B÷FF6ñvÊGW&R“G∑6V∆V7FVD6ˆ÷÷W76ñG”£¢G∂7W'&VÁDñ◊ñÁFê¢Ê÷ÇÜñ◊ñÁFÚí”‚∞¢6ˆÁ7B6ˆ˜&FñÊFW2“vWDñ◊ñÁFÙ÷6ˆ˜&FñÊFW2Üñ◊ñÁFÚì∞¢&WGW&‚G∂'Vñ∆Dñ◊ñÁFÙ∂WíÜñ◊ñÁFÚó◊¬G∂6ˆ˜&FñÊFW3ÚÂ≥“ÛÚ"'◊¬G∂6ˆ˜&FñÊFW3ÚÂ≥“ÛÚ"'÷∞¢“ê¢Á6˜'BÇê¢Ê¶ˆñ‚Ç#≤"ó÷∞¢∆WB÷&∂W$f˜$7FófTgV∆«67&VVÂ˜W“ÁV∆√∞†¢7W'&VÁDñ◊ñÁFíÊf˜$V6ÇÇÜñ◊ñÁFÚí”‚∞¢6ˆÁ7Bñ◊ñÁFÙ∂Wí“'Vñ∆Dñ◊ñÁFÙ∂WíÜñ◊ñÁFÚì∞¢6ˆÁ7B6ˆ˜&FñÊFW2“vWDñ◊ñÁFÙ÷6ˆ˜&FñÊFW2Üñ◊ñÁFÚì∞¢ñbÜ6ˆ˜&FñÊFW2íñ◊ñÁFî&˜VÊG2ÁW6ÇÜ6ˆ˜&FñÊFW2ì∞¢6ˆÁ7B6Ê˜u&ˆEFÇ“vWE6Ê˜u&ˆEFÇÜñ◊ñÁFÚì∞¢FE6Ê˜u&ˆEˆ«ñ∆ñÊUFÙ∆ñW"Üñ◊ñÁFÚ¬6Ê˜u&ˆD∆ñW"¬÷ì∞¢FE6Ê˜u&ˆEˆ«ñ∆ñÊUFÙ∆ñW"Üñ◊ñÁFÚ¬gV∆«67&VVÂ6Ê˜u&ˆD∆ñW"¬gV∆«67&VV‰÷ì∞¢6Ê˜u&ˆEFÇÊf˜$V6ÇÇáˆñÁBí”‚&˜VÊG2ÁW6ÇáˆñÁBíì∞¢6ˆÁ7B÷&∂W"“FDñ◊ñÁFÙ÷&∂W%FÙ÷∆ñW"Üñ◊ñÁFÚ¬÷&∂W$∆ñW"¬÷ì∞¢ñbÜ÷&∂W"íñ◊ñÁFÙ÷&∂W$'î∂WíÁ6WBÜñ◊ñÁFÙ∂Wí¬÷&∂W"ì∞¢6ˆÁ7BgV∆«67&VV‰÷&∂W"“FDñ◊ñÁFÙ÷&∂W%FÙ÷∆ñW"Üñ◊ñÁFÚ¬gV∆«67&VV‰÷&∂W$∆ñW"¬gV∆«67&VV‰÷ì∞¢ñbÜgV∆«67&VV‰÷&∂W"í∞¢ñbÜñ◊ñÁFÙ∂WíígV∆«67&VV‰ñ◊ñÁFÙ÷&∂W$'î∂WíÁ6WBÜñ◊ñÁFÙ∂Wí¬gV∆«67&VV‰÷&∂W"ì∞¢&˜VÊG2ÁW6ÇÜ6ˆ˜&FñÊFW2ì∞¢ñbÜñ◊ñÁFÙ∂Wíbbñ◊ñÁFÙ∂Wí””“6V∆V7FVDgV∆«67&VV‰ñ◊ñÁFÙñBí÷&∂W$f˜$7FófTgV∆«67&VVÂ˜W“gV∆«67&VV‰÷&∂W#∞¢–¢“ì∞†¢&VÊFW$˜W&F˜%˜6óFñˆ‰÷&∂W'2Ü&˜VÊG2ì∞†¢6ˆÁ7B6Ü˜V∆DWFÙfóEFÙ&˜VÊG2“ñ◊ñÁFî&˜VÊG2Ê∆VÊwFÇ‚bbÇ÷ñ‰÷fñWu7FFRÊÜ5W6W$÷˜fVB«¬÷WFÙfóE6ñvÊGW&R”“÷FF6ñvÊGW&Rì∞¢ñbá6Ü˜V∆DWFÙfóEFÙ&˜VÊG2í∞¢÷ÊfóD&˜VÊG2Üñ◊ñÁFî&˜VÊG2¬≤FFñÊs¢≥#B¬#E“¬÷Ö¶ˆˆ”¢¬Êñ÷FS¢f«6R“ì∞¢6ˆÁ7B6VÁFW"“÷ÊvWD6VÁFW"Çì∞¢÷ñ‰÷fñWu7FFRÊ6VÁFW"“∂6VÁFW"Ê∆B¬6VÁFW"Ê∆Êu”∞¢÷ñ‰÷fñWu7FFRÁ¶ˆˆ““÷ÊvWE¶ˆˆ“Çì∞¢÷ñ‰÷fñWu7FFRÊÜ5W6W$÷˜fVB“G'VS∞¢÷WFÙfóE6ñvÊGW&R“÷FF6ñvÊGW&S∞¢“V«6R∞¢÷Á6WEfñWrÜ÷ñ‰÷fñWu7FFRÊ6VÁFW"¬÷ñ‰÷fñWu7FFRÁ¶ˆˆ“¬≤Êñ÷FS¢f«6R“ì∞¢–¢gV∆«67&VV‰÷Á6WEfñWrÜ÷ñ‰÷fñWu7FFRÊ6VÁFW"¬÷ñ‰÷fñWu7FFRÁ¶ˆˆ“¬≤Êñ÷FS¢f«6R“ì∞¢7ñÊ56V∆V7FVDñ◊ñÁFÙFWFñƒgFW%&Vg&W6ÇÜ÷&∂W$f˜$7FófTgV∆«67&VVÂ˜Wì∞¢&V∆ˆD6ˆ÷÷W76vVFÜW$f˜%fó6ñ&∆Tñ◊ñÁFíÇì∞ß–†¶gVÊ7Fñˆ‚7ñÊ56V∆V7FVDñ◊ñÁFÙFWFñƒgFW%&Vg&W6ÇÜ÷&∂W$f˜%6V∆V7FVDñ◊ñÁFÚí∞¢ñbÇ6V∆V7FVDñ◊ñÁFÙñBí&WGW&„∞¢6ˆÁ7B∆FW7Dñ◊ñÁFÚ“fñÊD7W'&VÁDñ◊ñÁFÙ'î∂Wíá6V∆V7FVDñ◊ñÁFÙñBì∞¢ñbÇ∆FW7Dñ◊ñÁFÚí∞¢6∆˜6U6V∆V7FVDñ◊ñÁFÙFWFñ¬á≤6∆˜6U˜W¢G'VR“ì∞¢&WGW&„∞¢–¢6V∆V7FVDñ◊ñÁFÙFF“≤‚‚Ê∆FW7Dñ◊ñÁFÚ”∞¢6V∆V7FVDgV∆«67&VV‰ñ◊ñÁFÙñB“6V∆V7FVDñ◊ñÁFÙñC∞¢&VÊFW%6V∆V7FVDñ◊ñÁFÙFWFñ≈ÊV¬Çì∞¢∂VW6V∆V7FVDgV∆«67&VVÂ˜W˜V‚Ü÷&∂W$f˜%6V∆V7FVDñ◊ñÁFÚì∞ß–†¶gVÊ7Fñˆ‚∂VW6V∆V7FVDgV∆«67&VVÂ˜W˜V‚Ü÷&∂W$f˜%6V∆V7FVDñ◊ñÁFÚí∞¢ñbÇ6V∆V7FVDgV∆«67&VV‰ñ◊ñÁFÙñB«¬÷&∂W$f˜%6V∆V7FVDñ◊ñÁFÚí&WGW&„∞¢6ˆÁ7B&V˜VÂ˜W“Çí”‚∞¢ñbÇ6V∆V7FVDgV∆«67&VV‰ñ◊ñÁFÙñB«¬gV∆«67&VV‰÷&∂W$∆ñW"ÊÜ4∆ñW"Ü÷&∂W$f˜%6V∆V7FVDñ◊ñÁFÚí«¬÷&∂W$f˜%6V∆V7FVDñ◊ñÁFÚÊvWE˜WÚ‚Çíí&WGW&„∞¢÷&∂W$f˜%6V∆V7FVDñ◊ñÁFÚÊ˜VÂ˜WÇì∞¢”∞¢&WVW7DÊñ÷Fñˆ‰g&÷Rá&V˜VÂ˜Wì∞¢6WEFñ÷V˜WBá&V˜VÂ˜W¬Éì∞ß–††¶gVÊ7Fñˆ‚vWE6Ê˜u&ˆEFÇÜñ◊ñÁFÚí∞¢6ˆÁ7B&r“'&íÊó4'&íÜñ◊ñÁFÛÚÁ&˜WFUFÇíÚñ◊ñÁFÚÁ&˜WFUFÇ¢µ”∞¢&WGW&‚&rÊ÷ÇáˆñÁBí”‚'&íÊó4'&íáˆñÁBê¢Ú¥ÁV÷&W"áˆñÁE≥“í¬ÁV÷&W"áˆñÁE≥“ï–¢¢¥ÁV÷&W"áˆñÁCÚÊ∆BÛÚˆñÁCÚÊ∆FóGVFRí¬ÁV÷&W"áˆñÁCÚÊ∆ÊrÛÚˆñÁCÚÊ∆ˆ‚ÛÚˆñÁCÚÊ∆ˆÊvóGVFRï–¢íÊfñ«FW"ÇÖ∂∆B¬∆Êu“í”‚ÁV÷&W"Êó4fñÊóFRÜ∆BíbbÁV÷&W"Êó4fñÊóFRÜ∆Êríì∞ß–†¶gVÊ7Fñˆ‚FE6Ê˜u&ˆEˆ«ñ∆ñÊUFÙ∆ñW"Üñ◊ñÁFÚ¬F&vWD∆ñW"¬F&vWD÷í∞¢6ˆÁ7BFÇ“vWE6Ê˜u&ˆEFÇÜñ◊ñÁFÚì∞¢ñbÇñ◊ñÁFÛÚÁ6Ê˜u&ˆB«¬FÇÊ∆VÊwFÇ¬"í&WGW&‚ÁV∆√∞¢6ˆÁ7B6ˆ∆˜"“ñ◊ñÁFÚÊFˆÊRÚ"3f3F"¢"3#Sc6V"#∞¢6ˆÁ7B∆ñÊR“¬Áˆ«ñ∆ñÊRáFÇ¬≤6ˆ∆˜"¬vVñváC¢b¬˜6óGì¢„í¬∆ñÊT6¢'&˜VÊB"¬∆ñÊT¶ˆñ„¢'&˜VÊB"“ì∞¢ñbáF&vWD÷”“gV∆«67&VV‰÷í∆ñÊRÊ&ñÊE˜WÜ'Vñ∆Dñ◊ñÁFÙ÷˜WÜñ◊ñÁFÚ¬ñ◊ñÁFÚÁFóÙ÷ÁWFVÁ¶ñˆÊR«¬%fñÊWfR"íì∞¢∆ñÊRÊˆ‚Ç&6∆ñ6≤"¬Çí”‚6V∆V7Dñ◊ñÁFÙf˜$÷FWFñ¬Üñ◊ñÁFÚíì∞¢∆ñÊRÊFEFÚáF&vWD∆ñW"ì∞¢&WGW&‚∆ñÊS∞ß–†¶gVÊ7Fñˆ‚Fó7FÊ6T÷WFW'5Fı6Ê˜u&ˆBÜñ◊ñÁFÚ¬˜6óFñˆ‚í∞¢6ˆÁ7BFÇ“vWE6Ê˜u&ˆEFÇÜñ◊ñÁFÚì∞¢ñbÇ˜6óFñˆ‚«¬FÇÊ∆VÊwFÇ¬"í&WGW&‚ÁV÷&W"Âı4ïDïdUÙî‰dî‰ïEì∞¢&WGW&‚FÇÁ&VGV6RÇÜ&W7B¬ˆñÁBí”‚÷FÇÊ÷ñ‚Ü&W7B¬ÜfW'6ñÊRá˜6óFñˆ‚Ê∆B¬˜6óFñˆ‚Ê∆Êr¬ˆñÁE≥“¬ˆñÁE≥“í¢í¬ÁV÷&W"Âı4ïDïdUÙî‰dî‰ïEíì∞ß–†¶7ñÊ2gVÊ7Fñˆ‚WFÙ6ˆ◊∆WFU76VE6Ê˜u&ˆG2Çí∞¢ñbÇó56Ê˜u6W'fñ6T6ˆÁFWáBÇí«¬6V∆V7FVD6ˆ÷÷W76ñB«¬7W'&VÁEW6W%˜2í&WGW&„∞¢6ˆÁ7B76VB“7W'&VÁDñ◊ñÁFíÊfñ«FW"ÇÜñ◊ñÁFÚí”‚ñ◊ñÁFÛÚÁ6Ê˜u&ˆBbbñ◊ñÁFÚÊFˆÊRbbFó7FÊ6T÷WFW'5Fı6Ê˜u&ˆBÜñ◊ñÁFÚ¬7W'&VÁEW6W%˜2í√“#Rì∞¢ñbÇ76VBÊ∆VÊwFÇí&WGW&„∞¢vóB6WDñ◊ñÁFÙFˆÊRá6V∆V7FVD6ˆ÷÷W76ñB¬76VBÊ÷ÇÜñ◊ñÁFÚí”‚ñ◊ñÁFÚÊñBíÊfñ«FW"Ñ&ˆˆ∆V‚í¬G'VR¬≤FˆÊT'ì¢7W'&VÁEW6W#ÚÊFó7∆îÊ÷R«¬7W'&VÁEW6W#ÚÊV÷ñ¬«¬$˜W&F˜&RÊWfR"“ì∞ß–†¶gVÊ7Fñˆ‚FDñ◊ñÁFÙ÷&∂W%FÙ÷∆ñW"Üñ◊ñÁFÚ¬F&vWD∆ñW"¬F&vWD÷“÷í∞¢6ˆÁ7B6ˆ˜&FñÊFW2“vWDñ◊ñÁFÙ÷6ˆ˜&FñÊFW2Üñ◊ñÁFÚì∞¢ñbÇ6ˆ˜&FñÊFW2í&WGW&‚ÁV∆√∞†¢6ˆÁ7B÷&∂W$6∆72“vWD÷&∂W$6∆72Üñ◊ñÁFÚì∞¢6ˆÁ7B÷&∂W%6WVVÊ6R“÷÷&∂W%6WVVÊ6T'î∂WíÊvWBÜ'Vñ∆Dñ◊ñÁFÙ∂WíÜñ◊ñÁFÚíì∞¢6ˆÁ7B÷&∂W"“¬Ê÷&∂W"Ö∂ñ◊ñÁFÚÊw5í¬ñ◊ñÁFÚÊw5Ö“¬∞¢ñ6ˆ„¢¬ÊFódñ6ˆ‚á∞¢6∆74Ê÷S¢""¿¢áF÷√¢∆Fób6∆73“&÷&∂W"◊ñ‚G∂÷&∂W$6∆77“#‚G¥ÁV÷&W"Êó4fñÊóFRÜ÷&∂W%6WVVÊ6RíÚ«7‚6∆73“&÷&∂W"◊ñ‚÷ÁV÷&W"#‚G∂÷&∂W%6WVVÊ6W”¬˜7„Ê¢"'”¬ˆFócÊ¿¢ñ6ˆÂ6ó¶S¢≥Ç¬Ö“¿¢ñ6ˆ‰Ê6Ü˜#¢≥í¬ï–¢“ê¢“ì∞¢ñbáF&vWD÷”“gV∆«67&VV‰÷í∞¢6ˆÁ7BFóÚ“ñ◊ñÁFÚÁFóÙ÷ÁWFVÁ¶ñˆÊR«¬6∆76ñgïFóÙ÷ÁWFVÁ¶ñˆÊRÜñ◊ñÁFÚÊ6ˆFñ6U&Wß¶Úì∞¢÷&∂W"Ê&ñÊE˜WÜ'Vñ∆Dñ◊ñÁFÙ÷˜WÜñ◊ñÁFÚ¬FóÚí¬∞¢WFÙ6∆˜6S¢f«6R¿¢6∆˜6Tˆ‰6∆ñ6≥¢f«6R¿¢6∆˜6T'WGFˆ„¢f«6R¿¢∂VWñÂfñWs¢G'VR¿¢÷ÖvñGFÉ¢3C¿¢÷ñÂvñGFÉ¢##¿¢6∆74Ê÷S¢&ñ◊ñÁFÚ÷÷◊˜W ¢“ì∞¢–¢÷&∂W"Êˆ‚Ç&6∆ñ6≤"¬Çí”‚∞¢6ˆÁ7BÊWáDñ◊ñÁFÙñB“'Vñ∆Dñ◊ñÁFÙ∂WíÜñ◊ñÁFÚì∞¢ñbáF&vWD÷””“gV∆«67&VV‰÷bb6V∆V7FVDñ◊ñÁFÙñBbb6V∆V7FVDñ◊ñÁFÙñB”“ÊWáDñ◊ñÁFÙñBígV∆«67&VV‰÷Ê6∆˜6U˜WÇì∞¢6V∆V7Dñ◊ñÁFÙf˜$÷FWFñ¬Üñ◊ñÁFÚì∞¢fˆ7W4ñ◊ñÁFÙñ‰∆ó7BÜñ◊ñÁFÚ¬f«6Rì∞¢“ì∞¢÷&∂W"ÊFEFÚáF&vWD∆ñW"ì∞¢&WGW&‚÷&∂W#∞ß–†¶gVÊ7Fñˆ‚VÁ7W&UW6W$∆ˆ6Fñˆ‰÷&∂W"áF&vWD÷¬6ˆ˜&G2í∞¢6ˆÁ7B÷&∂W%&Vb“F&vWD÷””“gV∆«67&VV‰÷Ú&gV∆«67&VV‚"¢&÷ñ‚#∞¢6ˆÁ7BWÜó7FñÊr“÷&∂W%&Vb””“&gV∆«67&VV‚"ÚW6W$∆ˆ6Fñˆ‰gV∆«67&VV‰÷&∂W"¢W6W$∆ˆ6Fñˆ‰÷÷&∂W#∞¢ñbÜWÜó7FñÊríWÜó7FñÊrÁ6WD∆D∆ÊrÜ6ˆ˜&G2ì∞¢V«6R∞¢6ˆÁ7B÷&∂W"“¬Ê÷&∂W"Ü6ˆ˜&G2¬∞¢ñ6ˆ„¢¬ÊFódñ6ˆ‚á∞¢6∆74Ê÷S¢'W6W"÷∆ˆ6Fñˆ‚÷÷&∂W"◊w&"¿¢áF÷√¢#«7‚6∆73“wW6W"÷∆ˆ6Fñˆ‚÷÷&∂W"÷F˜Br&ñ÷ÜñFFV„“wG'VRs„¬˜7„‚"¿¢ñ6ˆÂ6ó¶S¢≥b¬e“¿¢ñ6ˆ‰Ê6Ü˜#¢≥Ç¬Ö–¢“ê¢“íÊFEFÚáF&vWD÷ì∞¢ñbÜ÷&∂W%&Vb””“&gV∆«67&VV‚"íW6W$∆ˆ6Fñˆ‰gV∆«67&VV‰÷&∂W"“÷&∂W#∞¢V«6RW6W$∆ˆ6Fñˆ‰÷÷&∂W"“÷&∂W#∞¢–ß–†¶gVÊ7Fñˆ‚6VÁFW$÷ˆÂW6W$∆ˆ6Fñˆ‚áF&vWD÷“÷í∞¢ñbÇÊfñvF˜"ÊvVˆ∆ˆ6Fñˆ‚«¬ó4∆ˆ6Fñˆ‰VÊ&∆VBí∞¢WFFT∆ˆ6FñˆÂv&ÊñÊu7FFRá≤f∆∆&6¥÷W76vS¢ÊfñvF˜"ÊvVˆ∆ˆ6Fñˆ‚Ú&Ê˜E˜7W˜'FVB"¢""“ì∞¢&WGW&„∞¢–¢ÊfñvF˜"ÊvVˆ∆ˆ6Fñˆ‚ÊvWD7W'&VÁE˜6óFñˆ‚Ä¢á7V66W72í”‚∞¢6ˆÁ7B∆B“7V66W72Ê6ˆ˜&G2Ê∆FóGVFS∞¢6ˆÁ7B∆Êr“7V66W72Ê6ˆ˜&G2Ê∆ˆÊvóGVFS∞¢F&vWD÷Á6WEfñWrÖ∂∆B¬∆Êu“¬bì∞¢VÁ7W&UW6W$∆ˆ6Fñˆ‰÷&∂W"áF&vWD÷¬∂∆B¬∆Êu“ì∞¢ñbáF&vWD÷””“÷íVÁ7W&UW6W$∆ˆ6Fñˆ‰÷&∂W"ÜgV∆«67&VV‰÷¬∂∆B¬∆Êu“ì∞¢ñbáF&vWD÷””“gV∆«67&VV‰÷íVÁ7W&UW6W$∆ˆ6Fñˆ‰÷&∂W"Ü÷¬∂∆B¬∆Êu“ì∞¢“¿¢Çí”‚∞¢WFFT∆ˆ6FñˆÂv&ÊñÊu7FFRá≤f∆∆&6¥÷W76vS¢&÷ÁV≈˜6WGFñÊw2"“ì∞¢“¿¢∞¢VÊ&∆TÜñvÑ67W&7ì¢G'VR¿¢Fñ÷V˜WC¢¿¢÷Üñ◊V‘vS¢3 ¢–¢ì∞ß–†¶gVÊ7Fñˆ‚&VÊFW$∆ˆ6FñˆÂv&ÊñÊrÇí∞¢6Ü˜t∆ˆ6FñˆÂv&ÊñÊr“ó4∆ˆ6Fñˆ‰VÊ&∆VB«¬∆ˆ6FñˆÂW&÷ó76ñˆ‚”“&w&ÁFVB#∞¢VíÊ÷∆ˆ6FñˆÂv&ÊñÊsÚÊ6∆74∆ó7BÁFˆvv∆RÇ&ÜñFFV‚"¬6Ü˜t∆ˆ6FñˆÂv&ÊñÊrì∞¢ñbÇ6Ü˜t∆ˆ6FñˆÂv&ÊñÊrí&WGW&„∞¢6ˆÁ7B&∆ˆ6∂VB“∆ˆ6FñˆÂv&ÊñÊt÷ˆFR””“&&∆ˆ6∂VB#∞¢ñbáVíÊ÷∆ˆ6FñˆÂv&ÊñÊuFóF∆RíVíÊ÷∆ˆ6FñˆÂv&ÊñÊuFóF∆RÁFWáD6ˆÁFVÁB“&∆ˆ6∂VBÚ/	˘8“˜6ó¶ñˆÊR&∆ˆ66F"¢%˜6ó¶ñˆÊRÊˆ‚GFóf#∞¢ñbáVíÊ÷∆ˆ6FñˆÂv&ÊñÊuFWáBí∞¢VíÊ÷∆ˆ6FñˆÂv&ÊñÊuFWáBÁFWáD6ˆÁFVÁB“&∆ˆ6∂V@¢Ú$∆˜6ó¶ñˆÊR:Ç&∆ˆ66FW"VW7FÚ'&˜w6W"‚W"W6&R∆˜6ó¶ñˆÊR¬FWfí&ñ∆óF&∆÷ÁV∆÷VÁFR‚ ¢¢%W"W6&RVW7FgVÁ¶ñˆÊR&ñ∆óF∆˜6ó¶ñˆÊRF¬FV∆VfˆÊÚ‚#∞¢–¢VíÊ÷VÊ&∆T∆ˆ6Fñˆ‰'F„ÚÊ6∆74∆ó7BÁFˆvv∆RÇ&ÜñFFV‚"¬&∆ˆ6∂VBì∞¢VíÊ÷&WG'î∆ˆ6Fñˆ‰'F„ÚÊ6∆74∆ó7BÁFˆvv∆RÇ&ÜñFFV‚"¬&∆ˆ6∂VBì∞¢VíÊ÷∆ˆ6Fñˆ‰ÜV«'F„ÚÊ6∆74∆ó7BÁFˆvv∆RÇ&ÜñFFV‚"¬&∆ˆ6∂VBì∞¢ñbáVíÊ÷∆ˆ6FñˆÂv&ÊñÊu∆Ff˜&“í∞¢VíÊ÷∆ˆ6FñˆÂv&ÊñÊu∆Ff˜&“Ê6∆74∆ó7BÁFˆvv∆RÇ&ÜñFFV‚"¬&∆ˆ6∂VBì∞¢VíÊ÷∆ˆ6FñˆÂv&ÊñÊu∆Ff˜&“ÁFWáD6ˆÁFVÁB“'&˜w6W"&ñ∆WfFÛ¢G∂∆ˆ6Fñˆ‰6∆ñVÁDñÊfÚÊ'&˜w6W'“(
"6ó7FV÷&ñ∆WfFÛ¢G∂∆ˆ6Fñˆ‰6∆ñVÁDñÊfÚÊ˜7÷∞¢–ß–†¶gVÊ7Fñˆ‚WFFT∆ˆ6FñˆÂv&ÊñÊu7FFRá≤W&÷ó76ñˆ‚¬VÊ&∆VB¬÷ˆFR““∑“í∞¢ñbáW&÷ó76ñˆ‚í∆ˆ6FñˆÂW&÷ó76ñˆ‚“W&÷ó76ñˆ„∞¢ñbáGóVˆbVÊ&∆VB””“&&ˆˆ∆V‚"íó4∆ˆ6Fñˆ‰VÊ&∆VB“VÊ&∆VC∞¢ñbÜ÷ˆFRí∆ˆ6FñˆÂv&ÊñÊt÷ˆFR“÷ˆFS∞¢ñbÜ∆ˆ6FñˆÂv&ÊñÊt÷ˆFR”“&&∆ˆ6∂VB"íVíÊ÷∆ˆ6Fñˆ‰ÜV«ÊV√ÚÊ6∆74∆ó7BÊFBÇ&ÜñFFV‚"ì∞¢&VÊFW$∆ˆ6FñˆÂv&ÊñÊrÇì∞ß–†¶gVÊ7Fñˆ‚FWFV7D∆ˆ6Fñˆ‰6∆ñVÁDñÊfÚÇí∞¢6ˆÁ7BV“7G&ñÊrÜÊfñvF˜"ÁW6W$vVÁB«¬""íÁFÙ∆˜vW$66RÇì∞¢6ˆÁ7BV'&ÊG2“'&íÊó4'&íÜÊfñvF˜"ÁW6W$vVÁDFFÚÊ'&ÊG2íÚÊfñvF˜"ÁW6W$vVÁDFFÊ'&ÊG2Ê÷ÇÜ"í”‚7G&ñÊrÜ"Ê'&ÊB«¬""íÁFÙ∆˜vW$66RÇíí¢µ”∞¢6ˆÁ7BÜ4'&ÊB“ÜÊ÷Rí”‚V'&ÊG2Á6ˆ÷RÇÜ'&ÊBí”‚'&ÊBÊñÊ6«VFW2ÜÊ÷Ríì∞¢6ˆÁ7B'&˜w6W"“VÊñÊ6«VFW2Ç'6◊7VÊv'&˜w6W""í«¬Ü4'&ÊBÇ'6◊7VÊr"íÚ%6◊7VÊrñÁFW&ÊWB ¢¢áVÊñÊ6«VFW2Ç&VFrÚ"í«¬Ü4'&ÊBÇ&VFvR"ííÚ$VFvR ¢¢áVÊñÊ6«VFW2Ç&fó&Vf˜Ç"í«¬Ü4'&ÊBÇ&fó&Vf˜Ç"ííÚ$fó&Vf˜Ç ¢¢áVÊñÊ6«VFW2Ç&6á&ˆ÷R"í«¬Ü4'&ÊBÇ&6á&ˆ÷R"ííÚ$6á&ˆ÷R ¢¢áVÊñÊ6«VFW2Ç'6f&í"í«¬Ü4'&ÊBÇ'6f&í"ííÚ%6f&í ¢¢$«G&Ú'&˜w6W"#∞¢6ˆÁ7B˜2“VÊñÊ6«VFW2Ç&ÊG&ˆñB"íÚ$ÊG&ˆñB ¢¢áVÊñÊ6«VFW2Ç&óÜˆÊR"í«¬VÊñÊ6«VFW2Ç&óB"í«¬VÊñÊ6«VFW2Ç&óˆB"ííÚ&ïÜˆÊRˆîı2 ¢¢VÊñÊ6«VFW2Ç'vñÊF˜w2"íÚ%vñÊF˜w2 ¢¢$«G&Ú6ó7FV÷#∞¢∆ˆ6Fñˆ‰6∆ñVÁDñÊfÚ“≤'&˜w6W"¬˜2”∞ß–†¶gVÊ7Fñˆ‚vWD∆ˆ6Fñˆ‰ÜV«7FW2Çí∞¢6ˆÁ7B∂Wí“G∂∆ˆ6Fñˆ‰6∆ñVÁDñÊfÚÊ˜7◊¬G∂∆ˆ6Fñˆ‰6∆ñVÁDñÊfÚÊ'&˜w6W'÷∞¢6ˆÁ7B÷2“∞¢$ÊG&ˆñGƒ6á&ˆ÷R#¢≤%Fˆ66ñ¬«V66ÜWGFÚfñ6ñÊÚ∆Œ(	ññÊFó&óß¶ÚFV¬6óFÚ"¬%Fˆ66W&÷W76í"¬$&í˜6ó¶ñˆÊR"¬%6V∆W¶ñˆÊ6ˆÁ6VÁFí"¬%&ñ6&ñ6∆vñÊ"¬%F˜&ÊÊV∆Œ(	ñR&V÷í(	≈&ó&˜f˜6ó¶ñˆÊ^(	“%“¿¢$ÊG&ˆñG≈6◊7VÊrñÁFW&ÊWB#¢≤$&íñ◊˜7F¶ñˆÊíFV¬FV∆VfˆÊÚ"¬%fí7R"¬$6W&66◊7VÊrñÁFW&ÊWB"¬$VÁG&ñ‚W&÷W76í"¬%Fˆ66˜6ó¶ñˆÊR"¬%6V∆W¶ñˆÊ6ˆÁ6VÁFí"¬%F˜&ÊÊV∆Œ(	ñR&V÷í(	≈&ó&˜f˜6ó¶ñˆÊ^(	“%“¿¢$ÊG&ˆñGƒVFvR#¢≤$&íñ◊˜7F¶ñˆÊíFV¬FV∆VfˆÊÚ"¬%fí7R"¬$6W&6÷ñ7&˜6ˆgBVFvR"¬$VÁG&ñ‚W&÷W76í"¬%Fˆ66˜6ó¶ñˆÊR"¬%6V∆W¶ñˆÊ6ˆÁ6VÁFí"¬%F˜&ÊÊV∆Œ(	ñR&V÷í(	≈&ó&˜f˜6ó¶ñˆÊ^(	“%“¿¢&ïÜˆÊRˆîı7≈6f&í#¢≤$&íñ◊˜7F¶ñˆÊíïÜˆÊR"¬%fí7R&óf7íR6ñ7W&Wß¶"¬%Fˆ66∆ˆ6∆óß¶¶ñˆÊR"¬$76ñ7W&Fí6ÜR∆ˆ6∆óß¶¶ñˆÊR6ñGFóf"¬$6W&66f&í"¬$ñ◊˜7F∆˜6ó¶ñˆÊR7R(	ƒ÷VÁG&RW6íŒ(	ñ(	“"¬%F˜&ÊÊV∆Œ(	ñR&V÷í(	≈&ó&˜f˜6ó¶ñˆÊ^(	“%“¿¢&ïÜˆÊRˆîı7ƒ6á&ˆ÷R#¢≤$&íñ◊˜7F¶ñˆÊíïÜˆÊR"¬%fí7R&óf7íR6ñ7W&Wß¶"¬%Fˆ66∆ˆ6∆óß¶¶ñˆÊR"¬$6W&66á&ˆ÷R"¬$ñ◊˜7F∆˜6ó¶ñˆÊR7R(	ƒ÷VÁG&RW6íŒ(	ñ(	“"¬%F˜&ÊÊV∆Œ(	ñR&V÷í(	≈&ó&˜f˜6ó¶ñˆÊ^(	“%–¢”∞¢&WGW&‚÷5∂∂Wï“«¬≤$&í∆Rñ◊˜7F¶ñˆÊíFV¬FV∆VfˆÊÚ"¬%fí7R"¬$6W&6ñ¬'&˜w6W"6ÜR7FíW6ÊFÚ"¬$VÁG&ñ‚W&÷W76í"¬$GFóf˜6ó¶ñˆÊR"¬%F˜&ÊÊV∆Œ(	ñ"¬%&V÷í(	≈&ó&˜f˜6ó¶ñˆÊ^(	“%”∞ß–†¶gVÊ7Fñˆ‚Fˆvv∆T∆ˆ6Fñˆ‰ÜV«ÊV¬Çí∞¢ñbÇVíÊ÷∆ˆ6Fñˆ‰ÜV«ÊV¬í&WGW&„∞¢6ˆÁ7BÊ˜tÜñFFV‚“VíÊ÷∆ˆ6Fñˆ‰ÜV«ÊV¬Ê6∆74∆ó7BÁFˆvv∆RÇ&ÜñFFV‚"ì∞¢ñbÇÊ˜tÜñFFV‚í∞¢6ˆÁ7BóFV◊2“vWD∆ˆ6Fñˆ‰ÜV«7FW2ÇíÊ÷Çá7FWí”‚∆∆ì‚G∂W66TÖD‘¬á7FWó”¬ˆ∆ìÊíÊ¶ˆñ‚Ç""ì∞¢VíÊ÷∆ˆ6Fñˆ‰ÜV«ÊV¬ÊñÊÊW$ÖD‘¬“∆ˆ√‚G∂óFV◊7”¬ˆˆ√Ê∞¢–ß–†¶7ñÊ2gVÊ7Fñˆ‚7ñÊ4∆ˆ6Fñˆ‰fñ∆&ñ∆óGíÇí∞¢G'í∞¢FWFV7D∆ˆ6Fñˆ‰6∆ñVÁDñÊfÚÇì∞¢ñbÇÊfñvF˜"ÊvVˆ∆ˆ6Fñˆ‚í∞¢6∆V$7W'&VÁEW6W%˜6óFñˆ‚Çì∞¢WFFT∆ˆ6FñˆÂv&ÊñÊu7FFRá≤W&÷ó76ñˆ„¢'VÊfñ∆&∆R"¬VÊ&∆VC¢f«6R¬÷ˆFS¢&&∆ˆ6∂VB"“ì∞¢&WGW&„∞¢–¢ñbÜÊfñvF˜"ÁW&÷ó76ñˆÁ3ÚÁVW'íí∞¢6ˆÁ7B7FGW2“vóBÊfñvF˜"ÁW&÷ó76ñˆÁ2ÁVW'íá≤Ê÷S¢&vVˆ∆ˆ6Fñˆ‚"“ì∞¢ñbá7FGW2Á7FFR””“&FVÊñVB"í∞¢6∆V$7W'&VÁEW6W%˜6óFñˆ‚Çì∞¢WFFT∆ˆ6FñˆÂv&ÊñÊu7FFRá≤W&÷ó76ñˆ„¢&FVÊñVB"¬VÊ&∆VC¢f«6R¬÷ˆFS¢&&∆ˆ6∂VB"“ì∞¢–¢V«6Rñbá7FGW2Á7FFR””“'&ˆ◊B"íWFFT∆ˆ6FñˆÂv&ÊñÊu7FFRá≤W&÷ó76ñˆ„¢'&ˆ◊B"¬VÊ&∆VC¢f«6R¬÷ˆFS¢'&ˆ◊B"“ì∞¢V«6RWFFT∆ˆ6FñˆÂv&ÊñÊu7FFRá≤W&÷ó76ñˆ„¢&w&ÁFVB"¬÷ˆFS¢'&ˆ◊B"“ì∞¢7FGW2ÊˆÊ6ÜÊvR“Çí”‚≤fˆñB7ñÊ4∆ˆ6Fñˆ‰fñ∆&ñ∆óGíÇì≤”∞¢–¢ñbÇÊfñvF˜"ÁW&÷ó76ñˆÁ3ÚÁVW'íí∞¢ÊfñvF˜"ÊvVˆ∆ˆ6Fñˆ‚ÊvWD7W'&VÁE˜6óFñˆ‚ÇÇí”‚∞¢WFFT∆ˆ6FñˆÂv&ÊñÊu7FFRá≤W&÷ó76ñˆ„¢&w&ÁFVB"¬VÊ&∆VC¢G'VR¬÷ˆFS¢'&ˆ◊B"“ì∞¢“¬ÜW'&˜"í”‚∞¢6∆V$7W'&VÁEW6W%˜6óFñˆ‚Çì∞¢6ˆÁ7BFVÊñVB“W'&˜#ÚÊ6ˆFR””“∞¢WFFT∆ˆ6FñˆÂv&ÊñÊu7FFRá≤W&÷ó76ñˆ„¢FVÊñVBÚ&FVÊñVB"¢'&ˆ◊B"¬VÊ&∆VC¢f«6R¬÷ˆFS¢FVÊñVBÚ&&∆ˆ6∂VB"¢'&ˆ◊B"“ì∞¢“¬≤VÊ&∆TÜñvÑ67W&7ì¢G'VR¬Fñ÷V˜WC¢¬÷Üñ◊V‘vS¢“ì∞¢–¢“6F6ÇÜW'&˜"í∞¢6ˆÁ6ˆ∆RÁv&‚Ç%fW&ñfñ6Fó7ˆÊñ&ñ∆óL:˜6ó¶ñˆÊRÊˆ‚&óW66óF¢"¬W'&˜"ì∞¢6∆V$7W'&VÁEW6W%˜6óFñˆ‚Çì∞¢WFFT∆ˆ6FñˆÂv&ÊñÊu7FFRá≤W&÷ó76ñˆ„¢'VÊfñ∆&∆R"¬VÊ&∆VC¢f«6R¬÷ˆFS¢&&∆ˆ6∂VB"“ì∞¢–ß–†¶7ñÊ2gVÊ7Fñˆ‚&WVW7D∆ˆ6Fñˆ‰VÊ&∆Tf∆˜rÜ˜FñˆÁ2“∑“í∞¢G'í∞¢ñbÇÊfñvF˜"ÊvVˆ∆ˆ6Fñˆ‚í∞¢WFFT∆ˆ6FñˆÂv&ÊñÊu7FFRá≤W&÷ó76ñˆ„¢'VÊfñ∆&∆R"¬VÊ&∆VC¢f«6R¬÷ˆFS¢&&∆ˆ6∂VB"“ì∞¢&WGW&„∞¢–¢6ˆÁ7Bó466óF˜"“&ˆˆ∆V‚ávñÊF˜r‰66óF˜#ÚÊó4ÊFófU∆Ff˜&”Ú‚Çíì∞¢ñbÜó466óF˜"bbvñÊF˜r‰66óF˜#ÚÂ«VvñÁ3Ú‰vVˆ∆ˆ6Fñˆ„ÚÁ&WVW7EW&÷ó76ñˆÁ2í∞¢vóBvñÊF˜r‰66óF˜"Â«VvñÁ2‰vVˆ∆ˆ6Fñˆ‚Á&WVW7EW&÷ó76ñˆÁ2Çì∞¢–¢ÊfñvF˜"ÊvVˆ∆ˆ6Fñˆ‚ÊvWD7W'&VÁE˜6óFñˆ‚Çá˜2í”‚∞¢WFFT7W'&VÁEW6W%˜6óFñˆ‚á˜2Ê6ˆ˜&G2¬˜2ÁFñ÷W7F◊¬≤&VÊFW#¢f«6R“ì∞¢VÁ7W&UW6W$∆ˆ6Fñˆ‰÷&∂W"Ü÷¬∑˜2Ê6ˆ˜&G2Ê∆FóGVFR¬˜2Ê6ˆ˜&G2Ê∆ˆÊvóGVFU“ì∞¢VÁ7W&UW6W$∆ˆ6Fñˆ‰÷&∂W"ÜgV∆«67&VV‰÷¬∑˜2Ê6ˆ˜&G2Ê∆FóGVFR¬˜2Ê6ˆ˜&G2Ê∆ˆÊvóGVFU“ì∞¢WFFT∆ˆ6FñˆÂv&ÊñÊu7FFRá≤W&÷ó76ñˆ„¢&w&ÁFVB"¬VÊ&∆VC¢G'VR¬÷ˆFS¢'&ˆ◊B"“ì∞¢÷Á6WEfñWrÖ∑˜2Ê6ˆ˜&G2Ê∆FóGVFR¬˜2Ê6ˆ˜&G2Ê∆ˆÊvóGVFU“¬bì∞¢gV∆«67&VV‰÷Á6WEfñWrÖ∑˜2Ê6ˆ˜&G2Ê∆FóGVFR¬˜2Ê6ˆ˜&G2Ê∆ˆÊvóGVFU“¬bì∞¢“¬7ñÊ2ÜW'&˜"í”‚∞¢6∆V$7W'&VÁEW6W%˜6óFñˆ‚Çì∞¢6ˆÁ7BFVÊñVB“W'&˜#ÚÊ6ˆFR””“∞¢6ˆÁ7B&∆ˆ6∂VD∆ñ∂R“FVÊñVB«¬W'&˜#ÚÊ6ˆFR””“"«¬W'&˜#ÚÊ6ˆFR””“2«¬˜FñˆÁ2Êf˜&6U&WG'ì∞¢WFFT∆ˆ6FñˆÂv&ÊñÊu7FFRá≤VÊ&∆VC¢f«6R¬W&÷ó76ñˆ„¢FVÊñVBÚ&FVÊñVB"¢∆ˆ6FñˆÂW&÷ó76ñˆ‚¬÷ˆFS¢&∆ˆ6∂VD∆ñ∂RÚ&&∆ˆ6∂VB"¢'&ˆ◊B"“ì∞¢ñbÜW'&˜#ÚÊ6ˆFR””“í∞¢ñbÜó466óF˜"bbvñÊF˜r‰66óF˜#ÚÂ«VvñÁ3Ú‰ÚÊ˜VÂ6WGFñÊw2í∞¢G'í∞¢vóBvñÊF˜r‰66óF˜"Â«VvñÁ2‰Ê˜VÂ6WGFñÊw2Çì∞¢“6F6Ç∞¢WFFT∆ˆ6FñˆÂv&ÊñÊu7FFRá≤÷ˆFS¢&&∆ˆ6∂VB"“ì∞¢–¢–¢–¢“¬≤VÊ&∆TÜñvÑ67W&7ì¢G'VR¬Fñ÷V˜WC¢¬÷Üñ◊V‘vS¢“ì∞¢“6F6ÇÜW'&˜"í∞¢6ˆÁ6ˆ∆RÁv&‚Ç%&ñ6ÜñW7FW&÷W76Ú˜6ó¶ñˆÊRÊˆ‚&óW66óF¢"¬W'&˜"ì∞¢WFFT∆ˆ6FñˆÂv&ÊñÊu7FFRá≤÷ˆFS¢&&∆ˆ6∂VB"“ì∞¢–ß–†¶gVÊ7Fñˆ‚vWDñ◊ñÁFı˜WFFÜñ◊ñÁFÚ¬FóÚ“""í∞¢6ˆÁ7BFˆÊTñÊfÚ“f˜&÷DFˆÊTFFUFñ÷RÜñ◊ñÁFÚÊFˆÊTBì∞¢6ˆÁ7BñE6“ñ◊ñÁFÚÊñE6«¬ñ◊ñÁFÚÊ6ˆFñ6U6«¬"“#∞¢6ˆÁ7Bfñ“ñ◊ñÁFÚÊñÊFó&óß¶Ú«¬ñ◊ñÁFÚÊFW67&ó¶ñˆÊUfñ«¬ñ◊ñÁFÚÁfñ«¬"“#∞¢6ˆÁ7BFóˆ∆ˆvñ“ñ◊ñÁFÚÁFóˆ∆ˆvññ◊ñÁFÚ«¬ñ◊ñÁFÚÁFóÙñ◊ñÁFÚ«¬ñ◊ñÁFÚÁFóˆ∆ˆvññÁFW'fVÁFÚ«¬ñ◊ñÁFÚÊ∆f˜&¶ñˆÊï&ñ6ÜñW7FR«¬FóÚ«¬"“#∞¢6ˆÁ7B˜W&F˜&R“ñ◊ñÁFÚÊFˆÊT'í«¬ñ◊ñÁFÚÊ˜W&F˜&R«¬ñ◊ñÁFÚÊ˜W&F˜"«¬ñ◊ñÁFÚÊÊfñvFVD'í«¬"“#∞¢6ˆÁ7B7VG&“ñ◊ñÁFÚÁ7VG&«¬ñ◊ñÁFÚÁ7VG&76VvÊF«¬ñ◊ñÁFÚÁFV“«¬"#∞¢&WGW&‚∞¢ñE6¿¢fñ¿¢Fóˆ∆ˆvñ¿¢7FFÛ¢ñ◊ñÁFÚÊFˆÊRÚ$fGFÚ"¢$Ff&R"¿¢FFfGFÛ¢FˆÊTñÊfÚÊFFR””“"“"Ú"“"¢G∂FˆÊTñÊfÚÊFFW“G∂FˆÊTñÊfÚÁFñ÷W÷¿¢˜W&F˜&U7VG&¢∂˜W&F˜&R¬7VG&“Êfñ«FW"Çáf«VRí”‚f«VRbbf«VR”“"“"íÊ¶ˆñ‚Ç"(
""í«¬"“"¿¢6ˆ˜&FñÊFW3¢ñ◊ñÁFÚÊw5í“ÁV∆¬bbñ◊ñÁFÚÊw5Ç“ÁV∆¿¢ÚG¥ÁV÷&W"Üñ◊ñÁFÚÊw5ííÁFÙfóÜVBÉbó“¬G¥ÁV÷&W"Üñ◊ñÁFÚÊw5ÇíÁFÙfóÜVBÉbó÷ ¢¢"“ ¢”∞ß–†¶gVÊ7Fñˆ‚'Vñ∆Dñ◊ñÁFı˜6óFñˆÂW&¬Üñ◊ñÁFÚí∞¢ñbÜñ◊ñÁFÚÊw5í”“ÁV∆¬«¬ñ◊ñÁFÚÊw5Ç”“ÁV∆¬í&WGW&‚"#∞¢&WGW&‚áGG3¢Ú˜wwrÊvˆˆv∆RÊ6ˆ“ˆ÷2˜6V&6ÇÛˆì”gVW'ì“G∂VÊ6ˆFUU$î6ˆ◊ˆÊVÁBÜG∂ñ◊ñÁFÚÊw5ó“¬G∂ñ◊ñÁFÚÊw5á÷ó÷∞ß–†¶gVÊ7Fñˆ‚'Vñ∆Dñ◊ñÁFÙW&¬Üñ◊ñÁFÚí∞¢6ˆÁ7Bñ◊ñÁFÙ∂Wí“'Vñ∆Dñ◊ñÁFÙ∂WíÜñ◊ñÁFÚì∞¢6ˆÁ7B&◊2“ÊWrU$≈6V&6Ö&◊2Çì∞¢ñbá6V∆V7FVD6ˆ÷÷W76ñBí&◊2Á6WBÇ&6ˆ÷÷W76"¬6V∆V7FVD6ˆ÷÷W76ñBì∞¢ñbÜñ◊ñÁFÙ∂Wíí&◊2Á6WBÇ&ñ◊ñÁFÚ"¬ñ◊ñÁFÙ∂Wíì∞¢6ˆÁ7BÜ6Ç“&◊2ÁFı7G&ñÊrÇí«¬&Üˆ÷R#∞¢&WGW&‚G∑vñÊF˜rÊ∆ˆ6Fñˆ‚Ê˜&ñvñÁ“G∑vñÊF˜rÊ∆ˆ6Fñˆ‚ÁFÜÊ÷W“2G∂Ü6á÷∞ß–†¶gVÊ7Fñˆ‚'Vñ∆DgV∆«67&VV‰ñ◊ñÁFıvÜG4÷W76vRÜñ◊ñÁFÚí∞¢6ˆÁ7BFóÚ“ñ◊ñÁFÚÁFóÙ÷ÁWFVÁ¶ñˆÊR«¬6∆76ñgïFóÙ÷ÁWFVÁ¶ñˆÊRÜñ◊ñÁFÚÊ6ˆFñ6U&Wß¶Úì∞¢6ˆÁ7BFF“vWDñ◊ñÁFı˜WFFÜñ◊ñÁFÚ¬FóÚì∞¢6ˆÁ7B∆ñÊ∂VDÊ˜FW2“vWD6ˆ÷÷W76Ê˜FT∆ñÊ∂VDÊ˜FW2Üñ◊ñÁFÚì∞¢6ˆÁ7B∆ñÊ∂VDÊ˜FW4∆ñÊW2“∆ñÊ∂VDÊ˜FW2Ê∆VÊwFÄ¢Ú∆ñÊ∂VDÊ˜FW2Êf∆D÷ÇÜÊ˜FRí”‚∞¢“G∂vWD6ˆ÷÷W76Ê˜FUFóF∆RÜÊ˜FRó÷¿¢‚‚‚Ö7G&ñÊrÜÊ˜FRÁFWáB«¬""íÁG&ñ“ÇíÚµ7G&ñÊrÜÊ˜FRÁFWáB«¬""íÁG&ñ“Çï“¢µ“ê¢“ê¢¢≤"“%”∞¢6ˆÁ7B˜6óFñˆÂW&¬“'Vñ∆Dñ◊ñÁFı˜6óFñˆÂW&¬Üñ◊ñÁFÚì∞¢6ˆÁ7BW&¬“'Vñ∆Dñ◊ñÁFÙW&¬Üñ◊ñÁFÚì∞¢&WGW&‚∞¢%FíñÊˆ«G&ÚíFWGFv∆íFV∆Œ(	ññ◊ñÁFÛ¢"¿¢""¿¢Êˆ÷Rñ◊ñÁFÛ¢G∂ñ◊ñÁFÚÊFVÊˆ÷ñÊ¶ñˆÊR«¬"“'÷¿¢îB4¢G∂FFÊñE6÷¿¢6ˆ◊VÊS¢G∂ñ◊ñÁFÚÊ6ˆ◊VÊR«¬"“'÷¿¢ñÊFó&óß¶ÚÚfñ¢G∂FFÁfñ÷¿¢Fóˆ∆ˆvññ◊ñÁFÛ¢G∂FFÁFóˆ∆ˆvñ÷¿¢7FFÚ∆f˜&Û¢G∂FFÁ7FF˜÷¿¢FFfGFÛ¢G∂FFÊFFfGF˜÷¿¢˜W&F˜&RÚ7VG&¢G∂FFÊ˜W&F˜&U7VG&÷¿¢%6VvÊ∆¶ñˆÊR6ˆ∆∆VvF¢"¿¢‚‚Ê∆ñÊ∂VDÊ˜FW4∆ñÊW2¿¢""¿¢˜6ó¶ñˆÊRñ◊ñÁFÛ¢G∂FFÊ6ˆ˜&FñÊFW7÷¿¢˜6óFñˆÂW&¬Úø	˘8“&í˜6ó¶ñˆÊRñ◊ñÁFı“ÇG∑˜6óFñˆÂW&«“ñ¢/	˘8“&í˜6ó¶ñˆÊRñ◊ñÁFÛ¢˜6ó¶ñˆÊRÊˆ‚Fó7ˆÊñ&ñ∆R"¿¢ø	˘Ir&íñ◊ñÁFÚÊV∆Œ(	ñ“ÇG∂W&«“ñ ¢“Ê¶ˆñ‚Ç%∆‚"ì∞ß–†¶gVÊ7Fñˆ‚˜V‰gV∆«67&VV‰ñ◊ñÁFıvÜG4Üñ◊ñÁFÚí∞¢6ˆÁ7B÷W76vR“'Vñ∆DgV∆«67&VV‰ñ◊ñÁFıvÜG4÷W76vRÜñ◊ñÁFÚì∞¢ñbÇ6fT˜VÂvÜG4÷W76vRÜ÷W76vRíí∆W'BÇ$ñ◊˜76ñ&ñ∆R&ó&RvÜG47RVW7FÚFó7˜6óFófÚ‚"ì∞ß–†¶gVÊ7Fñˆ‚'Vñ∆Dñ◊ñÁFÙ÷˜WÜñ◊ñÁFÚ¬FóÚ¬˜FñˆÁ2“∑“í∞¢6ˆÁ7Bñ◊ñÁFÙ∂Wí“'Vñ∆Dñ◊ñÁFÙ∂WíÜñ◊ñÁFÚì∞¢6ˆÁ7B∆ñÊ∂VDÊ˜FW2“vWD6ˆ÷÷W76Ê˜FT∆ñÊ∂VDÊ˜FW2Üñ◊ñÁFÚì∞¢6ˆÁ7B˜WFF“vWDñ◊ñÁFı˜WFFÜñ◊ñÁFÚ¬FóÚì∞¢6ˆÁ7BÊ˜FTñ◊ñÁFÚ“vWDñ◊ñÁFı˜WÊ˜FW2Üñ◊ñÁFÚì∞¢6ˆÁ7B6Ü˜u6ñ‰ÜVFW"“&ˆˆ∆V‚Ü˜FñˆÁ3ÚÁ6Ü˜u6ñ‰ÜVFW"ì∞¢6ˆÁ7BvÜG67Fñˆ‚“˜FñˆÁ3ÚÊgV∆«67&VVÂvÜG4Ú&gV∆«67&VV‚◊vÜG6"¢'vÜG6#∞¢6ˆÁ7B∆ñÊ∂VDÊ˜FW4÷&∑W“∆ñÊ∂VDÊ˜FW2Ê∆VÊwFÄ¢Ú∆ñÊ∂VDÊ˜FW2Ê÷ÇÜÊ˜FRí”‚ ¢∆'WGFˆ‚GóS“&'WGFˆ‚"6∆73“&÷◊˜W÷Ê˜FR÷'F‚"FF÷÷◊˜W÷7Fñˆ„“&Ê˜FR"FF÷Ê˜FR÷ñC“"G∂W66TÖD‘¬ÜÊ˜FRÊñB«¬""ó“#‡¢G∂W66TÖD‘¬ÜvWD6ˆ÷÷W76Ê˜FUFóF∆RÜÊ˜FRíó–¢¬ˆ'WGFˆ„‡¢íÊ¶ˆñ‚Ç""ê¢¢#«7„‚”¬˜7„‚#∞¢6ˆÁ7BÜVFW%7V'FóF∆R“6Ü˜u6ñ‰ÜVFW ¢Ú«6∆73“&÷◊˜W◊7V'FóF∆R#‰îB4¢G∂W66TÖD‘¬á˜WFFÊñE6ó”¬˜Ê ¢¢"#∞¢6ˆÁ7BñE6FWFñ¬“6Ü˜u6ñ‰ÜVFW ¢Ú" ¢¢∆Fóc„∆GC‰îB4¬ˆGC„∆FC‚G∂W66TÖD‘¬á˜WFFÊñE6ó”¬ˆFC„¬ˆFócÊ∞†¢&WGW&‚ ¢∆Fób6∆73“&÷◊˜W÷6&B"FF÷ñ◊ñÁFÚ÷∂Wì“"G∂W66TÖD‘¬Üñ◊ñÁFÙ∂Wíó“#‡¢∆Fób6∆73“&÷◊˜W÷ÜVFW"#‡¢∆Fób6∆73“&÷◊˜W◊FóF∆R#‡¢∆É3‚G∂'Vñ∆Dñ◊ñÁFÙ÷&∂W$&FvRÜñ◊ñÁFÚó“G∂W66TÖD‘¬Üñ◊ñÁFÚÊFVÊˆ÷ñÊ¶ñˆÊR«¬$ñ◊ñÁFÚ"ó”¬ˆÉ3‡¢G∂'Vñ∆Dñ◊ñÁFıvVFÜW$&FvT÷&∑WÜñ◊ñÁFÚó–¢G∂ÜVFW%7V'FóF∆W–¢¬ˆFóc‡¢∆'WGFˆ‚GóS“&'WGFˆ‚"6∆73“&÷◊˜W÷6∆˜6R÷'F‚"FF÷÷◊˜W÷7Fñˆ„“&6∆˜6R"&ñ÷∆&V√“$6ÜóVFí˜WFWGFv∆ñÚñ◊ñÁFÚ"FóF∆S“$6ÜóVFí#Ï9s¬ˆ'WGFˆ„‡¢¬ˆFóc‡¢∆Fób6∆73“&÷◊˜W◊67&ˆ∆¬#‡¢∆F¬6∆73“&÷◊˜W÷FWFñ«2#‡¢G∂ñE6FWFñ«–¢∆Fóc„∆GC‰6ˆ◊VÊS¬ˆGC„∆FC‚G∂W66TÖD‘¬Üñ◊ñÁFÚÊ6ˆ◊VÊR«¬"“"ó”¬ˆFC„¬ˆFóc‡¢∆Fóc„∆GC‰ñÊFó&óß¶ÚÚfñ¬ˆGC„∆FC‚G∂W66TÖD‘¬á˜WFFÁfñó”¬ˆFC„¬ˆFóc‡¢∆Fóc„∆GCÂFóˆ∆ˆvññ◊ñÁFÛ¬ˆGC„∆FC‚G∂W66TÖD‘¬á˜WFFÁFóˆ∆ˆvñó”¬ˆFC„¬ˆFóc‡¢∆Fóc„∆GCÂ7FFÚ∆f˜&Û¬ˆGC„∆FC‚G∂W66TÖD‘¬á˜WFFÁ7FFÚó”¬ˆFC„¬ˆFóc‡¢∆Fóc„∆GC‰FFfGFÛ¬ˆGC„∆FC‚G∂W66TÖD‘¬á˜WFFÊFFfGFÚó”¬ˆFC„¬ˆFóc‡¢∆Fóc„∆GC‰˜W&F˜&RÚ7VG&¬ˆGC„∆FC‚G∂W66TÖD‘¬á˜WFFÊ˜W&F˜&U7VG&ó”¬ˆFC„¬ˆFóc‡¢∆Fóc„∆GCÂ6VvÊ∆¶ñˆÊR6ˆ∆∆VvF¬ˆGC„∆FB6∆73“&÷◊˜W÷Ê˜FW2÷∆ó7B#‚G∂∆ñÊ∂VDÊ˜FW4÷&∑W”¬ˆFC„¬ˆFóc‡¢∆Fóc„∆GC‰Ê˜FRñ◊ñÁFÛ¬ˆGC„∆FC‚G∂W66TÖD‘¬ÜÊ˜FTñ◊ñÁFÚ«¬"“"ó”¬ˆFC„¬ˆFóc‡¢∆Fóc„∆GC‰6ˆ˜&FñÊFRu3¬ˆGC„∆FC‚G∂W66TÖD‘¬á˜WFFÊ6ˆ˜&FñÊFW2ó”¬ˆFC„¬ˆFóc‡¢¬ˆF√‡¢¬ˆFóc‡¢∆Fób6∆73“&÷◊˜W÷7FñˆÁ2#‡¢∆'WGFˆ‚GóS“&'WGFˆ‚"6∆73“&'F‚'F‚◊6÷∆¬'F‚◊&ñ÷'í"FF÷÷◊˜W÷7Fñˆ„“&ÊfñvFR"FF÷ñ◊ñÁFÚ÷∂Wì“"G∂W66TÖD‘¬Üñ◊ñÁFÙ∂Wíó“#‰‰dît¬ˆ'WGFˆ„‡¢∆'WGFˆ‚GóS“&'WGFˆ‚"6∆73“&'F‚'F‚◊6÷∆¬'F‚◊vÜG6"FF÷÷◊˜W÷7Fñˆ„“"G∂W66TÖD‘¬ávÜG67Fñˆ‚ó“"FF÷ñ◊ñÁFÚ÷∂Wì“"G∂W66TÖD‘¬Üñ◊ñÁFÙ∂Wíó“#ÂtÑE4¬ˆ'WGFˆ„‡¢∆'WGFˆ‚GóS“&'WGFˆ‚"6∆73“&'F‚'F‚◊6÷∆¬"FF÷÷◊˜W÷7Fñˆ„“&FWFñ¬"FF÷ñ◊ñÁFÚ÷∂Wì“"G∂W66TÖD‘¬Üñ◊ñÁFÙ∂Wíó“#‰DUEDtƒîÚî’îÂDÛ¬ˆ'WGFˆ„‡¢¬ˆFóc‡¢¬ˆFóc‡¢∞ß–†¶gVÊ7Fñˆ‚6V∆V7Dñ◊ñÁFÙf˜$÷FWFñ¬Üñ◊ñÁFÚí∞¢6ˆÁ7B∂Wí“'Vñ∆Dñ◊ñÁFÙ∂WíÜñ◊ñÁFÚì∞¢ñbÇ∂Wíí&WGW&„∞¢6V∆V7FVDñ◊ñÁFÙñB“∂Wì∞¢6V∆V7FVDgV∆«67&VV‰ñ◊ñÁFÙñB“∂Wì∞¢6V∆V7FVDñ◊ñÁFÙFF“≤‚‚Êñ◊ñÁFÚ”∞¢&VÊFW%6V∆V7FVDñ◊ñÁFÙFWFñ≈ÊV¬Çì∞ß–†¶gVÊ7Fñˆ‚6∆˜6U6V∆V7FVDñ◊ñÁFÙFWFñ¬á≤6∆˜6U˜W“f«6R““∑“í∞¢6V∆V7FVDñ◊ñÁFÙñB“"#∞¢6V∆V7FVDñ◊ñÁFÙFF“ÁV∆√∞¢6V∆V7FVDgV∆«67&VV‰ñ◊ñÁFÙñB“"#∞¢VíÊ÷ñ‰÷ñ◊ñÁFÙFWFñ≈ÊV√ÚÊ6∆74∆ó7BÊFBÇ&ÜñFFV‚"ì∞¢VíÊ÷ñ◊ñÁFÙFWFñ≈ÊV√ÚÊ6∆74∆ó7BÊFBÇ&ÜñFFV‚"ì∞¢ñbáVíÊ÷ñ‰÷ñ◊ñÁFÙFWFñƒ&ˆGííVíÊ÷ñ‰÷ñ◊ñÁFÙFWFñƒ&ˆGíÊñÊÊW$ÖD‘¬“"#∞¢ñbáVíÊ÷ñ◊ñÁFÙFWFñƒ&ˆGííVíÊ÷ñ◊ñÁFÙFWFñƒ&ˆGíÊñÊÊW$ÖD‘¬“"#∞¢ñbÜ6∆˜6U˜WígV∆«67&VV‰÷Ê6∆˜6U˜WÇì∞ß–†¶gVÊ7Fñˆ‚&VÊFW%6V∆V7FVDñ◊ñÁFÙFWFñ≈ÊV¬Çí∞¢6ˆÁ7BÊV«2“∞¢≤ÊV√¢VíÊ÷ñ‰÷ñ◊ñÁFÙFWFñ≈ÊV¬¬&ˆGì¢VíÊ÷ñ‰÷ñ◊ñÁFÙFWFñƒ&ˆGí“¿¢≤ÊV√¢VíÊ÷ñ◊ñÁFÙFWFñ≈ÊV¬¬&ˆGì¢VíÊ÷ñ◊ñÁFÙFWFñƒ&ˆGí–¢“Êfñ«FW"ÇÜVÁG'íí”‚VÁG'íÁÊV¬bbVÁG'íÊ&ˆGíì∞¢ñbÇÊV«2Ê∆VÊwFÇí&WGW&„∞¢ñbÇ6V∆V7FVDñ◊ñÁFÙñB«¬6V∆V7FVDñ◊ñÁFÙFFí∞¢ÊV«2Êf˜$V6ÇÇá≤ÊV¬¬&ˆGí“í”‚∞¢ÊV¬Ê6∆74∆ó7BÊFBÇ&ÜñFFV‚"ì∞¢&ˆGíÊñÊÊW$ÖD‘¬“"#∞¢“ì∞¢&WGW&„∞¢–¢6ˆÁ7BFóÚ“6V∆V7FVDñ◊ñÁFÙFFÁFóÙ÷ÁWFVÁ¶ñˆÊR«¬6∆76ñgïFóÙ÷ÁWFVÁ¶ñˆÊRá6V∆V7FVDñ◊ñÁFÙFFÊ6ˆFñ6U&Wß¶Úì∞¢ÊV«2Êf˜$V6ÇÇá≤ÊV¬¬&ˆGí“í”‚∞¢6ˆÁ7Bó4gV∆«67&VVÂÊV¬“ÊV¬””“VíÊ÷ñ◊ñÁFÙFWFñ≈ÊV√∞¢&ˆGíÊñÊÊW$ÖD‘¬“'Vñ∆Dñ◊ñÁFÙ÷˜Wá6V∆V7FVDñ◊ñÁFÙFF¬FóÚ¬∞¢6Ü˜u6ñ‰ÜVFW#¢ó4gV∆«67&VVÂÊV¬¿¢gV∆«67&VVÂvÜG4¢ó4gV∆«67&VVÂÊV¿¢“ì∞¢ÊV¬Ê6∆74∆ó7BÁ&V÷˜fRÇ&ÜñFFV‚"ì∞¢“ì∞¢&ñÊEW'6ó7FVÁDñ◊ñÁFÙFWFñƒ7FñˆÁ2Çì∞ß–†¶gVÊ7Fñˆ‚&ñÊEW'6ó7FVÁDñ◊ñÁFÙFWFñƒ7FñˆÁ2Çí∞¢∑VíÊ÷ñ‰÷ñ◊ñÁFÙFWFñ≈ÊV¬¬VíÊ÷ñ◊ñÁFÙFWFñ≈ÊV≈“Êfñ«FW"Ñ&ˆˆ∆V‚íÊf˜$V6ÇÇáÊV¬í”‚∞¢ÊV¬ÁVW'ï6V∆V7F˜$∆¬Ç%∂FF÷÷◊˜W÷7Fñˆ„“vÊfñvFRu“"íÊf˜$V6ÇÇÜ'WGFˆ‚í”‚∞¢'WGFˆ‚ÊFDWfVÁD∆ó7FVÊW"Ç&6∆ñ6≤"¬7ñÊ2Çí”‚∞¢6ˆÁ7B∂Wí“'WGFˆ‚ÊvWDGG&ñ'WFRÇ&FF÷ñ◊ñÁFÚ÷∂Wí"í«¬6V∆V7FVDñ◊ñÁFÙñC∞¢6ˆÁ7Bñ◊ñÁFÚ“fñÊD7W'&VÁDñ◊ñÁFÙ'î∂WíÜ∂Wíí«¬6V∆V7FVDñ◊ñÁFÙFF∞¢ñbÇñ◊ñÁFÚí&WGW&„∞¢vóBÊfñvFUFÙñ◊ñÁFÚÜñ◊ñÁFÚì∞¢“ì∞¢“ì∞¢ÊV¬ÁVW'ï6V∆V7F˜$∆¬Ç%∂FF÷÷◊˜W÷7Fñˆ„“wvÜG6u“"íÊf˜$V6ÇÇÜ'WGFˆ‚í”‚∞¢'WGFˆ‚ÊFDWfVÁD∆ó7FVÊW"Ç&6∆ñ6≤"¬7ñÊ2Çí”‚∞¢6ˆÁ7B∂Wí“'WGFˆ‚ÊvWDGG&ñ'WFRÇ&FF÷ñ◊ñÁFÚ÷∂Wí"í«¬6V∆V7FVDñ◊ñÁFÙñC∞¢6ˆÁ7Bñ◊ñÁFÚ“fñÊD7W'&VÁDñ◊ñÁFÙ'î∂WíÜ∂Wíí«¬6V∆V7FVDñ◊ñÁFÙFF∞¢ñbÇñ◊ñÁFÚí&WGW&„∞¢vóBÜÊF∆Tñ◊ñÁFıvÜG46∆ñ6≤Üñ◊ñÁFÚì∞¢“ì∞¢“ì∞¢ÊV¬ÁVW'ï6V∆V7F˜$∆¬Ç%∂FF÷÷◊˜W÷7Fñˆ„“vgV∆«67&VV‚◊vÜG6u“"íÊf˜$V6ÇÇÜ'WGFˆ‚í”‚∞¢'WGFˆ‚ÊFDWfVÁD∆ó7FVÊW"Ç&6∆ñ6≤"¬Çí”‚∞¢6ˆÁ7B∂Wí“'WGFˆ‚ÊvWDGG&ñ'WFRÇ&FF÷ñ◊ñÁFÚ÷∂Wí"í«¬6V∆V7FVDñ◊ñÁFÙñC∞¢6ˆÁ7Bñ◊ñÁFÚ“fñÊD7W'&VÁDñ◊ñÁFÙ'î∂WíÜ∂Wíí«¬6V∆V7FVDñ◊ñÁFÙFF∞¢ñbÇñ◊ñÁFÚí&WGW&„∞¢˜V‰gV∆«67&VV‰ñ◊ñÁFıvÜG4Üñ◊ñÁFÚì∞¢“ì∞¢“ì∞¢ÊV¬ÁVW'ï6V∆V7F˜$∆¬Ç%∂FF÷÷◊˜W÷7Fñˆ„“vFWFñ¬u“"íÊf˜$V6ÇÇÜ'WGFˆ‚í”‚∞¢'WGFˆ‚ÊFDWfVÁD∆ó7FVÊW"Ç&6∆ñ6≤"¬Çí”‚∞¢6ˆÁ7B∂Wí“'WGFˆ‚ÊvWDGG&ñ'WFRÇ&FF÷ñ◊ñÁFÚ÷∂Wí"í«¬6V∆V7FVDñ◊ñÁFÙñC∞¢6ˆÁ7Bñ◊ñÁFÚ“fñÊD7W'&VÁDñ◊ñÁFÙ'î∂WíÜ∂Wíí«¬6V∆V7FVDñ◊ñÁFÙFF∞¢ñbÇñ◊ñÁFÚí&WGW&„∞¢fˆ7W4ñ◊ñÁFÙñ‰∆ó7BÜñ◊ñÁFÚ¬G'VRì∞¢6∆˜6T÷gV∆«67&VVÂvRÇì∞¢“ì∞¢“ì∞¢ÊV¬ÁVW'ï6V∆V7F˜$∆¬Ç%∂FF÷÷◊˜W÷7Fñˆ„“vÊ˜FRu“"íÊf˜$V6ÇÇÜ'WGFˆ‚í”‚∞¢'WGFˆ‚ÊFDWfVÁD∆ó7FVÊW"Ç&6∆ñ6≤"¬Çí”‚∞¢6ˆÁ7BÊ˜FTñB“'WGFˆ‚ÊvWDGG&ñ'WFRÇ&FF÷Ê˜FR÷ñB"í«¬"#∞¢6ˆÁ7BÊ˜FR“7W'&VÁD6ˆ÷÷W76Ê˜FW2ÊfñÊBÇÜóFV“í”‚óFV“ÊñB””“Ê˜FTñBì∞¢ñbÇÊ˜FRí&WGW&„∞¢˜V‰6ˆ÷÷W76Ê˜FW5vRÇì∞¢6WEFñ÷V˜WBÇÇí”‚˜V‰6ˆ÷÷W76Ê˜FTFWFñ¬ÜÊ˜FRí¬Sì∞¢“ì∞¢“ì∞¢ÊV¬ÁVW'ï6V∆V7F˜$∆¬Ç%∂FF÷÷◊˜W÷7Fñˆ„“v6∆˜6Ru“"íÊf˜$V6ÇÇÜ'WGFˆ‚í”‚∞¢'WGFˆ‚ÊFDWfVÁD∆ó7FVÊW"Ç&6∆ñ6≤"¬Ü6∆ñ6¥WfVÁBí”‚∞¢6∆ñ6¥WfVÁBÁ&WfVÁDFVfV«BÇì∞¢6∆ñ6¥WfVÁBÁ7F˜&˜vFñˆ‚Çì∞¢6∆˜6U6V∆V7FVDñ◊ñÁFÙFWFñ¬á≤6∆˜6U˜W¢G'VR“ì∞¢“ì∞¢“ì∞¢“ì∞ß–†¶gVÊ7Fñˆ‚vWDñ◊ñÁFı˜WÊ˜FW2Üñ◊ñÁFÚí∞¢&WGW&‚∞¢ñ◊ñÁFÚÊÊ˜FTñ◊ñÁFÚ¿¢ñ◊ñÁFÚÊÊ˜FR¿¢ñ◊ñÁFÚÊÊ˜FW2¿¢ñ◊ñÁFÚÊÊÊ˜F¶ñˆÊí¿¢ñ◊ñÁFÚÊFW67&ó¶ñˆÊR¿¢ñ◊ñÁFÚÊFW67&ó¶ñˆÊTñ◊ñÁF¢“Ê÷Çáf«VRí”‚7G&ñÊráf«VR«¬""íÁG&ñ“ÇííÊfñÊBÑ&ˆˆ∆V‚í«¬"#∞ß–†¶gVÊ7Fñˆ‚fñÊD7W'&VÁDñ◊ñÁFÙ'î∂WíÜ∂Wíí∞¢&WGW&‚7W'&VÁDñ◊ñÁFíÊfñÊBÇÜóFV“í”‚'Vñ∆Dñ◊ñÁFÙ∂WíÜóFV“í””“∂Wíì∞ß–†¶gVÊ7Fñˆ‚&ñÊDñ◊ñÁFÙ÷˜W7FñˆÁ2ÜWfVÁB¬˜W÷í∞¢6ˆÁ7B˜WV∆V÷VÁB“WfVÁBÁ˜WÚÊvWDV∆V÷VÁBÇì∞¢ñbÇ˜WV∆V÷VÁBí&WGW&„∞¢6ˆÁ7B6&B“˜WV∆V÷VÁBÁVW'ï6V∆V7F˜"Ç"Ê÷◊˜W÷6&E∂FF÷ñ◊ñÁFÚ÷∂Wï“"ì∞¢6ˆÁ7B˜W∂Wí“6&CÚÊvWDGG&ñ'WFRÇ&FF÷ñ◊ñÁFÚ÷∂Wí"í«¬"#∞†¢˜WV∆V÷VÁBÁVW'ï6V∆V7F˜$∆¬Ç%∂FF÷÷◊˜W÷7Fñˆ„“vÊfñvFRu“"íÊf˜$V6ÇÇÜ'WGFˆ‚í”‚∞¢'WGFˆ‚ÊFDWfVÁD∆ó7FVÊW"Ç&6∆ñ6≤"¬7ñÊ2Çí”‚∞¢6ˆÁ7B∂Wí“'WGFˆ‚ÊvWDGG&ñ'WFRÇ&FF÷ñ◊ñÁFÚ÷∂Wí"í«¬˜W∂Wì∞¢6ˆÁ7Bñ◊ñÁFÚ“fñÊD7W'&VÁDñ◊ñÁFÙ'î∂WíÜ∂Wíì∞¢ñbÇñ◊ñÁFÚí&WGW&„∞¢vóBÊfñvFUFÙñ◊ñÁFÚÜñ◊ñÁFÚì∞¢“ì∞¢“ì∞¢˜WV∆V÷VÁBÁVW'ï6V∆V7F˜$∆¬Ç%∂FF÷÷◊˜W÷7Fñˆ„“wvÜG6u“"íÊf˜$V6ÇÇÜ'WGFˆ‚í”‚∞¢'WGFˆ‚ÊFDWfVÁD∆ó7FVÊW"Ç&6∆ñ6≤"¬7ñÊ2Çí”‚∞¢6ˆÁ7B∂Wí“'WGFˆ‚ÊvWDGG&ñ'WFRÇ&FF÷ñ◊ñÁFÚ÷∂Wí"í«¬˜W∂Wì∞¢6ˆÁ7Bñ◊ñÁFÚ“fñÊD7W'&VÁDñ◊ñÁFÙ'î∂WíÜ∂Wíì∞¢ñbÇñ◊ñÁFÚí&WGW&„∞¢vóBÜÊF∆Tñ◊ñÁFıvÜG46∆ñ6≤Üñ◊ñÁFÚì∞¢“ì∞¢“ì∞¢˜WV∆V÷VÁBÁVW'ï6V∆V7F˜$∆¬Ç%∂FF÷÷◊˜W÷7Fñˆ„“vgV∆«67&VV‚◊vÜG6u“"íÊf˜$V6ÇÇÜ'WGFˆ‚í”‚∞¢'WGFˆ‚ÊFDWfVÁD∆ó7FVÊW"Ç&6∆ñ6≤"¬Çí”‚∞¢6ˆÁ7B∂Wí“'WGFˆ‚ÊvWDGG&ñ'WFRÇ&FF÷ñ◊ñÁFÚ÷∂Wí"í«¬˜W∂Wì∞¢6ˆÁ7Bñ◊ñÁFÚ“fñÊD7W'&VÁDñ◊ñÁFÙ'î∂WíÜ∂Wíì∞¢ñbÇñ◊ñÁFÚí&WGW&„∞¢˜V‰gV∆«67&VV‰ñ◊ñÁFıvÜG4Üñ◊ñÁFÚì∞¢“ì∞¢“ì∞¢˜WV∆V÷VÁBÁVW'ï6V∆V7F˜$∆¬Ç%∂FF÷÷◊˜W÷7Fñˆ„“vFWFñ¬u“"íÊf˜$V6ÇÇÜ'WGFˆ‚í”‚∞¢'WGFˆ‚ÊFDWfVÁD∆ó7FVÊW"Ç&6∆ñ6≤"¬Çí”‚∞¢6ˆÁ7B∂Wí“'WGFˆ‚ÊvWDGG&ñ'WFRÇ&FF÷ñ◊ñÁFÚ÷∂Wí"í«¬˜W∂Wì∞¢6ˆÁ7Bñ◊ñÁFÚ“fñÊD7W'&VÁDñ◊ñÁFÙ'î∂WíÜ∂Wíì∞¢ñbÇñ◊ñÁFÚí&WGW&„∞¢6V∆V7FVDgV∆«67&VV‰ñ◊ñÁFÙñB“˜W÷””“gV∆«67&VV‰÷Ú∂Wí¢6V∆V7FVDgV∆«67&VV‰ñ◊ñÁFÙñC∞¢fˆ7W4ñ◊ñÁFÙñ‰∆ó7BÜñ◊ñÁFÚ¬G'VRì∞¢ñbá˜W÷””“gV∆«67&VV‰÷í6∆˜6T÷gV∆«67&VVÂvRÇì∞¢“ì∞¢“ì∞¢˜WV∆V÷VÁBÁVW'ï6V∆V7F˜$∆¬Ç%∂FF÷÷◊˜W÷7Fñˆ„“vÊ˜FRu“"íÊf˜$V6ÇÇÜ'WGFˆ‚í”‚∞¢'WGFˆ‚ÊFDWfVÁD∆ó7FVÊW"Ç&6∆ñ6≤"¬Çí”‚∞¢6ˆÁ7BÊ˜FTñB“'WGFˆ‚ÊvWDGG&ñ'WFRÇ&FF÷Ê˜FR÷ñB"í«¬"#∞¢6ˆÁ7BÊ˜FR“7W'&VÁD6ˆ÷÷W76Ê˜FW2ÊfñÊBÇÜóFV“í”‚óFV“ÊñB””“Ê˜FTñBì∞¢ñbÇÊ˜FRí&WGW&„∞¢˜V‰6ˆ÷÷W76Ê˜FW5vRÇì∞¢6WEFñ÷V˜WBÇÇí”‚˜V‰6ˆ÷÷W76Ê˜FTFWFñ¬ÜÊ˜FRí¬Sì∞¢“ì∞¢“ì∞¢˜WV∆V÷VÁBÁVW'ï6V∆V7F˜$∆¬Ç%∂FF÷÷◊˜W÷7Fñˆ„“v6∆˜6Ru“"íÊf˜$V6ÇÇÜ'WGFˆ‚í”‚∞¢'WGFˆ‚ÊFDWfVÁD∆ó7FVÊW"Ç&6∆ñ6≤"¬Ü6∆ñ6¥WfVÁBí”‚∞¢6∆ñ6¥WfVÁBÁ&WfVÁDFVfV«BÇì∞¢6∆ñ6¥WfVÁBÁ7F˜&˜vFñˆ‚Çì∞¢ñbá˜W÷””“gV∆«67&VV‰÷bb˜W∂Wí””“6V∆V7FVDñ◊ñÁFÙñBí6∆˜6U6V∆V7FVDñ◊ñÁFÙFWFñ¬á≤6∆˜6U˜W¢f«6R“ì∞¢˜W÷Ê6∆˜6U˜WÜWfVÁBÁ˜Wì∞¢“ì∞¢“ì∞ß–†¶÷Êˆ‚Ç'˜W˜V‚"¬ÜWfVÁBí”‚&ñÊDñ◊ñÁFÙ÷˜W7FñˆÁ2ÜWfVÁB¬÷íì∞¶gV∆«67&VV‰÷Êˆ‚Ç'˜W˜V‚"¬ÜWfVÁBí”‚&ñÊDñ◊ñÁFÙ÷˜W7FñˆÁ2ÜWfVÁB¬gV∆«67&VV‰÷íì∞¶÷Êˆ‚Ç&÷˜fVVÊB¶ˆˆ÷VÊB"¬Çí”‚&V∆ˆD6ˆ÷÷W76vVFÜW$f˜%fó6ñ&∆Tñ◊ñÁFíÇíì∞¶gV∆«67&VV‰÷Êˆ‚Ç&÷˜fVVÊB¶ˆˆ÷VÊB"¬Çí”‚&V∆ˆDñ◊ñÁFïvVFÜW"ÜvWEfó6ñ&∆T÷ñ◊ñÁFíÜgV∆«67&VV‰÷¬7W'&VÁDñ◊ñÁFíí¬≤∆ñ÷óC¢î’îÂDııtTDÑU%ı$Te$U4ÖÙƒî‘ïB¬&VfW$ÊV&W7C¢G'VR“íì∞¶v∆ˆ&ƒ÷Êˆ‚Ç&÷˜fVVÊB¶ˆˆ÷VÊB"¬Çí”‚&V∆ˆDñ◊ñÁFïvVFÜW"ÜvWEfó6ñ&∆T÷ñ◊ñÁFíÜv∆ˆ&ƒ÷¬v∆ˆ&ƒñ◊ñÁFíííì∞†¶gVÊ7Fñˆ‚fˆ7W4ñ◊ñÁFÙñ‰∆ó7BÜñ◊ñÁFÚ¬67&ˆ∆¬“G'VRí∞¢6ˆÁ7B∂Wí“'Vñ∆Dñ◊ñÁFÙ∂WíÜñ◊ñÁFÚì∞¢ÜñvÜ∆ñváFVDñ◊ñÁFÙ∂Wí“∂Wì∞¢WáÊFVDñ◊ñÁFÙ∂Wí“∂Wì∞¢&VÊFW$ñ◊ñÁFíÇì∞¢ñbÇ67&ˆ∆¬í&WGW&„∞¢6ˆÁ7B&˜r“VíÊñ◊ñÁFî∆ó7FÁVW'ï6V∆V7F˜"Ü∂FF÷ñ◊ñÁFÚ÷∂Wì’¬"G∂774W66Uf«VRÜ∂Wíó’¬%÷ì∞¢ñbÇ&˜rí&WGW&„∞¢VíÊñ◊ñÁFî∆ó7FÁVW'ï6V∆V7F˜$∆¬Ç"Êñ◊ñÁFÚ÷óFV“ÊÜñvÜ∆ñváB"íÊf˜$V6ÇÇÜV¬í”‚V¬Ê6∆74∆ó7BÁ&V÷˜fRÇ&ÜñvÜ∆ñváB"íì∞¢&˜rÊ6∆74∆ó7BÊFBÇ&ÜñvÜ∆ñváB"ì∞¢&˜rÁ67&ˆ∆ƒñÁFıfñWrá≤&VÜfñ˜#¢'6÷ˆ˜FÇ"¬&∆ˆ6≥¢&6VÁFW""“ì∞ß–†¶gVÊ7Fñˆ‚774W66Uf«VRáf«VRí∞¢ñbávñÊF˜r‰552bbGóVˆbvñÊF˜r‰552ÊW66R””“&gVÊ7Fñˆ‚"í&WGW&‚vñÊF˜r‰552ÊW66Ráf«VRì∞¢&WGW&‚7G&ñÊráf«VRíÁ&W∆6RÇı≤%≈≈“ˆr¬%≈¬Bb"ì∞ß–†¶gVÊ7Fñˆ‚vWD÷&∂W$6∆72Üñ◊ñÁFÚí∞¢6ˆÁ7BFˆÊR“&ˆˆ∆V‚Üñ◊ñÁFÚÊFˆÊRì∞¢6ˆÁ7B7G&˜&FñÊ&ñÚ“ñ◊ñÁFÚÊÜ57G&˜&FñÊ&ñÚÛÚÜ57G&˜&FñÊ&ñÚÜñ◊ñÁFÚÊ6ˆFñ6U&Wß¶Úì∞¢ñbÜFˆÊRí&WGW&‚&FˆÊR#∞¢ñbá7G&˜&FñÊ&ñÚí&WGW&‚'7G&˜&FñÊ&ñÚ#∞¢&WGW&‚'FˆFÚ#∞ß–†¶gVÊ7Fñˆ‚WFFTñ◊ñÁFÙ∆ˆ6≈7FFRÜñ◊ñÁFÙñG2¬F6Çí∞¢6ˆÁ7BñE6WB“ÊWr6WBÜñ◊ñÁFÙñG2ì∞¢7W'&VÁDñ◊ñÁFí“7W'&VÁDñ◊ñÁFíÊ÷ÇÜóFV“í”‚Ä¢vWDñ◊ñÁFÙFˆ4ñG2ÜóFV“íÁ6ˆ÷RÇÜñBí”‚ñE6WBÊÜ2ÜñBííÚ≤‚‚ÊóFV“¬‚‚ÁF6Ç“¢óFV–¢íì∞¢&VÊFW$ñ◊ñÁFíÇì∞¢&VÊFW$÷Çì∞ß–†¶gVÊ7Fñˆ‚vWDñ◊ñÁFÙFˆ4ñG2Üñ◊ñÁFÚí∞¢ñbÑ'&íÊó4'&íÜñ◊ñÁFÚÁ6˜W&6TñG2íbbñ◊ñÁFÚÁ6˜W&6TñG2Ê∆VÊwFÇí&WGW&‚ñ◊ñÁFÚÁ6˜W&6TñG3∞¢&WGW&‚ñ◊ñÁFÚÊñBÚ∂ñ◊ñÁFÚÊñE“¢µ”∞ß–†¶gVÊ7Fñˆ‚ó4'Vñ«DñÂ7WW$F÷ñ‰V÷ñ¬ÜV÷ñ¬í∞¢6ˆÁ7BÊ˜&÷∆ó¶VB“Ê˜&÷∆ó¶TV÷ñ¬ÜV÷ñ¬ì∞¢&WGW&‚%Tî≈EÙîÂı5UU%ÙD‘îÂÙT‘î≈2Á6ˆ÷RÇÜF÷ñ‰V÷ñ¬í”‚Ê˜&÷∆ó¶TV÷ñ¬ÜF÷ñ‰V÷ñ¬í””“Ê˜&÷∆ó¶VBì∞ß–†¶gVÊ7Fñˆ‚6‰÷ÊvTFFÇí∞¢6ˆÁ7BV÷ñ¬“Ê˜&÷∆ó¶TV÷ñ¬Ü7W'&VÁEW6W#ÚÊV÷ñ¬«¬""ì∞¢&WGW&‚ó4'Vñ«DñÂ7WW$F÷ñ‰V÷ñ¬ÜV÷ñ¬í«¬F÷ñ‰V÷ñ«2ÊÜ2ÜV÷ñ¬ì∞ß–†¶gVÊ7Fñˆ‚Ê˜&÷∆ó¶TV÷ñ¬ÜV÷ñ¬í∞¢&WGW&‚7G&ñÊrÜV÷ñ¬«¬""íÁG&ñ“ÇíÁFÙ∆˜vW$66RÇì∞ß–†¶gVÊ7Fñˆ‚6∆V$÷Çí∞¢ñ◊ñÁFÙ÷&∂W$'î∂WíÊ6∆V"Çì∞¢gV∆«67&VV‰ñ◊ñÁFÙ÷&∂W$'î∂WíÊ6∆V"Çì∞¢÷&∂W$∆ñW"Ê6∆V$∆ñW'2Çì∞¢6Ê˜u&ˆD∆ñW"Ê6∆V$∆ñW'2Çì∞¢gV∆«67&VVÂ6Ê˜u&ˆD∆ñW"Ê6∆V$∆ñW'2Çì∞¢gV∆«67&VV‰÷&∂W$∆ñW"Ê6∆V$∆ñW'2Çì∞ß–†¶gVÊ7Fñˆ‚vWD÷Wß¶Ù'î∆&V¬Ü∆&V¬í∞¢6ˆÁ7BÊ˜&÷∆ó¶VB“7G&ñÊrÜ∆&V¬«¬""íÁG&ñ“ÇíÁFÙ∆˜vW$66RÇì∞¢ñbÇÊ˜&÷∆ó¶VBí&WGW&‚ÁV∆√∞¢6ˆÁ7BWÜ7B“÷Wß¶ï&V6˜&G2ÊfñÊBÇÜ÷Wß¶Úí”‚∞¢6ˆÁ7B‰ñB“7G&ñÊrÜ÷Wß¶ÚÊ‰ñB«¬""íÁFÙ∆˜vW$66RÇì∞¢6ˆÁ7BÊˆ÷R“7G&ñÊrÜ÷Wß¶ÚÊÊˆ÷R«¬""íÁFÙ∆˜vW$66RÇì∞¢&WGW&‚‰ñB””“Ê˜&÷∆ó¶VB«¬Êˆ÷R””“Ê˜&÷∆ó¶VC∞¢“ì∞¢ñbÜWÜ7Bí&WGW&‚WÜ7C∞¢6ˆÁ7B'î‰ñD6ˆÁFñÁ2“÷Wß¶ï&V6˜&G2ÊfñÊBÇÜ÷Wß¶Úí”‚∞¢6ˆÁ7B‰ñB“7G&ñÊrÜ÷Wß¶ÚÊ‰ñB«¬""íÁFÙ∆˜vW$66RÇì∞¢&WGW&‚‰ñBbbÜ‰ñBÊñÊ6«VFW2ÜÊ˜&÷∆ó¶VBí«¬Ê˜&÷∆ó¶VBÊñÊ6«VFW2Ü‰ñBíì∞¢“ì∞¢&WGW&‚'î‰ñD6ˆÁFñÁ2«¬ÁV∆√∞ß–†¶7ñÊ2gVÊ7Fñˆ‚˜V‰gVV≈vRÜ÷Wß¶Ù∆&V¬í∞¢6V∆V7FVDgVVƒ÷Wß¶Ú“vWD÷Wß¶Ù'î∆&V¬Ü÷Wß¶Ù∆&V¬í«¬≤‰ñC¢÷Wß¶Ù∆&V¬¬Êˆ÷S¢÷Wß¶Ù∆&V¬”∞¢VíÊgVV≈vUFóF∆RÁFWáD6ˆÁFVÁB“Fó7G&ñ'WF˜&íÇÙT‰í(
"G∑6V∆V7FVDgVVƒ÷Wß¶ÚÊ‰ñB«¬6V∆V7FVDgVVƒ÷Wß¶ÚÊÊˆ÷R«¬$÷Wß¶Ú'÷∞¢VíÊgVVƒ÷Wß¶ÙFWFñ«46&BÊ6∆74∆ó7BÊFBÇ&ÜñFFV‚"ì∞¢ñbáVíÊgVV≈&FóW2íVíÊgVV≈&FóW2Áf«VR“#R#∞¢&VÊFW$gVVƒ÷Wß¶ÙFWFñ«2Çì∞¢vñÊF˜rÊ∆ˆ6Fñˆ‚ÊÜ6Ç“gVV√“G∂VÊ6ˆFUU$î6ˆ◊ˆÊVÁBá6V∆V7FVDgVVƒ÷Wß¶ÚÊ‰ñB«¬6V∆V7FVDgVVƒ÷Wß¶ÚÊÊˆ÷R«¬&÷Wß¶Ú"ó÷∞¢«ï&˜WFRÇì∞¢vóB∆ˆDÊV&'îgVV≈7FFñˆÁ2Çì∞ß–†¶gVÊ7Fñˆ‚Fˆvv∆TgVVƒ÷Wß¶ÙFWFñ«2Çí∞¢VíÊgVVƒ÷Wß¶ÙFWFñ«46&BÊ6∆74∆ó7BÁFˆvv∆RÇ&ÜñFFV‚"ì∞¢ñbÇVíÊgVVƒ÷Wß¶ÙFWFñ«46&BÊ6∆74∆ó7BÊ6ˆÁFñÁ2Ç&ÜñFFV‚"íí∞¢VíÊgVVƒ÷Wß¶ÙFWFñ«46&BÁ67&ˆ∆ƒñÁFıfñWrá≤&VÜfñ˜#¢'6÷ˆ˜FÇ"¬&∆ˆ6≥¢'7F'B"“ì∞¢–ß–†¶gVÊ7Fñˆ‚&VÊFW$gVVƒ÷Wß¶ÙFWFñ«2Çí∞¢ñbÇ6V∆V7FVDgVVƒ÷Wß¶Úí∞¢VíÊgVVƒ÷Wß¶ÙFWFñ«2ÊñÊÊW$ÖD‘¬“#«6∆73“v◊WFVBs‰ÊW77V‚÷Wß¶Ú6V∆W¶ñˆÊFÚ„¬˜‚#∞¢&WGW&„∞¢–¢6ˆÁ7B˜'FF∆&V¬“6V∆V7FVDgVVƒ÷Wß¶ÚÁ˜'FF6&ñ6Ú«¬6V∆V7FVDgVVƒ÷Wß¶ÚÁ˜'FF6&ñ6Ù∂r«¬6V∆V7FVDgVVƒ÷Wß¶ÚÁ˜'FF«¬"“#∞¢6ˆÁ7B÷76∆&V¬“6V∆V7FVDgVVƒ÷Wß¶ÚÊ÷766ˆ◊∆W76óf∂r«¬6V∆V7FVDgVVƒ÷Wß¶ÚÊ÷766ˆ◊∆W76óf«¬6V∆V7FVDgVVƒ÷Wß¶ÚÊ÷76«¬"“#∞¢6ˆÁ7B˜7Fî∆&V¬“vWD÷Wß¶ı˜7Fî∆&V¬á6V∆V7FVDgVVƒ÷Wß¶Úí«¬"“#∞¢VíÊgVVƒ÷Wß¶ÙFWFñ«2ÊñÊÊW$ÖD‘¬“ ¢«„∆#‰‚‚îC£¬ˆ#‚G∂W66TÖD‘¬á6V∆V7FVDgVVƒ÷Wß¶ÚÊ‰ñB«¬6V∆V7FVDgVVƒ÷Wß¶ÚÊÊˆ÷R«¬"“"ó”¬˜‡¢«„∆#‰÷&6£¬ˆ#‚G∂W66TÖD‘¬á6V∆V7FVDgVVƒ÷Wß¶ÚÊ÷&6«¬"“"ó”¬˜‡¢«„∆#‰÷ˆFV∆∆Û£¬ˆ#‚G∂W66TÖD‘¬á6V∆V7FVDgVVƒ÷Wß¶ÚÊ÷ˆFV∆∆Ú«¬"“"ó”¬˜‡¢«„∆#Â˜7Fì£¬ˆ#‚G∂W66TÖD‘¬á˜7Fî∆&V¬ó”¬˜‡¢«„∆#Â˜'FFÜ6&ñ6Úì£¬ˆ#‚G∂W66TÖD‘¬á˜'FF∆&V¬ó”¬˜‡¢«„∆#‰÷766ˆ◊∆W76ófÜ∂rì£¬ˆ#‚G∂W66TÖD‘¬Ü÷76∆&V¬ó”¬˜‡¢«„∆#‰∆ñ÷VÁF¶ñˆÊS£¬ˆ#‚G∂W66TÖD‘¬á6V∆V7FVDgVVƒ÷Wß¶ÚÊ∆ñ÷VÁF¶ñˆÊR«¬"“"ó”¬˜‡¢∆Fób6∆73“&óFV“÷7FñˆÁ2#‡¢∆'WGFˆ‚ñC“&gVV¬÷˜V‚◊ñ‚÷Fˆ2÷'F‚"6∆73“&'F‚"GóS“&'WGFˆ‚#Ô	˘8¬î‚6&'W&ÁFS¬ˆ'WGFˆ„‡¢¬ˆFóc‡¢∞¢6ˆÁ7B˜VÂñ‰'F‚“Fˆ7V÷VÁBÊvWDV∆V÷VÁD'îñBÇ&gVV¬÷˜V‚◊ñ‚÷Fˆ2÷'F‚"ì∞¢˜VÂñ‰'F„ÚÊFDWfVÁD∆ó7FVÊW"Ç&6∆ñ6≤"¬Çí”‚∞¢˜VÂ&ófFTFˆ75vRÇì∞¢«ï&ófFTFˆ5&W6WBÇ'ñ‚"ì∞¢6WEFñ÷V˜WBÇÇí”‚∞¢VíÁ&ófFTFˆ74f˜&”ÚÁ67&ˆ∆ƒñÁFıfñWrá≤&VÜfñ˜#¢'6÷ˆ˜FÇ"¬&∆ˆ6≥¢'7F'B"“ì∞¢VíÁ&ófFTFˆ74Ê÷SÚÊfˆ7W2Çì∞¢“¬Sì∞¢“ì∞ß–†¶7ñÊ2gVÊ7Fñˆ‚∆ˆDÊV&'îgVV≈7FFñˆÁ2Ü˜FñˆÁ2“∑“í∞¢ñbÜ˜FñˆÁ2Êf˜&6RbbgVV≈7FFñˆÁ4∆ˆE&ˆ÷ó6Rí∞¢gVV≈7FFñˆÁ4&˜'D6ˆÁG&ˆ∆∆W#ÚÊ&˜'BÇì∞¢G'í∞¢vóBgVV≈7FFñˆÁ4∆ˆE&ˆ÷ó6S∞¢“6F6ÇÜW'&˜"í∞¢ñbÜW'&˜#ÚÊÊ÷R”“$&˜'DW'&˜""í6ˆÁ6ˆ∆RÁv&‚Ç%&ñ6W&6Fó7G&ñ'WF˜&í&V6VFVÁFRÊˆ‚6ˆ◊∆WFF¢"¬W'&˜"ì∞¢–¢–¢ñbÜgVV≈7FFñˆÁ4∆ˆE&ˆ÷ó6Rí&WGW&‚gVV≈7FFñˆÁ4∆ˆE&ˆ÷ó6S∞¢gVV≈7FFñˆÁ4∆ˆE&ˆ÷ó6R“'V‰gVV≈7FFñˆÁ4∆ˆBÇíÊfñÊ∆«íÇÇí”‚≤gVV≈7FFñˆÁ4∆ˆE&ˆ÷ó6R“ÁV∆√≤“ì∞¢&WGW&‚gVV≈7FFñˆÁ4∆ˆE&ˆ÷ó6S∞ß–†¶7ñÊ2gVÊ7Fñˆ‚'V‰gVV≈7FFñˆÁ4∆ˆBÇí∞¢6ˆÁ7BgVV¬“vñÊF˜r‰ÜW&gVV≈7FFñˆÁ2ÊÊ˜&÷∆ó¶TgVV¬á6V∆V7FVDgVVƒ÷Wß¶ÛÚÊ∆ñ÷VÁF¶ñˆÊRì∞¢6ˆÁ7BgVVƒ∆&V¬“vñÊF˜r‰ÜW&gVV≈7FFñˆÁ2‰eTT≈Ùƒ$T≈5∂gVV≈“«¬7G&ñÊrá6V∆V7FVDgVVƒ÷Wß¶ÛÚÊ∆ñ÷VÁF¶ñˆÊR«¬&6&'W&ÁFR"íÁG&ñ“ÇíÁFÙ∆˜vW$66RÇì∞¢6ˆÁ7B&FóW5f«VR“ÁV÷&W"áVíÊgVV≈&FóW3ÚÁf«VR«¬Rì∞¢6ˆÁ7B&FóW4∂““÷FÇÊ÷ñ‚ÉS¬÷FÇÊ÷ÇÉR¬ÁV÷&W"Êó4fñÊóFRá&FóW5f«VRíÚ&FóW5f«VR¢Ríì∞¢6ˆÁ6ˆ∆RÊñÊfÚÇ%¥Fó7G&ñ'WF˜&ï“6&'W&ÁFRR&vvñÚ"¬≤gVV¬¬&FóW4∂““ì∞¢ñbáVíÊgVVƒfñ«FW%7V÷÷'íí∞¢VíÊgVVƒfñ«FW%7V÷÷'íÁFWáD6ˆÁFVÁB“gVV¿¢Ú6W&6Ú6ˆ∆ÚFó7G&ñ'WF˜&íÇÙT‰í6ÜRfVÊFˆÊÚG∂gVVƒ∆&V«“¬VÁG&ÚG∑&FóW4∂◊“∂“‚6ˆÁG&ˆ∆∆Ú&ñ÷Œ(	ñ&6ÜófñÚ‘î‘ïB6«fFÚÊ ¢¢∆ñ÷VÁF¶ñˆÊR÷Wß¶ÚÊˆ‚&ñ6ˆÊ˜66óWF(
"&vvñÚG∑&FóW4∂◊“∂÷∞¢–¢ñbÇgVV¬í∞¢6Ü˜tgVV≈7FFñˆÁ4W'&˜"Ç$∆ñ÷VÁF¶ñˆÊRFV¬÷Wß¶ÚÊˆ‚&ñ6ˆÊ˜66óWF‚"¬G'VRì∞¢&WGW&„∞¢–¢ñbÜÊfñvF˜"Êˆ‰∆ñÊR””“f«6Rí∞¢6Ü˜tgVV≈7FFñˆÁ4W'&˜"Ç$6ˆÊÊW76ñˆÊR76VÁFR‚6ˆÁG&ˆ∆∆ñÁFW&ÊWBR&ó&˜f‚"¬G'VRì∞¢&WGW&„∞¢–¢VíÊgVV≈7FFñˆÁ4∆ó7BÊñÊÊW$ÖD‘¬“#«6∆73“v◊WFVBs‰6&ñ6÷VÁFÚFó7G&ñ'WF˜&í6ˆ◊Fñ&ñ∆í‚‚„¬˜‚#∞¢ñbáVíÊgVV≈6V&6Ñ'F‚íVíÊgVV≈6V&6Ñ'F‚ÊFó6&∆VB“G'VS∞¢ñbáVíÊgVV≈&FóW2íVíÊgVV≈&FóW2ÊFó6&∆VB“G'VS∞¢VÁ7W&TgVVƒ÷Çì∞¢gVV≈7FFñˆÁ4∆ñW"Ê6∆V$∆ñW'2Çì∞¢G'í∞¢6ˆÁ7B˜6óFñˆ‚“vóB&WVW7Dg&W6ÑgVV≈˜6óFñˆ‚Çì∞¢WFFT7W'&VÁEW6W%˜6óFñˆ‚á˜6óFñˆ‚¬˜6óFñˆ‚ÁFñ÷W7F◊«¬FFRÊÊ˜rÇí¬≤&VÊFW#¢f«6R“ì∞¢6ˆÁ6ˆ∆RÊñÊfÚÇ%¥Fó7G&ñ'WF˜&ï“˜6ó¶ñˆÊR˜W&F˜&R"¬≤∆C¢˜6óFñˆ‚Ê∆B¬∆Ês¢˜6óFñˆ‚Ê∆Êr¬67W&7ì¢˜6óFñˆ‚Ê67W&7í“ì∞¢6ˆÁ7BñÊóFñ≈¶ˆˆ““&FóW4∂“√“RÚ"¢&FóW4∂“√“Ú¢&FóW4∂“√“#Ú¢ì∞¢gVVƒ÷ñÁ7FÊ6RÁ6WEfñWrÖ∑˜6óFñˆ‚Ê∆B¬˜6óFñˆ‚Ê∆Êu“¬ñÊóFñ≈¶ˆˆ“ì∞¢gVVƒ÷ñÁ7FÊ6RÊñÁf∆ñFFU6ó¶RÇì∞¢6ˆÁ7B6V&6Ö&W7V«B“vóBfWF6Ñ6ˆ◊Fñ&∆TgVV≈7FFñˆÁ2á˜6óFñˆ‚¬&FóW4∂“¬gVV¬ì∞¢6ˆÁ7B7FFñˆÁ2“6V&6Ö&W7V«BÁ7FFñˆÁ2Êfñ«FW"Çá7FFñˆ‚í”‚7FFñˆ‚ÊFó7FÊ6R√“&FóW4∂“ì∞¢6ˆÁ6ˆ∆RÊñÊfÚÇ%¥Fó7G&ñ'WF˜&ï“fñ«G&Ú&ó7V«FFí"¬≤gVV¬¬&FóW4∂“¬6˜W&6S¢6V&6Ö&W7V«BÁ6˜W&6R¬&V6VófVC¢6V&6Ö&W7V«BÁ&V6VófVB¬6ˆ◊Fñ&∆S¢7FFñˆÁ2Ê∆VÊwFÇ“ì∞¢ñbÇ7FFñˆÁ2Ê∆VÊwFÇí∞¢6Ü˜tgVV≈7FFñˆÁ4W'&˜"ÜÊW77V‚Fó7G&ñ'WF˜&RÇÙT‰í6ÜRfVÊFRG∂gVVƒ∆&V«“G&˜fFÚÊV¬&vvñÚFíG∑&FóW4∂◊“∂“Ê¬G'VRì∞¢&WGW&„∞¢–¢ñbáVíÊgVVƒfñ«FW%7V÷÷'íí∞¢VíÊgVVƒfñ«FW%7V÷÷'íÁFWáD6ˆÁFVÁB“G&˜fFíG∑7FFñˆÁ2Ê∆VÊwFá“Fó7G&ñ'WF˜&í6ÜRfVÊFˆÊÚG∂gVVƒ∆&V«“¬VÁG&ÚG∑&FóW4∂◊“∂“(
"fˆÁFS¢G∑6V&6Ö&W7V«BÁ6˜W&6W“Ê∞¢–¢&VÊFW$gVV≈7FFñˆÁ2á7FFñˆÁ2ì∞¢“6F6ÇÜW'&˜"í∞¢ñbÜW'&˜#ÚÊÊ÷R””“$&˜'DW'&˜""í&WGW&„∞¢6ˆÁ6ˆ∆RÊW'&˜"Ç%¥Fó7G&ñ'WF˜&ï“6&ñ6÷VÁFÚÊˆ‚&óW66óFÚ"¬≤GóS¢W'&˜#ÚÊgVVƒW'&˜%GóR«¬'6W'fñ6R"¬7FGW3¢W'&˜#ÚÁ7FGW2«¬ÁV∆¬¬÷W76vS¢W'&˜#ÚÊ÷W76vR“ì∞¢6ˆÁ7B÷W76vR“W'&˜#ÚÊgVVƒW'&˜%GóR””“&∆ˆ6Fñˆ‚ ¢Ú%˜6ó¶ñˆÊRÊˆ‚Fó7ˆÊñ&ñ∆R‚GFóf∆∆ˆ6∆óß¶¶ñˆÊRR&ó&˜f‚ ¢¢ÜÊfñvF˜"Êˆ‰∆ñÊR””“f«6R«¬W'&˜#ÚÊgVVƒW'&˜%GóR””“&ÊWGv˜&≤"ê¢Ú$6ˆÊÊW76ñˆÊR76VÁFR‚6ˆÁG&ˆ∆∆ñÁFW&ÊWBR&ó&˜f‚ ¢¢%6W'fó¶ñÚFó7G&ñ'WF˜&íFV◊˜&ÊV÷VÁFRÊˆ‚Fó7ˆÊñ&ñ∆R‚#∞¢6Ü˜tgVV≈7FFñˆÁ4W'&˜"Ü÷W76vR¬G'VRì∞¢“fñÊ∆«í∞¢ñbáVíÊgVV≈6V&6Ñ'F‚íVíÊgVV≈6V&6Ñ'F‚ÊFó6&∆VB“f«6S∞¢ñbáVíÊgVV≈&FóW2íVíÊgVV≈&FóW2ÊFó6&∆VB“f«6S∞¢–ß–†¶gVÊ7Fñˆ‚6Ü˜tgVV≈7FFñˆÁ4W'&˜"Ü÷W76vR¬&WG'íí∞¢ñbÜgVV≈7FFñˆÁ4∆ñW"ígVV≈7FFñˆÁ4∆ñW"Ê6∆V$∆ñW'2Çì∞¢VíÊgVV≈7FFñˆÁ4∆ó7BÊñÊÊW$ÖD‘¬“«6∆73“&◊WFVB#‚G∂W66TÖD‘¬Ü÷W76vRó”¬˜Ê∞¢ñbá&WG'ííVíÊgVV≈7FFñˆÁ4∆ó7BÊVÊD6Üñ∆BÜ7&VFT'WGFˆ‚Ç%&ó&˜f"¬Çí”‚∆ˆDÊV&'îgVV≈7FFñˆÁ2Çííì∞ß–†¶7ñÊ2gVÊ7Fñˆ‚&WVW7Dg&W6ÑgVV≈˜6óFñˆ‚Çí∞¢ñbÇÊfñvF˜"ÊvVˆ∆ˆ6Fñˆ‚íFá&˜rˆ&¶V7BÊ76ñv‚ÜÊWrW'&˜"Ç$vVˆ∆ˆ6Fñˆ‚VÁ7W˜'FVB"í¬≤gVVƒW'&˜%GóS¢&∆ˆ6Fñˆ‚"“ì∞¢G'í∞¢ñbÜÊfñvF˜"ÁW&÷ó76ñˆÁ3ÚÁVW'íí∞¢6ˆÁ7BW&÷ó76ñˆ‚“vóBÊfñvF˜"ÁW&÷ó76ñˆÁ2ÁVW'íá≤Ê÷S¢&vVˆ∆ˆ6Fñˆ‚"“ì∞¢ñbáW&÷ó76ñˆ‚Á7FFR””“&FVÊñVB"íFá&˜rˆ&¶V7BÊ76ñv‚ÜÊWrW'&˜"Ç$vVˆ∆ˆ6Fñˆ‚FVÊñVB"í¬≤gVVƒW'&˜%GóS¢&∆ˆ6Fñˆ‚"“ì∞¢–¢ñbávñÊF˜r‰66óF˜#ÚÊó4ÊFófU∆Ff˜&”Ú‚ÇíbbvñÊF˜r‰66óF˜#ÚÂ«VvñÁ3Ú‰vVˆ∆ˆ6Fñˆ„ÚÁ&WVW7EW&÷ó76ñˆÁ2í∞¢6ˆÁ7BW&÷ó76ñˆ‚“vóBvñÊF˜r‰66óF˜"Â«VvñÁ2‰vVˆ∆ˆ6Fñˆ‚Á&WVW7EW&÷ó76ñˆÁ2Çì∞¢ñbáW&÷ó76ñˆ„ÚÊ∆ˆ6Fñˆ‚””“&FVÊñVB"íFá&˜rˆ&¶V7BÊ76ñv‚ÜÊWrW'&˜"Ç$vVˆ∆ˆ6Fñˆ‚FVÊñVB"í¬≤gVVƒW'&˜%GóS¢&∆ˆ6Fñˆ‚"“ì∞¢–¢&WGW&‚vóBÊWr&ˆ÷ó6RÇá&W6ˆ«fR¬&V¶V7Bí”‚ÊfñvF˜"ÊvVˆ∆ˆ6Fñˆ‚ÊvWD7W'&VÁE˜6óFñˆ‚Ä¢á˜2í”‚&W6ˆ«fRá≤∆C¢˜2Ê6ˆ˜&G2Ê∆FóGVFR¬∆Ês¢˜2Ê6ˆ˜&G2Ê∆ˆÊvóGVFR¬67W&7ì¢˜2Ê6ˆ˜&G2Ê67W&7í«¬“í¿¢ÜW'&˜"í”‚&V¶V7BÑˆ&¶V7BÊ76ñv‚ÜW'&˜"«¬ÊWrW'&˜"Ç$vVˆ∆ˆ6Fñˆ‚VÊfñ∆&∆R"í¬≤gVVƒW'&˜%GóS¢&∆ˆ6Fñˆ‚"“íí¿¢≤VÊ&∆TÜñvÑ67W&7ì¢G'VR¬Fñ÷V˜WC¢#¬÷Üñ◊V‘vS¢–¢íì∞¢“6F6ÇÜW'&˜"í∞¢W'&˜"ÊgVVƒW'&˜%GóR“&∆ˆ6Fñˆ‚#∞¢Fá&˜rW'&˜#∞¢–ß–†¶7ñÊ2gVÊ7Fñˆ‚fWF6Ñ6ˆ◊Fñ&∆TgVV≈7FFñˆÁ2á˜6óFñˆ‚¬&FóW4∂“¬gVV¬í∞¢ñbÜgVV¬”“&V∆V7G&ñ2"í∞¢6ˆÁ7BÊFñˆÊƒ66ÜR“vñÊF˜r‰ÜW&gVVƒÊFñˆÊƒ66ÜS∞¢ñbÜÊFñˆÊƒ66ÜSÚÊfñÊDÊV&'íí∞¢6ˆÁ7B66ÜVB“vóBÊFñˆÊƒ66ÜRÊfñÊDÊV&'íÜgVV¬¬˜6óFñˆ‚¬&FóW4∂“¬ÜfW'6ñÊRì∞¢ñbÜ66ÜVBÁ7FFñˆÁ2Ê∆VÊwFÇí∞¢6ˆÁ6ˆ∆RÊñÊfÚÇ%¥Fó7G&ñ'WF˜&ï“&ó7V«FFíF∆Œ(	ñ&6ÜófñÚ‘î‘ïBÊ¶ñˆÊ∆R"¬∞¢gVV¬¿¢&FóW4∂“¿¢F˜Fƒ66ÜVC¢66ÜVBÁF˜Fƒ66ÜVB¿¢6ˆ◊Fñ&∆S¢66ÜVBÁ7FFñˆÁ2Ê∆VÊwFÄ¢“ì∞¢&WGW&‚∞¢7FFñˆÁ3¢66ÜVBÁ7FFñˆÁ2¿¢6˜W&6S¢$&6ÜófñÚ‘î‘ïB6«fFÚ"¿¢&V6VófVC¢66ÜVBÁF˜Fƒ66ÜV@¢”∞¢–¢ñbÇ66ÜVBÊfñ∆&∆RíÊFñˆÊƒ66ÜRÁ&Vg&W6ÉÚ‚ÇíÊ6F6ÇÇÇí”‚∑“ì∞¢–†¢G'í∞¢6ˆÁ7BFF“vóBfWF6ÑgVV≈7FFñˆÁ4g&ˆ‘÷ñ÷óBá˜6óFñˆ‚Ê∆B¬˜6óFñˆ‚Ê∆Êr¬&FóW4∂“ì∞¢6ˆÁ7B7FFñˆÁ2“vñÊF˜r‰ÜW&gVV≈7FFñˆÁ2Á'6T÷ñ÷óE7FFñˆÁ2ÜFFÁ&W7V«G2¬gVV¬¬˜6óFñˆ‚¬ÜfW'6ñÊRê¢Êfñ«FW"Çá7FFñˆ‚í”‚7FFñˆ‚ÊFó7FÊ6R√“&FóW4∂“ì∞¢ñbá7FFñˆÁ2Ê∆VÊwFÇí∞¢&WGW&‚≤7FFñˆÁ2¬6˜W&6S¢$‘î‘ïB˜76W'f&Wß¶í"¬&V6VófVC¢FFÁ&W7V«G2Ê∆VÊwFÇ”∞¢–¢6ˆÁ6ˆ∆RÊñÊfÚÇ%¥Fó7G&ñ'WF˜&ï“‘î‘ïB6VÁ¶&ó7V«FFí6ˆ◊Fñ&ñ∆í¬W6Ú∆&ó6W'f˜VÂ7G&VWD÷"¬≤gVV¬¬&FóW4∂“¬&V6VófVC¢FFÁ&W7V«G2Ê∆VÊwFÇ“ì∞¢“6F6ÇÜW'&˜"í∞¢ñbÜW'&˜#ÚÊÊ÷R””“$&˜'DW'&˜""íFá&˜rW'&˜#∞¢6ˆÁ6ˆ∆RÁv&‚Ç%¥Fó7G&ñ'WF˜&ï“‘î‘ïBÊˆ‚Fó7ˆÊñ&ñ∆R¬W6Ú∆&ó6W'f˜VÂ7G&VWD÷"¬≤7FGW3¢W'&˜#ÚÁ7FGW2«¬ÁV∆¬¬÷W76vS¢W'&˜#ÚÊ÷W76vR“ì∞¢–¢–†¢6ˆÁ7BFF“vóBfWF6ÑgVV≈7FFñˆÁ4g&ˆ‘˜fW'72á˜6óFñˆ‚Ê∆B¬˜6óFñˆ‚Ê∆Êr¬&FóW4∂“¬gVV¬ì∞¢6ˆÁ7B7FFñˆÁ2“vñÊF˜r‰ÜW&gVV≈7FFñˆÁ2Á'6U7FFñˆÁ2ÜFFÊV∆V÷VÁG2¬gVV¬¬˜6óFñˆ‚¬ÜfW'6ñÊRê¢Êfñ«FW"Çá7FFñˆ‚í”‚7FFñˆ‚ÊFó7FÊ6R√“&FóW4∂“ê¢Ê÷Çá7FFñˆ‚í”‚á≤‚‚Á7FFñˆ‚¬6˜W&6S¢$˜VÂ7G&VWD÷"“íì∞¢&WGW&‚≤7FFñˆÁ2¬6˜W&6S¢$˜VÂ7G&VWD÷"¬&V6VófVC¢FFÊV∆V÷VÁG3ÚÊ∆VÊwFÇ«¬”∞ß–†¶7ñÊ2gVÊ7Fñˆ‚fWF6ÑgVV≈7FFñˆÁ4g&ˆ‘÷ñ÷óBÜ∆B¬∆Êr¬&FóW4∂“í∞¢∆WBFñ÷V˜WC∞¢∆WBFñ÷VD˜WB“f«6S∞¢G'í∞¢gVV≈7FFñˆÁ4&˜'D6ˆÁG&ˆ∆∆W"“ÊWr&˜'D6ˆÁG&ˆ∆∆W"Çì∞¢Fñ÷V˜WB“6WEFñ÷V˜WBÇÇí”‚∞¢Fñ÷VD˜WB“G'VS∞¢gVV≈7FFñˆÁ4&˜'D6ˆÁG&ˆ∆∆W"Ê&˜'BÇì∞¢“¬#ì∞¢6ˆÁ7B&W7ˆÁ6R“vóBfWF6ÇávñÊF˜r‰ÜW&gVV≈7FFñˆÁ2‰‘î‘ïEÙïıU$¬¬∞¢÷WFÜˆC¢%ı5B"¿¢ÜVFW'3¢≤66WC¢&∆ñ6Fñˆ‚ˆß6ˆ‚"¬$6ˆÁFVÁB’GóR#¢&∆ñ6Fñˆ‚ˆß6ˆ‚"“¿¢&ˆGì¢•4Ù‚Á7G&ñÊvñgíávñÊF˜r‰ÜW&gVV≈7FFñˆÁ2Ê'Vñ∆D÷ñ÷óE&WVW7BÜ∆B¬∆Êr¬&FóW4∂“íí¿¢6ñvÊ√¢gVV≈7FFñˆÁ4&˜'D6ˆÁG&ˆ∆∆W"Á6ñvÊ¿¢“ì∞¢6∆V%Fñ÷V˜WBáFñ÷V˜WBì∞¢6ˆÁ6ˆ∆RÊñÊfÚÇ%¥Fó7G&ñ'WF˜&ï“&ó7˜7F‘î‘ïB"¬≤7FGW3¢&W7ˆÁ6RÁ7FGW2¬ˆ≥¢&W7ˆÁ6RÊˆ≤¬&FóW4∂““ì∞¢ñbÇ&W7ˆÁ6RÊˆ≤íFá&˜rˆ&¶V7BÊ76ñv‚ÜÊWrW'&˜"Ü‘î‘ïBÖEEG∑&W7ˆÁ6RÁ7FGW7÷í¬≤7FGW3¢&W7ˆÁ6RÁ7FGW2“ì∞¢6ˆÁ7BFF“vóB&W7ˆÁ6RÊß6ˆ‚Çì∞¢ñbÜFFÚÁ7V66W72””“f«6R«¬'&íÊó4'&íÜFFÚÁ&W7V«G2ííFá&˜rÊWrW'&˜"Ç%&ó7˜7F‘î‘ïBÊˆ‚f∆ñF"ì∞¢&WGW&‚FF∞¢“6F6ÇÜW'&˜"í∞¢6∆V%Fñ÷V˜WBáFñ÷V˜WBì∞¢ñbÜW'&˜#ÚÊÊ÷R””“$&˜'DW'&˜""bbFñ÷VD˜WBíFá&˜rˆ&¶V7BÊ76ñv‚ÜÊWrW'&˜"Ç%Fñ÷V˜WB‘î‘ïB"í¬≤7FGW3¢CÇ“ì∞¢Fá&˜rW'&˜#∞¢–ß–†¶7ñÊ2gVÊ7Fñˆ‚fWF6ÑgVV≈7FFñˆÁ4g&ˆ‘˜fW'72Ü∆B¬∆Êr¬&FóW4∂“¬gVV¬í∞¢6ˆÁ7BVW'í“vñÊF˜r‰ÜW&gVV≈7FFñˆÁ2Ê'Vñ∆EVW'íÜ∆B¬∆Êr¬&FóW4∂“¬gVV¬ì∞¢6ˆÁ7BVÊGˆñÁG2“∞¢&áGG3¢Úˆ˜fW'72÷íÊFRˆíˆñÁFW'&WFW""¿¢&áGG3¢Úˆ˜fW'72Ê∑V÷íÁ7ó7FV◊2ˆíˆñÁFW'&WFW" ¢”∞†¢∆WB∆7DW'&˜"“ÁV∆√∞¢f˜"Ü6ˆÁ7BVÊGˆñÁBˆbVÊGˆñÁG2í∞¢∆WBFñ÷V˜WC∞¢G'í∞¢gVV≈7FFñˆÁ4&˜'D6ˆÁG&ˆ∆∆W"“ÊWr&˜'D6ˆÁG&ˆ∆∆W"Çì∞¢Fñ÷V˜WB“6WEFñ÷V˜WBÇÇí”‚gVV≈7FFñˆÁ4&˜'D6ˆÁG&ˆ∆∆W"Ê&˜'BÇí¬#ì∞¢6ˆÁ7B&WVW7EW&¬“G∂VÊGˆñÁG”ˆFF“G∂VÊ6ˆFUU$î6ˆ◊ˆÊVÁBáVW'íó÷∞¢6ˆÁ6ˆ∆RÊñÊfÚÇ%¥Fó7G&ñ'WF˜&ï“&ñ6ÜñW7F"¬≤W&√¢VÊGˆñÁB¬∆B¬∆Êr¬&FóW4∂“¬gVV¬“ì∞¢6ˆÁ7B&W7ˆÁ6R“vóBfWF6Çá&WVW7EW&¬¬≤÷WFÜˆC¢$tUB"¬ÜVFW'3¢≤66WC¢&∆ñ6Fñˆ‚ˆß6ˆ‚"“¬6ñvÊ√¢gVV≈7FFñˆÁ4&˜'D6ˆÁG&ˆ∆∆W"Á6ñvÊ¬“ì∞¢6∆V%Fñ÷V˜WBáFñ÷V˜WBì∞¢6ˆÁ6ˆ∆RÊñÊfÚÇ%¥Fó7G&ñ'WF˜&ï“&ó7˜7FÖEE"¬≤W&√¢VÊGˆñÁB¬7FGW3¢&W7ˆÁ6RÁ7FGW2¬ˆ≥¢&W7ˆÁ6RÊˆ≤“ì∞¢ñbÇ&W7ˆÁ6RÊˆ≤íFá&˜rˆ&¶V7BÊ76ñv‚ÜÊWrW'&˜"Ü˜fW'72ÖEEG∑&W7ˆÁ6RÁ7FGW7÷í¬≤7FGW3¢&W7ˆÁ6RÁ7FGW2“ì∞¢6ˆÁ7BFF“vóB&W7ˆÁ6RÊß6ˆ‚Çì∞¢ñbÇ'&íÊó4'&íÜFFÚÊV∆V÷VÁG2ííFá&˜rÊWrW'&˜"Ç%&ó7˜7F˜fW'72Êˆ‚f∆ñF"ì∞¢&WGW&‚FF∞¢“6F6ÇÜW'&˜"í∞¢6∆V%Fñ÷V˜WBáFñ÷V˜WBì∞¢ñbÜW'&˜#ÚÊÊ÷R””“$&˜'DW'&˜""íW'&˜"“ˆ&¶V7BÊ76ñv‚ÜÊWrW'&˜"Ç%Fñ÷V˜WB˜fW'72"í¬≤7FGW3¢CÇ“ì∞¢ñbÜW'&˜"ñÁ7FÊ6VˆbGóTW'&˜"íW'&˜"ÊgVVƒW'&˜%GóR“&ÊWGv˜&≤#∞¢∆7DW'&˜"“W'&˜#∞¢–¢–¢Fá&˜r∆7DW'&˜"«¬ÊWrW'&˜"Ç$˜fW'72Êˆ‚Fó7ˆÊñ&ñ∆R"ì∞ß–†¶gVÊ7Fñˆ‚VÁ7W&TgVVƒ÷Çí∞¢ñbÜgVVƒ÷ñÁ7FÊ6Rí&WGW&„∞¢gVVƒ÷ñÁ7FÊ6R“¬Ê÷Ç&gVV¬÷÷"ì∞¢¬ÁFñ∆T∆ñW"Ç&áGG3¢Ú˜∑7“ÁFñ∆RÊ˜VÁ7G&VWF÷Ê˜&r˜∑ß“˜∑á“˜∑ó“ÁÊr"¬∞¢÷Ö¶ˆˆ”¢í¿¢GG&ñ'WFñˆ„¢"f6˜ì≤˜VÂ7G&VWD÷ ¢“íÊFEFÚÜgVVƒ÷ñÁ7FÊ6Rì∞¢gVV≈7FFñˆÁ4∆ñW"“¬Ê∆ñW$w&˜WÇíÊFEFÚÜgVVƒ÷ñÁ7FÊ6Rì∞ß–†¶gVÊ7Fñˆ‚&VÊFW$gVV≈7FFñˆÁ2á7FFñˆÁ2í∞¢VÁ7W&TgVVƒ÷Çì∞¢gVV≈7FFñˆÁ4∆ñW"Ê6∆V$∆ñW'2Çì∞¢VíÊgVV≈7FFñˆÁ4∆ó7BÊñÊÊW$ÖD‘¬“"#∞¢ñbÇ7FFñˆÁ2Ê∆VÊwFÇí∞¢&WGW&„∞¢–¢6ˆÁ7B&˜VÊG2“µ”∞¢7FFñˆÁ2Êf˜$V6ÇÇá7FFñˆ‚í”‚∞¢6ˆÁ7B÷&∂W"“¬Ê÷&∂W"Ö∑7FFñˆ‚Ê∆B¬7FFñˆ‚Ê∆ˆÂ“¬∞¢ñ6ˆ„¢7&VFTgVVƒ÷&∂W$ñ6ˆ‚á7FFñˆ‚Ê'&ÊD∆&V¬ê¢“íÊFEFÚÜgVV≈7FFñˆÁ4∆ñW"ì∞¢6ˆÁ7B6˜W&6T∆&V¬“7FFñˆ‚Á6˜W&6RÚ∆'#‰fˆÁFS¢G∂W66TÖD‘¬á7FFñˆ‚Á6˜W&6Ró÷¢"#∞¢÷&∂W"Ê&ñÊE˜WÜ∆#‚G∂W66TÖD‘¬á7FFñˆ‚ÊÊ÷Ró”¬ˆ#„∆'#‚G∂W66TÖD‘¬á7FFñˆ‚ÊFG&W72ó”∆'#‚G∂W66TÖD‘¬á7FFñˆ‚Êfñ∆&∆TgVV¬ó“(
"G∂f˜&÷DFó7FÊ6Rá7FFñˆ‚ÊFó7FÊ6Ró“G∑6˜W&6T∆&V«÷ì∞¢6ˆÁ7BÊd'F‚“7&VFT'WGFˆ‚Ç$‰dît"¬Çí”‚vñÊF˜rÊ˜V‚ÜáGG3¢Ú˜wwrÊvˆˆv∆RÊ6ˆ“ˆ÷2ˆFó"Ûˆì”fFW7FñÊFñˆ„“G∂VÊ6ˆFUU$î6ˆ◊ˆÊVÁBÜG∑7FFñˆ‚Ê∆G“¬G∑7FFñˆ‚Ê∆ˆÁ÷ó÷¬%ˆ&∆Ê≤"¬&Êˆ˜VÊW""íì∞¢6ˆÁ7B&˜r“Fˆ7V÷VÁBÊ7&VFTV∆V÷VÁBÇ&Fób"ì∞¢&˜rÊ6∆74Ê÷R“'6ñ◊∆R÷∆ó7B÷óFV“#∞¢&˜rÊñÊÊW$ÖD‘¬“«7„„∆#‚G∂W66TÖD‘¬á7FFñˆ‚ÊÊ÷Ró”¬ˆ#„∆'#„«6÷∆√‚G∂W66TÖD‘¬á7FFñˆ‚ÊFG&W72ó”∆'#‚G∂W66TÖD‘¬á7FFñˆ‚Êfñ∆&∆TgVV¬ó“(
"G∂f˜&÷DFó7FÊ6Rá7FFñˆ‚ÊFó7FÊ6Ró“G∑6˜W&6T∆&V«”¬˜6÷∆√„¬˜7„Ê∞¢&˜rÊVÊD6Üñ∆BÜÊd'F‚ì∞¢VíÊgVV≈7FFñˆÁ4∆ó7BÊVÊD6Üñ∆Bá&˜rì∞¢÷&∂W"Êˆ‚Ç&6∆ñ6≤"¬Çí”‚Êd'F‚Êfˆ7W2Çíì∞¢&˜VÊG2ÁW6ÇÖ∑7FFñˆ‚Ê∆B¬7FFñˆ‚Ê∆ˆÂ“ì∞¢“ì∞¢ñbÜ7W'&VÁEW6W%˜2í&˜VÊG2ÁW6ÇÖ∂7W'&VÁEW6W%˜2Ê∆B¬7W'&VÁEW6W%˜2Ê∆Êu“ì∞¢gVVƒ÷ñÁ7FÊ6RÊfóD&˜VÊG2Ü&˜VÊG2¬≤FFñÊs¢≥#B¬#E““ì∞ß–†¶gVÊ7Fñˆ‚7&VFTgVVƒ÷&∂W$ñ6ˆ‚Ü'&ÊD∆&V¬í∞¢6ˆÁ7B÷&∂W$6∆72“vWDgVVƒ÷&∂W$6∆72Ü'&ÊD∆&V¬ì∞¢&WGW&‚¬ÊFódñ6ˆ‚á∞¢6∆74Ê÷S¢&gVV¬÷÷&∂W"◊w&"¿¢áF÷√¢«7‚6∆73“&gVV¬÷÷&∂W"÷∆&V¬G∂÷&∂W$6∆77“#‚G∂W66TÖD‘¬Ü'&ÊD∆&V¬«¬$eTT¬"ó”¬˜7„Ê¿¢ñ6ˆÂ6ó¶S¢≥CB¬#E“¿¢ñ6ˆ‰Ê6Ü˜#¢≥#"¬%“¿¢˜WÊ6Ü˜#¢≥¬”–¢“ì∞ß–†¶gVÊ7Fñˆ‚vWDgVVƒ÷&∂W$6∆72Ü'&ÊD∆&V¬í∞¢6ˆÁ7BÊ˜&÷∆ó¶VB“7G&ñÊrÜ'&ÊD∆&V¬«¬""íÁFÙ∆˜vW$66RÇì∞¢ñbÜÊ˜&÷∆ó¶VBÊñÊ6«VFW2Ç'Ç"íí&WGW&‚&gVV¬÷÷&∂W"◊Ç#∞¢ñbÜÊ˜&÷∆ó¶VBÊñÊ6«VFW2Ç&VÊí"íí&WGW&‚&gVV¬÷÷&∂W"÷VÊí#∞¢&WGW&‚&gVV¬÷÷&∂W"÷FVfV«B#∞ß–†¶gVÊ7Fñˆ‚ˆÂW'6ˆÊ≈6W'fñ6T6FVv˜'î6∆ñ6≤ÜWfVÁBí∞¢6ˆÁ7B'F‚“WfVÁBÁF&vWBÊ6∆˜6W7BÇ"ÁW'6ˆÊ¬◊6W'fñ6R÷6FVv˜'í÷'F‚"ì∞¢ñbÇ'F‚í&WGW&„∞¢6ˆÁ7B6FVv˜'í“'F‚ÊFF6WBÁ6W'fñ6T6FVv˜'í«¬"#∞¢ñbÇ6FVv˜'íí&WGW&„∞¢vñÊF˜rÊ∆ˆ6Fñˆ‚ÊÜ6Ç“6W'fó¶í◊W'6ˆÊ∆ì“G∂6FVv˜'ó÷∞¢«ï&˜WFRÇì∞ß–†¶7ñÊ2gVÊ7Fñˆ‚∆ˆEW'6ˆÊ≈6W'fñ6W4'î6FVv˜'íÜ6FVv˜'íí∞¢ñbÇU%4Ù‰≈ı4U%dî4UÙ4DTtı$îU5∂6FVv˜'ï“í&WGW&„∞¢7FófUW'6ˆÊ≈6W'fñ6T6FVv˜'í“6FVv˜'ì∞¢WáÊFVEW'6ˆÊ≈6W'fñ6TñB“"#∞¢W'6ˆÊ≈6W'fñ6W5&W7V«G2“µ”∞¢VíÁW'6ˆÊ≈6W'fñ6W46FVv˜&ñW3ÚÁVW'ï6V∆V7F˜$∆¬Ç"ÁW'6ˆÊ¬◊6W'fñ6R÷6FVv˜'í÷'F‚"íÊf˜$V6ÇÇÜ'F‚í”‚∞¢'F‚Ê6∆74∆ó7BÁFˆvv∆RÇ&ó2÷7FófR"¬'F‚ÊFF6WBÁ6W'fñ6T6FVv˜'í””“6FVv˜'íì∞¢“ì∞¢6ˆÁ7B6fr“U%4Ù‰≈ı4U%dî4UÙ4DTtı$îU5∂6FVv˜'ï”∞¢6ˆÁ7B&FóW4÷WFW'2“vWE6V∆V7FVEW'6ˆÊ≈6W'fñ6W5&FóW2Çì∞¢VíÁW'6ˆÊ≈6W'fñ6W5vUFóF∆RÁFWáD6ˆÁFVÁB“G∂6frÊñ6ˆÁ“G∂6frÁFóF∆W÷∞¢VíÁW'6ˆÊ≈6W'fñ6W4∆ó7EFóF∆RÁFWáD6ˆÁFVÁB“ú;ífñ6ñÊíFR(
"G∂6frÁFóF∆W“(
"&vvñÚG¥÷FÇÁ&˜VÊBá&FóW4÷WFW'2Úó“∂÷∞¢ñbÇ7W'&VÁEW6W%˜2í∞¢VíÁW'6ˆÊ≈6W'fñ6W4fVVF&6≤ÁFWáD6ˆÁFVÁB“%˜6ó¶ñˆÊRÊˆ‚Fó7ˆÊñ&ñ∆R‚GFófu2W"W6&Rí6W'fó¶íW'6ˆÊ∆í‚#∞¢VíÁW'6ˆÊ≈6W'fñ6W4∆ó7BÊñÊÊW$ÖD‘¬“"#∞¢6∆V%W'6ˆÊ≈6W'fñ6W4÷Çì∞¢&WGW&„∞¢–¢VíÁW'6ˆÊ≈6W'fñ6W4fVVF&6≤ÁFWáD6ˆÁFVÁB“$6&ñ6÷VÁFÚ«VˆvÜíñ‚6˜'6Ú‚‚‚#∞¢VíÁW'6ˆÊ≈6W'fñ6W4∆ó7BÊñÊÊW$ÖD‘¬“"#∞¢G'í∞¢6ˆÁ7BFF“vóBfWF6ÖW'6ˆÊ≈6W'fñ6W4g&ˆ‘˜fW'72Ü6FVv˜'í¬7W'&VÁEW6W%˜2Ê∆B¬7W'&VÁEW6W%˜2Ê∆Êr¬&FóW4÷WFW'2ì∞¢6ˆÁ7B∆6W2“Ê˜&÷∆ó¶UW'6ˆÊ≈6W'fñ6W2ÜFFÊV∆V÷VÁG2«¬µ“¬6FVv˜'íì∞¢W'6ˆÊ≈6W'fñ6W5&W7V«G2“∆6W3∞¢&VÊFW%W'6ˆÊ≈6W'fñ6W4∆ó7BÇì∞¢&VÊFW%W'6ˆÊ≈6W'fñ6W4÷Çì∞¢ñbÇ∆6W2Ê∆VÊwFÇí∞¢VíÁW'6ˆÊ≈6W'fñ6W4fVVF&6≤ÁFWáD6ˆÁFVÁB“$ÊW77V‚&ó7V«FFÚG&˜fFÚÊV∆∆¶ˆÊ‚#∞¢“V«6RñbÜ6FVv˜'í””“&«VÊ6Ç"í∞¢6ˆÁ7B66WFVD6˜VÁB“∆6W2Êfñ«FW"Çá∆6Rí”‚ó4÷V≈f˜V6ÜW$66WFVBá∆6RÁFw2ííÊ∆VÊwFÉ∞¢VíÁW'6ˆÊ≈6W'fñ6W4fVVF&6≤ÁFWáD6ˆÁFVÁB“G&˜fFíG∑∆6W2Ê∆VÊwFá“«VˆvÜíÇG∂66WFVD6˜VÁG“6ˆ‚'VˆÊí7FÚíÊ∞¢“V«6R∞¢VíÁW'6ˆÊ≈6W'fñ6W4fVVF&6≤ÁFWáD6ˆÁFVÁB“G&˜fFíG∑∆6W2Ê∆VÊwFá“«VˆvÜífñ6ñÊÚFRÊ∞¢–¢“6F6ÇÜW'&˜"í∞¢6ˆÁ6ˆ∆RÊW'&˜"Ç$W'&˜&R6&ñ6÷VÁFÚ6W'fó¶íW'6ˆÊ∆ì¢"¬W'&˜"ì∞¢VíÁW'6ˆÊ≈6W'fñ6W4fVVF&6≤ÁFWáD6ˆÁFVÁB“$W'&˜&RGW&ÁFRñ¬6&ñ6÷VÁFÚ‚&ó&˜f‚#∞¢VíÁW'6ˆÊ≈6W'fñ6W4∆ó7BÊñÊÊW$ÖD‘¬“"#∞¢6∆V%W'6ˆÊ≈6W'fñ6W4÷Çì∞¢–ß–†¶gVÊ7Fñˆ‚Ê˜&÷∆ó¶UW'6ˆÊ≈6W'fñ6W2ÜóFV◊2¬6FVv˜'íí∞¢6ˆÁ7B6VV‚“ÊWr6WBÇì∞¢&WGW&‚óFV◊2Ê÷ÇÜóFV“í”‚∞¢6ˆÁ7B∆B“óFV“Ê∆B«¬ÜóFV“Ê6VÁFW"bbóFV“Ê6VÁFW"Ê∆Bì∞¢6ˆÁ7B∆ˆ‚“óFV“Ê∆ˆ‚«¬ÜóFV“Ê6VÁFW"bbóFV“Ê6VÁFW"Ê∆ˆ‚ì∞¢ñbÇ∆B«¬∆ˆ‚í&WGW&‚ÁV∆√∞¢6ˆÁ7BFw2“óFV“ÁFw2«¬∑”∞¢6ˆÁ7BÊ÷R“Fw2ÊÊ÷R«¬Fw2Ê'&ÊB«¬FVfV«EW'6ˆÊ≈6W'fñ6TÊ÷RÜ6FVv˜'íì∞¢6ˆÁ7B∂Wí“G∂Ê÷RÁFÙ∆˜vW$66RÇó““G¥÷FÇÁ&˜VÊBÜ∆B¢ó““G¥÷FÇÁ&˜VÊBÜ∆ˆ‚¢ó÷∞¢ñbá6VV‚ÊÜ2Ü∂Wííí&WGW&‚ÁV∆√∞¢6VV‚ÊFBÜ∂Wíì∞¢&WGW&‚∞¢ñC¢óFV“ÊñB«¬∂Wí¿¢6FVv˜'í¿¢Ê÷R¿¢∆B¿¢∆ˆ‚¿¢Fw2¿¢Fó7FÊ6S¢ÜfW'6ñÊRÜ7W'&VÁEW6W%˜2Ê∆B¬7W'&VÁEW6W%˜2Ê∆Êr¬∆B¬∆ˆ‚ê¢”∞¢“íÊfñ«FW"Ñ&ˆˆ∆V‚íÁ6˜'BÇÜ¬"í”‚ÊFó7FÊ6R“"ÊFó7FÊ6Rì∞ß–†¶gVÊ7Fñˆ‚FVfV«EW'6ˆÊ≈6W'fñ6TÊ÷RÜ6FVv˜'íí∞¢6ˆÁ7B6fr“U%4Ù‰≈ı4U%dî4UÙ4DTtı$îU5∂6FVv˜'ï”∞¢&WGW&‚6frÚ6frÁFóF∆R¢%6W'fó¶ñÚ#∞ß–†¶gVÊ7Fñˆ‚VÁ7W&UW'6ˆÊ≈6W'fñ6W4÷Çí∞¢ñbáW'6ˆÊ≈6W'fñ6W4÷ñÁ7FÊ6Rí&WGW&„∞¢W'6ˆÊ≈6W'fñ6W4÷ñÁ7FÊ6R“¬Ê÷Ç'W'6ˆÊ¬◊6W'fñ6W2÷÷"ì∞¢¬ÁFñ∆T∆ñW"Ç&áGG3¢Ú˜∑7“ÁFñ∆RÊ˜VÁ7G&VWF÷Ê˜&r˜∑ß“˜∑á“˜∑ó“ÁÊr"¬∞¢÷Ö¶ˆˆ”¢í¿¢GG&ñ'WFñˆ„¢"f6˜ì≤˜VÂ7G&VWD÷ ¢“íÊFEFÚáW'6ˆÊ≈6W'fñ6W4÷ñÁ7FÊ6Rì∞¢W'6ˆÊ≈6W'fñ6W4∆ñW"“¬Ê∆ñW$w&˜WÇíÊFEFÚáW'6ˆÊ≈6W'fñ6W4÷ñÁ7FÊ6Rì∞ß–†¶gVÊ7Fñˆ‚6∆V%W'6ˆÊ≈6W'fñ6W4÷Çí∞¢ñbÇW'6ˆÊ≈6W'fñ6W4∆ñW"í&WGW&„∞¢W'6ˆÊ≈6W'fñ6W4∆ñW"Ê6∆V$∆ñW'2Çì∞ß–†¶gVÊ7Fñˆ‚&VÊFW%W'6ˆÊ≈6W'fñ6W4÷Çí∞¢VÁ7W&UW'6ˆÊ≈6W'fñ6W4÷Çì∞¢6∆V%W'6ˆÊ≈6W'fñ6W4÷Çì∞¢ñbÇW'6ˆÊ≈6W'fñ6W5&W7V«G2Ê∆VÊwFÇí&WGW&„∞¢6ˆÁ7B&˜VÊG2“µ”∞¢W'6ˆÊ≈6W'fñ6W5&W7V«G2Êf˜$V6ÇÇá∆6Rí”‚∞¢6ˆÁ7B÷&∂W"“¬Ê÷&∂W"Ö∑∆6RÊ∆B¬∆6RÊ∆ˆÂ“¬∞¢ñ6ˆ„¢7&VFUW'6ˆÊ≈6W'fñ6T÷&∂W$ñ6ˆ‚á∆6RÊ6FVv˜'íê¢“íÊFEFÚáW'6ˆÊ≈6W'fñ6W4∆ñW"ì∞¢÷&∂W"Ê&ñÊE˜WÜ∆#‚G∂W66TÖD‘¬á∆6RÊÊ÷Ró”¬ˆ#„∆'#‚G∂f˜&÷DFó7FÊ6Rá∆6RÊFó7FÊ6Ró÷ì∞¢÷&∂W"Êˆ‚Ç&6∆ñ6≤"¬Çí”‚6V∆V7EW'6ˆÊ≈6W'fñ6Rá∆6RÊñBíì∞¢&˜VÊG2ÁW6ÇÖ∑∆6RÊ∆B¬∆6RÊ∆ˆÂ“ì∞¢“ì∞¢ñbÜ7W'&VÁEW6W%˜2í&˜VÊG2ÁW6ÇÖ∂7W'&VÁEW6W%˜2Ê∆B¬7W'&VÁEW6W%˜2Ê∆Êu“ì∞¢W'6ˆÊ≈6W'fñ6W4÷ñÁ7FÊ6RÊfóD&˜VÊG2Ü&˜VÊG2¬≤FFñÊs¢≥#B¬#E““ì∞ß–†¶gVÊ7Fñˆ‚7&VFUW'6ˆÊ≈6W'fñ6T÷&∂W$ñ6ˆ‚Ü6FVv˜'íí∞¢6ˆÁ7B6fr“U%4Ù‰≈ı4U%dî4UÙ4DTtı$îU5∂6FVv˜'ï“«¬∑”∞¢&WGW&‚¬ÊFódñ6ˆ‚á∞¢6∆74Ê÷S¢&gVV¬÷÷&∂W"◊w&"¿¢áF÷√¢«7‚6∆73“&gVV¬÷÷&∂W"÷∆&V¬G∂vWEW'6ˆÊ≈6W'fñ6T÷&∂W$6∆72Ü6FVv˜'íó“#‚G∂W66TÖD‘¬Ü6frÊñ6ˆ‚«¬/	˘8“"ó”¬˜7„Ê¿¢ñ6ˆÂ6ó¶S¢≥CB¬#E“¿¢ñ6ˆ‰Ê6Ü˜#¢≥#"¬%“¿¢˜WÊ6Ü˜#¢≥¬”–¢“ì∞ß–†¶gVÊ7Fñˆ‚vWEW'6ˆÊ≈6W'fñ6T÷&∂W$6∆72Ü6FVv˜'íí∞¢6ˆÁ7B∆WGFR“∞¢'&V∂f7C¢'2÷÷&∂W"÷'&V∂f7B"¿¢«VÊ6É¢'2÷÷&∂W"÷«VÊ6Ç"¿¢7WW&÷&∂WC¢'2÷÷&∂W"◊7WW&÷&∂WB"¿¢Fˆ&66Û¢'2÷÷&∂W"◊Fˆ&66Ú"¿¢v3¢'2÷÷&∂W"◊v2"¿¢F”¢'2÷÷&∂W"÷F“"¿¢Ü&÷7ì¢'2÷÷&∂W"◊Ü&÷7í"¿¢&∂ñÊs¢'2÷÷&∂W"◊&∂ñÊr ¢”∞¢&WGW&‚∆WGFU∂6FVv˜'ï“«¬'2÷÷&∂W"÷FVfV«B#∞ß–†¶gVÊ7Fñˆ‚&VÊFW%W'6ˆÊ≈6W'fñ6W4∆ó7BÇí∞¢VíÁW'6ˆÊ≈6W'fñ6W4∆ó7BÊñÊÊW$ÖD‘¬“"#∞¢ñbÇW'6ˆÊ≈6W'fñ6W5&W7V«G2Ê∆VÊwFÇí&WGW&„∞¢ñbÜ7FófUW'6ˆÊ≈6W'fñ6T6FVv˜'í””“&«VÊ6Ç"í∞¢&VÊFW$«VÊ6Ñw&˜WVD∆ó7BÇì∞¢&WGW&„∞¢–¢W'6ˆÊ≈6W'fñ6W5&W7V«G2Êf˜$V6ÇÇá∆6Rí”‚∞¢VíÁW'6ˆÊ≈6W'fñ6W4∆ó7BÊVÊD6Üñ∆BÜ'Vñ∆EW'6ˆÊ≈6W'fñ6U&˜rá∆6Ríì∞¢“ì∞ß–†¶gVÊ7Fñˆ‚&VÊFW$«VÊ6Ñw&˜WVD∆ó7BÇí∞¢6ˆÁ7B66WFVB“W'6ˆÊ≈6W'fñ6W5&W7V«G2Êfñ«FW"Çá∆6Rí”‚ó4÷V≈f˜V6ÜW$66WFVBá∆6RÁFw2íì∞¢6ˆÁ7B˜FÜW"“W'6ˆÊ≈6W'fñ6W5&W7V«G2Êfñ«FW"Çá∆6Rí”‚ó4÷V≈f˜V6ÜW$66WFVBá∆6RÁFw2íì∞¢6ˆÁ7B66WFVD6&B“Fˆ7V÷VÁBÊ7&VFTV∆V÷VÁBÇ&Fób"ì∞¢66WFVD6&BÊ6∆74Ê÷R“'6ñ◊∆R÷∆ó7B÷óFV“7F6∂VB#∞¢66WFVD6&BÊñÊÊW$ÖD‘¬“«7G&ˆÊsÓ)»R66WGFÊÚ'VˆÊí7FÚÇG∂66WFVBÊ∆VÊwFá“ì¬˜7G&ˆÊsÊ∞¢6ˆÁ7B66WFVD∆ó7B“Fˆ7V÷VÁBÊ7&VFTV∆V÷VÁBÇ&Fób"ì∞¢66WFVD∆ó7BÊ6∆74Ê÷R“'6ñ◊∆R÷∆ó7B#∞¢66WFVBÊf˜$V6ÇÇá∆6Rí”‚66WFVD∆ó7BÊVÊD6Üñ∆BÜ'Vñ∆EW'6ˆÊ≈6W'fñ6U&˜rá∆6Rííì∞¢66WFVD6&BÊVÊD6Üñ∆BÜ66WFVD∆ó7Bì∞¢VíÁW'6ˆÊ≈6W'fñ6W4∆ó7BÊVÊD6Üñ∆BÜ66WFVD6&Bì∞†¢6ˆÁ7B˜FÜW$6&B“Fˆ7V÷VÁBÊ7&VFTV∆V÷VÁBÇ&Fób"ì∞¢˜FÜW$6&BÊ6∆74Ê÷R“'6ñ◊∆R÷∆ó7B÷óFV“7F6∂VB#∞¢˜FÜW$6&BÊñÊÊW$ÖD‘¬“«7G&ˆÊsÓ(Kû˚àÚÊˆ‚66WGFÊÚÚÊˆ‚ñÊFñ6FÚÇG∂˜FÜW"Ê∆VÊwFá“ì¬˜7G&ˆÊsÊ∞¢6ˆÁ7B˜FÜW$∆ó7B“Fˆ7V÷VÁBÊ7&VFTV∆V÷VÁBÇ&Fób"ì∞¢˜FÜW$∆ó7BÊ6∆74Ê÷R“'6ñ◊∆R÷∆ó7B#∞¢˜FÜW"Êf˜$V6ÇÇá∆6Rí”‚˜FÜW$∆ó7BÊVÊD6Üñ∆BÜ'Vñ∆EW'6ˆÊ≈6W'fñ6U&˜rá∆6Rííì∞¢˜FÜW$6&BÊVÊD6Üñ∆BÜ˜FÜW$∆ó7Bì∞¢VíÁW'6ˆÊ≈6W'fñ6W4∆ó7BÊVÊD6Üñ∆BÜ˜FÜW$6&Bì∞ß–†¶gVÊ7Fñˆ‚'Vñ∆EW'6ˆÊ≈6W'fñ6U&˜rá∆6Rí∞¢6ˆÁ7B&˜r“Fˆ7V÷VÁBÊ7&VFTV∆V÷VÁBÇ&Fób"ì∞¢&˜rÊ6∆74Ê÷R“'6ñ◊∆R÷∆ó7B÷óFV“7F6∂VB#∞¢&˜rÊFF6WBÁ∆6TñB“7G&ñÊrá∆6RÊñBì∞¢6ˆÁ7BÜVB“Fˆ7V÷VÁBÊ7&VFTV∆V÷VÁBÇ&Fób"ì∞¢ÜVBÊ6∆74Ê÷R“'W'6ˆÊ¬◊6W'fñ6R◊&˜r÷ÜVB#∞¢6ˆÁ7Bñ6ˆ‰'F‚“7&VFT'WGFˆ‚ÖU%4Ù‰≈ı4U%dî4UÙ4DTtı$îU5∑∆6RÊ6FVv˜'ï”ÚÊñ6ˆ‚«¬/	˘8“"¬Çí”‚6V∆V7EW'6ˆÊ≈6W'fñ6Rá∆6RÊñBíì∞¢ñ6ˆ‰'F‚Ê6∆74∆ó7BÊFBÇ&7Fñˆ‚÷ñ6ˆ‚÷'F‚"ì∞¢6ˆÁ7BÊ÷T'F‚“7&VFT'WGFˆ‚á∆6RÊÊ÷R¬Çí”‚6V∆V7EW'6ˆÊ≈6W'fñ6Rá∆6RÊñBíì∞¢Ê÷T'F‚Ê6∆74∆ó7BÊFBÇ'W'6ˆÊ¬◊6W'fñ6R÷Ê÷R÷'F‚"ì∞¢6ˆÁ7B÷WF“Fˆ7V÷VÁBÊ7&VFTV∆V÷VÁBÇ'6÷∆¬"ì∞¢÷WFÊ6∆74Ê÷R“&◊WFVB#∞¢÷WFÁFWáD6ˆÁFVÁB“f˜&÷DFó7FÊ6Rá∆6RÊFó7FÊ6Rì∞¢6ˆÁ7BFWáEw&“Fˆ7V÷VÁBÊ7&VFTV∆V÷VÁBÇ&Fób"ì∞¢FWáEw&Ê6∆74Ê÷R“'W'6ˆÊ¬◊6W'fñ6R◊&˜r÷÷ñ‚#∞¢FWáEw&ÊVÊD6Üñ∆BÜÊ÷T'F‚ì∞¢FWáEw&ÊVÊD6Üñ∆BÜ÷WFì∞¢ÜVBÊVÊD6Üñ∆BÜñ6ˆ‰'F‚ì∞¢ÜVBÊVÊD6Üñ∆BáFWáEw&ì∞¢&˜rÊVÊD6Üñ∆BÜÜVBì∞¢ñbÜWáÊFVEW'6ˆÊ≈6W'fñ6TñB””“7G&ñÊrá∆6RÊñBíí∞¢&˜rÊ6∆74∆ó7BÊFBÇ&ó2◊6V∆V7FVB"ì∞¢&˜rÊVÊD6Üñ∆BÜ'Vñ∆DWáÊFVEW'6ˆÊ≈6W'fñ6TFWFñ«2á∆6Ríì∞¢–¢&WGW&‚&˜s∞ß–†¶gVÊ7Fñˆ‚6V∆V7EW'6ˆÊ≈6W'fñ6Rá∆6TñBí∞¢WáÊFVEW'6ˆÊ≈6W'fñ6TñB“WáÊFVEW'6ˆÊ≈6W'fñ6TñB””“7G&ñÊrá∆6TñBíÚ""¢7G&ñÊrá∆6TñBì∞¢&VÊFW%W'6ˆÊ≈6W'fñ6W4∆ó7BÇì∞ß–†¶gVÊ7Fñˆ‚'Vñ∆DWáÊFVEW'6ˆÊ≈6W'fñ6TFWFñ«2á∆6Rí∞¢6ˆÁ7Bw&“Fˆ7V÷VÁBÊ7&VFTV∆V÷VÁBÇ&Fób"ì∞¢w&Ê6∆74Ê÷R“'W'6ˆÊ¬◊6W'fñ6R÷WáÊFVB#∞¢6ˆÁ7BFw2“∆6RÁFw2«¬∑”∞¢6ˆÁ7BÊd'F‚“7&VFT'WGFˆ‚Ç$Êfñv"¬Çí”‚∞¢vñÊF˜rÊ˜V‚ÜáGG3¢Ú˜wwrÊvˆˆv∆RÊ6ˆ“ˆ÷2ˆFó"Ûˆì”fFW7FñÊFñˆ„“G∑∆6RÊ∆G“¬G∑∆6RÊ∆ˆÁ÷¬%ˆ&∆Ê≤"ì∞¢“ì∞¢Êd'F‚Ê6∆74∆ó7BÊFBÇ&'F‚◊&ñ÷'í"ì∞¢6ˆÁ7Bvˆˆv∆TFWFñ«4'F‚“7&VFT'WGFˆ‚Ç$vˆˆv∆RFWGFv∆í"¬Çí”‚∞¢6ˆÁ7BVW'í“VÊ6ˆFUU$î6ˆ◊ˆÊVÁBÜG∑∆6RÊÊ÷W“G∑∆6RÊ∆G“¬G∑∆6RÊ∆ˆÁ÷ì∞¢vñÊF˜rÊ˜V‚ÜáGG3¢Ú˜wwrÊvˆˆv∆RÊ6ˆ“ˆ÷2˜6V&6ÇÛˆì”gVW'ì“G∑VW'ó÷¬%ˆ&∆Ê≤"ì∞¢“ì∞¢6ˆÁ7B6∆˜6T'F‚“7&VFT'WGFˆ‚Ç$6ÜóVFíFWGFv∆í"¬Çí”‚6V∆V7EW'6ˆÊ≈6W'fñ6Rá∆6RÊñBíì∞¢6ˆÁ7B7FñˆÁ2“Fˆ7V÷VÁBÊ7&VFTV∆V÷VÁBÇ&Fób"ì∞¢7FñˆÁ2Ê6∆74Ê÷R“&óFV“÷7FñˆÁ2#∞¢7FñˆÁ2ÊVÊD6Üñ∆BÜÊd'F‚ì∞¢7FñˆÁ2ÊVÊD6Üñ∆BÜvˆˆv∆TFWFñ«4'F‚ì∞¢7FñˆÁ2ÊVÊD6Üñ∆BÜ6∆˜6T'F‚ì∞¢w&ÊñÊÊW$ÖD‘¬“ ¢«„∆#‰Êˆ÷S£¬ˆ#‚G∂W66TÖD‘¬á∆6RÊÊ÷Ró”¬˜‡¢«„∆#‰Fó7FÁ¶£¬ˆ#‚G∂W66TÖD‘¬Üf˜&÷DFó7FÊ6Rá∆6RÊFó7FÊ6Ríó”¬˜‡¢«„∆#‰ñÊFó&óß¶Û£¬ˆ#‚G∂W66TÖD‘¬Üf˜&÷DFG&W72áFw2íó”¬˜‡¢∞¢w&ÊVÊD6Üñ∆BÜ7FñˆÁ2ì∞¢6ˆÁ7BFWFñ«2“Fˆ7V÷VÁBÊ7&VFTV∆V÷VÁBÇ&Fób"ì∞¢FWFñ«2Ê6∆74Ê÷R“'6ñ◊∆R÷∆ó7B#∞¢FWFñ«2ÊñÊÊW$ÖD‘¬“&VÊFW$WáFVÊFVEW'6ˆÊ≈6W'fñ6TFWFñ«2á∆6Rì∞¢w&ÊVÊD6Üñ∆BÜFWFñ«2ì∞¢&WGW&‚w&∞ß–†¶gVÊ7Fñˆ‚&VÊFW$WáFVÊFVEW'6ˆÊ≈6W'fñ6TFWFñ«2á∆6Rí∞¢6ˆÁ7BFw2“∆6RÁFw2«¬∑”∞¢6ˆÁ7B6fr“U%4Ù‰≈ı4U%dî4UÙ4DTtı$îU5∑∆6RÊ6FVv˜'ï”∞¢6ˆÁ7B&˜w2“µ”∞¢ñbá∆6RÊ6FVv˜'í””“&«VÊ6Ç"í∞¢&˜w2ÁW6ÇÜ«„∆#‰'VˆÊí7FÛ£¬ˆ#‚G∂W66TÖD‘¬Üf˜&÷D÷V≈f˜V6ÜW%7FGW2áFw2íó”¬˜Êì∞¢–¢6ˆÁ7BFWFñƒfñV∆G2“'&íÊó4'&íÜ6fsÚÊFWFñƒfñV∆G2íÚ6frÊFWFñƒfñV∆G2¢µ”∞¢FWFñƒfñV∆G2Êf˜$V6ÇÇÜfñV∆Bí”‚∞¢6ˆÁ7B&uf«VR“Fw5∂fñV∆E”∞¢ñbá&uf«VR”“ÁV∆¬«¬&uf«VR””“""í&WGW&„∞¢&˜w2ÁW6ÇÜ«„∆#‚G∂W66TÖD‘¬Üf˜&÷DFWFñƒfñV∆D∆&V¬ÜfñV∆Bíó”£¬ˆ#‚G∂W66TÖD‘¬Ö7G&ñÊrá&uf«VRíó”¬˜Êì∞¢“ì∞¢6ˆÁ7B∆≈Fu&˜w2“ˆ&¶V7BÊVÁG&ñW2áFw2ê¢Êfñ«FW"ÇÖ∂∂Wí¬f«VU“í”‚f«VR“ÁV∆¬bb7G&ñÊráf«VRíÁG&ñ“Çí”“""ê¢Á6˜'BÇÖ∂“¬∂%“í”‚Ê∆ˆ6∆T6ˆ◊&RÜ"íê¢Ê÷ÇÖ∂∂Wí¬f«VU“í”‚«„∆#‚G∂W66TÖD‘¬Üf˜&÷DFWFñƒfñV∆D∆&V¬Ü∂Wííó”£¬ˆ#‚G∂W66TÖD‘¬Ö7G&ñÊráf«VRíó”¬˜Êì∞¢ñbÜ∆≈Fu&˜w2Ê∆VÊwFÇí∞¢&˜w2ÁW6ÇÇ#∆á#‚"ì∞¢&˜w2ÁW6ÇÇ#«„∆#ÂGWGFííFFíFó7ˆÊñ&ñ∆ì£¬ˆ#„¬˜‚"ì∞¢&˜w2ÁW6ÇÇ‚‚Ê∆≈Fu&˜w2ì∞¢–¢ñbÇ&˜w2Ê∆VÊwFÇí&˜w2ÁW6ÇÇ#«6∆73“v◊WFVBs‰ÊW77V‚FWGFv∆ñÚvvóVÁFófÚFó7ˆÊñ&ñ∆R„¬˜‚"ì∞¢&WGW&‚&˜w2Ê¶ˆñ‚Ç""ì∞ß–†¶gVÊ7Fñˆ‚f˜&÷DFWFñƒfñV∆D∆&V¬ÜfñV∆Bí∞¢6ˆÁ7B∆&V«2“∞¢˜VÊñÊuˆÜ˜W'3¢$˜&&í"¿¢7Vó6ñÊS¢%FóÚ7V6ñÊ"¿¢F∂Vvì¢%F∂R÷ví"¿¢FV∆ófW'ì¢$6ˆÁ6VvÊ"¿¢&6ˆÁF7CßÜˆÊR#¢%FV∆VfˆÊÚ"¿¢vV'6óFS¢%6óFÚvV""¿¢'ñ÷VÁC¶÷V≈˜f˜V6ÜW"#¢$'VˆÊí7FÚ"¿¢'ñ÷VÁCß6ˆFWÜÚ#¢%6ˆFWÜÚ"¿¢'ñ÷VÁC¶VFVÁ&VB#¢$VFVÁ&VB"¿¢'ñ÷VÁCßFñ6∂WE˜&W7FW&ÁB#¢%Fñ6∂WB&W7FW&ÁB"¿¢&FñWCßfVvWF&ñ‚#¢$˜¶ñˆÊífVvWF&ñÊR"¿¢fVS¢$v÷VÁFÚ"¿¢vÜVV∆6Üó#¢$66W76ñ&ñ∆óL:6'&˜ß¶ñÊ"¿¢˜W&F˜#¢$˜W&F˜&R"¿¢66Öˆñ„¢%fW'6÷VÁFÚ6ˆÁFÁFí"¿¢6ˆÁF7F∆W73¢$6ˆÁF7F∆W72"¿¢&7W'&VÊ7ì§UU"#¢$WW&Ú"¿¢Fó7VÁ6ñÊs¢$Fó7VÁ6¶ñˆÊR"¿¢66W73¢$66W76Ú"¿¢66óGì¢$66óL:"¿¢'&ÊC¢$÷&6ÜñÚ ¢”∞¢&WGW&‚∆&V«5∂fñV∆E“«¬fñV∆C∞ß–†¶gVÊ7Fñˆ‚f˜&÷DFG&W72áFw2í∞¢6ˆÁ7B'G2“∞¢Fw5≤&FG#ß7G&VWB%“¿¢Fw5≤&FG#¶Ü˜W6VÁV÷&W"%“¿¢Fw5≤&FG#¶6óGí%–¢“Êfñ«FW"Ñ&ˆˆ∆V‚ì∞¢&WGW&‚'G2Ê∆VÊwFÇÚ'G2Ê¶ˆñ‚Ç"¬"í¢$Êˆ‚Fó7ˆÊñ&ñ∆R#∞ß–†¶gVÊ7Fñˆ‚ó4÷V≈f˜V6ÜW$66WFVBáFw2í∞¢ñbÇFw2í&WGW&‚f«6S∞¢6ˆÁ7B˜6óFófTfñV∆G2“≤'ñ÷VÁC¶÷V≈˜f˜V6ÜW""¬'ñ÷VÁCß6ˆFWÜÚ"¬'ñ÷VÁC¶VFVÁ&VB"¬'ñ÷VÁCßFñ6∂WE˜&W7FW&ÁB%”∞¢&WGW&‚˜6óFófTfñV∆G2Á6ˆ÷RÇÜfñV∆Bí”‚7G&ñÊráFw5∂fñV∆E“«¬""íÁFÙ∆˜vW$66RÇí””“'ñW2"ì∞ß–†¶gVÊ7Fñˆ‚f˜&÷D÷V≈f˜V6ÜW%7FGW2áFw2í∞¢ñbÜó4÷V≈f˜V6ÜW$66WFVBáFw2íí&WGW&‚$66WGFFí#∞¢6ˆÁ7BfñV∆G2“≤'ñ÷VÁC¶÷V≈˜f˜V6ÜW""¬'ñ÷VÁCß6ˆFWÜÚ"¬'ñ÷VÁC¶VFVÁ&VB"¬'ñ÷VÁCßFñ6∂WE˜&W7FW&ÁB%”∞¢ñbÜfñV∆G2Á6ˆ÷RÇÜfñV∆Bí”‚7G&ñÊráFw5∂fñV∆E“«¬""íÁFÙ∆˜vW$66RÇí””“&ÊÚ"íí&WGW&‚$Êˆ‚66WGFFí#∞¢&WGW&‚$Êˆ‚7V6ñfñ6FÚ#∞ß–†¶gVÊ7Fñˆ‚vWE6V∆V7FVEW'6ˆÊ≈6W'fñ6W5&FóW2Çí∞¢6ˆÁ7Bf«VR“ÁV÷&W"áVíÁW'6ˆÊ≈6W'fñ6W5&FóW3ÚÁf«VR«¬3ì∞¢ñbÇÁV÷&W"Êó4fñÊóFRáf«VRí«¬f«VR¬Sí&WGW&‚3∞¢&WGW&‚f«VS∞ß–†¶7ñÊ2gVÊ7Fñˆ‚fWF6ÖW'6ˆÊ≈6W'fñ6W4g&ˆ‘˜fW'72Ü6FVv˜'í¬∆B¬∆Êr¬&FóW4÷WFW'2“3í∞¢6ˆÁ7B6fr“U%4Ù‰≈ı4U%dî4UÙ4DTtı$îU5∂6FVv˜'ï”∞¢ñbÇ6frí&WGW&‚≤V∆V÷VÁG3¢µ“”∞¢6ˆÁ7Bg&v÷VÁB“6frÁVW'ê¢Á&W∆6T∆¬Ç'∂∆G“"¬7G&ñÊrÜ∆Bíê¢Á&W∆6T∆¬Ç'∂∆Êw“"¬7G&ñÊrÜ∆Êríê¢Á&W∆6T∆¬Ç'∑&FóW7“"¬7G&ñÊrá&FóW4÷WFW'2íì∞¢6ˆÁ7BVW'í“ ¢∂˜WC¶ß6ˆÂ’∑Fñ÷V˜WC£#U”∞¢Ä¢G∂g&v÷VÁG–¢ì∞¢˜WB6VÁFW"Fw3∞¢∞¢6ˆÁ7Bfó'7E&W7V«B“vóBfWF6Ñ˜fW'75vóFÑf∆∆&6≤áVW'íì∞¢ñbÜ6FVv˜'í”“&«VÊ6Ç"«¬Üfó'7E&W7V«BÊV∆V÷VÁG2«¬µ“íÊ∆VÊwFÇí&WGW&‚fó'7E&W7V«C∞¢6ˆÁ7B'&ˆD«VÊ6ÖVW'í“ ¢∂˜WC¶ß6ˆÂ’∑Fñ÷V˜WC£#U”∞¢Ä¢ÊˆFU≤&÷VÊóGí'‚'&W7FW&ÁG∆f7EˆfˆˆG∆fˆˆEˆ6˜W'G∆6ÁFVVÁ∆6fW∆&"%“Ü&˜VÊC¢G¥÷FÇÊ÷ÇÉ#¬&FóW4÷WFW'2ó“¬G∂∆G“¬G∂∆Êw“ì∞¢vï≤&÷VÊóGí'‚'&W7FW&ÁG∆f7EˆfˆˆG∆fˆˆEˆ6˜W'G∆6ÁFVVÁ∆6fW∆&"%“Ü&˜VÊC¢G¥÷FÇÊ÷ÇÉ#¬&FóW4÷WFW'2ó“¬G∂∆G“¬G∂∆Êw“ì∞¢&V∆FñˆÂ≤&÷VÊóGí'‚'&W7FW&ÁG∆f7EˆfˆˆG∆fˆˆEˆ6˜W'G∆6ÁFVVÁ∆6fW∆&"%“Ü&˜VÊC¢G¥÷FÇÊ÷ÇÉ#¬&FóW4÷WFW'2ó“¬G∂∆G“¬G∂∆Êw“ì∞¢ì∞¢˜WB6VÁFW"Fw3∞¢∞¢&WGW&‚fWF6Ñ˜fW'75vóFÑf∆∆&6≤Ü'&ˆD«VÊ6ÖVW'íì∞ß–†¶7ñÊ2gVÊ7Fñˆ‚fWF6Ñ˜fW'75vóFÑf∆∆&6≤áVW'íí∞¢6ˆÁ7BVÊGˆñÁG2“∞¢&áGG3¢Úˆ˜fW'72÷íÊFRˆíˆñÁFW'&WFW""¿¢&áGG3¢Úˆ˜fW'72Ê∑V÷íÁ7ó7FV◊2ˆíˆñÁFW'&WFW" ¢”∞¢∆WB∆7DW'&˜"“ÁV∆√∞¢f˜"Ü6ˆÁ7BVÊGˆñÁBˆbVÊGˆñÁG2í∞¢G'í∞¢6ˆÁ7B6ˆÁG&ˆ∆∆W"“ÊWr&˜'D6ˆÁG&ˆ∆∆W"Çì∞¢6ˆÁ7BFñ÷V˜WB“6WEFñ÷V˜WBÇÇí”‚6ˆÁG&ˆ∆∆W"Ê&˜'BÇí¬#ì∞¢6ˆÁ7B&W7ˆÁ6R“vóBfWF6ÇÜVÊGˆñÁB¬∞¢÷WFÜˆC¢%ı5B"¿¢&ˆGì¢VW'í¿¢6ñvÊ√¢6ˆÁG&ˆ∆∆W"Á6ñvÊ¿¢“ì∞¢6∆V%Fñ÷V˜WBáFñ÷V˜WBì∞¢ñbÇ&W7ˆÁ6RÊˆ≤íFá&˜rÊWrW'&˜"Ü˜fW'72G∑&W7ˆÁ6RÁ7FGW7÷ì∞¢&WGW&‚vóB&W7ˆÁ6RÊß6ˆ‚Çì∞¢“6F6ÇÜW'&˜"í∞¢∆7DW'&˜"“W'&˜#∞¢–¢–¢Fá&˜r∆7DW'&˜"«¬ÊWrW'&˜"Ç$˜fW'72Êˆ‚Fó7ˆÊñ&ñ∆R"ì∞ß–†¶gVÊ7Fñˆ‚vWD7W'&VÁD˜W&F˜%˜6óFñˆ‰76ñvÊ÷VÁBÇí∞¢6ˆÁ7BFFT∂Wí“vWD7FófU7VG&TFFT∂WíÇì∞¢6ˆÁ7B76ñvÊ÷VÁB“vWD7W'&VÁEW6W$76ñvÊVD6ˆ÷÷W76Tf˜$FFRÜFFT∂Wíï≥”∞¢6ˆÁ7B÷F6ÜVE&˜r“76ñvÊ÷VÁCÚÊ÷F6ÜVE&˜w3ÚÂ≥“«¬ÁV∆√∞¢&WGW&‚∞¢˜W&F˜$Ê÷S¢÷F6ÜVE&˜sÚÊ÷F6ÜVDÊ÷R«¬7W'&VÁEW6W#ÚÊFó7∆îÊ÷R«¬7W'&VÁEW6W#ÚÊV÷ñ¬«¬$˜W&F˜&R"¿¢7VG&ñÊFWÉ¢÷F6ÜVE&˜sÚÁ7VG&ñÊFWÇ«¬""¿¢7VG&∆&V√¢÷F6ÜVE&˜sÚÁ7VG&∆&V¬«¬""¿¢7VG&Ê÷S¢÷F6ÜVE&˜sÚÁ&˜sÚÊÊˆ÷R«¬÷F6ÜVE&˜sÚÁ&˜sÚÊÊ÷R«¬""¿¢6ˆ÷÷W76ñC¢76ñvÊ÷VÁCÚÊ6ˆ÷÷W76ñB«¬6V∆V7FVD6ˆ÷÷W76ñB«¬""¿¢6ˆ÷÷W76Ê÷S¢76ñvÊ÷VÁCÚÊ6ˆ÷÷W76Ê÷R«¬6V∆V7FVD6ˆ÷÷W76Ê÷R«¬""¿¢&ñfW&ñ÷VÁFÙFF¢FFT∂Wí«¬" ¢”∞ß–†¶gVÊ7Fñˆ‚ñÊóDvVˆ∆ˆ6Fñˆ‚Çí∞¢ñbÇÊfñvF˜"ÊvVˆ∆ˆ6Fñˆ‚í∞¢VíÊw57FGW2ÁFWáD6ˆÁFVÁB“$vVˆ∆ˆ6∆óß¶¶ñˆÊRÊˆ‚7W˜'FFF¬'&˜w6W"‚#∞¢WFFT∆ˆ6FñˆÂv&ÊñÊu7FFRá≤W&÷ó76ñˆ„¢'VÊfñ∆&∆R"¬VÊ&∆VC¢f«6R¬÷ˆFS¢&&∆ˆ6∂VB"“ì∞¢&WGW&„∞¢–¢fˆñB7ñÊ4∆ˆ6Fñˆ‰fñ∆&ñ∆óGíÇì∞¢ñbÜvVˆ∆ˆ6FñˆÂvF6ÑñB“ÁV∆¬íÊfñvF˜"ÊvVˆ∆ˆ6Fñˆ‚Ê6∆V%vF6ÇÜvVˆ∆ˆ6FñˆÂvF6ÑñBì∞†¢6ˆÁ7BˆÂ˜6óFñˆ‚“á˜2í”‚∞¢∆FW7DvVˆ∆ˆ6Fñˆ‰6ˆ˜&G2“˜2Ê6ˆ˜&G3∞¢WFFT7W'&VÁEW6W%˜6óFñˆ‚á˜2Ê6ˆ˜&G2¬˜2ÁFñ÷W7F◊¬≤&VÊFW#¢f«6R“ì∞¢Wf«VFUFñ÷'&GW&&V÷ñÊFW'2Çì∞¢VíÊw57FGW2ÁFWáD6ˆÁFVÁB“%˜6ó¶ñˆÊRGFóf¢ñ◊ñÁFí˜&FñÊFíW"Fó7FÁ¶‚#∞¢WFFT∆ˆ6FñˆÂv&ÊñÊu7FFRá≤W&÷ó76ñˆ„¢&w&ÁFVB"¬VÊ&∆VC¢G'VR¬÷ˆFS¢'&ˆ◊B"“ì∞¢&VÊFW$ñ◊ñÁFíÇì∞¢&VÊFW$÷Çì∞¢Wf«VFTñ◊ñÁFı&˜Üñ÷óGî∆W'G2Çì∞¢WFÙ6ˆ◊∆WFU76VE6Ê˜u&ˆG2ÇíÊ6F6ÇÇÜW'&˜"í”‚6ˆÁ6ˆ∆RÁv&‚Ç$6ˆ◊∆WF÷VÁFÚWFˆ÷Fñ6ÚfñRÊWfRÊˆ‚&óW66óFÛ¢"¬W'&˜"íì∞¢”∞†¢ÊfñvF˜"ÊvVˆ∆ˆ6Fñˆ‚ÊvWD7W'&VÁE˜6óFñˆ‚Çá˜2í”‚∞¢ˆÂ˜6óFñˆ‚á˜2ì∞¢fWF6ÖvVFÜW"Çì∞¢“¬Çí”‚∞¢6∆V$7W'&VÁEW6W%˜6óFñˆ‚Çì∞¢VíÊw57FGW2ÁFWáD6ˆÁFVÁB“%˜6ó¶ñˆÊRÊˆ‚Fó7ˆÊñ&ñ∆R#∞¢WFFT∆ˆ6FñˆÂv&ÊñÊu7FFRá≤VÊ&∆VC¢f«6R¬÷ˆFS¢∆ˆ6FñˆÂW&÷ó76ñˆ‚””“&FVÊñVB"Ú&&∆ˆ6∂VB"¢'&ˆ◊B"“ì∞¢fWF6ÖvVFÜW"Çì∞¢“¬∞¢VÊ&∆TÜñvÑ67W&7ì¢G'VR¿¢Fñ÷V˜WC¢É ¢“ì∞†¢vVˆ∆ˆ6FñˆÂvF6ÑñB“ÊfñvF˜"ÊvVˆ∆ˆ6Fñˆ‚ÁvF6Ö˜6óFñˆ‚ÜˆÂ˜6óFñˆ‚¬Çí”‚∞¢6∆V$7W'&VÁEW6W%˜6óFñˆ‚Çì∞¢VíÊw57FGW2ÁFWáD6ˆÁFVÁB“%˜6ó¶ñˆÊRÊˆ‚Fó7ˆÊñ&ñ∆R#∞¢WFFT∆ˆ6FñˆÂv&ÊñÊu7FFRá≤VÊ&∆VC¢f«6R¬÷ˆFS¢∆ˆ6FñˆÂW&÷ó76ñˆ‚””“&FVÊñVB"Ú&&∆ˆ6∂VB"¢'&ˆ◊B"“ì∞¢“¬∞¢VÊ&∆TÜñvÑ67W&7ì¢G'VR¿¢÷Üñ◊V‘vS¢¿¢Fñ÷V˜WC¢ ¢“ì∞ß–†¶gVÊ7Fñˆ‚'Vñ∆Dñ◊ñÁFÙWfVÁD∆ˆ6ƒ∂WíÜWfVÁEGóR¬6ˆ÷÷W76ñB¬ñ◊ñÁFÙ∂Wíí∞¢&WGW&‚ÜW&Ê˜FñfñVC¢G∂WfVÁEGóW”¢G∂6ˆ÷÷W76ñB«¬"'”¢G∂ñ◊ñÁFÙ∂Wí«¬"'÷∞ß–†¶gVÊ7Fñˆ‚Ü4∆ˆ6ƒñ◊ñÁFÙWfVÁBÜWfVÁEGóR¬6ˆ÷÷W76ñB¬ñ◊ñÁFÙ∂Wíí∞¢&WGW&‚∆ˆ6≈7F˜&vRÊvWDóFV“Ü'Vñ∆Dñ◊ñÁFÙWfVÁD∆ˆ6ƒ∂WíÜWfVÁEGóR¬6ˆ÷÷W76ñB¬ñ◊ñÁFÙ∂Wííí””“##∞ß–†¶gVÊ7Fñˆ‚÷&¥∆ˆ6ƒñ◊ñÁFÙWfVÁBÜWfVÁEGóR¬6ˆ÷÷W76ñB¬ñ◊ñÁFÙ∂Wíí∞¢∆ˆ6≈7F˜&vRÁ6WDóFV“Ü'Vñ∆Dñ◊ñÁFÙWfVÁD∆ˆ6ƒ∂WíÜWfVÁEGóR¬6ˆ÷÷W76ñB¬ñ◊ñÁFÙ∂Wíí¬#"ì∞ß–†¶gVÊ7Fñˆ‚Wf«VFTñ◊ñÁFı&˜Üñ÷óGî∆W'G2Çí∞¢ñbÇ6V∆V7FVD6ˆ÷÷W76ñB«¬7W'&VÁEW6W%˜2«¬'&íÊó4'&íÜ7W'&VÁDñ◊ñÁFíí«¬7W'&VÁDñ◊ñÁFíÊ∆VÊwFÇí&WGW&„∞¢6ˆÁ7BFˆFı6˜'FVB“6ˆ÷&ñÊTñ◊ñÁFîf˜%fñWrÜ7W'&VÁDñ◊ñÁFíê¢Êfñ«FW"ÇÜñ◊ñÁFÚí”‚ñ◊ñÁFÚÊFˆÊRê¢Á6˜'BÇÜ¬"í”‚Fó7FÊ6Tg&ˆ’W6W"Üí“Fó7FÊ6Tg&ˆ’W6W"Ü"íì∞¢6ˆÁ7BÊV&W7B“FˆFı6˜'FVE≥”∞¢ñbÇÊV&W7Bí∞¢7FófTÊV&'îñ◊ñÁFÙ6ˆÁFWáB“ÁV∆√∞¢&WGW&„∞¢–¢6ˆÁ7BÊV&W7D∂Wí“'Vñ∆Dñ◊ñÁFÙ∂WíÜÊV&W7Bì∞¢6ˆÁ7BFó7FÊ6T∂““Fó7FÊ6Tg&ˆ’W6W"ÜÊV&W7Bì∞¢ñbÇÁV÷&W"Êó4fñÊóFRÜFó7FÊ6T∂“íí&WGW&„∞†¢ñbÇ7FófTÊV&'îñ◊ñÁFÙ6ˆÁFWáBbbFó7FÊ6T∂“√“$ıÑî‘ïEïÙ‰T%Ù¥“í∞¢7FófTÊV&'îñ◊ñÁFÙ6ˆÁFWáB“≤6ˆ÷÷W76ñC¢6V∆V7FVD6ˆ÷÷W76ñB¬ñ◊ñÁFÙ∂Wì¢ÊV&W7D∂Wí”∞¢ñbÇÜ4∆ˆ6ƒñ◊ñÁFÙWfVÁBÇ&ÊV""¬6V∆V7FVD6ˆ÷÷W76ñB¬ÊV&W7D∂Wííí∞¢÷&¥∆ˆ6ƒñ◊ñÁFÙWfVÁBÇ&ÊV""¬6V∆V7FVD6ˆ÷÷W76ñB¬ÊV&W7D∂Wíì∞¢6ˆÁ7B6ˆ÷÷W76Ê÷R“6V∆V7FVD6ˆ÷÷W76Ê÷R«¬$6ˆ÷÷W76#∞¢6ˆÁ7Bñ◊ñÁFÙÊ÷R“ÊV&W7BÊFVÊˆ÷ñÊ¶ñˆÊR«¬$ñ◊ñÁFÚ#∞¢V&∆ó6Ñv∆ˆ&ƒÊ˜Fñfñ6Fñˆ‰WfVÁBÇ&ñ◊ñÁFÚ÷ÊV""¬∞¢FóF∆S¢$˜W&F˜&Rfñ6ñÊÚñ◊ñÁFÚ"¿¢&ˆGì¢G∂7W'&VÁEW6W#ÚÊFó7∆îÊ÷R«¬7W'&VÁEW6W#ÚÊV÷ñ¬«¬$˜W&F˜&R'“:Çfñ6ñÊÚG∂ñ◊ñÁFÙÊ÷W“ÇG∂6ˆ÷÷W76Ê÷W“íÊ¿¢6ˆ÷÷W76ñC¢6V∆V7FVD6ˆ÷÷W76ñB¿¢6ˆ÷÷W76Ê÷R¿¢ñ◊ñÁFÙÊ÷R¿¢ñ◊ñÁFÙ∂Wì¢ÊV&W7D∂Wê¢“ì∞¢–¢&WGW&„∞¢–†¢ñbÇ7FófTÊV&'îñ◊ñÁFÙ6ˆÁFWáBí&WGW&„∞¢ñbÜ7FófTÊV&'îñ◊ñÁFÙ6ˆÁFWáBÊ6ˆ÷÷W76ñB”“6V∆V7FVD6ˆ÷÷W76ñB«¬7FófTÊV&'îñ◊ñÁFÙ6ˆÁFWáBÊñ◊ñÁFÙ∂Wí”“ÊV&W7D∂Wíí&WGW&„∞¢ñbÜFó7FÊ6T∂“¬$ıÑî‘ïEïÙtïÙ¥“í&WGW&„∞¢ñbÜÊV&W7BÊFˆÊRí∞¢7FófTÊV&'îñ◊ñÁFÙ6ˆÁFWáB“ÁV∆√∞¢&WGW&„∞¢–¢ñbÇÜ4∆ˆ6ƒñ◊ñÁFÙWfVÁBÇ&ví◊vóFÜ˜WB÷FˆÊR"¬6V∆V7FVD6ˆ÷÷W76ñB¬ÊV&W7D∂Wííí∞¢÷&¥∆ˆ6ƒñ◊ñÁFÙWfVÁBÇ&ví◊vóFÜ˜WB÷FˆÊR"¬6V∆V7FVD6ˆ÷÷W76ñB¬ÊV&W7D∂Wíì∞¢6ˆÁ7B6ˆ÷÷W76Ê÷R“6V∆V7FVD6ˆ÷÷W76Ê÷R«¬$6ˆ÷÷W76#∞¢6ˆÁ7Bñ◊ñÁFÙÊ÷R“ÊV&W7BÊFVÊˆ÷ñÊ¶ñˆÊR«¬$ñ◊ñÁFÚ#∞¢V&∆ó6Ñv∆ˆ&ƒÊ˜Fñfñ6Fñˆ‰WfVÁBÇ&ñ◊ñÁFÚ÷ví◊vóFÜ˜WB÷FˆÊR"¬∞¢FóF∆S¢$∆∆ˆÁFÊ÷VÁFÚ6VÁ¶dEDÚ"¿¢&ˆGì¢G∂7W'&VÁEW6W#ÚÊFó7∆îÊ÷R«¬7W'&VÁEW6W#ÚÊV÷ñ¬«¬$˜W&F˜&R'“6í:Ç∆∆ˆÁFÊFÚFG∂ñ◊ñÁFÙÊ÷W“6VÁ¶&V÷W&RdEDÚÊ¿¢6ˆ÷÷W76ñC¢6V∆V7FVD6ˆ÷÷W76ñB¿¢6ˆ÷÷W76Ê÷R¿¢ñ◊ñÁFÙÊ÷R¿¢ñ◊ñÁFÙ∂Wì¢ÊV&W7D∂Wê¢“ì∞¢–¢7FófTÊV&'îñ◊ñÁFÙ6ˆÁFWáB“ÁV∆√∞ß–†¶gVÊ7Fñˆ‚'Vñ∆EFñ÷'&GW&&V÷ñÊFW$∆ˆ6ƒ∂WíÜFFT∂Wí¬f66ñí∞¢&WGW&‚ÜW&Fñ÷'&GW&&V÷ñÊFW#¢G∂FFT∂Wó”¢G∂f66ñ÷∞ß–†¶gVÊ7Fñˆ‚vWD∆ˆ6ƒFFT∂WíÜFFRí∞¢6ˆÁ7BñV"“FFRÊvWDgV∆≈ñV"Çì∞¢6ˆÁ7B÷ˆÁFÇ“7G&ñÊrÜFFRÊvWD÷ˆÁFÇÇí≤íÁE7F'BÉ"¬#"ì∞¢6ˆÁ7BFí“7G&ñÊrÜFFRÊvWDFFRÇííÁE7F'BÉ"¬#"ì∞¢&WGW&‚G∑ñV'““G∂÷ˆÁFá““G∂Fó÷∞ß–†¶gVÊ7Fñˆ‚vWD÷ñÁWFW56ñÊ6T÷ñFÊñváBÜFFRí∞¢&WGW&‚FFRÊvWDÜ˜W'2Çí¢c≤FFRÊvWD÷ñÁWFW2Çì∞ß–†¶gVÊ7Fñˆ‚Wf«VFUFñ÷'&GW&&V÷ñÊFW'2ÜÊ˜r“ÊWrFFRÇíí∞¢ñbÇ7W'&VÁEW6W%˜2í&WGW&„∞¢6ˆÁ7BFó7FÊ6T÷WFW'2“ÜfW'6ñÊRÜ7W'&VÁEW6W%˜2Ê∆B¬7W'&VÁEW6W%˜2Ê∆Êr¬Dî‘%$EU$ıD$tUEÙƒB¬Dî‘%$EU$ıD$tUEÙƒ‰rí¢∞¢ñbÇÁV÷&W"Êó4fñÊóFRÜFó7FÊ6T÷WFW'2í«¬Fó7FÊ6T÷WFW'2‚Dî‘%$EU$ı$DïU5Ù“í&WGW&„∞†¢6ˆÁ7B÷ñÁWFW2“vWD÷ñÁWFW56ñÊ6T÷ñFÊñváBÜÊ˜rì∞¢∆WBf66ñ“"#∞¢∆WBÊ˜Fñfñ6FñˆÂFWáB“"#∞¢ñbÜ÷ñÁWFW2„“Dî‘%$EU$ÙTÂE$Dı5D%EÙ‘î‚bb÷ñÁWFW2√“Dî‘%$EU$ÙTÂE$DÙT‰EÙ‘î‚í∞¢f66ñ“&VÁG&F#∞¢Ê˜Fñfñ6FñˆÂFWáB“%$î4ı$DDíDíDî‘%$$R¬TÂE$D#∞¢“V«6RñbÜ÷ñÁWFW2„“Dî‘%$EU$ıU44ïDı5D%EÙ‘î‚bb÷ñÁWFW2√“Dî‘%$EU$ıU44ïDÙT‰EÙ‘î‚í∞¢f66ñ“'W66óF#∞¢Ê˜Fñfñ6FñˆÂFWáB“%$î4ı$DDíDíDî‘%$$R¬U44ïD#∞¢“V«6R∞¢&WGW&„∞¢–†¢6ˆÁ7BFFT∂Wí“vWD∆ˆ6ƒFFT∂WíÜÊ˜rì∞¢6ˆÁ7B&V÷ñÊFW$∂Wí“'Vñ∆EFñ÷'&GW&&V÷ñÊFW$∆ˆ6ƒ∂WíÜFFT∂Wí¬f66ñì∞¢ñbÜ∆ˆ6≈7F˜&vRÊvWDóFV“á&V÷ñÊFW$∂Wíí””“#"í&WGW&„∞¢∆ˆ6≈7F˜&vRÁ6WDóFV“á&V÷ñÊFW$∂Wí¬#"ì∞¢6Ü˜t∆ˆ6ƒÊ˜Fñfñ6Fñˆ‚Ç$ÜW&"¬∞¢&ˆGì¢Ê˜Fñfñ6FñˆÂFWáB¿¢Fs¢ÜW&◊Fñ÷'&GW&“G∂FFT∂Wó““G∂f66ñ÷¿¢&VÊ˜Fñgì¢f«6R¿¢FF¢≤W&√¢"‚ˆñÊFWÇÊáF÷¬"–¢“íÊ6F6ÇÇÜW'&˜"í”‚∞¢6ˆÁ6ˆ∆RÁv&‚Ç$ñÁfñÚÊ˜Fñfñ6Fñ÷'&GW&Êˆ‚&óW66óFÛ¢"¬W'&˜"ì∞¢“ì∞ß–†¶6ˆÁ7B4ïdî≈ı$ıDT5DîÙÂÙƒU%EıtR“&áGG3¢Úˆ÷RÁ&˜FW¶ñˆÊV6ófñ∆RÊv˜bÊóBˆóBˆ÷R◊&ó66Üíˆ&ˆ∆∆WGFñÊÚ÷Fí÷7&óFñ6óFÚ#∞¶6ˆÁ7B4ïdî≈ı$ıDT5DîÙÂÙtïDÖT%Ùí“&áGG3¢ÚˆíÊvóFáV"Ê6ˆ“˜&W˜2˜6“÷G2ÙE2‘&ˆ∆∆WGFñÊí‘7&óFñ6óF‘ñG&ˆvVˆ∆ˆvñ6‘ñG&V∆ñ6ˆ6ˆÁFVÁG2ˆfñ∆W2˜Ü÷√˜&Vc÷÷7FW"#∞¶6ˆÁ7B‘UDTıÛ4%Ù$4UıU$¬“&áGG3¢Ú˜wwr„6&÷WFVÚÊ6ˆ“ˆ÷WFVÚˆóF∆ñ#∞¶6ˆÁ7Btı$¥ƒî‘DUÙdı$T45EıU$¬“&áGG3¢ÚˆÁv˜&∂∆ñ÷FRÊóBˆ˜&FñÊÁ¶÷6∆FÚ÷∆f˜&Ú#∞¶6ˆÁ7BtTDÑU%ı$ıÖïıDÇ“"ˆí˜vVFÜW"#∞¶6ˆÁ7BtTDÑU%ı$ıÖïıT$ƒî5ıU$¬“&áGG3¢Úˆ7&VFófR◊7ó&Êñ∂í÷FFF&RÊÊWF∆ñgíÊˆí˜vVFÜW"#∞¶6ˆÁ7BtTDÑU%ÙdUD4ÖıDî‘TıUEÙ’2“S∞¶6ˆÁ7BƒU%EÙƒUdT≈Ù‘UD“∞¢w&VV„¢≤&Ê≥¢¬V÷ˆ¶ì¢/	˘˙""¬6∆74Ê÷S¢&∆W'B÷w&VV‚"¬∆&V√¢$ÊW77VÊ∆∆W'F"“¿¢ñV∆∆˜s¢≤&Ê≥¢¬V÷ˆ¶ì¢/	˘˙"¬6∆74Ê÷S¢&∆W'B◊ñV∆∆˜r"¬∆&V√¢$∆∆W'F&˜FW¶ñˆÊR6ófñ∆R"“¿¢˜&ÊvS¢≤&Ê≥¢"¬V÷ˆ¶ì¢/	˘˙"¬6∆74Ê÷S¢&∆W'B÷˜&ÊvR"¬∆&V√¢$∆∆W'F&˜FW¶ñˆÊR6ófñ∆R"“¿¢&VC¢≤&Ê≥¢2¬V÷ˆ¶ì¢/	˘KB"¬6∆74Ê÷S¢&∆W'B◊&VB"¬∆&V√¢$∆∆W'F&˜FW¶ñˆÊR6ófñ∆R"–ß”∞¶6ˆÁ7BƒU%EÙ¥Uïtı$E2“∞¢≤∂Wì¢'FV◊˜&∆í"¬∆&V√¢%FV◊˜&∆í"¬GFW&Á3¢≤'FV◊˜&∆í"¬'FV◊˜&∆R%““¿¢≤∂Wì¢'fVÁFÚ"¬∆&V√¢'fVÁFÚ"¬GFW&Á3¢≤'fVÁFÚ"¬'fVÁFí"¬&'W'&66%““¿¢≤∂Wì¢&ÊWfR"¬∆&V√¢&ÊWfR"¬GFW&Á3¢≤&ÊWfR"¬&ÊWfñ6FR%““¿¢≤∂Wì¢&vÜñ66ñÚ"¬∆&V√¢&vÜñ66ñÚ"¬GFW&Á3¢≤&vÜñ66ñÚ"¬&vV∆FR%““¿¢≤∂Wì¢&∆«WfñˆÊR"¬∆&V√¢&∆«WfñˆÊR"¬GFW&Á3¢≤&ñG&V∆ñ6Ú"¬&ñG&ˆvVˆ∆ˆvñ6Ú"¬&∆«WfñˆÊR"¬&∆∆v÷VÁFí%““¿¢≤∂Wì¢&ÊV&&ñ"¬∆&V√¢&ÊV&&ñ"¬GFW&Á3¢≤&ÊV&&ñ"¬&ÊV&&ñR%““¿¢≤∂Wì¢&6∆FÚ"¬∆&V√¢&6∆FÚW7G&V÷Ú"¬GFW&Á3¢≤&6∆FÚ"¬&ˆÊFFRFí6∆˜&R"¬'FV◊W&GW&RV∆WfFR%“–•”∞¶6ˆÁ7B‰dîtDîÙÂıtTDÑU%ı5E$Ù‰uıtî‰EÙ¥‘Ç“S∞¶6ˆÁ7B‰dîtDîÙÂıtTDÑU%ı5E$Ù‰uÙuU5EÙ¥‘Ç“s∞¶6ˆÁ7B‰dîtDîÙÂıtTDÑU%ı$TƒUdÂEı$îÂÙ‘““S∞¶6ˆÁ7B‰dîtDîÙÂıtTDÑU%ÙƒîtÖEı$îÂÙ‘ÖÙ‘““#∞¶6ˆÁ7B‰dîtDîÙÂıtTDÑU%Ù‰UÖEÙÑıU%ı$Ù$$îƒïEí“C∞¶6ˆÁ7B‰dîtDîÙÂıtTDÑU%ıDÖT‰DU%Ù4ÙDU2“ÊWr6WBÖ≥ìR¬ìb¬ìï“ì∞¶6ˆÁ7B‰dîtDîÙÂıtTDÑU%ı$îÂÙ4ÙDU2“ÊWr6WBÖ≥S¬S2¬SR¬Sb¬Sr¬c¬c2¬cR¬cb¬cr¬É¬É¬É%“ì∞¶6ˆÁ7Bî’îÂDııtTDÑU%Ù44ÑUıED≈Ù’2“R¢c¢∞¶6ˆÁ7Bî’îÂDııtTDÑU%ı$Te$U4ÖÙƒî‘ïB“C∞¶6ˆÁ7Bî’îÂDııtTDÑU%Ù$D4Öı4ï§R“#∞¶6ˆÁ7Bî’îÂDııtTDÑU%Ù4Ùı$Dî‰DUı$T4ï4îÙ‚“S∞¶6ˆÁ7Bî’îÂDııtTDÑU%ÙƒÙ4≈Ù44ÑUÙ‘ÖÙTÂE$îU2“C∞††¶gVÊ7Fñˆ‚'Vñ∆EvVFÜW$f˜&V67E&WVW7E&◊2áF&vWB¬≤˜W&FñˆÊ¬“f«6R““∑“í∞¢6ˆÁ7B&6U&◊2“∞¢∆FóGVFS¢7G&ñÊráF&vWBÊ∆Bí¿¢∆ˆÊvóGVFS¢7G&ñÊráF&vWBÊ∆ˆ‚í¿¢7W'&VÁC¢'FV◊W&GW&UÛ&“«vñÊE˜7VVEÛ“«vVFÜW%ˆ6ˆFR"¿¢Ü˜W&«ì¢'FV◊W&GW&UÛ&“«&V6óóFFñˆÂ˜&ˆ&&ñ∆óGí«6Ê˜vf∆¬«fó6ñ&ñ∆óGí«vVFÜW%ˆ6ˆFR«vñÊE˜7VVEÛ“"¿¢f˜&V67EˆFó3¢#R"¿¢Fñ÷W¶ˆÊS¢&WFÚ"¿¢÷ˆFV«3¢&&W7Eˆ÷F6Ç"¿¢6V∆≈˜6V∆V7Fñˆ„¢&∆ÊB"¿¢vñÊE˜7VVE˜VÊóC¢&∂÷Ç ¢”∞†¢ñbÇ˜W&FñˆÊ¬í&WGW&‚&6U&◊3∞†¢&WGW&‚∞¢‚‚Ê&6U&◊2¿¢7W'&VÁC¢'FV◊W&GW&UÛ&“∆&VÁE˜FV◊W&GW&R«&V∆FófUˆáV÷ñFóGïÛ&“«&V6óóFFñˆ‚«&ñ‚«6Ü˜vW'2«vVFÜW%ˆ6ˆFR«vñÊE˜7VVEÛ“«vñÊEˆFó&V7FñˆÂÛ“«vñÊEˆwW7G5Û“"¿¢÷ñÁWFV«ïÛS¢'FV◊W&GW&UÛ&“«&V∆FófUˆáV÷ñFóGïÛ&“∆&VÁE˜FV◊W&GW&R«&V6óóFFñˆ‚«&ñ‚«6Ü˜vW'2«6Ê˜vf∆¬«vVFÜW%ˆ6ˆFR«vñÊE˜7VVEÛ“«vñÊEˆFó&V7FñˆÂÛ“«vñÊEˆwW7G5Û“«fó6ñ&ñ∆óGí∆∆ñváFÊñÊu˜˜FVÁFñ¬"¿¢Ü˜W&«ì¢'FV◊W&GW&UÛ&“«&V∆FófUˆáV÷ñFóGïÛ&“∆FWu˜ˆñÁEÛ&“∆&VÁE˜FV◊W&GW&R«&V6óóFFñˆÂ˜&ˆ&&ñ∆óGí«&V6óóFFñˆ‚«&ñ‚«6Ü˜vW'2«6Ê˜vf∆¬«fó6ñ&ñ∆óGí«vVFÜW%ˆ6ˆFR«vñÊE˜7VVEÛ“«vñÊEˆFó&V7FñˆÂÛ“«vñÊEˆwW7G5Û“∆6R«WeˆñÊFWÇ"¿¢f˜&V67EˆÜ˜W'3¢#""¿¢f˜&V67Eˆ÷ñÁWFV«ïÛS¢#CÇ"¿¢f˜&V67EˆFó3¢#"¿¢Fñ÷W¶ˆÊS¢&WFÚ ¢”∞ß–†¶gVÊ7Fñˆ‚vWEvVFÜW%&˜áîVÊGˆñÁBÇí∞¢6ˆÁ7B&˜Fˆ6ˆ¬“vñÊF˜rÊ∆ˆ6Fñˆ‚Á&˜Fˆ6ˆ√∞¢6ˆÁ7BÜ˜7FÊ÷R“vñÊF˜rÊ∆ˆ6Fñˆ‚ÊÜ˜7FÊ÷S∞¢6ˆÁ7Bó4ÊFófR“&ˆˆ∆V‚Üv∆ˆ&≈FÜó2‰66óF˜#ÚÊó4ÊFófU∆Ff˜&”Ú‚Çíê¢«¬&˜Fˆ6ˆ¬””“&66óF˜#¢ ¢«¬&˜Fˆ6ˆ¬””“&ñˆÊñ3¢ ¢«¬á&˜Fˆ6ˆ¬””“&áGG3¢"bbÜÜ˜7FÊ÷R””“&∆ˆ6∆Ü˜7B"«¬Ü˜7FÊ÷R””“##r„„„"íì∞¢&WGW&‚ó4ÊFófRÚtTDÑU%ı$ıÖïıT$ƒî5ıU$¬¢tTDÑU%ı$ıÖïıDÉ∞ß–†¶7ñÊ2gVÊ7Fñˆ‚fWF6ÖvVFÜW%&W7ˆÁ6RáW&¬¬˜FñˆÁ2“∑“í∞¢6ˆÁ7B6ˆÁG&ˆ∆∆W"“ÊWr&˜'D6ˆÁG&ˆ∆∆W"Çì∞¢6ˆÁ7BFñ÷V˜WDñB“6WEFñ÷V˜WBÇÇí”‚6ˆÁG&ˆ∆∆W"Ê&˜'BÇí¬tTDÑU%ÙdUD4ÖıDî‘TıUEÙ’2ì∞¢G'í∞¢&WGW&‚vóBfWF6ÇáW&¬¬∞¢66ÜS¢˜FñˆÁ2Ê66ÜR«¬&ÊÚ◊7F˜&R"¿¢6ñvÊ√¢6ˆÁG&ˆ∆∆W"Á6ñvÊ¿¢“ì∞¢“6F6ÇÜW'&˜"í∞¢ñbÜW'&˜#ÚÊÊ÷R””“$&˜'DW'&˜""íFá&˜rÊWrW'&˜"Ç%Fñ÷V˜WB6W'fó¶ñÚ÷WFVÚ"ì∞¢Fá&˜rW'&˜#∞¢“fñÊ∆«í∞¢6∆V%Fñ÷V˜WBáFñ÷V˜WDñBì∞¢–ß–†¶7ñÊ2gVÊ7Fñˆ‚fWF6ÖvVFÜW$f˜&V67BáF&vWB¬˜FñˆÁ2“∑“í∞¢6ˆÁ7B&◊2“ÊWrU$≈6V&6Ö&◊2Ü'Vñ∆EvVFÜW$f˜&V67E&WVW7E&◊2áF&vWB¬˜FñˆÁ2íì∞¢6ˆÁ7BFó&V7EW&¬“áGG3¢ÚˆíÊ˜V‚÷÷WFVÚÊ6ˆ“˜cˆf˜&V67CÚG∑&◊2ÁFı7G&ñÊrÇó÷∞¢6ˆÁ7B&˜áï&◊2“ÊWrU$≈6V&6Ö&◊2á∞¢∆C¢7G&ñÊráF&vWBÊ∆Bí¿¢∆ˆ„¢7G&ñÊráF&vWBÊ∆ˆ‚í¿¢˜W&FñˆÊ√¢˜FñˆÁ2Ê˜W&FñˆÊ¬Ú#"¢# ¢“ì∞¢6ˆÁ7B&˜áïW&¬“G∂vWEvVFÜW%&˜áîVÊGˆñÁBÇó”ÚG∑&˜áï&◊2ÁFı7G&ñÊrÇó÷∞¢∆WB&˜áîW'&˜"“ÁV∆√∞†¢G'í∞¢6ˆÁ7B&W7ˆÁ6R“vóBfWF6ÖvVFÜW%&W7ˆÁ6Rá&˜áïW&¬¬˜FñˆÁ2ì∞¢ñbÇ&W7ˆÁ6RÊˆ≤íFá&˜rÊWrW'&˜"Ü&˜áí÷WFVÚÖEEG∑&W7ˆÁ6RÁ7FGW7÷ì∞¢6ˆÁ7BFF“vóB&W7ˆÁ6RÊß6ˆ‚Çì∞¢ñbÜ˜FñˆÁ2Ê˜W&FñˆÊ¬íf∆ñFFTñ◊ñÁFıvVFÜW%ñ∆ˆBÜFF¬FFÁ&˜fñFW"«¬'&˜áí÷WFVÚ"ì∞¢&WGW&‚∞¢‚‚ÊFF¿¢&˜fñFW#¢FFÁ&˜fñFW"«¬$˜V‚‘÷WFVÚ&W7B÷F6Ç"¿¢&˜fñFW%W&√¢FFÁ&˜fñFW%W&¬«¬&áGG3¢Úˆ˜V‚÷÷WFVÚÊ6ˆ“Ú"¿¢÷ˆFV≈6V∆V7Fñˆ„¢FFÊ÷ˆFV≈6V∆V7Fñˆ‚«¬&&W7Eˆ÷F6Ç"¿¢&WVW7EW&√¢&˜áïW&¬¿¢áGG7FGW3¢&W7ˆÁ6RÁ7FGW0¢”∞¢“6F6ÇÜW'&˜"í∞¢&˜áîW'&˜"“W'&˜#∞¢–†¢G'í∞¢6ˆÁ7B&W7ˆÁ6R“vóBfWF6ÖvVFÜW%&W7ˆÁ6RÜFó&V7EW&¬¬˜FñˆÁ2ì∞¢ñbÇ&W7ˆÁ6RÊˆ≤íFá&˜rÊWrW'&˜"Ü˜V‚‘÷WFVÚÖEEG∑&W7ˆÁ6RÁ7FGW7÷ì∞¢6ˆÁ7BFF“vóB&W7ˆÁ6RÊß6ˆ‚Çì∞¢ñbÜ˜FñˆÁ2Ê˜W&FñˆÊ¬íf∆ñFFTñ◊ñÁFıvVFÜW%ñ∆ˆBÜFF¬$˜V‚‘÷WFVÚFó&WGFÚ"ì∞¢&WGW&‚∞¢‚‚ÊFF¿¢&˜fñFW#¢$˜V‚‘÷WFVÚ&W7B÷F6ÇÜFó&WGFÚí"¿¢&˜fñFW%W&√¢&áGG3¢Úˆ˜V‚÷÷WFVÚÊ6ˆ“Ú"¿¢÷ˆFV≈6V∆V7Fñˆ„¢&&W7Eˆ÷F6Ç"¿¢&WVW7EW&√¢Fó&V7EW&¬¿¢áGG7FGW3¢&W7ˆÁ6RÁ7FGW2¿¢&˜áîW'&˜#¢&˜áîW'&˜#ÚÊ÷W76vR«¬%&˜áíÊˆ‚Fó7ˆÊñ&ñ∆R ¢”∞¢“6F6ÇÜFó&V7DW'&˜"í∞¢6ˆÁ7B&˜áî÷W76vR“&˜áîW'&˜#ÚÊ÷W76vR«¬&W'&˜&R&˜áí66ˆÊ˜66óWFÚ#∞¢6ˆÁ7BFó&V7D÷W76vR“Fó&V7DW'&˜#ÚÊ÷W76vR«¬&W'&˜&RFó&WGFÚ66ˆÊ˜66óWFÚ#∞¢Fá&˜rÊWrW'&˜"Ü÷WFVÚÊˆ‚Fó7ˆÊñ&ñ∆S¢G∑&˜áî÷W76vW”≤G∂Fó&V7D÷W76vW÷ì∞¢–ß–†¶gVÊ7Fñˆ‚ó5vVFÜW$FñvÊ˜7Fñ4VÊ&∆VBÇí∞¢6ˆÁ7BVW'í“ÊWrU$≈6V&6Ö&◊2ávñÊF˜rÊ∆ˆ6Fñˆ‚Á6V&6Ç«¬""íÊvWBÇ'vVFÜW$Fñr"ì∞¢6ˆÁ7B∆ˆ6¬“∆ˆ6≈7F˜&vRÊvWDóFV“Ç&ÜW&vVFÜW$FñvÊ˜7Fñ72"ì∞¢&WGW&‚VW'í””“#"«¬∆ˆ6¬””“##∞ß–†¶gVÊ7Fñˆ‚&VÊFW%vVFÜW$FñvÊ˜7Fñ72ÜFñr“∑“í∞¢ñbÇVíÁvVFÜW$FñvÊ˜7Fñ72í&WGW&„∞¢ñbÇó5vVFÜW$FñvÊ˜7Fñ4VÊ&∆VBÇíí∞¢VíÁvVFÜW$FñvÊ˜7Fñ72Ê6∆74∆ó7BÊFBÇ&ÜñFFV‚"ì∞¢VíÁvVFÜW$FñvÊ˜7Fñ72ÁFWáD6ˆÁFVÁB“"#∞¢&WGW&„∞¢–¢6ˆÁ7B∆ñÊW2“∞¢&˜fñFW#¢G∂FñrÁ&˜fñFW"«¬"“'÷¿¢6ˆ˜&FñÊFS¢G∂FñrÊ6ˆ˜&FñÊFW2«¬"“'÷¿¢U$¬ì¢G∂FñrÁW&¬«¬"“'÷¿¢ÖEE¢G∂FñrÊáGG7FGW2ÛÚ"“'÷¿¢W'&˜&S¢G∂FñrÊW'&˜"«¬"“'÷¿¢V«Fñ÷Úvvñ˜&Ê÷VÁFÛ¢G∂FñrÁWFFVDBÚÊWrFFRÜFñrÁWFFVDBíÁFÙ∆ˆ6∆U7G&ñÊrÇ&óB‘ïB"í¢"“'÷ ¢”∞¢VíÁvVFÜW$FñvÊ˜7Fñ72Ê6∆74∆ó7BÁ&V÷˜fRÇ&ÜñFFV‚"ì∞¢VíÁvVFÜW$FñvÊ˜7Fñ72ÁFWáD6ˆÁFVÁB“∆ñÊW2Ê¶ˆñ‚Ç%∆‚"ì∞ß–†¶7ñÊ2gVÊ7Fñˆ‚fWF6ÖvVFÜW"Çí∞¢6ˆÁ7BF&vWB“vWEvVFÜW%F&vWD6ˆ˜&FñÊFW2Çì∞¢7W'&VÁEvVFÜW%F&vWB“F&vWC∞¢&VÊFW$6ófñ≈&˜FV7Fñˆ‰∆W'Bá≤∆WfV√¢&w&VV‚"¬∆&V√¢%fW&ñfñ6&˜FW¶ñˆÊR6ófñ∆R‚‚‚"¬W&√¢4ïdî≈ı$ıDT5DîÙÂÙƒU%EıtR¬∆ˆFñÊs¢G'VR“ì∞†¢6ˆÁ7BFñvÊ˜7Fñ72“∞¢6ˆ˜&FñÊFW3¢G¥ÁV÷&W"áF&vWBÊ∆BíÁFÙfóÜVBÉRó“¬G¥ÁV÷&W"áF&vWBÊ∆ˆ‚íÁFÙfóÜVBÉRó“ÇG∑F&vWBÁ6˜W&6W“ñ¿¢WFFVDC¢FFRÊÊ˜rÇê¢”∞¢G'í∞¢6ˆÁ7BFF“vóBfWF6ÖvVFÜW$f˜&V67BáF&vWBì∞¢FñvÊ˜7Fñ72Á&˜fñFW"“FFÁ&˜fñFW"«¬$˜V‚‘÷WFVÚ&W7B÷F6Ç#∞¢FñvÊ˜7Fñ72ÁW&¬“FFÁ&WVW7EW&¬«¬"“#∞¢FñvÊ˜7Fñ72ÊáGG7FGW2“FFÊáGG7FGW2ÛÚ#∞¢ñbÜFFÁ&˜áîW'&˜"íFñvÊ˜7Fñ72ÊW'&˜"“&˜áì¢G∂FFÁ&˜áîW'&˜'÷∞¢6ˆÁ7B7W'&VÁB“FFÊ7W'&VÁB«¬∑”∞¢6ˆÁ7BvVFÜW$∆&V¬“vVFÜW$6ˆFT∆&V¬Ü7W'&VÁBÁvVFÜW%ˆ6ˆFRì∞¢VíÁvVFÜW%7V÷÷'íÁFWáD6ˆÁFVÁB“G∑vVFÜW$∆&V«“(
"G¥÷FÇÁ&˜VÊBÜ7W'&VÁBÁFV◊W&GW&UÛ&“ÛÚó‹+2(
"fVÁFÚG¥÷FÇÁ&˜VÊBÜ7W'&VÁBÁvñÊE˜7VVEÛ“ÛÚó“∂“ˆÜ∞¢vóB&VÊFW%vVFÜW$FWFñ«2ÜFF¬F&vWBì∞¢&VÊFW%vVFÜW$FñvÊ˜7Fñ72ÜFñvÊ˜7Fñ72ì∞¢“6F6ÇÜW'&˜"í∞¢VíÁvVFÜW%7V÷÷'íÁFWáD6ˆÁFVÁB“$÷WFVÚÊˆ‚Fó7ˆÊñ&ñ∆R‚#∞¢VíÁvVFÜW%&ó6∑2ÊñÊÊW$ÖD‘¬“«7‚6∆73“wvVFÜW"◊&ó6≤÷6ÜósÓ)™˚àÚÊW77V‚FFÚ&ó66ÜñÚFó7ˆÊñ&ñ∆S¬˜7„‚G∂'Vñ∆DÜˆ÷Uv˜&∂∆ñ÷FT'WGFˆ‚á≤F&vWB“ó÷∞¢&ñÊDÜˆ÷Uv˜&∂∆ñ÷FT'WGFˆ‚Çì∞¢&VÊFW$6ófñ≈&˜FV7Fñˆ‰∆W'Bá≤∆WfV√¢&w&VV‚"¬∆&V√¢%&˜FW¶ñˆÊR6ófñ∆RÊˆ‚Fó7ˆÊñ&ñ∆R"¬W&√¢4ïdî≈ı$ıDT5DîÙÂÙƒU%EıtR“ì∞¢VíÁvVFÜW$FWFñ«2ÊñÊÊW$ÖD‘¬“#«6∆73“v◊WFVBs‰ñ◊˜76ñ&ñ∆R6&ñ6&R&Wfó6ñˆÊíFWGFv∆ñFR„¬˜‚#∞¢FñvÊ˜7Fñ72ÊW'&˜"“W'&˜#ÚÊ÷W76vR«¬&W'&˜&R66ˆÊ˜66óWFÚ#∞¢ñbÇFñvÊ˜7Fñ72ÊáGG7FGW2bbÙ4ı%7ƒfñ∆VBFÚfWF6áƒÊWGv˜&¥W'&˜'∆fWF6ÇˆíÁFW7BÜFñvÊ˜7Fñ72ÊW'&˜"ííFñvÊ˜7Fñ72ÊáGG7FGW2“$4ı%2ÙÊWGv˜&≤#∞¢FñvÊ˜7Fñ72ÁWFFVDB“FFRÊÊ˜rÇì∞¢&VÊFW%vVFÜW$FñvÊ˜7Fñ72ÜFñvÊ˜7Fñ72ì∞¢–ß–†¶gVÊ7Fñˆ‚vWEvVFÜW%F&vWD6ˆ˜&FñÊFW2Çí∞¢ñbá6V∆V7FVEvVFÜW$∆ˆ6Fñˆ‚í&WGW&‚≤‚‚Á6V∆V7FVEvVFÜW$∆ˆ6Fñˆ‚¬6˜W&6S¢&÷ÁV¬"”∞¢ñbÜ7W'&VÁEW6W%˜2í&WGW&‚≤∆C¢ÁV÷&W"Ü7W'&VÁEW6W%˜2Ê∆Bí¬∆ˆ„¢ÁV÷&W"Ü7W'&VÁEW6W%˜2Ê∆Êrí¬6˜W&6S¢&w2"”∞¢6ˆÁ7Bw4ñ◊ñÁFí“7W'&VÁDñ◊ñÁFê¢Ê÷ÇÜñ◊ñÁFÚí”‚á≤∆C¢ÁV÷&W"Üñ◊ñÁFÚÊw5íí¬∆ˆ„¢ÁV÷&W"Üñ◊ñÁFÚÊw5Çí“íê¢Êfñ«FW"Çá˜2í”‚ÁV÷&W"Êó4fñÊóFRá˜2Ê∆BíbbÁV÷&W"Êó4fñÊóFRá˜2Ê∆ˆ‚íì∞¢ñbÜw4ñ◊ñÁFíÊ∆VÊwFÇí∞¢6ˆÁ7B7V““w4ñ◊ñÁFíÁ&VGV6RÇÜ62¬˜2í”‚á≤∆C¢62Ê∆B≤˜2Ê∆B¬∆ˆ„¢62Ê∆ˆ‚≤˜2Ê∆ˆ‚“í¬≤∆C¢¬∆ˆ„¢“ì∞¢&WGW&‚≤∆C¢7V“Ê∆BÚw4ñ◊ñÁFíÊ∆VÊwFÇ¬∆ˆ„¢7V“Ê∆ˆ‚Úw4ñ◊ñÁFíÊ∆VÊwFÇ¬6˜W&6S¢&6ˆ÷÷W76"”∞¢–¢&WGW&‚≤∆C¢CB„CìCí¬∆ˆ„¢„3C#b¬6˜W&6S¢&f∆∆&6≤"”∞ß–†¶7ñÊ2gVÊ7Fñˆ‚&VÊFW%vVFÜW$FWFñ«2ÜFF¬F&vWBí∞¢6ˆÁ7BFñ÷W2“ÜFFÊÜ˜W&«íbbFFÊÜ˜W&«íÁFñ÷Rí«¬µ”∞¢6ˆÁ7BFV◊2“ÜFFÊÜ˜W&«íbbFFÊÜ˜W&«íÁFV◊W&GW&UÛ&“í«¬µ”∞¢7W'&VÁDÜˆ÷UvVFÜW$f˜&V67B“∞¢Fñ÷W2¿¢FV◊2¿¢6ˆFW3¢ÜFFÊÜ˜W&«íbbFFÊÜ˜W&«íÁvVFÜW%ˆ6ˆFRí«¬µ“¿¢&ñÁ3¢ÜFFÊÜ˜W&«íbbFFÊÜ˜W&«íÁ&V6óóFFñˆÂ˜&ˆ&&ñ∆óGíí«¬µ“¿¢vñÊG3¢ÜFFÊÜ˜W&«íbbFFÊÜ˜W&«íÁvñÊE˜7VVEÛ“í«¬µ“¿¢WFFVDC¢FFRÊÊ˜rÇê¢”∞¢6ˆÁ7B&ñÁ2“ÜFFÊÜ˜W&«íbbFFÊÜ˜W&«íÁ&V6óóFFñˆÂ˜&ˆ&&ñ∆óGíí«¬µ”∞¢6ˆÁ7B6Ê˜w2“ÜFFÊÜ˜W&«íbbFFÊÜ˜W&«íÁ6Ê˜vf∆¬í«¬µ”∞¢6ˆÁ7Bfó6ñ&ñ∆óFñW2“ÜFFÊÜ˜W&«íbbFFÊÜ˜W&«íÁfó6ñ&ñ∆óGíí«¬µ”∞¢6ˆÁ7B6ˆFW2“ÜFFÊÜ˜W&«íbbFFÊÜ˜W&«íÁvVFÜW%ˆ6ˆFRí«¬µ”∞¢6ˆÁ7BvñÊG2“ÜFFÊÜ˜W&«íbbFFÊÜ˜W&«íÁvñÊE˜7VVEÛ“í«¬µ”∞¢6ˆÁ7B÷Ö&ñ‚“÷FÇÊ÷ÇÇ‚‚Á&ñÁ2Á6∆ñ6RÉ¬"íÊ÷Çáf«VRí”‚ÁV÷&W"áf«VRí«¬í¬ì∞¢6ˆÁ7B6Ê˜u7V““6Ê˜w2Á6∆ñ6RÉ¬"íÁ&VGV6RÇÜ62¬f«VRí”‚62≤ÑÁV÷&W"áf«VRí«¬í¬ì∞¢6ˆÁ7B÷ñÂfó6ñ&ñ∆óGí“÷FÇÊ÷ñ‚Ç‚‚Áfó6ñ&ñ∆óFñW2Á6∆ñ6RÉ¬"íÊ÷Çáf«VRí”‚ÁV÷&W"áf«VRí«¬ÁV÷&W"‰‘Öı4dUÙîÂDTtU"íì∞¢6ˆÁ7BÜ4fˆt6ˆFR“6ˆFW2Á6∆ñ6RÉ¬"íÁ6ˆ÷RÇáf«VRí”‚ÁV÷&W"áf«VRí””“CR«¬ÁV÷&W"áf«VRí””“CÇì∞¢6ˆÁ7B&ó6¥ñ6R“FV◊2Á6∆ñ6RÉ¬"íÁ6ˆ÷RÇáf«VR¬ñGÇí”‚ÁV÷&W"áf«VRí√“bbÁV÷&W"á&ñÁ5∂ñGÖ“«¬í„“Cì∞†¢6ˆÁ7B&ó6∑2“µ”∞¢&ó6∑2ÁW6ÇÜ÷Ö&ñ‚„“cÚ/	¯ ~˚àÚ&ó66ÜñÚñˆvvñ«F"¢/	¯ ~˚àÚ&ó66ÜñÚñˆvvñ&76"ì∞¢ñbá6Ê˜u7V“‚í&ó6∑2ÁW6ÇÇ.)ÿN˚àÚ˜76ñ&ñ∆RÊWfR"ì∞¢ñbÜÜ4fˆt6ˆFR«¬÷ñÂfó6ñ&ñ∆óGí¬#í&ó6∑2ÁW6ÇÇ/	¯ æ˚àÚ˜76ñ&ñ∆RÊV&&ñ"ì∞¢ñbá&ó6¥ñ6Rí&ó6∑2ÁW6ÇÇ/	˙x¢˜76ñ&ñ∆RvÜñ66ñÚ"ì∞†¢6ˆÁ7B∆W'B“vóBvWD6ófñ≈&˜FV7Fñˆ‰∆W'BáF&vWB¬≤FV◊2¬vñÊG2¬6Ê˜w2¬fó6ñ&ñ∆óFñW2¬6ˆFW2“ì∞¢6ˆÁ7B&ó6¥6Üó2“&ó6∑2Ê÷Çá&ó6≤í”‚«7‚6∆73“wvVFÜW"◊&ó6≤÷6Üós‚G∂W66TÖD‘¬á&ó6≤ó”¬˜7„ÊíÊ¶ˆñ‚Ç""ì∞¢VíÁvVFÜW%&ó6∑2ÊñÊÊW$ÖD‘¬“G∑&ó6¥6Üó7“G∂'Vñ∆D6ófñ≈&˜FV7Fñˆ‰∆W'D6ÜóÜ∆W'Bó“G∂'Vñ∆DÜˆ÷Uv˜&∂∆ñ÷FT'WGFˆ‚á≤FV◊2¬F&vWB“ó÷∞¢&ñÊDÜˆ÷Uv˜&∂∆ñ÷FT'WGFˆ‚Çì∞¢&VÊFW$6ófñ≈&˜FV7Fñˆ‰∆W'BÜ∆W'Bì∞†¢6ˆÁ7B&˜w2“Fñ÷W2Á6∆ñ6RÉ¬"íÊ÷ÇáFñ÷R¬ñGÇí”‚∞¢6ˆÁ7BÜ˜W"“ÊWrFFRáFñ÷RíÁFÙ∆ˆ6∆UFñ÷U7G&ñÊrÇ&óB‘ïB"¬≤Ü˜W#¢#"÷FñvóB"¬÷ñÁWFS¢#"÷FñvóB"¬Ü˜W##¢f«6R“ì∞¢6ˆÁ7Bfó4∂““ÇÑÁV÷&W"áfó6ñ&ñ∆óFñW5∂ñGÖ“í«¬íÚíÁFÙfóÜVBÉì∞¢6ˆÁ7B∆&V¬“vVFÜW$6ˆFT∆&V¬Ü6ˆFW5∂ñGÖ“ì∞¢&WGW&‚«„∆#‚G∂Ü˜W'”¬ˆ#‚(
"G∂∆&V«“(
"	¯ ˚àÚG¥÷FÇÁ&˜VÊBáFV◊5∂ñGÖ“ÛÚó‹+2(
"	¯ ~˚àÚG¥÷FÇÁ&˜VÊBá&ñÁ5∂ñGÖ“ÛÚó“R(
")ÿN˚àÚG¥ÁV÷&W"á6Ê˜w5∂ñGÖ“«¬íÁFÙfóÜVBÉó“÷“(
"	˘˚àÚG∑fó4∂◊“∂”¬˜Ê∞¢“íÊ¶ˆñ‚Ç""ì∞¢6ˆÁ7B&˜fñFW$Ê÷R“FFÁ&˜fñFW"«¬$˜V‚‘÷WFVÚ&W7B÷F6Ç#∞¢6ˆÁ7B&˜fñFW%W&¬“FFÁ&˜fñFW%W&¬«¬&áGG3¢Úˆ˜V‚÷÷WFVÚÊ6ˆ“Ú#∞¢6ˆÁ7BGG&ñ'WFñˆ‚“«6∆73“&◊WFVBvVFÜW"÷FF◊6˜W&6R#‰FFì¢∆á&Vc“"G∂W66TÖD‘¬á&˜fñFW%W&¬ó“"F&vWC“%ˆ&∆Ê≤"&V√“&Êˆ˜VÊW"Ê˜&VfW'&W"#‚G∂W66TÖD‘¬á&˜fñFW$Ê÷Ró”¬ˆ„¬˜Ê∞¢VíÁvVFÜW$FWFñ«2ÊñÊÊW$ÖD‘¬“G∑&˜w2«¬#«6∆73“v◊WFVBs‰ÊW77V‚FFÚ÷WFVÚ„¬˜‚'“G∂GG&ñ'WFñˆÁ÷∞ß–†¶7ñÊ2gVÊ7Fñˆ‚vWD6ófñ≈&˜FV7Fñˆ‰∆W'BáF&vWB¬f˜&V67B“∑“í∞¢6ˆÁ7B&Vvñˆ‚“vóB&WfW'6TvVˆ6ˆFU&Vvñˆ‚áF&vWBíÊ6F6ÇÇÇí”‚""ì∞¢6ˆÁ7Bˆffñ6ñ≈FWáB“vóBfWF6Ñ6ófñ≈&˜FV7Fñˆ‰ˆffñ6ñ≈FWáBÇíÊ6F6ÇÇÇí”‚""ì∞¢6ˆÁ7Bˆffñ6ñƒ∆W'B“'6T6ófñ≈&˜FV7Fñˆ‰∆W'EFWáBÜˆffñ6ñ≈FWáB¬&Vvñˆ‚ì∞¢6ˆÁ7Bf˜&V67D∆W'B“'Vñ∆D˜W&FñˆÊƒf˜&V67D∆W'BÜf˜&V67Bì∞¢6ˆÁ7B∆W'B“ñ6¥ÜñvÜW7D∆W'BÖ∂ˆffñ6ñƒ∆W'B¬f˜&V67D∆W'E“ì∞¢&WGW&‚∞¢‚‚Ê∆W'B¿¢&Vvñˆ‚¿¢W&√¢4ïdî≈ı$ıDT5DîÙÂÙƒU%EıtR¿¢∆&V√¢∆W'BÊ∆&V¬«¬$ÊW77VÊ∆∆W'F ¢”∞ß–†¶7ñÊ2gVÊ7Fñˆ‚&WfW'6TvVˆ6ˆFU&Vvñˆ‚áF&vWBí∞¢6ˆÁ7B∂Wí“ÜW&vVFÜW%&Vvñˆ„¢G∑F&vWBÊ∆BÁFÙfóÜVBÉ"ó”¢G∑F&vWBÊ∆ˆ‚ÁFÙfóÜVBÉ"ó÷∞¢6ˆÁ7B66ÜVB“∆ˆ6≈7F˜&vRÊvWDóFV“Ü∂Wíì∞¢ñbÜ66ÜVBí&WGW&‚66ÜVC∞¢6ˆÁ7BW&¬“áGG3¢ÚˆÊˆ÷ñÊFñ“Ê˜VÁ7G&VWF÷Ê˜&r˜&WfW'6Sˆf˜&÷C÷ß6ˆÁc"f∆C“G∂VÊ6ˆFUU$î6ˆ◊ˆÊVÁBáF&vWBÊ∆Bó“f∆ˆ„“G∂VÊ6ˆFUU$î6ˆ◊ˆÊVÁBáF&vWBÊ∆ˆ‚ó“g¶ˆˆ””ÇfFG&W76FWFñ«3”∞¢6ˆÁ7B&W7ˆÁ6R“vóBfWF6ÇáW&¬¬≤ÜVFW'3¢≤66WC¢&∆ñ6Fñˆ‚ˆß6ˆ‚"““ì∞¢ñbÇ&W7ˆÁ6RÊˆ≤í&WGW&‚"#∞¢6ˆÁ7BFF“vóB&W7ˆÁ6RÊß6ˆ‚Çì∞¢6ˆÁ7B&Vvñˆ‚“Ê˜&÷∆ó¶TóF∆ñÂ&Vvñˆ‰Ê÷RÜFFÚÊFG&W73ÚÁ7FFR«¬FFÚÊFG&W73ÚÁ&Vvñˆ‚«¬""ì∞¢ñbá&Vvñˆ‚í∆ˆ6≈7F˜&vRÁ6WDóFV“Ü∂Wí¬&Vvñˆ‚ì∞¢&WGW&‚&Vvñˆ„∞ß–†¶gVÊ7Fñˆ‚Ê˜&÷∆ó¶TóF∆ñÂ&Vvñˆ‰Ê÷Rá&Vvñˆ‚í∞¢&WGW&‚7G&ñÊrá&Vvñˆ‚«¬""ê¢Á&W∆6RÇıÁ&VvñˆÊU«2≤ˆí¬""ê¢Á&W∆6RÇˆV÷ñ∆ñ◊&ˆ÷vÊˆí¬$V÷ñ∆ñ&ˆ÷vÊ"ê¢Á&W∆6RÇ˜G&VÁFñÊÚ÷«FÚFñvU¬˜<;∆GFó&ˆ¬ˆí¬%G&VÁFñÊÚ«FÚFñvR"ê¢ÁG&ñ“Çì∞ß–†¶7ñÊ2gVÊ7Fñˆ‚fWF6Ñ6ófñ≈&˜FV7Fñˆ‰ˆffñ6ñ≈FWáBÇí∞¢6ˆÁ7B66ÜVB“vWD66ÜVD6ófñ≈&˜FV7FñˆÂFWáBÇì∞¢ñbÜ66ÜVBí&WGW&‚66ÜVC∞¢6ˆÁ7BvóFáV%FWáB“vóBfWF6Ñ∆FW7D6ófñ≈&˜FV7FñˆÂÜ÷≈FWáBÇíÊ6F6ÇÇÇí”‚""ì∞¢ñbÜvóFáV%FWáBí∞¢66ÜT6ófñ≈&˜FV7FñˆÂFWáBÜvóFáV%FWáBì∞¢&WGW&‚vóFáV%FWáC∞¢–¢6ˆÁ7B&W7ˆÁ6R“vóBfWF6ÇÑ4ïdî≈ı$ıDT5DîÙÂÙƒU%EıtR¬≤66ÜS¢&ÊÚ◊7F˜&R"“ì∞¢ñbÇ&W7ˆÁ6RÊˆ≤í&WGW&‚"#∞¢6ˆÁ7BáF÷¬“vóB&W7ˆÁ6RÁFWáBÇì∞¢6ˆÁ7BFˆ2“ÊWrDÙ’'6W"ÇíÁ'6Tg&ˆ’7G&ñÊrÜáF÷¬¬'FWáBˆáF÷¬"ì∞¢6ˆÁ7BFWáB“Fˆ2Ê&ˆGìÚÁFWáD6ˆÁFVÁB«¬áF÷√∞¢66ÜT6ófñ≈&˜FV7FñˆÂFWáBáFWáBì∞¢&WGW&‚FWáC∞ß–†¶7ñÊ2gVÊ7Fñˆ‚fWF6Ñ∆FW7D6ófñ≈&˜FV7FñˆÂÜ÷≈FWáBÇí∞¢6ˆÁ7B&W7ˆÁ6R“vóBfWF6ÇÑ4ïdî≈ı$ıDT5DîÙÂÙtïDÖT%Ùí¬≤ÜVFW'3¢≤66WC¢&∆ñ6Fñˆ‚˜fÊBÊvóFáV"∂ß6ˆ‚"““ì∞¢ñbÇ&W7ˆÁ6RÊˆ≤í&WGW&‚"#∞¢6ˆÁ7Bfñ∆W2“vóB&W7ˆÁ6RÊß6ˆ‚Çì∞¢6ˆÁ7B∆FW7B“Ñ'&íÊó4'&íÜfñ∆W2íÚfñ∆W2¢µ“ê¢Êfñ«FW"ÇÜfñ∆Rí”‚7G&ñÊrÜfñ∆RÊÊ÷R«¬""íÁFÙ∆˜vW$66RÇíÊVÊG5vóFÇÇ"ÁÜ÷¬"íbbfñ∆RÊF˜vÊ∆ˆE˜W&¬ê¢Á6˜'BÇÜ¬"í”‚7G&ñÊrÜ"ÊÊ÷RíÊ∆ˆ6∆T6ˆ◊&RÖ7G&ñÊrÜÊÊ÷Rííê¢ÊBÉì∞¢ñbÇ∆FW7Bí&WGW&‚"#∞¢6ˆÁ7BÜ÷≈&W7ˆÁ6R“vóBfWF6ÇÜ∆FW7BÊF˜vÊ∆ˆE˜W&¬¬≤66ÜS¢&ÊÚ◊7F˜&R"“ì∞¢&WGW&‚Ü÷≈&W7ˆÁ6RÊˆ≤ÚÜ÷≈&W7ˆÁ6RÁFWáBÇí¢"#∞ß–†¶gVÊ7Fñˆ‚vWD66ÜVD6ófñ≈&˜FV7FñˆÂFWáBÇí∞¢G'í∞¢6ˆÁ7B66ÜVB“•4Ù‚Á'6RÜ∆ˆ6≈7F˜&vRÊvWDóFV“Ç&ÜW&6ófñ≈&˜FV7Fñˆ‰∆W'EFWáB"í«¬&ÁV∆¬"ì∞¢ñbÜ66ÜVBbbFFRÊÊ˜rÇí“ÁV÷&W"Ü66ÜVBÁ6fVDB«¬í¬c¢c¢í&WGW&‚7G&ñÊrÜ66ÜVBÁFWáB«¬""ì∞¢“6F6ÇÜW'&˜"í∞¢6ˆÁ6ˆ∆RÁv&‚Ç$66ÜR&˜FW¶ñˆÊR6ófñ∆RÊˆ‚∆Vvvñ&ñ∆S¢"¬W'&˜"ì∞¢–¢&WGW&‚"#∞ß–†¶gVÊ7Fñˆ‚66ÜT6ófñ≈&˜FV7FñˆÂFWáBáFWáBí∞¢G'í∞¢∆ˆ6≈7F˜&vRÁ6WDóFV“Ç&ÜW&6ófñ≈&˜FV7Fñˆ‰∆W'EFWáB"¬•4Ù‚Á7G&ñÊvñgíá≤FWáB¬6fVDC¢FFRÊÊ˜rÇí“íì∞¢“6F6ÇÜW'&˜"í∞¢6ˆÁ6ˆ∆RÁv&‚Ç$66ÜR&˜FW¶ñˆÊR6ófñ∆RÊˆ‚6«fF¢"¬W'&˜"ì∞¢–ß–†¶gVÊ7Fñˆ‚'6T6ófñ≈&˜FV7Fñˆ‰∆W'EFWáBáFWáB¬&Vvñˆ‚í∞¢6ˆÁ7BÊ˜&÷∆ó¶VEFWáB“7G&ñÊráFWáB«¬""íÁ&W∆6RÇÛ≈µ„Â“≥‚ˆr¬%∆‚"ì∞¢6ˆÁ7B∆ñÊW2“Ê˜&÷∆ó¶VEFWáBÁ7∆óBÇı«#ı∆‚ÚíÊ÷ÇÜ∆ñÊRí”‚∆ñÊRÁ&W∆6RÇı«2≤ˆr¬""íÁG&ñ“ÇííÊfñ«FW"Ñ&ˆˆ∆V‚ì∞¢6ˆÁ7B&Vvñˆ‰ÊVVF∆R“Ê˜&÷∆ó¶Tf˜%6V&6Çá&Vvñˆ‚ì∞¢∆WB7W'&VÁD∆WfV¬“&w&VV‚#∞¢∆WB7W'&VÁEÜVÊˆ÷VÊˆ‚“%&˜FW¶ñˆÊR6ófñ∆R#∞¢6ˆÁ7B∆W'G2“µ”∞†¢∆ñÊW2Êf˜$V6ÇÇÜ∆ñÊRí”‚∞¢6ˆÁ7BWW"“∆ñÊRÁFıWW$66RÇì∞¢6ˆÁ7B∆WfV¬“∆WfVƒg&ˆ’FWáBáWW"ì∞¢ñbáWW"ÊñÊ6«VFW2Ç$ƒƒU%D"í«¬WW"ÊñÊ6«VFW2Ç$5$ïDî4ïD"íí∞¢7W'&VÁD∆WfV¬“∆WfV¬«¬7W'&VÁD∆WfV√∞¢7W'&VÁEÜVÊˆ÷VÊˆ‚“ÜVÊˆ÷VÊˆ‰g&ˆ’FWáBÜ∆ñÊRí«¬7W'&VÁEÜVÊˆ÷VÊˆ„∞¢–¢6ˆÁ7B∆ñÊTÜ5&Vvñˆ‚“&Vvñˆ‰ÊVVF∆RbbÊ˜&÷∆ó¶Tf˜%6V&6ÇÜ∆ñÊRíÊñÊ6«VFW2á&Vvñˆ‰ÊVVF∆Rì∞¢6ˆÁ7Bó4ÊFñˆÊƒÊÙ∆W'B“WW"ÊñÊ6«VFW2Ç$‰U55T‰ƒƒU%D"í«¬WW"ÊñÊ6«VFW2Ç$54TÂ§DídT‰Ù‘T‰í4ît‰îdî4Dïdí"ì∞¢ñbÜ∆ñÊTÜ5&Vvñˆ‚bbƒU%EÙƒUdT≈Ù‘UD∂7W'&VÁD∆WfV≈”ÚÁ&Ê≤‚í∞¢∆W'G2ÁW6Çá≤∆WfV√¢7W'&VÁD∆WfV¬¬∆&V√¢'Vñ∆D∆W'D∆&V¬Ü7W'&VÁD∆WfV¬¬7W'&VÁEÜVÊˆ÷VÊˆ‚í¬ÜVÊˆ÷VÊˆ„¢7W'&VÁEÜVÊˆ÷VÊˆ‚“ì∞¢“V«6RñbÇ&Vvñˆ‰ÊVVF∆Rbbó4ÊFñˆÊƒÊÙ∆W'Bí∞¢∆W'G2ÁW6Çá≤∆WfV√¢&w&VV‚"¬∆&V√¢$ÊW77VÊ∆∆W'F"¬ÜVÊˆ÷VÊˆ„¢""“ì∞¢–¢“ì∞†¢&WGW&‚ñ6¥ÜñvÜW7D∆W'BÜ∆W'G2Ê∆VÊwFÇÚ∆W'G2¢∑≤∆WfV√¢&w&VV‚"¬∆&V√¢$ÊW77VÊ∆∆W'F"¬ÜVÊˆ÷VÊˆ„¢""’“ì∞ß–†¶gVÊ7Fñˆ‚Ê˜&÷∆ó¶Tf˜%6V&6Çáf«VRí∞¢&WGW&‚7G&ñÊráf«VR«¬""íÊÊ˜&÷∆ó¶RÇ$‰dB"íÁ&W∆6RÇıµ«S3’«S3fe“ˆr¬""íÁFÙ∆˜vW$66RÇì∞ß–†¶gVÊ7Fñˆ‚∆WfVƒg&ˆ’FWáBáFWáBí∞¢ñbáFWáBÊñÊ6«VFW2Ç%$ı54"í«¬FWáBÊñÊ6«VFW2Ç$TƒUdD"íí&WGW&‚'&VB#∞¢ñbáFWáBÊñÊ6«VFW2Ç$$‰4îÙ‰R"í«¬FWáBÊñÊ6«VFW2Ç$‘ÙDU$D"íí&WGW&‚&˜&ÊvR#∞¢ñbáFWáBÊñÊ6«VFW2Ç$tîƒƒ"í«¬FWáBÊñÊ6«VFW2Ç$ı$Dî‰$î"íí&WGW&‚'ñV∆∆˜r#∞¢ñbáFWáBÊñÊ6«VFW2Ç%dU$DR"í«¬FWáBÊñÊ6«VFW2Ç$‰U55T‰ƒƒU%D"íí&WGW&‚&w&VV‚#∞¢&WGW&‚"#∞ß–†¶gVÊ7Fñˆ‚ÜVÊˆ÷VÊˆ‰g&ˆ’FWáBáFWáBí∞¢6ˆÁ7BÊ˜&÷∆ó¶VB“Ê˜&÷∆ó¶Tf˜%6V&6ÇáFWáBì∞¢6ˆÁ7B÷F6Ç“ƒU%EÙ¥Uïtı$E2ÊfñÊBÇÜóFV“í”‚óFV“ÁGFW&Á2Á6ˆ÷RÇáGFW&‚í”‚Ê˜&÷∆ó¶VBÊñÊ6«VFW2ÜÊ˜&÷∆ó¶Tf˜%6V&6ÇáGFW&‚íííì∞¢&WGW&‚÷F6ÉÚÊ∆&V¬«¬%&˜FW¶ñˆÊR6ófñ∆R#∞ß–†¶gVÊ7Fñˆ‚'Vñ∆D˜W&FñˆÊƒf˜&V67D∆W'Bá≤FV◊2“µ“¬vñÊG2“µ“¬6Ê˜w2“µ“¬fó6ñ&ñ∆óFñW2“µ“¬6ˆFW2“µ“““∑“í∞¢6ˆÁ7B÷ÖvñÊB“÷FÇÊ÷ÇÇ‚‚ÁvñÊG2Á6∆ñ6RÉ¬"íÊ÷Çáf«VRí”‚ÁV÷&W"áf«VRí«¬í¬ì∞¢6ˆÁ7B6Ê˜u7V““6Ê˜w2Á6∆ñ6RÉ¬"íÁ&VGV6RÇÜ62¬f«VRí”‚62≤ÑÁV÷&W"áf«VRí«¬í¬ì∞¢6ˆÁ7B÷ñÂFV◊“÷FÇÊ÷ñ‚Ç‚‚ÁFV◊2Á6∆ñ6RÉ¬"íÊ÷Çáf«VRí”‚ÁV÷&W"áf«VRí«¬ÁV÷&W"‰‘Öı4dUÙîÂDTtU"íì∞¢6ˆÁ7B÷ÖFV◊“÷FÇÊ÷ÇÇ‚‚ÁFV◊2Á6∆ñ6RÉ¬"íÊ÷Çáf«VRí”‚ÁV÷&W"áf«VRí«¬”í¬”ì∞¢6ˆÁ7B÷ñÂfó6ñ&ñ∆óGí“÷FÇÊ÷ñ‚Ç‚‚Áfó6ñ&ñ∆óFñW2Á6∆ñ6RÉ¬"íÊ÷Çáf«VRí”‚ÁV÷&W"áf«VRí«¬ÁV÷&W"‰‘Öı4dUÙîÂDTtU"íì∞¢6ˆÁ7BÜ57F˜&‘6ˆFR“6ˆFW2Á6∆ñ6RÉ¬"íÁ6ˆ÷RÇáf«VRí”‚≥ìR¬ìb¬ìï“ÊñÊ6«VFW2ÑÁV÷&W"áf«VRííì∞¢ñbÜ÷ÖvñÊB„“sRí&WGW&‚≤∆WfV√¢&˜&ÊvR"¬∆&V√¢'Vñ∆D∆W'D∆&V¬Ç&˜&ÊvR"¬'fVÁFÚ"í¬ÜVÊˆ÷VÊˆ„¢'fVÁFÚ"”∞¢ñbá6Ê˜u7V“„“#í&WGW&‚≤∆WfV√¢&˜&ÊvR"¬∆&V√¢'Vñ∆D∆W'D∆&V¬Ç&˜&ÊvR"¬&ÊWfR"í¬ÜVÊˆ÷VÊˆ„¢&ÊWfR"”∞¢ñbÜÜ57F˜&‘6ˆFRí&WGW&‚≤∆WfV√¢'ñV∆∆˜r"¬∆&V√¢'Vñ∆D∆W'D∆&V¬Ç'ñV∆∆˜r"¬%FV◊˜&∆í"í¬ÜVÊˆ÷VÊˆ„¢%FV◊˜&∆í"”∞¢ñbÜ÷ñÂFV◊√“”"í&WGW&‚≤∆WfV√¢'ñV∆∆˜r"¬∆&V√¢'Vñ∆D∆W'D∆&V¬Ç'ñV∆∆˜r"¬&vÜñ66ñÚ"í¬ÜVÊˆ÷VÊˆ„¢&vÜñ66ñÚ"”∞¢ñbÜ÷ñÂfó6ñ&ñ∆óGí¬Sí&WGW&‚≤∆WfV√¢'ñV∆∆˜r"¬∆&V√¢'Vñ∆D∆W'D∆&V¬Ç'ñV∆∆˜r"¬&ÊV&&ñ"í¬ÜVÊˆ÷VÊˆ„¢&ÊV&&ñ"”∞¢ñbÜ÷ÖFV◊„“3Rí&WGW&‚≤∆WfV√¢'&VB"¬∆&V√¢'Vñ∆D∆W'D∆&V¬Ç'&VB"¬&6∆FÚW7G&V÷Ú"í¬ÜVÊˆ÷VÊˆ„¢&6∆FÚW7G&V÷Ú"”∞¢ñbÜ÷ÖFV◊„“3"í&WGW&‚≤∆WfV√¢&˜&ÊvR"¬∆&V√¢'Vñ∆D∆W'D∆&V¬Ç&˜&ÊvR"¬&6∆FÚñÁFVÁ6Ú"í¬ÜVÊˆ÷VÊˆ„¢&6∆FÚñÁFVÁ6Ú"”∞¢ñbÜ÷ÖFV◊„“3í&WGW&‚≤∆WfV√¢'ñV∆∆˜r"¬∆&V√¢'Vñ∆D∆W'D∆&V¬Ç'ñV∆∆˜r"¬&6∆FÚ"í¬ÜVÊˆ÷VÊˆ„¢&6∆FÚ"”∞¢&WGW&‚≤∆WfV√¢&w&VV‚"¬∆&V√¢$ÊW77VÊ∆∆W'F"¬ÜVÊˆ÷VÊˆ„¢""”∞ß–†¶gVÊ7Fñˆ‚ñ6¥ÜñvÜW7D∆W'BÜ∆W'G2í∞¢&WGW&‚∆W'G2Á&VGV6RÇÜ&W7B¬∆W'Bí”‚∞¢6ˆÁ7B∆WfV¬“∆W'CÚÊ∆WfV¬«¬&w&VV‚#∞¢&WGW&‚ƒU%EÙƒUdT≈Ù‘UD∂∆WfV≈“Á&Ê≤‚ƒU%EÙƒUdT≈Ù‘UD∂&W7BÊ∆WfV≈“Á&Ê≤Ú∆W'B¢&W7C∞¢“¬≤∆WfV√¢&w&VV‚"¬∆&V√¢$ÊW77VÊ∆∆W'F"¬ÜVÊˆ÷VÊˆ„¢""“ì∞ß–†¶gVÊ7Fñˆ‚'Vñ∆D∆W'D∆&V¬Ü∆WfV¬¬ÜVÊˆ÷VÊˆ‚í∞¢ñbÜ∆WfV¬””“&w&VV‚"í&WGW&‚$ÊW77VÊ∆∆W'F#∞¢&WGW&‚ÜVÊˆ÷VÊˆ‚bbÜVÊˆ÷VÊˆ‚”“%&˜FW¶ñˆÊR6ófñ∆R"Ú∆∆W'FG∑ÜVÊˆ÷VÊˆÁ÷¢$∆∆W'F&˜FW¶ñˆÊR6ófñ∆R#∞ß–†¶gVÊ7Fñˆ‚vWDÜˆ÷Uv˜&∂∆ñ÷FU&ó6¥∆WfV¬áFV◊2“µ“í∞¢6ˆÁ7B÷ÖFV◊“÷FÇÊ÷ÇÇ‚‚ÁFV◊2Á6∆ñ6RÉ¬"íÊ÷Çáf«VRí”‚ÁV÷&W"áf«VRí«¬”í¬”ì∞¢ñbÜ÷ÖFV◊„“3Rí&WGW&‚'&˜76Ú#∞¢ñbÜ÷ÖFV◊„“3"í&WGW&‚&&Ê6ñˆÊR#∞¢ñbÜ÷ÖFV◊„“3í&WGW&‚&vñ∆∆Ú#∞¢&WGW&‚'fW&FR#∞ß–†¶gVÊ7Fñˆ‚vWDÜˆ÷Uv˜&∂∆ñ÷FT'WGFˆ‰∆&V¬Çí∞¢&WGW&‚'v˜&∂∆ñ÷FR#∞ß–†¶gVÊ7Fñˆ‚'Vñ∆DÜˆ÷Uv˜&∂∆ñ÷FT'WGFˆ‚á≤FV◊2“µ“¬F&vWB“ÁV∆¬““∑“í∞¢6ˆÁ7B&ó6¥∆WfV¬“vWDÜˆ÷Uv˜&∂∆ñ÷FU&ó6¥∆WfV¬áFV◊2ì∞¢6ˆÁ7B∆&V¬“vWDÜˆ÷Uv˜&∂∆ñ÷FT'WGFˆ‰∆&V¬á&ó6¥∆WfV¬¬F&vWBì∞¢&WGW&‚∆'WGFˆ‚GóS“&'WGFˆ‚"6∆73“'vVFÜW"◊&ó6≤÷6ÜóÜˆ÷R◊v˜&∂∆ñ÷FR÷'F‚&ó6≤“G∑&ó6¥∆WfV«“"FF÷Üˆ÷R◊v˜&∂∆ñ÷FR◊W&√“"G∂W66TÖD‘¬Ötı$¥ƒî‘DUÙdı$T45EıU$¬ó“"FF÷Üˆ÷R◊v˜&∂∆ñ÷FR◊&ó6≥“"G∂W66TÖD‘¬á&ó6¥∆WfV¬ó“"&ñ÷∆&V√“$&í&6ÜV6v˜&∂∆ñ÷FR#‚G∂W66TÖD‘¬Ü∆&V¬ó”¬ˆ'WGFˆ„Ê∞ß–†¶gVÊ7Fñˆ‚&ñÊDÜˆ÷Uv˜&∂∆ñ÷FT'WGFˆ‚Çí∞¢VíÁvVFÜW%&ó6∑3ÚÁVW'ï6V∆V7F˜"Ç%∂FF÷Üˆ÷R◊v˜&∂∆ñ÷FR◊W&≈“"ìÚÊFDWfVÁD∆ó7FVÊW"Ç&6∆ñ6≤"¬ÜWfVÁBí”‚∞¢WfVÁBÁ&WfVÁDFVfV«BÇì∞¢WfVÁBÁ7F˜&˜vFñˆ‚Çì∞¢6ˆÁ7B'WGFˆ‚“WfVÁBÊ7W'&VÁEF&vWC∞¢6ˆÁ7BW&¬“'WGFˆ‚ÊvWDGG&ñ'WFRÇ&FF÷Üˆ÷R◊v˜&∂∆ñ÷FR◊W&¬"í«¬tı$¥ƒî‘DUÙdı$T45EıU$√∞¢6ˆÁ7B&ó6¥∆WfV¬“'WGFˆ‚ÊvWDGG&ñ'WFRÇ&FF÷Üˆ÷R◊v˜&∂∆ñ÷FR◊&ó6≤"í«¬'fW&FR#∞¢˜V‰Üˆ÷Uv˜&∂∆ñ÷FT&ˆ&Bá≤&ó6¥∆WfV¬¬W&¬“ì∞¢“ì∞ß–†¶gVÊ7Fñˆ‚f˜&÷DÜVDF6Ü&ˆ&DFFRÜFFRí∞¢ñbÇÜFFRñÁ7FÊ6VˆbFFRí«¬ÁV÷&W"Êó4Ê‚ÜFFRÊvWEFñ÷RÇííí&WGW&‚"“#∞¢6ˆÁ7BFí“FFRÁFÙ∆ˆ6∆TFFU7G&ñÊrÇ&óB‘ïB"¬≤vVV∂Fì¢'6Ü˜'B"“ì∞¢6ˆÁ7BÁV÷W&ñ2“FFRÁFÙ∆ˆ6∆TFFU7G&ñÊrÇ&óB‘ïB"¬≤Fì¢#"÷FñvóB"¬÷ˆÁFÉ¢#"÷FñvóB"“ì∞¢&WGW&‚G∂FíÊ6Ü$BÉíÁFıWW$66RÇó“G∂FíÁ6∆ñ6RÉó“G∂ÁV÷W&ñ7÷∞ß–†¶gVÊ7Fñˆ‚vWDÜVE&ó6µ7VvvW7Fñˆ‚Ü∆WfV¬í∞¢6ˆÁ7BÊ˜&÷∆ó¶VB“Ê˜&÷∆ó¶Uv˜&∂∆ñ÷FT∆WfV¬Ü∆WfV¬ì∞¢ñbÜÊ˜&÷∆ó¶VB””“'&˜76Ú"í&WGW&‚$WfóF&RGFófóL:W6ÁFíÊV∆∆R˜&R6VÁG&∆í#∞¢ñbÜÊ˜&÷∆ó¶VB””“&&Ê6ñˆÊR"í&WGW&‚%&ˆw&÷÷&RW6RvvóVÁFófR#∞¢ñbÜÊ˜&÷∆ó¶VB””“&vñ∆∆Ú"í&WGW&‚$&W&R7W76ÚRf&RW6R'&Wfí#∞¢&WGW&‚$6ˆÊFó¶ñˆÊí˜&FñÊ&ñR¬÷ÁFVÊW&RñG&F¶ñˆÊR#∞ß–†¶gVÊ7Fñˆ‚vWDÜVE&ó6¥g&ˆ’FV◊Ü÷ÖFV◊í∞¢ñbÇÁV÷&W"Êó4fñÊóFRÑÁV÷&W"Ü÷ÖFV◊ííí&WGW&‚'fW&FR#∞¢&WGW&‚vWDÜˆ÷Uv˜&∂∆ñ÷FU&ó6¥∆WfV¬Ö¥ÁV÷&W"Ü÷ÖFV◊ï“ì∞ß–†¶gVÊ7Fñˆ‚'Vñ∆DÜVDf˜&V67DFó2Çí∞¢6ˆÁ7Bf˜&V67B“7W'&VÁDÜˆ÷UvVFÜW$f˜&V67B«¬∑”∞¢6ˆÁ7BFñ÷W2“'&íÊó4'&íÜf˜&V67BÁFñ÷W2íÚf˜&V67BÁFñ÷W2¢µ”∞¢6ˆÁ7BFV◊2“'&íÊó4'&íÜf˜&V67BÁFV◊2íÚf˜&V67BÁFV◊2¢µ”∞¢6ˆÁ7B6ˆFW2“'&íÊó4'&íÜf˜&V67BÊ6ˆFW2íÚf˜&V67BÊ6ˆFW2¢µ”∞¢6ˆÁ7B'îFFR“ÊWr÷Çì∞¢Fñ÷W2Êf˜$V6ÇÇáFñ÷R¬ñGÇí”‚∞¢6ˆÁ7BFFR“ÊWrFFRáFñ÷Rì∞¢ñbÑÁV÷&W"Êó4Ê‚ÜFFRÊvWEFñ÷RÇííí&WGW&„∞¢6ˆÁ7B∂Wí“vWDÜVDf˜&V67DFî∂WíÜFFRì∞¢6ˆÁ7BóFV““'îFFRÊvWBÜ∂Wíí«¬≤FFR¬÷ÖFV◊¢‘ñÊfñÊóGí¬6ˆFS¢6ˆFW5∂ñGÖ“”∞¢6ˆÁ7BFV◊“ÁV÷&W"áFV◊5∂ñGÖ“ì∞¢ñbÑÁV÷&W"Êó4fñÊóFRáFV◊íbbFV◊‚óFV“Ê÷ÖFV◊í∞¢óFV“Ê÷ÖFV◊“FV◊∞¢óFV“Ê6ˆFR“6ˆFW5∂ñGÖ”∞¢–¢'îFFRÁ6WBÜ∂Wí¬óFV“ì∞¢“ì∞¢6ˆÁ7BFˆFí“ÊWrFFRÇì∞¢FˆFíÁ6WDÜ˜W'2É¬¬¬ì∞¢&WGW&‚'&íÊg&ˆ“á≤∆VÊwFÉ¢R“¬ÖÚ¬ˆfg6WBí”‚∞¢6ˆÁ7BFFR“ÊWrFFRáFˆFíì∞¢FFRÁ6WDFFRáFˆFíÊvWDFFRÇí≤ˆfg6WBì∞¢6ˆÁ7B∂Wí“vWDÜVDf˜&V67DFî∂WíÜFFRì∞¢6ˆÁ7BóFV““'îFFRÊvWBÜ∂Wíì∞¢ñbÇóFV“«¬ÁV÷&W"Êó4fñÊóFRÜóFV“Ê÷ÖFV◊íí&WGW&‚≤FFR¬fñ∆&∆S¢f«6R”∞¢6ˆÁ7B∆WfV¬“vWDÜVE&ó6¥g&ˆ’FV◊ÜóFV“Ê÷ÖFV◊ì∞¢&WGW&‚≤FFR¬fñ∆&∆S¢G'VR¬÷ÖFV◊¢÷FÇÁ&˜VÊBÜóFV“Ê÷ÖFV◊í¬6ˆFS¢óFV“Ê6ˆFR¬∆WfV¬¬7VvvW7Fñˆ„¢vWDÜVE&ó6µ7VvvW7Fñˆ‚Ü∆WfV¬í”∞¢“ì∞ß–†¶gVÊ7Fñˆ‚vWDÜVDf˜&V67DFî∂WíÜFFRí∞¢ñbÇÜFFRñÁ7FÊ6VˆbFFRí«¬ÁV÷&W"Êó4Ê‚ÜFFRÊvWEFñ÷RÇííí&WGW&‚"#∞¢6ˆÁ7BñV"“FFRÊvWDgV∆≈ñV"Çì∞¢6ˆÁ7B÷ˆÁFÇ“7G&ñÊrÜFFRÊvWD÷ˆÁFÇÇí≤íÁE7F'BÉ"¬#"ì∞¢6ˆÁ7BFí“7G&ñÊrÜFFRÊvWDFFRÇííÁE7F'BÉ"¬#"ì∞¢&WGW&‚G∑ñV'““G∂÷ˆÁFá““G∂Fó÷∞ß–†¶gVÊ7Fñˆ‚vWDÜVDf˜&V67D7W'&VÁDFï7F'BÜFFT∂Wíí∞¢6ˆÁ7BÊ˜r“ÊWrFFRÇì∞¢ñbÜvWDÜVDf˜&V67DFî∂WíÜÊ˜rí”“FFT∂Wíí&WGW&‚ÁV∆√∞¢Ê˜rÁ6WD÷ñÁWFW2É¬¬ì∞¢&WGW&‚Ê˜s∞ß–†¶gVÊ7Fñˆ‚vWDÜVDÜ˜W&«îf˜&V67Df˜$FíÜFFT∂Wíí∞¢6ˆÁ7Bf˜&V67B“7W'&VÁDÜˆ÷UvVFÜW$f˜&V67B«¬∑”∞¢6ˆÁ7BFñ÷W2“'&íÊó4'&íÜf˜&V67BÁFñ÷W2íÚf˜&V67BÁFñ÷W2¢µ”∞¢6ˆÁ7BFV◊2“'&íÊó4'&íÜf˜&V67BÁFV◊2íÚf˜&V67BÁFV◊2¢µ”∞¢6ˆÁ7B6ˆFW2“'&íÊó4'&íÜf˜&V67BÊ6ˆFW2íÚf˜&V67BÊ6ˆFW2¢µ”∞¢6ˆÁ7B&ñÁ2“'&íÊó4'&íÜf˜&V67BÁ&ñÁ2íÚf˜&V67BÁ&ñÁ2¢µ”∞¢6ˆÁ7BvñÊG2“'&íÊó4'&íÜf˜&V67BÁvñÊG2íÚf˜&V67BÁvñÊG2¢µ”∞¢6ˆÁ7B7W'&VÁDFï7F'B“vWDÜVDf˜&V67D7W'&VÁDFï7F'BÜFFT∂Wíì∞¢&WGW&‚Fñ÷W2Ê÷ÇáFñ÷R¬ñGÇí”‚∞¢6ˆÁ7BFFR“ÊWrFFRáFñ÷Rì∞¢ñbÑÁV÷&W"Êó4Ê‚ÜFFRÊvWEFñ÷RÇíí«¬vWDÜVDf˜&V67DFî∂WíÜFFRí”“FFT∂Wíí&WGW&‚ÁV∆√∞¢ñbÜ7W'&VÁDFï7F'BbbFFR¬7W'&VÁDFï7F'Bí&WGW&‚ÁV∆√∞¢6ˆÁ7BFV◊“ÁV÷&W"áFV◊5∂ñGÖ“ì∞¢6ˆÁ7B∆WfV¬“vWDÜVE&ó6¥g&ˆ’FV◊áFV◊ì∞¢&WGW&‚≤FFR¬FV◊¬6ˆFS¢6ˆFW5∂ñGÖ“¬&ñ„¢&ñÁ5∂ñGÖ“¬vñÊC¢vñÊG5∂ñGÖ“¬∆WfV¬”∞¢“íÊfñ«FW"Ñ&ˆˆ∆V‚ì∞ß–†¶gVÊ7Fñˆ‚&VÊFW$ÜVDf˜&V67D6&G2Çí∞¢6ˆÁ7BFó2“'Vñ∆DÜVDf˜&V67DFó2Çì∞¢ñbÇFó2Á6ˆ÷RÇÜFíí”‚FíÊfñ∆&∆Ríí&WGW&‚«6∆73“'v˜&∂∆ñ÷FR÷V◊Gí◊7FFR#‰FFíÊˆ‚Fó7ˆÊñ&ñ∆í¬&ó&˜fú;íF&Fí„¬˜Ê∞¢&WGW&‚∆Fób6∆73“&ÜVB÷f˜&V67B◊7G&ó"&ñ÷∆&V√“%&Wfó6ñˆÊí&ó66ÜñÚ6∆˜&R&˜76ñ÷íRvñ˜&Êí#‚G∂Fó2Ê÷ÇÜFíí”‚∞¢ñbÇFíÊfñ∆&∆Rí&WGW&‚∆'Fñ6∆R6∆73“&ÜVB÷f˜&V67B÷6&BÜVB÷f˜&V67B◊VÊfñ∆&∆R#„«7G&ˆÊs‚G∂W66TÖD‘¬Üf˜&÷DÜVDF6Ü&ˆ&DFFRÜFíÊFFRíó”¬˜7G&ˆÊs„«‰FFíÊˆ‚Fó7ˆÊñ&ñ∆í¬&ó&˜fú;íF&Fí„¬˜„¬ˆ'Fñ6∆SÊ∞¢6ˆÁ7Bñ6ˆ‚“tTDÑU%ÙƒU%EÙî4ÙÂ∂FíÊ∆WfV≈“«¬/	˘˙"#∞¢6ˆÁ7BvVFÜW$ñ6ˆ‚“vVFÜW$6ˆFT∆&V¬ÜFíÊ6ˆFRíÁ7∆óBÇ""ï≥“«¬.)à˚àÚ#∞¢6ˆÁ7B∆&V¬“tı$¥ƒî‘DUÙ4Ùƒı%Ùƒ$T≈∂FíÊ∆WfV≈“«¬FíÊ∆WfV√∞¢6ˆÁ7BFî∂Wí“vWDÜVDf˜&V67DFî∂WíÜFíÊFFRì∞¢&WGW&‚∆'WGFˆ‚GóS“&'WGFˆ‚"6∆73“&ÜVB÷f˜&V67B÷6&B&ó6≤“G∂W66TÖD‘¬ÜFíÊ∆WfV¬ó“"FF÷ÜVB÷Fì“"G∂W66TÖD‘¬ÜFî∂Wíó“"&ñ÷∆&V√“$&í&Wfó6ñˆÊí˜&W"˜&W"G∂W66TÖD‘¬Üf˜&÷DÜVDF6Ü&ˆ&DFFRÜFíÊFFRíó“#„∆Fóc„«7G&ˆÊs‚G∂W66TÖD‘¬Üf˜&÷DÜVDF6Ü&ˆ&DFFRÜFíÊFFRíó”¬˜7G&ˆÊs„«7„‚G∂W66TÖD‘¬ávVFÜW$ñ6ˆ‚ó”¬˜7„„¬ˆFóc„«6∆73“&ÜVB◊&ó6≤◊ñ∆¬#‚G∂W66TÖD‘¬Üñ6ˆ‚ó“G∂W66TÖD‘¬Ü∆&V¬ó”¬˜„∆#‚G∂W66TÖD‘¬Ö7G&ñÊrÜFíÊ÷ÖFV◊íó‹+3¬ˆ#„«6÷∆√‚G∂W66TÖD‘¬ÜFíÁ7VvvW7Fñˆ‚ó”¬˜6÷∆√„∆V”‰˜&W"˜&¬ˆV”„¬ˆ'WGFˆ„Ê∞¢“íÊ¶ˆñ‚Ç""ó”¬ˆFócÊ∞ß–†¶gVÊ7Fñˆ‚˜V‰ÜVDÜ˜W&«îf˜&V67BÜFî∂Wíí∞¢6ˆÁ7B6∆˜G2“vWDÜVDÜ˜W&«îf˜&V67Df˜$FíÜFî∂Wíì∞¢6ˆÁ7BFîFFR“6∆˜G5≥”ÚÊFFR«¬ÊWrFFRÜG∂Fî∂Wó’C££ì∞¢6ˆÁ7BFóF∆R“f˜&÷DÜVDF6Ü&ˆ&DFFRÜFîFFRì∞¢6ˆÁ7B&˜w2“6∆˜G2Ê÷Çá6∆˜Bí”‚∞¢6ˆÁ7BÜ˜W"“6∆˜BÊFFRÁFÙ∆ˆ6∆UFñ÷U7G&ñÊrÇ&óB‘ïB"¬≤Ü˜W#¢#"÷FñvóB"¬÷ñÁWFS¢#"÷FñvóB"¬Ü˜W##¢f«6R“ì∞¢6ˆÁ7B∆&V¬“vVFÜW$6ˆFT∆&V¬á6∆˜BÊ6ˆFRì∞¢6ˆÁ7Bñ6ˆ‚“∆&V¬Á7∆óBÇ""ï≥“«¬.)à˚àÚ#∞¢6ˆÁ7B&ó6¥ñ6ˆ‚“tTDÑU%ÙƒU%EÙî4ÙÂ∑6∆˜BÊ∆WfV≈“«¬/	˘˙"#∞¢6ˆÁ7B&ó6¥∆&V¬“tı$¥ƒî‘DUÙ4Ùƒı%Ùƒ$T≈∑6∆˜BÊ∆WfV≈“«¬6∆˜BÊ∆WfV√∞¢6ˆÁ7BFV◊∆&V¬“ÁV÷&W"Êó4fñÊóFRá6∆˜BÁFV◊íÚG¥÷FÇÁ&˜VÊBá6∆˜BÁFV◊ó‹+6¢&‚ˆB#∞¢6ˆÁ7B&ñ‰∆&V¬“ÁV÷&W"Êó4fñÊóFRÑÁV÷&W"á6∆˜BÁ&ñ‚ííÚG¥÷FÇÁ&˜VÊBÑÁV÷&W"á6∆˜BÁ&ñ‚íó“V¢&‚ˆB#∞¢6ˆÁ7BvñÊD∆&V¬“ÁV÷&W"Êó4fñÊóFRÑÁV÷&W"á6∆˜BÁvñÊBííÚG¥÷FÇÁ&˜VÊBÑÁV÷&W"á6∆˜BÁvñÊBíó“∂“ˆÜ¢&‚ˆB#∞¢&WGW&‚∆'Fñ6∆R6∆73“&ÜVB÷Ü˜W&«í◊&˜r&ó6≤“G∂W66TÖD‘¬á6∆˜BÊ∆WfV¬ó“#„∆Fóc„«7G&ˆÊs‚G∂W66TÖD‘¬ÜÜ˜W"ó”¬˜7G&ˆÊs„«7‚&ñ÷ÜñFFV„“'G'VR#‚G∂W66TÖD‘¬Üñ6ˆ‚ó”¬˜7„„¬ˆFóc„«‚G∂W66TÖD‘¬Ü∆&V¬Á&W∆6RÇıÂµ‰’¶◊¨8‹;ı“µ«2¢Ú¬""íó”¬˜„∆#‚G∂W66TÖD‘¬áFV◊∆&V¬ó”¬ˆ#„«6÷∆√Ô	¯ ~˚àÚG∂W66TÖD‘¬á&ñ‰∆&V¬ó“+r	¯ Œ˚àÚG∂W66TÖD‘¬ávñÊD∆&V¬ó”¬˜6÷∆√„∆V”‚G∂W66TÖD‘¬á&ó6¥ñ6ˆ‚ó“G∂W66TÖD‘¬á&ó6¥∆&V¬ó”¬ˆV”„¬ˆ'Fñ6∆SÊ∞¢“íÊ¶ˆñ‚Ç""ì∞¢6ˆÁ7B˜fW&∆í“Fˆ7V÷VÁBÊ7&VFTV∆V÷VÁBÇ&Fób"ì∞¢˜fW&∆íÊ6∆74Ê÷R“'v˜&∂∆ñ÷FR÷÷ˆF¬÷˜fW&∆íÜVB÷Ü˜W&«í÷˜fW&∆í#∞¢˜fW&∆íÊñÊÊW$ÖD‘¬“∆Fób6∆73“'v˜&∂∆ñ÷FR÷÷ˆF¬ÜVB÷Ü˜W&«í÷÷ˆF¬"&ˆ∆S“&Fñ∆ˆr"&ñ÷÷ˆF√“'G'VR"&ñ÷∆&V√“%&Wfó6ñˆÊí˜&W"˜&&ó66ÜñÚ6∆˜&R#„∆ÜVFW"6∆73“'v˜&∂∆ñ÷FR◊vR÷ÜVFW"#„∆'WGFˆ‚GóS“&'WGFˆ‚"6∆73“'v˜&∂∆ñ÷FR◊vR÷&6≤"&ñ÷∆&V√“%F˜&Ê¬&ó66ÜñÚ6∆˜&R#Ó(i¬ˆ'WGFˆ„„∆Fóc„«Â&Wfó6ñˆÊí˜&W"˜&¬˜„«7G&ˆÊs‚G∂W66TÖD‘¬áFóF∆Ró”¬˜7G&ˆÊs„¬ˆFóc„∆'WGFˆ‚GóS“&'WGFˆ‚"6∆73“'v˜&∂∆ñ÷FR÷÷ˆF¬÷6∆˜6R"&ñ÷∆&V√“$6ÜóVFí#Ï9s¬ˆ'WGFˆ„„¬ˆÜVFW#„∆÷ñ‚6∆73“&ÜVB÷Ü˜W&«í÷∆ó7B#‚G∑&˜w2«¬#«6∆73“wv˜&∂∆ñ÷FR÷V◊Gí◊7FFRsÂ&Wfó6ñˆÊí˜&&ñRÊˆ‚Fó7ˆÊñ&ñ∆íW"VW7FÚvñ˜&ÊÚ„¬˜‚'”¬ˆ÷ñ„„¬ˆFócÊ∞¢6ˆÁ7B6∆˜6R“Çí”‚˜fW&∆íÁ&V÷˜fRÇì∞¢˜fW&∆íÊFDWfVÁD∆ó7FVÊW"Ç&6∆ñ6≤"¬ÜWfVÁBí”‚≤ñbÜWfVÁBÁF&vWB””“˜fW&∆íí6∆˜6RÇì≤“ì∞¢˜fW&∆íÁVW'ï6V∆V7F˜"Ç"Áv˜&∂∆ñ÷FR÷÷ˆF¬÷6∆˜6R"ìÚÊFDWfVÁD∆ó7FVÊW"Ç&6∆ñ6≤"¬6∆˜6Rì∞¢˜fW&∆íÁVW'ï6V∆V7F˜"Ç"Áv˜&∂∆ñ÷FR◊vR÷&6≤"ìÚÊFDWfVÁD∆ó7FVÊW"Ç&6∆ñ6≤"¬6∆˜6Rì∞¢Fˆ7V÷VÁBÊ&ˆGíÊVÊD6Üñ∆BÜ˜fW&∆íì∞¢˜fW&∆íÁVW'ï6V∆V7F˜"Ç"Áv˜&∂∆ñ÷FR◊vR÷&6≤"ìÚÊfˆ7W2Çì∞ß–†¶gVÊ7Fñˆ‚˜V‰Üˆ÷Uv˜&∂∆ñ÷FT&ˆ&Bá≤&ó6¥∆WfV¬“'fW&FR"¬W&¬“tı$¥ƒî‘DUÙdı$T45EıU$¬¬6ˆÁFWáDFF“ÁV∆¬““∑“í∞¢6ˆÁ7B∆WfV¬“Ê˜&÷∆ó¶Uv˜&∂∆ñ÷FT∆WfV¬á&ó6¥∆WfV¬ì∞¢6ˆÁ7B∆WfVƒ∆&V¬“tı$¥ƒî‘DUÙ4Ùƒı%Ùƒ$T≈∂∆WfV≈“«¬∆WfV√∞¢6ˆÁ7Bñ6ˆ‚“tTDÑU%ÙƒU%EÙî4ÙÂ∂∆WfV≈“«¬/	˘˙"#∞¢6ˆÁ7BF&vWB“7W'&VÁEvVFÜW%F&vWB«¬vWEvVFÜW%F&vWD6ˆ˜&FñÊFW2Çì∞¢6ˆÁ7B6ˆ˜&FñÊFW4∆&V¬“ÁV÷&W"Êó4fñÊóFRÑÁV÷&W"áF&vWCÚÊ∆BííbbÁV÷&W"Êó4fñÊóFRÑÁV÷&W"áF&vWCÚÊ∆ˆ‚ííÚG¥ÁV÷&W"áF&vWBÊ∆BíÁFÙfóÜVBÉBó“¬G¥ÁV÷&W"áF&vWBÊ∆ˆ‚íÁFÙfóÜVBÉBó÷¢"#∞¢6ˆÁ7B˜6óFñˆ‰∆&V¬“F&vWCÚÁ6˜W&6R””“&w2"Ú˜6ó¶ñˆÊRu2GGV∆RG∂6ˆ˜&FñÊFW4∆&V¬Ú(
"G∂6ˆ˜&FñÊFW4∆&V«÷¢"'÷¢F&vWCÚÁ6˜W&6R””“&÷ÁV¬"ÚG∑F&vWBÊÊ÷R«¬$∆ˆ6∆óL:66V«F'“G∂6ˆ˜&FñÊFW4∆&V¬Ú(
"G∂6ˆ˜&FñÊFW4∆&V«÷¢"'÷¢F&vWCÚÁ6˜W&6R””“&6ˆ÷÷W76"Ú6ˆ◊VÊR˜¶ˆÊFV∆∆6ˆ÷÷W76G∂6ˆ˜&FñÊFW4∆&V¬Ú(
"G∂6ˆ˜&FñÊFW4∆&V«÷¢"'÷¢%˜7F¶ñˆÊR&VFVfñÊóF#∞¢6ˆÁ7BWFFVD∆&V¬“7W'&VÁDÜˆ÷UvVFÜW$f˜&V67CÚÁWFFVDBÚÊWrFFRÜ7W'&VÁDÜˆ÷UvVFÜW$f˜&V67BÁWFFVDBíÁFÙ∆ˆ6∆U7G&ñÊrÇ&óB‘ïB"¬≤FFU7Gñ∆S¢'6Ü˜'B"¬Fñ÷U7Gñ∆S¢'6Ü˜'B"“í¢$FFíÊˆ‚Fó7ˆÊñ&ñ∆í#∞¢6ˆÁ7B˜W&FñˆÊƒ6&G2“∞¢≤ñ6ˆ„¢/	˘8≤"¬FóF∆S¢%&ñ÷FV¬GW&ÊÚ"¬óFV◊3¢≤$6ˆÁG&ˆ∆∆&Rv˜&∂∆ñ÷FR"¬%fW&ñfñ6&RFó7ˆÊñ&ñ∆óL:7V"¬$ñÊf˜&÷&R∆7VG&%““¿¢≤ñ6ˆ„¢.)à˚àÚ"¬FóF∆S¢$GW&ÁFRñ¬GW&ÊÚ"¬óFV◊3¢≤$f&RW6R&Vvˆ∆&í"¬$∆f˜&&R&VfW&ñ&ñ∆÷VÁFR∆Œ(	ñˆ÷'&%““¿¢≤ñ6ˆ„¢/	˘JR"¬FóF∆S¢$˜&R6VÁG&∆í#£3”c£"¬óFV◊3¢≤%&ñGW'&R∆RGFófóL:W6ÁFí"¬$V÷VÁF&R∆g&WVVÁ¶FV∆∆RW6R%““¿¢≤ñ6ˆ„¢/	˘™Ç"¬FóF∆S¢$ñ‚66ÚFí∆ófV∆∆Ú&˜76Ú"¬óFV◊3¢≤%&ñ÷ˆGV∆&R∆RGFófóL:"¬%7˜7F&Rí∆f˜&íw&f˜6í¬÷GFñÊÚ"¬%f«WF&R∆6˜7VÁ6ñˆÊRFVí∆f˜&íú;íW6ÁFí%“–¢”∞¢6ˆÁ7B˜W&FñˆÊƒ÷&∑W“˜W&FñˆÊƒ6&G2Ê÷ÇÜ6&Bí”‚∆'Fñ6∆R6∆73“&ÜVB÷7Fñˆ‚÷6&B#„«7‚&ñ÷ÜñFFV„“'G'VR#‚G∂W66TÖD‘¬Ü6&BÊñ6ˆ‚ó”¬˜7„„∆É3‚G∂W66TÖD‘¬Ü6&BÁFóF∆Ró”¬ˆÉ3„«V√‚G∂6&BÊóFV◊2Ê÷ÇÜóFV“í”‚∆∆ì‚G∂W66TÖD‘¬ÜóFV“ó”¬ˆ∆ìÊíÊ¶ˆñ‚Ç""ó”¬˜V√„¬ˆ'Fñ6∆SÊíÊ¶ˆñ‚Ç""ì∞¢6ˆÁ7B6ˆÁFWáD÷&∑W“6ˆÁFWáDFFÚ«6V7Fñˆ‚6∆73“&ÜVB÷6ˆÁFWáB÷6&B#„∆É#‰6ˆ÷÷W766V∆W¶ñˆÊF¬ˆÉ#„«„∆#‚G∂W66TÖD‘¬Ü6ˆÁFWáDFFÊ6ˆ÷÷W76«¬$6ˆ÷÷W76"ó”¬ˆ#‚G∂6ˆÁFWáDFFÊ6ˆFñ6T6ˆ÷÷W76Ú+rG∂W66TÖD‘¬Ü6ˆÁFWáDFFÊ6ˆFñ6T6ˆ÷÷W76ó÷¢"'”¬˜„«‚G∂W66TÖD‘¬Ü6ˆÁFWáDFFÊ6ˆ◊VÊR«¬$6ˆ◊VÊRÊˆ‚Fó7ˆÊñ&ñ∆R"ó“+rG∂W66TÖD‘¬Ü6ˆÁFWáDFFÁ6V∆V7FVDFFR«¬$FFÊˆ‚Fó7ˆÊñ&ñ∆R"ó“+r	¯ ˚àÚG∂W66TÖD‘¬Üf˜&÷Ev˜&∂∆ñ÷FUFV◊W&GW&RÜ6ˆÁFWáDFFÊfW&vUFV◊W&GW&Ríó”¬˜„«‰∆ófV∆∆Ú∆∆W'F¢∆#‰6ˆFñ6RG∂W66TÖD‘¬Ü6ˆÁFWáDFFÊ∆W'D∆WfV¬«¬∆WfV¬ó”¬ˆ#‚+rfˆÁFS¢G∂W66TÖD‘¬Ü6ˆÁFWáDFFÁ6˜W&6R«¬%v˜&∂∆ñ÷FRˆ÷WFVÚ"ó”¬˜„¬˜6V7Fñˆ„Ê¢"#∞¢6ˆÁ7B˜fW&∆í“Fˆ7V÷VÁBÊ7&VFTV∆V÷VÁBÇ&Fób"ì∞¢˜fW&∆íÊ6∆74Ê÷R“'v˜&∂∆ñ÷FR÷÷ˆF¬÷˜fW&∆ív˜&∂∆ñ÷FR◊vR÷˜fW&∆í#∞¢˜fW&∆íÊñÊÊW$ÖD‘¬“∆Fób6∆73“'v˜&∂∆ñ÷FR÷÷ˆF¬v˜&∂∆ñ÷FR÷&ˆ&B÷÷ˆF¬v˜&∂∆ñ÷FR÷&ˆ&B◊vRÜVB÷F6Ü&ˆ&B"&ˆ∆S“&Fñ∆ˆr"&ñ÷÷ˆF√“'G'VR"&ñ÷∆&V√“%&ˆ6VGW&6ñ7W&Wß¶&ó66ÜñÚ6∆˜&R#‡¢∆ÜVFW"6∆73“'v˜&∂∆ñ÷FR◊vR÷ÜVFW"#‡¢∆'WGFˆ‚GóS“&'WGFˆ‚"6∆73“'v˜&∂∆ñ÷FR◊vR÷&6≤"&ñ÷∆&V√“$6ÜóVFívñÊ&ó66ÜñÚ6∆˜&R#Ó(i¬ˆ'WGFˆ„‡¢∆Fóc„«Â&ˆ6VGW&6ñ7W&Wß¶¬˜„«7G&ˆÊsÂ&ó66ÜñÚ6∆˜&S¬˜7G&ˆÊs„¬ˆFóc‡¢∆'WGFˆ‚GóS“&'WGFˆ‚"6∆73“'v˜&∂∆ñ÷FR÷÷ˆF¬÷6∆˜6R"&ñ÷∆&V√“$6ÜóVFí#Ï9s¬ˆ'WGFˆ„‡¢¬ˆÜVFW#‡¢∆÷ñ‚FF◊v˜&∂∆ñ÷FR÷÷ñ‚6∆73“&ÜVB÷F6Ü&ˆ&B÷÷ñ‚#‡¢«6V7Fñˆ‚6∆73“&ÜVB◊7FGW2÷6&B&ó6≤“G∂W66TÖD‘¬Ü∆WfV¬ó“#‡¢∆Fób6∆73“&ÜVB◊7FGW2÷ñ6ˆ‚"&ñ÷ÜñFFV„“'G'VR#‚G∂W66TÖD‘¬Üñ6ˆ‚ó”¬ˆFóc‡¢∆Fóc„∆ÉÂ&ó66ÜñÚ6∆˜&S¬ˆÉ„«‚G∂W66TÖD‘¬á˜6óFñˆ‰∆&V¬ó”¬˜„«6÷∆√ÂV«Fñ÷Úvvñ˜&Ê÷VÁFÛ¢G∂W66TÖD‘¬áWFFVD∆&V¬ó”¬˜6÷∆√„¬ˆFóc‡¢∆'WGFˆ‚GóS“&'WGFˆ‚"6∆73“&ÜVB÷∆ˆ6Fñˆ‚÷'WGFˆ‚"FF÷ÜVB÷∆ˆ6Fñˆ‚◊ñ6∂W#Ô	˘8“6÷&ñ∆ˆ6∆óL:¬ˆ'WGFˆ„‡¢«7G&ˆÊs‚G∂W66TÖD‘¬Ü∆WfVƒ∆&V¬ó”¬˜7G&ˆÊs‡¢¬˜6V7Fñˆ„‡¢G∂6ˆÁFWáD÷&∑W–¢«6V7Fñˆ‚6∆73“&ÜVB◊6V7Fñˆ‚#„∆Fób6∆73“&ÜVB◊6V7Fñˆ‚◊FóF∆R#„«7‚&ñ÷ÜñFFV„“'G'VR#Ó)à˚àÛ¬˜7„„∆É#Â&Wfó6ñˆÊí&˜76ñ÷íRvñ˜&Êì¬ˆÉ#„¬ˆFóc‚G∑&VÊFW$ÜVDf˜&V67D6&G2Çó”¬˜6V7Fñˆ„‡¢«6V7Fñˆ‚6∆73“&ÜVB◊6V7Fñˆ‚#„∆Fób6∆73“&ÜVB◊6V7Fñˆ‚◊FóF∆R#„«7‚&ñ÷ÜñFFV„“'G'VR#Ó)™˚àÛ¬˜7„„∆É#‰ñÊFñ6¶ñˆÊí˜W&FófS¬ˆÉ#„¬ˆFóc„∆Fób6∆73“&ÜVB÷7FñˆÁ2÷w&ñB#‚G∂˜W&FñˆÊƒ÷&∑W”¬ˆFóc„¬˜6V7Fñˆ„‡¢«6V7Fñˆ‚6∆73“&ÜVB÷ñÊfÚ÷6&BáñG&Fñˆ‚#„∆É#Ô	˘*rñG&F¶ñˆÊS¬ˆÉ#„«‰&W&R∆÷VÊÚ#S÷¬Fí7VˆvÊí#÷ñÁWFí„¬˜„¬˜6V7Fñˆ„‡¢«6V7Fñˆ‚6∆73“&ÜVB÷∆r÷6ˆ◊7B#„∆Fóc„∆É#Ô	˘8¬˜&FñÊÁ¶&VvñˆÊ∆R‚„s"FV¬2ÛbÛ##c¬ˆÉ#„«‰÷ó7W&RFí&WfVÁ¶ñˆÊRW"GFófóL:∆f˜&FófRñ‚6ˆÊFó¶ñˆÊíFíW7˜6ó¶ñˆÊR¬6∆˜&R„¬˜„¬ˆFóc„∆'WGFˆ‚GóS“&'WGFˆ‚"6∆73“&'F‚'F‚◊&ñ÷'ív˜&∂∆ñ÷FR◊fó6óB÷'F‚"FF◊v˜&∂∆ñ÷FR◊fó6óC“"G∂W66TÖD‘¬áW&¬ó“#‰&íFˆ7V÷VÁFÛ¬ˆ'WGFˆ„„¬˜6V7Fñˆ„‡¢¬ˆ÷ñ„‡¢¬ˆFócÊ∞¢6ˆÁ7B6∆˜6R“Çí”‚∞¢Fˆ7V÷VÁBÊ&ˆGíÊ6∆74∆ó7BÁ&V÷˜fRÇ'v˜&∂∆ñ÷FR◊vR÷˜V‚"ì∞¢˜fW&∆íÁ&V÷˜fRÇì∞¢”∞¢Fˆ7V÷VÁBÊ&ˆGíÊ6∆74∆ó7BÊFBÇ'v˜&∂∆ñ÷FR◊vR÷˜V‚"ì∞¢˜fW&∆íÊFDWfVÁD∆ó7FVÊW"Ç&6∆ñ6≤"¬ÜWfVÁBí”‚≤ñbÜWfVÁBÁF&vWB””“˜fW&∆íí6∆˜6RÇì≤“ì∞¢˜fW&∆íÁVW'ï6V∆V7F˜"Ç"Áv˜&∂∆ñ÷FR÷÷ˆF¬÷6∆˜6R"ìÚÊFDWfVÁD∆ó7FVÊW"Ç&6∆ñ6≤"¬6∆˜6Rì∞¢˜fW&∆íÁVW'ï6V∆V7F˜"Ç"Áv˜&∂∆ñ÷FR◊vR÷&6≤"ìÚÊFDWfVÁD∆ó7FVÊW"Ç&6∆ñ6≤"¬6∆˜6Rì∞¢˜fW&∆íÁVW'ï6V∆V7F˜"Ç%∂FF÷ÜVB÷∆ˆ6Fñˆ‚◊ñ6∂W%“"ìÚÊFDWfVÁD∆ó7FVÊW"Ç&6∆ñ6≤"¬Çí”‚˜V‰ÜVD∆ˆ6FñˆÂñ6∂W"Ü˜fW&∆ííì∞¢˜fW&∆íÁVW'ï6V∆V7F˜$∆¬Ç%∂FF÷ÜVB÷Fï“"íÊf˜$V6ÇÇÜ6&Bí”‚∞¢6&BÊFDWfVÁD∆ó7FVÊW"Ç&6∆ñ6≤"¬Çí”‚˜V‰ÜVDÜ˜W&«îf˜&V67BÜ6&BÊvWDGG&ñ'WFRÇ&FF÷ÜVB÷Fí"í«¬""íì∞¢“ì∞¢˜fW&∆íÁVW'ï6V∆V7F˜"Ç%∂FF◊v˜&∂∆ñ÷FR◊fó6óE“"ìÚÊFDWfVÁD∆ó7FVÊW"Ç&6∆ñ6≤"¬ÜWfVÁBí”‚∞¢6ˆÁ7Bfó6óEW&¬“WfVÁBÊ7W'&VÁEF&vWBÊvWDGG&ñ'WFRÇ&FF◊v˜&∂∆ñ÷FR◊fó6óB"í«¬tı$¥ƒî‘DUÙdı$T45EıU$√∞¢vñÊF˜rÊ˜V‚áfó6óEW&¬¬%ˆ&∆Ê≤"¬&Êˆ˜VÊW"∆Ê˜&VfW'&W""ì∞¢“ì∞¢Fˆ7V÷VÁBÊ&ˆGíÊVÊD6Üñ∆BÜ˜fW&∆íì∞¢˜fW&∆íÁVW'ï6V∆V7F˜"Ç"Áv˜&∂∆ñ÷FR◊vR÷&6≤"ìÚÊfˆ7W2Çì∞ß–††¶gVÊ7Fñˆ‚'Vñ∆DÜVD∆ˆ6Fñˆ‰˜Fñˆ‰÷&∑WÇí∞¢6ˆÁ7B∆6W2“∞¢≤Ê÷S¢$&ˆ∆ˆvÊ"¬∆C¢CB„CìCí¬∆ˆ„¢„3C#b“¿¢≤Ê÷S¢$÷ˆFVÊ"¬∆C¢CB„cCs¬∆ˆ„¢„ì#S"“¿¢≤Ê÷S¢$fW'&&"¬∆C¢CB„É3É¬∆ˆ„¢„cìÇ“¿¢≤Ê÷S¢%&fVÊÊ"¬∆C¢CB„CÉB¬∆ˆ„¢"„#3R–¢”∞¢&WGW&‚∆6W2Ê÷Çá∆6Rí”‚∆'WGFˆ‚GóS“&'WGFˆ‚"6∆73“&ÜVB◊∆6R÷˜Fñˆ‚"FF◊∆6R÷Ê÷S“"G∂W66TÖD‘¬á∆6RÊÊ÷Ró“"FF◊∆6R÷∆C“"G∂W66TÖD‘¬á∆6RÊ∆Bó“"FF◊∆6R÷∆ˆ„“"G∂W66TÖD‘¬á∆6RÊ∆ˆ‚ó“#„«7G&ˆÊs‚G∂W66TÖD‘¬á∆6RÊÊ÷Ró”¬˜7G&ˆÊs„«7„‚G¥ÁV÷&W"á∆6RÊ∆BíÁFÙfóÜVBÉBó“¬G¥ÁV÷&W"á∆6RÊ∆ˆ‚íÁFÙfóÜVBÉBó”¬˜7„„¬ˆ'WGFˆ„ÊíÊ¶ˆñ‚Ç""ì∞ß–†¶gVÊ7Fñˆ‚˜V‰ÜVD∆ˆ6FñˆÂñ6∂W"Ü&ˆ&D˜fW&∆íí∞¢6ˆÁ7Bñ6∂W"“Fˆ7V÷VÁBÊ7&VFTV∆V÷VÁBÇ&Fób"ì∞¢ñ6∂W"Ê6∆74Ê÷R“'v˜&∂∆ñ÷FR÷÷ˆF¬÷˜fW&∆íÜVB÷∆ˆ6Fñˆ‚÷˜fW&∆í#∞¢ñ6∂W"ÊñÊÊW$ÖD‘¬“∆Fób6∆73“'v˜&∂∆ñ÷FR÷÷ˆF¬v˜&∂∆ñ÷FR÷&ˆ&B÷÷ˆF¬v˜&∂∆ñ÷FR÷&ˆ&B◊vRÜVB÷∆ˆ6Fñˆ‚◊fñWr"&ˆ∆S“&Fñ∆ˆr"&ñ÷÷ˆF√“'G'VR"&ñ÷∆&V√“%66Vv∆í∆ˆ6∆óL:&ó66ÜñÚ6∆˜&R#‡¢∆ÜVFW"6∆73“'v˜&∂∆ñ÷FR◊vR÷ÜVFW"#‡¢∆'WGFˆ‚GóS“&'WGFˆ‚"6∆73“'v˜&∂∆ñ÷FR◊vR÷&6≤"FF÷∆ˆ6Fñˆ‚÷6∆˜6R&ñ÷∆&V√“%F˜&Ê¬&ó66ÜñÚ6∆˜&R#Ó(i¬ˆ'WGFˆ„‡¢∆Fóc„«Â&ó66ÜñÚ6∆˜&S¬˜„«7G&ˆÊs‰6÷&ñ∆ˆ6∆óL:¬˜7G&ˆÊs„¬ˆFóc‡¢∆'WGFˆ‚GóS“&'WGFˆ‚"6∆73“'v˜&∂∆ñ÷FR÷÷ˆF¬÷6∆˜6R"FF÷∆ˆ6Fñˆ‚÷6∆˜6R&ñ÷∆&V√“$6ÜóVFí#Ï9s¬ˆ'WGFˆ„‡¢¬ˆÜVFW#‡¢∆÷ñ‚6∆73“&ÜVB÷∆ˆ6Fñˆ‚÷÷ñ‚#‡¢«6V7Fñˆ‚6∆73“&ÜVB÷∆ˆ6Fñˆ‚◊ÊV¬#‡¢∆É#„‚66Vv∆íVÊ∆ˆ6∆óL:¬ˆÉ#‡¢«Â6V∆W¶ñˆÊVÊ6óGL:&ñF˜W&RñÁ6W&ó66í6ˆ˜&FñÊFR&V6ó6R„¬˜‡¢∆Fób6∆73“&ÜVB◊∆6R÷w&ñB#‚G∂'Vñ∆DÜVD∆ˆ6Fñˆ‰˜Fñˆ‰÷&∑WÇó”¬ˆFóc‡¢∆f˜&“6∆73“&ÜVB÷6ˆ˜&FñÊFR÷f˜&“"FF÷6ˆ˜&FñÊFR÷f˜&”‡¢∆∆&V√‰Êˆ÷R∆ˆ6∆óL:∆ñÁWBÊ÷S“&Ê÷R"GóS“'FWáB"∆6VÜˆ∆FW#“$W2‚6ÁFñW&RÊ˜&B#„¬ˆ∆&V√‡¢∆∆&V√‰∆FóGVFñÊS∆ñÁWBÊ÷S“&∆B"GóS“&ÁV÷&W""7FW“#„"÷ñ„“"”ì"÷É“#ì"&WVó&VC„¬ˆ∆&V√‡¢∆∆&V√‰∆ˆÊvóGVFñÊS∆ñÁWBÊ÷S“&∆ˆ‚"GóS“&ÁV÷&W""7FW“#„"÷ñ„“"”É"÷É“#É"&WVó&VC„¬ˆ∆&V√‡¢∆'WGFˆ‚GóS“'7V&÷óB"6∆73“&'F‚'F‚◊&ñ÷'í#ÂW6VW7F∆ˆ6∆óL:¬ˆ'WGFˆ„‡¢¬ˆf˜&”‡¢¬˜6V7Fñˆ„‡¢«6V7Fñˆ‚6∆73“&ÜVB÷∆ˆ6Fñˆ‚◊ÊV¬ÜVB÷÷◊ÊV¬#‡¢∆É#„"‚66Vv∆í7V∆∆÷¬ˆÉ#‡¢«ÂFˆ66∆÷66ÜW&÷ÚñÁFW&ÚW"˜6ó¶ñˆÊ&Rñ¬VÁFÚ„¬˜‡¢∆'WGFˆ‚GóS“&'WGFˆ‚"6∆73“&ÜVB÷÷÷˜V‚"FF÷˜V‚÷gV∆¬÷÷„«7„Ô	˘{Æ˚àÛ¬˜7„‰&í÷66ÜW&÷ÚñÁFW&Û¬ˆ'WGFˆ„‡¢¬˜6V7Fñˆ„‡¢¬ˆ÷ñ„‡¢¬ˆFócÊ∞¢6ˆÁ7B6∆˜6Uñ6∂W"“Çí”‚ñ6∂W"Á&V÷˜fRÇì∞¢6ˆÁ7B«î∆ˆ6Fñˆ‚“7ñÊ2Ü∆ˆ6Fñˆ‚í”‚∞¢6V∆V7FVEvVFÜW$∆ˆ6Fñˆ‚“∆ˆ6Fñˆ„∞¢7W'&VÁEvVFÜW%F&vWB“≤‚‚Ê∆ˆ6Fñˆ‚¬6˜W&6S¢&÷ÁV¬"”∞¢6∆˜6Uñ6∂W"Çì∞¢&ˆ&D˜fW&∆ìÚÁ&V÷˜fRÇì∞¢Fˆ7V÷VÁBÊ&ˆGíÊ6∆74∆ó7BÁ&V÷˜fRÇ'v˜&∂∆ñ÷FR◊vR÷˜V‚"ì∞¢vóBfWF6ÖvVFÜW"Çì∞¢˜V‰Üˆ÷Uv˜&∂∆ñ÷FT&ˆ&Bá≤&ó6¥∆WfV√¢vWDÜˆ÷Uv˜&∂∆ñ÷FU&ó6¥∆WfV¬Ü7W'&VÁDÜˆ÷UvVFÜW$f˜&V67CÚÁFV◊2«¬µ“í¬W&√¢tı$¥ƒî‘DUÙdı$T45EıU$¬“ì∞¢”∞¢ñ6∂W"ÁVW'ï6V∆V7F˜$∆¬Ç%∂FF÷∆ˆ6Fñˆ‚÷6∆˜6U“"íÊf˜$V6ÇÇÜ'F‚í”‚'F‚ÊFDWfVÁD∆ó7FVÊW"Ç&6∆ñ6≤"¬6∆˜6Uñ6∂W"íì∞¢ñ6∂W"ÁVW'ï6V∆V7F˜$∆¬Ç%∂FF◊∆6R÷Ê÷U“"íÊf˜$V6ÇÇÜ'F‚í”‚'F‚ÊFDWfVÁD∆ó7FVÊW"Ç&6∆ñ6≤"¬Çí”‚«î∆ˆ6Fñˆ‚á≤Ê÷S¢'F‚ÊFF6WBÁ∆6TÊ÷R¬∆C¢ÁV÷&W"Ü'F‚ÊFF6WBÁ∆6T∆Bí¬∆ˆ„¢ÁV÷&W"Ü'F‚ÊFF6WBÁ∆6T∆ˆ‚í“ííì∞¢ñ6∂W"ÁVW'ï6V∆V7F˜"Ç%∂FF÷6ˆ˜&FñÊFR÷f˜&’“"ìÚÊFDWfVÁD∆ó7FVÊW"Ç'7V&÷óB"¬ÜWfVÁBí”‚∞¢WfVÁBÁ&WfVÁDFVfV«BÇì∞¢6ˆÁ7Bf˜&““ÊWrf˜&‘FFÜWfVÁBÊ7W'&VÁEF&vWBì∞¢«î∆ˆ6Fñˆ‚á≤Ê÷S¢7G&ñÊrÜf˜&“ÊvWBÇ&Ê÷R"í«¬$∆ˆ6∆óL:66V«F"íÁG&ñ“Çí¬∆C¢ÁV÷&W"Üf˜&“ÊvWBÇ&∆B"íí¬∆ˆ„¢ÁV÷&W"Üf˜&“ÊvWBÇ&∆ˆ‚"íí“ì∞¢“ì∞¢ñ6∂W"ÁVW'ï6V∆V7F˜"Ç%∂FF÷˜V‚÷gV∆¬÷÷“"ìÚÊFDWfVÁD∆ó7FVÊW"Ç&6∆ñ6≤"¬Çí”‚˜V‰ÜVDgV∆≈67&VV‰÷Ü«î∆ˆ6Fñˆ‚íì∞¢Fˆ7V÷VÁBÊ&ˆGíÊVÊD6Üñ∆Báñ6∂W"ì∞¢ñ6∂W"ÁVW'ï6V∆V7F˜"Ç"Áv˜&∂∆ñ÷FR◊vR÷&6≤"ìÚÊfˆ7W2Çì∞ß–†¶gVÊ7Fñˆ‚˜V‰ÜVDgV∆≈67&VV‰÷ÜˆÂ6V∆V7Bí∞¢6ˆÁ7B˜fW&∆í“Fˆ7V÷VÁBÊ7&VFTV∆V÷VÁBÇ&Fób"ì∞¢˜fW&∆íÊ6∆74Ê÷R“&ÜVB÷gV∆¬÷÷#∞¢˜fW&∆íÊñÊÊW$ÖD‘¬“∆'WGFˆ‚GóS“&'WGFˆ‚"6∆73“&ÜVB÷÷÷6∆˜6R"&ñ÷∆&V√“$6ÜóVFí÷#Ï9s¬ˆ'WGFˆ„„∆Fób6∆73“&ÜVB÷÷÷6Áf2"FF÷÷÷6Áf3„¬ˆFóc„∆Fób6∆73“&ÜVB÷÷◊Fˆˆ∆&"#„«7G&ˆÊsÂ66Vv∆í˜6ó¶ñˆÊR7V∆∆÷¬˜7G&ˆÊs„«7‚FF÷÷÷6ˆ˜&FñÊFW3ÂFˆ66V‚VÁFÛ¬˜7„„∆'WGFˆ‚GóS“&'WGFˆ‚"6∆73“&'F‚'F‚◊&ñ÷'í"FF÷6ˆÊfó&“÷÷Fó6&∆VC‰6ˆÊfW&÷˜6ó¶ñˆÊS¬ˆ'WGFˆ„„¬ˆFócÊ∞¢∆WB6V∆V7FVB“ÁV∆√∞¢∆WBÜVD÷“ÁV∆√∞¢∆WB÷&∂W"“ÁV∆√∞¢6ˆÁ7B6∆˜6T÷“Çí”‚∞¢ñbÜÜVD÷í∞¢ÜVD÷Á&V÷˜fRÇì∞¢ÜVD÷“ÁV∆√∞¢–¢˜fW&∆íÁ&V÷˜fRÇì∞¢”∞¢˜fW&∆íÁVW'ï6V∆V7F˜"Ç"ÊÜVB÷÷÷6∆˜6R"ìÚÊFDWfVÁD∆ó7FVÊW"Ç&6∆ñ6≤"¬6∆˜6T÷ì∞¢˜fW&∆íÁVW'ï6V∆V7F˜"Ç%∂FF÷6ˆÊfó&“÷÷“"ìÚÊFDWfVÁD∆ó7FVÊW"Ç&6∆ñ6≤"¬Çí”‚∞¢ñbÇ6V∆V7FVBí&WGW&„∞¢6∆˜6T÷Çì∞¢ˆÂ6V∆V7Bá6V∆V7FVBì∞¢“ì∞¢Fˆ7V÷VÁBÊ&ˆGíÊVÊD6Üñ∆BÜ˜fW&∆íì∞†¢ñbáGóVˆb¬””“'VÊFVfñÊVB"í∞¢6ˆÁ7Bf∆∆&6≤“˜fW&∆íÁVW'ï6V∆V7F˜"Ç%∂FF÷÷÷6Áf5“"ì∞¢ñbÜf∆∆&6≤íf∆∆&6≤ÁFWáD6ˆÁFVÁB“$÷Êˆ‚Fó7ˆÊñ&ñ∆S¢&ñ6&ñ6∆vñÊR&ó&˜f‚#∞¢&WGW&„∞¢–†¢6ˆÁ7BñÊóFñƒ∆ˆ6Fñˆ‚“6V∆V7FVEvVFÜW$∆ˆ6Fñˆ‚«¬7W'&VÁEvVFÜW%F&vWB«¬ÑTEÙDTdT≈EÙƒÙ4DîÙ„∞¢ÜVD÷“¬Ê÷Ü˜fW&∆íÁVW'ï6V∆V7F˜"Ç%∂FF÷÷÷6Áf5“"í¬∞¢¶ˆˆ‘6ˆÁG&ˆ√¢G'VR¿¢GG&ñ'WFñˆ‰6ˆÁG&ˆ√¢G'VP¢“íÁ6WEfñWrÖ¥ÁV÷&W"ÜñÊóFñƒ∆ˆ6Fñˆ‚Ê∆Bí«¬ÑTEÙDTdT≈EÙƒÙ4DîÙ‚Ê∆B¬ÁV÷&W"ÜñÊóFñƒ∆ˆ6Fñˆ‚Ê∆ˆ‚í«¬ÑTEÙDTdT≈EÙƒÙ4DîÙ‚Ê∆ˆÂ“¬ì∞¢¬ÁFñ∆T∆ñW"Ö5D‰D$EıDîƒUıU$¬¬5D‰D$EıDîƒUÙıDîÙÂ2íÊFEFÚÜÜVD÷ì∞¢6ˆÁ7B6V∆V7EˆñÁB“Ü∆F∆Êrí”‚∞¢6V∆V7FVB“≤Ê÷S¢%˜6ó¶ñˆÊR66V«F7R÷"¬∆C¢∆F∆ÊrÊ∆B¬∆ˆ„¢∆F∆ÊrÊ∆Êr”∞¢ñbÇ÷&∂W"í∞¢÷&∂W"“¬Ê÷&∂W"Ü∆F∆ÊríÊFEFÚÜÜVD÷ì∞¢“V«6R∞¢÷&∂W"Á6WD∆D∆ÊrÜ∆F∆Êrì∞¢–¢˜fW&∆íÁVW'ï6V∆V7F˜"Ç%∂FF÷÷÷6ˆ˜&FñÊFW5“"íÁFWáD6ˆÁFVÁB“G∑6V∆V7FVBÊ∆BÁFÙfóÜVBÉRó“¬G∑6V∆V7FVBÊ∆ˆ‚ÁFÙfóÜVBÉRó÷∞¢˜fW&∆íÁVW'ï6V∆V7F˜"Ç%∂FF÷6ˆÊfó&“÷÷“"íÊFó6&∆VB“f«6S∞¢”∞¢ÜVD÷Êˆ‚Ç&6∆ñ6≤"¬ÜWfVÁBí”‚6V∆V7EˆñÁBÜWfVÁBÊ∆F∆Êríì∞¢6WEFñ÷V˜WBÇÇí”‚ÜVD÷ÚÊñÁf∆ñFFU6ó¶RÇí¬cì∞ß–†¶gVÊ7Fñˆ‚'Vñ∆D6ófñ≈&˜FV7Fñˆ‰∆W'D6ÜóÜ∆W'Bí∞¢6ˆÁ7B∆WfV¬“∆W'CÚÊ∆WfV¬«¬&w&VV‚#∞¢6ˆÁ7B÷WF“ƒU%EÙƒUdT≈Ù‘UD∂∆WfV≈“«¬ƒU%EÙƒUdT≈Ù‘UDÊw&VV„∞¢6ˆÁ7BFWáB“∆W'CÚÊ∆ˆFñÊrÚ%fW&ñfñ6&˜FW¶ñˆÊR6ófñ∆R‚‚‚"¢G∂÷WFÊV÷ˆ¶ó“G∂∆W'CÚÊ∆&V¬«¬÷WFÊ∆&V«÷∞¢&WGW&‚«7‚6∆73“wvVFÜW"◊&ó6≤÷6ÜóG∂÷WFÊ6∆74Ê÷W“rFóF∆S“tgfó6Ú&˜FW¶ñˆÊR6ófñ∆RG∂∆W'CÚÁ&Vvñˆ‚Ú(
"G∂W66TÖD‘¬Ü∆W'BÁ&Vvñˆ‚ó÷¢"'“s‚G∂W66TÖD‘¬áFWáBó”¬˜7„Ê∞ß–†¶gVÊ7Fñˆ‚&VÊFW$6ófñ≈&˜FV7Fñˆ‰∆W'BÜ∆W'Bí∞¢7W'&VÁD6ófñ≈&˜FV7Fñˆ‰∆W'B“≤‚‚Ê7W'&VÁD6ófñ≈&˜FV7Fñˆ‰∆W'B¬‚‚‚Ü∆W'B«¬∑“í”∞¢6ˆÁ7B∆WfV¬“7W'&VÁD6ófñ≈&˜FV7Fñˆ‰∆W'BÊ∆WfV¬«¬&w&VV‚#∞¢6ˆÁ7B6Ü˜t&ÊÊW"“ƒU%EÙƒUdT≈Ù‘UD∂∆WfV≈”ÚÁ&Ê≤„“ƒU%EÙƒUdT≈Ù‘UDÊ˜&ÊvRÁ&Ê≥∞¢VíÁvVFÜW$∆W'D&ÊÊW#ÚÊ6∆74∆ó7BÁFˆvv∆RÇ&ÜñFFV‚"¬6Ü˜t&ÊÊW"ì∞ß–†¶gVÊ7Fñˆ‚˜VÂvVFÜW$WáFW&ÊƒFWFñ¬Çí∞¢6ˆÁ7BF&vWB“7W'&VÁEvVFÜW%F&vWB«¬vWEvVFÜW%F&vWD6ˆ˜&FñÊFW2Çì∞¢6ˆÁ7BÜ46ófñ≈&˜FV7Fñˆ‰∆W'B“ƒU%EÙƒUdT≈Ù‘UD∂7W'&VÁD6ófñ≈&˜FV7Fñˆ‰∆W'CÚÊ∆WfV¬«¬&w&VV‚%“Á&Ê≤‚∞¢6ˆÁ7BW&¬“Ü46ófñ≈&˜FV7Fñˆ‰∆W'@¢ÚÜ7W'&VÁD6ófñ≈&˜FV7Fñˆ‰∆W'BÁW&¬«¬4ïdî≈ı$ıDT5DîÙÂÙƒU%EıtRê¢¢G¥‘UDTıÛ4%Ù$4UıU$«”ˆ∆C“G∂VÊ6ˆFUU$î6ˆ◊ˆÊVÁBáF&vWBÊ∆Bó“f∆ˆ„“G∂VÊ6ˆFUU$î6ˆ◊ˆÊVÁBáF&vWBÊ∆ˆ‚ó÷∞¢vñÊF˜rÊ˜V‚áW&¬¬%ˆ&∆Ê≤"¬&Êˆ˜VÊW"∆Ê˜&VfW'&W""ì∞ß–†¶gVÊ7Fñˆ‚˜VÂvVFÜW$÷ˆF¬Çí∞¢VíÁvVFÜW$÷ˆF¬Ê6∆74∆ó7BÁ&V÷˜fRÇ&ÜñFFV‚"ì∞¢VíÁvVFÜW$÷ˆF¬Á6WDGG&ñ'WFRÇ&&ñ÷ÜñFFV‚"¬&f«6R"ì∞ß–†¶gVÊ7Fñˆ‚6∆˜6UvVFÜW$÷ˆF¬Çí∞¢VíÁvVFÜW$÷ˆF¬Ê6∆74∆ó7BÊFBÇ&ÜñFFV‚"ì∞¢VíÁvVFÜW$÷ˆF¬Á6WDGG&ñ'WFRÇ&&ñ÷ÜñFFV‚"¬'G'VR"ì∞ß–†¶gVÊ7Fñˆ‚Fó7FÊ6Tg&ˆ’W6W"Üñ◊ñÁFÚí∞¢ñbÇ7W'&VÁEW6W%˜2«¬ñ◊ñÁFÚÊw5í”“ÁV∆¬«¬ñ◊ñÁFÚÊw5Ç”“ÁV∆¬í&WGW&‚ÁV÷&W"‰‘Öı4dUÙîÂDTtU#∞¢&WGW&‚ÜfW'6ñÊRÜ7W'&VÁEW6W%˜2Ê∆B¬7W'&VÁEW6W%˜2Ê∆Êr¬ñ◊ñÁFÚÊw5í¬ñ◊ñÁFÚÊw5Çì∞ß–†¶gVÊ7Fñˆ‚ÜfW'6ñÊRÜ∆C¬∆ˆ„¬∆C"¬∆ˆ„"í∞¢6ˆÁ7BFı&B“ÜFVrí”‚ÜFVr¢÷FÇÂííÚÉ∞¢6ˆÁ7B"“c3s∞¢6ˆÁ7BD∆B“Fı&BÜ∆C"“∆Cì∞¢6ˆÁ7BD∆ˆ‚“Fı&BÜ∆ˆ„"“∆ˆ„ì∞¢6ˆÁ7B“÷FÇÁ6ñ‚ÜD∆BÚ"í¢¢ ¢≤÷FÇÊ6˜2áFı&BÜ∆Cíí¢÷FÇÊ6˜2áFı&BÜ∆C"íí¢÷FÇÁ6ñ‚ÜD∆ˆ‚Ú"í¢¢#∞¢&WGW&‚"¢"¢÷FÇÊ6ñ‚Ñ÷FÇÁ7'BÜíì∞ß–†¶gVÊ7Fñˆ‚f˜&÷DFó7FÊ6RÜ∂“í∞¢ñbÇÁV÷&W"Êó4fñÊóFRÜ∂“í«¬∂“‚Sí&WGW&‚$‚ÙB#∞¢ñbÜ∂“¬í&WGW&‚G¥÷FÇÁ&˜VÊBÜ∂“¢ó“÷∞¢&WGW&‚G∂∂“ÁFÙfóÜVBÉ"ó“∂÷∞ß–†¶gVÊ7Fñˆ‚vWEG&ffñ4ñÁFVÁ6óGî'îÜ˜W"ÜFFR“ÊWrFFRÇíí∞¢6ˆÁ7BÜ˜W"“FFRÊvWDÜ˜W'2Çì∞¢6ˆÁ7BFí“FFRÊvWDFíÇì∞¢6ˆÁ7Bó5vVV∂VÊB“Fí””“«¬Fí””“c∞†¢ñbÇó5vVV∂VÊBbbÇÜÜ˜W"„“rbbÜ˜W"√“íí«¬ÜÜ˜W"„“rbbÜ˜W"√“íííí&WGW&‚&ñÁFVÁ6Ú#∞¢ñbÜÜ˜W"„“#"«¬Ü˜W"√“Rí&WGW&‚&∆VvvW&Ú#∞¢ñbÜó5vVV∂VÊBbbÜ˜W"„“"bbÜ˜W"√“Bí&WGW&‚&÷ˆFW&FÚ#∞¢&WGW&‚&÷ˆFW&FÚ#∞ß–†¶gVÊ7Fñˆ‚vWDFó7FÊ6TñÁFVÁ6óGîˆfg6WBÜFó7FÊ6T∂“í∞¢ñbÇÁV÷&W"Êó4fñÊóFRÜFó7FÊ6T∂“íí&WGW&‚∞¢ñbÜFó7FÊ6T∂“¬„"í&WGW&‚”∞¢ñbÜFó7FÊ6T∂“¬bí&WGW&‚∞¢ñbÜFó7FÊ6T∂“¬#í&WGW&‚∞¢&WGW&‚∞ß–†¶gVÊ7Fñˆ‚vWE&˜WFUf&ñÊ6Tˆfg6WBÜFó7FÊ6T∂“í∞¢ñbÇÁV÷&W"Êó4fñÊóFRÜFó7FÊ6T∂“íí&WGW&‚∞¢6ˆÁ7BfñÊvW'&ñÁB“÷FÇÊf∆ˆ˜"ÜFó7FÊ6T∂“¢íRS∞¢ñbÜfñÊvW'&ñÁB””“í&WGW&‚”∞¢ñbÜfñÊvW'&ñÁB””“Bí&WGW&‚∞¢&WGW&‚∞ß–†¶gVÊ7Fñˆ‚Ê˜&÷∆ó¶UG&ffñ4ñÁFVÁ6óGíÜ&6TñÁFVÁ6óGí¬Fó7FÊ6T∂“í∞¢6ˆÁ7B∆WfV«2“≤&∆VvvW&Ú"¬&÷ˆFW&FÚ"¬&ñÁFVÁ6Ú%”∞¢6ˆÁ7B&6TñÊFWÇ“∆WfV«2ÊñÊFWÑˆbÜ&6TñÁFVÁ6óGíì∞¢ñbÜ&6TñÊFWÇ””“”í&WGW&‚&÷ˆFW&FÚ#∞†¢6ˆÁ7BFó7FÊ6Tˆfg6WB“vWDFó7FÊ6TñÁFVÁ6óGîˆfg6WBÜFó7FÊ6T∂“ì∞¢6ˆÁ7Bf&ñÊ6Tˆfg6WB“vWE&˜WFUf&ñÊ6Tˆfg6WBÜFó7FÊ6T∂“ì∞¢6ˆÁ7BÊ˜&÷∆ó¶VDñÊFWÇ“÷FÇÊ÷ÇÉ¬÷FÇÊ÷ñ‚Ü∆WfV«2Ê∆VÊwFÇ“¬&6TñÊFWÇ≤Fó7FÊ6Tˆfg6WB≤f&ñÊ6Tˆfg6WBíì∞¢&WGW&‚∆WfV«5∂Ê˜&÷∆ó¶VDñÊFWÖ”∞ß–†¶gVÊ7Fñˆ‚W7Fñ÷FUG&fVƒ÷WFÜFó7FÊ6T∂“í∞¢ñbÇÁV÷&W"Êó4fñÊóFRÜFó7FÊ6T∂“í«¬Fó7FÊ6T∂“‚Sí∞¢&WGW&‚≤ñÁFVÁ6óGî∂Wì¢&Ê"¬ñÁFVÁ6óGî∆&V√¢$‚ÙB"¬WF∆&V√¢$‚ÙB"”∞¢–¢6ˆÁ7B&6TñÁFVÁ6óGí“vWEG&ffñ4ñÁFVÁ6óGî'îÜ˜W"Çì∞¢6ˆÁ7BñÁFVÁ6óGî∂Wí“Ê˜&÷∆ó¶UG&ffñ4ñÁFVÁ6óGíÜ&6TñÁFVÁ6óGí¬Fó7FÊ6T∂“ì∞¢6ˆÁ7BñÁFVÁ6óGî∆&V¬“ñÁFVÁ6óGî∂WíÊ6Ü$BÉíÁFıWW$66RÇí≤ñÁFVÁ6óGî∂WíÁ6∆ñ6RÉì∞¢6ˆÁ7B7VVD'îñÁFVÁ6óGí“∞¢ñÁFVÁ6Û¢#R¿¢÷ˆFW&FÛ¢C¿¢∆VvvW&Û¢c ¢”∞¢6ˆÁ7Bfu7VVB“7VVD'îñÁFVÁ6óGï∂ñÁFVÁ6óGî∂Wï“«¬3S∞¢6ˆÁ7BWF÷ñÁWFW2“÷FÇÊ÷ÇÉ¬÷FÇÁ&˜VÊBÇÑ÷FÇÊ÷ÇÜFó7FÊ6T∂“¬íÚfu7VVBí¢cíì∞¢&WGW&‚∞¢ñÁFVÁ6óGî∂Wí¿¢ñÁFVÁ6óGî∆&V¬¿¢WF∆&V√¢G∂WF÷ñÁWFW7“÷ñÊ ¢”∞ß–†¶gVÊ7Fñˆ‚W66TÖD‘¬áf«VRí∞¢&WGW&‚7G&ñÊráf«VR«¬""ê¢Á&W∆6T∆¬Ç"b"¬"f◊≤"ê¢Á&W∆6T∆¬Ç#¬"¬"f«C≤"ê¢Á&W∆6T∆¬Ç#‚"¬"fwC≤"ê¢Á&W∆6T∆¬Çr"r¬"gV˜C≤"ê¢Á&W∆6T∆¬Ç"r"¬"b33ì≤"ì∞ß–†¶gVÊ7Fñˆ‚7V'67&ñ&T6ÜBÇí∞¢6ÜDÊ˜Fñfñ6FñˆÁ4ñÊóFñ∆ó¶VB“f«6S∞¢VÁ7V'67&ñ&T6ÜB“F ¢Ê6ˆ∆∆V7Fñˆ‚Ç&6ÜD÷W76vW2"ê¢Ê˜&FW$'íÇ&7&VFVDB"¬&FW62"ê¢Ê∆ñ÷óBÉSê¢ÊˆÂ6Ê6Ü˜BÇá6Ê6Ü˜Bí”‚∞¢6ÜD÷W76vW2“6Ê6Ü˜BÊFˆ70¢Ê÷ÇÜFˆ2í”‚á≤ñC¢Fˆ2ÊñB¬‚‚ÊFˆ2ÊFFÇí“íê¢Á&WfW'6RÇì∞¢Ê˜Fñgîf˜$ÊWt6ÜD÷W76vW2á6Ê6Ü˜BÊFˆ46ÜÊvW2Çíì∞¢&VÊFW$6ÜBÜ6ÜD÷W76vW2ì∞¢“¬ÜW'&˜"í”‚∞¢6ˆÁ6ˆ∆RÊW'&˜"ÜW'&˜"ì∞¢VíÊ6ÜDfVVF&6≤ÁFWáD6ˆÁFVÁB“$W'&˜&R6&ñ6÷VÁFÚ6ÜB‚#∞¢“ì∞ß–†¶7ñÊ2gVÊ7Fñˆ‚Ê˜Fñgîf˜$ÊWt6ÜD÷W76vW2Ü6ÜÊvW2“µ“í∞¢ñbÇ6ÜDÊ˜Fñfñ6FñˆÁ4ñÊóFñ∆ó¶VBí∞¢6ÜDÊ˜Fñfñ6FñˆÁ4ñÊóFñ∆ó¶VB“G'VS∞¢&WGW&„∞¢–¢6ˆÁ7BFFVD÷W76vW2“6ÜÊvW0¢Êfñ«FW"ÇÜ6ÜÊvRí”‚6ÜÊvRÁGóR””“&FFVB"ê¢Ê÷ÇÜ6ÜÊvRí”‚á≤ñC¢6ÜÊvRÊFˆ2ÊñB¬‚‚Ê6ÜÊvRÊFˆ2ÊFFÇí“íê¢Êfñ«FW"ÇÜ÷W76vRí”‚6‰Ê˜Fñgîf˜$6ÜD÷W76vRÜ÷W76vRíì∞†¢f˜"Ü6ˆÁ7B÷W76vRˆbFFVD÷W76vW2í∞¢6ˆÁ7B6VÊFW$Ê÷R“7G&ñÊrÜ÷W76vRÁ6VÊFW$Ê÷R«¬$˜W&F˜&R"íÁG&ñ“Çì∞¢6ˆÁ7B&ˆGí“vWD6ÜDÊ˜Fñfñ6Fñˆ‰&ˆGíÜ÷W76vRì∞¢G'í∞¢vóB6Ü˜t∆ˆ6ƒÊ˜Fñfñ6Fñˆ‚ÜÁV˜fÚ÷W76vvñÚFG∑6VÊFW$Ê÷W÷¬∞¢&ˆGí¿¢Fs¢ÜW&÷6ÜB“G∂÷W76vRÊñG÷¿¢FF¢≤W&√¢"‚ˆñÊFWÇÊáF÷¬66ÜB"–¢“ì∞¢“6F6ÇÜW'&˜"í∞¢6ˆÁ6ˆ∆RÁv&‚Ç$ñÁfñÚÊ˜Fñfñ66ÜBÊˆ‚&óW66óFÛ¢"¬W'&˜"ì∞¢–¢–ß–†¶gVÊ7Fñˆ‚6‰Ê˜Fñgîf˜$6ÜD÷W76vRÜ÷W76vRí∞¢ñbÇ7W'&VÁEW6W"«¬ó4˜v‰÷W76vRÜ÷W76vRíí&WGW&‚f«6S∞¢ñbÇ6ÂfñWt÷W76vRÜ÷W76vRí«¬ó46ÜD÷W76vTg&W6ÇÜ÷W76vRíí&WGW&‚f«6S∞¢ñbÇFˆ7V÷VÁBÊÜñFFV‚bbVíÊ6ÜD÷ˆF¬bbVíÊ6ÜD÷ˆF¬Ê6∆74∆ó7BÊ6ˆÁFñÁ2Ç&ÜñFFV‚"íí&WGW&‚f«6S∞¢&WGW&‚G'VS∞ß–†¶gVÊ7Fñˆ‚vWD6ÜDÊ˜Fñfñ6Fñˆ‰&ˆGíÜ÷W76vRí∞¢6ˆÁ7BFWáB“7G&ñÊrÜ÷W76vRÁFWáB«¬÷W76vRÊ÷W76vR«¬÷W76vRÊ&ˆGí«¬÷W76vRÊ6ˆÁFVÁB«¬""íÁG&ñ“Çì∞¢ñbáFWáBí&WGW&‚FWáBÊ∆VÊwFÇ‚ÉÚG∑FWáBÁ6∆ñ6RÉ¬sró“‚‚Ê¢FWáC∞¢6ˆÁ7BGóR“7G&ñÊrÜ÷W76vRÁGóR«¬""íÁFÙ∆˜vW$66RÇì∞¢ñbáGóR””“&ñ÷vR"í&WGW&‚$ÜñÁfñFÚVÊf˜FÚ‚#∞¢ñbáGóR””“'fñFVÚ"í&WGW&‚$ÜñÁfñFÚV‚fñFVÚ‚#∞¢ñbáGóR””“'fˆñ6R"í&WGW&‚$ÜñÁfñFÚV‚÷W76vvñÚfˆ6∆R‚#∞¢&WGW&‚$ÜíV‚ÁV˜fÚ÷W76vvñÚñ‚6ÜB‚#∞ß–†¶gVÊ7Fñˆ‚ó46ÜD÷W76vTg&W6ÇÜ÷W76vRí∞¢6ˆÁ7BWáó&W4D◊2“vWD6ÜD÷W76vTWáó'î◊2Ü÷W76vRì∞¢ñbÇWáó&W4D◊2í&WGW&‚G'VS∞¢&WGW&‚Wáó&W4D◊2„“FFRÊÊ˜rÇì∞ß–†¶gVÊ7Fñˆ‚vWD6ÜD÷W76vT7&VFVDD◊2Ü÷W76vRí∞¢ñbÜ÷W76vSÚÊ7&VFVDBbbGóVˆb÷W76vRÊ7&VFVDBÁFÙFFR””“&gVÊ7Fñˆ‚"í∞¢&WGW&‚÷W76vRÊ7&VFVDBÁFÙFFRÇíÊvWEFñ÷RÇì∞¢–¢&WGW&‚∞ß–†¶gVÊ7Fñˆ‚vWD6ÜD÷W76vTWáó'î◊2Ü÷W76vRí∞¢ñbÜ÷W76vSÚÊWáó&W4BbbGóVˆb÷W76vRÊWáó&W4BÁFÙFFR””“&gVÊ7Fñˆ‚"í∞¢&WGW&‚÷W76vRÊWáó&W4BÁFÙFFRÇíÊvWEFñ÷RÇì∞¢–¢6ˆÁ7B7&VFVDD◊2“vWD6ÜD÷W76vT7&VFVDD◊2Ü÷W76vRì∞¢ñbÇ7&VFVDD◊2í&WGW&‚∞¢&WGW&‚7&VFVDD◊2≤4ÑEı$UDTÂDîÙÂÙ’3∞ß–†¶gVÊ7Fñˆ‚7F'D6ÜE&WFVÁFñˆ‰∆ˆ˜Çí∞¢7F˜6ÜE&WFVÁFñˆ‰∆ˆ˜Çì∞¢W&vTˆ∆D6ÜD÷W76vW2Çì∞¢6ÜE&WFVÁFñˆÂFñ÷W"“6WDñÁFW'f¬ÇÇí”‚∞¢W&vTˆ∆D6ÜD÷W76vW2Çì∞¢“¬c¢c¢ì∞ß–†¶gVÊ7Fñˆ‚7F˜6ÜE&WFVÁFñˆ‰∆ˆ˜Çí∞¢ñbÜ6ÜE&WFVÁFñˆÂFñ÷W"í∞¢6∆V$ñÁFW'f¬Ü6ÜE&WFVÁFñˆÂFñ÷W"ì∞¢6ÜE&WFVÁFñˆÂFñ÷W"“ÁV∆√∞¢–ß–†¶7ñÊ2gVÊ7Fñˆ‚W&vTˆ∆D6ÜD÷W76vW2Çí∞¢ñbÇ6‰÷ÊvTFFÇíí&WGW&„∞¢6ˆÁ7B7WFˆfdFFR“ÊWrFFRÑFFRÊÊ˜rÇí“4ÑEı$UDTÂDîÙÂÙ’2ì∞¢6ˆÁ7BÊ˜tFFR“ÊWrFFRÇì∞¢G'í∞¢6ˆÁ7B∂∆Vv7ï6Ê6Ü˜B¬Wáó&W56Ê6Ü˜E““vóB&ˆ÷ó6RÊ∆¬Ö∞¢F ¢Ê6ˆ∆∆V7Fñˆ‚Ç&6ÜD÷W76vW2"ê¢ÁvÜW&RÇ&7&VFVDB"¬#√“"¬fó&V&6RÊfó&W7F˜&RÂFñ÷W7F◊Êg&ˆ‘FFRÜ7WFˆfdFFRíê¢Ê∆ñ÷óBÉ#ê¢ÊvWBÇí¿¢F ¢Ê6ˆ∆∆V7Fñˆ‚Ç&6ÜD÷W76vW2"ê¢ÁvÜW&RÇ&Wáó&W4B"¬#√“"¬fó&V&6RÊfó&W7F˜&RÂFñ÷W7F◊Êg&ˆ‘FFRÜÊ˜tFFRíê¢Ê∆ñ÷óBÉ#ê¢ÊvWBÇê¢“ì∞¢6ˆÁ7BFˆ75FÙFV∆WFR“ÊWr÷Çì∞¢∆Vv7ï6Ê6Ü˜BÊFˆ72Êf˜$V6ÇÇÜFˆ2í”‚∞¢6ˆÁ7BFF“Fˆ2ÊFFÇí«¬∑”∞¢ñbÜFFÊWáó&W4BbbGóVˆbFFÊWáó&W4BÁFÙFFR””“&gVÊ7Fñˆ‚"í&WGW&„∞¢Fˆ75FÙFV∆WFRÁ6WBÜFˆ2ÊñB¬Fˆ2ì∞¢“ì∞¢Wáó&W56Ê6Ü˜BÊFˆ72Êf˜$V6ÇÇÜFˆ2í”‚Fˆ75FÙFV∆WFRÁ6WBÜFˆ2ÊñB¬Fˆ2íì∞¢ñbÇFˆ75FÙFV∆WFRÁ6ó¶Rí&WGW&„∞¢6ˆÁ7B&F6Ç“F"Ê&F6ÇÇì∞¢Fˆ75FÙFV∆WFRÊf˜$V6ÇÇÜFˆ2í”‚&F6ÇÊFV∆WFRÜFˆ2Á&Vbíì∞¢vóB&F6ÇÊ6ˆ÷÷óBÇì∞¢“6F6ÇÜW'&˜"í∞¢6ˆÁ6ˆ∆RÁv&‚Ç%V∆ó¶ñ6ÜB#FÇÊˆ‚6ˆ◊∆WFF¢"¬W'&˜"ì∞¢–ß–†¶gVÊ7Fñˆ‚7F'DÜ˜W'4FVF∆ñÊT∆W'D∆ˆ˜Çí∞¢7F˜Ü˜W'4FVF∆ñÊT∆W'D∆ˆ˜Çì∞¢6ÜV6¥ÊE6VÊDÜ˜W'4FVF∆ñÊT∆W'G2Çì∞¢Ü˜W'4FVF∆ñÊT∆W'EFñ÷W"“6WDñÁFW'f¬ÇÇí”‚∞¢6ÜV6¥ÊE6VÊDÜ˜W'4FVF∆ñÊT∆W'G2Çì∞¢“¬R¢c¢ì∞ß–†¶gVÊ7Fñˆ‚7F˜Ü˜W'4FVF∆ñÊT∆W'D∆ˆ˜Çí∞¢ñbÜÜ˜W'4FVF∆ñÊT∆W'EFñ÷W"í∞¢6∆V$ñÁFW'f¬ÜÜ˜W'4FVF∆ñÊT∆W'EFñ÷W"ì∞¢Ü˜W'4FVF∆ñÊT∆W'EFñ÷W"“ÁV∆√∞¢–ß–†¶gVÊ7Fñˆ‚Ü57VG&U&˜w4f˜$FFá7VDFFí∞¢6ˆÁ7B&˜w2“'&íÊó4'&íá7VDFFÚÁ7VG&RíÚ7VDFFÁ7VG&R¢vWD∆Vv7ï7VG&U&˜w2á7VDFF«¬∑“ì∞¢&WGW&‚&˜w2Á6ˆ÷RÇá&˜rí”‚7G&ñÊrá&˜sÚÁW'6ˆÊ∆R«¬""íÁG&ñ“Çí«¬7G&ñÊrá&˜sÚÊ÷Wß¶í«¬""íÁG&ñ“Çíì∞ß–†¶gVÊ7Fñˆ‚Ü4Ü˜W'4f˜$6ˆ÷÷W76ñ‰VÁG&ñW2ÜVÁG&ñW2¬6ˆ÷÷W76ñBí∞¢ñbÇ'&íÊó4'&íÜVÁG&ñW2í«¬6ˆ÷÷W76ñBí&WGW&‚f«6S∞¢&WGW&‚VÁG&ñW2Á6ˆ÷RÇÜVÁG'íí”‚∞¢ñbÖ7G&ñÊrÜVÁG'ìÚÊ6ˆ÷÷W76ñB«¬""í”“7G&ñÊrÜ6ˆ÷÷W76ñBíí&WGW&‚f«6S∞¢&WGW&‚'&íÊó4'&íÜVÁG'íÁ&˜w2íbbVÁG'íÁ&˜w2Á6ˆ÷RÇá&˜rí”‚ÁV÷&W"á&˜sÚÊ˜&R«¬í‚ì∞¢“ì∞ß–†¶7ñÊ2gVÊ7Fñˆ‚6ÜV6¥ÊE6VÊDÜ˜W'4FVF∆ñÊT∆W'G2Çí∞¢ñbÇ7W'&VÁEW6W"«¬6‰÷ÊvTFFÇíí&WGW&„∞¢6ˆÁ7BÊ˜r“ÊWrFFRÇì∞¢ñbÜÊ˜rÊvWDÜ˜W'2Çí¬ÑıU%5ÙDTDƒî‰UÙƒU%EÙÑıU"í&WGW&„∞¢6ˆÁ7BFFT∂Wí“vWDFFT∂Wîg&ˆ‘∆ˆ6ƒFFRÜÊ˜rì∞¢6ˆÁ7B7F˜&ñ6ÙFVƒvñ˜&ÊÚ“7VG&TÜó7F˜'î'îFFRÊvWBÜFFT∂Wíí«¬ÊWr÷Çì∞¢6ˆÁ7B6ˆ÷÷W76T6ˆÂ7VG&“'&íÊg&ˆ“á7F˜&ñ6ÙFVƒvñ˜&ÊÚÁf«VW2Çíê¢Êfñ«FW"Çá7VDFFí”‚Ü57VG&U&˜w4f˜$FFá7VDFFíê¢Ê÷Çá7VDFFí”‚á∞¢6ˆ÷÷W76ñC¢7G&ñÊrá7VDFFÊ6ˆ÷÷W76ñB«¬""íÁG&ñ“Çí¿¢6ˆ÷÷W76Ê÷S¢7G&ñÊrá7VDFFÊ6ˆ÷÷W76Êˆ÷R«¬""íÁG&ñ“Çí«¬Ü6ˆ÷÷W76T'îñBÊvWBÖ7G&ñÊrá7VDFFÊ6ˆ÷÷W76ñB«¬""íÁG&ñ“Çíí«¬∑“íÊÊˆ÷R«¬$6ˆ÷÷W76 ¢“íê¢Êfñ«FW"Çá&˜rí”‚&˜rÊ6ˆ÷÷W76ñBì∞¢ñbÇ6ˆ÷÷W76T6ˆÂ7VG&Ê∆VÊwFÇí&WGW&„∞†¢6ˆÁ7B∑&W˜'G56Ê6Ü˜B¬&˜f≈6Ê6Ü˜E““vóB&ˆ÷ó6RÊ∆¬Ö∞¢F"Ê6ˆ∆∆V7Fñˆ‚ÜvWD˜&U&W˜'G46ˆ∆∆V7Fñˆ‰Ê÷RÇííÁvÜW&RÇ&FFR"¬#”“"¬FFT∂WííÊvWBÇí¿¢F"Ê6ˆ∆∆V7Fñˆ‚ÜvWD˜&T&˜f≈&WVW7G46ˆ∆∆V7Fñˆ‰Ê÷RÇííÁvÜW&RÇ&FFR"¬#”“"¬FFT∂WííÊvWBÇê¢“ì∞†¢6ˆÁ7B&W˜'G2“&W˜'G56Ê6Ü˜BÊFˆ72Ê÷ÇÜFˆ2í”‚á≤ñC¢Fˆ2ÊñB¬‚‚ÊFˆ2ÊFFÇí“íì∞¢6ˆÁ7B&˜f«2“&˜f≈6Ê6Ü˜BÊFˆ70¢Ê÷ÇÜFˆ2í”‚á≤ñC¢Fˆ2ÊñB¬‚‚ÊFˆ2ÊFFÇí“íê¢Êfñ«FW"Çá&WVW7Bí”‚7G&ñÊrá&WVW7BÁ7FGW2«¬""íÁG&ñ“Çí”“'&V¶V7FVB"ì∞†¢f˜"Ü6ˆÁ7B6ˆ÷÷W76ˆb6ˆ÷÷W76T6ˆÂ7VG&í∞¢6ˆÁ7BÜ4Ü˜W'56fVB“&W˜'G2Á6ˆ÷RÇá&W˜'Bí”‚Ü4Ü˜W'4f˜$6ˆ÷÷W76ñ‰VÁG&ñW2á&W˜'BÊVÁG&ñW2¬6ˆ÷÷W76Ê6ˆ÷÷W76ñBíì∞¢6ˆÁ7BÜ4Ü˜W'5VÊFñÊr“&˜f«2Á6ˆ÷RÇá&WVW7Bí”‚Ü4Ü˜W'4f˜$6ˆ÷÷W76ñ‰VÁG&ñW2á&WVW7BÊVÁG&ñW2¬6ˆ÷÷W76Ê6ˆ÷÷W76ñBíì∞¢ñbÜÜ4Ü˜W'56fVB«¬Ü4Ü˜W'5VÊFñÊrí6ˆÁFñÁVS∞¢vóB6VÊDÜ˜W'4FVF∆ñÊT∆W'Dñd÷ó76ñÊrá∞¢6ˆ÷÷W76ñC¢6ˆ÷÷W76Ê6ˆ÷÷W76ñB¿¢6ˆ÷÷W76Ê÷S¢6ˆ÷÷W76Ê6ˆ÷÷W76Ê÷R¿¢FFT∂Wê¢“ì∞¢–ß–†¶7ñÊ2gVÊ7Fñˆ‚6VÊDÜ˜W'4FVF∆ñÊT∆W'Dñd÷ó76ñÊrá≤6ˆ÷÷W76ñB¬6ˆ÷÷W76Ê÷R¬FFT∂Wí“í∞¢ñbÇ6ˆ÷÷W76ñB«¬FFT∂Wíí&WGW&„∞¢6ˆÁ7B∆W'DñB“G∂FFT∂Wó’ıÚG∂6ˆ÷÷W76ñG’ıˆ∆∆∞¢6ˆÁ7B∆W'E&Vb“F"Ê6ˆ∆∆V7Fñˆ‚Ç&Ü˜W'4FVF∆ñÊT∆W'G2"íÊFˆ2Ü∆W'DñBì∞¢6ˆÁ7BWÜó7FñÊr“vóB∆W'E&VbÊvWBÇì∞¢ñbÜWÜó7FñÊrÊWÜó7G2í&WGW&„∞†¢6ˆÁ7BFFT∆&V¬“ÊWrFFRÜG∂FFT∂Wó’C££íÁFÙ∆ˆ6∆TFFU7G&ñÊrÇ&óB‘ïB"ì∞¢6ˆÁ7BFWáB“)™˚àÚgfó6Ú˜&R÷Ê6ÁFì¢W"∆6ˆ÷÷W76G∂6ˆ÷÷W76Ê÷R«¬$6ˆ÷÷W76'“ÇG∂FFT∆&V«“íÊˆ‚&ó7V«FÊÚ˜&RñÁ6W&óFRVÁG&Ú∆Rì£Ê∞¢6ˆÁ7BWáó&W4B“ÊWrFFRÑFFRÊÊ˜rÇí≤ÑıU%5ÙDTDƒî‰UÙƒU%Eı$UDTÂDîÙÂÙ’2ì∞¢6ˆÁ7B6ÜDFˆ5&Vb“vóBF"Ê6ˆ∆∆V7Fñˆ‚Ç&6ÜD÷W76vW2"íÊFBá∞¢GóS¢'FWáB"¿¢FWáB¿¢&V6óñVÁDñC¢""¿¢6VÊFW$ñC¢'7ó7FV“"¿¢6VÊFW$Ê÷S¢%6ó7FV÷˜&R"¿¢6VÊFW$V÷ñ√¢""¿¢∂ñÊC¢'7ó7FV“"¿¢÷WFFF¢∞¢GóS¢&Ü˜W'5ˆFVF∆ñÊUˆ∆W'B"¿¢7Fñˆ„¢&˜VÂˆÜ˜W'2"¿¢6ˆ÷÷W76ñB¿¢6ˆ÷÷W76Ê÷S¢6ˆ÷÷W76Ê÷R«¬""¿¢FFS¢FFT∂Wê¢“¿¢Wáó&W4C¢fó&V&6RÊfó&W7F˜&RÂFñ÷W7F◊Êg&ˆ‘FFRÜWáó&W4Bí¿¢7&VFVDC¢fó&V&6RÊfó&W7F˜&R‰fñV∆Ef«VRÁ6W'fW%Fñ÷W7F◊Çê¢“ì∞†¢vóBF"Ê6ˆ∆∆V7Fñˆ‚Ç&Ê˜Fñfñ6FñˆÁ2"íÊFBá∞¢WfVÁEGóS¢&Ü˜W'2÷FVF∆ñÊR÷÷ó76ñÊr"¿¢FóF∆S¢$˜&R÷Ê6ÁFíVÁG&Ú∆Rí"¿¢&ˆGì¢FWáB¿¢6ˆ÷÷W76ñB¿¢6ˆ÷÷W76Ê÷S¢6ˆ÷÷W76Ê÷R«¬""¿¢ñ◊ñÁFÙÊ÷S¢""¿¢ñ◊ñÁFÙ∂Wì¢""¿¢7&VFVD'ïVñC¢'7ó7FV“"¿¢7&VFVD'îÊ÷S¢%6ó7FV÷˜&R"¿¢7&VFVD'îV÷ñ√¢""¿¢7&VFVDC¢fó&V&6RÊfó&W7F˜&R‰fñV∆Ef«VRÁ6W'fW%Fñ÷W7F◊Çê¢“ì∞†¢vóB∆W'E&VbÁ6WBá∞¢∆W'DñB¿¢&V6óñVÁC¢&∆¬"¿¢6ˆ÷÷W76ñB¿¢6ˆ÷÷W76Ê÷S¢6ˆ÷÷W76Ê÷R«¬""¿¢FFS¢FFT∂Wí¿¢6ÜD÷W76vTñC¢6ÜDFˆ5&VbÊñB¿¢7&VFVDC¢fó&V&6RÊfó&W7F˜&R‰fñV∆Ef«VRÁ6W'fW%Fñ÷W7F◊Çê¢“¬≤÷W&vS¢G'VR“ì∞ß–†¶7ñÊ2gVÊ7Fñˆ‚W6W'D7W'&VÁE∆Ff˜&’W6W"Çí∞¢ñbÇ7W'&VÁEW6W"í&WGW&„∞¢6ˆÁ7Bó57WW$F÷ñ‚“ó4'Vñ«DñÂ7WW$F÷ñ‰V÷ñ¬Ü7W'&VÁEW6W"ÊV÷ñ¬ì∞¢6ˆÁ7Bó4F÷ñÂW6W"“6‰÷ÊvTFFÇì∞¢6ˆÁ7B&ˆfñ∆UF6Ç“∞¢VñC¢7W'&VÁEW6W"ÁVñB¿¢V÷ñ√¢7W'&VÁEW6W"ÊV÷ñ¬«¬""¿¢Fó7∆îÊ÷S¢7W'&VÁEW6W"ÊFó7∆îÊ÷R«¬7W'&VÁEW6W"ÊV÷ñ¬«¬%WFVÁFR"¿¢Ü˜FıU$√¢7W'&VÁEW6W"ÁÜ˜FıU$¬«¬""¿¢&˜fñFW$ñC¢'&íÊó4'&íÜ7W'&VÁEW6W"Á&˜fñFW$FFê¢Ú7G&ñÊrÜ7W'&VÁEW6W"Á&˜fñFW$FFÊfñÊBÇá&˜fñFW"í”‚&˜fñFW#ÚÁ&˜fñFW$ñB””“&vˆˆv∆RÊ6ˆ“"ìÚÁ&˜fñFW$ñB«¬""ê¢¢""¿¢V÷ñ≈fW&ñfñVC¢7W'&VÁEW6W"ÊV÷ñ≈fW&ñfñVB””“G'VR¿¢ó4F÷ñ„¢ó4F÷ñÂW6W"¿¢‚‚‚Üó57WW$F÷ñ‚Ú∞¢&ˆ∆S¢'7WW%ˆF÷ñ‚"¿¢'Vˆ∆Û¢'7WW%ˆF÷ñ‚"¿¢&ñ∆óFFÛ¢G'VR¿¢VÊ&∆VC¢G'VR¿¢&˜fVC¢G'VP¢“¢∑“í¿¢Ê˜Fñfñ6FñˆÁ4WFÙVÊ&∆VC¢ó4WFÙÊ˜Fñfñ6Fñˆ‰VÊ&∆VBÇí¿¢∆7E6VV‰C¢fó&V&6RÊfó&W7F˜&R‰fñV∆Ef«VRÁ6W'fW%Fñ÷W7F◊Çê¢”∞¢vóBF"Ê6ˆ∆∆V7Fñˆ‚Ç'∆Ff˜&’W6W'2"íÊFˆ2Ü7W'&VÁEW6W"ÁVñBíÁ6WBá&ˆfñ∆UF6Ç¬≤÷W&vS¢G'VR“ì∞¢vóB6fUW'6ó7FVE6W76ñˆ‚Ü7W'&VÁEW6W"¬&ˆfñ∆UF6Çì∞ß–†¶gVÊ7Fñˆ‚7V'67&ñ&TF÷ñÂW6W'2Çí∞¢VÁ7V'67&ñ&TF÷ñÂW6W'2“F"Ê6ˆ∆∆V7Fñˆ‚Ç&6ˆÊfñr"íÊFˆ2Ç&F÷ñÂW6W'2"íÊˆÂ6Ê6Ü˜BÇÜFˆ2í”‚∞¢6ˆÁ7BFF“Fˆ2ÊWÜó7G2ÚFˆ2ÊFFÇí¢∑”∞¢6ˆÁ7B&t∆ó7B“'&íÊó4'&íÜFFÊV÷ñ«2íÚFFÊV÷ñ«2¢µ”∞¢6ˆÁ7BÊ˜&÷∆ó¶VB“&t∆ó7@¢Ê÷ÇÜV÷ñ¬í”‚Ê˜&÷∆ó¶TV÷ñ¬ÜV÷ñ¬íê¢Êfñ«FW"Ñ&ˆˆ∆V‚ì∞¢F÷ñ‰V÷ñ«2“ÊWr6WBÖ≤‚‚‰%Tî≈EÙîÂı5UU%ÙD‘îÂÙT‘î≈2Ê÷ÇÜV÷ñ¬í”‚Ê˜&÷∆ó¶TV÷ñ¬ÜV÷ñ¬íí¬‚‚ÊÊ˜&÷∆ó¶VE“ì∞¢WFFTF÷ñ‰6ˆÁG&ˆ«2Çì∞¢7V'67&ñ&U˜4Fˆ7V÷VÁG2Çì∞¢&VÊFW$6ˆ÷÷W76T÷ÊvV÷VÁD∆ó7BÇì∞¢&VÊFW$F÷ñÂW6W'2Çì∞¢ñbÜ7W'&VÁEW6W"í∞¢fˆñB6fUW'6ó7FVE6W76ñˆ‚Ü7W'&VÁEW6W"¬≤ó4F÷ñ„¢6‰÷ÊvTFFÇí¬&ˆ∆S¢6‰÷ÊvTFFÇíÚ&F÷ñ‚"¢'W6W""“ì∞¢7V'67&ñ&UW6W'2Çì∞¢7V'67&ñ&T˜W&F˜%˜6óFñˆÁ2Çì∞¢7V'67&ñ&U&ˆw&÷÷¶ñˆÊíÇì∞¢–¢“¬ÜW'&˜"í”‚∞¢6ˆÁ6ˆ∆RÊW'&˜"Ç$W'&˜&R6&ñ6÷VÁFÚF÷ñ‚W6W'3¢"¬W'&˜"ì∞¢F÷ñ‰V÷ñ«2“ÊWr6WBÑ%Tî≈EÙîÂı5UU%ÙD‘îÂÙT‘î≈2Ê÷ÇÜV÷ñ¬í”‚Ê˜&÷∆ó¶TV÷ñ¬ÜV÷ñ¬ííì∞¢WFFTF÷ñ‰6ˆÁG&ˆ«2Çì∞¢7V'67&ñ&U˜4Fˆ7V÷VÁG2Çì∞¢&VÊFW$6ˆ÷÷W76T÷ÊvV÷VÁD∆ó7BÇì∞¢&VÊFW$F÷ñÂW6W'2Çì∞¢ñbÜ7W'&VÁEW6W"í∞¢7V'67&ñ&UW6W'2Çì∞¢7V'67&ñ&T˜W&F˜%˜6óFñˆÁ2Çì∞¢7V'67&ñ&U&ˆw&÷÷¶ñˆÊíÇì∞¢–¢“ì∞ß–†¶gVÊ7Fñˆ‚7F˜F÷ñÂW6W'57V'67&óFñˆ‚Çí∞¢ñbáVÁ7V'67&ñ&TF÷ñÂW6W'2í∞¢VÁ7V'67&ñ&TF÷ñÂW6W'2Çì∞¢VÁ7V'67&ñ&TF÷ñÂW6W'2“ÁV∆√∞¢–¢F÷ñ‰V÷ñ«2“ÊWr6WBÑ%Tî≈EÙîÂı5UU%ÙD‘îÂÙT‘î≈2Ê÷ÇÜV÷ñ¬í”‚Ê˜&÷∆ó¶TV÷ñ¬ÜV÷ñ¬ííì∞¢&VÊFW$F÷ñÂW6W'2Çì∞¢&VÊFW%W6W$&‰∆ó7BÇì∞ß–†¶gVÊ7Fñˆ‚7V'67&ñ&UW6W'2Çí∞¢7F˜W6W'57V'67&óFñˆ‚Çì∞¢ñbÇ7W'&VÁEW6W"í&WGW&„∞¢6ˆÁ7B6˜W&6R“6‰÷ÊvTFFÇê¢ÚF"Ê6ˆ∆∆V7Fñˆ‚Ç'∆Ff˜&’W6W'2"ê¢¢F"Ê6ˆ∆∆V7Fñˆ‚Ç'∆Ff˜&’W6W'2"íÁvÜW&RÜfó&V&6RÊfó&W7F˜&R‰fñV∆EFÇÊFˆ7V÷VÁDñBÇí¬#”“"¬7W'&VÁEW6W"ÁVñBì∞¢6ˆÁ7B«ï6Ê6Ü˜B“á6Ê6Ü˜Bí”‚∞¢∆Ff˜&’W6W'2“6Ê6Ü˜BÊFˆ72Ê÷ÇÜFˆ2í”‚∞¢6ˆÁ7BFF“Fˆ2ÊFFÇí«¬∑”∞¢&WGW&‚∞¢ñC¢Fˆ2ÊñB¿¢Fó7∆îÊ÷S¢7G&ñÊrÜFFÊFó7∆îÊ÷R«¬FFÊV÷ñ¬«¬%WFVÁFR"í¿¢V÷ñ√¢7G&ñÊrÜFFÊV÷ñ¬«¬""í¿¢&ˆ∆S¢7G&ñÊrÜFFÁ&ˆ∆R«¬""í¿¢‚‚ÊFF¢”∞¢“íÁ6˜'BÇÜ¬"í”‚7G&ñÊrÜÊFó7∆îÊ÷R«¬""íÊ∆ˆ6∆T6ˆ◊&RÖ7G&ñÊrÜ"ÊFó7∆îÊ÷R«¬""í¬&óB"íì∞¢vñÊF˜r‰ÜW&66W74&˜f√ÚÁ&VÊFW$F÷ñ‚á∆Ff˜&’W6W'2¬6‰÷ÊvTFFÇíì∞¢7ñÊ4Ê˜Fñfñ6Fñˆ‰WFı&VfW&VÊ6Tg&ˆ’&ˆfñ∆RÇì∞¢÷ñ&TWFÙVÊ&∆TÊ˜Fñfñ6FñˆÁ2Çì∞¢FVÊñVDñ◊ñÁFÙ7FñˆÁ2“vWDFVÊñVD7FñˆÁ4f˜$7W'&VÁEW6W"Çì∞¢&VÊFW$6ÜE&V6óñVÁG2Çì∞¢6ˆÁ7B7W'&VÁE&ˆfñ∆R“∆Ff˜&’W6W'2ÊfñÊBÇáW6W"í”‚7G&ñÊráW6W"ÊñB«¬W6W"ÁVñB«¬""í””“7G&ñÊrÜ7W'&VÁEW6W#ÚÁVñB«¬""íì∞¢6ˆÁ7B7W'&VÁD66˜VÁE7FGW2“vñÊF˜r‰ÜW&66W74&˜f√ÚÁ7FGW4ˆbÜ7W'&VÁE&ˆfñ∆Rì∞¢ñbÜ7W'&VÁE&ˆfñ∆Rbb7W'&VÁD66˜VÁE7FGW2”“&GFófÚ"bb6‰÷ÊvTFFÇíí∞¢fˆñB6∆V%W'6ó7FVE6W76ñˆ‚Çì∞¢7F˜6ˆ÷÷W76U7V'67&óFñˆ‚Çì≤7F˜ñ◊ñÁFï7V'67&óFñˆ‚Çì≤7F˜7VG&U7V'67&óFñˆ‚Çì≤7F˜W'6ˆÊ∆U7V'67&óFñˆ‚Çì≤7F˜÷Wß¶ï7V'67&óFñˆ‚Çì≤7F˜v∆ˆ&ƒÊ˜Fñfñ6FñˆÁ57V'67&óFñˆ‚Çì∞¢fˆñBvñÊF˜r‰ÜW&66W74&˜f√ÚÁfW&ñgíÜ7W'&VÁEW6W"ì∞¢&WGW&„∞¢–¢ñbÜ7W'&VÁE&ˆfñ∆SÚÊ&ÊÊVBbb6‰÷ÊvTFFÇíí∞¢7W'&VÁEW6W$&Â&ˆfñ∆R“7W'&VÁE&ˆfñ∆S∞¢7F˜6ˆ÷÷W76U7V'67&óFñˆ‚Çì≤7F˜ñ◊ñÁFï7V'67&óFñˆ‚Çì≤7F˜7VG&U7V'67&óFñˆ‚Çì≤7F˜W'6ˆÊ∆U7V'67&óFñˆ‚Çì≤7F˜÷Wß¶ï7V'67&óFñˆ‚Çì≤7F˜v∆ˆ&ƒÊ˜Fñfñ6FñˆÁ57V'67&óFñˆ‚Çì∞¢6WDWFÜVÁFñ6Fñˆ‰vFU7FFRÇ&&ÊÊVB"ì∞¢&WGW&„∞¢–¢&VÊFW%W6W%W&÷ó76ñˆ‰∆ó7BÇì∞¢&VÊFW%W6W$&‰∆ó7BÇì∞¢&VÊFW$Ê˜Fñfñ6FñˆÂF&vWEW6W'2Çì∞¢&VÊFW$ÜVFW$7FófóGï7V÷÷'íÇì∞¢&VÊFW$WáFW&Êƒ2Çì∞¢&VÊFW$ñ◊ñÁFíÇì∞¢&VÊFW$÷Çì∞¢&VÊFW%FˆFï7V÷÷'íÇì∞¢ñbÜ7W'&VÁEW6W"bb7W'&VÁE&ˆfñ∆Rí∞¢fˆñB6fUW'6ó7FVE6W76ñˆ‚Ü7W'&VÁEW6W"¬∞¢‚‚Ê7W'&VÁE&ˆfñ∆R¿¢ó4F÷ñ„¢6‰÷ÊvTFFÇí¿¢&ˆ∆S¢7W'&VÁE&ˆfñ∆RÁ&ˆ∆R«¬7W'&VÁE&ˆfñ∆RÁ'Vˆ∆Ú«¬Ü6‰÷ÊvTFFÇíÚ&F÷ñ‚"¢'W6W""ê¢“ì∞¢–¢6ÜV6¥ÊE6VÊDÜ˜W'4FVF∆ñÊT∆W'G2Çì∞¢”∞¢ÚÚˆÂ6Ê6Ü˜BñÊ6«VFRvú:ñ¬6&ñ6÷VÁFÚñÊó¶ñ∆S¢ñ¬&V6VFVÁFRvWBÇíf6Wf¢ÚÚv&RGVRfˆ«FR∆7FW76∆ó7FˆvÊí66W76Ú‡¢G'í∞¢VÁ7V'67&ñ&UW6W'2“6˜W&6RÊˆÂ6Ê6Ü˜BÜ«ï6Ê6Ü˜B¬ÜW'&˜"í”‚∞¢∆ˆtfó&W7F˜&TW'&˜"Ç$ƒÙBUDTÂDíƒï5DT‰U""¬W'&˜"ì∞¢∆Ff˜&’W6W'2“µ”∞¢&VÊFW$6ÜE&V6óñVÁG2Çì∞¢&VÊFW$ÜVFW$7FófóGï7V÷÷'íÇì∞¢“ì∞¢“6F6ÇÜW'&˜"í∞¢∆ˆtfó&W7F˜&TW'&˜"Ç$ƒÙBUDTÂDíƒï5DT‰U"î‰ïB"¬W'&˜"ì∞¢–ß–†¶gVÊ7Fñˆ‚7V'67&ñ&U&ˆw&÷÷¶ñˆÊíá≤f˜&6R“f«6R““∑“í∞¢ñbÇ7W'&VÁEW6W"«¬6‰÷ÊvTFFÇíí&WGW&‚&ˆ÷ó6RÁ&W6ˆ«fRÇì∞¢ñbÇf˜&6Rbb&ˆw&÷÷¶ñˆÊî∆ˆFVDBbbFFRÊÊ˜rÇí“&ˆw&÷÷¶ñˆÊî∆ˆFVDB¬$Ùu$‘‘§îÙ‰ïÙ44ÑUıED≈Ù’2í∞¢&VÊFW%&ˆw&÷÷¶ñˆÊíÇì∞¢&WGW&‚&ˆ÷ó6RÁ&W6ˆ«fRÇì∞¢–¢ñbá&ˆw&÷÷¶ñˆÊî∆ˆE&ˆ÷ó6Rí&WGW&‚&ˆw&÷÷¶ñˆÊî∆ˆE&ˆ÷ó6S∞¢ñbáVÁ7V'67&ñ&U&ˆw&÷÷¶ñˆÊíí∞¢VÁ7V'67&ñ&U&ˆw&÷÷¶ñˆÊíÇì∞¢VÁ7V'67&ñ&U&ˆw&÷÷¶ñˆÊí“ÁV∆√∞¢–¢&ˆw&÷÷¶ñˆÊî∆ˆE&ˆ÷ó6R“F"Ê6ˆ∆∆V7Fñˆ‚Ç'&ˆw&÷÷¶ñˆÊí"íÊvWBÇê¢ÁFÜV‚Çá6Ê6Ü˜Bí”‚∞¢&ˆw&÷÷¶ñˆÊí“6Ê6Ü˜BÊFˆ72Ê÷ÇÜFˆ2í”‚á≤ñC¢Fˆ2ÊñB¬‚‚ÊFˆ2ÊFFÇí“íì∞¢&ˆw&÷÷¶ñˆÊî∆ˆFVDB“FFRÊÊ˜rÇì∞¢&VÊFW%&ˆw&÷÷¶ñˆÊíÇì∞¢“ê¢Ê6F6ÇÇÜW'&˜"í”‚∞¢6ˆÁ6ˆ∆RÊW'&˜"Ç$W'&˜&R6&ñ6÷VÁFÚ&ˆw&÷÷¶ñˆÊì¢"¬W'&˜"ì∞¢&VÊFW%&ˆw&÷÷¶ñˆÊíÇì∞¢“ê¢ÊfñÊ∆«íÇÇí”‚≤&ˆw&÷÷¶ñˆÊî∆ˆE&ˆ÷ó6R“ÁV∆√≤“ì∞¢&WGW&‚&ˆw&÷÷¶ñˆÊî∆ˆE&ˆ÷ó6S∞ß–†¶gVÊ7Fñˆ‚˜V∆FU&ˆw&÷÷¶ñˆÊTf˜&‘˜FñˆÁ2Çí∞¢ñbáVíÁ&ˆw&÷÷6ˆ÷÷W76í∞¢6ˆÁ7B&Wfñ˜W2“7G&ñÊráVíÁ&ˆw&÷÷6ˆ÷÷W76Áf«VR«¬""ì∞¢6ˆÁ7B6ˆ÷÷W76R“6˜'D6ˆ÷÷W76T'î7&VFVDDFW62Ñ'&íÊg&ˆ“Ü6ˆ÷÷W76T'îñBÁf«VW2Çííì∞¢VíÁ&ˆw&÷÷6ˆ÷÷W76ÊñÊÊW$ÖD‘¬“#∆˜Fñˆ‚f«VS“rs‰6ˆ÷÷W76¬ˆ˜Fñˆ„‚#∞¢6ˆ÷÷W76RÊf˜$V6ÇÇÜ6ˆ÷÷W76í”‚∞¢6ˆÁ7B˜Fñˆ‚“Fˆ7V÷VÁBÊ7&VFTV∆V÷VÁBÇ&˜Fñˆ‚"ì∞¢˜Fñˆ‚Áf«VR“7G&ñÊrÜ6ˆ÷÷W76ÊÊˆ÷R«¬""íÁG&ñ“Çì∞¢˜Fñˆ‚ÁFWáD6ˆÁFVÁB“7G&ñÊrÜ6ˆ÷÷W76ÊÊˆ÷R«¬$6ˆ÷÷W76"ì∞¢VíÁ&ˆw&÷÷6ˆ÷÷W76ÊVÊD6Üñ∆BÜ˜Fñˆ‚ì∞¢“ì∞¢ñbá&Wfñ˜W2íVíÁ&ˆw&÷÷6ˆ÷÷W76Áf«VR“&Wfñ˜W3∞¢–¢ñbáVíÁ&ˆw&÷÷˜W&F˜&íí∞¢6ˆÁ7B6V∆V7FVB“ÊWr6WBÑ'&íÊg&ˆ“áVíÁ&ˆw&÷÷˜W&F˜&íÁ6V∆V7FVD˜FñˆÁ2«¬µ“íÊ÷ÇÜ˜Bí”‚7G&ñÊrÜ˜BÁf«VR«¬""íÁG&ñ“Çííì∞¢VíÁ&ˆw&÷÷˜W&F˜&íÊñÊÊW$ÖD‘¬“"#∞¢W'6ˆÊ∆U&V6˜&G0¢Ê÷ÇáW'6ˆ‚í”‚vWEW'6ˆÊ∆TFó7∆îÊ÷RáW'6ˆ‚íê¢Êfñ«FW"Ñ&ˆˆ∆V‚ê¢Á6˜'BÇÜ¬"í”‚7G&ñÊrÜíÊ∆ˆ6∆T6ˆ◊&RÖ7G&ñÊrÜ"í¬&óB"íê¢Êf˜$V6ÇÇÜ˜W&F˜$Ê÷Rí”‚∞¢6ˆÁ7Bf«VR“7G&ñÊrÜ˜W&F˜$Ê÷R«¬""íÁG&ñ“Çì∞¢ñbÇf«VRí&WGW&„∞¢6ˆÁ7B˜Fñˆ‚“Fˆ7V÷VÁBÊ7&VFTV∆V÷VÁBÇ&˜Fñˆ‚"ì∞¢˜Fñˆ‚Áf«VR“f«VS∞¢˜Fñˆ‚ÁFWáD6ˆÁFVÁB“f«VS∞¢ñbá6V∆V7FVBÊÜ2áf«VRíí˜Fñˆ‚Á6V∆V7FVB“G'VS∞¢VíÁ&ˆw&÷÷˜W&F˜&íÊVÊD6Üñ∆BÜ˜Fñˆ‚ì∞¢“ì∞¢–¢ñbáVíÁ&ˆw&÷÷÷Wß¶íí∞¢6ˆÁ7B6V∆V7FVB“ÊWr6WBÑ'&íÊg&ˆ“áVíÁ&ˆw&÷÷÷Wß¶íÁ6V∆V7FVD˜FñˆÁ2«¬µ“íÊ÷ÇÜ˜Bí”‚7G&ñÊrÜ˜BÁf«VR«¬""ííì∞¢VíÁ&ˆw&÷÷÷Wß¶íÊñÊÊW$ÖD‘¬“"#∞¢÷Wß¶ï&V6˜&G2Êf˜$V6ÇÇÜ÷Wß¶Úí”‚∞¢6ˆÁ7B∆&V¬“7G&ñÊrÜ÷Wß¶ÚÊ‰ñB«¬÷Wß¶ÚÊÊˆ÷R«¬÷Wß¶ÚÊ÷ˆFV∆∆Ú«¬""íÁG&ñ“Çì∞¢ñbÇ∆&V¬í&WGW&„∞¢6ˆÁ7B˜Fñˆ‚“Fˆ7V÷VÁBÊ7&VFTV∆V÷VÁBÇ&˜Fñˆ‚"ì∞¢˜Fñˆ‚Áf«VR“∆&V√∞¢˜Fñˆ‚ÁFWáD6ˆÁFVÁB“∆&V√∞¢ñbá6V∆V7FVBÊÜ2Ü∆&V¬íí˜Fñˆ‚Á6V∆V7FVB“G'VS∞¢VíÁ&ˆw&÷÷÷Wß¶íÊVÊD6Üñ∆BÜ˜Fñˆ‚ì∞¢“ì∞¢–ß–†¶gVÊ7Fñˆ‚vWE&ˆw&÷÷¶ñˆÊT˜W&F˜$˜FñˆÁ2Çí∞¢&WGW&‚W'6ˆÊ∆U&V6˜&G0¢Ê÷ÇáW'6ˆ‚í”‚á≤f«VS¢vWEW'6ˆÊ∆TFó7∆îÊ÷RáW'6ˆ‚í¬fF#¢7G&ñÊráW'6ˆ‚ÊfF%W&¬«¬""íÁG&ñ“Çí“íê¢Êfñ«FW"ÇÜóFV“í”‚óFV“Áf«VRê¢Á6˜'BÇÜ¬"í”‚7G&ñÊrÜÁf«VRíÊ∆ˆ6∆T6ˆ◊&RÖ7G&ñÊrÜ"Áf«VRí¬&óB"íì∞ß–†¶gVÊ7Fñˆ‚vWE&ˆw&÷÷¶ñˆÊT÷Wß¶î˜FñˆÁ2Çí∞¢&WGW&‚÷Wß¶ï&V6˜&G0¢Ê÷ÇÜ÷Wß¶Úí”‚7G&ñÊrÜ÷Wß¶ÚÊ‰ñB«¬÷Wß¶ÚÊÊˆ÷R«¬÷Wß¶ÚÊ÷ˆFV∆∆Ú«¬""íÁG&ñ“Çíê¢Êfñ«FW"Ñ&ˆˆ∆V‚ê¢Á6˜'BÇÜ¬"í”‚Ê∆ˆ6∆T6ˆ◊&RÜ"¬&óB"íê¢Ê÷Çáf«VRí”‚á≤f«VR¬fF#¢""“íì∞ß–†¶gVÊ7Fñˆ‚'Vñ∆E&ˆw&÷÷¶ñˆÊTWFˆ6ˆ◊∆WFRá&ˆ˜B¬∆&V¬¬˜FñˆÁ2¬6V∆V7FVEf«VW2“µ“í∞¢ñbÇ&ˆ˜Bí&WGW&‚≤vWEf«VW3¢Çí”‚µ“”∞¢6ˆÁ7B6V∆V7FVB“ÊWr÷Çì∞¢6V∆V7FVEf«VW2Êf˜$V6ÇÇábí”‚6V∆V7FVBÁ6WBÖ7G&ñÊrábíÁFÙ∆˜vW$66RÇí¬≤f«VS¢b¬fF#¢""“íì∞¢&ˆ˜BÊñÊÊW$ÖD‘¬“∆∆&V√‚G∂∆&V«”¬ˆ∆&V√„∆ñÁWBGóS“'FWáB"∆6VÜˆ∆FW#“$6W&6‚‚‚"WFˆ6ˆ◊∆WFS“&ˆfb#„∆Fób6∆73“&WFˆ6ˆ◊∆WFR÷∆ó7BÜñFFV‚#„¬ˆFóc„∆Fób6∆73“&WFˆ6ˆ◊∆WFR÷6Üó2#„¬ˆFócÊ∞¢6ˆÁ7BñÁWB“&ˆ˜BÁVW'ï6V∆V7F˜"Ç&ñÁWB"ì∞¢6ˆÁ7B∆ó7B“&ˆ˜BÁVW'ï6V∆V7F˜"Ç"ÊWFˆ6ˆ◊∆WFR÷∆ó7B"ì∞¢6ˆÁ7B6Üó2“&ˆ˜BÁVW'ï6V∆V7F˜"Ç"ÊWFˆ6ˆ◊∆WFR÷6Üó2"ì∞¢gVÊ7Fñˆ‚&VÊFW$6Üó2Çí∞¢6Üó2ÊñÊÊW$ÖD‘¬“'&íÊg&ˆ“á6V∆V7FVBÁf«VW2ÇííÊ÷ÇÜóFV“í”‚«7‚6∆73“&WFˆ6ˆ◊∆WFR÷6Üó#‚G∂óFV“ÊfF"Ú∆ñ÷r7&3“"G∂W66TÖD‘¬ÜóFV“ÊfF"ó“"«C“""vñGFÉ“#Ç"ÜVñváC“#Ç#Ê¢"'“G∂W66TÖD‘¬ÜóFV“Áf«VRó”∆'WGFˆ‚GóS“&'WGFˆ‚"FF◊&V÷˜fS“"G∂W66TÖD‘¬ÜóFV“Áf«VRó“#Ó)…S¬ˆ'WGFˆ„„¬˜7„ÊíÊ¶ˆñ‚Ç""ì∞¢–¢gVÊ7Fñˆ‚&VÊFW$∆ó7BÇí∞¢6ˆÁ7B“7G&ñÊrÜñÁWBÁf«VR«¬""íÁG&ñ“ÇíÁFÙ∆˜vW$66RÇì∞¢6ˆÁ7Bfñ«FW&VB“˜FñˆÁ2Êfñ«FW"ÇÜóFV“í”‚óFV“Áf«VRÁFÙ∆˜vW$66RÇíÊñÊ6«VFW2áíbb6V∆V7FVBÊÜ2ÜóFV“Áf«VRÁFÙ∆˜vW$66RÇíííÁ6∆ñ6RÉ¬Çì∞¢∆ó7BÊñÊÊW$ÖD‘¬“fñ«FW&VBÊ÷ÇÜóFV“í”‚∆Fób6∆73“&WFˆ6ˆ◊∆WFR÷óFV“"FF◊f«VS“"G∂W66TÖD‘¬ÜóFV“Áf«VRó“#‚G∂W66TÖD‘¬ÜóFV“Áf«VRó”¬ˆFócÊíÊ¶ˆñ‚Ç""í«¬#∆Fób6∆73“vWFˆ6ˆ◊∆WFR÷óFV“◊WFVBs‰ÊW77V‚&ó7V«FFÛ¬ˆFóc‚#∞¢∆ó7BÊ6∆74∆ó7BÁ&V÷˜fRÇ&ÜñFFV‚"ì∞¢–¢ñÁWBÊFDWfVÁD∆ó7FVÊW"Ç&ñÁWB"¬&VÊFW$∆ó7Bì∞¢ñÁWBÊFDWfVÁD∆ó7FVÊW"Ç&fˆ7W2"¬&VÊFW$∆ó7Bì∞¢Fˆ7V÷VÁBÊFDWfVÁD∆ó7FVÊW"Ç&6∆ñ6≤"¬ÜWfVÁBí”‚≤ñbÇ&ˆ˜BÊ6ˆÁFñÁ2ÜWfVÁBÁF&vWBíí∆ó7BÊ6∆74∆ó7BÊFBÇ&ÜñFFV‚"ì≤“ì∞¢∆ó7BÊFDWfVÁD∆ó7FVÊW"Ç&6∆ñ6≤"¬ÜWfVÁBí”‚∞¢6ˆÁ7B&˜r“WfVÁBÁF&vWBÊ6∆˜6W7BÇ%∂FF◊f«VU“"ì∞¢ñbÇ&˜rí&WGW&„∞¢6ˆÁ7Bf«VR“7G&ñÊrá&˜rÊvWDGG&ñ'WFRÇ&FF◊f«VR"í«¬""ì∞¢6ˆÁ7Bf˜VÊB“˜FñˆÁ2ÊfñÊBÇÜóFV“í”‚óFV“Áf«VR””“f«VRí«¬≤f«VR¬fF#¢""”∞¢6V∆V7FVBÁ6WBáf«VRÁFÙ∆˜vW$66RÇí¬f˜VÊBì∞¢ñÁWBÁf«VR“"#∞¢&VÊFW$6Üó2Çì∞¢&VÊFW$∆ó7BÇì∞¢“ì∞¢6Üó2ÊFDWfVÁD∆ó7FVÊW"Ç&6∆ñ6≤"¬ÜWfVÁBí”‚∞¢6ˆÁ7B'F‚“WfVÁBÁF&vWBÊ6∆˜6W7BÇ%∂FF◊&V÷˜fU“"ì∞¢ñbÇ'F‚í&WGW&„∞¢6V∆V7FVBÊFV∆WFRÖ7G&ñÊrÜ'F‚ÊvWDGG&ñ'WFRÇ&FF◊&V÷˜fR"í«¬""íÁFÙ∆˜vW$66RÇíì∞¢&VÊFW$6Üó2Çì∞¢“ì∞¢&VÊFW$6Üó2Çì∞¢&WGW&‚≤vWEf«VW3¢Çí”‚'&íÊg&ˆ“á6V∆V7FVBÁf«VW2ÇííÊ÷ÇÜóFV“í”‚óFV“Áf«VRí”∞ß–†¶gVÊ7Fñˆ‚7V'67&ñ&T˜W&F˜%˜6óFñˆÁ2Çí∞¢7F˜˜W&F˜%˜6óFñˆÁ57V'67&óFñˆ‚Çì∞¢ÚÚ∆÷6˜'&VÁFR÷˜7G&6ˆ∆Ú∆˜6ó¶ñˆÊR∆ˆ6∆RFV∆¬wWFVÁFR‚ñ¬fV66Üñ¢ÚÚ∆ó7FVÊW"∆VvvWfGWGFR∆R˜6ó¶ñˆÊí6VÁ¶6ÜRfVÊó76W&Ú÷ífó7V∆óß¶FR‡¢˜W&F˜%˜6óFñˆÁ2“µ”∞¢&VÊFW$÷Çì∞ß–†¶gVÊ7Fñˆ‚7F˜˜W&F˜%˜6óFñˆÁ57V'67&óFñˆ‚Çí∞¢ñbáVÁ7V'67&ñ&T˜W&F˜%˜6óFñˆÁ2í∞¢VÁ7V'67&ñ&T˜W&F˜%˜6óFñˆÁ2Çì∞¢VÁ7V'67&ñ&T˜W&F˜%˜6óFñˆÁ2“ÁV∆√∞¢–¢˜W&F˜%˜6óFñˆÁ2“µ”∞ß–†¶gVÊ7Fñˆ‚7F˜W6W'57V'67&óFñˆ‚Çí∞¢ñbáVÁ7V'67&ñ&UW6W'2í∞¢VÁ7V'67&ñ&UW6W'2Çì∞¢VÁ7V'67&ñ&UW6W'2“ÁV∆√∞¢–¢∆Ff˜&’W6W'2“µ”∞¢FVÊñVDñ◊ñÁFÙ7FñˆÁ2“ÊWr6WBÇì∞¢&VÊFW$6ÜE&V6óñVÁG2Çì∞¢&VÊFW%W6W%W&÷ó76ñˆ‰∆ó7BÇì∞¢&VÊFW$Ê˜Fñfñ6FñˆÂF&vWEW6W'2Çì∞¢&VÊFW$ÜVFW$7FófóGï7V÷÷'íÇì∞¢&VÊFW$WáFW&Êƒ2Çì∞ß–†¶gVÊ7Fñˆ‚6‰&˜fTÜ˜W'4∆WfV√á&WVW7Bí∞¢ñbÇ7W'&VÁEW6W"«¬&WVW7Bí&WGW&‚f«6S∞¢ñbÜ6‰÷ÊvTFFÇíí&WGW&‚G'VS∞¢&WGW&‚7G&ñÊrá&WVW7BÊ7&VFVD'ïVñB«¬""í””“7G&ñÊrÜ7W'&VÁEW6W"ÁVñB«¬""ì∞ß–†¶gVÊ7Fñˆ‚Ü5f∆ñDÜ˜W'5&˜w2á&V6˜&B“∑“í∞¢&WGW&‚Ñ'&íÊó4'&íá&V6˜&BÊVÁG&ñW2íÚ&V6˜&BÊVÁG&ñW2¢µ“íÁ6ˆ÷RÇÜVÁG'íí”‚Ä¢7G&ñÊrÜVÁG'ìÚÊ6ˆ÷÷W76ñB«¬""íÁG&ñ“Çê¢bbÑ'&íÊó4'&íÜVÁG'ìÚÁ&˜w2íÚVÁG'íÁ&˜w2¢µ“íÁ6ˆ÷RÇá&˜rí”‚Ä¢7G&ñÊrá&˜sÚÊ˜W&F˜&R«¬""íÁG&ñ“ÇíbbÁV÷&W"á&˜sÚÊ˜&R«¬í‚ ¢íê¢íì∞ß–†¶gVÊ7Fñˆ‚ó4&∆ˆ6∂ñÊtÜ˜W'4&˜f≈vóFÜ˜WEfó6ñ&∆U&V6˜&Bá&WVW7B“∑“í∞¢6ˆÁ7B7FGW2“7G&ñÊrá&WVW7BÁ7FGW2«¬""íÁG&ñ“Çì∞¢ñbá7FGW2””“'&V¶V7FVB"«¬7FGW2””“&&˜fVB"í&WGW&‚f«6S∞¢&WGW&‚Ü5f∆ñDÜ˜W'5&˜w2á&WVW7Bì∞ß–†¶7ñÊ2gVÊ7Fñˆ‚VÊ&∆ˆ6¥ñÁf∆ñDÜ˜W'5&WVW7Bá&WVW7Bí∞¢ñbÇ6‰÷ÊvTFFÇíí∞¢∆W'BÇ%6ˆ∆Ú¬vF÷ñ‚\;"6&∆ˆ66&RVÊ&ñ6ÜñW7F˜&RñÊ6ˆ◊∆WF‚"ì∞¢&WGW&„∞¢–¢ñbÇ&WVW7CÚÊñBí&WGW&„∞¢6ˆÁ7Bˆ≤“vñÊF˜rÊ6ˆÊfó&“Ü6&∆ˆ66&R∆&ñ6ÜñW7F˜&RG∑&WVW7BÊñG”ÚfW',:÷&6F6ˆ÷R&ñfóWFFW&6å:íÊˆ‚6ˆÁFñVÊRV‚&V6˜&B˜&R6ˆ◊∆WFÚÊì∞¢ñbÇˆ≤í&WGW&„∞¢vóBF"Ê6ˆ∆∆V7Fñˆ‚ÜvWD˜&T&˜f≈&WVW7G46ˆ∆∆V7Fñˆ‰Ê÷RÇííÊFˆ2á&WVW7BÊñBíÁ6WBá∞¢7FGW3¢'&V¶V7FVB"¿¢&V¶V7FVD'ì¢7W'&VÁEW6W"ÊV÷ñ¬«¬7W'&VÁEW6W"ÊFó7∆îÊ÷R«¬&F÷ñ‚"¿¢&V¶V7FVDC¢fó&V&6RÊfó&W7F˜&R‰fñV∆Ef«VRÁ6W'fW%Fñ÷W7F◊Çí¿¢&V¶V7FñˆÂ&V6ˆ„¢%6&∆ˆ66ÚF÷ñ„¢&∆ˆ66Ú˜&R6VÁ¶&V6˜&B˜&R6ˆ◊∆WFÚfó6ñ&ñ∆R‚"¿¢WFFVDC¢fó&V&6RÊfó&W7F˜&R‰fñV∆Ef«VRÁ6W'fW%Fñ÷W7F◊Çê¢“¬≤÷W&vS¢G'VR“ì∞¢vóBWFFTÜ˜W'4∆ˆ6∑4f˜$VÁG&ñW2á&WVW7BÊFFR«¬""¬&WVW7BÊVÁG&ñW2«¬µ“¬∞¢7FGW3¢'&V¶V7FVB"¿¢&˜f≈&WVW7DñC¢&WVW7BÊñB«¬""¿¢&V¶V7FVD'ì¢7W'&VÁEW6W"ÊV÷ñ¬«¬7W'&VÁEW6W"ÊFó7∆îÊ÷R«¬&F÷ñ‚ ¢“ì∞ß–†¶gVÊ7Fñˆ‚&VÊFW$Ü˜W'4&˜f≈&WVW7G2Çí∞¢ñbÇVíÊÜ˜W'4&˜f«4∆ó7B«¬VíÊÜ˜W'4&˜f«4fVVF&6≤í&WGW&„∞¢ñbÇ7W'&VÁEW6W"í∞¢VíÊÜ˜W'4&˜f«4fVVF&6≤ÁFWáD6ˆÁFVÁB“$fí∆ˆvñ‚W"fVFW&R∆R&ñ6ÜñW7FR˜&R‚#∞¢VíÊÜ˜W'4&˜f«4∆ó7BÊñÊÊW$ÖD‘¬“"#∞¢&WGW&„∞¢–¢6ˆÁ7Bfó6ñ&∆R“Ü˜W'4&˜f≈&WVW7G2Êfñ«FW"Çá&WVW7Bí”‚6‰÷ÊvTFFÇí«¬7G&ñÊrá&WVW7BÊ7&VFVD'ïVñB«¬""í””“7G&ñÊrÜ7W'&VÁEW6W"ÁVñB«¬""íì∞¢ñbÇfó6ñ&∆RÊ∆VÊwFÇí∞¢VíÊÜ˜W'4&˜f«4fVVF&6≤ÁFWáD6ˆÁFVÁB“$ÊW77VÊ&ñ6ÜñW7F˜&Rñ‚&˜f¶ñˆÊR‚#∞¢VíÊÜ˜W'4&˜f«4∆ó7BÊñÊÊW$ÖD‘¬“"#∞¢&WGW&„∞¢–¢VíÊÜ˜W'4&˜f«4fVVF&6≤ÁFWáD6ˆÁFVÁB“&ñ6ÜñW7FRG&˜fFS¢G∑fó6ñ&∆RÊ∆VÊwFá“Ê∞¢VíÊÜ˜W'4&˜f«4∆ó7BÊñÊÊW$ÖD‘¬“"#∞¢fó6ñ&∆RÊf˜$V6ÇÇá&WVW7Bí”‚∞¢6ˆÁ7B6&B“Fˆ7V÷VÁBÊ7&VFTV∆V÷VÁBÇ&'Fñ6∆R"ì∞¢6&BÊ6∆74Ê÷R“&óFV“÷6&B#∞¢6ˆÁ7BFFT∆&V¬“&WVW7BÊFFRÚÊWrFFRÜG∑&WVW7BÊFFW’C££íÁFÙ∆ˆ6∆TFFU7G&ñÊrÇ&óB‘ïB"í¢"“#∞¢6ˆÁ7B7FGW4÷“∞¢VÊFñÊuˆ∆WfV√¢$ñ‚GFW6&ñ÷ÚÙ≤"¿¢VÊFñÊuˆF÷ñ„¢$ñ‚GFW6Ù≤F÷ñ‚fñÊ∆R"¿¢&˜fVC¢&˜fF)»Rá&W˜'BG∑&WVW7BÊfñÊ∆ó¶VE&W˜'DñB«¬"“'“ñ¿¢&V¶V7FVC¢%&ñfóWFF)ÿ¬ ¢”∞¢6ˆÁ7B7FGW5FWáB“7FGW4÷∑&WVW7BÁ7FGW5“«¬&WVW7BÁ7FGW2«¬"“#∞¢6ˆÁ7BWFÜ˜"“&WVW7BÊ7&VFVD'îÊ÷R«¬&WVW7BÊ7&VFVD'îV÷ñ¬«¬$˜W&F˜&R#∞¢6ˆÁ7B7V÷÷'í“Ñ'&íÊó4'&íá&WVW7BÊVÁG&ñW2íÚ&WVW7BÊVÁG&ñW2¢µ“íÊ÷ÇÜVÁG'íí”‚∞¢6ˆÁ7BF˜B“ÜVÁG'íÁ&˜w2«¬µ“íÁ&VGV6RÇá7V“¬&˜rí”‚7V“≤ÑÁV÷&W"á&˜rÊ˜&R«¬í«¬í¬ì∞¢&WGW&‚∆∆ì‚G∂W66TÖD‘¬ÜVÁG'íÊ6ˆ÷÷W76Ê÷R«¬$6ˆ÷÷W76"ó”¢G∂W66TÖD‘¬Ö7G&ñÊráF˜Bíó÷É¬ˆ∆ìÊ∞¢“íÊ¶ˆñ‚Ç""ì∞¢6ˆÁ7BÜ4ñÁf∆ñD&∆ˆ6≤“ó4&∆ˆ6∂ñÊtÜ˜W'4&˜f≈vóFÜ˜WEfó6ñ&∆U&V6˜&Bá&WVW7Bì∞¢6&BÊñÊÊW$ÖD‘¬“ ¢«„∆#‰îC£¬ˆ#‚G∂W66TÖD‘¬á&WVW7BÊñB«¬"“"ó”¬˜‡¢«„∆#‰FF£¬ˆ#‚G∂W66TÖD‘¬ÜFFT∆&V¬ó“(
"∆#‰7&VFÚF£¬ˆ#‚G∂W66TÖD‘¬ÜWFÜ˜"ó”¬˜‡¢«„∆#Â7FFÛ£¬ˆ#‚G∂W66TÖD‘¬á7FGW5FWáBó”¬˜‡¢«V√‚G∑7V÷÷'í«¬#∆∆ì‰ÊW77VÊ6ˆ÷÷W76¬ˆ∆ì‚'”¬˜V√‡¢G∂Ü4ñÁf∆ñD&∆ˆ6≤Ú«6∆73“'v&ÊñÊr#„∆#‰W'&˜&S£¬ˆ#‚&ó7V«FV‚&∆ˆ66Ú˜&R÷ñ¬&V6˜&B˜&RÊˆ‚:Ç7FFÚG&˜fFÚ‚6ˆÁFGF&R÷÷ñÊó7G&F˜&R„¬˜Ê¢"'–¢G∑&WVW7BÁ&V¶V7FñˆÂ&V6ˆ‚Ú«„∆#‰÷˜FófÚ&ñfóWFÛ£¬ˆ#‚G∂W66TÖD‘¬á&WVW7BÁ&V¶V7FñˆÂ&V6ˆ‚ó”¬˜Ê¢"'–¢∞¢6ˆÁ7B7FñˆÁ2“Fˆ7V÷VÁBÊ7&VFTV∆V÷VÁBÇ&Fób"ì∞¢7FñˆÁ2Ê6∆74Ê÷R“&óFV“÷7FñˆÁ2#∞¢ñbÜÜ4ñÁf∆ñD&∆ˆ6≤bb6‰÷ÊvTFFÇíí∞¢7FñˆÁ2ÊVÊD6Üñ∆BÜ7&VFT'WGFˆ‚Ç%6&∆ˆ66&∆ˆ66Ú˜&R"¬Çí”‚VÊ&∆ˆ6¥ñÁf∆ñDÜ˜W'5&WVW7Bá&WVW7Bííì∞¢–¢ñbÇÜ4ñÁf∆ñD&∆ˆ6≤bb&WVW7BÁ7FGW2””“'VÊFñÊuˆ∆WfV√"bb6‰&˜fTÜ˜W'4∆WfV√á&WVW7Bíí∞¢7FñˆÁ2ÊVÊD6Üñ∆BÜ7&VFT'WGFˆ‚Ç$66WGF∆ófV∆∆Ú"¬Çí”‚&˜fTÜ˜W'5&WVW7D∆WfV√á&WVW7Bííì∞¢7FñˆÁ2ÊVÊD6Üñ∆BÜ7&VFT'WGFˆ‚Ç%&ñfóWF"¬Çí”‚&V¶V7DÜ˜W'5&WVW7Bá&WVW7Bííì∞¢–¢ñbÇÜ4ñÁf∆ñD&∆ˆ6≤bb&WVW7BÁ7FGW2””“'VÊFñÊuˆF÷ñ‚"bb6‰÷ÊvTFFÇíí∞¢7FñˆÁ2ÊVÊD6Üñ∆BÜ7&VFT'WGFˆ‚Ç$66WGFF÷ñ‚fñÊ∆R"¬Çí”‚&˜fTÜ˜W'5&WVW7D∆WfV√"á&WVW7Bííì∞¢7FñˆÁ2ÊVÊD6Üñ∆BÜ7&VFT'WGFˆ‚Ç%&ñfóWF"¬Çí”‚&V¶V7DÜ˜W'5&WVW7Bá&WVW7Bííì∞¢–¢ñbÜ7FñˆÁ2Ê6Üñ∆G&V‚Ê∆VÊwFÇí6&BÊVÊD6Üñ∆BÜ7FñˆÁ2ì∞¢VíÊÜ˜W'4&˜f«4∆ó7BÊVÊD6Üñ∆BÜ6&Bì∞¢“ì∞ß–†††¶7ñÊ2gVÊ7Fñˆ‚&˜fTÜ˜W'5&WVW7D∆WfV√á&WVW7Bí∞¢ñbÇ6‰&˜fTÜ˜W'4∆WfV√á&WVW7Bíí∞¢∆W'BÇ$Êˆ‚WF˜&óß¶FÚ¬&ñ÷Ú∆ófV∆∆ÚFí&˜f¶ñˆÊR‚"ì∞¢&WGW&„∞¢–¢vóBF"Ê6ˆ∆∆V7Fñˆ‚ÜvWD˜&T&˜f≈&WVW7G46ˆ∆∆V7Fñˆ‰Ê÷RÇííÊFˆ2á&WVW7BÊñBíÁ6WBá∞¢7FGW3¢'VÊFñÊuˆF÷ñ‚"¿¢∆WfV√&˜fVD'ì¢7W'&VÁEW6W"ÊV÷ñ¬«¬7W'&VÁEW6W"ÊFó7∆îÊ÷R«¬'WFVÁFR"¿¢∆WfV√&˜fVDC¢fó&V&6RÊfó&W7F˜&R‰fñV∆Ef«VRÁ6W'fW%Fñ÷W7F◊Çí¿¢WFFVDC¢fó&V&6RÊfó&W7F˜&R‰fñV∆Ef«VRÁ6W'fW%Fñ÷W7F◊Çê¢“¬≤÷W&vS¢G'VR“ì∞¢vóBWFFTÜ˜W'4∆ˆ6∑4f˜$VÁG&ñW2á&WVW7BÊFFR«¬""¬&WVW7BÊVÁG&ñW2«¬µ“¬∞¢7FGW3¢'VÊFñÊuˆF÷ñ‚"¿¢&˜f≈&WVW7DñC¢&WVW7BÊñB«¬" ¢“ì∞¢vóBÊ˜FñgîF÷ñÁ4f˜$fñÊƒÜ˜W'4&˜f¬á&WVW7BÊñB¬&WVW7B¬7W'&VÁEW6W"ÊFó7∆îÊ÷R«¬7W'&VÁEW6W"ÊV÷ñ¬«¬'WFVÁFR"ì∞¢6ˆÁ7B&WVW7FW"“∆Ff˜&’W6W'2ÊfñÊBÇáW6W"í”‚7G&ñÊráW6W"ÁVñB«¬""í””“7G&ñÊrá&WVW7BÊ7&VFVD'ïVñB«¬""íì∞¢ñbá&WVW7FW#ÚÊñBí∞¢vóB6VÊE&ófFT6ÜDÊ˜Fñfñ6Fñˆ‚á∞¢&V6óñVÁDñC¢&WVW7FW"ÊñB¿¢FWáC¢)»R&ñ6ÜñW7F˜&RG∑&WVW7BÊñG”¢&ñ÷Ú∆ófV∆∆Ú&˜fFÚ‚ñ‚GFW66ˆÊfW&÷F÷ñ‚fñÊ∆RÊ¿¢6VÊFW$Ê÷S¢7W'&VÁEW6W"ÊFó7∆îÊ÷R«¬7W'&VÁEW6W"ÊV÷ñ¬«¬%6ó7FV÷ ¢“ì∞¢–ß–†¶7ñÊ2gVÊ7Fñˆ‚&˜fTÜ˜W'5&WVW7D∆WfV√"á&WVW7Bí∞¢G'í∞¢6ˆÁ7B&W7V«B“vóB6fT&˜fVDÜ˜W'5&WVW7Bá&WVW7Bì∞¢6ˆÁ7BF&vWEW6W"“∆Ff˜&’W6W'2ÊfñÊBÇáW6W"í”‚7G&ñÊráW6W"ÁVñB«¬""í””“7G&ñÊrá&WVW7BÊ7&VFVD'ïVñB«¬""íì∞¢ñbáF&vWEW6W#ÚÊñBí∞¢vóB6VÊE&ófFT6ÜDÊ˜Fñfñ6Fñˆ‚á∞¢&V6óñVÁDñC¢F&vWEW6W"ÊñB¿¢FWáC¢)»R&ñ6ÜñW7F˜&RG∑&WVW7BÊñG“&˜fFFVfñÊóFóf÷VÁFR‚&W˜'B6«fFÛ¢G∑&W7V«BÁ&W˜'DñG“Ê¿¢6VÊFW$Ê÷S¢7W'&VÁEW6W"ÊFó7∆îÊ÷R«¬7W'&VÁEW6W"ÊV÷ñ¬«¬$F÷ñ‚ ¢“ì∞¢–¢∆ˆE6fVDÜ˜W'5&W˜'G2Çì∞¢“6F6ÇÜW'&˜"í∞¢ñbÜW'&˜#ÚÊ6ˆFR””“&Ü˜W'2÷GW∆ñ6FR÷∆ˆ6≤"«¬W'&˜#ÚÊ6ˆFR””“&Ü˜W'2÷GW∆ñ6FR÷G&gB"í∞¢∆W'BÜW'&˜"Ê÷W76vR«¬f˜&÷DÜ˜W'4GW∆ñ6FT÷W76vRÜW'&˜"Ê6ˆÊf∆ñ7G2¬≤F÷ñ„¢G'VR“íì∞¢&WGW&„∞¢–¢∆W'BÜW'&˜#ÚÊ÷W76vR«¬$W'&˜&RGW&ÁFR∆6ˆÊfW&÷˜&R‚"ì∞¢Fá&˜rW'&˜#∞¢–ß–†¶7ñÊ2gVÊ7Fñˆ‚&˜fTÜ˜W'5&WVW7Dg&ˆ‘6ÜBá&WVW7DñBí∞¢ñbÇ6‰÷ÊvTFFÇíí∞¢Fá&˜rÊWrW'&˜"Ç%6ˆ∆ÚF÷ñ‚\;"6ˆÊfW&÷&R∆R˜&RF6ÜB‚"ì∞¢–¢ñbÇ&WVW7DñBí∞¢Fá&˜rÊWrW'&˜"Ç$îB&ñ6ÜñW7FÊˆ‚f∆ñFÚ‚"ì∞¢–¢6ˆÁ7B&WVW7B“vóBvWDÜ˜W'4&˜f≈&WVW7D'îñBá&WVW7DñBì∞¢ñbÇ&WVW7Bí∞¢Fá&˜rÊWrW'&˜"Ç%&ñ6ÜñW7F˜&RÊˆ‚G&˜fF‚"ì∞¢–¢ñbÖ7G&ñÊrá&WVW7BÁ7FGW2«¬""í”“'VÊFñÊuˆF÷ñ‚"í∞¢Fá&˜rÊWrW'&˜"ÜñFV◊˜FVÁ¶¢&ñ6ÜñW7FÊˆ‚VÊFñÊuˆF÷ñ‚á7FFÛ¢Gµ7G&ñÊrá&WVW7BÁ7FGW2«¬'66ˆÊ˜66óWFÚ"ó“íÊì∞¢–¢vóB&˜fTÜ˜W'5&WVW7D∆WfV√"á&WVW7Bì∞ß–†¶7ñÊ2gVÊ7Fñˆ‚&V¶V7DÜ˜W'5&WVW7Dg&ˆ‘6ÜBá&WVW7DñBí∞¢ñbÇ6‰÷ÊvTFFÇíí∞¢Fá&˜rÊWrW'&˜"Ç%6ˆ∆ÚF÷ñ‚\;"&ñfóWF&R∆R˜&RF6ÜB‚"ì∞¢–¢ñbÇ&WVW7DñBí∞¢Fá&˜rÊWrW'&˜"Ç$îB&ñ6ÜñW7FÊˆ‚f∆ñFÚ‚"ì∞¢–¢6ˆÁ7B&WVW7B“vóBvWDÜ˜W'4&˜f≈&WVW7D'îñBá&WVW7DñBì∞¢ñbÇ&WVW7Bí∞¢Fá&˜rÊWrW'&˜"Ç%&ñ6ÜñW7F˜&RÊˆ‚G&˜fF‚"ì∞¢–¢ñbÖ7G&ñÊrá&WVW7BÁ7FGW2«¬""í”“'VÊFñÊuˆF÷ñ‚"í∞¢Fá&˜rÊWrW'&˜"ÜñFV◊˜FVÁ¶¢&ñ6ÜñW7FÊˆ‚VÊFñÊuˆF÷ñ‚á7FFÛ¢Gµ7G&ñÊrá&WVW7BÁ7FGW2«¬'66ˆÊ˜66óWFÚ"ó“íÊì∞¢–¢vóBF"Ê6ˆ∆∆V7Fñˆ‚ÜvWD˜&T&˜f≈&WVW7G46ˆ∆∆V7Fñˆ‰Ê÷RÇííÊFˆ2á&WVW7BÊñBíÁ6WBá∞¢7FGW3¢'&V¶V7FVB"¿¢&V¶V7FVD'ì¢7W'&VÁEW6W"ÊV÷ñ¬«¬7W'&VÁEW6W"ÊFó7∆îÊ÷R«¬&F÷ñ‚"¿¢&V¶V7FVDC¢fó&V&6RÊfó&W7F˜&R‰fñV∆Ef«VRÁ6W'fW%Fñ÷W7F◊Çí¿¢&V¶V7FñˆÂ&V6ˆ„¢%&ñfóWFFF6ÜBF÷ñ‚"¿¢WFFVDC¢fó&V&6RÊfó&W7F˜&R‰fñV∆Ef«VRÁ6W'fW%Fñ÷W7F◊Çê¢“¬≤÷W&vS¢G'VR“ì∞¢vóBWFFTÜ˜W'4∆ˆ6∑4f˜$VÁG&ñW2á&WVW7BÊFFR«¬""¬&WVW7BÊVÁG&ñW2«¬µ“¬∞¢7FGW3¢'&V¶V7FVB"¿¢&˜f≈&WVW7DñC¢&WVW7BÊñB«¬""¿¢&V¶V7FVD'ì¢7W'&VÁEW6W"ÊV÷ñ¬«¬7W'&VÁEW6W"ÊFó7∆îÊ÷R«¬&F÷ñ‚ ¢“ì∞¢6ˆÁ7BF&vWEW6W"“∆Ff˜&’W6W'2ÊfñÊBÇáW6W"í”‚7G&ñÊráW6W"ÁVñB«¬""í””“7G&ñÊrá&WVW7BÊ7&VFVD'ïVñB«¬""íì∞¢ñbáF&vWEW6W#ÚÊñBí∞¢vóB6VÊE&ófFT6ÜDÊ˜Fñfñ6Fñˆ‚á∞¢&V6óñVÁDñC¢F&vWEW6W"ÊñB¿¢FWáC¢)ÿ¬&ñ6ÜñW7F˜&RG∑&WVW7BÊñG“&ñfóWFFF6ÜBF÷ñ‚Ê¿¢6VÊFW$Ê÷S¢7W'&VÁEW6W"ÊFó7∆îÊ÷R«¬7W'&VÁEW6W"ÊV÷ñ¬«¬$F÷ñ‚ ¢“ì∞¢–ß–†¶7ñÊ2gVÊ7Fñˆ‚&V¶V7DÜ˜W'5&WVW7Bá&WVW7Bí∞¢6ˆÁ7B6Â&V¶V7B“á&WVW7BÁ7FGW2””“'VÊFñÊuˆ∆WfV√"bb6‰&˜fTÜ˜W'4∆WfV√á&WVW7Bíê¢«¬á&WVW7BÁ7FGW2””“'VÊFñÊuˆF÷ñ‚"bb6‰÷ÊvTFFÇíì∞¢ñbÇ6Â&V¶V7Bí∞¢∆W'BÇ$Êˆ‚WF˜&óß¶FÚ&ñfóWF&RVW7F&ñ6ÜñW7F‚"ì∞¢&WGW&„∞¢–¢6ˆÁ7B&V6ˆ‚“vñÊF˜rÁ&ˆ◊BÇ$÷˜FófÚFV¬&ñfóWFÚÜ˜¶ñˆÊ∆Rì¢"¬""í«¬"#∞¢vóBF"Ê6ˆ∆∆V7Fñˆ‚ÜvWD˜&T&˜f≈&WVW7G46ˆ∆∆V7Fñˆ‰Ê÷RÇííÊFˆ2á&WVW7BÊñBíÁ6WBá∞¢7FGW3¢'&V¶V7FVB"¿¢&V¶V7FVD'ì¢7W'&VÁEW6W"ÊV÷ñ¬«¬7W'&VÁEW6W"ÊFó7∆îÊ÷R«¬'WFVÁFR"¿¢&V¶V7FVDC¢fó&V&6RÊfó&W7F˜&R‰fñV∆Ef«VRÁ6W'fW%Fñ÷W7F◊Çí¿¢&V¶V7FñˆÂ&V6ˆ„¢7G&ñÊrá&V6ˆ‚íÁG&ñ“Çí¿¢WFFVDC¢fó&V&6RÊfó&W7F˜&R‰fñV∆Ef«VRÁ6W'fW%Fñ÷W7F◊Çê¢“¬≤÷W&vS¢G'VR“ì∞¢vóBWFFTÜ˜W'4∆ˆ6∑4f˜$VÁG&ñW2á&WVW7BÊFFR«¬""¬&WVW7BÊVÁG&ñW2«¬µ“¬∞¢7FGW3¢'&V¶V7FVB"¿¢&˜f≈&WVW7DñC¢&WVW7BÊñB«¬""¿¢&V¶V7FVD'ì¢7W'&VÁEW6W"ÊV÷ñ¬«¬7W'&VÁEW6W"ÊFó7∆îÊ÷R«¬'WFVÁFR ¢“ì∞¢6ˆÁ7BF&vWEW6W"“∆Ff˜&’W6W'2ÊfñÊBÇáW6W"í”‚7G&ñÊráW6W"ÁVñB«¬""í””“7G&ñÊrá&WVW7BÊ7&VFVD'ïVñB«¬""íì∞¢ñbáF&vWEW6W#ÚÊñBí∞¢vóB6VÊE&ófFT6ÜDÊ˜Fñfñ6Fñˆ‚á∞¢&V6óñVÁDñC¢F&vWEW6W"ÊñB¿¢FWáC¢)ÿ¬&ñ6ÜñW7F˜&RG∑&WVW7BÊñG“&ñfóWFF‚G∑&V6ˆ‚Ú÷˜FófÛ¢Gµ7G&ñÊrá&V6ˆ‚íÁG&ñ“Çó÷¢"'÷¿¢6VÊFW$Ê÷S¢7W'&VÁEW6W"ÊFó7∆îÊ÷R«¬7W'&VÁEW6W"ÊV÷ñ¬«¬%6ó7FV÷ ¢“ì∞¢–ß–†¶gVÊ7Fñˆ‚7F'E&W6VÊ6TÜV'F&VBÇí∞¢7F˜&W6VÊ6TÜV'F&VBÇì∞¢ñbÇ7W'&VÁEW6W"í&WGW&„∞¢&W6VÊ6TÜV'F&VEFñ÷W"“6WDñÁFW'f¬ÇÇí”‚∞¢W6W'D7W'&VÁE∆Ff˜&’W6W"Çì∞¢“¬R¢c¢ì∞ß–†¶gVÊ7Fñˆ‚7F˜&W6VÊ6TÜV'F&VBÇí∞¢ñbá&W6VÊ6TÜV'F&VEFñ÷W"í∞¢6∆V$ñÁFW'f¬á&W6VÊ6TÜV'F&VEFñ÷W"ì∞¢&W6VÊ6TÜV'F&VEFñ÷W"“ÁV∆√∞¢–ß–†¶gVÊ7Fñˆ‚7V'67&ñ&Tw5&WVW7G2Çí∞¢ñbáVÁ7V'67&ñ&Tw5&WVW7G2íVÁ7V'67&ñ&Tw5&WVW7G2Çì∞¢VÁ7V'67&ñ&Tw5&WVW7G2“F ¢Ê6ˆ∆∆V7Fñˆ‚Ç&w5WFFU&WVW7G2"ê¢Ê˜&FW$'íÇ&7&VFVDB"¬&FW62"ê¢Ê∆ñ÷óBÉ#ê¢ÊˆÂ6Ê6Ü˜BÇá6Ê6Ü˜Bí”‚∞¢w5WFFU&WVW7G2“6Ê6Ü˜BÊFˆ72Ê÷ÇÜFˆ2í”‚á≤ñC¢Fˆ2ÊñB¬‚‚ÊFˆ2ÊFFÇí“íì∞¢&VÊFW$w5&WVW7G2Çì∞¢“¬ÜW'&˜"í”‚∞¢6ˆÁ6ˆ∆RÊW'&˜"Ç$W'&˜&R6&ñ6÷VÁFÚ&ñ6ÜñW7FRu3¢"¬W'&˜"ì∞¢ñbáVíÊw5&WVW7G4∆ó7BíVíÊw5&WVW7G4∆ó7BÊñÊÊW$ÖD‘¬“#«6∆73“v◊WFVBs‰W'&˜&R6&ñ6÷VÁFÚ&ñ6ÜñW7FRu2„¬˜‚#∞¢“ì∞ß–†¶gVÊ7Fñˆ‚7F˜w5&WVW7G57V'67&óFñˆ‚Çí∞¢ñbáVÁ7V'67&ñ&Tw5&WVW7G2í∞¢VÁ7V'67&ñ&Tw5&WVW7G2Çì∞¢VÁ7V'67&ñ&Tw5&WVW7G2“ÁV∆√∞¢–¢w5WFFU&WVW7G2“µ”∞¢&VÊFW$w5&WVW7G2Çì∞ß–†¶gVÊ7Fñˆ‚&VÊFW$w5&WVW7G2Çí∞¢ñbÇVíÊw5&WVW7G4∆ó7Bí&WGW&„∞¢ñbÇ7W'&VÁEW6W"í∞¢VíÊw5&WVW7G4∆ó7BÊñÊÊW$ÖD‘¬“#«6∆73“v◊WFVBs‰fí∆ˆvñ‚W"fó7V∆óß¶&R∆R&ñ6ÜñW7FRu2„¬˜‚#∞¢&WGW&„∞¢–¢ñbÇ6‰÷ÊvTFFÇíí∞¢VíÊw5&WVW7G4∆ó7BÊñÊÊW$ÖD‘¬“#«6∆73“v◊WFVBsÂ6ˆ∆Úv∆íF÷ñ‚˜76ˆÊÚvW7Fó&R∆R&ñ6ÜñW7FRu2„¬˜‚#∞¢&WGW&„∞¢–¢ñbÇw5WFFU&WVW7G2Ê∆VÊwFÇí∞¢VíÊw5&WVW7G4∆ó7BÊñÊÊW$ÖD‘¬“#«6∆73“v◊WFVBs‰ÊW77VÊ&ñ6ÜñW7Fu2„¬˜‚#∞¢&WGW&„∞¢–¢VíÊw5&WVW7G4∆ó7BÊñÊÊW$ÖD‘¬“"#∞¢w5WFFU&WVW7G2Êf˜$V6ÇÇá&WVW7Bí”‚∞¢6ˆÁ7B&˜r“Fˆ7V÷VÁBÊ7&VFTV∆V÷VÁBÇ&Fób"ì∞¢&˜rÊ6∆74Ê÷R“'6ñ◊∆R÷∆ó7B÷óFV“#∞¢6ˆÁ7BvÜV‚“&WVW7BÊ7&VFVDBbbGóVˆb&WVW7BÊ7&VFVDBÁFÙFFR””“&gVÊ7Fñˆ‚ ¢Ú&WVW7BÊ7&VFVDBÁFÙFFRÇíÁFÙ∆ˆ6∆U7G&ñÊrÇ&óB‘ïB"ê¢¢"“#∞¢6ˆÁ7B7FGW2“7G&ñÊrá&WVW7BÁ7FGW2«¬'VÊFñÊr"íÁFıWW$66RÇì∞¢6ˆÁ7BñÊfÚ“Fˆ7V÷VÁBÊ7&VFTV∆V÷VÁBÇ'7‚"ì∞¢ñÊfÚÊñÊÊW$ÖD‘¬“G∂W66TÖD‘¬á&WVW7BÊñ◊ñÁFÙFVÊˆ÷ñÊ¶ñˆÊR«¬$ñ◊ñÁFÚ"ó“(
"G∂W66TÖD‘¬á&WVW7BÊ˜W&F˜$Ê÷R«¬$˜W&F˜&R"ó”∆'#„«6÷∆√‚G∂W66TÖD‘¬á&WVW7BÊ˜W&F˜$∆Bó“¬G∂W66TÖD‘¬á&WVW7BÊ˜W&F˜$∆Êró“(
"G∑vÜVÁ“(
"G∑7FGW7”¬˜6÷∆√Ê∞¢&˜rÊVÊD6Üñ∆BÜñÊfÚì∞†¢6ˆÁ7B7FñˆÁ2“Fˆ7V÷VÁBÊ7&VFTV∆V÷VÁBÇ&Fób"ì∞¢7FñˆÁ2Ê6∆74Ê÷R“&óFV“÷7FñˆÁ2#∞¢6ˆÁ7B6‰FV6ñFR“7G&ñÊrá&WVW7BÁ7FGW2«¬'VÊFñÊr"í””“'VÊFñÊr#∞¢7FñˆÁ2ÊVÊD6Üñ∆BÜ7&VFT'WGFˆ‚Ç$66WGF"¬Çí”‚&˜fTw5&WVW7Bá&WVW7Bí¬6‰FV6ñFRíì∞¢7FñˆÁ2ÊVÊD6Üñ∆BÜ7&VFT'WGFˆ‚Ç%&ñfóWF"¬Çí”‚&V¶V7Dw5&WVW7Bá&WVW7Bí¬6‰FV6ñFRíì∞¢&˜rÊVÊD6Üñ∆BÜ7FñˆÁ2ì∞¢VíÊw5&WVW7G4∆ó7BÊVÊD6Üñ∆Bá&˜rì∞¢“ì∞ß–†¶7ñÊ2gVÊ7Fñˆ‚&˜fTw5&WVW7Bá&WVW7Bí∞¢ñbÇ6‰÷ÊvTFFÇíí&WGW&„∞¢6ˆÁ7Bñ◊ñÁFÙñG2“'&íÊó4'&íá&WVW7BÊñ◊ñÁFÙñG2íÚ&WVW7BÊñ◊ñÁFÙñG2Êfñ«FW"Ñ&ˆˆ∆V‚í¢µ”∞¢ñbÇ&WVW7BÊ6ˆ÷÷W76ñB«¬ñ◊ñÁFÙñG2Ê∆VÊwFÇí∞¢∆W'BÇ%&ñ6ÜñW7FÊˆ‚f∆ñF¢ñ◊ñÁFÚÊˆ‚G&˜fFÚ‚"ì∞¢&WGW&„∞¢–¢vóBWFFTñ◊ñÁFÙ6ˆ˜&FñÊFW2á&WVW7BÊ6ˆ÷÷W76ñB¬ñ◊ñÁFÙñG2¬&WVW7BÊ˜W&F˜$∆B¬&WVW7BÊ˜W&F˜$∆Êrì∞¢vóBF"Ê6ˆ∆∆V7Fñˆ‚Ç&w5WFFU&WVW7G2"íÊFˆ2á&WVW7BÊñBíÁ6WBá∞¢7FGW3¢&&˜fVB"¿¢&˜fVDC¢fó&V&6RÊfó&W7F˜&R‰fñV∆Ef«VRÁ6W'fW%Fñ÷W7F◊Çí¿¢&˜fVD'ì¢7W'&VÁEW6W#ÚÊV÷ñ¬«¬" ¢“¬≤÷W&vS¢G'VR“ì∞¢vóBÊ˜Fñgîw4FV6ó6ñˆ‚á&WVW7B¬G'VRì∞ß–†¶7ñÊ2gVÊ7Fñˆ‚&V¶V7Dw5&WVW7Bá&WVW7Bí∞¢ñbÇ6‰÷ÊvTFFÇíí&WGW&„∞¢vóBF"Ê6ˆ∆∆V7Fñˆ‚Ç&w5WFFU&WVW7G2"íÊFˆ2á&WVW7BÊñBíÁ6WBá∞¢7FGW3¢'&V¶V7FVB"¿¢&V¶V7FVDC¢fó&V&6RÊfó&W7F˜&R‰fñV∆Ef«VRÁ6W'fW%Fñ÷W7F◊Çí¿¢&V¶V7FVD'ì¢7W'&VÁEW6W#ÚÊV÷ñ¬«¬" ¢“¬≤÷W&vS¢G'VR“ì∞¢vóBÊ˜Fñgîw4FV6ó6ñˆ‚á&WVW7B¬f«6Rì∞ß–†¶7ñÊ2gVÊ7Fñˆ‚Ê˜Fñgîw4FV6ó6ñˆ‚á&WVW7B¬&˜fVBí∞¢ñbÇ&WVW7BÊ˜W&F˜$ñBí&WGW&„∞¢vóBF"Ê6ˆ∆∆V7Fñˆ‚Ç&6ÜD÷W76vW2"íÊFBá∞¢GóS¢'FWáB"¿¢FWáC¢&˜fV@¢Ú)»R&ñ6ÜñW7Fu266WGFFW"G∑&WVW7BÊñ◊ñÁFÙFVÊˆ÷ñÊ¶ñˆÊR«¬&ñ◊ñÁFÚ'“‚6ˆ˜&FñÊFRvvñ˜&ÊFRÊ ¢¢)ÿ¬&ñ6ÜñW7Fu2&ñfóWFFW"G∑&WVW7BÊñ◊ñÁFÙFVÊˆ÷ñÊ¶ñˆÊR«¬&ñ◊ñÁFÚ'“Ê¿¢&V6óñVÁDñC¢&WVW7BÊ˜W&F˜$ñB¿¢6VÊFW$ñC¢7W'&VÁEW6W#ÚÁVñB«¬""¿¢6VÊFW$Ê÷S¢7W'&VÁEW6W#ÚÊFó7∆îÊ÷R«¬7W'&VÁEW6W#ÚÊV÷ñ¬«¬$F÷ñ‚"¿¢6VÊFW$V÷ñ√¢7W'&VÁEW6W#ÚÊV÷ñ¬«¬""¿¢7&VFVDC¢fó&V&6RÊfó&W7F˜&R‰fñV∆Ef«VRÁ6W'fW%Fñ÷W7F◊Çê¢“ì∞ß–†¶7ñÊ2gVÊ7Fñˆ‚WFFTñ◊ñÁFÙ6ˆ˜&FñÊFW2Ü6ˆ÷÷W76ñB¬ñ◊ñÁFÙñG2¬∆B¬∆Êrí∞¢6ˆÁ7B&Vb“F"Ê6ˆ∆∆V7Fñˆ‚ÜvWD6ˆ÷÷W76T6ˆ∆∆V7Fñˆ‰Ê÷RÇííÊFˆ2Ü6ˆ÷÷W76ñBíÊ6ˆ∆∆V7Fñˆ‚Ç&ñ◊ñÁFí"ì∞¢vóB&ˆ÷ó6RÊ∆¬Üñ◊ñÁFÙñG2Ê÷ÇÜñ◊ñÁFÙñBí”‚&VbÊFˆ2Üñ◊ñÁFÙñBíÁWFFRá∞¢w5ì¢ÁV÷&W"Ü∆Bí¿¢w5É¢ÁV÷&W"Ü∆Êrí¿¢WFFVDC¢fó&V&6RÊfó&W7F˜&R‰fñV∆Ef«VRÁ6W'fW%Fñ÷W7F◊Çí¿¢WFFVD'ì¢7W'&VÁEW6W#ÚÊV÷ñ¬«¬&F÷ñ‚ ¢“ííì∞ß–†¶gVÊ7Fñˆ‚vWD7W'&VÁE∆Ff˜&’W6W%&˜rÇí∞¢ñbÇ7W'&VÁEW6W"í&WGW&‚ÁV∆√∞¢&WGW&‚∆Ff˜&’W6W'2ÊfñÊBÇáW6W"í”‚W6W"ÊñB””“7W'&VÁEW6W"ÁVñBí«¬ÁV∆√∞ß–†¶gVÊ7Fñˆ‚vWDFVÊñVD7FñˆÁ4f˜$7W'&VÁEW6W"Çí∞¢6ˆÁ7B&˜r“vWD7W'&VÁE∆Ff˜&’W6W%&˜rÇì∞¢6ˆÁ7BFVÊñVB“'&íÊó4'&íá&˜sÚÊFVÊñVDñ◊ñÁFÙ7FñˆÁ2íÚ&˜rÊFVÊñVDñ◊ñÁFÙ7FñˆÁ2¢µ”∞¢&WGW&‚ÊWr6WBÜFVÊñVBÊfñ«FW"ÇÜ7Fñˆ‚í”‚î’îÂDıÙ5DîÙÂ2ÊñÊ6«VFW2Ü7Fñˆ‚ííì∞ß–†¶gVÊ7Fñˆ‚vWD∆∆˜vVD7FñˆÁ4f˜$7W'&VÁEW6W"Çí∞¢6ˆÁ7B&˜r“vWD7W'&VÁE∆Ff˜&’W6W%&˜rÇì∞¢6ˆÁ7B∆∆˜vVB“'&íÊó4'&íá&˜sÚÊ∆∆˜vVDñ◊ñÁFÙ7FñˆÁ2íÚ&˜rÊ∆∆˜vVDñ◊ñÁFÙ7FñˆÁ2¢µ”∞¢&WGW&‚ÊWr6WBÜ∆∆˜vVBÊfñ«FW"ÇÜ7Fñˆ‚í”‚î’îÂDıÙ5DîÙÂ2ÊñÊ6«VFW2Ü7Fñˆ‚ííì∞ß–†¶gVÊ7Fñˆ‚6ÂW6Tñ◊ñÁFÙ7Fñˆ‚Ü7Fñˆ‚í∞¢ñbÇî’îÂDıÙ5DîÙÂ2ÊñÊ6«VFW2Ü7Fñˆ‚íí&WGW&‚f«6S∞¢ñbÜ6‰÷ÊvTFFÇíí&WGW&‚G'VS∞¢ñbÑD‘îÂÙÙ‰≈ïÙî’îÂDıÙ5DîÙÂ2ÊñÊ6«VFW2Ü7Fñˆ‚íí&WGW&‚vWD∆∆˜vVD7FñˆÁ4f˜$7W'&VÁEW6W"ÇíÊÜ2Ü7Fñˆ‚ì∞¢ñbÜ7Fñˆ‚””“&FˆÊR"«¬7Fñˆ‚””“'vÜG6"í&WGW&‚G'VS∞¢&WGW&‚FVÊñVDñ◊ñÁFÙ7FñˆÁ2ÊÜ2Ü7Fñˆ‚ì∞ß–†¶gVÊ7Fñˆ‚ó4ñ◊ñÁFÙ7Fñˆ‰FVÊñVBÜ7Fñˆ‚í∞¢&WGW&‚6ÂW6Tñ◊ñÁFÙ7Fñˆ‚Ü7Fñˆ‚ì∞ß–†¶gVÊ7Fñˆ‚7Fñˆ‰∆&V¬Ü7Fñˆ‚í∞¢6ˆÁ7B÷“∞¢FˆÊS¢.)»RfGFÚ"¿¢ÊfñvFS¢/	˙z“Êfñv"¿¢&W6WC¢.)õæ˚àÚ&W6WB"¿¢vÜG6¢/	˘*¬vÜG4"¿¢'&ˆ&∆V“◊&W˜'B#¢/	˘™Ç6VvÊ∆&ˆ&∆V÷"¿¢&w2◊WFFR#¢/	˘8“vvñ˜&Êu2"¿¢VFóC¢.)»˛˚àÚ÷ˆFñfñ6"¿¢FV∆WFS¢/	˘y˚àÚV∆ñ÷ñÊ ¢”∞¢&WGW&‚÷∂7FñˆÂ“«¬7Fñˆ„∞ß–†¶gVÊ7Fñˆ‚&VÊFW%W6W%W&÷ó76ñˆ‰∆ó7BÇí∞¢ñbÇVíÁW6W%W&÷ó76ñˆÁ4∆ó7Bí&WGW&„∞¢ñbÇ7W'&VÁEW6W"í∞¢VíÁW6W%W&÷ó76ñˆÁ4∆ó7BÊñÊÊW$ÖD‘¬“#«6∆73“v◊WFVBs‰fí∆ˆvñ‚W"vW7Fó&RíW&÷W76í„¬˜‚#∞¢&WGW&„∞¢–¢ñbÇ6‰÷ÊvTFFÇíí∞¢VíÁW6W%W&÷ó76ñˆÁ4∆ó7BÊñÊÊW$ÖD‘¬“#«6∆73“v◊WFVBsÂ6ˆ∆Úv∆íF÷ñ‚˜76ˆÊÚ6÷&ñ&RíW&÷W76í¶ñˆÊR„¬˜‚#∞¢&WGW&„∞¢–¢6ˆÁ7BW6W'2“∆Ff˜&’W6W'2Êfñ«FW"ÇáW6W"í”‚F÷ñ‰V÷ñ«2ÊÜ2ÜÊ˜&÷∆ó¶TV÷ñ¬áW6W"ÊV÷ñ¬ííì∞¢ñbÇW6W'2Ê∆VÊwFÇí∞¢VíÁW6W%W&÷ó76ñˆÁ4∆ó7BÊñÊÊW$ÖD‘¬“#«6∆73“v◊WFVBs‰ÊW77V‚WFVÁFRFó7ˆÊñ&ñ∆R„¬˜‚#∞¢&WGW&„∞¢–¢VíÁW6W%W&÷ó76ñˆÁ4∆ó7BÊñÊÊW$ÖD‘¬“"#∞¢W6W'2Êf˜$V6ÇÇáW6W"í”‚∞¢6ˆÁ7B&˜r“Fˆ7V÷VÁBÊ7&VFTV∆V÷VÁBÇ&Fób"ì∞¢&˜rÊ6∆74Ê÷R“'6ñ◊∆R÷∆ó7B÷óFV“7F6∂VB#∞¢6ˆÁ7BFóF∆R“Fˆ7V÷VÁBÊ7&VFTV∆V÷VÁBÇ'7G&ˆÊr"ì∞¢FóF∆RÁFWáD6ˆÁFVÁB“W6W"ÊFó7∆îÊ÷R«¬W6W"ÊV÷ñ¬«¬W6W"ÊñC∞¢&˜rÊVÊD6Üñ∆BáFóF∆Rì∞¢6ˆÁ7BÜñÁB“Fˆ7V÷VÁBÊ7&VFTV∆V÷VÁBÇ'6÷∆¬"ì∞¢ÜñÁBÊ6∆74Ê÷R“&◊WFVB#∞¢ÜñÁBÁFWáD6ˆÁFVÁB“%fW&FR“¶ñˆÊR6ˆÊ6W76˜fó6ñ&ñ∆R‚&«R“¶ñˆÊRÊ66˜7FÚÊˆ‚6ˆÊ6W76‚#∞¢&˜rÊVÊD6Üñ∆BÜÜñÁBì∞¢6ˆÁ7B7Fñˆ‰&˜Ç“Fˆ7V÷VÁBÊ7&VFTV∆V÷VÁBÇ&Fób"ì∞¢7Fñˆ‰&˜ÇÊ6∆74Ê÷R“&7FñˆÁ2◊&˜r#∞¢6ˆÁ7BFVÊñVB“ÊWr6WBÑ'&íÊó4'&íáW6W"ÊFVÊñVDñ◊ñÁFÙ7FñˆÁ2íÚW6W"ÊFVÊñVDñ◊ñÁFÙ7FñˆÁ2¢µ“ì∞¢6ˆÁ7B∆∆˜vVB“ÊWr6WBÑ'&íÊó4'&íáW6W"Ê∆∆˜vVDñ◊ñÁFÙ7FñˆÁ2íÚW6W"Ê∆∆˜vVDñ◊ñÁFÙ7FñˆÁ2¢µ“ì∞¢î’îÂDıÙ5DîÙÂ2Êf˜$V6ÇÇÜ7Fñˆ‚í”‚∞¢6ˆÁ7BF÷ñ‰ˆÊ«í“D‘îÂÙÙ‰≈ïÙî’îÂDıÙ5DîÙÂ2ÊñÊ6«VFW2Ü7Fñˆ‚ì∞¢6ˆÁ7BVÊ&∆VB“F÷ñ‰ˆÊ«íÚ∆∆˜vVBÊÜ2Ü7Fñˆ‚í¢FVÊñVBÊÜ2Ü7Fñˆ‚ì∞¢6ˆÁ7B'F‚“7&VFT'WGFˆ‚ÜG∂VÊ&∆VBÚ.)»R"¢/	˘™≤'“G∂7Fñˆ‰∆&V¬Ü7Fñˆ‚ó÷¬7ñÊ2Çí”‚∞¢6ˆÁ7BÊWáDFVÊñVB“ÊWr6WBÜFVÊñVBì∞¢6ˆÁ7BÊWáD∆∆˜vVB“ÊWr6WBÜ∆∆˜vVBì∞¢ñbÜF÷ñ‰ˆÊ«íí∞¢ñbÜÊWáD∆∆˜vVBÊÜ2Ü7Fñˆ‚ííÊWáD∆∆˜vVBÊFV∆WFRÜ7Fñˆ‚ì∞¢V«6RÊWáD∆∆˜vVBÊFBÜ7Fñˆ‚ì∞¢“V«6RñbÜÊWáDFVÊñVBÊÜ2Ü7Fñˆ‚íí∞¢ÊWáDFVÊñVBÊFV∆WFRÜ7Fñˆ‚ì∞¢“V«6R∞¢ÊWáDFVÊñVBÊFBÜ7Fñˆ‚ì∞¢–¢vóBF"Ê6ˆ∆∆V7Fñˆ‚Ç'∆Ff˜&’W6W'2"íÊFˆ2áW6W"ÊñBíÁ6WBá∞¢FVÊñVDñ◊ñÁFÙ7FñˆÁ3¢'&íÊg&ˆ“ÜÊWáDFVÊñVBí¿¢∆∆˜vVDñ◊ñÁFÙ7FñˆÁ3¢'&íÊg&ˆ“ÜÊWáD∆∆˜vVBí¿¢WFFVDC¢fó&V&6RÊfó&W7F˜&R‰fñV∆Ef«VRÁ6W'fW%Fñ÷W7F◊Çí¿¢WFFVD'ì¢7W'&VÁEW6W#ÚÊV÷ñ¬«¬" ¢“¬≤÷W&vS¢G'VR“ì∞¢“ì∞¢'F‚Ê6∆74∆ó7BÊFBÇ&'F‚◊6÷∆¬"¬VÊ&∆VBÚ&'F‚◊W&÷ó76ñˆ‚÷ˆ‚"¢&'F‚◊W&÷ó76ñˆ‚÷ˆfb"ì∞¢'F‚ÁFóF∆R“F÷ñ‰ˆÊ«ê¢ÚÜVÊ&∆VBÚ%W&÷W76ÚF÷ñ‚6ˆÊ6W76Û¢6∆ñ66W"&Wfˆ6&R"¢%W&÷W76ÚF÷ñ‚Êˆ‚6ˆÊ6W76Û¢6∆ñ66W"6ˆÊ6VFW&R"ê¢¢ÜVÊ&∆VBÚ%V«6ÁFRfó6ñ&ñ∆S¢6∆ñ66W"Ê66ˆÊFW&∆Ú"¢%V«6ÁFRÊ66˜7FÛ¢6∆ñ66W"÷˜7G&&∆Ú"ì∞¢7Fñˆ‰&˜ÇÊVÊD6Üñ∆BÜ'F‚ì∞¢“ì∞¢&˜rÊVÊD6Üñ∆BÜ7Fñˆ‰&˜Çì∞¢VíÁW6W%W&÷ó76ñˆÁ4∆ó7BÊVÊD6Üñ∆Bá&˜rì∞¢“ì∞ß–†¶gVÊ7Fñˆ‚&VÊFW%W6W$&‰∆ó7BÇí∞¢ñbÇVíÁW6W$&‰∆ó7Bí&WGW&„∞¢ñbÇ6‰÷ÊvTFFÇíí∞¢VíÁW6W$&‰∆ó7BÊñÊÊW$ÖD‘¬“#«6∆73“v◊WFVBsÂ6ˆ∆ÚV‚F÷ñ‚\;"&∆ˆ66&RÚ6&∆ˆ66&RWFVÁFí„¬˜‚#∞¢&WGW&„∞¢–¢6ˆÁ7BW6W'2“∆Ff˜&’W6W'2Êfñ«FW"ÇáW6W"í”‚7G&ñÊráW6W"ÊñB«¬W6W"ÁVñB«¬""í”“7G&ñÊrÜ7W'&VÁEW6W#ÚÁVñB«¬""íì∞¢ñbÇW6W'2Ê∆VÊwFÇí∞¢VíÁW6W$&‰∆ó7BÊñÊÊW$ÖD‘¬“#«6∆73“v◊WFVBs‰ÊW77V‚WFVÁFRFó7ˆÊñ&ñ∆R„¬˜‚#∞¢&WGW&„∞¢–¢VíÁW6W$&‰∆ó7BÊñÊÊW$ÖD‘¬“"#∞¢W6W'2Êf˜$V6ÇÇáW6W"í”‚∞¢6ˆÁ7B&ÊÊVB“&ˆˆ∆V‚áW6W"Ê&ÊÊVBì∞¢6ˆÁ7B&˜r“Fˆ7V÷VÁBÊ7&VFTV∆V÷VÁBÇ&Fób"ì∞¢&˜rÊ6∆74Ê÷R“'6ñ◊∆R÷∆ó7B÷óFV“7F6∂VB#∞¢6ˆÁ7BFóF∆R“Fˆ7V÷VÁBÊ7&VFTV∆V÷VÁBÇ'7G&ˆÊr"ì∞¢FóF∆RÁFWáD6ˆÁFVÁB“W6W"ÊFó7∆îÊ÷R«¬W6W"ÊV÷ñ¬«¬W6W"ÊñC∞¢&˜rÊVÊD6Üñ∆BáFóF∆Rì∞¢6ˆÁ7B7FGW2“Fˆ7V÷VÁBÊ7&VFTV∆V÷VÁBÇ'7‚"ì∞¢7FGW2Ê6∆74Ê÷R“&ÊÊVBÚ'W6W"÷&‚◊7FGW2W6W"÷&‚◊7FGW2“÷&ÊÊVB"¢'W6W"÷&‚◊7FGW2W6W"÷&‚◊7FGW2“÷7FófR#∞¢7FGW2ÁFWáD6ˆÁFVÁB“&ÊÊVBÚ/	˘KB66W76ÚÊVvFÚ"¢/	˘˙"GFófÚ#∞¢&˜rÊVÊD6Üñ∆Bá7FGW2ì∞¢6ˆÁ7B'F‚“7&VFT'WGFˆ‚Ü&ÊÊVBÚ%4$ƒÙ44UDTÂDR"¢$$‰‰UDTÂDR"¬Çí”‚6WEW6W$&ÊÊVBáW6W"¬&ÊÊVBíì∞¢'F‚Ê6∆74∆ó7BÊFBÇ&'F‚◊6÷∆¬"¬&ÊÊVBÚ&'F‚◊&ñ÷'í"¢&'F‚÷FÊvW""ì∞¢&˜rÊVÊD6Üñ∆BÜ'F‚ì∞¢VíÁW6W$&‰∆ó7BÊVÊD6Üñ∆Bá&˜rì∞¢“ì∞ß–†¶7ñÊ2gVÊ7Fñˆ‚6WEW6W$&ÊÊVBáW6W"¬&ÊÊVBí∞¢ñbÇ6‰÷ÊvTFFÇíí&WGW&‚∆W'BÇ%6ˆ∆ÚV‚F÷ñ‚\;"&∆ˆ66&RÚ6&∆ˆ66&RWFVÁFí‚"ì∞¢6ˆÁ7BW6W$ñB“7G&ñÊráW6W#ÚÊñB«¬W6W#ÚÁVñB«¬""ì∞¢ñbÇW6W$ñBí&WGW&„∞¢6ˆÁ7Bñ∆ˆB“&ÊÊV@¢Ú≤&ÊÊVC¢G'VR¬7FFÙ66˜VÁC¢&&∆ˆ66FÚ"¬66˜VÁE7FGW3¢&&∆ˆ66FÚ"¬&ÊÊVE&V6ˆ„¢$&∆ˆ66FÚF÷÷ñÊó7G&F˜&R"¬&ÊÊVDC¢fó&V&6RÊfó&W7F˜&R‰fñV∆Ef«VRÁ6W'fW%Fñ÷W7F◊Çí¬&ÊÊVD'ì¢7W'&VÁEW6W#ÚÊV÷ñ¬«¬7W'&VÁEW6W#ÚÁVñB«¬&F÷ñ‚"–¢¢≤&ÊÊVC¢f«6R¬7FFÙ66˜VÁC¢&GFófÚ"¬66˜VÁE7FGW3¢&GFófÚ"¬&ÊÊVE&V6ˆ„¢fó&V&6RÊfó&W7F˜&R‰fñV∆Ef«VRÊFV∆WFRÇí¬&ÊÊVDC¢fó&V&6RÊfó&W7F˜&R‰fñV∆Ef«VRÊFV∆WFRÇí¬&ÊÊVD'ì¢fó&V&6RÊfó&W7F˜&R‰fñV∆Ef«VRÊFV∆WFRÇí”∞¢6ˆÁ7B&F6Ç“F"Ê&F6ÇÇì∞¢&F6ÇÁ6WBÜF"Ê6ˆ∆∆V7Fñˆ‚Ç'∆Ff˜&’W6W'2"íÊFˆ2áW6W$ñBí¬∞¢‚‚Áñ∆ˆB¿¢WFFVDC¢fó&V&6RÊfó&W7F˜&R‰fñV∆Ef«VRÁ6W'fW%Fñ÷W7F◊Çí¿¢WFFVD'ì¢7W'&VÁEW6W#ÚÊV÷ñ¬«¬" ¢“¬≤÷W&vS¢G'VR“ì∞¢&F6ÇÁ6WBÜF"Ê6ˆ∆∆V7Fñˆ‚Ç'W6W$66W74VFóB"íÊFˆ2Çí¬≤W6W$ñB¬7Fñˆ„¢&ÊÊVBÚ&&∆ˆ66Ú"¢'&ñGFóf¶ñˆÊR"¬F÷ñÊó7G&F˜%VñC¢7W'&VÁEW6W#ÚÁVñB«¬""¬F÷ñÊó7G&F˜$V÷ñ√¢7W'&VÁEW6W#ÚÊV÷ñ¬«¬""¬7&VFVDC¢fó&V&6RÊfó&W7F˜&R‰fñV∆Ef«VRÁ6W'fW%Fñ÷W7F◊Çí“ì∞¢vóB&F6ÇÊ6ˆ÷÷óBÇì∞ß–†¶gVÊ7Fñˆ‚vWE∆Ff˜&’W6W$∆&V¬áW6W"í∞¢ñbÇW6W"í&WGW&‚%WFVÁFR66ˆÊ˜66óWFÚ#∞¢&WGW&‚7G&ñÊráW6W"ÊFó7∆îÊ÷R«¬W6W"ÊV÷ñ¬«¬W6W"ÊñB«¬%WFVÁFR"íÁG&ñ“Çí«¬%WFVÁFR#∞ß–†¶gVÊ7Fñˆ‚&VÊFW$Ê˜Fñfñ6FñˆÂF&vWEW6W'2Çí∞¢ñbÇVíÊÊ˜Fñfñ6FñˆÂW6W%6V∆V7Bí&WGW&„∞¢6ˆÁ7B&Wfñ˜W2“'&íÊg&ˆ“áVíÊÊ˜Fñfñ6FñˆÂW6W%6V∆V7BÁ6V∆V7FVD˜FñˆÁ2«¬µ“íÊ÷ÇÜ˜Bí”‚˜BÁf«VRì∞¢VíÊÊ˜Fñfñ6FñˆÂW6W%6V∆V7BÊñÊÊW$ÖD‘¬“"#∞¢6ˆÁ7B&V6óñVÁG2“∆Ff˜&’W6W'2Êfñ«FW"ÇáW6W"í”‚F÷ñ‰V÷ñ«2ÊÜ2ÜÊ˜&÷∆ó¶TV÷ñ¬áW6W"ÊV÷ñ¬ííì∞¢&V6óñVÁG2Êf˜$V6ÇÇáW6W"í”‚∞¢6ˆÁ7B˜B“Fˆ7V÷VÁBÊ7&VFTV∆V÷VÁBÇ&˜Fñˆ‚"ì∞¢˜BÁf«VR“W6W"ÊñC∞¢˜BÁFWáD6ˆÁFVÁB“G∂vWE∆Ff˜&’W6W$∆&V¬áW6W"ó“G∑W6W"ÊV÷ñ¬ÚÇG∑W6W"ÊV÷ñ«“ñ¢"'÷∞¢ñbá&Wfñ˜W2ÊñÊ6«VFW2áW6W"ÊñBíí˜BÁ6V∆V7FVB“G'VS∞¢VíÊÊ˜Fñfñ6FñˆÂW6W%6V∆V7BÊVÊD6Üñ∆BÜ˜Bì∞¢“ì∞¢ˆ‰Ê˜Fñfñ6FñˆÂ6VÊD∆ƒ6ÜÊvRÇì∞ß–†¶gVÊ7Fñˆ‚ˆ‰Ê˜Fñfñ6FñˆÂ6VÊD∆ƒ6ÜÊvRÇí∞¢6ˆÁ7B6VÊD∆¬“&ˆˆ∆V‚áVíÊÊ˜Fñfñ6FñˆÂ6VÊD∆≈Fˆvv∆SÚÊ6ÜV6∂VBì∞¢ñbáVíÊÊ˜Fñfñ6FñˆÂW6W%6V∆V7BíVíÊÊ˜Fñfñ6FñˆÂW6W%6V∆V7BÊFó6&∆VB“6VÊD∆√∞ß–†¶gVÊ7Fñˆ‚&VÊFW$Ê˜Fñfñ6FñˆÁ4∆ó7BÇí∞¢ñbÇVíÊÊ˜Fñfñ6FñˆÁ4∆ó7Bí&WGW&„∞¢ñbÇ7W'&VÁEW6W"í∞¢VíÊÊ˜Fñfñ6FñˆÁ4∆ó7BÊñÊÊW$ÖD‘¬“#«6∆73“v◊WFVBs‰fí∆ˆvñ‚W"fVFW&R∆RÊ˜Fñfñ6ÜR„¬˜‚#∞¢&WGW&„∞¢–¢ñbÇ6‰÷ÊvTFFÇíí∞¢VíÊÊ˜Fñfñ6FñˆÁ4∆ó7BÊñÊÊW$ÖD‘¬“#«6∆73“v◊WFVBsÂ6ˆ∆Úv∆íF÷ñ‚˜76ˆÊÚvW7Fó&R∆RÊ˜Fñfñ6ÜR„¬˜‚#∞¢&WGW&„∞¢–¢ñbÇW6W$∆W'G2Ê∆VÊwFÇí∞¢VíÊÊ˜Fñfñ6FñˆÁ4∆ó7BÊñÊÊW$ÖD‘¬“#«6∆73“v◊WFVBs‰ÊW77VÊÊ˜Fñfñ6Fó7ˆÊñ&ñ∆R„¬˜‚#∞¢&WGW&„∞¢–¢VíÊÊ˜Fñfñ6FñˆÁ4∆ó7BÊñÊÊW$ÖD‘¬“"#∞¢W6W$∆W'G2Êf˜$V6ÇÇÜóFV“í”‚∞¢6ˆÁ7BW6W$∆&V¬“f˜&÷DÊ˜Fñfñ6FñˆÂ&V6óñVÁG4∆&V¬ÜóFV“ì∞¢6ˆÁ7B&˜r“Fˆ7V÷VÁBÊ7&VFTV∆V÷VÁBÇ&'Fñ6∆R"ì∞¢&˜rÊ6∆74Ê÷R“'6ñ◊∆R÷∆ó7B÷óFV“7F6∂VB#∞¢6ˆÁ7B7FGW2“óFV“Á66ÜVGV∆VDFFT∂WíbbóFV“Á66ÜVGV∆VDFFT∂Wí‚vWDFFT∂Wîg&ˆ‘∆ˆ6ƒFFRÜÊWrFFRÇííÚ%&ˆw&÷“‚gWGW&"¢$GFóf#∞¢6ˆÁ7BFóF∆R“7G&ñÊrÜóFV“ÁFóF∆R«¬$Ê˜Fñfñ6"íÁG&ñ“Çì∞¢6ˆÁ7BGF6Ü÷VÁG2“'&íÊó4'&íÜóFV“ÊGF6Ü÷VÁG2íÚóFV“ÊGF6Ü÷VÁG2¢µ”∞¢&˜rÊñÊÊW$ÖD‘¬“«7G&ˆÊs‚G∂W66TÖD‘¬áFóF∆Ró”¬˜7G&ˆÊs„«‚G∂W66TÖD‘¬ÜóFV“Ê÷W76vR«¬""ó”¬˜„«6÷∆√‚G∂W66TÖD‘¬áW6W$∆&V¬ó“(
"G∂W66TÖD‘¬á7FGW2ó”¬˜6÷∆√Ê∞¢6ˆÁ7BGF6Ü÷VÁG4&˜Ç“Fˆ7V÷VÁBÊ7&VFTV∆V÷VÁBÇ&Fób"ì∞¢ñbÇGF6Ü÷VÁG2Ê∆VÊwFÇí∞¢GF6Ü÷VÁG4&˜ÇÊñÊÊW$ÖD‘¬“#«6÷∆√‰ÊW77V‚∆∆VvFÚ„¬˜6÷∆√‚#∞¢“V«6R∞¢6ˆÁ7B∆ó7B“Fˆ7V÷VÁBÊ7&VFTV∆V÷VÁBÇ'V¬"ì∞¢GF6Ü÷VÁG2Êf˜$V6ÇÇÜGBí”‚∞¢6ˆÁ7B∆í“Fˆ7V÷VÁBÊ7&VFTV∆V÷VÁBÇ&∆í"ì∞¢6ˆÁ7B∆ñÊ≤“Fˆ7V÷VÁBÊ7&VFTV∆V÷VÁBÇ&"ì∞¢∆ñÊ≤Êá&Vb“"2#∞¢∆ñÊ≤ÁFWáD6ˆÁFVÁB“	˘8‚G∂GBÊÊ÷R«¬$Fˆ7V÷VÁFÚ'÷∞¢∆ñÊ≤ÊFDWfVÁD∆ó7FVÊW"Ç&6∆ñ6≤"¬ÜWfVÁBí”‚∞¢WfVÁBÁ&WfVÁDFVfV«BÇì∞¢˜V‰Ê˜Fñfñ6Fñˆ‰Fˆ7V÷VÁEfñWvW"ÜGBÁW&¬«¬""¬GBÊÊ÷R«¬$Fˆ7V÷VÁFÚ"ì∞¢“ì∞¢∆íÊVÊD6Üñ∆BÜ∆ñÊ≤ì∞¢∆ó7BÊVÊD6Üñ∆BÜ∆íì∞¢“ì∞¢GF6Ü÷VÁG4&˜ÇÊVÊD6Üñ∆BÜ∆ó7Bì∞¢–¢&˜rÊVÊD6Üñ∆BÜGF6Ü÷VÁG4&˜Çì∞¢6ˆÁ7B7FñˆÁ2“Fˆ7V÷VÁBÊ7&VFTV∆V÷VÁBÇ&Fób"ì∞¢7FñˆÁ2Ê6∆74Ê÷R“&óFV“÷7FñˆÁ2#∞¢7FñˆÁ2ÊVÊD6Üñ∆BÜ7&VFT'WGFˆ‚Ç$V∆ñ÷ñÊ"¬Çí”‚FV∆WFUW6W$Ê˜Fñfñ6Fñˆ‚ÜóFV“ÊñBííì∞¢7FñˆÁ2ÊVÊD6Üñ∆BÜ7&VFT'WGFˆ‚Ç$FWGFv∆ñÚvñ˜&ÊÚ"¬Çí”‚˜V‰Ê˜Fñfñ6Fñˆ‰FîFWFñ¬ÜvWDÊ˜Fñfñ6FñˆÂ&ñ÷'îFFT∂WíÜóFV“íííì∞¢&˜rÊVÊD6Üñ∆BÜ7FñˆÁ2ì∞¢VíÊÊ˜Fñfñ6FñˆÁ4∆ó7BÊVÊD6Üñ∆Bá&˜rì∞¢“ì∞ß–†¶gVÊ7Fñˆ‚vWDÊ˜Fñfñ6FñˆÂ&V6óñVÁEW6W$ñG2Ü∆W'DóFV“í∞¢ñbÇ∆W'DóFV“í&WGW&‚µ”∞¢ñbÑ'&íÊó4'&íÜ∆W'DóFV“ÁF&vWEW6W$ñG2íí&WGW&‚∆W'DóFV“ÁF&vWEW6W$ñG2Ê÷ÇÜñBí”‚7G&ñÊrÜñB«¬""íÁG&ñ“ÇííÊfñ«FW"Ñ&ˆˆ∆V‚ì∞¢ñbÜ∆W'DóFV“ÁF&vWEW6W$ñBí&WGW&‚µ7G&ñÊrÜ∆W'DóFV“ÁF&vWEW6W$ñBíÁG&ñ“Çï“Êfñ«FW"Ñ&ˆˆ∆V‚ì∞¢&WGW&‚µ”∞ß–†¶gVÊ7Fñˆ‚ó4Ê˜Fñfñ6Fñˆ‰f˜$7W'&VÁEW6W"Ü∆W'DóFV“í∞¢ñbÇ∆W'DóFV“«¬7W'&VÁEW6W"í&WGW&‚f«6S∞¢ñbÜ∆W'DóFV“Á6VÊEFÙ∆≈&Vvó7FW&VBí&WGW&‚G'VS∞¢ñbÜvWDÊ˜Fñfñ6FñˆÂ&V6óñVÁEW6W$ñG2Ü∆W'DóFV“íÊñÊ6«VFW2Ü7W'&VÁEW6W"ÁVñBíí&WGW&‚G'VS∞¢6ˆÁ7B÷V÷&W$Ê÷W2“'&íÊó4'&íÜ∆W'DóFV“ÁF&vWD÷V÷&W$Ê÷W2íÚ∆W'DóFV“ÁF&vWD÷V÷&W$Ê÷W2¢µ”∞¢&WGW&‚÷V÷&W$Ê÷W2Á6ˆ÷RÇÜ÷V÷&W$Ê÷Rí”‚Fı7VG&÷V÷&W$ÊEW6W$÷F6ÇÜ÷V÷&W$Ê÷Ríì∞ß–†¶gVÊ7Fñˆ‚f˜&÷DÊ˜Fñfñ6FñˆÂ&V6óñVÁG4∆&V¬Ü∆W'DóFV“í∞¢ñbÜ∆W'DóFV”ÚÁ6VÊEFÙ∆≈&Vvó7FW&VBí&WGW&‚$FW7FñÊF&ì¢GWGFív∆íWFVÁFí&Vvó7G&Fí#∞¢6ˆÁ7BñG2“vWDÊ˜Fñfñ6FñˆÂ&V6óñVÁEW6W$ñG2Ü∆W'DóFV“ì∞¢ñbÇñG2Ê∆VÊwFÇí&WGW&‚$FW7FñÊF&íÊˆ‚ñ◊˜7FFí#∞¢6ˆÁ7B∆&V«2“ñG2Ê÷ÇÜñBí”‚vWE∆Ff˜&’W6W$∆&V¬á∆Ff˜&’W6W'2ÊfñÊBÇáRí”‚RÊñB””“ñBííì∞¢&WGW&‚FW7FñÊF&ì¢G∂∆&V«2Ê¶ˆñ‚Ç"¬"ó÷∞ß–†¶gVÊ7Fñˆ‚vWDÊ˜Fñfñ6FñˆÂ&ñ÷'îFFT∂WíÜóFV“í∞¢&WGW&‚7G&ñÊrÜóFV”ÚÁ66ÜVGV∆VDFFT∂Wí«¬óFV”ÚÊ7&VFVDFFT∂Wí«¬""íÁG&ñ“Çì∞ß–†¶gVÊ7Fñˆ‚&VÊFW$7FófUW6W$∆W'DGF6Ü÷VÁG2Çí∞¢ñbÇVíÁW6W$∆W'DGF6Ü÷VÁG2í&WGW&„∞¢6ˆÁ7BGF6Ü÷VÁG2“'&íÊó4'&íÜ7FófUW6W$∆W'CÚÊGF6Ü÷VÁG2íÚ7FófUW6W$∆W'BÊGF6Ü÷VÁG2¢µ”∞¢ñbÇGF6Ü÷VÁG2Ê∆VÊwFÇí∞¢VíÁW6W$∆W'DGF6Ü÷VÁG2ÊñÊÊW$ÖD‘¬“#«6∆73“v◊WFVBs‰ÊW77V‚Fˆ7V÷VÁFÚ∆∆VvFÚ„¬˜‚#∞¢&WGW&„∞¢–¢VíÁW6W$∆W'DGF6Ü÷VÁG2ÊñÊÊW$ÖD‘¬“"#∞¢GF6Ü÷VÁG2Êf˜$V6ÇÇÜGF6Ü÷VÁBí”‚∞¢6ˆÁ7B&˜r“Fˆ7V÷VÁBÊ7&VFTV∆V÷VÁBÇ&Fób"ì∞¢&˜rÊ6∆74Ê÷R“'6ñ◊∆R÷∆ó7B÷óFV“#∞¢6ˆÁ7B∆&V¬“Fˆ7V÷VÁBÊ7&VFTV∆V÷VÁBÇ'7‚"ì∞¢∆&V¬ÁFWáD6ˆÁFVÁB“	˘8‚G∂GF6Ü÷VÁBÊÊ÷R«¬$Fˆ7V÷VÁFÚ'÷∞¢&˜rÊVÊD6Üñ∆BÜ∆&V¬ì∞¢6ˆÁ7B˜V‰'F‚“7&VFT'WGFˆ‚Ç$&í"¬Çí”‚˜V‰Ê˜Fñfñ6Fñˆ‰Fˆ7V÷VÁEfñWvW"ÜGF6Ü÷VÁBÁW&¬«¬""¬GF6Ü÷VÁBÊÊ÷R«¬$Fˆ7V÷VÁFÚ"íì∞¢&˜rÊVÊD6Üñ∆BÜ˜V‰'F‚ì∞¢VíÁW6W$∆W'DGF6Ü÷VÁG2ÊVÊD6Üñ∆Bá&˜rì∞¢“ì∞ß–†¶gVÊ7Fñˆ‚'Vñ∆DÊ˜Fñfñ6Fñˆ‰V÷ñ≈ñ∆ˆBáF&vWEW6W"¬÷W76vR¬Ê˜Fñfñ6Fñˆ‰ñBí∞¢6ˆÁ7B&V6óñVÁDV÷ñ¬“7G&ñÊráF&vWEW6W#ÚÊV÷ñ¬«¬""íÁG&ñ“Çì∞¢6ˆÁ7BW&¬“G∑vñÊF˜rÊ∆ˆ6Fñˆ‚Ê˜&ñvñÁ“G∑vñÊF˜rÊ∆ˆ6Fñˆ‚ÁFÜÊ÷W“6Üˆ÷V∞¢&WGW&‚∞¢FÛ¢&V6óñVÁDV÷ñ¬¿¢7V&¶V7C¢$ÁV˜fÊ˜Fñfñ6FÜW&"¿¢FWáC¢6ñÚG∑F&vWEW6W#ÚÊFó7∆îÊ÷R«¬'WFVÁFR'“≈∆Â∆Ê¬v÷÷ñÊó7G&F˜&RFíÜñÁfñFÚVÊÁV˜fÊ˜Fñfñ6•∆Â¬"G∂÷W76vW’¬%∆Â∆‰&í¬vW"fó7V∆óß¶&∆¢G∂W&«’∆Â∆‰îBÊ˜Fñfñ6¢G∂Ê˜Fñfñ6Fñˆ‰ñG÷¿¢áF÷√¢ ¢«‰6ñÚG∂W66TÖD‘¬áF&vWEW6W#ÚÊFó7∆îÊ÷R«¬'WFVÁFR"ó“√¬˜‡¢«‰¬v÷÷ñÊó7G&F˜&RFíÜñÁfñFÚVÊÁV˜fÊ˜Fñfñ67RÜW&„¬˜‡¢«„∆#‰÷W76vvñÛ£¬ˆ#‚G∂W66TÖD‘¬Ü÷W76vRó”¬˜‡¢«„∆á&Vc“"G∂W&«“"7Gñ∆S“&Fó7∆ì¶ñÊ∆ñÊR÷&∆ˆ6≥∑FFñÊs£ÇGÉ∂&6∂w&˜VÊC¢3CFVCÉ∂6ˆ∆˜#¢6ffc∑FWáB÷FV6˜&Fñˆ„¶ÊˆÊS∂&˜&FW"◊&FóW3£áÉ≤#‰6∆ñ66W"&ó&R¬v¬ˆ„¬˜‡¢«7Gñ∆S“&fˆÁB◊6ó¶S£'É∂6ˆ∆˜#¢3cCsCÜ#≤#‰îBÊ˜Fñfñ6¢G∂W66TÖD‘¬ÜÊ˜Fñfñ6Fñˆ‰ñBó”¬˜‡¢¿¢W&¬¿¢Ê˜Fñfñ6Fñˆ‰ñ@¢”∞ß–†¶7ñÊ2gVÊ7Fñˆ‚VWVTÊ˜Fñfñ6Fñˆ‰V÷ñ¬áF&vWEW6W"¬÷W76vR¬Ê˜Fñfñ6Fñˆ‰ñBí∞¢ñbÇF&vWEW6W#ÚÊV÷ñ¬í&WGW&‚≤VWVVC¢f«6R¬&V6ˆ„¢&÷ó76ñÊuˆV÷ñ¬"”∞¢6ˆÁ7Bñ∆ˆB“'Vñ∆DÊ˜Fñfñ6Fñˆ‰V÷ñ≈ñ∆ˆBáF&vWEW6W"¬÷W76vR¬Ê˜Fñfñ6Fñˆ‰ñBì∞¢vóBF"Ê6ˆ∆∆V7Fñˆ‚Ç&Ê˜Fñfñ6Fñˆ‰V÷ñ≈VWVR"íÊFBá∞¢‚‚Áñ∆ˆB¿¢7FGW3¢'VÊFñÊr"¿¢7&VFVDC¢fó&V&6RÊfó&W7F˜&R‰fñV∆Ef«VRÁ6W'fW%Fñ÷W7F◊Çí¿¢7&VFVD'ì¢7W'&VÁEW6W#ÚÊV÷ñ¬«¬7W'&VÁEW6W#ÚÁVñB«¬&F÷ñ‚ ¢“ì∞¢6ˆÁ7BvV&ÜˆˆµW&¬“7G&ñÊrávñÊF˜sÚ‰ÑU$Ù‰ıDîdî4DîÙÂÙT‘î≈ıtT$ÑÙÙ≤«¬∆ˆ6≈7F˜&vRÊvWDóFV“Ç&ÜW&Ê˜Fñfñ6Fñˆ‰V÷ñ≈vV&Üˆˆ≤"í«¬""íÁG&ñ“Çì∞¢ñbÇvV&ÜˆˆµW&¬í&WGW&‚≤VWVVC¢G'VR¬FV∆ófW&VC¢f«6R”∞¢G'í∞¢vóBfWF6ÖvóFÖFñ÷V˜WDÊE&WG'íávV&ÜˆˆµW&¬¬∞¢÷WFÜˆC¢%ı5B"¿¢ÜVFW'3¢≤$6ˆÁFVÁB’GóR#¢&∆ñ6Fñˆ‚ˆß6ˆ‚"“¿¢&ˆGì¢•4Ù‚Á7G&ñÊvñgíáñ∆ˆBê¢“¬∞¢Fñ÷V˜WD◊3¢‰UEtı$µÙDTdT≈EıDî‘TıUEÙ’2¿¢&WG&ñW3¢¢“ì∞¢&WGW&‚≤VWVVC¢G'VR¬FV∆ófW&VC¢G'VR”∞¢“6F6ÇÜW'&˜"í∞¢6ˆÁ6ˆ∆RÁv&‚Ç%vV&Üˆˆ≤V÷ñ¬Ê˜Fñfñ6Êˆ‚&vvóVÊvñ&ñ∆S¢"¬W'&˜"ì∞¢&WGW&‚≤VWVVC¢G'VR¬FV∆ófW&VC¢f«6R¬W'&˜"”∞¢–ß–†¶gVÊ7Fñˆ‚÷fñ∆TÊ÷Tf˜$Ê˜Fñfñ6Fñˆ‰GF6Ü÷VÁBÜfñ∆TÊ÷R“""í∞¢6ˆÁ7B6fTÊ÷R“7G&ñÊrÜfñ∆TÊ÷R«¬&∆∆VvFÚ"íÁ&W∆6RÇı«2≤ˆr¬%Ú"íÁ&W∆6RÇıµÂ«rÂ¬’“≤ˆr¬%Ú"ì∞¢&WGW&‚Ê˜Fñfñ6ÜRÚG¥FFRÊÊ˜rÇó’ÚG∑6fTÊ÷W÷∞ß–†¶gVÊ7Fñˆ‚6WDÊ˜Fñfñ6FñˆÂW∆ˆE7FFRÜñÂ&ˆw&W72í∞¢Ê˜Fñfñ6FñˆÂW∆ˆDñÂ&ˆw&W72“&ˆˆ∆V‚ÜñÂ&ˆw&W72ì∞¢ñbáVíÊÊ˜Fñfñ6Fñˆ‰6Ê6V≈W∆ˆD'F‚íVíÊÊ˜Fñfñ6Fñˆ‰6Ê6V≈W∆ˆD'F‚ÊFó6&∆VB“Ê˜Fñfñ6FñˆÂW∆ˆDñÂ&ˆw&W72«¬6‰÷ÊvTFFÇì∞¢ñbáVíÊÊ˜Fñfñ6FñˆÂ7V&÷óBíVíÊÊ˜Fñfñ6FñˆÂ7V&÷óBÊFó6&∆VB“Ê˜Fñfñ6FñˆÂW∆ˆDñÂ&ˆw&W72«¬6‰÷ÊvTFFÇì∞ß–†¶gVÊ7Fñˆ‚6Ê6VƒÊ˜Fñfñ6FñˆÂW∆ˆBÇí∞¢ñbÇÊ˜Fñfñ6FñˆÂW∆ˆD&˜'D6ˆÁG&ˆ∆∆W"í&WGW&„∞¢Ê˜Fñfñ6FñˆÂW∆ˆD&˜'D6ˆÁG&ˆ∆∆W"Ê&˜'BÇì∞¢Ê˜Fñfñ6FñˆÂW∆ˆD&˜'D6ˆÁG&ˆ∆∆W"“ÁV∆√∞¢6WDÊ˜Fñfñ6FñˆÂW∆ˆE7FFRÜf«6Rì∞¢ñbáVíÊÊ˜Fñfñ6Fñˆ‰fVVF&6≤íVíÊÊ˜Fñfñ6Fñˆ‰fVVF&6≤ÁFWáD6ˆÁFVÁB“$6&ñ6÷VÁFÚ∆∆VvFíÊÁV∆∆FÚ‚#∞ß–†¶7ñÊ2gVÊ7Fñˆ‚W∆ˆDÊ˜Fñfñ6Fñˆ‰GF6Ü÷VÁG2Üfñ∆W2“µ“¬˜FñˆÁ2“∑“í∞¢ñbÇfñ∆W2Ê∆VÊwFÇí&WGW&‚µ”∞¢ñbÇó46VÁG&ƒG&ófT6ˆÊfñwW&VBÇíbbG&ófT66W75Fˆ∂V‚í∞¢Fá&˜rÊWrW'&˜"Ç$6∆˜VB÷÷ñÊó7G&F˜&RÊˆ‚6ˆÊfñwW&FÚ‚6ˆÊfñwW&G&ófRF÷ñ‚W"∆∆Vv&RFˆ7V÷VÁFí∆∆RÊ˜Fñfñ6ÜR‚"ì∞¢–¢ñbÇG&ófU&W˜'G4fˆ∆FW$ñBívóBVÁ7W&TG&ófTfˆ∆FW'2Çì∞¢6ˆÁ7B≤6ñvÊ¬“ÁV∆¬¬ˆÂ&ˆw&W72“ÁV∆¬““˜FñˆÁ3∞¢∆WB6ˆ◊∆WFVB“∞¢6ˆÁ7BW∆ˆG2“vóB&ˆ÷ó6RÊ∆¬Üfñ∆W2Ê÷Ü7ñÊ2Üfñ∆Rí”‚∞¢ñbá6ñvÊ√ÚÊ&˜'FVBíFá&˜rÊWrDÙ‘WÜ6WFñˆ‚Ç%W∆ˆBÊÁV∆∆FÚ"¬$&˜'DW'&˜""ì∞¢6ˆÁ7BW∆ˆB“vóBW∆ˆD&∆ˆ%FÙG&ófRÄ¢fñ∆R¿¢÷fñ∆TÊ÷Tf˜$Ê˜Fñfñ6Fñˆ‰GF6Ü÷VÁBÜfñ∆RÊÊ÷R«¬&∆∆VvFÚ"í¿¢fñ∆RÁGóR«¬&∆ñ6Fñˆ‚ˆˆ7FWB◊7G&V“"¿¢G&ófU&W˜'G4fˆ∆FW$ñB¿¢≤6ñvÊ¬–¢ì∞¢6ˆ◊∆WFVB≥“∞¢ñbáGóVˆbˆÂ&ˆw&W72””“&gVÊ7Fñˆ‚"íˆÂ&ˆw&W72Ü6ˆ◊∆WFVB¬fñ∆W2Ê∆VÊwFÇ¬fñ∆RÊÊ÷R«¬&∆∆VvFÚ"ì∞¢&WGW&‚∞¢Ê÷S¢fñ∆RÊÊ÷R«¬&∆∆VvFÚ"¿¢GóS¢fñ∆RÁGóR«¬&∆ñ6Fñˆ‚ˆˆ7FWB◊7G&V“"¿¢6ó¶S¢ÁV÷&W"Üfñ∆RÁ6ó¶R«¬í¿¢W&√¢W∆ˆBÁvV%fñWt∆ñÊ≤«¬W∆ˆBÊFó&V7EW&¬«¬""¿¢fñ∆TñC¢W∆ˆBÊfñ∆TñB«¬" ¢”∞¢“íì∞¢&WGW&‚W∆ˆG3∞ß–†¶gVÊ7Fñˆ‚'Vñ∆DFˆ7V÷VÁEfñWvW%W&¬á&uW&¬“""í∞¢6ˆÁ7BW&¬“7G&ñÊrá&uW&¬«¬""íÁG&ñ“Çì∞¢ñbÇW&¬í&WGW&‚"#∞¢ñbÇˆFˆ75¬Êvˆˆv∆U¬Ê6ˆ’¬˜7&VG6ÜVWG2ˆíÁFW7BáW&¬íí&WGW&‚W&√∞¢ñbÇˆG&ófU¬Êvˆˆv∆U¬Ê6ˆ“ˆíÁFW7BáW&¬íí∞¢&WGW&‚áGG3¢ÚˆFˆ72Êvˆˆv∆RÊ6ˆ“ˆwfñWsˆV÷&VFFVC”gW&√“G∂VÊ6ˆFUU$î6ˆ◊ˆÊVÁBáW&¬ó÷∞¢–¢&WGW&‚áGG3¢ÚˆFˆ72Êvˆˆv∆RÊ6ˆ“ˆwfñWsˆV÷&VFFVC”gW&√“G∂VÊ6ˆFUU$î6ˆ◊ˆÊVÁBáW&¬ó÷∞ß–†¶gVÊ7Fñˆ‚˜V‰Ê˜Fñfñ6Fñˆ‰Fˆ7V÷VÁEfñWvW"á&uW&¬¬FóF∆R“$Fˆ7V÷VÁFÚ"í∞¢6ˆÁ7BfñWvW%W&¬“'Vñ∆DFˆ7V÷VÁEfñWvW%W&¬á&uW&¬ì∞¢ñbÇfñWvW%W&¬í&WGW&„∞¢ñbáVíÊÊ˜Fñfñ6Fñˆ‰Fˆ5fñWvW%FóF∆RíVíÊÊ˜Fñfñ6Fñˆ‰Fˆ5fñWvW%FóF∆RÁFWáD6ˆÁFVÁB“FóF∆S∞¢ñbáVíÊÊ˜Fñfñ6Fñˆ‰Fˆ5fñWvW$g&÷RíVíÊÊ˜Fñfñ6Fñˆ‰Fˆ5fñWvW$g&÷RÁ7&2“fñWvW%W&√∞¢VíÊÊ˜Fñfñ6Fñˆ‰Fˆ5fñWvW$÷ˆF√ÚÊ6∆74∆ó7BÁ&V÷˜fRÇ&ÜñFFV‚"ì∞¢VíÊÊ˜Fñfñ6Fñˆ‰Fˆ5fñWvW$÷ˆF√ÚÁ6WDGG&ñ'WFRÇ&&ñ÷ÜñFFV‚"¬&f«6R"ì∞ß–†¶gVÊ7Fñˆ‚6∆˜6TÊ˜Fñfñ6Fñˆ‰Fˆ7V÷VÁEfñWvW"Çí∞¢ñbáVíÊÊ˜Fñfñ6Fñˆ‰Fˆ5fñWvW$g&÷RíVíÊÊ˜Fñfñ6Fñˆ‰Fˆ5fñWvW$g&÷RÁ7&2“"#∞¢VíÊÊ˜Fñfñ6Fñˆ‰Fˆ5fñWvW$÷ˆF√ÚÊ6∆74∆ó7BÊFBÇ&ÜñFFV‚"ì∞¢VíÊÊ˜Fñfñ6Fñˆ‰Fˆ5fñWvW$÷ˆF√ÚÁ6WDGG&ñ'WFRÇ&&ñ÷ÜñFFV‚"¬'G'VR"ì∞ß–†¶7ñÊ2gVÊ7Fñˆ‚7&VFUW6W$Ê˜Fñfñ6Fñˆ‚ÜWfVÁBí∞¢WfVÁBÁ&WfVÁDFVfV«BÇì∞¢ñbÜÊ˜Fñfñ6FñˆÂW∆ˆDñÂ&ˆw&W72í∞¢ñbáVíÊÊ˜Fñfñ6Fñˆ‰fVVF&6≤íVíÊÊ˜Fñfñ6Fñˆ‰fVVF&6≤ÁFWáD6ˆÁFVÁB“$6&ñ6÷VÁFÚvú:ñ‚6˜'6Ú‚‚‚#∞¢&WGW&„∞¢–¢ñbÇ7W'&VÁEW6W"«¬6‰÷ÊvTFFÇíí∞¢ñbáVíÊÊ˜Fñfñ6Fñˆ‰fVVF&6≤íVíÊÊ˜Fñfñ6Fñˆ‰fVVF&6≤ÁFWáD6ˆÁFVÁB“%6ˆ∆Úv∆íF÷ñ‚˜76ˆÊÚñÁfñ&Rgfó6í‚#∞¢&WGW&„∞¢–¢6ˆÁ7B6VÊEFÙ∆≈&Vvó7FW&VB“&ˆˆ∆V‚áVíÊÊ˜Fñfñ6FñˆÂ6VÊD∆≈Fˆvv∆SÚÊ6ÜV6∂VBì∞¢6ˆÁ7BF&vWEW6W$ñG2“6VÊEFÙ∆≈&Vvó7FW&V@¢Úµ–¢¢'&íÊg&ˆ“áVíÊÊ˜Fñfñ6FñˆÂW6W%6V∆V7CÚÁ6V∆V7FVD˜FñˆÁ2«¬µ“íÊ÷ÇÜ˜Bí”‚7G&ñÊrÜ˜BÁf«VR«¬""íÁG&ñ“ÇííÊfñ«FW"Ñ&ˆˆ∆V‚ì∞¢6ˆÁ7BFóF∆R“7G&ñÊráVíÊÊ˜Fñfñ6FñˆÂFóF∆SÚÁf«VR«¬""íÁG&ñ“Çì∞¢6ˆÁ7B÷W76vR“7G&ñÊráVíÊÊ˜Fñfñ6Fñˆ‰÷W76vSÚÁf«VR«¬""íÁG&ñ“Çì∞¢6ˆÁ7B66ÜVGV∆VDFFT∂Wí“7G&ñÊráVíÊÊ˜Fñfñ6Fñˆ‰FFSÚÁf«VR«¬""íÁG&ñ“Çì∞¢6ˆÁ7Bfñ∆W2“'&íÊg&ˆ“áVíÊÊ˜Fñfñ6Fñˆ‰GF6Ü÷VÁG3ÚÊfñ∆W2«¬µ“ì∞¢6ˆÁ7BF˜F≈6ó¶T÷"“fñ∆W2Á&VGV6RÇá7V“¬fñ∆Rí”‚7V“≤ÁV÷&W"Üfñ∆RÁ6ó¶R«¬í¬íÚÉ#B¢#Bì∞¢ñbáF˜F≈6ó¶T÷"‚Éí∞¢ñbáVíÊÊ˜Fñfñ6Fñˆ‰fVVF&6≤íVíÊÊ˜Fñfñ6Fñˆ‰fVVF&6≤ÁFWáD6ˆÁFVÁB“$∆∆VvFíG&˜Úw&ÊFíÉ„É‘"í‚&ñGV6íñ¬W6ÚW"fV∆ˆ6óß¶&Rñ¬6&ñ6÷VÁFÚ‚#∞¢&WGW&„∞¢–¢ñbÇFóF∆R«¬÷W76vRí∞¢ñbáVíÊÊ˜Fñfñ6Fñˆ‰fVVF&6≤íVíÊÊ˜Fñfñ6Fñˆ‰fVVF&6≤ÁFWáD6ˆÁFVÁB“$ñÁ6W&ó66íFóFˆ∆ÚRFW7FÚÊ˜Fñfñ6‚#∞¢&WGW&„∞¢–¢ñbÇ6VÊEFÙ∆≈&Vvó7FW&VBbbF&vWEW6W$ñG2Ê∆VÊwFÇí∞¢ñbáVíÊÊ˜Fñfñ6Fñˆ‰fVVF&6≤íVíÊÊ˜Fñfñ6Fñˆ‰fVVF&6≤ÁFWáD6ˆÁFVÁB“%6V∆W¶ñˆÊ∆÷VÊÚV‚WFVÁFRÚ66Vv∆íñÁfñÚGWGFí‚#∞¢&WGW&„∞¢–¢Ê˜Fñfñ6FñˆÂW∆ˆD&˜'D6ˆÁG&ˆ∆∆W"“ÊWr&˜'D6ˆÁG&ˆ∆∆W"Çì∞¢6WDÊ˜Fñfñ6FñˆÂW∆ˆE7FFRáG'VRì∞¢G'í∞¢ñbáVíÊÊ˜Fñfñ6Fñˆ‰fVVF&6≤íVíÊÊ˜Fñfñ6Fñˆ‰fVVF&6≤ÁFWáD6ˆÁFVÁB“fñ∆W2Ê∆VÊwFÇÚ6&ñ6÷VÁFÚ∆∆VvFíÉÚG∂fñ∆W2Ê∆VÊwFá“í‚‚Ê¢$ñÁfñÚÊ˜Fñfñ6‚‚‚#∞¢6ˆÁ7BGF6Ü÷VÁG2“vóBW∆ˆDÊ˜Fñfñ6Fñˆ‰GF6Ü÷VÁG2Üfñ∆W2¬∞¢6ñvÊ√¢Ê˜Fñfñ6FñˆÂW∆ˆD&˜'D6ˆÁG&ˆ∆∆W"Á6ñvÊ¬¿¢ˆÂ&ˆw&W73¢Ü6ˆ◊∆WFVB¬F˜F¬í”‚∞¢ñbáVíÊÊ˜Fñfñ6Fñˆ‰fVVF&6≤íVíÊÊ˜Fñfñ6Fñˆ‰fVVF&6≤ÁFWáD6ˆÁFVÁB“6&ñ6÷VÁFÚ∆∆VvFíÇG∂6ˆ◊∆WFVG“ÚG∑F˜F«“í‚‚Ê∞¢–¢“ì∞¢ñbáVíÊÊ˜Fñfñ6Fñˆ‰fVVF&6≤íVíÊÊ˜Fñfñ6Fñˆ‰fVVF&6≤ÁFWáD6ˆÁFVÁB“%6«fFvvñÚÊ˜Fñfñ6‚‚‚#∞¢6ˆÁ7B7&VFVDFFT∂Wí“vWDFFT∂Wîg&ˆ‘∆ˆ6ƒFFRÜÊWrFFRÇíì∞¢6ˆÁ7BÊ˜Fñfñ6FñˆÂ&Vb“vóBF"Ê6ˆ∆∆V7Fñˆ‚Ç'W6W$∆W'G2"íÊFBá∞¢FóF∆R¿¢6VÊEFÙ∆≈&Vvó7FW&VB¿¢F&vWEW6W$ñG2¿¢÷W76vR¿¢GF6Ü÷VÁG2¿¢66ÜVGV∆VDFFT∂Wì¢66ÜVGV∆VDFFT∂Wí«¬""¿¢7&VFVDFFT∂Wí¿¢7FGW3¢66ÜVGV∆VDFFT∂Wíbb66ÜVGV∆VDFFT∂Wí‚7&VFVDFFT∂WíÚ'66ÜVGV∆VB"¢&7FófR"¿¢7&VFVDC¢fó&V&6RÊfó&W7F˜&R‰fñV∆Ef«VRÁ6W'fW%Fñ÷W7F◊Çí¿¢7&VFVD'ì¢7W'&VÁEW6W"ÊV÷ñ¬«¬7W'&VÁEW6W"ÁVñB«¬&F÷ñ‚"¿¢6∂Ê˜v∆VFvVEW6W'3¢ ¢“ì∞¢6ˆÁ7BF&vWEW6W'2“6VÊEFÙ∆≈&Vvó7FW&V@¢Ú∆Ff˜&’W6W'2Êfñ«FW"ÇáW6W"í”‚F÷ñ‰V÷ñ«2ÊÜ2ÜÊ˜&÷∆ó¶TV÷ñ¬áW6W"ÊV÷ñ¬ííê¢¢∆Ff˜&’W6W'2Êfñ«FW"ÇáW6W"í”‚F&vWEW6W$ñG2ÊñÊ6«VFW2áW6W"ÊñBíì∞¢6ˆÁ7BV÷ñ≈&W7V«G2“vóB&ˆ÷ó6RÊ∆¬áF&vWEW6W'2Ê÷ÇáF&vWEW6W"í”‚VWVTÊ˜Fñfñ6Fñˆ‰V÷ñ¬áF&vWEW6W"¬÷W76vR¬Ê˜Fñfñ6FñˆÂ&VbÊñBííì∞¢6ˆÁ7BÜ5VWVVB“V÷ñ≈&W7V«G2Á6ˆ÷RÇá&W7V«Bí”‚&W7V«CÚÁVWVVBì∞¢ñbáVíÊÊ˜Fñfñ6Fñˆ‰f˜&“íVíÊÊ˜Fñfñ6Fñˆ‰f˜&“Á&W6WBÇì∞¢ˆ‰Ê˜Fñfñ6FñˆÂ6VÊD∆ƒ6ÜÊvRÇì∞¢ñbáVíÊÊ˜Fñfñ6Fñˆ‰fVVF&6≤í∞¢VíÊÊ˜Fñfñ6Fñˆ‰fVVF&6≤ÁFWáD6ˆÁFVÁB“Ü5VWVV@¢Ú$Ê˜Fñfñ66«fF‚V÷ñ¬÷W76Rñ‚6ˆFW"íFW7FñÊF&í6ˆ‚V÷ñ¬‚ ¢¢$Ê˜Fñfñ66«fF‚ÊW77VÊV÷ñ¬ñÁfñFÜFW7FñÊF&í6VÁ¶V÷ñ¬í‚#∞¢–¢“6F6ÇÜW'&˜"í∞¢ñbÜW'&˜#ÚÊÊ÷R””“$&˜'DW'&˜""í∞¢ñbáVíÊÊ˜Fñfñ6Fñˆ‰fVVF&6≤íVíÊÊ˜Fñfñ6Fñˆ‰fVVF&6≤ÁFWáD6ˆÁFVÁB“$6&ñ6÷VÁFÚÊÁV∆∆FÚ‚#∞¢&WGW&„∞¢–¢6ˆÁ6ˆ∆RÊW'&˜"Ç$W'&˜&RñÁfñÚgfó6ÚWFVÁFS¢"¬W'&˜"ì∞¢ñbáVíÊÊ˜Fñfñ6Fñˆ‰fVVF&6≤íVíÊÊ˜Fñfñ6Fñˆ‰fVVF&6≤ÁFWáD6ˆÁFVÁB“$W'&˜&RGW&ÁFRñ¬6«fFvvñÚFV∆¬vgfó6Ú‚#∞¢“fñÊ∆«í∞¢Ê˜Fñfñ6FñˆÂW∆ˆD&˜'D6ˆÁG&ˆ∆∆W"“ÁV∆√∞¢6WDÊ˜Fñfñ6FñˆÂW∆ˆE7FFRÜf«6Rì∞¢–ß–†¶7ñÊ2gVÊ7Fñˆ‚FV∆WFUW6W$Ê˜Fñfñ6Fñˆ‚ÜÊ˜Fñfñ6Fñˆ‰ñBí∞¢ñbÇ7W'&VÁEW6W"«¬6‰÷ÊvTFFÇí«¬Ê˜Fñfñ6Fñˆ‰ñBí&WGW&„∞¢6ˆÁ7B6ˆÊfó&÷VB“vñÊF˜rÊ6ˆÊfó&“Ç$V∆ñ÷ñÊ&RVW7FÊ˜Fñfñ6Ú"ì∞¢ñbÇ6ˆÊfó&÷VBí&WGW&„∞¢vóBF"Ê6ˆ∆∆V7Fñˆ‚Ç'W6W$∆W'G2"íÊFˆ2ÜÊ˜Fñfñ6Fñˆ‰ñBíÊFV∆WFRÇì∞ß–†¶gVÊ7Fñˆ‚7V'67&ñ&UW6W$∆W'G2Çí∞¢ñbáVÁ7V'67&ñ&UW6W$∆W'G2íVÁ7V'67&ñ&UW6W$∆W'G2Çì∞¢ñbÇ7W'&VÁEW6W"í&WGW&„∞¢6ˆÁ7B∆Vv7ï7VG&∆W'E6˜W&6R“≤'7VG&"¬&gfó6Ú%“Ê¶ˆñ‚Ç"“"ì∞¢ñbÜ6‰÷ÊvTFFÇíí∞¢VÁ7V'67&ñ&UW6W$∆W'G2“F"Ê6ˆ∆∆V7Fñˆ‚Ç'W6W$∆W'G2"ê¢Ê˜&FW$'íÇ&7&VFVDB"¬&FW62"ê¢Ê∆ñ÷óBÉ#ê¢ÊˆÂ6Ê6Ü˜BÇá6Ê6Ü˜Bí”‚∞¢W6W$∆W'G2“6Ê6Ü˜BÊFˆ70¢Ê÷ÇÜFˆ2í”‚á≤ñC¢Fˆ2ÊñB¬‚‚ÊFˆ2ÊFFÇí“íê¢Êfñ«FW"ÇÜóFV“í”‚óFV“Á6˜W&6R”“∆Vv7ï7VG&∆W'E6˜W&6Rì∞¢&VÊFW$Ê˜Fñfñ6FñˆÁ4∆ó7BÇì∞¢&VÊFW$Ê˜Fñfñ6Fñˆ‰6∆VÊF"Çì∞¢&VÊFW%FˆFï7V÷÷'íÇì∞¢“¬ÜW'&˜"í”‚∞¢6ˆÁ6ˆ∆RÊW'&˜"Ç$W'&˜&R6&ñ6÷VÁFÚÊ˜Fñfñ6ÜRF÷ñ„¢"¬W'&˜"ì∞¢“ì∞¢&WGW&„∞¢–¢VÁ7V'67&ñ&UW6W$∆W'G2“F"Ê6ˆ∆∆V7Fñˆ‚Ç'W6W$∆W'G2"ê¢Ê˜&FW$'íÇ&7&VFVDB"¬&FW62"ê¢Ê∆ñ÷óBÉê¢ÊˆÂ6Ê6Ü˜BÇá6Ê6Ü˜Bí”‚∞¢6ˆÁ7BFˆFî∂Wí“vWDFFT∂Wîg&ˆ‘∆ˆ6ƒFFRÜÊWrFFRÇíì∞¢W6W$∆W'G2“6Ê6Ü˜BÊFˆ70¢Ê÷ÇÜFˆ2í”‚á≤ñC¢Fˆ2ÊñB¬‚‚ÊFˆ2ÊFFÇí“íê¢Êfñ«FW"ÇÜóFV“í”‚∞¢ñbÜóFV“Á6˜W&6R””“∆Vv7ï7VG&∆W'E6˜W&6Rí&WGW&‚f«6S∞¢6ˆÁ7B&V∆ˆÊw5FıW6W"“ó4Ê˜Fñfñ6Fñˆ‰f˜$7W'&VÁEW6W"ÜóFV“ì∞¢ñbÇ&V∆ˆÊw5FıW6W"í&WGW&‚f«6S∞¢ñbÑ&ˆˆ∆V‚ÜóFV”ÚÊFó6÷ó76VD'ïW6W$ñG3ÚÂ∂7W'&VÁEW6W"ÁVñE“íí&WGW&‚f«6S∞¢6ˆÁ7BFFT∂Wí“7G&ñÊrÜóFV“Á66ÜVGV∆VDFFT∂Wí«¬""íÁG&ñ“Çì∞¢&WGW&‚óFV“Á6˜W&6R””“&6∆VÊF"÷'6VÊ6R"«¬FFT∂Wí«¬FFT∂Wí√“FˆFî∂Wì∞¢“ê¢Á6˜'BÇÜ¬"í”‚∞¢6ˆÁ7B◊2“ÚÊ7&VFVDBbbGóVˆbÊ7&VFVDBÁFÙ÷ñ∆∆ó2””“&gVÊ7Fñˆ‚"ÚÊ7&VFVDBÁFÙ÷ñ∆∆ó2Çí¢∞¢6ˆÁ7B$◊2“#ÚÊ7&VFVDBbbGóVˆb"Ê7&VFVDBÁFÙ÷ñ∆∆ó2””“&gVÊ7Fñˆ‚"Ú"Ê7&VFVDBÁFÙ÷ñ∆∆ó2Çí¢∞¢&WGW&‚$◊2“◊3∞¢“ì∞¢vñÊF˜r‰ÜW&Ê˜Fñfñ6FñˆÂ&VFW#ÚÊ&6ÜófSÚ‚áW6W$∆W'G2Ê÷ÇÜóFV“í”‚á∞¢ñC¢óFV“ÊñB¿¢FóF∆S¢óFV“ÁFóF∆R«¬$gfó6Úñ◊˜'FÁFR"¿¢&ˆGì¢óFV“Ê÷W76vR«¬""¿¢FW7FñÊFñˆ„¢&Üˆ÷R"¿¢&V6VófVDC¢óFV“Ê7&VFVDCÚÁFÙ÷ñ∆∆ó3Ú‚Çí«¬FFRÊÊ˜rÇê¢“ííì∞¢÷ñ&U6Ü˜uW6W$∆W'BÇì∞¢&VÊFW%FˆFï7V÷÷'íÇì∞¢“¬ÜW'&˜"í”‚∞¢6ˆÁ6ˆ∆RÊW'&˜"Ç$W'&˜&R6&ñ6÷VÁFÚÊ˜Fñfñ6ÜRWFVÁFS¢"¬W'&˜"ì∞¢“ì∞ß–†¶gVÊ7Fñˆ‚7F˜W6W$∆W'G57V'67&óFñˆ‚Çí∞¢ñbáVÁ7V'67&ñ&UW6W$∆W'G2í∞¢VÁ7V'67&ñ&UW6W$∆W'G2Çì∞¢VÁ7V'67&ñ&UW6W$∆W'G2“ÁV∆√∞¢–¢W6W$∆W'G2“µ”∞¢7FófUW6W$∆W'B“ÁV∆√∞¢&VÊFW$Ê˜Fñfñ6FñˆÁ4∆ó7BÇì∞ß–†¶gVÊ7Fñˆ‚÷ñ&U6Ü˜uW6W$∆W'BÇí∞¢ñbÇ7W'&VÁEW6W"«¬6‰÷ÊvTFFÇíí&WGW&„∞¢6ˆÁ7Bfó'7EVÊFñÊr“W6W$∆W'G2ÊfñÊBÇÜ∆W'DóFV“í”‚&ˆˆ∆V‚Ü∆W'DóFV”ÚÊ6¥'ïW6W$ñG3ÚÂ∂7W'&VÁEW6W"ÁVñE“íì∞¢ñbÇfó'7EVÊFñÊrí∞¢6∆˜6UW6W$∆W'D÷ˆF¬Çì∞¢&WGW&„∞¢–¢6∆˜6U6ñFT÷VÁRÇì∞¢6∆˜6T÷ÊvV÷VÁEÊV¬Çì∞¢7FófUW6W$∆W'B“fó'7EVÊFñÊs∞¢ñbáVíÁW6W$∆W'EFWáBíVíÁW6W$∆W'EFWáBÁFWáD6ˆÁFVÁB“G∂fó'7EVÊFñÊrÁFóF∆RÚG∂fó'7EVÊFñÊrÁFóF∆W’∆Â∆Ê¢"'“G∂fó'7EVÊFñÊrÊ÷W76vR«¬"'÷∞¢&VÊFW$7FófUW6W$∆W'DGF6Ü÷VÁG2Çì∞¢VíÁW6W$∆W'D÷ˆF√ÚÊ6∆74∆ó7BÁ&V÷˜fRÇ&ÜñFFV‚"ì∞¢VíÁW6W$∆W'D÷ˆF√ÚÁ6WDGG&ñ'WFRÇ&&ñ÷ÜñFFV‚"¬&f«6R"ì∞ß–†¶gVÊ7Fñˆ‚6∆˜6UW6W$∆W'D÷ˆF¬Çí∞¢ñbáVíÁW6W$∆W'DGF6Ü÷VÁG2íVíÁW6W$∆W'DGF6Ü÷VÁG2ÊñÊÊW$ÖD‘¬“"#∞¢VíÁW6W$∆W'D÷ˆF√ÚÊ6∆74∆ó7BÊFBÇ&ÜñFFV‚"ì∞¢VíÁW6W$∆W'D÷ˆF√ÚÁ6WDGG&ñ'WFRÇ&&ñ÷ÜñFFV‚"¬'G'VR"ì∞ß–†¶7ñÊ2gVÊ7Fñˆ‚6∂Ê˜v∆VFvT7FófUW6W$∆W'BÇí∞¢ñbÇ7W'&VÁEW6W"«¬7FófUW6W$∆W'CÚÊñBí&WGW&„∞¢6ˆÁ7B6∂Ê˜v∆VFvV÷VÁDñB“G∂7FófUW6W$∆W'BÊñG’ıÚG∂7W'&VÁEW6W"ÁVñG÷∞¢6ˆÁ7B6∂Ê˜v∆VFvV÷VÁE&Vb“F"Ê6ˆ∆∆V7Fñˆ‚Ç'W6W$∆W'D6∂Ê˜v∆VFvV÷VÁG2"íÊFˆ2Ü6∂Ê˜v∆VFvV÷VÁDñBì∞¢6ˆÁ7BWÜó7FñÊr“vóB6∂Ê˜v∆VFvV÷VÁE&VbÊvWBÇì∞¢6ˆÁ7BÊ˜r“ÊWrFFRÇì∞¢ñbÇWÜó7FñÊrÊWÜó7G2í∞¢vóB6∂Ê˜v∆VFvV÷VÁE&VbÁ6WBá∞¢Ê˜Fñfñ6Fñˆ‰ñC¢7FófUW6W$∆W'BÊñB¿¢W6W$ñC¢7W'&VÁEW6W"ÁVñB«¬""¿¢W6W$Ê÷S¢7W'&VÁEW6W"ÊFó7∆îÊ÷R«¬7W'&VÁEW6W"ÊV÷ñ¬«¬%WFVÁFR"¿¢6∂Ê˜v∆VFvVDC¢fó&V&6RÊfó&W7F˜&R‰fñV∆Ef«VRÁ6W'fW%Fñ÷W7F◊Çí¿¢6∂Ê˜v∆VFvVDFFT∂Wì¢vWDFFT∂Wîg&ˆ‘∆ˆ6ƒFFRÜÊ˜rê¢“¬≤÷W&vS¢G'VR“ì∞¢–¢vóBF"Ê6ˆ∆∆V7Fñˆ‚Ç'W6W$∆W'G2"íÊFˆ2Ü7FófUW6W$∆W'BÊñBíÁ6WBá∞¢∂6¥'ïW6W$ñG2‚G∂7W'&VÁEW6W"ÁVñG÷”¢fó&V&6RÊfó&W7F˜&R‰fñV∆Ef«VRÁ6W'fW%Fñ÷W7F◊Çí¿¢∂Fó6÷ó76VD'ïW6W$ñG2‚G∂7W'&VÁEW6W"ÁVñG÷”¢fó&V&6RÊfó&W7F˜&R‰fñV∆Ef«VRÁ6W'fW%Fñ÷W7F◊Çí¿¢∆7D6∂Ê˜v∆VFvVDC¢fó&V&6RÊfó&W7F˜&R‰fñV∆Ef«VRÁ6W'fW%Fñ÷W7F◊Çê¢“¬≤÷W&vS¢G'VR“ì∞¢ñbÇWÜó7FñÊrÊWÜó7G2ívóB6VÊDÊ˜Fñfñ6Fñˆ‰6µFÙF÷ñÁ2Ü7FófUW6W$∆W'B¬Ê˜rì∞¢7FófUW6W$∆W'B“ÁV∆√∞¢6∆˜6UW6W$∆W'D÷ˆF¬Çì∞ß–†¶7ñÊ2gVÊ7Fñˆ‚6VÊDÊ˜Fñfñ6Fñˆ‰6µFÙF÷ñÁ2Ü∆W'DóFV“¬6¥FFR“ÊWrFFRÇíí∞¢6ˆÁ7BF÷ñÂW6W'2“∆Ff˜&’W6W'2Êfñ«FW"ÇáW6W"í”‚F÷ñ‰V÷ñ«2ÊÜ2ÜÊ˜&÷∆ó¶TV÷ñ¬áW6W"ÊV÷ñ¬ííì∞¢ñbÇF÷ñÂW6W'2Ê∆VÊwFÇí&WGW&„∞¢6ˆÁ7BvÜV‰∆&V¬“6¥FFRÁFÙ∆ˆ6∆U7G&ñÊrÇ&óB‘ïB"ì∞¢6ˆÁ7BFWáB“)»R‰ıDîdî44Ù‰dU$‘D∆‰Œ(	óWFVÁFRG∂7W'&VÁEW6W#ÚÊFó7∆îÊ÷R«¬7W'&VÁEW6W#ÚÊV÷ñ¬«¬%WFVÁFR'“Ü&V◊WFÚ(	ƒÙ≤¬ÑÚ4ïD˛(	’∆‰Ê˜Fñfñ6¢G∂∆W'DóFV”ÚÁFóF∆R«¬$Ê˜Fñfñ6'’∆‰FFÙ˜&¢G∑vÜV‰∆&V«÷∞¢vóB&ˆ÷ó6RÊ∆¬ÜF÷ñÂW6W'2Ê÷ÇÜF÷ñÂW6W"í”‚F"Ê6ˆ∆∆V7Fñˆ‚Ç&6ÜD÷W76vW2"íÊFBá∞¢GóS¢'FWáB"¿¢FWáB¿¢&V6óñVÁDñC¢F÷ñÂW6W"ÊñB¿¢6VÊFW$ñC¢7W'&VÁEW6W#ÚÁVñB«¬'7ó7FV“"¿¢6VÊFW$Ê÷S¢%6ó7FV÷Ê˜Fñfñ6ÜR"¿¢6VÊFW$V÷ñ√¢7W'&VÁEW6W#ÚÊV÷ñ¬«¬""¿¢∂ñÊC¢'7ó7FV“"¿¢÷WFFF¢∞¢GóS¢&Ê˜Fñfñ6FñˆÂˆ6≤"¿¢Ê˜Fñfñ6Fñˆ‰ñC¢∆W'DóFV”ÚÊñB«¬""¿¢Ê˜Fñfñ6FñˆÂFóF∆S¢∆W'DóFV”ÚÁFóF∆R«¬""¿¢6∂Ê˜v∆VFvVD'ïW6W$ñC¢7W'&VÁEW6W#ÚÁVñB«¬""¿¢6∂Ê˜v∆VFvVD'ïW6W$Ê÷S¢7W'&VÁEW6W#ÚÊFó7∆îÊ÷R«¬7W'&VÁEW6W#ÚÊV÷ñ¬«¬%WFVÁFR"¿¢6∂Ê˜v∆VFvVDC¢6¥FFRÁFÙï4ı7G&ñÊrÇê¢“¿¢7&VFVDC¢fó&V&6RÊfó&W7F˜&R‰fñV∆Ef«VRÁ6W'fW%Fñ÷W7F◊Çê¢“ííì∞ß–†¶gVÊ7Fñˆ‚˜7GˆÊT7FófUW6W$∆W'BÇí∞¢7FófUW6W$∆W'B“ÁV∆√∞¢6∆˜6UW6W$∆W'D÷ˆF¬Çì∞ß–†¶gVÊ7Fñˆ‚˜V‰Ê˜Fñfñ6Fñˆ‰6∆VÊF%fñWrÇí∞¢ñbÇVíÊÊ˜Fñfñ6Fñˆ‰6∆VÊF%fñWr«¬VíÊÊ˜Fñfñ6Fñˆ‰÷ñÂfñWrí&WGW&„∞¢VíÊÊ˜Fñfñ6Fñˆ‰÷ñÂfñWrÊ6∆74∆ó7BÊFBÇ&ÜñFFV‚"ì∞¢VíÊÊ˜Fñfñ6Fñˆ‰6∆VÊF%fñWrÊ6∆74∆ó7BÁ&V÷˜fRÇ&ÜñFFV‚"ì∞¢&VÊFW$Ê˜Fñfñ6Fñˆ‰6∆VÊF"Çì∞ß–†¶gVÊ7Fñˆ‚6∆˜6TÊ˜Fñfñ6Fñˆ‰6∆VÊF%fñWrÇí∞¢ñbÇVíÊÊ˜Fñfñ6Fñˆ‰6∆VÊF%fñWr«¬VíÊÊ˜Fñfñ6Fñˆ‰÷ñÂfñWrí&WGW&„∞¢VíÊÊ˜Fñfñ6Fñˆ‰6∆VÊF%fñWrÊ6∆74∆ó7BÊFBÇ&ÜñFFV‚"ì∞¢VíÊÊ˜Fñfñ6Fñˆ‰÷ñÂfñWrÊ6∆74∆ó7BÁ&V÷˜fRÇ&ÜñFFV‚"ì∞ß–†¶gVÊ7Fñˆ‚÷˜fTÊ˜Fñfñ6Fñˆ‰6∆VÊF$÷ˆÁFÇÜˆfg6WBí∞¢Ê˜Fñfñ6Fñˆ‰6∆VÊF$7W'6˜"“ÊWrFFRÜÊ˜Fñfñ6Fñˆ‰6∆VÊF$7W'6˜"ÊvWDgV∆≈ñV"Çí¬Ê˜Fñfñ6Fñˆ‰6∆VÊF$7W'6˜"ÊvWD÷ˆÁFÇÇí≤ˆfg6WB¬ì∞¢&VÊFW$Ê˜Fñfñ6Fñˆ‰6∆VÊF"Çì∞ß–†¶gVÊ7Fñˆ‚&VÊFW$Ê˜Fñfñ6Fñˆ‰6∆VÊF"Çí∞¢ñbÇVíÊÊ˜Fñfñ6Fñˆ‰6∆VÊF$w&ñB«¬VíÊÊ˜Fñfñ6Fñˆ‰6∆VÊF$÷ˆÁFÑ∆&V¬í&WGW&„∞¢6ˆÁ7B÷ˆÁFÖ7F'B“ÊWrFFRÜÊ˜Fñfñ6Fñˆ‰6∆VÊF$7W'6˜"ÊvWDgV∆≈ñV"Çí¬Ê˜Fñfñ6Fñˆ‰6∆VÊF$7W'6˜"ÊvWD÷ˆÁFÇÇí¬ì∞¢6ˆÁ7B÷ˆÁFÑ∆&V¬“÷ˆÁFÖ7F'BÁFÙ∆ˆ6∆TFFU7G&ñÊrÇ&óB‘ïB"¬≤÷ˆÁFÉ¢&∆ˆÊr"¬ñV#¢&ÁV÷W&ñ2"“ì∞¢VíÊÊ˜Fñfñ6Fñˆ‰6∆VÊF$÷ˆÁFÑ∆&V¬ÁFWáD6ˆÁFVÁB“÷ˆÁFÑ∆&V¬Ê6Ü$BÉíÁFıWW$66RÇí≤÷ˆÁFÑ∆&V¬Á6∆ñ6RÉì∞¢6ˆÁ7BvVV∂Fî∆&V«2“≤$«V‚"¬$÷""¬$÷W""¬$vñÚ"¬%fV‚"¬%6""¬$Fˆ“%”∞¢6ˆÁ7BFî÷“ÊWr÷Çì∞¢W6W$∆W'G2Êf˜$V6ÇÇÜóFV“í”‚∞¢6ˆÁ7B∂Wí“vWDÊ˜Fñfñ6FñˆÂ&ñ÷'îFFT∂WíÜóFV“ì∞¢ñbÇ∂Wíí&WGW&„∞¢ñbÇFî÷ÊÜ2Ü∂WíííFî÷Á6WBÜ∂Wí¬µ“ì∞¢Fî÷ÊvWBÜ∂WííÁW6ÇÜóFV“ì∞¢“ì∞¢6ˆÁ7Bfó'7EvVV∂Fí“Ü÷ˆÁFÖ7F'BÊvWDFíÇí≤bíRs∞¢6ˆÁ7BFó4ñ‰÷ˆÁFÇ“ÊWrFFRÜ÷ˆÁFÖ7F'BÊvWDgV∆≈ñV"Çí¬÷ˆÁFÖ7F'BÊvWD÷ˆÁFÇÇí≤¬íÊvWDFFRÇì∞¢VíÊÊ˜Fñfñ6Fñˆ‰6∆VÊF$w&ñBÊñÊÊW$ÖD‘¬“"#∞¢vVV∂Fî∆&V«2Êf˜$V6ÇÇÜ∆&V¬í”‚∞¢6ˆÁ7BÜVFW"“Fˆ7V÷VÁBÊ7&VFTV∆V÷VÁBÇ&Fób"ì∞¢ÜVFW"Ê6∆74Ê÷R“&Ê˜Fñfñ6Fñˆ‚÷6∆VÊF"÷6V∆¬Ê˜Fñfñ6Fñˆ‚÷6∆VÊF"◊vVV∂Fí#∞¢ÜVFW"ÁFWáD6ˆÁFVÁB“∆&V√∞¢VíÊÊ˜Fñfñ6Fñˆ‰6∆VÊF$w&ñBÊVÊD6Üñ∆BÜÜVFW"ì∞¢“ì∞¢f˜"Ü∆WBí“≤í¬fó'7EvVV∂Fì≤í≥“í∞¢6ˆÁ7BV◊Gí“Fˆ7V÷VÁBÊ7&VFTV∆V÷VÁBÇ&Fób"ì∞¢V◊GíÊ6∆74Ê÷R“&Ê˜Fñfñ6Fñˆ‚÷6∆VÊF"÷6V∆¬ó2÷V◊Gí#∞¢VíÊÊ˜Fñfñ6Fñˆ‰6∆VÊF$w&ñBÊVÊD6Üñ∆BÜV◊Gíì∞¢–¢f˜"Ü∆WBFí“≤Fí√“Fó4ñ‰÷ˆÁFÉ≤Fí≥“í∞¢6ˆÁ7BFFR“ÊWrFFRÜ÷ˆÁFÖ7F'BÊvWDgV∆≈ñV"Çí¬÷ˆÁFÖ7F'BÊvWD÷ˆÁFÇÇí¬Fíì∞¢6ˆÁ7BFFT∂Wí“vWDFFT∂Wîg&ˆ‘∆ˆ6ƒFFRÜFFRì∞¢6ˆÁ7B'F‚“Fˆ7V÷VÁBÊ7&VFTV∆V÷VÁBÇ&'WGFˆ‚"ì∞¢'F‚ÁGóR“&'WGFˆ‚#∞¢'F‚Ê6∆74Ê÷R“&Ê˜Fñfñ6Fñˆ‚÷6∆VÊF"÷6V∆¬Ê˜Fñfñ6Fñˆ‚÷6∆VÊF"÷Fí#∞¢'F‚ÁFWáD6ˆÁFVÁB“7G&ñÊrÜFíì∞¢ñbÜFî÷ÊÜ2ÜFFT∂Wííí'F‚Ê6∆74∆ó7BÊFBÇ&Ü2÷Ê˜Fñfñ6Fñˆ‚"ì∞¢ñbá6V∆V7FVDÊ˜Fñfñ6Fñˆ‰6∆VÊF$FFT∂Wí””“FFT∂Wíí'F‚Ê6∆74∆ó7BÊFBÇ&ó2◊6V∆V7FVB"ì∞¢'F‚ÊFDWfVÁD∆ó7FVÊW"Ç&6∆ñ6≤"¬Çí”‚˜V‰Ê˜Fñfñ6Fñˆ‰FîFWFñ¬ÜFFT∂Wííì∞¢VíÊÊ˜Fñfñ6Fñˆ‰6∆VÊF$w&ñBÊVÊD6Üñ∆BÜ'F‚ì∞¢–ß–†¶7ñÊ2gVÊ7Fñˆ‚˜V‰Ê˜Fñfñ6Fñˆ‰FîFWFñ¬ÜFFT∂Wíí∞¢6V∆V7FVDÊ˜Fñfñ6Fñˆ‰6∆VÊF$FFT∂Wí“7G&ñÊrÜFFT∂Wí«¬""íÁG&ñ“Çì∞¢&VÊFW$Ê˜Fñfñ6Fñˆ‰6∆VÊF"Çì∞¢ñbÇVíÊÊ˜Fñfñ6Fñˆ‰FîFWFñ¬í&WGW&„∞¢ñbÇ6V∆V7FVDÊ˜Fñfñ6Fñˆ‰6∆VÊF$FFT∂Wíí∞¢VíÊÊ˜Fñfñ6Fñˆ‰FîFWFñ¬ÊñÊÊW$ÖD‘¬“"#∞¢&WGW&„∞¢–¢6ˆÁ7BFFT∆&V¬“ÊWrFFRÜG∑6V∆V7FVDÊ˜Fñfñ6Fñˆ‰6∆VÊF$FFT∂Wó’C££íÁFÙ∆ˆ6∆TFFU7G&ñÊrÇ&óB‘ïB"ì∞¢6ˆÁ7BFîóFV◊2“W6W$∆W'G2Êfñ«FW"ÇÜóFV“í”‚vWDÊ˜Fñfñ6FñˆÂ&ñ÷'îFFT∂WíÜóFV“í””“6V∆V7FVDÊ˜Fñfñ6Fñˆ‰6∆VÊF$FFT∂Wíì∞¢VíÊÊ˜Fñfñ6Fñˆ‰FîFWFñ¬ÊñÊÊW$ÖD‘¬“∆ÉC‰FWGFv∆ñÚG∂W66TÖD‘¬ÜFFT∆&V¬ó”¬ˆÉCÊ∞¢6ˆÁ7BFD'WGFˆ‚“7&VFT'WGFˆ‚Ç.)ÈR&ˆw&÷÷Ê˜Fñfñ6W"VW7FÚvñ˜&ÊÚ"¬Çí”‚˜VÂ66ÜVGV∆VDÊ˜Fñfñ6Fñˆ‰f˜&“á6V∆V7FVDÊ˜Fñfñ6Fñˆ‰6∆VÊF$FFT∂Wííì∞¢VíÊÊ˜Fñfñ6Fñˆ‰FîFWFñ¬ÊVÊD6Üñ∆BÜFD'WGFˆ‚ì∞¢ñbÇFîóFV◊2Ê∆VÊwFÇí∞¢6ˆÁ7BV◊Gí“Fˆ7V÷VÁBÊ7&VFTV∆V÷VÁBÇ'"ì∞¢V◊GíÊ6∆74Ê÷R“&◊WFVB#∞¢V◊GíÁFWáD6ˆÁFVÁB“$ÊW77VÊÊ˜Fñfñ6&Vvó7G&FW"VW7FÚvñ˜&ÊÚ‚#∞¢VíÊÊ˜Fñfñ6Fñˆ‰FîFWFñ¬ÊVÊD6Üñ∆BÜV◊Gíì∞¢&WGW&„∞¢–¢6ˆÁ7B6µ6Ê6Ü˜B“vóBF"Ê6ˆ∆∆V7Fñˆ‚Ç'W6W$∆W'D6∂Ê˜v∆VFvV÷VÁG2"íÁvÜW&RÇ&6∂Ê˜v∆VFvVDFFT∂Wí"¬#”“"¬6V∆V7FVDÊ˜Fñfñ6Fñˆ‰6∆VÊF$FFT∂WííÊvWBÇì∞¢6ˆÁ7B6∂Ê˜v∆VFvV÷VÁG2“6µ6Ê6Ü˜BÊFˆ72Ê÷ÇÜFˆ2í”‚á≤ñC¢Fˆ2ÊñB¬‚‚ÊFˆ2ÊFFÇí“íì∞¢FîóFV◊2Êf˜$V6ÇÇÜóFV“í”‚∞¢6ˆÁ7B6&B“Fˆ7V÷VÁBÊ7&VFTV∆V÷VÁBÇ&'Fñ6∆R"ì∞¢6&BÊ6∆74Ê÷R“'6ñ◊∆R÷∆ó7B÷óFV“7F6∂VB#∞¢6ˆÁ7B&V6óñVÁG4ñG2“óFV“Á6VÊEFÙ∆≈&Vvó7FW&V@¢Ú∆Ff˜&’W6W'2Êfñ«FW"ÇáW6W"í”‚F÷ñ‰V÷ñ«2ÊÜ2ÜÊ˜&÷∆ó¶TV÷ñ¬áW6W"ÊV÷ñ¬íííÊ÷ÇáW6W"í”‚W6W"ÊñBê¢¢vWDÊ˜Fñfñ6FñˆÂ&V6óñVÁEW6W$ñG2ÜóFV“ì∞¢6ˆÁ7B&V6óñVÁD∆&V«2“&V6óñVÁG4ñG2Ê÷ÇÜñBí”‚vWE∆Ff˜&’W6W$∆&V¬á∆Ff˜&’W6W'2ÊfñÊBÇáRí”‚RÊñB””“ñBííì∞¢6ˆÁ7BóFV‘6∑2“6∂Ê˜v∆VFvV÷VÁG2Êfñ«FW"ÇÜ6≤í”‚7G&ñÊrÜ6≤ÊÊ˜Fñfñ6Fñˆ‰ñB«¬""í””“7G&ñÊrÜóFV“ÊñB«¬""íì∞¢6ˆÁ7B6∂VDñG2“ÊWr6WBÜóFV‘6∑2Ê÷ÇÜ6≤í”‚7G&ñÊrÜ6≤ÁW6W$ñB«¬""ííì∞¢6ˆÁ7BVÊFñÊt∆&V«2“&V6óñVÁG4ñG2Êfñ«FW"ÇÜñBí”‚6∂VDñG2ÊÜ2ÜñBííÊ÷ÇÜñBí”‚vWE∆Ff˜&’W6W$∆&V¬á∆Ff˜&’W6W'2ÊfñÊBÇáRí”‚RÊñB””“ñBííì∞¢6ˆÁ7B6µ&˜w2“óFV‘6∑2Ê÷ÇÜ6≤í”‚∞¢6ˆÁ7BvÜV‚“6≤Ê6∂Ê˜v∆VFvVDCÚÁFÙFFRÚ6≤Ê6∂Ê˜v∆VFvVDBÁFÙFFRÇíÁFÙ∆ˆ6∆U7G&ñÊrÇ&óB‘ïB"í¢"“#∞¢&WGW&‚∆∆ì‚G∂W66TÖD‘¬Ü6≤ÁW6W$Ê÷R«¬6≤ÁW6W$ñB«¬%WFVÁFR"ó“(
"G∂W66TÖD‘¬ávÜV‚ó”¬ˆ∆ìÊ∞¢“íÊ¶ˆñ‚Ç""ì∞¢6&BÊñÊÊW$ÖD‘¬“ ¢«7G&ˆÊs‚G∂W66TÖD‘¬ÜóFV“ÁFóF∆R«¬$Ê˜Fñfñ6"ó”¬˜7G&ˆÊs‡¢«‚G∂W66TÖD‘¬ÜóFV“Ê÷W76vR«¬""ó”¬˜‡¢«6÷∆√‚G∂W66TÖD‘¬ÜFW7FñÊF&ì¢G∑&V6óñVÁD∆&V«2Ê¶ˆñ‚Ç"¬"í«¬"“'÷ó”¬˜6÷∆√‡¢«6÷∆√‰6ˆÊfW&÷Fì¢G∂óFV‘6∑2Ê∆VÊwFá”¬˜6÷∆√‡¢«6÷∆√‰Êˆ‚6ˆÊfW&÷Fì¢G∑VÊFñÊt∆&V«2Ê∆VÊwFÇÚW66TÖD‘¬áVÊFñÊt∆&V«2Ê¶ˆñ‚Ç"¬"íí¢$ÊW77VÊÚ'”¬˜6÷∆√‡¢«V√‚G∂6µ&˜w2«¬#∆∆ì‰ÊW77VÊ6ˆÊfW&÷¬ˆ∆ì‚'”¬˜V√‡¢∞¢VíÊÊ˜Fñfñ6Fñˆ‰FîFWFñ¬ÊVÊD6Üñ∆BÜ6&Bì∞¢“ì∞ß–†¶gVÊ7Fñˆ‚˜VÂ66ÜVGV∆VDÊ˜Fñfñ6Fñˆ‰f˜&“ÜFFT∂Wíí∞¢6∆˜6TÊ˜Fñfñ6Fñˆ‰6∆VÊF%fñWrÇì∞¢ñbáVíÊÊ˜Fñfñ6Fñˆ‰FFRíVíÊÊ˜Fñfñ6Fñˆ‰FFRÁf«VR“FFT∂Wì∞¢VíÊÊ˜Fñfñ6FñˆÂFóF∆SÚÊfˆ7W2Çì∞ß–†¶gVÊ7Fñˆ‚&VÊFW$WáFW&Êƒ2Çí∞¢ñbÇVíÊWáFW&Êƒ4∆ó7Bí&WGW&„∞¢ñbÇ7W'&VÁEW6W"í∞¢VíÊWáFW&Êƒ4∆ó7BÊñÊÊW$ÖD‘¬“#«6∆73“v◊WFVBs‰fí∆ˆvñ‚W"6ˆ∆∆Vv&RW7FW&ÊR„¬˜‚#∞¢&WGW&„∞¢–¢6ˆÁ7B&˜r“∆Ff˜&’W6W'2ÊfñÊBÇáW6W"í”‚W6W"ÊñB””“7W'&VÁEW6W"ÁVñBì∞¢6ˆÁ7B2“'&íÊó4'&íá&˜sÚÊWáFW&Êƒ2íÚ&˜rÊWáFW&Êƒ2¢µ”∞¢ñbÇ2Ê∆VÊwFÇí∞¢VíÊWáFW&Êƒ4∆ó7BÊñÊÊW$ÖD‘¬“#«6∆73“v◊WFVBs‰ÊW77VÊW7FW&Ê6ˆ∆∆VvF„¬˜‚#∞¢&WGW&„∞¢–¢VíÊWáFW&Êƒ4∆ó7BÊñÊÊW$ÖD‘¬“"#∞¢2Êf˜$V6ÇÇÜ¬ñÊFWÇí”‚∞¢6ˆÁ7BóFV““Fˆ7V÷VÁBÊ7&VFTV∆V÷VÁBÇ&Fób"ì∞¢óFV“Ê6∆74Ê÷R“'6ñ◊∆R÷∆ó7B÷óFV“#∞¢6ˆÁ7B∆&V¬“Fˆ7V÷VÁBÊ7&VFTV∆V÷VÁBÇ&"ì∞¢∆&V¬Êá&Vb“ÁW&√∞¢∆&V¬ÁF&vWB“%ˆ&∆Ê≤#∞¢∆&V¬Á&V¬“&Êˆ˜VÊW"Ê˜&VfW'&W"#∞¢∆&V¬ÁFWáD6ˆÁFVÁB“	˘IrG∂ÊÊ÷R«¬ÁW&«÷∞¢óFV“ÊVÊD6Üñ∆BÜ∆&V¬ì∞¢6ˆÁ7B&V÷˜fT'F‚“7&VFT'WGFˆ‚Ç%&ñ◊V˜fí"¬Çí”‚&V÷˜fTWáFW&ÊƒÜñÊFWÇíì∞¢óFV“ÊVÊD6Üñ∆Bá&V÷˜fT'F‚ì∞¢VíÊWáFW&Êƒ4∆ó7BÊVÊD6Üñ∆BÜóFV“ì∞¢“ì∞ß–†¶7ñÊ2gVÊ7Fñˆ‚6fTWáFW&Êƒf˜$7W'&VÁEW6W"ÜWfVÁBí∞¢WfVÁBÁ&WfVÁDFVfV«BÇì∞¢ñbÇ7W'&VÁEW6W"í&WGW&„∞¢6ˆÁ7BÊ÷R“7G&ñÊráVíÊWáFW&ÊƒÊ÷RÁf«VR«¬""íÁG&ñ“Çì∞¢6ˆÁ7B&uW&¬“7G&ñÊráVíÊWáFW&ÊƒW&¬Áf«VR«¬""íÁG&ñ“Çì∞¢ñbÇÊ÷R«¬&uW&¬í&WGW&„∞¢6ˆÁ7BW&¬“ıÊáGG3Û•¬ı¬ÚˆíÁFW7Bá&uW&¬íÚ&uW&¬¢áGG3¢ÚÚG∑&uW&«÷∞¢6ˆÁ7B&˜r“∆Ff˜&’W6W'2ÊfñÊBÇáW6W"í”‚W6W"ÊñB””“7W'&VÁEW6W"ÁVñBì∞¢6ˆÁ7B2“'&íÊó4'&íá&˜sÚÊWáFW&Êƒ2íÚ&˜rÊWáFW&Êƒ2¢µ”∞¢6ˆÁ7BÊWáB“≤‚‚Ê2¬≤Ê÷R¬W&¬’”∞¢vóBF"Ê6ˆ∆∆V7Fñˆ‚Ç'∆Ff˜&’W6W'2"íÊFˆ2Ü7W'&VÁEW6W"ÁVñBíÁ6WBá∞¢WáFW&Êƒ3¢ÊWáB¿¢WFFVDC¢fó&V&6RÊfó&W7F˜&R‰fñV∆Ef«VRÁ6W'fW%Fñ÷W7F◊Çí¿¢WFFVD'ì¢7W'&VÁEW6W"ÊV÷ñ¬«¬" ¢“¬≤÷W&vS¢G'VR“ì∞¢VíÊWáFW&Êƒf˜&“Á&W6WBÇì∞ß–†¶7ñÊ2gVÊ7Fñˆ‚&V÷˜fTWáFW&ÊƒÜñÊFWÇí∞¢ñbÇ7W'&VÁEW6W"í&WGW&„∞¢6ˆÁ7B&˜r“∆Ff˜&’W6W'2ÊfñÊBÇáW6W"í”‚W6W"ÊñB””“7W'&VÁEW6W"ÁVñBì∞¢6ˆÁ7B2“'&íÊó4'&íá&˜sÚÊWáFW&Êƒ2íÚ&˜rÊWáFW&Êƒ2¢µ”∞¢ñbÜñÊFWÇ¬«¬ñÊFWÇ„“2Ê∆VÊwFÇí&WGW&„∞¢6ˆÁ7BÊWáB“2Êfñ«FW"ÇÖÚ¬íí”‚í”“ñÊFWÇì∞¢vóBF"Ê6ˆ∆∆V7Fñˆ‚Ç'∆Ff˜&’W6W'2"íÊFˆ2Ü7W'&VÁEW6W"ÁVñBíÁ6WBá∞¢WáFW&Êƒ3¢ÊWáB¿¢WFFVDC¢fó&V&6RÊfó&W7F˜&R‰fñV∆Ef«VRÁ6W'fW%Fñ÷W7F◊Çí¿¢WFFVD'ì¢7W'&VÁEW6W"ÊV÷ñ¬«¬" ¢“¬≤÷W&vS¢G'VR“ì∞ß–†¶gVÊ7Fñˆ‚&VÊFW$6ÜE&V6óñVÁG2Çí∞¢VíÊ6ÜE&V6óñVÁBÊñÊÊW$ÖD‘¬“#∆˜Fñˆ‚f«VS“rs‰÷W76vvñÚW"GWGFì¬ˆ˜Fñˆ„‚#∞¢∆Ff˜&’W6W'2Êf˜$V6ÇÇáW6W"í”‚∞¢ñbÜ7W'&VÁEW6W"bbW6W"ÊñB””“7W'&VÁEW6W"ÁVñBí&WGW&„∞¢6ˆÁ7B˜Fñˆ‚“Fˆ7V÷VÁBÊ7&VFTV∆V÷VÁBÇ&˜Fñˆ‚"ì∞¢˜Fñˆ‚Áf«VR“W6W"ÊñC∞¢˜Fñˆ‚ÁFWáD6ˆÁFVÁB“W6W"ÊFó7∆îÊ÷R«¬W6W"ÊV÷ñ¬«¬%WFVÁFR#∞¢VíÊ6ÜE&V6óñVÁBÊVÊD6Üñ∆BÜ˜Fñˆ‚ì∞¢“ì∞ß–†¶6ˆÁ7BÙddî4î≈Ùî’îÂDïÙ4Ù≈T‘Â2“∞¢$‚‚"¬$îB4"¬$FVÊˆ÷ñÊ¶ñˆÊRñ◊ñÁFÚ"¬%Fóˆ∆ˆvññ◊ñÁFÚ"¬$6ˆ◊VÊR"¬$ñÊFó&óß¶Ú"¿¢$∆FóGVFñÊR"¬$∆ˆÊvóGVFñÊR"¬$6ˆFñ6R&Wß¶Ú"¬$&VÚ6ˆ◊WFVÁ¶"¬$FóGFW6V7WG&ñ6R"¿¢$FFW6V7W¶ñˆÊR"¬$˜&W6V7W¶ñˆÊR"¬$˜W&F˜&R"¬$Ê˜FR •”∞†¶6ˆÁ7B‘‰tT‘TÂEÙî’îÂDïÙ4Ù≈T‘Â2“∞¢≤&ÁV÷W&ı&ˆw&W76ófÚ"¬$‚‚%“¬≤&ñE6"¬$îB4%“¬≤&FVÊˆ÷ñÊ¶ñˆÊR"¬$FVÊˆ÷ñÊ¶ñˆÊRñ◊ñÁFÚ%“¿¢≤'Fóˆ∆ˆvññ◊ñÁFÚ"¬%Fóˆ∆ˆvññ◊ñÁFÚ%“¬≤&6ˆ◊VÊR"¬$6ˆ◊VÊR%“¬≤&ñÊFó&óß¶Ú"¬$ñÊFó&óß¶Ú%“¿¢≤&w5í"¬$∆FóGVFñÊR%“¬≤&w5Ç"¬$∆ˆÊvóGVFñÊR%“¬≤&6ˆFñ6U&Wß¶Ú"¬$6ˆFñ6R&Wß¶Ú%“¿¢≤&&V"¬$&VÚ6ˆ◊WFVÁ¶%“¬≤&FóGFW6V7WG&ñ6R"¬$FóGFW6V7WG&ñ6R%“¬≤&FFW6V7W¶ñˆÊR"¬$FFW6V7W¶ñˆÊR%“¿¢≤&˜&W6V7W¶ñˆÊR"¬$˜&W6V7W¶ñˆÊR%“¬≤&˜W&F˜&R"¬$˜W&F˜&R%“¬≤&Ê˜FR"¬$Ê˜FR%“¬≤'7FFÚ"¬%7FFÚ%–•”∞¶6ˆÁ7B‘‰tT‘TÂEı4ı%D$ƒUÙdîTƒE2“ÊWr6WBÖ≤&ÁV÷W&ı&ˆw&W76ófÚ"¬&ñE6"¬&FVÊˆ÷ñÊ¶ñˆÊR"¬&6ˆ◊VÊR"¬&FFW6V7W¶ñˆÊR"¬&˜W&F˜&R"¬'7FFÚ%“ì∞¶6ˆÁ7B‘‰tT‘TÂEı5DEU4U2“≤$Dd$R"¬$dEDÚ"¬$î‚ƒdı$§îÙ‰R"¬%4ı5U4Ú"¬$DdU$îdî4$R%”∞¶∆WB÷ÊvV÷VÁD6ˆ÷÷W76ñB“"#∞¶∆WB÷ÊvV÷VÁE6V&6ÖFñ÷W"“ÁV∆√∞¶∆WB÷ÊvV÷VÁE6˜'B“≤fñV∆C¢&ÁV÷W&ı&ˆw&W76ófÚ"¬Fó&V7Fñˆ„¢”∞¶∆WB÷ÊvV÷VÁEvR“∞¶∆WB÷ÊvV÷VÁDVFóFñÊtñB“"#∞¶∆WB÷ÊvV÷VÁE6fñÊtñB“"#∞¶6ˆÁ7B÷ÊvV÷VÁE6V∆V7FVDñG2“ÊWr6WBÇì∞†¶gVÊ7Fñˆ‚Ê˜&÷∆ó¶T÷ÊvV÷VÁE6V&6Çáf«VRí∞¢&WGW&‚7G&ñÊráf«VR«¬""íÁG&ñ“ÇíÁFÙ∆ˆ6∆T∆˜vW$66RÇ&óB"íÊÊ˜&÷∆ó¶RÇ$‰dB"íÁ&W∆6RÇıµ«S3’«S3fe“ˆr¬""ì∞ß–†¶gVÊ7Fñˆ‚vWD÷ÊvV÷VÁE∆ÁE7FGW2á∆ÁBí∞¢&WGW&‚∆ÁBÊFˆÊRÚ$dEDÚ"¢Ö7G&ñÊrá∆ÁBÁ7FFÚ«¬$Dd$R"íÁG&ñ“ÇíÁFıWW$66RÇí«¬$Dd$R"ì∞ß–†¶gVÊ7Fñˆ‚vWD÷ÊvV÷VÁDWÜV7WFñˆ‚á∆ÁBí∞¢6ˆÁ7B'G2“∆ÁBÊFˆÊTBÚf˜&÷E&ˆ÷TFˆÊU'G2á∆ÁBÊFˆÊTBí¢≤FFS¢""¬Fñ÷S¢""”∞¢&WGW&‚≤FFS¢∆ÁBÊFFW6V7W¶ñˆÊR«¬'G2ÊFFR¬Fñ÷S¢∆ÁBÊ˜&W6V7W¶ñˆÊR«¬'G2ÁFñ÷R¬˜W&F˜#¢∆ÁBÊ˜W&F˜&R«¬∆ÁBÊFˆÊT'í«¬""”∞ß–†¶gVÊ7Fñˆ‚Ü5f∆ñD÷ÊvV÷VÁD6ˆ˜&FñÊFW2á∆ÁBí∞¢6ˆÁ7B∆B“ÁV÷&W"á∆ÁBÊw5íí¬∆Êr“ÁV÷&W"á∆ÁBÊw5Çì∞¢&WGW&‚7G&ñÊrá∆ÁBÊw5íÛÚ""íÁG&ñ“Çí”“""bb7G&ñÊrá∆ÁBÊw5ÇÛÚ""íÁG&ñ“Çí”“""bbÁV÷&W"Êó4fñÊóFRÜ∆BíbbÁV÷&W"Êó4fñÊóFRÜ∆Êríbb∆B„“”ìbb∆B√“ìbb∆Êr„“”Ébb∆Êr√“É∞ß–†¶gVÊ7Fñˆ‚vWD÷ÊvV÷VÁE6V&6Ö&Ê≤á∆ÁB¬VW'íí∞¢ñbÇVW'íí&WGW&‚∞¢6ˆÁ7BWÜV7WFñˆ‚“vWD÷ÊvV÷VÁDWÜV7WFñˆ‚á∆ÁBì∞¢6ˆÁ7Bf«VW2“∑∆ÁBÊFVÊˆ÷ñÊ¶ñˆÊR¬∆ÁBÊñE6¬∆ÁBÊ6ˆ◊VÊR¬∆ÁBÊ6ˆFñ6U&Wß¶Ú¬∆ÁBÁFóˆ∆ˆvññ◊ñÁFÚ¬∆ÁBÊñÊFó&óß¶Ú¬∆ÁBÊ&V¬∆ÁBÊFóGFW6V7WG&ñ6R¬WÜV7WFñˆ‚Ê˜W&F˜"¬∆ÁBÊÊ˜FR«¬∆ÁBÊÊ˜FTñ◊ñÁFı“Ê÷ÜÊ˜&÷∆ó¶T÷ÊvV÷VÁE6V&6Çì∞¢ñbáf«VW5≥“Á7F'G5vóFÇáVW'ííí&WGW&‚∞¢ñbáf«VW5≥“Á7F'G5vóFÇáVW'ííí&WGW&‚#∞¢ñbáf«VW5≥%“Á7F'G5vóFÇáVW'ííí&WGW&‚3∞¢ñbáf«VW5≥5“Á7F'G5vóFÇáVW'ííí&WGW&‚C∞¢ñbáf«VW5≥“ÊñÊ6«VFW2áVW'ííí&WGW&‚S∞¢&WGW&‚f«VW2Á6ˆ÷RÇáf«VRí”‚f«VRÊñÊ6«VFW2áVW'íííÚb¢ñÊfñÊóGì∞ß–†¶gVÊ7Fñˆ‚÷ÊvV÷VÁDfñV∆Ef«VRá∆ÁB¬fñV∆Bí∞¢6ˆÁ7BWÜV7WFñˆ‚“vWD÷ÊvV÷VÁDWÜV7WFñˆ‚á∆ÁBì∞¢ñbÜfñV∆B””“'7FFÚ"í&WGW&‚vWD÷ÊvV÷VÁE∆ÁE7FGW2á∆ÁBì∞¢ñbÜfñV∆B””“&FFW6V7W¶ñˆÊR"í&WGW&‚WÜV7WFñˆ‚ÊFFS∞¢ñbÜfñV∆B””“&˜&W6V7W¶ñˆÊR"í&WGW&‚WÜV7WFñˆ‚ÁFñ÷S∞¢ñbÜfñV∆B””“&˜W&F˜&R"í&WGW&‚WÜV7WFñˆ‚Ê˜W&F˜#∞¢ñbÜfñV∆B””“&Ê˜FR"í&WGW&‚∆ÁBÊÊ˜FR«¬∆ÁBÊÊ˜FTñ◊ñÁFÚ«¬"#∞¢&WGW&‚∆ÁE∂fñV∆E“ÛÚ"#∞ß–†¶gVÊ7Fñˆ‚˜V∆FT÷ÊvV÷VÁDfñ«FW"ÜñB¬∆&V¬¬f«VW2í∞¢6ˆÁ7B6V∆V7B“Fˆ7V÷VÁBÊvWDV∆V÷VÁD'îñBÜñBì∞¢ñbÇ6V∆V7Bí&WGW&„∞¢6ˆÁ7B&Wfñ˜W2“6V∆V7BÁf«VS∞¢6V∆V7BÊñÊÊW$ÖD‘¬“∆˜Fñˆ‚f«VS“"#‚G∂∆&V«”¬ˆ˜Fñˆ„Ê≤≤‚‚ÊÊWr6WBáf«VW2Êfñ«FW"Ñ&ˆˆ∆V‚íï“Á6˜'BÇÜ¬"í”‚Ê∆ˆ6∆T6ˆ◊&RÜ"¬&óB"ííÊ÷Çáf«VRí”‚∆˜Fñˆ‚f«VS“"G∂W66TÖD‘¬áf«VRó“#‚G∂W66TÖD‘¬áf«VRó”¬ˆ˜Fñˆ„ÊíÊ¶ˆñ‚Ç""ì∞¢ñbÖ≤‚‚Á6V∆V7BÊ˜FñˆÁ5“Á6ˆ÷RÇÜ˜Fñˆ‚í”‚˜Fñˆ‚Áf«VR””“&Wfñ˜W2íí6V∆V7BÁf«VR“&Wfñ˜W3∞ß–†¶gVÊ7Fñˆ‚&VÊFW$ñ◊ñÁFî÷ÊvV÷VÁEF&∆RÇí∞¢6ˆÁ7BF&ˆGí“Fˆ7V÷VÁBÊvWDV∆V÷VÁD'îñBÇ&ñ◊ñÁFí÷÷ÊvV÷VÁB◊F&ˆGí"ì∞¢6ˆÁ7BFÜVB“Fˆ7V÷VÁBÊvWDV∆V÷VÁD'îñBÇ&ñ◊ñÁFí÷÷ÊvV÷VÁB◊FÜVB"ì∞¢ñbÇF&ˆGí«¬FÜVB«¬÷ÊvV÷VÁD6ˆ÷÷W76ñBí&WGW&„∞¢6ˆÁ7B∆¬“vWD6ˆ÷÷W7666ÜVDñ◊ñÁFíÜ÷ÊvV÷VÁD6ˆ÷÷W76ñBì∞¢6ˆÁ7BVW'í“Ê˜&÷∆ó¶T÷ÊvV÷VÁE6V&6ÇÜFˆ7V÷VÁBÊvWDV∆V÷VÁD'îñBÇ&ñ◊ñÁFí÷÷ÊvV÷VÁB◊6V&6Ç"ìÚÁf«VRì∞¢6ˆÁ7B7FGW4fñ«FW"“Fˆ7V÷VÁBÊvWDV∆V÷VÁD'îñBÇ&ñ◊ñÁFí◊7FGW2÷fñ«FW""ìÚÁf«VR«¬&∆¬#∞¢6ˆÁ7B6ˆ◊VÊR“Fˆ7V÷VÁBÊvWDV∆V÷VÁD'îñBÇ&ñ◊ñÁFí÷6ˆ◊VÊR÷fñ«FW""ìÚÁf«VR«¬"#∞¢6ˆÁ7BFóˆ∆ˆvñ“Fˆ7V÷VÁBÊvWDV∆V÷VÁD'îñBÇ&ñ◊ñÁFí◊Fóˆ∆ˆvñ÷fñ«FW""ìÚÁf«VR«¬"#∞¢6ˆÁ7B˜W&F˜&R“Fˆ7V÷VÁBÊvWDV∆V÷VÁD'îñBÇ&ñ◊ñÁFí÷˜W&F˜&R÷fñ«FW""ìÚÁf«VR«¬"#∞¢˜V∆FT÷ÊvV÷VÁDfñ«FW"Ç&ñ◊ñÁFí÷6ˆ◊VÊR÷fñ«FW""¬$6ˆ◊VÊR"¬∆¬Ê÷Çáí”‚7G&ñÊráÊ6ˆ◊VÊR«¬""ííì∞¢˜V∆FT÷ÊvV÷VÁDfñ«FW"Ç&ñ◊ñÁFí◊Fóˆ∆ˆvñ÷fñ«FW""¬%Fóˆ∆ˆvñ"¬∆¬Ê÷Çáí”‚7G&ñÊráÁFóˆ∆ˆvññ◊ñÁFÚ«¬""ííì∞¢˜V∆FT÷ÊvV÷VÁDfñ«FW"Ç&ñ◊ñÁFí÷˜W&F˜&R÷fñ«FW""¬$˜W&F˜&R"¬∆¬Ê÷Çáí”‚vWD÷ÊvV÷VÁDWÜV7WFñˆ‚áíÊ˜W&F˜"íì∞¢6ˆÁ7B÷F6ÜW57FGW2“á∆ÁBí”‚7FGW4fñ«FW"””“&∆¬ ¢«¬á7FGW4fñ«FW"””“&FˆÊR"bb∆ÁBÊFˆÊRí«¬á7FGW4fñ«FW"””“'FˆFÚ"bb∆ÁBÊFˆÊRê¢«¬á7FGW4fñ«FW"””“'vóFÇ÷6ˆ˜&FñÊFW2"bbÜ5f∆ñD÷ÊvV÷VÁD6ˆ˜&FñÊFW2á∆ÁBíê¢«¬á7FGW4fñ«FW"””“'vóFÜ˜WB÷6ˆ˜&FñÊFW2"bbÜ5f∆ñD÷ÊvV÷VÁD6ˆ˜&FñÊFW2á∆ÁBíê¢«¬á7FGW4fñ«FW"””“&W'&˜'2"bbÇ7G&ñÊrá∆ÁBÊFVÊˆ÷ñÊ¶ñˆÊR«¬""íÁG&ñ“Çí«¬ÇÖ7G&ñÊrá∆ÁBÊw5íÛÚ""íÁG&ñ“Çí«¬7G&ñÊrá∆ÁBÊw5ÇÛÚ""íÁG&ñ“ÇííbbÜ5f∆ñD÷ÊvV÷VÁD6ˆ˜&FñÊFW2á∆ÁBíííì∞¢6ˆÁ7Bfñ«FW&VB“∆¬Ê÷Çá∆ÁBí”‚á≤∆ÁB¬&Ê≥¢vWD÷ÊvV÷VÁE6V&6Ö&Ê≤á∆ÁB¬VW'íí“ííÊfñ«FW"Çá≤∆ÁB¬&Ê≤“í”‚&Ê≤¬ñÊfñÊóGíbb÷F6ÜW57FGW2á∆ÁBíbbÇ6ˆ◊VÊR«¬∆ÁBÊ6ˆ◊VÊR””“6ˆ◊VÊRíbbÇFóˆ∆ˆvñ«¬∆ÁBÁFóˆ∆ˆvññ◊ñÁFÚ””“Fóˆ∆ˆvñíbbÇ˜W&F˜&R«¬vWD÷ÊvV÷VÁDWÜV7WFñˆ‚á∆ÁBíÊ˜W&F˜"””“˜W&F˜&Ríì∞¢fñ«FW&VBÁ6˜'BÇÜ¬"í”‚áVW'íbbÁ&Ê≤”“"Á&Ê≤ÚÁ&Ê≤“"Á&Ê≤¢÷ÊvV÷VÁE6˜'BÊFó&V7Fñˆ‚¢7G&ñÊrÜ÷ÊvV÷VÁDfñV∆Ef«VRÜÁ∆ÁB¬÷ÊvV÷VÁE6˜'BÊfñV∆BííÊ∆ˆ6∆T6ˆ◊&RÖ7G&ñÊrÜ÷ÊvV÷VÁDfñV∆Ef«VRÜ"Á∆ÁB¬÷ÊvV÷VÁE6˜'BÊfñV∆Bíí¬&óB"¬≤ÁV÷W&ñ3¢G'VR“ííì∞¢6ˆÁ7BFˆÊR“∆¬Êfñ«FW"Çá∆ÁBí”‚∆ÁBÊFˆÊRíÊ∆VÊwFÉ∞¢Fˆ7V÷VÁBÊvWDV∆V÷VÁD'îñBÇ&ñ◊ñÁFí÷÷ÊvV÷VÁB◊7FG2"íÊñÊÊW$ÖD‘¬“«7„„∆#‚G∂∆¬Ê∆VÊwFá”¬ˆ#‚ñ◊ñÁFì¬˜7„„«7‚6∆73“&ó2÷FˆÊR#„∆#‚G∂FˆÊW”¬ˆ#‚fGFì¬˜7„„«7„„∆#‚G∂∆¬Ê∆VÊwFÇ“FˆÊW”¬ˆ#‚Ff&S¬˜7„Ê∞¢FÜVBÊñÊÊW$ÖD‘¬“«G#„«FÇ6∆73“'6ÜVWB◊6V∆V7B#„∆ñÁWBGóS“&6ÜV6∂&˜Ç"FF◊6V∆V7B÷∆¬&ñ÷∆&V√“%6V∆W¶ñˆÊGWGFí#„¬˜FÉ‚G¥‘‰tT‘TÂEÙî’îÂDïÙ4Ù≈T‘Â2Ê÷ÇÖ∂fñV∆B¬∆&V≈“í”‚«FÇFF◊6˜'C“"G¥‘‰tT‘TÂEı4ı%D$ƒUÙdîTƒE2ÊÜ2ÜfñV∆BíÚfñV∆B¢"'“"6∆73“"G∂fñV∆B””“&ÁV÷W&ı&ˆw&W76ófÚ"Ú'6ÜVWB÷ÁV÷&W""¢"'“#‚G∂∆&V«“G∂÷ÊvV÷VÁE6˜'BÊfñV∆B””“fñV∆BÚÜ÷ÊvV÷VÁE6˜'BÊFó&V7Fñˆ‚‚Ú"(i"¢"(i2"í¢"'”¬˜FÉÊíÊ¶ˆñ‚Ç""ó”«FÉ‰¶ñˆÊì¬˜FÉ„¬˜G#Ê∞¢6ˆÁ7BvU6ó¶R“¬vW2“÷FÇÊ÷ÇÉ¬÷FÇÊ6Vñ¬Üfñ«FW&VBÊ∆VÊwFÇÚvU6ó¶Ríì≤÷ÊvV÷VÁEvR“÷FÇÊ÷ñ‚Ü÷ÊvV÷VÁEvR¬vW2ì∞¢6ˆÁ7BvU&˜w2“fñ«FW&VBÁ6∆ñ6RÇÜ÷ÊvV÷VÁEvR“í¢vU6ó¶R¬÷ÊvV÷VÁEvR¢vU6ó¶Rì∞¢ñbÇvU&˜w2Ê∆VÊwFÇíF&ˆGíÊñÊÊW$ÖD‘¬“«G#„«FB6ˆ«7„“#Ç"6∆73“'6ÜVWB÷V◊Gí#‰ÊW77V‚ñ◊ñÁFÚG&˜fFÚ„¬˜FC„¬˜G#Ê∞¢V«6RF&ˆGíÊñÊÊW$ÖD‘¬“vU&˜w2Ê÷Çá≤∆ÁB“í”‚&VÊFW$÷ÊvV÷VÁE∆ÁE&˜rá∆ÁBííÊ¶ˆñ‚Ç""ì∞¢Fˆ7V÷VÁBÊvWDV∆V÷VÁD'îñBÇ&ñ◊ñÁFí÷÷ÊvV÷VÁB◊vñÊFñˆ‚"íÊñÊÊW$ÖD‘¬“vW2‚Ú∆'WGFˆ‚6∆73“&'F‚"FF◊vS“"G∂÷ÊvV÷VÁEvR““"G∂÷ÊvV÷VÁEvR””“Ú&Fó6&∆VB"¢"'”Ó(i¬ˆ'WGFˆ„„«7„ÂvñÊG∂÷ÊvV÷VÁEvW“FíG∑vW7“+rG∂fñ«FW&VBÊ∆VÊwFá“&ó7V«FFì¬˜7„„∆'WGFˆ‚6∆73“&'F‚"FF◊vS“"G∂÷ÊvV÷VÁEvR≤“"G∂÷ÊvV÷VÁEvR””“vW2Ú&Fó6&∆VB"¢"'”Ó(i#¬ˆ'WGFˆ„Ê¢«7„‚G∂fñ«FW&VBÊ∆VÊwFá“&ó7V«FFì¬˜7„Ê∞¢WFFT÷ÊvV÷VÁD'V∆¥&"Çì∞ß–†¶gVÊ7Fñˆ‚&VÊFW$÷ÊvV÷VÁE∆ÁE&˜rá∆ÁBí∞¢6ˆÁ7BVFóFñÊr“÷ÊvV÷VÁDVFóFñÊtñB””“∆ÁBÊñB¬6fñÊr“÷ÊvV÷VÁE6fñÊtñB””“∆ÁBÊñC∞¢6ˆÁ7BVFóF&∆TfñV∆G2“ÊWr6WBÑ‘‰tT‘TÂEÙî’îÂDïÙ4Ù≈T‘Â2Á6∆ñ6RÉíÊ÷ÇÖ∂fñV∆E“í”‚fñV∆Bíì∞¢6ˆÁ7B6V∆«2“‘‰tT‘TÂEÙî’îÂDïÙ4Ù≈T‘Â2Ê÷ÇÖ∂fñV∆E“í”‚∞¢6ˆÁ7Bf«VR“÷ÊvV÷VÁDfñV∆Ef«VRá∆ÁB¬fñV∆Bì∞¢ñbÇVFóFñÊr«¬VFóF&∆TfñV∆G2ÊÜ2ÜfñV∆Bíí&WGW&‚«FB6∆73“"G∂fñV∆B””“&ÁV÷W&ı&ˆw&W76ófÚ"Ú'6ÜVWB÷ÁV÷&W""¢"'“#‚G∂fñV∆B””“'7FFÚ"Ú«7‚6∆73“'∆ÁB◊7FGW27FGW2“G∂Ê˜&÷∆ó¶T÷ÊvV÷VÁE6V&6Çáf«VRíÁ&W∆6RÇı«2≤ˆr¬"“"ó“#‚G∂W66TÖD‘¬áf«VRó”¬˜7„Ê¢W66TÖD‘¬áf«VR«¬ÜfñV∆B””“&ÁV÷W&ı&ˆw&W76ófÚ"Ú.(	B"¢""íó”¬˜FCÊ∞¢ñbÜfñV∆B””“'7FFÚ"í&WGW&‚«FC„«6V∆V7BFF÷fñV∆C“'7FFÚ#‚G¥‘‰tT‘TÂEı5DEU4U2Ê÷Çá7FGW2í”‚∆˜Fñˆ‚G∑7FGW2””“f«VRÚ'6V∆V7FVB"¢"'”‚G∑7FGW7”¬ˆ˜Fñˆ„ÊíÊ¶ˆñ‚Ç""ó”¬˜6V∆V7C„¬˜FCÊ∞¢6ˆÁ7BGóR“fñV∆B””“&FFW6V7W¶ñˆÊR"Ú&FFR"¢fñV∆B””“&˜&W6V7W¶ñˆÊR"Ú'Fñ÷R"¢'FWáB#∞¢&WGW&‚«FC„∆ñÁWBFF÷fñV∆C“"G∂fñV∆G“"GóS“"G∑GóW“"f«VS“"G∂W66TÖD‘¬áf«VRó“"G∑6fñÊrÚ&Fó6&∆VB"¢"'”„¬˜FCÊ∞¢“íÊ¶ˆñ‚Ç""ì∞¢&WGW&‚«G"FF◊∆ÁB÷ñC“"G∂W66TÖD‘¬á∆ÁBÊñBó“"6∆73“"G∂VFóFñÊrÚ&ó2÷VFóFñÊr"¢"'“#„«FB6∆73“'6ÜVWB◊6V∆V7B#„∆ñÁWBGóS“&6ÜV6∂&˜Ç"FF◊6V∆V7B◊&˜rG∂÷ÊvV÷VÁE6V∆V7FVDñG2ÊÜ2á∆ÁBÊñBíÚ&6ÜV6∂VB"¢"'“&ñ÷∆&V√“%6V∆W¶ñˆÊñ◊ñÁFÚ#„¬˜FC‚G∂6V∆«7”«FB6∆73“'6ÜVWB÷7FñˆÁ2#‚G∂VFóFñÊrÚ∆'WGFˆ‚6∆73“&'F‚'F‚◊&ñ÷'í"FF◊&˜r÷7Fñˆ„“'6fR"G∑6fñÊrÚ&Fó6&∆VB"¢"'”‚G∑6fñÊrÚ%6«fFvvñ˛(
b"¢%6«f'”¬ˆ'WGFˆ„„∆'WGFˆ‚6∆73“&'F‚"FF◊&˜r÷7Fñˆ„“&6Ê6V¬"G∑6fñÊrÚ&Fó6&∆VB"¢"'”‰ÊÁV∆∆¬ˆ'WGFˆ„„«6∆73“'&˜r÷fVVF&6≤"&ˆ∆S“&∆W'B#„¬˜Ê¢Ü6‰÷ÊvTFFÇíÚ∆'WGFˆ‚6∆73“&'F‚"FF◊&˜r÷7Fñˆ„“&VFóB#Ó)»˛˚àÚ÷ˆFñfñ6¬ˆ'WGFˆ„Ê¢""ó”¬˜FC„¬˜G#Ê∞ß–†¶gVÊ7Fñˆ‚˜V‰ÊWt6ˆ÷÷W76Fñ∆ˆrÇí∞¢6ˆÁ7B÷ˆF¬“Fˆ7V÷VÁBÊvWDV∆V÷VÁD'îñBÇ&ÊWr÷6ˆ÷÷W76÷÷ˆF¬"ì∞¢÷ˆF√ÚÊ6∆74∆ó7BÁ&V÷˜fRÇ&ÜñFFV‚"ì∞¢÷ˆF√ÚÁ6WDGG&ñ'WFRÇ&&ñ÷ÜñFFV‚"¬&f«6R"ì∞¢Fˆ7V÷VÁBÊvWDV∆V÷VÁD'îñBÇ&6ˆ÷÷W76÷Ê÷R"ìÚÊfˆ7W2Çì∞ß–†¶gVÊ7Fñˆ‚6∆˜6TÊWt6ˆ÷÷W76Fñ∆ˆrÇí∞¢6ˆÁ7B÷ˆF¬“Fˆ7V÷VÁBÊvWDV∆V÷VÁD'îñBÇ&ÊWr÷6ˆ÷÷W76÷÷ˆF¬"ì∞¢÷ˆF√ÚÊ6∆74∆ó7BÊFBÇ&ÜñFFV‚"ì∞¢÷ˆF√ÚÁ6WDGG&ñ'WFRÇ&&ñ÷ÜñFFV‚"¬'G'VR"ì∞ß–†¶gVÊ7Fñˆ‚˜V‰ñ◊ñÁFî÷ÊvV÷VÁBÜ6ˆ÷÷W76í∞¢ñbávñÊF˜r‰66˜VÁFñÊuc"í&WGW&‚vñÊF˜r‰66˜VÁFñÊuc"Ê˜V‚Ü6ˆ÷÷W76ì∞¢ñbÇ6ˆ÷÷W76ÚÊñBí&WGW&„∞¢6ˆÁ7B67&VV‚“Fˆ7V÷VÁBÊvWDV∆V÷VÁD'îñBÇ&ñ◊ñÁFí÷÷ÊvV÷VÁB◊67&VV‚"ì∞¢Fˆ7V÷VÁBÊvWDV∆V÷VÁD'îñBÇ&6ˆ÷÷W76R÷÷ÊvR÷∆ó7B"ìÚÊ6∆74∆ó7BÊFBÇ&ÜñFFV‚"ì∞¢Fˆ7V÷VÁBÁVW'ï6V∆V7F˜"Ç"Ê6ˆ÷÷W76R÷÷ÊvV÷VÁB÷ÜVB"ìÚÊ6∆74∆ó7BÊFBÇ&ÜñFFV‚"ì∞¢Fˆ7V÷VÁBÁVW'ï6V∆V7F˜"Ç"Ê6ˆ÷÷W76R÷÷ÊvV÷VÁB◊6V&6Ç"ìÚÊ6∆74∆ó7BÊFBÇ&ÜñFFV‚"ì∞¢67&VV„ÚÊ6∆74∆ó7BÁ&V÷˜fRÇ&ÜñFFV‚"ì∞¢67&VV„ÚÁ6WDGG&ñ'WFRÇ&&ñ÷ÜñFFV‚"¬&f«6R"ì∞¢ñbáVíÊ6ˆ÷÷W76F&vWE6V∆V7BíVíÊ6ˆ÷÷W76F&vWE6V∆V7BÁf«VR“6ˆ÷÷W76ÊñC∞¢÷ÊvV÷VÁD6ˆ÷÷W76ñB“6ˆ÷÷W76ÊñC∞¢÷ÊvV÷VÁEvR“∞¢÷ÊvV÷VÁDVFóFñÊtñB“"#∞¢÷ÊvV÷VÁE6V∆V7FVDñG2Ê6∆V"Çì∞¢6ˆÁ7BF˜F¬“ÁV÷&W"ÜvWD6ˆ÷÷W767FG2Ü6ˆ÷÷W76ÊñBíÁF˜F¬«¬ì∞¢6ˆÁ7BFóF∆R“Fˆ7V÷VÁBÊvWDV∆V÷VÁD'îñBÇ&ñ◊ñÁFí÷÷ÊvV÷VÁB◊FóF∆R"ì∞¢6ˆÁ7B÷WF“Fˆ7V÷VÁBÊvWDV∆V÷VÁD'îñBÇ&ñ◊ñÁFí÷÷ÊvV÷VÁB÷÷WF"ì∞¢ñbáFóF∆RíFóF∆RÁFWáD6ˆÁFVÁB“6ˆ÷÷W76ÊÊˆ÷R«¬$vW7FñˆÊRñ◊ñÁFí#∞¢ñbÜ÷WFí÷WFÁFWáD6ˆÁFVÁB“6ˆB‚G∂6ˆ÷÷W76Ê6ˆFñ6R«¬.(	B'“(
"G∑F˜F«“G∑F˜F¬””“Ú&ñ◊ñÁFÚ"¢&ñ◊ñÁFí'÷∞¢&VÊFW$ñ◊ñÁFî÷ÊvV÷VÁEF&∆RÇì∞¢67&VV„ÚÁ67&ˆ∆ƒñÁFıfñWrá≤&∆ˆ6≥¢'7F'B"“ì∞ß–†¶gVÊ7Fñˆ‚6∆˜6Tñ◊ñÁFî÷ÊvV÷VÁBÇí∞¢6ˆÁ7B67&VV‚“Fˆ7V÷VÁBÊvWDV∆V÷VÁD'îñBÇ&ñ◊ñÁFí÷÷ÊvV÷VÁB◊67&VV‚"ì∞¢67&VV„ÚÊ6∆74∆ó7BÊFBÇ&ÜñFFV‚"ì∞¢67&VV„ÚÁ6WDGG&ñ'WFRÇ&&ñ÷ÜñFFV‚"¬'G'VR"ì∞¢Fˆ7V÷VÁBÊvWDV∆V÷VÁD'îñBÇ&6ˆ÷÷W76R÷÷ÊvR÷∆ó7B"ìÚÊ6∆74∆ó7BÁ&V÷˜fRÇ&ÜñFFV‚"ì∞¢Fˆ7V÷VÁBÁVW'ï6V∆V7F˜"Ç"Ê6ˆ÷÷W76R÷÷ÊvV÷VÁB÷ÜVB"ìÚÊ6∆74∆ó7BÁ&V÷˜fRÇ&ÜñFFV‚"ì∞¢Fˆ7V÷VÁBÁVW'ï6V∆V7F˜"Ç"Ê6ˆ÷÷W76R÷÷ÊvV÷VÁB◊6V&6Ç"ìÚÊ6∆74∆ó7BÁ&V÷˜fRÇ&ÜñFFV‚"ì∞¢÷ÊvV÷VÁD6ˆ÷÷W76ñB“"#∞ß–†¶gVÊ7Fñˆ‚f∆ñFFT÷ÊvV÷VÁE∆ÁBáF6Ç¬∆ÁDñB¬ÁV÷&W"í∞¢ñbÇF6ÇÊFVÊˆ÷ñÊ¶ñˆÊRí&WGW&‚ñ◊ñÁFÚ‚‚G∂ÁV÷&W"«¬.(	B'”¢FVÊˆ÷ñÊ¶ñˆÊRˆ&&∆ñvF˜&ñÊ∞¢6ˆÁ7BÜ4∆B“7G&ñÊráF6ÇÊw5íÛÚ""íÁG&ñ“Çí”“""¬Ü4∆Êr“7G&ñÊráF6ÇÊw5ÇÛÚ""íÁG&ñ“Çí”“"#∞¢ñbÜÜ4∆B”“Ü4∆Êr«¬ÇÜÜ4∆B«¬Ü4∆ÊríbbÜ5f∆ñD÷ÊvV÷VÁD6ˆ˜&FñÊFW2áF6Çííí&WGW&‚ñ◊ñÁFÚ‚‚G∂ÁV÷&W"«¬.(	B'”¢6ˆ˜&FñÊFRÊˆ‚f∆ñFRÊ∞¢ñbáF6ÇÊñE6í∞¢6ˆÁ7BGW∆ñ6FR“vWD6ˆ÷÷W7666ÜVDñ◊ñÁFíÜ÷ÊvV÷VÁD6ˆ÷÷W76ñBíÊfñÊBÇÜóFV“í”‚óFV“ÊñB”“∆ÁDñBbbÊ˜&÷∆ó¶T÷ÊvV÷VÁE6V&6ÇÜóFV“ÊñE6í””“Ê˜&÷∆ó¶T÷ÊvV÷VÁE6V&6ÇáF6ÇÊñE6íì∞¢ñbÜGW∆ñ6FRí&WGW&‚ñ◊ñÁFÚ‚‚G∂ÁV÷&W"«¬.(	B'”¢îB4vú:WFñ∆óß¶FÚF∆¬vñ◊ñÁFÚ‚‚G∂GW∆ñ6FRÊÁV÷W&ı&ˆw&W76ófÚ«¬.(	B'“Ê∞¢–¢&WGW&‚"#∞ß–†¶7ñÊ2gVÊ7Fñˆ‚6fT÷ÊvV÷VÁE∆ÁE&˜rá&˜rí∞¢ñbÇ6‰÷ÊvTFFÇí«¬÷ÊvV÷VÁE6fñÊtñBí&WGW&„∞¢6ˆÁ7B∆ÁDñB“&˜rÊFF6WBÁ∆ÁDñC∞¢6ˆÁ7B˜&ñvñÊ¬“vWD6ˆ÷÷W7666ÜVDñ◊ñÁFíÜ÷ÊvV÷VÁD6ˆ÷÷W76ñBíÊfñÊBÇá∆ÁBí”‚∆ÁBÊñB””“∆ÁDñBì∞¢ñbÇ˜&ñvñÊ¬í&WGW&„∞¢6ˆÁ7BF6Ç“∑”∞¢&˜rÁVW'ï6V∆V7F˜$∆¬Ç%∂FF÷fñV∆E“"íÊf˜$V6ÇÇÜñÁWBí”‚≤F6Ö∂ñÁWBÊFF6WBÊfñV∆E““7G&ñÊrÜñÁWBÁf«VR«¬""íÁG&ñ“Çì≤“ì∞¢F6ÇÊw5í“F6ÇÊw5í””“""ÚÁV∆¬¢ÁV÷&W"Ö7G&ñÊráF6ÇÊw5ííÁ&W∆6RÇ"¬"¬"‚"íì∞¢F6ÇÊw5Ç“F6ÇÊw5Ç””“""ÚÁV∆¬¢ÁV÷&W"Ö7G&ñÊráF6ÇÊw5ÇíÁ&W∆6RÇ"¬"¬"‚"íì∞¢6ˆÁ7BW'&˜"“f∆ñFFT÷ÊvV÷VÁE∆ÁBáF6Ç¬∆ÁDñB¬˜&ñvñÊ¬ÊÁV÷W&ı&ˆw&W76ófÚì∞¢ñbÜW'&˜"í≤&˜rÁVW'ï6V∆V7F˜"Ç"Á&˜r÷fVVF&6≤"íÁFWáD6ˆÁFVÁB“W'&˜#≤&WGW&„≤–¢÷ÊvV÷VÁE6fñÊtñB“∆ÁDñC≤&VÊFW$ñ◊ñÁFî÷ÊvV÷VÁEF&∆RÇì∞¢G'í∞¢6ˆÁ7BÊWáDFˆÊR“F6ÇÁ7FFÚ””“$dEDÚ#∞¢ñbÜÊWáDFˆÊR”“&ˆˆ∆V‚Ü˜&ñvñÊ¬ÊFˆÊRíívóB6WDñ◊ñÁFÙFˆÊRÜ÷ÊvV÷VÁD6ˆ÷÷W76ñB¬vWDñ◊ñÁFÙFˆ4ñG2Ü˜&ñvñÊ¬í¬ÊWáDFˆÊR¬≤FˆÊT'ì¢F6ÇÊ˜W&F˜&R«¬vWD˜W&F˜$Fó7∆îÊ÷RÇí“ì∞¢6ˆÁ7Bñ∆ˆB“∞¢ñE6¢F6ÇÊñE6¬FVÊˆ÷ñÊ¶ñˆÊS¢F6ÇÊFVÊˆ÷ñÊ¶ñˆÊR¬Fóˆ∆ˆvññ◊ñÁFÛ¢F6ÇÁFóˆ∆ˆvññ◊ñÁFÚ¿¢6ˆ◊VÊS¢F6ÇÊ6ˆ◊VÊR¬ñÊFó&óß¶Û¢F6ÇÊñÊFó&óß¶Ú¬w5ì¢F6ÇÊw5í¬w5É¢F6ÇÊw5Ç¿¢6ˆFñ6U&Wß¶Û¢F6ÇÊ6ˆFñ6U&Wß¶Ú¬&V¢F6ÇÊ&V¬FóGFW6V7WG&ñ6S¢F6ÇÊFóGFW6V7WG&ñ6R¿¢FFW6V7W¶ñˆÊS¢F6ÇÊFFW6V7W¶ñˆÊR¬˜&W6V7W¶ñˆÊS¢F6ÇÊ˜&W6V7W¶ñˆÊR¬˜W&F˜&S¢F6ÇÊ˜W&F˜&R¿¢Ê˜FS¢F6ÇÊÊ˜FR¬Ê˜FTñ◊ñÁFÛ¢F6ÇÊÊ˜FR¬7FFÛ¢F6ÇÁ7FFÚ¬WFFVDC¢fó&V&6RÊfó&W7F˜&R‰fñV∆Ef«VRÁ6W'fW%Fñ÷W7F◊Çí¿¢WFFVD'ïVñC¢7W'&VÁEW6W#ÚÁVñB«¬""¬WFFVD'îÊ÷S¢vWD˜W&F˜$Fó7∆îÊ÷RÇê¢”∞¢vóBF"Ê6ˆ∆∆V7Fñˆ‚ÜvWD6ˆ÷÷W76T6ˆ∆∆V7Fñˆ‰Ê÷RÇííÊFˆ2Ü÷ÊvV÷VÁD6ˆ÷÷W76ñBíÊ6ˆ∆∆V7Fñˆ‚Ç&ñ◊ñÁFí"íÊFˆ2á∆ÁDñBíÁ6WBáñ∆ˆB¬≤÷W&vS¢G'VR“ì∞¢ñÁf∆ñFFTñ◊ñÁFıvÜG4FV◊∆FRÜvWDñ◊ñÁFÙFˆ4ñG2Ü˜&ñvñÊ¬íì∞¢÷ÊvV÷VÁDVFóFñÊtñB“"#∞¢“6F6ÇÜW'&˜%6fRí∞¢6ˆÁ6ˆ∆RÊW'&˜"Ç%6«fFvvñÚñ◊ñÁFÚF∆∆F&V∆∆f∆∆óFÛ¢"¬W'&˜%6fRì∞¢6ˆÁ7BfVVF&6≤“Fˆ7V÷VÁBÁVW'ï6V∆V7F˜"Ü∂FF◊∆ÁB÷ñC“"G¥552ÊW66Rá∆ÁDñBó“%“Á&˜r÷fVVF&6∂ì∞¢ñbÜfVVF&6≤ífVVF&6≤ÁFWáD6ˆÁFVÁB“W'&˜%6fRÊ÷W76vR«¬%6«fFvvñÚÊˆ‚&óW66óFÚ‚&ó&˜f‚#∞¢“fñÊ∆«í≤÷ÊvV÷VÁE6fñÊtñB“"#≤&VÊFW$ñ◊ñÁFî÷ÊvV÷VÁEF&∆RÇì≤–ß–†¶7ñÊ2gVÊ7Fñˆ‚FD÷ÊvV÷VÁDñ◊ñÁFÚÇí∞¢ñbÇ6‰÷ÊvTFFÇí«¬÷ÊvV÷VÁD6ˆ÷÷W76ñBí&WGW&„∞¢6ˆÁ7BFVÊˆ÷ñÊ¶ñˆÊR“7G&ñÊrávñÊF˜rÁ&ˆ◊BÇ$FVÊˆ÷ñÊ¶ñˆÊRñ◊ñÁFÛ¢"¬""í«¬""íÁG&ñ“Çì∞¢ñbÇFVÊˆ÷ñÊ¶ñˆÊRí&WGW&„∞¢6ˆÁ7BñE6“7G&ñÊrávñÊF˜rÁ&ˆ◊BÇ$îB4Ü˜¶ñˆÊ∆Rì¢"¬""í«¬""íÁG&ñ“Çì∞¢6ˆÁ7B∆¬“vWD6ˆ÷÷W7666ÜVDñ◊ñÁFíÜ÷ÊvV÷VÁD6ˆ÷÷W76ñBì∞¢6ˆÁ7B6ˆ÷÷W76“6ˆ÷÷W76T'îñBÊvWBÜ÷ÊvV÷VÁD6ˆ÷÷W76ñBí«¬∑”∞¢6ˆÁ7B∆Vv7í“∆¬Ê∆VÊwFÇ‚bbÁV÷&W"Ü6ˆ÷÷W76ÊWÜ6Vƒ÷ˆFV≈fW'6ñˆ‚«¬í¬#∞¢6ˆÁ7BÁV÷W&ı&ˆw&W76ófÚ“∆Vv7íÚÁV∆¬¢÷FÇÊ÷ÇÉ¬‚‚Ê∆¬Ê÷Çá∆ÁBí”‚ÁV÷&W"á∆ÁBÊÁV÷W&ı&ˆw&W76ófÚ«¬ííí≤∞¢6ˆÁ7Bñ∆ˆB“≤FVÊˆ÷ñÊ¶ñˆÊR¬ñE6¬Fóˆ∆ˆvññ◊ñÁFÛ¢""¬6ˆ◊VÊS¢""¬ñÊFó&óß¶Û¢""¬w5ì¢ÁV∆¬¬w5É¢ÁV∆¬¬6ˆFñ6U&Wß¶Û¢""¬&V¢""¬FóGFW6V7WG&ñ6S¢""¬Ê˜FS¢""¬7FFÛ¢$Dd$R"¬FˆÊS¢f«6R¬FˆÊTC¢ÁV∆¬¬FˆÊT'ì¢""¬ÁV÷W&ı&ˆw&W76ófÚ¬6ˆ÷÷W76ñC¢÷ÊvV÷VÁD6ˆ÷÷W76ñB¬7&VFVDC¢fó&V&6RÊfó&W7F˜&R‰fñV∆Ef«VRÁ6W'fW%Fñ÷W7F◊Çí¬WFFVDC¢fó&V&6RÊfó&W7F˜&R‰fñV∆Ef«VRÁ6W'fW%Fñ÷W7F◊Çí¬WFFVD'ïVñC¢7W'&VÁEW6W#ÚÁVñB«¬""¬WFFVD'îÊ÷S¢vWD˜W&F˜$Fó7∆îÊ÷RÇí”∞¢6ˆÁ7Bf∆ñFFñˆ‚“f∆ñFFT÷ÊvV÷VÁE∆ÁBáñ∆ˆB¬""¬ÁV÷W&ı&ˆw&W76ófÚì∞¢ñbáf∆ñFFñˆ‚í&WGW&‚∆W'Báf∆ñFFñˆ‚ì∞¢6ˆÁ7B&Vb“vóBF"Ê6ˆ∆∆V7Fñˆ‚ÜvWD6ˆ÷÷W76T6ˆ∆∆V7Fñˆ‰Ê÷RÇííÊFˆ2Ü÷ÊvV÷VÁD6ˆ÷÷W76ñBíÊ6ˆ∆∆V7Fñˆ‚Ç&ñ◊ñÁFí"íÊFBáñ∆ˆBì∞¢÷ÊvV÷VÁDVFóFñÊtñB“&VbÊñC∞ß–†¶gVÊ7Fñˆ‚WFFT÷ÊvV÷VÁD'V∆¥&"Çí∞¢6ˆÁ7B&"“Fˆ7V÷VÁBÊvWDV∆V÷VÁD'îñBÇ&ñ◊ñÁFí÷'V∆≤÷7FñˆÁ2"ì∞¢&#ÚÊ6∆74∆ó7BÁFˆvv∆RÇ&ÜñFFV‚"¬÷ÊvV÷VÁE6V∆V7FVDñG2Á6ó¶R«¬6‰÷ÊvTFFÇíì∞¢6ˆÁ7B6˜VÁFW"“Fˆ7V÷VÁBÊvWDV∆V÷VÁD'îñBÇ&ñ◊ñÁFí◊6V∆V7FVB÷6˜VÁB"ì∞¢ñbÜ6˜VÁFW"í6˜VÁFW"ÁFWáD6ˆÁFVÁB“G∂÷ÊvV÷VÁE6V∆V7FVDñG2Á6ó¶W“6V∆W¶ñˆÊFñ∞ß–†¶7ñÊ2gVÊ7Fñˆ‚'V‰÷ÊvV÷VÁD'V∆¥7Fñˆ‚Ü7Fñˆ‚í∞¢ñbÇ6‰÷ÊvTFFÇí«¬÷ÊvV÷VÁE6V∆V7FVDñG2Á6ó¶Rí&WGW&„∞¢6ˆÁ7BñG2“≤‚‚Ê÷ÊvV÷VÁE6V∆V7FVDñG5”∞¢6ˆÁ7B&Vb“F"Ê6ˆ∆∆V7Fñˆ‚ÜvWD6ˆ÷÷W76T6ˆ∆∆V7Fñˆ‰Ê÷RÇííÊFˆ2Ü÷ÊvV÷VÁD6ˆ÷÷W76ñBíÊ6ˆ∆∆V7Fñˆ‚Ç&ñ◊ñÁFí"ì∞¢ñbÜ7Fñˆ‚””“&FV∆WFR"í∞¢ñbÇ6ˆÊfó&“ÜV∆ñ÷ñÊ&RFVfñÊóFóf÷VÁFRG∂ñG2Ê∆VÊwFá“ñ◊ñÁFí6V∆W¶ñˆÊFìˆíí&WGW&„∞¢6ˆÁ7B&F6Ç“F"Ê&F6ÇÇì≤ñG2Êf˜$V6ÇÇÜñBí”‚&F6ÇÊFV∆WFRá&VbÊFˆ2ÜñBííì≤vóB&F6ÇÊ6ˆ÷÷óBÇì∞¢“V«6RñbÜ7Fñˆ‚””“&Wá˜'B"í∞¢&WGW&‚Wá˜'D∆ƒñ◊ñÁFï7FGW2ÜñG2ì∞¢“V«6R∞¢6ˆÁ7B∆&V«2“≤7FGW3¢$ÁV˜fÚ7FFÚ"¬˜W&F˜#¢$ÁV˜fÚ˜W&F˜&R"¬&V¢$ÁV˜f&V"¬6ˆ◊Áì¢$ÁV˜fFóGF"”∞¢6ˆÁ7Bf«VR“7G&ñÊrá&ˆ◊BÜG∂∆&V«5∂7FñˆÂ◊”¶¬7Fñˆ‚””“'7FGW2"Ú$Dd$R"¢""í«¬""íÁG&ñ“Çì≤ñbÇf«VRí&WGW&„∞¢6ˆÁ7BfñV∆B“≤7FGW3¢'7FFÚ"¬˜W&F˜#¢&˜W&F˜&R"¬&V¢&&V"¬6ˆ◊Áì¢&FóGFW6V7WG&ñ6R"’∂7FñˆÂ”∞¢6ˆÁ7B&F6Ç“F"Ê&F6ÇÇì≤ñG2Êf˜$V6ÇÇÜñBí”‚&F6ÇÁ6WBá&VbÊFˆ2ÜñBí¬≤∂fñV∆E”¢f«VR¬‚‚‚Ü7Fñˆ‚””“'7FGW2"Ú≤FˆÊS¢f«VRÁFıWW$66RÇí””“$dEDÚ"“¢∑“í¬WFFVDC¢fó&V&6RÊfó&W7F˜&R‰fñV∆Ef«VRÁ6W'fW%Fñ÷W7F◊Çí¬WFFVD'ïVñC¢7W'&VÁEW6W#ÚÁVñB«¬""¬WFFVD'îÊ÷S¢vWD˜W&F˜$Fó7∆îÊ÷RÇí“¬≤÷W&vS¢G'VR“íì≤vóB&F6ÇÊ6ˆ÷÷óBÇì∞¢–¢÷ÊvV÷VÁE6V∆V7FVDñG2Ê6∆V"Çì≤&VÊFW$ñ◊ñÁFî÷ÊvV÷VÁEF&∆RÇì∞ß–†¶gVÊ7Fñˆ‚7Gñ∆Tˆffñ6ñ≈v˜&∑6ÜVWBá6ÜVWBí∞¢6ÜVWBÁfñWw2“∑≤7FFS¢&g&˜¶V‚"¬ï7∆óC¢’”∞¢6ÜVWBÊWFÙfñ«FW"“≤g&ˆ”¢$"¬FÛ¢$Û"”∞¢6ÜVWBÊ6ˆ«V÷Á2“≥r¬R¬3"¬#"¬#¬3"¬R¬R¬Ç¬#2¬#B¬Ç¬b¬#B¬3E“Ê÷ÇávñGFÇí”‚á≤vñGFÇ“íì∞¢6ˆÁ7BÜVFW"“6ÜVWBÊvWE&˜rÉì∞¢ÜVFW"ÊÜVñváB“#É∞¢ÜVFW"ÊV6Ñ6V∆¬ÇÜ6V∆¬í”‚∞¢6V∆¬ÊfˆÁB“≤&ˆ∆C¢G'VR¬6ˆ∆˜#¢≤&v#¢$dddddddb"“”∞¢6V∆¬Êfñ∆¬“≤GóS¢'GFW&‚"¬GFW&„¢'6ˆ∆ñB"¬ft6ˆ∆˜#¢≤&v#¢$dcsTSSB"“”∞¢6V∆¬Ê∆ñvÊ÷VÁB“≤fW'Fñ6√¢&÷ñFF∆R"¬Ü˜&ó¶ˆÁF√¢&6VÁFW""”∞¢“ì∞ß–†¶7ñÊ2gVÊ7Fñˆ‚6fTˆffñ6ñ≈v˜&∂&ˆˆ≤áv˜&∂&ˆˆ≤¬fñ∆VÊ÷Rí∞¢6ˆÁ7B'VffW"“vóBv˜&∂&ˆˆ≤ÁÜ«7ÇÁw&óFT'VffW"Çì∞¢6ˆÁ7BW&¬“U$¬Ê7&VFTˆ&¶V7EU$¬ÜÊWr&∆ˆ"Ö∂'VffW%“¬≤GóS¢&∆ñ6Fñˆ‚˜fÊBÊ˜VÁÜ÷∆f˜&÷G2÷ˆffñ6VFˆ7V÷VÁBÁ7&VG6ÜVWF÷¬Á6ÜVWB"“íì∞¢6ˆÁ7B∆ñÊ≤“Fˆ7V÷VÁBÊ7&VFTV∆V÷VÁBÇ&"ì∞¢∆ñÊ≤Êá&Vb“W&√∞¢∆ñÊ≤ÊF˜vÊ∆ˆB“fñ∆VÊ÷S∞¢Fˆ7V÷VÁBÊ&ˆGíÊVÊD6Üñ∆BÜ∆ñÊ≤ì∞¢∆ñÊ≤Ê6∆ñ6≤Çì∞¢∆ñÊ≤Á&V÷˜fRÇì∞¢6WEFñ÷V˜WBÇÇí”‚U$¬Á&Wfˆ∂Tˆ&¶V7EU$¬áW&¬í¬ì∞ß–†¶7ñÊ2gVÊ7Fñˆ‚F˜vÊ∆ˆDˆffñ6ñƒñ◊ñÁFïFV◊∆FRÇí∞¢6ˆÁ7Bv˜&∂&ˆˆ≤“ÊWrWÜ6Vƒ•2Âv˜&∂&ˆˆ≤Çì∞¢v˜&∂&ˆˆ≤Ê7&VF˜"“$ÜW&#∞¢6ˆÁ7B6ÜVWB“v˜&∂&ˆˆ≤ÊFEv˜&∑6ÜVWBÇ$î’îÂDí"ì∞¢6ÜVWBÊFE&˜rÑÙddî4î≈Ùî’îÂDïÙ4Ù≈T‘Â2ì∞¢6ÜVWBÊFE&˜rÖ≥¬%4”"¬$ñ◊ñÁFÚFíW6V◊ñÚ"¬$FWW&F˜&R"¬$&ˆ∆ˆvÊ"¬%fñ&ˆ÷"¬CB„CìCí¬„3C#b¬$"¬$&VÊ˜&B"¬$FóGFW6V◊ñÚ"¬""¬""¬""¬%&ñ◊V˜fW&RVW7F&ñv&ñ÷FV∆¬vñ◊˜'F¶ñˆÊR%“ì∞¢7Gñ∆Tˆffñ6ñ≈v˜&∑6ÜVWBá6ÜVWBì∞¢6ˆÁ7BñÁ7G'V7FñˆÁ2“v˜&∂&ˆˆ≤ÊFEv˜&∑6ÜVWBÇ$ï5E%U§îÙ‰í"ì∞¢ñÁ7G'V7FñˆÁ2Ê6ˆ«V÷Á2“∑≤vñGFÉ¢#Ç“¬≤vñGFÉ¢É’”∞¢ñÁ7G'V7FñˆÁ2ÊFE&˜w2Ö∞¢≤$‘ÙDTƒƒÚTddî4îƒRÑU$"¬$Êˆ‚÷ˆFñfñ6&RíÊˆ÷íFV∆∆R6ˆ∆ˆÊÊRFV¬fˆv∆ñÚî’îÂDí‚%“¿¢≤$6◊íˆ&&∆ñvF˜&í"¬$‚‚RFVÊˆ÷ñÊ¶ñˆÊRñ◊ñÁFÚ‚îB4:Ç6ˆÁ6ñv∆ñFÚW"&ñ6ˆÊ˜66W&RíGW∆ñ6Fí‚%“¿¢≤$‚‚"¬$ÁV÷W&Ú&ˆw&W76ófÚ7F&ñ∆RRVÊófˆ6ÚÊV∆∆6ˆ÷÷W76‚%“¿¢≤$6ˆ˜&FñÊFR"¬$∆FóGVFñÊRG&”ìRì≤∆ˆÊvóGVFñÊRG&”ÉRÉ‚%“¿¢≤$6ˆ◊∆WF÷VÁFÚ"¬$FFW6V7W¶ñˆÊR¬˜&W6V7W¶ñˆÊRR˜W&F˜&RFWfˆÊÚ&W7F&RgV˜Fì¢6&ÊÊÚ6ˆ◊ñ∆Fí&V÷VÊFÚdEDÚ‚%“¿¢≤$f˜&÷Fíñ◊˜'F&ñ∆í"¬%Ñ≈5Ç¬Ñ≈2¬55bRÙE2‚%–¢“ì∞¢ñÁ7G'V7FñˆÁ2ÊvWE&˜rÉíÊfˆÁB“≤&ˆ∆C¢G'VR¬6ˆ∆˜#¢≤&v#¢$dddddddb"“”∞¢ñÁ7G'V7FñˆÁ2ÊvWE&˜rÉíÊfñ∆¬“≤GóS¢'GFW&‚"¬GFW&„¢'6ˆ∆ñB"¬ft6ˆ∆˜#¢≤&v#¢$dcsTSSB"“”∞¢vóB6fTˆffñ6ñ≈v˜&∂&ˆˆ≤áv˜&∂&ˆˆ≤¬&÷ˆFV∆∆ı˜Vffñ6ñ∆Uˆñ◊ñÁFíÁÜ«7Ç"ì∞ß–†¶gVÊ7Fñˆ‚f˜&÷E&ˆ÷TFˆÊU'G2áf«VRí∞¢6ˆÁ7B÷ñ∆∆ó2“fó&W7F˜&TFFUFÙ÷ñ∆∆ó2áf«VRì∞¢ñbÇ÷ñ∆∆ó2í&WGW&‚≤FFS¢""¬Fñ÷S¢""”∞¢6ˆÁ7BFFR“ÊWrFFRÜ÷ñ∆∆ó2ì∞¢&WGW&‚∞¢FFS¢ÊWrñÁF¬‰FFUFñ÷Tf˜&÷BÇ&óB‘ïB"¬≤Fñ÷U¶ˆÊS¢$WW&˜Rı&ˆ÷R"¬Fì¢#"÷FñvóB"¬÷ˆÁFÉ¢#"÷FñvóB"¬ñV#¢&ÁV÷W&ñ2"“íÊf˜&÷BÜFFRí¿¢Fñ÷S¢ÊWrñÁF¬‰FFUFñ÷Tf˜&÷BÇ&óB‘ïB"¬≤Fñ÷U¶ˆÊS¢$WW&˜Rı&ˆ÷R"¬Ü˜W#¢#"÷FñvóB"¬÷ñÁWFS¢#"÷FñvóB"¬Ü˜W##¢f«6R“íÊf˜&÷BÜFFRê¢”∞ß–†¶7ñÊ2gVÊ7Fñˆ‚Wá˜'D∆ƒñ◊ñÁFï7FGW2á6V∆V7FVDñG2“ÁV∆¬í∞¢6ˆÁ7B6ˆ÷÷W76ñB“vWEF&vWD6ˆ÷÷W76ñBÇì∞¢6ˆÁ7B6ˆ÷÷W76“6ˆ÷÷W76T'îñBÊvWBÜ6ˆ÷÷W76ñBì∞¢ñbÇ6ˆ÷÷W76ñB«¬6ˆ÷÷W76í&WGW&‚∆W'BÇ%6V∆W¶ñˆÊVÊ6ˆ÷÷W76‚"ì∞¢6ˆÁ7B6Ê6Ü˜B“vóBF"Ê6ˆ∆∆V7Fñˆ‚ÜvWD6ˆ÷÷W76T6ˆ∆∆V7Fñˆ‰Ê÷RÇííÊFˆ2Ü6ˆ÷÷W76ñBíÊ6ˆ∆∆V7Fñˆ‚Ç&ñ◊ñÁFí"íÊvWBÇì∞¢6ˆÁ7B6V∆V7FVB“'&íÊó4'&íá6V∆V7FVDñG2íÚÊWr6WBá6V∆V7FVDñG2í¢ÁV∆√∞¢6ˆÁ7B∆ÁG2“6Ê6Ü˜BÊFˆ72Ê÷ÇÜFˆ2í”‚á≤ñC¢Fˆ2ÊñB¬‚‚ÊFˆ2ÊFFÇí“ííÊfñ«FW"Çá∆ÁBí”‚6V∆V7FVB«¬6V∆V7FVBÊÜ2á∆ÁBÊñBííÁ6˜'BÇÜ¬"í”‚ÁV÷&W"ÜÊÁV÷W&ı&ˆw&W76ófÚ«¬Á6˜'D˜&FW"«¬í“ÁV÷&W"Ü"ÊÁV÷W&ı&ˆw&W76ófÚ«¬"Á6˜'D˜&FW"«¬íì∞¢6ˆÁ7Bv˜&∂&ˆˆ≤“ÊWrWÜ6Vƒ•2Âv˜&∂&ˆˆ≤Çì∞¢6ˆÁ7B6ÜVWB“v˜&∂&ˆˆ≤ÊFEv˜&∑6ÜVWBÇ$î’îÂDí"ì∞¢6ÜVWBÊFE&˜rÑÙddî4î≈Ùî’îÂDïÙ4Ù≈T‘Â2ì∞¢∆ÁG2Êf˜$V6ÇÇá∆ÁB¬ñÊFWÇí”‚∞¢6ˆÁ7BFˆÊR“∆ÁBÊFˆÊRÚf˜&÷E&ˆ÷TFˆÊU'G2á∆ÁBÊFˆÊTBí¢≤FFS¢""¬Fñ÷S¢""”∞¢6ÜVWBÊFE&˜rÖ∑∆ÁBÊÁV÷W&ı&ˆw&W76ófÚ«¬∆ÁBÁ6˜'D˜&FW"«¬ñÊFWÇ≤¬∆ÁBÊñE6«¬""¬∆ÁBÊFVÊˆ÷ñÊ¶ñˆÊR«¬""¬∆ÁBÁFóˆ∆ˆvññ◊ñÁFÚ«¬""¬∆ÁBÊ6ˆ◊VÊR«¬""¬∆ÁBÊñÊFó&óß¶Ú«¬""¬∆ÁBÊw5íÛÚ""¬∆ÁBÊw5ÇÛÚ""¬∆ÁBÊ6ˆFñ6U&Wß¶Ú«¬""¬∆ÁBÊ&V«¬""¬∆ÁBÊFóGFW6V7WG&ñ6R«¬""¬FˆÊRÊFFR¬FˆÊRÁFñ÷R¬∆ÁBÊFˆÊRÚá∆ÁBÊFˆÊT'í«¬""í¢""¬∆ÁBÊÊ˜FR«¬∆ÁBÊÊ˜FTñ◊ñÁFÚ«¬"%“ì∞¢“ì∞¢7Gñ∆Tˆffñ6ñ≈v˜&∑6ÜVWBá6ÜVWBì∞¢6ˆÁ7BñÁ7G'V7FñˆÁ2“v˜&∂&ˆˆ≤ÊFEv˜&∑6ÜVWBÇ$ï5E%U§îÙ‰í"ì∞¢ñÁ7G'V7FñˆÁ2ÊFE&˜rÖ≤$W7˜'F¶ñˆÊR7FFÚñ◊ñÁFí"¬6ˆ÷÷W76¢G∂6ˆ÷÷W76ÊÊˆ÷R«¬"'“(
"GWGFív∆íñ◊ñÁFí6ˆÊÚñÊ6«W6íÊ“ì∞¢ñÁ7G'V7FñˆÁ2Ê6ˆ«V÷Á2“∑≤vñGFÉ¢3"“¬≤vñGFÉ¢É’”∞¢vóB6fTˆffñ6ñ≈v˜&∂&ˆˆ≤áv˜&∂&ˆˆ≤¬7FFıˆñ◊ñÁFïÚGµ7G&ñÊrÜ6ˆ÷÷W76ÊÊˆ÷R«¬&6ˆ÷÷W76"íÁ&W∆6RÇıµÊ◊£”ï“≤ˆví¬%Ú"ó“ÁÜ«7Üì∞ß–†¶gVÊ7Fñˆ‚&VÊFW$6ˆ÷÷W76T÷ÊvV÷VÁD∆ó7BÇí∞¢ñbÇVíÊ6ˆ÷÷W76T÷ÊvT∆ó7Bí&WGW&„∞¢ñbÇ6‰÷ÊvTFFÇíí∞¢VíÊ6ˆ÷÷W76T÷ÊvT∆ó7BÊñÊÊW$ÖD‘¬“#«6∆73“v◊WFVBsÂ6ˆ∆Úv∆íF÷ñ‚˜76ˆÊÚ&ñÊˆ÷ñÊ&R¬7gV˜F&RÚV∆ñ÷ñÊ&R6ˆ÷÷W76R„¬˜‚#∞¢&WGW&„∞¢–¢6ˆÁ7BVW'í“7G&ñÊrÜFˆ7V÷VÁBÊvWDV∆V÷VÁD'îñBÇ&6ˆ÷÷W76R÷÷ÊvV÷VÁB◊6V&6Ç"ìÚÁf«VR«¬""íÁG&ñ“ÇíÁFÙ∆ˆ6∆T∆˜vW$66RÇ&óB"ì∞¢6ˆÁ7B6ˆ÷÷W76R“6˜'D6ˆ÷÷W76T'î7&VFVDDFW62Ñ'&íÊg&ˆ“Ü6ˆ÷÷W76T'îñBÁf«VW2ÇíííÊfñ«FW"ÇÜ6ˆ÷÷W76í”‚Ä¢VW'í«¬G∂6ˆ÷÷W76ÊÊˆ÷R«¬"'“G∂6ˆ÷÷W76Ê6ˆFñ6R«¬"'÷ÁFÙ∆ˆ6∆T∆˜vW$66RÇ&óB"íÊñÊ6«VFW2áVW'íê¢íì∞¢ñbÇ6ˆ÷÷W76RÊ∆VÊwFÇí∞¢VíÊ6ˆ÷÷W76T÷ÊvT∆ó7BÊñÊÊW$ÖD‘¬“#«6∆73“v◊WFVBs‰ÊW77VÊ6ˆ÷÷W76Fó7ˆÊñ&ñ∆R„¬˜‚#∞¢&WGW&„∞¢–¢VíÊ6ˆ÷÷W76T÷ÊvT∆ó7BÊñÊÊW$ÖD‘¬“"#∞¢6ˆ÷÷W76RÊf˜$V6ÇÇÜ6ˆ÷÷W76í”‚∞¢6ˆÁ7B&˜r“Fˆ7V÷VÁBÊ7&VFTV∆V÷VÁBÇ&Fób"ì∞¢&˜rÊ6∆74Ê÷R“&6ˆ÷÷W76÷÷ÊvR÷6&B#∞¢6ˆÁ7BñÊfÚ“Fˆ7V÷VÁBÊ7&VFTV∆V÷VÁBÇ&Fób"ì∞¢ñÊfÚÊ6∆74Ê÷R“&6ˆ÷÷W76÷÷ÊvR÷ñÊfÚ#∞¢6ˆÁ7BFóF∆R“Fˆ7V÷VÁBÊ7&VFTV∆V÷VÁBÇ'7G&ˆÊr"ì∞¢6ˆÁ7B6ˆFñ6T6ˆ÷÷W76“7G&ñÊrÜ6ˆ÷÷W76Ê6ˆFñ6R«¬""íÁG&ñ“Çì∞¢6ˆÁ7BÜ57V&6ˆ÷÷W76R“vWE7V&6ˆ÷÷W76RÜ6ˆ÷÷W76ÊñBíÊ∆VÊwFÇ‚∞¢6ˆÁ7BF˜F¬“ÁV÷&W"ÜvWD6ˆ÷÷W767FG2Ü6ˆ÷÷W76ÊñBíÁF˜F¬«¬ì∞¢FóF∆RÊñÊÊW$ÖD‘¬“G∂W66TÖD‘¬Ü6ˆ÷÷W76ÊÊˆ÷R«¬$6ˆ÷÷W766VÁ¶Êˆ÷R"ó“G∂Ü57V&6ˆ÷÷W76RÚ«7‚6∆73“&6ˆ÷÷W76◊&VÁB÷ñÊFñ6F˜""FóF∆S“$6ˆÁFñVÊR7V&6ˆ÷÷W76R#Ô	˘8#¬˜7„Ê¢"'÷∞¢ñÊfÚÊVÊD6Üñ∆BáFóF∆Rì∞¢6ˆÁ7B÷WF“Fˆ7V÷VÁBÊ7&VFTV∆V÷VÁBÇ'"ì∞¢÷WFÊ6∆74Ê÷R“&6ˆ÷÷W76÷÷ÊvR÷÷WF#∞¢÷WFÊñÊÊW$ÖD‘¬“«7„‰6ˆB‚G∂W66TÖD‘¬Ü6ˆFñ6T6ˆ÷÷W76«¬.(	B"ó”¬˜7„„«7„‚G∑F˜F«“G∑F˜F¬””“Ú&ñ◊ñÁFÚ"¢&ñ◊ñÁFí'”¬˜7„Ê∞¢ñÊfÚÊVÊD6Üñ∆BÜ÷WFì∞†¢6ˆÁ7B7FñˆÁ2“Fˆ7V÷VÁBÊ7&VFTV∆V÷VÁBÇ&Fób"ì∞¢7FñˆÁ2Ê6∆74Ê÷R“&6ˆ÷÷W76÷6&B÷7FñˆÁ2#∞¢6ˆÁ7B˜V‚“7&VFT'WGFˆ‚Ç$$í"¬Çí”‚˜V‰ñ◊ñÁFî÷ÊvV÷VÁBÜ6ˆ÷÷W76íì∞¢˜V‚Ê6∆74∆ó7BÊFBÇ&'F‚◊&ñ÷'í"ì∞¢6ˆÁ7B÷VÁR“Fˆ7V÷VÁBÊ7&VFTV∆V÷VÁBÇ&FWFñ«2"ì∞¢÷VÁRÊ6∆74Ê÷R“&6ˆ÷÷W76÷7FñˆÁ2÷÷VÁR#∞¢÷VÁRÊñÊÊW$ÖD‘¬“«7V÷÷'í&ñ÷∆&V√“$¶ñˆÊíW"G∂W66TÖD‘¬Ü6ˆ÷÷W76ÊÊˆ÷R«¬&6ˆ÷÷W76"ó“#Ó(∫„¬˜7V÷÷'ì„∆Fóc„¬ˆFócÊ∞¢6ˆÁ7B÷VÁT&ˆGí“÷VÁRÁVW'ï6V∆V7F˜"Ç&Fób"ì∞¢÷VÁT&ˆGíÊVÊD6Üñ∆BÜ7&VFT'WGFˆ‚Ç$÷ˆFñfñ66ˆ÷÷W76"¬Çí”‚&VÊ÷T6ˆ÷÷W76Ü6ˆ÷÷W76ÊñB¬6ˆ÷÷W76ÊÊˆ÷R«¬$6ˆ÷÷W76"¬6ˆ÷÷W76Ê6ˆFñ6R«¬""ííì∞¢÷VÁT&ˆGíÊVÊD6Üñ∆BÜ7&VFT'WGFˆ‚Ç$vW7Fó66íñ◊ñÁFí"¬Çí”‚˜V‰ñ◊ñÁFî÷ÊvV÷VÁBÜ6ˆ÷÷W76ííì∞¢÷VÁT&ˆGíÊVÊD6Üñ∆BÜ7&VFT'WGFˆ‚Ç$vW7Fó66í&Wß¶ñ&ñÚ"¬7ñÊ2Çí”‚≤vóB˜V‰ñ◊ñÁFî÷ÊvV÷VÁBÜ6ˆ÷÷W76ì≤vñÊF˜r‰66˜VÁFñÊuc#ÚÊ˜VÂ&ñ6W2Çì≤“íì∞¢÷VÁT&ˆGíÊVÊD6Üñ∆BÜ7&VFT'WGFˆ‚Ç$ñ◊˜'FFFí"¬Çí”‚˜V‰ñ◊ñÁFî÷ÊvV÷VÁBÜ6ˆ÷÷W76ííì∞¢÷VÁT&ˆGíÊVÊD6Üñ∆BÜ7&VFT'WGFˆ‚Ç$W7˜'FFFí"¬7ñÊ2Çí”‚≤vóB˜V‰ñ◊ñÁFî÷ÊvV÷VÁBÜ6ˆ÷÷W76ì≤vñÊF˜r‰66˜VÁFñÊuc#ÚÊWá˜'Ev˜&∂&ˆˆ≤Çì≤“íì∞¢÷VÁT&ˆGíÊVÊD6Üñ∆BÜ7&VFT'WGFˆ‚Ç%7gV˜F6ˆ÷÷W76"¬Çí”‚vñÊF˜r‰66˜VÁFñÊuc"ÚvñÊF˜r‰66˜VÁFñÊuc"Ê˜V‰6∆V"Ü6ˆ÷÷W76í¢6∆V$6ˆ÷÷W76ñ◊ñÁFíÜ6ˆ÷÷W76ÊñB¬6ˆ÷÷W76ÊÊˆ÷R«¬$6ˆ÷÷W76"ííì∞¢÷VÁT&ˆGíÊVÊD6Üñ∆BÜ7&VFT'WGFˆ‚Ç$V∆ñ÷ñÊ6ˆ÷÷W76"¬Çí”‚FV∆WFT6ˆ÷÷W76Ü6ˆ÷÷W76ÊñB¬6ˆ÷÷W76ÊÊˆ÷R«¬$6ˆ÷÷W76"ííì∞¢7FñˆÁ2ÊVÊBÜ˜V‚¬÷VÁRì∞†¢&˜rÊVÊD6Üñ∆BÜñÊfÚì∞¢&˜rÊVÊD6Üñ∆BÜ7FñˆÁ2ì∞¢VíÊ6ˆ÷÷W76T÷ÊvT∆ó7BÊVÊD6Üñ∆Bá&˜rì∞¢“ì∞ß–†¶7ñÊ2gVÊ7Fñˆ‚&VÊ÷T6ˆ÷÷W76Ü6ˆ÷÷W76ñB¬7W'&VÁDÊ÷R¬7W'&VÁD6ˆFR“""í∞¢ñbÇ6‰÷ÊvTFFÇíí∞¢∆W'BÇ%6ˆ∆ÚV‚F÷ñ‚\;"&ñÊˆ÷ñÊ&R6ˆ÷÷W76R‚"ì∞¢&WGW&„∞¢–¢6ˆÁ7BÊWáDÊ÷R“vñÊF˜rÁ&ˆ◊BÇ$ÁV˜fÚÊˆ÷R6ˆ÷÷W76¢"¬7W'&VÁDÊ÷R«¬""ì∞¢ñbÜÊWáDÊ÷R”“ÁV∆¬í&WGW&„∞¢6ˆÁ7BÊWáD6ˆFR“vñÊF˜rÁ&ˆ◊BÇ$6ˆFñ6R6ˆ÷÷W76¢"¬7W'&VÁD6ˆFR«¬""ì∞¢ñbÜÊWáD6ˆFR”“ÁV∆¬í&WGW&„∞¢6ˆÁ7BÊ˜&÷∆ó¶VB“ÊWáDÊ÷RÁG&ñ“Çì∞¢6ˆÁ7BÊ˜&÷∆ó¶VD6ˆFR“7G&ñÊrÜÊWáD6ˆFR«¬""íÁG&ñ“Çì∞¢ñbÇÊ˜&÷∆ó¶VBí&WGW&„∞¢vóBF"Ê6ˆ∆∆V7Fñˆ‚ÜvWD6ˆ÷÷W76T6ˆ∆∆V7Fñˆ‰Ê÷RÇííÊFˆ2Ü6ˆ÷÷W76ñBíÁ6WBá≤Êˆ÷S¢Ê˜&÷∆ó¶VB¬6ˆFñ6S¢Ê˜&÷∆ó¶VD6ˆFR“¬≤÷W&vS¢G'VR“ì∞¢ñbá6V∆V7FVD6ˆ÷÷W76ñB””“6ˆ÷÷W76ñBí∞¢6V∆V7FVD6ˆ÷÷W76Ê÷R“Ê˜&÷∆ó¶VC∞¢VíÊ6ˆ÷÷W76GFófÁFWáD6ˆÁFVÁB“Ê˜&÷∆ó¶VD6ˆFRÚ6ˆ÷÷W766V∆W¶ñˆÊF¢G∂Ê˜&÷∆ó¶VG“(
"6ˆB‚6ˆ÷÷W76¢G∂Ê˜&÷∆ó¶VD6ˆFW÷¢6ˆ÷÷W766V∆W¶ñˆÊF¢G∂Ê˜&÷∆ó¶VG÷∞¢WFFT6ˆ÷÷W766ˆÁFWáETíÇì∞¢–ß–†¶7ñÊ2gVÊ7Fñˆ‚6∆V$6ˆ÷÷W76ñ◊ñÁFíÜ6ˆ÷÷W76ñB¬Êˆ÷Rí∞¢ñbÇ6‰÷ÊvTFFÇíí∞¢∆W'BÇ%6ˆ∆ÚV‚F÷ñ‚\;"7gV˜F&R6ˆ÷÷W76R‚"ì∞¢&WGW&„∞¢–¢6ˆÁ7Bˆ≤“vñÊF˜rÊ6ˆÊfó&“Ä¢7FíV∆ñ÷ñÊÊFÚGWGFív∆íñ◊ñÁFíFV∆∆6ˆ÷÷W76"G∂Êˆ÷W“"Â∆Â∆‰¬FW&÷ñÊRfW',:GFófFÚWFˆ÷Fñ6÷VÁFRñ¬ÁV˜fÚ÷ˆFV∆∆ÚWÜ6V¬Â∆Â∆‰˜W&¶ñˆÊRó'&WfW'6ñ&ñ∆RÂ∆Â∆Â&V÷íÙ≤W"7gV˜FRGFófÁV˜fÚ÷ˆFV∆∆ÚÊ ¢ì∞¢ñbÇˆ≤í&WGW&„∞¢6ˆÁ7Bñ◊ñÁFï&Vb“F"Ê6ˆ∆∆V7Fñˆ‚Ç&6ˆ÷÷W76R"íÊFˆ2Ü6ˆ÷÷W76ñBíÊ6ˆ∆∆V7Fñˆ‚Ç&ñ◊ñÁFí"ì∞¢vóBFV∆WFT6ˆ∆∆V7Fñˆ‰Fˆ72Üñ◊ñÁFï&Vbì∞¢vóBF"Ê6ˆ∆∆V7Fñˆ‚ÜvWD6ˆ÷÷W76T6ˆ∆∆V7Fñˆ‰Ê÷RÇííÊFˆ2Ü6ˆ÷÷W76ñBíÁ6WBá∞¢WÜ6Vƒ÷ˆFV≈fW'6ñˆ„¢"¿¢WÜ6Vƒ÷ˆFVƒ7FófFVDC¢fó&V&6RÊfó&W7F˜&R‰fñV∆Ef«VRÁ6W'fW%Fñ÷W7F◊Çí¿¢WÜ6Vƒ÷ˆFVƒ7FófFVD'ì¢WFÇÊ7W'&VÁEW6W#ÚÊV÷ñ¬«¬""¿¢ÊWáDñ◊ñÁFÙÁV÷&W#¢¢“¬≤÷W&vS¢G'VR“ì∞¢6ˆÁ7B6ˆ÷÷W76“6ˆ÷÷W76T'îñBÊvWBÜ6ˆ÷÷W76ñBí«¬≤ñC¢6ˆ÷÷W76ñB¬Êˆ÷R”∞¢˜V‰ñ◊ñÁFî÷ÊvV÷VÁBÜ6ˆ÷÷W76ì∞¢∆W'BÇ$6ˆ÷÷W76vvñ˜&ÊF6ˆ‚7V66W76Ú¬ÁV˜fÚ÷ˆFV∆∆ÚWÜ6V¬‚"ì∞ß–†¶7ñÊ2gVÊ7Fñˆ‚FV∆WFT6ˆ∆∆V7Fñˆ‰Fˆ72Ü6ˆ∆∆V7FñˆÂ&Vb¬&F6Ö6ó¶R“#í∞¢vÜñ∆RáG'VRí∞¢6ˆÁ7B6Ê6Ü˜B“vóB6ˆ∆∆V7FñˆÂ&VbÊ∆ñ÷óBÜ&F6Ö6ó¶RíÊvWBÇì∞¢ñbá6Ê6Ü˜BÊV◊Gíí'&V≥∞¢6ˆÁ7B&F6Ç“F"Ê&F6ÇÇì∞¢6Ê6Ü˜BÊFˆ72Êf˜$V6ÇÇÜFˆ2í”‚&F6ÇÊFV∆WFRÜFˆ2Á&Vbíì∞¢vóB&F6ÇÊ6ˆ÷÷óBÇì∞¢–ß–†¶gVÊ7Fñˆ‚&VÊFW$F÷ñÂW6W'2Çí∞¢ñbÇVíÊF÷ñÂW6W'4∆ó7Bí&WGW&„∞¢ñbÇ6‰÷ÊvTFFÇíí∞¢VíÊF÷ñÂW6W'4∆ó7BÊñÊÊW$ÖD‘¬“#«6∆73“v◊WFVBsÂ6ˆ∆ÚV‚F÷ñ‚\;"vW7Fó&RíW&÷W76íF÷ñ‚„¬˜‚#∞¢&WGW&„∞¢–¢6ˆÁ7BV÷ñ«2“'&íÊg&ˆ“ÜF÷ñ‰V÷ñ«2íÁ6˜'BÇÜ¬"í”‚Ê∆ˆ6∆T6ˆ◊&RÜ"¬&óB"íì∞¢VíÊF÷ñÂW6W'4∆ó7BÊñÊÊW$ÖD‘¬“"#∞¢V÷ñ«2Êf˜$V6ÇÇÜV÷ñ¬í”‚∞¢6ˆÁ7B&˜r“Fˆ7V÷VÁBÊ7&VFTV∆V÷VÁBÇ&Fób"ì∞¢&˜rÊ6∆74Ê÷R“'6ñ◊∆R÷∆ó7B÷óFV“#∞¢6ˆÁ7B∆&V¬“Fˆ7V÷VÁBÊ7&VFTV∆V÷VÁBÇ'7‚"ì∞¢∆&V¬ÁFWáD6ˆÁFVÁB“V÷ñ√∞¢&˜rÊVÊD6Üñ∆BÜ∆&V¬ì∞¢ñbÇó4'Vñ«DñÂ7WW$F÷ñ‰V÷ñ¬ÜV÷ñ¬íí∞¢6ˆÁ7B&Wfˆ∂T'F‚“7&VFT'WGFˆ‚Ç%&ñ◊V˜fí"¬Çí”‚&V÷˜fTF÷ñ‰V÷ñ¬ÜV÷ñ¬íì∞¢&˜rÊVÊD6Üñ∆Bá&Wfˆ∂T'F‚ì∞¢–¢VíÊF÷ñÂW6W'4∆ó7BÊVÊD6Üñ∆Bá&˜rì∞¢“ì∞ß–†¶7ñÊ2gVÊ7Fñˆ‚FDF÷ñÂW6W$'îV÷ñ¬ÜWfVÁBí∞¢WfVÁBÁ&WfVÁDFVfV«BÇì∞¢ñbÇ6‰÷ÊvTFFÇíí∞¢∆W'BÇ%6ˆ∆ÚV‚F÷ñ‚\;"vvóVÊvW&R«G&íF÷ñ‚‚"ì∞¢&WGW&„∞¢–¢6ˆÁ7BV÷ñ¬“Ê˜&÷∆ó¶TV÷ñ¬áVíÊF÷ñÂW6W$V÷ñ¬Áf«VRì∞¢ñbÇV÷ñ¬«¬V÷ñ¬ÊñÊ6«VFW2Ç$"íí∞¢∆W'BÇ$ñÁ6W&ó66íV‚vV÷ñ¬f∆ñF‚"ì∞¢&WGW&„∞¢–¢6ˆÁ7BÊWáB“'&íÊg&ˆ“ÜÊWr6WBÖ≤‚‚ÊF÷ñ‰V÷ñ«2¬V÷ñ≈“íì∞¢vóBF"Ê6ˆ∆∆V7Fñˆ‚Ç&6ˆÊfñr"íÊFˆ2Ç&F÷ñÂW6W'2"íÁ6WBá∞¢V÷ñ«3¢ÊWáB¿¢WFFVDC¢fó&V&6RÊfó&W7F˜&R‰fñV∆Ef«VRÁ6W'fW%Fñ÷W7F◊Çí¿¢WFFVD'ì¢7W'&VÁEW6W#ÚÊV÷ñ¬«¬" ¢“¬≤÷W&vS¢G'VR“ì∞¢VíÊF÷ñÂW6W$f˜&“Á&W6WBÇì∞ß–†¶7ñÊ2gVÊ7Fñˆ‚&V÷˜fTF÷ñ‰V÷ñ¬ÜV÷ñ¬í∞¢ñbÇ6‰÷ÊvTFFÇíí∞¢∆W'BÇ%6ˆ∆ÚV‚F÷ñ‚\;"&ñ◊V˜fW&RF÷ñ‚‚"ì∞¢&WGW&„∞¢–¢6ˆÁ7BÊ˜&÷∆ó¶VB“Ê˜&÷∆ó¶TV÷ñ¬ÜV÷ñ¬ì∞¢ñbÇÊ˜&÷∆ó¶VB«¬ó4'Vñ«DñÂ7WW$F÷ñ‰V÷ñ¬ÜÊ˜&÷∆ó¶VBíí&WGW&„∞¢6ˆÁ7Bˆ≤“vñÊF˜rÊ6ˆÊfó&“Ü&ñ◊V˜fW&RíW&÷W76íF÷ñ‚W"G∂Ê˜&÷∆ó¶VG”ˆì∞¢ñbÇˆ≤í&WGW&„∞¢6ˆÁ7BÊWáB“'&íÊg&ˆ“ÜF÷ñ‰V÷ñ«2íÊfñ«FW"ÇÜóFV“í”‚óFV“”“Ê˜&÷∆ó¶VBì∞¢vóBF"Ê6ˆ∆∆V7Fñˆ‚Ç&6ˆÊfñr"íÊFˆ2Ç&F÷ñÂW6W'2"íÁ6WBá∞¢V÷ñ«3¢ÊWáB¿¢WFFVDC¢fó&V&6RÊfó&W7F˜&R‰fñV∆Ef«VRÁ6W'fW%Fñ÷W7F◊Çí¿¢WFFVD'ì¢7W'&VÁEW6W#ÚÊV÷ñ¬«¬" ¢“¬≤÷W&vS¢G'VR“ì∞ß–†¶gVÊ7Fñˆ‚7F˜6ÜE7V'67&óFñˆ‚Çí∞¢ñbáVÁ7V'67&ñ&T6ÜBí∞¢VÁ7V'67&ñ&T6ÜBÇì∞¢VÁ7V'67&ñ&T6ÜB“ÁV∆√∞¢–¢6ÜD÷W76vW2“µ”∞¢6ÜDÊ˜Fñfñ6FñˆÁ4ñÊóFñ∆ó¶VB“f«6S∞ß–†¶7ñÊ2gVÊ7Fñˆ‚&VÊFW$6ÜBÜ÷W76vW2í∞¢ñbÇ7W'&VÁEW6W"í∞¢VíÊ6ÜD6˜VÁFW"Ê6∆74∆ó7BÊFBÇ&ÜñFFV‚"ì∞¢&VÊFW$6ÜDV◊Gï7FFRÇ$fí∆ˆvñ‚W"W6&R∆6ÜB‚"ì∞¢VíÊ6ÜE6VÊD'F‚ÊFó6&∆VB“G'VS∞¢VíÊ6ÜE&V6óñVÁBÊFó6&∆VB“G'VS∞¢VíÊ6ÜEFWáBÊFó6&∆VB“G'VS∞¢VíÊ6ÜD÷VFññÁWBÊFó6&∆VB“G'VS∞¢VíÊ6ÜEfˆñ6T'F‚ÊFó6&∆VB“G'VS∞¢&WGW&„∞¢–†¢VíÊ6ÜE6VÊD'F‚ÊFó6&∆VB“f«6S∞¢VíÊ6ÜE&V6óñVÁBÊFó6&∆VB“f«6S∞¢VíÊ6ÜEFWáBÊFó6&∆VB“f«6S∞¢VíÊ6ÜD÷VFññÁWBÊFó6&∆VB“f«6S∞¢VíÊ6ÜEfˆñ6T'F‚ÊFó6&∆VB“f«6S∞†¢6ˆÁ7Bfó6ñ&∆T÷W76vW2“÷W76vW2Êfñ«FW"ÇÜ÷W76vRí”‚6ÂfñWt÷W76vRÜ÷W76vRíbbó46ÜD÷W76vTg&W6ÇÜ÷W76vRíì∞†¢ñbÇfó6ñ&∆T÷W76vW2Ê∆VÊwFÇí∞¢VíÊ6ÜD6˜VÁFW"Ê6∆74∆ó7BÊFBÇ&ÜñFFV‚"ì∞¢&VÊFW$6ÜDV◊Gï7FFRÇì∞¢&WGW&„∞¢–†¢6ˆÁ7BVÁ&VD6˜VÁB“6˜VÁEVÁ&VD÷W76vW2áfó6ñ&∆T÷W76vW2ì∞¢ñbáVÁ&VD6˜VÁB‚í∞¢VíÊ6ÜD6˜VÁFW"Ê6∆74∆ó7BÁ&V÷˜fRÇ&ÜñFFV‚"ì∞¢VíÊ6ÜD6˜VÁFW"ÁFWáD6ˆÁFVÁB“VÁ&VD6˜VÁB‚ìíÚ#ìí≤"¢7G&ñÊráVÁ&VD6˜VÁBì∞¢“V«6R∞¢VíÊ6ÜD6˜VÁFW"Ê6∆74∆ó7BÊFBÇ&ÜñFFV‚"ì∞¢–†¢VíÊ6ÜDgV∆ƒ∆ó7BÊñÊÊW$ÖD‘¬“"#∞¢6ˆÁ7B÷W76vTV∆V÷VÁG2“vóB&ˆ÷ó6RÊ∆¬áfó6ñ&∆T÷W76vW2Ê÷ÇÜ÷W76vRí”‚7&VFT6ÜD÷W76vTV∆V÷VÁBÜ÷W76vRííì∞¢÷W76vTV∆V÷VÁG2Êf˜$V6ÇÇÜV∆V÷VÁBí”‚∞¢ñbÜV∆V÷VÁBíVíÊ6ÜDgV∆ƒ∆ó7BÊVÊD6Üñ∆BÜV∆V÷VÁBì∞¢“ì∞¢VíÊ6ÜDgV∆ƒ∆ó7BÁ67&ˆ∆≈F˜“VíÊ6ÜDgV∆ƒ∆ó7BÁ67&ˆ∆ƒÜVñváC∞†¢ñbÇVíÊ6ÜD÷ˆF¬Ê6∆74∆ó7BÊ6ˆÁFñÁ2Ç&ÜñFFV‚"íí∞¢÷&¥6ÜD5&VBÇì∞¢–ß–†¶gVÊ7Fñˆ‚&VÊFW$6ÜDV◊Gï7FFRÜ÷W76vR“$ÊW77V‚÷W76vvñÚ&W6VÁFR"í∞¢ñbÇVíÊ6ÜDgV∆ƒ∆ó7Bí&WGW&„∞¢VíÊ6ÜDgV∆ƒ∆ó7BÊñÊÊW$ÖD‘¬“ ¢∆Fób6∆73“&6ÜB÷V◊Gí◊7FFR#‡¢∆Fób6∆73“&6ÜB÷V◊Gí÷ñ6ˆ‚"&ñ÷ÜñFFV„“'G'VR#Ô	˘*√¬ˆFóc‡¢«‚G∂W66TÖD‘¬Ü÷W76vRó”¬˜‡¢¬ˆFóc‡¢∞ß–†¶gVÊ7Fñˆ‚˜V‰6ÜD6∆V$6ˆÊfó&‘÷ˆF¬Çí∞¢ñbÇ6‰÷ÊvTFFÇíí&WGW&„∞¢VíÊ6ÜD6∆V$6ˆÊfó&‘÷ˆF√ÚÊ6∆74∆ó7BÁ&V÷˜fRÇ&ÜñFFV‚"ì∞¢VíÊ6ÜD6∆V$6ˆÊfó&‘÷ˆF√ÚÁ6WDGG&ñ'WFRÇ&&ñ÷ÜñFFV‚"¬&f«6R"ì∞¢VíÊ6ÜD6∆V$6ˆÊfó&‘'F„ÚÊfˆ7W2Çì∞ß–†¶gVÊ7Fñˆ‚6∆˜6T6ÜD6∆V$6ˆÊfó&‘÷ˆF¬Çí∞¢VíÊ6ÜD6∆V$6ˆÊfó&‘÷ˆF√ÚÊ6∆74∆ó7BÊFBÇ&ÜñFFV‚"ì∞¢VíÊ6ÜD6∆V$6ˆÊfó&‘÷ˆF√ÚÁ6WDGG&ñ'WFRÇ&&ñ÷ÜñFFV‚"¬'G'VR"ì∞ß–†¶7ñÊ2gVÊ7Fñˆ‚6∆V$7W'&VÁD6ÜD÷W76vW2Çí∞¢ñbÇ6‰÷ÊvTFFÇíí∞¢∆W'BÇ%6ˆ∆ÚV‚F÷ñ‚\;"7gV˜F&R∆6ÜB‚"ì∞¢&WGW&„∞¢–†¢6∆˜6T6ÜD6∆V$6ˆÊfó&‘÷ˆF¬Çì∞¢6ˆÁ7Bfó6ñ&∆T÷W76vTñG2“6ÜD÷W76vW0¢Êfñ«FW"ÇÜ÷W76vRí”‚6ÂfñWt÷W76vRÜ÷W76vRíê¢Ê÷ÇÜ÷W76vRí”‚7G&ñÊrÜ÷W76vRÊñB«¬""íÁG&ñ“Çíê¢Êfñ«FW"Ñ&ˆˆ∆V‚ì∞†¢VíÊ6ÜD6∆V$'F‚ÊFó6&∆VB“G'VS∞¢VíÊ6ÜD6∆V$6ˆÊfó&‘'F‚ÊFó6&∆VB“G'VS∞¢VíÊ6ÜDfVVF&6≤ÁFWáD6ˆÁFVÁB“fó6ñ&∆T÷W76vTñG2Ê∆VÊwFÇÚ%7gV˜F÷VÁFÚ6ÜBñ‚6˜'6Ú‚‚‚"¢$6ÜBvú:gV˜F‚#∞†¢G'í∞¢vóBÊñ÷FT6ÜDFV∆WFñˆ‚Çì∞¢6ÜD÷W76vW2“6ÜD÷W76vW2Êfñ«FW"ÇÜ÷W76vRí”‚fó6ñ&∆T÷W76vTñG2ÊñÊ6«VFW2Ö7G&ñÊrÜ÷W76vRÊñB«¬""ííì∞¢vóB&VÊFW$6ÜBÜ6ÜD÷W76vW2ì∞†¢ñbáfó6ñ&∆T÷W76vTñG2Ê∆VÊwFÇí∞¢vóBFV∆WFT6ÜD÷W76vW4'îñG2áfó6ñ&∆T÷W76vTñG2ì∞¢VíÊ6ÜDfVVF&6≤ÁFWáD6ˆÁFVÁB“$6ÜB7gV˜FF‚#∞¢“V«6R∞¢&VÊFW$6ÜDV◊Gï7FFRÇì∞¢–¢“6F6ÇÜW'&˜"í∞¢6ˆÁ6ˆ∆RÊW'&˜"Ç$W'&˜&R7gV˜F÷VÁFÚ6ÜC¢"¬W'&˜"ì∞¢VíÊ6ÜDfVVF&6≤ÁFWáD6ˆÁFVÁB“W'&˜#ÚÊ÷W76vR«¬$ñ◊˜76ñ&ñ∆R7gV˜F&R∆6ÜB‚#∞¢“fñÊ∆«í∞¢VíÊ6ÜD6∆V$6ˆÊfó&‘'F‚ÊFó6&∆VB“f«6S∞¢VíÊ6ÜD6∆V$'F‚ÊFó6&∆VB“6‰÷ÊvTFFÇì∞¢–ß–†¶gVÊ7Fñˆ‚Êñ÷FT6ÜDFV∆WFñˆ‚Çí∞¢6ˆÁ7B÷W76vTÊˆFW2“'&íÊg&ˆ“áVíÊ6ÜDgV∆ƒ∆ó7CÚÁVW'ï6V∆V7F˜$∆¬Ç"Ê6ÜB÷÷W76vR"í«¬µ“ì∞¢ñbÇ÷W76vTÊˆFW2Ê∆VÊwFÇí&WGW&‚&ˆ÷ó6RÁ&W6ˆ«fRÇì∞¢÷W76vTÊˆFW2Êf˜$V6ÇÇÜÊˆFR¬ñÊFWÇí”‚∞¢ÊˆFRÁ7Gñ∆RÁ6WE&˜W'GíÇ"“÷6ÜB÷FV∆WFR÷FV∆í"¬G¥÷FÇÊ÷ñ‚ÜñÊFWÇ¢#B¬Éó÷◊6ì∞¢ÊˆFRÊ6∆74∆ó7BÊFBÇ&ó2÷FV∆WFñÊr"ì∞¢“ì∞¢&WGW&‚ÊWr&ˆ÷ó6RÇá&W6ˆ«fRí”‚vñÊF˜rÁ6WEFñ÷V˜WBá&W6ˆ«fR¬3cíì∞ß–†¶7ñÊ2gVÊ7Fñˆ‚FV∆WFT6ÜD÷W76vW4'îñG2Ü÷W76vTñG2í∞¢6ˆÁ7BVÊóVTñG2“'&íÊg&ˆ“ÜÊWr6WBÜ÷W76vTñG2íì∞¢f˜"Ü∆WBñÊFWÇ“≤ñÊFWÇ¬VÊóVTñG2Ê∆VÊwFÉ≤ñÊFWÇ≥“CSí∞¢6ˆÁ7B&F6Ç“F"Ê&F6ÇÇì∞¢VÊóVTñG2Á6∆ñ6RÜñÊFWÇ¬ñÊFWÇ≤CSíÊf˜$V6ÇÇÜ÷W76vTñBí”‚∞¢&F6ÇÊFV∆WFRÜF"Ê6ˆ∆∆V7Fñˆ‚Ç&6ÜD÷W76vW2"íÊFˆ2Ü÷W76vTñBíì∞¢“ì∞¢vóB&F6ÇÊ6ˆ÷÷óBÇì∞¢–ß–†¶gVÊ7Fñˆ‚6˜VÁEVÁ&VD÷W76vW2Ü÷W76vW2í∞¢&WGW&‚÷W76vW2Êfñ«FW"ÇÜ÷W76vRí”‚∞¢ñbÜó4˜v‰÷W76vRÜ÷W76vRíí&WGW&‚f«6S∞¢6ˆÁ7B7&VFVDB“÷W76vRÊ7&VFVDBbb÷W76vRÊ7&VFVDBÁFÙFFP¢Ú÷W76vRÊ7&VFVDBÁFÙFFRÇíÊvWEFñ÷RÇê¢¢∞¢&WGW&‚∆7E&VD6ÜDB«¬7&VFVDB‚∆7E&VD6ÜDC∞¢“íÊ∆VÊwFÉ∞ß–†¶gVÊ7Fñˆ‚6ÂfñWt÷W76vRÜ÷W76vRí∞¢ñbÇ7W'&VÁEW6W"í&WGW&‚f«6S∞¢6ˆÁ7B÷WFFFGóR“7G&ñÊrÜ÷W76vSÚÊ÷WFFFÚÁGóR«¬""ì∞¢ñbÜ÷WFFFGóR””“&Ê˜Fñfñ6FñˆÂˆ6≤"bb6‰÷ÊvTFFÇíí&WGW&‚f«6S∞¢ñbÇ÷W76vRÁ&V6óñVÁDñBí&WGW&‚G'VS∞¢&WGW&‚÷W76vRÁ&V6óñVÁDñB””“7W'&VÁEW6W"ÁVñB«¬÷W76vRÁ6VÊFW$ñB””“7W'&VÁEW6W"ÁVñC∞ß–†¶gVÊ7Fñˆ‚÷&¥6ÜD5&VBÇí∞¢ñbÇ6ÜD÷W76vW2Ê∆VÊwFÇí∞¢VíÊ6ÜD6˜VÁFW"Ê6∆74∆ó7BÊFBÇ&ÜñFFV‚"ì∞¢&WGW&„∞¢–†¢6ˆÁ7B∆FW7D÷W76vR“6ÜD÷W76vW5∂6ÜD÷W76vW2Ê∆VÊwFÇ“”∞¢6ˆÁ7B7&VFVDB“∆FW7D÷W76vRÊ7&VFVDBbb∆FW7D÷W76vRÊ7&VFVDBÁFÙFFP¢Ú∆FW7D÷W76vRÊ7&VFVDBÁFÙFFRÇíÊvWEFñ÷RÇê¢¢FFRÊÊ˜rÇì∞†¢∆7E&VD6ÜDB“7&VFVDC∞¢VíÊ6ÜD6˜VÁFW"Ê6∆74∆ó7BÊFBÇ&ÜñFFV‚"ì∞ß–†¶gVÊ7Fñˆ‚6‰6ˆÊfó&‘Ü˜W'4g&ˆ‘6ÜBÜ÷W76vRí∞¢ñbÖ7G&ñÊrÜ÷W76vSÚÊ∂ñÊB«¬""í”“'7ó7FV“"í&WGW&‚f«6S∞¢6ˆÁ7B÷WFFF“÷W76vSÚÊ÷WFFFbbGóVˆb÷W76vRÊ÷WFFF””“&ˆ&¶V7B"Ú÷W76vRÊ÷WFFF¢ÁV∆√∞¢ñbÇ÷WFFFí&WGW&‚f«6S∞¢ñbÜ÷WFFFÁGóR”“&Ü˜W'5ˆ&˜f¬"«¬÷WFFFÊ&˜f≈&WVW7DñBí&WGW&‚f«6S∞¢ñbÜ÷WFFFÊ7Fñˆ‚””“&∆WfV√ˆˆ≤"í&WGW&‚G'VS∞¢ñbÜ÷WFFFÊ7Fñˆ‚””“&F÷ñÂˆfñÊ≈ˆˆ≤"í&WGW&‚6‰÷ÊvTFFÇì∞¢&WGW&‚f«6S∞ß–†¶gVÊ7Fñˆ‚6‰˜V‰Ü˜W'4g&ˆ‘6ÜD∆W'BÜ÷W76vRí∞¢ñbÖ7G&ñÊrÜ÷W76vSÚÊ∂ñÊB«¬""í”“'7ó7FV“"í&WGW&‚f«6S∞¢ñbÇ6‰÷ÊvTFFÇíí&WGW&‚f«6S∞¢6ˆÁ7B÷WFFF“÷W76vSÚÊ÷WFFFbbGóVˆb÷W76vRÊ÷WFFF””“&ˆ&¶V7B"Ú÷W76vRÊ÷WFFF¢ÁV∆√∞¢ñbÇ÷WFFFí&WGW&‚f«6S∞¢&WGW&‚÷WFFFÁGóR””“&Ü˜W'5ˆFVF∆ñÊUˆ∆W'B"bb÷WFFFÊ7Fñˆ‚””“&˜VÂˆÜ˜W'2#∞ß–†¶gVÊ7Fñˆ‚6‰÷˜fTñ◊ñÁFÙFˆÊTg&ˆ‘6ÜBÜ÷W76vRí∞¢ñbÖ7G&ñÊrÜ÷W76vSÚÊ∂ñÊB«¬""í”“'7ó7FV“"í&WGW&‚f«6S∞¢ñbÇ6‰÷ÊvTFFÇíí&WGW&‚f«6S∞¢6ˆÁ7B÷WFFF“÷W76vSÚÊ÷WFFFbbGóVˆb÷W76vRÊ÷WFFF””“&ˆ&¶V7B"Ú÷W76vRÊ÷WFFF¢ÁV∆√∞¢ñbÇ÷WFFFí&WGW&‚f«6S∞¢ñbÜ÷WFFFÁGóR”“&ñ◊ñÁFıˆFˆÊU˜&V6˜fW'í"«¬÷WFFFÊ7Fñˆ‚”“&÷˜fUˆFˆÊR"í&WGW&‚f«6S∞¢6ˆÁ7B6ˆ÷÷W76ñB“7G&ñÊrÜ÷WFFFÊ6ˆ÷÷W76ñB«¬""íÁG&ñ“Çì∞¢6ˆÁ7Bñ◊ñÁFÙñG2“'&íÊó4'&íÜ÷WFFFÊñ◊ñÁFÙñG2íÚ÷WFFFÊñ◊ñÁFÙñG2Êfñ«FW"Ñ&ˆˆ∆V‚í¢µ”∞¢&WGW&‚&ˆˆ∆V‚Ü6ˆ÷÷W76ñBbbñ◊ñÁFÙñG2Ê∆VÊwFÇì∞ß–†¶7ñÊ2gVÊ7Fñˆ‚÷˜fTñ◊ñÁFıFÙFˆÊTg&ˆ‘6ÜBÜ÷W76vRí∞¢6ˆÁ7B÷WFFF“÷W76vSÚÊ÷WFFFbbGóVˆb÷W76vRÊ÷WFFF””“&ˆ&¶V7B"Ú÷W76vRÊ÷WFFF¢ÁV∆√∞¢ñbÇ÷WFFFíFá&˜rÊWrW'&˜"Ç$÷W76vvñÚ6ÜB6VÁ¶÷WFFFíf∆ñFí‚"ì∞¢6ˆÁ7B6ˆ÷÷W76ñB“7G&ñÊrÜ÷WFFFÊ6ˆ÷÷W76ñB«¬""íÁG&ñ“Çì∞¢6ˆÁ7B6ˆ÷÷W76Ê÷R“7G&ñÊrÜ÷WFFFÊ6ˆ÷÷W76Ê÷R«¬""íÁG&ñ“Çí«¬$6ˆ÷÷W76#∞¢6ˆÁ7Bñ◊ñÁFÙñG2“'&íÊó4'&íÜ÷WFFFÊñ◊ñÁFÙñG2íÚ÷WFFFÊñ◊ñÁFÙñG2Ê÷ÇÜñBí”‚7G&ñÊrÜñB«¬""íÁG&ñ“ÇííÊfñ«FW"Ñ&ˆˆ∆V‚í¢µ”∞¢ñbÇ6ˆ÷÷W76ñB«¬ñ◊ñÁFÙñG2Ê∆VÊwFÇíFá&˜rÊWrW'&˜"Ç$FFíñ◊ñÁFÚñÁ7Vffñ6ñVÁFíW"∆Ú7˜7F÷VÁFÚÊVídEDí‚"ì∞¢vóB6WDñ◊ñÁFÙFˆÊRÜ6ˆ÷÷W76ñB¬ñ◊ñÁFÙñG2¬G'VRì∞¢ñbá6V∆V7FVD6ˆ÷÷W76ñB””“6ˆ÷÷W76ñBí∞¢WFFTñ◊ñÁFÙ∆ˆ6≈7FFRÜñ◊ñÁFÙñG2¬∞¢FˆÊS¢G'VR¿¢FˆÊTC¢ÊWrFFRÇí¿¢FˆÊT'ì¢7W'&VÁEW6W#ÚÊFó7∆îÊ÷R«¬7W'&VÁEW6W#ÚÊV÷ñ¬«¬$F÷ñ‚ ¢“ì∞¢–¢66ÜVGV∆T6ˆ÷÷W766ÜVWE7ñÊ2Ü6ˆ÷÷W76ñB¬6ˆ÷÷W76Ê÷R¬#Sì∞ß–†¶7ñÊ2gVÊ7Fñˆ‚FV∆WFT6ÜD÷W76vT'îñBÜ÷W76vTñBí∞¢6ˆÁ7BÊ˜&÷∆ó¶VB“7G&ñÊrÜ÷W76vTñB«¬""íÁG&ñ“Çì∞¢ñbÇÊ˜&÷∆ó¶VBí&WGW&„∞¢vóBF"Ê6ˆ∆∆V7Fñˆ‚Ç&6ÜD÷W76vW2"íÊFˆ2ÜÊ˜&÷∆ó¶VBíÊFV∆WFRÇì∞ß–†¶7ñÊ2gVÊ7Fñˆ‚vWDÜ˜W'4&˜f≈&WVW7D'îñBá&WVW7DñBí∞¢6ˆÁ7BÊ˜&÷∆ó¶VE&WVW7DñB“7G&ñÊrá&WVW7DñB«¬""íÁG&ñ“Çì∞¢ñbÇÊ˜&÷∆ó¶VE&WVW7DñBí&WGW&‚ÁV∆√∞¢6ˆÁ7BFˆ56Ê“vóBF"Ê6ˆ∆∆V7Fñˆ‚ÜvWD˜&T&˜f≈&WVW7G46ˆ∆∆V7Fñˆ‰Ê÷RÇííÊFˆ2ÜÊ˜&÷∆ó¶VE&WVW7DñBíÊvWBÇì∞¢ñbÇFˆ56ÊÊWÜó7G2í&WGW&‚ÁV∆√∞¢&WGW&‚≤ñC¢Fˆ56ÊÊñB¬‚‚ÊFˆ56ÊÊFFÇí”∞ß–†¶gVÊ7Fñˆ‚6WD6ÜDÜ˜W'47Fñˆ‰'WGFˆÁ57FFRÜ66WD'WGFˆ‚¬&V¶V7D'WGFˆ‚¬7FFRí∞¢6ˆÁ7B&˜FÑ'WGFˆÁ2“∂66WD'WGFˆ‚¬&V¶V7D'WGFˆÂ“Êfñ«FW"Ñ&ˆˆ∆V‚ì∞¢&˜FÑ'WGFˆÁ2Êf˜$V6ÇÇÜ'WGFˆ‚í”‚∞¢'WGFˆ‚ÊFó6&∆VB“7FFS∞¢“ì∞ß–†¶7ñÊ2gVÊ7Fñˆ‚'Vñ∆D6ÜD÷W76vT7FñˆÁ2Ü÷W76vRí∞¢ñbÜ6‰÷˜fTñ◊ñÁFÙFˆÊTg&ˆ‘6ÜBÜ÷W76vRíí∞¢6ˆÁ7B7FñˆÁ2“Fˆ7V÷VÁBÊ7&VFTV∆V÷VÁBÇ&Fób"ì∞¢7FñˆÁ2Ê6∆74Ê÷R“&óFV“÷7FñˆÁ2#∞¢6ˆÁ7B÷˜fTFˆÊT'F‚“7&VFT'WGFˆ‚Ç%7˜7FÊVídEDí"¬7ñÊ2Çí”‚∞¢÷˜fTFˆÊT'F‚ÊFó6&∆VB“G'VS∞¢G'í∞¢vóB÷˜fTñ◊ñÁFıFÙFˆÊTg&ˆ‘6ÜBÜ÷W76vRì∞¢VíÊ6ÜDfVVF&6≤ÁFWáD6ˆÁFVÁB“$ñ◊ñÁFÚ7˜7FFÚÊVídEDí‚#∞¢“6F6ÇÜW'&˜"í∞¢6ˆÁ6ˆ∆RÊW'&˜"Ç$W'&˜&R7˜7F÷VÁFÚñ◊ñÁFÚÊVídEDíF6ÜC¢"¬W'&˜"ì∞¢VíÊ6ÜDfVVF&6≤ÁFWáD6ˆÁFVÁB“W'&˜#ÚÊ÷W76vR«¬$ñ◊˜76ñ&ñ∆R7˜7F&R¬vñ◊ñÁFÚÊVídEDí‚#∞¢÷˜fTFˆÊT'F‚ÊFó6&∆VB“f«6S∞¢–¢“ì∞¢7FñˆÁ2ÊVÊD6Üñ∆BÜ÷˜fTFˆÊT'F‚ì∞¢&WGW&‚7FñˆÁ3∞¢–†¢ñbÜ6‰˜V‰Ü˜W'4g&ˆ‘6ÜD∆W'BÜ÷W76vRíí∞¢6ˆÁ7B7FñˆÁ2“Fˆ7V÷VÁBÊ7&VFTV∆V÷VÁBÇ&Fób"ì∞¢7FñˆÁ2Ê6∆74Ê÷R“&óFV“÷7FñˆÁ2#∞¢6ˆÁ7B˜V‰Ü˜W'4'F‚“7&VFT'WGFˆ‚Ç$ñÁ6W&ó66í˜&R"¬Çí”‚∞¢ñbÇ7W'&VÁEW6W"í&WGW&„∞¢˜V‰Ü˜W'5vRÇì∞¢ñbáVíÊÜ˜W'4FFRbb÷W76vSÚÊ÷WFFFÚÊFFRíVíÊÜ˜W'4FFRÁf«VR“÷W76vRÊ÷WFFFÊFFS∞¢ñbáVíÊÜ˜W'4fVVF&6≤í∞¢6ˆÁ7B6ˆ÷÷W76Ê÷R“7G&ñÊrÜ÷W76vSÚÊ÷WFFFÚÊ6ˆ÷÷W76Ê÷R«¬""íÁG&ñ“Çì∞¢VíÊÜ˜W'4fVVF&6≤ÁFWáD6ˆÁFVÁB“6ˆ÷÷W76Ê÷P¢Úgfó6Ú˜&S¢6ˆ◊ñ∆∆6ˆ÷÷W76G∂6ˆ÷÷W76Ê÷W“Ê ¢¢$gfó6Ú˜&S¢6ˆ◊ñ∆∆R˜&R÷Ê6ÁFí‚#∞¢–¢“ì∞¢7FñˆÁ2ÊVÊD6Üñ∆BÜ˜V‰Ü˜W'4'F‚ì∞¢ñbÜ6‰÷ÊvTFFÇíí∞¢6ˆÁ7BFV∆WFT'F‚“7&VFT'WGFˆ‚Ç$V∆ñ÷ñÊ"¬7ñÊ2Çí”‚∞¢G'í∞¢vóBFV∆WFT6ÜD÷W76vT'îñBÜ÷W76vRÊñBì∞¢VíÊ6ÜDfVVF&6≤ÁFWáD6ˆÁFVÁB“$÷W76vvñÚgfó6ÚV∆ñ÷ñÊFÚ‚#∞¢“6F6ÇÜW'&˜"í∞¢6ˆÁ6ˆ∆RÊW'&˜"Ç$W'&˜&RV∆ñ÷ñÊ¶ñˆÊR÷W76vvñÚgfó6Û¢"¬W'&˜"ì∞¢VíÊ6ÜDfVVF&6≤ÁFWáD6ˆÁFVÁB“$ñ◊˜76ñ&ñ∆RV∆ñ÷ñÊ&Rñ¬÷W76vvñÚgfó6Ú‚#∞¢–¢“ì∞¢7FñˆÁ2ÊVÊD6Üñ∆BÜFV∆WFT'F‚ì∞¢–¢&WGW&‚7FñˆÁ3∞¢–†¢ñbÇ6‰6ˆÊfó&‘Ü˜W'4g&ˆ‘6ÜBÜ÷W76vRíí&WGW&‚ÁV∆√∞¢6ˆÁ7B7FñˆÂGóR“7G&ñÊrÜ÷W76vSÚÊ÷WFFFÚÊ7Fñˆ‚«¬""íÁG&ñ“Çì∞¢6ˆÁ7B&˜f≈&WVW7DñB“7G&ñÊrÜ÷W76vSÚÊ÷WFFFÚÊ&˜f≈&WVW7DñB«¬""íÁG&ñ“Çì∞¢6ˆÁ7B&W6ˆ«fVE&WVW7B“vóBvWDÜ˜W'4&˜f≈&WVW7D'îñBÜ&˜f≈&WVW7DñBì∞¢6ˆÁ7B7FñˆÁ2“Fˆ7V÷VÁBÊ7&VFTV∆V÷VÁBÇ&Fób"ì∞¢7FñˆÁ2Ê6∆74Ê÷R“&óFV“÷7FñˆÁ2#∞¢ñbÇ&W6ˆ«fVE&WVW7Bí∞¢7FñˆÁ2ÊVÊD6Üñ∆BÜ7&VFT'WGFˆ‚Ç%&ñ6ÜñW7FÊˆ‚G&˜fF"¬Çí”‚∑“¬G'VRíì∞¢&WGW&‚7FñˆÁ3∞¢–†¢6ˆÁ7BWáV7FVE7FGW2“7FñˆÂGóR””“&∆WfV√ˆˆ≤"Ú'VÊFñÊuˆ∆WfV√"¢'VÊFñÊuˆF÷ñ‚#∞¢6ˆÁ7B6‰7B“7FñˆÂGóR””“&∆WfV√ˆˆ≤ ¢Ú6‰&˜fTÜ˜W'4∆WfV√á&W6ˆ«fVE&WVW7Bê¢¢6‰÷ÊvTFFÇì∞†¢ñbá&W6ˆ«fVE&WVW7BÁ7FGW2”“WáV7FVE7FGW2í∞¢6ˆÁ7B7FGW4÷“∞¢VÊFñÊuˆ∆WfV√¢$ñ‚GFW6&ñ÷ÚÙ≤"¿¢VÊFñÊuˆF÷ñ„¢$ñ‚GFW6F÷ñ‚fñÊ∆R"¿¢&˜fVC¢$vú:&˜fF"¿¢&V¶V7FVC¢$vú:&ñfóWFF ¢”∞¢7FñˆÁ2ÊVÊD6Üñ∆BÜ7&VFT'WGFˆ‚á7FGW4÷∑&W6ˆ«fVE&WVW7BÁ7FGW5“«¬$vú:vW7FóF"¬Çí”‚∑“¬G'VRíì∞¢&WGW&‚7FñˆÁ3∞¢–¢ñbÇ6‰7Bí&WGW&‚ÁV∆√∞†¢6ˆÁ7B66WD'WGFˆ‚“7&VFT'WGFˆ‚Ç$66WGF"¬7ñÊ2Çí”‚∞¢6WD6ÜDÜ˜W'47Fñˆ‰'WGFˆÁ57FFRÜ66WD'WGFˆ‚¬&V¶V7D'WGFˆ‚¬G'VRì∞¢G'í∞¢ñbÜ7FñˆÂGóR””“&∆WfV√ˆˆ≤"í∞¢vóB&˜fTÜ˜W'5&WVW7D∆WfV√á&W6ˆ«fVE&WVW7Bì∞¢VíÊ6ÜDfVVF&6≤ÁFWáD6ˆÁFVÁB“%&ñ÷ÚÙ≤&Vvó7G&FÚ‚#∞¢“V«6R∞¢vóB&˜fTÜ˜W'5&WVW7Dg&ˆ‘6ÜBÜ&˜f≈&WVW7DñBì∞¢VíÊ6ÜDfVVF&6≤ÁFWáD6ˆÁFVÁB“$˜&R&ó7V«FFR6ˆÊfW&÷FR‚#∞¢–¢“6F6ÇÜW'&˜"í∞¢6ˆÁ6ˆ∆RÊW'&˜"Ç$W'&˜&R6ˆÊfW&÷˜&RF6ÜC¢"¬W'&˜"ì∞¢6ˆÁ7B∆FW7E7FFR“vóBvWDÜ˜W'4&˜f≈&WVW7D'îñBÜ&˜f≈&WVW7DñBì∞¢ñbÇ∆FW7E7FFR«¬∆FW7E7FFRÁ7FGW2”“WáV7FVE7FGW2í∞¢VíÊ6ÜDfVVF&6≤ÁFWáD6ˆÁFVÁB“%&ñ6ÜñW7Fvú:vW7FóF‚#∞¢6WD6ÜDÜ˜W'47Fñˆ‰'WGFˆÁ57FFRÜ66WD'WGFˆ‚¬&V¶V7D'WGFˆ‚¬G'VRì∞¢&WGW&„∞¢–¢VíÊ6ÜDfVVF&6≤ÁFWáD6ˆÁFVÁB“W'&˜#ÚÊ÷W76vR«¬$W'&˜&RGW&ÁFR∆6ˆÊfW&÷˜&RF∆∆6ÜB‚#∞¢6WD6ÜDÜ˜W'47Fñˆ‰'WGFˆÁ57FFRÜ66WD'WGFˆ‚¬&V¶V7D'WGFˆ‚¬f«6Rì∞¢–¢“ì∞†¢6ˆÁ7B&V¶V7D'WGFˆ‚“7&VFT'WGFˆ‚Ç%&ñfóWF"¬7ñÊ2Çí”‚∞¢6WD6ÜDÜ˜W'47Fñˆ‰'WGFˆÁ57FFRÜ66WD'WGFˆ‚¬&V¶V7D'WGFˆ‚¬G'VRì∞¢G'í∞¢ñbÜ7FñˆÂGóR””“&∆WfV√ˆˆ≤"í∞¢vóB&V¶V7DÜ˜W'5&WVW7Bá&W6ˆ«fVE&WVW7Bì∞¢“V«6R∞¢vóB&V¶V7DÜ˜W'5&WVW7Dg&ˆ‘6ÜBÜ&˜f≈&WVW7DñBì∞¢–¢VíÊ6ÜDfVVF&6≤ÁFWáD6ˆÁFVÁB“%&ñ6ÜñW7F˜&R&ñfóWFF‚#∞¢“6F6ÇÜW'&˜"í∞¢6ˆÁ6ˆ∆RÊW'&˜"Ç$W'&˜&R&ñfóWFÚ˜&RF6ÜC¢"¬W'&˜"ì∞¢6ˆÁ7B∆FW7E7FFR“vóBvWDÜ˜W'4&˜f≈&WVW7D'îñBÜ&˜f≈&WVW7DñBì∞¢ñbÇ∆FW7E7FFR«¬∆FW7E7FFRÁ7FGW2”“WáV7FVE7FGW2í∞¢VíÊ6ÜDfVVF&6≤ÁFWáD6ˆÁFVÁB“%&ñ6ÜñW7Fvú:vW7FóF‚#∞¢6WD6ÜDÜ˜W'47Fñˆ‰'WGFˆÁ57FFRÜ66WD'WGFˆ‚¬&V¶V7D'WGFˆ‚¬G'VRì∞¢&WGW&„∞¢–¢VíÊ6ÜDfVVF&6≤ÁFWáD6ˆÁFVÁB“W'&˜#ÚÊ÷W76vR«¬$W'&˜&RGW&ÁFRñ¬&ñfóWFÚ˜&RF∆∆6ÜB‚#∞¢6WD6ÜDÜ˜W'47Fñˆ‰'WGFˆÁ57FFRÜ66WD'WGFˆ‚¬&V¶V7D'WGFˆ‚¬f«6Rì∞¢–¢“ì∞¢7FñˆÁ2ÊVÊD6Üñ∆BÜ66WD'WGFˆ‚ì∞¢7FñˆÁ2ÊVÊD6Üñ∆Bá&V¶V7D'WGFˆ‚ì∞¢&WGW&‚7FñˆÁ3∞ß–†¶7ñÊ2gVÊ7Fñˆ‚7&VFT6ÜD÷W76vTV∆V÷VÁBÜ÷W76vRí∞¢6ˆÁ7BóFV““Fˆ7V÷VÁBÊ7&VFTV∆V÷VÁBÇ&'Fñ6∆R"ì∞¢óFV“Ê6∆74Ê÷R“&6ÜB÷÷W76vR"≤Üó4˜v‰÷W76vRÜ÷W76vRíÚ"˜v‚"¢""ì∞¢6ˆÁ7Bó4ñÊ6ˆ÷ñÊu&ófFR“&ˆˆ∆V‚Ü÷W76vRÁ&V6óñVÁDñBbbó4˜v‰÷W76vRÜ÷W76vRíì∞†¢6ˆÁ7B7&VFVDB“÷W76vRÊ7&VFVDBbb÷W76vRÊ7&VFVDBÁFÙFFP¢Ú÷W76vRÊ7&VFVDBÁFÙFFRÇê¢¢ÊWrFFRÇì∞†¢6ˆÁ7BF˜“Fˆ7V÷VÁBÊ7&VFTV∆V÷VÁBÇ&Fób"ì∞¢F˜Ê6∆74Ê÷R“&6ÜB÷÷W76vR◊F˜#∞¢F˜ÊñÊÊW$ÖD‘¬“ ¢«7G&ˆÊs‚G∂W66TÖD‘¬Ü÷W76vRÁ6VÊFW$Ê÷R«¬$˜W&F˜&R"ó”¬˜7G&ˆÊs‡¢«7„‚G∂7&VFVDBÁFÙ∆ˆ6∆U7G&ñÊrÇ&óB‘ïB"ó”¬˜7„‡¢∞¢óFV“ÊVÊD6Üñ∆BáF˜ì∞†¢ñbÜ÷W76vRÁ&V6óñVÁDñBí∞¢6ˆÁ7BFr“Fˆ7V÷VÁBÊ7&VFTV∆V÷VÁBÇ'"ì∞¢FrÊ6∆74Ê÷R“&6ÜB◊GóR÷&FvR#∞¢FrÁFWáD6ˆÁFVÁB“ó4˜v‰÷W76vRÜ÷W76vRíÚ/	˘:í÷W76vvñÚ&ófFÚ"¢/	˘I"&ófFÚW"FR#∞¢óFV“ÊVÊD6Üñ∆BáFrì∞¢–†¢6ˆÁ7B6ˆÁFVÁEw&“Fˆ7V÷VÁBÊ7&VFTV∆V÷VÁBÇ&Fób"ì∞¢6ˆÁFVÁEw&Ê6∆74Ê÷R“ó4ñÊ6ˆ÷ñÊu&ófFRÚ&6ÜB◊&ófFR÷6ˆÁFVÁB"¢&6ÜB÷÷W76vR÷6ˆÁFVÁB#∞†¢6ˆÁ7B÷W76vUFWáB“GóVˆb÷W76vRÁFWáB””“'7G&ñÊr"bb÷W76vRÁFWáBÁG&ñ“Çê¢Ú÷W76vRÁFWá@¢¢GóVˆb÷W76vRÊ÷W76vR””“'7G&ñÊr"bb÷W76vRÊ÷W76vRÁG&ñ“Çê¢Ú÷W76vRÊ÷W76vP¢¢GóVˆb÷W76vRÊ&ˆGí””“'7G&ñÊr"bb÷W76vRÊ&ˆGíÁG&ñ“Çê¢Ú÷W76vRÊ&ˆGê¢¢GóVˆb÷W76vRÊ6ˆÁFVÁB””“'7G&ñÊr"bb÷W76vRÊ6ˆÁFVÁBÁG&ñ“Çê¢Ú÷W76vRÊ6ˆÁFVÁ@¢¢"#∞†¢ñbÇÜ÷W76vRÁGóR””“'FWáB"«¬Ç÷W76vRÁGóRbb÷W76vUFWáBííbb÷W76vUFWáBí∞¢6ˆÁ7B“Fˆ7V÷VÁBÊ7&VFTV∆V÷VÁBÇ'"ì∞¢Ê6∆74Ê÷R“&6ÜB◊FWáB#∞¢ÁFWáD6ˆÁFVÁB“÷W76vUFWáC∞¢6ˆÁFVÁEw&ÊVÊD6Üñ∆Báì∞¢–†¢6ˆÁ7B÷VFñ6˜W&6R“÷W76vRÊ÷VFñW&¬«¬÷W76vRÊ÷VFñFFW&¬«¬"#∞¢6ˆÁ7B÷VFñ÷ñ÷UGóR“7G&ñÊrÜ÷W76vRÊ÷VFñ÷ñ÷UGóR«¬""íÁFÙ∆˜vW$66RÇì∞¢6ˆÁ7BÜ4ñ÷vT÷VFñ“÷W76vRÁGóR””“&ñ÷vR"«¬÷VFñ÷ñ÷UGóRÁ7F'G5vóFÇÇ&ñ÷vRÚ"ì∞¢6ˆÁ7BÜ5fñFVÙ÷VFñ“÷W76vRÁGóR””“'fñFVÚ"«¬÷VFñ÷ñ÷UGóRÁ7F'G5vóFÇÇ'fñFVÚÚ"ì∞¢6ˆÁ7BÜ5fˆñ6T÷VFñ“÷W76vRÁGóR””“'fˆñ6R"«¬÷VFñ÷ñ÷UGóRÁ7F'G5vóFÇÇ&VFñÚÚ"ì∞†¢ñbÜÜ4ñ÷vT÷VFñbb÷VFñ6˜W&6Rí∞¢6ˆÁ7Bñ÷r“Fˆ7V÷VÁBÊ7&VFTV∆V÷VÁBÇ&ñ÷r"ì∞¢ñ÷rÊ6∆74Ê÷R“&6ÜB÷÷VFñ◊&WfñWr#∞¢ñ÷rÁ7&2“÷VFñ6˜W&6S∞¢ñ÷rÊ«B“$ñ÷÷vñÊRñÁfñFñ‚6ÜB#∞¢6ˆÁFVÁEw&ÊVÊD6Üñ∆BÜñ÷rì∞¢–†¢ñbÜÜ5fñFVÙ÷VFñbb÷VFñ6˜W&6Rí∞¢6ˆÁ7BfñFVÚ“Fˆ7V÷VÁBÊ7&VFTV∆V÷VÁBÇ'fñFVÚ"ì∞¢fñFVÚÊ6∆74Ê÷R“&6ÜB÷÷VFñ◊&WfñWr#∞¢fñFVÚÁ7&2“÷VFñ6˜W&6S∞¢fñFVÚÊ6ˆÁG&ˆ«2“G'VS∞¢6ˆÁFVÁEw&ÊVÊD6Üñ∆BáfñFVÚì∞¢–†¢ñbÜÜ5fˆñ6T÷VFñbb÷VFñ6˜W&6Rí∞¢6ˆÁ7BVFñÚ“Fˆ7V÷VÁBÊ7&VFTV∆V÷VÁBÇ&VFñÚ"ì∞¢VFñÚÁ7&2“÷VFñ6˜W&6S∞¢VFñÚÊ6ˆÁG&ˆ«2“G'VS∞¢VFñÚÊ6∆74Ê÷R“&6ÜB÷VFñÚ#∞¢6ˆÁFVÁEw&ÊVÊD6Üñ∆BÜVFñÚì∞¢–†¢6ˆÁ7BvV%fñWt∆ñÊ≤“7G&ñÊrÜ÷W76vRÊ÷VFñG&ófUvV%fñWt∆ñÊ≤«¬""íÁG&ñ“Çì∞¢ñbÇ÷VFñ6˜W&6RbbvV%fñWt∆ñÊ≤í∞¢6ˆÁ7B˜V‰∆ñÊ≤“Fˆ7V÷VÁBÊ7&VFTV∆V÷VÁBÇ&"ì∞¢˜V‰∆ñÊ≤Ê6∆74Ê÷R“&'F‚#∞¢˜V‰∆ñÊ≤Êá&Vb“vV%fñWt∆ñÊ≥∞¢˜V‰∆ñÊ≤ÁF&vWB“%ˆ&∆Ê≤#∞¢˜V‰∆ñÊ≤Á&V¬“&Êˆ˜VÊW"Ê˜&VfW'&W"#∞¢˜V‰∆ñÊ≤ÁFWáD6ˆÁFVÁB“$&í∆∆VvFÚ#∞¢6ˆÁFVÁEw&ÊVÊD6Üñ∆BÜ˜V‰∆ñÊ≤ì∞¢–†¢6ˆÁ7BÜ46ˆÁFVÁB“6ˆÁFVÁEw&Ê6Üñ∆DV∆V÷VÁD6˜VÁB‚∞¢ñbÜó4ñÊ6ˆ÷ñÊu&ófFRí∞¢6ˆÁ7BFˆvv∆T'F‚“Fˆ7V÷VÁBÊ7&VFTV∆V÷VÁBÇ&'WGFˆ‚"ì∞¢Fˆvv∆T'F‚ÁGóR“&'WGFˆ‚#∞¢Fˆvv∆T'F‚Ê6∆74Ê÷R“&'F‚'F‚÷6ÜB◊&ófFR◊Fˆvv∆R#∞¢Fˆvv∆T'F‚ÁFWáD6ˆÁFVÁB“$&í÷W76vvñÚ#∞¢Fˆvv∆T'F‚ÊFDWfVÁD∆ó7FVÊW"Ç&6∆ñ6≤"¬Çí”‚∞¢6ˆÁ7B˜VÊVB“6ˆÁFVÁEw&Ê6∆74∆ó7BÁFˆvv∆RÇ&ó2÷˜V‚"ì∞¢Fˆvv∆T'F‚ÁFWáD6ˆÁFVÁB“˜VÊVBÚ$6ÜóVFí÷W76vvñÚ"¢$&í÷W76vvñÚ#∞¢“ì∞¢óFV“ÊVÊD6Üñ∆BáFˆvv∆T'F‚ì∞¢ñbÜÜ46ˆÁFVÁBí∞¢óFV“ÊVÊD6Üñ∆BÜ6ˆÁFVÁEw&ì∞¢“V«6R∞¢6ˆÁ7Bf∆∆&6≤“Fˆ7V÷VÁBÊ7&VFTV∆V÷VÁBÇ'"ì∞¢f∆∆&6≤Ê6∆74Ê÷R“&◊WFVB#∞¢f∆∆&6≤ÁFWáD6ˆÁFVÁB“$6ˆÁFVÁWFÚÊˆ‚Fó7ˆÊñ&ñ∆R‚#∞¢óFV“ÊVÊD6Üñ∆BÜf∆∆&6≤ì∞¢–¢6ˆÁ7B7FñˆÁ2“vóB'Vñ∆D6ÜD÷W76vT7FñˆÁ2Ü÷W76vRì∞¢ñbÜ7FñˆÁ2íóFV“ÊVÊD6Üñ∆BÜ7FñˆÁ2ì∞¢&WGW&‚óFV”∞¢–†¢ñbÜÜ46ˆÁFVÁBí∞¢óFV“ÊVÊD6Üñ∆BÜ6ˆÁFVÁEw&ì∞¢–¢6ˆÁ7B7FñˆÁ2“vóB'Vñ∆D6ÜD÷W76vT7FñˆÁ2Ü÷W76vRì∞¢ñbÜ7FñˆÁ2íóFV“ÊVÊD6Üñ∆BÜ7FñˆÁ2ì∞†¢&WGW&‚óFV”∞ß–†¶gVÊ7Fñˆ‚˜V‰6ÜD÷ˆF¬Çí∞¢VíÊ6ÜD÷ˆF¬Ê6∆74∆ó7BÁ&V÷˜fRÇ&ÜñFFV‚"ì∞¢VíÊ6ÜD÷ˆF¬Á6WDGG&ñ'WFRÇ&&ñ÷ÜñFFV‚"¬&f«6R"ì∞¢÷&¥6ÜD5&VBÇì∞ß–†¶gVÊ7Fñˆ‚6∆˜6T6ÜD÷ˆF¬Çí∞¢VíÊ6ÜD÷ˆF¬Ê6∆74∆ó7BÊFBÇ&ÜñFFV‚"ì∞¢VíÊ6ÜD÷ˆF¬Á6WDGG&ñ'WFRÇ&&ñ÷ÜñFFV‚"¬'G'VR"ì∞ß–†¶7ñÊ2gVÊ7Fñˆ‚6VÊEFWáD÷W76vRÜWfVÁBí∞¢WfVÁBÁ&WfVÁDFVfV«BÇì∞¢6ˆÁ7BFWáB“VíÊ6ÜEFWáBÁf«VRÁG&ñ“Çì∞¢ñbÇFWáBí&WGW&„∞†¢vóB6VÊD6ÜD÷W76vRá∞¢GóS¢'FWáB"¿¢FWáB¿¢&V6óñVÁDñC¢VíÊ6ÜE&V6óñVÁBÁf«VR«¬" ¢“ì∞†¢VíÊ6ÜEFWáBÁf«VR“"#∞ß–†¶7ñÊ2gVÊ7Fñˆ‚6VÊD÷VFñ÷W76vRÜWfVÁBí∞¢6ˆÁ7Bfñ∆R“WfVÁBÁF&vWBÊfñ∆W2bbWfVÁBÁF&vWBÊfñ∆W5≥”∞¢ñbÇfñ∆Rí&WGW&„∞†¢G'í∞¢ñbÇó46VÁG&ƒG&ófT6ˆÊfñwW&VBÇíbbG&ófT66W75Fˆ∂V‚í∞¢Fá&˜rÊWrW'&˜"Ç$6∆˜VB÷÷ñÊó7G&F˜&RÊˆ‚6ˆÊfñwW&FÚ‚6ˆÁFGFV‚F÷ñ‚W"ñÁfñ&Rf˜FÚ˜fñFVÚ‚"ì∞¢–†¢VÊf˜&6T÷VFñ6ó¶RÜfñ∆R¬E$ïdUÙ4ÑEÙ‘TDîÙ‘ÖÙ‘"ì∞¢6ˆÁ7BGóR“fñ∆RÁGóRÁ7F'G5vóFÇÇ'fñFVÚÚ"íÚ'fñFVÚ"¢&ñ÷vR#∞¢6ˆÁ7BW∆ˆB“vóBW∆ˆD&∆ˆ%FÙG&ófRÜfñ∆R¬fñ∆RÊÊ÷R¬fñ∆RÁGóR¬G&ófT6ÜDfˆ∆FW$ñBì∞†¢vóB6VÊD6ÜD÷W76vRá∞¢GóR¿¢FWáC¢""¿¢&V6óñVÁDñC¢VíÊ6ÜE&V6óñVÁBÁf«VR«¬""¿¢÷VFñW&√¢W∆ˆBÊFó&V7EW&¬¿¢÷VFñ÷ñ÷UGóS¢fñ∆RÁGóR¿¢÷VFñÊ÷S¢fñ∆RÊÊ÷R¿¢÷VFñG&ófTfñ∆TñC¢W∆ˆBÊfñ∆TñB¿¢÷VFñG&ófUvV%fñWt∆ñÊ≥¢W∆ˆBÁvV%fñWt∆ñÊ≤«¬" ¢“ì∞†¢VíÊ6ÜDfVVF&6≤ÁFWáD6ˆÁFVÁB“$÷VFññÁfñFÚ7Rvˆˆv∆RG&ófR‚#∞¢“6F6ÇÜW'&˜"í∞¢6ˆÁ6ˆ∆RÊW'&˜"ÜW'&˜"ì∞¢VíÊ6ÜDfVVF&6≤ÁFWáD6ˆÁFVÁB“W'&˜"Ê÷W76vR«¬$W'&˜&RñÁfñÚ÷VFñ‚#∞¢“fñÊ∆«í∞¢VíÊ6ÜD÷VFññÁWBÁf«VR“"#∞¢–ß–†¶7ñÊ2gVÊ7Fñˆ‚Fˆvv∆Ufˆñ6U&V6˜&FñÊrÇí∞¢ñbÇ7W'&VÁEW6W"í∞¢∆W'BÇ$FWfíf&R∆ˆvñ‚‚"ì∞¢&WGW&„∞¢–†¢ñbÇÊfñvF˜"Ê÷VFñFWfñ6W2«¬ÊfñvF˜"Ê÷VFñFWfñ6W2ÊvWEW6W$÷VFñí∞¢VíÊ6ÜDfVVF&6≤ÁFWáD6ˆÁFVÁB“%&Vvó7G&¶ñˆÊRfˆ6∆RÊˆ‚7W˜'FFFVW7FÚ'&˜w6W"‚#∞¢&WGW&„∞¢–†¢ñbÜó5&V6˜&FñÊrbb÷VFñ&V6˜&FW"í∞¢÷VFñ&V6˜&FW"Á7F˜Çì∞¢&WGW&„∞¢–†¢G'í∞¢6ˆÁ7B7G&V““vóBÊfñvF˜"Ê÷VFñFWfñ6W2ÊvWEW6W$÷VFñá≤VFñÛ¢G'VR“ì∞¢÷VFñ6áVÊ∑2“µ”∞¢÷VFñ&V6˜&FW"“ÊWr÷VFñ&V6˜&FW"á7G&V“ì∞†¢÷VFñ&V6˜&FW"ÊˆÊFFfñ∆&∆R“ÜWfVÁBí”‚∞¢ñbÜWfVÁBÊFFÁ6ó¶R‚í÷VFñ6áVÊ∑2ÁW6ÇÜWfVÁBÊFFì∞¢”∞†¢÷VFñ&V6˜&FW"ÊˆÁ7F˜“7ñÊ2Çí”‚∞¢G'í∞¢ñbÇó46VÁG&ƒG&ófT6ˆÊfñwW&VBÇíbbG&ófT66W75Fˆ∂V‚í∞¢Fá&˜rÊWrW'&˜"Ç$6∆˜VB÷÷ñÊó7G&F˜&RÊˆ‚6ˆÊfñwW&FÚ‚6ˆÁFGFV‚F÷ñ‚W"ñÁfñ&Rfˆ6∆í‚"ì∞¢–†¢6ˆÁ7B&∆ˆ"“ÊWr&∆ˆ"Ü÷VFñ6áVÊ∑2¬≤GóS¢÷VFñ&V6˜&FW"Ê÷ñ÷UGóR«¬&VFñÚ˜vV&“"“ì∞¢VÊf˜&6T÷VFñ6ó¶RÜ&∆ˆ"¬E$ïdUÙ4ÑEÙ‘TDîÙ‘ÖÙ‘"ì∞¢6ˆÁ7Bfñ∆TÊ÷R“fˆ6∆R“G¥FFRÊÊ˜rÇó“ÁvV&÷∞¢6ˆÁ7BW∆ˆB“vóBW∆ˆD&∆ˆ%FÙG&ófRÜ&∆ˆ"¬fñ∆TÊ÷R¬&∆ˆ"ÁGóR«¬&VFñÚ˜vV&“"¬G&ófT6ÜDfˆ∆FW$ñBì∞†¢vóB6VÊD6ÜD÷W76vRá∞¢GóS¢'fˆñ6R"¿¢FWáC¢""¿¢&V6óñVÁDñC¢VíÊ6ÜE&V6óñVÁBÁf«VR«¬""¿¢÷VFñW&√¢W∆ˆBÊFó&V7EW&¬¿¢÷VFñ÷ñ÷UGóS¢&∆ˆ"ÁGóR«¬&VFñÚ˜vV&“"¿¢÷VFñÊ÷S¢fñ∆TÊ÷R¿¢÷VFñG&ófTfñ∆TñC¢W∆ˆBÊfñ∆TñB¿¢÷VFñG&ófUvV%fñWt∆ñÊ≥¢W∆ˆBÁvV%fñWt∆ñÊ≤«¬" ¢“ì∞†¢VíÊ6ÜDfVVF&6≤ÁFWáD6ˆÁFVÁB“$÷W76vvñÚfˆ6∆RñÁfñFÚ7Rvˆˆv∆RG&ófR‚#∞¢“6F6ÇÜW'&˜"í∞¢6ˆÁ6ˆ∆RÊW'&˜"ÜW'&˜"ì∞¢VíÊ6ÜDfVVF&6≤ÁFWáD6ˆÁFVÁB“W'&˜"Ê÷W76vR«¬$W'&˜&RñÁfñÚfˆ6∆R‚#∞¢“fñÊ∆«í∞¢7G&V“ÊvWEG&6∑2ÇíÊf˜$V6ÇÇáG&6≤í”‚G&6≤Á7F˜Çíì∞¢÷VFñ&V6˜&FW"“ÁV∆√∞¢÷VFñ6áVÊ∑2“µ”∞¢ó5&V6˜&FñÊr“f«6S∞¢VíÊ6ÜEfˆñ6T'F‚ÁFWáD6ˆÁFVÁB“/	¯ÍBñÁfñfˆ6∆R#∞¢–¢”∞†¢÷VFñ&V6˜&FW"Á7F'BÇì∞¢ó5&V6˜&FñÊr“G'VS∞¢VíÊ6ÜEfˆñ6T'F‚ÁFWáD6ˆÁFVÁB“.(˚û˚àÚ7F˜RñÁfñ#∞¢VíÊ6ÜDfVVF&6≤ÁFWáD6ˆÁFVÁB“%&Vvó7G&¶ñˆÊRñ‚6˜'6Ú‚‚‚#∞¢“6F6ÇÜW'&˜"í∞¢6ˆÁ6ˆ∆RÊW'&˜"ÜW'&˜"ì∞¢VíÊ6ÜDfVVF&6≤ÁFWáD6ˆÁFVÁB“$ñ◊˜76ñ&ñ∆R66VFW&R¬÷ñ7&ˆfˆÊÚ‚#∞¢–ß–†¶7ñÊ2gVÊ7Fñˆ‚6VÊD6ÜD÷W76vRáñ∆ˆBí∞¢ñbÇ7W'&VÁEW6W"í∞¢∆W'BÇ$FWfíf&R∆ˆvñ‚‚"ì∞¢&WGW&„∞¢–†¢6ˆÁ7B6VÊFW$Ê÷R“7W'&VÁEW6W"ÊFó7∆îÊ÷R«¬7W'&VÁEW6W"ÊV÷ñ¬«¬$˜W&F˜&R#∞†¢vóBF"Ê6ˆ∆∆V7Fñˆ‚Ç&6ÜD÷W76vW2"íÊFBá∞¢‚‚Áñ∆ˆB¿¢6VÊFW$ñC¢7W'&VÁEW6W"ÁVñB¿¢6VÊFW$Ê÷R¿¢6VÊFW$V÷ñ√¢7W'&VÁEW6W"ÊV÷ñ¬«¬""¿¢7&VFVDC¢fó&V&6RÊfó&W7F˜&R‰fñV∆Ef«VRÁ6W'fW%Fñ÷W7F◊Çê¢“ì∞ß–†¶gVÊ7Fñˆ‚ó4˜v‰÷W76vRÜ÷W76vRí∞¢&WGW&‚&ˆˆ∆V‚Ü7W'&VÁEW6W"bb÷W76vRÁ6VÊFW$ñB””“7W'&VÁEW6W"ÁVñBì∞ß–†¶gVÊ7Fñˆ‚VÊf˜&6T÷VFñ6ó¶RÜfñ∆T˜$&∆ˆ"¬÷Ñ÷"í∞¢6ˆÁ7B÷Ñ'óFW2“÷Ñ÷"¢#B¢#C∞¢ñbÜfñ∆T˜$&∆ˆ"Á6ó¶R‚÷Ñ'óFW2í∞¢Fá&˜rÊWrW'&˜"Üfñ∆RG&˜Úw&ÊFR‚∆ñ÷óFS¢G∂÷Ñ÷'‘‘"Êì∞¢–ß–†¶gVÊ7Fñˆ‚&W6WDG&ófU7FFRÇí∞¢G&ófT66W75Fˆ∂V‚“"#∞¢G&ófU&ˆ˜Dfˆ∆FW$ñB“"#∞¢G&ófT6ÜDfˆ∆FW$ñB“"#∞¢G&ófU&W˜'G4fˆ∆FW$ñB“"#∞¢G&ófU7VG&Tfˆ∆FW$ñB“"#∞¢G&ófTÜV«6VÁFW$fˆ∆FW$ñB“"#∞¢6ˆ÷÷W766ÜVWD66ÜRÊ6∆V"Çì∞¢WFFTG&ófU7FGW2Üf«6Rì∞ß–†¶gVÊ7Fñˆ‚WFFTG&ófU7FGW2Üó46ˆÊÊV7FVBí∞¢6ˆÁ7B6ˆÊÊV7FVB“&ˆˆ∆V‚Üó46ˆÊÊV7FVB«¬G&ófT'&ñFvU7FFRÊ6ˆÊfñwW&VBì∞¢VíÊG&ófU7FGW2Ê6∆74∆ó7BÁFˆvv∆RÇ'7FGW2÷6Üó÷G&ófR"¬6ˆÊÊV7FVBì∞¢ñbÜ6ˆÊÊV7FVBí∞¢VíÊG&ófU7FGW2ÁFWáD6ˆÁFVÁB“6‰÷ÊvTFFÇê¢Ú6∆˜VB6VÁG&∆óß¶FÚGFófÚG∂G&ófT'&ñFvU7FFRÊ˜vÊW$V÷ñ¬ÚÇG∂G&ófT'&ñFvU7FFRÊ˜vÊW$V÷ñ«“ñ¢"'“Ê ¢¢$&6Üófñ¶ñˆÊR6∆˜VBGFóf(
"6∆˜VB6VÁG&∆óß¶FÚGFófÚ#∞¢“V«6R∞¢VíÊG&ófU7FGW2ÁFWáD6ˆÁFVÁB“vWD6VÁG&ƒG&ófTÊ˜D6ˆÊfñwW&VD÷W76vRÇì∞¢–ß–†¶7ñÊ2gVÊ7Fñˆ‚6ˆÊÊV7Dvˆˆv∆TG&ófRÇí∞¢ñbÇ6‰÷ÊvTFFÇíí∞¢∆W'BÇ%6ˆ∆ÚF÷ñ‚\;"6ˆÊfñwW&&Rvˆˆv∆RG&ófR‚"ì∞¢&WGW&„∞¢–¢G'í∞¢6ˆÁ7B&˜fñFW"“ÊWrfó&V&6RÊWFÇ‰vˆˆv∆TWFÖ&˜fñFW"Çì∞¢&˜fñFW"ÊFE66˜RÇ&áGG3¢Ú˜wwrÊvˆˆv∆Vó2Ê6ˆ“ˆWFÇˆG&ófRÊfñ∆R"ì∞¢6ˆÁ7B&W7V«B“vóBfó&V&6RÊWFÇÇíÁ6ñv‰ñÂvóFÖ˜Wá&˜fñFW"ì∞¢6ˆÁ7B7&VFVÁFñ¬“&W7V«BÊ7&VFVÁFñ¬«¬fó&V&6RÊWFÇ‰vˆˆv∆TWFÖ&˜fñFW"Ê7&VFVÁFñƒg&ˆ’&W7V«Bá&W7V«Bì∞¢6ˆÁ7B66W75Fˆ∂V‚“7&VFVÁFñ¬bb7&VFVÁFñ¬Ê66W75Fˆ∂V‚Ú7&VFVÁFñ¬Ê66W75Fˆ∂V‚¢ÁV∆√∞¢ñbÇ66W75Fˆ∂V‚í∞¢Fá&˜rÊWrW'&˜"Ç$66W72Fˆ∂V‚vˆˆv∆RG&ófRÊˆ‚˜GFVÁWFÚ"ì∞¢–†¢W'6ó7DG&ófT66W75Fˆ∂V‚Ü66W75Fˆ∂V‚ì∞¢vóBWFÙ6ˆÊÊV7DG&ófT'&ñFvRá≤Ê˜Fñgîˆ‰W'&˜#¢G'VR“ì∞¢ñbá6V∆V7FVD6ˆ÷÷W76ñBí∞¢66ÜVGV∆T6ˆ÷÷W766ÜVWE7ñÊ2á6V∆V7FVD6ˆ÷÷W76ñB¬6V∆V7FVD6ˆ÷÷W76Ê÷R¬#ì∞¢–¢∆W'BÇ$vˆˆv∆RG&ófR6ˆ∆∆VvFÚ6˜'&WGF÷VÁFR"ì∞¢“6F6ÇÜW'&˜"í∞¢6ˆÁ6ˆ∆RÊW'&˜"Ç$W'&˜&R6ˆ∆∆Vv÷VÁFÚvˆˆv∆RG&ófS¢"¬W'&˜"ì∞¢&V6˜fW$fó&W7F˜&UW'6ó7FVÊ6RÜW'&˜"ì∞¢∆W'BÇ$W'&˜&R6ˆ∆∆Vv÷VÁFÚvˆˆv∆RG&ófS¢"≤f˜&÷D∆ˆvñ‰W'&˜"ÜW'&˜"íì∞¢–ß–†¶gVÊ7Fñˆ‚WáG&7Dvˆˆv∆T66W75Fˆ∂V‚á&W7V«Bí∞¢ñbá&W7V«Bbb&W7V«BÊ7&VFVÁFñ¬bb&W7V«BÊ7&VFVÁFñ¬Ê66W75Fˆ∂V‚í∞¢&WGW&‚&W7V«BÊ7&VFVÁFñ¬Ê66W75Fˆ∂V„∞¢–¢ñbÄ¢fó&V&6P¢bbfó&V&6RÊWFÄ¢bbfó&V&6RÊWFÇ‰vˆˆv∆TWFÖ&˜fñFW ¢bbGóVˆbfó&V&6RÊWFÇ‰vˆˆv∆TWFÖ&˜fñFW"Ê7&VFVÁFñƒg&ˆ’&W7V«B””“&gVÊ7Fñˆ‚ ¢í∞¢6ˆÁ7B7&VFVÁFñ¬“fó&V&6RÊWFÇ‰vˆˆv∆TWFÖ&˜fñFW"Ê7&VFVÁFñƒg&ˆ’&W7V«Bá&W7V«Bì∞¢ñbÜ7&VFVÁFñ¬bb7&VFVÁFñ¬Ê66W75Fˆ∂V‚í&WGW&‚7&VFVÁFñ¬Ê66W75Fˆ∂V„∞¢–¢&WGW&‚"#∞ß–†¢ÚÚ6«fV‚ˆvvWGFÚ•4Ù‚ÊV¬G&ófRFV∆¬wWFVÁFR∆ˆvvFÚW6ÊFÚ◊V«Fó'BW∆ˆB‡¢ÚÚ&ñ6ÜñVFR6ÜRñ¬∆ˆvñ‚vˆˆv∆R&&ñ&W7FóGVóFÚV‚66W72Fˆ∂V‚6ˆ‚66˜RG&ófRÊfñ∆R‡¶7ñÊ2gVÊ7Fñˆ‚6fUFÙG&ófRÜFFí∞¢ñbÇG&ófT66W75Fˆ∂V‚í∞¢6ˆÁ6ˆ∆RÊW'&˜"Ç$vˆˆv∆RG&ófRÊˆ‚WF˜&óß¶FÛ¢÷Ê666W72Fˆ∂V‚‚&ñfí∆ˆvñ‚6ˆ‚vˆˆv∆R‚"ì∞¢&WGW&‚ÁV∆√∞¢–†¢6ˆÁ7B÷WFFF“∞¢Ê÷S¢'FW7BÊß6ˆ‚"¿¢÷ñ÷UGóS¢&∆ñ6Fñˆ‚ˆß6ˆ‚ ¢”∞¢6ˆÁ7Bfñ∆T6ˆÁFVÁB“•4Ù‚Á7G&ñÊvñgíÜFF¬ÁV∆¬¬"ì∞†¢6ˆÁ7B&˜VÊF'í“&ÜW&÷÷&˜VÊF'í“"≤FFRÊÊ˜rÇì∞¢6ˆÁ7B◊V«Fó'D&ˆGí“∞¢““G∂&˜VÊF'ó÷¿¢$6ˆÁFVÁB’GóS¢∆ñ6Fñˆ‚ˆß6ˆ„≤6Ü'6WC’UDb”Ç"¿¢""¿¢•4Ù‚Á7G&ñÊvñgíÜ÷WFFFí¿¢““G∂&˜VÊF'ó÷¿¢$6ˆÁFVÁB’GóS¢∆ñ6Fñˆ‚ˆß6ˆ‚"¿¢""¿¢fñ∆T6ˆÁFVÁB¿¢““G∂&˜VÊF'ó““÷ ¢“Ê¶ˆñ‚Ç%«%∆‚"ì∞†¢G'í∞¢6ˆÁ7B&W7ˆÁ6R“vóBfWF6ÇÇ&áGG3¢Ú˜wwrÊvˆˆv∆Vó2Ê6ˆ“˜W∆ˆBˆG&ófR˜c2ˆfñ∆W3˜W∆ˆEGóS÷◊V«Fó'B"¬∞¢÷WFÜˆC¢%ı5B"¿¢ÜVFW'3¢∞¢WFÜ˜&ó¶Fñˆ„¢&V&W"G∂G&ófT66W75Fˆ∂VÁ÷¿¢$6ˆÁFVÁB’GóR#¢◊V«Fó'B˜&V∆FVC≤&˜VÊF'ì“G∂&˜VÊF'ó÷ ¢“¿¢&ˆGì¢◊V«Fó'D&ˆGê¢“ì∞†¢ñbÇ&W7ˆÁ6RÊˆ≤í∞¢6ˆÁ7BW'&˜%FWáB“vóB&W7ˆÁ6RÁFWáBÇì∞¢6ˆÁ6ˆ∆RÊW'&˜"Ç$W'&˜&RW∆ˆB7Rvˆˆv∆RG&ófS¢"¬&W7ˆÁ6RÁ7FGW2¬W'&˜%FWáBì∞¢&WGW&‚ÁV∆√∞¢–†¢6ˆÁ7B&W7V«B“vóB&W7ˆÁ6RÊß6ˆ‚Çì∞¢6ˆÁ6ˆ∆RÊ∆ˆrÇ%W∆ˆB6ˆ◊∆WFFÚ7Rvˆˆv∆RG&ófR‚fñ∆TñC¢"¬&W7V«BÊñBì∞¢&WGW&‚≤fñ∆TñC¢&W7V«BÊñB”∞¢“6F6ÇÜW'&˜"í∞¢6ˆÁ6ˆ∆RÊW'&˜"Ç$W'&˜&RGW&ÁFRñ¬6«fFvvñÚ7Rvˆˆv∆RG&ófS¢"¬W'&˜"ì∞¢&WGW&‚ÁV∆√∞¢–ß–†¢ÚÚW6V◊ñÚBwW6Û†¢ÚÚ6fUFÙG&ófRá∞¢ÚÚ&˜f¢G'VR¿¢ÚÚFF¢ÊWrFFRÇíÁFÙï4ı7G&ñÊrÇê¢ÚÚ“ì∞†¶7ñÊ2gVÊ7Fñˆ‚VÁ7W&TG&ófTfˆ∆FW'2Çí∞¢G&ófU&ˆ˜Dfˆ∆FW$ñB“4TÂE$≈ÙE$ïdUı$ÙıEÙdÙƒDU%ÙîC∞¢G&ófT6ÜDfˆ∆FW$ñB“$dıDÚ#∞¢G&ófU&W˜'G4fˆ∆FW$ñB“%4Tt‰ƒ§îÙ‰í#∞¢G&ófU7VG&Tfˆ∆FW$ñB“$UÖı%B#∞¢G&ófTÜV«6VÁFW$fˆ∆FW$ñB“$UÖı%B#∞ß–†¶gVÊ7Fñˆ‚Ê˜&÷∆ó¶TfFFÜFFí∞¢6ˆÁ7Bñ∆ˆB“FFbbGóVˆbFF””“&ˆ&¶V7B"ÚFF¢∑”∞¢6ˆÁ7B&tóFV◊2“'&íÊó4'&íáñ∆ˆBÊóFV◊2íÚñ∆ˆBÊóFV◊2¢µ”∞¢6ˆÁ7BÊ˜&÷∆ó¶VDóFV◊2“&tóFV◊2Ê÷ÇÜóFV“¬ñÊFWÇí”‚á∞¢ñC¢7G&ñÊrÜóFV“ÊñB«¬f“G∂ñÊFWÇ≤÷í¿¢Fˆ÷ÊF¢7G&ñÊrÜóFV“ÊFˆ÷ÊF«¬óFV“ÁVW7Fñˆ‚«¬""íÁG&ñ“Çí¿¢&ó7˜7F¢7G&ñÊrÜóFV“Á&ó7˜7F«¬óFV“ÊÁ7vW"«¬""íÁG&ñ“Çí¿¢76ì¢'&íÊó4'&íÜóFV“Á76ííÚóFV“Á76íÊ÷Çá7FWí”‚7G&ñÊrá7FW«¬""íÁG&ñ“ÇííÊfñ«FW"Ñ&ˆˆ∆V‚í¢µ–¢“ííÊfñ«FW"ÇÜóFV“í”‚óFV“ÊFˆ÷ÊFbbóFV“Á&ó7˜7Fì∞†¢&WGW&‚∞¢fW'6ñˆ„¢ÁV÷&W"áñ∆ˆBÁfW'6ñˆ‚í‚ÚÁV÷&W"áñ∆ˆBÁfW'6ñˆ‚í¢ÑT≈Ù4TÂDU%ÙdÙdƒƒ$4≤ÁfW'6ñˆ‚¿¢WFFVDC¢ñ∆ˆBÁWFFVDB«¬ÁV∆¬¿¢WFFVD'ì¢7G&ñÊráñ∆ˆBÁWFFVD'í«¬""í¿¢óFV◊3¢Ê˜&÷∆ó¶VDóFV◊2Ê∆VÊwFÇ‚ÚÊ˜&÷∆ó¶VDóFV◊2¢ÑT≈Ù4TÂDU%ÙdÙdƒƒ$4≤ÊóFV◊0¢”∞ß–†¶gVÊ7Fñˆ‚W66TáF÷¬áf«VRí∞¢&WGW&‚7G&ñÊráf«VR«¬""ê¢Á&W∆6T∆¬Ç"b"¬"f◊≤"ê¢Á&W∆6T∆¬Ç#¬"¬"f«C≤"ê¢Á&W∆6T∆¬Ç#‚"¬"fwC≤"ê¢Á&W∆6T∆¬Ç%¬""¬"gV˜C≤"ê¢Á&W∆6T∆¬Ç"r"¬"b33ì≤"ì∞ß–†¶gVÊ7Fñˆ‚FÙó6ÙFFRáf«VRí∞¢ñbÇf«VRí&WGW&‚"#∞¢ñbáGóVˆbf«VR””“'7G&ñÊr"í&WGW&‚f«VS∞¢ñbáf«VRñÁ7FÊ6VˆbFFRí&WGW&‚f«VRÁFÙï4ı7G&ñÊrÇì∞¢ñbáf«VRbbGóVˆbf«VRÁFÙFFR””“&gVÊ7Fñˆ‚"í&WGW&‚f«VRÁFÙFFRÇíÁFÙï4ı7G&ñÊrÇì∞¢&WGW&‚"#∞ß–†¶7ñÊ2gVÊ7Fñˆ‚∆ˆDfg&ˆ‘fó&W7F˜&RÇí∞¢ñbÇ7W'&VÁEW6W"í&WGW&‚ÁV∆√∞¢G'í∞¢6ˆÁ7B6Ê6Ü˜B“vóB'V‰fó&W7F˜&TvWEvóFÖ&WG'íÄ¢F"Ê6ˆ∆∆V7Fñˆ‚ÑÑT≈Ù4TÂDU%Ù4Ù‰dîuıDÇÊ6ˆ∆∆V7Fñˆ‚íÁvÜW&RÜfó&V&6RÊfó&W7F˜&R‰fñV∆EFÇÊFˆ7V÷VÁDñBÇí¬#”“"¬ÑT≈Ù4TÂDU%Ù4Ù‰dîuıDÇÊFˆ2í¿¢≤∆&V√¢$ƒÙBÑT≈4TÂDU""¬Fñ÷V˜WD◊3¢ì¬&WG&ñW3¢–¢ì∞¢6ˆÁ7BFˆ2“6Ê6Ü˜BÊFˆ75≥“«¬≤WÜó7G3¢f«6R¬FF¢Çí”‚á∑“í”∞¢ñbÇFˆ2ÊWÜó7G2í&WGW&‚ÁV∆√∞¢6ˆÁ7BFF“Ê˜&÷∆ó¶TfFFÜFˆ2ÊFFÇíì∞¢fFF6WB“FF∞¢&WGW&‚FF∞¢“6F6ÇÜW'&˜"í∞¢6ˆÁ6ˆ∆RÁv&‚Ç$ÜV«6VÁFW"fó&W7F˜&RÊˆ‚Fó7ˆÊñ&ñ∆R¬W6Úf∆∆&6≤∆ˆ6∆S¢"¬W'&˜"ì∞¢&WGW&‚ÁV∆√∞¢–ß–†¶7ñÊ2gVÊ7Fñˆ‚6fTfFÙfó&W7F˜&RÜfFFí∞¢ñbÇ6‰÷ÊvTFFÇíí∞¢Fá&˜rÊWrW'&˜"Ç%6ˆ∆ÚV‚F÷ñ‚\;"vvñ˜&Ê&R∆Rd‚"ì∞¢–¢6ˆÁ7BÊ˜&÷∆ó¶VB“Ê˜&÷∆ó¶TfFFÜfFFì∞¢6ˆÁ7BWÜó7FñÊr“vóB∆ˆDfg&ˆ‘fó&W7F˜&RÇì∞¢6ˆÁ7BÊWáEfW'6ñˆ‚“ÜWÜó7FñÊrbbÁV÷&W"ÜWÜó7FñÊrÁfW'6ñˆ‚í‚ÚÁV÷&W"ÜWÜó7FñÊrÁfW'6ñˆ‚í¢í≤∞¢6ˆÁ7Bñ∆ˆB“∞¢‚‚ÊÊ˜&÷∆ó¶VB¿¢fW'6ñˆ„¢ÊWáEfW'6ñˆ‚¿¢WFFVDC¢fó&V&6RÊfó&W7F˜&R‰fñV∆Ef«VRÁ6W'fW%Fñ÷W7F◊Çí¿¢WFFVD'ì¢7W'&VÁEW6W#ÚÊV÷ñ¬«¬" ¢”∞¢vóBF"Ê6ˆ∆∆V7Fñˆ‚ÑÑT≈Ù4TÂDU%Ù4Ù‰dîuıDÇÊ6ˆ∆∆V7Fñˆ‚íÊFˆ2ÑÑT≈Ù4TÂDU%Ù4Ù‰dîuıDÇÊFˆ2íÁ6WBáñ∆ˆB¬≤÷W&vS¢G'VR“ì∞¢fFF6WB“≤‚‚ÊÊ˜&÷∆ó¶VB¬fW'6ñˆ„¢ÊWáEfW'6ñˆ‚¬WFFVDC¢ÊWrFFRÇíÁFÙï4ı7G&ñÊrÇí¬WFFVD'ì¢7W'&VÁEW6W#ÚÊV÷ñ¬«¬""”∞¢6ˆÁ7B6Ê6Ü˜B“vóBWá˜'Df6Ê6Ü˜EFÙG&ófRÜfFF6WBì∞¢&WGW&‚≤‚‚ÊfFF6WB¬6Ê6Ü˜B”∞ß–†¶gVÊ7Fñˆ‚&VÊFW$fÜV«6VÁFW"ÜfFFí∞¢6ˆÁ7BÊ˜&÷∆ó¶VB“Ê˜&÷∆ó¶TfFFÜfFFì∞¢fFF6WB“Ê˜&÷∆ó¶VC∞¢vñÊF˜rÊÜV«fFF“Ê˜&÷∆ó¶VC∞†¢6ˆÁ7B∆ó7B“Fˆ7V÷VÁBÊvWDV∆V÷VÁD'îñBÇ&ÜV«÷f÷∆ó7B"í«¬VíÊÜ˜wFÙf∆ó7C∞¢ñbÇ∆ó7Bí&WGW&„∞¢∆ó7BÊñÊÊW$ÖD‘¬“Ê˜&÷∆ó¶VBÊóFV◊2Ê÷ÇÜóFV“í”‚∞¢6ˆÁ7B7FW2“óFV“Á76íÊ∆VÊwFÄ¢Ú∆ˆ√‚G∂óFV“Á76íÊ÷Çá7FWí”‚∆∆ì‚G∂W66TáF÷¬á7FWó”¬ˆ∆ìÊíÊ¶ˆñ‚Ç""ó”¬ˆˆ√Ê ¢¢"#∞¢&WGW&‚∆'Fñ6∆R6∆73“&f÷óFV“#„∆É3‚G∂W66TáF÷¬ÜóFV“ÊFˆ÷ÊFó”¬ˆÉ3„«‚G∂W66TáF÷¬ÜóFV“Á&ó7˜7Fó”¬˜‚G∑7FW7”¬ˆ'Fñ6∆SÊ∞¢“íÊ¶ˆñ‚Ç""ì∞ß–†¶7ñÊ2gVÊ7Fñˆ‚ñÊóDÜV«6VÁFW$fÇí∞¢6ˆÁ7B&V÷˜FTf“vóB∆ˆDfg&ˆ‘fó&W7F˜&RÇì∞¢&VÊFW$fÜV«6VÁFW"á&V÷˜FTf«¬ÑT≈Ù4TÂDU%ÙdÙdƒƒ$4≤ì∞ß–†ßvñÊF˜rÊ∆ˆDfg&ˆ‘fó&W7F˜&R“∆ˆDfg&ˆ‘fó&W7F˜&S∞ßvñÊF˜rÁ6fTfFÙfó&W7F˜&R“6fTfFÙfó&W7F˜&S∞ßvñÊF˜rÊWá˜'Df6Ê6Ü˜EFÙG&ófR“Wá˜'Df6Ê6Ü˜EFÙG&ófS∞†¶7ñÊ2gVÊ7Fñˆ‚Wá˜'Df6Ê6Ü˜EFÙG&ófRÜfFF“fFF6WBí∞¢ñbÇ6‰÷ÊvTFFÇíí∞¢Fá&˜rÊWrW'&˜"Ç%6ˆ∆ÚV‚F÷ñ‚\;"W7˜'F&R6Ê6Ü˜Bd‚"ì∞¢–¢ñbÇG&ófT66W75Fˆ∂V‚í∞¢6ˆÁ6ˆ∆RÁv&‚Ç$6∆˜VB÷÷ñÊó7G&F˜&RÊˆ‚6ˆÊfñwW&FÛ¢6«FÚWá˜'B6Ê6Ü˜Bd‚"ì∞¢&WGW&‚ÁV∆√∞¢–†¢ñbÇG&ófTÜV«6VÁFW$fˆ∆FW$ñBívóBVÁ7W&TG&ófTfˆ∆FW'2Çì∞¢6ˆÁ7BÊ˜&÷∆ó¶VB“Ê˜&÷∆ó¶TfFFÜfFFì∞¢6ˆÁ7B÷WFFF“∞¢fW'6ñˆ„¢ÁV÷&W"ÜÊ˜&÷∆ó¶VBÁfW'6ñˆ‚í«¬¿¢WFFVDC¢FÙó6ÙFFRÜÊ˜&÷∆ó¶VBÁWFFVDBí«¬ÊWrFFRÇíÁFÙï4ı7G&ñÊrÇí¿¢WFFVD'ì¢Ê˜&÷∆ó¶VBÁWFFVD'í«¬7W'&VÁEW6W#ÚÊV÷ñ¬«¬" ¢”∞¢6ˆÁ7Bñ∆ˆB“≤‚‚Ê÷WFFF¬óFV◊3¢Ê˜&÷∆ó¶VBÊóFV◊2”∞¢6ˆÁ7B&∆ˆ"“ÊWr&∆ˆ"Ö¥•4Ù‚Á7G&ñÊvñgíáñ∆ˆB¬ÁV∆¬¬"ï“¬≤GóS¢&∆ñ6Fñˆ‚ˆß6ˆ‚"“ì∞¢6ˆÁ7Bfñ∆TÊ÷R“ÜV«÷6VÁFW"÷f◊bG∂÷WFFFÁfW'6ñˆÁ“Êß6ˆÊ∞¢6ˆÁ7BW∆ˆFVB“vóBW∆ˆD&∆ˆ%FÙG&ófRÜ&∆ˆ"¬fñ∆TÊ÷R¬&∆ñ6Fñˆ‚ˆß6ˆ‚"¬G&ófTÜV«6VÁFW$fˆ∆FW$ñBì∞¢vóBF"Ê6ˆ∆∆V7Fñˆ‚ÑÑT≈Ù4TÂDU%Ù4Ù‰dîuıDÇÊ6ˆ∆∆V7Fñˆ‚íÊFˆ2ÑÑT≈Ù4TÂDU%Ù4Ù‰dîuıDÇÊFˆ2íÁ6WBá∞¢∆FW7E6Ê6Ü˜C¢∞¢fñ∆TñC¢W∆ˆFVBÊfñ∆TñB¿¢W&√¢W∆ˆFVBÁvV%fñWt∆ñÊ≤«¬W∆ˆFVBÊFó&V7EW&¬«¬""¿¢Wá˜'FVDC¢fó&V&6RÊfó&W7F˜&R‰fñV∆Ef«VRÁ6W'fW%Fñ÷W7F◊Çí¿¢Wá˜'FVD'ì¢÷WFFFÁWFFVD'ê¢–¢“¬≤÷W&vS¢G'VR“ì∞¢&WGW&‚W∆ˆFVC∞ß–†¶7ñÊ2gVÊ7Fñˆ‚vWD˜$7&VFTG&ófTfˆ∆FW"ÜÊ÷R¬&VÁDñB“""í∞¢6ˆÁ7B&VÁEVW'í“&VÁDñBÚÊBrG∑&VÁDñG“rñ‚&VÁG6¢"#∞¢6ˆÁ7B6fTÊ÷R“W66TG&ófUVW'ïf«VRÜÊ÷Rì∞¢6ˆÁ7BVW'í“÷ñ÷UGóS“v∆ñ6Fñˆ‚˜fÊBÊvˆˆv∆R÷2Êfˆ∆FW"rÊBG&6ÜVC÷f«6RÊBÊ÷S“rG∑6fTÊ÷W“rG∑&VÁEVW'ó÷∞¢6ˆÁ7B6V&6ÖW&¬“áGG3¢Ú˜wwrÊvˆˆv∆Vó2Ê6ˆ“ˆG&ófR˜c2ˆfñ∆W3˜“G∂VÊ6ˆFUU$î6ˆ◊ˆÊVÁBáVW'íó“ffñV∆G3÷fñ∆W2ÜñB∆Ê÷R∆7&VFVEFñ÷Ríf˜&FW$'ì÷7&VFVEFñ÷RgvU6ó¶S”∞¢6ˆÁ7B6V&6Ö&W7ˆÁ6R“vóBG&ófTîfWF6Çá6V&6ÖW&¬¬≤÷WFÜˆC¢$tUB"“ì∞†¢ñbÑ'&íÊó4'&íá6V&6Ö&W7ˆÁ6RÊfñ∆W2íbb6V&6Ö&W7ˆÁ6RÊfñ∆W2Ê∆VÊwFÇ‚í∞¢&WGW&‚6V&6Ö&W7ˆÁ6RÊfñ∆W5≥“ÊñC∞¢–†¢6ˆÁ7B7&VFUñ∆ˆB“∞¢Ê÷R¿¢÷ñ÷UGóS¢&∆ñ6Fñˆ‚˜fÊBÊvˆˆv∆R÷2Êfˆ∆FW" ¢”∞¢ñbá&VÁDñBí7&VFUñ∆ˆBÁ&VÁG2“∑&VÁDñE”∞†¢6ˆÁ7B7&VFVB“vóBG&ófTîfWF6ÇÇ&áGG3¢Ú˜wwrÊvˆˆv∆Vó2Ê6ˆ“ˆG&ófR˜c2ˆfñ∆W2"¬∞¢÷WFÜˆC¢%ı5B"¿¢ÜVFW'3¢≤$6ˆÁFVÁB’GóR#¢&∆ñ6Fñˆ‚ˆß6ˆ‚"“¿¢&ˆGì¢•4Ù‚Á7G&ñÊvñgíÜ7&VFUñ∆ˆBê¢“ì∞†¢&WGW&‚7&VFVBÊñC∞ß–†¶7ñÊ2gVÊ7Fñˆ‚fñÊDG&ófTfˆ∆FW'4'îÊ÷RÜÊ÷R¬&VÁDñB“""í∞¢6ˆÁ7B&VÁEVW'í“&VÁDñBÚÊBrG∑&VÁDñG“rñ‚&VÁG6¢"#∞¢6ˆÁ7B6fTÊ÷R“W66TG&ófUVW'ïf«VRÜÊ÷Rì∞¢6ˆÁ7BVW'í“÷ñ÷UGóS“v∆ñ6Fñˆ‚˜fÊBÊvˆˆv∆R÷2Êfˆ∆FW"rÊBG&6ÜVC÷f«6RÊBÊ÷S“rG∑6fTÊ÷W“rG∑&VÁEVW'ó÷∞¢6ˆÁ7B6V&6ÖW&¬“áGG3¢Ú˜wwrÊvˆˆv∆Vó2Ê6ˆ“ˆG&ófR˜c2ˆfñ∆W3˜“G∂VÊ6ˆFUU$î6ˆ◊ˆÊVÁBáVW'íó“ffñV∆G3÷fñ∆W2ÜñB∆Ê÷R«&VÁG2∆7&VFVEFñ÷Ríf˜&FW$'ì÷7&VFVEFñ÷RgvU6ó¶S”∞¢6ˆÁ7B6V&6Ö&W7ˆÁ6R“vóBG&ófTîfWF6Çá6V&6ÖW&¬¬≤÷WFÜˆC¢$tUB"“ì∞¢&WGW&‚'&íÊó4'&íá6V&6Ö&W7ˆÁ6RÊfñ∆W2íÚ6V&6Ö&W7ˆÁ6RÊfñ∆W2¢µ”∞ß–†¶7ñÊ2gVÊ7Fñˆ‚÷˜fTG&ófTfñ∆UFÙfˆ∆FW"Üfñ∆TñB¬F&vWE&VÁDñB¬7W'&VÁE&VÁG2“µ“í∞¢ñbÇfñ∆TñB«¬F&vWE&VÁDñBí&WGW&„∞¢6ˆÁ7B&V÷˜fU&VÁG2“7W'&VÁE&VÁG2Êfñ«FW"Çá&VÁDñBí”‚&VÁDñBbb&VÁDñB”“F&vWE&VÁDñBíÊ¶ˆñ‚Ç"¬"ì∞¢6ˆÁ7B&◊2“ÊWrU$≈6V&6Ö&◊2á≤FE&VÁG3¢F&vWE&VÁDñB¬fñV∆G3¢&ñB«&VÁG2"“ì∞¢ñbá&V÷˜fU&VÁG2í&◊2Á6WBÇ'&V÷˜fU&VÁG2"¬&V÷˜fU&VÁG2ì∞¢vóBG&ófTîfWF6ÇÜáGG3¢Ú˜wwrÊvˆˆv∆Vó2Ê6ˆ“ˆG&ófR˜c2ˆfñ∆W2ÚG∂fñ∆TñG”ÚG∑&◊2ÁFı7G&ñÊrÇó÷¬∞¢÷WFÜˆC¢%D4Ç"¿¢ÜVFW'3¢≤$6ˆÁFVÁB’GóR#¢&∆ñ6Fñˆ‚ˆß6ˆ‚"“¿¢&ˆGì¢•4Ù‚Á7G&ñÊvñgíá∑“ê¢“ì∞ß–†¶7ñÊ2gVÊ7Fñˆ‚÷ñw&FT∆Vv7îG&ófTFFFÙ6VÁG&≈&ˆ˜BÜ˜FñˆÁ2“∑“í∞¢6ˆÁ7B≤f˜&6R“f«6R““˜FñˆÁ3∞¢ñbÇ6‰÷ÊvTFFÇí«¬G&ófT66W75Fˆ∂V‚í&WGW&„∞¢6ˆÁ7B÷ñw&Fñˆ‰∂Wí“G¥ƒTt5ïÙE$ïdUÙ‘îu$DîÙÂÙ¥Uó”¢G¥4TÂE$≈ÙE$ïdUı$ÙıEÙdÙƒDU%ÙîG÷∞¢ñbÇf˜&6Rbb6W76ñˆÂ7F˜&vRÊvWDóFV“Ü÷ñw&Fñˆ‰∂Wíí””“'G'VR"í&WGW&„∞¢G'í∞¢vóBVÁ7W&TG&ófTfˆ∆FW'2Çì∞¢6ˆÁ7B∆Vv7î6ˆÁFñÊW$ñB“vóBvWD˜$7&VFTG&ófTfˆ∆FW"Ñ4TÂE$≈ÙE$ïdUÙƒTt5ïÙdÙƒDU%Ù‰‘R¬4TÂE$≈ÙE$ïdUı$ÙıEÙdÙƒDU%ÙîBì∞¢f˜"Ü6ˆÁ7B∆Vv7îÊ÷RˆbƒTt5ïÙE$ïdUı$ÙıEÙdÙƒDU%Ù‰‘U2í∞¢6ˆÁ7B∆Vv7îfˆ∆FW'2“vóBfñÊDG&ófTfˆ∆FW'4'îÊ÷RÜ∆Vv7îÊ÷Rì∞¢f˜"Ü6ˆÁ7B∆Vv7îfˆ∆FW"ˆb∆Vv7îfˆ∆FW'2í∞¢ñbÇ∆Vv7îfˆ∆FW"ÊñB«¬∆Vv7îfˆ∆FW"ÊñB””“4TÂE$≈ÙE$ïdUı$ÙıEÙdÙƒDU%ÙîB«¬∆Vv7îfˆ∆FW"ÊñB””“∆Vv7î6ˆÁFñÊW$ñBí6ˆÁFñÁVS∞¢6ˆÁ7B&VÁG2“'&íÊó4'&íÜ∆Vv7îfˆ∆FW"Á&VÁG2íÚ∆Vv7îfˆ∆FW"Á&VÁG2¢µ”∞¢ñbá&VÁG2ÊñÊ6«VFW2Ü∆Vv7î6ˆÁFñÊW$ñBíí6ˆÁFñÁVS∞¢vóB÷˜fTG&ófTfñ∆UFÙfˆ∆FW"Ü∆Vv7îfˆ∆FW"ÊñB¬∆Vv7î6ˆÁFñÊW$ñB¬&VÁG2ì∞¢–¢–¢6W76ñˆÂ7F˜&vRÁ6WDóFV“Ü÷ñw&Fñˆ‰∂Wí¬'G'VR"ì∞¢“6F6ÇÜW'&˜"í∞¢6ˆÁ6ˆ∆RÁv&‚Ç$÷ñw&¶ñˆÊRFFíG&ófRfV66ÜíÊˆ‚6ˆ◊∆WFF¢"¬W'&˜"ì∞¢–ß–†¶gVÊ7Fñˆ‚vWD6ˆ÷÷W766ÜVWDÜVFW'2Çí∞¢&WGW&‚µ∞¢$6ˆ÷÷W76"¬$6ÁFñW&R"¬$Fó7G&WGFÚ"¬$îB4"¬$FVÊˆ÷ñÊ¶ñˆÊR"¬$6ˆ◊VÊR"¬$ñÊFó&óß¶Ú"¬%fˆ6R&ñfW&ñ÷VÁFÚ"¿¢$6ˆFñ6R&Wß¶Ú"¬%6f∆6í"¬$g&WVVÁ¶ÊÁV"¬%Fóˆ∆ˆvññÁFW'fVÁFÚ"¬$∆f˜&¶ñˆÊí&ñ6ÜñW7FR"¿¢$u2í"¬$u2Ç"¬%FóÚ÷ÁWFVÁ¶ñˆÊR"¬%7FFÚ"¬$FFW6V7W¶ñˆÊR"¬$˜&W6V7W¶ñˆÊR"¬$W6VwVóFÚF"¬$V÷ñ¬˜W&F˜&R ¢’”∞ß–†¶gVÊ7Fñˆ‚'Vñ∆E6ÜVWE&˜w4g&ˆ‘FˆÊTñ◊ñÁFíÜFˆÊTñ◊ñÁFí¬6ˆ÷÷W76Ê÷R¬˜W&F˜$V÷ñ¬“""í∞¢&WGW&‚FˆÊTñ◊ñÁFíÊf∆D÷ÇÜñ◊ñÁFÚí”‚'Vñ∆E&˜w4f˜$V6Ñ6ˆFñ6U&Wß¶ÚÜñ◊ñÁFÚííÊ÷Çá&˜tFFí”‚∞¢6ˆÁ7BFˆÊTñÊfÚ“f˜&÷DFˆÊTFFUFñ÷Rá&˜tFFÊFˆÊTBì∞¢&WGW&‚∞¢6ˆ÷÷W76Ê÷R«¬""¿¢&˜tFFÊ6ÁFñW&U&ñv«¬""¿¢&˜tFFÊFó7G&WGFÚ«¬""¿¢&˜tFFÊñE6«¬""¿¢&˜tFFÊFVÊˆ÷ñÊ¶ñˆÊR«¬""¿¢&˜tFFÊ6ˆ◊VÊR«¬""¿¢&˜tFFÊñÊFó&óß¶Ú«¬""¿¢&˜tFFÁfˆ6U&ñfW&ñ÷VÁFÚ«¬""¿¢&˜tFFÊ6ˆFñ6U&Wß¶ı6ñÊvˆ∆Ú«¬&˜tFFÊ6ˆFñ6U&Wß¶Ú«¬""¿¢&˜tFFÁ6f∆6í«¬""¿¢&˜tFFÊg&WVVÁ¶ÊÁV«¬""¿¢&˜tFFÁFóˆ∆ˆvññÁFW'fVÁFÚ«¬""¿¢&˜tFFÊ∆f˜&¶ñˆÊï&ñ6ÜñW7FR«¬""¿¢&˜tFFÊw5íÛÚ""¿¢&˜tFFÊw5ÇÛÚ""¿¢&˜tFFÁFóÙ÷ÁWFVÁ¶ñˆÊR«¬6∆76ñgïFóÙ÷ÁWFVÁ¶ñˆÊRá&˜tFFÊ6ˆFñ6U&Wß¶Úí¿¢$fGFÚ"¿¢FˆÊTñÊfÚÊFFR¿¢FˆÊTñÊfÚÁFñ÷R¿¢&˜tFFÊFˆÊT'í«¬""¿¢˜W&F˜$V÷ñ¬«¬" ¢”∞¢“ì∞ß–†¶7ñÊ2gVÊ7Fñˆ‚7ñÊ46ˆ÷÷W76FˆÊTñ◊ñÁFïFÙG&ófU6ÜVWBÜ6ˆ÷÷W76ñB“6V∆V7FVD6ˆ÷÷W76ñB¬f∆∆&6¥6ˆ÷÷W76Ê÷R“6V∆V7FVD6ˆ÷÷W76Ê÷Rí∞¢ñbÇG&ófT66W75Fˆ∂V‚í∞¢Fá&˜rÊWrW'&˜"Ç$G&ófR6VÁG&∆óß¶FÚÊˆ‚Fó7ˆÊñ&ñ∆R‚"ì∞¢–¢ñbÇ6ˆ÷÷W76ñBí&WGW&„∞¢ñbÇG&ófU&W˜'G4fˆ∆FW$ñBívóBVÁ7W&TG&ófTfˆ∆FW'2Çì∞†¢6ˆÁ7B6ˆ÷÷W76“6ˆ÷÷W76T'îñBÊvWBÜ6ˆ÷÷W76ñBí«¬∑”∞¢6ˆÁ7B6ˆ÷÷W76Ê÷R“6ˆ÷÷W76ÊÊˆ÷R«¬f∆∆&6¥6ˆ÷÷W76Ê÷R«¬$6ˆ÷÷W76#∞¢6ˆÁ7B6Ê6Ü˜B“vóBF"Ê6ˆ∆∆V7Fñˆ‚Ç&6ˆ÷÷W76R"íÊFˆ2Ü6ˆ÷÷W76ñBíÊ6ˆ∆∆V7Fñˆ‚Ç&ñ◊ñÁFí"íÊvWBÇì∞¢6ˆÁ7B&tñ◊ñÁFí“6Ê6Ü˜BÊFˆ72Ê÷ÇÜFˆ2í”‚á≤ñC¢Fˆ2ÊñB¬‚‚ÊFˆ2ÊFFÇí“íì∞¢6ˆÁ7BFˆÊTñ◊ñÁFí“6ˆ÷&ñÊTñ◊ñÁFîf˜%fñWrá&tñ◊ñÁFííÊfñ«FW"ÇÜóFV“í”‚óFV“ÊFˆÊRì∞¢6ˆÁ7B&˜w2“'Vñ∆E6ÜVWE&˜w4g&ˆ‘FˆÊTñ◊ñÁFíÜFˆÊTñ◊ñÁFí¬6ˆ÷÷W76Ê÷R¬7W'&VÁEW6W#ÚÊV÷ñ¬«¬""ì∞¢6ˆÁ7B7&VG6ÜVWB“vóBvWD˜$7&VFT6ˆ÷÷W767&VG6ÜVWBÜ6ˆ÷÷W76ñB¬6ˆ÷÷W76Ê÷Rì∞†¢vóBG&ófTîfWF6ÇÜáGG3¢Ú˜6ÜVWG2Êvˆˆv∆Vó2Ê6ˆ“˜cB˜7&VG6ÜVWG2ÚG∑7&VG6ÜVWBÊñG“˜f«VW2Ù•£¶6∆V&¬∞¢÷WFÜˆC¢%ı5B"¿¢ÜVFW'3¢≤$6ˆÁFVÁB’GóR#¢&∆ñ6Fñˆ‚ˆß6ˆ‚"“¿¢&ˆGì¢•4Ù‚Á7G&ñÊvñgíá∑“ê¢“ì∞†¢vóBG&ófTîfWF6ÇÜáGG3¢Ú˜6ÜVWG2Êvˆˆv∆Vó2Ê6ˆ“˜cB˜7&VG6ÜVWG2ÚG∑7&VG6ÜVWBÊñG“˜f«VW2Ù¶VÊC˜f«VTñÁWD˜Fñˆ„’$v¬∞¢÷WFÜˆC¢%ı5B"¿¢ÜVFW'3¢≤$6ˆÁFVÁB’GóR#¢&∆ñ6Fñˆ‚ˆß6ˆ‚"“¿¢&ˆGì¢•4Ù‚Á7G&ñÊvñgíá∞¢f«VW3¢≤‚‚ÊvWD6ˆ÷÷W766ÜVWDÜVFW'2Çí¬‚‚Á&˜w5–¢“ê¢“ì∞ß–†¶7ñÊ2gVÊ7Fñˆ‚vWD˜$7&VFT6VÁG&ƒG&ófUGóTfˆ∆FW"Ü6ˆ÷÷W76Ê÷R¬G&ófUGóRí∞¢vóBVÁ7W&TG&ófTfˆ∆FW'2Çì∞¢6ˆÁ7B6ˆ÷÷W76fˆ∆FW$ñB“vóBvWD˜$7&VFTG&ófTfˆ∆FW"ÜÊ˜&÷∆ó¶TG&ófTfˆ∆FW$Ê÷RÜ6ˆ÷÷W76Ê÷R¬4TÂE$≈ÙE$ïdUÙDTdT≈EÙ4Ù‘‘U54í¬G&ófU&ˆ˜Dfˆ∆FW$ñBì∞¢&WGW&‚vWD˜$7&VFTG&ófTfˆ∆FW"ÜÊ˜&÷∆ó¶TG&ófTfˆ∆FW$Ê÷RÜG&ófUGóR¬$UÖı%B"íÁFıWW$66RÇí¬6ˆ÷÷W76fˆ∆FW$ñBì∞ß–†¶7ñÊ2gVÊ7Fñˆ‚vWD˜$7&VFT6ˆ÷÷W767&VG6ÜVWBÜ6ˆ÷÷W76ñB¬6ˆ÷÷W76Ê÷Rí∞¢6ˆÁ7B66ÜVDñB“6ˆ÷÷W766ÜVWD66ÜRÊvWBÜ6ˆ÷÷W76ñBì∞¢ñbÜ66ÜVDñBí&WGW&‚≤ñC¢66ÜVDñB”∞†¢6ˆÁ7B7&VG6ÜVWDfˆ∆FW$ñB“vóBvWD˜$7&VFT6VÁG&ƒG&ófUGóTfˆ∆FW"Ü6ˆ÷÷W76Ê÷R¬$UÖı%B"ì∞¢6ˆÁ7B6ˆ÷÷W76FF“6ˆ÷÷W76T'îñBÊvWBÜ6ˆ÷÷W76ñBí«¬∑”∞¢6ˆÁ7B6ˆÊfñwW&VE6ÜVWDñB“7G&ñÊrÜ6ˆ÷÷W76FFÁ6ÜVWE7&VG6ÜVWDñB«¬""íÁG&ñ“Çì∞¢ñbÜ6ˆÊfñwW&VE6ÜVWDñBí∞¢G'í∞¢6ˆÁ7BWÜó7FñÊu6ÜVWB“vóBG&ófTîfWF6ÇÜáGG3¢Ú˜wwrÊvˆˆv∆Vó2Ê6ˆ“ˆG&ófR˜c2ˆfñ∆W2ÚG∂6ˆÊfñwW&VE6ÜVWDñG”ˆfñV∆G3÷ñB∆Ê÷R∆÷ñ÷UGóR«&VÁG6¬≤÷WFÜˆC¢$tUB"“ì∞¢ñbÜWÜó7FñÊu6ÜVWBbbWÜó7FñÊu6ÜVWBÊñBí∞¢6ˆÁ7B&VÁG2“'&íÊó4'&íÜWÜó7FñÊu6ÜVWBÁ&VÁG2íÚWÜó7FñÊu6ÜVWBÁ&VÁG2¢µ”∞¢ñbÇ&VÁG2ÊñÊ6«VFW2á7&VG6ÜVWDfˆ∆FW$ñBíí∞¢6ˆÁ7B&V÷˜fU&VÁG2“&VÁG2Êfñ«FW"Ñ&ˆˆ∆V‚íÊ¶ˆñ‚Ç"¬"ì∞¢6ˆÁ7B÷˜fU&◊2“ÊWrU$≈6V&6Ö&◊2á≤FE&VÁG3¢7&VG6ÜVWDfˆ∆FW$ñB¬fñV∆G3¢&ñB«&VÁG2"“ì∞¢ñbá&V÷˜fU&VÁG2í÷˜fU&◊2Á6WBÇ'&V÷˜fU&VÁG2"¬&V÷˜fU&VÁG2ì∞¢vóBG&ófTîfWF6ÇÜáGG3¢Ú˜wwrÊvˆˆv∆Vó2Ê6ˆ“ˆG&ófR˜c2ˆfñ∆W2ÚG∂6ˆÊfñwW&VE6ÜVWDñG”ÚG∂÷˜fU&◊2ÁFı7G&ñÊrÇó÷¬∞¢÷WFÜˆC¢%D4Ç"¿¢ÜVFW'3¢≤$6ˆÁFVÁB’GóR#¢&∆ñ6Fñˆ‚ˆß6ˆ‚"“¿¢&ˆGì¢•4Ù‚Á7G&ñÊvñgíá∑“ê¢“ì∞¢–¢6ˆ÷÷W766ÜVWD66ÜRÁ6WBÜ6ˆ÷÷W76ñB¬6ˆÊfñwW&VE6ÜVWDñBì∞¢&WGW&‚≤ñC¢6ˆÊfñwW&VE6ÜVWDñB”∞¢–¢“6F6ÇÜW'&˜"í∞¢6ˆÁ6ˆ∆RÁv&‚Ç$fˆv∆ñÚ6ˆÊfñwW&FÚÊˆ‚ú;íFó7ˆÊñ&ñ∆R¬&˜fÚ&ñ7&V¶ñˆÊRWFˆ÷Fñ6¢"¬W'&˜"ì∞¢vóBF"Ê6ˆ∆∆V7Fñˆ‚ÜvWD6ˆ÷÷W76T6ˆ∆∆V7Fñˆ‰Ê÷RÇííÊFˆ2Ü6ˆ÷÷W76ñBíÁ6WBá∞¢6ÜVWE7&VG6ÜVWDñC¢fó&V&6RÊfó&W7F˜&R‰fñV∆Ef«VRÊFV∆WFRÇí¿¢6ÜVWEWFFVDC¢fó&V&6RÊfó&W7F˜&R‰fñV∆Ef«VRÁ6W'fW%Fñ÷W7F◊Çê¢“¬≤÷W&vS¢G'VR“ì∞¢–¢–†¢6ˆÁ7B6fTÊ÷R“W66TG&ófUVW'ïf«VRÜ6ˆ÷÷W76Ê÷Rì∞¢6ˆÁ7BVW'í“∞¢&÷ñ÷UGóS“v∆ñ6Fñˆ‚˜fÊBÊvˆˆv∆R÷2Á7&VG6ÜVWBr"¿¢'G&6ÜVC÷f«6R"¿¢rG∑7&VG6ÜVWDfˆ∆FW$ñG“rñ‚&VÁG6¿¢Ê÷S“t6ˆ÷÷W76“G∑6fTÊ÷W“v ¢“Ê¶ˆñ‚Ç"ÊB"ì∞†¢6ˆÁ7B6V&6ÖW&¬“áGG3¢Ú˜wwrÊvˆˆv∆Vó2Ê6ˆ“ˆG&ófR˜c2ˆfñ∆W3˜“G∂VÊ6ˆFUU$î6ˆ◊ˆÊVÁBáVW'íó“ffñV∆G3÷fñ∆W2ÜñB∆Ê÷R∆7&VFVEFñ÷Ríf˜&FW$'ì÷7&VFVEFñ÷RgvU6ó¶S”∞¢6ˆÁ7B6V&6Ö&W7ˆÁ6R“vóBG&ófTîfWF6Çá6V&6ÖW&¬¬≤÷WFÜˆC¢$tUB"“ì∞†¢ñbÑ'&íÊó4'&íá6V&6Ö&W7ˆÁ6RÊfñ∆W2íbb6V&6Ö&W7ˆÁ6RÊfñ∆W2Ê∆VÊwFÇ‚í∞¢6ˆÁ7BWÜó7FñÊr“6V&6Ö&W7ˆÁ6RÊfñ∆W5≥”∞¢6ˆ÷÷W766ÜVWD66ÜRÁ6WBÜ6ˆ÷÷W76ñB¬WÜó7FñÊrÊñBì∞¢vóBF"Ê6ˆ∆∆V7Fñˆ‚ÜvWD6ˆ÷÷W76T6ˆ∆∆V7Fñˆ‰Ê÷RÇííÊFˆ2Ü6ˆ÷÷W76ñBíÁ6WBá∞¢6ÜVWE7&VG6ÜVWDñC¢WÜó7FñÊrÊñB¿¢6ÜVWEWFFVDC¢fó&V&6RÊfó&W7F˜&R‰fñV∆Ef«VRÁ6W'fW%Fñ÷W7F◊Çê¢“¬≤÷W&vS¢G'VR“ì∞¢&WGW&‚≤ñC¢WÜó7FñÊrÊñB”∞¢–†¢6ˆÁ7B7&VFVB“vóBG&ófTîfWF6ÇÇ&áGG3¢Ú˜wwrÊvˆˆv∆Vó2Ê6ˆ“ˆG&ófR˜c2ˆfñ∆W2"¬∞¢÷WFÜˆC¢%ı5B"¿¢ÜVFW'3¢≤$6ˆÁFVÁB’GóR#¢&∆ñ6Fñˆ‚ˆß6ˆ‚"“¿¢&ˆGì¢•4Ù‚Á7G&ñÊvñgíá∞¢Ê÷S¢6ˆ÷÷W76“G∂6ˆ÷÷W76Ê÷W÷¿¢÷ñ÷UGóS¢&∆ñ6Fñˆ‚˜fÊBÊvˆˆv∆R÷2Á7&VG6ÜVWB"¿¢&VÁG3¢∑7&VG6ÜVWDfˆ∆FW$ñE–¢“ê¢“ì∞†¢6ˆÁ7BÜVFW'2“vWD6ˆ÷÷W766ÜVWDÜVFW'2Çì∞†¢vóBG&ófTîfWF6ÇÜáGG3¢Ú˜6ÜVWG2Êvˆˆv∆Vó2Ê6ˆ“˜cB˜7&VG6ÜVWG2ÚG∂7&VFVBÊñG“˜f«VW2Ù¶VÊC˜f«VTñÁWD˜Fñˆ„’$v¬∞¢÷WFÜˆC¢%ı5B"¿¢ÜVFW'3¢≤$6ˆÁFVÁB’GóR#¢&∆ñ6Fñˆ‚ˆß6ˆ‚"“¿¢&ˆGì¢•4Ù‚Á7G&ñÊvñgíá≤f«VW3¢ÜVFW'2“ê¢“ì∞†¢6ˆ÷÷W766ÜVWD66ÜRÁ6WBÜ6ˆ÷÷W76ñB¬7&VFVBÊñBì∞¢vóBF"Ê6ˆ∆∆V7Fñˆ‚ÜvWD6ˆ÷÷W76T6ˆ∆∆V7Fñˆ‰Ê÷RÇííÊFˆ2Ü6ˆ÷÷W76ñBíÁ6WBá∞¢6ÜVWE7&VG6ÜVWDñC¢7&VFVBÊñB¿¢6ÜVWEWFFVDC¢fó&V&6RÊfó&W7F˜&R‰fñV∆Ef«VRÁ6W'fW%Fñ÷W7F◊Çê¢“¬≤÷W&vS¢G'VR“ì∞¢&WGW&‚≤ñC¢7&VFVBÊñB”∞ß–†¶gVÊ7Fñˆ‚&VD&∆ˆ$4&6ScBÜ&∆ˆ"í∞¢&WGW&‚ÊWr&ˆ÷ó6RÇá&W6ˆ«fR¬&V¶V7Bí”‚∞¢6ˆÁ7B&VFW"“ÊWrfñ∆U&VFW"Çì∞¢&VFW"ÊˆÊ∆ˆB“Çí”‚∞¢6ˆÁ7B&W7V«B“7G&ñÊrá&VFW"Á&W7V«B«¬""ì∞¢&W6ˆ«fRá&W7V«BÊñÊ6«VFW2Ç"¬"íÚ&W7V«BÁ7∆óBÇ"¬"íÁ˜Çí¢&W7V«Bì∞¢”∞¢&VFW"ÊˆÊW'&˜"“Çí”‚&V¶V7Bá&VFW"ÊW'&˜"«¬ÊWrW'&˜"Ç$∆WGGW&fñ∆RÊˆ‚&óW66óF‚"íì∞¢&VFW"Á&VD4FFU$¬Ü&∆ˆ"ì∞¢“ì∞ß–†¶7ñÊ2gVÊ7Fñˆ‚W∆ˆD&∆ˆ%Fá&˜VvÑ6VÁG&ƒ&6∂VÊBÜ&∆ˆ"¬fñ∆TÊ÷R¬÷ñ÷UGóR¬fˆ∆FW$ñB¬˜FñˆÁ2“∑“í∞¢ñbÇó46VÁG&ƒG&ófT6ˆÊfñwW&VBÇíí∞¢Fá&˜rÊWrW'&˜"ÜvWD6VÁG&ƒG&ófTÊ˜D6ˆÊfñwW&VD÷W76vRÇíì∞¢–¢ñbÇgVÊ7FñˆÁ2«¬GóVˆbgVÊ7FñˆÁ2ÊáGG46∆∆&∆R”“&gVÊ7Fñˆ‚"í∞¢Fá&˜rÊWrW'&˜"Ç$&6∂VÊBfó&V&6RW"W∆ˆB6VÁG&∆óß¶FÚÊˆ‚Fó7ˆÊñ&ñ∆R‚"ì∞¢–¢6ˆÁ7B&6ScB“vóB&VD&∆ˆ$4&6ScBÜ&∆ˆ"ì∞¢6ˆÁ7BW∆ˆD6VÁG&ƒG&ófTfñ∆R“gVÊ7FñˆÁ2ÊáGG46∆∆&∆RÇ'W∆ˆD6VÁG&ƒG&ófTfñ∆R"ì∞¢6ˆÁ7B&W7V«B“vóBW∆ˆD6VÁG&ƒG&ófTfñ∆Rá∞¢fñ∆TÊ÷S¢Ê˜&÷∆ó¶TG&ófTfˆ∆FW$Ê÷RÜfñ∆TÊ÷R¬&fñ∆R"í¿¢÷ñ÷UGóS¢÷ñ÷UGóR«¬&∆ñ6Fñˆ‚ˆˆ7FWB◊7G&V“"¿¢&6ScB¿¢6ˆ÷÷W76Ê÷S¢vWD7W'&VÁDG&ófT6ˆ÷÷W76Ê÷RÜ˜FñˆÁ2í¿¢G&ófUGóS¢ñÊfW$6VÁG&ƒG&ófUGóRÜfˆ∆FW$ñB¬˜FñˆÁ2ê¢“ì∞¢6ˆÁ7BFF“&W7V«CÚÊFF«¬∑”∞¢&WGW&‚∞¢fñ∆TñC¢FFÊfñ∆TñB«¬""¿¢vV%fñWt∆ñÊ≥¢FFÁvV%fñWt∆ñÊ≤«¬""¿¢Fó&V7EW&√¢" ¢”∞ß–†¶7ñÊ2gVÊ7Fñˆ‚W∆ˆD&∆ˆ$Fó&V7EFÙF÷ñ‰G&ófRÜ&∆ˆ"¬fñ∆TÊ÷R¬÷ñ÷UGóR¬fˆ∆FW$ñB¬˜FñˆÁ2“∑“í∞¢ñbÇG&ófT66W75Fˆ∂V‚í∞¢Fá&˜rÊWrW'&˜"ÜvWD6VÁG&ƒG&ófTÊ˜D6ˆÊfñwW&VD÷W76vRÇíì∞¢–¢vóBVÁ7W&TG&ófTfˆ∆FW'2Çì∞¢6ˆÁ7B≤6ñvÊ¬“ÁV∆¬““˜FñˆÁ3∞¢6ˆÁ7B6ˆ÷÷W76fˆ∆FW$ñB“vóBvWD˜$7&VFTG&ófTfˆ∆FW"ÜvWD7W'&VÁDG&ófT6ˆ÷÷W76Ê÷RÜ˜FñˆÁ2í¬4TÂE$≈ÙE$ïdUı$ÙıEÙdÙƒDU%ÙîBì∞¢6ˆÁ7BGóTfˆ∆FW$ñB“vóBvWD˜$7&VFTG&ófTfˆ∆FW"ÜñÊfW$6VÁG&ƒG&ófUGóRÜfˆ∆FW$ñB¬˜FñˆÁ2í¬6ˆ÷÷W76fˆ∆FW$ñBì∞¢6ˆÁ7B÷WFFF“∞¢Ê÷S¢Ê˜&÷∆ó¶TG&ófTfˆ∆FW$Ê÷RÜfñ∆TÊ÷R¬&fñ∆R"í¿¢÷ñ÷UGóS¢÷ñ÷UGóR«¬&∆ñ6Fñˆ‚ˆˆ7FWB◊7G&V“"¿¢&VÁG3¢∑GóTfˆ∆FW$ñE–¢”∞¢6ˆÁ7Bf˜&““ÊWrf˜&‘FFÇì∞¢f˜&“ÊVÊBÇ&÷WFFF"¬ÊWr&∆ˆ"Ö¥•4Ù‚Á7G&ñÊvñgíÜ÷WFFFï“¬≤GóS¢&∆ñ6Fñˆ‚ˆß6ˆ‚"“íì∞¢f˜&“ÊVÊBÇ&fñ∆R"¬&∆ˆ"¬÷WFFFÊÊ÷Rì∞¢6ˆÁ7BW∆ˆFVB“vóBG&ófTîfWF6ÇÇ&áGG3¢Ú˜wwrÊvˆˆv∆Vó2Ê6ˆ“˜W∆ˆBˆG&ófR˜c2ˆfñ∆W3˜W∆ˆEGóS÷◊V«Fó'BffñV∆G3÷ñB∆Ê÷R«vV%fñWt∆ñÊ≤«vV$6ˆÁFVÁD∆ñÊ≤"¬∞¢÷WFÜˆC¢%ı5B"¿¢&ˆGì¢f˜&“¿¢6ñvÊ¿¢“ì∞¢&WGW&‚∞¢fñ∆TñC¢W∆ˆFVBÊñB«¬""¿¢vV%fñWt∆ñÊ≥¢W∆ˆFVBÁvV%fñWt∆ñÊ≤«¬""¿¢Fó&V7EW&√¢" ¢”∞ß–†¶7ñÊ2gVÊ7Fñˆ‚W∆ˆD&∆ˆ%FÙG&ófRÜ&∆ˆ"¬fñ∆TÊ÷R¬÷ñ÷UGóR¬fˆ∆FW$ñB¬˜FñˆÁ2“∑“í∞¢ñbÇó46VÁG&ƒG&ófT6ˆÊfñwW&VBÇíbbG&ófT66W75Fˆ∂V‚í∞¢Fá&˜rÊWrW'&˜"ÜvWD6VÁG&ƒG&ófTÊ˜D6ˆÊfñwW&VD÷W76vRÇíì∞¢–¢ñbÜ6‰÷ÊvTFFÇíbbG&ófT66W75Fˆ∂V‚í∞¢G'í∞¢&WGW&‚vóBW∆ˆD&∆ˆ$Fó&V7EFÙF÷ñ‰G&ófRÜ&∆ˆ"¬fñ∆TÊ÷R¬÷ñ÷UGóR¬fˆ∆FW$ñB¬˜FñˆÁ2ì∞¢“6F6ÇÜW'&˜"í∞¢ñbÇgVÊ7FñˆÁ2«¬GóVˆbgVÊ7FñˆÁ2ÊáGG46∆∆&∆R”“&gVÊ7Fñˆ‚"íFá&˜rW'&˜#∞¢6ˆÁ6ˆ∆RÁv&‚Ç%W∆ˆBFó&WGFÚF÷ñ‚Êˆ‚&óW66óFÚ¬&˜fÚ&6∂VÊB6VÁG&∆óß¶FÛ¢"¬W'&˜"ì∞¢–¢–¢&WGW&‚W∆ˆD&∆ˆ%Fá&˜VvÑ6VÁG&ƒ&6∂VÊBÜ&∆ˆ"¬fñ∆TÊ÷R¬÷ñ÷UGóR¬fˆ∆FW$ñB¬˜FñˆÁ2ì∞ß–†¶7ñÊ2gVÊ7Fñˆ‚&6∑W7VG&U6Ê6Ü˜EFÙG&ófRÜFFT∂Wí¬7VG&ñ∆ˆBí∞¢ñbÇG&ófT66W75Fˆ∂V‚í&WGW&„∞¢ñbÇG&ófU7VG&Tfˆ∆FW$ñBívóBVÁ7W&TG&ófTfˆ∆FW'2Çì∞¢6ˆÁ7BWá˜'DFF“vóB'Vñ∆D&6∑Wñ∆ˆBÜFFT∂Wí¬7VG&ñ∆ˆBì∞¢6ˆÁ7B&∆ˆ"“ÊWr&∆ˆ"Ö¥•4Ù‚Á7G&ñÊvñgíÜWá˜'DFF¬ÁV∆¬¬"ï“¬≤GóS¢&∆ñ6Fñˆ‚ˆß6ˆ‚"“ì∞¢6ˆÁ7B6ˆ÷÷W76∆&V¬“7G&ñÊrá7VG&ñ∆ˆBÊ6ˆ÷÷W76Êˆ÷R«¬$6ˆ÷÷W76"íÁ&W∆6RÇıµÂ«u¬’“≤ˆr¬%Ú"ì∞¢6ˆÁ7Bfñ∆TÊ÷R“7VG&UÚG∂FFT∂Wó’ÚG∂6ˆ÷÷W76∆&V«“Êß6ˆÊ∞¢vóBW∆ˆD&∆ˆ%FÙG&ófRÜ&∆ˆ"¬fñ∆TÊ÷R¬&∆ñ6Fñˆ‚ˆß6ˆ‚"¬G&ófU7VG&Tfˆ∆FW$ñB¬≤G&ófUGóS¢$UÖı%B"¬6ˆ÷÷W76Ê÷S¢7VG&ñ∆ˆBÊ6ˆ÷÷W76Êˆ÷R«¬%7VG&R"“ì∞ß–†¶7ñÊ2gVÊ7Fñˆ‚'Vñ∆D&6∑Wñ∆ˆBÜFFT∂Wí¬7VG&ñ∆ˆBí∞¢6ˆÁ7B∂6ˆ÷÷W76U6Ê6Ü˜B¬W'6ˆÊ∆U6Ê6Ü˜B¬÷Wß¶ï6Ê6Ü˜B¬7VG&T6˜'&VÁFï6Ê6Ü˜B¬7VG&U7F˜&ñ6ı6Ê6Ü˜E““vóB&ˆ÷ó6RÊ∆¬Ö∞¢F"Ê6ˆ∆∆V7Fñˆ‚ÜvWD6ˆ÷÷W76T6ˆ∆∆V7Fñˆ‰Ê÷RÇííÊvWBÇí¿¢F"Ê6ˆ∆∆V7Fñˆ‚ÜvWEW'6ˆÊ∆T6ˆ∆∆V7Fñˆ‰Ê÷RÇííÊvWBÇí¿¢F"Ê6ˆ∆∆V7Fñˆ‚ÜvWD÷Wß¶î6ˆ∆∆V7Fñˆ‰Ê÷RÇííÊvWBÇí¿¢F"Ê6ˆ∆∆V7Fñˆ‚ÜvWE7VG&T7W'&VÁD6ˆ∆∆V7Fñˆ‰Ê÷RÇííÊvWBÇí¿¢F"Ê6ˆ∆∆V7Fñˆ‚ÜvWE7VG&TÜó7F˜'î6ˆ∆∆V7Fñˆ‰Ê÷RÇííÁvÜW&RÇ&FFT∂Wí"¬#”“"¬FFT∂WííÊvWBÇê¢“ì∞¢&WGW&‚∞¢Wá˜'FVDC¢ÊWrFFRÇíÁFÙï4ı7G&ñÊrÇí¿¢Wá˜'FVD'ì¢Ü7W'&VÁEW6W"bb7W'&VÁEW6W"ÊV÷ñ¬íÚ7W'&VÁEW6W"ÊV÷ñ¬¢""¿¢6V∆V7FVDFFS¢FFT∂Wí¿¢6fVD6ˆ◊˜6óFñˆ„¢7VG&ñ∆ˆB¿¢6ˆ÷÷W76S¢6ˆ÷÷W76U6Ê6Ü˜BÊFˆ72Ê÷ÇÜFˆ2í”‚á≤ñC¢Fˆ2ÊñB¬‚‚ÊFˆ2ÊFFÇí“íí¿¢W'6ˆÊ∆S¢W'6ˆÊ∆U6Ê6Ü˜BÊFˆ72Ê÷ÇÜFˆ2í”‚á≤ñC¢Fˆ2ÊñB¬‚‚ÊFˆ2ÊFFÇí“íí¿¢÷Wß¶ì¢÷Wß¶ï6Ê6Ü˜BÊFˆ72Ê÷ÇÜFˆ2í”‚á≤ñC¢Fˆ2ÊñB¬‚‚ÊFˆ2ÊFFÇí“íí¿¢7VG&T6˜'&VÁFì¢7VG&T6˜'&VÁFï6Ê6Ü˜BÊFˆ72Ê÷ÇÜFˆ2í”‚á≤ñC¢Fˆ2ÊñB¬‚‚ÊFˆ2ÊFFÇí“íí¿¢7VG&U7F˜&ñ6Ùvñ˜&ÊÛ¢7VG&U7F˜&ñ6ı6Ê6Ü˜BÊFˆ72Ê÷ÇÜFˆ2í”‚á≤ñC¢Fˆ2ÊñB¬‚‚ÊFˆ2ÊFFÇí“íê¢”∞ß–†¶7ñÊ2gVÊ7Fñˆ‚&Vg&W6ÑG&ófT66W75Fˆ∂V‚Çí∞¢ñbÜG&ófUFˆ∂VÂ&Vg&W6Ö&ˆ÷ó6Rí&WGW&‚G&ófUFˆ∂VÂ&Vg&W6Ö&ˆ÷ó6S∞¢ñbÇWFÇÊ7W'&VÁEW6W"í&WGW&‚f«6S∞¢G&ófUFˆ∂VÂ&Vg&W6Ö&ˆ÷ó6R“Ü7ñÊ2Çí”‚∞¢G'í∞¢6ˆÁ7B&˜fñFW"“ÊWrfó&V&6RÊWFÇ‰vˆˆv∆TWFÖ&˜fñFW"Çì∞¢&˜fñFW"ÊFE66˜RÇ&áGG3¢Ú˜wwrÊvˆˆv∆Vó2Ê6ˆ“ˆWFÇˆG&ófRÊfñ∆R"ì∞¢&˜fñFW"Á6WD7W7Fˆ’&÷WFW'2á≤&ˆ◊C¢&ÊˆÊR"“ì∞¢6ˆÁ7B&W7V«B“vóBWFÇÁ6ñv‰ñÂvóFÖ˜Wá&˜fñFW"ì∞¢6ˆÁ7B66W75Fˆ∂V‚“WáG&7Dvˆˆv∆T66W75Fˆ∂V‚á&W7V«Bì∞¢ñbÇ66W75Fˆ∂V‚í&WGW&‚f«6S∞¢W'6ó7DG&ófT66W75Fˆ∂V‚Ü66W75Fˆ∂V‚ì∞¢vóBWFÙ6ˆÊÊV7DG&ófT'&ñFvRá≤Ê˜Fñgîˆ‰W'&˜#¢f«6R“ì∞¢&WGW&‚G'VS∞¢“6F6ÇÜW'&˜"í∞¢6ˆÁ6ˆ∆RÁv&‚Ç%&Vg&W6ÇWFˆ÷Fñ6ÚFˆ∂V‚G&ófRÊˆ‚&óW66óFÛ¢"¬W'&˜"ì∞¢&WGW&‚f«6S∞¢“fñÊ∆«í∞¢G&ófUFˆ∂VÂ&Vg&W6Ö&ˆ÷ó6R“ÁV∆√∞¢–¢“íÇì∞¢&WGW&‚G&ófUFˆ∂VÂ&Vg&W6Ö&ˆ÷ó6S∞ß–†¶7ñÊ2gVÊ7Fñˆ‚G&ófTîfWF6ÇáW&¬¬˜FñˆÁ2“∑“í∞¢ñbÇG&ófT66W75Fˆ∂V‚í∞¢Fá&˜rÊWrW'&˜"ÜvWD6VÁG&ƒG&ófTÊ˜D6ˆÊfñwW&VD÷W76vRÇíì∞¢–†¢6ˆÁ7BÜVFW'2“ÊWrÜVFW'2Ü˜FñˆÁ2ÊÜVFW'2«¬∑“ì∞¢ÜVFW'2Á6WBÇ$WFÜ˜&ó¶Fñˆ‚"¬&V&W"G∂G&ófT66W75Fˆ∂VÁ÷ì∞†¢6ˆÁ7B&W7ˆÁ6R“vóBfWF6ÖvóFÖFñ÷V˜WDÊE&WG'íáW&¬¬∞¢‚‚Ê˜FñˆÁ2¿¢ÜVFW'0¢“¬∞¢Fñ÷V˜WD◊3¢‰UEtı$µÙDTdT≈EıDî‘TıUEÙ’2¿¢&WG&ñW3¢ ¢“ì∞†¢ñbá&W7ˆÁ6RÁ7FGW2””“C«¬&W7ˆÁ6RÁ7FGW2””“C2í∞¢6ˆÁ7B&Vg&W6ÜVB“vóB&Vg&W6ÑG&ófT66W75Fˆ∂V‚Çì∞¢ñbá&Vg&W6ÜVBí∞¢&WGW&‚G&ófTîfWF6ÇáW&¬¬˜FñˆÁ2ì∞¢–¢∆ˆ6≈7F˜&vRÁ&V÷˜fTóFV“Ç&vˆˆv∆TG&ófT66W75Fˆ∂V‚"ì∞¢WFFTG&ófU7FGW2Üf«6Rì∞¢Fá&˜rÊWrW'&˜"Ç%6W76ñˆÊRG&ófR÷÷ñÊó7G&F˜&R66GWF‚6ˆ∆ÚF÷ñ‚FWfR&ñ6ˆ∆∆Vv&Rvˆˆv∆RG&ófR‚"ì∞¢–†¢ñbÇ&W7ˆÁ6RÊˆ≤í∞¢6ˆÁ7BFWáB“vóB&W7ˆÁ6RÁFWáBÇì∞¢Fá&˜rÊWrW'&˜"ÜW'&˜&Rvˆˆv∆RG&ófRÇG∑&W7ˆÁ6RÁ7FGW7“ì¢G∑FWáBÁ6∆ñ6RÉ¬Éó÷ì∞¢–†¢ñbá&W7ˆÁ6RÁ7FGW2””“#Bí&WGW&‚∑”∞¢&WGW&‚&W7ˆÁ6RÊß6ˆ‚Çì∞ß–†¶gVÊ7Fñˆ‚ó5&WG'ñ&∆TÊWGv˜&¥W'&˜"ÜW'&˜"í∞¢ñbÇW'&˜"í&WGW&‚f«6S∞¢ñbÜW'&˜"ÊÊ÷R””“$&˜'DW'&˜""í&WGW&‚G'VS∞¢ñbÜW'&˜"ñÁ7FÊ6VˆbGóTW'&˜"í&WGW&‚G'VS∞¢&WGW&‚ˆÊWGv˜&∑∆fWF6á∆fñ∆VG«Fñ÷V˜WBˆíÁFW7BÖ7G&ñÊrÜW'&˜"Ê÷W76vR«¬""íì∞ß–†¶gVÊ7Fñˆ‚6Ü˜V∆E&WG'îáGG7FGW2á7FGW2í∞¢&WGW&‚‰UEtı$µı$UE%î$ƒUı5DEU5Ù4ÙDU2ÊÜ2ÑÁV÷&W"á7FGW2íì∞ß–†¶gVÊ7Fñˆ‚vóBÜ◊2í∞¢&WGW&‚ÊWr&ˆ÷ó6RÇá&W6ˆ«fRí”‚6WEFñ÷V˜WBá&W6ˆ«fR¬◊2íì∞ß–†¶7ñÊ2gVÊ7Fñˆ‚fWF6ÖvóFÖFñ÷V˜WDÊE&WG'íáW&¬¬˜FñˆÁ2“∑“¬6ˆÊfñr“∑“í∞¢6ˆÁ7B&WG&ñW2“ÁV÷&W"Êó4fñÊóFRÜ6ˆÊfñrÁ&WG&ñW2íÚ÷FÇÊ÷ÇÉ¬6ˆÊfñrÁ&WG&ñW2í¢∞¢6ˆÁ7BFñ÷V˜WD◊2“ÁV÷&W"Êó4fñÊóFRÜ6ˆÊfñrÁFñ÷V˜WD◊2íÚ÷FÇÊ÷ÇÉ¬6ˆÊfñrÁFñ÷V˜WD◊2í¢‰UEtı$µÙDTdT≈EıDî‘TıUEÙ’3∞¢6ˆÁ7B&6TFV∆î◊2“ÁV÷&W"Êó4fñÊóFRÜ6ˆÊfñrÊ&6TFV∆î◊2íÚ÷FÇÊ÷ÇÉ¬6ˆÊfñrÊ&6TFV∆î◊2í¢c∞†¢∆WBGFV◊B“∞¢vÜñ∆RÜGFV◊B√“&WG&ñW2í∞¢6ˆÁ7B6ˆÁG&ˆ∆∆W"“ÊWr&˜'D6ˆÁG&ˆ∆∆W"Çì∞¢6ˆÁ7BWáFW&Ê≈6ñvÊ¬“˜FñˆÁ3ÚÁ6ñvÊ¬«¬ÁV∆√∞¢ñbÜWáFW&Ê≈6ñvÊ√ÚÊ&˜'FVBí∞¢Fá&˜rÊWrDÙ‘WÜ6WFñˆ‚Ç$˜W&¶ñˆÊRÊÁV∆∆F"¬$&˜'DW'&˜""ì∞¢–¢6ˆÁ7Bˆ‰WáFW&Êƒ&˜'B“Çí”‚6ˆÁG&ˆ∆∆W"Ê&˜'BÇì∞¢ñbÜWáFW&Ê≈6ñvÊ¬íWáFW&Ê≈6ñvÊ¬ÊFDWfVÁD∆ó7FVÊW"Ç&&˜'B"¬ˆ‰WáFW&Êƒ&˜'B¬≤ˆÊ6S¢G'VR“ì∞¢6ˆÁ7BFñ÷V˜WDñB“6WEFñ÷V˜WBÇÇí”‚6ˆÁG&ˆ∆∆W"Ê&˜'BÇí¬Fñ÷V˜WD◊2ì∞¢G'í∞¢6ˆÁ7B÷W&vVD˜FñˆÁ2“≤‚‚Ê˜FñˆÁ2¬6ñvÊ√¢6ˆÁG&ˆ∆∆W"Á6ñvÊ¬”∞¢6ˆÁ7B&W7ˆÁ6R“vóBfWF6ÇáW&¬¬÷W&vVD˜FñˆÁ2ì∞¢6∆V%Fñ÷V˜WBáFñ÷V˜WDñBì∞¢ñbÜWáFW&Ê≈6ñvÊ¬íWáFW&Ê≈6ñvÊ¬Á&V÷˜fTWfVÁD∆ó7FVÊW"Ç&&˜'B"¬ˆ‰WáFW&Êƒ&˜'Bì∞¢ñbÜGFV◊B¬&WG&ñW2bb6Ü˜V∆E&WG'îáGG7FGW2á&W7ˆÁ6RÁ7FGW2íí∞¢vóBvóBÜ&6TFV∆î◊2¢ÜGFV◊B≤íì∞¢GFV◊B≥“∞¢6ˆÁFñÁVS∞¢–¢&WGW&‚&W7ˆÁ6S∞¢“6F6ÇÜW'&˜"í∞¢6∆V%Fñ÷V˜WBáFñ÷V˜WDñBì∞¢ñbÜWáFW&Ê≈6ñvÊ¬íWáFW&Ê≈6ñvÊ¬Á&V÷˜fTWfVÁD∆ó7FVÊW"Ç&&˜'B"¬ˆ‰WáFW&Êƒ&˜'Bì∞¢ñbÜGFV◊B„“&WG&ñW2«¬ó5&WG'ñ&∆TÊWGv˜&¥W'&˜"ÜW'&˜"íí∞¢Fá&˜rW'&˜#∞¢–¢vóBvóBÜ&6TFV∆î◊2¢ÜGFV◊B≤íì∞¢GFV◊B≥“∞¢–¢–ß–†¶gVÊ7Fñˆ‚ó5&ˆw&÷÷¶ñˆÊUfó6ñ&∆UFÙ7W'&VÁEW6W"ÜóFV“í∞¢ñbÜ6‰÷ÊvTFFÇíí&WGW&‚G'VS∞¢6ˆÁ7BV÷ñ¬“Ê˜&÷∆ó¶TV÷ñ¬Ü7W'&VÁEW6W#ÚÊV÷ñ¬«¬""ì∞¢6ˆÁ7BFó7∆îÊ÷R“7G&ñÊrÜ7W'&VÁEW6W#ÚÊFó7∆îÊ÷R«¬""íÁG&ñ“ÇíÁFÙ∆˜vW$66RÇì∞¢6ˆÁ7B˜W&F˜'2“'&íÊó4'&íÜóFV”ÚÊ˜W&F˜&î6ˆñÁfˆ«FííÚóFV“Ê˜W&F˜&î6ˆñÁfˆ«Fí¢µ”∞¢6ˆÁ7BñÁfˆ«fVB“˜W&F˜'2Á6ˆ÷RÇÜVÁG'íí”‚∞¢6ˆÁ7B&r“7G&ñÊrÜVÁG'í«¬""íÁG&ñ“Çì∞¢ñbÇ&rí&WGW&‚f«6S∞¢6ˆÁ7BÊ˜&÷∆ó¶VDV÷ñ¬“Ê˜&÷∆ó¶TV÷ñ¬á&rì∞¢6ˆÁ7BÊ˜&÷∆ó¶VDÊ÷R“&rÁFÙ∆˜vW$66RÇì∞¢&WGW&‚ÜV÷ñ¬bbÊ˜&÷∆ó¶VDV÷ñ¬””“V÷ñ¬í«¬ÜFó7∆îÊ÷RbbÊ˜&÷∆ó¶VDÊ÷R””“Fó7∆îÊ÷Rì∞¢“ì∞¢ñbÇñÁfˆ«fVBí&WGW&‚f«6S∞¢6ˆÁ7BFFT∂Wí“7G&ñÊrÜóFV”ÚÊFF«¬""ì∞¢ñbÇFFT∂Wíí&WGW&‚f«6S∞¢6ˆÁ7BÊ˜r“ÊWrFFRÇì∞¢6ˆÁ7BF&vWB“ÊWrFFRÜG∂FFT∂Wó’C££ì∞¢6ˆÁ7BFˆFí“ÊWrFFRÜÊ˜rÊvWDgV∆≈ñV"Çí¬Ê˜rÊvWD÷ˆÁFÇÇí¬Ê˜rÊvWDFFRÇíì∞¢6ˆÁ7BFñfb“÷FÇÊf∆ˆ˜"ÇáF&vWB“FˆFííÚÉcCì∞¢ñbáF&vWBÊvWDFíÇí””“í&WGW&‚Fñfb√“2bbFñfb„“∞¢&WGW&‚Fñfb√“bbFñfb„“∞ß–†¶gVÊ7Fñˆ‚&ˆw&÷÷¶ñˆÊU&V÷ñÊFW$&FvRÜFFT∂Wí¬FóÚ“""í∞¢6ˆÁ7BFˆFí“ÊWrFFRÇì∞¢6ˆÁ7BF&vWB“ÊWrFFRÜG∂FFT∂Wó’C££ì∞¢6ˆÁ7BFñfb“÷FÇÊf∆ˆ˜"ÇáF&vWB“ÊWrFFRáFˆFíÊvWDgV∆≈ñV"Çí¬FˆFíÊvWD÷ˆÁFÇÇí¬FˆFíÊvWDFFRÇíííÚÉcCì∞¢6ˆÁ7BFóÙÊ˜&““7G&ñÊráFóÚ«¬""íÁFÙ∆˜vW$66RÇì∞¢ñbáFóÙÊ˜&“””“&fW&ñR"í∞¢ñbÜFñfb¬í&WGW&‚"#∞¢ñbÜFñfb””“í&WGW&‚/	¯˘n˚àÚfW&ñRˆvví#∞¢ñbÜFñfb””“í&WGW&‚/	¯˘n˚àÚfW&ñRFˆ÷Êí#∞¢ñbÜFñfb√“rí&WGW&‚	¯˘n˚àÚfW&ñRG&G∂Fñfg“vñ˜&Êñ∞¢&WGW&‚"#∞¢–¢ñbÜFñfb””“í&WGW&‚/	˘8Rˆvví#∞¢ñbÜFñfb””“í&WGW&‚/	˘8¬Fˆ÷Êí#∞¢6ˆÁ7BFí“F&vWBÊvWDFíÇì∞¢ñbÜFí””“bbFñfb„“2bbFñfb√“Rí&WGW&‚/	˘8¬&ˆw&÷÷¶ñˆÊR«VÊVL:¬#∞¢&WGW&‚"#∞ß–†¶gVÊ7Fñˆ‚&VÊFW%&ˆw&÷÷¶ñˆÊíÇí∞¢&Vg&W6ÑfW&ñU&ˆw&÷÷¶ñˆÊUVíÇì∞¢&VÊFW$fW&ñT∆ó7BÇì∞¢6ˆÁ7Bfó6ñ&∆R“&ˆw&÷÷¶ñˆÊíÊfñ«FW"Üó5&ˆw&÷÷¶ñˆÊUfó6ñ&∆UFÙ7W'&VÁEW6W"ì∞¢6ˆÁ7Bfñ«FW"“7G&ñÊráVíÁ&ˆw&÷÷¶ñˆÊTfñ«FW#ÚÁf«VR«¬&∆¬"ì∞¢6ˆÁ7BFˆFí“ÊWrFFRÇíÁFÙï4ı7G&ñÊrÇíÁ6∆ñ6RÉ¬ì∞¢6ˆÁ7BFˆ÷˜'&˜r“ÊWrFFRÑFFRÊÊ˜rÇí≤ÉcCíÁFÙï4ı7G&ñÊrÇíÁ6∆ñ6RÉ¬ì∞¢6ˆÁ7Bfñ«FW&VB“fó6ñ&∆RÊfñ«FW"ÇÜóFV“í”‚∞¢ñbÜfñ«FW"””“&ˆvví"í&WGW&‚óFV“ÊFF””“FˆFì∞¢ñbÜfñ«FW"””“&Fˆ÷Êí"í&WGW&‚óFV“ÊFF””“Fˆ÷˜'&˜s∞¢ñbÜfñ«FW"””“'&ˆw&÷÷FÚ"í&WGW&‚7G&ñÊrÜóFV“Á7FFÚ«¬""í””“%&ˆw&÷÷FÚ#∞¢ñbÜfñ«FW"””“'W&vVÁFR"í&WGW&‚7G&ñÊrÜóFV“Á&ñ˜&óF«¬""í””“%W&vVÁFR"«¬7G&ñÊrÜóFV“ÁFóÚ«¬""í””“'W&vVÁFR#∞¢ñbÜfñ«FW"””“&fGFÚ"í&WGW&‚7G&ñÊrÜóFV“Á7FFÚ«¬""í””“$fGFÚ#∞¢ñbÜfñ«FW"””“&ÊÁV∆∆FÚ"í&WGW&‚7G&ñÊrÜóFV“Á7FFÚ«¬""í””“$ÊÁV∆∆FÚ#∞¢&WGW&‚G'VS∞¢“ì∞¢ñbáVíÁ&ˆw&÷÷¶ñˆÊT∆ó7Bí∞¢VíÁ&ˆw&÷÷¶ñˆÊT∆ó7BÊñÊÊW$ÖD‘¬“fñ«FW&VBÊ÷ÇÜóFV“í”‚∆'Fñ6∆R6∆73“'6ñ◊∆R÷∆ó7B÷óFV“Gµ7G&ñÊrÜóFV“Á7FF˜«¬""ì””“$fGFÚ#Ú'&ˆw&÷÷¶ñˆÊR÷FˆÊR#¢"'“#„«7G&ˆÊs‚G∂W66TÖD‘¬ÜóFV“Ê˜&«¬"“”¢““"ó““G∂W66TÖD‘¬ÜóFV“Ê˜&fñÊW«¬"“”¢““"ó“G∂W66TÖD‘¬ÜóFV“ÁFóÙ∆&V««∆óFV“ÁFó˜«¬""ó”¬˜7G&ˆÊs„«‚G∂W66TÖD‘¬ÜóFV“Ê6ˆ÷÷W76«¬""ó“(
"G∂W66TÖD‘¬ÜóFV“Á¶ˆÊ«¬""ó”¬˜„«‚G∂W66TÖD‘¬ÇÜóFV“ÊÊ˜FW«¬""íÁ6∆ñ6RÉ√Éíó”¬˜„«‚G∂W66TÖD‘¬ÜóFV“Á7FF˜«¬""ó“G∂W66TÖD‘¬á&ˆw&÷÷¶ñˆÊU&V÷ñÊFW$&FvRÜóFV“ÊFF¬óFV“ÁFóÚó«¬""ó”¬˜‚G∂6‰÷ÊvTFFÇìˆ∆Fób6∆73“vóFV“÷7FñˆÁ2s„∆'WGFˆ‚GóS“v'WGFˆ‚r6∆73“v'F‚rFF÷VFóB◊&ˆw&÷÷¶ñˆÊS“rG∂W66TÖD‘¬ÜóFV“ÊñG«¬""ó“s‰÷ˆFñfñ6¬ˆ'WGFˆ„„∆'WGFˆ‚GóS“v'WGFˆ‚r6∆73“v'F‚'F‚÷FÊvW"rFF÷FV∆WFR◊&ˆw&÷÷¶ñˆÊS“rG∂W66TÖD‘¬ÜóFV“ÊñG«¬""ó“s‰V∆ñ÷ñÊ¬ˆ'WGFˆ„„¬ˆFócÊ¢"'”¬ˆ'Fñ6∆SÊíÊ¶ˆñ‚Ç""í«¬#«6∆73“v◊WFVBs‰ÊW77VÊ&ˆw&÷÷¶ñˆÊRfó6ñ&ñ∆R„¬˜‚#∞¢VíÁ&ˆw&÷÷¶ñˆÊT∆ó7BÁVW'ï6V∆V7F˜$∆¬Ç%∂FF÷VFóB◊&ˆw&÷÷¶ñˆÊU“"íÊf˜$V6ÇÇÜ'F‚í”‚'F‚ÊFDWfVÁD∆ó7FVÊW"Ç&6∆ñ6≤"¬Çí”‚˜V‰VFóE&ˆw&÷÷¶ñˆÊRÜ'F‚ÊvWDGG&ñ'WFRÇ&FF÷VFóB◊&ˆw&÷÷¶ñˆÊR"íííì∞¢VíÁ&ˆw&÷÷¶ñˆÊT∆ó7BÁVW'ï6V∆V7F˜$∆¬Ç%∂FF÷FV∆WFR◊&ˆw&÷÷¶ñˆÊU“"íÊf˜$V6ÇÇÜ'F‚í”‚'F‚ÊFDWfVÁD∆ó7FVÊW"Ç&6∆ñ6≤"¬Çí”‚FV∆WFU&ˆw&÷÷¶ñˆÊT'îñBÜ'F‚ÊvWDGG&ñ'WFRÇ&FF÷FV∆WFR◊&ˆw&÷÷¶ñˆÊR"íííì∞¢–¢ñbáVíÁ&ˆw&÷÷¶ñˆÊîÜˆ÷T6&BbbVíÁ&ˆw&÷÷¶ñˆÊîÜˆ÷T∆ó7Bí∞¢6ˆÁ7BÜˆ÷TóFV◊2“fó6ñ&∆RÊfñ«FW"ÇÜóFV“í”‚&ˆˆ∆V‚á&ˆw&÷÷¶ñˆÊU&V÷ñÊFW$&FvRÜóFV“ÊFF¬óFV“ÁFóÚííì∞¢VíÁ&ˆw&÷÷¶ñˆÊîÜˆ÷T6&BÊ6∆74∆ó7BÁFˆvv∆RÇ&ÜñFFV‚"¬Üˆ÷TóFV◊2Ê∆VÊwFÇì∞¢VíÁ&ˆw&÷÷¶ñˆÊîÜˆ÷T6&BÁ6WDGG&ñ'WFRÇ&&ñ÷ÜñFFV‚"¬Üˆ÷TóFV◊2Ê∆VÊwFÇÚ&f«6R"¢'G'VR"ì∞¢VíÁ&ˆw&÷÷¶ñˆÊîÜˆ÷T∆ó7BÊñÊÊW$ÖD‘¬“Üˆ÷TóFV◊2Ê÷ÇÜóFV“í”‚∆'Fñ6∆R6∆73“'6ñ◊∆R÷∆ó7B÷óFV“#„«7G&ˆÊs‚G∂W66TÖD‘¬á&ˆw&÷÷¶ñˆÊU&V÷ñÊFW$&FvRÜóFV“ÊFF¬óFV“ÁFóÚíó”¬˜7G&ˆÊs„«‚G∂W66TÖD‘¬ÜóFV“Ê˜&«¬""ó“(
"G∂W66TÖD‘¬ÜóFV“ÁFóÙ∆&V««∆óFV“ÁFó˜«¬""ó“(
"G∂W66TÖD‘¬ÜóFV“Ê6ˆ÷÷W76«¬""ó”¬˜„¬ˆ'Fñ6∆SÊíÊ¶ˆñ‚Ç""ì∞¢–ß–††¶gVÊ7Fñˆ‚vWDfW&ñTV∆ñvñ&∆T˜W&F˜'2Çí∞¢6ˆÁ7B6ˆ÷÷W76TÊ÷W2“'&íÊg&ˆ“Ü6ˆ÷÷W76T'îñBÁf«VW2ÇííÊ÷ÇÜ2í”‚7G&ñÊrÜ3ÚÊÊˆ÷R«¬""íÁG&ñ“ÇííÊfñ«FW"Ñ&ˆˆ∆V‚ì∞¢&WGW&‚W'6ˆÊ∆U&V6˜&G2Êfñ«FW"ÇáW'6ˆ‚í”‚∞¢6ˆÁ7B∆ƒVÊ&∆VB“&ˆˆ∆V‚áW'6ˆ„ÚÊ&ñ∆óFFıGWGFT6ˆ÷÷W76R«¬W'6ˆ„ÚÊ∆ƒ6ˆ÷÷W76TVÊ&∆VBì∞¢ñbÜ∆ƒVÊ&∆VBí&WGW&‚G'VS∞¢6ˆÁ7BVÊ&∆VB“'&íÊó4'&íáW'6ˆ„ÚÊ6ˆ÷÷W76T&ñ∆óFFRê¢ÚW'6ˆ‚Ê6ˆ÷÷W76T&ñ∆óFFRÊ÷Çábí”‚7G&ñÊráb«¬""íÁG&ñ“ÇííÊfñ«FW"Ñ&ˆˆ∆V‚ê¢¢µ”∞¢ñbÇVÊ&∆VBÊ∆VÊwFÇí&WGW&‚f«6S∞¢ñbÇ6ˆ÷÷W76TÊ÷W2Ê∆VÊwFÇí&WGW&‚G'VS∞¢&WGW&‚6ˆ÷÷W76TÊ÷W2Á6ˆ÷RÇÜ6ˆ÷÷W76Ê÷Rí”‚ó5W'6ˆ‰&ñ∆óFFf˜$6ˆ÷÷W76áW'6ˆ‚¬6ˆ÷÷W76Ê÷Ríì∞¢“ì∞ß–†¶gVÊ7Fñˆ‚&Vg&W6ÑfW&ñU&ˆw&÷÷¶ñˆÊUVíÇí∞¢&Vg&W6ÑfW&ñT˜W&F˜$˜FñˆÁ2Çì∞ß–†¶gVÊ7Fñˆ‚&Vg&W6ÑfW&ñT˜W&F˜$˜FñˆÁ2Çí∞¢ñbÇVíÊfW&ñT˜W&F˜&Rí&WGW&„∞¢6ˆÁ7BV˜∆R“vWDfW&ñTV∆ñvñ&∆T˜W&F˜'2Çì∞¢6ˆÁ7B&Wb“VíÊfW&ñT˜W&F˜&RÁf«VS∞¢VíÊfW&ñT˜W&F˜&RÊñÊÊW$ÖD‘¬“s∆˜Fñˆ‚f«VS“"#‰˜W&F˜&S¬ˆ˜Fñˆ„‚r≤V˜∆P¢Ê÷Çáí”‚vWEW'6ˆÊ∆TFó7∆îÊ÷RáííÊfñ«FW"Ñ&ˆˆ∆V‚íÁ6˜'BÇÜ∆"ì”ÊÊ∆ˆ6∆T6ˆ◊&RÜ"¬vóBríê¢Ê÷ÇÜÊ÷Rì”Ê∆˜Fñˆ‚f«VS“"G∂W66TÖD‘¬ÜÊ÷Ró“#‚G∂W66TÖD‘¬ÜÊ÷Ró”¬ˆ˜Fñˆ„ÊíÊ¶ˆñ‚Ç""ì∞¢ñbá&WbíVíÊfW&ñT˜W&F˜&RÁf«VR“&Wc∞ß–†¶7ñÊ2gVÊ7Fñˆ‚6fTfW&ñT6ˆ∆∆VvÜWfVÁBí∞¢WfVÁBÁ&WfVÁDFVfV«BÇì∞¢ñbÇ6‰÷ÊvTFFÇíí&WGW&„∞¢6ˆÁ7B˜W&F˜&R“7G&ñÊráVíÊfW&ñT˜W&F˜&SÚÁf«VR«¬""íÁG&ñ“Çì∞¢6ˆÁ7BFFñÊó¶ñÚ“7G&ñÊráVíÊfW&ñTñÊó¶ñÛÚÁf«VR«¬""íÁG&ñ“Çì∞¢6ˆÁ7BFFfñÊR“7G&ñÊráVíÊfW&ñTfñÊSÚÁf«VR«¬""íÁG&ñ“Çì∞¢6ˆÁ7BÊ˜FR“7G&ñÊráVíÊfW&ñTÊ˜FSÚÁf«VR«¬""íÁG&ñ“Çì∞¢ñbÇ˜W&F˜&R«¬FFñÊó¶ñÚ«¬FFfñÊRí&WGW&‚∆W'BÇt6ˆ◊ñ∆GWGFíí6◊íˆ&&∆ñvF˜&ífW&ñR‚rì∞¢ñbÜFFfñÊR¬FFñÊó¶ñÚí&WGW&‚∆W'BÇt∆FFfñÊRfW&ñRFWfRW76W&R7V66W76ófÚVwV∆R∆∆FFñÊó¶ñÚ‚rì∞¢vóBF"Ê6ˆ∆∆V7Fñˆ‚ÇvfW&ñT6ˆ∆∆VvÜíríÊFBá≤˜W&F˜&R¬FFñÊó¶ñÚ¬FFfñÊR¬Ê˜FR¬7&VFVDC¢fó&V&6RÊfó&W7F˜&R‰fñV∆Ef«VRÁ6W'fW%Fñ÷W7F◊Çí¬7&VFVD'ì¢7W'&VÁEW6W#ÚÊV÷ñ¬«¬rr“ì∞¢VíÊfW&ñTf˜&”ÚÁ&W6WBÇì∞¢&VÊFW$fW&ñT∆ó7BÇì∞ß–†¶gVÊ7Fñˆ‚6ˆ◊WFTFï7FG2ÜFFT∂Wí¬fW&ñTóFV◊2í∞¢6ˆÁ7BVÊ&∆VEV˜∆R“vWDfW&ñTV∆ñvñ&∆T˜W&F˜'2Çì∞¢6ˆÁ7Bñ‰fW&ñR“ÊWr6WBÜfW&ñTóFV◊2Êfñ«FW"ÇÜbí”‚bÊFFñÊó¶ñÚ√“FFT∂WíbbbÊFFfñÊR„“FFT∂WííÊ÷ÇÜbí”‚Ê˜&÷∆ó¶U6fWGî∂WíÜbÊ˜W&F˜&Rííì∞¢6ˆÁ7Bfñ∆&∆R“VÊ&∆VEV˜∆RÊfñ«FW"Çáí”‚ñ‰fW&ñRÊÜ2ÜÊ˜&÷∆ó¶U6fWGî∂WíÜvWEW'6ˆÊ∆TFó7∆îÊ÷Ráíííì∞¢6ˆÁ7B&W6˜VÁG2“∞¢&ñ÷Û¢fñ∆&∆RÊfñ«FW"Çáí”‚Ü5&WVó&VEW'6ˆÊ∆T6˜W'6Rá¬w&ñ÷Ú6ˆ66˜'6ÚrííÊ∆VÊwFÇ¿¢ÁFñÊ6VÊFñÛ¢fñ∆&∆RÊfñ«FW"Çáí”‚Ü5&WVó&VEW'6ˆÊ∆T6˜W'6Rá¬vÁFñÊ6VÊFñÚrííÊ∆VÊwFÇ¿¢&W˜7FÛ¢fñ∆&∆RÊfñ«FW"Çáí”‚Ü5&WVó&VEW'6ˆÊ∆T6˜W'6Rá¬w&W˜7FÚrííÊ∆VÊwFÄ¢”∞¢6ˆÁ7B'ïV˜∆R“÷FÇÊf∆ˆ˜"Üfñ∆&∆RÊ∆VÊwFÇÚ"ì∞¢6ˆÁ7Bf∆ñEFV◊2“÷FÇÊ÷ÇÉ¬÷FÇÊ÷ñ‚Ü'ïV˜∆R¬&W6˜VÁG2Á&ñ÷Ú¬&W6˜VÁG2ÊÁFñÊ6VÊFñÚ¬&W6˜VÁG2Á&W˜7FÚíì∞¢&WGW&‚≤VÊ&∆VEV˜∆R¬ñ‰fW&ñR¬fñ∆&∆R¬f∆ñEFV◊2”∞ß–†¶gVÊ7Fñˆ‚Ü5&WVó&VEW'6ˆÊ∆T6˜W'6RáW'6ˆ‚¬∂Wóv˜&Bí∞¢6ˆÁ7BF&vWB“Ê˜&÷∆ó¶U6fWGî∂WíÜ∂Wóv˜&Bì∞¢ñbáF&vWBÊñÊ6«VFW2Çw&ñ÷Ú6ˆ66˜'6Úríí&WGW&‚Ü5&WVó&VD6˜W'6RáW'6ˆ‚¬w&ñ÷Ú6ˆ66˜'6Úrì∞¢ñbáF&vWBÊñÊ6«VFW2ÇvÁFñÊ6VÊFñÚríí&WGW&‚Ü5&WVó&VD6˜W'6RáW'6ˆ‚¬vÁFñÊ6VÊFñÚrì∞¢ñbáF&vWBÊñÊ6«VFW2Çw&W˜7FÚríí&WGW&‚Ü5&WVó&VD6˜W'6RáW'6ˆ‚¬w&W˜7FÚrì∞¢ñbáF&vWBÊñÊ6«VFW2ÇvFWÇríí&WGW&‚Ü5&WVó&VD6˜W'6RáW'6ˆ‚¬vFWÇrì∞¢6ˆÁ7B6˜'6í“Ê˜&÷∆ó¶UW'6ˆ‰6˜W'6W2áW'6ˆ‚ì∞¢&WGW&‚ˆ&¶V7BÁf«VW2Ü6˜'6í«¬∑“íÁ6ˆ÷RÇÜ6˜'6Úí”‚&ˆˆ∆V‚Ü6˜'6ÛÚÁ˜76ñVFRíbbÊ˜&÷∆ó¶U6fWGî∂WíÜ6˜'6ÛÚÊÊˆ÷R«¬rríÊñÊ6«VFW2áF&vWBíì∞ß–††¶7ñÊ2gVÊ7Fñˆ‚&VÊFW$fW&ñT∆ó7BÇí∞¢ñbÇVíÊfW&ñT∆ó7Bí&WGW&„∞¢ñbÇ6‰÷ÊvTFFÇíí≤VíÊfW&ñT∆ó7BÊñÊÊW$ÖD‘¬“#«6∆73“v◊WFVBsÂ6ˆ∆ÚF÷ñ‚\;"vW7Fó&RfW&ñR„¬˜‚#≤&WGW&„≤–¢6ˆÁ7B6Ê“vóBF"Ê6ˆ∆∆V7Fñˆ‚ÇvfW&ñT6ˆ∆∆VvÜíríÊ˜&FW$'íÇvFFñÊó¶ñÚr¬v62ríÊvWBÇíÊ6F6ÇÇÇì”ÊÁV∆¬ì∞¢ñbÇ6Êí&WGW&„∞¢6ˆÁ7B&˜w2“6ÊÊFˆ72Ê÷ÇÜBì”‚á∂ñC¶BÊñB¬‚‚ÊBÊFFÇó“íì∞¢VíÊfW&ñT∆ó7BÊñÊÊW$ÖD‘¬“&˜w2Ê÷Çá"ì”Ê∆'Fñ6∆R6∆73“w6ñ◊∆R÷∆ó7B÷óFV“s„«7G&ˆÊs‚G∂W66TÖD‘¬á"Ê˜W&F˜&W«¬r“ró”¬˜7G&ˆÊs„«‚G∂W66TÖD‘¬á"ÊFFñÊó¶ñ˜«¬r“ró“(i"G∂W66TÖD‘¬á"ÊFFfñÊW«¬r“ró”¬˜„«‚G∂W66TÖD‘¬á"ÊÊ˜FW«¬rró”¬˜„∆Fób6∆73“vóFV“÷7FñˆÁ2s„∆'WGFˆ‚GóS“v'WGFˆ‚r6∆73“v'F‚rFF÷VFóB÷fW&ñS“rG∂W66TÖD‘¬á"ÊñBó“s‰÷ˆFñfñ6¬ˆ'WGFˆ„„∆'WGFˆ‚GóS“v'WGFˆ‚r6∆73“v'F‚'F‚÷FÊvW"rFF÷FV¬÷fW&ñS“rG∂W66TÖD‘¬á"ÊñBó“s‰V∆ñ÷ñÊ¬ˆ'WGFˆ„„¬ˆFóc„¬ˆ'Fñ6∆SÊíÊ¶ˆñ‚Çrrí«¬#«6∆73“v◊WFVBs‰ÊW77VÊfW&ñRñÁ6W&óF„¬˜‚#∞¢VíÊfW&ñT∆ó7BÁVW'ï6V∆V7F˜$∆¬Çu∂FF÷FV¬÷fW&ñU“ríÊf˜$V6ÇÇÜ'F‚ì”Ê'F‚ÊFDWfVÁD∆ó7FVÊW"Çv6∆ñ6≤r¬7ñÊ2Çì”Á∞¢ñbÇ6‰÷ÊvTFFÇíí&WGW&„∞¢ñbÇ6ˆÊfó&“ÇtV∆ñ÷ñÊ&RfW&ñSÚríí&WGW&„∞¢vóBF"Ê6ˆ∆∆V7Fñˆ‚ÇvfW&ñT6ˆ∆∆VvÜíríÊFˆ2Ü'F‚ÊvWDGG&ñ'WFRÇvFF÷FV¬÷fW&ñRró«¬rríÊFV∆WFRÇì∞¢&VÊFW$fW&ñT∆ó7BÇì∞¢“íì∞¢VíÊfW&ñT∆ó7BÁVW'ï6V∆V7F˜$∆¬Çu∂FF÷VFóB÷fW&ñU“ríÊf˜$V6ÇÇÜ'F‚ì”Ê'F‚ÊFDWfVÁD∆ó7FVÊW"Çv6∆ñ6≤r¬7ñÊ2Çì”Á∞¢ñbÇ6‰÷ÊvTFFÇíí&WGW&„∞¢6ˆÁ7BñB“'F‚ÊvWDGG&ñ'WFRÇvFF÷VFóB÷fW&ñRrí«¬rs∞¢6ˆÁ7B&˜r“&˜w2ÊfñÊBÇáÇì”ÁÇÊñC””÷ñBì∞¢ñbÇ&˜rí&WGW&„∞¢6ˆÁ7BÊ˜FR“&ˆ◊BÇt÷ˆFñfñ6Ê˜FRfW&ñRr¬&˜rÊÊ˜FR«¬rrì∞¢ñbÜÊ˜FR””“ÁV∆¬í&WGW&„∞¢vóBF"Ê6ˆ∆∆V7Fñˆ‚ÇvfW&ñT6ˆ∆∆VvÜíríÊFˆ2ÜñBíÁ6WBá≤Ê˜FR¬WFFVDC¢fó&V&6RÊfó&W7F˜&R‰fñV∆Ef«VRÁ6W'fW%Fñ÷W7F◊Çí“¬≤÷W&vS¢G'VR“ì∞¢&VÊFW$fW&ñT∆ó7BÇì∞¢“íì∞ß–††¶gVÊ7Fñˆ‚'Vñ∆EFV‘6ˆ÷&ñÊFñˆÁ2Üfñ∆&∆UV˜∆R¬÷ÖFV◊2í∞¢6ˆÁ7B&V÷ñÊñÊr“≤‚‚Êfñ∆&∆UV˜∆U”∞¢6ˆÁ7B6ˆ÷&˜2“µ”∞¢6ˆÁ7BÜ5&W“áW'6ˆ‚¬∂Wíí”‚Ü5&WVó&VEW'6ˆÊ∆T6˜W'6RáW'6ˆ‚¬∂Wíì∞¢f˜"Ü∆WBñGÇ“≤ñGÇ¬÷ÖFV◊3≤ñGÇ≥“í∞¢ñbá&V÷ñÊñÊrÊ∆VÊwFÇ¬"í'&V≥∞¢∆WB∆VDñÊFWÇ“”∞¢∆WB∆VE66˜&R“”∞¢&V÷ñÊñÊrÊf˜$V6ÇÇáW'6ˆ‚¬íí”‚∞¢6ˆÁ7B66˜&R“ÁV÷&W"ÜÜ5&WáW'6ˆ‚¬'&ñ÷Ú6ˆ66˜'6Ú"íí≤ÁV÷&W"ÜÜ5&WáW'6ˆ‚¬&ÁFñÊ6VÊFñÚ"íí≤ÁV÷&W"ÜÜ5&WáW'6ˆ‚¬'&W˜7FÚ"íì∞¢ñbá66˜&R‚∆VE66˜&Rí≤∆VE66˜&R“66˜&S≤∆VDñÊFWÇ“ì≤–¢“ì∞¢ñbÜ∆VDñÊFWÇ¬í'&V≥∞¢6ˆÁ7B∆VB“&V÷ñÊñÊrÁ7∆ñ6RÜ∆VDñÊFWÇ¬ï≥”∞¢∆WB÷FTñÊFWÇ“&V÷ñÊñÊrÊfñÊDñÊFWÇÇáW'6ˆ‚í”‚∞¢6ˆÁ7B&ñ÷Ùˆ≤“Ü5&WÜ∆VB¬'&ñ÷Ú6ˆ66˜'6Ú"í«¬Ü5&WáW'6ˆ‚¬'&ñ÷Ú6ˆ66˜'6Ú"ì∞¢6ˆÁ7BÁFîˆ≤“Ü5&WÜ∆VB¬&ÁFñÊ6VÊFñÚ"í«¬Ü5&WáW'6ˆ‚¬&ÁFñÊ6VÊFñÚ"ì∞¢6ˆÁ7B&Wˆ≤“Ü5&WÜ∆VB¬'&W˜7FÚ"í«¬Ü5&WáW'6ˆ‚¬'&W˜7FÚ"ì∞¢&WGW&‚&ñ÷Ùˆ≤bbÁFîˆ≤bb&Wˆ≥∞¢“ì∞¢ñbÜ÷FTñÊFWÇ¬í÷FTñÊFWÇ“∞¢6ˆÁ7B÷FR“&V÷ñÊñÊrÁ7∆ñ6RÜ÷FTñÊFWÇ¬ï≥”∞¢6ˆ÷&˜2ÁW6ÇÖ∂∆VB¬÷FU“Êfñ«FW"Ñ&ˆˆ∆V‚íì∞¢–¢&WGW&‚6ˆ÷&˜3∞ß–†¶gVÊ7Fñˆ‚f˜&÷EW'6ˆÂ&W&FvW2áW'6ˆ‚í∞¢6ˆÁ7B'G2“µ”∞¢ñbÜÜ5&WVó&VEW'6ˆÊ∆T6˜W'6RáW'6ˆ‚¬'&ñ÷Ú6ˆ66˜'6Ú"íí'G2ÁW6ÇÇ%2"ì∞¢ñbÜÜ5&WVó&VEW'6ˆÊ∆T6˜W'6RáW'6ˆ‚¬&ÁFñÊ6VÊFñÚ"íí'G2ÁW6ÇÇ$í"ì∞¢ñbÜÜ5&WVó&VEW'6ˆÊ∆T6˜W'6RáW'6ˆ‚¬'&W˜7FÚ"íí'G2ÁW6ÇÇ%""ì∞¢&WGW&‚'G2Ê∆VÊwFÇÚ≤G∑'G2Ê¶ˆñ‚Ç"Ú"ó’÷¢"#∞ß–†¶7ñÊ2gVÊ7Fñˆ‚&VÊFW$fW&ñTFó7ˆÊñ&ñ∆óF6∆VÊF"Çí∞¢ñbÇVíÊfW&ñT6∆VÊF%&W7V«Bí&WGW&„∞¢ñbÇ6‰÷ÊvTFFÇíí&WGW&„∞¢6ˆÁ7B7F'B“7G&ñÊráVíÊfW&ñT6ÜV6µ7F'CÚÁf«VR«¬rríÁG&ñ“Çì∞¢6ˆÁ7BVÊB“7G&ñÊráVíÊfW&ñT6ÜV6¥VÊCÚÁf«VR«¬rríÁG&ñ“Çì∞¢ñbÇ7F'B«¬VÊBí&WGW&‚∆W'BÇu6V∆W¶ñˆÊW&ñˆFÚ‚rì∞¢ñbÜVÊB¬7F'Bí&WGW&‚∆W'BÇtñÁFW'f∆∆ÚFFRÊˆ‚f∆ñFÚ‚rì∞†¢6ˆÁ7BfW&ñU6Ê“vóBF"Ê6ˆ∆∆V7Fñˆ‚ÇvfW&ñT6ˆ∆∆VvÜíríÊvWBÇì∞¢6ˆÁ7BfW&ñTóFV◊2“fW&ñU6ÊÊFˆ72Ê÷ÇÜBí”‚á≤ñC¢BÊñB¬‚‚ÊBÊFFÇí“íì∞†¢6ˆÁ7B7F'DFFR“ÊWrFFRÜG∑7F'G’C££ì∞¢6ˆÁ7BVÊDFFR“ÊWrFFRÜG∂VÊG’C££ì∞¢6ˆÁ7BFï7FG2“ÊWr÷Çì∞¢f˜"Ü∆WBB“ÊWrFFRá7F'DFFRì≤B√“VÊDFFS≤BÁ6WDFFRÜBÊvWDFFRÇí≤íí∞¢6ˆÁ7B∂Wí“BÁFÙï4ı7G&ñÊrÇíÁ6∆ñ6RÉ¬ì∞¢6ˆÁ7B7FG2“6ˆ◊WFTFï7FG2Ü∂Wí¬fW&ñTóFV◊2ì∞¢6ˆÁ7B6ˆ÷&˜2“'Vñ∆EFV‘6ˆ÷&ñÊFñˆÁ2á7FG2Êfñ∆&∆R¬7FG2Áf∆ñEFV◊2ì∞¢Fï7FG2Á6WBÜ∂Wí¬≤7FG2¬6ˆ÷&˜2“ì∞¢–†¢6ˆÁ7B÷ˆÁFÖ7F'B“ÊWrFFRá7F'DFFRÊvWDgV∆≈ñV"Çí¬7F'DFFRÊvWD÷ˆÁFÇÇí¬ì∞¢6ˆÁ7B÷ˆÁFÑVÊB“ÊWrFFRÜVÊDFFRÊvWDgV∆≈ñV"Çí¬VÊDFFRÊvWD÷ˆÁFÇÇí¬ì∞¢6ˆÁ7B÷ˆÁFÑ&∆ˆ6∑2“µ”∞†¢f˜"Ü∆WB““ÊWrFFRÜ÷ˆÁFÖ7F'Bì≤“√“÷ˆÁFÑVÊC≤“Á6WD÷ˆÁFÇÜ“ÊvWD÷ˆÁFÇÇí≤íí∞¢6ˆÁ7BñV"““ÊvWDgV∆≈ñV"Çì∞¢6ˆÁ7B÷ˆÁFÇ““ÊvWD÷ˆÁFÇÇì∞¢6ˆÁ7Bfó'7DFí“ÊWrFFRáñV"¬÷ˆÁFÇ¬ì∞¢6ˆÁ7BFó4ñ‰÷ˆÁFÇ“ÊWrFFRáñV"¬÷ˆÁFÇ≤¬íÊvWDFFRÇì∞¢6ˆÁ7B7F'Dˆfg6WB“Üfó'7DFíÊvWDFíÇí≤bíRs∞¢6ˆÁ7B÷ˆÁFÑ∆&V¬“fó'7DFíÁFÙ∆ˆ6∆TFFU7G&ñÊrÇvóB‘ïBr¬≤÷ˆÁFÉ¢v∆ˆÊrr¬ñV#¢vÁV÷W&ñ2r“ì∞¢6ˆÁ7B6V∆«2“µ”∞†¢f˜"Ü∆WBí“≤í¬7F'Dˆfg6WC≤í≥“í6V∆«2ÁW6ÇÇs∆Fób6∆73“&fW&ñR÷÷ˆÁFÇ÷6V∆¬fW&ñR÷÷ˆÁFÇ÷6V∆¬“÷V◊Gí#„¬ˆFóc‚rì∞†¢f˜"Ü∆WBFí“≤Fí√“Fó4ñ‰÷ˆÁFÉ≤Fí≥“í∞¢6ˆÁ7BFîFFR“ÊWrFFRáñV"¬÷ˆÁFÇ¬Fíì∞¢6ˆÁ7B∂Wí“FîFFRÁFÙï4ı7G&ñÊrÇíÁ6∆ñ6RÉ¬ì∞¢6ˆÁ7BñÂ&ÊvR“FîFFR„“7F'DFFRbbFîFFR√“VÊDFFS∞¢6ˆÁ7Bñ∆ˆB“ñÂ&ÊvRÚFï7FG2ÊvWBÜ∂Wíí¢ÁV∆√∞¢6ˆÁ7Bf∆ñEFV◊2“ñ∆ˆCÚÁ7FG3ÚÁf∆ñEFV◊2«¬∞¢6ˆÁ7B7FGW46∆72“ñÂ&ÊvRÚvfW&ñR÷÷ˆÁFÇ÷6V∆¬“÷˜WBr¢áf∆ñEFV◊2‚ÚvfW&ñR÷÷ˆÁFÇ÷6V∆¬“÷ˆ≤r¢vfW&ñR÷÷ˆÁFÇ÷6V∆¬“÷∂Úrì∞¢6ˆÁ7BFWFñƒñB“fW&ñR÷Fí÷FWFñ¬“G∂∂Wó÷∞¢6ˆÁ7B7FG2“ñ∆ˆCÚÁ7FG3∞¢6ˆÁ7B6ˆ÷&˜2“ñ∆ˆCÚÊ6ˆ÷&˜2«¬µ”∞¢6ˆÁ7B6ˆ÷&ı&˜w2“6ˆ÷&˜2Ê÷ÇáFV“¬ñGÇí”‚∆∆ì„∆#Â7VG&G∂ñGÇ≤”¬ˆ#„¢G∑FV“Ê÷Çáí”‚G∂W66TÖD‘¬ÜvWEW'6ˆÊ∆TFó7∆îÊ÷Ráí«¬r“ró“G∂W66TÖD‘¬Üf˜&÷EW'6ˆÂ&W&FvW2áíó÷íÊ¶ˆñ‚Çr≤ró”¬ˆ∆ìÊíÊ¶ˆñ‚Çrrì∞¢6ˆÁ7Bñ‰fW&ñTÊ÷W2“7FG0¢Ú7FG2ÊVÊ&∆VEV˜∆RÊfñ«FW"Çáí”‚7FG2Êñ‰fW&ñRÊÜ2ÜÊ˜&÷∆ó¶U6fWGî∂WíÜvWEW'6ˆÊ∆TFó7∆îÊ÷RáííííÊ÷Çáí”‚vWEW'6ˆÊ∆TFó7∆îÊ÷RáííÊfñ«FW"Ñ&ˆˆ∆V‚ê¢¢µ”∞¢6ˆÁ7BFWFñ¬“7FG2Ú∆FóbñC“"G∂W66TÖD‘¬ÜFWFñƒñBó“"6∆73“&fW&ñR÷Fí÷FWFñ¬ÜñFFV‚#„«„∆#‰FF£¬ˆ#‚G∂W66TÖD‘¬Ü∂Wíó”¬˜„«‰&ñ∆óFFì¢G∑7FG2ÊVÊ&∆VEV˜∆RÊ∆VÊwFá“(
"ñ‚fW&ñS¢G∑7FG2Êñ‰fW&ñRÁ6ó¶W“(
"Fó7ˆÊñ&ñ∆ì¢G∑7FG2Êfñ∆&∆RÊ∆VÊwFá”¬˜„«Ó)»R7VG&R6ˆ◊∆WFR7&V&ñ∆ì¢G∑f∆ñEFV◊7”¬˜„«‚G∑f∆ñEFV◊2””“bb7FG2Êfñ∆&∆RÊ∆VÊwFÇ‚Ú)™˚àÚW'6ˆÊRFó7ˆÊñ&ñ∆í÷&WVó6óFí÷Ê6ÁFì¢G∑7FG2Êfñ∆&∆RÊ∆VÊwFá÷¢~)™˚àÚW'6ˆÊRFó7ˆÊñ&ñ∆í÷&WVó6óFí÷Ê6ÁFì¢w”¬˜„«‚G∑f∆ñEFV◊2””“Ú~)ÿ¬vñ˜&ÊÚ66˜W'FÚr¢rw”¬˜„«„∆#‰6ˆ∆∆VvÜíñ‚fW&ñS£¬ˆ#‚G∂W66TÖD‘¬Üñ‰fW&ñTÊ÷W2Ê¶ˆñ‚Çr¬rí«¬r“ró”¬˜‚G∂6ˆ÷&ı&˜w2Ú«V√‚G∂6ˆ÷&ı&˜w7”¬˜V√Ê¢s«6∆73“&◊WFVB#‰ÊW77VÊ6ˆ÷&ñÊ¶ñˆÊRf∆ñF„¬˜‚w”¬ˆFócÊ¢rs∞¢6V∆«2ÁW6ÇÜ∆'WGFˆ‚GóS“&'WGFˆ‚"6∆73“&fW&ñR÷÷ˆÁFÇ÷6V∆¬G∑7FGW46∆77“"G∂ñÂ&ÊvRÚFF÷fW&ñR◊Fˆvv∆S“"G∂W66TÖD‘¬ÜFWFñƒñBó“&¢vFó6&∆VBw”„«7‚6∆73“&fW&ñR÷÷ˆÁFÇ÷FñÁV“#‚G∂Fó”¬˜7„‚G∂ñÂ&ÊvRÚ«6÷∆√Â7¢G∑f∆ñEFV◊7”¬˜6÷∆√Ê¢rw”¬ˆ'WGFˆ„‚G∂FWFñ«÷ì∞¢–†¢÷ˆÁFÑ&∆ˆ6∑2ÁW6ÇÜ«6V7Fñˆ‚6∆73“&fW&ñR÷÷ˆÁFÇ#„∆ÉS‚G∂W66TÖD‘¬Ü÷ˆÁFÑ∆&V¬Ê6Ü$BÉíÁFıWW$66RÇí≤÷ˆÁFÑ∆&V¬Á6∆ñ6RÉíó”¬ˆÉS„∆Fób6∆73“&fW&ñR÷÷ˆÁFÇ◊vVV∂Fó2#„«7„‰«V„¬˜7„„«7„‰÷#¬˜7„„«7„‰÷W#¬˜7„„«7„‰vñÛ¬˜7„„«7„ÂfV„¬˜7„„«7„Â6#¬˜7„„«7„‰Fˆ”¬˜7„„¬ˆFóc„∆Fób6∆73“&fW&ñR÷÷ˆÁFÇ÷w&ñB#‚G∂6V∆«2Ê¶ˆñ‚Çrró”¬ˆFóc„¬˜6V7Fñˆ„Êì∞¢–†¢VíÊfW&ñT6∆VÊF%&W7V«BÊñÊÊW$ÖD‘¬“∆Fób6∆73“&fW&ñR÷÷ˆÁFÇ◊w&#‚G∂÷ˆÁFÑ&∆ˆ6∑2Ê¶ˆñ‚Çrró”¬ˆFócÊ∞¢VíÊfW&ñT6∆VÊF%&W7V«BÁVW'ï6V∆V7F˜$∆¬Çu∂FF÷fW&ñR◊Fˆvv∆U“ríÊf˜$V6ÇÇÜ'F‚í”‚∞¢'F‚ÊFDWfVÁD∆ó7FVÊW"Çv6∆ñ6≤r¬Çí”‚∞¢6ˆÁ7BñB“'F‚ÊvWDGG&ñ'WFRÇvFF÷fW&ñR◊Fˆvv∆Rrí«¬rs∞¢6ˆÁ7BFWFñ¬“VíÊfW&ñT6∆VÊF%&W7V«BÁVW'ï6V∆V7F˜"Ü2G∂774W66Uf«VRÜñBó÷ì∞¢ñbÇFWFñ¬í&WGW&„∞¢FWFñ¬Ê6∆74∆ó7BÁFˆvv∆RÇvÜñFFV‚rì∞¢“ì∞¢“ì∞ß–††¶gVÊ7Fñˆ‚˜V‰VFóE&ˆw&÷÷¶ñˆÊRÜñBí∞¢ñbÇ6‰÷ÊvTFFÇíí&WGW&„∞¢6ˆÁ7BóFV““&ˆw&÷÷¶ñˆÊíÊfñÊBÇá&˜rí”‚&˜rÊñB””“ñBì∞¢ñbÇóFV“í&WGW&„∞¢˜V∆FU&ˆw&÷÷¶ñˆÊTf˜&‘˜FñˆÁ2Çì∞¢VíÁ&ˆw&÷÷ñBÁf«VR“óFV“ÊñB«¬"#∞¢Fˆ7V÷VÁBÊvWDV∆V÷VÁD'îñBÇ'&ˆw&÷÷÷FF"íÁf«VR“óFV“ÊFF«¬"#∞¢Fˆ7V÷VÁBÊvWDV∆V÷VÁD'îñBÇ'&ˆw&÷÷÷˜&"íÁf«VR“óFV“Ê˜&«¬"#∞¢Fˆ7V÷VÁBÊvWDV∆V÷VÁD'îñBÇ'&ˆw&÷÷÷˜&÷fñÊR"íÁf«VR“óFV“Ê˜&fñÊR«¬"#∞¢Fˆ7V÷VÁBÊvWDV∆V÷VÁD'îñBÇ'&ˆw&÷÷◊¶ˆÊ"íÁf«VR“óFV“Á¶ˆÊ«¬"#∞¢Fˆ7V÷VÁBÊvWDV∆V÷VÁD'îñBÇ'&ˆw&÷÷◊FóÚ"íÁf«VR“óFV“ÁFóÚ«¬'6f∆6ñÚ#∞¢Fˆ7V÷VÁBÊvWDV∆V÷VÁD'îñBÇ'&ˆw&÷÷◊&ñ˜&óF"íÁf«VR“óFV“Á&ñ˜&óF«¬&Ê˜&÷∆R#∞¢Fˆ7V÷VÁBÊvWDV∆V÷VÁD'îñBÇ'&ˆw&÷÷÷Ê˜FR"íÁf«VR“óFV“ÊÊ˜FR«¬"#∞¢Fˆ7V÷VÁBÊvWDV∆V÷VÁD'îñBÇ'&ˆw&÷÷÷6ˆ÷÷W76"íÁf«VR“óFV“Ê6ˆ÷÷W76«¬"#∞¢&ˆw&÷÷¶ñˆÊT˜W&F˜$WFˆ6ˆ◊∆WFR“'Vñ∆E&ˆw&÷÷¶ñˆÊTWFˆ6ˆ◊∆WFRáVíÁ&ˆw&÷÷˜W&F˜&îWFˆ6ˆ◊∆WFR¬$˜W&F˜&í6ˆñÁfˆ«Fí"¬vWE&ˆw&÷÷¶ñˆÊT˜W&F˜$˜FñˆÁ2Çí¬óFV“Ê˜W&F˜&î6ˆñÁfˆ«Fí«¬µ“ì∞¢&ˆw&÷÷¶ñˆÊT÷Wß¶îWFˆ6ˆ◊∆WFR“'Vñ∆E&ˆw&÷÷¶ñˆÊTWFˆ6ˆ◊∆WFRáVíÁ&ˆw&÷÷÷Wß¶îWFˆ6ˆ◊∆WFR¬$÷Wß¶íÚGG&Wß¶GW&R"¬vWE&ˆw&÷÷¶ñˆÊT÷Wß¶î˜FñˆÁ2Çí¬óFV“Ê÷Wß¶î76VvÊFí«¬µ“ì∞¢VíÁ&ˆw&÷÷¶ñˆÊTFV∆WFT'F„ÚÊ6∆74∆ó7BÁ&V÷˜fRÇ&ÜñFFV‚"ì∞¢VíÁ&ˆw&÷÷¶ñˆÊTFñ∆ˆsÚÁ6Ü˜t÷ˆF¬Çì∞ß–†¶7ñÊ2gVÊ7Fñˆ‚6fU&ˆw&÷÷¶ñˆÊRÜWfVÁBí∞¢WfVÁBÁ&WfVÁDFVfV«BÇì∞¢ñbÇ6‰÷ÊvTFFÇíí&WGW&„∞¢6ˆÁ7Bñ∆ˆB“∞¢WFFVDC¢fó&V&6RÊfó&W7F˜&R‰fñV∆Ef«VRÁ6W'fW%Fñ÷W7F◊Çí¿¢FF¢Fˆ7V÷VÁBÊvWDV∆V÷VÁD'îñBÇ'&ˆw&÷÷÷FF"ìÚÁf«VR«¬""¿¢˜&¢Fˆ7V÷VÁBÊvWDV∆V÷VÁD'îñBÇ'&ˆw&÷÷÷˜&"ìÚÁf«VR«¬""¿¢˜&fñÊS¢Fˆ7V÷VÁBÊvWDV∆V÷VÁD'îñBÇ'&ˆw&÷÷÷˜&÷fñÊR"ìÚÁf«VR«¬""¿¢FóÛ¢Fˆ7V÷VÁBÊvWDV∆V÷VÁD'îñBÇ'&ˆw&÷÷◊FóÚ"ìÚÁf«VR«¬""¿¢FóÙ∆&V√¢Fˆ7V÷VÁBÊvWDV∆V÷VÁD'îñBÇ'&ˆw&÷÷◊FóÚ"ìÚÁ6V∆V7FVD˜FñˆÁ3ÚÂ≥”ÚÁFWáD6ˆÁFVÁB«¬""¿¢6ˆ÷÷W76¢Fˆ7V÷VÁBÊvWDV∆V÷VÁD'îñBÇ'&ˆw&÷÷÷6ˆ÷÷W76"ìÚÁf«VR«¬""¿¢¶ˆÊ¢Fˆ7V÷VÁBÊvWDV∆V÷VÁD'îñBÇ'&ˆw&÷÷◊¶ˆÊ"ìÚÁf«VR«¬""¿¢˜W&F˜&î6ˆñÁfˆ«Fì¢&ˆw&÷÷¶ñˆÊT˜W&F˜$WFˆ6ˆ◊∆WFSÚÊvWEf«VW3Ú‚Çí«¬µ“¿¢&ñ˜&óF¢Fˆ7V÷VÁBÊvWDV∆V÷VÁD'îñBÇ'&ˆw&÷÷◊&ñ˜&óF"ìÚÁf«VR«¬&Ê˜&÷∆R"¿¢Ê˜FS¢Fˆ7V÷VÁBÊvWDV∆V÷VÁD'îñBÇ'&ˆw&÷÷÷Ê˜FR"ìÚÁf«VR«¬""¿¢7FFÛ¢%&ˆw&÷÷FÚ"¿¢÷Wß¶î76VvÊFì¢&ˆw&÷÷¶ñˆÊT÷Wß¶îWFˆ6ˆ◊∆WFSÚÊvWEf«VW3Ú‚Çí«¬µ“¿¢7&VFVD'ì¢7W'&VÁEW6W#ÚÊV÷ñ¬«¬" ¢”∞¢6ˆÁ7BñB“7G&ñÊráVíÁ&ˆw&÷÷ñCÚÁf«VR«¬""íÁG&ñ“Çì∞¢ñbÜñBí∞¢vóBF"Ê6ˆ∆∆V7Fñˆ‚Ç'&ˆw&÷÷¶ñˆÊí"íÊFˆ2ÜñBíÁ6WBáñ∆ˆB¬≤÷W&vS¢G'VR“ì∞¢∆W'BÇ%&ˆw&÷÷¶ñˆÊRvvñ˜&ÊF6˜'&WGF÷VÁFR"ì∞¢vóB7ñÊ5&ˆw&÷÷¶ñˆÊUFı7VG&ÜñB¬ñ∆ˆB¬≤&V÷˜fS¢f«6R“ì∞¢“V«6R∞¢ñ∆ˆBÊ7&VFVDB“fó&V&6RÊfó&W7F˜&R‰fñV∆Ef«VRÁ6W'fW%Fñ÷W7F◊Çì∞¢6ˆÁ7BFˆ5&Vb“vóBF"Ê6ˆ∆∆V7Fñˆ‚Ç'&ˆw&÷÷¶ñˆÊí"íÊFBáñ∆ˆBì∞¢vóB7ñÊ5&ˆw&÷÷¶ñˆÊUFı7VG&ÜFˆ5&VbÊñB¬ñ∆ˆB¬≤&V÷˜fS¢f«6R“ì∞¢–¢VíÁ&ˆw&÷÷¶ñˆÊTFñ∆ˆsÚÊ6∆˜6RÇì∞¢VíÁ&ˆw&÷÷¶ñˆÊTf˜&”ÚÁ&W6WBÇì∞¢vóB7V'67&ñ&U&ˆw&÷÷¶ñˆÊíá≤f˜&6S¢G'VR“ì∞ß–†¶7ñÊ2gVÊ7Fñˆ‚7ñÊ5&ˆw&÷÷¶ñˆÊUFı7VG&á&ˆw&÷÷¶ñˆÊTñB¬ñ∆ˆB¬≤&V÷˜fR“f«6R““∑“í∞¢6ˆÁ7B6ˆ÷÷W76“'&íÊg&ˆ“Ü6ˆ÷÷W76T'îñBÁf«VW2ÇííÊfñÊBÇá&˜rí”‚7G&ñÊrá&˜rÊÊˆ÷R«¬""íÁG&ñ“Çí””“7G&ñÊráñ∆ˆBÊ6ˆ÷÷W76«¬""íÁG&ñ“Çíì∞¢ñbÇ6ˆ÷÷W76ÚÊñB«¬ñ∆ˆCÚÊFFí&WGW&„∞¢6ˆÁ7BÜó7F˜'ï&Vb“F"Ê6ˆ∆∆V7Fñˆ‚Ç'7VG&U7F˜&ñ6Ú"íÊFˆ2ÜG∑ñ∆ˆBÊFF’ıÚG∂6ˆ÷÷W76ÊñG÷ì∞¢ñbá&V÷˜fRí∞¢vóBÜó7F˜'ï&VbÁ6WBá≤WFˆvVÊW&FVDg&ˆ’&ˆw&÷÷¶ñˆÊS¢fó&V&6RÊfó&W7F˜&R‰fñV∆Ef«VRÊFV∆WFRÇí¬WFı&ˆw&÷÷¶ñˆÊTñC¢fó&V&6RÊfó&W7F˜&R‰fñV∆Ef«VRÊFV∆WFRÇí“¬≤÷W&vS¢G'VR“ì∞¢&WGW&„∞¢–¢vóBÜó7F˜'ï&VbÁ6WBá∞¢6ˆ÷÷W76ñC¢6ˆ÷÷W76ÊñB¿¢6ˆ÷÷W76Êˆ÷S¢6ˆ÷÷W76ÊÊˆ÷R«¬$6ˆ÷÷W76"¿¢&ñfW&ñ÷VÁFÙFF¢ñ∆ˆBÊFF¿¢FFT∂Wì¢ñ∆ˆBÊFF¿¢WFı&ˆw&÷÷¶ñˆÊTñC¢&ˆw&÷÷¶ñˆÊTñB¿¢WFˆvVÊW&FVDg&ˆ’&ˆw&÷÷¶ñˆÊS¢G'VR¿¢7VG&S¢∑≤W'6ˆÊ∆S¢áñ∆ˆBÊ˜W&F˜&î6ˆñÁfˆ«Fí«¬µ“íÊ¶ˆñ‚Ç"¬"í¬÷Wß¶ì¢áñ∆ˆBÊ÷Wß¶î76VvÊFí«¬µ“íÊ¶ˆñ‚Ç"¬"í¬ñ◊ñÁFì¢ñ∆ˆBÁ¶ˆÊ«¬""¬Ê˜FS¢ñ∆ˆBÊÊ˜FR«¬""¬˜&&ñÛ¢G∑ñ∆ˆBÊ˜&«¬"'““G∑ñ∆ˆBÊ˜&fñÊR«¬"'÷’–¢“¬≤÷W&vS¢G'VR“ì∞ß–†¶7ñÊ2gVÊ7Fñˆ‚FV∆WFU&ˆw&÷÷¶ñˆÊT'îñBÜñBí∞¢ñbÇ6‰÷ÊvTFFÇí«¬ñBí&WGW&„∞¢6ˆÁ7BóFV““&ˆw&÷÷¶ñˆÊíÊfñÊBÇá&˜rí”‚&˜rÊñB””“ñBì∞¢ñbÇóFV“í&WGW&„∞¢ñbÇvñÊF˜rÊ6ˆÊfó&“Ç%6Ví6ñ7W&ÚFífˆ∆W"V∆ñ÷ñÊ&RVW7F&ˆw&÷÷¶ñˆÊSÚ"íí&WGW&„∞¢vóB7ñÊ5&ˆw&÷÷¶ñˆÊUFı7VG&ÜñB¬óFV“¬≤&V÷˜fS¢G'VR“ì∞¢vóBF"Ê6ˆ∆∆V7Fñˆ‚Ç'&ˆw&÷÷¶ñˆÊí"íÊFˆ2ÜñBíÊFV∆WFRÇì∞¢vóB7V'67&ñ&U&ˆw&÷÷¶ñˆÊíá≤f˜&6S¢G'VR“ì∞ß–¶7ñÊ2gVÊ7Fñˆ‚FV∆WFU&ˆw&÷÷¶ñˆÊTg&ˆ‘f˜&“Çí≤&WGW&‚FV∆WFU&ˆw&÷÷¶ñˆÊT'îñBáVíÁ&ˆw&÷÷ñCÚÁf«VR«¬""ì≤–††¶Fˆ7V÷VÁBÊFDWfVÁD∆ó7FVÊW"Ç&6∆ñ6≤"¬ÜWfVÁBí”‚∞¢ñbÜWfVÁBÁF&vWCÚÊñB”“&&ñˆv2÷6VÁFW"÷÷÷'F‚"í&WGW&„∞¢ñbÇ&ñˆv4÷ñÁ7FÊ6R«¬&ñˆv5W6W$÷&∂W"í&WGW&„∞¢&ñˆv4÷ñÁ7FÊ6RÁ6WEfñWrÜ&ñˆv5W6W$÷&∂W"ÊvWD∆D∆ÊrÇí¬÷FÇÊ÷ÇÜ&ñˆv4÷ñÁ7FÊ6RÊvWE¶ˆˆ“Çí¬Çíì∞ß“ì∞†¶6ˆÁ7B4‰ıuı4U%dî4UÙ4ÙƒƒT5DîÙÂ2“∞¢6∆ñVÁG3¢'6W'fó¶ñÙÊWfT6∆ñVÁFí"¿¢&˜WFW3¢'6W'fó¶ñÙÊWfUW&6˜'6í"¿¢fVÜñ6∆W3¢'6W'fó¶ñÙÊWfT÷Wß¶í"¿¢˜W&F˜'3¢'6W'fó¶ñÙÊWfT˜W&F˜&í"¿¢&W˜'G3¢'6W'fó¶ñÙÊWfU6VvÊ∆¶ñˆÊí"¿¢6WGFñÊw3¢'6W'fó¶ñÙÊWfTñ◊˜7F¶ñˆÊí ß”∞¶6ˆÁ7B6Ê˜u6W'fñ6U7FFR“≤6∆ñVÁG3¢µ“¬&˜WFW3¢µ“¬fVÜñ6∆W3¢µ“¬˜W&F˜'3¢µ“¬&W˜'G3¢µ“”∞¶6ˆÁ7B4‰ıuı4U%dî4UÙ4Ù‘‘U54R“∞¢≤Êˆ÷S¢$6ˆ◊VÊRFífW'&&"¬7VG&¢%7VG&ÊWfR"¬˜W&F˜&ì¢$6˜7VG&¬WFó7F∆÷¬7W˜'FÚ6∆R"¬÷Wß¶ì¢$∆÷¬7&vó6∆R"“¿¢≤Êˆ÷S¢$6ˆ◊VÊRFífñv&ÊÚ÷ñÊ&F"¬7VG&¢%7VG&ÊWfR""¬˜W&F˜&ì¢$6˜7VG&¬WFó7FG&GF˜&R¬7W˜'FÚfñ&ñ∆óL:"¬÷Wß¶ì¢%G&GF˜&RÊWfR¬7&vó6∆R""“¿¢≤Êˆ÷S¢$6ˆ◊VÊRFí&ˆÊFVÊÚ"¬7VG&¢%7VG&ÊWfR2"¬˜W&F˜&ì¢$6˜7VG&¬WFó7F∆¬˜W&F˜&R÷ÁV∆R"¬÷Wß¶ì¢%∆vˆ÷÷F¬ñ6≤◊WÊWfR"“¿¢≤Êˆ÷S¢$6ˆ◊VÊRFí6VÁFÚ"¬7VG&¢%7VG&ÊWfRB"¬˜W&F˜&ì¢$6˜7VG&¬WFó7F∆÷¬7W˜'FÚV÷W&vVÁ¶R"¬÷Wß¶ì¢$∆÷"¬7&vó6∆R2"“¿¢≤Êˆ÷S¢$6ˆ◊VÊRFí6˜&Ú"¬7VG&¢%7VG&ÊWfRR"¬˜W&F˜&ì¢$6˜7VG&¬WFó7FG&GF˜&R¬˜W&F˜&R6∆R"¬÷Wß¶ì¢%G&GF˜&R∆÷¬G&÷ˆvvñ6∆R"“¿¢≤Êˆ÷S¢$6ˆ◊VÊRFí&vVÁF"¬7VG&¢%7VG&ÊWfRb"¬˜W&F˜&ì¢$6˜7VG&¬WFó7F÷Wß¶ÚGÉB¬7W˜'FÚ&WW&ñ&ñ∆óL:"¬÷Wß¶ì¢#GÉBÊWfR¬7&vó6∆R6ˆ◊GFÚ"“¿¢≤Êˆ÷S¢$6ˆ◊VÊRFí˜'Fˆ÷vvñ˜&R"¬7VG&¢%7VG&ÊWfRr"¬˜W&F˜&ì¢$6˜7VG&¬WFó7F∆÷¬7W˜'FÚ7G&FR"¬÷Wß¶ì¢$∆÷2¬ñ6≤◊W6∆R"“¿¢≤Êˆ÷S¢$6ˆ◊VÊRFí6ˆ÷66ÜñÚ"¬7VG&¢%7VG&ÊWfRÇ"¬˜W&F˜&ì¢$6˜7VG&¬WFó7F∆¬7W˜'FÚ∆óF˜&∆R"¬÷Wß¶ì¢%∆6ˆ◊GF¬7&vó6∆RB"“¿¢≤Êˆ÷S¢$6ˆ◊VÊRFíˆ66Üñˆ&V∆∆Ú"¬7VG&¢%7VG&ÊWfRí"¬˜W&F˜&ì¢$6˜7VG&¬WFó7FG&GF˜&R¬7W˜'FÚˆÁFí"¬÷Wß¶ì¢%G&GF˜&RÊWfR¬∆÷∆FW&∆R"–•”∞¶∆WB6Ê˜u6W'fñ6UVÁ7V'67&ñ&W'2“µ”∞¶∆WB6Ê˜t÷ÁV≈7VG&Tfñ«FW$FFT∂Wí“"#∞¶∆WB6Ê˜u6Ü&VE7VG&TFFT∂Wí“"#∞†¶gVÊ7Fñˆ‚ó56Ê˜u6W'fñ6T6ˆÁFWáBÇí∞¢&WGW&‚vñÊF˜rÊ∆ˆ6Fñˆ‚ÊÜ6Ç””“"76W'fó¶ñÚ÷ÊWfR"«¬Fˆ7V÷VÁBÊ&ˆGíÊ6∆74∆ó7BÊ6ˆÁFñÁ2Ç'6Ê˜r÷÷ÊvV÷VÁB÷6ˆÁFWáB"ì∞ß–†¶gVÊ7Fñˆ‚vWD6ˆ÷÷W76T6ˆ∆∆V7Fñˆ‰Ê÷RÇí≤&WGW&‚ó56Ê˜u6W'fñ6T6ˆÁFWáBÇíÚ&ÊWfUˆ6ˆ÷÷W76R"¢&6ˆ÷÷W76R#≤–¶gVÊ7Fñˆ‚vWE7VG&T7W'&VÁD6ˆ∆∆V7Fñˆ‰Ê÷RÇí≤&WGW&‚ó56Ê˜u6W'fñ6T6ˆÁFWáBÇíÚ&ÊWfU˜7VG&R"¢'7VG&T6ˆ÷÷W76R#≤–¶gVÊ7Fñˆ‚vWE7VG&TÜó7F˜'î6ˆ∆∆V7Fñˆ‰Ê÷RÇí≤&WGW&‚ó56Ê˜u6W'fñ6T6ˆÁFWáBÇíÚ&ÊWfU˜7VG&U˜7F˜&ñ6Ú"¢'7VG&U7F˜&ñ6Ú#≤–¶gVÊ7Fñˆ‚vWD˜&U&W˜'G46ˆ∆∆V7Fñˆ‰Ê÷RÇí≤&WGW&‚ó56Ê˜u6W'fñ6T6ˆÁFWáBÇíÚ&ÊWfUˆ˜&R"¢&˜&U&W˜'G2#≤–¶gVÊ7Fñˆ‚vWD˜&T&˜f≈&WVW7G46ˆ∆∆V7Fñˆ‰Ê÷RÇí≤&WGW&‚ó56Ê˜u6W'fñ6T6ˆÁFWáBÇíÚ&ÊWfUˆ˜&U˜&ñ6ÜñW7FR"¢&˜&T&˜f≈&WVW7G2#≤–¶gVÊ7Fñˆ‚vWEW'6ˆÊ∆T6ˆ∆∆V7Fñˆ‰Ê÷RÇí≤&WGW&‚ó56Ê˜u6W'fñ6T6ˆÁFWáBÇíÚ&ÊWfU˜W'6ˆÊ∆R"¢'W'6ˆÊ∆R#≤–¶gVÊ7Fñˆ‚vWD÷Wß¶î6ˆ∆∆V7Fñˆ‰Ê÷RÇí≤&WGW&‚ó56Ê˜u6W'fñ6T6ˆÁFWáBÇíÚ&ÊWfUˆ÷Wß¶í"¢&÷Wß¶í#≤–†¶gVÊ7Fñˆ‚˜VÂ6Ê˜u6W'fñ6UvRÇí∞¢ñbÇ6‰÷ÊvTFFÇíí∞¢∆W'BÇ%6ˆ∆Ú¬vF÷ñ‚\;"66VFW&R¬6W'fó¶ñÚÊWfR‚"ì∞¢&WGW&„∞¢–¢Fˆ7V÷VÁBÊ&ˆGíÊ6∆74∆ó7BÊFBÇ'6Ê˜r÷÷ÊvV÷VÁB÷6ˆÁFWáB"ì∞¢vñÊF˜rÊ∆ˆ6Fñˆ‚ÊÜ6Ç“'6W'fó¶ñÚ÷ÊWfR#∞¢7F˜Ê˜&÷ƒFF7V'67&óFñˆÁ4f˜%6Ê˜t÷ˆFRÇì∞¢∆ˆE6Ê˜t÷ˆFTFFÇì∞¢«ï&˜WFRÇì∞ß–†¶gVÊ7Fñˆ‚6∆˜6U6Ê˜u6W'fñ6UvRÇí∞¢7F˜6Ê˜t÷ˆFTFFÇì∞¢Fˆ7V÷VÁBÊ&ˆGíÊ6∆74∆ó7BÁ&V÷˜fRÇ'6Ê˜r÷÷ÊvV÷VÁB÷6ˆÁFWáB"ì∞¢6ˆÊfñwW&U6Ê˜u6ñFT÷VÁRÜf«6Rì∞¢vñÊF˜rÊ∆ˆ6Fñˆ‚ÊÜ6Ç“"#∞¢&V∆ˆDÊ˜&÷ƒ÷ˆFTFFÇì∞¢«ï&˜WFRÇì∞ß–†¶gVÊ7Fñˆ‚ó56Ê˜u6W'fñ6U&˜WFRÇí∞¢&WGW&‚vñÊF˜rÊ∆ˆ6Fñˆ‚ÊÜ6Ç””“"76W'fó¶ñÚ÷ÊWfR#∞ß–†¶gVÊ7Fñˆ‚&VÊFW%6Ê˜u6W'fñ6T∆ó7BÜV∆V÷VÁB¬&˜w2¬V◊GïFWáB¬&VÊFW%&˜rí∞¢ñbÇV∆V÷VÁBí&WGW&„∞¢V∆V÷VÁBÊñÊÊW$ÖD‘¬“&˜w2Ê∆VÊwFÄ¢Ú&˜w2Ê÷á&VÊFW%&˜ríÊ¶ˆñ‚Ç""ê¢¢«6∆73“v◊WFVBs‚G∂W66TÖD‘¬ÜV◊GïFWáBó”¬˜Ê∞ß–†¶gVÊ7Fñˆ‚&VÊFW%6Ê˜u6W'fñ6T6ˆ÷÷W76RÇí∞¢6ˆÁ7B∆ó7B“Fˆ7V÷VÁBÊvWDV∆V÷VÁD'îñBÇ'6Ê˜r◊7VG&R÷∆ó7F"ì∞¢ñbÇ∆ó7Bí&WGW&„∞¢∆ó7BÊñÊÊW$ÖD‘¬“"#∞¢ñbÜ&U7F'GW6˜&T6ˆ∆∆V7FñˆÁ4∆ˆFñÊrÇíí∞¢∆ó7BÊñÊÊW$ÖD‘¬“«6∆73“v◊WFVBs‚G∂W66TÖD‘¬á7F'GW6˜&T6ˆ∆∆V7FñˆÁ4∆ˆE7FFRÊ÷W76vR«¬$6&ñ6÷VÁFÚFFí7VG&ÊWfR‚‚‚"ó”¬˜Ê∞¢&WGW&„∞¢–¢ñbá7VG&T∆ˆE7FFRÁ7FGW2””“&∆ˆFñÊr"í∞¢∆ó7BÊñÊÊW$ÖD‘¬“«6∆73“v◊WFVBs‚G∂W66TÖD‘¬á7VG&T∆ˆE7FFRÊ÷W76vR«¬$6&ñ6÷VÁFÚ7VG&RÊWfR‚‚‚"ó”¬˜Ê∞¢&WGW&„∞¢–¢ñbá7VG&T∆ˆE7FFRÁ7FGW2””“&WFÇ◊&WVó&VB"í∞¢∆ó7BÊñÊÊW$ÖD‘¬“«6∆73“v◊WFVBs‚G∂W66TÖD‘¬á7VG&T∆ˆE7FFRÊ÷W76vR«¬$fí∆ˆvñ‚W"6&ñ6&R∆R7VG&RÊWfR‚"ó”¬˜Ê∞¢&WGW&„∞¢–¢ñbá7VG&T∆ˆE7FFRÁ7FGW2””“&W'&˜""í∞¢∆ó7BÊñÊÊW$ÖD‘¬“«6∆73“v◊WFVBs‚G∂W66TÖD‘¬á7VG&T∆ˆE7FFRÊ÷W76vR«¬$W'&˜&R6&ñ6÷VÁFÚFFí"ó”¬˜„∆'WGFˆ‚ñC“w6Ê˜r◊7VG&R◊&WG'í÷'F‚r6∆73“v'F‚'F‚◊&ñ÷'írGóS“v'WGFˆ‚sÂ&ó&˜f¬ˆ'WGFˆ„Ê∞¢∆ó7BÁVW'ï6V∆V7F˜"Ç"76Ê˜r◊7VG&R◊&WG'í÷'F‚"ìÚÊFDWfVÁD∆ó7FVÊW"Ç&6∆ñ6≤"¬Çí”‚7V'67&ñ&U7VG&RÇíì∞¢&WGW&„∞¢–¢6ˆÁ7B6V∆V7FVDFFT∂Wí“vWD7FófU7VG&TFFT∂WíÇì∞¢ñbÇ6V∆V7FVDFFT∂Wíí&WGW&„∞¢6ˆÁ7B7F˜&ñ6ÙFVƒvñ˜&ÊÚ“7VG&TÜó7F˜'î'îFFRÊvWBá6V∆V7FVDFFT∂Wíí«¬ÊWr÷Çì∞¢6ˆÁ7B6ˆ÷÷W76TÊWfR“'&íÊg&ˆ“Ü6ˆ÷÷W76T'îñBÁf«VW2ÇííÊfñ«FW"ÇÜ6ˆ÷÷W76í”‚∞¢6ˆÁ7B7VB“7F˜&ñ6ÙFVƒvñ˜&ÊÚÊvWBÜ6ˆ÷÷W76ÊñBí«¬∑”∞¢6ˆÁ7B&˜w2“'&íÊó4'&íá7VBÁ7VG&RíÚ7VBÁ7VG&R¢vWD∆Vv7ï7VG&U&˜w2á7VBì∞¢&WGW&‚&˜w2Á6ˆ÷RÜó57VG&&˜tfñ∆∆VBì∞¢“ì∞¢ñbÇ6ˆ÷÷W76TÊWfRÊ∆VÊwFÇí∞¢∆ó7BÊñÊÊW$ÖD‘¬“#«6∆73“v◊WFVBs‰ÊW77VÊ7VG&ÊWfR7&VFW"VW7FÚvñ˜&ÊÛ¬˜‚#∞¢&WGW&„∞¢–¢6ˆ÷÷W76TÊWfRÊf˜$V6ÇÇÜ6ˆ÷÷W76í”‚∞¢6ˆÁ7BóFV““Fˆ7V÷VÁBÊ7&VFTV∆V÷VÁBÇ&'Fñ6∆R"ì∞¢óFV“Ê6∆74Ê÷R“'7VG&÷óFV“#∞¢6ˆÁ7B7VB“7F˜&ñ6ÙFVƒvñ˜&ÊÚÊvWBÜ6ˆ÷÷W76ÊñBí«¬∑”∞¢6ˆÁ7B7VE&˜w2“'&íÊó4'&íá7VBÁ7VG&RíÚ7VBÁ7VG&R¢vWD∆Vv7ï7VG&U&˜w2á7VBì∞¢6ˆÁ7B&ñfW&ñ÷VÁFÚ“7VBÁ&ñfW&ñ÷VÁFÙFF¢ÚÊWrFFRÜG∑7VBÁ&ñfW&ñ÷VÁFÙFF’C££íÁFÙ∆ˆ6∆TFFU7G&ñÊrÇ&óB‘ïB"ê¢¢f˜&÷DFFT∂Wîf˜$Fó7∆íá6V∆V7FVDFFT∂Wíì∞¢6ˆÁ7B&˜w4áF÷¬“7VE&˜w2Ê÷Çá&˜r¬ñGÇí”‚∞¢6ˆÁ7B˜&&ñÙ∆&V¬“f˜&÷E7VG&˜&&ñÚá&˜rì∞¢6ˆÁ7BFWFñ«2“∞¢&˜rÊ6˜7VG&Ú∆'#„∆#Ô	˙y(ﬁ)»é˚àÚ6˜7VG&£¬ˆ#‚G∂W66TÖD‘¬á&˜rÊ6˜7VG&ó÷¢""¿¢˜&&ñÙ∆&V¬Ú∆'#„∆#Ô	˘Y#¬ˆ#‚G∂W66TÖD‘¬Ü˜&&ñÙ∆&V¬ó÷¢""¿¢&˜rÊñ◊ñÁFíÚ∆'#„∆#Ô	˘8“ñ◊ñÁFì£¬ˆ#‚G∂W66TÖD‘¬á&˜rÊñ◊ñÁFíó÷¢""¿¢&˜rÊÊ˜FRÚ∆'#„∆#Ô	˘9“Ê˜FS£¬ˆ#‚G∂W66TÖD‘¬á&˜rÊÊ˜FRó÷¢" ¢“Ê¶ˆñ‚Ç""ì∞¢&WGW&‚∆Fób6∆73“'7VG&◊6fVB◊&˜r"FF◊7VG&÷ñÊFWÉ“"G∂ñGá“#„«„∆'WGFˆ‚GóS“&'WGFˆ‚"6∆73“'7VG&÷VFóB÷∆ñÊ≤"FF÷6ˆ÷÷W76÷ñC“"G∂W66TÖD‘¬Ü6ˆ÷÷W76ÊñBó“"FF÷FFR÷∂Wì“"G∂W66TÖD‘¬á6V∆V7FVDFFT∂Wíó“"FF◊7VG&÷ñÊFWÉ“"G∂ñGá“"&ñ÷∆&V√“$÷ˆFñfñ67VG&ÊWfRG∂ñGÇ≤“FíG∂W66TÖD‘¬Ü6ˆ÷÷W76ÊÊˆ÷R«¬&6ˆ÷÷W76"ó“#Ô	˘R7VG&G∂ñGÇ≤”£¬ˆ'WGFˆ„‚G∂W66TÖD‘¬á&˜rÁW'6ˆÊ∆R«¬"“"ó“G∂FWFñ«7”∆'#„∆#Ô	˘©¢÷Wß¶íG∂ñGÇ≤”£¬ˆ#‚G∑&VÊFW$÷Wß¶î'WGFˆÁ4÷&∑Wá&˜rÊ÷Wß¶íó”¬˜„¬ˆFócÊ∞¢“íÊ¶ˆñ‚Ç""ì∞¢6ˆÁ7Bv&ÊñÊtó77VW2“'Vñ∆E7VG&v&ÊñÊtFWFñ«2Ü6ˆ÷÷W76¬7VE&˜w2ì∞¢6ˆÁ7Bv&ÊñÊt÷&∑W“v&ÊñÊtó77VW2Ê∆VÊwFÄ¢Ú∆Fób6∆73“'7VG&◊v&ÊñÊr◊w&#„∆'WGFˆ‚GóS“&'WGFˆ‚"6∆73“'7VG&◊v&ÊñÊr◊Fˆvv∆R"&ñ÷WáÊFVC“&f«6R"&ñ÷∆&V√“$÷˜7G&6ˆÁG&ˆ∆∆Ú7VG&ÊWfR#Ó)™˚àÛ¬ˆ'WGFˆ„„∆Fób6∆73“'7VG&◊v&ÊñÊr÷FWFñ«2ÜñFFV‚#„«„∆#Ó)™˚àÚ6ˆÁG&ˆ∆∆Ú7VG&¬ˆ#„¬˜„«V√‚G∑v&ÊñÊtó77VW2Ê÷ÇÜó77VRí”‚∆∆ì‚G∂W66TÖD‘¬Üó77VRÁ&W∆6RÇıÓ)™˚àı«2¢Ú¬""íó”¬ˆ∆ìÊíÊ¶ˆñ‚Ç""ó”¬˜V√„¬ˆFóc„¬ˆFócÊ ¢¢"#∞¢6ˆÁ7B6ˆFñ6T6ˆ÷÷W76“7G&ñÊrÜ6ˆ÷÷W76Ê6ˆFñ6R«¬""íÁG&ñ“Çì∞¢óFV“ÊñÊÊW$ÖD‘¬“ ¢∆Fób6∆73“'7VG&÷óFV“÷ÜVB7VG&÷6ˆ÷÷W76÷∆ñÊ≤"&ˆ∆S“&'WGFˆ‚"F&ñÊFWÉ“#"&ñ÷∆&V√“$&íFWGFv∆ñÚ6ˆ÷÷W76ÊWfRG∂W66TÖD‘¬Ü6ˆ÷÷W76ÊÊˆ÷R«¬$6ˆ÷÷W766VÁ¶Êˆ÷R"ó“#‡¢∆Fób6∆73“'7VG&÷6ˆ÷÷W76◊FóF∆R◊w&#‡¢«7G&ˆÊsÔ	˘8G∂W66TÖD‘¬Ü6ˆ÷÷W76ÊÊˆ÷R«¬$6ˆ÷÷W76ÊWfR"ó”¬˜7G&ˆÊs‡¢G∂vWE7VG&v˜&∂∆ñ÷FT6ˆFT∆ñÊT÷&∑WÜ6ˆ÷÷W76¬6ˆFñ6T6ˆ÷÷W76ó–¢∆Fób6∆73“'6Ê˜r◊7VG&÷÷WF#„«7‚6∆73“'ñ∆¬#Ó)ÿN˚àÚ6W'fó¶ñÚÊWfS¬˜7„„¬ˆFóc‡¢¬ˆFóc‡¢G∑v&ÊñÊt÷&∑W–¢¬ˆFóc‡¢«„∆#Ô	˘8Rvñ˜&ÊÛ£¬ˆ#‚G∂W66TÖD‘¬á&ñfW&ñ÷VÁFÚó”¬˜‡¢G∑&˜w4áF÷«–¢∞¢6ˆÁ7BÜVB“óFV“ÁVW'ï6V∆V7F˜"Ç"Á7VG&÷óFV“÷ÜVB"ì∞¢VÊE7VG&TÜVFW%&ó6¥7FñˆÁ2ÜÜVB¬6ˆ÷÷W76¬6V∆V7FVDFFT∂Wíì∞¢VÊDFDÜ˜W'4'WGFˆ‰ñd∆∆˜vVBÜÜVB¬6ˆ÷÷W76¬6V∆V7FVDFFT∂Wíì∞¢ÜVCÚÊFDWfVÁD∆ó7FVÊW"Ç&6∆ñ6≤"¬ÜWfVÁBí”‚∞¢ñbÜWfVÁBÁF&vWBÊ6∆˜6W7BÇ&'WGFˆ‚¬¬ñÁWB¬6V∆V7B¬FWáF&V"íí&WGW&„∞¢˜V‰6ˆ÷÷W76g&ˆ’7VG&RÜ6ˆ÷÷W76ì∞¢“ì∞¢ÜVCÚÊFDWfVÁD∆ó7FVÊW"Ç&∂WñF˜v‚"¬ÜWfVÁBí”‚∞¢ñbÜWfVÁBÊ∂Wí”“$VÁFW""bbWfVÁBÊ∂Wí”“""í&WGW&„∞¢ñbÜWfVÁBÁF&vWBÊ6∆˜6W7BÇ&'WGFˆ‚¬¬ñÁWB¬6V∆V7B¬FWáF&V"íí&WGW&„∞¢WfVÁBÁ&WfVÁDFVfV«BÇì∞¢˜V‰6ˆ÷÷W76g&ˆ’7VG&RÜ6ˆ÷÷W76ì∞¢“ì∞¢óFV“ÁVW'ï6V∆V7F˜$∆¬Ç"Á7VG&÷VFóB÷∆ñÊ≤"íÊf˜$V6ÇÇÜ'F‚í”‚∞¢'F‚ÊFDWfVÁD∆ó7FVÊW"Ç&6∆ñ6≤"¬ÜWfVÁBí”‚∞¢WfVÁBÁ&WfVÁDFVfV«BÇì∞¢WfVÁBÁ7F˜&˜vFñˆ‚Çì∞¢˜VÂ7VG&6ˆ◊˜6óFñˆ‰VFóF˜"Ü'F‚ÊFF6WBÊ6ˆ÷÷W76ñB«¬6ˆ÷÷W76ÊñB¬'F‚ÊFF6WBÊFFT∂Wí«¬6V∆V7FVDFFT∂Wí¬ÁV÷&W"Ü'F‚ÊFF6WBÁ7VG&ñÊFWÇí«¬ì∞¢“ì∞¢“ì∞¢óFV“ÁVW'ï6V∆V7F˜$∆¬Ç"Ê÷Wß¶Ú÷6Üó÷'F‚"íÊf˜$V6ÇÇÜ'F‚í”‚∞¢'F‚ÊFDWfVÁD∆ó7FVÊW"Ç&6∆ñ6≤"¬Çí”‚˜V‰gVV≈vRÜ'F‚ÊFF6WBÊ÷Wß¶Ú«¬""íì∞¢“ì∞¢óFV“ÁVW'ï6V∆V7F˜"Ç%∂FF◊v˜&∂∆ñ÷FR÷6ˆ÷÷W76“"ìÚÊFDWfVÁD∆ó7FVÊW"Ç&6∆ñ6≤"¬ÜWfVÁBí”‚∞¢WfVÁBÁ&WfVÁDFVfV«BÇì∞¢WfVÁBÁ7F˜&˜vFñˆ‚Çì∞¢˜VÂ7VG&v˜&∂∆ñ÷FU6fWGíÜ6ˆ÷÷W76¬6V∆V7FVDFFT∂Wíì∞¢“ì∞¢óFV“ÁVW'ï6V∆V7F˜"Ç%∂FF◊v˜&∂∆ñ÷FR◊FV◊W&GW&R÷6ˆ÷÷W76“"ìÚÊFDWfVÁD∆ó7FVÊW"Ç&6∆ñ6≤"¬ÜWfVÁBí”‚∞¢WfVÁBÁ&WfVÁDFVfV«BÇì∞¢WfVÁBÁ7F˜&˜vFñˆ‚Çì∞¢˜VÂ7VG&v˜&∂∆ñ÷FU6fWGíÜ6ˆ÷÷W76¬6V∆V7FVDFFT∂Wí¬≤&VfW$÷¶˜&óGî∆ˆ6Fñˆ„¢G'VR¬&VfW$fW&vUFV◊W&GW&S¢G'VR“ì∞¢“ì∞¢6ˆÁ7Bv&ÊñÊuFˆvv∆R“óFV“ÁVW'ï6V∆V7F˜"Ç"Á7VG&◊v&ÊñÊr◊Fˆvv∆R"ì∞¢v&ÊñÊuFˆvv∆SÚÊFDWfVÁD∆ó7FVÊW"Ç&6∆ñ6≤"¬ÜWfVÁBí”‚∞¢WfVÁBÁ7F˜&˜vFñˆ‚Çì∞¢6ˆÁ7BFWFñ«2“óFV“ÁVW'ï6V∆V7F˜"Ç"Á7VG&◊v&ÊñÊr÷FWFñ«2"ì∞¢6ˆÁ7Bó4ÜñFFV‚“FWFñ«3ÚÊ6∆74∆ó7BÊ6ˆÁFñÁ2Ç&ÜñFFV‚"ì∞¢FWFñ«3ÚÊ6∆74∆ó7BÁFˆvv∆RÇ&ÜñFFV‚"¬ó4ÜñFFV‚ì∞¢v&ÊñÊuFˆvv∆RÁ6WDGG&ñ'WFRÇ&&ñ÷WáÊFVB"¬ó4ÜñFFV‚Ú'G'VR"¢&f«6R"ì∞¢“ì∞¢∆ó7BÊVÊD6Üñ∆BÜóFV“ì∞¢“ì∞ß–†¶gVÊ7Fñˆ‚7ñÊ56Ê˜uvVFÜW%ÊV¬Çí∞¢6ˆÁ7B7V÷÷'í“Fˆ7V÷VÁBÊvWDV∆V÷VÁD'îñBÇ'6Ê˜r◊vVFÜW"◊7V÷÷'í"ì∞¢6ˆÁ7B&ó6∑2“Fˆ7V÷VÁBÊvWDV∆V÷VÁD'îñBÇ'6Ê˜r◊vVFÜW"◊&ó6∑2"ì∞¢ñbá7V÷÷'íbbVíÁvVFÜW%7V÷÷'íí7V÷÷'íÁFWáD6ˆÁFVÁB“VíÁvVFÜW%7V÷÷'íÁFWáD6ˆÁFVÁB«¬$6&ñ6÷VÁFÚ÷WFVÚ‚‚‚#∞¢ñbá&ó6∑2bbVíÁvVFÜW%&ó6∑2í&ó6∑2ÊñÊÊW$ÖD‘¬“VíÁvVFÜW%&ó6∑2ÊñÊÊW$ÖD‘√∞ß–†¶gVÊ7Fñˆ‚6ˆÊfñwW&U6Ê˜u6ñFT÷VÁRÜó56Ê˜rí∞¢6ˆÁ7B∆∆˜vVB“ÊWr6WBÖ≤&˜V‚◊ÊV¬÷6ˆ÷÷W76R"¬&˜V‚◊ÊV¬◊7VG&R"¬&˜V‚◊ÊV¬◊W'6ˆÊ∆R"¬&˜V‚◊ÊV¬÷÷Wß¶í"¬&˜V‚◊ÊV¬◊WFVÁFí"¬&˜V‚÷Ü˜W'2÷'F‚%“ì∞¢Fˆ7V÷VÁBÊ&ˆGíÊ6∆74∆ó7BÁFˆvv∆RÇ'6Ê˜r÷÷ÊvV÷VÁB÷6ˆÁFWáB"¬&ˆˆ∆V‚Üó56Ê˜ríì∞¢Fˆ7V÷VÁBÁVW'ï6V∆V7F˜$∆¬Ç"76ñFR÷÷VÁRÊ÷VÁR◊FóF∆R÷'F‚"íÊf˜$V6ÇÇÜ'WGFˆ‚í”‚∞¢'WGFˆ‚Ê6∆74∆ó7BÁFˆvv∆RÇ&ÜñFFV‚"¬&ˆˆ∆V‚Üó56Ê˜ríbb∆∆˜vVBÊÜ2Ü'WGFˆ‚ÊñBíì∞¢“ì∞ß–†¶gVÊ7Fñˆ‚&VÊFW%6Ê˜u6W'fñ6RÇí∞¢7ñÊ56Ê˜uvVFÜW%ÊV¬Çì∞¢7ñÊ57VG&TFFTñÁWG2Çì∞¢6ˆÊfñwW&U6Ê˜u6ñFT÷VÁRáG'VRì∞¢&VÊFW%6Ê˜u6W'fñ6T6ˆ÷÷W76RÇì∞ß–¶gVÊ7Fñˆ‚7V'67&ñ&U6Ê˜u6W'fñ6T6ˆ∆∆V7FñˆÁ2Çí∞¢7F˜6Ê˜u6W'fñ6T6ˆ∆∆V7FñˆÁ2Çì∞¢ñbÇF"«¬7W'&VÁEW6W"«¬ó56Ê˜u6W'fñ6T6ˆÁFWáBÇíí&WGW&‚&ˆ÷ó6RÁ&W6ˆ«fRÜf«6Rì∞¢ˆ&¶V7BÊVÁG&ñW2á≤6∆ñVÁG3¢4‰ıuı4U%dî4UÙ4ÙƒƒT5DîÙÂ2Ê6∆ñVÁG2¬&˜WFW3¢4‰ıuı4U%dî4UÙ4ÙƒƒT5DîÙÂ2Á&˜WFW2¬fVÜñ6∆W3¢4‰ıuı4U%dî4UÙ4ÙƒƒT5DîÙÂ2ÁfVÜñ6∆W2¬˜W&F˜'3¢4‰ıuı4U%dî4UÙ4ÙƒƒT5DîÙÂ2Ê˜W&F˜'2¬&W˜'G3¢4‰ıuı4U%dî4UÙ4ÙƒƒT5DîÙÂ2Á&W˜'G2“íÊf˜$V6ÇÇÖ∂∂Wí¬6ˆ∆∆V7Fñˆ‰Ê÷U“í”‚∞¢6Ê˜u6W'fñ6UVÁ7V'67&ñ&W'2ÁW6ÇÜF"Ê6ˆ∆∆V7Fñˆ‚Ü6ˆ∆∆V7Fñˆ‰Ê÷RíÊ˜&FW$'íÇ&7&VFVDB"¬&FW62"íÊˆÂ6Ê6Ü˜BÇá6Ê6Ü˜Bí”‚∞¢6Ê˜u6W'fñ6U7FFU∂∂Wï““6Ê6Ü˜BÊFˆ72Ê÷ÇÜFˆ2í”‚á≤ñC¢Fˆ2ÊñB¬‚‚ÊFˆ2ÊFFÇí“íì∞¢&VÊFW%6Ê˜u6W'fñ6RÇì∞¢“¬ÜW'&˜"í”‚6ˆÁ6ˆ∆RÁv&‚ÜW'&˜&R6&ñ6÷VÁFÚG∂6ˆ∆∆V7Fñˆ‰Ê÷W”¶¬W'&˜"ííì∞¢“ì∞¢&WGW&‚&ˆ÷ó6RÁ&W6ˆ«fRáG'VRì∞ß–†¶7ñÊ2gVÊ7Fñˆ‚FE6Ê˜u6W'fñ6TóFV“áGóRí∞¢ñbÇ6‰÷ÊvTFFÇíí&WGW&‚∆W'BÇ%6ˆ∆ÚF÷ñ‚\;"÷ˆFñfñ6&Rñ¬6W'fó¶ñÚÊWfR‚"ì∞¢6ˆÁ7B6ˆÊfñr“∞¢6∆ñVÁG3¢≤6ˆ∆∆V7Fñˆ„¢4‰ıuı4U%dî4UÙ4ÙƒƒT5DîÙÂ2Ê6∆ñVÁG2¬&ˆ◊C¢$Êˆ÷R6ˆ◊VÊRÚ6∆ñVÁFRÊWfR"¬fñV∆C¢&Êˆ÷R"“¿¢&˜WFW3¢≤6ˆ∆∆V7Fñˆ„¢4‰ıuı4U%dî4UÙ4ÙƒƒT5DîÙÂ2Á&˜WFW2¬&ˆ◊C¢$Êˆ÷RW&6˜'6ÚÊWfR"¬fñV∆C¢&Êˆ÷R"“¿¢fVÜñ6∆W3¢≤6ˆ∆∆V7Fñˆ„¢4‰ıuı4U%dî4UÙ4ÙƒƒT5DîÙÂ2ÁfVÜñ6∆W2¬&ˆ◊C¢$Êˆ÷R÷Wß¶ÚÊWfR"¬fñV∆C¢&Êˆ÷R"“¿¢˜W&F˜'3¢≤6ˆ∆∆V7Fñˆ„¢4‰ıuı4U%dî4UÙ4ÙƒƒT5DîÙÂ2Ê˜W&F˜'2¬&ˆ◊C¢$Êˆ÷R˜W&F˜&RÊWfR"¬fñV∆C¢&Êˆ÷R"“¿¢&W˜'G3¢≤6ˆ∆∆V7Fñˆ„¢4‰ıuı4U%dî4UÙ4ÙƒƒT5DîÙÂ2Á&W˜'G2¬&ˆ◊C¢%FóFˆ∆Ú6VvÊ∆¶ñˆÊRÊWfR"¬fñV∆C¢'FóFˆ∆Ú"–¢’∑GóU”∞¢ñbÇ6ˆÊfñrí&WGW&„∞¢6ˆÁ7Bf«VR“7G&ñÊrávñÊF˜rÁ&ˆ◊BÜ6ˆÊfñrÁ&ˆ◊B¬""í«¬""íÁG&ñ“Çì∞¢ñbÇf«VRí&WGW&„∞¢6ˆÁ7BÊ˜FR“7G&ñÊrávñÊF˜rÁ&ˆ◊BÇ$Ê˜FRÚFWGFv∆íÜ˜¶ñˆÊ∆Rí"¬""í«¬""íÁG&ñ“Çì∞¢vóBF"Ê6ˆ∆∆V7Fñˆ‚Ü6ˆÊfñrÊ6ˆ∆∆V7Fñˆ‚íÊFBá≤∂6ˆÊfñrÊfñV∆E”¢f«VR¬Ê˜FR¬7&VFVDC¢fó&V&6RÊfó&W7F˜&R‰fñV∆Ef«VRÁ6W'fW%Fñ÷W7F◊Çí¬7&VFVD'ì¢7W'&VÁEW6W#ÚÊV÷ñ¬«¬""“ì∞ß–†¶7ñÊ2gVÊ7Fñˆ‚FV∆WFU6Ê˜u6W'fñ6TóFV“áGóR¬ñBí∞¢ñbÇ6‰÷ÊvTFFÇíí&WGW&‚∆W'BÇ%6ˆ∆ÚF÷ñ‚\;"V∆ñ÷ñÊ&RV∆V÷VÁFíFV¬6W'fó¶ñÚÊWfR‚"ì∞¢6ˆÁ7B6ˆ∆∆V7Fñˆ‚“4‰ıuı4U%dî4UÙ4ÙƒƒT5DîÙÂ5∑GóU”∞¢ñbÇ6ˆ∆∆V7Fñˆ‚«¬ñB«¬vñÊF˜rÊ6ˆÊfó&“Ç$V∆ñ÷ñÊ&RVW7FÚV∆V÷VÁFÚÊWfSÚ"íí&WGW&„∞¢vóBF"Ê6ˆ∆∆V7Fñˆ‚Ü6ˆ∆∆V7Fñˆ‚íÊFˆ2ÜñBíÊFV∆WFRÇì∞ß–†¶gVÊ7Fñˆ‚ÜÊF∆U6Ê˜u6W'fñ6T÷VÁT7Fñˆ‚Ü7Fñˆ‚í∞¢ñbÜ7Fñˆ‚””“'6WGFñÊw2"í&WGW&‚∆W'BÇ$ñ◊˜7F¶ñˆÊí6W'fó¶ñÚÊWfS¢÷ˆGV∆ÚFVFñ6FÚR6W&FÚF∆∆÷ÁWFVÁ¶ñˆÊRfW&FR‚"ì∞¢ñbÜ7Fñˆ‚””“&FB÷6∆ñVÁB"í&WGW&‚FE6Ê˜u6W'fñ6TóFV“Ç&6∆ñVÁG2"ì∞¢ñbÜ7Fñˆ‚””“&FB◊&˜WFR"í&WGW&‚FE6Ê˜u6W'fñ6TóFV“Ç'&˜WFW2"ì∞¢ñbÜ7Fñˆ‚””“&÷ÊvR◊fVÜñ6∆W2"í&WGW&‚FE6Ê˜u6W'fñ6TóFV“Ç'fVÜñ6∆W2"ì∞¢ñbÜ7Fñˆ‚””“&÷ÊvR÷˜W&F˜'2"í&WGW&‚FE6Ê˜u6W'fñ6TóFV“Ç&˜W&F˜'2"ì∞ß–†¶Fˆ7V÷VÁBÊvWDV∆V÷VÁD'îñBÇ'6Ê˜r◊6W'fñ6R÷'F‚"ìÚÊFDWfVÁD∆ó7FVÊW"Ç&6∆ñ6≤"¬˜VÂ6Ê˜u6W'fñ6UvRì∞¶Fˆ7V÷VÁBÊvWDV∆V÷VÁD'îñBÇ'6Ê˜r◊6W'fñ6R÷&6≤÷'F‚"ìÚÊFDWfVÁD∆ó7FVÊW"Ç&6∆ñ6≤"¬6∆˜6U6Ê˜u6W'fñ6UvRì∞¶Fˆ7V÷VÁBÊvWDV∆V÷VÁD'îñBÇ'6Ê˜r◊6W'fñ6R÷÷VÁR÷'F‚"ìÚÊFDWfVÁD∆ó7FVÊW"Ç&6∆ñ6≤"¬Çí”‚∞¢6ˆÊfñwW&U6Ê˜u6ñFT÷VÁRáG'VRì∞¢˜VÂ6ñFT÷VÁRÇì∞¢&WGW&„∞¢6ˆÁ7B÷VÁR“Fˆ7V÷VÁBÊvWDV∆V÷VÁD'îñBÇ'6Ê˜r◊6W'fñ6R÷÷VÁR"ì∞¢6ˆÁ7Bó4˜V‚“÷VÁSÚÊ6∆74∆ó7BÊ6ˆÁFñÁ2Ç&ÜñFFV‚"ì∞¢÷VÁSÚÊ6∆74∆ó7BÁFˆvv∆RÇ&ÜñFFV‚"¬ó4˜V‚ì∞¢Fˆ7V÷VÁBÊvWDV∆V÷VÁD'îñBÇ'6Ê˜r◊6W'fñ6R÷÷VÁR÷'F‚"ìÚÁ6WDGG&ñ'WFRÇ&&ñ÷WáÊFVB"¬7G&ñÊrÇó4˜V‚íì∞ß“ì∞¶Fˆ7V÷VÁBÊvWDV∆V÷VÁD'îñBÇ'6Ê˜r◊6W'fñ6R÷÷VÁR"ìÚÊFDWfVÁD∆ó7FVÊW"Ç&6∆ñ6≤"¬ÜWfVÁBí”‚∞¢6ˆÁ7B'WGFˆ‚“WfVÁBÁF&vWCÚÊ6∆˜6W7CÚ‚Ç%∂FF◊6Ê˜r÷7FñˆÂ“"ì∞¢ñbÇ'WGFˆ‚í&WGW&„∞¢Fˆ7V÷VÁBÊvWDV∆V÷VÁD'îñBÇ'6Ê˜r◊6W'fñ6R÷÷VÁR"ìÚÊ6∆74∆ó7BÊFBÇ&ÜñFFV‚"ì∞¢ÜÊF∆U6Ê˜u6W'fñ6T÷VÁT7Fñˆ‚Ü'WGFˆ‚ÊFF6WBÁ6Ê˜t7Fñˆ‚«¬""ì∞ß“ì∞¶Fˆ7V÷VÁBÊvWDV∆V÷VÁD'îñBÇ'6Ê˜r◊6W'fñ6R◊vR"ìÚÊFDWfVÁD∆ó7FVÊW"Ç&6∆ñ6≤"¬ÜWfVÁBí”‚∞¢6ˆÁ7BFV¬“WfVÁBÁF&vWCÚÊ6∆˜6W7CÚ‚Ç%∂FF◊6Ê˜r÷FV∆WFU“"ì∞¢ñbÜFV¬í∞¢6ˆÁ7B∑GóR¬ñE““7G&ñÊrÜFV¬ÊFF6WBÁ6Ê˜tFV∆WFR«¬""íÁ7∆óBÇ#¢"ì∞¢FV∆WFU6Ê˜u6W'fñ6TóFV“áGóR¬ñBì∞¢&WGW&„∞¢–¢6ˆÁ7BÊb“WfVÁBÁF&vWCÚÊ6∆˜6W7CÚ‚Ç%∂FF◊6Ê˜r÷ÊfñvFU“"ì∞¢ñbÜÊbí∆W'BÇ$Êfñv¶ñˆÊRW&6˜'6ÚÊWfR&ˆÁF¢6ˆ∆∆Vv6ˆ˜&FñÊFRFVFñ6FRÊV∆∆÷ÊWfR‚"ì∞ß“ì∞†¶Fˆ7V÷VÁBÊvWDV∆V÷VÁD'îñBÇ'6Ê˜r◊&Vg&W6Ç÷÷'F‚"ìÚÊFDWfVÁD∆ó7FVÊW"Ç&6∆ñ6≤"¬&Vg&W6Ñ∆ñ6Fñˆ‰FFì∞†¶gVÊ7Fñˆ‚vWD7FófóGïW&ñˆE&ÊvRÇí∞¢6ˆÁ7BÊ˜r“ÊWrFFRÇì∞¢6ˆÁ7B7F'B“ÊWrFFRÜÊ˜rì∞¢6ˆÁ7BVÊB“ÊWrFFRÜÊ˜rì∞¢6ˆÁ7B÷ˆFR“VíÊ7FófUW6W'5W&ñˆE6V∆V7CÚÁf«VR«¬'FˆFí#∞¢7F'BÁ6WDÜ˜W'2É¬¬¬ì∞¢VÊBÁ6WDÜ˜W'2É#2¬Sí¬Sí¬ììíì∞¢ñbÜ÷ˆFR””“'ñW7FW&Fí"í≤7F'BÁ6WDFFRá7F'BÊvWDFFRÇí“ì≤VÊBÁ6WDFFRÜVÊBÊvWDFFRÇí“ì≤–¢ñbÜ÷ˆFR””“&∆7Cr"í7F'BÁ6WDFFRá7F'BÊvWDFFRÇí“bì∞¢ñbÜ÷ˆFR””“&÷ˆÁFÇ"í7F'BÁ6WDFFRÉì∞¢ñbÜ÷ˆFR””“&7W7Fˆ“"í∞¢ñbáVíÊ7FófUW6W'57F'DFFSÚÁf«VRí7F'BÁ6WEFñ÷RÜÊWrFFRÜG∑VíÊ7FófUW6W'57F'DFFRÁf«VW’C££íÊvWEFñ÷RÇíì∞¢ñbáVíÊ7FófUW6W'4VÊDFFSÚÁf«VRíVÊBÁ6WEFñ÷RÜÊWrFFRÜG∑VíÊ7FófUW6W'4VÊDFFRÁf«VW’C#3£Sì£SñíÊvWEFñ÷RÇíì∞¢–¢&WGW&‚≤7F'B¬VÊB”∞ß–†¶gVÊ7Fñˆ‚f˜&÷D7FófóGîFFRáf«VRí∞¢6ˆÁ7B◊2“fó&W7F˜&TFFUFÙ÷ñ∆∆ó2áf«VRì∞¢&WGW&‚◊2ÚÊWrFFRÜ◊2íÁFÙ∆ˆ6∆U7G&ñÊrÇ&óB‘ïB"í¢"“#∞ß–†¶gVÊ7Fñˆ‚vWEW6W$Fó7∆îÊ÷RáW6W"“∑“í∞¢&WGW&‚W6W"ÊFó7∆îÊ÷R«¬W6W"ÊÊ÷R«¬W6W"ÁW6W$Ê÷R«¬W6W"ÊV÷ñ¬«¬%WFVÁFR#∞ß–†¶gVÊ7Fñˆ‚vWEW6W%&ˆ∆RáW6W"“∑“í∞¢&WGW&‚F÷ñ‰V÷ñ«2ÊÜ2ÜÊ˜&÷∆ó¶TV÷ñ¬áW6W"ÊV÷ñ¬ííÚ$F÷ñ‚"¢$˜W&F˜&R#∞ß–††¶6ˆÁ7B5DïdUÙƒÙuı5DEU5ÙıDîÙÂ2“≤$W'FÚ"¬$ñ‚6ˆÁG&ˆ∆∆Ú"¬%&ó6ˆ«FÚ"¬$ñvÊ˜&FÚ%”∞¶6ˆÁ7B5DïdUÙƒÙuı5DEU5ı5Dı$tUÙ¥Uí“&7FófT∆ˆu&ˆ&∆V’7FGW6W2#∞†¶gVÊ7Fñˆ‚vWD7W'&VÁEfñWtÊ÷RÇí∞¢6ˆÁ7BÜ6Ç“7G&ñÊrávñÊF˜rÊ∆ˆ6Fñˆ‚ÊÜ6Ç«¬""íÁ&W∆6RÇı‚2Ú¬""ì∞¢ñbÜÜ6ÇÁ7F'G5vóFÇÇ&6ˆ÷÷W76“"íí&WGW&‚'6T6ˆ÷÷W76Ü6ÇÜ2G∂Ü6á÷íÊñ◊ñÁFÚÚ$FWGFv∆ñÚñ◊ñÁFÚ"¢$6ˆ÷÷W76Úñ◊ñÁFí#∞¢ñbÜÜ6Ç””“&FWGFv∆ñÚ◊WFVÁFí÷GFófí"í&WGW&‚$6ˆÁ6ˆ∆R÷÷ñÊó7G&F˜&RÚ&Vvó7G&Ú¶ñˆÊí#∞¢ñbÜÜ6Ç””“&˜&R"í&WGW&‚$vW7FñˆÊR˜&R#∞¢ñbÜÜ6Ç””“'6VvÊ∆¶ñˆÊí"í&WGW&‚%6VvÊ∆¶ñˆÊí#∞¢ñbÜÜ6Ç””“&Fˆ7V÷VÁFí"í&WGW&‚$Fˆ7V÷VÁFí#∞¢ñbÜÜ6Ç””“&6∆VÊF&ñÚ"í&WGW&‚$6∆VÊF&ñÚ6ˆÊFófó6Ú#∞¢&WGW&‚Ü6Ç«¬$Üˆ÷R#∞ß–†¶gVÊ7Fñˆ‚&VD7FófT∆ˆu7FGW6W2Çí∞¢G'í≤&WGW&‚•4Ù‚Á'6RÜ∆ˆ6≈7F˜&vRÊvWDóFV“Ñ5DïdUÙƒÙuı5DEU5ı5Dı$tUÙ¥Uíí«¬'∑“"ì≤–¢6F6ÇÖÚí≤&WGW&‚∑”≤–ß–†¶gVÊ7Fñˆ‚vWD7FófT∆ˆu7FGW2ÜñBí∞¢&WGW&‚&VD7FófT∆ˆu7FGW6W2Çï∂ñE“«¬$W'FÚ#∞ß–†¶gVÊ7Fñˆ‚6WD7FófT∆ˆu7FGW2ÜñB¬7FGW2í∞¢ñbÇñBí&WGW&„∞¢6ˆÁ7B7FGW6W2“&VD7FófT∆ˆu7FGW6W2Çì∞¢7FGW6W5∂ñE““5DïdUÙƒÙuı5DEU5ÙıDîÙÂ2ÊñÊ6«VFW2á7FGW2íÚ7FGW2¢$W'FÚ#∞¢∆ˆ6≈7F˜&vRÁ6WDóFV“Ñ5DïdUÙƒÙuı5DEU5ı5Dı$tUÙ¥Uí¬•4Ù‚Á7G&ñÊvñgíá7FGW6W2íì∞ß–†¶gVÊ7Fñˆ‚ó4W'&˜$7FófóGíÜ∆ˆr“∑“í∞¢&WGW&‚ˆW'&˜&W∆W'&˜'∆fñ«∆fñ∆VG«W&÷ó76ñˆÁ∆FVÊñVBˆíÁFW7BÜG∂∆ˆrÊ7FñˆÂGóR«¬"'“G∂∆ˆrÊ7Fñˆ‰FW67&óFñˆ‚«¬"'“G∂∆ˆrÊFWFñ¬«¬"'“G∂∆ˆrÊW'&˜$6ˆFR«¬"'÷ì∞ß–†¶gVÊ7Fñˆ‚vWE&ˆ&∆V‘w&˜W∂WíÜ∆ˆr“∑“í∞¢ñbÇó4W'&˜$7FófóGíÜ∆ˆríí&WGW&‚∆ˆrÊñB«¬G∂∆ˆrÊ7FñˆÂGóW““G∂fó&W7F˜&TFFUFÙ÷ñ∆∆ó2Ü∆ˆrÊ7&VFVDBó÷∞¢&WGW&‚∂∆ˆrÊ7FñˆÂGóR¬∆ˆrÊW'&˜$6ˆFR¬∆ˆrÊfó&W7F˜&T6ˆ∆∆V7Fñˆ‚¬∆ˆrÊfó&W7F˜&T˜W&Fñˆ‚¬∆ˆrÊFWFñ¬«¬∆ˆrÊ7Fñˆ‰FW67&óFñˆ‚¬∆ˆrÊ6ˆ÷÷W76ñB«¬∆ˆrÊ6ˆ÷÷W76Ê÷R¬∆ˆrÊñ◊ñÁFÙñB«¬∆ˆrÊñ◊ñÁFÙÊ÷U“Ê÷áb”‚7G&ñÊráb«¬"“"íÁFÙ∆˜vW$66RÇíÁG&ñ“ÇííÊ¶ˆñ‚Ç'¬"ì∞ß–†¶gVÊ7Fñˆ‚'Vñ∆D7FófT∆ˆtw&˜W2Ü∆ˆw2“µ“í∞¢6ˆÁ7Bw&˜W2“ÊWr÷Çì∞¢∆ˆw2Êf˜$V6ÇÇÜ∆ˆrí”‚∞¢6ˆÁ7B∂Wí“vWE&ˆ&∆V‘w&˜W∂WíÜ∆ˆrì∞¢ñbÇw&˜W2ÊÜ2Ü∂Wíííw&˜W2Á6WBÜ∂Wí¬≤∂Wí¬&ñ÷'ì¢∆ˆr¬∆ˆw3¢µ““ì∞¢w&˜W2ÊvWBÜ∂WííÊ∆ˆw2ÁW6ÇÜ∆ˆrì∞¢“ì∞¢&WGW&‚'&íÊg&ˆ“Üw&˜W2Áf«VW2ÇííÊ÷ÇÜw&˜Wí”‚∞¢w&˜WÊ∆ˆw2Á6˜'BÇÜ¬"í”‚fó&W7F˜&TFFUFÙ÷ñ∆∆ó2Ü"Ê7&VFVDBí“fó&W7F˜&TFFUFÙ÷ñ∆∆ó2ÜÊ7&VFVDBíì∞¢w&˜WÁ&ñ÷'í“w&˜WÊ∆ˆw5≥“«¬w&˜WÁ&ñ÷'ì∞¢&WGW&‚w&˜W∞¢“ì∞ß–†¶gVÊ7Fñˆ‚7FófT∆ˆu7V÷÷'ïFWáBÜw&˜Wí∞¢6ˆÁ7B∆ˆr“w&˜WÁ&ñ÷'í«¬∑”∞¢6ˆÁ7B∆ñÊW2“∞¢FóÛ¢G∂∆ˆrÊ7FñˆÂGóR«¬&¶ñˆÊR'÷¿¢7FFÛ¢G∂vWD7FófT∆ˆu7FGW2Ü∆ˆrÊñB«¬w&˜WÊ∂Wíó÷¿¢6˜6:Ç7V66W76Û¢G∂∆ˆrÊ7Fñˆ‰FW67&óFñˆ‚«¬∆ˆrÊFWFñ¬«¬$WfVÁFÚ&Vvó7G&FÚ'÷¿¢WFVÁFS¢G∂∆ˆrÁW6W$Ê÷R«¬∆ˆrÁW6W$V÷ñ¬«¬"“'÷¿¢VÊFÛ¢G∂f˜&÷D7FófóGîFFRÜ∆ˆrÊ7&VFVDBó÷¿¢fñWr˜vñÊ¢G∂∆ˆrÁfñWtÊ÷R«¬$FfW&ñfñ6&R'÷¿¢V«6ÁFR&V◊WFÛ¢G∂∆ˆrÊ'WGFˆ‰∆&V¬«¬$Êˆ‚ñÊFñ6FÚ'÷¿¢6ˆ÷÷W76¢G∂∆ˆrÊ6ˆ÷÷W76Ê÷R«¬∆ˆrÊ6ˆ÷÷W76ñB«¬"“'÷¿¢ñ◊ñÁFÛ¢G∂∆ˆrÊñ◊ñÁFÙÊ÷R«¬∆ˆrÊñ◊ñÁFÙñB«¬"“'÷¿¢W'&˜&RFV6Êñ6Ú6ˆ◊∆WFÛ¢G∂∆ˆrÁFV6ÜÊñ6ƒW'&˜"«¬∆ˆrÊFWFñ¬«¬"“'÷¿¢˜76ñ&ñ∆R6W6¢G∂∆ˆrÁ˜76ñ&∆T6W6R«¬Üó4W'&˜$7FófóGíÜ∆ˆríÚ$FÊ∆óß¶&S¢fVFW&RW'&˜&RFV6Êñ6Ú¬WFVÁFR¬fñWrRW&÷W76í‚"¢"“"ó÷¿¢7VvvW&ñ÷VÁFÛ¢G∂∆ˆrÁ&W6ˆ«WFñˆ‰ÜñÁB«¬Üó4W'&˜$7FófóGíÜ∆ˆríÚ%&ó&ˆGW'&Rñ¬f«W76Ú¬fW&ñfñ6&RW&÷W76íˆFFíR6ˆÁG&ˆ∆∆&R6ˆÁ6ˆ∆R‚"¢"“"ó÷ ¢”∞¢ñbÖ7G&ñÊrÜ∆ˆrÊ7FñˆÂGóR«¬""íÊñÊ6«VFW2Ç&fó&W7F˜&R"íí∞¢∆ñÊW2ÁW6ÇÜ6ˆ∆∆V7Fñˆ‚fó&W7F˜&S¢G∂∆ˆrÊfó&W7F˜&T6ˆ∆∆V7Fñˆ‚«¬$FfW&ñfñ6&R'÷ì∞¢∆ñÊW2ÁW6ÇÜ˜W&¶ñˆÊRFVÁFF¢G∂∆ˆrÊfó&W7F˜&T˜W&Fñˆ‚«¬$FfW&ñfñ6&R'÷ì∞¢∆ñÊW2ÁW6ÇÜFóÚ&∆ˆ66Û¢G∂∆ˆrÊfó&W7F˜&Tfñ«W&UGóR«¬'W&÷W76ÚÊVvFÚÚ&WFRÚFFÚ÷Ê6ÁFRFfW&ñfñ6&R'÷ì∞¢∆ñÊW2ÁW6ÇÜFFíÊˆ‚6«fFì¢G∂∆ˆrÁVÁ6fVDFF«¬$FfW&ñfñ6&R'÷ì∞¢∆ñÊW2ÁW6ÇÜ&Vvˆ∆fó&W7F˜&R˜76ñ&ñ∆S¢G∂∆ˆrÊfó&W7F˜&U'V∆TÜñÁB«¬$6ˆÁG&ˆ∆∆&R∆∆˜r&VB˜w&óFRFV∆∆6ˆ∆∆V7Fñˆ‚ñÁFW&W76F‚'÷ì∞¢–¢ñbÜw&˜WÊ∆ˆw2Ê∆VÊwFÇ‚í∆ñÊW2ÁW6ÇÜW'&˜&R&óWGWFÚG∂w&˜WÊ∆ˆw2Ê∆VÊwFá“fˆ«FS¢G∂w&˜WÊ∆ˆw2Ê÷Ü¬”‚G∂¬ÁW6W$Ê÷R«¬¬ÁW6W$V÷ñ¬«¬"“'“G∂f˜&÷D7FófóGîFFRÜ¬Ê7&VFVDBó“ÇG∂¬ÁfñWtÊ÷R«¬'fñWsÚ'“ñíÊ¶ˆñ‚Ç#≤"ó÷ì∞¢&WGW&‚∆ñÊW2Ê¶ˆñ‚Ç%∆‚"ì∞ß–†¶gVÊ7Fñˆ‚6˜ïFWáEFÙ6∆ó&ˆ&BáFWáBí∞¢ñbÜÊfñvF˜"Ê6∆ó&ˆ&CÚÁw&óFUFWáBí&WGW&‚ÊfñvF˜"Ê6∆ó&ˆ&BÁw&óFUFWáBáFWáBì∞¢6ˆÁ7BFWáF&V“Fˆ7V÷VÁBÊ7&VFTV∆V÷VÁBÇ'FWáF&V"ì∞¢FWáF&VÁf«VR“FWáC∞¢Fˆ7V÷VÁBÊ&ˆGíÊVÊD6Üñ∆BáFWáF&Vì∞¢FWáF&VÁ6V∆V7BÇì∞¢Fˆ7V÷VÁBÊWÜV46ˆ÷÷ÊBÇ&6˜í"ì∞¢FWáF&VÁ&V÷˜fRÇì∞¢&WGW&‚&ˆ÷ó6RÁ&W6ˆ«fRÇì∞ß–†¶7ñÊ2gVÊ7Fñˆ‚∆ˆt7FófóGíÜ7FñˆÂGóR¬7Fñˆ‰FW67&óFñˆ‚¬WáG&“∑“í∞¢ñbÇF"«¬7W'&VÁEW6W"í&WGW&„∞¢6ˆÁ7BÊ˜&÷∆ó¶VEGóR“7G&ñÊrÜ7FñˆÂGóR«¬&¶ñˆÊR"íÁG&ñ“Çì∞¢G'í∞¢vóBF"Ê6ˆ∆∆V7Fñˆ‚Ç&7FófóGî∆ˆw2"íÊFBá∞¢W6W$ñC¢7W'&VÁEW6W"ÁVñB«¬""¿¢W6W$Ê÷S¢7W'&VÁEW6W"ÊFó7∆îÊ÷R«¬7W'&VÁEW6W"ÊV÷ñ¬«¬%WFVÁFR"¿¢W6W$V÷ñ√¢7W'&VÁEW6W"ÊV÷ñ¬«¬""¿¢W6W%&ˆ∆S¢6‰÷ÊvTFFÇíÚ&F÷ñ‚"¢&˜W&F˜&R"¿¢7FñˆÂGóS¢Ê˜&÷∆ó¶VEGóR¿¢7Fñˆ‰FW67&óFñˆ„¢7G&ñÊrÜ7Fñˆ‰FW67&óFñˆ‚«¬Ê˜&÷∆ó¶VEGóRí¿¢6ˆ÷÷W76ñC¢WáG&Ê6ˆ÷÷W76ñB«¬6V∆V7FVD6ˆ÷÷W76ñB«¬""¿¢6ˆ÷÷W76Ê÷S¢WáG&Ê6ˆ÷÷W76Ê÷R«¬6V∆V7FVD6ˆ÷÷W76Ê÷R«¬""¿¢ñ◊ñÁFÙñC¢WáG&Êñ◊ñÁFÙñB«¬""¿¢ñ◊ñÁFÙÊ÷S¢WáG&Êñ◊ñÁFÙÊ÷R«¬""¿¢FWFñ√¢WáG&ÊFWFñ¬«¬""¿¢fñWtÊ÷S¢WáG&ÁfñWtÊ÷R«¬vWD7W'&VÁEfñWtÊ÷RÇí¿¢'WGFˆ‰∆&V√¢WáG&Ê'WGFˆ‰∆&V¬«¬""¿¢FV6ÜÊñ6ƒW'&˜#¢WáG&ÁFV6ÜÊñ6ƒW'&˜"«¬WáG&ÊW'&˜"«¬""¿¢W'&˜$6ˆFS¢WáG&ÊW'&˜$6ˆFR«¬""¿¢˜76ñ&∆T6W6S¢WáG&Á˜76ñ&∆T6W6R«¬""¿¢&W6ˆ«WFñˆ‰ÜñÁC¢WáG&Á&W6ˆ«WFñˆ‰ÜñÁB«¬""¿¢fó&W7F˜&T6ˆ∆∆V7Fñˆ„¢WáG&Êfó&W7F˜&T6ˆ∆∆V7Fñˆ‚«¬""¿¢fó&W7F˜&T˜W&Fñˆ„¢WáG&Êfó&W7F˜&T˜W&Fñˆ‚«¬""¿¢fó&W7F˜&Tfñ«W&UGóS¢WáG&Êfó&W7F˜&Tfñ«W&UGóR«¬""¿¢VÁ6fVDFF¢WáG&ÁVÁ6fVDFF«¬""¿¢fó&W7F˜&U'V∆TÜñÁC¢WáG&Êfó&W7F˜&U'V∆TÜñÁB«¬""¿¢FWfñ6TñÊfÛ¢ÊfñvF˜"ÁW6W$vVÁB«¬""¿¢7&VFVDC¢fó&V&6RÊfó&W7F˜&R‰fñV∆Ef«VRÁ6W'fW%Fñ÷W7F◊Çê¢“ì∞¢“6F6ÇÜW'&˜"í∞¢6ˆÁ6ˆ∆RÁv&‚Ç%6«fFvvñÚ7FófóGî∆ˆw2Êˆ‚&óW66óFÛ¢"¬W'&˜"ì∞¢–ß–†¶gVÊ7Fñˆ‚fñ«FW$7FófUW6W'4∆ˆw2Çí∞¢6ˆÁ7BW6W%FW&““7G&ñÊráVíÊ7FófUW6W'56V&6ÖW6W#ÚÁf«VR«¬""íÁFÙ∆˜vW$66RÇì∞¢6ˆÁ7B6ˆ÷÷W76FW&““7G&ñÊráVíÊ7FófUW6W'56V&6Ñ6ˆ÷÷W76ÚÁf«VR«¬""íÁFÙ∆˜vW$66RÇì∞¢6ˆÁ7Bñ◊ñÁFıFW&““7G&ñÊráVíÊ7FófUW6W'56V&6Ññ◊ñÁFÛÚÁf«VR«¬""íÁFÙ∆˜vW$66RÇì∞¢6ˆÁ7B˜W&F˜"“VíÊ7FófUW6W'4fñ«FW$˜W&F˜#ÚÁf«VR«¬"#∞¢6ˆÁ7B7Fñˆ‚“VíÊ7FófUW6W'4fñ«FW$7Fñˆ„ÚÁf«VR«¬"#∞¢6ˆÁ7BW'&˜'4ˆÊ«í“&ˆˆ∆V‚áVíÊ7FófUW6W'4W'&˜'4ˆÊ«ìÚÊ6ÜV6∂VBì∞¢&WGW&‚7FófUW6W'4∆ˆw2Êfñ«FW"ÇÜ∆ˆrí”‚∞¢ñbáW6W%FW&“bbG∂∆ˆrÁW6W$Ê÷R«¬"'“G∂∆ˆrÁW6W$V÷ñ¬«¬"'÷ÁFÙ∆˜vW$66RÇíÊñÊ6«VFW2áW6W%FW&“íí&WGW&‚f«6S∞¢ñbÜ6ˆ÷÷W76FW&“bbG∂∆ˆrÊ6ˆ÷÷W76Ê÷R«¬"'“G∂∆ˆrÊ6ˆ÷÷W76ñB«¬"'÷ÁFÙ∆˜vW$66RÇíÊñÊ6«VFW2Ü6ˆ÷÷W76FW&“íí&WGW&‚f«6S∞¢ñbÜñ◊ñÁFıFW&“bbG∂∆ˆrÊñ◊ñÁFÙÊ÷R«¬"'“G∂∆ˆrÊñ◊ñÁFÙñB«¬"'÷ÁFÙ∆˜vW$66RÇíÊñÊ6«VFW2Üñ◊ñÁFıFW&“íí&WGW&‚f«6S∞¢ñbÜ˜W&F˜"bb7G&ñÊrÜ∆ˆrÁW6W$ñB«¬∆ˆrÁW6W$V÷ñ¬«¬""í”“˜W&F˜"í&WGW&‚f«6S∞¢ñbÜ7Fñˆ‚bb∆ˆrÊ7FñˆÂGóR”“7Fñˆ‚í&WGW&‚f«6S∞¢ñbÜW'&˜'4ˆÊ«íbbó4W'&˜$7FófóGíÜ∆ˆríí&WGW&‚f«6S∞¢&WGW&‚G'VS∞¢“ì∞ß–†¶gVÊ7Fñˆ‚'Vñ∆D7FófUW6W'5&˜w2áW6W'2“µ“¬∆ˆw2“7FófUW6W'4∆ˆw2í∞¢&WGW&‚W6W'2Ê÷áW6W"”‚∞¢6ˆÁ7BW6W$∆ˆw2“∆ˆw2Êfñ«FW"Ü¬”‚Ü¬ÁW6W$ñBbbÜ¬ÁW6W$ñB””“W6W"ÊñB«¬¬ÁW6W$ñB””“W6W"ÁVñBíí«¬Ê˜&÷∆ó¶TV÷ñ¬Ü¬ÁW6W$V÷ñ¬í””“Ê˜&÷∆ó¶TV÷ñ¬áW6W"ÊV÷ñ¬íì∞¢6ˆÁ7B∆7B“W6W$∆ˆw5≥“«¬∑”∞¢6ˆÁ7BˆÊ∆ñÊR“FFRÊÊ˜rÇí“fó&W7F˜&TFFUFÙ÷ñ∆∆ó2áW6W"Ê∆7E6VV‰Bí√“¢c¢∞¢6ˆÁ7B&ˆ&∆V◊2“W6W$∆ˆw2Êfñ«FW"Üó4W'&˜$7FófóGííÊ∆VÊwFÉ∞¢&WGW&‚∆'Fñ6∆R6∆73“&7FófR◊W6W'2◊&˜ró2÷6∆ñ6∂&∆R"&ˆ∆S“&'WGFˆ‚"F&ñÊFWÉ“#"FF÷7FófR◊W6W"÷ñC“"G∂W66TÖD‘¬áW6W"ÊñB«¬W6W"ÁVñB«¬W6W"ÊV÷ñ¬«¬""ó“"&ñ÷∆&V√“$&íGFófóL:FíG∂W66TÖD‘¬ÜvWEW6W$Fó7∆îÊ÷RáW6W"íó“#„«7G&ˆÊs‚G∂W66TÖD‘¬ÜvWEW6W$Fó7∆îÊ÷RáW6W"íó”¬˜7G&ˆÊs„∆Fób6∆73“&7FófR◊W6W'2◊&˜r÷w&ñB#„«7„‚G∂W66TÖD‘¬áW6W"ÊV÷ñ¬«¬"“"ó”¬˜7„„«7„‚G∂vWEW6W%&ˆ∆RáW6W"ó”¬˜7„„«7‚6∆73“&7FófR◊7FGW2÷F˜BG∂ˆÊ∆ñÊRÚ&ˆÊ∆ñÊR"¢&ˆff∆ñÊR'“#‚G∂ˆÊ∆ñÊRÚ/	˘˙"ˆÊ∆ñÊR"¢.)™¢ˆff∆ñÊR'”¬˜7„„«7„ÂV«Fñ÷Ú66W76Û¢G∂f˜&÷D7FófóGîFFRáW6W"Ê∆7D∆ˆvñ‰B«¬W6W"Ê7&VFVDBó”¬˜7„„«7„ÂV«Fñ÷GFófóL:¢G∂f˜&÷D7FófóGîFFRáW6W"Ê∆7E6VV‰Bó”¬˜7„„«7„ÂV«Fñ÷¶ñˆÊS¢G∂W66TÖD‘¬Ü∆7BÊ7Fñˆ‰FW67&óFñˆ‚«¬∆7BÊ7FñˆÂGóR«¬"“"ó”¬˜7„„«7„‰6ˆ÷÷W76¢G∂W66TÖD‘¬Ü∆7BÊ6ˆ÷÷W76Ê÷R«¬∆7BÊ6ˆ÷÷W76ñB«¬"“"ó”¬˜7„„«7„Â&ˆ&∆V÷ì¢G∑&ˆ&∆V◊7”¬˜7„„¬ˆFóc„¬ˆ'Fñ6∆SÊ∞¢“íÊ¶ˆñ‚Ç""í«¬s«6∆73“&◊WFVB#‰ÊW77V‚WFVÁFR„¬˜‚s∞ß–†¶gVÊ7Fñˆ‚'Vñ∆D7FófT∆ˆt∆ó7BÜ∆ˆw2“fñ«FW$7FófUW6W'4∆ˆw2Çíí∞¢6ˆÁ7B∆ˆtw&˜W2“'Vñ∆D7FófT∆ˆtw&˜W2Ü∆ˆw2ì∞¢vñÊF˜rÊ7FófUW6W'5&VÊFW&VD∆ˆtw&˜W4'î∂Wí“vñÊF˜rÊ7FófUW6W'5&VÊFW&VD∆ˆtw&˜W4'î∂Wí«¬∑”∞¢&WGW&‚∆ˆtw&˜W2Ê÷ÇÜw&˜W¬ñÊFWÇí”‚∞¢6ˆÁ7B∆ˆr“w&˜WÁ&ñ÷'í«¬∑”∞¢6ˆÁ7BGóT6∆72“ó4W'&˜$7FófóGíÜ∆ˆríÚ&W'&˜""¢Ö≤'&W76ñˆÊUˆf˜'¶"¬&7&V¶ñˆÊU˜7VG&R"¬&÷ˆFñfñ6ˆ˜&R%“ÊñÊ6«VFW2Ü∆ˆrÊ7FñˆÂGóRíÚ&ñ◊˜'FÁB"¢&Ê˜&÷¬"ì∞¢6ˆÁ7B7FGW2“vWD7FófT∆ˆu7FGW2Ü∆ˆrÊñB«¬w&˜WÊ∂Wíì∞¢6ˆÁ7B&WVFVB“w&˜WÊ∆ˆw2Ê∆VÊwFÇ‚Ú«7G&ˆÊr6∆73“&7FófR÷∆ˆr◊&WVB#‰W'&˜&R&óWGWFÚG∂w&˜WÊ∆ˆw2Ê∆VÊwFá“fˆ«FS¬˜7G&ˆÊsÊ¢"#∞¢6ˆÁ7B&˜t∂Wí“∆ˆr“G∂ñÊFWá““Gµ7G&ñÊrÜw&˜WÊ∂Wí«¬∆ˆrÊñB«¬""íÁ&W∆6RÇıµÊ◊£”ïÚ’“ˆví¬"“"ó÷∞¢vñÊF˜rÊ7FófUW6W'5&VÊFW&VD∆ˆtw&˜W4'î∂Wï∑&˜t∂Wï““w&˜W∞¢&WGW&‚∆'Fñ6∆R6∆73“&7FófR◊W6W'2÷∆ˆr◊&˜ró2÷6∆ñ6∂&∆R"&ˆ∆S“&'WGFˆ‚"F&ñÊFWÉ“#"FF÷7FófR÷∆ˆr÷∂Wì“"G∂W66TÖD‘¬á&˜t∂Wíó“#„∆Fóc„«7‚6∆73“&7FófR÷∆ˆr◊GóRG∑GóT6∆77“#‚G∑GóT6∆72””“&W'&˜""Ú/	˘KB"¢GóT6∆72””“&ñ◊˜'FÁB"Ú/	˘˙"¢/	˘KR'“G∂W66TÖD‘¬Ü∆ˆrÊ7FñˆÂGóR«¬&¶ñˆÊR"ó”¬˜7„‚(
"G∂f˜&÷D7FófóGîFFRÜ∆ˆrÊ7&VFVDBó“«7‚6∆73“&7FófR÷∆ˆr◊7FGW2#‚G∂W66TÖD‘¬á7FGW2ó”¬˜7„„¬ˆFóc„∆Fóc‚G∂W66TÖD‘¬Ü∆ˆrÁW6W$Ê÷R«¬∆ˆrÁW6W$V÷ñ¬«¬"“"ó“(
"fñWs¢G∂W66TÖD‘¬Ü∆ˆrÁfñWtÊ÷R«¬"“"ó“(
"V«6ÁFS¢G∂W66TÖD‘¬Ü∆ˆrÊ'WGFˆ‰∆&V¬«¬"“"ó”¬ˆFóc„∆Fóc‰6ˆ÷÷W76¢G∂W66TÖD‘¬Ü∆ˆrÊ6ˆ÷÷W76Ê÷R«¬∆ˆrÊ6ˆ÷÷W76ñB«¬"“"ó“(
"ñ◊ñÁFÛ¢G∂W66TÖD‘¬Ü∆ˆrÊñ◊ñÁFÙÊ÷R«¬∆ˆrÊñ◊ñÁFÙñB«¬"“"ó”¬ˆFóc„«6∆73“&◊WFVB#‚G∂W66TÖD‘¬Ü∆ˆrÊ7Fñˆ‰FW67&óFñˆ‚«¬∆ˆrÊFWFñ¬«¬"“"ó”¬˜‚G∑&WVFVG”¬ˆ'Fñ6∆SÊ∞¢“íÊ¶ˆñ‚Ç""í«¬s«6∆73“&◊WFVB#‰ÊW77VÊ¶ñˆÊRÊV¬W&ñˆFÚ„¬˜‚s∞ß–†¶gVÊ7Fñˆ‚&VÊFW$7FófUW6W'46&DFWFñ¬Ü÷WG&ñ74'î∂Wí“∑“í∞¢ñbÇVíÊ7FófUW6W'46&DFWFñ¬í&WGW&„∞¢ñbÇ6V∆V7FVD7FófUW6W'46&Bí&WGW&‚VíÊ7FófUW6W'46&DFWFñ¬Ê6∆74∆ó7BÊFBÇ&ÜñFFV‚"ì∞¢6ˆÁ7B6fr“÷WG&ñ74'î∂Wï∑6V∆V7FVD7FófUW6W'46&E”∞¢ñbÇ6frí&WGW&‚VíÊ7FófUW6W'46&DFWFñ¬Ê6∆74∆ó7BÊFBÇ&ÜñFFV‚"ì∞¢VíÊ7FófUW6W'46&DFWFñ¬Ê6∆74∆ó7BÁ&V÷˜fRÇ&ÜñFFV‚"ì∞¢VíÊ7FófUW6W'46&DFWFñ¬ÊñÊÊW$ÖD‘¬“∆Fób6∆73“'6V7Fñˆ‚÷ÜVB#„∆É#‚G∂W66TÖD‘¬Ü6frÁFóF∆Ró”¬ˆÉ#„∆'WGFˆ‚6∆73“&'F‚"GóS“&'WGFˆ‚"FF÷7FófR÷6&B÷6∆˜6S‰6ÜóVFì¬ˆ'WGFˆ„„¬ˆFóc‚G∂6frÊáF÷«÷∞ß–†¶gVÊ7Fñˆ‚&VÊFW$7FófUW6W'4FWFñ¬Çí∞¢6ˆÁ7B∆ˆw2“fñ«FW$7FófUW6W'4∆ˆw2Çì∞¢6ˆÁ7BFˆFï7F'B“ÊWrFFRÇì≤FˆFï7F'BÁ6WDÜ˜W'2É√√√ì∞¢6ˆÁ7BFˆFî∆ˆw2“7FófUW6W'4∆ˆw2Êfñ«FW"Ü¬”‚fó&W7F˜&TFFUFÙ÷ñ∆∆ó2Ü¬Ê7&VFVDBí„“FˆFï7F'BÊvWEFñ÷RÇíì∞¢6ˆÁ7BW6W'5FˆFîñG2“ÊWr6WBáFˆFî∆ˆw2Ê÷Ü¬”‚¬ÁW6W$ñB«¬Ê˜&÷∆ó¶TV÷ñ¬Ü¬ÁW6W$V÷ñ¬ííÊfñ«FW"Ñ&ˆˆ∆V‚íì∞¢6ˆÁ7BˆÊ∆ñÊUW6W'2“∆Ff˜&’W6W'2Êfñ«FW"ÇáW6W"í”‚FFRÊÊ˜rÇí“fó&W7F˜&TFFUFÙ÷ñ∆∆ó2áW6W"Ê∆7E6VV‰Bí√“¢c¢ì∞¢6ˆÁ7BFˆFïW6W'2“∆Ff˜&’W6W'2Êfñ«FW"áR”‚W6W'5FˆFîñG2ÊÜ2áRÊñBí«¬W6W'5FˆFîñG2ÊÜ2áRÁVñBí«¬W6W'5FˆFîñG2ÊÜ2ÜÊ˜&÷∆ó¶TV÷ñ¬áRÊV÷ñ¬ííì∞¢6ˆÁ7B'ïGóR“áGóRí”‚∆ˆw2Êfñ«FW"Ü¬”‚¬Ê7FñˆÂGóR””“GóRì∞¢6ˆÁ7BW'&˜$∆ˆw2“∆ˆw2Êfñ«FW"Üó4W'&˜$7FófóGíì∞¢vñÊF˜rÊ7FófUW6W'5&VÊFW&VD∆ˆtw&˜W4'î∂Wí“∑”∞¢6ˆÁ7B÷WG&ñ72“∞¢≤∂Wì¢&ˆÊ∆ñÊR"¬∆&V√¢/	˘˙"WFVÁFíˆÊ∆ñÊRFW76Ú"¬f«VS¢ˆÊ∆ñÊUW6W'2Ê∆VÊwFÇ“¿¢≤∂Wì¢'FˆFí"¬∆&V√¢/	˘RWFVÁFíˆvví"¬f«VS¢W6W'5FˆFîñG2Á6ó¶R“¿¢≤∂Wì¢&7FñˆÁ2"¬∆&V√¢/	˘8¢F˜F∆R¶ñˆÊíˆvví"¬f«VS¢FˆFî∆ˆw2Ê∆VÊwFÇ“¿¢≤∂Wì¢&FˆÊR"¬∆&V√¢.)»Rñ◊ñÁFí6ˆ◊∆WFFí"¬f«VS¢'ïGóRÇ'&W76ñˆÊUˆfGFÚ"íÊ∆VÊwFÇ“¿¢≤∂Wì¢&Ü˜W'2"¬∆&V√¢.(˚˜&RñÁ6W&óFR"¬f«VS¢'ïGóRÇ&ñÁ6W&ñ÷VÁFıˆ˜&R"íÊ∆VÊwFÇ“¿¢≤∂Wì¢&Êb"¬∆&V√¢/	˙z“Êfñv¶ñˆÊígfñFR"¬f«VS¢'ïGóRÇ'&W76ñˆÊUˆÊfñv"íÊ∆VÊwFÇ“¿¢≤∂Wì¢'7ñÊ2"¬∆&V√¢/	˘HBV«Fñ÷6ñÊ7&ˆÊóß¶¶ñˆÊR"¬f«VS¢ÊWrFFRÇíÁFÙ∆ˆ6∆UFñ÷U7G&ñÊrÇ&óB‘ïB"í“¿¢≤∂Wì¢&W'&˜'2"¬∆&V√¢.)™˚àÚWfVÁGV∆íW'&˜&í"¬f«VS¢W'&˜$∆ˆw2Ê∆VÊwFÇ–¢”∞¢ñbáVíÊ7FófUW6W'5F˜7V÷÷'ííVíÊ7FófUW6W'5F˜7V÷÷'íÊñÊÊW$ÖD‘¬“÷WG&ñ72Á6∆ñ6RÉ√"íÊ÷Ü“”‚∆'WGFˆ‚6∆73“&6&B7FófR◊W6W'2÷÷WG&ñ27FófR◊W6W'2÷÷WG&ñ2÷'WGFˆ‚"GóS“&'WGFˆ‚"FF÷7FófR÷6&C“"G∂“Ê∂Wó“#„«7G&ˆÊs‚G∂W66TÖD‘¬Ü“Ê∆&V¬ó”¬˜7G&ˆÊs„«‚G∂W66TÖD‘¬Ö7G&ñÊrÜ“Áf«VRíó”¬˜„¬ˆ'WGFˆ„ÊíÊ¶ˆñ‚Ç""ì∞¢VíÊ7FófUW6W'4F6Ü&ˆ&BÊñÊÊW$ÖD‘¬“÷WG&ñ72Ê÷Ü“”‚∆'WGFˆ‚6∆73“&6&B7FófR◊W6W'2÷÷WG&ñ27FófR◊W6W'2÷÷WG&ñ2÷'WGFˆ‚"GóS“&'WGFˆ‚"FF÷7FófR÷6&C“"G∂“Ê∂Wó“#„«7G&ˆÊs‚G∂W66TÖD‘¬Ü“Ê∆&V¬ó”¬˜7G&ˆÊs„«‚G∂W66TÖD‘¬Ö7G&ñÊrÜ“Áf«VRíó”¬˜„¬ˆ'WGFˆ„ÊíÊ¶ˆñ‚Ç""ì∞¢6ˆÁ7B÷WG&ñ74'î∂Wí“∞¢ˆÊ∆ñÊS¢≤FóF∆S¢%WFVÁFí6ˆ∆∆VvFí˜&"¬áF÷√¢'Vñ∆D7FófUW6W'5&˜w2ÜˆÊ∆ñÊUW6W'2¬∆ˆw2í“¿¢FˆFì¢≤FóF∆S¢%WFVÁFí6ÜRÜÊÊÚW6FÚ¬vˆvví"¬áF÷√¢'Vñ∆D7FófUW6W'5&˜w2áFˆFïW6W'2¬∆ˆw2í“¿¢7FñˆÁ3¢≤FóF∆S¢$7&ˆÊˆ∆ˆvñ6ˆ◊∆WF¶ñˆÊí"¬áF÷√¢∆Fób6∆73“&7FófR◊W6W'2÷∆ˆr÷∆ó7B#‚G∂'Vñ∆D7FófT∆ˆt∆ó7BáFˆFî∆ˆw2ó”¬ˆFócÊ“¿¢FˆÊS¢≤FóF∆S¢$ñ◊ñÁFí6ˆ◊∆WFFí"¬áF÷√¢∆Fób6∆73“&7FófR◊W6W'2÷∆ˆr÷∆ó7B#‚G∂'Vñ∆D7FófT∆ˆt∆ó7BÜ'ïGóRÇ'&W76ñˆÊUˆfGFÚ"íó”¬ˆFócÊ“¿¢Ü˜W'3¢≤FóF∆S¢$˜&RñÁ6W&óFR"¬áF÷√¢∆Fób6∆73“&7FófR◊W6W'2÷∆ˆr÷∆ó7B#‚G∂'Vñ∆D7FófT∆ˆt∆ó7BÜ'ïGóRÇ&ñÁ6W&ñ÷VÁFıˆ˜&R"íó”¬ˆFócÊ“¿¢Êc¢≤FóF∆S¢$Êfñv¶ñˆÊígfñFR"¬áF÷√¢∆Fób6∆73“&7FófR◊W6W'2÷∆ˆr÷∆ó7B#‚G∂'Vñ∆D7FófT∆ˆt∆ó7BÜ'ïGóRÇ'&W76ñˆÊUˆÊfñv"íó”¬ˆFócÊ“¿¢7ñÊ3¢≤FóF∆S¢%6ñÊ7&ˆÊóß¶¶ñˆÊí&óW66óFRRFó7˜6óFófí"¬áF÷√¢∆Fób6∆73“&7FófR◊W6W'2÷∆ˆr÷∆ó7B#‚G∂'Vñ∆D7FófT∆ˆt∆ó7BÜ∆ˆw2Êfñ«FW"Ü¬”‚˜7ñÊ7«6ñÊ7&ˆ‚ˆíÁFW7BÜG∂¬Ê7FñˆÂGóR«¬"'“G∂¬Ê7Fñˆ‰FW67&óFñˆ‚«¬"'÷ííó”¬ˆFócÊ“¿¢W'&˜'3¢≤FóF∆S¢$∆ó7F&ˆ&∆V÷í"¬áF÷√¢∆Fób6∆73“&7FófR◊W6W'2÷∆ˆr÷∆ó7B#‚G∂'Vñ∆D7FófT∆ˆt∆ó7BÜW'&˜$∆ˆw2ó”¬ˆFócÊ–¢”∞¢&VÊFW$7FófUW6W'46&DFWFñ¬Ü÷WG&ñ74'î∂Wíì∞¢ñbáVíÊ7FófUW6W'4gV∆ƒ∆ó7Bí∞¢VíÊ7FófUW6W'4gV∆ƒ∆ó7BÊ6∆74∆ó7BÁFˆvv∆RÇ&ÜñFFV‚"¬7FófUW6W'4gV∆ƒ∆ó7D˜V‚ì∞¢VíÊ7FófUW6W'4gV∆ƒ∆ó7BÊñÊÊW$ÖD‘¬“7FófUW6W'4gV∆ƒ∆ó7D˜V‚Ú'Vñ∆D7FófUW6W'5&˜w2á∆Ff˜&’W6W'2¬∆ˆw2í¢"#∞¢–¢ñbáVíÊ7FófUW6W'4∆ˆt∆ó7Bí∞¢VíÊ7FófUW6W'4∆ˆt∆ó7BÊ6∆74∆ó7BÁFˆvv∆RÇ&ÜñFFV‚"¬7FófUW6W'4∆ˆt∆ó7D˜V‚ì∞¢VíÊ7FófUW6W'4∆ˆt∆ó7BÊñÊÊW$ÖD‘¬“7FófUW6W'4∆ˆt∆ó7D˜V‚Ú'Vñ∆D7FófT∆ˆt∆ó7BÜ∆ˆw2í¢"#∞¢VíÊ7FófUW6W'4∆ˆt∆ó7BÊFF6WBÊw&˜W2“'&VÊFW&VB#∞¢–¢VíÊ7FófUW6W'4fñ«FW%ÊV√ÚÊ6∆74∆ó7BÁFˆvv∆RÇ&ÜñFFV‚"¬7FófUW6W'4fñ«FW$˜V‚ì∞¢VíÊ7FófUW6W'4fñ«FW%Fˆvv∆SÚÁ6WDGG&ñ'WFRÇ&&ñ÷WáÊFVB"¬7G&ñÊrÜ7FófUW6W'4fñ«FW$˜V‚íì∞¢VíÊ7FófUW6W'4gV∆≈Fˆvv∆SÚÁ6WDGG&ñ'WFRÇ&&ñ÷WáÊFVB"¬7G&ñÊrÜ7FófUW6W'4gV∆ƒ∆ó7D˜V‚íì∞¢VíÊ7FófUW6W'4∆ˆuFˆvv∆SÚÁ6WDGG&ñ'WFRÇ&&ñ÷WáÊFVB"¬7G&ñÊrÜ7FófUW6W'4∆ˆt∆ó7D˜V‚íì∞¢&VÊFW%6V∆V7FVD7FófUW6W$FWFñ¬Çì∞ß–††¶gVÊ7Fñˆ‚˜V‰7FófT∆ˆu&ˆ&∆V‘FWFñ¬Üw&˜Wí∞¢6ˆÁ7B∆ˆr“w&˜WÚÁ&ñ÷'í«¬∑”∞¢ñbÇ∆ˆrí&WGW&„∞¢6ˆÁ7B÷ˆF¬“Fˆ7V÷VÁBÊ7&VFTV∆V÷VÁBÇ&Fób"ì∞¢÷ˆF¬Ê6∆74Ê÷R“&7FófR÷∆ˆr÷FWFñ¬÷÷ˆF¬#∞¢6ˆÁ7Bó4fó&W7F˜&R“7G&ñÊrÜ∆ˆrÊ7FñˆÂGóR«¬""íÊñÊ6«VFW2Ç&fó&W7F˜&R"ì∞¢6ˆÁ7BFWFñ≈FWáB“7FófT∆ˆu7V÷÷'ïFWáBÜw&˜Wì∞¢6ˆÁ7B7FGW4˜FñˆÁ2“5DïdUÙƒÙuı5DEU5ÙıDîÙÂ2Ê÷á7FGW2”‚∆˜Fñˆ‚f«VS“"G∂W66TÖD‘¬á7FGW2ó“"G∂vWD7FófT∆ˆu7FGW2Ü∆ˆrÊñB«¬w&˜WÊ∂Wíí””“7FGW2Ú'6V∆V7FVB"¢"'”‚G∂W66TÖD‘¬á7FGW2ó”¬ˆ˜Fñˆ„ÊíÊ¶ˆñ‚Ç""ì∞¢6ˆÁ7B&WVG2“w&˜WÊ∆ˆw2Ê∆VÊwFÇ‚Ú«6V7Fñˆ„„∆ÉC‰W'&˜&R&óWGWFÚG∂w&˜WÊ∆ˆw2Ê∆VÊwFá“fˆ«FS¬ˆÉC„«V√‚G∂w&˜WÊ∆ˆw2Ê÷Ü¬”‚∆∆ì‚G∂W66TÖD‘¬Ü¬ÁW6W$Ê÷R«¬¬ÁW6W$V÷ñ¬«¬"“"ó“(
"G∂f˜&÷D7FófóGîFFRÜ¬Ê7&VFVDBó“(
"G∂W66TÖD‘¬Ü¬ÁfñWtÊ÷R«¬"“"ó”¬ˆ∆ìÊíÊ¶ˆñ‚Ç""ó”¬˜V√„¬˜6V7Fñˆ„Ê¢"#∞¢÷ˆF¬ÊñÊÊW$ÖD‘¬“∆Fób6∆73“&7FófR÷∆ˆr÷FWFñ¬÷6&B"&ˆ∆S“&Fñ∆ˆr"&ñ÷÷ˆF√“'G'VR"&ñ÷∆&V√“$FWGFv∆ñÚ&ˆ&∆V÷Ú¶ñˆÊR#„∆Fób6∆73“'6V7Fñˆ‚÷ÜVB#„∆É#‰FWGFv∆ñÚ&ˆ&∆V÷Ú¶ñˆÊS¬ˆÉ#„∆'WGFˆ‚GóS“&'WGFˆ‚"6∆73“&'F‚"FF÷7FófR÷∆ˆr÷6∆˜6S‰6ÜóVFì¬ˆ'WGFˆ„„¬ˆFóc„∆Fób6∆73“&7FófR÷∆ˆr÷FWFñ¬÷w&ñB#„«6V7Fñˆ„„∆ÉC‰6˜6:Ç7V66W76Û¬ˆÉC„«‚G∂W66TÖD‘¬Ü∆ˆrÊ7Fñˆ‰FW67&óFñˆ‚«¬∆ˆrÊFWFñ¬«¬$WfVÁFÚ&Vvó7G&FÚ"ó”¬˜„¬˜6V7Fñˆ„„«6V7Fñˆ„„∆ÉCÂWFVÁFS¬ˆÉC„«‚G∂W66TÖD‘¬Ü∆ˆrÁW6W$Ê÷R«¬∆ˆrÁW6W$V÷ñ¬«¬"“"ó”¬˜„¬˜6V7Fñˆ„„«6V7Fñˆ„„∆ÉCÂVÊFÛ¬ˆÉC„«‚G∂f˜&÷D7FófóGîFFRÜ∆ˆrÊ7&VFVDBó”¬˜„¬˜6V7Fñˆ„„«6V7Fñˆ„„∆ÉCÂfñWrÚvñÊ¬ˆÉC„«‚G∂W66TÖD‘¬Ü∆ˆrÁfñWtÊ÷R«¬$FfW&ñfñ6&R"ó”¬˜„¬˜6V7Fñˆ„„«6V7Fñˆ„„∆ÉCÂV«6ÁFR&V◊WFÛ¬ˆÉC„«‚G∂W66TÖD‘¬Ü∆ˆrÊ'WGFˆ‰∆&V¬«¬$Êˆ‚ñÊFñ6FÚ"ó”¬˜„¬˜6V7Fñˆ„„«6V7Fñˆ„„∆ÉC‰6ˆ÷÷W76Úñ◊ñÁFÛ¬ˆÉC„«‚G∂W66TÖD‘¬Ü∆ˆrÊ6ˆ÷÷W76Ê÷R«¬∆ˆrÊ6ˆ÷÷W76ñB«¬"“"ó“ÚG∂W66TÖD‘¬Ü∆ˆrÊñ◊ñÁFÙÊ÷R«¬∆ˆrÊñ◊ñÁFÙñB«¬"“"ó”¬˜„¬˜6V7Fñˆ„„«6V7Fñˆ‚6∆73“'vñFR#„∆ÉC‰W'&˜&RFV6Êñ6Ú6ˆ◊∆WFÛ¬ˆÉC„«&S‚G∂W66TÖD‘¬Ü∆ˆrÁFV6ÜÊñ6ƒW'&˜"«¬∆ˆrÊFWFñ¬«¬"“"ó”¬˜&S„¬˜6V7Fñˆ„„«6V7Fñˆ„„∆ÉCÂ˜76ñ&ñ∆R6W6¬ˆÉC„«‚G∂W66TÖD‘¬Ü∆ˆrÁ˜76ñ&∆T6W6R«¬Üó4W'&˜$7FófóGíÜ∆ˆríÚ$FÊ∆óß¶&Rñ‚&6RW'&˜&RFV6Êñ6Ú¬W&÷W76í¬FFíRfñWr‚"¢"“"íó”¬˜„¬˜6V7Fñˆ„„«6V7Fñˆ„„∆ÉCÂ7VvvW&ñ÷VÁFÚW"&ó6ˆ«fW&S¬ˆÉC„«‚G∂W66TÖD‘¬Ü∆ˆrÁ&W6ˆ«WFñˆ‰ÜñÁB«¬Üó4W'&˜$7FófóGíÜ∆ˆríÚ%&ó&ˆGW'&Rñ¬f«W76Ú¬6ˆÁG&ˆ∆∆&RW&÷W76íˆFFíR6ˆÁ6ˆ∆R¬ˆí6˜'&VvvW&Rñ¬6ˆFñ6RÚ∆R&Vvˆ∆R‚"¢"“"íó”¬˜„¬˜6V7Fñˆ„‚G∂ó4fó&W7F˜&RÚ«6V7Fñˆ‚6∆73“'vñFRfó&W7F˜&R÷FV'Vr#„∆ÉC‰FWGFv∆ífó&W7F˜&S¬ˆÉC„«„«7G&ˆÊs‰6ˆ∆∆V7Fñˆ„£¬˜7G&ˆÊs‚G∂W66TÖD‘¬Ü∆ˆrÊfó&W7F˜&T6ˆ∆∆V7Fñˆ‚«¬$FfW&ñfñ6&R"ó”¬˜„«„«7G&ˆÊs‰˜W&¶ñˆÊRFVÁFF£¬˜7G&ˆÊs‚G∂W66TÖD‘¬Ü∆ˆrÊfó&W7F˜&T˜W&Fñˆ‚«¬$FfW&ñfñ6&R"ó”¬˜„«„«7G&ˆÊsÂW&÷W76ÚÊVvFÚÚ&WFRÚFFÚ÷Ê6ÁFS£¬˜7G&ˆÊs‚G∂W66TÖD‘¬Ü∆ˆrÊfó&W7F˜&Tfñ«W&UGóR«¬$FfW&ñfñ6&R"ó”¬˜„«„«7G&ˆÊs‰FFíÊˆ‚6«fFì£¬˜7G&ˆÊs‚G∂W66TÖD‘¬Ü∆ˆrÁVÁ6fVDFF«¬$FfW&ñfñ6&R"ó”¬˜„«„«7G&ˆÊsÂ&Vvˆ∆fó&W7F˜&R6ÜR˜G&V&&R&∆ˆ66&S£¬˜7G&ˆÊs‚G∂W66TÖD‘¬Ü∆ˆrÊfó&W7F˜&U'V∆TÜñÁB«¬$6ˆÁG&ˆ∆∆&R∆∆˜r&VB˜w&óFRFV∆∆6ˆ∆∆V7Fñˆ‚ñÁFW&W76F‚"ó”¬˜„¬˜6V7Fñˆ„Ê¢"'“G∑&WVG7”¬ˆFóc„∆∆&V¬6∆73“&7FófR÷∆ˆr◊7FGW2÷VFóF˜"#Â7FFÚ&ˆ&∆V÷«6V∆V7BFF÷7FófR÷∆ˆr◊7FGW3‚G∑7FGW4˜FñˆÁ7”¬˜6V∆V7C„¬ˆ∆&V√„∆Fób6∆73“&óFV“÷7FñˆÁ2#„∆'WGFˆ‚GóS“&'WGFˆ‚"6∆73“&'F‚"FF÷6˜í÷FWFñ«3‰6˜ñFWGFv∆ì¬ˆ'WGFˆ„„∆'WGFˆ‚GóS“&'WGFˆ‚"6∆73“&'F‚'F‚◊&ñ÷'í"FF÷6˜í÷6ˆFWÉ‚G∂ó4fó&W7F˜&RÚ$6˜ñW'&˜&RW"6ˆFWÇ"¢$6˜ñW"6ˆFWÇ'”¬ˆ'WGFˆ„„∆'WGFˆ‚GóS“&'WGFˆ‚"6∆73“&'F‚"FF÷÷&≤◊&W6ˆ«fVCÂ6VvÊ&ó6ˆ«FÛ¬ˆ'WGFˆ„„∆'WGFˆ‚GóS“&'WGFˆ‚"6∆73“&'F‚"FF÷˜V‚÷6ˆ÷÷W76G∂∆ˆrÊ6ˆ÷÷W76ñBÚ""¢&Fó6&∆VB'”‰&í6ˆ÷÷W76¬ˆ'WGFˆ„„∆'WGFˆ‚GóS“&'WGFˆ‚"6∆73“&'F‚"FF÷˜V‚÷ñ◊ñÁFÚG∂∆ˆrÊ6ˆ÷÷W76ñBbbÜ∆ˆrÊñ◊ñÁFÙñB«¬∆ˆrÊñ◊ñÁFÙÊ÷RíÚ""¢&Fó6&∆VB'”‰&íñ◊ñÁFÛ¬ˆ'WGFˆ„„¬ˆFóc„¬ˆFócÊ∞¢Fˆ7V÷VÁBÊ&ˆGíÊVÊD6Üñ∆BÜ÷ˆF¬ì∞¢6ˆÁ7B6∆˜6R“Çí”‚≤÷ˆF¬Á&V÷˜fRÇì≤&VÊFW$7FófUW6W'4FWFñ¬Çì≤”∞¢÷ˆF¬ÁVW'ï6V∆V7F˜"Ç%∂FF÷7FófR÷∆ˆr÷6∆˜6U“"ìÚÊFDWfVÁD∆ó7FVÊW"Ç&6∆ñ6≤"¬6∆˜6Rì∞¢÷ˆF¬ÊFDWfVÁD∆ó7FVÊW"Ç&6∆ñ6≤"¬ÜWfVÁBí”‚≤ñbÜWfVÁBÁF&vWB””“÷ˆF¬í6∆˜6RÇì≤“ì∞¢÷ˆF¬ÁVW'ï6V∆V7F˜"Ç%∂FF÷7FófR÷∆ˆr◊7FGW5“"ìÚÊFDWfVÁD∆ó7FVÊW"Ç&6ÜÊvR"¬ÜWfVÁBí”‚6WD7FófT∆ˆu7FGW2Ü∆ˆrÊñB«¬w&˜WÊ∂Wí¬WfVÁBÁF&vWBÁf«VRíì∞¢÷ˆF¬ÁVW'ï6V∆V7F˜"Ç%∂FF÷6˜í÷FWFñ«5“"ìÚÊFDWfVÁD∆ó7FVÊW"Ç&6∆ñ6≤"¬Çí”‚6˜ïFWáEFÙ6∆ó&ˆ&BÜFWFñ≈FWáBíì∞¢÷ˆF¬ÁVW'ï6V∆V7F˜"Ç%∂FF÷6˜í÷6ˆFWÖ“"ìÚÊFDWfVÁD∆ó7FVÊW"Ç&6∆ñ6≤"¬Çí”‚6˜ïFWáEFÙ6∆ó&ˆ&BÜÊ∆óß¶R6˜'&VvvíVW7FÚ&ˆ&∆V÷FV∆¬vÜW&•∆Â∆‚G∂FWFñ≈FWáG÷íì∞¢÷ˆF¬ÁVW'ï6V∆V7F˜"Ç%∂FF÷÷&≤◊&W6ˆ«fVE“"ìÚÊFDWfVÁD∆ó7FVÊW"Ç&6∆ñ6≤"¬Çí”‚≤6WD7FófT∆ˆu7FGW2Ü∆ˆrÊñB«¬w&˜WÊ∂Wí¬%&ó6ˆ«FÚ"ì≤6∆˜6RÇì≤“ì∞¢÷ˆF¬ÁVW'ï6V∆V7F˜"Ç%∂FF÷˜V‚÷6ˆ÷÷W76“"ìÚÊFDWfVÁD∆ó7FVÊW"Ç&6∆ñ6≤"¬Çí”‚≤ñbÜ∆ˆrÊ6ˆ÷÷W76ñBí≤vñÊF˜rÊ∆ˆ6Fñˆ‚ÊÜ6Ç“6ˆ÷÷W76“G∂VÊ6ˆFUU$î6ˆ◊ˆÊVÁBÜ∆ˆrÊ6ˆ÷÷W76ñBó÷≤6∆˜6RÇì≤““ì∞¢÷ˆF¬ÁVW'ï6V∆V7F˜"Ç%∂FF÷˜V‚÷ñ◊ñÁFı“"ìÚÊFDWfVÁD∆ó7FVÊW"Ç&6∆ñ6≤"¬Çí”‚≤ñbÜ∆ˆrÊ6ˆ÷÷W76ñBí≤6ˆÁ7Bñ◊ñÁFÚ“∆ˆrÊñ◊ñÁFÙñB«¬∆ˆrÊñ◊ñÁFÙÊ÷R«¬"#≤vñÊF˜rÊ∆ˆ6Fñˆ‚ÊÜ6Ç“6ˆ÷÷W76“G∂VÊ6ˆFUU$î6ˆ◊ˆÊVÁBÜ∆ˆrÊ6ˆ÷÷W76ñBó“G∂ñ◊ñÁFÚÚfñ◊ñÁFÛ“G∂VÊ6ˆFUU$î6ˆ◊ˆÊVÁBÜñ◊ñÁFÚó÷¢"'÷≤6∆˜6RÇì≤““ì∞ß–†¶gVÊ7Fñˆ‚&VÊFW%6V∆V7FVD7FófUW6W$FWFñ¬Çí∞¢ñbÇVíÊ7FófUW6W'5W6W$FWFñ¬«¬6V∆V7FVD7FófUW6W'5W6W$ñBí&WGW&‚VíÊ7FófUW6W'5W6W$FWFñ√ÚÊ6∆74∆ó7BÊFBÇ&ÜñFFV‚"ì∞¢6ˆÁ7BW6W"“∆Ff˜&’W6W'2ÊfñÊBáR”‚7G&ñÊráRÊñB«¬RÁVñB«¬RÊV÷ñ¬í””“6V∆V7FVD7FófUW6W'5W6W$ñBì∞¢ñbÇW6W"í&WGW&„∞¢6ˆÁ7BW6W$∆ˆw2“7FófUW6W'4∆ˆw2Êfñ«FW"Ü¬”‚Ü¬ÁW6W$ñBbb¬ÁW6W$ñB””“W6W"ÊñBí«¬Ê˜&÷∆ó¶TV÷ñ¬Ü¬ÁW6W$V÷ñ¬í””“Ê˜&÷∆ó¶TV÷ñ¬áW6W"ÊV÷ñ¬íì∞¢6ˆÁ7B6˜VÁB“áGóRí”‚W6W$∆ˆw2Êfñ«FW"Ü¬”‚¬Ê7FñˆÂGóR””“GóRíÊ∆VÊwFÉ∞¢VíÊ7FófUW6W'5W6W$FWFñ¬Ê6∆74∆ó7BÁ&V÷˜fRÇ&ÜñFFV‚"ì∞¢VíÊ7FófUW6W'5W6W$FWFñ¬ÊñÊÊW$ÖD‘¬“∆Fób6∆73“'6V7Fñˆ‚÷ÜVB#„∆É#‰FWGFv∆ñÚ6ñÊvˆ∆ÚWFVÁFS¬ˆÉ#„«7‚6∆73“'ñ∆¬#‚G∂W66TÖD‘¬ÜvWEW6W%&ˆ∆RáW6W"íó”¬˜7„„¬ˆFóc„∆Fób6∆73“&7FófR◊W6W'2◊&˜r÷w&ñB#„«7„‰Êˆ÷S¢G∂W66TÖD‘¬ÜvWEW6W$Fó7∆îÊ÷RáW6W"íó”¬˜7„„«7„‰V÷ñ√¢G∂W66TÖD‘¬áW6W"ÊV÷ñ¬«¬"“"ó”¬˜7„„«7„Â7FFÛ¢G¥FFRÊÊ˜rÇí“fó&W7F˜&TFFUFÙ÷ñ∆∆ó2áW6W"Ê∆7E6VV‰Bí√“£c£Ú/	˘˙"ˆÊ∆ñÊR"¢.)™¢ˆff∆ñÊR'”¬˜7„„«7„ÂV«Fñ÷Ú66W76Û¢G∂f˜&÷D7FófóGîFFRáW6W"Ê∆7D∆ˆvñ‰B«¬W6W"Ê7&VFVDBó”¬˜7„„«7„‰˜&RñÁ6W&óFS¢G∂6˜VÁBÇ&ñÁ6W&ñ÷VÁFıˆ˜&R"ó”¬˜7„„«7„‰ñ◊ñÁFídEDÛ¢G∂6˜VÁBÇ'&W76ñˆÊUˆfGFÚ"ó”¬˜7„„«7„‰dı%§¢G∂6˜VÁBÇ'&W76ñˆÊUˆf˜'¶"ó”¬˜7„„«7„‰‰dît¢G∂6˜VÁBÇ'&W76ñˆÊUˆÊfñv"ó”¬˜7„„«7„‰6ˆ÷÷W76R6ˆÁ7V«FFS¢G∂ÊWr6WBáW6W$∆ˆw2Ê÷Ü¬”‚¬Ê6ˆ÷÷W76ñB«¬¬Ê6ˆ÷÷W76Ê÷RíÊfñ«FW"Ñ&ˆˆ∆V‚ííÁ6ó¶W”¬˜7„„«7„‰W'&˜&ì¢G∑W6W$∆ˆw2Êfñ«FW"Üó4W'&˜$7FófóGííÊ∆VÊwFá”¬˜7„„¬ˆFóc„∆É3‰7&ˆÊˆ∆ˆvñGFófóL:¬ˆÉ3‚G∑W6W$∆ˆw2Á6∆ñ6RÉ√3íÊ÷Ü¬”‚«6∆73“&◊WFVB#‚G∂f˜&÷D7FófóGîFFRÜ¬Ê7&VFVDBó“(	BG∂W66TÖD‘¬Ü¬Ê7Fñˆ‰FW67&óFñˆ‚«¬¬Ê7FñˆÂGóR«¬"“"ó”¬˜ÊíÊ¶ˆñ‚Ç""í«¬s«6∆73“&◊WFVB#‰ÊW77VÊGFófóL:„¬˜‚w÷∞ß–†¶7ñÊ2gVÊ7Fñˆ‚∆ˆD7FófUW6W'4∆ˆw2Çí∞¢ñbÇ6‰÷ÊvTFFÇíí&WGW&„∞¢6ˆÁ7B≤7F'B¬VÊB““vWD7FófóGïW&ñˆE&ÊvRÇì∞¢6ˆÁ7B6Ê“vóBF"Ê6ˆ∆∆V7Fñˆ‚Ç&7FófóGî∆ˆw2"íÁvÜW&RÇ&7&VFVDB"¬#„“"¬7F'BíÁvÜW&RÇ&7&VFVDB"¬#√“"¬VÊBíÊ˜&FW$'íÇ&7&VFVDB"¬&FW62"íÊ∆ñ÷óBÉSíÊvWBÇì∞¢7FófUW6W'4∆ˆw2“6ÊÊFˆ72Ê÷ÜFˆ2”‚á≤ñC¢Fˆ2ÊñB¬‚‚ÊFˆ2ÊFFÇí“íì∞¢6ˆÁ7B˜W&F˜'2“ÊWr÷Çì≤6ˆÁ7B7FñˆÁ2“ÊWr6WBÇì∞¢7FófUW6W'4∆ˆw2Êf˜$V6ÇÜ∆ˆr”‚≤˜W&F˜'2Á6WBÜ∆ˆrÁW6W$ñB«¬∆ˆrÁW6W$V÷ñ¬«¬""¬∆ˆrÁW6W$Ê÷R«¬∆ˆrÁW6W$V÷ñ¬«¬$˜W&F˜&R"ì≤ñbÜ∆ˆrÊ7FñˆÂGóRí7FñˆÁ2ÊFBÜ∆ˆrÊ7FñˆÂGóRì≤“ì∞¢ñbáVíÊ7FófUW6W'4fñ«FW$˜W&F˜"íVíÊ7FófUW6W'4fñ«FW$˜W&F˜"ÊñÊÊW$ÖD‘¬“s∆˜Fñˆ‚f«VS“"#ÂGWGFí˜W&F˜&ì¬ˆ˜Fñˆ„‚r≤'&íÊg&ˆ“Ü˜W&F˜'2¬Ö∂ñB∆Ê÷U“í”‚∆˜Fñˆ‚f«VS“"G∂W66TÖD‘¬ÜñBó“#‚G∂W66TÖD‘¬ÜÊ÷Ró”¬ˆ˜Fñˆ„ÊíÊ¶ˆñ‚Ç""ì∞¢ñbáVíÊ7FófUW6W'4fñ«FW$7Fñˆ‚íVíÊ7FófUW6W'4fñ«FW$7Fñˆ‚ÊñÊÊW$ÖD‘¬“s∆˜Fñˆ‚f«VS“"#ÂGWGFR¶ñˆÊì¬ˆ˜Fñˆ„‚r≤'&íÊg&ˆ“Ü7FñˆÁ2íÁ6˜'BÇíÊ÷Ü”‚∆˜Fñˆ‚f«VS“"G∂W66TÖD‘¬Üó“#‚G∂W66TÖD‘¬Üó”¬ˆ˜Fñˆ„ÊíÊ¶ˆñ‚Ç""ì∞¢&VÊFW$7FófUW6W'4FWFñ¬Çì∞ß–†¶7ñÊ2gVÊ7Fñˆ‚˜V‰7FófUW6W'4FWFñ≈fñWrÇí∞¢ñbÇVíÊ7FófUW6W'4FWFñ≈vRí&WGW&„∞¢6ˆÁ7B∆∆˜vVB“6‰÷ÊvTFFÇì∞¢VíÊ7FófUW6W'466W74÷W76vRÁFWáD6ˆÁFVÁB“∆∆˜vVBÚ$6ˆÁ6ˆ∆RFí6ˆÁG&ˆ∆∆ÚWFñ∆óß¶Ú¬WFVÁFíRGFófóL:‚"¢$66W76Ú6ˆÁ6VÁFóFÚ6ˆ∆Úv∆í÷÷ñÊó7G&F˜&í‚#∞¢VíÊ7FófUW6W'4F÷ñ‰6ˆÁ6ˆ∆SÚÊ6∆74∆ó7BÁFˆvv∆RÇ&ÜñFFV‚"¬∆∆˜vVBì∞¢ñbÇ∆∆˜vVBí&WGW&„∞¢ñbÜ7FófUW6W'4∆ˆFVBí&WGW&„∞¢7FófUW6W'4∆ˆFVB“G'VS∞¢VíÊ7FófUW6W'4∆ˆt∆ó7BÊñÊÊW$ÖD‘¬“s«6∆73“&◊WFVB#‰6&ñ6÷VÁFÚGFófóL:‚‚„¬˜‚s∞¢G'í≤vóB∆ˆD7FófUW6W'4∆ˆw2Çì≤–¢6F6ÇÜW'&˜"í≤VíÊ7FófUW6W'4∆ˆt∆ó7BÊñÊÊW$ÖD‘¬“s«6∆73“&◊WFVB#‰W'&˜&R6&ñ6÷VÁFÚ&Vvó7G&Ú¶ñˆÊí„¬˜‚s≤6ˆÁ6ˆ∆RÊW'&˜"Ç$W'&˜&R7FófóGî∆ˆw3¢"¬W'&˜"ì≤–ß–†¶gVÊ7Fñˆ‚∆ˆ6ƒFFUf«VRÜFFR“ÊWrFFRÇíí∞¢6ˆÁ7Bˆfg6WB“FFRÊvWEFñ÷W¶ˆÊTˆfg6WBÇí¢c∞¢&WGW&‚ÊWrFFRÜFFRÊvWEFñ÷RÇí“ˆfg6WBíÁFÙï4ı7G&ñÊrÇíÁ6∆ñ6RÉ¬ì∞ß–†¶gVÊ7Fñˆ‚W6W$÷F6ÜW47FófóGî∆ˆrÜ∆ˆr¬W6W"í∞¢6ˆÁ7BñB“7G&ñÊráW6W#ÚÊñB«¬W6W#ÚÁVñB«¬""ì∞¢&WGW&‚&ˆˆ∆V‚ÇÜñBbb7G&ñÊrÜ∆ˆrÁW6W$ñB«¬""í””“ñBí«¬áW6W#ÚÊV÷ñ¬bbÊ˜&÷∆ó¶TV÷ñ¬Ü∆ˆrÁW6W$V÷ñ¬í””“Ê˜&÷∆ó¶TV÷ñ¬áW6W"ÊV÷ñ¬ííì∞ß–†¶gVÊ7Fñˆ‚&VÊFW%W6W$7FófóGïfñWrÇí∞¢ñbÇ6V∆V7FVEW6W$7FófóGïW6W"«¬VíÁW6W$7FófóGï7V÷÷'íí&WGW&„∞¢6ˆÁ7BW6W"“6V∆V7FVEW6W$7FófóGïW6W#∞¢6ˆÁ7BˆÊ∆ñÊR“FFRÊÊ˜rÇí“fó&W7F˜&TFFUFÙ÷ñ∆∆ó2áW6W"Ê∆7E6VV‰Bí√“¢c¢∞¢6ˆÁ7BW'&˜'2“6V∆V7FVEW6W$7FófóGî∆ˆw2Êfñ«FW"Üó4W'&˜$7FófóGííÊ∆VÊwFÉ∞¢6ˆÁ7B7V÷÷'îfñV∆G2“∞¢≤$Êˆ÷RR6ˆvÊˆ÷R"¬vWEW6W$Fó7∆îÊ÷RáW6W"ï“¬≤$V÷ñ¬"¬W6W"ÊV÷ñ≈“¬≤%'Vˆ∆Ú"¬vWEW6W%&ˆ∆RáW6W"ï“¿¢≤%7FFÚ"¬ˆÊ∆ñÊRÚ/	˘˙"ˆÊ∆ñÊR"¢.)™¢ˆff∆ñÊR%“¬≤%V«Fñ÷GFófóL:"¬W6W"Ê∆7E6VV‰BÚf˜&÷D7FófóGîFFRáW6W"Ê∆7E6VV‰Bí¢"%“¿¢≤$¶ñˆÊíÊV¬vñ˜&ÊÚ"¬7G&ñÊrá6V∆V7FVEW6W$7FófóGî∆ˆw2Ê∆VÊwFÇï“¬≤$W'&˜&í"¬7G&ñÊrÜW'&˜'2ï–¢“Êfñ«FW"ÇÖ≤¬f«VU“í”‚f«VR”“""bbf«VR“ÁV∆¬ì∞¢VíÁW6W$7FófóGï7V÷÷'íÊñÊÊW$ÖD‘¬“∆Fób6∆73“'6V7Fñˆ‚÷ÜVB#„∆É#‚G∂W66TÖD‘¬ÜvWEW6W$Fó7∆îÊ÷RáW6W"íó”¬ˆÉ#„«7‚6∆73“&7FófR◊7FGW2÷F˜BG∂ˆÊ∆ñÊRÚ&ˆÊ∆ñÊR"¢&ˆff∆ñÊR'“#‚G∂ˆÊ∆ñÊRÚ$ˆÊ∆ñÊR"¢$ˆff∆ñÊR'”¬˜7„„¬ˆFóc„∆F√‚G∑7V÷÷'îfñV∆G2Ê÷ÇÖ∂∆&V¬¬f«VU“í”‚∆Fóc„∆GC‚G∂W66TÖD‘¬Ü∆&V¬ó”¬ˆGC„∆FC‚G∂W66TÖD‘¬áf«VRó”¬ˆFC„¬ˆFócÊíÊ¶ˆñ‚Ç""ó”¬ˆF√Ê∞¢VíÁW6W$7FófóGî6˜VÁBÁFWáD6ˆÁFVÁB“G∑6V∆V7FVEW6W$7FófóGî∆ˆw2Ê∆VÊwFá“¶ñˆÊñ∞¢VíÁW6W$7FófóGïFñ÷V∆ñÊRÊñÊÊW$ÖD‘¬“6V∆V7FVEW6W$7FófóGî∆ˆw2Ê÷Ü∆ˆr”‚∞¢6ˆÁ7BW'&˜"“ó4W'&˜$7FófóGíÜ∆ˆrì∞¢6ˆÁ7BWÜ7EFñ÷R“ÊWrFFRÜfó&W7F˜&TFFUFÙ÷ñ∆∆ó2Ü∆ˆrÊ7&VFVDBííÁFÙ∆ˆ6∆UFñ÷U7G&ñÊrÇ&óB‘ïB"¬≤Ü˜W#¢#"÷FñvóB"¬÷ñÁWFS¢#"÷FñvóB"¬6V6ˆÊC¢#"÷FñvóB"“ì∞¢6ˆÁ7BfñV∆G2“∞¢≤%FóÚFí¶ñˆÊR"¬∆ˆrÊ7FñˆÂGóU“¬≤$6ˆ÷÷W76"¬∆ˆrÊ6ˆ÷÷W76Ê÷R«¬∆ˆrÊ6ˆ÷÷W76ñE“¿¢≤$ñ◊ñÁFÚ"¬∆ˆrÊñ◊ñÁFÙÊ÷U“¬≤$îB4"¬∆ˆrÊñ◊ñÁFı6«¬∆ˆrÊñE6«¬∆ˆrÊñE4“¿¢≤$FWGFv∆í"¬∆ˆrÊ7Fñˆ‰FW67&óFñˆ‚«¬∆ˆrÊFWFñ≈“¬≤$W6óFÚ"¬W'&˜"Ú$W'&˜&R"¢%&óW66óF%“¿¢≤$÷W76vvñÚFV6Êñ6Ú"¬W'&˜"ÚÜ∆ˆrÁFV6ÜÊñ6ƒW'&˜"«¬∆ˆrÊW'&˜"«¬∆ˆrÊFWFñ¬í¢"%–¢“Êfñ«FW"ÇÖ≤¬f«VU“í”‚f«VR”“""bbf«VR“ÁV∆¬ì∞¢&WGW&‚∆'Fñ6∆R6∆73“'W6W"÷7FófóGí÷WfVÁBG∂W'&˜"Ú&ó2÷W'&˜""¢"'“#„«Fñ÷RFFWFñ÷S“"G∂ÊWrFFRÜfó&W7F˜&TFFUFÙ÷ñ∆∆ó2Ü∆ˆrÊ7&VFVDBííÁFÙï4ı7G&ñÊrÇó“#‚G∂W66TÖD‘¬ÜWÜ7EFñ÷Ró”¬˜Fñ÷S„∆F√‚G∂fñV∆G2Ê÷ÇÖ∂∆&V¬¬f«VU“í”‚∆Fóc„∆GC‚G∂W66TÖD‘¬Ü∆&V¬ó”¬ˆGC„∆FC‚G∂W66TÖD‘¬Ö7G&ñÊráf«VRíó”¬ˆFC„¬ˆFócÊíÊ¶ˆñ‚Ç""ó”¬ˆF√„¬ˆ'Fñ6∆SÊ∞¢“íÊ¶ˆñ‚Ç""í«¬s«6∆73“&◊WFVB#‰ÊW77VÊGFófóL:&Vvó7G&Fñ‚VW7FFF¬˜‚s∞ß–†¶7ñÊ2gVÊ7Fñˆ‚∆ˆE6V∆V7FVEW6W$7FófóGíÇí∞¢ñbÇ6‰÷ÊvTFFÇí«¬6V∆V7FVEW6W$7FófóGïW6W"«¬F"í&WGW&„∞¢6ˆÁ7BFFUf«VR“VíÁW6W$7FófóGîFFSÚÁf«VR«¬∆ˆ6ƒFFUf«VRÇì∞¢6ˆÁ7B7F'B“ÊWrFFRÜG∂FFUf«VW’C££ì∞¢6ˆÁ7BVÊB“ÊWrFFRá7F'Bì≤VÊBÁ6WDFFRÜVÊBÊvWDFFRÇí≤ì∞¢VíÁW6W$7FófóGïFñ÷V∆ñÊRÊñÊÊW$ÖD‘¬“s«6∆73“&◊WFVB#‰6&ñ6÷VÁFÚGFófóL:‚‚„¬˜‚s∞¢6ˆÁ7B6Ê“vóBF"Ê6ˆ∆∆V7Fñˆ‚Ç&7FófóGî∆ˆw2"íÁvÜW&RÇ&7&VFVDB"¬#„“"¬7F'BíÁvÜW&RÇ&7&VFVDB"¬#¬"¬VÊBíÊ˜&FW$'íÇ&7&VFVDB"¬&FW62"íÊ∆ñ÷óBÉSíÊvWBÇì∞¢6V∆V7FVEW6W$7FófóGî∆ˆw2“6ÊÊFˆ72Ê÷ÜFˆ2”‚á≤ñC¢Fˆ2ÊñB¬‚‚ÊFˆ2ÊFFÇí“ííÊfñ«FW"Ü∆ˆr”‚W6W$÷F6ÜW47FófóGî∆ˆrÜ∆ˆr¬6V∆V7FVEW6W$7FófóGïW6W"íì∞¢&VÊFW%W6W$7FófóGïfñWrÇì∞ß–†¶7ñÊ2gVÊ7Fñˆ‚˜VÂW6W$7FófóGïfñWráW6W$ñBí∞¢6ˆÁ7B∆∆˜vVB“6‰÷ÊvTFFÇì∞¢VíÁW6W$7FófóGî66W74÷W76vRÁFWáD6ˆÁFVÁB“∆∆˜vVBÚ""¢$66W76Ú6ˆÁ6VÁFóFÚ6ˆ∆Úv∆í÷÷ñÊó7G&F˜&í‚#∞¢VíÁW6W$7FófóGî66W74÷W76vRÊ6∆74∆ó7BÁFˆvv∆RÇ&ÜñFFV‚"¬∆∆˜vVBì∞¢VíÁW6W$7FófóGîF÷ñ‰6ˆÁFVÁCÚÊ6∆74∆ó7BÁFˆvv∆RÇ&ÜñFFV‚"¬∆∆˜vVBì∞¢ñbÇ∆∆˜vVBí&WGW&„∞¢6V∆V7FVEW6W$7FófóGïW6W"“∆Ff˜&’W6W'2ÊfñÊBáW6W"”‚7G&ñÊráW6W"ÊñB«¬W6W"ÁVñB«¬W6W"ÊV÷ñ¬í””“7G&ñÊráW6W$ñBíí«¬ÁV∆√∞¢ñbÇ6V∆V7FVEW6W$7FófóGïW6W"í∞¢VíÁW6W$7FófóGî66W74÷W76vRÁFWáD6ˆÁFVÁB“%WFVÁFRÊˆ‚Fó7ˆÊñ&ñ∆R‚#∞¢VíÁW6W$7FófóGî66W74÷W76vRÊ6∆74∆ó7BÁ&V÷˜fRÇ&ÜñFFV‚"ì∞¢VíÁW6W$7FófóGîF÷ñ‰6ˆÁFVÁBÊ6∆74∆ó7BÊFBÇ&ÜñFFV‚"ì∞¢&WGW&„∞¢–¢ñbáVíÁW6W$7FófóGîFFRbbVíÁW6W$7FófóGîFFRÁf«VRíVíÁW6W$7FófóGîFFRÁf«VR“∆ˆ6ƒFFUf«VRÇì∞¢G'í≤vóB∆ˆE6V∆V7FVEW6W$7FófóGíÇì≤–¢6F6ÇÜW'&˜"í≤VíÁW6W$7FófóGïFñ÷V∆ñÊRÊñÊÊW$ÖD‘¬“s«6∆73“&◊WFVB#‰W'&˜&RGW&ÁFRñ¬6&ñ6÷VÁFÚFV∆∆RGFófóL:„¬˜‚s≤6ˆÁ6ˆ∆RÊW'&˜"Ç$W'&˜&RGFófóL:WFVÁFS¢"¬W'&˜"ì≤–ß–†ßVíÊ7FófUW6W'57V÷÷'ìÚÊFDWfVÁD∆ó7FVÊW"Ç&6∆ñ6≤"¬Çí”‚∞¢ñbÜ6‰÷ÊvTFFÇíívñÊF˜rÊ∆ˆ6Fñˆ‚ÊÜ6Ç“"6FWGFv∆ñÚ◊WFVÁFí÷GFófí#∞¢V«6R∆W'BÇ$66W76Ú6ˆÁ6VÁFóFÚ6ˆ∆Úv∆í÷÷ñÊó7G&F˜&í‚"ì∞ß“ì∞ßVíÊ7FófUW6W'57V÷÷'ìÚÊFDWfVÁD∆ó7FVÊW"Ç&∂WñF˜v‚"¬ÜWfVÁBí”‚≤ñbÇÜWfVÁBÊ∂Wí””“$VÁFW""«¬WfVÁBÊ∂Wí””“""íbb6‰÷ÊvTFFÇíí≤WfVÁBÁ&WfVÁDFVfV«BÇì≤vñÊF˜rÊ∆ˆ6Fñˆ‚ÊÜ6Ç“"6FWGFv∆ñÚ◊WFVÁFí÷GFófí#≤““ì∞ßVíÊ7FófUW6W'4&6¥'F„ÚÊFDWfVÁD∆ó7FVÊW"Ç&6∆ñ6≤"¬Çí”‚≤vñÊF˜rÊ∆ˆ6Fñˆ‚ÊÜ6Ç“"#≤“ì∞ßVíÁW6W$7FófóGî&6¥'F„ÚÊFDWfVÁD∆ó7FVÊW"Ç&6∆ñ6≤"¬Çí”‚≤vñÊF˜rÊ∆ˆ6Fñˆ‚ÊÜ6Ç“"6FWGFv∆ñÚ◊WFVÁFí÷GFófí#≤“ì∞ßVíÁW6W$7FófóGîFFSÚÊFDWfVÁD∆ó7FVÊW"Ç&6ÜÊvR"¬∆ˆE6V∆V7FVEW6W$7FófóGíì∞ßVíÊ7FófUW6W'5&Vg&W6Ñ'F„ÚÊFDWfVÁD∆ó7FVÊW"Ç&6∆ñ6≤"¬Çí”‚≤7FófUW6W'4∆ˆFVB“f«6S≤˜V‰7FófUW6W'4FWFñ≈fñWrÇì≤“ì∞ßVíÊ7FófUW6W'4fñ«FW%Fˆvv∆SÚÊFDWfVÁD∆ó7FVÊW"Ç&6∆ñ6≤"¬Çí”‚≤7FófUW6W'4fñ«FW$˜V‚“7FófUW6W'4fñ«FW$˜V„≤&VÊFW$7FófUW6W'4FWFñ¬Çì≤“ì∞ßVíÊ7FófUW6W'4gV∆≈Fˆvv∆SÚÊFDWfVÁD∆ó7FVÊW"Ç&6∆ñ6≤"¬Çí”‚≤7FófUW6W'4gV∆ƒ∆ó7D˜V‚“7FófUW6W'4gV∆ƒ∆ó7D˜V„≤&VÊFW$7FófUW6W'4FWFñ¬Çì≤“ì∞ßVíÊ7FófUW6W'4∆ˆuFˆvv∆SÚÊFDWfVÁD∆ó7FVÊW"Ç&6∆ñ6≤"¬Çí”‚≤7FófUW6W'4∆ˆt∆ó7D˜V‚“7FófUW6W'4∆ˆt∆ó7D˜V„≤&VÊFW$7FófUW6W'4FWFñ¬Çì≤“ì∞•∑VíÊ7FófUW6W'56V&6ÖW6W"¬VíÊ7FófUW6W'56V&6Ñ6ˆ÷÷W76¬VíÊ7FófUW6W'56V&6Ññ◊ñÁFÚ¬VíÊ7FófUW6W'4fñ«FW$˜W&F˜"¬VíÊ7FófUW6W'4fñ«FW$7Fñˆ‚¬VíÊ7FófUW6W'4W'&˜'4ˆÊ«ï“Êf˜$V6ÇÜV¬”‚V√ÚÊFDWfVÁD∆ó7FVÊW"Ç&ñÁWB"¬&VÊFW$7FófUW6W'4FWFñ¬íì∞ßVíÊ7FófUW6W'4F÷ñ‰6ˆÁ6ˆ∆SÚÊFDWfVÁD∆ó7FVÊW"Ç&6∆ñ6≤"¬ÜWfVÁBí”‚∞¢6ˆÁ7B6&B“WfVÁBÁF&vWBÊ6∆˜6W7BÇ%∂FF÷7FófR÷6&E“"ì∞¢ñbÜ6&Bí≤6V∆V7FVD7FófUW6W'46&B“6&BÊvWDGG&ñ'WFRÇ&FF÷7FófR÷6&B"í«¬"#≤&VÊFW$7FófUW6W'4FWFñ¬Çì≤&WGW&„≤–¢ñbÜWfVÁBÁF&vWBÊ6∆˜6W7BÇ%∂FF÷7FófR÷6&B÷6∆˜6U“"íí≤6V∆V7FVD7FófUW6W'46&B“"#≤&VÊFW$7FófUW6W'4FWFñ¬Çì≤&WGW&„≤–¢6ˆÁ7BW6W%&˜r“WfVÁBÁF&vWBÊ6∆˜6W7BÇ%∂FF÷7FófR◊W6W"÷ñE“"ì∞¢ñbáW6W%&˜rí≤vñÊF˜rÊ∆ˆ6Fñˆ‚ÊÜ6Ç“GFófóF◊WFVÁFS“G∂VÊ6ˆFUU$î6ˆ◊ˆÊVÁBáW6W%&˜rÊvWDGG&ñ'WFRÇ&FF÷7FófR◊W6W"÷ñB"í«¬""ó÷≤&WGW&„≤–¢6ˆÁ7B∆ˆu&˜r“WfVÁBÁF&vWBÊ6∆˜6W7BÇ%∂FF÷7FófR÷∆ˆr÷∂Wï“"ì∞¢ñbÜ∆ˆu&˜rí˜V‰7FófT∆ˆu&ˆ&∆V‘FWFñ¬ávñÊF˜rÊ7FófUW6W'5&VÊFW&VD∆ˆtw&˜W4'î∂WìÚÂ∂∆ˆu&˜rÊvWDGG&ñ'WFRÇ&FF÷7FófR÷∆ˆr÷∂Wí"ï“ì∞ß“ì∞ßVíÊ7FófUW6W'4F÷ñ‰6ˆÁ6ˆ∆SÚÊFDWfVÁD∆ó7FVÊW"Ç&∂WñF˜v‚"¬ÜWfVÁBí”‚∞¢ñbÜWfVÁBÊ∂Wí”“$VÁFW""bbWfVÁBÊ∂Wí”“""í&WGW&„∞¢6ˆÁ7BW6W%&˜r“WfVÁBÁF&vWBÊ6∆˜6W7BÇ%∂FF÷7FófR◊W6W"÷ñE“"ì∞¢ñbáW6W%&˜rí≤WfVÁBÁ&WfVÁDFVfV«BÇì≤vñÊF˜rÊ∆ˆ6Fñˆ‚ÊÜ6Ç“GFófóF◊WFVÁFS“G∂VÊ6ˆFUU$î6ˆ◊ˆÊVÁBáW6W%&˜rÊvWDGG&ñ'WFRÇ&FF÷7FófR◊W6W"÷ñB"í«¬""ó÷≤&WGW&„≤–¢6ˆÁ7B&˜r“WfVÁBÁF&vWBÊ6∆˜6W7BÇ%∂FF÷7FófR÷∆ˆr÷∂Wï“"ì∞¢ñbÇ&˜rí&WGW&„∞¢WfVÁBÁ&WfVÁDFVfV«BÇì∞¢˜V‰7FófT∆ˆu&ˆ&∆V‘FWFñ¬ávñÊF˜rÊ7FófUW6W'5&VÊFW&VD∆ˆtw&˜W4'î∂WìÚÂ∑&˜rÊvWDGG&ñ'WFRÇ&FF÷7FófR÷∆ˆr÷∂Wí"ï“ì∞ß“ì∞††¶Fˆ7V÷VÁBÊFDWfVÁD∆ó7FVÊW"Ç&6∆ñ6≤"¬ÜWfVÁBí”‚∞¢6ˆÁ7B'WGFˆ‚“WfVÁBÁF&vWBÊ6∆˜6W7BÇ&'WGFˆ‚¬"ì∞¢ñbÇ'WGFˆ‚«¬7W'&VÁEW6W"í&WGW&„∞¢6ˆÁ7B∆&V¬“7G&ñÊrÜ'WGFˆ‚ÁFWáD6ˆÁFVÁB«¬'WGFˆ‚ÊvWDGG&ñ'WFRÇ&&ñ÷∆&V¬"í«¬""íÁG&ñ“ÇíÁFÙ∆˜vW$66RÇì∞¢6ˆÁ7BWáG&“∞¢'WGFˆ‰∆&V√¢7G&ñÊrÜ'WGFˆ‚ÁFWáD6ˆÁFVÁB«¬'WGFˆ‚ÊvWDGG&ñ'WFRÇ&&ñ÷∆&V¬"í«¬'WGFˆ‚ÁFóF∆R«¬""íÁG&ñ“Çí¿¢fñWtÊ÷S¢vWD7W'&VÁEfñWtÊ÷RÇí¿¢6ˆ÷÷W76ñC¢6V∆V7FVD6ˆ÷÷W76ñB«¬""¿¢6ˆ÷÷W76Ê÷S¢6V∆V7FVD6ˆ÷÷W76Ê÷R«¬""¿¢ñ◊ñÁFÙñC¢'WGFˆ‚ÊvWDGG&ñ'WFRÇ&FF÷ñ◊ñÁFÚ÷∂Wí"í«¬'WGFˆ‚ÊvWDGG&ñ'WFRÇ&FF÷ñ◊ñÁFÚ÷ñB"í«¬""¿¢ñ◊ñÁFÙÊ÷S¢'WGFˆ‚ÊvWDGG&ñ'WFRÇ&FF÷ñ◊ñÁFÚ÷Ê÷R"í«¬" ¢”∞¢ñbÜ∆&V¬ÊñÊ6«VFW2Ç&Êfñv"íí∆ˆt7FófóGíÇ'&W76ñˆÊUˆÊfñv"¬%&W76ñˆÊR‰dît"¬WáG&ì∞¢ñbÜ∆&V¬ÊñÊ6«VFW2Ç&fGFÚ"íí∆ˆt7FófóGíÇ'&W76ñˆÊUˆfGFÚ"¬%&W76ñˆÊRdEDÚ"¬WáG&ì∞¢ñbÜ∆&V¬ÊñÊ6«VFW2Ç&f˜'¶"íí∆ˆt7FófóGíÇ'&W76ñˆÊUˆf˜'¶"¬%&W76ñˆÊRdı%§"¬WáG&ì∞¢ñbÜ∆&V¬ÊñÊ6«VFW2Ç'vÜG6"íí∆ˆt7FófóGíÇ&ñÁfñı˜vÜG6"¬$ñÁfñÚvÜG4"¬WáG&ì∞¢ñbÜ∆&V¬ÊñÊ6«VFW2Ç&÷"íí∆ˆt7FófóGíÇ&W'GW&ˆ÷"¬$W'GW&÷"¬WáG&ì∞ß“ì∞†