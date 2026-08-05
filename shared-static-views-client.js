(() => {
  "use strict";

  const original = {
    personale: subscribePersonale,
    mezzi: subscribeMezzi
  };
  const sourceSubscriptions = {
    users: typeof subscribeUsers === "function" ? subscribeUsers : null,
    hoursStats: typeof subscribeHoursStats === "function" ? subscribeHoursStats : null,
    commessaStats: typeof subscribeStatsForCommesse === "function" ? subscribeStatsForCommesse : null
  };
  const FALLBACK_MS = 9000;
  const registry = {
    unsubscribe: null,
    active: new Set(),
    sourceFallback: new Set(),
    timers: new Map(),
    pending: new Map(),
    lastView: null
  };
  const lazyStartup = {
    installed: true,
    version: "1.1.1",
    fullUsersEnabled: false,
    fullUsersStarted: false,
    usersSessionUid: "",
    lightUsersUid: "",
    hoursSourceEnabled: false,
    commessaStatsEnabled: false,
    blockedHoursStarts: 0,
    blockedProgrammaticHoursStarts: 0,
    blockedCommessaStatsStarts: 0,
    ignoredStaticCalendarUpdates: 0,
    calendarUnsubscribe: null,
    calendarMonth: ""
  };

  function settle(kind, value) {
    const pending = registry.pending.get(kind) || [];
    registry.pending.delete(kind);
    pending.forEach((resolve) => resolve(value));
  }

  function clearFallbackTimer(kind) {
    const timer = registry.timers.get(kind);
    if (timer) clearTimeout(timer);
    registry.timers.delete(kind);
  }

  function release(kind) {
    clearFallbackTimer(kind);
    registry.active.delete(kind);
    if (registry.active.size || !registry.unsubscribe) return;
    registry.unsubscribe();
    registry.unsubscribe = null;
    registry.lastView = null;
  }

  function useOriginal(kind, reason) {
    if (registry.sourceFallback.has(kind)) return;
    registry.sourceFallback.add(kind);
    clearFallbackTimer(kind);
    registry.active.delete(kind);
    console.warn("[SHARED REGISTRIES] fallback sorgente", { kind, reason });
    Promise.resolve(original[kind]())
      .then(() => settle(kind, false))
      .catch((error) => {
        console.error("[SHARED REGISTRIES] fallback fallito", { kind, error });
        settle(kind, false);
      });
    if (!registry.active.size && registry.unsubscribe) {
      registry.unsubscribe();
      registry.unsubscribe = null;
      registry.lastView = null;
    }
  }

  function validView(view) {
    return Boolean(
      view &&
      view.payload &&
      Array.isArray(view.payload.personale) &&
      Array.isArray(view.payload.mezzi)
    );
  }

  function applyPersonale(records) {
    personaleRecords = records.map((record) =>
      normalizePersonaleDocument({ id: record.id, data: () => record })
    );
    personaleLoadState = { status: "loaded", message: "" };
    renderPersonaleList(ui.personaleLista, personaleRecords, deletePersonale);
    refreshResolvedUserIdentity();
    renderHoursOperatoriOptions();
    renderSquadre();
    renderTodaySummary();
    updateSquadraHintFromSources();
    updateSuggestionLists();
  }

  function applyMezzi(records) {
    mezziRecords = records.map((record) =>
      normalizeMezzoDocument({ id: record.id, data: () => record })
    );
    mezziLoadState = { status: "loaded", message: "" };
    renderMezziList(ui.mezziLista, mezziRecords, deleteMezzo);
    renderTodaySummary();
    updateSquadraHintFromSources();
    updateSuggestionLists();
  }

  function applyView(view, metadata = {}) {
    if (!validView(view)) {
      [...registry.active].forEach((kind) => useOriginal(kind, "payload-non-valido"));
      return;
    }
    registry.lastView = view;
    if (registry.active.has("personale") && !registry.sourceFallback.has("personale")) {
      applyPersonale(view.payload.personale);
      clearFallbackTimer("personale");
      settle("personale", true);
    }
    if (registry.active.has("mezzi") && !registry.sourceFallback.has("mezzi")) {
      applyMezzi(view.payload.mezzi);
      clearFallbackTimer("mezzi");
      settle("mezzi", true);
    }
    console.debug("[SHARED REGISTRIES] registri", {
      personale: view.payload.personale.length,
      mezzi: view.payload.mezzi.length,
      source: metadata.source || "firestore"
    });
  }

  function ensureListener() {
    if (registry.unsubscribe) return;
    if (typeof db === "undefined" || !db?.collection) {
      [...registry.active].forEach((kind) => useOriginal(kind, "firestore-non-disponibile"));
      return;
    }
    registry.unsubscribe = db.collection("sharedStaticViews").doc("registri__corrente")
      .onSnapshot((snapshot) => {
        if (!snapshot.exists) {
          [...registry.active].forEach((kind) => useOriginal(kind, "documento-mancante"));
          return;
        }
        applyView({ id: snapshot.id, ...(snapshot.data() || {}) }, {
          source: snapshot.metadata?.fromCache ? "firestore-cache" : "firestore"
        });
      }, (error) => {
        console.warn("[SHARED REGISTRIES] vista non disponibile", error);
        [...registry.active].forEach((kind) => useOriginal(kind, "errore-listener"));
      });
  }

  function subscribe(kind) {
    if (!currentUser) return Promise.resolve(false);
    if (kind === "personale") stopPersonaleSubscription();
    else stopMezziSubscription();

    registry.active.add(kind);
    registry.sourceFallback.delete(kind);
    const promise = new Promise((resolve) => {
      const pending = registry.pending.get(kind) || [];
      pending.push(resolve);
      registry.pending.set(kind, pending);
    });

    if (kind === "personale") unsubscribePersonale = () => release(kind);
    else unsubscribeMezzi = () => release(kind);

    registry.timers.set(kind, setTimeout(
      () => useOriginal(kind, "timeout"),
      FALLBACK_MS
    ));

    if (registry.lastView) applyView(registry.lastView, { source: "memory" });
    else ensureListener();
    return promise;
  }

  function runSafely(callback, label) {
    try {
      return callback?.();
    } catch (error) {
      console.error(`[LIGHT STARTUP] ${label}`, error);
      return null;
    }
  }

  function subscribeCurrentPlatformUserOnly() {
    if (!sourceSubscriptions.users || !currentUser?.uid) return null;
    const uid = String(currentUser.uid);

    if (lazyStartup.usersSessionUid && lazyStartup.usersSessionUid !== uid) {
      lazyStartup.fullUsersEnabled = false;
      lazyStartup.fullUsersStarted = false;
      lazyStartup.lightUsersUid = "";
    }
    lazyStartup.usersSessionUid = uid;

    if (lazyStartup.fullUsersEnabled) {
      if (lazyStartup.fullUsersStarted && typeof unsubscribeUsers === "function") return unsubscribeUsers;
      lazyStartup.fullUsersStarted = true;
      return sourceSubscriptions.users();
    }

    if (lazyStartup.lightUsersUid === uid && typeof unsubscribeUsers === "function") {
      return unsubscribeUsers;
    }

    const originalCanManageData = canManageData;
    try {
      canManageData = () => false;
      const result = sourceSubscriptions.users();
      lazyStartup.lightUsersUid = uid;
      console.debug("[LIGHT STARTUP] platformUsers limitato all'utente corrente", { uid });
      return result;
    } finally {
      canManageData = originalCanManageData;
    }
  }

  function enableFullUsers() {
    if (!sourceSubscriptions.users || lazyStartup.fullUsersStarted) return;
    lazyStartup.fullUsersEnabled = true;
    lazyStartup.fullUsersStarted = true;
    lazyStartup.usersSessionUid = String(currentUser?.uid || "");
    lazyStartup.lightUsersUid = "";
    runSafely(() => sourceSubscriptions.users(), "caricamento completo utenti");
    console.debug("[LIGHT STARTUP] elenco utenti completo attivato dalla sezione Gestione utenti");
  }

  function gatedHoursStats() {
    if (!lazyStartup.hoursSourceEnabled) {
      lazyStartup.blockedHoursStarts += 1;
      console.debug("[LIGHT STARTUP] oreReports rinviato", {
        tentativiBloccati: lazyStartup.blockedHoursStarts
      });
      return null;
    }
    return sourceSubscriptions.hoursStats?.();
  }

  function stopStaticCalendarForFullHours() {
    if (!lazyStartup.calendarUnsubscribe) return;
    lazyStartup.calendarUnsubscribe();
    lazyStartup.calendarUnsubscribe = null;
    console.debug("[LIGHT STARTUP] vista calendario ridotta fermata: priorità alle ore complete");
  }

  function enableHoursSource(trigger = null) {
    const isClickEvent = Boolean(
      trigger &&
      typeof trigger === "object" &&
      "isTrusted" in trigger
    );

    if (isClickEvent && trigger.isTrusted !== true) {
      lazyStartup.blockedProgrammaticHoursStarts += 1;
      console.debug("[LIGHT STARTUP] avvio automatico oreReports bloccato", {
        tentativiBloccati: lazyStartup.blockedProgrammaticHoursStarts
      });
      return;
    }

    if (!sourceSubscriptions.hoursStats || lazyStartup.hoursSourceEnabled) return;
    lazyStartup.hoursSourceEnabled = true;
    stopStaticCalendarForFullHours();
    runSafely(() => sourceSubscriptions.hoursStats(), "caricamento ore");
    console.debug("[LIGHT STARTUP] oreReports attivato da azione esplicita");
  }

  function gatedCommessaStats() {
    if (!lazyStartup.commessaStatsEnabled) {
      lazyStartup.blockedCommessaStatsStarts += 1;
      console.debug("[LIGHT STARTUP] statistiche impianti rinviate", {
        tentativiBloccati: lazyStartup.blockedCommessaStatsStarts
      });
      return null;
    }
    return sourceSubscriptions.commessaStats?.();
  }

  function enableCommessaStats() {
    if (!sourceSubscriptions.commessaStats || lazyStartup.commessaStatsEnabled) return;
    lazyStartup.commessaStatsEnabled = true;
    runSafely(() => sourceSubscriptions.commessaStats(), "caricamento statistiche commesse");
    console.debug("[LIGHT STARTUP] statistiche impianti attivate su richiesta");
  }

  function currentRomeMonth() {
    return new Intl.DateTimeFormat("en-CA", {
      timeZone: "Europe/Rome",
      year: "numeric",
      month: "2-digit"
    }).format(new Date()).slice(0, 7);
  }

  function normalizeStaticReport(record) {
    if (!record || typeof record !== "object") return null;
    const copy = { ...record };
    delete copy.sourceCollection;
    delete copy.sourceKey;
    return copy;
  }

  function applyStaticCalendar(view, metadata = {}) {
    if (lazyStartup.hoursSourceEnabled) {
      lazyStartup.ignoredStaticCalendarUpdates += 1;
      console.debug("[LIGHT STARTUP] aggiornamento calendario ridotto ignorato: ore complete attive", {
        ignorati: lazyStartup.ignoredStaticCalendarUpdates,
        source: metadata.source || "shared-view"
      });
      return;
    }

    const reports = Array.isArray(view?.payload?.reports) ? view.payload.reports : [];
    const directReports = reports
      .filter((record) => !record?.sourceCollection || record.sourceCollection === "oreReports")
      .map(normalizeStaticReport)
      .filter(Boolean);
    const approvalReports = reports
      .filter((record) => record?.sourceCollection === "oreApprovalRequests")
      .map(normalizeStaticReport)
      .filter(Boolean);

    allHoursReports = directReports;
    allHoursApprovalRequests = approvalReports;
    hoursReportsLoaded = true;

    runSafely(() => renderTodaySummary(), "render riepilogo ore statiche");
    if (!ui.calendarPage?.classList.contains("hidden") && calendarMode === "hours") {
      runSafely(() => renderCalendar(), "render calendario statico");
    }

    console.debug("[LIGHT STARTUP] calendario condiviso applicato", {
      mese: view?.payload?.month || lazyStartup.calendarMonth,
      ore: directReports.length,
      richieste: approvalReports.length,
      source: metadata.source || "shared-view"
    });
  }

  function subscribeStaticCalendar() {
    const api = window.HeraSharedStaticViews;
    if (lazyStartup.hoursSourceEnabled || !api?.subscribe || lazyStartup.calendarUnsubscribe) return;
    lazyStartup.calendarMonth = currentRomeMonth();
    lazyStartup.calendarUnsubscribe = api.subscribe(
      "calendario",
      lazyStartup.calendarMonth,
      applyStaticCalendar
    );
  }

  function bindCapture(id, handler) {
    const node = document.getElementById(id);
    if (!node || node.dataset.lightStartupBound === "true") return false;
    node.dataset.lightStartupBound = "true";
    node.addEventListener("click", handler, true);
    return true;
  }

  function installLazyTriggers() {
    [
      "open-panel-utenti",
      "open-panel-notifiche",
      "chat-open-btn",
      "open-private-docs-btn",
      "open-private-docs-upload-btn",
      "documents-new-btn"
    ].forEach((id) => bindCapture(id, enableFullUsers));

    bindCapture("open-hours-btn", enableHoursSource);

    bindCapture("open-panel-commesse", enableCommessaStats);
    bindCapture("toggle-commesse-home-btn", enableCommessaStats);

    subscribeStaticCalendar();
  }

  subscribePersonale = () => subscribe("personale");
  subscribeMezzi = () => subscribe("mezzi");

  if (sourceSubscriptions.users) subscribeUsers = subscribeCurrentPlatformUserOnly;
  if (sourceSubscriptions.hoursStats) subscribeHoursStats = gatedHoursStats;
  if (sourceSubscriptions.commessaStats) subscribeStatsForCommesse = gatedCommessaStats;

  window.HeraSharedRegistries = {
    installed: true,
    version: "4.0.0",
    getRecords: (kind) => {
      const records = registry.lastView?.payload?.[kind];
      return Array.isArray(records) ? records.slice() : null;
    },
    getState: () => ({
      active: [...registry.active],
      fallback: [...registry.sourceFallback],
      listening: Boolean(registry.unsubscribe)
    })
  };

  window.HeraLightStartup = {
    installed: true,
    version: lazyStartup.version,
    enableFullUsers,
    enableHoursSource,
    enableCommessaStats,
    getState: () => ({
      fullUsersEnabled: lazyStartup.fullUsersEnabled,
      fullUsersStarted: lazyStartup.fullUsersStarted,
      usersSessionUid: lazyStartup.usersSessionUid,
      lightUsersUid: lazyStartup.lightUsersUid,
      hoursSourceEnabled: lazyStartup.hoursSourceEnabled,
      commessaStatsEnabled: lazyStartup.commessaStatsEnabled,
      blockedHoursStarts: lazyStartup.blockedHoursStarts,
      blockedProgrammaticHoursStarts: lazyStartup.blockedProgrammaticHoursStarts,
      blockedCommessaStatsStarts: lazyStartup.blockedCommessaStatsStarts,
      ignoredStaticCalendarUpdates: lazyStartup.ignoredStaticCalendarUpdates,
      calendarMonth: lazyStartup.calendarMonth,
      calendarSharedViewActive: Boolean(lazyStartup.calendarUnsubscribe)
    })
  };

  if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", installLazyTriggers, { once: true });
  } else {
    installLazyTriggers();
  }
})();
