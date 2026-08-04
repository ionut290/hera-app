(() => {
  "use strict";

  function apiReady(timeoutMs = 10000) {
    if (window.HeraSharedStaticViews?.installed) return Promise.resolve(window.HeraSharedStaticViews);
    return new Promise((resolve, reject) => {
      const startedAt = Date.now();
      const timer = setInterval(() => {
        if (window.HeraSharedStaticViews?.installed) {
          clearInterval(timer);
          resolve(window.HeraSharedStaticViews);
        } else if (Date.now() - startedAt >= timeoutMs) {
          clearInterval(timer);
          reject(new Error("Servizio viste condivise non disponibile."));
        }
      }, 50);
    });
  }

  function subscribeRegistry(kind) {
    const isPersonale = kind === "personale";
    if (!currentUser) return Promise.resolve(false);
    if (isPersonale) stopPersonaleSubscription(); else stopMezziSubscription();
    return apiReady().then((api) => new Promise((resolve) => {
      let initial = true;
      const unsubscribe = api.subscribe("registri", "corrente", (view, metadata = {}) => {
        const records = Array.isArray(view?.payload?.[kind]) ? view.payload[kind] : [];
        if (isPersonale) {
          personaleRecords = records.map((record) => normalizePersonaleDocument({ id: record.id, data: () => record }));
          personaleLoadState = { status: "loaded", message: "" };
          renderPersonaleList(ui.personaleLista, personaleRecords, deletePersonale);
          refreshResolvedUserIdentity();
          renderHoursOperatoriOptions();
          renderSquadre();
        } else {
          mezziRecords = records.map((record) => normalizeMezzoDocument({ id: record.id, data: () => record }));
          mezziLoadState = { status: "loaded", message: "" };
          renderMezziList(ui.mezziLista, mezziRecords, deleteMezzo);
        }
        console.debug("[SHARED VIEWS] registri", { kind, records: records.length, source: metadata.source });
        renderTodaySummary();
        updateSquadraHintFromSources();
        updateSuggestionLists();
        if (initial) { initial = false; resolve(true); }
      });
      if (isPersonale) unsubscribePersonale = unsubscribe; else unsubscribeMezzi = unsubscribe;
    })).catch((error) => {
      logFirestoreError(`LOAD ${kind.toUpperCase()} SHARED VIEW`, error);
      if (isPersonale) personaleLoadState = { status: "error", message: "Vista personale non disponibile." };
      else mezziLoadState = { status: "error", message: "Vista mezzi non disponibile." };
      return false;
    });
  }

  subscribePersonale = () => subscribeRegistry("personale");
  subscribeMezzi = () => subscribeRegistry("mezzi");

  subscribeSquadre = function subscribeSquadreFromSharedViews() {
    if (!currentUser) return Promise.resolve(false);
    stopSquadreSubscription();
    const dates = [...new Set([getActiveSquadreDateKey(), getTodayDateKey()].filter(Boolean))];
    squadreLoadState = { status: "loading", message: "Caricamento squadre condivise..." };
    renderSquadre();
    return apiReady().then((api) => new Promise((resolve) => {
      let pending = dates.length;
      const stops = dates.map((dateKey) => api.subscribe("squadre", dateKey, (view, metadata = {}) => {
        const items = Array.isArray(view?.payload?.squadre) ? view.payload.squadre : [];
        const history = new Map();
        items.forEach((item) => {
          const commessaId = String(item?.commessaId || item?.commessa || "");
          if (!commessaId) return;
          if (Array.isArray(item.squadre)) history.set(commessaId, { ...item, dateKey, commessaId });
          else {
            const current = history.get(commessaId) || { dateKey, commessaId, squadre: [] };
            current.squadre.push(item);
            history.set(commessaId, current);
          }
        });
        squadreHistoryByDate.set(dateKey, history);
        squadreLoadState = { status: "loaded", message: "" };
        console.debug("[SHARED VIEWS] squadre", { dateKey, records: items.length, source: metadata.source });
        renderTodaySummary();
        renderSquadre();
        updateCommessaDashboard();
        renderCommesseHomeList();
        autofillSquadraForm();
        if (--pending <= 0) resolve(true);
      }));
      unsubscribeSquadreHistory = () => stops.forEach((stop) => stop?.());
    })).catch((error) => {
      logFirestoreError("LOAD SQUADRE SHARED VIEW", error);
      squadreLoadState = { status: "error", message: "Vista squadre non disponibile; nessuna raccolta sorgente è stata letta." };
      renderSquadre();
      return false;
    });
  };

  subscribeHoursStats = function subscribeHoursFromSharedView() {
    if (unsubscribeHoursStats || !currentUser) return;
    const month = new Date().toISOString().slice(0, 7);
    apiReady().then((api) => {
      unsubscribeHoursStats = api.subscribe("calendario", month, (view, metadata = {}) => {
        const reports = Array.isArray(view?.payload?.reports) ? view.payload.reports : [];
        allHoursReports = reports.filter((item) => String(item.sourceCollection || "oreReports") === "oreReports");
        allHoursApprovalRequests = reports.filter((item) => String(item.sourceCollection || "") === "oreApprovalRequests");
        hoursApprovalRequests = allHoursApprovalRequests;
        hoursReportsLoaded = true;
        hoursApprovalsLoaded = true;
        console.debug("[SHARED VIEWS] calendario", { month, reports: reports.length, source: metadata.source });
        renderTodaySummary();
        recalculateCommessaWorkSummaries();
        renderParentCommessaOverview();
        renderHoursApprovalRequests();
        renderSquadre();
        updateCommessaDashboard();
        if (!ui.calendarPage?.classList.contains("hidden") && calendarMode === "hours") renderCalendar();
      });
      unsubscribeHoursApprovals = () => {};
    }).catch((error) => console.error("Errore vista calendario condivisa:", error));
  };
})();
