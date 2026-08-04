(() => {
  "use strict";

  const original = {
    personale: subscribePersonale,
    mezzi: subscribeMezzi,
    squadre: subscribeSquadre
  };
  const FALLBACK_MS = 1400;

  const romeDateKey = () => new Intl.DateTimeFormat("en-CA", {
    timeZone: "Europe/Rome", year: "numeric", month: "2-digit", day: "2-digit"
  }).format(new Date());
  const visibleMonth = () => String(ui?.hoursStatsMonth?.value || ui?.hoursTableMonth?.value || romeDateKey().slice(0, 7));

  function apiReady(timeoutMs = FALLBACK_MS) {
    if (window.HeraSharedStaticViews?.installed) return Promise.resolve(window.HeraSharedStaticViews);
    return new Promise((resolve, reject) => {
      const startedAt = Date.now();
      const timer = setInterval(() => {
        if (window.HeraSharedStaticViews?.installed) { clearInterval(timer); resolve(window.HeraSharedStaticViews); }
        else if (Date.now() - startedAt >= timeoutMs) { clearInterval(timer); reject(new Error("Viste condivise non disponibili.")); }
      }, 50);
    });
  }

  function registry(kind) {
    const isPersonale = kind === "personale";
    if (!currentUser) return Promise.resolve(false);
    if (isPersonale) stopPersonaleSubscription(); else stopMezziSubscription();
    return apiReady().then((api) => new Promise((resolve) => {
      let settled = false;
      let stopShared = null;
      const fallback = () => {
        if (settled) return;
        settled = true;
        stopShared?.();
        console.warn("[SHARED VIEWS] fallback sorgente", { kind });
        Promise.resolve(original[kind]()).then(resolve);
      };
      const timer = setTimeout(fallback, FALLBACK_MS);
      stopShared = api.subscribe("registri", "corrente", (view, metadata = {}) => {
        const records = Array.isArray(view?.payload?.[kind]) ? view.payload[kind] : [];
        if (!records.length) return fallback();
        if (settled) return;
        settled = true;
        clearTimeout(timer);
        if (isPersonale) {
          personaleRecords = records.map((record) => normalizePersonaleDocument({ id: record.id, data: () => record }));
          personaleLoadState = { status: "loaded", message: "" };
          renderPersonaleList(ui.personaleLista, personaleRecords, deletePersonale);
          refreshResolvedUserIdentity(); renderHoursOperatoriOptions(); renderSquadre();
          unsubscribePersonale = stopShared;
        } else {
          mezziRecords = records.map((record) => normalizeMezzoDocument({ id: record.id, data: () => record }));
          mezziLoadState = { status: "loaded", message: "" };
          renderMezziList(ui.mezziLista, mezziRecords, deleteMezzo);
          unsubscribeMezzi = stopShared;
        }
        console.debug("[SHARED VIEWS] registri", { kind, records: records.length, source: metadata.source });
        renderTodaySummary(); updateSquadraHintFromSources(); updateSuggestionLists();
        resolve(true);
      });
    })).catch(() => original[kind]());
  }

  subscribePersonale = () => registry("personale");
  subscribeMezzi = () => registry("mezzi");

  subscribeSquadre = function subscribeSquadreSafe() {
    if (!currentUser) return Promise.resolve(false);
    stopSquadreSubscription();
    const dates = [...new Set([getActiveSquadreDateKey(), getTodayDateKey()].filter(Boolean))];
    return apiReady().then((api) => new Promise((resolve) => {
      let completed = 0;
      let usedShared = false;
      let settled = false;
      const stops = [];
      const fallback = () => {
        if (settled) return;
        settled = true;
        stops.forEach((stop) => stop?.());
        console.warn("[SHARED VIEWS] fallback squadre mirato", { dates });
        Promise.resolve(original.squadre()).then(resolve);
      };
      const timer = setTimeout(fallback, FALLBACK_MS);
      dates.forEach((dateKey) => {
        stops.push(api.subscribe("squadre", dateKey, (view, metadata = {}) => {
          const items = Array.isArray(view?.payload?.squadre) ? view.payload.squadre : [];
          if (!items.length) return fallback();
          usedShared = true;
          const history = new Map();
          items.forEach((item) => {
            const commessaId = String(item?.commessaId || item?.commessa || "");
            if (!commessaId) return;
            if (Array.isArray(item.squadre)) history.set(commessaId, { ...item, dateKey, commessaId });
            else {
              const current = history.get(commessaId) || { dateKey, commessaId, squadre: [] };
              current.squadre.push(item); history.set(commessaId, current);
            }
          });
          squadreHistoryByDate.set(dateKey, history);
          squadreLoadState = { status: "loaded", message: "" };
          console.debug("[SHARED VIEWS] squadre", { dateKey, records: items.length, source: metadata.source });
          renderTodaySummary(); renderSquadre(); updateCommessaDashboard(); renderCommesseHomeList(); autofillSquadraForm();
          completed += 1;
          if (!settled && usedShared && completed >= dates.length) {
            settled = true; clearTimeout(timer);
            unsubscribeSquadreHistory = () => stops.forEach((stop) => stop?.());
            resolve(true);
          }
        }));
      });
    })).catch(() => original.squadre());
  };

  subscribeHoursStats = function subscribeHoursSafe() {
    if (unsubscribeHoursStats || !currentUser) return;
    const month = visibleMonth();
    let received = false;
    const apply = (reports, source) => {
      allHoursReports = reports.filter((item) => String(item.sourceCollection || "oreReports") === "oreReports");
      allHoursApprovalRequests = reports.filter((item) => String(item.sourceCollection || "") === "oreApprovalRequests");
      hoursApprovalRequests = allHoursApprovalRequests;
      hoursReportsLoaded = true; hoursApprovalsLoaded = true;
      console.debug("[SHARED VIEWS] calendario", { month, reports: reports.length, source });
      renderTodaySummary(); recalculateCommessaWorkSummaries(); renderParentCommessaOverview();
      renderHoursApprovalRequests(); renderSquadre(); updateCommessaDashboard();
      if (!ui.calendarPage?.classList.contains("hidden") && calendarMode === "hours") renderCalendar();
    };
    const fallback = async () => {
      if (received) return;
      const meta = getMonthMeta(month);
      const reports = await fetchHoursReportsForMonth(month, meta, { includePendingApprovals: true });
      apply(reports, "source-month-fallback");
    };
    apiReady().then((api) => {
      unsubscribeHoursStats = api.subscribe("calendario", month, (view, metadata = {}) => {
        const reports = Array.isArray(view?.payload?.reports) ? view.payload.reports : [];
        if (!reports.length) return void fallback();
        received = true; apply(reports, metadata.source);
      });
      unsubscribeHoursApprovals = () => {};
      setTimeout(() => void fallback(), FALLBACK_MS);
    }).catch(() => void fallback());
  };

  window.HeraSharedViewFallback = { installed: true, romeDateKey, visibleMonth };
})();
