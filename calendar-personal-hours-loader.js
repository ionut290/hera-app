(() => {
  "use strict";

  const normalize = (value) => String(value || "")
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .toLocaleLowerCase("it-IT")
    .replace(/[^a-z0-9@._-]+/g, " ")
    .replace(/\s+/g, " ")
    .trim();

  const compact = (value) => normalize(value).replace(/[^a-z0-9]/g, "");

  function dateKey(value) {
    if (typeof normalizeHoursReportDateKey === "function") {
      try {
        const normalized = normalizeHoursReportDateKey(value);
        if (normalized) return normalized;
      } catch (_) {}
    }
    if (value?.toDate instanceof Function) value = value.toDate();
    if (value instanceof Date && !Number.isNaN(value.getTime())) {
      const year = value.getFullYear();
      const month = String(value.getMonth() + 1).padStart(2, "0");
      const day = String(value.getDate()).padStart(2, "0");
      return `${year}-${month}-${day}`;
    }
    const text = String(value || "").trim();
    const iso = text.match(/^(\d{4})-(\d{2})-(\d{2})/);
    if (iso) return `${iso[1]}-${iso[2]}-${iso[3]}`;
    const italian = text.match(/^(\d{1,2})[\/.\-](\d{1,2})[\/.\-](\d{4})/);
    if (italian) return `${italian[3]}-${italian[2].padStart(2, "0")}-${italian[1].padStart(2, "0")}`;
    return text;
  }

  function collectCurrentUserIdentity() {
    const exact = new Set();
    const names = new Set();
    const addExact = (value) => {
      const key = compact(value);
      if (key) exact.add(key);
    };
    const addName = (value) => {
      const key = normalize(value);
      if (!key || key.includes("@")) return;
      const tokens = key.split(" ").filter((token) => token.length > 1).sort();
      if (tokens.length >= 2) names.add(tokens.join("|"));
    };
    const addRecord = (record) => {
      if (!record || typeof record !== "object") return;
      [record.id, record.uid, record.userId, record.operatoreId, record.personaleId,
       record.email, record.emailAccessoApp, record.linkedUserEmail].forEach(addExact);
      [record.displayName, record.nomeCompleto, record.fullName, record.operatore, record.name,
       `${record.nome || ""} ${record.cognome || ""}`,
       `${record.cognome || ""} ${record.nome || ""}`].forEach(addName);
    };

    if (typeof currentUser !== "undefined") addRecord(currentUser);
    if (typeof getCurrentUserSquadraIdentity === "function") {
      try { addRecord(getCurrentUserSquadraIdentity()); } catch (_) {}
    }

    const userEmail = normalize(typeof currentUser !== "undefined" ? currentUser?.email : "");
    const userUid = compact(typeof currentUser !== "undefined" ? currentUser?.uid : "");
    const sources = [];
    if (typeof personaleRecords !== "undefined" && Array.isArray(personaleRecords)) sources.push(...personaleRecords);
    if (typeof platformUsers !== "undefined" && Array.isArray(platformUsers)) sources.push(...platformUsers);
    sources.forEach((record) => {
      const recordEmails = [record?.email, record?.emailAccessoApp, record?.linkedUserEmail].map(normalize);
      const recordIds = [record?.id, record?.uid, record?.userId].map(compact);
      if ((userEmail && recordEmails.includes(userEmail)) || (userUid && recordIds.includes(userUid))) addRecord(record);
    });

    return { exact, names };
  }

  function rowMatchesCurrentUser(row, entry, identity) {
    const values = [
      row?.operatoreId, row?.personaleId, row?.uid, row?.userId, row?.email,
      entry?.operatoreId, entry?.personaleId, entry?.uid, entry?.userId, entry?.email
    ];
    if (values.some((value) => identity.exact.has(compact(String(value || "").replace(/^utente:/i, ""))))) return true;

    const nameValues = [
      row?.operatore, row?.nomeOperatore, row?.nome, row?.name, row?.displayName,
      `${row?.nome || ""} ${row?.cognome || ""}`,
      `${row?.cognome || ""} ${row?.nome || ""}`,
      entry?.operatore, entry?.nomeOperatore
    ];
    return nameValues.some((value) => {
      const tokens = normalize(value).split(" ").filter((token) => token.length > 1).sort();
      return tokens.length >= 2 && identity.names.has(tokens.join("|"));
    });
  }

  function getHoursValue(row) {
    const candidates = [row?.ore, row?.hours, row?.totaleOre, row?.oreTotali, row?.quantita];
    for (const candidate of candidates) {
      const number = Number(String(candidate ?? "").replace(",", "."));
      if (Number.isFinite(number) && number > 0) return number;
    }
    return 0;
  }

  function collectLoadedReports() {
    const result = [];
    const seen = new Set();
    const append = (items) => {
      if (!Array.isArray(items)) return;
      items.forEach((item) => {
        if (!item || String(item.status || "").toLowerCase() === "rejected") return;
        const key = String(item.id || `${dateKey(item.date)}-${result.length}`);
        if (seen.has(key)) return;
        seen.add(key);
        result.push(item);
      });
    };
    if (typeof allHoursReports !== "undefined") append(allHoursReports);
    if (typeof allHoursApprovalRequests !== "undefined") append(allHoursApprovalRequests);
    return result;
  }

  function personalRowsForDate(selectedDate) {
    if (!selectedDate) return [];
    const identity = collectCurrentUserIdentity();
    const rows = [];
    collectLoadedReports().forEach((report) => {
      if (dateKey(report.date || report.data || report.giorno) !== selectedDate) return;
      const entries = Array.isArray(report.entries) ? report.entries : [report];
      entries.forEach((entry) => {
        const entryRows = Array.isArray(entry?.rows)
          ? entry.rows
          : (Array.isArray(entry?.operatori) ? entry.operatori : []);
        entryRows.forEach((row) => {
          const hours = getHoursValue(row);
          if (hours > 0 && rowMatchesCurrentUser(row, entry, identity)) rows.push({ report, entry, row, hours });
        });
      });
    });
    return rows;
  }

  function refreshCalendar() {
    if (typeof calendarMode === "undefined" || calendarMode !== "hours") return;
    if (typeof renderCalendar === "function") renderCalendar();
  }

  function installCompatibilityMatcher() {
    if (typeof getPersonalHoursRowsForDate !== "function") return false;
    try {
      getPersonalHoursRowsForDate = personalRowsForDate;
      window.__personalHoursCalendarUsesLoadedData = true;
      refreshCalendar();
      return true;
    } catch (error) {
      console.warn("Impossibile installare il matching compatibile delle ore personali", error);
      return false;
    }
  }

  let attempts = 0;
  const timer = window.setInterval(() => {
    attempts += 1;
    if (installCompatibilityMatcher() || attempts >= 80) window.clearInterval(timer);
  }, 250);

  document.addEventListener("click", (event) => {
    const button = event.target instanceof Element ? event.target.closest("button") : null;
    if (!button) return;
    if (["calendar-choice-hours-btn", "calendar-hours-tab", "calendar-prev-btn", "calendar-next-btn", "calendar-today-btn"].includes(button.id)) {
      window.setTimeout(() => {
        installCompatibilityMatcher();
        refreshCalendar();
      }, 0);
    }
  }, true);

  document.addEventListener("visibilitychange", () => {
    if (!document.hidden) window.setTimeout(refreshCalendar, 0);
  });

  // Il file di ripristino esisteva nel repository ma non veniva mai caricato.
  // Lo avvia una sola volta con una versione nuova per evitare la cache PWA precedente.
  if (!document.querySelector('script[data-ripristino-id-personale]')) {
    const restoreScript = document.createElement('script');
    restoreScript.src = './rubrica-personale-restore.js?v=20260802c';
    restoreScript.defer = true;
    restoreScript.dataset.ripristinoIdPersonale = '1';
    document.head.appendChild(restoreScript);
  }
})();
