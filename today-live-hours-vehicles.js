"use strict";

(() => {
  if (window.__heraTodayLiveHoursVehiclesInstalled) return;
  window.__heraTodayLiveHoursVehiclesInstalled = true;

  const ROME_TIME_ZONE = "Europe/Rome";
  let minuteTimeout = null;
  let minuteInterval = null;

  const escapeMarkup = (value) => String(value ?? "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#039;");

  const normalizeIdentity = (value) => String(value || "")
    .toLocaleLowerCase("it-IT")
    .replace(/\s+/g, " ")
    .trim();

  function getLiveElement(id, fallback) {
    const element = document.getElementById(id);
    if (element) return element;
    return fallback?.isConnected ? fallback : null;
  }

  function getAssignmentsSafe() {
    try {
      if (typeof findCurrentUserSquadreForDate !== "function" || typeof getTodayDateKey !== "function") return [];
      const result = findCurrentUserSquadreForDate(getTodayDateKey());
      return Array.isArray(result) ? result : [];
    } catch (error) {
      console.warn("Riepilogo Oggi: impossibile leggere la squadra assegnata", error);
      return [];
    }
  }

  function getRomeClockMinutes(now = new Date()) {
    const parts = new Intl.DateTimeFormat("it-IT", {
      timeZone: ROME_TIME_ZONE,
      hour: "2-digit",
      minute: "2-digit",
      hourCycle: "h23"
    }).formatToParts(now);
    const values = Object.fromEntries(parts.map(({ type, value }) => [type, value]));
    return (Number(values.hour) * 60) + Number(values.minute);
  }

  function parseTimeToMinutes(value) {
    if (typeof parseSquadraTimeToMinutes === "function") {
      const parsed = parseSquadraTimeToMinutes(value);
      if (parsed !== null && Number.isFinite(Number(parsed))) return Number(parsed);
    }
    const match = String(value || "").trim().match(/^(\d{1,2}):(\d{2})$/);
    if (!match) return null;
    const hours = Number(match[1]);
    const minutes = Number(match[2]);
    if (hours < 0 || hours > 23 || minutes < 0 || minutes > 59) return null;
    return (hours * 60) + minutes;
  }

  function getRowStartMinutes(row) {
    try {
      if (typeof getSquadraOrarioParts === "function") {
        const start = getSquadraOrarioParts(row)?.start;
        const parsed = parseTimeToMinutes(start);
        if (parsed !== null) return parsed;
      }
    } catch (_) {
      // Prosegue con i campi compatibili usati dalle versioni precedenti.
    }
    const fallback = row?.oraInizio ?? row?.orarioInizio ?? row?.inizio ?? row?.start ?? "";
    return parseTimeToMinutes(fallback);
  }

  function getAssignedStartMinutes(assignments) {
    const starts = [];
    assignments.forEach((assignment) => {
      (assignment?.matchedRows || []).forEach(({ row }) => {
        const start = getRowStartMinutes(row);
        if (start !== null) starts.push(start);
      });
    });
    return starts.length ? Math.min(...starts) : null;
  }

  function getElapsedWorkMinutes(assignments) {
    const start = getAssignedStartMinutes(assignments);
    if (start === null) return null;
    const elapsed = Math.max(0, getRomeClockMinutes() - start);
    return elapsed > (8 * 60) ? elapsed - 60 : elapsed;
  }

  function formatHoursMinutes(minutes) {
    const safeMinutes = Math.max(0, Math.round(Number(minutes) || 0));
    return `${String(Math.floor(safeMinutes / 60)).padStart(2, "0")}:${String(safeMinutes % 60).padStart(2, "0")}`;
  }

  function isCurrentUserOperator(operatorName) {
    if (typeof doSquadraMemberAndUserMatch === "function") {
      try {
        return Boolean(doSquadraMemberAndUserMatch(operatorName));
      } catch (_) {
        // Prosegue con il controllo compatibile locale.
      }
    }

    const operator = normalizeIdentity(operatorName);
    if (!operator) return false;
    const userValues = [
      typeof currentUser !== "undefined" ? currentUser?.displayName : "",
      typeof currentUser !== "undefined" ? currentUser?.email : "",
      typeof currentUser !== "undefined" ? currentUser?.uid : ""
    ].map(normalizeIdentity).filter(Boolean);
    return userValues.some((value) => value === operator || value.includes(operator) || operator.includes(value));
  }

  function getSavedHoursForCurrentUser(assignments) {
    const reportsLoaded = typeof hoursReportsLoaded === "undefined" || Boolean(hoursReportsLoaded);
    const approvalsLoaded = typeof hoursApprovalsLoaded === "undefined" || Boolean(hoursApprovalsLoaded);
    if (!reportsLoaded || !approvalsLoaded) return { loaded: false, found: false, minutes: 0 };

    const dateKey = typeof getTodayDateKey === "function" ? String(getTodayDateKey() || "") : "";
    if (!dateKey) return { loaded: true, found: false, minutes: 0 };

    const assignedRows = new Map(assignments.map((assignment) => [
      String(assignment?.commessaId || ""),
      new Set((assignment?.matchedRows || []).map(({ squadraIndex }) => String(squadraIndex ?? "")))
    ]));

    const reports = typeof allHoursReports !== "undefined" && Array.isArray(allHoursReports)
      ? allHoursReports
      : [];
    const approvals = typeof allHoursApprovalRequests !== "undefined" && Array.isArray(allHoursApprovalRequests)
      ? allHoursApprovalRequests.filter((request) => String(request?.status || "").trim().toLowerCase() !== "rejected")
      : [];

    let sources = [...reports, ...approvals];
    if (typeof deduplicateHoursRecordsForDisplay === "function") {
      try {
        sources = deduplicateHoursRecordsForDisplay(sources);
      } catch (_) {
        // Mantiene le sorgenti originali se la deduplicazione non è disponibile.
      }
    }

    let minutes = 0;
    let found = false;
    sources.forEach((record) => {
      if (String(record?.date || "").trim() !== dateKey) return;
      (Array.isArray(record?.entries) ? record.entries : []).forEach((entry) => {
        const commessaId = String(entry?.commessaId || entry?.commessa?.id || "");
        const teamIndexes = assignedRows.get(commessaId);
        if (!teamIndexes) return;

        (Array.isArray(entry?.rows) ? entry.rows : []).forEach((row) => {
          const rowTeam = String(row?.squadraIndex ?? entry?.squadraIndex ?? "").trim();
          if (rowTeam && !teamIndexes.has(rowTeam)) return;
          if (!isCurrentUserOperator(row?.operatore || row?.operator || row?.nomeOperatore || "")) return;
          const hours = Number(row?.ore);
          if (!Number.isFinite(hours) || hours <= 0) return;
          found = true;
          minutes += Math.round(hours * 60);
        });
      });
    });

    return { loaded: true, found, minutes: Math.max(0, minutes) };
  }

  function normalizeVehicleValue(value) {
    if (value && typeof value === "object") {
      return String(value.codice || value.targa || value.nome || value.label || value.id || "").trim();
    }
    return String(value || "").trim();
  }

  function parseVehicles(value) {
    if (typeof parseMultiEntryValue === "function") {
      try {
        const parsed = parseMultiEntryValue(value);
        if (Array.isArray(parsed)) return parsed.map(normalizeVehicleValue).filter(Boolean);
      } catch (_) {
        // Prosegue con il parser locale.
      }
    }
    if (Array.isArray(value)) return value.map(normalizeVehicleValue).filter(Boolean);
    return String(value || "")
      .split(/[\n,;|]+/)
      .map(normalizeVehicleValue)
      .filter(Boolean);
  }

  function getAssignedVehicles(assignments) {
    const vehicles = new Map();
    const visitedRows = new Set();
    assignments.forEach((assignment) => {
      (assignment?.matchedRows || []).forEach(({ squadraIndex, row }) => {
        const rowKey = `${assignment?.commessaId || ""}:${squadraIndex ?? ""}`;
        if (visitedRows.has(rowKey)) return;
        visitedRows.add(rowKey);
        const rawVehicles = row?.mezzi ?? row?.mezzo ?? row?.veicoli ?? row?.mezziAssegnati ?? "";
        parseVehicles(rawVehicles).forEach((vehicle) => {
          const key = typeof normalizeSquadraMemberIdentity === "function"
            ? normalizeSquadraMemberIdentity(vehicle)
            : normalizeIdentity(vehicle);
          if (key && !vehicles.has(key)) vehicles.set(key, vehicle);
        });
      });
    });
    return [...vehicles.values()];
  }

  function getVehicleMeta(vehicle) {
    const code = String(vehicle || "").trim().toUpperCase();
    if (code.startsWith("MA")) return { icon: "🏗️", label: "Escavatore" };
    if (code.startsWith("A")) return { icon: "🚛", label: "Camion" };
    if (code.startsWith("T")) return { icon: "🚜", label: "Trattore grande" };
    if (code.startsWith("R")) return { icon: "🚜", label: "Trattorino" };
    return { icon: "🛠️", label: "Mezzo" };
  }

  function renderVehicleBadge(vehicle) {
    const meta = getVehicleMeta(vehicle);
    return `<span class="today-vehicle-badge" title="${escapeMarkup(meta.label)}"><span class="today-vehicle-icon" aria-hidden="true">${meta.icon}</span><span>${escapeMarkup(vehicle)}</span></span>`;
  }

  function installStyle() {
    if (document.getElementById("today-live-hours-vehicles-style")) return;
    const style = document.createElement("style");
    style.id = "today-live-hours-vehicles-style";
    style.textContent = `
      #today-mezzi-count{display:block!important;visibility:visible!important;opacity:1!important;width:100%!important;color:#172033!important;font-weight:900!important;text-align:center!important;line-height:1.15!important}
      @media(max-width:480px){#today-mezzi-count{font-size:.68rem!important}}
    `;
    document.head.appendChild(style);
  }

  function renderLiveHours(assignments) {
    const hoursButton = getLiveElement("today-hours-btn", typeof ui !== "undefined" ? ui.todayHoursBtn : null);
    const hoursCount = getLiveElement("today-hours-count", typeof ui !== "undefined" ? ui.todayHoursCount : null);
    if (!hoursButton || !hoursCount) return;

    const savedHours = getSavedHoursForCurrentUser(assignments);
    if (savedHours.loaded && savedHours.found) {
      const savedText = formatHoursMinutes(savedHours.minutes);
      hoursCount.textContent = savedText;
      hoursButton.classList.add("is-complete");
      hoursButton.dataset.liveCounter = "stopped";
      hoursButton.setAttribute("aria-label", `Ore inserite oggi: ${savedText}. Premi per visualizzarle o modificarle`);
      return;
    }

    hoursButton.classList.remove("is-complete");
    hoursButton.dataset.liveCounter = "running";
    const elapsedMinutes = getElapsedWorkMinutes(assignments);
    const hoursText = elapsedMinutes === null ? "--:--" : formatHoursMinutes(elapsedMinutes);
    hoursCount.textContent = hoursText;
    hoursButton.setAttribute("aria-label", elapsedMinutes === null
      ? "Inserisci ore. Nessun orario di inizio assegnato alla squadra"
      : `Inserisci ore. Tempo trascorso dall'inizio turno: ${hoursText}`);
  }

  function renderAssignedVehicles(assignments) {
    const vehicleButton = getLiveElement("today-mezzi-btn", typeof ui !== "undefined" ? ui.todayMezziBtn : null);
    const vehicleCount = getLiveElement("today-mezzi-count", typeof ui !== "undefined" ? ui.todayMezziCount : null);
    if (!vehicleButton || !vehicleCount) return;

    const vehicles = getAssignedVehicles(assignments);
    vehicleCount.hidden = false;
    vehicleCount.removeAttribute("aria-hidden");
    vehicleCount.textContent = vehicles.length ? vehicles.join("  ") : "MEZZI";
    vehicleButton.disabled = false;
    vehicleButton.setAttribute("aria-label", vehicles.length
      ? `Mezzi assegnati oggi: ${vehicles.join(", ")}`
      : "Apri i mezzi assegnati oggi");
  }

  function applyTodayLiveSummary() {
    installStyle();
    const assignments = getAssignmentsSafe();
    renderLiveHours(assignments);
    renderAssignedVehicles(assignments);
  }

  if (typeof renderTodaySummary === "function") {
    const originalRenderTodaySummary = renderTodaySummary;
    renderTodaySummary = function renderTodaySummaryWithLiveHoursAndVehicles(...args) {
      const result = originalRenderTodaySummary.apply(this, args);
      queueMicrotask(applyTodayLiveSummary);
      return result;
    };
  }

  function clearClockSchedule() {
    if (minuteTimeout !== null) window.clearTimeout(minuteTimeout);
    if (minuteInterval !== null) window.clearInterval(minuteInterval);
    minuteTimeout = null;
    minuteInterval = null;
  }

  function refreshSummary() {
    if (typeof renderTodaySummary === "function") renderTodaySummary();
    else applyTodayLiveSummary();
  }

  function startClockSchedule() {
    clearClockSchedule();
    refreshSummary();
    const waitForNextMinute = 60000 - (Date.now() % 60000) + 75;
    minuteTimeout = window.setTimeout(() => {
      refreshSummary();
      minuteInterval = window.setInterval(refreshSummary, 60000);
    }, waitForNextMinute);
  }

  document.addEventListener("visibilitychange", () => {
    if (document.visibilityState === "visible") startClockSchedule();
    else clearClockSchedule();
  });
  window.addEventListener("pageshow", startClockSchedule);
  window.addEventListener("pagehide", clearClockSchedule);

  if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", startClockSchedule, { once: true });
  } else {
    startClockSchedule();
  }
})();
