"use strict";

(() => {
  const ROME_TIME_ZONE = "Europe/Rome";
  let minuteTimeout = null;
  let minuteInterval = null;

  const escapeMarkup = (value) => String(value ?? "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#039;");

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
    const safeMinutes = Math.max(0, Math.floor(Number(minutes) || 0));
    return `${String(Math.floor(safeMinutes / 60)).padStart(2, "0")}:${String(safeMinutes % 60).padStart(2, "0")}`;
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
            : vehicle.toLocaleLowerCase("it-IT").replace(/\s+/g, " ").trim();
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
      #today-mezzi-action.today-mezzi-inline{display:flex;align-items:center;justify-content:center;flex-wrap:wrap;gap:3px;max-width:100%;line-height:1.05}
      #today-mezzi-action.today-mezzi-inline .today-vehicle-badge{display:inline-flex;align-items:center;gap:2px;min-height:20px;padding:1px 5px;border:1px solid #cfd9e6;border-radius:7px;background:#f8fafc;color:#172033;font-size:.72rem;font-weight:800;line-height:1;white-space:nowrap;text-transform:none}
      #today-mezzi-action.today-mezzi-inline .today-vehicle-icon{font-size:.78rem;line-height:1}
      @media(max-width:480px){#today-mezzi-action.today-mezzi-inline{gap:2px}#today-mezzi-action.today-mezzi-inline .today-vehicle-badge{font-size:.66rem;padding:1px 4px}}
    `;
    document.head.appendChild(style);
  }

  function renderLiveHours(assignments) {
    const hoursButton = (typeof ui !== "undefined" && ui.todayHoursBtn) || document.getElementById("today-hours-btn");
    const hoursCount = (typeof ui !== "undefined" && ui.todayHoursCount) || document.getElementById("today-hours-count");
    if (!hoursButton || !hoursCount) return;

    // Quando le ore sono già state inserite, mantiene il valore registrato e ferma il conteggio visivo.
    if (hoursButton.classList.contains("is-complete")) return;

    const elapsedMinutes = getElapsedWorkMinutes(assignments);
    const hoursText = elapsedMinutes === null ? "--:--" : formatHoursMinutes(elapsedMinutes);
    hoursCount.textContent = hoursText;
    hoursButton.setAttribute("aria-label", elapsedMinutes === null
      ? "Inserisci ore. Nessun orario di inizio assegnato alla squadra"
      : `Inserisci ore. Tempo trascorso dall'inizio turno: ${hoursText}`);
  }

  function renderAssignedVehicles(assignments) {
    const vehicleButton = (typeof ui !== "undefined" && ui.todayMezziBtn) || document.getElementById("today-mezzi-btn");
    const vehicleAction = document.getElementById("today-mezzi-action");
    const vehicleCount = (typeof ui !== "undefined" && ui.todayMezziCount) || document.getElementById("today-mezzi-count");
    if (!vehicleButton || !vehicleAction || !vehicleCount) return;

    const vehicles = getAssignedVehicles(assignments);
    vehicleAction.classList.toggle("today-mezzi-inline", vehicles.length > 0);
    if (vehicles.length) {
      vehicleAction.innerHTML = vehicles.map(renderVehicleBadge).join("");
      vehicleCount.textContent = "";
      vehicleButton.disabled = false;
      vehicleButton.setAttribute("aria-label", `Mezzi assegnati oggi: ${vehicles.join(", ")}`);
      return;
    }

    vehicleAction.textContent = "NESSUN MEZZO";
    vehicleCount.textContent = "";
    vehicleButton.disabled = true;
    vehicleButton.setAttribute("aria-label", "Nessun mezzo assegnato oggi");
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
