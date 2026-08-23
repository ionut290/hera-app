(() => {
  "use strict";
  if (window.HeraRecommendedPlants?.installed) return;

  const CONFIG = Object.freeze({
    start: { name: "Avola Coop", address: "Via Galliera 14/A, 40013 Castel Maggiore (BO)", lat: 44.5790, lng: 11.3635 },
    baselineTeamSize: 3,
    machine: "trincia",
    baselineMinutesPer100Mq: 8,
    setupMinutesPerPlant: 10,
    contingencyPct: 15,
    roadFactor: 1.22,
    averageRoadSpeedKmh: 52,
    planningMinutes: 8 * 60,
    maxVisible: 8
  });

  const STORAGE_ROUTE_DAY = "heraRecommendedPlantsRouteDay";
  const STORAGE_ROUTE_STARTED = "heraRecommendedPlantsRouteStarted";
  const state = { open: false, lastPlan: [], originMode: "auto", teamSize: CONFIG.baselineTeamSize };

  const numberValue = value => {
    if (value == null || value === "") return null;
    const parsed = Number(String(value).trim().replace(/\s/g, "").replace(/\./g, "").replace(",", "."));
    return Number.isFinite(parsed) ? parsed : null;
  };
  const normalizeText = value => String(value ?? "").trim();
  const readGlobal = name => {
    try {
      if (name === "currentImpianti" && typeof currentImpianti !== "undefined") return currentImpianti;
      if (name === "currentSquadre" && typeof currentSquadre !== "undefined") return currentSquadre;
    } catch (_) {}
    return window[name];
  };

  function todayKey() {
    const now = new Date();
    return `${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, "0")}-${String(now.getDate()).padStart(2, "0")}`;
  }
  function hasRouteStartedToday() {
    try {
      return localStorage.getItem(STORAGE_ROUTE_DAY) === todayKey()
        && localStorage.getItem(STORAGE_ROUTE_STARTED) === "1";
    } catch (_) {
      return false;
    }
  }
  function markRouteStartedToday() {
    try {
      localStorage.setItem(STORAGE_ROUTE_DAY, todayKey());
      localStorage.setItem(STORAGE_ROUTE_STARTED, "1");
    } catch (_) {}
  }
  function selectedCommessa() {
    try { return String(selectedCommessaId || window.selectedCommessaId || "").trim(); }
    catch (_) { return String(window.selectedCommessaId || "").trim(); }
  }
  function operatorCountFromSquad(squad) {
    const arrays = [squad?.operatori, squad?.operators, squad?.membri, squad?.members, squad?.persone, squad?.componenti, squad?.utenti];
    for (const value of arrays) {
      if (Array.isArray(value) && value.length) return value.filter(Boolean).length;
    }
    for (const value of [squad?.numeroOperatori, squad?.operatorCount, squad?.teamSize, squad?.numeroPersone]) {
      const parsed = numberValue(value);
      if (parsed > 0) return Math.round(parsed);
    }
    return 0;
  }
  function detectTeamSize() {
    const commessaId = selectedCommessa();
    const today = todayKey();
    const sources = [readGlobal("currentSquadre"), window.squadre, window.squadreOggi, window.todaySquads, window.squadreByCommessa];
    for (const source of sources) {
      let rows = [];
      if (Array.isArray(source)) rows = source;
      else if (source instanceof Map) rows = [...(source.get(commessaId) || [])];
      else if (source && typeof source === "object") rows = Array.isArray(source[commessaId]) ? source[commessaId] : Object.values(source);
      for (const squad of rows) {
        const squadCommessa = String(squad?.commessaId ?? squad?.idCommessa ?? squad?.commessa ?? "").trim();
        const date = String(squad?.data ?? squad?.date ?? squad?.giorno ?? "").slice(0, 10);
        if (squadCommessa && commessaId && squadCommessa !== commessaId) continue;
        if (date && date !== today) continue;
        const count = operatorCountFromSquad(squad);
        if (count > 0) return count;
      }
    }
    return CONFIG.baselineTeamSize;
  }

  function firstValue(item, names) {
    for (const name of names) {
      const value = item?.[name];
      if (value !== undefined && value !== null && String(value).trim() !== "") return value;
    }
    return "";
  }

  function coordsOf(item) {
    if (!item || typeof item !== "object") return null;

    const latitudeRaw = firstValue(item, [
      "latitudine", "lat", "latitude", "gpsY", "GPSY", "gps_y", "coordinateY", "coordY",
      "y", "latitudineGps", "latGps", "gpsLat", "gpsLatitude"
    ]);
    const longitudeRaw = firstValue(item, [
      "longitudine", "lng", "lon", "longitude", "gpsX", "GPSX", "gps_x", "coordinateX", "coordX",
      "x", "longitudineGps", "lngGps", "lonGps", "gpsLng", "gpsLon", "gpsLongitude"
    ]);

    const repair = window.HeraCoordinateRepair;
    if (repair?.diagnose) {
      const diagnosed = repair.diagnose(latitudeRaw, longitudeRaw);
      if (diagnosed?.valid) return { lat: diagnosed.latitude, lng: diagnosed.longitude };
    }

    let lat = numberValue(latitudeRaw);
    let lng = numberValue(longitudeRaw);
    if (Number.isFinite(lat) && Number.isFinite(lng)) {
      const directItaly = lat >= 35 && lat <= 48.8 && lng >= 5 && lng <= 20;
      const swappedItaly = lng >= 35 && lng <= 48.8 && lat >= 5 && lat <= 20;
      if (!directItaly && swappedItaly) [lat, lng] = [lng, lat];
      if (Math.abs(lat) <= 90 && Math.abs(lng) <= 180 && lat !== 0 && lng !== 0) return { lat, lng };
    }

    const combinedCandidates = [
      item.coordinate, item.coordinates, item.coordinateGps, item.coordinateGPS, item.coordinate_gps,
      item.gps, item.GPS, item.posizioneGps, item.posizioneGPS, item.coordinateImpianto,
      item["Coordinate GPS"], item["COORDINATE GPS"], item["Coordinate GPS(X)/GPS(Y)"],
      item["GPS(X)/GPS(Y)"], item["GPS X/Y"]
    ];
    for (const combined of combinedCandidates) {
      if (combined == null || String(combined).trim() === "") continue;
      if (repair?.diagnose) {
        const diagnosed = repair.diagnose(combined, "");
        if (diagnosed?.valid) return { lat: diagnosed.latitude, lng: diagnosed.longitude };
      }
      const matches = String(combined).match(/[-+]?\d{1,3}(?:[.,]\d+)?/g) || [];
      if (matches.length >= 2) {
        let first = numberValue(matches[0]);
        let second = numberValue(matches[1]);
        if (!Number.isFinite(first) || !Number.isFinite(second)) continue;
        const directItaly = first >= 35 && first <= 48.8 && second >= 5 && second <= 20;
        const swappedItaly = second >= 35 && second <= 48.8 && first >= 5 && first <= 20;
        if (!directItaly && swappedItaly) [first, second] = [second, first];
        if (Math.abs(first) <= 90 && Math.abs(second) <= 180) return { lat: first, lng: second };
      }
    }

    const normalizedEntries = Object.entries(item).map(([key, value]) => [
      String(key).toLowerCase().replace(/[^a-z0-9]/g, ""),
      value
    ]);
    const dynamicLat = normalizedEntries.find(([key]) => /^(lat|latitude|latitudine|gpsy|coordinatey|coordy)$/.test(key))?.[1];
    const dynamicLng = normalizedEntries.find(([key]) => /^(lng|lon|longitude|longitudine|gpsx|coordinatex|coordx)$/.test(key))?.[1];
    if (dynamicLat != null || dynamicLng != null) {
      if (repair?.diagnose) {
        const diagnosed = repair.diagnose(dynamicLat, dynamicLng);
        if (diagnosed?.valid) return { lat: diagnosed.latitude, lng: diagnosed.longitude };
      }
      const parsedLat = numberValue(dynamicLat);
      const parsedLng = numberValue(dynamicLng);
      if (Number.isFinite(parsedLat) && Number.isFinite(parsedLng)) return { lat: parsedLat, lng: parsedLng };
    }

    return null;
  }

  function areaMqOf(item) {
    for (const value of [item?.areaMq, item?.mq, item?.superficieMq, item?.metriQuadri, item?.superficie, item?.quantitaMq, item?.quantita]) {
      const parsed = numberValue(value);
      if (parsed > 0) return parsed;
    }
    return 0;
  }
  function isDone(item) {
    const status = normalizeText(item?.stato ?? item?.status).toUpperCase();
    return Boolean(item?.done) || ["FATTO", "DONE", "COMPLETATO"].includes(status);
  }
  const titleOf = item => normalizeText(item?.denominazione ?? item?.nome ?? item?.impianto ?? item?.name) || "Impianto";
  const municipalityOf = item => normalizeText(item?.comune ?? item?.municipality ?? item?.citta);
  const idOf = item => normalizeText(item?.id ?? item?.impiantoId ?? item?.idSap ?? item?.idSAP ?? titleOf(item));

  function haversineKm(a, b) {
    const radius = 6371;
    const rad = value => value * Math.PI / 180;
    const dLat = rad(b.lat - a.lat);
    const dLng = rad(b.lng - a.lng);
    const h = Math.sin(dLat / 2) ** 2 + Math.cos(rad(a.lat)) * Math.cos(rad(b.lat)) * Math.sin(dLng / 2) ** 2;
    return 2 * radius * Math.asin(Math.sqrt(h));
  }
  const roadEstimateKm = (a, b) => haversineKm(a, b) * CONFIG.roadFactor;
  const travelMinutes = km => Math.max(3, Math.round(km / CONFIG.averageRoadSpeedKmh * 60));
  const minutesPer100Mq = teamSize => CONFIG.baselineMinutesPer100Mq * CONFIG.baselineTeamSize / Math.max(1, teamSize);
  function workMinutes(area, teamSize) {
    if (!area) return CONFIG.setupMinutesPerPlant + 30;
    const productive = area / 100 * minutesPer100Mq(teamSize);
    return Math.round((CONFIG.setupMinutesPerPlant + productive) * (1 + CONFIG.contingencyPct / 100));
  }
  function formatMinutes(total) {
    const min = Math.max(0, Math.round(total || 0));
    const hours = Math.floor(min / 60);
    const minutes = min % 60;
    return !hours ? `${minutes} min` : !minutes ? `${hours} h` : `${hours} h ${minutes} min`;
  }
  const formatKm = km => `${Number(km || 0).toLocaleString("it-IT", { maximumFractionDigits: 1 })} km`;

  function getCurrentPosition() {
    return new Promise(resolve => {
      if (!navigator.geolocation) return resolve(null);
      navigator.geolocation.getCurrentPosition(
        position => resolve({ lat: position.coords.latitude, lng: position.coords.longitude }),
        () => resolve(null),
        { enableHighAccuracy: true, timeout: 4500, maximumAge: 120000 }
      );
    });
  }
  async function resolveOrigin() {
    if (state.originMode === "avola") return { ...CONFIG.start, mode: "avola" };
    if (state.originMode === "position") {
      const live = await getCurrentPosition();
      return live ? { ...live, name: "Posizione attuale", mode: "position" } : { ...CONFIG.start, mode: "avola" };
    }
    if (!hasRouteStartedToday()) return { ...CONFIG.start, mode: "avola" };
    const live = await getCurrentPosition();
    return live ? { ...live, name: "Posizione attuale", mode: "position" } : { ...CONFIG.start, mode: "avola" };
  }

  function buildPlan(items, origin, teamSize) {
    const remaining = items
      .filter(item => !isDone(item))
      .map(item => ({ item, coords: coordsOf(item), areaMq: areaMqOf(item) }))
      .filter(entry => entry.coords);
    const plan = [];
    let cursor = { lat: origin.lat, lng: origin.lng };
    let cumulative = 0;
    while (remaining.length) {
      remaining.sort((a, b) => roadEstimateKm(cursor, a.coords) - roadEstimateKm(cursor, b.coords));
      const next = remaining.shift();
      const km = roadEstimateKm(cursor, next.coords);
      const drive = travelMinutes(km);
      const work = workMinutes(next.areaMq, teamSize);
      cumulative += drive + work;
      plan.push({ ...next, km, driveMinutes: drive, workMinutes: work, cumulativeMinutes: cumulative, fitsDay: cumulative <= CONFIG.planningMinutes });
      cursor = next.coords;
    }
    return plan;
  }

  function ensurePanel() {
    const card = document.getElementById("impianti-card");
    if (!card) return null;
    let panel = document.getElementById("recommended-plants-panel");
    if (panel) return panel;
    panel = document.createElement("section");
    panel.id = "recommended-plants-panel";
    panel.className = "recommended-plants-panel hidden";
    panel.setAttribute("aria-live", "polite");
    const list = document.getElementById("impianti-lista");
    list ? card.insertBefore(panel, list) : card.appendChild(panel);
    return panel;
  }
  function ensureButton() {
    const tabs = document.querySelector("#impianti-card .view-tabs");
    if (!tabs || document.getElementById("recommended-plants-btn")) return;
    const button = document.createElement("button");
    button.id = "recommended-plants-btn";
    button.className = "btn recommended-plants-btn";
    button.type = "button";
    button.innerHTML = "✨ Impianti consigliati";
    const todo = document.getElementById("view-todo-btn");
    todo?.nextSibling ? tabs.insertBefore(button, todo.nextSibling) : tabs.appendChild(button);
    button.addEventListener("click", () => {
      state.open = !state.open;
      state.originMode = "auto";
      render();
    });
  }
  const escapeHtml = value => String(value ?? "").replace(/[&<>\"']/g, char => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#039;" }[char]));
  function navigationUrl(entry) {
    return `https://www.google.com/maps/dir/?api=1&destination=${encodeURIComponent(`${entry.coords.lat},${entry.coords.lng}`)}&travelmode=driving`;
  }

  async function render() {
    ensureButton();
    const panel = ensurePanel();
    const button = document.getElementById("recommended-plants-btn");
    if (!panel || !button) return;
    panel.classList.toggle("hidden", !state.open);
    button.classList.toggle("btn-primary", state.open);
    button.setAttribute("aria-pressed", String(state.open));
    if (!state.open) return;

    panel.innerHTML = '<div class="recommended-loading">Calcolo gli impianti consigliati…</div>';
    const current = readGlobal("currentImpianti");
    const items = Array.isArray(current) ? current : [];
    const origin = await resolveOrigin();
    const teamSize = detectTeamSize();
    const plan = buildPlan(items, origin, teamSize);
    state.teamSize = teamSize;
    state.lastPlan = plan;

    if (!items.length) {
      panel.innerHTML = '<div class="recommended-empty"><strong>Nessun impianto disponibile.</strong><span>Attendi il caricamento della commessa e riprova.</span></div>';
      return;
    }
    if (!plan.length) {
      panel.innerHTML = '<div class="recommended-empty"><strong>Nessun impianto da pianificare.</strong><span>Non riesco a leggere le coordinate degli impianti di questa commessa.</span></div>';
      return;
    }

    const feasible = plan.filter(entry => entry.fitsDay);
    const visible = (feasible.length ? feasible : plan.slice(0, 1)).slice(0, CONFIG.maxVisible);
    const totalDrive = visible.reduce((sum, entry) => sum + entry.driveMinutes, 0);
    const totalWork = visible.reduce((sum, entry) => sum + entry.workMinutes, 0);
    const originLabel = origin.mode === "avola" ? "Avola Coop · primo impianto" : "Posizione attuale · squadra già in giro";
    const rate = minutesPer100Mq(teamSize);

    panel.innerHTML = `<header class="recommended-head"><div><span class="recommended-kicker">PIANIFICAZIONE GIORNATA</span><h3>✨ Impianti consigliati</h3><p>${escapeHtml(originLabel)}</p></div><button id="recommended-close-btn" class="btn btn-small" type="button">Chiudi</button></header><div class="recommended-summary"><span><strong>${visible.length}</strong> consigliati</span><span>👥 ${teamSize} persone</span><span>🚐 ${formatMinutes(totalDrive)}</span><span>🌿 ${formatMinutes(totalWork)}</span></div><p class="recommended-rule">Stima automatica sulla squadra di oggi: ${teamSize} persone + ${CONFIG.machine}, ${rate.toLocaleString("it-IT", { maximumFractionDigits: 1 })} min/100 m², margine ${CONFIG.contingencyPct}%.</p><div class="recommended-origin-actions"><button id="recommended-origin-avola" class="btn btn-small ${origin.mode === "avola" ? "btn-primary" : ""}" type="button">🏠 Primo giro da Avola</button><button id="recommended-origin-live" class="btn btn-small ${origin.mode === "position" ? "btn-primary" : ""}" type="button">📍 Sono già in giro</button></div><div class="recommended-list">${visible.map((entry, index) => `<article class="recommended-item" data-recommended-id="${escapeHtml(idOf(entry.item))}"><div class="recommended-rank">${index + 1}</div><div class="recommended-main"><strong>${escapeHtml(titleOf(entry.item))}</strong><span>${escapeHtml(municipalityOf(entry.item))}</span><small>🚐 ${formatKm(entry.km)} · ${formatMinutes(entry.driveMinutes)} &nbsp; 🌿 ${entry.areaMq ? `${Math.round(entry.areaMq).toLocaleString("it-IT")} m² · ` : ""}${formatMinutes(entry.workMinutes)}</small></div><a class="btn recommended-nav-btn" href="${navigationUrl(entry)}" target="_blank" rel="noopener" data-recommended-nav="1">NAVIGA</a></article>`).join("")}</div><div class="recommended-footer"><button id="recommended-start-route" class="btn btn-primary" type="button">AVVIA GIRO CONSIGLIATO</button><small>La lista normale resta ordinata per distanza come prima.</small></div>`;

    panel.querySelector("#recommended-close-btn")?.addEventListener("click", () => { state.open = false; render(); });
    panel.querySelector("#recommended-origin-avola")?.addEventListener("click", () => { state.originMode = "avola"; render(); });
    panel.querySelector("#recommended-origin-live")?.addEventListener("click", () => { state.originMode = "position"; render(); });
    panel.querySelector("#recommended-start-route")?.addEventListener("click", () => {
      markRouteStartedToday();
      state.originMode = "position";
      const currentButton = panel.querySelector("#recommended-start-route");
      if (currentButton) currentButton.textContent = "✓ GIRO AVVIATO";
    });
    panel.querySelectorAll("[data-recommended-nav='1']").forEach(link => link.addEventListener("click", markRouteStartedToday, { passive: true }));
  }

  function install() {
    ensureButton();
    ensurePanel();
    const page = document.getElementById("impianti-page");
    if (page) new MutationObserver(() => { ensureButton(); ensurePanel(); }).observe(page, { childList: true, subtree: true });
  }

  if (document.readyState === "loading") document.addEventListener("DOMContentLoaded", install, { once: true });
  else install();

  window.HeraRecommendedPlants = {
    installed: true,
    version: "1.2.0",
    config: CONFIG,
    open: () => { state.open = true; state.originMode = "auto"; return render(); },
    refresh: render,
    markRouteStartedToday,
    getState: () => ({ ...state, lastPlan: state.lastPlan.slice() })
  };
})();