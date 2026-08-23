(() => {
  "use strict";

  if (window.HeraRecommendedPlants?.installed) return;

  const CONFIG = Object.freeze({
    start: {
      name: "Avola Coop",
      address: "Via Galliera 14/A, 40013 Castel Maggiore (BO)",
      lat: 44.5790,
      lng: 11.3635
    },
    teamSize: 3,
    machine: "trincia",
    minutesPer100Mq: 8,
    setupMinutesPerPlant: 10,
    contingencyPct: 15,
    roadFactor: 1.22,
    averageRoadSpeedKmh: 52,
    planningMinutes: 8 * 60,
    maxVisible: 8
  });

  const STORAGE_ROUTE_DAY = "heraRecommendedPlantsRouteDay";
  const STORAGE_ROUTE_STARTED = "heraRecommendedPlantsRouteStarted";
  const state = { open: false, lastPlan: [], originMode: "auto" };

  const numberValue = value => {
    if (value == null || value === "") return null;
    const normalized = String(value).trim().replace(/\s/g, "").replace(/\./g, "").replace(",", ".");
    const parsed = Number(normalized);
    return Number.isFinite(parsed) ? parsed : null;
  };

  const readGlobal = name => {
    try {
      if (name === "currentImpianti" && typeof currentImpianti !== "undefined") return currentImpianti;
    } catch (_) {}
    return window[name];
  };

  const normalizeText = value => String(value ?? "").trim();

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

  function coordsOf(item) {
    const lat = numberValue(item?.latitudine ?? item?.lat ?? item?.latitude ?? item?.gpsY ?? item?.y);
    const lng = numberValue(item?.longitudine ?? item?.lng ?? item?.lon ?? item?.longitude ?? item?.gpsX ?? item?.x);
    if (!Number.isFinite(lat) || !Number.isFinite(lng)) return null;
    if (Math.abs(lat) > 90 || Math.abs(lng) > 180) return null;
    return { lat, lng };
  }

  function areaMqOf(item) {
    const candidates = [
      item?.areaMq, item?.mq, item?.superficieMq, item?.metriQuadri,
      item?.superficie, item?.quantitaMq, item?.quantita
    ];
    for (const value of candidates) {
      const parsed = numberValue(value);
      if (Number.isFinite(parsed) && parsed > 0) return parsed;
    }
    return 0;
  }

  function isDone(item) {
    const status = normalizeText(item?.stato ?? item?.status).toUpperCase();
    return Boolean(item?.done) || ["FATTO", "DONE", "COMPLETATO"].includes(status);
  }

  function titleOf(item) {
    return normalizeText(item?.denominazione ?? item?.nome ?? item?.impianto ?? item?.name) || "Impianto";
  }

  function municipalityOf(item) {
    return normalizeText(item?.comune ?? item?.municipality ?? item?.citta);
  }

  function idOf(item) {
    return normalizeText(item?.id ?? item?.impiantoId ?? item?.idSap ?? item?.idSAP ?? titleOf(item));
  }

  function haversineKm(a, b) {
    const r = 6371;
    const toRad = value => value * Math.PI / 180;
    const dLat = toRad(b.lat - a.lat);
    const dLng = toRad(b.lng - a.lng);
    const lat1 = toRad(a.lat);
    const lat2 = toRad(b.lat);
    const h = Math.sin(dLat / 2) ** 2 + Math.cos(lat1) * Math.cos(lat2) * Math.sin(dLng / 2) ** 2;
    return 2 * r * Math.asin(Math.sqrt(h));
  }

  function roadEstimateKm(a, b) {
    return haversineKm(a, b) * CONFIG.roadFactor;
  }

  function travelMinutes(km) {
    return Math.max(3, Math.round((km / CONFIG.averageRoadSpeedKmh) * 60));
  }

  function workMinutes(areaMq) {
    if (!areaMq) return CONFIG.setupMinutesPerPlant + 30;
    const productive = (areaMq / 100) * CONFIG.minutesPer100Mq;
    const base = CONFIG.setupMinutesPerPlant + productive;
    return Math.round(base * (1 + CONFIG.contingencyPct / 100));
  }

  function formatMinutes(total) {
    const min = Math.max(0, Math.round(total || 0));
    const h = Math.floor(min / 60);
    const m = min % 60;
    if (!h) return `${m} min`;
    if (!m) return `${h} h`;
    return `${h} h ${m} min`;
  }

  function formatKm(km) {
    return `${Number(km || 0).toLocaleString("it-IT", { maximumFractionDigits: 1 })} km`;
  }

  function getCurrentPosition() {
    return new Promise(resolve => {
      if (!navigator.geolocation) return resolve(null);
      navigator.geolocation.getCurrentPosition(
        pos => resolve({ lat: pos.coords.latitude, lng: pos.coords.longitude }),
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

  function buildGreedyPlan(items, origin) {
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
      const work = workMinutes(next.areaMq);
      cumulative += drive + work;
      plan.push({
        ...next,
        km,
        driveMinutes: drive,
        workMinutes: work,
        cumulativeMinutes: cumulative,
        fitsDay: cumulative <= CONFIG.planningMinutes
      });
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
    if (list) card.insertBefore(panel, list);
    else card.appendChild(panel);
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
    if (todo?.nextSibling) tabs.insertBefore(button, todo.nextSibling);
    else tabs.appendChild(button);
    button.addEventListener("click", () => {
      state.open = !state.open;
      state.originMode = "auto";
      render();
    });
  }

  function escapeHtml(value) {
    return String(value ?? "").replace(/[&<>\"']/g, char => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#039;" }[char]));
  }

  function navigationUrl(entry) {
    const { lat, lng } = entry.coords;
    return `https://www.google.com/maps/dir/?api=1&destination=${encodeURIComponent(`${lat},${lng}`)}&travelmode=driving`;
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

    panel.innerHTML = `<div class="recommended-loading">Calcolo gli impianti consigliati…</div>`;
    const current = readGlobal("currentImpianti");
    const items = Array.isArray(current) ? current : [];
    const origin = await resolveOrigin();
    const plan = buildGreedyPlan(items, origin);
    state.lastPlan = plan;

    if (!items.length) {
      panel.innerHTML = `<div class="recommended-empty"><strong>Nessun impianto disponibile.</strong><span>Attendi il caricamento della commessa e riprova.</span></div>`;
      return;
    }
    if (!plan.length) {
      panel.innerHTML = `<div class="recommended-empty"><strong>Nessun impianto da pianificare.</strong><span>Gli impianti da fare devono avere coordinate valide.</span></div>`;
      return;
    }

    const feasible = plan.filter(entry => entry.fitsDay);
    const visible = (feasible.length ? feasible : plan.slice(0, 1)).slice(0, CONFIG.maxVisible);
    const totalDrive = visible.reduce((sum, x) => sum + x.driveMinutes, 0);
    const totalWork = visible.reduce((sum, x) => sum + x.workMinutes, 0);
    const originLabel = origin.mode === "avola" ? "Avola Coop · primo impianto" : "Posizione attuale · squadra già in giro";

    panel.innerHTML = `
      <header class="recommended-head">
        <div>
          <span class="recommended-kicker">PIANIFICAZIONE GIORNATA</span>
          <h3>✨ Impianti consigliati</h3>
          <p>${escapeHtml(originLabel)}</p>
        </div>
        <button id="recommended-close-btn" class="btn btn-small" type="button">Chiudi</button>
      </header>
      <div class="recommended-summary">
        <span><strong>${visible.length}</strong> consigliati</span>
        <span>🚐 ${formatMinutes(totalDrive)}</span>
        <span>🌿 ${formatMinutes(totalWork)}</span>
      </div>
      <p class="recommended-rule">Stima sfalcio: ${CONFIG.teamSize} persone + ${CONFIG.machine}, ${CONFIG.minutesPer100Mq} min/100 m², margine ${CONFIG.contingencyPct}%.</p>
      <div class="recommended-origin-actions">
        <button id="recommended-origin-avola" class="btn btn-small ${origin.mode === "avola" ? "btn-primary" : ""}" type="button">🏠 Primo giro da Avola</button>
        <button id="recommended-origin-live" class="btn btn-small ${origin.mode === "position" ? "btn-primary" : ""}" type="button">📍 Sono già in giro</button>
      </div>
      <div class="recommended-list">
        ${visible.map((entry, index) => `
          <article class="recommended-item" data-recommended-id="${escapeHtml(idOf(entry.item))}">
            <div class="recommended-rank">${index + 1}</div>
            <div class="recommended-main">
              <strong>${escapeHtml(titleOf(entry.item))}</strong>
              <span>${escapeHtml(municipalityOf(entry.item))}</span>
              <small>🚐 ${formatKm(entry.km)} · ${formatMinutes(entry.driveMinutes)} &nbsp; 🌿 ${entry.areaMq ? `${Math.round(entry.areaMq).toLocaleString("it-IT")} m² · ` : ""}${formatMinutes(entry.workMinutes)}</small>
            </div>
            <a class="btn recommended-nav-btn" href="${navigationUrl(entry)}" target="_blank" rel="noopener" data-recommended-nav="1">NAVIGA</a>
          </article>
        `).join("")}
      </div>
      <div class="recommended-footer">
        <button id="recommended-start-route" class="btn btn-primary" type="button">AVVIA GIRO CONSIGLIATO</button>
        <small>La lista normale resta ordinata per distanza come prima.</small>
      </div>`;

    panel.querySelector("#recommended-close-btn")?.addEventListener("click", () => {
      state.open = false;
      render();
    });
    panel.querySelector("#recommended-origin-avola")?.addEventListener("click", () => {
      state.originMode = "avola";
      render();
    });
    panel.querySelector("#recommended-origin-live")?.addEventListener("click", () => {
      state.originMode = "position";
      render();
    });
    panel.querySelector("#recommended-start-route")?.addEventListener("click", () => {
      markRouteStartedToday();
      state.originMode = "position";
      const btn = panel.querySelector("#recommended-start-route");
      if (btn) btn.textContent = "✓ GIRO AVVIATO";
    });
    panel.querySelectorAll("[data-recommended-nav='1']").forEach(link => {
      link.addEventListener("click", markRouteStartedToday, { passive: true });
    });
  }

  function install() {
    ensureButton();
    ensurePanel();
    const page = document.getElementById("impianti-page");
    if (page) {
      new MutationObserver(() => {
        ensureButton();
        ensurePanel();
      }).observe(page, { childList: true, subtree: true });
    }
  }

  if (document.readyState === "loading") document.addEventListener("DOMContentLoaded", install, { once: true });
  else install();

  window.HeraRecommendedPlants = {
    installed: true,
    version: "1.0.0",
    config: CONFIG,
    open: () => { state.open = true; state.originMode = "auto"; return render(); },
    refresh: render,
    markRouteStartedToday,
    getState: () => ({ ...state, lastPlan: state.lastPlan.slice() })
  };
})();