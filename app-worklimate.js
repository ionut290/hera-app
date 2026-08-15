"use strict";
(function installVargaWorklimateModule(global) {
  if (global.VargaWorklimateModule) return;
  const api = {};
  function normalizeWorklimateLevel(value) {
    const key = normalizeWeatherAlertKey(value);
    if (key.includes("ross") || key.includes("alto") || key.includes("emerg")) return "rosso";
    if (key.includes("aranc") || key.includes("medio") || key.includes("moderat")) return "arancione";
    if (key.includes("giall") || key.includes("atten") || key.includes("basso")) return "giallo";
    return "verde";
  }
  api.normalizeWorklimateLevel = normalizeWorklimateLevel;
  function getMostSevereWorklimateRisk(items = []) {
    return [...items].sort((a, b) => {
      const severity = (WEATHER_ALERT_PRIORITY[normalizeWorklimateLevel(b.riskLevel || b.livello)] || 0) - (WEATHER_ALERT_PRIORITY[normalizeWorklimateLevel(a.riskLevel || a.livello)] || 0);
      return severity || getRiskTimestampMs(b.updatedAt || b.forecastAt) - getRiskTimestampMs(a.updatedAt || a.forecastAt);
    })[0] || null;
  }
  api.getMostSevereWorklimateRisk = getMostSevereWorklimateRisk;
  function getWorklimateRiskForDateHour(items = [], dateKey = getActiveSquadreDateKey(), hour = 12) {
    const matchingItems = items.filter((item) => {
      const forecast = getRomeDateHourFromTimestamp(item?.forecastAt);
      return forecast?.dateKey === dateKey && forecast.hour === hour;
    });
    return getMostSevereWorklimateRisk(matchingItems);
  }
  api.getWorklimateRiskForDateHour = getWorklimateRiskForDateHour;
  function getWorklimateContextForCommessaAtNoon(commessa, dateKey = getCommessaScheduledDate(commessa?.id)) {
    if (!commessa?.id) return null;
    const risk = getWorklimateRiskForDateHour(worklimateRiskByCommessaId.get(commessa.id) || [], dateKey, 12);
    if (!risk) return null;
    return { commessa, risk, alert: null, riskLevel: normalizeWorklimateLevel(risk.riskLevel || risk.livello || "verde") };
  }
  api.getWorklimateContextForCommessaAtNoon = getWorklimateContextForCommessaAtNoon;
  function normalizeWorklimateRiskDoc(doc) {
    const data = doc.data?.() || doc || {};
    const path = String(data.impiantoPath || "");
    const match = path.match(/commesse\/([^/]+)\/impianti\//);
    const commessaId = data.commessaId || match?.[1] || "";
    return {
      id: doc.id || data.id || "",
      commessaId,
      comune: data.comune || data.zona || "",
      impiantoName: data.impiantoName || data.denominazione || "",
      riskLevel: normalizeWorklimateLevel(data.riskLevel || data.livello || data.level),
      tipoRischio: data.tipoRischio || data.riskType || data.tipoAllerta || data.raw?.riskType || data.raw?.tipoRischio || "caldo",
      forecastAt: data.forecastAt || data.validAt || data.updatedAt || null,
      updatedAt: data.updatedAt || data.forecastAt || null,
      operationalAdvice: Array.isArray(data.operationalAdvice) && data.operationalAdvice.length ? data.operationalAdvice : WORKLIMATE_DEFAULT_ADVICE,
      source: data.source || data.fonte || "Worklimate",
      averageTemperature: getFirstFiniteNumber(
        data.averageTemperature,
        data.temperaturaMedia,
        data.temperatureAvg,
        data.tempMedia,
        data.temperature,
        data.temperatura,
        data.raw?.averageTemperature,
        data.raw?.temperaturaMedia,
        data.raw?.temperature
      )
    };
  }
  api.normalizeWorklimateRiskDoc = normalizeWorklimateRiskDoc;
  function getWorklimateAverageTemperature(context = {}) {
    const risk = context.risk || {};
    const alert = context.alert || {};
    return getFirstFiniteNumber(
      risk.averageTemperature,
      risk.temperaturaMedia,
      risk.temperature,
      alert.averageTemperature,
      alert.temperaturaMedia,
      alert.temperature
    );
  }
  api.getWorklimateAverageTemperature = getWorklimateAverageTemperature;
  function formatWorklimateTemperature(value) {
    return Number.isFinite(Number(value)) ? `${Math.round(Number(value))}°C` : "–°C";
  }
  api.formatWorklimateTemperature = formatWorklimateTemperature;
  function getSquadraWorklimateCodeLineMarkup(commessa, codiceCommessa) {
    const scheduledDateKey = getCommessaScheduledDate(commessa?.id);
    const noonContext = getWorklimateContextForCommessaAtNoon(commessa, scheduledDateKey);
    const context = noonContext || getWorklimateContextForCommessa(commessa) || { commessa, riskLevel: "verde" };
    const temperature = getCommessaAverageImpiantiTemperature(commessa);
    const temperatureLabel = formatSquadraAverageTemperature(temperature);
    const level = normalizeWorklimateLevel(context.riskLevel);
    const temperatureLevel = getSquadraAverageTemperatureLevel(temperature);
    const badgeMarkup = WEATHER_ALERT_PRIORITY[level] > 0
      ? `<button type="button" class="squadra-worklimate-code-badge risk-${escapeHTML(level)}" data-worklimate-commessa="${escapeHTML(commessa.id || "")}" aria-label="Apri sicurezza Worklimate ore 12:00 del ${escapeHTML(scheduledDateKey)}: Codice ${escapeHTML(level)}">Codice ${escapeHTML(level)}</button>`
      : "";
    return `<span class="squadra-commessa-code-line"><span class="squadra-commessa-code-text" aria-label="Codice commessa ${escapeHTML(codiceCommessa || "non disponibile")}">${escapeHTML(codiceCommessa || "-")}</span><button type="button" class="squadra-commessa-temperature risk-${escapeHTML(temperatureLevel)}" data-worklimate-temperature-commessa="${escapeHTML(commessa.id || "")}" aria-label="Apri Worklimate ore 12:00 del ${escapeHTML(scheduledDateKey)}: temperatura media ${escapeHTML(temperatureLabel)}, codice temperatura ${escapeHTML(temperatureLevel)}">🌡️ Media impianti: ${escapeHTML(temperatureLabel)}</button>${badgeMarkup}</span>`;
  }
  api.getSquadraWorklimateCodeLineMarkup = getSquadraWorklimateCodeLineMarkup;
  async function openSquadraWorklimateSafety(commessa, dateKey = getActiveSquadreDateKey(), options = {}) {
    const context = getWorklimateContextForCommessaAtNoon(commessa, dateKey) || getWorklimateContextForCommessa(commessa) || { commessa, riskLevel: "verde" };
    const riskLevel = normalizeWorklimateLevel(context.riskLevel);
    const majorityLocation = getCommessaMajorityImpiantiLocation(commessa, dateKey);
    const comune = options.preferMajorityLocation
      ? (majorityLocation?.comune || context.risk?.comune || context.alert?.comune || getCommessaAlertComuni(commessa)[0] || "Non disponibile")
      : (context.risk?.comune || context.alert?.comune || majorityLocation?.comune || getCommessaAlertComuni(commessa)[0] || "Non disponibile");
    const temperature = options.preferAverageTemperature ? getCommessaAverageImpiantiTemperature(commessa) : getWorklimateAverageTemperature(context);
    if (options.preferMajorityLocation && Number.isFinite(Number(majorityLocation?.lat)) && Number.isFinite(Number(majorityLocation?.lon))) {
      selectedWeatherLocation = { name: majorityLocation.comune, lat: Number(majorityLocation.lat), lon: Number(majorityLocation.lon), source: "commessa" };
      currentWeatherTarget = { ...selectedWeatherLocation };
      await fetchWeather();
    }
    openHomeWorklimateBoard({
      riskLevel,
      url: WORKLIMATE_FORECAST_URL,
      contextData: {
        commessa: commessa.nome || "Commessa",
        codiceCommessa: commessa.codice || "",
        comune,
        selectedDate: dateKey || "",
        averageTemperature: temperature,
        alertLevel: riskLevel,
        source: options.preferMajorityLocation ? `Località prevalente impianti (${majorityLocation?.count || 0})` : (context.risk?.source || context.alert?.fonte || "Worklimate/meteo")
      }
    });
  }
  api.openSquadraWorklimateSafety = openSquadraWorklimateSafety;
  function loadWorklimateRiskCacheBackground() {
    if (!db || !currentUser || worklimateRiskCacheLoading) return Promise.resolve([]);
    worklimateRiskCacheLoading = true;
    return db.collection("worklimateRiskByImpianto").get()
      .then((snapshot) => {
        const grouped = new Map();
        snapshot.forEach((doc) => {
          const risk = normalizeWorklimateRiskDoc(doc);
          if (!risk.commessaId) return;
          if (!grouped.has(risk.commessaId)) grouped.set(risk.commessaId, []);
          grouped.get(risk.commessaId).push(risk);
        });
        worklimateRiskByCommessaId.clear();
        grouped.forEach((items, commessaId) => worklimateRiskByCommessaId.set(commessaId, items));
        worklimateRiskCacheLoaded = true;
        renderSquadre();
        return snapshot.docs;
      })
      .catch((error) => console.error("Cache Worklimate non disponibile:", error))
      .finally(() => { worklimateRiskCacheLoading = false; });
  }
  api.loadWorklimateRiskCacheBackground = loadWorklimateRiskCacheBackground;
  function getWorklimateContextForCommessa(commessa) {
    if (!commessa?.id) return null;
    const risk = getMostSevereWorklimateRisk(worklimateRiskByCommessaId.get(commessa.id) || []);
    const alert = getMostSevereWeatherAlert(getAlertsForCommessa(commessa));
    if (!risk && !alert) return null;
    const riskLevel = normalizeWorklimateLevel(risk?.riskLevel || alert?.livello || "verde");
    return { commessa, risk, alert, riskLevel };
  }
  api.getWorklimateContextForCommessa = getWorklimateContextForCommessa;
  function getHomeWorklimateRiskLevel(temps = []) {
    const maxTemp = Math.max(...temps.slice(0, 12).map((value) => Number(value) || -100), -100);
    if (maxTemp >= 35) return "rosso";
    if (maxTemp >= 32) return "arancione";
    if (maxTemp >= 30) return "giallo";
    return "verde";
  }
  api.getHomeWorklimateRiskLevel = getHomeWorklimateRiskLevel;
  function buildHomeWorklimateButton({ temps = [], target = null } = {}) {
    const riskLevel = getHomeWorklimateRiskLevel(temps);
    const label = getHomeWorklimateButtonLabel(riskLevel, target);
    return `<button type="button" class="weather-risk-chip home-worklimate-btn risk-${riskLevel}" data-home-worklimate-url="${escapeHTML(WORKLIMATE_FORECAST_URL)}" data-home-worklimate-risk="${escapeHTML(riskLevel)}" aria-label="Apri bacheca Worklimate">${escapeHTML(label)}</button>`;
  }
  api.buildHomeWorklimateButton = buildHomeWorklimateButton;
  function bindHomeWorklimateButton() {
    ui.weatherRisks?.querySelector("[data-home-worklimate-url]")?.addEventListener("click", (event) => {
      event.preventDefault();
      event.stopPropagation();
      const button = event.currentTarget;
      const url = button.getAttribute("data-home-worklimate-url") || WORKLIMATE_FORECAST_URL;
      const riskLevel = button.getAttribute("data-home-worklimate-risk") || "verde";
      openHomeWorklimateBoard({ riskLevel, url });
    });
  }
  api.bindHomeWorklimateButton = bindHomeWorklimateButton;
  function openHomeWorklimateBoard({ riskLevel = "verde", url = WORKLIMATE_FORECAST_URL, contextData = null } = {}) {
    const level = normalizeWorklimateLevel(riskLevel);
    const levelLabel = WORKLIMATE_COLOR_LABEL[level] || level;
    const icon = WEATHER_ALERT_ICON[level] || "🟢";
    const target = currentWeatherTarget || getWeatherTargetCoordinates();
    const coordinatesLabel = Number.isFinite(Number(target?.lat)) && Number.isFinite(Number(target?.lon)) ? `${Number(target.lat).toFixed(4)}, ${Number(target.lon).toFixed(4)}` : "";
    const positionLabel = target?.source === "gps" ? `Posizione GPS attuale${coordinatesLabel ? ` • ${coordinatesLabel}` : ""}` : target?.source === "manual" ? `${target.name || "Località scelta"}${coordinatesLabel ? ` • ${coordinatesLabel}` : ""}` : target?.source === "commessa" ? `Comune/zona della commessa${coordinatesLabel ? ` • ${coordinatesLabel}` : ""}` : "Postazione predefinita";
    const updatedLabel = currentHomeWeatherForecast?.updatedAt ? new Date(currentHomeWeatherForecast.updatedAt).toLocaleString("it-IT", { dateStyle: "short", timeStyle: "short" }) : "Dati non disponibili";
    const operationalCards = [
      { icon: "📋", title: "Prima del turno", items: ["Controllare Worklimate", "Verificare disponibilità acqua", "Informare la squadra"] },
      { icon: "☀️", title: "Durante il turno", items: ["Fare pause regolari", "Lavorare preferibilmente all’ombra"] },
      { icon: "🔥", title: "Ore centrali 12:30-16:00", items: ["Ridurre le attività pesanti", "Aumentare la frequenza delle pause"] },
      { icon: "🚨", title: "In caso di livello rosso", items: ["Rimodulare le attività", "Spostare i lavori gravosi al mattino", "Valutare la sospensione dei lavori più pesanti"] }
    ];
    const operationalMarkup = operationalCards.map((card) => `<article class="heat-action-card"><span aria-hidden="true">${escapeHTML(card.icon)}</span><h3>${escapeHTML(card.title)}</h3><ul>${card.items.map((item) => `<li>${escapeHTML(item)}</li>`).join("")}</ul></article>`).join("");
    const contextMarkup = contextData ? `<section class="heat-context-card"><h2>Commessa selezionata</h2><p><b>${escapeHTML(contextData.commessa || "Commessa")}</b>${contextData.codiceCommessa ? ` · ${escapeHTML(contextData.codiceCommessa)}` : ""}</p><p>${escapeHTML(contextData.comune || "Comune non disponibile")} · ${escapeHTML(contextData.selectedDate || "Data non disponibile")} · 🌡️ ${escapeHTML(formatWorklimateTemperature(contextData.averageTemperature))}</p><p>Livello allerta: <b>Codice ${escapeHTML(contextData.alertLevel || level)}</b> · Fonte: ${escapeHTML(contextData.source || "Worklimate/meteo")}</p></section>` : "";
    const overlay = document.createElement("div");
    overlay.className = "worklimate-modal-overlay worklimate-page-overlay";
    overlay.innerHTML = `<div class="worklimate-modal worklimate-board-modal worklimate-board-page heat-dashboard" role="dialog" aria-modal="true" aria-label="Procedura sicurezza rischio calore">
      <header class="worklimate-page-header">
        <button type="button" class="worklimate-page-back" aria-label="Chiudi pagina rischio calore">←</button>
        <div><p>Procedura sicurezza</p><strong>Rischio calore</strong></div>
        <button type="button" class="worklimate-modal-close" aria-label="Chiudi">×</button>
      </header>
      <main data-worklimate-main class="heat-dashboard-main">
        <section class="heat-status-card risk-${escapeHTML(level)}">
          <div class="heat-status-icon" aria-hidden="true">${escapeHTML(icon)}</div>
          <div><h1>Rischio calore</h1><p>${escapeHTML(positionLabel)}</p><small>Ultimo aggiornamento: ${escapeHTML(updatedLabel)}</small></div>
          <button type="button" class="heat-location-button" data-heat-location-picker>📍 Cambia località</button>
          <strong>${escapeHTML(levelLabel)}</strong>
        </section>
        ${contextMarkup}
        <section class="heat-section"><div class="heat-section-title"><span aria-hidden="true">☀️</span><h2>Previsioni prossimi 5 giorni</h2></div>${renderHeatForecastCards()}</section>
        <section class="heat-section"><div class="heat-section-title"><span aria-hidden="true">⚠️</span><h2>Indicazioni operative</h2></div><div class="heat-actions-grid">${operationalMarkup}</div></section>
        <section class="heat-info-card hydration"><h2>💧 Idratazione</h2><p>Bere almeno 250 ml di acqua ogni 20 minuti.</p></section>
        <section class="heat-law-compact"><div><h2>📌 Ordinanza Regionale n.72 del 03/06/2026</h2><p>Misure di prevenzione per attività lavorative in condizioni di esposizione al calore.</p></div><button type="button" class="btn btn-primary worklimate-visit-btn" data-worklimate-visit="${escapeHTML(url)}">Apri documento</button></section>
      </main>
    </div>`;
    const close = () => {
      document.body.classList.remove("worklimate-page-open");
      overlay.remove();
    };
    document.body.classList.add("worklimate-page-open");
    overlay.addEventListener("click", (event) => { if (event.target === overlay) close(); });
    overlay.querySelector(".worklimate-modal-close")?.addEventListener("click", close);
    overlay.querySelector(".worklimate-page-back")?.addEventListener("click", close);
    overlay.querySelector("[data-heat-location-picker]")?.addEventListener("click", () => openHeatLocationPicker(overlay));
    overlay.querySelectorAll("[data-heat-day]").forEach((card) => {
      card.addEventListener("click", () => openHeatHourlyForecast(card.getAttribute("data-heat-day") || ""));
    });
    overlay.querySelector("[data-worklimate-visit]")?.addEventListener("click", (event) => {
      const visitUrl = event.currentTarget.getAttribute("data-worklimate-visit") || WORKLIMATE_FORECAST_URL;
      window.open(visitUrl, "_blank", "noopener,noreferrer");
    });
    document.body.appendChild(overlay);
    overlay.querySelector(".worklimate-page-back")?.focus();
  }
  api.openHomeWorklimateBoard = openHomeWorklimateBoard;
  Object.assign(global, api);
  global.VargaWorklimateModule = Object.freeze({ ...api });
})(window);
