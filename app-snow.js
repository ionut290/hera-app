"use strict";
(function installVargaSnowModule(global) {
  if (global.VargaSnowModule) return;
  const api = {};
  function stopNormalDataSubscriptionsForSnowMode() {
    stopCommesseSubscription();
    stopSquadreSubscription();
    stopPersonaleSubscription();
    stopMezziSubscription();
  }
  api.stopNormalDataSubscriptionsForSnowMode = stopNormalDataSubscriptionsForSnowMode;
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
  api.clearSnowModeRuntimeData = clearSnowModeRuntimeData;
  function stopSnowServiceCollections() {
    snowServiceUnsubscribers.forEach((unsubscribe) => unsubscribe && unsubscribe());
    snowServiceUnsubscribers = [];
  }
  api.stopSnowServiceCollections = stopSnowServiceCollections;
  function stopSnowModeData() {
    stopSnowServiceCollections();
    stopCommesseSubscription();
    stopSquadreSubscription();
    stopPersonaleSubscription();
    stopMezziSubscription();
    clearSnowModeRuntimeData();
  }
  api.stopSnowModeData = stopSnowModeData;
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
  api.loadSnowModeData = loadSnowModeData;
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
    setFullscreenFeedback("Tracciato via neve salvato: la linea diventerà verde quando l’operatore passa sulla strada.");
    renderMap();
  }
  api.saveDrawnSnowRoadPath = saveDrawnSnowRoadPath;
  function parseSnowRoadLines(value) {
    return String(value || "")
      .split(/\r?\n/)
      .map((line) => line.trim())
      .filter(Boolean)
      .map((line, index) => ({
        denominazione: line,
        indirizzo: line,
        descrizioneVia: line,
        tipologiaImpianto: "Via neve",
        codicePrezzo: "NEVE-STRADA",
        snowRoad: true,
        roadStatus: "todo",
        routePath: [],
        done: false,
        doneAt: null,
        doneBy: "",
        sortOrder: index + 1,
        createdAt: firebase.firestore.FieldValue.serverTimestamp(),
        createdBy: currentUser?.email || ""
      }));
  }
  api.parseSnowRoadLines = parseSnowRoadLines;
  async function addSnowRoadsToSelectedCommessa(event) {
    event?.preventDefault?.();
    if (!canManageData()) return alert("Solo un admin può aggiungere vie neve.");
    if (!isSnowServiceContext()) return alert("Apri Servizio Neve per aggiungere vie neve separate.");
    const commessaId = String(ui.commessaTargetSelect?.value || selectedCommessaId || "").trim();
    if (!commessaId) return alert("Seleziona una commessa neve.");
    const textarea = document.getElementById("snow-roads-list");
    const rows = parseSnowRoadLines(textarea?.value || "");
    if (!rows.length) return alert("Inserisci almeno una via, una per riga.");
    const batch = db.batch();
    const ref = db.collection("neve_commesse").doc(commessaId).collection("impianti");
    rows.forEach((row) => batch.set(ref.doc(), row));
    await batch.commit();
    if (textarea) textarea.value = "";
    const feedback = document.getElementById("snow-roads-feedback");
    if (feedback) feedback.textContent = `${rows.length} vie neve aggiunte come cantieri da pulire.`;
  }
  api.addSnowRoadsToSelectedCommessa = addSnowRoadsToSelectedCommessa;
  function getSnowSquadreDateKey() {
    if (snowManualSquadreFilterDateKey) return snowManualSquadreFilterDateKey;
    if (snowSharedSquadreDateKey) return snowSharedSquadreDateKey;
    if (!automaticSquadreDateKey) automaticSquadreDateKey = getAutomaticSquadreDateKey();
    return automaticSquadreDateKey;
  }
  api.getSnowSquadreDateKey = getSnowSquadreDateKey;
  function onSnowSquadreFilterDateChange() {
    setSquadreDateOverride(ui.snowSquadreFilterDate?.value || "", { snow: true });
  }
  api.onSnowSquadreFilterDateChange = onSnowSquadreFilterDateChange;
  function getSnowRoadPath(impianto) {
    const raw = Array.isArray(impianto?.routePath) ? impianto.routePath : [];
    return raw.map((point) => Array.isArray(point)
      ? [Number(point[0]), Number(point[1])]
      : [Number(point?.lat ?? point?.latitude), Number(point?.lng ?? point?.lon ?? point?.longitude)]
    ).filter(([lat, lng]) => Number.isFinite(lat) && Number.isFinite(lng));
  }
  api.getSnowRoadPath = getSnowRoadPath;
  function addSnowRoadPolylineToLayer(impianto, targetLayer, targetMap) {
    const path = getSnowRoadPath(impianto);
    if (!impianto?.snowRoad || path.length < 2) return null;
    const color = impianto.done ? "#16a34a" : "#2563eb";
    const line = L.polyline(path, { color, weight: 6, opacity: 0.9, lineCap: "round", lineJoin: "round" });
    if (targetMap !== fullscreenMap) line.bindPopup(buildImpiantoMapPopup(impianto, impianto.tipoManutenzione || "Via neve"));
    line.on("click", () => selectImpiantoForMapDetail(impianto));
    line.addTo(targetLayer);
    return line;
  }
  api.addSnowRoadPolylineToLayer = addSnowRoadPolylineToLayer;
  function distanceMetersToSnowRoad(impianto, position) {
    const path = getSnowRoadPath(impianto);
    if (!position || path.length < 2) return Number.POSITIVE_INFINITY;
    return path.reduce((best, point) => Math.min(best, haversine(position.lat, position.lng, point[0], point[1]) * 1000), Number.POSITIVE_INFINITY);
  }
  api.distanceMetersToSnowRoad = distanceMetersToSnowRoad;
  async function autoCompletePassedSnowRoads() {
    if (!isSnowServiceContext() || !selectedCommessaId || !currentUserPos) return;
    const passed = currentImpianti.filter((impianto) => impianto?.snowRoad && !impianto.done && distanceMetersToSnowRoad(impianto, currentUserPos) <= 25);
    if (!passed.length) return;
    await setImpiantoDone(selectedCommessaId, passed.map((impianto) => impianto.id).filter(Boolean), true, { doneBy: currentUser?.displayName || currentUser?.email || "Operatore neve" });
  }
  api.autoCompletePassedSnowRoads = autoCompletePassedSnowRoads;
  function isSnowServiceContext() {
    return window.location.hash === "#servizio-neve" || document.body.classList.contains("snow-management-context");
  }
  api.isSnowServiceContext = isSnowServiceContext;
  function openSnowServicePage() {
    if (!canManageData()) {
      alert("Solo l'admin può accedere al Servizio Neve.");
      return;
    }
    document.body.classList.add("snow-management-context");
    window.location.hash = "servizio-neve";
    stopNormalDataSubscriptionsForSnowMode();
    loadSnowModeData();
    applyRoute();
  }
  api.openSnowServicePage = openSnowServicePage;
  function closeSnowServicePage() {
    stopSnowModeData();
    document.body.classList.remove("snow-management-context");
    configureSnowSideMenu(false);
    window.location.hash = "";
    reloadNormalModeData();
    applyRoute();
  }
  api.closeSnowServicePage = closeSnowServicePage;
  function isSnowServiceRoute() {
    return window.location.hash === "#servizio-neve";
  }
  api.isSnowServiceRoute = isSnowServiceRoute;
  function renderSnowServiceCommesse() {
    const list = document.getElementById("snow-squadre-lista");
    if (!list) return;
    list.innerHTML = "";
    if (areStartupCoreCollectionsLoading()) {
      list.innerHTML = `<p class='muted'>${escapeHTML(startupCoreCollectionsLoadState.message || "Caricamento dati squadra neve...")}</p>`;
      return;
    }
    if (squadreLoadState.status === "loading") {
      list.innerHTML = `<p class='muted'>${escapeHTML(squadreLoadState.message || "Caricamento squadre neve...")}</p>`;
      return;
    }
    if (squadreLoadState.status === "auth-required") {
      list.innerHTML = `<p class='muted'>${escapeHTML(squadreLoadState.message || "Fai login per caricare le squadre neve.")}</p>`;
      return;
    }
    if (squadreLoadState.status === "error") {
      list.innerHTML = `<p class='muted'>${escapeHTML(squadreLoadState.message || "Errore caricamento dati")}</p><button id='snow-squadre-retry-btn' class='btn btn-primary' type='button'>Riprova</button>`;
      list.querySelector("#snow-squadre-retry-btn")?.addEventListener("click", () => subscribeSquadre());
      return;
    }
    const selectedDateKey = getActiveSquadreDateKey();
    if (!selectedDateKey) return;
    const storicoDelGiorno = squadreHistoryByDate.get(selectedDateKey) || new Map();
    const commesseNeve = Array.from(commesseById.values()).filter((commessa) => {
      const squad = storicoDelGiorno.get(commessa.id) || {};
      const rows = Array.isArray(squad.squadre) ? squad.squadre : getLegacySquadreRows(squad);
      return rows.some(isSquadraRowFilled);
    });
    if (!commesseNeve.length) {
      list.innerHTML = "<p class='muted'>Nessuna squadra neve creata per questo giorno</p>";
      return;
    }
    commesseNeve.forEach((commessa) => {
      const item = document.createElement("article");
      item.className = "squadra-item";
      const squad = storicoDelGiorno.get(commessa.id) || {};
      const squadRows = Array.isArray(squad.squadre) ? squad.squadre : getLegacySquadreRows(squad);
      const riferimento = squad.riferimentoData
        ? new Date(`${squad.riferimentoData}T00:00:00`).toLocaleDateString("it-IT")
        : formatDateKeyForDisplay(selectedDateKey);
      const rowsHtml = squadRows.map((row, idx) => {
        const orarioLabel = formatSquadraOrario(row);
        const details = [
          row.caposquadra ? `<br><b>🧑‍✈️ Caposquadra:</b> ${escapeHTML(row.caposquadra)}` : "",
          orarioLabel ? `<br><b>🕒</b> ${escapeHTML(orarioLabel)}` : "",
          row.impianti ? `<br><b>📍 Impianti:</b> ${escapeHTML(row.impianti)}` : "",
          row.note ? `<br><b>📝 Note:</b> ${escapeHTML(row.note)}` : ""
        ].join("");
        return `<div class="squadra-saved-row" data-squadra-index="${idx}"><p><button type="button" class="squadra-edit-link" data-commessa-id="${escapeHTML(commessa.id)}" data-date-key="${escapeHTML(selectedDateKey)}" data-squadra-index="${idx}" aria-label="Modifica Squadra neve ${idx + 1} di ${escapeHTML(commessa.nome || "commessa")}">👥 Squadra ${idx + 1}:</button> ${escapeHTML(row.personale || "-")}${details}<br><b>🚚 Mezzi ${idx + 1}:</b> ${renderMezziButtonsMarkup(row.mezzi)}</p></div>`;
      }).join("");
      const warningIssues = buildSquadraWarningDetails(commessa, squadRows);
      const warningMarkup = warningIssues.length
        ? `<div class="squadra-warning-wrap"><button type="button" class="squadra-warning-toggle" aria-expanded="false" aria-label="Mostra controllo squadra neve">⚠️</button><div class="squadra-warning-details hidden"><p><b>⚠️ Controllo squadra</b></p><ul>${warningIssues.map((issue) => `<li>${escapeHTML(issue.replace(/^⚠️\s*/, ""))}</li>`).join("")}</ul></div></div>`
        : "";
      const codiceCommessa = String(commessa.codice || "").trim();
      item.innerHTML = `
        <div class="squadra-item-head squadra-commessa-link" role="button" tabindex="0" aria-label="Apri dettaglio commessa neve ${escapeHTML(commessa.nome || "Commessa senza nome")}">
          <div class="squadra-commessa-title-wrap">
            <strong>📁 ${escapeHTML(commessa.nome || "Commessa neve")}</strong>
            ${getSquadraWorklimateCodeLineMarkup(commessa, codiceCommessa)}
            <div class="snow-squadra-meta"><span class="pill">❄️ Servizio neve</span></div>
          </div>
          ${warningMarkup}
        </div>
        <p><b>📅 Giorno:</b> ${escapeHTML(riferimento)}</p>
        ${rowsHtml}
      `;
      const head = item.querySelector(".squadra-item-head");
      appendSquadreHeaderRiskActions(head, commessa, selectedDateKey);
      appendAddHoursButtonIfAllowed(head, commessa, selectedDateKey);
      head?.addEventListener("click", (event) => {
        if (event.target.closest("button, a, input, select, textarea")) return;
        openCommessaFromSquadre(commessa);
      });
      head?.addEventListener("keydown", (event) => {
        if (event.key !== "Enter" && event.key !== " ") return;
        if (event.target.closest("button, a, input, select, textarea")) return;
        event.preventDefault();
        openCommessaFromSquadre(commessa);
      });
      item.querySelectorAll(".squadra-edit-link").forEach((btn) => {
        btn.addEventListener("click", (event) => {
          event.preventDefault();
          event.stopPropagation();
          openSquadraCompositionEditor(btn.dataset.commessaId || commessa.id, btn.dataset.dateKey || selectedDateKey, Number(btn.dataset.squadraIndex) || 0);
        });
      });
      item.querySelectorAll(".mezzo-chip-btn").forEach((btn) => {
        btn.addEventListener("click", () => openFuelPage(btn.dataset.mezzo || ""));
      });
      item.querySelector("[data-worklimate-commessa]")?.addEventListener("click", (event) => {
        event.preventDefault();
        event.stopPropagation();
        openSquadraWorklimateSafety(commessa, selectedDateKey);
      });
      item.querySelector("[data-worklimate-temperature-commessa]")?.addEventListener("click", (event) => {
        event.preventDefault();
        event.stopPropagation();
        openSquadraWorklimateSafety(commessa, selectedDateKey, { preferMajorityLocation: true, preferAverageTemperature: true });
      });
      const warningToggle = item.querySelector(".squadra-warning-toggle");
      warningToggle?.addEventListener("click", (event) => {
        event.stopPropagation();
        const details = item.querySelector(".squadra-warning-details");
        const isHidden = details?.classList.contains("hidden");
        details?.classList.toggle("hidden", !isHidden);
        warningToggle.setAttribute("aria-expanded", isHidden ? "true" : "false");
      });
      list.appendChild(item);
    });
  }
  api.renderSnowServiceCommesse = renderSnowServiceCommesse;
  function syncSnowWeatherPanel() {
    const summary = document.getElementById("snow-weather-summary");
    const risks = document.getElementById("snow-weather-risks");
    if (summary && ui.weatherSummary) summary.textContent = ui.weatherSummary.textContent || "Caricamento meteo...";
    if (risks && ui.weatherRisks) risks.innerHTML = ui.weatherRisks.innerHTML;
  }
  api.syncSnowWeatherPanel = syncSnowWeatherPanel;
  function configureSnowSideMenu(isSnow) {
    const allowed = new Set(["open-panel-commesse", "open-panel-squadre", "open-panel-personale", "open-panel-mezzi", "open-panel-utenti", "open-hours-btn"]);
    document.body.classList.toggle("snow-management-context", Boolean(isSnow));
    document.querySelectorAll("#side-menu .menu-title-btn").forEach((button) => {
      button.classList.toggle("hidden", Boolean(isSnow) && !allowed.has(button.id));
    });
  }
  api.configureSnowSideMenu = configureSnowSideMenu;
  function renderSnowService() {
    syncSnowWeatherPanel();
    syncSquadreDateInputs();
    configureSnowSideMenu(true);
    renderSnowServiceCommesse();
  }
  api.renderSnowService = renderSnowService;
  function subscribeSnowServiceCollections() {
    stopSnowServiceCollections();
    if (!db || !currentUser || !isSnowServiceContext()) return Promise.resolve(false);
    Object.entries({ clients: SNOW_SERVICE_COLLECTIONS.clients, routes: SNOW_SERVICE_COLLECTIONS.routes, vehicles: SNOW_SERVICE_COLLECTIONS.vehicles, operators: SNOW_SERVICE_COLLECTIONS.operators, reports: SNOW_SERVICE_COLLECTIONS.reports }).forEach(([key, collectionName]) => {
      snowServiceUnsubscribers.push(db.collection(collectionName).orderBy("createdAt", "desc").onSnapshot((snapshot) => {
        snowServiceState[key] = snapshot.docs.map((doc) => ({ id: doc.id, ...doc.data() }));
        renderSnowService();
      }, (error) => console.warn(`Errore caricamento ${collectionName}:`, error)));
    });
    return Promise.resolve(true);
  }
  api.subscribeSnowServiceCollections = subscribeSnowServiceCollections;
  async function addSnowServiceItem(type) {
    if (!canManageData()) return alert("Solo admin può modificare il Servizio Neve.");
    const config = {
      clients: { collection: SNOW_SERVICE_COLLECTIONS.clients, prompt: "Nome Comune / Cliente neve", field: "nome" },
      routes: { collection: SNOW_SERVICE_COLLECTIONS.routes, prompt: "Nome percorso neve", field: "nome" },
      vehicles: { collection: SNOW_SERVICE_COLLECTIONS.vehicles, prompt: "Nome mezzo neve", field: "nome" },
      operators: { collection: SNOW_SERVICE_COLLECTIONS.operators, prompt: "Nome operatore neve", field: "nome" },
      reports: { collection: SNOW_SERVICE_COLLECTIONS.reports, prompt: "Titolo segnalazione neve", field: "titolo" }
    }[type];
    if (!config) return;
    const value = String(window.prompt(config.prompt, "") || "").trim();
    if (!value) return;
    const note = String(window.prompt("Note / dettagli (opzionale)", "") || "").trim();
    await db.collection(config.collection).add({ [config.field]: value, note, createdAt: firebase.firestore.FieldValue.serverTimestamp(), createdBy: currentUser?.email || "" });
  }
  api.addSnowServiceItem = addSnowServiceItem;
  async function deleteSnowServiceItem(type, id) {
    if (!canManageData()) return alert("Solo admin può eliminare elementi del Servizio Neve.");
    const collection = SNOW_SERVICE_COLLECTIONS[type];
    if (!collection || !id || !window.confirm("Eliminare questo elemento neve?")) return;
    await db.collection(collection).doc(id).delete();
  }
  api.deleteSnowServiceItem = deleteSnowServiceItem;
  function handleSnowServiceMenuAction(action) {
    if (action === "settings") return alert("Impostazioni servizio neve: modulo dedicato e separato dalla Manutenzione Verde.");
    if (action === "add-client") return addSnowServiceItem("clients");
    if (action === "add-route") return addSnowServiceItem("routes");
    if (action === "manage-vehicles") return addSnowServiceItem("vehicles");
    if (action === "manage-operators") return addSnowServiceItem("operators");
  }
  api.handleSnowServiceMenuAction = handleSnowServiceMenuAction;
  Object.assign(global, api);
  global.VargaSnowModule = Object.freeze({ ...api });
})(window);
