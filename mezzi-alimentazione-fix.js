(function () {
  "use strict";

  function normalizeAlimentazione(value) {
    const original = String(value || "").trim();
    const normalized = original.toLowerCase();
    const hasBenzina = /\bbenzina\b/.test(normalized);
    const hasMetano = /\bmetano\b/.test(normalized);
    const hasGpl = /\bgpl\b/.test(normalized);
    if (!hasBenzina || (!hasMetano && !hasGpl)) return original;
    return [hasMetano ? "METANO" : "", hasGpl ? "GPL" : ""].filter(Boolean).join(" + ");
  }

  const originalNormalizeMezzoDocument = window.normalizeMezzoDocument;
  if (typeof originalNormalizeMezzoDocument === "function") {
    window.normalizeMezzoDocument = function (doc) {
      const mezzo = originalNormalizeMezzoDocument(doc);
      return { ...mezzo, alimentazione: normalizeAlimentazione(mezzo.alimentazione) };
    };
  }

  document.getElementById("mezzi-form")?.addEventListener("submit", function () {
    const input = document.getElementById("mezzo-alimentazione");
    if (input) input.value = normalizeAlimentazione(input.value);
  }, true);

  const originalSubscribeMezzi = window.subscribeMezzi;
  const originalSubscribeSquadre = window.subscribeSquadre;
  const originalOpenManagementPanel = window.openManagementPanel;
  const originalCloseManagementPanel = window.closeManagementPanel;
  let fullMezziMode = false;
  let assignedLoadSequence = 0;

  const normalizeVehicleKey = (value) => String(value || "").trim().toLocaleUpperCase("it-IT");

  function parseAssignedVehicles(value) {
    if (Array.isArray(value)) return value.flatMap(parseAssignedVehicles);
    if (value && typeof value === "object") {
      return parseAssignedVehicles(value.nId || value.codice || value.targa || value.nome || value.label || value.id || "");
    }
    return String(value || "").split(/[\n,;|]+/).map((item) => item.trim()).filter(Boolean);
  }

  function collectAssignedVehicleCodes() {
    const dateKey = typeof getActiveSquadreDateKey === "function"
      ? getActiveSquadreDateKey()
      : (typeof getTodayDateKey === "function" ? getTodayDateKey() : "");
    const history = typeof squadreHistoryByDate !== "undefined" && squadreHistoryByDate instanceof Map
      ? squadreHistoryByDate.get(dateKey)
      : null;
    const codes = new Map();
    const addRow = (row) => {
      if (!row || typeof row !== "object") return;
      parseAssignedVehicles(row.mezzi ?? row.mezzo ?? row.veicoli ?? row.mezziAssegnati ?? "").forEach((value) => {
        const key = normalizeVehicleKey(value);
        if (key && !codes.has(key)) codes.set(key, value);
      });
    };
    if (history instanceof Map) {
      history.forEach((entry) => (Array.isArray(entry?.squadre) ? entry.squadre : []).forEach(addRow));
    }
    return [...codes.values()];
  }

  function applyVehicleRecords(records) {
    mezziRecords = Array.isArray(records) ? records : [];
    mezziLoadState = { status: "loaded", message: "" };
    if (typeof renderTodaySummary === "function") renderTodaySummary();
    if (typeof updateSquadraHintFromSources === "function") updateSquadraHintFromSources();
    if (typeof updateSuggestionLists === "function") updateSuggestionLists();
  }

  async function getSnapshot(query, label) {
    return typeof runFirestoreGetWithRetry === "function"
      ? runFirestoreGetWithRetry(query, { label, timeoutMs: 9000, retries: 2 })
      : query.get();
  }

  async function queryByField(collection, field, values, sequence) {
    const docs = [];
    for (let start = 0; start < values.length; start += 10) {
      if (sequence !== assignedLoadSequence || fullMezziMode) return docs;
      const snapshot = await getSnapshot(
        collection.where(field, "in", values.slice(start, start + 10)),
        `LOAD MEZZI ASSEGNATI ${field}`
      );
      snapshot.docs.forEach((doc) => docs.push(doc));
    }
    return docs;
  }

  async function loadAssignedVehicles() {
    if (fullMezziMode || typeof currentUser === "undefined" || !currentUser) return false;
    const sequence = ++assignedLoadSequence;
    const values = [...new Set(collectAssignedVehicleCodes().map((value) => String(value).trim()).filter(Boolean))];
    if (!values.length) {
      applyVehicleRecords([]);
      return true;
    }

    mezziLoadState = { status: "loading", message: "Caricamento mezzi assegnati..." };
    try {
      const collection = db.collection(getMezziCollectionName());
      const byNId = await queryByField(collection, "nId", values, sequence);
      if (sequence !== assignedLoadSequence || fullMezziMode) return false;

      const matched = new Set(byNId.map((doc) => {
        const data = doc.data() || {};
        return normalizeVehicleKey(data.nId || data.numero || "");
      }));
      const unresolved = values.filter((value) => !matched.has(normalizeVehicleKey(value)));
      const byTarga = unresolved.length ? await queryByField(collection, "targa", unresolved, sequence) : [];
      if (sequence !== assignedLoadSequence || fullMezziMode) return false;

      const uniqueDocs = new Map([...byNId, ...byTarga].map((doc) => [doc.id, doc]));
      const normalizer = window.normalizeMezzoDocument || normalizeMezzoDocument;
      const records = [...uniqueDocs.values()].map((doc) => normalizer(doc));
      applyVehicleRecords(records);
      console.log("Mezzi assegnati caricati", { richiesti: values.length, trovati: records.length });
      return true;
    } catch (error) {
      console.warn("Caricamento mirato mezzi non riuscito", error);
      if (sequence === assignedLoadSequence && !fullMezziMode) applyVehicleRecords([]);
      return false;
    }
  }

  if (typeof originalSubscribeMezzi === "function") {
    window.subscribeMezzi = function () {
      if (fullMezziMode) return originalSubscribeMezzi.apply(this, arguments);
      applyVehicleRecords([]);
      return Promise.resolve(true);
    };
  }

  if (typeof originalSubscribeSquadre === "function") {
    window.subscribeSquadre = function () {
      const result = originalSubscribeSquadre.apply(this, arguments);
      return Promise.resolve(result).then((value) => {
        if (!fullMezziMode) void loadAssignedVehicles();
        return value;
      });
    };
  }

  function enableFullMezziMode() {
    if (fullMezziMode) return;
    fullMezziMode = true;
    assignedLoadSequence += 1;
    if (typeof originalSubscribeMezzi === "function") void originalSubscribeMezzi();
  }

  function restoreAssignedMezziMode() {
    if (!fullMezziMode) return;
    fullMezziMode = false;
    assignedLoadSequence += 1;
    if (typeof unsubscribeMezzi === "function") {
      unsubscribeMezzi();
      unsubscribeMezzi = null;
    }
    void loadAssignedVehicles();
  }

  if (typeof originalOpenManagementPanel === "function") {
    window.openManagementPanel = function (panelName) {
      if (["mezzi", "squadre", "programmazione"].includes(String(panelName || ""))) enableFullMezziMode();
      return originalOpenManagementPanel.apply(this, arguments);
    };
  }

  if (typeof originalCloseManagementPanel === "function") {
    window.closeManagementPanel = function () {
      const result = originalCloseManagementPanel.apply(this, arguments);
      restoreAssignedMezziMode();
      return result;
    };
  }

  function loadOptionalScript(selector, src, datasetKey) {
    if (document.querySelector(selector)) return;
    const script = document.createElement("script");
    script.src = src;
    script.defer = true;
    script.dataset[datasetKey] = "1";
    document.head.appendChild(script);
  }

  loadOptionalScript('script[data-squadre-mezzi-pictograms]', "./squadre-mezzi-pictograms.js?v=20260727a", "squadreMezziPictograms");
  loadOptionalScript('script[data-today-live-hours-vehicles]', "./today-live-hours-vehicles.js?v=20260730b", "todayLiveHoursVehicles");
  loadOptionalScript('script[data-squad-operator-profile]', "./squad-operator-profile.js?v=20260731a", "squadOperatorProfile");
  loadOptionalScript('script[data-calendar-personal-hours-loader]', "./calendar-personal-hours-loader.js?v=20260802b", "calendarPersonalHoursLoader");
  loadOptionalScript('script[data-rubrica-personale-restore]', "./rubrica-personale-restore.js?v=20260802-clear2", "rubricaPersonaleRestore");
})();
