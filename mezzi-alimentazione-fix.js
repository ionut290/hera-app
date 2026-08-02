(function () {
  "use strict";

  function normalizeAlimentazione(value) {
    const original = String(value || "").trim();
    const normalized = original.toLowerCase();
    const hasBenzina = /\bbenzina\b/.test(normalized);
    const hasMetano = /\bmetano\b/.test(normalized);
    const hasGpl = /\bgpl\b/.test(normalized);

    if (!hasBenzina || (!hasMetano && !hasGpl)) return original;

    const alimentazioni = [];
    if (hasMetano) alimentazioni.push("METANO");
    if (hasGpl) alimentazioni.push("GPL");
    return alimentazioni.join(" + ");
  }

  const originalNormalizeMezzoDocument = window.normalizeMezzoDocument;
  if (typeof originalNormalizeMezzoDocument === "function") {
    window.normalizeMezzoDocument = function (doc) {
      const mezzo = originalNormalizeMezzoDocument(doc);
      return {
        ...mezzo,
        alimentazione: normalizeAlimentazione(mezzo.alimentazione)
      };
    };
  }

  document.getElementById("mezzi-form")?.addEventListener("submit", function () {
    const input = document.getElementById("mezzo-alimentazione");
    if (input) input.value = normalizeAlimentazione(input.value);
  }, true);

  // La bonifica storica è conclusa: nessuna scansione automatica della collezione mezzi.
  // La normalizzazione resta attiva durante lettura e salvataggio dei singoli mezzi.

  /*
   * Ottimizzazione Firestore mezzi:
   * - all'avvio non scarica più l'intera anagrafica;
   * - dopo il caricamento delle squadre legge solo i mezzi assegnati alla data attiva;
   * - l'anagrafica completa e il listener live vengono attivati soltanto nelle sezioni
   *   amministrative che devono scegliere o gestire qualunque mezzo.
   */
  const originalSubscribeMezzi = window.subscribeMezzi;
  const originalSubscribeSquadre = window.subscribeSquadre;
  const originalOpenManagementPanel = window.openManagementPanel;
  const originalCloseManagementPanel = window.closeManagementPanel;
  let fullMezziMode = false;
  let assignedLoadSequence = 0;

  const normalizeVehicleKey = (value) => String(value || "")
    .trim()
    .toLocaleUpperCase("it-IT");

  function parseAssignedVehicles(value) {
    if (Array.isArray(value)) return value.flatMap(parseAssignedVehicles);
    if (value && typeof value === "object") {
      return parseAssignedVehicles(value.nId || value.codice || value.targa || value.nome || value.label || value.id || "");
    }
    return String(value || "")
      .split(/[\n,;|]+/)
      .map((item) => item.trim())
      .filter(Boolean);
  }

  function collectAssignedVehicleCodes() {
    const dateKey = typeof getActiveSquadreDateKey === "function"
      ? getActiveSquadreDateKey()
      : (typeof getTodayDateKey === "function" ? getTodayDateKey() : "");
    const history = typeof squadreHistoryByDate !== "undefined" && squadreHistoryByDate instanceof Map
      ? squadreHistoryByDate.get(dateKey)
      : null;
    const codes = new Map();

    const addFromRow = (row) => {
      if (!row || typeof row !== "object") return;
      const raw = row.mezzi ?? row.mezzo ?? row.veicoli ?? row.mezziAssegnati ?? "";
      parseAssignedVehicles(raw).forEach((value) => {
        const key = normalizeVehicleKey(value);
        if (key && !codes.has(key)) codes.set(key, value);
      });
    };

    if (history instanceof Map) {
      history.forEach((entry) => {
        const rows = Array.isArray(entry?.squadre) ? entry.squadre : [];
        rows.forEach(addFromRow);
      });
    }
    return [...codes.values()];
  }

  function applyAssignedVehicleRecords(records) {
    mezziRecords = Array.isArray(records) ? records : [];
    mezziLoadState = { status: "loaded", message: "" };
    if (typeof renderTodaySummary === "function") renderTodaySummary();
    if (typeof updateSquadraHintFromSources === "function") updateSquadraHintFromSources();
    if (typeof updateSuggestionLists === "function") updateSuggestionLists();
  }

  async function getQuerySnapshot(query, label) {
    if (typeof runFirestoreGetWithRetry === "function") {
      return runFirestoreGetWithRetry(query, { label, timeoutMs: 9000, retries: 2 });
    }
    return query.get();
  }

  async function queryAssignedVehiclesByField(collection, field, values, sequence) {
    const found = [];
    for (let start = 0; start < values.length; start += 10) {
      if (sequence !== assignedLoadSequence || fullMezziMode) return found;
      const chunk = values.slice(start, start + 10);
      const snapshot = await getQuerySnapshot(
        collection.where(field, "in", chunk),
        `LOAD MEZZI ASSEGNATI ${field}`
      );
      snapshot.docs.forEach((doc) => found.push(doc));
    }
    return found;
  }

  async function loadAssignedVehicles() {
    if (fullMezziMode || !window.currentUser) return false;
    const sequence = ++assignedLoadSequence;
    const assignedValues = collectAssignedVehicleCodes();
    if (!assignedValues.length) {
      applyAssignedVehicleRecords([]);
      return true;
    }

    mezziLoadState = { status: "loading", message: "Caricamento mezzi assegnati..." };
    try {
      const collection = db.collection(getMezziCollectionName());
      const exactValues = [...new Set(assignedValues.map((value) => String(value).trim()).filter(Boolean))];
      const byNId = await queryAssignedVehiclesByField(collection, "nId", exactValues, sequence);
      if (sequence !== assignedLoadSequence || fullMezziMode) return false;

      const matchedKeys = new Set();
      byNId.forEach((doc) => {
        const data = doc.data() || {};
        matchedKeys.add(normalizeVehicleKey(data.nId || data.numero || ""));
      });
      const unresolved = exactValues.filter((value) => !matchedKeys.has(normalizeVehicleKey(value)));
      const byTarga = unresolved.length
        ? await queryAssignedVehiclesByField(collection, "targa", unresolved, sequence)
        : [];
      if (sequence !== assignedLoadSequence || fullMezziMode) return false;

      const uniqueDocs = new Map([...byNId, ...byTarga].map((doc) => [doc.id, doc]));
      const records = [...uniqueDocs.values()].map(window.normalizeMezzoDocument || normalizeMezzoDocument);
      applyAssignedVehicleRecords(records);
      console.log("Mezzi assegnati caricati", { richiesti: exactValues.length, trovati: records.length });
      return true;
    } catch (error) {
      console.warn("Caricamento mirato mezzi non riuscito", error);
      if (sequence === assignedLoadSequence && !fullMezziMode) {
        // Nessun fallback all'intera collezione: il riepilogo Oggi usa già i codici presenti nelle squadre.
        applyAssignedVehicleRecords([]);
      }
      return false;
    }
  }

  if (typeof originalSubscribeMezzi === "function") {
    window.subscribeMezzi = function subscribeOnlyNeededVehicles() {
      if (fullMezziMode) return originalSubscribeMezzi.apply(this, arguments);
      // Durante l'avvio le squadre vengono caricate in parallelo. Evita qui la lettura di tutti i mezzi;
      // il caricamento mirato parte subito dopo il primo snapshot delle squadre.
      applyAssignedVehicleRecords([]);
      return Promise.resolve(true);
    };
  }

  if (typeof originalSubscribeSquadre === "function") {
    window.subscribeSquadre = function subscribeSquadreAndAssignedVehicles() {
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
    window.openManagementPanel = function openManagementPanelWithLazyMezzi(panelName) {
      if (["mezzi", "squadre", "programmazione"].includes(String(panelName || ""))) {
        enableFullMezziMode();
      }
      return originalOpenManagementPanel.apply(this, arguments);
    };
  }

  if (typeof originalCloseManagementPanel === "function") {
    window.closeManagementPanel = function closeManagementPanelAndReleaseMezzi() {
      const result = originalCloseManagementPanel.apply(this, arguments);
      restoreAssignedMezziMode();
      return result;
    };
  }

  if (!document.querySelector('script[data-squadre-mezzi-pictograms]')) {
    const script = document.createElement("script");
    script.src = "./squadre-mezzi-pictograms.js?v=20260727a";
    script.defer = true;
    script.dataset.squadreMezziPictograms = "1";
    document.head.appendChild(script);
  }

  if (!document.querySelector('script[data-today-live-hours-vehicles]')) {
    const script = document.createElement("script");
    script.src = "./today-live-hours-vehicles.js?v=20260730b";
    script.defer = true;
    script.dataset.todayLiveHoursVehicles = "1";
    document.head.appendChild(script);
  }

  if (!document.querySelector('script[data-squad-operator-profile]')) {
    const script = document.createElement("script");
    script.src = "./squad-operator-profile.js?v=20260731a";
    script.defer = true;
    script.dataset.squadOperatorProfile = "1";
    document.head.appendChild(script);
  }
})();
