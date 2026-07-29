/* Sincronizzazione bidirezionale tra Gestione impianti e contabilità e Google Sheet.
   Lettura: proxy GViz già presente. Scrittura: Netlify Function -> Google Apps Script doPost. */
(() => {
  "use strict";

  const VERSION = 1;
  const DEFAULT_INTERVAL_MINUTES = 15;
  const PUSH_DEBOUNCE_MS = 3500;
  const MAX_ROWS = 20000;
  const SYSTEM_HEADERS = ["SYNC_KEY", "IMPIANTO_KEY"];
  const DATA_HEADERS = [
    "N.", "Distretto", "ID SAP", "Denominazione Impianto", "Comune ubicazione Impianto",
    "Via e civico di ubicazione Impianto", "Voce di Riferimento Elenco Prezzi", "Quantità",
    "Frequenza annua minima", "Tipologia di lavorazione / sfalcio", "Coordinate GPS Y / Latitudine",
    "Coordinate GPS X / Longitudine", "u.m.", "Unitario Base d’asta",
    "RIBASSO / Prezzo unitario ribassato", "Totali", "Data esecuzione", "Ora esecuzione",
    "Operatore", "Note", "Stato"
  ];
  const ALL_HEADERS = [...SYSTEM_HEADERS, ...DATA_HEADERS];

  const state = {
    installed: false,
    currentCommessa: null,
    config: null,
    baseOpen: null,
    subscriptions: [],
    pullTimer: null,
    pushTimer: null,
    hydratedCollections: new Set(),
    cache: { commessa: null, plants: [], work: [], prices: [] },
    pushRunning: false,
    pullRunning: false,
    suppressPushUntil: 0
  };

  const aliases = {
    syncKey: ["synckey", "syncid", "id riga", "idriga"],
    plantSyncKey: ["impiantokey", "impiantoid", "idimpianto"],
    numeroProgressivoRiga: ["n", "numero", "numeroprogressivo"],
    distretto: ["distretto"],
    idSap: ["idsap"],
    denominazione: ["denominazioneimpianto", "denominazione", "impianto"],
    comune: ["comuneubicazioneimpianto", "comune"],
    indirizzo: ["viaecivicodiubicazioneimpianto", "indirizzo", "viaecivico"],
    codiceVocePrezzo: ["vocediriferimentoelencoprezzi", "codicevoce", "codiceprezzo", "tipoattivitacodicevoce"],
    quantita: ["quantita"],
    frequenzaAnnua: ["frequenzaannuaminima", "frequenzaannua"],
    tipologiaLavorazione: ["tipologiadilavorazionesfalcio", "tipologialavorazione", "descrizionetipologiaattivita"],
    latitudine: ["coordinategpsylatitudine", "latitudine", "gpsy"],
    longitudine: ["coordinategpsxlongitudine", "longitudine", "gpsx"],
    coordinateUnica: ["coordinategps", "coordinate", "coordinates", "gps"],
    unitaMisura: ["um", "unitadimisura"],
    prezzoBase: ["unitariobasedasta", "prezzobase"],
    prezzoRibassato: ["ribassoprezzounitarioribassato", "prezzoribassato"],
    totale: ["totali", "totale"],
    dataEsecuzione: ["dataesecuzione", "datafatto"],
    oraEsecuzione: ["oraesecuzione", "orafatto"],
    operatoreNome: ["operatore", "operatorenome"],
    note: ["note"],
    stato: ["stato"]
  };

  function collectionName() {
    return typeof getCommesseCollectionName === "function" ? getCommesseCollectionName() : "commesse";
  }

  function canManage() {
    return typeof canManageData !== "function" || canManageData();
  }

  function normalize(value) {
    return String(value ?? "")
      .trim()
      .toLocaleLowerCase("it-IT")
      .normalize("NFD")
      .replace(/[\u0300-\u036f]/g, "")
      .replace(/[^a-z0-9]/g, "");
  }

  function clean(value) {
    return String(value ?? "").trim().replace(/\s+/g, " ");
  }

  function parseFlexibleNumber(value) {
    if (value == null || String(value).trim() === "") return null;
    let text = String(value).trim().replace(/\s/g, "").replace(/[€%]/g, "");
    if (text.includes(",") && text.includes(".")) {
      text = text.lastIndexOf(",") > text.lastIndexOf(".")
        ? text.replace(/\./g, "").replace(",", ".")
        : text.replace(/,/g, "");
    } else if (text.includes(",")) {
      text = text.replace(",", ".");
    }
    const parsed = Number(text);
    return Number.isFinite(parsed) ? parsed : null;
  }

  function fnv1a(value) {
    let hash = 0x811c9dc5;
    const text = String(value ?? "");
    for (let index = 0; index < text.length; index += 1) {
      hash ^= text.charCodeAt(index);
      hash = Math.imul(hash, 0x01000193);
    }
    return (hash >>> 0).toString(36);
  }

  function stableJsonHash(value) {
    const normalizeValue = (entry) => {
      if (Array.isArray(entry)) return entry.map(normalizeValue);
      if (entry && typeof entry === "object") {
        return Object.fromEntries(Object.keys(entry).sort().map((key) => [key, normalizeValue(entry[key])]));
      }
      return entry ?? null;
    };
    return fnv1a(JSON.stringify(normalizeValue(value)));
  }

  function sheetIdentity(rawUrl) {
    const parsed = new URL(String(rawUrl || "").trim());
    if (parsed.protocol !== "https:" || parsed.hostname !== "docs.google.com") {
      throw new Error("Inserisci un link Google Sheet valido.");
    }
    const match = parsed.pathname.match(/^\/spreadsheets\/d\/([A-Za-z0-9_-]+)/);
    if (!match) throw new Error("ID del Google Sheet non riconosciuto.");
    const hashParams = new URLSearchParams(parsed.hash.replace(/^#/, ""));
    const requestedGid = parsed.searchParams.get("gid") || hashParams.get("gid") || "0";
    const gid = /^\d+$/.test(requestedGid) ? requestedGid : "0";
    return { spreadsheetId: match[1], gid, sourceId: `${match[1]}:${gid}` };
  }

  function apiOrigin() {
    const configured = String(window.HERA_API_ORIGIN || "").trim();
    if (configured) return configured.replace(/\/$/, "");
    return /(?:^|\.)netlify\.app$/i.test(location.hostname)
      ? location.origin
      : "https://creative-syrniki-dddbae.netlify.app";
  }

  function endpoint(path) {
    return new URL(path, `${apiOrigin()}/`).href;
  }

  function getFeedback() {
    return document.querySelector("#sheet-two-way-feedback") || document.querySelector("#sheet-url-feedback");
  }

  function setFeedback(message, kind = "") {
    const feedback = getFeedback();
    if (!feedback) return;
    feedback.textContent = message;
    feedback.dataset.kind = kind;
    feedback.style.color = kind === "error" ? "#b91c1c" : kind === "success" ? "#047857" : "";
  }

  function setBusy(button, busy, busyText) {
    if (!button) return;
    if (busy) {
      button.dataset.originalText = button.textContent;
      button.disabled = true;
      button.textContent = busyText;
    } else {
      button.disabled = false;
      button.textContent = button.dataset.originalText || button.textContent;
      delete button.dataset.originalText;
    }
  }

  function injectUi() {
    if (document.querySelector("#sheet-two-way-controls")) return;
    const details = document.querySelector("#sheet-url-import-btn")?.closest("details.import-mode-card");
    const row = details?.querySelector(".import-mode-details");
    if (!details || !row) return;

    const controls = document.createElement("div");
    controls.id = "sheet-two-way-controls";
    controls.style.cssText = "display:grid;grid-template-columns:repeat(auto-fit,minmax(170px,1fr));gap:8px;width:100%;margin-top:10px;padding-top:10px;border-top:1px solid #dbe7e3";
    controls.innerHTML = `
      <button id="sheet-save-link-btn" class="btn" type="button">Salva collegamento</button>
      <button id="sheet-pull-btn" class="btn" type="button">Aggiorna app dal foglio</button>
      <button id="sheet-push-btn" class="btn btn-primary" type="button">Invia app al foglio</button>
      <label style="display:flex;align-items:center;gap:8px;padding:8px 10px;border:1px solid #dbe7e3;border-radius:10px">
        <input id="sheet-auto-sync" type="checkbox">
        <span>Sync automatica</span>
      </label>
      <label style="display:flex;align-items:center;gap:8px;padding:8px 10px;border:1px solid #dbe7e3;border-radius:10px">
        <span>Controlla ogni</span>
        <select id="sheet-sync-interval" style="min-width:80px">
          <option value="5">5 min</option><option value="15">15 min</option>
          <option value="30">30 min</option><option value="60">60 min</option>
        </select>
      </label>
      <p id="sheet-two-way-feedback" class="muted" role="status" aria-live="polite" style="grid-column:1/-1;margin:2px 0 0"></p>
      <small class="muted" style="grid-column:1/-1">Lettura foglio → app tramite GViz. Scrittura app → foglio tramite Apps Script. La prima direzione va confermata manualmente.</small>`;
    row.appendChild(controls);

    document.querySelector("#sheet-save-link-btn")?.addEventListener("click", () => void saveConfigFromUi());
    document.querySelector("#sheet-pull-btn")?.addEventListener("click", (event) => void pullFromSheet({ manual: true, button: event.currentTarget }));
    document.querySelector("#sheet-push-btn")?.addEventListener("click", (event) => void pushToSheet({ manual: true, button: event.currentTarget }));
    document.querySelector("#sheet-auto-sync")?.addEventListener("change", () => void saveConfigFromUi());
    document.querySelector("#sheet-sync-interval")?.addEventListener("change", () => void saveConfigFromUi());
  }

  function renderConfig() {
    injectUi();
    const input = document.querySelector("#sheet-url");
    const auto = document.querySelector("#sheet-auto-sync");
    const interval = document.querySelector("#sheet-sync-interval");
    const controls = document.querySelector("#sheet-two-way-controls");
    if (!controls) return;
    controls.style.display = canManage() ? "grid" : "none";
    if (input && state.config?.sheetUrl && !input.value.trim()) input.value = state.config.sheetUrl;
    if (auto) auto.checked = Boolean(state.config?.enabled);
    if (interval) interval.value = String(state.config?.intervalMinutes || DEFAULT_INTERVAL_MINUTES);
    if (!state.config?.sheetUrl) setFeedback("Incolla il link del Google Sheet e premi “Salva collegamento”.");
    else if (!state.config?.initialized) setFeedback("Collegamento salvato. Scegli una prima direzione: foglio → app oppure app → foglio.");
    else if (state.config.enabled) setFeedback(`Sincronizzazione automatica attiva ogni ${state.config.intervalMinutes} minuti.`, "success");
    else setFeedback("Collegamento attivo. La sincronizzazione automatica è disattivata.");
  }

  async function readConfig(commessa) {
    const raw = commessa?.googleSheetSync || {};
    let sheetUrl = clean(raw.sheetUrl);
    if (!sheetUrl && commessa?.sheetSpreadsheetId) {
      sheetUrl = `https://docs.google.com/spreadsheets/d/${commessa.sheetSpreadsheetId}/edit`;
    }
    return {
      version: Number(raw.version) || VERSION,
      sheetUrl,
      spreadsheetId: clean(raw.spreadsheetId || commessa?.sheetSpreadsheetId),
      gid: String(raw.gid || "0"),
      enabled: Boolean(raw.enabled),
      initialized: Boolean(raw.initialized),
      intervalMinutes: [5, 15, 30, 60].includes(Number(raw.intervalMinutes)) ? Number(raw.intervalMinutes) : DEFAULT_INTERVAL_MINUTES,
      lastPushHash: clean(raw.lastPushHash),
      lastPullHash: clean(raw.lastPullHash)
    };
  }

  async function saveConfigFromUi() {
    if (!state.currentCommessa || !canManage()) return;
    const input = document.querySelector("#sheet-url");
    const sheetUrl = clean(input?.value);
    if (!sheetUrl) return setFeedback("Incolla prima il link del Google Sheet.", "error");
    let identity;
    try {
      identity = sheetIdentity(sheetUrl);
    } catch (error) {
      return setFeedback(error.message, "error");
    }
    const enabled = Boolean(document.querySelector("#sheet-auto-sync")?.checked);
    const intervalMinutes = Number(document.querySelector("#sheet-sync-interval")?.value) || DEFAULT_INTERVAL_MINUTES;
    const previousSourceId = state.config?.spreadsheetId ? `${state.config.spreadsheetId}:${state.config.gid || "0"}` : "";
    const sourceChanged = previousSourceId && previousSourceId !== identity.sourceId;
    const nextConfig = {
      ...(state.config || {}),
      version: VERSION,
      sheetUrl,
      spreadsheetId: identity.spreadsheetId,
      gid: identity.gid,
      enabled,
      intervalMinutes,
      initialized: sourceChanged ? false : Boolean(state.config?.initialized),
      lastPushHash: sourceChanged ? "" : clean(state.config?.lastPushHash),
      lastPullHash: sourceChanged ? "" : clean(state.config?.lastPullHash)
    };
    const operatorName = typeof getOperatorDisplayName === "function" ? getOperatorDisplayName() : "";
    await db.collection(collectionName()).doc(state.currentCommessa.id).set({
      sheetSpreadsheetId: identity.spreadsheetId,
      googleSheetSync: {
        ...nextConfig,
        updatedAt: firebase.firestore.FieldValue.serverTimestamp(),
        updatedBy: auth.currentUser?.uid || "",
        updatedByName: operatorName
      }
    }, { merge: true });
    state.config = nextConfig;
    state.currentCommessa = { ...state.currentCommessa, sheetSpreadsheetId: identity.spreadsheetId, googleSheetSync: nextConfig };
    setupTimersAndSubscriptions();
    renderConfig();
    setFeedback(sourceChanged
      ? "Nuovo foglio collegato. Conferma manualmente la prima direzione di sincronizzazione."
      : "Collegamento Google Sheet salvato.", "success");
  }

  function stopSubscriptions() {
    state.subscriptions.splice(0).forEach((unsubscribe) => {
      try { unsubscribe?.(); } catch (_) { /* nessuna azione */ }
    });
    clearInterval(state.pullTimer);
    clearTimeout(state.pushTimer);
    state.pullTimer = null;
    state.pushTimer = null;
    state.hydratedCollections.clear();
    state.cache = { commessa: null, plants: [], work: [], prices: [] };
  }

  function markHydrated(name) {
    state.hydratedCollections.add(name);
    if (state.hydratedCollections.size >= 4) schedulePush();
  }

  function setupTimersAndSubscriptions() {
    stopSubscriptions();
    if (!state.currentCommessa || !state.config?.sheetUrl || !canManage()) return;
    const ref = db.collection(collectionName()).doc(state.currentCommessa.id);
    state.subscriptions.push(ref.onSnapshot((snapshot) => {
      state.cache.commessa = { id: snapshot.id, ...snapshot.data() };
      markHydrated("commessa");
      schedulePush();
    }, (error) => console.error("[Sheet Sync] commessa:", error)));
    state.subscriptions.push(ref.collection("impiantiFisici").onSnapshot((snapshot) => {
      state.cache.plants = snapshot.docs.map((doc) => ({ id: doc.id, ...doc.data() }));
      markHydrated("plants");
      schedulePush();
    }, (error) => console.error("[Sheet Sync] impianti:", error)));
    state.subscriptions.push(ref.collection("lavorazioni").onSnapshot((snapshot) => {
      state.cache.work = snapshot.docs.map((doc) => ({ id: doc.id, ...doc.data() }));
      markHydrated("work");
      schedulePush();
    }, (error) => console.error("[Sheet Sync] lavorazioni:", error)));
    state.subscriptions.push(ref.collection("prezziario").onSnapshot((snapshot) => {
      state.cache.prices = snapshot.docs.map((doc) => ({ id: doc.id, ...doc.data() }));
      markHydrated("prices");
      schedulePush();
    }, (error) => console.error("[Sheet Sync] prezziario:", error)));

    if (state.config.enabled && state.config.initialized) {
      const milliseconds = Math.max(5, Number(state.config.intervalMinutes) || DEFAULT_INTERVAL_MINUTES) * 60 * 1000;
      state.pullTimer = window.setInterval(() => void pullFromSheet({ manual: false }), milliseconds);
    }
  }

  function schedulePush() {
    clearTimeout(state.pushTimer);
    if (!state.config?.enabled || !state.config?.initialized || !navigator.onLine || !canManage()) return;
    if (Date.now() < state.suppressPushUntil || state.hydratedCollections.size < 4) return;
    state.pushTimer = window.setTimeout(() => void pushToSheet({ manual: false }), PUSH_DEBOUNCE_MS);
  }

  function buildRows() {
    const core = window.InreteWorkItemsV2;
    const plants = new Map(state.cache.plants.map((plant) => [plant.id, plant]));
    const priceMap = core?.buildPriceMap ? core.buildPriceMap(state.cache.prices) : new Map();
    const generalDiscount = state.cache.commessa?.percentualeRibassoGenerale ?? 0;
    return state.cache.work
      .slice()
      .sort((a, b) => (parseFlexibleNumber(a.numeroProgressivoRiga) || 0) - (parseFlexibleNumber(b.numeroProgressivoRiga) || 0))
      .map((rawWork, index) => {
        const work = core?.enrichWorkItem ? core.enrichWorkItem(rawWork, priceMap, generalDiscount) : rawWork;
        const plant = plants.get(work.impiantoId) || {};
        const syncKey = clean(work.sheetSyncKey || work.id || `work-${index + 1}`);
        const plantSyncKey = clean(plant.sheetSyncKey || plant.id || work.impiantoId || `plant-${index + 1}`);
        return [
          syncKey, plantSyncKey, work.numeroProgressivoRiga || index + 1, plant.distretto || work.distretto || "",
          plant.idSap || work.idSap || "", plant.denominazione || work.denominazione || "",
          plant.comune || work.comune || "", plant.indirizzo || work.indirizzo || "",
          work.codiceVocePrezzo || work.codicePrezzo || "", work.quantita ?? "", work.frequenzaAnnua ?? "",
          work.tipologiaLavorazione || work.tipologiaIntervento || "",
          plant.latitudine ?? work.latitudine ?? plant.gpsY ?? "", plant.longitudine ?? work.longitudine ?? plant.gpsX ?? "",
          work.unitaMisura || "", work.prezzoBase ?? "", work.prezzoRibassato ?? "", work.totale ?? "",
          work.stato === "FATTO" ? (work.dataEsecuzione || "") : "",
          work.stato === "FATTO" ? (work.oraEsecuzione || "") : "",
          work.stato === "FATTO" ? (work.operatoreNome || work.operatore || "") : "",
          work.note || "", String(work.stato || "DA FARE").toUpperCase()
        ];
      });
  }

  async function authorizedFetch(url, options = {}) {
    const user = auth.currentUser;
    if (!user) throw new Error("Accedi di nuovo prima di sincronizzare il foglio.");
    const token = await user.getIdToken();
    return fetch(url, {
      ...options,
      headers: { ...(options.headers || {}), Authorization: `Bearer ${token}` },
      cache: "no-store"
    });
  }

  async function responseError(response) {
    try {
      const payload = await response.json();
      return payload?.error || payload?.detail || `Errore HTTP ${response.status}`;
    } catch (_) {
      return `Errore HTTP ${response.status}`;
    }
  }

  async function persistSyncResult(patch) {
    if (!state.currentCommessa) return;
    const nextConfig = { ...(state.config || {}), ...patch };
    await db.collection(collectionName()).doc(state.currentCommessa.id).set({
      googleSheetSync: {
        ...nextConfig,
        updatedAt: firebase.firestore.FieldValue.serverTimestamp(),
        updatedBy: auth.currentUser?.uid || "",
        updatedByName: typeof getOperatorDisplayName === "function" ? getOperatorDisplayName() : ""
      }
    }, { merge: true });
    state.config = nextConfig;
  }

  async function pushToSheet({ manual = false, button = null } = {}) {
    if (state.pushRunning || state.pullRunning || !state.currentCommessa || !canManage()) return;
    if (!navigator.onLine) return manual && setFeedback("Sei offline. I dati restano in Firebase e saranno inviati quando torni online.", "error");
    const sheetUrl = clean(document.querySelector("#sheet-url")?.value || state.config?.sheetUrl);
    if (!sheetUrl) return manual && setFeedback("Salva prima il collegamento al Google Sheet.", "error");
    if (!state.config?.sheetUrl || state.config.sheetUrl !== sheetUrl) {
      await saveConfigFromUi();
      if (!state.config?.sheetUrl) return;
    }
    let identity;
    try { identity = sheetIdentity(sheetUrl); } catch (error) { return setFeedback(error.message, "error"); }
    if (state.hydratedCollections.size < 4) {
      setupTimersAndSubscriptions();
      return manual && setFeedback("Caricamento dati della commessa in corso. Riprova tra pochi secondi.");
    }

    state.pushRunning = true;
    setBusy(button, true, "Invio…");
    if (manual) setFeedback("Invio dei dati dell’app al Google Sheet…");
    try {
      const rows = buildRows();
      if (rows.length > MAX_ROWS) throw new Error(`Troppe righe da sincronizzare: massimo ${MAX_ROWS}.`);
      const hash = stableJsonHash({ headers: ALL_HEADERS, rows });
      if (!manual && hash === state.config?.lastPushHash) return;
      const response = await authorizedFetch(endpoint("/api/google-sheet-sync"), {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          action: "replaceRows",
          sheetUrl,
          spreadsheetId: identity.spreadsheetId,
          gid: identity.gid,
          commessaId: state.currentCommessa.id,
          commessaName: state.currentCommessa.nome || "",
          headers: ALL_HEADERS,
          rows,
          operationId: `${state.currentCommessa.id}-${Date.now()}-${hash}`
        })
      });
      if (!response.ok) throw new Error(await responseError(response));
      const payload = await response.json();
      if (!payload?.ok) throw new Error(payload?.error || "Scrittura Google Sheet non riuscita.");
      await persistSyncResult({
        version: VERSION,
        sheetUrl,
        spreadsheetId: identity.spreadsheetId,
        gid: identity.gid,
        initialized: true,
        lastDirection: "APP_TO_SHEET",
        lastPushHash: hash,
        lastPushAt: firebase.firestore.FieldValue.serverTimestamp()
      });
      setFeedback(`${rows.length} righe inviate al Google Sheet.`, "success");
      renderConfig();
    } catch (error) {
      console.error("[Sheet Sync] invio non riuscito", error);
      setFeedback(`${error.message || error}`, "error");
    } finally {
      state.pushRunning = false;
      setBusy(button, false);
    }
  }

  function mapped(row, field) {
    const normalized = Object.fromEntries(Object.entries(row).map(([key, value]) => [normalize(key), value]));
    for (const alias of aliases[field] || []) {
      const key = normalize(alias);
      if (Object.prototype.hasOwnProperty.call(normalized, key)) return normalized[key];
    }
    return "";
  }

  function parseSheetRows(csv) {
    const workbook = XLSX.read(csv, { type: "string" });
    const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
    const rawRows = XLSX.utils.sheet_to_json(firstSheet, { defval: "", raw: false });
    return rawRows.map((raw, index) => {
      const singleCoordinates = mapped(raw, "coordinateUnica");
      const latRaw = mapped(raw, "latitudine") || singleCoordinates;
      const lonRaw = mapped(raw, "longitudine");
      const gps = window.HeraCoordinateRepair?.diagnose(latRaw, lonRaw) || {
        valid: false, latitude: null, longitude: null, status: "MISSING", message: "Coordinate non disponibili",
        rawLatitude: String(latRaw || ""), rawLongitude: String(lonRaw || "")
      };
      const row = {
        syncKey: clean(mapped(raw, "syncKey")),
        plantSyncKey: clean(mapped(raw, "plantSyncKey")),
        numeroProgressivoRiga: parseFlexibleNumber(mapped(raw, "numeroProgressivoRiga")) || index + 1,
        distretto: clean(mapped(raw, "distretto")), idSap: clean(mapped(raw, "idSap")),
        denominazione: clean(mapped(raw, "denominazione")), comune: clean(mapped(raw, "comune")),
        indirizzo: clean(mapped(raw, "indirizzo")), codiceVocePrezzo: clean(mapped(raw, "codiceVocePrezzo")).toUpperCase(),
        quantita: parseFlexibleNumber(mapped(raw, "quantita")), frequenzaAnnua: parseFlexibleNumber(mapped(raw, "frequenzaAnnua")),
        tipologiaLavorazione: clean(mapped(raw, "tipologiaLavorazione")),
        latitudine: gps.valid ? gps.latitude : latRaw, longitudine: gps.valid ? gps.longitude : lonRaw,
        coordinateStatus: gps.status, coordinateIssue: gps.message,
        coordinateLatitudineOriginale: gps.rawLatitude, coordinateLongitudineOriginale: gps.rawLongitude,
        unitaMisura: clean(mapped(raw, "unitaMisura")), prezzoBase: parseFlexibleNumber(mapped(raw, "prezzoBase")),
        prezzoRibassato: parseFlexibleNumber(mapped(raw, "prezzoRibassato")), totale: parseFlexibleNumber(mapped(raw, "totale")),
        dataEsecuzione: clean(mapped(raw, "dataEsecuzione")), oraEsecuzione: clean(mapped(raw, "oraEsecuzione")),
        operatoreNome: clean(mapped(raw, "operatoreNome")), note: clean(mapped(raw, "note")),
        stato: clean(mapped(raw, "stato") || "DA FARE").toUpperCase()
      };
      if (!row.plantSyncKey) row.plantSyncKey = row.idSap || `plant-${fnv1a(`${row.denominazione}|${row.comune}|${row.indirizzo}`)}`;
      if (!row.syncKey) row.syncKey = `work-${fnv1a(`${row.plantSyncKey}|${row.codiceVocePrezzo}|${row.numeroProgressivoRiga}`)}`;
      return row;
    }).filter((row) => row.denominazione);
  }

  function documentId(prefix, sourceId, key) {
    return `${prefix}_${fnv1a(`${sourceId}|${key}`)}`;
  }

  async function commitOperations(operations) {
    for (let index = 0; index < operations.length; index += 350) {
      const batch = db.batch();
      operations.slice(index, index + 350).forEach(({ type, ref, data }) => {
        if (type === "delete") batch.delete(ref);
        else batch.set(ref, data, { merge: true });
      });
      await batch.commit();
    }
  }

  async function applyRowsToFirestore(rows, identity) {
    const commessaRef = db.collection(collectionName()).doc(state.currentCommessa.id);
    const [plantSnapshot, workSnapshot, priceSnapshot] = await Promise.all([
      commessaRef.collection("impiantiFisici").get(),
      commessaRef.collection("lavorazioni").get(),
      commessaRef.collection("prezziario").get()
    ]);
    const plants = plantSnapshot.docs.map((doc) => ({ id: doc.id, ...doc.data() }));
    const works = workSnapshot.docs.map((doc) => ({ id: doc.id, ...doc.data() }));
    const prices = priceSnapshot.docs.map((doc) => ({ id: doc.id, ...doc.data() }));
    const priceMap = window.InreteWorkItemsV2?.buildPriceMap?.(prices) || new Map();
    const generalDiscount = state.cache.commessa?.percentualeRibassoGenerale ?? state.currentCommessa.percentualeRibassoGenerale ?? 0;
    const sourceId = identity.sourceId;
    const operations = [];
    const incomingPlantIds = new Set();
    const incomingWorkIds = new Set();
    const existingPlantBySyncKey = new Map(plants.filter((item) => item.sheetSyncSourceId === sourceId).map((item) => [item.sheetSyncKey, item]));
    const existingPlantBySap = new Map(plants.filter((item) => item.idSap).map((item) => [normalize(item.idSap), item]));
    const existingWorkBySyncKey = new Map(works.filter((item) => item.sheetSyncSourceId === sourceId).map((item) => [item.sheetSyncKey, item]));
    const existingWorkByIdentity = new Map(works.map((item) => [
      `${item.impiantoId || ""}|${normalize(item.codiceVocePrezzo || item.codicePrezzo)}|${Number(item.numeroProgressivoRiga) || ""}`,
      item
    ]));
    const now = firebase.firestore.FieldValue.serverTimestamp();
    const userId = auth.currentUser?.uid || "";
    const userName = typeof getOperatorDisplayName === "function" ? getOperatorDisplayName() : "";

    for (const row of rows) {
      let plant = existingPlantBySyncKey.get(row.plantSyncKey)
        || (row.idSap ? existingPlantBySap.get(normalize(row.idSap)) : null);
      const plantId = plant?.id || documentId("sheetplant", sourceId, row.plantSyncKey);
      incomingPlantIds.add(plantId);
      const plantRef = commessaRef.collection("impiantiFisici").doc(plantId);
      const plantData = {
        commessaId: state.currentCommessa.id,
        sheetManaged: true, sheetSyncSourceId: sourceId, sheetSyncKey: row.plantSyncKey,
        numeroProgressivoImpianto: plant?.numeroProgressivoImpianto || row.numeroProgressivoRiga,
        distretto: row.distretto, idSap: row.idSap, denominazione: row.denominazione,
        comune: row.comune, indirizzo: row.indirizzo,
        latitudine: row.latitudine, longitudine: row.longitudine,
        coordinateStatus: row.coordinateStatus, coordinateIssue: row.coordinateIssue,
        coordinateLatitudineOriginale: row.coordinateLatitudineOriginale,
        coordinateLongitudineOriginale: row.coordinateLongitudineOriginale,
        updatedAt: now, updatedBy: userId, updatedByName: userName
      };
      if (!plant) Object.assign(plantData, { createdAt: now, createdBy: userId });
      operations.push({ type: "set", ref: plantRef, data: plantData });

      const oldWork = existingWorkBySyncKey.get(row.syncKey)
        || existingWorkByIdentity.get(`${plantId}|${normalize(row.codiceVocePrezzo)}|${Number(row.numeroProgressivoRiga) || ""}`);
      const workId = oldWork?.id || documentId("sheetwork", sourceId, row.syncKey);
      incomingWorkIds.add(workId);
      const linkedPrice = window.InreteWorkItemsV2?.resolvePriceItem?.(priceMap, row.codiceVocePrezzo);
      const importedEconomic = row.prezzoBase != null || row.prezzoRibassato != null || row.totale != null;
      let economic = {
        unitaMisura: row.unitaMisura,
        prezzoBase: row.prezzoBase,
        prezzoRibassato: row.prezzoRibassato,
        totale: row.totale,
        priceOverride: importedEconomic,
        priceListLinkStatus: row.codiceVocePrezzo ? (linkedPrice ? "LINKED" : "MISSING") : "EMPTY"
      };
      if (linkedPrice && !importedEconomic && window.InreteWorkItemsV2?.enrichWorkItem) {
        economic = {
          ...economic,
          ...window.InreteWorkItemsV2.enrichWorkItem({ ...row, priceOverride: false }, priceMap, generalDiscount),
          priceOverride: false,
          priceListLinkStatus: "LINKED"
        };
      }
      const workRef = commessaRef.collection("lavorazioni").doc(workId);
      const workData = {
        commessaId: state.currentCommessa.id, impiantoId: plantId,
        sheetManaged: true, sheetSyncSourceId: sourceId, sheetSyncKey: row.syncKey,
        numeroProgressivoRiga: row.numeroProgressivoRiga, codiceVocePrezzo: row.codiceVocePrezzo,
        quantita: row.quantita, frequenzaAnnua: row.frequenzaAnnua,
        tipologiaLavorazione: row.tipologiaLavorazione,
        dataEsecuzione: row.stato === "FATTO" ? row.dataEsecuzione : "",
        oraEsecuzione: row.stato === "FATTO" ? row.oraEsecuzione : "",
        operatoreNome: row.stato === "FATTO" ? row.operatoreNome : "",
        note: row.note, stato: row.stato || "DA FARE", done: row.stato === "FATTO",
        ...economic, updatedAt: now, updatedBy: userId, updatedByName: userName
      };
      if (!oldWork) Object.assign(workData, { createdAt: now, createdBy: userId });
      operations.push({ type: "set", ref: workRef, data: workData });
    }

    works.filter((item) => item.sheetManaged && item.sheetSyncSourceId === sourceId && !incomingWorkIds.has(item.id))
      .forEach((item) => operations.push({ type: "delete", ref: commessaRef.collection("lavorazioni").doc(item.id) }));
    plants.filter((item) => item.sheetManaged && item.sheetSyncSourceId === sourceId && !incomingPlantIds.has(item.id))
      .forEach((item) => {
        operations.push({ type: "delete", ref: commessaRef.collection("impiantiFisici").doc(item.id) });
        operations.push({ type: "delete", ref: commessaRef.collection("impianti").doc(item.id) });
      });

    await commitOperations(operations);
    return { rows: rows.length, operations: operations.length };
  }

  async function refreshAccountingScreen() {
    if (!state.currentCommessa || !state.baseOpen) return;
    await state.baseOpen(state.currentCommessa);
    if (window.AccountingV2?.synchronizeOperationalModel) {
      await window.AccountingV2.synchronizeOperationalModel({ debug: true });
      await state.baseOpen(state.currentCommessa);
    }
  }

  async function pullFromSheet({ manual = false, button = null } = {}) {
    if (state.pullRunning || state.pushRunning || !state.currentCommessa || !canManage()) return;
    if (!navigator.onLine) return manual && setFeedback("Sei offline. Impossibile leggere il Google Sheet.", "error");
    const sheetUrl = clean(document.querySelector("#sheet-url")?.value || state.config?.sheetUrl);
    if (!sheetUrl) return manual && setFeedback("Salva prima il collegamento al Google Sheet.", "error");
    if (!state.config?.sheetUrl || state.config.sheetUrl !== sheetUrl) {
      await saveConfigFromUi();
      if (!state.config?.sheetUrl) return;
    }
    let identity;
    try { identity = sheetIdentity(sheetUrl); } catch (error) { return setFeedback(error.message, "error"); }

    state.pullRunning = true;
    setBusy(button, true, "Lettura…");
    if (manual) setFeedback("Lettura del Google Sheet e aggiornamento dell’app…");
    try {
      const importUrl = new URL(endpoint("/api/google-sheet-import"));
      importUrl.searchParams.set("url", sheetUrl);
      const response = await fetch(importUrl.href, { headers: { Accept: "text/csv" }, cache: "no-store" });
      if (!response.ok) throw new Error(await responseError(response));
      const csv = await response.text();
      if (!csv.trim()) throw new Error("Il Google Sheet è vuoto.");
      const rows = parseSheetRows(csv);
      if (rows.length > MAX_ROWS) throw new Error(`Troppe righe: massimo ${MAX_ROWS}.`);
      const hash = stableJsonHash(rows);
      if (!manual && hash === state.config?.lastPullHash) return;
      const report = await applyRowsToFirestore(rows, identity);
      state.suppressPushUntil = Date.now() + 10000;
      await persistSyncResult({
        version: VERSION,
        sheetUrl,
        spreadsheetId: identity.spreadsheetId,
        gid: identity.gid,
        initialized: true,
        lastDirection: "SHEET_TO_APP",
        lastPullHash: hash,
        lastPullAt: firebase.firestore.FieldValue.serverTimestamp()
      });
      await refreshAccountingScreen();
      setFeedback(`${report.rows} righe lette dal foglio e applicate all’app.`, "success");
      renderConfig();
    } catch (error) {
      console.error("[Sheet Sync] lettura non riuscita", error);
      setFeedback(`${error.message || error}`, "error");
    } finally {
      state.pullRunning = false;
      setBusy(button, false);
    }
  }

  async function selectCommessa(commessa) {
    stopSubscriptions();
    state.currentCommessa = commessa;
    state.config = await readConfig(commessa);
    injectUi();
    renderConfig();
    setupTimersAndSubscriptions();
  }

  function install() {
    if (state.installed) return;
    if (typeof db === "undefined" || typeof auth === "undefined" || !window.AccountingV2 || !window.XLSX) {
      window.setTimeout(install, 250);
      return;
    }
    state.installed = true;
    state.baseOpen = window.AccountingV2.open.bind(window.AccountingV2);
    window.AccountingV2.open = async function openWithGoogleSheetSync(commessa) {
      const result = await state.baseOpen(commessa);
      await selectCommessa(commessa);
      return result;
    };
    window.addEventListener("online", () => {
      setFeedback("Connessione ripristinata. Controllo sincronizzazione…");
      schedulePush();
      if (state.config?.enabled && state.config?.initialized) void pullFromSheet({ manual: false });
    });
    window.addEventListener("offline", () => setFeedback("Offline: le modifiche restano in Firebase e saranno sincronizzate più tardi."));
    window.HeraGoogleSheetSync = Object.freeze({
      push: () => pushToSheet({ manual: true }),
      pull: () => pullFromSheet({ manual: true }),
      selectCommessa,
      version: VERSION
    });
  }

  install();
})();
