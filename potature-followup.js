(() => {
  "use strict";

  const COMMESSA_ID = "potature-abbattimenti";
  const PHASES = Object.freeze({
    raccolta: Object.freeze({
      label: "Raccolta",
      taskLabel: "Mucchia",
      code: "MUCCHIA",
      idSuffix: "raccolta"
    }),
    ceppi: Object.freeze({
      label: "Ceppi",
      taskLabel: "Ceppo",
      code: "CEPPO",
      idSuffix: "ceppi"
    })
  });
  const METHODS = Object.freeze({
    raccolta: Object.freeze({
      ragno: "Con ragno",
      mano: "A mano"
    }),
    ceppi: Object.freeze({
      robotino: "Con robotino",
      t15: "Con T15"
    })
  });

  const text = (value) => String(value ?? "").trim();
  const esc = (value) => text(value).replace(/[&<>'"]/g, (character) => ({
    "&": "&amp;", "<": "&lt;", ">": "&gt;", "'": "&#39;", '"': "&quot;"
  }[character]));

  function stableHash(value) {
    let hash = 2166136261;
    for (const character of text(value)) {
      hash ^= character.charCodeAt(0);
      hash = Math.imul(hash, 16777619);
    }
    return (hash >>> 0).toString(36);
  }

  function slug(value) {
    return text(value)
      .normalize("NFD")
      .replace(/[\u0300-\u036f]/g, "")
      .toLowerCase()
      .replace(/[^a-z0-9]+/g, "-")
      .replace(/^-|-$/g, "")
      .slice(0, 80);
  }

  function sourceIdentity(source = {}) {
    return text(source.potatureOrigineId)
      || text(source.id)
      || text(Array.isArray(source.sourceIds) ? source.sourceIds[0] : "")
      || text(source.idSap)
      || `${text(source.denominazione)}|${text(source.comune)}|${text(source.indirizzo)}`;
  }

  function phaseFor(item = {}) {
    const phase = text(item.potatureFase).toLowerCase();
    return Object.prototype.hasOwnProperty.call(PHASES, phase) ? phase : "";
  }

  function isOriginal(item = {}) {
    const special = item.potatureAbbattimenti === true
      || text(item.messaggioWhazzupTipo) === "POTATURE_ABBATTIMENTI";
    return special && !phaseFor(item);
  }

  function taskDocumentId(source, phase) {
    const config = PHASES[phase];
    if (!config) throw new Error("Fase Potature Abbattimenti non valida.");
    const identity = sourceIdentity(source);
    const base = slug(identity) || `albero-${stableHash(identity)}`;
    return `${base}--${config.idSuffix}`.slice(0, 140);
  }

  function taskIdSap(source, phase) {
    const sourceIdSap = text(source?.idSap);
    const base = sourceIdSap || `ALB-${stableHash(sourceIdentity(source)).toUpperCase()}`;
    return `${base}-${phase === "raccolta" ? "RAC" : "CEP"}`;
  }

  function copyTreeFields(source = {}) {
    const keys = [
      "comune", "indirizzo", "descrizioneVia", "area", "competenza",
      "gpsY", "gpsX", "latitudine", "longitudine", "coordinate",
      "tipologia", "tipologiaImpianto", "alberoCatasto", "schedaParziale",
      "schedaVersione", "campiDaCompletare", "numeroPunto", "codiceAlbero",
      "specieAlbero", "dettagliCatastoPrimiSei", "catastoFonte", "catastoDataset",
      "noteOperatore"
    ];
    return keys.reduce((copy, key) => {
      if (source[key] !== undefined && source[key] !== null) copy[key] = source[key];
      return copy;
    }, {});
  }

  function buildTaskPayload(source, phase, method, metadata = {}) {
    const config = PHASES[phase];
    const methodLabel = METHODS[phase]?.[method];
    if (!config || !methodLabel) throw new Error("Scelta di lavorazione non valida.");

    const name = text(source?.denominazione || source?.nome || "Albero");
    const requestedWork = `${config.taskLabel} ${methodLabel.toLowerCase()}`;
    const operatorName = text(metadata.operatorName || "Operatore");
    const timestamp = metadata.timestamp || new Date();

    return {
      ...copyTreeFields(source),
      commessaId: COMMESSA_ID,
      idSap: taskIdSap(source, phase),
      denominazione: name,
      nome: name,
      potatureAbbattimenti: true,
      potatureFase: phase,
      potatureFaseLabel: config.label,
      potatureMetodo: method,
      potatureMetodoLabel: methodLabel,
      potatureOrigineId: sourceIdentity(source),
      potatureOrigineIdSap: text(source?.idSap),
      tipologiaIntervento: config.code,
      tipologiaLavorazione: requestedWork,
      lavorazioniRichieste: requestedWork,
      codicePrezzo: config.code,
      voceRiferimento: config.code,
      hasOrdinario: false,
      hasStraordinario: true,
      tipoManutenzione: "Straordinaria",
      hasNote: Boolean(text(source?.noteOperatore)),
      preparatoDaUid: text(metadata.operatorUid),
      preparatoDaNome: operatorName,
      preparatoAt: timestamp,
      updatedAt: timestamp
    };
  }

  function buildTasks(source, selections = {}, metadata = {}) {
    return Object.keys(PHASES).flatMap((phase) => {
      const method = text(selections[phase]).toLowerCase();
      if (!METHODS[phase]?.[method]) return [];
      return [{
        id: taskDocumentId(source, phase),
        phase,
        method,
        payload: buildTaskPayload(source, phase, method, metadata)
      }];
    });
  }

  function existingSelections(source, items = []) {
    const identity = sourceIdentity(source);
    return items.reduce((result, item) => {
      const phase = phaseFor(item);
      if (!phase || text(item.potatureOrigineId) !== identity) return result;
      const method = text(item.potatureMetodo).toLowerCase();
      if (METHODS[phase]?.[method]) result[phase] = method;
      return result;
    }, { raccolta: "", ceppi: "" });
  }

  async function saveTasks(options, tasks) {
    if (!tasks.length) return;
    const store = options.store;
    if (!store?.collection || !store?.batch) throw new Error("Database non disponibile. Controlla la connessione e riprova.");
    const collectionName = text(options.collectionName || "commesse");
    const commessaId = text(options.commessaId);
    if (commessaId !== COMMESSA_ID) throw new Error("Il form è disponibile solo nella commessa Potature Abbattimenti.");
    const ref = store.collection(collectionName).doc(commessaId).collection("impianti");
    const batch = store.batch();
    tasks.forEach((task) => batch.set(ref.doc(task.id), task.payload, { merge: true }));
    await batch.commit();
  }

  function ensureModal() {
    let modal = document.getElementById("potature-followup-modal");
    if (modal) return modal;
    modal = document.createElement("section");
    modal.id = "potature-followup-modal";
    modal.className = "potature-followup-modal hidden";
    modal.setAttribute("role", "dialog");
    modal.setAttribute("aria-modal", "true");
    modal.setAttribute("aria-hidden", "true");
    modal.setAttribute("aria-labelledby", "potature-followup-title");
    modal.innerHTML = `
      <form class="potature-followup-screen">
        <header class="potature-followup-head">
          <button class="btn potature-followup-close" type="button" aria-label="Chiudi">← Indietro</button>
          <div><p>Potature Abbattimenti</p><h2 id="potature-followup-title">Prepara la fine</h2></div>
        </header>
        <div class="potature-followup-body">
          <p class="potature-followup-plant"></p>
          <p class="potature-followup-intro">Le due domande sono facoltative. Seleziona solo le attività che devono essere create.</p>
          <fieldset>
            <legend><span>1</span><strong>Mucchia</strong></legend>
            <label><input type="radio" name="raccolta" value="" checked><span><b>Nessuna</b><small>Non creare un impianto in Raccolta</small></span></label>
            <label><input type="radio" name="raccolta" value="ragno"><span><b>Con ragno</b><small>Crea l’impianto nella vista Raccolta</small></span></label>
            <label><input type="radio" name="raccolta" value="mano"><span><b>A mano</b><small>Crea l’impianto nella vista Raccolta</small></span></label>
          </fieldset>
          <fieldset>
            <legend><span>2</span><strong>Ceppo</strong></legend>
            <label><input type="radio" name="ceppi" value="" checked><span><b>Nessuno</b><small>Non creare un impianto in Ceppi</small></span></label>
            <label><input type="radio" name="ceppi" value="robotino"><span><b>Con robotino</b><small>Crea l’impianto nella vista Ceppi</small></span></label>
            <label><input type="radio" name="ceppi" value="t15"><span><b>Con T15</b><small>Crea l’impianto nella vista Ceppi</small></span></label>
          </fieldset>
          <p class="potature-followup-note">Le schede già create vengono aggiornate senza duplicati. Lasciare una scelta vuota non elimina attività esistenti.</p>
          <p class="potature-followup-feedback" role="status" aria-live="polite"></p>
        </div>
        <footer class="potature-followup-actions">
          <p>Dopo il ritorno alla scheda, premi il pulsante speciale <strong>TERMINATO</strong>.</p>
          <button class="btn btn-primary potature-followup-save" type="submit">SALVA E TORNA</button>
        </footer>
      </form>`;
    document.body.appendChild(modal);
    return modal;
  }

  function close() {
    const modal = document.getElementById("potature-followup-modal");
    if (!modal) return;
    modal.classList.add("hidden");
    modal.setAttribute("aria-hidden", "true");
    document.body.classList.remove("potature-followup-open");
  }

  function open(options = {}) {
    if (!isOriginal(options.source)) throw new Error("Il form è disponibile solo per gli alberi originali di Potature Abbattimenti.");
    const modal = ensureModal();
    const form = modal.querySelector("form");
    const feedback = modal.querySelector(".potature-followup-feedback");
    const saveButton = modal.querySelector(".potature-followup-save");
    const previous = existingSelections(options.source, options.existingItems || []);
    form.reset();
    form.elements.raccolta.value = previous.raccolta;
    form.elements.ceppi.value = previous.ceppi;
    modal.querySelector(".potature-followup-plant").innerHTML = `<strong>${esc(options.source.denominazione || options.source.nome || "Albero")}</strong><span>${esc(options.source.comune || "")} ${esc(options.source.indirizzo || "")}</span>`;
    feedback.textContent = "";
    feedback.className = "potature-followup-feedback";
    saveButton.disabled = false;

    const finish = (result) => {
      close();
      if (typeof options.onComplete === "function") options.onComplete(result);
    };
    const onClose = () => done({ tasks: [], skipped: true });
    const onBackdrop = (event) => { if (event.target === modal) onClose(); };
    const onKeydown = (event) => { if (event.key === "Escape") onClose(); };
    const cleanup = () => {
      modal.querySelector(".potature-followup-close")?.removeEventListener("click", onClose);
      modal.removeEventListener("click", onBackdrop);
      document.removeEventListener("keydown", onKeydown);
      form.removeEventListener("submit", onSubmit);
    };
    const done = (result) => {
      cleanup();
      finish(result);
    };
    const onSubmit = async (event) => {
      event.preventDefault();
      const values = new FormData(form);
      const timestamp = options.timestampFactory ? options.timestampFactory() : new Date();
      const tasks = buildTasks(options.source, {
        raccolta: values.get("raccolta"),
        ceppi: values.get("ceppi")
      }, {
        operatorUid: options.operatorUid,
        operatorName: options.operatorName,
        timestamp
      });
      saveButton.disabled = true;
      feedback.textContent = tasks.length ? "Salvataggio delle attività in corso…" : "Nessuna attività aggiuntiva selezionata.";
      try {
        await saveTasks(options, tasks);
        feedback.classList.add("success");
        feedback.textContent = tasks.length
          ? `Salvate: ${tasks.map((task) => PHASES[task.phase].label).join(" e ")}. Ora puoi premere TERMINATO.`
          : "Nessuna attività creata. Ora puoi premere TERMINATO.";
        window.setTimeout(() => done({ tasks, skipped: false }), 650);
      } catch (error) {
        feedback.classList.add("error");
        feedback.textContent = error?.message || "Impossibile salvare le attività.";
        saveButton.disabled = false;
      }
    };

    modal.querySelector(".potature-followup-close")?.addEventListener("click", onClose);
    modal.addEventListener("click", onBackdrop);
    document.addEventListener("keydown", onKeydown);
    form.addEventListener("submit", onSubmit);
    modal.classList.remove("hidden");
    modal.setAttribute("aria-hidden", "false");
    document.body.classList.add("potature-followup-open");
    modal.scrollTop = 0;
    modal.querySelector(".potature-followup-close")?.focus({ preventScroll: true });
  }

  const api = Object.freeze({
    commessaId: COMMESSA_ID,
    phases: PHASES,
    methods: METHODS,
    phaseFor,
    isOriginal,
    sourceIdentity,
    taskDocumentId,
    taskIdSap,
    buildTaskPayload,
    buildTasks,
    existingSelections,
    saveTasks,
    open,
    close
  });

  if (typeof window !== "undefined") window.HeraPotatureFollowup = api;
  if (typeof module !== "undefined" && module.exports) module.exports = api;
})();

(() => {
  "use strict";

  if (typeof window === "undefined" || typeof document === "undefined") return;
  if (window.HeraSpecialTerminato?.installed) return;

  const STYLE_ID = "hera-special-terminato-style";
  const TERMINATED_FIELD = "specialTerminato";
  const FINISHED_BUTTON_ID = "view-done-btn";
  const PROGRAM_BUTTON_ID = "view-todo-btn";
  const EXPORT_BUTTON_ID = "export-current-commessa-btn";
  const POTATURE_ID = "potature-abbattimenti";
  const COBO_ID = "sfalcio-cobo";
  const PENDING_STORAGE_KEY = "heraSpecialTerminatoPendingV2";
  const RETRY_DELAYS_MS = Object.freeze([0, 1000, 2000]);
  const text = (value) => String(value ?? "").trim();
  let specialViewMode = "program";
  let scheduled = false;
  let pendingSyncRunning = false;
  const processingPlants = new Set();

  function selectedId() {
    try {
      if (typeof selectedCommessaId !== "undefined" && text(selectedCommessaId)) return text(selectedCommessaId);
    } catch (_) {}
    const match = String(window.location.hash || "").match(/(?:^|[&#])commessa=([^&]+)/);
    if (!match?.[1]) return "";
    try { return decodeURIComponent(match[1]); } catch (_) { return match[1]; }
  }

  function isSpecialCommessa() {
    const id = text(selectedId()).toLowerCase();
    const potatureId = text(window.HeraPotatureFollowup?.commessaId || POTATURE_ID).toLowerCase();
    const coboId = text(window.HeraCoboMowing?.commessaId || COBO_ID).toLowerCase();
    return id === potatureId || id === coboId;
  }

  function currentPlants() {
    try {
      return Array.isArray(currentImpianti) ? currentImpianti : [];
    } catch (_) {
      return [];
    }
  }

  function plantKey(plant) {
    try {
      if (typeof buildImpiantoKey === "function") return text(buildImpiantoKey(plant));
    } catch (_) {}
    const idSap = text(plant?.idSap).toLowerCase();
    return idSap ? `sap:${idSap}` : text(plant?.id);
  }

  function plantForCard(card) {
    const key = text(card?.dataset?.impiantoKey);
    return currentPlants().find((plant) => plantKey(plant) === key) || null;
  }

  function plantDocumentIds(plant) {
    return [...new Set([
      text(plant?.id),
      ...(Array.isArray(plant?.sourceIds) ? plant.sourceIds.map(text) : [])
    ].filter(Boolean))];
  }

  function database() {
    try {
      if (typeof db !== "undefined" && db?.collection) return db;
    } catch (_) {}
    return window.db?.collection ? window.db : null;
  }

  function collectionName() {
    try {
      return typeof getCommesseCollectionName === "function" ? getCommesseCollectionName() : "commesse";
    } catch (_) {
      return "commesse";
    }
  }

  function serverTimestamp() {
    try {
      if (typeof firebase !== "undefined" && firebase.firestore?.FieldValue?.serverTimestamp) {
        return firebase.firestore.FieldValue.serverTimestamp();
      }
    } catch (_) {}
    return new Date();
  }

  function firestoreTimestamp(date) {
    try {
      if (typeof firebase !== "undefined" && firebase.firestore?.Timestamp?.fromDate) {
        return firebase.firestore.Timestamp.fromDate(date);
      }
    } catch (_) {}
    return date;
  }

  function isOffline() {
    try {
      return typeof navigator !== "undefined" && navigator.onLine === false;
    } catch (_) {
      return false;
    }
  }

  function wait(milliseconds) {
    return milliseconds > 0
      ? new Promise((resolve) => window.setTimeout(resolve, milliseconds))
      : Promise.resolve();
  }

  function localExecutionParts(date) {
    try {
      const parts = new Intl.DateTimeFormat("it-IT", {
        timeZone: "Europe/Rome",
        year: "numeric",
        month: "2-digit",
        day: "2-digit",
        hour: "2-digit",
        minute: "2-digit",
        hour12: false
      }).formatToParts(date).reduce((result, part) => {
        if (part.type !== "literal") result[part.type] = part.value;
        return result;
      }, {});
      return {
        date: `${parts.year}-${parts.month}-${parts.day}`,
        time: `${parts.hour === "24" ? "00" : parts.hour}:${parts.minute}`
      };
    } catch (_) {
      return {
        date: date.toISOString().slice(0, 10),
        time: date.toTimeString().slice(0, 5)
      };
    }
  }

  function operatorIdentity() {
    let user = null;
    try {
      if (typeof auth !== "undefined" && auth?.currentUser) user = auth.currentUser;
      else if (typeof currentUser !== "undefined" && currentUser) user = currentUser;
    } catch (_) {}
    user ||= window.currentUser || null;
    let name = text(user?.displayName || user?.email || "Operatore");
    try {
      if (typeof getOperatorDisplayName === "function") name = text(getOperatorDisplayName()) || name;
    } catch (_) {}
    return { uid: text(user?.uid), name, email: text(user?.email) };
  }

  function isTerminated(plant) {
    return plant?.[TERMINATED_FIELD] === true;
  }

  function getDisplayState(plant) {
    const active = isSpecialCommessa();
    const terminated = active && isTerminated(plant);
    const pending = terminated && (plant?.specialTerminatoPending === true || Boolean(pendingActionForPlant(plant)));
    return {
      active,
      terminated,
      pending,
      state: terminated ? "Finito" : "In programma",
      completedAt: terminated ? formatTimestamp(plant?.specialTerminatoAt) : "-",
      operator: terminated ? text(plant?.specialTerminatoBy || plant?.specialOperatore || "Operatore") : "-",
      action: terminated ? "FINITO" : "TERMINATO"
    };
  }

  function formatTimestamp(value) {
    let date = value;
    if (value?.toDate) date = value.toDate();
    else if (value?.seconds) date = new Date(Number(value.seconds) * 1000);
    if (!(date instanceof Date) || Number.isNaN(date.getTime())) return "Data non disponibile";
    return date.toLocaleString("it-IT", {
      day: "2-digit", month: "2-digit", year: "numeric", hour: "2-digit", minute: "2-digit", hour12: false
    });
  }

  function finishedStatusLabel(value) {
    let date = value;
    if (value?.toDate) date = value.toDate();
    else if (value?.seconds) date = new Date(Number(value.seconds) * 1000);
    if (!(date instanceof Date) || Number.isNaN(date.getTime())) return "✅ TERMINATO";
    const day = String(date.getDate()).padStart(2, "0");
    const month = new Intl.DateTimeFormat("it-IT", { month: "long" }).format(date).toUpperCase();
    return `✅ TERMINATO DAL ${day} ${month}`;
  }

  function finishedStatusDateLabel(value) {
    let date = value;
    if (value?.toDate) date = value.toDate();
    else if (value?.seconds) date = new Date(Number(value.seconds) * 1000);
    if (!(date instanceof Date) || Number.isNaN(date.getTime())) return "DATA NON DISPONIBILE";
    const day = String(date.getDate()).padStart(2, "0");
    const month = new Intl.DateTimeFormat("it-IT", { month: "short" })
      .format(date)
      .replace(".", "")
      .toUpperCase();
    return `${day} ${month} ${date.getFullYear()}`;
  }

  function esc(value) {
    return text(value).replace(/[&<>'"]/g, (character) => ({
      "&": "&amp;", "<": "&lt;", ">": "&gt;", "'": "&#39;", '"': "&quot;"
    }[character]));
  }

  function specialTypeLabel(plant) {
    const phase = text(plant?.potatureFase).toLowerCase();
    if (phase === "raccolta") return `Raccolta · ${text(plant.potatureMetodoLabel) || "Metodo non indicato"}`;
    if (phase === "ceppi") return `Ceppo · ${text(plant.potatureMetodoLabel) || "Metodo non indicato"}`;
    if (selectedId().toLowerCase() === COBO_ID) return "Verde COBO";
    return text(plant?.tipologiaIntervento || "Potatura / Abbattimento");
  }

  function selectedName() {
    try {
      if (typeof selectedCommessaName !== "undefined" && text(selectedCommessaName)) return text(selectedCommessaName);
      if (typeof commesseById !== "undefined") return text(commesseById.get(selectedId())?.nome) || "Commessa";
    } catch (_) {}
    return "Commessa";
  }

  function authenticatedUser() {
    try {
      if (typeof auth !== "undefined" && auth?.currentUser) return auth.currentUser;
    } catch (_) {}
    return null;
  }

  function validateCoordinates(plant) {
    const rawLatitude = plant?.gpsY;
    const rawLongitude = plant?.gpsX;
    const latitude = Number(rawLatitude);
    const longitude = Number(rawLongitude);
    const valid = rawLatitude !== null
      && rawLatitude !== undefined
      && text(rawLatitude) !== ""
      && rawLongitude !== null
      && rawLongitude !== undefined
      && text(rawLongitude) !== ""
      && Number.isFinite(latitude)
      && Number.isFinite(longitude)
      && latitude >= -90
      && latitude <= 90
      && longitude >= -180
      && longitude <= 180;
    return { valid, latitude, longitude };
  }

  function storage() {
    try {
      return window.localStorage || globalThis.localStorage || null;
    } catch (_) {
      return null;
    }
  }

  function loadPendingActions() {
    try {
      const parsed = JSON.parse(storage()?.getItem(PENDING_STORAGE_KEY) || "[]");
      return Array.isArray(parsed)
        ? parsed.filter((action) => action?.id && action?.commessaId && Array.isArray(action.documentIds) && action.documentIds.length)
        : [];
    } catch (error) {
      console.warn("Coda TERMINATO speciale non leggibile:", error);
      return [];
    }
  }

  function savePendingActions(actions) {
    try {
      storage()?.setItem(PENDING_STORAGE_KEY, JSON.stringify(Array.isArray(actions) ? actions : []));
    } catch (error) {
      console.warn("Coda TERMINATO speciale non salvata:", error);
    }
  }

  function actionId(commessaId, plant) {
    const operator = operatorIdentity();
    return `${operator.uid || "user"}:${commessaId}:${plantKey(plant) || plantDocumentIds(plant)[0] || "cantiere"}`;
  }

  function buildPendingAction(plant, commessaId, completedAt, operator) {
    const execution = localExecutionParts(completedAt);
    return {
      id: actionId(commessaId, plant),
      version: 2,
      commessaId,
      commessaName: selectedName(),
      documentIds: plantDocumentIds(plant),
      plantKey: plantKey(plant),
      plantName: text(plant?.denominazione || plant?.nome || "Cantiere"),
      plantIdSap: text(plant?.idSap),
      completedAt: completedAt.toISOString(),
      completionDate: execution.date,
      completionTime: execution.time,
      operatorUid: operator.uid,
      operatorName: operator.name,
      operatorEmail: operator.email,
      createdAt: new Date().toISOString(),
      attempts: 0,
      lastError: ""
    };
  }

  function upsertPendingAction(action) {
    const actions = loadPendingActions();
    const index = actions.findIndex((item) => item.id === action.id);
    if (index >= 0) actions.splice(index, 1, { ...actions[index], ...action, updatedAt: new Date().toISOString() });
    else actions.push(action);
    savePendingActions(actions);
    return action;
  }

  function updatePendingAction(actionIdValue, patch) {
    const actions = loadPendingActions();
    const index = actions.findIndex((item) => item.id === actionIdValue);
    if (index < 0) return null;
    actions[index] = { ...actions[index], ...patch, updatedAt: new Date().toISOString() };
    savePendingActions(actions);
    return actions[index];
  }

  function removePendingAction(actionIdValue) {
    const actions = loadPendingActions();
    const next = actions.filter((item) => item.id !== actionIdValue);
    if (next.length !== actions.length) savePendingActions(next);
  }

  function pendingActionForPlant(plant, commessaId = selectedId()) {
    const ids = new Set(plantDocumentIds(plant));
    const key = plantKey(plant);
    return loadPendingActions().find((action) => action.commessaId === commessaId && (
      (key && action.plantKey === key)
      || action.documentIds.some((id) => ids.has(id))
    )) || null;
  }

  function localPatchFromAction(action, pending = true) {
    return {
      [TERMINATED_FIELD]: true,
      specialStato: "FINITO",
      specialTerminatoAt: new Date(action.completedAt),
      specialTerminatoBy: action.operatorName,
      specialTerminatoByUid: action.operatorUid,
      specialTerminatoByEmail: action.operatorEmail,
      specialDataEsecuzione: action.completionDate,
      specialOraEsecuzione: action.completionTime,
      specialOperatore: action.operatorName,
      specialTerminatoVersione: 2,
      specialTerminatoPending: pending,
      specialTerminatoLastError: pending ? text(action.lastError) : ""
    };
  }

  function firestorePatchFromAction(action) {
    const completedAt = new Date(action.completedAt);
    return {
      [TERMINATED_FIELD]: true,
      specialStato: "FINITO",
      specialTerminatoAt: firestoreTimestamp(completedAt),
      specialTerminatoBy: action.operatorName,
      specialTerminatoByUid: action.operatorUid,
      specialTerminatoByEmail: action.operatorEmail,
      specialDataEsecuzione: action.completionDate,
      specialOraEsecuzione: action.completionTime,
      specialOperatore: action.operatorName,
      specialTerminatoVersione: 2,
      specialTerminatoPending: false,
      specialTerminatoLastError: "",
      updatedAt: serverTimestamp()
    };
  }

  function applyPendingStateForSelectedCommessa() {
    const commessaId = selectedId();
    if (!commessaId || !isSpecialCommessa()) return;
    loadPendingActions().filter((action) => action.commessaId === commessaId).forEach((action) => {
      const actionIds = new Set(action.documentIds);
      currentPlants().forEach((plant) => {
        if ((action.plantKey && plantKey(plant) === action.plantKey)
          || plantDocumentIds(plant).some((id) => actionIds.has(id))) {
          Object.assign(plant, localPatchFromAction(action, true));
        }
      });
    });
  }

  async function persistActionOnce(action) {
    const store = database();
    if (!store?.collection) throw new Error("Database non disponibile. Controlla la connessione e riprova.");
    const documentIds = [...new Set(action.documentIds.map(text).filter(Boolean))];
    if (!documentIds.length) throw new Error("Nessun cantiere disponibile per il salvataggio.");
    if (documentIds.length > 500) throw new Error("Troppi documenti collegati per un singolo salvataggio.");
    const reference = store.collection(collectionName()).doc(action.commessaId).collection("impianti");
    const patch = firestorePatchFromAction(action);
    const batch = store.batch?.();
    if (batch) {
      documentIds.forEach((id) => batch.set(reference.doc(id), patch, { merge: true }));
      await batch.commit();
    } else {
      await Promise.all(documentIds.map((id) => reference.doc(id).set(patch, { merge: true })));
    }
  }

  async function persistActionWithRetry(action) {
    let lastError = null;
    for (let attempt = 0; attempt < RETRY_DELAYS_MS.length; attempt += 1) {
      await wait(RETRY_DELAYS_MS[attempt]);
      try {
        await persistActionOnce(action);
        return true;
      } catch (error) {
        lastError = error;
        updatePendingAction(action.id, {
          attempts: Number(action.attempts || 0) + attempt + 1,
          lastError: text(error?.message || error).slice(0, 500)
        });
      }
    }
    throw lastError || new Error("Salvataggio TERMINATO non riuscito.");
  }

  function refreshSpecialViews() {
    try { if (typeof renderMap === "function") renderMap(); } catch (_) {}
    if (specialViewMode === "finished") renderFinishedList();
    else scheduleApply();
  }

  function publishCompletionEffects(action) {
    try {
      if (typeof logActivity === "function") {
        void Promise.resolve(logActivity("pressione_terminato", "Pressione TERMINATO", {
          buttonLabel: "TERMINATO",
          commessaId: action.commessaId,
          commessaName: action.commessaName,
          impiantoId: action.plantKey || action.documentIds[0] || "",
          impiantoName: action.plantName,
          detail: `Cantiere spostato nei FINITI il ${action.completionDate} alle ${action.completionTime}`
        })).catch((error) => console.warn("Registro attività TERMINATO non salvato:", error));
      }
    } catch (error) {
      console.warn("Registro attività TERMINATO non disponibile:", error);
    }
    try {
      if (typeof publishGlobalNotificationEvent === "function") {
        void Promise.resolve(publishGlobalNotificationEvent("impianto-done", {
          title: "Cantiere finito",
          body: `${action.operatorName || "Operatore"} ha premuto TERMINATO su ${action.plantName || "Cantiere"} (${action.commessaName || "Commessa"}).`,
          commessaId: action.commessaId,
          commessaName: action.commessaName,
          impiantoName: action.plantName,
          impiantoKey: action.plantKey
        })).catch((error) => console.warn("Notifica TERMINATO non salvata:", error));
      }
    } catch (error) {
      console.warn("Notifica TERMINATO non disponibile:", error);
    }
  }

  function markActionSyncedLocally(action) {
    if (selectedId() !== action.commessaId) return;
    const actionIds = new Set(action.documentIds);
    currentPlants().forEach((plant) => {
      if ((action.plantKey && plantKey(plant) === action.plantKey)
        || plantDocumentIds(plant).some((id) => actionIds.has(id))) {
        Object.assign(plant, localPatchFromAction(action, false));
      }
    });
    refreshSpecialViews();
  }

  async function syncPendingActions() {
    if (pendingSyncRunning || isOffline()) return;
    const store = database();
    const user = authenticatedUser();
    if (!store?.collection || !user) return;
    pendingSyncRunning = true;
    try {
      const userUid = text(user.uid);
      const actions = loadPendingActions().filter((action) => !action.operatorUid || action.operatorUid === userUid);
      for (const action of actions) {
        if (processingPlants.has(action.id)) continue;
        processingPlants.add(action.id);
        try {
          await persistActionWithRetry(action);
          removePendingAction(action.id);
          markActionSyncedLocally(action);
          publishCompletionEffects(action);
        } catch (error) {
          const message = text(error?.message || error).slice(0, 500);
          updatePendingAction(action.id, { lastError: message });
          if (selectedId() === action.commessaId) {
            const actionIds = new Set(action.documentIds);
            currentPlants().forEach((plant) => {
              if (plantDocumentIds(plant).some((id) => actionIds.has(id))) {
                Object.assign(plant, localPatchFromAction({ ...action, lastError: message }, true));
              }
            });
            refreshSpecialViews();
          }
        } finally {
          processingPlants.delete(action.id);
        }
      }
    } finally {
      pendingSyncRunning = false;
    }
  }

  function buildFinishedRowsForExport(commessaName) {
    const email = text(authenticatedUser()?.email);
    return currentPlants().filter(isTerminated)
      .flatMap((plant) => buildRowsForEachCodicePrezzo(plant))
      .map((plant) => {
        const finishedInfo = formatDoneDateTime(plant.specialTerminatoAt);
        return {
          "Commessa padre": "",
          Commessa: commessaName,
          Cantiere: plant.cantiereRiga || "",
          Distretto: plant.distretto || "",
          "ID SAP": plant.idSap || "",
          Denominazione: plant.denominazione || "",
          Comune: plant.comune || "",
          Indirizzo: plant.indirizzo || "",
          "Voce riferimento": plant.voceRiferimento || "",
          "Codice prezzo": plant.codicePrezzoSingolo || plant.codicePrezzo || "",
          Sfalci: plant.sfalci || "",
          "Frequenza annua": plant.frequenzaAnnua || "",
          "Tipologia intervento": plant.tipologiaIntervento || specialTypeLabel(plant),
          "Lavorazioni richieste": plant.lavorazioniRichieste || "",
          "GPS Y": plant.gpsY ?? "",
          "GPS X": plant.gpsX ?? "",
          "Tipo manutenzione": plant.tipoManutenzione || classifyTipoManutenzione(plant.codicePrezzo),
          Stato: "Finito",
          "Data esecuzione": finishedInfo.date,
          "Ora esecuzione": finishedInfo.time,
          "Eseguito da": plant.specialTerminatoBy || "-",
          "Email operatore": email
        };
      });
  }

  function exportFinishedSummary() {
    if (!authenticatedUser()) {
      window.alert("Devi fare login per esportare il riepilogo.");
      return;
    }

    try {
      const commessaName = selectedName();
      const rows = buildFinishedRowsForExport(commessaName);
      if (!rows.length) {
        window.alert(`Nessun impianto FINITO da esportare per la commessa "${commessaName}".`);
        return;
      }

      const worksheet = XLSX.utils.json_to_sheet(rows);
      const workbook = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(workbook, worksheet, "Riepilogo impianti");
      const safeName = String(commessaName || "commessa")
        .trim()
        .replace(/[\\/:*?"<>|]/g, "_")
        .replace(/\s+/g, "_");
      const timestamp = new Date().toISOString().slice(0, 19).replace(/[:T]/g, "-");
      XLSX.writeFile(workbook, `riepilogo_impianti_${safeName}_${timestamp}.xlsx`);
    } catch (error) {
      console.error(error);
      window.alert("Errore durante l'esportazione del riepilogo in Excel.");
    }
  }

  function ensureStyle() {
    if (document.getElementById(STYLE_ID)) return;
    const style = document.createElement("style");
    style.id = STYLE_ID;
    style.textContent = `
      .special-core-action-hidden { display: none !important; }
      .special-terminato-program-hidden { display: none !important; }
      .impianto-main-column > .item-actions.impianto-actions .impianto-primary-actions .special-terminato-btn { order: 2; grid-column: 2; grid-row: 1; display: inline-flex; align-items: center; justify-content: center; box-sizing: border-box; width: 100%; max-width: none; min-width: 0; min-height: 38px !important; margin: 0; padding: 6px 8px; overflow: hidden; border: 1px solid #16a34a; border-radius: 12px; color: #fff; background: #16a34a; font-size: clamp(.7rem, 3.1vw, .9rem); font-weight: 900; line-height: 1; letter-spacing: .015em; white-space: nowrap; text-overflow: clip; box-shadow: 0 4px 12px rgba(8,120,63,.24); }
      .special-terminato-btn:disabled { opacity: .66; cursor: wait; }
      .special-terminated-card { border: 2px solid #8bd3ae; background: #f2fbf6; }
      .special-terminated-card .impianto-summary-title-wrap strong { color: #075f34; }
      .special-terminated-card .impianto-details p { margin: 4px 0; }
      .special-terminated-badge { display: inline-flex; width: fit-content; padding: 5px 9px; border-radius: 999px; color: #075f34; background: #d7f4e4; font-size: .78rem; font-weight: 900; }
      .impianto-main-column > .item-actions.impianto-actions .impianto-primary-actions .special-finished-status-btn { order: 2; grid-column: 2; grid-row: 1; display: inline-flex; flex-direction: column; align-items: center; justify-content: center; box-sizing: border-box; width: 100%; max-width: none; min-width: 0; min-height: 38px !important; margin: 0; padding: 4px 6px; overflow: hidden; border: 1px solid #f59e0b; border-radius: 12px; color: #78350f; background: #fbbf24; font-size: .68rem; font-weight: 900; line-height: 1.05; white-space: normal; opacity: 1; }
      .special-finished-status-btn span, .special-finished-status-btn small { display: block; max-width: 100%; overflow: hidden; white-space: nowrap; text-overflow: clip; }
      .special-finished-status-btn small { margin-top: 2px; font-size: .55rem; letter-spacing: .02em; }
    `;
    document.head.appendChild(style);
  }

  function updateLocalPlant(plant, patch) {
    Object.assign(plant, patch);
    const ids = new Set(plantDocumentIds(plant));
    currentPlants().forEach((item) => {
      if (item === plant || plantDocumentIds(item).some((id) => ids.has(id))) Object.assign(item, patch);
    });
  }

  function removeProgramCardImmediately(plant) {
    const list = document.getElementById("impianti-lista") || document.querySelector(".impianti-lista");
    if (!list) return;
    const key = plantKey(plant);
    list.querySelectorAll(".impianto-item[data-impianto-key]").forEach((card) => {
      if (text(card.dataset.impiantoKey) !== key) return;
      card.hidden = true;
      card.remove();
    });
  }

  async function terminatePlant(plant, button) {
    if (!plant || isTerminated(plant)) return;
    if (!isSpecialCommessa()) throw new Error("TERMINATO è disponibile solo nelle due commesse speciali.");
    const user = authenticatedUser();
    if (!user) throw new Error("Sessione scaduta: esegui nuovamente il login.");
    const commessaId = text(selectedId());
    const documentIds = plantDocumentIds(plant);
    if (!database()?.collection || !commessaId || !documentIds.length) {
      throw new Error("Cantiere non disponibile per il salvataggio.");
    }
    if (!validateCoordinates(plant).valid) {
      throw new Error("La posizione nella scheda del cantiere è mancante o non valida. Correggila prima di premere TERMINATO.");
    }

    const operator = operatorIdentity();
    const processingKey = actionId(commessaId, plant);
    if (processingPlants.has(processingKey)) return;
    processingPlants.add(processingKey);
    const completedAt = new Date();
    const action = buildPendingAction(plant, commessaId, completedAt, operator);

    button.disabled = true;
    button.textContent = "SALVATAGGIO…";
    upsertPendingAction(action);
    updateLocalPlant(plant, localPatchFromAction(action, true));
    removeProgramCardImmediately(plant);
    showFinishedList();
    try { if (typeof renderMap === "function") renderMap(); } catch (_) {}

    if (isOffline()) {
      processingPlants.delete(processingKey);
      window.alert("Sei offline: il cantiere è già nei FINITI. Il salvataggio si sincronizzerà automaticamente quando torna Internet.");
      return true;
    }

    try {
      await persistActionWithRetry(action);
      removePendingAction(action.id);
      updateLocalPlant(plant, localPatchFromAction(action, false));
      publishCompletionEffects(action);
      refreshSpecialViews();
      return true;
    } catch (error) {
      const message = text(error?.message || error).slice(0, 500);
      const queuedAction = updatePendingAction(action.id, { lastError: message }) || { ...action, lastError: message };
      updateLocalPlant(plant, localPatchFromAction(queuedAction, true));
      refreshSpecialViews();
      window.alert("Il cantiere è nei FINITI, ma il salvataggio online non è stato confermato. Riproverò automaticamente quando torna la connessione.");
      return true;
    } finally {
      processingPlants.delete(processingKey);
    }
  }

  async function terminateFromMap(plant, button) {
    if (!plant || isTerminated(plant)) return false;
    const targetButton = button || { disabled: false, textContent: "TERMINATO" };
    return terminatePlant(plant, targetButton);
  }

  function createTerminateButton(card, plant) {
    const primary = card.querySelector(".impianto-primary-actions");
    if (!primary || primary.querySelector(".special-terminato-btn")) return;
    const button = document.createElement("button");
    button.type = "button";
    button.className = "btn special-terminato-btn";
    button.textContent = "TERMINATO";
    button.setAttribute("aria-label", "Sposta il cantiere nei Finiti della commessa speciale");
    button.addEventListener("click", async (event) => {
      event.preventDefault();
      event.stopPropagation();
      try {
        await terminatePlant(plant, button);
      } catch (error) {
        button.disabled = false;
        button.textContent = "TERMINATO";
        window.alert(error?.message || "Impossibile terminare il cantiere.");
      }
    });
    primary.appendChild(button);
  }

  function hideLegacySpecialActions(card) {
    card.querySelectorAll([
      '.impianto-primary-actions [data-action-key="whatsapp"]',
      '.impianto-primary-actions [data-action-key="whatsapp-attachment"]',
      ".impianto-force-done-btn",
      '[data-action-key="reset"]'
    ].join(",")).forEach((element) => element.classList.add("special-core-action-hidden"));
    card.querySelectorAll("button").forEach((button) => {
      const label = text(button.textContent || button.getAttribute("aria-label")).toUpperCase();
      if (label.includes("FORZA IN FATTI") || label.includes("FORZA CHIUSURA IMPIANTO COME FATTO")) {
        button.classList.add("special-core-action-hidden");
      }
    });
  }

  function restoreLegacyActions() {
    document.querySelectorAll(".special-core-action-hidden").forEach((element) => element.classList.remove("special-core-action-hidden"));
    document.querySelectorAll(".special-terminato-btn").forEach((button) => button.remove());
  }

  function saveOriginalLabel(button) {
    if (!button || button.dataset.specialOriginalLabel !== undefined) return;
    button.dataset.specialOriginalLabel = button.textContent || "";
  }

  function restoreSpecialTabs() {
    [FINISHED_BUTTON_ID, PROGRAM_BUTTON_ID, EXPORT_BUTTON_ID].forEach((id) => {
      const button = document.getElementById(id);
      if (!button || button.dataset.specialOriginalLabel === undefined) return;
      button.textContent = button.dataset.specialOriginalLabel;
      delete button.dataset.specialOriginalLabel;
    });
    document.getElementById(EXPORT_BUTTON_ID)?.classList.remove("special-core-action-hidden");
  }

  function selectSpecialTab(mode) {
    const finishedButton = document.getElementById(FINISHED_BUTTON_ID);
    const programButton = document.getElementById(PROGRAM_BUTTON_ID);
    finishedButton?.classList.toggle("btn-primary", mode === "finished");
    programButton?.classList.toggle("btn-primary", mode === "program");
  }

  function showFinishedList() {
    specialViewMode = "finished";
    selectSpecialTab("finished");
    renderFinishedList();
  }

  function configureSpecialTabs() {
    const toolbar = document.querySelector("#impianti-card .view-tabs") || document.querySelector(".view-tabs");
    if (!toolbar) return null;
    const finishedButton = document.getElementById(FINISHED_BUTTON_ID);
    const programButton = document.getElementById(PROGRAM_BUTTON_ID);
    const exportButton = document.getElementById(EXPORT_BUTTON_ID);
    saveOriginalLabel(finishedButton);
    saveOriginalLabel(programButton);
    saveOriginalLabel(exportButton);
    if (finishedButton) finishedButton.textContent = "✅ Finiti";
    if (programButton) programButton.textContent = "🛠️ In programma";
    if (exportButton) {
      exportButton.textContent = "📤 Esporta finiti";
      exportButton.classList.remove("special-core-action-hidden");
    }

    if (finishedButton && !finishedButton.dataset.specialFinishedHandler) {
      finishedButton.addEventListener("click", (event) => {
        if (!isSpecialCommessa()) return;
        event.preventDefault();
        event.stopImmediatePropagation();
        showFinishedList();
      }, true);
      finishedButton.dataset.specialFinishedHandler = "1";
    }
    if (programButton && !programButton.dataset.specialProgramHandler) {
      programButton.addEventListener("click", () => {
        if (!isSpecialCommessa()) return;
        specialViewMode = "program";
        selectSpecialTab("program");
        window.setTimeout(apply, 0);
      }, true);
      programButton.dataset.specialProgramHandler = "1";
    }
    if (exportButton && !exportButton.dataset.specialFinishedExportHandler) {
      exportButton.addEventListener("click", (event) => {
        if (!isSpecialCommessa()) return;
        event.preventDefault();
        event.stopImmediatePropagation();
        exportFinishedSummary();
      }, true);
      exportButton.dataset.specialFinishedExportHandler = "1";
    }
    if (!toolbar.dataset.specialTerminatoTabs) {
      toolbar.addEventListener("click", (event) => {
        const button = event.target.closest("button");
        if (!button || !isSpecialCommessa() || button.id === FINISHED_BUTTON_ID || button.id === PROGRAM_BUTTON_ID) return;
        specialViewMode = "program";
        window.setTimeout(apply, 0);
      });
      toolbar.dataset.specialTerminatoTabs = "1";
    }
    return { finishedButton, programButton };
  }

  function canShowFinishedManagementAction(actionKey) {
    try {
      if (typeof isImpiantoActionDenied === "function" && isImpiantoActionDenied(actionKey)) return false;
      if (["edit", "delete"].includes(actionKey) && typeof canUseImpiantoAction === "function") {
        return canUseImpiantoAction(actionKey);
      }
      return true;
    } catch (_) {
      return false;
    }
  }

  function appendFinishedManagementAction(container, label, actionKey, handler) {
    if (!container || typeof handler !== "function" || !canShowFinishedManagementAction(actionKey)) return;
    const button = document.createElement("button");
    button.type = "button";
    button.className = "btn";
    button.dataset.actionKey = actionKey;
    button.textContent = label;
    button.addEventListener("click", async (event) => {
      event.preventDefault();
      event.stopPropagation();
      if (button.disabled) return;
      button.disabled = true;
      try {
        await handler();
      } catch (error) {
        window.alert(error?.message || "Azione non disponibile per questo cantiere.");
      } finally {
        button.disabled = false;
      }
    });
    container.appendChild(button);
  }

  function createFinishedManagementMenu(plant) {
    const managementStack = document.createElement("div");
    managementStack.className = "impianto-management-stack";
    const managementActions = document.createElement("div");
    managementActions.className = "item-actions item-actions-gestione special-finished-management-actions hidden";

    appendFinishedManagementAction(
      managementActions,
      "Segnala problema",
      "problem-report",
      typeof openImpiantoReportModal === "function" ? () => openImpiantoReportModal(plant) : null
    );
    appendFinishedManagementAction(
      managementActions,
      "Aggiorna GPS",
      "gps-update",
      typeof requestGpsUpdate === "function" ? () => requestGpsUpdate(plant) : null
    );
    appendFinishedManagementAction(
      managementActions,
      "Modifica",
      "edit",
      typeof openImpiantoEditor === "function" ? () => openImpiantoEditor(plant) : null
    );
    appendFinishedManagementAction(
      managementActions,
      "Elimina",
      "delete",
      typeof deleteImpianto === "function" ? () => deleteImpianto(plant) : null
    );

    let adminCanUpload = false;
    try { adminCanUpload = typeof canManageData === "function" && canManageData(); } catch (_) {}
    if (adminCanUpload) {
      appendFinishedManagementAction(
        managementActions,
        "Inserisci PDF richiesta",
        "request-pdf",
        typeof setImpiantoRequestDriveLink === "function" ? () => setImpiantoRequestDriveLink(plant) : null
      );
    }

    if (!managementActions.childElementCount) return null;
    const manageButton = document.createElement("button");
    manageButton.type = "button";
    manageButton.className = "btn gestione-toggle-btn";
    manageButton.textContent = "⚙️";
    manageButton.title = "Gestione";
    manageButton.setAttribute("aria-label", "Gestione cantiere finito");
    manageButton.setAttribute("aria-expanded", "false");
    manageButton.addEventListener("click", (event) => {
      event.preventDefault();
      event.stopPropagation();
      const expanded = manageButton.getAttribute("aria-expanded") === "true";
      manageButton.setAttribute("aria-expanded", expanded ? "false" : "true");
      managementActions.classList.toggle("hidden", expanded);
    });
    managementStack.appendChild(manageButton);
    return { managementStack, managementActions };
  }

  function renderFinishedList() {
    if (!isSpecialCommessa() || specialViewMode !== "finished") return;
    const list = document.getElementById("impianti-lista") || document.querySelector(".impianti-lista");
    if (!list) return;
    const plants = currentPlants()
      .filter(isTerminated)
      .filter((plant) => {
        try { return typeof matchesImpiantoSearch === "function" ? matchesImpiantoSearch(plant) : true; } catch (_) { return true; }
      })
      .sort((first, second) => {
        try {
          if (typeof distanceFromUser === "function") return distanceFromUser(first) - distanceFromUser(second);
        } catch (_) {}
        return text(first?.denominazione).localeCompare(text(second?.denominazione), "it");
      });
    const signature = plants.map((plant) => `${plantKey(plant)}:${formatTimestamp(plant.specialTerminatoAt)}:${text(plant.specialTerminatoBy)}:${plant.specialTerminatoPending === true}:${text(plant.specialTerminatoLastError)}`).join("|");
    if (list.dataset.specialTerminatoSignature === signature && list.querySelector("[data-special-terminated-render]")) return;
    list.dataset.specialTerminatoSignature = signature;
    list.innerHTML = "";
    const wrapper = document.createElement("section");
    wrapper.dataset.specialTerminatedRender = "1";
    wrapper.className = "impianti-lista";
    if (!plants.length) {
      wrapper.innerHTML = '<p class="muted">Nessun cantiere finito in questa commessa speciale.</p>';
    } else {
      plants.forEach((plant) => {
        const card = document.createElement("article");
        card.className = "impianto-item card-impianto done special-terminated-card";
        card.dataset.impiantoKey = plantKey(plant);
        const displayState = getDisplayState(plant);

        const mainColumn = document.createElement("div");
        mainColumn.className = "impianto-main-column impianto-left";
        const summary = document.createElement("button");
        summary.type = "button";
        summary.className = "impianto-summary-btn";
        summary.setAttribute("aria-expanded", "false");
        summary.innerHTML = `
          <span class="impianto-summary-topline">
            <span class="impianto-summary-title-wrap"><strong>${esc(plant.denominazione || plant.nome || "Cantiere")}</strong></span>
          </span>
          <small class="impianto-summary-meta">
            <span class="special-terminated-badge">✅ Nei FINITI</span>
            ${displayState.pending ? '<span class="badge badge-whatsapp-pending">⏳ Sincronizzazione in attesa</span>' : ""}
            <span>${esc(specialTypeLabel(plant))}</span>
          </small>`;

        const details = document.createElement("div");
        details.className = "impianto-details";
        details.hidden = true;
        details.innerHTML = `
          <p><b>Tipo:</b> ${esc(specialTypeLabel(plant))}</p>
          <p><b>Comune:</b> ${esc(plant.comune || "-")}</p>
          <p><b>Indirizzo:</b> ${esc(plant.indirizzo || "-")}</p>
          <p><b>Codice prezzo:</b> ${esc(plant.codicePrezzo || plant.voceRiferimento || "-")}</p>
          <p><b>Lavorazioni richieste:</b> ${esc(plant.lavorazioniRichieste || plant.tipologiaIntervento || "-")}</p>
          <p><b>Stato:</b> Finito</p>
          <p><b>Data e ora terminato:</b> ${esc(formatTimestamp(plant.specialTerminatoAt))}</p>
          <p><b>Eseguito da:</b> ${esc(plant.specialTerminatoBy || "Operatore")}</p>
          ${displayState.pending ? `<p><b>Sincronizzazione:</b> In attesa${plant.specialTerminatoLastError ? ` · ${esc(plant.specialTerminatoLastError)}` : ""}</p>` : ""}`;
        summary.addEventListener("click", () => {
          const expanded = summary.getAttribute("aria-expanded") === "true";
          summary.setAttribute("aria-expanded", expanded ? "false" : "true");
          details.hidden = expanded;
          card.classList.toggle("is-expanded", !expanded);
        });

        const actions = document.createElement("div");
        actions.className = "item-actions impianto-actions";
        const primary = document.createElement("div");
        primary.className = "impianto-primary-actions";
        const navigateButton = document.createElement("button");
        navigateButton.type = "button";
        navigateButton.className = "btn action-icon-btn";
        navigateButton.dataset.actionKey = "navigate";
        navigateButton.textContent = "🗺️";
        navigateButton.title = "Naviga";
        navigateButton.setAttribute("aria-label", "Naviga");
        navigateButton.addEventListener("click", () => {
          try {
            if (typeof navigateToImpianto === "function") navigateToImpianto(plant);
          } catch (_) {
            window.alert("Navigazione non disponibile per questo cantiere.");
          }
        });
        const statusButton = document.createElement("button");
        statusButton.type = "button";
        statusButton.className = "btn special-finished-status-btn";
        statusButton.innerHTML = displayState.pending
          ? "<span>⏳ TERMINATO</span><small>DA SINCRONIZZARE</small>"
          : `<span>✅ TERMINATO</span><small>${esc(finishedStatusDateLabel(plant.specialTerminatoAt))}</small>`;
        statusButton.disabled = true;
        statusButton.setAttribute("aria-label", `${displayState.pending ? "Salvataggio TERMINATO in attesa" : finishedStatusLabel(plant.specialTerminatoAt)}. Nessun messaggio Whazzup viene aperto.`);
        primary.appendChild(navigateButton);
        primary.appendChild(statusButton);
        actions.appendChild(primary);

        const managementMenu = createFinishedManagementMenu(plant);
        if (managementMenu) {
          const secondary = document.createElement("div");
          secondary.className = "impianto-secondary-actions";
          secondary.appendChild(managementMenu.managementStack);
          actions.appendChild(secondary);
        }

        mainColumn.appendChild(summary);
        mainColumn.appendChild(details);
        mainColumn.appendChild(actions);
        if (managementMenu) mainColumn.appendChild(managementMenu.managementActions);
        card.appendChild(mainColumn);
        wrapper.appendChild(card);
      });
    }
    list.appendChild(wrapper);
  }

  function renderCurrentListIfOwned() {
    if (!isSpecialCommessa() || specialViewMode !== "finished") return false;
    renderFinishedList();
    return true;
  }

  function applyStandardList() {
    const list = document.getElementById("impianti-lista") || document.querySelector(".impianti-lista");
    if (!list) return;
    list.querySelectorAll(".impianto-item[data-impianto-key]").forEach((card) => {
      const plant = plantForCard(card);
      if (!plant) return;
      hideLegacySpecialActions(card);
      card.querySelectorAll(".impianto-details p").forEach((row) => {
        const label = text(row.querySelector("b")?.textContent).toLowerCase();
        if (label === "stato:") row.innerHTML = "<b>Stato:</b> In programma";
        if (label === "eseguito da:") row.innerHTML = "<b>Eseguito da:</b> -";
      });
      if (isTerminated(plant)) {
        card.classList.add("special-terminato-program-hidden");
        card.hidden = true;
        card.remove();
        return;
      }
      card.classList.remove("special-terminato-program-hidden");
      card.hidden = false;
      createTerminateButton(card, plant);
    });
  }

  function apply() {
    ensureStyle();
    if (!isSpecialCommessa()) {
      specialViewMode = "program";
      restoreSpecialTabs();
      restoreLegacyActions();
      return;
    }
    applyPendingStateForSelectedCommessa();
    configureSpecialTabs();
    if (specialViewMode === "finished") renderFinishedList();
    else applyStandardList();
  }

  function scheduleApply() {
    if (scheduled) return;
    scheduled = true;
    window.requestAnimationFrame(() => {
      scheduled = false;
      apply();
    });
  }

  function install() {
    ensureStyle();
    const list = document.getElementById("impianti-lista") || document.querySelector(".impianti-lista");
    if (list && !list.dataset.specialTerminatoObserver) {
      const observer = new MutationObserver(scheduleApply);
      observer.observe(list, { childList: true, subtree: true });
      list.dataset.specialTerminatoObserver = "1";
    }
    apply();
    void syncPendingActions();
  }

  window.addEventListener("hashchange", () => {
    specialViewMode = "program";
    scheduleApply();
  });
  window.addEventListener("hera:data-ready", install);
  window.addEventListener("online", () => { void syncPendingActions(); });
  if (document.readyState === "loading") document.addEventListener("DOMContentLoaded", install, { once: true });
  else install();

  window.HeraSpecialTerminato = Object.freeze({
    installed: true,
    field: TERMINATED_FIELD,
    isSpecialCommessa,
    isTerminated,
    getDisplayState,
    terminatePlant,
    terminateFromMap,
    syncPendingActions,
    loadPendingActions,
    exportFinishedSummary,
    apply,
    renderFinishedList,
    renderCurrentListIfOwned
  });
})();
