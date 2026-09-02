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
  const text = (value) => String(value ?? "").trim();
  let specialViewMode = "program";
  let scheduled = false;

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
    if (Array.isArray(plant?.sourceIds) && plant.sourceIds.length) {
      return [...new Set(plant.sourceIds.map(text).filter(Boolean))];
    }
    return text(plant?.id) ? [text(plant.id)] : [];
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
    return { uid: text(user?.uid), name };
  }

  function isTerminated(plant) {
    return plant?.[TERMINATED_FIELD] === true;
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
      .special-terminato-btn { min-width: 126px; min-height: 46px; border: 0; border-radius: 13px; color: #fff; background: #08783f; font-weight: 900; letter-spacing: .035em; box-shadow: 0 4px 12px rgba(8,120,63,.24); }
      .special-terminato-btn:disabled { opacity: .66; cursor: wait; }
      .special-terminated-card { border: 2px solid #8bd3ae; background: #f2fbf6; }
      .special-terminated-card header { display: grid; gap: 5px; }
      .special-terminated-card header strong { color: #075f34; font-size: 1.02rem; }
      .special-terminated-card p { margin: 4px 0; }
      .special-terminated-badge { display: inline-flex; width: fit-content; padding: 5px 9px; border-radius: 999px; color: #075f34; background: #d7f4e4; font-size: .78rem; font-weight: 900; }
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

  async function terminatePlant(plant, button) {
    if (!plant || isTerminated(plant)) return;
    if (!isSpecialCommessa()) throw new Error("TERMINATO è disponibile solo nelle due commesse speciali.");
    const store = database();
    const commessaId = text(selectedId());
    const documentIds = plantDocumentIds(plant);
    if (!store?.collection || !commessaId || !documentIds.length) {
      throw new Error("Cantiere non disponibile per il salvataggio.");
    }

    const operator = operatorIdentity();
    const timestamp = serverTimestamp();
    const localTimestamp = new Date();
    const patch = {
      [TERMINATED_FIELD]: true,
      specialTerminatoAt: timestamp,
      specialTerminatoBy: operator.name,
      specialTerminatoByUid: operator.uid,
      specialTerminatoVersione: 1,
      updatedAt: timestamp
    };

    button.disabled = true;
    button.textContent = "SALVATAGGIO…";
    const reference = store.collection(collectionName()).doc(commessaId).collection("impianti");
    const batch = store.batch?.();
    if (batch) {
      documentIds.forEach((id) => batch.set(reference.doc(id), patch, { merge: true }));
      await batch.commit();
    } else {
      await Promise.all(documentIds.map((id) => reference.doc(id).set(patch, { merge: true })));
    }

    updateLocalPlant(plant, {
      [TERMINATED_FIELD]: true,
      specialTerminatoAt: localTimestamp,
      specialTerminatoBy: operator.name,
      specialTerminatoByUid: operator.uid,
      specialTerminatoVersione: 1
    });
    showFinishedList();
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
      ".impianto-force-done-btn"
    ].join(",")).forEach((element) => element.classList.add("special-core-action-hidden"));
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

  function renderFinishedList() {
    if (!isSpecialCommessa() || specialViewMode !== "finished") return;
    const list = document.getElementById("impianti-lista") || document.querySelector(".impianti-lista");
    if (!list) return;
    const plants = currentPlants().filter(isTerminated);
    const signature = plants.map((plant) => `${plantKey(plant)}:${formatTimestamp(plant.specialTerminatoAt)}:${text(plant.specialTerminatoBy)}`).join("|");
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
        card.className = "impianto-item card-impianto special-terminated-card";
        card.innerHTML = `
          <header>
            <span class="special-terminated-badge">✅ TERMINATO</span>
            <strong>${esc(plant.denominazione || plant.nome || "Cantiere")}</strong>
          </header>
          <p><b>Tipo:</b> ${esc(specialTypeLabel(plant))}</p>
          <p><b>Comune:</b> ${esc(plant.comune || "-")}</p>
          <p><b>Indirizzo:</b> ${esc(plant.indirizzo || "-")}</p>
          <p><b>Terminato da:</b> ${esc(plant.specialTerminatoBy || "Operatore")}</p>
          <p><b>Data e ora:</b> ${esc(formatTimestamp(plant.specialTerminatoAt))}</p>`;
        wrapper.appendChild(card);
      });
    }
    list.appendChild(wrapper);
  }

  function applyStandardList() {
    const list = document.getElementById("impianti-lista") || document.querySelector(".impianti-lista");
    if (!list) return;
    list.querySelectorAll(".impianto-item[data-impianto-key]").forEach((card) => {
      const plant = plantForCard(card);
      if (!plant) return;
      hideLegacySpecialActions(card);
      if (isTerminated(plant)) {
        card.hidden = true;
        return;
      }
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
  }

  window.addEventListener("hashchange", () => {
    specialViewMode = "program";
    scheduleApply();
  });
  window.addEventListener("hera:data-ready", install);
  if (document.readyState === "loading") document.addEventListener("DOMContentLoaded", install, { once: true });
  else install();

  window.HeraSpecialTerminato = Object.freeze({
    installed: true,
    field: TERMINATED_FIELD,
    isSpecialCommessa,
    isTerminated,
    terminatePlant,
    exportFinishedSummary,
    apply,
    renderFinishedList
  });
})();
