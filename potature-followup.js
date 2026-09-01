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
          <p>Dopo il ritorno alla scheda, premi il pulsante verde <strong>FATTO</strong>.</p>
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
          ? `Salvate: ${tasks.map((task) => PHASES[task.phase].label).join(" e ")}. Ora puoi premere FATTO.`
          : "Nessuna attività creata. Ora puoi premere FATTO.";
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
