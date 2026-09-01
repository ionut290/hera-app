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
          <p>Le domande sono facoltative. Confermando, il cantiere verrà completato con <strong>TERMINATO</strong>.</p>
          <button class="btn btn-primary potature-followup-save" type="submit">SALVA E TERMINA</button>
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
          ? `Salvate: ${tasks.map((task) => PHASES[task.phase].label).join(" e ")}. Chiusura del cantiere in corso…`
          : "Nessuna attività creata. Chiusura del cantiere in corso…";
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

  const POTATURE_ID = "potature-abbattimenti";
  const TERMINATO_ACTION = "special-terminato";
  const text = (value) => String(value ?? "").trim();
  const normalize = (value) => text(value)
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, "-")
    .replace(/^-|-$/g, "");

  function selectedId() {
    const match = String(window.location.hash || "").match(/(?:^|[&#])commessa=([^&]+)/);
    if (match?.[1]) {
      try { return decodeURIComponent(match[1]); } catch (_) { return match[1]; }
    }
    try {
      if (typeof selectedCommessaId !== "undefined") return text(selectedCommessaId);
    } catch (_) {}
    return text(window.selectedCommessaId);
  }

  function selectedName() {
    try {
      if (typeof selectedCommessaName !== "undefined") return text(selectedCommessaName);
    } catch (_) {}
    return text(window.selectedCommessaName || document.getElementById("impianti-page-title")?.textContent);
  }

  function isSfalcioCobo(commessaId = selectedId(), commessaName = selectedName()) {
    const value = `${normalize(commessaId)} ${normalize(commessaName)}`;
    return value.includes("sfalcio") && value.includes("cobo");
  }

  function isSpecialCommessa(commessaId = selectedId(), commessaName = selectedName()) {
    return normalize(commessaId) === POTATURE_ID || isSfalcioCobo(commessaId, commessaName);
  }

  function database() {
    try {
      if (typeof db !== "undefined" && db?.collection && db?.batch) return db;
    } catch (_) {}
    return window.db?.collection && window.db?.batch ? window.db : null;
  }

  function collectionName() {
    try {
      if (typeof getCommesseCollectionName === "function") return text(getCommesseCollectionName()) || "commesse";
    } catch (_) {}
    return "commesse";
  }

  function authUser() {
    try {
      if (typeof auth !== "undefined" && auth?.currentUser) return auth.currentUser;
    } catch (_) {}
    return window.auth?.currentUser || window.currentUser || null;
  }

  function operatorName() {
    try {
      if (typeof getOperatorDisplayName === "function") {
        const value = text(getOperatorDisplayName());
        if (value) return value;
      }
    } catch (_) {}
    const user = authUser();
    return text(user?.displayName || user?.email || "Operatore");
  }

  function currentItems() {
    try {
      if (typeof currentImpianti !== "undefined" && Array.isArray(currentImpianti)) return currentImpianti;
    } catch (_) {}
    return Array.isArray(window.currentImpianti) ? window.currentImpianti : [];
  }

  function itemKey(item) {
    try {
      if (typeof buildImpiantoKey === "function") return text(buildImpiantoKey(item));
    } catch (_) {}
    return text(item?.id || item?.idSap || item?.denominazione);
  }

  function findItem(card) {
    const key = text(card?.dataset?.impiantoKey);
    if (!key) return null;
    return currentItems().find((item) => itemKey(item) === key) || null;
  }

  function documentIds(item = {}) {
    const ids = [];
    if (Array.isArray(item.sourceIds)) item.sourceIds.forEach((id) => { if (text(id)) ids.push(text(id)); });
    if (text(item.id)) ids.push(text(item.id));
    return [...new Set(ids)];
  }

  function isDone(item = {}) {
    if (item.done === true || item.fatto === true || item.completed === true) return true;
    return ["fatto", "done", "completed", "completato", "terminato"].includes(text(item.stato || item.status).toLowerCase());
  }

  function completedWhazzup(item, doneAt, doneBy) {
    let handler = null;
    try {
      if (typeof handleCompletedImpiantoWhatsAppClick === "function") handler = handleCompletedImpiantoWhatsAppClick;
    } catch (_) {}
    if (!handler && typeof window.handleCompletedImpiantoWhatsAppClick === "function") handler = window.handleCompletedImpiantoWhatsAppClick;
    if (!handler) {
      console.warn("TERMINATO salvato: apertura Whazzup completato non disponibile.");
      return false;
    }
    try {
      const result = handler({ ...item, done: true, stato: "fatto", doneAt, doneBy });
      if (result && typeof result.catch === "function") result.catch((error) => console.warn("Apertura Whazzup dopo TERMINATO non riuscita:", error));
      return true;
    } catch (error) {
      console.warn("Apertura Whazzup dopo TERMINATO non riuscita:", error);
      return false;
    }
  }

  async function terminate(item, options = {}) {
    const commessaId = text(options.commessaId || selectedId());
    if (!isSpecialCommessa(commessaId, selectedName())) throw new Error("TERMINATO è disponibile solo per Potature Abbattimenti e Sfalcio COBO.");
    const store = database();
    if (!store) throw new Error("Database non disponibile. Controlla la connessione e riprova.");
    const ids = documentIds(item);
    if (!ids.length) throw new Error("Impossibile identificare il cantiere da terminare.");

    const user = authUser();
    const doneBy = operatorName();
    const doneAt = new Date();
    const payload = {
      done: true,
      stato: "fatto",
      doneAt,
      doneBy,
      doneByUid: text(user?.uid),
      doneByEmail: text(user?.email),
      terminato: true,
      terminatoAt: doneAt,
      terminatoBy: doneBy,
      terminatoByUid: text(user?.uid),
      terminatoByEmail: text(user?.email),
      completionMode: "TERMINATO_SPECIAL",
      updatedAt: doneAt
    };

    const ref = store.collection(collectionName()).doc(commessaId).collection("impianti");
    const batch = store.batch();
    ids.forEach((id) => batch.set(ref.doc(id), payload, { merge: true }));
    await batch.commit();
    completedWhazzup(item, doneAt, doneBy);
    return { doneAt, doneBy, ids, payload };
  }

  function adaptPotatureModal() {
    const modal = document.getElementById("potature-followup-modal");
    if (!modal) return;
    const title = modal.querySelector("#potature-followup-title");
    const footerText = modal.querySelector(".potature-followup-actions p");
    const saveButton = modal.querySelector(".potature-followup-save");
    if (title) title.textContent = "Termina cantiere";
    if (footerText) footerText.innerHTML = "Le domande sono facoltative. Premendo <strong>TERMINA CANTIERE</strong> il cantiere passa nei Fatti.";
    if (saveButton) saveButton.textContent = "✅ TERMINA CANTIERE";
  }

  function potatureOptions(item, onComplete) {
    const user = authUser();
    return {
      source: item,
      existingItems: currentItems(),
      store: database(),
      collectionName: collectionName(),
      commessaId: POTATURE_ID,
      operatorUid: text(user?.uid),
      operatorName: operatorName(),
      onComplete
    };
  }

  async function perform(button, item) {
    button.disabled = true;
    const originalText = button.textContent;
    button.textContent = "SALVATAGGIO…";
    try {
      await terminate(item);
      button.textContent = "✅ TERMINATO";
      button.classList.add("is-completed-done");
    } catch (error) {
      console.error("TERMINATO speciale non riuscito:", error);
      button.disabled = false;
      button.textContent = originalText;
      window.alert(error?.message || "Impossibile terminare il cantiere.");
    }
  }

  function startPotature(button, item) {
    const followup = window.HeraPotatureFollowup;
    if (!followup?.isOriginal?.(item)) {
      void perform(button, item);
      return;
    }
    button.disabled = true;
    try {
      followup.open(potatureOptions(item, (result) => {
        if (result?.skipped) {
          button.disabled = false;
          return;
        }
        void perform(button, item);
      }));
      adaptPotatureModal();
    } catch (error) {
      button.disabled = false;
      console.error("Apertura form TERMINATO Potature non riuscita:", error);
      window.alert(error?.message || "Impossibile aprire il form di chiusura.");
    }
  }

  function completionLabel(button) {
    return `${button?.textContent || ""} ${button?.title || ""} ${button?.getAttribute?.("aria-label") || ""}`
      .toLowerCase()
      .replace(/\s+/g, " ")
      .trim();
  }

  function isLegacyCompletionButton(button) {
    if (!button || button.dataset?.actionKey === TERMINATO_ACTION) return false;
    if (button.matches?.(".impianto-force-done-btn, [data-hidden-move-done-btn=\"1\"]")) return true;
    const label = completionLabel(button);
    if (label.includes("prepara fine") || label.includes("forza in fatti") || label.includes("whazzup / fatto")) return true;
    if (label === "fatto" || label === "✅ fatto" || label === "✓ fatto") return true;
    if (button.matches?.('.action-icon-btn[data-action-key="whatsapp"]') && label.includes("fatto")) return true;
    return false;
  }

  function legacyCompletionButtons(card) {
    return [...card.querySelectorAll("button")].filter(isLegacyCompletionButton);
  }

  function primaryLegacyButton(card) {
    const buttons = legacyCompletionButtons(card);
    if (!buttons.length) return null;
    return buttons.find((button) => button.closest(".impianto-primary-actions"))
      || buttons.find((button) => completionLabel(button).includes("whazzup / fatto"))
      || buttons[0];
  }

  function makeInvisible(element) {
    if (!element) return;
    element.hidden = true;
    element.setAttribute("aria-hidden", "true");
    element.setAttribute("tabindex", "-1");
    element.dataset.specialTerminatoReplaced = "1";
    element.classList.add("special-terminato-legacy-hidden");
    element.style.setProperty("display", "none", "important");
  }

  function hideLegacyCompletion(card, item) {
    if (isDone(item)) return [];
    const buttons = legacyCompletionButtons(card);
    buttons.forEach(makeInvisible);
    return buttons;
  }

  function terminatoClassName(legacyTarget) {
    const legacyClasses = text(legacyTarget?.className)
      .split(/\s+/)
      .filter(Boolean)
      .filter((name) => !["special-terminato-legacy-hidden", "impianto-force-done-btn"].includes(name));
    const classes = new Set(legacyClasses.length ? legacyClasses : ["btn", "btn-primary"]);
    classes.add("special-terminato-btn");
    classes.add("btn-primary");
    return [...classes].join(" ");
  }

  function placeAtLegacyPosition(button, legacyTarget, card) {
    if (legacyTarget?.parentElement) {
      legacyTarget.parentElement.insertBefore(button, legacyTarget);
      button.dataset.replacesAction = "fatto";
      return true;
    }
    const row = card.querySelector(".impianto-primary-actions") || card.querySelector(".impianto-actions");
    if (!row) return false;
    row.appendChild(button);
    return true;
  }

  function ensureButton(card, item) {
    const old = card.querySelector(`[data-action-key="${TERMINATO_ACTION}"]`);
    if (isDone(item)) {
      old?.remove();
      return;
    }

    const legacyTarget = primaryLegacyButton(card);
    const desiredClassName = terminatoClassName(legacyTarget);
    hideLegacyCompletion(card, item);

    if (old) {
      old.className = desiredClassName;
      if (legacyTarget?.parentElement && (old.parentElement !== legacyTarget.parentElement || old.nextElementSibling !== legacyTarget)) {
        legacyTarget.parentElement.insertBefore(old, legacyTarget);
      }
      old.dataset.replacesAction = legacyTarget ? "fatto" : "fallback";
      return;
    }

    const button = document.createElement("button");
    button.type = "button";
    button.className = desiredClassName;
    button.dataset.actionKey = TERMINATO_ACTION;
    button.textContent = "✅ TERMINATO";
    button.setAttribute("aria-label", "Termina il cantiere e spostalo nei Fatti");
    button.title = "Termina il cantiere";
    button.addEventListener("click", (event) => {
      event.preventDefault();
      event.stopPropagation();
      if (normalize(selectedId()) === POTATURE_ID) startPotature(button, item);
      else void perform(button, item);
    });
    placeAtLegacyPosition(button, legacyTarget, card);
  }

  function apply() {
    if (!isSpecialCommessa()) return;
    const list = document.getElementById("impianti-lista") || document.querySelector(".impianti-lista");
    if (!list) return;
    list.querySelectorAll("[data-impianto-key]").forEach((card) => {
      const item = findItem(card);
      if (item) ensureButton(card, item);
    });
  }

  let scheduled = false;
  function scheduleApply() {
    if (scheduled) return;
    scheduled = true;
    window.requestAnimationFrame(() => {
      scheduled = false;
      apply();
    });
  }

  const observer = new MutationObserver(scheduleApply);
  observer.observe(document.documentElement, { childList: true, subtree: true });
  window.addEventListener("hashchange", scheduleApply);
  window.addEventListener("hera:data-ready", scheduleApply);
  document.addEventListener("DOMContentLoaded", scheduleApply, { once: true });
  scheduleApply();

  window.HeraSpecialTerminato = Object.freeze({
    installed: true,
    isSpecialCommessa,
    isSfalcioCobo,
    documentIds,
    terminate,
    apply
  });
})();
