(() => {
  "use strict";

  if (window.HeraTreeWorkOrders?.installed) return;

  const COMMESSA_ID = "potature-abbattimenti";
  const COMMESSA_NAME = "Potature Abbattimenti";
  const COMMESSA_CODE = "POT-ABB";
  const DATASET_NAME = "alberi-manutenzioni";

  const text = (value) => String(value ?? "").trim();
  const esc = (value) => text(value).replace(/[&<>'"]/g, (char) => ({
    "&": "&amp;", "<": "&lt;", ">": "&gt;", "'": "&#39;", '"': "&quot;"
  }[char]));

  function database() {
    try {
      if (typeof db !== "undefined" && db?.collection) return db;
    } catch (_) {}
    return window.db?.collection ? window.db : null;
  }

  function authenticatedUser() {
    try {
      if (typeof auth !== "undefined" && auth?.currentUser) return auth.currentUser;
      if (typeof currentUser !== "undefined" && currentUser) return currentUser;
    } catch (_) {}
    return window.currentUser || null;
  }

  function canCreateCommessa() {
    try {
      return typeof canManageData === "function" && canManageData();
    } catch (_) {
      return false;
    }
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
      .slice(0, 70);
  }

  function treeIdentifier(context) {
    return text(context?.tree?.num_pt || context?.tree?.cod_alb || context?.details?.[0]?.value || "albero");
  }

  function plantIdFor(context) {
    const municipality = slug(context?.municipality || "comune") || "comune";
    const identifier = slug(treeIdentifier(context)) || "albero";
    const point = context?.point || {};
    const coordinateKey = `${Number(point.lat).toFixed(6)}|${Number(point.lon).toFixed(6)}`;
    return `albero-${municipality}-${identifier}-${stableHash(coordinateKey)}`;
  }

  function firstSixDetails(context) {
    return (Array.isArray(context?.details) ? context.details : []).slice(0, 6).map((entry) => ({
      campo: text(entry?.key),
      etichetta: text(entry?.label),
      valore: text(entry?.value)
    }));
  }

  function navigationUrl(context) {
    if (text(context?.navigationUrl)) return text(context.navigationUrl);
    const lat = Number(context?.point?.lat);
    const lon = Number(context?.point?.lon);
    return Number.isFinite(lat) && Number.isFinite(lon)
      ? `https://www.google.com/maps/dir/?api=1&destination=${lat},${lon}`
      : "";
  }

  function buildTreeMessageBlock(context, intervention, requestedWork, operatorNote) {
    return [
      "🌳 SCHEDA POTATURE / ABBATTIMENTI",
      ...firstSixDetails(context).map((entry) => `• ${entry.etichetta}: ${entry.valore}`),
      `🪚 Intervento richiesto: ${intervention}`,
      `📋 Lavorazione: ${requestedWork}`,
      operatorNote ? `📝 Nota operatore: ${operatorNote}` : "",
      navigationUrl(context) ? `📍 Naviga verso l’albero: ${navigationUrl(context)}` : ""
    ].filter(Boolean).join("\n");
  }

  function buildPlantPayload(context, values) {
    const tree = context.tree || {};
    const point = context.point || {};
    const identifier = treeIdentifier(context);
    const species = text(tree.classe || tree.nome_scientifico || tree.nome_comune || "Specie non disponibile");
    const municipality = text(context.municipality || "Bologna");
    const address = text(tree.via || tree.indirizzo || tree.localizzazione || tree.quartiere || "Posizione da censimento comunale");
    const requestedWork = text(values.requestedWork) || `${values.intervention} dell’albero ${identifier}`;
    const dedicatedMessage = buildTreeMessageBlock(context, values.intervention, requestedWork, values.operatorNote);
    const plantId = plantIdFor(context);
    const user = authenticatedUser();
    let operatorName = text(user?.displayName || user?.email || "Operatore");
    try {
      if (typeof getOperatorDisplayName === "function") operatorName = text(getOperatorDisplayName()) || operatorName;
    } catch (_) {}

    return {
      id: plantId,
      commessaId: COMMESSA_ID,
      idSap: `ALB-${slug(municipality).toUpperCase()}-${slug(identifier).toUpperCase()}`,
      denominazione: `ALBERO #${identifier} — ${species}`,
      nome: `ALBERO #${identifier} — ${species}`,
      comune: municipality,
      indirizzo: address,
      descrizioneVia: address,
      area: text(tree.quartiere),
      competenza: `COMUNE DI ${municipality.toUpperCase()}`,
      gpsY: Number(point.lat),
      gpsX: Number(point.lon),
      latitudine: Number(point.lat),
      longitudine: Number(point.lon),
      coordinate: `${Number(point.lat).toFixed(6)}, ${Number(point.lon).toFixed(6)}`,
      tipologia: "ALBERO",
      tipologiaImpianto: "POTATURE / ABBATTIMENTI",
      tipologiaIntervento: text(values.intervention),
      tipologiaLavorazione: requestedWork,
      lavorazioniRichieste: requestedWork,
      codicePrezzo: text(values.intervention),
      voceRiferimento: text(values.intervention),
      hasOrdinario: false,
      hasStraordinario: true,
      tipoManutenzione: "Straordinaria",
      note: dedicatedMessage,
      noteImpianto: dedicatedMessage,
      noteOperatore: text(values.operatorNote),
      hasNote: true,
      alberoCatasto: true,
      potatureAbbattimenti: true,
      schedaParziale: true,
      schedaVersione: 1,
      campiDaCompletare: true,
      numeroPunto: text(tree.num_pt),
      codiceAlbero: text(tree.cod_alb),
      specieAlbero: species,
      dettagliCatastoPrimiSei: firstSixDetails(context),
      catastoFonte: "Comune di Bologna · censimento ufficiale",
      catastoDataset: DATASET_NAME,
      messaggioWhazzupTipo: "POTATURE_ABBATTIMENTI",
      messaggioWhazzupAlbero: dedicatedMessage,
      updatedByUid: text(user?.uid),
      updatedByName: operatorName
    };
  }

  function sameValue(left, right) {
    if (left === right) return true;
    if (left == null && right == null) return true;
    if (typeof left === "number" || typeof right === "number") return Number(left) === Number(right);
    if (typeof left === "object" || typeof right === "object") {
      try { return JSON.stringify(left) === JSON.stringify(right); } catch (_) { return false; }
    }
    return text(left) === text(right);
  }

  function hasPayloadChanges(existing, payload) {
    return Object.entries(payload).some(([key, value]) => !sameValue(existing?.[key], value));
  }

  async function saveWorkOrder(context, values) {
    const store = database();
    const user = authenticatedUser();
    if (!user) throw new Error("Devi accedere all’app prima di creare il cantiere.");
    if (!store) throw new Error("Database non disponibile. Controlla la connessione e riprova.");
    if (!text(values.intervention)) throw new Error("Seleziona Potatura oppure Abbattimento.");

    const payload = buildPlantPayload(context, values);
    const plantId = payload.id;
    const commessaRef = store.collection(collectionName()).doc(COMMESSA_ID);
    const plantRef = commessaRef.collection("impianti").doc(plantId);
    const [commessaSnapshot, plantSnapshot] = await Promise.all([commessaRef.get(), plantRef.get()]);

    if (!commessaSnapshot.exists && !canCreateCommessa()) {
      throw new Error("La commessa Potature Abbattimenti deve essere inizializzata una volta da un amministratore.");
    }

    const batch = store.batch();
    let writes = 0;
    if (!commessaSnapshot.exists) {
      batch.set(commessaRef, {
        nome: COMMESSA_NAME,
        codice: COMMESSA_CODE,
        speciale: true,
        tipoSpeciale: "POTATURE_ABBATTIMENTI",
        catastoAlberi: true,
        attiva: true,
        parentCommessaId: null,
        excelModelVersion: 2,
        priceListVersion: 2,
        workItemsModelVersion: 2,
        percentualeRibassoGenerale: 0.01,
        nextImpiantoNumber: 1,
        creatoDa: text(user.email),
        createdAt: serverTimestamp()
      }, { merge: true });
      writes += 1;
    }

    const existing = plantSnapshot.exists ? (plantSnapshot.data() || {}) : null;
    if (!existing || hasPayloadChanges(existing, payload)) {
      const timestamps = existing
        ? { updatedAt: serverTimestamp() }
        : { stato: "da_fare", done: false, doneAt: null, doneBy: "", createdAt: serverTimestamp(), updatedAt: serverTimestamp() };
      batch.set(plantRef, { ...payload, ...timestamps }, { merge: true });
      writes += 1;
    }

    if (writes) await batch.commit();
    return {
      plantId,
      payload,
      created: !plantSnapshot.exists,
      updated: Boolean(plantSnapshot.exists && writes),
      unchanged: Boolean(plantSnapshot.exists && !writes)
    };
  }

  function closePanel() {
    document.querySelector(".tree-work-order-panel")?.remove();
  }

  function openCreatedCantiere(saved) {
    document.getElementById("tree-search-back-btn")?.click();
    try {
      if (typeof selectCommessa === "function") selectCommessa(COMMESSA_ID, COMMESSA_NAME, COMMESSA_CODE);
      const impiantoKey = typeof buildImpiantoKey === "function"
        ? buildImpiantoKey(saved.payload)
        : `sap:${text(saved.payload.idSap).toLowerCase()}`;
      window.location.hash = `commessa=${encodeURIComponent(COMMESSA_ID)}&impianto=${encodeURIComponent(impiantoKey)}`;
      if (typeof applyRoute === "function") applyRoute();
    } catch (error) {
      console.error("Apertura cantiere Potature Abbattimenti non riuscita:", error);
      window.location.hash = `commessa=${encodeURIComponent(COMMESSA_ID)}`;
    }
  }

  function closePanelAndRestore() {
    closePanel();
    document.querySelector(".tree-work-order-open")?.focus();
  }

  function renderPanel(context) {
    const host = document.getElementById("tree-result");
    if (!host) return;
    closePanel();
    const details = firstSixDetails(context);
    const panel = document.createElement("section");
    panel.className = "tree-work-order-panel";
    panel.setAttribute("aria-labelledby", "tree-work-order-title");
    panel.innerHTML = `
      <header class="tree-work-order-head">
        <div><small>Commessa speciale · ${COMMESSA_NAME}</small><h2 id="tree-work-order-title">✂️ Nuovo cantiere albero</h2></div>
        <button class="btn tree-work-order-close" type="button">CHIUDI</button>
      </header>
      <p class="tree-work-order-intro">La scheda iniziale è già compilata con i primi 6 dettagli del catasto. Gli altri campi verranno aggiunti nella fase successiva.</p>
      <form class="tree-work-order-form">
        <fieldset class="tree-work-order-prefill"><legend>Dati copiati dal Catasto alberi</legend>
          ${details.map((entry) => `<label><span>${esc(entry.etichetta)}</span><input type="text" value="${esc(entry.valore)}" readonly></label>`).join("")}
        </fieldset>
        <div class="tree-work-order-editable">
          <label for="tree-work-order-intervention"><span>Intervento richiesto *</span><select id="tree-work-order-intervention" name="intervention" required><option value="">Seleziona</option><option value="POTATURA">Potatura</option><option value="ABBATTIMENTO">Abbattimento</option><option value="POTATURA E ABBATTIMENTO">Potatura e abbattimento</option></select></label>
          <label for="tree-work-order-request"><span>Lavorazione richiesta *</span><textarea id="tree-work-order-request" name="requestedWork" rows="3" required placeholder="Es. Rimonda del secco e contenimento chioma"></textarea></label>
          <label for="tree-work-order-note"><span>Nota operatore</span><textarea id="tree-work-order-note" name="operatorNote" rows="3" placeholder="Informazioni utili per la squadra"></textarea></label>
        </div>
        <aside class="tree-work-order-message-preview"><strong>Messaggio Whazzup dedicato</strong><p>Conterrà la scheda albero, i primi 6 dettagli, l’intervento richiesto e il collegamento per navigare verso l’albero.</p></aside>
        <p class="tree-work-order-feedback" role="status" aria-live="polite"></p>
        <button class="btn btn-primary tree-work-order-save" type="submit">SALVA E CREA IL CANTIERE</button>
      </form>`;
    host.appendChild(panel);
    panel.querySelector(".tree-work-order-close")?.addEventListener("click", closePanelAndRestore);
    const intervention = panel.querySelector("#tree-work-order-intervention");
    const requestedWork = panel.querySelector("#tree-work-order-request");
    intervention?.addEventListener("change", () => {
      if (!requestedWork || requestedWork.dataset.userEdited === "1") return;
      requestedWork.value = intervention.value ? `${intervention.options[intervention.selectedIndex].text} dell’albero` : "";
    });
    requestedWork?.addEventListener("input", () => { requestedWork.dataset.userEdited = "1"; });
    panel.querySelector("form")?.addEventListener("submit", async (event) => {
      event.preventDefault();
      const saveButton = panel.querySelector(".tree-work-order-save");
      const feedback = panel.querySelector(".tree-work-order-feedback");
      const form = new FormData(event.currentTarget);
      saveButton.disabled = true;
      feedback.className = "tree-work-order-feedback";
      feedback.textContent = "Salvataggio del cantiere in corso…";
      try {
        const saved = await saveWorkOrder(context, {
          intervention: text(form.get("intervention")),
          requestedWork: text(form.get("requestedWork")),
          operatorNote: text(form.get("operatorNote"))
        });
        feedback.classList.add("success");
        feedback.innerHTML = `${saved.created ? "✅ Albero trasformato in cantiere." : saved.updated ? "✅ Scheda del cantiere aggiornata." : "✅ Il cantiere era già aggiornato."}<br><span>Aprilo nella commessa Potature Abbattimenti per usare NAVIGA, ALLEGA e FATTO.</span>`;
        const openButton = document.createElement("button");
        openButton.type = "button";
        openButton.className = "btn tree-work-order-open-created";
        openButton.textContent = "APRI IL CANTIERE";
        openButton.addEventListener("click", () => openCreatedCantiere(saved));
        feedback.appendChild(openButton);
      } catch (error) {
        feedback.classList.add("error");
        feedback.textContent = error?.message || "Impossibile creare il cantiere.";
      } finally {
        saveButton.disabled = false;
      }
    });
    panel.scrollIntoView({ behavior: "smooth", block: "start" });
  }

  function open(context) {
    if (!context?.tree || !context?.point) return;
    renderPanel(context);
  }

  window.HeraTreeWorkOrders = Object.freeze({
    installed: true,
    version: "1.0.0",
    commessaId: COMMESSA_ID,
    commessaName: COMMESSA_NAME,
    open,
    close: closePanel,
    buildPlantPayload,
    plantIdFor
  });
})();
