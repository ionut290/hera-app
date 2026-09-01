(() => {
  "use strict";

  if (window.HeraCoboMowing?.installed) return;

  const COMMESSA_ID = "sfalcio-cobo";
  const COMMESSA_NAME = "Sfalcio COBO";
  const COMMESSA_CODE = "COBO-SFALCIO";
  const SOURCE_DATASET = "carta-tecnica-comunale-toponimi-parchi-e-giardini";
  const MODAL_ID = "cobo-mowing-modal";

  let returnFocusTarget = null;

  const text = (value) => String(value ?? "").trim();
  const esc = (value) => text(value).replace(/[&<>'"]/g, (character) => ({
    "&": "&amp;", "<": "&lt;", ">": "&gt;", "'": "&#39;", '"': "&quot;"
  }[character]));

  function normalize(value) {
    return text(value)
      .normalize("NFD")
      .replace(/[\u0300-\u036f]/g, "")
      .toLocaleLowerCase("it-IT")
      .replace(/[^a-z0-9]/g, "");
  }

  function fieldValue(record, names) {
    const values = new Map(Object.entries(record || {}).map(([key, value]) => [normalize(key), value]));
    for (const name of names) {
      const value = values.get(normalize(name));
      if (value !== undefined && value !== null && text(value)) return value;
    }
    return "";
  }

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

  function pointOf(context) {
    const lat = Number(context?.point?.lat ?? fieldValue(context?.record, ["lat", "latitude", "y"]));
    const lon = Number(context?.point?.lon ?? context?.point?.lng ?? fieldValue(context?.record, ["lon", "lng", "longitude", "x"]));
    if (!Number.isFinite(lat) || !Number.isFinite(lon)) return null;
    if (lat < -90 || lat > 90 || lon < -180 || lon > 180) return null;
    return { lat, lon };
  }

  function parkNameOf(context) {
    return text(context?.parkName || fieldValue(context?.record, [
      "nomevia", "nome_via", "nome via", "completo", "porzione", "denominazione", "nome", "name", "toponimo"
    ]) || "Parco / giardino");
  }

  function parkCodeOf(context) {
    return text(context?.parkCode || fieldValue(context?.record, ["codvia", "cod_via", "codice via"]));
  }

  function parkFallbackCodeOf(context) {
    return text(fieldValue(context?.record, ["cod_ogg", "codogg", "idtopon", "id", "objectid"]));
  }

  function parkQuarterOf(context) {
    return text(context?.quarter || fieldValue(context?.record, ["quartiere", "nome_quartiere", "quart"]));
  }

  function parkAddressOf(context) {
    return text(context?.address || fieldValue(context?.record, [
      "ubicazione", "indirizzo", "via", "localizzazione", "nomevia", "nome_via"
    ]) || parkNameOf(context));
  }

  function identifierOf(context) {
    const officialCode = parkCodeOf(context);
    if (officialCode && !/^0(?:[.,]0+)?$/.test(officialCode)) return officialCode;
    const fallbackCode = parkFallbackCodeOf(context);
    if (fallbackCode) return fallbackCode;
    const point = pointOf(context);
    return stableHash(`${parkNameOf(context)}|${point?.lat || ""}|${point?.lon || ""}`);
  }

  function plantIdFor(context) {
    const identifier = slug(identifierOf(context)) || stableHash(parkNameOf(context));
    return `cobo-parco-${identifier}`;
  }

  function isGeometryField(key, value) {
    const normalizedKey = normalize(key);
    return normalizedKey.includes("geoshape")
      || normalizedKey === "geometry"
      || normalizedKey === "geom"
      || normalizedKey.includes("geopoint")
      || (value && typeof value === "object" && (value.type || value.geometry || value.coordinates));
  }

  function displayValue(value) {
    if (value === null || value === undefined || value === "") return "";
    if (typeof value === "boolean") return value ? "Sì" : "No";
    if (typeof value === "object") return "";
    return text(value);
  }

  function firstParkDetails(context) {
    const record = context?.record || {};
    const priority = ["codvia", "nomevia", "tipo", "quartiere", "ubicazione", "indirizzo"];
    return Object.entries(record)
      .filter(([key, value]) => !String(key).startsWith("__vb") && !isGeometryField(key, value))
      .map(([key, value]) => ({ key, value: displayValue(value) }))
      .filter((entry) => entry.value)
      .sort((left, right) => {
        const leftIndex = priority.indexOf(normalize(left.key));
        const rightIndex = priority.indexOf(normalize(right.key));
        if (leftIndex >= 0 || rightIndex >= 0) return (leftIndex < 0 ? 999 : leftIndex) - (rightIndex < 0 ? 999 : rightIndex);
        return left.key.localeCompare(right.key, "it", { sensitivity: "base" });
      })
      .slice(0, 6)
      .map((entry) => ({ campo: text(entry.key), etichetta: text(entry.key).replace(/_/g, " "), valore: entry.value }));
  }

  function parseOptionalNumber(value) {
    const normalized = text(value).replace(/\s/g, "").replace(",", ".");
    if (!normalized) return null;
    const parsed = Number(normalized);
    return Number.isFinite(parsed) && parsed >= 0 ? parsed : null;
  }

  function getOperatorName(user) {
    let operatorName = text(user?.displayName || user?.email || "Operatore");
    try {
      if (typeof getOperatorDisplayName === "function") operatorName = text(getOperatorDisplayName()) || operatorName;
    } catch (_) {}
    return operatorName;
  }

  function buildPlantPayload(context, values = {}) {
    const point = pointOf(context);
    if (!point) throw new Error("Il parco non dispone di coordinate valide per creare il cantiere.");
    const user = authenticatedUser();
    const parkName = parkNameOf(context);
    const parkCode = parkCodeOf(context) || parkFallbackCodeOf(context) || identifierOf(context);
    const quarter = parkQuarterOf(context);
    const address = parkAddressOf(context);
    const areaMq = parseOptionalNumber(values.areaMq);
    const workType = text(values.workType || "SFALCIO COMPLETO");
    const requestedWork = text(values.requestedWork || `${workType} del parco`);
    const operatorNote = text(values.operatorNote);

    return {
      id: plantIdFor(context),
      commessaId: COMMESSA_ID,
      idSap: `COBO-${slug(parkCode).toUpperCase()}`,
      denominazione: parkName,
      nome: parkName,
      comune: "Bologna",
      indirizzo: address,
      descrizioneVia: address,
      area: quarter,
      competenza: "COMUNE DI BOLOGNA",
      gpsY: point.lat,
      gpsX: point.lon,
      latitudine: point.lat,
      longitudine: point.lon,
      coordinate: `${point.lat.toFixed(6)}, ${point.lon.toFixed(6)}`,
      tipologia: "PARCO / GIARDINO",
      tipologiaImpianto: "SFALCIO COBO",
      tipologiaIntervento: workType,
      lavorazioniRichieste: requestedWork,
      codicePrezzo: "A11",
      voceRiferimento: "SFALCIO COBO",
      hasOrdinario: true,
      hasStraordinario: false,
      tipoManutenzione: "Ordinaria",
      sfalci: areaMq,
      sfalciMq: areaMq,
      note: operatorNote,
      noteImpianto: operatorNote,
      noteOperatore: operatorNote,
      hasNote: Boolean(operatorNote),
      sfalcioCobo: true,
      parcoCobo: true,
      parcoCodvia: parkCode,
      parcoNomevia: parkName,
      parcoQuartiere: quarter,
      parcoConfiniDisponibili: Boolean(context?.boundaryAvailable),
      parcoDettagliPrimiSei: firstParkDetails(context),
      parcoFonte: "Comune di Bologna · Open Data ufficiali",
      parcoDataset: SOURCE_DATASET,
      schedaVersione: 1,
      updatedByUid: text(user?.uid),
      updatedByName: getOperatorName(user)
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

    const payload = buildPlantPayload(context, values);
    const plantId = payload.id;
    const commessaRef = store.collection(collectionName()).doc(COMMESSA_ID);
    const plantRef = commessaRef.collection("impianti").doc(plantId);
    const [commessaSnapshot, plantSnapshot] = await Promise.all([commessaRef.get(), plantRef.get()]);

    if (!commessaSnapshot.exists && !canCreateCommessa()) {
      throw new Error("La commessa Sfalcio COBO deve essere inizializzata una volta da un amministratore.");
    }

    const batch = store.batch();
    let writes = 0;
    if (!commessaSnapshot.exists) {
      batch.set(commessaRef, {
        nome: COMMESSA_NAME,
        codice: COMMESSA_CODE,
        speciale: true,
        tipoSpeciale: "SFALCIO_COBO",
        verdeBologna: true,
        parchiGiardini: true,
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
      unchanged: Boolean(plantSnapshot.exists && !writes),
      alreadyDone: Boolean(existing?.done)
    };
  }

  function closeModal({ restoreFocus = true } = {}) {
    document.getElementById(MODAL_ID)?.remove();
    document.body.classList.remove("cobo-mowing-modal-open");
    if (restoreFocus) returnFocusTarget?.focus?.();
    returnFocusTarget = null;
  }

  function openCreatedCantiere(saved) {
    closeModal({ restoreFocus: false });
    window.HeraVerdeBologna?.close?.();
    try {
      if (typeof selectCommessa === "function") selectCommessa(COMMESSA_ID, COMMESSA_NAME, COMMESSA_CODE);
      const impiantoKey = typeof buildImpiantoKey === "function"
        ? buildImpiantoKey(saved.payload)
        : `sap:${text(saved.payload.idSap).toLowerCase()}`;
      window.location.hash = `commessa=${encodeURIComponent(COMMESSA_ID)}&impianto=${encodeURIComponent(impiantoKey)}`;
      if (typeof applyRoute === "function") applyRoute();
    } catch (error) {
      console.error("Apertura cantiere Sfalcio COBO non riuscita:", error);
      window.location.hash = `commessa=${encodeURIComponent(COMMESSA_ID)}`;
    }
  }

  function createModal(title, subtitle) {
    document.getElementById(MODAL_ID)?.remove();
    document.body.classList.remove("cobo-mowing-modal-open");
    const modal = document.createElement("section");
    modal.id = MODAL_ID;
    modal.className = "cobo-mowing-modal";
    modal.innerHTML = `
      <div class="cobo-mowing-screen" role="dialog" aria-modal="true" aria-labelledby="cobo-mowing-title">
        <header class="cobo-mowing-head">
          <div><p>${esc(subtitle)}</p><h2 id="cobo-mowing-title">${esc(title)}</h2></div>
          <button class="btn" type="button" data-cobo-close>CHIUDI</button>
        </header>
        <div class="cobo-mowing-content"></div>
      </div>`;
    document.body.appendChild(modal);
    document.body.classList.add("cobo-mowing-modal-open");
    modal.querySelector("[data-cobo-close]")?.addEventListener("click", () => closeModal());
    modal.addEventListener("keydown", (event) => {
      if (event.key === "Escape") closeModal();
    });
    return modal;
  }

  function openCreate(context = {}) {
    returnFocusTarget = context.triggerButton || document.activeElement;
    const point = pointOf(context);
    const details = firstParkDetails(context);
    const modal = createModal("Nuovo cantiere parco", `Commessa · ${COMMESSA_NAME}`);
    const content = modal.querySelector(".cobo-mowing-content");
    content.innerHTML = `
      <section class="cobo-mowing-park-summary">
        <strong>${esc(parkNameOf(context))}</strong>
        <span>CODVIA ${esc(parkCodeOf(context) || "—")} · ${esc(parkQuarterOf(context) || "Quartiere non indicato")}</span>
        <span>${point ? `${point.lat.toFixed(6)}, ${point.lon.toFixed(6)}` : "Coordinate non disponibili"}</span>
      </section>
      <form class="cobo-mowing-form">
        <fieldset class="cobo-mowing-prefill"><legend>Dati del parco</legend>
          ${details.length ? details.map((entry) => `<label><span>${esc(entry.etichetta)}</span><input type="text" value="${esc(entry.valore)}" readonly></label>`).join("") : "<p>Dati essenziali copiati dalla scheda ufficiale.</p>"}
        </fieldset>
        <label><span>Tipo di lavorazione *</span><select name="workType" required>
          <option value="SFALCIO COMPLETO">Sfalcio completo</option>
          <option value="SFALCIO PARZIALE">Sfalcio parziale</option>
          <option value="RIFINITURA BORDI">Rifinitura bordi</option>
          <option value="RACCOLTA RESIDUI">Raccolta residui</option>
        </select></label>
        <label><span>Lavorazione richiesta *</span><textarea name="requestedWork" rows="3" required>Sfalcio completo del parco</textarea></label>
        <label><span>Superficie prevista (mq, facoltativa)</span><input name="areaMq" type="number" min="0" step="0.01" inputmode="decimal"></label>
        <label><span>Nota operatore</span><textarea name="operatorNote" rows="3" placeholder="Informazioni utili per la squadra"></textarea></label>
        <p class="cobo-mowing-feedback" role="status" aria-live="polite"></p>
        <button class="btn btn-primary cobo-mowing-save" type="submit">CREA CANTIERE IN SFALCIO COBO</button>
      </form>`;
    const workType = content.querySelector("[name='workType']");
    const requestedWork = content.querySelector("[name='requestedWork']");
    workType?.addEventListener("change", () => {
      if (requestedWork?.dataset.userEdited === "1") return;
      requestedWork.value = `${workType.options[workType.selectedIndex].text} del parco`;
    });
    requestedWork?.addEventListener("input", () => { requestedWork.dataset.userEdited = "1"; });
    content.querySelector("form")?.addEventListener("submit", async (event) => {
      event.preventDefault();
      const button = event.currentTarget.querySelector(".cobo-mowing-save");
      const feedback = event.currentTarget.querySelector(".cobo-mowing-feedback");
      const form = new FormData(event.currentTarget);
      button.disabled = true;
      feedback.className = "cobo-mowing-feedback";
      feedback.textContent = "Salvataggio del cantiere in corso…";
      try {
        const saved = await saveWorkOrder(context, {
          workType: text(form.get("workType")),
          requestedWork: text(form.get("requestedWork")),
          areaMq: text(form.get("areaMq")),
          operatorNote: text(form.get("operatorNote"))
        });
        feedback.classList.add("success");
        const resultText = saved.created
          ? "✅ Parco trasformato in cantiere."
          : saved.updated
            ? "✅ Cantiere Sfalcio COBO aggiornato."
            : "✅ Il cantiere era già aggiornato.";
        feedback.innerHTML = `${resultText}<br><span>${saved.alreadyDone ? "Il cantiere è già presente nei FATTI." : "Usa REGISTRA SFALCIO e poi il FATTO originale nella commessa."}</span>`;
        const openButton = document.createElement("button");
        openButton.type = "button";
        openButton.className = "btn cobo-mowing-open-created";
        openButton.textContent = "APRI IL CANTIERE";
        openButton.addEventListener("click", () => openCreatedCantiere(saved));
        feedback.appendChild(openButton);
      } catch (error) {
        feedback.classList.add("error");
        feedback.textContent = error?.message || "Impossibile creare il cantiere Sfalcio COBO.";
      } finally {
        button.disabled = false;
      }
    });
    content.querySelector("select, textarea, input")?.focus?.();
  }

  function isWorkOrder(plant, commessaId = "") {
    const sameCommessa = !commessaId || text(commessaId) === COMMESSA_ID;
    return sameCommessa && (plant?.sfalcioCobo === true || text(plant?.commessaId) === COMMESSA_ID);
  }

  function registrationSummary(plant = {}) {
    return {
      registered: plant.sfalcioCoboRegistrato === true,
      workType: text(plant.sfalcioCoboTipoEsecuzione),
      areaMq: parseOptionalNumber(plant.sfalcioCoboSuperficieEseguitaMq),
      equipment: text(plant.sfalcioCoboMezzo),
      note: text(plant.sfalcioCoboNota),
      operator: text(plant.sfalcioCoboRegistratoDa)
    };
  }

  function registrationButtonLabel(plant = {}) {
    return plant.sfalcioCoboRegistrato === true ? "✏️ MODIFICA SFALCIO" : "🧾 REGISTRA SFALCIO";
  }

  async function saveRegistration(options, values) {
    const store = options?.store || database();
    const plant = options?.plant || {};
    const commessaId = text(options?.commessaId || COMMESSA_ID);
    const plantId = text(plant.id);
    if (!store?.collection) throw new Error("Database non disponibile. Controlla la connessione e riprova.");
    if (commessaId !== COMMESSA_ID || !isWorkOrder(plant, commessaId)) throw new Error("Questo modulo è disponibile solo per Sfalcio COBO.");
    if (!plantId) throw new Error("Cantiere non valido: identificativo mancante.");

    const areaMq = parseOptionalNumber(values.areaMq);
    const workType = text(values.workType);
    const equipment = text(values.equipment);
    const operatorNote = text(values.note);
    const registrationNote = [
      workType ? `Sfalcio: ${workType}` : "",
      equipment ? `Mezzo: ${equipment}` : "",
      areaMq == null ? "" : `Superficie: ${areaMq} mq`,
      operatorNote ? `Nota: ${operatorNote}` : ""
    ].filter(Boolean).join(" · ");
    const comparable = {
      sfalcioCoboRegistrato: true,
      sfalcioCoboTipoEsecuzione: workType,
      sfalcioCoboSuperficieEseguitaMq: areaMq,
      sfalcioCoboMezzo: equipment,
      sfalcioCoboNota: operatorNote,
      note: registrationNote,
      noteImpianto: registrationNote,
      noteOperatore: operatorNote,
      hasNote: Boolean(registrationNote),
      sfalcioCoboRegistratoDa: text(options.operatorName || getOperatorName(authenticatedUser())),
      sfalcioCoboRegistratoDaUid: text(options.operatorUid || authenticatedUser()?.uid),
      sfalcioCoboRegistrazioneVersione: 1
    };
    if (!hasPayloadChanges(plant, comparable)) return { changed: false, updates: comparable };

    const timestampFactory = typeof options.timestampFactory === "function" ? options.timestampFactory : serverTimestamp;
    const updates = { ...comparable, sfalcioCoboRegistratoAt: timestampFactory(), updatedAt: timestampFactory() };
    await store.collection(options.collectionName || collectionName()).doc(COMMESSA_ID).collection("impianti").doc(plantId).set(updates, { merge: true });
    return { changed: true, updates: comparable };
  }

  function openRegistration(options = {}) {
    const plant = options.plant || {};
    if (!isWorkOrder(plant, options.commessaId)) return;
    returnFocusTarget = options.triggerButton || document.activeElement;
    const summary = registrationSummary(plant);
    const modal = createModal("Registra lo sfalcio", `${COMMESSA_NAME} · preparazione prima di FATTO`);
    const content = modal.querySelector(".cobo-mowing-content");
    content.innerHTML = `
      <section class="cobo-mowing-park-summary">
        <strong>${esc(plant.denominazione || plant.nome || "Parco / giardino")}</strong>
        <span>${esc(plant.parcoQuartiere || plant.area || "Bologna")}</span>
      </section>
      <p class="cobo-mowing-hint">Registra i dati della lavorazione. Lo spostamento nei FATTI e Whazzup rimangono affidati al pulsante FATTO originale.</p>
      <form class="cobo-mowing-form">
        <label><span>Lavorazione eseguita *</span><select name="workType" required>
          <option value="SFALCIO COMPLETO"${summary.workType === "SFALCIO COMPLETO" ? " selected" : ""}>Sfalcio completo</option>
          <option value="SFALCIO PARZIALE"${summary.workType === "SFALCIO PARZIALE" ? " selected" : ""}>Sfalcio parziale</option>
          <option value="RIFINITURA BORDI"${summary.workType === "RIFINITURA BORDI" ? " selected" : ""}>Rifinitura bordi</option>
          <option value="RACCOLTA RESIDUI"${summary.workType === "RACCOLTA RESIDUI" ? " selected" : ""}>Raccolta residui</option>
        </select></label>
        <label><span>Superficie eseguita (mq, facoltativa)</span><input name="areaMq" type="number" min="0" step="0.01" inputmode="decimal" value="${summary.areaMq == null ? "" : esc(summary.areaMq)}"></label>
        <label><span>Mezzo utilizzato</span><select name="equipment">
          <option value="">Non indicato</option>
          ${["DECESPUGLIATORE", "RASAERBA", "TRINCIA", "MANUALE", "MISTO"].map((value) => `<option value="${value}"${summary.equipment === value ? " selected" : ""}>${value.charAt(0) + value.slice(1).toLowerCase()}</option>`).join("")}
        </select></label>
        <label><span>Nota operatore</span><textarea name="note" rows="4" placeholder="Lavorazioni eseguite, problemi o parti rimaste">${esc(summary.note || plant.noteOperatore || "")}</textarea></label>
        <p class="cobo-mowing-feedback" role="status" aria-live="polite"></p>
        <button class="btn btn-primary cobo-mowing-save" type="submit">SALVA REGISTRAZIONE SFALCIO</button>
      </form>`;
    content.querySelector("form")?.addEventListener("submit", async (event) => {
      event.preventDefault();
      const button = event.currentTarget.querySelector(".cobo-mowing-save");
      const feedback = event.currentTarget.querySelector(".cobo-mowing-feedback");
      const form = new FormData(event.currentTarget);
      button.disabled = true;
      feedback.className = "cobo-mowing-feedback";
      feedback.textContent = "Salvataggio della registrazione…";
      try {
        const result = await saveRegistration(options, {
          workType: text(form.get("workType")),
          areaMq: text(form.get("areaMq")),
          equipment: text(form.get("equipment")),
          note: text(form.get("note"))
        });
        Object.assign(plant, result.updates);
        feedback.classList.add("success");
        feedback.textContent = result.changed
          ? "✅ Sfalcio registrato. Ora puoi usare il FATTO originale."
          : "✅ I dati erano già aggiornati. Ora puoi usare il FATTO originale.";
        options.onComplete?.(result.updates);
      } catch (error) {
        feedback.classList.add("error");
        feedback.textContent = error?.message || "Impossibile registrare lo sfalcio.";
      } finally {
        button.disabled = false;
      }
    });
    content.querySelector("select, textarea, input")?.focus?.();
  }

  window.HeraCoboMowing = Object.freeze({
    installed: true,
    version: "1.0.0",
    commessaId: COMMESSA_ID,
    commessaName: COMMESSA_NAME,
    commessaCode: COMMESSA_CODE,
    sourceDataset: SOURCE_DATASET,
    openCreate,
    openRegistration,
    close: closeModal,
    isWorkOrder,
    registrationSummary,
    registrationButtonLabel,
    buildPlantPayload,
    plantIdFor
  });
})();
