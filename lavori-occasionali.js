(() => {
  "use strict";

  if (window.HeraLavoriOccasionali?.installed) return;

  const COMMESSA_ID = "lavori-occasionali";
  const COMMESSA_NOME = "LAVORI OCCASIONALI";
  const state = {
    installed: false,
    rowsWrapped: false,
    virtualCommessaReady: false,
    selectorObserved: false,
    lastError: null
  };
  let selectorObserver = null;
  let cardsObserver = null;
  let pendingPreventivoFile = null;
  let pendingOccasionalSharePlant = null;
  let preventivoViewer = null;
  const leafletMaps = new Set();
  const mapLayers = new WeakMap();

  const normalizeName = (value) => String(value || "")
    .trim()
    .replace(/\s+/g, " ")
    .toLocaleUpperCase("it-IT");

  function isOccasionalSelected() {
    return document.getElementById("squadra-commessa")?.value === COMMESSA_ID;
  }

  function getWorkName() {
    const editor = document.getElementById("lavoro-occasionale-nome");
    return normalizeName(editor?.textContent || editor?.value);
  }

  function getEditorText(id) {
    const editor = document.getElementById(id);
    return String(editor?.textContent || editor?.value || "").trim();
  }

  function parseCoordinates(value) {
    const matches = String(value || "").replace(/;/g, ",").match(/-?\d+(?:[.,]\d+)?/g) || [];
    if (matches.length < 2) return null;
    const lat = Number(matches[0].replace(",", "."));
    const lng = Number(matches[1].replace(",", "."));
    if (!Number.isFinite(lat) || !Number.isFinite(lng) || Math.abs(lat) > 90 || Math.abs(lng) > 180) return null;
    return { lat, lng, text: `${lat.toFixed(6)}, ${lng.toFixed(6)}` };
  }

  function getWorkMetadata() {
    const coordinates = parseCoordinates(getEditorText("lavoro-occasionale-coordinate"));
    return {
      nome: getWorkName(),
      descrizione: getEditorText("lavoro-occasionale-descrizione"),
      comune: getEditorText("lavoro-occasionale-comune"),
      indirizzo: getEditorText("lavoro-occasionale-indirizzo"),
      codicePrezzo: getEditorText("lavoro-occasionale-codice-prezzo"),
      numeroPreventivo: getEditorText("lavoro-occasionale-numero-preventivo"),
      coordinates
    };
  }

  function stableHash(value) {
    let hash = 2166136261;
    const text = String(value || "");
    for (let index = 0; index < text.length; index += 1) {
      hash ^= text.charCodeAt(index);
      hash = Math.imul(hash, 16777619);
    }
    return (hash >>> 0).toString(36);
  }

  function occasionalPlantId(metadata) {
    const slug = metadata.nome.toLocaleLowerCase("it-IT")
      .normalize("NFD").replace(/[\u0300-\u036f]/g, "")
      .replace(/[^a-z0-9]+/g, "-").replace(/^-|-$/g, "").slice(0, 48) || "lavoro";
    return `occasionale-${slug}-${stableHash(`${metadata.nome}|${metadata.coordinates?.text || ""}`)}`;
  }

  async function savePreventivoFirestore(plantRef, file) {
    const base64 = await fileToBase64(file);
    const chunkSize = 700000;
    const uploadId = `${Date.now()}-${stableHash(file.name)}`;
    const chunks = [];
    for (let offset = 0; offset < base64.length; offset += chunkSize) chunks.push(base64.slice(offset, offset + chunkSize));
    for (let index = 0; index < chunks.length; index += 1) {
      await plantRef.collection("preventiviPdf").doc(`${uploadId}-${String(index).padStart(3, "0")}`).set({
        uploadId, index, data: chunks[index]
      });
    }
    return {
      preventivoPdfFirestore: true,
      preventivoPdfUploadId: uploadId,
      preventivoPdfChunks: chunks.length,
      preventivoPdfNome: file.name,
      preventivoPdfTipo: "application/pdf",
      preventivoPdfDimensione: file.size
    };
  }

  function occasionalPlantRef(plant) {
    const collectionName = typeof getCommesseCollectionName === "function" ? getCommesseCollectionName() : "commesse";
    const plantId = String(plant?.id || plant?.docId || "");
    return db.collection(collectionName).doc(COMMESSA_ID).collection("impianti").doc(plantId);
  }

  async function loadPreventivoBlob(plant) {
    if (plant?.preventivoPdfUrl) {
      const response = await fetch(String(plant.preventivoPdfUrl), { cache: "no-store" });
      if (!response.ok) throw new Error("Preventivo PDF non raggiungibile");
      return response.blob();
    }
    const uploadId = String(plant?.preventivoPdfUploadId || "");
    const count = Number(plant?.preventivoPdfChunks || 0);
    if (!uploadId || !count) throw new Error("Preventivo PDF non disponibile");
    const parts = [];
    for (let index = 0; index < count; index += 1) {
      const snapshot = await occasionalPlantRef(plant).collection("preventiviPdf")
        .doc(`${uploadId}-${String(index).padStart(3, "0")}`).get();
      if (!snapshot.exists) throw new Error("Una parte del preventivo PDF è mancante");
      parts.push(String(snapshot.data()?.data || ""));
    }
    const binary = atob(parts.join(""));
    const bytes = new Uint8Array(binary.length);
    for (let index = 0; index < binary.length; index += 1) bytes[index] = binary.charCodeAt(index);
    return new Blob([bytes], { type: "application/pdf" });
  }

  async function upsertNormalOccasionalPlant(metadata = getWorkMetadata()) {
    if (!metadata.nome || !metadata.coordinates) throw new Error("Nome o coordinate del lavoro occasionale mancanti.");
    if (typeof db === "undefined" || !db) throw new Error("Database non disponibile.");
    const collectionName = typeof getCommesseCollectionName === "function"
      ? getCommesseCollectionName()
      : "commesse";
    const plantId = occasionalPlantId(metadata);
    const now = typeof firebase !== "undefined" && firebase.firestore?.FieldValue?.serverTimestamp
      ? firebase.firestore.FieldValue.serverTimestamp()
      : new Date();
    const operatorName = typeof getOperatorDisplayName === "function" ? getOperatorDisplayName() : "";
    const userId = typeof currentUser !== "undefined" ? String(currentUser?.uid || "") : "";
    const plantRef = db.collection(collectionName).doc(COMMESSA_ID).collection("impianti").doc(plantId);
    const payload = {
      id: plantId,
      commessaId: COMMESSA_ID,
      idSap: "",
      denominazione: metadata.nome,
      nome: metadata.nome,
      comune: metadata.comune,
      indirizzo: metadata.indirizzo,
      descrizioneVia: metadata.indirizzo,
      latitudine: metadata.coordinates.lat,
      longitudine: metadata.coordinates.lng,
      gpsY: metadata.coordinates.lat,
      gpsX: metadata.coordinates.lng,
      coordinate: metadata.coordinates.text,
      tipologiaImpianto: "LAVORO OCCASIONALE",
      codicePrezzo: metadata.codicePrezzo,
      codiceVocePrezzo: metadata.codicePrezzo,
      numeroPreventivo: metadata.numeroPreventivo,
      tipologiaIntervento: metadata.descrizione,
      tipologiaLavorazione: metadata.descrizione,
      lavorazioniRichieste: metadata.descrizione,
      note: metadata.descrizione,
      lavoroOccasionale: true,
      updatedAt: now,
      updatedBy: userId,
      updatedByName: operatorName
    };
    await plantRef.set(payload, { merge: true });
    let preventivo = {};
    if (pendingPreventivoFile) {
      if (pendingPreventivoFile.type !== "application/pdf" && !pendingPreventivoFile.name.toLowerCase().endsWith(".pdf")) {
        throw new Error("Il preventivo deve essere un file PDF.");
      }
      if (pendingPreventivoFile.size > 20 * 1024 * 1024) throw new Error("Il PDF supera il limite di 20 MB.");
      if (typeof firebase === "undefined" || !firebase.storage) throw new Error("Archivio PDF non disponibile.");
      const safeName = pendingPreventivoFile.name.replace(/[^a-zA-Z0-9._-]/g, "_").slice(-120) || "preventivo.pdf";
      const storagePath = `lavori-occasionali/preventivi/${plantId}/${Date.now()}-${safeName}`;
      try {
        const storageRef = firebase.storage().ref(storagePath);
        const uploadTask = storageRef.put(pendingPreventivoFile, { contentType: "application/pdf" });
        let timeoutId;
        await Promise.race([
          uploadTask,
          new Promise((_, reject) => {
            timeoutId = window.setTimeout(() => {
              try { uploadTask.cancel(); } catch (_) {}
              reject(new Error("Firebase Storage non raggiungibile: uso archivio alternativo."));
            }, 15000);
          })
        ]).finally(() => window.clearTimeout(timeoutId));
        preventivo = {
          preventivoPdfUrl: await storageRef.getDownloadURL(),
          preventivoPdfNome: pendingPreventivoFile.name,
          preventivoPdfStoragePath: storagePath,
          preventivoPdfTipo: "application/pdf",
          preventivoPdfDimensione: pendingPreventivoFile.size
        };
      } catch (storageError) {
        state.lastError = storageError;
        preventivo = await savePreventivoFirestore(plantRef, pendingPreventivoFile);
      }
    }
    if (Object.keys(preventivo).length) await plantRef.set(preventivo, { merge: true });
    pendingPreventivoFile = null;
    return { plantId, metadata, ...preventivo };
  }

  function installVirtualCommessa() {
    const virtualCommessa = {
      id: COMMESSA_ID,
      nome: COMMESSA_NOME,
      codice: "OCCASIONALI",
      virtuale: true,
      lavoriOccasionali: true,
      attiva: true
    };

    try {
      if (typeof commesseById !== "undefined" && commesseById instanceof Map
        && !commesseById.has(COMMESSA_ID)) {
        commesseById.set(COMMESSA_ID, virtualCommessa);
      }
      if (typeof commesse !== "undefined" && Array.isArray(commesse)
        && !commesse.some((item) => String(item?.id || "") === COMMESSA_ID)) {
        commesse.push(virtualCommessa);
      }
    } catch (error) {
      state.lastError = error;
    }

    document.querySelectorAll("#squadra-commessa, #hours-table-commessa-select").forEach((select) => {
      if (select.querySelector(`option[value="${COMMESSA_ID}"]`)) return;
      const option = document.createElement("option");
      option.value = COMMESSA_ID;
      option.textContent = COMMESSA_NOME;
      select.appendChild(option);
    });

    state.virtualCommessaReady = Boolean(document.querySelector(`#squadra-commessa option[value="${COMMESSA_ID}"]`));
  }

  function observeCommessaSelector() {
    const select = document.getElementById("squadra-commessa");
    if (!select || selectorObserver) return;
    selectorObserver = new MutationObserver(() => {
      if (select.querySelector(`option[value="${COMMESSA_ID}"]`)) return;
      queueMicrotask(() => {
        installVirtualCommessa();
        installWorkField();
      });
    });
    selectorObserver.observe(select, { childList: true });
    state.selectorObserved = true;
  }

  function installWorkField() {
    const form = document.getElementById("squadra-form");
    const commessaSelect = document.getElementById("squadra-commessa");
    if (!form || !commessaSelect || document.getElementById("lavoro-occasionale-field")) return;

    const field = document.createElement("div");
    field.id = "lavoro-occasionale-field";
    field.className = "squadra-date-field lavoro-occasionale-field hidden";
    field.innerHTML = [
      "<span>Commessa o luogo del lavoro occasionale *</span>",
      '<div id="lavoro-occasionale-nome" class="lavoro-occasionale-editor" contenteditable="true"',
      ' role="textbox" aria-label="Commessa o luogo del lavoro occasionale"',
      ' data-placeholder="Es. Parco Zucca, Scuole Granarolo" spellcheck="true"></div>',
      '<small>Il nome verrà mostrato nella scheda della squadra.</small>',
      '<span class="lavoro-occasionale-subtitle">Descrizione del lavoro</span>',
      '<div id="lavoro-occasionale-descrizione" class="lavoro-occasionale-editor lavoro-occasionale-description"',
      ' contenteditable="true" role="textbox" aria-label="Descrizione del lavoro"',
      ' data-placeholder="Es. Sfalcio, raccolta, potatura..." spellcheck="true"></div>',
      '<span class="lavoro-occasionale-subtitle">Coordinate GPS *</span>',
      '<div id="lavoro-occasionale-coordinate" class="lavoro-occasionale-editor"',
      ' contenteditable="true" role="textbox" aria-label="Coordinate GPS"',
      ' data-placeholder="Es. 44.494887, 11.342616" inputmode="decimal"></div>',
      '<small>Con coordinate valide il lavoro comparirà anche sulla mappa.</small>',
      '<span class="lavoro-occasionale-subtitle">Comune</span>',
      '<div id="lavoro-occasionale-comune" class="lavoro-occasionale-editor" contenteditable="true"',
      ' role="textbox" data-placeholder="Es. Bologna"></div>',
      '<span class="lavoro-occasionale-subtitle">Indirizzo</span>',
      '<div id="lavoro-occasionale-indirizzo" class="lavoro-occasionale-editor" contenteditable="true"',
      ' role="textbox" data-placeholder="Es. Via del Frullo 5"></div>',
      '<div class="lavoro-occasionale-grid">',
      '<label><span class="lavoro-occasionale-subtitle">Codice prezzo</span>',
      '<div id="lavoro-occasionale-codice-prezzo" class="lavoro-occasionale-editor" contenteditable="true"',
      ' role="textbox" data-placeholder="Codice prezzo"></div></label>',
      '<label><span class="lavoro-occasionale-subtitle">Numero preventivo</span>',
      '<div id="lavoro-occasionale-numero-preventivo" class="lavoro-occasionale-editor" contenteditable="true"',
      ' role="textbox" data-placeholder="Numero preventivo"></div></label>',
      '</div>',
      '<label class="lavoro-occasionale-pdf-field"><span class="lavoro-occasionale-subtitle">Preventivo PDF</span>',
      '<input id="lavoro-occasionale-preventivo" type="file" accept="application/pdf,.pdf">',
      '<small id="lavoro-occasionale-preventivo-status">Nessun PDF selezionato.</small></label>',
      '<datalist id="lavori-occasionali-options"></datalist>'
    ].join("");

    const dateField = form.querySelector(".squadra-date-field");
    (dateField || commessaSelect).insertAdjacentElement("afterend", field);

    const input = field.querySelector("#lavoro-occasionale-nome");
    field.querySelectorAll("[contenteditable]").forEach((editor) => {
      ["keydown", "keyup", "keypress", "beforeinput", "input"].forEach((eventName) => {
        editor.addEventListener(eventName, (event) => event.stopPropagation());
      });
    });
    const pdfInput = field.querySelector("#lavoro-occasionale-preventivo");
    pdfInput?.addEventListener("change", () => {
      const file = pdfInput.files?.[0] || null;
      const status = field.querySelector("#lavoro-occasionale-preventivo-status");
      if (!file) {
        pendingPreventivoFile = null;
        if (status) status.textContent = "Nessun PDF selezionato.";
        return;
      }
      if (file.type !== "application/pdf" && !file.name.toLowerCase().endsWith(".pdf")) {
        pdfInput.value = "";
        pendingPreventivoFile = null;
        if (status) status.textContent = "Seleziona un file PDF valido.";
        return;
      }
      pendingPreventivoFile = file;
      if (status) status.textContent = `${file.name} • ${(file.size / 1048576).toLocaleString("it-IT", { maximumFractionDigits: 2 })} MB`;
    });
    input.addEventListener("input", () => {
      if (input.textContent.length > 120) input.textContent = input.textContent.slice(0, 120);
    });
    const refresh = () => {
      const active = isOccasionalSelected();
      field.classList.toggle("hidden", !active);
      input.setAttribute("aria-required", active ? "true" : "false");
      if (active) refreshSuggestions();
    };
    commessaSelect.addEventListener("change", refresh);
    refresh();
  }

  function getCompositions() {
    const output = [];
    try {
      if (!(squadreHistoryByDate instanceof Map)) return output;
      squadreHistoryByDate.forEach((byCommessa, dateKey) => {
        if (!(byCommessa instanceof Map)) return;
        const composition = byCommessa.get(COMMESSA_ID);
        if (!composition) return;
        const rows = Array.isArray(composition.squadre) ? composition.squadre : [];
        rows.forEach((row, index) => {
          const nome = normalizeName(row?.lavoroOccasionaleNome || composition?.lavoroOccasionaleNome);
          if (nome) output.push({ dateKey, nome, row, index });
        });
      });
    } catch (error) {
      state.lastError = error;
    }
    return output;
  }

  function refreshSuggestions() {
    const datalist = document.getElementById("lavori-occasionali-options");
    if (!datalist) return;
    const names = [...new Set(getCompositions().map((item) => item.nome))].sort((a, b) => a.localeCompare(b, "it"));
    datalist.replaceChildren(...names.map((name) => {
      const option = document.createElement("option");
      option.value = name;
      return option;
    }));
  }

  function wrapRowReader() {
    if (state.rowsWrapped || typeof readSquadraRows !== "function") return;
    const original = readSquadraRows;
    readSquadraRows = function readSquadraRowsWithOccasionalWork() {
      const rows = original.apply(this, arguments);
      if (!isOccasionalSelected()) return rows;
      const metadata = getWorkMetadata();
      return Array.isArray(rows) ? rows.map((row) => ({
        ...row,
        lavoroOccasionale: true,
        lavoroOccasionaleNome: metadata.nome,
        lavoroOccasionaleDescrizione: metadata.descrizione,
        lavoroOccasionaleComune: metadata.comune,
        lavoroOccasionaleIndirizzo: metadata.indirizzo,
        lavoroOccasionaleCodicePrezzo: metadata.codicePrezzo,
        lavoroOccasionaleNumeroPreventivo: metadata.numeroPreventivo,
        lavoroOccasionaleCoordinate: metadata.coordinates?.text || "",
        lavoroOccasionaleLat: metadata.coordinates?.lat ?? null,
        lavoroOccasionaleLng: metadata.coordinates?.lng ?? null
      })) : rows;
    };
    state.rowsWrapped = true;
  }

  function restoreWorkNameFromComposition() {
    if (!isOccasionalSelected()) return;
    const dateKey = document.getElementById("squadra-riferimento")?.value || "";
    try {
      const composition = squadreHistoryByDate instanceof Map
        ? squadreHistoryByDate.get(dateKey)?.get(COMMESSA_ID)
        : null;
      const first = Array.isArray(composition?.squadre) ? composition.squadre[0] : null;
      const nome = normalizeName(first?.lavoroOccasionaleNome || composition?.lavoroOccasionaleNome);
      const input = document.getElementById("lavoro-occasionale-nome");
      if (input && !normalizeName(input.textContent)) input.textContent = nome;
      const description = document.getElementById("lavoro-occasionale-descrizione");
      if (description && !description.textContent.trim()) {
        description.textContent = String(first?.lavoroOccasionaleDescrizione || "").trim();
      }
      const coordinate = document.getElementById("lavoro-occasionale-coordinate");
      if (coordinate && !coordinate.textContent.trim()) {
        coordinate.textContent = String(first?.lavoroOccasionaleCoordinate || "").trim();
      }
      const restore = (id, value) => {
        const editor = document.getElementById(id);
        if (editor && !editor.textContent.trim()) editor.textContent = String(value || "").trim();
      };
      restore("lavoro-occasionale-comune", first?.lavoroOccasionaleComune);
      restore("lavoro-occasionale-indirizzo", first?.lavoroOccasionaleIndirizzo);
      restore("lavoro-occasionale-codice-prezzo", first?.lavoroOccasionaleCodicePrezzo);
      restore("lavoro-occasionale-numero-preventivo", first?.lavoroOccasionaleNumeroPreventivo);
      if (composition && typeof setSquadraRowsFromData === "function") {
        window.setTimeout(() => {
          try {
            setSquadraRowsFromData(composition);
            if (typeof updateSquadraAutofillHint === "function") {
              updateSquadraAutofillHint(`Composizione salvata per ${nome || "questo lavoro occasionale"}.`);
            }
          } catch (error) {
            state.lastError = error;
          }
        }, 0);
      }
    } catch (error) {
      state.lastError = error;
    }
  }

  function validateBeforeCoreSave(event) {
    if (!isOccasionalSelected()) return;
    const rows = typeof readSquadraRows === "function" ? readSquadraRows() : [];
    if (!Array.isArray(rows) || !rows.length) return;
    const input = document.getElementById("lavoro-occasionale-nome");
    const metadata = getWorkMetadata();
    const nome = metadata.nome;
    if (nome) {
      const coordinateText = getEditorText("lavoro-occasionale-coordinate");
      if (!coordinateText || !metadata.coordinates) {
        event.preventDefault();
        event.stopImmediatePropagation();
        document.getElementById("lavoro-occasionale-coordinate")?.focus();
        const feedback = document.getElementById("squadra-feedback");
        if (feedback) {
          feedback.dataset.type = "error";
          feedback.textContent = "Inserisci coordinate valide nel formato: 44.494887, 11.342616";
        }
        return;
      }
      input.textContent = nome;
      return;
    }
    event.preventDefault();
    event.stopImmediatePropagation();
    input?.focus();
    const feedback = document.getElementById("squadra-feedback");
    if (feedback) {
      feedback.dataset.type = "error";
      feedback.textContent = "Inserisci il nome della commessa o del luogo, per esempio Parco Zucca.";
    }
  }

  function countOperators(row) {
    const candidates = [row?.operatori, row?.componenti, row?.persone, row?.personale];
    const value = candidates.find(Array.isArray);
    return value?.length || 0;
  }

  function getHoursFor(dateKey) {
    try {
      if (!Array.isArray(oreReports)) return 0;
      return oreReports.reduce((total, report) => {
        const reportDate = String(report?.dateKey || report?.data || report?.date || "").slice(0, 10);
        const commessaId = String(report?.commessaId || report?.commessa?.id || "");
        if (reportDate !== dateKey || commessaId !== COMMESSA_ID) return total;
        const raw = report?.ore ?? report?.hours ?? report?.totaleOre ?? 0;
        const numeric = Number(String(raw).replace(",", "."));
        return total + (Number.isFinite(numeric) ? numeric : 0);
      }, 0);
    } catch (_) {
      return 0;
    }
  }

  function renderHistory() {
    const host = document.getElementById("lavori-occasionali-history-list");
    if (!host) return;
    const groups = new Map();
    getCompositions().forEach((item) => {
      const current = groups.get(item.nome) || { nome: item.nome, dates: new Set(), operators: 0, hours: 0 };
      current.dates.add(item.dateKey);
      current.operators += countOperators(item.row);
      groups.set(item.nome, current);
    });
    groups.forEach((group) => {
      group.hours = [...group.dates].reduce((sum, dateKey) => sum + getHoursFor(dateKey), 0);
    });
    const rows = [...groups.values()].sort((a, b) => a.nome.localeCompare(b.nome, "it"));
    if (!rows.length) {
      host.innerHTML = '<p class="muted">Nessun lavoro occasionale registrato.</p>';
      return;
    }
    host.replaceChildren(...rows.map((group) => {
      const article = document.createElement("article");
      article.className = "item-card lavoro-occasionale-history-item";
      const hours = group.hours ? ` • ${group.hours.toLocaleString("it-IT")} ore registrate` : "";
      article.innerHTML = `<strong>${group.nome}</strong><p>${group.dates.size} intervent${group.dates.size === 1 ? "o" : "i"}${hours}</p>`;
      return article;
    }));
  }

  function installHistory() {
    const panel = document.getElementById("panel-squadre");
    if (!panel || document.getElementById("lavori-occasionali-history")) return;
    const section = document.createElement("section");
    section.id = "lavori-occasionali-history";
    section.className = "squadra-calendar-box lavori-occasionali-history";
    section.innerHTML = '<h3>Storico lavori occasionali</h3><p class="muted">Interventi raggruppati per commessa o luogo.</p><div id="lavori-occasionali-history-list"></div>';
    panel.appendChild(section);
    renderHistory();
  }

  function getOccasionalPlants() {
    const items = [];
    try {
      if (Array.isArray(currentImpianti)) items.push(...currentImpianti);
    } catch (_) {}
    try {
      const cached = impiantiByCommessaId instanceof Map ? impiantiByCommessaId.get(COMMESSA_ID) : null;
      if (Array.isArray(cached)) items.push(...cached);
    } catch (_) {}
    return [...new Map(items.filter((item) => item?.lavoroOccasionale === true)
      .map((item) => [String(item.id || item.docId || item.denominazione), item])).values()];
  }

  function closePreventivoViewer() {
    preventivoViewer?.remove();
    preventivoViewer = null;
    document.body.style.overflow = "";
  }

  async function openPreventivoViewer(plant) {
    if (!plant?.preventivoPdfUrl && !plant?.preventivoPdfFirestore) return;
    closePreventivoViewer();
    let url = String(plant?.preventivoPdfUrl || "").trim();
    let temporaryUrl = "";
    if (!url) {
      const blob = await loadPreventivoBlob(plant);
      temporaryUrl = URL.createObjectURL(blob);
      url = temporaryUrl;
    }
    preventivoViewer = document.createElement("section");
    preventivoViewer.className = "preventivo-pdf-fullscreen";
    preventivoViewer.setAttribute("role", "dialog");
    preventivoViewer.setAttribute("aria-modal", "true");
    preventivoViewer.innerHTML = `
      <header class="preventivo-pdf-toolbar"><strong>📄 ${escapeHtml(plant.numeroPreventivo ? `Preventivo ${plant.numeroPreventivo}` : "Preventivo")}</strong>
      <div><a class="btn" href="${escapeHtml(url)}" target="_blank" rel="noopener" download>Scarica</a>
      <button class="btn btn-primary" type="button" data-share>Condividi</button>
      <button class="btn" type="button" data-close>✕</button></div></header>
      <iframe title="Preventivo PDF" src="${escapeHtml(url)}"></iframe>`;
    preventivoViewer.querySelector("[data-close]").addEventListener("click", () => {
      if (temporaryUrl) URL.revokeObjectURL(temporaryUrl);
      closePreventivoViewer();
    });
    preventivoViewer.querySelector("[data-share]").addEventListener("click", async () => {
      try {
        const blob = await loadPreventivoBlob(plant);
        const file = new File([blob], plant.preventivoPdfNome || "preventivo.pdf", { type: "application/pdf" });
        if (navigator.share && navigator.canShare?.({ files: [file] })) {
          await navigator.share({ title: "Preventivo", files: [file] });
        } else window.open(url, "_blank", "noopener");
      } catch (error) {
        window.alert("Non è stato possibile condividere il preventivo.");
      }
    });
    document.body.appendChild(preventivoViewer);
    document.body.style.overflow = "hidden";
    preventivoViewer.querySelector("[data-close]").focus();
  }

  function decorateOccasionalPlantCards() {
    getOccasionalPlants().forEach((plant) => {
      const name = normalizeName(plant.denominazione || plant.nome);
      document.querySelectorAll("button").forEach((fattoButton) => {
        if (normalizeName(fattoButton.textContent) !== "FATTO") return;
        let card = fattoButton.parentElement;
        while (card && card !== document.body) {
          const hasName = normalizeName(card.textContent).includes(name);
          const hasNavigate = [...card.querySelectorAll("button")].some((button) => normalizeName(button.textContent) === "NAVIGA");
          if (hasName && hasNavigate) break;
          card = card.parentElement;
        }
        if (!card || card === document.body || card.dataset.occasionalPlantDecorated === String(plant.id || plant.docId || name)) return;
        card.dataset.occasionalPlantDecorated = String(plant.id || plant.docId || name);
        card.querySelectorAll(".preventivo-number-row").forEach((row, index) => { if (index) row.remove(); });
        const labelRow = (label) => {
          const walker = document.createTreeWalker(card, NodeFilter.SHOW_TEXT);
          while (walker.nextNode()) {
            if (!normalizeName(walker.currentNode.nodeValue).startsWith(normalizeName(label))) continue;
            const parent = walker.currentNode.parentElement;
            const row = /^(STRONG|B)$/.test(parent?.tagName || "") ? parent.parentElement : parent;
            return row && row !== card ? row : null;
          }
          return null;
        };
        const codeRow = labelRow("Codice prezzo:");
        if (codeRow) codeRow.hidden = !String(plant.codicePrezzo || plant.codiceVocePrezzo || "").trim();
        let numberRow = card.querySelector(".preventivo-number-row");
        const numeroPreventivo = String(plant.numeroPreventivo || "").trim();
        if (!numeroPreventivo) numberRow?.remove();
        else if (!numberRow) {
          numberRow = document.createElement("div");
          numberRow.className = "preventivo-number-row";
          numberRow.innerHTML = `<strong>Numero preventivo:</strong> ${escapeHtml(numeroPreventivo)}`;
          const workRow = labelRow("Lavorazioni richieste:");
          (workRow || card.firstElementChild)?.insertAdjacentElement(workRow ? "beforebegin" : "afterend", numberRow);
        }
        if ((!plant.preventivoPdfUrl && !plant.preventivoPdfFirestore) || card.querySelector(".preventivo-open-btn")) return;
        const actions = fattoButton.parentElement;
        if (!actions) return;
        const button = document.createElement("button");
        button.type = "button";
        button.className = "btn preventivo-open-btn";
        button.textContent = "📄 VEDI PREVENTIVO";
        button.addEventListener("click", () => openPreventivoViewer(plant));
        actions.insertAdjacentElement("afterend", button);
      });
    });
  }

  function formatOccasionalDateTime(date = new Date()) {
    const parts = new Intl.DateTimeFormat("it-IT", {
      timeZone: "Europe/Rome", day: "2-digit", month: "2-digit", year: "numeric",
      hour: "2-digit", minute: "2-digit", hour12: false
    }).formatToParts(date).reduce((output, item) => ({ ...output, [item.type]: item.value }), {});
    return { date: `${parts.day}/${parts.month}/${parts.year}`, time: `${parts.hour}:${parts.minute}` };
  }

  function buildOccasionalWhatsAppMessage(plant) {
    const when = formatOccasionalDateTime();
    const operator = typeof getOperatorDisplayName === "function"
      ? getOperatorDisplayName()
      : String(typeof currentUser !== "undefined" ? currentUser?.displayName || currentUser?.email || "-" : "-");
    const lines = [
      "✅ Attività: INTERVENTO DI MANUTENZIONE VERDE"
    ];
    const numero = String(plant?.numeroPreventivo || "").trim();
    if (numero) lines.push(`🆔 Numero preventivo: ${numero}`);
    const codice = String(plant?.codicePrezzo || plant?.codiceVocePrezzo || "").trim();
    if (codice) lines.push(`🏷️ Codice prezzo: ${codice}`);
    lines.push(
      `🏗️ Cantiere: ${plant?.denominazione || plant?.nome || "-"}`,
      `📍 Comune: ${plant?.comune || "-"}`,
      `🛣️ Via: ${plant?.indirizzo || plant?.descrizioneVia || "-"}`,
      `🛠️ Lavorazione: ${plant?.lavorazioniRichieste || plant?.tipologiaLavorazione || plant?.tipologiaIntervento || plant?.note || "-"}`,
      `👷 Operatore: ${operator || "-"}`,
      `📅 Data: ${when.date}`,
      `🕒 Ora: ${when.time}`
    );
    return lines.join("\n");
  }

  function fileToBase64(file) {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = () => resolve(String(reader.result || "").split(",")[1] || "");
      reader.onerror = () => reject(reader.error || new Error("File non leggibile"));
      reader.readAsDataURL(file);
    });
  }

  async function sharePreventivoNative(plugin, plant) {
    const blob = await loadPreventivoBlob(plant);
    const file = new File([blob], plant.preventivoPdfNome || "preventivo.pdf", { type: "application/pdf" });
    const session = await plugin.begin();
    const sessionId = String(session?.sessionId || "");
    if (!sessionId) throw new Error("Sessione PDF non disponibile");
    try {
      await plugin.addDocument({ sessionId, fileName: file.name, mimeType: "application/pdf", data: await fileToBase64(file) });
      return await plugin.shareDocument({ sessionId });
    } catch (error) {
      try { await plugin.discard?.({ sessionId }); } catch (_) {}
      throw error;
    }
  }

  function installOccasionalFattoPdfFlow() {
    if (window.__heraOccasionalFattoPdfInstalled) return true;
    if (typeof safeOpenWhatsAppMessage !== "function") return false;
    const originalSafeOpen = safeOpenWhatsAppMessage;
    const schedulePdfShare = (result) => {
      const pending = pendingOccasionalSharePlant;
      if ((!pending?.plant?.preventivoPdfUrl && !pending?.plant?.preventivoPdfFirestore) || pending.pdfPromise || !result) return result;
      pending.messageOpened = true;
      pending.pdfPromise = new Promise((resolve) => window.setTimeout(resolve, 8000))
        .then(async () => {
          const plugin = typeof getDedicatedAndroidWhazzupPhotoPlugin === "function"
            ? getDedicatedAndroidWhazzupPhotoPlugin()
            : null;
          if (!plugin?.addDocument || !plugin?.shareDocument) throw new Error("Plugin PDF Android non disponibile");
          return sharePreventivoNative(plugin, pending.plant);
        })
        .then(() => new Promise((resolve) => window.setTimeout(resolve, 8000)))
        .catch((error) => {
          state.lastError = error;
          if (typeof showToast === "function") showToast(`Preventivo non allegato: ${error?.message || error}`);
        });
      return result;
    };
    safeOpenWhatsAppMessage = function safeOpenWhatsAppMessageWithOccasionalPdf(message) {
      const pending = pendingOccasionalSharePlant;
      const outgoingMessage = pending?.plant ? buildOccasionalWhatsAppMessage(pending.plant) : message;
      const result = originalSafeOpen.call(this, outgoingMessage);
      return schedulePdfShare(result);
    };

    const installOpenWhatsAppBridge = () => {
      const originalOpen = window.openWhatsApp;
      if (typeof originalOpen !== "function" || originalOpen.__occasionalMessageWrapped) return false;
      const wrappedOpen = function openWhatsAppWithOccasionalMessage(value) {
        const pending = pendingOccasionalSharePlant;
        let nextValue = value;
        if (pending?.plant) {
          const customMessage = buildOccasionalWhatsAppMessage(pending.plant);
          try {
            const parsed = new URL(String(value || ""));
            parsed.searchParams.set("text", customMessage);
            nextValue = parsed.toString();
          } catch (_) {
            nextValue = `whatsapp://send?text=${encodeURIComponent(customMessage)}`;
          }
        }
        return schedulePdfShare(originalOpen.call(this, nextValue));
      };
      wrappedOpen.__occasionalMessageWrapped = true;
      wrappedOpen.__original = originalOpen;
      window.openWhatsApp = wrappedOpen;
      return true;
    };
    installOpenWhatsAppBridge();

    const installPhotoBridge = () => {
      if (typeof shareWhazzupPhotosNativeAndroid !== "function" || shareWhazzupPhotosNativeAndroid.__occasionalPdfWrapped) return false;
      const originalPhotoShare = shareWhazzupPhotosNativeAndroid;
      const wrapped = async function shareWhazzupPhotosAfterOccasionalPdf(files, message) {
        const pending = pendingOccasionalSharePlant;
        if (!pending?.plant?.preventivoPdfUrl && !pending?.plant?.preventivoPdfFirestore) return originalPhotoShare.apply(this, arguments);
        if (!pending.messageOpened) safeOpenWhatsAppMessage(message);
        await pending.pdfPromise;
        const plugin = typeof getDedicatedAndroidWhazzupPhotoPlugin === "function" ? getDedicatedAndroidWhazzupPhotoPlugin() : null;
        try {
          if (plugin && typeof shareWhazzupPhotosDedicatedAndroid === "function") {
            return await shareWhazzupPhotosDedicatedAndroid(plugin, files);
          }
          return await originalPhotoShare(files, "");
        } finally {
          pendingOccasionalSharePlant = null;
        }
      };
      wrapped.__occasionalPdfWrapped = true;
      shareWhazzupPhotosNativeAndroid = wrapped;
      return true;
    };
    installPhotoBridge();
    [500, 1500, 4000].forEach((delay) => window.setTimeout(() => {
      installOpenWhatsAppBridge();
      installPhotoBridge();
    }, delay));
    window.__heraOccasionalFattoPdfInstalled = true;
    return true;
  }

  function captureOccasionalFatto(event) {
    const button = event.target.closest("button");
    if (!button || normalizeName(button.textContent) !== "FATTO") return;
    try {
      if (String(selectedCommessaId || "") !== COMMESSA_ID) return;
    } catch (_) { return; }
    let card = button.parentElement;
    let plant = null;
    const plants = getOccasionalPlants();
    while (card && card !== document.body && !plant) {
      const cardText = normalizeName(card.textContent);
      plant = plants.find((item) => cardText.includes(normalizeName(item.denominazione || item.nome))) || null;
      if (!plant) card = card.parentElement;
    }
    if (!plant) return;
    pendingOccasionalSharePlant = { plant, messageOpened: false, pdfPromise: null };
    window.setTimeout(() => {
      if (pendingOccasionalSharePlant?.plant === plant) pendingOccasionalSharePlant = null;
    }, 90000);
  }

  function getLatestWork(dateKey = "") {
    const items = getCompositions()
      .filter((item) => !dateKey || item.dateKey === dateKey)
      .sort((a, b) => String(b.dateKey).localeCompare(String(a.dateKey)));
    return items[0] || null;
  }

  function applyWorkNamesToData() {
    const items = getCompositions();
    items.forEach((item) => {
      try {
        const composition = squadreHistoryByDate.get(item.dateKey)?.get(COMMESSA_ID);
        if (!composition) return;
        composition.commessaNome = item.nome;
        composition.lavoroOccasionaleNome = item.nome;
        composition.lavoroOccasionaleDescrizione = item.row?.lavoroOccasionaleDescrizione || "";
      } catch (error) {
        state.lastError = error;
      }
    });
  }

  function dateFromCardText(text) {
    const match = String(text || "").match(/(\d{2})\/(\d{2})\/(\d{4})/);
    return match ? `${match[3]}-${match[2]}-${match[1]}` : "";
  }

  function decorateSquadCards() {
    applyWorkNamesToData();
    const walker = document.createTreeWalker(document.body, NodeFilter.SHOW_TEXT);
    const matches = [];
    while (walker.nextNode()) {
      const node = walker.currentNode;
      const parent = node.parentElement;
      if (!parent || parent.closest("#squadra-form, select, option, script, style")) continue;
      if (normalizeName(node.nodeValue) === COMMESSA_NOME) matches.push(node);
    }
    matches.forEach((textNode) => {
      const title = textNode.parentElement;
      const card = title.closest("article, section, .card, [class*='commessa']") || title.parentElement;
      const work = getLatestWork(dateFromCardText(card?.textContent));
      if (!work?.nome) return;
      textNode.nodeValue = textNode.nodeValue.replace(/LAVORI\s+OCCASIONALI/i, work.nome);
      card?.querySelectorAll("span, small").forEach((badge) => {
        if (normalizeName(badge.textContent) === "OCCASIONALI") badge.textContent = "OCCASIONALE";
      });
    });
  }

  function escapeHtml(value) {
    return String(value || "").replace(/[&<>"']/g, (char) => ({
      "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;"
    }[char]));
  }

  function registerMap(map) {
    if (map && typeof map.eachLayer === "function" && typeof map.getContainer === "function") leafletMaps.add(map);
  }

  function discoverMaps() {
    ["map", "fullscreenMap", "commessaMap", "impiantiMap", "globalMap"].forEach((name) => {
      try { registerMap(window[name]); } catch (_) {}
    });
  }

  function syncMapMarkers() {
    if (typeof L === "undefined") return;
    discoverMaps();
    const points = getCompositions().filter((item) => {
      const lat = Number(item.row?.lavoroOccasionaleLat);
      const lng = Number(item.row?.lavoroOccasionaleLng);
      return Number.isFinite(lat) && Number.isFinite(lng);
    });
    leafletMaps.forEach((map) => {
      try {
        mapLayers.get(map)?.remove();
        const group = L.layerGroup();
        points.forEach((item) => {
          const marker = L.circleMarker(
            [Number(item.row.lavoroOccasionaleLat), Number(item.row.lavoroOccasionaleLng)],
            { radius: 9, color: "#b45309", weight: 3, fillColor: "#f59e0b", fillOpacity: 0.9 }
          );
          const description = item.row?.lavoroOccasionaleDescrizione
            ? `<br><span>${escapeHtml(item.row.lavoroOccasionaleDescrizione)}</span>`
            : "";
          marker.bindPopup(`<strong>${escapeHtml(item.nome)}</strong>${description}<br><small>${escapeHtml(item.dateKey)}</small>`);
          marker.addTo(group);
        });
        group.addTo(map);
        mapLayers.set(map, group);
      } catch (error) {
        state.lastError = error;
      }
    });
  }

  function installMapCapture() {
    try {
      if (typeof L === "undefined" || typeof L.map !== "function" || L.map.__heraOccasionalWrapped) return;
      const originalMap = L.map;
      const wrappedMap = function heraOccasionalLeafletMap() {
        const map = originalMap.apply(this, arguments);
        registerMap(map);
        window.setTimeout(syncMapMarkers, 0);
        return map;
      };
      Object.assign(wrappedMap, originalMap);
      wrappedMap.__heraOccasionalWrapped = true;
      L.map = wrappedMap;
    } catch (error) {
      state.lastError = error;
    }
  }

  function observeSquadCards() {
    if (cardsObserver || !document.body) return;
    let queued = false;
    cardsObserver = new MutationObserver(() => {
      if (queued) return;
      queued = true;
      queueMicrotask(() => {
        queued = false;
        installVirtualCommessa();
        installWorkField();
        decorateSquadCards();
        decorateOccasionalPlantCards();
      });
    });
    cardsObserver.observe(document.body, { childList: true, subtree: true });
  }

  function refresh() {
    installVirtualCommessa();
    observeCommessaSelector();
    installWorkField();
    wrapRowReader();
    installHistory();
    observeSquadCards();
    installOccasionalFattoPdfFlow();
    restoreWorkNameFromComposition();
    applyWorkNamesToData();
    refreshSuggestions();
    renderHistory();
    decorateSquadCards();
    decorateOccasionalPlantCards();
    state.installed = state.virtualCommessaReady && state.rowsWrapped;
  }

  document.addEventListener("DOMContentLoaded", refresh, { once: true });
  window.addEventListener("load", () => {
    refresh();
    window.setTimeout(refresh, 1200);
    window.setTimeout(refresh, 4000);
  }, { once: true });

  document.getElementById("squadra-form")?.addEventListener("submit", validateBeforeCoreSave, true);
  document.getElementById("squadra-form")?.addEventListener("submit", () => {
    const feedback = document.getElementById("squadra-feedback");
    if (!feedback || !isOccasionalSelected()) return;
    const rows = typeof readSquadraRows === "function" ? readSquadraRows() : [];
    const deletingComposition = !Array.isArray(rows) || !rows.length;
    const metadata = getWorkMetadata();
    const observer = new MutationObserver(async () => {
      if (feedback.dataset.type !== "success") return;
      observer.disconnect();
      if (!deletingComposition) {
        try {
          const saved = await upsertNormalOccasionalPlant(metadata);
          const pdfLoaded = Boolean(saved.preventivoPdfUrl || saved.preventivoPdfFirestore);
          const pdfInput = document.getElementById("lavoro-occasionale-preventivo");
          const pdfStatus = document.getElementById("lavoro-occasionale-preventivo-status");
          if (pdfLoaded) {
            if (pdfInput) pdfInput.value = "";
            if (pdfStatus) {
              pdfStatus.className = "preventivo-status-success";
              pdfStatus.textContent = `✅ PDF caricato: ${saved.preventivoPdfNome || "preventivo.pdf"}`;
            }
          }
          feedback.dataset.type = "success";
          feedback.textContent = `✅ Squadra e impianto ${metadata.nome} salvati.${pdfLoaded ? " ✅ Preventivo PDF caricato." : ""}`;
        } catch (error) {
          state.lastError = error;
          feedback.dataset.type = "error";
          const pdfStatus = document.getElementById("lavoro-occasionale-preventivo-status");
          if (pendingPreventivoFile && pdfStatus) {
            pdfStatus.className = "preventivo-status-error";
            pdfStatus.textContent = "❌ PDF non caricato. Il file è ancora selezionato: premi Fine per riprovare.";
          }
          feedback.textContent = `Squadra salvata, ma aggiornamento impianto non riuscito: ${error?.message || error}. Premi Fine per riprovare.`;
        }
      }
      window.setTimeout(() => {
        installVirtualCommessa();
        installWorkField();
        restoreWorkNameFromComposition();
        renderHistory();
        decorateSquadCards();
        decorateOccasionalPlantCards();
      }, 0);
      [150, 600].forEach((delay) => window.setTimeout(() => {
        installVirtualCommessa();
        installWorkField();
      }, delay));
    });
    observer.observe(feedback, { childList: true, subtree: true, attributes: true });
    window.setTimeout(() => observer.disconnect(), 20000);
  });
  document.getElementById("squadra-commessa")?.addEventListener("change", () => {
    restoreWorkNameFromComposition();
    renderHistory();
  });
  document.getElementById("squadra-riferimento")?.addEventListener("change", restoreWorkNameFromComposition);
  document.addEventListener("click", captureOccasionalFatto, true);

  const style = document.createElement("style");
  style.textContent = `
    .lavoro-occasionale-editor {
      box-sizing: border-box;
      width: 100%;
      min-height: 42px;
      padding: 10px 13px;
      border: 1px solid #cbd8ec;
      border-radius: 12px;
      background: #fff;
      color: #172033;
      font: inherit;
      line-height: 1.35;
      cursor: text;
      outline: none;
      white-space: pre-wrap;
      overflow-wrap: anywhere;
    }
    .lavoro-occasionale-editor:focus {
      border-color: #2563eb;
      box-shadow: 0 0 0 3px rgba(37, 99, 235, .14);
    }
    .lavoro-occasionale-editor:empty::before {
      content: attr(data-placeholder);
      color: #7b879c;
      pointer-events: none;
    }
    .lavoro-occasionale-description {
      min-height: 72px;
    }
    .lavoro-occasionale-subtitle {
      display: block;
      margin-top: 10px;
      margin-bottom: 5px;
      font-weight: 700;
    }
    .lavoro-occasionale-card-description {
      margin: 7px 0;
      color: #334155;
    }
    .lavoro-occasionale-grid {
      display: grid;
      grid-template-columns: repeat(2, minmax(0, 1fr));
      gap: 10px 14px;
    }
    .lavoro-occasionale-pdf-field {
      display: grid;
      gap: 6px;
      margin-top: 10px;
    }
    .lavoro-occasionale-pdf-field input[type="file"] {
      box-sizing: border-box;
      width: 100%;
      padding: 9px;
      border: 1px solid #cbd8ec;
      border-radius: 12px;
      background: #fff;
    }
    .preventivo-status-success,
    .preventivo-status-error {
      display: block;
      padding: 9px 11px;
      border-radius: 10px;
      font-weight: 800;
    }
    .preventivo-status-success {
      border: 1px solid #22c55e;
      background: #dcfce7;
      color: #166534;
    }
    .preventivo-status-error {
      border: 1px solid #ef4444;
      background: #fee2e2;
      color: #991b1b;
    }
    .preventivo-number-row { margin: 2px 0; }
    .preventivo-open-btn {
      width: 100%;
      margin-top: 8px;
      border-color: #d6a900 !important;
      background: #fff7c7 !important;
      color: #594500 !important;
      font-weight: 800;
    }
    .preventivo-pdf-fullscreen {
      position: fixed;
      inset: 0;
      z-index: 2147483000;
      display: grid;
      grid-template-rows: auto 1fr;
      background: #101827;
    }
    .preventivo-pdf-toolbar {
      display: flex;
      align-items: center;
      gap: 9px;
      padding: max(10px, env(safe-area-inset-top)) 12px 10px;
      background: #fff;
      box-shadow: 0 2px 12px rgba(0, 0, 0, .25);
    }
    .preventivo-pdf-toolbar strong {
      min-width: 0;
      flex: 1;
      overflow: hidden;
      text-overflow: ellipsis;
      white-space: nowrap;
    }
    .preventivo-pdf-toolbar button {
      min-height: 42px;
      padding: 8px 13px;
      border: 1px solid #cbd8ec;
      border-radius: 11px;
      background: #fff;
      font: inherit;
      font-weight: 800;
    }
    .preventivo-pdf-fullscreen iframe {
      width: 100%;
      height: 100%;
      border: 0;
      background: #fff;
    }
    @media (max-width: 680px) {
      .lavoro-occasionale-grid { grid-template-columns: 1fr; }
      .preventivo-pdf-toolbar { flex-wrap: wrap; }
      .preventivo-pdf-toolbar strong { flex-basis: 100%; order: -1; }
    }
  `;
  document.head.appendChild(style);

  window.HeraLavoriOccasionali = {
    installed: true,
    version: "1.3.0",
    commessaId: COMMESSA_ID,
    firestoreScope: "commesse/lavori-occasionali/impianti",
    refresh,
    getState: () => ({ ...state, lastError: state.lastError ? String(state.lastError?.message || state.lastError) : null })
  };
})();
