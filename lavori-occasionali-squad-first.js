(() => {
  "use strict";

  if (window.HeraOccasionalSquadFirstFlow?.installed) return;

  const COMMESSA_ID = "lavori-occasionali";
  const PANEL_ID = "lavori-occasionali-squad-first-panel";
  const FIELD_ID = "lavoro-occasionale-field";
  const state = {
    busy: false,
    observer: null,
    refreshQueued: false,
    lastError: null,
    lastSaved: null
  };

  const text = (value) => String(value ?? "").trim();
  const normalize = (value) => text(value)
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/\s+/g, " ")
    .toLocaleUpperCase("it-IT");

  function isOccasionalSelected() {
    return document.getElementById("squadra-commessa")?.value === COMMESSA_ID;
  }

  function currentDateKey() {
    const value = document.getElementById("squadra-riferimento")?.value
      || document.getElementById("squadre-filter-date")?.value
      || "";
    return /^\d{4}-\d{2}-\d{2}$/.test(value) ? value : "";
  }

  function currentComposition() {
    const dateKey = currentDateKey();
    if (!dateKey) return null;
    try {
      const history = typeof squadreHistoryByDate !== "undefined" ? squadreHistoryByDate : null;
      return history instanceof Map ? history.get(dateKey)?.get(COMMESSA_ID) || null : null;
    } catch (_) {
      return null;
    }
  }

  function compositionRows() {
    const composition = currentComposition();
    return Array.isArray(composition?.squadre) ? composition.squadre : [];
  }

  function personLabel(value) {
    if (value && typeof value === "object") {
      return text(
        value.nomeCompleto
        || value.nomeCognome
        || value.displayName
        || value.nome
        || value.label
        || value.email
        || value.id
      );
    }
    return text(value);
  }

  function rowOperators(row) {
    const values = [];
    const add = (value) => {
      if (Array.isArray(value)) {
        value.forEach(add);
        return;
      }
      const label = personLabel(value);
      if (!label) return;
      label.split(/[\n,;|]+/).map((item) => item.trim()).filter(Boolean).forEach((item) => values.push(item));
    };

    [
      row?.operatori,
      row?.componenti,
      row?.persone,
      row?.personale,
      row?.membri,
      row?.addetti,
      row?.operatoriNomi
    ].forEach(add);

    if (!values.length && row && typeof row === "object") {
      Object.entries(row).forEach(([key, value]) => {
        if (/^(?:operatore|addetto|componente|persona)\d*$/i.test(key)) add(value);
      });
    }

    const seen = new Set();
    return values.filter((value) => {
      const key = normalize(value);
      if (!key || seen.has(key)) return false;
      seen.add(key);
      return true;
    });
  }

  function squadraLabel(row, index) {
    const operators = rowOperators(row);
    return operators.length
      ? `Squadra ${index + 1} — ${operators.join(", ")}`
      : `Squadra ${index + 1}`;
  }

  function squadraKey(row, index) {
    const raw = text(
      row?.squadraId
      || row?.squadraID
      || row?.rowId
      || row?.uuid
      || row?.key
      || row?.id
      || `squadra-${index + 1}`
    );
    return raw
      .normalize("NFD")
      .replace(/[\u0300-\u036f]/g, "")
      .replace(/[^a-zA-Z0-9_-]+/g, "-")
      .replace(/^-+|-+$/g, "")
      .slice(0, 100) || `squadra-${index + 1}`;
  }

  function installStyle() {
    if (document.getElementById("lavori-occasionali-squad-first-style")) return;
    const style = document.createElement("style");
    style.id = "lavori-occasionali-squad-first-style";
    style.textContent = `
      .occasional-squad-first-panel{display:grid;gap:14px;margin-top:18px;padding:16px;border:1px solid #bfd2ef;border-radius:18px;background:#f8fbff}
      .occasional-squad-first-head{display:flex;align-items:flex-start;justify-content:space-between;gap:10px;flex-wrap:wrap}
      .occasional-squad-first-head h3{margin:0;color:#172033}
      .occasional-squad-first-step{display:inline-flex;align-items:center;min-height:28px;padding:4px 9px;border-radius:999px;background:#dbeafe;color:#1d4ed8;font-size:.78rem;font-weight:900}
      .occasional-squad-first-instructions{margin:0;color:#475569;font-weight:650;line-height:1.45}
      .occasional-squad-first-lock{padding:13px 14px;border:1px dashed #94a3b8;border-radius:14px;background:#fff;color:#334155;font-weight:800;line-height:1.4}
      .occasional-squad-first-editor{display:grid;gap:12px}
      .occasional-squad-first-team{display:grid;gap:6px;font-weight:850;color:#172033}
      .occasional-squad-first-team select{width:100%;min-height:48px;padding:9px 12px;border:1px solid #bfd2ef;border-radius:12px;background:#fff;color:#172033;font:inherit;font-weight:750}
      .occasional-squad-first-field-slot{display:grid;gap:10px}
      .occasional-squad-first-field-slot > #lavoro-occasionale-field{margin:0;padding:0;border:0;background:transparent}
      .occasional-squad-first-save{width:100%;min-height:52px;font-weight:900}
      .occasional-squad-first-feedback{min-height:22px;margin:0;font-weight:800;line-height:1.4}
      .occasional-squad-first-feedback[data-type="success"]{color:#166534}
      .occasional-squad-first-feedback[data-type="error"]{color:#b91c1c}
      .occasional-squad-first-feedback[data-type="info"]{color:#1d4ed8}
      #${PANEL_ID} .occasional-multi-controls{display:none!important}
      @media(max-width:600px){.occasional-squad-first-panel{padding:13px;border-radius:15px}.occasional-squad-first-step{width:100%;justify-content:center}}
    `;
    document.head.appendChild(style);
  }

  function ensurePanel() {
    let panel = document.getElementById(PANEL_ID);
    if (panel) return panel;
    const form = document.getElementById("squadra-form");
    if (!form) return null;

    panel = document.createElement("section");
    panel.id = PANEL_ID;
    panel.className = "occasional-squad-first-panel hidden";
    panel.innerHTML = `
      <header class="occasional-squad-first-head">
        <h3>📍 Cantieri della squadra</h3>
        <span class="occasional-squad-first-step">2° PASSAGGIO</span>
      </header>
      <p class="occasional-squad-first-instructions">Prima crea e salva normalmente la squadra. Dopo il salvataggio scegli la squadra e aggiungi uno o più cantieri.</p>
      <div class="occasional-squad-first-lock" data-occasional-squad-lock>Prima salva la squadra del giorno.</div>
      <div class="occasional-squad-first-editor hidden" data-occasional-squad-editor>
        <label class="occasional-squad-first-team">Squadra salvata
          <select id="lavoro-occasionale-squadra-select" aria-label="Seleziona la squadra a cui assegnare il cantiere"></select>
        </label>
        <div class="occasional-squad-first-field-slot" data-occasional-field-slot></div>
        <button id="lavoro-occasionale-assegna-btn" class="btn btn-primary occasional-squad-first-save" type="button">➕ AGGIUNGI CANTIERE ALLA SQUADRA</button>
        <p id="lavoro-occasionale-assegna-feedback" class="occasional-squad-first-feedback" data-type="info" role="status" aria-live="polite"></p>
      </div>
    `;
    form.insertAdjacentElement("afterend", panel);
    panel.querySelector("#lavoro-occasionale-assegna-btn")?.addEventListener("click", saveCantiereForSelectedSquad);
    panel.querySelector("#lavoro-occasionale-squadra-select")?.addEventListener("change", scheduleRefresh);
    return panel;
  }

  function adaptFieldCopy(field) {
    if (!field || field.dataset.squadFirstAdapted === "1") return;
    const firstLabel = Array.from(field.children).find((element) =>
      element.tagName === "SPAN" && /commessa o luogo/i.test(element.textContent || "")
    );
    if (firstLabel) firstLabel.textContent = "Nome del cantiere *";
    const nameEditor = field.querySelector("#lavoro-occasionale-nome");
    const hint = nameEditor?.nextElementSibling;
    if (hint?.tagName === "SMALL") hint.textContent = "Il cantiere verrà associato alla squadra selezionata.";
    field.dataset.squadFirstAdapted = "1";
  }

  function moveFieldToPanel(panel) {
    const field = document.getElementById(FIELD_ID);
    const slot = panel?.querySelector("[data-occasional-field-slot]");
    if (!field || !slot) return false;
    adaptFieldCopy(field);
    if (field.parentElement !== slot) slot.appendChild(field);
    return true;
  }

  function updateSelect(select, rows) {
    if (!select) return;
    const signature = rows.map((row, index) => `${squadraKey(row, index)}:${squadraLabel(row, index)}`).join("|");
    if (select.dataset.signature === signature) return;
    const previous = Number(select.value);
    select.replaceChildren(...rows.map((row, index) => {
      const option = document.createElement("option");
      option.value = String(index);
      option.textContent = squadraLabel(row, index);
      return option;
    }));
    select.dataset.signature = signature;
    select.value = Number.isInteger(previous) && previous >= 0 && previous < rows.length
      ? String(previous)
      : "0";
  }

  function setFeedback(type, message) {
    const feedback = document.getElementById("lavoro-occasionale-assegna-feedback");
    if (!feedback) return;
    feedback.dataset.type = type;
    if (feedback.textContent !== message) feedback.textContent = message;
  }

  function refreshPanel() {
    installStyle();
    const panel = ensurePanel();
    if (!panel) return;
    const active = isOccasionalSelected();
    panel.classList.toggle("hidden", !active);
    if (!active) return;

    const fieldReady = moveFieldToPanel(panel);
    const rows = compositionRows();
    const lock = panel.querySelector("[data-occasional-squad-lock]");
    const editor = panel.querySelector("[data-occasional-squad-editor]");
    const select = panel.querySelector("#lavoro-occasionale-squadra-select");
    const saveButton = panel.querySelector("#lavoro-occasionale-assegna-btn");
    const ready = Boolean(currentDateKey() && rows.length && fieldReady);

    updateSelect(select, rows);
    lock?.classList.toggle("hidden", ready);
    editor?.classList.toggle("hidden", !ready);
    if (lock && !ready) {
      const message = !currentDateKey()
        ? "Seleziona il giorno della composizione."
        : rows.length
          ? "Preparo il modulo cantieri…"
          : "Prima completa la squadra e premi FINE. Dopo il salvataggio potrai aggiungere i cantieri.";
      if (lock.textContent !== message) lock.textContent = message;
    }

    const field = document.getElementById(FIELD_ID);
    if (field) field.classList.toggle("hidden", !ready);
    if (saveButton) saveButton.disabled = !ready || state.busy;
  }

  function scheduleRefresh() {
    if (state.refreshQueued) return;
    state.refreshQueued = true;
    requestAnimationFrame(() => {
      state.refreshQueued = false;
      refreshPanel();
    });
  }

  function sanitizeForFirestore(value, seen = new WeakSet()) {
    if (value === null || value === undefined) return value === undefined ? null : value;
    if (["string", "number", "boolean"].includes(typeof value)) return value;
    if (typeof value === "function") return null;
    if (value instanceof Date) return value.toISOString();
    if (value && typeof value.toDate === "function") return value;
    if (Array.isArray(value)) return value.map((item) => sanitizeForFirestore(item, seen));
    if (typeof value !== "object") return text(value);
    if (seen.has(value)) return null;
    seen.add(value);
    const output = {};
    Object.entries(value).forEach(([key, item]) => {
      if (typeof item === "function" || item === undefined) return;
      output[key] = sanitizeForFirestore(item, seen);
    });
    seen.delete(value);
    return output;
  }

  function activeCollectionName() {
    try {
      return typeof getCommesseCollectionName === "function"
        ? text(getCommesseCollectionName()) || "commesse"
        : "commesse";
    } catch (_) {
      return "commesse";
    }
  }

  function serverNow() {
    try {
      return firebase.firestore.FieldValue.serverTimestamp();
    } catch (_) {
      return new Date();
    }
  }

  function currentUserData() {
    return {
      uid: text(typeof currentUser !== "undefined" ? currentUser?.uid : ""),
      name: typeof getOperatorDisplayName === "function"
        ? text(getOperatorDisplayName())
        : text(typeof currentUser !== "undefined" ? currentUser?.displayName || currentUser?.email : "")
    };
  }

  async function saveAssignment(saved, metadata, row, index) {
    if (typeof db === "undefined" || !db?.collection) throw new Error("Database non disponibile.");
    const dateKey = currentDateKey();
    const plantId = text(saved?.plantId);
    if (!dateKey || !plantId) throw new Error("Dati della squadra o del cantiere incompleti.");

    const teamKey = squadraKey(row, index);
    const safePlantId = plantId.replace(/[^a-zA-Z0-9_-]+/g, "-").slice(0, 220);
    const documentId = `${dateKey}_sq-${index + 1}_${teamKey}_${safePlantId}`.slice(0, 600);
    const user = currentUserData();
    const now = serverNow();
    const operators = rowOperators(row);

    await db.collection(activeCollectionName())
      .doc(COMMESSA_ID)
      .collection("assegnazioniOccasionali")
      .doc(documentId)
      .set({
        assignmentVersion: 2,
        dateKey,
        plantId,
        commessaId: COMMESSA_ID,
        cantiere: metadata.nome,
        comune: metadata.comune || "",
        indirizzo: metadata.indirizzo || "",
        coordinates: metadata.coordinates?.text || "",
        descrizione: metadata.descrizione || "",
        codicePrezzo: metadata.codicePrezzo || "",
        numeroPreventivo: metadata.numeroPreventivo || "",
        lavoroOccasionale: true,
        squadraIndex: index,
        squadraNumero: index + 1,
        squadraKey: teamKey,
        squadraLabel: squadraLabel(row, index),
        squadraOperatori: operators,
        squadra: [sanitizeForFirestore(row)],
        createdAt: now,
        updatedAt: now,
        createdBy: user.uid,
        createdByName: user.name,
        updatedBy: user.uid,
        updatedByName: user.name
      }, { merge: true });
  }

  function clearCantiereFields() {
    [
      "lavoro-occasionale-nome",
      "lavoro-occasionale-descrizione",
      "lavoro-occasionale-coordinate",
      "lavoro-occasionale-comune",
      "lavoro-occasionale-indirizzo",
      "lavoro-occasionale-codice-prezzo",
      "lavoro-occasionale-numero-preventivo"
    ].forEach((id) => {
      const editor = document.getElementById(id);
      if (!editor) return;
      if ("value" in editor && editor.tagName !== "DIV") editor.value = "";
      else editor.textContent = "";
    });
    const pdf = document.getElementById("lavoro-occasionale-preventivo");
    if (pdf) pdf.value = "";
    const pdfStatus = document.getElementById("lavoro-occasionale-preventivo-status");
    if (pdfStatus) {
      pdfStatus.className = "";
      pdfStatus.textContent = "Nessun PDF selezionato.";
    }
  }

  async function saveCantiereForSelectedSquad() {
    if (state.busy) return;
    const rows = compositionRows();
    const select = document.getElementById("lavoro-occasionale-squadra-select");
    const index = Number(select?.value || 0);
    const row = rows[index];
    if (!row) {
      setFeedback("error", "Prima salva la squadra e poi selezionala.");
      return;
    }

    const core = window.HeraLavoriOccasionali;
    if (typeof core?.getWorkMetadata !== "function" || typeof core?.upsertPlant !== "function") {
      setFeedback("error", "Il modulo cantieri non è ancora pronto. Attendi un momento e riprova.");
      window.setTimeout(scheduleRefresh, 500);
      return;
    }

    const metadata = core.getWorkMetadata();
    if (!metadata?.nome) {
      setFeedback("error", "Inserisci o seleziona il nome del cantiere.");
      document.getElementById("lavoro-occasionale-nome")?.focus();
      return;
    }
    if (!metadata?.coordinates) {
      setFeedback("error", "Inserisci coordinate valide oppure scegli il cantiere sulla mappa Google.");
      document.getElementById("lavoro-occasionale-coordinate")?.focus();
      return;
    }

    const button = document.getElementById("lavoro-occasionale-assegna-btn");
    state.busy = true;
    if (button) {
      button.disabled = true;
      button.textContent = "SALVATAGGIO CANTIERE…";
    }
    setFeedback("info", `Aggiungo ${metadata.nome} alla ${squadraLabel(row, index)}…`);

    try {
      const saved = await core.upsertPlant(metadata);
      await saveAssignment(saved, metadata, row, index);
      state.lastSaved = { dateKey: currentDateKey(), plantId: saved.plantId, squadraIndex: index, nome: metadata.nome };
      clearCantiereFields();
      setFeedback("success", `✅ ${metadata.nome} aggiunto alla ${squadraLabel(row, index)}. Puoi aggiungere subito un altro cantiere.`);
      try { window.HeraOccasionalSquadSites?.refresh?.(); } catch (_) {}
      window.dispatchEvent(new CustomEvent("hera:occasional-assignment-saved", { detail: { ...state.lastSaved } }));
      document.getElementById("lavoro-occasionale-nome")?.focus();
    } catch (error) {
      state.lastError = error;
      setFeedback("error", `Cantiere non assegnato: ${text(error?.message || error) || "errore sconosciuto"}. I dati restano compilati per riprovare.`);
    } finally {
      state.busy = false;
      if (button) {
        button.textContent = "➕ AGGIUNGI CANTIERE ALLA SQUADRA";
        button.disabled = false;
      }
      scheduleRefresh();
    }
  }

  function observeCoreFeedback() {
    const feedback = document.getElementById("squadra-feedback");
    if (!feedback || feedback.dataset.occasionalSquadFirstObserved === "1") return;
    feedback.dataset.occasionalSquadFirstObserved = "1";
    new MutationObserver(() => {
      if (!isOccasionalSelected() || feedback.dataset.type !== "success") return;
      setFeedback("info", "Squadra salvata. Ora seleziona la squadra e aggiungi uno o più cantieri qui sotto.");
      [0, 200, 700, 1500].forEach((delay) => window.setTimeout(scheduleRefresh, delay));
    }).observe(feedback, { childList: true, subtree: true, attributes: true, attributeFilter: ["data-type"] });
  }

  function installObserver() {
    if (state.observer) return;
    const root = document.getElementById("panel-squadre") || document.body;
    if (!root) return;
    state.observer = new MutationObserver(scheduleRefresh);
    state.observer.observe(root, { childList: true, subtree: true });
  }

  function start() {
    installStyle();
    ensurePanel();
    observeCoreFeedback();
    installObserver();
    document.getElementById("squadra-commessa")?.addEventListener("change", scheduleRefresh);
    document.getElementById("squadra-riferimento")?.addEventListener("change", scheduleRefresh);
    document.getElementById("squadra-form")?.addEventListener("submit", () => {
      [100, 500, 1200, 2500].forEach((delay) => window.setTimeout(scheduleRefresh, delay));
    });
    scheduleRefresh();
    [50, 250, 800, 2000, 5000].forEach((delay) => window.setTimeout(scheduleRefresh, delay));
  }

  window.HeraOccasionalSquadFirstFlow = {
    installed: true,
    version: "1.0.0",
    commessaId: COMMESSA_ID,
    refresh: scheduleRefresh,
    getState: () => ({
      busy: state.busy,
      lastSaved: state.lastSaved ? { ...state.lastSaved } : null,
      lastError: state.lastError ? text(state.lastError?.message || state.lastError) : null
    })
  };

  if (document.readyState === "loading") document.addEventListener("DOMContentLoaded", start, { once: true });
  else start();
})();
