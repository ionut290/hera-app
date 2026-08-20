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
      coordinates
    };
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

    const field = document.createElement("label");
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
      '<span class="lavoro-occasionale-subtitle">Coordinate GPS</span>',
      '<div id="lavoro-occasionale-coordinate" class="lavoro-occasionale-editor"',
      ' contenteditable="true" role="textbox" aria-label="Coordinate GPS"',
      ' data-placeholder="Es. 44.494887, 11.342616" inputmode="decimal"></div>',
      '<small>Con coordinate valide il lavoro comparirà anche sulla mappa.</small>',
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
    const input = document.getElementById("lavoro-occasionale-nome");
    const metadata = getWorkMetadata();
    const nome = metadata.nome;
    if (nome) {
      const coordinateText = getEditorText("lavoro-occasionale-coordinate");
      if (coordinateText && !metadata.coordinates) {
        event.preventDefault();
        event.stopImmediatePropagation();
        document.getElementById("lavoro-occasionale-coordinate")?.focus();
        const feedback = document.getElementById("squadra-feedback");
        if (feedback) {
          feedback.dataset.type = "error";
          feedback.textContent = "Coordinate non valide. Usa il formato: 44.494887, 11.342616";
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
      if (work.row?.lavoroOccasionaleDescrizione && !card?.querySelector(".lavoro-occasionale-card-description")) {
        const description = document.createElement("p");
        description.className = "lavoro-occasionale-card-description";
        description.textContent = `📋 ${work.row.lavoroOccasionaleDescrizione}`;
        title.parentElement?.insertAdjacentElement("afterend", description);
      }
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
        decorateSquadCards();
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
    installMapCapture();
    observeSquadCards();
    restoreWorkNameFromComposition();
    applyWorkNamesToData();
    refreshSuggestions();
    renderHistory();
    decorateSquadCards();
    syncMapMarkers();
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
    const observer = new MutationObserver(() => {
      if (feedback.dataset.type !== "success") return;
      observer.disconnect();
      window.setTimeout(() => {
        renderHistory();
        decorateSquadCards();
        syncMapMarkers();
      }, 0);
    });
    observer.observe(feedback, { childList: true, subtree: true, attributes: true });
    window.setTimeout(() => observer.disconnect(), 20000);
  });
  document.getElementById("squadra-commessa")?.addEventListener("change", () => {
    restoreWorkNameFromComposition();
    renderHistory();
  });
  document.getElementById("squadra-riferimento")?.addEventListener("change", restoreWorkNameFromComposition);
  document.addEventListener("click", () => window.setTimeout(syncMapMarkers, 350), true);

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
  `;
  document.head.appendChild(style);

  window.HeraLavoriOccasionali = {
    installed: true,
    version: "1.1.0",
    commessaId: COMMESSA_ID,
    refresh,
    getState: () => ({ ...state, lastError: state.lastError ? String(state.lastError?.message || state.lastError) : null })
  };
})();
