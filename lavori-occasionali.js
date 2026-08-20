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

  const normalizeName = (value) => String(value || "")
    .trim()
    .replace(/\s+/g, " ")
    .toLocaleUpperCase("it-IT");

  function isOccasionalSelected() {
    return document.getElementById("squadra-commessa")?.value === COMMESSA_ID;
  }

  function getWorkName() {
    return normalizeName(document.getElementById("lavoro-occasionale-nome")?.value);
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
      if (typeof commesseById !== "undefined" && commesseById instanceof Map) {
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
      '<input id="lavoro-occasionale-nome" type="text" list="lavori-occasionali-options"',
      ' maxlength="120" autocomplete="off" placeholder="Es. Parco Zucca, Scuole Granarolo">',
      '<datalist id="lavori-occasionali-options"></datalist>',
      '<small>Il nome verrà salvato nella squadra e nello storico del lavoro.</small>'
    ].join("");

    const dateField = form.querySelector(".squadra-date-field");
    (dateField || commessaSelect).insertAdjacentElement("afterend", field);

    const input = field.querySelector("input");
    const refresh = () => {
      const active = isOccasionalSelected();
      field.classList.toggle("hidden", !active);
      input.required = active;
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
      const nome = getWorkName();
      return Array.isArray(rows) ? rows.map((row) => ({
        ...row,
        lavoroOccasionale: true,
        lavoroOccasionaleNome: nome
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
      if (input) input.value = nome;
    } catch (error) {
      state.lastError = error;
    }
  }

  function validateBeforeCoreSave(event) {
    if (!isOccasionalSelected()) return;
    const input = document.getElementById("lavoro-occasionale-nome");
    const nome = getWorkName();
    if (nome) {
      input.value = nome;
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

  function refresh() {
    installVirtualCommessa();
    observeCommessaSelector();
    installWorkField();
    wrapRowReader();
    installHistory();
    restoreWorkNameFromComposition();
    refreshSuggestions();
    renderHistory();
    state.installed = state.virtualCommessaReady && state.rowsWrapped;
  }

  document.addEventListener("DOMContentLoaded", refresh, { once: true });
  window.addEventListener("load", () => {
    refresh();
    window.setTimeout(refresh, 1200);
    window.setTimeout(refresh, 4000);
  }, { once: true });

  document.getElementById("squadra-form")?.addEventListener("submit", validateBeforeCoreSave, true);
  document.getElementById("squadra-commessa")?.addEventListener("change", () => {
    restoreWorkNameFromComposition();
    renderHistory();
  });
  document.getElementById("squadra-riferimento")?.addEventListener("change", restoreWorkNameFromComposition);

  window.HeraLavoriOccasionali = {
    installed: true,
    version: "1.0.0",
    commessaId: COMMESSA_ID,
    refresh,
    getState: () => ({ ...state, lastError: state.lastError ? String(state.lastError?.message || state.lastError) : null })
  };
})();
