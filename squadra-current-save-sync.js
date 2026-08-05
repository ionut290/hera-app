(() => {
  "use strict";

  if (typeof autofillSquadraForm !== "function" || window.HeraSquadraCurrentSaveSync?.installed) return;

  const originalAutofillSquadraForm = autofillSquadraForm;
  const cloneRows = (rows) => (Array.isArray(rows) ? rows.map((row) => ({ ...row })) : []);

  function currentComposition(commessaId, dateKey) {
    const history = squadreHistoryByDate instanceof Map ? squadreHistoryByDate.get(dateKey) : null;
    return history instanceof Map ? history.get(commessaId) || null : null;
  }

  function showCurrentComposition(commessaId, dateKey, composition) {
    if (!composition || ui.squadraCommessa?.value !== commessaId || ui.squadraRiferimento?.value !== dateKey) return false;
    latestSquadraAutofillRequestId += 1;
    setSquadraRowsFromData(composition);
    updateSquadraAutofillHint(
      `Composizione salvata per questa commessa (${formatDateKeyForDisplay(dateKey)}). Puoi continuare a modificarla e premere “Fine” per salvare.`
    );
    return true;
  }

  autofillSquadraForm = async function autofillCurrentSquadraFirst() {
    const commessaId = ui.squadraCommessa?.value || "";
    setDefaultSquadraCompositionDate();
    const dateKey = ui.squadraRiferimento?.value || getNextDayDateKey();
    const current = currentComposition(commessaId, dateKey);
    if (commessaId && current && showCurrentComposition(commessaId, dateKey, current)) return true;
    return originalAutofillSquadraForm.apply(this, arguments);
  };

  ui.squadraForm?.addEventListener("submit", () => {
    const commessaId = ui.squadraCommessa?.value || "";
    const dateKey = ui.squadraRiferimento?.value || "";
    const rows = typeof readSquadraRows === "function" ? cloneRows(readSquadraRows()) : [];
    if (!commessaId || !dateKey || !rows.length || !ui.squadraFeedback) return;

    const feedback = ui.squadraFeedback;
    let finished = false;
    const stop = () => {
      if (finished) return;
      finished = true;
      observer.disconnect();
      clearTimeout(timeout);
    };
    const applyAfterSuccess = () => {
      if (feedback.dataset.type !== "success" || !feedback.textContent.includes("Composizione salvata")) return false;
      const history = squadreHistoryByDate.get(dateKey) instanceof Map
        ? new Map(squadreHistoryByDate.get(dateKey))
        : new Map();
      const existing = history.get(commessaId) || {};
      const composition = {
        ...existing,
        commessaId,
        commessaNome: (commesseById.get(commessaId) || {}).nome || "Commessa",
        riferimentoData: dateKey,
        dateKey,
        squadre: cloneRows(rows)
      };
      history.set(commessaId, composition);
      squadreHistoryByDate.set(dateKey, history);
      showCurrentComposition(commessaId, dateKey, composition);
      renderSquadre();
      renderTodaySummary();
      updateCommessaDashboard();
      renderCommesseHomeList();
      stop();
      return true;
    };

    const observer = new MutationObserver(() => {
      if (applyAfterSuccess()) return;
      if (feedback.dataset.type === "error") stop();
    });
    observer.observe(feedback, { childList: true, characterData: true, subtree: true, attributes: true });
    const timeout = setTimeout(stop, 20000);
    applyAfterSuccess();
  });

  window.HeraSquadraCurrentSaveSync = { installed: true };
})();

// Quando l'utente entra in “Le mie ore” dal calendario, usa la sorgente completa
// oreReports e riporta sempre mese e giorno alla data odierna. La vista statica
// mensile resta disponibile all'avvio, ma non può più nascondere giornate già salvate.
(() => {
  "use strict";

  if (window.HeraPersonalHoursCalendarEntryFix?.installed) return;

  const targetIds = ["calendar-choice-hours-btn", "calendar-hours-tab"];

  function openPersonalHoursOnToday(event) {
    try {
      window.HeraLightStartup?.enableHoursSource?.(event);
    } catch (error) {
      console.error("[LE MIE ORE] attivazione sorgente completa non riuscita", error);
    }

    try {
      if (typeof showCalendarToday === "function") showCalendarToday();
    } catch (error) {
      console.error("[LE MIE ORE] apertura sulla data odierna non riuscita", error);
    }
  }

  targetIds.forEach((id) => {
    const button = document.getElementById(id);
    if (!button || button.dataset.personalHoursEntryFixBound === "true") return;
    button.dataset.personalHoursEntryFixBound = "true";
    button.addEventListener("click", openPersonalHoursOnToday, true);
  });

  window.HeraPersonalHoursCalendarEntryFix = {
    installed: true,
    version: "1.0.0",
    targets: targetIds.slice()
  };
})();
