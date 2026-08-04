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
