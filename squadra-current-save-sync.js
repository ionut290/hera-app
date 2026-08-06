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

// La chat operatori non viene più utilizzata. Questo guard viene caricato subito
// dopo app.js: blocca il listener chat e la pulizia oraria prima che l'avvio
// autenticato possa eseguirli. I messaggi già presenti vengono eliminati una sola
// volta dal primo amministratore che apre l'app dopo questo aggiornamento.
(() => {
  "use strict";

  if (window.HeraChatDisabledGuard?.installed) return;

  const DELETE_MARKER_KEY = "heraChatMessagesDeletedV1";
  const state = {
    listenerStopped: false,
    retentionStopped: false,
    deletionAttempted: false,
    deletedMessages: 0,
    deletionComplete: false,
    lastError: null
  };

  function stopChatRuntime() {
    try {
      if (typeof unsubscribeChat === "function") {
        unsubscribeChat();
        state.listenerStopped = true;
      }
      if (typeof unsubscribeChat !== "undefined") unsubscribeChat = null;
    } catch (error) {
      state.lastError = error;
      console.warn("[CHAT DISATTIVATA] chiusura listener non riuscita", error);
    }

    try {
      if (typeof chatRetentionTimer !== "undefined" && chatRetentionTimer) {
        clearInterval(chatRetentionTimer);
        chatRetentionTimer = null;
        state.retentionStopped = true;
      }
    } catch (error) {
      state.lastError = error;
      console.warn("[CHAT DISATTIVATA] chiusura timer non riuscita", error);
    }
  }

  function hideChatInterface() {
    ["chat-open-btn", "chat-modal", "chat-clear-confirm-modal"].forEach((id) => {
      const node = document.getElementById(id);
      if (!node) return;
      node.classList.add("hidden");
      node.setAttribute("aria-hidden", "true");
      node.style.display = "none";
    });
  }

  function disableChatFunctions() {
    try {
      if (typeof subscribeChat === "function") {
        subscribeChat = function subscribeChatDisabled() {
          stopChatRuntime();
          try { chatMessages = []; } catch (_) {}
          return () => {};
        };
      }
      if (typeof startChatRetentionLoop === "function") {
        startChatRetentionLoop = function startChatRetentionLoopDisabled() {
          stopChatRuntime();
          return false;
        };
      }
      if (typeof stopChatRetentionLoop === "function") {
        stopChatRetentionLoop = function stopChatRetentionLoopDisabled() {
          stopChatRuntime();
          return true;
        };
      }
      if (typeof purgeOldChatMessages === "function") {
        purgeOldChatMessages = async function purgeOldChatMessagesDisabled() {
          return { disabled: true, deleted: 0 };
        };
      }
    } catch (error) {
      state.lastError = error;
      console.error("[CHAT DISATTIVATA] sostituzione funzioni non riuscita", error);
    }
  }

  async function deleteStoredMessagesOnce() {
    if (state.deletionAttempted) return;
    try {
      if (localStorage.getItem(DELETE_MARKER_KEY) === "done") {
        state.deletionComplete = true;
        return;
      }
    } catch (_) {}

    if (typeof canManageData !== "function" || !canManageData()) return;
    if (typeof db === "undefined" || !db) return;

    state.deletionAttempted = true;
    try {
      const snapshot = await db.collection("chatMessages").get();
      const docs = snapshot.docs || [];
      for (let index = 0; index < docs.length; index += 450) {
        const batch = db.batch();
        docs.slice(index, index + 450).forEach((doc) => batch.delete(doc.ref));
        await batch.commit();
      }
      state.deletedMessages = docs.length;
      state.deletionComplete = true;
      try { chatMessages = []; } catch (_) {}
      try { localStorage.setItem(DELETE_MARKER_KEY, "done"); } catch (_) {}
      console.info(`[CHAT DISATTIVATA] eliminati definitivamente ${docs.length} messaggi.`);
    } catch (error) {
      state.lastError = error;
      state.deletionAttempted = false;
      console.warn("[CHAT DISATTIVATA] eliminazione messaggi rinviata: serve un amministratore autorizzato", error);
    }
  }

  stopChatRuntime();
  disableChatFunctions();
  hideChatInterface();

  if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", hideChatInterface, { once: true });
  }

  try {
    const authInstance = typeof firebase !== "undefined" && firebase.auth ? firebase.auth() : null;
    authInstance?.onAuthStateChanged((user) => {
      if (!user) return;
      window.setTimeout(() => {
        stopChatRuntime();
        disableChatFunctions();
        hideChatInterface();
        deleteStoredMessagesOnce();
      }, 1200);
    });
  } catch (error) {
    state.lastError = error;
    console.warn("[CHAT DISATTIVATA] controllo autenticazione non disponibile", error);
  }

  window.HeraChatDisabledGuard = {
    installed: true,
    version: "1.0.0",
    collection: "chatMessages",
    getState: () => ({ ...state, lastError: state.lastError ? String(state.lastError?.message || state.lastError) : null }),
    retryDeletion: deleteStoredMessagesOnce
  };
})();
