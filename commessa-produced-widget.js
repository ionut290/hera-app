/* Funzione PRODOTTO rimossa: nessun listener Firestore viene avviato. */
(() => {
  "use strict";

  const widget = document.getElementById("commessa-produced-widget");
  const toggle = document.getElementById("commessa-produced-toggle");
  const popover = document.getElementById("commessa-produced-popover");
  let sharedStatsCacheSignature = "";

  function hideWidget() {
    if (widget) widget.hidden = true;
    if (toggle) {
      toggle.setAttribute("aria-expanded", "false");
      toggle.disabled = true;
    }
    if (popover) {
      popover.setAttribute("aria-hidden", "true");
      popover.classList.remove("is-open");
    }
  }

  function select() {
    hideWidget();
  }

  function stop() {
    hideWidget();
  }

  function updateRemainingImpiantiStat() {
    if (typeof getSelectedCommessaDashboardStats !== "function") return;

    const value = document.getElementById("commessa-stat-ore");
    const item = value?.closest?.(".commessa-stat-item");
    const label = item?.querySelector?.(".commessa-stat-label");
    if (!value || !item || !label) return;

    const stats = getSelectedCommessaDashboardStats();
    const remaining = Math.max(0, Number(stats?.total || 0) - Number(stats?.done || 0));

    value.textContent = String(remaining);
    label.textContent = remaining === 1 ? "impianto da fare" : "impianti da fare";
    item.dataset.statAction = "impianti";
    item.setAttribute("aria-label", `Vai ai ${remaining} impianti da fare`);
    item.title = `${remaining} impiant${remaining === 1 ? "o" : "i"} ancora da fare`;
  }

  function getCachedImpiantiSignature(impianti = []) {
    return impianti
      .map((impianto) => {
        const id = String(impianto?.id || impianto?.idSap || impianto?.numero || "");
        const done = impianto?.done ? "1" : "0";
        const doneAt = typeof firestoreDateToMillis === "function"
          ? firestoreDateToMillis(impianto?.doneAt)
          : String(impianto?.doneAt || "");
        return `${id}:${done}:${doneAt}`;
      })
      .sort()
      .join("|");
  }

  function canReuseSharedImpiantiListener(commessaId) {
    return Boolean(
      commessaId
      && typeof unsubscribeCommessaStats !== "undefined"
      && unsubscribeCommessaStats instanceof Map
      && unsubscribeCommessaStats.has(commessaId)
      && typeof impiantiByCommessaId !== "undefined"
      && impiantiByCommessaId instanceof Map
      && impiantiByCommessaId.has(commessaId)
    );
  }

  function hydrateSelectedCommessaFromSharedCache(force = false) {
    if (!selectedCommessaId || !canReuseSharedImpiantiListener(selectedCommessaId)) return false;

    const cached = impiantiByCommessaId.get(selectedCommessaId);
    if (!Array.isArray(cached)) return false;

    const signature = `${selectedCommessaId}::${getCachedImpiantiSignature(cached)}`;
    if (!force && signature === sharedStatsCacheSignature) return true;
    sharedStatsCacheSignature = signature;

    currentImpianti = typeof applyPendingActionsToImpianti === "function"
      ? applyPendingActionsToImpianti(cached, selectedCommessaId)
      : cached.slice();

    if (typeof refreshImpiantoWhatsAppTemplateCache === "function") {
      refreshImpiantoWhatsAppTemplateCache(currentImpianti);
    }
    if (typeof renderSquadre === "function") renderSquadre();
    if (typeof renderHeaderActivitySummary === "function") renderHeaderActivitySummary();
    if (typeof updateCommessaDashboard === "function") updateCommessaDashboard();
    if (typeof renderImpianti === "function") renderImpianti();
    if (typeof renderMap === "function") renderMap();
    return true;
  }

  if (typeof subscribeImpianti === "function") {
    const originalSubscribeImpianti = subscribeImpianti;
    subscribeImpianti = function subscribeImpiantiWithoutDuplicateListener(...args) {
      const commessaId = selectedCommessaId;
      if (!canReuseSharedImpiantiListener(commessaId)) {
        return originalSubscribeImpianti.apply(this, args);
      }

      if (typeof subscribeFattoVisualEvidence === "function") {
        subscribeFattoVisualEvidence(commessaId);
      }
      hydrateSelectedCommessaFromSharedCache(true);
      console.log("Listener impianti dettaglio non duplicato: uso listener statistiche già attivo", {
        commessaId
      });
      return undefined;
    };
  }

  if (typeof renderCommesseHomeList === "function") {
    const originalRenderCommesseHomeList = renderCommesseHomeList;
    renderCommesseHomeList = function renderCommesseHomeListWithSelectedCacheSync(...args) {
      const result = originalRenderCommesseHomeList.apply(this, args);
      if (!unsubscribeImpianti) hydrateSelectedCommessaFromSharedCache(false);
      return result;
    };
  }

  if (typeof updateCommessaDashboard === "function") {
    const originalUpdateCommessaDashboard = updateCommessaDashboard;
    updateCommessaDashboard = function updateCommessaDashboardWithRemainingImpianti(...args) {
      const result = originalUpdateCommessaDashboard.apply(this, args);
      updateRemainingImpiantiStat();
      return result;
    };
  }

  hideWidget();
  updateRemainingImpiantiStat();
  window.CommessaProducedWidget = { select, stop, removed: true };
})();
