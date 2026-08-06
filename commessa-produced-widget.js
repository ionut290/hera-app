/* Funzione PRODOTTO rimossa: nessun listener Firestore viene avviato. */
(() => {
  "use strict";

  const widget = document.getElementById("commessa-produced-widget");
  const toggle = document.getElementById("commessa-produced-toggle");
  const popover = document.getElementById("commessa-produced-popover");

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
