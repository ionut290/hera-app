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

  function formatEuro(value) {
    const numeric = Number(value);
    return new Intl.NumberFormat("it-IT", {
      style: "currency",
      currency: "EUR",
      minimumFractionDigits: 2,
      maximumFractionDigits: 2
    }).format(Number.isFinite(numeric) ? numeric : 0);
  }

  function parseItalianCurrency(value) {
    const raw = String(value ?? "").trim();
    if (!raw) return NaN;
    const normalized = raw
      .replace(/\s/g, "")
      .replace(/€/g, "")
      .replace(/\./g, "")
      .replace(",", ".")
      .replace(/[^0-9.-]/g, "");
    const numeric = Number(normalized);
    return Number.isFinite(numeric) ? numeric : NaN;
  }

  function readRenderedCompletedSubtotal() {
    try {
      const elements = Array.from(document.querySelectorAll("body *"));
      const match = elements.find((element) => {
        if (element.children?.length) return false;
        return /subtotale\s+lavorazioni\s+fatte/i.test(element.textContent || "");
      });
      if (!match) return NaN;
      return parseItalianCurrency(match.textContent);
    } catch (_) {
      return NaN;
    }
  }

  function getSelectedCompletedSubtotal() {
    try {
      if (typeof selectedCommessaId !== "undefined" && selectedCommessaId && typeof commesseById !== "undefined") {
        const commessa = commesseById?.get?.(selectedCommessaId) || {};
        const candidates = [
          commessa.completedSubtotal,
          commessa.subtotalCompleted,
          commessa.subtotaleLavorazioniFatte,
          commessa.subtotaleFatto,
          commessa.completedWorkSubtotal,
          commessa.importoLavorazioniFatte
        ];
        for (const candidate of candidates) {
          const numeric = Number(candidate);
          if (Number.isFinite(numeric)) return numeric;
        }
      }
    } catch (_) {}

    const rendered = readRenderedCompletedSubtotal();
    return Number.isFinite(rendered) ? rendered : 0;
  }

  function getDashboardStatItems() {
    const remainingValue = document.getElementById("commessa-stat-ore");
    const remainingItem = remainingValue?.closest?.(".commessa-stat-item");
    const container = remainingItem?.parentElement;
    if (!container) return [];
    return Array.from(container.children).filter((element) => element.classList?.contains("commessa-stat-item"));
  }

  function updateCompletedAmountStat() {
    const items = getDashboardStatItems();
    if (items.length < 3) return;

    const item = items[1];
    const label = item.querySelector?.(".commessa-stat-label");
    const value = item.querySelector?.(".commessa-stat-value")
      || Array.from(item.querySelectorAll?.("[id]") || []).find((element) => element !== label)
      || Array.from(item.children || []).find((element) => element !== label && !element.classList?.contains("commessa-stat-icon"));
    if (!value || !label) return;

    const subtotal = getSelectedCompletedSubtotal();
    value.textContent = formatEuro(subtotal);
    label.textContent = "lavorazioni fatte";
    item.dataset.statAction = "lavorazioni-fatte";
    item.setAttribute("aria-label", `Lavorazioni fatte: ${formatEuro(subtotal)}`);
    item.title = `Subtotale lavorazioni fatte: ${formatEuro(subtotal)}`;
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

  function updateDashboardReplacementStats() {
    updateCompletedAmountStat();
    updateRemainingImpiantiStat();
  }

  if (typeof updateCommessaDashboard === "function") {
    const originalUpdateCommessaDashboard = updateCommessaDashboard;
    updateCommessaDashboard = function updateCommessaDashboardWithReplacementStats(...args) {
      const result = originalUpdateCommessaDashboard.apply(this, args);
      updateDashboardReplacementStats();
      return result;
    };
  }

  hideWidget();
  updateDashboardReplacementStats();
  window.CommessaProducedWidget = { select, stop, removed: true };
})();
