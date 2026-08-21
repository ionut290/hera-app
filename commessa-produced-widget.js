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

  function select() { hideWidget(); }
  function stop() { hideWidget(); }

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
      const nodes = Array.from(document.querySelectorAll("body *"));
      const match = nodes.find((el) => /subtotale\s+lavorazioni\s+fatte/i.test(el.textContent || ""));
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

  function findCompletedAmountStatItem() {
    const labels = Array.from(document.querySelectorAll(".commessa-stat-label"));
    const label = labels.find((el) => /avanzamento|lavorazioni\s+fatte|€\s*guadagnati|guadagnati/i.test(el.textContent || ""));
    if (label) return label.closest(".commessa-stat-item");

    const candidates = Array.from(document.querySelectorAll(".commessa-stat-item"));
    const semanticMatch = candidates.find((item) => /avanzamento|lavorazioni\s+fatte|€\s*guadagnati|guadagnati/i.test(item.textContent || ""));
    if (semanticMatch) return semanticMatch;

    const remainingValue = document.getElementById("commessa-stat-ore");
    const container = remainingValue?.closest?.(".commessa-stat-item")?.parentElement;
    if (!container) return null;
    const items = Array.from(container.children).filter((el) => el.classList?.contains("commessa-stat-item"));
    return items.length >= 3 ? items[1] : null;
  }

  function updateCompletedAmountStat() {
    const item = findCompletedAmountStatItem();
    if (!item) return false;

    const label = item.querySelector(".commessa-stat-label") || Array.from(item.children).find((el) => /avanzamento|lavorazioni\s+fatte|guadagnati/i.test(el.textContent || ""));
    const value = item.querySelector(".commessa-stat-value")
      || Array.from(item.querySelectorAll("[id]")).find((el) => el !== label && /%|€|\d/.test(el.textContent || ""))
      || Array.from(item.children).find((el) => el !== label && !el.classList?.contains("commessa-stat-icon"));
    if (!label || !value) return false;

    const subtotal = getSelectedCompletedSubtotal();
    const formatted = formatEuro(subtotal);
    value.textContent = formatted;
    label.textContent = "€ guadagnati";
    item.dataset.statAction = "euro-guadagnati";
    item.setAttribute("aria-label", `Euro guadagnati: ${formatted}`);
    item.title = `Subtotale lavorazioni fatte: ${formatted}`;
    return true;
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
      queueMicrotask(updateDashboardReplacementStats);
      return result;
    };
  }

  let refreshTimer = null;
  const scheduleRefresh = () => {
    clearTimeout(refreshTimer);
    refreshTimer = setTimeout(updateDashboardReplacementStats, 40);
  };

  const observer = new MutationObserver(scheduleRefresh);
  observer.observe(document.documentElement, { childList: true, subtree: true, characterData: true });

  document.addEventListener("click", () => setTimeout(updateDashboardReplacementStats, 80), true);
  window.addEventListener("hashchange", scheduleRefresh);
  window.addEventListener("popstate", scheduleRefresh);

  hideWidget();
  updateDashboardReplacementStats();
  setTimeout(updateDashboardReplacementStats, 250);
  setTimeout(updateDashboardReplacementStats, 1000);

  window.CommessaProducedWidget = { select, stop, removed: true, refresh: updateDashboardReplacementStats };
})();
