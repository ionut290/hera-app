/* Funzione PRODOTTO rimossa: nessun listener Firestore viene avviato. */
(() => {
  "use strict";

  const widget = document.getElementById("commessa-produced-widget");
  const toggle = document.getElementById("commessa-produced-toggle");
  const popover = document.getElementById("commessa-produced-popover");
  let refreshTimer = 0;
  let pageVisibilityObserver = null;

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
      const root = document.getElementById("impianti-page");
      if (!root) return NaN;
      const nodes = Array.from(root.querySelectorAll([
        "[data-completed-subtotal]",
        ".commessa-dashboard-card p",
        ".commessa-dashboard-card span",
        ".commessa-dashboard-card strong",
        ".commessa-dashboard-card small",
        ".accounting-summary p",
        ".accounting-summary span",
        ".accounting-summary strong"
      ].join(",")));
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

  function setStatIcon(item, type) {
    const icon = item?.querySelector?.(".commessa-stat-icon");
    if (!icon || icon.dataset.dashboardIcon === type) return;

    if (type === "euro") {
      icon.innerHTML = `
        <svg viewBox="0 0 64 64" width="1em" height="1em" aria-hidden="true" focusable="false">
          <circle cx="32" cy="32" r="25" fill="none" stroke="currentColor" stroke-width="4"/>
          <text x="32" y="41" text-anchor="middle" font-size="31" font-weight="700" font-family="Arial, Helvetica, sans-serif" fill="currentColor">€</text>
        </svg>`;
    } else if (type === "done") {
      icon.innerHTML = `
        <svg viewBox="0 0 64 64" width="1em" height="1em" aria-hidden="true" focusable="false">
          <circle cx="32" cy="32" r="25" fill="none" stroke="currentColor" stroke-width="4"/>
          <path d="M20 33.5 28.5 42 45 23.5" fill="none" stroke="currentColor" stroke-width="5" stroke-linecap="round" stroke-linejoin="round"/>
        </svg>`;
    }

    icon.dataset.dashboardIcon = type;
  }

  function findCompletedAmountStatItem() {
    const root = document.getElementById("impianti-page") || document;
    const labels = Array.from(root.querySelectorAll(".commessa-stat-label"));
    const label = labels.find((el) => /avanzamento|lavorazioni\s+fatte|€\s*guadagnati|guadagnati/i.test(el.textContent || ""));
    if (label) return label.closest(".commessa-stat-item");

    const candidates = Array.from(root.querySelectorAll(".commessa-stat-item"));
    const semanticMatch = candidates.find((item) => /avanzamento|lavorazioni\s+fatte|€\s*guadagnati|guadagnati/i.test(item.textContent || ""));
    if (semanticMatch) return semanticMatch;

    const doneValue = document.getElementById("commessa-stat-ore");
    const container = doneValue?.closest?.(".commessa-stat-item")?.parentElement;
    if (!container) return null;
    const items = Array.from(container.children).filter((el) => el.classList?.contains("commessa-stat-item"));
    return items.length >= 3 ? items[1] : null;
  }

  function updateCompletedAmountStat() {
    const item = findCompletedAmountStatItem();
    if (!item) return false;

    const label = item.querySelector(".commessa-stat-label")
      || Array.from(item.children).find((el) => /avanzamento|lavorazioni\s+fatte|guadagnati/i.test(el.textContent || ""));
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
    setStatIcon(item, "euro");
    return true;
  }

  function updateDoneImpiantiStat() {
    if (typeof getSelectedCommessaDashboardStats !== "function") return;
    const value = document.getElementById("commessa-stat-ore");
    const item = value?.closest?.(".commessa-stat-item");
    const label = item?.querySelector?.(".commessa-stat-label");
    if (!value || !item || !label) return;

    const stats = getSelectedCommessaDashboardStats();
    const done = Math.max(0, Number(stats?.done || 0));
    value.textContent = String(done);
    label.textContent = done === 1 ? "impianto fatto" : "impianti fatti";
    item.dataset.statAction = "impianti-fatti";
    item.setAttribute("aria-label", `${done} impiant${done === 1 ? "o fatto" : "i fatti"}`);
    item.title = `${done} impiant${done === 1 ? "o completato" : "i completati"}`;
    setStatIcon(item, "done");
  }

  function updateDashboardReplacementStats() {
    hideWidget();
    updateCompletedAmountStat();
    updateDoneImpiantiStat();
  }

  function scheduleRefresh(delay = 40) {
    window.clearTimeout(refreshTimer);
    refreshTimer = window.setTimeout(updateDashboardReplacementStats, delay);
  }

  if (typeof updateCommessaDashboard === "function") {
    const originalUpdateCommessaDashboard = updateCommessaDashboard;
    updateCommessaDashboard = function updateCommessaDashboardWithReplacementStats(...args) {
      const result = originalUpdateCommessaDashboard.apply(this, args);
      queueMicrotask(updateDashboardReplacementStats);
      return result;
    };
  }

  function isImpiantiPageVisible(page) {
    if (!page || page.hidden || page.classList.contains("hidden")) return false;
    if (page.getAttribute("aria-hidden") === "true") return false;
    return page.style?.display !== "none" && page.style?.visibility !== "hidden";
  }

  function installPageVisibilityObserver() {
    const page = document.getElementById("impianti-page");
    if (!page || pageVisibilityObserver) return;
    const refreshWhenVisible = () => {
      if (isImpiantiPageVisible(page)) scheduleRefresh(0);
    };
    pageVisibilityObserver = new MutationObserver(refreshWhenVisible);
    pageVisibilityObserver.observe(page, {
      attributes: true,
      attributeFilter: ["class", "hidden", "aria-hidden", "style"]
    });
    refreshWhenVisible();
  }

  function shouldRefreshAfterClick(target) {
    return Boolean(target?.closest?.([
      "#back-to-home-btn",
      ".commessa-stat-item",
      "#view-todo-btn",
      "#view-done-btn",
      "[data-action-key=\"done\"]",
      "[data-action=\"done\"]",
      ".impianto-fatto-btn",
      ".fatto-btn"
    ].join(",")));
  }

  document.addEventListener("click", (event) => {
    if (shouldRefreshAfterClick(event.target)) scheduleRefresh(80);
  }, true);
  window.addEventListener("hashchange", () => scheduleRefresh(0));
  window.addEventListener("popstate", () => scheduleRefresh(0));

  if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", installPageVisibilityObserver, { once: true });
  } else {
    installPageVisibilityObserver();
  }

  hideWidget();
  scheduleRefresh(0);
  window.setTimeout(updateDashboardReplacementStats, 250);

  window.CommessaProducedWidget = {
    select,
    stop,
    removed: true,
    refresh: updateDashboardReplacementStats
  };
})();

// Carica la funzione opzionale "Impianti consigliati" solo quando si entra
// nella pagina Impianti. La Home e le letture Firestore restano prioritarie.
(() => {
  "use strict";

  const VERSION = "20260823-stability2";
  let pageObserver = null;
  let recommendedTask = 0;
  let recommendedTaskKind = "";

  function loadStyleOnce() {
    if (document.querySelector("link[data-recommended-plants-style]")) return;
    const link = document.createElement("link");
    link.rel = "stylesheet";
    link.href = `recommended-plants.css?v=${VERSION}`;
    link.dataset.recommendedPlantsStyle = "1";
    document.head.appendChild(link);
  }

  function loadScriptOnce(selector, src, datasetKey, options = {}) {
    if (document.querySelector(selector)) return null;
    const script = document.createElement("script");
    script.src = src;
    script.async = options.async !== false;
    script.dataset[datasetKey] = "1";
    script.onerror = () => console.warn(options.errorMessage || `Modulo non caricato: ${src}`);
    document.body.appendChild(script);
    return script;
  }

  function loadFattoScrollGuard() {
    loadScriptOnce(
      "script[data-fatto-scroll-guard]",
      `fatto-scroll-guard.js?v=${VERSION}`,
      "fattoScrollGuard",
      { async: false, errorMessage: "Protezione pulsante FATTO non caricata." }
    );
  }

  function loadRecommendedCore() {
    recommendedTask = 0;
    recommendedTaskKind = "";
    if (window.HeraRecommendedPlants?.installed) return;
    loadStyleOnce();
    loadScriptOnce(
      "script[data-recommended-plants-script]",
      `recommended-plants.js?v=${VERSION}`,
      "recommendedPlantsScript",
      { async: false, errorMessage: "Impianti consigliati non caricati." }
    );
  }

  function cancelRecommendedTask() {
    if (!recommendedTask) return;
    if (recommendedTaskKind === "idle" && typeof window.cancelIdleCallback === "function") {
      window.cancelIdleCallback(recommendedTask);
    } else {
      window.clearTimeout(recommendedTask);
    }
    recommendedTask = 0;
    recommendedTaskKind = "";
  }

  function scheduleRecommendedCore(options = {}) {
    if (window.HeraRecommendedPlants?.installed || document.querySelector("script[data-recommended-plants-script]")) return;
    if (options.immediate) {
      cancelRecommendedTask();
      loadRecommendedCore();
      return;
    }
    if (recommendedTask) return;

    if (typeof window.requestIdleCallback === "function") {
      recommendedTaskKind = "idle";
      recommendedTask = window.requestIdleCallback(loadRecommendedCore, { timeout: 1200 });
    } else {
      recommendedTaskKind = "timeout";
      recommendedTask = window.setTimeout(loadRecommendedCore, 700);
    }
  }

  function isPageVisible(page) {
    if (!page || page.hidden || page.classList.contains("hidden")) return false;
    if (page.getAttribute("aria-hidden") === "true") return false;
    return page.style?.display !== "none" && page.style?.visibility !== "hidden";
  }

  function bindImpiantiPage() {
    const page = document.getElementById("impianti-page");
    if (!page) return false;

    const activateWhenVisible = () => {
      if (!isPageVisible(page)) return;
      scheduleRecommendedCore();
      pageObserver?.disconnect();
      pageObserver = null;
    };

    activateWhenVisible();
    if (!pageObserver && !isPageVisible(page)) {
      pageObserver = new MutationObserver(activateWhenVisible);
      pageObserver.observe(page, {
        attributes: true,
        attributeFilter: ["class", "hidden", "aria-hidden", "style"]
      });
    }
    return true;
  }

  function install() {
    loadFattoScrollGuard();
    if (bindImpiantiPage()) return;
    let attempts = 0;
    const timer = window.setInterval(() => {
      attempts += 1;
      if (bindImpiantiPage() || attempts >= 6) window.clearInterval(timer);
    }, 500);
  }

  document.addEventListener("click", (event) => {
    if (event.target?.closest?.("#recommended-plants-btn")) scheduleRecommendedCore({ immediate: true });
  }, true);

  if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", install, { once: true });
  } else {
    install();
  }
})();
