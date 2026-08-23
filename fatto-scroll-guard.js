(() => {
  "use strict";

  const MOVE_THRESHOLD_PX = 8;
  const BLOCK_AFTER_SCROLL_MS = 450;
  const HELPER_VERSION = "20260823-stability2";
  let lastTouchScrollAt = 0;
  let touchStartX = 0;
  let touchStartY = 0;
  let trackingTouch = false;
  let moved = false;
  let helperTask = 0;
  let helperTaskKind = "";
  let trafficTask = 0;
  let streetViewTask = 0;
  let impiantiPageObserver = null;

  function cleanText(value) {
    return String(value || "").replace(/\s+/g, " ").trim().toUpperCase();
  }

  function findFattoButton(target) {
    const button = target?.closest?.('button, [role="button"], input[type="button"], input[type="submit"]');
    if (!button) return null;
    const label = cleanText(button.value || button.getAttribute("aria-label") || button.textContent);
    return /(^|\s)FATTO($|\s)/.test(label) ? button : null;
  }

  document.addEventListener("touchstart", (event) => {
    const touch = event.touches?.[0];
    if (!touch) return;
    trackingTouch = true;
    moved = false;
    touchStartX = touch.clientX;
    touchStartY = touch.clientY;
  }, { capture: true, passive: true });

  document.addEventListener("touchmove", (event) => {
    if (!trackingTouch) return;
    const touch = event.touches?.[0];
    if (!touch) return;
    const dx = touch.clientX - touchStartX;
    const dy = touch.clientY - touchStartY;
    if (!moved && Math.hypot(dx, dy) >= MOVE_THRESHOLD_PX) moved = true;
    if (moved) lastTouchScrollAt = Date.now();
  }, { capture: true, passive: true });

  document.addEventListener("touchend", () => {
    if (moved) lastTouchScrollAt = Date.now();
    trackingTouch = false;
    moved = false;
  }, { capture: true, passive: true });

  document.addEventListener("touchcancel", () => {
    if (moved) lastTouchScrollAt = Date.now();
    trackingTouch = false;
    moved = false;
  }, { capture: true, passive: true });

  document.addEventListener("click", (event) => {
    const button = findFattoButton(event.target);
    if (!button || Date.now() - lastTouchScrollAt > BLOCK_AFTER_SCROLL_MS) return;
    event.preventDefault();
    event.stopPropagation();
    event.stopImmediatePropagation();
    button.classList.remove("fatto-scroll-guard-blocked");
    void button.offsetWidth;
    button.classList.add("fatto-scroll-guard-blocked");
    window.setTimeout(() => button.classList.remove("fatto-scroll-guard-blocked"), 260);
  }, true);

  function loadScriptOnce(selector, src, datasetKey, errorMessage, onload) {
    const existing = document.querySelector(selector);
    if (existing) {
      onload?.();
      return existing;
    }
    const script = document.createElement("script");
    script.src = src;
    script.async = true;
    script.dataset[datasetKey] = "true";
    script.onerror = () => console.warn(errorMessage);
    if (onload) script.addEventListener("load", onload, { once: true });
    document.head.appendChild(script);
    return script;
  }

  function loadSquadContext() {
    if (window.HeraSquadContext?.installed) return;
    loadScriptOnce(
      'script[data-squad-context-bridge]',
      `./squad-context-bridge.js?v=${HELPER_VERSION}`,
      "squadContextBridge",
      "Contesto squadra reale non caricato."
    );
  }

  function loadTrafficWeather() {
    trafficTask = 0;
    if (window.HeraRecommendedTrafficWeather?.installed) {
      window.HeraRecommendedTrafficWeather.refresh?.();
      return;
    }
    loadScriptOnce(
      'script[data-recommended-traffic-weather]',
      `./recommended-traffic-weather.js?v=${HELPER_VERSION}`,
      "recommendedTrafficWeather",
      "Traffico/meteo Impianti consigliati non caricato.",
      () => window.HeraRecommendedTrafficWeather?.refresh?.()
    );
  }

  function recommendedPanelOpen() {
    const panel = document.getElementById("recommended-plants-panel");
    return Boolean(panel && !panel.classList.contains("hidden"));
  }

  function loadStreetView() {
    streetViewTask = 0;
    if (recommendedPanelOpen()) {
      scheduleStreetView(3500);
      return;
    }
    if (window.HeraStreetViewCards?.installed) return;
    loadScriptOnce(
      'script[data-street-view-cards]',
      `./street-view-cards.js?v=${HELPER_VERSION}`,
      "streetViewCards",
      "Street View card impianti non caricato."
    );
  }

  function cancelScheduledHelpers() {
    if (helperTask) {
      if (helperTaskKind === "idle" && typeof window.cancelIdleCallback === "function") {
        window.cancelIdleCallback(helperTask);
      } else {
        window.clearTimeout(helperTask);
      }
    }
    helperTask = 0;
    helperTaskKind = "";
  }

  function cancelTrafficTask() {
    if (!trafficTask) return;
    window.clearTimeout(trafficTask);
    trafficTask = 0;
  }

  function cancelStreetViewTask() {
    if (!streetViewTask) return;
    window.clearTimeout(streetViewTask);
    streetViewTask = 0;
  }

  function scheduleTraffic(delay = 700) {
    cancelTrafficTask();
    trafficTask = window.setTimeout(() => {
      if (!recommendedPanelOpen()) return;
      if (typeof window.requestIdleCallback === "function") {
        window.requestIdleCallback(loadTrafficWeather, { timeout: 1600 });
      } else {
        loadTrafficWeather();
      }
    }, delay);
  }

  function scheduleStreetView(delay = 5500) {
    cancelStreetViewTask();
    streetViewTask = window.setTimeout(() => {
      if (typeof window.requestIdleCallback === "function") {
        window.requestIdleCallback(loadStreetView, { timeout: 3000 });
      } else {
        loadStreetView();
      }
    }, delay);
  }

  function loadRecommendedHelpers() {
    helperTask = 0;
    helperTaskKind = "";
    loadSquadContext();
    if (recommendedPanelOpen()) scheduleTraffic(650);
  }

  function scheduleRecommendedHelpers(options = {}) {
    if (options.immediate) {
      cancelScheduledHelpers();
      loadRecommendedHelpers();
      return;
    }
    if (helperTask) return;
    if (typeof window.requestIdleCallback === "function") {
      helperTaskKind = "idle";
      helperTask = window.requestIdleCallback(loadRecommendedHelpers, { timeout: 2500 });
    } else {
      helperTaskKind = "timeout";
      helperTask = window.setTimeout(loadRecommendedHelpers, 1200);
    }
  }

  function isImpiantiPageVisible(page) {
    if (!page || page.hidden || page.classList.contains("hidden")) return false;
    if (page.getAttribute("aria-hidden") === "true") return false;
    return page.style?.display !== "none" && page.style?.visibility !== "hidden";
  }

  function bindImpiantiPage() {
    const page = document.getElementById("impianti-page");
    if (!page) return false;
    const checkVisibility = () => {
      if (!isImpiantiPageVisible(page)) {
        cancelTrafficTask();
        cancelStreetViewTask();
        return;
      }
      scheduleRecommendedHelpers();
      scheduleStreetView();
    };
    if (!impiantiPageObserver) {
      impiantiPageObserver = new MutationObserver(checkVisibility);
      impiantiPageObserver.observe(page, {
        attributes: true,
        attributeFilter: ["class", "hidden", "aria-hidden", "style"]
      });
    }
    checkVisibility();
    return true;
  }

  function initializeLazyHelpers() {
    if (bindImpiantiPage()) return;
    let attempts = 0;
    const timer = window.setInterval(() => {
      attempts += 1;
      if (bindImpiantiPage() || attempts >= 6) window.clearInterval(timer);
    }, 500);
  }

  document.addEventListener("click", (event) => {
    if (event.target?.closest?.("#recommended-plants-btn")) {
      scheduleRecommendedHelpers({ immediate: true });
      cancelStreetViewTask();
      scheduleStreetView(7000);
    }
  }, true);

  window.addEventListener("hera:recommended-ready", () => scheduleTraffic(500));
  window.addEventListener("hera:recommended-closed", () => {
    cancelTrafficTask();
    scheduleStreetView(1800);
  });

  if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", initializeLazyHelpers, { once: true });
  } else {
    initializeLazyHelpers();
  }
})();
