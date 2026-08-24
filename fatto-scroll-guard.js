(() => {
  "use strict";

  const HELPER_VERSION = "20260824-oneclick2";
  let helpersLoaded = false;
  let helperTask = 0;
  let helperTaskKind = "";
  let impiantiPageObserver = null;

  // Non intercettare touch/click di FATTO: Safari gestisce gia la distinzione
  // fra scorrimento e tocco. Il vecchio blocco di 450 ms annullava il primo
  // click reale quando il dito si spostava di pochi pixel.

  function loadScriptOnce(selector, src, datasetKey, errorMessage) {
    if (document.querySelector(selector)) return;
    const script = document.createElement("script");
    script.src = src;
    script.async = true;
    script.dataset[datasetKey] = "true";
    script.onerror = () => console.warn(errorMessage);
    document.head.appendChild(script);
  }

  function loadRecommendedHelpers() {
    if (helpersLoaded) return;
    helpersLoaded = true;
    helperTask = 0;
    helperTaskKind = "";

    if (!window.HeraSquadContext?.installed) {
      loadScriptOnce(
        'script[data-squad-context-bridge]',
        `./squad-context-bridge.js?v=${HELPER_VERSION}`,
        "squadContextBridge",
        "Contesto squadra reale non caricato."
      );
    }

    // Il motore stabile integra direttamente stime adattive e attrezzature.
    // Non vengono più caricati i due vecchi observer ricorsivi che congelavano la PWA.
    if (!window.HeraRecommendedTrafficWeather?.installed) {
      loadScriptOnce(
        'script[data-recommended-traffic-weather]',
        `./recommended-traffic-weather.js?v=${HELPER_VERSION}`,
        "recommendedTrafficWeather",
        "Traffico/meteo Impianti consigliati non caricato."
      );
    }

    if (!window.HeraStreetViewCards?.installed) {
      loadScriptOnce(
        'script[data-street-view-cards]',
        `./street-view-cards.js?v=${HELPER_VERSION}`,
        "streetViewCards",
        "Street View card impianti non caricato."
      );
    }
  }

  function cancelScheduledHelpers() {
    if (!helperTask) return;
    if (helperTaskKind === "idle" && typeof window.cancelIdleCallback === "function") {
      window.cancelIdleCallback(helperTask);
    } else {
      window.clearTimeout(helperTask);
    }
    helperTask = 0;
    helperTaskKind = "";
  }

  function scheduleRecommendedHelpers(options = {}) {
    if (helpersLoaded) return;
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
      helperTask = window.setTimeout(loadRecommendedHelpers, 1000);
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
      if (!isImpiantiPageVisible(page)) return;
      scheduleRecommendedHelpers();
      impiantiPageObserver?.disconnect();
      impiantiPageObserver = null;
    };
    checkVisibility();
    if (!helpersLoaded && !helperTask && !impiantiPageObserver) {
      impiantiPageObserver = new MutationObserver(checkVisibility);
      impiantiPageObserver.observe(page, {
        attributes: true,
        attributeFilter: ["class", "hidden", "aria-hidden", "style"]
      });
    }
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
    }
  }, true);

  if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", initializeLazyHelpers, { once: true });
  } else {
    initializeLazyHelpers();
  }
})();
