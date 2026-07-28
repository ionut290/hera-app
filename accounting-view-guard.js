/* Mantiene stabile la nuova vista Gestione impianti e contabilità.
   Se uno script non è disponibile dalla cache, lo ricarica prima di aprire la schermata. */
(() => {
  "use strict";

  let loadingPromise = null;

  function loadScript(src) {
    return new Promise((resolve, reject) => {
      const script = document.createElement("script");
      script.src = `${src}${src.includes("?") ? "&" : "?"}retry=${Date.now()}`;
      script.async = false;
      script.onload = resolve;
      script.onerror = () => reject(new Error(`Caricamento non riuscito: ${src}`));
      document.head.appendChild(script);
    });
  }

  async function ensureAccountingView() {
    if (window.InreteWorkItemsV2 && window.AccountingV2) return;
    if (loadingPromise) return loadingPromise;
    loadingPromise = (async () => {
      if (!window.InreteWorkItemsV2) {
        await loadScript("inrete-work-items-v2.js?v=20260728b");
      }
      if (!window.AccountingV2) {
        await loadScript("accounting-v2.js?v=20260728e");
      }
      if (!window.InreteWorkItemsV2 || !window.AccountingV2) {
        throw new Error("La vista contabile non è disponibile.");
      }
    })().finally(() => {
      loadingPromise = null;
    });
    return loadingPromise;
  }

  window.openImpiantiManagement = async function openStableAccountingManagement(commessa) {
    try {
      await ensureAccountingView();
      return await window.AccountingV2.open(commessa);
    } catch (error) {
      console.error("Apertura Gestione impianti e contabilità non riuscita:", error);
      alert("Non è stato possibile caricare la contabilità. Controlla la connessione e riprova. La vecchia tabella non verrà aperta.");
      return null;
    }
  };
})();

(function installMapAutoFitGuard(root) {
  "use strict";

  function parseCoordinate(value, min, max) {
    if (value == null || String(value).trim() === "") return null;
    const coordinate = Number(String(value).trim().replace(",", "."));
    if (!Number.isFinite(coordinate) || coordinate < min || coordinate > max || coordinate === 0) return null;
    return coordinate;
  }

  function getPlantCoordinates(impianto) {
    const lat = parseCoordinate(impianto?.gpsY, -90, 90);
    const lng = parseCoordinate(impianto?.gpsX, -180, 180);
    return lat == null || lng == null ? null : [lat, lng];
  }

  function sanitizeBounds(rawBounds) {
    if (!Array.isArray(rawBounds)) return [];
    return rawBounds
      .map((point) => {
        if (!Array.isArray(point)) return null;
        const lat = parseCoordinate(point[0], -90, 90);
        const lng = parseCoordinate(point[1], -180, 180);
        return lat == null || lng == null ? null : [lat, lng];
      })
      .filter(Boolean);
  }

  const originalMarkerFactory = root.addImpiantoMarkerToMapLayer;
  if (typeof originalMarkerFactory === "function") {
    root.addImpiantoMarkerToMapLayer = function addValidatedImpiantoMarker(impianto, ...args) {
      const coordinates = getPlantCoordinates(impianto);
      if (!coordinates) return null;
      return originalMarkerFactory.call(this, {
        ...impianto,
        gpsY: coordinates[0],
        gpsX: coordinates[1]
      }, ...args);
    };
  }

  const originalRenderMap = root.renderMap;
  if (typeof originalRenderMap === "function") {
    root.renderMap = function renderMapWithValidBounds(...args) {
      const targetMap = typeof map !== "undefined" ? map : null;
      const originalFitBounds = targetMap?.fitBounds;
      if (typeof originalFitBounds !== "function") return originalRenderMap.apply(this, args);

      targetMap.fitBounds = function fitOnlyValidPlantBounds(rawBounds, options) {
        const validBounds = sanitizeBounds(rawBounds);
        if (!validBounds.length) return this;
        return originalFitBounds.call(this, validBounds, options);
      };

      try {
        return originalRenderMap.apply(this, args);
      } finally {
        targetMap.fitBounds = originalFitBounds;
      }
    };
  }

  root.HeraMapAutoFit = Object.freeze({
    getPlantCoordinates,
    sanitizeBounds
  });
})(window);
