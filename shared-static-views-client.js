(() => {
  "use strict";

  const CORE_URL = "./shared-static-views-client-core.js?v=20260806-explicit-hours-v1";
  const HOURS_GUARD_URL = "./hours-source-explicit-guard.js?v=20260806a";
  const api = window.HeraSharedStaticViews;

  function loadHoursGuard() {
    if (document.querySelector('script[data-hours-source-explicit-guard]')) return;
    const script = document.createElement("script");
    script.src = HOURS_GUARD_URL;
    script.async = false;
    script.dataset.hoursSourceExplicitGuard = "1";
    document.head.appendChild(script);
  }

  function loadCore() {
    if (document.readyState === "loading") {
      document.write(`<script src="${CORE_URL}"><\/script>`);
      document.write(`<script src="${HOURS_GUARD_URL}" data-hours-source-explicit-guard="1"><\/script>`);
      return;
    }
    const script = document.createElement("script");
    script.src = CORE_URL;
    script.async = false;
    script.addEventListener("load", loadHoursGuard, { once: true });
    document.head.appendChild(script);
  }

  if (!api || typeof api.subscribe !== "function") {
    console.warn("[SAFE CALENDAR GUARD] API viste condivise non disponibile; carico il client originale.");
    loadCore();
    return;
  }

  const originalSubscribe = api.subscribe.bind(api);

  api.subscribe = function guardedSubscribe(type, key, callback, ...rest) {
    if (type !== "calendario" || typeof callback !== "function") {
      return originalSubscribe(type, key, callback, ...rest);
    }

    return originalSubscribe(type, key, (view, metadata = {}) => {
      const complete = Boolean(
        view &&
        view.schemaVersion === 2 &&
        view.completeRecords === true &&
        view.payload &&
        view.payload.schemaVersion === 2 &&
        view.payload.completeRecords === true &&
        Array.isArray(view.payload.reports)
      );

      if (complete) {
        callback(view, metadata);
        return;
      }

      console.warn("[SAFE CALENDAR GUARD] Vista calendario non completa: attivo il fallback Firestore originale.", {
        key,
        schemaVersion: view?.schemaVersion,
        completeRecords: view?.completeRecords,
        payloadSchemaVersion: view?.payload?.schemaVersion,
        payloadCompleteRecords: view?.payload?.completeRecords
      });

      queueMicrotask(() => {
        const fallback = window.HeraLightStartup?.enableHoursSource;
        if (typeof fallback === "function") {
          fallback({
            forceSharedCalendarFallback: true,
            reason: "shared-calendar-incomplete",
            month: key
          });
        }
      });
    }, ...rest);
  };

  window.HeraSafeCalendarGuard = {
    installed: true,
    version: "1.1.0"
  };

  loadCore();
})();