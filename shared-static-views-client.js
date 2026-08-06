(() => {
  "use strict";

  const CORE_URL = "./shared-static-views-client-core.js?v=20260806-explicit-hours-v3";
  const HOURS_GUARD_URL = "./hours-source-explicit-guard.js?v=20260806b";
  const api = window.HeraSharedStaticViews;
  const prematureHoursState = {
    stoppedHoursListener: false,
    stoppedApprovalsListener: false,
    errors: []
  };

  // app.js viene eseguito prima di questo client e può aprire i listener completi
  // delle ore durante l'avvio. Li chiudiamo prima che consegnino l'intero storico;
  // la vista mensile condivisa alimenterà Home e calendario. La sorgente completa
  // potrà essere riaperta dal core soltanto dopo un'azione esplicita dell'utente.
  function stopPrematureHoursSubscriptions() {
    try {
      if (typeof unsubscribeHoursStats === "function") {
        unsubscribeHoursStats();
        unsubscribeHoursStats = null;
        prematureHoursState.stoppedHoursListener = true;
      }
    } catch (error) {
      prematureHoursState.errors.push(`oreReports: ${error?.message || error}`);
    }

    try {
      if (typeof unsubscribeHoursApprovals === "function") {
        unsubscribeHoursApprovals();
        unsubscribeHoursApprovals = null;
        prematureHoursState.stoppedApprovalsListener = true;
      }
    } catch (error) {
      prematureHoursState.errors.push(`oreApprovalRequests: ${error?.message || error}`);
    }

    if (prematureHoursState.stoppedHoursListener || prematureHoursState.stoppedApprovalsListener) {
      console.debug("[SAFE CALENDAR GUARD] listener ore completi fermati durante l'avvio", {
        oreReports: prematureHoursState.stoppedHoursListener,
        approvazioni: prematureHoursState.stoppedApprovalsListener
      });
    }
  }

  function loadHoursGuard(callback) {
    const existing = document.querySelector('script[data-hours-source-explicit-guard]');
    if (existing) {
      if (window.HeraHoursSourceExplicitGuard?.installed) callback?.();
      else existing.addEventListener("load", () => callback?.(), { once: true });
      return;
    }
    const script = document.createElement("script");
    script.src = HOURS_GUARD_URL;
    script.async = false;
    script.dataset.hoursSourceExplicitGuard = "1";
    script.addEventListener("load", () => callback?.(), { once: true });
    document.head.appendChild(script);
  }

  function loadCore() {
    if (document.readyState === "loading") {
      document.write(`<script src="${HOURS_GUARD_URL}" data-hours-source-explicit-guard="1"><\/script>`);
      document.write(`<script src="${CORE_URL}"><\/script>`);
      return;
    }
    loadHoursGuard(() => {
      const script = document.createElement("script");
      script.src = CORE_URL;
      script.async = false;
      document.head.appendChild(script);
    });
  }

  stopPrematureHoursSubscriptions();

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
    version: "1.3.0",
    getState: () => ({ ...prematureHoursState, errors: prematureHoursState.errors.slice() })
  };

  loadCore();
})();
