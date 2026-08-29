(() => {
  "use strict";

  const CORE_URL = "./shared-static-views-client-core.js?v=20260806-explicit-hours-v4-add-hours-rewrite1";
  const HOURS_GUARD_URL = "./hours-source-explicit-guard.js?v=20260806c";
  const api = window.HeraSharedStaticViews;
  const prematureHoursState = {
    stoppedHoursListener: false,
    stoppedApprovalsListener: false,
    completeCalendarViews: 0,
    legacyCalendarViewsAccepted: 0,
    invalidCalendarViewsIgnored: 0,
    automaticFallbacksBlocked: 0,
    errors: []
  };

  // Blocca immediatamente le vecchie sottoscrizioni complete di personale e
  // mezzi mentre viene caricato il client delle viste condivise. Su dispositivi
  // veloci l'autenticazione può completarsi prima del caricamento del core:
  // senza questa guardia app.js apre i listener originali e scarica entrambe le
  // collezioni complete. Le chiamate vengono accodate e riprodotte dal core.
  const registryStartupGate = (() => {
    const originals = {
      personale: typeof subscribePersonale === "function" ? subscribePersonale : null,
      mezzi: typeof subscribeMezzi === "function" ? subscribeMezzi : null
    };
    const pending = new Map([
      ["personale", []],
      ["mezzi", []]
    ]);
    const state = {
      installed: false,
      released: false,
      queuedCalls: 0,
      replayedCalls: 0,
      fallbackCalls: 0,
      errors: []
    };
    let fallbackTimer = null;

    function stopNative(kind) {
      try {
        const stop = kind === "personale"
          ? (typeof stopPersonaleSubscription === "function" ? stopPersonaleSubscription : null)
          : (typeof stopMezziSubscription === "function" ? stopMezziSubscription : null);
        if (stop) {
          stop();
          return;
        }

        const unsubscribe = kind === "personale"
          ? (typeof unsubscribePersonale === "function" ? unsubscribePersonale : null)
          : (typeof unsubscribeMezzi === "function" ? unsubscribeMezzi : null);
        unsubscribe?.();
        if (kind === "personale" && typeof unsubscribePersonale !== "undefined") unsubscribePersonale = null;
        if (kind === "mezzi" && typeof unsubscribeMezzi !== "undefined") unsubscribeMezzi = null;
      } catch (error) {
        state.errors.push(`${kind}: ${error?.message || error}`);
      }
    }

    function settle(kind, value) {
      const resolvers = pending.get(kind) || [];
      pending.set(kind, []);
      resolvers.forEach((resolve) => resolve(value));
    }

    function queue(kind) {
      state.queuedCalls += 1;
      stopNative(kind);
      return new Promise((resolve) => {
        const resolvers = pending.get(kind) || [];
        resolvers.push(resolve);
        pending.set(kind, resolvers);
      });
    }

    function restoreAndReplay() {
      if (state.released) return;
      state.released = true;
      state.fallbackCalls += 1;
      if (originals.personale) subscribePersonale = originals.personale;
      if (originals.mezzi) subscribeMezzi = originals.mezzi;

      for (const kind of ["personale", "mezzi"]) {
        const resolvers = pending.get(kind) || [];
        if (!resolvers.length) continue;
        const original = originals[kind];
        if (!original) {
          settle(kind, false);
          continue;
        }
        Promise.resolve(original())
          .then((value) => settle(kind, value))
          .catch((error) => {
            state.errors.push(`${kind}-fallback: ${error?.message || error}`);
            settle(kind, false);
          });
      }
    }

    function install() {
      stopNative("personale");
      stopNative("mezzi");
      if (originals.personale) subscribePersonale = () => queue("personale");
      if (originals.mezzi) subscribeMezzi = () => queue("mezzi");
      state.installed = Boolean(originals.personale || originals.mezzi);

      if (state.installed && typeof window.setTimeout === "function") {
        fallbackTimer = window.setTimeout(restoreAndReplay, 12000);
      }
      return state.installed;
    }

    function release(sharedSubscriptions = {}) {
      if (state.released) return;
      state.released = true;
      if (fallbackTimer && typeof window.clearTimeout === "function") {
        window.clearTimeout(fallbackTimer);
      }
      fallbackTimer = null;

      if (typeof sharedSubscriptions.personale === "function") {
        subscribePersonale = sharedSubscriptions.personale;
      }
      if (typeof sharedSubscriptions.mezzi === "function") {
        subscribeMezzi = sharedSubscriptions.mezzi;
      }

      for (const kind of ["personale", "mezzi"]) {
        const resolvers = pending.get(kind) || [];
        if (!resolvers.length) continue;
        const shared = sharedSubscriptions[kind];
        if (typeof shared !== "function") {
          settle(kind, false);
          continue;
        }
        state.replayedCalls += 1;
        Promise.resolve(shared())
          .then((value) => settle(kind, value))
          .catch((error) => {
            state.errors.push(`${kind}-shared: ${error?.message || error}`);
            settle(kind, false);
          });
      }
    }

    return {
      installed: true,
      version: "1.0.0",
      originals,
      install,
      release,
      failOpen: restoreAndReplay,
      getState: () => ({ ...state, errors: state.errors.slice() })
    };
  })();

  window.HeraRegistryStartupGate = registryStartupGate;
  registryStartupGate.install();

  // app.js viene eseguito prima di questo client e può aprire i listener completi
  // delle ore durante l'avvio. Li chiudiamo prima che consegnino l'intero storico;
  // la vista mensile condivisa alimenterà Home e calendario. La sorgente completa
  // potrà essere riaperta soltanto da una vera azione dell'utente.
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

  function isCompleteCalendarView(view) {
    return Boolean(
      view &&
      Number(view.schemaVersion || 0) >= 2 &&
      view.completeRecords === true &&
      view.payload &&
      Number(view.payload.schemaVersion || 0) >= 2 &&
      view.payload.completeRecords === true &&
      Array.isArray(view.payload.reports)
    );
  }

  function hasUsableLegacyCalendar(view) {
    return Boolean(
      view &&
      view.payload &&
      Array.isArray(view.payload.reports)
    );
  }

  function markStaticApprovalsReady() {
    try { hoursApprovalsLoaded = true; } catch (_) {}
  }

  function syncStaticApprovalList() {
    try { hoursApprovalRequests = allHoursApprovalRequests; } catch (_) {}
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
      if (isCompleteCalendarView(view)) {
        prematureHoursState.completeCalendarViews += 1;
        markStaticApprovalsReady();
        callback(view, metadata);
        syncStaticApprovalList();
        return;
      }

      // Le versioni precedenti salvavano nel telefono una vista ridotta senza
      // schemaVersion/completeRecords. La usiamo solo come dato temporaneo e
      // continuiamo ad attendere il documento completo dal listener Firestore.
      // Non deve mai aprire automaticamente oreReports/oreApprovalRequests.
      if (hasUsableLegacyCalendar(view)) {
        prematureHoursState.legacyCalendarViewsAccepted += 1;
        prematureHoursState.automaticFallbacksBlocked += 1;
        markStaticApprovalsReady();
        callback(view, { ...metadata, legacyReduced: true });
        syncStaticApprovalList();
        console.warn("[SAFE CALENDAR GUARD] Vista calendario legacy accettata temporaneamente; fallback completo bloccato.", {
          key,
          source: metadata.source || "sconosciuta",
          reports: view.payload.reports.length
        });
        return;
      }

      prematureHoursState.invalidCalendarViewsIgnored += 1;
      prematureHoursState.automaticFallbacksBlocked += 1;
      console.warn("[SAFE CALENDAR GUARD] Vista calendario non valida ignorata; ore complete disponibili solo da Gestione ore.", {
        key,
        source: metadata.source || "sconosciuta",
        schemaVersion: view?.schemaVersion,
        completeRecords: view?.completeRecords,
        payloadSchemaVersion: view?.payload?.schemaVersion,
        payloadCompleteRecords: view?.payload?.completeRecords
      });
    }, ...rest);
  };

  window.HeraSafeCalendarGuard = {
    installed: true,
    version: "1.4.0",
    getState: () => ({ ...prematureHoursState, errors: prematureHoursState.errors.slice() })
  };

  loadCore();
})();
