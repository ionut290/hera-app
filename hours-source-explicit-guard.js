(() => {
  "use strict";

  if (window.HeraHoursSourceExplicitGuard?.installed) return;

  const state = {
    installed: true,
    version: "1.0.0",
    wrapped: false,
    blockedAutomaticStarts: 0,
    allowedTrustedStarts: 0,
    allowedFallbackStarts: 0
  };

  function install(attempt = 0) {
    const api = window.HeraLightStartup;
    const current = api?.enableHoursSource;

    if (!api || typeof current !== "function") {
      if (attempt < 120) window.setTimeout(() => install(attempt + 1), 50);
      return;
    }

    if (current.__explicitHoursSourceGuard) {
      state.wrapped = true;
      return;
    }

    const original = current.bind(api);

    function guardedEnableHoursSource(trigger = null) {
      const trustedUserAction = Boolean(
        trigger &&
        typeof trigger === "object" &&
        "isTrusted" in trigger &&
        trigger.isTrusted === true
      );
      const verifiedSharedViewFallback = Boolean(
        trigger &&
        typeof trigger === "object" &&
        trigger.forceSharedCalendarFallback === true
      );

      if (!trustedUserAction && !verifiedSharedViewFallback) {
        state.blockedAutomaticStarts += 1;
        console.debug("[HOURS SOURCE GUARD] caricamento completo ore bloccato perché non richiesto dall’utente", {
          blockedAutomaticStarts: state.blockedAutomaticStarts
        });
        return null;
      }

      if (trustedUserAction) state.allowedTrustedStarts += 1;
      if (verifiedSharedViewFallback) state.allowedFallbackStarts += 1;
      return original(trigger);
    }

    guardedEnableHoursSource.__explicitHoursSourceGuard = true;
    guardedEnableHoursSource.__original = current;
    api.enableHoursSource = guardedEnableHoursSource;
    state.wrapped = true;

    console.info("[HOURS SOURCE GUARD] oreReports completi disponibili solo da azione esplicita o fallback verificato.");
  }

  window.HeraHoursSourceExplicitGuard = {
    ...state,
    getState: () => ({ ...state })
  };

  install();
})();
