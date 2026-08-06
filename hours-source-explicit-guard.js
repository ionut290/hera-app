(() => {
  "use strict";

  if (window.HeraHoursSourceExplicitGuard?.installed) return;

  const state = {
    installed: true,
    version: "2.1.0",
    wrapped: false,
    blockedAutomaticStarts: 0,
    blockedFallbackStarts: 0,
    allowedTrustedStarts: 0,
    assignmentsIntercepted: 0
  };

  function wrapApi(api) {
    if (!api || typeof api.enableHoursSource !== "function") return api;
    const current = api.enableHoursSource;
    if (current.__explicitHoursSourceGuard) {
      state.wrapped = true;
      return api;
    }

    const original = current.bind(api);
    function guardedEnableHoursSource(trigger = null) {
      const trustedUserAction = Boolean(
        trigger &&
        typeof trigger === "object" &&
        "isTrusted" in trigger &&
        trigger.isTrusted === true
      );
      const requestedAutomaticFallback = Boolean(
        trigger &&
        typeof trigger === "object" &&
        trigger.forceSharedCalendarFallback === true
      );

      if (!trustedUserAction) {
        state.blockedAutomaticStarts += 1;
        if (requestedAutomaticFallback) state.blockedFallbackStarts += 1;
        console.debug("[HOURS SOURCE GUARD] avvio automatico delle ore complete bloccato", {
          blockedAutomaticStarts: state.blockedAutomaticStarts,
          blockedFallbackStarts: state.blockedFallbackStarts
        });
        return null;
      }

      state.allowedTrustedStarts += 1;
      return original(trigger);
    }

    guardedEnableHoursSource.__explicitHoursSourceGuard = true;
    guardedEnableHoursSource.__original = current;
    api.enableHoursSource = guardedEnableHoursSource;
    state.wrapped = true;
    return api;
  }

  let storedApi = window.HeraLightStartup;
  const existingDescriptor = Object.getOwnPropertyDescriptor(window, "HeraLightStartup");

  if (!existingDescriptor || existingDescriptor.configurable !== false) {
    Object.defineProperty(window, "HeraLightStartup", {
      configurable: true,
      enumerable: true,
      get() {
        return storedApi;
      },
      set(value) {
        state.assignmentsIntercepted += 1;
        storedApi = wrapApi(value);
      }
    });
    if (storedApi) storedApi = wrapApi(storedApi);
  } else {
    storedApi = wrapApi(storedApi);
  }

  function install(attempt = 0) {
    const api = window.HeraLightStartup;
    if (api && typeof api.enableHoursSource === "function") {
      wrapApi(api);
      return;
    }
    if (attempt < 200) window.setTimeout(() => install(attempt + 1), 25);
  }

  window.HeraHoursSourceExplicitGuard = {
    installed: true,
    version: state.version,
    getState: () => ({ ...state })
  };

  install();
})();
