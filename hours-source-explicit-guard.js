(() => {
  "use strict";

  if (window.HeraHoursSourceExplicitGuard?.installed) return;

  const state = {
    installed: true,
    version: "2.1.1",
    wrapped: false,
    blockedAutomaticStarts: 0,
    blockedFallbackStarts: 0,
    allowedTrustedStarts: 0,
    assignmentsIntercepted: 0,
    staticCalendarHoursReconciled: 0
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

  function reconcileStaticCalendarHoursState() {
    try {
      const startupState = window.HeraLightStartup?.getState?.();
      if (!startupState?.calendarSharedViewActive || startupState.hoursSourceEnabled) return false;
      if (typeof hoursReportsLoaded === "undefined" || hoursReportsLoaded !== true) return false;
      if (typeof allHoursApprovalRequests === "undefined" || !Array.isArray(allHoursApprovalRequests)) return false;

      if (typeof hoursApprovalsLoaded !== "undefined" && hoursApprovalsLoaded !== true) {
        hoursApprovalsLoaded = true;
        if (typeof hoursApprovalRequests !== "undefined") {
          hoursApprovalRequests = allHoursApprovalRequests;
        }
        state.staticCalendarHoursReconciled += 1;
        console.debug("[HOURS SOURCE GUARD] stato richieste ore allineato dalla vista calendario condivisa");
        if (typeof renderHoursApprovalRequests === "function") renderHoursApprovalRequests();
        if (typeof renderSquadre === "function") renderSquadre();
      }
      return typeof hoursApprovalsLoaded !== "undefined" && hoursApprovalsLoaded === true;
    } catch (error) {
      console.debug("[HOURS SOURCE GUARD] allineamento calendario ore rinviato", error);
      return false;
    }
  }

  function startStaticCalendarHoursReconciliation(attempt = 0) {
    if (reconcileStaticCalendarHoursState()) return;
    if (attempt < 240) {
      window.setTimeout(() => startStaticCalendarHoursReconciliation(attempt + 1), 250);
    }
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
        startStaticCalendarHoursReconciliation();
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
      startStaticCalendarHoursReconciliation();
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
