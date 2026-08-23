(() => {
  "use strict";

  const VERSION = "2.0.0-safe-shim";
  if (window.HeraAdaptiveWorkLearning?.version === VERSION) return;

  const PROFILE_KEY = "heraAdaptiveWorkProfilesV1";
  const ACTIVE_KEY = "heraAdaptiveWorkActiveV1";

  function readJson(key, fallback) {
    try {
      const raw = localStorage.getItem(key);
      return raw ? JSON.parse(raw) : fallback;
    } catch (_) {
      return fallback;
    }
  }

  function refreshStablePanel() {
    try {
      window.HeraRecommendedPlants?.refreshDecorations?.();
    } catch (_) {}
  }

  window.HeraAdaptiveWorkLearning = {
    installed: true,
    version: VERSION,
    safeShim: true,
    getProfiles: () => JSON.parse(JSON.stringify(readJson(PROFILE_KEY, {}))),
    getActiveSession: () => readJson(ACTIVE_KEY, null),
    reset: () => {
      try {
        localStorage.removeItem(PROFILE_KEY);
        localStorage.removeItem(ACTIVE_KEY);
      } catch (_) {}
      refreshStablePanel();
    },
    applyToRecommendedPanel: refreshStablePanel,
    adaptiveMinutes: (_item, _team, fallback) => ({
      minutes: Math.max(0, Math.round(Number(fallback || 0))),
      samples: 0,
      confidence: 0
    }),
    finishActive: () => null
  };
})();
