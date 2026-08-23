(() => {
  "use strict";

  const VERSION = "2.0.0-safe-shim";
  if (window.HeraEquipmentAdvisor?.version === VERSION) return;

  const emptyTeam = Object.freeze({
    teamSize: 2,
    equipmentText: "",
    hasBigTrincia: false,
    hasSmallTrincia: false,
    hasTrincia: false,
    hasDaily: false,
    hasDecespugliatore: false,
    hasSoffiatore: false,
    hasTagliasiepe: false,
    hasMotosega: false,
    hasSpazzatrice: false,
    hasPiattaforma: false,
    bigCode: "",
    smallCode: "",
    dailyCode: ""
  });

  function getTeamInfo() {
    try {
      return window.HeraRecommendedPlants?.getTeamInfo?.() || { ...emptyTeam };
    } catch (_) {
      return { ...emptyTeam };
    }
  }

  function refresh() {
    try {
      window.HeraRecommendedPlants?.refreshDecorations?.();
    } catch (_) {}
  }

  window.HeraEquipmentAdvisor = {
    installed: true,
    version: VERSION,
    safeShim: true,
    getTeamInfo,
    recommendations: () => [],
    getProfiles: () => ({}),
    refresh
  };
})();
