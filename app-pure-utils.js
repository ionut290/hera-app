"use strict";
(function exposeVargaPureUtils(global) {
  const api = {};
  function toggleOperatorPositionsVisibility() {
    return;
  }
  api.toggleOperatorPositionsVisibility = toggleOperatorPositionsVisibility;
  function createEmptyCorsoState() {
    return { possiede: false };
  }
  api.createEmptyCorsoState = createEmptyCorsoState;
  function getHomeWorklimateButtonLabel() {
    return "worklimate";
  }
  api.getHomeWorklimateButtonLabel = getHomeWorklimateButtonLabel;
  function normalizeEmail(email) {
    return String(email || "").trim().toLowerCase();
  }
  api.normalizeEmail = normalizeEmail;
  Object.assign(global, api);
  global.VargaPureUtils = Object.freeze({ ...(global.VargaPureUtils || {}), ...api });
})(window);
