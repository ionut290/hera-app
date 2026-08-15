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

// Phase 27: pure utilities extracted from app.js
(function exposeMoreVargaPureUtils(global) {
  const api = {};
  function sanitizePhone(value) {
    return String(value || "").replace(/[^0-9+]/g, "");
  }
  api.sanitizePhone = sanitizePhone;
  function normalizeMezzoNId(value) {
    return String(value || "").trim().toLowerCase();
  }
  api.normalizeMezzoNId = normalizeMezzoNId;
  function formatPhoneHref(phone = "") {
    return String(phone || "").replace(/[^+\d]/g, "");
  }
  api.formatPhoneHref = formatPhoneHref;
  function getMinutesSinceMidnight(date) {
    return date.getHours() * 60 + date.getMinutes();
  }
  api.getMinutesSinceMidnight = getMinutesSinceMidnight;
  Object.assign(global, api);
  global.VargaPureUtils = Object.freeze({ ...(global.VargaPureUtils || {}), ...api });
})(window);
