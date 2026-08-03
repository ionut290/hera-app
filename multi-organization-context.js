(function () {
  "use strict";

  const DEFAULT_ORGANIZATION_ID = "varga";
  const STORAGE_KEY = "hera.activeOrganizationId";

  function normalizeOrganizationId(value) {
    return String(value || "")
      .trim()
      .toLowerCase()
      .replace(/[^a-z0-9_-]+/g, "-")
      .replace(/^-+|-+$/g, "") || DEFAULT_ORGANIZATION_ID;
  }

  function getStoredOrganizationId() {
    try {
      return normalizeOrganizationId(window.localStorage.getItem(STORAGE_KEY));
    } catch (_error) {
      return DEFAULT_ORGANIZATION_ID;
    }
  }

  let activeOrganizationId = getStoredOrganizationId();

  function setActiveOrganizationId(value) {
    const nextOrganizationId = normalizeOrganizationId(value);
    activeOrganizationId = nextOrganizationId;

    try {
      window.localStorage.setItem(STORAGE_KEY, nextOrganizationId);
    } catch (_error) {
      // L'app deve continuare a funzionare anche se localStorage non è disponibile.
    }

    window.dispatchEvent(new CustomEvent("hera:organization-changed", {
      detail: { organizationId: nextOrganizationId }
    }));

    return nextOrganizationId;
  }

  function getActiveOrganizationId() {
    return activeOrganizationId || DEFAULT_ORGANIZATION_ID;
  }

  function isDefaultOrganization() {
    return getActiveOrganizationId() === DEFAULT_ORGANIZATION_ID;
  }

  function organizationCollectionPath(collectionName) {
    const safeCollectionName = String(collectionName || "").trim();
    if (!safeCollectionName) {
      throw new Error("Nome collezione organizzazione mancante.");
    }

    return `organizations/${getActiveOrganizationId()}/${safeCollectionName}`;
  }

  window.HeraOrganizationContext = Object.freeze({
    DEFAULT_ORGANIZATION_ID,
    getActiveOrganizationId,
    setActiveOrganizationId,
    isDefaultOrganization,
    organizationCollectionPath,
    normalizeOrganizationId
  });
})();
