(function (root, factory) {
  const api = factory();
  if (typeof module === "object" && module.exports) {
    module.exports = api;
  } else {
    root.HeraSuperAdminOrganizations = api;
  }
})(typeof globalThis !== "undefined" ? globalThis : this, function () {
  "use strict";

  function normalizeOrganization(input) {
    const source = input || {};
    const id = String(source.id || "").trim().toLowerCase();
    if (!id) throw new Error("ID organizzazione mancante.");

    return {
      id,
      name: String(source.name || id).trim(),
      status: source.status === "suspended" ? "suspended" : "active",
      adminUserIds: Array.isArray(source.adminUserIds) ? Array.from(new Set(source.adminUserIds.map(String))) : [],
      createdAt: source.createdAt || null,
      updatedAt: source.updatedAt || null
    };
  }

  function canCreateOrganization(userAccess) {
    return Boolean(userAccess && userAccess.isSuperAdmin === true);
  }

  function canManageOrganization(userAccess, organizationId) {
    if (!userAccess) return false;
    if (userAccess.isSuperAdmin === true) return true;

    const memberships = userAccess.memberships || {};
    const membership = memberships[String(organizationId || "")];
    return Boolean(membership && membership.active !== false && membership.role === "admin");
  }

  function createOrganizationDraft(input, userAccess) {
    if (!canCreateOrganization(userAccess)) {
      throw new Error("Solo il Super Admin può creare organizzazioni.");
    }

    const organization = normalizeOrganization(input);
    return Object.assign({}, organization, {
      status: "active",
      createdByUid: String(userAccess.uid || ""),
      schemaVersion: 1
    });
  }

  return Object.freeze({
    normalizeOrganization,
    canCreateOrganization,
    canManageOrganization,
    createOrganizationDraft
  });
});
