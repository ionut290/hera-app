(function (root, factory) {
  const api = factory();
  if (typeof module === "object" && module.exports) module.exports = api;
  if (root) root.HeraMultiOrganizationMigrationPlan = api;
})(typeof window !== "undefined" ? window : globalThis, function () {
  "use strict";

  const DEFAULT_ORGANIZATION_ID = "varga";

  function normalizeId(value) {
    return String(value || "").trim().toLowerCase().replace(/[^a-z0-9_-]+/g, "-").replace(/^-+|-+$/g, "");
  }

  function inferLegacyRole(profile) {
    const role = String(profile?.role || profile?.ruolo || "").trim().toLowerCase();
    if (profile?.isAdmin === true || profile?.admin === true || role === "admin" || role === "amministratore") return "admin";
    return "operatore";
  }

  function buildUserMigration(userId, profile) {
    if (!userId) throw new Error("userId mancante");
    const existingOrganizations = profile?.organizations && typeof profile.organizations === "object" ? profile.organizations : null;
    if (existingOrganizations && Object.keys(existingOrganizations).length > 0) {
      return { userId, action: "skip", reason: "already-migrated", writes: [] };
    }

    const role = inferLegacyRole(profile || {});
    return {
      userId,
      action: "migrate",
      reason: "legacy-user-defaults-to-varga",
      writes: [
        {
          path: `platformUsers/${userId}`,
          merge: true,
          data: {
            defaultOrganizationId: DEFAULT_ORGANIZATION_ID,
            organizations: {
              [DEFAULT_ORGANIZATION_ID]: {
                role,
                status: "active"
              }
            }
          }
        },
        {
          path: `organizations/${DEFAULT_ORGANIZATION_ID}/members/${userId}`,
          merge: true,
          data: {
            userId,
            role,
            status: "active"
          }
        }
      ]
    };
  }

  function buildOrganizationBootstrap(superAdminUid) {
    if (!superAdminUid) throw new Error("superAdminUid mancante");
    return [
      {
        path: `organizations/${DEFAULT_ORGANIZATION_ID}`,
        merge: true,
        data: {
          id: DEFAULT_ORGANIZATION_ID,
          name: "Varga Cantieri",
          status: "active",
          legacyDataMode: true
        }
      },
      {
        path: `platformSuperAdmins/${superAdminUid}`,
        merge: true,
        data: { active: true }
      },
      {
        path: `organizations/${DEFAULT_ORGANIZATION_ID}/members/${superAdminUid}`,
        merge: true,
        data: { userId: superAdminUid, role: "admin", status: "active" }
      }
    ];
  }

  function createDryRunReport(input) {
    const users = Array.isArray(input?.users) ? input.users : [];
    const superAdminUid = String(input?.superAdminUid || "").trim();
    const userPlans = users.map((entry) => buildUserMigration(entry.userId, entry.profile || {}));
    const bootstrapWrites = superAdminUid ? buildOrganizationBootstrap(superAdminUid) : [];
    const plannedWrites = bootstrapWrites.concat(userPlans.flatMap((plan) => plan.writes));

    return {
      mode: "dry-run",
      defaultOrganizationId: DEFAULT_ORGANIZATION_ID,
      usersScanned: users.length,
      usersToMigrate: userPlans.filter((plan) => plan.action === "migrate").length,
      usersSkipped: userPlans.filter((plan) => plan.action === "skip").length,
      plannedWriteCount: plannedWrites.length,
      userPlans,
      writes: plannedWrites
    };
  }

  return Object.freeze({
    DEFAULT_ORGANIZATION_ID,
    normalizeId,
    inferLegacyRole,
    buildUserMigration,
    buildOrganizationBootstrap,
    createDryRunReport
  });
});
