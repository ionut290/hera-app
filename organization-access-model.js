(function () {
  "use strict";

  const DEFAULT_ORGANIZATION_ID = "varga";
  const VALID_ROLES = Object.freeze(["super_admin", "admin", "operatore"]);

  function normalizeId(value) {
    return String(value || "")
      .trim()
      .toLowerCase()
      .replace(/[^a-z0-9_-]+/g, "-")
      .replace(/^-+|-+$/g, "");
  }

  function normalizeRole(value) {
    const role = String(value || "").trim().toLowerCase();
    return VALID_ROLES.includes(role) ? role : "operatore";
  }

  function normalizeMembership(rawMembership, fallbackOrganizationId) {
    const source = rawMembership && typeof rawMembership === "object" ? rawMembership : {};
    const organizationId = normalizeId(
      source.organizationId || source.organizzazioneId || source.id || fallbackOrganizationId
    ) || DEFAULT_ORGANIZATION_ID;

    return Object.freeze({
      organizationId,
      role: normalizeRole(source.role || source.ruolo),
      active: source.active !== false && source.attivo !== false,
      organizationName: String(source.organizationName || source.nomeOrganizzazione || "").trim()
    });
  }

  function legacyVargaMembership(profile) {
    const source = profile && typeof profile === "object" ? profile : {};
    const legacyRole = source.isAdmin === true
      || source.admin === true
      || String(source.role || source.ruolo || "").toLowerCase() === "admin"
      ? "admin"
      : "operatore";

    return normalizeMembership({
      organizationId: DEFAULT_ORGANIZATION_ID,
      organizationName: "Varga",
      role: legacyRole,
      active: true
    });
  }

  function extractMemberships(profile) {
    const source = profile && typeof profile === "object" ? profile : {};
    const rawMemberships = Array.isArray(source.organizationMemberships)
      ? source.organizationMemberships
      : Array.isArray(source.organizzazioni)
        ? source.organizzazioni
        : [];

    const membershipMap = new Map();
    rawMemberships.forEach((membership) => {
      const normalized = normalizeMembership(membership);
      if (normalized.active) {
        membershipMap.set(normalized.organizationId, normalized);
      }
    });

    // Compatibilità obbligatoria: i profili storici senza membership restano in Varga.
    if (membershipMap.size === 0) {
      const fallback = legacyVargaMembership(source);
      membershipMap.set(fallback.organizationId, fallback);
    }

    return Object.freeze(Array.from(membershipMap.values()));
  }

  function resolveOrganizationAccess(profile, requestedOrganizationId) {
    const memberships = extractMemberships(profile);
    const requestedId = normalizeId(requestedOrganizationId);
    const requestedMembership = memberships.find((item) => item.organizationId === requestedId);
    const selectedMembership = requestedMembership || memberships[0];

    return Object.freeze({
      memberships,
      activeMembership: selectedMembership,
      activeOrganizationId: selectedMembership.organizationId,
      requiresSelection: memberships.length > 1 && !requestedMembership,
      canAccessRequestedOrganization: Boolean(requestedMembership)
    });
  }

  function canManageOrganization(membership) {
    const normalized = normalizeMembership(membership);
    return normalized.active && (normalized.role === "admin" || normalized.role === "super_admin");
  }

  function isSuperAdmin(profile) {
    const source = profile && typeof profile === "object" ? profile : {};
    return source.platformRole === "super_admin"
      || source.ruoloPiattaforma === "super_admin"
      || extractMemberships(source).some((membership) => membership.role === "super_admin");
  }

  window.HeraOrganizationAccessModel = Object.freeze({
    DEFAULT_ORGANIZATION_ID,
    VALID_ROLES,
    normalizeMembership,
    legacyVargaMembership,
    extractMemberships,
    resolveOrganizationAccess,
    canManageOrganization,
    isSuperAdmin
  });
})();
