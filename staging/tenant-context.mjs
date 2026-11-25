// Pure staging tenant context. No network, storage, or production dependencies.
// UI context is not authorization: Firestore rules must enforce every operation.
export function resolveTenantContext({ uid, claims, tenantId, tenant, member, membership }) {
  if (!uid || claims?.email_verified !== true || claims?.apple_review === true) return null;
  if (typeof tenantId !== "string" || !/^[a-z0-9-]{3,64}$/.test(tenantId)) return null;
  if (tenant?.demo !== true || tenant?.status !== "active") return null;
  if (member?.status !== "active" || membership?.status !== "active") return null;
  if (membership?.tenantId !== tenantId || member.role !== membership.role) return null;
  if (!["admin", "supervisor", "employee"].includes(member.role)) return null;
  return Object.freeze({ uid, tenantId, role: member.role, superAdmin: claims.superAdmin === true });
}
export function tenantCollectionPath(context, collectionName) {
  if (!context || !/^[a-z0-9-]{3,64}$/.test(context.tenantId)) throw new Error("Tenant context required");
  if (!["sites", "members"].includes(collectionName)) throw new Error("Collection not approved for staging");
  return ["tenants", context.tenantId, collectionName];
}
