const fs = require("node:fs");
const path = require("node:path");
const vm = require("node:vm");
const assert = require("node:assert/strict");

const source = fs.readFileSync(path.join(__dirname, "..", "organization-access-model.js"), "utf8");
const sandbox = { window: {} };
vm.runInNewContext(source, sandbox, { filename: "organization-access-model.js" });

const model = sandbox.window.HeraOrganizationAccessModel;

assert.ok(model, "Il modello deve essere esposto su window");

const legacyOperator = model.extractMemberships({ role: "user" });
assert.equal(legacyOperator.length, 1);
assert.equal(legacyOperator[0].organizationId, "varga");
assert.equal(legacyOperator[0].role, "operatore");

const legacyAdmin = model.extractMemberships({ isAdmin: true });
assert.equal(legacyAdmin[0].organizationId, "varga");
assert.equal(legacyAdmin[0].role, "admin");

const multi = model.resolveOrganizationAccess({
  organizationMemberships: [
    { organizationId: "varga", role: "operatore", active: true },
    { organizationId: "levato", role: "admin", active: true }
  ]
}, "levato");
assert.equal(multi.memberships.length, 2);
assert.equal(multi.activeOrganizationId, "levato");
assert.equal(multi.canAccessRequestedOrganization, true);
assert.equal(multi.requiresSelection, false);

const inaccessible = model.resolveOrganizationAccess({
  organizationMemberships: [{ organizationId: "varga", role: "operatore" }]
}, "levato");
assert.equal(inaccessible.activeOrganizationId, "varga");
assert.equal(inaccessible.canAccessRequestedOrganization, false);

const selectionRequired = model.resolveOrganizationAccess({
  organizationMemberships: [
    { organizationId: "varga", role: "operatore" },
    { organizationId: "levato", role: "operatore" }
  ]
});
assert.equal(selectionRequired.requiresSelection, true);

assert.equal(model.canManageOrganization({ organizationId: "varga", role: "admin" }), true);
assert.equal(model.canManageOrganization({ organizationId: "varga", role: "operatore" }), false);
assert.equal(model.isSuperAdmin({ platformRole: "super_admin" }), true);

console.log("organization-access-model: tutti i test superati");
