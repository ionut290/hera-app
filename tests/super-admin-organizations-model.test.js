const assert = require("assert");
const model = require("../super-admin-organizations-model.js");

const superAdmin = {
  uid: "uid-super-admin",
  isSuperAdmin: true,
  memberships: {}
};

const vargaAdmin = {
  uid: "uid-varga-admin",
  isSuperAdmin: false,
  memberships: {
    varga: { role: "admin", active: true }
  }
};

const operator = {
  uid: "uid-operatore",
  isSuperAdmin: false,
  memberships: {
    varga: { role: "operatore", active: true }
  }
};

assert.strictEqual(model.canCreateOrganization(superAdmin), true);
assert.strictEqual(model.canCreateOrganization(vargaAdmin), false);
assert.strictEqual(model.canManageOrganization(superAdmin, "levato"), true);
assert.strictEqual(model.canManageOrganization(vargaAdmin, "varga"), true);
assert.strictEqual(model.canManageOrganization(vargaAdmin, "levato"), false);
assert.strictEqual(model.canManageOrganization(operator, "varga"), false);

const levato = model.createOrganizationDraft({
  id: "LEVATO",
  name: "Levato Cantieri",
  adminUserIds: ["uid-levato", "uid-levato"]
}, superAdmin);

assert.strictEqual(levato.id, "levato");
assert.strictEqual(levato.name, "Levato Cantieri");
assert.strictEqual(levato.status, "active");
assert.deepStrictEqual(levato.adminUserIds, ["uid-levato"]);
assert.strictEqual(levato.createdByUid, "uid-super-admin");

assert.throws(
  () => model.createOrganizationDraft({ id: "rossi" }, vargaAdmin),
  /Solo il Super Admin/
);

console.log("super-admin-organizations-model: tutti i test superati");
