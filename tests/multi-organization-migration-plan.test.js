"use strict";

const assert = require("assert");
const migration = require("../multi-organization-migration-plan.js");

const legacyOperator = migration.buildUserMigration("u1", { role: "user" });
assert.equal(legacyOperator.action, "migrate");
assert.equal(legacyOperator.writes[0].data.organizations.varga.role, "operatore");
assert.equal(legacyOperator.writes.length, 2);

const legacyAdmin = migration.buildUserMigration("u2", { isAdmin: true });
assert.equal(legacyAdmin.writes[0].data.organizations.varga.role, "admin");

const migrated = migration.buildUserMigration("u3", {
  organizations: { varga: { role: "operatore", status: "active" } }
});
assert.equal(migrated.action, "skip");
assert.equal(migrated.writes.length, 0);

const bootstrap = migration.buildOrganizationBootstrap("super-uid");
assert.equal(bootstrap.length, 3);
assert.equal(bootstrap[0].path, "organizations/varga");
assert.equal(bootstrap[1].path, "platformSuperAdmins/super-uid");

const report = migration.createDryRunReport({
  superAdminUid: "super-uid",
  users: [
    { userId: "u1", profile: { role: "user" } },
    { userId: "u2", profile: { admin: true } },
    { userId: "u3", profile: { organizations: { varga: { role: "operatore" } } } }
  ]
});

assert.equal(report.mode, "dry-run");
assert.equal(report.usersScanned, 3);
assert.equal(report.usersToMigrate, 2);
assert.equal(report.usersSkipped, 1);
assert.equal(report.plannedWriteCount, 7);

console.log("multi-organization-migration-plan: tutti i test superati");
