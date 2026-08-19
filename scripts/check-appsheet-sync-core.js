"use strict";

const assert = require("assert");
const core = require("../functions/appsheet-sync-core");

const config = {
  enabled: true,
  appId: "a33fc9cd-0a18-4aa8-b70a-c067c0c6c278",
  tables: {
    commesse: {
      tableName: "COMMESSE",
      keyColumn: "VARGA_ID",
      fields: {
        VARGA_ID: "$id",
        NOME: "nome",
        CODICE: "codice"
      }
    },
    impianti: {
      tableName: "IMPIANTI",
      keyColumn: "VARGA_ID",
      fields: {
        VARGA_ID: "$id",
        COMMESSA_ID: "$parentId",
        STATO: "stato",
        DONE_AT: "doneAt"
      }
    }
  }
};

assert.deepStrictEqual(core.validateConfig(config), { valid: true, reason: "" });
assert.strictEqual(core.shouldSyncWrite(false, true, config.tables.impianti), "Add");
assert.strictEqual(core.shouldSyncWrite(true, true, config.tables.impianti), "Edit");
assert.strictEqual(core.shouldSyncWrite(true, false, config.tables.impianti), "Ignore");

const row = core.buildMappedRow(
  { stato: "FATTO", doneAt: new Date("2026-08-19T16:00:00.000Z") },
  config.tables.impianti,
  { id: "impianto-1", parentId: "commessa-1", path: "commesse/commessa-1/impianti/impianto-1" }
);
assert.deepStrictEqual(row, {
  VARGA_ID: "impianto-1",
  COMMESSA_ID: "commessa-1",
  STATO: "FATTO",
  DONE_AT: "2026-08-19T16:00:00.000Z"
});

assert.strictEqual(
  core.buildActionUrl(config, "IMPIANTI"),
  "https://www.appsheet.com/api/v2/apps/a33fc9cd-0a18-4aa8-b70a-c067c0c6c278/tables/IMPIANTI/Action"
);

console.log("AppSheet sync core: OK");
