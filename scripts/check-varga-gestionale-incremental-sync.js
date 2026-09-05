"use strict";

const assert = require("node:assert/strict");
const sync = require("../functions/varga-gestionale-sync");

const {
  SCHEMA_VERSION,
  ROOT_SOURCES,
  sourceDefinitionForPath,
  changeVersion,
  trackedChangeOperation,
  triggerPathForSource,
  triggerKeyForSource,
  deletionRecord
} = sync.__test;

const snap = (exists, data = {}) => ({ exists, data: () => data });

assert.equal(SCHEMA_VERSION, 2);
assert.equal(ROOT_SOURCES.length, 23);
assert.equal(Object.keys(sync.vgDelta).length, 29);
assert.equal(typeof sync.getVargaGestionaleChanges, "function");
assert.equal(typeof sync.syncVargaGestionaleAt0530, "function");
assert.equal(typeof sync.syncVargaGestionaleAt1600, "function");

assert.equal(sourceDefinitionForPath("impianti/abc").name, "impianti");
assert.equal(sourceDefinitionForPath("commesse/abc/lavorazioni/def").name, "commesse");
assert.equal(sourceDefinitionForPath("commesse/abc/a/def/b/ghi/c/lmn").name, "commesse");
assert.equal(sourceDefinitionForPath("commesse/abc/a/def/b/ghi/c/lmn/d/too-deep"), null);
assert.equal(sourceDefinitionForPath("impianti/abc/lavorazioni/def"), null);
assert.equal(sourceDefinitionForPath("commesse/abc/privateDocuments/def"), null);
assert.equal(sourceDefinitionForPath("operatorPositions/abc"), null);

assert.equal(trackedChangeOperation({ name: "impianti" }, snap(false), snap(true, { value: 1 })), "upsert");
assert.equal(trackedChangeOperation({ name: "impianti" }, snap(true, { value: 1 }), snap(false)), "delete");
const publicDocuments = { name: "documents", filter: (data) => data.visibility !== "personal" };
assert.equal(trackedChangeOperation(publicDocuments, snap(false), snap(true, { visibility: "personal" })), "");
assert.equal(
  trackedChangeOperation(
    publicDocuments,
    snap(true, { visibility: "shared" }),
    snap(true, { visibility: "personal" })
  ),
  "delete"
);

const earlier = changeVersion({ seconds: 100, nanoseconds: 1 }, "impianti/a");
const later = changeVersion({ seconds: 100, nanoseconds: 2 }, "impianti/a");
assert.ok(earlier < later);
assert.notEqual(
  changeVersion({ seconds: 100, nanoseconds: 2 }, "impianti/a"),
  changeVersion({ seconds: 100, nanoseconds: 2 }, "impianti/b")
);

assert.equal(triggerPathForSource("commesse", 0), "commesse/{documentId0}");
assert.equal(
  triggerPathForSource("commesse", 2),
  "commesse/{documentId0}/{collectionId1}/{documentId1}/{collectionId2}/{documentId2}"
);
assert.equal(triggerKeyForSource("neve_commesse", 3), "neveCommesseD3");

assert.deepEqual(deletionRecord({
  sourcePath: "personale/user-1",
  rootCollection: "personale",
  sourceId: "user-1",
  version: later
}), {
  sourcePath: "personale/user-1",
  rootCollection: "personale",
  id: "user-1",
  operation: "delete",
  deleted: true,
  deletedAt: later,
  data: null
});

console.log("Sincronizzazione incrementale Varga Gestionale: controlli superati.");
