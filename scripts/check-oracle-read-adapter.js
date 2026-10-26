const assert = require("node:assert");
const fs = require("node:fs");
const vm = require("node:vm");

const source = fs.readFileSync("oracle-read-adapter.js", "utf8");
assert.ok(!/\.collection\s*\(/.test(source), "L'adattatore non deve leggere Firestore.");
assert.ok(!/\.onSnapshot\s*\(/.test(source), "L'adattatore non deve creare listener Firestore.");
assert.ok(!/\b(?:set|add|update|delete)Doc\s*\(/.test(source), "L'adattatore deve restare in sola lettura.");

let fetchCalls = 0;
const sandbox = {
  window: {
    __HERA_ORACLE_PILOT__: { enabled: false },
    firebase: { auth: () => ({ currentUser: null }) }
  },
  fetch: async () => { fetchCalls += 1; throw new Error("unexpected fetch"); },
  Object, String, Number, Set, Error, Promise, encodeURIComponent
};
vm.runInNewContext(source, sandbox, { filename: "oracle-read-adapter.js" });
assert.equal(sandbox.window.HeraOracleReadAdapter.installed, true);
assert.equal(sandbox.window.HeraOracleReadAdapter.enabled, false);
assert.equal(sandbox.window.HeraOracleReadAdapter.mode, "read-only");
assert.equal(fetchCalls, 0, "L'installazione non deve avviare letture automatiche.");

sandbox.window.HeraOracleReadAdapter.loadCatalog().then(
  () => assert.fail("Il pilota disattivato non deve caricare dati."),
  (error) => {
    assert.equal(error.message, "oracle_pilot_disabled");
    assert.equal(fetchCalls, 0);
    console.log("Oracle read adapter: disattivato per default, nessuna lettura/listener/scrittura Firebase.");
  }
);
