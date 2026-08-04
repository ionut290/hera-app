"use strict";

const assert = require("assert");
const fs = require("fs");

const client = fs.readFileSync("shared-static-views-client.js", "utf8");

assert.match(client, /subscribePersonale = \(\) => subscribe\("personale"\)/);
assert.match(client, /subscribeMezzi = \(\) => subscribe\("mezzi"\)/);
assert.match(client, /collection\("sharedStaticViews"\)\.doc\("registri__corrente"\)/);
assert.strictEqual((client.match(/\.onSnapshot\(/g) || []).length, 1);
assert.match(client, /Array\.isArray\(view\.payload\.personale\)/);
assert.match(client, /Array\.isArray\(view\.payload\.mezzi\)/);
assert.match(client, /documento-mancante/);
assert.match(client, /payload-non-valido/);
assert.match(client, /const FALLBACK_MS = 9000/);
assert.match(client, /timeout/);
assert.match(client, /getRecords: \(kind\)/);
assert.doesNotMatch(client, /subscribeSquadre\s*=/);
assert.doesNotMatch(client, /subscribeHoursStats\s*=/);
assert.doesNotMatch(client, /subscribeHoursApprovals\s*=/);
assert.doesNotMatch(client, /squadreStorico|squadreCommesse|oreReports|oreApprovalRequests/);

console.log("Shared registries client checks passed.");
