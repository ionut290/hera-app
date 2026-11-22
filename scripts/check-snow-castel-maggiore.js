/* Regression checks: three commesse, four sottopercorsi, 45 PDF stops. */
"use strict";
const fs = require("node:fs");
const vm = require("node:vm");
const assert = require("node:assert/strict");
const context = {window:{}};
vm.runInNewContext(fs.readFileSync("snow-castel-maggiore-catalog.js","utf8"),context);
const catalog = context.window.VARGA_SNOW_CM_CATALOG;
assert.ok(catalog, "catalog missing");
assert.equal(catalog.groups.length,3);
assert.equal(catalog.sections.length,4);
assert.deepEqual(Array.from(catalog.sections,x=>x.stops.length),[14,19,6,6]);
assert.deepEqual(Array.from(catalog.groups,g=>Array.from(g.indices)),[[0],[1],[2,3]]);
assert.deepEqual(Array.from(catalog.groups,g=>g.name),["Castel Maggiore 1","Castel Maggiore 2","Castel Maggiore 3"]);
const all=catalog.sections.flatMap(x=>x.stops);
assert.equal(all.length,45);
const keys=all.map(x=>x.key);
assert.equal(new Set(keys).size,keys.length,"duplicate stop keys");
for(const stop of all){
  assert.ok(stop.note.length>4,"missing instruction "+stop.key);
  assert.ok(stop.parts.length>0,"missing geometry "+stop.key);
  assert.ok(stop.marker && stop.marker.length===2,"missing marker "+stop.key);
  for(const part of stop.parts){
    assert.ok(part.coords.length>=2, "route part too short "+stop.key);
    for(const pt of part.coords)assert.ok(pt.length===2 && Number.isFinite(pt[0]) && Number.isFinite(pt[1]) && pt[0]>43 && pt[0]<46 && pt[1]>9 && pt[1]<13, "invalid point "+stop.key);
  }
}
const html=fs.readFileSync("index.html","utf8");
for(const asset of ["snow-castel-maggiore-catalog.js","snow-plan-operations.js","snow-plan-operations.css"])assert.ok(html.includes(asset), "asset not referenced "+asset);
const snow=fs.readFileSync("app-snow.js","utf8");
assert.ok(snow.includes("getSnowRoadParts(impianto)"),"multipath render not installed");
assert.ok(snow.includes("!impianto.snowPlanStopKey"),"official stops must not be GPS auto-completed");
const rules=fs.readFileSync("firestore.rules","utf8");
assert.ok(rules.includes("match /snowRouteWork/{recordId}"),"progress rules missing");
// The requested 3 groups must use normal Varga Cantieri commesse, never the
// dedicated snow-service Firestore collection or its separated page.
const operations=fs.readFileSync("snow-plan-operations.js","utf8");

assert.ok(operations.includes('db.collection("commesse")'),"must import into normal commesse");
assert.ok(!operations.includes('db.collection("neve_commesse")'),"must not import into separate neve_commesse");
assert.ok(operations.includes('id("impianti-page")'),"route map must be on normal commessa page");
assert.ok(operations.includes('id("panel-commesse")'),"admin import button belongs in ordinary commesse management");
assert.ok(!operations.includes('id("snow-service-page")'),"separate snow page must not host these routes");
assert.ok(operations.includes("normalPageVisible()"),"ordinary detail page guard missing");
assert.ok(operations.includes("snowPlanStopKey"),"stable stop IDs required");
console.log("OK: Castel Maggiore 3 commesse, 4 sections, 45 route stops, normal commesse, map and progress checks");
