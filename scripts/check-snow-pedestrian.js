/* Regressione mappa neve professionale: precisione, backup, integrazione. */
"use strict";
const fs=require("node:fs");
const vm=require("node:vm");
const assert=require("node:assert/strict");
const context={window:{}};
vm.runInNewContext(fs.readFileSync("snow-pedestrian-bulk.js","utf8"),context);
const auto=context.window.VargaSnowAutomaticSidewalks;
assert.ok(auto && typeof auto.automaticPart==="function" && typeof auto.run==="function");
const src=[[44.57600,11.36000],[44.57605,11.36000],[44.57610,11.36000]];
const localSidewalk={id:99,points:[[44.57600,11.36001],[44.57605,11.36001],[44.57610,11.36001]]};
const good=auto.automaticPart(src,[localSidewalk]);
assert.ok(good, "accept a continuous close dedicated sidewalk");
assert.equal(good.moved,3, "all points should be shifted");
assert.equal(good.original.length,src.length);
assert.equal(src[0][1],11.36,"the source data must not be mutated");
const wrongSide={id:100,points:[[44.57600,11.36020],[44.57605,11.36020],[44.57610,11.36020]]};
assert.equal(auto.automaticPart(src,[wrongSide]),null, "never snap to distant opposite side");
const distant=[[44.57600,11.36],[44.58000,11.36]];
assert.equal(auto.automaticPart(distant,[localSidewalk]),null,"avoid long straight shortcuts");
assert.equal(auto.automaticPart(src,[]),null,"do not fabricate geometry without reference");
const js=fs.readFileSync("snow-plan-operations.js","utf8");
assert.ok(js.includes("VargaSnowAutomaticSidewalks.run"),"admin automation must be connected");
assert.ok(js.includes("autoPrecisionStarted"),"avoid duplicate session runs");
assert.ok(js.includes("source-original-mymaps-kml"),"original KML geometry must be solid, not automatically treated as uncertain");
assert.ok(js.includes("VargaSnowPedestrianAssist?.detach"),"overlays must be cleaned up");
const html=fs.readFileSync("index.html","utf8");
for(const name of ["snow-castel-maggiore-kml-real.js","snow-pedestrian-bulk.js","snow-pedestrian-assist.js","snow-plan-operations.js"])
  assert.ok(html.includes(name),"script missing: "+name);
assert.ok(html.indexOf("snow-castel-maggiore-kml-real.js")<html.indexOf("snow-pedestrian-bulk.js"));
assert.ok(html.indexOf("snow-pedestrian-bulk.js")<html.indexOf("snow-plan-operations.js"));
assert.ok(fs.readFileSync("snow-pedestrian-assist.js","utf8").includes("if (isSidewalk) for"),
  "snap must be restricted to explicitly mapped sidewalks");
console.log("OK: conservative auto-snap, unchanged originals, no cross-road snapping, admin-only integration.");
