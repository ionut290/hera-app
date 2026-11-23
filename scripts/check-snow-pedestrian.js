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
const genericFootway={id:101,tags:{highway:"footway"},points:[[44.57600,11.36008],[44.57605,11.36008],[44.57610,11.36008]]};
const genericMatch=auto.automaticPart(src,[genericFootway]);
assert.ok(genericMatch,"nearby footways must be accepted as pedestrian sidewalks");
assert.ok(genericMatch.maxShift>5 && genericMatch.maxShift<12,"snap can shift a blue line several metres toward a mapped pedestrian path");
assert.ok(genericMatch.candidate.length>=src.length,"preserve all original route coverage");
const genericPedestrian={id:102,tags:{highway:"pedestrian"},points:[[44.57600,11.360075],[44.57605,11.360075],[44.57610,11.360075]]};
assert.ok(auto.automaticPart(src,[genericPedestrian]),"pedestrian streets count as snapping references");
const genericPath={id:103,tags:{highway:"path",foot:"yes"},points:[[44.57600,11.36007],[44.57605,11.36007],[44.57610,11.36007]]};
assert.ok(auto.automaticPart(src,[genericPath]),"walkable OSM paths count too");
const wrongSide={id:100,points:[[44.57600,11.36020],[44.57605,11.36020],[44.57610,11.36020]]};
assert.equal(auto.automaticPart(src,[wrongSide]),null, "never snap to distant opposite side");
const distant=[[44.57600,11.36],[44.58000,11.36]];
assert.equal(auto.automaticPart(distant,[localSidewalk]),null,"avoid long straight shortcuts");
assert.equal(auto.automaticPart(src,[]),null,"do not fabricate geometry without reference");
const rules=fs.readFileSync("firestore.rules","utf8");
assert.ok(rules.includes('request.resource.data.keys().hasOnly(["stopKey", "parts", "updatedAt", "updatedBy"])'),
  "guard against adding extra fields forbidden by Firestore rules");
const bulk=fs.readFileSync("snow-pedestrian-bulk.js","utf8");
assert.ok(!bulk.includes("originalParts:"),"no extra Firestore fields are allowed");
assert.ok(!bulk.includes("autoPrecision:"),"do not write forbidden metadata");
const js=fs.readFileSync("snow-plan-operations.js","utf8");
assert.ok(js.includes("VargaSnowRedFootpaths.persist"),"verified OSM red footpaths must be persisted by admin");
assert.ok(js.includes("autoPrecisionStarted"),"avoid duplicate session runs");
assert.ok(js.includes("source-original-mymaps-kml"),"original KML geometry must be solid, not automatically treated as uncertain");
assert.ok(js.includes("VargaSnowPedestrianAssist?.detach"),"overlays must be cleaned up");
const html=fs.readFileSync("index.html","utf8");
for(const name of ["snow-castel-maggiore-kml-real.js","snow-pedestrian-bulk.js","snow-pedestrian-assist.js","snow-plan-operations.js"])
  assert.ok(html.includes(name),"script missing: "+name);
assert.ok(html.indexOf("snow-castel-maggiore-kml-real.js")<html.indexOf("snow-pedestrian-bulk.js"));
assert.ok(html.indexOf("snow-pedestrian-bulk.js")<html.indexOf("snow-plan-operations.js"));
assert.ok(fs.readFileSync("snow-pedestrian-assist.js","utf8").includes("if (isPedestrianPath) for"),
  "snap must also accept pedestrian paths and not only mapped sidewalks");
console.log("OK: pedestrian paths, Firestore rules, original-preserving red-path auto-persistence.");
