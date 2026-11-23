/* Castel Maggiore: reconstituted routes must follow stored OSM pedestrian geometry. */
"use strict";
const fs=require("node:fs");
const vm=require("node:vm");
const assert=require("node:assert/strict");
const W={};
const context={window:W};
function execute(path){vm.runInNewContext(fs.readFileSync(path,"utf8"),context,{filename:path});}
execute("snow-castel-maggiore-catalog.js");
const original=W.VARGA_SNOW_CM_CATALOG.sections.flatMap(s=>s.stops)
  .map(s=>({key:s.key,note:s.note,priority:s.priority}));
execute("snow-footway-source-snapshot.js");
const allRefs=Object.values(W.VARGA_SNOW_CM_PEDESTRIAN_REFERENCES).flat().filter(p=>p.source==="osm");
assert.ok(allRefs.length>250,"original OSM pedestrian geometry snapshot missing");
execute("snow-castel-maggiore-kml-real.js");
assert.equal(W.VARGA_SNOW_CM_KML_SOURCE.appliedStops,43);
execute("snow-red-footpaths-rebuild.js");
const rebuilt=W.VargaSnowRedFootpaths;
assert.ok(rebuilt?.results,"reconstruction module missing");
const stops=W.VARGA_SNOW_CM_CATALOG.sections.flatMap(s=>s.stops);
assert.equal(stops.length,45);
assert.deepEqual(stops.map(s=>s.key).join(","),original.map(s=>s.key).join(","));
assert.deepEqual(stops.map(s=>s.note).join("|"),original.map(s=>s.note).join("|"));
assert.deepEqual(stops.map(s=>String(s.priority)).join("|"),original.map(s=>String(s.priority)).join("|"));
assert.ok(rebuilt.results.rebuilt>=25,"too few routes reconstructed on actual OSM vectors");
assert.ok(rebuilt.results.unmatched>0,"must not pretend unverified routes have a red OSM footpath");
assert.ok(rebuilt.results.complete>=1,"at least one original route must match OSM in full");
assert.equal(rebuilt.results.rebuilt+rebuilt.results.unmatched,45);
assert.equal(rebuilt.results.partial+rebuilt.results.complete,rebuilt.results.rebuilt);
const metres=(a,b)=>Math.hypot((a[0]-b[0])*111195,(a[1]-b[1])*111195*Math.cos(a[0]*Math.PI/180));
function project(p,a,b){
 const la=Math.PI/180*p[0],q=Math.cos(la);
 const xy=x=>[x[1]*q,x[0]];
 const P=xy(p),A=xy(a),B=xy(b),dx=B[0]-A[0],dy=B[1]-A[1],denom=dx*dx+dy*dy;
 const t=denom?Math.max(0,Math.min(1,((P[0]-A[0])*dx+(P[1]-A[1])*dy)/denom)):0;
 return [a[0]+(b[0]-a[0])*t,a[1]+(b[1]-a[1])*t];
}
const segments=[];
for(const p of allRefs)for(let i=1;i<p.coords.length;i++)
  segments.push([p.coords[i-1],p.coords[i]]);
let checked=0, problems=0;
for(const stop of stops){
 if(!stop.redFootwayRebuild)continue;
 const parts=stop.parts;
 assert.ok(parts.every(p=>p.coords.length>=2),"empty rebuilt segment in "+stop.key);
 for(const part of parts){
   if(!part.matched)continue;
   for(const point of part.coords){
     if(checked>=300)break;
     const shortest=segments.reduce((best,s)=>Math.min(best,metres(point,project(point,s[0],s[1]))),Infinity);
     if(shortest>.25)problems++;
     checked++;
   }
 }
}
assert.ok(checked>100,"native walkway geometry coverage test insufficient");
assert.equal(problems,0,"one or more rebuilt vertices are not on OSM pedestrian geometry");
// Manual edits must not be replaced by an admin automation.
const ops=fs.readFileSync("snow-plan-operations.js","utf8");
assert.ok(ops.includes("effectiveSnowParts(stop)"),"manual geometry precedence is missing");
assert.ok(ops.includes("VargaSnowRedFootpaths.persist"),"permanent saving is not wired");
assert.ok(!ops.includes("VargaSnowAutomaticSidewalks.run"),"old broad street snapping must be disabled");
const idx=fs.readFileSync("index.html","utf8");
const sourceOrder=["snow-castel-maggiore-catalog.js","snow-footway-source-snapshot.js",
  "snow-castel-maggiore-kml-real.js","snow-red-footpaths-rebuild.js","app-snow.js"];
let prev=-1;
for(const filename of sourceOrder){
  const at=idx.indexOf(filename);
  assert.ok(at>prev,"wrong import order: "+filename);
  prev=at;
}
console.log("OK: "+rebuilt.results.rebuilt+" routes with pedestrian OSM segments, "+rebuilt.results.complete+
  " entirely reconstructed, "+rebuilt.results.unmatched+" deliberately left unverified.");
