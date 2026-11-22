#!/usr/bin/env node
"use strict";
/* Review-only PDF audit patch; changes only PDF source descriptions, red priority
 * indicators and squad personnel/equipment on 3 specifically pinned normal
 * Varga Cantieri commesse; never changes geometry, completion or assignments.
 * All reads occur before transaction writes; abort on any data drift. */
const fs=require("node:fs"),vm=require("node:vm"),assert=require("node:assert/strict");
const context={window:{}};
vm.runInNewContext(fs.readFileSync("snow-castel-maggiore-catalog.js","utf8"),context,{timeout:5000});
const catalog=context.window.VARGA_SNOW_CM_CATALOG;
const legacy=JSON.parse(fs.readFileSync("scripts/snow-castel-maggiore-legacy-metadata.json","utf8"));
const sectionStops=g=>g.indices.flatMap(i=>catalog.sections[i].stops.map(stop=>({...stop,sectionIndex:i,sectionName:catalog.sections[i].name})));
assert.equal(catalog.sections.length,4);
assert.deepEqual(Array.from(catalog.sections,s=>s.stops.length),[14,19,6,6]);
const expectedPriority=[[3,4,5,6,7],[2,5,6,7,9],[2,3,5],[]];
assert.deepEqual(Array.from(catalog.sections,s=>Array.from(s.stops.filter(x=>x.priority),x=>x.number)),expectedPriority);
assert.equal(Object.keys(legacy.stops).length,45);
assert.equal(catalog.groups.length,3);
for(const g of catalog.groups)assert.ok(g.equipment?.length && g.people);
if(process.argv.includes("--validate")){
 console.log("PDF fidelity metadata verified: 45 descriptions, 13 red priority items, three squad sizes, equipment; zero writes.");
 process.exit(0);
}
if(!process.argv.includes("--execute") || process.env.CONFIRM_PDF_AUDIT_PATCH!=="CORREGGI-PDF-ALLEGATO-1")
 throw new Error("Refusing database writes without --execute and confirmation");
const project="hera-app-6cd2b";
if(process.env.GOOGLE_CLOUD_PROJECT&&process.env.GOOGLE_CLOUD_PROJECT!==project)throw new Error("Project mismatch");
const admin=require("../functions/node_modules/firebase-admin");
admin.initializeApp({credential:admin.credential.applicationDefault(),projectId:project});
const db=admin.firestore();
const exp=(o)=>JSON.stringify(o);
const patch=async()=>{
 const results=await db.runTransaction(async(tx)=>{
  const docs=[],roots=[];
  for(const group of catalog.groups){
   const ref=db.collection("commesse").doc("normal-"+group.id);
   roots.push({group,ref});
   for(const stop of sectionStops(group)){
    docs.push({stop,group,ref:ref.collection("impianti").doc("piano-neve-"+stop.key)});
   }
  }
  // Firestore requires all transaction reads before writes.
  const rootSnaps=[],docSnaps=[];
  for(const r of roots)rootSnaps.push(await tx.get(r.ref));
  for(const r of docs)docSnaps.push(await tx.get(r.ref));
  let corrected=0,already=0,rootUpdates=0;
  for(let i=0;i<roots.length;i++){
   const {group}=roots[i],snap=rootSnaps[i];
   if(!snap.exists)throw new Error("Missing normal commessa: "+group.name);
   const v=snap.data();
   if(v.nome!==group.name||v.codice!==group.code||v.snowPlan!=="castel-maggiore-viabilita-minore")throw new Error("Mismatched normal commessa: "+group.name);
   if(v.createdBy!=="seed-castel-maggiore-normal-20261122")throw new Error("Expected imported source commessa not found: "+group.name);
   const desired={personalePrevisto:group.people,attrezzaturaPrevista:group.equipment,
    notaAttrezzaturaCondivisa:"1 trattore leggero con spandisale a servizio dei 3 percorsi",
    snowPlanPdfReference:"Allegato 1 - Smaltimento neve della viabilita minore",snowPlanPdfReviewed:"2026-11-22"};
   for(const [k,val] of Object.entries(desired)){
    if(v[k]!==undefined && exp(v[k])!==exp(val))throw new Error("Commessa metadata edited since import: "+group.name+" "+k);
   }
   if(Object.keys(desired).some(k=>v[k]===undefined)){tx.update(snap.ref,desired);rootUpdates++}
  }
  for(let i=0;i<docs.length;i++){
   const {stop,group}=docs[i],snap=docSnaps[i];
   if(!snap.exists)throw new Error("Missing imported stop "+stop.key);
   const v=snap.data(),old=legacy.stops[stop.key];
   if(!old)throw new Error("Legacy key absent "+stop.key);
   if(v.snowPlanStopKey!==stop.key || v.createdBy!=="seed-castel-maggiore-normal-20261122")throw new Error("Non-original stop "+stop.key);
   const desiredName=(stop.sectionName+" · "+stop.number+" · "+stop.note).slice(0,140);
   const oldSection=stop.sectionIndex===3?"1 Maggio":stop.sectionName;
   const originalName=(oldSection+" · "+stop.number+" · "+old.note).slice(0,140);
   if(v.snowPlanSource===catalog.version){
     if(v.indirizzo!==stop.note || v.descrizioneVia!==stop.note || v.noteOperative!==stop.note || v.snowPriority!==stop.priority)
       throw new Error("PDF-reviewed stop unexpectedly changed "+stop.key);
     already++;continue;
   }
   if(v.snowPlanSource!==legacy.version || v.indirizzo!==old.note || v.descrizioneVia!==old.note || v.noteOperative!==old.note ||
      v.snowPriority!==old.priority || v.denominazione!==originalName)throw new Error("Stop has new edits; refusing overwrite "+stop.key);
   tx.update(snap.ref,{
    denominazione:desiredName,
    indirizzo:stop.note,descrizioneVia:stop.note,noteOperative:stop.note,
    snowPriority:stop.priority,snowPlanSource:catalog.version,
    snowPlanPdfReference:"Allegato 1 - Smaltimento neve della viabilita minore",
    snowPlanGeometryWarning:"Geometrie OSM/DBTR indicative; verificare lato marciapiede e accessi sul posto"
   });
   corrected++;
  }
  return {normalCommesse:roots.length,items:docs.length,corrected,already,rootUpdates,
    priorityByGroup:catalog.sections.map(s=>({area:s.name,numbers:s.stops.filter(x=>x.priority).map(x=>x.number)}))};
 });
 console.log("PDF_AUDIT_COMPLETED:",JSON.stringify(results));
 process.exit(0);
};
patch().catch(e=>{console.error("PDF_AUDIT_ABORT:",e.message||e);process.exit(1)});
