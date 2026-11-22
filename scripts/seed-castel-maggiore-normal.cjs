#!/usr/bin/env node
"use strict";
// One-time, idempotent import into ORDINARY commesse, NOT neve_commesse.
// Does not update/delete any existing commessa or any existing impianto.
// Must run under the explicitly named Firebase project with a service account.
const fs=require("node:fs");
const vm=require("node:vm");
const assert=require("node:assert/strict");
const context={window:{}};
vm.runInNewContext(fs.readFileSync("snow-castel-maggiore-catalog.js","utf8"),context,{timeout:5000});
const catalog=context.window.VARGA_SNOW_CM_CATALOG;
assert.ok(catalog && catalog.groups.length===3 && catalog.sections.length===4,"unexpected catalog");
assert.deepEqual(Array.from(catalog.sections,s=>s.stops.length),[14,19,6,6]);
const groups=catalog.groups;
const keys=new Set();
for(const group of groups){
  assert.ok(group.code && group.name && group.indices.length,"incomplete group");
  for(const sectionIdx of group.indices)for(const stop of catalog.sections[sectionIdx].stops){
    assert.ok(!keys.has(stop.key),"duplicate stop "+stop.key);
    keys.add(stop.key);
    assert.ok(Array.isArray(stop.marker) && stop.marker.length===2 && stop.marker.every(Number.isFinite),"missing marker "+stop.key);
    assert.ok(stop.parts?.some(p=>p.coords?.length>=2),"no route lines "+stop.key);
  }
}
assert.equal(keys.size,45,"not 45 stops");
if(process.argv.includes("--validate")){
  console.log("VALIDATED: 3 ordinary commesse, 4 sections, 45 stops; NO Firebase writes.");
  process.exit(0);
}
if(!process.argv.includes("--execute") || process.env.CONFIRM_NORMAL_SNOW_SEED!=="IMPORTA-3-COMMESSE-NORMALI"){
  throw new Error("Production write requires --execute and explicit confirmation flag");
}
const project="hera-app-6cd2b";
if(process.env.GCLOUD_PROJECT && process.env.GCLOUD_PROJECT!==project)throw new Error("Wrong cloud project");
if(process.env.GOOGLE_CLOUD_PROJECT && process.env.GOOGLE_CLOUD_PROJECT!==project)throw new Error("Wrong Google project");
const admin=require("../functions/node_modules/firebase-admin");
admin.initializeApp({credential:admin.credential.applicationDefault(),projectId:project});
const db=admin.firestore(),serverTime=()=>admin.firestore.FieldValue.serverTimestamp();
const normalize=x=>String(x||"").trim().toLowerCase();
const sectionStops=group=>group.indices.flatMap(sectionIndex=>catalog.sections[sectionIndex].stops.map(stop=>({sectionIndex,stop})));
const main=async()=>{
  const snapshot=await db.collection("commesse").get();
  const existing=snapshot.docs;
  const plan=[];
  for(const group of groups){
    const matching=existing.filter(doc=>normalize(doc.data().nome)===normalize(group.name) || normalize(doc.data().codice)===normalize(group.code));
    if(matching.length>1)throw new Error("Ambiguous normal commessa name/code for "+group.name);
    const found=matching[0];
    if(found && normalize(found.data().nome)!==normalize(group.name))throw new Error("Code conflict for "+group.name+"; refusing write");
    const ref=found?.ref||db.collection("commesse").doc("normal-"+group.id);
    const root=await ref.get();
    if(root.exists && !found)throw new Error("Reserved ID belongs to a different normal commessa: "+ref.id);
    const stops=await ref.collection("impianti").get();
    const keyIndex=new Map(),ids=new Set();
    for(const doc of stops.docs){
      ids.add(doc.id);
      const key=doc.data().snowPlanStopKey;
      if(key){if(keyIndex.has(key) && keyIndex.get(key)!==doc.id)throw new Error("Duplicated existing stop key "+key);keyIndex.set(key,doc.id)}
    }
    const pending=[];
    for(const {sectionIndex,stop} of sectionStops(group)){
      const id="piano-neve-"+stop.key;
      if(keyIndex.has(stop.key) || ids.has(id)){
        if(ids.has(id) && !keyIndex.has(stop.key))throw new Error("Conflicting stop document ID "+id);
        continue;
      }
      const first=stop.parts.find(part=>Array.isArray(part.coords) && part.coords.length>=2);
      if(!first)throw new Error("No route for "+stop.key);
      const site=catalog.sections[sectionIndex].name,marker=stop.marker;
      const record={
        denominazione:(site+" · "+stop.number+" · "+stop.note).slice(0,140),
        indirizzo:stop.note,descrizioneVia:stop.note,noteOperative:stop.note,
        comune:"Castel Maggiore",localita:site,tipologiaImpianto:"Via neve",codicePrezzo:"NEVE-STRADA",
        snowRoad:true,snowPlanStopKey:stop.key,snowPlanSection:sectionIndex,
        snowPriority:stop.priority===true,snowPlanSource:catalog.version,
        routePath:first.coords.map(pt=>({lat:Number(pt[0]),lng:Number(pt[1])})),
        roadStatus:"todo",done:false,doneAt:null,doneBy:"",
        latitudine:marker[0],longitudine:marker[1],lat:marker[0],lng:marker[1],
        sortOrder:sectionIndex===3?6+stop.number:stop.number,
        createdAt:serverTime(),createdBy:"seed-castel-maggiore-normal-20261122"
      };
      pending.push({ref:ref.collection("impianti").doc(id),record});
    }
    plan.push({group,ref,createParent:!root.exists,pending,already:sectionStops(group).length-pending.length});
  }
  console.log("PREFLIGHT:",JSON.stringify(plan.map(p=>({name:p.group.name,createParent:p.createParent,newItems:p.pending.length,alreadyPresent:p.already}))));
  for(const p of plan){
    if(p.createParent){
      await p.ref.create({nome:p.group.name,codice:p.group.code,tipo:"main",snowPlan:"castel-maggiore-viabilita-minore",createdAt:serverTime(),createdBy:"seed-castel-maggiore-normal-20261122"});
    }
    if(p.pending.length){
      const batch=db.batch();p.pending.forEach(({ref,record})=>batch.create(ref,record));await batch.commit();
    }
  }
  console.log("COMPLETED:",JSON.stringify(plan.map(p=>({commessa:p.group.name,id:p.ref.id,newItems:p.pending.length,alreadyPresent:p.already}))));
  process.exit(0);
};
main().catch(e=>{console.error("ABORT:",e.message||e);process.exit(1)});
