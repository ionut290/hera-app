/* Three salatura commesse are idempotent and never overwrite KML commesse. */
"use strict";
const fs=require("node:fs"),vm=require("node:vm"),assert=require("node:assert/strict");
const window={},document={readyState:"loading",addEventListener:()=>{}};
vm.runInNewContext(fs.readFileSync("salatura-strada-routes.js","utf8"),{window});
const routes=window.VARGA_SALATURA_STRADA_ROUTES;
assert.equal(routes.length,3);
assert.deepEqual(Array.from(routes,r=>r.name),["Salatura Strada 1","Salatura Strada 2","Salatura Strada 3"]);
assert.deepEqual(Array.from(routes,r=>r.code),["SALATURA-STRADA-1","SALATURA-STRADA-2","SALATURA-STRADA-3"]);
assert.deepEqual(Array.from(routes,r=>r.streets.length),[48,37,17]);
assert.deepEqual(Array.from(routes,r=>r.routeLengthMeters),[19887,29108,24855]);
assert.deepEqual(Array.from(routes,r=>r.sourceName),["Percorso Rosso n° 1","Percorso Blu n° 2","Percorso Giallo n° 3"]);
const docs=new Map([["existing-minor",{nome:"Castel Maggiore 1",codice:"NEVE-CASTEL-MAGGIORE-1",valuable:"KEEP"}]]);
const reference=id=>({
  id,
  async get(){return {exists:docs.has(id),data:()=>docs.get(id)}}
});
const db={
  collection(name){
    assert.equal(name,"commesse");
    return {
      doc:id=>reference(id),
      async get(){return {docs:Array.from(docs,([id,data])=>({id,data:()=>data,ref:reference(id)}))}}
    }
  },
  batch(){
    const changes=[];
    return {set(ref,record){changes.push([ref.id,record])},
      async commit(){for(const [id,record] of changes)docs.set(id,record)}};
  }
};
const firebase={firestore:{FieldValue:{serverTimestamp:()=>"now"}}};
vm.runInNewContext(fs.readFileSync("salatura-strada-integration.js","utf8"),
  {window,document,db,firebase,currentUser:{uid:"admin",email:"test@example.com"},
    canManageData:()=>true,console,Intl,Number,setInterval:()=>0});
const integration=window.VargaSalaturaStrada;
assert.ok(module&&typeof integration.seed==="function");
(async()=>{
  assert.equal(integration.url(0),"https://ionut290.github.io/Varga-piano-neve/?vargaMajorRoute=0");
  assert.equal(integration.url(1),"https://ionut290.github.io/Varga-piano-neve/?vargaMajorRoute=1");
  assert.equal(integration.url(2),"https://ionut290.github.io/Varga-piano-neve/?vargaMajorRoute=2");
  await integration.seed();
  assert.equal(docs.size,4);
  assert.equal(docs.get("existing-minor").valuable,"KEEP");
  routes.forEach((route,i)=>{
    const item=docs.get("normal-salatura-strada-"+(i+1));
    assert.equal(item.nome,route.name);
    assert.equal(item.codice,route.code);
    assert.equal(item.pianoNeveMajorRouteIndex,i);
  });
  await integration.seed();
  assert.equal(docs.size,4,"repeat admin import must not create duplicates");
  const html=fs.readFileSync("index.html","utf8");
  assert.ok(html.includes("snow-castel-maggiore-kml-real.js?v=20261123-viabilita-minore1"));
  assert.ok(html.includes("salatura-strada-routes.js?v=20261123-1"));
  assert.ok(html.includes("salatura-strada-integration.js?v=20261123-1"));
  assert.ok(html.indexOf("salatura-strada-routes.js")<html.indexOf("salatura-strada-integration.js"));
  const sw=fs.readFileSync("sw.js","utf8");
  assert.ok(sw.includes("salatura-strada-integration.js?v=20261123-1"));
  const source=fs.readFileSync("salatura-strada-integration.js","utf8");
  assert.ok(!source.includes('collection("snowRouteGeometry")'),"do not modify VIABILITA MINORE KML");
  assert.ok(!source.includes('collection("pianoNeveStreetWork")'),"do not duplicate original Piano Neve progress");
  console.log("OK: Salatura Strada 1, 2, 3, 102 streets, idempotent commesse, safe deep links.");
})().catch(error=>{console.error(error);process.exitCode=1});
