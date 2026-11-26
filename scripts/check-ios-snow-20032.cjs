"use strict";
const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const vm = require("node:vm");
const root = path.resolve(process.argv[2] || path.join(__dirname, ".."));
const read = p => fs.readFileSync(path.join(root,p),"utf8");
const check = p => assert.ok(fs.statSync(path.join(root,p)).size > 0, "Empty asset: "+p);
const required = ["index.html","snow-plan-operations.js","snow-plan-operations.css","app-snow.js","snow-castel-maggiore-catalog.js","snow-castel-maggiore-kml-real.js","salatura-strada-routes.js","salatura-strada-integration.js","salatura-strada-integration.css","piano-neve/index.html","piano-neve/app.js","piano-neve/snow-auth-bridge.js","piano-neve/major-offline-nav.js","piano-neve/major-network.js","piano-neve/major-routes.js"];
for(let i=1;i<=11;i++) required.push("piano-neve/major-network-shard-"+String(i).padStart(2,"0")+".js");
required.forEach(check);
const html=read("index.html");
assert.ok(html.includes("snow-plan-operations.js?v=20261126-ios20032"));
assert.ok(html.includes("salatura-strada-integration.js?v=20261126-ios20032"));
const nested=read("piano-neve/index.html");
for(const m of nested.matchAll(/<script\b[^>]*\bsrc=["']([^"']+)["']/gi)){
  const src=m[1];
  if(!/^(?:https?:|\/\/|data:)/i.test(src))check(path.posix.join("piano-neve",src.split(/[?#]/)[0]));
}
const win={}, page={hidden:false,classList:{contains:()=>false}};
let snowContext=true;
const document={baseURI:"https://localhost/index.html",readyState:"loading",
  getElementById:id=>id==="impianti-page"?page:null,
  body:{classList:{contains:v=>snowContext&&v==="snow-management-context"}},
  addEventListener:()=>{}};
const selectedCommessaId="route",commesseById=new Map();
const context=vm.createContext({window:win,document,URL,selectedCommessaId,commesseById});
for(const f of ["salatura-strada-routes.js","snow-castel-maggiore-catalog.js","salatura-strada-integration.js","snow-plan-operations.js"]){
  vm.runInContext(read(f),context,{filename:f});
}
const big=win.VargaSalaturaStrada,small=win.VargaSnowPlanUI;
assert.ok(big&&small,"Both operation modules must initialize");
let checks=0;
function eq(actual,expected,label){assert.equal(actual,expected,label);checks++;}
for(let i=0;i<3;i++){
  for(const mode of [true,false]){
    snowContext=mode;
    commesseById.set(selectedCommessaId,{nome:"Salatura Strada "+(i+1)});
    eq(big.currentRouteIndex(),i,"salatura name in navigation context");
    eq(small.currentRouteIndex(),-1,"no minor panel");
    commesseById.set(selectedCommessaId,{nome:"Castel Maggiore "+(i+1)});
    eq(small.currentRouteIndex(),i,"minor name in navigation context");
    eq(big.currentRouteIndex(),-1,"no major panel");
  }
  eq(big.resolveCommessa({codice:"salatura strada "+(i+1)},"x"),i,"normalized major code");
  eq(small.resolveCommessa({codice:"NEVE_CM_"+(i+1)},"x"),i,"normalized minor code");
  eq(big.resolveCommessa({pianoNeveService:"salatura-viabilita-maggiore",pianoNeveMajorRouteIndex:i},"x"),i,"major metadata");
  eq(big.resolveCommessa({},"normal-salatura-strada-"+(i+1)),i,"stable major id");
  eq(small.resolveCommessa({},"normal-neve-castel-maggiore-"+(i+1)),i,"stable minor id");
  eq(big.url(i),"https://localhost/piano-neve/index.html?vargaMajorRoute="+i,"native route link");
}
commesseById.set(selectedCommessaId,{nome:"Altra commessa"});
eq(big.currentRouteIndex(),-1,"unknown major");
eq(small.currentRouteIndex(),-1,"unknown minor");
eq(big.resolveCommessa({codice:"SALATURA-STRADA-1",nome:"Salatura Strada 2"},"x"),-1,"ambiguous major");
eq(small.resolveCommessa({codice:"NEVE-CM-1",nome:"Castel Maggiore 2"},"x"),-1,"ambiguous minor");
eq(big.resolveCommessa({codice:"SALATURA-STRADA-1",pianoNeveService:"salatura-viabilita-maggiore",pianoNeveMajorRouteIndex:1},"x"),-1,"conflicting metadata");
commesseById.set(selectedCommessaId,{nome:"Salatura Strada 1"});page.hidden=true;
eq(big.currentRouteIndex(),-1,"hidden page major");
commesseById.set(selectedCommessaId,{nome:"Castel Maggiore 1"});
eq(small.currentRouteIndex(),-1,"hidden page minor");
console.log("PASS: "+checks+" commission routing assertions plus local asset presence checks in "+root+". Device-level GPS, Firebase and navigation tests remain mandatory.");
