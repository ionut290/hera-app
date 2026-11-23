/* Unified Salatura/Piano Neve: original offline graph and true shared auth. */
"use strict";
const fs=require("node:fs"),vm=require("node:vm"),assert=require("node:assert/strict"),path=require("node:path");
const root=path.join(__dirname,"..","piano-neve");
const idx=fs.readFileSync(path.join(root,"index.html"),"utf8");
assert.ok(idx.includes("snow-auth-bridge.js"),"Auth bridge must replace separate Piano Neve login");
assert.ok(!idx.includes('id="loginEmail"')&&!idx.includes('id="loginGoogleButton"'),"Second login controls must be absent");
assert.ok(!idx.includes("https://ionut290.github.io/Varga-piano-neve"),"No dependency on standalone app origin");
const assets=[...idx.matchAll(/(?:src|href)="([^"]+)"/g)].map(m=>m[1]);
for(const asset of assets){
 if(/^(https?:|data:|#)/.test(asset))continue;
 const local=asset.split("?")[0];
 assert.ok(fs.existsSync(path.join(root,local)),"Missing bundled Piano Neve asset: "+local);
}
const graph={window:{}};
for(const file of ["major-network.js",...Array.from({length:11},(_,i)=>"major-network-shard-"+String(i+1).padStart(2,"0")+".js")]){
 const s=fs.readFileSync(path.join(root,file),"utf8");
 vm.runInNewContext(s,graph,{filename:file,timeout:12000});
}
assert.equal(graph.window.MAJOR_OFFLINE.version,3);
assert.equal(graph.window.MAJOR_OFFLINE.roads.length,1547,"Original road geometry must survive split");
assert.equal(graph.window.MAJOR_OFFLINE.routingRoads.length,20115,"All offline transfer road edges must survive split");
assert.equal(new Set(Array.from(graph.window.MAJOR_OFFLINE.routingRoads,r=>r.id)).size,20115,"Offline road IDs must remain unique");
const app=fs.readFileSync(path.join(root,"app.js"),"utf8");
new Function(app);
assert.ok(app.includes('if(window.__VARGA_CANTIERI_INTEGRATED__)return'),"Nested service worker must not replace Varga controller");
assert.ok(app.includes('get('+"'vargaMajorRoute'"+')'),"Existing major-route selection must be preserved");
const authSource=fs.readFileSync(path.join(root,"snow-auth-bridge.js"),"utf8");
new Function(authSource);
assert.ok(!authSource.includes("signInWithEmailAndPassword")&&!authSource.includes("signInWithPopup"),"No second login API call is allowed");
assert.ok(!authSource.includes("await auth.signOut()"),"Rejecting a snow account must not log out the gestionale");
const handlers={};
const classes=new Set(["auth-locked"]);
const loginScreen={hidden:false},snowApp={hidden:true},msg={textContent:""};
const elements={loginScreen,snowApp,loginMessage:msg};
const document={readyState:"loading",body:{classList:{
 add:n=>classes.add(n),remove:n=>classes.delete(n),toggle:(n,v)=>v?classes.add(n):classes.delete(n)
}},getElementById:id=>elements[id]||null,addEventListener:(type,cb)=>handlers[type]=cb};
let authListener=null,sessionSave=null,signOutCalls=0;
const auth={onAuthStateChanged:cb=>{authListener=cb},signOut:()=>{signOutCalls++;return Promise.resolve()}};
const db={
 collection(name){
   if(name==="platformUsers")return {doc:uid=>({get:async()=>({exists:true,data:()=>({statoAccount:"attivo",role:"operatore"})})})};
   if(name==="pianoNeveUsers")return {doc:uid=>({collection:sub=>({doc:id=>({refId:uid+"/"+sub+"/"+id})})})};
   throw Error("Unexpected collection: "+name);
 },
 runTransaction:async callback=>callback({get:async()=>({exists:false}),set:(ref,data)=>sessionSave={ref,data}})
};
const fb={auth:()=>auth,firestore:()=>db};
fb.firestore.FieldValue={serverTimestamp:()=>({__timestampFromParent:true})};
const location={origin:"https://varga.example"};
const snowWindow={location,VARGA_FIREBASE_CONFIG:{projectId:"hera-app-6cd2b"},VargaSnowOutbox:{setUser:()=>{}},
 dispatchEvent:()=>{}};
snowWindow.parent={location,firebase:fb};
const context={window:snowWindow,document,localStorage:{getItem:()=>null,setItem:()=>{}},
 crypto:{randomUUID:()=>"123"},firebase:{apps:[],initializeApp:()=>{throw Error("second Firebase app forbidden in same-origin frame")}},
 console,CustomEvent:class{constructor(type,options){this.type=type;this.detail=options.detail}}};
vm.runInNewContext(authSource,context,{filename:"snow-auth-bridge.js"});
assert.ok(typeof handlers.DOMContentLoaded==="function");
handlers.DOMContentLoaded();
assert.ok(authListener,"Must subscribe to parent Firebase Auth");
(async()=>{
 await authListener({uid:"uid-1",email:"operator@example.com"});
 assert.equal(snowWindow.VargaSnowCloud.user.uid,"uid-1");
 assert.equal(snowWindow.VargaSnowCloud.role,"operatore");
 assert.equal(loginScreen.hidden,true,"No second sign-in screen");
 assert.equal(snowApp.hidden,false,"Navigator should open immediately after shared auth");
 await snowWindow.VargaSnowCloud.saveSession({active:true,routeIndex:0});
 assert.equal(sessionSave.data.uid,"uid-1");
 assert.equal(sessionSave.data.updatedAt.__timestampFromParent,true,"Transactions must use matching Firestore SDK sentinels");
 assert.equal(signOutCalls,0);
 const rules=fs.readFileSync(path.join(__dirname,"..","firestore.rules"),"utf8");
 assert.ok(rules.includes("match /pianoNeveUsers/{uid}"),"Cloud sessions require owner-scoped rules");
 assert.ok(rules.includes("match /pianoNeveRouteEdits/{editId}"),"Admin GPS route edits require scoped rules");
 assert.ok(rules.includes("allow create, update, delete: if isAdmin();"),"Route edits must require admin privileges");
 const prepared=fs.readFileSync(path.join(__dirname,"prepare-capacitor-web.js"),"utf8");
 assert.ok(prepared.includes('["icons", "piano-neve"]'),"Native iOS/Android bundles must carry GPS app");
 const sw=fs.readFileSync(path.join(__dirname,"..","sw.js"),"utf8");
 assert.ok(sw.includes("./piano-neve/major-network-shard-11.js"),"Offline service worker must precache road graph");
 console.log("OK: one Firebase Auth session, local GPS + 20,115 offline edges, native bundle, secure cloud data.");
})().catch(err=>{console.error(err);process.exitCode=1});
