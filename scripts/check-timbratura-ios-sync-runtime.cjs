"use strict";
// Simulated Capacitor plugin, no iPhone, Firebase, network or real notifications.
// Ensure native cancel/schedule sequencing and denied-permission safety.
const assert=require("node:assert/strict");
const fs=require("node:fs");
const vm=require("node:vm");

const calls=[];
let granted=true;
let gate=null;
const nativePlugin={
  async registerActionTypes(){ calls.push("registerActions"); },
  async addListener(){ calls.push("listener"); },
  async checkPermissions(){
    calls.push("permission");
    if(gate){ const g=gate; gate=null; await g; }
    return {display:granted?"granted":"denied"};
  },
  async requestPermissions(){ calls.push("requestPermission"); return {display:granted?"granted":"denied"}; },
  async cancel({notifications}){
    calls.push(["cancel",notifications.map(n=>n.id)]);
  },
  async schedule({notifications}){
    calls.push(["schedule",notifications.map(n=>({id:n.id,hour:n.schedule.on.hour}))]);
  }
};
const window={
  Capacitor:{
    isNativePlatform:()=>true,getPlatform:()=>"ios",
    isPluginAvailable:(name)=>name==="LocalNotifications",
    Plugins:{LocalNotifications:nativePlugin}
  },
  dispatchEvent(){}
};
const ctx={window,console,CustomEvent:class CustomEvent{},Promise,Set,Date,Array,Number,Math};
vm.runInNewContext(fs.readFileSync("timbratura-ios-notifications.js","utf8"),ctx);
const api=window.VargaTimbraturaIOS;
const a={iosScheduled:true,days:[1,2,3,4,5],arrivalFrom:"06:30",returnFrom:"15:15"};
const b={...a,arrivalFrom:"08:05",returnFrom:"17:30"};
const count=tag=>calls.filter(x=>Array.isArray(x)&&x[0]===tag).length;

(async()=>{
  const first=await api.sync(a,{prompt:false});
  assert.equal(first.scheduled,10);
  assert.equal(count("cancel"),1);
  assert.equal(count("schedule"),1);
  assert.equal(calls.some(x=>Array.isArray(x)&&x[0]==="schedule"&&x[1][0].hour===6),true);

  // Permission denial must not cancel 10 previously scheduled reminders.
  granted=false;
  const before=count("cancel");
  const denied=await api.sync(b,{prompt:false});
  assert.equal(denied.permissionDenied,true);
  assert.equal(count("cancel"),before);
  const invalid=await api.sync({...a,days:[8]},{prompt:true});
  assert.equal(invalid.permissionDenied,true);
  assert.equal(count("cancel"),before);

  granted=true;
  const empty=await api.sync({...a,days:[8]},{prompt:false});
  assert.match(empty.error,/Nessun orario/);
  assert.equal(count("cancel"),before);

  // Two rapid Save calls must both be executed in order, with latest last.
  let release;
  gate=new Promise(resolve=>{release=resolve;});
  const firstPending=api.sync(a,{prompt:false});
  const secondPending=api.sync(b,{prompt:false});
  await Promise.resolve();await Promise.resolve();await Promise.resolve();
  release();
  const results=await Promise.all([firstPending,secondPending]);
  assert.deepEqual(results.map(x=>x.scheduled),[10,10]);
  const scheduled=calls.filter(x=>Array.isArray(x)&&x[0]==="schedule");
  assert.equal(scheduled.length,3);
  assert.equal(scheduled[2][1][0].hour,8);
  assert.equal(scheduled[2][1][1].hour,17);

  // Disable invoked during scheduling must not be silently discarded.
  const beforeDisable=count("cancel");
  const dis=await api.sync({...a,iosScheduled:false},{prompt:false});
  assert.equal(dis.scheduled,0);
  assert.equal(count("cancel"),beforeDisable+1);
  assert.equal(calls.filter(x=>Array.isArray(x)&&x[0]==="cancel").at(-1)[1].length,14);
  console.log("PASS: iOS local reminders serialize rapid saves, preserve pending on denial/invalid config, and obey explicit disable.");
})().catch(error=>{console.error(error);process.exitCode=1;});
