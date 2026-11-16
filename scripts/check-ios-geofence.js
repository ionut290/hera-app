"use strict";
const assert = require("node:assert/strict");
const fs = require("node:fs");
const vm = require("node:vm");

const code = fs.readFileSync("timbratura-ios-geofence.js", "utf8");
const ui = fs.readFileSync("timbratura-reminders.js", "utf8");
const html = fs.readFileSync("index.html", "utf8");
const native = fs.readFileSync("ios-native/VargaGeofenceManager.swift", "utf8");
const plugin = fs.readFileSync("ios-native/VargaGeofencePlugin.swift", "utf8");
const delegate = fs.readFileSync("scripts/install-ios-native-geofence.rb", "utf8");
const plist = fs.readFileSync("scripts/configure-ios-plist.py", "utf8");

const calls = [];
const fake = {
  configure: async payload => { calls.push(["configure",payload]); return { supported:true, permission:"always", monitoring:true }; },
  disable: async () => { calls.push(["disable"]); return { monitoring:false }; },
  status: async x => { calls.push(["status",x]); return { monitoring:true, pending: { which:"arrival" } }; }
};
const events = [];
const context = {
  window: {
    Capacitor: { isNativePlatform:()=>true, getPlatform:()=>"ios", registerPlugin:(name)=>name==="VargaGeofence"?fake:null },
    dispatchEvent: e => events.push(e)
  },
  document: { addEventListener:()=>{} },
  CustomEvent: class { constructor(type,options){this.type=type;this.detail=options.detail} },
  console, Date
};
vm.runInNewContext(code,context,{filename:"timbratura-ios-geofence.js"});
(async () => {
  const api=context.window.VargaTimbraturaGeofence;
  assert.equal(api.isAvailable(),true);
  const conf={iosGeofence:true,lat:"44.5100",lon:"11.345",radius:150,days:[1,2,3,4,5],arrivalFrom:"06:30",arrivalTo:"07:15",returnFrom:"15:15",returnTo:"16:15"};
  const result=await api.sync(conf);
  assert.equal(result.monitoring,true);
  assert.equal(calls[0][0],"configure");
  assert.equal(calls[0][1].latitude,44.51);
  assert.equal(calls[0][1].radius,150);
  assert.equal((await api.sync({...conf,radius:60})).error.includes("100"),true);
  assert.equal(calls.length,1,"invalid radius must never configure a geofence");
  await api.sync({...conf,iosGeofence:false});
  assert.equal(calls[1][0],"disable");
  await api.status({consumePending:true});
  assert.equal(events[0].detail.which,"arrival");
  assert.ok(ui.includes("timbratura-ios-geofence"));
  assert.ok(ui.includes("next.iosScheduled=false"),"GPS mode must disable time-only reminders");
  assert.ok(ui.includes("VargaTimbraturaGeofence.sync(config)"));
  assert.ok(html.includes('src="timbratura-ios-geofence.js'));
  assert.ok(html.indexOf('src="timbratura-ios-geofence.js')<html.indexOf('src="timbratura-reminders.js'));
  assert.ok(native.includes("didEnterRegion"));
  assert.ok(native.includes("eligibleReminder(settings, now: Date())"));
  assert.ok(native.includes("locationManager.requestAlwaysAuthorization()"));
  assert.ok(native.includes("UNUserNotificationCenter.current().add(request)"));
  assert.ok(native.includes("readSettings()?.enabled"));
  assert.ok(native.includes("defaults.removeObject(forKey: key)"));
  assert.ok(plugin.includes('jsName = "VargaGeofence"'));
  assert.ok(delegate.includes("VargaGeofenceManager.shared.bootstrap()"));
  assert.ok(delegate.includes("VargaBridgeViewController"));
  assert.ok(plist.includes("NSLocationAlwaysAndWhenInUseUsageDescription"));
  for (const filename of [".github/workflows/build-ios-simulator.yml",".github/workflows/ios-private-testflight.yml",".github/workflows/ios-appstore-candidate.yml"])
    assert.ok(fs.readFileSync(filename,"utf8").includes("ruby scripts/install-ios-native-geofence.rb"),filename);
  console.log("PASS: GPS geofence iOS native, permesso Sempre, filtro orari, apertura Easy-Box, nessun avviso solo orario");
})().catch(error=>{ console.error(error); process.exitCode=1; });
