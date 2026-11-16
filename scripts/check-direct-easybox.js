"use strict";
const assert = require("node:assert/strict");
const fs = require("node:fs");
const vm = require("node:vm");

const source = fs.readFileSync("timbratura-ios-notifications.js", "utf8");
const ui = fs.readFileSync("timbratura-reminders.js", "utf8");
const swift = fs.readFileSync("ios-native/VargaGeofencePlugin.swift", "utf8");

let listener, nativeOpenCalls=0;
const events=[];
const notifications={
  registerActionTypes:async()=>{},
  addListener:async(name,callback)=>{assert.equal(name,"localNotificationActionPerformed");listener=callback;},
  cancel:async()=>{}, checkPermissions:async()=>({display:"granted"}),schedule:async()=>{}
};
const native={
  openEasyBoxShortcut:async()=>{nativeOpenCalls++;return {opened:true};}
};
const win={
  Capacitor:{
    isNativePlatform:()=>true,getPlatform:()=>"ios",
    isPluginAvailable:(name)=>name==="LocalNotifications",
    Plugins:{LocalNotifications:notifications,VargaGeofence:native},
    registerPlugin:(name)=>({LocalNotifications:notifications,VargaGeofence:native})[name]
  },
  dispatchEvent:e=>events.push(e),
  open:()=>{},location:{href:""}
};
const context={window:win,CustomEvent:class{constructor(type,options){this.type=type;this.detail=options.detail;}},console};
vm.runInNewContext(source,context);
(async()=>{
  const api=win.VargaTimbraturaIOS;
  assert.equal(await api.init(),true);
  listener({ notification: { extra:{ feature:"other-feature",which:"arrival" } } });
  assert.equal(events.length,0,"unrelated notification must be ignored");
  listener({ notification: { extra:{ feature:"varga-timbratura",which:"return" } } });
  assert.equal(events.length,1);
  assert.equal(events[0].detail.source,"notification-tap");
  assert.equal(events[0].detail.tapped,true);
  assert.equal(events[0].detail.which,"return");
  assert.equal(nativeOpenCalls,0,"notification listener must NOT launch Easy-Box by itself");
  assert.equal((await api.launchEasyBoxFromNotification()).opened,true);
  assert.equal(nativeOpenCalls,1);
  assert.ok(swift.includes('CAPPluginMethod(name: "openEasyBoxShortcut"'));
  assert.ok(swift.includes('UIApplication.shared.open(url, options: [:])'));
  assert.ok(swift.includes("shortcuts://run-shortcut?name=Apri%20Easy-Box"));
  assert.ok(ui.includes('id="timbratura-easybox-direct"'));
  assert.ok(ui.includes('id="timbratura-easybox-test"'));
  assert.ok(ui.includes("next.directEasyBoxTap="));
  assert.ok(ui.includes('detail.source !== "notification-tap" || detail.tapped !== true'),
      "geofence triggers and ordinary app startup must never open Easy-Box");
  assert.ok(ui.includes('if (!config.directEasyBoxTap) showEasyBoxReminder(detail.which)'));
  assert.ok(ui.includes("window.addEventListener(\"varga:apri-timbratura\""));
  assert.ok(!swift.includes('it.ictgroup.eggspresenze://'),"never invent an Easy-Box deep link");
  console.log("PASS: explicit notification tap only, user opt-in, Apple Shortcut handoff, privacy fallback");
})().catch(e=>{console.error(e);process.exitCode=1;});
