"use strict";
const assert = require("node:assert/strict");
const fs = require("node:fs");
const vm = require("node:vm");

const source = fs.readFileSync("timbratura-ios-notifications.js", "utf8");
const html = fs.readFileSync("index.html", "utf8");
const ui = fs.readFileSync("timbratura-reminders.js", "utf8");
const manifest = require("../package.json");
const lock = require("../package-lock.json");

const context = { window: {}, console, CustomEvent: class {}, Set, Date };
vm.runInNewContext(source, context, {filename:"timbratura-ios-notifications.js"});
const api = context.window.VargaTimbraturaIOS;
assert.ok(api && api.makeNotifications);
const config = {
  iosScheduled: true,
  days: [1, 2, 3, 4, 5],
  arrivalFrom: "06:30", returnFrom: "15:15"
};
const notifications = api.makeNotifications(config);
assert.equal(notifications.length, 10, "2 notifications on each of 5 weekdays");
assert.equal(new Set(notifications.map(n=>n.id)).size, 10);
assert.equal(notifications[0].schedule.on.weekday, 2);
assert.equal(notifications[0].schedule.on.hour, 6);
assert.equal(notifications[1].schedule.on.hour, 15);
assert.equal(notifications[8].schedule.on.weekday, 6);
assert.ok(notifications.every(n=>n.sound==="default" && n.schedule.repeats));
assert.equal(api.makeNotifications({...config,iosScheduled:false}).length, 0);
assert.equal(api.makeNotifications({...config,days:[]}).length, 0);
assert.equal(api.makeNotifications({...config,days:[8]}).length, 0);
assert.equal(api.allIDs().length, 14);
assert.equal(api.APP_STORE,"https://apps.apple.com/it/app/easy-box/id1612276209");
assert.ok(!notifications.some(n=>JSON.stringify(n).includes("coordinate")), "scheduled alerts must not claim GPS verification");
assert.ok(ui.includes("timbratura-ios-scheduled"));
assert.ok(ui.includes("VargaTimbraturaIOS.sync(config"));
assert.ok(ui.includes("timbratura-shortcut") && ui.includes("timbratura-appstore"));
assert.ok(html.indexOf('src="timbratura-ios-notifications.js') > 0);
assert.ok(html.indexOf('src="timbratura-ios-notifications.js') < html.indexOf('src="timbratura-reminders.js'));
assert.equal(manifest.dependencies["@capacitor/local-notifications"],"^7.0.7");
assert.equal(lock.packages["node_modules/@capacitor/local-notifications"].version,"7.0.7");

const bad = api.makeNotifications({...config,arrivalFrom:"25:99"});
assert.equal(bad.length,5);
console.log("OK: notifiche iOS offline, orari, giorni, ID univoci, Easy-Box e bundle");
