const assert = require("node:assert/strict");
const fs = require("node:fs");

const app = fs.readFileSync("app.js", "utf8");
const reader = fs.readFileSync("notification-session-enhancements.js", "utf8");
const center = fs.readFileSync("notification-center.js", "utf8");
const serviceWorker = fs.readFileSync("sw.js", "utf8");
const userNotifications = fs.readFileSync("functions/user-notifications.js", "utf8");
const doneNotifications = fs.readFileSync("functions/index.js", "utf8");

const mergeConflictMarker = /^(?:<<<<<<<|=======|>>>>>>>)(?: .*)?$/m;
[app, reader, center, serviceWorker, userNotifications, doneNotifications].forEach((source) => {
  assert.doesNotMatch(source, mergeConflictMarker);
});

assert.match(app, /dismissedByUserIds\.\$\{currentUser\.uid\}/);
assert.match(app, /Boolean\(item\?\.dismissedByUserIds\?\.\[currentUser\.uid\]\)/);
assert.match(app, /HeraNotificationReader\?\.archive/);

assert.match(reader, /hera_notification_history_v1/);
assert.match(reader, /showNotificationInbox/);
assert.doesNotMatch(reader, /dialog\.id\s*=\s*"received-notification-dialog"/);
assert.match(reader, /HeraNotificationCenter\?\.open/);
assert.match(reader, /data\.fullMessage/);

assert.match(center, /L="userAlerts"/);
assert.match(center, /C="notifications"/);
assert.match(center, /data-central-notification-bell/);
assert.match(center, /userAlertAcknowledgements/);
assert.match(center, /OK, HO CAPITO/);

// Verifica che il service worker usi una cache versionata senza bloccare gli aggiornamenti futuri.
assert.match(serviceWorker, /const CACHE_NAME = "hera-app-shell-v\d+";/);
assert.doesNotMatch(serviceWorker, /notification-session-enhancements\.js\?v=/);
assert.match(serviceWorker, /notification-center\.js\?v=20260730b/);

assert.match(userNotifications, /ti ha segnato/);
assert.match(userNotifications, /fullMessage/);
assert.match(doneNotifications, /fullMessage/);

console.log("Notification experience checks passed.");
