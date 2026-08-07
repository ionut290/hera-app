const assert = require("node:assert/strict");
const fs = require("node:fs");

const disabled = fs.readFileSync("notifications-disabled.js", "utf8");
const firebaseConfig = fs.readFileSync("firebase-config.js", "utf8");
const reader = fs.readFileSync("notification-session-enhancements.js", "utf8");
const center = fs.readFileSync("notification-center.js", "utf8");
const css = fs.readFileSync("notification-center.css", "utf8");

const mergeConflictMarker = /^(?:<<<<<<<|=======|>>>>>>>)(?: .*)?$/m;
[disabled, firebaseConfig, reader, center, css].forEach((source) => {
  assert.doesNotMatch(source, mergeConflictMarker);
});

// Il blocco deve essere installato prima di app.js e impedire traffico Firestore
// sulle raccolte appartenenti al vecchio sistema notifiche.
assert.match(firebaseConfig, /HERA_NOTIFICATIONS_DISABLED_SRC/);
assert.match(firebaseConfig, /notifications-disabled\.js\?v=/);
assert.doesNotMatch(firebaseConfig, /loadOnce\("notification-session-enhancements\.js/);

for (const collection of ["notifications", "userAlerts", "appNotifications", "userAlertAcknowledgements"]) {
  assert.match(disabled, new RegExp(`\\"${collection}\\"`));
}
assert.match(disabled, /QueryProto\.onSnapshot/);
assert.match(disabled, /QueryProto\.get/);
assert.match(disabled, /DocumentProto\.onSnapshot/);
assert.match(disabled, /DocumentProto\.get/);
assert.match(disabled, /CollectionProto\.add/);
assert.match(disabled, /return Promise\.resolve\(null\)/);
assert.match(disabled, /unsubscribeBrowserPush/);
assert.match(disabled, /open-panel-notifiche/);
assert.match(disabled, /panel-notifiche/);
assert.match(disabled, /today-alerts-btn/);

// I vecchi asset restano soltanto come stub per compatibilità con index/build.
assert.doesNotMatch(center, /\.collection\(/);
assert.doesNotMatch(center, /\.onSnapshot\(/);
assert.doesNotMatch(center, /HeraNotificationCenter\s*=\s*\{/);
assert.match(center, /HeraNotificationCenter = undefined/);
assert.match(center, /heraPushFcmToken/);
assert.match(center, /unsubscribePush/);

assert.doesNotMatch(reader, /addEventListener\("message"/);
assert.doesNotMatch(reader, /pushNotificationReceived/);
assert.match(reader, /HeraNotificationReader = undefined/);

for (const id of ["open-panel-notifiche", "panel-notifiche", "user-alert-modal", "notification-doc-viewer-modal", "pwa-notification-status", "enable-notifications-btn", "test-notification-btn", "today-alerts-btn"]) {
  assert.match(css, new RegExp(`#${id}`));
}

console.log("Notification removal checks passed: UI, push and Firestore traffic are disabled.");
