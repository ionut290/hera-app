const assert = require("node:assert/strict");
const fs = require("node:fs");
const app = fs.readFileSync("app.js", "utf8");
const reader = fs.readFileSync("notification-session-enhancements.js", "utf8");
const userNotifications = fs.readFileSync("functions/user-notifications.js", "utf8");
const doneNotifications = fs.readFileSync("functions/index.js", "utf8");

const mergeConflictMarker = /^(?:<<<<<<<|=======|>>>>>>>)(?: .*)?$/m;
assert.doesNotMatch(app, mergeConflictMarker);
assert.doesNotMatch(reader, mergeConflictMarker);
assert.doesNotMatch(userNotifications, mergeConflictMarker);
assert.doesNotMatch(doneNotifications, mergeConflictMarker);

assert.match(app, /dismissedByUserIds\.\$\{currentUser\.uid\}/);
assert.match(app, /Boolean\(item\?\.dismissedByUserIds\?\.\[currentUser\.uid\]\)/);
assert.match(app, /HeraNotificationReader\?\.archive/);
assert.match(reader, /hera_notification_history_v1/);
assert.match(reader, /showNotificationInbox/);
assert.match(reader, />CHIUDI<\/button>/);
assert.match(reader, /data\.fullMessage/);
assert.match(userNotifications, /ti ha segnato/);
assert.match(userNotifications, /fullMessage/);
assert.match(doneNotifications, /fullMessage/);
console.log("Notification experience checks passed.");
