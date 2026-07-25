const assert = require("node:assert/strict");
const fs = require("node:fs");

const backend = fs.readFileSync("functions/index.js", "utf8");
const runtime = fs.readFileSync("native-android-runtime.js", "utf8");
const app = fs.readFileSync("app.js", "utf8");

assert.match(app, /publishGlobalNotificationEvent\("impianto-done"/);
assert.match(backend, /document\("appNotifications\/\{notificationId\}"\)/);
assert.match(backend, /event\.eventType !== "impianto-done"/);
assert.match(backend, /sendEachForMulticast/);
assert.match(backend, /priority: "high"/);
assert.match(backend, /channelId: "hera_operational_updates"/);
assert.match(runtime, /PushNotifications\.createChannel/);
assert.match(runtime, /id: "hera_operational_updates"/);
assert.match(runtime, /auth\.onAuthStateChanged/);

console.log("Android FATTO push check passed.");
