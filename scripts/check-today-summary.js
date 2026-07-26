"use strict";

const assert = require("node:assert/strict");
const fs = require("node:fs");

const app = fs.readFileSync("app.js", "utf8");
const interactions = fs.readFileSync("today-summary-interactions.js", "utf8");
const notifications = fs.readFileSync("functions/user-notifications.js", "utf8");
const index = fs.readFileSync("index.html", "utf8");
const serviceWorker = fs.readFileSync("sw.js", "utf8");

assert.match(app, /const subscribedDateKeys = \[\.\.\.new Set\(\[selectedDateKey, todayDateKey\]/);
assert.match(app, /where\("dateKey", "in", subscribedDateKeys\)/);
assert.match(app, /class="squadra-avviso-input"/);
assert.match(app, /avviso: String\(row\.querySelector\("\.squadra-avviso-input"\)/);
assert.match(app, /createSquadraAlertsForChangedRows/);
assert.match(app, /source: "squadra-avviso"/);
assert.match(app, /targetMemberNames/);

assert.match(interactions, /function getPlannedHours/);
assert.match(interactions, /getSquadraWorkedHours\(row\)/);
assert.doesNotMatch(interactions, /getRecordedHours/);
assert.match(interactions, /replaceSummaryButton\("todayMezziBtn", openAssignedVehicles\)/);
assert.match(interactions, /row\?\.avviso/);

assert.match(notifications, /function extractChangedSquadraAlerts/);
assert.match(notifications, /eventType: "squadra-alert"/);
assert.match(notifications, /title: "⚠️ Avviso squadra"/);
assert.match(notifications, /squadAlert\.operators\.some/);

assert.match(index, /today-summary-interactions\.js\?v=20260726b/);
assert.match(serviceWorker, /hera-app-shell-v31/);
assert.match(serviceWorker, /today-summary-interactions\.js\?v=20260726b/);

console.log("Today summary and team alert checks passed.");
