"use strict";

const assert = require("node:assert/strict");
const fs = require("node:fs");

const app = fs.readFileSync("app.js", "utf8");
const interactions = fs.readFileSync("today-summary-interactions.js", "utf8");
const notifications = fs.readFileSync("functions/user-notifications.js", "utf8");
const index = fs.readFileSync("index.html", "utf8");
const serviceWorker = fs.readFileSync("sw.js", "utf8");
const layout = fs.readFileSync("squadre-restyle.css", "utf8");
const androidWorkflow = fs.readFileSync(".github/workflows/build-android-aab.yml", "utf8");
const capacitorBundle = fs.readFileSync("scripts/prepare-capacitor-web.js", "utf8");

assert.match(app, /const subscribedDateKeys = \[\.\.\.new Set\(\[selectedDateKey, todayDateKey\]/);
assert.match(app, /where\("dateKey", "in", subscribedDateKeys\)/);
assert.match(app, /class="squadra-avviso-input"/);
assert.match(app, /avviso: String\(row\.querySelector\("\.squadra-avviso-input"\)/);
assert.match(app, /createSquadraAlertsForChangedRows/);
assert.match(app, /source: "squadra-avviso"/);
assert.match(app, /targetMemberNames/);

assert.match(interactions, /function getCurrentUserSavedHours/);
assert.match(interactions, /timeZone: ROME_TIME_ZONE/);
assert.match(interactions, /elapsed > 8 \* 60 \? elapsed - 60 : elapsed/);
assert.match(interactions, /window\.setInterval\(\(\) => renderTodaySummary\(\), 60 \* 1000\)/);
assert.match(interactions, /stopTodayHoursCounter/);
assert.match(interactions, /replaceSummaryButton\("todayMezziBtn", openAssignedVehicles\)/);
assert.match(interactions, /getUnreadPersonalAlerts/);
assert.match(interactions, /doSquadraMemberAndUserMatch\(row\?\.operatore/);
assert.match(interactions, /openHoursPageForCommessa\(assignments\[0\]\.commessaId/);
assert.match(interactions, /isNotificationForCurrentUser\(alertItem\)/);

assert.match(notifications, /function extractChangedSquadraAlerts/);
assert.match(notifications, /eventType: "squadra-alert"/);
assert.match(notifications, /title: "⚠️ Avviso squadra"/);
assert.match(notifications, /squadAlert\.operators\.some/);

assert.match(index, /today-summary-interactions\.js\?v=20260728c/);
assert.match(index, /id="today-commesse-action">APRI/);
assert.doesNotMatch(index, /Nessuna commessa assegnata|Apri la tua commessa/);
assert.match(index, />Inserisci ore<\/span>/);
assert.match(index, /id="today-mezzi-action">ELENCO MEZZI/);
assert.doesNotMatch(index, /Nessun mezzo assegnato|I tuoi mezzi|Visualizza i mezzi assegnati alla tua squadra/);
assert.match(index, />I tuoi avvisi<\/span>/);
assert.match(index, /squadre-restyle\.css\?v=20260728c/);
assert.match(serviceWorker, /hera-app-shell-v\d+/);
assert.match(serviceWorker, /today-summary-interactions\.js\?v=20260728c/);
assert.match(serviceWorker, /squadre-restyle\.css\?v=20260728c/);
assert.match(layout, /#today-summary-card \.today-summary-grid\s*\{\s*grid-template-columns: repeat\(4, minmax\(0, 1fr\)\);/);
assert.doesNotMatch(layout, /#today-summary-card \.today-summary-grid\s*\{\s*grid-template-columns: repeat\(2,/);
assert.match(layout, /#today-summary-card \.today-summary-item\s*\{[\s\S]*?min-height: 54px;[\s\S]*?padding: 6px 4px;/);
assert.doesNotMatch(layout, /#today-summary-card #today-commesse-count,[\s\S]*?text-overflow: ellipsis/);
assert.match(interactions, /commessaNames\.join\(" • "\)/);
assert.match(interactions, /\[\.\.\.mezzi\.values\(\)\]\.join\(" • "\)/);
assert.match(interactions, /getNotificationPrimaryDateKey\(alertItem\) === dateKey/);
assert.match(androidWorkflow, /npm run android:aab:prepare/);
assert.match(capacitorBundle, /"squadre-restyle\.css"/);
assert.match(capacitorBundle, /"today-summary-interactions\.js"/);

console.log("Today summary and team alert checks passed.");
