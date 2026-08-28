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

function sourceBetween(startMarker, endMarker) {
  const start = app.indexOf(startMarker);
  const end = app.indexOf(endMarker, start + startMarker.length);
  assert.notEqual(start, -1, `Missing source marker: ${startMarker}`);
  assert.notEqual(end, -1, `Missing source marker: ${endMarker}`);
  return app.slice(start, end);
}

const subscribeCommesseSource = sourceBetween("function subscribeCommesse()", "function stopCommesseSubscription()");
const subscribePersonaleSource = sourceBetween("function subscribePersonale()", "const DEFAULT_COMMESSE_ABILITAZIONI");
const subscribeMezziSource = sourceBetween("function subscribeMezzi()", "function clearSquadreLoadTimeout()");
const subscribeSquadreSource = sourceBetween("function subscribeSquadre()", "function stopSquadreSubscription()");

assert.match(subscribeCommesseSource, /const applyCommesseSnapshot[\s\S]*?renderNextActionCard\(\);\s*renderTodaySummary\(\);/);
assert.match(subscribeCommesseSource, /query\.onSnapshot[\s\S]*?\(error\) => \{[\s\S]*?commesseLoadState = \{ status: "error"[\s\S]*?renderTodaySummary\(\);/);
assert.match(subscribeCommesseSource, /\.catch\(\(error\) => \{[\s\S]*?commesseLoadState = \{ status: "error"[\s\S]*?renderTodaySummary\(\);/);
assert.match(subscribePersonaleSource, /const applySnapshot[\s\S]*?refreshResolvedUserIdentity\(\);\s*renderTodaySummary\(\);/);
assert.match(subscribePersonaleSource, /query\.onSnapshot\(applySnapshot, \(error\) => \{[\s\S]*?personaleLoadState = \{ status: "error"[\s\S]*?renderTodaySummary\(\);/);
assert.match(subscribePersonaleSource, /\.catch\(\(error\) => \{[\s\S]*?personaleLoadState = \{ status: "error"[\s\S]*?renderTodaySummary\(\);/);
assert.match(subscribeMezziSource, /const applySnapshot[\s\S]*?mezziLoadState = \{ status: "loaded"[^\n]*\n\s*renderTodaySummary\(\);/);
assert.match(subscribeMezziSource, /query\.onSnapshot\(applySnapshot, \(error\) => \{[\s\S]*?mezziLoadState = \{ status: "error"[\s\S]*?renderTodaySummary\(\);/);
assert.match(subscribeMezziSource, /\.catch\(\(error\) => \{[\s\S]*?mezziLoadState = \{ status: "error"[\s\S]*?renderTodaySummary\(\);/);
assert.match(subscribeSquadreSource, /const applySquadreSnapshot[\s\S]*?squadreHistoryByDate\.set[\s\S]*?squadreLoadState = \{ status: "loaded"[^\n]*\n\s*renderTodaySummary\(\);/);
assert.match(subscribeSquadreSource, /squadreQuery\.onSnapshot[\s\S]*?\(error\) => \{[\s\S]*?squadreLoadState = \{ status: "error"[\s\S]*?renderTodaySummary\(\);/);

assert.match(app, /const subscribedDateKeys = \[\.\.\.new Set\(\[selectedDateKey, todayDateKey, tomorrowDateKey\]/);
assert.match(app, /where\("dateKey", "in", subscribedDateKeys\)/);
assert.doesNotMatch(app, /squadra-avviso-input|createSquadraAlertsForChangedRows|source: "squadra-avviso"|row(?:\?\.|\.)avviso\b|alertsCreated/);
assert.doesNotMatch(interactions, /assignedAlerts|row(?:\?\.|\.)avviso\b/);
assert.doesNotMatch(notifications, /extractChangedSquadraAlerts|row(?:\?\.|\.)avviso\b|eventType: "squadra-alert"|title: "⚠️ Avviso squadra"/);
assert.match(app, /legacySquadraAlertSource[\s\S]*?filter\(\(item\) => item\.source !== legacySquadraAlertSource\)/);
assert.match(app, /if \(item\.source === legacySquadraAlertSource\) return false;/);
assert.match(app, /avvisoAutomaticoAssenze/);
assert.match(app, /calendarAbsenceEventIds/);
assert.match(app, /function buildSquadraWarningDetails/);
assert.match(app, /function isPersonAbilitataForCommessa/);
assert.match(app, /hasRequiredPersonaleCourse|hasRequiredCourse/);
assert.doesNotMatch(interactions, /function getAlertGroups\(/);
assert.doesNotMatch(interactions, /function getPlannedHours\(/);
assert.match(notifications, /source: "calendar-absence"/);
assert.match(app, /function findCurrentUserSquadreForDate/);
assert.match(app, /function getSquadrePerCommessaForDate/);
assert.match(app, /candidate\.uids[\s\S]*candidate\.personaleIds[\s\S]*candidate\.emails[\s\S]*candidate\.names/);
assert.match(app, /parseMultiEntryValue\(member \|\| ""\)/);
assert.match(app, /getSquadraNameVariants/);
assert.match(app, /squadrePerCommessa\.forEach/);
assert.match(app, /function updateTodaySummary/);
assert.match(app, /getOreReportsCollectionName\(\)[\s\S]*?allHoursReports = snapshot\.docs\.map[\s\S]*?hoursReportsLoaded = true;\s*renderTodaySummary\(\);/);
assert.match(app, /getOreApprovalRequestsCollectionName\(\)[\s\S]*?allHoursApprovalRequests = snapshot\.docs\.map[\s\S]*?hoursApprovalsLoaded = true;\s*renderTodaySummary\(\);/);

assert.match(interactions, /function getCurrentUserSavedHours/);
assert.match(interactions, /if \(!hoursReportsLoaded \|\| !hoursApprovalsLoaded\) \{\s*return \{ loaded: false, found: false, minutes: 0 \};/);
assert.match(interactions, /return \{ loaded: true, found, minutes: Math\.max\(0, Math\.round\(total \* 60\)\) \};/);
assert.match(interactions, /const hoursText = !savedHours\.loaded[\s\S]*?savedHours\.found/);
assert.match(interactions, /\? "Dati ore in caricamento"/);
assert.match(interactions, /timeZone: ROME_TIME_ZONE/);
assert.doesNotMatch(interactions, /function getLiveWorkedMinutes\(/);
assert.match(interactions, /assignedStart === null \? "--:--" : formatHoursMinutes\(assignedStart\)/);
assert.match(interactions, /window\.setInterval\(\(\) => renderTodaySummary\(\), 60 \* 1000\)/);
assert.match(interactions, /stopTodayHoursCounter/);
assert.match(interactions, /replaceSummaryButton\("todayMezziBtn", openAssignedVehicles\)/);
assert.match(interactions, /getUnreadPersonalAlerts/);
assert.match(interactions, /doSquadraMemberAndUserMatch\(row\?\.operatore/);
assert.match(interactions, /openHoursPageForCommessa\(assignments\[0\]\.commessaId/);
assert.match(interactions, /isNotificationForCurrentUser\(alertItem\)/);
assert.match(notifications, /title: "👷 Aggiunto a una squadra"/);
assert.match(notifications, /avvisoAutomaticoAssenze/);

const getAssetVersion = (source, assetName) => {
  const escapedName = assetName.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
  return source.match(new RegExp(`["'](?:\\./)?${escapedName}\\?v=([^"']+)["']`))?.[1];
};
const assertAssetIsSynchronized = (assetName) => {
  const indexVersion = getAssetVersion(index, assetName);
  const serviceWorkerVersion = getAssetVersion(serviceWorker, assetName);
  assert.ok(indexVersion, `index.html non carica ${assetName} con cache-busting`);
  assert.equal(serviceWorkerVersion, indexVersion, `${assetName} non è sincronizzato in sw.js`);
};

assertAssetIsSynchronized("today-summary-interactions.js");
assert.match(index, /id="today-commesse-action">APRI/);
assert.doesNotMatch(index, /Nessuna commessa assegnata|Apri la tua commessa/);
assert.match(index, />Inserisci ore<\/span>/);
assert.match(index, /id="today-mezzi-btn"[^>]*today-mezzi-only[^>]*>[\s\S]*?id="today-mezzi-count">NESSUN MEZZO/);
assert.doesNotMatch(index, /Nessun mezzo assegnato|I tuoi mezzi|Visualizza i mezzi assegnati alla tua squadra/);
assert.match(index, />I tuoi avvisi<\/span>/);
assertAssetIsSynchronized("squadre-restyle.css");
assert.match(serviceWorker, /[a-z0-9-]+-shell-v\d+/i);
assert.match(layout, /#today-summary-card \.today-summary-grid\s*\{\s*grid-template-columns: repeat\(4, minmax\(0, 1fr\)\);/);
assert.doesNotMatch(layout, /#today-summary-card \.today-summary-grid\s*\{\s*grid-template-columns: repeat\(2,/);
assert.match(layout, /#today-summary-card \.today-summary-item\s*\{[\s\S]*?min-height: 54px;[\s\S]*?padding: 6px 4px;/);
assert.doesNotMatch(layout, /#today-summary-card #today-commesse-count,[\s\S]*?text-overflow: ellipsis/);
assert.match(interactions, /commessaNames\.join\(" • "\)/);
assert.match(interactions, /function renderVehicleBadge\(vehicle\)/);
assert.match(interactions, /vehicles\.map\(renderVehicleBadge\)\.join\(""\)/);
assert.match(interactions, /getNotificationPrimaryDateKey\(alertItem\) === dateKey/);
assert.match(interactions, /const getSummaryDateKey = \(\) => getTodayDateKey\(\);/);
assert.doesNotMatch(interactions, /getActiveSquadreDateKey\(/);
assert.match(interactions, /findCurrentUserSquadreForDate\(getSummaryDateKey\(\)\)/);
assert.match(interactions, /function openAssignedCommessa\(\) \{[\s\S]*?assignments: getAssignments\(\)/);
assert.match(interactions, /function openAssignedHours\(\) \{[\s\S]*?const assignments = getAssignments\(\);[\s\S]*?openHoursPageForCommessa\(assignments\[0\]\.commessaId, getSummaryDateKey\(\)\)/);
assert.match(interactions, /function openAssignedVehicles\(\) \{[\s\S]*?const assignments = getAssignments\(\);/);
assert.match(interactions, /function getUnreadPersonalAlerts\(dateKey = getSummaryDateKey\(\)\)/);
assert.match(interactions, /function renderInteractiveTodaySummary\(\) \{[\s\S]*?const dateKey = getSummaryDateKey\(\);[\s\S]*?findCurrentUserSquadreForDate\(dateKey\)/);
const summaryDateExpression = interactions.match(/const getSummaryDateKey = \(\) => ([^;]+);/)?.[1];
assert.ok(summaryDateExpression, "getSummaryDateKey must be defined");
const evaluateSummaryDate = Function("getTodayDateKey", "getActiveSquadreDateKey", `return (${summaryDateExpression});`);
assert.equal(evaluateSummaryDate(() => "2026-07-28", () => "2026-07-27"), "2026-07-28", "the today card must ignore a different active squadre date");
assert.match(app, /function getSquadraRowMembers/);
assert.match(app, /row\.personale, row\.operatori, row\.caposquadra/);
assert.match(app, /Il riepilogo usa gli stessi dati e la stessa data appena renderizzati qui/);
assert.match(androidWorkflow, /npm run android:aab:prepare/);
assert.match(capacitorBundle, /"squadre-restyle\.css"/);
assert.match(capacitorBundle, /"today-summary-interactions\.js"/);

console.log("Today summary and team alert checks passed.");
