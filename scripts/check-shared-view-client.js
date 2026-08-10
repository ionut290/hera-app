"use strict";

const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const vm = require("node:vm");

const root = process.cwd();
const read = (file) => fs.readFileSync(path.join(root, file), "utf8");
const wrapper = read("shared-static-views-client.js");
const explicitGuard = read("hours-source-explicit-guard.js");
const sharedViews = read("shared-static-views.js");
const corePath = path.join(root, "shared-static-views-client-core.js");
const core = fs.existsSync(corePath) ? fs.readFileSync(corePath, "utf8") : "";

async function main() {
// Contratto statico: nessun fallback automatico può riaprire l'intero storico.
assert.match(wrapper, /function stopPrematureHoursSubscriptions\(\)/);
assert.match(wrapper, /unsubscribeHoursStats\(\);\s*unsubscribeHoursStats = null/);
assert.match(wrapper, /unsubscribeHoursApprovals\(\);\s*unsubscribeHoursApprovals = null/);
assert.match(wrapper, /version: "1\.4\.0"/);
assert.match(wrapper, /shared-static-views-client-core\.js\?v=20260806-explicit-hours-v4/);
assert.match(wrapper, /automaticFallbacksBlocked/);
assert.match(wrapper, /fallback completo bloccato/);
assert.match(wrapper, /hoursApprovalsLoaded = true/);
assert.doesNotMatch(wrapper, /enableHoursSource\(\{[\s\S]*forceSharedCalendarFallback/);

assert.match(explicitGuard, /version: "2\.1\.0"/);
assert.match(explicitGuard, /trigger\.isTrusted === true/);
assert.match(explicitGuard, /blockedFallbackStarts/);
assert.match(explicitGuard, /if \(!trustedUserAction\)/);
assert.doesNotMatch(explicitGuard, /if \(!trustedUserAction && !verifiedSharedViewFallback\)/);

assert.match(sharedViews, /const CALENDAR_SCHEMA_VERSION = 2/);
assert.match(sharedViews, /function isCompleteCalendarView\(value\)/);
assert.match(sharedViews, /invalidCalendarCacheDropped/);
assert.match(sharedViews, /cache calendario legacy eliminata/);
assert.match(sharedViews, /snapshot calendario incompleto ignorato/);
assert.match(sharedViews, /typeof allHoursReports !== "undefined" \? allHoursReports/);
assert.match(sharedViews, /typeof allHoursApprovalRequests !== "undefined" \? allHoursApprovalRequests/);
assert.match(sharedViews, /sourceCollection,/);
assert.match(sharedViews, /schemaVersion: CALENDAR_SCHEMA_VERSION/);
assert.match(sharedViews, /completeRecords: true/);
assert.match(sharedViews, /aggiornamento calendario affidato alla Cloud Function/);
assert.doesNotMatch(sharedViews, /window\.allHoursReports/);
assert.doesNotMatch(sharedViews, /window\.allHoursApprovalRequests/);
assert.doesNotMatch(sharedViews, /hera-hours-saved[\s\S]{0,180}publishCalendar/);

if (core) {
  assert.match(core, /subscribeHoursStats = gatedHoursStats/);
  assert.match(core, /bindCapture\("open-hours-btn", enableHoursSource\)/);
  assert.match(core, /stopStaticCalendarForFullHours\(\);[\s\S]*sourceSubscriptions\.hoursStats\(\)/);
  assert.doesNotMatch(core, /monthly-query-fallback/);
}

// Test runtime del guard: fallback programmatico bloccato, click reale consentito.
{
  let fullHoursStarts = 0;
  const window = {
    HeraLightStartup: {
      enableHoursSource() {
        fullHoursStarts += 1;
        return "started";
      }
    },
    setTimeout(callback) { callback(); }
  };
  const context = vm.createContext({ window, console });
  vm.runInContext(explicitGuard, context, { filename: "hours-source-explicit-guard.js" });

  window.HeraLightStartup.enableHoursSource({
    forceSharedCalendarFallback: true,
    reason: "shared-calendar-incomplete"
  });
  assert.equal(fullHoursStarts, 0, "Il fallback automatico non deve aprire le ore complete");

  const result = window.HeraLightStartup.enableHoursSource({ isTrusted: true });
  assert.equal(result, "started");
  assert.equal(fullHoursStarts, 1, "Il click reale su Gestione ore deve restare consentito");
  const state = window.HeraHoursSourceExplicitGuard.getState();
  assert.equal(state.blockedFallbackStarts, 1);
  assert.equal(state.allowedTrustedStarts, 1);
}

// Test runtime del wrapper: una cache legacy può alimentare temporaneamente la UI,
// ma non deve mai chiamare enableHoursSource né riaprire oreReports.
{
  let hoursUnsubscribed = 0;
  let approvalsUnsubscribed = 0;
  let fullHoursStarts = 0;
  let subscribedCallback = null;
  const delivered = [];
  const writtenScripts = [];
  const api = {
    subscribe(_type, _key, callback) {
      subscribedCallback = callback;
      return () => {};
    }
  };
  const window = {
    HeraSharedStaticViews: api,
    HeraLightStartup: {
      enableHoursSource() { fullHoursStarts += 1; }
    }
  };
  const document = {
    readyState: "loading",
    querySelector() { return null; },
    write(value) { writtenScripts.push(value); },
    createElement() { return { dataset: {}, addEventListener() {} }; },
    head: { appendChild() {} }
  };
  const context = vm.createContext({
    window,
    document,
    console,
    queueMicrotask,
    unsubscribeHoursStats: () => { hoursUnsubscribed += 1; },
    unsubscribeHoursApprovals: () => { approvalsUnsubscribed += 1; },
    hoursApprovalsLoaded: false,
    hoursApprovalRequests: [],
    allHoursApprovalRequests: [{ id: "approval-1" }]
  });
  vm.runInContext(wrapper, context, { filename: "shared-static-views-client.js" });

  assert.equal(hoursUnsubscribed, 1);
  assert.equal(approvalsUnsubscribed, 1);
  assert.equal(writtenScripts.length, 2);
  api.subscribe("calendario", "2026-08", (view, metadata) => delivered.push({ view, metadata }));

  subscribedCallback({ payload: { reports: [{ id: "legacy" }] } }, { source: "local" });
  assert.equal(delivered.length, 1);
  assert.equal(delivered[0].metadata.legacyReduced, true);
  assert.equal(fullHoursStarts, 0);
  assert.equal(context.hoursApprovalsLoaded, true);

  subscribedCallback({
    schemaVersion: 2,
    completeRecords: true,
    payload: { schemaVersion: 2, completeRecords: true, reports: [{ id: "full" }] }
  }, { source: "firestore" });
  assert.equal(delivered.length, 2);
  assert.equal(fullHoursStarts, 0);

  subscribedCallback({ payload: null }, { source: "firestore" });
  assert.equal(delivered.length, 2, "Una vista senza reports deve essere ignorata");
  assert.equal(fullHoursStarts, 0);
}

// Test runtime della vista condivisa: cache legacy eliminata e payload v2 completo.
{
  const storage = new Map();
  const listeners = new Map();
  let firestoreWrites = 0;
  let lastWrite = null;
  const localStorage = {
    getItem(key) { return storage.has(key) ? storage.get(key) : null; },
    setItem(key, value) { storage.set(key, String(value)); },
    removeItem(key) { storage.delete(key); }
  };
  const firestore = {
    collection() {
      return {
        doc(id) {
          return {
            onSnapshot() { return () => {}; },
            async set(value) {
              firestoreWrites += 1;
              lastWrite = { id, value };
            }
          };
        }
      };
    }
  };
  const firestoreFactory = () => firestore;
  firestoreFactory.FieldValue = { serverTimestamp: () => ({ server: true }) };
  const window = {
    firebase: {
      firestore: firestoreFactory,
      auth: () => ({ currentUser: { email: "admin@example.com" } })
    },
    canManageData: () => true,
    addEventListener(type, callback) { listeners.set(type, callback); },
    dispatchEvent() {}
  };
  const document = {
    readyState: "complete",
    addEventListener() {}
  };
  const context = vm.createContext({
    window,
    document,
    console,
    localStorage,
    Blob,
    CustomEvent: class CustomEvent { constructor(type, init) { this.type = type; this.detail = init?.detail; } },
    setTimeout,
    clearTimeout
  });
  vm.runInContext(`
    let allHoursReports = [{
      id: "report-1",
      date: "2026-08-06",
      operatore: "Mario",
      customField: "preserved",
      entries: [{ hours: 8 }]
    }];
    let allHoursApprovalRequests = [{
      id: "approval-1",
      date: "2026-08-06",
      status: "pending",
      entries: [{ hours: 2 }]
    }];
  `, context);
  vm.runInContext(sharedViews, context, { filename: "shared-static-views.js" });

  const api = window.HeraSharedStaticViews;
  const payload = api.collectCalendar("2026-08");
  assert.equal(payload.schemaVersion, 2);
  assert.equal(payload.completeRecords, true);
  assert.equal(payload.reports.length, 2);
  assert.equal(payload.reports[0].sourceCollection, "oreApprovalRequests");
  assert.equal(payload.reports[1].sourceCollection, "oreReports");
  assert.equal(payload.reports[1].customField, "preserved");

  const legacyKey = "hera-shared-static-view:calendario__2026-08";
  storage.set(legacyKey, JSON.stringify({ payload: { reports: [] } }));
  assert.equal(api.getCached("calendario", "2026-08"), null);
  assert.equal(storage.has(legacyKey), false);

  const returnPromise = api.publishCalendar("2026-08");
  await returnPromise;
  assert.equal(firestoreWrites, 1);
  assert.equal(lastWrite.value.schemaVersion, 2);
  assert.equal(lastWrite.value.completeRecords, true);
  assert.equal(lastWrite.value.payload.completeRecords, true);

  const writesBeforeEvent = firestoreWrites;
  listeners.get("hera-hours-saved")?.({ detail: { month: "2026-08" } });
  assert.equal(firestoreWrites, writesBeforeEvent, "L'evento locale non deve sovrascrivere la Cloud Function");
}

console.log("Shared views startup, cache migration and explicit full-hours checks passed.");
}

main().catch((error) => {
  console.error(error);
  process.exitCode = 1;
});
