#!/usr/bin/env node
"use strict";

const fs = require("node:fs");
const assert = require("node:assert/strict");
const vm = require("node:vm");

const html = fs.readFileSync("index.html", "utf8");
const calendar = fs.readFileSync("app-calendar.js", "utf8");
const feature = fs.readFileSync("administrative-calendar.js", "utf8");
const css = fs.readFileSync("calendar-feature.css", "utf8");
const sharedClient = fs.readFileSync("shared-static-views-client-core.js", "utf8");
const serviceWorker = fs.readFileSync("sw.js", "utf8");
const androidBundle = fs.readFileSync("scripts/prepare-capacitor-web.js", "utf8");

assert.match(html, /id="calendar-choice-administrative-btn"/);
assert.match(html, /id="calendar-administrative-tab"/);
assert.match(html, /administrative-calendar\.js\?v=/);
assert.match(calendar, /mode !== "administrative"/);
assert.match(calendar, /calendarMode === "administrative"/);
assert.match(calendar, /HeraAdministrativeCalendar\?\.render/);

assert.match(feature, /allHoursReports/);
assert.match(feature, /impiantiByCommessaId/);
assert.match(feature, /heraImpiantiPersistentCacheV1/);
assert.match(feature, /groupDayData/);
assert.match(feature, /knownEarnings/);
assert.match(feature, /Ore non ancora inserite/);
assert.match(feature, /importo da valorizzare/);
assert.match(feature, /setCalendarMode\("administrative"\)/);

assert.doesNotMatch(feature, /\bdb\s*\./);
assert.doesNotMatch(feature, /\.onSnapshot\s*\(/);
assert.doesNotMatch(feature, /HeraSharedStaticViews\s*\.\s*subscribe/);
assert.doesNotMatch(feature, /firebase\s*\.\s*firestore/i);
assert.doesNotMatch(feature, /canManageData\s*\(/);

assert.match(css, /\.administrative-day-kpis/);
assert.match(css, /\.administrative-commessa-card/);
assert.match(css, /@media \(max-width: 600px\)/);
assert.match(sharedClient, /calendarMode === "hours" \|\| calendarMode === "administrative"/);
assert.match(serviceWorker, /administrative-calendar\.js\?v=/);
assert.match(androidBundle, /"administrative-calendar\.js"/);

const context = {
  console,
  Intl,
  Date,
  Map,
  Set,
  Object,
  Array,
  Number,
  String,
  JSON,
  Math,
  currentUser: { uid: "user-1" },
  getCommesseCollectionName: () => "commesse",
  commesseById: new Map([["c1", { id: "c1", nome: "HERA DEPURAZIONE", codice: "H-01" }]]),
  impiantiByCommessaId: new Map([["c1", [{
    id: "i1", denominazione: "Depuratore Nord", idSap: "SAP-1", comune: "Bologna",
    done: true, doneAt: "2026-08-26T08:30:00+02:00", doneBy: "Mario Rossi", totale: 150
  }]]]),
  allHoursReports: [{
    date: "2026-08-26",
    entries: [{ commessaId: "c1", commessaName: "HERA DEPURAZIONE", rows: [{ operatore: "Mario Rossi", ore: 8 }] }]
  }],
  normalizeHoursReportDateKey: (value) => String(value || "").slice(0, 10),
  firestoreDateToMillis: (value) => Date.parse(value) || 0,
  formatCalendarDateKey: (value) => new Date(value).toISOString().slice(0, 10),
  formatCalendarLongDate: (value) => value,
  document: { getElementById: () => null },
  localStorage: { length: 0, key: () => null, getItem: () => null }
};
context.window = context;
vm.runInNewContext(feature, context, { filename: "administrative-calendar.js" });
const grouped = context.HeraAdministrativeCalendar.groupDayData("2026-08-26", context.impiantiByCommessaId);
assert.equal(grouped.length, 1);
assert.equal(grouped[0].commessaName, "HERA DEPURAZIONE");
assert.equal(grouped[0].plants.length, 1);
assert.equal(grouped[0].totalHours, 8);
assert.equal(grouped[0].knownEarnings, 150);
assert.equal(grouped[0].operators[0].name, "Mario Rossi");

console.log("✓ Calendario amministrativo visibile a tutti e calcolato soltanto da dati già caricati/cache locale.");
console.log("✓ Nessuna lettura, query o listener Firestore aggiunto dal calendario amministrativo.");
