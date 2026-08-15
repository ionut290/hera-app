#!/usr/bin/env node
"use strict";

const fs = require("node:fs");

const app = fs.readFileSync("app.js", "utf8");
const index = fs.readFileSync("index.html", "utf8");
const sw = fs.readFileSync("sw.js", "utf8");

const modules = [
  {
    file: "app-worklimate.js",
    global: "VargaWorklimateModule",
    functions: [
      "normalizeWorklimateLevel",
      "getMostSevereWorklimateRisk",
      "loadWorklimateRiskCacheBackground",
      "openHomeWorklimateBoard",
      "openSquadraWorklimateSafety"
    ]
  },
  {
    file: "app-atex.js",
    global: "VargaAtexModule",
    functions: [
      "renderAtexProcedurePage",
      "saveAtexProcedureForm",
      "openAtexProcedurePage",
      "shouldShowAtexButtonForImpianto",
      "getAtexIllustrationSvg"
    ]
  },
  {
    file: "app-documents.js",
    global: "VargaDocumentsModule",
    functions: [
      "renderPosDocuments",
      "savePosDocument",
      "renderPrivateDocsList",
      "savePrivateDocument",
      "openNotificationDocumentViewer"
    ]
  },
  {
    file: "app-calendar.js",
    global: "VargaCalendarModule",
    functions: [
      "renderCalendar",
      "saveCalendarEvent",
      "subscribeCalendarEvents",
      "renderNotificationCalendar",
      "applyCalendarAbsenceWarningsToSquadraRows"
    ]
  },
  {
    file: "app-snow.js",
    global: "VargaSnowModule",
    functions: [
      "renderSnowService",
      "renderSnowServiceCommesse",
      "subscribeSnowServiceCollections",
      "saveDrawnSnowRoadPath",
      "openSnowServicePage"
    ]
  }
];

function fail(message) {
  console.error(`❌ ${message}`);
  process.exitCode = 1;
}

function exactScriptIndex(file) {
  const re = new RegExp(`<script[^>]+src=["'][^"']*${file.replace(/[.*+?^${}()|[\]\\]/g, "\\$&")}(?:\\?[^"']*)?["'][^>]*><\\/script>`, "i");
  const match = index.match(re);
  return match ? index.indexOf(match[0]) : -1;
}

const appScriptIndex = exactScriptIndex("app.js");
if (appScriptIndex < 0) fail("app.js non è caricato da index.html");

for (const module of modules) {
  if (!fs.existsSync(module.file)) {
    fail(`${module.file} mancante`);
    continue;
  }
  const source = fs.readFileSync(module.file, "utf8");
  if (!source.includes(module.global)) fail(`${module.file} non espone ${module.global}`);

  for (const name of module.functions) {
    if (!source.includes(`function ${name}`)) fail(`${name} manca da ${module.file}`);
    if (app.includes(`function ${name}`)) fail(`${name} è tornata dentro app.js`);
  }

  const moduleIndex = exactScriptIndex(module.file);
  if (moduleIndex < 0) fail(`${module.file} non è caricato da index.html`);
  else if (appScriptIndex >= 0 && moduleIndex > appScriptIndex) fail(`${module.file} deve essere caricato prima di app.js`);

  if (!sw.includes(`./${module.file}`)) fail(`${module.file} manca dalla shell PWA`);
}

for (const coreName of [
  "normalizeCommessaDocument",
  "normalizeMezzoDocument",
  "normalizePersonaleDocument",
  "normalizeSquadraStoricoDocument",
  "renderPersonalHoursCalendar",
  "renderFerieDisponibilitaCalendar",
  "publishGlobalNotificationEvent"
]) {
  if (!app.includes(`function ${coreName}`)) fail(`${coreName} deve restare nel core app.js`);
}

if (process.exitCode) process.exit(process.exitCode);
console.log("✅ Confini modulari Worklimate + ATEX + Documenti + Calendario + Neve verificati.");
