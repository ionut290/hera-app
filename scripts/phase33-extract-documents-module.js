#!/usr/bin/env node
"use strict";

const fs = require("node:fs");
const path = require("node:path");
const acorn = require("acorn");

const ROOT = path.resolve(__dirname, "..");
const APP = path.join(ROOT, "app.js");
const INDEX = path.join(ROOT, "index.html");
const SW = path.join(ROOT, "sw.js");
const OUT = path.join(ROOT, "app-documents.js");

const TARGETS = [
  "buildDocumentViewerUrl",
  "closeNotificationDocumentViewer",
  "closePosDocumentForm",
  "closePrivateDocsPage",
  "createPosDocumentCard",
  "deletePosDocument",
  "deletePrivateDocument",
  "getFilteredPosDocuments",
  "getOrCreatePrivateDocsFolder",
  "getPrivateDocsDriveToken",
  "openDocumentLink",
  "openNotificationDocumentViewer",
  "openPosDocumentForm",
  "openPrivateDocsPage",
  "openPrivateDocsUploadPage",
  "renderPosDocuments",
  "renderPrivateDocsList",
  "savePosDocument",
  "savePrivateDocument",
  "stopPosDocumentsSubscription",
  "stopPrivateDocsSubscription",
  "subscribePosDocuments",
  "subscribePrivateDocs",
  "uploadPrivateDocumentToDrive"
];

let source = fs.readFileSync(APP, "utf8");
const ast = acorn.parse(source, { ecmaVersion: "latest", sourceType: "script", allowHashBang: true });
const found = ast.body.filter((node) => node.type === "FunctionDeclaration" && TARGETS.includes(node.id?.name));
const names = found.map((node) => node.id.name).sort();
const missing = TARGETS.filter((name) => !names.includes(name));
if (missing.length) throw new Error(`Funzioni documenti mancanti: ${missing.join(", ")}`);
if (found.length !== TARGETS.length) throw new Error(`Attese ${TARGETS.length} funzioni, trovate ${found.length}`);

const moduleParts = [
  '"use strict";',
  '(function installVargaDocumentsModule(global) {',
  '  if (global.VargaDocumentsModule) return;',
  '  const api = {};'
];
for (const fn of found.sort((a, b) => a.start - b.start)) {
  const text = source.slice(fn.start, fn.end);
  moduleParts.push(text.split("\n").map((line) => "  " + line).join("\n"));
  moduleParts.push(`  api.${fn.id.name} = ${fn.id.name};`);
}
moduleParts.push('  Object.assign(global, api);');
moduleParts.push('  global.VargaDocumentsModule = Object.freeze({ ...api });');
moduleParts.push('})(window);');
moduleParts.push('');
const moduleSource = moduleParts.join("\n");
acorn.parse(moduleSource, { ecmaVersion: "latest", sourceType: "script", allowHashBang: true });
fs.writeFileSync(OUT, moduleSource, "utf8");

const ranges = found.map((fn) => ({ start: fn.start, end: fn.end })).sort((a, b) => b.start - a.start);
for (const range of ranges) {
  let end = range.end;
  while (end < source.length && /[ \t]/.test(source[end])) end += 1;
  if (source[end] === "\r") end += 1;
  if (source[end] === "\n") end += 1;
  source = source.slice(0, range.start) + source.slice(end);
}
acorn.parse(source, { ecmaVersion: "latest", sourceType: "script", allowHashBang: true });
fs.writeFileSync(APP, source, "utf8");

let index = fs.readFileSync(INDEX, "utf8");
if (!index.includes("app-documents.js")) {
  const appTag = index.match(/<script[^>]+src=["'][^"']*app\.js[^"']*["'][^>]*><\/script>/i);
  if (!appTag) throw new Error("Tag app.js non trovato in index.html");
  index = index.replace(appTag[0], `<script src="app-documents.js?v=20260815-mod1"></script>\n${appTag[0]}`);
  fs.writeFileSync(INDEX, index, "utf8");
}

let sw = fs.readFileSync(SW, "utf8");
if (!sw.includes("app-documents.js")) {
  const marker = /(["']\.\/app\.js[^"']*["'],?)/;
  if (!marker.test(sw)) throw new Error("app.js non trovato nella shell PWA");
  sw = sw.replace(marker, `"./app-documents.js?v=20260815-mod1",\n  $1`);
  fs.writeFileSync(SW, sw, "utf8");
}

console.log(`Estratto modulo Documenti: ${TARGETS.length} funzioni, ${moduleSource.length} byte.`);
