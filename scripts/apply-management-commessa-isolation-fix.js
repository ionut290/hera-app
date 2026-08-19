#!/usr/bin/env node
"use strict";

const fs = require("node:fs");

const path = "operational-import-repair.js";
let source = fs.readFileSync(path, "utf8");

const before = `  function applyManagementStats(canonical = findCanonicalModenaCommessa(), summary = lastSummary) {\n    if (!canonical?.id || !summary) return false;\n    let isCurrent = false;\n    try {\n      if (typeof managementCommessaId !== \"undefined\" && managementCommessaId === canonical.id) isCurrent = true;\n    } catch (_) {}\n    const meta = document.querySelector?.(\"#impianti-management-meta\");\n    if (!isCurrent && meta && normalizeCode(meta.textContent).includes(CANONICAL_COMMESSA_CODE)) isCurrent = true;\n    if (!isCurrent) return false;\n`;

const after = `  function getManagementMetaCode() {\n    const meta = document.querySelector?.(\"#impianti-management-meta\");\n    const value = text(meta?.textContent);\n    const match = value.match(/(?:Cod\\.?|Codice)\\s*[:.]?\\s*([A-Za-z0-9_-]+)/i);\n    return normalizeCode(match?.[1] || \"\");\n  }\n\n  function isManagementForCommessa(commessa = {}) {\n    if (!commessa?.id) return false;\n    try {\n      const activeId = typeof managementCommessaId !== \"undefined\" ? text(managementCommessaId) : \"\";\n      if (activeId) return activeId === text(commessa.id);\n    } catch (_) {}\n    const targetCode = normalizeCode(commessa.codice || commessa.code);\n    return Boolean(targetCode && getManagementMetaCode() === targetCode);\n  }\n\n  function applyManagementStats(canonical = findCanonicalModenaCommessa(), summary = lastSummary) {\n    if (!canonical?.id || !summary || !isManagementForCommessa(canonical)) return false;\n`;

if (!source.includes(after)) {
  if (!source.includes(before)) throw new Error("Blocco applyManagementStats atteso non trovato");
  source = source.replace(before, after);
}

const testingBefore = `      summarizeData,\n      removeSyntheticCommessaFromLocalState,\n      applyMetadataFallback\n`;
const testingAfter = `      summarizeData,\n      removeSyntheticCommessaFromLocalState,\n      applyMetadataFallback,\n      getManagementMetaCode,\n      isManagementForCommessa,\n      applyManagementStats\n`;
if (!source.includes(testingAfter)) {
  if (!source.includes(testingBefore)) throw new Error("Blocco testing atteso non trovato");
  source = source.replace(testingBefore, testingAfter);
}

if (source.includes("normalizeCode(meta.textContent).includes(CANONICAL_COMMESSA_CODE)")) {
  throw new Error("Il confronto per sottostringa del codice commessa è ancora presente");
}

fs.writeFileSync(path, source);
console.log("Correzione isolamento Gestione commesse applicata.");
