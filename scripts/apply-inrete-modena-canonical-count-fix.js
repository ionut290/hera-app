#!/usr/bin/env node
"use strict";

const fs = require("node:fs");

function replaceOnce(source, before, after, label) {
  if (source.includes(after)) return source;
  const first = source.indexOf(before);
  if (first < 0) throw new Error(`Blocco non trovato: ${label}`);
  if (source.indexOf(before, first + before.length) >= 0) throw new Error(`Blocco non univoco: ${label}`);
  return source.slice(0, first) + after + source.slice(first + before.length);
}

function replaceRegexOnce(source, pattern, after, label) {
  const flags = pattern.flags.includes("g") ? pattern.flags : `${pattern.flags}g`;
  const globalPattern = new RegExp(pattern.source, flags);
  const matches = [...source.matchAll(globalPattern)];
  if (matches.length !== 1) throw new Error(`Blocco ${label}: attese 1 corrispondenza, trovate ${matches.length}`);
  return source.replace(pattern, after);
}

function write(path, content) {
  fs.writeFileSync(path, content, "utf8");
}

// 1) Il runtime storico non deve più creare una seconda commessa INRETE Modena.
{
  const path = "operational-import-repair.js";
  let source = fs.readFileSync(path, "utf8");

  source = replaceOnce(
    source,
    '  const COMMESSA_ID = "inrete_modena_agosto_2026";\n',
    '  const COMMESSA_ID = "inrete_modena_agosto_2026";\n  const CANONICAL_COMMESSA_CODE = "28015";\n  const CANONICAL_COMMESSA_NAME = "INRETE MODENA";\n  let duplicateCleanupRunning = false;\n  let duplicateArchivedInto = "";\n',
    "costanti commessa canonica"
  );

  const canWriteBlock = `  function canWriteParentCommessa() {\n    try {\n      if (typeof canManageData === "function") return canManageData() === true;\n      if (typeof auth !== "undefined") {\n        return text(auth?.currentUser?.email).toLowerCase() === "ionut29019@gmail.com";\n      }\n    } catch (_) {}\n    return false;\n  }\n`;

  const canonicalHelpers = `${canWriteBlock}\n  function normalizeCommessaCode(value) {\n    return text(value).toLocaleUpperCase("it-IT").replace(/[^A-Z0-9]+/g, "");\n  }\n\n  function findCanonicalModenaCommessa() {\n    try {\n      if (typeof commesseById === "undefined" || !commesseById?.entries) return null;\n      let nameMatch = null;\n      for (const [id, raw] of commesseById.entries()) {\n        const commessa = { id, ...(raw || {}) };\n        if (!id || id === COMMESSA_ID || commessa.hiddenFromHome === true || commessa.mergedIntoCommessaId) continue;\n        const code = normalizeCommessaCode(commessa.codice || commessa.code);\n        const name = norm(commessa.nome || commessa.name);\n        if (code === normalizeCommessaCode(CANONICAL_COMMESSA_CODE)) return commessa;\n        if (name === norm(CANONICAL_COMMESSA_NAME)) nameMatch = nameMatch || commessa;\n      }\n      return nameMatch;\n    } catch (_) {\n      return null;\n    }\n  }\n\n  function removeSyntheticCommessaFromLocalState() {\n    let changed = false;\n    try {\n      if (typeof commesseById !== "undefined" && commesseById?.delete) changed = commesseById.delete(COMMESSA_ID) || changed;\n      if (typeof impiantiByCommessaId !== "undefined" && impiantiByCommessaId?.delete) changed = impiantiByCommessaId.delete(COMMESSA_ID) || changed;\n      if (typeof commessaStatsById !== "undefined" && commessaStatsById?.delete) changed = commessaStatsById.delete(COMMESSA_ID) || changed;\n      if (typeof commessaWorkSummariesById !== "undefined" && commessaWorkSummariesById?.delete) changed = commessaWorkSummariesById.delete(COMMESSA_ID) || changed;\n      if (typeof commessaHoursById !== "undefined" && commessaHoursById?.delete) changed = commessaHoursById.delete(COMMESSA_ID) || changed;\n    } catch (_) {}\n    if (changed) {\n      try { if (typeof renderCommesseHomeList === "function") renderCommesseHomeList(); } catch (_) {}\n      try { if (typeof renderCommesseManagementList === "function") renderCommesseManagementList(); } catch (_) {}\n      try { if (typeof refreshCommesseDependentUI === "function") refreshCommesseDependentUI(false); } catch (_) {}\n    }\n    return changed;\n  }\n\n  async function archiveSyntheticDuplicate(canonical) {\n    if (!canonical?.id || canonical.id === COMMESSA_ID) return false;\n    removeSyntheticCommessaFromLocalState();\n    if (duplicateArchivedInto === canonical.id) return true;\n    if (duplicateCleanupRunning || !canWriteParentCommessa()) return false;\n    if (typeof db === "undefined" || !db || typeof auth === "undefined" || !auth.currentUser) return false;\n\n    duplicateCleanupRunning = true;\n    try {\n      const ref = db.collection(collectionName()).doc(COMMESSA_ID);\n      const snapshot = await ref.get();\n      if (!snapshot.exists) {\n        duplicateArchivedInto = canonical.id;\n        return true;\n      }\n      const current = snapshot.data() || {};\n      const alreadyArchived = current.attiva === false\n        && current.hiddenFromHome === true\n        && text(current.mergedIntoCommessaId) === text(canonical.id);\n      if (!alreadyArchived) {\n        const now = firebase.firestore.FieldValue.serverTimestamp();\n        await ref.set({\n          attiva: false,\n          stato: "Archiviata",\n          hiddenFromHome: true,\n          mergedIntoCommessaId: canonical.id,\n          mergedIntoCommessaCode: canonical.codice || canonical.code || CANONICAL_COMMESSA_CODE,\n          mergeReason: "Duplicato tecnico sostituito dalla commessa INRETE MODENA canonica",\n          mergedAt: now,\n          updatedAt: now,\n          updatedBy: auth.currentUser.uid\n        }, { merge: true });\n      }\n      duplicateArchivedInto = canonical.id;\n      removeSyntheticCommessaFromLocalState();\n      console.info("[INRETE Modena] duplicato tecnico archiviato", { duplicateId: COMMESSA_ID, canonicalId: canonical.id });\n      return true;\n    } catch (error) {\n      console.warn("[INRETE Modena] archiviazione duplicato non riuscita", error);\n      return false;\n    } finally {\n      duplicateCleanupRunning = false;\n    }\n  }\n`;

  source = replaceOnce(source, canWriteBlock, canonicalHelpers, "helper commessa canonica");

  source = replaceRegexOnce(
    source,
    /  function applyMergedVisiblePlants\(\) \{[\s\S]*?\n  function mergeCommessaLocally\(existing\) \{/,
    `  function applyMergedVisiblePlants() {\n    const canonical = findCanonicalModenaCommessa();\n    if (!canonical?.id) return [];\n    const commessaId = canonical.id;\n    let cached = [];\n    try {\n      if (typeof impiantiByCommessaId !== "undefined" && impiantiByCommessaId?.get) {\n        cached = impiantiByCommessaId.get(commessaId) || [];\n      }\n    } catch (_) {}\n\n    let current = [];\n    try {\n      if (\n        typeof selectedCommessaId !== "undefined"\n        && selectedCommessaId === commessaId\n        && typeof currentImpianti !== "undefined"\n        && Array.isArray(currentImpianti)\n      ) current = currentImpianti;\n    } catch (_) {}\n\n    const merged = mergePlants(cached, current);\n    if (!merged.length) return merged;\n    try {\n      if (\n        typeof impiantiByCommessaId !== "undefined"\n        && impiantiByCommessaId?.set\n        && signature(cached) !== signature(merged)\n      ) impiantiByCommessaId.set(commessaId, merged.map((item) => ({ ...item, commessaId, localFallback: false })));\n      if (typeof commessaStatsById !== "undefined" && commessaStatsById?.set && typeof calculateImpiantiStats === "function") {\n        commessaStatsById.set(commessaId, calculateImpiantiStats(merged));\n      }\n    } catch (_) {}\n\n    try {\n      if (\n        typeof selectedCommessaId !== "undefined"\n        && selectedCommessaId === commessaId\n        && typeof currentImpianti !== "undefined"\n        && signature(current) !== signature(merged)\n      ) {\n        currentImpianti = merged.map((item) => ({ ...item, commessaId, localFallback: false }));\n        if (typeof renderImpianti === "function") renderImpianti();\n        if (typeof renderMap === "function") renderMap();\n        if (typeof renderHeaderActivitySummary === "function") renderHeaderActivitySummary();\n        if (typeof updateCommessaDashboard === "function") updateCommessaDashboard();\n      }\n    } catch (error) {\n      console.warn("[INRETE Modena merge impianti UI]", error);\n    }\n    return merged;\n  }\n\n  async function refreshWorkKinds() {\n    const canonical = findCanonicalModenaCommessa();\n    if (\n      !canonical?.id\n      || mixedRefreshRunning\n      || typeof db === "undefined"\n      || typeof auth === "undefined"\n      || !auth.currentUser\n      || (typeof document !== "undefined" && document.hidden)\n    ) return false;\n\n    mixedRefreshRunning = true;\n    try {\n      const ref = db.collection(collectionName()).doc(canonical.id);\n      const snapshot = await ref.collection("lavorazioni").get();\n      workItemsCache = snapshot.docs.map((doc) => ({ id: doc.id, ...doc.data() }));\n      applyMergedVisiblePlants();\n      return true;\n    } catch (error) {\n      console.warn("[INRETE Modena] lettura lavorazioni ordinario/straordinario non riuscita", error);\n      return false;\n    } finally {\n      mixedRefreshRunning = false;\n    }\n  }\n\n  function mergeCommessaLocally(existing) {`,
    "merge canonico e lettura lavorazioni"
  );

  source = replaceRegexOnce(
    source,
    /  function ensureVisibleLocally\(\) \{[\s\S]*?\n  function buildCanonicalIndex\(existingCollections = \[\]\) \{/,
    `  function ensureVisibleLocally() {\n    try {\n      const canonical = findCanonicalModenaCommessa();\n      if (!canonical?.id) return false;\n      removeSyntheticCommessaFromLocalState();\n      archiveSyntheticDuplicate(canonical).catch((error) => {\n        console.warn("[INRETE Modena] pulizia duplicato differita", error);\n      });\n      return true;\n    } catch (error) {\n      console.warn("[INRETE Modena canonical cleanup]", error);\n      return false;\n    }\n  }\n\n  function buildCanonicalIndex(existingCollections = []) {`,
    "stop creazione fallback sintetico"
  );

  source = replaceRegexOnce(
    source,
    /  async function createInFirestore\(force = false\) \{[\s\S]*?\n  async function tryRun\(\) \{/,
    `  async function createInFirestore(force = false) {\n    void force;\n    if (running || typeof db === "undefined" || typeof auth === "undefined" || !auth.currentUser) return false;\n    const canonical = findCanonicalModenaCommessa();\n    if (!canonical?.id) return false;\n    running = true;\n    try {\n      removeSyntheticCommessaFromLocalState();\n      await archiveSyntheticDuplicate(canonical);\n      return true;\n    } finally {\n      running = false;\n    }\n  }\n\n  async function tryRun() {`,
    "disabilita scrittura commessa sintetica"
  );

  source = replaceOnce(
    source,
    `      buildExistingPlantPatch,\n      setWorkItems: setWorkItemsForTesting`,
    `      buildExistingPlantPatch,\n      findCanonicalModenaCommessa,\n      removeSyntheticCommessaFromLocalState,\n      setWorkItems: setWorkItemsForTesting`,
    "export test canonico"
  );

  source = replaceOnce(
    source,
    `      if (\n        typeof selectedCommessaId !== "undefined"\n        && selectedCommessaId === COMMESSA_ID\n        && (typeof document === "undefined" || !document.hidden)\n      ) refreshWorkKinds();`,
    `      const canonical = findCanonicalModenaCommessa();\n      if (\n        canonical?.id\n        && typeof selectedCommessaId !== "undefined"\n        && selectedCommessaId === canonical.id\n        && (typeof document === "undefined" || !document.hidden)\n      ) refreshWorkKinds();`,
    "refresh periodico canonico"
  );

  source = replaceOnce(
    source,
    `        if (typeof selectedCommessaId !== "undefined" && selectedCommessaId === COMMESSA_ID) refreshWorkKinds();`,
    `        const canonical = findCanonicalModenaCommessa();\n        if (canonical?.id && typeof selectedCommessaId !== "undefined" && selectedCommessaId === canonical.id) refreshWorkKinds();`,
    "refresh visibilità canonico"
  );

  write(path, source);
}

// 2) Le statistiche della home usano anche i contatori affidabili salvati sulla commessa.
{
  const path = "commessa-stats-cache-optimizer.js";
  let source = fs.readFileSync(path, "utf8");

  source = replaceOnce(
    source,
    `  const originalSubscribeStatsForCommesse = subscribeStatsForCommesse;\n  const state = {`,
    `  const originalSubscribeStatsForCommesse = subscribeStatsForCommesse;\n  const originalGetCommessaStats = typeof getCommessaStats === "function" ? getCommessaStats : null;\n\n  function metadataNumber(...values) {\n    const numbers = values.map((value) => Number(value)).filter((value) => Number.isFinite(value) && value >= 0);\n    return numbers.length ? Math.max(...numbers) : 0;\n  }\n\n  function statsFromCommessaMetadata(commessaId, baseStats = {}) {\n    const commessa = commesseById.get(commessaId) || {};\n    const total = metadataNumber(commessa.impiantiCount, commessa.totalPlants, commessa.uniquePlantsCount);\n    const done = Math.min(total, metadataNumber(commessa.impiantiFattiCount, commessa.donePlantsCount));\n    return total > 0 ? { ...baseStats, total, done } : baseStats;\n  }\n\n  if (originalGetCommessaStats) {\n    getCommessaStats = function getCommessaStatsWithMetadataFallback(commessaId) {\n      const current = originalGetCommessaStats(commessaId);\n      if (Number(current?.total || 0) > 0) return current;\n      return statsFromCommessaMetadata(commessaId, current);\n    };\n  }\n\n  const state = {`,
    "fallback statistiche da metadati"
  );

  source = replaceOnce(
    source,
    `    changedDocumentsRead: 0,\n    errors: []`,
    `    changedDocumentsRead: 0,\n    metadataFallbackEnabled: Boolean(originalGetCommessaStats),\n    errors: []`,
    "stato fallback metadati"
  );

  source = replaceOnce(
    source,
    `  function targetCommessaIds() {\n    const selected = String(selectedCommessaId || "").trim();\n    if (!selected || typeof getSubcommesse !== "function") return [];\n    const children = getSubcommesse(selected) || [];\n    return children.map((item) => String(item?.id || "").trim()).filter(Boolean);\n  }`,
    `  function targetCommessaIds() {\n    const selected = String(selectedCommessaId || "").trim();\n    if (!selected) return [];\n    const ids = new Set([selected]);\n    if (typeof getSubcommesse === "function") {\n      (getSubcommesse(selected) || []).forEach((item) => {\n        const id = String(item?.id || "").trim();\n        if (id) ids.add(id);\n      });\n    }\n    return [...ids];\n  }`,
    "include commessa selezionata nelle statistiche"
  );

  source = replaceOnce(
    source,
    `    originalSubscribeStatsForCommesse,\n    getState:`,
    `    originalSubscribeStatsForCommesse,\n    originalGetCommessaStats,\n    statsFromCommessaMetadata,\n    getState:`,
    "export diagnostica metadati"
  );

  source = replaceOnce(
    source,
    `  window.HeraCommessaStatsCacheOptimizer = {\n    installed: true,`,
    `  setTimeout(() => refreshStatsUI(), 0);\n\n  window.HeraCommessaStatsCacheOptimizer = {\n    installed: true,`,
    "refresh iniziale statistiche"
  );

  write(path, source);
}

// 3) Conteggio riutilizzabile: impianti fisici unici separati dalle lavorazioni.
{
  const path = "inrete-work-items-v2.js";
  let source = fs.readFileSync(path, "utf8");

  const marker = `  const derivePlantStatus = items => {\n    if (items.some(item => item.economicStatus === "DA_VERIFICARE")) return "DA VERIFICARE";\n    const states=items.map(item=>normalizeStatus(item.stato));\n    if (states.includes("IN LAVORAZIONE")) return "IN LAVORAZIONE";\n    if (states.every(s=>s==="FATTO")) return "FATTO";\n    if (states.every(s=>s==="DA FARE")) return "DA FARE";\n    if (states.includes("FATTO")) return "PARZIALMENTE FATTO";\n    return states[0] || "DA FARE";\n  };\n`;

  const helpers = `${marker}\n  const physicalPlantIdentity = row => {\n    const sap = normalizePriceCode(row?.idSap ?? row?.idSAP ?? row?.["ID SAP"]).replace(/[^A-Z0-9]+/g, "");\n    if (sap) return `SAP:${sap}`;\n    const direct = String(row?.impiantoId ?? row?.physicalPlantId ?? row?.migrationSourceId ?? "").trim().split("::")[0];\n    if (direct) return `ID:${direct}`;\n    const normalize = value => String(value ?? "").trim().toLocaleLowerCase("it-IT").normalize("NFD").replace(/[\\u0300-\\u036f]/g, "").replace(/[^a-z0-9]+/g, "");\n    const anagrafica = normalize(`${row?.denominazione ?? row?.nome ?? ""}|${row?.comune ?? ""}|${row?.indirizzo ?? row?.via ?? ""}`);\n    if (anagrafica) return `ANAG:${anagrafica}`;\n    return `ROW:${String(row?.id ?? row?.numeroProgressivoRiga ?? "").trim()}`;\n  };\n\n  const summarizeWorkItems = (items = [], plants = []) => {\n    const groups = new Map();\n    const ensure = row => {\n      const key = physicalPlantIdentity(row);\n      if (!groups.has(key)) groups.set(key, []);\n      return groups.get(key);\n    };\n    (plants || []).filter(Boolean).forEach(ensure);\n    let doneWork = 0;\n    let todoWork = 0;\n    (items || []).filter(Boolean).forEach(item => {\n      ensure(item).push(item);\n      const status = normalizeStatus(item.stato);\n      if (status === "FATTO") doneWork += 1;\n      else if (status === "DA FARE") todoWork += 1;\n    });\n    const donePlants = [...groups.values()].filter(rows => rows.length > 0 && rows.every(item => normalizeStatus(item.stato) === "FATTO")).length;\n    return {\n      uniquePlants: groups.size,\n      workRows: (items || []).length,\n      doneWork,\n      todoWork,\n      donePlants,\n      todoPlants: Math.max(0, groups.size - donePlants)\n    };\n  };\n`;

  source = replaceOnce(source, marker, helpers, "riepilogo impianti fisici");
  source = replaceOnce(
    source,
    `calculateCompletedSubtotal,isInreteCommessa,derivePlantStatus,migrateInreteCommesseToWorkItemsV2`,
    `calculateCompletedSubtotal,isInreteCommessa,derivePlantStatus,physicalPlantIdentity,summarizeWorkItems,migrateInreteCommesseToWorkItemsV2`,
    "export riepilogo impianti"
  );
  write(path, source);
}

// 4) La gestione mostra chiaramente impianti unici e lavorazioni e riallinea i contatori parent.
{
  const path = "accounting-v2.js";
  let source = fs.readFileSync(path, "utf8");

  source = replaceOnce(
    source,
    `  const actor=()=>({updatedAt:server(),updatedBy:currentUser?.uid||"",updatedByName:getOperatorDisplayName()});\n`,
    `  const actor=()=>({updatedAt:server(),updatedBy:currentUser?.uid||"",updatedByName:getOperatorDisplayName()});\n  const managementSummary=()=>core.summarizeWorkItems(state.work.map(joined),state.plants);\n\n  function applyParentSummaryLocally(summary){\n    if(!state.commessa)return;\n    const patch={impiantiCount:summary.uniquePlants,totalPlants:summary.uniquePlants,workItemsCount:summary.workRows,impiantiFattiCount:summary.donePlants,impiantiDaFareCount:summary.todoPlants,workItemsFattiCount:summary.doneWork,workItemsDaFareCount:summary.todoWork};\n    Object.assign(state.commessa,patch);\n    try{\n      if(typeof commesseById!=="undefined"&&commesseById?.set){\n        const cached=commesseById.get(state.commessa.id)||state.commessa;\n        commesseById.set(state.commessa.id,{...cached,...patch});\n      }\n      if(typeof commessaStatsById!=="undefined"&&commessaStatsById?.set){\n        const current=commessaStatsById.get(state.commessa.id)||{};\n        commessaStatsById.set(state.commessa.id,{...current,total:summary.uniquePlants,done:summary.donePlants});\n      }\n      if(typeof renderCommesseHomeList==="function")renderCommesseHomeList();\n      if(typeof renderCommesseManagementList==="function")renderCommesseManagementList();\n    }catch(_){}\n    return patch;\n  }\n\n  async function syncParentSummary(){\n    if(!state.commessa)return null;\n    const summary=managementSummary();\n    const before={...state.commessa};\n    const patch=applyParentSummaryLocally(summary);\n    const fields=Object.keys(patch);\n    const changed=fields.some(field=>Number(before[field]??-1)!==Number(patch[field]));\n    if(changed&&typeof canManageData==="function"&&canManageData()){\n      await commRef().set({...patch,summaryModelVersion:2,summarySyncedAt:server(),...actor()},{merge:true});\n    }\n    return summary;\n  }\n`,
    "sincronizzazione contatori parent"
  );

  source = replaceOnce(
    source,
    `    render();\n  }\n  const clean=`,
    `    render();\n    syncParentSummary().catch(error=>console.warn("Riepilogo commessa non sincronizzato:",error));\n  }\n  const clean=`,
    "sync riepilogo dopo caricamento"
  );

  source = replaceOnce(
    source,
    `  function render(){\n    const done=state.work.filter(w=>w.stato==="FATTO").length,todo=state.work.filter(w=>w.stato==="DA FARE").length;`,
    `  function render(){\n    const summary=managementSummary(),done=summary.doneWork,todo=summary.todoWork;`,
    "render usa riepilogo univoco"
  );

  source = replaceOnce(
    source,
    '    document.querySelector("#impianti-management-stats").innerHTML=`<span><b>${state.work.length}</b> righe</span><span class="is-done"><b>${done}</b> fatte</span><span><b>${todo}</b> da fare</span>${coordinateProblems?`<span class="coordinate-warning-stat"><b>${coordinateProblems}</b> problemi coordinate</span>`:""}<span><b>${money.format(subtotal())}</b> Subtotale lavorazioni fatte</span>`;',
    '    document.querySelector("#impianti-management-stats").innerHTML=`<span><b>${summary.uniquePlants}</b> impianti</span><span><b>${summary.workRows}</b> lavorazioni</span><span class="is-done"><b>${done}</b> lavorazioni fatte</span><span><b>${todo}</b> lavorazioni da fare</span>${coordinateProblems?`<span class="coordinate-warning-stat"><b>${coordinateProblems}</b> problemi coordinate</span>`:""}<span><b>${money.format(subtotal())}</b> Subtotale lavorazioni fatte</span>`;',
    "etichette riepilogo contabilità"
  );

  source = replaceOnce(
    source,
    `  const api={repairImportedMatrixPlants,synchronizeOperationalModel,migrateInreteCommesseToWorkItemsV2:migrateInrete,calculations:`,
    `  const api={repairImportedMatrixPlants,synchronizeOperationalModel,migrateInreteCommesseToWorkItemsV2:migrateInrete,testing:{summarizeWorkItems:core.summarizeWorkItems,physicalPlantIdentity:core.physicalPlantIdentity},calculations:`,
    "export test riepilogo"
  );

  write(path, source);
}

// 5) Test di regressione per il caso mostrato: 42 lavorazioni, 30 impianti, una sola commessa Modena.
{
  const path = "scripts/check-commesse-impianti-integrity.js";
  let source = fs.readFileSync(path, "utf8");

  source = replaceOnce(
    source,
    `const sandbox = {\n  window: {`,
    `const sandbox = {\n  commesseById: new Map([\n    ["inrete_modena_agosto_2026", { id: "inrete_modena_agosto_2026", nome: "INRETE MODENA - AGOSTO 2026", codice: "INRETE-MO-AGO-2026", attiva: true }],\n    ["canonical-modena", { id: "canonical-modena", nome: "INRETE MODENA", codice: "28015", attiva: true, impiantiCount: 30, impiantiFattiCount: 3 }]\n  ]),\n  impiantiByCommessaId: new Map([["inrete_modena_agosto_2026", [{ id: "legacy" }]]]),\n  commessaStatsById: new Map([["inrete_modena_agosto_2026", { total: 30, done: 0 }]]),\n  window: {`,
    "sandbox commessa canonica"
  );

  source = replaceOnce(
    source,
    `assert.ok(api, "Runtime INRETE Modena non esportato");\n`,
    `assert.ok(api, "Runtime INRETE Modena non esportato");\nassert.equal(api.testing.findCanonicalModenaCommessa().id, "canonical-modena");\nassert.equal(api.ensureVisibleLocally(), true);\nassert.equal(sandbox.commesseById.has("inrete_modena_agosto_2026"), false, "La commessa sintetica non deve restare visibile");\nassert.equal(sandbox.commesseById.has("canonical-modena"), true, "La commessa 28015 deve restare attiva");\nassert.equal(sandbox.impiantiByCommessaId.has("inrete_modena_agosto_2026"), false);\n`,
    "test rimozione duplicato"
  );

  source = replaceOnce(
    source,
    `const statsCacheSource = fs.readFileSync("commessa-stats-cache-optimizer.js", "utf8");\n`,
    `const workItemsCore = require("../inrete-work-items-v2.js");\nconst workRows = Array.from({ length: 42 }, (_, index) => ({\n  id: `work-${index + 1}`,\n  idSap: String(1000000 + (index < 30 ? index : index - 30)),\n  stato: index < 4 ? "FATTO" : "DA FARE"\n}));\nconst workSummary = workItemsCore.summarizeWorkItems(workRows);\nassert.equal(workSummary.uniquePlants, 30, "42 lavorazioni ripetute devono corrispondere a 30 impianti fisici");\nassert.equal(workSummary.workRows, 42);\nassert.equal(workSummary.doneWork, 4);\nassert.equal(workSummary.todoWork, 38);\n\nconst statsCacheSource = fs.readFileSync("commessa-stats-cache-optimizer.js", "utf8");\n`,
    "test 42 righe 30 impianti"
  );

  source = replaceOnce(
    source,
    `assert.match(statsCacheSource, /function commessaRef\\(commessaId\\)/, "Manca il riferimento centralizzato alla commessa attiva");\n`,
    `assert.match(statsCacheSource, /function commessaRef\\(commessaId\\)/, "Manca il riferimento centralizzato alla commessa attiva");\nassert.match(statsCacheSource, /function statsFromCommessaMetadata\\(commessaId, baseStats = \\{\\}\\)/, "La home deve usare i contatori parent quando la cache non è pronta");\nassert.match(statsCacheSource, /const ids = new Set\\(\\[selected\\]\\)/, "La commessa aperta deve essere inclusa nel caricamento statistiche");\n`,
    "test fallback statistiche"
  );

  source = replaceOnce(
    source,
    `const globalArchiveSource = fs.readFileSync("global-archive-sync.js", "utf8");\n`,
    `const accountingSource = fs.readFileSync("accounting-v2.js", "utf8");\nassert.match(accountingSource, /summary\\.uniquePlants/, "La gestione deve mostrare il numero di impianti unici");\nassert.match(accountingSource, /summary\\.workRows/, "La gestione deve mostrare separatamente le lavorazioni");\nassert.match(accountingSource, /syncParentSummary/, "I contatori della commessa devono essere riallineati");\nassert.doesNotMatch(accountingSource, /<b>\\$\\{state\\.work\\.length\\}<\\/b> righe/, "La UI non deve più chiamare impianti le righe di lavorazione");\n\nassert.match(source, /CANONICAL_COMMESSA_CODE = "28015"/);\nassert.match(source, /archiveSyntheticDuplicate/);\nconst ensureVisibleBody = source.match(/function ensureVisibleLocally\\(\\) \\{[\\s\\S]*?\\n  \\}/)?.[0] || "";\nassert.doesNotMatch(ensureVisibleBody, /commesseById\\.set\\(COMMESSA_ID/, "Il runtime non deve ricreare la commessa sintetica");\n\nconst globalArchiveSource = fs.readFileSync("global-archive-sync.js", "utf8");\n`,
    "test UI e duplicato"
  );

  source = source.replace(
    "cache, Global e backfill non distruttivo verificati.",
    "conteggio univoco, commessa canonica, cache, Global e backfill non distruttivo verificati."
  );
  write(path, source);
}

// 6) Versioni PWA allineate per distribuire subito la correzione.
{
  const version = "20260819-inrete-canonical-count1";
  const indexPath = "index.html";
  let index = fs.readFileSync(indexPath, "utf8");
  const replacements = [
    ["commessa-stats-cache-optimizer.js?v=20260815a", `commessa-stats-cache-optimizer.js?v=${version}`],
    ["inrete-work-items-v2.js?v=20260728b", `inrete-work-items-v2.js?v=${version}`],
    ["accounting-v2.js?v=20260812-modena2", `accounting-v2.js?v=${version}`],
    ["operational-import-repair.js?v=20260728a", `operational-import-repair.js?v=${version}`],
    ["navigator.serviceWorker.register(\"./sw.js?v=20260814-loading-humor1\", { updateViaCache: \"none\" })", `navigator.serviceWorker.register("./sw.js?v=${version}", { updateViaCache: "none" })`]
  ];
  for (const [before, after] of replacements) index = replaceOnce(index, before, after, `index ${before}`);
  write(indexPath, index);

  const swPath = "sw.js";
  let sw = fs.readFileSync(swPath, "utf8");
  sw = sw.replace(/const CACHE_NAME = "varga-cantieri-shell-v\d+";/, 'const CACHE_NAME = "varga-cantieri-shell-v137";');
  sw = sw.replace(/const CACHE_RESET_VERSION = "[^"]+";/, `const CACHE_RESET_VERSION = "${version}";`);
  for (const [before, after] of replacements.slice(0, 4)) sw = replaceOnce(sw, `./${before}`, `./${after}`, `sw ${before}`);
  write(swPath, sw);
}

console.log("Correzione commessa INRETE Modena canonica e conteggi univoci applicata.");
