#!/usr/bin/env node
"use strict";

const assert = require("node:assert/strict");
const fs = require("node:fs");
const vm = require("node:vm");

const html = fs.readFileSync("index.html", "utf8");
const feature = fs.readFileSync("commessa-impianti-menu.js", "utf8");
const accounting = fs.readFileSync("accounting-v2.js", "utf8");
const css = fs.readFileSync("style.css", "utf8");
const serviceWorker = fs.readFileSync("sw.js", "utf8");

assert.match(html, /id="commessa-plants-menu-btn"/);
for (const action of ["add", "edit", "import", "export", "prices", "advanced"]) {
  assert.match(feature, new RegExp(`data-commessa-mobile-action="${action}"`));
}
assert.match(feature, /canManageData\(\)/);
assert.match(feature, /openManagementPanel\("commesse"\)/);
assert.match(feature, /AccountingV2\.openMobileHub\(commessa\)/);
assert.match(feature, /id="commessa-mobile-management"/);
assert.match(accounting, /mobileGroups/);
assert.match(accounting, /Dati impianto/);
assert.match(accounting, /Lavorazione/);
assert.match(accounting, /Esecuzione/);
assert.match(accounting, /async function createMobilePlant/);
assert.match(accounting, /async function createMobileWork/);
assert.match(accounting, /function operationalProjectionDrift/);
assert.match(accounting, /async function saveMobilePlant/);
assert.match(accounting, /function ensureMobileAddWorkButton/);
assert.match(accounting, /async function autofillMobileAddress/);
assert.match(accounting, /reverseGeocodeMobileCoordinates/);
assert.match(accounting, /mobileAddressFromGeocode/);
assert.match(accounting, /id="commessa-mobile-geocode-status"/);
assert.match(accounting, /Comune e Via compilati automaticamente/);
assert.match(accounting, /mobileManualValue/);
assert.match(accounting, /accept-language=it/);
assert.match(accounting, /document\.createElement\("button"\)/);
assert.match(feature, /id="commessa-mobile-add-work"/);
assert.match(accounting, /codiceVocePrezzo:""/);
assert.match(accounting, /quantita/);
const createWorkBody = accounting.slice(accounting.indexOf("async function createMobileWork"), accounting.indexOf("async function saveMobilePlant"));
assert.match(createWorkBody, /physicalPlantId=String\(source\.impiantoId\|\|source\.physicalPlantId\|\|""\)\.trim\(\)/);
assert.match(createWorkBody, /impiantoId:physicalPlantId/);
assert.match(createWorkBody, /batch=db\.batch\(\)/);
assert.match(createWorkBody, /batch\.set\(ref,workDocument\)/);
assert.match(createWorkBody, /matchingOperationalPlants\(physicalPlantId,plant,previousItems\)/);
assert.match(createWorkBody, /operationalTargetIds\(physicalPlantId,plant,previousItems\)/);
assert.match(createWorkBody, /buildProjectionData\(updatedPlant,plantItems,matches,\{reopen:true\}\)/);
assert.match(createWorkBody, /projectionOperations\(commRef\(\),targetIds,projection,"management-new-work"\)/);
assert.match(createWorkBody, /batch\.set\(operation\.ref,operation\.data,\{merge:true\}\)/);
assert.match(createWorkBody, /collection\("impiantiFisici"\)\.doc\(physicalPlantId\)/);
assert.match(createWorkBody, /stato:"DA FARE",dataEsecuzione:"",oraEsecuzione:"",operatoreNome:""/);
assert.match(createWorkBody, /await load\(\{autoRepair:false\}\)/);
assert.match(createWorkBody, /state\.work\.some\(work=>work\.id===ref\.id/);
assert.match(createWorkBody, /targetIds\.every/);
assert.match(createWorkBody, /item\.numeroLavorazioni/);
assert.match(createWorkBody, /item\.done===false/);
assert.match(createWorkBody, /codeSignature\(item\.codicePrezzo\)/);
assert.match(createWorkBody, /L’impianto è tornato nell’elenco Da fare/);
assert.doesNotMatch(createWorkBody, /impiantoId:source\.impiantoId/);
assert.doesNotMatch(createWorkBody, /\.get\(\)|\.onSnapshot\(/);
assert.match(createWorkBody, /collection\("lavorazioni"\)\.doc\(\)/);
assert.doesNotMatch(createWorkBody, /\.delete\(/);
const createPlantBody = accounting.slice(accounting.indexOf("async function createMobilePlant"), accounting.indexOf("async function createMobileWork"));
assert.match(createPlantBody, /buildProjectionData\(\{\.\.\.plantDocument,id:p\.id\}/);
assert.match(createPlantBody, /projectionOperations\(commRef\(\),\[p\.id\],projection,"management-new-plant"\)/);
assert.match(createPlantBody, /batch\.set\(p,plantDocument\);batch\.set\(w,workDocument\)/);
assert.match(createPlantBody, /state\.operationalPlants\.some\(item=>item\.id===p\.id&&item\.done===false/);
assert.match(createPlantBody, /await load\(\{autoRepair:false\}\)/);
assert.doesNotMatch(createPlantBody, /Promise\.all|\.get\(\)|\.onSnapshot\(|\.delete\(/);
const loadBody = accounting.slice(accounting.indexOf("async function load"), accounting.indexOf("const clean="));
assert.match(loadBody, /operationalProjectionDrift\(\)/);
assert.match(loadBody, /commitOperations\(drift\.flatMap/);
assert.match(loadBody, /projectionOperations\(requestedRef,item\.targetIds,item\.data,"management-auto-repair"\)/);
assert.match(loadBody, /syncOperationalStateToMainView\(\)/);
assert.match(loadBody, /riallineati \$\{drift\.length\} impianti operativi/);
assert.doesNotMatch(loadBody, /operationalProjectionDrift\(\)[\s\S]*\.get\(\)|operationalProjectionDrift\(\)[\s\S]*\.onSnapshot\(/);
const projectionBody = accounting.slice(accounting.indexOf("function operationalProjectionDrift"), accounting.indexOf("async function commitOperations"));
assert.match(projectionBody, /matchingOperationalPlants\(id,plant,items\)/);
assert.match(projectionBody, /state\.work\.every\(item=>item\.legacy\)\)return \[\]/);
assert.match(projectionBody, /detailsMismatch/);
assert.match(projectionBody, /items\.length>1&&items\.length>maxCount/);
assert.match(projectionBody, /visibleOperationalItems\(items\)\.map\(item=>item\.codiceVocePrezzo\)/);
assert.match(projectionBody, /buildProjectionData\(\{\.\.\.plant,id\},items,matches,\{reopen\}\)/);
assert.doesNotMatch(projectionBody, /\.collection\(|\.onSnapshot\(|db\.batch\(/);
const linkageBody = accounting.slice(accounting.indexOf("const operationalIdentity="), accounting.indexOf("function operationalProjectionDrift"));
assert.match(linkageBody, /matchingOperationalPlants/);
assert.match(linkageBody, /String\(item\.physicalPlantId\|\|""\)===id/);
assert.match(linkageBody, /operationalIdentity\(item\)===identity/);
assert.match(linkageBody, /resetAt:server\(\)/);
assert.match(linkageBody, /stato:"DA FARE",statoGenerale:"DA FARE"/);
assert.match(linkageBody, /collection\("impiantoChangeIndex"\)\.doc\(id\)/);
assert.match(linkageBody, /changedAt:server\(\)/);
assert.doesNotMatch(linkageBody, /\.get\(\)|\.onSnapshot\(/);
assert.match(accounting, /renderImpiantiAfterRemoteSync\(state\.operationalPlants,\{value:null\}\)/);

const linkageRuntime = accounting.slice(accounting.indexOf("const operationalIdentity="), accounting.indexOf("const timestampMs="));
const linkageState = {operationalPlants:[{id:"legacy_doc",idSap:"SAP-1",denominazione:"Impianto",comune:"Galliera",indirizzo:"Via 1"}]};
const linkageApi = vm.runInNewContext(`(()=>{${linkageRuntime};return {matchingOperationalPlants,operationalTargetIds,codeSignature};})()`, {
  state: linkageState,
  norm: value => String(value || "").trim().toLowerCase()
});
const physicalPlant = {id:"physical_doc",idSap:"SAP-1",denominazione:"Impianto",comune:"Galliera",indirizzo:"Via 1"};
assert.deepEqual(Array.from(linkageApi.operationalTargetIds("physical_doc",physicalPlant)), ["legacy_doc"], "L’ID SAP deve riusare il documento operativo storico senza crearne uno parallelo.");
linkageState.operationalPlants.push({id:"physical_doc",physicalPlantId:"physical_doc",idSap:"SAP-1"});
assert.deepEqual(Array.from(linkageApi.operationalTargetIds("physical_doc",physicalPlant)).sort(), ["legacy_doc","physical_doc"], "Tutti i documenti già duplicati dello stesso impianto devono essere riallineati.");
assert.equal(linkageApi.codeSignature(["A11; B10","B10 | A11"]), "a11|b10");
const visibleRuntime = accounting.slice(accounting.indexOf("const visibleOperationalItems="), accounting.indexOf("const operationalPayload="));
const visibleOperationalItems = vm.runInNewContext(`(()=>{${visibleRuntime};return visibleOperationalItems;})()`);
assert.deepEqual(
  Array.from(visibleOperationalItems([{codiceVocePrezzo:"VECCHIO",stato:"FATTO"},{codiceVocePrezzo:"NUOVO",stato:"DA FARE"}])).map(item=>item.codiceVocePrezzo),
  ["NUOVO"],
  "Dopo la riapertura la scheda operativa deve mostrare la nuova lavorazione ancora da fare."
);
assert.equal(visibleOperationalItems([{stato:"FATTO"}]).length, 1, "Un impianto completato deve conservare la propria lavorazione visibile.");
const buildProjectionRuntime = accounting.slice(accounting.indexOf("function buildProjectionData"), accounting.indexOf("const changeIndexPayload="));
const buildProjectionDataRuntime = vm.runInNewContext(`(()=>{${buildProjectionRuntime};return buildProjectionData;})()`, {
  operationalPayload: () => ({done:true,stato:"FATTO",statoGenerale:"FATTO"}),
  plantStatus: items => items.some(item=>item.stato!=="FATTO") ? "DA FARE" : "FATTO",
  isOperationalDone: item => item.done===true,
  server: () => "SERVER_TIME",
  getOperatorDisplayName: () => "Admin",
  currentUser: {uid:"admin-1"}
});
const reopenedProjection = buildProjectionDataRuntime({}, [{stato:"FATTO"},{stato:"DA FARE"}], [{done:true}], {reopen:true});
assert.equal(reopenedProjection.stato, "DA FARE");
assert.equal(reopenedProjection.statoGenerale, "DA FARE");
assert.equal(reopenedProjection.done, false);
assert.equal(reopenedProjection.resetAt, "SERVER_TIME");
assert.equal(reopenedProjection.reopenedByManagement, true);
const saveRowStart = accounting.indexOf("async function saveRow");
const saveRowBody = accounting.slice(saveRowStart, accounting.indexOf("const feedback=", saveRowStart));
assert.match(saveRowBody, /physicalPlantId=String\(plant\?\.id\|\|old\.impiantoId\|\|""\)\.trim\(\)/);
assert.match(saveRowBody, /if\(!old\.legacy&&!physicalPlantId\)return feedback/);
assert.match(saveRowBody, /collection\("impiantiFisici"\)\.doc\(physicalPlantId\)/);
assert.match(saveRowBody, /commessaId:state\.commessa\.id/);
assert.match(saveRowBody, /batch=db\.batch\(\)/);
assert.match(saveRowBody, /projectionOperations\(commRef\(\),targetIds,projection,"management-edit"\)/);
assert.match(saveRowBody, /await load\(\{autoRepair:false\}\)/);
assert.doesNotMatch(saveRowBody, /Promise\.all|\.get\(\)|\.onSnapshot\(|\.delete\(/);
assert.doesNotMatch(saveRowBody, /\.doc\(plant\.id\)/);
const joinedBody = accounting.slice(accounting.indexOf("function joined"), accounting.indexOf("function rank"));
assert.match(joinedBody, /const plant=state\.plants\.find/);
assert.match(joinedBody, /plantFields\.forEach\(field=>\{if\(plant\[field\]!==undefined\)row\[field\]=plant\[field\]/);
assert.match(css, /\.commessa-dashboard-head \.commessa-plants-menu-wrap\s*{[^}]*position:\s*absolute/s);
assert.match(css, /\.commessa-mobile-plant-form/);
assert.doesNotMatch(feature, /\bdb\.|\bfirebase\.|\.collection\(|\.onSnapshot\(/);
assert.match(serviceWorker, /commessa-impianti-menu\.js\?v=20260826c/);
assert.match(serviceWorker, /accounting-v2\.js\?v=20260829-management-projection1/);
assert.match(serviceWorker, /style\.css\?v=20260826-new-work-sync1/);
assert.match(html, /accounting-v2\.js\?v=20260829-management-projection1/);

const accountingIndex = html.indexOf("accounting-v2.js");
const featureIndex = html.indexOf("commessa-impianti-menu.js");
const fattoIndex = html.indexOf("fatto-button-immediate.js");
assert.ok(accountingIndex >= 0 && featureIndex > accountingIndex, "Il menu deve riusare la gestione impianti già caricata.");
assert.ok(fattoIndex > featureIndex, "Il nuovo menu deve restare esterno al componente FATTO sigillato.");

console.log("✅ Gestione commessa mobile collegata ai dati esistenti con pulsante sovrapposto alla testata.");
