#!/usr/bin/env node
"use strict";

const fs = require("node:fs");
const path = require("node:path");

const ROOT = path.resolve(__dirname, "..");

const REMOVALS = {
  "app.js": [
`function openBannedAccessRequest() {
  window.open(buildBannedWhatsAppUrl(), "_blank", "noopener,noreferrer");
}
`,
`function isCurrentUserBanned() {
  return Boolean(currentUserBanProfile?.banned);
}
`,
`function urlBase64ToUint8Array(base64String) {
  const padding = "=".repeat((4 - (base64String.length % 4)) % 4);
  const base64 = (base64String + padding).replace(/-/g, "+").replace(/_/g, "/");
  const rawData = atob(base64);
  return Uint8Array.from([...rawData].map((char) => char.charCodeAt(0)));
}
`,
`function getCurrentUserIdentityParts() {
  if (!currentUser) return [];
  const currentProfile = platformUsers.find((user) => String(user.id || user.uid || "") === String(currentUser.uid || ""));
  return getPlatformUserIdentityParts(currentProfile || currentUser);
}
`,
`function getHoursOperatorForCurrentUser(commessaId, dateValue = "") {
  const assignment = getCurrentUserSquadraAssignment(commessaId, dateValue);
  return assignment?.matchedName || getCurrentUserResolvedName();
}
`,
`function getCommesseErrorMessage() {
  return "Impossibile caricare le commesse online. Mostro dati salvati localmente.";
}
`,
`function getCommessaHoursTotal(commessaId) {
  return Number(getCommessaWorkSummary(commessaId).totalHours || 0);
}
`,
`function shareGlobalImpiantoViaWhatsapp(impianto) {
  handleOpenGlobalSegnalazioneClick(impianto);
}
`,
`function toggleCommessaNoteForm() {
  openCommessaNotesPage();
}
`,
`function formatWeatherDetailValue(value, suffix = "") {
  if (!isPresentFiniteNumber(value)) return "-";
  return \`${Math.round(Number(value))}${suffix}\`;
}
`,
`function formatWeatherAmount(value) {
  const amount = Number(value);
  if (!Number.isFinite(amount)) return "";
  return amount >= 10 ? String(Math.round(amount)) : amount.toFixed(1);
}
`,
`function getTimestampDate(value) {
  if (!value) return null;
  if (typeof value.toDate === "function") return value.toDate();
  if (typeof value.seconds === "number") return new Date(value.seconds * 1000);
  const parsed = new Date(value);
  return Number.isNaN(parsed.getTime()) ? null : parsed;
}
`,
`function openWeatherModal() {
  ui.weatherModal.classList.remove("hidden");
  ui.weatherModal.setAttribute("aria-hidden", "false");
}
`,
`function renderSnowServiceList(element, rows, emptyText, renderRow) {
  if (!element) return;
  element.innerHTML = rows.length
    ? rows.map(renderRow).join("")
    : \`<p class='muted'>${escapeHTML(emptyText)}</p>\`;
}
`
  ],
  "today-summary-interactions.js": [
`  function getLiveWorkedMinutes(assignments) {
    const start = getAssignedStartMinutes(assignments);
    if (start === null) return null;
    const elapsed = Math.max(0, getRomeClockMinutes() - start);
    return elapsed > 8 * 60 ? elapsed - 60 : elapsed;
  }
`
  ]
};

function removeExactBlock(source, block, fileName) {
  const normalizedSource = source.replace(/\r\n/g, "\n");
  const normalizedBlock = block.replace(/\r\n/g, "\n");
  const first = normalizedSource.indexOf(normalizedBlock);
  if (first < 0) throw new Error(`Blocco non trovato in ${fileName}: ${normalizedBlock.split("\n")[0]}`);
  if (normalizedSource.indexOf(normalizedBlock, first + normalizedBlock.length) >= 0) {
    throw new Error(`Blocco duplicato in ${fileName}: ${normalizedBlock.split("\n")[0]}`);
  }
  let result = normalizedSource.slice(0, first) + normalizedSource.slice(first + normalizedBlock.length);
  if (result[first] === "\n") result = result.slice(0, first) + result.slice(first + 1);
  return result;
}

let removed = 0;
for (const [relativePath, blocks] of Object.entries(REMOVALS)) {
  const filePath = path.join(ROOT, relativePath);
  let source = fs.readFileSync(filePath, "utf8");
  for (const block of blocks) {
    source = removeExactBlock(source, block, relativePath);
    removed += 1;
    console.log(`RIMOSSO ${relativePath} :: ${block.match(/function\s+([\w$]+)/)?.[1] || "blocco"}`);
  }
  fs.writeFileSync(filePath, source, "utf8");
}

console.log(`Rimosse ${removed} funzioni semplici confermate come inutilizzate.`);
