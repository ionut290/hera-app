#!/usr/bin/env node
"use strict";

const fs = require("node:fs");
const source = fs.readFileSync("admin-user-access-tools.js", "utf8");

function requireText(text, message) {
  if (!source.includes(text)) throw new Error(message);
}

function forbidText(text, message) {
  if (source.includes(text)) throw new Error(message);
}

requireText("let panelActivated = false;", "Manca il gate del pannello Gestione utenti.");
requireText("if (!panelActivated || !isManager()) return;", "enhance() non è protetto dal gate on-demand.");
requireText("observer.observe(panel, { childList: true, subtree: true });", "L'observer deve restare limitato al pannello utenti.");
requireText("window.setTimeout(activateUserPanel, 120);", "L'attivazione on-demand dal pulsante Gestione utenti è assente.");
requireText("firebase.firestore().collection(\"platformUsers\").get()", "Il fallback profili su richiesta è stato rimosso: Gestione utenti deve poter caricare i profili quando aperta.");
forbidText("observer.observe(document.body", "Non reintrodurre un observer globale che può avviare letture profili durante il rendering.");

const initializeMatch = source.match(/function initialize\(\) \{([\s\S]*?)\n  \}\n\n  window\.HeraAdminUserAccessTools/);
if (!initializeMatch) throw new Error("Funzione initialize() non trovata.");
const initializeBody = initializeMatch[1];
if (/\bvoid\s+enhance\(\)/.test(initializeBody)) {
  throw new Error("Gestione utenti non deve eseguire enhance() direttamente all'avvio.");
}
if (/\bloadProfiles\(\)/.test(initializeBody)) {
  throw new Error("Gestione utenti non deve leggere platformUsers direttamente all'avvio.");
}

console.log("OK: i profili completi platformUsers restano on-demand nella Gestione utenti.");
