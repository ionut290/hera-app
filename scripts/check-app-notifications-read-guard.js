"use strict";

const fs = require("node:fs");

const guard = fs.readFileSync("app-notifications-read-guard.js", "utf8");
const config = fs.readFileSync("firebase-config.js", "utf8");
const sw = fs.readFileSync("sw.js", "utf8");

const checks = [
  [guard.includes('collection: "appNotifications"'), "la guardia deve puntare ad appNotifications"],
  [guard.includes('property === "onSnapshot"'), "la guardia deve bloccare i listener legacy"],
  [guard.includes('property === "get"'), "la guardia deve bloccare le letture one-shot legacy"],
  [guard.includes("FirestoreCtor") && guard.includes("prototype.collection"), "la guardia deve intercettare soltanto l'accesso alla collection"],
  [!guard.includes('property === "add"'), "la guardia non deve bloccare add/scritture"],
  [!guard.includes('property === "set"'), "la guardia non deve bloccare set/scritture"],
  [!guard.includes('property === "update"'), "la guardia non deve bloccare update/scritture"],
  [!guard.includes('property === "delete"'), "la guardia non deve bloccare delete/scritture"],
  [config.includes("HERA_APP_NOTIFICATIONS_READ_GUARD_SRC"), "firebase-config deve caricare la guardia"],
  [config.includes('data-app-notifications-read-guard="1"'), "la guardia deve essere caricata prima di app.js"],
  [sw.includes("app-notifications-read-guard.js?v=20260815a"), "la shell PWA deve includere la guardia"]
];

const failed = checks.filter(([ok]) => !ok).map(([, message]) => message);
if (failed.length) {
  console.error("Controllo appNotifications read guard NON superato:");
  failed.forEach((message) => console.error(`- ${message}`));
  process.exit(1);
}

console.log("Controllo appNotifications read guard superato.");
