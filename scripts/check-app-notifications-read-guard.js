"use strict";

const fs = require("node:fs");

const guard = fs.readFileSync("app-notifications-read-guard.js", "utf8");
const config = fs.readFileSync("firebase-config.js", "utf8");
const sw = fs.readFileSync("sw.js", "utf8");

const required = [
  [guard.includes('collection: "appNotifications"'), "la guardia deve puntare ad appNotifications"],
  [guard.includes('property === "onSnapshot"'), "la guardia deve bloccare i listener"],
  [guard.includes('property === "get"'), "la guardia deve bloccare le letture one-shot"],
  [!guard.includes('property === "add"'), "la guardia non deve bloccare add/scritture"],
  [!guard.includes('property === "set"'), "la guardia non deve bloccare set/scritture"],
  [config.includes("HERA_APP_NOTIFICATIONS_READ_GUARD_SRC"), "firebase-config deve caricare la guardia"],
  [config.indexOf("HERA_APP_NOTIFICATIONS_READ_GUARD_SRC") < config.indexOf("HERA_FIRESTORE_OPERATION_DIAGNOSTICS_SRC"), "la guardia deve essere dichiarata prima della diagnostica"],
  [config.includes('data-app-notifications-read-guard="1"'), "la guardia deve essere caricata nel percorso iniziale"],
  [sw.includes("app-notifications-read-guard.js?v=20260815a"), "il service worker deve includere la guardia nella shell"]
];

const failed = required.filter(([ok]) => !ok).map(([, message]) => message);
if (failed.length) {
  console.error("Controllo appNotifications read guard NON superato:");
  failed.forEach((message) => console.error(`- ${message}`));
  process.exit(1);
}

console.log("Controllo appNotifications read guard superato.");
