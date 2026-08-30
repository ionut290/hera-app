"use strict";

const fs = require("fs");
const path = require("path");

const root = path.resolve(__dirname, "..");
const read = (file) => fs.readFileSync(path.join(root, file), "utf8");
const assert = (condition, message) => {
  if (!condition) throw new Error(message);
};

const html = read("index.html");
const client = read("green-assistant.js");
const css = read("green-assistant.css");
const backend = read("netlify/functions/green-assistant.js");
const netlify = read("netlify.toml");
const serviceWorker = read("sw.js");

assert(html.includes('id="open-gardening-assistant-btn"'), "Voce menu giardinaggio assente.");
assert(html.includes('id="open-equipment-assistant-btn"'), "Voce menu mezzi assente.");
assert(html.includes('green-assistant.css?v=20260830-brave-manuals3'), "CSS assistente non caricato.");
assert(html.includes('green-assistant.js?v=20260830-brave-manuals3'), "Client assistente non caricato.");
assert(netlify.includes('from = "/api/green-assistant"'), "Redirect Netlify assistente assente.");
assert(netlify.includes('to = "/.netlify/functions/green-assistant"'), "Funzione Netlify assistente non collegata.");

for (const action of ["status", "identifyPlant", "identifyDisease", "searchPlant", "plantDetails", "equipmentInfo", "equipmentManuals"]) {
  assert(backend.includes(`action === "${action}"`), `Azione backend mancante: ${action}`);
}
for (const secret of ["PLANTNET_API_KEY", "TREFLE_API_TOKEN", "GEMINI_API_KEY", "BRAVE_SEARCH_API_KEY"]) {
  assert(backend.includes(`process.env.${secret}`) || backend.includes(`requireSecret("${secret}"`), `Segreto protetto mancante: ${secret}`);
  assert(!client.includes(secret), `Il segreto ${secret} non deve essere nel client.`);
  assert(!html.includes(secret), `Il segreto ${secret} non deve essere nell'HTML.`);
}
assert(backend.includes("authenticateEvent(event)"), "La funzione deve richiedere autenticazione Firebase.");
assert(backend.includes('"Cache-Control": "no-store"'), "Le risposte API non devono essere memorizzate in cache.");
assert(backend.includes('"gemini-3.5-flash-lite"'), "Il modello Gemini gratuito corrente deve essere configurato.");
assert(!backend.includes('"gemini-2.5-flash-lite"'), "Il modello Gemini non disponibile ai nuovi utenti non deve essere usato.");
assert(client.includes('const CACHE_KEY = "heraGreenAssistantArchiveV1"'), "Archivio locale assistente assente.");
assert(client.includes('window.openManagementPanel("mezzi")'), "Copia controllata nel form mezzi assente.");
assert(client.includes('typeof mezziRecords !== "undefined"'), "Riutilizzo locale del registro mezzi assente.");
assert(client.includes("openEquipment"), "Collegamento Mezzi → Assistente assente.");
assert(client.includes('id="green-equipment-code" type="text"'), "Il codice mezzo deve essere digitabile anche sui browser mobili.");
assert(!client.includes('id="green-equipment-code" type="search"'), "Il codice mezzo non deve usare il campo search incompatibile con alcuni datalist mobili.");
assert(client.includes("[payload.brand, payload.model, payload.year, payload.type]"), "La cache manuali deve distinguere il tipo di mezzo.");
assert(client.includes("API PROTETTE"), "L'interfaccia non deve presentare Brave come servizio gratuito.");
assert(client.includes("data.notice"), "L'avviso sul piano Brave deve essere mostrato all'utente.");
assert(!backend.includes("hostnameKey.includes(brandKey)"), "Un dominio non deve essere considerato ufficiale solo perché contiene il marchio.");
assert(html.includes("management-v2.js?v=20260830-brave-manuals3"), "Versione gestione mezzi non aggiornata.");
assert(read("management-v2.js").includes("data-equipment-assistance"), "Pulsante assistenza nelle schede mezzo assente.");

for (const forbidden of [".collection(", ".doc(", ".onSnapshot(", "firebase.firestore(", "getFirestore(", "setDoc(", "addDoc("]) {
  assert(!client.includes(forbidden), `Operazione Firestore vietata nel client assistente: ${forbidden}`);
  assert(!backend.includes(forbidden), `Operazione Firestore vietata nel backend assistente: ${forbidden}`);
}
for (const protectedToken of ["fatto-button", "whazzup", "whatsapp", "done-button"]) {
  assert(!client.toLowerCase().includes(protectedToken), `Riferimento al flusso protetto rilevato: ${protectedToken}`);
  assert(!backend.toLowerCase().includes(protectedToken), `Riferimento backend al flusso protetto rilevato: ${protectedToken}`);
}

assert(css.includes(".green-assistant-overlay"), "Stile overlay assistente assente.");
assert(serviceWorker.includes('"./green-assistant.js?v=20260830-brave-manuals3"'), "Client assistente assente dalla shell PWA.");
assert(serviceWorker.includes('"./green-assistant.css?v=20260830-brave-manuals3"'), "CSS assistente assente dalla shell PWA.");

console.log("Assistenti giardinaggio e mezzi: controlli statici superati.");
