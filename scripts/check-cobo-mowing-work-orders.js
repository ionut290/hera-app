#!/usr/bin/env node
"use strict";

const assert = require("node:assert/strict");
const fs = require("node:fs");
const vm = require("node:vm");

const read = (file) => fs.readFileSync(file, "utf8");
const source = read("cobo-mowing-work-orders.js");
const css = read("cobo-mowing-work-orders.css");
const app = read("app.js");
const baseParks = read("verde-bologna.js");
const mobileParks = read("verde-bologna-parchi-mobile.js");
const html = read("index.html");
const serviceWorker = read("sw.js");

const checks = [
  [source.includes('const COMMESSA_ID = "sfalcio-cobo"'), "ID stabile della commessa Sfalcio COBO"],
  [source.includes('const COMMESSA_NAME = "Sfalcio COBO"'), "nome della commessa"],
  [source.includes('const COMMESSA_CODE = "COBO-SFALCIO"'), "codice della commessa"],
  [source.includes('.doc(COMMESSA_ID)') && source.includes('.collection("impianti").doc(plantId)'), "percorso operativo standard commessa/impianti"],
  [source.includes('tipoSpeciale: "SFALCIO_COBO"') && source.includes("parchiGiardini: true"), "collegamento esplicito a Parchi e giardini"],
  [source.includes('stato: "da_fare", done: false, doneAt: null, doneBy: ""') && source.includes("const timestamps = existing"), "stato iniziale applicato solo ai nuovi cantieri"],
  [source.includes("hasPayloadChanges(existing, payload)") && source.includes("if (writes) await batch.commit()"), "creazione idempotente senza scritture duplicate"],
  [source.includes("sfalcioCoboRegistrato: true") && source.includes("sfalcioCoboTipoEsecuzione"), "registrazione preparatoria dello sfalcio"],
  [source.includes("if (!hasPayloadChanges(plant, comparable)) return { changed: false"), "registrazione invariata senza nuova scrittura"],
  [baseParks.includes("data-vb-create-cobo") && baseParks.includes("data-vb-cobo-index"), "CREA CANTIERE nella ricerca desktop"],
  [mobileParks.includes("data-vb-create-cobo") && mobileParks.includes("openCoboWorkOrder"), "CREA CANTIERE nella scheda mobile"],
  [app.includes("function isCoboMowingWorkOrder(impianto)") && app.includes("openCoboMowingRegistrationForm"), "integrazione limitata alla commessa COBO"],
  [app.includes("🧾 REGISTRA SFALCIO") && app.includes("cobo-register-sfalcio-btn"), "pulsante preparatorio nella normale scheda commessa"],
  [app.includes("buildCoboMowingCardDetailsMarkup") && app.includes("badge-cobo-sfalcio"), "riepilogo registrazione nel cantiere"],
  [html.includes("cobo-mowing-work-orders.css?v=20260901-cobo-sfalcio1") && html.includes("cobo-mowing-work-orders.js?v=20260902-special-terminato1"), "modulo e stile caricati nell’app"],
  [serviceWorker.includes("cobo-mowing-work-orders.css?v=20260901-cobo-sfalcio1") && serviceWorker.includes("cobo-mowing-work-orders.js?v=20260902-special-terminato1"), "modulo disponibile nella PWA"],
  [css.includes(".cobo-mowing-modal") && css.includes(".cobo-register-sfalcio-btn"), "interfaccia responsive dedicata"],
  [!source.includes(".onSnapshot(") && !source.includes("setInterval(") && !source.includes("watchPosition("), "nessun listener o polling aggiunto"],
  [!/(?:function\s+markImpiantoDone|function\s+handleImpiantoWhatsAppClick|window\.openWhatsApp\s*=|safeOpenWhatsAppMessage\s*=)/.test(source), "nessuna ridefinizione FATTO o Whazzup"]
];

const registrationFunction = source.match(/async function saveRegistration\(options, values\) \{[\s\S]*?\n  \}\n\n  function openRegistration/);
checks.push([Boolean(registrationFunction), "funzione di registrazione individuata"]);
if (registrationFunction) {
  checks.push([!/["'](?:done|doneAt|doneBy|stato)["']?\s*:/.test(registrationFunction[0]), "REGISTRA SFALCIO non modifica lo stato FATTO"]);
  checks.push([!registrationFunction[0].includes("Whazzup") && !registrationFunction[0].includes("WhatsApp"), "REGISTRA SFALCIO non apre Whazzup"]);
  checks.push([!registrationFunction[0].includes(".get("), "REGISTRA SFALCIO non aggiunge letture Firestore"]);
}

const sandbox = {
  window: { currentUser: { uid: "utente-1", email: "operatore@example.com", displayName: "Operatore Test" } },
  console,
  Date,
  FormData: class {},
  URLSearchParams
};
vm.runInNewContext(source, sandbox, { filename: "cobo-mowing-work-orders.js" });
const api = sandbox.window.HeraCoboMowing;
const sampleContext = {
  record: { codvia: "1234", nomevia: "Giardino Test", tipo: "GIARDINO" },
  parkName: "Giardino Test",
  parkCode: "1234",
  quarter: "Navile",
  point: { lat: 44.5123, lon: 11.3456 },
  boundaryAvailable: true
};
const payload = api.buildPlantPayload(sampleContext, {
  workType: "SFALCIO COMPLETO",
  requestedWork: "Sfalcio completo del parco",
  areaMq: "1250,5",
  operatorNote: "Ingresso da via Test"
});

checks.push([api.installed === true && api.commessaId === "sfalcio-cobo", "API COBO installata"]);
checks.push([api.plantIdFor(sampleContext) === api.plantIdFor({ ...sampleContext }), "ID cantiere deterministico"]);
checks.push([payload.denominazione === "Giardino Test" && payload.parcoCodvia === "1234", "dati ufficiali del parco copiati"]);
checks.push([payload.gpsY === 44.5123 && payload.gpsX === 11.3456, "coordinate copiate nel formato standard"]);
checks.push([payload.sfalciMq === 1250.5 && payload.tipoManutenzione === "Ordinaria" && payload.codicePrezzo === "A11", "superficie e manutenzione ordinaria valorizzate"]);
checks.push([!("done" in payload) && !("doneAt" in payload) && !("doneBy" in payload), "payload di aggiornamento non sovrascrive FATTO"]);
checks.push([api.isWorkOrder(payload, "sfalcio-cobo") === true && api.isWorkOrder(payload, "altra-commessa") === false, "pulsante limitato alla commessa corretta"]);

const failed = checks.filter(([ok]) => !ok);
checks.forEach(([ok, label]) => console.log(`${ok ? "OK" : "FAIL"} ${label}`));
if (failed.length) process.exit(1);
console.log("✅ Sfalcio COBO: ricerca parchi, creazione cantiere, registrazione preparatoria e isolamento FATTO verificati.");
