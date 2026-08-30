const fs = require("fs");
const path = require("path");
const root = path.resolve(__dirname, "..");
const code = fs.readFileSync(path.join(root, "lavori-occasionali-workflow-v2.js"), "utf8");
const loader = fs.readFileSync(path.join(root, "firebase-config.js"), "utf8");
const checks = [
  [code.includes('Lavori occasionali'), "manca la nuova voce/menu Lavori occasionali"],
  [code.includes('root().collection("impianti")'), "il catalogo deve usare solo gli impianti della commessa occasionale"],
  [code.includes('collection("assegnazioniOccasionali")'), "manca l'associazione cantieri alle squadre"],
  [code.includes('assignmentVersion:3'), "manca il formato assegnazione v3"],
  [code.includes('Fine: salva composizione'), "la scelta cantieri deve restare dentro il flusso composizione normale"],
  [code.includes('className="btn btn-primary ocv2-hours-btn"'), "manca + ORE sui cantieri della scheda squadra"],
  [code.includes('collection("oreCantiere")'), "le ore devono essere salvate per singolo cantiere"],
  [code.includes('Totale giornata Lavori occasionali'), "manca il totale ore giornaliero"],
  [code.includes('Ore Lavori occasionali'), "manca il riepilogo ore nel calendario personale"],
  [code.includes('hoursFor(p.id,day,true)') && code.includes('operatoreUid'), "il calendario deve filtrare le ore dell'utente corrente"],
  [code.includes('let plantsPromise = null') && code.includes('if (plantsPromise) return plantsPromise'), "le richieste simultanee del catalogo devono condividere una sola lettura"],
  [loader.includes('HERA_OCCASIONAL_WORKFLOW_V2_SRC'), "firebase-config non carica il nuovo workflow"],
  [!/(?:safeOpenWhatsAppMessage|openWhatsApp|markImpiantoDone|forceMoveImpiantoToFatti|handleImpiantoWhatsAppClick)/.test(code), "il workflow non deve modificare FATTO o Whazzup"],
  [!/(?:currentImpianti\s*=|impiantiByCommessaId\.(?:set|delete|clear)\s*\()/.test(code), "il workflow non deve contaminare le cache globali"],
  [!code.includes('onSnapshot('), "il workflow non deve introdurre listener Firestore persistenti"]
];
const failed = checks.filter(([ok]) => !ok).map(([,msg]) => msg);
if (failed.length) {
  console.error("❌ Workflow Lavori occasionali v2 non valido:");
  failed.forEach((m) => console.error(`- ${m}`));
  process.exit(1);
}
console.log("✅ Workflow Lavori occasionali v2 OK: catalogo, squadre, cantieri, ore e calendario isolati.");
