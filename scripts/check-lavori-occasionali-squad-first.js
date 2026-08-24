const fs = require("fs");
const path = require("path");

const root = path.resolve(__dirname, "..");
const flow = fs.readFileSync(path.join(root, "lavori-occasionali-squad-first.js"), "utf8");
const core = fs.readFileSync(path.join(root, "lavori-occasionali.js"), "utf8");
const sites = fs.readFileSync(path.join(root, "squadre-lavori-occasionali-cantieri.js"), "utf8");
const loader = fs.readFileSync(path.join(root, "firebase-config.js"), "utf8");

const checks = [
  [flow.includes('form.insertAdjacentElement("afterend", panel)'), "il modulo cantieri deve stare dopo il form squadra"],
  [flow.includes('.collection("assegnazioniOccasionali")'), "il cantiere deve essere assegnato nella raccolta dedicata"],
  [flow.includes("squadraIndex: index"), "l'assegnazione deve conservare la squadra scelta"],
  [flow.includes("compositionRows()"), "il salvataggio deve richiedere una squadra già salvata"],
  [flow.includes("core.upsertPlant(metadata)"), "il modulo deve creare l'impianto nella commessa Lavori occasionali"],
  [core.includes("if (window.HeraOccasionalSquadFirstFlow?.installed) return;\n    if (!isOccasionalSelected()) return;"), "il salvataggio squadra non deve richiedere il cantiere"],
  [core.includes("upsertPlant: upsertNormalOccasionalPlant"), "l'API cantieri deve esporre il salvataggio controllato"],
  [sites.includes("function extractSiteName(value, seen = new WeakSet())"), "l'elenco deve estrarre il nome anche dai vecchi oggetti cantiere"],
  [sites.includes('raw.toLocaleUpperCase("it-IT") === "[OBJECT OBJECT]"'), "l'elenco deve scartare il testo tecnico object Object"],
  [sites.includes("name: extractSiteName(data.cantiere)"), "le assegnazioni legacy devono leggere il nome interno del cantiere"],
  [!sites.includes('String(item?.name || item?.cantiere || item?.nome || item || "")'), "l'elenco non deve convertire direttamente un oggetto in testo"],
  [loader.includes("HERA_OCCASIONAL_SQUAD_FIRST_SRC"), "firebase-config deve caricare il nuovo modulo"],
  [loader.includes('data-occasional-squad-first="1"'), "il caricamento sincrono deve installare il nuovo flusso prima del modulo storico"],
  [loader.includes('squadre-lavori-occasionali-cantieri.js?v=20260824b'), "il correttivo nomi cantieri deve usare una nuova versione cache"],
  [!flow.includes("onSnapshot("), "il modulo non deve aggiungere listener Firestore persistenti"],
  [!/(?:currentImpianti\s*=|currentImpianti\.(?:push|splice|pop|shift|unshift|sort|reverse)\s*\(|impiantiByCommessaId\.(?:set|delete|clear)\s*\()/.test(flow), "il modulo non deve mutare le cache globali degli impianti"],
  [!/(?:safeOpenWhatsAppMessage|openWhatsApp|markImpiantoDone|forceMoveImpiantoToFatti|handleImpiantoWhatsAppClick)/.test(flow), "il modulo non deve modificare FATTO o Whazzup"]
];

const failed = checks.filter(([ok]) => !ok).map(([, message]) => message);
if (failed.length) {
  console.error("❌ Controllo flusso Lavori occasionali fallito:");
  failed.forEach((message) => console.error(`- ${message}`));
  process.exit(1);
}

console.log("✅ Flusso Lavori occasionali OK: prima squadra, poi uno o più cantieri associati, senza [object Object].");
