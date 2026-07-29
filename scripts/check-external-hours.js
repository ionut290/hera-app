const fs = require("fs");
const path = require("path");

const root = path.resolve(__dirname, "..");
const app = fs.readFileSync(path.join(root, "app.js"), "utf8");
const pkg = JSON.parse(fs.readFileSync(path.join(root, "package.json"), "utf8"));

const checks = [
  ["presenza dell’opzione COMMESSA ESTERNA", app.includes('externalOption.textContent = "COMMESSA ESTERNA"') && app.includes('data-hours-commessa-option="${HOURS_EXTERNAL_COMMESSA_VALUE}"')],
  ["presenza del campo Nome commessa esterna", app.includes("hours-external-commessa-name") && app.includes("NOME COMMESSA ESTERNA") && app.includes("Scrivi il nome della commessa")],
  ["validazione del nome obbligatorio", app.includes('externalNameInput.required = external') && app.includes("Inserisci il nome della commessa esterna.")],
  ["collectHoursEntries conserva una registrazione esterna senza commessaId", app.includes("commessaEsterna, tipoCommessa") && app.includes("commessaId: realCommessaId") && app.includes("entry.commessaEsterna || entry.rows.length")],
  ["finalizeHoursReport non filtra le commesse esterne", app.includes('(entry.commessaId || (entry.commessaEsterna === true && String(entry.commessaName || "").trim())) && entry.rows.length')],
  ["getHoursCommessaDisplayName produce ESTERNA · Nome", app.includes("function getHoursCommessaDisplayName") && app.includes("return `ESTERNA · ${name")],
  ["le commesse esterne non vengono associate alla contabilità interna", app.includes("if (isExternalHoursEntry(entry)) return false") && app.includes('return { id: "", nome: name, codice: "", key: name ? `external:')],
  ["due nomi esterni diversi generano due gruppi export distinti", app.includes('return `external:${normalizeExternalHoursCommessaName') && app.includes("const commessaId = getHoursEntryGroupingKey(entry)")],
  ["script npm configurato", pkg.scripts?.["check:external-hours"] === "node scripts/check-external-hours.js"]
];

checks.forEach(([label, ok]) => console.log(`${ok ? "✓" : "✗"} ${label}`));
if (checks.some(([, ok]) => !ok)) process.exit(1);
