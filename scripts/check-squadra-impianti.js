const fs = require("fs");
const path = require("path");

const appSource = fs.readFileSync(path.join(__dirname, "..", "app.js"), "utf8");
const cssSource = fs.readFileSync(path.join(__dirname, "..", "style.css"), "utf8");

const checks = [
  ["ricerca impianti limitata alla commessa", /getCommessaCachedImpianti\(ui\.squadraCommessa\?\.value\)/],
  ["ricerca per prefisso senza distinzione maiuscole/minuscole", /toLocaleLowerCase\("it-IT"\)\.startsWith\(term\)/],
  ["risultati con Comune e ID SAP", /impianto\.comune[\s\S]*ID SAP[\s\S]*impianto\.idSap/],
  ["blocco duplicati nella stessa squadra", /selectedIds\.has\(getSquadraImpiantoId\(impianto\)\)/],
  ["salvataggio snapshot e ID univoco", /snapshot\.impiantoId = getSquadraImpiantoId\(impianto\)/],
  ["apertura della commessa sull'impianto selezionato", /openImpiantiPage\(`&impianto=\$\{encodeURIComponent\(impiantoKey\)\}`\)/],
  ["rimozione visuale degli impianti FATTO", /return !live\?\.done/],
  ["chip compatti nella vista squadre", /\.squadra-impianto-link[\s\S]*max-width:\s*150px/]
];

let failed = false;
checks.forEach(([label, pattern]) => {
  const source = label.includes("chip") ? cssSource : appSource;
  if (pattern.test(source)) {
    console.log(`✅ ${label}`);
  } else {
    failed = true;
    console.error(`❌ ${label}`);
  }
});

if (failed) process.exit(1);
console.log("✅ Controlli impianti assegnati completati senza conflitti.");
