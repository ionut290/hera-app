const fs = require("fs");

const html = fs.readFileSync("index.html", "utf8");
const client = fs.readFileSync("green-areas.js", "utf8");
const serviceWorker = fs.readFileSync("sw.js", "utf8");
const assert = (condition, message) => {
  if (!condition) throw new Error(message);
};

assert(html.includes('id="green-areas-name"') && !html.match(/id="green-areas-name"[^>]*required/), "Il nome deve essere facoltativo.");
assert(html.includes('id="green-areas-municipality"'), "Campo Comune assente.");
assert(client.includes("searchMunicipalGreenAreas"), "Ricerca comunale completa assente.");
assert(client.includes("Area verde senza nome"), "Le aree senza nome non vengono incluse.");
assert(client.includes('"ISO3166-2"="IT-45"'), "Ricerca non limitata all’Emilia-Romagna.");
assert(client.includes("PSR_Area_verde"), "Livello ufficiale regionale assente.");
assert(!client.includes("firebase") && !client.includes("firestore"), "La pagina aree verdi non deve usare Firestore.");
assert(serviceWorker.includes("./green-areas.js?v=20260829b"), "Cache PWA della ricerca aree verdi non aggiornata.");

console.log("Controlli aree verdi comunali superati.");
