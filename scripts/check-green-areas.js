const fs = require("fs");

const html = fs.readFileSync("index.html", "utf8");
const client = fs.readFileSync("green-areas.js", "utf8");
const serviceWorker = fs.readFileSync("sw.js", "utf8");
const assert = (condition, message) => {
  if (!condition) throw new Error(message);
};

assert(html.includes('id="green-areas-name"') && !html.match(/id="green-areas-name"[^>]*required/), "Il nome deve essere facoltativo.");
assert(html.includes('<select id="green-areas-municipality"'), "Selettore Comune assente.");
for (const municipality of ["Castel Maggiore", "Bologna", "Argelato", "Minerbio", "Zola Predosa"]) {
  assert(html.includes(`value="${municipality}"`), `Comune non disponibile: ${municipality}.`);
}
assert(client.includes("searchMunicipalGreenAreas"), "Ricerca comunale completa assente.");
assert(client.includes("Area verde senza nome"), "Le aree senza nome non vengono incluse.");
assert(client.includes("showAllAreas"), "La mappa non mostra insieme tutte le aree comunali.");
assert(client.includes("APRI SCHEDA") && client.includes("openAreaSheet"), "Scheda area verde non collegata alla mappa.");
assert(client.includes("ALLOWED_MUNICIPALITIES"), "Lista chiusa dei cinque Comuni assente.");
assert(!client.includes("MAX_DISTANCE_KM") && !client.includes("verifyMunicipalityRadius"), "La vecchia ricerca nel raggio di 50 km è ancora presente.");
assert(client.includes("village_green") && client.includes("dog_park"), "Categorie verdi comunali incomplete.");
assert(client.includes('"ISO3166-2"="IT-45"'), "Ricerca non limitata all’Emilia-Romagna.");
assert(client.includes("PSR_Area_verde"), "Livello ufficiale regionale assente.");
assert(!client.includes("firebase") && !client.includes("firestore"), "La pagina aree verdi non deve usare Firestore.");
assert(serviceWorker.includes("./green-area-sheet.css?v=20260829a"), "Stile scheda area verde assente dalla cache PWA.");
assert(serviceWorker.includes("./green-areas.js?v=20260829e"), "Cache PWA della ricerca aree verdi non aggiornata.");

console.log("Controlli aree verdi comunali superati.");
