const fs = require("fs");
const path = require("path");

const root = path.resolve(__dirname, "..");

function replaceOnce(source, search, replacement, label) {
  const first = source.indexOf(search);
  if (first < 0) throw new Error(`Blocco non trovato: ${label}`);
  if (source.indexOf(search, first + search.length) >= 0) {
    throw new Error(`Blocco duplicato, patch interrotta: ${label}`);
  }
  return source.slice(0, first) + replacement + source.slice(first + search.length);
}

function patchOccasionalCore() {
  const file = path.join(root, "lavori-occasionali.js");
  let source = fs.readFileSync(file, "utf8");

  source = replaceOnce(
    source,
    `  function wrapRowReader() {\n    if (state.rowsWrapped || typeof readSquadraRows !== "function") return;`,
    `  function wrapRowReader() {\n    if (window.HeraOccasionalSquadFirstFlow?.installed) {\n      state.rowsWrapped = true;\n      return;\n    }\n    if (state.rowsWrapped || typeof readSquadraRows !== "function") return;`,
    "wrapRowReader"
  );

  source = replaceOnce(
    source,
    `  function restoreWorkNameFromComposition() {\n    if (!isOccasionalSelected()) return;`,
    `  function restoreWorkNameFromComposition() {\n    if (window.HeraOccasionalSquadFirstFlow?.installed) {\n      if (!isOccasionalSelected()) return;\n      const dateKey = document.getElementById("squadra-riferimento")?.value || "";\n      try {\n        const composition = squadreHistoryByDate instanceof Map\n          ? squadreHistoryByDate.get(dateKey)?.get(COMMESSA_ID)\n          : null;\n        if (composition && typeof setSquadraRowsFromData === "function") {\n          window.setTimeout(() => {\n            try {\n              setSquadraRowsFromData(composition);\n              if (typeof updateSquadraAutofillHint === "function") {\n                updateSquadraAutofillHint("Composizione salvata per LAVORI OCCASIONALI.");\n              }\n            } catch (error) {\n              state.lastError = error;\n            }\n          }, 0);\n        }\n      } catch (error) {\n        state.lastError = error;\n      }\n      return;\n    }\n    if (!isOccasionalSelected()) return;`,
    "restoreWorkNameFromComposition"
  );

  source = replaceOnce(
    source,
    `  function validateBeforeCoreSave(event) {\n    if (!isOccasionalSelected()) return;`,
    `  function validateBeforeCoreSave(event) {\n    if (window.HeraOccasionalSquadFirstFlow?.installed) return;\n    if (!isOccasionalSelected()) return;`,
    "validateBeforeCoreSave"
  );

  source = replaceOnce(
    source,
    `  function applyWorkNamesToData() {\n    const items = getCompositions();`,
    `  function applyWorkNamesToData() {\n    if (window.HeraOccasionalSquadFirstFlow?.installed) return;\n    const items = getCompositions();`,
    "applyWorkNamesToData"
  );

  source = replaceOnce(
    source,
    `  function decorateSquadCards() {\n    applyWorkNamesToData();`,
    `  function decorateSquadCards() {\n    if (window.HeraOccasionalSquadFirstFlow?.installed) return;\n    applyWorkNamesToData();`,
    "decorateSquadCards"
  );

  source = replaceOnce(
    source,
    `  document.getElementById("squadra-form")?.addEventListener("submit", () => {\n    const feedback = document.getElementById("squadra-feedback");`,
    `  document.getElementById("squadra-form")?.addEventListener("submit", () => {\n    if (window.HeraOccasionalSquadFirstFlow?.installed) return;\n    const feedback = document.getElementById("squadra-feedback");`,
    "salvataggio automatico impianto durante submit squadra"
  );

  source = replaceOnce(
    source,
    `    version: "1.3.0",\n    commessaId: COMMESSA_ID,\n    firestoreScope: "commesse/lavori-occasionali/impianti",\n    refresh,`,
    `    version: "1.4.0",\n    commessaId: COMMESSA_ID,\n    firestoreScope: "commesse/lavori-occasionali/impianti",\n    getWorkMetadata,\n    upsertPlant: upsertNormalOccasionalPlant,\n    refresh,`,
    "API pubblica lavori occasionali"
  );

  fs.writeFileSync(file, source);
}

function patchFirebaseLoader() {
  const file = path.join(root, "firebase-config.js");
  let source = fs.readFileSync(file, "utf8");

  source = replaceOnce(
    source,
    `const HERA_OCCASIONAL_GOOGLE_PLACES_SRC = "lavori-occasionali-google-places.js?v=20260823-map2";`,
    `const HERA_OCCASIONAL_SQUAD_FIRST_SRC = "lavori-occasionali-squad-first.js?v=20260824a";\nconst HERA_OCCASIONAL_GOOGLE_PLACES_SRC = "lavori-occasionali-google-places.js?v=20260823-map2";`,
    "costante loader squadra prima"
  );

  source = replaceOnce(
    source,
    `  document.write(\`<script src="\${HERA_OCCASIONAL_GOOGLE_PLACES_SRC}" data-occasional-google-places="1"><\\/script>\`);`,
    `  document.write(\`<script src="\${HERA_OCCASIONAL_SQUAD_FIRST_SRC}" data-occasional-squad-first="1"><\\/script>\`);\n  document.write(\`<script src="\${HERA_OCCASIONAL_GOOGLE_PLACES_SRC}" data-occasional-google-places="1"><\\/script>\`);`,
    "document.write loader squadra prima"
  );

  source = replaceOnce(
    source,
    `  loadOnce(HERA_OCCASIONAL_GOOGLE_PLACES_SRC, "occasional-google-places", () => Boolean(window.HeraLavoriOccasionaliGooglePlaces?.installed));`,
    `  loadOnce(HERA_OCCASIONAL_SQUAD_FIRST_SRC, "occasional-squad-first", () => Boolean(window.HeraOccasionalSquadFirstFlow?.installed));\n  loadOnce(HERA_OCCASIONAL_GOOGLE_PLACES_SRC, "occasional-google-places", () => Boolean(window.HeraLavoriOccasionaliGooglePlaces?.installed));`,
    "loadOnce loader squadra prima"
  );

  fs.writeFileSync(file, source);
}

patchOccasionalCore();
patchFirebaseLoader();
console.log("✅ Patch applicata: prima si salva la squadra, poi si aggiungono i cantieri.");
