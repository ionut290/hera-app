"use strict";

const fs = require("node:fs");
const path = require("node:path");
const root = path.resolve(__dirname, "..");
const read = (file) => fs.readFileSync(path.join(root, file), "utf8");
const html = read("index.html");
const capture = read("private-documents-v2.js");
const viewer = read("identity-card-feature.js");
const css = `${read("private-documents-v2.css")}\n${read("identity-card-feature.css")}`;
const checks = [
  ["acquisizione posteriore", html.includes('capture="environment"') && capture.includes("facingMode") && capture.includes("getUserMedia")],
  ["rotazione EXIF", capture.includes("readExifOrientation") && capture.includes("orientedCanvas")],
  ["rilevamento e fallback manuale", capture.includes("detectDocumentEdges") && capture.includes("Rilevamento incerto") && html.includes("identity-crop-handles")],
  ["correzione prospettica", capture.includes("homography") && capture.includes("perspectiveCorrect")],
  ["ritaglio tessera isolato", capture.includes("if (identityPreset)") && capture.includes("IDENTITY_RATIO")],
  ["anteprima e conferma", html.includes("identity-crop-preview") && html.includes("identity-crop-use") && html.includes("identity-crop-retry")],
  ["fullscreen zoom e trascinamento", viewer.includes("openFullscreenImage") && viewer.includes("pointermove") && viewer.includes("popstate")],
  ["viewport e contenimento", css.includes("100dvh") && css.includes("object-fit: contain") && css.includes("safe-area-inset")]
];
const failed = checks.filter(([, ok]) => !ok);
checks.forEach(([label, ok]) => console.log(`${ok ? "✓" : "✗"} ${label}`));
if (failed.length) throw new Error(`Controlli tesserino falliti: ${failed.map(([label]) => label).join(", ")}`);
