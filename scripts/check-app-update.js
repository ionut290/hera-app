const assert = require("node:assert/strict");
const fs = require("node:fs");

const html = fs.readFileSync("index.html", "utf8");
const css = fs.readFileSync("style.css", "utf8");
const app = fs.readFileSync("app.js", "utf8");
const capacitor = JSON.parse(fs.readFileSync("capacitor.config.json", "utf8"));

assert.match(html, /id="update-app-btn"[^>]*>[\s\S]*?Aggiorna app[\s\S]*?<\/button>/);
assert.match(css, /\.update-app-btn\s*\{[\s\S]*?height:\s*34px;/);
assert.match(css, /\.logo-head\s*\{[\s\S]*?grid-template-columns:\s*minmax\(0, 1fr\) auto minmax\(0, 1fr\)/);
assert.match(app, /updateAppBtn\?\.addEventListener\("click", openApplicationUpdate\)/);
assert.match(app, new RegExp(`play\\.google\\.com/store/apps/details\\?id=${capacitor.appId.replaceAll(".", "\\.")}`));
assert.match(app, /https:\/\/creative-syrniki-dddbae\.netlify\.app\//);
assert.match(app, /Capacitor\?\.getPlatform\?\.\(\) === "android"/);
assert.match(app, /updateUrl\.searchParams\.set\("update", String\(Date\.now\(\)\)\)/);

console.log("App update button checks passed.");
