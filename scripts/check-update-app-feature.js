const assert = require("node:assert/strict");
const fs = require("node:fs");

const loader = fs.readFileSync("hours-export-range.js", "utf8");
const feature = fs.readFileSync("update-app-feature.js", "utf8");
const capacitor = JSON.parse(fs.readFileSync("capacitor.config.json", "utf8"));

assert.match(loader, /update-app-feature\.js\?v=20260727a/);
assert.match(feature, /id = "update-app-btn"/);
assert.match(feature, /width:\s*34px;[\s\S]*height:\s*34px;/);
assert.match(feature, /left:\s*40px;/);
assert.match(feature, new RegExp(`play\\.google\\.com/store/apps/details\\?id=${capacitor.appId.replaceAll(".", "\\.")}`));
assert.match(feature, /navigator\.serviceWorker\?\.getRegistration/);
assert.match(feature, /window\.location\.reload\(\)/);

console.log("Update app feature checks passed.");
