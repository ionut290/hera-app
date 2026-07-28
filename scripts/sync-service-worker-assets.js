const fs = require("node:fs");

const indexPath = "index.html";
const serviceWorkerPath = "sw.js";
const assetNames = ["app.js"];

const index = fs.readFileSync(indexPath, "utf8");
let serviceWorker = fs.readFileSync(serviceWorkerPath, "utf8");
let changed = false;

for (const assetName of assetNames) {
  const escapedName = assetName.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
  const indexMatch = index.match(new RegExp(`["'](${escapedName}\\?v=[^"']+)["']`));

  if (!indexMatch) {
    throw new Error(`${indexPath} non contiene una versione cache-busting per ${assetName}`);
  }

  const serviceWorkerPattern = new RegExp(`\\./${escapedName}(?:\\?v=[^"']+)?`, "g");
  const expectedAsset = `./${indexMatch[1]}`;

  if (!serviceWorkerPattern.test(serviceWorker)) {
    throw new Error(`${serviceWorkerPath} non precarica ${assetName}`);
  }

  serviceWorkerPattern.lastIndex = 0;
  serviceWorker = serviceWorker.replace(serviceWorkerPattern, (currentAsset) => {
    if (currentAsset === expectedAsset) return currentAsset;
    changed = true;
    return expectedAsset;
  });
}

if (changed) {
  serviceWorker = serviceWorker.replace(/hera-app-shell-v(\d+)/, (_, version) => (
    `hera-app-shell-v${Number(version) + 1}`
  ));
  fs.writeFileSync(serviceWorkerPath, serviceWorker);
  console.log("Service Worker sincronizzato con le versioni degli asset in index.html.");
} else {
  console.log("Le versioni degli asset del Service Worker sono già sincronizzate.");
}
