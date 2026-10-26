const assert = require("node:assert");
const fs = require("node:fs");

const bootstrap = fs.readFileSync("oracle-canary-bootstrap.js", "utf8");
const nginx = fs.readFileSync("oracle-original-canary/nginx.conf", "utf8");

assert.ok(!/innerHTML|appendChild|insertAdjacentHTML|classList|\.style\b/.test(bootstrap), "Il canary non deve modificare la UI.");
assert.ok(!/\.collection\s*\(|\.onSnapshot\s*\(/.test(bootstrap), "Il canary non deve aggiungere accessi Firestore.");
assert.ok(!/\b(?:set|add|update|delete)Doc\s*\(/.test(bootstrap), "Il canary deve essere in sola lettura.");
assert.ok(bootstrap.includes("onlyFirebase") && bootstrap.includes("onlyOracle"), "Manca il confronto bidirezionale degli ID.");
assert.ok(bootstrap.includes("visibleCommesseCovered: onlyFirebase.length === 0"), "Il canary deve verificare che Oracle copra tutte le commesse visibili all'utente.");
assert.ok(bootstrap.includes("exactCatalogMatch"), "Manca la diagnostica separata di uguaglianza del catalogo completo.");
assert.ok(nginx.includes("sub_filter '</head>'") && nginx.includes("sub_filter '</body>'"), "Gli script canary devono essere iniettati senza modificare index.html.");
assert.ok(nginx.includes("deny all"), "Mancano i blocchi per file e directory non pubblici.");
console.log("Original UI Oracle canary: nessuna modifica UI, nessuna scrittura/listener Firestore.");
