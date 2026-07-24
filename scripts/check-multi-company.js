const fs = require("node:fs");
const assert = require("node:assert/strict");

const app = fs.readFileSync("app.js", "utf8");
const html = fs.readFileSync("index.html", "utf8");
const rules = fs.readFileSync("firestore.rules", "utf8");

assert.match(app, /const DEFAULT_COMPANY_ID = "avola"/);
assert.match(app, /legacyCompanyMigration: true/);
assert.match(app, /return showCompanyCodeGate\(\)/);
assert.match(app, /TENANT_SCOPED_COLLECTIONS/);
assert.match(app, /companyId !== DEFAULT_COMPANY_ID/);
assert.match(app, /currentCompanyRole === "company_admin"/);
assert.match(html, /id="company-code-gate"/);
assert.match(html, /id="open-panel-aziende"/);
assert.match(html, /id="panel-aziende"/);
assert.match(rules, /function belongsToCompany\(companyId\)/);
assert.match(rules, /match \/companies\/\{companyId\}/);
assert.match(rules, /allow list: if false/);

console.log("✅ Base multi-azienda, migrazione Avola e isolamento tenant presenti.");
