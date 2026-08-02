const fs = require('fs');
const assert = require('assert');

const files = [
  'multi-organization-context.js',
  'organization-access-model.js',
  'organization-selector.js',
  'super-admin-organizations-model.js',
  'multi-organization-migration-plan.js',
  'multi-organization-admin.js'
];

for (const file of files) {
  assert(fs.existsSync(file), `File mancante: ${file}`);
  const content = fs.readFileSync(file, 'utf8');
  assert(!/onSnapshot\s*\(/.test(content), `${file}: listener Firestore realtime non consentito`);
  assert(!/setInterval\s*\(/.test(content), `${file}: polling non consentito`);
}

const admin = fs.readFileSync('multi-organization-admin.js', 'utf8');
assert(admin.includes('buildMigrationDryRun'), 'Dry-run migrazione mancante');
assert(admin.includes('window.confirm'), 'Conferma esplicita migrazione mancante');
assert(admin.includes("DEFAULT_ORGANIZATION_ID = 'varga'"), 'Fallback Varga mancante');
assert(!admin.includes('fatto-button-immediate'), 'Riferimento vietato alla logica FATTO');
assert(!/whats?app|whazzup/i.test(admin), 'Riferimento vietato alla logica WhatsApp/WHAZZUP');
console.log('Multi-organizzazione: controlli statici superati.');
