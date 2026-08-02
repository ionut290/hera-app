const fs = require('fs');
const path = require('path');

const root = path.resolve(__dirname, '..');
const rulesPath = path.join(root, 'firestore.rules');
const fragmentPath = path.join(root, 'firestore.multi-organization.fragment.rules');
const outputPath = path.join(root, 'firestore.multi-organization.generated.rules');
const apply = process.argv.includes('--apply');

const rules = fs.readFileSync(rulesPath, 'utf8');
const fragment = fs.readFileSync(fragmentPath, 'utf8').trim();
const startMarker = '// HERA_MULTI_ORGANIZATION_RULES_START';
const endMarker = '// HERA_MULTI_ORGANIZATION_RULES_END';

if (rules.includes(startMarker) || rules.includes(endMarker)) {
  throw new Error('La patch multi-organizzazione risulta già presente in firestore.rules.');
}

const genericMatch = '    match /{collection}/{docId} {';
const insertAt = rules.indexOf(genericMatch);
if (insertAt < 0) {
  throw new Error('Match generico finale non trovato: nessuna modifica eseguita.');
}

const indentedFragment = fragment
  .split('\n')
  .map((line) => line ? `    ${line}` : '')
  .join('\n');
const generated = `${rules.slice(0, insertAt)}${indentedFragment}\n\n${rules.slice(insertAt)}`;

fs.writeFileSync(outputPath, generated, 'utf8');
console.log(`Anteprima generata: ${path.relative(root, outputPath)}`);
console.log('Nessuna regola è stata distribuita su Firebase.');

if (apply) {
  fs.writeFileSync(rulesPath, generated, 'utf8');
  console.log('firestore.rules aggiornato localmente. Eseguire i test prima del deploy.');
} else {
  console.log('Per applicare al file locale: node scripts/apply-multi-organization-rules-fragment.js --apply');
}
