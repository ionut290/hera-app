const fs = require('fs');
const path = require('path');

const root = path.resolve(__dirname, '..');
const files = fs.readdirSync(root)
  .filter((name) => /^lavori-occasionali.*\.js$/i.test(name) || name === 'impianti-pdf-storage.js')
  .sort();

// Regola architetturale: i moduli dei Lavori occasionali e il modulo PDF impianti
// possono LEGGERE le cache globali degli impianti, ma non possono mai riscriverle
// o modificarle. La visibilità degli impianti resta di app.js.
const forbidden = [
  { re: /\bcurrentImpianti\s*=/g, reason: 'non deve sostituire currentImpianti' },
  { re: /\bcurrentImpianti\.(?:push|splice|pop|shift|unshift|sort|reverse)\s*\(/g, reason: 'non deve mutare currentImpianti' },
  { re: /\bimpiantiByCommessaId\.(?:set|delete|clear)\s*\(/g, reason: 'non deve mutare impiantiByCommessaId' },
];

const findings = [];
for (const file of files) {
  const source = fs.readFileSync(path.join(root, file), 'utf8');
  for (const rule of forbidden) {
    rule.re.lastIndex = 0;
    let match;
    while ((match = rule.re.exec(source))) {
      const line = source.slice(0, match.index).split('\n').length;
      findings.push(`${file}:${line} — ${rule.reason}`);
    }
  }
}

if (findings.length) {
  console.error('\n❌ BLOCCO SICUREZZA CANTIERI');
  console.error('Un modulo accessorio sta tentando di alterare le cache globali degli impianti.');
  console.error('Questo può far sparire o mescolare i cantieri. La PR deve essere corretta prima del merge.\n');
  findings.forEach((item) => console.error(`- ${item}`));
  process.exit(1);
}

console.log(`✅ Isolamento cantieri OK: controllati ${files.length} moduli protetti.`);
