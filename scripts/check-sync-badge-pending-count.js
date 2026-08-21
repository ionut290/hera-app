const fs = require('fs');
const vm = require('vm');

const source = fs.readFileSync('data-durability-runtime.js', 'utf8');

if (!source.includes('heraPendingOfflineMutations')) throw new Error('Manca coda offline reale nel contatore');
if (!source.includes('heraPendingImpiantoActions')) throw new Error('Manca coda FATTO reale nel contatore');
if (!source.includes('heraPendingSheetExports')) throw new Error('Manca coda export reale nel contatore');
if (!source.includes('syncPendingImpiantoActions')) throw new Error('La sync manuale non richiama la coda impianti');
if (!source.includes('processPendingSheetExports')) throw new Error('La sync manuale non richiama la coda export');

new vm.Script(source, { filename: 'data-durability-runtime.js' });
console.log('Sync badge pending count guard OK');
