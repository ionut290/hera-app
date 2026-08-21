const fs = require('fs');
const vm = require('vm');

const source = fs.readFileSync('sync-badge-pending-fix.js', 'utf8');
const loader = fs.readFileSync('update-app-feature.js', 'utf8');

if (!source.includes('heraPendingOfflineMutations')) throw new Error('Manca coda offline reale nel contatore');
if (!source.includes('heraPendingImpiantoActions')) throw new Error('Manca coda FATTO reale nel contatore');
if (!source.includes('heraPendingSheetExports')) throw new Error('Manca coda export reale nel contatore');
if (!source.includes('syncPendingImpiantoActions')) throw new Error('La sync manuale non richiama la coda impianti');
if (!source.includes('processPendingSheetExports')) throw new Error('La sync manuale non richiama la coda export');
if (!loader.includes('sync-badge-pending-fix.js')) throw new Error('La correzione badge non viene caricata');

new vm.Script(source, { filename: 'sync-badge-pending-fix.js' });
new vm.Script(loader, { filename: 'update-app-feature.js' });
console.log('Sync badge pending count guard OK');
