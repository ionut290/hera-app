const fs = require('fs');
const s = fs.readFileSync('app.js', 'utf8');
if (!s.includes('function yieldToBrowserDuringStartup()')) process.exit(1);
let ready = s.indexOf('console.log("APP READY")');
if (ready < 0) ready = s.indexOf("console.log('APP READY')");
if (ready < 0) process.exit(1);
const auth = s.lastIndexOf('onAuthStateChanged', ready);
if (auth < 0) process.exit(1);
const block = s.slice(auth, ready);
if ((block.match(/await yieldToBrowserDuringStartup\(\);/g) || []).length !== 3) process.exit(1);
for (const token of ['operatorPositions = [];', 'closeCommessaResourceViewer();', 'clearMap();']) {
  if (!block.includes(token)) process.exit(1);
}
console.log('Auth startup scoped yields OK');
