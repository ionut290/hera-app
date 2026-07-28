const fs=require('fs');
const js=fs.readFileSync('accounting-v2.js','utf8');const html=fs.readFileSync('index.html','utf8');
const checks=[
 ['titolo gestione',html.includes('Gestione impianti e contabilità')],
 ['22 colonne',js.includes('"Totali"')&&js.includes('"Voce di Riferimento Elenco Prezzi"')],
 ['prezziario per commessa',js.includes('collection("prezziario")')],
 ['impianti fisici separati',js.includes('collection("impiantiFisici")')],
 ['lavorazioni separate',js.includes('collection("lavorazioni")')],
 ['subtotale solo FATTO',js.includes('filter(w=>w.stato==="FATTO")')],
 ['formula AC',js.includes('toUpperCase()==="AC"')],
 ['formula Excel SUMIF',js.includes('SUMIF(U2:U')],
 ['timezone Roma',js.includes('Europe/Rome') || fs.readFileSync('app.js','utf8').includes('Europe/Rome')],
 ['svuotamento selettivo',js.includes('PLANTS_AND_WORK_ITEMS')&&js.includes('PRICE_LIST')&&js.includes('clearOperations')],
 ['compatibilità legacy',js.includes('compatibilità: i documenti storici restano immutati')],
 ['import formati',html.includes('accept=".xlsx,.xls,.csv,.ods"')]
];let failed=false;for(const [name,ok] of checks){console.log(`${ok?'✅':'❌'} ${name}`);failed||=!ok;}if(failed)process.exit(1);
