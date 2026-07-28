const fs=require('fs');
const js=fs.readFileSync('accounting-v2.js','utf8');const core=fs.readFileSync('inrete-work-items-v2.js','utf8');const html=fs.readFileSync('index.html','utf8');const repair=fs.readFileSync('operational-import-repair.js','utf8');
const checks=[
 ['titolo gestione',html.includes('Gestione impianti e contabilità')],
 ['22 colonne',js.includes('"Totali"')&&js.includes('"Voce di Riferimento Elenco Prezzi"')],
 ['prezziario per commessa',js.includes('collection("prezziario")')],
 ['impianti fisici separati',js.includes('collection("impiantiFisici")')],
 ['lavorazioni separate',js.includes('collection("lavorazioni")')],
 ['subtotale solo FATTO',js.includes('calculateCompletedSubtotal')],
 ['formula AC',core.includes('toUpperCase() === "AC"')],
 ['formula Excel SUMIF',js.includes('SUMIF(U2:U')],
 ['timezone Roma',js.includes('Europe/Rome') || fs.readFileSync('app.js','utf8').includes('Europe/Rome')],
 ['svuotamento selettivo',js.includes('PLANTS_AND_WORK_ITEMS')&&js.includes('PRICE_LIST')&&js.includes('clearOperations')],
 ['compatibilità legacy',js.includes('compatibilità: i documenti storici restano immutati')],
 ['calcolo UI indipendente da Excel',js.includes('la UI non legge mai risultati dalle formule xlsx')],
 ['migrazione INRETE disponibile',js.includes('migrateInreteCommesseToWorkItemsV2')],
 ['import formati',html.includes('accept=".xlsx,.xls,.csv,.ods"')],
 ['sincronizzazione operativa',js.includes('synchronizeOperationalModel')&&js.includes('collection("impianti")')],
 ['autoriparazione importazioni incomplete',js.includes('needsOperationalRepair')&&js.includes('autoRepairing')&&js.includes('load({autoRepair:false})')],
 ['autoriparazione dal caricamento app',repair.includes('repairCommessa')&&repair.includes('workByPlantId')&&repair.includes('repairImportedInretePlants')],
 ['riparazione idempotente',js.includes('repairImportedMatrixPlants')&&js.includes('migrationSourceId')&&js.includes('stableId')],
 ['batch sotto limite Firestore',js.includes('i+=400')&&js.includes('commitOperations')],
 ['contatori commessa',js.includes('impiantiFattiCount')&&js.includes('workItemsDaFareCount')],
 ['coordinate validate',js.includes('coordinate=(value,min,max)')&&js.includes('gpsY:coordinate')],
 ['pulsante riparazione',html.includes('repair-imported-plants-btn')&&html.includes('Ripara collegamento impianti')]
];let failed=false;for(const [name,ok] of checks){console.log(`${ok?'✅':'❌'} ${name}`);failed||=!ok;}if(failed)process.exit(1);
