#!/usr/bin/env node
"use strict";

const fs = require("node:fs");

const path = "accounting-v2.js";
let source = fs.readFileSync(path, "utf8");

const before = `    if(!state.work.length){ // compatibilità: i documenti storici restano immutati e sono solo adattati in memoria\n      const legacy=getCommessaCachedImpianti(state.commessa.id)||[];\n      state.plants=legacy.map(p=>({id:p.id,legacy:true,numeroProgressivoImpianto:p.numeroProgressivo,...p,latitudine:p.coordinateLatitudineOriginale??p.gpsY,longitudine:p.coordinateLongitudineOriginale??p.gpsX}));\n      state.work=legacy.flatMap(p=>core.adaptLegacyPlantToWorkItems({...p,numeroProgressivoRiga:p.numeroProgressivo,frequenzaAnnua:p.frequenzaAnnua||\"\",note:p.note||p.noteImpianto||\"\"}));\n    }`;

const after = `    if(!state.work.length){ // compatibilità: nessuna scrittura; gli impianti operativi diventano righe legacy solo in memoria\n      const cachedLegacy=getCommessaCachedImpianti(state.commessa.id)||[];\n      const legacy=state.operationalPlants.length?state.operationalPlants:cachedLegacy;\n      const physicalById=new Map(state.plants.map(p=>[String(p.id||\"\"),p]));\n      state.plants=legacy.map((p,index)=>{\n        const linkedId=String(p.physicalPlantId||p.id||\"\").trim();\n        const physical=physicalById.get(linkedId)||{};\n        const merged={...physical,...p};\n        const id=linkedId||String(physical.id||\`legacy_\${index+1}\`);\n        return {...merged,id,legacy:true,numeroProgressivoImpianto:merged.numeroProgressivoImpianto??merged.numeroProgressivo??index+1,latitudine:merged.coordinateLatitudineOriginale??merged.latitudine??merged.gpsY,longitudine:merged.coordinateLongitudineOriginale??merged.longitudine??merged.gpsX};\n      });\n      state.work=state.plants.flatMap(p=>core.adaptLegacyPlantToWorkItems({...p,numeroProgressivoRiga:p.numeroProgressivoRiga??p.numeroProgressivo??p.numeroProgressivoImpianto,frequenzaAnnua:p.frequenzaAnnua||\"\",note:p.note||p.noteImpianto||\"\"}));\n    }`;

if (source.includes(after)) {
  console.log("Fallback tabella contabilità già applicato.");
  process.exit(0);
}

if (!source.includes(before)) {
  throw new Error("Blocco legacy atteso non trovato in accounting-v2.js");
}

source = source.replace(before, after);
fs.writeFileSync(path, source);
console.log("Fallback tabella contabilità applicato.");
