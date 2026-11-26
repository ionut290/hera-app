// Local demo only: plan all rows before mutating a single contract.
export function planMatrixImport(input,existing,{mapped,norm,num,normalizeCode}){
 if(!Array.isArray(input)||input.length>500)throw Error('Massimo 500 lavorazioni.');
 const text=(row,f)=>String(mapped(row,f)??'').trim();
 const identity=(p)=>p.idSap?'sap:'+norm(p.idSap):'site:'+JSON.stringify([norm(p.denominazione),norm(p.comune),norm(p.indirizzo)]);
 const location=p=>JSON.stringify([norm(p.denominazione),norm(p.comune),norm(p.indirizzo)]);
 const used=new Set(existing.map(identity)),usedLocation=new Set(existing.map(location)),groups=new Map(),locationIds=new Map();
 let workCount=0;
 for(const [index,row] of input.entries()){
  if(Object.values(row).some(v=>norm(v).includes('esempio - non importare')))continue;
  const p={denominazione:text(row,'denominazione'),comune:text(row,'comune'),indirizzo:text(row,'indirizzo'),idSap:text(row,'idSap'),distretto:text(row,'distretto')};
  const code=normalizeCode(text(row,'codiceVocePrezzo')),raw=text(row,'quantita'),q=raw?num(raw):null;
  const status=text(row,'stato').toUpperCase(),freq=text(row,'frequenzaAnnua'),description=text(row,'tipologiaLavorazione');
  if(!p.denominazione||p.denominazione.length>250||p.comune.length>250||p.indirizzo.length>500||p.idSap.length>150||p.distretto.length>150||!code||code.length>150||/[|;,\r\n]/.test(code)||raw&&(!Number.isFinite(q)||q<0||q>100000000)||freq.length>100||description.length>500)throw Error('Riga '+(index+2)+': dati o quantità non validi.');
  if(status&&status!=='DA FARE')throw Error('Riga '+(index+2)+': importare soltanto nuove lavorazioni DA FARE.');
  const id=identity(p),loc=location(p);
  if(used.has(id)||usedLocation.has(loc))throw Error('Riga '+(index+2)+': impianto già presente. Nessun dato modificato.');
  if(locationIds.has(loc)&&locationIds.get(loc)!==id)throw Error('Riga '+(index+2)+': identificativi impianto discordanti.');
  locationIds.set(loc,id);
  const group=groups.get(id)||{plant:p,work:[]};
  if(location(group.plant)!==loc||norm(group.plant.distretto)!==norm(p.distretto))throw Error('Riga '+(index+2)+': dati impianto discordanti.');
  if(group.work.some(w=>w.code===code))throw Error('Riga '+(index+2)+': codice ripetuto nello stesso impianto.');
  group.work.push({code,quantity:q,description,frequency:freq});groups.set(id,group);workCount++;
 }
 if(!workCount)throw Error('Nessuna lavorazione importabile.');
 return {groups:Array.from(groups.values()),workCount};
}
