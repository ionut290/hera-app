// Projection only. It never changes the operational state, execution dates or source plants.
export function projectPreviewWork(plants,prices,generalDiscount,quantities,core){
 const counts=new Map();
 for(const item of prices){
  const code=core.normalizePriceCode(item.codiceVoce);
  if(code)counts.set(code,(counts.get(code)||0)+1);
 }
 const priceMap=core.buildPriceMap(prices.filter(p=>{
  const code=core.normalizePriceCode(p.codiceVoce);
  return code&&counts.get(code)===1&&typeof p.prezzoBase==='number'&&Number.isFinite(p.prezzoBase)&&p.prezzoBase>=0&&String(p.unitaMisura||'').trim()&&(p.percentualeRibasso==null||typeof p.percentualeRibasso==='number'&&p.percentualeRibasso>=0&&p.percentualeRibasso<=1);
 }));
 const rows=[];
 for(const plant of plants){
  const seen=new Set();
  for(const item of core.adaptLegacyPlantToWorkItems(plant)){
   const code=core.normalizePriceCode(item.codiceVocePrezzo);
   if(seen.has(code))continue;
   seen.add(code);
   const id=JSON.stringify([plant.id,code]),raw=quantities.has(id)?quantities.get(id):item.quantita;
   const parsed=raw==null||String(raw).trim()===''?null:core.parseNumber(raw);
   const quantita=Number.isFinite(parsed)&&parsed>=0&&parsed<=100000000?parsed:null;
   const enriched=core.enrichWorkItem({...item,quantita},priceMap,generalDiscount);
   const total=Number.isFinite(enriched.totale)&&enriched.totale>=0?enriched.totale:null;
   rows.push({...enriched,id,quantita,totale:total,economicStatus:total===null?'DA_VERIFICARE':'OK'});
  }
 }
 return rows;
}
export function summarizePreviewWork(rows){
 return {rows:rows.length,verified:rows.filter(r=>r.totale!==null).length,missing:rows.filter(r=>r.totale===null).length,estimated:rows.reduce((sum,r)=>sum+(r.totale??0),0)};
}
