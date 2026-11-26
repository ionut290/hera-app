'use strict';
// On-demand, tenant-only read projections. Never use global registries or legacy paths.
const {SUPPORT,readSupport}=require('./support-modules.cjs');
const REGISTRIES=Object.freeze({
  personnel:['personale','neve_personale'],
  vehicles:['mezzi','neve_mezzi'],
  teams:['squadreCommesse','neve_squadre']
});
const ECONOMICS=Object.freeze({accounting:'lavorazioni',prices:'prezziario'});
function fail(code){throw Object.assign(new Error(code),{code});}
function own(obj,key){return Object.prototype.hasOwnProperty.call(obj,key);}
function str(x){return typeof x==='string'?x.slice(0,250):'';}
function number(x){return typeof x==='number'&&Number.isFinite(x)?x:null;}
function row(module,id,data){
  if(module==='personnel')return {id,name:str(data.fullName)||[str(data.cognome),str(data.nome)].filter(Boolean).join(' ')||id};
  if(module==='vehicles')return {id,name:[str(data.nId)||str(data.nome),str(data.marca),str(data.modello)].filter(Boolean).join(' · ')||id};
  if(module==='teams')return {id,name:str(data.commessaNome)||str(data.commessaId)||id,date:str(data.riferimentoData)||str(data.dateKey),count:Array.isArray(data.squadre)?data.squadre.length:0};
  if(module==='prices')return {id,name:str(data.descrizione)||str(data.codiceVoce)||id,code:str(data.codiceVoce),unit:str(data.unitaMisura),base:number(data.prezzoBase),discounted:number(data.prezzoRibassato)};
  return {id,name:str(data.denominazione)||str(data.tipologiaLavorazione)||id,code:str(data.codiceVocePrezzo),status:str(data.stato),total:number(data.totale)};
}
async function readModule({tx,access,caller,payload,page,id}){
  const module=payload.module,kind=payload.kind;
  if(SUPPORT.has(module))return readSupport({tx,access,caller,payload,page,id});
  if(!['ordinary','snow'].includes(kind))fail('invalid-kind');
  if(typeof module!=='string'||(!own(REGISTRIES,module)&&!own(ECONOMICS,module)))fail('invalid-module');
  const {limit,after}=page(payload);
  let collection,commessaId=null;
  if(own(ECONOMICS,module)){
    commessaId=id(payload.commessaId);
    const parent=access.root.collection(kind==='snow'?'neve_commesse':'commesse').doc(commessaId);
    if(!(await tx.get(parent)).exists)fail('not-found');
    collection=parent.collection(ECONOMICS[module]);
  }else{
    collection=access.root.collection(REGISTRIES[module][kind==='snow'?1:0]);
  }
  // Existing aggregate has priority for ordinary personnel/vehicles. Its contents
  // are projected, never returned as raw documents. No polling or startup reads.
  if(kind==='ordinary'&&(module==='personnel'||module==='vehicles')){
    const view=await tx.get(access.root.collection('sharedStaticViews').doc('registri__corrente'));
    const records=view.exists?view.data()?.payload?.[REGISTRIES[module][0]]:null;
    if(Array.isArray(records)){
      const sorted=records.filter(x=>x&&typeof x.id==='string'&&/^[A-Za-z0-9_-]{1,128}$/.test(x.id)&&!/^__.*__$/.test(x.id)).sort((a,b)=>a.id<b.id?-1:a.id>b.id?1:0);
      const filtered=after?sorted.filter(x=>x.id>after):sorted;
      const rows=filtered.slice(0,limit);
      return {company:access.company,module,kind,commessaId,rows:rows.map(x=>({...row(module,x.id,x),...(x.tenantLabManaged===true?{editable:true}:{})})),nextCursor:filtered.length>limit?rows.at(-1).id:null};
    }
  }
  let query=collection.orderBy('__name__').limit(limit+1);if(after)query=query.startAfter(after);
  const snap=await tx.get(query),rows=snap.docs.slice(0,limit);
  return {company:access.company,module,kind,commessaId,rows:rows.map(x=>({...row(module,x.id,x.data()),...(x.data().tenantLabManaged===true?{editable:true}:{})})),nextCursor:snap.size>limit?rows.at(-1).id:null};
}
module.exports={readModule};
