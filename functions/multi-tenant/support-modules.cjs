'use strict';
// Tenant-scoped support read service. No Storage URLs, recipient details or external calls.
const SUPPORT=new Set(['documents','calendar','events','hours','approvals','notifications','integrations']);
const COMMON=new Set(['documents','calendar','events','notifications','integrations']);
function fail(code){throw Object.assign(new Error(code),{code});}
function str(x){return typeof x==='string'?x.slice(0,250):'';}
function date(x){return typeof x==='string'&&/^\d{4}-\d{2}-\d{2}$/.test(x)?x:'';}
function safeMonth(x){if(typeof x!=='string'||!/^\d{4}-(0[1-9]|1[0-2])$/.test(x))fail('invalid-month');return x;}
function project(module,id,d){
  if(module==='documents')return {id,name:str(d.name)||id,category:str(d.category),date:date(d.expirationDate),status:str(d.visibility)};
  if(module==='events')return {id,name:str(d.title)||id,date:date(d.startDate),status:str(d.type)};
  if(module==='notifications')return {id,name:str(d.title)||'Notifica',status:str(d.eventType)};
  return {id,name:str(d.commessaName)||str(d.commessaNome)||'Rapporto ore',date:date(d.date),status:str(d.status)||str(d.stato),count:Array.isArray(d.entries)?d.entries.length:0,...(module==='approvals'&&d.tenantLabHours===true&&d.tenantLabManaged===true&&Array.isArray(d.entries)&&d.entries.length===1&&Array.isArray(d.entries[0].rows)&&d.entries[0].rows.length===1?{reviewable:true,operator:str(d.entries[0].rows[0].operatore),hours:d.entries[0].rows[0].ore}: {})};
}
async function readSupport({tx,access,caller,payload,page,id}){
  const module=payload.module,kind=payload.kind;
  if(!SUPPORT.has(module))fail('invalid-module');
  if(!['ordinary','snow'].includes(kind)||(COMMON.has(module)&&kind!=='ordinary'))fail('invalid-kind');
  const {limit,after}=page(payload),uid=id(caller.uid),root=access.root;
  const result=(rows,nextCursor=null)=>({company:access.company,module,kind,commessaId:null,rows,nextCursor});
  if(payload.commessaId!=null)fail('invalid-id');
  // Until tenant-specific credentials, ownership checks and delivery rules exist,
  // no external integration can execute or retrieve legacy configuration.
  if(module==='integrations'){
    return result(['Google Drive','Google Sheets','Email / webhook','Notifiche push','File Storage'].map((name,i)=>({id:'integration-'+i,name,status:'Non attiva — isolamento da verificare'})));
  }
  if(module==='calendar'){
    const month=safeMonth(payload.month);
    const snap=await tx.get(root.collection('sharedStaticViews').doc('calendario__'+month));
    const data=snap.exists?snap.data()?.payload:null;
    if(!data||data.month!==month)return result([]);
    return result([{id:month,name:'Riepilogo '+month,count:Array.isArray(data.reports)?data.reports.length:0,activitiesCount:Array.isArray(data.activities)?data.activities.length:0}]);
  }
  let query;
  // Even the company owner can only see metadata for documents they own here.
  // Shared and global-document permissions require a separate verified implementation.
  if(module==='documents')query=root.collection('documents').where('ownerUserId','==',uid);
  else if(module==='events')query=root.collection('calendarEvents').where('createdByUid','==',uid);
  else if(module==='notifications')query=root.collection('appNotifications').where('createdByUid','==',uid);
  else query=root.collection(module==='hours'?(kind==='snow'?'neve_ore':'oreReports'):(kind==='snow'?'neve_ore_richieste':'oreApprovalRequests'));
  query=query.orderBy('__name__').limit(limit+1);if(after)query=query.startAfter(after);
  const snap=await tx.get(query),rows=snap.docs.slice(0,limit);
  return result(rows.map(x=>project(module,x.id,x.data())),snap.size>limit?rows.at(-1).id:null);
}
module.exports={SUPPORT,readSupport};
