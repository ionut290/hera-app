// Local emulator UI. Server remains authoritative; never exposed in live build.
export function createCompanyDocuments({host,auth,isActive,tenantId,plant=null}){
 let closed=false,busy=false,cursor=null,writable=false,sequence=0,previewUrl=null;
 const el=(tag,text)=>{const n=document.createElement(tag);if(text)n.textContent=text;return n;};
 const title=el('h2',plant?'Fotografie impianto':'Documenti aziendali'),status=el('p'),usage=el('p'),bar=el('progress'),list=el('div'),input=el('input'),upload=el('button','Carica documento'),more=el('button','Altri documenti'),close=el('button','Torna al cantiere');
 status.setAttribute('role','status');usage.setAttribute('aria-live','polite');bar.setAttribute('aria-label','Spazio archiviazione');input.type='file';input.accept=plant?'.jpg,.jpeg,.png,.webp':'.pdf,.jpg,.jpeg,.png,.webp';input.setAttribute('aria-label','Seleziona documento');for(const b of [upload,more,close])b.type='button';
 const visibility=el('select');visibility.setAttribute('aria-label','Visibilità documento');for(const [v,label] of [['private','Privato'],['company','Condiviso con l’azienda']]){const opt=el('option',label);opt.value=v;visibility.append(opt);}
 const preview=el('img');preview.hidden=true;preview.alt='';preview.style.maxWidth='100%';preview.style.maxHeight='60vh';
 const uid=auth.currentUser?.uid;
 const valid=()=>!closed&&isActive()&&auth.currentUser?.uid===uid;
 const ensure=()=>{if(!valid())throw Error('SESSION_CHANGED');};
 const format=n=>Number.isSafeInteger(n)?(n/1000000).toLocaleString('it-IT',{maximumFractionDigits:4})+' MB':'non verificato';
 const errors={STORAGE_LIMIT:'Spazio disponibile esaurito.',PLAN_EXPIRED:'Abbonamento scaduto: solo consultazione.',FILE_PENDING_RECONCILIATION:'Operazione in verifica. Non ripetere il caricamento.',FORBIDDEN:'Accesso non autorizzato.',INVALID_FILE:'File non valido o troppo grande (massimo 15 MiB).',INVALID_FILE_TYPE:'Formato non supportato.'};
 async function api(action,data={}){
  ensure();const token=await auth.currentUser.getIdToken();ensure();
  const response=await fetch('/__v2_documents/'+action,{method:'POST',headers:{'Content-Type':'application/json',Authorization:'Bearer '+token},body:JSON.stringify({...data,tenantId})});
  const result=await response.json();ensure();if(!response.ok)throw Error(result.code||'REQUEST_FAILED');return result;
 }
 function sync(){input.disabled=visibility.disabled=upload.disabled=busy||!writable;more.disabled=busy;more.hidden=!cursor;for(const b of list.querySelectorAll('button'))b.disabled=busy||(b.dataset.remove==='true'&&!writable);}
 async function task(fn){if(busy||!valid())return;busy=true;sync();status.textContent='Operazione in corso…';try{await fn();if(valid())status.textContent='';}catch(e){if(valid())status.textContent=errors[e.message]||'Operazione non riuscita. Riprova.';}finally{busy=false;if(valid())sync();}}
 function renderFile(f){
  const row=el('article'),label=el('p',f.name+' · '+format(f.size)+(f.status==='ready'?' · '+(f.visibility==='company'?'Condiviso':'Privato'):' · in verifica'));row.append(label);
  if(f.status==='ready'){const download=el('button','Scarica');download.onclick=()=>task(async()=>{const d=await api('download',{fileId:f.id}),bytes=Uint8Array.from(atob(d.base64),c=>c.charCodeAt(0)),url=URL.createObjectURL(new Blob([bytes],{type:d.type}));const a=el('a');a.href=url;a.download=d.name;a.click();setTimeout(()=>URL.revokeObjectURL(url),1000);});row.append(download);if(f.type.startsWith('image/')){const show=el('button','Anteprima foto');show.onclick=()=>task(async()=>{const d=await api('download',{fileId:f.id}),bytes=Uint8Array.from(atob(d.base64),c=>c.charCodeAt(0));if(previewUrl)URL.revokeObjectURL(previewUrl);previewUrl=URL.createObjectURL(new Blob([bytes],{type:d.type}));preview.src=previewUrl;preview.alt=d.name;preview.hidden=false;});row.append(show);}}
  if(f.canManage&&f.status==='ready'){const share=el('button',f.visibility==='company'?'Rendi privato':'Condividi con azienda');share.dataset.remove='true';share.onclick=()=>{const v=f.visibility==='company'?'private':'company';if(v==='company'&&!confirm('Rendere il file visibile a tutti gli utenti attivi di questa azienda?'))return;task(async()=>{await api('share',{fileId:f.id,visibility:v});await refresh();});};row.append(share);}
  if(f.canManage&&['ready','deleting'].includes(f.status)){const remove=el('button',f.status==='deleting'?'Completa eliminazione':'Elimina');remove.dataset.remove='true';remove.onclick=()=>{if(!confirm('Eliminare definitivamente '+f.name+'?'))return;task(async()=>{await api('remove',{fileId:f.id});await refresh();});};row.append(remove);}
  list.append(row);
 }
 async function refresh(next=false){
  const attempt=++sequence,result=await api('list',{cursor:next?cursor:null,plant});if(attempt!==sequence)return;
  writable=result.writable===true;usage.textContent=format(result.used)+' utilizzati su '+format(result.limit)+(writable?'':' · solo consultazione');bar.max=Math.max(1,result.limit||1);bar.value=Math.min(bar.max,result.used||0);
  if(!next){list.replaceChildren();if(previewUrl)URL.revokeObjectURL(previewUrl);previewUrl=null;preview.removeAttribute('src');preview.hidden=true;}for(const f of result.files)renderFile(f);cursor=result.nextCursor;if(!list.children.length)list.append(el('p','Nessun documento disponibile.'));sync();
 }
 upload.onclick=()=>task(async()=>{const f=input.files?.[0];if(!f||!f.size||f.size>15*1024*1024)throw Error('INVALID_FILE');const bytes=new Uint8Array(await f.arrayBuffer());ensure();let binary='';for(let i=0;i<bytes.length;i+=32768)binary+=String.fromCharCode(...bytes.subarray(i,i+32768));await api('upload',{fileId:crypto.randomUUID(),name:f.name,type:f.type,base64:btoa(binary),visibility:visibility.value,plant});input.value='';await refresh();});
 more.onclick=()=>task(()=>refresh(true));close.onclick=()=>destroy();
 function destroy(){closed=true;sequence++;if(previewUrl)URL.revokeObjectURL(previewUrl);previewUrl=null;host.hidden=true;host.replaceChildren();}
 host.replaceChildren(title,close,usage,bar,input,visibility,upload,status,preview,list,more);host.hidden=false;sync();task(()=>refresh());
 return Object.freeze({close:destroy});
}
