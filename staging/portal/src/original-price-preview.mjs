import {parsePriceCsv,planPriceImport} from './price-import.mjs';
import {inspectXlsxZip,verifyXlsxInflation,readPriceXlsx} from './price-xlsx.mjs';
import {projectPreviewWork,summarizePreviewWork} from './work-price-projection.mjs';
// Local-only adapter for the original accounting-v2 price screen and handlers.
export function pricePreviewScript(original){
 return String.raw`
const pricePreview=(()=>{
 const state={commessa:null,prices:[],work:[],priceEditing:''};
 const archive=new Map(),quantityArchive=new Map();
 const currentUser={uid:'demo-operator'};
 const getOperatorDisplayName=()=> 'Operatore demo';
 const total=()=>{throw Error('Linked work not enabled in preview');};
 let seq=0,busy=false,pending=null;
 function rows(id){if(!archive.has(id))archive.set(id,new Map());return archive.get(id);}
 function valid(row){if(row.prezzoBase!=null&&(!Number.isFinite(row.prezzoBase)||row.prezzoBase<0))throw Error('Prezzo non valido');if(row.percentualeRibasso!=null&&(!Number.isFinite(row.percentualeRibasso)||row.percentualeRibasso<0||row.percentualeRibasso>1))throw Error('Ribasso non valido');}
 const db={collection(name){if(name!=='commesse')throw Error('Archive denied');return {doc(id){if(id!==state.commessa?.id||!commesseById.has(id))throw Error('Scope denied');return {
 collection(name){if(name!=='prezziario')throw Error('Archive denied');return {doc(key='price-'+(++seq)){return {id:key,async set(data,options){const map=rows(id);if(options?.merge&&!map.has(key))throw Error('Missing price');const row={...(options?.merge?map.get(key):{}),...data,id:key};if(row.commessaId!==id)throw Error('Scope denied');valid(row);map.set(key,row);}};}};},
 async set(data,options){if(options?.merge!==true||Object.keys(data).some(k=>!['percentualeRibassoGenerale','updatedAt','updatedBy','updatedByName'].includes(k))||!Number.isFinite(data.percentualeRibassoGenerale)||data.percentualeRibassoGenerale<0||data.percentualeRibassoGenerale>1)throw Error('Ribasso non valido');Object.assign(commesseById.get(id),data);}
 };}};},batch(){const writes=[];return {set(ref,data,options){writes.push(()=>ref.set(data,options));},async commit(){for(const write of writes)await write();}};}};
 async function load(){state.prices=Array.from(rows(state.commessa.id).values(),p=>({...p}));}
 const notify=message=>{get('price-feedback').textContent=message;};
 const alert=notify;
 const confirm=()=>false;
 `+'\n'+original.priceHandlers+'\n'+parsePriceCsv.toString()+'\n'+planPriceImport.toString()+'\n'+inspectXlsxZip.toString()+'\n'+verifyXlsxInflation.toString()+'\n'+readPriceXlsx.toString()+'\n'+projectPreviewWork.toString()+'\n'+summarizePreviewWork.toString()+'\n'+(original.excelHandlers||'')+String.raw`
 function clearPending(){pending=null;importControls.hidden=true;}
 const workPanel=document.createElement('details');workPanel.className='import-mode-card';workPanel.id='preview-work-panel';
 const heading=document.createElement('summary');heading.textContent='Lavorazioni e importi stimati';workPanel.append(heading);
 const workBody=document.createElement('div');workPanel.append(workBody);
 get('impianti-management-meta').closest('header').after(workPanel);
 function quantitiesFor(id){if(!quantityArchive.has(id))quantityArchive.set(id,new Map());return quantityArchive.get(id);}
 function renderWork(){
  const id=managementCommessaId,contract=commesseById.get(id);
  workBody.replaceChildren();if(!contract)return;
  const quantities=quantitiesFor(id),items=projectPreviewWork(impiantiByCommessaId.get(id)||[],Array.from(rows(id).values()),contract.percentualeRibassoGenerale,quantities,core);
  const report=summarizePreviewWork(items);
  const note=document.createElement('p');note.id='preview-work-summary';
  note.textContent='Importo stimato: '+money.format(report.estimated)+' • '+report.verified+'/'+report.rows+' lavorazioni valorizzate • '+report.missing+' da verificare. Non è un consuntivo.';
  workBody.append(note);
  const wrap=document.createElement('div');wrap.className='impianti-sheet-wrap';const table=document.createElement('table');table.className='impianti-sheet';table.id='preview-work-table';wrap.append(table);workBody.append(wrap);
  const header=table.createTHead().insertRow();
  for(const label of ['Impianto','Codice voce','Quantità','u.m.','Prezzo ribassato','Importo stimato','Verifica','Azioni']){const th=document.createElement('th');th.textContent=label;header.append(th);}
  const tbody=table.createTBody();
  for(const item of items){
   const tr=tbody.insertRow();
   const cell=value=>{const td=tr.insertCell();td.textContent=value;return td;};
   cell(item.denominazione||'');cell(item.codiceVocePrezzo||'—');
   const quantityCell=tr.insertCell(),input=document.createElement('input');input.type='number';input.min='0';input.max='100000000';input.step='any';input.value=item.quantita??'';input.setAttribute('aria-label','Quantità '+(item.denominazione||'')+' '+(item.codiceVocePrezzo||''));quantityCell.append(input);
   const fixed=String(item.unitaMisura||'').trim().toUpperCase()==='AC';input.disabled=fixed;
   cell(item.unitaMisura||'—');cell(item.prezzoRibassato==null?'—':money.format(item.prezzoRibassato));cell(item.totale==null?'—':money.format(item.totale));cell(item.priceListLinkStatus==='MISSING'?'Prezzo mancante o ambiguo':item.totale==null?'Quantità da verificare':'OK');
   const save=document.createElement('button');save.type='button';save.className='btn';save.textContent='Salva quantità';save.disabled=fixed;
   tr.insertCell().append(save);
   save.addEventListener('click',()=>{
    if(managementCommessaId!==id)return;
    const value=input.value.trim()===''?null:num(input.value);
    if(!input.validity.valid||value!==null&&(!Number.isFinite(value)||value<0||value>100000000)){input.reportValidity();return;}
    quantities.set(item.id,value);renderWork();
   });
  }
 }

 function render(){renderPrices();for(const b of get('price-list-tbody').querySelectorAll('[data-price-action="delete"]'))b.disabled=true;get('prezziario-meta').textContent=state.commessa.nome+' • '+state.prices.length+' voci';}
 async function open(){state.commessa=commesseById.get(managementCommessaId);if(!state.commessa)return;clearPending();state.priceEditing='';get('price-search').value='';notify('');await load();openPrices();get('impianti-management-screen').setAttribute('aria-hidden','true');get('prezziario-management-screen').setAttribute('aria-hidden','false');render();}
 async function act(fn){if(busy)return;clearPending();busy=true;notify('');try{await fn();render();}catch(e){notify(e.message||'Operazione non riuscita.');}finally{busy=false;}}
 get('open-prezziario-btn').addEventListener('click',()=>act(open));
 get('close-prezziario-btn').addEventListener('click',()=>{if(busy)return;clearPending();state.priceEditing='';get('prezziario-management-screen').classList.add('hidden');get('prezziario-management-screen').setAttribute('aria-hidden','true');get('impianti-management-screen').classList.remove('hidden');get('impianti-management-screen').setAttribute('aria-hidden','false');renderWork();});
 get('add-price-item-btn').addEventListener('click',()=>act(()=>addPrice()));
 get('price-search').addEventListener('input',()=>{if(!busy)render();});
 get('save-general-discount-btn').addEventListener('click',()=>act(async()=>{const d=num(get('general-discount').value);if(d==null||d<0||d>100)return notify('Percentuale ribasso non valida.');await commRef().set({percentualeRibassoGenerale:d/100,...actor()},{merge:true});state.commessa.percentualeRibassoGenerale=d/100;notify('Ribasso generale salvato.');}));
 const excelReady=Boolean(globalThis.ExcelJS?.Workbook);
 for(const [id,template] of [['download-price-template-btn',true],['export-price-list-btn',false]]){
  const button=get(id);button.disabled=!excelReady;
  button.addEventListener('click',()=>act(async()=>{if(!excelReady||state.priceEditing)throw Error('Concludi la modifica prima di esportare.');await exportWorkbook(true,template);notify('File Excel preparato.');}));
 }
 const fileInput=get('price-file');fileInput.accept=excelReady?'.xlsx,.csv':'.csv,text/csv';fileInput.title=excelReady?'Anteprima locale: XLSX e CSV':'Anteprima locale: solo CSV';
 const importControls=document.createElement('div');importControls.hidden=true;
 const apply=document.createElement('button');apply.type='button';apply.className='btn btn-primary';apply.textContent='Conferma importazione CSV';
 const cancel=document.createElement('button');cancel.type='button';cancel.className='btn';cancel.textContent='Annulla importazione CSV';
 importControls.append(apply,cancel);get('price-feedback').after(importControls);
 fileInput.addEventListener('change',async()=>{
  const file=fileInput.files[0];fileInput.value='';if(!file||busy||!state.commessa)return;
  clearPending();notify('');busy=true;
  try{
   if(state.priceEditing)throw Error('Concludi la modifica della voce prima di importare.');
   const isExcel=/\.xlsx$/i.test(file.name),format=isExcel?'XLSX':'CSV';
   if(isExcel?(!excelReady||file.size>5000000):(!/\.csv$/i.test(file.name)||file.size>1000000))throw Error('Seleziona un CSV (max 1 MB) o XLSX (max 5 MB, se disponibile).');
   const commessaId=state.commessa.id;
   const data=isExcel?await readPriceXlsx(await file.arrayBuffer(),globalThis.ExcelJS,aliases,key):parsePriceCsv(await file.text());
   const plan=planPriceImport(data,state.prices,{mapped,num,norm,discounted});
   if(commessaId!==state.commessa?.id||commessaId!==managementCommessaId)throw Error('Commessa cambiata.');
   pending={commessaId,plan};const updates=plan.filter(p=>p.id!==null).length;
   apply.textContent='Conferma importazione '+format;cancel.textContent='Annulla importazione '+format;
   notify('Anteprima '+format+': '+(plan.length-updates)+' nuove voci, '+updates+' aggiornamenti. Confermare?');importControls.hidden=false;
  }catch(e){notify(e.message||'Importazione non riuscita.');}finally{busy=false;}
 });
 cancel.addEventListener('click',()=>{clearPending();notify('Importazione annullata.');});
 apply.addEventListener('click',()=>{
  if(!pending||busy)return;const {commessaId,plan}=pending;
  act(async()=>{
   if(commessaId!==state.commessa?.id||commessaId!==managementCommessaId)throw Error('Commessa cambiata.');
   const nextRows=new Map(rows(commessaId));
   for(const item of plan){
    const id=item.id??'price-'+(++seq),old=nextRows.get(id);
    if(item.id&&!old)throw Error('Voce non trovata.');
    const row={...(old||{createdAt:server(),createdBy:currentUser.uid}),...item.data,...actor(),commessaId,id};
    valid(row);nextRows.set(id,row);
   }
   archive.set(commessaId,nextRows);await load();notify('Importate '+plan.length+' voci nel prezziario demo.');
  });
 });
 get('price-list-tbody').addEventListener('click',event=>{const b=event.target.closest('[data-price-action]');if(!b||b.disabled)return;const tr=b.closest('tr'),id=tr.dataset.priceId,action=b.dataset.priceAction;act(async()=>{if(action==='edit')state.priceEditing=id;else if(action==='cancel')state.priceEditing='';else if(action==='save')await savePrice(tr);else if(action==='duplicate')await addPrice(state.prices.find(p=>p.id===id));});});
 return Object.freeze({open,renderWork});
})();
`;
}
