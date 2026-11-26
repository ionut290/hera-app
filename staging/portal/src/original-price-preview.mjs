import {parsePriceCsv,planPriceImport} from './price-import.mjs';
import {inspectXlsxZip,verifyXlsxInflation,readPriceXlsx} from './price-xlsx.mjs';
import {projectPreviewWork,summarizePreviewWork} from './work-price-projection.mjs';
import {planMatrixImport} from './matrix-import.mjs';
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
 `+'\n'+original.priceHandlers+'\n'+parsePriceCsv.toString()+'\n'+planPriceImport.toString()+'\n'+inspectXlsxZip.toString()+'\n'+verifyXlsxInflation.toString()+'\n'+readPriceXlsx.toString()+'\n'+projectPreviewWork.toString()+'\n'+summarizePreviewWork.toString()+'\n'+planMatrixImport.toString()+'\n'+(original.excelHandlers||'')+String.raw`
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
 const matrixInput=get('excel-file'),matrixButton=get('import-btn'),matrixFeedback=get('import-feedback');
 let matrixFile=null,matrixPending=null,matrixBusy=false,matrixGeneration=0;
 const matrixConfirm=document.createElement('button');matrixConfirm.type='button';matrixConfirm.className='btn btn-primary';matrixConfirm.textContent='Conferma matrice cantieri';matrixConfirm.hidden=true;matrixFeedback.after(matrixConfirm);
 const matrixCancel=document.createElement('button');matrixCancel.type='button';matrixCancel.className='btn';matrixCancel.textContent='Annulla matrice';matrixCancel.hidden=true;matrixConfirm.after(matrixCancel);
 const clearMatrix=()=>{matrixGeneration++;matrixPending=null;matrixFile=null;matrixInput.value='';matrixConfirm.hidden=true;matrixCancel.hidden=true;matrixButton.disabled=true;};
 matrixInput.accept=excelReady?'.xlsx,.csv':'.csv';
 get('toggle-import-impianti-btn').addEventListener('click',()=>get('impianti-import-card').classList.toggle('hidden'));
 matrixInput.addEventListener('change',()=>{if(matrixBusy)return;const file=matrixInput.files[0];clearMatrix();matrixFile=file||null;matrixButton.disabled=!matrixFile;matrixFeedback.textContent=file?'File selezionato: '+file.name:'';});
 matrixButton.addEventListener('click',async()=>{
  if(matrixBusy||!matrixFile)return;
  const id=managementCommessaId,contract=commesseById.get(id),file=matrixFile;
  if(!contract)return;
  matrixBusy=true;clearMatrix();const generation=matrixGeneration;matrixFeedback.textContent='';
  try{
   const isExcel=/\.xlsx$/i.test(file.name);
   if(isExcel?(!excelReady||file.size>5000000):(!/\.csv$/i.test(file.name)||file.size>1000000))throw Error('Seleziona un CSV o XLSX valido.');
   const existing=impiantiByCommessaId.get(id)||[],snapshot=JSON.stringify(existing);
   const input=isExcel?await readPriceXlsx(await file.arrayBuffer(),globalThis.ExcelJS,aliases,key,'matrix'):parsePriceCsv(await file.text());
   const plan=planMatrixImport(input,existing,{mapped,norm,num,normalizeCode:core.normalizePriceCode});
   if(generation!==matrixGeneration||id!==managementCommessaId||get('impianti-management-screen').classList.contains('hidden'))return;
   matrixPending={id,snapshot,plan};matrixConfirm.hidden=false;matrixCancel.hidden=false;
   matrixFeedback.textContent='Anteprima: '+plan.groups.length+' nuovi cantieri, '+plan.workCount+' lavorazioni. Coordinate, stati e dati storici non importati.';
  }catch(e){if(generation===matrixGeneration)matrixFeedback.textContent=e.message||'Importazione non riuscita.';}finally{matrixBusy=false;}
 });
 matrixCancel.addEventListener('click',()=>{clearMatrix();matrixFeedback.textContent='Importazione annullata.';});
 matrixConfirm.addEventListener('click',()=>{
  const selected=matrixPending;if(!selected||matrixBusy)return;clearMatrix();
  const {id,snapshot,plan}=selected,contract=commesseById.get(id),existing=impiantiByCommessaId.get(id)||[];
  if(!contract||id!==managementCommessaId||get('impianti-management-screen').classList.contains('hidden')||JSON.stringify(existing)!==snapshot){matrixFeedback.textContent='Commessa o cantieri cambiati: ripetere anteprima.';return;}
  let number=Math.max(0,(Number(contract.nextImpiantoNumber)||1)-1,...existing.map(p=>Number(p.numeroProgressivo)||0));
  const next=existing.slice(),quantities=new Map(quantitiesFor(id));
  for(const group of plan.groups){
   const codes=group.work.map(w=>w.code).join(' | '),ordinary=hasOrdinario(codes),extra=hasStraordinario(codes),plantId='plant-'+(++nextId);
   const progress=++number;
   const plant={...group.plant,id:plantId,codicePrezzo:codes,voceRiferimento:'',sfalci:'',areaMq:null,frequenzaAnnua:group.work.map(w=>w.frequency).join(' | '),tipologiaIntervento:group.work.map(w=>w.description).join(' | '),lavorazioniRichieste:'',gpsY:null,gpsX:null,numeroProgressivo:progress,sortOrder:progress,hasOrdinario:ordinary,hasStraordinario:extra,tipoManutenzione:ordinary?(extra?'Ordinaria + Straordinaria':'Ordinaria'):(extra?'Straordinaria':'Non specificata'),done:false,doneAt:null,doneBy:'',createdAt:new Date()};
   next.push(plant);
   for(const w of group.work)quantities.set(JSON.stringify([plantId,w.code]),w.quantity);
  }
  impiantiByCommessaId.set(id,next);quantityArchive.set(id,quantities);contract.nextImpiantoNumber=number+1;
  refreshPlants();matrixFeedback.textContent='Importati '+plan.groups.length+' cantieri e '+plan.workCount+' lavorazioni demo.';
 });
 for(const id of ['close-impianti-management-btn','open-prezziario-btn'])get(id).addEventListener('click',clearMatrix);
 get('download-excel-template-btn').disabled=!excelReady;
 get('download-excel-template-btn').addEventListener('click',async()=>{
  const id=managementCommessaId,contract=commesseById.get(id);if(!contract||!excelReady)return;
  try{const wb=new ExcelJS.Workbook(),ws=wb.addWorksheet('Matrice impianti');ws.addRow(['Distretto','ID SAP','Denominazione Impianto','Comune ubicazione Impianto','Via e civico di ubicazione Impianto','Voce di Riferimento Elenco Prezzi','Quantità','Frequenza annua minima','Tipologia di lavorazione / sfalcio','Stato']);ws.addRow(['Nord','SAP-ESEMPIO','ESEMPIO - NON IMPORTARE','Bologna','Via Roma 1','A11',100,1,'Sfalcio','DA FARE']);sheetStyle(ws,[14,18,32,28,32,24,12,20,32,16]);await download(await wb.xlsx.writeBuffer(),'MODELLO_MATRICE_CANTIERI.xlsx');}catch(e){matrixFeedback.textContent=e.message||'Esportazione non riuscita.';}
 });
 return Object.freeze({open,renderWork});
})();
`;
}
