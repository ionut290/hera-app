// No auth data, real databases, network, storage or external integrations enter this frame.
const nonce='original-demo-7f3a';
export function buildOriginalPreviewFrame(original,role){
 if(role!=='admin')throw Error('Administrator required');
 const script=String.raw`
const get=id=>document.getElementById(id);
const ui={commessaName:get('commessa-name'),commessaCode:get('commessa-code'),commessaType:get('commessa-type'),commessaParent:get('commessa-parent'),commessaForm:get('commessa-form'),commesseManageList:get('commesse-manage-list')};
Object.assign(ui,{commessaTargetSelect:get('commessa-target-select'),manualImpiantoForm:get('manual-impianto-form'),manualImpiantoDenominazione:get('manual-impianto-denominazione'),manualImpiantoComune:get('manual-impianto-comune'),manualImpiantoIndirizzo:get('manual-impianto-indirizzo'),manualImpiantoCodice:get('manual-impianto-codice'),manualImpiantoFeedback:get('manual-impianto-feedback')});
const commesseById=new Map();
const impiantiByCommessaId=new Map();
let selectedCommessaId='',selectedCommessaName='';const currentImpianti=[];
const auth={currentUser:{email:'demo@example.test'}};
const canManageData=()=>true;
const getCommesseCollectionName=()=> 'commesse';
const driveAccessToken=null;
const firebase={firestore:{FieldValue:{serverTimestamp:()=>new Date()}}};
const firestoreDateToMillis=v=>v instanceof Date?v.getTime():0;
const getCommessaStats=id=>({total:(impiantiByCommessaId.get(id)||[]).length});
const unavailable=()=>{get('preview-feedback').textContent='Funzione non ancora collegata in questa anteprima.';};
const renameCommessa=unavailable,clearCommessaImpianti=unavailable,deleteCommessa=unavailable;
let nextId=0;
const db={collection(n){
 if(n!=='commesse')throw Error('Archive denied');
 return {
 async add(data){const id='demo-'+(++nextId);commesseById.set(id,normalizeCommessaDocument({id,data:()=>data}));renderList();return {id};},
 doc(id){
 if(!commesseById.has(id))throw Error('Parent missing');
 return {
 async set(data,options){if(options?.merge!==true||Object.keys(data).some(k=>k!=='nextImpiantoNumber'))throw Error('Update denied');Object.assign(commesseById.get(id),data);},
 collection(name){
 if(name!=='impianti')throw Error('Archive denied');
 if(!impiantiByCommessaId.has(id))impiantiByCommessaId.set(id,[]);
 const rows=impiantiByCommessaId.get(id);
 return {
 async get(){return {docs:rows.map(row=>({id:row.id,data:()=>({...row})}))};},
 async add(data){if(data.done!==false||data.doneAt!==null)throw Error('Completion denied');const plantId='plant-'+(++nextId);rows.push({...data,id:plantId});return {id:plantId};},
 doc(plantId){return {async update(data){const row=rows.find(r=>r.id===plantId);if(!row||Object.keys(data).some(k=>!['codicePrezzo','voceRiferimento','tipologiaIntervento','lavorazioniRichieste','frequenzaAnnua','hasOrdinario','hasStraordinario','tipoManutenzione'].includes(k)))throw Error('Update denied');Object.assign(row,data);}};}
 };
 }
 };
 }
 };
}};
function bindLocalForm(form,handler,feedback){
 let running=false;
 form.addEventListener('submit',async event=>{
  event.preventDefault();
  if(running)return;
  running=true;
  const buttons=Array.from(form.querySelectorAll('button[type="submit"]'));
  const originalStates=buttons.map(b=>b.disabled);
  buttons.forEach(b=>b.disabled=true);
  try{await handler(event);}catch{feedback.textContent='Operazione demo non riuscita.';}
  finally{running=false;buttons.forEach((b,i)=>b.disabled=originalStates[i]);}
 });
}
function renderList(){
 renderCommesseManagementList();
 for(const el of ui.commesseManageList.querySelectorAll('button')){
 if(!['APRI','Gestisci impianti'].includes(el.textContent.trim())){el.disabled=true;el.title='Non ancora disponibile in questa anteprima';}
 else el.addEventListener('click',()=>{selectedCommessaId=managementCommessaId;selectedCommessaName=commesseById.get(managementCommessaId)?.nome||'';refreshPlants();});
 }
}
function refreshPlants(){
 managementEditingId='';
 renderImpiantiManagementTable();
 const n=getCommessaStats(managementCommessaId).total;
 get('impianti-management-meta').textContent='Cod. '+(commesseById.get(managementCommessaId)?.codice||'—')+' • '+n+' '+(n===1?'impianto':'impianti');
 for(const el of get('impianti-management-screen').querySelectorAll('[data-row-action],[data-select-row],[data-select-all]'))el.disabled=true;
}

`+'\n'+original.handlers+String.raw`
get('panel-commesse').classList.remove('hidden');
get('open-organize-commesse-btn').disabled=true;
get('open-new-commessa-btn').addEventListener('click',openNewCommessaDialog);
for(const id of ['close-new-commessa-btn','cancel-new-commessa-btn'])get(id).addEventListener('click',closeNewCommessaDialog);
get('commesse-management-search').addEventListener('input',renderList);
bindLocalForm(ui.commessaForm,createCommessa,get('preview-feedback'));
document.addEventListener('submit',event=>event.preventDefault());
updateCommessaParentField();renderList();
get('close-impianti-management-btn').addEventListener('click',()=>{closeImpiantiManagement();selectedCommessaId='';selectedCommessaName='';renderList();});
const manualDetails=ui.manualImpiantoForm.closest('details');
get('add-management-impianto-btn').addEventListener('click',()=>{manualDetails.open=true;ui.manualImpiantoDenominazione.focus();});
for(const id of ['toggle-import-impianti-btn','export-all-impianti-btn','open-prezziario-btn','sheet-url-import-btn','download-excel-template-btn','excel-file','sheet-url'])get(id).disabled=true;
get('snow-roads-section').classList.add('hidden');
for(const id of ['parent-commessa-form','move-subcommesse-form','snow-roads-form'])for(const el of get(id).querySelectorAll('input,select,textarea,button'))el.disabled=true;
bindLocalForm(ui.manualImpiantoForm,async event=>{await addManualImpianto(event);refreshPlants();},ui.manualImpiantoFeedback);
get('impianti-management-search').addEventListener('input',refreshPlants);
get('clear-impianti-management-search').addEventListener('click',()=>{get('impianti-management-search').value='';refreshPlants();});
for(const id of ['impianti-status-filter','impianti-comune-filter','impianti-tipologia-filter','impianti-operatore-filter'])get(id).addEventListener('change',refreshPlants);
`;
 const markup=original.markup;
 const css=original.css.replace(/<\/style/gi,'<\\/style');
 return '<!doctype html><html lang="it"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><meta http-equiv="Content-Security-Policy" content="default-src \'none\'; script-src \'nonce-'+nonce+'\'; style-src \'unsafe-inline\'; img-src data:; connect-src \'none\'; form-action \'none\'; base-uri \'none\'; object-src \'none\'"><style>'+css+'\nbody{padding:12px!important;overflow:auto!important;min-height:100vh!important}#panel-commesse{display:block!important;position:relative!important}#preview-feedback{padding:10px}</style></head><body><p id="preview-feedback" role="status" class="muted">Anteprima originale: dati temporanei, nessuna scrittura online.</p>'+markup+'<script nonce="'+nonce+'">'+script.replace(/<\/script/gi,'<\\/script')+'</script></body></html>';
}
