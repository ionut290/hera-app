// No auth data, real databases, network, storage or external integrations enter this frame.
const nonce='original-demo-7f3a';
export function buildOriginalPreviewFrame(original,role){
 if(role!=='admin')throw Error('Administrator required');
 const script=String.raw`
const get=id=>document.getElementById(id);
const ui={commessaName:get('commessa-name'),commessaCode:get('commessa-code'),commessaType:get('commessa-type'),commessaParent:get('commessa-parent'),commessaForm:get('commessa-form'),commesseManageList:get('commesse-manage-list')};
const commesseById=new Map();
const auth={currentUser:{email:'demo@example.test'}};
const canManageData=()=>true;
const getCommesseCollectionName=()=> 'commesse';
const driveAccessToken=null;
const firebase={firestore:{FieldValue:{serverTimestamp:()=>new Date()}}};
const firestoreDateToMillis=v=>v instanceof Date?v.getTime():0;
const getCommessaStats=()=>({total:0});
const unavailable=()=>{get('preview-feedback').textContent='Funzione non ancora collegata in questa anteprima.';};
const openImpiantiManagement=unavailable,renameCommessa=unavailable,clearCommessaImpianti=unavailable,deleteCommessa=unavailable;
let nextId=0;
const db={collection(n){if(n!=='commesse')throw Error('Archive denied');return {async add(data){const id='demo-'+(++nextId);commesseById.set(id,normalizeCommessaDocument({id,data:()=>data}));renderList();return {id};}};}};
function renderList(){renderCommesseManagementList();for(const el of ui.commesseManageList.querySelectorAll('button')){el.disabled=true;el.title='Non ancora disponibile in questa anteprima';}}
`+'\n'+original.handlers+String.raw`
get('panel-commesse').classList.remove('hidden');
get('open-organize-commesse-btn').disabled=true;
get('open-new-commessa-btn').addEventListener('click',openNewCommessaDialog);
for(const id of ['close-new-commessa-btn','cancel-new-commessa-btn'])get(id).addEventListener('click',closeNewCommessaDialog);
get('commesse-management-search').addEventListener('input',renderList);
ui.commessaForm.addEventListener('submit',event=>createCommessa(event).catch(()=>{get('preview-feedback').textContent='Creazione demo non riuscita.';}));
document.addEventListener('submit',event=>event.preventDefault());
updateCommessaParentField();renderList();
`;
 const markup=original.markup;
 const css=original.css.replace(/<\/style/gi,'<\\/style');
 return '<!doctype html><html lang="it"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><meta http-equiv="Content-Security-Policy" content="default-src \'none\'; script-src \'nonce-'+nonce+'\'; style-src \'unsafe-inline\'; img-src data:; connect-src \'none\'; form-action \'none\'; base-uri \'none\'; object-src \'none\'"><style>'+css+'\nbody{padding:12px!important;overflow:auto!important;min-height:100vh!important}#panel-commesse{display:block!important;position:relative!important}#preview-feedback{padding:10px}</style></head><body><p id="preview-feedback" role="status" class="muted">Anteprima originale: dati temporanei, nessuna scrittura online.</p>'+markup+'<script nonce="'+nonce+'">'+script.replace(/<\/script/gi,'<\\/script')+'</script></body></html>';
}
