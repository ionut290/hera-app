// Local-only adapter for the original accounting-v2 price screen and handlers.
export function pricePreviewScript(original){
 return String.raw`
const pricePreview=(()=>{
 const state={commessa:null,prices:[],work:[],priceEditing:''};
 const archive=new Map();
 const currentUser={uid:'demo-operator'};
 const getOperatorDisplayName=()=> 'Operatore demo';
 const total=()=>{throw Error('Linked work not enabled in preview');};
 let seq=0,busy=false;
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
 `+'\n'+original.priceHandlers+String.raw`
 function render(){renderPrices();for(const b of get('price-list-tbody').querySelectorAll('[data-price-action="delete"]'))b.disabled=true;get('prezziario-meta').textContent=state.commessa.nome+' • '+state.prices.length+' voci';}
 async function open(){state.commessa=commesseById.get(managementCommessaId);if(!state.commessa)return;state.priceEditing='';get('price-search').value='';notify('');await load();openPrices();render();}
 async function act(fn){if(busy)return;busy=true;notify('');try{await fn();render();}catch(e){notify(e.message||'Operazione non riuscita.');}finally{busy=false;}}
 get('open-prezziario-btn').addEventListener('click',()=>act(open));
 get('close-prezziario-btn').addEventListener('click',()=>{if(busy)return;state.priceEditing='';get('prezziario-management-screen').classList.add('hidden');get('impianti-management-screen').classList.remove('hidden');});
 get('add-price-item-btn').addEventListener('click',()=>act(()=>addPrice()));
 get('price-search').addEventListener('input',()=>{if(!busy)render();});
 get('save-general-discount-btn').addEventListener('click',()=>act(async()=>{const d=num(get('general-discount').value);if(d==null||d<0||d>100)return notify('Percentuale ribasso non valida.');await commRef().set({percentualeRibassoGenerale:d/100,...actor()},{merge:true});state.commessa.percentualeRibassoGenerale=d/100;notify('Ribasso generale salvato.');}));
 for(const id of ['download-price-template-btn','price-file','export-price-list-btn'])get(id).disabled=true;
 get('price-list-tbody').addEventListener('click',event=>{const b=event.target.closest('[data-price-action]');if(!b||b.disabled)return;const tr=b.closest('tr'),id=tr.dataset.priceId,action=b.dataset.priceAction;act(async()=>{if(action==='edit')state.priceEditing=id;else if(action==='cancel')state.priceEditing='';else if(action==='save')await savePrice(tr);else if(action==='duplicate')await addPrice(state.prices.find(p=>p.id===id));});});
 return Object.freeze({open});
})();
`;
}
