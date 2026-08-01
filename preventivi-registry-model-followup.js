(()=>{'use strict';
const P=window.HeraPreventivi,M=window.HeraPreventiviModels,R=window.HeraPreventiviRegistry;if(!P||!M||!R)return;
/*
 * HOTFIX BLOCCO PREVENTIVI
 * Il modulo precedente aggiorna le etichette dei pulsanti tramite MutationObserver.
 * Impostare textContent allo stesso valore generava comunque una nuova mutazione,
 * riattivando l'observer senza fine. Per i soli pulsanti di esportazione evitiamo
 * la scrittura quando il testo è già identico.
 */
if(!window.__preventiviExportTextGuard&&window.Node&&window.HTMLButtonElement){
  const descriptor=Object.getOwnPropertyDescriptor(Node.prototype,'textContent');
  if(descriptor?.get&&descriptor?.set){
    window.__preventiviExportTextGuard=true;
    Object.defineProperty(HTMLButtonElement.prototype,'textContent',{
      configurable:true,
      enumerable:descriptor.enumerable,
      get(){return descriptor.get.call(this)},
      set(value){
        const next=String(value??'');
        if(this.matches?.('[data-pvm-export]')&&descriptor.get.call(this)===next)return;
        descriptor.set.call(this,value);
      }
    });
  }
}
const C=v=>String(v??'').trim(),enrich=(record,fd)=>{if(!record||!fd)return record;const data=R.registry(),cid=C(fd.get('commessaId')),pid=C(fd.get('plantId')),c=data.commesse.find(x=>x.id===cid),p=data.plants.find(x=>x.id===pid);return Object.assign(record,{commessaId:cid,commessaName:c?.name||record.commessaName||'',commessaCode:C(fd.get('commessaCode'))||c?.code||record.commessaCode||'',plantId:pid,plantName:p?.name||record.plantName||'',plantSap:C(fd.get('plantSap'))||p?.sap||record.plantSap||'',workLocation:C(fd.get('workLocation'))||[p?.address,p?.city].filter(Boolean).join(', ')||record.workLocation||'',city:C(fd.get('city'))||p?.city||record.city||'',plantType:C(fd.get('plantType'))||p?.type||record.plantType||''})};
let pending=null;const remote=P.saveRemote?.bind(P);if(remote)P.saveRemote=(collection,record)=>{if(collection===P.collections.quotes&&pending){enrich(record,pending);const local=(P.state.quotes||[]).find(x=>x.id===record.id);if(local)Object.assign(local,record);P.persistLocal?.();pending=null}return remote(collection,record)};const save=P.saveQuote?.bind(P);if(save)P.saveQuote=async form=>{const fd=new FormData(form);pending=fd;const out=await save(form);setTimeout(()=>{if(pending===fd)pending=null},30000);return out};
const exact=M.exportOriginal?.bind(M);if(exact)M.exportOriginal=(doc,type)=>{const model=M.getModel(doc?.modelId);if(!model?.builtIn)return exact(doc,type);const html=`<!doctype html><html lang="it"><head><meta charset="utf-8"><title>${P.escapeHtml(doc.number||'Documento')}</title></head><body>${M.previewHtml(doc,type)}</body></html>`;M.download(new Blob([html],{type:'text/html;charset=utf-8'}),`${M.fileName(doc.number)}.html`);P.setFeedback('Modello standard compilato.','success')};
const DOCX='https://cdn.jsdelivr.net/npm/docxtemplater@3.51.0/build/docxtemplater.js';M.exportDocx=async(model,stored,data)=>{if(!(model.fields||[]).some(f=>f.token))throw Error(`Il modello originale “${model.originalName||model.name}” non contiene segnaposto compilabili.`);await M.ensureScript('PizZip',M.constants.ZIP_LIB,'pizzip');const D=await M.ensureScript(['docxtemplater','Docxtemplater'],DOCX,'docxtemplater'),zip=new PizZip(stored.buffer);Object.keys(zip.files).filter(n=>/^word\/.*\.xml$/i.test(n)).forEach(n=>{const e=zip.file(n);if(e)zip.file(n,e.asText().replace(/\[\[\s*([^\[\]]+?)\s*\]\]/g,'{{$1}}'))});const doc=new D(zip,{paragraphLoop:true,linebreaks:true,delimiters:{start:'{{',end:'}}'},nullGetter:()=>''});doc.render(data);const ext=String(model.format||'docx'),name=String(model.originalName||model.name||'modello.docx').replace(/[\\/:*?"<>|]/g,'_').replace(/\.docx$/i,'');M.download(doc.getZip().generate({type:'blob',mimeType:stored.type||'application/vnd.openxmlformats-officedocument.wordprocessingml.document'}),`${name}-compilato.${ext}`)};
const loadScript=(marker,src,error)=>new Promise(resolve=>{
  const existing=document.querySelector(`script[${marker}]`);if(existing){if(existing.dataset.loaded==='1')resolve();else existing.addEventListener('load',resolve,{once:true});return;}
  const script=document.createElement('script');script.src=src;script.async=false;script.setAttribute(marker,'1');script.addEventListener('load',()=>{script.dataset.loaded='1';resolve()},{once:true});script.addEventListener('error',()=>{console.warn(error);resolve();},{once:true});document.head.appendChild(script);
});
const loadModelModules=async()=>{
  await loadScript('data-preventivi-matrix-profile','./preventivi-matrix-form-profile.js?v=20260801a','Profilo form matrice non caricato.');
  await loadScript('data-preventivi-matrix-xlsx-fields','./preventivi-matrix-xlsx-fields.js?v=20260801a','Campi matrice XLSX non caricati.');
  await loadScript('data-preventivi-model-field-recovery','./preventivi-model-field-recovery.js?v=20260801c','Riconoscimento avanzato campi modello non caricato.');
  await loadScript('data-preventivi-model-driven-form','./preventivi-model-driven-form.js?v=20260801a','Form dinamico dei modelli non caricato.');
};
const existingExact=document.querySelector('script[data-preventivi-exact-xlsx]');
if(existingExact){loadModelModules();}
else{
  const script=document.createElement('script');script.src='./preventivi-exact-xlsx.js?v=20260801a';script.async=false;script.dataset.preventiviExactXlsx='1';script.addEventListener('load',loadModelModules,{once:true});script.addEventListener('error',()=>{console.warn('Compilatore matrice XLSX non caricato.');loadModelModules();},{once:true});document.head.appendChild(script);
}
})();
