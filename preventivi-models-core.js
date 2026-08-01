(() => {
  'use strict';
  const PV = window.HeraPreventivi;
  if (!PV) throw new Error('Preventivi core non caricato.');

  const M = window.HeraPreventiviModels = window.HeraPreventiviModels || {};
  const MODELS_KEY = 'hera_preventivi_modelli_v2';
  const DB_NAME = 'hera-preventivi-modelli';
  const STORE = 'files';
  const MODEL_COLLECTION = 'commesse/preventivi_app/modelli';
  const CONS_COLLECTION = 'commesse/preventivi_app/consuntivi';
  const SUPPORTED = ['docx','xlsx','pdf','odt','ods','html','htm','rtf','txt','json'];
  const ZIP_LIB = 'https://cdn.jsdelivr.net/npm/pizzip@3.1.8/dist/pizzip.min.js';
  const PDF_LIB = 'https://cdn.jsdelivr.net/npm/pdf-lib@1.17.1/dist/pdf-lib.min.js';

  const CATALOG = {
    numero_documento:['Numero documento','text','number'], numero_preventivo:['Numero preventivo','text','number'], numero_consuntivo:['Numero consuntivo','text','number'],
    numero_contratto:['Numero contratto','text','contractNumber'], numero_ordine:['Numero ordine','text','orderNumber'], numero_richiesta:['Numero richiesta','text','requestNumber'],
    richiedente:['Richiedente','text','requester'], cliente:['Cliente','text','clientName'], committente:['Committente','text','clientName'], data_documento:['Data documento','date','date'],
    data_inizio_lavori:['Data inizio lavori','date','startDate'], data_fine_lavori:['Data fine lavori','date','endDate'], validita_preventivo:['Validità preventivo','date','validUntil'],
    commessa:['Commessa','text','commessaName'], codice_commessa:['Codice commessa','text','commessaCode'], impianto:['Impianto','text','plantName'],
    denominazione_impianto:['Denominazione impianto','text','plantName'], id_sap:['ID SAP','text','plantSap'], indirizzo:['Indirizzo','text','workLocation'], comune:['Comune','text','city'],
    tipologia_impianto:['Tipologia impianto','text','plantType'], oggetto:['Oggetto','text','subject'], descrizione_intervento:['Descrizione intervento','textarea','description'],
    descrizione_lavori:['Descrizione lavori','textarea','description'], operatori:['Operatori','number','operators'], ore_impiegate:['Ore impiegate','number','hours'],
    mezzi_utilizzati:['Mezzi utilizzati','textarea','vehicles'], materiali_utilizzati:['Materiali utilizzati','textarea','materials'], prezziario:['Prezziario di riferimento','text','priceListNames'],
    note:['Note','textarea','notes'], totale_imponibile:['Totale imponibile','currency','totale_imponibile'], aliquota_iva:['Aliquota IVA','number','aliquota_iva'],
    iva:['IVA','currency','iva'], totale_generale:['Totale generale','currency','totale_generale'], firma_cliente:['Firma cliente','signature','firma_cliente'],
    firma_operatore:['Firma operatore','signature','firma_operatore'], timbro_cliente:['Timbro cliente','image','timbro_cliente'], lavorazioni:['Lavorazioni','repeater','lavorazioni']
  };
  const ALIASES = {
    'contratto numero':'numero_contratto','numero contratto':'numero_contratto',contratto:'numero_contratto',richiedente:'richiedente','data inizio lavori':'data_inizio_lavori',
    'inizio lavori':'data_inizio_lavori','data fine lavori':'data_fine_lavori','fine lavori':'data_fine_lavori',cliente:'cliente',committente:'committente',commessa:'commessa',
    'codice commessa':'codice_commessa',impianto:'impianto','denominazione impianto':'denominazione_impianto','id sap':'id_sap',indirizzo:'indirizzo',comune:'comune',
    oggetto:'oggetto','descrizione intervento':'descrizione_intervento','descrizione lavori':'descrizione_lavori',prezziario:'prezziario','prezziario di riferimento':'prezziario',
    note:'note',imponibile:'totale_imponibile','totale imponibile':'totale_imponibile',iva:'iva',totale:'totale_generale','totale generale':'totale_generale',
    lavorazioni:'lavorazioni','firma cliente':'firma_cliente','timbro cliente':'timbro_cliente'
  };

  M.version = '20260801b';
  M.PV = PV;
  M.constants = { MODELS_KEY, MODEL_COLLECTION, CONS_COLLECTION, SUPPORTED, ZIP_LIB, PDF_LIB };
  M.runtime = M.runtime || { db:null, observer:null, decorating:false, remoteSubscribed:false, previewDoc:null, previewType:'preventivo' };
  M.clean = (value) => String(value ?? '').trim().replace(/\s+/g, ' ');
  M.normalize = (value) => PV.normalizeText(value).replace(/[’']/g, '').replace(/[^a-z0-9]+/g, ' ').trim();
  M.slug = (value) => M.normalize(value).replace(/\s+/g, '_') || 'campo';
  M.extension = (name) => String(name || '').split('.').pop().toLowerCase();
  M.formatDateTime = (value) => { const d = new Date(value || 0); return Number.isNaN(d.getTime()) ? '—' : new Intl.DateTimeFormat('it-IT',{dateStyle:'short',timeStyle:'short'}).format(d); };
  M.canManage = () => {
    if (window.isAdmin === true || window.currentUser?.isAdmin === true) return true;
    const role = M.normalize(window.currentUserRole || window.userRole || document.body?.dataset?.role || window.currentUser?.role || '');
    return !role || /admin|amministratore|responsabile/.test(role);
  };
  M.fieldDefinition = (raw) => {
    const direct = M.slug(raw); const key = CATALOG[direct] ? direct : ALIASES[M.normalize(raw)] || direct; const item = CATALOG[key];
    return { key, label:item?.[0] || M.clean(raw).replace(/_/g,' ').replace(/\b\w/g,c=>c.toUpperCase()), type:item?.[1] || (/data/.test(key)?'date':/note|descrizione|lavor/.test(key)?'textarea':/totale|prezzo|importo/.test(key)?'currency':'text'), source:item?.[2] || '', required:false, automatic:Boolean(item?.[2]) };
  };
  M.uniqueFields = (fields) => {
    const map = new Map(); (fields || []).forEach((field) => { const base=M.fieldDefinition(field.key || field.label || 'campo'); const key=M.slug(field.key || base.key); if(!map.has(key)) map.set(key,{...base,...field,key}); }); return [...map.values()];
  };
  M.extractFields = (text, includeLabels=true) => {
    const source=String(text || ''); const map=new Map(); const patterns=[/\{\{\s*([#/^]?[^{}]+?)\s*\}\}/g,/\[\[\s*([^\[\]]+?)\s*\]\]/g];
    patterns.forEach((pattern)=>{ let match; while((match=pattern.exec(source))){ const raw=M.clean(match[1]); if(!raw || /^[#/^]/.test(raw)) continue; const field=M.fieldDefinition(raw); if(!map.has(field.key)) map.set(field.key,{...field,token:match[0]}); } });
    if(includeLabels){ const normalized=M.normalize(source.slice(0,250000)); Object.entries(ALIASES).forEach(([label,key])=>{ if(normalized.includes(M.normalize(label))&&!map.has(key)) map.set(key,{...M.fieldDefinition(key),detectedFromLabel:label}); }); }
    return [...map.values()];
  };

  M.ensureScript = (globals, src, marker) => new Promise((resolve,reject) => {
    const names=Array.isArray(globals)?globals:[globals]; const found=names.map(n=>window[n]).find(Boolean); if(found) return resolve(found);
    const existing=document.querySelector(`script[data-pvm-lib="${marker}"]`); if(existing){ existing.addEventListener('load',()=>resolve(names.map(n=>window[n]).find(Boolean)),{once:true}); existing.addEventListener('error',reject,{once:true}); return; }
    const script=document.createElement('script'); script.src=src; script.dataset.pvmLib=marker; script.onload=()=>resolve(names.map(n=>window[n]).find(Boolean)); script.onerror=()=>reject(new Error(`Libreria ${marker} non disponibile.`)); document.head.appendChild(script);
  });
  M.openDb = () => {
    if(M.runtime.db) return M.runtime.db;
    M.runtime.db=new Promise((resolve,reject)=>{ const request=indexedDB.open(DB_NAME,1); request.onupgradeneeded=()=>{ if(!request.result.objectStoreNames.contains(STORE)) request.result.createObjectStore(STORE); }; request.onsuccess=()=>resolve(request.result); request.onerror=()=>reject(request.error); });
    return M.runtime.db;
  };
  M.putBinary = async (id,file) => { const db=await M.openDb(); const buffer=await file.arrayBuffer(); return new Promise((resolve,reject)=>{ const tx=db.transaction(STORE,'readwrite'); tx.objectStore(STORE).put({buffer,name:file.name,type:file.type,size:file.size},id); tx.oncomplete=resolve; tx.onerror=()=>reject(tx.error); }); };
  M.getBinary = async (id) => { const db=await M.openDb(); return new Promise((resolve,reject)=>{ const request=db.transaction(STORE).objectStore(STORE).get(id); request.onsuccess=()=>resolve(request.result||null); request.onerror=()=>reject(request.error); }); };
  M.deleteBinary = async (id) => { const db=await M.openDb(); return new Promise((resolve,reject)=>{ const tx=db.transaction(STORE,'readwrite'); tx.objectStore(STORE).delete(id); tx.oncomplete=resolve; tx.onerror=()=>reject(tx.error); }); };

  M.analyzeFile = async (file) => {
    const ext=M.extension(file.name); if(!SUPPORTED.includes(ext)) throw new Error(`Formato .${ext || '?'} non supportato.`); const buffer=await file.arrayBuffer(); let fields=[]; const analysis={format:ext,size:file.size};
    if(['html','htm','rtf','txt','json'].includes(ext)){ const text=await file.text(); fields=M.extractFields(text,true); analysis.textPreview=text.slice(0,3000); }
    else if(ext==='xlsx'){ if(!window.XLSX) throw new Error('Lettore Excel non disponibile.'); const wb=XLSX.read(buffer,{type:'array',cellText:true}); wb.SheetNames.forEach(name=>{ const rows=XLSX.utils.sheet_to_json(wb.Sheets[name],{header:1,raw:false,defval:''}); fields.push(...M.extractFields(rows.flat().join('\n'),true)); }); analysis.sheets=wb.SheetNames; }
    else if(['docx','odt','ods'].includes(ext)){ await M.ensureScript('PizZip',ZIP_LIB,'pizzip'); const zip=new PizZip(buffer); const names=Object.keys(zip.files).filter(name=>/\.(xml|rels)$/i.test(name)); let text=''; names.slice(0,80).forEach(name=>{ const entry=zip.file(name); if(entry) text += `\n${entry.asText()}`; }); fields=M.extractFields(text.replace(/<[^>]+>/g,' '),true); analysis.entries=names.length; }
    else if(ext==='pdf'){ await M.ensureScript('PDFLib',PDF_LIB,'pdf-lib'); const pdf=await PDFLib.PDFDocument.load(buffer,{ignoreEncryption:true}); const pdfFields=pdf.getForm().getFields(); fields=pdfFields.map(field=>({...M.fieldDefinition(field.getName()),pdfFieldName:field.getName()})); analysis.pages=pdf.getPageCount(); analysis.fillableFields=pdfFields.length; }
    analysis.fields=M.uniqueFields(fields); return analysis;
  };
  M.compatibility = (format,analysis={}) => ({docx:'DOCX compilato + PDF + stampa',xlsx:'XLSX compilato + PDF + stampa',ods:'ODS compilato + PDF + stampa',odt:'ODT compilato + PDF + stampa',pdf:analysis.fillableFields?'PDF compilabile + stampa':'PDF generico + stampa',html:'HTML compilato + PDF + stampa',htm:'HTML compilato + PDF + stampa',rtf:'RTF compilato + PDF + stampa',txt:'TXT compilato + PDF + stampa',json:'JSON compilato + PDF + stampa'}[format] || 'PDF + stampa');
  M.sanitizeModel = (model) => { const copy={...model}; delete copy.analysis?.textPreview; return copy; };
  M.uploadModelFile = async (model,file) => { try { if(!window.firebase?.storage) return ''; const safe=file.name.replace(/[^a-zA-Z0-9._-]/g,'_'); const snap=await firebase.storage().ref().child(`preventivi-models/${model.id}/${safe}`).put(file,{contentType:file.type}); return await snap.ref.getDownloadURL(); } catch(error){ console.warn('Modello salvato solo sul dispositivo.',error); return ''; } };
  M.getModelBlob = async (model) => { if(model?.builtIn) return null; const local=await M.getBinary(`model:${model.id}`).catch(()=>null); if(local?.buffer) return {blob:new Blob([local.buffer],{type:local.type}),buffer:local.buffer,name:local.name,type:local.type}; if(model?.fileUrl){ const response=await fetch(model.fileUrl); if(!response.ok) throw new Error('File modello non disponibile online.'); const blob=await response.blob(); return {blob,buffer:await blob.arrayBuffer(),name:model.originalName,type:blob.type}; } throw new Error('File originale non disponibile su questo dispositivo.'); };
  M.getModel = (id) => (PV.state.models || []).find(model=>model.id===id) || null;
  M.activeModels = (type) => (PV.state.models || []).filter(model=>model.active!==false && (model.documentType==='both'||model.documentType===type));
  M.seed = () => { if((PV.state.models||[]).some(m=>m.builtIn)) return; PV.state.models=PV.state.models||[]; PV.state.models.unshift({id:'model-standard-varga',name:'Modello standard VARGA CANTIERI',documentType:'both',description:'Modello interno pronto per preventivi e consuntivi.',format:'html',originalName:'modello-standard.html',active:true,isDefault:true,builtIn:true,version:1,fields:['numero_contratto','richiedente','data_inizio_lavori','data_fine_lavori','cliente','commessa','codice_commessa','impianto','id_sap','indirizzo','comune','oggetto','descrizione_intervento','prezziario','note'].map(key=>({...M.fieldDefinition(key),required:['commessa','impianto'].includes(key)})),compatibility:'Stampa, PDF e HTML',createdAt:PV.nowIso(),updatedAt:PV.nowIso(),syncPending:true}); PV.persistLocal?.(); };

  PV.keys.models=MODELS_KEY; PV.collections.models=MODEL_COLLECTION; PV.collections.consuntivi=CONS_COLLECTION; PV.state.models=PV.state.models||[]; PV.state.editingModelId=''; PV.state.modelSearch=''; PV.state.deletions.models=PV.state.deletions.models||{};
  const load=PV.loadLocal.bind(PV); PV.loadLocal=()=>{ load(); const models=PV.readJson(MODELS_KEY,[]); PV.state.models=Array.isArray(models)?models:[]; PV.state.deletions.models=PV.state.deletions.models||{}; M.seed(); };
  const persist=PV.persistLocal.bind(PV); PV.persistLocal=()=>{ persist(); PV.writeJson(MODELS_KEY,PV.state.models||[]); };
  const sync=PV.syncPending.bind(PV); PV.syncPending=async()=>{ await sync(); if(!PV.state.firestore) return; for(const model of (PV.state.models||[]).filter(m=>m.syncPending)){ if(await PV.saveRemote(PV.collections.models,{...M.sanitizeModel(model),syncPending:false})) model.syncPending=false; } for(const id of Object.keys(PV.state.deletions.models||{})){ if(await PV.deleteRemote(PV.collections.models,id)) delete PV.state.deletions.models[id]; } PV.persistLocal(); };
  M.subscribe = (attempt=0,options={}) => { if(!PV.state.isOpen)return; if(!PV.state.firestore){ if(attempt<20)setTimeout(()=>M.subscribe(attempt+1,options),500); return; } PV.subscribeCollection(PV.collections.models,'models','models',options); };
  const connect=PV.connectFirebase.bind(PV); PV.connectFirebase=(options={})=>{ connect(options); M.subscribe(0,options); };
  M.seed();
})();
