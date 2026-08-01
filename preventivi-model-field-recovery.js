(()=>{
  'use strict';
  const M=window.HeraPreventiviModels,P=M?.PV||window.HeraPreventivi;
  if(!M||!P||M.runtime.fieldRecovery)return;
  M.runtime.fieldRecovery=true;

  const VERSION='20260801c';
  const C=v=>String(v??'').trim().replace(/\s+/g,' ');
  const N=v=>C(v).normalize('NFD').replace(/[\u0300-\u036f]/g,'').toLowerCase().replace(/[^a-z0-9%]+/g,' ').trim();
  const PLACEHOLDER=/^(?:[-_.:/\\ ]+|x+|n\/?a|da compilare|da inserire|inserire|compilare|selezionare)$/i;

  const DEFINITIONS=[
    {labels:['nome impianto','denominazione impianto'],key:'denominazione_impianto',label:'Nome impianto',type:'text',source:'plantName',required:true,automatic:true,plant:true},
    {labels:['id sap','codice sap'],key:'id_sap',label:'ID SAP',type:'text',source:'plantSap',automatic:true,plant:true},
    {labels:['comune'],key:'comune',label:'Comune',type:'text',source:'city',automatic:true,plant:true},
    {labels:['indirizzo','luogo lavori','luogo dei lavori','cantiere luogo dei lavori'],key:'indirizzo',label:'Indirizzo / luogo lavori',type:'text',source:'workLocation',automatic:true,plant:true},
    {labels:['tipologia impianto'],key:'tipologia_impianto',label:'Tipologia impianto',type:'text',source:'plantType',automatic:true,plant:true},
    {labels:['codice commessa'],key:'codice_commessa',label:'Codice commessa',type:'text',source:'commessaCode',automatic:true,commessa:true},
    {labels:['commessa'],key:'commessa',label:'Commessa',type:'text',source:'commessaName',automatic:true,commessa:true},
    {labels:['odl','ordine di lavoro','numero ordine','n ordine'],key:'numero_ordine',label:'ODL / Numero ordine',type:'text',source:'orderNumber',required:false},
    {labels:['numero richiesta','n richiesta'],key:'numero_richiesta',label:'Numero richiesta',type:'text',source:'requestNumber'},
    {labels:['numero contratto','contratto n','contratto numero'],key:'numero_contratto',label:'Numero contratto',type:'text',source:'contractNumber'},
    {labels:['oggetto','oggetto del preventivo'],key:'oggetto',label:'Oggetto',type:'text',source:'subject',required:true},
    {labels:['descrizione intervento'],key:'descrizione_intervento',label:'Descrizione intervento',type:'textarea',source:'description',required:true},
    {labels:['descrizione lavori','attivita eseguite','attivita svolta'],key:'descrizione_lavori',label:'Descrizione lavori / attività',type:'textarea',source:'description'},
    {labels:['richiedente intervento'],key:'richiedente_intervento',label:'Richiedente intervento',type:'text',source:'requester'},
    {labels:['richiedente'],key:'richiedente',label:'Richiedente',type:'text',source:'requester'},
    {labels:['ditta esecutrice se diversa da avola','ditta esecutrice'],key:'ditta_esecutrice',label:'Ditta esecutrice se diversa da Avola',type:'text',source:'issuerName'},
    {labels:['cliente ragione sociale','ragione sociale','cliente'],key:'cliente',label:'Cliente / Ragione sociale',type:'text',source:'clientName'},
    {labels:['committente','spett le','spettabile'],key:'committente',label:'Committente',type:'text',source:'clientName'},
    {labels:['partita iva codice fiscale cliente','partita iva','codice fiscale cliente'],key:'partita_iva',label:'Partita IVA / Codice fiscale',type:'text',source:'clientTaxCode'},
    {labels:['intestazione azienda fornitore','fornitore'],key:'intestazione_azienda',label:'Intestazione azienda / fornitore',type:'text',source:'issuerName'},
    {labels:['data richiesta'],key:'data_richiesta',label:'Data richiesta',type:'date'},
    {labels:['data esecuzione'],key:'data_esecuzione',label:'Data esecuzione',type:'date'},
    {labels:['data inizio lavori','inizio lavori'],key:'data_inizio_lavori',label:'Data inizio lavori',type:'date',source:'startDate'},
    {labels:['data fine lavori','fine lavori'],key:'data_fine_lavori',label:'Data fine lavori',type:'date',source:'endDate'},
    {labels:['data documento','data preventivo','data consuntivo'],key:'data_documento',label:'Data documento',type:'date',source:'date',calculated:true},
    {labels:['validita preventivo','validita giorni'],key:'validita_preventivo',label:'Validità preventivo',type:'text',source:'validityDays'},
    {labels:['competenza bologna ovest'],key:'competenza_bologna_ovest',label:'Competenza Bologna Ovest',type:'checkbox'},
    {labels:['competenza bologna est'],key:'competenza_bologna_est',label:'Competenza Bologna Est',type:'checkbox'},
    {labels:['operatori','numero operatori'],key:'operatori',label:'Operatori',type:'number',source:'operators'},
    {labels:['ore impiegate','ore lavorate'],key:'ore_impiegate',label:'Ore impiegate',type:'number',source:'hours'},
    {labels:['mezzi utilizzati'],key:'mezzi_utilizzati',label:'Mezzi utilizzati',type:'textarea',source:'vehicles'},
    {labels:['materiali utilizzati'],key:'materiali_utilizzati',label:'Materiali utilizzati',type:'textarea',source:'materials'},
    {labels:['note operative','note condizioni','note'],key:'note',label:'Note',type:'textarea',source:'notes'},
    {labels:['firma cliente','firma committente'],key:'firma_cliente',label:'Firma cliente',type:'signature'},
    {labels:['firma operatore','firma impresa'],key:'firma_operatore',label:'Firma operatore / impresa',type:'signature'},
    {labels:['timbro cliente'],key:'timbro_cliente',label:'Timbro cliente',type:'image'}
  ];
  const ECONOMIC_LABELS=['cod prest est','codice prestazione','codice lavorazione','testo breve','testo esteso','unita di misura','quantita','prezzo capitolato','prezzo unitario','ribasso si no','prezzo netto ribassato','importo prestazione','note descrizioni attivita svolta'];
  const TOTAL_LABELS=['totale intervento','totale preventivo','totale consuntivo','totale generale','totale imponibile'];

  const definitionField=definition=>({key:definition.key,label:definition.label,type:definition.type||'text',source:definition.source||'',required:Boolean(definition.required),automatic:Boolean(definition.automatic),detectedFromLabel:true,detectionVersion:VERSION});
  const matchDefinition=value=>{
    const normalized=N(value).replace(/\bn\b$/,'').trim();
    let best=null;
    DEFINITIONS.forEach(definition=>definition.labels.forEach(label=>{
      const normalizedLabel=N(label);
      if(normalized===normalizedLabel||normalized.startsWith(`${normalizedLabel} `)){
        if(!best||normalizedLabel.length>best.label.length)best={definition,label:normalizedLabel};
      }
    }));
    return best;
  };
  const inlineTail=(raw,label)=>{
    const cleaned=C(raw),normalized=N(cleaned),index=normalized.indexOf(label);
    if(index!==0)return'';
    const colon=cleaned.indexOf(':');
    if(colon>=0)return C(cleaned.slice(colon+1));
    const words=label.split(' ').length;
    return C(cleaned.split(/\s+/).slice(words).join(' '));
  };
  const meaningful=value=>{const text=C(value);return Boolean(text)&&!PLACEHOLDER.test(text)&&!/^(?:\{\{|\[\[)/.test(text)};
  const exactKnownLabel=value=>Boolean(matchDefinition(value))||ECONOMIC_LABELS.some(label=>N(value)===N(label))||TOTAL_LABELS.some(label=>N(value)===N(label));

  function detectRows(rows){
    const map=new Map();
    let economic=false,plant=false;
    const add=field=>{if(!map.has(field.key))map.set(field.key,field)};
    (rows||[]).forEach(row=>{
      const cells=Array.isArray(row)?row:[];
      cells.forEach((raw,column)=>{
        const value=C(raw),normalized=N(value);
        if(!value)return;
        if(ECONOMIC_LABELS.some(label=>normalized===N(label))){economic=true;return;}
        if(TOTAL_LABELS.some(label=>normalized===N(label))){add({...M.fieldDefinition('totale_imponibile'),calculated:true,detectedFromLabel:value,detectionVersion:VERSION});return;}
        const matched=matchDefinition(value);if(!matched)return;
        const {definition,label}=matched;
        const tail=inlineTail(value,label);
        let fixed=meaningful(tail);
        if(!fixed){
          for(let index=column+1;index<Math.min(cells.length,column+8);index+=1){
            const candidate=C(cells[index]);
            if(candidate&&exactKnownLabel(candidate))break;
            if(meaningful(candidate)){fixed=true;break;}
          }
        }
        if(definition.calculated){add({...definitionField(definition),calculated:true});return;}
        if(definition.type==='checkbox'||!fixed){add(definitionField(definition));if(definition.plant)plant=true;}
      });
    });
    if(plant){
      add({key:'commessa',label:'Commessa',type:'text',source:'commessaName',required:true,automatic:true,helperOnly:true,detectedFromLabel:'collegamento impianto',detectionVersion:VERSION});
    }
    if(economic)add({...M.fieldDefinition('lavorazioni'),type:'repeater',source:'lavorazioni',automatic:true,calculated:true,detectedFromLabel:'tabella lavorazioni',detectionVersion:VERSION});
    return [...map.values()];
  }

  const decodeXml=text=>{const node=new DOMParser().parseFromString(`<div>${text}</div>`,'text/html');return node.body?.textContent||''};
  function rowsFromXmlText(text,format){
    let source=String(text||'');
    if(format==='docx')source=source.replace(/<\/w:tc>/gi,'\t').replace(/<\/w:tr>/gi,'\n').replace(/<\/w:p>/gi,'\n');
    else source=source.replace(/<\/table:table-cell>/gi,'\t').replace(/<\/table:table-row>/gi,'\n').replace(/<\/text:p>/gi,'\n');
    source=decodeXml(source.replace(/<[^>]+>/g,' '));
    return source.split(/\r?\n/).map(line=>line.split(/\t/));
  }
  async function rowsFromStored(stored,format){
    if(format==='xlsx'){
      if(!window.XLSX)throw new Error('Lettore Excel non disponibile.');
      const workbook=XLSX.read(stored.buffer,{type:'array',cellText:true,cellDates:false});
      const rows=[];workbook.SheetNames.forEach(name=>rows.push(...XLSX.utils.sheet_to_json(workbook.Sheets[name],{header:1,raw:false,defval:''}),[]));return rows;
    }
    if(['docx','odt','ods'].includes(format)){
      await M.ensureScript('PizZip',M.constants.ZIP_LIB,'pizzip');
      const zip=new PizZip(stored.buffer),paths=format==='docx'?['word/document.xml']:['content.xml'],rows=[];
      paths.forEach(path=>{const entry=zip.file(path);if(entry)rows.push(...rowsFromXmlText(entry.asText(),format));});return rows;
    }
    if(['html','htm','rtf','txt','json'].includes(format)){
      const text=await stored.blob.text();
      return decodeXml(text.replace(/<\/?(?:tr|p|div|br|li|h\d)[^>]*>/gi,'\n').replace(/<\/?(?:td|th)[^>]*>/gi,'\t')).split(/\r?\n/).map(line=>line.split(/\t/));
    }
    return[];
  }
  const mergeFields=(existing,recovered)=>{
    const map=new Map();
    (existing||[]).forEach(field=>{const key=M.slug(field.key||field.label||'campo');map.set(key,{...M.fieldDefinition(key),...field,key})});
    (recovered||[]).forEach(field=>{const key=M.slug(field.key||field.label||'campo'),current=map.get(key);map.set(key,current?{...field,...current,key}:{...field,key})});
    return [...map.values()];
  };
  const signature=fields=>JSON.stringify((fields||[]).map(field=>[field.key,field.label,field.type,Boolean(field.required),field.source]).sort((a,b)=>String(a[0]).localeCompare(String(b[0]))));

  async function recoverStored(model,stored){
    const format=String(model.format||M.extension(model.originalName)).toLowerCase();
    if(format==='pdf')return model.fields||[];
    const rows=await rowsFromStored(stored,format),recovered=detectRows(rows);
    return mergeFields(model.fields,recovered);
  }
  async function recoverModel(model,form=null){
    if(!model||model.builtIn||model.fieldDetectionVersion===VERSION||M.runtime.recoveringModel===model.id)return false;
    M.runtime.recoveringModel=model.id;
    try{
      const stored=await M.getModelBlob(model),before=signature(model.fields),fields=await recoverStored(model,stored),after=signature(fields);
      model.fieldDetectionVersion=VERSION;
      model.analysis={...(model.analysis||{}),fieldDetectionVersion:VERSION,recoveredFields:fields.length};
      if(before!==after){model.fields=fields;model.version=Math.max(1,Number(model.version)||1)+1;model.updatedAt=P.nowIso();model.updatedBy=P.currentUser();model.syncPending=true;}
      P.persistLocal?.();if(before!==after)P.scheduleSync?.();
      if(form&&before!==after){form.dataset.pvdSignature='';const select=form.querySelector('[data-pvm-model-select]');select?.dispatchEvent(new Event('change',{bubbles:true}));P.setFeedback?.(`Modello “${model.name}” aggiornato: riconosciuti ${fields.length} campi.`, 'success');}
      return before!==after;
    }catch(error){console.warn(`Riconoscimento campi non riuscito per ${model.name}:`,error);return false}
    finally{if(M.runtime.recoveringModel===model.id)M.runtime.recoveringModel=''}
  }

  const originalAnalyze=M.analyzeFile?.bind(M);
  if(originalAnalyze)M.analyzeFile=async file=>{
    const analysis=await originalAnalyze(file),format=M.extension(file.name);
    if(format==='pdf')return analysis;
    try{
      const buffer=await file.arrayBuffer(),stored={buffer,blob:new Blob([buffer],{type:file.type}),name:file.name,type:file.type};
      analysis.fields=await recoverStored({format,originalName:file.name,fields:analysis.fields||[]},stored);
      analysis.fieldDetectionVersion=VERSION;
    }catch(error){console.warn('Riconoscimento avanzato campi non riuscito:',error)}
    return analysis;
  };

  const originalPayload=M.modelPayload?.bind(M);
  if(originalPayload)M.modelPayload=form=>{
    const payload=originalPayload(form);
    form?.querySelectorAll('input[type="checkbox"][data-pvm-field]').forEach(input=>{payload.modelFields[input.dataset.pvmField]=input.checked?'X':''});
    return payload;
  };

  function selectedForm(target){return target?.closest?.('[data-pv-quote-form],[data-cons-form]')||P.page()?.querySelector('[data-pv-quote-form],[data-cons-form]')}
  function queueRecovery(form){
    if(!form)return;const model=M.selectedModel?.(form);if(!model||model.builtIn)return;
    clearTimeout(form._pvmRecoveryTimer);form._pvmRecoveryTimer=setTimeout(()=>recoverModel(model,form),80);
  }
  document.addEventListener('change',event=>{if(event.target.matches?.('[data-pvm-model-select]'))queueRecovery(selectedForm(event.target))},true);
  document.addEventListener('click',event=>{if(event.target.closest?.('[data-pv-action="new-quote"],[data-pv-action="edit-quote"],[data-cons-new],[data-cons-edit]'))setTimeout(()=>queueRecovery(selectedForm()),250)},true);

  let queued=false;
  const scan=()=>{if(queued)return;queued=true;requestAnimationFrame(()=>{queued=false;P.page()?.querySelectorAll('[data-pv-quote-form],[data-cons-form]').forEach(queueRecovery)})};
  const start=()=>{new MutationObserver(scan).observe(document.body,{childList:true,subtree:true});scan();[700,1800,4000].forEach(ms=>setTimeout(scan,ms))};
  document.readyState==='loading'?document.addEventListener('DOMContentLoaded',start,{once:true}):start();

  window.HeraPreventiviModelFieldRecovery={version:VERSION,detectRows,recoverModel};
})();
