(()=>{
  'use strict';
  const M=window.HeraPreventiviModels,P=M?.PV||window.HeraPreventivi;
  if(!M||!P||M.runtime.modelDrivenForm)return;
  M.runtime.modelDrivenForm=true;

  const C=v=>String(v??'').trim();
  const N=v=>C(v).normalize('NFD').replace(/[\u0300-\u036f]/g,'').toLowerCase().replace(/[^a-z0-9]+/g,' ').trim();
  const E=P.escapeHtml;
  const COMMESSA=new Set(['commessa','codice_commessa']);
  const PLANT=new Set(['impianto','denominazione_impianto','id_sap','indirizzo','comune','tipologia_impianto']);
  const CALCULATED=new Set(['numero_documento','numero_preventivo','numero_consuntivo','data_documento','totale_imponibile','iva','totale_generale','lavorazioni','prezziario']);
  const VALUE_MAP={numero_contratto:'contractNumber',numero_ordine:'orderNumber',numero_richiesta:'requestNumber',richiedente:'requester',cliente:'clientName',committente:'clientName',data_inizio_lavori:'startDate',data_fine_lavori:'endDate',commessa:'commessaName',codice_commessa:'commessaCode',impianto:'plantName',denominazione_impianto:'plantName',id_sap:'plantSap',indirizzo:'workLocation',comune:'city',tipologia_impianto:'plantType',oggetto:'subject',descrizione_intervento:'description',descrizione_lavori:'description',operatori:'operators',ore_impiegate:'hours',mezzi_utilizzati:'vehicles',materiali_utilizzati:'materials',note:'notes',aliquota_iva:'vatRate'};
  const ECONOMIC=/lavor|prezz|quant|importo|totale|ribasso|unita|misura|cod.*prest|testo.*breve|testo.*esteso/i;
  const today=()=>new Date().toISOString().slice(0,10);
  const typeOf=form=>form?.matches('[data-cons-form]')?'consuntivo':'preventivo';
  const selectedModel=form=>M.selectedModel?.(form)||null;
  const currentDoc=type=>type==='consuntivo'?(P.state.editingConsuntivoId&&P.state.editingConsuntivoId!=='new'?(P.state.consuntivi||[]).find(x=>x.id===P.state.editingConsuntivoId)||{}:{}):(P.state.editingQuoteId&&P.state.editingQuoteId!=='new'?P.getQuote(P.state.editingQuoteId)||{}:{});
  const registry=()=>{try{return window.HeraPreventiviRegistry?.registry?.()||{commesse:[],plants:[]}}catch(_){return{commesse:[],plants:[]}}};
  const fields=model=>{const map=new Map();(model?.fields||[]).forEach(f=>{const key=M.slug(f.key||f.label||'campo');if(!map.has(key))map.set(key,{...f,key});});return[...map.values()]};
  const keysOf=model=>new Set(fields(model).map(f=>f.key));
  const isEconomicField=f=>ECONOMIC.test(N(`${f.key} ${f.label||''}`));
  const needsEconomic=model=>fields(model).some(isEconomicField);
  const isDepGas=model=>String(model?.format||'').toLowerCase()==='xlsx'&&N(`${model?.name||''} ${model?.originalName||''} ${model?.description||''}`).includes('depurazione')&&N(`${model?.name||''} ${model?.originalName||''} ${model?.description||''}`).includes('gas');
  const valueOf=(doc,key)=>doc?.modelFields?.[key]??doc?.[VALUE_MAP[key]]??'';
  const setNamed=(form,name,value,onlyEmpty=false)=>{const el=form.elements?.namedItem(name);if(!el||!('value'in el))return;if(!onlyEmpty||!C(el.value))el.value=value??''};
  const modelField=(form,key)=>form.querySelector(`[data-pvm-field="${key}"]`);
  const dataValue=(form,...keys)=>{for(const key of keys){const el=modelField(form,key);if(C(el?.value))return C(el.value)}return''};
  const setModelField=(form,key,value,onlyEmpty=false)=>{const el=modelField(form,key);if(el&&(!onlyEmpty||!C(el.value)))el.value=value??''};
  const liveValue=(form,name,fallback='')=>{const el=form.elements?.namedItem(name);return el&&'value'in el?el.value:fallback};
  function liveDoc(form,type,base=currentDoc(type)){
    const modelFields={...(base.modelFields||{})};
    form.querySelectorAll('[data-pvm-field]').forEach(el=>{modelFields[el.dataset.pvmField]=el.type==='checkbox'?(el.checked?'X':''):el.value});
    const commessaId=form.querySelector('[data-pvd-backing="commessaId"]')?.value??base.commessaId??'';
    const plantId=form.querySelector('[data-pvd-backing="plantId"]')?.value??base.plantId??'';
    const plantSearch=form.querySelector('[data-pvd-plant-search]')?.value||base.plantName||'';
    return {...base,modelFields,commessaId,plantId,plantName:plantSearch.split(' — ID SAP ')[0],date:liveValue(form,'date',base.date),issuerName:liveValue(form,'issuerName',base.issuerName),clientName:liveValue(form,'clientName',base.clientName),clientTaxCode:liveValue(form,'clientTaxCode',base.clientTaxCode),workLocation:liveValue(form,'workLocation',base.workLocation),subject:liveValue(form,'subject',base.subject),notes:liveValue(form,'notes',base.notes),validityDays:liveValue(form,'validityDays',base.validityDays),vatRate:liveValue(form,'vatRate',base.vatRate),contractNumber:liveValue(form,'contractNumber',base.contractNumber),requester:liveValue(form,'requester',base.requester),startDate:liveValue(form,'startDate',base.startDate),endDate:liveValue(form,'endDate',base.endDate),description:liveValue(form,'description',base.description),hours:liveValue(form,'hours',base.hours),operators:liveValue(form,'operators',base.operators),vehicles:liveValue(form,'vehicles',base.vehicles),materials:liveValue(form,'materials',base.materials),compiler:liveValue(form,'compiler',base.compiler),commessaCode:liveValue(form,'commessaCode',base.commessaCode),plantSap:liveValue(form,'plantSap',base.plantSap),city:liveValue(form,'city',base.city),plantType:liveValue(form,'plantType',base.plantType)};
  }

  function backing(form,name){
    let hidden=form.querySelector(`input[type="hidden"][data-pvd-backing="${name}"]`);
    if(hidden)return hidden;
    const original=[...form.elements].find(el=>el.name===name&&!el.dataset.pvdBacking);
    if(original){original.dataset.pvdOriginalName=name;original.removeAttribute('name');}
    hidden=document.createElement('input');hidden.type='hidden';hidden.name=name;hidden.dataset.pvdBacking=name;form.appendChild(hidden);return hidden;
  }
  function restoreBackings(form){form.querySelectorAll('[data-pvd-backing]').forEach(el=>el.remove());form.querySelectorAll('[data-pvd-original-name]').forEach(el=>{el.name=el.dataset.pvdOriginalName;delete el.dataset.pvdOriginalName;});}
  function suspendRequired(root){root?.querySelectorAll('[required]').forEach(el=>{if(!el.dataset.pvdWasRequired){el.dataset.pvdWasRequired='1';el.required=false;}})}
  function restoreRequired(root){root?.querySelectorAll('[data-pvd-was-required]').forEach(el=>{el.required=true;delete el.dataset.pvdWasRequired;})}
  function hide(el,on){if(!el)return;if(on){if(!el.dataset.pvdHidden)el.dataset.pvdHidden='1';el.hidden=true}else if(el.dataset.pvdHidden){el.hidden=false;delete el.dataset.pvdHidden}}

  function fieldControl(field,doc){
    const key=field.key,label=field.label||key.replace(/_/g,' ').replace(/\b\w/g,c=>c.toUpperCase()),required=field.required?'required':'',value=valueOf(doc,key),help=field.automatic?'<small class="pv-muted">Compilazione automatica quando il dato è disponibile.</small>':'';
    if(field.type==='textarea')return `<label class="pv-label pv-span-2"><span class="${field.required?'pv-required':''}">${E(label)}</span><textarea data-pvm-field="${E(key)}" ${required}>${E(value)}</textarea>${help}</label>`;
    if(field.type==='date')return `<label class="pv-label"><span class="${field.required?'pv-required':''}">${E(label)}</span><input type="date" data-pvm-field="${E(key)}" value="${E(value)}" ${required}>${help}</label>`;
    if(['number','currency'].includes(field.type))return `<label class="pv-label"><span class="${field.required?'pv-required':''}">${E(label)}</span><input type="number" step="${field.type==='currency'?'0.01':'any'}" data-pvm-field="${E(key)}" value="${E(value)}" ${required}>${help}</label>`;
    const inputType=['email','tel'].includes(field.type)?field.type:'text',placeholder=field.type==='signature'?'Nome del firmatario':field.type==='image'?'Link o riferimento immagine':'';
    return `<label class="pv-label"><span class="${field.required?'pv-required':''}">${E(label)}</span><input type="${inputType}" data-pvm-field="${E(key)}" value="${E(value)}" placeholder="${E(placeholder)}" ${required}>${help}</label>`;
  }
  function commessaMarkup(model,doc,r){
    const keys=keysOf(model),required=fields(model).some(f=>COMMESSA.has(f.key)&&f.required),selected=doc.commessaId||'',c=r.commesse.find(x=>x.id===selected);
    const empty=r.commesse.length?'Seleziona commessa':'Commesse in sincronizzazione…';
    return `<label class="pv-label"><span class="${required?'pv-required':''}">Commessa</span><select data-pvd-commessa ${required?'required':''}><option value="">${empty}</option>${r.commesse.map(x=>`<option value="${E(x.id)}" ${x.id===selected?'selected':''}>${E(x.name)}${x.code?` — ${E(x.code)}`:''}</option>`).join('')}</select></label>${keys.has('commessa')?`<input type="hidden" data-pvm-field="commessa" value="${E(c?.name||doc.commessaName||doc.modelFields?.commessa||'')}" ${required?'required':''}>`:''}${keys.has('codice_commessa')?`<label class="pv-label"><span>Codice commessa</span><input data-pvm-field="codice_commessa" value="${E(c?.code||doc.commessaCode||doc.modelFields?.codice_commessa||'')}" readonly></label>`:''}`;
  }
  function plantMarkup(model,doc,r){
    const keys=keysOf(model),required=fields(model).some(f=>PLANT.has(f.key)&&f.required),selected=doc.plantId||'',p=r.plants.find(x=>x.id===selected),outputs=[];
    if(keys.has('id_sap'))outputs.push(`<label class="pv-label"><span>ID SAP</span><input data-pvm-field="id_sap" value="${E(p?.sap||doc.plantSap||doc.modelFields?.id_sap||'')}" ${required?'required':''} readonly></label>`);
    if(keys.has('comune'))outputs.push(`<label class="pv-label"><span>Comune</span><input data-pvm-field="comune" value="${E(p?.city||doc.city||doc.modelFields?.comune||'')}" readonly></label>`);
    if(keys.has('indirizzo'))outputs.push(`<label class="pv-label pv-span-2"><span>Indirizzo</span><input data-pvm-field="indirizzo" value="${E(p?.address||doc.workLocation||doc.modelFields?.indirizzo||'')}" readonly></label>`);
    if(keys.has('tipologia_impianto'))outputs.push(`<label class="pv-label"><span>Tipologia impianto</span><input data-pvm-field="tipologia_impianto" value="${E(p?.type||doc.plantType||doc.modelFields?.tipologia_impianto||'')}" readonly></label>`);
    return `<label class="pv-label"><span class="${required?'pv-required':''}">Impianto</span><div class="pvd-plant-wrap"><input type="search" data-pvd-plant-search autocomplete="off" value="${E(p?`${p.name} — ID SAP ${p.sap||'—'}`:doc.plantName||doc.modelFields?.impianto||doc.modelFields?.denominazione_impianto||'')}" placeholder="Cerca nome impianto o ID SAP" ${required?'required':''}><div class="pvd-results hidden" data-pvd-results></div></div></label>${keys.has('impianto')?`<input type="hidden" data-pvm-field="impianto" value="${E(p?.name||doc.plantName||doc.modelFields?.impianto||'')}" ${required?'required':''}>`:''}${keys.has('denominazione_impianto')?`<input type="hidden" data-pvm-field="denominazione_impianto" value="${E(p?.name||doc.plantName||doc.modelFields?.denominazione_impianto||'')}" ${required?'required':''}>`:''}${outputs.join('')}`;
  }
  function renderSection(form,model,doc,r){
    form.querySelector('[data-pvd-section]')?.remove();
    const all=fields(model),keys=keysOf(model),ordinary=all.filter(f=>!CALCULATED.has(f.key)&&!COMMESSA.has(f.key)&&!PLANT.has(f.key)&&!isEconomicField(f));
    let controls='';
    if([...keys].some(k=>COMMESSA.has(k)))controls+=commessaMarkup(model,doc,r);
    if([...keys].some(k=>PLANT.has(k)))controls+=plantMarkup(model,doc,r);
    controls+=ordinary.map(f=>fieldControl(f,doc)).join('');
    const message=all.length?(!controls?'<p class="pv-muted">Tutti i dati del modello vengono compilati automaticamente oppure tramite le lavorazioni.</p>':''):'<p class="pv-feedback" data-type="warning">Nel modello non sono stati riconosciuti campi compilabili. Apri “Gestisci modelli” per verificare i campi riconosciuti.</p>';
    const html=`<section class="pv-form-card pvd-section" data-pvd-section><div class="pv-section-head"><div><h3>Dati richiesti dal modello</h3><p class="pv-muted">Sono mostrati esclusivamente i dati utilizzati da <strong>${E(model.name||'questo modello')}</strong>.</p></div></div><div class="pv-form-grid">${controls}</div>${message}</section>`;
    form.querySelector('[data-pvm-document-section]')?.insertAdjacentHTML('afterend',html);
  }
  function standardHeader(form){const modelSection=form.querySelector('[data-pvm-document-section]');let el=modelSection?.previousElementSibling;while(el&&!el.classList.contains('pv-form-card'))el=el.previousElementSibling;return el}
  function economicSections(form,type,show){
    if(type==='preventivo'){hide(form.querySelector('[data-pv-price-list-choices]')?.closest('.pv-form-card'),!show);hide(form.querySelector('[data-pv-lines]')?.closest('.pv-form-card'),!show);hide(form.querySelector('[data-pv-summary]'),!show);}
    else{hide(form.querySelector('.pvm-cons-price-section'),!show);hide(form.querySelector('[data-cons-lines]')?.closest('.pv-form-card'),!show);}
  }
  function syncSelection(form){
    const r=registry(),cid=backing(form,'commessaId').value,pid=backing(form,'plantId').value,c=r.commesse.find(x=>x.id===cid),p=r.plants.find(x=>x.id===pid);
    if(c){setModelField(form,'commessa',c.name||'');setModelField(form,'codice_commessa',c.code||'');setNamed(form,'commessaCode',c.code||'');}
    if(p){setModelField(form,'impianto',p.name||'');setModelField(form,'denominazione_impianto',p.name||'');setModelField(form,'id_sap',p.sap||'');setModelField(form,'comune',p.city||'');setModelField(form,'indirizzo',p.address||'');setModelField(form,'tipologia_impianto',p.type||'');setNamed(form,'plantSap',p.sap||'');setNamed(form,'city',p.city||'');setNamed(form,'plantType',p.type||'');}
    return{c,p};
  }
  function syncTechnical(form){
    if(!form.classList.contains('pvd-active'))return;
    const type=typeOf(form),model=selectedModel(form),{c,p}=syncSelection(form),fallback='Non richiesto dal modello';
    setNamed(form,'date',dataValue(form,'data_documento','data_richiesta')||form.elements.namedItem('date')?.value||today());
    if(type==='preventivo'){
      setNamed(form,'issuerName',dataValue(form,'intestazione_azienda','ditta_esecutrice')||form.elements.namedItem('issuerName')?.value||'Avola Società Cooperativa');
      setNamed(form,'clientName',dataValue(form,'cliente','committente')||c?.client||form.elements.namedItem('clientName')?.value||fallback);
      setNamed(form,'clientTaxCode',dataValue(form,'partita_iva','codice_fiscale_cliente')||form.elements.namedItem('clientTaxCode')?.value||'');
      setNamed(form,'workLocation',dataValue(form,'indirizzo')||[p?.address,p?.city].filter(Boolean).join(', ')||p?.name||form.elements.namedItem('workLocation')?.value||fallback);
      setNamed(form,'subject',dataValue(form,'oggetto','descrizione_intervento','descrizione_lavori')||form.elements.namedItem('subject')?.value||model?.name||fallback);
      setNamed(form,'notes',dataValue(form,'note')||form.elements.namedItem('notes')?.value||'');
      setNamed(form,'vatRate',dataValue(form,'aliquota_iva')||form.elements.namedItem('vatRate')?.value||'0');
      setNamed(form,'validityDays',form.elements.namedItem('validityDays')?.value||'0');
    }else{
      setNamed(form,'contractNumber',dataValue(form,'numero_contratto')||c?.contract||form.elements.namedItem('contractNumber')?.value||fallback);
      setNamed(form,'requester',dataValue(form,'richiedente','richiedente_intervento')||c?.requester||form.elements.namedItem('requester')?.value||fallback);
      setNamed(form,'startDate',dataValue(form,'data_inizio_lavori','data_richiesta')||form.elements.namedItem('startDate')?.value||today());
      setNamed(form,'endDate',dataValue(form,'data_fine_lavori','data_esecuzione')||form.elements.namedItem('endDate')?.value||today());
      setNamed(form,'clientName',dataValue(form,'cliente','committente')||c?.client||form.elements.namedItem('clientName')?.value||fallback);
      setNamed(form,'workLocation',dataValue(form,'indirizzo')||[p?.address,p?.city].filter(Boolean).join(', ')||p?.name||form.elements.namedItem('workLocation')?.value||fallback);
      setNamed(form,'subject',dataValue(form,'oggetto','descrizione_intervento','descrizione_lavori')||form.elements.namedItem('subject')?.value||model?.name||fallback);
      setNamed(form,'description',dataValue(form,'descrizione_intervento','descrizione_lavori','oggetto')||form.elements.namedItem('description')?.value||fallback);
      setNamed(form,'hours',dataValue(form,'ore_impiegate')||form.elements.namedItem('hours')?.value||'');
      setNamed(form,'operators',dataValue(form,'operatori')||form.elements.namedItem('operators')?.value||'');
      setNamed(form,'vehicles',dataValue(form,'mezzi_utilizzati')||form.elements.namedItem('vehicles')?.value||'');
      setNamed(form,'materials',dataValue(form,'materiali_utilizzati')||form.elements.namedItem('materials')?.value||'');
      setNamed(form,'compiler',form.elements.namedItem('compiler')?.value||P.currentUser?.()?.displayName||'Operatore');
      setNamed(form,'notes',dataValue(form,'note')||form.elements.namedItem('notes')?.value||'');
      setNamed(form,'vatRate',dataValue(form,'aliquota_iva')||form.elements.namedItem('vatRate')?.value||'0');
    }
  }
  function restore(form,doc=currentDoc(typeOf(form))){
    if(!form?.classList.contains('pvd-active'))return;
    form.classList.remove('pvd-active');form.querySelector('[data-pvd-section]')?.remove();form.querySelectorAll('[data-pvd-hidden]').forEach(el=>hide(el,false));
    const header=standardHeader(form);restoreRequired(header);restoreBackings(form);
    const dyn=form.querySelector('[data-pvm-dynamic-fields]');if(dyn){delete dyn.dataset.pvdOwned;M.renderDynamic?.(form,doc);}
    delete form.dataset.pvdSignature;
  }
  function apply(form){
    if(!form)return;
    const type=typeOf(form),model=selectedModel(form),doc=liveDoc(form,type);
    if(!model)return restore(form,doc);
    if(type==='preventivo'&&isDepGas(model)){restore(form,doc);return;}
    const r=registry(),sig=[model.id,model.version,fields(model).map(f=>`${f.key}:${f.required?'1':'0'}:${f.type}`).join(','),r.commesse.map(x=>x.id).join(','),r.plants.map(x=>x.id).join(',')].join('|');
    if(form.dataset.pvdSignature===sig){syncTechnical(form);return;}
    const header=standardHeader(form);hide(header,true);suspendRequired(header);form.classList.add('pvd-active');
    backing(form,'commessaId').value=doc.commessaId||'';backing(form,'plantId').value=doc.plantId||'';
    const dyn=form.querySelector('[data-pvm-dynamic-fields]');if(dyn){dyn.innerHTML='<p class="pv-muted">I campi vengono mostrati nella sezione dedicata al modello.</p>';dyn.dataset.pvdOwned='1';}
    economicSections(form,type,needsEconomic(model));renderSection(form,model,doc,r);form.dataset.pvdSignature=sig;syncTechnical(form);
  }
  function plantResults(input){
    const form=input.closest('form'),box=form.querySelector('[data-pvd-results]'),r=registry(),cid=backing(form,'commessaId').value,c=r.commesse.find(x=>x.id===cid),targets=[cid,c?.code,c?.name].map(N).filter(Boolean),q=N(input.value);
    const plants=r.plants.filter(p=>{const pc=N(p.commessaId),ok=!cid||!pc||targets.some(t=>t===pc||t.includes(pc)||pc.includes(t));return ok&&(!q||N(`${p.name} ${p.sap} ${p.city}`).includes(q))}).slice(0,80);
    box.innerHTML=plants.length?plants.map(p=>`<button type="button" class="pvd-result" data-pvd-plant="${E(p.id)}" data-pvd-cid="${E(p.commessaId||'')}"><strong>${E(p.name)}</strong><span>ID SAP ${E(p.sap||'—')}</span></button>`).join(''):'<p class="pv-muted">Nessun impianto trovato. Cerca per nome oppure ID SAP.</p>';box.classList.remove('hidden');
  }
  function choosePlant(form,id,cid=''){
    const r=registry(),p=r.plants.find(x=>x.id===id&&(!cid||x.commessaId===cid))||r.plants.find(x=>x.id===id);if(!p)return;
    backing(form,'plantId').value=p.id;const input=form.querySelector('[data-pvd-plant-search]');if(input)input.value=`${p.name} — ID SAP ${p.sap||'—'}`;form.querySelector('[data-pvd-results]')?.classList.add('hidden');syncTechnical(form);
  }

  const originalValidate=P.validateQuote?.bind(P);
  if(originalValidate)P.validateQuote=(form,lines,ids)=>{if(!form?.classList.contains('pvd-active'))return originalValidate(form,lines,ids);syncTechnical(form);const model=selectedModel(form);if(needsEconomic(model))return originalValidate(form,lines,ids);const missing=[...form.querySelectorAll('[data-pvd-section] [required]')].find(el=>el.type==='checkbox'?!el.checked:!C(el.value));if(missing){missing.focus();return'Compila i campi obbligatori richiesti dal modello.'}return''};
  const originalModelData=M.modelData?.bind(M);
  if(originalModelData)M.modelData=(doc,type)=>({...originalModelData(doc,type),...(doc?.modelFields||{})});
  const originalDraft=M.collectDraft?.bind(M);
  if(originalDraft)M.collectDraft=type=>{const doc=originalDraft(type),form=P.page()?.querySelector(type==='consuntivo'?'[data-cons-form]':'[data-pv-quote-form]');if(!doc||!form?.classList.contains('pvd-active'))return doc;syncTechnical(form);const r=registry(),cid=backing(form,'commessaId').value,pid=backing(form,'plantId').value,c=r.commesse.find(x=>x.id===cid),p=r.plants.find(x=>x.id===pid);return{...doc,commessaId:cid,commessaName:c?.name||doc.commessaName||'',commessaCode:c?.code||doc.commessaCode||'',plantId:pid,plantName:p?.name||doc.plantName||'',plantSap:p?.sap||doc.plantSap||'',city:p?.city||doc.city||'',workLocation:p?.address||doc.workLocation||'',modelFields:M.modelPayload(form).modelFields}};

  document.addEventListener('change',e=>{
    const form=e.target.closest?.('[data-pv-quote-form],[data-cons-form]');if(!form)return;
    if(e.target.matches('[data-pvm-model-select]')){form.dataset.pvdSignature='';setTimeout(()=>apply(form));return;}
    if(e.target.matches('[data-pvd-commessa]')){backing(form,'commessaId').value=e.target.value;backing(form,'plantId').value='';const search=form.querySelector('[data-pvd-plant-search]');if(search)search.value='';syncTechnical(form);return;}
    if(e.target.matches('[data-pvm-field]'))syncTechnical(form);
  },true);
  document.addEventListener('input',e=>{
    if(e.target.matches('[data-pvd-plant-search]')){const form=e.target.closest('form');backing(form,'plantId').value='';plantResults(e.target);return;}
    const form=e.target.closest?.('[data-pv-quote-form],[data-cons-form]');if(form&&e.target.matches('[data-pvm-field]'))syncTechnical(form);
  },true);
  document.addEventListener('focusin',e=>{if(e.target.matches?.('[data-pvd-plant-search]'))plantResults(e.target)},true);
  document.addEventListener('click',e=>{const button=e.target.closest?.('[data-pvd-plant]');if(!button)return;e.preventDefault();e.stopImmediatePropagation();choosePlant(button.closest('form'),button.dataset.pvdPlant,button.dataset.pvdCid)},true);

  let queued=false;
  const queue=()=>{if(queued)return;queued=true;requestAnimationFrame(()=>{queued=false;apply(P.page()?.querySelector('[data-pv-quote-form]'));apply(P.page()?.querySelector('[data-cons-form]'));});};
  const start=()=>{
    if(!document.querySelector('style[data-pvd-style]')){const style=document.createElement('style');style.dataset.pvdStyle='1';style.textContent='.pvd-section{border-left:4px solid #16733b}.pvd-plant-wrap{position:relative}.pvd-results{position:absolute;left:0;right:0;top:calc(100% + 4px);z-index:180;max-height:52vh;overflow:auto;background:#fff;border:1px solid #cbd5e1;border-radius:10px;box-shadow:0 18px 40px #0f172e2e;padding:6px}.pvd-results.hidden{display:none}.pvd-result{display:flex;width:100%;justify-content:space-between;gap:12px;padding:10px;border:0;border-bottom:1px solid #e2e8f0;background:#fff;text-align:left;cursor:pointer}.pvd-result:hover{background:#f0fdf4}.pvd-result span{font-size:.82rem;font-weight:700;color:#475569}@media(max-width:760px){.pvd-result{flex-direction:column;gap:3px}}';document.head.appendChild(style)}
    new MutationObserver(queue).observe(document.body,{childList:true,subtree:true});queue();[600,1600,3500,7000].forEach(ms=>setTimeout(queue,ms));
  };
  document.readyState==='loading'?document.addEventListener('DOMContentLoaded',start,{once:true}):start();
})();
