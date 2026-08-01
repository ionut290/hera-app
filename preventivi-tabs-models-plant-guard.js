(()=>{
  'use strict';

  const PV=window.HeraPreventivi;
  const M=window.HeraPreventiviModels;
  if(!PV||!M)return;

  const clean=value=>String(value??'').trim().replace(/\s+/g,' ');
  const normalize=value=>clean(value).normalize('NFD').replace(/[\u0300-\u036f]/g,'').toLowerCase();
  const invalidName=value=>{
    const name=normalize(value).replace(/[._-]+/g,' ').replace(/\s+/g,' ').trim();
    return !name||['senza nome','impianto senza nome','nome non disponibile','non disponibile','non definito','undefined','null','n a','na','-'].includes(name);
  };
  const nameKeys=[
    'Denominazione Impianto','DENOMINAZIONE IMPIANTO','denominazione impianto','denominazioneImpianto','denominazione_impianto',
    'nomeImpianto','Nome Impianto','NOME IMPIANTO','nome_impianto','impiantoNome','impianto_nome','plantName','siteName',
    'denominazione','Denominazione','DENOMINAZIONE','titolo','title','label','nome','name','descrizioneImpianto','descrizione_impianto'
  ];
  const nestedKeys=['anagrafica','datiImpianto','dati_impianto','impianto','plant','site','properties','dati','record'];

  function candidateName(raw,depth=0,seen=new WeakSet()){
    if(!raw||typeof raw!=='object'||depth>3||seen.has(raw))return'';
    seen.add(raw);
    for(const key of nameKeys){
      const value=raw[key];
      if(typeof value==='string'&&!invalidName(value))return clean(value);
    }
    for(const key of nestedKeys){
      const value=raw[key];
      if(value&&typeof value==='object'){
        const found=candidateName(value,depth+1,seen);
        if(found)return found;
      }
    }
    return'';
  }

  function plantFallback(plant){
    const sap=clean(plant.sap||plant.idSap||plant.id_sap);
    const type=invalidName(plant.type)?'':clean(plant.type);
    const city=invalidName(plant.city)?'':clean(plant.city);
    if(type&&city)return`${type} – ${city}${sap?` — ID SAP ${sap}`:''}`;
    if(city)return`Impianto di ${city}${sap?` — ID SAP ${sap}`:''}`;
    return sap?`Impianto SAP ${sap}`:'Impianto da identificare';
  }

  function sanitizeRegistry(data){
    const commesse=Array.isArray(data?.commesse)?data.commesse:[];
    const plants=(Array.isArray(data?.plants)?data.plants:[]).map(plant=>{
      const recovered=!invalidName(plant.name)?clean(plant.name):candidateName(plant.raw);
      return{...plant,name:recovered||plantFallback(plant)};
    });
    const unique=new Map();
    plants.forEach(plant=>{
      const key=`${clean(plant.commessaId)}:${clean(plant.sap)||clean(plant.id)}`;
      const current=unique.get(key);
      const score=item=>(invalidName(item.name)?0:10)+(clean(item.sap)?3:0)+(clean(item.address)?2:0)+(clean(item.city)?1:0);
      if(!current||score(plant)>score(current))unique.set(key,plant);
    });
    return{...data,commesse,plants:[...unique.values()]};
  }

  const Registry=window.HeraPreventiviRegistry;
  if(Registry?.registry&&!Registry.__readableNames){
    const originalRegistry=Registry.registry.bind(Registry);
    Registry.registry=()=>sanitizeRegistry(originalRegistry());
    Registry.plantLabel=plant=>`${plant.name} — ID SAP ${clean(plant.sap)||'—'}`;
    Registry.__readableNames=true;
  }
  if(PV.getOperationalData&&!PV.__readablePlantNames){
    const originalOperational=PV.getOperationalData.bind(PV);
    PV.getOperationalData=()=>{
      const value=originalOperational()||{};
      const normalized=sanitizeRegistry({commesse:value.commesse||[],plants:value.plants||[]});
      return{...value,commesse:normalized.commesse,plants:normalized.plants};
    };
    PV.__readablePlantNames=true;
  }

  let queued=false;
  const desiredTabs=[
    ['quotes','Preventivi'],
    ['prices','Prezziari'],
    ['consuntivi','Consuntivi'],
    ['models','Modelli']
  ];

  function ensureTabs(){
    const page=PV.page?.()||document.getElementById('preventivi-page');
    const nav=page?.querySelector('.pv-nav');
    if(!nav)return;
    desiredTabs.forEach(([view,label])=>{
      if(nav.querySelector(`[data-pv-view="${view}"]`))return;
      nav.insertAdjacentHTML('beforeend',`<button type="button" class="pv-tab" data-pv-view="${view}">${PV.escapeHtml(label)}</button>`);
    });
    nav.querySelectorAll('[data-pv-view]').forEach(button=>button.classList.toggle('active',button.dataset.pvView===PV.state.view));
  }

  function currentDoc(form){
    if(form.matches('[data-cons-form]')){
      const id=PV.state.editingConsuntivoId;
      return id&&id!=='new'?(PV.state.consuntivi||[]).find(item=>item.id===id)||{}:{};
    }
    const id=PV.state.editingQuoteId;
    return id&&id!=='new'?PV.getQuote?.(id)||{}:{};
  }

  function ensureModelSelector(form){
    if(!form||form.querySelector('[data-pvm-model-select]')||typeof M.modelSection!=='function')return;
    const type=form.matches('[data-cons-form]')?'consuntivo':'preventivo';
    const firstCard=form.querySelector('.pv-form-card');
    if(!firstCard)return;
    firstCard.insertAdjacentHTML('afterend',M.modelSection(type,currentDoc(form)));
    M.renderDynamic?.(form,currentDoc(form));
    form.dataset.pvmDecorated='1';
  }

  function sanitizeVisibleOptions(page){
    page.querySelectorAll('option').forEach(option=>{
      if(!/senza nome/i.test(option.textContent||''))return;
      const sap=(option.textContent||'').match(/(?:ID\s*SAP\s*)?([A-Z0-9][A-Z0-9._/-]{2,})\s*$/i)?.[1]||clean(option.value);
      option.textContent=sap?`Impianto SAP ${sap} — ID SAP ${sap}`:'Impianto da identificare';
    });
    page.querySelectorAll('input[list]').forEach(input=>{
      const list=input.list;
      list?.querySelectorAll('option').forEach(option=>{
        if(!/senza nome/i.test(`${option.value} ${option.textContent}`))return;
        const sap=(`${option.textContent} ${option.value}`).match(/([A-Z0-9][A-Z0-9._/-]{2,})/i)?.[1]||'';
        option.value=sap?`Impianto SAP ${sap}`:'Impianto da identificare';
        option.textContent=sap?`ID SAP ${sap}`:'';
      });
    });
  }

  function repair(){
    queued=false;
    const page=PV.page?.()||document.getElementById('preventivi-page');
    if(!page)return;
    ensureTabs();
    ensureModelSelector(page.querySelector('[data-pv-quote-form]'));
    ensureModelSelector(page.querySelector('[data-cons-form]'));
    sanitizeVisibleOptions(page);
  }

  function queue(){
    if(queued)return;
    queued=true;
    requestAnimationFrame(repair);
  }

  const originalEnsure=PV.ensurePage?.bind(PV);
  if(originalEnsure)PV.ensurePage=()=>{const result=originalEnsure();queue();return result};
  const originalRender=PV.renderCurrentView?.bind(PV);
  if(originalRender)PV.renderCurrentView=()=>{const result=originalRender();queue();return result};

  const observer=new MutationObserver(queue);
  const start=()=>{
    observer.observe(document.body,{childList:true,subtree:true});
    queue();
    [250,750,1500,3000].forEach(delay=>setTimeout(queue,delay));
  };
  if(document.readyState==='loading')document.addEventListener('DOMContentLoaded',start,{once:true});
  else start();
})();
