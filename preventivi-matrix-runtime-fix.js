(()=>{
  'use strict';

  if(window.__preventiviMatrixRuntimeFix)return;
  window.__preventiviMatrixRuntimeFix=true;

  const P=window.HeraPreventivi;
  if(!P)return;

  const clean=value=>String(value??'').trim().replace(/\s+/g,' ');
  const norm=value=>clean(value).normalize('NFD').replace(/[\u0300-\u036f]/g,'').toLowerCase().replace(/[^a-z0-9]+/g,' ').trim();
  const esc=value=>String(value??'').replace(/&/g,'&amp;').replace(/"/g,'&quot;').replace(/</g,'&lt;').replace(/>/g,'&gt;');
  const first=(object,keys)=>{
    for(const key of keys){
      const value=object?.[key];
      if(value!==undefined&&value!==null&&clean(value))return value;
    }
    return'';
  };
  const invalidName=value=>{
    const name=norm(value);
    return !name||[
      'senza nome','impianto senza nome','nome non disponibile','non disponibile','non definito',
      'undefined','null','n a','na','impianto da identificare','-'
    ].includes(name);
  };

  const COMMESSA_KEYS=['id','uid','commessaId','projectId','codiceCommessa','codice','numeroCommessa','numero','idCommessa'];
  const COMMESSA_NAME_KEYS=['nome','name','denominazione','titolo','commessa','descrizione','nomeCommessa','commessaNome'];
  const PLANT_ID_KEYS=['id','uid','impiantoId','plantId','idSap','ID SAP','ID_SAP','id_sap','sap','sapId','codiceSap','codiceImpianto'];
  const PLANT_NAME_KEYS=[
    'Denominazione Impianto','DENOMINAZIONE IMPIANTO','denominazione impianto','denominazioneImpianto','denominazione_impianto',
    'nomeImpianto','Nome Impianto','NOME IMPIANTO','nome_impianto','impiantoNome','impianto_nome','plantName','siteName',
    'denominazione','Denominazione','DENOMINAZIONE','titolo','title','label','nome','name','descrizioneImpianto','descrizione_impianto'
  ];
  const SAP_KEYS=['idSap','ID SAP','ID_SAP','id_sap','sap','sapId','codiceSap','codiceImpianto'];
  const NESTED_PLANTS=['impianti','plants','sites','impiantiDaFare','impiantiFatti','daFare','fatti','items','children','records','data','elencoImpianti','impiantiGlobal'];

  let registryCache=null;
  let registryCacheAt=0;

  function safeSource(target,name,getter){
    try{
      const value=getter();
      if(value&&typeof value==='object')target.push([name,value]);
    }catch(_){ }
  }

  function sourceList(){
    const sources=[];
    safeSource(sources,'HeraPreventiviRegistry',()=>window.HeraPreventiviRegistry?.registry?.());
    safeSource(sources,'HeraPreventiviOperational',()=>P.getOperationalData?.());
    safeSource(sources,'preventiviState',()=>P.state);

    safeSource(sources,'commesseById',()=>typeof commesseById!=='undefined'?commesseById:null);
    safeSource(sources,'commesse',()=>typeof commesse!=='undefined'?commesse:null);
    safeSource(sources,'commesseList',()=>typeof commesseList!=='undefined'?commesseList:null);
    safeSource(sources,'impiantiById',()=>typeof impiantiById!=='undefined'?impiantiById:null);
    safeSource(sources,'impianti',()=>typeof impianti!=='undefined'?impianti:null);
    safeSource(sources,'impiantiDaFare',()=>typeof impiantiDaFare!=='undefined'?impiantiDaFare:null);
    safeSource(sources,'impiantiFatti',()=>typeof impiantiFatti!=='undefined'?impiantiFatti:null);

    [
      'commesseById','commesse','commesseList','projects','cantieri','appCommesse','globalCommesse','commesseGlobal',
      'impiantiById','impianti','impiantiDaFare','impiantiFatti','allImpianti','plants','sites','globalImpianti','impiantiGlobal'
    ].forEach(name=>safeSource(sources,`window.${name}`,()=>window[name]));

    const readStorage=(storage,label)=>{
      try{
        for(let index=0;index<storage.length;index+=1){
          const key=storage.key(index)||'';
          const raw=storage.getItem(key);
          if(!raw||raw.length>25_000_000||!/^[\[{]/.test(raw.trim()))continue;
          try{
            const value=JSON.parse(raw);
            if(value&&typeof value==='object')sources.push([`${label}.${key}`,value]);
          }catch(_){ }
        }
      }catch(_){ }
    };
    readStorage(window.localStorage,'localStorage');
    readStorage(window.sessionStorage,'sessionStorage');
    return sources;
  }

  function registry(force=false){
    const now=Date.now();
    if(!force&&registryCache&&now-registryCacheAt<1500)return registryCache;

    const commesse=[];
    const plants=[];
    const seen=new WeakSet();

    const addCommessa=(raw,fallbackId='',path='')=>{
      if(!raw||typeof raw!=='object')return null;
      const name=clean(first(raw,COMMESSA_NAME_KEYS));
      const code=clean(first(raw,['codiceCommessa','codice','code','numeroCommessa','numero','contractCode']));
      const id=clean(first(raw,COMMESSA_KEYS)||fallbackId||code||(name?`commessa-${norm(name).replace(/\s/g,'-')}`:''));
      if(!id&&!name)return null;
      const item={
        raw,id:id||name,name:name||`Commessa ${code||id}`,code,
        client:clean(first(raw,['cliente','committente','clientName','ragioneSociale'])),
        contract:clean(first(raw,['numeroContratto','contratto','contractNumber'])),
        requester:clean(first(raw,['richiedente','referente','referenteCliente'])),
        path
      };
      const aliases=[item.id,item.code,item.name].map(norm).filter(Boolean);
      const existing=commesse.find(entry=>{
        const other=[entry.id,entry.code,entry.name].map(norm).filter(Boolean);
        return aliases.some(alias=>other.includes(alias));
      });
      if(existing){
        Object.keys(item).forEach(key=>{if(!clean(existing[key])&&clean(item[key]))existing[key]=item[key];});
        return existing;
      }
      commesse.push(item);
      return item;
    };

    const candidateName=raw=>{
      for(const key of PLANT_NAME_KEYS){
        const value=raw?.[key];
        if(typeof value==='string'&&!invalidName(value))return clean(value);
      }
      const nested=['anagrafica','datiImpianto','dati_impianto','impianto','plant','site','properties','dati','record'];
      for(const key of nested){
        const value=raw?.[key];
        if(!value||typeof value!=='object')continue;
        for(const nameKey of PLANT_NAME_KEYS){
          const found=value?.[nameKey];
          if(typeof found==='string'&&!invalidName(found))return clean(found);
        }
      }
      return'';
    };

    const addPlant=(raw,parent=null,path='')=>{
      if(!raw||typeof raw!=='object')return null;
      const sap=clean(first(raw,SAP_KEYS));
      const explicitId=clean(first(raw,PLANT_ID_KEYS));
      const name=candidateName(raw);
      const type=clean(first(raw,['tipologia','Tipologia impianto','tipologiaImpianto','tipo','type']));
      const city=clean(first(raw,['comune','Comune','city','localita','municipality']));
      const address=clean(first(raw,['indirizzo','Descrizione via','descrizioneVia','via','address','ubicazione']));
      const plantSignal=Boolean(sap||first(raw,['impiantoId','plantId','Denominazione Impianto','denominazioneImpianto','nomeImpianto','Tipologia impianto']));
      if(!plantSignal&&(!name||(!type&&!city&&!address)))return null;
      if(!name&&!sap&&!explicitId)return null;
      const id=explicitId||sap||`${parent?.id||'impianto'}-${norm(name).replace(/\s/g,'-')}`;
      const commessaId=clean(first(raw,['commessaId','projectId','idCommessa','codiceCommessa','commessaUid','commessa_id'])||parent?.id||'');
      const commessaName=clean(first(raw,['nomeCommessa','commessaNome','commessa_name'])||parent?.name||'');
      const displayName=name||(sap?`Impianto SAP ${sap}`:type&&city?`${type} – ${city}`:city?`Impianto di ${city}`:'Impianto da identificare');
      const item={raw,id,commessaId,commessaName,name:displayName,sap,address,city,type,path};
      const key=`${norm(commessaId||commessaName)}:${norm(sap||id)}`;
      const existing=plants.find(entry=>entry._key===key);
      const score=value=>(invalidName(value.name)?0:10)+(clean(value.sap)?4:0)+(clean(value.address)?2:0)+(clean(value.city)?2:0)+(clean(value.commessaId)?2:0);
      item._key=key;
      if(existing){
        if(score(item)>score(existing))Object.assign(existing,item);
        else Object.keys(item).forEach(field=>{if(!clean(existing[field])&&clean(item[field]))existing[field]=item[field];});
        return existing;
      }
      plants.push(item);
      return item;
    };

    const walk=(value,path='',parent=null,depth=0,fallbackId='')=>{
      if(value==null||depth>10)return;
      if(value instanceof Map){
        value.forEach((child,key)=>walk(child,`${path}.${String(key)}`,parent,depth+1,String(key)));
        return;
      }
      if(Array.isArray(value)){
        value.forEach((child,index)=>walk(child,`${path}[${index}]`,parent,depth+1,''));
        return;
      }
      if(typeof value!=='object'||seen.has(value))return;
      seen.add(value);

      const last=path.split(/[.\[]/).pop()?.replace(/\]$/,'')||'';
      const hasPlantSignal=Boolean(first(value,[...SAP_KEYS,'impiantoId','plantId','Denominazione Impianto','denominazioneImpianto','nomeImpianto','Tipologia impianto']));
      const nestedPlants=NESTED_PLANTS.some(key=>value[key]&&typeof value[key]==='object');
      const hasCommessaSignal=Boolean(first(value,['codiceCommessa','numeroCommessa','numeroContratto','committente','projectId']))||nestedPlants||/commess|project|cantier/i.test(last);
      let current=parent;
      if(hasCommessaSignal&&!hasPlantSignal)current=addCommessa(value,fallbackId,path)||parent;
      if(hasPlantSignal||/impiant|plant|site/i.test(last))addPlant(value,current,path);

      Object.entries(value).forEach(([key,child])=>{
        if(child&&typeof child==='object')walk(child,path?`${path}.${key}`:key,current,depth+1,key);
      });
    };

    sourceList().forEach(([name,value])=>walk(value,name,null,0,''));
    registryCache={
      commesse:commesse.sort((a,b)=>a.name.localeCompare(b.name,'it')),
      plants:plants.map(({_key,...plant})=>plant).sort((a,b)=>a.name.localeCompare(b.name,'it'))
    };
    registryCacheAt=now;
    return registryCache;
  }

  function isMatrixForm(form){
    if(!form?.matches('[data-pv-quote-form]'))return false;
    if(form.classList.contains('pvm-matrix-active')||form.querySelector('[data-matrix-profile]'))return true;
    const option=form.querySelector('[data-pvm-model-select] option:checked');
    const text=norm(option?.textContent||'');
    return text.includes('depurazione')&&text.includes('gas');
  }

  function cleanupDuplicateFields(form){
    if(!isMatrixForm(form))return;

    const profiles=[...form.querySelectorAll('[data-matrix-profile]')];
    profiles.slice(1).forEach(section=>section.remove());
    form.querySelectorAll('[data-pvd-section]').forEach(section=>section.remove());

    const dynamic=form.querySelector('[data-pvm-dynamic-fields]');
    if(dynamic&&!dynamic.querySelector('[data-matrix-runtime-placeholder]')){
      dynamic.innerHTML='<p class="pv-muted" data-matrix-runtime-placeholder>I campi del modello sono compilati una sola volta nella sezione “Dati richiesti dalla matrice”.</p>';
      dynamic.dataset.matrix='1';
    }
  }

  function commessaAliases(form,data){
    const select=form.querySelector('[data-matrix-commessa]');
    const value=clean(select?.value);
    const text=clean(select?.selectedOptions?.[0]?.textContent);
    const selected=data.commesse.find(item=>clean(item.id)===value)||data.commesse.find(item=>{
      const aliases=[item.id,item.code,item.name].map(norm).filter(Boolean);
      return aliases.includes(norm(value))||aliases.some(alias=>norm(text).includes(alias));
    });
    const aliases=new Set([value,text,selected?.id,selected?.code,selected?.name].map(norm).filter(Boolean));
    if(text.includes('—'))text.split('—').forEach(part=>aliases.add(norm(part)));
    if(text.includes('-'))text.split('-').forEach(part=>aliases.add(norm(part)));
    return{selected,aliases:[...aliases].filter(Boolean)};
  }

  function plantsForForm(form,data){
    const {aliases}=commessaAliases(form,data);
    if(!aliases.length)return{plants:data.plants,fallback:false};
    const related=data.plants.filter(plant=>{
      const relation=[plant.commessaId,plant.commessaName].map(norm).filter(Boolean);
      const path=norm(plant.path);
      return relation.some(value=>aliases.some(alias=>value===alias||value.includes(alias)||alias.includes(value)))||aliases.some(alias=>path&&path.includes(alias));
    });
    if(related.length)return{plants:related,fallback:false};
    const unassigned=data.plants.filter(plant=>!norm(plant.commessaId)&&!norm(plant.commessaName));
    if(unassigned.length)return{plants:unassigned,fallback:true};
    return{plants:data.plants,fallback:true};
  }

  function resultKey(plant,index){
    return `${clean(plant.id)||clean(plant.sap)||'plant'}::${index}`;
  }

  function renderPlantResults(input,force=false){
    const form=input?.closest?.('[data-pv-quote-form]');
    const box=form?.querySelector('[data-matrix-plant-results]');
    if(!form||!box||!isMatrixForm(form))return;

    cleanupDuplicateFields(form);
    const data=registry(force);
    const resolved=plantsForForm(form,data);
    const query=norm(input.value);
    const filtered=resolved.plants.filter(plant=>!query||norm(`${plant.name} ${plant.sap} ${plant.city} ${plant.address}`).includes(query)).slice(0,120);
    const map=new Map();
    form.__preventiviMatrixPlants=map;
    box.innerHTML=filtered.length?filtered.map((plant,index)=>{
      const key=resultKey(plant,index);
      map.set(key,plant);
      const details=[plant.sap?`ID SAP ${plant.sap}`:'',plant.city,plant.address].filter(Boolean).join(' • ');
      return `<button type="button" class="pv-plant-result" data-matrix-runtime-plant="${esc(key)}"><strong>${esc(plant.name)}</strong><span>${esc(details||'Dati impianto disponibili')}</span></button>`;
    }).join(''):'<p class="pv-muted">Nessun impianto trovato. Verifica la sincronizzazione oppure cerca per nome o ID SAP.</p>';
    box.classList.remove('hidden');

    const feedback=form.querySelector('[data-matrix-feedback]');
    if(feedback){
      if(!data.plants.length){
        feedback.textContent='Gli impianti non sono ancora disponibili nell’archivio locale. Premi Aggiorna app e attendi la sincronizzazione.';
        feedback.dataset.type='error';
      }else if(resolved.fallback){
        feedback.textContent='Il collegamento automatico con la commessa non era presente: la ricerca usa l’elenco completo degli impianti.';
        feedback.dataset.type='warning';
      }else if(/collegamento automatico|non sono ancora disponibili/i.test(feedback.textContent||'')){
        feedback.textContent='';
        delete feedback.dataset.type;
      }
    }
  }

  function setValue(form,selector,value){
    form.querySelectorAll(selector).forEach(element=>{
      if('value'in element)element.value=value??'';
    });
  }

  function choosePlant(button){
    const form=button.closest('[data-pv-quote-form]');
    const plant=form?.__preventiviMatrixPlants?.get(button.dataset.matrixRuntimePlant);
    if(!form||!plant)return;

    let backing=form.querySelector('[data-matrix-source-plant], input[name="plantId"], select[name="plantId"]');
    if(!backing){
      backing=document.createElement('input');
      backing.type='hidden';
      backing.name='plantId';
      backing.dataset.matrixSourcePlant='1';
      form.appendChild(backing);
    }
    backing.value=plant.id||plant.sap||'';

    const search=form.querySelector('[data-matrix-plant-search]');
    if(search){
      search.value=`${plant.name} — ID SAP ${plant.sap||'—'}`;
      search.setCustomValidity('');
    }
    setValue(form,'[data-pvm-field="denominazione_impianto"], [data-pvm-field="impianto"]',plant.name);
    setValue(form,'[data-pvm-field="id_sap"]',plant.sap);
    setValue(form,'[data-pvm-field="comune"]',plant.city);
    setValue(form,'[data-pvm-field="indirizzo"]',plant.address);
    setValue(form,'[data-pvm-field="tipologia_impianto"]',plant.type);
    setValue(form,'[name="plantSap"], [name="idSap"]',plant.sap);
    setValue(form,'[name="plantName"], [name="nomeImpianto"]',plant.name);
    setValue(form,'[name="city"], [name="comune"]',plant.city);
    setValue(form,'[name="workLocation"]',[plant.address,plant.city].filter(Boolean).join(', '));
    setValue(form,'[name="plantType"], [name="tipologiaImpianto"]',plant.type);
    form.querySelector('[data-matrix-plant-results]')?.classList.add('hidden');
    const feedback=form.querySelector('[data-matrix-feedback]');
    if(feedback&&/collegamento automatico|non sono ancora disponibili/i.test(feedback.textContent||'')){
      feedback.textContent='';
      delete feedback.dataset.type;
    }
  }

  function repairVisibleForms(){
    document.querySelectorAll('[data-pv-quote-form]').forEach(form=>cleanupDuplicateFields(form));
  }

  document.addEventListener('input',event=>{
    const input=event.target.closest?.('[data-matrix-plant-search]');
    if(!input)return;
    queueMicrotask(()=>renderPlantResults(input,false));
  },true);

  document.addEventListener('focusin',event=>{
    const input=event.target.closest?.('[data-matrix-plant-search]');
    if(input)setTimeout(()=>renderPlantResults(input,false),0);
  },true);

  document.addEventListener('click',event=>{
    const button=event.target.closest?.('[data-matrix-runtime-plant]');
    if(button){
      event.preventDefault();
      event.stopImmediatePropagation();
      choosePlant(button);
      return;
    }
    const input=event.target.closest?.('[data-matrix-plant-search]');
    if(input)setTimeout(()=>renderPlantResults(input,false),0);
  },true);

  document.addEventListener('change',event=>{
    const form=event.target.closest?.('[data-pv-quote-form]');
    if(!form)return;
    if(event.target.matches('[data-matrix-commessa]')){
      registryCache=null;
      const search=form.querySelector('[data-matrix-plant-search]');
      if(search){
        search.value='';
        search.setCustomValidity('Seleziona un impianto dai risultati.');
        setTimeout(()=>renderPlantResults(search,true),0);
      }
      setValue(form,'[data-pvm-field="denominazione_impianto"], [data-pvm-field="impianto"], [data-pvm-field="id_sap"], [data-pvm-field="comune"], [data-pvm-field="indirizzo"], [data-pvm-field="tipologia_impianto"]','');
    }
    if(event.target.matches('[data-pvm-model-select]'))setTimeout(()=>cleanupDuplicateFields(form),0);
  },true);

  window.addEventListener('storage',()=>{registryCache=null;});

  let queued=false;
  const observer=new MutationObserver(()=>{
    if(queued)return;
    queued=true;
    requestAnimationFrame(()=>{
      queued=false;
      repairVisibleForms();
    });
  });
  const start=()=>{
    observer.observe(document.body,{childList:true,subtree:true});
    repairVisibleForms();
    [150,500,1200,2500].forEach(delay=>setTimeout(repairVisibleForms,delay));
  };
  if(document.readyState==='loading')document.addEventListener('DOMContentLoaded',start,{once:true});
  else start();
})();
