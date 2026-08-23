(() => {
  'use strict';
  const VERSION = '3.0.0-unfreeze1';
  if (window.HeraRecommendedPlants?.version === VERSION) return;

  const CFG = Object.freeze({
    start: { name: 'Avola Coop', lat: 44.579, lng: 11.3635 },
    team: 2, setup: 10, margin: 1.15, road: 1.22, speed: 52,
    day: 480, max: 8, dailyMq: 5000, unload: 20, yieldRows: 120
  });
  const ROUTE_DAY = 'heraRecommendedPlantsRouteDay';
  const ROUTE_STARTED = 'heraRecommendedPlantsRouteStarted';
  const PROFILES = 'heraAdaptiveWorkProfilesV1';
  const ACTIVE = 'heraAdaptiveWorkActiveV1';
  const state = { open: false, seq: 0, plan: [], originMode: 'auto', origin: null, team: null, signature: '' };
  const txt = v => String(v ?? '').replace(/\s+/g, ' ').trim();
  const up = v => txt(v).toLocaleUpperCase('it-IT');
  const esc = v => String(v ?? '').replace(/[&<>"']/g, c => ({ '&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#039;' }[c]));
  const clamp = (n, a, b) => Math.min(b, Math.max(a, n));
  const sleep0 = () => new Promise(r => setTimeout(r, 0));
  const json = (key, fallback) => { try { return JSON.parse(localStorage.getItem(key) || 'null') ?? fallback; } catch (_) { return fallback; } };
  const save = (key, value) => { try { localStorage.setItem(key, JSON.stringify(value)); } catch (_) {} };

  function number(v) {
    if (typeof v === 'number') return Number.isFinite(v) ? v : 0;
    let s = txt(v).replace(/\s/g, '').replace(/[^0-9.,+-]/g, '');
    if (!s) return 0;
    if (s.includes(',') && s.includes('.')) s = s.lastIndexOf(',') > s.lastIndexOf('.') ? s.replace(/\./g, '').replace(',', '.') : s.replace(/,/g, '');
    else if (s.includes(',')) s = s.replace(',', '.');
    else if (/^[+-]?\d{1,3}(?:\.\d{3})+$/.test(s)) s = s.replace(/\./g, '');
    const n = Number(s); return Number.isFinite(n) ? n : 0;
  }
  function coord(v) {
    if (typeof v === 'number') return Number.isFinite(v) ? v : null;
    const n = Number(txt(v).replace(',', '.').replace(/[^0-9.+-]/g, ''));
    return Number.isFinite(n) ? n : null;
  }
  function globalValue(name) {
    try {
      if (name === 'currentImpianti' && typeof currentImpianti !== 'undefined') return currentImpianti;
      if (name === 'currentSquadre' && typeof currentSquadre !== 'undefined') return currentSquadre;
    } catch (_) {}
    return window[name];
  }
  function today() {
    const d = new Date();
    return `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,'0')}-${String(d.getDate()).padStart(2,'0')}`;
  }
  function commessaId() {
    try { if (typeof selectedCommessaId !== 'undefined') return txt(selectedCommessaId); } catch (_) {}
    return txt(window.selectedCommessaId);
  }
  function inrete() {
    const title = document.getElementById('impianti-page-title')?.textContent || '';
    const label = document.getElementById('commessa-focus-label')?.textContent || '';
    return up(`${title} ${label}`).includes('INRETE');
  }
  function rows(source, id) {
    if (Array.isArray(source)) return source;
    if (source instanceof Map) {
      const direct = source.get(id); if (direct) return Array.isArray(direct) ? direct : [direct];
      return [...source.values()].flatMap(v => Array.isArray(v) ? v : [v]);
    }
    if (source && typeof source === 'object') {
      const direct = source[id]; if (direct) return Array.isArray(direct) ? direct : [direct];
      return Object.values(source).flatMap(v => Array.isArray(v) ? v : [v]);
    }
    return [];
  }
  function countPeople(s) {
    for (const v of [s?.operatori,s?.operators,s?.membri,s?.members,s?.persone,s?.componenti,s?.utenti]) if (Array.isArray(v) && v.filter(Boolean).length) return v.filter(Boolean).length;
    return Math.round(number(s?.numeroOperatori ?? s?.operatorCount ?? s?.teamSize ?? s?.numeroPersone)) || CFG.team;
  }
  function teamInfo() {
    const id = commessaId(), day = today();
    const sources = [globalValue('currentSquadre'),window.squadre,window.squadreOggi,window.todaySquads,window.squadreByCommessa];
    let found = null;
    for (const source of sources) for (const squad of rows(source, id)) {
      if (!squad || typeof squad !== 'object') continue;
      const cid = txt(squad.commessaId ?? squad.idCommessa ?? squad.commessa);
      const date = txt(squad.data ?? squad.date ?? squad.giorno).slice(0,10);
      if (cid && id && cid !== id) continue;
      if (date && date !== day) continue;
      const e = up(JSON.stringify(squad));
      const T = e.match(/(?:^|[^A-Z0-9])T\s*[-_.]?\s*\d{1,5}(?:[^A-Z0-9]|$)/);
      const R = e.match(/(?:^|[^A-Z0-9])R\s*[-_.]?\s*\d{1,5}(?:[^A-Z0-9]|$)/);
      const A = e.match(/(?:^|[^A-Z0-9])A\s*[-_.]?\s*\d{1,5}(?:[^A-Z0-9]|$)/);
      found = {
        teamSize: countPeople(squad), equipmentText: e,
        big: Boolean(T), small: Boolean(R), trincia: Boolean(T || R || /TRINCIA|TRINCIATR/.test(e)),
        daily: Boolean(A || /DAILY|IVECO/.test(e)), decesp: /DECESPUGLI/.test(e), soff: /SOFFIAT/.test(e),
        siepe: /TAGLIASIEP/.test(e), moto: /MOTOSEG/.test(e), spazz: /SPAZZATR/.test(e), piattaforma: /PIATTAFORM/.test(e),
        bigCode: T ? T[0].replace(/[^A-Z0-9]/g,'') : '', smallCode: R ? R[0].replace(/[^A-Z0-9]/g,'') : '', dailyCode: A ? A[0].replace(/[^A-Z0-9]/g,'') : ''
      };
      if (found.trincia || found.daily) return found;
    }
    return found || { teamSize:2,equipmentText:'',big:false,small:false,trincia:false,daily:false,decesp:false,soff:false,siepe:false,moto:false,spazz:false,piattaforma:false,bigCode:'',smallCode:'',dailyCode:'' };
  }

  function coordinates(item) {
    try {
      const fixed = window.HeraCoordinateRepair?.getCoordinates?.(item);
      if (Number.isFinite(Number(fixed?.lat)) && Number.isFinite(Number(fixed?.lng))) return { lat:Number(fixed.lat), lng:Number(fixed.lng) };
    } catch (_) {}
    const latRaw = item?.latitudine ?? item?.lat ?? item?.latitude ?? item?.gpsY ?? item?.GPSY ?? item?.gps_y;
    const lngRaw = item?.longitudine ?? item?.lng ?? item?.lon ?? item?.longitude ?? item?.gpsX ?? item?.GPSX ?? item?.gps_x;
    try {
      const fixed = window.HeraCoordinateRepair?.diagnose?.(latRaw,lngRaw);
      if (fixed?.valid) return { lat:Number(fixed.latitude), lng:Number(fixed.longitude) };
    } catch (_) {}
    let lat = coord(latRaw), lng = coord(lngRaw);
    if (Number.isFinite(lat) && Number.isFinite(lng)) {
      if (!(lat>=35&&lat<=49&&lng>=5&&lng<=20) && lng>=35&&lng<=49&&lat>=5&&lat<=20) [lat,lng]=[lng,lat];
      if (Math.abs(lat)<=90 && Math.abs(lng)<=180 && lat && lng) return {lat,lng};
    }
    const raw = txt(item?.coordinate ?? item?.coordinates ?? item?.coordinateGps ?? item?.coordinateGPS ?? item?.gps ?? item?.GPS ?? item?.['Coordinate GPS'] ?? item?.['Coordinate GPS(X)/GPS(Y)']);
    const m = raw.match(/[-+]?\d{1,3}(?:[.,]\d+)?/g) || [];
    if (m.length >= 2) {
      lat=coord(m[0]); lng=coord(m[1]);
      if (!(lat>=35&&lat<=49&&lng>=5&&lng<=20) && lng>=35&&lng<=49&&lat>=5&&lat<=20) [lat,lng]=[lng,lat];
      if (Math.abs(lat)<=90 && Math.abs(lng)<=180) return {lat,lng};
    }
    return null;
  }
  const done = item => Boolean(item?.done) || ['FATTO','DONE','COMPLETATO'].includes(up(item?.stato ?? item?.status));
  const plantId = item => txt(item?.physicalPlantId ?? item?.impiantoId ?? item?.id ?? item?.idSap ?? item?.idSAP ?? item?.sap ?? item?.denominazione ?? item?.nome);
  const title = item => txt(item?.denominazione ?? item?.nome ?? item?.impianto ?? item?.name) || 'Impianto';
  const comune = item => txt(item?.comune ?? item?.municipality ?? item?.citta);
  function workSource(item) {
    const fields = ['lavorazioni','lavorazioniRichieste','lavorazioniStraordinarie','extra','extras','lavoriExtra','interventi','attivita','workItems','tipologiaIntervento','tipologiaLavorazione','lavorazione','attivitaRichiesta','descrizioneLavoro','descrizione','noteLavoro','note','noteImpianto','modalitaEsecuzione','accesso','difficolta','atex','ATEX'];
    return up(fields.map(k => { const v=item?.[k]; try { return typeof v === 'string' ? v : v ? JSON.stringify(v) : ''; } catch (_) { return ''; } }).join(' '));
  }
  function area(item, source) {
    for (const v of [item?.areaMq,item?.mq,item?.superficieMq,item?.metriQuadri,item?.superficie,item?.quantitaMq,item?.quantita]) { const n=number(v); if(n>0) return n; }
    const m=source.match(/(\d+(?:[.,]\d+)?)\s*(?:MQ|M2|M²)/); return m ? number(m[1]) : 0;
  }
  function adaptive(kind, baseMinutes, baseRate, quantity, team) {
    const key = `${kind}|team=${Math.max(1,Math.round(team.teamSize||2))}${kind==='sfalcio'?`|trincia=${team.trincia?1:0}`:''}`;
    const p = json(PROFILES,{})[key];
    if (!p || !(Number(p.emaRate)>0)) return { minutes:baseMinutes,samples:0,key };
    const confidence = clamp(Number(p.weightedCount||p.count||0)/8,0,1);
    const ratio = 1-confidence + clamp(Number(p.emaRate)/Math.max(1,baseRate),.55,1.8)*confidence;
    return { minutes:Math.max(5,Math.round(baseMinutes*ratio)),samples:Number(p.count||0),key,quantity,baseRate };
  }
  function estimate(item, team) {
    const s=workSource(item), mq=area(item,s), tf=clamp(Math.sqrt(2/Math.max(1,team.teamSize)),.58,1.35);
    let difficulty=1;
    if(/SCARPAT|PENDEN|FOSSO|ARGINE/.test(s)) difficulty+=.18;
    if(/STRETT|ACCESSO DIFFIC|OSTACOL|RECINZ/.test(s)) difficulty+=.12;
    if(/ERBA ALTA|INFESTANT|ABBANDON/.test(s)) difficulty+=.18;
    if(/ATEX|AREA GAS|CABINA GAS|REMI|GRF/.test(s)) difficulty+=.08;
    let kind='generico',label='Lavorazione',q=1,rate=30;
    if(/SIEP/.test(s)&&/POTAT|TAGLI|RIFIL/.test(s)){kind='potatura_siepe';label='Potatura siepe';q=Math.max(1,number(item?.metriSiepe??item?.lunghezzaSiepe??item?.ml)/10||1);rate=/ALTA|DIFFIC|PIATTAFORMA/.test(s)?32:20;}
    else if(/POTAT/.test(s)&&/ALBER|PIANT|ARBOR/.test(s)){kind='potatura_alberi';label='Potatura alberi';q=Math.max(1,number(item?.numeroPiante??item?.piante??item?.numeroAlberi)||1);rate=/ALTO|ALTA|DIFFIC|PIATTAFORMA/.test(s)?60:35;}
    else if(/ABBATT/.test(s)){kind='abbattimento';label='Abbattimento';q=Math.max(1,number(item?.numeroPiante??item?.piante)||1);rate=75;}
    else if(/ROV|BONIFIC/.test(s)){kind='bonifica_rovi';label='Bonifica rovi/vegetazione';q=Math.max(1,(mq||100)/100);rate=32;}
    else if(/A MANO|MANUAL|DECESP|RIFINIT/.test(s)){kind=/RIFINIT/.test(s)?'rifiniture':'lavoro_manual';label=/RIFINIT/.test(s)?'Rifiniture/decespugliatore':'Lavorazione a mano';q=Math.max(1,(mq||100)/100);rate=/A MANO|MANUAL/.test(s)?28:20;}
    else if(/SFALC|TRINCI|ERBA/.test(s)||mq){kind='sfalcio';label=team.trincia?'Sfalcio con trincia':'Sfalcio senza trincia';q=Math.max(1,(mq||100)/100);rate=team.trincia?12:20;}
    else if(/PULIZ|RACCOLT|ASPORT|SMALT/.test(s)){kind='pulizia_raccolta';label='Pulizia/raccolta materiale';q=Math.max(1,(mq||100)/100);rate=8;}
    let machine=1;
    if(kind==='sfalcio'&&team.big) machine=.55; else if(kind==='sfalcio'&&team.small) machine=.68;
    else if(kind==='potatura_siepe'&&team.siepe) machine=.82; else if((kind==='potatura_alberi'||kind==='abbattimento')&&team.moto) machine=.82;
    const raw=Math.max(5,Math.round(q*rate*tf*difficulty*machine));
    const learned=adaptive(kind,raw,rate,q,team);
    const atex=/ATEX/.test(s)?10:0;
    return { minutes:Math.round((CFG.setup+learned.minutes+atex)*CFG.margin), collectionMq:mq, samples:learned.samples, units:[{kind,quantity:q,baseRate:rate,minutes:learned.minutes}], breakdown:[{label,minutes:learned.minutes},...(atex?[{label:'Preparazione sicurezza ATEX',minutes:atex}]:[])] };
  }
  function advice(item, team) {
    const s=workSource(item), mq=area(item,s), out=[];
    const add=(label,ok)=>{if(!out.some(x=>x.label===label))out.push({label,ok:Boolean(ok)});};
    if(/SFALC|TRINCI|ERBA|VERDE/.test(s)||mq){ if(mq>=2500&&!/STRETT|FOSSO|SCARPAT/.test(s))add('Trattore grande con trincia (T###)',team.big); else add('Trattore piccolo con trincia (R###)',team.small); add('Decespugliatore',team.decesp); add('Soffiatore',team.soff); }
    if(/SIEP/.test(s))add('Tagliasiepe',team.siepe);
    if(/ALBER|ARBOR|MOTOSEG|ABBATT/.test(s))add('Motosega / attrezzatura potatura',team.moto);
    if(/ALTO|ALTA|PIATTAFORM/.test(s))add('Piattaforma',team.piattaforma);
    if(/SPAZZATR/.test(s))add('Spazzatrice',team.spazz);
    if(/ATEX|AREA GAS|CABINA GAS|REMI|GRF/.test(s))add('Kit e DPI ATEX',/ATEX/.test(team.equipmentText));
    if(/RACCOLT|ASPORT|SMALT/.test(s)||inrete())add('Daily (A###)',team.daily);
    return out;
  }

  function km(a,b){const r=6371,rad=v=>v*Math.PI/180,dlat=rad(b.lat-a.lat),dlng=rad(b.lng-a.lng),h=Math.sin(dlat/2)**2+Math.cos(rad(a.lat))*Math.cos(rad(b.lat))*Math.sin(dlng/2)**2;return 2*r*Math.asin(Math.sqrt(h))*CFG.road;}
  const drive = distance => Math.max(3,Math.round(distance/CFG.speed*60));
  const fmtMin = n => {n=Math.max(0,Math.round(n||0));const h=Math.floor(n/60),m=n%60;return h?(m?`${h} h ${m} min`:`${h} h`):`${m} min`;};
  const fmtKm = n => `${Number(n||0).toLocaleString('it-IT',{maximumFractionDigits:1})} km`;
  function geolocation(){return new Promise(resolve=>{if(!navigator.geolocation)return resolve(null);navigator.geolocation.getCurrentPosition(p=>resolve({lat:p.coords.latitude,lng:p.coords.longitude}),()=>resolve(null),{enableHighAccuracy:false,timeout:2800,maximumAge:120000});});}
  function routeStarted(){try{return localStorage.getItem(ROUTE_DAY)===today()&&localStorage.getItem(ROUTE_STARTED)==='1';}catch(_){return false;}}
  function markRoute(){try{localStorage.setItem(ROUTE_DAY,today());localStorage.setItem(ROUTE_STARTED,'1');}catch(_){}}
  async function origin(current){
    if(state.originMode==='avola'||(state.originMode==='auto'&&!routeStarted()))return {...CFG.start,mode:'avola'};
    const live=await geolocation(); if(!current())return null;
    return live?{...live,name:'Posizione attuale',mode:'position'}:{...CFG.start,mode:'avola'};
  }
  async function plan(items, start, team, current){
    const candidates=[];
    for(let i=0;i<items.length;i++){if(!current())return[];const item=items[i];if(!done(item)){const c=coordinates(item);if(c)candidates.push({item,coords:c});}if(i&&i%CFG.yieldRows===0)await sleep0();}
    const out=[];let cursor={lat:start.lat,lng:start.lng},total=0,load=0;
    while(candidates.length&&out.length<CFG.max){if(!current())return[];let ni=0,nk=Infinity;for(let i=0;i<candidates.length;i++){const d=km(cursor,candidates[i].coords);if(d<nk){nk=d;ni=i;}if(i&&i%CFG.yieldRows===0)await sleep0();}
      const next=candidates.splice(ni,1)[0],work=estimate(next.item,team),travel=drive(nk);let nextTotal=total+travel+work.minutes,nextLoad=load+(inrete()?work.collectionMq:0),unload=null;
      if(inrete()&&nextLoad>=CFG.dailyMq){const back=km(next.coords,CFG.start);unload={km:back,driveMinutes:drive(back),unloadMinutes:CFG.unload,loadMq:nextLoad};nextTotal+=unload.driveMinutes+unload.unloadMinutes;}
      if(out.length&&nextTotal>CFG.day)break;
      if(unload){nextLoad%=CFG.dailyMq;cursor={lat:CFG.start.lat,lng:CFG.start.lng};}else cursor=next.coords;
      total=nextTotal;load=nextLoad;out.push({...next,km:nk,driveMinutes:travel,work,workMinutes:work.minutes,unload,cumulativeMinutes:total,fitsDay:total<=CFG.day,loadMqAfter:load,equipmentAdvice:advice(next.item,team)});await sleep0();
    }
    return out;
  }

  function panel(){const card=document.getElementById('impianti-card');if(!card)return null;let p=document.getElementById('recommended-plants-panel');if(!p){p=document.createElement('section');p.id='recommended-plants-panel';p.className='recommended-plants-panel hidden';p.setAttribute('aria-live','polite');p.setAttribute('aria-busy','false');const list=document.getElementById('impianti-lista');list?card.insertBefore(p,list):card.appendChild(p);bindPanel(p);}return p;}
  function button(){const tabs=document.querySelector('#impianti-card .view-tabs');if(!tabs)return null;let b=document.getElementById('recommended-plants-btn');if(!b){b=document.createElement('button');b.id='recommended-plants-btn';b.className='btn recommended-plants-btn';b.type='button';b.textContent='✨ Impianti consigliati';b.setAttribute('aria-pressed','false');const todo=document.getElementById('view-todo-btn');todo?.nextSibling?tabs.insertBefore(b,todo.nextSibling):tabs.appendChild(b);b.addEventListener('click',()=>state.open?close():(state.open=true,state.originMode='auto',render()));}return b;}
  const navUrl=e=>`https://www.google.com/maps/dir/?api=1&destination=${encodeURIComponent(`${e.coords.lat},${e.coords.lng}`)}&travelmode=driving`;
  function teamLabel(t){if(t.big)return`🚜 ${t.bigCode||'T'} · trincia grande`;if(t.small)return`🚜 ${t.smallCode||'R'} · trincia piccola`;if(t.trincia)return'🚜 Trincia presente';if(t.decesp)return'🧑‍🌾 Decespugliatore presente';return'🧑‍🌾 Nessuna trincia rilevata';}
  function signature(){return JSON.stringify({o:state.origin?.mode,t:state.team?.teamSize,e:state.team?.equipmentText,p:state.plan.map(x=>[plantId(x.item),Math.round(x.km*10),x.driveMinutes,x.workMinutes,x.traffic?.durationMinutes||0,x.weather?.label||'',x.unload?.driveMinutes||0])});}
  function draw(force=false){const p=panel(),b=button();if(!p||!b||!state.open||!state.team||!state.origin)return;const sig=signature();if(!force&&sig===state.signature&&p.querySelector('.recommended-list'))return;state.signature=sig;
    const totalDrive=state.plan.reduce((s,x)=>s+x.driveMinutes+(x.unload?.driveMinutes||0),0),totalWork=state.plan.reduce((s,x)=>s+x.workMinutes+(x.unload?.unloadMinutes||0),0),originLabel=state.origin.mode==='avola'?'Avola Coop · primo impianto':'Posizione attuale · squadra già in giro';
    p.innerHTML=`<header class="recommended-head"><div><span class="recommended-kicker">PIANIFICAZIONE GIORNATA</span><h3>✨ Impianti consigliati</h3><p>${esc(originLabel)}</p></div><button id="recommended-close-btn" class="btn btn-small" type="button">Chiudi</button></header>
    <div class="recommended-summary"><span><strong>${state.plan.length}</strong> consigliati</span><span>👥 ${state.team.teamSize} persone</span><span>${esc(teamLabel(state.team))}</span><span>🚐 ${fmtMin(totalDrive)}</span><span>🌿 ${fmtMin(totalWork)}</span></div>
    <p class="recommended-rule">Calcolo base: 2 persone. Considera squadra, mezzi, lavorazioni, extra, accesso, difficoltà e sicurezza. Margine 15%.</p>
    ${inrete()?`<p class="recommended-rule">INRETE: raccolta inclusa. Ogni 5.000 m² il Daily viene considerato pieno e viene inserito il rientro ad Avola Coop con 20 minuti di scarico.${state.team.daily?'':' ⚠️ Daily non rilevato nella squadra.'}</p>`:''}
    <div class="recommended-origin-actions"><button id="recommended-origin-avola" class="btn btn-small ${state.origin.mode==='avola'?'btn-primary':''}" type="button">🏠 Primo giro da Avola</button><button id="recommended-origin-live" class="btn btn-small ${state.origin.mode==='position'?'btn-primary':''}" type="button">📍 Sono già in giro</button></div>
    <div class="recommended-list">${state.plan.length?state.plan.map((e,i)=>{const ratio=e.work.minutes?e.workMinutes/e.work.minutes:1,details=(e.work.breakdown||[]).map(x=>`${esc(x.label)}: ${fmtMin(x.minutes*ratio)}`).join(' · '),missing=e.equipmentAdvice.filter(x=>!x.ok);return`<article class="recommended-item" data-recommended-id="${esc(plantId(e.item))}"><div class="recommended-rank">${i+1}</div><div class="recommended-main"><strong>${esc(title(e.item))}</strong><span>${esc(comune(e.item))}</span><small>🚐 ${fmtKm(e.km)} · ${fmtMin(e.driveMinutes)}${e.traffic?` · 🚦 ${fmtMin(e.traffic.durationMinutes)}${e.traffic.delayMinutes?` (+${e.traffic.delayMinutes} min)`:''}`:''} &nbsp; 🌿 lavoro ${fmtMin(e.workMinutes)}</small><small>${details}</small><small>${e.work.samples?`🧠 Stima adattiva · ${e.work.samples} campioni reali`:'🧠 Stima iniziale · apprende dai lavori reali'}${e.weather?` · 🌦️ ${esc(e.weather.label)}`:''}</small>${inrete()&&e.work.collectionMq?`<small>🟩 Raccolta stimata: ${Math.round(e.work.collectionMq).toLocaleString('it-IT')} m² · carico Daily: ${Math.round(e.loadMqAfter).toLocaleString('it-IT')} / 5.000 m²</small>`:''}${e.unload?`<small>🚛 Daily pieno → rientro Avola: ${fmtKm(e.unload.km)} · ${fmtMin(e.unload.driveMinutes)} + scarico ${fmtMin(e.unload.unloadMinutes)}</small>`:''}<div class="recommended-equipment-advice"><strong>🧰 Attrezzature consigliate</strong><br>${e.equipmentAdvice.length?e.equipmentAdvice.map(x=>`${x.ok?'✅':'⚠️'} ${esc(x.label)}`).join(' · '):'Nessuna attrezzatura speciale rilevata'}${missing.length?`<br><strong>Da verificare:</strong> ${missing.map(x=>esc(x.label)).join(', ')}`:'<br><strong>✅ Attrezzatura completa</strong>'}</div></div><a class="btn recommended-nav-btn" href="${navUrl(e)}" target="_blank" rel="noopener" data-recommended-nav="1">NAVIGA</a></article>`;}).join(''):'<div class="recommended-empty"><strong>Nessun impianto da pianificare.</strong><span>Gli impianti da fare devono avere coordinate valide.</span></div>'}</div>
    ${state.plan.length?'<div class="recommended-footer"><button id="recommended-start-route" class="btn btn-primary" type="button">AVVIA GIRO CONSIGLIATO</button><small>La lista normale resta disponibile sotto questo pannello.</small></div>':''}`;
    p.setAttribute('aria-busy','false');
  }
  function close(){state.seq++;state.open=false;const p=document.getElementById('recommended-plants-panel'),b=document.getElementById('recommended-plants-btn');p?.classList.add('hidden');p?.setAttribute('aria-busy','false');b?.classList.remove('btn-primary');b?.setAttribute('aria-pressed','false');}
  function entryFrom(node){const id=txt(node?.closest?.('[data-recommended-id]')?.dataset?.recommendedId);return state.plan.find(e=>plantId(e.item)===id);}
  function startLearning(e){if(!e)return;save(ACTIVE,{date:today(),plantId:plantId(e.item),plantTitle:title(e.item),navigateAt:Date.now(),travelMinutes:e.driveMinutes,team:state.team,units:e.work.units});}
  function finishLearning(node){const active=json(ACTIVE,null);if(!active||active.date!==today()||!active.navigateAt)return;const body=up(node?.closest?.('article,li,.card,.impianto-card,.impianto-item')?.textContent||'');if(active.plantTitle&&body&&!body.includes(up(active.plantTitle)))return;const total=(Date.now()-active.navigateAt)/60000;if(total<8||total>720)return;const minutes=Math.max(5,total-Number(active.travelMinutes||0)),units=active.units?.length?active.units:[{kind:'generico',quantity:1,baseRate:30,minutes:30}],base=units.reduce((s,u)=>s+Math.max(1,u.minutes||u.baseRate||1),0),profiles=json(PROFILES,{});for(const u of units){const observed=minutes*(Math.max(1,u.minutes||u.baseRate||1)/base)/Math.max(.1,u.quantity||1),key=`${u.kind}|team=${Math.max(1,Math.round(active.team?.teamSize||2))}${u.kind==='sfalcio'?`|trincia=${active.team?.trincia?1:0}`:''}`,p=profiles[key]||{count:0,weightedCount:0,emaRate:u.baseRate||30},a=clamp(.16+Math.min(.12,(p.count||0)*.01),.16,.28);profiles[key]={...p,count:(p.count||0)+1,weightedCount:(p.weightedCount||0)+.8,emaRate:p.count?p.emaRate*(1-a)+observed*a:observed,lastUpdated:new Date().toISOString()};}save(PROFILES,profiles);try{localStorage.removeItem(ACTIVE);}catch(_){};}
  function bindPanel(p){if(p.dataset.bound)return;p.dataset.bound='1';p.addEventListener('click',ev=>{if(ev.target.closest?.('#recommended-close-btn')){ev.preventDefault();close();return;}if(ev.target.closest?.('#recommended-origin-avola')){state.originMode='avola';render();return;}if(ev.target.closest?.('#recommended-origin-live')){state.originMode='position';render();return;}const start=ev.target.closest?.('#recommended-start-route');if(start){markRoute();state.originMode='position';start.textContent='✓ GIRO AVVIATO';return;}const nav=ev.target.closest?.('[data-recommended-nav="1"]');if(nav){markRoute();startLearning(entryFrom(nav));}});}
  async function render(){const seq=++state.seq,current=()=>seq===state.seq&&state.open,p=panel(),b=button();if(!p||!b||!state.open)return;state.signature='';p.classList.remove('hidden');p.setAttribute('aria-busy','true');b.classList.add('btn-primary');b.setAttribute('aria-pressed','true');p.innerHTML='<div class="recommended-loading">Calcolo gli impianti consigliati…</div>';await sleep0();if(!current())return;const items=globalValue('currentImpianti');if(!Array.isArray(items)||!items.length){p.innerHTML='<div class="recommended-empty"><strong>Nessun impianto disponibile.</strong><span>Attendi il caricamento della commessa e riprova.</span></div>';p.setAttribute('aria-busy','false');return;}const o=await origin(current);if(!o||!current())return;const t=teamInfo(),list=await plan(items,o,t,current);if(!current())return;state.origin=o;state.team=t;state.plan=list;draw(true);}
  const pageVisible=p=>Boolean(p&&!p.hidden&&!p.classList.contains('hidden')&&p.getAttribute('aria-hidden')!=='true'&&p.style?.display!=='none'&&p.style?.visibility!=='hidden');
  function install(){button();panel();const page=document.getElementById('impianti-page');if(page)new MutationObserver(()=>pageVisible(page)?(button(),panel()):close()).observe(page,{attributes:true,attributeFilter:['class','hidden','aria-hidden','style'],childList:true,subtree:true});}
  document.addEventListener('click',ev=>{const b=ev.target?.closest?.('button,[role="button"],input[type="button"],input[type="submit"]');if(b&&/(^|\s)FATTO($|\s)/.test(up(b.textContent||b.getAttribute('aria-label')||b.value||'')))setTimeout(()=>finishLearning(b),900);},true);
  window.addEventListener('popstate',()=>setTimeout(()=>{if(!pageVisible(document.getElementById('impianti-page')))close();},0));
  window.addEventListener('hashchange',()=>setTimeout(()=>{if(!pageVisible(document.getElementById('impianti-page')))close();},0));

  window.HeraRecommendedPlants={installed:true,version:VERSION,config:{...CFG,baselineTeamSize:2,planningMinutes:CFG.day,maxVisible:CFG.max,inreteDailyCapacityMq:CFG.dailyMq,inreteUnloadMinutes:CFG.unload,start:CFG.start},open:()=>{state.open=true;state.originMode='auto';return render();},close,refresh:render,refreshDecorations:()=>draw(false),markRouteStartedToday:markRoute,getTeamInfo:teamInfo,getState:()=>({open:state.open,lastPlan:state.plan.slice(),originMode:state.originMode,teamSize:state.team?.teamSize||2,hasTrincia:Boolean(state.team?.trincia),teamInfo:state.team,origin:state.origin})};
  window.HeraEquipmentAdvisor={installed:true,version:VERSION,getTeamInfo:teamInfo,recommendations:advice,getProfiles:()=>({}),refresh:()=>draw(false)};
  window.HeraAdaptiveWorkLearning={installed:true,version:VERSION,getProfiles:()=>json(PROFILES,{}),getActiveSession:()=>json(ACTIVE,null),reset:()=>{try{localStorage.removeItem(PROFILES);localStorage.removeItem(ACTIVE);}catch(_){}state.signature='';draw(true);},applyToRecommendedPanel:()=>{state.signature='';draw(true);},finishActive:()=>finishLearning(null)};
  if(document.readyState==='loading')document.addEventListener('DOMContentLoaded',install,{once:true});else install();
})();
