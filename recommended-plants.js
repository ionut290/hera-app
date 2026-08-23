(() => {
  "use strict";
  if (window.HeraRecommendedPlants?.installed) return;

  const CONFIG = Object.freeze({
    start: { name:"Avola Coop", address:"Via Galliera 14/A, 40013 Castel Maggiore (BO)", lat:44.5790, lng:11.3635 },
    baselineTeamSize: 2,
    setupMinutesPerPlant: 10,
    contingencyPct: 15,
    roadFactor: 1.22,
    averageRoadSpeedKmh: 52,
    planningMinutes: 8 * 60,
    maxVisible: 8,
    inreteDailyCapacityMq: 5000,
    inreteUnloadMinutes: 20
  });

  const STORAGE_ROUTE_DAY = "heraRecommendedPlantsRouteDay";
  const STORAGE_ROUTE_STARTED = "heraRecommendedPlantsRouteStarted";
  const state = { open:false, lastPlan:[], originMode:"auto", teamSize:CONFIG.baselineTeamSize, hasTrincia:false };

  const normalizeText = value => String(value ?? "").trim();
  const upper = value => normalizeText(value).toLocaleUpperCase("it-IT");
  const numberValue = value => {
    if (value == null || value === "") return null;
    const text = String(value).trim().replace(/\s/g,"").replace(/\./g,"").replace(",",".").replace(/[^0-9.+-]/g,"");
    const parsed = Number(text);
    return Number.isFinite(parsed) ? parsed : null;
  };
  const readGlobal = name => {
    try {
      if (name === "currentImpianti" && typeof currentImpianti !== "undefined") return currentImpianti;
      if (name === "currentSquadre" && typeof currentSquadre !== "undefined") return currentSquadre;
    } catch (_) {}
    return window[name];
  };

  function todayKey() {
    const now = new Date();
    return `${now.getFullYear()}-${String(now.getMonth()+1).padStart(2,"0")}-${String(now.getDate()).padStart(2,"0")}`;
  }
  function hasRouteStartedToday() {
    try { return localStorage.getItem(STORAGE_ROUTE_DAY) === todayKey() && localStorage.getItem(STORAGE_ROUTE_STARTED) === "1"; }
    catch (_) { return false; }
  }
  function markRouteStartedToday() {
    try { localStorage.setItem(STORAGE_ROUTE_DAY,todayKey()); localStorage.setItem(STORAGE_ROUTE_STARTED,"1"); } catch (_) {}
  }
  function selectedCommessaIdValue() {
    try { return String(selectedCommessaId || window.selectedCommessaId || "").trim(); }
    catch (_) { return String(window.selectedCommessaId || "").trim(); }
  }
  function selectedCommessaRecord() {
    const id = selectedCommessaIdValue();
    try {
      if (typeof commesseById !== "undefined" && commesseById?.get) return commesseById.get(id) || {};
    } catch (_) {}
    const candidates = [window.commesseById, window.currentCommessa, window.selectedCommessa, window.commessaAttiva];
    for (const value of candidates) {
      if (!value) continue;
      if (value instanceof Map) {
        const found = value.get(id);
        if (found) return found;
      } else if (typeof value === "object" && !Array.isArray(value)) {
        if (value[id]) return value[id];
        if (String(value.id ?? "") === id) return value;
      }
    }
    return {};
  }
  function selectedCommessaText() {
    const c = selectedCommessaRecord();
    const dom = [
      document.getElementById("commessa-focus-label")?.textContent,
      document.getElementById("commessa-focus-code")?.textContent,
      document.getElementById("impianti-page-title")?.textContent
    ].filter(Boolean).join(" ");
    return upper([
      c?.nome,c?.codice,c?.categoria,c?.tipo,c?.commessaPadre,c?.parentName,c?.descrizione,dom
    ].filter(Boolean).join(" "));
  }
  function isInreteSelected() {
    return selectedCommessaText().includes("INRETE");
  }

  function rowsFromSource(source, commessaId) {
    if (Array.isArray(source)) return source;
    if (source instanceof Map) {
      const direct = source.get(commessaId);
      return Array.isArray(direct) ? direct : direct ? [direct] : [...source.values()].flatMap(v => Array.isArray(v) ? v : [v]);
    }
    if (source && typeof source === "object") {
      const direct = source[commessaId];
      if (Array.isArray(direct)) return direct;
      if (direct) return [direct];
      return Object.values(source).flatMap(v => Array.isArray(v) ? v : [v]);
    }
    return [];
  }
  function operatorCountFromSquad(squad) {
    for (const value of [squad?.operatori,squad?.operators,squad?.membri,squad?.members,squad?.persone,squad?.componenti,squad?.utenti]) {
      if (Array.isArray(value) && value.length) return value.filter(Boolean).length;
    }
    for (const value of [squad?.numeroOperatori,squad?.operatorCount,squad?.teamSize,squad?.numeroPersone]) {
      const parsed = numberValue(value);
      if (parsed > 0) return Math.round(parsed);
    }
    return 0;
  }
  function collectStrings(value, out = [], depth = 0) {
    if (depth > 5 || value == null) return out;
    if (typeof value === "string" || typeof value === "number") { out.push(String(value)); return out; }
    if (Array.isArray(value)) { value.forEach(v => collectStrings(v,out,depth+1)); return out; }
    if (typeof value === "object") Object.entries(value).forEach(([k,v]) => {
      if (/mez|attrezz|macchin|trinc|decesp|strument|daily|soffiat|tagliasiep|motoseg|spazzatr|piattaform/i.test(k)) collectStrings(v,out,depth+1);
    });
    return out;
  }
  function detectTeamInfo() {
    const commessaId = selectedCommessaIdValue();
    const today = todayKey();
    const sources = [readGlobal("currentSquadre"),window.squadre,window.squadreOggi,window.todaySquads,window.squadreByCommessa];
    let best = null;
    for (const source of sources) {
      for (const squad of rowsFromSource(source,commessaId)) {
        if (!squad || typeof squad !== "object") continue;
        const squadCommessa = String(squad.commessaId ?? squad.idCommessa ?? squad.commessa ?? "").trim();
        const date = String(squad.data ?? squad.date ?? squad.giorno ?? "").slice(0,10);
        if (squadCommessa && commessaId && squadCommessa !== commessaId) continue;
        if (date && date !== today) continue;
        const count = operatorCountFromSquad(squad);
        const equipmentText = upper(collectStrings(squad).join(" "));
        const info = {
          teamSize: count > 0 ? count : CONFIG.baselineTeamSize,
          hasTrincia: equipmentText.includes("TRINCIA") || equipmentText.includes("TRINCIATR"),
          hasDecespugliatore: equipmentText.includes("DECESPUGLI"),
          hasDaily: equipmentText.includes("DAILY") || equipmentText.includes("IVECO"),
          hasSoffiatore: equipmentText.includes("SOFFIAT"),
          hasTagliasiepe: equipmentText.includes("TAGLIASIEP"),
          hasMotosega: equipmentText.includes("MOTOSEG"),
          hasSpazzatrice: equipmentText.includes("SPAZZATR"),
          hasPiattaforma: equipmentText.includes("PIATTAFORM"),
          equipmentText
        };
        if (!best || count > 0 || info.hasTrincia || info.hasDaily) best = info;
        if (count > 0 && (info.hasTrincia || info.hasDaily)) return info;
      }
    }
    return best || {
      teamSize:CONFIG.baselineTeamSize, hasTrincia:false, hasDecespugliatore:false, hasDaily:false,
      hasSoffiatore:false, hasTagliasiepe:false, hasMotosega:false, hasSpazzatrice:false, hasPiattaforma:false, equipmentText:""
    };
  }

  function firstValue(item,names) {
    for (const name of names) {
      const value = item?.[name];
      if (value !== undefined && value !== null && String(value).trim() !== "") return value;
    }
    return "";
  }
  function coordsOf(item) {
    if (!item || typeof item !== "object") return null;
    const latRaw = firstValue(item,["latitudine","lat","latitude","gpsY","GPSY","gps_y","coordinateY","coordY","y","latitudineGps","latGps","gpsLat","gpsLatitude"]);
    const lngRaw = firstValue(item,["longitudine","lng","lon","longitude","gpsX","GPSX","gps_x","coordinateX","coordX","x","longitudineGps","lngGps","lonGps","gpsLng","gpsLon","gpsLongitude"]);
    const repair = window.HeraCoordinateRepair;
    if (repair?.diagnose) {
      const diagnosed = repair.diagnose(latRaw,lngRaw);
      if (diagnosed?.valid) return { lat:diagnosed.latitude, lng:diagnosed.longitude };
    }
    let lat = numberValue(latRaw), lng = numberValue(lngRaw);
    if (Number.isFinite(lat) && Number.isFinite(lng)) {
      const directItaly = lat >= 35 && lat <= 48.8 && lng >= 5 && lng <= 20;
      const swappedItaly = lng >= 35 && lng <= 48.8 && lat >= 5 && lat <= 20;
      if (!directItaly && swappedItaly) [lat,lng] = [lng,lat];
      if (Math.abs(lat)<=90 && Math.abs(lng)<=180 && lat!==0 && lng!==0) return {lat,lng};
    }
    const combined = [
      item.coordinate,item.coordinates,item.coordinateGps,item.coordinateGPS,item.coordinate_gps,item.gps,item.GPS,
      item.posizioneGps,item.posizioneGPS,item.coordinateImpianto,item["Coordinate GPS"],item["COORDINATE GPS"],
      item["Coordinate GPS(X)/GPS(Y)"],item["GPS(X)/GPS(Y)"],item["GPS X/Y"]
    ];
    for (const raw of combined) {
      if (raw == null || String(raw).trim() === "") continue;
      if (repair?.diagnose) {
        const diagnosed = repair.diagnose(raw,"");
        if (diagnosed?.valid) return { lat:diagnosed.latitude, lng:diagnosed.longitude };
      }
      const matches = String(raw).match(/[-+]?\d{1,3}(?:[.,]\d+)?/g) || [];
      if (matches.length >= 2) {
        let a = numberValue(matches[0]), b = numberValue(matches[1]);
        if (!Number.isFinite(a) || !Number.isFinite(b)) continue;
        if (!(a>=35&&a<=48.8&&b>=5&&b<=20) && (b>=35&&b<=48.8&&a>=5&&a<=20)) [a,b]=[b,a];
        if (Math.abs(a)<=90 && Math.abs(b)<=180) return {lat:a,lng:b};
      }
    }
    return null;
  }

  function isDone(item) {
    const status = upper(item?.stato ?? item?.status);
    return Boolean(item?.done) || ["FATTO","DONE","COMPLETATO"].includes(status);
  }
  const titleOf = item => normalizeText(item?.denominazione ?? item?.nome ?? item?.impianto ?? item?.name) || "Impianto";
  const municipalityOf = item => normalizeText(item?.comune ?? item?.municipality ?? item?.citta);
  const idOf = item => normalizeText(item?.id ?? item?.impiantoId ?? item?.idSap ?? item?.idSAP ?? titleOf(item));

  function rawWorkItems(item) {
    const fields = [
      "lavorazioni","lavorazioniRichieste","lavorazioniStraordinarie","extra","extras","lavoriExtra","interventi",
      "attivita","activities","workItems","tipologieIntervento","operazioni","servizi","prestazioni"
    ];
    const out = [];
    fields.forEach(field => {
      const value = item?.[field];
      if (Array.isArray(value)) out.push(...value);
      else if (value && typeof value === "object") out.push(...Object.values(value));
      else if (typeof value === "string" && value.trim()) out.push(...value.split(/[;\n|]+/).filter(Boolean));
    });
    if (!out.length) {
      const text = firstValue(item,[
        "tipologiaIntervento","tipologiaLavorazione","lavorazione","attivitaRichiesta","descrizioneLavoro",
        "descrizione","noteLavoro","note","noteImpianto","modalitaEsecuzione"
      ]);
      if (text) out.push(text);
    }
    return out;
  }
  function workText(work) {
    if (typeof work === "string") return work;
    return [
      work?.nome,work?.titolo,work?.descrizione,work?.tipologia,work?.tipologiaLavorazione,work?.tipologiaIntervento,
      work?.lavorazione,work?.attivita,work?.voce,work?.note,work?.modalita,work?.unitaMisura,work?.codiceVocePrezzo
    ].filter(Boolean).join(" ");
  }
  function itemContextText(item) {
    return upper([
      item?.tipologiaIntervento,item?.tipologiaLavorazione,item?.lavorazione,item?.attivitaRichiesta,item?.descrizioneLavoro,
      item?.descrizione,item?.noteLavoro,item?.note,item?.noteImpianto,item?.modalitaEsecuzione,item?.accesso,item?.difficolta,
      item?.tipologiaImpianto,item?.area,item?.zona,item?.vincoli,item?.sicurezza,item?.atex,item?.ATEX
    ].filter(Boolean).join(" "));
  }
  function quantityOf(work,item,unit) {
    const own = typeof work === "object" ? [
      work.quantita,work.quantity,work.mq,work.areaMq,work.superficieMq,work.metri,work.ml,work.lunghezza,
      work.numero,work.numeroPiante,work.piante
    ] : [];
    const itemVals = unit === "mq"
      ? [item?.areaMq,item?.mq,item?.superficieMq,item?.metriQuadri,item?.superficie,item?.quantitaMq,item?.quantita]
      : unit === "m"
        ? [item?.metriSiepe,item?.lunghezzaSiepe,item?.ml,item?.metriLineari]
        : [item?.numeroPiante,item?.piante,item?.numeroAlberi];
    for (const v of [...own,...itemVals]) {
      const n = numberValue(v);
      if (n > 0) return n;
    }
    const text = workText(work);
    const unitPattern = unit === "mq" ? /(\d+(?:[.,]\d+)?)\s*(?:M2|MQ|M²)/i : unit === "m" ? /(\d+(?:[.,]\d+)?)\s*(?:ML|M\b|METR)/i : /(\d+)\s*(?:PIANT|ALBER|ARBUST)/i;
    const match = text.match(unitPattern);
    return match ? numberValue(match[1]) || 0 : 0;
  }
  function teamFactor(teamSize) {
    return Math.min(1.35,Math.max(0.58,Math.sqrt(CONFIG.baselineTeamSize / Math.max(1,teamSize))));
  }
  function siteDifficultyFactor(item) {
    const text = itemContextText(item);
    let factor = 1;
    if (/SCARPAT|PENDEN|DISLIVELL|FOSSO|ARGINE/.test(text)) factor += 0.18;
    if (/OSTACOL|STRETTO|ACCESSO DIFFIC|PASSAGGIO STRETTO|RECINZ/.test(text)) factor += 0.12;
    if (/ERBA ALTA|VEGETAZIONE ALTA|INFESTANT|MOLTO SPORC|ABBANDON/.test(text)) factor += 0.18;
    if (/ATEX|AREA GAS|CABINA GAS|REMI|GRF/.test(text)) factor += 0.08;
    return Math.min(1.55,factor);
  }
  function equipmentFactor(label,teamInfo) {
    const text = upper(label);
    let factor = 1;
    if (/SIEP/.test(text) && teamInfo.hasTagliasiepe) factor *= 0.82;
    if (/ALBER|ABBATT/.test(text) && teamInfo.hasMotosega) factor *= 0.82;
    if (/PULIZ|RACCOLT/.test(text) && teamInfo.hasSoffiatore) factor *= 0.88;
    if (/SPAZZATR/.test(text) && teamInfo.hasSpazzatrice) factor *= 0.72;
    if (/ALBER|SIEP.*ALTA/.test(text) && teamInfo.hasPiattaforma) factor *= 0.88;
    return factor;
  }
  function estimateOneWork(work,item,teamInfo) {
    const text = upper(workText(work));
    const factor = teamFactor(teamInfo.teamSize) * siteDifficultyFactor(item);
    let label = workText(work).trim() || "Lavorazione";
    let minutes = 0;

    if (/SIEP/.test(text) && /POTAT|TAGLI|RIFIL/.test(text)) {
      const meters = quantityOf(work,item,"m") || 10;
      const difficult = /ALT[AE]|DIFFIC|SCALA|PIATTAFORMA|LECCIO|LAURO.*ALT/.test(text);
      minutes = (meters/10) * (difficult ? 32 : 20) * factor;
      label = `Potatura siepe${difficult ? " difficile/alta" : ""}`;
    } else if (/POTAT/.test(text) && /ALBER|PIANT|ARBOR/.test(text)) {
      const count = quantityOf(work,item,"count") || 1;
      const difficult = /ALTO|ALTA|DIFFIC|PIATTAFORMA|ABBATT/.test(text);
      minutes = count * (difficult ? 60 : 35) * factor;
      label = `Potatura alberi (${Math.round(count)})`;
    } else if (/ROV|BONIFIC/.test(text)) {
      const mq = quantityOf(work,item,"mq") || 100;
      minutes = (mq/100) * 32 * factor;
      label = "Bonifica rovi/vegetazione";
    } else if (/DECESP|A MANO|MANUAL|RIFINIT/.test(text)) {
      const mq = quantityOf(work,item,"mq") || 100;
      const base = /A MANO|MANUAL/.test(text) ? 28 : 20;
      minutes = (mq/100) * base * factor;
      label = /RIFINIT/.test(text) ? "Rifiniture/decespugliatore" : "Sfalcio manuale";
    } else if (/SFALC|TRINCI/.test(text)) {
      const mq = quantityOf(work,item,"mq") || 100;
      const base = teamInfo.hasTrincia ? 12 : 20;
      minutes = (mq/100) * base * factor;
      label = teamInfo.hasTrincia ? "Sfalcio con trincia" : "Sfalcio senza trincia";
    } else if (/PULIZ|RACCOLT|ASPORT|SMALT/.test(text)) {
      const mq = quantityOf(work,item,"mq");
      minutes = (mq ? Math.max(15,(mq/100)*8) : 20) * factor;
      label = "Pulizia/raccolta materiale";
    } else if (/SPAZZATR/.test(text)) {
      const mq = quantityOf(work,item,"mq");
      minutes = (mq ? Math.max(15,(mq/100)*6) : 20) * factor;
      label = "Spazzatrice";
    } else if (/ABBATT/.test(text)) {
      const count = quantityOf(work,item,"count") || 1;
      minutes = count * 75 * factor;
      label = `Abbattimento (${Math.round(count)})`;
    } else {
      minutes = 30 * factor;
      label = label.slice(0,60);
    }

    minutes *= equipmentFactor(label,teamInfo);
    return { label, minutes:Math.max(5,Math.round(minutes)) };
  }
  function collectionMqOf(item) {
    const explicit = [
      item?.mqRaccolta,item?.raccoltaMq,item?.superficieRaccolta,item?.mqDaRaccogliere
    ].map(numberValue).find(v => v > 0);
    if (explicit) return explicit;
    const works = rawWorkItems(item);
    let total = 0;
    for (const work of works) {
      const text = upper(workText(work));
      if (/SFALC|TRINCI|DECESP|BONIFIC|ROV/.test(text)) {
        const mq = quantityOf(work,item,"mq");
        if (mq > 0) total += mq;
      }
    }
    if (total > 0) return total;
    return quantityOf({},item,"mq") || 0;
  }
  function estimateWork(item,teamInfo) {
    const works = rawWorkItems(item);
    const breakdown = works.map(work => estimateOneWork(work,item,teamInfo));
    if (!breakdown.length) {
      const mq = quantityOf({},item,"mq");
      const base = teamInfo.hasTrincia ? 12 : 20;
      const minutes = mq ? (mq/100)*base*teamFactor(teamInfo.teamSize)*siteDifficultyFactor(item) : 30;
      breakdown.push({ label:teamInfo.hasTrincia ? "Sfalcio con trincia" : "Lavorazione manuale", minutes:Math.round(minutes) });
    }
    const context = itemContextText(item);
    if (/ATEX/.test(context)) breakdown.push({label:"Preparazione sicurezza ATEX",minutes:10});
    const productive = breakdown.reduce((sum,x)=>sum+x.minutes,0);
    const total = Math.round((CONFIG.setupMinutesPerPlant + productive) * (1 + CONFIG.contingencyPct/100));
    return { minutes:total, breakdown, collectionMq:collectionMqOf(item) };
  }

  function haversineKm(a,b) {
    const r=6371, rad=v=>v*Math.PI/180, dLat=rad(b.lat-a.lat), dLng=rad(b.lng-a.lng);
    const h=Math.sin(dLat/2)**2+Math.cos(rad(a.lat))*Math.cos(rad(b.lat))*Math.sin(dLng/2)**2;
    return 2*r*Math.asin(Math.sqrt(h));
  }
  const roadEstimateKm = (a,b) => haversineKm(a,b)*CONFIG.roadFactor;
  const travelMinutes = km => Math.max(3,Math.round(km/CONFIG.averageRoadSpeedKmh*60));
  function formatMinutes(total) {
    const min=Math.max(0,Math.round(total||0)), h=Math.floor(min/60), m=min%60;
    return !h ? `${m} min` : !m ? `${h} h` : `${h} h ${m} min`;
  }
  const formatKm = km => `${Number(km||0).toLocaleString("it-IT",{maximumFractionDigits:1})} km`;

  function getCurrentPosition() {
    return new Promise(resolve => {
      if (!navigator.geolocation) return resolve(null);
      navigator.geolocation.getCurrentPosition(
        p=>resolve({lat:p.coords.latitude,lng:p.coords.longitude}),
        ()=>resolve(null),
        {enableHighAccuracy:true,timeout:4500,maximumAge:120000}
      );
    });
  }
  async function resolveOrigin() {
    if (state.originMode==="avola") return {...CONFIG.start,mode:"avola"};
    if (state.originMode==="position") {
      const live=await getCurrentPosition();
      return live ? {...live,name:"Posizione attuale",mode:"position"} : {...CONFIG.start,mode:"avola"};
    }
    if (!hasRouteStartedToday()) return {...CONFIG.start,mode:"avola"};
    const live=await getCurrentPosition();
    return live ? {...live,name:"Posizione attuale",mode:"position"} : {...CONFIG.start,mode:"avola"};
  }

  function buildPlan(items,origin,teamInfo) {
    const remaining=items.filter(i=>!isDone(i)).map(item=>({
      item, coords:coordsOf(item), work:estimateWork(item,teamInfo)
    })).filter(e=>e.coords);
    const plan=[];
    let cursor={lat:origin.lat,lng:origin.lng}, cumulative=0, loadMq=0;
    const inrete = isInreteSelected();

    while (remaining.length) {
      remaining.sort((a,b)=>roadEstimateKm(cursor,a.coords)-roadEstimateKm(cursor,b.coords));
      const next=remaining.shift();
      const km=roadEstimateKm(cursor,next.coords), drive=travelMinutes(km), work=next.work.minutes;
      cumulative += drive + work;
      loadMq += inrete ? next.work.collectionMq : 0;

      let unload = null;
      if (inrete && loadMq >= CONFIG.inreteDailyCapacityMq) {
        const toAvolaKm = roadEstimateKm(next.coords,CONFIG.start);
        const toAvolaMinutes = travelMinutes(toAvolaKm);
        unload = {
          km:toAvolaKm,
          driveMinutes:toAvolaMinutes,
          unloadMinutes:CONFIG.inreteUnloadMinutes,
          loadMq
        };
        cumulative += toAvolaMinutes + CONFIG.inreteUnloadMinutes;
        loadMq = loadMq % CONFIG.inreteDailyCapacityMq;
        cursor={lat:CONFIG.start.lat,lng:CONFIG.start.lng};
      } else {
        cursor=next.coords;
      }

      plan.push({
        ...next,km,driveMinutes:drive,workMinutes:work,unload,
        cumulativeMinutes:cumulative,fitsDay:cumulative<=CONFIG.planningMinutes,
        loadMqAfter:loadMq
      });
    }
    return plan;
  }

  function ensurePanel() {
    const card=document.getElementById("impianti-card"); if(!card) return null;
    let panel=document.getElementById("recommended-plants-panel"); if(panel) return panel;
    panel=document.createElement("section"); panel.id="recommended-plants-panel"; panel.className="recommended-plants-panel hidden"; panel.setAttribute("aria-live","polite");
    const list=document.getElementById("impianti-lista"); list ? card.insertBefore(panel,list) : card.appendChild(panel); return panel;
  }
  function ensureButton() {
    const tabs=document.querySelector("#impianti-card .view-tabs");
    if(!tabs || document.getElementById("recommended-plants-btn")) return;
    const b=document.createElement("button"); b.id="recommended-plants-btn"; b.className="btn recommended-plants-btn"; b.type="button"; b.innerHTML="✨ Impianti consigliati";
    const todo=document.getElementById("view-todo-btn"); todo?.nextSibling ? tabs.insertBefore(b,todo.nextSibling) : tabs.appendChild(b);
    b.addEventListener("click",()=>{state.open=!state.open;state.originMode="auto";render();});
  }
  const escapeHtml=v=>String(v??"").replace(/[&<>\"']/g,c=>({"&":"&amp;","<":"&lt;",">":"&gt;",'\"':"&quot;","'":"&#039;"}[c]));
  function navigationUrl(entry) {
    return `https://www.google.com/maps/dir/?api=1&destination=${encodeURIComponent(`${entry.coords.lat},${entry.coords.lng}`)}&travelmode=driving`;
  }

  async function render() {
    ensureButton();
    const panel=ensurePanel(), button=document.getElementById("recommended-plants-btn");
    if(!panel||!button) return;
    panel.classList.toggle("hidden",!state.open); button.classList.toggle("btn-primary",state.open); button.setAttribute("aria-pressed",String(state.open));
    if(!state.open) return;

    panel.innerHTML='<div class="recommended-loading">Calcolo gli impianti consigliati…</div>';
    const current=readGlobal("currentImpianti"), items=Array.isArray(current)?current:[], origin=await resolveOrigin(), teamInfo=detectTeamInfo(), plan=buildPlan(items,origin,teamInfo);
    state.teamSize=teamInfo.teamSize; state.hasTrincia=teamInfo.hasTrincia; state.lastPlan=plan;

    if(!items.length){panel.innerHTML='<div class="recommended-empty"><strong>Nessun impianto disponibile.</strong><span>Attendi il caricamento della commessa e riprova.</span></div>';return;}
    if(!plan.length){panel.innerHTML='<div class="recommended-empty"><strong>Nessun impianto da pianificare.</strong><span>Gli impianti da fare devono avere coordinate valide.</span></div>';return;}

    const feasible=plan.filter(e=>e.fitsDay), visible=(feasible.length?feasible:plan.slice(0,1)).slice(0,CONFIG.maxVisible);
    const totalDrive=visible.reduce((s,x)=>s+x.driveMinutes+(x.unload?.driveMinutes||0),0);
    const totalWork=visible.reduce((s,x)=>s+x.workMinutes+(x.unload?.unloadMinutes||0),0);
    const originLabel=origin.mode==="avola"?"Avola Coop · primo impianto":"Posizione attuale · squadra già in giro";
    const equipmentLabel=teamInfo.hasTrincia?"🚜 Trincia presente":teamInfo.hasDecespugliatore?"🧑‍🌾 Decespugliatore presente":"🧑‍🌾 Nessuna trincia rilevata";
    const inrete=isInreteSelected();

    panel.innerHTML=`
      <header class="recommended-head"><div><span class="recommended-kicker">PIANIFICAZIONE GIORNATA</span><h3>✨ Impianti consigliati</h3><p>${escapeHtml(originLabel)}</p></div><button id="recommended-close-btn" class="btn btn-small" type="button">Chiudi</button></header>
      <div class="recommended-summary"><span><strong>${visible.length}</strong> consigliati</span><span>👥 ${teamInfo.teamSize} persone</span><span>${equipmentLabel}</span><span>🚐 ${formatMinutes(totalDrive)}</span><span>🌿 ${formatMinutes(totalWork)}</span></div>
      <p class="recommended-rule">Calcolo base: 2 persone. Usa squadra, attrezzature, lavorazioni, extra, difficoltà/accesso e sicurezza dell'impianto. Margine ${CONFIG.contingencyPct}%.</p>
      ${inrete ? `<p class="recommended-rule">INRETE: raccolta inclusa. Ogni ${CONFIG.inreteDailyCapacityMq.toLocaleString("it-IT")} m² il Daily viene considerato pieno e il giro inserisce rientro ad Avola Coop + ${CONFIG.inreteUnloadMinutes} min di scarico.${teamInfo.hasDaily ? "" : " ⚠️ Daily non rilevato nella squadra: verificare mezzo assegnato."}</p>` : ""}
      <div class="recommended-origin-actions"><button id="recommended-origin-avola" class="btn btn-small ${origin.mode==="avola"?"btn-primary":""}" type="button">🏠 Primo giro da Avola</button><button id="recommended-origin-live" class="btn btn-small ${origin.mode==="position"?"btn-primary":""}" type="button">📍 Sono già in giro</button></div>
      <div class="recommended-list">
        ${visible.map((e,index)=>`
          <article class="recommended-item" data-recommended-id="${escapeHtml(idOf(e.item))}">
            <div class="recommended-rank">${index+1}</div>
            <div class="recommended-main">
              <strong>${escapeHtml(titleOf(e.item))}</strong><span>${escapeHtml(municipalityOf(e.item))}</span>
              <small>🚐 ${formatKm(e.km)} · ${formatMinutes(e.driveMinutes)} &nbsp; 🌿 lavoro ${formatMinutes(e.workMinutes)}</small>
              <small>${e.work.breakdown.map(x=>`${escapeHtml(x.label)}: ${formatMinutes(x.minutes)}`).join(" · ")}</small>
              ${inrete && e.work.collectionMq ? `<small>🟩 Raccolta stimata: ${Math.round(e.work.collectionMq).toLocaleString("it-IT")} m² · carico residuo Daily: ${Math.round(e.loadMqAfter).toLocaleString("it-IT")} / ${CONFIG.inreteDailyCapacityMq.toLocaleString("it-IT")} m²</small>` : ""}
              ${e.unload ? `<small>🚛 Daily pieno → rientro Avola: ${formatKm(e.unload.km)} · ${formatMinutes(e.unload.driveMinutes)} + scarico ${formatMinutes(e.unload.unloadMinutes)}</small>` : ""}
            </div>
            <a class="btn recommended-nav-btn" href="${navigationUrl(e)}" target="_blank" rel="noopener" data-recommended-nav="1">NAVIGA</a>
          </article>`).join("")}
      </div>
      <div class="recommended-footer"><button id="recommended-start-route" class="btn btn-primary" type="button">AVVIA GIRO CONSIGLIATO</button><small>La lista normale resta ordinata per distanza come prima.</small></div>`;

    panel.querySelector("#recommended-close-btn")?.addEventListener("click",()=>{state.open=false;render();});
    panel.querySelector("#recommended-origin-avola")?.addEventListener("click",()=>{state.originMode="avola";render();});
    panel.querySelector("#recommended-origin-live")?.addEventListener("click",()=>{state.originMode="position";render();});
    panel.querySelector("#recommended-start-route")?.addEventListener("click",()=>{markRouteStartedToday();state.originMode="position";const b=panel.querySelector("#recommended-start-route");if(b)b.textContent="✓ GIRO AVVIATO";});
    panel.querySelectorAll("[data-recommended-nav='1']").forEach(link=>link.addEventListener("click",markRouteStartedToday,{passive:true}));
  }

  function install() {
    ensureButton(); ensurePanel();
    const page=document.getElementById("impianti-page");
    if(page) new MutationObserver(()=>{ensureButton();ensurePanel();}).observe(page,{childList:true,subtree:true});
  }
  if(document.readyState==="loading") document.addEventListener("DOMContentLoaded",install,{once:true}); else install();

  window.HeraRecommendedPlants={
    installed:true, version:"2.1.0", config:CONFIG,
    open:()=>{state.open=true;state.originMode="auto";return render();},
    refresh:render, markRouteStartedToday,
    getState:()=>({...state,lastPlan:state.lastPlan.slice()})
  };
})();