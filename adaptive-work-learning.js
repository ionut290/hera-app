(() => {
  "use strict";
  if (window.HeraAdaptiveWorkLearning?.installed) return;

  const VERSION = "1.0.0";
  const STORAGE_PROFILES = "heraAdaptiveWorkProfilesV1";
  const STORAGE_ACTIVE = "heraAdaptiveWorkActiveV1";
  const MIN_TOTAL_MINUTES = 8;
  const MAX_TOTAL_MINUTES = 12 * 60;
  const MIN_WORK_MINUTES = 5;
  const MAX_PROFILE_SAMPLES = 80;
  const DAY_MS = 24 * 60 * 60 * 1000;
  let applyingPanel = false;

  const text = value => String(value ?? "").trim();
  const upper = value => text(value).toLocaleUpperCase("it-IT");
  const num = value => {
    if (value == null || value === "") return 0;
    const parsed = Number(String(value).trim().replace(/\s/g, "").replace(/\./g, "").replace(",", ".").replace(/[^0-9.+-]/g, ""));
    return Number.isFinite(parsed) ? parsed : 0;
  };
  const clamp = (value, min, max) => Math.min(max, Math.max(min, value));
  const escapeHtml = value => String(value ?? "").replace(/[&<>\"']/g, ch => ({"&":"&amp;","<":"&lt;",">":"&gt;",'\"':"&quot;","'":"&#039;"}[ch]));

  function readJson(key, fallback) {
    try {
      const raw = localStorage.getItem(key);
      return raw ? JSON.parse(raw) : fallback;
    } catch (_) { return fallback; }
  }
  function writeJson(key, value) {
    try { localStorage.setItem(key, JSON.stringify(value)); return true; }
    catch (_) { return false; }
  }
  function profiles() { return readJson(STORAGE_PROFILES, {}); }
  function saveProfiles(value) { writeJson(STORAGE_PROFILES, value); }
  function activeSession() { return readJson(STORAGE_ACTIVE, null); }
  function saveActive(value) {
    if (value) writeJson(STORAGE_ACTIVE, value);
    else { try { localStorage.removeItem(STORAGE_ACTIVE); } catch (_) {} }
  }

  function currentPlants() {
    try {
      if (typeof currentImpianti !== "undefined" && Array.isArray(currentImpianti)) return currentImpianti;
    } catch (_) {}
    return Array.isArray(window.currentImpianti) ? window.currentImpianti : [];
  }
  function selectedCommessaId() {
    try { return text(selectedCommessaId || window.selectedCommessaId); }
    catch (_) { return text(window.selectedCommessaId); }
  }
  function plantId(item) {
    return text(item?.physicalPlantId || item?.impiantoId || item?.migrationSourceId || item?.id || item?.idSap || item?.idSAP || item?.sap || item?.denominazione || item?.nome);
  }
  function plantTitle(item) { return text(item?.denominazione || item?.nome || item?.impianto || item?.name || item?.idSap) || "Impianto"; }

  function collectStrings(value, out = [], depth = 0) {
    if (depth > 5 || value == null) return out;
    if (["string","number"].includes(typeof value)) { out.push(String(value)); return out; }
    if (Array.isArray(value)) { value.forEach(v => collectStrings(v, out, depth + 1)); return out; }
    if (typeof value === "object") Object.entries(value).forEach(([key, val]) => {
      if (/mez|attrezz|macchin|trinc|decesp|strument|daily|iveco|soffiat|tagliasiep|motoseg|spazzatr|piattaform/i.test(key)) collectStrings(val, out, depth + 1);
    });
    return out;
  }
  function rowsFromSource(source, commessaId) {
    if (Array.isArray(source)) return source;
    if (source instanceof Map) {
      const direct = source.get(commessaId);
      if (Array.isArray(direct)) return direct;
      if (direct) return [direct];
      return [...source.values()].flatMap(v => Array.isArray(v) ? v : [v]);
    }
    if (source && typeof source === "object") {
      const direct = source[commessaId];
      if (Array.isArray(direct)) return direct;
      if (direct) return [direct];
      return Object.values(source).flatMap(v => Array.isArray(v) ? v : [v]);
    }
    return [];
  }
  function operatorCount(squad) {
    for (const value of [squad?.operatori,squad?.operators,squad?.membri,squad?.members,squad?.persone,squad?.componenti,squad?.utenti]) {
      if (Array.isArray(value) && value.filter(Boolean).length) return value.filter(Boolean).length;
    }
    for (const value of [squad?.numeroOperatori,squad?.operatorCount,squad?.teamSize,squad?.numeroPersone]) {
      const n = num(value); if (n > 0) return Math.round(n);
    }
    return 0;
  }
  function todayKey() {
    const d = new Date();
    return `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,"0")}-${String(d.getDate()).padStart(2,"0")}`;
  }
  function detectTeamInfo() {
    const commessaId = selectedCommessaId();
    const today = todayKey();
    const sources = [];
    try { if (typeof currentSquadre !== "undefined") sources.push(currentSquadre); } catch (_) {}
    sources.push(window.currentSquadre, window.squadre, window.squadreOggi, window.todaySquads, window.squadreByCommessa);
    let best = null;
    for (const source of sources) {
      for (const squad of rowsFromSource(source, commessaId)) {
        if (!squad || typeof squad !== "object") continue;
        const squadCommessa = text(squad.commessaId ?? squad.idCommessa ?? squad.commessa);
        const date = text(squad.data ?? squad.date ?? squad.giorno).slice(0, 10);
        if (squadCommessa && commessaId && squadCommessa !== commessaId) continue;
        if (date && date !== today) continue;
        const teamSize = operatorCount(squad) || 2;
        const equipmentText = upper(collectStrings(squad).join(" "));
        const info = {
          teamSize,
          hasTrincia: /TRINCIA|TRINCIATR/.test(equipmentText),
          hasDecespugliatore: /DECESPUGLI/.test(equipmentText),
          hasDaily: /DAILY|IVECO/.test(equipmentText) || /(?:^|\s)A\s*[-_.]?\s*\d{1,5}(?:\s|$)/.test(equipmentText),
          hasTagliasiepe: /TAGLIASIEP/.test(equipmentText),
          hasMotosega: /MOTOSEG/.test(equipmentText),
          hasSoffiatore: /SOFFIAT/.test(equipmentText),
          hasSpazzatrice: /SPAZZATR/.test(equipmentText),
          hasPiattaforma: /PIATTAFORM/.test(equipmentText),
          equipmentText
        };
        if (!best || teamSize > 0 || info.hasTrincia) best = info;
      }
    }
    const rp = window.HeraRecommendedPlants?.getState?.();
    if (rp?.teamSize) {
      best = best || { teamSize: rp.teamSize };
      best.teamSize = rp.teamSize;
      if (typeof rp.hasTrincia === "boolean") best.hasTrincia = rp.hasTrincia;
    }
    return Object.assign({teamSize:2,hasTrincia:false,hasDecespugliatore:false,hasDaily:false,hasTagliasiepe:false,hasMotosega:false,hasSoffiatore:false,hasSpazzatrice:false,hasPiattaforma:false,equipmentText:""}, best || {});
  }

  function rawWorkItems(item) {
    const fields = ["lavorazioni","lavorazioniRichieste","lavorazioniStraordinarie","extra","extras","lavoriExtra","interventi","attivita","activities","workItems","tipologieIntervento","operazioni","servizi","prestazioni"];
    const out = [];
    fields.forEach(field => {
      const value = item?.[field];
      if (Array.isArray(value)) out.push(...value);
      else if (value && typeof value === "object") out.push(...Object.values(value));
      else if (typeof value === "string" && value.trim()) out.push(...value.split(/[;\n|]+/).filter(Boolean));
    });
    if (!out.length) {
      const fallback = [item?.tipologiaIntervento,item?.tipologiaLavorazione,item?.lavorazione,item?.attivitaRichiesta,item?.descrizioneLavoro,item?.descrizione,item?.noteLavoro,item?.note].filter(Boolean).join(" ");
      if (fallback) out.push(fallback);
    }
    return out;
  }
  function workText(work) {
    if (typeof work === "string") return work;
    return [work?.nome,work?.titolo,work?.descrizione,work?.tipologia,work?.tipologiaLavorazione,work?.tipologiaIntervento,work?.lavorazione,work?.attivita,work?.voce,work?.note,work?.modalita,work?.unitaMisura,work?.codiceVocePrezzo].filter(Boolean).join(" ");
  }
  function quantity(work, item, unit) {
    const own = typeof work === "object" ? [work.quantita,work.quantity,work.mq,work.areaMq,work.superficieMq,work.metri,work.ml,work.lunghezza,work.numero,work.numeroPiante,work.piante] : [];
    const itemValues = unit === "mq" ? [item?.areaMq,item?.mq,item?.superficieMq,item?.metriQuadri,item?.superficie,item?.quantitaMq,item?.quantita]
      : unit === "m" ? [item?.metriSiepe,item?.lunghezzaSiepe,item?.ml,item?.metriLineari]
      : [item?.numeroPiante,item?.piante,item?.numeroAlberi];
    for (const value of [...own, ...itemValues]) { const n = num(value); if (n > 0) return n; }
    const t = workText(work);
    const regex = unit === "mq" ? /(\d+(?:[.,]\d+)?)\s*(?:M2|MQ|M²)/i : unit === "m" ? /(\d+(?:[.,]\d+)?)\s*(?:ML|METR|M\b)/i : /(\d+)\s*(?:PIANT|ALBER|ARBUST)/i;
    const match = t.match(regex); return match ? num(match[1]) : 0;
  }

  function classifyWork(work, item, team) {
    const t = upper(workText(work));
    if (/SIEP/.test(t) && /POTAT|TAGLI|RIFIL/.test(t)) return {kind:"potatura_siepe",unit:"10m",quantity:(quantity(work,item,"m")||10)/10,baseRate: /ALT[AE]|DIFFIC|PIATTAFORMA/.test(t)?32:20};
    if (/POTAT/.test(t) && /ALBER|PIANT|ARBOR/.test(t)) return {kind:"potatura_alberi",unit:"pianta",quantity:quantity(work,item,"count")||1,baseRate:/ALTO|ALTA|DIFFIC|PIATTAFORMA/.test(t)?60:35};
    if (/ABBATT/.test(t)) return {kind:"abbattimento",unit:"pianta",quantity:quantity(work,item,"count")||1,baseRate:75};
    if (/ROV|BONIFIC/.test(t)) return {kind:"bonifica_rovi",unit:"100mq",quantity:(quantity(work,item,"mq")||100)/100,baseRate:32};
    if (/RIFINIT/.test(t)) return {kind:"rifiniture",unit:"100mq",quantity:(quantity(work,item,"mq")||100)/100,baseRate:20};
    if (/DECESP|A MANO|MANUAL/.test(t)) return {kind:/A MANO|MANUAL/.test(t)?"lavoro_manual":"sfalcio_decespugliatore",unit:"100mq",quantity:(quantity(work,item,"mq")||100)/100,baseRate:/A MANO|MANUAL/.test(t)?28:20};
    if (/SFALC|TRINCI/.test(t)) return {kind:team.hasTrincia?"sfalcio_trincia":"sfalcio_senza_trincia",unit:"100mq",quantity:(quantity(work,item,"mq")||100)/100,baseRate:team.hasTrincia?12:20};
    if (/PULIZ|RACCOLT|ASPORT|SMALT/.test(t)) return {kind:"pulizia_raccolta",unit:"100mq",quantity:Math.max(1,(quantity(work,item,"mq")||100)/100),baseRate:8};
    if (/SPAZZATR/.test(t)) return {kind:"spazzatrice",unit:"100mq",quantity:Math.max(1,(quantity(work,item,"mq")||100)/100),baseRate:6};
    return {kind:"lavorazione_generica",unit:"intervento",quantity:1,baseRate:30};
  }
  function workUnits(item, team) {
    const works = rawWorkItems(item);
    const units = (works.length ? works : [""]).map(work => classifyWork(work,item,team));
    const seen = new Set();
    return units.filter(unit => {
      const key = `${unit.kind}|${unit.unit}|${Math.round(unit.quantity*100)/100}`;
      if (seen.has(key)) return false;
      seen.add(key); return true;
    });
  }
  function profileKey(unit, team) {
    const equipment = unit.kind === "sfalcio_trincia" || unit.kind === "sfalcio_senza_trincia" ? `|trincia=${team.hasTrincia?1:0}` : "";
    return `${unit.kind}|team=${Math.max(1,Math.round(team.teamSize||2))}${equipment}`;
  }
  function learnedRate(unit, team) {
    const p = profiles()[profileKey(unit,team)];
    if (!p || !Number.isFinite(p.emaRate) || p.emaRate <= 0) return {rate:unit.baseRate,confidence:0,count:0};
    const confidence = clamp((p.weightedCount || p.count || 0) / 8, 0, 1);
    const safe = clamp(p.emaRate, unit.baseRate * 0.5, unit.baseRate * 2.0);
    return {rate:unit.baseRate*(1-confidence)+safe*confidence,confidence,count:p.count||0};
  }
  function adaptiveMinutes(item, team, fallback) {
    const units = workUnits(item,team);
    let total = 0, samples = 0, maxConfidence = 0;
    units.forEach(unit => {
      const learned = learnedRate(unit,team);
      total += unit.quantity * learned.rate;
      samples += learned.count;
      maxConfidence = Math.max(maxConfidence, learned.confidence);
    });
    if (!total) return {minutes:fallback,samples,confidence:maxConfidence};
    const baselineRaw = units.reduce((sum,u)=>sum+u.quantity*u.baseRate,0);
    const baseline = Math.max(1, baselineRaw);
    const fallbackCore = Math.max(1, Number(fallback||baseline));
    const ratio = clamp(total / baseline, 0.55, 1.8);
    return {minutes:Math.round(fallbackCore*ratio),samples,confidence:maxConfidence,ratio};
  }

  function findPlantByElement(element) {
    if (!(element instanceof Element)) return null;
    const card = element.closest("[data-recommended-id],[data-impianto-id],[data-plant-id],[data-id-sap],.impianto-card,.impianto-item,.plant-card,article,li,.card") || element.parentElement;
    const candidates = [card?.dataset?.recommendedId,card?.dataset?.impiantoId,card?.dataset?.plantId,card?.dataset?.idSap,element.dataset?.impiantoId,element.dataset?.plantId].map(text).filter(Boolean);
    const plants = currentPlants();
    for (const id of candidates) {
      const exact = plants.find(p => [plantId(p),text(p?.id),text(p?.idSap),text(p?.idSAP),text(p?.sap)].includes(id));
      if (exact) return exact;
    }
    const body = upper(card?.textContent || "");
    let best = null, bestLen = 0;
    for (const plant of plants) {
      for (const token of [plantId(plant),plant?.idSap,plant?.denominazione,plant?.nome].map(text).filter(v=>v.length>=3)) {
        if (body.includes(upper(token)) && token.length > bestLen) { best=plant; bestLen=token.length; }
      }
    }
    return best;
  }
  function labelOf(element) { return upper(element?.textContent || element?.getAttribute?.("aria-label") || element?.value || ""); }
  function isNavigate(element) {
    const node = element?.closest?.("a,button,[role='button'],input[type='button'],input[type='submit']");
    if (!node) return null;
    const label = labelOf(node);
    return /\bNAVIGA\b|NAVIGA VERSO/.test(label) ? node : null;
  }
  function isFatto(element) {
    const node = element?.closest?.("button,[role='button'],input[type='button'],input[type='submit']");
    if (!node) return null;
    return /(^|\s)FATTO($|\s)/.test(labelOf(node)) ? node : null;
  }
  function estimateTravelForPlant(item, card) {
    const state = window.HeraRecommendedPlants?.getState?.();
    const id = plantId(item);
    const entry = state?.lastPlan?.find?.(row => plantId(row?.item) === id);
    if (entry?.driveMinutes > 0) return {minutes:entry.driveMinutes,source:"recommended"};
    const body = text(card?.textContent);
    const eta = body.match(/ETA\s*(\d+)\s*MIN/i) || body.match(/(\d+)\s*MIN\s*(?:DI\s*)?(?:VIAGGIO|AUTO)/i);
    if (eta) return {minutes:Number(eta[1]),source:"card"};
    return {minutes:0,source:"unknown"};
  }

  function startSession(item, sourceNode) {
    if (!item) return;
    const team = detectTeamInfo();
    const travel = estimateTravelForPlant(item, sourceNode?.closest?.("article,li,.card,.impianto-card,.impianto-item"));
    const now = Date.now();
    const session = {
      version:1,
      commessaId:selectedCommessaId(),
      plantId:plantId(item),
      plantTitle:plantTitle(item),
      navigateAt:now,
      date:todayKey(),
      travelMinutes:travel.minutes,
      travelSource:travel.source,
      team,
      units:workUnits(item,team),
      itemSnapshot:{
        areaMq:num(item?.areaMq||item?.mq||item?.superficieMq||item?.metriQuadri||item?.superficie),
        metriSiepe:num(item?.metriSiepe||item?.lunghezzaSiepe||item?.ml||item?.metriLineari),
        numeroPiante:num(item?.numeroPiante||item?.piante||item?.numeroAlberi)
      }
    };
    saveActive(session);
    window.dispatchEvent(new CustomEvent("hera:adaptive-learning-start",{detail:{plantId:session.plantId,navigateAt:session.navigateAt}}));
  }

  function updateProfile(store, key, unit, observedRate, quality, meta) {
    const previous = store[key] || {kind:unit.kind,unit:unit.unit,count:0,weightedCount:0,emaRate:unit.baseRate,meanRate:unit.baseRate,totalWeight:0,totalRateWeight:0};
    let rate = clamp(observedRate, unit.baseRate*0.3, unit.baseRate*3.0);
    if ((previous.count||0) >= 3 && previous.emaRate > 0) rate = clamp(rate, previous.emaRate*0.5, previous.emaRate*2.0);
    const q = clamp(quality,0.15,1);
    const alpha = clamp(0.12 + q*0.18,0.12,0.30);
    const nextWeight = (previous.totalWeight||0)+q;
    const nextRateWeight = (previous.totalRateWeight||0)+rate*q;
    store[key] = {
      ...previous,
      count:Math.min(MAX_PROFILE_SAMPLES,(previous.count||0)+1),
      weightedCount:Math.min(MAX_PROFILE_SAMPLES,(previous.weightedCount||0)+q),
      emaRate:(previous.count||0) ? previous.emaRate*(1-alpha)+rate*alpha : rate,
      meanRate:nextRateWeight/Math.max(0.001,nextWeight),
      totalWeight:nextWeight,
      totalRateWeight:nextRateWeight,
      lastRate:rate,
      lastUpdated:new Date().toISOString(),
      lastPlantId:meta.plantId,
      lastTeamSize:meta.teamSize,
      lastDurationMinutes:meta.workMinutes
    };
  }

  function finishSession(item) {
    const session = activeSession();
    if (!session || !session.navigateAt) return null;
    if (Date.now()-session.navigateAt > DAY_MS || session.date !== todayKey()) { saveActive(null); return null; }
    if (item && plantId(item) && session.plantId && plantId(item) !== session.plantId) return null;
    const totalMinutes = (Date.now()-session.navigateAt)/60000;
    if (totalMinutes < MIN_TOTAL_MINUTES || totalMinutes > MAX_TOTAL_MINUTES) { saveActive(null); return null; }
    const travelKnown = session.travelMinutes > 0;
    const workMinutes = Math.max(MIN_WORK_MINUTES,totalMinutes-(travelKnown?session.travelMinutes:0));
    const units = Array.isArray(session.units)&&session.units.length ? session.units : [{kind:"lavorazione_generica",unit:"intervento",quantity:1,baseRate:30}];
    const baselineWeights = units.map(u=>Math.max(1,u.quantity*u.baseRate));
    const weightTotal = baselineWeights.reduce((a,b)=>a+b,0)||1;
    let quality = travelKnown ? 0.9 : 0.55;
    const expected = weightTotal;
    const ratio = workMinutes/Math.max(1,expected);
    if (ratio<0.35 || ratio>3.2) quality*=0.45;
    else if (ratio<0.55 || ratio>2.2) quality*=0.7;
    const store = profiles();
    units.forEach((unit,index)=>{
      const allocated = workMinutes*(baselineWeights[index]/weightTotal);
      const observedRate = allocated/Math.max(0.1,unit.quantity);
      updateProfile(store,profileKey(unit,session.team||{teamSize:2}),unit,observedRate,quality,{plantId:session.plantId,teamSize:session.team?.teamSize||2,workMinutes});
    });
    saveProfiles(store);
    saveActive(null);
    const detail={plantId:session.plantId,totalMinutes:Math.round(totalMinutes),travelMinutes:Math.round(session.travelMinutes||0),workMinutes:Math.round(workMinutes),quality,profiles:units.map(u=>profileKey(u,session.team||{teamSize:2}))};
    window.dispatchEvent(new CustomEvent("hera:adaptive-learning-sample",{detail}));
    window.setTimeout(applyToRecommendedPanel,120);
    return detail;
  }

  function formatMinutes(value) {
    const min=Math.max(0,Math.round(value||0)),h=Math.floor(min/60),m=min%60;
    return !h?`${m} min`:!m?`${h} h`:`${h} h ${m} min`;
  }
  function formatKm(km) { return `${Number(km||0).toLocaleString("it-IT",{maximumFractionDigits:1})} km`; }
  function navUrl(entry) { return `https://www.google.com/maps/dir/?api=1&destination=${encodeURIComponent(`${entry.coords.lat},${entry.coords.lng}`)}&travelmode=driving`; }

  function applyToRecommendedPanel() {
    if (applyingPanel) return;
    const panel=document.getElementById("recommended-plants-panel");
    if (!panel || panel.classList.contains("hidden")) return;
    const api=window.HeraRecommendedPlants;
    const state=api?.getState?.();
    const plan=state?.lastPlan;
    if (!Array.isArray(plan)||!plan.length) return;
    applyingPanel=true;
    try {
      const team=detectTeamInfo();
      let cumulative=0;
      const adaptive=plan.map(entry=>{
        const adjusted=adaptiveMinutes(entry.item,team,entry.workMinutes);
        cumulative+=entry.driveMinutes+adjusted.minutes+(entry.unload?.driveMinutes||0)+(entry.unload?.unloadMinutes||0);
        return {...entry,adaptiveWorkMinutes:adjusted.minutes,adaptiveSamples:adjusted.samples,adaptiveConfidence:adjusted.confidence,adaptiveCumulative:cumulative,adaptiveFits:cumulative<=8*60};
      });
      let visible=adaptive.filter(e=>e.adaptiveFits).slice(0,8);
      if (!visible.length) visible=adaptive.slice(0,1);
      const list=panel.querySelector(".recommended-list");
      if (!list) return;
      list.innerHTML=visible.map((e,index)=>{
        const ratio=(e.workMinutes>0)?e.adaptiveWorkMinutes/e.workMinutes:1;
        const detail=(e.work?.breakdown||[]).map(x=>`${escapeHtml(x.label)}: ${formatMinutes((x.minutes||0)*ratio)}`).join(" · ");
        const learned=e.adaptiveSamples>0?`<small>🧠 Stima adattiva · ${e.adaptiveSamples} campioni reali</small>`:"<small>🧠 Stima iniziale · apprenderà dai lavori reali</small>";
        return `<article class="recommended-item" data-recommended-id="${escapeHtml(plantId(e.item))}"><div class="recommended-rank">${index+1}</div><div class="recommended-main"><strong>${escapeHtml(plantTitle(e.item))}</strong><span>${escapeHtml(e.item?.comune||e.item?.municipality||e.item?.citta||"")}</span><small>🚐 ${formatKm(e.km)} · ${formatMinutes(e.driveMinutes)} &nbsp; 🌿 lavoro ${formatMinutes(e.adaptiveWorkMinutes)}</small>${detail?`<small>${detail}</small>`:""}${learned}${e.unload?`<small>🚛 Daily pieno → rientro Avola: ${formatKm(e.unload.km)} · ${formatMinutes(e.unload.driveMinutes)} + scarico ${formatMinutes(e.unload.unloadMinutes)}</small>`:""}</div><a class="btn recommended-nav-btn" href="${navUrl(e)}" target="_blank" rel="noopener" data-recommended-nav="1">NAVIGA</a></article>`;
      }).join("");
      const summary=panel.querySelector(".recommended-summary");
      if(summary){
        const totalWork=visible.reduce((s,e)=>s+e.adaptiveWorkMinutes+(e.unload?.unloadMinutes||0),0);
        const spans=summary.querySelectorAll("span");
        if(spans[0]) spans[0].innerHTML=`<strong>${visible.length}</strong> consigliati`;
        if(spans.length) {
          let badge=summary.querySelector("[data-adaptive-summary]");
          if(!badge){badge=document.createElement("span");badge.dataset.adaptiveSummary="1";summary.appendChild(badge);}
          const sampleCount=visible.reduce((s,e)=>s+(e.adaptiveSamples||0),0);
          badge.textContent=sampleCount?`🧠 ${sampleCount} dati reali · lavoro ${formatMinutes(totalWork)}`:"🧠 Apprendimento attivo";
        }
      }
      list.querySelectorAll("[data-recommended-nav='1']").forEach(link=>link.addEventListener("click",()=>{
        const item=findPlantByElement(link); if(item) startSession(item,link);
        try{api?.markRouteStartedToday?.();}catch(_){}
      },{passive:true}));
    } finally { window.setTimeout(()=>{applyingPanel=false;},60); }
  }

  document.addEventListener("click",event=>{
    const nav=isNavigate(event.target);
    if(nav){const item=findPlantByElement(nav);if(item) startSession(item,nav);return;}
    const fatto=isFatto(event.target);
    if(!fatto)return;
    const item=findPlantByElement(fatto);
    const session=activeSession();
    if(!session)return;
    if(item&&plantId(item)!==session.plantId)return;
    window.setTimeout(()=>finishSession(item),900);
  },true);

  function installObserver(){
    const root=document.getElementById("recommended-plants-panel")||document.body;
    if(!root)return;
    let timer=0;
    new MutationObserver(()=>{clearTimeout(timer);timer=window.setTimeout(applyToRecommendedPanel,80);}).observe(root,{childList:true,subtree:true,attributes:true,attributeFilter:["class"]});
  }
  if(document.readyState==="loading")document.addEventListener("DOMContentLoaded",installObserver,{once:true});else installObserver();

  window.HeraAdaptiveWorkLearning={
    installed:true,version:VERSION,
    getProfiles:()=>JSON.parse(JSON.stringify(profiles())),
    getActiveSession:()=>activeSession(),
    reset:()=>{saveProfiles({});saveActive(null);applyToRecommendedPanel();},
    applyToRecommendedPanel,
    adaptiveMinutes:(item,team,fallback)=>adaptiveMinutes(item,team||detectTeamInfo(),fallback),
    finishActive:()=>finishSession(null)
  };
})();
